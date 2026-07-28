/*******************************************************************************
 * Copyright (c) 2026 Maxprograms.
 *
 * This program and the accompanying materials
 * are made available under the terms of the Eclipse Public License 1.0
 * which accompanies this distribution, and is available at
 * https://www.eclipse.org/org/documents/epl-v10.html
 *
 * Contributors:
 *     Maxprograms - initial API and implementation
 *******************************************************************************/

import { readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { DOMBuilder, SAXParser, XMLAttribute, XMLDocument, XMLElement } from 'typesxml';
import { CellValue } from './cellValue.js';
import { I18n } from './i18n.js';
import { ColumnUtils } from './columnUtils.js';
import { ExcelSheet } from './excelSheet.js';
import { ExcelWorkbook } from './excelWorkbook.js';
import { ZipUtils } from './zipUtils.js';

const TEMPLATE_PATH: string = join(dirname(fileURLToPath(import.meta.url)), 'template.xlsx');
const DEFAULT_COLUMN_WIDTH: string = '30';

export class ExcelWriter {
    private i18n: I18n;
    private columnUtils: ColumnUtils;
    private zipUtils: ZipUtils;

    constructor(language: string) {
        this.i18n = new I18n(this.resolveMessagesPath(language));
        this.columnUtils = new ColumnUtils();
        this.zipUtils = new ZipUtils();
    }

    private resolveMessagesPath(language: string): string {
        return join(dirname(fileURLToPath(import.meta.url)), 'excel_' + language + '.json');
    }

    writeFile(filePath: string, data: ExcelSheet | ExcelWorkbook): void {
        let templateBuffer: Buffer = readFileSync(TEMPLATE_PATH);
        let entries: Map<string, Buffer>;
        try {
            entries = this.zipUtils.readZip(templateBuffer);
        } catch (error) {
            throw new Error(this.i18n.format(this.i18n.getString('ExcelWriter', 'corruptTemplate'), [error instanceof Error ? error.message : String(error)]));
        }

        if (data instanceof ExcelWorkbook) {
            let sheets: ExcelSheet[] = data.getSheets();
            if (sheets.length === 0) {
                throw new Error(this.i18n.getString('ExcelWriter', 'emptyWorkbook'));
            }
            this.writeMultipleSheets(entries, sheets);
        } else {
            this.setSheetName(entries, data.getName());
            this.setSheetData(entries, data);
        }

        writeFileSync(filePath, this.zipUtils.writeZip(entries));
    }

    private writeMultipleSheets(entries: Map<string, Buffer>, sheets: ExcelSheet[]): void {
        let WORKSHEET_REL_TYPE: string = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
        let WORKSHEET_CONTENT_TYPE: string = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';

        let hasDate: boolean = false;
        for (let sheet of sheets) {
            let allRows: CellValue[][] = [sheet.getHeaders(), ...sheet.getRows()];
            if (allRows.some((row: CellValue[]) => row.some((v: CellValue) => v instanceof Date))) {
                hasDate = true;
                break;
            }
        }
        let dateStyleIndex: number = hasDate ? this.ensureDateStyleIndex(entries) : -1;

        let hasBoldHeader: boolean = sheets.some((sheet: ExcelSheet) => sheet.getBoldHeader());
        let headerStyleIndex: number = hasBoldHeader ? this.ensureHeaderStyleIndex(entries) : -1;

        let unique: Map<string, number> = new Map();
        let totalStringCount: number = 0;
        for (let sheet of sheets) {
            let allRows: CellValue[][] = [sheet.getHeaders(), ...sheet.getRows()];
            for (let row of allRows) {
                for (let v of row) {
                    if (typeof v === 'string') {
                        if (!unique.has(v)) {
                            unique.set(v, unique.size);
                        }
                        totalStringCount++;
                    }
                }
            }
        }
        let sstDocument: XMLDocument = this.parseEntry(entries, 'xl/sharedStrings.xml');
        let sstRoot: XMLElement = this.requireRoot(sstDocument, 'xl/sharedStrings.xml');
        sstRoot.setContent([]);
        for (let [text] of unique) {
            let si: XMLElement = new XMLElement('si');
            let t: XMLElement = new XMLElement('t');
            if (text.indexOf('\n') !== -1) {
                t.setAttribute(new XMLAttribute('xml:space', 'preserve'));
            }
            t.addString(text);
            si.addElement(t);
            sstRoot.addElement(si);
        }
        sstRoot.setAttribute(new XMLAttribute('count', '' + totalStringCount));
        sstRoot.setAttribute(new XMLAttribute('uniqueCount', '' + unique.size));
        this.writeEntry(entries, 'xl/sharedStrings.xml', sstDocument);

        for (let i = 0; i < sheets.length; i++) {
            let sheet: ExcelSheet = sheets[i];
            let sheetPath: string = 'xl/worksheets/sheet' + (i + 1) + '.xml';
            let sheetDocument: XMLDocument = this.parseEntry(entries, 'xl/worksheets/sheet1.xml');
            let sheetRoot: XMLElement = this.requireRoot(sheetDocument, 'xl/worksheets/sheet1.xml');

            if (i > 0) {
                sheetRoot.getChild('sheetViews')?.getChild('sheetView')?.removeAttribute('tabSelected');
                sheetRoot.removeAttribute('xr:uid');
            }

            let headers: string[] = sheet.getHeaders();
            let columnWidths: number[] | undefined = sheet.getColumnWidths();
            let columnsElement: XMLElement | undefined = sheetRoot.getChild('cols');
            if (columnsElement) {
                columnsElement.setContent([]);
                for (let c = 0; c < headers.length; c++) {
                    let col: XMLElement = new XMLElement('col');
                    col.setAttribute(new XMLAttribute('min', '' + (c + 1)));
                    col.setAttribute(new XMLAttribute('max', '' + (c + 1)));
                    col.setAttribute(new XMLAttribute('width', '' + (columnWidths?.[c] ?? DEFAULT_COLUMN_WIDTH)));
                    col.setAttribute(new XMLAttribute('customWidth', '1'));
                    columnsElement.addElement(col);
                }
            }

            let sheetDataElement: XMLElement | undefined = sheetRoot.getChild('sheetData');
            if (!sheetDataElement) {
                throw new Error(this.i18n.getString('ExcelWriter', 'missingSheetDataElement'));
            }
            sheetDataElement.setContent([]);

            let allRows: CellValue[][] = [headers, ...sheet.getRows()];
            for (let r = 0; r < allRows.length; r++) {
                let rowValues: CellValue[] = allRows[r];
                let spans: string = '1:' + rowValues.length;
                let rowElement: XMLElement = new XMLElement('row');
                rowElement.setAttribute(new XMLAttribute('r', '' + (r + 1)));
                rowElement.setAttribute(new XMLAttribute('spans', spans));

                for (let c = 0; c < rowValues.length; c++) {
                    let cellValue: CellValue = rowValues[c];
                    if (cellValue === null || cellValue === undefined) {
                        continue;
                    }
                    let cellElement: XMLElement = new XMLElement('c');
                    cellElement.setAttribute(new XMLAttribute('r', this.columnUtils.columnName(c) + (r + 1)));

                    if (typeof cellValue === 'string') {
                        let isBoldHeaderRow: boolean = r === 0 && sheet.getBoldHeader();
                        cellElement.setAttribute(new XMLAttribute('t', 's'));
                        cellElement.setAttribute(new XMLAttribute('s', '' + (isBoldHeaderRow ? headerStyleIndex : 1)));
                        let vElement: XMLElement = new XMLElement('v');
                        vElement.addString('' + unique.get(cellValue));
                        cellElement.addElement(vElement);
                    } else if (typeof cellValue === 'number') {
                        cellElement.setAttribute(new XMLAttribute('s', '0'));
                        let vElement: XMLElement = new XMLElement('v');
                        vElement.addString('' + cellValue);
                        cellElement.addElement(vElement);
                    } else if (typeof cellValue === 'boolean') {
                        cellElement.setAttribute(new XMLAttribute('t', 'b'));
                        let vElement: XMLElement = new XMLElement('v');
                        vElement.addString(cellValue ? '1' : '0');
                        cellElement.addElement(vElement);
                    } else if (cellValue instanceof Date) {
                        cellElement.setAttribute(new XMLAttribute('s', '' + dateStyleIndex));
                        let vElement: XMLElement = new XMLElement('v');
                        vElement.addString('' + this.dateToSerial(cellValue));
                        cellElement.addElement(vElement);
                    }
                    rowElement.addElement(cellElement);
                }
                sheetDataElement.addElement(rowElement);
            }
            entries.set(sheetPath, Buffer.from(sheetDocument.toString(), 'utf8'));
        }

        let workbookDocument: XMLDocument = this.parseEntry(entries, 'xl/workbook.xml');
        let workbookRoot: XMLElement = this.requireRoot(workbookDocument, 'xl/workbook.xml');
        let sheetsElement: XMLElement | undefined = workbookRoot.getChild('sheets');
        if (!sheetsElement) {
            throw new Error(this.i18n.getString('ExcelWriter', 'missingSheetElement'));
        }
        let existingSheetEls: XMLElement[] = sheetsElement.getChildren().filter((child: XMLElement) => child.getName() === 'sheet');
        if (existingSheetEls.length > 0) {
            existingSheetEls[0].setAttribute(new XMLAttribute('name', sheets[0].getName()));
        }
        let maxSheetId: number = 0;
        for (let el of existingSheetEls) {
            let sid: number = parseInt(el.getAttribute('sheetId')?.getValue() || '0', 10);
            if (sid > maxSheetId) {
                maxSheetId = sid;
            }
        }

        let relsDocument: XMLDocument = this.parseEntry(entries, 'xl/_rels/workbook.xml.rels');
        let relsRoot: XMLElement = this.requireRoot(relsDocument, 'xl/_rels/workbook.xml.rels');
        let maxRId: number = 0;
        for (let rel of relsRoot.getChildren().filter((child: XMLElement) => child.getName() === 'Relationship')) {
            let num: number = parseInt((rel.getAttribute('Id')?.getValue() || '').replace('rId', ''), 10);
            if (!isNaN(num) && num > maxRId) {
                maxRId = num;
            }
        }

        for (let i = 1; i < sheets.length; i++) {
            let newRId: string = 'rId' + (maxRId + i);
            let sheetEl: XMLElement = new XMLElement('sheet');
            sheetEl.setAttribute(new XMLAttribute('name', sheets[i].getName()));
            sheetEl.setAttribute(new XMLAttribute('sheetId', '' + (maxSheetId + i)));
            sheetEl.setAttribute(new XMLAttribute('r:id', newRId));
            sheetsElement.addElement(sheetEl);

            let rel: XMLElement = new XMLElement('Relationship');
            rel.setAttribute(new XMLAttribute('Id', newRId));
            rel.setAttribute(new XMLAttribute('Type', WORKSHEET_REL_TYPE));
            rel.setAttribute(new XMLAttribute('Target', 'worksheets/sheet' + (i + 1) + '.xml'));
            relsRoot.addElement(rel);
        }
        this.writeEntry(entries, 'xl/workbook.xml', workbookDocument);
        this.writeEntry(entries, 'xl/_rels/workbook.xml.rels', relsDocument);

        if (sheets.length > 1) {
            let ctDocument: XMLDocument = this.parseEntry(entries, '[Content_Types].xml');
            let ctRoot: XMLElement = this.requireRoot(ctDocument, '[Content_Types].xml');
            for (let i = 1; i < sheets.length; i++) {
                let override: XMLElement = new XMLElement('Override');
                override.setAttribute(new XMLAttribute('PartName', '/xl/worksheets/sheet' + (i + 1) + '.xml'));
                override.setAttribute(new XMLAttribute('ContentType', WORKSHEET_CONTENT_TYPE));
                ctRoot.addElement(override);
            }
            this.writeEntry(entries, '[Content_Types].xml', ctDocument);
        }
    }

    private setSheetName(entries: Map<string, Buffer>, name: string): void {
        let document: XMLDocument = this.parseEntry(entries, 'xl/workbook.xml');
        let root: XMLElement = this.requireRoot(document, 'xl/workbook.xml');
        let sheets: XMLElement | undefined = root.getChild('sheets');
        let sheetElement: XMLElement | undefined = sheets?.getChild('sheet');
        if (!sheetElement) {
            throw new Error(this.i18n.getString('ExcelWriter', 'missingSheetElement'));
        }
        sheetElement.setAttribute(new XMLAttribute('name', name));
        this.writeEntry(entries, 'xl/workbook.xml', document);
    }

    private setSheetData(entries: Map<string, Buffer>, sheet: ExcelSheet): void {
        let headers: string[] = sheet.getHeaders();
        let rows: CellValue[][] = sheet.getRows();

        let sstDocument: XMLDocument = this.parseEntry(entries, 'xl/sharedStrings.xml');
        let sstRoot: XMLElement = this.requireRoot(sstDocument, 'xl/sharedStrings.xml');
        sstRoot.setContent([]);

        let sheetDocument: XMLDocument = this.parseEntry(entries, 'xl/worksheets/sheet1.xml');
        let sheetRoot: XMLElement = this.requireRoot(sheetDocument, 'xl/worksheets/sheet1.xml');

        let columnWidths: number[] | undefined = sheet.getColumnWidths();
        let columnsElement: XMLElement | undefined = sheetRoot.getChild('cols');
        if (columnsElement) {
            columnsElement.setContent([]);
            for (let i = 0; i < headers.length; i++) {
                let col: XMLElement = new XMLElement('col');
                col.setAttribute(new XMLAttribute('min', '' + (i + 1)));
                col.setAttribute(new XMLAttribute('max', '' + (i + 1)));
                col.setAttribute(new XMLAttribute('width', '' + (columnWidths?.[i] ?? DEFAULT_COLUMN_WIDTH)));
                col.setAttribute(new XMLAttribute('customWidth', '1'));
                columnsElement.addElement(col);
            }
        }

        let sheetDataElement: XMLElement | undefined = sheetRoot.getChild('sheetData');
        if (!sheetDataElement) {
            throw new Error(this.i18n.getString('ExcelWriter', 'missingSheetDataElement'));
        }
        sheetDataElement.setContent([]);

        let hasDate: boolean = rows.some((row: CellValue[]) => row.some((v: CellValue) => v instanceof Date));
        let dateStyleIndex: number = hasDate ? this.ensureDateStyleIndex(entries) : -1;

        let headerStyleIndex: number = sheet.getBoldHeader() ? this.ensureHeaderStyleIndex(entries) : -1;

        let unique: Map<string, number> = new Map();
        let stringCount: number = 0;
        let allRows: CellValue[][] = [headers, ...rows];

        for (let r = 0; r < allRows.length; r++) {
            let rowValues: CellValue[] = allRows[r];
            let spans: string = '1:' + rowValues.length;
            let rowElement: XMLElement = new XMLElement('row');
            rowElement.setAttribute(new XMLAttribute('r', '' + (r + 1)));
            rowElement.setAttribute(new XMLAttribute('spans', spans));

            for (let c = 0; c < rowValues.length; c++) {
                let cellValue: CellValue = rowValues[c];
                if (cellValue === null || cellValue === undefined) {
                    continue;
                }

                let cellElement: XMLElement = new XMLElement('c');
                cellElement.setAttribute(new XMLAttribute('r', this.columnUtils.columnName(c) + (r + 1)));

                if (typeof cellValue === 'string') {
                    if (!unique.has(cellValue)) {
                        unique.set(cellValue, unique.size);
                        let si: XMLElement = new XMLElement('si');
                        let t: XMLElement = new XMLElement('t');
                        if (cellValue.indexOf('\n') !== -1) {
                            t.setAttribute(new XMLAttribute('xml:space', 'preserve'));
                        }
                        t.addString(cellValue);
                        si.addElement(t);
                        sstRoot.addElement(si);
                    }
                    let isBoldHeaderRow: boolean = r === 0 && sheet.getBoldHeader();
                    cellElement.setAttribute(new XMLAttribute('t', 's'));
                    cellElement.setAttribute(new XMLAttribute('s', '' + (isBoldHeaderRow ? headerStyleIndex : 1)));
                    let vElement: XMLElement = new XMLElement('v');
                    vElement.addString('' + unique.get(cellValue));
                    cellElement.addElement(vElement);
                    stringCount++;
                } else if (typeof cellValue === 'number') {
                    cellElement.setAttribute(new XMLAttribute('s', '0'));
                    let vElement: XMLElement = new XMLElement('v');
                    vElement.addString('' + cellValue);
                    cellElement.addElement(vElement);
                } else if (typeof cellValue === 'boolean') {
                    cellElement.setAttribute(new XMLAttribute('t', 'b'));
                    let vElement: XMLElement = new XMLElement('v');
                    vElement.addString(cellValue ? '1' : '0');
                    cellElement.addElement(vElement);
                } else if (cellValue instanceof Date) {
                    cellElement.setAttribute(new XMLAttribute('s', '' + dateStyleIndex));
                    let vElement: XMLElement = new XMLElement('v');
                    vElement.addString('' + this.dateToSerial(cellValue));
                    cellElement.addElement(vElement);
                }

                rowElement.addElement(cellElement);
            }
            sheetDataElement.addElement(rowElement);
        }

        sstRoot.setAttribute(new XMLAttribute('count', '' + stringCount));
        sstRoot.setAttribute(new XMLAttribute('uniqueCount', '' + unique.size));

        this.writeEntry(entries, 'xl/sharedStrings.xml', sstDocument);
        this.writeEntry(entries, 'xl/worksheets/sheet1.xml', sheetDocument);
    }

    private ensureDateStyleIndex(entries: Map<string, Buffer>): number {
        let stylesDocument: XMLDocument = this.parseEntry(entries, 'xl/styles.xml');
        let root: XMLElement = this.requireRoot(stylesDocument, 'xl/styles.xml');
        let builtinDateFmtIds: Set<number> = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47]);
        let cellXfsElement: XMLElement | undefined = root.getChild('cellXfs');
        if (!cellXfsElement) {
            return 0;
        }
        let xfs: XMLElement[] = cellXfsElement.getChildren().filter((child: XMLElement) => child.getName() === 'xf');
        for (let i = 0; i < xfs.length; i++) {
            let numFmtId: number = parseInt(xfs[i].getAttribute('numFmtId')?.getValue() || '0', 10);
            if (builtinDateFmtIds.has(numFmtId)) {
                return i;
            }
        }
        let newXf: XMLElement = new XMLElement('xf');
        newXf.setAttribute(new XMLAttribute('numFmtId', '14'));
        newXf.setAttribute(new XMLAttribute('fontId', '0'));
        newXf.setAttribute(new XMLAttribute('fillId', '0'));
        newXf.setAttribute(new XMLAttribute('borderId', '0'));
        newXf.setAttribute(new XMLAttribute('xfId', '0'));
        newXf.setAttribute(new XMLAttribute('applyNumberFormat', '1'));
        cellXfsElement.addElement(newXf);
        cellXfsElement.setAttribute(new XMLAttribute('count', '' + (xfs.length + 1)));
        this.writeEntry(entries, 'xl/styles.xml', stylesDocument);
        return xfs.length;
    }

    private ensureHeaderStyleIndex(entries: Map<string, Buffer>): number {
        let stylesDocument: XMLDocument = this.parseEntry(entries, 'xl/styles.xml');
        let root: XMLElement = this.requireRoot(stylesDocument, 'xl/styles.xml');

        let fontsElement: XMLElement | undefined = root.getChild('fonts');
        let cellXfsElement: XMLElement | undefined = root.getChild('cellXfs');
        if (!fontsElement || !cellXfsElement) {
            return 1;
        }

        let fonts: XMLElement[] = fontsElement.getChildren().filter((child: XMLElement) => child.getName() === 'font');
        let boldFont: XMLElement = new XMLElement('font');
        boldFont.addElement(new XMLElement('b'));
        let sz: XMLElement = new XMLElement('sz');
        sz.setAttribute(new XMLAttribute('val', '12'));
        boldFont.addElement(sz);
        let color: XMLElement = new XMLElement('color');
        color.setAttribute(new XMLAttribute('theme', '1'));
        boldFont.addElement(color);
        let name: XMLElement = new XMLElement('name');
        name.setAttribute(new XMLAttribute('val', 'Calibri'));
        boldFont.addElement(name);
        let family: XMLElement = new XMLElement('family');
        family.setAttribute(new XMLAttribute('val', '2'));
        boldFont.addElement(family);
        let scheme: XMLElement = new XMLElement('scheme');
        scheme.setAttribute(new XMLAttribute('val', 'minor'));
        boldFont.addElement(scheme);

        let boldFontId: number = fonts.length;
        fontsElement.addElement(boldFont);
        fontsElement.setAttribute(new XMLAttribute('count', '' + (fonts.length + 1)));

        let xfs: XMLElement[] = cellXfsElement.getChildren().filter((child: XMLElement) => child.getName() === 'xf');
        let newXf: XMLElement = new XMLElement('xf');
        newXf.setAttribute(new XMLAttribute('numFmtId', '0'));
        newXf.setAttribute(new XMLAttribute('fontId', '' + boldFontId));
        newXf.setAttribute(new XMLAttribute('fillId', '0'));
        newXf.setAttribute(new XMLAttribute('borderId', '0'));
        newXf.setAttribute(new XMLAttribute('xfId', '0'));
        newXf.setAttribute(new XMLAttribute('applyFont', '1'));
        newXf.setAttribute(new XMLAttribute('applyAlignment', '1'));
        let alignment: XMLElement = new XMLElement('alignment');
        alignment.setAttribute(new XMLAttribute('vertical', 'top'));
        alignment.setAttribute(new XMLAttribute('wrapText', '1'));
        newXf.addElement(alignment);
        cellXfsElement.addElement(newXf);
        cellXfsElement.setAttribute(new XMLAttribute('count', '' + (xfs.length + 1)));

        this.writeEntry(entries, 'xl/styles.xml', stylesDocument);
        return xfs.length;
    }

    private dateToSerial(date: Date): number {
        return (date.getTime() - Date.UTC(1899, 11, 30)) / 86400000;
    }

    private parseEntry(entries: Map<string, Buffer>, name: string): XMLDocument {
        let data: Buffer | undefined = entries.get(name);
        if (!data) {
            throw new Error(this.i18n.format(this.i18n.getString('ExcelWriter', 'missingTemplatePart'), [name]));
        }
        let contentHandler: DOMBuilder = new DOMBuilder();
        let parser: SAXParser = new SAXParser();
        parser.setSchemaLoadingEnabled(false);
        parser.setValidating(false);
        parser.setContentHandler(contentHandler);
        parser.parseString(data.toString('utf8'));
        let document: XMLDocument | undefined = contentHandler.getDocument();
        if (!document) {
            throw new Error(this.i18n.format(this.i18n.getString('ExcelWriter', 'unparsableTemplatePart'), [name]));
        }
        return document;
    }

    private requireRoot(document: XMLDocument, name: string): XMLElement {
        let root: XMLElement | undefined = document.getRoot();
        if (!root) {
            throw new Error(this.i18n.format(this.i18n.getString('ExcelWriter', 'missingRootElement'), [name]));
        }
        return root;
    }

    private writeEntry(entries: Map<string, Buffer>, name: string, document: XMLDocument): void {
        entries.set(name, Buffer.from(document.toString(), 'utf8'));
    }
}
