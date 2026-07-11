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

import { readFileSync } from 'node:fs';
import { dirname, join, posix } from 'node:path';
import { fileURLToPath } from 'node:url';
import { DOMBuilder, SAXParser, XMLElement } from 'typesxml';
import { CellValue } from './cellValue.js';
import { I18n } from './i18n.js';
import { ColumnUtils } from './columnUtils.js';
import { ExcelSheet } from './excelSheet.js';
import { ExcelWorkbook } from './excelWorkbook.js';
import { ZipUtils } from './zipUtils.js';

export interface ExcelSheetData {
    name: string;
    rows: CellValue[][];
}

interface Relationship {
    id: string;
    type: string;
    target: string;
}

export class ExcelReader {
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

    readSheets(filePath: string): ExcelSheetData[] {
        let { entries, workbookDir, workbookRelationships, relsById, sheetsElement } = this.loadWorkbookContext(filePath);
        let sharedStrings: string[] = this.loadSharedStrings(entries, workbookDir, workbookRelationships);
        let dateStyleIndices: Set<number> = this.parseDateStyleIndices(entries);
        let sheets: ExcelSheetData[] = [];
        for (let sheetElement of sheetsElement.getChildren().filter((child: XMLElement) => child.getName() === 'sheet')) {
            let state: string | undefined = sheetElement.getAttribute('state')?.getValue();
            if (state === 'hidden' || state === 'veryHidden') {
                continue;
            }
            let rId: string | undefined = sheetElement.getAttribute('r:id')?.getValue();
            let target: string | undefined = rId ? relsById.get(rId) : undefined;
            if (!target) {
                continue;
            }
            let name: string = sheetElement.getAttribute('name')?.getValue() || '';
            let sheetPath: string = this.resolveZipPath(workbookDir, target);
            let sheetRoot: XMLElement = this.parseXml(this.requireEntry(entries, sheetPath));
            sheets.push({ name, rows: this.parseSheetRows(sheetRoot, sharedStrings, dateStyleIndices) });
        }
        if (sheets.length === 0) {
            throw new Error(this.i18n.getString('ExcelReader', 'noSheets'));
        }
        return sheets;
    }

    listSheets(filePath: string): string[] {
        let { sheetsElement } = this.loadWorkbookContext(filePath);
        let names: string[] = [];
        for (let sheetElement of sheetsElement.getChildren().filter((child: XMLElement) => child.getName() === 'sheet')) {
            let state: string | undefined = sheetElement.getAttribute('state')?.getValue();
            if (state === 'hidden' || state === 'veryHidden') {
                continue;
            }
            let name: string = sheetElement.getAttribute('name')?.getValue() || '';
            names.push(name);
        }
        return names;
    }

    readSheet(filePath: string, sheetName: string): ExcelSheetData {
        let { entries, workbookDir, workbookRelationships, relsById, sheetsElement } = this.loadWorkbookContext(filePath);
        for (let sheetElement of sheetsElement.getChildren().filter((child: XMLElement) => child.getName() === 'sheet')) {
            let state: string | undefined = sheetElement.getAttribute('state')?.getValue();
            if (state === 'hidden' || state === 'veryHidden') {
                continue;
            }
            let name: string = sheetElement.getAttribute('name')?.getValue() || '';
            if (name !== sheetName) {
                continue;
            }
            let rId: string | undefined = sheetElement.getAttribute('r:id')?.getValue();
            let target: string | undefined = rId ? relsById.get(rId) : undefined;
            if (!target) {
                continue;
            }
            let sharedStrings: string[] = this.loadSharedStrings(entries, workbookDir, workbookRelationships);
            let dateStyleIndices: Set<number> = this.parseDateStyleIndices(entries);
            let sheetPath: string = this.resolveZipPath(workbookDir, target);
            let sheetRoot: XMLElement = this.parseXml(this.requireEntry(entries, sheetPath));
            return { name, rows: this.parseSheetRows(sheetRoot, sharedStrings, dateStyleIndices) };
        }
        throw new Error(this.i18n.format(this.i18n.getString('ExcelReader', 'sheetNotFound'), [sheetName]));
    }

    readWorkbook(filePath: string): ExcelWorkbook {
        let { entries, workbookDir, workbookRelationships, relsById, sheetsElement } = this.loadWorkbookContext(filePath);
        let sharedStrings: string[] = this.loadSharedStrings(entries, workbookDir, workbookRelationships);
        let dateStyleIndices: Set<number> = this.parseDateStyleIndices(entries);
        let sheets: ExcelSheet[] = [];
        for (let sheetElement of sheetsElement.getChildren().filter((child: XMLElement) => child.getName() === 'sheet')) {
            let state: string | undefined = sheetElement.getAttribute('state')?.getValue();
            if (state === 'hidden' || state === 'veryHidden') {
                continue;
            }
            let rId: string | undefined = sheetElement.getAttribute('r:id')?.getValue();
            let target: string | undefined = rId ? relsById.get(rId) : undefined;
            if (!target) {
                continue;
            }
            let name: string = sheetElement.getAttribute('name')?.getValue() || '';
            let sheetPath: string = this.resolveZipPath(workbookDir, target);
            let sheetRoot: XMLElement = this.parseXml(this.requireEntry(entries, sheetPath));
            let rows: CellValue[][] = this.parseSheetRows(sheetRoot, sharedStrings, dateStyleIndices);
            let headers: string[] = rows.length > 0 ? rows[0].map((v: CellValue) => v !== null ? String(v) : '') : [];
            let dataRows: CellValue[][] = rows.slice(1);
            sheets.push(new ExcelSheet(name, headers, dataRows));
        }
        return new ExcelWorkbook(sheets);
    }

    private loadWorkbookContext(filePath: string): {
        entries: Map<string, Buffer>,
        workbookDir: string,
        workbookRelationships: Relationship[],
        relsById: Map<string, string>,
        sheetsElement: XMLElement
    } {
        let entries: Map<string, Buffer>;
        try {
            entries = this.zipUtils.readZip(readFileSync(filePath));
        } catch (error) {
            throw new Error(this.i18n.format(this.i18n.getString('ExcelReader', 'notAWorkbook'), [error instanceof Error ? error.message : String(error)]));
        }
        let rootRelationships: Relationship[] = this.parseRelationships(entries, '_rels/.rels');
        let officeDocument: Relationship | undefined = rootRelationships.find((relationship: Relationship) => relationship.type.endsWith('/officeDocument'));
        if (!officeDocument) {
            throw new Error(this.i18n.format(this.i18n.getString('ExcelReader', 'missingWorkbookPart'), ['_rels/.rels']));
        }
        let workbookPath: string = this.resolveZipPath('', officeDocument.target);
        let workbookDir: string = posix.dirname(workbookPath);
        let workbookRelsPath: string = posix.join(workbookDir, '_rels', posix.basename(workbookPath) + '.rels');
        let workbookRelationships: Relationship[] = this.parseRelationships(entries, workbookRelsPath);
        let relsById: Map<string, string> = new Map();
        for (let relationship of workbookRelationships) {
            relsById.set(relationship.id, relationship.target);
        }
        let workbookRoot: XMLElement = this.parseXml(this.requireEntry(entries, workbookPath));
        let sheetsElement: XMLElement | undefined = workbookRoot.getChild('sheets');
        if (!sheetsElement) {
            throw new Error(this.i18n.getString('ExcelReader', 'noSheets'));
        }
        return { entries, workbookDir, workbookRelationships, relsById, sheetsElement };
    }

    private loadSharedStrings(entries: Map<string, Buffer>, workbookDir: string, workbookRelationships: Relationship[]): string[] {
        let sharedStrings: string[] = [];
        let sharedStringsRelationship: Relationship | undefined = workbookRelationships.find((relationship: Relationship) => relationship.type.endsWith('/sharedStrings'));
        if (sharedStringsRelationship) {
            let sharedStringsPath: string = this.resolveZipPath(workbookDir, sharedStringsRelationship.target);
            let data: Buffer | undefined = entries.get(sharedStringsPath);
            if (data) {
                sharedStrings = this.parseSharedStrings(this.parseXml(data));
            }
        }
        return sharedStrings;
    }

    private parseSheetRows(sheetRoot: XMLElement, sharedStrings: string[], dateStyleIndices: Set<number>): CellValue[][] {
        let sheetDataElement: XMLElement | undefined = sheetRoot.getChild('sheetData');
        if (!sheetDataElement) {
            return [];
        }

        let rawRows: Map<number, CellValue>[] = [];
        let maxColumn: number = -1;
        for (let rowElement of sheetDataElement.getChildren().filter((child: XMLElement) => child.getName() === 'row')) {
            let rowValues: Map<number, CellValue> = new Map();
            for (let cellElement of rowElement.getChildren().filter((child: XMLElement) => child.getName() === 'c')) {
                let reference: string | undefined = cellElement.getAttribute('r')?.getValue();
                let match: RegExpMatchArray | null = reference ? reference.match(/^[A-Z]+/) : null;
                if (!match) {
                    continue;
                }
                let value: CellValue = this.cellValue(cellElement, sharedStrings, dateStyleIndices);
                if (value === null) {
                    continue;
                }
                let index: number = this.columnUtils.columnIndex(match[0]);
                rowValues.set(index, value);
                if (index > maxColumn) {
                    maxColumn = index;
                }
            }
            rawRows.push(rowValues);
        }

        let rows: CellValue[][] = [];
        for (let rowValues of rawRows) {
            let row: CellValue[] = [];
            for (let c = 0; c <= maxColumn; c++) {
                row.push(rowValues.get(c) ?? null);
            }
            rows.push(row);
        }
        return rows;
    }

    private cellValue(cellElement: XMLElement, sharedStrings: string[], dateStyleIndices: Set<number>): CellValue {
        let type: string = cellElement.getAttribute('t')?.getValue() || '';
        if (type === 's') {
            let v: XMLElement | undefined = cellElement.getChild('v');
            if (!v) {
                return null;
            }
            return sharedStrings[parseInt(v.getText(), 10)] ?? null;
        }
        if (type === 'inlineStr') {
            let is: XMLElement | undefined = cellElement.getChild('is');
            return is ? this.fullText(is) : null;
        }
        if (type === 'b') {
            let v: XMLElement | undefined = cellElement.getChild('v');
            return v ? v.getText() !== '0' : null;
        }
        if (type === 'e') {
            let v: XMLElement | undefined = cellElement.getChild('v');
            return v ? v.getText() : null;
        }
        let v: XMLElement | undefined = cellElement.getChild('v');
        if (!v) {
            return null;
        }
        let n: number = parseFloat(v.getText());
        if (isNaN(n)) {
            return null;
        }
        let styleIndex: number = parseInt(cellElement.getAttribute('s')?.getValue() || '0', 10);
        if (dateStyleIndices.has(styleIndex)) {
            return this.serialToDate(n);
        }
        return n;
    }

    private serialToDate(serial: number): Date {
        return new Date(serial * 86400000 + Date.UTC(1899, 11, 30));
    }

    private parseDateStyleIndices(entries: Map<string, Buffer>): Set<number> {
        let builtinDateFmtIds: Set<number> = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47]);
        let dateFmtIds: Set<number> = new Set(builtinDateFmtIds);
        let dateStyleIndices: Set<number> = new Set();
        let data: Buffer | undefined = entries.get('xl/styles.xml');
        if (!data) {
            return dateStyleIndices;
        }
        let root: XMLElement = this.parseXml(data);
        let numFmtsElement: XMLElement | undefined = root.getChild('numFmts');
        if (numFmtsElement) {
            for (let numFmt of numFmtsElement.getChildren().filter((child: XMLElement) => child.getName() === 'numFmt')) {
                let id: number = parseInt(numFmt.getAttribute('numFmtId')?.getValue() || '0', 10);
                let code: string = numFmt.getAttribute('formatCode')?.getValue() || '';
                if (/[ydhm]/i.test(code) && !/^[0#,\.\s"]+$/.test(code)) {
                    dateFmtIds.add(id);
                }
            }
        }
        let cellXfsElement: XMLElement | undefined = root.getChild('cellXfs');
        if (cellXfsElement) {
            let index: number = 0;
            for (let xf of cellXfsElement.getChildren().filter((child: XMLElement) => child.getName() === 'xf')) {
                let numFmtId: number = parseInt(xf.getAttribute('numFmtId')?.getValue() || '0', 10);
                if (dateFmtIds.has(numFmtId)) {
                    dateStyleIndices.add(index);
                }
                index++;
            }
        }
        return dateStyleIndices;
    }

    private parseSharedStrings(root: XMLElement): string[] {
        let strings: string[] = [];
        for (let si of root.getChildren().filter((child: XMLElement) => child.getName() === 'si')) {
            strings.push(this.fullText(si));
        }
        return strings;
    }

    private parseRelationships(entries: Map<string, Buffer>, relsPath: string): Relationship[] {
        let data: Buffer | undefined = entries.get(relsPath);
        if (!data) {
            return [];
        }
        let root: XMLElement = this.parseXml(data);
        let relationships: Relationship[] = [];
        for (let element of root.getChildren().filter((child: XMLElement) => child.getName() === 'Relationship')) {
            relationships.push({
                id: element.getAttribute('Id')?.getValue() || '',
                type: element.getAttribute('Type')?.getValue() || '',
                target: element.getAttribute('Target')?.getValue() || ''
            });
        }
        return relationships;
    }

    private requireEntry(entries: Map<string, Buffer>, path: string): Buffer {
        let data: Buffer | undefined = entries.get(path);
        if (!data) {
            throw new Error(this.i18n.format(this.i18n.getString('ExcelReader', 'missingWorkbookPart'), [path]));
        }
        return data;
    }

    private parseXml(data: Buffer): XMLElement {
        let contentHandler: DOMBuilder = new DOMBuilder();
        let parser: SAXParser = new SAXParser();
        parser.setSchemaLoadingEnabled(false);
        parser.setValidating(false);
        parser.setContentHandler(contentHandler);
        parser.parseString(data.toString('utf8'));
        let root: XMLElement | undefined = contentHandler.getDocument()?.getRoot();
        if (!root) {
            throw new Error(this.i18n.getString('ExcelReader', 'unparsableXmlPart'));
        }
        return root;
    }

    private fullText(container: XMLElement): string {
        let text: string = '';
        for (let child of container.getChildren()) {
            let name: string = child.getName();
            if (name === 'rPh' || name === 'phoneticPr') {
                continue; // skip phonetic (ruby) annotations
            }
            text += this.firstText(child);
        }
        return text;
    }

    private firstText(element: XMLElement): string {
        if (element.getName() === 't') {
            return element.getText();
        }
        for (let child of element.getChildren()) {
            let name: string = child.getName();
            if (name === 'rPh' || name === 'phoneticPr') {
                continue; // skip phonetic (ruby) annotations
            }
            let text: string = this.firstText(child);
            if (text) {
                return text;
            }
        }
        return '';
    }

    private resolveZipPath(baseDir: string, target: string): string {
        if (target.startsWith('/')) {
            return posix.normalize(target.slice(1));
        }
        return posix.normalize(posix.join(baseDir, target));
    }
}
