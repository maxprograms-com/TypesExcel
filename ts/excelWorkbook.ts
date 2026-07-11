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

import { ExcelSheet } from './excelSheet.js';

export class ExcelWorkbook {
    private sheets: ExcelSheet[];

    constructor(sheets: ExcelSheet[] = []) {
        this.sheets = sheets;
    }

    addSheet(sheet: ExcelSheet): void {
        this.sheets.push(sheet);
    }

    removeSheet(name: string): void {
        this.sheets = this.sheets.filter((s: ExcelSheet) => s.getName() !== name);
    }

    getSheet(name: string): ExcelSheet | undefined {
        return this.sheets.find((s: ExcelSheet) => s.getName() === name);
    }

    getSheets(): ExcelSheet[] {
        return [...this.sheets];
    }
}
