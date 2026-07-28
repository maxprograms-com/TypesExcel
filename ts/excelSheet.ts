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

import { CellValue } from './cellValue.js';

export interface ExcelSheetOptions {
    columnWidths?: number[];
    boldHeader?: boolean;
}

export class ExcelSheet {
    private name: string;
    private headers: string[];
    private rows: CellValue[][];
    private options: ExcelSheetOptions;

    constructor(name: string, headers: string[], rows: CellValue[][], options?: ExcelSheetOptions) {
        this.name = name;
        this.headers = headers;
        this.rows = rows;
        this.options = options ?? {};
    }

    getName(): string {
        return this.name;
    }

    getHeaders(): string[] {
        return this.headers;
    }

    getRows(): CellValue[][] {
        return this.rows;
    }

    getColumnWidths(): number[] | undefined {
        return this.options.columnWidths;
    }

    getBoldHeader(): boolean {
        return this.options.boldHeader ?? false;
    }
}
