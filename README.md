# TypesExcel

[![npm version](https://img.shields.io/npm/v/typesexcel)](https://www.npmjs.com/package/typesexcel)
[![npm license](https://img.shields.io/npm/l/typesexcel)](LICENSE)
[![TypeScript](https://img.shields.io/badge/implementation-native%20TypeScript-3178c6)](https://www.typescriptlang.org/)

TypesExcel is a TypeScript / Node.js library for reading and writing Excel (.xlsx) files. It works directly with the OOXML ZIP structure — no native bindings, no external spreadsheet engine.

The library is built on [TypesXML](https://github.com/maxprograms-com/TypesXML) for XML parsing and exposes a simple, strongly typed API for extracting sheet data or generating new workbooks from code, with typed cell values (strings, numbers, booleans, dates) and support for multi-sheet workbooks.

## Quick example

Read every sheet from an existing workbook:

```ts
import { ExcelReader, ExcelSheetData } from 'typesexcel';

const reader = new ExcelReader('en');
const sheets: ExcelSheetData[] = reader.readSheets('data.xlsx');

for (const sheet of sheets) {
    console.log('Sheet:', sheet.name);
    for (const row of sheet.rows) {
        console.log(row.join('\t'));
    }
}
```

Or work with a typed `ExcelWorkbook` (headers separated from data rows):

```ts
import { ExcelReader, ExcelWorkbook, ExcelSheet } from 'typesexcel';

const reader = new ExcelReader('en');
const workbook: ExcelWorkbook = reader.readWorkbook('data.xlsx');

for (const sheet of workbook.getSheets()) {
    console.log('Sheet:', sheet.getName());
    console.log('Headers:', sheet.getHeaders());
    for (const row of sheet.getRows()) {
        console.log(row);
    }
}
```

Write a workbook with multiple sheets and mixed cell types:

```ts
import { ExcelSheet, ExcelWorkbook, ExcelWriter } from 'typesexcel';

const employees = new ExcelSheet(
    'Employees',
    ['Name', 'Department', 'Hired', 'Active'],
    [
        ['Alice', 'Engineering', new Date('2021-03-15'), true],
        ['Bob',   'Marketing',   new Date('2022-07-01'), false]
    ]
);

const offices = new ExcelSheet(
    'Offices',
    ['City', 'Headcount'],
    [
        ['Madrid', 12],
        ['Berlin', 7]
    ]
);

const workbook = new ExcelWorkbook([employees, offices]);

const writer = new ExcelWriter('en');
writer.writeFile('output.xlsx', workbook);
```

`ExcelWriter.writeFile()` also accepts a single `ExcelSheet` directly when only one sheet is needed.

## Why TypesExcel

- No native bindings — runs anywhere Node.js runs
- Lightweight: reads only what is needed from the ZIP (shared strings, sheet XML, relationships, styles)
- Strongly typed API — cell values are `string | number | boolean | Date | null`
- Full multi-sheet support for both reading and writing
- Built on [TypesXML](https://github.com/maxprograms-com/TypesXML), a strict, W3C-validated XML parser
- Designed for automation and tooling scenarios (data extraction, report generation, localization pipelines)

## Features

- **Read sheets** — Extract all worksheets from an `.xlsx` file, by name, or as a typed `ExcelWorkbook`
- **Write workbooks** — Generate a new `.xlsx` file from one `ExcelSheet` or a multi-sheet `ExcelWorkbook`
- **Typed cell values** — Strings, numbers, booleans, and dates are read and written with their native type, not just as strings
- **Hidden sheets skipped** — Sheets marked `hidden` or `veryHidden` are excluded when reading
- **Column utilities** — Convert between column letters (`A`, `B`, `AA`) and zero-based indices
- **Localized error messages** — Built-in i18n support for `en`, `es`, and `fr`

## Installation

```bash
npm install typesexcel
```

## API

### `ExcelReader`

```ts
new ExcelReader(language: string)
```

`language` selects the locale (`en`, `es`, or `fr`) used for error messages thrown by the reader; it has no effect on the data that is read.

| Method | Description |
| ------ | ----------- |
| `readSheets(filePath: string): ExcelSheetData[]` | Reads all worksheets from the given `.xlsx` file. Returns an array of `ExcelSheetData` objects, one per sheet. |
| `readSheet(filePath: string, sheetName: string): ExcelSheetData` | Reads a single worksheet by name. Throws if the sheet does not exist. |
| `readWorkbook(filePath: string): ExcelWorkbook` | Reads all worksheets and returns them as a typed `ExcelWorkbook`, with the first row of each sheet used as headers. |
| `listSheets(filePath: string): string[]` | Returns the names of all visible worksheets without parsing their data. |

Hidden and very-hidden sheets are skipped by all of the methods above.

### `ExcelSheetData`

```ts
interface ExcelSheetData {
    name: string;        // worksheet name as it appears in the workbook
    rows: CellValue[][]; // all rows including the header row, each cell typed
}
```

### `CellValue`

```ts
type CellValue = string | number | boolean | Date | null;
```

Dates are recognized from the cell's number format (built-in and custom date formats) and converted to `Date` instances; empty cells are `null`.

### `ExcelSheet`

```ts
new ExcelSheet(name: string, headers: string[], rows: CellValue[][])
```

| Method | Description |
| ------ | ----------- |
| `getName(): string` | Returns the sheet name. |
| `getHeaders(): string[]` | Returns the header row. |
| `getRows(): CellValue[][]` | Returns the data rows (not including the header). |

### `ExcelWorkbook`

```ts
new ExcelWorkbook(sheets: ExcelSheet[] = [])
```

| Method | Description |
| ------ | ----------- |
| `addSheet(sheet: ExcelSheet): void` | Appends a sheet to the workbook. |
| `removeSheet(name: string): void` | Removes the sheet with the given name, if present. |
| `getSheet(name: string): ExcelSheet \| undefined` | Returns the sheet with the given name. |
| `getSheets(): ExcelSheet[]` | Returns all sheets in the workbook. |

### `ExcelWriter`

```ts
new ExcelWriter(language: string)
```

`language` selects the locale (`en`, `es`, or `fr`) used for error messages thrown by the writer; it has no effect on the workbook that is generated.

| Method | Description |
| ------ | ----------- |
| `writeFile(filePath: string, data: ExcelSheet \| ExcelWorkbook): void` | Writes a workbook to `filePath`. Accepts either a single `ExcelSheet` or a multi-sheet `ExcelWorkbook`. |

Strings, numbers, booleans, and dates are written with the correct cell type and, for dates, a date number format.

### `ColumnUtils`

| Method | Description |
| ------ | ----------- |
| `columnName(index: number): string` | Converts a zero-based column index to a column letter (`0` → `A`, `26` → `AA`). |
| `columnIndex(name: string): number` | Converts a column letter to a zero-based index (`A` → `0`, `AA` → `26`). |

## License

[Eclipse Public License 1.0](LICENSE)
