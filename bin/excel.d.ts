import { DataRow } from 'hsdatab';
import { WorkSheet } from 'xlsx/types';
export interface TableStruct {
    names: string[];
    sheetName: string;
    headerRow: number;
    colIndex: string[];
}
export interface Table {
    columns: TableStruct;
    table: DataRow[];
}
export interface ExcelFile {
    getSheetNames: () => string[];
    getTableColumns: (sheetName: string, startCol?: string, row?: number) => TableStruct;
    getRowsForTable: (table: TableStruct, maxRows?: number) => DataRow[];
    getTable: (sheetID: string | number, startCol?: string, startRow?: number) => Table;
    nextExcelColIndex: (startCol?: string) => IterableIterator<string>;
    getCellValue: (sheet: string | WorkSheet, col: string, row: number) => string;
}
export declare class File implements ExcelFile {
    private workbook;
    constructor(name: string, options?: any);
    getSheetNames(): string[];
    getTableColumns(sheetName: string, startCol?: string, row?: number): TableStruct;
    getRowsForTable(table: TableStruct, maxRows?: number): DataRow[];
    nextExcelColIndex(startCol?: string): IterableIterator<string>;
    getTable(sheetID: string | number, startCol?: string, startRow?: number): {
        columns: TableStruct;
        table: (string | number | Date)[][];
    };
    getCellValue(sheet: string | WorkSheet, col: string, row: number): string;
    private getCellValues;
    private getConsecutiveColumnNames;
    private getRow;
    private constructCol;
}
