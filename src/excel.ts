/**
 * # Excel 
 * Convenience functions to access tables in Excel files.
 * Uses the {@link https://github.com/SheetJS/js-xlsx Sheet JS xlsx parser and writer}.
 * 
 */

/** */
import { Log }  from 'hsutil'; const log = new Log('Excel'); 
import XLSX     from 'xlsx';

import { WorkBook,
         WorkSheet,
         CellObject
       }            from 'xlsx/types';

type DataRow = Array<number | string | Date>;

/**
 * A structure describing an Excel table
 */
export interface TableStruct {
    names:string[];
    sheetName:string;
    headerRow:number;
    colIndex:string[]; 
}

/**
 * 
 */
export interface Table {
    columns:TableStruct;
    table:DataRow[];
}

export interface ExcelFile {
    getSheetNames:		() =>string[];
    getTableColumns:	(sheetName:string, startCol?:string, row?:number) => TableStruct;
    getRowsForTable:	(table:TableStruct, maxRows?:number) => DataRow[];
    getTable:			(sheetID:string|number, startCol?:string, startRow?:number) => Table;
    nextExcelColIndex:	(startCol?:string) => IterableIterator<string>;
    getCellValue:		(sheet:string|WorkSheet, col:string, row:number) => string;		
}

/**
 * reads and returns a promise for an {@link #/hsLog/hsNode.excelFile excel file}.
 * ```
 * {
 * 	  	{@link excel.File.getSheetNames getSheetNames},
 * 		{@link excel.File.getTableColumns getTableColumns},
 * 		{@link excel.File.getRowsForTable getRowsForTable},
 * 		{@link excel.File.getTable getTable},
 * 		{@link excel.File.nextExcelColIndex nextExcelColIndex},
 * 		{@link excel.File.getCellValue getCellValue}
 * }
 * ```
 * # Usage
 * ```
 * const excel = require('./hsNode.excel');
 * const excelFile = excel.excelFile('./aFile.xlsx');
 * ``` 
 * @param name the name of the Excel file to read
 * @returns an object of functions providing access to the contents of the excel file.
 */
export class Excel implements ExcelFile { 
	private wb:WorkBook;

    constructor(name?:string, options?:any) {
        if (name!==undefined) {
            this.readFile(name, options);
        }
    }
     
    public readFile(name:string, options?:any) {
        this.wb = XLSX.readFile(name, options);
    }

    public readData(name:any, options?:any) {
        this.wb = XLSX.read(name, options);
    }

    public get workbook() { return this.wb; }

	/**
	 * retrieves sheet names from a file
	 * @returns {[string]} an array of sheet names
	 */
	public getSheetNames():string[] {
		return this.wb.SheetNames;
	}

	/**
	 * getTableColumns retrieves an array of consecutive valid column names.
	 * @param sheetName the sheet name to retrieve cells from
	 * @param startCol the first column of the table; defaults to 'A'.
	 * @param row the row to iterate over; defaults to 1.
	 * @returns an excel table description
	 */
	public getTableColumns(sheetName:string, startCol='A', row=1):TableStruct {
        const sheet:WorkSheet = this.wb.Sheets[sheetName];
		return this.constructCol(sheetName, row, this.getConsecutiveColumnNames(sheet, row, startCol));
	}
	
	/**
	 * getRowsForTable returns a 2D array[r][c] of row values, where the columns match the provided 
	 * columns names. 
	 * @param table an array of column descriptors. 
	 * @param maxRows if specified, determines the maximum number of rows to scan for. 
	 * If omitted, iteration stops when the first row of empty values is encountered.
	 */
	public getRowsForTable(table:TableStruct, maxRows=0):DataRow[] {
		if (!table.sheetName) { throw new Error('illegal table parameter in getRowsForTable'); }
		const sheet:WorkSheet = this.wb.Sheets[table.sheetName];
		const result:DataRow[] = [];
		let row=0; 
		while (true) { try {
			let newRow = this.getRow(sheet, row+table.headerRow+1, table.colIndex);
            let filledCells = 0;
            // only return non-empty rows
            if (newRow.some((c:string) => c.length>0)) { result.push(newRow); }
            // if no maxRows specified: break upon first empty row
            else if (maxRows<=0) {  break; }
            row++;
            // if rows exceed maxRows: break;
			if (maxRows>0 && row>=maxRows) { break; }
        } catch(e) {
            log.error(`processing row ${row+table.headerRow+1} for sheet ${table.sheetName}: ${e}`);
            throw e;
        }}
		return result;
	}
	
    /**
     * **generator** for Excel column indices starting at startCol. 
     * Following 'Z' the next column generated is 'AA' and so on. The generator 
     * starts with producing startCol as first index.
     * # Usage
     * ```
     * for (col of file.nextExcelColIndex(startCol='Y') { 
     *    printf("%s, ", col);      // -> Y, Z, AA, AB
     *    if (col === 'AB')  { break; }
     * }
     * ```
     * @param startCol the first column index ('A', ....) to yield; defaults to 'A'
     */
    public* nextExcelColIndex(startCol='A'):IterableIterator<string> {
        function nextChar(c:string):string { return String.fromCharCode(c.charCodeAt(0) + 1); }
        
        let c = startCol;
        while (true) {
            yield c;
            if (c.length === 1) {
                c = (c < 'Z')? nextChar(c[0]) : 'AA';
            } else {
                var ch = nextChar(c[1]);
                c = (ch > 'Z')? nextChar(c[0])+'A' : c[0] + ch;
            }
        }
    }

	/**
	 * gets a table of values, starting at the startCol and startRow.
	 * The table includes all consecutive columns with valid names, and all consecutive
	 * rows with at least one valid cell value.
	 * @param sheetID the sheet name or index from which to get the table
	 * @param startCol determines the left edge of the table; defaults to 'A'
	 * @param startRow determines the top edge of the table; defaults to 1
	 * @returns a tuple of {columns, table} 
	 */
	public getTable(sheetID:string|number, startCol='A', startRow=1) {
        const sheetName = (typeof sheetID === 'string')? sheetID : this.getSheetNames()[sheetID];
		const columns:TableStruct = this.getTableColumns(sheetName, startCol, startRow);
		const table:DataRow[]     = this.getRowsForTable(columns);
		return {columns, table};
    }

    /**
	 * returns the value of a cell, or undefined
	 * @param sheet the sheet object or sheet name to retrieve cells from;
	 * @param col the column index ('A',...)
	 * @param row index (1,...)
	 * @returns the value of a cell, or undefined
	 */
	public getCellValue(sheet:string|WorkSheet, col:string, row:number):string {
		if (typeof sheet === 'string') { sheet = this.wb.Sheets[sheet]; } 
        let c:CellObject;
		if (sheet[col+row] && sheet[col+row].v!=='') { 
            c = sheet[col+row];
            let val = c.w!==undefined? c.w : c.v;
			if (c) { switch(c.t) {
				case 's': return (<string>val).replace(/,/g,';').replace(/[\n\r]+/g,' ').trim();
				case 'n': /* falls through */ 
				default: return c.w.replace(/,/g,'');
			}}
		}
		return ''; 
	}

    //----------- private methods ------------------
		
	/**
	 * **Generator**, yields consecutive cell values over a row
	 * @param sheet the sheet object or sheet name to retrieve cells from;
	 * @param row the row to iterate over
	 * @param colIterator iterable over columns;
	 * or an iterable that generates column indices.
	 */
	private* getCellValues(sheet:WorkSheet, row:number, colIterator:string[]) {
		for (let col of colIterator) {
			yield this.getCellValue(sheet, col, row); 
		}
	}
	
	/**
	 * **Generator**, yields consecutive column names as an 
	 * {col, name} object. 
	 * The generator exits when the first empty column name is encountered.
	 * @param sheet the sheet to scan
	 * @param row the row to scan
	 * @param startCol defaults to 'A'
	 */
	private* getConsecutiveColumnNames(sheet:WorkSheet, row:number, startCol:string) {
		for (let col of this.nextExcelColIndex(startCol)) {
			if (!this.getCellValue(sheet, col, row)) { break; }
			yield {col:col, name:this.getCellValue(sheet, col, row)}; 
		}
	}

	/**
	 * returns an array[c] of values from columns that match indices provided `columns`.
	 * @param sheet the sheet object or sheet name to retrieve cells from;
	 * @param row the row to iterate over
	 * @param columns a) an array of column names. b) an {from:'A', to:'Z'} object 
	 * @return array of column values in the row
	 */
	private getRow(sheet:WorkSheet, row:number, columns:string[]) {
		const result = [...this.getCellValues(sheet, row, columns)];
		return result;
	}

	/**
	 * returns the value of a cell, or undefined
	 * @param sheetName the sheet object or sheet name to retrieve cells from;
	 * @param row index (1,...)
     * @param it an iterator over columns
	 * @returns the value of a cell, or undefined
	 */
	private constructCol(sheetName:string, row:number, it:any):TableStruct {
		const result:TableStruct = {
			names:[],
			sheetName: sheetName,
			headerRow: row,
			colIndex:  <string[]>[]
		};
		for (let col of it) {
			result.names.push(col.name);
			result.colIndex.push(col.col);
		}
		return result;
	}

}

