import * as excel from './excel';
import { o }      from 'hsutil/mithril';

const TEST_FILE = '../../example/test.xlsx';

o.spec("excel", () => {
	o('should have excel defined', () => {
		o(excel).notEquals(undefined);
	});
  
	o.spec("test file", () => {
		const NAME  = TEST_FILE;
		const SHEET = 'Closed';
		let file:any;
		
		o.beforeEach(() => {
			file = excel.readFile(__dirname+'/'+NAME);
		});
		
		o('should have read '+NAME, () => {
			o(file).notEquals(undefined);
		});	
		
		o.spec('sheets', () => {
			let sheets:any;
			
			o.beforeEach(() => { 
				sheets = file.getSheetNames(); 
			});
			
			o('should have 2 sheets', () => {
				o(sheets.length).equals(2);
			});
			
			o(`should have sheet "${SHEET}"`, () => {
				o(sheets.indexOf(SHEET)).equals(1);
			});			
		});
		
		o.spec('getCellValue', () => {
			o(`should have cell value`, () => {
				o(file.getCellValue(SHEET, 'B', 5)).equals('Ringo');
			});			
		});
		
		o.spec('header row', () => {
			o('should have a column name "Topic" on 4th position', () => {
				let columns = file.getTableColumns(SHEET, 'A', 1);
				o(columns.names[2]).equals('Topic');
			});
			
			o('should accept string as row number', () => {
				let columns = file.getTableColumns(SHEET, 'A', "1");
				o(columns.names[2]).equals('Topic'); 
			});
		}); 
		
		o.spec('getRowsForTable', () => {
            let err = '';
			o('should fail for illegal columns', (done:any) => {
				function noSheet(done:any) {
                    try { file.getRowsForTable({}); }
                    catch(e) { err = e.toString(); };
                    done(); 
                }
                noSheet(done);
				o(err).notEquals('');
			});
			
			o('should return maxRows', () => {
				let columns = file.getTableColumns(SHEET, 'A', 1);
				let rows = file.getRowsForTable(columns, 1);
				o(rows.length).equals(1);
			});
			
			o('should return less than maxRows=10', () => {
				let columns = file.getTableColumns(SHEET, 'A', 1);
				let rows = file.getRowsForTable(columns, 10);
				o(rows.length).equals(4);
			});
		});
		
		o.spec('table rows', () => {
			let columns:any;
			let rows:any;
			
			o.beforeEach(() => { 
				columns = file.getTableColumns(SHEET, 'A', 1);
				rows = file.getRowsForTable(columns);
			});
			
			o('should have 4 rows', () => {
				o(rows.length).equals(4);
			});
			
			o('should have "Start" value in 4th row', () => {
				let col = columns.names.indexOf('Start');
				o(col).equals(3);
				o(rows[3][col]).equals('03/01/14');
			});
		});
		
		o.spec('entire table', () => {			
			o('should have 5 columns', () => {
				let {columns} = file.getTable(SHEET, 'A', 1);
				o(columns.names.length).equals(5);
			});
			
			o('should have 4 rows', () => {
				let {table} = file.getTable(SHEET, 'A', 1);
				o(table.length).equals(4);
			});
		});
		
		o.spec('nextExcelColIndex', () => {
			function nextIndex(startCol:string) { 
				const gen = file.nextExcelColIndex(startCol);
				gen.next(); // reproduces startCol;
				return gen.next().value;
			}
			
			o("should produce column 'N' after column 'M'", () => {
				o(nextIndex('M')).equals('N');
			});

			o("should produce column 'AA' after column 'Z'", () => {
				o(nextIndex('Z')).equals('AA');
			});

			o("should produce column 'BA' after column 'AZ'", () => {
				o(nextIndex('AZ')).equals('BA');
			});

			o("should produce column 'BN' after column 'BM'", () => {
				o(nextIndex('BM')).equals('BN');
			});
		});
	});	
});
