import * as excel from './excel';

const TEST_FILE = './example/test.xlsx';

describe('excel', () => {
	it('should have excel defined', () => {
		expect(excel).toBeDefined();
	});
  
	describe('test file', () => {
		const NAME  = TEST_FILE;
		const SHEET = 'Closed';
		let file:any;
		
		beforeEach(() => {
			file = new excel.File(__dirname+'/'+NAME);
		});
		
		it('should have read '+NAME, () => {
			expect(file).toBeDefined();
		});	
		
		describe('sheets', () => {
			let sheets:any;
			
			beforeEach(() => { 
				sheets = file.getSheetNames(); 
			});
			
			it('should have 2 sheets', () => {
				expect(sheets.length).toBe(2);
			});
			
			it(`should have sheet "${SHEET}"`, () => {
				expect(sheets.indexOf(SHEET)).toBe(1);
			});			
		});
		
		describe('getCellValue', () => {
			it(`should have cell value`, () => {
				expect(file.getCellValue(SHEET, 'B', 5)).toBe('Ringo');
			});			
		});
		
		describe('header row', () => {
			it('should have a column name "Topic" on 4th position', () => {
				const columns = file.getTableColumns(SHEET, 'A', 1);
				expect(columns.names[2]).toBe('Topic');
			});
			it('should accept string as row number', () => {
				const columns = file.getTableColumns(SHEET, 'A', "1");
				expect(columns.names[2]).toBe('Topic'); 
			});
			it('should read table at default position A1', () => {
				const columns = file.getTableColumns(SHEET);
				expect(columns.names[2]).toBe('Topic');
			});
			
		}); 
		
		describe('getRowsForTable', () => {
            let err = '';
			it('should fail for illegal columns', () => {
				function noSheet() {
                    try { file.getRowsForTable({}); }
                    catch(e) { err = e.toString(); };
                     
                }
                noSheet();
				expect(err).not.toBe('');
			});
			
			it('should return maxRows', () => {
				let columns = file.getTableColumns(SHEET, 'A', 1);
				let rows = file.getRowsForTable(columns, 1);
				expect(rows.length).toBe(1);
			});
			
			it('should return less than maxRows=10', () => {
				let columns = file.getTableColumns(SHEET, 'A', 1);
				let rows = file.getRowsForTable(columns, 10);
				expect(rows.length).toBe(4);
			});
		});
		
		describe('table rows', () => {
			let columns:any;
			let rows:any;
			
			beforeEach(() => { 
				columns = file.getTableColumns(SHEET, 'A', 1);
				rows = file.getRowsForTable(columns);
			});
			
			it('should have 4 rows', () => {
				expect(rows.length).toBe(4);
			});
			
			it('should have "Start" value in 4th row', () => {
				let col = columns.names.indexOf('Start');
				expect(col).toBe(3);
				expect(rows[3][col]).toBe('03/01/14');
			});
		});
		
		describe('entire table', () => {			
			it('should have 5 columns', () => {
				let {columns} = file.getTable(SHEET, 'A', 1);
				expect(columns.names.length).toBe(5);
			});
			
			it('should have 4 rows', () => {
				let {table} = file.getTable(SHEET, 'A', 1);
				expect(table.length).toBe(4);
			});
			
			it('should have 4 rows in first sheet at default position', () => {
				let {table} = file.getTable(0);
				expect(table.length).toBe(4);
			});
		});
		
		describe('nextExcelColIndex', () => {
			function nextIndex(startCol?:string) { 
				const gen = file.nextExcelColIndex(startCol);
				gen.next(); // reproduces startCol;
				return gen.next().value;
			}
			
			it(`should produce column 'N' after column 'M'`, () => {
				expect(nextIndex('M')).toBe('N');
			});

			it(`should produce column 'AA' after column 'Z'`, () => {
				expect(nextIndex('Z')).toBe('AA');
			});

			it(`should produce column 'BA' after column 'AZ'`, () => {
				expect(nextIndex('AZ')).toBe('BA');
			});

			it(`should produce column 'BN' after column 'BM'`, () => {
				expect(nextIndex('BM')).toBe('BN');
			});

			it(`should produce column 'B' after default column`, () => {
				expect(nextIndex()).toBe('B');
			});
		});
	});	
});
