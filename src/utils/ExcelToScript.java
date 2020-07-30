package utils;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
/**
 * 
 * @author A707963 Note : Excel Sheet should be of the format xlsx and in that Sheet name MUST be the table name , Header of the table inside each sheet MUST be the column name, Also
 * Create new sheet even if existing sheet is present, remove the existing one and create a new one
 *
 */
public class ExcelToScript {
	private static final String EXCEL_FILE = "C:\\Users\\A707963\\Desktop\\db_update.xlsx";
	private static final String SQL_FILE ="C:\\Users\\A707963\\Desktop\\INSERT_SERVICE_4885_D2.sql";
	static String tableName;
	static String columnNames;
	static XSSFWorkbook myWorkBook;
	static FileInputStream fis;
	static StringBuilder script;
	static BufferedWriter bwr;
	public static void main(String[] args) throws IOException {
		try {
			fis = new FileInputStream(new File(EXCEL_FILE));
			myWorkBook = new XSSFWorkbook(fis);
			if (null != myWorkBook) {
				script = new StringBuilder();
				for (int i = 0; i < myWorkBook.getNumberOfSheets(); i++) {
					Sheet sheet = myWorkBook.getSheetAt(i);
					processSheet(sheet,script);
				}
				script.append("COMMIT;");
				bwr= new BufferedWriter(new FileWriter(new File(SQL_FILE)));
				bwr.write(script.toString());
				bwr.flush();
				System.out.println("SQL file generated "+SQL_FILE);
				//System.out.println(script);
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if(null !=bwr){
				bwr.close();
			}
			if (null != myWorkBook) {
				myWorkBook.close();
			}
			if (null != fis) {
				fis.close();
			}
		}
	}

	private static void processSheet(Sheet sheet,StringBuilder script) {
		if (null != sheet) {
			tableName = sheet.getSheetName();
			script.append("-----------------------------------"+tableName+"------------------------------------------------\n");
			for (Row row : sheet) {
				processRow(row,script);
			}
			
		}
	}

	

	@SuppressWarnings("deprecation")
	private static void processRow(Row row,StringBuilder script) {
		if (null != row) {
			if (0 == row.getRowNum()) { //Populating Column Names from Header
				columnNames = "(";
				for (int cell = 0; cell < row.getLastCellNum(); cell++) {
					Cell c = row.getCell(cell);
					if (null != c) {
						if(cell == row.getLastCellNum()-1) {
							columnNames = columnNames + c.getStringCellValue();
						}
						else {
							columnNames = columnNames + c.getStringCellValue()+ ",";
						}
					}
				}
				columnNames = columnNames +")";
			} else { //Populating values from 2nd row onwards
				script.append("INSERT INTO "+tableName+columnNames+" VALUES "+"(");
				for (int cell = 0; cell < row.getLastCellNum(); cell++) {
					Cell c = row.getCell(cell);
					if (null != c) {
						if(cell == row.getLastCellNum()-1) {
							if(CellType.NUMERIC==c.getCellType()) {
								script.append((int)c.getNumericCellValue());
							}
							else if(CellType.STRING==c.getCellType()) {
								script.append("'"+c.getStringCellValue()+"'");
							}
							
						}
						else {
							if(CellType.NUMERIC==c.getCellType()) { //0 for number
								script.append((int)c.getNumericCellValue()+" , ");
							}
							else if(CellType.STRING==c.getCellType()) { //1 for varchar
								script.append("'"+c.getStringCellValue()+"', ");
							}

						}
					}
					else {
						if(cell == row.getLastCellNum()-1) {
							script.append("null");
						}
						else {
							script.append("null, ");
						}
					}
				}
				script.append(");");
				script.append("\n");				
			}
		}
	}
}
