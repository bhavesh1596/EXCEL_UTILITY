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
 * Please note :: Scripts will be generated.. Please check the order of the execution. say 3 table scripts are generated.. depending on the primary key foreign key relation re arrange the order
 *jdk 8 needed for compilation ; apache poi jars
 * poi-4.1.2\poi-4.1.2.jar
 * poi-examples-4.1.2.jar
 * poi-excelant-4.1.2.jar
 * poi-ooxml-4.1.2.jar
 * poi-ooxml-schemas-4.1.2.jar
 * poi-scratchpad-4.1.2.jar
 */
public class ExcelToScriptV2 {
    private static final String EXCEL_FILE = "C:\\Users\\A707963\\Desktop\\ongoingCR's\\DSL_CR5183\\IMP_DOCS\\RATEPLAN_ADDON_2_5183.xlsx";
    private static final String SQL_FILE ="C:\\Users\\A707963\\Desktop\\insert_RATEPLAN_ADDON_5183_2.sql";
    private static final String ROLLBACK_SQL_FILE ="C:\\Users\\A707963\\Desktop\\rollback_RATEPLAN_ADDON_5183_2.sql";
    static String tableName;
    static String columnNames;
    static XSSFWorkbook myWorkBook;
    static FileInputStream fis;
    static StringBuilder script;
    static BufferedWriter bwr;
    static StringBuilder rollbackScript;
    static BufferedWriter rollbackBwr;
    static String firstColumn;
    static String secondColumn;
    public static void main(String[] args) throws IOException {
        try {
            fis = new FileInputStream(new File(EXCEL_FILE));
            myWorkBook = new XSSFWorkbook(fis);
            if (null != myWorkBook) {
                script = new StringBuilder();
                rollbackScript = new StringBuilder();
                for (int i = 0; i < myWorkBook.getNumberOfSheets(); i++) {
                    Sheet sheet = myWorkBook.getSheetAt(i);
                    processSheet(sheet,script,rollbackScript);
                }
                script.append("COMMIT;");
                rollbackScript.append("COMMIT;");
                bwr = new BufferedWriter(new FileWriter(new File(SQL_FILE)));
                rollbackBwr= new BufferedWriter(new FileWriter(new File(ROLLBACK_SQL_FILE)));
                bwr.write(script.toString());
                rollbackBwr.write(rollbackScript.toString());
                bwr.flush();
                rollbackBwr.flush();
                System.out.println("SQL file generated "+SQL_FILE);
                System.out.println("Rollback SQL file generated "+ROLLBACK_SQL_FILE);
                //System.out.println(script);
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if(null !=bwr){
                bwr.close();
            }
            if(null !=rollbackBwr){
                rollbackBwr.close();
            }
            if (null != myWorkBook) {
                myWorkBook.close();
            }
            if (null != fis) {
                fis.close();
            }
        }
    }

    private static void processSheet(Sheet sheet, StringBuilder script, StringBuilder rollbackScript) {
        if (null != sheet) {
            tableName = sheet.getSheetName();
            script.append("-----------------------------------"+tableName+"------------------------------------------------\n");
            rollbackScript.append("-----------------------------------"+tableName+"------------------------------------------------\n");
            for (Row row : sheet) {
                processRow(row,script,rollbackScript);
            }

        }
    }



    @SuppressWarnings("deprecation")
    private static void processRow(Row row, StringBuilder script, StringBuilder rollbackScript) {
        if (null != row) {
            if (0 == row.getRowNum()) { //Populating Column Names from Header
                columnNames = "(";
                for (int cell = 0; cell < row.getLastCellNum(); cell++) {
                    Cell c = row.getCell(cell);
                    if (null != c) {
                        if(0==cell) {
                            firstColumn = c.getStringCellValue();
                        }
                        if(1==cell) {
                            secondColumn = c.getStringCellValue();
                        }
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
                rollbackScript.append("DELETE FROM "+tableName+" WHERE "+firstColumn+"= ");
                for (int cell = 0; cell < row.getLastCellNum(); cell++) {
                    Cell c = row.getCell(cell);
                    if (null != c) {
                        if(0==cell) {
                            if(CellType.NUMERIC==c.getCellType()) {
                                rollbackScript.append((int)c.getNumericCellValue()+" AND "+secondColumn+"= ");
                            }
                            else if(CellType.STRING==c.getCellType()) {
                                rollbackScript.append("'"+c.getStringCellValue()+"'");
                            }
                        }
                        if(1==cell) {
                            if(CellType.NUMERIC==c.getCellType()) {
                                rollbackScript.append((int)c.getNumericCellValue()+";");
                            }
                            else if(CellType.STRING==c.getCellType()) {
                                rollbackScript.append("'"+c.getStringCellValue()+"';");
                            }
                        }
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
                        if(0==cell) {
                            rollbackScript.append("null AND "+secondColumn+"= ");
                        }
                        if(1==cell) {
                            rollbackScript.append("null;");
                        }
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
                rollbackScript.append("\n");
            }
        }
    }
}
