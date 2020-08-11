package utils;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.*;

public class CheckCountriesInDB {
    private static final String COUNTRY_FILE = "C:\\Users\\A707963\\Desktop\\ongoingCR's\\DSL_CR5183\\IMP_DOCS\\CountriesToBeCheckedInCountryTable.xlsx";
    private static int count = 0;
    public static void main(String[] args) {
        new CheckCountriesInDB().checkFileDataInDB();
    }

    private void checkFileDataInDB() {
            try (final Connection con = getMajorDBConnection();
                 final FileInputStream fis = new FileInputStream(new File(COUNTRY_FILE));
                  final XSSFWorkbook myWorkBook = new XSSFWorkbook(fis)
                  ) {
                final Sheet countrySheet = getCountrySheet("Country",myWorkBook);
                if(null != countrySheet) {
                    for(final Row row : countrySheet) {
                        if(null != row) {
                            final String countryName = String.valueOf(row.getCell(0));
                           // System.out.println(countryName);
                            if(null != countryName && countryName.length()>0) {
                                searchCountryInDB(countryName,con);
                            }
                        }
                    }
                }
                System.out.println("Total of "+count+" records not found in DB");
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            } catch (SQLException e) {
                e.printStackTrace();
            }
    }

    private  void searchCountryInDB(final String countryName, final Connection con) {
        //stem.out.println("Checking country "+countryName);
        try(final PreparedStatement ps = con.prepareStatement("select COUNTRY_NAME from COUNTRY where COUNTRY_NAME = ?")) {
            if(null != ps) {
                ps.setString(1,countryName.toUpperCase());
                try(final ResultSet rs = ps.executeQuery()) {
                    if(null != rs) {
                        if(rs.next()) {
                            final String dbCountry = rs.getString(1);
                           // System.out.println(dbCountry);
                            if(stringNotNullEmpty(dbCountry)) {
                                //System.out.println("countryName :: " +countryName+" from excel :: FOUND :: in Database "+dbCountry);
                            }
                        }
                        else {
                            count++;
                            System.out.println(count+" countryName :: " +countryName.toUpperCase()+" from excel !! NOT FOUND !! in Database ");
                        }
                    }
                }
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }
    }

    private  Sheet getCountrySheet(final String sheetName, final XSSFWorkbook myWorkBook) {
        Sheet countrySheet = null;
        if(null != myWorkBook) {
             countrySheet = myWorkBook.getSheet(sheetName);
        }
        return countrySheet;
    }


    private  Connection getMajorDBConnection() {
        Connection con = null;
        try{
            String URL = "jdbc:oracle:thin:@172.24.108.182:1525:scrregr";
            String USER = "DSL_TQC_MAJOR_SERVICES";
            String PASS = "Welcome#123";
            Class.forName("oracle.jdbc.driver.OracleDriver");
            con= DriverManager.getConnection(URL,USER,PASS);
        }
        catch (Exception e) {
            e.printStackTrace();
        }
        return con;
    }

    private  boolean stringNotNullEmpty(final String var) {
        return null != var && var.length()>0;
    }
}
