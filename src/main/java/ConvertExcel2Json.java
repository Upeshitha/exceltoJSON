/*
 Author: Eranda Upeshitha
 Modified date: 2021/02/06
*/
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.fasterxml.jackson.databind.ObjectMapper;

public class ConvertExcel2Json {

    public static void main(String[] args) {
        // Step 1: Read Excel File into Java List Objects
        List<Customer> customers = readExcelFile("E:\\User\\Downloads\\customers-1.xlsx");

        // Step 2: Write Java List Objects to JSON File
        writeObjects2JsonFile(customers, "customers.json");

        System.out.println("Done");
    }

    /**
     * Read Excel File into Java List Objects
     *
     * @param filePath
     * @return
     */
    private static List<Customer> readExcelFile(String filePath){
        try {
            FileInputStream excelFile = new FileInputStream(new File(filePath));
            Workbook workbook = new XSSFWorkbook(excelFile);

            Sheet sheet = workbook.getSheet("Customers");
            Iterator<Row> rows = sheet.iterator();

            List<Customer> lstCustomers = new ArrayList<Customer>();

            int rowNumber = 0;
            while (rows.hasNext()) {
                Row currentRow = rows.next();

                // skip header
                if(rowNumber == 0) {
                    rowNumber++;
                    continue;
                }

                Iterator<Cell> cellsInRow = currentRow.iterator();

                Customer cust = new Customer();

                int cellIndex = 0;
                while (cellsInRow.hasNext()) {
                    Cell currentCell = cellsInRow.next();

                    if(cellIndex==0) { // ID
                        cust.setId(String.valueOf(currentCell.getNumericCellValue()));
                    } else if(cellIndex==1) { // Name
                        cust.setName(currentCell.getStringCellValue());
                    } else if(cellIndex==2) { // Address
                        cust.setAddress(currentCell.getStringCellValue());
                    } else if(cellIndex==3) { // Age
                        cust.setAge((int) currentCell.getNumericCellValue());
                    }

                    cellIndex++;
                }

                lstCustomers.add(cust);
            }

            // Close WorkBook
            workbook.close();

            return lstCustomers;
        } catch (IOException e) {
            throw new RuntimeException("FAIL! -> message = " + e.getMessage());
        }
    }

    /**
     *
     * Convert Java Objects to JSON File
     *
     * @param customers
     * @param pathFile
     */
    private static void writeObjects2JsonFile(List<Customer> customers, String pathFile) {
        ObjectMapper mapper = new ObjectMapper();

        File file = new File(pathFile);
        try {
            // Serialize Java object info JSON file.
            mapper.writeValue(file, customers);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
