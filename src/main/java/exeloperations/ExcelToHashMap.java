package exeloperations;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class ExcelToHashMap {
    public static void main(String[] args) throws IOException {
        FileInputStream inputStream = new FileInputStream(".\\DataFiles\\Students.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
        XSSFSheet sheet = workbook.getSheet("Student Sheet");

        Map<String, String> data = new HashMap<>();

        int rowNumber = sheet.getLastRowNum();

        for (int r = 0; r <= rowNumber; r++) {
            XSSFRow row = sheet.getRow(r);
            data.put(row.getCell(0).getStringCellValue(), row.getCell(1).getStringCellValue());
        }

        System.out.println(data);

        workbook.close();
        inputStream.close();

        System.out.println("Hash Map read...");
    }
}
