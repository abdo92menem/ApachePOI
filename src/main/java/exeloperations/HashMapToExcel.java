package exeloperations;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class HashMapToExcel {
    public static void main(String[] args) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Student Sheet");

        Map<String, String> data = new HashMap<>();
        data.put("101", "John");
        data.put("102", "Smith");
        data.put("103", "Scott");
        data.put("104", "Kim");
        data.put("105", "Mary");

        int rowNumber = 0;

        for (Map.Entry entry: data.entrySet()) {
            XSSFRow row = sheet.createRow(rowNumber++);
            row.createCell(0).setCellValue((String) entry.getKey());
            row.createCell(1).setCellValue((String) entry.getValue());
        }

        FileOutputStream outputStream = new FileOutputStream(".\\DataFiles\\Students.xlsx");
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();

        System.out.println("Excel Sheet written successfully...");
    }
}
