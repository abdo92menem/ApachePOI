package exeloperations;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class WriteFormulaDataToCell {
    public static void main(String[] args) throws IOException {

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Numbers");

        XSSFRow row = sheet.createRow(0);

        row.createCell(0).setCellValue(10);
        row.createCell(1).setCellValue(20);
        row.createCell(2).setCellValue(30);

        row.createCell(3).setCellFormula("A1*B1*C1");

        FileOutputStream outputStream = new FileOutputStream(".\\DataFiles\\WriteFormula.xlsx");
        workbook.write(outputStream);

        outputStream.close();

        System.out.println("WriteFormula.xlsx created with formula cell...");
    }
}
