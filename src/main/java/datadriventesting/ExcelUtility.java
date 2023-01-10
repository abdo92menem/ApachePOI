package datadriventesting;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelUtility {
    public FileInputStream inputStream;
    public FileOutputStream outputStream;
    public XSSFWorkbook workbook;
    public XSSFSheet sheet;
    public XSSFRow row;
    public XSSFCell cell;
    public CellStyle style;
    String path = null;

    public ExcelUtility(String path) {
        this.path = path;
    }

    public int getRowCount(String sheetName) throws IOException {
        inputStream = new FileInputStream(path);
        workbook = new XSSFWorkbook(inputStream);
        sheet = workbook.getSheet(sheetName);
        int rowCount = sheet.getLastRowNum();
        workbook.close();
        inputStream.close();
        return rowCount;
    }

    public int getCellCount(String sheetName, int rowNumber) throws IOException {
        inputStream = new FileInputStream(path);
        workbook = new XSSFWorkbook(inputStream);
        sheet = workbook.getSheet(sheetName);
        row = sheet.getRow(rowNumber);
        int cellCount = row.getLastCellNum();
        workbook.close();
        inputStream.close();
        return cellCount;
    }

    public String getCellData(String sheetName, int rowNumber, int colNumber) throws IOException {
        inputStream = new FileInputStream(path);
        workbook = new XSSFWorkbook(inputStream);
        sheet = workbook.getSheet(sheetName);
        row = sheet.getRow(rowNumber);
        cell = row.getCell(colNumber);

        DataFormatter formatter = new DataFormatter();
        String data;

        try {
            data = formatter.formatCellValue(cell);
        } catch (Exception e) {
            data = "";
        }

        workbook.close();
        inputStream.close();

        return data;
    }

    public void setCellData(String sheetName, int rowNumber, int colNumber, String data) throws IOException {
        File xlFile = new File(path);

        // if the file is not existed
        if (!xlFile.exists()) {
            workbook = new XSSFWorkbook();
            outputStream = new FileOutputStream(path);
            workbook.write(outputStream);
        }

        inputStream = new FileInputStream(path);
        workbook = new XSSFWorkbook(inputStream);

        // if sheet not existed
        if (workbook.getSheetIndex(sheetName) == -1)
            workbook.createSheet(sheetName);

        sheet = workbook.getSheet(sheetName);

        // if row not existed
        if (sheet.getRow(rowNumber) == null)
            sheet.createRow(rowNumber);

        row = sheet.getRow(rowNumber);
        cell = row.createCell(colNumber);
        cell.setCellValue(data);

        outputStream = new FileOutputStream(path);
        workbook.write(outputStream);
        workbook.close();
        inputStream.close();
        outputStream.close();
    }

    public void fillGreenColor(String sheetName, int rowNumber, int colNumber) throws IOException{
        inputStream = new FileInputStream(path);
        workbook = new XSSFWorkbook(inputStream);
        sheet = workbook.getSheet(sheetName);
        row = sheet.getRow(rowNumber);
        cell = row.getCell(colNumber);

        style = workbook.createCellStyle();

        style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        cell.setCellStyle(style);

        outputStream = new FileOutputStream(path);
        workbook.write(outputStream);

        workbook.close();
        inputStream.close();
        outputStream.close();
    }

    public void fillRedColor(String sheetName, int rowNumber, int colNumber) throws IOException{
        inputStream = new FileInputStream(path);
        workbook = new XSSFWorkbook(inputStream);
        sheet = workbook.getSheet(sheetName);
        row = sheet.getRow(rowNumber);
        cell = row.getCell(colNumber);

        style = workbook.createCellStyle();

        style.setFillForegroundColor(IndexedColors.RED.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        cell.setCellStyle(style);

        outputStream = new FileOutputStream(path);
        workbook.write(outputStream);

        workbook.close();
        inputStream.close();
        outputStream.close();
    }
}
