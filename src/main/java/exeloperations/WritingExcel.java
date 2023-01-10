package exeloperations;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public class WritingExcel {
    public static void main(String[] args) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Emp Info");

        ArrayList<Object[]> empData = new ArrayList<>();

        empData.add(new Object[] {"EmpID", "Name", "Job"});
        empData.add(new Object[] {101, "David", "Engineer"});
        empData.add(new Object[] {102, "Smith", "Manager"});
        empData.add(new Object[] {103, "Scott", "Analyst"});

//        Object[][] empData = {
//                {"EmpID", "Name", "Job"},
//                {101, "David", "Engineer"},
//                {102, "Smith", "Manager"},
//                {103, "Scott", "Analyst"}
//        };

//        int rows = empData.length;
//        int cols = empData[0].length;
//
//        System.out.println(rows);
//        System.out.println(cols);
//
//        for (int r = 0; r < rows; r++) {
//            XSSFRow row = sheet.createRow(r);
//
//            for (int c = 0; c < cols; c++) {
//                XSSFCell cell = row.createCell(c);
//                Object value = empData[r][c];
//
//                if (value instanceof String)
//                    cell.setCellValue((String) value);
//                if (value instanceof Integer)
//                    cell.setCellValue((Integer) value);
//                if (value instanceof Boolean)
//                    cell.setCellValue((Boolean) value);
//            }
//        }

        // For Each loop

        int rowCount = 0;

        for (Object[] emp : empData) {
            XSSFRow row = sheet.createRow(rowCount++);

            int colCount = 0;

            for (Object value : emp) {
                XSSFCell cell = row.createCell(colCount++);

                if (value instanceof String)
                    cell.setCellValue((String) value);
                if (value instanceof Integer)
                    cell.setCellValue((Integer) value);
                if (value instanceof Boolean)
                    cell.setCellValue((Boolean) value);
            }
        }

        String filePath = ".\\DataFiles\\Employee.xlsx";
        FileOutputStream outputStream = new FileOutputStream(filePath);
        workbook.write(outputStream);

        outputStream.close();

        System.out.println("Employee.xlsx file written successfully...");
    }
}
