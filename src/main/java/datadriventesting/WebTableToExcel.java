package datadriventesting;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import java.io.IOException;
import java.time.Duration;

public class WebTableToExcel {
    public static void main(String[] args) throws IOException {
        System.setProperty("webdriver.chrome.driver", "Drivers\\chromedriver.exe");
        WebDriver driver = new ChromeDriver();

        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));
        driver.manage().window().maximize();

        driver.get("https://en.wikipedia.org/wiki/List_of_countries_and_dependencies_by_population");

        String path = ".\\DataFiles\\Population.xlsx";
        ExcelUtility excelUtility = new ExcelUtility(path);

        // Write header in excel sheet
        excelUtility.setCellData("Sheet1", 0, 0, "Country");
        excelUtility.setCellData("Sheet1", 0, 1, "Population");
        excelUtility.setCellData("Sheet1", 0, 2, "% of world");
        excelUtility.setCellData("Sheet1", 0, 3, "Date");
        excelUtility.setCellData("Sheet1", 0, 4, "Source");
        excelUtility.setCellData("Sheet1", 0, 5, "Notes");

        // Changing Header color to green
        excelUtility.fillGreenColor("Sheet1", 0, 0);
        excelUtility.fillGreenColor("Sheet1", 0, 1);
        excelUtility.fillGreenColor("Sheet1", 0, 2);
        excelUtility.fillGreenColor("Sheet1", 0, 3);
        excelUtility.fillGreenColor("Sheet1", 0, 4);
        excelUtility.fillGreenColor("Sheet1", 0, 5);

        // Capture Table rows
        WebElement table = driver.findElement(By.xpath("//*[@id=\"mw-content-text\"]/div[1]/table/tbody"));

        int rows = table.findElements(By.xpath("tr")).size();

        // Iterating through table to get table data
        for (int r = 1; r < rows; r++) {
            for (int c = 0; c < 6; c++) {
                WebElement td = table.findElement(By.xpath("tr[" + (r + 1) + "]/td[" + (c + 1) + "]"));
                excelUtility.setCellData("Sheet1", r, c, td.getText());
            }
        }
        System.out.println("Table Data imported successfully...");

        driver.close();
    }
}
