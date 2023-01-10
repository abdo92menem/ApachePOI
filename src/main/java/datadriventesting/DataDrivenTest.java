package datadriventesting;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.Assert;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.IOException;
import java.time.Duration;

public class DataDrivenTest {

    WebDriver driver;

    @BeforeTest
    public void setup() {
        System.setProperty("webdriver.chrome.driver", "Drivers\\chromedriver.exe");
        driver = new ChromeDriver();
        driver.manage().timeouts().implicitlyWait(Duration.ofMillis(2000));
        driver.manage().window().maximize();
        driver.get("https://admin-demo.nopcommerce.com/login");
    }

    @Test(dataProvider = "LoginData")
    public void loginTest(String username, String password, String expected) {
        WebElement emailTxt = driver.findElement(By.id("Email"));
        emailTxt.clear();
        emailTxt.sendKeys(username);

        WebElement passwordTxt = driver.findElement(By.id("Password"));
        passwordTxt.clear();
        passwordTxt.sendKeys(password);

        driver.findElement(By.cssSelector("button.button-1.login-button")).click();

        String expectedTitle = "Dashboard / nopCommerce administration";
        String actualTitle = driver.getTitle();

        if (expected.equals("Valid")) {
            if (expectedTitle.equals(actualTitle)) {
                driver.findElement(By.linkText("Logout")).click();
                Assert.assertTrue(true);
            } else {
                Assert.assertTrue(false);
            }
        } else if (expected.equals("Invalid")) {
            if (expectedTitle.equals(actualTitle)) {
                driver.findElement(By.linkText("Logout")).click();
                Assert.assertTrue(false);
            } else {
                Assert.assertTrue(true);
            }
        }
    }

    @AfterClass
    public void tearDown() {
        driver.close();
    }

    @DataProvider(name = "LoginData")
    public Object[][] getData() throws IOException {
        /* String[][] loginData = {
                {"admin@yourstore.com", "admin", "Valid"},
                {"admin@yourstore.com", "adm", "Invalid"},
                {"adm@yourstore.com", "admin", "Invalid"},
                {"adm@yourstore.com", "adm", "Invalid"}
        };
        */

        // get the data from Excel File
        String path = ".\\DataFiles\\LoginData.xlsx";
        ExcelUtility excelUtility = new ExcelUtility(path);

        int totalRows = excelUtility.getRowCount("Sheet1");
        int totalCols = excelUtility.getCellCount("Sheet1", 1);

        String[][] loginData = new String[totalRows][totalCols];

        for (int r = 1; r <= totalRows; r++) {
            for (int c = 0; c < totalCols; c++) {
                loginData[r - 1][c] = excelUtility.getCellData("Sheet1", r, c);
            }
        }

        return loginData;
    }
}
