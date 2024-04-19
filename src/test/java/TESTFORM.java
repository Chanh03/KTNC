import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.edge.EdgeDriver;
import org.testng.Assert;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import java.util.List;
import java.util.concurrent.TimeUnit;

public class TESTFORM {
    public String baseUrl = "https://joymall.vn/";
    public String driverPath = "D:\\Vanh\\Kiểm Thử Nâng Cao\\msedgedriver.exe";
    public WebDriver driver;


    String ho = "";
    String ten = "";
    String sdt = "";
    String username = "";
    String password = "";
    @BeforeClass
    public void setUp() {
        System.setProperty("webdriver.edge.driver", driverPath); // Đây là EdgeDriver
        driver = new EdgeDriver();
        driver.manage().window().maximize();
        driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
    }

    @Test
    public void runer() {
        driver.get(baseUrl);
        driver.findElement(By.xpath("/html/body/header/div/div/div/div[4]/ul/li[2]/div/small/a")).click();
        driver.findElement(By.xpath("/html/body/section[2]/div/div/div[1]/p/a")).click();
        driver.findElement(By.id("lastName")).sendKeys(ho);
        driver.findElement(By.id("firstName")).sendKeys(ten);
        driver.findElement(By.id("Phone")).sendKeys(sdt);
        driver.findElement(By.id("email")).sendKeys(username);
        driver.findElement(By.id("password")).sendKeys(password);
        driver.findElement(By.xpath("//*[@id=\"customer_register\"]/div[2]/div[3]/button")).click();
        List<WebElement> inputElements = driver.findElements(By.cssSelector("input:required"));

        boolean foundError = false;

        for (WebElement element : inputElements) {
            String validationMessage = element.getAttribute("validationMessage");
            if (validationMessage != null && !validationMessage.isEmpty()) {
                System.out.println(validationMessage);
                foundError = true;
                break;
            }
        }


    }
    @AfterClass
    public void tearDown() {
        driver.close();
    }
}
