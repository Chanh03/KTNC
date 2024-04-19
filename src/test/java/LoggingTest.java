import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.edge.EdgeDriver;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileInputStream;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.concurrent.TimeUnit;

public class LoggingTest {
    private final String EXCEL_DIR = "D:\\Vanh\\Kiểm Thử Nâng Cao\\data\\";
    public String baseUrl = "https://joymall.vn/";
    public String driverPath = "D:\\Vanh\\Kiểm Thử Nâng Cao\\msedgedriver.exe";
    public WebDriver driver;
    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    private Map<String, Object[]> TestNGResult_Logging;
    private LinkedHashMap<String, String[]> dataLoginTest;

    private void readDataFromExcel() {
        try {
            dataLoginTest = new LinkedHashMap<String, String[]>();
            if (sheet != null) {
                Iterator<Row> rowIterator = sheet.iterator();
                DataFormatter dataformat = new DataFormatter();
                while (rowIterator.hasNext()) {
                    Row row = rowIterator.next();
                    if (row.getRowNum() >= 1) {
                        Iterator<Cell> cellIterator = row.cellIterator();
                        String key = "";
                        String username = "";
                        String password = "";
                        String expected = "";

                        while (cellIterator.hasNext()) {
                            Cell cell = cellIterator.next();
                            if (cell.getColumnIndex() == 0) {
                                key = dataformat.formatCellValue(cell);
                            } else if (cell.getColumnIndex() == 1) {
                                username = dataformat.formatCellValue(cell);
                            } else if (cell.getColumnIndex() == 2) {
                                password = dataformat.formatCellValue(cell);
                            } else if (cell.getColumnIndex() == 3) {
                                expected = dataformat.formatCellValue(cell);
                            }

                        }
                        String[] myArr = {username, password, expected};
                        dataLoginTest.put(key, myArr);
                    }
                }
            } else {
                System.out.println("Không tìm thấy sheet có tên 'sheet' read.");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @BeforeClass
    public void setUp() {
        try {
            TestNGResult_Logging = new LinkedHashMap<>();
            System.setProperty("webdriver.edge.driver", driverPath); // Đây là EdgeDriver
            driver = new EdgeDriver();
            driver.manage().window().maximize();
            driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
            workbook = new XSSFWorkbook(new FileInputStream(new File(EXCEL_DIR + "test_log.xlsx")));
            sheet = workbook.getSheet("Sheet1");
            if (sheet == null) {
                System.out.println("Không tìm thấy sheet có tên 'sheet'.");
            } else {
                readDataFromExcel();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @DataProvider(name = "loggingData")
    public Object[][] getLoggingData() {
        // Gọi phương thức để đọc dữ liệu từ file Excel
        readDataFromExcel();

        // Tạo mảng 2 chiều để lưu dữ liệu từ file Excel
        Object[][] registrationData = new Object[dataLoginTest.size()][4]; // số cột dữ liệu Excel

        // Duyệt qua dataLoginTest và gán dữ liệu vào mảng registrationData
        int rowIndex = 0;
        for (String key : dataLoginTest.keySet()) {
            String[] rowData = dataLoginTest.get(key);
            if (rowData.length >= 2) { // Đảm bảo mỗi dòng trong Excel có ít nhất 7 cột dữ liệu
                registrationData[rowIndex][0] = key; // Key
                registrationData[rowIndex][1] = rowData[0]; // Username
                registrationData[rowIndex][2] = rowData[1]; // Password
                registrationData[rowIndex][3] = rowData[2]; // Expected
            }
            rowIndex++;
        }

        return registrationData;
    }


    @Test(dataProvider = "loggingData", priority = 2)
    public void logTest(String key, String username, String password, String expected) {
        String testSteps = "";
        String errorMessage = "";
        String status = "";
        String actual = "";
        String combinedValues = "";
        driver.get(baseUrl);
        testSteps += "Bước 1 : Truy cập web : '" + baseUrl + "'\n";
        try {
            WebElement loginLink = driver.findElement(By.xpath("//a[@href='/account/login' and contains(text(), 'Đăng nhập')]"));
            loginLink.click();
            testSteps += "Bước 2 : Click nút 'Đăng Nhập'\n";
            WebElement usernameInput = driver.findElement(By.id("customer_email"));
            WebElement passwordInput = driver.findElement(By.id("customer_password"));
            usernameInput.sendKeys(username);
            testSteps += "Bước 3 : Nhập email : '" + username + "'\n";
            passwordInput.sendKeys(password);
            testSteps += "Bước 4 : Nhập mật khẩu : '" + password + "'\n";

            WebElement loginButton = driver.findElement(By.cssSelector("button.btn-login"));
            loginButton.click();
            testSteps += "Bước 5 : Click nút 'Đăng nhập'\n";

        } catch (Exception e) {
            errorMessage = "Login Failed";
        }

        // Kiểm tra xem có thông báo lỗi nào xuất hiện không
        boolean errorDisplayed = false;

        if (username.isEmpty() || password.isEmpty()) {
            errorMessage = "Register fail - Empty data";
            errorDisplayed = true;
        } else {
            try {
                WebElement errorElement = driver.findElement(By.cssSelector("div.form-signup.margin-bottom-15"));
                // Lấy nội dung của phần tử và trích xuất thông điệp lỗi
                errorMessage = errorElement.getText().trim();
                if (!errorMessage.isEmpty()) {
                    errorDisplayed = true;
                }
            } catch (org.openqa.selenium.NoSuchElementException e) {

            }
        }
        if (errorDisplayed) {
            actual = errorMessage;
        } else if (driver.getCurrentUrl().equalsIgnoreCase("https://joymall.vn/account")) {
            // Click vào menu account
            try {
                // Xử lý kết quả
                WebElement emailElement = driver.findElement(By.xpath("//div[@class='form-signup name-account m992']/p[contains(strong, 'Email:')]"));
                String emailText = emailElement.getText().trim();
                String[] parts = emailText.split(":");
                String emailOut = parts[1].trim();
                testSteps += "Bước 6 : Thực hiện kiểm tra dữ liệu\n";
                if (emailOut.equalsIgnoreCase(username)) {
                    actual = "Login successful account: " + username;
                    driver.findElement(By.xpath("/html/body/header/div/div/div/div[4]/ul/li[2]/div/small/a"))
                            .click();
                } else {
                    actual = "Tên đăng nhập không trùng";
                }
            } catch (Exception e) {
                actual = "Login Failed";
            }
        } else {
            actual = "Login Failed";
        }

        // Đưa kết quả vào map TestNGResult
        combinedValues = "Username: " + username + ", Password: " + password + "\n";
        System.out.println("ACTUAL TEST : " + actual);

        LocalDateTime myDateObj = LocalDateTime.now();
        DateTimeFormatter myFormatObj = DateTimeFormatter.ofPattern("HH:mm:ss | dd-MM-yyyy");
        String formattedDate = myDateObj.format(myFormatObj);
        String dateCheck = formattedDate; // Lấy thời gian kiểm tra
        String resultStatus = actual.equals(expected) ? "PASS" : "FAIL";
        TestNGResult_Logging.put(key, new Object[]{key, combinedValues, testSteps, actual, expected, resultStatus, dateCheck});
    }


    @AfterClass
    public void tearDown() {
        try {
            ExcelWriter.writeToExcel(EXCEL_DIR + "TestNG_Result_JoyMall_Logging.xlsx", "Sheet1", TestNGResult_Logging);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (driver != null) {
                driver.quit();
            }
        }
    }

}
