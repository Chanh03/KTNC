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

public class RegisterTest {
    private final String EXCEL_DIR = "D:\\Vanh\\Kiểm Thử Nâng Cao\\data\\";
    public String baseUrl = "https://joymall.vn/";
    public String driverPath = "D:\\Vanh\\Kiểm Thử Nâng Cao\\msedgedriver.exe";
    public WebDriver driver;
    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    private Map<String, Object[]> TestNGResult;
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
                        String ho = "";
                        String ten = "";
                        String sdt = "";
                        String username = "";
                        String password = "";
                        String expected = "";
                        String Status = "";
                        String Actual = "";
                        String DateCheck = "";

                        while (cellIterator.hasNext()) {
                            Cell cell = cellIterator.next();
                            if (cell.getColumnIndex() == 0) {
                                key = dataformat.formatCellValue(cell);
                            } else if (cell.getColumnIndex() == 1) {
                                ho = dataformat.formatCellValue(cell);
                            } else if (cell.getColumnIndex() == 2) {
                                ten = dataformat.formatCellValue(cell);
                            } else if (cell.getColumnIndex() == 3) {
                                sdt = dataformat.formatCellValue(cell);
                            } else if (cell.getColumnIndex() == 4) {
                                username = dataformat.formatCellValue(cell);
                            } else if (cell.getColumnIndex() == 5) {
                                password = dataformat.formatCellValue(cell);
                            } else if (cell.getColumnIndex() == 6) {
                                expected = dataformat.formatCellValue(cell);
                            }
                        }
                        String[] myArr = {ho, ten, sdt, username, password, expected, Status, Actual, DateCheck};
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
            TestNGResult = new LinkedHashMap<>();
            System.setProperty("webdriver.edge.driver", driverPath); // Đây là EdgeDriver
            driver = new EdgeDriver();
            driver.manage().window().maximize();
            driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
            workbook = new XSSFWorkbook(new FileInputStream(new File(EXCEL_DIR + "test_reg_3.xlsx")));
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

    @DataProvider(name = "registrationData")
    public Object[][] getRegistrationData() {
        // Gọi phương thức để đọc dữ liệu từ file Excel
        readDataFromExcel();

        // Tạo mảng 2 chiều để lưu dữ liệu từ file Excel
        Object[][] registrationData = new Object[dataLoginTest.size()][7]; // số cột dữ liệu Excel

        // Duyệt qua dataLoginTest và gán dữ liệu vào mảng registrationData
        int rowIndex = 0;
        for (String key : dataLoginTest.keySet()) {
            String[] rowData = dataLoginTest.get(key);
            if (rowData.length >= 2) { // Đảm bảo mỗi dòng trong Excel có ít nhất 7 cột dữ liệu
                registrationData[rowIndex][0] = key; // Key
                registrationData[rowIndex][1] = rowData[0]; // Ho
                registrationData[rowIndex][2] = rowData[1]; // Ten
                registrationData[rowIndex][3] = rowData[2]; // SDT
                registrationData[rowIndex][4] = rowData[3]; // Username
                registrationData[rowIndex][5] = rowData[4]; // Password
                registrationData[rowIndex][6] = rowData[5]; // Expected
            }
            rowIndex++;
        }

        return registrationData;
    }


    @Test(dataProvider = "registrationData")
    public void regTest(String key, String ho, String ten, String sdt, String username, String password, String expected) {
        String testSteps = "";
        String errorMessage = "";
        String actual;
        String combinedValues = "";
        driver.get(baseUrl);
        try {

            testSteps += "Bước 1 : Truy cập web : '" + baseUrl + "'\n";
            driver.findElement(By.xpath("/html/body/header/div/div/div/div[4]/ul/li[2]/div/small/a")).click();
            testSteps += "Bước 2 : Click nút 'Đăng Nhập'\n";
            driver.findElement(By.xpath("/html/body/section[2]/div/div/div[1]/p/a")).click();
            testSteps += "Bước 3 : Click 'Đăng ký tại đây'\n";
            driver.findElement(By.id("lastName")).sendKeys(ho);
            testSteps += "Bước 4 : Nhập họ : '" + ho + "'\n";
            driver.findElement(By.id("firstName")).sendKeys(ten);
            testSteps += "Bước 5 : Nhập tên : '" + ten + "'\n";
            driver.findElement(By.id("Phone")).sendKeys(sdt);
            testSteps += "Bước 6 : Nhập số điện thoại : '" + sdt + "'\n";
            driver.findElement(By.id("email")).sendKeys(username);
            testSteps += "Bước 7 : Nhập email : '" + username + "'\n";
            driver.findElement(By.id("password")).sendKeys(password);
            testSteps += "Bước 8 : Nhập mật khẩu : '" + password + "'\n";
            driver.findElement(By.xpath("//*[@id=\"customer_register\"]/div[2]/div[3]/button")).click();
            testSteps += "Bước 9 : Click nút 'Đăng ký'\n";
        } catch (org.openqa.selenium.NoSuchElementException e) {
            errorMessage = "Register Failed";
        }
        // Kiểm tra xem có thông báo lỗi nào xuất hiện không
        boolean errorDisplayed = false;
        if (ho.isEmpty() || ten.isEmpty() || sdt.isEmpty() || username.isEmpty() || password.isEmpty()) {
            errorMessage = "Register fail - Empty data";
            errorDisplayed = true;
        } else {
            try {
                WebElement errorMessageElement = driver.findElement(By.xpath("//div[@class='errors']"));
                if (errorMessageElement.isDisplayed()) {
                    errorMessage = errorMessageElement.getText();
                    errorDisplayed = true;
                }
            } catch (org.openqa.selenium.NoSuchElementException e) {

            }
        }

        if (errorDisplayed) {
            actual = errorMessage;
        } else {
            // Click vào menu account
            try {
                WebElement accountMenu = driver.findElement(By.xpath("/html/body/header/div/div/div/div[4]/ul/li[2]/div/a"));
                if (accountMenu != null) {
                    accountMenu.click();
                    testSteps += "Bước 10 : Click vào menu account\n";
                }
                WebElement emailElement = driver.findElement(
                        By.xpath("//div[@class='form-signup name-account m992']/p[contains(strong, 'Email:')]"));
                testSteps += "Bước 11 : Thực hiện kiểm tra dữ liệu\n";
                String emailText = emailElement.getText().trim();
                String[] parts = emailText.split(":");
                String emailOut = parts[1].trim();
                if (emailOut.equalsIgnoreCase(username)) {
                    actual = "Register successful account :" + username;
                    driver.findElement(By.xpath("/html/body/header/div/div/div/div[4]/ul/li[2]/div/small/a"))
                            .click();
                    driver.findElement(By.xpath("/html/body/header/div/div/div/div[4]/ul/li[2]/div/small/a"))
                            .click();
                    driver.findElement(By.xpath("/html/body/section[2]/div/div/div[1]/p/a")).click();
                    driver.findElement(By.id("lastName")).sendKeys(ho);
                    driver.findElement(By.id("firstName")).sendKeys(ten);
                    driver.findElement(By.id("Phone")).sendKeys(sdt);
                    driver.findElement(By.id("email")).sendKeys(username);
                    driver.findElement(By.id("password")).sendKeys(password);
                    driver.findElement(By.xpath("//*[@id=\"customer_register\"]/div[2]/div[3]/button")).click();
                } else {
                    actual = "Tên đăng nhập không trùng";
                }
            } catch (Exception e) {
                actual = "Register Failed";
            }
        }

        // Đặt kết quả vào map TestNGResult
        combinedValues = "Họ: " + (ho.isEmpty() ? "''" : "'" + ho + "'") + ",\nTên: " + (ten.isEmpty() ? "''" : "'" + ten + "'") + ",\nSố điện thoại: " + (sdt.isEmpty() ? "''" : "'" + sdt + "'") + ",\nUsername: " + (username.isEmpty() ? "''" : "'" + username + "'") + ",\nPassword: " + (password.isEmpty() ? "''" : "'" + password + "'");
        System.out.println("ACTUAL TEST : " + actual);

        LocalDateTime myDateObj = LocalDateTime.now();
        DateTimeFormatter myFormatObj = DateTimeFormatter.ofPattern("HH:mm:ss | dd-MM-yyyy");
        String dateCheck = myDateObj.format(myFormatObj);
        String resultStatus = actual.equals(expected) ? "PASS" : "FAIL";
        TestNGResult.put(key, new Object[]{key, combinedValues, testSteps, actual, expected, resultStatus, dateCheck});
        //        Assert.assertEquals(actual, expected, "Kết quả test FAIL");
    }

    @AfterClass
    public void tearDown() {
        try {
            ExcelWriter.writeToExcel(EXCEL_DIR + "TestNG_Result_JoyMall_Register.xlsx", "Sheet1", TestNGResult);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (driver != null) {
                driver.quit();
            }
        }
    }

}