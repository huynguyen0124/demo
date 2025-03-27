using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.ComponentModel;
using OfficeOpenXml;
using System.Data;
using ExcelDataReader;
using System.Linq;


namespace TestScript
{
    public class TestTimKiemPhong
    {
        private IWebDriver driver;
        private WebDriverWait wait;
        private string excelFilePath;

        [SetUp]
        public void Setup()
        {
            driver = new ChromeDriver();
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));

            driver.Navigate().GoToUrl("http://localhost:3000/login");
            driver.Manage().Window.Maximize();
            // Nhập thông tin đăng nhập
            IWebElement usernameInput = driver.FindElement(By.Id("email"));
            usernameInput.SendKeys("admin@gmail.com");
            IWebElement passwordInput = driver.FindElement(By.Id("password"));
            passwordInput.SendKeys("minhhuy24001");

            IWebElement loginButton = driver.FindElement(By.Id("button"));
            loginButton.Click();
            wait.Until(d => d.Url.Contains("/dashboard") || d.Url.Contains("/Owner"));
        }

        public class ExcelDataProvider
        { 
            private static DataTable _excelDataTable;

            private static DataTable ReadExcel(string filePath)
            {
                if (_excelDataTable != null)
                {
                    return _excelDataTable;
                }
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        DataSet dataSet = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                            {
                                UseHeaderRow = true
                            }
                        });

                        _excelDataTable = dataSet.Tables[0];
                        return _excelDataTable;
                    }
                }
            }

            public static IEnumerable<TestCaseData> GetTestCasesFromExcel(string filePath)
            {
                var testCases = new List<TestCaseData>();

                if (!File.Exists(filePath))
                    throw new FileNotFoundException($"File '{filePath}' không tồn tại!");

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    int rowCount = worksheet.Dimension.Rows;

                    for (int row = 2; row <= rowCount; row++) // Bỏ qua header
                    {
                        int rowIndex = row - 2; //  Đảm bảo lưu chỉ số dòng
                        string name = worksheet.Cells[row, 2].Text;
                        string email = worksheet.Cells[row, 3].Text;
                        string birthday = worksheet.Cells[row, 4].Text;
                        string phone = worksheet.Cells[row, 5].Text;
                        string idencard = worksheet.Cells[row, 6].Text;
                        string avatar = worksheet.Cells[row, 7].Text;
                        bool expectedResult = worksheet.Cells[row, 8].Text.ToLower() == "true";

                        testCases.Add(new TestCaseData(rowIndex, name, email, birthday, phone, idencard, avatar, expectedResult));
                    }
                }
                return testCases;
            }

            private static int rowStart = 2; // Vị trí dòng bắt đầu ghi kết quả
            private static int colIndexActual = 9; // Cột kết quả thực tế

            public static void WriteResultToExcel(string filePath, string sheetName, int rowIndex, bool actuals, string result)
            {
                try
                {
                    Console.WriteLine($" Đang ghi kết quả vào file: {filePath}");
                    Console.WriteLine($" Sheet: {sheetName}, Row: {rowIndex}, Actual: {actuals}, Result: {result}");

                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($" File không tồn tại: {filePath}");
                        return;
                    }

                    using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetName];

                        if (worksheet == null)
                        {
                            Console.WriteLine($" Không tìm thấy sheet '{sheetName}'");
                            return;
                        }

                        int rowToWrite = rowStart + rowIndex;
                        Console.WriteLine($" Ghi giá trị {actuals} vào hàng {rowToWrite}, cột {colIndexActual}");
                        Console.WriteLine($" Ghi giá trị '{result}' vào hàng {rowToWrite}, cột {colIndexActual + 1}");

                        worksheet.Cells[rowToWrite, colIndexActual].Value = actuals;
                        worksheet.Cells[rowToWrite, colIndexActual].Value = result;

                        package.Save();
                        Console.WriteLine(" File đã được cập nhật!");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($" Lỗi khi ghi Excel: {ex.Message}");
                }
            }
        }
        static void Main()
        {
            IWebDriver driver = new ChromeDriver();
            driver.Navigate().GoToUrl("http://localhost:3000/");

            // Xử lý Textbox
            IWebElement firstNameTextbox = driver.FindElement(By.Name("ant-input-affix-wrapper css-dev-only-do-not-override-49qm ant-dropdown-trigger rounded-none h-full"));
            firstNameTextbox.SendKeys("Thành phố Hồ Chí Minh");

            // Xử lý Dropdown
            IWebElement monthDropdown = driver.FindElement(By.Name("ant-picker ant-picker-range css-dev-only-do-not-override-49qm rounded-none "));
            SelectElement monthSelect = new SelectElement(monthDropdown);
            monthSelect.SelectByText("January");

            IWebElement dayDropdown = driver.FindElement(By.Name("ant-picker ant-picker-range css-dev-only-do-not-override-49qm rounded-none"));
            SelectElement daySelect = new SelectElement(dayDropdown);
            daySelect.SelectByValue("15");

            IWebElement yearDropdown = driver.FindElement(By.Name("birthday_year"));
            SelectElement yearSelect = new SelectElement(yearDropdown);
            yearSelect.SelectByValue("1990");

            // Xử lý Button
            IWebElement signUpButton = driver.FindElement(By.Name("websubmit"));
            signUpButton.Click();

            // Đóng trình duyệt
            driver.Quit();
        }
            [TearDown]
            public void TearDown()
            {
                if (driver != null)
                {
                    driver.Quit();
                    driver.Dispose();
                }
            }
        }
    }
}

