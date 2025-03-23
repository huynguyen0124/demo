using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using ExcelDataReader;
using OfficeOpenXml;
using OfficeOpenXml.Core.ExcelPackage;
using System.ComponentModel;

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

            driver.Navigate().GoToUrl("http://localhost:3000/search");
            driver.Manage().Window.Maximize();

            excelFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TestData", "testcaseSearch.xlsx");
        }

        [TearDown]
        public void TearDown()
        {
            driver.Quit();
        }

        public static class ExcelDataProvider
        {
            public static IEnumerable<TestCaseData> GetTestCasesFromExcel(string filePath, string sheetName)
            {
                var testCases = new List<TestCaseData>();

                if (!File.Exists(filePath))
                    throw new FileNotFoundException($"File '{filePath}' không tồn tại!");

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets[sheetName];
                    int rowCount = worksheet.Dimension.Rows;

                    for (int row = 2; row <= rowCount; row++)
                    {
                        int rowIndex = row - 2;
                        string province = worksheet.Cells[row, 2].Text;
                        string checkinDate = worksheet.Cells[row, 3].Text;
                        string checkoutDate = worksheet.Cells[row, 4].Text;
                        string guests = worksheet.Cells[row, 5].Text;
                        string roomType = worksheet.Cells[row, 6].Text;
                        string priceRange = worksheet.Cells[row, 7].Text;
                        string expectedResult = worksheet.Cells[row, 8].Text;

                        testCases.Add(new TestCaseData(rowIndex, province, checkinDate, checkoutDate, guests, roomType, priceRange, expectedResult));
                    }
                }
                return testCases;
            }

            public static void WriteResultToExcel(string filePath, string sheetName, int rowIndex, string actualResult)
            {
                try
                {
                    if (!File.Exists(filePath))
                        throw new FileNotFoundException($"File '{filePath}' không tồn tại!");

                    using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetName];

                        if (worksheet == null)
                            throw new Exception($"Sheet '{sheetName}' không tồn tại!");

                        int rowToWrite = 2 + rowIndex;
                        worksheet.Cells[rowToWrite, 9].Value = actualResult;
                        worksheet.Cells[rowToWrite, 10].Value = actualResult == worksheet.Cells[rowToWrite, 8].Text ? "Passed" : "Failed";

                        package.Save();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Lỗi khi ghi Excel: {ex.Message}");
                }
            }
        }

        public static IEnumerable<TestCaseData> GetSearchTestCases()
        {
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TestData", "testcaseSearch.xlsx");
            return ExcelDataProvider.GetTestCasesFromExcel(filePath, "Search");
        }

        private void ClearAndType(By locator, string text)
        {
            var element = driver.FindElement(locator);
            element.Clear();
            element.SendKeys(text);
        }

        [Test, TestCaseSource(nameof(GetSearchTestCases))]
        public void TestTimKiemPhong(int rowIndex, string province, string checkinDate, string checkoutDate, string guests, string roomType, string priceRange, string expectedResult)
        {
            try
            {
                // Nhập dữ liệu tìm kiếm
                if (!string.IsNullOrEmpty(province)) ClearAndType(By.Id("province"), province);
                if (!string.IsNullOrEmpty(checkinDate)) ClearAndType(By.Id("checkin"), checkinDate);
                if (!string.IsNullOrEmpty(checkoutDate)) ClearAndType(By.Id("checkout"), checkoutDate);
                if (!string.IsNullOrEmpty(guests)) ClearAndType(By.Id("guests"), guests);
                if (!string.IsNullOrEmpty(roomType)) ClearAndType(By.Id("roomType"), roomType);
                if (!string.IsNullOrEmpty(priceRange)) ClearAndType(By.Id("priceRange"), priceRange);

                // Nhấn nút tìm kiếm
                driver.FindElement(By.XPath("//button[contains(text(), 'Tìm kiếm')]")).Click();
                Thread.Sleep(2000);

                // Kiểm tra kết quả
                string actualResult = "";
                try
                {
                    // Nếu có danh sách phòng hiển thị
                    IWebElement resultsSection = wait.Until(d =>
                    {
                        var elements = d.FindElements(By.XPath("//div[contains(@class, 'room-card')]"));
                        return elements.Count > 0 ? elements[0] : null;
                    });

                    actualResult = resultsSection != null ? "Hiển thị danh sách phòng phù hợp" : "Không có phòng phù hợp";
                }
                catch (WebDriverTimeoutException)
                {
                    // Nếu có thông báo lỗi
                    try
                    {
                        actualResult = driver.FindElement(By.CssSelector(".error-message")).Text;
                    }
                    catch (NoSuchElementException)
                    {
                        actualResult = "Không có kết quả";
                    }
                }

                // Ghi kết quả vào file Excel
                ExcelDataProvider.WriteResultToExcel(excelFilePath, "Search", rowIndex, actualResult);

                // Kiểm tra kết quả mong đợi
                Assert.That(actualResult, Is.EqualTo(expectedResult), $"Expected: {expectedResult}, but got: {actualResult}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Lỗi xảy ra: " + ex.Message);
                Assert.Fail("Test case gặp lỗi: " + ex.Message);

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
}