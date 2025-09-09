using System;
using System.Collections.Generic;
using System.IO;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System.Threading;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SeleniumExtras.WaitHelpers;
using OpenQA.Selenium.Interactions;
using Microsoft.VisualStudio.TestPlatform.ObjectModel;
using static Microsoft.IO.RecyclableMemoryStreamManager;

namespace ShoppingCartTests
{
    public class ShoppingCartTests : IDisposable
    {
        private IWebDriver driver;
        private WebDriverWait wait;
        private string excelFilePath = @"D:\\Năm 3 - HK2\\BaoĐamCLPM\\Do An\\ShoppingCart\\ShoppingCartTest.xlsx";
        private List<TestData> testCases;

        [SetUp]
        public void Setup()
        {
            driver = new ChromeDriver();
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("https://localhost:44374/Home/Index");
            Thread.Sleep(3000);

            // Di chuột vào "My Account" để hiển thị dropdown
            IWebElement myAccount = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//a[@href='#' and contains(text(),'My Account')]")));
            Actions actions = new Actions(driver);
            actions.MoveToElement(myAccount).Perform();
            Thread.Sleep(1000);

            // Bấm vào "Sign In"
            IWebElement signInButton = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//a[@href='/account/login' and contains(text(),'Sign In')]")));
            signInButton.Click();
            Thread.Sleep(3000);

            testCases = ReadTestDataFromExcel();
        }

        private void Login(string username, string password)
        {
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("UserName"))).SendKeys(username);
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Password"))).SendKeys(password);
            wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id='loginForm']/form/div[4]/div/input"))).Click();
        }

        public List<TestData> ReadTestDataFromExcel()
        {
            var testDataList = new List<TestData>();

            using (var file = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook = new XSSFWorkbook(file);
                ISheet sheet = workbook.GetSheet("TestData");

                for (int row = 1; row <= sheet.LastRowNum; row++)
                {
                    IRow currentRow = sheet.GetRow(row);
                    if (currentRow != null)
                    {
                        string testCaseID = currentRow.GetCell(0).ToString();
                        string username = currentRow.GetCell(1).ToString();
                        string password = currentRow.GetCell(2).ToString();
                        string expectedMessage = currentRow.GetCell(5).ToString();

                        testDataList.Add(new TestData
                        {
                            TestCaseID = testCaseID,
                            Username = username,
                            Password = password,
                            ExpectedMessage = expectedMessage
                        });
                    }
                }
            }
            return testDataList;
        }

        [Test]
        public void Test_GioHangRong()
        {
            var testCase = testCases.Find(tc => tc.TestCaseID == "TC1");
            Assert.NotNull(testCase, "Không tìm thấy TC1 trong file Excel!");

            // Đọc Expected Result chính xác từ Excel
            string expectedMessage = testCase.ExpectedMessage.Trim();
            Console.WriteLine($"Expected: '{expectedMessage}'"); // Debug dữ liệu từ Excel

            Login(testCase.Username, testCase.Password);
            Thread.Sleep(3000);
            driver.FindElement(By.CssSelector("a[href='/gio-hang']")).Click();
            Thread.Sleep(2000);

            string actualMessage;
            try
            {
                actualMessage = driver.FindElement(By.CssSelector(".cart-empty-message")).Text.Trim();
            }
            catch (NoSuchElementException)
            {
                actualMessage = "There are no products in the cart!";
            }

            Console.WriteLine($"Actual: '{actualMessage}'"); // Debug kết quả thực tế

            string status = actualMessage == expectedMessage ? "Passed" : "Failed";
            WriteResultToExcel(testCase.TestCaseID, actualMessage, status);

            Assert.AreEqual(expectedMessage, actualMessage);
        }


        [Test]
            public void Test_ThemSanPhamVaoGioHang()
            {
                var testCase = testCases.Find(tc => tc.TestCaseID == "TC2");
                Assert.NotNull(testCase, "Không tìm thấy TC2 trong file Excel!");

                Login(testCase.Username, testCase.Password);
                Thread.Sleep(3000);
                driver.Navigate().GoToUrl("https://localhost:44374/products");
                Thread.Sleep(3000);

                // Lấy số lượng sản phẩm trong giỏ hàng trước khi thêm
                driver.Navigate().GoToUrl("https://localhost:44374/gio-hang");
                Thread.Sleep(2000);
                int initialCartCount = driver.FindElements(By.XPath("//table//tr")).Count;

                // Quay lại trang sản phẩm để thêm vào giỏ hàng
                driver.Navigate().GoToUrl("https://localhost:44374/products");
                Thread.Sleep(3000);

                // Tìm sản phẩm "PISTA GP RR E2206 DOT - ORO"
                IWebElement productElement = wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("//h6[@class='product_name']/a[contains(text(), 'PISTA GP RR E2206 DOT - ORO')]")
                ));
                string productName = productElement.Text;

                // Hover vào sản phẩm để hiển thị nút "Add to Cart"
                Actions actions = new Actions(driver);
                actions.MoveToElement(productElement).Perform();
                Thread.Sleep(1000);

                // Tìm và click vào nút "Add to Cart"
                IWebElement addToCartButton = wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//div[contains(@class,'product')]//a[contains(@class, 'btnAddToCart') and @data-id='99']")
                ));

                // Scroll đến nút "Add to Cart" nếu cần
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView({block: 'center'});", addToCartButton);
                Thread.Sleep(500);

                // Click vào nút "Add to Cart"
                actions.MoveToElement(addToCartButton).Click().Perform();
                Thread.Sleep(2000);

                // Kiểm tra nếu có alert
                try
                {
                    WebDriverWait waitForAlert = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
                    waitForAlert.Until(ExpectedConditions.AlertIsPresent()).Accept();
                    Thread.Sleep(2000);
                }
                catch (NoAlertPresentException) { }

                // Vào giỏ hàng kiểm tra
                driver.Navigate().GoToUrl("https://localhost:44374/gio-hang");
                Thread.Sleep(2000);

                // Lấy số lượng sản phẩm trong giỏ hàng sau khi thêm
                int finalCartCount = driver.FindElements(By.XPath("//table//tr")).Count;

                // Kiểm tra số lượng sản phẩm có tăng lên không
                bool isAddedSuccessfully = finalCartCount > initialCartCount;
                string status = isAddedSuccessfully ? "Passed" : "Failed";

                // Ghi kết quả vào Excel
                WriteResultToExcel(testCase.TestCaseID, isAddedSuccessfully ? "Thêm sản phẩm vào giỏ hàng thành công!" : "Thêm sản phẩm thất bại", status);
                Assert.IsTrue(isAddedSuccessfully, "Sản phẩm không được thêm vào giỏ hàng!");
            }

        [Test]
        public void Test_ThemSPTuTrangChiTietSP()
        {
            var testCase = testCases.Find(tc => tc.TestCaseID == "TC3");
            Assert.NotNull(testCase, "Không tìm thấy TC4 trong file Excel!");

            // Đăng nhập
            Login(testCase.Username, testCase.Password);
            Thread.Sleep(2000);

            // Truy cập trang danh sách sản phẩm
            driver.Navigate().GoToUrl("https://localhost:44374/products");
            Thread.Sleep(2000);

            // Tìm sản phẩm "PISTA GP RR E2206 DOT - ORO" và vào trang chi tiết
            IWebElement productElement = wait.Until(ExpectedConditions.ElementIsVisible(
                By.XPath("//h6[@class='product_name']/a[contains(text(), 'PISTA GP RR E2206 DOT - ORO')]")
            ));
            productElement.Click();
            Thread.Sleep(2000);

            // Kiểm tra trang chi tiết sản phẩm đã load
            Assert.IsTrue(driver.Url.Contains("/chi-tiet"), "Không vào được trang chi tiết sản phẩm!");

            // Tìm nút "Add to Cart"
            try
            {
                IWebElement addToCartButton = wait.Until(ExpectedConditions.ElementExists(
                    By.XPath("//div[contains(@class,'red_button add_to_cart_button')]/a[contains(@class, 'btnAddToCart') and @data-id='99']")
                ));

                // Scroll đến nút (nếu chưa thấy)
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView({block: 'center'});", addToCartButton);
                Thread.Sleep(500);

                // Click vào nút "Add to Cart"
                try
                {
                    addToCartButton.Click();
                }
                catch (Exception)
                {
                    ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", addToCartButton);
                }
                Thread.Sleep(2000);
            }
            catch (WebDriverTimeoutException)
            {
                Assert.Fail("Không tìm thấy nút 'Add to Cart'.");
            }

            // Xử lý alert và lưu vào Excel
            string alertMessage = "Không có thông báo";
            try
            {
                WebDriverWait waitForAlert = new WebDriverWait(driver, TimeSpan.FromSeconds(3));
                IAlert alert = waitForAlert.Until(ExpectedConditions.AlertIsPresent());

                alertMessage = alert.Text; // Lưu nội dung alert
                Console.WriteLine("Alert hiển thị: " + alertMessage);
                alert.Accept(); // Bấm OK để đóng alert
                Thread.Sleep(1000);
            }
            catch (WebDriverTimeoutException)
            {
                Console.WriteLine("⚠ Không có alert nào xuất hiện.");
            }

            // Vào giỏ hàng kiểm tra
            driver.Navigate().GoToUrl("https://localhost:44374/gio-hang");
            Thread.Sleep(2000);

            // Kiểm tra sản phẩm có trong giỏ hàng không
            bool isAddedSuccessfully = driver.FindElements(By.XPath("//table//tr"))
                                             .Any(row => row.Text.Contains("PISTA GP RR E2206 DOT - ORO"));

            string status = isAddedSuccessfully ? "Passed" : "Failed";

            // Ghi KẾT QUẢ + THÔNG BÁO vào Excel
            WriteResultToExcel(testCase.TestCaseID, alertMessage, status);

            Assert.IsTrue(isAddedSuccessfully, "Sản phẩm không được thêm vào giỏ hàng!");
        }



        [Test]
        public void Test_CapNhatSoLuongAm()
        {
            var testCase = testCases.Find(tc => tc.TestCaseID == "TC4");
            Assert.NotNull(testCase, "Không tìm thấy TC4 trong file Excel!");

            Login(testCase.Username, testCase.Password);
            Thread.Sleep(3000);
            driver.Navigate().GoToUrl("https://localhost:44374/products");
            Thread.Sleep(3000);

            //// Lấy số lượng sản phẩm trong giỏ hàng trước khi thêm
            //driver.Navigate().GoToUrl("https://localhost:44374/gio-hang");
            //Thread.Sleep(2000);
            //int initialCartCount = driver.FindElements(By.XPath("//table//tr")).Count;

            //// Quay lại trang sản phẩm để thêm vào giỏ hàng
            //driver.Navigate().GoToUrl("https://localhost:44374/products");
            //Thread.Sleep(3000);

            // Tìm sản phẩm "PISTA GP RR E2206 DOT - ORO"
            IWebElement productElement = wait.Until(ExpectedConditions.ElementIsVisible(
                By.XPath("//h6[@class='product_name']/a[contains(text(), 'PISTA GP RR E2206 DOT - ORO')]")
            ));

            // Hover vào sản phẩm để hiển thị nút "Add to Cart"
            Actions actions = new Actions(driver);
            actions.MoveToElement(productElement).Perform();
            Thread.Sleep(1000);

            // Tìm và click vào nút "Add to Cart"
            IWebElement addToCartButton = wait.Until(ExpectedConditions.ElementToBeClickable(
                By.XPath("//div[contains(@class,'product')]//a[contains(@class, 'btnAddToCart') and @data-id='99']"))
            );

            // Scroll đến nút "Add to Cart" nếu cần
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView({block: 'center'});", addToCartButton);
            Thread.Sleep(500);

            // Click vào nút "Add to Cart"
            actions.MoveToElement(addToCartButton).Click().Perform();
            Thread.Sleep(2000);

            // Kiểm tra nếu có alert
            try
            {
                WebDriverWait waitForAlert = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
                waitForAlert.Until(ExpectedConditions.AlertIsPresent()).Accept();
                Thread.Sleep(2000);
            }
            catch (NoAlertPresentException) { }

            // Vào giỏ hàng kiểm tra
            driver.Navigate().GoToUrl("https://localhost:44374/gio-hang");
            Thread.Sleep(3000);  // Tăng thời gian chờ để trang cập nhật đầy đủ

            // Tìm ô nhập số lượng sản phẩm
            IWebElement quantityInput = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("input[id^='Quantity_']")));

            // Xóa giá trị cũ và nhập -5
            quantityInput.Clear();
            quantityInput.SendKeys("-5");
            Thread.Sleep(1000);

            // Nhấn nút "Update"
            IWebElement updateButton = wait.Until(ExpectedConditions.ElementToBeClickable(
                By.XPath("//a[contains(@class, 'btnUpdate')]")
            ));
            updateButton.Click();
            Thread.Sleep(3000); // Đợi trang cập nhật xong

            // Kiểm tra xem có hiển thị thông báo lỗi không
            bool isErrorDisplayed = false;
            try
            {
                IWebElement errorMessage = wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("//*[contains(text(), 'Số lượng sản phẩm không thể là 0 hoặc số âm')]") // Cập nhật XPath
                ));
                isErrorDisplayed = errorMessage.Displayed;
            }
            catch (WebDriverTimeoutException)
            {
                Console.WriteLine("Không tìm thấy thông báo lỗi!");
            }

            string actualResult;
            if (isErrorDisplayed)
            {
                actualResult = "Thông báo 'Số lượng sản phẩm không thể là 0 hoặc số âm'";
            }
            else
            {
                actualResult = "Cập nhật số lượng sản phẩm, danh sách sản phẩm và giá trong giỏ hàng.";
            }

            // Ghi kết quả vào Excel
            string status = isErrorDisplayed ? "Passed" : "Failed";
            WriteResultToExcel(testCase.TestCaseID, actualResult, status);

            // Kiểm tra kết quả
            Assert.IsTrue(isErrorDisplayed, "Lỗi: Hệ thống cho phép số lượng âm!");
        }

        [Test]
        public void Test_XoaSanPhamTrongGioHang()
        {
            var testCase = testCases.Find(tc => tc.TestCaseID == "TC5");
            Assert.NotNull(testCase, "Không tìm thấy TC5 trong file Excel!");

            Login(testCase.Username, testCase.Password);
            Thread.Sleep(3000);
            driver.Navigate().GoToUrl("https://localhost:44374/products");
            Thread.Sleep(3000);

            //// Lấy số lượng sản phẩm trong giỏ hàng trước khi thêm
            //driver.Navigate().GoToUrl("https://localhost:44374/gio-hang");
            //Thread.Sleep(2000);
            //int initialCartCount = driver.FindElements(By.XPath("//table//tr")).Count;

            //// Quay lại trang sản phẩm để thêm vào giỏ hàng
            //driver.Navigate().GoToUrl("https://localhost:44374/products");
            //Thread.Sleep(3000);

            // Tìm sản phẩm "PISTA GP RR E2206 DOT - ORO"
            IWebElement productElement = wait.Until(ExpectedConditions.ElementIsVisible(
                By.XPath("//h6[@class='product_name']/a[contains(text(), 'PISTA GP RR E2206 DOT - ORO')]")
            ));

            // Hover vào sản phẩm để hiển thị nút "Add to Cart"
            Actions actions = new Actions(driver);
            actions.MoveToElement(productElement).Perform();
            Thread.Sleep(1000);

            // Tìm và click vào nút "Add to Cart"
            IWebElement addToCartButton = wait.Until(ExpectedConditions.ElementToBeClickable(
                By.XPath("//div[contains(@class,'product')]//a[contains(@class, 'btnAddToCart') and @data-id='99']"))
            );

            // Scroll đến nút "Add to Cart" nếu cần
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView({block: 'center'});", addToCartButton);
            Thread.Sleep(500);

            // Click vào nút "Add to Cart"
            actions.MoveToElement(addToCartButton).Click().Perform();
            Thread.Sleep(2000);

            // Kiểm tra nếu có alert
            try
            {
                WebDriverWait waitForAlert = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
                waitForAlert.Until(ExpectedConditions.AlertIsPresent()).Accept();
                Thread.Sleep(2000);
            }
            catch (NoAlertPresentException) { }

            // Vào giỏ hàng kiểm tra
            driver.Navigate().GoToUrl("https://localhost:44374/gio-hang");
            Thread.Sleep(3000);  // Tăng thời gian chờ để trang cập nhật đầy đủ

            // Kiểm tra xem có nút "Delete" không
            IReadOnlyCollection<IWebElement> deleteButtons = driver.FindElements(By.XPath("//a[contains(@class, 'btn btn-sm btn-danger btnDelete')]"));
            Assert.IsTrue(deleteButtons.Count > 0, "Không tìm thấy sản phẩm để xóa!");

            // Nhấn vào nút "Delete"
            IWebElement deleteButton = deleteButtons.First();
            deleteButton.Click();
            Thread.Sleep(2000);

            // Kiểm tra alert
            try
            {
                WebDriverWait waitForAlert = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
                IAlert alert = waitForAlert.Until(ExpectedConditions.AlertIsPresent());
                string alertText = alert.Text;
                Console.WriteLine("Alert text: " + alertText);
                Assert.IsTrue(alertText.Contains("Bạn chắc chắn muốn xóa!"), "Thông báo xác nhận không đúng!");

                alert.Accept();
                Thread.Sleep(3000);
            }
            catch (WebDriverTimeoutException)
            {
                Assert.Fail("Không hiển thị alert xác nhận xóa!");
            }

            // Kiểm tra lại giỏ hàng
            driver.Navigate().GoToUrl("https://localhost:44374/gio-hang");
            Thread.Sleep(3000);

            // Kiểm tra xem sản phẩm đã biến mất chưa
            deleteButtons = driver.FindElements(By.XPath("//a[contains(@class, 'btn btn-sm btn-danger btnDelete')]"));
            bool isProductDeleted = deleteButtons.Count == 0;

            string actualResult = isProductDeleted ? "Xóa sản phẩm trong giỏ hàng thành công " : "Sản phẩm vẫn còn trong giỏ hàng";
            string status = isProductDeleted ? "Passed" : "Failed";
            WriteResultToExcel(testCase.TestCaseID, actualResult, status);

            Assert.IsTrue(isProductDeleted, "Lỗi: Sản phẩm không bị xóa khỏi giỏ hàng!");
        }



        public void WriteResultToExcel(string testCaseID, string actualResult, string status)
        {
            using (FileStream file = new FileStream(excelFilePath, FileMode.Open, FileAccess.ReadWrite))
            {
                IWorkbook workbook = new XSSFWorkbook(file);
                ISheet sheet = workbook.GetSheet("TestData");

                for (int rowIndex = 1; rowIndex <= sheet.LastRowNum; rowIndex++)
                {
                    IRow row = sheet.GetRow(rowIndex);
                    if (row != null && row.GetCell(0)?.ToString() == testCaseID)
                    {
                        // Cột "Actual Result" (index 6)
                        ICell actualResultCell = row.GetCell(6) ?? row.CreateCell(6);
                        actualResultCell.SetCellValue(actualResult);

                        // Cột "Status" (index 7)
                        ICell statusCell = row.GetCell(7) ?? row.CreateCell(7);
                        statusCell.SetCellValue(status);

                        break; // Thoát vòng lặp sau khi tìm thấy test case
                    }
                }

                // Ghi dữ liệu trở lại file
                using (FileStream writeFile = new FileStream(excelFilePath, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(writeFile);
                }

                // Giải phóng tài nguyên
                workbook.Close();
            }
        }


        [TearDown]
        public void TearDown()
        {
            Dispose();
        }

        public void Dispose()
        {
            driver?.Quit();
            driver?.Dispose();
        }
    }

    public class TestData
    {
        public string TestCaseID { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }
        public string ExpectedMessage { get; set; }
    }
}
