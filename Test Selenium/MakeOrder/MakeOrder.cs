using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using OpenQA.Selenium.Interactions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Microsoft.VisualStudio.TestPlatform.ObjectModel;


namespace NUnitAutomation
{
    [TestFixture]
    public class MakeOrderTest : IDisposable
    {
        private ChromeDriver? driver;
        private WebDriverWait? wait;
        private static readonly string excelFilePath = @"D:\BDCLPM\DoAn\MakeOrderTest.xlsx";
        private List<TestData> testCases;

        [SetUp]
        public void Setup()
        {
            driver = new ChromeDriver();
            driver.Manage().Window.Maximize();
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            // Đọc dữ liệu test từ file excel
            testCases = ReadTestDataFromExcel();
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
                        string testCaseID = currentRow.GetCell(0)?.ToString() ?? string.Empty;
                        string username = currentRow.GetCell(1)?.ToString() ?? string.Empty;
                        string password = currentRow.GetCell(2)?.ToString() ?? string.Empty;
                        string expectedMessage = currentRow.GetCell(3)?.ToString() ?? string.Empty;

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
        public void Test1_DatHangDeTrongHoTen()
        {
            driver!.Navigate().GoToUrl("https://localhost:44374/Home/Index");
            Assert.That(driver.Title, Does.Contain("HelmetShop"));

            var testCase = testCases.Find(tc => tc.TestCaseID == "TC1");
            Assert.That(testCase, Is.Not.Null, "Không tìm thấy TC1 trong file Excel!");

            Actions actions = new Actions(driver);
            IWebElement myAccount = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//a[contains(text(),'My Account')]")));

            // Hover chuột để mở dropdown
            actions.MoveToElement(myAccount).Perform();

            // Đợi 'Sign In' xuất hiện rồi mới click
            IWebElement signIn = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//a[contains(text(),'Sign In')]")));

            // Click vào 'Sign In'
            signIn.Click();

            // Đăng nhập
            wait.Until(ExpectedConditions.ElementExists(By.Id("UserName"))).SendKeys("Quynh");
            driver.FindElement(By.Id("Password")).SendKeys("123456");
            driver.FindElement(By.XPath("//*[@id='loginForm']/form/div[4]/div/input")).Click();

            // Tìm sản phẩm "PISTA GP RR E2206 DOT - ORO"
            IWebElement productElement = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//h6[@class='product_name']/a[contains(text(), 'PISTA GP RR E2206 DOT - ORO')]")));
            string productName = productElement.Text;

            // Hover vào sản phẩm để hiển thị nút "Add to Cart"
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
            Thread.Sleep(1000);

            // Nhấn vào nút Pay
            IWebElement payButton = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("/html/body/div[1]/div[3]/div/div/div[2]/div[2]/div/a[2]")));
            payButton.Click();

            // Điền thông tin, bỏ trống Full Name
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id=\"myForm\"]/div[2]/input"))).SendKeys("0775015583");
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id=\"myForm\"]/div[3]/input"))).SendKeys("50 Lò Siêu");
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id=\"myForm\"]/div[4]/input"))).SendKeys("lntquynh30@gmail.com");

            // Chọn phương thức thanh toán
            IWebElement paymentDropdown = driver.FindElement(By.XPath("//*[@id='myForm']/div[5]/select"));
            SelectElement selectPayment = new SelectElement(paymentDropdown);
            paymentDropdown.Click(); // Nhấn mở dropdown
            IWebElement bankingOption = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id=\"myForm\"]/div[5]/select/option[2]")));
            bankingOption.Click(); // Chọn "Banking"

            // Click ra ngoài để đóng dropdown
            IWebElement pageBody = driver.FindElement(By.TagName("body"));
            pageBody.Click();

            // Nhấn nút đặt hàng
            driver.FindElement(By.XPath("//*[@id='myForm']/div[6]/button")).Click();

            // Kiểm tra thông báo lỗi
            string expectedMessage = "Bạn không được để trống trường này";

            // Tìm phần tử thông báo lỗi
            IWebElement actualMessageElement = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id=\"CustumerName-error\"]")));

            // Lấy nội dung thông báo lỗi
            string actualMessage = actualMessageElement.Text;
            Thread.Sleep(500);

            // Kiểm tra kết quả
            bool isAddedSuccessfully = actualMessage.Contains(expectedMessage);
            // Ghi vào Excel
            WriteResultToExcel(testCase.TestCaseID, isAddedSuccessfully ? "Pass" : "Fail", "Completed");
            // Kiểm tra Assert
            Assert.That(isAddedSuccessfully, Is.True, "Không được để trống trường này!");
        }

        [Test]
        public void Test2_DatHangDeTrongSDT()
        {
            driver!.Navigate().GoToUrl("https://localhost:44374/Home/Index");
            Assert.That(driver.Title, Does.Contain("HelmetShop"));

            var testCase = testCases.Find(tc => tc.TestCaseID == "TC2");
            Assert.That(testCase, Is.Not.Null, "Không tìm thấy TC2 trong file Excel!");
                
            Actions actions = new Actions(driver);
            IWebElement myAccount = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//a[contains(text(),'My Account')]")));

            // Hover chuột để mở dropdown
            actions.MoveToElement(myAccount).Perform();

            // Đợi 'Sign In' xuất hiện rồi mới click
            IWebElement signIn = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//a[contains(text(),'Sign In')]")));

            // Click vào 'Sign In'
            signIn.Click();

            // Đăng nhập
            wait.Until(ExpectedConditions.ElementExists(By.Id("UserName"))).SendKeys("Quynh");
            driver.FindElement(By.Id("Password")).SendKeys("123456");
            driver.FindElement(By.XPath("//*[@id='loginForm']/form/div[4]/div/input")).Click();

            // Tìm sản phẩm "PISTA GP RR E2206 DOT - ORO"
            IWebElement productElement = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//h6[@class='product_name']/a[contains(text(), 'PISTA GP RR E2206 DOT - ORO')]")));
            string productName = productElement.Text;

            // Hover vào sản phẩm để hiển thị nút "Add to Cart"
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
            Thread.Sleep(1000);

            // Nhấn vào nút Pay
            IWebElement payButton = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("/html/body/div[1]/div[3]/div/div/div[2]/div[2]/div/a[2]")));
            payButton.Click();

            // Điền thông tin, bỏ trống số điện thoại
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id=\"myForm\"]/div[1]/input"))).SendKeys("Lê Nguyễn Trúc Quỳnh");
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id=\"myForm\"]/div[3]/input"))).SendKeys("50 Lò Siêu");
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id=\"myForm\"]/div[4]/input"))).SendKeys("lntquynh30@gmail.com");

            // Chọn phương thức thanh toán
            IWebElement paymentDropdown = driver.FindElement(By.XPath("//*[@id='myForm']/div[5]/select"));
            SelectElement selectPayment = new SelectElement(paymentDropdown);
            paymentDropdown.Click(); // Nhấn mở dropdown
            IWebElement bankingOption = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id=\"myForm\"]/div[5]/select/option[2]")));
            bankingOption.Click(); // Chọn "Banking"

            // Click ra ngoài để đóng dropdown
            IWebElement pageBody = driver.FindElement(By.TagName("body"));
            pageBody.Click();

            // Nhấn nút đặt hàng
            driver.FindElement(By.XPath("//*[@id='myForm']/div[6]/button")).Click();

            // Kiểm tra thông báo lỗi
            string expectedMessage = "Bạn không được để trống trường này";

            // Tìm phần tử thông báo lỗi
            IWebElement actualMessageElement = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id=\"Phone-error\"]")));

            // Lấy nội dung thông báo lỗi
            string actualMessage = actualMessageElement.Text;
            Thread.Sleep(500);

            // Kiểm tra kết quả
            bool isAddedSuccessfully = actualMessage.Contains(expectedMessage);
            // Ghi vào Excel
            WriteResultToExcel(testCase.TestCaseID, isAddedSuccessfully ? "Pass" : "Fail", "Completed");
            // Kiểm tra Assert
            Assert.That(isAddedSuccessfully, Is.True, "Không được để trống trường này!");
        }

        [Test]
        public void Test3_DatHangDeTrongEmail()
        {
            driver!.Navigate().GoToUrl("https://localhost:44374/Home/Index");
            Assert.That(driver.Title, Does.Contain("HelmetShop"));

            var testCase = testCases.Find(tc => tc.TestCaseID == "TC3");
            Assert.That(testCase, Is.Not.Null, "Không tìm thấy TC3 trong file Excel!");

            Actions actions = new Actions(driver);
            IWebElement myAccount = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//a[contains(text(),'My Account')]")));

            // Hover chuột để mở dropdown
            actions.MoveToElement(myAccount).Perform();

            // Đợi 'Sign In' xuất hiện rồi mới click
            IWebElement signIn = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//a[contains(text(),'Sign In')]")));

            // Click vào 'Sign In'
            signIn.Click();

            // Đăng nhập
            wait.Until(ExpectedConditions.ElementExists(By.Id("UserName"))).SendKeys("Quynh");
            driver.FindElement(By.Id("Password")).SendKeys("123456");
            driver.FindElement(By.XPath("//*[@id='loginForm']/form/div[4]/div/input")).Click();

            // Tìm sản phẩm "PISTA GP RR E2206 DOT - ORO"
            IWebElement productElement = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//h6[@class='product_name']/a[contains(text(), 'PISTA GP RR E2206 DOT - ORO')]")));
            string productName = productElement.Text;

            // Hover vào sản phẩm để hiển thị nút "Add to Cart"
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
            Thread.Sleep(1000);

            // Nhấn vào nút Pay
            IWebElement payButton = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("/html/body/div[1]/div[3]/div/div/div[2]/div[2]/div/a[2]")));
            payButton.Click();

            // Điền thông tin, bỏ trống email
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id=\"myForm\"]/div[1]/input"))).SendKeys("Lê Nguyễn Trúc Quỳnh");
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id=\"myForm\"]/div[2]/input"))).SendKeys("0775015583");
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id=\"myForm\"]/div[3]/input"))).SendKeys("50 Lò Siêu");

            // Chọn phương thức thanh toán
            IWebElement paymentDropdown = driver.FindElement(By.XPath("//*[@id='myForm']/div[5]/select"));
            SelectElement selectPayment = new SelectElement(paymentDropdown);
            paymentDropdown.Click(); // Nhấn mở dropdown
            IWebElement bankingOption = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id=\"myForm\"]/div[5]/select/option[2]")));
            bankingOption.Click(); // Chọn "Banking"

            // Click ra ngoài để đóng dropdown
            IWebElement pageBody = driver.FindElement(By.TagName("body"));
            pageBody.Click();

            // Nhấn nút đặt hàng
            driver.FindElement(By.XPath("//*[@id='myForm']/div[6]/button")).Click();

            // Kiểm tra thông báo lỗi
            string expectedMessage = "Email không hợp lệ";

            // Tìm phần tử thông báo lỗi
            IWebElement actualMessageElement = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id=\"Email-error\"]")));

            // Lấy nội dung thông báo lỗi
            string actualMessage = actualMessageElement.Text;
            Thread.Sleep(500);

            // Kiểm tra kết quả
            bool isAddedSuccessfully = actualMessage.Contains(expectedMessage);
            // Ghi vào Excel
            WriteResultToExcel(testCase.TestCaseID, isAddedSuccessfully ? "Pass" : "Fail", "Completed");
            // Kiểm tra Assert
            Assert.That(isAddedSuccessfully, Is.True, "Email không hợp lệ");
        }

        [Test]
        public void Test4_DatHangNhapHoTenChuaSo()
        {
            driver!.Navigate().GoToUrl("https://localhost:44374/Home/Index");
            Assert.That(driver.Title, Does.Contain("HelmetShop"));

            var testCase = testCases.Find(tc => tc.TestCaseID == "TC4");
            Assert.That(testCase, Is.Not.Null, "Không tìm thấy TC4 trong file Excel!");

            Actions actions = new Actions(driver);
            IWebElement myAccount = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//a[contains(text(),'My Account')]")));

            // Hover chuột để mở dropdown
            actions.MoveToElement(myAccount).Perform();

            // Đợi 'Sign In' xuất hiện rồi mới click
            IWebElement signIn = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//a[contains(text(),'Sign In')]")));

            // Click vào 'Sign In'
            signIn.Click();

            // Đăng nhập
            wait.Until(ExpectedConditions.ElementExists(By.Id("UserName"))).SendKeys("Quynh");
            driver.FindElement(By.Id("Password")).SendKeys("123456");
            driver.FindElement(By.XPath("//*[@id='loginForm']/form/div[4]/div/input")).Click();

            // Tìm sản phẩm "PISTA GP RR E2206 DOT - ORO"
            IWebElement productElement = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//h6[@class='product_name']/a[contains(text(), 'PISTA GP RR E2206 DOT - ORO')]")));
            string productName = productElement.Text;

            // Hover vào sản phẩm để hiển thị nút "Add to Cart"
            actions.MoveToElement(productElement).Perform();
            Thread.Sleep(1000);

            // Tìm và click vào nút "Add to Cart"
            IWebElement addToCartButton = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[contains(@class,'product')]//a[contains(@class, 'btnAddToCart') and @data-id='99']")));

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
            Thread.Sleep(1000);

            // Nhấn vào nút Pay
            IWebElement payButton = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("/html/body/div[1]/div[3]/div/div/div[2]/div[2]/div/a[2]")));
            payButton.Click();

            // Điền thông tin, Họ Tên chứa số
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id=\"myForm\"]/div[1]/input"))).SendKeys("Ronaldo 7");
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id=\"myForm\"]/div[2]/input"))).SendKeys("0775015583");
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id=\"myForm\"]/div[3]/input"))).SendKeys("50 Lò Siêu");
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id=\"myForm\"]/div[4]/input"))).SendKeys("lntquynh30@gmail.com");

            // Chọn phương thức thanh toán
            IWebElement paymentDropdown = driver.FindElement(By.XPath("//*[@id='myForm']/div[5]/select"));
            SelectElement selectPayment = new SelectElement(paymentDropdown);
            paymentDropdown.Click(); // Nhấn mở dropdown
            IWebElement bankingOption = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id=\"myForm\"]/div[5]/select/option[2]")));
            bankingOption.Click(); // Chọn "Banking"

            // Click ra ngoài để đóng dropdown
            IWebElement pageBody = driver.FindElement(By.TagName("body"));
            pageBody.Click();

            // Nhấn nút đặt hàng
            driver.FindElement(By.XPath("//*[@id='myForm']/div[6]/button")).Click();

            // Kiểm tra lỗi họ tên chứa số
            try
            {
                IWebElement errorMessageElement = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("CustumerName-error")));
                string actualMessage = errorMessageElement.Text.Trim();

                if (actualMessage.Contains("Họ tên không được chứa số"))
                {
                    WriteResultToExcel(testCase.TestCaseID, "Pass", "Completed");
                    return; // Test Passed, kết thúc kiểm tra
                }
            }
            catch (WebDriverTimeoutException)
            {
                // Không tìm thấy thông báo lỗi, tiếp tục kiểm tra kết quả đặt hàng
            }

            // Nếu không có lỗi hiển thị, kiểm tra xem đơn hàng có được tạo không
            bool isOrderPlaced = driver.PageSource.Contains("Đặt hàng thành công");

            // Nếu đơn hàng được tạo thành công -> TEST FAIL
            WriteResultToExcel(testCase.TestCaseID, isOrderPlaced ? "Fail" : "Pass", "Completed");

            // Kiểm tra Assert
            Assert.That(isOrderPlaced, Is.False, "Hệ thống không kiểm tra lỗi họ tên chứa số!");

        }

        [Test]
        public void Test5_DatHangThanhCong()
        {
            driver!.Navigate().GoToUrl("https://localhost:44374/Home/Index");
            Assert.That(driver.Title, Does.Contain("HelmetShop"));

            var testCase = testCases.Find(tc => tc.TestCaseID == "TC5");
            Assert.That(testCase, Is.Not.Null, "Không tìm thấy TC5 trong file Excel!");

            Actions actions = new Actions(driver);
            IWebElement myAccount = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//a[contains(text(),'My Account')]")));

            // Hover chuột để mở dropdown
            actions.MoveToElement(myAccount).Perform();

            // Đợi 'Sign In' xuất hiện rồi mới click
            IWebElement signIn = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//a[contains(text(),'Sign In')]")));

            // Click vào 'Sign In'
            signIn.Click();

            // Đăng nhập
            wait.Until(ExpectedConditions.ElementExists(By.Id("UserName"))).SendKeys("Quynh");
            driver.FindElement(By.Id("Password")).SendKeys("123456");
            driver.FindElement(By.XPath("//*[@id='loginForm']/form/div[4]/div/input")).Click();

            // Tìm sản phẩm "PISTA GP RR E2206 DOT - ORO"
            IWebElement productElement = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//h6[@class='product_name']/a[contains(text(), 'PISTA GP RR E2206 DOT - ORO')]")));
            string productName = productElement.Text;

            // Hover vào sản phẩm để hiển thị nút "Add to Cart"
            actions.MoveToElement(productElement).Perform();
            Thread.Sleep(1000);

            // Tìm và click vào nút "Add to Cart"
            IWebElement addToCartButton = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[contains(@class,'product')]//a[contains(@class, 'btnAddToCart') and @data-id='99']")));

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
            Thread.Sleep(1000);

            // Nhấn vào nút Pay
            IWebElement payButton = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("/html/body/div[1]/div[3]/div/div/div[2]/div[2]/div/a[2]")));
            payButton.Click();

            // Điền thông tin, Họ Tên chứa số
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id=\"myForm\"]/div[1]/input"))).SendKeys("Lê Nguyễn Trúc Quỳnh");
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id=\"myForm\"]/div[2]/input"))).SendKeys("0775015583");
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id=\"myForm\"]/div[3]/input"))).SendKeys("50 Lò Siêu");
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id=\"myForm\"]/div[4]/input"))).SendKeys("lntquynh30@gmail.com");

            // Chọn phương thức thanh toán
            IWebElement paymentDropdown = driver.FindElement(By.XPath("//*[@id='myForm']/div[5]/select"));
            SelectElement selectPayment = new SelectElement(paymentDropdown);
            paymentDropdown.Click(); // Nhấn mở dropdown
            IWebElement bankingOption = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id=\"myForm\"]/div[5]/select/option[2]")));
            bankingOption.Click(); // Chọn "Banking"

            // Click ra ngoài để đóng dropdown
            IWebElement pageBody = driver.FindElement(By.TagName("body"));
            pageBody.Click();

            // Nhấn nút đặt hàng
            driver.FindElement(By.XPath("//*[@id='myForm']/div[6]/button")).Click();

            // Kiểm tra xem đơn hàng có được đặt thành công không
            bool isOrderPlaced = driver.FindElements(By.XPath("//*[@id='load_data']/h1")).Count > 0;
            // Ghi kết quả vào Excel
            WriteResultToExcel(testCase.TestCaseID, isOrderPlaced ? "Pass" : "Fail", "Completed");
            Thread.Sleep(2000); // Chờ thêm thời gian nếu hệ thống load chậm  
            // Kiểm tra kết quả bằng Assert
            Assert.That(isOrderPlaced, Is.True, "Đặt hàng không thành công!");
        }

        // Ghi kết quả vào Excel
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
        public void TearDown() => Dispose();
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