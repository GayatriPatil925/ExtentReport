using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;
using AventStack.ExtentReports.Reporter.Config;
using Microsoft.VisualBasic;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using NUnit.Framework.Interfaces;

namespace ExtentReport
{
    public class Tests
    {
        public ExtentReports extent;
        public ExtentSparkReporter reporter;//it's actual report file
        public ExtentTest testLog; //it's like Console.Writline for report file
        public IWebDriver driver;
        public IJavaScriptExecutor js;
        public Screenshot file;

        [OneTimeSetUp]
        public void OneTimeSetUp() 
        { 
            reporter =new ExtentSparkReporter(@"C:\Report\ProjectReport.html");
            reporter.Config.Theme = Theme.Dark;
            reporter.Config.DocumentTitle = "My Report";
            reporter.Config.ReportName="test";

            extent= new ExtentReports();
            extent.AttachReporter(reporter);//by using this will copy the properties of ExtentSparkReporter
            extent.AddSystemInfo("Environment", "QA");
            extent.AddSystemInfo("Tester", Environment.UserName);
            extent.AddSystemInfo("MachineName", Environment.MachineName);
        }
        [OneTimeTearDown]
        public void OneTimeTearDown()
        {
            extent.Flush();
        }
        [TearDown]
        public void TearDown() 
        {
            LoggingTestStatusExtentReport();
            driver.Dispose();
            driver.Quit();
        }
        [SetUp]
        public void Setup()
        {
            driver = new ChromeDriver();
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
            driver.Manage().Window.Maximize();
            driver.Url = "https://demoqa.com/accordian";
            js = (IJavaScriptExecutor)driver;
            testLog = extent.CreateTest(TestContext.CurrentContext.Test.Name); //it'll pickup method name and create entry for it in report
        }

        [Test]
        [Order(0)]
        public void VerifyBuyBoxPassTest()
        {
            IWebElement actualText = driver.FindElement(By.XPath("//div[@id='section1Heading']/following-sibling::div/div/p"));
            js.ExecuteScript("arguments[0].scrollIntoView(true);", actualText);
            testLog.Info("Actual text got for testing is: " + actualText.Text);
            string expectedText = "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with desktop publishing software like Aldus PageMaker including versions of Lorem Ipsum.";
            testLog.Info("Expected text for testing is: " + expectedText);
            Assert.IsTrue(actualText.Text == expectedText, "Text not matching.");
            testLog.Info("Text matched successfull.");
        }

        [Test]
        [Order(1)]
        public void VerifyBuyBoxFailTest()
        {
            IWebElement actualText = driver.FindElement(By.XPath("//div[@id='section1Heading']/following-sibling::div/div/p"));
            js.ExecuteScript("arguments[0].scrollIntoView(true);", actualText);
            testLog.Info("Actual text got for testing is: " + actualText.Text);
            string expectedText = "Lorem Ipsum is dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with desktop publishing software like Aldus PageMaker including versions of Lorem Ipsum.";
            testLog.Info("Expected text for testing is: " + expectedText);
            Assert.IsTrue(actualText.Text == expectedText, "Text not matching.");
            testLog.Info("Text matched successfull.");
        }
        private void LoggingTestStatusExtentReport()
        {
            try
            {
                var status = TestContext.CurrentContext.Result.Outcome.Status;//getting the result of the test method pass or fail 
                var stacktrace = string.Empty + TestContext.CurrentContext.Result.StackTrace + string.Empty;//detail error info with line no and any exception i.e throw
                var errorMessage = TestContext.CurrentContext.Result.Message;// it will pick output msg
                Status logstatus;
                switch (status)
                {
                    case TestStatus.Failed:
                        file = ((ITakesScreenshot)driver).GetScreenshot();//to take screenshot
                        file.SaveAsFile(@"C:\Report\ProjectReport.png");//it will save data in file
                        //the way to add screenshot to extent report
                        testLog.Info("Text verification failed.", MediaEntityBuilder.CreateScreenCaptureFromPath(@"C:\Report\ProjectReport.png").Build());
                        logstatus = Status.Fail;
                        testLog.Log(logstatus, "Test failed with stacktrace - " + stacktrace);
                        testLog.Log(logstatus, "Test ended with " + logstatus + " – " + errorMessage);
                        break;
                    case TestStatus.Skipped:
                        logstatus = Status.Skip;
                        testLog.Log(logstatus, "Test ended with " + logstatus);
                        break;
                    default:
                        logstatus = Status.Pass;
                        testLog.Log(logstatus, "Test steps finished for test case " + TestContext.CurrentContext.Test.Name);
                        testLog.Log(logstatus, "Test ended with " + logstatus);
                        break;
                }
            }
            catch (WebDriverException ex)
            {
                Console.WriteLine(ex.ToString());
            }

        }
    }
}