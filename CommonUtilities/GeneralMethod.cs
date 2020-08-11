
using AventStack.ExtentReports;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using System.Reflection;
using Syncfusion.XlsIO;
using Newtonsoft.Json;
using System.Net;
using Microsoft.Extensions.Options;
using System.Diagnostics;


namespace LexBaseFramework.CommonUtilities
{
    public class GeneralMethod
    {
        public static IWebDriver driver;
        public IList<IWebElement> row = null;
        public IList<string> NotSorted = null;
        public IList<string> SortedDescending = null;
        public IList<string> SortedAscending = null;
        public ExtentTest test;
        private static Microsoft.Office.Interop.Excel.Workbook mWorkBook;
        private static Microsoft.Office.Interop.Excel.Sheets mWorkSheets;
        private static Microsoft.Office.Interop.Excel.Worksheet mWSheet1;
        private static Microsoft.Office.Interop.Excel.Application oXL;

        static DateTime TodaysDate;
        static DateTime FivedaysDate;

        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public string ScenarioName = string.Empty;
        public string TestCaseName = string.Empty;
        public static string browserName = string.Empty;
        Enum.LogStatus status;
        public Enum.LogStatus openBrowser(string bType)
        {
            ExtentTestManager._parentTest.Log(Status.Info, "Opening browser- " + bType);
            log.Info("Framework Initalized  !!! ");
            try
            {
                
                if (bType.Equals(Enum.BrowserName.firefox))
                    driver = new FirefoxDriver();
                else if (bType.Equals(Enum.BrowserName.chrome.ToString()))
                     driver = new ChromeDriver(GetDriversPath());
                else if (bType.Equals(Enum.BrowserName.ie.ToString()))
                    driver = new InternetExplorerDriver(GetDriversPath());
                log.Info("Opened the browser !!! ");
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(100);
                driver.Manage().Window.Maximize();
                log.Info("Maximize the browser !!! ");
                //driver.Manage().Cookies.DeleteAllCookies();
                //log.Info("Delete All Cookies from the browser !!! ");
                return Enum.LogStatus.Passed;
            }
            catch (Exception ex)
            {
                log.Info(ex.Message);
                return Enum.LogStatus.Failed;
            }
        }
        public Enum.LogStatus navigateprod()
        {
            try
            {
                //test.Log(Status.Info, "Navigating to - " + ConfigurationManager.AppSettings["URLProd"]);
                driver.Url = ConfigurationManager.AppSettings["URLProd"];
                log.Info("User Pass the URL!!! " + driver.Url);
                return Enum.LogStatus.Passed;
            }
            catch (Exception ex)
            {
                log.Info(ex.Message);
                return Enum.LogStatus.Failed;
            }

        }
        public Enum.LogStatus navigateTest()
        {
            try
            {
                //test.Log(Status.Info, "Navigating to - " + ConfigurationManager.AppSettings["URLTesting"]);
                driver.Url = ConfigurationManager.AppSettings["URLTesting"];
                log.Info("User Pass the URL!!! " + driver.Url);
                return Enum.LogStatus.Passed;
            }
            catch (Exception ex)
            {
                log.Info(ex.Message);
                return Enum.LogStatus.Failed;
            }

        }
        public Enum.LogStatus navigateStage()
        {
            try
            {
                //test.Log(Status.Info, "Navigating to - " + ConfigurationManager.AppSettings["URLStage"]);
                driver.Url = ConfigurationManager.AppSettings["URLStage"];
                log.Info("User Pass the URL!!! " + driver.Url);
                return Enum.LogStatus.Passed;
            }
            catch (Exception ex)
            {
                log.Info(ex.Message);
                return Enum.LogStatus.Failed;
            }

        }
       

        public void WaitforElement(int Seconds_time, int Pollingtime, string locatorValue)
        {
            try
            {
                PageLoadWait(20);
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(Seconds_time));
                wait.PollingInterval = TimeSpan.FromMilliseconds(Pollingtime);
                wait.IgnoreExceptionTypes(typeof(NoSuchElementException), typeof(ElementNotVisibleException));
                wait.Until(WaitforElements(locatorValue));
                ExtentTestManager._parentTest.Log(Status.Pass, "Element visiable to UI");
            }
            catch (Exception ex)
            {
                GeneralMethod.ScreenShotCapture();
                ExtentTestManager._parentTest.Log(Status.Fail, ex.ToString() + " || " + ex.Message);
                Assert.Fail(ex.Message);
            }
        }

        public void WaitforElementbool(int Seconds_time, int Pollingtime, string locatorValue)
        {
            try
            {
                PageLoadWait(20);
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(Seconds_time));
                wait.PollingInterval = TimeSpan.FromMilliseconds(Pollingtime);
                wait.IgnoreExceptionTypes(typeof(NoSuchElementException), typeof(ElementNotVisibleException));
                wait.Until(WaitforElement_icon(locatorValue));
                ExtentTestManager._parentTest.Log(Status.Pass, "Element visiable to UI");
            }
            catch (Exception ex)
            {
                GeneralMethod.ScreenShotCapture();
                ExtentTestManager._parentTest.Log(Status.Fail, ex.ToString() + " || " + ex.Message);
                Assert.Fail(ex.Message);
            }
        }

        private Func<IWebDriver, IWebElement> WaitforElements(string locatorValue)
        {
            return ((x) =>
            {

                ExtentTestManager._parentTest.Log(Status.Pass, "Waiting for element");
                if (x.FindElements(By.XPath(locatorValue)).Count == 1)
                    return x.FindElement(By.XPath(locatorValue));
                return null;
            });
        }

        private Func<IWebDriver, bool> WaitforElement_icon(string locatorValue)
        {

            return ((x) =>
            {
                ExtentTestManager._parentTest.Log(Status.Pass, "Waiting for Icon to be visible");
                return (x.FindElements(By.XPath(locatorValue)).Count == 1);
            });

            
        }

        public void WaitforElement_ExpectedConditions(int Seconds_time, int Pollingtime, string locatorValue)
        {
            try
            {
                PageLoadWait(20);
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(Seconds_time));
                wait.PollingInterval = TimeSpan.FromMilliseconds(Pollingtime);
                wait.IgnoreExceptionTypes(typeof(NoSuchElementException), typeof(ElementNotVisibleException));
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(locatorValue)));
                ExtentTestManager._parentTest.Log(Status.Pass, "Element visiable to UI");
            }
            catch (Exception ex)
            {
                GeneralMethod.ScreenShotCapture();
                ExtentTestManager._parentTest.Log(Status.Fail, ex.ToString() + " || " + ex.Message);
                Assert.Fail(ex.Message);
            }
        }

        public void WaitforElement_Visibility(int Seconds_time, int Pollingtime, string locatorValue)
        {
            try
            {
                PageLoadWait(20);
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(Seconds_time));
                wait.PollingInterval = TimeSpan.FromMilliseconds(Pollingtime);
                wait.IgnoreExceptionTypes(typeof(NoSuchElementException), typeof(ElementNotVisibleException));
                wait.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.XPath(locatorValue)));
                ExtentTestManager._parentTest.Log(Status.Pass, "Element visiable to UI");
            }
            catch (Exception ex)
            {
                GeneralMethod.ScreenShotCapture();
                ExtentTestManager._parentTest.Log(Status.Fail, ex.ToString() + " || " + ex.Message);
                Assert.Fail(ex.Message);
            }
        }

        public void WaitforElement_Exists(int Seconds_time, int Pollingtime, string locatorValue)
        {
            try
            {
                PageLoadWait(20);
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(Seconds_time));
                wait.PollingInterval = TimeSpan.FromMilliseconds(Pollingtime);
                wait.IgnoreExceptionTypes(typeof(NoSuchElementException), typeof(ElementNotVisibleException));
                wait.Until(ExpectedConditions.ElementExists(By.XPath(locatorValue)));
                ExtentTestManager._parentTest.Log(Status.Pass, "Element visiable to UI");
            }
            catch (Exception ex)
            {
                GeneralMethod.ScreenShotCapture();
                ExtentTestManager._parentTest.Log(Status.Fail, ex.ToString() + " || " + ex.Message);
                Assert.Fail(ex.Message);
            }
        }

        public IWebElement getElement(string locatorKey, string locatorValue)
        {
            IWebElement e = null;
            try
            {

                if (locatorKey == Enum.LocatorType.id.ToString())
                    e = driver.FindElement(By.Id(locatorValue));
                else if (locatorKey == Enum.LocatorType.xpath.ToString())
                    e = driver.FindElement(By.XPath(locatorValue));
                else if (locatorKey == Enum.LocatorType.name.ToString())
                    e = driver.FindElement(By.Name(locatorValue));
                else if (locatorKey == Enum.LocatorType.linktext.ToString())
                    e = driver.FindElement(By.LinkText(locatorValue));
                else if (locatorKey == Enum.LocatorType.cssselector.ToString())
                    e = driver.FindElement(By.CssSelector(locatorValue));
            }
            catch (Exception ex)
            {
                //   reportFailure("Failure in element extraction ");
                log.Info(ex.Message);
                Assert.Fail("Failure in element extraction " + locatorValue);
            }
            return e;
        }
        public IList<IWebElement> getElements(string locatorKey, string locatorValue)
        {
            IList<IWebElement> e = null;
            try
            {

                if (locatorKey == Enum.LocatorType.id.ToString())
                    e = driver.FindElements(By.Id(locatorValue));
                else if (locatorKey == Enum.LocatorType.xpath.ToString())
                    e = driver.FindElements(By.XPath(locatorValue));
                else if (locatorKey == Enum.LocatorType.name.ToString())
                    e = driver.FindElements(By.Name(locatorValue));

            }
            catch (Exception ex)
            {
                //   reportFailure("Failure in element extraction ");
                log.Info(ex.Message);
                Assert.Fail("Failure in element extraction " + locatorValue);
            }
            return e;
        }
        public void maximiseBrowser()
        {
            driver.Manage().Window.Maximize();
        }
        public void explicitWait(By aStringElement, int aWaitTimeInSec)
        {
            try

            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(aWaitTimeInSec));
                wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(aStringElement));
                log.Info("Automation waited for Element to be visible");
                ExtentTestManager._parentTest.Log(Status.Pass, "Automation waited for Element to be visible");

            }
            catch (TimeoutException e)
            {
                Console.WriteLine(e.ToString());
                log.Info(e.Message);
                ExtentTestManager._parentTest.Log(Status.Pass, e.ToString() + " || " + e.Message);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                log.Info(e.Message);
                ExtentTestManager._parentTest.Log(Status.Pass, e.ToString());
            }
        }
        public Enum.LogStatus threadWait(int aWaitTimeInSec)
        {
            Thread.Sleep(aWaitTimeInSec);
            return Enum.LogStatus.Passed;
        }
        public bool WaitUntillElementIsVisible(String aStringType, String aStringeValue, int timeOut)
        {
            int iTimer = 0;
            while (iTimer <= timeOut)
            {
                bool status = IsElementVisible(aStringType, aStringeValue);
                if (status)
                    return true;
            }
            return false;
        }
        public void browserBackNavigation()
        {
            driver.Navigate().Back();
        }
        public void browserForwardNavigation()
        {
            driver.Navigate().Forward();
        }
        public string getAttributeOfWebelementByLocator(By aStringValue, String aAttribute)
        {
            explicitWait(aStringValue, 300);
            IWebElement element = driver.FindElement(aStringValue);
            return element.GetAttribute(aAttribute);
        }
        public IWebElement getWebElementByLocator(String aStringType, String aStringValue)
        {

            IWebElement webElement = getElement(aStringType, aStringValue);
            return webElement;
        }
        public string getTextOfWebElementByLocator(String aStringType, String aWebElementID)
        {
            return getWebElementByLocator(aStringType, aWebElementID).Text;
        }
        public string getWebElementByLocatortooltip(String aStringType, String aStringValue) // Tooltip 
        {
            IWebElement webElement = getElement(aStringType, aStringValue);
            string Tooltiptext = webElement.GetAttribute("title");
            return Tooltiptext;
        }

        public string getTextOfWebElementByLocatortooltip(String aStringType, String aWebElementID) // Tooltip 
        {

            return getWebElementByLocatortooltip(aStringType, aWebElementID);
        }
        public Enum.LogStatus AssertAreEqualtooltip(String aStringType, String expected, string actual)// Tooltip 
        {
            ExtentTestManager._parentTest.Log(Status.Info, "Started Comparing Expected vs Actual ");
            try
            {
                Assert.AreEqual(getTextOfWebElementByLocatortooltip(aStringType, expected), actual);
                ExtentTestManager._parentTest.Log(Status.Pass, "Expected matched with Actual");
                return status = Enum.LogStatus.Passed;

            }
            catch (Exception)
            {
                ExtentTestManager._parentTest.Log(Status.Fail, "Expected not matched with actual " + actual);
                return status = Enum.LogStatus.Failed;
            }

        }

        public Enum.LogStatus ClickOnElementWhenElementFound(String aStringType, String aStringValue, String aStringName)
        {
            //ExtentTestManager._parentTest.Log(Status.Info, "Clicking on -" + aStringName);
            try
            {

                IWebElement webElement = getWebElementByLocator(aStringType, aStringValue);
                webElement.Click();
                ExtentTestManager._parentTest.Log(Status.Pass, "Automation Clicked Web Element : " + aStringName);
                return status = Enum.LogStatus.Passed;
            }
            catch (Exception e)
            {
                ExtentTestManager._parentTest.Log(Status.Fail, "Automation unable to Click Web Element : " + aStringName);
                return status = Enum.LogStatus.Failed;
            }
        }

        public Enum.LogStatus ClickOnElementWhenElementFound_Search(String aStringType, String aStringValue, String aStringName)
        {
            ExtentTestManager._parentTest.Log(Status.Info, "Clicking on -" + aStringName);
            try
            {

                IWebElement webElement = getWebElementByLocator(aStringType, aStringValue);
                webElement.Click();
                ExtentTestManager._parentTest.Log(Status.Pass, "Automation Clicked Web Element : " + aStringName);
                return status = Enum.LogStatus.Passed;
            }
            catch (Exception e)
            {
                ExtentTestManager._parentTest.Log(Status.Debug, "Automation  to Click Web Element : " + aStringName);
                return status = Enum.LogStatus.Failed;
            }
        }


        public Enum.LogStatus PageLoadWait(int iWaitsec)
        {
            try
            {
                driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(iWaitsec);
                ExtentTestManager._parentTest.Log(Status.Pass, "Automation Wait until Page get loaded :- > Max time to Page Load Wait will be " + iWaitsec);
                return Enum.LogStatus.Passed;
            }
            catch (TimeoutException t)
            {
                Console.WriteLine(t.ToString());
                log.Info(t.Message);
                ExtentTestManager._parentTest.Log(Status.Fail, t.ToString() + " || " + t.Message);
                return Enum.LogStatus.Passed;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                log.Info(e.Message);
                ExtentTestManager._parentTest.Log(Status.Fail, e.ToString() + " || " + e.Message);
                return Enum.LogStatus.Passed;
            }
        }
        public Enum.LogStatus ElementWaitTime(int iWaitsec)
        {
            try
            {
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(iWaitsec);
                ExtentTestManager._parentTest.Log(Status.Pass, "Automation Wait until Element loaded :- > Max time to Page Load Wait will be " + iWaitsec);
                return Enum.LogStatus.Passed;
            }
            catch (TimeoutException t)
            {
                Console.WriteLine(t.ToString());
                log.Info(t.Message);
                ExtentTestManager._parentTest.Log(Status.Fail, t.ToString() + " || " + t.Message);
                return Enum.LogStatus.Failed;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                log.Info(e.Message);
                ExtentTestManager._parentTest.Log(Status.Fail, e.ToString() + " || " + e.Message);
                return Enum.LogStatus.Failed;
            }
        }

        public Enum.LogStatus SendKeysForElement(String aStringType, String aStringValue, String aTestData, String aStringName)
        {
            ExtentTestManager._parentTest.Log(Status.Info, "Entering into - " + aStringName);
            try
            {

                getElement(aStringType, aStringValue).SendKeys(aTestData);
                ExtentTestManager._parentTest.Log(Status.Pass, "User Entered the Value  " + aStringName);
                log.Info("Entered  value : " + aStringName);
                return status = Enum.LogStatus.Passed;
            }
            catch (ElementNotVisibleException e)
            {
                Console.WriteLine(e.ToString());
                log.Info(e.Message);
                ExtentTestManager._parentTest.Log(Status.Fail, "User unable to Enter the Value  " + aStringName);
                return status = Enum.LogStatus.Failed;
            }
            catch (Exception e)
            {
                ExtentTestManager._parentTest.Log(Status.Fail, "User unable to Enter the Value  " + aStringName);
                log.Info(e.Message);
                return status = Enum.LogStatus.Failed;
            }
        }

        public Enum.LogStatus SendKeysForElementClick(String aStringType, String aStringValue, String aTestData, String aStringName)
        {
            ExtentTestManager._parentTest.Log(Status.Info, "Entering into - " + aStringName);
            try
            {
               
                getElement(aStringType, aStringValue).SendKeys(Keys.Enter + aTestData);
                ExtentTestManager._parentTest.Log(Status.Pass, "User Entered the Value  " + aStringName);
                log.Info("Entered  value : " + aStringName);
                return status = Enum.LogStatus.Passed;
            }
            catch (ElementNotVisibleException e)
            {
                Console.WriteLine(e.ToString());
                log.Info(e.Message);
                ExtentTestManager._parentTest.Log(Status.Fail, "User unable to Enter the Value  " + aStringName);
                return status = Enum.LogStatus.Failed;
            }
            catch (Exception e)
            {
                ExtentTestManager._parentTest.Log(Status.Fail, "User unable to Enter the Value  " + aStringName);
                log.Info(e.Message);
                return status = Enum.LogStatus.Failed;
            }
        }
        public Enum.LogStatus SendKeysForAElement(String aStringType, String aStringValue, String aTestData, String aStringName)
        {
            ExtentTestManager._parentTest.Log(Status.Info, "Entering into - " + aStringName);
            try
            {
                getElement(aStringType, aStringValue).SendKeys(aTestData + Keys.Enter);
                ExtentTestManager._parentTest.Log(Status.Pass, "User Entering the desire value :" + aStringName);
                ExtentTestManager._parentTest.Log(Status.Pass, "User Click the Button");
                log.Info("Entered  value : " + aStringName);
                return status = Enum.LogStatus.Passed;
            }
            catch (ElementNotVisibleException e)
            {
                Console.WriteLine(e.ToString());
                log.Info(e.Message);
                return status = Enum.LogStatus.Failed;
            }
            catch (Exception e)
            {
                ExtentTestManager._parentTest.Log(Status.Fail, "Unbale to Enter the test data: " + aTestData);
                log.Info(e.Message);
                return status = Enum.LogStatus.Failed;
            }
        }
        public Enum.LogStatus SendKeysForWebElement(String aStringType, String aStringValue, String aTestData, String aStringName)
        {
            ExtentTestManager._parentTest.Log(Status.Info, "Entering into - " + aStringName);
            try
            {
                getElement(aStringType, aStringValue).Clear();
                getElement(aStringType, aStringValue).SendKeys(aTestData);
                ExtentTestManager._parentTest.Log(Status.Pass, "User Entering the desire value :" + aTestData);
                log.Info("Entered  value : " + aTestData);
                return status = Enum.LogStatus.Passed;
            }
            catch (ElementNotVisibleException e)
            {
                Console.WriteLine(e.ToString());
                ExtentTestManager._parentTest.Log(Status.Fail, aTestData + e.Message);
                log.Info(e.Message);
                return status = Enum.LogStatus.Warning;
            }
            catch (Exception e)
            {
                ExtentTestManager._parentTest.Log(Status.Fail, "Unbale to Enter the test data: " + aTestData);
                log.Info(e.Message);
                return status = Enum.LogStatus.Failed;
            }
        }
        public bool IsElementVisible(String aStringType, String aStringValue)
        {
            try
            {
                IWebElement element = getElement(aStringType, aStringValue);
                if (element.Displayed)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (NoSuchElementException)
            {
                return false;
            }
            catch (Exception e)
            {
                return false;
            }
        }

        // start
        public Enum.LogStatus ClearValueOnElementWhenElementFound(String aStringType, String aStringValue, String aStringName)
        {

            try
            {

                IWebElement webElement = getWebElementByLocator(aStringType, aStringValue);
                webElement.Clear();
                ExtentTestManager._parentTest.Log(Status.Pass, "Automation Clicked Web Element : " + aStringName);
                return status = Enum.LogStatus.Passed;
            }
            catch (Exception e)
            {
                ExtentTestManager._parentTest.Log(Status.Fail, "Automation unable to Click Web Element : " + aStringName);
                return status = Enum.LogStatus.Failed;
            }
        }
        public Enum.LogStatus ScrollToElement(String aStringType, String aStringValue, String aStringName)
        {
            Actions actions = new Actions(driver);

            actions.MoveToElement(getElement(aStringType, aStringValue));

            actions.Perform();
            ExtentTestManager._parentTest.Log(Status.Pass, "User Scrolled to the WebElement");
            return Enum.LogStatus.Passed;
        }


        /// <summary>
        /// Desc: Method is used to verify Date List and Time in Descending Order or not
        /// </summary>
        /// <returns>Status</returns>
        /// <Parameter></Parameter>
        public bool VerifyListSortingDescendingOrder(List<DateTime> dateTimeList)
        {
            var previousDateTimeItem = dateTimeList.FirstOrDefault();

            foreach (DateTime currentDateTimeItem in dateTimeList)
            {
                if (currentDateTimeItem.CompareTo(previousDateTimeItem) > 0)
                    return false;
            }

            return true;
        }

        /// <summary>
        /// Desc: Method is used to verify Date List and Time in Ascending Order or not
        /// </summary>
        /// <returns>Status</returns>
        /// <Parameter></Parameter>
        public bool VerifyListSortingAscendingOrder(List<DateTime> dateTimeList)
        {
            var previousDateTimeItem = dateTimeList.FirstOrDefault();

            foreach (DateTime currentDateTimeItem in dateTimeList)
            {
                if (currentDateTimeItem.CompareTo(previousDateTimeItem) < 0)
                    return false;
            }

            return true;
        }

        //end 
        public bool waitUntilElementNotVisible(String aStringType, String aStringeValue, int timeOut)
        {

            int iTimer = 0;
            while (iTimer <= timeOut)
            {
                bool status = !(IsElementVisible(aStringType, aStringeValue));
                if (status)
                    return true;
            }
            return false;
        }
        public bool IsElementPresent(String aStringType, String aStringValue)
        {
            try
            {
                getElement(aStringType, aStringValue);
                return true;
            }
            catch (Exception e)
            {
                return false;
            }

        }

        public String selectValueFromDropdown(String aStringValue, string value, String aStringType)
        {
            try
            {
                SelectElement oSelect = new SelectElement(getElement(aStringType, aStringValue));
                oSelect.SelectByText(value);
                return value;
            }
            catch (NoSuchElementException)
            {
                return null;
            }
        }
        public Enum.LogStatus mouseHover(String aStringType, String aStringValue)
        {
            Actions act = new Actions(driver);
            act.MoveToElement(getElement(aStringType, aStringValue)).Click();
            return Enum.LogStatus.Passed;
        }
        public string selectValueFromDropdownByText(String aStringType, String aStringValue, string value)
        {
            try
            {
                SelectElement oSelect = new SelectElement(getElement(aStringType, aStringValue));
                oSelect.SelectByText(value);
                return value;
            }
            catch (NoSuchElementException)
            {
                return null;
            }
        }
        public int selectValueByIndex(String aStringType, String aStringValue, int index)
        {
            try
            {
                SelectElement oSelect = new SelectElement(getElement(aStringType, aStringValue));
                oSelect.SelectByIndex(index);
                return index;
            }
            catch (NoSuchElementException)
            {
                return -1;
            }
        }
        public Enum.LogStatus AssertAreEqual(String aStringType, String expected, string actual)
        {
            //ExtentTestManager._parentTest.Log(Status.Info, "Entering into - " + actual);
            try
            {
                Assert.AreEqual(getTextOfWebElementByLocator(aStringType, expected), actual);
                ExtentTestManager._parentTest.Log(Status.Pass, "Expected matched with actual - " + actual);
                return status = Enum.LogStatus.Passed;

            }
            catch (Exception)
            {
                ExtentTestManager._parentTest.Log(Status.Fail, "Expected not matched with actual " + actual);
                GeneralMethod.ScreenShotCapture();
                return status = Enum.LogStatus.Failed;
            }

        }

        
        public Enum.LogStatus AssertIsTrue(String type, String eWebElement, String aStringName)
        {
            try
            {
                Assert.IsTrue(IsElementVisible(type, eWebElement));
                ExtentTestManager._parentTest.Log(Status.Pass, "Web Element is visible hence Expected condition is met : " + aStringName);
                return status = Enum.LogStatus.Passed;
            }
            catch (Exception)
            {
                ExtentTestManager._parentTest.Log(Status.Fail, "Web Element is not visible hence Expected condition is not met. " + aStringName);
                GeneralMethod.ScreenShotCapture();
                return status = Enum.LogStatus.Failed;
            }
        }
        public Enum.LogStatus AssertIsFalse(String type, String eWebElement)
        {
            try
            {
                Assert.IsFalse(IsElementVisible(type, eWebElement));
                ExtentTestManager._parentTest.Log(Status.Pass, "Web Element is not visible hence Expected condition is met.");
                return status = Enum.LogStatus.Passed;
            }
            catch (Exception)
            {
                ExtentTestManager._parentTest.Log(Status.Fail, "Web Element is visible hence Expected condition is not met.");
                return status = Enum.LogStatus.Failed;
            }
        }

        public static Enum.LogStatus AssertIgnore()
        {
            try
            {
                Assert.Ignore("Skipping the test");
                ExtentTestManager._parentTest.Log(Status.Pass, "Skipping the Test.");
                return Enum.LogStatus.Skipped;
            }
            catch (Exception)
            {
                ExtentTestManager._parentTest.Log(Status.Fail, "Skipping the Test.");
                return Enum.LogStatus.Skipped;
            }
        }
        public Enum.LogStatus jsClick(String aStringType, String aStringValue, String aStringName)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("arguments[0].click();", aStringValue);
            ExtentTestManager._parentTest.Log(Status.Pass, "User Clicked the WebElement");
            return Enum.LogStatus.Passed;
        }

        /// <summary>
        /// Desc:Method is used to capture the ScreenShots
        /// </summary>
        /// <returns></returns>
        public static string ScreenShotCapture()
        {
            try
            {
                // driver.Manage().Window.FullScreen();
                string CurrentDates = string.Format("{0:dd-MM-yyy h-mm-ss}", DateTime.Now);
                string filename = CurrentDates + ".jpeg";
                string finalpath = GetScreenshotPath() + filename;
                ITakesScreenshot screenshotDriver = driver as ITakesScreenshot;
                Screenshot screenshot = screenshotDriver.GetScreenshot();
                screenshot.SaveAsFile(finalpath, ScreenshotImageFormat.Jpeg);
                finalpath = "Screenshots//" + filename;
                return finalpath;
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Desc:Method is used to GetDrivers path
        /// </summary>
        /// <returns></returns>
        public string GetDriversPath()
        {
            string path = System.Reflection.Assembly.GetCallingAssembly().CodeBase;
            string actualPath = path.Substring(0, path.LastIndexOf("bin"));
            string driverPath = new Uri(actualPath).LocalPath;
            driverPath = driverPath + "Driver";
            return driverPath;
        }

        /// <summary>
        /// Desc:Method is used to get generated report's path
        /// </summary>
        /// <returns></returns>
        public static string GetReportPath()
        {
            string path = System.Reflection.Assembly.GetCallingAssembly().CodeBase;
            string actualPath = path.Substring(0, path.LastIndexOf("bin"));
            string reportPath = new Uri(actualPath).LocalPath;
            string CurrentDates = string.Format("{0:dd-MM-yyy h-mm-ss}", DateTime.Now);
            reportPath = reportPath + @"ResultReport\\ExtentReport" + " " + CurrentDates + ".html";
            return reportPath;
        }

        /// <summary>
        /// Desc:Method is used to get json path
        /// </summary>
        /// <returns></returns>
        public static string GetJsonpath()
        {
            string path = System.Reflection.Assembly.GetCallingAssembly().CodeBase;
            string actualPath = path.Substring(0, path.LastIndexOf("bin"));
            string jsonpath = new Uri(actualPath).LocalPath;
            jsonpath = jsonpath + "Util.Action.Package\\Locator.json";
            return jsonpath;
        }

        /// <summary>
        /// Desc:Method is used to Get Screenshot's Path
        /// </summary>
        /// <returns></returns>
        public static string GetScreenshotPath()
        {
            string path = System.Reflection.Assembly.GetCallingAssembly().CodeBase;
            string actualPath = path.Substring(0, path.LastIndexOf("bin"));
            string screenshotPath = new Uri(actualPath).LocalPath;
            string CurrentDates = string.Format("{0:dd-MM-yyy h-mm-ss}", DateTime.Now);
            screenshotPath = screenshotPath + @"ResultReport\Screenshots\";
            return screenshotPath;
        }

        /// <summary>
        /// Desc:Method is used to Get Report Folder's Path
        /// </summary>
        /// <returns></returns>
        public static string GetReportFolderPath()
        {
            string path = System.Reflection.Assembly.GetCallingAssembly().CodeBase;
            string actualPath = path.Substring(0, path.LastIndexOf("bin"));
            string reportPath = new Uri(actualPath).LocalPath;
            reportPath = reportPath + "ResultReport";
            return reportPath;
        }

        /// <summary>
        /// Desc:Method is used to Get zip Folder's Path
        /// </summary>
        /// <returns></returns>
        public static string GetZipFolderPath()
        {
            string path = System.Reflection.Assembly.GetCallingAssembly().CodeBase;
            string actualPath = path.Substring(0, path.LastIndexOf("bin"));
            string reportPath = new Uri(actualPath).LocalPath;
            reportPath = reportPath + "ZipFolder";
            return reportPath;
        }
        /// <summary>
        /// Desc:Method is used to set report into zip file
        /// </summary>
        /// <returns></returns>
        public static void createZipFile()
        {
            try
            {
                string reportPath = GetReportFolderPath();
                string zipFilePath = GetZipFolderPath();
                bool exists = System.IO.Directory.Exists(reportPath);
                if (!exists)
                    System.IO.Directory.CreateDirectory(reportPath);
                addIntoZip(reportPath, zipFilePath);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
        public static void addIntoZip(string directoryPath, string zipFilePath)
        {
            DirectoryInfo dirZipPath = new DirectoryInfo(zipFilePath);
            foreach (FileInfo fi in dirZipPath.GetFiles())
            {
                fi.Delete();
            }
            string CurrentDates = string.Format("{0:dd-MM-yyy h-mm-ss}", DateTime.Now);
            ZipFile.CreateFromDirectory(directoryPath, Path.Combine(zipFilePath, "ExtentReport" + "  " + CurrentDates + ".zip"), CompressionLevel.Optimal, true);
        }
        /// <summary>
        /// Desc:Method is used to get excelsheet's path
        /// </summary>
        /// <returns></returns>
        public static string GetExcelPath()
        {
            string path = System.Reflection.Assembly.GetCallingAssembly().CodeBase;
            string actualPath = path.Substring(0, path.LastIndexOf("bin"));
            string excelPath = new Uri(actualPath).LocalPath;
            excelPath = excelPath + "ExcelSheet\\AllTestData.xlsx";
            return excelPath;
        }

        /// <summary>
        /// Desc:Method is used to Open New tab
        /// </summary>
        /// <returns></returns>
        public void MethodtToOpenNewtab()
        {
            ((IJavaScriptExecutor)driver).ExecuteScript("window.open();");
            driver.SwitchTo().Window(driver.WindowHandles.Last());
            driver.Navigate().GoToUrl("");
        }
       
        /// <summary>
        /// Desc:Method is used to mouse hover without the click
        /// </summary>
        /// <param name="aStringType"></param>
        /// <param name="aStringValue"></param>
        /// <param name="aStringName"></param>
        /// <returns></returns>
        public Enum.LogStatus mouseHoverWithoutClick(String aStringType, String aStringValue, String aStringName)
        {
            //test.Log(Status.Info, "hover on - " + aStringName);
            try
            {
                Actions action = new Actions(driver);
                action.MoveToElement(getElement(aStringType, aStringValue)).Build().Perform();
                return Enum.LogStatus.Passed;
            }
            catch (Exception e)
            {
                return Enum.LogStatus.Failed;
            }
        }


        /// <summary>
        /// Desc: Method is used 
        /// </summary>
        /// <param name="aStringType_LocatorName"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public string Locatore(string aStringType_LocatorName, string value)
        {
            try
            {
                value = ConfigurationManager.AppSettings[aStringType_LocatorName];
                return value;
            }

            catch
            {
                ExtentTestManager._parentTest.Log(Status.Fail, "Expected not matched with actual ");
                return value;
            }
        }


        /// <summary>
        /// Desc: Method is used read Excel Data from Json
        /// </summary>
        /// <param name="aStringType_LocatorName"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public void JsonFunctionLibray_Reader()
        {
            try
            {
                using (ExcelEngine excelEngine = new ExcelEngine())
                {
                    IApplication application = excelEngine.Excel;

                    //The workbook is opened.
                    FileStream fileStream = new FileStream(GetExcelPath(), FileMode.Open);

                    IWorkbook workbook = application.Workbooks.Open(fileStream, ExcelOpenType.Automatic);
                    IWorksheet worksheet = workbook.Worksheets["Locators"];

                    //Export worksheet data into CLR Objects
                    IList<JsonReadLibrary> jsonfiles = worksheet.ExportData<JsonReadLibrary>(1, 1, worksheet.UsedRange.LastRow, workbook.Worksheets[0].UsedRange.LastColumn);

                    //open file stream
                    using (StreamWriter file = File.CreateText(GetJsonpath()))
                    {
                        JsonSerializer serializer = new JsonSerializer();

                        //serialize object directly into file stream
                        serializer.Serialize(file, jsonfiles);
                    }
                }
            }
            catch (Exception ex)
            {
                GeneralMethod.ScreenShotCapture();
                Assert.Fail(ex.Message);
            }
        }

        /// <summary>
        /// Desc: Method is used read Json to code for Locator Types
        /// </summary>
        /// <param name="LocatorTypes_value"></param>
        /// <param name="LocatorNames"></param>
        /// <returns></returns>
        public string JsonTocode_LT(string LocatorTypes_value, string LocatorNames)
        {
            try
            {
                var client = new WebClient();
                var text = client.DownloadString(GetJsonpath());
                var serializers = new System.Web.Script.Serialization.JavaScriptSerializer();
                List<Jsontocode> objlist = (List<Jsontocode>)serializers.Deserialize(text, typeof(List<Jsontocode>));
                foreach (Jsontocode obj in objlist)
                {
                    if (LocatorNames.Equals(obj.LocatorName))
                    {
                        LocatorTypes_value = obj.LocatorType;
                    }
                    else
                    { }
                    //Console.WriteLine("Locator Type :" + obj.LocatorType + "Locator value : " + obj.LocatorValue + "Locator name :" + obj.LocatorName);
                }
                return LocatorTypes_value;
            }
            catch (Exception ex)
            {
                GeneralMethod.ScreenShotCapture();
                Assert.Fail(ex.Message);
                return LocatorTypes_value;
            }
        }

        /// <summary>
        /// Desc: Method is used read Json to code for Locator Value
        /// </summary>
        /// <param name="LocatorValues_value"></param>
        /// <param name="LocatorNames"></param>
        /// <returns></returns>
        public string JsonTocode_LV(string LocatorValues_value, string LocatorNames)
        {
            try
            {
                var client = new WebClient();
                var text = client.DownloadString(GetJsonpath());
                var serializers = new System.Web.Script.Serialization.JavaScriptSerializer();
                List<Jsontocode> objlist = (List<Jsontocode>)serializers.Deserialize(text, typeof(List<Jsontocode>));
                foreach (Jsontocode obj in objlist)
                {
                    if (LocatorNames.Equals(obj.LocatorName))
                    {
                        LocatorValues_value = obj.LocatorValue;
                    }
                    else { }

                }
                return LocatorValues_value;
            }
            catch (Exception ex)
            {
                GeneralMethod.ScreenShotCapture();
                NUnit.Framework.Assert.Fail(ex.Message);
                return LocatorValues_value;
            }
        }


        /// <summary>
        /// Desc:Method is used to Reading Existing excel sheets and execute the keywords accordingly
        /// Update Excel TestData :- For New Cases
        /// </summary>
        public static void ReadExistingExcel()
        {

            string path = System.Reflection.Assembly.GetCallingAssembly().CodeBase;
            string actualPath = path.Substring(0, path.LastIndexOf("bin"));
            string excelPath = new Uri(actualPath).LocalPath;
            excelPath = excelPath + "ExcelSheet\\AllTestData.xlsx";
            log.Info("All Test Data Excel file Initalized  !!!");
            Random random = new Random();
            int num = random.Next(1000);
            oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = true;
            oXL.DisplayAlerts = false;
            mWorkBook = oXL.Workbooks.Open(excelPath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            mWorkSheets = mWorkBook.Worksheets;
            mWSheet1 = (Microsoft.Office.Interop.Excel.Worksheet)mWorkSheets.get_Item("TestData");
            Microsoft.Office.Interop.Excel.Range range = mWSheet1.UsedRange;
            int colCount = range.Columns.Count;
            int rowCount = range.Rows.Count;
            for (int index = 1; index <= 1; index++)
            {
                mWSheet1.Cells[3, 13] = "ClientAutomation" + num;  // TextBox_ContractShortName
                ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet in TextBox_ContractShortName Column : " + " || " + "TestingNewEndflow001" + num);
                mWSheet1.Cells[3, 14] = "LegalNameAutomation" + num; //TextBox_ContractLegalName
                ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet in TextBox_ContractLegalName Column : " + " || " + num);
                mWSheet1.Cells[3, 17] = "Automation" + num; //TextBox_TotalContractValue
                ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet in TextBox_TotalContractValue Column : " + " || " + num);
                mWSheet1.Cells[3, 24] = "AutomationL1" + num; //Textbox_DocumentNameL1
                ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet in Textbox_DocumentNameL1 Column : " + " || " + num);
                mWSheet1.Cells[3, 25] = "AutomationL1Modify" + num; //Textbox_ModifyDocumentNameL1
                ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet in Textbox_ModifyDocumentNameL1 Column : " + " || " + num);
                mWSheet1.Cells[3, 26] = "Automation Enter Additional Information for Modify L1 Document" + num; //Textarea_AdditionalInformationL1
                ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet in Textarea_AdditionalInformationL1 Column : " + " || " + num);
                mWSheet1.Cells[3, 27] = "AutomationObligation" + num; //TextArea_ObligationExtract
                ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet in TextArea_ObligationExtract Column : " + " || " + num);
                mWSheet1.Cells[3, 28] = "Automation add new obligation for newly created Contract" + num; //Textbox_ObligationTopic&Description
                ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet in Textbox_ObligationTopic&Description Column : " + " || " + num);
                mWSheet1.Cells[3, 36] = "AutomationReference" + num; //Textbox_ReferenceSection
                ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet in Textbox_ReferenceSection Column : " + " || " + num);
                mWSheet1.Cells[3, 43] = "AutomationObligationModify" + num; //TextArea_ModifyObligationExtract
                ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet in TextArea_ModifyObligationExtract Column : " + " || " + num);
                mWSheet1.Cells[3, 44] = "Adding Modify Obligation Extract" + num; //TextArea_AdditionalInformationModifyObligation
                ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet in TextArea_AdditionalInformationModifyObligation Column : " + " || " + num);
                mWSheet1.Cells[3, 58] = "Description for High Risk " + num; //Description_InputText
                ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet in Description_InputText Column : " + " || " + num);

            }

            mWorkBook.Save();
            ExtentTestManager._parentTest.Log(Status.Pass, "Automation Save the Excel Sheet.");
            mWorkBook.Close(Missing.Value, Missing.Value, Missing.Value);
            ExtentTestManager._parentTest.Log(Status.Pass, "Automation Close the Excel Sheet.");
            mWSheet1 = null;
            mWorkBook = null;
            oXL.Quit();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        /// <summary>
        /// Desc:Method is used to Reading Existing excel sheets and execute the keywords accordingly
        /// Update Excel TestData :- For old Test case
        /// </summary>
        public string ReadOldExistingExcel(string Openacasename, string TestData_BillingID, string LawSessionName)
        {
            try
            {
                string path = System.Reflection.Assembly.GetCallingAssembly().CodeBase;
                string actualPath = path.Substring(0, path.LastIndexOf("bin"));
                string excelPath = new Uri(actualPath).LocalPath;
                excelPath = excelPath + "ExcelSheet\\AllTestData.xlsx";
                log.Info("All Test Data Excel file Initalized  !!!");
                Random random = new Random();
                int num = random.Next(1000);
                int num1 = num + 1;
                oXL = new Microsoft.Office.Interop.Excel.Application();
                oXL.Visible = true;
                oXL.DisplayAlerts = false;
                mWorkBook = oXL.Workbooks.Open(excelPath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                mWorkSheets = mWorkBook.Worksheets;
                mWSheet1 = (Microsoft.Office.Interop.Excel.Worksheet)mWorkSheets.get_Item("TestData");
                Microsoft.Office.Interop.Excel.Range range = mWSheet1.UsedRange;
                int colCount = range.Columns.Count;
                int rowCount = range.Rows.Count;
                for (int index = 1; index <= 1; index++)
                {
                    mWSheet1.Cells[3, 19] = Openacasename;  // Openacasename Column
                    ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet in Openacasename Column : " + " || " + Openacasename);
                    mWSheet1.Cells[3, 21] = "CustodianKev" + num; // CustodianFirstName Column
                    ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet in CustodianFirstName Column : " + " || " + "CustodianKev" + num);



                }
                mWorkBook.Save();
                ExtentTestManager._parentTest.Log(Status.Pass, "Automation Save the Excel Sheet.");
                Thread.Sleep(5000);
                mWorkBook.Close(Missing.Value, Missing.Value, Missing.Value);
                ExtentTestManager._parentTest.Log(Status.Pass, "Automation Close the Excel Sheet.");
                mWSheet1 = null;
                mWorkBook = null;
                oXL.Quit();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                return TestData_BillingID;
            }
            catch (Exception ex)
            {
                GeneralMethod.ScreenShotCapture();
                return TestData_BillingID;
                Assert.Fail(ex.Message);
            }

        }


        // File Updload action
        public Enum.LogStatus Fileupload(string TestData)
        {
            try
            {
                string path = System.Reflection.Assembly.GetCallingAssembly().CodeBase;
                string actualPath = path.Substring(0, path.LastIndexOf("bin"));
                string FileUpload = new Uri(actualPath).LocalPath;
                FileUpload = FileUpload + "FileUpload\\FileUploadScrpit.exe";
                ExtentTestManager._parentTest.Log(Status.Pass, "Automation started reading File upload Exe");
                string ArgumentFileUpload = new Uri(actualPath).LocalPath;
                ArgumentFileUpload = "\""+ ArgumentFileUpload + "FileUpload\\" + TestData + "\""; 
                ExtentTestManager._parentTest.Log(Status.Pass, "Automation started reading test data which need to upload");
                ProcessStartInfo processinfo = new ProcessStartInfo();
                processinfo.FileName = FileUpload;
                ExtentTestManager._parentTest.Log(Status.Pass, "Automation excecuted file upload exe.");
                processinfo.Arguments = ArgumentFileUpload;
                //processinfo.Arguments = @"C:\AddL1.xlsx";
                Process process = Process.Start(processinfo);
                ExtentTestManager._parentTest.Log(Status.Pass, "Automation uploaded the File");
                process.WaitForExit();
                process.Close();
                ExtentTestManager._parentTest.Log(Status.Pass, "File Upload prcoess done Automation closed the process");
                return status = Enum.LogStatus.Passed;
            }
            catch (Exception ex)
            {

                ExtentTestManager._parentTest.Log(Status.Fail, "File Upload prcoess fail");
                GeneralMethod.ScreenShotCapture();
                return status = Enum.LogStatus.Failed;
            }
        }


        /// <summary>
        /// Desc:Method is used to update write Dcoument
        /// </summary>
        public static void UpdateUploadExcelfile(Dictionary<string, string> testData, string DocumentUpload, string TestData)
        {
            try
            {
                string ValueName1 = testData["UploadHader_AddL1Document"];
                string ValueName2 = testData["UploadHader_AddL2Document"];
                string NameofExcel1 = testData["Upload_AddL1Document"];
                string NameofExcel2 = testData["Upload_AddL2Document"];

                if (DocumentUpload.Equals(ValueName1) & NameofExcel1.Equals(TestData))
                {
                    string path = System.Reflection.Assembly.GetCallingAssembly().CodeBase;
                    string actualPath = path.Substring(0, path.LastIndexOf("bin"));
                    string excelPath = new Uri(actualPath).LocalPath;
                    excelPath = excelPath + "FileUpload\\" + TestData;
                    log.Info("All Test Data Excel file Initalized  !!!");
                    Random random = new Random();
                    int num = random.Next(1000);
                    oXL = new Microsoft.Office.Interop.Excel.Application();
                    oXL.Visible = true;
                    oXL.DisplayAlerts = false;
                    mWorkBook = oXL.Workbooks.Open(excelPath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    mWorkSheets = mWorkBook.Worksheets;
                    mWSheet1 = (Microsoft.Office.Interop.Excel.Worksheet)mWorkSheets.get_Item("Sheet1");
                    Microsoft.Office.Interop.Excel.Range range = mWSheet1.UsedRange;
                    int colCount = range.Columns.Count;
                    int rowCount = range.Rows.Count;
                    for (int index = 1; index <= 1; index++)
                    {
                        mWSheet1.Cells[2, 2] = "Automation_AddL1DocumentWithExcel" + num;  // Document Name
                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet in Document Namee  : " + " || " + "Automation_AddL1DocumentWithExcel" + num);
                        TodaysDate = DateTime.Now;
                        string TodaysDate_effectiveDate = Convert.ToString(TodaysDate);
                        string TodaysDate_effectiveDateMain = Convert.ToDateTime(TodaysDate_effectiveDate).ToString("MM/dd/yyyy");
                        mWSheet1.Cells[3, 2] = TodaysDate_effectiveDateMain; //Effective Date
                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet in Effective Date: " + TodaysDate_effectiveDateMain);
                        FivedaysDate = DateTime.Now;
                        FivedaysDate = FivedaysDate.AddDays(5);
                        string FivedaysDate_expDate = Convert.ToString(FivedaysDate);
                        string FivedaysDate_expDateMain = Convert.ToDateTime(FivedaysDate_expDate).ToString("MM/dd/yyyy");
                        mWSheet1.Cells[4, 2] = FivedaysDate_expDateMain;  //Expiration Date
                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet in Expiration Date : " + " || " + FivedaysDate_expDateMain);
                        mWSheet1.Cells[5, 2] = "Automation_AgreementNumber" + num; //Agreement Number
                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet in Agreement Number : " + " || " + "Automation_AgreementNumber" + num);
                    }
                    mWorkBook.Save();
                    ExtentTestManager._parentTest.Log(Status.Pass, "Automation Save the Excel Sheet.");
                    mWorkBook.Close(Missing.Value, Missing.Value, Missing.Value);
                    ExtentTestManager._parentTest.Log(Status.Pass, "Automation Close the Excel Sheet.");
                }
                else if (DocumentUpload.Equals(ValueName2) & NameofExcel2.Equals(TestData))
                {
                    string path = System.Reflection.Assembly.GetCallingAssembly().CodeBase;
                    string actualPath = path.Substring(0, path.LastIndexOf("bin"));
                    string excelPath = new Uri(actualPath).LocalPath;
                    excelPath = excelPath + "FileUpload\\" + TestData;
                    log.Info("All Test Data Excel file Initalized  !!!");
                    Random random = new Random();
                    int num = random.Next(1000);
                    oXL = new Microsoft.Office.Interop.Excel.Application();
                    oXL.Visible = true;
                    oXL.DisplayAlerts = false;
                    mWorkBook = oXL.Workbooks.Open(excelPath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    mWorkSheets = mWorkBook.Worksheets;
                    mWSheet1 = (Microsoft.Office.Interop.Excel.Worksheet)mWorkSheets.get_Item("Sheet1");
                    Microsoft.Office.Interop.Excel.Range range = mWSheet1.UsedRange;
                    int colCount = range.Columns.Count;
                    int rowCount = range.Rows.Count;
                    for (int index = 1; index <= 1; index++)
                    {
                        mWSheet1.Cells[2, 2] = "Automation_AddL2DocumentWithExcel" + num;  // Document Name
                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet in Document Namee  : " + " || " + "Automation_AddL2DocumentWithExcel" + num);
                        TodaysDate = DateTime.Now;
                        string TodaysDate_effectiveDate = Convert.ToString(TodaysDate);
                        string TodaysDate_effectiveDateMain = Convert.ToDateTime(TodaysDate_effectiveDate).ToString("MM/dd/yyyy");
                        mWSheet1.Cells[3, 2] = TodaysDate_effectiveDateMain; //Effective Date
                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet in Effective Date: " + TodaysDate_effectiveDateMain);
                        FivedaysDate = DateTime.Now;
                        FivedaysDate = FivedaysDate.AddDays(5);
                        string FivedaysDate_expDate = Convert.ToString(FivedaysDate);
                        string FivedaysDate_expDateMain = Convert.ToDateTime(FivedaysDate_expDate).ToString("MM/dd/yyyy");
                        mWSheet1.Cells[4, 2] = FivedaysDate_expDateMain;  //Expiration Date
                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet in Expiration Date : " + " || " + FivedaysDate_expDateMain);
                    }
                    mWorkBook.Save();
                    ExtentTestManager._parentTest.Log(Status.Pass, "Automation Save the Excel Sheet.");
                    mWorkBook.Close(Missing.Value, Missing.Value, Missing.Value);
                    ExtentTestManager._parentTest.Log(Status.Pass, "Automation Close the Excel Sheet.");
                }


            }
            catch (Exception ex)
            {
                ExtentTestManager._parentTest.Log(Status.Fail, "Excel not updated");
                GeneralMethod.ScreenShotCapture();

            }
            finally
            {

                mWSheet1 = null;
                mWorkBook = null;
                oXL.Quit();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }

        public static void UpdateExcelfilewithMultipledata(Dictionary<string, string> testData, string DocumentUpload, string TestData)
        {
            try
            {
                string ValueName1 = testData["UploadHader_Obligation"];
                string NameofExcel1 = testData["Upload_Obligation"];
                string NumberofDocument = testData["NumberofDocumnetCount"];
                int Size = Int32.Parse(NumberofDocument);
                string AlertedCount = testData["ObligationManagementType_AlertedCount"];
                int Sizeindividual_AlertedCount = Int32.Parse(AlertedCount);
                string MonitoredCount = testData["ObligationManagementType_MonitoredCount"];
                int Sizeindividual_MonitoredCount = Int32.Parse(MonitoredCount);
                string ReportedCount = testData["ObligationManagementType_ReportedCount"];
                int Sizeindividual_ReportedCount = Int32.Parse(ReportedCount);
                string Alerted = "Alerted";
                string Monitored = "Monitored";
                string Reported = "Reported";
                if (Size != 0)
                {
                    if (DocumentUpload.Equals(ValueName1) & NameofExcel1.Equals(TestData))
                    {
                        string path = System.Reflection.Assembly.GetCallingAssembly().CodeBase;
                        string actualPath = path.Substring(0, path.LastIndexOf("bin"));
                        string excelPath = new Uri(actualPath).LocalPath;
                        excelPath = excelPath + "FileUpload\\" + TestData;
                        log.Info("All Test Data Excel file Initalized  !!!");
                        Random random = new Random();
                        int num = random.Next(1000);
                        oXL = new Microsoft.Office.Interop.Excel.Application();
                        oXL.Visible = true;
                        oXL.DisplayAlerts = false;
                        mWorkBook = oXL.Workbooks.Open(excelPath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                        mWorkSheets = mWorkBook.Worksheets;
                        mWSheet1 = (Microsoft.Office.Interop.Excel.Worksheet)mWorkSheets.get_Item("Sheet1");
                        Microsoft.Office.Interop.Excel.Range range = mWSheet1.UsedRange;
                        int colCount = range.Columns.Count;
                        int rowCount = range.Rows.Count;
                        string[] UploadExcelColumnNames = { Alerted, Monitored, Reported };
                        int sizeoftestdata = UploadExcelColumnNames.Count();
                        for (int x = 1; x <= sizeoftestdata; x++)
                        {
                            int UplaodNamesgrid = x - 1;
                            for (int i = 1; i <= Sizeindividual_AlertedCount; i++)
                            {                                
                                if (UploadExcelColumnNames[UplaodNamesgrid].Equals(testData["AlertedColumn"]))
                                {
                                    num = num + 5;
                                    int Rowcount = i + 1;
                                    for (int index = 1; index <= 1; index++)
                                    {
                                        mWSheet1.Cells[Rowcount, 1] = "Automation_AlertedExtract" + num;  // Obligation Extract  >>  Extract
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + "Automation_AlertedExtract" + num);
                                        mWSheet1.Cells[Rowcount, 2] = "Automation_AlertedTopic" + num; //Topic
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet   : " + " || " + "Automation_AlertedTopic" + num);
                                        mWSheet1.Cells[Rowcount, 3] = "Automation_AlertedReferenceDocument" + num; //ReferenceDocument
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet   : " + " || " + "Automation_AlertedReferenceDocument" + num);
                                        mWSheet1.Cells[Rowcount, 4] = "Automation_AlertedReferenceSection" + num; //ReferenceSection
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + "Automation_AlertedReferenceSection" + num);
                                        mWSheet1.Cells[Rowcount, 5] = "Automation_AlertedCrossReference" + num; //CrossReference
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + "Automation_AlertedCrossReference" + num);
                                        mWSheet1.Cells[Rowcount, 6] = "YES"; //IsTrigger
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + "YES");
                                        mWSheet1.Cells[Rowcount, 7] = "YES"; //TriggerEvent
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + "YES");                       
                                        mWSheet1.Cells[Rowcount, 8] = "Alerted"; //ManagementType
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + "Alerted");
                                        mWSheet1.Cells[Rowcount, 9] = "Transition"; //Phase
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet : " + " || " + "Transition");
                                        mWSheet1.Cells[Rowcount, 10] = "One Time Type I"; //Frequency
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + "One Time Type I");
                                        mWSheet1.Cells[Rowcount, 11] = "YES"; //FinancialImpact
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + "YES");
                                        mWSheet1.Cells[Rowcount, 12] = "Automation_AlertedFinancialAmountSaved" + num; //FinancialAmountSaved
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + "Automation_AlertedFinancialAmountSaved" + num);
                                        mWSheet1.Cells[Rowcount, 13] = "YES"; //TransitionMilestone
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + "YES");
                                        mWSheet1.Cells[Rowcount, 14] = "YES"; //CustomerApprovalRequired
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + "YES");
                                        mWSheet1.Cells[Rowcount, 15] = "YES"; //CustomerApprovalRequired
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + "YES");
                                        TodaysDate = DateTime.Now;
                                        string TodaysDate_effectiveDate = Convert.ToString(TodaysDate);
                                        string TodaysDate_effectiveDateMain = Convert.ToDateTime(TodaysDate_effectiveDate).ToString("MM/dd/yyyy");
                                        mWSheet1.Cells[Rowcount, 16] = TodaysDate_effectiveDateMain; //Effective Date
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet in Effective Date: " + TodaysDate_effectiveDateMain);
                                        FivedaysDate = DateTime.Now;
                                        FivedaysDate = FivedaysDate.AddDays(5);
                                        string FivedaysDate_expDate = Convert.ToString(FivedaysDate);
                                        string FivedaysDate_expDateMain = Convert.ToDateTime(FivedaysDate_expDate).ToString("MM/dd/yyyy");
                                        mWSheet1.Cells[Rowcount, 17] = FivedaysDate_expDateMain;  //Expiration Date
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet in Expiration Date : " + " || " + FivedaysDate_expDateMain);
                                        mWSheet1.Cells[Rowcount, 18] = "PMO"; //Obligation Functional Group
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + "PMO");
                                        mWSheet1.Cells[Rowcount, 19] = "DXC and Customer"; // Responsibility
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + "DXC and Customer");
                                        mWSheet1.Cells[Rowcount, 20] = testData["Dropdown_ObligationOwner"]; //Obligation Owner
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + testData["Dropdown_ObligationOwner"]);
                                        mWSheet1.Cells[Rowcount, 21] = testData["Dropdown_ObligationOwner"]; //EscalationOwner
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + testData["Dropdown_ObligationOwner"]);

                                    }
                                }
                                else if (UploadExcelColumnNames[UplaodNamesgrid].Equals(testData["MonitoredColumn"]))
                                {
                                    num = num + 10;
                                    int y = x - 1;
                                    int Rowcount = y+i+ Sizeindividual_AlertedCount;
                                    for (int index = 1; index <= 1; index++)
                                    {
                                        mWSheet1.Cells[Rowcount, 1] = "Automation_MonitoredExtract" + num;  // Obligation Extract  >>  Extract
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + "Automation_MonitoredExtract" + num);
                                        mWSheet1.Cells[Rowcount, 2] = "Automation_MonitoredTopic" + num; //Topic
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet   : " + " || " + "Automation_MonitoredTopic" + num);
                                        mWSheet1.Cells[Rowcount, 3] = "Automation_MonitoredReferenceDocument" + num; //ReferenceDocument
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet   : " + " || " + "Automation_MonitoredReferenceDocument" + num);
                                        mWSheet1.Cells[Rowcount, 4] = "Automation_MonitoredReferenceSection" + num; //ReferenceSection
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + "Automation_MonitoredReferenceSection" + num);
                                        mWSheet1.Cells[Rowcount, 5] = "Automation_MonitoredCrossReference" + num; //CrossReference
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + "Automation_MonitoredCrossReference" + num);
                                        mWSheet1.Cells[Rowcount, 6] = "YES"; //IsTrigger
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + "YES");
                                        mWSheet1.Cells[Rowcount, 7] = "YES"; //TriggerEvent
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + "YES");
                                        mWSheet1.Cells[Rowcount, 8] = "Monitored"; //ManagementType
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + "Monitored");
                                        mWSheet1.Cells[Rowcount, 9] = "Steady"; //Phase
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet : " + " || " + "Steady");
                                        mWSheet1.Cells[Rowcount, 10] = "Monthly"; //Frequency
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + "Monthly");
                                        mWSheet1.Cells[Rowcount, 11] = "YES"; //FinancialImpact
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + "YES");
                                        mWSheet1.Cells[Rowcount, 12] = "Automation_MonitoredFinancialAmountSaved" + num; //FinancialAmountSaved
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + "Automation_MonitoredFinancialAmountSaved" + num);
                                        mWSheet1.Cells[Rowcount, 13] = "YES"; //TransitionMilestone
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + "YES");
                                        mWSheet1.Cells[Rowcount, 14] = "YES"; //CustomerApprovalRequired
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + "YES");
                                        mWSheet1.Cells[Rowcount, 15] = "YES"; //CustomerApprovalRequired
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + "YES");
                                        TodaysDate = DateTime.Now;
                                        string TodaysDate_effectiveDate = Convert.ToString(TodaysDate);
                                        string TodaysDate_effectiveDateMain = Convert.ToDateTime(TodaysDate_effectiveDate).ToString("MM/dd/yyyy");
                                        mWSheet1.Cells[Rowcount, 16] = TodaysDate_effectiveDateMain; //Effective Date
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet in Effective Date: " + TodaysDate_effectiveDateMain);
                                        FivedaysDate = DateTime.Now;
                                        FivedaysDate = FivedaysDate.AddDays(5);
                                        string FivedaysDate_expDate = Convert.ToString(FivedaysDate);
                                        string FivedaysDate_expDateMain = Convert.ToDateTime(FivedaysDate_expDate).ToString("MM/dd/yyyy");
                                        mWSheet1.Cells[Rowcount, 17] = FivedaysDate_expDateMain;  //Expiration Date
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet in Expiration Date : " + " || " + FivedaysDate_expDateMain);
                                        mWSheet1.Cells[Rowcount, 18] = "HR"; //Obligation Functional Group
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + "HR");
                                        mWSheet1.Cells[Rowcount, 19] = "DXC and DXC Subcontracor"; //Responsibility
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + "DXC and DXC Subcontracor");
                                        mWSheet1.Cells[Rowcount, 20] = testData["Dropdown_ObligationOwner"]; //Obligation Owner
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + testData["Dropdown_ObligationOwner"]);
                                        mWSheet1.Cells[Rowcount, 21] = testData["Dropdown_ObligationOwner"]; //EscalationOwner
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + testData["Dropdown_ObligationOwner"]);
                                    }
                                }
                                else if (UploadExcelColumnNames[UplaodNamesgrid].Equals(testData["ReportedColumn"]))
                                {
                                    num = num + 15;
                                    int y = x - 2;
                                    int Rowcount = y + i + Sizeindividual_MonitoredCount + Sizeindividual_ReportedCount;
                                    for (int index = 1; index <= 1; index++)
                                    {
                                        mWSheet1.Cells[Rowcount, 1] = "Automation_ReportExtract" + num;  // Obligation Extract  >>  Extract
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + "Automation_ReportExtract" + num);
                                        mWSheet1.Cells[Rowcount, 2] = "Automation_ReportTopic" + num; //Topic
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet   : " + " || " + "Automation_ReportTopic" + num);
                                        mWSheet1.Cells[Rowcount, 3] = "Automation_ReportReferenceDocument" + num; //ReferenceDocument
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet   : " + " || " + "Automation_ReportReferenceDocument" + num);
                                        mWSheet1.Cells[Rowcount, 4] = "Automation_ReportReferenceSection" + num; //ReferenceSection
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + "Automation_ReportReferenceSection" + num);
                                        mWSheet1.Cells[Rowcount, 5] = "Automation_ReportCrossReference" + num; //CrossReference
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + "Automation_ReportCrossReference" + num);
                                        mWSheet1.Cells[Rowcount, 6] = "YES"; //IsTrigger
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + "YES");
                                        mWSheet1.Cells[Rowcount, 7] = "YES"; //TriggerEvent
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + "YES");
                                        mWSheet1.Cells[Rowcount, 8] = "Reported"; //ManagementType
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + "Reported");
                                        mWSheet1.Cells[Rowcount, 9] = "Transition"; //Phase
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet : " + " || " + "Transition");
                                        mWSheet1.Cells[Rowcount, 10] = "One Time Type II"; //Frequency
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + "One Time Type II");
                                        mWSheet1.Cells[Rowcount, 11] = "YES"; //FinancialImpact
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + "YES");
                                        mWSheet1.Cells[Rowcount, 12] = "Automation_ReportFinancialAmountSaved" + num; //FinancialAmountSaved
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + "Automation_ReportFinancialAmountSaved" + num);
                                        mWSheet1.Cells[Rowcount, 13] = "YES"; //TransitionMilestone
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + "YES");
                                        mWSheet1.Cells[Rowcount, 14] = "YES"; //CustomerApprovalRequired
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + "YES");
                                        mWSheet1.Cells[Rowcount, 15] = "YES"; //CustomerApprovalRequired
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + "YES");
                                        TodaysDate = DateTime.Now;
                                        string TodaysDate_effectiveDate = Convert.ToString(TodaysDate);
                                        string TodaysDate_effectiveDateMain = Convert.ToDateTime(TodaysDate_effectiveDate).ToString("MM/dd/yyyy");
                                        mWSheet1.Cells[Rowcount, 16] = TodaysDate_effectiveDateMain; //Effective Date
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet in Effective Date: " + TodaysDate_effectiveDateMain);
                                        FivedaysDate = DateTime.Now;
                                        FivedaysDate = FivedaysDate.AddDays(5);
                                        string FivedaysDate_expDate = Convert.ToString(FivedaysDate);
                                        string FivedaysDate_expDateMain = Convert.ToDateTime(FivedaysDate_expDate).ToString("MM/dd/yyyy");
                                        mWSheet1.Cells[Rowcount, 17] = FivedaysDate_expDateMain;  //Expiration Date
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet in Expiration Date : " + " || " + FivedaysDate_expDateMain);
                                        mWSheet1.Cells[Rowcount, 18] = "Delivery"; //Obligation Functional Group
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + "Delivery");
                                        mWSheet1.Cells[Rowcount, 19] = "DXC Subcontractor"; //Responsibility
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + "DXC Subcontractor");
                                        mWSheet1.Cells[Rowcount, 20] = testData["Dropdown_ObligationOwner"]; //Obligation Owner
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + testData["Dropdown_ObligationOwner"]);
                                        mWSheet1.Cells[Rowcount, 21] = testData["Dropdown_ObligationOwner"]; //EscalationOwner
                                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Updated the Excel Sheet  : " + " || " + testData["Dropdown_ObligationOwner"]);
                                    }

                                }
                            }
                        }
                        mWorkBook.Save();
                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Save the Excel Sheet.");
                        mWorkBook.Close(Missing.Value, Missing.Value, Missing.Value);
                        ExtentTestManager._parentTest.Log(Status.Pass, "Automation Close the Excel Sheet.");
                    }
                }
                else
                {
                    ExtentTestManager._parentTest.Log(Status.Info, "Number of Document is 0 Please add count in NumberofDocumnetCount column");
                }
            }
            catch (Exception ex)
            {
                ExtentTestManager._parentTest.Log(Status.Fail, "Excel not updated");
                GeneralMethod.ScreenShotCapture();

            }
            finally
            {

                mWSheet1 = null;
                mWorkBook = null;
                oXL.Quit();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }




        /// <summary>
        /// Desc: Method is used to validate PageSorting
        /// </summary>
        /// <param name="testData"></param>
        public void PageSort(Dictionary<string, string> testData, string gridAttribute, string gridvalue, string haderAttribute, string hadervalue)
        {
            try
            {
                PageLoadWait(100);
                List<String> displayNames = new List<string>();
                row = getElements(gridAttribute, gridvalue);
                foreach (IWebElement cell in row)
                {
                    displayNames.Add(cell.Text);
                }
                for (int j = 1; j <= 2; j++)
                {
                    List<String> displayNamesSorted = new List<string>(displayNames);
                    displayNamesSorted.Sort();
                    ClickOnElementWhenElementFound(haderAttribute, hadervalue, "Automation clicked hader");
                    if (row.OrderBy(c => c.Text).SequenceEqual(row))
                    {
                        ExtentTestManager._parentTest.Log(Status.Pass, "Ascending    " + row.OrderBy(c => c.Text).SequenceEqual(row).ToString());
                    }
                    else if (row.OrderByDescending(c => c.Text).SequenceEqual(row))
                    {
                        ExtentTestManager._parentTest.Log(Status.Pass, "Descending   " + row.OrderByDescending(c => c.Text).SequenceEqual(row));
                    }
                }
            }

            catch (Exception ex)
            {
                GeneralMethod.ScreenShotCapture();
                Assert.Fail(ex.Message);
            }
        }


        public void Pagination(Dictionary<string, string> testData, string attribute_lengthofpagination, string Value_lengthofpagination)
        {
            try
            {
                IList<IWebElement> pagination = getElements(attribute_lengthofpagination, Value_lengthofpagination);
                int size = pagination.Count();
                ExtentTestManager._parentTest.Log(Status.Info, "Count of pagination : " + size);
                if (size > 0)
                {
                    ExtentTestManager._parentTest.Log(Status.Pass, "Pagination exists ");
                    for (int i = 1; i <= size; i++)
                    {
                        try
                        {
                            string ValueofPaginagtion = Value_lengthofpagination;
                            ClickOnElementWhenElementFound(attribute_lengthofpagination, ValueofPaginagtion + "[" + i + "]", "Click the pagination button " + i);
                            threadWait(1200);
                            ExtentTestManager._parentTest.Log(Status.Info, "Automation at pagination : " + i);
                        }
                        catch (Exception e)
                        {
                            ExtentTestManager._parentTest.Log(Status.Fail, " " + e.ToString());
                            GeneralMethod.ScreenShotCapture();
                        }
                    }
                }
                else 
                {
                    ExtentTestManager._parentTest.Log(Status.Info, "Pagination not exists ");
                }

            }
            catch (Exception ex)
            {
                GeneralMethod.ScreenShotCapture();
                ExtentTestManager._parentTest.Log(Status.Fail, "Expected not matched " + ex.Message);
            }
        }



    }
}

