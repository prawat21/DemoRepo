using AventStack.ExtentReports;
using Newtonsoft.Json.Linq;
using NPOI.SS.Util;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using LexBaseFramework.CommonUtilities;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OpenQA.Selenium.Support.UI;

namespace LexBaseFramework.LexBaseLibrary
{
    public class HomePage_FunctionLibrary : Keywords
    {
        public IList<IWebElement> row = null;


        /// <summary>
        /// Desc: Method is used for validate Home Page
        /// Validation Point : CCS Dashboard Grid 
        /// >> Filter>>Label Name 
        /// </summary>
        /// <param name="testData"></param>
        public void HomePage_CCSDashboard(Dictionary<string, string> testData)
        {
            try
            {
                PageLoadWait(100);
                AssertAreEqual(JsonTocode_LT("LocatorTypes_value", "HyperReflabel_CCSDashboard"), JsonTocode_LV("LocatorValues_value", "HyperReflabel_CCSDashboard"), testData["HyperReflabel_CCSDashboard"]);
                AssertAreEqual(JsonTocode_LT("LocatorTypes_value", "HyperReflabel_OMSDashboard"), JsonTocode_LV("LocatorValues_value", "HyperReflabel_OMSDashboard"), testData["HyperReflabel_OMSDashboard"]);
                
               
            }
            catch (Exception ex)
            {
                GeneralMethod.ScreenShotCapture();
                Assert.Fail(ex.Message);
            }
        }

  
        /// <summary>
        /// Desc: Method is used for validate PDF Icon
        /// </summary>
        /// <param name="testData"></param>
        public void PdfIconinGrid(Dictionary<string, string> testData)
        {
            try 
            {
                PageLoadWait(40);
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(50));
                wait.PollingInterval = TimeSpan.FromMilliseconds(250);
                wait.IgnoreExceptionTypes(typeof(NoSuchElementException), typeof(ElementNotVisibleException));
                wait.Until(WaitforIcon());
                AssertIsTrue("xpath", "//button[contains(@class,'k-button-icon k-button k-grid-excel')]" , "Excel Element Icon is visible");
                
            }
            catch (Exception ex) 
            {
                GeneralMethod.ScreenShotCapture();
                Assert.Fail(ex.Message);
            }
        }


        /// <summary>
        /// Desc: Method is used for validate PDF Icon
        /// </summary>
        /// <param name="testData"></param>
        public void Sorting(Dictionary<string, string> testData)
        {
            try
            {
                
                PageLoadWait(40);
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(50));
                wait.PollingInterval = TimeSpan.FromMilliseconds(250);
                wait.IgnoreExceptionTypes(typeof(NoSuchElementException), typeof(ElementNotVisibleException));
                wait.Until(WaitforIcon());
                PageSort(testData, "xpath", "//*[contains(@class,'k-grid-content k-virtual-content')]//descendant::tr/td[1]//span", "xpath", "//*[contains(@class,'k-grid-header-wrap')]//descendant::tr/th[1]//a[@class='k-link']");  
            }
            catch (Exception ex)
            {
                GeneralMethod.ScreenShotCapture();
                Assert.Fail(ex.Message);
            }
        }


        /// <summary>
        /// Desc: Method is used for validate Home Page
        /// Validation Point : CCS Dashboard Filter Icon
        /// </summary>
        /// <param name="testData"></param>
        public void HomePage_FilterIcon(Dictionary<string, string> testData)
        {
            try
            {
                PageLoadWait(100);
                row = getElements("xpath", "//div[contains(@class,'k-grid-aria-root')]//descendant::tr/th"); // Filter Icon List
                IList<IWebElement> list = getElements("xpath","//*[@class='k-grid-header-wrap']//descendant::tr/th//span[contains(@class,'k-icon k-i-filter')]");
                int size = row.Count;
                var flag = 0;
                for (int i=0; i<=size; i++)
                {
                    string HaderName = testData["Dashboard_PageFilterIcon"]; 
                    if (HaderName.Equals(row.ElementAt(i).Text))
                    {
                        ExtentTestManager._parentTest.Log(Status.Pass, "Web Element is visible hence Expected condition is met : " + HaderName);
                        flag = flag + 1;
                        list.ElementAt(i).Click();
                        ExtentTestManager._parentTest.Log(Status.Pass, "Web Element is visible hence Expected condition is met Automation clicked  : " + HaderName + " Filter Icon");
                        SendKeysForElement("xpath", "//kendo-grid-string-filter-menu-input[1]/kendo-grid-filter-menu-input-wrapper/input", testData["Dashboard_filterIteamName"], "Test Data passed");
                        ClickOnElementWhenElementFound("xpath", "//*[@class='k-button k-primary']", "Click on Filter button");
                        break;
                    }
                    else if (HaderName.Equals(row.ElementAt(i).Text))
                    {
                        ExtentTestManager._parentTest.Log(Status.Pass, "Web Element is visible hence Expected condition is met : " + HaderName);
                        flag = flag + 1;
                        list.ElementAt(i).Click();
                        ExtentTestManager._parentTest.Log(Status.Pass, "Web Element is visible hence Expected condition is met Automation clicked  : " + HaderName + " Filter Icon");
                        SendKeysForElement("xpath", "//kendo-grid-string-filter-menu-input[1]/kendo-grid-filter-menu-input-wrapper/input", testData["Dashboard_filterIteamName"], "Test Data passed");
                        ClickOnElementWhenElementFound("xpath", "//*[@class='k-button k-primary']", "Click on Filter button");
                        break;
                    }
                    else if (HaderName.Equals(row.ElementAt(i).Text))
                    {
                        ExtentTestManager._parentTest.Log(Status.Pass, "Web Element is visible hence Expected condition is met : " + HaderName);
                        flag = flag + 1;
                        list.ElementAt(i).Click();
                        ExtentTestManager._parentTest.Log(Status.Pass, "Web Element is visible hence Expected condition is met Automation clicked  : " + HaderName + " Filter Icon");
                        SendKeysForElement("xpath", "//kendo-grid-string-filter-menu-input[1]/kendo-grid-filter-menu-input-wrapper/input", testData["Dashboard_filterIteamName"], "Test Data passed");
                        ClickOnElementWhenElementFound("xpath", "//*[@class='k-button k-primary']", "Click on Filter button");
                        break;
                    }
                    else if (HaderName.Equals(row.ElementAt(i).Text))
                    {
                        ExtentTestManager._parentTest.Log(Status.Pass, "Web Element is visible hence Expected condition is met : " + HaderName);
                        flag = flag + 1;
                        list.ElementAt(i).Click();
                        ExtentTestManager._parentTest.Log(Status.Pass, "Web Element is visible hence Expected condition is met Automation clicked  : " + HaderName + " Filter Icon");
                        SendKeysForElement("xpath", "//kendo-grid-string-filter-menu-input[1]/kendo-grid-filter-menu-input-wrapper/input", testData["Dashboard_filterIteamName"], "Test Data passed");
                        ClickOnElementWhenElementFound("xpath", "//*[@class='k-button k-primary']", "Click on Filter button");
                        break;
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
        /// Desc: Method is used for Navigate to Contracts
        /// </summary>
        /// <param name="testData"></param>
        public void NavigationContracts(Dictionary<string, string> testData)
        {
            try
            {
                ElementWaitTime(4);
                AssertAreEqual(JsonTocode_LT("LocatorTypes_value", "Nav_Contracts"), JsonTocode_LV("LocatorValues_value", "Nav_Contracts"), testData["Nav_Contracts"]);
                WebDriverWait wait = new WebDriverWait(driver,TimeSpan.FromSeconds(50));
                wait.PollingInterval = TimeSpan.FromMilliseconds(250);
                wait.IgnoreExceptionTypes(typeof(NoSuchElementException), typeof(ElementNotVisibleException));
                IWebElement Contratsbtn= wait.Until(waitforElement());
                threadWait(1200);
                Contratsbtn.Click();
            }
            catch (Exception ex)
            {
                GeneralMethod.ScreenShotCapture();
                Assert.Fail(ex.Message);
            }
        }
      
        private Func<IWebDriver, IWebElement> waitforElement()
        {
            return ((x) =>
            {
                ExtentTestManager._parentTest.Log(Status.Pass, "Waiting for element");
                if (x.FindElements(By.XPath("//li[2]//a[1]")).Count == 1)
                return x.FindElement(By.XPath("//*[text()='Contracts ']"));
                return null;
            });
        }

        private Func<IWebDriver, bool> WaitforIcon()
        {
            return ((x) =>
            {
                ExtentTestManager._parentTest.Log(Status.Pass, "Waiting for PDF Element Icon to be visible");
                return x.FindElements(By.XPath("//button[contains(@class,'k-button-icon k-button k-grid-pdf')]")).Count == 1;
            });
        }
     
    }   

}