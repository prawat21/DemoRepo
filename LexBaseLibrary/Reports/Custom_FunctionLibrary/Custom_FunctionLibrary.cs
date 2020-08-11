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
    public class Custom_FunctionLibrary : Keywords
    {
        public IList<IWebElement> row = null;
        public DateTime TodaysDate;
        public DateTime FivedaysDate;
        
        
        /// <summary>
        /// Desc: Method validate Page navigate to Contracts
        /// </summary>
        /// <param name="testData"></param>
        public void CustomReportPage(Dictionary<string, string> testData)
        {
            try 
            {
                PageLoadWait(10);
                AssertIsTrue("id", "navComponentDropDown3", "Report Nav Button");
                ClickOnElementWhenElementFound("id", "navComponentDropDown3", "Report Nav Button");
                WaitforElement(50, 250, "//a[contains(text(),'Custom')]");
                AssertIsTrue("xpath", "//a[contains(text(),'Custom')]", "Custom Nav drop Down");
                ClickOnElementWhenElementFound("xpath", "//a[contains(text(),'Custom')]", "Custom Nav drop Down");
            }
            catch (Exception ex) 
            {
                GeneralMethod.ScreenShotCapture();
                ExtentTestManager._parentTest.Log(Status.Fail, "Expected not matched " + ex.Message);
                Assert.Fail(ex.Message);
            }
        }

        //
        public void CustomReportGrid(Dictionary<string, string> testData)
        {
            try 
            {
                WaitforElementbool(20, 250, "//div[@class='col-sm col-sm-3 key']");
                AssertIsTrue("xpath", "//div[@class='col-sm col-sm-3 key']", "Report Type");
                WaitforElement_ExpectedConditions(20,250, "//span[contains(@class,'k-i-arrow-s k-icon')]");
                var kendoElement_ReportType = getElement("xpath", "//span[contains(@class,'k-i-arrow-s k-icon')]");
                IWebElement currentItem_ReportType = kendoElement_ReportType.FindElement(By.XPath("//*[contains(@class,'k-dropdown-wrap k-state-default')]"));
                while (currentItem_ReportType.Text != testData["Dropdown_ReportType"])
                {
                    currentItem_ReportType.SendKeys(Keys.ArrowDown);
                    ExtentTestManager._parentTest.Log(Status.Pass, "Expected Account Name Matched " + testData["Dropdown_ReportType"] + " Automation selected the data ");
                }// Report Type
                WaitforElement_ExpectedConditions(20,250, "//button[contains(text(),'Save As Report Template')]");threadWait(900);
                WaitforElementbool(20,250, "//button[@class='btn btn-dark ng-star-inserted']");
            }
            catch (Exception ex) 
            {
                GeneralMethod.ScreenShotCapture();
                ExtentTestManager._parentTest.Log(Status.Fail, "Expected not matched " + ex.Message);
                Assert.Fail(ex.Message);
            }
        }


    }

}