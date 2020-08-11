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
    public class ObligationSearch_FunctionLibrary : Keywords
    {
        public IList<IWebElement> row = null;
        public DateTime TodaysDate;


        /// <summary>
        /// Desc: Method validate Page navigate to Search Obligation Page
        /// </summary>
        /// <param name="testData"></param>
        public void SearchPage_Obligation(Dictionary<string, string> testData)
        {
            try 
            {
                WaitforElement_ExpectedConditions(30, 250, "//a[@id='navComponentDropDown5']");
                WaitforElement_Exists(30, 250, "//a[@id='navComponentDropDown5']");
                AssertIsTrue("xpath", "//a[@id='navComponentDropDown5']", "Search Nav Button");
                ClickOnElementWhenElementFound("xpath", "//a[@id='navComponentDropDown5']", "Search Nav Button");
                WaitforElement(50, 250, "//a[@class='dropdown-item'][contains(text(),'Obligation')]");
                AssertIsTrue("xpath", "//a[@class='dropdown-item'][contains(text(),'Obligation')]", "Obligation Search Nav Button");
                ClickOnElementWhenElementFound("xpath", "//a[@class='dropdown-item'][contains(text(),'Obligation')]", "Obligation Search Nav Button");

            }
            catch (Exception ex) 
            {
                GeneralMethod.ScreenShotCapture();
                ExtentTestManager._parentTest.Log(Status.Fail, "Expected not matched " + ex.Message);
                Assert.Fail(ex.Message);
            }
        }

        /// <summary>
        /// Desc: Method validate Obligation Filters
        /// </summary>
        /// <param name="testData"></param>
        public void SearchPage_ObligationFilters(Dictionary<string, string> testData)
        {
            try
            {
                WaitforElement_ExpectedConditions(30, 250, "//kendo-panelbar-item[@id='k-panelbar-1-item-default-3']");
                WaitforElement_Exists(30, 250, "//kendo-panelbar-item[@id='k-panelbar-1-item-default-3']");
                AssertIsTrue("xpath", "//kendo-panelbar-item[@id='k-panelbar-1-item-default-3']", "Obligation Filters");
                WaitforElement_ExpectedConditions(30, 250, "//div[contains(text(),'Functional Group')]");
                TodaysDate = DateTime.Now;
                string TodaysDate_effectiveDate = Convert.ToString(TodaysDate);
                string TodaysDate_effectiveDateMain = Convert.ToDateTime(TodaysDate_effectiveDate).ToString("MM/dd/yyyy");
                string effectivedate = TodaysDate_effectiveDateMain.Replace("/", "");
                ClearValueOnElementWhenElementFound("xpath", "//kendo-datepicker[contains(@name,'dueDateUpto')]//input[contains(@class,'k-input')]", "Effective Date value cleared");
                SendKeysForElement("xpath", "//kendo-datepicker[contains(@name,'dueDateUpto')]//input[contains(@class,'k-input')]", effectivedate, "Effective Date");
                WaitforElement_ExpectedConditions(30, 250, "//button[@class='btn btn-dark']");threadWait(900);
                ClickOnElementWhenElementFound("xpath", "//button[@class='btn btn-dark']", "Search Button");
                WaitforElement_ExpectedConditions(30, 250, "//div[@class='ulx-panel-header ng-star-inserted']");
                WaitforElementbool(20,250, "//*[contains(@class,'k-grid-header-wrap')]//descendant::tr/th[2]//a[contains(@class,'k-link ng-star-inserted')]");
                PageSort(testData, "xpath", "//*[contains(@class,'k-grid-content k-virtual-content')]//descendant::tr/td[2]//span", "xpath", "//*[contains(@class,'k-grid-header-wrap')]//descendant::tr/th[2]//a[contains(@class,'k-link ng-star-inserted')]");
                Pagination(testData, "xpath", "//*[contains(@class,'k-widget k-grid')]//descendant::ul/li");
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