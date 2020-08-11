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
    public class StandardReports_FunctionLibrary : Keywords
    {
        public IList<IWebElement> row = null;
        public DateTime TodaysDate;
        public DateTime FivedaysDate;
        
        
        /// <summary>
        /// Desc: Method validate Page navigate to Contracts
        /// </summary>
        /// <param name="testData"></param>
        public void StandardReportPage(Dictionary<string, string> testData)
        {
            try 
            {
                PageLoadWait(10);
                AssertIsTrue("id", "navComponentDropDown3", "Report Nav Button");
                ClickOnElementWhenElementFound("id", "navComponentDropDown3", "Report Nav Button");
                WaitforElement(50, 250, "//a[contains(text(),'Standard')]");
                AssertIsTrue("xpath", "//a[contains(text(),'Standard')]", "Standard Nav drop Down");
                ClickOnElementWhenElementFound("xpath", "//a[contains(text(),'Standard')]", "Standard Nav drop Down");
            }
            catch (Exception ex) 
            {
                GeneralMethod.ScreenShotCapture();
                ExtentTestManager._parentTest.Log(Status.Fail, "Expected not matched " + ex.Message);
                Assert.Fail(ex.Message);
            }
        }

        //
        public void StandardReportGrid(Dictionary<string, string> testData)
        {
            try 
            {
                WaitforElementbool(20,250, "//div[@class='ulx-panel-header ng-star-inserted']");
                string HaderName = "Select a Contract or Obligation Report";
                AssertAreEqual("xpath", "//div[@class='ulx-panel-header ng-star-inserted']", HaderName);
                AssertIsTrue("xpath", "//div[contains(text(),'Contract Report')]","Contract Report");
                AssertIsTrue("xpath", "//*[@class='d-flex align-items-stretch ulx-panel-contents']//div[contains(text(),'Obligation Report')]", "Obligation Reports");
            }
            catch (Exception ex) 
            {
                GeneralMethod.ScreenShotCapture();
                ExtentTestManager._parentTest.Log(Status.Fail, "Expected not matched " + ex.Message);
                Assert.Fail(ex.Message);
            }
        }





        public void PDFandExcelIcon(Dictionary<string, string> testData)
        {
            try
            {
                WaitforElementbool(50, 250, "//button[contains(@class,'k-button-icon k-button k-grid-pdf')]");
                AssertIsTrue("xpath", "//button[contains(@class,'k-button-icon k-button k-grid-excel')]", "Excel Element Icon is visible");
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