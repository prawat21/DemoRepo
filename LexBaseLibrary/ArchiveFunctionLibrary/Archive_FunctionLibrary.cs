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
    public class Archive_FunctionLibrary : Keywords
    {
        public IList<IWebElement> row = null;
        //public DateTime TodaysDate;
        //public DateTime FivedaysDate;
        
        
        /// <summary>
        /// Desc: Method validate Page navigate to Archive Page
        /// </summary>
        /// <param name="testData"></param>
        public void ArchivePage(Dictionary<string, string> testData)
        {
            try 
            {
                WaitforElement_ExpectedConditions(20,250, "//a[contains(text(),'Archived Contracts')]");
                AssertIsTrue("xpath", "//a[contains(text(),'Archived Contracts')]", "Archived contract Nav hader");
                ClickOnElementWhenElementFound("xpath" , "//a[contains(text(),'Archived Contracts')]", "Obligation Nav Hader");
                PDFandExcelIcon_Archive(testData);
                PageSort(testData, "xpath", "//*[contains(@class,'k-grid-content k-virtual-content')]//descendant::tr/td[1]", "xpath", "//*[contains(@class,'k-grid-header-wrap')]//descendant::tr/th[1]//a[@class='k-link ng-star-inserted']");
                WaitforElementbool(20, 250, "//span[contains(@class,'k-icon k-i-seek-w')]");
                Pagination(testData, "xpath", "//*[contains(@class,'ng-star-inserted')]//descendant::ul/li");
            }
            catch (Exception ex) 
            {
                GeneralMethod.ScreenShotCapture();
                ExtentTestManager._parentTest.Log(Status.Fail, "Expected not matched " + ex.Message);
                Assert.Fail(ex.Message);
            }
        }


        public void PDFandExcelIcon_Archive(Dictionary<string, string> testData)
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