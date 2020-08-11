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
    public class Obligation_FunctionLibrary : Keywords
    {
        public IList<IWebElement> row = null;
        public DateTime TodaysDate;
        public DateTime FivedaysDate;


        /// <summary>
        /// Desc: Method validate Page navigate to Obligation
        /// </summary>
        /// <param name="testData"></param>
        public void ObligationPage(Dictionary<string, string> testData)
        {
            try 
            {
                PageLoadWait(100);
                AssertIsTrue("xpath", "//a[contains(text(),'Obligations')]", "Obligation Nav hader");
                ClickOnElementWhenElementFound("xpath" , "//a[contains(text(),'Obligations')]" ,"Obligation Nav Hader");
                PDFandExcelIcon(testData);
                AssertAreEqual("xpath", "//a[contains(text(),'Alerted Obligations')]" , "Alerted Obligations");
                AssertAreEqual("xpath", "//a[contains(text(),'Reported Obligations')]", "Reported Obligations");
                AssertAreEqual("xpath", "//a[contains(text(),'Triggered Obligations')]", "Triggered Obligations");
                AssertAreEqual("xpath", "//a[contains(text(),'Checklist')]", "Checklist");
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