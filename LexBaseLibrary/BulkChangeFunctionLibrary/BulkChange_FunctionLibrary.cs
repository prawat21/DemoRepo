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
    public class BulkChange_FunctionLibrary : Keywords
    {
        public IList<IWebElement> row = null;
        
        
        
        /// <summary>
        /// Desc: Method validate Bulk change Icon
        /// </summary>
        /// <param name="testData"></param>
        public void BulkChangePage(Dictionary<string, string> testData)
        {
            try 
            {
                WaitforElement_ExpectedConditions(20,250, "//a[contains(text(),'Bulk Change')]");
                AssertIsTrue("xpath", "//a[contains(text(),'Bulk Change')]", "Bulk Change Nav hader");
                ClickOnElementWhenElementFound("xpath" , "//a[contains(text(),'Bulk Change')]", "Bulk Change Nav hader");
                WaitforElement_ExpectedConditions(20,250, "//span[contains(text(),'--Select an Action--')]");
            }
            catch (Exception ex) 
            {
                GeneralMethod.ScreenShotCapture();
                ExtentTestManager._parentTest.Log(Status.Fail, "Expected not matched " + ex.Message);
                Assert.Fail(ex.Message);
            }
        }



        /// <summary>
        /// Desc: Method validate Bulk change Grid
        /// </summary>
        /// <param name="testData"></param>
        public void BulkchangeGrid(Dictionary<string, string> testData)
        {
            try
            {
                WaitforElementbool(20,250, "//kendo-dropdownlist[contains(@name,'mt')]//span[contains(@class,'k-input')]");
                var kendoElement_Action = getElement("xpath", "//kendo-dropdownlist[@name='a']//span[contains(@class,'k-i-arrow-s k-icon')]");
                IWebElement currentItem_Action = kendoElement_Action.FindElement(By.XPath("//*[contains(@name,'a')]//*[contains(@class,'k-dropdown-wrap k-state-default')]"));
                while (currentItem_Action.Text != testData["Dropdown_Action"])
                { currentItem_Action.SendKeys(Keys.ArrowDown); ExtentTestManager._parentTest.Log(Status.Pass, "Expected Contract Type Matched " + testData["Dropdown_ContractType"] + " Automation selected the data "); }
                var kendoElement_ObligationMT = getElement("xpath", "//kendo-dropdownlist[@name='mt']//span[contains(@class,'k-i-arrow-s k-icon')]");
                IWebElement currentItem_ObligationMT = kendoElement_ObligationMT.FindElement(By.XPath("//*[contains(@name,'mt')]//*[contains(@class,'k-dropdown-wrap k-state-default')]"));
                while (currentItem_ObligationMT.Text != testData["Dropdown_ObligationMT"])
                { currentItem_ObligationMT.SendKeys(Keys.ArrowDown); ExtentTestManager._parentTest.Log(Status.Pass, "Expected Contract Type Matched " + testData["Dropdown_ContractType"] + " Automation selected the data "); } //Obligation Management Type 
                WaitforElement_Visibility(20, 250, "//button[contains(@class,'btn btn-dark')]");
                WaitforElement_ExpectedConditions(20, 250, "//button[contains(@class,'btn btn-dark')]");
                ClickOnElementWhenElementFound("xpath", "//button[contains(@class,'btn btn-dark')]", "Search button");
                PDFandExcelIcon_Bulk(testData);
                PageSort(testData, "xpath", "//*[contains(@class,'k-grid-content k-virtual-content')]//descendant::tr/td[1]", "xpath", "//*[contains(@class,'k-grid-header-wrap')]//descendant::tr/th[2]//a[@class='k-link ng-star-inserted']");
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

        /// <summary>
        /// Desc: Method validate Bulk change Grid
        /// </summary>
        /// <param name="testData"></param>
        public void UploadandAddtional(Dictionary<string, string> testData)
        {
            try
            {
                AssertAreEqual("xpath", "//div[contains(text(),'Upload Evidence')]", "Upload Evidence");
                AssertAreEqual("xpath", "//div[contains(text(),'Additional Information')]", "Additional Information");
                AssertIsTrue("xpath", "//button[contains(@class,'btn btn-dark ng-star-inserted')]", "Complete Obligations");
                AssertIsTrue("xpath", "//button[contains(@class,'btn btn-outline-dark')]", "Reset");
            }
            catch (Exception ex)
            {
                GeneralMethod.ScreenShotCapture();
                ExtentTestManager._parentTest.Log(Status.Fail, "Expected not matched " + ex.Message);
                Assert.Fail(ex.Message);
            }
        }

        public void PDFandExcelIcon_Bulk(Dictionary<string, string> testData)
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