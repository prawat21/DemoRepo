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
    public class ChangeRequest_FunctionLibrary : Keywords
    {
        public IList<IWebElement> row = null;


        /// <summary>
        /// Desc: Method validate Page navigate to change Request
        /// </summary>
        /// <param name="testData"></param>
        public void ChangeRequestPage(Dictionary<string, string> testData)
        {
            try
            {
                WaitforElement_ExpectedConditions(30, 250, "//*[contains(text(),'Change Request')]");
                WaitforElement_Exists(30, 250, "//*[contains(text(),'Change Request')]");
                AssertIsTrue("xpath", "//*[contains(text(),'Change Request')]", "Change Request");threadWait(1200);
                ClickOnElementWhenElementFound("xpath", "//*[contains(text(),'Change Request')]", "Change Request");

            }
            catch (Exception ex)
            {
                GeneralMethod.ScreenShotCapture();
                ExtentTestManager._parentTest.Log(Status.Fail, "Expected not matched " + ex.Message);
                Assert.Fail(ex.Message);
            }
        }

        /// <summary>
        /// Desc: Method validate Grid of change Request
        /// </summary>
        /// <param name="testData"></param>
        public void HaderforChangeRequest(Dictionary<string, string> testData)
        {
            try
            {
                PdfandExcelIconforChangeRequest(testData);
                string HaderName = getElement("xpath", "//div[contains(@class,'titleLeft')]").Text;
                if (HaderName.Contains("Change Request Log"))
                {
                    ExtentTestManager._parentTest.Log(Status.Pass, "Expected matched with actual - " + HaderName);
                    row = getElements("xpath", "//*[@class='k-grid-header-wrap']//descendant::tr/th//a[contains(@class,'k-link ng-star-inserted')]");
                    int size = row.Count;
                    for (int i = 1; i <= size; i++)
                    {
                        string[] HaderNames = { "Contract", "Request ID", "Lead CCM", "Request Category", "Request Type", "Status", "Request Date", "Target Date" };
                        string HeaderName = HaderNames[i - 1];
                        AssertAreEqual("xpath", "//*[@class='k-grid-header-wrap']//descendant::tr/th[" + i + "]//a[contains(@class,'k-link ng-star-inserted')]", HeaderName);
                    }
                }
                else
                {
                    ExtentTestManager._parentTest.Log(Status.Fail, "Expected not  matched with actual  :  " + HaderName);
                    GeneralMethod.ScreenShotCapture();
                }

            }
            catch (Exception ex)
            {
                GeneralMethod.ScreenShotCapture();
                ExtentTestManager._parentTest.Log(Status.Fail, "Expected not matched " + ex.Message);
                Assert.Fail(ex.Message);
            }
        }


        /// <summary>
        /// Desc: Method validate change Request TAT window and Open Change Request Window
        /// Validation point >> column sorting >> Pagination
        /// </summary>
        /// <param name="testData"></param>
        public void OtherWindowChangeRequest(Dictionary<string, string> testData)
        {
            try
            {
                AssertAreEqual("xpath", "//div[contains(text(),'Change Request TAT')]", "Change Request TAT");
                AssertAreEqual("xpath", "//div[contains(text(),'Open Change Request')]", "Open Change Request");
                AssertIsTrue("xpath", "//a[contains(text(),'Add Change Request')]" ,"Add Change Request");
                PageSort(testData,"xpath", "//*[contains(@class,'k-grid-content k-virtual-content')]//descendant::tr/td[2]","xpath", "//*[contains(@class,'k-grid-header-wrap')]//descendant::tr/th[2]//a[@class='k-link ng-star-inserted']");
                WaitforElementbool(20,250, "//span[contains(@class,'k-icon k-i-seek-w')]");
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
        /// Desc: Method validate Add Change Request
        /// Validation point >> column sorting >> Pagination
        /// </summary>
        /// <param name="testData"></param>
        public void AddChangeRequest(Dictionary<string, string> testData)
        {
            try 
            {
                ClickOnElementWhenElementFound("xpath", "//a[contains(text(),'Add Change Request')]", "Add Change Request");
                WaitforElementbool(20,250, "//div[contains(@class,'titleLeft')]");
                WaitforElement_Exists(20,250, "//button[contains(@class,'btn btn-outline-dark')]");
                AssertAreEqual("xpath", "//div[contains(@class,'titleLeft')]", "Add Change Request");
                AssertAreEqual("xpath", "//div[contains(text(),'Upload Reference Document')]", "Upload Reference Document");
                AssertAreEqual("xpath", "//div[contains(text(),'Additional Information')]", "Additional Information");
                AssertIsTrue("xpath", "//button[contains(@class,'btn btn-dark')]","Add Request"); threadWait(1900);
                WaitforElement_ExpectedConditions(20,250, "//a[contains(text(),'Change Request Log')]");
                WaitforElement_Exists(10,250, "//a[contains(text(),'Change Request Log')]");
                var kendoElement_RequestCategory = getElement("xpath", "//kendo-dropdownlist[contains(@name,'requestCategory')]//span[contains(@class,'k-i-arrow-s k-icon')]");
                IWebElement currentItem_RequestCategory = kendoElement_RequestCategory.FindElement(By.XPath("//*[contains(@name,'requestCategory')]//*[contains(@class,'k-dropdown-wrap k-state-default')]"));
                while (currentItem_RequestCategory.Text != testData["Dropdown_RequestCategory"])
                {
                    currentItem_RequestCategory.SendKeys(Keys.ArrowDown); 
                    ExtentTestManager._parentTest.Log(Status.Pass, "Expected Account Name Matched " + testData["Dropdown_RequestCategory"] + " Automation selected the data ");
                }// Request Category
                WaitforElement_ExpectedConditions(20, 250, "//div[@class='col-sm col-sm-3 key required ng-star-inserted']");
                var kendoElement_RequestType = getElement("xpath", "//kendo-dropdownlist[contains(@name,'requestType')]//span[contains(@class,'k-i-arrow-s k-icon')]");
                IWebElement currentItem_RequestType = kendoElement_RequestType.FindElement(By.XPath("//*[contains(@name,'requestType')]//*[contains(@class,'k-dropdown-wrap k-state-default')]"));
                while (currentItem_RequestType.Text != testData["Dropdown_RequestType"])
                {
                    currentItem_RequestCategory.SendKeys(Keys.ArrowDown);
                    ExtentTestManager._parentTest.Log(Status.Pass, "Expected Account Name Matched " + testData["Dropdown_RequestType"] + " Automation selected the data ");
                }//Request Type
                WaitforElement_ExpectedConditions(20,250, "//div[contains(text(),'Contract')]");
                var kendoElement_contract = getElement("xpath", "//kendo-dropdownlist[contains(@name,'account')]//span[contains(@class,'k-i-arrow-s k-icon')]");
                IWebElement currentItem_contract = kendoElement_contract.FindElement(By.XPath("//*[contains(@name,'account')]//*[contains(@class,'k-dropdown-wrap k-state-default')]"));
                while (currentItem_contract.Text != testData["Dropdown_contract"])
                {
                    currentItem_RequestCategory.SendKeys(Keys.ArrowDown);
                    ExtentTestManager._parentTest.Log(Status.Pass, "Expected Account Name Matched " + testData["Dropdown_contract"] + " Automation selected the data ");
                }//contract
                SendKeysForElement("xpath", "//input[contains(@placeholder,'Enter Subject Of Email here...')]",testData["EmailSubjectCR"],"Subject of Email");
                SendKeysForElement("xpath", "//textarea[contains(@placeholder,'Enter Details here...')]", testData["DetailsCR"], "Details");
                SendKeysForElement("xpath", "//input[contains(@placeholder,'Enter Reason here...')]", testData["ReasonCR"], "Reason");
                ClickOnElementWhenElementFound("xpath", "//button[contains(text(),'Add Request')]","Add Request");
                WaitforElement(10, 250, "//*[@class='ulx-notification-title']");
                PdfandExcelIconforChangeRequest(testData);
            }
            catch (Exception ex)
            {
                GeneralMethod.ScreenShotCapture();
                ExtentTestManager._parentTest.Log(Status.Fail, "Expected not matched " + ex.Message);
                Assert.Fail(ex.Message);
            }
        }


        /// <summary>
        /// Desc: Method validate PDF and Excel Icon.
        /// </summary>
        /// <param name="testData"></param>
        public void PdfandExcelIconforChangeRequest(Dictionary<string, string> testData)
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