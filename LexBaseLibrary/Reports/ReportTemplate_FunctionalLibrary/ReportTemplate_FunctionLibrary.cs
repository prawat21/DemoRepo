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
    public class ReportTemplate_FunctionLibrary : Keywords
    {
        public IList<IWebElement> row = null;
       
        
        
        /// <summary>
        /// Desc: Method validate Page navigate to Contracts
        /// </summary>
        /// <param name="testData"></param>
        public void ReportPage(Dictionary<string, string> testData)
        {
            try
            {
                WaitforElement_ExpectedConditions(30, 250, "//*[contains(@id,'navComponentDropDown3e')]");
                WaitforElement_Exists(30, 250, "//*[contains(@id,'navComponentDropDown3e')]");
                ClickOnElementWhenElementFound("xpath", "//*[contains(@id,'navComponentDropDown3e')]", "Report Nav Button");
                ElementWaitTime(4);
                ClickOnElementWhenElementFound("xpath", "//a[contains(text(),'Report Template')]", "Report Template Nav drop Down"); 

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