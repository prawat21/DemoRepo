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
    public class UserSearch_FunctionLibrary : Keywords
    {
        public IList<IWebElement> row = null;
        
        
        /// <summary>
        /// Desc: Method validate Page navigate to Search User Page
        /// </summary>
        /// <param name="testData"></param>
        public void SearchPage_User(Dictionary<string, string> testData)
        {
            try 
            {
                WaitforElement_ExpectedConditions(30, 250, "//a[@id='navComponentDropDown5']");
                WaitforElement_Exists(30, 250, "//a[@id='navComponentDropDown5']");
                AssertIsTrue("xpath", "//a[@id='navComponentDropDown5']", "Search Nav Button");
                ClickOnElementWhenElementFound("xpath", "//a[@id='navComponentDropDown5']", "Search Nav Button");
                WaitforElement(50, 250, "//div[@class='dropdown-menu show']//a[@class='dropdown-item'][contains(text(),'User')]");
                AssertIsTrue("xpath", "//div[@class='dropdown-menu show']//a[@class='dropdown-item'][contains(text(),'User')]", "User Search Nav Button");
                ClickOnElementWhenElementFound("xpath", "//div[@class='dropdown-menu show']//a[@class='dropdown-item'][contains(text(),'User')]", "User Search Nav Button");
            }
            catch (Exception ex) 
            {
                GeneralMethod.ScreenShotCapture();
                ExtentTestManager._parentTest.Log(Status.Fail, "Expected not matched " + ex.Message);
                Assert.Fail(ex.Message);
            }
        }


        /// <summary>
        /// Desc: Method validate validate the Search user grid and enter the rqueried for the same.
        /// </summary>
        /// <param name="testData"></param>
        public void SearchwithName(Dictionary<string, string> testData)
        {
            try 
            {
                WaitforElement_ExpectedConditions(30, 250, "//button[@class='btn btn-dark']");
                SendKeysForElement("xpath", "//input[@name='firstName']", testData["firstName_Search"], "First Name");
                SendKeysForElement("xpath", "//input[@name='lastName']", testData["lastName_Search"], "Last Name");
                SendKeysForElement("xpath", "//input[@name='email']", testData["EmailID_Search"], "Email ID");threadWait(900);
                WaitforElement_ExpectedConditions(30,250, "//button[text()='Search']");
                ClickOnElementWhenElementFound("xpath", "//button[text()='Search']", "Search Button");
                AssertIsTrue("xpath", "//button[@class='btn btn-outline-dark']", "Reset Button");
                WaitforElementbool(20,250, "//div[@class='ulx-panel-header ng-star-inserted']");
                Pagination(testData, "xpath", "//*[contains(@class,'k-widget k-grid')]//descendant::ul/li");
                ClickOnElementWhenElementFound("xpath", "//button[@class='btn btn-outline-dark']", "Reset Button");
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