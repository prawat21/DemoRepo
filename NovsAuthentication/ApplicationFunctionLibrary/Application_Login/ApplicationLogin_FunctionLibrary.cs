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



namespace LexBaseFramework.NovsAuthentication
{
    public class ApplicationLogin_FunctionLibrary : Keywords
    {
        public IList<IWebElement> row = null;
        

        /// <summary>
        /// Desc: Method is used for Read all Excel data to json file[Locators] AND Data Setup will Update the Excel sheet[TestData].
        /// Method start from Reading Existing contract
        /// Data Setup : Update TestData Sheet.
        /// Json Read Excel File : Locators
        /// </summary>
        /// <param name="testData"></param>
        public void TDupdation_CCMUserLogin(Dictionary<string, string> testData)
        {
            try
            {
                PageLoadWait(100);
                ReadOldExistingExcel(testData["CaseName"], testData["BillingID"], testData["DBName"]); // to update test data excel file.
                JsonFunctionLibray_Reader(); // to read all excel data to json file.
            }
            catch (Exception ex)
            {
                GeneralMethod.ScreenShotCapture();
                Assert.Fail(ex.Message);
            }
        }

        /// <summary>
        /// Desc: Method is used for Read all Excel data to json file[Locators] AND Data Setup will Update the Excel sheet[TestData].
        /// Method start from created new case 
        /// Data Setup : Update TestData Sheet.
        /// Json Read Excel File : Locators
        /// </summary>
        /// <param name="testData"></param>
        /// <param name="CaseName">Column For Openacasename Column </param>
        public void TDupdation_AzureADLogin(Dictionary<string, string> testData)
        {
            try
            {
                PageLoadWait(100);
                ReadExistingExcel(); // to update test data excel file.
                JsonFunctionLibray_Reader(); // to read all excel data to json file.
            }
            catch (Exception ex)
            {
                GeneralMethod.ScreenShotCapture();
                Assert.Fail(ex.Message);
            }
        }


        /// <summary>
        /// Desc: Method is used for Application Login for Azure AD user
        /// </summary>
        /// <param name="testData"></param>
        public void AzureADLogin(Dictionary<string, string> testData)
        {
            try
            {
                PageLoadWait(100);
                WaitforElementbool(40,250, "//h3[contains(text(),'External Login')]");
                ClickOnElementWhenElementFound(JsonTocode_LT("LocatorTypes_value", "Label_Externalheader"), JsonTocode_LV("LocatorValues_value", "Label_Externalheader"), "External Login UlxAzureAD button");
                AssertAreEqual(JsonTocode_LT("LocatorTypes_value", "Label_Externalheader"), JsonTocode_LV("LocatorValues_value", "Label_Externalheader"), testData["LabelValidation_ExternalLogin"]);
                AssertIsTrue(JsonTocode_LT("LocatorTypes_value", "Btn_ExternalLoginUlxAzureAD"), JsonTocode_LV("LocatorValues_value", "Btn_ExternalLoginUlxAzureAD"), "External Login UlxAzureAD button");
                ClickOnElementWhenElementFound(JsonTocode_LT("LocatorTypes_value", "Btn_ExternalLoginUlxAzureAD"), JsonTocode_LV("LocatorValues_value", "Btn_ExternalLoginUlxAzureAD"), "External Login UlxAzureAD button");   
            }
            catch (Exception ex)
            {
                GeneralMethod.ScreenShotCapture();
                Assert.Fail(ex.Message);
            }
        }

        /// <summary>
        /// Desc: Method is used for Application Login for CCM User
        /// </summary>
        /// <param name="testData"></param>
        public void CCM_UserLogin(Dictionary<string, string> testData)
        {
            try
            {
                PageLoadWait(100);
                AssertIsTrue(JsonTocode_LT("LocatorTypes_value", "txt_Email"), JsonTocode_LV("LocatorValues_value", "txt_Email"), "User Name");
                SendKeysForElement(JsonTocode_LT("LocatorTypes_value", "txt_Email"), JsonTocode_LV("LocatorValues_value", "txt_Email"), testData["Username"], "User Name");
                AssertIsTrue(JsonTocode_LT("LocatorTypes_value", "txt_Password"), JsonTocode_LV("LocatorValues_value", "txt_Password"), "Password");
                SendKeysForAElement(JsonTocode_LT("LocatorTypes_value", "txt_Password"), JsonTocode_LV("LocatorValues_value", "txt_Password"), testData["Password"], "Password");
            }
            catch (Exception ex)
            {
                GeneralMethod.ScreenShotCapture();
                Assert.Fail(ex.Message);
            }
        }

        /// <summary>
        /// Desc: Method is used to validate user has logined with valid credential 
        /// </summary>
        /// <param name="testData"></param>
        public void Credentail_Verfiy(Dictionary<string, string> testData)
        {
            try
            {
                PageLoadWait(100);
                WaitforElementbool(40,250, "//a[@id='navUserDropDown']");
                AssertAreEqual(JsonTocode_LT("LocatorTypes_value", "Label_UserName_Nav"),JsonTocode_LV("LocatorValues_value", "Label_UserName_Nav"), testData["UserName_HomePage"]);
                GeneralMethod.ScreenShotCapture();
            }
            catch (Exception ex)
            {
                GeneralMethod.ScreenShotCapture();
                Assert.Fail(ex.Message);
            }
        }


        //verfiy the Scroll
        public void Scrollanywhere(Dictionary<string, string> testData)
        {
            IJavaScriptExecutor jse = (IJavaScriptExecutor)driver;
            jse.ExecuteScript("scroll(0,910)");
        }
    }

}