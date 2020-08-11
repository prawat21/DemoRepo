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
using NUnit.Framework.Internal;

namespace LexBaseFramework.LexBaseLibrary
{
    public class Contracts_FunctionLibrary : Keywords
    {
        public IList<IWebElement> row = null;
        public DateTime TodaysDate;
        public DateTime FivedaysDate;
        /// <summary>
        /// Desc: Method validate Page navigate to Contracts
        /// </summary>
        /// <param name="testData"></param>
        public void ContractsPage(Dictionary<string, string> testData)
        {
            try 
            {
                PageLoadWait(100);
                string Label_Contracts = getElement(JsonTocode_LT("LocatorTypes_value", "Label_Contracts"), JsonTocode_LV("LocatorValues_value", "Label_Contracts")).Text;
                if (Label_Contracts.Contains(testData["Nav_Contracts"]))
                {
                    ExtentTestManager._parentTest.Log(Status.Pass, "Expected Contracts label is visible, contratcs navigation landed in correct Page ");
                    ElementWaitTime(100);
                    AssertIsTrue(JsonTocode_LT("LocatorTypes_value", "Btn_AddContract"), JsonTocode_LV("LocatorValues_value", "Btn_AddContract"),"Add contract button");
                    ClickOnElementWhenElementFound(JsonTocode_LT("LocatorTypes_value", "Btn_AddContract"), JsonTocode_LV("LocatorValues_value", "Btn_AddContract"), "Add contract button is clicked by Automation");
                }
                else 
                {
                    ExtentTestManager._parentTest.Log(Status.Fail, "Expected Contracts label is not visible");
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
        /// Desc: Method validate New Add Contract Tab
        /// </summary>
        /// <param name="testData"></param>
        public void NewAddContratctTab(Dictionary<string, string> testData)
        {
            try 
            {
                PageLoadWait(100);
                AssertAreEqual(JsonTocode_LT("LocatorTypes_value", "Label_Contracts"), JsonTocode_LV("LocatorValues_value", "Label_Contracts"),testData["Label_AddContracts"]);
                threadWait(1900);
                var kendoElement_Account = getElement("xpath", "//kendo-dropdownlist[@name='account_mg']//span[@class='k-i-arrow-s k-icon']");
                IWebElement currentItem_Account= kendoElement_Account.FindElement(By.XPath("//*[contains(@name,'account_mg')]//*[contains(@class,'k-dropdown-wrap k-state-default')]"));
                while (currentItem_Account.Text != testData["Dropdown_Account"])
                {currentItem_Account.SendKeys(Keys.ArrowDown); ExtentTestManager._parentTest.Log(Status.Pass, "Expected Account Name Matched " + testData["Dropdown_Account"] +" Automation selected the data "); } // Account 
                SendKeysForElement("xpath", "//input[@name='shortName']" ,testData["TextBox_ContractShortName"], "Contract Short Name Text box"); //Contract Short Name
                SendKeysForElement("xpath", "//input[@name='legalName']", testData["TextBox_ContractLegalName"], "Contract Legal Name Text box"); //Contract Legal Name  // Customer Contracting Entity
                var kendoElement_ContractType = getElement("xpath", "//kendo-dropdownlist[@name='accountType']//span[contains(@class,'k-i-arrow-s k-icon')]");
                IWebElement currentItem_ContractType = kendoElement_ContractType.FindElement(By.XPath("//*[contains(@name,'accountType')]//*[contains(@class,'k-dropdown-wrap k-state-default')]"));
                while (currentItem_ContractType.Text != testData["Dropdown_ContractType"])
                {  currentItem_ContractType.SendKeys(Keys.ArrowDown); ExtentTestManager._parentTest.Log(Status.Pass, "Expected Contract Type Matched " + testData["Dropdown_ContractType"] + " Automation selected the data "); } //Contract Type 
                var kendoElement_ContractClassification = getElement("xpath", "//kendo-dropdownlist[@name='accountClassification']//span[contains(@class,'k-i-arrow-s k-icon')]");
                IWebElement currentItem_ContractClassification = kendoElement_ContractClassification.FindElement(By.XPath("//*[contains(@name,'accountClassification')]//*[contains(@class,'k-dropdown-wrap k-state-default')]")); 
                while (currentItem_ContractClassification.Text != testData["Dropdown_ContractClassification"])
                { currentItem_ContractClassification.SendKeys(Keys.ArrowDown); ExtentTestManager._parentTest.Log(Status.Pass, "Expected Contract Classification Type Matched " + testData["Dropdown_ContractClassification"] + " Automation selected the data "); } //Contract Classification
              
                SendKeysForElement("xpath", "//*[@class='k-numeric-wrap']//*[@class='k-input k-formatted-value']", testData["TextBox_TotalContractValue"], "Total Contract Value Text box"); // Total Contract Value
                Scrollanywheres(testData);threadWait(900);
                SendKeysForElement("xpath", "//*[contains(@name,'leadCCM')]//*[@class='k-input']", testData["Dropdown_LeadCMM"], "Lead CMM");
                ClickOnElementWhenElementFound("xpath", "//*[contains(@name,'leadCCM')]//*[@class='k-input']", "Lead CMM");//Lead CCM 
                var kendoElement_LeadCountry = getElement("xpath", "//*[contains(@name,'leadCountry')]//span[contains(@class,'k-i-arrow-s k-icon')]");
                IWebElement currentItem_LeadCountry = kendoElement_LeadCountry.FindElement(By.XPath("//*[contains(@name,'leadCountry')]//*[contains(@class,'k-dropdown-wrap k-state-default')]"));
                while (currentItem_LeadCountry.Text != testData["Dropdown_LeadCountry"])
                { currentItem_LeadCountry.SendKeys(Keys.ArrowDown); ExtentTestManager._parentTest.Log(Status.Pass, "Expected lead Countryn Type Matched " + testData["Dropdown_LeadCountry"] + " Automation selected the data "); }  //Lead Country
                var kendoElement_Coverage = getElement("xpath", "//*[contains(@name,'coverage')]//span[contains(@class,'k-i-arrow-s k-icon')]");
                IWebElement currentItem_Coverage = kendoElement_Coverage.FindElement(By.XPath("//*[contains(@name,'coverage')]//*[contains(@class,'k-dropdown-wrap k-state-default')]"));
                while (currentItem_Coverage.Text != testData["Dropdown_coverage"])
                { currentItem_Coverage.SendKeys(Keys.ArrowDown); ExtentTestManager._parentTest.Log(Status.Pass, "Expected coverage Type Matched " + testData["Dropdown_coverage"] + " Automation selected the data "); } // coverage
                var kendoElement_Region = getElement("xpath", "//*[@name='region']//span[contains(@class,'k-i-arrow-s k-icon')]");
                IWebElement currentItem_Region = kendoElement_Region.FindElement(By.XPath("//*[@name='region']//*[contains(@class,'k-dropdown-wrap k-state-default')]"));
                while (currentItem_Region.Text != testData["Dropdown_Region"])
                { currentItem_Region.SendKeys(Keys.ArrowDown); ExtentTestManager._parentTest.Log(Status.Pass, "Expected Region Matched " + testData["Dropdown_Region"] + " Automation selected the data "); } //Region
                SendKeysForElement("xpath", "//*[contains(@name,'requestedBy')]//*[@class='k-input']", testData["Dropdown_LeadCMM"], "Requested By");
                ClickOnElementWhenElementFound("xpath", "//*[contains(@name,'requestedBy')]//*[@class='k-input']", "Requested By");//Lead CCM 
                ClickOnElementWhenElementFound("xpath", "//button[contains(text(),'Submit')]","Submit Button");
            }
            catch (Exception ex) 
            {
                GeneralMethod.ScreenShotCapture();
                ExtentTestManager._parentTest.Log(Status.Fail, "Expected not matched " + ex.Message);
                Assert.Fail(ex.Message);
            }
        }

        /// <summary>
        /// Desc: Method validate Contratcs grid Name after creating new Contracts 
        /// </summary>
        /// <param name="testData"></param>
        public void Contracts_GridHader(Dictionary<string, string> testData)
        {
            try
            {
                PageLoadWait(100);
                row = getElements("xpath", "//*[@class='k-grid-header-wrap']//descendant::tr/th//a[contains(@class,'k-link ng-star-inserted')]");
                int size = row.Count;
                for (int i = 1; i <= size; i++)
                {
                    string[] HaderNames = { "Contract Name", "Account", "Creation Date", "Region", "Lead Country", "Lead CCM", "Contract Status" };
                    string HeaderName = HaderNames[i - 1];
                    AssertAreEqual("xpath", "//*[@class='k-grid-header-wrap']//descendant::tr/th["+i+"]//a[contains(@class,'k-link ng-star-inserted')]", HeaderName);
                }
                ClickOnElementWhenElementFound("xpath", "//*[@class='k-grid-header-wrap']//descendant::tr/th[1]//span[contains(@class,'k-icon k-i-filter')]", "Filter Icon");
                SendKeysForElement("xpath", "//kendo-grid-string-filter-menu-input[1]/kendo-grid-filter-menu-input-wrapper/input", testData["TextBox_ContractShortName"], "Test Data passed");
                ClickOnElementWhenElementFound("xpath", "//*[@class='k-button k-primary']", "Click on Filter button");
            }
            catch (Exception ex)
            {
                GeneralMethod.ScreenShotCapture();
                ExtentTestManager._parentTest.Log(Status.Fail, "Expected not matched " + ex.Message);
                Assert.Fail(ex.Message);
            }
        }


        /// <summary>
        /// Desc: Method for Newly Created contract displayed
        /// </summary>
        /// <param name="testData"></param>
        public void NewlyCreatedGrid(Dictionary<string, string> testData)
        {
            try 
            {
                PageLoadWait(100);
                row = getElements("xpath", "//*[@class='k-grid-table-wrap']//descendant::tr/td");
                int size = row.Count;
                int ListName  = 0;
                for (int i = 1; i <= size; i++)
                {
                    string List = getElement("xpath", "//*[@class='k-grid-table-wrap']//descendant::tr/td["+i+"]").Text;
                    if (List.Equals(testData["TextBox_ContractShortName"]))
                    {
                        ExtentTestManager._parentTest.Log(Status.Pass, "Expected " + List + " matched with Actual " + testData["TextBox_ContractShortName"]);
                        ListName = ListName + 1;
                    }
                    else if (List.Equals(testData["Dropdown_Account"]))
                    {
                        ExtentTestManager._parentTest.Log(Status.Pass, "Expected " + List + " matched with Actual " + testData["Dropdown_Account"]);
                        ListName = ListName + 1;
                    }
                    else if (List.Equals(testData["Dropdown_Region"]))
                    {
                        ExtentTestManager._parentTest.Log(Status.Pass, "Expected " + List + " matched with Actual " + testData["Dropdown_Region"]);
                        ListName = ListName + 1;
                    }
                    else if (List.Equals(testData["Dropdown_LeadCountry"]))
                    {
                        ExtentTestManager._parentTest.Log(Status.Pass, "Expected " + List + " matched with Actual " + testData["Dropdown_LeadCountry"]);
                        ListName = ListName + 1;
                    }
                    else if (List.Equals(testData["Dropdown_LeadCMM"]))
                    {
                        ExtentTestManager._parentTest.Log(Status.Pass, "Expected " + List + " matched with Actual " + testData["Dropdown_LeadCMM"]);
                        ListName = ListName + 1;
                    }
                    else if (List.Contains(testData["NewlyCreatedContractStatus"]))
                    {
                        ExtentTestManager._parentTest.Log(Status.Pass, "Expected " + List + " matched with Actual " + testData["NewlyCreatedContractStatus"]);
                        ListName = ListName + 1;
                        ClickOnElementWhenElementFound("xpath", "//*[@class='k-grid-table-wrap']//descendant::tr/td[1]//span", "Contract Short Name");
                    }
                }
                if (ListName.Equals(0))
                {
                    ExtentTestManager._parentTest.Log(Status.Fail, "Expected not matched in entried grid");
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
        /// Desc: Method for Validate ADD L1 Document 
        /// </summary>
        /// <param name="testData"></param>
        public void AddL1Document(Dictionary<string, string> testData)
        {
            try 
            {
                PageLoadWait(10);
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(50));
                wait.PollingInterval = TimeSpan.FromMilliseconds(250);
                wait.IgnoreExceptionTypes(typeof(NoSuchElementException), typeof(ElementNotVisibleException));
                wait.Until(waitforElement_Delete());
                AssertIsTrue("xpath", "//button[contains(text(),'Add L1 Document')]", "Add L1 Document Button visble");
                ClickOnElementWhenElementFound("xpath", "//button[contains(text(),'Add L1 Document')]", "Add L1 Document Button visble");
                wait.Until(WaitforLabel());
                SendKeysForElement(JsonTocode_LT("LocatorTypes_value", "Textbox_DocumentName"), JsonTocode_LV("LocatorValues_value", "Textbox_DocumentName"),testData["Textbox_DocumentNameL1"], "Automation Enter Document Name");
                AssertAreEqual("xpath", "//div[contains(text(),'"+testData["Dropdown_Region"]+"')]", testData["Dropdown_Region"]);
                TodaysDate = DateTime.Now;
                string TodaysDate_effectiveDate = Convert.ToString(TodaysDate);
                string TodaysDate_effectiveDateMain = Convert.ToDateTime(TodaysDate_effectiveDate).ToString("MM/dd/yyyy");
                string effectivedate = TodaysDate_effectiveDateMain.Replace("/", "");
                ClearValueOnElementWhenElementFound("xpath", "//kendo-datepicker[contains(@name,'effectiveDate')]//input[contains(@class,'k-input')]", "Effective Date value cleared");
                SendKeysForElement("xpath", "//kendo-datepicker[contains(@name,'effectiveDate')]//input[contains(@class,'k-input')]", effectivedate, "Effective Date");
                FivedaysDate = DateTime.Now;
                FivedaysDate = FivedaysDate.AddDays(5);
                string FivedaysDate_expDate = Convert.ToString(FivedaysDate);
                string FivedaysDate_expDateMain = Convert.ToDateTime(FivedaysDate_expDate).ToString("MM/dd/yyyy");
                string expDate = FivedaysDate_expDateMain.Replace("/", "");
                ClearValueOnElementWhenElementFound("xpath", "//kendo-datepicker[@name='expDate']//input[contains(@class,'k-input')]", "exp Date value cleared");
                SendKeysForElement("xpath", "//kendo-datepicker[@name='expDate']//input[contains(@class,'k-input')]", expDate, "exp Date");
                SendKeysForElement("xpath", "//kendo-combobox[contains(@placeholder,'Select User')]//*[@class='k-input']", testData["Dropdown_LeadCMM"], "Requested By");
                getElement("xpath", "//kendo-combobox[contains(@placeholder,'Select User')]//*[@class='k-input']").SendKeys(Keys.Enter); //Requested By 
                getElement("xpath", "//button[contains(text(),'Submit')]").SendKeys(Keys.Enter); //Submit Button
                wait.Until(waitforElement_KeyInfoDocName(testData)); // Document Name data 
                ValdationContractTree(testData); // Validate Tree view and Modify L1 Document.
                AddObligation(testData); // Add Obligation After Modify L1 Document.
                ValidateObligations(testData); // Validate Added Obligation and Modify the Obligation 
                AddwithExcel(testData); // Upload Document via Excel L1 Dcoument, L2 Document, Obligation
                NavigateBacktoCSName(testData); // Navigate Back to Contract Name.
                HighRiskItemsTab(testData);//Navigate to High Risk Items Tab
                ActiveContract(testData);//Activate the contract
                AuditLogs(testData);
            }
            catch (Exception ex)
            {
                GeneralMethod.ScreenShotCapture();
                ExtentTestManager._parentTest.Log(Status.Fail, "Expected not matched " + ex.Message);
                Assert.Fail(ex.Message);
            }
        }

        /// <summary>
        /// Desc: Method for Validate Contract Short Name and ADD L1 Document in tree view. 
        /// Valdation >> Modfiy L1 document
        /// </summary>
        /// <param name="testData"></param>
        public void ValdationContractTree(Dictionary<string, string> testData)
        {
            try
            {
                PageLoadWait(10);
                AssertAreEqual("xpath", "//span[contains(text(),'" + testData["TextBox_ContractShortName"] + "')]", testData["TextBox_ContractShortName"]);
                AssertAreEqual("xpath", "//span[contains(@class,'k-in k-state-selected')]//span[contains(@class,'ng-star-inserted')][contains(text(),'" + testData["Textbox_DocumentNameL1"] + "')]", testData["Textbox_DocumentNameL1"]); 
                ClickOnElementWhenElementFound(JsonTocode_LT("LocatorTypes_value", "Btn_ModfiyL1"), JsonTocode_LV("LocatorValues_value", "Btn_ModfiyL1"), "Btn_ModfiyL1");
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(50));
                wait.PollingInterval = TimeSpan.FromMilliseconds(250);
                wait.IgnoreExceptionTypes(typeof(NoSuchElementException), typeof(ElementNotVisibleException));
                wait.Until(waitforElement_DocumentTitleModifys());
                wait.Until(WaitforLabel());
                SendKeysForWebElement(JsonTocode_LT("LocatorTypes_value", "Textbox_DocumentNameModify"), JsonTocode_LV("LocatorValues_value", "Textbox_DocumentNameModify"), testData["Textbox_ModifyDocumentNameL1"], "Automation Enter Modify Document Name ");
                SendKeysForWebElement(JsonTocode_LT("LocatorTypes_value", "TextArea_AdditionalInformation"), JsonTocode_LV("LocatorValues_value", "TextArea_AdditionalInformation"), testData["Textarea_AdditionalInformationL1"], "Automation Enter Additional Information for Modify L1 Dcoument");
                SendKeysForElement("xpath", "//kendo-combobox[contains(@placeholder,'Select User')]//*[@class='k-input']", testData["Dropdown_LeadCMM"], "Requested By");
                getElement("xpath", "//kendo-combobox[contains(@placeholder,'Select User')]//*[@class='k-input']").SendKeys(Keys.Enter); //Requested By 
                getElement("xpath", "//button[contains(text(),'Submit')]").SendKeys(Keys.Enter); //Submit Button

            }
            catch (Exception ex) 
            {
                GeneralMethod.ScreenShotCapture();
                ExtentTestManager._parentTest.Log(Status.Fail, "Expected not matched " + ex.Message);
                Assert.Fail(ex.Message);

            }
        }

        /// <summary>
        /// Desc: Method for Add Obligatio 
        /// Valdation >> Add Obligatio After  Modfiy L1 document
        /// </summary>
        /// <param name="testData"></param>
        public void AddObligation(Dictionary<string, string> testData)
        {
            try
            {
                PageLoadWait(10);
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(50));
                wait.PollingInterval = TimeSpan.FromMilliseconds(250);
                wait.IgnoreExceptionTypes(typeof(NoSuchElementException), typeof(ElementNotVisibleException));
                wait.Until(waitforElement_AuditLogs());
                wait.Until(waitforElement_treeview());
                AssertAreEqual("xpath", "//span[contains(text(),'" + testData["TextBox_ContractShortName"] + "')]", testData["TextBox_ContractShortName"]);
                AssertAreEqual("xpath", "//span[contains(@class,'k-in k-state-selected')]//span[contains(@class,'ng-star-inserted')][contains(text(),'" + testData["Textbox_ModifyDocumentNameL1"] + "')]", testData["Textbox_ModifyDocumentNameL1"]);
                ClickOnElementWhenElementFound(JsonTocode_LT("LocatorTypes_value", "Btn_AddObligation"), JsonTocode_LV("LocatorValues_value", "Btn_AddObligation"),"Add Obligation");
                wait.Until(waitforElement_ModifyDocumentName(testData));
                SendKeysForElement(JsonTocode_LT("LocatorTypes_value", "TextArea_ObligationExtract"), JsonTocode_LV("LocatorValues_value", "TextArea_ObligationExtract"), testData["TextArea_ObligationExtract"],"Obligation Extract"); // Obligation Extract
                SendKeysForElement(JsonTocode_LT("LocatorTypes_value", "Textbox_ObligationTopic&Description"), JsonTocode_LV("LocatorValues_value", "Textbox_ObligationTopic&Description"), testData["Textbox_ObligationTopic&Description"], "Obligation Topic and Description"); // Obligation Topic and Description 
                var kendoElement_Frequency = getElement("xpath", "//kendo-dropdownlist[@id='frequency']//span[@class='k-i-arrow-s k-icon']");
                IWebElement currentItem_Frequency = kendoElement_Frequency.FindElement(By.XPath("//*[contains(@id,'frequency')]//*[contains(@class,'k-dropdown-wrap k-state-default')]"));
                while (currentItem_Frequency.Text != testData["Dropdown_Frequency"])
                { currentItem_Frequency.SendKeys(Keys.ArrowDown); ExtentTestManager._parentTest.Log(Status.Pass, "Expected Account Name Matched " + testData["Dropdown_Frequency"] + " Automation selected the data "); } // Frequency
                var kendoElement_ObligationManagementType = getElement("xpath", "//kendo-dropdownlist[@id='managementType']//span[@class='k-i-arrow-s k-icon']");
                IWebElement currentItem_ObligationManagementType = kendoElement_ObligationManagementType.FindElement(By.XPath("//*[contains(@id,'managementType')]//*[contains(@class,'k-dropdown-wrap k-state-default')]"));
                while (currentItem_ObligationManagementType.Text != testData["Dropdown_ObligationManagementType"])
                { currentItem_ObligationManagementType.SendKeys(Keys.ArrowDown); ExtentTestManager._parentTest.Log(Status.Pass, "Expected Account Name Matched " + testData["Dropdown_ObligationManagementType"] + " Automation selected the data "); } // Obligation Management Type
                var kendoElement_ObligationFunctionalGroup = getElement("xpath", "//kendo-dropdownlist[@id='functionalGroup']//span[@class='k-i-arrow-s k-icon']");
                IWebElement currentItem_ObligationFunctionalGroup = kendoElement_ObligationFunctionalGroup.FindElement(By.XPath("//*[contains(@id,'functionalGroup')]//*[contains(@class,'k-dropdown-wrap k-state-default')]"));
                while (currentItem_ObligationFunctionalGroup.Text != testData["Dropdown_ObligationFunctionalGroup"])
                { currentItem_ObligationFunctionalGroup.SendKeys(Keys.ArrowDown); ExtentTestManager._parentTest.Log(Status.Pass, "Expected Account Name Matched " + testData["Dropdown_ObligationFunctionalGroup"] + " Automation selected the data "); } // Obligation Functional Group 
                TodaysDate = DateTime.Now;
                string TodaysDate_InitialDueDates = Convert.ToString(TodaysDate);
                string TodaysDate_InitialDueDate = Convert.ToDateTime(TodaysDate_InitialDueDates).ToString("MM/dd/yyyy");
                string InitialDueDate = TodaysDate_InitialDueDate.Replace("/", "");
                ClickOnElementWhenElementFound("xpath", "//kendo-datepicker[contains(@id,'dueDate')]//input[contains(@class,'k-input')]", "Initial Due Date date picker clicked");
                ClearValueOnElementWhenElementFound("xpath", "//kendo-datepicker[contains(@id,'dueDate')]//input[contains(@class,'k-input')]", "Initial Due Date value cleared");
                SendKeysForElement("xpath", "//kendo-datepicker[contains(@id,'dueDate')]//input[contains(@class,'k-input')]", InitialDueDate, "Initial Due Date"); // Initial Due Date
                FivedaysDate = DateTime.Now;
                FivedaysDate = FivedaysDate.AddDays(5);
                string FivedaysDate_CompletionDate = Convert.ToString(FivedaysDate);
                string FivedaysDate_CompletionDates = Convert.ToDateTime(FivedaysDate_CompletionDate).ToString("MM/dd/yyyy");
                string CompletionDate = FivedaysDate_CompletionDates.Replace("/", "");
                ClickOnElementWhenElementFound("xpath", "//kendo-datepicker[contains(@id,'completionDate')]//input[contains(@class,'k-input')]", "Completion Date picker clicked");
                ClearValueOnElementWhenElementFound("xpath", "//kendo-datepicker[@id='completionDate']//input[contains(@class,'k-input')]", "Completion Date value cleared");
                SendKeysForElement("xpath", "//kendo-datepicker[@id='completionDate']//input[contains(@class,'k-input')]", CompletionDate, "Completion Date Date"); // Completion Date
                SendKeysForElement("xpath", "//kendo-combobox[contains(@id,'owner')]//input[contains(@class,'k-input')]", testData["Dropdown_ObligationOwner"], "Obligation Owner");
                //getElement("xpath", "//kendo-combobox[contains(@id,'owner')]//*[contains(@class,'k-searchbar')]").SendKeys(Keys.Enter); //Obligation Owner 
                SendKeysForElement("xpath", "//kendo-combobox[contains(@id,'escalationOwner')]//input[contains(@class,'k-input')]", testData["Dropdown_EscalationOwner"], "Escalation Owner");
                //getElement("xpath", "//kendo-combobox[contains(@id,'escalationOwner')]//*[contains(@class,'k-searchbar')]").SendKeys(Keys.Enter); //Escalation Owner
                var kendoElement_Responsibility = getElement("xpath", "//kendo-dropdownlist[@id='responsibility']//span[@class='k-i-arrow-s k-icon']");
                IWebElement currentItem_Responsibility = kendoElement_Responsibility.FindElement(By.XPath("//*[contains(@id,'responsibility')]//*[contains(@class,'k-dropdown-wrap k-state-default')]"));
                while (currentItem_Responsibility.Text != testData["Dropdown_Responsibility"])
                { currentItem_Responsibility.SendKeys(Keys.ArrowDown); ExtentTestManager._parentTest.Log(Status.Pass, "Expected Account Name Matched " + testData["Dropdown_Responsibility"] + " Automation selected the data "); } // Responsibility
                var kendoElement_Phase = getElement("xpath", "//kendo-dropdownlist[@id='phase']//span[@class='k-i-arrow-s k-icon']");
                IWebElement currentItem_Phase = kendoElement_Phase.FindElement(By.XPath("//*[contains(@id,'phase')]//*[contains(@class,'k-dropdown-wrap k-state-default')]"));
                while (currentItem_Phase.Text != testData["DropDown_Phase"])
                { currentItem_Phase.SendKeys(Keys.ArrowDown); ExtentTestManager._parentTest.Log(Status.Pass, "Expected Account Name Matched " + testData["DropDown_Phase"] + " Automation selected the data "); } // Phase
                Scrollanywheres(testData);
                SendKeysForElement(JsonTocode_LT("LocatorTypes_value", "Textbox_ReferenceSection"), JsonTocode_LV("LocatorValues_value", "Textbox_ReferenceSection"), testData["Textbox_ReferenceSection"], "Reference Section"); // Reference Section
                SendKeysForElement("xpath", "//kendo-combobox[contains(@id,'users')]//input[contains(@class,'k-input')]", testData["Dropdown_LeadCMM"], "Requested By");
                getElement("xpath", "//kendo-combobox[contains(@id,'users')]//input[contains(@class,'k-input')]").SendKeys(Keys.Enter); //Requested By 
                getElement("xpath", "//button[contains(text(),'Submit')]").SendKeys(Keys.Enter); //Submit Button
            }
            catch (Exception ex)
            {
                GeneralMethod.ScreenShotCapture();
                ExtentTestManager._parentTest.Log(Status.Fail, "Expected not matched " + ex.Message);
                Assert.Fail(ex.Message);

            }
        }


        /// <summary>
        /// Desc: Method for Obligations 
        /// validation >> Obligations
        /// validation >> Modify the Obligation
        /// </summary>
        /// <param name="testData"></param>
        public void ValidateObligations(Dictionary<string, string> testData)
        {
            try 
            {
                PageLoadWait(10);
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(50));
                wait.PollingInterval = TimeSpan.FromMilliseconds(250);
                wait.IgnoreExceptionTypes(typeof(NoSuchElementException), typeof(ElementNotVisibleException));
                wait.Until(waitforElement_AuditLogs());
                wait.Until(waitforElement_treeview());
                IWebElement Obligations = wait.Until(waitforElement_Obligations());
                Obligations.Click();
                wait.Until(waitforElement_ObligationsOwner());
                row = getElements("xpath", "//*[@class='k-grid-content k-virtual-content']//descendant::tr/td[2]");
                int size = row.Count;
                int ObligationColumn = 0;
                for (int i=1; i<=size; i++)
                {
                    string ObligationName = getElement("xpath", "//*[@class='k-grid-content k-virtual-content']//descendant::tr["+i+"]/td[2]").Text;
                    if (ObligationName.Equals(testData["Textbox_ObligationTopic&Description"]))
                    {

                        ExtentTestManager._parentTest.Log(Status.Pass, "Expected " + ObligationName + " matched with Actual " + testData["Textbox_ObligationTopic&Description"]);
                        ObligationColumn =  ObligationColumn + 1;
                        ClickOnElementWhenElementFound("xpath", "//*[@class='k-grid-content k-virtual-content']//descendant::tr["+i+"]/td[1]//a", "MasterID");
                        wait.Until(waitforElement_ObligationOwner());
                        AssertAreEqual("xpath", "//div[contains(text(),'"+ testData["TextBox_ContractShortName"] + "')]",testData["TextBox_ContractShortName"]);
                        AssertAreEqual("xpath", "//div[contains(text(),'"+ testData["Textbox_ModifyDocumentNameL1"]+"')]", testData["Textbox_ModifyDocumentNameL1"]);
                        AssertAreEqual("xpath", "//div[contains(@class,'col-sm col-sm-9 value')]", testData["TextArea_ObligationExtract"]);
                        AssertAreEqual("xpath", "//div[contains(text(),'" + testData["Textbox_ObligationTopic&Description"] + "')]", testData["Textbox_ObligationTopic&Description"]);
                        AssertAreEqual("xpath", "//div[contains(text(),'" + testData["TextBox_ContractShortName"] + "')]", testData["TextBox_ContractShortName"]);
                        AssertAreEqual("xpath", "//div[contains(text(),'" + testData["Dropdown_ObligationFunctionalGroup"] + "')]", testData["Dropdown_ObligationFunctionalGroup"]);
                        AssertAreEqual("xpath", "//div[contains(text(),'" + testData["Dropdown_Frequency"] + "')]", testData["Dropdown_Frequency"]);
                        AssertAreEqual("xpath", "//div[contains(text(),'" + testData["Dropdown_ObligationManagementType"] + "')]", testData["Dropdown_ObligationManagementType"]);
                        AssertAreEqual("xpath", "//div[12]", testData["Dropdown_ObligationOwner"]);
                        AssertAreEqual("xpath", "//div[14]", testData["Dropdown_EscalationOwner"]);
                        AssertAreEqual("xpath", "//div[contains(text(),'" + testData["Dropdown_Responsibility"] + "')]", testData["Dropdown_Responsibility"]);
                        AssertAreEqual("xpath", "//div[contains(text(),'" + testData["DropDown_Phase"] + "')]", testData["DropDown_Phase"]);
                        //ClickOnElementWhenElementFound("xpath", "","Modify Obligation");
                        getElement("xpath", "//button[contains(text(),'Modify Obligation')]").SendKeys(Keys.Enter);
                        wait.Until(waitforElement_Frequency());
                        SendKeysForWebElement(JsonTocode_LT("LocatorTypes_value", "TextArea_ObligationExtract"), JsonTocode_LV("LocatorValues_value", "TextArea_ObligationExtract"), testData["TextArea_ModifyObligationExtract"], "Obligation Extract modify"); // Obligation Extract
                        SendKeysForElement("xpath", "//textarea[@id='additionalInfo']", testData["TextArea_AdditionalInformationModifyObligation"], "Additional Information");
                        RequestedBy(testData);

                    }
                    else 
                    {
                        ExtentTestManager._parentTest.Log(Status.Debug, "Expected " + ObligationName + " matched with Actual " + testData["Textbox_ObligationTopic"]);
                    }
                }
                if (ObligationColumn.Equals(0))
                {
                    ExtentTestManager._parentTest.Log(Status.Fail, "Expected not matched in entried grid");
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
        /// Desc: Method for Add Obligatio 
        /// Valdation >> Add Obligatio After  Modfiy L1 document , validating obligation AND Modify Obligation 
        /// </summary>
        /// <param name="testData"></param>
        public void AddwithExcel(Dictionary<string, string> testData)
        {
            try
            {
                PageLoadWait(10);
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(50));
                wait.PollingInterval = TimeSpan.FromMilliseconds(250);
                wait.IgnoreExceptionTypes(typeof(NoSuchElementException), typeof(ElementNotVisibleException));
                wait.Until(waitforElement_AuditLogs());
                wait.Until(waitforElement_treeview());
                wait.Until(waitforElement_Obligations());
                IWebElement AddwithExcelTab = wait.Until(waitforElement_AddwithExcel());
                AddwithExcelTab.Click();
                IWebElement AddwithExcelgridDownloadL1 = wait.Until(waitforElement_AddwithExcelgridDownload());
                string AddL1DocumentUpload = getElement("xpath", "//div[contains(text(),'Add L1 Document')]").Text;
                string AddL2DocumentUpload = getElement("xpath", "//div[contains(text(),'Add L2 Document')]").Text;
                string AddObligation = getElement("xpath", "//div[contains(text(),'Add Obligation')]").Text;
                // All Add with Excel file Read
                string[] UploadExcelColumnNames = { testData["UploadHader_AddL1Document"], testData["UploadHader_AddL2Document"],
                                                   testData["UploadHader_Obligation"]};
                int sizeoftestdata = UploadExcelColumnNames.Count();
                for (int x = 1; x <= sizeoftestdata; x++)
                {
                    int UplaodNamesgrid = x - 1;
                    if (AddL1DocumentUpload.Equals(UploadExcelColumnNames[UplaodNamesgrid]))
                    {
                        string Test = testData["Upload_AddL1Document"];
                        UpdateUploadExcelfile(testData,AddL1DocumentUpload,Test);
                        ClickOnElementWhenElementFound("xpath", "//*[contains(@headertext,'Add L1 Document')]//*[contains(@class,'k-button k-upload-button')]", "Add L1 Document");
                        Fileupload(Test);
                        ClickOnElementWhenElementFound("xpath", "//button[contains(@class,'k-button k-primary k-upload-selected')]", "Upload Button");
                        wait.Until(waitforElement_Delete());
                        wait.Until(WaitforLabel());
                        IWebElement AddwithExcelTabforAddL1Document = wait.Until(waitforElement_AddwithExcel());
                        AddwithExcelTabforAddL1Document.Click();
                    }
                    else if (AddL2DocumentUpload.Equals(UploadExcelColumnNames[UplaodNamesgrid]))
                    {
                        string Test = testData["Upload_AddL2Document"];
                        UpdateUploadExcelfile(testData,AddL2DocumentUpload, Test);
                        ClickOnElementWhenElementFound("xpath", "//*[contains(@headertext,'Add L2 Document')]//*[contains(@class,'k-button k-upload-button')]", "Add L2 Document");
                        Fileupload(Test);
                        ClickOnElementWhenElementFound("xpath", "//button[contains(@class,'k-button k-primary k-upload-selected')]", "Upload Button");
                        wait.Until(waitforElement_Delete());
                        wait.Until(WaitforLabel());
                        IWebElement AddwithExcelTabforAddL2Document = wait.Until(waitforElement_AddwithExcel());
                        AddwithExcelTabforAddL2Document.Click();
                    }
                    else if (AddObligation.Equals(UploadExcelColumnNames[UplaodNamesgrid]))
                    {
                        string Test = testData["Upload_Obligation"];
                        UpdateExcelfilewithMultipledata(testData, AddObligation, Test);
                        ClickOnElementWhenElementFound("xpath", "//*[contains(@headertext,'Add Obligation')]//*[contains(@class,'k-button k-upload-button')]", "Add obligation");
                        Fileupload(Test);
                        ClickOnElementWhenElementFound("xpath", "//button[contains(@class,'k-button k-primary k-upload-selected')]", "Upload Button");
                        //ClickOnElementWhenElementFound("xpath", "//span[contains(text(),'Key Information')]", "Key Information");
                    }
                    else
                    {
                            
                    }
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
        /// Desc: Method for navigate to Contract Short Name via Tree View
        /// validation >> Contract Short Name in tree view >> Click the tree
        /// </summary>
        /// <param name="testData"></param>
        public void NavigateBacktoCSName(Dictionary<string, string> testData)
        {
            try 
            {
                PageLoadWait(10);
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(50));
                wait.PollingInterval = TimeSpan.FromMilliseconds(250);
                wait.IgnoreExceptionTypes(typeof(NoSuchElementException), typeof(ElementNotVisibleException));
                wait.Until(waitforElement_AuditLogs());
                wait.Until(waitforElement_treeview());
                AssertAreEqual("xpath", "//span[contains(text(),'" + testData["TextBox_ContractShortName"] + "')]",testData["TextBox_ContractShortName"]);
                AssertAreEqual("xpath", "//span[contains(@class,'k-in k-state-selected')]//span[contains(@class,'ng-star-inserted')][contains(text(),'" + testData["Textbox_ModifyDocumentNameL1"] + "')]", testData["Textbox_ModifyDocumentNameL1"]);
                ClickOnElementWhenElementFound("xpath", "//span[contains(text(),'" + testData["TextBox_ContractShortName"] + "')]", "Tree view");
            }
            catch (Exception ex)
            {
                GeneralMethod.ScreenShotCapture();
                ExtentTestManager._parentTest.Log(Status.Fail, "Expected not matched " + ex.Message);
                Assert.Fail(ex.Message);
            }
        }


        /// <summary>
        /// Desc: Method created for High Risk Items
        /// </summary>
        /// <param name="testData"></param>
        public void HighRiskItemsTab(Dictionary<string, string> testData)
        {

            try
            {
                WaitforElement(50, 250, "//span[contains(text(),'Audit Logs')]");
                WaitforElement(50, 250, "//button[contains(text(),'Delete')]");
                WaitforElement(50, 250, "//div[contains(text(),'RAG Status')]");
                WaitforElement(50, 250, "//span[contains(text(),'High Risk Items')]");
                ClickOnElementWhenElementFound("xpath", "//span[contains(text(),'High Risk Items')]", "High Risk Items Tab");
                WaitforElementbool(50, 250, "//button[contains(@class,'k-button-icon k-button k-grid-pdf')]");
                AssertIsTrue("xpath", "//button[contains(@class,'k-button-icon k-button k-grid-excel')]", "Excel Element Icon is visible");
                threadWait(1000);
                ClickOnElementWhenElementFound("xpath", "//span[contains(text(),'Key Information')]", "Key Information");

                
            }
            catch (Exception ex)
            {
                GeneralMethod.ScreenShotCapture();
                ExtentTestManager._parentTest.Log(Status.Fail, "Expected not matched " + ex.Message);
                Assert.Fail(ex.Message);
            }
        }


        /// <summary>
        /// Desc: Method created for Active the Contract 
        /// Phase of Active >> Alert CCM >> 
        /// </summary>
        /// <param name="testData"></param>
        public void ActiveContract(Dictionary<string, string> testData)
        {
            try 
            {
                PageLoadWait(10);
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(50));
                wait.PollingInterval = TimeSpan.FromMilliseconds(250);
                wait.IgnoreExceptionTypes(typeof(NoSuchElementException), typeof(ElementNotVisibleException));
                wait.Until(waitforElement_AuditLogs());
                wait.Until(waitforElement_Delete());
                wait.Until(waitforRagStatus());
                SendKeysForElement("xpath", "//kendo-combobox[contains(@name,'requestedBy')]//input[contains(@class,'k-input')]", testData["Dropdown_LeadCMM"], "Requested By");
                getElement("xpath", "//kendo-combobox[contains(@name,'requestedBy')]//input[contains(@class,'k-input')]").SendKeys(Keys.Enter); //Requested By 
                getElement("xpath", "//button[contains(text(),'Alert CCM')]").SendKeys(Keys.Enter); //Automation clicked Alert CCM
                wait.Until(waitforPreactivationMessgae());
                AssertAreEqual("xpath", "//div[@class='k-content k-window-content k-dialog-content']", testData["PreActiveMsg_AlertCCM"]);
                IWebElement WaitClick = wait.Until(waitforPreactivationMessgae_Okay());
                WaitClick.Click();
                wait.Until(waitforRagStatus());
                SendKeysForElement("xpath", "//kendo-combobox[contains(@name,'requestedBy')]//input[contains(@class,'k-input')]", testData["Dropdown_LeadCMM"], "Requested By");
                getElement("xpath", "//kendo-combobox[contains(@name,'requestedBy')]//input[contains(@class,'k-input')]").SendKeys(Keys.Enter); //Requested By 
                getElement("xpath", "//button[contains(text(),'Pre-Activate Contract')]").SendKeys(Keys.Enter); //Pre-Activation Contract
                wait.Until(waitforPreActivateContractMessgae());
                AssertAreEqual("xpath", "//div[@class='k-content k-window-content k-dialog-content']", testData["PreActiveMsg_Pre-ActivateContract"]);
                wait.Until(waitforPreactivationMessgae_Okay()).Click();
                wait.Until(waitforRagStatus());
                wait.Until(WaitforRequestBydropdown());
                threadWait(2400);
                ActiveContract_Last(testData);

            }
            catch (Exception ex)
            {
                GeneralMethod.ScreenShotCapture();
                ExtentTestManager._parentTest.Log(Status.Fail, "Expected not matched " + ex.Message);
                Assert.Fail(ex.Message);
            }
        }


        public void ActiveContract_Last(Dictionary<string, string> testData)
        {

            try
            {
                PageLoadWait(10);
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(50));
                wait.PollingInterval = TimeSpan.FromMilliseconds(250);
                wait.IgnoreExceptionTypes(typeof(NoSuchElementException), typeof(ElementNotVisibleException));
                wait.Until(WaitforActiveContract());
                wait.Until(WaitforRequestLable());
                SendKeysForElement("xpath", "//kendo-combobox[contains(@name,'requestedBy')]//input[contains(@class,'k-input')]", testData["Dropdown_LeadCMM"], "Requested By");
                getElement("xpath", "//kendo-combobox[contains(@name,'requestedBy')]//input[contains(@class,'k-input')]").SendKeys(Keys.Enter); //Requested By 
                getElement("xpath", "//button[contains(text(),'Activate Contract')]").SendKeys(Keys.Enter); //Activate Contract
                wait.Until(waitforActivateContractMessgae());
                AssertAreEqual("xpath", "//div[contains(@class,'k-content k-window-content k-dialog-content')]", testData["ContractActivateMsg_Activate"]);
                wait.Until(waitforPreactivationMessgae_Okay()).Click();
                wait.Until(WaitforContractActive_LastMsg());
                AssertAreEqual("xpath", "//span[contains(text(),'Since the contract is in process hence, no action')]", testData["ContractActive_waitMsg"]);

            }
            catch (Exception ex)
            {
                GeneralMethod.ScreenShotCapture();
                ExtentTestManager._parentTest.Log(Status.Fail, "Expected not matched " + ex.Message);
                Assert.Fail(ex.Message);
            }
        }

        /// <summary>
        /// Desc: Method created for Audit Logs
        /// </summary>
        /// <param name="testData"></param>
        public void AuditLogs(Dictionary<string, string> testData)
        {
            try
            {
                WaitforElement(30,250, "//span[contains(text(),'Audit Logs')]");
                AssertIsTrue("xpath", "//span[contains(text(),'Audit Logs')]", "Audit Logs Tab");
                ClickOnElementWhenElementFound("xpath", "//span[contains(text(),'Audit Logs')]", "Audit Logs Tab");
                PdfandExcelIcon(testData); 
                row = getElements("xpath", "//*[@class='k-grid-header-wrap']//descendant::tr/th//a[contains(@class,'k-link ng-star-inserted')]");
                int size = row.Count;
                for (int i = 1; i <= size; i++)
                {
                    string[] HaderNames = { "Event", "Requested By", "Completed By", "Audit Date", "Additional Info"};
                    string HeaderName = HaderNames[i - 1];
                    AssertAreEqual("xpath", "//*[@class='k-grid-header-wrap']//descendant::tr/th[" + i + "]//a[contains(@class,'k-link ng-star-inserted')]", HeaderName);
                   
                }
                AuditLogs_gridvlidate(testData);
            }
            catch (Exception ex)
            {
                GeneralMethod.ScreenShotCapture();
                ExtentTestManager._parentTest.Log(Status.Fail, "Expected not matched " + ex.Message);
                Assert.Fail(ex.Message);
            }
        }

        /// <summary>
        /// Desc: Method created for Audit Logs grid
        /// </summary>
        /// <param name="testData"></param>
        public void AuditLogs_gridvlidate(Dictionary<string, string> testData)
        {
            try
            {
                WaitforElement(30, 250, "//span[contains(text(),'Audit Logs')]");
                AssertIsTrue("xpath", "//span[contains(text(),'Audit Logs')]", "Audit Logs Tab");
                ClickOnElementWhenElementFound("xpath", "//span[contains(text(),'Audit Logs')]", "Audit Logs Tab");
                PdfandExcelIcon(testData);
                row = getElements("xpath", "//*[@class='k-grid-container ng-star-inserted']//descendant::tr/td//span");
                int size = row.Count;
                for (int j = 1; j <= size; j++)
                {
                    string[] GridNames = { "Contract has been activated", "Contract has been pre-activated", "Pre-activation request sent to CCM", "Contract has been added" };
                    string GridName = GridNames[j - 1];
                    AssertAreEqual("xpath", "//*[@class='k-grid-container ng-star-inserted']//descendant::tr["+j+"]/td//span", GridName);
                    AssertAreEqual("xpath", "//*[@class='k-grid-container ng-star-inserted']//descendant::tr["+j+"]/td[2]", testData["Dropdown_ObligationOwner"]);
                    AssertAreEqual("xpath", "//*[@class='k-grid-container ng-star-inserted']//descendant::tr["+j+"]/td[3]", testData["Dropdown_ObligationOwner"]);
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
        /// Desc: Method for Generate Contract Dump
        /// </summary>
        /// <param name="testData"></param>
        public void ContractDump(Dictionary<string, string> testData)
        {
            try
            {
                WaitforElement(50, 250, "//span[contains(text(),'Audit Logs')]");
                WaitforElement(50, 250, "//button[contains(text(),'Delete')]");
                WaitforElement(50, 250, "//div[contains(text(),'RAG Status')]");
                WaitforElementbool(20,250, "//button[contains(text(),'Generate Contract Dump')]");
                ClickOnElementWhenElementFound("xpath", "//button[contains(text(),'Generate Contract Dump')]", "Generate Contract Dump"); //Generate Contract Dump
                WaitforElement(10,250, "//*[@class='ulx-notification-title']");
            }
            catch (Exception ex)
            {
                GeneralMethod.ScreenShotCapture();
                ExtentTestManager._parentTest.Log(Status.Fail, "Expected not matched " + ex.Message);
                Assert.Fail(ex.Message);
            }
        }


        /// <summary>
        /// Desc: Method for Delete  Newly Created contract
        /// </summary>
        /// <param name="testData"></param>
        public void DeleteContract(Dictionary<string, string> testData)
        {
            try 
            {
                PageLoadWait(10);
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(50));
                wait.PollingInterval = TimeSpan.FromMilliseconds(250);
                wait.IgnoreExceptionTypes(typeof(NoSuchElementException), typeof(ElementNotVisibleException));
                IWebElement Delete = wait.Until(waitforElement_Delete());
                SendKeysForElement("xpath", "//*[contains(@name,'requestedBy')]//*[@class='k-input']", testData["Dropdown_LeadCMM"], "Requested By");
                ClickOnElementWhenElementFound("xpath", "//*[contains(@name,'requestedBy')]//*[@class='k-input']", "Requested By"); //Requested By
                Delete.Click();
                ExtentTestManager._parentTest.Log(Status.Pass, "Delete Button hit ");//Delete Button hit
                ElementWaitTime(10);
                AssertAreEqual("xpath", "//div[contains(text(),'Do you wish to delete this item?')]" ,testData["DeleteContract_Popupwindow"]);
                ClickOnElementWhenElementFound("xpath", "//button[contains(@class,'k-button k-primary ng-star-inserted')]", "Delete confrim");
                IWebElement YesConformation = wait.Until(waitforElement_AfterDeleteMessage());
                YesConformation.Click();
                ExtentTestManager._parentTest.Log(Status.Pass, "Delete conformation ");//Delete Button hit
            }
            catch (Exception ex) 
            {
                GeneralMethod.ScreenShotCapture();
                ExtentTestManager._parentTest.Log(Status.Fail, "Expected not matched " + ex.Message);
                Assert.Fail(ex.Message);
            }
        }
        
        //verfiy the Scroll
        public void Scrollanywheres(Dictionary<string, string> testData)
        {
            IJavaScriptExecutor jse = (IJavaScriptExecutor)driver;
            jse.ExecuteScript("scroll(0,910)");
        }

        public void RequestedBy(Dictionary<string, string> testData)
        {
            try 
            { 
                SendKeysForElement("xpath", "//kendo-combobox[contains(@id,'users')]//input[contains(@class,'k-input')]", testData["Dropdown_LeadCMM"], "Requested By");
                getElement("xpath", "//kendo-combobox[contains(@id,'users')]//input[contains(@class,'k-input')]").SendKeys(Keys.Enter); //Requested By 
                getElement("xpath", "//button[contains(text(),'Submit')]").SendKeys(Keys.Enter); //Submit Button
            }
            catch (Exception ex)
            {
                GeneralMethod.ScreenShotCapture();
                ExtentTestManager._parentTest.Log(Status.Fail, "Expected not matched " + ex.Message);
                Assert.Fail(ex.Message);

            }
        }



        /// <summary>
        /// Desc: Method for Validate Audit Log tab 
        /// </summary>
        /// <param name="testData"></param>
        public void AuditLog(Dictionary<string, string> testData)
        {
            try
            {
                PageLoadWait(10);
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(50));
                wait.PollingInterval = TimeSpan.FromMilliseconds(250);
                wait.IgnoreExceptionTypes(typeof(NoSuchElementException), typeof(ElementNotVisibleException));
                IWebElement Delete = wait.Until(waitforElement_Delete());
                ClickOnElementWhenElementFound("xpath", "//*[@class='k-link'][contains(text(),'Audit Logs')]", "Audit Log Tab");
                PdfandExcelIcon(testData);
                AssertIsTrue("xpath", "//*[@class='ulx-panel-header ng-star-inserted']", "Audit Log Header");
                row = getElements("xpath", "//*[@class='k-grid-header-wrap']//descendant::tr/th");
                for (int i = 1; i <= row.Count; i++)
                {
                    string[] HaderNames = { "Event", "Requested By", "Completed By", "Audit Date", "Modifications Made" };
                    string HeaderName = HaderNames[i - 1];
                    AssertAreEqual("xpath", "//*[@class='k-grid-header-wrap']//descendant::tr/th[" + i + "]", HeaderName);
                }
                row = getElements("xpath", "//*[@class='k-grid-table-wrap']");
                for (int j = 1; j <= row.Count; j++)
                {
                    string Header = getElement("xpath", "//*[@class='k-grid-table-wrap']//descendant::tr/td[" + j + "]").Text;
                    string customName = testData["Event_Column"];
                    if (Header.Equals(testData["Event_Column"]))
                    {
                        ExtentTestManager._parentTest.Log(Status.Pass, "Expected  matched ");
                    }
                    else
                    {
                        ExtentTestManager._parentTest.Log(Status.Fail, "Expected not matched ");
                        Console.WriteLine("Actual value: " + Header + ", expected value: " + customName);
                        Console.WriteLine("Test Case Not Passed");
                    }
                }

                //AssertIsTrue("xpath", "//a[@class='k-link ng-star-inserted'][contains(text(),'Event')]", "Event");
                //AssertIsTrue("xpath", "//th[1]//kendo-grid-filter-menu[1]//a[1]//span[1]", "FilterIcon[1]");
                //AssertIsTrue("xpath", "//a[@class='k-link ng-star-inserted'][contains(text(),'Requested By')]", "Requested By");
                //AssertIsTrue("xpath", "//th[2]//kendo-grid-filter-menu[1]//a[1]", "FilterIcon[2]");
                //AssertIsTrue("xpath", "//a[@class='k-link ng-star-inserted   Hover']", "Completed By");
                //AssertIsTrue("xpath", "//th[3]//kendo-grid-filter-menu[1]//a[1]//span[1]", "FilterIcon[3]");
                //AssertIsTrue("xpath", "//a[@class='k-link ng-star-inserted'][contains(text(),'Audit Date')]", "Audit Date");
                //AssertIsTrue("xpath", "//th[4]//kendo-grid-filter-menu[1]//a[1]//span[1]", "FilterIcon[4]");
                //AssertIsTrue("xpath", "//a[@class='k-link ng-star-inserted'][contains(text(),'Additional Info')]", "Additional Info");
                //AssertIsTrue("xpath", "//th[5]//kendo-grid-filter-menu[1]//a[1]//span[1]", "FilterIcon[5]");
                //AssertIsTrue("xpath", "//*[@class='k-button-icon k-button k-grid-pdf ng-star-inserted']/span", "Pdf Icon");
                //AssertIsTrue("xpath", "//*[@class='k-button-icon k-button k-grid-excel ng-star-inserted']/span", "Excel Icon");
                ////AssertIsTrue("xpath", "//a[@class='k-link k-state-selected']", "Pagination");
                ////AssertIsTrue("xpath", "//*[@class='k-pager-info k-label ng-star-inserted']", "Pagination Items Grid");
                AssertIsTrue("xpath", "//*[@class='row w-100 p-2 ng-star-inserted']", "Requested By");
                SendKeysForElement("xpath", "//span[@class='k-i-arrow-s k-icon']", testData["Dropdown_LeadCMM"], "Select User");
                AssertIsTrue("xpath", "//*[@class='text-center mt-2 ng-star-inserted']//descendant::button[1]", "Modify Contract");
                AssertIsTrue("xpath", "//*[@class='text-center mt-2 ng-star-inserted']//descendant::button[2]", "Add L1 Document");
                AssertIsTrue("xpath", "//*[@class='text-center mt-2 ng-star-inserted']//descendant::button[3]", "Add L2 Document");
                AssertIsTrue("xpath", "//*[@class='text-center mt-2 ng-star-inserted']//descendant::button[4]", "Modify Contract Tree");
                AssertIsTrue("xpath", "//*[@class='text-center mt-2 ng-star-inserted']//descendant::button[5]", "Alert CCM");
                AssertIsTrue("xpath", "//*[@class='text-center mt-2 ng-star-inserted']//descendant::button[6]", "Archive Contract");
                AssertIsTrue("xpath", "//*[@class='text-center mt-2 ng-star-inserted']//descendant::button[7]", "Delete");
                AssertIsTrue("xpath", "//*[@class='text-center mt-2 ng-star-inserted']//descendant::button[8]", "Generate Contract Dump");
                AssertIsTrue("xpath", "//*[@class='text-center mt-2 ng-star-inserted']//descendant::button[9]", "Obligation Calendar");
            }
            catch (Exception ex)
            {
                GeneralMethod.ScreenShotCapture();
                ExtentTestManager._parentTest.Log(Status.Fail, "Expected not matched " + ex.Message);
                Assert.Fail(ex.Message);
            }

            }

        /// <summary>
        /// Desc: Method for Validate File tab 
        /// </summary>
        /// <param name="testData"></param>
        public void FileTab(Dictionary<string, string> testData)
        {
            try
            {
                PageLoadWait(10);
                string Test = testData["Upload_SelectFiles"];
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(50));
                wait.PollingInterval = TimeSpan.FromMilliseconds(250);
                wait.IgnoreExceptionTypes(typeof(NoSuchElementException), typeof(ElementNotVisibleException));
                IWebElement Delete = wait.Until(waitforElement_Delete());
                ClickOnElementWhenElementFound("xpath", "//*[@class='text-center mt-2 ng-star-inserted']//descendant::button[2]", "Add L1 Document");
                ClickOnElementWhenElementFound("xpath", "//*[@id='k-tabstrip-tab-2']", "Files");
                AssertIsTrue("xpath", "//*[@class='ulx-panel-header ng-star-inserted']", "Upload Files Tab Header");
                ClickOnElementWhenElementFound("xpath", "//*[@class='k-button k-upload-button']", "Select Files Button");
                Fileupload(Test);
                wait.Until(waitforElement_Delete());
                wait.Until(WaitforLabel());
                ClickOnElementWhenElementFound("xpath", "//button[@class='k-button k-primary k-upload-selected']", "Upload Button");
                SendKeysForElement("xpath", "//*[@class='row w-100 p-2 ng-star-inserted']", testData["Dropdown_LeadCMM"], "Requested By");
                getElement("xpath", "//kendo-combobox[contains(@id,'users')]//input[contains(@class,'k-input')]").SendKeys(Keys.Enter); //Requested By 
                getElement("xpath", "//button[contains(text(),'Submit')]").SendKeys(Keys.Enter); //Submit Button

            }
            catch (Exception ex)
            {
                GeneralMethod.ScreenShotCapture();
                ExtentTestManager._parentTest.Log(Status.Fail, "Expected not matched " + ex.Message);
                Assert.Fail(ex.Message);
            }

        }

        /// <summary>
        /// Desc: Method for Validate Other Information tab 
        /// </summary>
        /// <param name="testData"></param>
        public void OtherInformationTab(Dictionary<string, string> testData)
        {
            try
            {
                PageLoadWait(10);
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(50));
                wait.PollingInterval = TimeSpan.FromMilliseconds(250);
                wait.IgnoreExceptionTypes(typeof(NoSuchElementException), typeof(ElementNotVisibleException));
                IWebElement Delete = wait.Until(waitforElement_Delete());
                PdfandExcelIcon(testData);
                ClickOnElementWhenElementFound("xpath", "//span[contains(text(),'Other Information')]", "Other Information Tab");
                ClickOnElementWhenElementFound("xpath", "//button[contains(text(),'Expand All')]", "Expand All");
                AssertIsTrue("xpath", "//span[@class='ng-tns-c10-43 k-link k-header k-state-selected']", "Governing Law");
            }
            catch (Exception ex)
            {
                GeneralMethod.ScreenShotCapture();
                ExtentTestManager._parentTest.Log(Status.Fail, "Expected not matched " + ex.Message);
                Assert.Fail(ex.Message);
            }

        }
























































        public void PdfandExcelIcon(Dictionary<string, string> testData)
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

        // Wait part 
        private Func<IWebDriver, IWebElement> waitforElement_Delete()
        {
            return ((x) =>
            {
                ExtentTestManager._parentTest.Log(Status.Pass, "Waiting for element");
                if (x.FindElements(By.XPath("//button[contains(text(),'Delete')]")).Count == 1)
                    return x.FindElement(By.XPath("//button[contains(text(),'Delete')]"));
                return null;
            });
        }
        private Func<IWebDriver, IWebElement> waitforElement_AfterDeleteMessage()
        {
            return ((x) =>
            {
                ExtentTestManager._parentTest.Log(Status.Pass, "Waiting for element");
                if (x.FindElements(By.XPath("//button[@class='k-button k-primary ng-star-inserted']")).Count == 1)
                    return x.FindElement(By.XPath("//button[@class='k-button k-primary ng-star-inserted']"));
                return null;
            });
        }


        private Func<IWebDriver, IWebElement> WaitforLabel()
        {
            return ((x) =>
            {

                ExtentTestManager._parentTest.Log(Status.Pass, "Waiting for element");
                if (x.FindElements(By.XPath("//div[contains(text(),'Agreement Number')]")).Count == 1)
                    return x.FindElement(By.XPath("//div[contains(text(),'Agreement Number')]"));
                return null;


            });
        }

        private Func<IWebDriver, IWebElement> waitforElement_treeview()
        {
            return ((x) =>
            {
                ExtentTestManager._parentTest.Log(Status.Pass, "Waiting for element");
                if (x.FindElements(By.XPath("//div[contains(text(),'Contract Short Name')]")).Count == 1)
                    return x.FindElement(By.XPath("//div[contains(text(),'Contract Short Name')]"));
                return null;
            });
        }
        private Func<IWebDriver, IWebElement> waitforElement_AddwithExcel()
        {
            return ((x) =>
            {
                ExtentTestManager._parentTest.Log(Status.Pass, "Waiting for element");
                if (x.FindElements(By.XPath("//span[contains(text(),'Add with Excel')]")).Count == 1)
                    return x.FindElement(By.XPath("//span[contains(text(),'Add with Excel')]"));
                return null;
            });
        }

        private Func<IWebDriver, IWebElement> waitforElement_AddwithExcelgridDownload()
        {
            return ((x) =>
            {
                ExtentTestManager._parentTest.Log(Status.Pass, "Waiting for element");
                if (x.FindElements(By.XPath("//ulx-panel[contains(@headertext,'Add L1 Document')]//button[contains(@class,'btn btn-dark')]")).Count == 1)
                    return x.FindElement(By.XPath("//ulx-panel[contains(@headertext,'Add L1 Document')]//button[contains(@class,'btn btn-dark')]"));
                return null;
            });
        }

        
        private Func<IWebDriver, IWebElement> waitforElement_AuditLogs()
        {
            return ((x) =>
            {
                ExtentTestManager._parentTest.Log(Status.Pass, "Waiting for element");
                if (x.FindElements(By.XPath("//span[contains(text(),'Audit Logs')]")).Count == 1)
                    return x.FindElement(By.XPath("//span[contains(text(),'Audit Logs')]"));
                return null;
            });
        }
        private Func<IWebDriver, IWebElement> waitforElement_KeyInfoDocName(Dictionary<string, string> testData)
        {
            return ((x) =>
            {
                ExtentTestManager._parentTest.Log(Status.Pass, "Waiting for element");
                if (x.FindElements(By.XPath("//div[contains(text(),'" + testData["Textbox_DocumentNameL1"] + "')]")).Count == 1)
                    return x.FindElement(By.XPath("//div[contains(text(),'" + testData["Textbox_DocumentNameL1"] + "')]"));
                return null;
            });
        }
        private Func<IWebDriver, IWebElement> waitforElement_DocumentTitleModifys()
        {
            return ((x) =>
            {
                ExtentTestManager._parentTest.Log(Status.Pass, "Waiting for element");
                if (x.FindElements(By.XPath("//div[contains(text(),'Document Title')]")).Count == 1)
                    return x.FindElement(By.XPath("//div[contains(text(),'Document Title')]"));
                return null;
            });
        }

        private Func<IWebDriver, IWebElement> waitforElement_ModifyDocumentName(Dictionary<string, string> testData)
        {
            return ((x) =>
            {
                ExtentTestManager._parentTest.Log(Status.Pass, "Waiting for element");
                if (x.FindElements(By.XPath("//div[contains(text(),'" + testData["Textbox_ModifyDocumentNameL1"] + "')]")).Count == 1)
                    return x.FindElement(By.XPath("//div[contains(text(),'" + testData["Textbox_ModifyDocumentNameL1"] + "')]"));
                return null;
            });
        }

        private Func<IWebDriver, IWebElement> waitforElement_Obligations()
        {
            return ((x) =>
            {
                ExtentTestManager._parentTest.Log(Status.Pass, "Waiting for element");
                if (x.FindElements(By.XPath("//span[contains(text(),'Obligations')]")).Count == 1)
                    return x.FindElement(By.XPath("//span[contains(text(),'Obligations')]"));
                return null;
            });
        }

        private Func<IWebDriver, IWebElement> waitforElement_ObligationsOwner()
        {
            return ((x) =>
            {
                ExtentTestManager._parentTest.Log(Status.Pass, "Waiting for element");
                if (x.FindElements(By.XPath("//th[contains(text(),'Owner')]")).Count == 1)
                    return x.FindElement(By.XPath("//th[contains(text(),'Owner')]"));
                return null;
            });
        }


        private Func<IWebDriver, IWebElement> waitforElement_Triggered()
        {
            return ((x) =>
            {
                ExtentTestManager._parentTest.Log(Status.Pass, "Waiting for element");
                if (x.FindElements(By.XPath("//th[contains(text(),'Owner')]")).Count == 1)
                    return x.FindElement(By.XPath("//th[contains(text(),'Owner')]"));
                return null;
            });
        }

        private Func<IWebDriver, IWebElement> waitforElement_Frequency()
        {
            return ((x) =>
            {
                ExtentTestManager._parentTest.Log(Status.Pass, "Waiting for element");
                if (x.FindElements(By.XPath("//div[contains(text(),'Frequency')]")).Count == 1)
                    return x.FindElement(By.XPath("//div[contains(text(),'Frequency')]"));
                return null;
            });
        }


        private Func<IWebDriver, IWebElement> waitforElement_ObligationOwner()
        {
            return ((x) =>
            {
                ExtentTestManager._parentTest.Log(Status.Pass, "Waiting for element");
                if (x.FindElements(By.XPath("//div[contains(text(),'Obligation Owner')]")).Count == 1)
                    return x.FindElement(By.XPath("//div[contains(text(),'Obligation Owner')]"));
                return null;
            });
        }

        private Func<IWebDriver, IWebElement> waitforRagStatus()
        {
            return ((x) =>
            {
                ExtentTestManager._parentTest.Log(Status.Pass, "Waiting for element");
                if (x.FindElements(By.XPath("//div[contains(text(),'RAG Status')]")).Count == 1)
                    return x.FindElement(By.XPath("//div[contains(text(),'RAG Status')]"));
                return null;
            });
        }


        private Func<IWebDriver, IWebElement> waitforPreactivationMessgae()
        {
            return ((x) =>
            {
                ExtentTestManager._parentTest.Log(Status.Pass, "Waiting for element");
                if (x.FindElements(By.XPath("//div[@class='k-content k-window-content k-dialog-content']")).Count == 1)
                    return x.FindElement(By.XPath("//div[@class='k-content k-window-content k-dialog-content']"));
                return null;
            });
        }

        private Func<IWebDriver, IWebElement> waitforPreactivationMessgae_Okay()
        {
            return ((x) =>
            {
                ExtentTestManager._parentTest.Log(Status.Pass, "Waiting for element");
                if (x.FindElements(By.XPath("//button[@class='k-button k-primary ng-star-inserted']")).Count == 1)
                    return x.FindElement(By.XPath("//button[@class='k-button k-primary ng-star-inserted']"));
                return null;
            });
        }


        private Func<IWebDriver, IWebElement> waitforPreActivateContractMessgae()
        {
            return ((x) =>
            {
                ExtentTestManager._parentTest.Log(Status.Pass, "Waiting for element");
                if (x.FindElements(By.XPath("//div[@class='k-content k-window-content k-dialog-content']")).Count == 1)
                    return x.FindElement(By.XPath("//div[@class='k-content k-window-content k-dialog-content']"));
                return null;
            });
        }

        private Func<IWebDriver, IWebElement> waitforActivateContractMessgae()
        {
            return ((x) =>
            {
                ExtentTestManager._parentTest.Log(Status.Pass, "Waiting for element");
                if (x.FindElements(By.XPath("//div[@class='k-content k-window-content k-dialog-content']")).Count == 1)
                    return x.FindElement(By.XPath("//div[@class='k-content k-window-content k-dialog-content']"));
                return null;
            });
        }

        private Func<IWebDriver, IWebElement> WaitforContractActive_LastMsg()
        {
            return ((x) =>
            {
                ExtentTestManager._parentTest.Log(Status.Pass, "Waiting for element");
                if (x.FindElements(By.XPath("//span[contains(text(),'Since the contract is in process hence, no action')]")).Count == 1)
                    return x.FindElement(By.XPath("//span[contains(text(),'Since the contract is in process hence, no action')]"));
                return null;
            });
        }

        private Func<IWebDriver, IWebElement> WaitforRequestBydropdown()
        {
            return ((x) =>
            {
                ExtentTestManager._parentTest.Log(Status.Pass, "Waiting for element");
                if (x.FindElements(By.XPath("//kendo-combobox[contains(@name,'requestedBy')]//input[contains(@class,'k-input')]")).Count == 1)
                    return x.FindElement(By.XPath("//kendo-combobox[contains(@name,'requestedBy')]//input[contains(@class,'k-input')]"));
                return null;
            });
        }
        private Func<IWebDriver, IWebElement> WaitforActiveContract()
        {
            return ((x) =>
            {
                ExtentTestManager._parentTest.Log(Status.Pass, "Waiting for element");
                if (x.FindElements(By.XPath("//button[contains(text(),'Activate Contract')]")).Count == 1)
                    return x.FindElement(By.XPath("//button[contains(text(),'Activate Contract')]"));
                return null;
            });
        }
        private Func<IWebDriver, IWebElement> WaitforRequestLable()
        {
            return ((x) =>
            {
                ExtentTestManager._parentTest.Log(Status.Pass, "Waiting for element");
                if (x.FindElements(By.XPath("//div[contains(text(),'Requested By')]")).Count == 1)
                    return x.FindElement(By.XPath("//div[contains(text(),'Requested By')]"));
                return null;
            });
        }
    }

}