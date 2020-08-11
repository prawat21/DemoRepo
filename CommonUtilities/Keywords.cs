using AventStack.ExtentReports;

using NUnit.Framework;
using LexBaseFramework.LexBaseLibrary;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LexBaseFramework.NovsAuthentication;

namespace LexBaseFramework.CommonUtilities
{
    public class Keywords : GeneralMethod
    {
     

        /// <summary>
        /// Desc:Method is used to read excel sheets and execute the keywords accordingly
        /// </summary>
        /// <param name="testUnderExecution"></param>
        /// <param name="xls"></param>
        /// <param name="testData"></param>
        public void executeKeywords(string testUnderExecution, ExcelReader xls, Dictionary<string, string> testData)
        {

            //  LexBaseLibrary Object.

            ApplicationLogin_FunctionLibrary Appobj = new ApplicationLogin_FunctionLibrary();
            HomePage_FunctionLibrary homePageobj = new HomePage_FunctionLibrary();
            Contracts_FunctionLibrary contractsobj = new Contracts_FunctionLibrary();
            Obligation_FunctionLibrary Obligationobj = new Obligation_FunctionLibrary();
            StandardReports_FunctionLibrary  StandardReportsObj = new StandardReports_FunctionLibrary();
            Custom_FunctionLibrary CustomReportsObj = new Custom_FunctionLibrary();
            ReportTemplate_FunctionLibrary ReprtsTemplateObj = new ReportTemplate_FunctionLibrary();
            ChangeRequest_FunctionLibrary ChangeRequetsObj = new ChangeRequest_FunctionLibrary();
            ObligationSearch_FunctionLibrary ObligationSearchObj = new ObligationSearch_FunctionLibrary();
            UserSearch_FunctionLibrary USerSearchObj = new UserSearch_FunctionLibrary();
            Archive_FunctionLibrary ArchiveObj = new Archive_FunctionLibrary();
            BulkChange_FunctionLibrary bulkChangeObj = new BulkChange_FunctionLibrary();
            
            string KeywordsSheet = Enum.Sheets.EnvShakeout.ToString();
                int rows = xls.getRowCount(KeywordsSheet);

                for (int rNum = 2; rNum <= rows; rNum++)
                {
                    string tcid = xls.getCellData(KeywordsSheet, Enum.KeywordsColumn.TC_Id.ToString(), rNum);
                    if (tcid.Equals(testUnderExecution))
                    {
                        string data = null;
                        string keyword = xls.getCellData(KeywordsSheet, Enum.KeywordsColumn.Keyword.ToString(), rNum);
                        string locatorName = xls.getCellData(KeywordsSheet, Enum.KeywordsColumn.LocatorName.ToString(), rNum);
                        string key = xls.getCellData(KeywordsSheet, Enum.KeywordsColumn.Data.ToString(), rNum);
                        string runMode = xls.getCellData(KeywordsSheet, Enum.KeywordsColumn.RunMode.ToString(), rNum);
                        Dictionary<string, string> locatorData = DataUtility.locatorData(xls, locatorName);
                        if (!key.Equals(""))
                        {
                            data = testData[key];
                        }
                        Enum.LogStatus resultStatus = Enum.LogStatus.Passed;
                        switch (keyword)
                        {
                            case "openbrowser":
                                resultStatus = openBrowser(data);
                                break;
                            case "navigateprod":
                                resultStatus = navigateprod();
                                break;
                            case "navigateTest":
                                resultStatus = navigateTest();
                                break;
                            case "navigateStage":
                                resultStatus = navigateStage();
                                break;
                            case "click":
                                resultStatus = ClickOnElementWhenElementFound(locatorData.Keys.FirstOrDefault(), locatorData.Values.FirstOrDefault(), locatorName);
                                break;
                            case "input":
                                resultStatus = SendKeysForElement(locatorData.Keys.FirstOrDefault(), locatorData.Values.FirstOrDefault(), data, locatorName);
                                break;
                            case "selectValueFromDropdownByText":
                                resultStatus = SendKeysForElement(locatorData.Keys.FirstOrDefault(), locatorData.Values.FirstOrDefault(), data, locatorName);
                                break;
                            case "IsElementPresent":
                                resultStatus = SendKeysForAElement(locatorData.Keys.FirstOrDefault(), locatorData.Values.FirstOrDefault(), data, locatorName);
                                break;
                            case "wait":
                                resultStatus = threadWait(7500);
                                break;
                            case "waits":
                                resultStatus = threadWait(20000);
                                break;
                            case "MaxWait":
                                resultStatus = threadWait(80000);
                                break;
                            case "mimwait":
                                resultStatus = threadWait(3000);
                                break;
                            case "hover":
                                resultStatus = mouseHoverWithoutClick(locatorData.Keys.FirstOrDefault(), locatorData.Values.FirstOrDefault(), locatorName);
                                break;
                            case "hoverClick":
                                resultStatus = mouseHover(locatorData.Keys.FirstOrDefault(), locatorData.Values.FirstOrDefault());
                                break;
                            case "jsClick":
                                resultStatus = jsClick(locatorData.Keys.FirstOrDefault(), locatorData.Values.FirstOrDefault(), locatorName);
                                break;
                            case "inputAndEnter":
                                resultStatus = SendKeysForAElement(locatorData.Keys.FirstOrDefault(), locatorData.Values.FirstOrDefault(), data, locatorName);
                                break;
                            case "TDupdation_CCMUserLogin":
                                Appobj.TDupdation_CCMUserLogin(testData);
                                break;
                            case "TDupdation_AzureADLogin":
                                Appobj.TDupdation_AzureADLogin(testData);
                                break;
                            case "AzureADLogin":
                                Appobj.AzureADLogin(testData);
                                break;
                            case "CCM_UserLogin":
                                Appobj.CCM_UserLogin(testData);
                                break;
                            case "Credentail_Verfiy":
                                Appobj.Credentail_Verfiy(testData);
                                break;
                            case "HomePage_CCSDashboard":
                                homePageobj.HomePage_CCSDashboard(testData);
                                break;
                            case "NavigationContracts":
                                homePageobj.NavigationContracts(testData);
                                break;
                            case "PdfIconinGrid":
                                homePageobj.PdfIconinGrid(testData);
                                break;
                            case "Sorting":
                                homePageobj.Sorting(testData);
                                break;
                            case "HomePage_FilterIcon":
                                    homePageobj.HomePage_FilterIcon(testData);
                                break;
                            case "NewAddContratctTab":
                                    contractsobj.NewAddContratctTab(testData);
                                break;
                            case "ContractsPage":
                                contractsobj.ContractsPage(testData);
                                break;
                            case "Scrollanywheres":
                                contractsobj.Scrollanywheres(testData);
                                break;
                            case "Contracts_GridHader":
                                contractsobj.Contracts_GridHader(testData);
                                break;
                            case "NewlyCreatedGrid":
                                contractsobj.NewlyCreatedGrid(testData);
                                break;
                            case "AddL1Document":
                                contractsobj.AddL1Document(testData);
                                break;
                            case "ValdationContractTree":
                                contractsobj.ValdationContractTree(testData);
                                break;
                            case "AddObligation":
                                contractsobj.AddObligation(testData);
                                break;
                            case "ValidateObligations":
                                contractsobj.ValidateObligations(testData);
                                break;
                            case "AddwithExcel":
                                contractsobj.AddwithExcel(testData);
                                break;
                            case "NavigateBacktoCSName":
                                contractsobj.NavigateBacktoCSName(testData);
                                break;
                            case "ActiveContract":
                                contractsobj.ActiveContract(testData);
                                break;
                            case "ActiveContract_Last":
                                contractsobj.ActiveContract_Last(testData);
                                break;
                            case "DeleteContract":
                                contractsobj.DeleteContract(testData);
                                break;
                            case "RequestedBy":
                                contractsobj.RequestedBy(testData);
                                break;
                            case "HighRiskItemsTab":
                                contractsobj.HighRiskItemsTab(testData);
                                break;
                            case "PdfandExcelIcon":
                                contractsobj.PdfandExcelIcon(testData);
                                break;
                            case "AuditLogs":
                                contractsobj.AuditLogs(testData);
                                break;
                            case "AuditLogs_gridvlidate":
                                contractsobj.AuditLogs_gridvlidate(testData);
                                break;
                            case "ContractDump":
                                contractsobj.ContractDump(testData);
                                break;
                            case "AuditLog":
                            contractsobj.AuditLog(testData);
                                break;
                        case "FileTab":
                            contractsobj.FileTab(testData);
                            break;
                        case "ObligationPage":
                                Obligationobj.ObligationPage(testData);
                                break;
                            case "StandardReportPage":
                                StandardReportsObj.StandardReportPage(testData);
                                break;
                            case "StandardReportGrid":
                                StandardReportsObj.StandardReportGrid(testData);
                                break;
                            case "CustomReportPage":
                                CustomReportsObj.CustomReportPage(testData);
                                break;
                            case "CustomReportGrid":
                                CustomReportsObj.CustomReportGrid(testData);
                                break;
                            case "ReportPage":
                                ReprtsTemplateObj.ReportPage(testData);
                                break;
                            case "ChangeRequestPage":
                                ChangeRequetsObj.ChangeRequestPage(testData);
                                break;
                            case "HaderforChangeRequest":
                                ChangeRequetsObj.HaderforChangeRequest(testData);
                                break;
                            case "OtherWindowChangeRequest":
                                ChangeRequetsObj.OtherWindowChangeRequest(testData);
                                break;
                            case "PdfandExcelIconforChangeRequest":
                                ChangeRequetsObj.PdfandExcelIconforChangeRequest(testData);
                                break;
                            case "AddChangeRequest":
                                ChangeRequetsObj.AddChangeRequest(testData);
                                break;
                            case "SearchPage_Obligation":
                                ObligationSearchObj.SearchPage_Obligation(testData);
                                break;
                            case "SearchPage_ObligationFilters":
                                ObligationSearchObj.SearchPage_ObligationFilters(testData);
                                break;
                            case "SearchPage_User":
                                USerSearchObj.SearchPage_User(testData);
                                break;
                            case "SearchwithName":
                                USerSearchObj.SearchwithName(testData);
                                break;
                            case "ArchivePage":
                                ArchiveObj.ArchivePage(testData);
                                break;
                            case "PDFandExcelIcon_Archive":
                                ArchiveObj.PDFandExcelIcon_Archive(testData);
                                break;
                            case "BulkChangePage":
                                bulkChangeObj.BulkChangePage(testData);
                                break;
                             case "BulkchangeGrid":
                                bulkChangeObj.BulkchangeGrid(testData);
                                break;
                            case "UploadandAddtional":
                                bulkChangeObj.UploadandAddtional(testData);
                                break;
                            case "PDFandExcelIcon_Bulk":
                                bulkChangeObj.PDFandExcelIcon_Bulk(testData);
                                break;
                            
                    }   
                        if (!resultStatus.Equals(Enum.LogStatus.Passed))
                        {
                            ExtentTestManager._parentTest.Log(Status.Fail, "Automation Unable to locate the Element " + "  " + tcid + " -- " + keyword + " -- " + data);
                            GeneralMethod.ScreenShotCapture();
                            Assert.Fail(resultStatus.ToString());
                        }
                    }
                }
            
        }
    }
}
