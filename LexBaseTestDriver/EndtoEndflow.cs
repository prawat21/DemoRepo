using LexBaseFramework.CommonUtilities;
using NUnit.Framework;
//using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AventStack.ExtentReports;

namespace LexBaseFramework.LexBaseTestDriver
{
    [TestFixture]
    public class EndtoEndflow : BaseTest
    {
        [Test, TestCaseSource("getData")]
        public void endtoendflow(Dictionary<string, string> data)
        {
            ExtentTestManager.CreateParentTest(GetType().Name + '-' + data["Browser"].ToString());
            try
            {
                app = new Keywords();
                app.executeKeywords(Enum.TestCaseName.EndtoEndflow.ToString(), xls, data);
                ExtentManager.Instance.Flush();
            }
            catch (System.Exception)
            {
                ExtentManager.Instance.Flush();
                Assert.Fail("");
            }
        }

        /// <summary>
        /// Desc: Method is used to get data from the excel sheet
        /// </summary>
        /// <returns></returns>
        public static object[] getData()
        {
            return DataUtility.getData(xls, Enum.TestCaseName.EndtoEndflow.ToString());
        }
    }
}
