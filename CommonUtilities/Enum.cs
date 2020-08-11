﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LexBaseFramework.CommonUtilities
{
    public class Enum
    {
        public enum BrowserName
        {
            ie,
            firefox,
            chrome,
            Headless
        }
        public enum LogStatus
        {
            Info,
            Inconclusive,
            Skipped,
            Passed,
            Warning,
            Failed
        }
        public enum LocatorType
        {
            id,
            xpath,
            name,
            classname,
            linktext,
            cssselector
        }
        public enum KeywordsColumn
        {
            TC_Id,
            Keyword,
            Description,
            LocatorName,
            Data,
            RunMode
        }
        public enum TestCaseName
        {
            EndtoEndflow,
            MSA
        }
        public enum TestCasesColumn
        {
            TC_Id,
            RunMode
        }
        public enum LocatorsColumn
        {
            LocatorType,
            LocatorName,
            LocatorValue
        }

        public enum Sheets
        {
            EnvShakeout,
            Locators,
            TestData,
            RunMode
        }
 
    }
}
