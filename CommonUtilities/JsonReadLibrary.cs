using NUnit.Framework;

using System;
using System.Collections.Generic;

using Newtonsoft.Json;
using System.IO;
using System.ComponentModel;

namespace LexBaseFramework.CommonUtilities
{
    public class JsonReadLibrary
    {

        #region Members

        private string Locator_Type;
        private string Locator_Name;
        private string Locator_Value;
        #endregion

        #region Prperties
        [DisplayNameAttribute("LocatorType LocatorName LocatorValue")]
        public string LocatorType
        {
            get
            {
                return Locator_Type;
            }
            set
            {
                Locator_Type = value;
            }          
        }
        public string LocatorName
        {
            get
            {
                return Locator_Name;
            }
            set
            {
                Locator_Name = value;
            }
        }
        public string LocatorValue
        {
            get
            {
                return Locator_Value;
            }
            set
            {
                Locator_Value = value;
            }
        }
        #endregion

    
    }
}
