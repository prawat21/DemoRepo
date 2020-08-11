using NUnit.Framework;
using System;
using System.Configuration;
using System.IO;
using System.Net.Mail;
using System.Net.Mime;
using LexBaseFramework.CommonUtilities;


namespace LexBaseFramework.CommonUtilities
{
    [SetUpFixture]
    public class ZipUtils
    {
        string body = string.Empty;

        [OneTimeTearDown]
        public void Zip()
        {
            GeneralMethod.createZipFile();
            int portNumber = Convert.ToInt32(ConfigurationManager.AppSettings["Port"]);
            try
            {
                ContentType ctype = new ContentType("application/zip");
                string zipFile = GeneralMethod.GetZipFolderPath() + @"\ExtentReport.zip";
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
        }
    }
}