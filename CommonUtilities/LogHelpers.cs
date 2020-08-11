using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LexBaseFramework.CommonUtilities
{
    public class LogHelpers
    {
        // logfile :-- Global Declaration  
        public static string _logFileName = string.Format("{0:MM-dd-yyyy-h-mm-ss}", DateTime.Now);
        private static StreamWriter _streamw = null;

        //create file which store log file 
        public static void createLogfile()
        {

            string dir = @"Log\";
            if (Directory.Exists(dir))
            {   
                _streamw = File.AppendText(dir + _logFileName + ".log");
            }
            else
            {
                Directory.CreateDirectory(dir);
                _streamw = File.AppendText(dir + _logFileName + ".log");
            }
        }

        //create a method which can write  the text in the log file 

        public static void Write(String logMessage)
        {
            createLogfile();
            _streamw.Write("{0} {1}", DateTime.Now.ToLongTimeString(), DateTime.Now.ToLongDateString());
            _streamw.WriteLine("   {0}", logMessage);
            _streamw.Flush();

        }

    }
}
