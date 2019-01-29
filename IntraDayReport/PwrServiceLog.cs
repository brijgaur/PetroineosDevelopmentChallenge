using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IntraDayReport
{
    public static class PwrServiceLog
    {
        /// <summary>
        /// Write Event Logs into the log file 
        /// </summary>
        /// <parameter name="strMessage">string message to write into the log file</parameter>
        /// <returns></returns>
        public static void WriteEventLog(string strMessage)
        {
            StreamWriter SW = null;
            try
            {
                // get current date time 
                string strDate = System.DateTime.Now.ToString("dd-MM-yyyy");

                // retrieving LogFile location from application configuration file
                string strTimeInterval = System.Configuration.ConfigurationManager.AppSettings["LogFileLocation"];
                SW = new StreamWriter(strTimeInterval + strDate + ".txt", true);

                // Writing logging message into the log file 
                SW.WriteLine(DateTime.Now.ToString() + " : " + strMessage);
            }
            catch (Exception objEx)
            {
                throw new Exception("%%% EXCEPTION ERROR %%% - An generic error occurred in WriteEventLog function during writing logs in txt file. Exception Message:" + objEx.Message.ToString());
            }
            finally
            {
                // Clearing buffer for current Stream Writer
                SW.Flush();
                // Closing the current Stream Writer
                SW.Close();
            }
        }


    }
}
