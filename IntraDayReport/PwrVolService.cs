using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using Services;


namespace IntraDayReport
{
    public partial class PwrVolService : ServiceBase
    {
        private Timer timer = null;
        public PwrVolService()
        {
            InitializeComponent();
        }

        /// <summary>
        /// The OnStart Event Method when Windows Service Starts.
        /// </summary>
        protected override void OnStart(string[] args)
        {
            // variable Declaration
            int intIntervalMinutes = 0, intIntervalMiliSeconds = 0;
			string strTimeInterval;

            try
            {
                PwrServiceLog.WriteEventLog("");
                PwrServiceLog.WriteEventLog("========== PWR PERIOD VOLUME AGGREGATION WINDOWS SERVICE EVENT: START ==========");
                PwrServiceLog.WriteEventLog("");

                // Function Call for Power period Volume Aggregation
                PwrServiceLog.WriteEventLog("");
                PwrServiceLog.WriteEventLog("****** Calling Power Period Volume Aggregation function to CSV Extract on instant Windows Service Start. ******");
                PwrServiceLog.WriteEventLog("");
                PwrPositionAggCalc();

                // Retrieving running time interval from Application Configuration File
                strTimeInterval = System.Configuration.ConfigurationManager.AppSettings["TimeInterval_Minutes"];
                PwrServiceLog.WriteEventLog("Retrieved Windows Service Re-Run Time interval from Application Configuration, Interval in Minutes: " + strTimeInterval);

                intIntervalMinutes = 0;
                Int32.TryParse(strTimeInterval, out intIntervalMinutes);
                // 1 Second = 1000 MiliSecond, Converting into MiliSeconds
                intIntervalMiliSeconds = intIntervalMinutes * 60 * 1000;
                PwrServiceLog.WriteEventLog("Re-Run Time Interval in Minutes: " + strTimeInterval + " ,converted into MiliSeconds: " + intIntervalMiliSeconds);

                // initialize the timer 
                timer = new Timer();

                // Set the timer with value in MiliSeconds
                this.timer.Interval = intIntervalMiliSeconds;
                PwrServiceLog.WriteEventLog("Windows Service Re-run timer has been set for minutes: " + strTimeInterval + " OR MiliSeconds: " + intIntervalMiliSeconds);

                // Calling Power Period Calculation Method through time elapsed event handler method
                PwrServiceLog.WriteEventLog("");
                PwrServiceLog.WriteEventLog("****** Calling Power Period Volume Aggregation function through time elapsed event handler. ******");
                PwrServiceLog.WriteEventLog("");
                this.timer.Elapsed += new ElapsedEventHandler(this.TimerElapsedEventMethod);
                this.timer.Enabled = true;
               
            }
            catch (Exception objEx)
            {
                PwrServiceLog.WriteEventLog("%%% EXCEPTION ERROR %%% - Exception occurred in OnStart Event Method. Exception Message: " + objEx.Message.ToString());
            }
         }

        /// <summary>
        /// This Method triggers when time elapsed event handler Starts at fixed time interval.
        /// </summary>             
        private void TimerElapsedEventMethod(object sender, ElapsedEventArgs e)
        {
            // Function Call for Power period Volume Aggregation
            PwrServiceLog.WriteEventLog("Method PwrPositionAggCalc() triggers when time elapsed event handler starts.");
            PwrPositionAggCalc();
        }

        /// <summary>
        /// Aggregate Power period volumes and extract to CSV file 
        /// </summary>
        /// <parameter name></parameter>
        /// <returns></returns>
        private void PwrPositionAggCalc()
        {
            PwrServiceLog.WriteEventLog("");
            PwrServiceLog.WriteEventLog("$$$$$$ Power Trades Period Volume Aggregation Operation Starts. $$$$$$");

            double dblVol1 = 0.0, dblVol2 = 0.0, dblVol3 = 0.0, dblVol4 = 0.0, dblVol5 = 0.0, dblVol6 = 0.0, dblVol7 = 0.0, dblVol8 = 0.0,
                   dblVol9 = 0.0, dblVol10 = 0.0, dblVol11 = 0.0, dblVol12 = 0.0, dblVol13 = 0.0, dblVol14 = 0.0, dblVol15 = 0.0, dblVol16 = 0.0,
                   dblVol17 = 0.0, dblVol18 = 0.0, dblVol19 = 0.0, dblVol20 = 0.0, dblVol21 = 0.0, dblVol22 = 0.0, dblVol23 = 0.0, dblVol24 = 0.0;

            int intCounter = 0, intTradePeriod=0;
			string strExtractYear, strExtractMonth, strExtractDate, strEroorMsg, strFileLocation, strFileName, strPath; 

            // DataTable Object Declaration
            DataTable tblPwrVolume = null;

            try
            {
                // Retrieving extract date from Application Configuration File
                PwrServiceLog.WriteEventLog("Retrieving extract date from Application Configuration File ");
                strExtractYear = System.Configuration.ConfigurationManager.AppSettings["ExtractRunYear"];
                strExtractMonth = System.Configuration.ConfigurationManager.AppSettings["ExtractRunMonth"];
                strExtractDate = System.Configuration.ConfigurationManager.AppSettings["ExtractRunDate"];

                DateTime dtExtractDate = new DateTime(Int32.Parse(strExtractYear), Int32.Parse(strExtractMonth), Int32.Parse(strExtractDate));
                PwrServiceLog.WriteEventLog("Service running for extract Date: " + dtExtractDate.ToString());

                // initialize object for Power Service
                PowerService pwrService = new PowerService();
                PwrServiceLog.WriteEventLog("PowerService DLL Object initialized.");

                // Calling PowerService DLL Method GetTrades() 
                PwrServiceLog.WriteEventLog("PowerService DLL Method GetTrades() Call - To retrieve all day-ahead trades and their period/volumes.");
                var TradesAll = pwrService.GetTrades(dtExtractDate);
                
                foreach (var tradeCount in TradesAll)
                {
                    // setting trade counter
                    intCounter += 1;
                    PwrServiceLog.WriteEventLog("****** Power Trade Number: " + intCounter + " - Period & Volume Retrieval Starts. ******");

                    foreach (var powerTrade in tradeCount.Periods)
                    {
                        PwrServiceLog.WriteEventLog("Power Trade " + intCounter + " ,Period: " + powerTrade.Period.ToString() + " ,Volume: " + powerTrade.Volume.ToString());
                        // Aggregation of Power Trade Volumes as per the period respectively.
                        intTradePeriod = powerTrade.Period;
                        switch (intTradePeriod)
                        {
                            case 1:
                                dblVol1 += powerTrade.Volume;
                                break;
                            case 2:
                                dblVol2 += powerTrade.Volume;
                                break;
                            case 3:
                                dblVol3 += powerTrade.Volume;
                                break;
                            case 4:
                                dblVol4 += powerTrade.Volume;
                                break;
                            case 5:
                                dblVol5 += powerTrade.Volume;
                                break;
                            case 6:
                                dblVol6 += powerTrade.Volume;
                                break;
                            case 7:
                                dblVol7 += powerTrade.Volume;
                                break;
                            case 8:
                                dblVol8 += powerTrade.Volume;
                                break;
                            case 9:
                                dblVol9 += powerTrade.Volume;
                                break;
                            case 10:
                                dblVol10 += powerTrade.Volume;
                                break;
                            case 11:
                                dblVol11 += powerTrade.Volume;
                                break;
                            case 12:
                                dblVol12 += powerTrade.Volume;
                                break;
                            case 13:
                                dblVol13 += powerTrade.Volume;
                                break;
                            case 14:
                                dblVol14 += powerTrade.Volume;
                                break;
                            case 15:
                                dblVol15 += powerTrade.Volume;
                                break;
                            case 16:
                                dblVol16 += powerTrade.Volume;
                                break;
                            case 17:
                                dblVol17 += powerTrade.Volume;
                                break;
                            case 18:
                                dblVol18 += powerTrade.Volume;
                                break;
                            case 19:
                                dblVol19 += powerTrade.Volume;
                                break;
                            case 20:
                                dblVol20 += powerTrade.Volume;
                                break;
                            case 21:
                                dblVol21 += powerTrade.Volume;
                                break;
                            case 22:
                                dblVol22 += powerTrade.Volume;
                                break;
                            case 23:
                                dblVol23 += powerTrade.Volume;
                                break;
                            case 24:
                                dblVol24 += powerTrade.Volume;
                                break;
                        }
                    }
                    PwrServiceLog.WriteEventLog("****** Power Trade Number: " + intCounter + " - All Period & Respective Volume Retrieval Completed. ******");
                 }
 
                // if PowerService DLL Method GetTrades() returns 0 trades then throw an exception.
                if (intCounter == 0)
                {
                    strEroorMsg = "%%% ERROR: %%% - There is no day ahead power trades retrieved by calling PowerService DLL Method GetTrades(), Trade Count = " + intCounter;
                    PwrServiceLog.WriteEventLog(strEroorMsg);
                    throw new Exception("An generic error occurred in PwrPositionAggCalc function. Exception Message: " + strEroorMsg);
                }
                else
                {
                    PwrServiceLog.WriteEventLog("@@@@@@ Day Ahead Power Trade's - Period wise Volume Aggregation Completed, Total Trade Count: " + intCounter + " @@@@@@");
                }

                // DataTable Initialization to store Power Periods & Aggregated Volumes
                tblPwrVolume = new DataTable();
                // Column Name Addition in DataTable tblPwrVolume
                tblPwrVolume.Columns.Add("Local Time", typeof(string));
                tblPwrVolume.Columns.Add("Volume", typeof(Double));
                PwrServiceLog.WriteEventLog("Data Table Created to hold Power Trade's Power Periods and respective aggregated volumes.");

                // Adding Power Periods & Aggregated Volumes into DataTable
                tblPwrVolume.Rows.Add("23:00", dblVol1);
                tblPwrVolume.Rows.Add("00:00", dblVol2);
                tblPwrVolume.Rows.Add("01:00", dblVol3);
                tblPwrVolume.Rows.Add("02:00", dblVol4);
                tblPwrVolume.Rows.Add("03:00", dblVol5);
                tblPwrVolume.Rows.Add("04:00", dblVol6);
                tblPwrVolume.Rows.Add("05:00", dblVol7);
                tblPwrVolume.Rows.Add("06:00", dblVol8);
                tblPwrVolume.Rows.Add("07:00", dblVol9);
                tblPwrVolume.Rows.Add("08:00", dblVol10);
                tblPwrVolume.Rows.Add("09:00", dblVol11);
                tblPwrVolume.Rows.Add("10:00", dblVol12);
                tblPwrVolume.Rows.Add("11:00", dblVol13);
                tblPwrVolume.Rows.Add("12:00", dblVol14);
                tblPwrVolume.Rows.Add("13:00", dblVol15);
                tblPwrVolume.Rows.Add("14:00", dblVol16);
                tblPwrVolume.Rows.Add("15:00", dblVol17);
                tblPwrVolume.Rows.Add("16:00", dblVol18);
                tblPwrVolume.Rows.Add("17:00", dblVol19);
                tblPwrVolume.Rows.Add("18:00", dblVol20);
                tblPwrVolume.Rows.Add("19:00", dblVol21);
                tblPwrVolume.Rows.Add("20:00", dblVol22);
                tblPwrVolume.Rows.Add("21:00", dblVol23);
                tblPwrVolume.Rows.Add("22:00", dblVol24);

                PwrServiceLog.WriteEventLog("Data Table filled with Day Ahead Power Periods and aggregated volumes.");

                // CSV File location retrieve from Application Configuration file
                strFileLocation = System.Configuration.ConfigurationManager.AppSettings["CsvFileLocation"];
                PwrServiceLog.WriteEventLog("CSV File location retrieved from Application Configuration file as: " + strFileLocation);
                strFileName = "\\PowerPosition_" + DateTime.Now.ToString("yyyyMMdd_HHmm") + ".csv";

                strPath = strFileLocation + strFileName;
                // Function Call to Create CSV File
                CsvExtract(strPath, tblPwrVolume, true);

                // Release DataTable
                tblPwrVolume.Dispose();

                PwrServiceLog.WriteEventLog("$$$$$$ Power Trades Period Volume Aggregation Operation completed Successfully.. $$$$$$");
                PwrServiceLog.WriteEventLog("");
            }
            catch (Exception objEx)
            {
                PwrServiceLog.WriteEventLog("%%% EXCEPTION ERROR %%% - Exception occurred in PwrPositionAggCalc Method. Exception Message: " + objEx.Message.ToString());
                throw new Exception("An generic error occurred in PwrPositionAggCalc function. Exception Message:" + objEx.Message.ToString());
            }
            finally
            {
                // Release DataTable
                tblPwrVolume.Dispose();
            }
        }

        /// <summary>
        /// Write Data and Create CSV File on location provided  
        /// </summary>
        /// <parameter name="strPath">CSV path location</parameter>
        /// <parameter name="tblDataTable">data table object containing data needed to be written into CSV File</parameter>
        /// <parameter name="boolIsFirstRowHeader">boolean parameter to check header is needed to set or not in CSV file</parameter>
        /// <returns></returns>
        static void CsvExtract(string strPath, DataTable tblDataTable, bool boolIsFirstRowHeader)
        {
            PwrServiceLog.WriteEventLog("CSV Extract Function Starts.");
            // list of strings object declaration
            List<string> lstStringList = null;

            try
            {
                // object initialization
                lstStringList = new List<string>();

                // if there are headers add them to the file first
                if (boolIsFirstRowHeader)
                {
                    string[] strColNames = tblDataTable.Columns.Cast<DataColumn>().Select(column => column.ColumnName).ToArray();
                    string strHeader = string.Join(",", strColNames);
                    lstStringList.Add(strHeader);
                    PwrServiceLog.WriteEventLog("Header has added into CSV File.");
                }

                // Place commas between the elements of each row
                var valueLines = tblDataTable.AsEnumerable().Cast<DataRow>().Select(row => string.Join(",", row.ItemArray.Select(o => o.ToString()).ToArray()));

                // Stuff the rows into a string joined by new line characters
                var allLines = string.Join(Environment.NewLine, valueLines.ToArray<string>());
                lstStringList.Add(allLines);

                // Write Data and Create File
                File.WriteAllLines(strPath, lstStringList.ToArray());
                PwrServiceLog.WriteEventLog("@@@@@@ CSV File Successfully created, Path Location: " + strPath + " @@@@@@");
                PwrServiceLog.WriteEventLog("CSV Extract Function Ends.");
            }
            catch (Exception objEx)
            {
                PwrServiceLog.WriteEventLog("%%% EXCEPTION ERROR %%% - Exception occurred in CsvExtract Method. Exception Message: " + objEx.Message.ToString());
                throw new Exception("An generic error occurred in CsvExtract function. Exception Message:" + objEx.Message.ToString());
            }
            finally
            {
                // remove elements
                lstStringList.Clear();
            }
        }

        /// <summary>
        /// The OnStop Event Method when Windows Service Stops.
        /// </summary>
        protected override void OnStop()
        {
            // Service Stop Activity
            PwrServiceLog.WriteEventLog("****** Attempt to shut down the PWR PERIOD VOLUME AGGREGATION WINDOWS SERVICE. ******");
                        
            // Stop the timer and set to null
            timer.Stop();
            timer = null;
            PwrServiceLog.WriteEventLog("****** WINDOWS SERVICE shutdown by user. ******");
            PwrServiceLog.WriteEventLog("========== PWR PERIOD VOLUME AGGREGATION WINDOWS SERVICE EVENT: STOPS ==========");
            PwrServiceLog.WriteEventLog("");
        }
    }
}
