using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using System.Data.OleDb;
using System.Threading;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace INFLOApps
{
    using INFLOClassLib;

    public partial class frmINFLOApps : Form
    {
        //private double FeetPerMile = 5280;

        //clsDatabase related variables
        //private CExcelDoc m_ExcelDoc = new CExcelDoc();
        


        //private List<Microsoft.Office.Interop.Excel.Worksheet> m_worksheets;
        private Microsoft.Office.Interop.Excel.Range m_workSheet_range = null;

        private int m_row;
        private int m_col;
        private int m_startupRow;
        private int m_startupColomn;

        private clsDatabase DB;

        //Configuration files variables
        private string INFLOConfigFile = string.Empty;

        //Name fo file to sync simulation program with INFLOApps program
        private string SyncFileName = string.Empty;
        private bool CVDataFlag = false;
        private bool TSSDataFlag = false;

        //Roadway entity lists
        private clsRoadway Roadway = new clsRoadway();
        private List<clsRoadway> RList = new List<clsRoadway>();
        private List<clsRoadwayLink> RLList = new List<clsRoadwayLink>();
        private List<clsRoadwaySubLink> RSLList = new List<clsRoadwaySubLink>();
        private List<clsDetectorStation> DSList = new List<clsDetectorStation>();
        private List<clsDetectionZone> DZList = new List<clsDetectionZone>();

        private StreamWriter TSSDataProcessor;
        private StreamWriter CVDataProcessor;
        private StreamWriter QueueLog;
        private StreamWriter SubLinKDataLog;
        private StreamWriter FillDataSetLog;

        private int NumberNoCVDataIntervals = 0;
        private int NumberNoTSSDataIntervals = 0;
        
        private DateTime CVPrevWakeupTime = DateTime.Now;
        private DateTime CVCurrWakeupTime = DateTime.Now;
        private double CVTimeDiff = 0;

        private clsRoadwaySubLink QueuedSubLink;
        private clsRoadwayLink QueuedLink;

        //CV Data Aggregator related variables
        private List<clsCVData> CurrIntervalCVList = new List<clsCVData>();
        private DateTime DateGenerated = DateTime.Now;

        //TSS data aggregator interval related variables
        private List<clsDetectionZone> CurrIntervalTSSDataList = new List<clsDetectionZone>();

        private double TroupingEndMM = 0;
        private double TroupingEndSpeed = 0;

        private frmINFLODisplay DisplayForm = new frmINFLODisplay();

        private bool Stopped = false;

        private void LogTxtMsg(System.Windows.Forms.TextBox txtControl, string Text)
        {
            Text = "\r\n" + Text;
            //Text = Environment.NewLine + "\t" + Text;
            if (txtControl.Text.Length > 30000)
                txtControl.Text = "";

            txtControl.SelectionStart = txtControl.Text.Length;
            txtControl.SelectedText = Text;
            txtControl.SelectionStart = txtControl.Text.Length;
        }
        
        public frmINFLOApps()
        {
            InitializeComponent();
        }

        //Excel
        Microsoft.Office.Interop.Excel.Application CVQWARNExcelApp = new Microsoft.Office.Interop.Excel.Application();
        Microsoft.Office.Interop.Excel.Workbook CVWorkbook = null;
        Microsoft.Office.Interop.Excel.Worksheet[] CVWorkSheets = new Microsoft.Office.Interop.Excel.Worksheet[3];

        Microsoft.Office.Interop.Excel.Application CVSPDHarmExcelApp = new Microsoft.Office.Interop.Excel.Application();
        Microsoft.Office.Interop.Excel.Workbook CVSPDHarmWorkbook = null;
        Microsoft.Office.Interop.Excel.Worksheet[] CVSPDHarmWorkSheets = new Microsoft.Office.Interop.Excel.Worksheet[3];

        Microsoft.Office.Interop.Excel.Application TSSQWARNExcelApp = new Microsoft.Office.Interop.Excel.Application();
        Microsoft.Office.Interop.Excel.Workbook TSSWorkbook = null;
        Microsoft.Office.Interop.Excel.Worksheet[] TSSWorkSheets = new Microsoft.Office.Interop.Excel.Worksheet[3];

        int TSSWSCurrRow = 0;
        int CVWSCurrRow = 0;
        int CVSPDHarmWSCurrRow = 0;

        private void frmINFLOApps_Load(object sender, EventArgs e)
        {
            string retValue = string.Empty;
            //Declare Microsoft Excel workbooks used for displaying link and sublink data.
            //Microsoft.Office.Interop.Excel.Application CVQWARNExcelApp = new Microsoft.Office.Interop.Excel.Application();
            //Microsoft.Office.Interop.Excel.Application TSSQWARNExcelApp = new Microsoft.Office.Interop.Excel.Application();
            DisplayForm.Show();
            //DisplayForm.Refresh();

            this.Show();
            this.Refresh();

            //Excel
            /*CVQWARNExcelApp.Visible = true;
            CVSPDHarmExcelApp.Visible = true;
            TSSQWARNExcelApp.Visible = true;
            CVQWARNExcelApp.StandardFont = "Arial Narrow";
            CVSPDHarmExcelApp.StandardFont = "Arial Narrow";
            TSSQWARNExcelApp.StandardFont = "Arial Narrow";
            CVQWARNExcelApp.StandardFontSize = 12;
            CVSPDHarmExcelApp.StandardFontSize = 12;
            TSSQWARNExcelApp.StandardFontSize = 12;
            */
            //Microsoft.Office.Interop.Excel.Workbook CVWorkbook = null;
            //Microsoft.Office.Interop.Excel.Workbook TSSWorkbook = null;
            
            
            /*CVWorkbook = CVQWARNExcelApp.Workbooks.Add(1);
            CVSPDHarmWorkbook = CVSPDHarmExcelApp.Workbooks.Add(1);
            TSSWorkbook = TSSQWARNExcelApp.Workbooks.Add(1);
            */
            
            //Microsoft.Office.Interop.Excel.Worksheet[] CVWorkSheets = new Microsoft.Office.Interop.Excel.Worksheet[3];
            //Microsoft.Office.Interop.Excel.Worksheet[] TSSWorkSheets = new Microsoft.Office.Interop.Excel.Worksheet[3];

            splitContainer1.SplitterDistance = 150;
            splitContainer4.SplitterDistance = 50;
            splitContainer5.SplitterDistance = 50;

            string tmpFileName = DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day + "-" + DateTime.Now.Hour + "-" + DateTime.Now.Minute + "-" + DateTime.Now.Second;
            FillDataSetLog = new StreamWriter(System.Windows.Forms.Application.StartupPath + "\\FillDataSet-" + tmpFileName + ".csv");
            FillDataSetLog.WriteLine(DateTime.Now);
            FillDataSetLog.WriteLine("Date/Time, DataType, NumberRecords,");

            #region "Read the INFLO configuration file and initialize the INFLO thresholds and parameters"
            LogTxtMsg(txtINFLOConfigLog, "\r\nRead the INFLO configuration file and intialize the INFLO algorithms thresholds and parameters.");

            INFLOConfigFile = System.Windows.Forms.Application.StartupPath + "\\Config\\INFLOConfig.xml";
            txtINFLOConfigFile.Text = INFLOConfigFile;
            txtINFLOConfigFile.Refresh();
            if (INFLOConfigFile.Length > 0)
            {
                retValue = clsMiscFunctions.ReadINFLOConfigFile(INFLOConfigFile, ref Roadway);
                if (retValue.Length > 0)
                {
                    LogTxtMsg(txtINFLOConfigLog, retValue);
                    LogTxtMsg(txtINFLOConfigLog, "\r\nThe INFLO application is terminating.");
                    return;
                }

                txtRoadwayLinkConfigFile.Text = clsGlobalVars.RoadwayLinkConfigFile;
                txtRoadwayLinkConfigFile.Refresh();
                txtDetectionStationConfigFile.Text = clsGlobalVars.DetectionZoneConfigFile;
                txtDetectionStationConfigFile.Refresh();
                txtDetectionZoneConfigFile.Text = clsGlobalVars.DetectionZoneConfigFile;
                txtDetectionZoneConfigFile.Refresh();
                DisplayForm.txtFOQ.Text = Roadway.RecurringCongestionMMLocation.ToString("0.00");
            }
            else
            {
                LogTxtMsg(txtINFLOConfigLog, "\tThe INFLO configuration file name was not specified in the Global variables class.\r\nThe INFLO Application is terminating.");
                return;
            }
            #endregion

            if (clsGlobalVars.CVDataSmoothedSpeedArraySize == 0)
            {
                clsGlobalVars.CVDataSmoothedSpeedArraySize  = (int)(Math.Ceiling((double)clsGlobalVars.TSSDataPollingFrequency / (double)clsGlobalVars.CVDataPollingFrequency));
            }

            #region "Establish connection to INFLO database"
            LogTxtMsg(txtINFLOConfigLog, "\r\n\tEstablish connection to the INFLO database: " +
                              "\r\n\t\tDBInterfaceType: " + clsGlobalVars.DBInterfaceType +
                              "\r\n\t\tAccessDBFileName: " + clsGlobalVars.AccessDBFileName +
                              "\r\n\t\tAccessDBFileName: " + clsGlobalVars.DSNName +
                              "\r\n\t\tAccessDBFileName: " + clsGlobalVars.SqlServer +
                              "\r\n\t\tAccessDBFileName: " + clsGlobalVars.SqlServerDatabase +
                              "\r\n\t\tAccessDBFileName: " + clsGlobalVars.SqlServerUserId +
                              "\r\n\t\tAccessDBFileName: " + clsGlobalVars.SqlStrConnection);
            if (clsGlobalVars.DBInterfaceType.Length > 0)
            {
                DB = new clsDatabase(clsGlobalVars.DBInterfaceType);
                if (DB.ConnectionStr.Length > 0)
                {
                    LogTxtMsg(txtINFLOConfigLog, "\r\n\tDatabase Connection string: " + DB.ConnectionStr);
                    retValue = DB.OpenDBConnection();
                    if (retValue.Length > 0)
                    {
                        LogTxtMsg(txtINFLOConfigLog, retValue);
                        LogTxtMsg(txtINFLOConfigLog, "\r\nThe INFLO application is terminating.");
                        return;
                    }
                }
                else
                {
                    LogTxtMsg(txtINFLOConfigLog, "\r\n\tError in generating connection string to INFLO database.");
                    return;
                }
            }
            else
            {
                LogTxtMsg(txtINFLOConfigLog, "\tThe INFLO Application can not connect to the INFLO database. " +
                                             "\r\n\tThe DBInterfaceType= " + clsGlobalVars.DBInterfaceType + "  is not specified.");
                return;
            }
            #endregion

            #region "Get available Roadway info from INFLO database"
            
            //LogTxtMsg(txtINFLOConfigLog, "Get list of available Roadways from the INFLO database: ");
            //retValue = GetRoadwayInfo(DB, ref RList);
            //if (retValue .Length > 0)
            //{
            //    LogTxtMsg(txtINFLOConfigLog, "\tError in getting the available roadways from the INFLO database: \r\n" + retValue);
            //    return;
            //}
            #endregion

            #region "Get available Roadway Links from INFLO database"
            //Get list of available infrastructure roadway links from the INFLO database
            LogTxtMsg(txtINFLOConfigLog, "Get list of available Roadway Links from the INFLO database: ");
            RLList.Clear();
            retValue = GetRoadwayLinks(DB, ref RLList);
            if (retValue.Length > 0)
            {
                LogTxtMsg(txtINFLOConfigLog, "\tError in getting the available roadway links from the INFLO database: \r\n" + retValue);
                return;
            }
            #endregion

            #region "Get available Roadway Sub-Links from INFLO database"
            //Get list of available infrastructure roadway sub-links from the INFLO database
            LogTxtMsg(txtINFLOConfigLog, "Get list of available Roadway Sub-Links from the INFLO database: ");
            RSLList.Clear();
            retValue = GetRoadwaySubLinks(DB, ref RSLList);
            if (retValue.Length > 0)
            {
                LogTxtMsg(txtINFLOConfigLog, "\tError in getting the available roadway sub-links from the INFLO database: \r\n" + retValue);
                return;
            }

            foreach (clsRoadwaySubLink rsl in RSLList)
            {
                rsl.SmoothedSpeed = new double[clsGlobalVars.CVDataSmoothedSpeedArraySize];
                rsl.SmoothedSpeedIndex = 0;
            }
            #endregion

            #region "Get available Detector Stations from INFLO database"
            //Get list of available infrastructure detector stations from the INFLO database
            LogTxtMsg(txtINFLOConfigLog, "Get list of available Detector stations from the INFLO database: ");
            DSList.Clear();
            retValue = GetDetectorStations(DB, ref DSList);
            if (retValue.Length > 0)
            {
                LogTxtMsg(txtINFLOConfigLog, "\tError in getting the available detector stations from the INFLO database: \r\n" + retValue);
                return;
            }
            #endregion

            #region "Get available Detection Zones from INFLO database"
            //Get list of available infrastructure detector zones from the INFLO database
            LogTxtMsg(txtINFLOConfigLog, "Get list of available Detector stations from the INFLO database: ");
            DZList.Clear();
            retValue = GetDetectionZones(DB, ref DZList);
            if (retValue.Length > 0)
            {
                LogTxtMsg(txtINFLOConfigLog,  retValue);
                return;
            }
            #endregion


            //Open CV and TSS data processing log files

            TSSDataProcessor = new StreamWriter(System.Windows.Forms.Application.StartupPath + "\\TSSDataProcessor-" + tmpFileName + ".Txt");
            TSSDataProcessor.WriteLine(DateTime.Now);
            TSSDataProcessor.WriteLine(DateTime.Now + "\r\n\tTSS BOQ: " + clsGlobalVars.InfrastructureBOQMMLocation);
            TSSDataProcessor.WriteLine(DateTime.Now + "\r\n\tTSS BOQTime: " + clsGlobalVars.InfrastructureBOQTime);
            TSSDataProcessor.WriteLine(DateTime.Now + "\r\n\tTSS BOQ: " + clsGlobalVars.CVBOQMMLocation);
            TSSDataProcessor.WriteLine(DateTime.Now + "\r\n\tTSS BOQTime: " + clsGlobalVars.CVBOQTime);
            TSSDataProcessor.WriteLine(DateTime.Now + "\r\n\tTSS BOQ: " + clsGlobalVars.BOQMMLocation);
            TSSDataProcessor.WriteLine(DateTime.Now + "\r\n\tTSS BOQ: " + clsGlobalVars.BOQTime);
            TSSDataProcessor.WriteLine(DateTime.Now + "\r\n\tTSS BOQ: " + clsGlobalVars.QueueRate);
            TSSDataProcessor.WriteLine(DateTime.Now + "\r\n\tTSS BOQ: " + clsGlobalVars.QueueChange);
            TSSDataProcessor.WriteLine(DateTime.Now + "\r\n\tTSS BOQ: " + clsGlobalVars.QueueSource);

            CVDataProcessor = new StreamWriter(System.Windows.Forms.Application.StartupPath + "\\CVDataProcessor-" + tmpFileName + ".Txt");
            CVDataProcessor.WriteLine(DateTime.Now);

            SubLinKDataLog = new StreamWriter(System.Windows.Forms.Application.StartupPath + "\\SubLinkDataLog-" + tmpFileName + ".csv");
            SubLinKDataLog.WriteLine(DateTime.Now);


            //Excel
            //Initialize Microsoft CV and TSS worksheets used to display link and sublink queue data
            /*TSSWorkSheets[1] = TSSWorkbook.Worksheets.Add();
            TSSWorkSheets[1].Name = "LinkQueuedStatus";
            retValue = InitializeLinkQueuedStateWorksheet(ref TSSWorkSheets[1], RLList);
            TSSWSCurrRow = 2;

            CVSPDHarmWorkSheets[1] = CVSPDHarmWorkbook.Worksheets.Add();
            CVSPDHarmWorkSheets[1].Name = "SublinkTroupeStatus";
            retValue = InitializeSublinkTroupeWorksheet(ref CVSPDHarmWorkSheets[1], RSLList);
            CVSPDHarmWSCurrRow = 2;

            CVWorkSheets[1] = CVWorkbook.Worksheets.Add();
            CVWorkSheets[1].Name = "SubLinkQueuedState";
            retValue =  InitializeSublinkQueuedStateWorksheet(ref CVWorkSheets[1], RSLList);
            CVWSCurrRow = 2;*/

            //Set the infrastructure data availability and initialize the infrastructure BOQ variables
            if (RLList.Count > 0)
            {
                clsGlobalVars.InfrastructureDataAvailable = true;
                TSSDataProcessor.WriteLine(DateTime.Now + ",  Infrastructure data available");
            }
            else
            {
                clsGlobalVars.InfrastructureDataAvailable = false;
                TSSDataProcessor.WriteLine(DateTime.Now + ",  Infrastructure data not available");
            }

            clsGlobalVars.CVDataAvailable = true;

            LogTxtMsg(txtINFLOConfigLog, "\r\nFinished processing the INFLO Application configuration files");


        // Start Insert Code -SS

            //             /*
            
            string retValue1 = string.Empty;

            if ((txtSyncFileName.Text.Trim()).Length > 0)
            {
                SyncFileName = txtSyncFileName.Text.Trim();
            }
            else
            {
                MessageBox.Show("Please enter the Name and Location of the file used to sync the Simulatin" +
                                 "program and the INFLOApps program in the Text box next to the Start button and then click the Start button again.");
                return;
            }
            LogTxtMsg(txtINFLOLog, "\r\nStart the connected vehicle, infrastructure, and weather data processors if data sources are available");

            btnStartINFLO.Enabled = false;
            retValue1 = StartTrafficDataProcessors();
            if (retValue1.Length > 0)
            {
                LogTxtMsg(txtINFLOLog, "\r\n\tError in starting the INFLO Q-WARN and SPD-HARM algorithms.\r\n\t" + retValue1);
                btnStartINFLO.Enabled = true;
                btnStopINFLO.Enabled = false;
                tmrCVData.Enabled = false;
                tmrTSSData.Enabled = false;
            }
            else
            {
                DisplayForm.txtBOQ.Text = string.Empty;
                DisplayForm.txtCVBOQ.Text = string.Empty;
                DisplayForm.txtCVDate.Text = string.Empty;
                DisplayForm.txtTSSBOQ.Text = string.Empty;
                DisplayForm.txtTSSDate.Text = string.Empty;
                DisplayForm.ClearCVSubLinkQueuedStatus();
                DisplayForm.ClearCVSubLinkTroupeStatus();
                DisplayForm.ClearCVSubLinkSPDHarmStatus();
                DisplayForm.ClearTSSQueuedLinkStatus();
                DisplayForm.ClearTSSSPDHarmLinkStatus();

                btnStopINFLO.Enabled = true;
                Stopped = false;
                tmrfile.Interval = 5000;
                tmrfile.Enabled = true;
            }

            LogTxtMsg(txtINFLOLog, "\r\nFinished Starting the CV, infrastructure and weather data processors");

            // */

        // End  Inser Code - SS
        }

        private void btnStartINFLO_Click_1(object sender, EventArgs e)
        {
            string retValue = string.Empty;

            if ((txtSyncFileName.Text.Trim()).Length > 0)
            {
                SyncFileName = txtSyncFileName.Text.Trim();
            }
            else
            {
                MessageBox.Show("Please enter the Name and Location of the file used to sync the Simulatin" + 
                                 "program and the INFLOApps program in the Text box next to the Start button and then click the Start button again.");
                return;
            }
            LogTxtMsg(txtINFLOLog, "\r\nStart the connected vehicle, infrastructure, and weather data processors if data sources are available");

            btnStartINFLO.Enabled = false;
            retValue = StartTrafficDataProcessors();
            if (retValue.Length > 0)
            {
                LogTxtMsg(txtINFLOLog, "\r\n\tError in starting the INFLO Q-WARN and SPD-HARM algorithms.\r\n\t" + retValue);
                btnStartINFLO.Enabled = true;
                btnStopINFLO.Enabled = false;
                tmrCVData.Enabled = false;
                tmrTSSData.Enabled = false;
            }
            else
            {
                DisplayForm.txtBOQ.Text = string.Empty;
                DisplayForm.txtCVBOQ.Text = string.Empty;
                DisplayForm.txtCVDate.Text = string.Empty;
                DisplayForm.txtTSSBOQ.Text = string.Empty;
                DisplayForm.txtTSSDate.Text = string.Empty;
                DisplayForm.ClearCVSubLinkQueuedStatus();
                DisplayForm.ClearCVSubLinkTroupeStatus();
                DisplayForm.ClearCVSubLinkSPDHarmStatus();
                DisplayForm.ClearTSSQueuedLinkStatus();
                DisplayForm.ClearTSSSPDHarmLinkStatus();

                btnStopINFLO.Enabled = true;
                Stopped = false;
                tmrfile.Interval = 5000;
                tmrfile.Enabled = true;
            }

            LogTxtMsg(txtINFLOLog, "\r\nFinished Starting the CV, infrastructure and weather data processors");
        }

        private void btnStopINFLO_Click_1(object sender, EventArgs e)
        {
            btnStopINFLO.Enabled = false;
            tmrCVData.Enabled = false;
            tmrTSSData.Enabled = false;
            if (QueueLog != null)
            {
                QueueLog.Close();
            }
            //if (TSSDataProcessor != null)
            //{
            //    TSSDataProcessor.Close();
            //}
            //if (CVDataProcessor != null)
            //{
            //    CVDataProcessor.Close();
            //}
            Stopped = true;
            btnStartINFLO.Enabled = true;
        }

        private string StartTrafficDataProcessors()
        {
            string retValue = string.Empty; ;


            //tabINFLOApps.SelectTab("tabPgINFLOConfiguration");
            //Open QueueLog log file
            string tmpFileName = DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day + "-" + DateTime.Now.Hour + "-" + DateTime.Now.Minute + "-" + DateTime.Now.Second;

            QueueLog = new StreamWriter(System.Windows.Forms.Application.StartupPath + "\\QueueLog-" + tmpFileName + ".csv");
            QueueLog.WriteLine(DateTime.Now);
            QueueLog.WriteLine("HH:MM:SS::MMM,PrevBOQ,PrevBOQTime,BOQ,BOQTime,MMChange, TimeDiff, QueueRate,QueueChange,Source, CVBOQ, TSSBOQ, TotalCVs, PrevTotalCVs, VolumeDiff, FlowRate, Density, PrevDensity, DensityDiff, ShockwaveRate");


            //Start the TSS, CV, and Weather data processors

            bool CVDataProcessorStarted = false;
            bool TSSDataProcessorStarted = false;

            long SSM = 0;   //seconds since midnight

            #region "Start the TSS data processor"
            while (TSSDataProcessorStarted == false)
            {
                LogTxtMsg(txtTSSDataLog, "\r\n" + DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + "::" + DateTime.Now.Millisecond +
                                         "\tWaiting for the start of the TSS data processing module......");
                if (clsGlobalVars.InfrastructureDataAvailable == true)
                {
                    //SSM = DateTime.Now.Hour * 3600 + DateTime.Now.Minute * 60 + DateTime.Now.Second;
                    //if ((SSM % clsGlobalVars.TSSDataLoadingFrequency) == 0)
                    //{
                        //tabINFLOApps.SelectTab("tabPgTSSDataAggregation");
                        TSSDataProcessorStarted = true;
                        LogTxtMsg(txtTSSDataLog, "\r\n\t\t" + DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + "::" + DateTime.Now.Millisecond + "\t\tStarting the TSS Data processor.");
                        TSSDataProcessor.WriteLine("\r\n" + DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + "::" + DateTime.Now.Millisecond + "\t\tStarted the TSS Data processor.");
                        tmrTSSData.Interval = clsGlobalVars.TSSDataLoadingFrequency * 1000;         //in milliseconds
                        Thread.Sleep(1000);

                        clsGlobalVars.InfrastructureBOQMMLocation = -1;
                        clsGlobalVars.InfrastructureBOQTime = DateTime.Now;
                        clsGlobalVars.PrevInfrastructureBOQMMLocation = -1;
                        clsGlobalVars.PrevInfrastructureBOQTime = DateTime.Now;

                        //tmrTSSData.Enabled = true;
                        //DateGenerated = DateTime.Now;
                        //tmrTSSData_Tick(sender, e);
                    //}
                }
                else
                {
                    LogTxtMsg(txtTSSDataLog, "\r\n\t\t" + DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + "::" + DateTime.Now.Millisecond +
                                             "\tThe TSS data processing module did not start. No Infrastructure data is available in the INFLO database.");
                    TSSDataProcessor.WriteLine("\r\n" + DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + "::" + DateTime.Now.Millisecond +
                                               "\t\tThe TSS data processor was not started because no infrastructure traffic sensor data is available.");
                    //MessageBox.Show("No Infrastructure data is available in the INFLO database.");
                    break;
                }
            }
            #endregion

            #region "Start the CV data processor"
            while (CVDataProcessorStarted == false)
            {
                LogTxtMsg(txtCVDataLog, "\r\n" + DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + "::" + DateTime.Now.Millisecond +
                                       "\tWaiting for the start of the CV data processing module.......");
                //SSM = DateTime.Now.Hour * 3600 + DateTime.Now.Minute * 60 + DateTime.Now.Second;
                //if ((SSM % clsGlobalVars.CVDataPollingFrequency) == 0)
                //{
                    CVDataProcessorStarted = true;
                    LogTxtMsg(txtCVDataLog, "\r\n\t\t" + DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + "::" + DateTime.Now.Millisecond + "\t\tStarting the CV data processor.");
                    CVDataProcessor.WriteLine("\r\n" + DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + "::" + DateTime.Now.Millisecond + "\t\tStarted the CV Data processor.");
                    tmrCVData.Interval = clsGlobalVars.CVDataPollingFrequency * 1000;         //in milliseconds
                    //tmrCVData.Interval = 10 * 1000;         //in milliseconds
                    Thread.Sleep(500);

                    clsGlobalVars.CVBOQMMLocation = -1;
                    clsGlobalVars.CVBOQTime = DateTime.Now;
                    clsGlobalVars.PrevCVBOQMMLocation = -1;
                    clsGlobalVars.PrevCVBOQTime = DateTime.Now;

                    //tmrCVData.Enabled = true;
                    //DateGenerated = DateTime.Now;
                    //tmrCVData_Tick(sender, e);
                //}
            }
            #endregion

            //Reset BOQ values
            clsGlobalVars.BOQMMLocation = -1;
            clsGlobalVars.BOQTime = DateTime.Now;

            clsGlobalVars.PrevBOQMMLocation = -1;
            clsGlobalVars.PrevBOQTime = DateTime.Now;

            clsGlobalVars.QueueRate = 0;
            clsGlobalVars.QueueChange = clsEnums.enQueueCahnge.NA;
            clsGlobalVars.QueueSource = clsEnums.enQueueSource.NA;
            //tmrBOQ.Interval = 5000;
            //tmrBOQ.Enabled = true;
            return retValue;
        }
        
        //Excel
        private string InitializeLinkQueuedStateWorksheet(ref Microsoft.Office.Interop.Excel.Worksheet TSSWS, List<clsRoadwayLink> RLList)
        {
            string retValue = string.Empty;

            try
            {
                TSSWS.Cells[1, 1] = "Date/Time";
                TSSWS.Cells[1, 2] = "StartInterval";
                TSSWS.Cells[1, 3] = "EndInterval";
                for (int i = 0; i < RLList.Count; i++)
                {
                    Excel.Range tmpRange = (Excel.Range)TSSWS.Range[TSSWS.Cells[1, (i * 5) + 4], TSSWS.Cells[1, (i + 1) * 5 + 3]];
                    //tmpRange = tmpRange.Select();
                    //tmpRange.ReadingOrder
                    tmpRange.Orientation = 90;
                    tmpRange.RowHeight = 95;
                    //tmpRange.AddIndent = false;
                    //tmpRange.IndentLevel = 0;
                    tmpRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    tmpRange.WrapText = false;
                    tmpRange.ShrinkToFit = false;
                    tmpRange.MergeCells = true;
                    TSSWS.Cells[1, (i * 5) + 4] = (i + 1) + " - " + RLList[i].BeginMM + "-TO-" + RLList[i].EndMM;
                    TSSWS.Cells[1, (i*5) + 4].Interior.Color = System.Drawing.Color.Orange;
                }

                //TSSWS.Range[TSSWS.Cells[1, 1], TSSWS.Cells[1, RLList.Count]].Merge();
                TSSWS.Range[TSSWS.Cells[1, 1], TSSWS.Cells[1, RLList.Count*5]].Interior.Color = System.Drawing.Color.Orange;
                TSSWS.Range[TSSWS.Cells[1, 1], TSSWS.Cells[1, RLList.Count*5]].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                TSSWS.Range[TSSWS.Cells[1, 1], TSSWS.Cells[1, RLList.Count*5]].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick;
                TSSWS.Range[TSSWS.Cells[1, 1], TSSWS.Cells[1, RLList.Count*5]].Borders.Colorindex  = 1;
                TSSWS.Range[TSSWS.Cells[1, 1], TSSWS.Cells[1, RLList.Count*5]].Style.Font.Name = "Arial Narrow";
                TSSWS.Range[TSSWS.Cells[1, 1], TSSWS.Cells[1, RLList.Count*5]].Style.Font.Bold = true;
                TSSWS.Range[TSSWS.Cells[1, 1], TSSWS.Cells[1, RLList.Count*5]].Style.Font.Size = 12;
                TSSWS.Range[TSSWS.Cells[1, 1], TSSWS.Cells[1, RLList.Count]].EntireRow.Autofit();
                TSSWS.Range[TSSWS.Cells[1, 1], TSSWS.Cells[1, RLList.Count*5]].Rows.Autofit();
                TSSWS.Range[TSSWS.Cells[1, 1], TSSWS.Cells[1, RLList.Count*5]].ColumnWidth = 3;
                TSSWS.Range[TSSWS.Cells[1, 1], TSSWS.Cells[1, RLList.Count*5]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                Excel.Range RowRange = (Excel.Range)TSSWS.Range[TSSWS.Cells[1, 1], TSSWS.Cells[1, 3]];
                RowRange.EntireColumn.ColumnWidth = 10;
                RowRange.EntireRow.RowHeight = 95;

            }
            catch (Exception ex)
            {
                retValue = "Error in intializing the TSS Link queued state Excel worksheet" + "\r\n\t" + ex.Message;
            }

            return retValue;
        }
        private string InitializeSublinkQueuedStateWorksheet(ref Microsoft.Office.Interop.Excel.Worksheet CVWS, List<clsRoadwaySubLink> RSLList)
        {
            string retValue = string.Empty;
            CVWS.Cells[1, 1] = "Date/Time";
            //TSSWS.Cells[1, 1] = "Start Interval";
            //TSSWS.Cells[1, 1] = "EndInterval";
            try
            {
                for (int i = 0; i < RSLList.Count; i++)
                {
                    CVWS.Cells[1, i+2] = (i + 1).ToString() + " - " + RSLList[i].BeginMM.ToString() +  "-TO-" + RSLList[i].EndMM.ToString();
                    
                    CVWS.Cells[1, i + 2].Interior.Color = System.Drawing.Color.Orange;
                }
                Excel.Range tmpRange = (Excel.Range)CVWS.Range[CVWS.Cells[1, 1], CVWS.Cells[1, RSLList.Count +1]];
                tmpRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                tmpRange.WrapText = false;
                tmpRange.Orientation = 90;
                tmpRange.AddIndent = false;
                tmpRange.IndentLevel = 0;
                tmpRange.ShrinkToFit = false;
                //tmpRange.ReadingOrder
                //tmpRange.MergeCells = true;
                tmpRange.Interior.Color = System.Drawing.Color.Orange;
                tmpRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                tmpRange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick;
                tmpRange.Borders.ColorIndex = 1;
                tmpRange.Style.Font.Name = "Arial Narrow";
                tmpRange.Style.Font.Bold = true;
                tmpRange.Style.Font.Size = 12;
                tmpRange.EntireRow.AutoFit();
                tmpRange.Rows.AutoFit();
                tmpRange.ColumnWidth = 3;
                tmpRange.RowHeight = 95;
                Excel.Range RowRange = (Excel.Range)CVWS.Range[CVWS.Cells[1, 1], CVWS.Cells[1000, 1]];
                RowRange.EntireColumn.ColumnWidth = 15;

            }
            catch (Exception ex)
            {
                retValue = "Error in intializing the CV SubLink queued state Excel worksheet" + "\r\n\t" + ex.Message; 
            }


            return retValue;
        }
        private string InitializeSublinkTroupeWorksheet(ref Microsoft.Office.Interop.Excel.Worksheet CVWS, List<clsRoadwaySubLink> RSLList)
        {
            string retValue = string.Empty;
            CVWS.Cells[1, 1] = "Date/Time";
            CVWS.Cells[1, 2] = "BOQ/Trouping";
            //TSSWS.Cells[1, 1] = "Start Interval";
            //TSSWS.Cells[1, 1] = "EndInterval";
            try
            {
                for (int i = 0; i < RSLList.Count; i++)
                {
                    CVWS.Cells[1, i + 3] = (i + 1).ToString() + " - " + RSLList[i].BeginMM.ToString() + "-TO-" + RSLList[i].EndMM.ToString();

                    CVWS.Cells[1, i + 3].Interior.Color = System.Drawing.Color.Orange;
                }
                Excel.Range tmpRange = (Excel.Range)CVWS.Range[CVWS.Cells[1, 1], CVWS.Cells[1, RSLList.Count + 2]];
                tmpRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                tmpRange.WrapText = false;
                tmpRange.Orientation = 90;
                tmpRange.AddIndent = false;
                tmpRange.IndentLevel = 0;
                tmpRange.ShrinkToFit = false;
                //tmpRange.ReadingOrder
                //tmpRange.MergeCells = true;
                tmpRange.Interior.Color = System.Drawing.Color.Orange;
                tmpRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                tmpRange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick;
                tmpRange.Borders.ColorIndex = 1;
                tmpRange.Style.Font.Name = "Arial Narrow";
                tmpRange.Style.Font.Bold = true;
                tmpRange.Style.Font.Size = 12;
                tmpRange.EntireRow.AutoFit();
                tmpRange.Rows.AutoFit();
                tmpRange.ColumnWidth = 3;
                tmpRange.RowHeight = 95;
                Excel.Range RowRange = (Excel.Range)CVWS.Range[CVWS.Cells[1, 1], CVWS.Cells[1, 2]];
                RowRange.EntireColumn.ColumnWidth = 15;

            }
            catch (Exception ex)
            {
                retValue = "Error in intializing the CV SubLink queued state Excel worksheet" + "\r\n\t" + ex.Message;
            }


            return retValue;
        }

        private string GetRoadwayInfo(clsDatabase DB, ref List<clsRoadway> RList)
        {
            string retValue = string.Empty;
            string sqlQuery = string.Empty;

            DataSet RoadwayDataSet = new DataSet("Roadway");
            sqlQuery = "Select * from Configuration_Roadway";

            retValue = string.Empty;
            retValue = DB.FillDataSet(sqlQuery, ref RoadwayDataSet);
            FillDataSetLog.WriteLine(DateTime.Now + ",GetRoadwayInfo," + RoadwayDataSet.Tables[0].Rows.Count);
            if (retValue.Length > 0)
            {
                return retValue;
            }

            //display roadway information if it already exists
            LogTxtMsg(txtINFLOConfigLog, "\tAvailable Roadway Info: ");

            if (RoadwayDataSet.Tables[0].Rows.Count > 0)
            {
                foreach (DataRow row in RoadwayDataSet.Tables[0].Rows)
                {
                    clsRoadway tmpRd = new clsRoadway();
                    string tmpRoadID = string.Empty;
                    string tmpRdName = string.Empty;
                    clsEnums.enDirection tmpDirection = clsEnums.enDirection.NA;
                    double tmpBeginMM = 0;
                    double tmpEndMM = 0;
                    double tmpRdGrade = 0;
                    double tmpRecurringCongestionMMLocation = 0;
                    clsEnums.enDirection tmpIncreasingDirection = clsEnums.enDirection.NA;
                    foreach (DataColumn col in RoadwayDataSet.Tables[0].Columns)
                    {
                        switch (col.ColumnName.ToString().ToLower())
                        {
                            case "roadwayid":
                                tmpRoadID = row[col].ToString();
                                tmpRd.Identifier = int.Parse(tmpRoadID);
                                break;
                            case "name":
                                tmpRdName = row[col].ToString();
                                tmpRd.Name = tmpRdName;
                                break;
                            case "grade":
                                tmpRdGrade = double.Parse(row[col].ToString());
                                tmpRd.Grade = tmpRdGrade;
                                break;
                            case "beginmm":
                                tmpBeginMM = double.Parse(row[col].ToString());
                                tmpRd.BeginMM = tmpBeginMM;
                                break;
                            case "endmm":
                                tmpEndMM = double.Parse(row[col].ToString());
                                tmpRd.EndMM = tmpEndMM;
                                break;
                            case "direction":
                                tmpDirection = (clsEnums.enDirection)(row[col]);
                                tmpRd.Direction = tmpDirection;
                                break;
                            case "mmincreasingdirection":
                                tmpIncreasingDirection = (clsEnums.enDirection)(row[col]);
                                tmpRd.MMIncreasingDirection = tmpIncreasingDirection;
                                break;
                            case "recurringcongestionmmlocation":
                                tmpRecurringCongestionMMLocation = double.Parse(row[col].ToString());
                                tmpRd.RecurringCongestionMMLocation = tmpRecurringCongestionMMLocation;
                                break;
                        }
                    }
                    LogTxtMsg(txtINFLOConfigLog, "\t\t" + tmpRoadID + ", " + tmpRdName + ", " + tmpDirection + ", " +
                                                          tmpRdGrade + ", " + tmpBeginMM + ", " + tmpEndMM + ", " + 
                                                          tmpRecurringCongestionMMLocation + ", " + tmpIncreasingDirection);
                    RList.Add(tmpRd);
                }
            }
            return retValue;
        }
        private string GetRoadwayLinks(clsDatabase DB, ref List<clsRoadwayLink> RLList)
        {
            string retValue = string.Empty;
            string sqlQuery = string.Empty;

            DataSet RoadwayLinksDataSet = new DataSet("RoadwayLinks");
            sqlQuery = "Select * from Configuration_RoadwayLinks where RoadwayID='" + Roadway.Identifier + "' and BeginMM>=" + Roadway.BeginMM + " and EndMM<=" + Roadway.EndMM;

            retValue = string.Empty;
            retValue = DB.FillDataSet(sqlQuery, ref RoadwayLinksDataSet);
            try
            {
                FillDataSetLog.WriteLine(DateTime.Now + ",GetRoadwayLinks," + RoadwayLinksDataSet.Tables[0].Rows.Count);
            }
            catch (Exception exc)
            {
                retValue = "Error reading roadway link data: " + exc.Message;
            }
            if (retValue.Length > 0)
            {
                return retValue;
            }

            //display roadway link information if it already exists
            LogTxtMsg(txtINFLOConfigLog, "\tAvailable Roadway links: ");

            if (RoadwayLinksDataSet.Tables[0].Rows.Count > 0)
            {
                foreach (DataRow row in RoadwayLinksDataSet.Tables[0].Rows)
                {
                    clsRoadwayLink tmpRLL = new clsRoadwayLink();
                    string RoadID = string.Empty;
                    string RoadLinkID = string.Empty;
                    double BeginMM = 0;
                    double EndMM = 0;
                    string Dir = string.Empty;
                    int NumLanes = 0;
                    int NumDetectorStations = 0;
                    string DetectorStations = string.Empty;
                    foreach (DataColumn col in RoadwayLinksDataSet.Tables[0].Columns)
                    {
                        switch (col.ColumnName.ToString().ToLower())
                        {
                            case "roadwayid":
                                RoadID = row[col].ToString();
                                tmpRLL.RoadwayID = int.Parse(RoadID);
                                break;
                            case "linkid":
                                RoadLinkID = row[col].ToString();
                                tmpRLL.Identifier = int.Parse(RoadLinkID);
                                break;
                            case "beginmm":
                                BeginMM = double.Parse(row[col].ToString());
                                tmpRLL.BeginMM = BeginMM;
                                break;
                            case "endmm":
                                EndMM = double.Parse(row[col].ToString());
                                tmpRLL.EndMM = EndMM;
                                break;
                            case "numberlanes":
                                NumLanes = int.Parse(row[col].ToString());
                                tmpRLL.NumberLanes = NumLanes;
                                break;
                            case "numberdetectorstations":
                                NumDetectorStations = int.Parse(row[col].ToString());
                                tmpRLL.NumberDetectionStations = NumDetectorStations;
                                break;
                            case "detectorstations":
                                DetectorStations = row[col].ToString();
                                tmpRLL.DetectionStations = DetectorStations;
                                break;
                        }
                    }
                    LogTxtMsg(txtINFLOConfigLog, "\t\t" + RoadID + ", " + RoadLinkID + ", " + BeginMM + ", " + EndMM + ", " + NumLanes + ", " + NumDetectorStations + ", " + DetectorStations);
                    tmpRLL.Direction = Roadway.Direction;
                    RLList.Add(tmpRLL);
                }
            }
            return retValue;
        }
        private string GetRoadwaySubLinks(clsDatabase DB, ref List<clsRoadwaySubLink> RSLList)
        {
            string retValue = string.Empty;
            string sqlQuery = string.Empty;

            DataSet RoadwaySubLinksDataSet = new DataSet("RoadwaySubLinks");
            sqlQuery = "Select * from Configuration_RoadwaySubLinks where RoadwayID='" + Roadway.Identifier + "' and BeginMM>=" + Roadway.BeginMM + " and EndMM<=" + Roadway.EndMM;

            retValue = string.Empty;
            retValue = DB.FillDataSet(sqlQuery, ref RoadwaySubLinksDataSet);
            FillDataSetLog.WriteLine(DateTime.Now + ",GetRoadwaySubLinks," + RoadwaySubLinksDataSet.Tables[0].Rows.Count);
            if (retValue.Length > 0)
            {
                return retValue;
            }

            //display roadway sublink information if it already exists
            LogTxtMsg(txtINFLOConfigLog, "\tAvailable Roadway sublinks: ");

            if (RoadwaySubLinksDataSet.Tables[0].Rows.Count > 0)
            {
                foreach (DataRow row in RoadwaySubLinksDataSet.Tables[0].Rows)
                {
                    clsRoadwaySubLink tmpRSL = new clsRoadwaySubLink();
                    string RoadID = string.Empty;
                    string RoadSubLinkID = string.Empty;
                    double BeginMM = 0;
                    double EndMM = 0;
                    string Dir = string.Empty;
                    int NumLanes = 0;
                    foreach (DataColumn col in RoadwaySubLinksDataSet.Tables[0].Columns)
                    {
                        switch (col.ColumnName.ToString().ToLower())
                        {
                            case "roadwayid":
                                RoadID = row[col].ToString();
                                if (row[col].ToString().Length > 0)
                                {
                                    tmpRSL.RoadwayID = int.Parse(RoadID);
                                }
                                break;
                            case "sublinkid":
                                RoadSubLinkID = row[col].ToString();
                                tmpRSL.Identifier = int.Parse(RoadSubLinkID);
                                break;
                            case "beginmm":
                                if (row[col].ToString().Length > 0)
                                {
                                    BeginMM = double.Parse(row[col].ToString());
                                    tmpRSL.BeginMM = BeginMM;
                                }
                                break;
                            case "endmm":
                                if (row[col].ToString().Length > 0)
                                {
                                    EndMM = double.Parse(row[col].ToString());
                                    tmpRSL.EndMM = EndMM;
                                }
                                break;
                            case "numberlanes":
                                if (row[col].ToString().Length > 0)
                                {
                                    NumLanes = int.Parse(row[col].ToString());
                                    tmpRSL.NumberLanes = NumLanes;
                                }
                                break;
                            case "direction":
                                Dir = row[col].ToString();
                                tmpRSL.Direction = clsEnums.GetDirIndexFromString(Dir);
                                break;
                        }
                    }
                    LogTxtMsg(txtINFLOConfigLog, "\t\t" + RoadID + ", " + RoadSubLinkID + ", " + BeginMM + ", " + EndMM + ", " + NumLanes + ", " + Dir);
                    RSLList.Add(tmpRSL);
                }
            }
            return retValue;
        }
        private string GetDetectorStations(clsDatabase DB, ref List<clsDetectorStation> DSList)
        {
            string retValue = string.Empty;
            string sqlQuery = string.Empty;

            DataSet DetectorStationsDataSet = new DataSet("DetectorStations");
            sqlQuery = "Select * from Configuration_TSSDetectorStation";

            retValue = string.Empty;
            retValue = DB.FillDataSet(sqlQuery, ref DetectorStationsDataSet);
            FillDataSetLog.WriteLine(DateTime.Now + ",GetDetectorStations," + DetectorStationsDataSet.Tables[0].Rows.Count);
            if (retValue.Length > 0)
            {
                return retValue;
            }

            //display Detector station information 
            LogTxtMsg(txtINFLOConfigLog, "\tAvailable detector stations: ");

            if (DetectorStationsDataSet.Tables[0].Rows.Count > 0)
            {
                foreach (DataRow row in DetectorStationsDataSet.Tables[0].Rows)
                {
                    clsDetectorStation tmpDS = new clsDetectorStation();
                    string RoadLinkID = string.Empty;
                    string DSID = string.Empty;
                    double MMLocation = 0;
                    int NumDetectionZones = 0;
                    string DetectionZones = string.Empty;
                    foreach (DataColumn col in DetectorStationsDataSet.Tables[0].Columns)
                    {
                        switch (col.ColumnName.ToString().ToLower())
                        {
                            case "linkid":
                                RoadLinkID = row[col].ToString();
                                tmpDS.LinkIdentifier = int.Parse(RoadLinkID);
                                break;
                            case "dsid":
                                DSID = row[col].ToString();
                                tmpDS.Identifier = int.Parse(DSID);
                                break;
                            case "mmlocation":
                                MMLocation = double.Parse(row[col].ToString());
                                tmpDS.MMLocation = MMLocation;
                                break;
                            case "numberdetectionzones":
                                NumDetectionZones = int.Parse(row[col].ToString());
                                tmpDS.NumberDetectionZones = NumDetectionZones;
                                break;
                            case "detectionzones":
                                DetectionZones = row[col].ToString();
                                tmpDS.DetectionZones = DetectionZones;
                                break;
                        }
                    }
                    LogTxtMsg(txtINFLOConfigLog, "\t\t" + RoadLinkID + ", " + DSID + ", " + MMLocation + ", " + NumDetectionZones + ", " + DetectionZones);
                    DSList.Add(tmpDS);
                }
            }
            return retValue;
        }
        private string GetDetectionZones(clsDatabase DB, ref List<clsDetectionZone> DZList)
        {
            string retValue = string.Empty;
            string sqlQuery = string.Empty;
            string CurrentSection = string.Empty;

            try
            {
                CurrentSection = "Get DZ Info from database";

                DataSet DetectionZonesDataSet = new DataSet("DetectionZones");
                sqlQuery = "Select * from Configuration_TSSDetectionZone";

                retValue = string.Empty;
                retValue = DB.FillDataSet(sqlQuery, ref DetectionZonesDataSet);
                FillDataSetLog.WriteLine(DateTime.Now + ",GetDetectionZones," + DetectionZonesDataSet.Tables[0].Rows.Count);
                if (retValue.Length > 0)
                {
                    return retValue;
                }

                //display Detector station information 
                LogTxtMsg(txtINFLOConfigLog, "\tAvailable detection zones: ");
                int i = 0;
                if (DetectionZonesDataSet.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow row in DetectionZonesDataSet.Tables[0].Rows)
                    {
                        clsDetectionZone tmpDZ = new clsDetectionZone();
                        string DSID = string.Empty;
                        string DZID = string.Empty;
                        double MMLocation = 0;
                        int LaneNo = 0;
                        string DZType = string.Empty;
                        string LaneType = string.Empty;
                        string LaneDesc = string.Empty;
                        string DataType = string.Empty;
                        clsEnums.enDirection Direction = clsEnums.enDirection.NA;
                        CurrentSection = "Row: " + i;

                        foreach (DataColumn col in DetectionZonesDataSet.Tables[0].Columns)
                        {
                            switch (col.ColumnName.ToString().ToLower())
                            {
                                case "dsid":
                                    CurrentSection = "Row: " + i + "\t Col: " +col.ColumnName + "\tValue: " + row[col].ToString();
                                    if (row[col].ToString().Length > 0)
                                    {
                                        DSID = row[col].ToString();
                                        tmpDZ.DSIdentifier = int.Parse(DSID);
                                    }
                                    break;
                                case "dzid":
                                    CurrentSection = "Row: " + i + "\t Col: " +col.ColumnName + "\tValue: " + row[col].ToString();
                                    if (row[col].ToString().Length > 0)
                                    {
                                        DZID = row[col].ToString();
                                        tmpDZ.Identifier = int.Parse(DZID);
                                    }
                                    break;
                                case "mmlocation":
                                    CurrentSection = "Row: " + i + "\t Col: " +col.ColumnName + "\tValue: " + row[col].ToString();
                                    if (row[col].ToString().Length > 0)
                                    {
                                        MMLocation = double.Parse(row[col].ToString());
                                        tmpDZ.MMLocation = MMLocation;
                                    }
                                    break;
                                case "LaneNumber":
                                    CurrentSection = "Row: " + i + "\t Col: " +col.ColumnName + "\tValue: " + row[col].ToString();
                                    if (row[col].ToString().Length > 0)
                                    {
                                        LaneNo = int.Parse(row[col].ToString());
                                        tmpDZ.LaneNo = LaneNo;
                                    }
                                    break;
                                case "dztype":
                                    CurrentSection = "Row: " + i + "\t Col: " +col.ColumnName + "\tValue: " + row[col].ToString();
                                    DZType = row[col].ToString();
                                    tmpDZ.Type = DZType;
                                    break;
                                case "lanetype":
                                    CurrentSection = "Row: " + i + "\t Col: " +col.ColumnName + "\tValue: " + row[col].ToString();
                                    LaneType = row[col].ToString();
                                    tmpDZ.LaneType = LaneType;
                                    break;
                                case "lanedescription":
                                    CurrentSection = "Row: " + i + "\t Col: " +col.ColumnName + "\tValue: " + row[col].ToString();
                                    LaneDesc = row[col].ToString();
                                    tmpDZ.LaneDesc = LaneDesc;
                                    break;
                                case "datatype":
                                    CurrentSection = "Row: " + i + "\t Col: " +col.ColumnName + "\tValue: " + row[col].ToString();
                                    DataType = row[col].ToString();
                                    tmpDZ.DataType = DataType;
                                    break;
                                case "direction":
                                    CurrentSection = "Row: " + i + "\t Col: " +col.ColumnName + "\tValue: " + row[col].ToString();
                                    Direction = clsEnums.GetDirIndexFromString(row[col].ToString());
                                    tmpDZ.Direction = Direction;
                                    break;
                            }
                        }
                        LogTxtMsg(txtINFLOConfigLog, "\t\t" + DSID + ", " + DZID + ", " + DZType + ", " + DataType + ", " +
                                                              MMLocation + ", " + LaneNo + ", " + LaneType + ", " + LaneDesc + ", " + Direction);
                        DZList.Add(tmpDZ);
                    }
                }
            }
            catch (Exception ex)
            {
                retValue = "\tError in getting the Detection zones info from the INFLO database. Error Hint: Current Section: " + CurrentSection + "\r\n\t\t" + ex.Message;
                return retValue;
            }
            return retValue;
        }

        private double GetMinimumSpeed(double TSSAvgSpeed, double WRTMSpeed, double CVAvgSpeed)
        {
            double[] spdArray = new double[3];
            int Indx = 0;
            double MinSpeed = 0;

            if (TSSAvgSpeed > 0)
            {
                spdArray[Indx++] = TSSAvgSpeed;
            }
            else
            {
                spdArray[Indx++] = clsGlobalVars.MaximumDisplaySpeed;
            }


            if (WRTMSpeed > 0)
            {
                spdArray[Indx++] = WRTMSpeed;
            }
            else 
            {
                spdArray[Indx++] = clsGlobalVars.MaximumDisplaySpeed;
            }
            if (CVAvgSpeed > 0)
            {
                spdArray[Indx++] = CVAvgSpeed;
            }
            else
            {
                spdArray[Indx++] = clsGlobalVars.MaximumDisplaySpeed;
            }
            if (Indx > 0)
            {
                MinSpeed = spdArray.Min(); 
            }
            return MinSpeed;
        }
        private void DetermineBOQ_Demo(clsRoadwaySubLink QuedSubLink, clsRoadwayLink QuedLink, clsRoadway Roadway)
        {
            LogTxtMsg(txtINFLOLog, "\t" + DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + ":" + DateTime.Now.Millisecond +
                                     "\tReconcile CV BOQ MM and TSS BOQ MM locations:");
            LogTxtMsg(txtINFLOLog, "\t\tCurr  CV BOQ MM: " + clsGlobalVars.CVBOQMMLocation.ToString("0.0") + "\t CV BOQ Speed: " + clsGlobalVars.CVBOQSublinkSpeed.ToString("0"));
            LogTxtMsg(txtINFLOLog, "\t\tCurr TSS BOQ MM: " + clsGlobalVars.InfrastructureBOQMMLocation.ToString("0.0") + "\t TSS BOQ Speed: " + clsGlobalVars.InfrastructureBOQMMLocation.ToString("0"));

            #region "Reconcile CV BOQ MM Location and TSS data BOQ MM Location"

            clsGlobalVars.QueueSource = clsEnums.enQueueSource.NA;
            if (Roadway.Direction == Roadway.MMIncreasingDirection)
            {
                #region "Mile marker increasing direction = roadway direction"
                if ((clsGlobalVars.InfrastructureBOQMMLocation != -1) && (clsGlobalVars.CVBOQMMLocation != -1))
                {
                    if (clsGlobalVars.CVBOQMMLocation <= clsGlobalVars.InfrastructureBOQMMLocation)
                    {
                        clsGlobalVars.QueueSource = clsEnums.enQueueSource.CV;
                    }
                    else
                    {
                        clsGlobalVars.QueueSource = clsEnums.enQueueSource.TSS;
                    }
                }
                else if (clsGlobalVars.CVBOQMMLocation != -1)
                {
                    clsGlobalVars.QueueSource = clsEnums.enQueueSource.CV;
                }
                else if (clsGlobalVars.InfrastructureBOQMMLocation != -1)
                {
                    clsGlobalVars.QueueSource = clsEnums.enQueueSource.TSS;
                }
                else
                {
                    clsGlobalVars.QueueSource = clsEnums.enQueueSource.NA;
                }
                #endregion
            }
            else if (Roadway.Direction != Roadway.MMIncreasingDirection)
            {
                #region "Mile marker increasing direction = opposite to roadway direction
                if ((clsGlobalVars.InfrastructureBOQMMLocation != -1) && (clsGlobalVars.CVBOQMMLocation != -1))
                {
                    if (clsGlobalVars.CVBOQMMLocation >= clsGlobalVars.InfrastructureBOQMMLocation)
                    {
                        clsGlobalVars.QueueSource = clsEnums.enQueueSource.CV;
                    }
                    else
                    {
                        clsGlobalVars.QueueSource = clsEnums.enQueueSource.TSS;
                    }
                }
                else if (clsGlobalVars.CVBOQMMLocation != -1)
                {
                    clsGlobalVars.QueueSource = clsEnums.enQueueSource.CV;
                }
                else if (clsGlobalVars.InfrastructureBOQMMLocation != -1)
                {
                    clsGlobalVars.QueueSource = clsEnums.enQueueSource.TSS;
                }
                else
                {
                    clsGlobalVars.QueueSource = clsEnums.enQueueSource.NA;
                }
                #endregion
            }

            if (clsGlobalVars.QueueSource == clsEnums.enQueueSource.CV)
            {
                clsGlobalVars.BOQMMLocation = clsGlobalVars.CVBOQMMLocation;
                clsGlobalVars.BOQTime = DateTime.Now;
                clsGlobalVars.BOQSpeed = clsGlobalVars.CVBOQSublinkSpeed;
            }
            else if (clsGlobalVars.QueueSource == clsEnums.enQueueSource.TSS)
            {
                clsGlobalVars.BOQMMLocation = clsGlobalVars.InfrastructureBOQMMLocation;
                clsGlobalVars.BOQTime = DateTime.Now;
                clsGlobalVars.BOQSpeed = clsGlobalVars.InfrastructureBOQLinkSpeed;
            }
            else if (clsGlobalVars.QueueSource == clsEnums.enQueueSource.NA)
            {
                clsGlobalVars.BOQMMLocation = -1;
                clsGlobalVars.BOQTime = DateTime.Now;
                clsGlobalVars.BOQSpeed = 0;
            }
            #endregion

            double tmpBOQMMLocationChange = 0;
            clsGlobalVars.QueueChange = clsEnums.enQueueCahnge.Same;
            clsGlobalVars.QueueRate = 0;

            #region "Determine queue rate of change"
            TimeSpan span = clsGlobalVars.BOQTime.Subtract(clsGlobalVars.PrevBOQTime);
            if (Roadway.Direction == Roadway.MMIncreasingDirection)
            {
                if ((clsGlobalVars.PrevBOQMMLocation != -1) && (clsGlobalVars.BOQMMLocation != -1))
                {
                    tmpBOQMMLocationChange = clsGlobalVars.BOQMMLocation - clsGlobalVars.PrevBOQMMLocation;
                    if (chkFilterQueues.Checked == true)
                    {
                        if ((Math.Abs(tmpBOQMMLocationChange) > (2 * clsGlobalVars.LinkLength)))
                        {
                            tmpBOQMMLocationChange = 0;
                            clsGlobalVars.BOQMMLocation = clsGlobalVars.PrevBOQMMLocation;
                        }
                    }
                    clsGlobalVars.QueueRate = ((Math.Abs(tmpBOQMMLocationChange) * 3600) / span.TotalSeconds);
                    if (tmpBOQMMLocationChange < 0)
                    {
                        clsGlobalVars.QueueChange = clsEnums.enQueueCahnge.Growing;
                    }
                    else if (tmpBOQMMLocationChange > 0)
                    {
                        clsGlobalVars.QueueChange = clsEnums.enQueueCahnge.Dissipating;
                    }
                    else if (tmpBOQMMLocationChange == 0)
                    {
                        clsGlobalVars.QueueChange = clsEnums.enQueueCahnge.Same;
                    }
                }
                else if ((clsGlobalVars.PrevBOQMMLocation != -1) && (clsGlobalVars.BOQMMLocation == -1))
                {
                    tmpBOQMMLocationChange = (Roadway.RecurringCongestionMMLocation - clsGlobalVars.PrevBOQMMLocation);
                    if (chkFilterQueues.Checked == true)
                    {
                        if ((Math.Abs(tmpBOQMMLocationChange) > (2 * clsGlobalVars.LinkLength)))
                        {
                            tmpBOQMMLocationChange = 0;
                            clsGlobalVars.BOQMMLocation = clsGlobalVars.PrevBOQMMLocation;
                        }
                    }

                    clsGlobalVars.QueueRate = ((Math.Abs(tmpBOQMMLocationChange) * 3600) / span.TotalSeconds);
                    if (tmpBOQMMLocationChange > 0)
                    {
                        clsGlobalVars.QueueChange = clsEnums.enQueueCahnge.Dissipating;
                    }
                    else if (tmpBOQMMLocationChange == 0)
                    {
                        clsGlobalVars.QueueChange = clsEnums.enQueueCahnge.Same;
                    }
                    //clsGlobalVars.QueueChange = clsEnums.enQueueCahnge.Dissipating;
                }
                else if ((clsGlobalVars.PrevBOQMMLocation == -1) && (clsGlobalVars.BOQMMLocation != -1))
                {
                    tmpBOQMMLocationChange = (Roadway.RecurringCongestionMMLocation - clsGlobalVars.BOQMMLocation);
                    if (chkFilterQueues.Checked == true)
                    {
                        if ((Math.Abs(tmpBOQMMLocationChange) > (2 * clsGlobalVars.LinkLength)))
                        {
                            tmpBOQMMLocationChange = 0;
                            clsGlobalVars.BOQMMLocation = clsGlobalVars.PrevBOQMMLocation;
                        }
                    }
                    clsGlobalVars.QueueRate = ((Math.Abs(tmpBOQMMLocationChange) * 3600) / span.TotalSeconds);
                    if (tmpBOQMMLocationChange > 0)
                    {
                        clsGlobalVars.QueueChange = clsEnums.enQueueCahnge.Growing;
                    }
                    else if (tmpBOQMMLocationChange == 0)
                    {
                        clsGlobalVars.QueueChange = clsEnums.enQueueCahnge.Same;
                    }
                    //clsGlobalVars.QueueChange = clsEnums.enQueueCahnge.Growing;
                }
                else
                {
                    clsGlobalVars.QueueChange = clsEnums.enQueueCahnge.Same;
                    clsGlobalVars.QueueRate = 0;
                }

                string tmpSublinkData = string.Empty;
                if (QueuedSubLink != null)
                {
                    tmpSublinkData = QueuedSubLink.TotalNumberCVs + "," + QueuedSubLink.PrevTotalNumberCVs + "," + QueuedSubLink.VolumeDiff + "," + QueuedSubLink.FlowRate.ToString("0.00") + "," +
                                     QueuedSubLink.Density.ToString("0.00") + "," + QueuedSubLink.PrevDensity.ToString("0.00") + "," + QueuedSubLink.DensityDiff + "," + QueuedSubLink.ShockWaveRate.ToString("0.00");
                }

                if (Stopped == false)
                {
                    QueueLog.WriteLine(DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + "::" + DateTime.Now.Millisecond + "," +
                                       clsGlobalVars.PrevBOQMMLocation.ToString("0.0") + "," +
                                       clsGlobalVars.PrevBOQTime.Hour + ":" + clsGlobalVars.PrevBOQTime.Minute + ":" + clsGlobalVars.PrevBOQTime.Second + ":" + clsGlobalVars.PrevBOQTime.Millisecond + "," +
                                       clsGlobalVars.BOQMMLocation.ToString("0.0") + "," +
                                       clsGlobalVars.BOQTime.Hour + ":" + clsGlobalVars.BOQTime.Minute + ":" + clsGlobalVars.BOQTime.Second + ":" + clsGlobalVars.BOQTime.Millisecond + "," +
                                       tmpBOQMMLocationChange + "," + span.TotalSeconds.ToString("0") + "," + clsGlobalVars.QueueRate.ToString("0") + "," + clsGlobalVars.QueueChange.ToString() + "," +
                                       clsGlobalVars.QueueSource.ToString() + "," + clsGlobalVars.CVBOQMMLocation + "," + clsGlobalVars.InfrastructureBOQMMLocation + "," + tmpSublinkData);
                }
                LogTxtMsg(txtINFLOLog, "\t\tPrev BOQ MM: " + clsGlobalVars.PrevBOQMMLocation.ToString("0.0") + "\t\tTime: " + clsGlobalVars.PrevBOQTime);
                if (clsGlobalVars.PrevBOQMMLocation != clsGlobalVars.BOQMMLocation)
                {
                    LogTxtMsg(txtINFLOLog, "\t\tCurr BOQ MM:  " + clsGlobalVars.BOQMMLocation + "\t\tTime: " + clsGlobalVars.BOQTime + "\tSource: " + clsGlobalVars.QueueSource.ToString());
                    clsGlobalVars.PrevBOQMMLocation = clsGlobalVars.BOQMMLocation;
                    clsGlobalVars.PrevBOQTime = clsGlobalVars.BOQTime;
                }
                if (clsGlobalVars.BOQMMLocation != -1)
                {
                    InsertQueueInfoIntoINFLODatabase(clsGlobalVars.BOQMMLocation, clsGlobalVars.BOQTime, clsGlobalVars.QueueRate, clsGlobalVars.QueueChange, clsGlobalVars.QueueSource, clsGlobalVars.QueueSpeed, Roadway);
                }
                //else
                //{
                //    InsertQueueInfoIntoINFLODatabase(0, clsGlobalVars.BOQTime, clsGlobalVars.QueueRate, clsGlobalVars.QueueChange, clsGlobalVars.QueueSource, clsGlobalVars.QueueSpeed, Roadway);
                //}
            }
            else if (Roadway.Direction != Roadway.MMIncreasingDirection)
            {
                if ((clsGlobalVars.PrevBOQMMLocation != -1) && (clsGlobalVars.BOQMMLocation != -1))
                {
                    tmpBOQMMLocationChange = clsGlobalVars.BOQMMLocation - clsGlobalVars.PrevBOQMMLocation;
                    if (chkFilterQueues.Checked == true)
                    {
                        if (Math.Abs(tmpBOQMMLocationChange) > (2 * clsGlobalVars.LinkLength))
                        {
                            tmpBOQMMLocationChange = 0;
                            clsGlobalVars.BOQMMLocation = clsGlobalVars.PrevBOQMMLocation;
                        }
                    }
                    clsGlobalVars.QueueRate = ((Math.Abs(tmpBOQMMLocationChange) * 3600) / span.TotalSeconds);
                    if (tmpBOQMMLocationChange < 0)
                    {
                        clsGlobalVars.QueueChange = clsEnums.enQueueCahnge.Dissipating;
                    }
                    else if (tmpBOQMMLocationChange > 0)
                    {
                        clsGlobalVars.QueueChange = clsEnums.enQueueCahnge.Growing;
                    }
                    else if (tmpBOQMMLocationChange == 0)
                    {
                        clsGlobalVars.QueueChange = clsEnums.enQueueCahnge.Same;
                    }
                }
                else if ((clsGlobalVars.PrevBOQMMLocation != -1) && (clsGlobalVars.BOQMMLocation == -1))
                {
                    tmpBOQMMLocationChange = (clsGlobalVars.PrevBOQMMLocation - Roadway.RecurringCongestionMMLocation);
                    if (chkFilterQueues.Checked == true)
                    {
                        if (Math.Abs(tmpBOQMMLocationChange) > (2 * clsGlobalVars.LinkLength))
                        {
                            tmpBOQMMLocationChange = 0;
                            clsGlobalVars.BOQMMLocation = clsGlobalVars.PrevBOQMMLocation;
                        }
                    }
                    clsGlobalVars.QueueRate = ((Math.Abs(tmpBOQMMLocationChange) * 3600) / span.TotalSeconds);
                    if (tmpBOQMMLocationChange > 0)
                    {
                        clsGlobalVars.QueueChange = clsEnums.enQueueCahnge.Growing;
                    }
                    else if (tmpBOQMMLocationChange == 0)
                    {
                        clsGlobalVars.QueueChange = clsEnums.enQueueCahnge.Same;
                    }
                    //clsGlobalVars.QueueChange = clsEnums.enQueueCahnge.Dissipating;
                }
                else if ((clsGlobalVars.PrevBOQMMLocation == -1) && (clsGlobalVars.BOQMMLocation != -1))
                {
                    tmpBOQMMLocationChange = (clsGlobalVars.BOQMMLocation - Roadway.RecurringCongestionMMLocation);
                    if (chkFilterQueues.Checked == true)
                    {
                        if (Math.Abs(tmpBOQMMLocationChange) > (2 * clsGlobalVars.LinkLength))
                        {
                            tmpBOQMMLocationChange = 0;
                            clsGlobalVars.BOQMMLocation = clsGlobalVars.PrevBOQMMLocation;
                        }
                    }
                    clsGlobalVars.QueueRate = ((Math.Abs(tmpBOQMMLocationChange) * 3600) / span.TotalSeconds);
                    if (tmpBOQMMLocationChange > 0)
                    {
                        clsGlobalVars.QueueChange = clsEnums.enQueueCahnge.Growing;
                    }
                    else if (tmpBOQMMLocationChange == 0)
                    {
                        clsGlobalVars.QueueChange = clsEnums.enQueueCahnge.Same;
                    }
                    //clsGlobalVars.QueueChange = clsEnums.enQueueCahnge.Growing;
                }
                else
                {
                    clsGlobalVars.QueueChange = clsEnums.enQueueCahnge.Same;
                    clsGlobalVars.QueueRate = 0;
                }

                string tmpSublinkData = string.Empty;
                if (QueuedSubLink != null)
                {
                    tmpSublinkData = QueuedSubLink.TotalNumberCVs + "," + QueuedSubLink.PrevTotalNumberCVs + "," + QueuedSubLink.VolumeDiff + "," + QueuedSubLink.FlowRate.ToString("0.00") + "," +
                                     QueuedSubLink.Density.ToString("0.00") + "," + QueuedSubLink.PrevDensity.ToString("0.00") + "," + QueuedSubLink.DensityDiff + "," + QueuedSubLink.ShockWaveRate.ToString("0.00");
                }

                QueueLog.WriteLine(DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + "::" + DateTime.Now.Millisecond + "," +
                                   clsGlobalVars.PrevBOQMMLocation.ToString("0.0") + "," +
                                   clsGlobalVars.PrevBOQTime.Hour + ":" + clsGlobalVars.PrevBOQTime.Minute + ":" + clsGlobalVars.PrevBOQTime.Second + ":" + clsGlobalVars.PrevBOQTime.Millisecond + "," +
                                   clsGlobalVars.BOQMMLocation.ToString("0.0") + "," +
                                   clsGlobalVars.BOQTime.Hour + ":" + clsGlobalVars.BOQTime.Minute + ":" + clsGlobalVars.BOQTime.Second + ":" + clsGlobalVars.BOQTime.Millisecond + "," +
                                   tmpBOQMMLocationChange + "," + span.TotalSeconds.ToString("0") + "," + clsGlobalVars.QueueRate.ToString("0") + "," + clsGlobalVars.QueueChange.ToString() + "," +
                                   clsGlobalVars.QueueSource.ToString() + "," + clsGlobalVars.CVBOQMMLocation + "," + clsGlobalVars.InfrastructureBOQMMLocation + "," + tmpSublinkData);

                LogTxtMsg(txtINFLOLog, "\t\tPrev BOQ MM: " + clsGlobalVars.PrevBOQMMLocation.ToString("0.0") + "\t\tTime: " + clsGlobalVars.PrevBOQTime);
                if (clsGlobalVars.PrevBOQMMLocation != clsGlobalVars.BOQMMLocation)
                {
                    LogTxtMsg(txtINFLOLog, "\t\tCurr BOQ MM:  " + clsGlobalVars.BOQMMLocation + "\t\tTime: " + clsGlobalVars.BOQTime + "\tSource: " + clsGlobalVars.QueueSource.ToString());
                    clsGlobalVars.PrevBOQMMLocation = clsGlobalVars.BOQMMLocation;
                    clsGlobalVars.PrevBOQTime = clsGlobalVars.BOQTime;
                }
                if (clsGlobalVars.BOQMMLocation != -1)
                {
                    InsertQueueInfoIntoINFLODatabase(clsGlobalVars.BOQMMLocation, clsGlobalVars.BOQTime, clsGlobalVars.QueueRate, clsGlobalVars.QueueChange, clsGlobalVars.QueueSource, clsGlobalVars.QueueSpeed, Roadway);
                }
                //else
                //{
                //    InsertQueueInfoIntoINFLODatabase(0, clsGlobalVars.BOQTime, clsGlobalVars.QueueRate, clsGlobalVars.QueueChange, clsGlobalVars.QueueSource, clsGlobalVars.QueueSpeed, Roadway);
                //}
            }
            #endregion

            //LogTxtMsg(txtTMEQWARNLog, "\r\n\t\tPrev BOQ MM location: " + clsGlobalVars.PrevBOQMMLocation + "\tTime: " + clsGlobalVars.PrevBOQTime);
            //CVDataProcessor.WriteLine("\tPrev BOQ MM location: " + clsGlobalVars.PrevBOQMMLocation + "\tTime: " + clsGlobalVars.PrevBOQTime);

            //LogTxtMsg(txtINFLOLog, "\t\tCurr BOQ MM location: " + clsGlobalVars.BOQMMLocation + "\tTime: " + clsGlobalVars.BOQTime + "\tSource: " + clsGlobalVars.QueueSource.ToString());
            //CVDataProcessor.WriteLine("\tCurr BOQ MM location: " + clsGlobalVars.BOQMMLocation + "\tTime: " + clsGlobalVars.BOQTime + "\tSource: " + clsGlobalVars.QueueSource.ToString());

            //Current BOQ
            if (clsGlobalVars.BOQMMLocation == -1)
            {
                clsGlobalVars.QueueLength = 0;
                clsGlobalVars.QueueSpeed = 0;
                DisplayForm.txtBOQ.Text = "No Queue";
                DisplayForm.txtQueueLength.Text = "";
                DisplayForm.txtQueueGrowthRate.Text = "";
                DisplayForm.txtQueueSpeed.Text = "";

            }
            else
            {
                double TotalQuedSublinksCVSpeed = 0;
                double TotalQuedSublinksVolume = 0;
                foreach (clsRoadwaySubLink RSL in RSLList)
                {
                    if (Roadway.Direction == Roadway.MMIncreasingDirection)
                    {
                        if ((RSL.BeginMM >= clsGlobalVars.BOQMMLocation) && (RSL.BeginMM < Roadway.RecurringCongestionMMLocation))
                        {
                            TotalQuedSublinksCVSpeed = TotalQuedSublinksCVSpeed + RSL.TotalNumberCVs * RSL.CVAvgSpeed;
                            TotalQuedSublinksVolume = TotalQuedSublinksVolume + RSL.TotalNumberCVs;
                        }
                    }
                    else if (Roadway.Direction != Roadway.MMIncreasingDirection)
                    {
                        if ((RSL.BeginMM <= clsGlobalVars.BOQMMLocation) && (RSL.BeginMM > Roadway.RecurringCongestionMMLocation))
                        {
                            TotalQuedSublinksCVSpeed = TotalQuedSublinksCVSpeed + RSL.TotalNumberCVs * RSL.CVAvgSpeed;
                            TotalQuedSublinksVolume = TotalQuedSublinksVolume + RSL.TotalNumberCVs;
                        }
                    }
                }
                clsGlobalVars.QueueSpeed = TotalQuedSublinksCVSpeed / TotalQuedSublinksVolume;
                clsGlobalVars.QueueLength = Roadway.RecurringCongestionMMLocation - clsGlobalVars.BOQMMLocation;
                DisplayForm.txtBOQ.Text = clsGlobalVars.BOQMMLocation.ToString();
                DisplayForm.txtQueueLength.Text = (Roadway.RecurringCongestionMMLocation - clsGlobalVars.BOQMMLocation).ToString("0.00");
                DisplayForm.txtQueueGrowthRate.Text = clsGlobalVars.QueueRate.ToString("0");
                DisplayForm.txtQueueSpeed.Text = clsGlobalVars.QueueSpeed.ToString("0");
            }
            LogTxtMsg(txtINFLOLog, "\t\tCurr BOQ rate of growth: " + clsGlobalVars.QueueRate.ToString("0.00") + "\tQueue direction: " + clsGlobalVars.QueueChange.ToString() + "\tQueue source: " + clsGlobalVars.QueueSource.ToString());
            //CVDataProcessor.WriteLine("\tCurr BOQ rate of growth: " + clsGlobalVars.QueueRate.ToString("0.00") + "\tQueue direction: " + clsGlobalVars.QueueChange.ToString() + "\tQueue source: " + clsGlobalVars.QueueSource.ToString());
        }
        private void DetermineBOQ(clsRoadwaySubLink QuedSubLink, clsRoadwayLink QuedLink, clsRoadway Roadway)
        {
            LogTxtMsg(txtINFLOLog, "\t" + DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + ":" + DateTime.Now.Millisecond +
                                     "\tReconcile CV BOQ MM and TSS BOQ MM locations:");
            LogTxtMsg(txtINFLOLog, "\r\n\t\tCurr  CV BOQ MM: " + clsGlobalVars.CVBOQMMLocation.ToString("0.0") + "\t CV BOQ Speed: " + clsGlobalVars.CVBOQSublinkSpeed.ToString("0"));
            LogTxtMsg(txtINFLOLog, "\t\tCurr TSS BOQ MM: " + clsGlobalVars.InfrastructureBOQMMLocation.ToString("0.0") + "\t TSS BOQ Speed: " + clsGlobalVars.InfrastructureBOQMMLocation.ToString("0"));

            #region "Reconcile CV BOQ MM Location and TSS data BOQ MM Location"

            clsEnums.enQueueSource tmpQueueSource = clsEnums.enQueueSource.NA;
            double tmpBOQMMLocationChange = 0;

            double tmpQueDiff = 0;
            double tmpTSSQueDiff = 0;
            double tmpCVQueDiff = 0;
            clsEnums.enQueueCahnge tmpQueChangeDirection = clsEnums.enQueueCahnge.NA;

            //Mile marker increasing direction = roadway direction
            if (Roadway.Direction == Roadway.MMIncreasingDirection)
            {
                #region "Queue Exists"
                if (clsGlobalVars.BOQMMLocation != -1)
                {
                    tmpTSSQueDiff = clsGlobalVars.InfrastructureBOQMMLocation - clsGlobalVars.BOQMMLocation;
                    tmpCVQueDiff = clsGlobalVars.CVBOQMMLocation - clsGlobalVars.BOQMMLocation;
                    if ((clsGlobalVars.InfrastructureBOQMMLocation != -1) && (clsGlobalVars.CVBOQMMLocation != -1))
                    {
                        if (chkFilterQueues.Checked == true)
                        {
                            if ((Math.Abs(tmpTSSQueDiff) <= clsGlobalVars.LinkLength) && (Math.Abs(tmpCVQueDiff) <= clsGlobalVars.LinkLength))
                            {
                                if (tmpCVQueDiff <= tmpTSSQueDiff)
                                {
                                    tmpQueueSource = clsEnums.enQueueSource.CV;
                                    tmpBOQMMLocationChange = tmpCVQueDiff;
                                }
                                else
                                {
                                    tmpQueueSource = clsEnums.enQueueSource.TSS;
                                    tmpBOQMMLocationChange = tmpTSSQueDiff;
                                }
                            }
                            else if (Math.Abs(tmpTSSQueDiff) <= clsGlobalVars.LinkLength)
                            {
                                tmpQueueSource = clsEnums.enQueueSource.TSS;
                                tmpBOQMMLocationChange = tmpTSSQueDiff;
                            }
                            else if (Math.Abs(tmpCVQueDiff) <= clsGlobalVars.LinkLength)
                            {
                                tmpQueueSource = clsEnums.enQueueSource.CV;
                                tmpBOQMMLocationChange = tmpCVQueDiff;
                            }
                        }
                        else
                        {
                            if (tmpCVQueDiff <= tmpTSSQueDiff)
                            {
                                tmpQueueSource = clsEnums.enQueueSource.CV;
                                tmpBOQMMLocationChange = tmpCVQueDiff;
                            }
                            else
                            {
                                tmpQueueSource = clsEnums.enQueueSource.TSS;
                                tmpBOQMMLocationChange = tmpTSSQueDiff;
                            }
                        }
                    }
                    else if (clsGlobalVars.CVBOQMMLocation != -1)
                    {
                        if (chkFilterQueues.Checked == true)
                        {
                            if (Math.Abs(tmpCVQueDiff) <= clsGlobalVars.LinkLength)
                            {
                                tmpQueueSource = clsEnums.enQueueSource.CV;
                                tmpBOQMMLocationChange = tmpCVQueDiff;
                            }
                        }
                        else
                        {
                            tmpQueueSource = clsEnums.enQueueSource.CV;
                            tmpBOQMMLocationChange = tmpCVQueDiff;
                        }
                    }
                    else if (clsGlobalVars.InfrastructureBOQMMLocation != -1)
                    {
                        if (chkFilterQueues.Checked == true)
                        {
                            if (Math.Abs(tmpTSSQueDiff) <= clsGlobalVars.LinkLength)
                            {
                                tmpQueueSource = clsEnums.enQueueSource.TSS;
                                tmpBOQMMLocationChange = tmpTSSQueDiff;
                            }
                        }
                        else
                        {
                            tmpQueueSource = clsEnums.enQueueSource.TSS;
                            tmpBOQMMLocationChange = tmpTSSQueDiff;
                        }
                    }
                    else
                    {
                        tmpQueDiff = Roadway.RecurringCongestionMMLocation - clsGlobalVars.BOQMMLocation;
                        if (chkFilterQueues.Checked == true)
                        {
                            if (Math.Abs(tmpQueDiff) <= clsGlobalVars.LinkLength)
                            {
                                tmpQueueSource = clsEnums.enQueueSource.NA;
                                tmpBOQMMLocationChange = tmpQueDiff;
                            }
                        }
                        else
                        {
                            tmpQueueSource = clsEnums.enQueueSource.NA;
                            tmpBOQMMLocationChange = tmpQueDiff;
                        }
                    }
                    if (tmpBOQMMLocationChange < 0)
                    {
                        tmpQueChangeDirection = clsEnums.enQueueCahnge.Growing;
                    }
                    else if (tmpBOQMMLocationChange > 0)
                    {
                        tmpQueChangeDirection = clsEnums.enQueueCahnge.Dissipating;
                    }
                    else if (tmpBOQMMLocationChange == 0)
                    {
                        tmpQueChangeDirection = clsEnums.enQueueCahnge.Same;
                    }
                }
                #endregion
                #region "No Queue"
                else if (clsGlobalVars.BOQMMLocation == -1)
                {
                    tmpTSSQueDiff = clsGlobalVars.InfrastructureBOQMMLocation - Roadway.RecurringCongestionMMLocation;
                    tmpCVQueDiff = clsGlobalVars.CVBOQMMLocation - Roadway.RecurringCongestionMMLocation;
                    if ((clsGlobalVars.InfrastructureBOQMMLocation != -1) && (clsGlobalVars.CVBOQMMLocation != -1))
                    {
                        if (chkFilterQueues.Checked == true)
                        {
                            if ((Math.Abs(tmpTSSQueDiff) <= 2*clsGlobalVars.LinkLength) && (Math.Abs(tmpCVQueDiff) <= 2*clsGlobalVars.LinkLength))
                            {
                                if ((tmpCVQueDiff <= tmpTSSQueDiff) && (clsGlobalVars.CVBOQMMLocation <= Roadway.RecurringCongestionMMLocation))
                                {
                                    tmpQueueSource = clsEnums.enQueueSource.CV;
                                    tmpBOQMMLocationChange = tmpCVQueDiff;
                                }
                                else if (clsGlobalVars.InfrastructureBOQMMLocation <= Roadway.RecurringCongestionMMLocation)
                                {
                                    tmpQueueSource = clsEnums.enQueueSource.TSS;
                                    tmpBOQMMLocationChange = tmpTSSQueDiff;
                                }
                            }
                            else if ((Math.Abs(tmpTSSQueDiff) <= 2 * clsGlobalVars.LinkLength) && (clsGlobalVars.InfrastructureBOQMMLocation <= Roadway.RecurringCongestionMMLocation))
                            {
                                tmpQueueSource = clsEnums.enQueueSource.TSS;
                                tmpBOQMMLocationChange = tmpTSSQueDiff;
                            }
                            else if ((Math.Abs(tmpCVQueDiff) <= 2 * clsGlobalVars.LinkLength) && (clsGlobalVars.CVBOQMMLocation <= Roadway.RecurringCongestionMMLocation))
                            {
                                tmpQueueSource = clsEnums.enQueueSource.CV;
                                tmpBOQMMLocationChange = tmpCVQueDiff;
                            }
                        }
                        else
                        {
                            if ((tmpCVQueDiff <= tmpTSSQueDiff) && (clsGlobalVars.CVBOQMMLocation <= Roadway.RecurringCongestionMMLocation))
                            {
                                tmpQueueSource = clsEnums.enQueueSource.CV;
                                tmpBOQMMLocationChange = tmpCVQueDiff;
                            }
                            else if (clsGlobalVars.InfrastructureBOQMMLocation <= Roadway.RecurringCongestionMMLocation)
                            {
                                tmpQueueSource = clsEnums.enQueueSource.TSS;
                                tmpBOQMMLocationChange = tmpTSSQueDiff;
                            }
                        }
                    }
                    else if (clsGlobalVars.CVBOQMMLocation != -1)
                    {
                        if (chkFilterQueues.Checked == true)
                        {
                            if ((Math.Abs(tmpCVQueDiff) <= 2 * clsGlobalVars.LinkLength) && (clsGlobalVars.CVBOQMMLocation <= Roadway.RecurringCongestionMMLocation))
                            {
                                tmpQueueSource = clsEnums.enQueueSource.CV;
                                tmpBOQMMLocationChange = tmpCVQueDiff;
                            }
                        }
                        else if (clsGlobalVars.CVBOQMMLocation <= Roadway.RecurringCongestionMMLocation)
                        {
                            tmpQueueSource = clsEnums.enQueueSource.CV;
                            tmpBOQMMLocationChange = tmpCVQueDiff;
                        }
                    }
                    else if (clsGlobalVars.InfrastructureBOQMMLocation != -1)
                    {
                        if (chkFilterQueues.Checked == true)
                        {
                            if ((Math.Abs(tmpTSSQueDiff) <= 2 * clsGlobalVars.LinkLength) && (clsGlobalVars.InfrastructureBOQMMLocation <= Roadway.RecurringCongestionMMLocation))
                            {
                                tmpQueueSource = clsEnums.enQueueSource.TSS;
                                tmpBOQMMLocationChange = tmpTSSQueDiff;
                            }
                        }
                        else if (clsGlobalVars.InfrastructureBOQMMLocation <= Roadway.RecurringCongestionMMLocation)
                        {
                            tmpQueueSource = clsEnums.enQueueSource.TSS;
                            tmpBOQMMLocationChange = tmpTSSQueDiff;
                        }
                    }
                    else
                    {
                        tmpQueueSource = clsEnums.enQueueSource.NA;
                        tmpBOQMMLocationChange = 0;
                    }
                    if (tmpBOQMMLocationChange <= 0)
                    {
                        tmpQueChangeDirection = clsEnums.enQueueCahnge.Growing;
                    }
                    else if (tmpBOQMMLocationChange > 0)
                    {
                        tmpQueChangeDirection = clsEnums.enQueueCahnge.Dissipating;
                    }
                }
                #endregion
            }
            //Mile marker increasing direction = opposite to roadway direction
            else if (Roadway.Direction != Roadway.MMIncreasingDirection)
            {
                #region "Queue Exists"
                if (clsGlobalVars.BOQMMLocation != -1)
                {
                    tmpTSSQueDiff = clsGlobalVars.InfrastructureBOQMMLocation - clsGlobalVars.BOQMMLocation;
                    tmpCVQueDiff = clsGlobalVars.CVBOQMMLocation - clsGlobalVars.BOQMMLocation;
                    if ((clsGlobalVars.InfrastructureBOQMMLocation != -1) && (clsGlobalVars.CVBOQMMLocation != -1))
                    {
                        if (chkFilterQueues.Checked == true)
                        {
                            if ((Math.Abs(tmpTSSQueDiff) <= clsGlobalVars.LinkLength) && (Math.Abs(tmpCVQueDiff) <= clsGlobalVars.LinkLength))
                            {
                                if (tmpCVQueDiff >= tmpTSSQueDiff)
                                {
                                    tmpQueueSource = clsEnums.enQueueSource.CV;
                                    tmpBOQMMLocationChange = tmpCVQueDiff;
                                }
                                else
                                {
                                    tmpQueueSource = clsEnums.enQueueSource.TSS;
                                    tmpBOQMMLocationChange = tmpTSSQueDiff;
                                }
                            }
                            else if (Math.Abs(tmpTSSQueDiff) <= clsGlobalVars.LinkLength)
                            {
                                tmpQueueSource = clsEnums.enQueueSource.TSS;
                                tmpBOQMMLocationChange = tmpTSSQueDiff;
                            }
                            else if (Math.Abs(tmpCVQueDiff) <= clsGlobalVars.LinkLength)
                            {
                                tmpQueueSource = clsEnums.enQueueSource.CV;
                                tmpBOQMMLocationChange = tmpCVQueDiff;
                            }
                        }
                        else
                        {
                            if (tmpCVQueDiff >= tmpTSSQueDiff)
                            {
                                tmpQueueSource = clsEnums.enQueueSource.CV;
                                tmpBOQMMLocationChange = tmpCVQueDiff;
                            }
                            else
                            {
                                tmpQueueSource = clsEnums.enQueueSource.TSS;
                                tmpBOQMMLocationChange = tmpTSSQueDiff;
                            }
                        }
                    }
                    else if (clsGlobalVars.CVBOQMMLocation != -1)
                    {
                        if (chkFilterQueues.Checked == true)
                        {
                            if (Math.Abs(tmpCVQueDiff) <= clsGlobalVars.LinkLength)
                            {
                                tmpQueueSource = clsEnums.enQueueSource.CV;
                                tmpBOQMMLocationChange = tmpCVQueDiff;
                            }
                        }
                        else
                        {
                            tmpQueueSource = clsEnums.enQueueSource.CV;
                            tmpBOQMMLocationChange = tmpCVQueDiff;
                        }
                    }
                    else if (clsGlobalVars.InfrastructureBOQMMLocation != -1)
                    {
                        if (chkFilterQueues.Checked == true)
                        {
                            if (Math.Abs(tmpTSSQueDiff) <= clsGlobalVars.LinkLength)
                            {
                                tmpQueueSource = clsEnums.enQueueSource.TSS;
                                tmpBOQMMLocationChange = tmpTSSQueDiff;
                            }
                        }
                        else
                        {
                            tmpQueueSource = clsEnums.enQueueSource.TSS;
                            tmpBOQMMLocationChange = tmpTSSQueDiff;
                        }
                    }
                    else
                    {
                        tmpQueDiff = Roadway.RecurringCongestionMMLocation - clsGlobalVars.BOQMMLocation;
                        if (chkFilterQueues.Checked == true)
                        {
                            if (Math.Abs(tmpQueDiff) <= clsGlobalVars.LinkLength)
                            {
                                tmpQueueSource = clsEnums.enQueueSource.NA;
                                tmpBOQMMLocationChange = tmpQueDiff;
                            }
                        }
                        else
                        {
                            tmpQueueSource = clsEnums.enQueueSource.NA;
                            tmpBOQMMLocationChange = tmpQueDiff;
                        }
                    }
                    if (tmpBOQMMLocationChange > 0)
                    {
                        tmpQueChangeDirection = clsEnums.enQueueCahnge.Growing;
                    }
                    else if (tmpBOQMMLocationChange < 0)
                    {
                        tmpQueChangeDirection = clsEnums.enQueueCahnge.Dissipating;
                    }
                    else if (tmpBOQMMLocationChange == 0)
                    {
                        tmpQueChangeDirection = clsEnums.enQueueCahnge.Same;
                    }
                }
                #endregion
                #region "No Queue"
                else if (clsGlobalVars.BOQMMLocation == -1)
                {
                    tmpTSSQueDiff = clsGlobalVars.InfrastructureBOQMMLocation - Roadway.RecurringCongestionMMLocation;
                    tmpCVQueDiff = clsGlobalVars.CVBOQMMLocation - Roadway.RecurringCongestionMMLocation;
                    if ((clsGlobalVars.InfrastructureBOQMMLocation != -1) && (clsGlobalVars.CVBOQMMLocation != -1))
                    {
                        if (chkFilterQueues.Checked == true)
                        {
                            if ((Math.Abs(tmpTSSQueDiff) <= 2 * clsGlobalVars.LinkLength) && (Math.Abs(tmpCVQueDiff) <= 2 * clsGlobalVars.LinkLength))
                            {
                                if (tmpCVQueDiff >= tmpTSSQueDiff)
                                {
                                    tmpQueueSource = clsEnums.enQueueSource.CV;
                                    tmpBOQMMLocationChange = tmpCVQueDiff;
                                }
                                else
                                {
                                    tmpQueueSource = clsEnums.enQueueSource.TSS;
                                    tmpBOQMMLocationChange = tmpTSSQueDiff;
                                }
                            }
                            else if (Math.Abs(tmpTSSQueDiff) <= 2 * clsGlobalVars.LinkLength)
                            {
                                tmpQueueSource = clsEnums.enQueueSource.TSS;
                                tmpBOQMMLocationChange = tmpTSSQueDiff;
                            }
                            else if (Math.Abs(tmpCVQueDiff) <= 2 * clsGlobalVars.LinkLength)
                            {
                                tmpQueueSource = clsEnums.enQueueSource.CV;
                                tmpBOQMMLocationChange = tmpCVQueDiff;
                            }
                        }
                        else
                        {
                            if (tmpCVQueDiff >= tmpTSSQueDiff)
                            {
                                tmpQueueSource = clsEnums.enQueueSource.CV;
                                tmpBOQMMLocationChange = tmpCVQueDiff;
                            }
                            else
                            {
                                tmpQueueSource = clsEnums.enQueueSource.TSS;
                                tmpBOQMMLocationChange = tmpTSSQueDiff;
                            }
                        }
                    }
                    else if (clsGlobalVars.CVBOQMMLocation != -1)
                    {
                        if (chkFilterQueues.Checked == true)
                        {
                            if (Math.Abs(tmpCVQueDiff) <= 2 * clsGlobalVars.LinkLength)
                            {
                                tmpQueueSource = clsEnums.enQueueSource.CV;
                                tmpBOQMMLocationChange = tmpCVQueDiff;
                            }
                        }
                        else
                        {
                            tmpQueueSource = clsEnums.enQueueSource.CV;
                            tmpBOQMMLocationChange = tmpCVQueDiff;
                        }
                    }
                    else if (clsGlobalVars.InfrastructureBOQMMLocation != -1)
                    {
                        if (chkFilterQueues.Checked == true)
                        {
                            if (Math.Abs(tmpTSSQueDiff) <= 2 * clsGlobalVars.LinkLength)
                            {
                                tmpQueueSource = clsEnums.enQueueSource.TSS;
                                tmpBOQMMLocationChange = tmpTSSQueDiff;
                            }
                        }
                        else
                        {
                            tmpQueueSource = clsEnums.enQueueSource.TSS;
                            tmpBOQMMLocationChange = tmpTSSQueDiff;
                        }
                    }
                    else
                    {
                        tmpQueueSource = clsEnums.enQueueSource.NA;
                        tmpBOQMMLocationChange = 0;
                    }
                    if (tmpBOQMMLocationChange > 0)
                    {
                        tmpQueChangeDirection = clsEnums.enQueueCahnge.Growing;
                    }
                    else if (tmpBOQMMLocationChange <= 0)
                    {
                        tmpQueChangeDirection = clsEnums.enQueueCahnge.Dissipating;
                    }
                }
                #endregion
            }

            if (tmpQueueSource == clsEnums.enQueueSource.CV)
            {
                clsGlobalVars.BOQMMLocation = clsGlobalVars.CVBOQMMLocation;
                clsGlobalVars.BOQTime = DateTime.Now;
                clsGlobalVars.BOQSpeed = clsGlobalVars.CVBOQSublinkSpeed;
            }
            else if (tmpQueueSource == clsEnums.enQueueSource.TSS)
            {
                clsGlobalVars.BOQMMLocation = clsGlobalVars.InfrastructureBOQMMLocation;
                clsGlobalVars.BOQTime = DateTime.Now;
                clsGlobalVars.BOQSpeed = clsGlobalVars.InfrastructureBOQLinkSpeed;
            }
            else if ((tmpQueueSource == clsEnums.enQueueSource.NA) && (tmpBOQMMLocationChange != 0))
            {
                clsGlobalVars.BOQMMLocation = -1;
                clsGlobalVars.BOQTime = DateTime.Now;
                clsGlobalVars.BOQSpeed = 0;
            }
            #endregion

            TimeSpan span = clsGlobalVars.BOQTime.Subtract(clsGlobalVars.PrevBOQTime);
            clsGlobalVars.QueueRate = ((Math.Abs(tmpBOQMMLocationChange) * 3600) / span.TotalSeconds);
            clsGlobalVars.QueueChange = tmpQueChangeDirection;

            #region "Determine queue rate of change"
                string tmpSublinkData = string.Empty;
                if (QueuedSubLink != null)
                {
                    tmpSublinkData = QueuedSubLink.TotalNumberCVs + "," + QueuedSubLink.PrevTotalNumberCVs + "," + QueuedSubLink.VolumeDiff + "," + QueuedSubLink.FlowRate.ToString("0.00") + "," +
                                     QueuedSubLink.Density.ToString("0.00") + "," + QueuedSubLink.PrevDensity.ToString("0.00") + "," + QueuedSubLink.DensityDiff + "," + QueuedSubLink.ShockWaveRate.ToString("0.00");
                }

                if (Stopped == false)
                {
                    QueueLog.WriteLine(DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + "::" + DateTime.Now.Millisecond + "," +
                                       clsGlobalVars.PrevBOQMMLocation.ToString("0.0") + "," +
                                       clsGlobalVars.PrevBOQTime.Hour + ":" + clsGlobalVars.PrevBOQTime.Minute + ":" + clsGlobalVars.PrevBOQTime.Second + ":" + clsGlobalVars.PrevBOQTime.Millisecond + "," +
                                       clsGlobalVars.BOQMMLocation.ToString("0.0") + "," +
                                       clsGlobalVars.BOQTime.Hour + ":" + clsGlobalVars.BOQTime.Minute + ":" + clsGlobalVars.BOQTime.Second + ":" + clsGlobalVars.BOQTime.Millisecond + "," +
                                       tmpBOQMMLocationChange + "," + span.TotalSeconds.ToString("0") + "," + clsGlobalVars.QueueRate.ToString("0") + "," + clsGlobalVars.QueueChange.ToString() + "," +
                                       clsGlobalVars.QueueSource.ToString() + "," + clsGlobalVars.CVBOQMMLocation + "," + clsGlobalVars.InfrastructureBOQMMLocation + "," + tmpSublinkData);
                }
                LogTxtMsg(txtINFLOLog, "\t\tPrev BOQ MM: " + clsGlobalVars.PrevBOQMMLocation.ToString("0.0") + "\t\tTime: " + clsGlobalVars.PrevBOQTime);
                if (clsGlobalVars.PrevBOQMMLocation != clsGlobalVars.BOQMMLocation)
                {
                    LogTxtMsg(txtINFLOLog, "\t\tCurr BOQ MM:  " + clsGlobalVars.BOQMMLocation + "\t\tTime: " + clsGlobalVars.BOQTime + "\tSource: " + clsGlobalVars.QueueSource.ToString());
                    clsGlobalVars.PrevBOQMMLocation = clsGlobalVars.BOQMMLocation;
                    clsGlobalVars.PrevBOQTime = clsGlobalVars.BOQTime;
                }
                if (clsGlobalVars.BOQMMLocation != -1)
                {
                    InsertQueueInfoIntoINFLODatabase(clsGlobalVars.BOQMMLocation, clsGlobalVars.BOQTime, clsGlobalVars.QueueRate, clsGlobalVars.QueueChange, clsGlobalVars.QueueSource, clsGlobalVars.QueueSpeed, Roadway);
                }
                //else
                //{
                //    InsertQueueInfoIntoINFLODatabase(0, clsGlobalVars.BOQTime, clsGlobalVars.QueueRate, clsGlobalVars.QueueChange, clsGlobalVars.QueueSource, clsGlobalVars.QueueSpeed, Roadway);
                //}
            #endregion

            //LogTxtMsg(txtTMEQWARNLog, "\r\n\t\tPrev BOQ MM location: " + clsGlobalVars.PrevBOQMMLocation + "\tTime: " + clsGlobalVars.PrevBOQTime);
            //CVDataProcessor.WriteLine("\tPrev BOQ MM location: " + clsGlobalVars.PrevBOQMMLocation + "\tTime: " + clsGlobalVars.PrevBOQTime);

            //LogTxtMsg(txtINFLOLog, "\t\tCurr BOQ MM location: " + clsGlobalVars.BOQMMLocation + "\tTime: " + clsGlobalVars.BOQTime + "\tSource: " + clsGlobalVars.QueueSource.ToString());
            //CVDataProcessor.WriteLine("\tCurr BOQ MM location: " + clsGlobalVars.BOQMMLocation + "\tTime: " + clsGlobalVars.BOQTime + "\tSource: " + clsGlobalVars.QueueSource.ToString());

            //Current BOQ
            if (clsGlobalVars.BOQMMLocation == -1)
            {
                clsGlobalVars.QueueLength = 0;
                clsGlobalVars.QueueSpeed = 0;
                DisplayForm.txtBOQ.Text = "No Queue";
                DisplayForm.txtQueueLength.Text = "";
                DisplayForm.txtQueueGrowthRate.Text = "";
                DisplayForm.txtQueueSpeed.Text = "";

            }
            else
            {
                double TotalQuedSublinksCVSpeed = 0;
                double TotalQuedSublinksVolume = 0;
                foreach (clsRoadwaySubLink RSL in RSLList)
                {
                    if (Roadway.Direction == Roadway.MMIncreasingDirection)
                    {
                        if ((RSL.BeginMM >= clsGlobalVars.BOQMMLocation) && (RSL.BeginMM < Roadway.RecurringCongestionMMLocation))
                        {
                            TotalQuedSublinksCVSpeed = TotalQuedSublinksCVSpeed + RSL.TotalNumberCVs * RSL.CVAvgSpeed;
                            TotalQuedSublinksVolume = TotalQuedSublinksVolume + RSL.TotalNumberCVs;
                        }
                    }
                    else if (Roadway.Direction != Roadway.MMIncreasingDirection)
                    {
                        if ((RSL.BeginMM <= clsGlobalVars.BOQMMLocation) && (RSL.BeginMM > Roadway.RecurringCongestionMMLocation))
                        {
                            TotalQuedSublinksCVSpeed = TotalQuedSublinksCVSpeed + RSL.TotalNumberCVs * RSL.CVAvgSpeed;
                            TotalQuedSublinksVolume = TotalQuedSublinksVolume + RSL.TotalNumberCVs;
                        }
                    }
                }
                clsGlobalVars.QueueSpeed = TotalQuedSublinksCVSpeed / TotalQuedSublinksVolume;
                clsGlobalVars.QueueLength = Roadway.RecurringCongestionMMLocation - clsGlobalVars.BOQMMLocation;
                DisplayForm.txtBOQ.Text = clsGlobalVars.BOQMMLocation.ToString();
                DisplayForm.txtQueueLength.Text = (Roadway.RecurringCongestionMMLocation - clsGlobalVars.BOQMMLocation).ToString("0.00");
                DisplayForm.txtQueueGrowthRate.Text = clsGlobalVars.QueueRate.ToString("0");
                DisplayForm.txtQueueSpeed.Text = clsGlobalVars.QueueSpeed.ToString("0");
            }
            LogTxtMsg(txtINFLOLog, "\t\tCurr BOQ rate of growth: " + clsGlobalVars.QueueRate.ToString("0.00") + "\tQueue direction: " + clsGlobalVars.QueueChange.ToString() + "\tQueue source: " + clsGlobalVars.QueueSource.ToString());
            //CVDataProcessor.WriteLine("\tCurr BOQ rate of growth: " + clsGlobalVars.QueueRate.ToString("0.00") + "\tQueue direction: " + clsGlobalVars.QueueChange.ToString() + "\tQueue source: " + clsGlobalVars.QueueSource.ToString());
        }
        private void ApplyWRTMSpeed(ref List<clsRoadwayLink> RLList, double WRTMRecommendedSpeed, double WRTMBeginMM, double WRTMEndMM)
        {
            foreach (clsRoadwayLink RL in RLList)
            {
                if ((RL.BeginMM >= WRTMBeginMM) && (RL.BeginMM < WRTMEndMM))
                {
                    RL.WRTMSpeed = WRTMRecommendedSpeed;
                }
            }
        }
        private void DetermineLinkHarmonizedSpeed(List<clsRoadwaySubLink> RSLList, ref List<clsRoadwayLink> RLList, double TroupingEndMM)
        {
            foreach (clsRoadwayLink RL in RLList)
            {
                RL.RecommendedSpeed = 0;
                foreach (clsRoadwaySubLink RSL in RSLList)
                {
                    if (RL.BeginMM < TroupingEndMM)
                    {
                        if (RSL.BeginMM == RL.BeginMM)
                        {
                            RL.RecommendedSpeed = RSL.HarmonizedSpeed;
                            break;
                        }
                    }
                    else
                    {
                        if (RSL.BeginMM == RL.BeginMM)
                        {
                            RL.RecommendedSpeed = RSL.RecommendedSpeed;
                            break;
                        }
                    }
                }
            }
        }

        private void tmrCVData_Tick(object sender, EventArgs e)
        {
        }
        private void ProcessCVData()
        {
            string retValue = string.Empty;
            string CurrentSection = string.Empty;

            tmrCVData.Enabled = false;

            //Calculate time difference between previous wakeuptime of the CV Timer for processing CV data and the current wakeupt time
            CVCurrWakeupTime = DateTime.Now;
            TimeSpan spandiff = CVCurrWakeupTime.Subtract(CVPrevWakeupTime);
            CVTimeDiff = spandiff.TotalSeconds;
            CVPrevWakeupTime = CVCurrWakeupTime;

            DateTime currDateTime = DateTime.Now;
            currDateTime = TimeZoneInfo.ConvertTimeToUtc(currDateTime, TimeZoneInfo.Local);

            System.Windows.Forms.Application.DoEvents();
            System.Windows.Forms.Application.DoEvents();
            //tabINFLOApps.SelectTab("tabPgCVDataAggregation");
            
            LogTxtMsg(txtCVDataLog, "\r\n------------------------------");
            LogTxtMsg(txtCVDataLog, "\t\tTime difference between CV Timer concecutive wakeup times: " + CVTimeDiff.ToString("0"));
            LogTxtMsg(txtCVDataLog,DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + "::" + DateTime.Now.Millisecond +
                                    "\tGet CV traffic data for the last " + clsGlobalVars.CVDataPollingFrequency + " seconds interval From: " +
                                    currDateTime.AddSeconds(-clsGlobalVars.CVDataPollingFrequency) + "\tTo: " + currDateTime);
            CVDataProcessor.WriteLine(DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + "::" + DateTime.Now.Millisecond +
                                    "\tGet CV traffic data for the last " + clsGlobalVars.CVDataPollingFrequency + " seconds interval From: " + 
                                    currDateTime.AddSeconds(-clsGlobalVars.CVDataPollingFrequency) + "\tTo: " + currDateTime);

            //Get Last interval CV data
            #region "Get last interval CV data
            TimeSpan GetCVDataTime = new TimeSpan(DateTime.Now.Ticks);
            int NumberRecordsRetrieved = 0;
            retValue = GetLastIntervalCVData(DB, currDateTime, ref CurrIntervalCVList, Roadway.Direction, Roadway.LowerHeading, Roadway.UpperHeading, ref NumberRecordsRetrieved);
            if (retValue.Length > 0)
            {
                LogTxtMsg(txtCVDataLog, "\r\n\t" + retValue);
                CVDataProcessor.WriteLine("\r\n\t" + retValue);
            }
            TimeSpan EndGetCVDataTime = new TimeSpan(DateTime.Now.Ticks);
            LogTxtMsg(txtCVDataLog, "\t\tTime for retrieving " + CurrIntervalCVList.Count + " CV traffic data records from INFLO database: \t" + (EndGetCVDataTime.TotalMilliseconds - GetCVDataTime.TotalMilliseconds).ToString("0") + " msecs");
            LogTxtMsg(txtINFLOLog, "\t\tTime for retrieving " + CurrIntervalCVList.Count + " CV traffic data records from INFLO database: \t" + (EndGetCVDataTime.TotalMilliseconds - GetCVDataTime.TotalMilliseconds).ToString("0") + " msecs");
            CVDataProcessor.WriteLine("\r\n" + DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + "::" + DateTime.Now.Millisecond +
                                      "\tRetrieved " + CurrIntervalCVList.Count + " CV traffic data records for the last " + clsGlobalVars.CVDataPollingFrequency + " seconds interval from INFLO database.");
            #endregion

            if (NumberRecordsRetrieved > 0)
            {
                NumberNoCVDataIntervals = 0;
                DisplayForm.DisableNoCVDataMessage();
                //Reset Troupe speed, and inclusion flag for every sublink
                foreach (clsRoadwaySubLink rsl in RSLList)
                {
                    rsl.TroupeSpeed = 0;
                    rsl.TroupeInclusionOverride = false;
                    rsl.BeginTroupe = false;
                    rsl.TroupeProcessed = DateTime.Now;
                    rsl.HarmonizedSpeed = 0;
                    rsl.SpdHarmInclusionOverride = false;
                    rsl.BeginSpdHarm = false;
                }

                #region "If CV data was found for the last 5 seconds"

                #region "Log received CV Data"
                //CVDataProcessor.WriteLine("\r\n\tNomadicDeviceID, DateGenerated, Heading, latitude, Longitude, Speed, MMlocation, Queued, Temperature, CoefficientFriction, RoadwayId, SubLinkId, Direction");
                //foreach (clsCVData CV in CurrIntervalCVList)
                //{
                //    CVDataProcessor.WriteLine("\t" + CV.NomadicDeviceID + ", " + CV.DateGenerated + ", " + CV.Heading + ", " + CV.Latitude + ", " + CV.Longitude + ", " + CV.Speed + ", " + CV.MMLocation + ", " +
                //                              CV.Queued + ", " + CV.Temperature + ", " + CV.CoefficientFriction + ", " + CV.RoadwayID + ", " + CV.SublinkID + ", " + CV.Direction);
                //}
                #endregion

                #region "Calculate sublink speed and queued state using CV data
                LogTxtMsg(txtCVDataLog, "\r\n\t\tUpdate roadway sublink traffic parameters using CV data from the last  " + clsGlobalVars.CVDataPollingFrequency + " seconds interval.");
                retValue = ProcessRoadwaySublinkQueuedStatus(ref CurrIntervalCVList, ref RSLList);
                if (retValue.Length > 0)
                {
                    LogTxtMsg(txtCVDataLog, "\r\n\t" + retValue);
                    return;
                }
                #endregion

                #region "Log Updated sublink data"
                //CVDataProcessor.WriteLine("\r\n" + DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + "::" + DateTime.Now.Millisecond + 
                //                                   "\tUpdated Roadway SubLink status using CV data.");
                //CVDataProcessor.WriteLine("\r\n\tRoadwayId, SubLinkId, DateProcessed, BeginMM, EndMM, NumberQueuedCVs, TotalNumberCVS, PercentQueuedCVs, Queued, AvgSpeed, Direction");
                //foreach (clsRoadwaySubLink rsl in RSLList)
                //{
                //    CVDataProcessor.WriteLine("\t" + rsl.RoadwayID + ", " + rsl.Identifier + ", " + rsl.DateProcessed + ", " + rsl.BeginMM + ", " + rsl.EndMM + ", " + rsl.NumberQueuedCVs + ", " +
                //                              rsl.TotalNumberCVs + ", " + rsl.PercentQueuedCVs + ", " + rsl.Queued + ", " + rsl.CVAvgSpeed.ToString("0") + ", " + rsl.Direction);
                //}
                #endregion

                //Excel - Log Sublink CV speed, CV % queued vehicles, CV # Queued vehicles, and CV Total # vehicles into Excel worksheet
                if (chkExcelDataLogging.Checked == true)
                {
                    Color myPurpleColor = Color.FromArgb(255, 192, 255);
                    #region "Log Sublink CV speed, CV % queued vehicles, CV # Queued vehicles, and CV Total # vehicles into Excel worksheet"
                    int s = 0;
                    CVWorkSheets[1].Cells[CVWSCurrRow, 1] = currDateTime;
                    foreach (clsRoadwaySubLink rsl in RSLList)
                    {
                        //Excel
                        CVWorkSheets[1].Cells[CVWSCurrRow, s + 2] = rsl.CVAvgSpeed.ToString("0") + "::" + rsl.PercentQueuedCVs.ToString("0") + "::" + rsl.NumberQueuedCVs.ToString("0") + "::" + rsl.TotalNumberCVs;
                        if (rsl.Queued == true)
                        {
                            CVWorkSheets[1].Cells[CVWSCurrRow, s + 2].Interior.Color = System.Drawing.Color.Red;
                        }
                        else 
                        {
                            CVWorkSheets[1].Cells[CVWSCurrRow, s + 2].Interior.Color = myPurpleColor;
                        }
                        s = s + 1;
                    }
                    CVWSCurrRow = CVWSCurrRow + 1;
                    #endregion
                }
                TimeSpan Processingtime = new TimeSpan(DateTime.Now.Ticks);
                LogTxtMsg(txtCVDataLog, "\t\tTime for processing " + CurrIntervalCVList.Count + " CV traffic data records from the last interval: \t" + (Processingtime.TotalMilliseconds - EndGetCVDataTime.TotalMilliseconds).ToString("0") + " msecs");
                //CVDataProcessor.WriteLine("\r\n" + DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + "::" + DateTime.Now.Millisecond +
                //                          "\tTime for processing: " + CurrIntervalCVList.Count + " CV traffic data records from the last " + clsGlobalVars.CVDataPollingFrequency + " seconds interval:= " +
                //                                                    (Processingtime.TotalMilliseconds - EndGetCVDataTime.TotalMilliseconds).ToString("0") + " msecs");
                
                #region "Insert Sublink Status Info into INFLO Database"
                ///retValue = InsertSubLinkStatusIntoINFLODatabase();
                //if (retValue.Length > 0)
                //{
                //    LogTxtMsg(txtCVDataLog, retValue);
                //}
                //else
                //{
                //    LogTxtMsg(txtTSSDataLog, DateTime.Now + "\t\tFinish inserting processed TSS roadway links status into INFLO database");
                //}
                #endregion

                System.Windows.Forms.Application.DoEvents();
                System.Windows.Forms.Application.DoEvents();

                //Determine CV sublink BOQ location
                TimeSpan BeginCVBOQ = new TimeSpan(DateTime.Now.Ticks);
                LogTxtMsg(txtCVDataLog, "\r\n\t\tDetermine CV SubLink BOQ location:");
                CVDataProcessor.WriteLine("\r\n" + DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + "::" + DateTime.Now.Millisecond + "\tDetermine CV SubLink BOQ location:");

                clsGlobalVars.PrevCVBOQMMLocation = clsGlobalVars.CVBOQMMLocation;
                clsGlobalVars.PrevCVBOQTime = clsGlobalVars.CVBOQTime;
                clsGlobalVars.CVBOQMMLocation = -1;

                #region "Determine CV Sublink BOQ MM Location"
                //TotalCVs = 0;
                //QueuedCVs = 0;
                QueuedSubLink = new clsRoadwaySubLink();
                RSLList.Sort((l, r) => l.BeginMM.CompareTo(r.BeginMM));
                foreach (clsRoadwaySubLink RSL in RSLList)
                {
                    if (RSL.Queued == true)
                    {
                        if (Roadway.MMIncreasingDirection == Roadway.Direction)
                        {
                            if (RSL.BeginMM < Roadway.RecurringCongestionMMLocation)
                            {
                                if (clsGlobalVars.CVBOQMMLocation != -1)
                                {
                                    if (RSL.BeginMM < clsGlobalVars.CVBOQMMLocation)
                                    {
                                        clsGlobalVars.CVBOQMMLocation = RSL.BeginMM;
                                        clsGlobalVars.CVBOQTime = DateTime.Now;
                                        clsGlobalVars.CVBOQSublinkSpeed = RSL.CVAvgSpeed;
                                        //QueuedSubLink = new clsRoadwaySubLink();
                                        QueuedSubLink = RSL;
                                    }
                                }
                                else
                                {
                                    clsGlobalVars.CVBOQMMLocation = RSL.BeginMM;
                                    clsGlobalVars.CVBOQTime = DateTime.Now;
                                    clsGlobalVars.CVBOQSublinkSpeed = RSL.CVAvgSpeed;
                                    //QueuedSubLink = new clsRoadwaySubLink();
                                    QueuedSubLink = RSL;
                                }
                            }
                        }
                        else if (Roadway.MMIncreasingDirection != Roadway.Direction)
                        {
                            if (RSL.BeginMM > Roadway.RecurringCongestionMMLocation)
                            {
                                if (clsGlobalVars.CVBOQMMLocation != -1)
                                {
                                    if (RSL.BeginMM > clsGlobalVars.CVBOQMMLocation)
                                    {
                                        clsGlobalVars.CVBOQMMLocation = RSL.BeginMM;
                                        clsGlobalVars.CVBOQTime = DateTime.Now;
                                        clsGlobalVars.CVBOQSublinkSpeed = RSL.CVAvgSpeed;
                                        //QueuedSubLink = new clsRoadwaySubLink();
                                        QueuedSubLink = RSL;
                                    }
                                }
                                else
                                {
                                    clsGlobalVars.CVBOQMMLocation = RSL.BeginMM;
                                    clsGlobalVars.CVBOQTime = DateTime.Now;
                                    clsGlobalVars.CVBOQSublinkSpeed = RSL.CVAvgSpeed;
                                    //QueuedSubLink = new clsRoadwaySubLink();
                                    QueuedSubLink = RSL;
                                }
                            }
                        }
                    }
                }
                #endregion

                DisplayForm.ClearCVSubLinkQueuedStatus();
                DisplayForm.DisplayCVSubLinkQueuedStatus(RSLList);

                LogTxtMsg(txtCVDataLog, "\t\t\tPrev CV  BOQ MM location: " + clsGlobalVars.PrevCVBOQMMLocation + "\tTime: " + clsGlobalVars.PrevCVBOQTime);
                LogTxtMsg(txtCVDataLog, "\t\t\tCurr CV  BOQ MM location: " + clsGlobalVars.CVBOQMMLocation + "\tTime: " + clsGlobalVars.CVBOQTime);
                CVDataProcessor.WriteLine("\tPrev CV  BOQ MM location: " + clsGlobalVars.PrevCVBOQMMLocation + "\tTime: " + clsGlobalVars.PrevCVBOQTime);
                CVDataProcessor.WriteLine("\tCurr CV  BOQ MM location: " + clsGlobalVars.CVBOQMMLocation + "\tTime: " + clsGlobalVars.CVBOQTime);

                //CV BOQ
                DisplayForm.txtCVDate.Text = DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second;
                if (clsGlobalVars.CVBOQMMLocation == -1)
                {
                    txtCVBOQMMLocation.Text = "No Queue";
                    txtCVBOQDate.Text = clsGlobalVars.CVBOQTime.ToString();
                    DisplayForm.txtCVBOQ.Text = "No Queue";
                }
                else
                {
                    txtCVBOQMMLocation.Text = clsGlobalVars.CVBOQMMLocation.ToString();
                    txtCVBOQDate.Text = clsGlobalVars.CVBOQTime.ToString();
                    DisplayForm.txtCVBOQ.Text = clsGlobalVars.CVBOQMMLocation.ToString();
                }

                //CV Previou BOQ
                if (clsGlobalVars.PrevCVBOQMMLocation == -1)
                {
                    txtCVPrevBOQMMLocation.Text = "No Queue";
                    txtCVPrevBOQDate.Text = clsGlobalVars.PrevCVBOQTime.ToString() ;
                }
                else
                {
                    txtCVPrevBOQMMLocation.Text = clsGlobalVars.PrevCVBOQMMLocation.ToString();
                    txtCVPrevBOQDate.Text = clsGlobalVars.PrevCVBOQTime.ToString();
                }

                TimeSpan EndCVBOQ = new TimeSpan(DateTime.Now.Ticks);
                LogTxtMsg(txtCVDataLog, "\t\tTime for processing CVBOQ MM Location: \t" + (EndCVBOQ.TotalMilliseconds - BeginCVBOQ.TotalMilliseconds).ToString("0") + " msecs");

                #region "Reconcile the CV BOQ and TSS BOQ - Old Algorithm"
                /*
                LogTxtMsg(txtCVDataLog, "\tDetermine BOQ MM location from CV and TSS BOQ MM locations:");
                CVDataProcessor.WriteLine("\r\n" + DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + "::" + DateTime.Now.Millisecond +
                                          "\tDetermine BOQ MM location from CV and TSS BOQ MM locations:");

                #region "Reconcile CV BOQ MM Location and TSS data BOQ MM Location"
                if (Roadway.Direction == Roadway.MMIncreasingDirection)
                {
                    if ((clsGlobalVars.InfrastructureBOQMMLocation != -1) && (clsGlobalVars.CVBOQMMLocation != -1))
                    {
                        if (clsGlobalVars.CVBOQMMLocation < clsGlobalVars.InfrastructureBOQMMLocation)
                        {
                            clsGlobalVars.BOQMMLocation = clsGlobalVars.CVBOQMMLocation;
                            clsGlobalVars.BOQTime = DateTime.Now;
                            clsGlobalVars.QueueSource = clsEnums.enQueueSource.CV;
                        }
                        else
                        {
                            clsGlobalVars.BOQMMLocation = clsGlobalVars.InfrastructureBOQMMLocation;
                            clsGlobalVars.BOQTime = DateTime.Now;
                            clsGlobalVars.QueueSource = clsEnums.enQueueSource.TSS;
                        }
                    }
                    else if (clsGlobalVars.CVBOQMMLocation != -1)
                    {
                        clsGlobalVars.BOQMMLocation = clsGlobalVars.CVBOQMMLocation;
                        clsGlobalVars.BOQTime = DateTime.Now;
                        clsGlobalVars.QueueSource = clsEnums.enQueueSource.CV;
                    }
                    else if (clsGlobalVars.InfrastructureBOQMMLocation != -1)
                    {
                        clsGlobalVars.BOQMMLocation = clsGlobalVars.InfrastructureBOQMMLocation;
                        clsGlobalVars.BOQTime = DateTime.Now;
                        clsGlobalVars.QueueSource = clsEnums.enQueueSource.TSS;
                    }
                    else
                    {
                        clsGlobalVars.BOQMMLocation = -1;
                        clsGlobalVars.BOQTime = DateTime.Now;
                        clsGlobalVars.QueueSource = clsEnums.enQueueSource.NA;
                    }
                }
                else if (Roadway.Direction != Roadway.MMIncreasingDirection)
                {
                    if ((clsGlobalVars.InfrastructureBOQMMLocation != -1) && (clsGlobalVars.CVBOQMMLocation != -1))
                    {
                        if (clsGlobalVars.CVBOQMMLocation > clsGlobalVars.InfrastructureBOQMMLocation)
                        {
                            clsGlobalVars.BOQMMLocation = clsGlobalVars.CVBOQMMLocation;
                            clsGlobalVars.BOQTime = DateTime.Now;
                            clsGlobalVars.QueueSource = clsEnums.enQueueSource.CV;
                        }
                        else
                        {
                            clsGlobalVars.BOQMMLocation = clsGlobalVars.InfrastructureBOQMMLocation;
                            clsGlobalVars.BOQTime = DateTime.Now;
                            clsGlobalVars.QueueSource = clsEnums.enQueueSource.TSS;
                        }
                    }
                    else if (clsGlobalVars.CVBOQMMLocation != -1)
                    {
                        clsGlobalVars.BOQMMLocation = clsGlobalVars.CVBOQMMLocation;
                        clsGlobalVars.BOQTime = DateTime.Now;
                        clsGlobalVars.QueueSource = clsEnums.enQueueSource.CV;
                    }
                    else if (clsGlobalVars.InfrastructureBOQMMLocation != -1)
                    {
                        clsGlobalVars.BOQMMLocation = clsGlobalVars.InfrastructureBOQMMLocation;
                        clsGlobalVars.BOQTime = DateTime.Now;
                        clsGlobalVars.QueueSource = clsEnums.enQueueSource.TSS;
                    }
                    else
                    {
                        clsGlobalVars.BOQMMLocation = -1;
                        clsGlobalVars.BOQTime = DateTime.Now;
                        clsGlobalVars.QueueSource = clsEnums.enQueueSource.NA;
                    }
                }
                #endregion

                double tmpBOQMMLocationChange = 0;

                #region "Determine queue rate of change"
                    TimeSpan span = clsGlobalVars.BOQTime.Subtract(clsGlobalVars.PrevBOQTime);
                    if (Roadway.Direction == Roadway.MMIncreasingDirection)
                    {
                        if ((clsGlobalVars.PrevBOQMMLocation != -1) && (clsGlobalVars.BOQMMLocation != -1))
                        {
                            tmpBOQMMLocationChange = clsGlobalVars.BOQMMLocation - clsGlobalVars.PrevBOQMMLocation;
                            clsGlobalVars.QueueRate = ((Math.Abs(tmpBOQMMLocationChange) * 3600) / span.TotalSeconds);
                            if (tmpBOQMMLocationChange < 0)
                            {
                                clsGlobalVars.QueueChange = clsEnums.enQueueCahnge.Growing;
                            }
                            else if (tmpBOQMMLocationChange > 0)
                            {
                                clsGlobalVars.QueueChange = clsEnums.enQueueCahnge.Dissipating;
                            }
                            else if (tmpBOQMMLocationChange == 0)
                            {
                                clsGlobalVars.QueueChange = clsEnums.enQueueCahnge.Same;
                            }
                        }
                        else if ((clsGlobalVars.PrevBOQMMLocation != -1) && (clsGlobalVars.BOQMMLocation == -1))
                        {
                            tmpBOQMMLocationChange = (Roadway.RecurringCongestionMMLocation - clsGlobalVars.PrevBOQMMLocation);
                            clsGlobalVars.QueueRate = ((Math.Abs(tmpBOQMMLocationChange) * 3600) / span.TotalSeconds);
                            clsGlobalVars.QueueChange = clsEnums.enQueueCahnge.Dissipating;
                        }
                        else if ((clsGlobalVars.PrevBOQMMLocation == -1) && (clsGlobalVars.BOQMMLocation != -1))
                        {
                            tmpBOQMMLocationChange = (Roadway.RecurringCongestionMMLocation - clsGlobalVars.BOQMMLocation);
                            clsGlobalVars.QueueRate = ((Math.Abs(tmpBOQMMLocationChange) * 3600) / span.TotalSeconds);
                            clsGlobalVars.QueueChange = clsEnums.enQueueCahnge.Growing;
                        }
                        else
                        {
                            clsGlobalVars.QueueChange = clsEnums.enQueueCahnge.Same;
                            clsGlobalVars.QueueRate = 0;
                        }
                        
                        QueueLog.WriteLine(DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + "::" + DateTime.Now.Millisecond + ",CV," + 
                                           clsGlobalVars.PrevBOQMMLocation.ToString("0.0") + "," + clsGlobalVars.PrevBOQTime.ToString() + "," + 
                                           clsGlobalVars.BOQMMLocation.ToString("0.0") + "," + clsGlobalVars.BOQTime.ToString() + "," + clsGlobalVars.QueueRate.ToString("0.00") + "," + 
                                           clsGlobalVars.QueueChange.ToString() + "," + clsGlobalVars.QueueSource.ToString());
                        if (clsGlobalVars.PrevBOQMMLocation != clsGlobalVars.BOQMMLocation)
                        {
                            clsGlobalVars.PrevBOQMMLocation = clsGlobalVars.BOQMMLocation;
                            clsGlobalVars.PrevBOQTime = clsGlobalVars.BOQTime;
                        }
                    }
                    else if (Roadway.Direction != Roadway.MMIncreasingDirection)
                    {
                        if ((clsGlobalVars.PrevBOQMMLocation != -1) && (clsGlobalVars.BOQMMLocation != -1))
                        {
                            tmpBOQMMLocationChange = clsGlobalVars.BOQMMLocation - clsGlobalVars.PrevBOQMMLocation;
                            clsGlobalVars.QueueRate = ((Math.Abs(tmpBOQMMLocationChange) * 3600) / span.TotalSeconds);
                            if (tmpBOQMMLocationChange < 0)
                            {
                                clsGlobalVars.QueueChange = clsEnums.enQueueCahnge.Dissipating;
                            }
                            else if (tmpBOQMMLocationChange > 0)
                            {
                                clsGlobalVars.QueueChange = clsEnums.enQueueCahnge.Growing;
                            }
                            else if (tmpBOQMMLocationChange == 0)
                            {
                                clsGlobalVars.QueueChange = clsEnums.enQueueCahnge.Same;
                            }
                        }
                        else if ((clsGlobalVars.PrevBOQMMLocation != -1) && (clsGlobalVars.BOQMMLocation == -1))
                        {
                            tmpBOQMMLocationChange = (clsGlobalVars.PrevBOQMMLocation - Roadway.RecurringCongestionMMLocation);
                            clsGlobalVars.QueueRate = ((Math.Abs(tmpBOQMMLocationChange) * 3600) / span.TotalSeconds);
                            clsGlobalVars.QueueChange = clsEnums.enQueueCahnge.Dissipating;
                        }
                        else if ((clsGlobalVars.PrevBOQMMLocation == -1) && (clsGlobalVars.BOQMMLocation != -1))
                        {
                            tmpBOQMMLocationChange = (clsGlobalVars.BOQMMLocation - Roadway.RecurringCongestionMMLocation);
                            clsGlobalVars.QueueRate = ((Math.Abs(tmpBOQMMLocationChange) * 3600) / span.TotalSeconds);
                            clsGlobalVars.QueueChange = clsEnums.enQueueCahnge.Growing;
                        }
                        else
                        {
                            clsGlobalVars.QueueChange = clsEnums.enQueueCahnge.Same;
                            clsGlobalVars.QueueRate = 0;
                        }
                        QueueLog.WriteLine(DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + "::" + DateTime.Now.Millisecond + ",CV," +
                                            clsGlobalVars.PrevBOQMMLocation.ToString("0.0") + "," + clsGlobalVars.PrevBOQTime.ToString() + "," +
                                            clsGlobalVars.BOQMMLocation.ToString("0.0") + "," + clsGlobalVars.BOQTime.ToString() + "," + clsGlobalVars.QueueRate.ToString("0.00") + "," +
                                            clsGlobalVars.QueueChange.ToString() + "," + clsGlobalVars.QueueSource.ToString());
                        if (clsGlobalVars.PrevBOQMMLocation != clsGlobalVars.BOQMMLocation)
                        {
                            clsGlobalVars.PrevBOQMMLocation = clsGlobalVars.BOQMMLocation;
                            clsGlobalVars.PrevBOQTime = clsGlobalVars.BOQTime;
                        }
                    }
                #endregion

                LogTxtMsg(txtCVDataLog, "\t\tPrev BOQ MM location: " + clsGlobalVars.PrevBOQMMLocation + "\tTime: " + clsGlobalVars.PrevBOQTime);
                CVDataProcessor.WriteLine("\tPrev BOQ MM location: " + clsGlobalVars.PrevBOQMMLocation + "\tTime: " + clsGlobalVars.PrevBOQTime);

                LogTxtMsg(txtCVDataLog, "\t\tCurr BOQ MM location: " + clsGlobalVars.BOQMMLocation + "\tTime: " + clsGlobalVars.BOQTime + "\tSource: " + clsGlobalVars.QueueSource.ToString());
                CVDataProcessor.WriteLine("\tCurr BOQ MM location: " + clsGlobalVars.BOQMMLocation + "\tTime: " + clsGlobalVars.BOQTime + "\tSource: " + clsGlobalVars.QueueSource.ToString());


                //Previous BOQ
                if (clsGlobalVars.PrevBOQMMLocation == -1)
                {
                    txtPrevBOQMMLocation.Text = "No Queue";
                    txtPrevBOQDate.Text = string.Empty;
                }
                else
                {
                    txtPrevBOQMMLocation.Text = clsGlobalVars.PrevBOQMMLocation.ToString();
                    txtPrevBOQDate.Text = clsGlobalVars.PrevBOQTime.ToString();
                }

                //Current BOQ
                if (clsGlobalVars.BOQMMLocation == -1)
                {
                    txtBOQMMLocation.Text = "No Queue";
                    txtBOQDate.Text = string.Empty;
                    DisplayForm.txtBOQ.Text = "No Queue";
                }
                else
                {
                    txtBOQMMLocation.Text = clsGlobalVars.BOQMMLocation.ToString();
                    txtBOQDate.Text = clsGlobalVars.BOQTime.ToString();
                    DisplayForm.txtBOQ.Text = clsGlobalVars.BOQMMLocation.ToString();
                }
                txtBOQGrowthRate.Text = clsGlobalVars.QueueRate.ToString(); ;
                txtBOQGrowthType.Text = clsGlobalVars.QueueChange.ToString();

                LogTxtMsg(txtCVDataLog, "\t\tCurr BOQ rate of growth: " + clsGlobalVars.QueueRate.ToString("0.00") + "\tQueue direction: " + clsGlobalVars.QueueChange.ToString() + "\tQueue source: " + clsGlobalVars.QueueSource.ToString());
                CVDataProcessor.WriteLine("\tCurr BOQ rate of growth: " + clsGlobalVars.QueueRate.ToString("0.00") + "\tQueue direction: " + clsGlobalVars.QueueChange.ToString() + "\tQueue source: " + clsGlobalVars.QueueSource.ToString());
                */
                #endregion



                #region "Reconcile the CV BOQ and TSS BOQ"
                TimeSpan BeginBOQ = new TimeSpan(DateTime.Now.Ticks);
                DetermineBOQ(QueuedSubLink, QueuedLink, Roadway);
                TimeSpan EndBOQ = new TimeSpan(DateTime.Now.Ticks);
                LogTxtMsg(txtCVDataLog, "\t\tTime for reconciling BOQ MM Location: \t" + (EndBOQ.TotalMilliseconds - BeginBOQ.TotalMilliseconds).ToString("0") + " msecs");
                #endregion

                //Start the sublink trouping process
                TimeSpan StartTrouping = new TimeSpan(DateTime.Now.Ticks);

                //Determine the Recommended speed for each sublink
                #region "Determine sublink recommended speed"
                foreach (clsRoadwayLink rl in RLList)
                {
                    foreach (clsRoadwaySubLink rsl in RSLList)
                    { 
                        if (Roadway.Direction == Roadway.MMIncreasingDirection)
                        {
                            if ((rsl.BeginMM >= rl.BeginMM) && (rsl.BeginMM < rl.EndMM))
                            {
                                rsl.TSSAvgSpeed = rl.TSSAvgSpeed;
                                rsl.WRTMSpeed = rl.WRTMSpeed;
                                rsl.RecommendedSpeed = GetMinimumSpeed(rsl.TSSAvgSpeed, rsl.WRTMSpeed, rsl.CVAvgSpeed);
                            }
                        }
                        else if (Roadway.Direction != Roadway.MMIncreasingDirection)
                        {
                            if ((rsl.BeginMM <= rl.BeginMM) && (rsl.BeginMM > rl.EndMM))
                            {
                                rsl.TSSAvgSpeed = rl.TSSAvgSpeed;
                                rsl.WRTMSpeed = rl.WRTMSpeed;
                                rsl.RecommendedSpeed = GetMinimumSpeed(rsl.TSSAvgSpeed, rsl.WRTMSpeed, rsl.CVAvgSpeed);
                            }
                        }
                    }
                }
                #endregion

                //Determine the ending MM for the trouping process
                #region "Determine the Trouping Ending MM Location"
                TroupingEndMM = 0;
                string Justification = "Normal";
                if (clsGlobalVars.BOQMMLocation == -1)
                {
                    //Locate the first sublink upstream of the Recurring congestion MMLocation that is congested 
                    for (int r = RSLList.Count - 1; r >= 0; r--)
                    {
                        if (Roadway.Direction == Roadway.MMIncreasingDirection)
                        {
                            if (RSLList[r].BeginMM < Roadway.RecurringCongestionMMLocation)
                            {
                                if ((RSLList[r].RecommendedSpeed > clsGlobalVars.LinkQueuedSpeedThreshold) && (RSLList[r].RecommendedSpeed <= clsGlobalVars.LinkCongestedSpeedThreshold))
                                {
                                    TroupingEndMM = RSLList[r].BeginMM;
                                    TroupingEndSpeed = RSLList[r].RecommendedSpeed;
                                    Justification = "Congestion";
                                    break;
                                }
                            }
                        }
                        else if (Roadway.Direction != Roadway.MMIncreasingDirection)
                        {
                            if (RSLList[r].BeginMM > Roadway.RecurringCongestionMMLocation)
                            {
                                if ((RSLList[r].RecommendedSpeed > clsGlobalVars.LinkQueuedSpeedThreshold) && (RSLList[r].RecommendedSpeed <= clsGlobalVars.LinkCongestedSpeedThreshold))
                                {
                                    TroupingEndMM = RSLList[r].BeginMM;
                                    TroupingEndSpeed = RSLList[r].RecommendedSpeed;
                                    Justification = "Congestion";
                                    break;
                                }
                            }
                        }
                    }
                }
                else if (clsGlobalVars.BOQMMLocation > 0)
                {
                    TroupingEndMM = clsGlobalVars.BOQMMLocation;
                    TroupingEndSpeed = clsGlobalVars.BOQSpeed;
                    Justification = "Queue";
                }
                #endregion

                DisplayForm.ClearCVSubLinkTroupeStatus();
                DisplayForm.ClearCVSubLinkSPDHarmStatus();
                DisplayForm.ClearTSSSPDHarmLinkStatus();

                string SpeedType = "Recommended";
                //perform the Trouping and Harmonization processes
                if (TroupingEndMM > 0)
                {
                    #region "Perform trouping and harmonization"
                    DisplayForm.DisableNoTroupingMessage();
                    DisplayForm.DisableNoSpdHarmMessage();
                    retValue = CalculateSublinkTroupeSpeed(ref RSLList, TroupingEndMM, TroupingEndSpeed);
                    if (retValue.Length > 0)
                    {
                        LogTxtMsg(txtCVDataLog, retValue);
                    }
                    retValue = CalculateSublinkHarmonizedSpeed(ref RSLList, TroupingEndMM, TroupingEndSpeed);
                    if (retValue.Length > 0)
                    {
                        LogTxtMsg(txtCVDataLog, retValue);
                    }

                    #region "Log Trouping and Harmonized sublink speeds"
                    CVDataProcessor.WriteLine("\r\n" + DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + "::" + DateTime.Now.Millisecond +
                                   "\tUpdated Roadway SubLink Recommended, Troupe, and Harmonized Speeds using CV data.");
                    CVDataProcessor.WriteLine("\r\n\tRoadwayId, SubLinkId, DateProcessed, BeginMM, EndMM, NumberQueuedCVs, TotalNumberCVS, PercentQueuedCVs, Queued, CVAvgSpd, TSSAvgSpd, RecommendedSpd, TroupeSpd, TroupePverride, HarmonizedSpeed, HarmonizationOverride, Direction");
                    foreach (clsRoadwaySubLink rsl in RSLList)
                    {
                        CVDataProcessor.WriteLine("\t" + rsl.RoadwayID + ", " + rsl.Identifier + ", " + rsl.DateProcessed + ", " + rsl.BeginMM + ", " + rsl.EndMM + ", " + rsl.NumberQueuedCVs + ", " +
                                                  rsl.TotalNumberCVs + ", " + rsl.PercentQueuedCVs + ", " + rsl.Queued + ", " + rsl.CVAvgSpeed.ToString("0") + ", " + rsl.TSSAvgSpeed.ToString("0") + ", " +
                                                  rsl.RecommendedSpeed.ToString("0") + ", " + rsl.TroupeSpeed.ToString("0") + ", " + rsl.TroupeInclusionOverride.ToString() + ", " +
                                                   rsl.HarmonizedSpeed.ToString() + ", " + rsl.SpdHarmInclusionOverride.ToString() + ", " + rsl.Direction);
                    }
                    #endregion

                    DisplayForm.DisplaySublinkTroupeSpeed(RSLList, TroupingEndMM, Roadway);
                    DisplayForm.DisplaySublinkHarmonizedSpeed(RSLList, TroupingEndMM, Roadway);
                    DetermineLinkHarmonizedSpeed(RSLList, ref RLList, TroupingEndMM);
                    DisplayForm.DisplayTSSSPDHarmLinkStatus(RLList);

                    SpeedType = "Harmonized";
                    LogTxtMsg(txtINFLOLog, "\t\tTrouping Required --Generating sublink Speed Messages");
                    GenerateSPDHarmMessages_Kittelson(RSLList, TroupingEndMM, clsGlobalVars.BOQMMLocation, Roadway, SpeedType, Justification);

                    #endregion
                }
                else
                {
                    #region "If no trouping is required"
                    LogTxtMsg(txtCVDataLog, "\t\tNo Trouping is required for sublinks because no queued or congested sublinks were detected");
                    DisplayForm.EnableNoTroupingMessage(Roadway.RecurringCongestionMMLocation);
                    DisplayForm.EnableNoSpdHarmMessage(Roadway.RecurringCongestionMMLocation);
                    LogTxtMsg(txtINFLOLog, "\t\tNo Trouping Required --Generating sublink Speed Messages");
                    GenerateSPDHarmMessages_Kittelson(RSLList, TroupingEndMM, clsGlobalVars.BOQMMLocation, Roadway, SpeedType, Justification);
                    #endregion
                }

                if (chkExcelDataLogging.Checked == true)
                {
                    Color myBlueColor = Color.FromArgb(128, 128, 255);
                    #region "Log Trouping and Harmonized sublink speed into Excel worksheet"
                    //Excel
                    int t = 0;
                    CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow, 1] = currDateTime;
                    CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow, 2] = "CVsSpd" + ":" + clsGlobalVars.BOQMMLocation + ":" + TroupingEndMM;
                    CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 1, 1] = currDateTime;
                    CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 1, 2] = "TSSSpd" + ":" + clsGlobalVars.BOQMMLocation + ":" + TroupingEndMM; ;
                    CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 2, 1] = currDateTime;
                    CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 2, 2] = "RecSpd" + ":" + clsGlobalVars.BOQMMLocation + ":" + TroupingEndMM; ;
                    CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 3, 1] = currDateTime;
                    CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 3, 2] = "TroSpd" + ":" + clsGlobalVars.BOQMMLocation + ":" + TroupingEndMM; ;
                    CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 4, 1] = currDateTime;
                    CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 4, 2] = "HarSpd" + ":" + clsGlobalVars.BOQMMLocation + ":" + TroupingEndMM; ;
                    foreach (clsRoadwaySubLink rsl in RSLList)
                    {
                        //Excel
                        CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow, t + 3] = rsl.CVAvgSpeed.ToString("0.00");
                        #region "CVSpeed Color"
                        if (rsl.CVAvgSpeed > 60.0)
                        {
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow, t + 3].Interior.Color = myBlueColor;
                        }
                        else if (rsl.CVAvgSpeed > 55.0)
                        {
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow, t + 3].Interior.Color = System.Drawing.Color.PaleTurquoise;
                        }
                        else if (rsl.CVAvgSpeed > 50.0)
                        {
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow, t + 3].Interior.Color = System.Drawing.Color.LimeGreen;
                        }
                        else if (rsl.CVAvgSpeed > 45.0)
                        {
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow, t + 3].Interior.Color = System.Drawing.Color.YellowGreen;
                        }
                        else if (rsl.CVAvgSpeed > 40.0)
                        {
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow, t + 3].Interior.Color = System.Drawing.Color.Yellow;
                        }
                        else if (rsl.CVAvgSpeed > 35.0)
                        {
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow, t + 3].Interior.Color = System.Drawing.Color.DarkOrange;
                        }
                        else if (rsl.CVAvgSpeed > 30.0)
                        {
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow, t + 3].Interior.Color = System.Drawing.Color.Tomato;
                        }
                        else if (rsl.CVAvgSpeed <= 30.0)
                        {
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow, t + 3].Interior.Color = System.Drawing.Color.Red;
                        }
                        #endregion
                        CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 1, t + 3] = rsl.TSSAvgSpeed.ToString("0.00");
                        #region "TSSSpeed Color"
                        if (rsl.TSSAvgSpeed > 60.0)
                        {
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 1, t + 3].Interior.Color = myBlueColor;
                        }
                        else if (rsl.TSSAvgSpeed > 55.0)
                        {
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 1, t + 3].Interior.Color = System.Drawing.Color.PaleTurquoise;
                        }
                        else if (rsl.TSSAvgSpeed > 50.0)
                        {
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 1, t + 3].Interior.Color = System.Drawing.Color.LimeGreen;
                        }
                        else if (rsl.TSSAvgSpeed > 45.0)
                        {
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 1, t + 3].Interior.Color = System.Drawing.Color.YellowGreen;
                        }
                        else if (rsl.TSSAvgSpeed > 40.0)
                        {
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 1, t + 3].Interior.Color = System.Drawing.Color.Yellow;
                        }
                        else if (rsl.TSSAvgSpeed > 35.0)
                        {
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 1, t + 3].Interior.Color = System.Drawing.Color.DarkOrange;
                        }
                        else if (rsl.TSSAvgSpeed > 30.0)
                        {
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 1, t + 3].Interior.Color = System.Drawing.Color.Tomato;
                        }
                        else if (rsl.TSSAvgSpeed <= 30.0)
                        {
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 1, t + 3].Interior.Color = System.Drawing.Color.Red;
                        }
                        #endregion
                        CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 2, t + 3] = rsl.RecommendedSpeed.ToString("0.00");
                        #region "RecommendedSpeed Color"
                        if (rsl.RecommendedSpeed > 60.0)
                        {
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 2, t + 3].Interior.Color = myBlueColor;
                        }
                        else if (rsl.RecommendedSpeed > 55.0)
                        {
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 2, t + 3].Interior.Color = System.Drawing.Color.PaleTurquoise;
                        }
                        else if (rsl.RecommendedSpeed > 50.0)
                        {
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 2, t + 3].Interior.Color = System.Drawing.Color.LimeGreen;
                        }
                        else if (rsl.RecommendedSpeed > 45.0)
                        {
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 2, t + 3].Interior.Color = System.Drawing.Color.YellowGreen;
                        }
                        else if (rsl.RecommendedSpeed > 40.0)
                        {
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 2, t + 3].Interior.Color = System.Drawing.Color.Yellow;
                        }
                        else if (rsl.RecommendedSpeed > 35.0)
                        {
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 2, t + 3].Interior.Color = System.Drawing.Color.DarkOrange;
                        }
                        else if (rsl.RecommendedSpeed > 30.0)
                        {
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 2, t + 3].Interior.Color = System.Drawing.Color.Tomato;
                        }
                        else if (rsl.RecommendedSpeed <= 30.0)
                        {
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 2, t + 3].Interior.Color = System.Drawing.Color.Red;
                        }
                        #endregion
                        CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 3, t + 3] = rsl.TroupeSpeed.ToString("0.00");
                        #region "Troupe Color"
                        CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 3, t + 3].Interior.Color = System.Drawing.Color.White;
                        if (rsl.TroupeInclusionOverride == true)
                        {
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 3, t + 3].Interior.Color = System.Drawing.Color.Silver;
                        }
                        if (rsl.BeginTroupe == true)
                        {
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 3, t + 3].Interior.Color = System.Drawing.Color.Pink;
                        }
                        #endregion
                        CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 4, t + 3] = rsl.HarmonizedSpeed.ToString("0.00");
                        #region "SpdHarm Color"
                        CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 4, t + 3].Interior.Color = System.Drawing.Color.White;
                        if (rsl.SpdHarmInclusionOverride == true)
                        {
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 4, t + 3].Interior.Color = System.Drawing.Color.Silver;
                        }
                        if (rsl.BeginSpdHarm == true)
                        {
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 4, t + 3].Interior.Color = System.Drawing.Color.Pink;
                        }
                        #endregion

                        t = t + 1;
                    }
                    //Excel
                    CVSPDHarmWSCurrRow = CVSPDHarmWSCurrRow + 6;
                    #endregion
                }
                TimeSpan EndTrouping = new TimeSpan(DateTime.Now.Ticks);
                LogTxtMsg(txtCVDataLog, "\t\tTime for sublink speed harmonization: \t" + (EndTrouping.TotalMilliseconds - StartTrouping.TotalMilliseconds).ToString("0") + " msecs");

                #endregion
            }
            else
            {
                #region "If NO CV data was retrieved for the last five seconds"
                NumberNoCVDataIntervals = NumberNoCVDataIntervals + 1;
                LogTxtMsg(txtINFLOLog, "\r\n\t\tNo CV traffic data was retrieved for the last " + clsGlobalVars.CVDataPollingFrequency + " second interval. " +
                                        "\r\n\t\tNumber of " + clsGlobalVars.CVDataPollingFrequency + " second intervals with no CV data retrieved: \t" + NumberNoCVDataIntervals);
                LogTxtMsg(txtCVDataLog, "\t\tNo CV traffic data was retrieved for the last " + clsGlobalVars.CVDataPollingFrequency + " second interval. " +
                                        "\r\n\t\tNumber of " + clsGlobalVars.CVDataPollingFrequency + " second intervals with no CV data retrieved: \t" + NumberNoCVDataIntervals);
                DisplayForm.txtCVDate.Text = DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second;
                if (NumberNoCVDataIntervals > clsGlobalVars.CVDataSmoothedSpeedArraySize)
                {
                    if (clsGlobalVars.CVBOQMMLocation != -1)
                    {
                        clsGlobalVars.CVBOQMMLocation = -1;
                        clsGlobalVars.CVBOQTime = DateTime.Now;
                    }
                    foreach (clsRoadwaySubLink RSL in RSLList)
                    {
                        RSL.CVAvgSpeed = clsGlobalVars.MaximumDisplaySpeed;
                        RSL.Queued = false;
                        RSL.Congested = false;
                        //RSL.HarmonizedSpeed = clsGlobalVars.MaximumDisplaySpeed;
                        //RSL.TroupeSpeed = clsGlobalVars.MaximumDisplaySpeed;
                    }
                    DisplayForm.EnableNoCVDataMessage(NumberNoCVDataIntervals);
                    DisplayForm.txtCVBOQ.Text = "No Queue";
                    DisplayForm.ClearCVSubLinkQueuedStatus();

                    if (NumberNoTSSDataIntervals > 0)
                    {
                        DisplayForm.txtBOQ.Text = "No Queue";
                        DisplayForm.txtCVBOQ.Text = "No Queue";
                        DisplayForm.txtTSSBOQ.Text = "No Queue";
                        DisplayForm.txtQueueGrowthRate.Text = "";
                        DisplayForm.txtQueueSpeed.Text = "";
                        DisplayForm.txtQueueLength.Text = "";

                        DisplayForm.ClearCVSubLinkSPDHarmStatus();
                        DisplayForm.ClearCVSubLinkTroupeStatus();
                        DisplayForm.ClearTSSSPDHarmLinkStatus();
                        //Reset Troupe speed, and inclusion flag for every sublink
                        foreach (clsRoadwaySubLink rsl in RSLList)
                        {
                            rsl.TroupeSpeed = 0;
                            rsl.TroupeInclusionOverride = false;
                            rsl.BeginTroupe = false;
                            rsl.TroupeProcessed = DateTime.Now;
                            rsl.HarmonizedSpeed = 0;
                            rsl.SpdHarmInclusionOverride = false;
                            rsl.BeginSpdHarm = false;
                        }
                        DisplayForm.EnableNoSpdHarmMessage(Roadway.RecurringCongestionMMLocation);
                        DisplayForm.EnableNoTroupingMessage(Roadway.RecurringCongestionMMLocation);

                        if (chkExcelDataLogging.Checked == true)
                        {
                            Color myBlueColor = Color.FromArgb(128, 128, 255);
                            #region "Log Trouping and Harmonized sublink speed into Excel worksheet"
                        //Excel
                            int t = 0;
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow, 1] = currDateTime;
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow, 2] = "CVsSpd" + ":" + clsGlobalVars.BOQMMLocation + ":" + TroupingEndMM;
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 1, 1] = currDateTime;
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 1, 2] = "TSSSpd" + ":" + clsGlobalVars.BOQMMLocation + ":" + TroupingEndMM; ;
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 2, 1] = currDateTime;
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 2, 2] = "RecSpd" + ":" + clsGlobalVars.BOQMMLocation + ":" + TroupingEndMM; ;
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 3, 1] = currDateTime;
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 3, 2] = "TroSpd" + ":" + clsGlobalVars.BOQMMLocation + ":" + TroupingEndMM; ;
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 4, 1] = currDateTime;
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 4, 2] = "HarSpd" + ":" + clsGlobalVars.BOQMMLocation + ":" + TroupingEndMM; ;
                            foreach (clsRoadwaySubLink rsl in RSLList)
                            {
                                //Excel
                                CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow, t + 3] = rsl.CVAvgSpeed.ToString("0.00");
                                #region "CVSpeed Color"
                                if (rsl.CVAvgSpeed > 60.0)
                                {
                                    CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow, t + 3].Interior.Color = myBlueColor;
                                }
                                else if (rsl.CVAvgSpeed > 55.0)
                                {
                                    CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow, t + 3].Interior.Color = System.Drawing.Color.PaleTurquoise;
                                }
                                else if (rsl.CVAvgSpeed > 50.0)
                                {
                                    CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow, t + 3].Interior.Color = System.Drawing.Color.LimeGreen;
                                }
                                else if (rsl.CVAvgSpeed > 45.0)
                                {
                                    CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow, t + 3].Interior.Color = System.Drawing.Color.YellowGreen;
                                }
                                else if (rsl.CVAvgSpeed >= 40.0)
                                {
                                    CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow, t + 3].Interior.Color = System.Drawing.Color.Yellow;
                                }
                                else if (rsl.CVAvgSpeed > 35.0)
                                {
                                    CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow, t + 3].Interior.Color = System.Drawing.Color.DarkOrange;
                                }
                                else if (rsl.CVAvgSpeed > 30.0)
                                {
                                    CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow, t + 3].Interior.Color = System.Drawing.Color.Tomato;
                                }
                                else if (rsl.CVAvgSpeed <= 30.0)
                                {
                                    CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow, t + 3].Interior.Color = System.Drawing.Color.Red;
                                }
                                #endregion

                                CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 1, t + 3] = rsl.TSSAvgSpeed.ToString("0.00");
                                #region "TSSSpeed Color"
                                if (rsl.TSSAvgSpeed > 60.0)
                                {
                                    CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 1, t + 3].Interior.Color = myBlueColor;
                                }
                                else if (rsl.TSSAvgSpeed > 55.0)
                                {
                                    CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 1, t + 3].Interior.Color = System.Drawing.Color.PaleTurquoise;
                                }
                                else if (rsl.TSSAvgSpeed > 50.0)
                                {
                                    CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 1, t + 3].Interior.Color = System.Drawing.Color.LimeGreen;
                                }
                                else if (rsl.TSSAvgSpeed > 45.0)
                                {
                                    CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 1, t + 3].Interior.Color = System.Drawing.Color.YellowGreen;
                                }
                                else if (rsl.TSSAvgSpeed > 40.0)
                                {
                                    CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 1, t + 3].Interior.Color = System.Drawing.Color.Yellow;
                                }
                                else if (rsl.TSSAvgSpeed > 35.0)
                                {
                                    CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 1, t + 3].Interior.Color = System.Drawing.Color.DarkOrange;
                                }
                                else if (rsl.TSSAvgSpeed > 30.0)
                                {
                                    CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 1, t + 3].Interior.Color = System.Drawing.Color.Tomato;
                                }
                                else if (rsl.TSSAvgSpeed <= 30.0)
                                {
                                    CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 1, t + 3].Interior.Color = System.Drawing.Color.Red;
                                }
                                #endregion

                                CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 2, t + 3] = rsl.RecommendedSpeed.ToString("0.00");
                                #region "RecommendedSpeed Color"
                                if (rsl.RecommendedSpeed > 60.0)
                                {
                                    CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 2, t + 3].Interior.Color = myBlueColor;
                                }
                                else if (rsl.RecommendedSpeed > 55.0)
                                {
                                    CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 2, t + 3].Interior.Color = System.Drawing.Color.PaleTurquoise;
                                }
                                else if (rsl.RecommendedSpeed > 50.0)
                                {
                                    CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 2, t + 3].Interior.Color = System.Drawing.Color.LimeGreen;
                                }
                                else if (rsl.RecommendedSpeed > 45.0)
                                {
                                    CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 2, t + 3].Interior.Color = System.Drawing.Color.YellowGreen;
                                }
                                else if (rsl.RecommendedSpeed > 40.0)
                                {
                                    CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 2, t + 3].Interior.Color = System.Drawing.Color.Yellow;
                                }
                                else if (rsl.RecommendedSpeed > 35.0)
                                {
                                    CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 2, t + 3].Interior.Color = System.Drawing.Color.DarkOrange;
                                }
                                else if (rsl.RecommendedSpeed > 30.0)
                                {
                                    CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 2, t + 3].Interior.Color = System.Drawing.Color.Tomato;
                                }
                                else if (rsl.RecommendedSpeed <= 30.0)
                                {
                                    CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 2, t + 3].Interior.Color = System.Drawing.Color.Red;
                                }
                                #endregion

                                CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 3, t + 3] = rsl.TroupeSpeed.ToString("0.00");
                                #region "Troupe Color"
                                CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 3, t + 3].Interior.Color = System.Drawing.Color.White;
                                if (rsl.TroupeInclusionOverride == true)
                                {
                                    CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 3, t + 3].Interior.Color = System.Drawing.Color.Silver;
                                }
                                if (rsl.BeginTroupe == true)
                                {
                                    CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 3, t + 3].Interior.Color = System.Drawing.Color.Pink;
                                }
                                #endregion

                                CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 4, t + 3] = rsl.HarmonizedSpeed.ToString("0.00");
                                #region "SpdHarm Color"
                                CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 4, t + 3].Interior.Color = System.Drawing.Color.White;
                                if (rsl.SpdHarmInclusionOverride == true)
                                {
                                    CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 4, t + 3].Interior.Color = System.Drawing.Color.Silver;
                                }
                                if (rsl.BeginSpdHarm == true)
                                {
                                    CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 4, t + 3].Interior.Color = System.Drawing.Color.Pink;
                                }
                                #endregion

                                t = t + 1;
                            }
                            //Excel
                            CVSPDHarmWSCurrRow = CVSPDHarmWSCurrRow + 6;
                            #endregion
                        }
                    }
                    //else
                    //{
                    //    DisplayForm.DisableNoSpdHarmMessage();
                    //    DisplayForm.DisableNoTroupingMessage();
                    //}
                }
                #endregion
            }

            System.Windows.Forms.Application.DoEvents();
            System.Windows.Forms.Application.DoEvents();
            System.Windows.Forms.Application.DoEvents();

            DateTime CVEndWakeupTime = DateTime.Now;
            TimeSpan WakeupTime = CVEndWakeupTime.Subtract(CVCurrWakeupTime);
            if (WakeupTime.TotalMilliseconds < (clsGlobalVars.CVDataPollingFrequency * 1000))
            {
                tmrCVData.Interval = (int)(Math.Ceiling((clsGlobalVars.CVDataPollingFrequency * 1000) - WakeupTime.TotalMilliseconds));
                LogTxtMsg (txtCVDataLog, "\t\tTmrCVData Next sleep interval = " + tmrCVData.Interval);
            }
            else
            {
                tmrCVData.Interval = 100;
                LogTxtMsg (txtCVDataLog, "\t\tTmrCVData Next sleep interval = " + tmrCVData.Interval);
            }

            LogTxtMsg(txtCVDataLog, "\r\n\r\n" + DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + "::" + DateTime.Now.Millisecond +
                                    "\tFinished Processing CV traffic data for the last: " + clsGlobalVars.CVDataPollingFrequency + " seconds interval.");
            CVDataProcessor.WriteLine("\r\n\r\n" + DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + "::" + DateTime.Now.Millisecond +
                                      "\tFinished Processing CV traffic data for the last: " + clsGlobalVars.CVDataPollingFrequency + " seconds interval.");

            if (Stopped == false)
            {
                //tmrCVData.Enabled = true;
            }
        }

        private void tmrTSSData_Tick(object sender, EventArgs e)
        {
        }
        private void ProcessTSSData()
        {
            string retValue = string.Empty;
            string CurrentSection = string.Empty;

            DateTime currDateTime = DateTime.Now;
            currDateTime = TimeZoneInfo.ConvertTimeToUtc(currDateTime, TimeZoneInfo.Local);


            LogTxtMsg(txtTSSDataLog, "\r\n\r\n------------------------------");
            LogTxtMsg(txtTSSDataLog, DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + "::" + DateTime.Now.Millisecond +
                                     "\tGet TSS detection zone traffic data for the last " + clsGlobalVars.TSSDataLoadingFrequency + " seconds interval from: " +
                                     currDateTime.AddSeconds(-clsGlobalVars.TSSDataLoadingFrequency) + "\tTo: " + currDateTime);
            TSSDataProcessor.WriteLine(DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + "::" + DateTime.Now.Millisecond +
                                     "\tGet TSS detection zone traffic data for the last " + clsGlobalVars.TSSDataLoadingFrequency + " seconds interval from: " +
                                     currDateTime.AddSeconds(-clsGlobalVars.TSSDataLoadingFrequency) + "\tTo: " + currDateTime);

            //Get Last interval TSS data
            int NumberRecordsRetrieved = 0;

            //Commented by Hassan Charara on 07/01/2014 because Kittelson does not provide detection zone data in the database
            #region "Get last interval TSS detection zone data"
            //TimeSpan GetTSSDataTime = new TimeSpan(DateTime.Now.Ticks);
            //retValue = GetLastIntervalDetectionZoneStatus(DB, currDateTime, ref DZList, ref NumberRecordsRetrieved);
            //if (retValue.Length > 0)
            //{
            //    LogTxtMsg(txtTSSDataLog, retValue);
            //    TSSDataProcessor.WriteLine("\t" + retValue);
            //}
            //TimeSpan EndGetTSSDataTime = new TimeSpan(DateTime.Now.Ticks);
            #endregion


            //Commented by Hassan Charara on 07/01/2014 to read directly from the database the TSS detection station data generated by Kittelson
            #region "Get last interval TSS data"
            TimeSpan GetTSSDataTime = new TimeSpan(DateTime.Now.Ticks);
            retValue = GetLastIntervalDetectionStationData(DB, ref DSList, ref NumberRecordsRetrieved);
            if (retValue.Length > 0)
            {
                LogTxtMsg(txtTSSDataLog, retValue);
                TSSDataProcessor.WriteLine("\t" + retValue);
            }
            TimeSpan EndGetTSSDataTime = new TimeSpan(DateTime.Now.Ticks);
            #endregion

            if (NumberRecordsRetrieved > 0)
            {
                LogTxtMsg(txtTSSDataLog, "\t\tTime for retrieving " + DSList.Count + " detector station traffic data records from INFLO database: " + (EndGetTSSDataTime.TotalMilliseconds - GetTSSDataTime.TotalMilliseconds).ToString("0") + " msecs");
                LogTxtMsg(txtINFLOLog, "\t\tTime for retrieving " + DSList.Count + " detector station traffic data records from INFLO database: " + (EndGetTSSDataTime.TotalMilliseconds - GetTSSDataTime.TotalMilliseconds).ToString("0") + " msecs");
                TSSDataProcessor.WriteLine("\tTime for retrieving " + DSList.Count + " detector station traffic data records from INFLO database: \t " + (EndGetTSSDataTime.TotalMilliseconds - GetTSSDataTime.TotalMilliseconds).ToString("0") + " msecs");
                
                #region "Process detection station traffic data"
                //TimeSpan StartProcessingDSStatus = new TimeSpan(DateTime.Now.Ticks);
                //LogTxtMsg(txtTSSDataLog, "\r\n\t\tCalculate detector station traffic parameters (average speed, occupancy and volume) using detection zone traffic data.");
                //TSSDataProcessor.WriteLine("\r\n\tCalculate detector station traffic parameters (average speed, occupancy and volume) using detection zone traffic data.");
                //retValue = ProcessDetectorStationStatus();
                //TimeSpan EndProcessingDSStatus = new TimeSpan(DateTime.Now.Ticks);
                //if (retValue.Length > 0)
                //{
                //    LogTxtMsg(txtTSSDataLog, retValue);
                //    TSSDataProcessor.WriteLine("\t" + retValue);
                //}
                //else
                //{
                //    LogTxtMsg(txtTSSDataLog, "\t\tTime for calculating TSS detector station traffic parameters: " + EndProcessingDSStatus.Subtract(StartProcessingDSStatus).TotalMilliseconds + " milliseconds");
                //    TSSDataProcessor.WriteLine("\t\tTime for calculating TSS detector station traffic parameters: " + EndProcessingDSStatus.Subtract(StartProcessingDSStatus).TotalMilliseconds + " milliseconds");
                //}
                #endregion

                #region "Process Roadway Link traffic data"
                TimeSpan StartProcessingRLStatus = new TimeSpan(DateTime.Now.Ticks);
                LogTxtMsg(txtTSSDataLog, "\r\n\t\tDetermine roadway link average speed, occupancy, volume and queued state.");
                TSSDataProcessor.WriteLine("\r\n\tDetermine roadway link average speed, occupancy, volume and queued state.");
                retValue = ProcessLinkInfrastructureStatus();
                TimeSpan EndProcessingRLStatus = new TimeSpan(DateTime.Now.Ticks);
                if (retValue.Length > 0)
                {
                    LogTxtMsg(txtTSSDataLog, retValue);
                    TSSDataProcessor.WriteLine(retValue);
                }
                else
                {
                    LogTxtMsg(txtTSSDataLog, "\r\n\tTime for setting TSS Link traffic parameters: " + EndProcessingRLStatus.Subtract(StartProcessingRLStatus).TotalMilliseconds + " milliseconds");
                    TSSDataProcessor.WriteLine("\r\n\tTime for setting TSS Link traffic parameters: " + EndProcessingRLStatus.Subtract(StartProcessingRLStatus).TotalMilliseconds + " milliseconds");
                }
                #endregion

                if (chkExcelDataLogging.Checked == true)
                {
                    //Excel
                    TSSWorkSheets[1].Cells[TSSWSCurrRow, 1] = currDateTime;
                    for (int s = 0; s < RLList.Count; s++)
                    {
                        Excel.Range tmpRange = (Excel.Range)TSSWorkSheets[1].Range[TSSWorkSheets[1].Cells[TSSWSCurrRow, (s * 5) + 4], TSSWorkSheets[1].Cells[TSSWSCurrRow, (s + 1) * 5 + 3]];
                        //tmpRange = tmpRange.Select();
                        //tmpRange.ReadingOrder
                        //tmpRange.Orientation = 0;
                        //tmpRange.AddIndent = false;
                        //tmpRange.IndentLevel = 0;
                        tmpRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        tmpRange.MergeCells = true;
                        TSSWorkSheets[1].Cells[TSSWSCurrRow, (s * 5) + 4] = RLList[s].TSSAvgSpeed.ToString("0") + "-" + RLList[s].Queued;
                        if (RLList[s].Queued == true)
                        {
                            TSSWorkSheets[1].Cells[TSSWSCurrRow, (s * 5) + 4].Interior.Color = System.Drawing.Color.Red;
                        }

                        TSSWorkSheets[1].Cells[TSSWSCurrRow, 2] = RLList[s].StartInterval;
                        TSSWorkSheets[1].Cells[TSSWSCurrRow, 3] = RLList[s].EndInterval;
                    }
                    TSSWorkSheets[1].Range[TSSWorkSheets[1].Cells[TSSWSCurrRow, 1], TSSWorkSheets[1].Cells[TSSWSCurrRow, (RLList.Count * 5) + 3]].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    TSSWorkSheets[1].Range[TSSWorkSheets[1].Cells[TSSWSCurrRow, 1], TSSWorkSheets[1].Cells[TSSWSCurrRow, (RLList.Count * 5) + 3]].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick;
                    TSSWSCurrRow = TSSWSCurrRow + 1;
                }

                TimeSpan endtime = new TimeSpan(DateTime.Now.Ticks);
                LogTxtMsg(txtTSSDataLog, "\t\tTime for processing: TSS records for the last " + clsGlobalVars.TSSDataLoadingFrequency + " seconds interval: " + (endtime.TotalMilliseconds - EndGetTSSDataTime.TotalMilliseconds).ToString("0") + " msecs");

                /*LogTxtMsg(txtTSSDataLog, "\r\n\tLoad roadway link traffic data into INFLO database.");
                retValue = InsertLinkStatusIntoINFLODatabase();
                if (retValue.Length > 0)
                {
                    LogTxtMsg(txtTSSDataLog, retValue);
                }
                else
                {
                    LogTxtMsg(txtTSSDataLog, "\r\n\tFinish loading TSS roadway links traffic data into INFLO database");
                }*/
                
                //Determine TSS link BOQ location
                TimeSpan BeginTSSBOQ = new TimeSpan(DateTime.Now.Ticks);
                LogTxtMsg(txtTSSDataLog, "\r\n\t\tDetermine TSS BOQ location:");
                TSSDataProcessor.WriteLine("\r\n" + DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + "::" + DateTime.Now.Millisecond + "\tDetermine CV SubLink BOQ location:");

                clsGlobalVars.PrevInfrastructureBOQMMLocation = clsGlobalVars.InfrastructureBOQMMLocation;
                clsGlobalVars.PrevInfrastructureBOQTime = clsGlobalVars.InfrastructureBOQTime;
                clsGlobalVars.InfrastructureBOQMMLocation = -1;

                #region "Determine TSS BOQ MM Location"
                QueuedLink = new clsRoadwayLink();
                RLList.Sort((l, r) => l.BeginMM.CompareTo(r.BeginMM));
                foreach (clsRoadwayLink RL in RLList)
                {
                    if (RL.Queued == true)
                    {
                        if (Roadway.MMIncreasingDirection == Roadway.Direction)
                        {
                            if (RL.BeginMM < Roadway.RecurringCongestionMMLocation)
                            {
                                if (clsGlobalVars.InfrastructureBOQMMLocation != -1)
                                {
                                    if (RL.BeginMM < clsGlobalVars.InfrastructureBOQMMLocation)
                                    {
                                        clsGlobalVars.InfrastructureBOQMMLocation = RL.BeginMM;
                                        clsGlobalVars.InfrastructureBOQTime = DateTime.Now;
                                        clsGlobalVars.InfrastructureBOQLinkSpeed = RL.TSSAvgSpeed;
                                        //QueuedLink = new clsRoadwayLink();
                                        QueuedLink = RL;
                                        //LogTxtMsg(txtTSSDataLog, "\t\t\tTSS BOQ MM Location: " + clsGlobalVars.InfrastructureBOQMMLocation);
                                    }
                                }
                                else
                                {
                                    clsGlobalVars.InfrastructureBOQMMLocation = RL.BeginMM;
                                    clsGlobalVars.InfrastructureBOQTime = DateTime.Now;
                                    clsGlobalVars.InfrastructureBOQLinkSpeed = RL.TSSAvgSpeed;
                                    //QueuedLink = new clsRoadwayLink();
                                    QueuedLink = RL;
                                    //LogTxtMsg(txtTSSDataLog, "\t\t\tTSS BOQ MM Location: " + clsGlobalVars.InfrastructureBOQMMLocation);
                                }
                            }
                        }
                        else if (Roadway.MMIncreasingDirection != Roadway.Direction)
                        {
                            if (RL.BeginMM > Roadway.RecurringCongestionMMLocation)
                            {
                                if (clsGlobalVars.InfrastructureBOQMMLocation != -1)
                                {
                                    if (RL.BeginMM > clsGlobalVars.InfrastructureBOQMMLocation)
                                    {
                                        clsGlobalVars.InfrastructureBOQMMLocation = RL.BeginMM;
                                        clsGlobalVars.InfrastructureBOQTime = DateTime.Now;
                                        clsGlobalVars.InfrastructureBOQLinkSpeed = RL.TSSAvgSpeed;
                                        //QueuedLink = new clsRoadwayLink();
                                        QueuedLink = RL;
                                        //LogTxtMsg(txtTSSDataLog, "\t\t\tTSS BOQ MM Location: " + clsGlobalVars.InfrastructureBOQMMLocation);
                                    }
                                }
                                else
                                {
                                    clsGlobalVars.InfrastructureBOQMMLocation = RL.BeginMM;
                                    clsGlobalVars.InfrastructureBOQTime = DateTime.Now;
                                    clsGlobalVars.InfrastructureBOQLinkSpeed = RL.TSSAvgSpeed;
                                    //QueuedLink = new clsRoadwayLink();
                                    QueuedLink = RL;
                                    //LogTxtMsg(txtTSSDataLog, "\t\t\tTSS BOQ MM Location: " + clsGlobalVars.InfrastructureBOQMMLocation);
                                }
                            }
                        }
                    }
                }
                #endregion

                DisplayForm.txtTSSDate.Text = DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second;
                DisplayForm.ClearTSSQueuedLinkStatus();
                DisplayForm.DisplayTSSQueuedLinkStatus(RLList);

                LogTxtMsg(txtTSSDataLog, "\t\tPrev TSS BOQ MM location: " + clsGlobalVars.PrevInfrastructureBOQMMLocation + "\tTime: " + clsGlobalVars.PrevInfrastructureBOQTime);
                LogTxtMsg(txtTSSDataLog, "\t\tCurr TSS BOQ MM location: " + clsGlobalVars.InfrastructureBOQMMLocation + "\tTime: " + clsGlobalVars.InfrastructureBOQTime);
                TSSDataProcessor.WriteLine("\t\tPrev TSS BOQ MM location: " + clsGlobalVars.PrevInfrastructureBOQMMLocation + "\tTime: " + clsGlobalVars.PrevInfrastructureBOQTime);
                TSSDataProcessor.WriteLine("\t\tCurr TSS BOQ MM location: " + clsGlobalVars.InfrastructureBOQMMLocation + "\tTime: " + clsGlobalVars.InfrastructureBOQTime);

                //Infrastructure BOQ
                if (clsGlobalVars.InfrastructureBOQMMLocation == -1)
                {
                    txtTSSBOQMMLocation1.Text = "No Queue";
                    txtTSSBOQDate1.Text = string.Empty;
                    DisplayForm.txtTSSBOQ.Text = "No Queue";
                }
                else
                {
                    txtTSSBOQMMLocation1.Text = clsGlobalVars.InfrastructureBOQMMLocation.ToString();
                    txtTSSBOQDate1.Text = clsGlobalVars.InfrastructureBOQTime.ToString();
                    DisplayForm.txtTSSBOQ.Text = clsGlobalVars.InfrastructureBOQMMLocation.ToString();
                }
                //Previous Infrastructure BOQ
                if (clsGlobalVars.PrevInfrastructureBOQMMLocation == -1)
                {
                    txtTSSPrevBOQMMLocation1.Text = "No Queue";
                    txtTSSPrevBOQDate1.Text = string.Empty;
                }
                else
                {
                    txtTSSPrevBOQMMLocation1.Text = clsGlobalVars.PrevInfrastructureBOQMMLocation.ToString();
                    txtTSSPrevBOQDate1.Text = clsGlobalVars.PrevInfrastructureBOQTime.ToString();
                }

                TimeSpan EndTSSBOQ = new TimeSpan(DateTime.Now.Ticks);
                LogTxtMsg(txtTSSDataLog, "\t\tTime for processing  TSS BOQ MM Location: \t" + (EndTSSBOQ.TotalMilliseconds - BeginTSSBOQ.TotalMilliseconds).ToString("0") + " msecs");


                if (NumberNoCVDataIntervals > 0)
                {
                    string SpeedType = "Recommended";

                    DisplayForm.ClearCVSubLinkTroupeStatus();
                    DisplayForm.ClearCVSubLinkSPDHarmStatus();
                    DisplayForm.ClearTSSSPDHarmLinkStatus();
                    //Reset Troupe speed, and inclusion flag for every sublink
                    foreach (clsRoadwaySubLink rsl in RSLList)
                    {
                        rsl.TroupeSpeed = 0;
                        rsl.TroupeInclusionOverride = false;
                        rsl.BeginTroupe = false;
                        rsl.TroupeProcessed = DateTime.Now;
                        rsl.HarmonizedSpeed = 0;
                        rsl.SpdHarmInclusionOverride = false;
                        rsl.BeginSpdHarm = false;
                    }
                    #region "Reconcile the CV BOQ and TSS BOQ"
                    TimeSpan BeginBOQ = new TimeSpan(DateTime.Now.Ticks);
                    DetermineBOQ(QueuedSubLink, QueuedLink, Roadway);
                    TimeSpan EndBOQ = new TimeSpan(DateTime.Now.Ticks);
                    LogTxtMsg(txtTSSDataLog, "\t\tTime for reconciling BOQ MM Location: \t" + (EndBOQ.TotalMilliseconds - BeginBOQ.TotalMilliseconds).ToString("0") + " msecs");
                    #endregion

                    //Start the sublink trouping process
                    TimeSpan StartTrouping = new TimeSpan(DateTime.Now.Ticks);

                    //Determine the Recommended speed for each sublink
                    #region "Determine sublink recommended speed"
                    foreach (clsRoadwayLink rl in RLList)
                    {
                        foreach (clsRoadwaySubLink rsl in RSLList)
                        {
                            if (Roadway.Direction == Roadway.MMIncreasingDirection)
                            {
                                if ((rsl.BeginMM >= rl.BeginMM) && (rsl.BeginMM < rl.EndMM))
                                {
                                    rsl.TSSAvgSpeed = rl.TSSAvgSpeed;
                                    rsl.WRTMSpeed = rl.WRTMSpeed;
                                    rsl.RecommendedSpeed = GetMinimumSpeed(rsl.TSSAvgSpeed, rsl.WRTMSpeed, rsl.CVAvgSpeed);
                                }
                            }
                            else if (Roadway.Direction != Roadway.MMIncreasingDirection)
                            {
                                if ((rsl.BeginMM <= rl.BeginMM) && (rsl.BeginMM > rl.EndMM))
                                {
                                    rsl.TSSAvgSpeed = rl.TSSAvgSpeed;
                                    rsl.WRTMSpeed = rl.WRTMSpeed;
                                    rsl.RecommendedSpeed = GetMinimumSpeed(rsl.TSSAvgSpeed, rsl.WRTMSpeed, rsl.CVAvgSpeed);
                                }
                            }
                        }
                    }
                    #endregion

                    //Determine the ending MM for the trouping process
                    #region "Determine the Trouping Ending MM Location"
                    TroupingEndMM = 0;
                    string Justification = "Normal";
                    if (clsGlobalVars.BOQMMLocation == -1)
                    {
                        //Locate the first sublink upstream of the Recurring congestion MMLocation that is congested 
                        for (int r = RSLList.Count - 1; r >= 0; r--)
                        {
                            if (Roadway.Direction == Roadway.MMIncreasingDirection)
                            {
                                if (RSLList[r].BeginMM < Roadway.RecurringCongestionMMLocation)
                                {
                                    if ((RSLList[r].RecommendedSpeed > clsGlobalVars.LinkQueuedSpeedThreshold) && (RSLList[r].RecommendedSpeed <= clsGlobalVars.LinkCongestedSpeedThreshold))
                                    {
                                        TroupingEndMM = RSLList[r].BeginMM;
                                        TroupingEndSpeed = RSLList[r].RecommendedSpeed;
                                        Justification = "Congestion";
                                        break;
                                    }
                                }
                            }
                            else if (Roadway.Direction != Roadway.MMIncreasingDirection)
                            {
                                if (RSLList[r].BeginMM > Roadway.RecurringCongestionMMLocation)
                                {
                                    if ((RSLList[r].RecommendedSpeed > clsGlobalVars.LinkQueuedSpeedThreshold) && (RSLList[r].RecommendedSpeed <= clsGlobalVars.LinkCongestedSpeedThreshold))
                                    {
                                        TroupingEndMM = RSLList[r].BeginMM;
                                        TroupingEndSpeed = RSLList[r].RecommendedSpeed;
                                        Justification = "Congestion";
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    else if (clsGlobalVars.BOQMMLocation > 0)
                    {
                        TroupingEndMM = clsGlobalVars.BOQMMLocation;
                        TroupingEndSpeed = clsGlobalVars.BOQSpeed;
                        Justification = "Queue";
                    }
                    #endregion

                    //perform the Trouping and Harmonization processes
                    if (TroupingEndMM > 0)
                    {
                        #region "Perform trouping and harmonization"
                        DisplayForm.DisableNoTroupingMessage();
                        DisplayForm.DisableNoSpdHarmMessage();
                        retValue = CalculateSublinkTroupeSpeed(ref RSLList, TroupingEndMM, TroupingEndSpeed);
                        if (retValue.Length > 0)
                        {
                            LogTxtMsg(txtTSSDataLog, retValue);
                        }
                        retValue = CalculateSublinkHarmonizedSpeed(ref RSLList, TroupingEndMM, TroupingEndSpeed);
                        if (retValue.Length > 0)
                        {
                            LogTxtMsg(txtTSSDataLog, retValue);
                        }
                        TSSDataProcessor.WriteLine("\r\n" + DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + "::" + DateTime.Now.Millisecond +
                                       "\tUpdated Roadway SubLink Recommended, Troupe, and Harmonized Speeds using CV data.");
                        TSSDataProcessor.WriteLine("\r\n\tRoadwayId, SubLinkId, DateProcessed, BeginMM, EndMM, NumberQueuedCVs, TotalNumberCVS, PercentQueuedCVs, Queued, CVAvgSpd, TSSAvgSpd, RecommendedSpd, TroupeSpd, TroupePverride, HarmonizedSpeed, HarmonizationOverride, Direction");
                        foreach (clsRoadwaySubLink rsl in RSLList)
                        {
                            TSSDataProcessor.WriteLine("\t" + rsl.RoadwayID + ", " + rsl.Identifier + ", " + rsl.DateProcessed + ", " + rsl.BeginMM + ", " + rsl.EndMM + ", " + rsl.NumberQueuedCVs + ", " +
                                                      rsl.TotalNumberCVs + ", " + rsl.PercentQueuedCVs + ", " + rsl.Queued + ", " + rsl.CVAvgSpeed.ToString("0") + ", " + rsl.TSSAvgSpeed.ToString("0") + ", " +
                                                      rsl.RecommendedSpeed.ToString("0") + ", " + rsl.TroupeSpeed.ToString("0") + ", " + rsl.TroupeInclusionOverride.ToString() + ", " +
                                                       rsl.HarmonizedSpeed.ToString() + ", " + rsl.SpdHarmInclusionOverride.ToString() + ", " + rsl.Direction);
                        }
                        DisplayForm.DisplaySublinkTroupeSpeed(RSLList, TroupingEndMM, Roadway);
                        DisplayForm.DisplaySublinkHarmonizedSpeed(RSLList, TroupingEndMM, Roadway);
                        DetermineLinkHarmonizedSpeed(RSLList, ref RLList, TroupingEndMM);
                        DisplayForm.DisplayTSSSPDHarmLinkStatus(RLList);

                        SpeedType = "Harmonized";
                        GenerateSPDHarmMessages_Kittelson(RSLList, TroupingEndMM, clsGlobalVars.BOQMMLocation, Roadway, SpeedType, Justification);
                        #endregion
                    }
                    else
                    {
                        #region "If no trouping is required"
                        LogTxtMsg(txtTSSDataLog, "\t\tNo Trouping is required for sublinks because no queued or congested sublinks were detected");

                        DisplayForm.EnableNoTroupingMessage(Roadway.RecurringCongestionMMLocation);
                        DisplayForm.EnableNoSpdHarmMessage(Roadway.RecurringCongestionMMLocation);
                        GenerateSPDHarmMessages_Kittelson(RSLList, TroupingEndMM, clsGlobalVars.BOQMMLocation, Roadway, SpeedType, Justification);


                        #endregion
                    }

                    if (chkExcelDataLogging.Checked == true)
                    {
                        #region "Log Trouping and Harmonized sublink speed into Excel worksheet"
                        //Excel
                        int t = 0;
                        CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow, 1] = currDateTime;
                        CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow, 2] = "CVsSpd" + ":" + clsGlobalVars.BOQMMLocation + ":" + TroupingEndMM;
                        CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 1, 1] = currDateTime;
                        CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 1, 2] = "TSSSpd" + ":" + clsGlobalVars.BOQMMLocation + ":" + TroupingEndMM; ;
                        CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 2, 1] = currDateTime;
                        CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 2, 2] = "RecSpd" + ":" + clsGlobalVars.BOQMMLocation + ":" + TroupingEndMM; ;
                        CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 3, 1] = currDateTime;
                        CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 3, 2] = "TroSpd" + ":" + clsGlobalVars.BOQMMLocation + ":" + TroupingEndMM; ;
                        CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 4, 1] = currDateTime;
                        CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 4, 2] = "HarSpd" + ":" + clsGlobalVars.BOQMMLocation + ":" + TroupingEndMM; ;
                        foreach (clsRoadwaySubLink rsl in RSLList)
                        {
                            //Excel
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow, t + 3] = rsl.CVAvgSpeed.ToString("0");
                            #region "CVSpeed Color"
                            if (rsl.CVAvgSpeed > 65.0)
                            {
                                CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow, t + 3].Interior.Color = System.Drawing.Color.LightSeaGreen;
                            }
                            else if (rsl.CVAvgSpeed > 60.0)
                            {
                                CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow, t + 3].Interior.Color = System.Drawing.Color.PaleTurquoise;
                            }
                            else if (rsl.CVAvgSpeed > 55.0)
                            {
                                CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow, t + 3].Interior.Color = System.Drawing.Color.LimeGreen;
                            }
                            else if (rsl.CVAvgSpeed > 50.0)
                            {
                                CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow, t + 3].Interior.Color = System.Drawing.Color.YellowGreen;
                            }
                            else if (rsl.CVAvgSpeed >= 45.0)
                            {
                                CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow, t + 3].Interior.Color = System.Drawing.Color.Yellow;
                            }
                            else if (rsl.CVAvgSpeed > 40.0)
                            {
                                CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow, t + 3].Interior.Color = System.Drawing.Color.Gold;
                            }
                            else if (rsl.CVAvgSpeed > 35.0)
                            {
                                CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow, t + 3].Interior.Color = System.Drawing.Color.DarkOrange;
                            }
                            else if (rsl.CVAvgSpeed > 30.0)
                            {
                                CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow, t + 3].Interior.Color = System.Drawing.Color.Tomato;
                            }
                            else if (rsl.CVAvgSpeed <= 30.0)
                            {
                                CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow, t + 3].Interior.Color = System.Drawing.Color.Red;
                            }
                            #endregion
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 1, t + 3] = rsl.TSSAvgSpeed.ToString("0");
                            #region "TSSSpeed Color"
                            if (rsl.TSSAvgSpeed > 65.0)
                            {
                                CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 1, t + 3].Interior.Color = System.Drawing.Color.LightSeaGreen;
                            }
                            else if (rsl.TSSAvgSpeed > 60.0)
                            {
                                CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 1, t + 3].Interior.Color = System.Drawing.Color.PaleTurquoise;
                            }
                            else if (rsl.TSSAvgSpeed > 55.0)
                            {
                                CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 1, t + 3].Interior.Color = System.Drawing.Color.LimeGreen;
                            }
                            else if (rsl.TSSAvgSpeed > 50.0)
                            {
                                CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 1, t + 3].Interior.Color = System.Drawing.Color.YellowGreen;
                            }
                            else if (rsl.TSSAvgSpeed > 45.0)
                            {
                                CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 1, t + 3].Interior.Color = System.Drawing.Color.Yellow;
                            }
                            else if (rsl.TSSAvgSpeed > 40.0)
                            {
                                CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 1, t + 3].Interior.Color = System.Drawing.Color.Gold;
                            }
                            else if (rsl.TSSAvgSpeed > 35.0)
                            {
                                CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 1, t + 3].Interior.Color = System.Drawing.Color.DarkOrange;
                            }
                            else if (rsl.TSSAvgSpeed > 30.0)
                            {
                                CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 1, t + 3].Interior.Color = System.Drawing.Color.Tomato;
                            }
                            else if (rsl.TSSAvgSpeed <= 30.0)
                            {
                                CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 1, t + 3].Interior.Color = System.Drawing.Color.Red;
                            }
                            #endregion
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 2, t + 3] = rsl.RecommendedSpeed.ToString("0");
                            #region "RecommendedSpeed Color"
                            if (rsl.RecommendedSpeed > 65.0)
                            {
                                CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 2, t + 3].Interior.Color = System.Drawing.Color.LightSeaGreen;
                            }
                            else if (rsl.RecommendedSpeed > 60.0)
                            {
                                CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 2, t + 3].Interior.Color = System.Drawing.Color.PaleTurquoise;
                            }
                            else if (rsl.RecommendedSpeed > 55.0)
                            {
                                CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 2, t + 3].Interior.Color = System.Drawing.Color.LimeGreen;
                            }
                            else if (rsl.RecommendedSpeed > 50.0)
                            {
                                CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 2, t + 3].Interior.Color = System.Drawing.Color.YellowGreen;
                            }
                            else if (rsl.RecommendedSpeed > 45.0)
                            {
                                CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 2, t + 3].Interior.Color = System.Drawing.Color.Yellow;
                            }
                            else if (rsl.RecommendedSpeed > 40.0)
                            {
                                CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 2, t + 3].Interior.Color = System.Drawing.Color.Gold;
                            }
                            else if (rsl.RecommendedSpeed > 35.0)
                            {
                                CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 2, t + 3].Interior.Color = System.Drawing.Color.DarkOrange;
                            }
                            else if (rsl.RecommendedSpeed > 30.0)
                            {
                                CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 2, t + 3].Interior.Color = System.Drawing.Color.Tomato;
                            }
                            else if (rsl.RecommendedSpeed <= 30.0)
                            {
                                CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 2, t + 3].Interior.Color = System.Drawing.Color.Red;
                            }
                            #endregion
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 3, t + 3] = rsl.TroupeSpeed.ToString("0");
                            #region "Troupe Color"
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 3, t + 3].Interior.Color = System.Drawing.Color.White;
                            if (rsl.TroupeInclusionOverride == true)
                            {
                                CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 3, t + 3].Interior.Color = System.Drawing.Color.Silver;
                            }
                            if (rsl.BeginTroupe == true)
                            {
                                CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 3, t + 3].Interior.Color = System.Drawing.Color.Pink;
                            }
                            #endregion
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 4, t + 3] = rsl.HarmonizedSpeed.ToString("0");
                            #region "Troupe Color"
                            CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 4, t + 3].Interior.Color = System.Drawing.Color.White;
                            if (rsl.SpdHarmInclusionOverride == true)
                            {
                                CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 4, t + 3].Interior.Color = System.Drawing.Color.Silver;
                            }
                            if (rsl.BeginSpdHarm == true)
                            {
                                CVSPDHarmWorkSheets[1].Cells[CVSPDHarmWSCurrRow + 4, t + 3].Interior.Color = System.Drawing.Color.Pink;
                            }
                            #endregion

                            t = t + 1;
                        }
                        //Excel
                        CVSPDHarmWSCurrRow = CVSPDHarmWSCurrRow + 6;
                        #endregion
                    }
                    TimeSpan EndTrouping = new TimeSpan(DateTime.Now.Ticks);
                    LogTxtMsg(txtTSSDataLog, "\t\tTime for sublink speed harmonization: \t" + (EndTrouping.TotalMilliseconds - StartTrouping.TotalMilliseconds).ToString("0") + " msecs");
                }
                
                System.Windows.Forms.Application.DoEvents();
                LogTxtMsg(txtTSSDataLog, "\r\n" +  DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + ":" + DateTime.Now.Millisecond +
                          "\tFinished processing TSS data for the last " + clsGlobalVars.TSSDataLoadingFrequency + " seconds interval");
                TSSDataProcessor.WriteLine("\r\n\r\n" + DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + "::" + DateTime.Now.Millisecond +
                                        "\tFinished Processing TSS traffic data for the last: " + clsGlobalVars.TSSDataLoadingFrequency + " seconds interval.");
                NumberNoTSSDataIntervals = 0;
                DisplayForm.DisableNoTSSDataMessage();
            }
            else
            {
                NumberNoTSSDataIntervals = NumberNoTSSDataIntervals + 1;
                LogTxtMsg(txtTSSDataLog, "\r\n\t\tNo TSS traffic data was retrieved for the last " + clsGlobalVars.TSSDataPollingFrequency + " second interval." +
                                        "\r\n\t\tNumber of " + clsGlobalVars.TSSDataLoadingFrequency + " second intervals with no TSS data retrieved: " + NumberNoTSSDataIntervals);
                LogTxtMsg(txtINFLOLog, "\r\n\t\tNo TSS traffic data was retrieved for the last " + clsGlobalVars.TSSDataPollingFrequency + " second interval." +
                                        "\r\n\t\tNumber of " + clsGlobalVars.TSSDataLoadingFrequency + " second intervals with no TSS data retrieved: " + NumberNoTSSDataIntervals);
                DisplayForm.txtTSSDate.Text = DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second;
                if (NumberNoTSSDataIntervals >= (60/clsGlobalVars.TSSDataLoadingFrequency))
                {
                    #region "If NO TSS data was retrieved for the last five seconds"
                    if (clsGlobalVars.InfrastructureBOQMMLocation != -1)
                    {
                        clsGlobalVars.InfrastructureBOQMMLocation = -1;
                        clsGlobalVars.InfrastructureBOQTime = DateTime.Now;
                    }
                    DisplayForm.EnableNoTSSDataMessage(NumberNoTSSDataIntervals);
                    DisplayForm.txtTSSBOQ.Text = "No Queue";
                    DisplayForm.ClearTSSQueuedLinkStatus();
                    //if (clsGlobalVars.CVBOQMMLocation == -1)
                    //{
                    //    if (clsGlobalVars.BOQMMLocation != -1)
                    //    {
                    //        clsGlobalVars.BOQMMLocation = -1;
                    //        clsGlobalVars.BOQTime = DateTime.Now;
                    //    }
                    //    DisplayForm.txtBOQ.Text = "No Queue";
                    //}
                    //if (NumberNoCVDataIntervals > 0)
                    //{
                        foreach (clsRoadwayLink RL in RLList)
                        {
                            RL.TSSAvgSpeed = clsGlobalVars.MaximumDisplaySpeed;
                        }
                        DisplayForm.ClearCVSubLinkSPDHarmStatus();
                        DisplayForm.ClearCVSubLinkTroupeStatus();
                        DisplayForm.EnableNoSpdHarmMessage(Roadway.RecurringCongestionMMLocation);
                        DisplayForm.EnableNoTroupingMessage(Roadway.RecurringCongestionMMLocation);
                    //}
                    #endregion
                }
            }
            System.Windows.Forms.Application.DoEvents();
            System.Windows.Forms.Application.DoEvents();
        }

        private string GetLastIntervalCVData(clsDatabase DB, DateTime IntervalUTCTime, ref List<clsCVData> CVList, clsEnums.enDirection Direction, double LHeading, double UHeading, ref int NumberRecordsRetrieved)
        {
            string retValue = string.Empty;
            string sqlQuery = string.Empty;

            CVList.Clear();
            
            DataSet CVDataSet = new DataSet("CVDataSet");
            //IntervalUTCTime = IntervalUTCTime.AddSeconds(-clsGlobalVars.CVDataPollingFrequency);
            if (clsGlobalVars.DBInterfaceType.ToLower() == "oledb")
            {
                //sqlQuery = "Select * from TME_CVData_Input where DateGenerated>=#" + IntervalUTCTime.AddSeconds(-clsGlobalVars.CVDataPollingFrequency) + "# and DateGenerated<=#" + IntervalUTCTime + "#" + " and MMLocation>=" + Roadway.BeginMM + " and mmlocation <=" + Roadway.RecurringCongestionMMLocation;
                sqlQuery = "Select * from TME_CVData_Input where MMLocation>=" + Roadway.BeginMM + " and mmlocation <=" + Roadway.RecurringCongestionMMLocation;
            }
            else if (clsGlobalVars.DBInterfaceType.ToLower() == "sqlserver")
            {

                //sqlQuery = "Select * from TME_CVData_Input where DateGenerated>='" + IntervalUTCTime.AddSeconds(-clsGlobalVars.CVDataPollingFrequency) + "' and DateGenerated<='" + IntervalUTCTime + "'" + " and MMLocation>=" + Roadway.BeginMM + " and mmlocation <=" + Roadway.RecurringCongestionMMLocation;
                sqlQuery = "Select * from TME_CVData_Input where MMLocation>=" + Roadway.BeginMM + " and mmlocation <=" + Roadway.RecurringCongestionMMLocation;
            }
            
            retValue = string.Empty;
            try
            {
                retValue = DB.FillDataSet(sqlQuery, ref CVDataSet);
                if (CVDataSet.Tables[0] != null)
                {
                    FillDataSetLog.WriteLine(DateTime.Now + ",GetLastIntervalCVData," + CVDataSet.Tables[0].Rows.Count + "," + IntervalUTCTime.AddSeconds(-clsGlobalVars.CVDataPollingFrequency) + "," + IntervalUTCTime);
                }
                if (retValue.Length > 0)
                {
                    return retValue;
                }

                //display Detector station information 
                //LogTxtMsg(txtCVDataLog, "\tAvailable CV data records: ");
                if (CVDataSet.Tables[0] != null)
                {
                    if (CVDataSet.Tables[0].Rows.Count > 0)
                    {
                        NumberRecordsRetrieved = CVDataSet.Tables[0].Rows.Count;
                        foreach (DataRow row in CVDataSet.Tables[0].Rows)
                        {
                            clsCVData tmpCV = new clsCVData();
                            string NomadicDeviceId = string.Empty;
                            string RoadwayId = string.Empty;
                            DateTime DateGenerated = DateTime.UtcNow;
                            int Speed = 0;
                            int Heading = 0;
                            double Latitude = 0;
                            double Longitude = 0;
                            double MMLocation = 0;
                            double CoefficientOfFriction = 0;
                            bool CVQueuedState = false;
                            int Temperature = 0;

                            foreach (DataColumn col in CVDataSet.Tables[0].Columns)
                            {
                                switch (col.ColumnName.ToString().ToLower())
                                {
                                    case "nomadicdeviceid":
                                        NomadicDeviceId = row[col].ToString();
                                        tmpCV.NomadicDeviceID = NomadicDeviceId;
                                        break;
                                    case "roadwayid":
                                        RoadwayId = row[col].ToString();
                                        tmpCV.RoadwayID = RoadwayId;
                                        break;
                                    case "speed":
                                        if (row[col].ToString().Length > 0)
                                        {
                                            Speed = (int)(double.Parse(row[col].ToString()));
                                            tmpCV.Speed = Speed;
                                        }
                                        break;
                                    case "heading":
                                        if (row[col].ToString().Length > 0)
                                        {
                                            Heading = int.Parse(row[col].ToString());
                                            tmpCV.Heading = Heading;
                                        }
                                        break;
                                    case "latitude":
                                        if (row[col].ToString().Length > 0)
                                        {
                                            Latitude = double.Parse(row[col].ToString());
                                            tmpCV.Latitude = Latitude;
                                        }
                                        break;
                                    case "longitude":
                                        if (row[col].ToString().Length > 0)
                                        {
                                            Longitude = double.Parse(row[col].ToString());
                                            tmpCV.Longitude = Longitude;
                                        }
                                        break;
                                    case "mmlocation":
                                        if (row[col].ToString().Length > 0)
                                        {
                                            MMLocation = double.Parse(row[col].ToString());
                                            tmpCV.MMLocation = MMLocation;
                                        }
                                        break;
                                    case "cvqueuedstate":
                                        CVQueuedState = bool.Parse(row[col].ToString());
                                        tmpCV.Queued = CVQueuedState;
                                        break;
                                    case "coefficientOffriction":
                                        if (row[col].ToString().Length > 0)
                                        {
                                            CoefficientOfFriction = double.Parse(row[col].ToString());
                                            tmpCV.CoefficientFriction = CoefficientOfFriction;
                                        }
                                        break;
                                    case "temperature":
                                        if (row[col].ToString().Length > 0)
                                        {
                                            Temperature = int.Parse(row[col].ToString());
                                            tmpCV.Temperature = Temperature;
                                        }
                                        break;
                                }
                            }
                            if (tmpCV.Heading > 0)
                            {
                                if ((tmpCV.Heading >= 337.5) || (tmpCV.Heading < 22.5))
                                {
                                    tmpCV.Direction = clsEnums.GetDirIndexFromString("North");
                                }
                                else if ((tmpCV.Heading >= 22.5) || (tmpCV.Heading < 67.5))
                                {
                                    tmpCV.Direction = clsEnums.GetDirIndexFromString("NE");
                                }
                                else if ((tmpCV.Heading >= 67.5) || (tmpCV.Heading < 112.5))
                                {
                                    tmpCV.Direction = clsEnums.GetDirIndexFromString("East");
                                }
                                else if ((tmpCV.Heading >= 112.5) || (tmpCV.Heading < 157.5))
                                {
                                    tmpCV.Direction = clsEnums.GetDirIndexFromString("SE");
                                }
                                else if ((tmpCV.Heading >= 157.5) || (tmpCV.Heading < 202.5))
                                {
                                    tmpCV.Direction = clsEnums.GetDirIndexFromString("South");
                                }
                                else if ((tmpCV.Heading >= 202.5) || (tmpCV.Heading < 247.5))
                                {
                                    tmpCV.Direction = clsEnums.GetDirIndexFromString("SW");
                                }
                                else if ((tmpCV.Heading >= 247.5) || (tmpCV.Heading < 292.5))
                                {
                                    tmpCV.Direction = clsEnums.GetDirIndexFromString("West");
                                }
                                else if ((tmpCV.Heading >= 292.5) || (tmpCV.Heading < 337.5))
                                {
                                    tmpCV.Direction = clsEnums.GetDirIndexFromString("NW");
                                }
                            }
                            else
                            {
                                tmpCV.Direction = Direction;
                            }
                            //LogTxtMsg(txtCVDataLog, "\t\t" + NomadicDeviceId + ", " + RoadwayId + ", " + Speed + ", " + Latitude + ", " + Longitude + ", " + Heading + ", " + DateGenerated + ", " +
                            //                  MMLocation + ", " + CVQueuedState + ", " + CoefficientOfFriction + ", " + Temperature + ", " + Direction);

                            CVList.Add(tmpCV);
                        }
                    }
                    else
                    {
                        NumberRecordsRetrieved = 0;
                    }
                }
                else
                {
                    retValue = "Error in retrieving CV data for the last interval from: " + IntervalUTCTime.AddSeconds(-clsGlobalVars.CVDataPollingFrequency) + " to: " + IntervalUTCTime + ". No data was retrieved.";
                    return retValue;
                }
            }
            catch (Exception exc)
            {
                retValue = "Error in retrieving CV data for the last interval from: " +  IntervalUTCTime.AddSeconds(-clsGlobalVars.CVDataPollingFrequency) + " to: " + IntervalUTCTime + "\r\n\t" + exc.Message;
                return retValue;
            }

            return retValue;
        }
        private string ProcessRoadwaySublinkQueuedStatus(ref List<clsCVData> CVList, ref List<clsRoadwaySubLink> RSLList)
        {
            string retValue = string.Empty;
            string CVID = string.Empty;
            string CurrRecord = string.Empty;
            double TotalSpeed = 0;
            string tmpSublinkData = string.Empty;

            try
            {
                CVList.Sort((l, r) => l.MMLocation.CompareTo(r.MMLocation));
                RSLList.Sort((l, r) => l.BeginMM.CompareTo(r.BeginMM));

                foreach (clsRoadwaySubLink RSL in RSLList)
                {
                    RSL.CVAvgSpeed = 0;
                    RSL.TotalNumberCVs = 0;
                    RSL.Congested = false;
                    RSL.Queued = false;
                    RSL.TotalNumberCVs = 0;
                    RSL.NumberQueuedCVs = 0;
                    RSL.PercentQueuedCVs = 0;
                    RSL.SmoothedSpeed[clsGlobalVars.CVDataSmoothedSpeedIndex] = 0;
                    TotalSpeed = 0;

                    RSL.CVList = new List<clsCVData>();

                    foreach (clsCVData CV in CVList)
                    {
                        //LogTxtMsg(txtCVDataLog, "\tProcessing roadway sublink: " + RSL.Identifier + "\tFromMM: " + RSL.BeginMM + "\tToMM: " + RSL.EndMM + "\tDirection: " + RSL.Direction + "\tCV: " + CV.NomadicDeviceID + " - " + CV.Direction + " - " + CV.MMLocation);
                        CurrRecord = "RSL: " + RSL.Identifier + "   CV: " + CV.NomadicDeviceID;
                        if ((CV.MMLocation >= RSL.BeginMM) && (CV.MMLocation < RSL.EndMM) && (RSL.Direction == CV.Direction))
                        {
                            CVID = CV.NomadicDeviceID;
                            CV.SublinkID = RSL.Identifier.ToString();
                            RSL.CVList.Add(CV);
                            RSL.TotalNumberCVs = RSL.TotalNumberCVs + 1;
                            if (CV.Queued == true)
                            {
                                RSL.NumberQueuedCVs = RSL.NumberQueuedCVs + 1;
                            }
                            TotalSpeed = TotalSpeed + CV.Speed;
                        }
                        //else if (CV.MMLocation >= RSL.EndMM)
                        //{
                        //    break;
                        //}
                    }
                    
                    if (RSL.TotalNumberCVs > 0)
                    {
                        RSL.SmoothedSpeed[clsGlobalVars.CVDataSmoothedSpeedIndex] = TotalSpeed / RSL.TotalNumberCVs;
                        RSL.PercentQueuedCVs = (RSL.NumberQueuedCVs * 100) / RSL.TotalNumberCVs;
                        RSL.VolumeDiff = RSL.TotalNumberCVs - RSL.PrevTotalNumberCVs;

                        RSL.FlowRate = (RSL.VolumeDiff / CVTimeDiff) * 3600;
                        RSL.Density = (RSL.TotalNumberCVs / clsGlobalVars.SubLinkLength);
                        RSL.DensityDiff = RSL.Density - RSL.PrevDensity;
                        if (RSL.DensityDiff > 0)
                        {
                            RSL.ShockWaveRate = (RSL.FlowRate / RSL.DensityDiff);
                        }
                        tmpSublinkData = tmpSublinkData + RSL.Identifier + "," + RSL.PrevTotalNumberCVs + ";" + RSL.TotalNumberCVs + ";" + RSL.VolumeDiff + ";" + RSL.PercentQueuedCVs + ";" +
                                         CVTimeDiff + ";" + RSL.FlowRate.ToString("0") + ";" + RSL.Density.ToString("0") + ";" + RSL.PrevDensity.ToString("0") + ";" + 
                                         RSL.DensityDiff + ";" + RSL.ShockWaveRate.ToString("0") + ",";
                        RSL.PrevDensity = RSL.Density;
                        RSL.PrevTotalNumberCVs = RSL.TotalNumberCVs;
                    }

                    TotalSpeed = 0;
                    int NoIntervals = 0;
                    for (int i = 0; i < clsGlobalVars.CVDataSmoothedSpeedArraySize; i++)
                    {
                        TotalSpeed = TotalSpeed + RSL.SmoothedSpeed[i];
                        if (RSL.SmoothedSpeed[i] > 0)
                        {
                            NoIntervals++;
                        }
                    }

                    if (NoIntervals > 0)
                    {
                        RSL.CVAvgSpeed = TotalSpeed / NoIntervals;
                    }

                    if (RSL.CVAvgSpeed == 0)
                    {
                        RSL.CVAvgSpeed = clsGlobalVars.MaximumDisplaySpeed;
                    }
                    else if (RSL.CVAvgSpeed > clsGlobalVars.MaximumDisplaySpeed)
                    {
                        RSL.CVAvgSpeed = clsGlobalVars.MaximumDisplaySpeed;
                    }
                    if (RSL.CVAvgSpeed >= clsGlobalVars.LinkCongestedSpeedThreshold)
                    {
                        RSL.Congested = false;
                    }
                    else if (RSL.CVAvgSpeed > clsGlobalVars.LinkQueuedSpeedThreshold)
                    {
                        RSL.Congested = true;
                    }
                    RSL.Queued = false;
                    if (RSL.PercentQueuedCVs >= clsGlobalVars.SubLinkPercentQueuedCV)
                    {
                        RSL.Queued = true;
                    }
                    RSL.CVDateProcessed = DateTime.UtcNow;
                    //LogTxtMsg(txtCVDataLog, "\t\t\tRSL CVAvg speed: " + RSL.CVAvgSpeed.ToString("0") + "\tNumCVs: " + RSL.TotalNumberCVs + "\tNumQueuedCVs: " + RSL.NumberQueuedCVs + "\t\t%QueuedCVs: " + RSL.PercentQueuedCVs.ToString("00") +
                    //                             "\tQueued " + RSL.Queued + "\tDateProcessed: " + RSL.CVDateProcessed + "\tBeginMM: " + RSL.BeginMM);
                }
                SubLinKDataLog.WriteLine(tmpSublinkData);
            }
            catch (Exception ex)
            {
                retValue = "Error in processing roadway sublink status from CV data. Error Hint - Current record: " + CurrRecord + "\r\n" + ex.Message;
                return retValue;
            }
            return retValue;
        }
        private string CalculateSublinkTroupeSpeed(ref List<clsRoadwaySubLink> RSLList, double TroupingEndMM, double TroupingEndSpeed)
        {
            string retValue = string.Empty;
            string CVID = string.Empty;
            string CurrRecord = string.Empty;
            int TroupeStartSublink = -1;
            double LastTroupeAvgSpeed = 0;
            clsTroupe tmpTroupe = new clsTroupe();

            try
            {
                //Sort sublinks based on sublink identifier assuming that sublinks are numberered starting from the sublink which is furthest upstream from the reuccuring congestion location
                //The furthest upstream sublink from the recurring congestion location is numbered as 1.
                RSLList.Sort((l, r) => l.Identifier.CompareTo(r.Identifier));


                //Find the furthest upstream link from the recurring congestion location with a recommended speed > 0 to start the first troupe
                for (int i = 0; i < RSLList.Count; i++)
                {
                    if (Roadway.Direction == Roadway.MMIncreasingDirection)
                    {
                        if ((RSLList[i].RecommendedSpeed > 0) && (RSLList[i].BeginMM < TroupingEndMM))
                        {
                            TroupeStartSublink = i;
                            tmpTroupe.MaxSpeed = RSLList[i].RecommendedSpeed;
                            tmpTroupe.MinSpeed = RSLList[i].RecommendedSpeed;
                            tmpTroupe.AvgSpeed = 0;
                            tmpTroupe.NumberSubLinks = 1;
                            tmpTroupe.SubLinks.Add(RSLList[i]);
                            RSLList[i].BeginTroupe = true;
                            break;
                        }
                    }
                    else
                    {
                        if ((RSLList[i].RecommendedSpeed > 0) && (RSLList[i].BeginMM > TroupingEndMM))
                        {
                            TroupeStartSublink = i;
                            tmpTroupe.MaxSpeed = RSLList[i].RecommendedSpeed;
                            tmpTroupe.MinSpeed = RSLList[i].RecommendedSpeed;
                            tmpTroupe.AvgSpeed = 0;
                            tmpTroupe.NumberSubLinks = 1;
                            tmpTroupe.SubLinks.Add(RSLList[i]);
                            RSLList[i].BeginTroupe = true;
                            break;
                        }
                    }
                }

                //Process the rest of the sublinks
                if (TroupeStartSublink != -1)
                {
                    for (int i = TroupeStartSublink + 1; i < RSLList.Count; i++)
                    {
                        if (Roadway.Direction == Roadway.MMIncreasingDirection)
                        {
                            if (RSLList[i].BeginMM < TroupingEndMM)
                            {
                                if ((RSLList[i].RecommendedSpeed >= (tmpTroupe.MaxSpeed - clsGlobalVars.TroupeRange)) &&
                                   (RSLList[i].RecommendedSpeed <= (tmpTroupe.MinSpeed + clsGlobalVars.TroupeRange)))
                                {
                                    tmpTroupe.SubLinks.Add(RSLList[i]);
                                    RSLList[i].TroupeInclusionOverride = false;
                                    if (RSLList[i].RecommendedSpeed > tmpTroupe.MaxSpeed)
                                    {
                                        tmpTroupe.MaxSpeed = RSLList[i].RecommendedSpeed;
                                    }
                                    if (RSLList[i].RecommendedSpeed < tmpTroupe.MinSpeed)
                                    {
                                        tmpTroupe.MinSpeed = RSLList[i].RecommendedSpeed;
                                    }
                                }
                                else
                                {
                                    tmpTroupe.CalculateTroupeAvgSpeed();
                                    tmpTroupe.CalculateTroupeLength();
                                    tmpTroupe.CalculateTroupeTravelTime();

                                    if (tmpTroupe.TravelTime < clsGlobalVars.DSD)
                                    {
                                        tmpTroupe.SubLinks.Add(RSLList[i]);
                                        RSLList[i].TroupeInclusionOverride = true;
                                        if (RSLList[i].RecommendedSpeed > tmpTroupe.MaxSpeed)
                                        {
                                            tmpTroupe.MaxSpeed = RSLList[i].RecommendedSpeed;
                                        }
                                        if (RSLList[i].RecommendedSpeed < tmpTroupe.MinSpeed)
                                        {
                                            tmpTroupe.MinSpeed = RSLList[i].RecommendedSpeed;
                                        }
                                    }
                                    else
                                    {
                                        //Assign the troupe avgspeed to all sublinks in the troupe
                                        LastTroupeAvgSpeed = tmpTroupe.AvgSpeed;
                                        foreach (clsRoadwaySubLink TSL in tmpTroupe.SubLinks)
                                        {
                                            foreach (clsRoadwaySubLink RSL in RSLList)
                                            {
                                                if (TSL.Identifier == RSL.Identifier)
                                                {
                                                    RSL.TroupeSpeed = tmpTroupe.AvgSpeed;
                                                    break;
                                                }
                                            }
                                        }

                                        tmpTroupe.SubLinks.Clear();
                                        tmpTroupe.MaxSpeed = RSLList[i].RecommendedSpeed;
                                        tmpTroupe.MinSpeed = RSLList[i].RecommendedSpeed;
                                        tmpTroupe.AvgSpeed = 0;
                                        tmpTroupe.NumberSubLinks = 1;
                                        tmpTroupe.SubLinks.Add(RSLList[i]);
                                        RSLList[i].BeginTroupe = true;
                                    }
                                }
                            }
                            else
                            {
                                break;
                            }
                        }
                        else if (Roadway.Direction != Roadway.MMIncreasingDirection)
                        {
                            if (RSLList[i].BeginMM > TroupingEndMM)
                            {
                                if ((RSLList[i].RecommendedSpeed >= (tmpTroupe.MaxSpeed - clsGlobalVars.TroupeRange)) &&
                                (RSLList[i].RecommendedSpeed <= (tmpTroupe.MinSpeed + clsGlobalVars.TroupeRange)))
                                {
                                    tmpTroupe.SubLinks.Add(RSLList[i]);
                                    RSLList[i].TroupeInclusionOverride = false;
                                    if (RSLList[i].RecommendedSpeed > tmpTroupe.MaxSpeed)
                                    {
                                        tmpTroupe.MaxSpeed = RSLList[i].RecommendedSpeed;
                                    }
                                    if (RSLList[i].RecommendedSpeed < tmpTroupe.MinSpeed)
                                    {
                                        tmpTroupe.MinSpeed = RSLList[i].RecommendedSpeed;
                                    }
                                }
                                else
                                {
                                    tmpTroupe.CalculateTroupeAvgSpeed();
                                    tmpTroupe.CalculateTroupeLength();
                                    tmpTroupe.CalculateTroupeTravelTime();

                                    if (tmpTroupe.TravelTime < clsGlobalVars.DSD)
                                    {
                                        tmpTroupe.SubLinks.Add(RSLList[i]);
                                        RSLList[i].TroupeInclusionOverride = true;
                                        if (RSLList[i].RecommendedSpeed > tmpTroupe.MaxSpeed)
                                        {
                                            tmpTroupe.MaxSpeed = RSLList[i].RecommendedSpeed;
                                        }
                                        if (RSLList[i].RecommendedSpeed < tmpTroupe.MinSpeed)
                                        {
                                            tmpTroupe.MinSpeed = RSLList[i].RecommendedSpeed;
                                        }
                                    }
                                    else
                                    {
                                        //Assign the troupe avgspeed to all sublinks in the troupe
                                        LastTroupeAvgSpeed = tmpTroupe.AvgSpeed;
                                        foreach (clsRoadwaySubLink TSL in tmpTroupe.SubLinks)
                                        {
                                            foreach (clsRoadwaySubLink RSL in RSLList)
                                            {
                                                if (TSL.Identifier == RSL.Identifier)
                                                {
                                                    RSL.TroupeSpeed = tmpTroupe.AvgSpeed;
                                                    break;
                                                }
                                            }
                                        }

                                        tmpTroupe.SubLinks.Clear();
                                        tmpTroupe.MaxSpeed = RSLList[i].RecommendedSpeed;
                                        tmpTroupe.MinSpeed = RSLList[i].RecommendedSpeed;
                                        tmpTroupe.AvgSpeed = 0;
                                        tmpTroupe.NumberSubLinks = 1;
                                        tmpTroupe.SubLinks.Add(RSLList[i]);
                                        RSLList[i].BeginTroupe = true;
                                    }
                                }
                            }
                            else
                            {
                                break;
                            }
                        }
                        //LogTxtMsg(txtCVDataLog, "\t\t\tRSL CVAvg speed: " + RSL.CVAvgSpeed.ToString("0") + "\tNumCVs: " + RSL.TotalNumberCVs + "\tNumQueuedCVs: " + RSL.NumberQueuedCVs + "\t\t%QueuedCVs: " + RSL.PercentQueuedCVs.ToString("00") +
                        //                             "\tQueued " + RSL.Queued + "\tDateProcessed: " + RSL.CVDateProcessed + "\tBeginMM: " + RSL.BeginMM);
                    }

                    if (tmpTroupe.SubLinks.Count > 0)
                    {
                        tmpTroupe.CalculateTroupeAvgSpeed();
                        tmpTroupe.CalculateTroupeLength();
                        tmpTroupe.CalculateTroupeTravelTime();

                        //travel time of the troupe based on the average speed of troupe and length of troupe
                        if (tmpTroupe.TravelTime < clsGlobalVars.DSD)
                        {
                            //Join the sublinks to the previous troupe and assign the last troupe avgspeed to the sublinks in the current troupe since they cannot be a troupe on their own
                            foreach (clsRoadwaySubLink TSL in tmpTroupe.SubLinks)
                            {
                                foreach (clsRoadwaySubLink RSL in RSLList)
                                {
                                    if (TSL.Identifier == RSL.Identifier)
                                    {
                                        RSL.TroupeSpeed = LastTroupeAvgSpeed;
                                        RSL.TroupeInclusionOverride = true;
                                        RSL.BeginTroupe = false;
                                        break;
                                    }
                                }
                            }
                        }
                        else
                        {
                            //Assign the troupe avgspeed to all sublinks in the troupe
                            foreach (clsRoadwaySubLink TSL in tmpTroupe.SubLinks)
                            {
                                foreach (clsRoadwaySubLink RSL in RSLList)
                                {
                                    if (TSL.Identifier == RSL.Identifier)
                                    {
                                        RSL.TroupeSpeed = tmpTroupe.AvgSpeed;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    LogTxtMsg(txtSpdHarmLog, "\r\n" + DateTime.Now + "\tNo sublink was found with speed > 0 upstream of the congestion or queue location");
                }
            }
            catch (Exception ex)
            {
                retValue = "Error in processing roadway sublinks to determine troupes of sublinks." + "\r\n" + ex.Message;
                return retValue;
            }
            return retValue;
        }
 
        private string CalculateSublinkHarmonizedSpeed(ref List<clsRoadwaySubLink> RSLList, double TroupingEndMM, double TroupingEndSpeed)
        {
            string retValue = string.Empty;
            string CVID = string.Empty;
            string CurrRecord = string.Empty;
            int SPDHarmStartSublink = 0;
            double SPDHarmGroupHarmonizedSpeed = 0;
            int SPDHarmGroupNumberOfSubLinks = 0;
            int tmpNumberofSublinks = 0;
            double LastSPDHarmGroupHarmonizedSpeed = 0;
            double SpeedDiff = 0;

            try
            {
                //Sort sublinks based on sublink identifier assuming that sublinks are numberered starting from the sublink which is furthest upstream from the reuccuring congestion location
                //The furthest upstream sublink from the recurring congestion location is numbered as 1.
                RSLList.Sort((l, r) => l.Identifier.CompareTo(r.Identifier));

                List<clsRoadwaySubLink> tmpSPDHarmSublinkGroup = new List<clsRoadwaySubLink>();

                //Find the furthest link upstream from the recurring congestion location where Trouping ended to start the speed harmonization process
                for (int i = RSLList.Count - 1; i >= 0; i--)
                {
                    if (Roadway.Direction == Roadway.MMIncreasingDirection)
                    {
                        if ((RSLList[i].RecommendedSpeed > 0) && (RSLList[i].BeginMM < TroupingEndMM))
                        {
                            SPDHarmStartSublink = i;
                            break;
                        }
                    }
                    else if (Roadway.Direction != Roadway.MMIncreasingDirection)
                    {
                        if ((RSLList[i].RecommendedSpeed > 0) && (RSLList[i].BeginMM > TroupingEndMM))
                        {
                            SPDHarmStartSublink = i;
                            break;
                        }
                    }
                }
                //Reset all the sublink troupe speed that are less than the minimum display speed to the minimum display speed
                for (int j = SPDHarmStartSublink; j >= 0; j--)
                {
                        if (RSLList[j].TroupeSpeed < clsGlobalVars.LinkQueuedSpeedThreshold)
                        {
                            RSLList[j].TroupeSpeed = clsGlobalVars.LinkQueuedSpeedThreshold;
                        }
                }

                SPDHarmGroupHarmonizedSpeed = TroupingEndSpeed;
                if (TroupingEndSpeed < clsGlobalVars.LinkQueuedSpeedThreshold)
                {
                    SPDHarmGroupHarmonizedSpeed = clsGlobalVars.LinkQueuedSpeedThreshold;
                }

                if ((SPDHarmGroupHarmonizedSpeed % clsGlobalVars.TroupeRange) != 0)
                {
                    SPDHarmGroupHarmonizedSpeed = (int)((SPDHarmGroupHarmonizedSpeed + 5) / 5) * 5;
                }

                SpeedDiff = RSLList[SPDHarmStartSublink].TroupeSpeed - SPDHarmGroupHarmonizedSpeed;
                if (SpeedDiff > 5) // greater than five miles per hour 
                {
                    SPDHarmGroupHarmonizedSpeed = SPDHarmGroupHarmonizedSpeed + 5;
                }
                else
                {
                    SPDHarmGroupHarmonizedSpeed = RSLList[SPDHarmStartSublink].TroupeSpeed;
                }

                RSLList[SPDHarmStartSublink].HarmonizedSpeed = SPDHarmGroupHarmonizedSpeed;
                LastSPDHarmGroupHarmonizedSpeed = SPDHarmGroupHarmonizedSpeed;
                RSLList[SPDHarmStartSublink].BeginSpdHarm = true;
                SPDHarmGroupNumberOfSubLinks = (int)(Math.Ceiling((SPDHarmGroupHarmonizedSpeed * clsGlobalVars.DSD) / (clsGlobalVars.SubLinkLength * 3600)));
                tmpNumberofSublinks = 1;
                tmpSPDHarmSublinkGroup.Add(RSLList[SPDHarmStartSublink]);

                /*for (int j = SPDHarmStartSublink-1; j >= 0; j--)
                {
                    if (Roadway.Direction == Roadway.MMIncreasingDirection)
                    {
                        if ((RSLList[j].RecommendedSpeed > 0) && (RSLList[j].BeginMM < TroupingEndMM))
                        {
                            SPDHarmStartSublink = i;

                            SpeedDiff = RSLList[i].TroupeSpeed - SPDHarmGroupHarmonizedSpeed;

                            LastSPDHarmGroupHarmonizedSpeed = SPDHarmGroupHarmonizedSpeed;
                            if ((SpeedDiff > 5) || (SpeedDiff < 0))// greater than five miles per hour or < 0 in case the speed was less than the previous 
                            {
                                SPDHarmGroupHarmonizedSpeed = SPDHarmGroupHarmonizedSpeed + 5;
                            }
                            else
                            {
                                SPDHarmGroupHarmonizedSpeed = RSLList[i].TroupeSpeed;
                            }

                            
                            SPDHarmGroupHarmonizedSpeed = RSLList[i].TroupeSpeed;
                            RSLList[i].HarmonizedSpeed = SPDHarmGroupHarmonizedSpeed;
                            LastSPDHarmGroupHarmonizedSpeed = SPDHarmGroupHarmonizedSpeed;
                            RSLList[i].BeginSpdHarm = true;
                            SPDHarmGroupNumberOfSubLinks = (int)(Math.Ceiling((SPDHarmGroupHarmonizedSpeed * clsGlobalVars.DSD)/(clsGlobalVars.SubLinkLength * 3600)));
                            tmpNumberofSublinks = 1;
                            tmpSPDHarmSublinkGroup.Add(RSLList[i]);
                            break;
                        }
                    }
                    else
                    {
                        if ((RSLList[i].RecommendedSpeed > 0) && (RSLList[i].BeginMM > TroupingEndMM))
                        {
                            SPDHarmStartSublink = i;
                            SPDHarmGroupHarmonizedSpeed = RSLList[i].TroupeSpeed;
                            RSLList[i].HarmonizedSpeed = SPDHarmGroupHarmonizedSpeed;
                            LastSPDHarmGroupHarmonizedSpeed = SPDHarmGroupHarmonizedSpeed;
                            RSLList[i].BeginSpdHarm = true;
                            SPDHarmGroupNumberOfSubLinks = (int)(Math.Ceiling((SPDHarmGroupHarmonizedSpeed * clsGlobalVars.DSD) / (clsGlobalVars.SubLinkLength * 3600)));
                            tmpNumberofSublinks = 1;
                            tmpSPDHarmSublinkGroup.Add(RSLList[i]);
                            break;
                        }
                    }
                }*/

                //Start processing the rest of the sublinks
                if (SPDHarmStartSublink != 0)
                {
                    for (int i = SPDHarmStartSublink - 1; i >= 0; i--)
                    {
                        if (Roadway.Direction == Roadway.MMIncreasingDirection)
                        {
                            if (RSLList[i].BeginMM < TroupingEndMM)
                            {
                                if (RSLList[i].TroupeSpeed == SPDHarmGroupHarmonizedSpeed)
                                {
                                    //Current sublink troupe speed = current sublink harmonization group  speed
                                    //Add the sublink to the sublink harmonization group, Change the sublink harmonizedspeed and continue
                                    RSLList[i].HarmonizedSpeed = SPDHarmGroupHarmonizedSpeed;
                                    tmpSPDHarmSublinkGroup.Add(RSLList[i]);
                                    tmpNumberofSublinks++;
                                }
                                //If the current sublink troupe speed > current sublink harmonization group speed
                                else if (RSLList[i].TroupeSpeed > SPDHarmGroupHarmonizedSpeed) 
                                {   //if the number of sublink in the harmonization group > number sublinks required for DSD
                                    if (tmpNumberofSublinks >= SPDHarmGroupNumberOfSubLinks)
                                    {
                                        //End the current sublink harmonization group and start a new sublink harmonization group and add the current sublink to it.
                                        SpeedDiff = RSLList[i].TroupeSpeed - SPDHarmGroupHarmonizedSpeed;

                                        LastSPDHarmGroupHarmonizedSpeed = SPDHarmGroupHarmonizedSpeed;
                                        if (SpeedDiff > 5) // greater than fivemiles per hour
                                        {
                                            SPDHarmGroupHarmonizedSpeed = SPDHarmGroupHarmonizedSpeed + 5;
                                        }
                                        else
                                        {
                                            SPDHarmGroupHarmonizedSpeed = RSLList[i].TroupeSpeed;
                                        }
                                        RSLList[i].HarmonizedSpeed = SPDHarmGroupHarmonizedSpeed;
                                        RSLList[i].BeginSpdHarm = true;
                                        tmpNumberofSublinks = 1;
                                        tmpSPDHarmSublinkGroup.Clear();
                                        tmpSPDHarmSublinkGroup.Add(RSLList[i]);
                                        SPDHarmGroupNumberOfSubLinks = (int)(Math.Ceiling((SPDHarmGroupHarmonizedSpeed * clsGlobalVars.DSD) / (clsGlobalVars.SubLinkLength * 3600)));
                                    }
                                    // else if the number of sublinks in the harmonization group < number of sublinks required for DSD
                                    else if (tmpNumberofSublinks < SPDHarmGroupNumberOfSubLinks)
                                    {
                                        //Lower the speed of the current sublink to the current harmonization group speed, add the sublink to the group and continue
                                        tmpNumberofSublinks++;
                                        RSLList[i].HarmonizedSpeed = SPDHarmGroupHarmonizedSpeed;
                                        tmpSPDHarmSublinkGroup.Add(RSLList[i]);
                                        RSLList[i].SpdHarmInclusionOverride = true;
                                    }
                                }
                                //else if current sublink troupe speed is < current harmonization group speed 
                                else if (RSLList[i].TroupeSpeed < (SPDHarmGroupHarmonizedSpeed)) 
                                {
                                    //if (tmpNumberofSublinks >= SPDHarmGroupNumberOfSubLinks)
                                    //{
                                    //}
                                    //if the number of sublinks in the current harmonization group is < number of sublinks required for DSD
                                    if (tmpNumberofSublinks < SPDHarmGroupNumberOfSubLinks)
                                    {
                                        //Assign previous group harmonized speed to the sublink if the speed of the last speed harmonization group is < than the sublink troupe speed
                                        if (LastSPDHarmGroupHarmonizedSpeed <= SPDHarmGroupHarmonizedSpeed)
                                        {
                                            foreach (clsRoadwaySubLink TSL in tmpSPDHarmSublinkGroup)
                                            {
                                                foreach (clsRoadwaySubLink RSL in RSLList)
                                                {
                                                    if (TSL.Identifier == RSL.Identifier)
                                                    {
                                                        RSL.HarmonizedSpeed = LastSPDHarmGroupHarmonizedSpeed;
                                                        RSL.SpdHarmInclusionOverride = true;
                                                        break;
                                                    }
                                                }
                                            }
                                            LastSPDHarmGroupHarmonizedSpeed = SPDHarmGroupHarmonizedSpeed;
                                            SPDHarmGroupHarmonizedSpeed = RSLList[i].TroupeSpeed;
                                            RSLList[i].HarmonizedSpeed = SPDHarmGroupHarmonizedSpeed;
                                            RSLList[i].BeginSpdHarm = true;
                                            tmpNumberofSublinks = 1;
                                            tmpSPDHarmSublinkGroup.Clear();
                                            tmpSPDHarmSublinkGroup.Add(RSLList[i]);
                                            SPDHarmGroupNumberOfSubLinks = (int)(Math.Ceiling((SPDHarmGroupHarmonizedSpeed * clsGlobalVars.DSD) / (clsGlobalVars.SubLinkLength * 3600)));
                                        }
                                        else
                                        {
                                            SPDHarmGroupHarmonizedSpeed = RSLList[i].TroupeSpeed;
                                            foreach (clsRoadwaySubLink TSL in tmpSPDHarmSublinkGroup)
                                            {
                                                foreach (clsRoadwaySubLink RSL in RSLList)
                                                {
                                                    if (TSL.Identifier == RSL.Identifier)
                                                    {
                                                        RSL.HarmonizedSpeed = SPDHarmGroupHarmonizedSpeed;
                                                        RSL.SpdHarmInclusionOverride = true;
                                                        break;
                                                    }
                                                }
                                            }
                                            tmpSPDHarmSublinkGroup.Add(RSLList[i]);
                                            tmpNumberofSublinks++;
                                        }
                                    }
                                    else if (tmpNumberofSublinks >= SPDHarmGroupNumberOfSubLinks)
                                    {
                                        //start a new spd harm group
                                        LastSPDHarmGroupHarmonizedSpeed = SPDHarmGroupHarmonizedSpeed;
                                        SPDHarmGroupHarmonizedSpeed = RSLList[i].TroupeSpeed;
                                        RSLList[i].HarmonizedSpeed = SPDHarmGroupHarmonizedSpeed;
                                        RSLList[i].BeginSpdHarm = true;
                                        tmpNumberofSublinks = 1;
                                        tmpSPDHarmSublinkGroup.Clear();
                                        tmpSPDHarmSublinkGroup.Add(RSLList[i]);
                                        SPDHarmGroupNumberOfSubLinks = (int)(Math.Ceiling((SPDHarmGroupHarmonizedSpeed * clsGlobalVars.DSD) / (clsGlobalVars.SubLinkLength * 3600)));
                                    }
                                }
                            }
                        }
                        else if (Roadway.Direction != Roadway.MMIncreasingDirection)
                        {
                            if (RSLList[i].BeginMM > TroupingEndMM)
                            {
                                if (RSLList[i].TroupeSpeed == SPDHarmGroupHarmonizedSpeed)
                                {
                                    //Current sublink troupe speed = current sublink harmonization group  speed
                                    //Add the sublink to the sublink harmonization group, Change the sublink harmonizedspeed and continue
                                    RSLList[i].HarmonizedSpeed = SPDHarmGroupHarmonizedSpeed;
                                    tmpSPDHarmSublinkGroup.Add(RSLList[i]);
                                    tmpNumberofSublinks++;
                                }
                                //If the current sublink troupe speed > current sublink harmonization group speed
                                else if (RSLList[i].TroupeSpeed > SPDHarmGroupHarmonizedSpeed)
                                {   //if the number of sublink in the harmonization group > number sublinks required for DSD
                                    if (tmpNumberofSublinks >= SPDHarmGroupNumberOfSubLinks)
                                    {
                                        //End the current sublink harmonization group and start a new sublink harmonization group and add the current sublink to it.
                                        SpeedDiff = RSLList[i].TroupeSpeed - SPDHarmGroupHarmonizedSpeed;

                                        LastSPDHarmGroupHarmonizedSpeed = SPDHarmGroupHarmonizedSpeed;
                                        if (SpeedDiff > 5) // greater than fivemiles per hour
                                        {
                                            SPDHarmGroupHarmonizedSpeed = SPDHarmGroupHarmonizedSpeed + 5;
                                        }
                                        else
                                        {
                                            SPDHarmGroupHarmonizedSpeed = RSLList[i].TroupeSpeed;
                                        }
                                        RSLList[i].HarmonizedSpeed = SPDHarmGroupHarmonizedSpeed;
                                        RSLList[i].BeginSpdHarm = true;
                                        tmpNumberofSublinks = 1;
                                        tmpSPDHarmSublinkGroup.Clear();
                                        tmpSPDHarmSublinkGroup.Add(RSLList[i]);
                                        SPDHarmGroupNumberOfSubLinks = (int)(Math.Ceiling((SPDHarmGroupHarmonizedSpeed * clsGlobalVars.DSD) / (clsGlobalVars.SubLinkLength * 3600)));
                                    }
                                    // else if the number of sublinks in the harmonization group < number of sublinks required for DSD
                                    else if (tmpNumberofSublinks < SPDHarmGroupNumberOfSubLinks)
                                    {
                                        //Lower the speed of the current sublink to the current harmonization group speed, add the sublink to the group and continue
                                        tmpNumberofSublinks++;
                                        RSLList[i].HarmonizedSpeed = SPDHarmGroupHarmonizedSpeed;
                                        tmpSPDHarmSublinkGroup.Add(RSLList[i]);
                                        RSLList[i].SpdHarmInclusionOverride = true;
                                    }
                                }
                                //else if current sublink troupe speed is < current harmonization group speed 
                                else if (RSLList[i].TroupeSpeed < (SPDHarmGroupHarmonizedSpeed))
                                {
                                    //if (tmpNumberofSublinks >= SPDHarmGroupNumberOfSubLinks)
                                    //{
                                    //}
                                    //if the number of sublinks in the current harmonization group is < number of sublinks required for DSD
                                    if (tmpNumberofSublinks < SPDHarmGroupNumberOfSubLinks)
                                    {
                                        //Assign previous group harmonized speed to the sublink if the speed of the last speed harmonization group is < than the sublink troupe speed
                                        if (LastSPDHarmGroupHarmonizedSpeed <= SPDHarmGroupHarmonizedSpeed)
                                        {
                                            foreach (clsRoadwaySubLink TSL in tmpSPDHarmSublinkGroup)
                                            {
                                                foreach (clsRoadwaySubLink RSL in RSLList)
                                                {
                                                    if (TSL.Identifier == RSL.Identifier)
                                                    {
                                                        RSL.HarmonizedSpeed = LastSPDHarmGroupHarmonizedSpeed;
                                                        RSL.SpdHarmInclusionOverride = true;
                                                        break;
                                                    }
                                                }
                                            }
                                            LastSPDHarmGroupHarmonizedSpeed = SPDHarmGroupHarmonizedSpeed;
                                            SPDHarmGroupHarmonizedSpeed = RSLList[i].TroupeSpeed;
                                            RSLList[i].HarmonizedSpeed = SPDHarmGroupHarmonizedSpeed;
                                            RSLList[i].BeginSpdHarm = true;
                                            tmpNumberofSublinks = 1;
                                            tmpSPDHarmSublinkGroup.Clear();
                                            tmpSPDHarmSublinkGroup.Add(RSLList[i]);
                                            SPDHarmGroupNumberOfSubLinks = (int)(Math.Ceiling((SPDHarmGroupHarmonizedSpeed * clsGlobalVars.DSD) / (clsGlobalVars.SubLinkLength * 3600)));
                                        }
                                        else
                                        {
                                            SPDHarmGroupHarmonizedSpeed = RSLList[i].TroupeSpeed;
                                            foreach (clsRoadwaySubLink TSL in tmpSPDHarmSublinkGroup)
                                            {
                                                foreach (clsRoadwaySubLink RSL in RSLList)
                                                {
                                                    if (TSL.Identifier == RSL.Identifier)
                                                    {
                                                        RSL.HarmonizedSpeed = SPDHarmGroupHarmonizedSpeed;
                                                        RSL.SpdHarmInclusionOverride = true;
                                                        break;
                                                    }
                                                }
                                            }
                                            tmpSPDHarmSublinkGroup.Add(RSLList[i]);
                                            tmpNumberofSublinks++;
                                        }
                                    }
                                    else if (tmpNumberofSublinks >= SPDHarmGroupNumberOfSubLinks)
                                    {
                                        //start a new spd harm group
                                        LastSPDHarmGroupHarmonizedSpeed = SPDHarmGroupHarmonizedSpeed;
                                        SPDHarmGroupHarmonizedSpeed = RSLList[i].TroupeSpeed;
                                        RSLList[i].HarmonizedSpeed = SPDHarmGroupHarmonizedSpeed;
                                        RSLList[i].BeginSpdHarm = true;
                                        tmpNumberofSublinks = 1;
                                        tmpSPDHarmSublinkGroup.Clear();
                                        tmpSPDHarmSublinkGroup.Add(RSLList[i]);
                                        SPDHarmGroupNumberOfSubLinks = (int)(Math.Ceiling((SPDHarmGroupHarmonizedSpeed * clsGlobalVars.DSD) / (clsGlobalVars.SubLinkLength * 3600)));
                                    }
                                }
                            }
                        }
                        //LogTxtMsg(txtCVDataLog, "\t\t\tRSL CVAvg speed: " + RSL.CVAvgSpeed.ToString("0") + "\tNumCVs: " + RSL.TotalNumberCVs + "\tNumQueuedCVs: " + RSL.NumberQueuedCVs + "\t\t%QueuedCVs: " + RSL.PercentQueuedCVs.ToString("00") +
                        //                             "\tQueued " + RSL.Queued + "\tDateProcessed: " + RSL.CVDateProcessed + "\tBeginMM: " + RSL.BeginMM);
                    }

                    if ((tmpSPDHarmSublinkGroup.Count > 0) && (SPDHarmGroupHarmonizedSpeed > 0))
                    {
                        //Assign previous group harmonized speed to the sublink if the speed of the last speed harmonization group is < than the sublink troupe speed
                        if ((LastSPDHarmGroupHarmonizedSpeed > 0)  && (LastSPDHarmGroupHarmonizedSpeed < SPDHarmGroupHarmonizedSpeed))
                        {
                            if (tmpNumberofSublinks < SPDHarmGroupNumberOfSubLinks)
                            {
                                foreach (clsRoadwaySubLink TSL in tmpSPDHarmSublinkGroup)
                                {
                                    foreach (clsRoadwaySubLink RSL in RSLList)
                                    {
                                        if (TSL.Identifier == RSL.Identifier)
                                        {
                                            RSL.HarmonizedSpeed = LastSPDHarmGroupHarmonizedSpeed;
                                            RSL.SpdHarmInclusionOverride = true;
                                            break;
                                        }
                                    }

                                }
                            }
                        }
                    }
                }
                else
                {
                    LogTxtMsg(txtSpdHarmLog, "\r\n" + DateTime.Now + "\tNo sublink was found with speed > 0 upstream of the congestion or queue location");
                }
            }
            catch (Exception ex)
            {
                retValue = "Error in processing roadway sublinks to determine harmonized speed." + "\r\n" + ex.Message;
                return retValue;
            }
            return retValue;
        }
        private string InsertSubLinkStatusIntoINFLODatabase()
        {
            string retValue = string.Empty;
            string DSID = string.Empty;
            string CurrRecord = string.Empty;
            try
            {
                string sqlStr = string.Empty;
                TimeSpan starttime = new TimeSpan(DateTime.Now.Ticks);
                LogTxtMsg(txtCVDataLog, "Adding roadway sublink dynamic information to INFLO database: ");

                //OleDbParameter parm = new OleDbParameter("?", OleDbType.Date);
                //parm.Value = DateTime.Now;
                //cmd.Parameters.Add(parm); 
                foreach (clsRoadwaySubLink tmpRSL in RSLList)
                {
                    sqlStr = "INSERT INTO TME_CVData_SubLink(RoadwayId, SubLinkId, DateProcessed, IntervalLength, TSSAvgSpeed, CVAvgSpeed, WRTMSpeed, RecommendedSpeed, RecommendedSpeedSource, CVQueued, CVCongested, NumberCVs,NumberQueuedCVs, PercentQueuedCVs) " +
                                "Values('" + tmpRSL.RoadwayID + "', '" + tmpRSL.Identifier + "', #" + tmpRSL.DateProcessed + "#, " + clsGlobalVars.CVDataPollingFrequency + ", " + tmpRSL.TSSAvgSpeed.ToString("0") + ", '" +
                                             tmpRSL.CVAvgSpeed + ", " + tmpRSL.WRTMSpeed.ToString("0.0") + ", " + tmpRSL.RecommendedSpeed.ToString("0") + ", '" + tmpRSL.RecommendedSpeedSource.ToString() + "', " + 
                                             tmpRSL.Congested + ", " + tmpRSL.Queued + ", " + tmpRSL.TotalNumberCVs + ", " + tmpRSL.NumberQueuedCVs + ", " + tmpRSL.PercentQueuedCVs.ToString("00") + ")";
                    retValue = DB.InsertRow(sqlStr);
                    if (retValue.Length > 0)
                    {
                        return retValue;
                    }
                    LogTxtMsg(txtCVDataLog, "\t\tRoadway Sublink data added to INFLO database: " + tmpRSL.RoadwayID + ", " + tmpRSL.Identifier + ", " + tmpRSL.DateProcessed + ", " + clsGlobalVars.CVDataPollingFrequency + ", " + tmpRSL.TSSAvgSpeed + ", " + 
                                                tmpRSL.CVAvgSpeed + ", " + tmpRSL.WRTMSpeed + ", " + tmpRSL.RecommendedSpeed + ", " + tmpRSL.RecommendedSpeedSource + ", " + 
                                                tmpRSL.Congested + ", " + tmpRSL.Queued + ", " + tmpRSL.TotalNumberCVs + ", " + tmpRSL.NumberQueuedCVs + ", " + tmpRSL.PercentQueuedCVs.ToString("00") );
                }
                TimeSpan endtime = new TimeSpan(DateTime.Now.Ticks);
                LogTxtMsg(txtCVDataLog, "Time for adding: " + RSLList.Count + " Sublink records into database" + (endtime.TotalMilliseconds - starttime.TotalMilliseconds).ToString("0") + " msecs");
            }
            catch (Exception ex)
            {
                retValue = "Error in adding Sublink data into INFLO database. Error Hint - Current record: " + CurrRecord + "\r\n" + ex.Message;
                return retValue;
            }
            return retValue;
        }

        private string GetLastIntervalDetectionStationData(clsDatabase DB, ref List<clsDetectorStation> DSList, ref int NumberRecordsRetrieved)
        {
            string retValue = string.Empty;
            string sqlQuery = string.Empty;
            //List<clsDetectionZone> LastIntervalDZList = new List<clsDetectionZone>();
            //clear old detection station records
            DSList.Clear();

            DataSet DetectorStationDataSet = new DataSet("DetectorStation");
            //IntervalUTCTime = IntervalUTCTime.AddSeconds(-clsGlobalVars.TSSDataPollingFrequency);
            if (clsGlobalVars.DBInterfaceType.ToLower() == "oledb")
            {
                sqlQuery = "Select * from TME_TSSData_Input";
                //sqlQuery = "Select * from TME_TSSData_Input where DateReceived>=#" + IntervalUTCTime.AddSeconds(-clsGlobalVars.TSSDataPollingFrequency) + "#";
            }
            else if (clsGlobalVars.DBInterfaceType.ToLower() == "sqlserver")
            {
                sqlQuery = "Select * from TME_TSSData_Input";
                //sqlQuery = "Select * from TME_TSSData_Input where DateReceived>='" + IntervalUTCTime.AddSeconds(-clsGlobalVars.TSSDataPollingFrequency) + "'";
            }

            retValue = string.Empty;
            try
            {
                retValue = DB.FillDataSet(sqlQuery, ref DetectorStationDataSet);
                if (DetectorStationDataSet.Tables[0] != null)
                {
                    FillDataSetLog.WriteLine(DateTime.Now + ",GetLastIntervalDetectionZoneStatus," + DetectorStationDataSet.Tables[0].Rows.Count);
                    if (retValue.Length > 0)
                    {
                        return retValue;
                    }

                    //display Detector station information 
                    LogTxtMsg(txtINFLOConfigLog, "\tAvailable detector station records: ");

                    NumberRecordsRetrieved = 0;
                    if (DetectorStationDataSet.Tables[0].Rows.Count > 0)
                    {
                        foreach (DataRow row in DetectorStationDataSet.Tables[0].Rows)
                        {
                            clsDetectorStation tmpDS = new clsDetectorStation();
                            string DSID = string.Empty;
                            string DZID = string.Empty;
                            double MMLocation = 0;
                            string DZStatus = string.Empty;
                            string DataType = string.Empty;
                            DateTime DateReceived = DateTime.UtcNow;
                            DateTime BeginTime = DateTime.UtcNow;
                            DateTime EndTime = DateTime.UtcNow;
                            int StartInterval = 0;
                            int EndInterval = 0;
                            int Volume = 0;
                            int AvgSpeed = 0;
                            double Occ = 0;
                            bool Queued = false;
                            bool Congested = false;
                            int IntervalLength = 0;

                            foreach (DataColumn col in DetectorStationDataSet.Tables[0].Columns)
                            {
                                switch (col.ColumnName.ToString().ToLower())
                                {
                                    case "dsid":
                                        if (row[col].ToString().Length > 0)
                                        {
                                            DSID = row[col].ToString();
                                            tmpDS.Identifier = int.Parse(DSID);
                                        }
                                        break;
                                    case "mmlocation":
                                        if (row[col].ToString().Length > 0)
                                        {
                                            MMLocation = double.Parse(row[col].ToString());
                                            tmpDS.MMLocation = MMLocation;
                                        }
                                        break;
                                    case "datereceived":
                                        if (row[col].ToString().Length > 0)
                                        {
                                            DateReceived = DateTime.Parse(row[col].ToString());
                                            tmpDS.DateReceived = DateReceived;
                                        }
                                        break;
                                    case "startinterval":
                                        if (row[col].ToString().Length > 0)
                                        {
                                            StartInterval = int.Parse(row[col].ToString());
                                            tmpDS.StartInterval = StartInterval;
                                        }
                                        break;
                                    case "endinterval":
                                        if (row[col].ToString().Length > 0)
                                        {
                                            EndInterval = int.Parse(row[col].ToString());
                                            tmpDS.EndInterval = EndInterval;
                                        }
                                        break;
                                    case "intervallength":
                                        if (row[col].ToString().Length > 0)
                                        {
                                            IntervalLength = int.Parse(row[col].ToString());
                                            tmpDS.IntervalLength = IntervalLength;
                                        }
                                        break;
                                    case "begintime":
                                        if (row[col].ToString().Length > 0)
                                        {
                                            BeginTime = DateTime.Parse(row[col].ToString());
                                            tmpDS.BeginTime = BeginTime;
                                        }
                                        break;
                                    case "endtime":
                                        if (row[col].ToString().Length > 0)
                                        {
                                            EndTime = DateTime.Parse(row[col].ToString());
                                            tmpDS.EndTime = EndTime;
                                        }
                                        break;
                                    case "volume":
                                        if (row[col].ToString().Length > 0)
                                        {
                                            Volume = int.Parse(row[col].ToString());
                                            tmpDS.Volume = Volume;
                                        }
                                        break;
                                    case "avgspeed":
                                        if (row[col].ToString().Length > 0)
                                        {
                                            AvgSpeed = int.Parse(row[col].ToString());
                                            tmpDS.AvgSpeed = AvgSpeed;
                                        }
                                        break;
                                    case "occupancy":
                                        if (row[col].ToString().Length > 0)
                                        {
                                            Occ = double.Parse(row[col].ToString());
                                            tmpDS.Occupancy = Occ;
                                        }
                                        break;
                                    case "queued":
                                        Queued = bool.Parse(row[col].ToString());
                                        tmpDS.Queued = Queued;
                                        break;
                                    case "congested":
                                        Congested = bool.Parse(row[col].ToString());
                                        tmpDS.Congested = Congested;
                                        break;
                                }
                            }
                            //LogTxtMsg(txtTSSDataLog, "\t\t" + DSID + ", " + DZID + ", " + DateReceived + ", " + StartInterval + ", " + EndInterval + ", " + BeginTime + ", " + EndTime + ", " +
                            //                                      IntervalLength + ", " + Volume + ", " + AvgSpeed + ", " + Occ + ", " + Queued + ", " + Congested + ", " + DZStatus + ", " + DataType + ", " + MMLocation);
                            DSList.Add(tmpDS);
                            NumberRecordsRetrieved = NumberRecordsRetrieved + 1;
                        }
                    }
                }
                else
                {
                    retValue = "Error in retrieving Detector Station data for the last interval. No data was retrieved.";
                    return retValue;
                }
            }
            catch (Exception exc)
            {
                retValue = "Error in retrieving Detector Station data for the last interval." + "\r\n\t" + exc.Message;
                return retValue;
            }

            return retValue;
        }
        private string GetLastIntervalDetectionZoneStatus(clsDatabase DB, DateTime IntervalUTCTime, ref List<clsDetectionZone> DZList, ref int NumberRecordsRetrieved)
        {
            string retValue = string.Empty;
            string sqlQuery = string.Empty;
            List<clsDetectionZone> LastIntervalDZList = new List<clsDetectionZone>();

            DataSet DetectionZonesDataSet = new DataSet("DetectionZones");
            //IntervalUTCTime = IntervalUTCTime.AddSeconds(-clsGlobalVars.TSSDataPollingFrequency);
            if (clsGlobalVars.DBInterfaceType.ToLower() == "oledb")
            {
                sqlQuery = "Select * from TME_TSSData_Input where DateReceived>=#" + IntervalUTCTime.AddSeconds(-clsGlobalVars.TSSDataPollingFrequency) + "#";
            }
            else if (clsGlobalVars.DBInterfaceType.ToLower() == "sqlserver")
            {
                sqlQuery = "Select * from TME_TSSData_Input where DateReceived>='" + IntervalUTCTime.AddSeconds(-clsGlobalVars.TSSDataPollingFrequency) + "'";
            }

            retValue = string.Empty;
            try
            {
                retValue = DB.FillDataSet(sqlQuery, ref DetectionZonesDataSet);
                if (DetectionZonesDataSet.Tables[0] != null)
                {
                    FillDataSetLog.WriteLine(DateTime.Now + ",GetLastIntervalDetectionZoneStatus," + DetectionZonesDataSet.Tables[0].Rows.Count + "," + IntervalUTCTime);
                    if (retValue.Length > 0)
                    {
                        return retValue;
                    }

                    //display Detector station information 
                    LogTxtMsg(txtINFLOConfigLog, "\tAvailable detection zone records: ");

                    NumberRecordsRetrieved = 0;
                    if (DetectionZonesDataSet.Tables[0].Rows.Count > 0)
                    {
                        foreach (DataRow row in DetectionZonesDataSet.Tables[0].Rows)
                        {
                            clsDetectionZone tmpDZ = new clsDetectionZone();
                            string DSID = string.Empty;
                            string DZID = string.Empty;
                            double MMLocation = 0;
                            string DZStatus = string.Empty;
                            string DataType = string.Empty;
                            DateTime DateReceived = DateTime.UtcNow;
                            DateTime BeginTime = DateTime.UtcNow;
                            DateTime EndTime = DateTime.UtcNow;
                            int StartInterval = 0;
                            int EndInterval = 0;
                            int Volume = 0;
                            int AvgSpeed = 0;
                            double Occ = 0;
                            bool Queued = false;
                            bool Congested = false;
                            int IntervalLength = 0;

                            foreach (DataColumn col in DetectionZonesDataSet.Tables[0].Columns)
                            {
                                switch (col.ColumnName.ToString().ToLower())
                                {
                                    case "dsid":
                                        if (row[col].ToString().Length > 0)
                                        {
                                            DSID = row[col].ToString();
                                            tmpDZ.DSIdentifier = int.Parse(DSID);
                                        }
                                        break;
                                    case "dzid":
                                        if (row[col].ToString().Length > 0)
                                        {
                                            DZID = row[col].ToString();
                                            tmpDZ.Identifier = int.Parse(DZID);
                                        }
                                        break;
                                    case "mmlocation":
                                        if (row[col].ToString().Length > 0)
                                        {
                                            MMLocation = double.Parse(row[col].ToString());
                                            tmpDZ.MMLocation = MMLocation;
                                        }
                                        break;
                                    case "datereceived":
                                        if (row[col].ToString().Length > 0)
                                        {
                                            DateReceived = DateTime.Parse(row[col].ToString());
                                            tmpDZ.DateReceived = DateReceived;
                                        }
                                        break;
                                    case "startinterval":
                                        if (row[col].ToString().Length > 0)
                                        {
                                            StartInterval = int.Parse(row[col].ToString());
                                            tmpDZ.StartInterval = StartInterval;
                                        }
                                        break;
                                    case "endinterval":
                                        if (row[col].ToString().Length > 0)
                                        {
                                            EndInterval = int.Parse(row[col].ToString());
                                            tmpDZ.EndInterval = EndInterval;
                                        }
                                        break;
                                    case "intervallength":
                                        if (row[col].ToString().Length > 0)
                                        {
                                            IntervalLength = int.Parse(row[col].ToString());
                                            tmpDZ.IntervalLength = IntervalLength;
                                        }
                                        break;
                                    case "begintime":
                                        if (row[col].ToString().Length > 0)
                                        {
                                            BeginTime = DateTime.Parse(row[col].ToString());
                                            tmpDZ.BeginTime = BeginTime;
                                        }
                                        break;
                                    case "endtime":
                                        if (row[col].ToString().Length > 0)
                                        {
                                            EndTime = DateTime.Parse(row[col].ToString());
                                            tmpDZ.EndTime = EndTime;
                                        }
                                        break;
                                    case "volume":
                                        if (row[col].ToString().Length > 0)
                                        {
                                            Volume = int.Parse(row[col].ToString());
                                            tmpDZ.Volume = Volume;
                                        }
                                        break;
                                    case "avgspeed":
                                        if (row[col].ToString().Length > 0)
                                        {
                                            AvgSpeed = int.Parse(row[col].ToString());
                                            tmpDZ.AvgSpeed = AvgSpeed;
                                        }
                                        break;
                                    case "occupancy":
                                        if (row[col].ToString().Length > 0)
                                        {
                                            Occ = double.Parse(row[col].ToString());
                                            tmpDZ.Occupancy = Occ;
                                        }
                                        break;
                                    case "queued":
                                        Queued = bool.Parse(row[col].ToString());
                                        tmpDZ.Queued = Queued;
                                        break;
                                    case "congested":
                                        Congested = bool.Parse(row[col].ToString());
                                        tmpDZ.Congested = Congested;
                                        break;
                                    case "DZStatus":
                                        DZStatus = row[col].ToString();
                                        tmpDZ.DZStatus = DZStatus;
                                        break;
                                    case "DataType":
                                        DataType = row[col].ToString();
                                        tmpDZ.DataType = DataType;
                                        break;
                                }
                            }
                            //LogTxtMsg(txtTSSDataLog, "\t\t" + DSID + ", " + DZID + ", " + DateReceived + ", " + StartInterval + ", " + EndInterval + ", " + BeginTime + ", " + EndTime + ", " +
                            //                                      IntervalLength + ", " + Volume + ", " + AvgSpeed + ", " + Occ + ", " + Queued + ", " + Congested + ", " + DZStatus + ", " + DataType + ", " + MMLocation);
                            LastIntervalDZList.Add(tmpDZ);
                        }
                        foreach (clsDetectionZone DZ in DZList)
                        {
                            DZ.AvgSpeed = 0;
                            DZ.Occupancy = 0;
                            DZ.Volume = 0;
                            DZ.Queued = false;
                            DZ.Congested = false;
                            DZ.DateReceived = DateTime.Now;
                            DZ.BeginTime = DateTime.Now;
                            DZ.EndTime = DateTime.Now;
                            DZ.StartInterval = 0;
                            DZ.EndInterval = 0;
                            DZ.IntervalLength = 0;
                            DZ.DZStatus = "NoNewData";

                            foreach (clsDetectionZone NewDZ in LastIntervalDZList)
                            {
                                if (DZ.Identifier == NewDZ.Identifier)
                                {
                                    DZ.AvgSpeed = NewDZ.AvgSpeed;
                                    DZ.Occupancy = NewDZ.Occupancy;
                                    DZ.Volume = NewDZ.Volume;
                                    DZ.Queued = NewDZ.Queued;
                                    DZ.Congested = NewDZ.Congested;
                                    DZ.DateReceived = NewDZ.DateReceived;
                                    DZ.BeginTime = NewDZ.BeginTime;
                                    DZ.EndTime = NewDZ.EndTime;
                                    DZ.StartInterval = NewDZ.StartInterval;
                                    DZ.EndInterval = NewDZ.EndInterval;
                                    DZ.IntervalLength = NewDZ.IntervalLength;
                                    DZ.DZStatus = "NewData";
                                    NumberRecordsRetrieved = NumberRecordsRetrieved + 1;
                                    break;
                                }
                            }
                            if (DZ.DZStatus.ToUpper() == "NewData".ToUpper())
                            {
                                DZ.NoNewData = false;
                                DZ.NumberNoNewDataIntervals = 0;
                            }
                            else if (DZ.DZStatus.ToUpper() == "NoNewData".ToUpper())
                            {
                                DZ.NoNewData = true;
                                DZ.NumberNoNewDataIntervals = DZ.NumberNoNewDataIntervals + 1;
                            }

                        }
                    }
                }
                else
                {
                    retValue = "Error in retrieving Detection Zone data for the last interval from: " + IntervalUTCTime.AddSeconds(-clsGlobalVars.TSSDataPollingFrequency) + ". No data was retrieved.";
                    return retValue;
                }
            }
            catch (Exception exc)
            {
                retValue = "Error in retrieving Detection Zone data for the last interval from: " + IntervalUTCTime.AddSeconds(-clsGlobalVars.TSSDataPollingFrequency) + "\r\n\t" + exc.Message;
                return retValue;
            }

            return retValue;
        }

        private string ProcessDetectorStationStatus()
        {
            string retValue = string.Empty;
            double TotalSpeed = 0;
            int TotalVol = 0;
            double TotalOcc = 0;
            double MaxOcc = 0;
            int NumDZs = 0;
            string DZID = string.Empty;
            string CurrRecord = string.Empty;
            DateTime DateReceived = DateTime.Now;
            DateTime BeginTime = DateTime.Now;
            DateTime EndTime = DateTime.Now;
            int StartInterval = 0;
            int EndInterval = 0;
            int IntervalLength = 0;
            try
            {
                foreach (clsDetectorStation DS in DSList)
                {
                    TotalSpeed = 0;
                    TotalVol = 0;
                    TotalOcc = 0;
                    MaxOcc = 0;
                    NumDZs = 0;
                    DS.AvgSpeed = 0;
                    DateReceived = DateTime.Now.Date;
                    BeginTime = DateTime.Now.Date;
                    EndTime = DateTime.Now.Date;
                    StartInterval = 0;
                    EndInterval = 0;
                    IntervalLength = 0;
                    DS.Volume = 0;
                    DS.Congested = false;
                    DS.Queued = false;
                    DS.Occupancy = 0;
                    DS.DateReceived = DateTime.Now.Date;
                    DS.BeginTime = DateTime.Now.Date;
                    DS.EndTime = DateTime.Now.Date;
                    DS.StartInterval = 0;
                    DS.EndInterval = 0;
                    DS.IntervalLength = 0;
                    //LogTxtMsg(txtTSSDataLog, "\t\t\tProcessing detector station: " + DS.Identifier);

                    string[] sField;
                    sField = DS.DetectionZones.Split(',');

                    foreach (clsDetectionZone DZ in DZList)
                    {
                        //LogTxtMsg(txtTSSDataLog, "\t\t\t\tDetection zone: " + DZ.Identifier);
                        CurrRecord = "DS: " + DS.Identifier + "   DZ: " + DZ.Identifier;
                        DZID = DZ.Identifier.ToString();
                        if (sField.Length > 0)
                        {
                            for (int j = 0; j < sField.Length; j++)
                            {
                                if (sField[j].ToUpper() == DZID.ToUpper())
                                {
                                    if (Roadway.Direction == DZ.Direction)
                                    {
                                        //if (DZ.AvgSpeed > 0)
                                        //{
                                        TotalSpeed = TotalSpeed + (DZ.AvgSpeed * DZ.Volume);
                                        TotalVol = TotalVol + DZ.Volume;
                                        if (DZ.Occupancy > MaxOcc)
                                        {
                                            MaxOcc = DZ.Occupancy;
                                        }
                                        TotalOcc = TotalOcc + DZ.Occupancy;
                                        NumDZs = NumDZs + 1;
                                        DateReceived = DZ.DateReceived;
                                        BeginTime = DZ.BeginTime;
                                        EndTime = DZ.EndTime;
                                        StartInterval = DZ.StartInterval;
                                        EndInterval = DZ.EndInterval;
                                        IntervalLength = DZ.IntervalLength;
                                        //}
                                    }
                                    break;
                                }
                            }
                        }
                    }
                    if (NumDZs > 0)
                    {
                        //DS.AvgSpeed = TotalSpeed / NumDZs;
                        if (TotalVol > 0)
                        {
                            DS.AvgSpeed = TotalSpeed / TotalVol;
                        }
                        else
                        {
                            DS.AvgSpeed = 0;
                        }
                        DS.Volume = TotalVol;
                        DS.Occupancy = MaxOcc;
                        //DS.Occupancy = TotalOcc / NumDZs;
                        DS.EndTime = EndTime;
                        DS.BeginTime = BeginTime;
                        DS.StartInterval = StartInterval;
                        DS.EndInterval = EndInterval;
                        DS.DateReceived = DateReceived;
                        DS.IntervalLength = IntervalLength;
                        //LogTxtMsg(txtTSSDataLog, "\t\t\tDS Avg speed: " + DS.AvgSpeed.ToString("0") + "\tVolume: " + DS.Volume + "\tOccupancy: " + DS.Occupancy.ToString("0.0") + " " + DS.BeginTime + " " + DS.EndTime + " " +
                       //                                                   DS.StartInterval + " " + DS.EndInterval + " " + DS.IntervalLength + " " + DS.DateReceived);
                        //Update Congested State and Queued state
                    }
                }
            }
            catch (Exception ex)
            {
                retValue = "Error in processing detector station status from detection zones. Current record: " + CurrRecord + "\r\n" + ex.Message;
                return retValue;
            }
            return retValue;
        }
        private string ProcessLinkInfrastructureStatus()
        {
            string retValue = string.Empty;
            string DSID = string.Empty;
            string CurrRecord = string.Empty;
            string RLDS = string.Empty;
            try
            {
                foreach (clsRoadwayLink RL in RLList)
                {
                    RL.TSSAvgSpeed = 0;
                    RL.Volume = 0;
                    RL.Occupancy = 0;
                    RL.Congested = false;
                    RL.Queued = false;
                    RL.Congested = false;
                    RLDS = RL.DetectionStations;
                    foreach (clsDetectorStation DS in DSList)
                    {
                        //LogTxtMsg(txtTSSDataLog, "\t\t\t\tDetector station: " + DS.Identifier);
                        CurrRecord = "RL: " + RL.Identifier + "   DS: " + DS.Identifier;
                        DSID = DS.Identifier.ToString();
                        if (RLDS == DSID)
                        {
                            //LogTxtMsg(txtTSSDataLog, "\t\t\tProcessing roadway link: " + RL.Identifier + "\tFromMM: " + RL.BeginMM + "\tToMM: " + RL.EndMM + "\tRLDS: " + RL.DetectionStations + "\tDS: " + DS.Identifier);
                            if ((DS.Occupancy < clsGlobalVars.OccupancyThreshold) && (DS.Volume < clsGlobalVars.VolumeThreshold))
                            {
                                RL.TSSAvgSpeed = clsGlobalVars.MaximumDisplaySpeed;
                                RL.Volume = DS.Volume;
                                RL.Occupancy = DS.Occupancy;
                            }
                            else if ((DS.Occupancy < clsGlobalVars.OccupancyThreshold) || (DS.Volume < clsGlobalVars.VolumeThreshold))
                            {
                                RL.TSSAvgSpeed = DS.AvgSpeed;
                                RL.Volume = DS.Volume;
                                RL.Occupancy = DS.Occupancy;
                            }
                            else if ((DS.Occupancy >= clsGlobalVars.OccupancyThreshold) && (DS.Volume >= clsGlobalVars.VolumeThreshold))
                            {
                                RL.TSSAvgSpeed = DS.AvgSpeed;
                                RL.Volume = DS.Volume;
                                RL.Occupancy = DS.Occupancy;
                            }
                            if (RL.TSSAvgSpeed == 0)
                            {
                                RL.TSSAvgSpeed = clsGlobalVars.MaximumDisplaySpeed;
                            }
                            else if (RL.TSSAvgSpeed > clsGlobalVars.MaximumDisplaySpeed)
                            {
                                RL.TSSAvgSpeed = clsGlobalVars.MaximumDisplaySpeed;
                            }

                            if (RL.TSSAvgSpeed <= clsGlobalVars.LinkQueuedSpeedThreshold)
                            {
                                RL.Queued = true;
                            }
                            else if (RL.TSSAvgSpeed <= clsGlobalVars.LinkCongestedSpeedThreshold)
                            {
                                RL.Congested = true;
                            }
                            RL.DateProcessed = DateTime.UtcNow;
                            RL.StartInterval = DS.StartInterval;
                            RL.EndInterval = DS.EndInterval;
                            //LogTxtMsg(txtTSSDataLog, "\t\t\tRL Avg speed: " + RL.TSSAvgSpeed + "\tVolume: " + RL.Volume + "\tOccupancy: " + RL.Occupancy + "\tQueued: " + RL.Queued +
                            //                             "\tCongested: " + RL.Congested + "\tDateProcessed: " + RL.DateProcessed);
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                retValue = "Error in processing roadway link status from detector station. Error Hint - Current record: " + CurrRecord + "\r\n" + ex.Message;
                return retValue;
            }
            return retValue;
        }
        private string InsertLinkStatusIntoINFLODatabase()
        {
            string retValue = string.Empty;
            string DSID = string.Empty;
            string CurrRecord = string.Empty;
            try
            {
                string sqlStr = string.Empty;
                TimeSpan starttime = new TimeSpan(DateTime.Now.Ticks);

                //OleDbParameter parm = new OleDbParameter("?", OleDbType.Date);
                //parm.Value = DateTime.Now;
                //cmd.Parameters.Add(parm); 
                foreach (clsRoadwayLink tmpRL in RLList)
                {
                    //DateTime tmpDate = DateTime.Parse(DateTime.Now.Year + "/" + DateTime.Now.Month + "/" + DateTime.Now.Day + " " + DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second);

                    sqlStr = "INSERT INTO TME_TSS_Link(RoadwayId, LinkId, DateProcessed, IntervalLength, TSSSpeed, TSSVolume, TSSOccupancy, WRTMSpeed, RecommendedSpeed, RecommendedSpeedSource, Congested, Queued) " +
                                "Values('" + tmpRL.RoadwayID + "', '" + tmpRL.Identifier + "', #" + tmpRL.DateProcessed + "#, " + clsGlobalVars.TSSDataLoadingFrequency + ", " + tmpRL.TSSAvgSpeed.ToString("0") + ", " +
                                             tmpRL.Volume + ", " + tmpRL.Occupancy.ToString("0.0") + ", " + tmpRL.WRTMSpeed.ToString("0") + ", " + tmpRL.RecommendedSpeed.ToString("0") + ", '" + 
                                             tmpRL.RecommendedSpeedSource.ToString() + "', " + tmpRL.Congested + ", " + tmpRL.Queued + ")";
                    retValue = DB.InsertRow(sqlStr);
                    if (retValue.Length > 0)
                    {
                        return retValue;
                    }
                    LogTxtMsg(txtTSSDataLog, "\t\tRoadway link TSS data added to database: " + tmpRL.RoadwayID + ", " + tmpRL.Identifier + ", " + tmpRL.TSSAvgSpeed + ", " + tmpRL.Volume + ", " +
                                                                                    tmpRL.Occupancy + ", " + tmpRL.WRTMSpeed + ", " + tmpRL.RecommendedSpeedSource + ", " + tmpRL.RecommendedSpeedSource + ", " + 
                                                                                    tmpRL.Congested + ", " + tmpRL.Queued + ", " + tmpRL.DateProcessed + ", " + clsGlobalVars.TSSDataLoadingFrequency);
                }
                TimeSpan endtime = new TimeSpan(DateTime.Now.Ticks);
                LogTxtMsg(txtTSSDataLog, "\r\n\tTime for adding: " + RLList.Count + " records into database" + (endtime.TotalMilliseconds - starttime.TotalMilliseconds).ToString("0") + " msecs");
            }
            catch (Exception ex)
            {
                retValue = "Error in adding TSS link data into INFLO database. Error Hint - Current record: " + CurrRecord + "\r\n" + ex.Message;
                return retValue;
            }
            return retValue;
        }

        private string InsertQueueInfoIntoINFLODatabase(double BOQMMLocation, DateTime BOQTime, double QueueRate, clsEnums.enQueueCahnge QueueChange, clsEnums.enQueueSource QueueSource, double QueueSpeed, clsRoadway Roadway)
        {
            string retValue = string.Empty;
            try
            {
                string sqlStr = string.Empty;
                TimeSpan starttime = new TimeSpan(DateTime.Now.Ticks);

                DateTime DateGenerated = TimeZoneInfo.ConvertTimeToUtc(BOQTime, TimeZoneInfo.Local);

                if (clsGlobalVars.DBInterfaceType.ToLower() == "sqlserver")
                {
                    //sqlStr = "INSERT INTO TMEOutput_QWARNMessage_CV(RoadwayID, DateGenerated, BOQMMLocation, FOQMMLocation, RateOfQueueGrowth, QueueGrowthDirection, SpeedInQueue, ValidityDuration) " +
                    //           "Values('" + Roadway.Identifier.ToString() + "', '" + DateGenerated + "', " + BOQMMLocation + ", " + Roadway.RecurringCongestionMMLocation + ", " +
                    //                        QueueRate.ToString("0") + ", '" + QueueChange.ToString() + "', " + QueueSpeed.ToString("0") + ", " + "20" + ")";
                    sqlStr = "INSERT INTO TMEOutput_QWARNMessage_CV(RoadwayID, DateGenerated, BOQMMLocation, FOQMMLocation, RateOfQueueGrowth, SpeedInQueue, ValidityDuration) " +
                                "Values('" + Roadway.Identifier.ToString() + "', '" + DateGenerated + "', " + BOQMMLocation + ", " + Roadway.RecurringCongestionMMLocation + ", " +
                                             QueueRate.ToString("0") + ", " + QueueSpeed.ToString("0") + ", " + "60" + ")";
                    retValue = DB.InsertRow(sqlStr);
                    if (retValue.Length > 0)
                    {
                        return retValue;
                    }
                    LogTxtMsg(txtINFLOLog, "\t\tQueue Info added to database: " + Roadway.Identifier.ToString() + ", " + BOQTime.ToString() + ", " + BOQMMLocation + ", " + Roadway.RecurringCongestionMMLocation + ", " +
                                                                                    QueueRate.ToString("0") + ", " + QueueChange + ", " + QueueSpeed.ToString("0") + ", " + "60");
                }
                else if (clsGlobalVars.DBInterfaceType.ToLower() == "oledb")
                {
                    sqlStr = "INSERT INTO TMEOutput_QWARNMessage_CV(RoadwayID, DateGenerated, BOQMMLocation, FOQMMLocation, RateOfQueueGrowth, QueueGrowthDirection, SpeedInQueue, ValidityDuration) " +
                                "Values('" + Roadway.Identifier.ToString() + "', #" + DateGenerated + "#, " + BOQMMLocation + ", " + Roadway.RecurringCongestionMMLocation + ", " +
                                             QueueRate.ToString("0") + ", '" + QueueChange.ToString() + "', " + QueueSpeed.ToString("0") + ", " + "60" + ")";
                    retValue = DB.InsertRow(sqlStr);
                    if (retValue.Length > 0)
                    {
                        return retValue;
                    }
                    LogTxtMsg(txtINFLOLog, "\t\tQueue Info added to database: " + Roadway.Identifier.ToString() + ", " + BOQTime.ToString() + ", " + BOQMMLocation + ", " + Roadway.RecurringCongestionMMLocation + ", " +
                                                                                    QueueRate.ToString("0") + ", " + QueueChange + ", " + QueueSpeed.ToString("0") + ", " + "60");
                }
                TimeSpan endtime = new TimeSpan(DateTime.Now.Ticks);
                LogTxtMsg(txtINFLOLog, "\t\tTime for adding Queue info to database: " + (endtime.TotalMilliseconds - starttime.TotalMilliseconds).ToString("0") + " msecs");
            }
            catch (Exception ex)
            {
                retValue = "\tError in adding Queue info into INFLO database." + "\r\n\t" + ex.Message;
                return retValue;
            }
            return retValue;
        }
        private string GenerateSPDHarmMessages(List<clsRoadwaySubLink> RSLList, double TroupingEndMM, double BOQMMLocation, clsRoadway Roadway)
        {
            string retValue = string.Empty;

            double RecommendedSpeed = 0;
            double BeginMM = 0;
            string Justification = string.Empty;
            double EndMM = 0;

            try
            {
                RSLList.Sort((l, r) => l.BeginMM.CompareTo(r.BeginMM));
                if (BOQMMLocation > 0)
                {
                    Justification = "Queue";
                }
                else
                {
                    Justification = "Congestion";
                }

                RecommendedSpeed = RSLList[0].HarmonizedSpeed;
                BeginMM = RSLList[0].BeginMM;
                EndMM = RSLList[0].EndMM;

                for (int i = 1; i < RSLList.Count; i++)
                {
                    if (RSLList[i].BeginMM < TroupingEndMM)
                    {
                        if (RSLList[i].HarmonizedSpeed == RecommendedSpeed)
                        {
                            EndMM = RSLList[i].EndMM;
                        }
                        else if (RSLList[i].HarmonizedSpeed != RecommendedSpeed)
                        {
                            retValue = InsertSPDHarmInfoIntoINFLODatabase(Roadway.Identifier.ToString(), DateTime.Now.AddMilliseconds(100), RecommendedSpeed, BeginMM, EndMM, Justification);
                            if (retValue.Length > 0)
                            {
                                LogTxtMsg(txtCVDataLog, "Insert SPDHarmMsg: " + Roadway.Identifier + ", " + DateTime.Now.ToString() + ", " + RecommendedSpeed + ", " + BeginMM + ", " +
                                                        EndMM + ", " + Justification + ", 60" + "\r\n\t" + retValue);
                            }
                            RecommendedSpeed = RSLList[i].HarmonizedSpeed;
                            BeginMM = RSLList[i].BeginMM;
                            EndMM = RSLList[i].EndMM;
                        }
                    }
                    else if (RSLList[i].BeginMM >= TroupingEndMM)
                    {
                        break;
                    }
                }
                retValue = InsertSPDHarmInfoIntoINFLODatabase(Roadway.Identifier.ToString(), DateTime.Now.AddMilliseconds(100), RecommendedSpeed, BeginMM, EndMM, Justification);
                if (retValue.Length > 0)
                {
                    LogTxtMsg(txtCVDataLog, "Insert SPDHarmMsg: " + Roadway.Identifier + ", " + DateTime.Now.ToString() + ", " + RecommendedSpeed + ", " + BeginMM + ", " +
                                            EndMM + ", " + Justification + ", 60" + "\r\n\t" + retValue);
                }
            }
            catch (Exception ex)
            {
                retValue = "\tError in g SPDenerating SPD Harm message info." + "\r\n\t" + ex.Message;
                return retValue;
            }
            return retValue;
        }
        private string GenerateSPDHarmMessages_Kittelson(List<clsRoadwaySubLink> RSLList, double TroupingEndMM, double BOQMMLocation, clsRoadway Roadway, string SpeedType, string Justification)
        {
            string retValue = string.Empty;

            double RecommendedSpeed = 0;
            double BeginMM = 0;
            double EndMM = 0;

            try
            {
                RSLList.Sort((l, r) => l.BeginMM.CompareTo(r.BeginMM));
                //Justification = "Normal";
                //if (BOQMMLocation > 0)
                //{
                //    Justification = "Queue";
                //}
                //else
                //{
                //    Justification = "Congestion";
                //}
                for (int i = 0; i < RSLList.Count; i++)
                {
                    if (SpeedType.ToLower() == "harmonized")
                    {
                        RecommendedSpeed = RSLList[i].HarmonizedSpeed;
                    }
                    else if (SpeedType.ToLower() == "recommended")
                    {
                        RecommendedSpeed = RSLList[i].RecommendedSpeed;
                    }
                    BeginMM = RSLList[i].BeginMM;
                    EndMM = RSLList[i].EndMM;

                    if (RSLList[i].BeginMM < Roadway.RecurringCongestionMMLocation)
                    {
                        retValue = InsertSPDHarmInfoIntoINFLODatabase(Roadway.Identifier.ToString(), DateTime.Now.AddMilliseconds(100), RecommendedSpeed, BeginMM, EndMM, Justification);
                        LogTxtMsg(txtINFLOLog, "Start: " + TroupingEndMM.ToString("00.00") +  "   Sublink: " + BeginMM.ToString("00.00") + " -to- " + EndMM.ToString("00.00") + "\t" + SpeedType + ": " + RecommendedSpeed.ToString("00") +  "\tJustification: " + Justification + "  -% Queued: " + RSLList[i].PercentQueuedCVs.ToString("00") +
                                               "\tCVAvgSpeed: " + RSLList[i].CVAvgSpeed.ToString("00") + "\tCVCount: " + RSLList[i].CVList.Count.ToString("00") + "\tQuedCVs: " + RSLList[i].NumberQueuedCVs.ToString("00") + "  \tTotalCVs: " + RSLList[i].TotalNumberCVs.ToString("00"));
                        if ((Justification.ToUpper() == "Normal".ToUpper()) && (RSLList[i].RecommendedSpeed < 30))
                        {
                            //LogTxtMsg(txtINFLOLog, "\tCVAvgSpeed: " + RSLList[i].CVAvgSpeed.ToString("00") + "\tCVCount: " + RSLList[i].CVList.Count.ToString("00") + "\tQuedCVs: " + RSLList[i].NumberQueuedCVs + "  \tTotalCVs: " + RSLList[i].TotalNumberCVs.ToString("00"));
                        }
                        if (retValue.Length > 0)
                        {
                            LogTxtMsg(txtCVDataLog, "Insert SPDHarmMsg: " + Roadway.Identifier + ", " + DateTime.Now.ToString() + ", " + RecommendedSpeed + ", " + BeginMM + ", " +
                                                    EndMM + ", " + Justification + ", 60" + "\r\n\t" + retValue);
                        }
                    }
                    else
                    {
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                retValue = "\tError in g SPDenerating SPD Harm message info." + "\r\n\t" + ex.Message;
                return retValue;
            }
            return retValue;
        }
        private string InsertSPDHarmInfoIntoINFLODatabase(string RoadwayId, DateTime CurrentTime, double RecommendedSpeed, double BeginMM, double EndMM, string Justification)
        {
            string retValue = string.Empty;
            string sqlStr = string.Empty;
            try
            {
                TimeSpan starttime = new TimeSpan(DateTime.Now.Ticks);
                DateTime DateGenerated = TimeZoneInfo.ConvertTimeToUtc(CurrentTime, TimeZoneInfo.Local);

                if (clsGlobalVars.DBInterfaceType.ToLower() == "sqlserver")
                {
                    sqlStr = "INSERT INTO TMEOutput_SPDHARMMessage_CV(RoadwayID, DateGenerated, RecommendedSpeed, BeginMM, EndMM, Justification, ValidityDuration) " +
                                "Values('" + RoadwayId + "', '" + DateGenerated + "', " + RecommendedSpeed + ", " + BeginMM.ToString("0.0") + ", " +
                                             EndMM.ToString("0.0") + ", '" + Justification + "', " + "60" + ")";
                    retValue = DB.InsertRow(sqlStr);
                    if (retValue.Length > 0)
                    {
                        return retValue;
                    }
                    LogTxtMsg(txtCVDataLog, "\t\tSPDHarm Message added to database: " + RoadwayId + ", " + DateGenerated.ToString() + ", " + RecommendedSpeed.ToString("0") + ", " + BeginMM.ToString("0.0") + ", " +
                                                                                    EndMM.ToString("0") + ", " + Justification  + ", " + "60");
                }
                else if (clsGlobalVars.DBInterfaceType.ToLower() == "oledb")
                {
                    sqlStr = "INSERT INTO TMEOutput_SPDHARMMessage_CV(RoadwayID, DateGenerated, RecommendedSpeed, BeginMM, EndMM, Justification, ValidityDuration) " +
                                "Values('" + RoadwayId + "', #" + DateGenerated + "#, " + RecommendedSpeed + ", " + BeginMM.ToString("0.0") + ", " +
                                             EndMM.ToString("0.0") + ", '" + Justification + "', " + "60" + ")";
                    retValue = DB.InsertRow(sqlStr);
                    if (retValue.Length > 0)
                    {
                        return retValue;
                    }
                    LogTxtMsg(txtCVDataLog, "\t\tSPDHarm Message added to database: " + RoadwayId + ", " + DateGenerated.ToString() + ", " + RecommendedSpeed.ToString("0") + ", " + BeginMM.ToString("0.0") + ", " +
                                                                                    EndMM.ToString("0") + ", " + Justification + ", " + "60");
                }
                TimeSpan endtime = new TimeSpan(DateTime.Now.Ticks);
                //LogTxtMsg(txtINFLOLog, "\t\tTime for adding SPDHarm Messages into database: " + (endtime.TotalMilliseconds - starttime.TotalMilliseconds).ToString("0") + " msecs");
            }
            catch (Exception ex)
            {
                retValue = "\tError in adding SPDHarm message info into INFLO database." + "\r\n\t" + ex.Message;
                return retValue;
            }
            return retValue;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            tmrCVData.Enabled = false;
        }
        private void button2_Click(object sender, EventArgs e)
        {
            tmrTSSData.Enabled = false;
        }

        private void frmINFLOApps_FormClosing(object sender, FormClosingEventArgs e)
        {
            tmrBOQ.Enabled = false;
            tmrCVData.Enabled = false;
            tmrTSSData.Enabled = false;

            if (CVDataProcessor != null)
            {
                CVDataProcessor.Close();
            }
            if (TSSDataProcessor != null)
            {
                TSSDataProcessor.Close();
            }
            if (QueueLog != null)
            {
                QueueLog.Close();
            }
            if (FillDataSetLog != null)
            {
                FillDataSetLog.Close();
            }
            if (SubLinKDataLog != null)
            {
                SubLinKDataLog.Close();
            }

            //workbook.Save();
            //workbook.Close(null, null, null);
            //excelApp.Quit();
        }

        private void btnUpdatePercentQueuedCVs_Click(object sender, EventArgs e)
        {
            int tmpPercent = 0;

            if ((txtSubLinkPercentQueuedCVs.Text.Trim()).Length > 0)
            {
                tmpPercent = int.Parse(txtSubLinkPercentQueuedCVs.Text.Trim());
                if ((tmpPercent > 0) && (tmpPercent <= 100))
                {
                    clsGlobalVars.SubLinkPercentQueuedCV = tmpPercent;
                }
            }
        }

        private void btnUpdatePercentQueuedCVs_Click_1(object sender, EventArgs e)
        {
            string tmpStrQueuedPercent = string.Empty;
            int tmpIntQueuedPercent = 0;
            tmpStrQueuedPercent = txtSubLinkPercentQueuedCVs.Text.Trim();
            if (tmpStrQueuedPercent.Length > 0)
            {
                try
                {
                    tmpIntQueuedPercent = int.Parse(tmpStrQueuedPercent);
                    if ((tmpIntQueuedPercent > 0) && (tmpIntQueuedPercent <= 100))
                    {
                        clsGlobalVars.SubLinkPercentQueuedCV = tmpIntQueuedPercent;
                    }
                    else
                    {
                        MessageBox.Show("Please check the value entered for sublink % queued CVs and reneter a value between 1 - 100");
                    }
                }
                catch (Exception exc)
                {
                    MessageBox.Show("Please check the value entered for sublink % queued CVs and reneter a value between 1 - 100 \r\n\t" + exc.Message);
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string tmpStrDSD = string.Empty;
            int tmpIntDSD = 0;
            tmpStrDSD = txtDSD.Text.Trim();
            if (tmpStrDSD.Length > 0)
            {
                try
                {
                    tmpIntDSD = int.Parse(tmpStrDSD);
                    if ((tmpIntDSD > 0) && (tmpIntDSD <= 300))
                    {
                        clsGlobalVars.DSD = tmpIntDSD;
                    }
                    else
                    {
                        MessageBox.Show("Please check the value entered for DSD and reneter a value between 1 - 300");
                    }
                }
                catch (Exception exc)
                {
                    MessageBox.Show("Please check the value entered for DSD and reneter a value between 1 - 300 \r\n\t" + exc.Message);
                }
            }
        }

        private void txtUpdateWeather_Click(object sender, EventArgs e)
        {
            string tmpStrBeginMM = string.Empty;
            double tmpDoubleBeginMM = 0;
            string tmpStrEndMM = string.Empty;
            double tmpDoubleEndMM = 0;
            string tmpStrVisibility = string.Empty;
            int tmpIntVisibility = 0;
            string tmpStrGripFactor = string.Empty;
            double tmpDoubleGripFactor = 0;

            tmpStrVisibility = txtVisibility.Text.Trim();
            tmpStrGripFactor = txtGripFactor.Text.Trim();
            tmpStrBeginMM = txtWRTMBeginMM.Text.Trim();
            tmpStrEndMM = txtWRTMEndMM.Text.Trim();
            
            txtRecommendedWRTMSpeed.Text = "";

            if (tmpStrBeginMM.Length > 0)
            {
                try
                {
                    tmpDoubleBeginMM = double.Parse(tmpStrBeginMM);
                    if (tmpDoubleBeginMM < 0)
                    {
                        MessageBox.Show("Please check the value entered for WRTM Begin MM and reneter a value > 0");
                        return;
                    }
                    else
                    {
                        clsGlobalVars.WRTMBeginMM = tmpDoubleBeginMM;
                    }
                }
                catch (Exception exc)
                {
                    MessageBox.Show("Please check the value entered for WRTM Begin MM and reneter a value > 0 \r\n\t" + exc.Message);
                    return;
                }
            }
            if (tmpStrEndMM.Length > 0)
            {
                try
                {
                    tmpDoubleEndMM = double.Parse(tmpStrEndMM);
                    if (tmpDoubleEndMM < 0)
                    {
                        MessageBox.Show("Please check the value entered for WRTM End MM and reneter a value > 0");
                        return;
                    }
                    else
                    {
                        clsGlobalVars.WRTMEndMM = tmpDoubleEndMM;
                    }
                }
                catch (Exception exc)
                {
                    MessageBox.Show("Please check the value entered for WRTM End MM and reneter a value > 0 \r\n\t" + exc.Message);
                    return;
                }
            }
            if (tmpStrVisibility.Length > 0)
            {
                try
                {
                    tmpIntVisibility = int.Parse(tmpStrVisibility);
                    if (tmpIntVisibility < 0)
                    {
                        MessageBox.Show("Please check the value entered for Visibility and reneter a value > 0");
                        return;
                    }
                }
                catch (Exception exc)
                {
                    MessageBox.Show("Please check the value entered for Visibility and reneter a value > 0 \r\n\t" + exc.Message);
                    return;
                }
            }
            if (tmpStrGripFactor.Length > 0)
            {
                try
                {
                    tmpDoubleGripFactor = double.Parse(tmpStrGripFactor);
                    if (tmpDoubleGripFactor < 0)
                    {
                        MessageBox.Show("Please check the value entered for Coefficient of Friction and reneter a value > 0");
                        return;
                    }
                }
                catch (Exception exc)
                {
                    MessageBox.Show("Please check the value entered for Coefficient of Friction and reneter a value > 0 \r\n\t" + exc.Message);
                    return;
                }
            }
            if (tmpIntVisibility >= 500)
            {
                if (tmpDoubleGripFactor >= 0.7)
                {
                    ApplyWRTMSpeed(ref RLList, clsGlobalVars.WRTMMaxRecommendedSpeed, clsGlobalVars.WRTMBeginMM, clsGlobalVars.WRTMEndMM);
                    txtRecommendedWRTMSpeed.Text = "WRTMMaxSpeed: " + clsGlobalVars.WRTMMaxRecommendedSpeed.ToString();
                }
                else if ((tmpDoubleGripFactor > 0.3) && (tmpDoubleGripFactor < 0.7))
                {
                    ApplyWRTMSpeed(ref RLList, clsGlobalVars.WRTMRecommendedSpeedLevel1, clsGlobalVars.WRTMBeginMM, clsGlobalVars.WRTMEndMM);
                    txtRecommendedWRTMSpeed.Text = "WRTMSpeedLevel1: " + clsGlobalVars.WRTMRecommendedSpeedLevel1.ToString();
                }
                else if (tmpDoubleGripFactor <= 0.3)
                {
                    ApplyWRTMSpeed(ref RLList, clsGlobalVars.WRTMRecommendedSpeedLevel2, clsGlobalVars.WRTMBeginMM, clsGlobalVars.WRTMEndMM);
                    txtRecommendedWRTMSpeed.Text = "WRTMSpeedLevel2: " + clsGlobalVars.WRTMRecommendedSpeedLevel2.ToString();
                }
            }
            else if ( tmpIntVisibility < 500)
            {
                if (tmpDoubleGripFactor >= 0.7)
                {
                    ApplyWRTMSpeed(ref RLList, clsGlobalVars.WRTMRecommendedSpeedLevel3, clsGlobalVars.WRTMBeginMM, clsGlobalVars.WRTMEndMM);
                    txtRecommendedWRTMSpeed.Text = "WRTMSpeedLevel3: " + clsGlobalVars.WRTMRecommendedSpeedLevel3.ToString();
                }
                else if ((tmpDoubleGripFactor > 0.3) && (tmpDoubleGripFactor < 0.7))
                {
                    ApplyWRTMSpeed(ref RLList, clsGlobalVars.WRTMRecommendedSpeedLevel4, clsGlobalVars.WRTMBeginMM, clsGlobalVars.WRTMEndMM);
                    txtRecommendedWRTMSpeed.Text = "WRTMSpeedLevel4: " + clsGlobalVars.WRTMRecommendedSpeedLevel4.ToString();
                }
                else if (tmpDoubleGripFactor <= 0.3)
                {
                    ApplyWRTMSpeed(ref RLList, clsGlobalVars.WRTMMinRecommendedSpeed, clsGlobalVars.WRTMBeginMM, clsGlobalVars.WRTMEndMM);
                    txtRecommendedWRTMSpeed.Text = "WRTMMinSpeed: " + clsGlobalVars.WRTMMinRecommendedSpeed.ToString();
                }
            }
        }

        //Added by Hassan Charara on 07/07 to make the INFLOApps program work with the Kittelson simulation
        private void tmrfile_Tick(object sender, EventArgs e)
        {
            string retValue = string.Empty;
            string SQLQuery = string.Empty;

            if (File.Exists(SyncFileName))
            {
                tmrfile.Enabled = false;
                LogTxtMsg(txtINFLOLog, "\r\n" + DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + ":" + DateTime.Now.Millisecond + 
                                       " -Sync File: " + SyncFileName + " was found.");
                LogTxtMsg(txtINFLOLog, "\t" + DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + ":" + DateTime.Now.Millisecond + 
                                       " -Disable Timer");

                StreamReader fReader = null;
                try
                {
                    fReader = File.OpenText(SyncFileName);

                    CVDataFlag = false;
                    TSSDataFlag = false;

                    string sLine;
                    while ((sLine = fReader.ReadLine()) != null)
                    {
                        if (sLine.Trim().Length == 0)
                            continue;

                        if (sLine.ToLower() == "tssdata")
                        {
                            TSSDataFlag = true;
                            LogTxtMsg(txtINFLOLog, "\t" + DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + ":" + DateTime.Now.Millisecond + 
                                                   " -TSS Data Enabled");
                        }
                        else if (sLine.ToLower() == "cvdata")
                        {
                            CVDataFlag = true;
                            LogTxtMsg(txtINFLOLog, "\t" + DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + ":" + DateTime.Now.Millisecond + 
                                                   " -CV Data Enabled");
                        }
                    }
                    fReader.Close();

                    //Delete the syncfile
                    LogTxtMsg(txtINFLOLog, "\t" + DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + ":" + DateTime.Now.Millisecond +
                                           " -Delete Data.txt file");
                    retValue = DeleteSyncFile(SyncFileName);
                    if (retValue.Length > 0)
                    {
                        MessageBox.Show(retValue);
                    }

                    if (CVDataFlag == true)
                    {
                        //Call the CVData function to retrieve the CV data from database and process it
                        LogTxtMsg(txtINFLOLog, "\t" + DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + ":" + DateTime.Now.Millisecond + 
                                               " -Process CV Data");
                        ProcessCVData();
                        //Delete the CV data from the database
                        SQLQuery = "Delete  from TME_CVData_Input";

                        retValue = string.Empty;
                        //Delete the CVData in the INFLO database
                        DataSet CVDataDataSet = new DataSet("CVData");
                        retValue = DB.FillDataSet(SQLQuery, ref CVDataDataSet);
                        if (retValue.Length > 0)
                        {
                            LogTxtMsg(txtINFLOLog, "\t" + DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + ":" + DateTime.Now.Millisecond + 
                                                   " -Error in deleting CVData records\r\n" + retValue); ;
                        }
                    }

                    if (TSSDataFlag == true)
                    {
                        //Call the TSSData function to retrieve the TSS data from database and process it
                        LogTxtMsg(txtINFLOLog, "\t" + DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + ":" + DateTime.Now.Millisecond + 
                                               " -Process TSS Data");
                        ProcessTSSData();

                        //Delete the TSS data from the database
                        SQLQuery = "Delete  from TME_TSSData_Input";

                        retValue = string.Empty;
                        //Delete the TSSData already in the INFLO database
                        DataSet TSSDataDataSet = new DataSet("TSSData");
                        retValue = DB.FillDataSet(SQLQuery, ref TSSDataDataSet);
                        if (retValue.Length > 0)
                        {
                            MessageBox.Show("Error in deleting TSSData records\r\n" + retValue);;
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogTxtMsg(txtINFLOLog, "\t" + DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + ":" + DateTime.Now.Millisecond + 
                                           " -Error in reading the sync file: " + SyncFileName + "\r\n" + ex.Message);
                }

            }
            else
            {
                //LogTxtMsg(txtINFLOLog, DateTime.Now + ":" + DateTime.Now.Millisecond.ToString("0000") + " -No Sync File: " + SyncFileName + " was found.");
            }
            if (tmrfile.Enabled == false)
            {
                LogTxtMsg(txtINFLOLog, "\t" + DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second + ":" + DateTime.Now.Millisecond + 
                                       " -Enable Timer");
                tmrfile.Enabled = true;
            }
        }

        private string DeleteSyncFile(string Filename)
        {
            string retValue = string.Empty;

            try
            {
                File.Delete(Filename);
            }
            catch (IOException ex)
            {
                retValue = "Error in deleting syncfile: " + Filename + "\r\n\t" + ex.Message;
            }
            return retValue;
        }

        private void txtSyncFileName_TextChanged(object sender, EventArgs e)
        {

        }

    }
}
