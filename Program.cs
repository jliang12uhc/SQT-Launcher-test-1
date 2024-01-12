using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Security.Cryptography;
using System.Diagnostics;
using System.Net;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace SQT_Launcher
{
    internal class Program
    {
        //static async Task Main(string[] args)
        const int versionNumber = 1;

        static void Main(string[] args)
        {
            // decide on a common directory eg. (C:\Users\[MSID]\Documents\something...)

            // THE INITIAL INSTALL
            // Running Installer 1.0 will result in a copy of itself (from the shared drive) and SQT-Launcher to be put in a common directory.

            // Any subsequent runs of SQT-Launcher should do the following:

            // UPON RUNNING SQT-Launcher:
            // This program should first be installed with an INSTALLER (another program which also acts as the updater). 
            // When this program starts, check against SQL if an updated version of the INSTALLER exists.
                // If so, update the INSTALLER.
            // Then, launch the INSTALLER, passing the name, PROCESS ID, and filepath (System.Reflection.Assembly.GetEntryAssembly().Location;) of this program (SQT Launcher) to it.
            // The INSTALLER should also make a read to SQL for an updated version of SQT LAUNCHER.
                // If there is, close THIS process (SQT LAUNCHER) and update it from the INSTALLER
                // Then relaunch the new SQT LAUNCHER with the same filepath
            // Close the INSTALLER

            // If the INSTALLER is run standalone:
                // Check for SQT-Launcher updates and replace accordingly

            ExcelLauncher launcher = new ExcelLauncher();
            //string sqtUrl = @"\\nas01773pn\UMR Pricing\a UMR Model Team\UMR\Sales Quote Tool\Dev\Sales Quote Tool - Dev 31.xlsm";
            //string sqtUrl = launcher.LookupSqtUrl();
            //string localTempPath = Path.GetTempFileName();
            
            // replace with get local path
            string localPath = @"C:\Users\jliang12\Documents\2023\SQT.xlsm";
            string sqtUrl = launcher.LookupSqtUrl();
            //launcher.OpenSqt(sqtUrl, localTempPath);

            Application excel = new Application();
            //Workbook sqt = excel.Workbooks.Open(localTempPath);

            using (WebClient wc = new WebClient())
            {
                wc.DownloadFile(sqtUrl, localPath);
            }

            Workbook sqt = excel.Workbooks.Open(localPath);
            sqt.BeforeSave += new WorkbookEvents_BeforeSaveEventHandler(ThisWorkbook_BeforeSave);
            excel.DisplayAlerts = false;
            excel.Visible = true;

            //FileInfo fi = new FileInfo(localTempPath);
            //System.Diagnostics.Process.Start(localTempPath);

            Console.WriteLine("Done");
            Console.ReadLine();
        }

        static void ThisWorkbook_BeforeSave(bool SaveAsUI, ref bool Cancel)
        {
            Cancel = true;
        }
    }

    internal class ExcelLauncher
    {
        const string DatabaseName = "WN000053767";
        const string InitialCatalog = "UMR";
        const string InitialSchema = "SALES";
        const string InitialTable = "SQT_PARAMETERS";
        const string LookupField_SqtUrl = "SQT_URL";
        const string LookupField_SqtVersion = "SQT_VERSION";
        const int ConnectTimeout = 300;

        public ExcelLauncher() {
            InitSqlConnection();
            LookupSqtVersion();
            LookupSqtUrl();
        }
        private SqlConnection Sc;
        public void InitSqlConnection()
        {
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
            builder.DataSource = DatabaseName;
            builder.InitialCatalog = InitialCatalog;
            builder.IntegratedSecurity = true;  // true == use Windows Authentication
            builder.ConnectTimeout = ConnectTimeout;    // time limit on queries 
            builder.MultipleActiveResultSets = true;
            Sc = new SqlConnection(builder.ConnectionString);
            Sc.Open();
        }

        public void OpenSqt(string networkPath, string localTempPath)
        {
            
            var uri = new Uri(networkPath);
            var fName = Path.GetFullPath(localTempPath);
            using (WebClient wc = new WebClient())
            {
                wc.DownloadFile(uri, fName);
            }
        }
    public string LookupSqtVersion() 
        {
            // Look up from sql
            string retrievedString = RetrieveStringFromSql(LookupField_SqtVersion);
            return retrievedString; 
        }
        public string LookupSqtUrl()
        {
            string retrievedString = RetrieveStringFromSql(LookupField_SqtUrl); //null;
            //SqtUrl = await RetrieveStringFromSql(LookupField_SqtUrl);
            return retrievedString;
        }        
        private string RetrieveStringFromSql(string field)
        {
            string retrievedString = null;
            string query = "SELECT " + field + " FROM " + InitialSchema + "." + InitialTable + " ORDER BY ID DESC";
            StringBuilder errorMessages = new StringBuilder();
            try
            {
                if (Sc.State == System.Data.ConnectionState.Closed)
                {
                    Sc.Open();
                }
                using (SqlCommand command = new SqlCommand(query, Sc))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            if (!reader.IsDBNull(0)) { retrievedString = reader.GetString(0); break; }
                        }
                    }

                }
            } catch (SqlException ex)
            {
                for (int i = 0; i < ex.Errors.Count; i++)
                {
                    errorMessages.Append("Index #" + i + "\n" +
                        "Message: " + ex.Errors[i].Message + "\n" +
                        "LineNumber: " + ex.Errors[i].LineNumber + "\n" +
                        "Source: " + ex.Errors[i].Source + "\n" +
                        "Procedure: " + ex.Errors[i].Procedure + "\n");
                }
                Console.WriteLine(errorMessages.ToString());
            }
            //Console.WriteLine(field + ": " + retrievedString);
            return retrievedString;
        }

    }
}