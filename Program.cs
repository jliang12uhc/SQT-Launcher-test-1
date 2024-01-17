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
using System.Threading;

namespace SQT_Launcher
{
    internal class Program
    {
        //static async Task Main(string[] args)
        const int versionNumber = 1;
        const string sqtFilename = "SQT.xlsm";

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

            string user = Environment.UserName;
            string installerFilepath = @"C:\Users\" + user + @"\Documents\2023\UMR (local)\Sales Tool\Installer\installer dev 1.vbs"; //.Replace(@"\\", @"\");
            if (File.Exists(installerFilepath))
            {
                Console.WriteLine(installerFilepath + " does exist");
            }
            else
            {
                Console.WriteLine("doesn't exist");
            }

            //string processArgs = @"cscript " + @"//B " + @" //Nologo " + installerFilepath;
            //System.Diagnostics.Process.Start(@processArgs);

            Process installerProcess = new Process();
            installerProcess.StartInfo.FileName = @installerFilepath;
            installerProcess.Start();
            installerProcess.WaitForExit();
            installerProcess.Close();

            ExcelLauncher launcher = new ExcelLauncher();

            string localDir = @"C:\Users\" + user + @"\Documents\2023\SQT\";
            Directory.CreateDirectory(localDir);

            string localPath = localDir + sqtFilename;
            string sqtUrl = launcher.LookupSqtUrl();

            Application excel = new Application();

            using (WebClient wc = new WebClient())
            {
                wc.DownloadFile(sqtUrl, localPath);
            }

            Workbook sqt = excel.Workbooks.Open(localPath);
            sqt.BeforeSave += new WorkbookEvents_BeforeSaveEventHandler(ThisWorkbook_BeforeSave);
            excel.DisplayAlerts = false;
            excel.Visible = true;

            FileInfo f = new FileInfo(localPath);
            while (IsFileLocked(f))
            {
                Thread.Sleep(1000);
            }
            DeleteFile(localPath);
        }

        static void ThisWorkbook_BeforeSave(bool SaveAsUI, ref bool Cancel)
        {
            Cancel = true;
        }
        static void DeleteFile(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                    //Console.WriteLine("Deleted " + path);
                }
                else
                {
                    //Console.WriteLine(path + " does not exist.");
                }
            }
            catch (Exception e)
            {
                //Console.WriteLine(e);
            }
        }
        static bool IsFileLocked(FileInfo file)
        {
            FileStream stream = null;

            try
            {
                stream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            }
            catch (IOException)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }

            //file is not locked
            return false;
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