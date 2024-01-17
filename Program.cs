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
using System.Threading;
using Microsoft.Office.Interop.Excel;

namespace SQT_Launcher
{
    internal class Program
    {
        //static async Task Main(string[] args)
        const int versionNumber = 1;
        
        static void Main(string[] args)
        {
            string user = Environment.UserName;

            Process p = new Process();
            p.StartInfo.Arguments = @"""C:\Users\jliang12\Documents\2023\UMR (local)\Sales Tool\Installer\installer dev 1.vbs""";
            p.StartInfo.FileName = @"C:\Windows\System32\cscript.exe";
            p.StartInfo.Arguments = versionNumber.ToString();   

            p.Start();
            p.WaitForExit();
            p.Close();

            ExcelLauncher launcher = new ExcelLauncher();

            string localDir = @"C:\Users\" + user + @"\Documents\2023\SQT\";
            Directory.CreateDirectory(localDir);

            string sqtFilename = launcher.LookupSqtFilename();
            string localPathSqt = localDir + sqtFilename;
            string sqtUrl = launcher.LookupSqtUrl();

            string proposalTemplateFilename = launcher.LookupProposalTemplateFilename();
            string localPathProposalTemplate = localDir + proposalTemplateFilename;
            string proposalTemplateUrl = launcher.LookupProposalTemplateUrl();

            Application excel = new Application();

            using (WebClient wc = new WebClient())
            {
                wc.DownloadFile(sqtUrl, localPathSqt);
            }

            Workbook sqt = excel.Workbooks.Open(localPathSqt);
            sqt.BeforeSave += new WorkbookEvents_BeforeSaveEventHandler(ThisWorkbook_BeforeSave);
            excel.DisplayAlerts = false;
            excel.Visible = true;

            FileInfo f = new FileInfo(localPathSqt);
            while (IsFileLocked(f))
            {
                Thread.Sleep(1000);
            }
            DeleteFile(localPathSqt);
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
        const string LookupField_SqtFilename = "SQT_FILENAME";
        const string LookupField_SqtVersion = "SQT_VERSION";
        const string LookupField_ProposalTemplateUrl = "PROPOSAL_TEMPLATE_URL";
        const string LookupField_ProposalTemplateFilename = "PROPOSAL_TEMPLATE_FILENAME";
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
            return RetrieveStringFromSql(LookupField_SqtVersion); 
        }
        public string LookupSqtUrl()
        {
            return RetrieveStringFromSql(LookupField_SqtUrl);
        }
        public string LookupSqtFilename()
        {
            return RetrieveStringFromSql(LookupField_SqtFilename);
        }
        public string LookupProposalTemplateUrl()
        {
            return RetrieveStringFromSql(LookupField_ProposalTemplateUrl);
        }
        public string LookupProposalTemplateFilename()
        {
            return RetrieveStringFromSql(LookupField_ProposalTemplateUrl);
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