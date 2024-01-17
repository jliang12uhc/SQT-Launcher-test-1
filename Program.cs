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
        const int versionNumber = 1;

        static void Main(string[] args)
        {
            string user = Environment.UserName;

            /*
            // LAUNCH INSTALLER
            Process p = new Process();
            p.StartInfo.FileName = @"""C:\Users\jliang12\Documents\2023\UMR (local)\Sales Tool\Installer\installer dev 1.vbs""";
            p.StartInfo.Arguments = versionNumber.ToString();   

            p.Start();
            p.WaitForExit();
            p.Close();
            */

            Console.WriteLine("Welcome to the Sales Quote Tool (SQT). Please wait while the file is loaded.\n");

            ProgressUpdate("Updating local SQT directory...", 1);
            ExcelLauncher launcher = new ExcelLauncher();
            string localDir = @"C:\Users\" + user + @"\Documents\2023\SQT\";
            Directory.CreateDirectory(localDir);
            Console.WriteLine(" Done.");

            ProgressUpdate("Retrieving SQT network parameters...", 1);
            string sqtFilename = launcher.LookupSqtFilename();
            string localPathSqt = localDir + sqtFilename;
            string sqtUrl = launcher.LookupSqtUrl();

            string proposalTemplateFilename = launcher.LookupProposalTemplateFilename();
            string localPathProposalTemplate = localDir + proposalTemplateFilename;
            string proposalTemplateUrl = launcher.LookupProposalTemplateUrl();
            Console.WriteLine(" Done.");

            ProgressUpdate("Retrieving the latest version of SQT...", 1);
            Application excel = new Application();

            using (WebClient wc = new WebClient())
            {
                wc.DownloadFile(sqtUrl, localPathSqt);
                wc.DownloadFile(proposalTemplateUrl, localPathProposalTemplate);
            }
            Console.WriteLine(" Done.");

            ProgressUpdate("Launching Excel...", 1);
            Workbook sqt = excel.Workbooks.Open(localPathSqt);
            sqt.BeforeSave += new WorkbookEvents_BeforeSaveEventHandler(ThisWorkbook_BeforeSave);
            excel.DisplayAlerts = false;
            excel.Visible = true;
            Console.WriteLine(" Done.");

            Console.WriteLine();
            Console.WriteLine("SQT is currently in use by " + user + ". Please do not close this window.");
            Console.WriteLine();

            FileInfo f = new FileInfo(localPathSqt);
            while (IsFileLocked(f))
            {
                Thread.Sleep(1000);
            }
            ProgressUpdate("Cleaning up local files...", 1);
            DeleteFile(localPathSqt);
            Console.WriteLine(" Done.");
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
        static void ProgressUpdate(string action, int threadTimer)
        {
            Console.Write(action);
            using (var pb = new ProgressBar())
            {
                //var pb = new ProgressBar();
                for (int i = 0; i <= 100; i++)
                {
                    pb.Report((double)i / 100);
                    Thread.Sleep(threadTimer);
                }
            }
            //Console.WriteLine("\b");
            //}
            //Console.WriteLine(" Done.");
        }
    }

}