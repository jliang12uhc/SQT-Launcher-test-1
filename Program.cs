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
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace SQT_Launcher
{
    internal class Program
    {
        const int versionNumber = 1;
        const string shortcutName = "Sales Quote Tool";

        static void Main(string[] args)
        {
            string user = Environment.UserName;

            // updating this exe - just send out an email/link to the app manifest on the shared drive
            //Console.WriteLine("Welcome to the Sales Quote Tool (SQT). Please wait while the file is loaded.\n");

            //ProgressUpdate("Updating local SQT directory...", 1);
            ExcelLauncher launcher = new ExcelLauncher();
            string localDir = @"C:\Users\" + user + @"\Documents\SQT\";
            Directory.CreateDirectory(localDir);
            //Console.WriteLine(" Done.");

            //ProgressUpdate("Retrieving SQT network parameters...", 1);
            string sqtFilename = launcher.LookupSqtFilename();
            string localPathSqt = localDir + sqtFilename;
            string sqtUrl = launcher.LookupSqtUrl();

            string proposalTemplateFilename = launcher.LookupProposalTemplateFilename();
            string localPathProposalTemplate = localDir + proposalTemplateFilename;
            string proposalTemplateUrl = launcher.LookupProposalTemplateUrl();
            //Console.WriteLine(" Done.");

            //ProgressUpdate("Retrieving the latest version of SQT...", 1);
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            using (WebClient wc = new WebClient())
            {
                wc.DownloadFile(sqtUrl, localPathSqt);
                wc.DownloadFile(proposalTemplateUrl, localPathProposalTemplate);
            }
            //Console.WriteLine(" Done.");

            //ProgressUpdate("Launching Excel...", 1);
            Workbook sqt = excel.Workbooks.Open(localPathSqt);
            sqt.BeforeSave += new WorkbookEvents_BeforeSaveEventHandler(ThisWorkbook_BeforeSave);
            excel.DisplayAlerts = false;
            excel.Visible = true;
            //Console.WriteLine(" Done.");

            //Console.WriteLine();
            //Console.WriteLine("SQT is currently in use by " + user + ". Please do not close this window.");
            //Console.WriteLine();

            FileInfo f = new FileInfo(localPathSqt);
            while (IsFileLocked(f))
            {
                Thread.Sleep(1000);
            }
            ////ProgressUpdate("Cleaning up local files...", 1);
            //DeleteFile(localPathSqt);
            //DeleteFile(localPathProposalTemplate);

            //string fullPath = Process.GetCurrentProcess().MainModule.FileName;
            string deskDir = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            bool shortcutExists = File.Exists(deskDir + @"\" + shortcutName + ".url");
            AddDesktopShortcut(shortcutName);
            if (!shortcutExists)
            {
                MessageBox.Show("Installed a shortcut to the Sales Quote Tool (SQT) on your desktop.", "Sales Quote Tool (SQT)");
            }
            //Console.WriteLine(" Done.");
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
                    ////Console.WriteLine("Deleted " + path);
                }
                else
                {
                    ////Console.WriteLine(path + " does not exist.");
                }
            }
            catch (Exception e)
            {
                ////Console.WriteLine(e);
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
            ////Console.WriteLine("\b");
            //}
            ////Console.WriteLine(" Done.");
        }
        static void AddDesktopShortcut(string linkName)
        {
            string deskDir = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);

            using (StreamWriter writer = new StreamWriter(deskDir + "\\" + linkName + ".url"))
            {
                string app = System.Reflection.Assembly.GetExecutingAssembly().Location;
                writer.WriteLine("[InternetShortcut]");
                writer.WriteLine("URL=file:///" + app);
                writer.WriteLine("IconIndex=0");
                string icon = app.Replace('\\', '/');
                writer.WriteLine("IconFile=" + icon);
            }
        }
    }
}