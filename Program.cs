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
    internal class Programs
    {
        const int versionNumber = 1;
        const string shortcutName = "Sales Quote Tool";

        static void Main(string[] args)
        {
            string user = Environment.UserName;

            // updating this exe - just send out an email/link to the app manifest on the shared drive
            ExcelLauncher launcher = new ExcelLauncher();
            string localDir = @"C:\Users\" + user + @"\Documents\SQT\";
            Directory.CreateDirectory(localDir);

            string sqtFilename = launcher.LookupSqtFilename();
            string localPathSqt = localDir + sqtFilename;
            string sqtUrl = launcher.LookupSqtUrl();

            string proposalTemplateFilename = launcher.LookupProposalTemplateFilename();
            string localPathProposalTemplate = localDir + proposalTemplateFilename;
            string proposalTemplateUrl = launcher.LookupProposalTemplateUrl();
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            using (WebClient wc = new WebClient())
            {
                wc.DownloadFile(sqtUrl, localPathSqt);
                wc.DownloadFile(proposalTemplateUrl, localPathProposalTemplate);
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
            string deskDir = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string shortcutPath = deskDir + @"\" + shortcutName + ".url";
            bool shortcutExists = File.Exists(shortcutPath);
            AddDesktopShortcut(shortcutName);
            /*
            if (!shortcutExists)
            {
                DialogResult result = MessageBox.Show("Would you like to install a shortcut to the Sales Quote Tool (SQT) on your desktop?", "Sales Quote Tool (SQT)", MessageBoxButtons.YesNo);
                if (result == System.Windows.Forms.DialogResult.No)
                {
                    File.Delete(shortcutPath);
                }
                else
                {
                    MessageBox.Show("Successfully created a desktop shortcut to SQT.", "Sales Quote Tool (SQT)");
                }
            }
            */
        }

        static void ThisWorkbook_BeforeSave(bool SaveAsUI, ref bool Cancel)
        {
            Cancel = true;
        }
        static void ThisWorkbook_BeforeClose(ref bool Cancel)
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
                }
                else
                {
                }
            }
            catch (Exception e)
            {
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