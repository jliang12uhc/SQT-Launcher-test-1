using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace SQT_Launcher
{
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

        public ExcelLauncher()
        {
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
            return RetrieveStringFromSql(LookupField_ProposalTemplateFilename);
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
            }
            catch (SqlException ex)
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
