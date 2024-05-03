using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Data.SqlClient;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Lieferanten_Dokumente_Ablegen {
    class InformationFromDataBase {

        const int WM_GETTEXT = 0x0D;

        [DllImport("user32.dll", EntryPoint = "FindWindow", SetLastError = true)]
        static extern IntPtr FindWindowByCaption(string ZeroOnly, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern int SendMessage(IntPtr hWnd, int msg, int Param, StringBuilder text);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        private static string GetCSGConnectionString() {
            IntPtr hWnd = FindWindow(null, "CSGClientConnection");
            if (hWnd != IntPtr.Zero) {
                IntPtr hEdit = FindWindowEx(hWnd, IntPtr.Zero, "ThunderRT6Frame", null);
                IntPtr ConString = FindWindowEx(hEdit, IntPtr.Zero, "ThunderRT6TextBox", null);
                StringBuilder connectionString = new StringBuilder(255);
                int RetVal2 = SendMessage(ConString, WM_GETTEXT, connectionString.Capacity, connectionString);
                IEnumerable<string> result = from Match match in Regex.Matches(connectionString.ToString(), "\"([^\"]*)\"")
                                             select match.ToString();

                List<string> infolist = result.ToList();
                string sqlConnectionString;
                if (infolist[4].Trim('"') == "Windows Security")
                    sqlConnectionString = "Data Source=" + infolist[9].Trim('"') +
                                               ";Initial Catalog=" + infolist[7].Trim('"') +
                                               ";Integrated Security = SSPI;";

                else
                    sqlConnectionString = "Data Source=" + infolist[9].Trim('"') +
                                               ";Initial Catalog=" + infolist[7].Trim('"') +
                                               ";User id=" + infolist[3].Trim('"') +
                                               ";Password=" + infolist[4].Trim('"') + ";";

                return sqlConnectionString;
            }
            return "CSGClientConnection_Notfound";
        }
        public static List<string> getCustomerPathFromDatabase(string customerEmailAddress, string documentType) {
            List<string> results = new List<string>();

            string connectionString = GetCSGConnectionString();
            //customerEmailAddress = "j.welling@csg-ms.de";
            if (connectionString == "CSGClientConnection_Notfound") {
                results.Add("CSGClientConnection_Notfound");
                return results;
            }

          


            string queryString = "SELECT bs_lfrt_texte.text " +
                          "FROM bs_lfrt_ansprech " +
                          "INNER JOIN bs_lfrt_texte ON bs_lfrt_ansprech.lfrt_key = bs_lfrt_texte.lfrt_key " +
                          $"WHERE bs_lfrt_ansprech.e_mail = '{customerEmailAddress}' " +
                          $"AND bs_lfrt_texte.typ = '{documentType}'";

            SqlDataReader myreader;

            try {
                using (SqlConnection connection = new SqlConnection(
                          connectionString)) {
                    SqlCommand command = new SqlCommand(queryString, connection);
                    command.Connection.Open();
                    myreader = command.ExecuteReader();

                    if (myreader.HasRows) {
                        foreach (DbDataRecord s in myreader) {
                            results.Add(s.GetString(0));
                        }
                        myreader.Close();
                        return results;
                    }
                    else {
                        myreader.Close();
                        results.Add("pathnotfound");
                        return results;

                    }
                }
            }
            catch (Exception ex) {
                results.Add("ConnectionError");
                return results;
            }
        }

        public static string getCustomerNameFromDatabase(string customerEmailAddress) {

            string connectionString = GetCSGConnectionString();

            if (connectionString == "CSGClientConnection_Notfound")
                return "CSGClientConnection_Notfound";

            string queryString = "select kurzbez from bs_lfrt_ansprech LEFT JOIN bs_lfrt ON bs_lfrt_ansprech.lfrt_key = bs_lfrt.lfrt_key where e_mail ='" + customerEmailAddress + "'";
            SqlDataReader myreader;

            try {
                using (SqlConnection connection = new SqlConnection(
                          connectionString)) {
                    SqlCommand command = new SqlCommand(queryString, connection);
                    command.Connection.Open();
                    myreader = command.ExecuteReader();

                    while (myreader.Read()) {

                        if (!string.IsNullOrEmpty(myreader[0].ToString()))
                            return myreader[0].ToString();
                        return "EmailNotFound";
                    }
                    myreader.Close();
                }
            }
            catch (Exception) {
                return "ConnectionError";
            }
            return "EmailNotFound";
        }

    }
}
