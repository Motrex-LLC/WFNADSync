using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.Net.Mail;
using System.IO;

namespace WFNADSync
{
    public sealed class Utility
    {
        private static System.Configuration.AppSettingsReader appSettingRdr = new System.Configuration.AppSettingsReader();

        //public static SqlConnection getWFNDBConn()
        //{
        //    SqlConnection Conn = new SqlConnection(appSettingRdr.GetValue("WFNConnStr", typeof(string)).ToString());
        //    Conn.Open();
        //    return Conn;
        //}

        private static string getMailHost()
        {
            //return System.Configuration.ConfigurationManager.AppSettings.Get("MailHost");
            return appSettingRdr.GetValue("MailHost", typeof(string)).ToString();
        }

        private static string getEnvironment()
        {
            //return System.Configuration.ConfigurationManager.AppSettings.Get("Environment");
            return appSettingRdr.GetValue("Environment", typeof(string)).ToString();
        }

        public static string getAlertTo()
        {
            //return System.Configuration.ConfigurationManager.AppSettings.Get("AlertEmail");
            return appSettingRdr.GetValue("AlertEmail", typeof(string)).ToString();
        }

        public static string getDumpADInfoToDB()
        {
            //return System.Configuration.ConfigurationManager.AppSettings.Get("AlertEmail");
            return appSettingRdr.GetValue("DumpADInfoToDB", typeof(string)).ToString();
        }        
        public static void SendSMTPEmail(string strAddressTo, string strSubject, string strBody, string strAddressCC = "", HashSet<string> hashAttachment = null)
        {
            try
            {
                MailMessage message = new MailMessage();
                message.IsBodyHtml = true;
                message.Priority = MailPriority.High;
                message.From = new MailAddress("webadmin@motrexllc.com", "MDM Admin");
                message.To.Add(new MailAddress(strAddressTo));
                message.Subject = strSubject + " [" + getEnvironment() + "]";
                message.Body = strBody;
                if (strAddressCC != "")
                {
                    message.CC.Add(new MailAddress(strAddressCC));
                }
                if (hashAttachment != null)
                {
                    foreach (string strAttachment in hashAttachment)
                    {
                        if (File.Exists(strAttachment))
                        {
                            message.Attachments.Add(new Attachment(strAttachment));
                        }
                    }
                }

                //message.Attachments.Add(new Attachment("D:\\WebApps\\MDM\\bin\\AppOutput\\ConfigIDLiftTruck_ExportAll_20200805231037.xlsx", MediaTypeNames.Application.Octet));            
                SmtpClient mclient = new SmtpClient();
                mclient.Host = getMailHost();
                //mclient.Port = System.Configuration.ConfigurationManager.AppSettings.Get("MailHost");

                mclient.Send(message);
            }
            catch (Exception e) { }
            //catch (Exception e) { SendSMTPEmail(getAlertTo(), "Exception in SendSMTPEmail", e.ToString(), ""); }
        }

        public static string GetAppBinDir()
        {
            //CodeBase give path in URI format
            string codeBase = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;
            UriBuilder uri = new UriBuilder(codeBase);
            string strBinDir = Path.GetDirectoryName(Uri.UnescapeDataString(uri.Path));

            //if (strBinDir.Substring(strBinDir.Length - 1) != Path.DirectorySeparatorChar.ToString())
            //{
            //    strBinDir = strBinDir + Path.DirectorySeparatorChar.ToString();
            //}

            if (strBinDir.Substring(strBinDir.Length - 1) == Path.DirectorySeparatorChar.ToString())
            {
                strBinDir = strBinDir.Substring(0, strBinDir.Length - 1);
            }

            return strBinDir;
        }
        public static void AppLog3Pass(string strPassin1, string strPassin2, string strPassin3, string FileName)
        {
            try
            {
                StreamWriter swDumper;
                string strDumpDirectory = GetAppBinDir() + Path.DirectorySeparatorChar.ToString() + "app_log" + Path.DirectorySeparatorChar.ToString();
                if (!Directory.Exists(strDumpDirectory))
                {
                    Directory.CreateDirectory(strDumpDirectory);
                }
                swDumper = File.AppendText(strDumpDirectory + FileName);
                //swDumper.WriteLine(strPassin1 + "\t\t" + strPassin2 + "\t\t" + strPassin3);
                swDumper.WriteLine(strPassin1 + "," + strPassin2 + "," + strPassin3);
                swDumper.Flush();
                swDumper.Close();
            }

            catch (Exception e)
            {
                //SendSMTPEmail(getAlertTo(), "Exception: Dumper3PassFilePath", e.ToString(), "", "");
            }
        }

        public static void AppLog3PassReplace(string strPassin1, string strPassin2, string strPassin3, string FileName)
        {
            try
            {
                StreamWriter swDumper;
                string strDumpDirectory = GetAppBinDir() + Path.DirectorySeparatorChar.ToString() + "app_log" + Path.DirectorySeparatorChar.ToString();
                if (!Directory.Exists(strDumpDirectory))
                {
                    Directory.CreateDirectory(strDumpDirectory);
                }
                swDumper = File.CreateText(strDumpDirectory + FileName);
                swDumper.WriteLine(strPassin1 + "\t\t" + strPassin2 + "\t\t" + strPassin3);
                swDumper.Flush();
                swDumper.Close();
            }

            catch (Exception e)
            {
                //SendSMTPEmail(getAlertTo(), "Exception: Dumper3PassFilePath", e.ToString(), "", "");
            }
        }

        public static string GetSessionIDFromSQL()
        {
            String strSQL = "";
            SqlCommand Cmd = new SqlCommand("", AppDat.getWFNDBConn());
            SqlDataReader Rdr;
            string strResult = "1/1/1900";

            try
            {
                strSQL = "Select Convert(varChar(25),GetDate(),121) As SessionID";
                Cmd.CommandText = strSQL;
                Rdr = Cmd.ExecuteReader();
                if (Rdr.HasRows)
                {
                    Rdr.Read();
                    strResult = (Rdr["SessionID"] + "").Trim().ToString();
                }
                Rdr.Close();
                return strResult;
            }

            catch (Exception ex)
            {
                return "1/1/1900";
            }

            finally
            {
                Cmd.Connection.Close();
                Cmd = null;
                Rdr = null;
            }
        }

    }
}
