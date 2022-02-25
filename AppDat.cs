using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;

namespace WFNADSync
{
    class AppDat
    {
        private static System.Configuration.AppSettingsReader appSettingRdr = new System.Configuration.AppSettingsReader();

        public static SqlConnection getWFNDBConn()
        {
            SqlConnection Conn = new SqlConnection(appSettingRdr.GetValue("WFNConnStr", typeof(string)).ToString());
            Conn.Open();
            return Conn;
        }

    }
}
