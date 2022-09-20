using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;

namespace Zarks_POS_ERP_System_v2.Classes
{
    public class ConnectionClass
    {
        #region

        public OleDbConnection con;
        public OleDbCommand cmd;
        public OleDbDataReader dbdr;
        public OleDbDataReader dbdr1;

        public string rmpath { get; set; }
        public string mdbpath;

        #endregion

        public ConnectionClass()
        {
            mdbpath = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;" +
                @"Data Source=|DataDirectory|\Preferences\ZarksPOSERPSystemDB.mdb;Jet OLEDB:Database Password=cpos6163";
        }

        public void OpenCon()
        {
            con = new OleDbConnection(mdbpath);
            if (con.State == ConnectionState.Closed)
            {
                con = new OleDbConnection(mdbpath);
                con.Open();
            }
        }

        public void closeConn()
        {
            con.Close();
            con.Dispose();
        }
    }
}
