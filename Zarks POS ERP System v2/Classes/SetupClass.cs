using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;

namespace Zarks_POS_ERP_System_v2.Classes
{
    public class SetupClass : ConnectionClass
    {
        #region "DECLARATION"
        public string admcode { get; set; }
        public string shccode { get; set; }

        public string storecodemobes { get; set; }
        public string extractmobeh { get; set; }
        public string rmpathmobeh { get; set; }

        public string foodgtplu{ get; set; }
        public string bevgtplu { get; set; }
        public string beergtplu { get; set; }
        public string openitemgtplu { get; set; }
        public string deliveriesgtplu { get; set; }
        public string othersgtplu { get; set; }
        public string mktggcplu { get; set; }
        public string bfstime { get; set; }
        public string bfetime { get; set; }
        public string lstime { get; set; }
        public string letime { get; set; }
        public string dstime { get; set; }
        public string detime { get; set; }
        public string tosales { get; set; }
        public string creditcardsales { get; set; }
        public string firstcust { get; set; }
        public string beencust { get; set; }
        public string scdrc { get; set; }
        public string pwdrc { get; set; }
        public string zrrc { get; set; }
        public string scddisc { get; set; }
        public string lessvatdisc { get; set; }
        public string pwddisc { get; set; }
        public string zrdisc { get; set; }
        public string mktgdisc { get; set; }
        public string tenpcntdisc { get; set; }
        public string hndrdpcntdisc { get; set; }
        public string gcpayment { get; set; }
        public string gcashpayment { get; set; }
        public string grabpaypayment { get; set; }
        public string grabpayarpayment { get; set; }
        public string gcasharpayment { get; set; }
        public string wechatpayment { get; set; }
        public string wechatarpayment { get; set; }
        public string fparpayment { get; set; }
        public string llfarpayment { get; set; }
        public string gfarpayment { get; set; }
        public string zomarpayment { get; set; }
        public string SFarpayment { get; set; }
        public string MGCpayment { get; set; }
        public string hbrevc { get; set; }
        public string fprevc { get; set; }
        public string gfrevc { get; set; }
        public string llfrevc { get; set; }
        public string zomatorevc { get; set; }
        public string manganrevc { get; set; }
        public string complirevc { get; set; }
        public string drivebyrevc { get; set; }
        public string zarksmodgrpi { get; set; }
        public string drivebygrpi { get; set; }
        public string menuadjgrpi { get; set; }
        #endregion

        //for Login Form
        public void AuthUser()
        {
            using (OleDbConnection mdbcon = new OleDbConnection(mdbpath))
            {
                mdbcon.Open();
                cmd = new OleDbCommand("Select * from tblAuth", mdbcon);
                dbdr = cmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    admcode = dbdr.GetValue(1).ToString();
                    shccode = dbdr.GetValue(2).ToString();
                }
                mdbcon.Close();
            }
        }
        
        //for Admin Form
        public void AdminSetup()
        {
            using (OleDbConnection mdbcon = new OleDbConnection(mdbpath))
            {
                mdbcon.Open();
                cmd = new OleDbCommand("Select * from tblAdmin", mdbcon);
                dbdr = cmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    storecodemobes = dbdr.GetValue(1).ToString();
                    extractmobeh = dbdr.GetValue(2).ToString();
                    rmpathmobeh = dbdr.GetValue(3).ToString();
                }
                mdbcon.Close();
            }
        }
        
        //Group Types
        public void LoadMappings()
        {
            using (OleDbConnection mdbcon = new OleDbConnection(mdbpath))
            {
                mdbcon.Open();
                cmd = new OleDbCommand("Select * from tblMappings", mdbcon);
                dbdr = cmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    //grouptype
                    foodgtplu = dbdr.GetValue(1).ToString();
                    bevgtplu = dbdr.GetValue(2).ToString();
                    beergtplu = dbdr.GetValue(3).ToString();
                    openitemgtplu = dbdr.GetValue(4).ToString();
                    deliveriesgtplu = dbdr.GetValue(5).ToString();
                    othersgtplu = dbdr.GetValue(6).ToString();

                    bfstime = dbdr.GetValue(8).ToString();
                    bfetime = dbdr.GetValue(9).ToString();
                    lstime = dbdr.GetValue(10).ToString();
                    letime = dbdr.GetValue(11).ToString();
                    dstime = dbdr.GetValue(12).ToString();
                    detime = dbdr.GetValue(13).ToString();

                    //menu item plu
                    mktggcplu = dbdr.GetValue(7).ToString();
                    tosales = dbdr.GetValue(14).ToString();
                    firstcust = dbdr.GetValue(16).ToString();
                    beencust = dbdr.GetValue(17).ToString();

                    //discounts
                    scddisc = dbdr.GetValue(21).ToString();
                    pwddisc = dbdr.GetValue(22).ToString();
                    zrdisc = dbdr.GetValue(23).ToString();
                    mktgdisc = dbdr.GetValue(24).ToString();
                    tenpcntdisc = dbdr.GetValue(25).ToString();
                    lessvatdisc = dbdr.GetValue(26).ToString();
                    hndrdpcntdisc = dbdr.GetValue(40).ToString();

                    //payment
                    creditcardsales = dbdr.GetValue(15).ToString();
                    gcpayment = dbdr.GetValue(27).ToString();
                    gcashpayment = dbdr.GetValue(28).ToString();
                    grabpaypayment = dbdr.GetValue(41).ToString();
                    grabpayarpayment = dbdr.GetValue(42).ToString();
                    gcasharpayment = dbdr.GetValue(43).ToString();
                    wechatpayment = dbdr.GetValue(44).ToString();
                    wechatarpayment = dbdr.GetValue(45).ToString();
                    fparpayment = dbdr.GetValue(46).ToString();
                    llfarpayment = dbdr.GetValue(47).ToString();
                    gfarpayment = dbdr.GetValue(48).ToString();
                    zomarpayment = dbdr.GetValue(49).ToString();
                    SFarpayment = dbdr.GetValue(50).ToString();
                    MGCpayment = dbdr.GetValue(51).ToString();

                    //revcent
                    scdrc = dbdr.GetValue(18).ToString();
                    pwdrc = dbdr.GetValue(19).ToString();
                    zrrc = dbdr.GetValue(20).ToString();
                    hbrevc = dbdr.GetValue(29).ToString();
                    fprevc = dbdr.GetValue(30).ToString();
                    gfrevc = dbdr.GetValue(31).ToString();
                    llfrevc = dbdr.GetValue(32).ToString();
                    zomatorevc = dbdr.GetValue(33).ToString();
                    manganrevc = dbdr.GetValue(34).ToString();
                    complirevc = dbdr.GetValue(35).ToString();
                    drivebyrevc = dbdr.GetValue(36).ToString();

                    //menu group item
                    zarksmodgrpi = dbdr.GetValue(37).ToString();
                    drivebygrpi = dbdr.GetValue(38).ToString();
                    menuadjgrpi = dbdr.GetValue(39).ToString();
                }
                mdbcon.Close();
            }
        }

        //update mappings
        public void UpdateMappings(string str, OleDbType typ, int sze, string fld, string vlue)
        {
            OpenCon();
            cmd = new OleDbCommand("UPDATE tblMappings SET " + fld + "=" + str, con);
            cmd.Parameters.Add(str, typ, sze, fld).Value = vlue;
            cmd.ExecuteNonQuery();
            closeConn();
        }

        //update admin passcode
        public void UpdateAdminCode(string str, OleDbType typ, int sze, string fld, string vlue)
        {
            OpenCon();
            cmd = new OleDbCommand("UPDATE tblAuth SET " + fld + "=" + str, con);
            cmd.Parameters.Add(str, typ, sze, fld).Value = vlue;
            cmd.ExecuteNonQuery();
            closeConn();
        }

        //update admin account
        public void UpdateAdminAcct(string str, OleDbType typ, int sze, string fld, string vlue)
        {
            OpenCon();
            cmd = new OleDbCommand("UPDATE tblAdmin SET " + fld + "=" + str, con);
            cmd.Parameters.Add(str, typ, sze, fld).Value = vlue;
            cmd.ExecuteNonQuery();
            closeConn();
        }
    }
}
