using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Zarks_POS_ERP_System_v2.Classes;

namespace Zarks_POS_ERP_System_v2
{
    public partial class GenerationForm : Form
    {
        public string rmconnect;
        public OleDbCommand cmdsales;
        public OleDbCommand cmdregdisc;
        public OleDbCommand cmdsrdisc;
        public OleDbCommand cmdpwddisc;
        public OleDbCommand cmdcomp;
        public OleDbCommand cmdhndrddisc;
        public OleDbCommand cmdgc;
        public OleDbCommand cmddsr;
        SetupClass setupcls;

        public GenerationForm()
        {
            InitializeComponent();
            setupcls = new SetupClass();
            setupcls.AdminSetup();
            rmconnect = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + setupcls.rmpathmobeh +
                       ";Extended Properties=dBASE IV;User ID=Admin;Password=;";
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            btnGenerate.Enabled = false;
            DateTime sfrom, efrom;
            sfrom = dtpfrom.Value;
            efrom = dtpto.Value;
            if (sfrom <= efrom)
            {
                progressBar1.Maximum = Convert.ToInt32((efrom - sfrom).TotalDays + 1);
                do
                {
                    Application.DoEvents();
                    lblstats.Text = sfrom.ToString("MM/dd/yyyy");
                    CreateExcelFile cdata = new CreateExcelFile();
                    cdata.datestart = sfrom.ToShortDateString();
                    cdata.dateend = efrom.ToShortDateString();
                    cdata.GetSession();
                    progressBar1.Value = progressBar1.Value + 1;
                    sfrom = sfrom.AddDays(1);

                    ConnectionClass cc = new ConnectionClass();
                    cc.OpenCon();
                    /*
                    cmdsales = new OleDbCommand("Delete from tblSales", cc.con);
                    cmdsales.ExecuteNonQuery();
                    cmdregdisc = new OleDbCommand("Delete from tblRegDiscSales", cc.con);
                    cmdregdisc.ExecuteNonQuery();
                    cmdsrdisc = new OleDbCommand("Delete from tblSrDiscSales", cc.con);
                    cmdsrdisc.ExecuteNonQuery();
                    cmdpwddisc = new OleDbCommand("Delete from tblPwdDiscSales", cc.con);
                    cmdpwddisc.ExecuteNonQuery();
                    cmdcomp = new OleDbCommand("Delete from tblComplimentary", cc.con);
                    cmdcomp.ExecuteNonQuery();
                    cmdhndrddisc = new OleDbCommand("Delete from tblHunDisc", cc.con);
                    cmdhndrddisc.ExecuteNonQuery();
                    cmdgc = new OleDbCommand("Delete from tblGC", cc.con);
                    cmdgc.ExecuteNonQuery();
                    */
                    cmddsr = new OleDbCommand("Delete from tblDSR", cc.con);
                    cmddsr.ExecuteNonQuery();
                    cc.closeConn();
                }
                while (sfrom <= efrom);
                MessageBox.Show(this, "Process Complete!");
                progressBar1.Maximum = 0;
                progressBar1.Value = 0;
                lblstats.Text = "00/00/0000";
                btnGenerate.Enabled = true;
            }
        }

        public void GetSessionStart()
        {
            using (OleDbConnection rmconn = new OleDbConnection(rmconnect))
            {
                rmconn.Open();
                OleDbCommand oleCmd = new OleDbCommand("SELECT SESSION_NO, first_bill, last_bill FROM REP" + dtpfrom.Value.ToString("yy") +
                                                       " WHERE date_start = #" + dtpfrom.Value.ToShortDateString() + "# " +
                                                       "AND first_bill > 0 and last_bill > 0", rmconn);
                OleDbDataReader dbdr = oleCmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    lblsesfrom.Text = dbdr.GetValue(0).ToString();
                }
                else
                {
                    lblsesfrom.Text = "00";
                }
                rmconn.Close();
            }
        }

        public void GetSessionEnd()
        {
            using (OleDbConnection rmconn = new OleDbConnection(rmconnect))
            {
                rmconn.Open();
                OleDbCommand oleCmd = new OleDbCommand("SELECT SESSION_NO FROM REP" + dtpto.Value.ToString("yy") +
                                            " WHERE date_start = #" + dtpto.Value.ToShortDateString() + "# AND " +
                                            "first_bill > 0 and last_bill > 0", rmconn);
                OleDbDataReader dbdr = oleCmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    lblsesto.Text = dbdr.GetValue(0).ToString();
                }
                else
                {
                    lblsesto.Text = "00";
                }
                rmconn.Close();
            }
        }

        private void dtpfrom_ValueChanged(object sender, EventArgs e)
        {
            GetSessionStart();
        }

        private void dtpto_ValueChanged(object sender, EventArgs e)
        {
            GetSessionEnd();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dtpfrom.Value = DateTime.Now;
            dtpto.Value = DateTime.Now;
            GetSessionStart();
            GetSessionEnd();
        }

        private void llExit_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Application.Exit();
        }

        private void llAdmin_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            LoginForm logf = new LoginForm();
            logf.Show();
            this.Hide();
        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
