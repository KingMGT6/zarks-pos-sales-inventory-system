using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Zarks_POS_ERP_System_v2.Classes;

namespace Zarks_POS_ERP_System_v2
{
    public partial class SetupForm : Form
    {
        SetupClass scls;
        public SetupForm()
        {
            InitializeComponent();
        }

        private void SetupForm_Load(object sender, EventArgs e)
        {
            txtstorecode.Select();

            scls = new SetupClass();
            scls.AuthUser();
            scls.AdminSetup();
            scls.LoadMappings();
            //passcode
            txtadmin.Text = scls.admcode;
            //adminsetup
            txtstorecode.Text = scls.storecodemobes;
            txtextrationpath.Text = scls.extractmobeh;
            txtrm.Text = scls.rmpathmobeh;
            //grouptype
            txtGTFood.Text = scls.foodgtplu;
            txtGTBev.Text = scls.bevgtplu;
            txtGTBeer.Text = scls.beergtplu;
            txtGTOI.Text = scls.openitemgtplu;
            txtGTDeli.Text = scls.deliveriesgtplu;
            txtGTO.Text = scls.othersgtplu;
            //payment
            txtPTCC.Text = scls.creditcardsales;
            txtPTGC.Text = scls.gcpayment;
            txtPTGcash.Text = scls.gcashpayment;

            txtPTGcashAR.Text = scls.gcasharpayment;
            txtPTGrabpay.Text = scls.grabpaypayment;
            txtPTGrabpayAR.Text = scls.grabpayarpayment;
            txtPTWechat.Text = scls.wechatpayment;
            txtPTWechatAR.Text = scls.wechatarpayment;
            txtPTFPAR.Text = scls.fparpayment;
            txtPTLLFAR.Text = scls.llfarpayment;
            txtPTGFAR.Text = scls.gfarpayment;
            txtPTZomAR.Text = scls.zomarpayment;
            txtPTSFAR.Text = scls.SFarpayment;
            txtPTMGC.Text = scls.MGCpayment;
            //revcenter
            txtRCSR.Text = scls.scdrc;
            txtRCPWD.Text = scls.pwdrc;
            txtRCZR.Text = scls.zrrc;
            txtRCHB.Text = scls.hbrevc;
            txtRCFP.Text = scls.fprevc;
            txtRCGF.Text = scls.gfrevc;
            txtRCLLF.Text = scls.llfrevc;
            txtRCZOM.Text = scls.zomatorevc;
            txtRCM.Text = scls.manganrevc;
            txtRCCOMP.Text = scls.complirevc;
            txtRCDBy.Text = scls.drivebyrevc;
            txtRCTO.Text = scls.tosales;
            //menugrpitems
            txtMEZM.Text = scls.zarksmodgrpi;
            txtMEDBy.Text = scls.drivebygrpi;
            txtMEMA.Text = scls.menuadjgrpi;
            //discounts
            txtDTScd.Text = scls.scddisc;
            txtDTLV.Text = scls.lessvatdisc;
            txtDTPWD.Text = scls.pwddisc;
            txtDTZR.Text = scls.zrdisc;
            txtDTMKTG.Text = scls.mktgdisc;
            txtDTTEN.Text = scls.tenpcntdisc;
            txtDT100.Text = scls.hndrdpcntdisc;
            //menuitems
            txtMFirst.Text = scls.firstcust;
            txtMBeen.Text = scls.beencust;
            txtMMktgGC.Text = scls.mktggcplu;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
            GenerationForm f1 = new GenerationForm();
            f1.Show();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                //admin pass code
                scls.UpdateAdminCode("@AdminCodeDB", OleDbType.Char, 100, "AdminCode", txtadmin.Text);
                //admin setup mapping
                scls.UpdateAdminAcct("@storecodingDB", OleDbType.Char, 500, "storecoding", txtstorecode.Text);
                scls.UpdateAdminAcct("@extractbehDB", OleDbType.Char, 500, "extractbeh", txtextrationpath.Text);
                scls.UpdateAdminAcct("@rmwinbehDB", OleDbType.Char, 500, "rmwinbeh", txtrm.Text);
                //mappings
                scls.UpdateMappings("@foodGTDB", OleDbType.Char, 500, "FoodGTplu", txtGTFood.Text);
                scls.UpdateMappings("@BevGTpludDB", OleDbType.Char, 500, "BevGTplu", txtGTBev.Text);
                scls.UpdateMappings("@BeerGTpludDB", OleDbType.Char, 500, "BeerGTplu", txtGTBeer.Text);
                scls.UpdateMappings("@OpenItemGTpluDB", OleDbType.Char, 500, "OpenItemGTplu", txtGTOI.Text);
                scls.UpdateMappings("@DeliveriesGTpluDB", OleDbType.Char, 500, "DeliveriesGTplu", txtGTDeli.Text);
                scls.UpdateMappings("@OthersGTpluDB", OleDbType.Char, 500, "OthersGTplu", txtGTO.Text);

                scls.UpdateMappings("@mktgGCpluDB", OleDbType.Char, 500, "mktgGCplu", txtMMktgGC.Text);
                scls.UpdateMappings("@firstcustDB", OleDbType.Char, 500, "firstcust", txtMFirst.Text);
                scls.UpdateMappings("@beencustDB", OleDbType.Char, 500, "beencust", txtMBeen.Text);

                scls.UpdateMappings("@ccardpluDB", OleDbType.Char, 500, "ccardplu", txtPTCC.Text);
                scls.UpdateMappings("@gcmopDB", OleDbType.Char, 500, "gcmop", txtPTGC.Text);
                scls.UpdateMappings("@gcashmopDB", OleDbType.Char, 500, "gcashmop", txtPTGcash.Text);
                scls.UpdateMappings("@gcasarhmopDB", OleDbType.Char, 500, "gcashARmop", txtPTGcashAR.Text);
                scls.UpdateMappings("@grabaymopDB", OleDbType.Char, 500, "grabpaymop", txtPTGrabpay.Text);
                scls.UpdateMappings("@grabayarmopDB", OleDbType.Char, 500, "grabpayARmop", txtPTGrabpayAR.Text);
                scls.UpdateMappings("@wechatmopDB", OleDbType.Char, 500, "wechatmop", txtPTWechat.Text);
                scls.UpdateMappings("@wechatARmopDB", OleDbType.Char, 500, "wechatARmop", txtPTWechatAR.Text);
                scls.UpdateMappings("@fpARmopDB", OleDbType.Char, 500, "fpARmop", txtPTFPAR.Text);
                scls.UpdateMappings("@llfARmopDB", OleDbType.Char, 500, "llfARmop", txtPTLLFAR.Text);
                scls.UpdateMappings("@gfAR", OleDbType.Char, 500, "gfAR", txtPTGFAR.Text);
                scls.UpdateMappings("@zomAR", OleDbType.Char, 500, "zomAR", txtPTZomAR.Text);
                scls.UpdateMappings("@sfAR", OleDbType.Char, 500, "speedfoodAR", txtPTSFAR.Text);
                scls.UpdateMappings("@mgc", OleDbType.Char, 500, "menugocash", txtPTMGC.Text);

                scls.UpdateMappings("@torcpluDB", OleDbType.Char, 500, "torcplu", txtRCTO.Text);
                scls.UpdateMappings("@scdrcpluDB", OleDbType.Char, 500, "scdrcplu", txtRCSR.Text);
                scls.UpdateMappings("@pwdrcpluDB", OleDbType.Char, 500, "pwdrcplu", txtRCPWD.Text);
                scls.UpdateMappings("@zrrcpluDB", OleDbType.Char, 500, "zrrcplu", txtRCZR.Text);
                scls.UpdateMappings("@hbrcDB", OleDbType.Char, 500, "hbrc", txtRCHB.Text);
                scls.UpdateMappings("@fprcDB", OleDbType.Char, 500, "fprc", txtRCFP.Text);
                scls.UpdateMappings("@gfrcDB", OleDbType.Char, 500, "gfrc", txtRCGF.Text);
                scls.UpdateMappings("@llfrcDB", OleDbType.Char, 500, "llfrc", txtRCLLF.Text);
                scls.UpdateMappings("@zomatorcDB", OleDbType.Char, 500, "zomatorc", txtRCZOM.Text);
                scls.UpdateMappings("@manganrcDB", OleDbType.Char, 500, "manganrc", txtRCM.Text);
                scls.UpdateMappings("@complircDB", OleDbType.Char, 500, "complirc", txtRCCOMP.Text);
                scls.UpdateMappings("@drivebyrcDB", OleDbType.Char, 500, "drivebyrc", txtRCDBy.Text);

                scls.UpdateMappings("@scddiscpluDB", OleDbType.Char, 500, "scddiscplu", txtDTScd.Text);
                scls.UpdateMappings("@lessvatdiscpluDB", OleDbType.Char, 500, "lessvatdiscplu", txtDTLV.Text);
                scls.UpdateMappings("@pwddiscpluDB", OleDbType.Char, 500, "pwddiscplu", txtDTPWD.Text);
                scls.UpdateMappings("@zrdiscpluDB", OleDbType.Char, 500, "zrdiscplu", txtDTZR.Text);
                scls.UpdateMappings("@hndrdpcntdiscDB", OleDbType.Char, 500, "hndrdpcntdisc", txtDT100.Text);
                scls.UpdateMappings("@tendiscpluDB", OleDbType.Char, 500, "tendiscplu", txtDTTEN.Text);
                scls.UpdateMappings("@mktgdiscDB", OleDbType.Char, 500, "mktgdisc", txtDTMKTG.Text);

                scls.UpdateMappings("@zarksmodgrpitemsDB", OleDbType.Char, 500, "zarksmodgrpitems", txtMEZM.Text);
                scls.UpdateMappings("@drivebygrpitemsDB", OleDbType.Char, 500, "drivebygrpitems", txtMEDBy.Text);
                scls.UpdateMappings("@menuadjgrpitemsDB", OleDbType.Char, 500, "menuadjgrpitems", txtMEMA.Text);

                MessageBox.Show("Successfully updated!");
                //this.Dispose();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnselectexpath_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog esearch = new FolderBrowserDialog();
            esearch.ShowDialog();
            if (esearch.SelectedPath != null)
            {
                txtextrationpath.Text = esearch.SelectedPath.ToString();
            }
        }

        private void btnselectrm_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog rmsearch = new FolderBrowserDialog();
            rmsearch.ShowDialog();
            if (rmsearch.SelectedPath != null)
            {
                txtrm.Text = rmsearch.SelectedPath.ToString();
            }
        }
    }
}
