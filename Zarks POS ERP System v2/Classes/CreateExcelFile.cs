using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace Zarks_POS_ERP_System_v2.Classes
{
    public class CreateExcelFile : SetupClass
    {
        private string rmstr;

        public CreateExcelFile()
        {
            AdminSetup();
            LoadMappings();
            rmstr = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + rmpathmobeh +
                ";Extended Properties=dBASE IV;User ID=Admin;Password=;";
        }

        #region "VARIABLES"
        #region EXCEL
        public ExcelPackage xclpck;
        public ExcelWorksheet Salexclwrksht;
        public ExcelWorksheet RegDiscxclwrksht;
        public ExcelWorksheet SrDiscxclwrksht;
        public ExcelWorksheet PwdDiscxclwrksht;
        public ExcelWorksheet Compxclwrksht;
        public ExcelWorksheet Hndrdxclwrksht;
        public ExcelWorksheet GCxclwrksht;
        public ExcelWorksheet DSRxclwrksht;
        #endregion
        #region SESSION
        public string datestart { get; set; }
        public string dateend { get; set; }
        public string repyear { get; set; }
        public string repmonth { get; set; }
        public string transdate { get; set; }
        public DateTime strdate { get; set; }
        public DateTime enddate { get; set; }
        public Int32 sessnum { get; set; }
        public string salesstring;
        #endregion
        #region SALES
        public int itempluv { get; set; }
        public string itemdescv { get; set; }
        public double itemqtyv { get; set; }
        public double itemunitpricev { get; set; }
        public double itemtotalv { get; set; }
        #endregion
        #region REG
        //Check
        public int itemplurds { get; set; }
        public string itemdescrds { get; set; }
        public double itemqtyrds { get; set; }
        public double itemunitPrds { get; set; }
        public double itemtotalrds { get; set; }
        public double itemtotalvatrds { get; set; }
        public double itemdiscamtrds { get; set; }
        //Check
        public int citemplurds { get; set; }
        public string citemdescrds { get; set; }
        public double citemqtyrds { get; set; }
        public double citemunitPrds { get; set; }
        public double citemtotalrds { get; set; }
        public double citemtotalvatrds { get; set; }
        public double citemchkdiscamtrds { get; set; }
        public double citemdiscamtrds { get; set; }
        #endregion
        #region SR
        public int itemplusr { get; set; }
        public string itemdescsr { get; set; }
        public double itemqtysr { get; set; }
        public double itemunitpriceVEsr { get; set; }
        public double itemdiscamtsr { get; set; }
        public double itemtotalsr { get; set; }

        public int checkplusr { get; set; }
        public string checkdescsr { get; set; }
        public double checkqtysr { get; set; }
        public double checkunitpriceVEsr { get; set; }
        public double checkdiscamtsr { get; set; }
        public double checktotalsr { get; set; }
        #endregion
        #region PWD
        public int itemplupwd { get; set; }
        public string itemdescpwd { get; set; }
        public double itemqtypwd { get; set; }
        public double itemunitpriceVEpwd { get; set; }
        public double itemdiscamtpwd { get; set; }
        public double itemtotalpwd { get; set; }

        public int checkplupwd { get; set; }
        public string checkdescpwd { get; set; }
        public double checkqtypwd { get; set; }
        public double checkunitpriceVEpwd { get; set; }
        public double checkdiscamtpwd { get; set; }
        public double checktotalpwd { get; set; }
        #endregion
        #region COMPLI
        public int compitemplu { get; set; }
        public string compitemdesc { get; set; }
        public double compitemqty { get; set; }
        public double compitemunitPR { get; set; }
        public int compcode { get; set; }
        public string compname { get; set; }
        #endregion
        #region HUNDISC
        public int hitemplu { get; set; }
        public string hitemdesc { get; set; }
        public double hitemqty { get; set; }
        public double hitemunitPR { get; set; }
        public int hitemcode { get; set; }
        public string hitemname { get; set; }

        public int hchkplu { get; set; }
        public string hchkdesc { get; set; }
        public double hchkqty { get; set; }
        public double hchkunitPR { get; set; }
        public int hchkcode { get; set; }
        public string hchkname { get; set; }
        #endregion
        #region GC
        public int gcplu { get; set; }
        public string gcdesc { get; set; }
        public string gcnum { get; set; }
        public double gcamt { get; set; }
        #endregion
        #region DSR
        public double food { get; set; }
        public double beverage { get; set; }
        public double beer { get; set; }
        public double openItem { get; set; }
        public double delivery { get; set; }
        public double other { get; set; }
        public double mktgamt { get; set; }
        public double bfstotal { get; set; }
        public double lstotal { get; set; }
        public double lscnt { get; set; }
        public double dstotal { get; set; }
        public double tototal { get; set; }
        public int tocnt { get; set; }
        public double ccardsales { get; set; }
        public int ccardcnt { get; set; }
        public int custcnt { get; set; }
        public int firstcustcnt { get; set; }
        public int beencustcnt { get; set; }
        public int tcnt { get; set; }
        public double grosssale { get; set; }
        public double vatables { get; set; }
        public double vatexsales { get; set; }
        public double vatzerosales { get; set; }
        public double outputvats { get; set; }
        public double scdcheck { get; set; }
        public double scditem { get; set; }
        public double pwdcheck { get; set; }
        public double pwditem { get; set; }
        public double mktgcheck { get; set; }
        public double mktgitem { get; set; }
        public double tencheck { get; set; }
        public double tenitem { get; set; }
        public double cashofsales { get; set; }
        public double gcpmt { get; set; }
        public double gcashpmt { get; set; }
        public int gcashpmtcnt { get; set; }
        public double hbrcsales { get; set; }
        public int hbtcsales { get; set; }
        public double fprcsales { get; set; }
        public int fptcsales { get; set; }
        public double gfrcsales { get; set; }
        public int gftcsales { get; set; }
        public double llfrcsales { get; set; }
        public int llftcsales { get; set; }
        public double zomrcsales { get; set; }
        public int zomtcsales { get; set; }
        public double manganrcsales { get; set; }
        public int mangantcsales { get; set; }
        public double excesssales { get; set; }

        public double grabpaypmt { get; set; }
        public int grabpaytcpmt { get; set; }
        public double grabpayarpmt { get; set; }
        public int grabpayartcpmt { get; set; }
        public double gcasharpmt { get; set; }
        public int gcashartcpmt { get; set; }
        public double wechatpmt { get; set; }
        public int wechattcpmt { get; set; }
        public double wechatarpmt { get; set; }
        public int wechatartcpmt { get; set; }
        public double fparpmt { get; set; }
        public int fpartcpmt { get; set; }
        public double llfarpmt { get; set; }
        public int llfartcpmt { get; set; }
        public double gfarpmt { get; set; }
        public int gfartcpmt { get; set; }
        public double zomarpmt { get; set; }
        public int zomartcpmt { get; set; }
        public double sfarpmt { get; set; }
        public int sfartcpmt { get; set; }
        public double mgcpmt { get; set; }
        public int mgctcpmt { get; set; }
        #endregion
        #endregion

        public void GetSession()
        {
            enddate = Convert.ToDateTime(dateend);
            strdate = Convert.ToDateTime(datestart);
            repyear = strdate.ToString("yy");

            using (OleDbConnection rmconn = new OleDbConnection(rmstr))
            {
                rmconn.Open();
                OleDbCommand oleCmd = new OleDbCommand("SELECT SESSION_NO, first_bill, last_bill FROM REP" + repyear +
                                                     " WHERE date_start = #" + strdate + "# and first_bill > 0 and last_bill > 0", rmconn);
                dbdr = oleCmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    sessnum = Convert.ToInt32(dbdr.GetValue(0).ToString());
                    
                    repmonth = strdate.ToString("MM");

                    salesstring = extractmobeh + "\\" + storecodemobes + "-" + strdate.ToString("MMddyyyy") + ".xlsx";
                    if (File.Exists(salesstring))
                    {
                        File.Delete(salesstring);
                    }
                    #region "Remove file creation"
                    /* 
                    xclpck = new ExcelPackage(new FileInfo(salesstring));
                    Salexclwrksht = xclpck.Workbook.Worksheets.Add("Sales");
                    SalesHeader();
                    CreateSales();
                    CloseExcel();

                    xclpck = new ExcelPackage(new FileInfo(salesstring));
                    RegDiscxclwrksht = xclpck.Workbook.Worksheets.Add("SalesWithRegDisc");
                    RegDiscSalesHeader();
                    CreateRegDiscSales();
                    CloseExcel();

                    xclpck = new ExcelPackage(new FileInfo(salesstring));
                    SrDiscxclwrksht = xclpck.Workbook.Worksheets.Add("SalesWithSrDisc");
                    SrDiscSalesHeader();
                    CreateSrDiscSales();
                    CloseExcel();

                    xclpck = new ExcelPackage(new FileInfo(salesstring));
                    PwdDiscxclwrksht = xclpck.Workbook.Worksheets.Add("SalesWithPWDDisc");
                    PwdDiscSalesHeader();
                    CreatePwdDiscSales();
                    CloseExcel();

                    xclpck = new ExcelPackage(new FileInfo(salesstring));
                    Compxclwrksht = xclpck.Workbook.Worksheets.Add("Complimentary");
                    CompHeader();
                    CreateComp();
                    CloseExcel();

                    xclpck = new ExcelPackage(new FileInfo(salesstring));
                    Hndrdxclwrksht = xclpck.Workbook.Worksheets.Add("100%Discount");
                    HndrdDiscSalesHeader();
                    CreateHunDisc();
                    CloseExcel();

                    xclpck = new ExcelPackage(new FileInfo(salesstring));
                    GCxclwrksht = xclpck.Workbook.Worksheets.Add("GC");
                    GCSalesHeader();
                    CreateGC();
                    CloseExcel();
                    */
                    #endregion
                    xclpck = new ExcelPackage(new FileInfo(salesstring));
                    DSRxclwrksht = xclpck.Workbook.Worksheets.Add("DSR");
                    DSRSalesHeader();
                    CreateDSR();
                    CloseExcel();
                }
            }
        }

        public void CloseExcel()
        {
            xclpck.Save();
            xclpck.Dispose();
        }

        #region "HEADERS"
        /* Remove header
        public void SalesHeader()
        {
            Salexclwrksht.Cells["A1"].Value = "ItemCode";
            Salexclwrksht.Cells["B1"].Value = "ItemName";
            Salexclwrksht.Cells["C1"].Value = "Qty";
            Salexclwrksht.Cells["D1"].Value = "UnitPrice";
            Salexclwrksht.Cells["E1"].Value = "TotalSales";
        }
        public void RegDiscSalesHeader()
        {
            RegDiscxclwrksht.Cells["A1"].Value = "ItemCode";
            RegDiscxclwrksht.Cells["B1"].Value = "ItemName";
            RegDiscxclwrksht.Cells["C1"].Value = "Qty";
            RegDiscxclwrksht.Cells["D1"].Value = "UnitPrice";
            RegDiscxclwrksht.Cells["E1"].Value = "TotalSales";
            RegDiscxclwrksht.Cells["F1"].Value = "TotalVat";
            RegDiscxclwrksht.Cells["G1"].Value = "DiscountAmt";
        }
        public void SrDiscSalesHeader()
        {
            SrDiscxclwrksht.Cells["A1"].Value = "ItemCode";
            SrDiscxclwrksht.Cells["B1"].Value = "ItemName";
            SrDiscxclwrksht.Cells["C1"].Value = "Qty";
            SrDiscxclwrksht.Cells["D1"].Value = "UnitPriceVatEx";
            SrDiscxclwrksht.Cells["E1"].Value = "SRDiscount";
            SrDiscxclwrksht.Cells["F1"].Value = "TotalSales";
        }
        public void PwdDiscSalesHeader()
        {
            PwdDiscxclwrksht.Cells["A1"].Value = "ItemCode";
            PwdDiscxclwrksht.Cells["B1"].Value = "ItemName";
            PwdDiscxclwrksht.Cells["C1"].Value = "Qty";
            PwdDiscxclwrksht.Cells["D1"].Value = "UnitPriceVatEx";
            PwdDiscxclwrksht.Cells["E1"].Value = "PWDDiscount";
            PwdDiscxclwrksht.Cells["F1"].Value = "TotalSales";
        }
        public void CompHeader()
        {
            Compxclwrksht.Cells["A1"].Value = "ItemCode";
            Compxclwrksht.Cells["B1"].Value = "ItemName";
            Compxclwrksht.Cells["C1"].Value = "Qty";
            Compxclwrksht.Cells["D1"].Value = "UnitPrice";
            Compxclwrksht.Cells["E1"].Value = "ComCode";
            Compxclwrksht.Cells["F1"].Value = "ComName";
        }
        public void HndrdDiscSalesHeader()
        {
            Hndrdxclwrksht.Cells["A1"].Value = "ItemCode";
            Hndrdxclwrksht.Cells["B1"].Value = "ItemName";
            Hndrdxclwrksht.Cells["C1"].Value = "Qty";
            Hndrdxclwrksht.Cells["D1"].Value = "UnitPrice";
            Hndrdxclwrksht.Cells["E1"].Value = "100%DiscCode";
            Hndrdxclwrksht.Cells["F1"].Value = "100%DiscName";
        }
        public void GCSalesHeader()
        {
            GCxclwrksht.Cells["A1"].Value = "GCCode";
            GCxclwrksht.Cells["B1"].Value = "GCName";
            GCxclwrksht.Cells["C1"].Value = "GCNumber";
            GCxclwrksht.Cells["D1"].Value = "GCAmount";
        }
        */
        public void DSRSalesHeader()
        {
            DSRxclwrksht.Cells["A1"].Value = "FOOD";
            DSRxclwrksht.Cells["B1"].Value = "BEVERAGE";
            DSRxclwrksht.Cells["C1"].Value = "BEER";
            DSRxclwrksht.Cells["D1"].Value = "TOTAL_NET_SALES";
            DSRxclwrksht.Cells["E1"].Value = "BREAKFAST_SALES";
            DSRxclwrksht.Cells["F1"].Value = "LUNCH_SALES";
            DSRxclwrksht.Cells["G1"].Value = "DINNER_SALES";
            DSRxclwrksht.Cells["H1"].Value = "TAKEOUT_SALES";
            DSRxclwrksht.Cells["I1"].Value = "TAKEOUT_TC";
            DSRxclwrksht.Cells["J1"].Value = "CREDIT_CARD_TC";
            DSRxclwrksht.Cells["K1"].Value = "LUNCH_CC";
            DSRxclwrksht.Cells["L1"].Value = "DAILY_CC";
            DSRxclwrksht.Cells["M1"].Value = "FIRST";
            DSRxclwrksht.Cells["N1"].Value = "BEEN";
            DSRxclwrksht.Cells["O1"].Value = "DAILY_PPA";
            DSRxclwrksht.Cells["P1"].Value = "TC";
            DSRxclwrksht.Cells["Q1"].Value = "DAILY_PTA";
            DSRxclwrksht.Cells["R1"].Value = "DAILY_TO_PTA";
            DSRxclwrksht.Cells["S1"].Value = "DAILY_CREDIT_CARD_PTA";
            DSRxclwrksht.Cells["T1"].Value = "AVERAGE_CHECK";
            DSRxclwrksht.Cells["U1"].Value = "TOTAL_GROSS_SALES";
            DSRxclwrksht.Cells["V1"].Value = "VATABLE_SALES";
            DSRxclwrksht.Cells["W1"].Value = "VAT_EXEMPT_SALES";
            DSRxclwrksht.Cells["X1"].Value = "VAT_ZERO_RATED_SALES";
            DSRxclwrksht.Cells["Y1"].Value = "OUTPUT_VAT";
            DSRxclwrksht.Cells["Z1"].Value = "SCD";
            DSRxclwrksht.Cells["AA1"].Value = "PWD";
            DSRxclwrksht.Cells["AB1"].Value = "MKTG_DISCOUNTS";
            DSRxclwrksht.Cells["AC1"].Value = "10%";
            DSRxclwrksht.Cells["AD1"].Value = "CASH_OF_SALES";
            DSRxclwrksht.Cells["AE1"].Value = "CREDIT_CARD_SALES";
            DSRxclwrksht.Cells["AF1"].Value = "GC1";
            DSRxclwrksht.Cells["AG1"].Value = "GC2";
            DSRxclwrksht.Cells["AH1"].Value = "G-Cash";
            DSRxclwrksht.Cells["AI1"].Value = "GcashTC";
            DSRxclwrksht.Cells["AJ1"].Value = "HB";
            DSRxclwrksht.Cells["AK1"].Value = "HBTC";
            DSRxclwrksht.Cells["AL1"].Value = "FP";
            DSRxclwrksht.Cells["AM1"].Value = "FPTC";
            DSRxclwrksht.Cells["AN1"].Value = "Mangan";
            DSRxclwrksht.Cells["AO1"].Value = "ManganTC";
            DSRxclwrksht.Cells["AP1"].Value = "GFSales";
            DSRxclwrksht.Cells["AQ1"].Value = "GFTC";
            DSRxclwrksht.Cells["AR1"].Value = "ZomatoSales";
            DSRxclwrksht.Cells["AS1"].Value = "ZomatoTC";
            DSRxclwrksht.Cells["AT1"].Value = "LLFSales";
            DSRxclwrksht.Cells["AU1"].Value = "LLFTC";
            DSRxclwrksht.Cells["AV1"].Value = "TIPS";
            DSRxclwrksht.Cells["AW1"].Value = "GCashAR";//"Grabpay";
            DSRxclwrksht.Cells["AX1"].Value = "GCashARTC";//"GrabpayTC";
            DSRxclwrksht.Cells["AY1"].Value = "WeChatAR";//"GrabPayAR";
            DSRxclwrksht.Cells["AZ1"].Value = "WeChatARTC";//"GrabPayARTC";
            DSRxclwrksht.Cells["BA1"].Value = "FPAR";
            DSRxclwrksht.Cells["BB1"].Value = "FPARTC";
            DSRxclwrksht.Cells["BC1"].Value = "LalafoodCash";//"WeChat";
            DSRxclwrksht.Cells["BD1"].Value = "LalafoodCashTC";//"WeChatTC";
            DSRxclwrksht.Cells["BE1"].Value = "GrabFoodAR";
            DSRxclwrksht.Cells["BF1"].Value = "GrabFoodARTC";
            DSRxclwrksht.Cells["BG1"].Value = "ZomatoAR";
            DSRxclwrksht.Cells["BH1"].Value = "ZomatoARTC";
            DSRxclwrksht.Cells["BI1"].Value = "SpeedFoodAR";
            DSRxclwrksht.Cells["BJ1"].Value = "SpeedFoodARTC";
            DSRxclwrksht.Cells["BK1"].Value = "MenuGoCash";
            DSRxclwrksht.Cells["BL1"].Value = "MenuGoCashTC";
        }
        #endregion

        /* Remove
        #region "SALES"
        public void CreateSales()
        {
            //Sales
            using (OleDbConnection rmconn = new OleDbConnection(rmstr))
            {
                rmconn.Open();
                //Vatable
                cmd = new OleDbCommand("SELECT a.ref_no, c.descript, Sum(a.quanty), " +
                     "(a.raw_price-a.vat_adj), Sum((ABS(a.vat_adj) + a.raw_price) * a.quanty) " +
                     "FROM (sdet" + repmonth + repyear + " a " +
                     "LEFT JOIN sls" + repmonth + repyear + " b ON a.bill_no=b.bill_no) " +
                     "LEFT JOIN menu c ON a.ref_no=c.ref_no " +
                     "WHERE b.Session_No = " + sessnum + " AND b.pay_type <> 5 " +
                     "AND b.rev_center NOT IN (" + scdrc + "," + zrrc + "," + complirevc + "," + pwdrc + "," + drivebyrevc + ") " +
                     "AND a.disc_no NOT IN (" + scddisc + "," + lessvatdisc + "," + pwddisc + "," + zrdisc + "," + hndrdpcntdisc + "," + tenpcntdisc + "," + mktgdisc + ") " +
                     "AND b.disc_type NOT IN (" + scddisc + "," + lessvatdisc + "," + pwddisc + "," + zrdisc + "," + hndrdpcntdisc + "," + tenpcntdisc + "," + mktgdisc + ") " +
                     "AND a.ref_no NOT IN ("+ mktggcplu +") " +
                     "AND c.Page_num NOT IN (" + zarksmodgrpi + "," + drivebygrpi + "," + menuadjgrpi + ") " +
                     "GROUP BY a.ref_no, c.descript, (a.raw_price-a.vat_adj)", rmconn);

                dbdr = cmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    while (dbdr.Read())
                    {
                        itempluv = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToInt32(ReturnData(dbdr[0].ToString())) : 0;
                        itemdescv = (!DBNull.Value.Equals(dbdr[1])) ? Convert.ToString(dbdr[1].ToString()) : " ";
                        itemqtyv = (!DBNull.Value.Equals(dbdr[2])) ? Convert.ToDouble(ReturnData(dbdr[2].ToString())) : 0;
                        itemunitpricev = (!DBNull.Value.Equals(dbdr[3])) ? Convert.ToDouble(ReturnData(dbdr[3].ToString())) : 0;
                        itemtotalv = (!DBNull.Value.Equals(dbdr[4])) ? Convert.ToDouble(ReturnData(dbdr[4].ToString())) : 0;
                        SaveSales();
                    }
                    dbdr.Close();
                }
                rmconn.Close();
            }
            //GetSalesfromMDB
            using (OleDbConnection mdbkonek = new OleDbConnection(mdbpath))
            {
                mdbkonek.Open();
                string getSales = "Select Sitemcode, Sitemname, Sqty, Sunitprice, Stotalsales from tblSales " +
                    "ORDER BY Sitemcode ASC";
                OleDbDataAdapter ldSales = new OleDbDataAdapter(getSales, mdbkonek);
                DataTable dtSales = new DataTable();
                ldSales.Fill(dtSales);
                Salexclwrksht.Cells["A2"].LoadFromDataTable(dtSales, true);
                Salexclwrksht.DeleteRow(2);
                Salexclwrksht.Column(4).Style.Numberformat.Format = "##0.00";
                Salexclwrksht.Column(5).Style.Numberformat.Format = "##0.00";
                Salexclwrksht.Cells.AutoFitColumns();
                mdbkonek.Close();
            }
        }

        public void SaveSales()
        {
            OpenCon();
            OleDbCommand Scmd = new OleDbCommand("Insert into tblSales(Sitemcode, Sitemname, Sqty, Sunitprice, Stotalsales) " +
                "values(@plunum, @itemdesc, @itemqty, @unitprice, @totasales)", con);
            Scmd.Parameters.AddWithValue("@plunum", itempluv);
            Scmd.Parameters.AddWithValue("@itemdesc", itemdescv);
            Scmd.Parameters.AddWithValue("@itemqty", itemqtyv);
            Scmd.Parameters.AddWithValue("@unitprice", string.Format("{0:##0.#0}", itemunitpricev));
            Scmd.Parameters.AddWithValue("@totasales", string.Format("{0:##0.#0}", itemtotalv));
            Scmd.ExecuteNonQuery();
            closeConn();
        }
        #endregion

        #region "REGDISCSALES"
        public void CreateRegDiscSales()
        {
            //RegDiscSales
            using (OleDbConnection rmconn = new OleDbConnection(rmstr))
            {
                rmconn.Open();
                //Check Discount
                cmd = new OleDbCommand("SELECT a.ref_no, c.descript, Sum(a.quanty), (a.raw_price-a.vat_adj), " +
                     "Sum(a.price_paid * a.quanty), Sum(ABS(a.vat_adj) * a.quanty), " +
                     "Sum((a.quanty * (a.disc_adj * -1))) " + 
                     "FROM (sdet" + repmonth + repyear + " a " +
                     "LEFT JOIN sls" + repmonth + repyear + " b ON a.bill_no=b.bill_no) " +
                     "LEFT JOIN menu c ON a.ref_no=c.ref_no " +
                     "WHERE b.Session_No = " + sessnum + " AND b.pay_type <> 5 " +
                     "AND c.Page_num NOT IN (" + zarksmodgrpi + "," + drivebygrpi + "," + menuadjgrpi + ") " +
                     "AND b.rev_center NOT IN (" + scdrc + "," + zrrc + "," + complirevc + "," + pwdrc + "," + drivebyrevc + ") " +
                     "AND b.disc_type IN (" + tenpcntdisc + "," + hndrdpcntdisc + "," + mktgdisc + ") " +
                     "GROUP BY a.ref_no, c.descript, (a.raw_price-a.vat_adj)", rmconn);
                dbdr = cmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    while (dbdr.Read())
                    {
                        citemplurds = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToInt32(ReturnData(dbdr[0].ToString())) : 0;
                        citemdescrds = (!DBNull.Value.Equals(dbdr[1])) ? Convert.ToString(dbdr[1].ToString()) : " ";
                        citemqtyrds = (!DBNull.Value.Equals(dbdr[2])) ? Convert.ToDouble(ReturnData(dbdr[2].ToString())) : 0;
                        citemunitPrds = (!DBNull.Value.Equals(dbdr[3])) ? Convert.ToDouble(ReturnData(dbdr[3].ToString())) : 0;
                        citemtotalrds = (!DBNull.Value.Equals(dbdr[4])) ? Convert.ToDouble(ReturnData(dbdr[4].ToString())) : 0;
                        citemtotalvatrds = (!DBNull.Value.Equals(dbdr[5])) ? Convert.ToDouble(ReturnData(dbdr[5].ToString())) : 0;
                        citemchkdiscamtrds = (!DBNull.Value.Equals(dbdr[6])) ? Convert.ToDouble(ReturnData(dbdr[6].ToString())) : 0;
                        SaveRegDiscCheckSales();
                    }
                }
                //Item Discount
                cmd = new OleDbCommand("SELECT a.ref_no, c.descript, Sum(a.quanty), (a.raw_price-a.vat_adj), " +
                     "Sum(a.price_paid * a.quanty), Sum(ABS(a.vat_adj) * a.quanty), " +
                     "Sum((a.quanty * (a.item_adj * -1))) " +
                     "FROM (sdet" + repmonth + repyear + " a " +
                     "LEFT JOIN sls" + repmonth + repyear + " b ON a.bill_no=b.bill_no) " +
                     "LEFT JOIN menu c ON a.ref_no=c.ref_no " +
                     "WHERE b.Session_No = " + sessnum + " AND b.pay_type <> 5 " +
                     "AND c.Page_num NOT IN (" + zarksmodgrpi + "," + drivebygrpi + "," + menuadjgrpi + ") " +
                     "AND b.rev_center NOT IN (" + scdrc + "," + zrrc + "," + complirevc + "," + pwdrc + "," + drivebyrevc + ") " +
                     "AND a.disc_no IN (" + tenpcntdisc + "," + hndrdpcntdisc + "," + mktgdisc + ") " +
                     "GROUP BY a.ref_no, c.descript, (a.raw_price-a.vat_adj)", rmconn);
                dbdr = cmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    while (dbdr.Read())
                    {
                        itemplurds = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToInt32(ReturnData(dbdr[0].ToString())) : 0;
                        itemdescrds = (!DBNull.Value.Equals(dbdr[1])) ? Convert.ToString(dbdr[1].ToString()) : " ";
                        itemqtyrds = (!DBNull.Value.Equals(dbdr[2])) ? Convert.ToDouble(ReturnData(dbdr[2].ToString())) : 0;
                        itemunitPrds = (!DBNull.Value.Equals(dbdr[3])) ? Convert.ToDouble(ReturnData(dbdr[3].ToString())) : 0;
                        itemtotalrds = (!DBNull.Value.Equals(dbdr[4])) ? Convert.ToDouble(ReturnData(dbdr[4].ToString())) : 0;
                        itemtotalvatrds = (!DBNull.Value.Equals(dbdr[5])) ? Convert.ToDouble(ReturnData(dbdr[5].ToString())) : 0;
                        itemdiscamtrds = (!DBNull.Value.Equals(dbdr[6])) ? Convert.ToDouble(ReturnData(dbdr[6].ToString())) : 0;
                        SaveRegDiscItemSales();
                    }
                }
                rmconn.Close();
            }
            //GetSalesfromMDB
            using (OleDbConnection mdbkonek = new OleDbConnection(mdbpath))
            {
                mdbkonek.Open();
                string getRDS = "Select Ritemcode, Ritemname, Sum(Rqty), Sum(Runitprice), Sum(Rtotalsales), Sum(Rtotalvat), " +
                    "Sum(Rdiscamt) from tblRegDiscSales GROUP BY Ritemcode, Ritemname ORDER BY Ritemcode ASC";
                OleDbDataAdapter ldRDS = new OleDbDataAdapter(getRDS, mdbkonek);
                DataTable dtRDS = new DataTable();
                ldRDS.Fill(dtRDS);
                RegDiscxclwrksht.Cells["A2"].LoadFromDataTable(dtRDS, true);
                RegDiscxclwrksht.DeleteRow(2);
                RegDiscxclwrksht.Column(4).Style.Numberformat.Format = "##0.00";
                RegDiscxclwrksht.Column(5).Style.Numberformat.Format = "##0.00";
                RegDiscxclwrksht.Column(6).Style.Numberformat.Format = "##0.00";
                RegDiscxclwrksht.Column(7).Style.Numberformat.Format = "##0.00";
                RegDiscxclwrksht.Cells.AutoFitColumns();
                mdbkonek.Close();
            }
        }

        public void SaveRegDiscCheckSales()
        {
            OpenCon();
            OleDbCommand RDScmd = new OleDbCommand("Insert into tblRegDiscSales(Ritemcode, Ritemname, Rqty, Runitprice, " +
                "Rtotalsales, Rtotalvat, Rdiscamt) " +
                "values(@plunum, @itemdesc, @itemqty, @unitprice, @totasales, @totalvat, @discamt)", con);
            RDScmd.Parameters.AddWithValue("@plunum", citemplurds);
            RDScmd.Parameters.AddWithValue("@itemdesc", citemdescrds);
            RDScmd.Parameters.AddWithValue("@itemqty", citemqtyrds);
            RDScmd.Parameters.AddWithValue("@unitprice", string.Format("{0:##0.#0}", citemunitPrds));
            RDScmd.Parameters.AddWithValue("@totasales", string.Format("{0:##0.#0}", citemtotalrds));
            RDScmd.Parameters.AddWithValue("@totalvat", string.Format("{0:##0.#0}", citemtotalvatrds));
            RDScmd.Parameters.AddWithValue("@discamt", string.Format("{0:##0.#0}", citemchkdiscamtrds));
            RDScmd.ExecuteNonQuery();
            closeConn();
        }

        public void SaveRegDiscItemSales()
        {
            OpenCon();
            OleDbCommand RDScmd = new OleDbCommand("Insert into tblRegDiscSales(Ritemcode, Ritemname, Rqty, Runitprice, " +
                "Rtotalsales, Rtotalvat, Rdiscamt) " +
                "values(@plunum, @itemdesc, @itemqty, @unitprice, @totasales, @totalvat, @discamt)", con);
            RDScmd.Parameters.AddWithValue("@plunum", itemplurds);
            RDScmd.Parameters.AddWithValue("@itemdesc", itemdescrds);
            RDScmd.Parameters.AddWithValue("@itemqty", itemqtyrds);
            RDScmd.Parameters.AddWithValue("@unitprice", string.Format("{0:##0.#0}", itemunitPrds));
            RDScmd.Parameters.AddWithValue("@totasales", string.Format("{0:##0.#0}", itemtotalrds));
            RDScmd.Parameters.AddWithValue("@totalvat", string.Format("{0:##0.#0}", itemtotalvatrds));
            RDScmd.Parameters.AddWithValue("@discamt", string.Format("{0:##0.#0}", itemdiscamtrds));
            RDScmd.ExecuteNonQuery();
            closeConn();
        }
        #endregion

        #region "SRDISCSALES"
        public void CreateSrDiscSales()
        {
            //SrDiscSales
            using (OleDbConnection rmconn = new OleDbConnection(rmstr))
            {
                rmconn.Open();
                //item disc
                OleDbCommand srcmd = new OleDbCommand("SELECT a.ref_no, c.descript, Sum(a.quanty), (a.raw_price + a.item_adj), " +
                    "Sum((a.quanty * (a.item_adj * -1))), Sum((a.quanty * a.price_paid)) " +
                     "FROM (sdet" + repmonth + repyear + " a " +
                     "LEFT JOIN sls" + repmonth + repyear + " b ON a.bill_no=b.bill_no) " +
                     "LEFT JOIN menu c ON a.ref_no=c.ref_no " +
                     "WHERE b.Session_No = " + sessnum + " AND b.pay_type <> 5 " +
                     "AND a.disc_no IN (" + scddisc + ") " +
                     "AND b.rev_center IN (" + scdrc + ") " +
                     "AND c.Page_num NOT IN (" + zarksmodgrpi + "," + drivebygrpi + "," + menuadjgrpi + ") " +
                     "GROUP BY a.ref_no, c.descript, (a.raw_price + a.item_adj)", rmconn);
                dbdr = srcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    while (dbdr.Read())
                    {
                        itemplusr = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToInt32(ReturnData(dbdr[0].ToString())) : 0;
                        itemdescsr = (!DBNull.Value.Equals(dbdr[1])) ? Convert.ToString(dbdr[1].ToString()) : " ";
                        itemqtysr = (!DBNull.Value.Equals(dbdr[2])) ? Convert.ToDouble(ReturnData(dbdr[2].ToString())) : 0;
                        itemunitpriceVEsr = (!DBNull.Value.Equals(dbdr[3])) ? Convert.ToDouble(ReturnData(dbdr[3].ToString())) : 0;
                        itemdiscamtsr = (!DBNull.Value.Equals(dbdr[4])) ? Convert.ToDouble(ReturnData(dbdr[4].ToString())) : 0;
                        itemtotalsr = (!DBNull.Value.Equals(dbdr[5])) ? Convert.ToDouble(ReturnData(dbdr[5].ToString())) : 0;
                        SaveSrItemDiscSales();
                    }
                }
                //check disc
                srcmd = new OleDbCommand("SELECT a.ref_no, c.descript, Sum(a.quanty), (a.raw_price + a.item_adj), " +
                    "Sum((a.quanty * (a.disc_adj * -1))), Sum((a.quanty * a.price_paid)) " +
                     "FROM (sdet" + repmonth + repyear + " a " +
                     "LEFT JOIN sls" + repmonth + repyear + " b ON a.bill_no=b.bill_no) " +
                     "LEFT JOIN menu c ON a.ref_no=c.ref_no " +
                     "WHERE b.Session_No = " + sessnum + " AND b.pay_type <> 5 " +
                     "AND b.disc_type IN (" + scddisc + ") " +
                     "AND b.rev_center IN (" + scdrc + ") " +
                     "AND c.Page_num NOT IN (" + zarksmodgrpi + "," + drivebygrpi + "," + menuadjgrpi + ") " +
                     "GROUP BY a.ref_no, c.descript, (a.raw_price + a.item_adj)", rmconn);
                dbdr = srcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    while (dbdr.Read())
                    {
                        checkplusr = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToInt32(ReturnData(dbdr[0].ToString())) : 0;
                        checkdescsr = (!DBNull.Value.Equals(dbdr[1])) ? Convert.ToString(dbdr[1].ToString()) : " ";
                        checkqtysr = (!DBNull.Value.Equals(dbdr[2])) ? Convert.ToDouble(ReturnData(dbdr[2].ToString())) : 0;
                        checkunitpriceVEsr = (!DBNull.Value.Equals(dbdr[3])) ? Convert.ToDouble(ReturnData(dbdr[3].ToString())) : 0;
                        checkdiscamtsr = (!DBNull.Value.Equals(dbdr[4])) ? Convert.ToDouble(ReturnData(dbdr[4].ToString())) : 0;
                        checktotalsr = (!DBNull.Value.Equals(dbdr[5])) ? Convert.ToDouble(ReturnData(dbdr[5].ToString())) : 0;
                        SaveSrCheckDiscSales();
                    }
                }
                rmconn.Close();
            }
            //GetSalesfromMDB
            using (OleDbConnection mdbkonek = new OleDbConnection(mdbpath))
            {
                mdbkonek.Open();
                string getSR = "Select SRitemcode, SRitemname, Sum(SRqty), Sum(SRunitPRVEx), Sum(SRdisc), Sum(SRtotal) " +
                    "from tblSrDiscSales GROUP BY SRitemcode, SRitemname ORDER BY SRitemcode";
                OleDbDataAdapter ldSR = new OleDbDataAdapter(getSR, mdbkonek);
                DataTable dtSR = new DataTable();
                ldSR.Fill(dtSR);
                SrDiscxclwrksht.Cells["A2"].LoadFromDataTable(dtSR, true);
                SrDiscxclwrksht.DeleteRow(2);
                SrDiscxclwrksht.Column(4).Style.Numberformat.Format = "##0.00";
                SrDiscxclwrksht.Column(5).Style.Numberformat.Format = "##0.00";
                SrDiscxclwrksht.Column(6).Style.Numberformat.Format = "##0.00";
                SrDiscxclwrksht.Cells.AutoFitColumns();
                mdbkonek.Close();
            }
        }

        public void SaveSrItemDiscSales()
        {
            OpenCon();
            OleDbCommand SRcmd = new OleDbCommand("Insert into tblSrDiscSales(SRitemcode, SRitemname, SRqty, " +
                "SRunitPRVEx, SRdisc, SRtotal) " +
                "values(@plunum, @itemdesc, @itemqty, @unitpriceVE, @srdisc, @totalsales)", con);
            SRcmd.Parameters.AddWithValue("@plunum", itemplusr);
            SRcmd.Parameters.AddWithValue("@itemdesc", itemdescsr);
            SRcmd.Parameters.AddWithValue("@itemqty", itemqtysr);
            SRcmd.Parameters.AddWithValue("@unitpriceVE", string.Format("{0:##0.#0}", itemunitpriceVEsr));
            SRcmd.Parameters.AddWithValue("@srdisc", string.Format("{0:##0.#0}", itemdiscamtsr));
            SRcmd.Parameters.AddWithValue("@totalsales", string.Format("{0:##0.#0}", itemtotalsr));
            SRcmd.ExecuteNonQuery();
            closeConn();
        }
        public void SaveSrCheckDiscSales()
        {
            OpenCon();
            OleDbCommand SRcmd1 = new OleDbCommand("Insert into tblSrDiscSales(SRitemcode, SRitemname, SRqty, " +
                "SRunitPRVEx, SRdisc, SRtotal) " +
                "values(@plunum1, @itemdesc1, @itemqty1, @unitpriceVE1, @srdisc1, @totalsales1)", con);
            SRcmd1.Parameters.AddWithValue("@plunum1", checkplusr);
            SRcmd1.Parameters.AddWithValue("@itemdesc1", checkdescsr);
            SRcmd1.Parameters.AddWithValue("@itemqty1", checkqtysr);
            SRcmd1.Parameters.AddWithValue("@unitpriceVE1", string.Format("{0:##0.#0}", checkunitpriceVEsr));
            SRcmd1.Parameters.AddWithValue("@srdisc1", string.Format("{0:##0.#0}", checkdiscamtsr));
            SRcmd1.Parameters.AddWithValue("@totalsales1", string.Format("{0:##0.#0}", checktotalsr));
            SRcmd1.ExecuteNonQuery();
            closeConn();
        }
        #endregion

        #region "PWDDISCSALES"
        public void CreatePwdDiscSales()
        {
            //PwdDiscSales
            using (OleDbConnection rmconn = new OleDbConnection(rmstr))
            {
                rmconn.Open();
                //item disc
                OleDbCommand pwdcmd = new OleDbCommand("SELECT a.ref_no, c.descript, Sum(a.quanty), (a.raw_price + a.item_adj), " +
                    "Sum((a.quanty * (a.item_adj * -1))), Sum((a.quanty * a.price_paid)) " +
                     "FROM (sdet" + repmonth + repyear + " a " +
                     "LEFT JOIN sls" + repmonth + repyear + " b ON a.bill_no=b.bill_no) " +
                     "LEFT JOIN menu c ON a.ref_no=c.ref_no " +
                     "WHERE b.Session_No = " + sessnum + " AND b.pay_type <> 5 " +
                     "AND a.disc_no IN (" + pwddisc + ") " +
                     "AND b.rev_center IN (" + pwdrc + ") " +
                     "AND c.Page_num NOT IN (" + zarksmodgrpi + "," + drivebygrpi + "," + menuadjgrpi + ") " +
                     "GROUP BY a.ref_no, c.descript, (a.raw_price + a.item_adj)", rmconn);
                dbdr = pwdcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    while (dbdr.Read())
                    {
                        itemplupwd = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToInt32(ReturnData(dbdr[0].ToString())) : 0;
                        itemdescpwd = (!DBNull.Value.Equals(dbdr[1])) ? Convert.ToString(dbdr[1].ToString()) : " ";
                        itemqtypwd = (!DBNull.Value.Equals(dbdr[2])) ? Convert.ToDouble(ReturnData(dbdr[2].ToString())) : 0;
                        itemunitpriceVEpwd = (!DBNull.Value.Equals(dbdr[3])) ? Convert.ToDouble(ReturnData(dbdr[3].ToString())) : 0;
                        itemdiscamtpwd = (!DBNull.Value.Equals(dbdr[4])) ? Convert.ToDouble(ReturnData(dbdr[4].ToString())) : 0;
                        itemtotalpwd = (!DBNull.Value.Equals(dbdr[5])) ? Convert.ToDouble(ReturnData(dbdr[5].ToString())) : 0;
                        SavePwdItemDiscSales();
                    }
                }
                //check disc
                pwdcmd = new OleDbCommand("SELECT a.ref_no, c.descript, Sum(a.quanty), (a.raw_price + a.item_adj), " +
                    "Sum((a.quanty * (a.disc_adj * -1))), Sum((a.quanty * a.price_paid)) " +
                     "FROM (sdet" + repmonth + repyear + " a " +
                     "LEFT JOIN sls" + repmonth + repyear + " b ON a.bill_no=b.bill_no) " +
                     "LEFT JOIN menu c ON a.ref_no=c.ref_no " +
                     "WHERE b.Session_No = " + sessnum + " AND b.pay_type <> 5 " +
                     "AND b.disc_type IN (" + pwddisc + ") " +
                     "AND b.rev_center IN (" + pwdrc + ") " +
                     "AND c.Page_num NOT IN (" + zarksmodgrpi + "," + drivebygrpi + "," + menuadjgrpi + ") " +
                     "GROUP BY a.ref_no, c.descript, (a.raw_price + a.item_adj)", rmconn);
                dbdr = pwdcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    while (dbdr.Read())
                    {
                        checkplupwd = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToInt32(ReturnData(dbdr[0].ToString())) : 0;
                        checkdescpwd = (!DBNull.Value.Equals(dbdr[1])) ? Convert.ToString(dbdr[1].ToString()) : " ";
                        checkqtypwd = (!DBNull.Value.Equals(dbdr[2])) ? Convert.ToDouble(ReturnData(dbdr[2].ToString())) : 0;
                        checkunitpriceVEpwd = (!DBNull.Value.Equals(dbdr[3])) ? Convert.ToDouble(ReturnData(dbdr[3].ToString())) : 0;
                        checkdiscamtpwd = (!DBNull.Value.Equals(dbdr[4])) ? Convert.ToDouble(ReturnData(dbdr[4].ToString())) : 0;
                        checktotalpwd = (!DBNull.Value.Equals(dbdr[5])) ? Convert.ToDouble(ReturnData(dbdr[5].ToString())) : 0;
                        SavePwdCheckDiscSales();
                    }
                }
                rmconn.Close();
            }
            //GetSalesfromMDB
            using (OleDbConnection mdbkonek = new OleDbConnection(mdbpath))
            {
                mdbkonek.Open();
                string getPWD = "Select PWDitemcode, PWDitemname, Sum(PWDqty), Sum(PWDunitPRVEx), Sum(PWDdisc), Sum(PWDtotal) " +
                    "from tblPwdDiscSales GROUP BY PWDitemcode, PWDitemname ORDER BY PWDitemcode";
                OleDbDataAdapter ldPWD = new OleDbDataAdapter(getPWD, mdbkonek);
                DataTable dtPWD = new DataTable();
                ldPWD.Fill(dtPWD);
                PwdDiscxclwrksht.Cells["A2"].LoadFromDataTable(dtPWD, true);
                PwdDiscxclwrksht.DeleteRow(2);
                PwdDiscxclwrksht.Column(4).Style.Numberformat.Format = "##0.00";
                PwdDiscxclwrksht.Column(5).Style.Numberformat.Format = "##0.00";
                PwdDiscxclwrksht.Column(6).Style.Numberformat.Format = "##0.00";
                PwdDiscxclwrksht.Cells.AutoFitColumns();
                mdbkonek.Close();
            }
        }

        public void SavePwdItemDiscSales()
        {
            OpenCon();
            OleDbCommand PWDcmd = new OleDbCommand("Insert into tblPwdDiscSales(PWDitemcode, PWDitemname, PWDqty, " +
                "PWDunitPRVEx, PWDdisc, PWDtotal) " +
                "values(@plunum, @itemdesc, @itemqty, @unitpriceVE, @srdisc, @totalsales)", con);
            PWDcmd.Parameters.AddWithValue("@plunum", itemplupwd);
            PWDcmd.Parameters.AddWithValue("@itemdesc", itemdescpwd);
            PWDcmd.Parameters.AddWithValue("@itemqty", itemqtypwd);
            PWDcmd.Parameters.AddWithValue("@unitpriceVE", string.Format("{0:##0.#0}", itemunitpriceVEpwd));
            PWDcmd.Parameters.AddWithValue("@srdisc", string.Format("{0:##0.#0}", itemdiscamtpwd));
            PWDcmd.Parameters.AddWithValue("@totalsales", string.Format("{0:##0.#0}", itemtotalpwd));
            PWDcmd.ExecuteNonQuery();
            closeConn();
        }
        public void SavePwdCheckDiscSales()
        {
            OpenCon();
            OleDbCommand PWDcmd1 = new OleDbCommand("Insert into tblPwdDiscSales(PWDitemcode, PWDitemname, PWDqty, " +
                "PWDunitPRVEx, PWDdisc, PWDtotal) " +
                "values(@plunum1, @itemdesc1, @itemqty1, @unitpriceVE1, @srdisc1, @totalsales1)", con);
            PWDcmd1.Parameters.AddWithValue("@plunum1", checkplupwd);
            PWDcmd1.Parameters.AddWithValue("@itemdesc1", checkdescpwd);
            PWDcmd1.Parameters.AddWithValue("@itemqty1", checkqtypwd);
            PWDcmd1.Parameters.AddWithValue("@unitpriceVE1", string.Format("{0:##0.#0}", checkunitpriceVEpwd));
            PWDcmd1.Parameters.AddWithValue("@srdisc1", string.Format("{0:##0.#0}", checkdiscamtpwd));
            PWDcmd1.Parameters.AddWithValue("@totalsales1", string.Format("{0:##0.#0}", checktotalpwd));
            PWDcmd1.ExecuteNonQuery();
            closeConn();
        }
        #endregion

        #region "COMPLIMENTARY"
        public void CreateComp()
        {
            //Comp
            using (OleDbConnection rmconn = new OleDbConnection(rmstr))
            {
                rmconn.Open();
                //Comp
                cmd = new OleDbCommand("SELECT a.ref_no, c.descript, Sum(a.quanty), c.price1, b.rev_center, d.rc_name " +
                     "FROM ((sdet" + repmonth + repyear + " a " +
                     "LEFT JOIN sls" + repmonth + repyear + " b ON a.bill_no=b.bill_no) " +
                     "LEFT JOIN menu c ON a.ref_no=c.ref_no) " +
                     "LEFT JOIN revcent d ON b.rev_center=d.rev_center " +
                     "WHERE b.Session_No = " + sessnum + " AND b.pay_type <> 5 " +
                     "AND b.rev_center IN (" + complirevc + ") " +
                     "AND c.Page_num NOT IN (" + zarksmodgrpi + "," + drivebygrpi + "," + menuadjgrpi + ") " +
                     "GROUP BY a.ref_no, b.rev_center, c.descript, c.price1, d.rc_name", rmconn);

                dbdr = cmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    while (dbdr.Read())
                    {
                        compitemplu = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToInt32(ReturnData(dbdr[0].ToString())) : 0;
                        compitemdesc = (!DBNull.Value.Equals(dbdr[1])) ? Convert.ToString(dbdr[1].ToString()) : " ";
                        compitemqty = (!DBNull.Value.Equals(dbdr[2])) ? Convert.ToDouble(ReturnData(dbdr[2].ToString())) : 0;
                        compitemunitPR = (!DBNull.Value.Equals(dbdr[3])) ? Convert.ToDouble(ReturnData(dbdr[3].ToString())) : 0;
                        compcode = (!DBNull.Value.Equals(dbdr[4])) ? Convert.ToInt32(ReturnData(dbdr[4].ToString())) : 0;
                        compname = (!DBNull.Value.Equals(dbdr[5])) ? Convert.ToString(dbdr[5].ToString()) : " ";
                        SaveComp();
                    }
                    dbdr.Close();
                }
                rmconn.Close();
            }
            //GetSalesfromMDB
            using (OleDbConnection mdbkonek = new OleDbConnection(mdbpath))
            {
                mdbkonek.Open();
                string getComp = "Select Compitemcode, Compitemname, Compqty, CompunitPR, CompCode, CompName " +
                    "from tblComplimentary ORDER BY Compitemcode ASC, CompCode ASC";
                OleDbDataAdapter ldComp = new OleDbDataAdapter(getComp, mdbkonek);
                DataTable dtComp = new DataTable();
                ldComp.Fill(dtComp);
                Compxclwrksht.Cells["A2"].LoadFromDataTable(dtComp, true);
                Compxclwrksht.DeleteRow(2);
                Compxclwrksht.Column(4).Style.Numberformat.Format = "##0.00";
                Compxclwrksht.Cells.AutoFitColumns();
                mdbkonek.Close();
            }
        }

        public void SaveComp()
        {
            OpenCon();
            OleDbCommand Compcmd = new OleDbCommand("Insert into tblComplimentary(Compitemcode, Compitemname, Compqty, " +
                "CompunitPR, CompCode, CompName) values(@Cplu, @Cdesc, @Cqty, @CunitPR, @Ccode, @Cname)", con);
            Compcmd.Parameters.AddWithValue("@Cplu", compitemplu);
            Compcmd.Parameters.AddWithValue("@Cdesc", compitemdesc);
            Compcmd.Parameters.AddWithValue("@Cqty", compitemqty);
            Compcmd.Parameters.AddWithValue("@CunitPR", string.Format("{0:##0.#0}", compitemunitPR));
            Compcmd.Parameters.AddWithValue("@Ccode", compcode);
            Compcmd.Parameters.AddWithValue("@Cname", compname);
            Compcmd.ExecuteNonQuery();
            closeConn();
        }
        #endregion

        #region "100%DISCOUNT"
        public void CreateHunDisc()
        {
            //100%disc
            using (OleDbConnection rmconn = new OleDbConnection(rmstr))
            {
                rmconn.Open();
                //item
                OleDbCommand Hcmd = new OleDbCommand("SELECT a.ref_no, c.descript, Sum(a.quanty), a.raw_price, a.disc_no, d.disc_name " +
                     "FROM ((sdet" + repmonth + repyear + " a " +
                     "LEFT JOIN sls" + repmonth + repyear + " b ON a.bill_no=b.bill_no) " +
                     "LEFT JOIN menu c ON a.ref_no=c.ref_no) " +
                     "LEFT JOIN discount d ON a.disc_no=d.disc_no " +
                     "WHERE b.Session_No = " + sessnum + " AND b.pay_type <> 5 AND a.disc_no IN (" + hndrdpcntdisc + ") " +
                     "AND b.rev_center NOT IN (" + scdrc + "," + pwdrc + "," + zrrc + "," + complirevc + "," + drivebyrevc + ") " +
                     "AND c.Page_num NOT IN (" + zarksmodgrpi + "," + drivebygrpi + "," + menuadjgrpi + ") " +
                     "GROUP BY a.ref_no, a.disc_no, c.descript, a.raw_price, d.disc_name", rmconn);

                dbdr = Hcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    while (dbdr.Read())
                    {
                        hitemplu = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToInt32(ReturnData(dbdr[0].ToString())) : 0;
                        hitemdesc = (!DBNull.Value.Equals(dbdr[1])) ? Convert.ToString(dbdr[1].ToString()) : " ";
                        hitemqty = (!DBNull.Value.Equals(dbdr[2])) ? Convert.ToDouble(ReturnData(dbdr[2].ToString())) : 0;
                        hitemunitPR = (!DBNull.Value.Equals(dbdr[3])) ? Convert.ToDouble(ReturnData(dbdr[3].ToString())) : 0;
                        hitemcode = (!DBNull.Value.Equals(dbdr[4])) ? Convert.ToInt32(ReturnData(dbdr[4].ToString())) : 0;
                        hitemname = (!DBNull.Value.Equals(dbdr[5])) ? Convert.ToString(dbdr[5].ToString()) : " ";
                        SaveItemHunDisc();
                    }
                    dbdr.Close();
                }
                //check
                Hcmd = new OleDbCommand("SELECT a.ref_no, c.descript, Sum(a.quanty), c.price1, a.disc_no, d.disc_name " +
                     "FROM ((sdet" + repmonth + repyear + " a " +
                     "LEFT JOIN sls" + repmonth + repyear + " b ON a.bill_no=b.bill_no) " +
                     "LEFT JOIN menu c ON a.ref_no=c.ref_no) " +
                     "LEFT JOIN discount d ON a.disc_no=d.disc_no " +
                     "WHERE b.Session_No = " + sessnum + " AND b.pay_type <> 5 AND b.disc_type IN (" + hndrdpcntdisc + ") " +
                     "AND b.rev_center NOT IN (" + scdrc + "," + pwdrc + "," + zrrc + "," + complirevc + "," + drivebyrevc + ") " +
                     "AND c.Page_num NOT IN (" + zarksmodgrpi + "," + drivebygrpi + "," + menuadjgrpi + ") " +
                     "GROUP BY a.ref_no, a.disc_no, c.descript, c.price1, d.disc_name", rmconn);

                dbdr = Hcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    while (dbdr.Read())
                    {
                        hchkplu = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToInt32(ReturnData(dbdr[0].ToString())) : 0;
                        hchkdesc = (!DBNull.Value.Equals(dbdr[1])) ? Convert.ToString(dbdr[1].ToString()) : " ";
                        hchkqty = (!DBNull.Value.Equals(dbdr[2])) ? Convert.ToDouble(ReturnData(dbdr[2].ToString())) : 0;
                        hchkunitPR = (!DBNull.Value.Equals(dbdr[3])) ? Convert.ToDouble(ReturnData(dbdr[3].ToString())) : 0;
                        hchkcode = (!DBNull.Value.Equals(dbdr[4])) ? Convert.ToInt32(ReturnData(dbdr[4].ToString())) : 0;
                        hchkname = (!DBNull.Value.Equals(dbdr[5])) ? Convert.ToString(dbdr[5].ToString()) : " ";
                        SaveCheckHunDisc();
                    }
                    dbdr.Close();
                }
                rmconn.Close();
            }
            //GetSalesfromMDB
            using (OleDbConnection mdbkonek = new OleDbConnection(mdbpath))
            {
                mdbkonek.Open();
                string getHunDisc = "Select Hitemcode, Hitemname, Sum(Hqty), Sum(HunitPR), Hcode, Hname " +
                    "from tblHunDisc GROUP BY Hitemcode, Hcode, Hitemname, Hname ORDER BY Hitemcode ASC, Hcode ASC";
                OleDbDataAdapter ldHunDisc = new OleDbDataAdapter(getHunDisc, mdbkonek);
                DataTable dtHunDisc = new DataTable();
                ldHunDisc.Fill(dtHunDisc);
                Hndrdxclwrksht.Cells["A2"].LoadFromDataTable(dtHunDisc, true);
                Hndrdxclwrksht.DeleteRow(2);
                Hndrdxclwrksht.Column(4).Style.Numberformat.Format = "##0.00";
                Hndrdxclwrksht.Cells.AutoFitColumns();
                mdbkonek.Close();
            }
        }

        public void SaveItemHunDisc()
        {
            OpenCon();
            OleDbCommand Compcmd = new OleDbCommand("Insert into tblHunDisc(Hitemcode, Hitemname, Hqty, " +
                "HunitPR, Hcode, Hname) values(@Hplu, @Hdesc, @Hqty, @HunitPR, @Hcode, @Hname)", con);
            Compcmd.Parameters.AddWithValue("@Hplu", hitemplu);
            Compcmd.Parameters.AddWithValue("@Hdesc", hitemdesc);
            Compcmd.Parameters.AddWithValue("@Hqty", hitemqty);
            Compcmd.Parameters.AddWithValue("@HunitPR", string.Format("{0:##0.#0}", hitemunitPR));
            Compcmd.Parameters.AddWithValue("@Hcode", hitemcode);
            Compcmd.Parameters.AddWithValue("@Hname", hitemname);
            Compcmd.ExecuteNonQuery();
            closeConn();
        }
        public void SaveCheckHunDisc()
        {
            OpenCon();
            OleDbCommand CompChkcmd = new OleDbCommand("Insert into tblHunDisc(Hitemcode, Hitemname, Hqty, " +
                "HunitPR, Hcode, Hname) values(@Hplu, @Hdesc, @Hqty, @HunitPR, @Hcode, @Hname)", con);
            CompChkcmd.Parameters.AddWithValue("@Hplu", hchkplu);
            CompChkcmd.Parameters.AddWithValue("@Hdesc", hchkdesc);
            CompChkcmd.Parameters.AddWithValue("@Hqty", hchkqty);
            CompChkcmd.Parameters.AddWithValue("@HunitPR", string.Format("{0:##0.#0}", hchkunitPR));
            CompChkcmd.Parameters.AddWithValue("@Hcode", hchkcode);
            CompChkcmd.Parameters.AddWithValue("@Hname", hchkname);
            CompChkcmd.ExecuteNonQuery();
            closeConn();
        }
        #endregion

        #region "GC"
        public void CreateGC()
        {
            //Sales
            using (OleDbConnection rmconn = new OleDbConnection(rmstr))
            {
                rmconn.Open();
                //Vatable
                cmd = new OleDbCommand("SELECT a.ref_no, c.descript, a.spec_inst, a.price_paid " +
                     "FROM (sdet" + repmonth + repyear + " a " +
                     "LEFT JOIN sls" + repmonth + repyear + " b ON a.bill_no=b.bill_no) " +
                     "LEFT JOIN menu c ON a.ref_no=c.ref_no " +
                     "WHERE b.Session_No = " + sessnum + " AND b.pay_type <> 5 AND a.ref_no IN (" + mktggcplu + ") " +
                     "AND b.rev_center NOT IN (" + complirevc + "," + drivebyrevc + ") " +
                     "AND c.Page_num NOT IN (" + zarksmodgrpi + "," + drivebygrpi + "," + menuadjgrpi + ") " +
                     "GROUP BY a.ref_no, c.descript, a.spec_inst, a.price_paid", rmconn);

                dbdr = cmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    while (dbdr.Read())
                    {
                        gcplu = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToInt32(ReturnData(dbdr[0].ToString())) : 0;
                        gcdesc = (!DBNull.Value.Equals(dbdr[1])) ? Convert.ToString(dbdr[1].ToString()) : " ";
                        gcnum = (!DBNull.Value.Equals(dbdr[2])) ? Convert.ToString(dbdr[2].ToString()) : " ";
                        gcamt = (!DBNull.Value.Equals(dbdr[3])) ? Convert.ToDouble(ReturnData(dbdr[3].ToString())) : 0;
                        SaveGCSales();
                    }
                    dbdr.Close();
                }
                rmconn.Close();
            }
            //GetSalesfromMDB
            using (OleDbConnection mdbkonek = new OleDbConnection(mdbpath))
            {
                mdbkonek.Open();
                string getGCSales = "Select GCcode, GCname, GCnumber, GCamount from tblGC " +
                    "ORDER BY GCcode ASC";
                OleDbDataAdapter ldGCSales = new OleDbDataAdapter(getGCSales, mdbkonek);
                DataTable dtGCSales = new DataTable();
                ldGCSales.Fill(dtGCSales);
                GCxclwrksht.Cells["A2"].LoadFromDataTable(dtGCSales, true);
                GCxclwrksht.DeleteRow(2);
                GCxclwrksht.Column(4).Style.Numberformat.Format = "##0.00";
                GCxclwrksht.Cells.AutoFitColumns();
                mdbkonek.Close();
            }
        }

        public void SaveGCSales()
        {
            OpenCon();
            OleDbCommand GCcmd = new OleDbCommand("Insert into tblGC(GCcode, GCname, GCnumber, GCamount) " +
                "values(@gcplu, @gcdesc, @gcnum, @gcamt)", con);
            GCcmd.Parameters.AddWithValue("@gcplu", gcplu);
            GCcmd.Parameters.AddWithValue("@gcdesc", gcdesc);
            GCcmd.Parameters.AddWithValue("@gcnum", gcnum);
            GCcmd.Parameters.AddWithValue("@gcamt", string.Format("{0:##0.#0}", gcamt));
            GCcmd.ExecuteNonQuery();
            closeConn();
        }
        #endregion
    */
        #region "DSR"
        public void CreateDSR()
        {
            //DSR
            using (OleDbConnection rmconn = new OleDbConnection(rmstr))
            {
                rmconn.Open();
                //Food
                OleDbCommand dsrcmd = new OleDbCommand("SELECT " +
                    "Sum(a.quanty * a.raw_price), Sum(ABS(a.item_adj) * a.quanty), Sum(ABS(a.disc_adj) * a.quanty) " +
                    "FROM (((sdet" + repmonth + repyear + " a " +
                    "LEFT JOIN sls" + repmonth + repyear + " b ON a.bill_no = b.bill_no) " +
                    "LEFT JOIN menu c ON c.ref_no = a.ref_no) " +
                    "LEFT JOIN pages d ON d.page_num = c.page_num) " +
                    "LEFT JOIN pagetype e ON e.page_type = d.page_type " +
                    "WHERE b.Session_No = " + sessnum + " AND b.pay_type <> 5 " +
                    "AND d.page_type IN (" + foodgtplu +") ", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        double rp = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                        double ia = (!DBNull.Value.Equals(dbdr[1])) ? Convert.ToDouble(ReturnData(dbdr[1].ToString())) : 0;
                        double da = (!DBNull.Value.Equals(dbdr[2])) ? Convert.ToDouble(ReturnData(dbdr[2].ToString())) : 0;
                        food = rp - (ia + da);
                    }
                    dbdr.Close();
                }
                //Bev
                dsrcmd = new OleDbCommand("SELECT " +
                    "Sum(a.quanty * a.raw_price), Sum(ABS(a.item_adj) * a.quanty), Sum(ABS(a.disc_adj) * a.quanty) " +
                    "FROM (((sdet" + repmonth + repyear + " a " +
                    "LEFT JOIN sls" + repmonth + repyear + " b ON a.bill_no = b.bill_no) " +
                    "LEFT JOIN menu c ON c.ref_no = a.ref_no) " +
                    "LEFT JOIN pages d ON d.page_num = c.page_num) " +
                    "LEFT JOIN pagetype e ON e.page_type = d.page_type " +
                    "WHERE b.Session_No = " + sessnum + " AND b.pay_type <> 5 " +
                    "AND d.page_type IN (" + bevgtplu + ") ", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        double rp = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                        double ia = (!DBNull.Value.Equals(dbdr[1])) ? Convert.ToDouble(ReturnData(dbdr[1].ToString())) : 0;
                        double da = (!DBNull.Value.Equals(dbdr[2])) ? Convert.ToDouble(ReturnData(dbdr[2].ToString())) : 0;
                        beverage = rp - (ia + da);
                    }
                    dbdr.Close();
                }
                //Beer
                dsrcmd = new OleDbCommand("SELECT " +
                    "Sum(a.quanty * a.raw_price), Sum(ABS(a.item_adj) * a.quanty), Sum(ABS(a.disc_adj) * a.quanty) " +
                    "FROM (((sdet" + repmonth + repyear + " a " +
                    "LEFT JOIN sls" + repmonth + repyear + " b ON a.bill_no = b.bill_no) " +
                    "LEFT JOIN menu c ON c.ref_no = a.ref_no) " +
                    "LEFT JOIN pages d ON d.page_num = c.page_num) " +
                    "LEFT JOIN pagetype e ON e.page_type = d.page_type " +
                    "WHERE b.Session_No = " + sessnum + " AND b.pay_type <> 5 " +
                    "AND d.page_type IN (" + beergtplu + ") ", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        double rp = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                        double ia = (!DBNull.Value.Equals(dbdr[1])) ? Convert.ToDouble(ReturnData(dbdr[1].ToString())) : 0;
                        double da = (!DBNull.Value.Equals(dbdr[2])) ? Convert.ToDouble(ReturnData(dbdr[2].ToString())) : 0;
                        beer = rp - (ia + da);
                    }
                    dbdr.Close();
                }
                //Open Item
                dsrcmd = new OleDbCommand("SELECT " +
                    "Sum(a.quanty * a.raw_price), Sum(ABS(a.item_adj) * a.quanty), Sum(ABS(a.disc_adj) * a.quanty) " +
                    "FROM (((sdet" + repmonth + repyear + " a " +
                    "LEFT JOIN sls" + repmonth + repyear + " b ON a.bill_no = b.bill_no) " +
                    "LEFT JOIN menu c ON c.ref_no = a.ref_no) " +
                    "LEFT JOIN pages d ON d.page_num = c.page_num) " +
                    "LEFT JOIN pagetype e ON e.page_type = d.page_type " +
                    "WHERE b.Session_No = " + sessnum + " AND b.pay_type <> 5 " +
                    "AND d.page_type IN (" + openitemgtplu + ") ", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        double rp = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                        double ia = (!DBNull.Value.Equals(dbdr[1])) ? Convert.ToDouble(ReturnData(dbdr[1].ToString())) : 0;
                        double da = (!DBNull.Value.Equals(dbdr[2])) ? Convert.ToDouble(ReturnData(dbdr[2].ToString())) : 0;
                        openItem = rp - (ia + da);
                    }
                    dbdr.Close();
                }
                //Deliveries
                dsrcmd = new OleDbCommand("SELECT " +
                    "Sum(a.quanty * a.raw_price), Sum(ABS(a.item_adj) * a.quanty), Sum(ABS(a.disc_adj) * a.quanty) " +
                    "FROM (((sdet" + repmonth + repyear + " a " +
                    "LEFT JOIN sls" + repmonth + repyear + " b ON a.bill_no = b.bill_no) " +
                    "LEFT JOIN menu c ON c.ref_no = a.ref_no) " +
                    "LEFT JOIN pages d ON d.page_num = c.page_num) " +
                    "LEFT JOIN pagetype e ON e.page_type = d.page_type " +
                    "WHERE b.Session_No = " + sessnum + " AND b.pay_type <> 5 " +
                    "AND d.page_type IN (" + deliveriesgtplu + ") ", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        double rp = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                        double ia = (!DBNull.Value.Equals(dbdr[1])) ? Convert.ToDouble(ReturnData(dbdr[1].ToString())) : 0;
                        double da = (!DBNull.Value.Equals(dbdr[2])) ? Convert.ToDouble(ReturnData(dbdr[2].ToString())) : 0;
                        delivery = rp - (ia + da);
                    }
                    dbdr.Close();
                }
                //others
                dsrcmd = new OleDbCommand("SELECT " +
                    "Sum(a.quanty * a.raw_price), Sum(ABS(a.item_adj) * a.quanty), Sum(ABS(a.disc_adj) * a.quanty) " +
                    "FROM (((sdet" + repmonth + repyear + " a " +
                    "LEFT JOIN sls" + repmonth + repyear + " b ON a.bill_no = b.bill_no) " +
                    "LEFT JOIN menu c ON c.ref_no = a.ref_no) " +
                    "LEFT JOIN pages d ON d.page_num = c.page_num) " +
                    "LEFT JOIN pagetype e ON e.page_type = d.page_type " +
                    "WHERE b.Session_No = " + sessnum + " AND b.pay_type <> 5 " +
                    "AND d.page_type IN (" + othersgtplu + ") ", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        double rp = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                        double ia = (!DBNull.Value.Equals(dbdr[1])) ? Convert.ToDouble(ReturnData(dbdr[1].ToString())) : 0;
                        double da = (!DBNull.Value.Equals(dbdr[2])) ? Convert.ToDouble(ReturnData(dbdr[2].ToString())) : 0;
                        other = rp - (ia + da);

                    }
                    dbdr.Close();
                }
                //MKTGGC
                dsrcmd = new OleDbCommand("SELECT Sum(ABS(a.price_paid) * a.quanty)" +
                    "FROM sdet" + repmonth + repyear + " a " +
                    "LEFT JOIN sls" + repmonth + repyear + " b ON a.bill_no = b.bill_no " +
                    "WHERE b.Session_No = " + sessnum + " AND b.pay_type <> 5 " +
                    "AND a.ref_no IN (" + mktggcplu + ") ", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        mktgamt = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                //Breakfast
                string breakfaststart = "6:00:00";
                string breakfastend = "9:59:59";
                dsrcmd = new OleDbCommand("SELECT Sum(a.total), Count(a.open_time) " +
                    "FROM sls" + repmonth + repyear + " a WHERE a.Session_No = " + sessnum +
                    " AND a.open_time BETWEEN '" + breakfaststart + "' AND '" + breakfastend + "'", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        bfstotal = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                //Lunch
                string lunchstart = "10:00:00";
                string lunchend = "15:59:59";
                dsrcmd = new OleDbCommand("SELECT Sum(a.total), Count(a.open_time) " +
                    "FROM sls" + repmonth + repyear + " a WHERE a.Session_No = " + sessnum +
                    " AND a.open_time BETWEEN '" + lunchstart + "' AND '" + lunchend + "'", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        lstotal = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                        lscnt = (!DBNull.Value.Equals(dbdr[1])) ? Convert.ToDouble(ReturnData(dbdr[1].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                //Dinner
                string dinnerstart = "16:00:00";
                string dinnerend = "5:59:59";
                dsrcmd = new OleDbCommand("SELECT Sum(a.total), Count(a.open_time) " +
                    "FROM sls" + repmonth + repyear + " a WHERE a.Session_No = " + sessnum +
                    " AND a.open_time BETWEEN '" + dinnerstart + "' AND '" + dinnerend + "'", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        dstotal = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                dsrcmd = new OleDbCommand("SELECT Sum(a.total), Sum(a.people_no) " +
                    "FROM sls" + repmonth + repyear + " a " +
                    "WHERE a.Session_No = " + sessnum + " AND a.pay_type <> 5 " +
                    "AND a.rev_center IN (" + tosales + ") ", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        tototal = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                        tocnt = (!DBNull.Value.Equals(dbdr[1])) ? Convert.ToInt32(ReturnData(dbdr[1].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                //Credit Card
                dsrcmd = new OleDbCommand("SELECT Sum(p.base_amt + p.tip_amt), Count(p.pay_type) " +
                    "FROM pmt" + repmonth + repyear + " p " +
                    "LEFT JOIN sls" + repmonth + repyear + " s ON p.bill_no = s.bill_no " +
                    "WHERE s.Session_No = " + sessnum + " AND p.pay_type <> 5 " +
                    "AND p.pay_type IN (" + creditcardsales + ")", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        ccardsales = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                        ccardcnt = (!DBNull.Value.Equals(dbdr[1])) ? Convert.ToInt32(ReturnData(dbdr[1].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                //Cust Count
                dsrcmd = new OleDbCommand("SELECT Sum(s.people_no) " +
                    "FROM sls" + repmonth + repyear + " s " +
                    "LEFT JOIN pmt" + repmonth + repyear + " p ON p.bill_no = s.bill_no " +
                    "WHERE s.Session_No = " + sessnum + " AND p.pay_type <> 5", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        custcnt = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToInt32(ReturnData(dbdr[0].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                //First Cust
                dsrcmd = new OleDbCommand("SELECT Count(b.bill_no) " +
                    "FROM sls" + repmonth + repyear + " b " +
                    "LEFT JOIN sdet" + repmonth + repyear + " a ON b.bill_no = a.bill_no " +
                    "WHERE b.Session_No = " + sessnum + " AND a.ref_no IN (" + firstcust + ")", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        firstcustcnt = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToInt32(ReturnData(dbdr[0].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                //Been Cust
                dsrcmd = new OleDbCommand("SELECT Count(b.bill_no) " +
                    "FROM sls" + repmonth + repyear + " b " +
                    "LEFT JOIN sdet" + repmonth + repyear + " a ON b.bill_no = a.bill_no " +
                    "WHERE b.Session_No = " + sessnum + " AND a.ref_no IN (" + beencust + ")", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        beencustcnt = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToInt32(ReturnData(dbdr[0].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                //TC
                dsrcmd = new OleDbCommand("SELECT Count(a.bill_no) FROM (SELECT DISTINCT a.bill_no " +
                    "FROM sdet" + repmonth + repyear + " a " +
                    "LEFT JOIN sls" + repmonth + repyear + " b ON a.bill_no = b.bill_no " +
                    "WHERE b.Session_No = " + sessnum + ")", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        tcnt = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToInt32(ReturnData(dbdr[0].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                //Gross Sales
                dsrcmd = new OleDbCommand("SELECT Sum(p.tip_amt), Sum(p.base_amt) " +
                    "FROM sls" + repmonth + repyear + " s " +
                    "LEFT JOIN pmt" + repmonth + repyear + " p ON s.bill_no = p.bill_no " +
                    "WHERE s.Session_No = " + sessnum + " AND p.pay_type <> 5 AND s.pay_type <> 5", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        double pmttip = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                        double pmtbaseamt = (!DBNull.Value.Equals(dbdr[1])) ? Convert.ToDouble(ReturnData(dbdr[1].ToString())) : 0;
                        //double slstaxes = (!DBNull.Value.Equals(dbdr[2])) ? Convert.ToDouble(ReturnData(dbdr[2].ToString())) : 0;
                        grosssale = pmttip + pmtbaseamt;//slstotal + slstaxes;
                    }
                    dbdr.Close();
                }
                //Vatable Sales
                dsrcmd = new OleDbCommand("SELECT Sum(total) " +
                    "FROM sls" + repmonth + repyear + " WHERE Session_No = " + sessnum + " AND pay_type <> 5 " +
                    "AND taxable = true", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        vatables = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                //Vat Exempt Sales
                dsrcmd = new OleDbCommand("SELECT Sum(total + discount) " +
                    "FROM sls" + repmonth + repyear + " WHERE Session_No = " + sessnum + " AND pay_type <> 5 " +
                    "AND taxable = false AND rev_center IN (" + scdrc + "," + pwdrc + ")", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        vatexsales = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                //Zero Rated Sales
                dsrcmd = new OleDbCommand("SELECT Sum(total + discount) " +
                    "FROM sls" + repmonth + repyear + " WHERE Session_No = " + sessnum + " AND pay_type <> 5 " +
                    "AND taxable = false AND rev_center IN (" + zrrc + ")", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        vatzerosales = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                //Output Vat
                dsrcmd = new OleDbCommand("SELECT Sum(taxes) " +
                    "FROM rep" + repyear + " WHERE Session_No = " + sessnum + "", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        outputvats = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                //SCD CHECK
                dsrcmd = new OleDbCommand("SELECT Sum(ABS(a.disc_adj) * a.quanty) " +
                    "FROM sdet" + repmonth + repyear + " a " +
                    "LEFT JOIN sls" + repmonth + repyear + " b ON a.bill_no = b.bill_no " +
                    "WHERE b.Session_No = " + sessnum + " AND b.pay_type <> 5 " +
                    "AND b.disc_type IN (" + scddisc + ")", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        scdcheck = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                //SCD ITEM
                dsrcmd = new OleDbCommand("SELECT Sum(ABS(a.item_adj) * a.quanty) " +
                    "FROM sdet" + repmonth + repyear + " a " +
                    "LEFT JOIN sls" + repmonth + repyear + " b ON a.bill_no = b.bill_no " +
                    "WHERE b.Session_No = " + sessnum + " AND b.pay_type <> 5 " +
                    "AND a.disc_no IN (" + scddisc + ")", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        scditem = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                //PWD CHECK
                dsrcmd = new OleDbCommand("SELECT Sum(ABS(a.disc_adj) * a.quanty) " +
                    "FROM sdet" + repmonth + repyear + " a " +
                    "LEFT JOIN sls" + repmonth + repyear + " b ON a.bill_no = b.bill_no " +
                    "WHERE b.Session_No = " + sessnum + " AND b.pay_type <> 5 " +
                    "AND b.disc_type IN (" + pwddisc + ")", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        pwdcheck = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                //PWD ITEM
                dsrcmd = new OleDbCommand("SELECT Sum(ABS(a.item_adj) * a.quanty) " +
                    "FROM sdet" + repmonth + repyear + " a " +
                    "LEFT JOIN sls" + repmonth + repyear + " b ON a.bill_no = b.bill_no " +
                    "WHERE b.Session_No = " + sessnum + " AND b.pay_type <> 5 " +
                    "AND a.disc_no IN (" + pwddisc + ")", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        pwditem = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                //MKTG CHECK
                dsrcmd = new OleDbCommand("SELECT Sum(ABS(a.disc_adj) * a.quanty) " +
                    "FROM sdet" + repmonth + repyear + " a " +
                    "LEFT JOIN sls" + repmonth + repyear + " b ON a.bill_no = b.bill_no " +
                    "WHERE b.Session_No = " + sessnum + " AND b.pay_type <> 5 " +
                    "AND b.disc_type IN (" + mktgdisc + ")", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        mktgcheck = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                //MKTG ITEM
                dsrcmd = new OleDbCommand("SELECT Sum(ABS(a.item_adj) * a.quanty) " +
                    "FROM sdet" + repmonth + repyear + " a " +
                    "LEFT JOIN sls" + repmonth + repyear + " b ON a.bill_no = b.bill_no " +
                    "WHERE b.Session_No = " + sessnum + " AND b.pay_type <> 5 " +
                    "AND a.disc_no IN (" + mktgdisc + ")", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        mktgitem = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                //TEN PCNT DISC CHECK
                dsrcmd = new OleDbCommand("SELECT Sum(ABS(a.disc_adj) * a.quanty) " +
                    "FROM sdet" + repmonth + repyear + " a " +
                    "LEFT JOIN sls" + repmonth + repyear + " b ON a.bill_no = b.bill_no " +
                    "WHERE b.Session_No = " + sessnum + " AND b.pay_type <> 5 " +
                    "AND b.disc_type IN (" + tenpcntdisc + ")", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        tencheck = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                //TEN PCNT DISC ITEM
                dsrcmd = new OleDbCommand("SELECT Sum(ABS(a.item_adj) * a.quanty) " +
                    "FROM sdet" + repmonth + repyear + " a " +
                    "LEFT JOIN sls" + repmonth + repyear + " b ON a.bill_no = b.bill_no " +
                    "WHERE b.Session_No = " + sessnum + " AND b.pay_type <> 5 " +
                    "AND a.disc_no IN (" + tenpcntdisc + ")", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        tenitem = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                //CASH SALES
                dsrcmd = new OleDbCommand("SELECT Sum(p.base_amt + p.tip_amt) " +
                    "FROM pmt" + repmonth + repyear + " p " +
                    "LEFT JOIN sls" + repmonth + repyear + " s ON p.bill_no = s.bill_no " +
                    "WHERE s.Session_No = " + sessnum + " AND p.pay_type IN (1)", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        cashofsales = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                //GC MOP
                dsrcmd = new OleDbCommand("SELECT Sum(p.base_amt + p.tip_amt) " +
                    "FROM pmt" + repmonth + repyear + " p " +
                    "LEFT JOIN sls" + repmonth + repyear + " s ON p.bill_no = s.bill_no " +
                    "WHERE s.Session_No = " + sessnum + " AND p.pay_type IN (" + gcpayment + ")", rmconn);

                dbdr = dsrcmd.ExecuteReader();  
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        gcpmt = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                //GCASH MOP
                dsrcmd = new OleDbCommand("SELECT Sum(p.base_amt + p.tip_amt) " +
                    "FROM pmt" + repmonth + repyear + " p " +
                    "LEFT JOIN sls" + repmonth + repyear + " s ON p.bill_no = s.bill_no " +
                    "WHERE s.Session_No = " + sessnum + " AND p.pay_type IN (" + gcashpayment + ")", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        gcashpmt = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                //GCASH COUNT
                dsrcmd = new OleDbCommand("SELECT Count(p.pay_type) " +
                    "FROM pmt" + repmonth + repyear + " p " +
                    "LEFT JOIN sls" + repmonth + repyear + " s ON p.bill_no = s.bill_no " +
                    "WHERE s.Session_No = " + sessnum + " AND p.pay_type IN (" + gcashpayment + ")", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        gcashpmtcnt = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToInt32(ReturnData(dbdr[0].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                //HB SALES
                dsrcmd = new OleDbCommand("SELECT Sum(total), Count(people_no) " +
                    "FROM sls" + repmonth + repyear + " " +
                    "WHERE Session_No = " + sessnum + " AND pay_type <> 5 AND rev_center IN (" + hbrevc + ")", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        hbrcsales = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                        hbtcsales = (!DBNull.Value.Equals(dbdr[1])) ? Convert.ToInt32(ReturnData(dbdr[1].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                //FB SALES
                dsrcmd = new OleDbCommand("SELECT Sum(total), Count(people_no) " +
                    "FROM sls" + repmonth + repyear + " " +
                    "WHERE Session_No = " + sessnum + " AND pay_type <> 5 AND rev_center IN (" + fprevc + ")", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        fprcsales = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                        fptcsales = (!DBNull.Value.Equals(dbdr[1])) ? Convert.ToInt32(ReturnData(dbdr[1].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                //GF SALES
                dsrcmd = new OleDbCommand("SELECT Sum(total), Count(people_no) " +
                    "FROM sls" + repmonth + repyear + " " +
                    "WHERE Session_No = " + sessnum + " AND pay_type <> 5 AND rev_center IN (" + gfrevc + ")", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        gfrcsales = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                        gftcsales = (!DBNull.Value.Equals(dbdr[1])) ? Convert.ToInt32(ReturnData(dbdr[1].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                //LLF SALES
                dsrcmd = new OleDbCommand("SELECT Sum(total), Count(people_no) " +
                    "FROM sls" + repmonth + repyear + " " +
                    "WHERE Session_No = " + sessnum + " AND pay_type <> 5 AND rev_center IN (" + llfrevc + ")", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        llfrcsales = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                        llftcsales = (!DBNull.Value.Equals(dbdr[1])) ? Convert.ToInt32(ReturnData(dbdr[1].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                //ZOMATO SALES
                dsrcmd = new OleDbCommand("SELECT Sum(total), Count(people_no) " +
                    "FROM sls" + repmonth + repyear + " " +
                    "WHERE Session_No = " + sessnum + " AND pay_type <> 5 AND rev_center IN (" + zomatorevc + ")", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        zomrcsales = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                        zomtcsales = (!DBNull.Value.Equals(dbdr[1])) ? Convert.ToInt32(ReturnData(dbdr[1].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                //MANGAN SALES
                dsrcmd = new OleDbCommand("SELECT Sum(total), Count(people_no) " +
                    "FROM sls" + repmonth + repyear + " " +
                    "WHERE Session_No = " + sessnum + " AND pay_type <> 5 AND rev_center IN (" + manganrevc + ")", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        manganrcsales = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                        mangantcsales = (!DBNull.Value.Equals(dbdr[1])) ? Convert.ToInt32(ReturnData(dbdr[1].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                //EXCESS
                dsrcmd = new OleDbCommand("SELECT Sum(p.tip_amt) " +
                    "FROM pmt" + repmonth + repyear + " p " +
                    "LEFT JOIN sls" + repmonth + repyear + " s ON p.bill_no = s.bill_no " +
                    "WHERE s.Session_No = " + sessnum + " AND p.pay_type <> 5", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        excesssales = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                /*
                //Grabpay
                dsrcmd = new OleDbCommand("SELECT Sum(p.base_amt + p.tip_amt), Count(p.pay_type) " +
                    "FROM pmt" + repmonth + repyear + " p " +
                    "LEFT JOIN sls" + repmonth + repyear + " s ON p.bill_no = s.bill_no " +
                    "WHERE s.Session_No = " + sessnum + " AND p.pay_type IN (" + grabpaypayment + ")", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        grabpaypmt = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                        grabpaytcpmt = (!DBNull.Value.Equals(dbdr[1])) ? Convert.ToInt32(ReturnData(dbdr[1].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                //Grabpay AR
                dsrcmd = new OleDbCommand("SELECT Sum(p.base_amt + p.tip_amt), Count(p.pay_type) " +
                    "FROM pmt" + repmonth + repyear + " p " +
                    "LEFT JOIN sls" + repmonth + repyear + " s ON p.bill_no = s.bill_no " +
                    "WHERE s.Session_No = " + sessnum + " AND p.pay_type IN (" + grabpayarpayment + ")", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        grabpayarpmt = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                        grabpayartcpmt = (!DBNull.Value.Equals(dbdr[1])) ? Convert.ToInt32(ReturnData(dbdr[1].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                */
                //G-Cash AR
                dsrcmd = new OleDbCommand("SELECT Sum(p.base_amt + p.tip_amt), Count(p.pay_type) " +
                    "FROM pmt" + repmonth + repyear + " p " +
                    "LEFT JOIN sls" + repmonth + repyear + " s ON p.bill_no = s.bill_no " +
                    "WHERE s.Session_No = " + sessnum + " AND p.pay_type IN (" + gcasharpayment + ")", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        gcasharpmt = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                        gcashartcpmt = (!DBNull.Value.Equals(dbdr[1])) ? Convert.ToInt32(ReturnData(dbdr[1].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                /*
                //We Chat
                dsrcmd = new OleDbCommand("SELECT Sum(p.base_amt + p.tip_amt), Count(p.pay_type) " +
                    "FROM pmt" + repmonth + repyear + " p " +
                    "LEFT JOIN sls" + repmonth + repyear + " s ON p.bill_no = s.bill_no " +
                    "WHERE s.Session_No = " + sessnum + " AND p.pay_type IN (" + wechatpayment + ")", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        wechatpmt = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                        wechattcpmt = (!DBNull.Value.Equals(dbdr[1])) ? Convert.ToInt32(ReturnData(dbdr[1].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                */
                //We Chat AR
                dsrcmd = new OleDbCommand("SELECT Sum(p.base_amt + p.tip_amt), Count(p.pay_type) " +
                    "FROM pmt" + repmonth + repyear + " p " +
                    "LEFT JOIN sls" + repmonth + repyear + " s ON p.bill_no = s.bill_no " +
                    "WHERE s.Session_No = " + sessnum + " AND p.pay_type IN (" + wechatarpayment + ")", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        wechatarpmt = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                        wechatartcpmt = (!DBNull.Value.Equals(dbdr[1])) ? Convert.ToInt32(ReturnData(dbdr[1].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                //FP AR
                dsrcmd = new OleDbCommand("SELECT Sum(p.base_amt + p.tip_amt), Count(p.pay_type) " +
                    "FROM pmt" + repmonth + repyear + " p " +
                    "LEFT JOIN sls" + repmonth + repyear + " s ON p.bill_no = s.bill_no " +
                    "WHERE s.Session_No = " + sessnum + " AND p.pay_type IN (" + fparpayment + ")", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        fparpmt = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                        fpartcpmt = (!DBNull.Value.Equals(dbdr[1])) ? Convert.ToInt32(ReturnData(dbdr[1].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                //Lalafood AR
                dsrcmd = new OleDbCommand("SELECT Sum(p.base_amt + p.tip_amt), Count(p.pay_type) " +
                    "FROM pmt" + repmonth + repyear + " p " +
                    "LEFT JOIN sls" + repmonth + repyear + " s ON p.bill_no = s.bill_no " +
                    "WHERE s.Session_No = " + sessnum + " AND p.pay_type IN (" + llfarpayment + ")", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        llfarpmt = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                        llfartcpmt = (!DBNull.Value.Equals(dbdr[1])) ? Convert.ToInt32(ReturnData(dbdr[1].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                //Grabfood AR
                dsrcmd = new OleDbCommand("SELECT Sum(p.base_amt + p.tip_amt), Count(p.pay_type) " +
                    "FROM pmt" + repmonth + repyear + " p " +
                    "LEFT JOIN sls" + repmonth + repyear + " s ON p.bill_no = s.bill_no " +
                    "WHERE s.Session_No = " + sessnum + " AND p.pay_type IN (" + gfarpayment + ")", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        gfarpmt = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                        gfartcpmt = (!DBNull.Value.Equals(dbdr[1])) ? Convert.ToInt32(ReturnData(dbdr[1].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                //Zomato AR
                dsrcmd = new OleDbCommand("SELECT Sum(p.base_amt + p.tip_amt), Count(p.pay_type) " +
                    "FROM pmt" + repmonth + repyear + " p " +
                    "LEFT JOIN sls" + repmonth + repyear + " s ON p.bill_no = s.bill_no " +
                    "WHERE s.Session_No = " + sessnum + " AND p.pay_type IN (" + zomarpayment + ")", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        zomarpmt = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                        zomartcpmt = (!DBNull.Value.Equals(dbdr[1])) ? Convert.ToInt32(ReturnData(dbdr[1].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                //Speed Food AR
                dsrcmd = new OleDbCommand("SELECT Sum(p.base_amt + p.tip_amt), Count(p.pay_type) " +
                    "FROM pmt" + repmonth + repyear + " p " +
                    "LEFT JOIN sls" + repmonth + repyear + " s ON p.bill_no = s.bill_no " +
                    "WHERE s.Session_No = " + sessnum + " AND p.pay_type IN (" + SFarpayment + ")", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        sfarpmt = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                        sfartcpmt = (!DBNull.Value.Equals(dbdr[1])) ? Convert.ToInt32(ReturnData(dbdr[1].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                //MenuGoCash
                dsrcmd = new OleDbCommand("SELECT Sum(p.base_amt + p.tip_amt), Count(p.pay_type) " +
                    "FROM pmt" + repmonth + repyear + " p " +
                    "LEFT JOIN sls" + repmonth + repyear + " s ON p.bill_no = s.bill_no " +
                    "WHERE s.Session_No = " + sessnum + " AND p.pay_type IN (" + MGCpayment + ")", rmconn);

                dbdr = dsrcmd.ExecuteReader();
                if (dbdr.HasRows)
                {
                    dbdr.Read();
                    {
                        mgcpmt = (!DBNull.Value.Equals(dbdr[0])) ? Convert.ToDouble(ReturnData(dbdr[0].ToString())) : 0;
                        mgctcpmt = (!DBNull.Value.Equals(dbdr[1])) ? Convert.ToInt32(ReturnData(dbdr[1].ToString())) : 0;
                    }
                    dbdr.Close();
                }
                SaveDSR();
                rmconn.Close();
            }
            //GetSalesfromMDB
            using (OleDbConnection mdbkonek = new OleDbConnection(mdbpath))
            {
                mdbkonek.Open();
                string getDSR = "Select dsrfood, dsrbev, dsrbeer, dsrnetsales, dsrbfsales, dsrlunchsales, " +
                    "dsrdinnersales, dsrtosales, dsrtotc, dsrccardtc, dsrlunchcc, dsrdailycc, dsrfirst, dsrbeen, " +
                    "dsrdailyppa, dsrtc, dsrdailypta, dsrdailytopta, dsrdailyccardpta, dsraveragecheck, dsrgrosssales, " +
                    "dsrvatsales, dsrvatexempt, dsrzerorated, dsroutputvat, dsrscd, dsrpwd, dsrmktgdisc, dsrtendisc, " +
                    "dsrcashsales, dsrccardsales, dsrgcone, dsrgctwo, dsrgcash, dsrgcashtc, dsrhb, dsrhbtc, dsrfp, " +
                    "dsrfptc, dsrmangan, dsrmangantc, dsrgf, dsrgftc, dsrzomato, dsrzomatotc, dsrllf, dsrllftc, " +
                    "dsrtips, dsrgrabpay, dsrgrabpaytc, dsrgrabpayAR, dsrgrabpayARTC, dsrgcashAR, dsrgcashARTC, " +
                    "dsrwechat, dsrwechatTC, dsrwechatAR, dsrwechatARTC, dsrfpAR, dsrfpARTC, dsrllfAR, dsrllfARTC, dsrgfAR, " +
                    "dsrgfARTC, dsrzomAR, dsrzomARTC, dsrsfAR, dsrsfARTC, dsrmgc, dsrmgcTC from tblDSR";
                OleDbDataAdapter ldDSR = new OleDbDataAdapter(getDSR, mdbkonek);
                DataTable dtDSR = new DataTable();
                ldDSR.Fill(dtDSR);
                DSRxclwrksht.Cells["A2"].LoadFromDataTable(dtDSR, true);
                DSRxclwrksht.DeleteRow(2);
                DSRxclwrksht.Column(1).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(2).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(3).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(4).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(5).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(6).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(7).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(8).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(15).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(17).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(18).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(19).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(20).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(21).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(22).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(23).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(24).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(25).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(26).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(27).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(28).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(29).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(30).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(31).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(32).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(33).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(34).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(36).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(38).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(40).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(42).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(44).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(46).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(48).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(49).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(51).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(53).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(55).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(57).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(59).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(61).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(63).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(65).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(67).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Column(69).Style.Numberformat.Format = "##0.00";
                DSRxclwrksht.Cells.AutoFitColumns();
                mdbkonek.Close();
            }
        }

        public void SaveDSR()
        {
            OpenCon();
            OleDbCommand dsrcmd = new OleDbCommand("Insert into tblDSR(dsrfood, dsrbev, dsrbeer, dsrnetsales, dsrbfsales, " +
                "dsrlunchsales, dsrdinnersales, dsrtosales, dsrtotc, dsrccardtc, dsrlunchcc, dsrdailycc, dsrfirst, dsrbeen, " +
                "dsrdailyppa, dsrtc, dsrdailypta, dsrdailytopta, dsrdailyccardpta, dsraveragecheck, dsrgrosssales, dsrvatsales, " +
                "dsrvatexempt, dsrzerorated, dsroutputvat, dsrscd, dsrpwd, dsrmktgdisc, dsrtendisc, dsrcashsales, dsrccardsales, " +
                "dsrgcone, dsrgctwo, dsrgcash, dsrgcashtc, dsrhb, dsrhbtc, dsrfp, dsrfptc, dsrmangan, dsrmangantc, dsrgf, " +
                "dsrgftc, dsrzomato, dsrzomatotc, dsrllf, dsrllftc, dsrtips, dsrgrabpay, dsrgrabpaytc, dsrgrabpayAR, " +
                "dsrgrabpayARTC, dsrgcashAR, dsrgcashARTC, dsrwechat, dsrwechatTC, dsrwechatAR, dsrwechatARTC, dsrfpAR, " +
                "dsrfpARTC, dsrllfAR, dsrllfARTC, dsrgfAR, dsrgfARTC, dsrzomAR, dsrzomARTC, dsrsfAR, dsrsfARTC, dsrmgc, dsrmgcTC) " +
                "values(@gtfood, @gtbev, @gtbeer, @netsales, @bfsales, @lunchsales, @dinnersales, @tosales, @tocnt, @ccardtc, " +
                "@lunchcc, @dailycc, @first, @been, @dailyppa, @tc, @dailypta, @dailytopta, @dailyccardpta, @averagecheck, " +
                "@grosssales, @vatsales, @vatexempt, @zerorated, @outputvat, @scd, @pwd, @mktgdisc, @tenpcntdisc, @cos, " +
                "@ccardsales, @mktggcsales, @gctwo, @gcash, @gcashcnt, @hb, @hbtc, @fp, @fptc, @mangan, @mangantc, " +
                "@gf, @gftc, @zomato, @zomatotc, @llf, @llftc, @tips, @gp, @gptc, @gpar, @gpartc, @gcashar, @gcashartc, " +
                "@wechat, @wechattc, @wechatar, @wechatartc, @fpar, @fpartc, @llfar, @llfartc, @gfar, " +
                "@gfartc, @zomar, @zomartc, @sfAR, @sfARTC, @mgc, @mgcTC)", con);
            //dnetsales
            double dnetsales = ((food + beverage + beer + openItem + delivery + other) - mktgamt);
            //ddailyppa
            double ddailyppa;
            if (custcnt != 0)
            {
                ddailyppa = (((food + beverage + beer + openItem + delivery + other) - mktgamt) / custcnt);
            }
            else
            {
                ddailyppa = 0;
            }
            //ddailypta
            double ddailypta;
            if (tcnt != 0)
            {
                ddailypta = (((food + beverage + beer + openItem + delivery + other) - mktgamt) / tcnt);
            }
            else
            {
                ddailypta = 0;
            }
            //ddailytopta
            double ddailytopta;
            if (tocnt != 0)
            {
                ddailytopta = (tototal / tocnt);
            }
            else
            {
                ddailytopta = 0;
            }
            //ddailyccardpta
            double ddailyccardpta;
            if (ccardcnt !=0)
            {
                ddailyccardpta = (ccardsales / ccardcnt);
            }
            else
            {
                ddailyccardpta = 0;
            }
            //daveragecheck
            double daveragecheck;
            if (custcnt != 0)
            {
                daveragecheck = (vatables / custcnt);
            }
            else
            {
                daveragecheck = 0;
            }
            //others
            double dscd = (scdcheck + scditem);
            double dpwd = (pwdcheck + pwditem);
            double dmktgdisc = (mktgcheck + mktgitem);
            double dtenpcntdisc = (tencheck + tenitem);
            double dtips = (openItem + excesssales);

            dsrcmd.Parameters.AddWithValue("@gtfood", string.Format("{0:##0.#0}", food));
            dsrcmd.Parameters.AddWithValue("@gtbev", string.Format("{0:##0.#0}", beverage));
            dsrcmd.Parameters.AddWithValue("@gtbeer", string.Format("{0:##0.#0}", beer));
            dsrcmd.Parameters.AddWithValue("@netsales", string.Format("{0:##0.#0}", dnetsales));
            dsrcmd.Parameters.AddWithValue("@bfsales", string.Format("{0:##0.#0}", bfstotal));
            dsrcmd.Parameters.AddWithValue("@lunchsales", string.Format("{0:##0.#0}", lstotal));
            dsrcmd.Parameters.AddWithValue("@dinnersales", string.Format("{0:##0.#0}", dstotal));
            dsrcmd.Parameters.AddWithValue("@tosales", string.Format("{0:##0.#0}", tototal));
            dsrcmd.Parameters.AddWithValue("@tocnt", tocnt);
            dsrcmd.Parameters.AddWithValue("@ccardtc", ccardcnt);
            dsrcmd.Parameters.AddWithValue("@lunchcc", lscnt);
            dsrcmd.Parameters.AddWithValue("@dailycc", custcnt);
            dsrcmd.Parameters.AddWithValue("@first", firstcustcnt);
            dsrcmd.Parameters.AddWithValue("@been", beencustcnt);
            dsrcmd.Parameters.AddWithValue("@dailyppa", string.Format("{0:##0.#0}", ddailyppa));
            dsrcmd.Parameters.AddWithValue("@tc", tcnt);
            dsrcmd.Parameters.AddWithValue("@dailypta", string.Format("{0:##0.#0}", ddailypta));
            dsrcmd.Parameters.AddWithValue("@dailytopta", string.Format("{0:##0.#0}", ddailytopta));
            dsrcmd.Parameters.AddWithValue("@dailyccardpta", string.Format("{0:##0.#0}", ddailyccardpta));
            dsrcmd.Parameters.AddWithValue("@averagecheck", string.Format("{0:##0.#0}", daveragecheck));
            dsrcmd.Parameters.AddWithValue("@grosssales", string.Format("{0:##0.#0}", grosssale));
            dsrcmd.Parameters.AddWithValue("@vatsales", string.Format("{0:##0.#0}", vatables));
            dsrcmd.Parameters.AddWithValue("@vatexempt", string.Format("{0:##0.#0}", vatexsales));
            dsrcmd.Parameters.AddWithValue("@zerorated", string.Format("{0:##0.#0}", vatzerosales));
            dsrcmd.Parameters.AddWithValue("@outputvat", string.Format("{0:##0.#0}", outputvats));
            dsrcmd.Parameters.AddWithValue("@scd", string.Format("{0:##0.#0}", dscd));
            dsrcmd.Parameters.AddWithValue("@pwd", string.Format("{0:##0.#0}", dpwd));
            dsrcmd.Parameters.AddWithValue("@mktgdisc", string.Format("{0:##0.#0}", dmktgdisc));
            dsrcmd.Parameters.AddWithValue("@tenpcntdisc", string.Format("{0:##0.#0}", dtenpcntdisc));
            dsrcmd.Parameters.AddWithValue("@cos", string.Format("{0:##0.#0}", cashofsales));
            dsrcmd.Parameters.AddWithValue("@ccardsales", string.Format("{0:##0.#0}", ccardsales));
            dsrcmd.Parameters.AddWithValue("@mktggcsales", string.Format("{0:##0.#0}", mktgamt));
            dsrcmd.Parameters.AddWithValue("@gctwo", string.Format("{0:##0.#0}", gcpmt));
            dsrcmd.Parameters.AddWithValue("@gcash", string.Format("{0:##0.#0}", gcashpmt));
            dsrcmd.Parameters.AddWithValue("@gcashcnt", gcashpmtcnt);
            dsrcmd.Parameters.AddWithValue("@hb", string.Format("{0:##0.#0}", hbrcsales));
            dsrcmd.Parameters.AddWithValue("@hbtc", hbtcsales);
            dsrcmd.Parameters.AddWithValue("@fp", string.Format("{0:##0.#0}", fprcsales));
            dsrcmd.Parameters.AddWithValue("@fptc", fptcsales);
            dsrcmd.Parameters.AddWithValue("@mangan", string.Format("{0:##0.#0}", manganrcsales));
            dsrcmd.Parameters.AddWithValue("@mangantc", mangantcsales);
            dsrcmd.Parameters.AddWithValue("@gf", string.Format("{0:##0.#0}", gfrcsales));
            dsrcmd.Parameters.AddWithValue("@gftc", gftcsales);
            dsrcmd.Parameters.AddWithValue("@zomato", string.Format("{0:##0.#0}", zomrcsales));
            dsrcmd.Parameters.AddWithValue("@zomatotc", zomtcsales);
            dsrcmd.Parameters.AddWithValue("@llf", string.Format("{0:##0.#0}", llfrcsales));
            dsrcmd.Parameters.AddWithValue("@llftc", llftcsales);
            dsrcmd.Parameters.AddWithValue("@tips", string.Format("{0:##0.#0}", dtips));
            dsrcmd.Parameters.AddWithValue("@gp", string.Format("{0:##0.#0}", grabpaypmt));
            dsrcmd.Parameters.AddWithValue("@gptc", grabpaytcpmt);
            dsrcmd.Parameters.AddWithValue("@gpartc", string.Format("{0:##0.#0}", grabpayarpmt));
            dsrcmd.Parameters.AddWithValue("@gptc", grabpayartcpmt);
            dsrcmd.Parameters.AddWithValue("@gcashar", string.Format("{0:##0.#0}", gcasharpmt));
            dsrcmd.Parameters.AddWithValue("@gcashartc", gcashartcpmt);
            dsrcmd.Parameters.AddWithValue("@wechat", string.Format("{0:##0.#0}", wechatpmt));
            dsrcmd.Parameters.AddWithValue("@wechattc", wechattcpmt);
            dsrcmd.Parameters.AddWithValue("@wechatar", string.Format("{0:##0.#0}", wechatarpmt));
            dsrcmd.Parameters.AddWithValue("@wechatartc", wechatartcpmt);
            dsrcmd.Parameters.AddWithValue("@fpar", string.Format("{0:##0.#0}", fparpmt));
            dsrcmd.Parameters.AddWithValue("@fpartc", fpartcpmt);
            dsrcmd.Parameters.AddWithValue("@llfar", string.Format("{0:##0.#0}", llfarpmt));
            dsrcmd.Parameters.AddWithValue("@llfartc", llfartcpmt);
            dsrcmd.Parameters.AddWithValue("@gfar", string.Format("{0:##0.#0}", gfarpmt));
            dsrcmd.Parameters.AddWithValue("@gfartc", gfartcpmt);
            dsrcmd.Parameters.AddWithValue("@zomar", string.Format("{0:##0.#0}", zomarpmt));
            dsrcmd.Parameters.AddWithValue("@zomartc", zomartcpmt);
            dsrcmd.Parameters.AddWithValue("@sfAR", string.Format("{0:##0.#0}", sfarpmt));
            dsrcmd.Parameters.AddWithValue("@sfARTC", sfartcpmt);
            dsrcmd.Parameters.AddWithValue("@mgc", string.Format("{0:##0.#0}", mgcpmt));
            dsrcmd.Parameters.AddWithValue("@mgcTC", mgctcpmt);
            dsrcmd.ExecuteNonQuery();
            closeConn();
        }
        #endregion

        //Return a Value if Value is null
        public Double ReturnData(string param)
        {
            if (param != "")
            {
                return Convert.ToDouble(param);
            }
            else
            {
                return 0;
            }
        }
    }
}
