using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using ClosedXML.Excel;

namespace CWA
{
    public partial class Form11 : Form
    {
        public Form11()
        {
            InitializeComponent();
        }

        public Form1 form1;


        public Form11(Form form)
        {
            InitializeComponent();
            form1 = (Form1)form;
        }



        public string FName = "";

        public Excel.Application xlApp { get; set; }
        public Excel.Workbook xlWorkBook { get; set; }
        public Excel.Worksheet Sheet { get; set; }

        public string Technology = "2G";



        public string File_Name = "";


        public DataTable Table_2G = new DataTable();
        public DataTable Table_3G = new DataTable();
        public DataTable Table_4G = new DataTable();



        public XLWorkbook Source_workbook = new XLWorkbook();
        public IXLWorksheet Source_worksheet = null;



        private void Form11_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {


            openFileDialog1.DefaultExt = "xlsx";
            openFileDialog1.Filter = "Excel File|*.xlsx";
            DialogResult result = openFileDialog1.ShowDialog();
            File_Name = openFileDialog1.SafeFileName.ToString();




            if (result == DialogResult.OK)
            {
                string file = openFileDialog1.FileName;
                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(file);
                Sheet = xlWorkBook.Worksheets[1];
                int RowCounts = Sheet.UsedRange.Rows.Count;

                if (Technology=="2G")
                {

                    Table_2G.Columns.Add("Cell", typeof(string));
                    Table_2G.Columns.Add("Date", typeof(DateTime));
                    Table_2G.Columns.Add("2G_Voice_Call_Drop_Rate>2", typeof(double));
                    Table_2G.Columns.Add("2G_Voice_Call_Setup_Success_Rate<94", typeof(double));
                    Table_2G.Columns.Add("IHSR<94", typeof(double));
                    Table_2G.Columns.Add("OHSR<94", typeof(double));
                    Table_2G.Columns.Add("RX_QUAL_DL<94", typeof(double));
                    Table_2G.Columns.Add("RX_QUAL_UL<94", typeof(double));
                    Table_2G.Columns.Add("SDCCH_Access_Success_Rate<94", typeof(double));
                    Table_2G.Columns.Add("SDCCH_Congestion_Rate>2", typeof(double));
                    Table_2G.Columns.Add("SDCCH_Drop_Rate>2", typeof(double));
                    Table_2G.Columns.Add("SDCCH_Traffic=0", typeof(double));
                    Table_2G.Columns.Add("TCH_Assignment_FR>2", typeof(double));
                    Table_2G.Columns.Add("TCH_Availability<99", typeof(double));
                    Table_2G.Columns.Add("TCH_Cong>2", typeof(double));
                    Table_2G.Columns.Add("TCH_Traffic (Erlang)=0", typeof(double));
                }


                if (Technology == "3G")
                {

                    Table_3G.Columns.Add("Cell", typeof(string));
                    Table_3G.Columns.Add("Date", typeof(DateTime));
                    Table_3G.Columns.Add("3G_Voice_Traffic (Erlang)=0", typeof(double));
                    Table_3G.Columns.Add("3G_Voice_Call_Drop_Rate>2", typeof(double));
                    Table_3G.Columns.Add("HSDPA_THR (Kbps)<200 Kbps", typeof(double));
                    Table_3G.Columns.Add("CS_IRAT_HO_SR<80", typeof(double));
                    Table_3G.Columns.Add("CS_RAB_SR<96", typeof(double));
                    Table_3G.Columns.Add("RX_RRC_SR<96", typeof(double));
                    Table_3G.Columns.Add("HSUPA_THR (Kbps)<30 Kbps", typeof(double));
                    Table_3G.Columns.Add("RTWP (dBm)>-80", typeof(double));
                    Table_3G.Columns.Add("PS_Drop_Rate>5", typeof(double));
                    Table_3G.Columns.Add("PS_RAB_SR<96", typeof(double));
                    Table_3G.Columns.Add("3G_Payload (GB)=0", typeof(double));
                    Table_3G.Columns.Add("Availability<99", typeof(double));
                    Table_3G.Columns.Add("Soft_HO_SR<94", typeof(double));
                }

                if (Technology == "4G")
                {
                    Table_4G.Columns.Add("Cell", typeof(string));
                    Table_4G.Columns.Add("Date", typeof(DateTime));
                    Table_4G.Columns.Add("DL_Latency (ms)>1000", typeof(double));
                    Table_4G.Columns.Add("DL_THR (Mbps)<4", typeof(double));
                    Table_4G.Columns.Add("Avg_Num_of_Users>300", typeof(double));
                    Table_4G.Columns.Add("UL_THR (Mbps)<0.5", typeof(double));
                    Table_4G.Columns.Add("Availability<99", typeof(double));
                    Table_4G.Columns.Add("CSFB_SR<96", typeof(double));
                    Table_4G.Columns.Add("DL_PRB_Utilization>99", typeof(double));
                    Table_4G.Columns.Add("ERAB_Drop_Rate>2", typeof(double));
                    Table_4G.Columns.Add("ERAB_SR<97", typeof(double));
                    Table_4G.Columns.Add("InterFreq_HO_SR<92", typeof(double));
                    Table_4G.Columns.Add("IntraFreq_HO_SR<92=0", typeof(double));
                    Table_4G.Columns.Add("RRC_Connection_SR<97", typeof(double));
                    Table_4G.Columns.Add("S1_Signal_SR<97", typeof(double));
                    Table_4G.Columns.Add("4G_Payload=0", typeof(double));
                }


                if (Technology == "2G")
                {
                    Excel.Range KPI_Val = Sheet.get_Range("A2", "P" + Sheet.UsedRange.Rows.Count);
                    object[,] KPI_Vals = (object[,])KPI_Val.Value2;


                    for (int k = 0; k < RowCounts - 1; k++)
                    {

                        double d = Convert.ToDouble(KPI_Vals[k + 1, 1]);
                        DateTime d1 = DateTime.FromOADate(d);

                        string Cell = Convert.ToString(KPI_Vals[k + 1, 2]);


                        string CDR_2G = "";
                        double CDR_2G_Issue = -1000;
                        if (Convert.ToString(KPI_Vals[k + 1, 3]) != "")
                        {
                            CDR_2G = Convert.ToString(KPI_Vals[k + 1, 3]);
                            if (Convert.ToDouble(CDR_2G) > 2)
                            {
                                CDR_2G_Issue = Convert.ToDouble(CDR_2G);
                            }
                        }

                        string CSSR_2G = "";
                        double CSSR_2G_Issue = -1000;
                        if (Convert.ToString(KPI_Vals[k + 1, 4]) != "")
                        {
                            CSSR_2G = Convert.ToString(KPI_Vals[k + 1, 4]);
                            if (Convert.ToDouble(CSSR_2G) < 94)
                            {
                                CSSR_2G_Issue = Convert.ToDouble(CSSR_2G);
                            }
                        }

                        string IHSR_2G = "";
                        double IHSR_2G_Issue = -1000;
                        if (Convert.ToString(KPI_Vals[k + 1, 5]) != "")
                        {
                            IHSR_2G = Convert.ToString(KPI_Vals[k + 1, 5]);
                            if (Convert.ToDouble(IHSR_2G) < 94)
                            {
                                IHSR_2G_Issue = Convert.ToDouble(IHSR_2G);
                            }
                        }

                        string OHSR_2G = "";
                        double OHSR_2G_Issue = -1000;
                        if (Convert.ToString(KPI_Vals[k + 1, 6]) != "")
                        {
                            OHSR_2G = Convert.ToString(KPI_Vals[k + 1, 6]);
                            if (Convert.ToDouble(OHSR_2G) < 94)
                            {
                                OHSR_2G_Issue = Convert.ToDouble(OHSR_2G);
                            }
                        }

                        string RX_QUAL_DL_2G = "";
                        double RX_QUAL_DL_2G_Issue = -1000;
                        if (Convert.ToString(KPI_Vals[k + 1, 7]) != "")
                        {
                            RX_QUAL_DL_2G = Convert.ToString(KPI_Vals[k + 1, 7]);
                            if (Convert.ToDouble(RX_QUAL_DL_2G) < 94)
                            {
                                RX_QUAL_DL_2G_Issue = Convert.ToDouble(RX_QUAL_DL_2G);
                            }
                        }

                        string RX_QUAL_UL_2G = "";
                        double RX_QUAL_UL_2G_Issue = -1000;
                        if (Convert.ToString(KPI_Vals[k + 1, 8]) != "")
                        {
                            RX_QUAL_UL_2G = Convert.ToString(KPI_Vals[k + 1, 8]);
                            if (Convert.ToDouble(RX_QUAL_UL_2G) < 94)
                            {
                                RX_QUAL_UL_2G_Issue = Convert.ToDouble(RX_QUAL_UL_2G);
                            }
                        }

                        string SDCCH_SR_2G = "";
                        double SDCCH_SR_2G_Issue = -1000;
                        if (Convert.ToString(KPI_Vals[k + 1, 9]) != "")
                        {
                            SDCCH_SR_2G = Convert.ToString(KPI_Vals[k + 1, 9]);
                            if (Convert.ToDouble(SDCCH_SR_2G) < 94)
                            {
                                SDCCH_SR_2G_Issue = Convert.ToDouble(SDCCH_SR_2G);
                            }
                        }

                        string SDCCH_Cong_2G = "";
                        double SDCCH_Cong_2G_Issue = -1000;
                        if (Convert.ToString(KPI_Vals[k + 1, 10]) != "")
                        {
                            SDCCH_Cong_2G = Convert.ToString(KPI_Vals[k + 1, 10]);
                            if (Convert.ToDouble(SDCCH_Cong_2G) > 2)
                            {
                                SDCCH_Cong_2G_Issue = Convert.ToDouble(SDCCH_Cong_2G);
                            }
                        }

                        string SDCCH_Drop_2G = "";
                        double SDCCH_Drop_2G_Issue = -1000;
                        if (Convert.ToString(KPI_Vals[k + 1, 11]) != "")
                        {
                            SDCCH_Drop_2G = Convert.ToString(KPI_Vals[k + 1, 11]);
                            if (Convert.ToDouble(SDCCH_Drop_2G) > 2)
                            {
                                SDCCH_Drop_2G_Issue = Convert.ToDouble(SDCCH_Drop_2G);
                            }
                        }

                        string SDCCH_Traffic_2G = "";
                        double SDCCH_Traffic_2G_Issue = -1000;
                        if (Convert.ToString(KPI_Vals[k + 1, 12]) != "")
                        {
                            SDCCH_Traffic_2G = Convert.ToString(KPI_Vals[k + 1, 12]);
                            if (Convert.ToDouble(SDCCH_Traffic_2G) == 0)
                            {
                                SDCCH_Traffic_2G_Issue = Convert.ToDouble(SDCCH_Traffic_2G);
                            }
                        }

                        string TCH_ASFR_2G = "";
                        double TCH_ASFR_2G_Issue = -1000;
                        if (Convert.ToString(KPI_Vals[k + 1, 13]) != "")
                        {
                            TCH_ASFR_2G = Convert.ToString(KPI_Vals[k + 1, 13]);
                            if (Convert.ToDouble(TCH_ASFR_2G) > 2)
                            {
                                TCH_ASFR_2G_Issue = Convert.ToDouble(TCH_ASFR_2G);
                            }
                        }
                        string TCH_Availability_2G = "";
                        double TCH_Availability_2G_Issue = -1000;
                        if (Convert.ToString(KPI_Vals[k + 1, 14]) != "")
                        {
                            TCH_Availability_2G = Convert.ToString(KPI_Vals[k + 1, 14]);
                            if (Convert.ToDouble(TCH_Availability_2G) < 99)
                            {
                                TCH_Availability_2G_Issue = Convert.ToDouble(TCH_Availability_2G);
                            }
                        }

                        string TCH_Cong_2G = "";
                        double TCH_Cong_2G_Issue = -1000;
                        if (Convert.ToString(KPI_Vals[k + 1, 15]) != "")
                        {
                            TCH_Cong_2G = Convert.ToString(KPI_Vals[k + 1, 15]);
                            if (Convert.ToDouble(TCH_Cong_2G) > 2)
                            {
                                TCH_Cong_2G_Issue = Convert.ToDouble(TCH_Cong_2G);
                            }
                        }

                        string TCH_Traffic_2G = "";
                        double TCH_Traffic_2G_Issue = -1000;
                        if (Convert.ToString(KPI_Vals[k + 1, 16]) != "")
                        {
                            TCH_Traffic_2G = Convert.ToString(KPI_Vals[k + 1, 16]);
                            if (Convert.ToDouble(TCH_Traffic_2G) == 0)
                            {
                                TCH_Traffic_2G_Issue = Convert.ToDouble(TCH_Traffic_2G);
                            }
                        }


                        if (CDR_2G_Issue != -1000 || CSSR_2G_Issue != -1000 || IHSR_2G_Issue != -1000 || OHSR_2G_Issue != -1000 || RX_QUAL_DL_2G_Issue != -1000 || RX_QUAL_UL_2G_Issue != -1000 || SDCCH_SR_2G_Issue != -1000 || SDCCH_Cong_2G_Issue != -1000 || SDCCH_Drop_2G_Issue != -1000 || SDCCH_Traffic_2G_Issue != -1000 || TCH_ASFR_2G_Issue != -1000 || TCH_Availability_2G_Issue != -1000 || TCH_Cong_2G_Issue != -1000 || TCH_Traffic_2G_Issue != -1000)
                        {
                            Table_2G.Rows.Add(Cell, d1, CDR_2G_Issue, CSSR_2G_Issue, IHSR_2G_Issue, OHSR_2G_Issue, RX_QUAL_DL_2G_Issue, RX_QUAL_UL_2G_Issue, SDCCH_SR_2G_Issue, SDCCH_Cong_2G_Issue, SDCCH_Drop_2G_Issue, SDCCH_Traffic_2G_Issue, TCH_ASFR_2G_Issue, TCH_Availability_2G_Issue, TCH_Cong_2G_Issue, TCH_Traffic_2G_Issue);

                        }


                    }
                }






                if (Technology == "3G")
                {
                    Excel.Range KPI_Val = Sheet.get_Range("A2", "O" + Sheet.UsedRange.Rows.Count);
                    object[,] KPI_Vals = (object[,])KPI_Val.Value2;


                    for (int k = 0; k < RowCounts - 1; k++)
                    {

                        double d = Convert.ToDouble(KPI_Vals[k + 1, 1]);
                        DateTime d1 = DateTime.FromOADate(d);


                            string Cell = Convert.ToString(KPI_Vals[k + 1, 2]);


                            string Traffic_3G = "";
                            double Traffic_3G_Issue = -1000;
                            if (Convert.ToString(KPI_Vals[k + 1, 3]) != "")
                            {
                                Traffic_3G = Convert.ToString(KPI_Vals[k + 1, 3]);
                                if (Convert.ToDouble(Traffic_3G) == 0)
                                {
                                    Traffic_3G_Issue = Convert.ToDouble(Traffic_3G);
                                }
                            }

                            string CDR_3G = "";
                        double CDR_3G_Issue = -1000;
                        if (Convert.ToString(KPI_Vals[k + 1, 4]) != "")
                            {
                                CDR_3G = Convert.ToString(KPI_Vals[k + 1, 4]);
                                if (Convert.ToDouble(CDR_3G) >2)
                                {
                                    CDR_3G_Issue = Convert.ToDouble(CDR_3G);
                                }
                            }

                            string HSDPA_THR_3G = "";
                        double HSDPA_THR_3G_Issue = -1000;
                            if (Convert.ToString(KPI_Vals[k + 1, 5]) != "")
                            {
                                HSDPA_THR_3G = Convert.ToString(KPI_Vals[k + 1, 5]);
                                if (Convert.ToDouble(HSDPA_THR_3G) < 200)
                                {
                                    HSDPA_THR_3G_Issue = Convert.ToDouble(HSDPA_THR_3G);
                                }
                            }

                            string CS_IRAT_3G = "";
                        double CS_IRAT_3G_Issue = -1000;
                            if (Convert.ToString(KPI_Vals[k + 1, 6]) != "")
                            {
                                CS_IRAT_3G = Convert.ToString(KPI_Vals[k + 1, 6]);
                                if (Convert.ToDouble(CS_IRAT_3G) < 80)
                                {
                                    CS_IRAT_3G_Issue = Convert.ToDouble(CS_IRAT_3G);
                                }
                            }

                            string CS_RAB_SR_3G = "";
                        double CS_RAB_SR_3G_Issue = -1000;
                            if (Convert.ToString(KPI_Vals[k + 1, 7]) != "")
                            {
                                CS_RAB_SR_3G = Convert.ToString(KPI_Vals[k + 1, 7]);
                                if (Convert.ToDouble(CS_RAB_SR_3G) < 96)
                                {
                                    CS_RAB_SR_3G_Issue = Convert.ToDouble(CS_RAB_SR_3G);
                                }
                            }

                            string CS_RRC_SR_3G = "";
                        double CS_RRC_SR_3G_Issue = -1000;
                            if (Convert.ToString(KPI_Vals[k + 1, 8]) != "")
                            {
                                CS_RRC_SR_3G = Convert.ToString(KPI_Vals[k + 1, 8]);
                                if (Convert.ToDouble(CS_RRC_SR_3G) < 96)
                                {
                                    CS_RRC_SR_3G_Issue = Convert.ToDouble(CS_RRC_SR_3G);
                                }
                            }

                            string HSUPA_THR_3G = "";
                        double HSUPA_THR_3G_Issue = -1000;
                            if (Convert.ToString(KPI_Vals[k + 1, 9]) != "")
                            {
                                HSUPA_THR_3G = Convert.ToString(KPI_Vals[k + 1, 9]);
                                if (Convert.ToDouble(HSUPA_THR_3G) < 30)
                                {
                                    HSUPA_THR_3G_Issue = Convert.ToDouble(HSUPA_THR_3G);
                                }
                            }

                            string RTWP_3G = "";
                        double RTWP_3G_Issue = -1000;
                            if (Convert.ToString(KPI_Vals[k + 1, 10]) != "")
                            {
                                RTWP_3G = Convert.ToString(KPI_Vals[k + 1, 10]);
                                if (Convert.ToDouble(RTWP_3G) > -80)
                                {
                                    RTWP_3G_Issue = Convert.ToDouble(RTWP_3G);
                                }
                            }

                            string PS_Drop_3G = "";
                        double PS_Drop_3G_Issue = -1000;
                            if (Convert.ToString(KPI_Vals[k + 1, 11]) != "")
                            {
                                PS_Drop_3G = Convert.ToString(KPI_Vals[k + 1, 11]);
                                if (Convert.ToDouble(PS_Drop_3G) > 5)
                                {
                                    PS_Drop_3G_Issue = Convert.ToDouble(PS_Drop_3G);
                                }
                            }

                            string PS_RAB_SR_3G = "";
                        double PS_RAB_SR_3G_Issue = -1000;
                            if (Convert.ToString(KPI_Vals[k + 1, 12]) != "")
                            {
                                PS_RAB_SR_3G = Convert.ToString(KPI_Vals[k + 1, 12]);
                                if (Convert.ToDouble(PS_RAB_SR_3G) <96)
                                {
                                    PS_RAB_SR_3G_Issue = Convert.ToDouble(PS_RAB_SR_3G);
                                }
                            }

                            string Payload_3G = "";
                        double Payload_3G_Issue = -1000;
                            if (Convert.ToString(KPI_Vals[k + 1, 13]) != "")
                            {
                                Payload_3G = Convert.ToString(KPI_Vals[k + 1, 13]);
                                if (Convert.ToDouble(Payload_3G) ==0)
                                {
                                    Payload_3G_Issue = Convert.ToDouble(Payload_3G);
                                }
                            }
                            string Availability_3G = "";
                        double Availability_3G_Issue = -1000;
                            if (Convert.ToString(KPI_Vals[k + 1, 14]) != "")
                            {
                                Availability_3G = Convert.ToString(KPI_Vals[k + 1, 14]);
                                if (Convert.ToDouble(Availability_3G) < 99)
                                {
                                    Availability_3G_Issue = Convert.ToDouble(Availability_3G);
                                }
                            }

                            string Soft_HO_3G = "";
                        double Soft_HO_3G_Issue = -1000;
                            if (Convert.ToString(KPI_Vals[k + 1, 15]) != "")
                            {
                                Soft_HO_3G = Convert.ToString(KPI_Vals[k + 1, 15]);
                                if (Convert.ToDouble(Soft_HO_3G) <94)
                                {
                                    Soft_HO_3G_Issue = Convert.ToDouble(Soft_HO_3G);
                                }
                            }



                        if (Traffic_3G_Issue != -1000 || CDR_3G_Issue != -1000 || HSDPA_THR_3G_Issue != -1000 || CS_IRAT_3G_Issue != -1000 || CS_RAB_SR_3G_Issue != -1000 || CS_RRC_SR_3G_Issue != -1000 || HSUPA_THR_3G_Issue != -1000 || RTWP_3G_Issue != -1000 || PS_Drop_3G_Issue != -1000 || PS_RAB_SR_3G_Issue != -1000 || Payload_3G_Issue != -1000 || Availability_3G_Issue != -1000 || Soft_HO_3G_Issue != -1000)
                        {
                            Table_3G.Rows.Add(Cell, d1, Traffic_3G_Issue, CDR_3G_Issue, HSDPA_THR_3G_Issue, CS_IRAT_3G_Issue, CS_RAB_SR_3G_Issue, CS_RRC_SR_3G_Issue, HSUPA_THR_3G_Issue, RTWP_3G_Issue, PS_Drop_3G_Issue, PS_RAB_SR_3G_Issue, Payload_3G_Issue, Availability_3G_Issue, Soft_HO_3G_Issue);

                        }


                    }
                }







                if (Technology == "4G")
                {
                    Excel.Range KPI_Val = Sheet.get_Range("A2", "P" + Sheet.UsedRange.Rows.Count);
                    object[,] KPI_Vals = (object[,])KPI_Val.Value2;


                    for (int k = 0; k < RowCounts - 1; k++)
                    {

                        double d = Convert.ToDouble(KPI_Vals[k + 1, 1]);
                        DateTime d1 = DateTime.FromOADate(d);


                        string Cell = Convert.ToString(KPI_Vals[k + 1, 2]);


                        string DL_Latency_4G = "";
                        double DL_Latency_4G_Issue = -1000;
                        if (Convert.ToString(KPI_Vals[k + 1, 3]) != "")
                        {
                            DL_Latency_4G = Convert.ToString(KPI_Vals[k + 1, 3]);
                            if (Convert.ToDouble(DL_Latency_4G) >1000)
                            {
                                DL_Latency_4G_Issue = Convert.ToDouble(DL_Latency_4G);
                            }
                        }

                        string DL_THR_4G = "";
                        double DL_THR_4G_Issue = -1000;
                        if (Convert.ToString(KPI_Vals[k + 1, 4]) != "")
                        {
                            DL_THR_4G = Convert.ToString(KPI_Vals[k + 1, 4]);
                            if (Convert.ToDouble(DL_THR_4G) <4)
                            {
                                DL_THR_4G_Issue = Convert.ToDouble(DL_THR_4G);
                            }
                        }

                        string Averge_Num_User_4G = "";
                        double Averge_Num_User_4G_Issue = -1000;
                        if (Convert.ToString(KPI_Vals[k + 1, 5]) != "")
                        {
                            Averge_Num_User_4G = Convert.ToString(KPI_Vals[k + 1, 5]);
                            if (Convert.ToDouble(Averge_Num_User_4G) >300)
                            {
                                Averge_Num_User_4G_Issue = Convert.ToDouble(Averge_Num_User_4G);
                            }
                        }

                        string UL_THR_4G = "";
                        double UL_THR_4G_Issue = -1000;
                        if (Convert.ToString(KPI_Vals[k + 1, 6]) != "")
                        {
                            UL_THR_4G = Convert.ToString(KPI_Vals[k + 1, 6]);
                            if (Convert.ToDouble(UL_THR_4G) <0.5)
                            {
                                UL_THR_4G_Issue = Convert.ToDouble(UL_THR_4G);
                            }
                        }

                        string Availability_4G = "";
                        double Availability_4G_Issue = -1000;
                        if (Convert.ToString(KPI_Vals[k + 1, 7]) != "")
                        {
                            Availability_4G = Convert.ToString(KPI_Vals[k + 1, 7]);
                            if (Convert.ToDouble(Availability_4G) < 99)
                            {
                                Availability_4G_Issue = Convert.ToDouble(Availability_4G);
                            }
                        }

                        string CSFB_4G = "";
                        double CSFB_4G_Issue = -1000;
                        if (Convert.ToString(KPI_Vals[k + 1, 8]) != "")
                        {
                            CSFB_4G = Convert.ToString(KPI_Vals[k + 1, 8]);
                            if (Convert.ToDouble(CSFB_4G) < 96)
                            {
                                CSFB_4G_Issue = Convert.ToDouble(CSFB_4G);
                            }
                        }

                        string DL_PRB_Utilization_4G = "";
                        double DL_PRB_Utilization_4G_Issue = -1000;
                        if (Convert.ToString(KPI_Vals[k + 1, 9]) != "")
                        {
                            DL_PRB_Utilization_4G = Convert.ToString(KPI_Vals[k + 1, 9]);
                            if (Convert.ToDouble(DL_PRB_Utilization_4G) >99)
                            {
                                DL_PRB_Utilization_4G_Issue = Convert.ToDouble(DL_PRB_Utilization_4G);
                            }
                        }

                        string ERAB_Drop_4G = "";
                        double ERAB_Drop_4G_Issue = -1000;
                        if (Convert.ToString(KPI_Vals[k + 1, 10]) != "")
                        {
                            ERAB_Drop_4G = Convert.ToString(KPI_Vals[k + 1, 10]);
                            if (Convert.ToDouble(ERAB_Drop_4G) > 2)
                            {
                                ERAB_Drop_4G_Issue = Convert.ToDouble(ERAB_Drop_4G);
                            }
                        }

                        string ERAB_SR_4G = "";
                        double ERAB_SR_4G_Issue = -1000;
                        if (Convert.ToString(KPI_Vals[k + 1, 11]) != "")
                        {
                            ERAB_SR_4G = Convert.ToString(KPI_Vals[k + 1, 11]);
                            if (Convert.ToDouble(ERAB_SR_4G) <97)
                            {
                                ERAB_SR_4G_Issue = Convert.ToDouble(ERAB_SR_4G);
                            }
                        }

                        string InterFreq_HO_SR_4G = "";
                        double InterFreq_HO_SR_4G_Issue = -1000;
                        if (Convert.ToString(KPI_Vals[k + 1, 12]) != "")
                        {
                            InterFreq_HO_SR_4G = Convert.ToString(KPI_Vals[k + 1, 12]);
                            if (Convert.ToDouble(InterFreq_HO_SR_4G) < 92)
                            {
                                InterFreq_HO_SR_4G_Issue = Convert.ToDouble(InterFreq_HO_SR_4G);
                            }
                        }

                        string IntraFreq_HO_SR_4G = "";
                        double IntraFreq_HO_SR_4G_Issue = -1000;
                        if (Convert.ToString(KPI_Vals[k + 1, 13]) != "")
                        {
                            IntraFreq_HO_SR_4G = Convert.ToString(KPI_Vals[k + 1, 13]);
                            if (Convert.ToDouble(IntraFreq_HO_SR_4G) <92)
                            {
                                IntraFreq_HO_SR_4G_Issue = Convert.ToDouble(IntraFreq_HO_SR_4G);
                            }
                        }
                        string RRC_Connection_SR_4G = "";
                        double RRC_Connection_SR_4G_Issue = -1000;
                        if (Convert.ToString(KPI_Vals[k + 1, 14]) != "")
                        {
                            RRC_Connection_SR_4G = Convert.ToString(KPI_Vals[k + 1, 14]);
                            if (Convert.ToDouble(RRC_Connection_SR_4G) < 97)
                            {
                                RRC_Connection_SR_4G_Issue = Convert.ToDouble(RRC_Connection_SR_4G);
                            }
                        }

                        string S1_Signal_SR_4G = "";
                        double S1_Signal_SR_4G_Issue = -1000;
                        if (Convert.ToString(KPI_Vals[k + 1, 15]) != "")
                        {
                            S1_Signal_SR_4G = Convert.ToString(KPI_Vals[k + 1, 15]);
                            if (Convert.ToDouble(S1_Signal_SR_4G) < 97)
                            {
                                S1_Signal_SR_4G_Issue = Convert.ToDouble(S1_Signal_SR_4G);
                            }
                        }


                        string Payload_4G = "";
                        double Payload_4G_Issue = -1000;
                        if (Convert.ToString(KPI_Vals[k + 1, 16]) != "")
                        {
                            Payload_4G = Convert.ToString(KPI_Vals[k + 1, 16]);
                            if (Convert.ToDouble(Payload_4G) ==0)
                            {
                                Payload_4G_Issue = Convert.ToDouble(Payload_4G);
                            }
                        }



                        if (DL_Latency_4G_Issue != -1000 || DL_THR_4G_Issue != -1000 || Averge_Num_User_4G_Issue != -1000 || UL_THR_4G_Issue != -1000 || Availability_4G_Issue != -1000 || CSFB_4G_Issue != -1000 || DL_PRB_Utilization_4G_Issue != -1000 || ERAB_Drop_4G_Issue != -1000 || ERAB_SR_4G_Issue != -1000 || InterFreq_HO_SR_4G_Issue != -1000 || IntraFreq_HO_SR_4G_Issue != -1000 || RRC_Connection_SR_4G_Issue != -1000 || S1_Signal_SR_4G_Issue != -1000 || Payload_4G_Issue!=-1000)
                        {
                            Table_4G.Rows.Add(Cell, d1, DL_Latency_4G_Issue, DL_THR_4G_Issue, Averge_Num_User_4G_Issue, UL_THR_4G_Issue, Availability_4G_Issue, CSFB_4G_Issue, DL_PRB_Utilization_4G_Issue, ERAB_Drop_4G_Issue, ERAB_SR_4G_Issue, InterFreq_HO_SR_4G_Issue, IntraFreq_HO_SR_4G_Issue, RRC_Connection_SR_4G_Issue, S1_Signal_SR_4G_Issue, Payload_4G_Issue);

                        }


                    }
                }






                if (Technology=="2G")
                {


                    XLWorkbook wb = new XLWorkbook();
                    wb.Worksheets.Add(Table_2G, "Arbaein_2G_Hourly_KPI_Issues");
                    var saveFileDialog = new SaveFileDialog
                    {
                        FileName = "Arbaein_2G_Hourly_KPI_Issues",
                        Filter = "Excel files|*.xlsx",
                        Title = "Save an Excel File"
                    };

                    saveFileDialog.ShowDialog();

                    if (!String.IsNullOrWhiteSpace(saveFileDialog.FileName))
                        wb.SaveAs(saveFileDialog.FileName);
                }

                if (Technology == "3G")
                {
                          
                    XLWorkbook wb = new XLWorkbook();
                    wb.Worksheets.Add(Table_3G, "Arbaein_3G_Hourly_KPI_Issues");
                    var saveFileDialog = new SaveFileDialog
                    {
                        FileName = "Arbaein_3G_Hourly_KPI_Issues",
                        Filter = "Excel files|*.xlsx",
                        Title = "Save an Excel File"
                    };

                    saveFileDialog.ShowDialog();

                    if (!String.IsNullOrWhiteSpace(saveFileDialog.FileName))
                        wb.SaveAs(saveFileDialog.FileName);
                }


                if (Technology == "4G")
                {

                    XLWorkbook wb = new XLWorkbook();
                    wb.Worksheets.Add(Table_4G, "Arbaein_4G_Hourly_KPI_Issues");
                    var saveFileDialog = new SaveFileDialog
                    {
                        FileName = "Arbaein_4G_Hourly_KPI_Issues",
                        Filter = "Excel files|*.xlsx",
                        Title = "Save an Excel File"
                    };

                    saveFileDialog.ShowDialog();

                    if (!String.IsNullOrWhiteSpace(saveFileDialog.FileName))
                        wb.SaveAs(saveFileDialog.FileName);
                }



                MessageBox.Show("Finished");


            }


        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                Technology = "2G";
                checkBox2.Checked = false;
                checkBox3.Checked = false;
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                Technology = "3G";
                checkBox1.Checked = false;
                checkBox3.Checked = false;
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked == true)
            {
                Technology = "4G";
                checkBox1.Checked = false;
                checkBox2.Checked = false;
            }
        }
    }


}
