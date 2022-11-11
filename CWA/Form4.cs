using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Windows.Forms.DataVisualization.Charting;

namespace CWA
{
    public partial class Form4 : Form
    {
        public Form4()
        {
            InitializeComponent();
        }


        public Form1 form1;


        public Form4(Form form)
        {
            InitializeComponent();
            form1 = (Form1)form;
        }


        string RNC_in_LTE = "";
        public double BL = -1000;
      


        private void Form4_Load(object sender, EventArgs e)
        {

            string Cell = form1.Selected_Cell;
            string KPI1= form1.Selected_KPI;
            DataTable KPI_Table_E = form1.KPI_Table_E1;
            DataTable KPI_Table_H = form1.KPI_Table_H1;
            DataTable KPI_Table_N = form1.KPI_Table_N1;

            DataTable KPI_Table_BL = form1.BL_Table;

            chart1.Series.Clear();
            chart1.Titles.Clear();
            Series newSeries1 = new Series();
            chart1.Series.Add(newSeries1);
            chart1.Series[0].ChartType = SeriesChartType.Line;
            chart1.Series[0].BorderWidth = 3;
            chart1.ChartAreas[0].AxisX.Interval = 5;
            chart1.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart1.Series[0].EmptyPointStyle.Color = Color.Transparent;
            chart1.Series[0].XValueType = ChartValueType.DateTime;
            chart1.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";
            chart1.Series[0].ToolTip = "#VALY{F}\n#VALX{dd/MM}";
            //Title title1 = chart1.Titles.Add(form1.Selected_Cell+"_"+form1.Selected_KPI);
            //title1.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            chart1.Series[0].IsValueShownAsLabel = false;


            Series newSeries2 = new Series();
            chart1.Series.Add(newSeries2);
            chart1.Series[1].ChartType = SeriesChartType.Line;
            chart1.Series[1].BorderWidth = 3;
            chart1.ChartAreas[0].AxisX.Interval = 5;
            chart1.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart1.Series[1].EmptyPointStyle.Color = Color.Transparent;
            chart1.Series[1].Color = Color.Green;
            chart1.Series[1].XValueType = ChartValueType.DateTime;
            chart1.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";
            chart1.Series[1].ToolTip = "#VALY{F}\n#VALX{dd/MM}";
            //Title title2 = chart1.Titles.Add(form1.Selected_Cell + "_" + form1.Selected_KPI);
            //title2.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            chart1.Series[1].IsValueShownAsLabel = false;

            Series newSeries3 = new Series();
            chart1.Series.Add(newSeries3);
            chart1.Series[2].ChartType = SeriesChartType.Line;
            chart1.Series[2].BorderWidth = 3;
            chart1.ChartAreas[0].AxisX.Interval = 5;
            chart1.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart1.Series[2].EmptyPointStyle.Color = Color.Transparent;
            chart1.Series[2].Color = Color.Red;
            chart1.Series[2].XValueType = ChartValueType.DateTime;
            chart1.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";
            chart1.Series[2].ToolTip = "#VALY{F}\n#VALX{dd/MM}";
            //Title title3 = chart1.Titles.Add(form1.Selected_Cell + "_" + form1.Selected_KPI);
            //title3.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            chart1.Series[2].IsValueShownAsLabel = false;


            Series newSeries4 = new Series();
            chart1.Series.Add(newSeries4);
            chart1.Series[3].ChartType = SeriesChartType.Line;
            chart1.Series[3].BorderWidth = 3;
            chart1.ChartAreas[0].AxisX.Interval = 5;
            chart1.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart1.Series[3].EmptyPointStyle.Color = Color.Transparent;
            chart1.Series[3].Color = Color.Black;
            chart1.Series[3].XValueType = ChartValueType.DateTime;
            chart1.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";
            chart1.Series[3].ToolTip = "#VALY{F}\n#VALX{dd/MM}";
            //Title title3 = chart1.Titles.Add(form1.Selected_Cell + "_" + form1.Selected_KPI);
            //title3.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            chart1.Series[3].IsValueShownAsLabel = false;
            chart1.Series[3].Enabled = false;


            if (form1.Technology == "2G_CS")
            {

                var cell_data = (from p in form1.Last_Day_List.AsEnumerable()
                                 where p.Field<string>("Cell") == Cell
                                 select p).ToList();

                var bl_data = (from p in KPI_Table_BL.AsEnumerable()
                                 where p.Field<string>("Cell") == Cell
                                 select p).ToList();

                dataGridView1.ColumnCount = 2;
                dataGridView1.Columns[0].Width = 150;
                dataGridView1.Columns[1].Width = 60;
                dataGridView1.RowCount = 22;
                DateTime d = Convert.ToDateTime(cell_data[0].ItemArray[0]);
                dataGridView1.Rows[0].Cells[0].Value = "Date"; dataGridView1.Rows[0].Cells[1].Value = Convert.ToString(d.Month) + "/" + Convert.ToString(d.Day) + "/" + Convert.ToString(d.Year);
                dataGridView1.Rows[1].Cells[0].Value = "Cell"; dataGridView1.Rows[1].Cells[1].Value = cell_data[0].ItemArray[1];
                dataGridView1.Rows[2].Cells[0].Value = "Level"; dataGridView1.Rows[2].Cells[1].Value = cell_data[0].ItemArray[5];
                dataGridView1.Rows[3].Cells[0].Value = "Worst"; dataGridView1.Rows[3].Cells[1].Value = cell_data[0].ItemArray[6];
                dataGridView1.Rows[4].Cells[0].Value = "Effeciency (%)"; dataGridView1.Rows[4].Cells[1].Value = cell_data[0].ItemArray[7];
                dataGridView1.Rows[5].Cells[0].Value = "Status"; dataGridView1.Rows[5].Cells[1].Value = cell_data[0].ItemArray[8];
                dataGridView1.Rows[6].Cells[0].Value = "QIX"; dataGridView1.Rows[6].Cells[1].Value = cell_data[0].ItemArray[9];
                dataGridView1.Rows[7].Cells[0].Value = "QIxP"; dataGridView1.Rows[7].Cells[1].Value = cell_data[0].ItemArray[10];
                dataGridView1.Rows[8].Cells[0].Value = "QIxBL"; dataGridView1.Rows[8].Cells[1].Value = cell_data[0].ItemArray[11];
                dataGridView1.Rows[9].Cells[0].Value = "Traffic (Erlang)"; dataGridView1.Rows[9].Cells[1].Value = cell_data[0].ItemArray[12];
                dataGridView1.Rows[10].Cells[0].Value = "Repeated Daily at 10 days(QIx<QIb)"; dataGridView1.Rows[10].Cells[1].Value = cell_data[0].ItemArray[13];
                dataGridView1.Rows[11].Cells[0].Value = "Total Days Availability>90"; dataGridView1.Rows[11].Cells[1].Value = cell_data[0].ItemArray[14];
                dataGridView1.Rows[12].Cells[0].Value = "CSSR (% of WPC)"; dataGridView1.Rows[12].Cells[1].Value = cell_data[0].ItemArray[15];
                dataGridView1.Rows[13].Cells[0].Value = "OHSR (% of WPC)"; dataGridView1.Rows[13].Cells[1].Value = cell_data[0].ItemArray[16];
                dataGridView1.Rows[14].Cells[0].Value = "CDR (% of WPC)"; dataGridView1.Rows[14].Cells[1].Value = cell_data[0].ItemArray[17];
                dataGridView1.Rows[15].Cells[0].Value = "TCH_ASFR (% of WPC)"; dataGridView1.Rows[15].Cells[1].Value = cell_data[0].ItemArray[18];
                dataGridView1.Rows[16].Cells[0].Value = "RXDL (% of WPC)"; dataGridView1.Rows[16].Cells[1].Value = cell_data[0].ItemArray[19];
                dataGridView1.Rows[17].Cells[0].Value = "RXUL (% of WPC)"; dataGridView1.Rows[17].Cells[1].Value = cell_data[0].ItemArray[20];
                dataGridView1.Rows[18].Cells[0].Value = "SDCCH_CONG (% of WPC)"; dataGridView1.Rows[18].Cells[1].Value = cell_data[0].ItemArray[21];
                dataGridView1.Rows[19].Cells[0].Value = "SDCCH_SR (% of WPC)"; dataGridView1.Rows[19].Cells[1].Value = cell_data[0].ItemArray[22];
                dataGridView1.Rows[20].Cells[0].Value = "SDCCH_DROP (% of WPC)"; dataGridView1.Rows[20].Cells[1].Value = cell_data[0].ItemArray[23];
                dataGridView1.Rows[21].Cells[0].Value = "IHSR (% of WPC)"; dataGridView1.Rows[21].Cells[1].Value = cell_data[0].ItemArray[24];
           
                int KPI_Index =0;
                int BL_Index = 0;
                if (KPI1=="CSSR")
                {
                    KPI_Index = 12;
                    BL_Index = 7;
                }
                if (KPI1 == "OHSR")
                {
                    KPI_Index = 13;
                    BL_Index = 12;
                }
                if (KPI1 == "CDR")
                {
                    KPI_Index = 14;
                    BL_Index = 14;
                }
                if (KPI1 == "TCH_ASFR")
                {
                    KPI_Index = 15;
                    BL_Index = 17;
                }
                if (KPI1 == "RXDL")
                {
                    KPI_Index = 16;
                    BL_Index = 15;
                }
                if (KPI1 == "RXUL")
                {
                    KPI_Index = 17;
                    BL_Index = 16;
                }
                if (KPI1 == "SDCCH_CONG")
                {
                    KPI_Index = 18;
                    BL_Index = 8;
                }
                if (KPI1 == "SDCCH_SR")
                {
                    KPI_Index = 19;
                    BL_Index = 13;
                }
                if (KPI1 == "SDCCH_DROP")
                {
                    KPI_Index = 20;
                    BL_Index = 9;
                }
                if (KPI1 == "IHSR")
                {
                    KPI_Index = 21;
                    BL_Index = 11;
                }
                if (KPI_Index != 0)
                {
                    dataGridView1.Rows[KPI_Index].Cells[0].Style.BackColor = Color.Yellow;
                    dataGridView1.Rows[KPI_Index].Cells[1].Style.BackColor = Color.Yellow;
                }

                if (KPI_Index != 0 && bl_data.Count != 0)
                {
                    BL = Convert.ToDouble(bl_data[0].ItemArray[BL_Index]);
                }

            }




            if (form1.Technology == "2G_PS")
            {

                var cell_data = (from p in form1.Last_Day_List.AsEnumerable()
                                 where p.Field<string>("Cell") == Cell
                                 select p).ToList();

                //var bl_data = (from p in KPI_Table_BL.AsEnumerable()
                //               where p.Field<string>("Cell") == Cell
                //               select p).ToList();

                dataGridView1.ColumnCount = 2;
                dataGridView1.Columns[0].Width = 150;
                dataGridView1.Columns[1].Width = 60;
                dataGridView1.RowCount = 18;
                DateTime d = Convert.ToDateTime(cell_data[0].ItemArray[0]);
                dataGridView1.Rows[0].Cells[0].Value = "Date"; dataGridView1.Rows[0].Cells[1].Value = Convert.ToString(d.Month) + "/" + Convert.ToString(d.Day) + "/" + Convert.ToString(d.Year);
                dataGridView1.Rows[1].Cells[0].Value = "Cell"; dataGridView1.Rows[1].Cells[1].Value = cell_data[0].ItemArray[1];
                dataGridView1.Rows[2].Cells[0].Value = "Level"; dataGridView1.Rows[2].Cells[1].Value = cell_data[0].ItemArray[5];
                dataGridView1.Rows[3].Cells[0].Value = "Worst"; dataGridView1.Rows[3].Cells[1].Value = cell_data[0].ItemArray[6];
                dataGridView1.Rows[4].Cells[0].Value = "Effeciency (%)"; dataGridView1.Rows[4].Cells[1].Value = cell_data[0].ItemArray[7];
                dataGridView1.Rows[5].Cells[0].Value = "Status"; dataGridView1.Rows[5].Cells[1].Value = cell_data[0].ItemArray[8];
                dataGridView1.Rows[6].Cells[0].Value = "QIX"; dataGridView1.Rows[6].Cells[1].Value = cell_data[0].ItemArray[9];
                dataGridView1.Rows[7].Cells[0].Value = "QIxP"; dataGridView1.Rows[7].Cells[1].Value = cell_data[0].ItemArray[10];
                dataGridView1.Rows[8].Cells[0].Value = "QIxBL"; dataGridView1.Rows[8].Cells[1].Value = cell_data[0].ItemArray[11];
                dataGridView1.Rows[9].Cells[0].Value = "Avg. Payload of Cell(GB)"; dataGridView1.Rows[9].Cells[1].Value = cell_data[0].ItemArray[12];
                dataGridView1.Rows[10].Cells[0].Value = "Repeated Daily at 10 days (QIx<QIb)"; dataGridView1.Rows[10].Cells[1].Value = cell_data[0].ItemArray[13];
                dataGridView1.Rows[11].Cells[0].Value = "Total Days Availability > 90"; dataGridView1.Rows[11].Cells[1].Value = cell_data[0].ItemArray[14];
                dataGridView1.Rows[12].Cells[0].Value = "Worst(%) of DL EGPRS Accessibility"; dataGridView1.Rows[12].Cells[1].Value = cell_data[0].ItemArray[15];
                dataGridView1.Rows[13].Cells[0].Value = "Worst(%) of DL TBF Drop Rate"; dataGridView1.Rows[13].Cells[1].Value = cell_data[0].ItemArray[16];
                dataGridView1.Rows[14].Cells[0].Value = "Worst(%) of DL GPRS Throughput (kbps)"; dataGridView1.Rows[14].Cells[1].Value = cell_data[0].ItemArray[17];
                dataGridView1.Rows[15].Cells[0].Value = "Worst(%) of DL EGPRS Throughput (kbps)"; dataGridView1.Rows[15].Cells[1].Value = cell_data[0].ItemArray[18];
                dataGridView1.Rows[16].Cells[0].Value = "Worst(%) of THR_DL_GPRS_PER_TS"; dataGridView1.Rows[16].Cells[1].Value = cell_data[0].ItemArray[19];
                dataGridView1.Rows[17].Cells[0].Value = "Worst(%) of THR_DL_EGPRS_PER_TS"; dataGridView1.Rows[17].Cells[1].Value = cell_data[0].ItemArray[20];


    


                int KPI_Index = 0;
                int BL_Index = 0;
                if (KPI1 == "TBF_Establish")
                {
                    KPI_Index = 12;
                   // BL_Index = 7;
                }
                if (KPI1 == "TBF_Drop")
                {
                    KPI_Index = 13;
                  //  BL_Index = 12;
                }
                if (KPI1 == "GPRS_THR")
                {
                    KPI_Index = 14;
               //     BL_Index = 14;
                }
                if (KPI1 == "EGPRS_THR")
                {
                    KPI_Index = 15;
                 //   BL_Index = 17;
                }
                if (KPI1 == "GPRS_THR_per_TS")
                {
                    KPI_Index = 16;
                   // BL_Index = 15;
                }
                if (KPI1 == "EGPRS_THR_per_TS")
                {
                    KPI_Index = 17;
                  //  BL_Index = 16;
                }


                if (KPI_Index != 0)
                {
                    dataGridView1.Rows[KPI_Index].Cells[0].Style.BackColor = Color.Yellow;
                    dataGridView1.Rows[KPI_Index].Cells[1].Style.BackColor = Color.Yellow;
                }

                //if (KPI_Index != 0 && bl_data.Count != 0)
                //{
                //    BL = Convert.ToDouble(bl_data[0].ItemArray[BL_Index]);
                //}

            }







            if (form1.Technology == "3G_CS")
            {

                var cell_data = (from p in form1.Last_Day_List.AsEnumerable()
                                 where p.Field<string>("Cell") == Cell
                                 select p).ToList();

                var bl_data = (from p in KPI_Table_BL.AsEnumerable()
                               where p.Field<string>("Cell") == Cell
                               select p).ToList();


                dataGridView1.ColumnCount = 2;
                dataGridView1.Columns[0].Width = 150;
                dataGridView1.Columns[1].Width = 60;
                dataGridView1.RowCount = 17;
                DateTime d = Convert.ToDateTime(cell_data[0].ItemArray[0]);
                dataGridView1.Rows[0].Cells[0].Value = "Date"; dataGridView1.Rows[0].Cells[1].Value = Convert.ToString(d.Month) + "/" + Convert.ToString(d.Day) + "/" + Convert.ToString(d.Year);
                dataGridView1.Rows[1].Cells[0].Value = "Cell"; dataGridView1.Rows[1].Cells[1].Value = cell_data[0].ItemArray[1];
                dataGridView1.Rows[2].Cells[0].Value = "LEVEL"; dataGridView1.Rows[2].Cells[1].Value = cell_data[0].ItemArray[5];
                dataGridView1.Rows[3].Cells[0].Value = "Worst"; dataGridView1.Rows[3].Cells[1].Value = cell_data[0].ItemArray[6];
                dataGridView1.Rows[4].Cells[0].Value = "Effeciency (%)"; dataGridView1.Rows[4].Cells[1].Value = cell_data[0].ItemArray[7];
                dataGridView1.Rows[5].Cells[0].Value = "Status"; dataGridView1.Rows[5].Cells[1].Value = cell_data[0].ItemArray[8];
                dataGridView1.Rows[6].Cells[0].Value = "QIx"; dataGridView1.Rows[6].Cells[1].Value = cell_data[0].ItemArray[9];
                dataGridView1.Rows[7].Cells[0].Value = "QIxP"; dataGridView1.Rows[7].Cells[1].Value = cell_data[0].ItemArray[10];
                dataGridView1.Rows[8].Cells[0].Value = "QIxBL"; dataGridView1.Rows[8].Cells[1].Value = cell_data[0].ItemArray[11];
                dataGridView1.Rows[9].Cells[0].Value = "Traffic (Erlang)"; dataGridView1.Rows[9].Cells[1].Value = cell_data[0].ItemArray[12];
                dataGridView1.Rows[10].Cells[0].Value = "Repeated Daily at 10 days (QIx<QIb)"; dataGridView1.Rows[10].Cells[1].Value = cell_data[0].ItemArray[13];
                dataGridView1.Rows[11].Cells[0].Value = "Total Days Availability > 90"; dataGridView1.Rows[11].Cells[1].Value = cell_data[0].ItemArray[14];
                dataGridView1.Rows[12].Cells[0].Value = "CS_RAB_Establish (% of WPC)"; dataGridView1.Rows[12].Cells[1].Value = cell_data[0].ItemArray[15];
                dataGridView1.Rows[13].Cells[0].Value = "CS_IRAT_HO_SR (% of WPC)"; dataGridView1.Rows[13].Cells[1].Value = cell_data[0].ItemArray[16];
                dataGridView1.Rows[14].Cells[0].Value = "CS_Drop_Rate (% of WPC)"; dataGridView1.Rows[14].Cells[1].Value = cell_data[0].ItemArray[17];
                dataGridView1.Rows[15].Cells[0].Value = "Soft_HO_SR (% of WPC)"; dataGridView1.Rows[15].Cells[1].Value = cell_data[0].ItemArray[18];
                dataGridView1.Rows[16].Cells[0].Value = "CS_RRC_SR (% of WPC)"; dataGridView1.Rows[16].Cells[1].Value = cell_data[0].ItemArray[19];

                int KPI_Index = 0;
                int BL_Index = 0;
                if (KPI1 == "CS_RAB_Establish")
                {
                    KPI_Index = 12;
                    BL_Index = 3;
                }
                if (KPI1 == "CS_IRAT_HO_SR")
                {
                    KPI_Index = 13;
                    BL_Index = 4;
                }
                if (KPI1 == "CS_Drop_Rate")
                {
                    KPI_Index = 14;
                    BL_Index = 5;
                }
                if (KPI1 == "Soft_HO_SR")
                {
                    KPI_Index = 15;
                    BL_Index = 6;
                }
                if (KPI1 == "CS_RRC_SR")
                {
                    KPI_Index = 16;
                    BL_Index = 7;
                }
               
                if (KPI_Index!=0)
                {
                    dataGridView1.Rows[KPI_Index].Cells[0].Style.BackColor = Color.Yellow;
                    dataGridView1.Rows[KPI_Index].Cells[1].Style.BackColor = Color.Yellow;
                }

                if (KPI_Index!=0 && bl_data.Count != 0)
                {
                    BL = Convert.ToDouble(bl_data[0].ItemArray[BL_Index]);
                }

            }




            if (form1.Technology == "3G_PS")
            {

                var cell_data = (from p in form1.Last_Day_List.AsEnumerable()
                                 where p.Field<string>("Cell") == Cell
                                 select p).ToList();

                var bl_data = (from p in KPI_Table_BL.AsEnumerable()
                               where p.Field<string>("Cell") == Cell
                               select p).ToList();


                dataGridView1.ColumnCount = 2;
                dataGridView1.Columns[0].Width = 150;
                dataGridView1.Columns[1].Width = 60;
                dataGridView1.RowCount = 26;
                DateTime d = Convert.ToDateTime(cell_data[0].ItemArray[0]);
                dataGridView1.Rows[0].Cells[0].Value = "Date"; dataGridView1.Rows[0].Cells[1].Value = Convert.ToString(d.Month) + "/" + Convert.ToString(d.Day) + "/" + Convert.ToString(d.Year);
                dataGridView1.Rows[1].Cells[0].Value = "Cell"; dataGridView1.Rows[1].Cells[1].Value = cell_data[0].ItemArray[1];
                dataGridView1.Rows[2].Cells[0].Value = "LEVEL"; dataGridView1.Rows[2].Cells[1].Value = cell_data[0].ItemArray[5];
                dataGridView1.Rows[3].Cells[0].Value = "Worst"; dataGridView1.Rows[3].Cells[1].Value = cell_data[0].ItemArray[6];
                dataGridView1.Rows[4].Cells[0].Value = "Effeciency (%)"; dataGridView1.Rows[4].Cells[1].Value = cell_data[0].ItemArray[7];
                dataGridView1.Rows[5].Cells[0].Value = "Status"; dataGridView1.Rows[5].Cells[1].Value = cell_data[0].ItemArray[8];
                dataGridView1.Rows[6].Cells[0].Value = "QIx"; dataGridView1.Rows[6].Cells[1].Value = cell_data[0].ItemArray[9];
                dataGridView1.Rows[7].Cells[0].Value = "QIxP"; dataGridView1.Rows[7].Cells[1].Value = cell_data[0].ItemArray[10];
                dataGridView1.Rows[8].Cells[0].Value = "QIxBL"; dataGridView1.Rows[8].Cells[1].Value = cell_data[0].ItemArray[11];
                dataGridView1.Rows[9].Cells[0].Value = "Payload (GB)"; dataGridView1.Rows[9].Cells[1].Value = cell_data[0].ItemArray[12];
                dataGridView1.Rows[10].Cells[0].Value = "Repeated Daily at 10 days (QIx<QIb)"; dataGridView1.Rows[10].Cells[1].Value = cell_data[0].ItemArray[13];
                dataGridView1.Rows[11].Cells[0].Value = "Total Days Availability > 90"; dataGridView1.Rows[11].Cells[1].Value = cell_data[0].ItemArray[14];
                dataGridView1.Rows[12].Cells[0].Value = "HSDPA_SR (% of WPC)"; dataGridView1.Rows[12].Cells[1].Value = cell_data[0].ItemArray[15];
                dataGridView1.Rows[13].Cells[0].Value = "HSUPA_SR (% of WPC)"; dataGridView1.Rows[13].Cells[1].Value = cell_data[0].ItemArray[16];
                dataGridView1.Rows[14].Cells[0].Value = "UL_User_THR (% of WPC)"; dataGridView1.Rows[14].Cells[1].Value = cell_data[0].ItemArray[17];
                dataGridView1.Rows[15].Cells[0].Value = "DL_User_THR (% of WPC)"; dataGridView1.Rows[15].Cells[1].Value = cell_data[0].ItemArray[18];
                dataGridView1.Rows[16].Cells[0].Value = "HSDAP_Drop_Rate (% of WPC)"; dataGridView1.Rows[16].Cells[1].Value = cell_data[0].ItemArray[19];
                dataGridView1.Rows[17].Cells[0].Value = "HSUPA_Drop_Rate (% of WPC)"; dataGridView1.Rows[17].Cells[1].Value = cell_data[0].ItemArray[20];
                dataGridView1.Rows[18].Cells[0].Value = "MultiRAB_SR (% of WPC)"; dataGridView1.Rows[18].Cells[1].Value = cell_data[0].ItemArray[21];
                dataGridView1.Rows[19].Cells[0].Value = "PS_RRC_SR (% of WPC)"; dataGridView1.Rows[19].Cells[1].Value = cell_data[0].ItemArray[22];
                dataGridView1.Rows[20].Cells[0].Value = "Ps_RAB_Establish (% of WPC)"; dataGridView1.Rows[20].Cells[1].Value = cell_data[0].ItemArray[23];
                dataGridView1.Rows[21].Cells[0].Value = "PS_MultiRAB_Establish (% of WPC)"; dataGridView1.Rows[21].Cells[1].Value = cell_data[0].ItemArray[24];
                dataGridView1.Rows[22].Cells[0].Value = "PS_Drop_Rate (% of WPC)"; dataGridView1.Rows[22].Cells[1].Value = cell_data[0].ItemArray[25];
                dataGridView1.Rows[23].Cells[0].Value = "HSDPA_Cell_Change_SR (% of WPC)"; dataGridView1.Rows[23].Cells[1].Value = cell_data[0].ItemArray[26];
                dataGridView1.Rows[24].Cells[0].Value = "HS_Share_Payload (% of WPC)"; dataGridView1.Rows[24].Cells[1].Value = cell_data[0].ItemArray[27];
                dataGridView1.Rows[25].Cells[0].Value = "DL_Cell_THR (% of WPC)"; dataGridView1.Rows[25].Cells[1].Value = cell_data[0].ItemArray[28];

                int KPI_Index = 0;
                int BL_Index = 0;
                if (KPI1 == "HSDPA_SR")
                {
                    KPI_Index = 12;
                    BL_Index = 4;
                }
                if (KPI1 == "HSUPA_SR")
                {
                    KPI_Index = 13;
                    BL_Index = 5;
                }
                if (KPI1 == "UL_User_THR")
                {
                    KPI_Index = 14;
                    BL_Index = 6;
                }
                if (KPI1 == "DL_User_THR")
                {
                    KPI_Index = 15;
                    BL_Index = 7;
                }
                if (KPI1 == "HSDAP_Drop_Rate")
                {
                    KPI_Index = 16;
                    BL_Index = 8;
                }
                if (KPI1 == "HSUPA_Drop_Rate")
                {
                    KPI_Index = 17;
                    BL_Index = 9;
                }
                if (KPI1 == "MultiRAB_SR")
                {
                    KPI_Index = 18;
                    BL_Index = 10;
                }
                if (KPI1 == "PS_RRC_SR")
                {
                    KPI_Index = 19;
                    BL_Index = 11;
                }
                if (KPI1 == "Ps_RAB_Establish")
                {
                    KPI_Index = 20;
                    BL_Index = 12;
                }
                if (KPI1 == "PS_MultiRAB_Establish")
                {
                    KPI_Index = 21;
                    BL_Index =13;
                }
                if (KPI1 == "PS_Drop_Rate")
                {
                    KPI_Index = 22;
                    BL_Index = 14;
                }
                if (KPI1 == "HSDPA_Cell_Change_SR")
                {
                    KPI_Index = 23;
                    BL_Index = 15;
                }
                if (KPI1 == "HS_Share_Payload")
                {
                    KPI_Index = 24;
                    BL_Index =16;
                }
                if (KPI1 == "DL_Cell_THR")
                {
                    KPI_Index = 25;
                    BL_Index = 17;
                }
                if (KPI_Index != 0)
                {
                    dataGridView1.Rows[KPI_Index].Cells[0].Style.BackColor = Color.Yellow;
                    dataGridView1.Rows[KPI_Index].Cells[1].Style.BackColor = Color.Yellow;
                }

                if (KPI_Index != 0 && bl_data.Count != 0)
                {
                    BL = Convert.ToDouble(bl_data[0].ItemArray[BL_Index]);
                }





            }


            if (form1.Technology == "4G")
            {

                var cell_data = (from p in form1.Last_Day_List.AsEnumerable()
                                 where p.Field<string>("eNodeB") == Cell
                                 select p).ToList();

                var bl_data = (from p in KPI_Table_BL.AsEnumerable()
                               where p.Field<string>("eNodeB") == Cell
                               select p).ToList();

                dataGridView1.ColumnCount = 2;
                dataGridView1.Columns[0].Width = 150;
                dataGridView1.Columns[1].Width = 60;
                dataGridView1.RowCount = 23;

                DateTime d = Convert.ToDateTime(cell_data[0].ItemArray[0]);
                if (cell_data[0].ItemArray[2].ToString()!="")
                {
                    RNC_in_LTE = cell_data[0].ItemArray[2].ToString();
                }
                else
                {
                    RNC_in_LTE = "";
                }
                dataGridView1.Rows[0].Cells[0].Value = "Date"; dataGridView1.Rows[0].Cells[1].Value = Convert.ToString(d.Month) + "/" + Convert.ToString(d.Day) + "/" + Convert.ToString(d.Year);
                dataGridView1.Rows[1].Cells[0].Value = "Cell"; dataGridView1.Rows[1].Cells[1].Value = cell_data[0].ItemArray[1];
                dataGridView1.Rows[2].Cells[0].Value = "LEVEL"; dataGridView1.Rows[2].Cells[1].Value = cell_data[0].ItemArray[6];
                dataGridView1.Rows[3].Cells[0].Value = "Worst"; dataGridView1.Rows[3].Cells[1].Value = cell_data[0].ItemArray[7];
                dataGridView1.Rows[4].Cells[0].Value = "Effeciency (%)"; dataGridView1.Rows[4].Cells[1].Value = cell_data[0].ItemArray[8];
                dataGridView1.Rows[5].Cells[0].Value = "Status"; dataGridView1.Rows[5].Cells[1].Value = cell_data[0].ItemArray[9];
                dataGridView1.Rows[6].Cells[0].Value = "QIx"; dataGridView1.Rows[6].Cells[1].Value = cell_data[0].ItemArray[10];
                dataGridView1.Rows[7].Cells[0].Value = "QIxP"; dataGridView1.Rows[7].Cells[1].Value = cell_data[0].ItemArray[11];
                dataGridView1.Rows[8].Cells[0].Value = "QIxBL"; dataGridView1.Rows[8].Cells[1].Value = cell_data[0].ItemArray[12];
                dataGridView1.Rows[9].Cells[0].Value = "Payload (GB)"; dataGridView1.Rows[9].Cells[1].Value = cell_data[0].ItemArray[13];
                dataGridView1.Rows[10].Cells[0].Value = "Repeated Daily at 10 days (QIx<QIb)"; dataGridView1.Rows[10].Cells[1].Value = cell_data[0].ItemArray[14];
                dataGridView1.Rows[11].Cells[0].Value = "Total Days Availability > 90"; dataGridView1.Rows[11].Cells[1].Value = cell_data[0].ItemArray[15];
                dataGridView1.Rows[12].Cells[0].Value = "RRC_Connection_SR (% of WPC)"; dataGridView1.Rows[12].Cells[1].Value = cell_data[0].ItemArray[16];
                dataGridView1.Rows[13].Cells[0].Value = "ERAB_SR_Initial (% of WPC)"; dataGridView1.Rows[13].Cells[1].Value = cell_data[0].ItemArray[17];
                dataGridView1.Rows[14].Cells[0].Value = "ERAB_SR_Added (% of WPC)"; dataGridView1.Rows[14].Cells[1].Value = cell_data[0].ItemArray[18];
                dataGridView1.Rows[15].Cells[0].Value = "DL_THR (% of WPC)"; dataGridView1.Rows[15].Cells[1].Value = cell_data[0].ItemArray[19];
                dataGridView1.Rows[16].Cells[0].Value = "UL_THR (% of WPC)"; dataGridView1.Rows[16].Cells[1].Value = cell_data[0].ItemArray[20];
                dataGridView1.Rows[17].Cells[0].Value = "HO_SR (% of WPC)"; dataGridView1.Rows[17].Cells[1].Value = cell_data[0].ItemArray[21];
                dataGridView1.Rows[18].Cells[0].Value = "ERAB_Drop_Rate (% of WPC)"; dataGridView1.Rows[18].Cells[1].Value = cell_data[0].ItemArray[22];
                dataGridView1.Rows[19].Cells[0].Value = "S1_Signalling_SR (% of WPC)"; dataGridView1.Rows[19].Cells[1].Value = cell_data[0].ItemArray[24];
                dataGridView1.Rows[20].Cells[0].Value = "Inter_Freq_SR (% of WPC)"; dataGridView1.Rows[20].Cells[1].Value = cell_data[0].ItemArray[25];
                dataGridView1.Rows[21].Cells[0].Value = "Intra_Freq_SR (% of WPC)"; dataGridView1.Rows[21].Cells[1].Value = cell_data[0].ItemArray[26];
                dataGridView1.Rows[22].Cells[0].Value = "UL_Packet_Loss (% of WPC)"; dataGridView1.Rows[22].Cells[1].Value = cell_data[0].ItemArray[27];

                                  



                int KPI_Index = 0;
                int BL_Index = 0;
                if (KPI1 == "RRC_Connection_SR")
                {
                    KPI_Index = 12;
                    BL_Index = 4;
                }
                if (KPI1 == "ERAB_SR_Initial")
                {
                    KPI_Index = 13;
                    BL_Index = 5;
                }
                if (KPI1 == "ERAB_SR_Added")
                {
                    KPI_Index = 14;
                    BL_Index = 6;
                }
                if (KPI1 == "DL_THR")
                {
                    KPI_Index = 15;
                    BL_Index = 7;
                }
                if (KPI1 == "UL_THR")
                {
                    KPI_Index = 16;
                    BL_Index = 8;
                }
                if (KPI1 == "HO_SR")
                {
                    KPI_Index = 17;
                    BL_Index = 9;
                }
                if (KPI1 == "ERAB_Drop_Rate")
                {
                    KPI_Index = 18;
                    BL_Index = 10;
                }
                if (KPI1 == "S1_Signalling_SR")
                {
                    KPI_Index = 19;
                    BL_Index =11;
                }
                if (KPI1 == "Inter_Freq_SR")
                {
                    KPI_Index = 20;
                    BL_Index = 12;
                }
                if (KPI1 == "Intra_Freq_SR")
                {
                    KPI_Index = 21;
                    BL_Index = 13;
                }
                if (KPI1 == "UL_Packet_Loss")
                {
                    KPI_Index = 22;
                    BL_Index = 14;
                }

                if (KPI_Index != 0)
                {
                    dataGridView1.Rows[KPI_Index].Cells[0].Style.BackColor = Color.Yellow;
                    dataGridView1.Rows[KPI_Index].Cells[1].Style.BackColor = Color.Yellow;
                }

                if (KPI_Index != 0 && bl_data.Count != 0)
                {
                    BL = Convert.ToDouble(bl_data[0].ItemArray[BL_Index]);
                }


            }




            double min_value_KPI = 100000;
            double max_value_KPI = -100000;

            string Node = "";
            for (int i = 0; i < KPI_Table_E.Rows.Count; i++)
            {
                DateTime dt = Convert.ToDateTime((KPI_Table_E.Rows[i]).ItemArray[1]);
                if (form1.Technology!="4G")
                {
                    Node = (KPI_Table_E.Rows[i]).ItemArray[0].ToString();
                }
               else
                {
                    Node = (KPI_Table_E.Rows[i]).ItemArray[2].ToString();
                }
                if ((KPI_Table_E.Rows[i]).ItemArray[2].ToString() != "")
                {
                    double KPI = Convert.ToDouble((KPI_Table_E.Rows[i]).ItemArray[2]);
                    chart1.Series[0].Points.AddXY(dt, KPI);
                
                    if (BL != -1000)
                    {
                       if (dt.Month==8 && dt.Day >= 22)
                        {
                            chart1.Series[3].Points.AddXY(dt, BL);
                        }
                        if (dt.Month > 8)
                        {
                            chart1.Series[3].Points.AddXY(dt, BL);
                        }
                    }
                    if (KPI > max_value_KPI)
                    {
                        //max_value_KPI = Math.Round(KPI, MidpointRounding.AwayFromZero);
                        max_value_KPI = KPI;
                    }
                    if (KPI < min_value_KPI)
                    {
                        //min_value_KPI = Math.Round(KPI, MidpointRounding.AwayFromZero);
                        min_value_KPI = KPI;
                    }

                }
                
            }

            for (int i = 0; i < KPI_Table_N.Rows.Count; i++)
            {
                DateTime dt = Convert.ToDateTime((KPI_Table_N.Rows[i]).ItemArray[1]);
                if (form1.Technology != "4G")
                {
                    Node = (KPI_Table_N.Rows[i]).ItemArray[0].ToString();
                }
                else
                {
                    Node = (KPI_Table_N.Rows[i]).ItemArray[2].ToString();
                }
                if ((KPI_Table_N.Rows[i]).ItemArray[2].ToString() != "")
                {
                    double KPI = Convert.ToDouble((KPI_Table_N.Rows[i]).ItemArray[2]);
                    chart1.Series[2].Points.AddXY(dt, KPI);
                 
                    if (BL != -1000)
                    {
                        if (dt.Month == 8 && dt.Day >= 22)
                        {
                            chart1.Series[3].Points.AddXY(dt, BL);
                        }
                        if (dt.Month > 8)
                        {
                            chart1.Series[3].Points.AddXY(dt, BL);
                        }
                    }
                    if (KPI > max_value_KPI)
                    {
                        //max_value_KPI = Math.Round(KPI, MidpointRounding.AwayFromZero);
                        max_value_KPI = KPI;
                    }
                    if (KPI < min_value_KPI)
                    {
                        //min_value_KPI = Math.Round(KPI, MidpointRounding.AwayFromZero);
                        min_value_KPI = KPI;
                    }

                }
            }

            for (int i = 0; i < KPI_Table_H.Rows.Count; i++)
            {
                DateTime dt = Convert.ToDateTime((KPI_Table_H.Rows[i]).ItemArray[1]);
                if (form1.Technology != "4G")
                {
                    Node = (KPI_Table_H.Rows[i]).ItemArray[0].ToString();
                }
                else
                {
                    Node = (KPI_Table_H.Rows[i]).ItemArray[2].ToString();
                }
                if ((KPI_Table_H.Rows[i]).ItemArray[2].ToString() != "")
                {
                    double KPI = Convert.ToDouble((KPI_Table_H.Rows[i]).ItemArray[2]);
                    chart1.Series[1].Points.AddXY(dt, KPI);
                    if (BL != -1000)
                    {
                        if (dt.Month == 8 && dt.Day >= 22)
                        {
                            chart1.Series[3].Points.AddXY(dt, BL);
                        }
                        if (dt.Month > 8)
                        {
                            chart1.Series[3].Points.AddXY(dt, BL);
                        }
                    }
                    if (KPI > max_value_KPI)
                    {
                        //max_value_KPI = Math.Round(KPI, MidpointRounding.AwayFromZero);
                        max_value_KPI = KPI;
                    }
                    if (KPI < min_value_KPI)
                    {
                        //min_value_KPI = Math.Round(KPI, MidpointRounding.AwayFromZero);
                        min_value_KPI = KPI;
                    }

                }

            
            }


            if (form1.Technology!="4G")
            {
                Title title = chart1.Titles.Add(Node + "_" + form1.Selected_Cell + "_" + form1.Selected_KPI);
                title.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            }
    else
            {
                Title title = chart1.Titles.Add(RNC_in_LTE + "_" + form1.Selected_Cell + "_" + form1.Selected_KPI);
                title.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            }
            chart1.Series[0].IsVisibleInLegend = false;
            chart1.Series[1].IsVisibleInLegend = false;
            chart1.Series[2].IsVisibleInLegend = false;
            chart1.Series[3].IsVisibleInLegend = false;
            chart1.ChartAreas[0].AxisY.Maximum = max_value_KPI;
            chart1.ChartAreas[0].AxisY.Minimum = min_value_KPI;





        }

        private void button1_Click(object sender, EventArgs e)
        {
            chart1.SaveImage("Image.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
            Image img = Image.FromFile("Image.jpg");
            System.Windows.Forms.Clipboard.SetImage(img);
        }
    }
}
