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
using Excel = Microsoft.Office.Interop.Excel;
using ClosedXML;
using ClosedXML.Excel;

namespace CWA
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            ConnectionString = @"Server=" + Server_Name + "; Database=" + DataBase_Name + "; User ID=cwpcApp; Password=cwpcApp@830625#Ahmad";
            connection = new SqlConnection(ConnectionString);
            connection.Open();


            //ConnectionString = @"Server=" + Server_Name + "; Database=" + DataBase_Name + "; Integrated Security=True";
            //connection = new SqlConnection(ConnectionString);
            //connection.Open();


            string userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            string currentUser = userName.Substring(8, userName.Length - 8);

            //if (currentUser != null)
            //{


            //    var excelApplication = new Excel.Application();

            //    var excelWorkBook = excelApplication.Application.Workbooks.Add(Type.Missing);

            //    excelApplication.Cells[1, 1] = currentUser;
            //    excelApplication.Cells[1, 2] = DateTime.Now;

            //    string CR_PATH = string.Format(@"\\DFS\fs\NPO\6. Performance\Quary.xlsx");
            //    excelApplication.ActiveWorkbook.SaveCopyAs(CR_PATH);
            //    //excelApplication.ActiveWorkbook.SaveCopyAs(@"F:\test.xlsx");

            //    excelApplication.ActiveWorkbook.Saved = true;

            //    // Close the Excel Application
            //    excelApplication.Quit();
            //}

        }


        public double Max_Y1 = 0;
        public double Max_Y2 = 0;
        public double Max_Y3 = 0;
        public double Max_Y4 = 0;
        public DateTime Min_X = DateTime.Today;
        public DateTime Max_X = DateTime.Today;
        public DateTime Min_X1 = DateTime.Today;
        public DateTime Max_X1 = DateTime.Today;
        public double Min_X2 = 10000000000;
        public double Max_X2 = -10000000000;
        public DateTime Min_X_Date = DateTime.Today;
        public DateTime Max_X_Date = DateTime.Today;

        public string ConnectionString = "";
        public SqlConnection connection = new SqlConnection();

        // Technology
        public DataTable Node_Table = new DataTable();
        public string Technology = "";
        public string Region = "";
        public string Province = "";
        public string Node = "";

        // Parameters for Selection Time 
        public DateTime Node_Selection_Time = DateTime.Now;
        public DateTime Owner_Selection_Time = DateTime.Now;
        public DateTime Region_Selection_Time = DateTime.Now;

        public DateTime checking_date = DateTime.Today;
        public int checking_day_num = 0;

        //public string Server_Name = @"NAKPRG-NB1243\" + "AHMAD";
        //public string DataBase_Name = "Contract";


        public string Server_Name = "PERFORMANCEDB01";
        public string DataBase_Name = "Performance_NAK";

        //public string Server_Name = "core";
        //public string DataBase_Name = "Core_Performance_Mohammad";


        //public string Server_Name = "172.26.7.159";
        //public string DataBase_Name = "Performance_NAK";

        public DataTable Date_Table1 = new DataTable();
        public string Solved_Quary = "";
        public string Raised_Quary = "";
        public string Repeated_Quary = "";
        public DataTable Solved_Data_Table = new DataTable();
        public DataTable Raised_Data_Table = new DataTable();
        public DataTable Repeated_Date_Table = new DataTable();
        public DataTable Repeated_Data_Table = new DataTable();
        public DataTable Repeated_Data_Table_Other = new DataTable();


        public string Selected_KPI = "";
        public string Selected_Cell = "";
        public string KPI_CSSR_E = "";
        public string KPI_CSSR_H = "";
        public string KPI_CSSR_N = "";

        public DataTable KPI_Table_E1 = new DataTable();
        public DataTable KPI_Table_H1 = new DataTable();
        public DataTable KPI_Table_N1 = new DataTable();

        public string chart_type = "";

        public int Level_Counter = 0;
        public double Score_Cell_Sum1 = 0;
        public double Score_Cell_Sum = 0;
        public int Region_Show = 0;
        public int Province_Show = 0;
        public int Node_Show = 0;

        public DataTable Last_Day_List = new DataTable();
        public DataTable BL_Table = new DataTable();


        public DataTable Single_Region_Table_Contractual = new DataTable();
        public DataTable Single_Province_Table_Contractual = new DataTable();
        public DataTable Single_Node_Table_Contractual = new DataTable();

        void comboBox3_MouseWheel(object sender, MouseEventArgs e)
        {
            ((HandledMouseEventArgs)e).Handled = true;
        }
        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

            Region_Show = 0;
            Province_Show = 0;
            Node_Show = 0;


            comboBox3.MouseWheel += new MouseEventHandler(comboBox3_MouseWheel);
            Technology = comboBox3.SelectedItem.ToString();
            if (Technology == "2G_CS")
            {
                comboBox7.Items.Clear();
                comboBox7.Items.Add("CSSR");
                comboBox7.Items.Add("OHSR");
                comboBox7.Items.Add("CDR");
                comboBox7.Items.Add("TCH_ASFR");
                comboBox7.Items.Add("RXDL");
                comboBox7.Items.Add("RXUL");
                comboBox7.Items.Add("SDCCH_CONG");
                comboBox7.Items.Add("SDCCH_SR");
                comboBox7.Items.Add("SDCCH_DROP");
                comboBox7.Items.Add("IHSR");
                comboBox7.Items.Add("TCH Traffic (Erlang)");
                comboBox7.Items.Add("Availability");
            }
            if (Technology == "2G_PS")
            {
                comboBox7.Items.Clear();
                comboBox7.Items.Add("TBF_Establish");
                comboBox7.Items.Add("TBF_Drop");
                comboBox7.Items.Add("GPRS_THR");
                comboBox7.Items.Add("EGPRS_THR");
                comboBox7.Items.Add("GPRS_THR_per_TS");
                comboBox7.Items.Add("EGPRS_THR_per_TS");
                comboBox7.Items.Add("PS Traffic (KB)");
                comboBox7.Items.Add("Availability");
            }
            if (Technology == "3G_CS")
            {
                comboBox7.Items.Clear();
                comboBox7.Items.Add("CS_RAB_Establish");
                comboBox7.Items.Add("CS_IRAT_HO_SR");
                comboBox7.Items.Add("CS_Drop_Rate");
                comboBox7.Items.Add("Soft_HO_SR");
                comboBox7.Items.Add("CS_RRC_SR");
                comboBox7.Items.Add("CS Traffic (Erlang)");
                comboBox7.Items.Add("Availability");
            }
            if (Technology == "3G_PS")
            {
                comboBox7.Items.Clear();
                comboBox7.Items.Add("HSDPA_SR");
                comboBox7.Items.Add("HSUPA_SR");
                comboBox7.Items.Add("UL_User_THR");
                comboBox7.Items.Add("DL_User_THR");
                comboBox7.Items.Add("HSDAP_Drop_Rate");
                comboBox7.Items.Add("HSUPA_Drop_Rate");
                comboBox7.Items.Add("MultiRAB_SR");
                comboBox7.Items.Add("PS_RRC_SR");
                comboBox7.Items.Add("Ps_RAB_Establish");
                comboBox7.Items.Add("PS_MultiRAB_Establish");
                comboBox7.Items.Add("PS_Drop_Rate");
                comboBox7.Items.Add("HSDPA_Cell_Change_SR");
                comboBox7.Items.Add("HS_Share_Payload");
                comboBox7.Items.Add("HS_Share_Payload");
                comboBox7.Items.Add("DL_Cell_THR");
                comboBox7.Items.Add("PS Traffic (GB)");
                comboBox7.Items.Add("Availability");
            }
            if (Technology == "4G")
            {
                comboBox7.Items.Clear();
                comboBox7.Items.Add("RRC_Connection_SR");
                comboBox7.Items.Add("ERAB_SR_Initial");
                comboBox7.Items.Add("ERAB_SR_Added");
                comboBox7.Items.Add("DL_THR");
                comboBox7.Items.Add("UL_THR");
                comboBox7.Items.Add("HO_SR");
                comboBox7.Items.Add("ERAB_Drop_Rate");
                comboBox7.Items.Add("S1_Signalling_SR");
                comboBox7.Items.Add("Inter_Freq_SR");
                comboBox7.Items.Add("Intra_Freq_SR");
                comboBox7.Items.Add("UL_Packet_Loss");
                comboBox7.Items.Add("Data Traffic (GB)");
                comboBox7.Items.Add("Availability");
            }
        }


        // BSC/RNC Selection
        void comboBox4_MouseWheel(object sender, MouseEventArgs e)
        {
            ((HandledMouseEventArgs)e).Handled = true;
        }
        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox4.MouseWheel += new MouseEventHandler(comboBox4_MouseWheel);
            Region_Show = 0;
            Province_Show = 0;
            Node_Show = 1;
            label1.BackColor = Color.PaleGoldenrod;
            label2.BackColor = Color.PaleGoldenrod;
            label4.BackColor = Color.Yellow;
            label5.BackColor = Color.PaleGoldenrod;
            Node_Selection_Time = DateTime.Now;

            Node = comboBox4.SelectedItem.ToString();
            string Single_Node_List_Quary_Contractual = "";
            string Single_Node_List_Quary_NearContractual = "";
            string Single_Node_Misssing_Quary_Contractual = "";
            string Single_Node_Missing_Quary_NearContractual = "";
            if (Technology == "2G_CS")
            {
                Single_Node_List_Quary_Contractual = "select [Date], [BSC], [Level], [Status], count(Level) as 'Worst_Count' from [Contractual_WPC_2G_CS] where [BSC] = '" + Node + "' and Status = 'Contractual WPC' group by[Date], [BSC], [Level], [Status]  order by Date";
                Single_Node_List_Quary_NearContractual = "select [Date], [BSC], [Level], [Status], count(Level) as 'Worst_Count' from [Contractual_WPC_2G_CS] where [BSC] = '" + Node + "' and Status = 'Near Contractual WPC' group by[Date], [BSC], [Level], [Status]  order by Date";
                Single_Node_Misssing_Quary_Contractual = "select [Date], [BSC], [Status], [Level], sum([Effeciency Index])*1.1 as 'Missing_Score' from [Contractual_WPC_2G_CS] where [BSC] = '" + Node + "' and Status = 'Contractual WPC' group by [Date], [BSC], [Status] ,[Level] order by Date";
                Single_Node_Missing_Quary_NearContractual = "select [Date], [BSC], [Status], [Level], sum([Effeciency Index]) as 'Missing_Score' from [Contractual_WPC_2G_CS] where [BSC] = '" + Node + "' and Status = 'Near Contractual WPC'  and QIxP<100 group by [Date], [BSC], [Status] ,[Level] order by Date";
            }
            if (Technology == "2G_PS")
            {
                Single_Node_List_Quary_Contractual = "select [Date], [BSC], [Level], [Status], count(Level) as 'Worst_Count' from [Contractual_WPC_2G_PS] where [BSC] = '" + Node + "' and Status = 'Contractual WPC' group by[Date], [BSC], [Level], [Status]  order by Date";
                Single_Node_List_Quary_NearContractual = "select [Date], [BSC], [Level], [Status], count(Level) as 'Worst_Count' from [Contractual_WPC_2G_PS] where [BSC] = '" + Node + "' and Status = 'Near WPC' group by[Date], [BSC], [Level], [Status]  order by Date";
                Single_Node_Misssing_Quary_Contractual = "select [Date], [BSC], [Status], [Level], sum([Effeciency Index])*1.1 as 'Missing_Score' from [Contractual_WPC_2G_PS] where [BSC] = '" + Node + "' and Status = 'Contractual WPC' group by [Date], [BSC], [Status] ,[Level] order by Date";
                Single_Node_Missing_Quary_NearContractual = "select [Date], [BSC], [Status], [Level], sum([Effeciency Index]) as 'Missing_Score' from [Contractual_WPC_2G_PS] where [BSC] = '" + Node + "' and Status = 'Near WPC'  and QIxP<100 group by [Date], [BSC], [Status] ,[Level] order by Date";
            }
            if (Technology == "3G_CS")
            {
                Single_Node_List_Quary_Contractual = "select [Date], [RNC], [Level], [Status], count(Level) as 'Worst_Count' from [Contractual_WPC_3G_CS] where [RNC] = '" + Node + "' and Status = 'Contractual WPC' group by[Date], [RNC], [Level], [Status]  order by Date";
                Single_Node_List_Quary_NearContractual = "select [Date], [RNC], [Level], [Status], count(Level) as 'Worst_Count' from [Contractual_WPC_3G_CS] where [RNC] = '" + Node + "' and Status = 'Near Contractual WPC' group by[Date], [RNC], [Level], [Status]  order by Date";
                Single_Node_Misssing_Quary_Contractual = "select [Date], [RNC], [Status], [Level], sum([Effeciency Index])*1.1 as 'Missing_Score' from [Contractual_WPC_3G_CS] where [RNC] = '" + Node + "' and Status = 'Contractual WPC' group by [Date], [RNC], [Status] ,[Level] order by Date";
                Single_Node_Missing_Quary_NearContractual = "select [Date], [RNC], [Status], [Level], sum([Effeciency Index]) as 'Missing_Score' from [Contractual_WPC_3G_CS] where [RNC] = '" + Node + "' and Status = 'Near Contractual WPC' and QIxP<100 group by [Date], [RNC], [Status] ,[Level] order by Date";
            }
            if (Technology == "3G_PS")
            {
                Single_Node_List_Quary_Contractual = "select [Date], [RNC], [Level], [Status], count(Level) as 'Worst_Count' from [Contractual_WPC_3G_PS] where [RNC] = '" + Node + "' and Status = 'Contractual WPC' group by[Date], [RNC], [Level], [Status]  order by Date";
                Single_Node_List_Quary_NearContractual = "select [Date], [RNC], [Level], [Status], count(Level) as 'Worst_Count' from [Contractual_WPC_3G_PS] where [RNC] = '" + Node + "' and Status = 'Near Contractual WPC' group by[Date], [RNC], [Level], [Status]  order by Date";
                Single_Node_Misssing_Quary_Contractual = "select [Date], [RNC], [Status], [Level], sum([Effeciency Index])*1.1 as 'Missing_Score' from [Contractual_WPC_3G_PS] where [RNC] = '" + Node + "' and Status = 'Contractual WPC' group by [Date], [RNC], [Status] ,[Level] order by Date";
                Single_Node_Missing_Quary_NearContractual = "select [Date], [RNC], [Status], [Level], sum([Effeciency Index]) as 'Missing_Score' from [Contractual_WPC_3G_PS] where [RNC] = '" + Node + "' and Status = 'Near Contractual WPC' and QIxP<100 group by [Date], [RNC], [Status] ,[Level] order by Date";
            }
            if (Technology == "4G")
            {
                Single_Node_List_Quary_Contractual = "select [Date],  [RNC], [Level], [Status], count(Level) as 'Worst_Count' from [Contractual_WPC_4G] where [RNC] = '" + Node + "'  and Status = 'Contractual WPC' group by [Date],  [RNC], [Level], [Status]  order by Date";
                Single_Node_List_Quary_NearContractual = "select [Date], [RNC], [Level], [Status], count(Level) as 'Worst_Count' from [Contractual_WPC_4G] where [RNC] = '" + Node + "'  and Status = 'Near Contractual WPC' group by[Date], [RNC], [Level], [Status]  order by Date";
                Single_Node_Misssing_Quary_Contractual = "select [Date],[RNC], [Status], [Level], sum([Effeciency Index])*1.1 as 'Missing_Score' from [Contractual_WPC_4G] where [RNC] = '" + Node + "'  and Status = 'Contractual WPC' group by [Date], [RNC], [Status] ,[Level] order by Date";
                Single_Node_Missing_Quary_NearContractual = "select [Date], [RNC],  [Status], [Level], sum([Effeciency Index]) as 'Missing_Score' from [Contractual_WPC_4G] where [RNC] = '" + Node + "'  and Status = 'Near Contractual WPC' and QIxP<100 group by [Date],  [RNC], [Status] ,[Level] order by Date";
            }

            // Worst Cells Count in Contractual WPC
            SqlCommand Single_Node_List_Quary_Contractual1 = new SqlCommand(Single_Node_List_Quary_Contractual, connection);
            Single_Node_List_Quary_Contractual1.ExecuteNonQuery();
            Single_Node_Table_Contractual = new DataTable();
            SqlDataAdapter dataAdapter_Contractual = new SqlDataAdapter(Single_Node_List_Quary_Contractual1);
            dataAdapter_Contractual.Fill(Single_Node_Table_Contractual);

            // Worst Cells Count in Near Contractual WPC
            SqlCommand Single_Node_List_Quary_NearContractual1 = new SqlCommand(Single_Node_List_Quary_NearContractual, connection);
            Single_Node_List_Quary_NearContractual1.ExecuteNonQuery();
            DataTable Single_Node_Table_NearContractual = new DataTable();
            SqlDataAdapter dataAdapter_NearContractual = new SqlDataAdapter(Single_Node_List_Quary_NearContractual1);
            dataAdapter_NearContractual.Fill(Single_Node_Table_NearContractual);

            // Worst Cells Missing Score in Contractual WPC
            SqlCommand Single_Node_Missing_Quary_Contractual1 = new SqlCommand(Single_Node_Misssing_Quary_Contractual, connection);
            Single_Node_Missing_Quary_Contractual1.ExecuteNonQuery();
            DataTable Single_Node_Table_Contractual1 = new DataTable();
            SqlDataAdapter dataAdapter_Contractual1 = new SqlDataAdapter(Single_Node_Missing_Quary_Contractual1);
            dataAdapter_Contractual1.Fill(Single_Node_Table_Contractual1);

            // Worst Cells Missing Score in Near Contractual WPC
            SqlCommand Single_Node_Missing_Quary_NearContractual1 = new SqlCommand(Single_Node_Missing_Quary_NearContractual, connection);
            Single_Node_Missing_Quary_NearContractual1.ExecuteNonQuery();
            DataTable Single_Node_Table_NearContractual1 = new DataTable();
            SqlDataAdapter dataAdapter_NearContractual1 = new SqlDataAdapter(Single_Node_Missing_Quary_NearContractual1);
            dataAdapter_NearContractual1.Fill(Single_Node_Table_NearContractual1);


            chart1.Series.Clear();
            chart1.Titles.Clear();

            Series newSeries1 = new Series();
            chart1.Series.Add(newSeries1);
            newSeries1.IsXValueIndexed = false;
            chart1.Series[0].ChartType = SeriesChartType.Line;
            chart1.Series[0].Color = Color.Red;
            chart1.Series[0].BorderWidth = 3;
            chart1.ChartAreas[0].AxisX.Interval = 5;
            chart1.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart1.Series[0].EmptyPointStyle.Color = Color.Transparent;
            chart1.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";
            chart1.Series[0].ToolTip = "#VALX [#VALY]";
            chart1.Series[0].IsValueShownAsLabel = false;
            chart1.Series[0].LegendText = "Level 1";
            chart1.Legends["Legend1"].Docking = Docking.Bottom;
            newSeries1.MarkerStyle = MarkerStyle.Circle;
            newSeries1.MarkerSize = 6;

            Series newSeries2 = new Series();
            chart1.Series.Add(newSeries2);
            newSeries2.IsXValueIndexed = false;
            chart1.Series[1].ChartType = SeriesChartType.Line;
            chart1.Series[1].Color = Color.Orange;
            chart1.Series[1].BorderWidth = 3;
            chart1.Series[1].EmptyPointStyle.Color = Color.Transparent;
            chart1.Series[1].ToolTip = "#VALX [#VALY]";
            chart1.Series[1].IsValueShownAsLabel = false;
            chart1.Series[1].LegendText = "Level 2";
            newSeries2.MarkerStyle = MarkerStyle.Circle;
            newSeries2.MarkerSize = 6;

            Series newSeries3 = new Series();
            chart1.Series.Add(newSeries3);
            newSeries3.IsXValueIndexed = false;
            chart1.Series[2].ChartType = SeriesChartType.Line;
            chart1.Series[2].Color = Color.Yellow;
            chart1.Series[2].BorderWidth = 3;
            chart1.Series[2].EmptyPointStyle.Color = Color.Transparent;
            chart1.Series[2].ToolTip = "#VALX [#VALY]";
            chart1.Series[2].IsValueShownAsLabel = false;
            chart1.Series[2].LegendText = "Level 3";
            newSeries3.MarkerStyle = MarkerStyle.Circle;
            newSeries3.MarkerSize = 6;


            Series newSeries4 = new Series();
            chart1.Series.Add(newSeries4);
            newSeries4.IsXValueIndexed = false;
            chart1.Series[3].ChartType = SeriesChartType.Line;
            chart1.Series[3].Color = Color.Blue;
            chart1.Series[3].BorderWidth = 3;
            chart1.Series[3].EmptyPointStyle.Color = Color.Transparent;
            chart1.Series[3].ToolTip = "#VALX [#VALY]";
            chart1.Series[3].IsValueShownAsLabel = false;
            chart1.Series[3].LegendText = "Level 4";
            newSeries4.MarkerStyle = MarkerStyle.Circle;
            newSeries4.MarkerSize = 6;


            Series newSeries5 = new Series();
            chart1.Series.Add(newSeries5);
            newSeries5.IsXValueIndexed = false;
            chart1.Series[4].ChartType = SeriesChartType.Line;
            chart1.Series[4].Color = Color.Green;
            chart1.Series[4].BorderWidth = 3;
            chart1.Series[4].EmptyPointStyle.Color = Color.Transparent;
            chart1.Series[4].ToolTip = "#VALX [#VALY]";
            chart1.Series[4].IsValueShownAsLabel = false;
            chart1.Series[4].LegendText = "Level 5";
            newSeries5.MarkerStyle = MarkerStyle.Circle;
            newSeries5.MarkerSize = 6;


            chart2.Series.Clear();
            chart2.Titles.Clear();

            Series newSeries6 = new Series();
            chart2.Series.Add(newSeries6);
            newSeries6.IsXValueIndexed = false;
            chart2.Series[0].ChartType = SeriesChartType.Line;
            chart2.Series[0].Color = Color.Red;
            chart2.Series[0].BorderWidth = 3;
            chart2.ChartAreas[0].AxisX.Interval = 5;
            chart2.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart2.Series[0].EmptyPointStyle.Color = Color.Transparent;
            chart2.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";
            chart2.Series[0].ToolTip = "#VALX [#VALY]";
            chart2.Series[0].IsValueShownAsLabel = false;
            chart2.Series[0].LegendText = "Level 1";
            chart2.Legends["Legend1"].Docking = Docking.Bottom;
            newSeries6.MarkerStyle = MarkerStyle.Circle;
            newSeries6.MarkerSize = 6;


            Series newSeries7 = new Series();
            chart2.Series.Add(newSeries7);
            newSeries7.IsXValueIndexed = false;
            chart2.Series[1].ChartType = SeriesChartType.Line;
            chart2.Series[1].Color = Color.Orange;
            chart2.Series[1].BorderWidth = 3;
            chart2.Series[1].EmptyPointStyle.Color = Color.Transparent;
            chart2.Series[1].ToolTip = "#VALX [#VALY]";
            chart2.Series[1].IsValueShownAsLabel = false;
            chart2.Series[1].LegendText = "Level 2";
            newSeries7.MarkerStyle = MarkerStyle.Circle;
            newSeries7.MarkerSize = 6;


            Series newSeries8 = new Series();
            chart2.Series.Add(newSeries8);
            newSeries8.IsXValueIndexed = false;
            chart2.Series[2].ChartType = SeriesChartType.Line;
            chart2.Series[2].Color = Color.Yellow;
            chart2.Series[2].BorderWidth = 3;
            chart2.Series[2].EmptyPointStyle.Color = Color.Transparent;
            chart2.Series[2].ToolTip = "#VALX [#VALY]";
            chart2.Series[2].IsValueShownAsLabel = false;
            chart2.Series[2].LegendText = "Level 3";
            newSeries8.MarkerStyle = MarkerStyle.Circle;
            newSeries8.MarkerSize = 6;


            Series newSeries9 = new Series();
            chart2.Series.Add(newSeries9);
            newSeries9.IsXValueIndexed = false;
            chart2.Series[3].ChartType = SeriesChartType.Line;
            chart2.Series[3].Color = Color.Blue;
            chart2.Series[3].BorderWidth = 3;
            chart2.Series[3].EmptyPointStyle.Color = Color.Transparent;
            chart2.Series[3].ToolTip = "#VALX [#VALY]";
            chart2.Series[3].IsValueShownAsLabel = false;
            chart2.Series[3].LegendText = "Level 4";
            newSeries9.MarkerStyle = MarkerStyle.Circle;
            newSeries9.MarkerSize = 6;


            Series newSeries10 = new Series();
            chart2.Series.Add(newSeries10);
            newSeries10.IsXValueIndexed = false;
            chart2.Series[4].ChartType = SeriesChartType.Line;
            chart2.Series[4].Color = Color.Green;
            chart2.Series[4].BorderWidth = 3;
            chart2.Series[4].EmptyPointStyle.Color = Color.Transparent;
            chart2.Series[4].ToolTip = "#VALX [#VALY]";
            chart2.Series[4].IsValueShownAsLabel = false;
            chart2.Series[4].LegendText = "Level 5";
            newSeries10.MarkerStyle = MarkerStyle.Circle;
            newSeries10.MarkerSize = 6;




            Title title1 = chart1.Titles.Add("Number of '" + Node + "' Worst Cells (Contractual)");
            title1.Font = new System.Drawing.Font("Arial", 12, FontStyle.Regular);
            Title title2 = chart2.Titles.Add("Number of '" + Node + "' Worst Cells (Near Contractual)");
            title2.Font = new System.Drawing.Font("Arial", 12, FontStyle.Regular);




            chart3.Series.Clear();
            chart3.Titles.Clear();

            Series newSeries11 = new Series();
            chart3.Series.Add(newSeries11);
            newSeries11.IsXValueIndexed = false;
            chart3.Series[0].ChartType = SeriesChartType.StackedColumn;
            chart3.Series[0].Color = Color.Brown;
            chart3.Series[0].BorderWidth = 3;
            chart3.ChartAreas[0].AxisX.Interval = 5;
            chart3.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart3.Series[0].EmptyPointStyle.Color = Color.Transparent;
            chart3.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";
            chart3.Series[0].ToolTip = "#VALX [#VALY]";
            chart3.Series[0].IsValueShownAsLabel = false;
            chart3.Series[0].IsVisibleInLegend = false;
            newSeries11.MarkerStyle = MarkerStyle.Circle;
            newSeries11.MarkerSize = 6;


            chart4.Series.Clear();
            chart4.Titles.Clear();

            Series newSeries12 = new Series();
            chart4.Series.Add(newSeries12);
            newSeries12.IsXValueIndexed = false;
            chart4.Series[0].ChartType = SeriesChartType.StackedColumn;
            chart4.Series[0].Color = Color.Brown;
            chart4.Series[0].BorderWidth = 3;
            chart4.ChartAreas[0].AxisX.Interval = 5;
            chart4.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart4.Series[0].EmptyPointStyle.Color = Color.Transparent;
            chart4.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";
            chart4.Series[0].ToolTip = "#VALX [#VALY]";
            chart4.Series[0].IsValueShownAsLabel = false;
            chart4.Series[0].IsVisibleInLegend = false;
            newSeries12.MarkerStyle = MarkerStyle.Circle;
            newSeries12.MarkerSize = 6;



            Title title3 = chart3.Titles.Add("Sum of '" + Node + "' Missing Score (Contractual)");
            title3.Font = new System.Drawing.Font("Arial", 12, FontStyle.Regular);
            Title title4 = chart4.Titles.Add("Sum of '" + Node + "' Missing Score (Near Contractual)");
            title4.Font = new System.Drawing.Font("Arial", 12, FontStyle.Regular);


            int y1 = 0;
            for (int k = 1; k <= Single_Node_Table_Contractual.Rows.Count; k++)
            {
                DateTime dt1 = Convert.ToDateTime((Single_Node_Table_Contractual.Rows[k - 1]).ItemArray[0]);
                dt1 = dt1.AddHours(23);
                double Worst_Cell_Count1 = 0;
                double Level1 = 0;
                if (Technology == "2G_CS" || Technology == "2G_PS" || Technology == "3G_CS" || Technology == "3G_PS")
                {
                    Worst_Cell_Count1 = Convert.ToDouble((Single_Node_Table_Contractual.Rows[k - 1]).ItemArray[4]);
                    Level1 = Convert.ToDouble((Single_Node_Table_Contractual.Rows[k - 1]).ItemArray[2]);
                }
                if (Technology == "4G")
                {
                    Worst_Cell_Count1 = Convert.ToDouble((Single_Node_Table_Contractual.Rows[k - 1]).ItemArray[4]);
                    Level1 = Convert.ToDouble((Single_Node_Table_Contractual.Rows[k - 1]).ItemArray[2]);
                }



                // Setting of Intervals
                // **************************************************************
                double dt1_double = dt1.Year * 10000 + dt1.Month * 100 + dt1.Day;

                if (dt1_double > Max_X2)
                {
                    Max_X2 = dt1_double;
                    Max_X_Date = dt1;
                }
                if (dt1_double < Min_X2)
                {
                    Min_X2 = dt1_double;
                    Min_X_Date = dt1;
                }
                // **************************************************************








                if (Level1 == 1 && checkBox6.Checked == true)
                {
                    chart1.Series[0].Points.AddXY(dt1, Worst_Cell_Count1);
                }
                if (Level1 == 2 && checkBox7.Checked == true)
                {
                    chart1.Series[1].Points.AddXY(dt1, Worst_Cell_Count1);
                }
                if (Level1 == 3 && checkBox8.Checked == true)
                {
                    chart1.Series[2].Points.AddXY(dt1, Worst_Cell_Count1);
                }
                if (Level1 == 4 && checkBox9.Checked == true)
                {
                    chart1.Series[3].Points.AddXY(dt1, Worst_Cell_Count1);
                }
                if (Level1 == 5 && checkBox10.Checked == true)
                {
                    chart1.Series[4].Points.AddXY(dt1, Worst_Cell_Count1);
                }
                Max_Y1 = chart1.ChartAreas[0].AxisY.Maximum;
                if (k == 1)
                {
                    Min_X = dt1;
                }

                y1 = k;
            }
            Max_X = Convert.ToDateTime((Single_Node_Table_Contractual.Rows[y1 - 1]).ItemArray[0]);






            // Setting of Intervals
            // **************************************************************
            double difference_day = (Max_X_Date - Min_X_Date).TotalDays;
            double day_interval = Math.Round(difference_day / 20);
            if (day_interval == 0)
            {
                day_interval = 1;
            }
            chart1.ChartAreas[0].AxisX.Interval = day_interval;
            chart2.ChartAreas[0].AxisX.Interval = day_interval;
            chart3.ChartAreas[0].AxisX.Interval = day_interval;
            chart4.ChartAreas[0].AxisX.Interval = day_interval;
            // **************************************************************




            int y2 = 0;
            for (int k = 1; k <= Single_Node_Table_NearContractual.Rows.Count; k++)
            {

                DateTime dt2 = Convert.ToDateTime((Single_Node_Table_NearContractual.Rows[k - 1]).ItemArray[0]);
                dt2 = dt2.AddHours(23);
                double Worst_Cell_Count2 = 0;
                double Level2 = 0;
                if (Technology == "2G_CS" || Technology == "2G_PS" || Technology == "3G_CS" || Technology == "3G_PS")
                {
                    Worst_Cell_Count2 = Convert.ToDouble((Single_Node_Table_NearContractual.Rows[k - 1]).ItemArray[4]);
                    Level2 = Convert.ToDouble((Single_Node_Table_NearContractual.Rows[k - 1]).ItemArray[2]);
                }
                if (Technology == "4G")
                {
                    Worst_Cell_Count2 = Convert.ToDouble((Single_Node_Table_NearContractual.Rows[k - 1]).ItemArray[4]);
                    Level2 = Convert.ToDouble((Single_Node_Table_NearContractual.Rows[k - 1]).ItemArray[2]);
                }

                if (Level2 == 1 && checkBox6.Checked == true)
                {
                    chart2.Series[0].Points.AddXY(dt2, Worst_Cell_Count2);
                }
                if (Level2 == 2 && checkBox7.Checked == true)
                {
                    chart2.Series[1].Points.AddXY(dt2, Worst_Cell_Count2);
                }
                if (Level2 == 3 && checkBox8.Checked == true)
                {
                    chart2.Series[2].Points.AddXY(dt2, Worst_Cell_Count2);
                }
                if (Level2 == 4 && checkBox9.Checked == true)
                {
                    chart2.Series[3].Points.AddXY(dt2, Worst_Cell_Count2);
                }
                if (Level2 == 5 && checkBox10.Checked == true)
                {
                    chart2.Series[4].Points.AddXY(dt2, Worst_Cell_Count2);
                }
                Max_Y2 = chart2.ChartAreas[0].AxisY.Maximum;
                if (k == 1)
                {
                    Min_X = dt2;
                }
                y2 = k;
            }
            Max_X = Convert.ToDateTime((Single_Node_Table_NearContractual.Rows[y2 - 1]).ItemArray[0]);

            for (int k = 1; k <= Single_Node_Table_Contractual1.Rows.Count; k++)
            {
                DateTime dt1 = Convert.ToDateTime((Single_Node_Table_Contractual1.Rows[k - 1]).ItemArray[0]);
                dt1 = dt1.AddHours(23);
                double Level = 0;
                if (Technology == "2G_CS" || Technology == "2G_PS" || Technology == "3G_CS" || Technology == "3G_PS")
                {
                    Score_Cell_Sum = Convert.ToDouble((Single_Node_Table_Contractual1.Rows[k - 1]).ItemArray[4]);
                    Level = Convert.ToDouble((Single_Node_Table_Contractual1.Rows[k - 1]).ItemArray[3]);

                }
                if (Technology == "4G")
                {
                    Score_Cell_Sum = Convert.ToDouble((Single_Node_Table_Contractual1.Rows[k - 1]).ItemArray[4]);
                    Level = Convert.ToDouble((Single_Node_Table_Contractual1.Rows[k - 1]).ItemArray[3]);
                }


                if (k == 1)
                {
                    Score_Cell_Sum1 = 0;
                }
                if (k > 1)
                {
                    DateTime dt_old = Convert.ToDateTime((Single_Node_Table_Contractual1.Rows[k - 2]).ItemArray[0]);
                    dt_old = dt_old.AddHours(23);
                    if (dt1 != dt_old)
                    {
                        Score_Cell_Sum1 = 0;
                    }
                }

                if (Level == 1 && checkBox6.Checked == true)
                {
                    Score_Cell_Sum1 = Score_Cell_Sum1 + Score_Cell_Sum;
                }
                if (Level == 2 && checkBox7.Checked == true)
                {
                    Score_Cell_Sum1 = Score_Cell_Sum1 + Score_Cell_Sum;
                }
                if (Level == 3 && checkBox8.Checked == true)
                {
                    Score_Cell_Sum1 = Score_Cell_Sum1 + Score_Cell_Sum;
                }
                if (Level == 4 && checkBox9.Checked == true)
                {
                    Score_Cell_Sum1 = Score_Cell_Sum1 + Score_Cell_Sum;
                }
                if (Level == 5 && checkBox10.Checked == true)
                {
                    Score_Cell_Sum1 = Score_Cell_Sum1 + Score_Cell_Sum;
                }

                chart3.Series[0].Points.AddXY(dt1, Score_Cell_Sum1);

            }

            for (int k = 1; k <= Single_Node_Table_NearContractual1.Rows.Count; k++)
            {
                DateTime dt1 = Convert.ToDateTime((Single_Node_Table_NearContractual1.Rows[k - 1]).ItemArray[0]);
                dt1 = dt1.AddHours(23);
                double Level = 0;
                if (Technology == "2G_CS" || Technology == "2G_PS" || Technology == "3G_CS" || Technology == "3G_PS")
                {
                    Score_Cell_Sum = Convert.ToDouble((Single_Node_Table_NearContractual1.Rows[k - 1]).ItemArray[4]);
                    Level = Convert.ToDouble((Single_Node_Table_NearContractual1.Rows[k - 1]).ItemArray[3]);

                }
                if (Technology == "4G")
                {
                    Score_Cell_Sum = Convert.ToDouble((Single_Node_Table_NearContractual1.Rows[k - 1]).ItemArray[4]);
                    Level = Convert.ToDouble((Single_Node_Table_NearContractual1.Rows[k - 1]).ItemArray[3]);

                }

                if (k == 1)
                {
                    Score_Cell_Sum1 = 0;
                }
                if (k > 1)
                {
                    DateTime dt_old = Convert.ToDateTime((Single_Node_Table_NearContractual1.Rows[k - 2]).ItemArray[0]);
                    dt_old = dt_old.AddHours(23);
                    if (dt1 != dt_old)
                    {
                        Score_Cell_Sum1 = 0;
                    }
                }

                if (Level == 1 && checkBox6.Checked == true)
                {
                    Score_Cell_Sum1 = Score_Cell_Sum1 + Score_Cell_Sum;
                }
                if (Level == 2 && checkBox7.Checked == true)
                {
                    Score_Cell_Sum1 = Score_Cell_Sum1 + Score_Cell_Sum;
                }
                if (Level == 3 && checkBox8.Checked == true)
                {
                    Score_Cell_Sum1 = Score_Cell_Sum1 + Score_Cell_Sum;
                }
                if (Level == 4 && checkBox9.Checked == true)
                {
                    Score_Cell_Sum1 = Score_Cell_Sum1 + Score_Cell_Sum;
                }
                if (Level == 5 && checkBox10.Checked == true)
                {
                    Score_Cell_Sum1 = Score_Cell_Sum1 + Score_Cell_Sum;
                }

                chart4.Series[0].Points.AddXY(dt1, Score_Cell_Sum1);

            }



        }







        // Owner Selection
        void comboBox5_MouseWheel(object sender, MouseEventArgs e)
        {
            ((HandledMouseEventArgs)e).Handled = true;
        }
        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox5.MouseWheel += new MouseEventHandler(comboBox5_MouseWheel);
            Region_Show = 0;
            label1.BackColor = Color.PaleGoldenrod;
            label2.BackColor = Color.PaleGoldenrod;
            label4.BackColor = Color.PaleGoldenrod;
            label5.BackColor = Color.Yellow;
            Owner_Selection_Time = DateTime.Now;

            string Owner = comboBox5.SelectedItem.ToString();
            string Single_Owner_List_Quary_Contractual = "";
            string Single_Owner_List_Quary_NearContractual = "";
            string Single_Owner_Misssing_Quary_Contractual = "";
            string Single_Owner_Missing_Quary_NearContractual = "";
            if (Technology == "2G_CS")
            {
                Single_Owner_List_Quary_Contractual = "select [Date], [Owner], [Level], [Status], count(Level) as 'Worst_Count' from(select[Contractual_WPC_2G_CS].[Date], [Contractual_WPC_2G_CS].[BSC], [Contractual_WPC_2G_CS].[Vendor], [Contractual_WPC_2G_CS].[Level], [Contractual_WPC_2G_CS].[Effeciency Index], [Contractual_WPC_2G_CS].[Status],[Owner_TBL].[Owner] from[Contractual_WPC_2G_CS] inner join[Owner_TBL] on[Contractual_WPC_2G_CS].[BSC]=[Node]) tbl where Owner='" + Owner + "' and Status = 'Contractual WPC' group by[Date], [Owner], [Level], [Status] order by Date";
                Single_Owner_List_Quary_NearContractual = "select [Date], [Owner], [Level], [Status], count(Level) as 'Worst_Count' from(select[Contractual_WPC_2G_CS].[Date], [Contractual_WPC_2G_CS].[BSC], [Contractual_WPC_2G_CS].[Vendor], [Contractual_WPC_2G_CS].[Level], [Contractual_WPC_2G_CS].[Effeciency Index], [Contractual_WPC_2G_CS].[Status], [Owner_TBL].[Owner] from[Contractual_WPC_2G_CS] inner join[Owner_TBL] on[Contractual_WPC_2G_CS].[BSC]=[Node]) tbl where Owner='" + Owner + "' and Status = 'Near Contractual WPC' group by[Date], [Owner], [Level], [Status] order by Date";
                Single_Owner_Misssing_Quary_Contractual = "select[Date], [Owner], [Status], sum([Effeciency Index])*1.1 as 'Missing Score' from(select[Contractual_WPC_2G_CS].[Date], [Contractual_WPC_2G_CS].[BSC], [Contractual_WPC_2G_CS].[Vendor], [Contractual_WPC_2G_CS].[Level], [Contractual_WPC_2G_CS].[Effeciency Index], [Contractual_WPC_2G_CS].[Status], [Owner_TBL].[Owner] from [Contractual_WPC_2G_CS] inner join[Owner_TBL] on [Contractual_WPC_2G_CS].[BSC]=[Node]) tbl where Owner='" + Owner + "' and Status = 'Contractual WPC' group by[Date], [Owner], [Status] order by Date";
                Single_Owner_Missing_Quary_NearContractual = "select[Date], [Owner], [Status], sum([Effeciency Index]) as 'Missing Score' from(select[Contractual_WPC_2G_CS].[Date], [Contractual_WPC_2G_CS].[BSC], [Contractual_WPC_2G_CS].[Vendor], [Contractual_WPC_2G_CS].[Level], [Contractual_WPC_2G_CS].[Effeciency Index], [Contractual_WPC_2G_CS].[Status],  [Contractual_WPC_2G_CS].[QIXP], [Owner_TBL].[Owner] from [Contractual_WPC_2G_CS] inner join[Owner_TBL] on [Contractual_WPC_2G_CS].[BSC]=[Node]) tbl where Owner='" + Owner + "' and Status = 'Near Contractual WPC' and QIxP<100 group by[Date], [Owner], [Status] order by Date";
            }
            if (Technology == "2G_PS")
            {
                Single_Owner_List_Quary_Contractual = "select [Date], [Owner], [Level], [Status], count(Level) as 'Worst_Count' from(select[Contractual_WPC_2G_PS].[Date], [Contractual_WPC_2G_PS].[BSC], [Contractual_WPC_2G_PS].[Vendor], [Contractual_WPC_2G_PS].[Level], [Contractual_WPC_2G_PS].[Effeciency Index], [Contractual_WPC_2G_PS].[Status],[Owner_TBL].[Owner] from[Contractual_WPC_2G_PS] inner join[Owner_TBL] on[Contractual_WPC_2G_PS].[BSC]=[Node]) tbl where Owner='" + Owner + "' and Status = 'Contractual WPC' group by[Date], [Owner], [Level], [Status] order by Date";
                Single_Owner_List_Quary_NearContractual = "select [Date], [Owner], [Level], [Status], count(Level) as 'Worst_Count' from(select[Contractual_WPC_2G_PS].[Date], [Contractual_WPC_2G_PS].[BSC], [Contractual_WPC_2G_PS].[Vendor], [Contractual_WPC_2G_PS].[Level], [Contractual_WPC_2G_PS].[Effeciency Index], [Contractual_WPC_2G_PS].[Status], [Owner_TBL].[Owner] from[Contractual_WPC_2G_PS] inner join[Owner_TBL] on[Contractual_WPC_2G_PS].[BSC]=[Node]) tbl where Owner='" + Owner + "' and Status = 'Near WPC' group by[Date], [Owner], [Level], [Status] order by Date";
                Single_Owner_Misssing_Quary_Contractual = "select[Date], [Owner], [Status], sum([Effeciency Index])*1.1 as 'Missing Score' from(select[Contractual_WPC_2G_PS].[Date], [Contractual_WPC_2G_PS].[BSC], [Contractual_WPC_2G_PS].[Vendor], [Contractual_WPC_2G_PS].[Level], [Contractual_WPC_2G_PS].[Effeciency Index], [Contractual_WPC_2G_PS].[Status], [Owner_TBL].[Owner] from [Contractual_WPC_2G_PS] inner join[Owner_TBL] on [Contractual_WPC_2G_PS].[BSC]=[Node]) tbl where Owner='" + Owner + "' and Status = 'Contractual WPC' group by[Date], [Owner], [Status] order by Date";
                Single_Owner_Missing_Quary_NearContractual = "select[Date], [Owner], [Status], sum([Effeciency Index]) as 'Missing Score' from(select[Contractual_WPC_2G_PS].[Date], [Contractual_WPC_2G_PS].[BSC], [Contractual_WPC_2G_PS].[Vendor], [Contractual_WPC_2G_PS].[Level], [Contractual_WPC_2G_PS].[Effeciency Index], [Contractual_WPC_2G_PS].[Status],  [Contractual_WPC_2G_PS].[QIXP], [Owner_TBL].[Owner] from [Contractual_WPC_2G_PS] inner join[Owner_TBL] on [Contractual_WPC_2G_PS].[BSC]=[Node]) tbl where Owner='" + Owner + "' and Status = 'Near WPC' and QIxP<100 group by[Date], [Owner], [Status] order by Date";
            }
            if (Technology == "3G_CS")
            {
                Single_Owner_List_Quary_Contractual = "select [Date], [Owner], [Level], [Status], count(Level) as 'Worst_Count' from(select[Contractual_WPC_3G_CS].[Date], [Contractual_WPC_3G_CS].[RNC], [Contractual_WPC_3G_CS].[Vendor], [Contractual_WPC_3G_CS].[Level], [Contractual_WPC_3G_CS].[Effeciency Index], [Contractual_WPC_3G_CS].[Status], [Owner_TBL].[Owner] from[Contractual_WPC_3G_CS] inner join[Owner_TBL] on[Contractual_WPC_3G_CS].[RNC]=[Node]) tbl where Owner='" + Owner + "' and Status = 'Contractual WPC' group by[Date], [Owner], [Level], [Status] order by Date";
                Single_Owner_List_Quary_NearContractual = "select [Date], [Owner], [Level], [Status], count(Level) as 'Worst_Count' from(select[Contractual_WPC_3G_CS].[Date], [Contractual_WPC_3G_CS].[RNC], [Contractual_WPC_3G_CS].[Vendor], [Contractual_WPC_3G_CS].[Level], [Contractual_WPC_3G_CS].[Effeciency Index], [Contractual_WPC_3G_CS].[Status], [Owner_TBL].[Owner] from[Contractual_WPC_3G_CS] inner join[Owner_TBL] on[Contractual_WPC_3G_CS].[RNC]=[Node]) tbl where Owner='" + Owner + "' and Status = 'Near Contractual WPC' group by[Date], [Owner], [Level], [Status] order by Date";
                Single_Owner_Misssing_Quary_Contractual = "select[Date], [Owner], [Status], sum([Effeciency Index])*1.1 as 'Missing Score' from(select[Contractual_WPC_3G_CS].[Date], [Contractual_WPC_3G_CS].[RNC], [Contractual_WPC_3G_CS].[Vendor], [Contractual_WPC_3G_CS].[Level], [Contractual_WPC_3G_CS].[Effeciency Index], [Contractual_WPC_3G_CS].[Status], [Owner_TBL].[Owner] from [Contractual_WPC_3G_CS] inner join[Owner_TBL] on [Contractual_WPC_3G_CS].[RNC]=[Node]) tbl where Owner='" + Owner + "' and Status = 'Contractual WPC' group by[Date], [Owner], [Status] order by Date";
                Single_Owner_Missing_Quary_NearContractual = "select[Date], [Owner], [Status], sum([Effeciency Index]) as 'Missing Score' from(select[Contractual_WPC_3G_CS].[Date], [Contractual_WPC_3G_CS].[RNC], [Contractual_WPC_3G_CS].[Vendor], [Contractual_WPC_3G_CS].[Level], [Contractual_WPC_3G_CS].[Effeciency Index], [Contractual_WPC_3G_CS].[Status],  [Contractual_WPC_3G_CS].[QIXP], [Owner_TBL].[Owner] from [Contractual_WPC_3G_CS] inner join[Owner_TBL] on [Contractual_WPC_3G_CS].[RNC]=[Node]) tbl where Owner='" + Owner + "' and Status = 'Near Contractual WPC' and QIxP<100 group by[Date], [Owner], [Status] order by Date";
            }
            if (Technology == "3G_PS")
            {
                Single_Owner_List_Quary_Contractual = "select [Date], [Owner], [Level], [Status], count(Level) as 'Worst_Count' from(select[Contractual_WPC_3G_PS].[Date], [Contractual_WPC_3G_PS].[RNC], [Contractual_WPC_3G_PS].[Vendor], [Contractual_WPC_3G_PS].[Level], [Contractual_WPC_3G_PS].[Effeciency Index], [Contractual_WPC_3G_PS].[Status], [Owner_TBL].[Owner] from[Contractual_WPC_3G_PS] inner join[Owner_TBL] on[Contractual_WPC_3G_PS].[RNC]=[Node]) tbl where Owner='" + Owner + "' and Status = 'Contractual WPC' group by[Date], [Owner], [Level], [Status] order by Date";
                Single_Owner_List_Quary_NearContractual = "select [Date], [Owner], [Level], [Status], count(Level) as 'Worst_Count' from(select[Contractual_WPC_3G_PS].[Date], [Contractual_WPC_3G_PS].[RNC], [Contractual_WPC_3G_PS].[Vendor], [Contractual_WPC_3G_PS].[Level], [Contractual_WPC_3G_PS].[Effeciency Index], [Contractual_WPC_3G_PS].[Status], [Owner_TBL].[Owner] from[Contractual_WPC_3G_PS] inner join[Owner_TBL] on[Contractual_WPC_3G_PS].[RNC]=[Node]) tbl where Owner='" + Owner + "' and Status = 'Near Contractual WPC' group by[Date], [Owner], [Level], [Status] order by Date";
                Single_Owner_Misssing_Quary_Contractual = "select[Date], [Owner], [Status], sum([Effeciency Index])*1.1 as 'Missing Score' from(select[Contractual_WPC_3G_PS].[Date], [Contractual_WPC_3G_PS].[RNC], [Contractual_WPC_3G_PS].[Vendor], [Contractual_WPC_3G_PS].[Level], [Contractual_WPC_3G_PS].[Effeciency Index], [Contractual_WPC_3G_PS].[Status], [Owner_TBL].[Owner] from [Contractual_WPC_3G_PS] inner join[Owner_TBL] on [Contractual_WPC_3G_PS].[RNC]=[Node]) tbl where Owner='" + Owner + "' and Status = 'Contractual WPC' group by[Date], [Owner], [Status] order by Date";
                Single_Owner_Missing_Quary_NearContractual = "select[Date], [Owner], [Status], sum([Effeciency Index]) as 'Missing Score' from(select[Contractual_WPC_3G_PS].[Date], [Contractual_WPC_3G_PS].[RNC], [Contractual_WPC_3G_PS].[Vendor], [Contractual_WPC_3G_PS].[Level], [Contractual_WPC_3G_PS].[Effeciency Index], [Contractual_WPC_3G_PS].[Status],  [Contractual_WPC_3G_PS].[QIXP], [Owner_TBL].[Owner] from [Contractual_WPC_3G_PS] inner join[Owner_TBL] on [Contractual_WPC_3G_PS].[RNC]=[Node]) tbl where Owner='" + Owner + "' and Status = 'Near Contractual WPC' and QIxP<100 group by[Date], [Owner], [Status] order by Date";
            }
            if (Technology == "4G")
            {
                Single_Owner_List_Quary_Contractual = "select [Date], [Owner], [Level], [Status], count(Level) as 'Worst_Count' from(select[Contractual_WPC_4G].[Date], [Contractual_WPC_4G].[Province], [Contractual_WPC_4G].[Vendor], [Contractual_WPC_4G].[Level], [Contractual_WPC_4G].[Effeciency Index], [Contractual_WPC_4G].[Status], [Owner_TBL].[Owner] from[Contractual_WPC_4G] inner join[Owner_TBL] on[Contractual_WPC_4G].[RNC]=[Node]) tbl where Owner='" + Owner + "' and Status = 'Contractual WPC' group by[Date], [Owner], [Level], [Status] order by Date";
                Single_Owner_List_Quary_NearContractual = "select [Date], [Owner], [Level], [Status], count(Level) as 'Worst_Count' from(select[Contractual_WPC_4G].[Date], [Contractual_WPC_4G].[Province], [Contractual_WPC_4G].[Vendor], [Contractual_WPC_4G].[Level], [Contractual_WPC_4G].[Effeciency Index], [Contractual_WPC_4G].[Status], [Owner_TBL].[Owner] from[Contractual_WPC_4G] inner join[Owner_TBL] on[Contractual_WPC_4G].[RNC]=[Node]) tbl where Owner='" + Owner + "' and Status = 'Near Contractual WPC' group by[Date], [Owner], [Level], [Status] order by Date";
                Single_Owner_Misssing_Quary_Contractual = "select[Date], [Owner], [Status], sum([Effeciency Index])*1.1 as 'Missing Score' from(select[Contractual_WPC_4G].[Date], [Contractual_WPC_4G].[Province], [Contractual_WPC_4G].[Vendor], [Contractual_WPC_4G].[Level], [Contractual_WPC_4G].[Effeciency Index], [Contractual_WPC_4G].[Status], [Owner_TBL].[Owner] from [Contractual_WPC_4G] inner join[Owner_TBL] on [Contractual_WPC_4G].[RNC]=[Node]) tbl where Owner='" + Owner + "' and Status = 'Contractual WPC' group by[Date], [Owner], [Status] order by Date";
                Single_Owner_Missing_Quary_NearContractual = "select[Date], [Owner], [Status], sum([Effeciency Index]) as 'Missing Score' from(select[Contractual_WPC_4G].[Date], [Contractual_WPC_4G].[Province], [Contractual_WPC_4G].[Vendor], [Contractual_WPC_4G].[Level], [Contractual_WPC_4G].[Effeciency Index], [Contractual_WPC_4G].[Status],  [Contractual_WPC_4G].[QIXP], [Owner_TBL].[Owner] from [Contractual_WPC_4G] inner join[Owner_TBL] on [Contractual_WPC_4G].[RNC]=[Node]) tbl where Owner='" + Owner + "' and Status = 'Near Contractual WPC' and QIxP<100 group by[Date], [Owner], [Status] order by Date";
            }

            // Worst Cells Count in Contractual WPC
            SqlCommand Single_Owner_List_Quary_Contractual1 = new SqlCommand(Single_Owner_List_Quary_Contractual, connection);
            Single_Owner_List_Quary_Contractual1.ExecuteNonQuery();
            DataTable Single_Owner_Table_Contractual = new DataTable();
            SqlDataAdapter dataAdapter_Contractual = new SqlDataAdapter(Single_Owner_List_Quary_Contractual1);
            dataAdapter_Contractual.Fill(Single_Owner_Table_Contractual);

            // Worst Cells Count in Near Contractual WPC
            SqlCommand Single_Owner_List_Quary_NearContractual1 = new SqlCommand(Single_Owner_List_Quary_NearContractual, connection);
            Single_Owner_List_Quary_NearContractual1.ExecuteNonQuery();
            DataTable Single_Owner_Table_NearContractual = new DataTable();
            SqlDataAdapter dataAdapter_NearContractual = new SqlDataAdapter(Single_Owner_List_Quary_NearContractual1);
            dataAdapter_NearContractual.Fill(Single_Owner_Table_NearContractual);

            // Worst Cells Missing Score in Contractual WPC
            SqlCommand Single_Owner_Missing_Quary_Contractual1 = new SqlCommand(Single_Owner_Misssing_Quary_Contractual, connection);
            Single_Owner_Missing_Quary_Contractual1.ExecuteNonQuery();
            DataTable Single_Owner_Table_Contractual1 = new DataTable();
            SqlDataAdapter dataAdapter_Contractual1 = new SqlDataAdapter(Single_Owner_Missing_Quary_Contractual1);
            dataAdapter_Contractual1.Fill(Single_Owner_Table_Contractual1);

            // Worst Cells Missing Score in Near Contractual WPC
            SqlCommand Single_Owner_Missing_Quary_NearContractual1 = new SqlCommand(Single_Owner_Missing_Quary_NearContractual, connection);
            Single_Owner_Missing_Quary_NearContractual1.ExecuteNonQuery();
            DataTable Single_Owner_Table_NearContractual1 = new DataTable();
            SqlDataAdapter dataAdapter_NearContractual1 = new SqlDataAdapter(Single_Owner_Missing_Quary_NearContractual1);
            dataAdapter_NearContractual1.Fill(Single_Owner_Table_NearContractual1);




            chart1.Series.Clear();
            chart1.Titles.Clear();

            Series newSeries1 = new Series();
            chart1.Series.Add(newSeries1);
            newSeries1.IsXValueIndexed = false;
            chart1.Series[0].ChartType = SeriesChartType.Line;
            chart1.Series[0].Color = Color.Red;
            chart1.Series[0].BorderWidth = 3;
            chart1.ChartAreas[0].AxisX.Interval = 5;
            chart1.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart1.Series[0].EmptyPointStyle.Color = Color.Transparent;
            chart1.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";
            chart1.Series[0].ToolTip = "#VALX [#VALY]";
            chart1.Series[0].IsValueShownAsLabel = false;
            chart1.Series[0].LegendText = "Level 1";
            chart1.Legends["Legend1"].Docking = Docking.Bottom;
            newSeries1.MarkerStyle = MarkerStyle.Circle;
            newSeries1.MarkerSize = 6;


            Series newSeries2 = new Series();
            chart1.Series.Add(newSeries2);
            newSeries2.IsXValueIndexed = false;
            chart1.Series[1].ChartType = SeriesChartType.Line;
            chart1.Series[1].Color = Color.Orange;
            chart1.Series[1].BorderWidth = 3;
            chart1.Series[1].EmptyPointStyle.Color = Color.Transparent;
            chart1.Series[1].ToolTip = "#VALX [#VALY]";
            chart1.Series[1].IsValueShownAsLabel = false;
            chart1.Series[1].LegendText = "Level 2";
            newSeries2.MarkerStyle = MarkerStyle.Circle;
            newSeries2.MarkerSize = 6;


            Series newSeries3 = new Series();
            chart1.Series.Add(newSeries3);
            newSeries3.IsXValueIndexed = false;
            chart1.Series[2].ChartType = SeriesChartType.Line;
            chart1.Series[2].Color = Color.Yellow;
            chart1.Series[2].BorderWidth = 3;
            chart1.Series[2].EmptyPointStyle.Color = Color.Transparent;
            chart1.Series[2].ToolTip = "#VALX [#VALY]";
            chart1.Series[2].IsValueShownAsLabel = false;
            chart1.Series[2].LegendText = "Level 3";
            newSeries3.MarkerStyle = MarkerStyle.Circle;
            newSeries3.MarkerSize = 6;


            Series newSeries4 = new Series();
            chart1.Series.Add(newSeries4);
            newSeries4.IsXValueIndexed = false;
            chart1.Series[3].ChartType = SeriesChartType.Line;
            chart1.Series[3].Color = Color.Blue;
            chart1.Series[3].BorderWidth = 3;
            chart1.Series[3].EmptyPointStyle.Color = Color.Transparent;
            chart1.Series[3].ToolTip = "#VALX [#VALY]";
            chart1.Series[3].IsValueShownAsLabel = false;
            chart1.Series[3].LegendText = "Level 4";
            newSeries4.MarkerStyle = MarkerStyle.Circle;
            newSeries4.MarkerSize = 6;


            Series newSeries5 = new Series();
            chart1.Series.Add(newSeries5);
            newSeries5.IsXValueIndexed = false;
            chart1.Series[4].ChartType = SeriesChartType.Line;
            chart1.Series[4].Color = Color.Green;
            chart1.Series[4].BorderWidth = 3;
            chart1.Series[4].EmptyPointStyle.Color = Color.Transparent;
            chart1.Series[4].ToolTip = "#VALX [#VALY]";
            chart1.Series[4].IsValueShownAsLabel = false;
            chart1.Series[4].LegendText = "Level 5";
            newSeries5.MarkerStyle = MarkerStyle.Circle;
            newSeries5.MarkerSize = 6;


            chart2.Series.Clear();
            chart2.Titles.Clear();

            Series newSeries6 = new Series();
            chart2.Series.Add(newSeries6);
            newSeries6.IsXValueIndexed = false;
            chart2.Series[0].ChartType = SeriesChartType.Line;
            chart2.Series[0].Color = Color.Red;
            chart2.Series[0].BorderWidth = 3;
            chart2.ChartAreas[0].AxisX.Interval = 5;
            chart2.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart2.Series[0].EmptyPointStyle.Color = Color.Transparent;
            chart2.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";
            chart2.Series[0].ToolTip = "#VALX [#VALY]";
            chart2.Series[0].IsValueShownAsLabel = false;
            chart2.Series[0].LegendText = "Level 1";
            chart2.Legends["Legend1"].Docking = Docking.Bottom;
            newSeries6.MarkerStyle = MarkerStyle.Circle;
            newSeries6.MarkerSize = 6;


            Series newSeries7 = new Series();
            chart2.Series.Add(newSeries7);
            newSeries7.IsXValueIndexed = false;
            chart2.Series[1].ChartType = SeriesChartType.Line;
            chart2.Series[1].Color = Color.Orange;
            chart2.Series[1].BorderWidth = 3;
            chart2.Series[1].EmptyPointStyle.Color = Color.Transparent;
            chart2.Series[1].ToolTip = "#VALX [#VALY]";
            chart2.Series[1].IsValueShownAsLabel = false;
            chart2.Series[1].LegendText = "Level 2";
            newSeries7.MarkerStyle = MarkerStyle.Circle;
            newSeries7.MarkerSize = 6;


            Series newSeries8 = new Series();
            chart2.Series.Add(newSeries8);
            newSeries8.IsXValueIndexed = false;
            chart2.Series[2].ChartType = SeriesChartType.Line;
            chart2.Series[2].Color = Color.Yellow;
            chart2.Series[2].BorderWidth = 3;
            chart2.Series[2].EmptyPointStyle.Color = Color.Transparent;
            chart2.Series[2].ToolTip = "#VALX [#VALY]";
            chart2.Series[2].IsValueShownAsLabel = false;
            chart2.Series[2].LegendText = "Level 3";
            newSeries8.MarkerStyle = MarkerStyle.Circle;
            newSeries8.MarkerSize = 6;


            Series newSeries9 = new Series();
            chart2.Series.Add(newSeries9);
            newSeries9.IsXValueIndexed = false;
            chart2.Series[3].ChartType = SeriesChartType.Line;
            chart2.Series[3].Color = Color.Blue;
            chart2.Series[3].BorderWidth = 3;
            chart2.Series[3].EmptyPointStyle.Color = Color.Transparent;
            chart2.Series[3].ToolTip = "#VALX [#VALY]";
            chart2.Series[3].IsValueShownAsLabel = false;
            chart2.Series[3].LegendText = "Level 4";
            newSeries9.MarkerStyle = MarkerStyle.Circle;
            newSeries9.MarkerSize = 6;


            Series newSeries10 = new Series();
            chart2.Series.Add(newSeries10);
            newSeries10.IsXValueIndexed = false;
            chart2.Series[4].ChartType = SeriesChartType.Line;
            chart2.Series[4].Color = Color.Green;
            chart2.Series[4].BorderWidth = 3;
            chart2.Series[4].EmptyPointStyle.Color = Color.Transparent;
            chart2.Series[4].ToolTip = "#VALX [#VALY]";
            chart2.Series[4].IsValueShownAsLabel = false;
            chart2.Series[4].LegendText = "Level 5";
            newSeries10.MarkerStyle = MarkerStyle.Circle;
            newSeries10.MarkerSize = 6;




            Title title1 = chart1.Titles.Add("Number of '" + Owner + "' Worst Cells (Contractual)");
            title1.Font = new System.Drawing.Font("Arial", 12, FontStyle.Regular);
            Title title2 = chart2.Titles.Add("Number of '" + Owner + "' Worst Cells (Near Contractual)");
            title2.Font = new System.Drawing.Font("Arial", 12, FontStyle.Regular);




            chart3.Series.Clear();
            chart3.Titles.Clear();

            Series newSeries11 = new Series();
            chart3.Series.Add(newSeries11);
            newSeries11.IsXValueIndexed = false;
            chart3.Series[0].ChartType = SeriesChartType.Line;
            chart3.Series[0].Color = Color.Brown;
            chart3.Series[0].BorderWidth = 3;
            chart3.ChartAreas[0].AxisX.Interval = 5;
            chart3.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart3.Series[0].EmptyPointStyle.Color = Color.Transparent;
            chart3.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";
            chart3.Series[0].ToolTip = "#VALX [#VALY]";
            chart3.Series[0].IsValueShownAsLabel = false;
            chart3.Series[0].IsVisibleInLegend = false;
            newSeries11.MarkerStyle = MarkerStyle.Circle;
            newSeries11.MarkerSize = 6;


            chart4.Series.Clear();
            chart4.Titles.Clear();

            Series newSeries12 = new Series();
            chart4.Series.Add(newSeries12);
            newSeries12.IsXValueIndexed = false;
            chart4.Series[0].ChartType = SeriesChartType.Line;
            chart4.Series[0].Color = Color.Brown;
            chart4.Series[0].BorderWidth = 3;
            chart4.ChartAreas[0].AxisX.Interval = 5;
            chart4.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart4.Series[0].EmptyPointStyle.Color = Color.Transparent;
            chart4.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";
            chart4.Series[0].ToolTip = "#VALX [#VALY]";
            chart4.Series[0].IsValueShownAsLabel = false;
            chart4.Series[0].IsVisibleInLegend = false;
            newSeries12.MarkerStyle = MarkerStyle.Circle;
            newSeries12.MarkerSize = 6;

            Title title3 = chart3.Titles.Add("Sum of '" + Owner + "' Missing Score (Contractual)");
            title3.Font = new System.Drawing.Font("Arial", 12, FontStyle.Regular);
            Title title4 = chart4.Titles.Add("Sum of '" + Owner + "' Missing Score (Near Contractual)");
            title4.Font = new System.Drawing.Font("Arial", 12, FontStyle.Regular);


            int y1 = 0;
            for (int k = 1; k <= Single_Owner_Table_Contractual.Rows.Count; k++)
            {
                DateTime dt1 = Convert.ToDateTime((Single_Owner_Table_Contractual.Rows[k - 1]).ItemArray[0]);
                dt1 = dt1.AddHours(23);
                double Worst_Cell_Count1 = 0;
                double Level1 = 0;
                if (Technology == "2G_CS" || Technology == "2G_PS" || Technology == "3G_CS" || Technology == "3G_PS")
                {
                    Worst_Cell_Count1 = Convert.ToDouble((Single_Owner_Table_Contractual.Rows[k - 1]).ItemArray[4]);
                    Level1 = Convert.ToDouble((Single_Owner_Table_Contractual.Rows[k - 1]).ItemArray[2]);
                }
                if (Technology == "4G")
                {
                    Worst_Cell_Count1 = Convert.ToDouble((Single_Owner_Table_Contractual.Rows[k - 1]).ItemArray[4]);
                    Level1 = Convert.ToDouble((Single_Owner_Table_Contractual.Rows[k - 1]).ItemArray[2]);
                }

                if (Level1 == 1 && checkBox6.Checked == true)
                {
                    chart1.Series[0].Points.AddXY(dt1, Worst_Cell_Count1);
                }
                if (Level1 == 2 && checkBox7.Checked == true)
                {
                    chart1.Series[1].Points.AddXY(dt1, Worst_Cell_Count1);
                }
                if (Level1 == 3 && checkBox8.Checked == true)
                {
                    chart1.Series[2].Points.AddXY(dt1, Worst_Cell_Count1);
                }
                if (Level1 == 4 && checkBox9.Checked == true)
                {
                    chart1.Series[3].Points.AddXY(dt1, Worst_Cell_Count1);
                }
                if (Level1 == 5 && checkBox10.Checked == true)
                {
                    chart1.Series[4].Points.AddXY(dt1, Worst_Cell_Count1);
                }
                Max_Y1 = chart1.ChartAreas[0].AxisY.Maximum;
                if (k == 1)
                {
                    Min_X = dt1;
                }

                y1 = k;
            }
            Max_X = Convert.ToDateTime((Single_Owner_Table_Contractual.Rows[y1 - 1]).ItemArray[0]);

            int y2 = 0;
            for (int k = 1; k <= Single_Owner_Table_NearContractual.Rows.Count; k++)
            {

                DateTime dt2 = Convert.ToDateTime((Single_Owner_Table_NearContractual.Rows[k - 1]).ItemArray[0]);
                dt2 = dt2.AddHours(23);
                double Worst_Cell_Count2 = 0;
                double Level2 = 0;
                if (Technology == "2G_CS" || Technology == "2G_PS" || Technology == "3G_CS" || Technology == "3G_PS")
                {
                    Worst_Cell_Count2 = Convert.ToDouble((Single_Owner_Table_NearContractual.Rows[k - 1]).ItemArray[4]);
                    Level2 = Convert.ToDouble((Single_Owner_Table_NearContractual.Rows[k - 1]).ItemArray[2]);
                }
                if (Technology == "4G")
                {
                    Worst_Cell_Count2 = Convert.ToDouble((Single_Owner_Table_NearContractual.Rows[k - 1]).ItemArray[4]);
                    Level2 = Convert.ToDouble((Single_Owner_Table_NearContractual.Rows[k - 1]).ItemArray[2]);
                }

                if (Level2 == 1 && checkBox6.Checked == true)
                {
                    chart2.Series[0].Points.AddXY(dt2, Worst_Cell_Count2);
                }
                if (Level2 == 2 && checkBox7.Checked == true)
                {
                    chart2.Series[1].Points.AddXY(dt2, Worst_Cell_Count2);
                }
                if (Level2 == 3 && checkBox8.Checked == true)
                {
                    chart2.Series[2].Points.AddXY(dt2, Worst_Cell_Count2);
                }
                if (Level2 == 4 && checkBox9.Checked == true)
                {
                    chart2.Series[3].Points.AddXY(dt2, Worst_Cell_Count2);
                }
                if (Level2 == 5 && checkBox10.Checked == true)
                {
                    chart2.Series[4].Points.AddXY(dt2, Worst_Cell_Count2);
                }
                Max_Y2 = chart2.ChartAreas[0].AxisY.Maximum;
                if (k == 1)
                {
                    Min_X = dt2;
                }

                y2 = k;
            }
            Max_X = Convert.ToDateTime((Single_Owner_Table_NearContractual.Rows[y2 - 1]).ItemArray[0]);



            for (int k = 1; k <= Single_Owner_Table_Contractual1.Rows.Count; k++)
            {
                DateTime dt1 = Convert.ToDateTime((Single_Owner_Table_Contractual1.Rows[k - 1]).ItemArray[0]);
                dt1 = dt1.AddHours(23);
                double Score_Cell_Sum = 0;
                if (Technology == "2G_CS" || Technology == "2G_PS" || Technology == "3G_CS" || Technology == "3G_PS")
                {
                    Score_Cell_Sum = Convert.ToDouble((Single_Owner_Table_Contractual1.Rows[k - 1]).ItemArray[3]);
                }
                if (Technology == "4G")
                {
                    Score_Cell_Sum = Convert.ToDouble((Single_Owner_Table_Contractual1.Rows[k - 1]).ItemArray[3]);
                }
                chart3.Series[0].Points.AddXY(dt1, Score_Cell_Sum);
            }

            for (int k = 1; k <= Single_Owner_Table_NearContractual1.Rows.Count; k++)
            {
                DateTime dt1 = Convert.ToDateTime((Single_Owner_Table_NearContractual1.Rows[k - 1]).ItemArray[0]);
                dt1 = dt1.AddHours(23);
                double Score_Cell_Sum = 0;
                if (Technology == "2G_CS" || Technology == "2G_PS" || Technology == "3G_CS" || Technology == "3G_PS")
                {
                    Score_Cell_Sum = Convert.ToDouble((Single_Owner_Table_NearContractual1.Rows[k - 1]).ItemArray[3]);
                }
                if (Technology == "4G")
                {
                    Score_Cell_Sum = Convert.ToDouble((Single_Owner_Table_NearContractual1.Rows[k - 1]).ItemArray[3]);
                }
                chart4.Series[0].Points.AddXY(dt1, Score_Cell_Sum);
            }



        }




        // Region Selection

        void comboBox1_MouseWheel(object sender, MouseEventArgs e)
        {
            ((HandledMouseEventArgs)e).Handled = true;
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Technology == "2G_CS" || Technology == "2G_PS" || Technology == "3G_CS" || Technology == "3G_PS" || Technology == "4G")
            {
                comboBox1.MouseWheel += new MouseEventHandler(comboBox1_MouseWheel);
                progressBar1.Minimum = 0;
                progressBar1.Maximum = 5;

                Region_Show = 1;
                Province_Show = 0;
                Node_Show = 0;
                label1.BackColor = Color.Yellow;
                label2.BackColor = Color.PaleGoldenrod;
                label4.BackColor = Color.PaleGoldenrod;
                label5.BackColor = Color.PaleGoldenrod;
                comboBox2.Items.Clear();
                comboBox4.Items.Clear();
                comboBox5.Items.Clear();
                Region_Selection_Time = DateTime.Now;


                Region = comboBox1.SelectedItem.ToString();

                // Data  Tabe of Last Day
                string Last_Day_List_Quary = "";
                if (Technology == "2G_CS")
                {
                    Last_Day_List_Quary = "select * from  [Contractual_WPC_2G_CS] where Contractor='" + Region + "' and date=(select max(distinct([Date])) from [Contractual_WPC_2G_CS]  where Contractor='" + Region + "')";
                }
                if (Technology == "2G_PS")
                {
                    Last_Day_List_Quary = "select * from  [Contractual_WPC_2G_PS] where Contractor='" + Region + "' and date=(select max(distinct([Date])) from [Contractual_WPC_2G_PS]  where Contractor='" + Region + "')";
                }
                if (Technology == "3G_CS")
                {
                    Last_Day_List_Quary = "select * from  [Contractual_WPC_3G_CS] where Contractor='" + Region + "' and date=(select max(distinct([Date])) from [Contractual_WPC_3G_CS]  where Contractor='" + Region + "')";
                }
                if (Technology == "3G_PS")
                {
                    Last_Day_List_Quary = "select * from  [Contractual_WPC_3G_PS] where Contractor='" + Region + "' and date=(select max(distinct([Date])) from [Contractual_WPC_3G_PS]  where Contractor='" + Region + "')";
                }
                if (Technology == "4G")
                {
                    Last_Day_List_Quary = "select * from  [Contractual_WPC_4G] where Contractor='" + Region + "' and date=(select max(distinct([Date])) from [Contractual_WPC_4G]  where Contractor='" + Region + "')";
                }
                SqlCommand Last_Day_List_Quary1 = new SqlCommand(Last_Day_List_Quary, connection);
                Last_Day_List_Quary1.CommandTimeout = 0;
                Last_Day_List_Quary1.ExecuteNonQuery();
                Last_Day_List = new DataTable();
                SqlDataAdapter Last_Day_List1 = new SqlDataAdapter(Last_Day_List_Quary1);
                Last_Day_List1.Fill(Last_Day_List);


                progressBar1.Value = 1;

                if (Region == "NAK-Nokia")
                {
                    comboBox2.Items.Add("Khorasan Razavi");
                    comboBox2.Items.Add("Kerman");
                    comboBox2.Items.Add("Yazd");
                    comboBox2.Items.Add("Chahar Mahal Va Bakhtiari");
                    comboBox2.Items.Add("Semnan");
                }

                if (Region == "NAK-North")
                {
                    comboBox2.Items.Add("Gilan");
                    comboBox2.Items.Add("Golestan");
                    comboBox2.Items.Add("Mazandaran");
                }
                if (Region == "NAK-Huawei")
                {
                    comboBox2.Items.Add("East Azarbaijan");
                    comboBox2.Items.Add("West Azarbaijan");
                    comboBox2.Items.Add("Khuzestan");
                }

                // Fill The Blank Status in DataBase with "Near Contractual WPC"
                string Update_Quary = "";
                string Update_Quary_CS = "";
                string Update_Quary_PS = "";
                string Node_List_Quary = "";
                string Date_Quary = "";
                string Cell_Quary = "";
                if (Technology == "2G_CS")
                {
                    Node_List_Quary = "select distinct([BSC]) from [Contractual_WPC_2G_CS] where Contractor='" + Region + "' and date=(select max(distinct([Date])) from [Contractual_WPC_2G_CS]  where Contractor='" + Region + "') order by BSC";

                    Date_Quary = "select distinct([Date]) from [Contractual_WPC_2G_CS] where Contractor='" + Region + "' order by Date";
                    SqlCommand Date_Quary1 = new SqlCommand(Date_Quary, connection);
                    Date_Quary1.ExecuteNonQuery();
                    Date_Table1 = new DataTable();
                    SqlDataAdapter Date_Table = new SqlDataAdapter(Date_Quary1);
                    Date_Table.Fill(Date_Table1);

                    Cell_Quary = "select Cell from [Contractual_WPC_2G_CS] where Contractor='" + Region + "' and date=(select max(distinct([Date])) from [Contractual_WPC_2G_CS] where Contractor='" + Region + "')";
                    SqlCommand Cell_Quary1 = new SqlCommand(Cell_Quary, connection);
                    Cell_Quary1.ExecuteNonQuery();
                    DataTable Cell_Table1 = new DataTable();
                    SqlDataAdapter Cell_Table = new SqlDataAdapter(Cell_Quary1);
                    Cell_Table.Fill(Cell_Table1);
                    comboBox6.Items.Clear();
                    for (int k = 0; k < Cell_Table1.Rows.Count; k++)
                    {
                        string cell = (Cell_Table1.Rows[k]).ItemArray[0].ToString();
                        comboBox6.Items.Add(cell);
                    }
                }

                if (Technology == "2G_PS")
                {
                    Node_List_Quary = "select distinct([BSC]) from [Contractual_WPC_2G_PS] where Contractor='" + Region + "' and date=(select max(distinct([Date])) from [Contractual_WPC_2G_PS]  where Contractor='" + Region + "') order by BSC";

                    Date_Quary = "select distinct([Date]) from [Contractual_WPC_2G_PS] where Contractor='" + Region + "' order by Date";
                    SqlCommand Date_Quary1 = new SqlCommand(Date_Quary, connection);
                    Date_Quary1.ExecuteNonQuery();
                    Date_Table1 = new DataTable();
                    SqlDataAdapter Date_Table = new SqlDataAdapter(Date_Quary1);
                    Date_Table.Fill(Date_Table1);

                    Cell_Quary = "select Cell from [Contractual_WPC_2G_PS] where Contractor='" + Region + "' and date=(select max(distinct([Date])) from [Contractual_WPC_2G_PS] where Contractor='" + Region + "')";
                    SqlCommand Cell_Quary1 = new SqlCommand(Cell_Quary, connection);
                    Cell_Quary1.ExecuteNonQuery();
                    DataTable Cell_Table1 = new DataTable();
                    SqlDataAdapter Cell_Table = new SqlDataAdapter(Cell_Quary1);
                    Cell_Table.Fill(Cell_Table1);
                    comboBox6.Items.Clear();
                    for (int k = 0; k < Cell_Table1.Rows.Count; k++)
                    {
                        string cell = (Cell_Table1.Rows[k]).ItemArray[0].ToString();
                        comboBox6.Items.Add(cell);
                    }
                }


                if (Technology == "3G_CS")
                {
                    Node_List_Quary = "select distinct([RNC]) from [Contractual_WPC_3G_CS] where Contractor='" + Region + "' and  date=(select max(distinct([Date])) from [Contractual_WPC_3G_CS]  where Contractor='" + Region + "') order by RNC";

                    Date_Quary = "select distinct([Date]) from [Contractual_WPC_3G_CS] where Contractor='" + Region + "' order by Date";
                    SqlCommand Date_Quary1 = new SqlCommand(Date_Quary, connection);
                    Date_Quary1.ExecuteNonQuery();
                    Date_Table1 = new DataTable();
                    SqlDataAdapter Date_Table = new SqlDataAdapter(Date_Quary1);
                    Date_Table.Fill(Date_Table1);

                    Cell_Quary = "select Cell from [Contractual_WPC_3G_CS] where Contractor='" + Region + "' and date=(select max(distinct([Date])) from [Contractual_WPC_3G_CS] where Contractor='" + Region + "')";
                    SqlCommand Cell_Quary1 = new SqlCommand(Cell_Quary, connection);
                    Cell_Quary1.ExecuteNonQuery();
                    DataTable Cell_Table1 = new DataTable();
                    SqlDataAdapter Cell_Table = new SqlDataAdapter(Cell_Quary1);
                    Cell_Table.Fill(Cell_Table1);
                    comboBox6.Items.Clear();
                    for (int k = 0; k < Cell_Table1.Rows.Count; k++)
                    {
                        string cell = (Cell_Table1.Rows[k]).ItemArray[0].ToString();
                        comboBox6.Items.Add(cell);
                    }

                }
                if (Technology == "3G_PS")
                {
                    Node_List_Quary = "select distinct([RNC]) from [Contractual_WPC_3G_PS] where Contractor='" + Region + "' and  date=(select max(distinct([Date])) from [Contractual_WPC_3G_PS] where Contractor='" + Region + "') order by RNC";

                    Date_Quary = "select distinct([Date]) from [Contractual_WPC_3G_PS] where Contractor='" + Region + "' order by Date";
                    SqlCommand Date_Quary1 = new SqlCommand(Date_Quary, connection);
                    Date_Quary1.ExecuteNonQuery();
                    Date_Table1 = new DataTable();
                    SqlDataAdapter Date_Table = new SqlDataAdapter(Date_Quary1);
                    Date_Table.Fill(Date_Table1);

                    Cell_Quary = "select Cell from[Contractual_WPC_3G_PS] where Contractor='" + Region + "' and date=(select max(distinct([Date])) from [Contractual_WPC_3G_PS]  where Contractor='" + Region + "')";
                    SqlCommand Cell_Quary1 = new SqlCommand(Cell_Quary, connection);
                    Cell_Quary1.ExecuteNonQuery();
                    DataTable Cell_Table1 = new DataTable();
                    SqlDataAdapter Cell_Table = new SqlDataAdapter(Cell_Quary1);
                    Cell_Table.Fill(Cell_Table1);
                    comboBox6.Items.Clear();
                    for (int k = 0; k < Cell_Table1.Rows.Count; k++)
                    {
                        string cell = (Cell_Table1.Rows[k]).ItemArray[0].ToString();
                        comboBox6.Items.Add(cell);
                    }

                }
                if (Technology == "4G")
                {
                    Node_List_Quary = "select distinct([RNC]) from [Contractual_WPC_4G] where Contractor='" + Region + "' and  date=(select max(distinct([Date])) from [Contractual_WPC_4G] where Contractor='" + Region + "') order by RNC";


                    Date_Quary = "select distinct([Date]) from [Contractual_WPC_4G] where Contractor='" + Region + "' order by Date";
                    SqlCommand Date_Quary1 = new SqlCommand(Date_Quary, connection);
                    Date_Quary1.ExecuteNonQuery();
                    Date_Table1 = new DataTable();
                    SqlDataAdapter Date_Table = new SqlDataAdapter(Date_Quary1);
                    Date_Table.Fill(Date_Table1);

                    Cell_Quary = "select eNodeB from [Contractual_WPC_4G] where Contractor='" + Region + "' and date=(select max(distinct([Date])) from [Contractual_WPC_4G] where Contractor='" + Region + "')";
                    SqlCommand Cell_Quary1 = new SqlCommand(Cell_Quary, connection);
                    Cell_Quary1.ExecuteNonQuery();
                    DataTable Cell_Table1 = new DataTable();
                    SqlDataAdapter Cell_Table = new SqlDataAdapter(Cell_Quary1);
                    Cell_Table.Fill(Cell_Table1);
                    comboBox6.Items.Clear();
                    for (int k = 0; k < Cell_Table1.Rows.Count; k++)
                    {
                        string cell = (Cell_Table1.Rows[k]).ItemArray[0].ToString();
                        comboBox6.Items.Add(cell);
                    }


                }



                progressBar1.Value = 2;

                // List of Nodes (BSC List or RNC List)
                if (Technology == "2G_CS" || Technology == "2G_PS" || Technology == "3G_CS" || Technology == "3G_PS" || Technology == "4G")
                {
                    SqlCommand Node_List_Quary1 = new SqlCommand(Node_List_Quary, connection);
                    Node_List_Quary1.ExecuteNonQuery();

                    Node_Table = new DataTable();
                    SqlDataAdapter dataAdapter = new SqlDataAdapter(Node_List_Quary1);
                    dataAdapter.Fill(Node_Table);

                }


                comboBox4.Items.Clear();
                for (int k = 1; k <= Node_Table.Rows.Count; k++)
                {
                    string Node_Name = (Node_Table.Rows[k - 1]).ItemArray[0].ToString();
                    if (Node_Name != "")
                    {
                        comboBox4.Items.Add(Node_Name);
                    }
                }


                string Technology1 = "";
                if (Technology == "3G_CS" || Technology == "3G_PS")
                {
                    Technology1 = "3G";
                }
                else
                {
                    Technology1 = Technology;
                }
                if (Technology == "2G_CS" || Technology == "2G_PS")
                {
                    Technology1 = "2G";
                }
                else
                {
                    Technology1 = Technology;
                }

                if (Region == "NAK-Tehran")
                {
                    string Owner_List_Quary = "select distinct([Owner]) from [Owner_TBL] where [Technology] =" + "'" + Technology1 + "'";

                    // List of Owners (BSC or RNC)
                    SqlCommand Owner_List_Quary1 = new SqlCommand(Owner_List_Quary, connection);
                    Owner_List_Quary1.ExecuteNonQuery();

                    DataTable Owner_Table = new DataTable();
                    SqlDataAdapter dataAdapter1 = new SqlDataAdapter(Owner_List_Quary1);
                    dataAdapter1.Fill(Owner_Table);


                    comboBox5.Items.Clear();
                    for (int k = 1; k <= Owner_Table.Rows.Count; k++)
                    {
                        string Owner_Name = (Owner_Table.Rows[k - 1]).ItemArray[0].ToString();
                        if (Owner_Name != "")
                        {
                            comboBox5.Items.Add(Owner_Name);
                        }
                    }

                }

                progressBar1.Value = 3;

                if (Technology == "2G_CS")
                {
                    string BL_Quary = "select * from BL_2G_CS where Contractor='" + Region + "'";
                    SqlCommand BL_Quary1 = new SqlCommand(BL_Quary, connection);
                    BL_Quary1.CommandTimeout = 0;
                    BL_Quary1.ExecuteNonQuery();
                    BL_Table = new DataTable();
                    SqlDataAdapter BL_Table1 = new SqlDataAdapter(BL_Quary1);
                    BL_Table1.Fill(BL_Table);
                }
                if (Technology == "3G_CS")
                {
                    string BL_Quary = "select * from BL_3G_CS where Contractor='" + Region + "'";
                    SqlCommand BL_Quary1 = new SqlCommand(BL_Quary, connection);
                    BL_Quary1.CommandTimeout = 0;
                    BL_Quary1.ExecuteNonQuery();
                    BL_Table = new DataTable();
                    SqlDataAdapter BL_Table1 = new SqlDataAdapter(BL_Quary1);
                    BL_Table1.Fill(BL_Table);
                }
                if (Technology == "3G_PS")
                {
                    string BL_Quary = "select * from BL_3G_PS where Contractor='" + Region + "'";
                    SqlCommand BL_Quary1 = new SqlCommand(BL_Quary, connection);
                    BL_Quary1.CommandTimeout = 0;
                    BL_Quary1.ExecuteNonQuery();
                    BL_Table = new DataTable();
                    SqlDataAdapter BL_Table1 = new SqlDataAdapter(BL_Quary1);
                    BL_Table1.Fill(BL_Table);
                }
                if (Technology == "4G")
                {
                    string BL_Quary = "select * from BL_4G_PS where Contractor='" + Region + "'";
                    SqlCommand BL_Quary1 = new SqlCommand(BL_Quary, connection);
                    BL_Quary1.CommandTimeout = 0;
                    BL_Quary1.ExecuteNonQuery();
                    BL_Table = new DataTable();
                    SqlDataAdapter BL_Table1 = new SqlDataAdapter(BL_Quary1);
                    BL_Table1.Fill(BL_Table);
                }





                string Single_Region_List_Quary_Contractual = "";
                string Single_Region_List_Quary_NearContractual = "";
                string Single_Region_Misssing_Quary_Contractual = "";
                string Single_Region_Missing_Quary_NearContractual = "";
                if (Technology == "2G_CS")
                {
                    Single_Region_List_Quary_Contractual = "select [Date], [Contractor], [Level], [Status], count(Level) as 'Worst_Count' from [Contractual_WPC_2G_CS] where [Contractor] = '" + Region + "' and Status = 'Contractual WPC' group by[Date], [Contractor], [Level], [Status]  order by Date";
                    Single_Region_List_Quary_NearContractual = "select [Date], [Contractor], [Level], [Status], count(Level) as 'Worst_Count' from [Contractual_WPC_2G_CS] where [Contractor] = '" + Region + "' and Status = 'Near Contractual WPC' group by[Date], [Contractor], [Level], [Status]  order by Date";
                    Single_Region_Misssing_Quary_Contractual = "select [Date], [Contractor], [Status], [Level], sum([Effeciency Index])*1.1 as 'Missing_Score' from [Contractual_WPC_2G_CS] where [Contractor] = '" + Region + "' and Status = 'Contractual WPC' group by [Date], [Contractor], [Status], [Level]  order by Date";
                    Single_Region_Missing_Quary_NearContractual = "select [Date], [Contractor], [Status], [Level], sum([Effeciency Index]) as 'Missing_Score' from [Contractual_WPC_2G_CS] where [Contractor] = '" + Region + "' and Status = 'Near Contractual WPC' and QIxP<100 group by [Date], [Contractor], [Status], [Level]  order by Date";
                }
                if (Technology == "2G_PS")
                {
                    Single_Region_List_Quary_Contractual = "select [Date], [Contractor], [Level], [Status], count(Level) as 'Worst_Count' from [Contractual_WPC_2G_PS] where [Contractor] = '" + Region + "' and Status = 'Contractual WPC' group by[Date], [Contractor], [Level], [Status]  order by Date";
                    Single_Region_List_Quary_NearContractual = "select [Date], [Contractor], [Level], [Status], count(Level) as 'Worst_Count' from [Contractual_WPC_2G_PS] where [Contractor] = '" + Region + "' and Status = 'Near WPC' group by[Date], [Contractor], [Level], [Status]  order by Date";
                    Single_Region_Misssing_Quary_Contractual = "select [Date], [Contractor], [Status], [Level], sum([Effeciency Index])*1.1 as 'Missing_Score' from [Contractual_WPC_2G_PS] where [Contractor] = '" + Region + "' and Status = 'Contractual WPC' group by [Date], [Contractor], [Status], [Level]  order by Date";
                    Single_Region_Missing_Quary_NearContractual = "select [Date], [Contractor], [Status], [Level], sum([Effeciency Index]) as 'Missing_Score' from [Contractual_WPC_2G_PS] where [Contractor] = '" + Region + "' and Status = 'Near WPC' and QIxP<100 group by [Date], [Contractor], [Status], [Level]  order by Date";
                }
                if (Technology == "3G_CS")
                {
                    Single_Region_List_Quary_Contractual = "select [Date], [Contractor], [Level], [Status], count(Level) as 'Worst_Count' from [Contractual_WPC_3G_CS] where [Contractor] = '" + Region + "' and Status = 'Contractual WPC' group by[Date], [Contractor], [Level], [Status]  order by Date";
                    Single_Region_List_Quary_NearContractual = "select [Date], [Contractor], [Level], [Status], count(Level) as 'Worst_Count' from [Contractual_WPC_3G_CS] where [Contractor] = '" + Region + "' and Status = 'Near Contractual WPC' group by[Date], [Contractor], [Level], [Status]  order by Date";
                    Single_Region_Misssing_Quary_Contractual = "select [Date], [Contractor], [Status], [Level], sum([Effeciency Index])*1.1 as 'Missing_Score' from [Contractual_WPC_3G_CS] where [Contractor] = '" + Region + "' and Status = 'Contractual WPC' group by [Date], [Contractor], [Status], [Level]  order by Date";
                    Single_Region_Missing_Quary_NearContractual = "select [Date], [Contractor], [Status], [Level], sum([Effeciency Index]) as 'Missing_Score' from [Contractual_WPC_3G_CS] where [Contractor] = '" + Region + "' and Status = 'Near Contractual WPC' and QIxP<100 group by [Date], [Contractor], [Status], [Level]  order by Date";
                }
                if (Technology == "3G_PS")
                {
                    Single_Region_List_Quary_Contractual = "select [Date], [Contractor], [Level], [Status], count(Level) as 'Worst_Count' from [Contractual_WPC_3G_PS] where [Contractor] = '" + Region + "' and Status = 'Contractual WPC' group by[Date], [Contractor], [Level], [Status]  order by Date, Status, Level";
                    Single_Region_List_Quary_NearContractual = "select [Date], [Contractor], [Level], [Status], count(Level) as 'Worst_Count' from [Contractual_WPC_3G_PS] where [Contractor] = '" + Region + "' and Status = 'Near Contractual WPC' group by[Date], [Contractor], [Level], [Status]  order by Date, Status, Level";
                    Single_Region_Misssing_Quary_Contractual = "select [Date], [Contractor], [Status], [Level], sum([Effeciency Index])*1.1 as 'Missing_Score' from [Contractual_WPC_3G_PS] where [Contractor] = '" + Region + "' and Status = 'Contractual WPC' group by [Date], [Contractor], [Status], [Level]  order by Date, Status, Level";
                    Single_Region_Missing_Quary_NearContractual = "select [Date], [Contractor], [Status], [Level], sum([Effeciency Index]) as 'Missing_Score' from [Contractual_WPC_3G_PS] where [Contractor] = '" + Region + "' and Status = 'Near Contractual WPC' and QIxP<100 group by [Date], [Contractor], [Status], [Level]  order by Date, Status, Level";
                }
                if (Technology == "4G")
                {
                    Single_Region_List_Quary_Contractual = "select [Date], [Contractor], [Level], [Status], count(Level) as 'Worst_Count' from [Contractual_WPC_4G] where [Contractor] = '" + Region + "' and Status = 'Contractual WPC' group by [Date], [Contractor], [Level], [Status]  order by Date";
                    Single_Region_List_Quary_NearContractual = "select [Date], [Contractor], [Level], [Status], count(Level) as 'Worst_Count' from [Contractual_WPC_4G] where [Contractor] = '" + Region + "' and Status = 'Near Contractual WPC' group by[Date], [Contractor], [Level], [Status]  order by Date";
                    Single_Region_Misssing_Quary_Contractual = "select [Date], [Contractor], [Status], [Level], sum([Effeciency Index])*1.1 as 'Missing_Score' from [Contractual_WPC_4G] where [Contractor] = '" + Region + "' and Status = 'Contractual WPC' group by [Date], [Contractor], [Status] , [Level] order by Date";
                    Single_Region_Missing_Quary_NearContractual = "select [Date], [Contractor], [Status], [Level], sum([Effeciency Index]) as 'Missing_Score' from [Contractual_WPC_4G] where [Contractor] = '" + Region + "' and Status = 'Near Contractual WPC' and QIxP<100 group by [Date], [Contractor], [Status], [Level]  order by Date";
                }

                // Worst Cells Count in Contractual WPC
                SqlCommand Single_Region_List_Quary_Contractual1 = new SqlCommand(Single_Region_List_Quary_Contractual, connection);
                Single_Region_List_Quary_Contractual1.ExecuteNonQuery();
                Single_Region_Table_Contractual = new DataTable();
                SqlDataAdapter dataAdapter_Contractual = new SqlDataAdapter(Single_Region_List_Quary_Contractual1);
                dataAdapter_Contractual.Fill(Single_Region_Table_Contractual);

                // Worst Cells Count in Near Contractual WPC
                SqlCommand Single_Region_List_Quary_NearContractual1 = new SqlCommand(Single_Region_List_Quary_NearContractual, connection);
                Single_Region_List_Quary_NearContractual1.ExecuteNonQuery();
                DataTable Single_Region_Table_NearContractual = new DataTable();
                SqlDataAdapter dataAdapter_NearContractual = new SqlDataAdapter(Single_Region_List_Quary_NearContractual1);
                dataAdapter_NearContractual.Fill(Single_Region_Table_NearContractual);

                // Worst Cells Missing Score in Contractual WPC
                SqlCommand Single_Region_Missing_Quary_Contractual1 = new SqlCommand(Single_Region_Misssing_Quary_Contractual, connection);
                Single_Region_Missing_Quary_Contractual1.ExecuteNonQuery();
                DataTable Single_Region_Table_Contractual1 = new DataTable();
                SqlDataAdapter dataAdapter_Contractual1 = new SqlDataAdapter(Single_Region_Missing_Quary_Contractual1);
                dataAdapter_Contractual1.Fill(Single_Region_Table_Contractual1);

                // Worst Cells Missing Score in Near Contractual WPC
                SqlCommand Single_Region_Missing_Quary_NearContractual1 = new SqlCommand(Single_Region_Missing_Quary_NearContractual, connection);
                Single_Region_Missing_Quary_NearContractual1.ExecuteNonQuery();
                DataTable Single_Region_Table_NearContractual1 = new DataTable();
                SqlDataAdapter dataAdapter_NearContractual1 = new SqlDataAdapter(Single_Region_Missing_Quary_NearContractual1);
                dataAdapter_NearContractual1.Fill(Single_Region_Table_NearContractual1);


                progressBar1.Value = 4;

                chart1.Series.Clear();
                chart1.Titles.Clear();

                Series newSeries1 = new Series();
                chart1.Series.Add(newSeries1);
                newSeries1.IsXValueIndexed = false;
                chart1.Series[0].ChartType = SeriesChartType.StackedColumn;
                if (chart_type == "SeriesChartType.Line")
                {
                    chart1.Series[0].ChartType = SeriesChartType.Line;
                }
                if (chart_type == "SeriesChartType.StackedColumn")
                {
                    chart1.Series[0].ChartType = SeriesChartType.StackedColumn;
                }
                chart1.Series[0].Color = Color.Red;
                chart1.Series[0].BorderWidth = 3;
                chart1.ChartAreas[0].AxisX.Interval = 5;
                chart1.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
                chart1.Series[0].EmptyPointStyle.Color = Color.Transparent;
                chart1.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";
                chart1.Series[0].ToolTip = "#VALX [#VALY]";
                chart1.Series[0].IsValueShownAsLabel = false;
                chart1.Series[0].LegendText = "Level 1";
                chart1.Legends["Legend1"].Docking = Docking.Bottom;




                Series newSeries2 = new Series();
                chart1.Series.Add(newSeries2);
                newSeries2.IsXValueIndexed = false;
                chart1.Series[1].ChartType = SeriesChartType.StackedColumn;
                if (chart_type == "SeriesChartType.Line")
                {
                    chart1.Series[1].ChartType = SeriesChartType.Line;
                }
                if (chart_type == "SeriesChartType.StackedColumn")
                {
                    chart1.Series[1].ChartType = SeriesChartType.StackedColumn;
                }
                chart1.Series[1].Color = Color.Orange;
                chart1.Series[1].BorderWidth = 3;
                chart1.Series[1].EmptyPointStyle.Color = Color.Transparent;
                chart1.Series[1].ToolTip = "#VALX [#VALY]";
                chart1.Series[1].IsValueShownAsLabel = false;
                chart1.Series[1].LegendText = "Level 2";

                Series newSeries3 = new Series();
                chart1.Series.Add(newSeries3);
                newSeries3.IsXValueIndexed = false;
                chart1.Series[2].ChartType = SeriesChartType.StackedColumn;
                if (chart_type == "SeriesChartType.Line")
                {
                    chart1.Series[2].ChartType = SeriesChartType.Line;
                }
                if (chart_type == "SeriesChartType.StackedColumn")
                {
                    chart1.Series[2].ChartType = SeriesChartType.StackedColumn;
                }
                chart1.Series[2].Color = Color.Yellow;
                chart1.Series[2].BorderWidth = 3;
                chart1.Series[2].EmptyPointStyle.Color = Color.Transparent;
                chart1.Series[2].ToolTip = "#VALX [#VALY]";
                chart1.Series[2].IsValueShownAsLabel = false;
                chart1.Series[2].LegendText = "Level 3";

                Series newSeries4 = new Series();
                chart1.Series.Add(newSeries4);
                newSeries4.IsXValueIndexed = false;
                chart1.Series[3].ChartType = SeriesChartType.StackedColumn;
                if (chart_type == "SeriesChartType.Line")
                {
                    chart1.Series[3].ChartType = SeriesChartType.Line;
                }
                if (chart_type == "SeriesChartType.StackedColumn")
                {
                    chart1.Series[3].ChartType = SeriesChartType.StackedColumn;
                }
                chart1.Series[3].Color = Color.Blue;
                chart1.Series[3].BorderWidth = 3;
                chart1.Series[3].EmptyPointStyle.Color = Color.Transparent;
                chart1.Series[3].ToolTip = "#VALX [#VALY]";
                chart1.Series[3].IsValueShownAsLabel = false;
                chart1.Series[3].LegendText = "Level 4";

                Series newSeries5 = new Series();
                chart1.Series.Add(newSeries5);
                newSeries5.IsXValueIndexed = false;
                chart1.Series[4].ChartType = SeriesChartType.StackedColumn;
                if (chart_type == "SeriesChartType.Line")
                {
                    chart1.Series[4].ChartType = SeriesChartType.Line;
                }
                if (chart_type == "SeriesChartType.StackedColumn")
                {
                    chart1.Series[4].ChartType = SeriesChartType.StackedColumn;
                }
                chart1.Series[4].Color = Color.Green;
                chart1.Series[4].BorderWidth = 3;
                chart1.Series[4].EmptyPointStyle.Color = Color.Transparent;
                chart1.Series[4].ToolTip = "#VALX [#VALY]";
                chart1.Series[4].IsValueShownAsLabel = false;
                chart1.Series[4].LegendText = "Level 5";



                chart2.Series.Clear();
                chart2.Titles.Clear();

                Series newSeries6 = new Series();
                chart2.Series.Add(newSeries6);
                newSeries6.IsXValueIndexed = false;
                chart2.Series[0].ChartType = SeriesChartType.StackedColumn;
                if (chart_type == "SeriesChartType.Line")
                {
                    chart2.Series[0].ChartType = SeriesChartType.Line;
                }
                if (chart_type == "SeriesChartType.StackedColumn")
                {
                    chart2.Series[0].ChartType = SeriesChartType.StackedColumn;
                }
                chart2.Series[0].Color = Color.Red;
                chart2.Series[0].BorderWidth = 3;
                chart2.ChartAreas[0].AxisX.Interval = 5;
                chart2.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
                chart2.Series[0].EmptyPointStyle.Color = Color.Transparent;
                chart2.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";
                chart2.Series[0].ToolTip = "#VALX [#VALY]";
                chart2.Series[0].IsValueShownAsLabel = false;
                chart2.Series[0].LegendText = "Level 1";
                chart2.Legends["Legend1"].Docking = Docking.Bottom;

                Series newSeries7 = new Series();
                chart2.Series.Add(newSeries7);
                newSeries7.IsXValueIndexed = false;
                chart2.Series[1].ChartType = SeriesChartType.StackedColumn;
                if (chart_type == "SeriesChartType.Line")
                {
                    chart2.Series[1].ChartType = SeriesChartType.Line;
                }
                if (chart_type == "SeriesChartType.StackedColumn")
                {
                    chart2.Series[1].ChartType = SeriesChartType.StackedColumn;
                }
                chart2.Series[1].Color = Color.Orange;
                chart2.Series[1].BorderWidth = 3;
                chart2.Series[1].EmptyPointStyle.Color = Color.Transparent;
                chart2.Series[1].ToolTip = "#VALX [#VALY]";
                chart2.Series[1].IsValueShownAsLabel = false;
                chart2.Series[1].LegendText = "Level 2";

                Series newSeries8 = new Series();
                chart2.Series.Add(newSeries8);
                newSeries8.IsXValueIndexed = false;
                chart2.Series[2].ChartType = SeriesChartType.StackedColumn;
                if (chart_type == "SeriesChartType.Line")
                {
                    chart2.Series[2].ChartType = SeriesChartType.Line;
                }
                if (chart_type == "SeriesChartType.StackedColumn")
                {
                    chart2.Series[2].ChartType = SeriesChartType.StackedColumn;
                }
                chart2.Series[2].Color = Color.Yellow;
                chart2.Series[2].BorderWidth = 3;
                chart2.Series[2].EmptyPointStyle.Color = Color.Transparent;
                chart2.Series[2].ToolTip = "#VALX [#VALY]";
                chart2.Series[2].IsValueShownAsLabel = false;
                chart2.Series[2].LegendText = "Level 3";

                Series newSeries9 = new Series();
                chart2.Series.Add(newSeries9);
                newSeries9.IsXValueIndexed = false;
                chart2.Series[3].ChartType = SeriesChartType.StackedColumn;
                if (chart_type == "SeriesChartType.Line")
                {
                    chart2.Series[3].ChartType = SeriesChartType.Line;
                }
                if (chart_type == "SeriesChartType.StackedColumn")
                {
                    chart2.Series[3].ChartType = SeriesChartType.StackedColumn;
                }
                chart2.Series[3].Color = Color.Blue;
                chart2.Series[3].BorderWidth = 3;
                chart2.Series[3].EmptyPointStyle.Color = Color.Transparent;
                chart2.Series[3].ToolTip = "#VALX [#VALY]";
                chart2.Series[3].IsValueShownAsLabel = false;
                chart2.Series[3].LegendText = "Level 4";

                Series newSeries10 = new Series();
                chart2.Series.Add(newSeries10);
                newSeries10.IsXValueIndexed = false;
                chart2.Series[4].ChartType = SeriesChartType.StackedColumn;
                if (chart_type == "SeriesChartType.Line")
                {
                    chart2.Series[4].ChartType = SeriesChartType.Line;
                }
                if (chart_type == "SeriesChartType.StackedColumn")
                {
                    chart2.Series[4].ChartType = SeriesChartType.StackedColumn;
                }
                chart2.Series[4].Color = Color.Green;
                chart2.Series[4].BorderWidth = 3;
                chart2.Series[4].EmptyPointStyle.Color = Color.Transparent;
                chart2.Series[4].ToolTip = "#VALX [#VALY]";
                chart2.Series[4].IsValueShownAsLabel = false;
                chart2.Series[4].LegendText = "Level 5";





                Title title1 = chart1.Titles.Add("Number of '" + Region + "-" + Technology + "' Worst Cells (Contractual)");
                title1.Font = new System.Drawing.Font("Arial", 12, FontStyle.Regular);
                Title title2 = chart2.Titles.Add("Number of '" + Region + "-" + Technology + "' Worst Cells (Near Contractual)");
                title2.Font = new System.Drawing.Font("Arial", 12, FontStyle.Regular);




                chart3.Series.Clear();
                chart3.Titles.Clear();

                Series newSeries11 = new Series();
                chart3.Series.Add(newSeries11);
                newSeries11.IsXValueIndexed = false;
                chart3.Series[0].ChartType = SeriesChartType.StackedColumn;
                chart3.Series[0].Color = Color.Brown;
                chart3.Series[0].BorderWidth = 3;
                chart3.ChartAreas[0].AxisX.Interval = 5;
                chart3.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
                chart3.Series[0].EmptyPointStyle.Color = Color.Transparent;
                chart3.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";
                chart3.Series[0].ToolTip = "#VALX [#VALY]";
                chart3.Series[0].IsValueShownAsLabel = false;
                chart3.Series[0].IsVisibleInLegend = false;

                Series newSeries12 = new Series();
                chart3.Series.Add(newSeries12);
                newSeries12.IsXValueIndexed = false;
                chart3.Series[1].ChartType = SeriesChartType.StackedColumn;
                chart3.Series[1].Color = Color.Brown;
                chart3.Series[1].BorderWidth = 3;
                chart3.ChartAreas[0].AxisX.Interval = 5;
                chart3.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
                chart3.Series[1].EmptyPointStyle.Color = Color.Transparent;
                chart3.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";
                chart3.Series[1].ToolTip = "#VALX [#VALY]";
                chart3.Series[1].IsValueShownAsLabel = false;
                chart3.Series[1].IsVisibleInLegend = false;


                Series newSeries13 = new Series();
                chart3.Series.Add(newSeries13);
                newSeries13.IsXValueIndexed = false;
                chart3.Series[2].ChartType = SeriesChartType.StackedColumn;
                chart3.Series[2].Color = Color.Brown;
                chart3.Series[2].BorderWidth = 3;
                chart3.ChartAreas[0].AxisX.Interval = 5;
                chart3.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
                chart3.Series[2].EmptyPointStyle.Color = Color.Transparent;
                chart3.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";
                chart3.Series[2].ToolTip = "#VALX [#VALY]";
                chart3.Series[2].IsValueShownAsLabel = false;
                chart3.Series[2].IsVisibleInLegend = false;

                Series newSeries14 = new Series();
                chart3.Series.Add(newSeries14);
                newSeries14.IsXValueIndexed = false;
                chart3.Series[3].ChartType = SeriesChartType.StackedColumn;
                chart3.Series[3].Color = Color.Brown;
                chart3.Series[3].BorderWidth = 3;
                chart3.ChartAreas[0].AxisX.Interval = 5;
                chart3.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
                chart3.Series[3].EmptyPointStyle.Color = Color.Transparent;
                chart3.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";
                chart3.Series[3].ToolTip = "#VALX [#VALY]";
                chart3.Series[3].IsValueShownAsLabel = false;
                chart3.Series[3].IsVisibleInLegend = false;


                Series newSeries15 = new Series();
                chart3.Series.Add(newSeries15);
                newSeries15.IsXValueIndexed = false;
                chart3.Series[4].ChartType = SeriesChartType.StackedColumn;
                chart3.Series[4].Color = Color.Brown;
                chart3.Series[4].BorderWidth = 3;
                chart3.ChartAreas[0].AxisX.Interval = 5;
                chart3.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
                chart3.Series[4].EmptyPointStyle.Color = Color.Transparent;
                chart3.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";
                chart3.Series[4].ToolTip = "#VALX [#VALY]";
                chart3.Series[4].IsValueShownAsLabel = false;
                chart3.Series[4].IsVisibleInLegend = false;


                chart4.Series.Clear();
                chart4.Titles.Clear();


                Series newSeries16 = new Series();
                chart4.Series.Add(newSeries16);
                newSeries16.IsXValueIndexed = false;
                chart4.Series[0].ChartType = SeriesChartType.StackedColumn;
                chart4.Series[0].Color = Color.Brown;
                chart4.Series[0].BorderWidth = 3;
                chart4.ChartAreas[0].AxisX.Interval = 5;
                chart4.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
                chart4.Series[0].EmptyPointStyle.Color = Color.Transparent;
                chart4.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";
                chart4.Series[0].ToolTip = "#VALX [#VALY]";
                chart4.Series[0].IsValueShownAsLabel = false;
                chart4.Series[0].IsVisibleInLegend = false;

                Series newSeries17 = new Series();
                chart4.Series.Add(newSeries17);
                newSeries17.IsXValueIndexed = false;
                chart4.Series[1].ChartType = SeriesChartType.StackedColumn;
                chart4.Series[1].Color = Color.Brown;
                chart4.Series[1].BorderWidth = 3;
                chart4.ChartAreas[0].AxisX.Interval = 5;
                chart4.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
                chart4.Series[1].EmptyPointStyle.Color = Color.Transparent;
                chart4.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";
                chart4.Series[1].ToolTip = "#VALX [#VALY]";
                chart4.Series[1].IsValueShownAsLabel = false;
                chart4.Series[1].IsVisibleInLegend = false;


                Series newSeries18 = new Series();
                chart4.Series.Add(newSeries18);
                newSeries18.IsXValueIndexed = false;
                chart4.Series[2].ChartType = SeriesChartType.StackedColumn;
                chart4.Series[2].Color = Color.Brown;
                chart4.Series[2].BorderWidth = 3;
                chart4.ChartAreas[0].AxisX.Interval = 5;
                chart4.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
                chart4.Series[2].EmptyPointStyle.Color = Color.Transparent;
                chart4.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";
                chart4.Series[2].ToolTip = "#VALX [#VALY]";
                chart4.Series[2].IsValueShownAsLabel = false;
                chart4.Series[2].IsVisibleInLegend = false;

                Series newSeries19 = new Series();
                chart4.Series.Add(newSeries19);
                newSeries19.IsXValueIndexed = false;
                chart4.Series[3].ChartType = SeriesChartType.StackedColumn;
                chart4.Series[3].Color = Color.Brown;
                chart4.Series[3].BorderWidth = 3;
                chart4.ChartAreas[0].AxisX.Interval = 5;
                chart4.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
                chart4.Series[3].EmptyPointStyle.Color = Color.Transparent;
                chart4.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";
                chart4.Series[3].ToolTip = "#VALX [#VALY]";
                chart4.Series[3].IsValueShownAsLabel = false;
                chart4.Series[3].IsVisibleInLegend = false;


                Series newSeries20 = new Series();
                chart4.Series.Add(newSeries20);
                newSeries20.IsXValueIndexed = false;
                chart4.Series[4].ChartType = SeriesChartType.StackedColumn;
                chart4.Series[4].Color = Color.Brown;
                chart4.Series[4].BorderWidth = 3;
                chart4.ChartAreas[0].AxisX.Interval = 5;
                chart4.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
                chart4.Series[4].EmptyPointStyle.Color = Color.Transparent;
                chart4.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";
                chart4.Series[4].ToolTip = "#VALX [#VALY]";
                chart4.Series[4].IsValueShownAsLabel = false;
                chart4.Series[4].IsVisibleInLegend = false;




                // chart1.DataManipulator.InsertEmptyPoints(0, IntervalType.Number, "Series1,Series2");
                //chart1.DataManipulator.InsertEmptyPoints(1, IntervalType.Days, newSeries2);
                //chart1.DataManipulator.InsertEmptyPoints(1, IntervalType.Days, newSeries3);
                //chart1.DataManipulator.InsertEmptyPoints(1, IntervalType.Days, newSeries4);
                //chart1.DataManipulator.InsertEmptyPoints(1, IntervalType.Days, newSeries5);



                Title title3 = chart3.Titles.Add("Sum of '" + Region + "-" + Technology + "' Missing Score (Contractual)");
                title3.Font = new System.Drawing.Font("Arial", 12, FontStyle.Regular);
                Title title4 = chart4.Titles.Add("Sum of '" + Region + "-" + Technology + "' Missing Score (Near Contractual)");
                title4.Font = new System.Drawing.Font("Arial", 12, FontStyle.Regular);


                int y1 = 0;

                DateTime Old_Date1= Convert.ToDateTime((Single_Region_Table_Contractual.Rows[0]).ItemArray[0]);
                Old_Date1 = Old_Date1.AddHours(23);
                int date_counter1 = 0;

                DateTime Old_Date2 = Convert.ToDateTime((Single_Region_Table_NearContractual.Rows[0]).ItemArray[0]);
                Old_Date2 = Old_Date2.AddHours(23);
                int date_counter2= 0;

                DateTime Old_Date3 = Convert.ToDateTime((Single_Region_Table_Contractual1.Rows[0]).ItemArray[0]);
                Old_Date3 = Old_Date3.AddHours(23);
                int date_counter3= 0;

                DateTime Old_Date4 = Convert.ToDateTime((Single_Region_Table_NearContractual1.Rows[0]).ItemArray[0]);
                Old_Date4 = Old_Date4.AddHours(23);
                int date_counter4 = 0;


                for (int k = 1; k <= Single_Region_Table_Contractual.Rows.Count; k++)
                {

                    DateTime dt1 = Convert.ToDateTime((Single_Region_Table_Contractual.Rows[k - 1]).ItemArray[0]);
                    dt1 = dt1.AddHours(23);

                    if (dt1 == Old_Date1)
                    {
                        date_counter1++;
                    }
                    else
                    {
                        if (date_counter1!=5)
                        {
                            chart1.Series[0].Points.AddXY(Old_Date1, 0);
                        }
                        date_counter1 = 1;
                        Old_Date1 = dt1;
                    }


                    // Setting of Intervals
                    // **************************************************************
                    double dt1_double = dt1.Year * 10000 + dt1.Month * 100 + dt1.Day;

                    if (dt1_double > Max_X2)
                    {
                        Max_X2 = dt1_double;
                        Max_X_Date = dt1;
                    }
                    if (dt1_double < Min_X2)
                    {
                        Min_X2 = dt1_double;
                        Min_X_Date = dt1;
                    }
                    // **************************************************************




                    double Worst_Cell_Count1 = Convert.ToDouble((Single_Region_Table_Contractual.Rows[k - 1]).ItemArray[4]);
                    double Level1 = Convert.ToDouble((Single_Region_Table_Contractual.Rows[k - 1]).ItemArray[2]);


                    if (Level1 == 1 && checkBox6.Checked == true)
                    {
                        chart1.Series[0].Points.AddXY(dt1, Worst_Cell_Count1);
                    }
                    if (Level1 == 2 && checkBox7.Checked == true)
                    {
                        chart1.Series[1].Points.AddXY(dt1, Worst_Cell_Count1);
                    }
                    if (Level1 == 3 && checkBox8.Checked == true)
                    {
                        chart1.Series[2].Points.AddXY(dt1, Worst_Cell_Count1);
                    }
                    if (Level1 == 4 && checkBox9.Checked == true)
                    {
                        chart1.Series[3].Points.AddXY(dt1, Worst_Cell_Count1);
                    }
                    if (Level1 == 5 && checkBox10.Checked == true)
                    {
                        chart1.Series[4].Points.AddXY(dt1, Worst_Cell_Count1);
                    }
                    Max_Y1 = chart1.ChartAreas[0].AxisY.Maximum;
                    if (k == 1)
                    {
                        Min_X = dt1;
                    }

                    y1 = k;
                }
                Max_X = Convert.ToDateTime((Single_Region_Table_Contractual.Rows[y1 - 1]).ItemArray[0]);



                // Setting of Intervals
                // **************************************************************
                double difference_day = (Max_X_Date - Min_X_Date).TotalDays;
                double day_interval = Math.Round(difference_day / 20);
                if (day_interval == 0)
                {
                    day_interval = 1;
                }
                chart1.ChartAreas[0].AxisX.Interval = day_interval;
                chart2.ChartAreas[0].AxisX.Interval = day_interval;
                chart3.ChartAreas[0].AxisX.Interval = day_interval;
                chart4.ChartAreas[0].AxisX.Interval = day_interval;
                // **************************************************************


                int y2 = 0;
                for (int k = 1; k <= Single_Region_Table_NearContractual.Rows.Count; k++)
                {

                    DateTime dt2 = Convert.ToDateTime((Single_Region_Table_NearContractual.Rows[k - 1]).ItemArray[0]);
                    dt2 = dt2.AddHours(23);


                    if (dt2 == Old_Date2)
                    {
                        date_counter2++;
                    }
                    else
                    {
                        if (date_counter2 != 5)
                        {
                            chart2.Series[0].Points.AddXY(Old_Date2, 0);
                        }
                        date_counter2 = 1;
                        Old_Date2 = dt2;
                    }


                    double Worst_Cell_Count2 = Convert.ToDouble((Single_Region_Table_NearContractual.Rows[k - 1]).ItemArray[4]);
                    double Level2 = Convert.ToDouble((Single_Region_Table_NearContractual.Rows[k - 1]).ItemArray[2]);

                    if (Level2 == 1 && checkBox6.Checked == true)
                    {
                        chart2.Series[0].Points.AddXY(dt2, Worst_Cell_Count2);
                    }
                    if (Level2 == 2 && checkBox7.Checked == true)
                    {
                        chart2.Series[1].Points.AddXY(dt2, Worst_Cell_Count2);
                    }
                    if (Level2 == 3 && checkBox8.Checked == true)
                    {
                        chart2.Series[2].Points.AddXY(dt2, Worst_Cell_Count2);
                    }
                    if (Level2 == 4 && checkBox9.Checked == true)
                    {
                        chart2.Series[3].Points.AddXY(dt2, Worst_Cell_Count2);
                    }
                    if (Level2 == 5 && checkBox10.Checked == true)
                    {
                        chart2.Series[4].Points.AddXY(dt2, Worst_Cell_Count2);
                    }
                    Max_Y2 = chart2.ChartAreas[0].AxisY.Maximum;
                    if (k == 1)
                    {
                        Min_X = dt2;
                    }

                    y2 = k;
                }
                Max_X = Convert.ToDateTime((Single_Region_Table_NearContractual.Rows[y2 - 1]).ItemArray[0]);



                for (int k = 1; k <= Single_Region_Table_Contractual1.Rows.Count; k++)
                {
                    DateTime dt1 = Convert.ToDateTime((Single_Region_Table_Contractual1.Rows[k - 1]).ItemArray[0]);
                    dt1 = dt1.AddHours(23);

                    if (dt1 == Old_Date3)
                    {
                        date_counter3++;
                    }
                    else
                    {
                        if (date_counter3 != 5)
                        {
                            chart3.Series[0].Points.AddXY(Old_Date3, 0);
                        }
                        date_counter3 = 1;
                        Old_Date3 = dt1;
                    }


                    double Score_Cell_Sum = Convert.ToDouble((Single_Region_Table_Contractual1.Rows[k - 1]).ItemArray[4]);
                    double Level = Convert.ToDouble((Single_Region_Table_Contractual1.Rows[k - 1]).ItemArray[3]);
                    if (Level == 1 && checkBox6.Checked == true)
                    {
                        chart3.Series[0].Points.AddXY(dt1, Score_Cell_Sum);
                    }
                    if (Level == 2 && checkBox7.Checked == true)
                    {
                        chart3.Series[1].Points.AddXY(dt1, Score_Cell_Sum);
                    }
                    if (Level == 3 && checkBox8.Checked == true)
                    {
                        chart3.Series[2].Points.AddXY(dt1, Score_Cell_Sum);
                    }
                    if (Level == 4 && checkBox9.Checked == true)
                    {
                        chart3.Series[3].Points.AddXY(dt1, Score_Cell_Sum);
                    }
                    if (Level == 5 && checkBox10.Checked == true)
                    {
                        chart3.Series[4].Points.AddXY(dt1, Score_Cell_Sum);
                    }
                    //chart3.Series[0].Points.AddXY(dt1, Score_Cell_Sum);
                    Max_Y3 = chart3.ChartAreas[0].AxisY.Maximum;
                }

                for (int k = 1; k <= Single_Region_Table_NearContractual1.Rows.Count; k++)
                {
                    DateTime dt1 = Convert.ToDateTime((Single_Region_Table_NearContractual1.Rows[k - 1]).ItemArray[0]);
                    dt1 = dt1.AddHours(23);

                    if (dt1 == Old_Date4)
                    {
                        date_counter4++;
                    }
                    else
                    {
                        if (date_counter4 != 5)
                        {
                            chart4.Series[0].Points.AddXY(Old_Date4, 0);
                        }
                        date_counter4 = 1;
                        Old_Date4 = dt1;
                    }


                    double Score_Cell_Sum = Convert.ToDouble((Single_Region_Table_NearContractual1.Rows[k - 1]).ItemArray[4]);
                    //chart4.Series[0].Points.AddXY(dt1, Score_Cell_Sum);
                    double Level = Convert.ToDouble((Single_Region_Table_NearContractual1.Rows[k - 1]).ItemArray[3]);
                    if (Level == 1 && checkBox6.Checked == true)
                    {
                        chart4.Series[0].Points.AddXY(dt1, Score_Cell_Sum);
                    }
                    if (Level == 2 && checkBox7.Checked == true)
                    {
                        chart4.Series[1].Points.AddXY(dt1, Score_Cell_Sum);
                    }
                    if (Level == 3 && checkBox8.Checked == true)
                    {
                        chart4.Series[2].Points.AddXY(dt1, Score_Cell_Sum);
                    }
                    if (Level == 4 && checkBox9.Checked == true)
                    {
                        chart4.Series[3].Points.AddXY(dt1, Score_Cell_Sum);
                    }
                    if (Level == 5 && checkBox10.Checked == true)
                    {
                        chart4.Series[4].Points.AddXY(dt1, Score_Cell_Sum);
                    }
                    Max_Y4 = chart4.ChartAreas[0].AxisY.Maximum;

                }



                progressBar1.Value = 5;

            }
        }


        // Province Selection
        void comboBox2_MouseWheel(object sender, MouseEventArgs e)
        {
            ((HandledMouseEventArgs)e).Handled = true;
        }
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox2.MouseWheel += new MouseEventHandler(comboBox2_MouseWheel);
            //comboBox4.Items.Clear();
            Region_Show = 0;
            Province_Show = 1;
            Node_Show = 0;
            label1.BackColor = Color.PaleGoldenrod;
            label2.BackColor = Color.Yellow;
            label4.BackColor = Color.PaleGoldenrod;
            label5.BackColor = Color.PaleGoldenrod;
            Province = comboBox2.SelectedItem.ToString();


            // Node Update BASED ON PROVINCE"
            //string Update_Quary = "";
            //string Update_Quary_CS = "";
            // string Update_Quary_PS = "";
            string Node_List_Quary = "";
            string Date_Quary = "";
            string Cell_Quary = "";
            if (Technology == "2G_CS")
            {
                //Update_Quary = "update [Contractual_WPC_2G_CS] set[Contractual_WPC_2G_CS].[Status] = 'Near Contractual WPC' where [Contractual_WPC_2G_CS].[Status] is null";
                Node_List_Quary = "select distinct([BSC]) from [Contractual_WPC_2G_CS] where Contractor='" + Region + "' and Province='" + Province + "' and date=(select max(distinct([Date])) from [Contractual_WPC_2G_CS]  where Contractor='" + Region + "') order by BSC";
                //   SqlCommand Update_Quary1 = new SqlCommand(Update_Quary, connection);
                //    Update_Quary1.ExecuteNonQuery();

                //Date_Quary = "select distinct([Date]) from [Contractual_WPC_2G_CS] where Contractor='" + Region + "' order by Date";
                //SqlCommand Date_Quary1 = new SqlCommand(Date_Quary, connection);
                //Date_Quary1.ExecuteNonQuery();
                //Date_Table1 = new DataTable();
                //SqlDataAdapter Date_Table = new SqlDataAdapter(Date_Quary1);
                //Date_Table.Fill(Date_Table1);

                Cell_Quary = "select Cell from [Contractual_WPC_2G_CS] where Contractor='" + Region + "' and Province='" + Province + "' and date=(select max(distinct([Date])) from [Contractual_WPC_2G_CS])";
                SqlCommand Cell_Quary1 = new SqlCommand(Cell_Quary, connection);
                Cell_Quary1.ExecuteNonQuery();
                DataTable Cell_Table1 = new DataTable();
                SqlDataAdapter Cell_Table = new SqlDataAdapter(Cell_Quary1);
                Cell_Table.Fill(Cell_Table1);
                comboBox6.Items.Clear();
                for (int k = 0; k < Cell_Table1.Rows.Count; k++)
                {
                    string cell = (Cell_Table1.Rows[k]).ItemArray[0].ToString();
                    comboBox6.Items.Add(cell);
                }

            }
            if (Technology == "2G_PS")
            {
                //Update_Quary = "update [Contractual_WPC_2G_PS] set[Contractual_WPC_2G_PS].[Status] = 'Near Contractual WPC' where [Contractual_WPC_2G_PS].[Status] is null";
                Node_List_Quary = "select distinct([BSC]) from [Contractual_WPC_2G_PS] where Contractor='" + Region + "' and Province='" + Province + "' and date=(select max(distinct([Date])) from [Contractual_WPC_2G_PS]  where Contractor='" + Region + "') order by BSC";
                //   SqlCommand Update_Quary1 = new SqlCommand(Update_Quary, connection);
                //    Update_Quary1.ExecuteNonQuery();

                //Date_Quary = "select distinct([Date]) from [Contractual_WPC_2G_PS] where Contractor='" + Region + "' order by Date";
                //SqlCommand Date_Quary1 = new SqlCommand(Date_Quary, connection);
                //Date_Quary1.ExecuteNonQuery();
                //Date_Table1 = new DataTable();
                //SqlDataAdapter Date_Table = new SqlDataAdapter(Date_Quary1);
                //Date_Table.Fill(Date_Table1);

                Cell_Quary = "select Cell from [Contractual_WPC_2G_PS] where Contractor='" + Region + "' and Province='" + Province + "' and date=(select max(distinct([Date])) from [Contractual_WPC_2G_PS])";
                SqlCommand Cell_Quary1 = new SqlCommand(Cell_Quary, connection);
                Cell_Quary1.ExecuteNonQuery();
                DataTable Cell_Table1 = new DataTable();
                SqlDataAdapter Cell_Table = new SqlDataAdapter(Cell_Quary1);
                Cell_Table.Fill(Cell_Table1);
                comboBox6.Items.Clear();
                for (int k = 0; k < Cell_Table1.Rows.Count; k++)
                {
                    string cell = (Cell_Table1.Rows[k]).ItemArray[0].ToString();
                    comboBox6.Items.Add(cell);
                }

            }
            if (Technology == "3G_CS")
            {
                //  Update_Quary_CS = "update [Contractual_WPC_3G_CS] set[Contractual_WPC_3G_CS].[Status] = 'Near Contractual WPC' where [Contractual_WPC_3G_CS].[Status] is null";
                Node_List_Quary = "select distinct([RNC]) from [Contractual_WPC_3G_CS] where Contractor='" + Region + "' and Province='" + Province + "' and date=(select max(distinct([Date])) from [Contractual_WPC_3G_CS]  where Contractor='" + Region + "') order by RNC";
                //  SqlCommand Update_Quary_CS1 = new SqlCommand(Update_Quary_CS, connection);
                //   Update_Quary_CS1.ExecuteNonQuery();

                //Date_Quary = "select distinct([Date]) from [Contractual_WPC_3G_CS] where Contractor='" + Region + "' order by Date";
                //SqlCommand Date_Quary1 = new SqlCommand(Date_Quary, connection);
                //Date_Quary1.ExecuteNonQuery();
                //Date_Table1 = new DataTable();
                //SqlDataAdapter Date_Table = new SqlDataAdapter(Date_Quary1);
                //Date_Table.Fill(Date_Table1);

                Cell_Quary = "select Cell from [Contractual_WPC_3G_CS] where Contractor='" + Region + "' and Province='" + Province + "' and date=(select max(distinct([Date])) from [Contractual_WPC_3G_CS])";
                SqlCommand Cell_Quary1 = new SqlCommand(Cell_Quary, connection);
                Cell_Quary1.ExecuteNonQuery();
                DataTable Cell_Table1 = new DataTable();
                SqlDataAdapter Cell_Table = new SqlDataAdapter(Cell_Quary1);
                Cell_Table.Fill(Cell_Table1);
                comboBox6.Items.Clear();
                for (int k = 0; k < Cell_Table1.Rows.Count; k++)
                {
                    string cell = (Cell_Table1.Rows[k]).ItemArray[0].ToString();
                    comboBox6.Items.Add(cell);
                }

            }
            if (Technology == "3G_PS")
            {
                // Update_Quary_PS = "update [Contractual_WPC_3G_PS] set[Contractual_WPC_3G_PS].[Status] = 'Near Contractual WPC' where [Contractual_WPC_3G_PS].[Status] is null";
                Node_List_Quary = "select distinct([RNC]) from [Contractual_WPC_3G_PS] where Contractor='" + Region + "' and Province='" + Province + "'  and date=(select max(distinct([Date])) from [Contractual_WPC_3G_PS] where Contractor='" + Region + "') order by RNC";
                // SqlCommand Update_Quary_PS1 = new SqlCommand(Update_Quary_PS, connection);
                // Update_Quary_PS1.ExecuteNonQuery();


                //Date_Quary = "select distinct([Date]) from [Contractual_WPC_3G_PS] where Contractor='" + Region + "' order by Date";
                //SqlCommand Date_Quary1 = new SqlCommand(Date_Quary, connection);
                //Date_Quary1.ExecuteNonQuery();
                //Date_Table1 = new DataTable();
                //SqlDataAdapter Date_Table = new SqlDataAdapter(Date_Quary1);
                //Date_Table.Fill(Date_Table1);

                Cell_Quary = "select Cell from[Contractual_WPC_3G_PS] where Contractor='" + Region + "' and Province='" + Province + "' and date=(select max(distinct([Date])) from [Contractual_WPC_3G_PS]  where Contractor='" + Region + "')";
                SqlCommand Cell_Quary1 = new SqlCommand(Cell_Quary, connection);
                Cell_Quary1.ExecuteNonQuery();
                DataTable Cell_Table1 = new DataTable();
                SqlDataAdapter Cell_Table = new SqlDataAdapter(Cell_Quary1);
                Cell_Table.Fill(Cell_Table1);
                comboBox6.Items.Clear();
                for (int k = 0; k < Cell_Table1.Rows.Count; k++)
                {
                    string cell = (Cell_Table1.Rows[k]).ItemArray[0].ToString();
                    comboBox6.Items.Add(cell);
                }

            }
            if (Technology == "4G")
            {
                // Update_Quary = "update [Contractual_WPC_4G] set[Contractual_WPC_4G].[Status] = 'Near Contractual WPC' where [Contractual_WPC_4G].[Status] is null";
                Node_List_Quary = "select distinct([RNC]) from [Contractual_WPC_4G] where Contractor='" + Region + "' and Province='" + Province + "'  and date=(select max(distinct([Date])) from [Contractual_WPC_4G] where Contractor='" + Region + "') order by RNC";
                //  SqlCommand Update_Quary1 = new SqlCommand(Update_Quary, connection);
                //   Update_Quary1.ExecuteNonQuery();


                //Date_Quary = "select distinct([Date]) from [Contractual_WPC_4G] where Contractor='" + Region + "' and Province='" + Province + "' order by Date";
                //SqlCommand Date_Quary1 = new SqlCommand(Date_Quary, connection);
                //Date_Quary1.ExecuteNonQuery();
                //Date_Table1 = new DataTable();
                //SqlDataAdapter Date_Table = new SqlDataAdapter(Date_Quary1);
                //Date_Table.Fill(Date_Table1);

                Cell_Quary = "select eNodeB from [Contractual_WPC_4G] where Contractor='" + Region + "' and Province='" + Province + "' and date=(select max(distinct([Date])) from [Contractual_WPC_4G])";
                SqlCommand Cell_Quary1 = new SqlCommand(Cell_Quary, connection);
                Cell_Quary1.ExecuteNonQuery();
                DataTable Cell_Table1 = new DataTable();
                SqlDataAdapter Cell_Table = new SqlDataAdapter(Cell_Quary1);
                Cell_Table.Fill(Cell_Table1);
                comboBox6.Items.Clear();
                for (int k = 0; k < Cell_Table1.Rows.Count; k++)
                {
                    string cell = (Cell_Table1.Rows[k]).ItemArray[0].ToString();
                    comboBox6.Items.Add(cell);
                }


            }


            // List of Nodes (BSC List or RNC List)
            if (Technology == "2G_CS" || Technology == "2G_PS" || Technology == "3G_CS" || Technology == "3G_PS" || Technology == "4G")
            {
                SqlCommand Node_List_Quary1 = new SqlCommand(Node_List_Quary, connection);
                Node_List_Quary1.ExecuteNonQuery();

                Node_Table = new DataTable();
                SqlDataAdapter dataAdapter = new SqlDataAdapter(Node_List_Quary1);
                dataAdapter.Fill(Node_Table);

            }


            comboBox4.Items.Clear();
            for (int k = 1; k <= Node_Table.Rows.Count; k++)
            {
                string Node_Name = (Node_Table.Rows[k - 1]).ItemArray[0].ToString();
                if (Node_Name != "")
                {
                    comboBox4.Items.Add(Node_Name);
                }
            }


            string Technology1 = "";
            if (Technology == "3G_CS" || Technology == "3G_PS")
            {
                Technology1 = "3G";
            }
            else
            {
                Technology1 = Technology;
            }

            if (Technology == "2G_CS" || Technology == "2G_PS")
            {
                Technology1 = "2G";
            }
            else
            {
                Technology1 = Technology;
            }




            //string Province = comboBox2.SelectedItem.ToString();
            string Single_Province_List_Quary_Contractual = "";
            string Single_Province_List_Quary_NearContractual = "";
            string Single_Province_Misssing_Quary_Contractual = "";
            string Single_Province_Missing_Quary_NearContractual = "";
            if (Technology == "2G_CS")
            {
                Single_Province_List_Quary_Contractual = "select [Date], [Province], [Level], [Status], count(Level) as 'Worst_Count' from [Contractual_WPC_2G_CS] where [Province] = '" + Province + "' and Status = 'Contractual WPC' group by[Date], [Province], [Level], [Status]  order by Date";
                Single_Province_List_Quary_NearContractual = "select [Date], [Province], [Level], [Status], count(Level) as 'Worst_Count' from [Contractual_WPC_2G_CS] where [Province] = '" + Province + "' and Status = 'Near Contractual WPC' group by[Date], [Province], [Level], [Status]  order by Date";
                Single_Province_Misssing_Quary_Contractual = "select [Date], [Province], [Status], [Level], sum([Effeciency Index])*1.1 as 'Missing_Score' from [Contractual_WPC_2G_CS] where [Province] = '" + Province + "' and Status = 'Contractual WPC' group by [Date], [Province], [Status] ,[Level] order by Date";
                Single_Province_Missing_Quary_NearContractual = "select [Date], [Province], [Status], [Level], sum([Effeciency Index]) as 'Missing_Score' from [Contractual_WPC_2G_CS] where [Province] = '" + Province + "' and Status = 'Near Contractual WPC'  and QIxP<100 group by [Date], [Province], [Status] ,[Level] order by Date";
            }
            if (Technology == "2G_PS")
            {
                Single_Province_List_Quary_Contractual = "select [Date], [Province], [Level], [Status], count(Level) as 'Worst_Count' from [Contractual_WPC_2G_PS] where [Province] = '" + Province + "' and Status = 'Contractual WPC' group by[Date], [Province], [Level], [Status]  order by Date";
                Single_Province_List_Quary_NearContractual = "select [Date], [Province], [Level], [Status], count(Level) as 'Worst_Count' from [Contractual_WPC_2G_PS] where [Province] = '" + Province + "' and Status = 'Near WPC' group by[Date], [Province], [Level], [Status]  order by Date";
                Single_Province_Misssing_Quary_Contractual = "select [Date], [Province], [Status], [Level], sum([Effeciency Index])*1.1 as 'Missing_Score' from [Contractual_WPC_2G_PS] where [Province] = '" + Province + "' and Status = 'Contractual WPC' group by [Date], [Province], [Status] ,[Level] order by Date";
                Single_Province_Missing_Quary_NearContractual = "select [Date], [Province], [Status], [Level], sum([Effeciency Index]) as 'Missing_Score' from [Contractual_WPC_2G_PS] where [Province] = '" + Province + "' and Status = 'Near WPC'  and QIxP<100 group by [Date], [Province], [Status] ,[Level] order by Date";
            }
            if (Technology == "3G_CS")
            {
                Single_Province_List_Quary_Contractual = "select [Date], [Province], [Level], [Status], count(Level) as 'Worst_Count' from [Contractual_WPC_3G_CS] where [Province] = '" + Province + "' and Status = 'Contractual WPC' group by[Date], [Province], [Level], [Status]  order by Date";
                Single_Province_List_Quary_NearContractual = "select [Date], [Province], [Level], [Status], count(Level) as 'Worst_Count' from [Contractual_WPC_3G_CS] where [Province] = '" + Province + "' and Status = 'Near Contractual WPC' group by[Date], [Province], [Level], [Status]  order by Date";
                Single_Province_Misssing_Quary_Contractual = "select [Date], [Province], [Status], [Level], sum([Effeciency Index])*1.1 as 'Missing_Score' from [Contractual_WPC_3G_CS] where [Province] = '" + Province + "' and Status = 'Contractual WPC' group by [Date], [Province], [Status] ,[Level] order by Date";
                Single_Province_Missing_Quary_NearContractual = "select [Date], [Province], [Status], [Level], sum([Effeciency Index]) as 'Missing_Score' from [Contractual_WPC_3G_CS] where [Province] = '" + Province + "' and Status = 'Near Contractual WPC' and QIxP<100 group by [Date], [Province], [Status] ,[Level] order by Date";
            }
            if (Technology == "3G_PS")
            {
                Single_Province_List_Quary_Contractual = "select [Date], [Province], [Level], [Status], count(Level) as 'Worst_Count' from [Contractual_WPC_3G_PS] where [Province] = '" + Province + "' and Status = 'Contractual WPC' group by[Date], [Province], [Level], [Status]  order by Date";
                Single_Province_List_Quary_NearContractual = "select [Date], [Province], [Level], [Status], count(Level) as 'Worst_Count' from [Contractual_WPC_3G_PS] where [Province] = '" + Province + "' and Status = 'Near Contractual WPC' group by[Date], [Province], [Level], [Status]  order by Date";
                Single_Province_Misssing_Quary_Contractual = "select [Date], [Province], [Status], [Level], sum([Effeciency Index])*1.1 as 'Missing_Score' from [Contractual_WPC_3G_PS] where [Province] = '" + Province + "' and Status = 'Contractual WPC' group by [Date], [Province], [Status] ,[Level] order by Date";
                Single_Province_Missing_Quary_NearContractual = "select [Date], [Province], [Status], [Level], sum([Effeciency Index]) as 'Missing_Score' from [Contractual_WPC_3G_PS] where [Province] = '" + Province + "' and Status = 'Near Contractual WPC' and QIxP<100 group by [Date], [Province], [Status] ,[Level] order by Date";
            }
            if (Technology == "4G")
            {
                Single_Province_List_Quary_Contractual = "select [Date], [Province],  [Level], [Status], count(Level) as 'Worst_Count' from [Contractual_WPC_4G] where [Province] = '" + Province + "' and  Status = 'Contractual WPC' group by [Date], [Province],  [Level], [Status]  order by Date";
                Single_Province_List_Quary_NearContractual = "select [Date], [Province], [Level], [Status], count(Level) as 'Worst_Count' from [Contractual_WPC_4G] where [Province] = '" + Province + "' and  Status = 'Near Contractual WPC' group by[Date], [Province],  [Level], [Status]  order by Date";
                Single_Province_Misssing_Quary_Contractual = "select [Date], [Province],  [Status], [Level], sum([Effeciency Index])*1.1 as 'Missing_Score' from [Contractual_WPC_4G] where [Province] = '" + Province + "' and  Status = 'Contractual WPC' group by [Date], [Province],  [Status] ,[Level] order by Date";
                Single_Province_Missing_Quary_NearContractual = "select [Date], [Province],  [Status], [Level], sum([Effeciency Index]) as 'Missing_Score' from [Contractual_WPC_4G] where [Province] = '" + Province + "' and  Status = 'Near Contractual WPC' and QIxP<100 group by [Date], [Province],  [Status] ,[Level] order by Date";
            }

            // Worst Cells Count in Contractual WPC
            SqlCommand Single_Province_List_Quary_Contractual1 = new SqlCommand(Single_Province_List_Quary_Contractual, connection);
            Single_Province_List_Quary_Contractual1.ExecuteNonQuery();
            Single_Province_Table_Contractual = new DataTable();
            SqlDataAdapter dataAdapter_Contractual = new SqlDataAdapter(Single_Province_List_Quary_Contractual1);
            dataAdapter_Contractual.Fill(Single_Province_Table_Contractual);

            // Worst Cells Count in Near Contractual WPC
            SqlCommand Single_Province_List_Quary_NearContractual1 = new SqlCommand(Single_Province_List_Quary_NearContractual, connection);
            Single_Province_List_Quary_NearContractual1.ExecuteNonQuery();
            DataTable Single_Province_Table_NearContractual = new DataTable();
            SqlDataAdapter dataAdapter_NearContractual = new SqlDataAdapter(Single_Province_List_Quary_NearContractual1);
            dataAdapter_NearContractual.Fill(Single_Province_Table_NearContractual);

            // Worst Cells Missing Score in Contractual WPC
            SqlCommand Single_Province_Missing_Quary_Contractual1 = new SqlCommand(Single_Province_Misssing_Quary_Contractual, connection);
            Single_Province_Missing_Quary_Contractual1.ExecuteNonQuery();
            DataTable Single_Province_Table_Contractual1 = new DataTable();
            SqlDataAdapter dataAdapter_Contractual1 = new SqlDataAdapter(Single_Province_Missing_Quary_Contractual1);
            dataAdapter_Contractual1.Fill(Single_Province_Table_Contractual1);

            // Worst Cells Missing Score in Near Contractual WPC
            SqlCommand Single_Province_Missing_Quary_NearContractual1 = new SqlCommand(Single_Province_Missing_Quary_NearContractual, connection);
            Single_Province_Missing_Quary_NearContractual1.ExecuteNonQuery();
            DataTable Single_Province_Table_NearContractual1 = new DataTable();
            SqlDataAdapter dataAdapter_NearContractual1 = new SqlDataAdapter(Single_Province_Missing_Quary_NearContractual1);
            dataAdapter_NearContractual1.Fill(Single_Province_Table_NearContractual1);


            chart1.Series.Clear();
            chart1.Titles.Clear();

            Series newSeries1 = new Series();
            chart1.Series.Add(newSeries1);
            newSeries1.IsXValueIndexed = false;
            chart1.Series[0].ChartType = SeriesChartType.Line;
            chart1.Series[0].Color = Color.Red;
            chart1.Series[0].BorderWidth = 3;
            chart1.ChartAreas[0].AxisX.Interval = 5;
            chart1.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart1.Series[0].EmptyPointStyle.Color = Color.Transparent;
            chart1.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";
            chart1.Series[0].ToolTip = "#VALX [#VALY]";
            chart1.Series[0].IsValueShownAsLabel = false;
            chart1.Series[0].LegendText = "Level 1";
            chart1.Legends["Legend1"].Docking = Docking.Bottom;
            newSeries1.MarkerStyle = MarkerStyle.Circle;
            newSeries1.MarkerSize = 6;

            Series newSeries2 = new Series();
            chart1.Series.Add(newSeries2);
            newSeries2.IsXValueIndexed = false;
            chart1.Series[1].ChartType = SeriesChartType.Line;
            chart1.Series[1].Color = Color.Orange;
            chart1.Series[1].BorderWidth = 3;
            chart1.Series[1].EmptyPointStyle.Color = Color.Transparent;
            chart1.Series[1].ToolTip = "#VALX [#VALY]";
            chart1.Series[1].IsValueShownAsLabel = false;
            chart1.Series[1].LegendText = "Level 2";
            newSeries2.MarkerStyle = MarkerStyle.Circle;
            newSeries2.MarkerSize = 6;

            Series newSeries3 = new Series();
            chart1.Series.Add(newSeries3);
            newSeries3.IsXValueIndexed = false;
            chart1.Series[2].ChartType = SeriesChartType.Line;
            chart1.Series[2].Color = Color.Yellow;
            chart1.Series[2].BorderWidth = 3;
            chart1.Series[2].EmptyPointStyle.Color = Color.Transparent;
            chart1.Series[2].ToolTip = "#VALX [#VALY]";
            chart1.Series[2].IsValueShownAsLabel = false;
            chart1.Series[2].LegendText = "Level 3";
            newSeries3.MarkerStyle = MarkerStyle.Circle;
            newSeries3.MarkerSize = 6;


            Series newSeries4 = new Series();
            chart1.Series.Add(newSeries4);
            newSeries4.IsXValueIndexed = false;
            chart1.Series[3].ChartType = SeriesChartType.Line;
            chart1.Series[3].Color = Color.Blue;
            chart1.Series[3].BorderWidth = 3;
            chart1.Series[3].EmptyPointStyle.Color = Color.Transparent;
            chart1.Series[3].ToolTip = "#VALX [#VALY]";
            chart1.Series[3].IsValueShownAsLabel = false;
            chart1.Series[3].LegendText = "Level 4";
            newSeries4.MarkerStyle = MarkerStyle.Circle;
            newSeries4.MarkerSize = 6;


            Series newSeries5 = new Series();
            chart1.Series.Add(newSeries5);
            newSeries5.IsXValueIndexed = false;
            chart1.Series[4].ChartType = SeriesChartType.Line;
            chart1.Series[4].Color = Color.Green;
            chart1.Series[4].BorderWidth = 3;
            chart1.Series[4].EmptyPointStyle.Color = Color.Transparent;
            chart1.Series[4].ToolTip = "#VALX [#VALY]";
            chart1.Series[4].IsValueShownAsLabel = false;
            chart1.Series[4].LegendText = "Level 5";
            newSeries5.MarkerStyle = MarkerStyle.Circle;
            newSeries5.MarkerSize = 6;


            chart2.Series.Clear();
            chart2.Titles.Clear();

            Series newSeries6 = new Series();
            chart2.Series.Add(newSeries6);
            newSeries6.IsXValueIndexed = false;
            chart2.Series[0].ChartType = SeriesChartType.Line;
            chart2.Series[0].Color = Color.Red;
            chart2.Series[0].BorderWidth = 3;
            chart2.ChartAreas[0].AxisX.Interval = 5;
            chart2.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart2.Series[0].EmptyPointStyle.Color = Color.Transparent;
            chart2.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";
            chart2.Series[0].ToolTip = "#VALX [#VALY]";
            chart2.Series[0].IsValueShownAsLabel = false;
            chart2.Series[0].LegendText = "Level 1";
            chart2.Legends["Legend1"].Docking = Docking.Bottom;
            newSeries6.MarkerStyle = MarkerStyle.Circle;
            newSeries6.MarkerSize = 6;


            Series newSeries7 = new Series();
            chart2.Series.Add(newSeries7);
            newSeries7.IsXValueIndexed = false;
            chart2.Series[1].ChartType = SeriesChartType.Line;
            chart2.Series[1].Color = Color.Orange;
            chart2.Series[1].BorderWidth = 3;
            chart2.Series[1].EmptyPointStyle.Color = Color.Transparent;
            chart2.Series[1].ToolTip = "#VALX [#VALY]";
            chart2.Series[1].IsValueShownAsLabel = false;
            chart2.Series[1].LegendText = "Level 2";
            newSeries7.MarkerStyle = MarkerStyle.Circle;
            newSeries7.MarkerSize = 6;


            Series newSeries8 = new Series();
            chart2.Series.Add(newSeries8);
            newSeries8.IsXValueIndexed = false;
            chart2.Series[2].ChartType = SeriesChartType.Line;
            chart2.Series[2].Color = Color.Yellow;
            chart2.Series[2].BorderWidth = 3;
            chart2.Series[2].EmptyPointStyle.Color = Color.Transparent;
            chart2.Series[2].ToolTip = "#VALX [#VALY]";
            chart2.Series[2].IsValueShownAsLabel = false;
            chart2.Series[2].LegendText = "Level 3";
            newSeries8.MarkerStyle = MarkerStyle.Circle;
            newSeries8.MarkerSize = 6;


            Series newSeries9 = new Series();
            chart2.Series.Add(newSeries9);
            newSeries9.IsXValueIndexed = false;
            chart2.Series[3].ChartType = SeriesChartType.Line;
            chart2.Series[3].Color = Color.Blue;
            chart2.Series[3].BorderWidth = 3;
            chart2.Series[3].EmptyPointStyle.Color = Color.Transparent;
            chart2.Series[3].ToolTip = "#VALX [#VALY]";
            chart2.Series[3].IsValueShownAsLabel = false;
            chart2.Series[3].LegendText = "Level 4";
            newSeries9.MarkerStyle = MarkerStyle.Circle;
            newSeries9.MarkerSize = 6;


            Series newSeries10 = new Series();
            chart2.Series.Add(newSeries10);
            newSeries10.IsXValueIndexed = false;
            chart2.Series[4].ChartType = SeriesChartType.Line;
            chart2.Series[4].Color = Color.Green;
            chart2.Series[4].BorderWidth = 3;
            chart2.Series[4].EmptyPointStyle.Color = Color.Transparent;
            chart2.Series[4].ToolTip = "#VALX [#VALY]";
            chart2.Series[4].IsValueShownAsLabel = false;
            chart2.Series[4].LegendText = "Level 5";
            newSeries10.MarkerStyle = MarkerStyle.Circle;
            newSeries10.MarkerSize = 6;




            Title title1 = chart1.Titles.Add("Number of '" + Province + "-" + Technology + "' Worst Cells (Contractual)");
            title1.Font = new System.Drawing.Font("Arial", 12, FontStyle.Regular);
            Title title2 = chart2.Titles.Add("Number of '" + Province + "-" + Technology + "' Worst Cells (Near Contractual)");
            title2.Font = new System.Drawing.Font("Arial", 12, FontStyle.Regular);




            chart3.Series.Clear();
            chart3.Titles.Clear();

            Series newSeries11 = new Series();
            chart3.Series.Add(newSeries11);
            newSeries11.IsXValueIndexed = false;
            chart3.Series[0].ChartType = SeriesChartType.StackedColumn;
            chart3.Series[0].Color = Color.Brown;
            chart3.Series[0].BorderWidth = 3;
            chart3.ChartAreas[0].AxisX.Interval = 5;
            chart3.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart3.Series[0].EmptyPointStyle.Color = Color.Transparent;
            chart3.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";
            chart3.Series[0].ToolTip = "#VALX [#VALY]";
            chart3.Series[0].IsValueShownAsLabel = false;
            chart3.Series[0].IsVisibleInLegend = false;
            newSeries11.MarkerStyle = MarkerStyle.Circle;
            newSeries11.MarkerSize = 6;


            chart4.Series.Clear();
            chart4.Titles.Clear();

            Series newSeries12 = new Series();
            chart4.Series.Add(newSeries12);
            newSeries12.IsXValueIndexed = false;
            chart4.Series[0].ChartType = SeriesChartType.StackedColumn;
            chart4.Series[0].Color = Color.Brown;
            chart4.Series[0].BorderWidth = 3;
            chart4.ChartAreas[0].AxisX.Interval = 5;
            chart4.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart4.Series[0].EmptyPointStyle.Color = Color.Transparent;
            chart4.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";
            chart4.Series[0].ToolTip = "#VALX [#VALY]";
            chart4.Series[0].IsValueShownAsLabel = false;
            chart4.Series[0].IsVisibleInLegend = false;
            newSeries12.MarkerStyle = MarkerStyle.Circle;
            newSeries12.MarkerSize = 6;



            Title title3 = chart3.Titles.Add("Sum of '" + Province + "-" + Technology + "' Missing Score (Contractual)");
            title3.Font = new System.Drawing.Font("Arial", 12, FontStyle.Regular);
            Title title4 = chart4.Titles.Add("Sum of '" + Province + "-" + Technology + "' Missing Score (Near Contractual)");
            title4.Font = new System.Drawing.Font("Arial", 12, FontStyle.Regular);



            int y1 = 0;
            for (int k = 1; k <= Single_Province_Table_Contractual.Rows.Count; k++)
            {
                DateTime dt1 = Convert.ToDateTime((Single_Province_Table_Contractual.Rows[k - 1]).ItemArray[0]);
                dt1 = dt1.AddHours(23);
                double Worst_Cell_Count1 = 0;
                double Level1 = 0;
                if (Technology == "2G_CS" || Technology == "2G_PS" || Technology == "3G_CS" || Technology == "3G_PS")
                {
                    Worst_Cell_Count1 = Convert.ToDouble((Single_Province_Table_Contractual.Rows[k - 1]).ItemArray[4]);
                    Level1 = Convert.ToDouble((Single_Province_Table_Contractual.Rows[k - 1]).ItemArray[2]);
                }
                if (Technology == "4G")
                {
                    Worst_Cell_Count1 = Convert.ToDouble((Single_Province_Table_Contractual.Rows[k - 1]).ItemArray[4]);
                    Level1 = Convert.ToDouble((Single_Province_Table_Contractual.Rows[k - 1]).ItemArray[2]);
                }


                // Setting of Intervals
                // **************************************************************
                double dt1_double = dt1.Year * 10000 + dt1.Month * 100 + dt1.Day;

                if (dt1_double > Max_X2)
                {
                    Max_X2 = dt1_double;
                    Max_X_Date = dt1;
                }
                if (dt1_double < Min_X2)
                {
                    Min_X2 = dt1_double;
                    Min_X_Date = dt1;
                }
                // **************************************************************





                if (Level1 == 1 && checkBox6.Checked == true)
                {
                    chart1.Series[0].Points.AddXY(dt1, Worst_Cell_Count1);
                }
                if (Level1 == 2 && checkBox7.Checked == true)
                {
                    chart1.Series[1].Points.AddXY(dt1, Worst_Cell_Count1);
                }
                if (Level1 == 3 && checkBox8.Checked == true)
                {
                    chart1.Series[2].Points.AddXY(dt1, Worst_Cell_Count1);
                }
                if (Level1 == 4 && checkBox9.Checked == true)
                {
                    chart1.Series[3].Points.AddXY(dt1, Worst_Cell_Count1);
                }
                if (Level1 == 5 && checkBox10.Checked == true)
                {
                    chart1.Series[4].Points.AddXY(dt1, Worst_Cell_Count1);
                }
                Max_Y1 = chart1.ChartAreas[0].AxisY.Maximum;
                if (k == 1)
                {
                    Min_X = dt1;
                }

                y1 = k;
            }
            Max_X = Convert.ToDateTime((Single_Province_Table_Contractual.Rows[y1 - 1]).ItemArray[0]);



            // Setting of Intervals
            // **************************************************************
            double difference_day = (Max_X_Date - Min_X_Date).TotalDays;
            double day_interval = Math.Round(difference_day / 20);
            if (day_interval == 0)
            {
                day_interval = 1;
            }
            chart1.ChartAreas[0].AxisX.Interval = day_interval;
            chart2.ChartAreas[0].AxisX.Interval = day_interval;
            chart3.ChartAreas[0].AxisX.Interval = day_interval;
            chart4.ChartAreas[0].AxisX.Interval = day_interval;
            // **************************************************************




            int y2 = 0;
            for (int k = 1; k <= Single_Province_Table_NearContractual.Rows.Count; k++)
            {

                DateTime dt2 = Convert.ToDateTime((Single_Province_Table_NearContractual.Rows[k - 1]).ItemArray[0]);
                dt2 = dt2.AddHours(23);
                double Worst_Cell_Count2 = 0;
                double Level2 = 0;
                if (Technology == "2G_CS" || Technology == "2G_PS" || Technology == "3G_CS" || Technology == "3G_PS")
                {
                    Worst_Cell_Count2 = Convert.ToDouble((Single_Province_Table_NearContractual.Rows[k - 1]).ItemArray[4]);
                    Level2 = Convert.ToDouble((Single_Province_Table_NearContractual.Rows[k - 1]).ItemArray[2]);
                }
                if (Technology == "4G")
                {
                    Worst_Cell_Count2 = Convert.ToDouble((Single_Province_Table_NearContractual.Rows[k - 1]).ItemArray[4]);
                    Level2 = Convert.ToDouble((Single_Province_Table_NearContractual.Rows[k - 1]).ItemArray[2]);
                }

                if (Level2 == 1 && checkBox6.Checked == true)
                {
                    chart2.Series[0].Points.AddXY(dt2, Worst_Cell_Count2);
                }
                if (Level2 == 2 && checkBox7.Checked == true)
                {
                    chart2.Series[1].Points.AddXY(dt2, Worst_Cell_Count2);
                }
                if (Level2 == 3 && checkBox8.Checked == true)
                {
                    chart2.Series[2].Points.AddXY(dt2, Worst_Cell_Count2);
                }
                if (Level2 == 4 && checkBox9.Checked == true)
                {
                    chart2.Series[3].Points.AddXY(dt2, Worst_Cell_Count2);
                }
                if (Level2 == 5 && checkBox10.Checked == true)
                {
                    chart2.Series[4].Points.AddXY(dt2, Worst_Cell_Count2);
                }
                Max_Y2 = chart2.ChartAreas[0].AxisY.Maximum;
                if (k == 1)
                {
                    Min_X = dt2;
                }
                y2 = k;
            }
            Max_X = Convert.ToDateTime((Single_Province_Table_NearContractual.Rows[y2 - 1]).ItemArray[0]);

            for (int k = 1; k <= Single_Province_Table_Contractual1.Rows.Count; k++)
            {
                DateTime dt1 = Convert.ToDateTime((Single_Province_Table_Contractual1.Rows[k - 1]).ItemArray[0]);
                dt1 = dt1.AddHours(23);
                double Level = 0;
                if (Technology == "2G_CS" || Technology == "2G_PS" || Technology == "3G_CS" || Technology == "3G_PS")
                {
                    Score_Cell_Sum = Convert.ToDouble((Single_Province_Table_Contractual1.Rows[k - 1]).ItemArray[4]);
                    Level = Convert.ToDouble((Single_Province_Table_Contractual1.Rows[k - 1]).ItemArray[3]);

                }
                if (Technology == "4G")
                {
                    Score_Cell_Sum = Convert.ToDouble((Single_Province_Table_Contractual1.Rows[k - 1]).ItemArray[4]);
                    Level = Convert.ToDouble((Single_Province_Table_Contractual1.Rows[k - 1]).ItemArray[3]);
                }


                if (k == 1)
                {
                    Score_Cell_Sum1 = 0;
                }
                if (k > 1)
                {
                    DateTime dt_old = Convert.ToDateTime((Single_Province_Table_Contractual1.Rows[k - 2]).ItemArray[0]);
                    dt_old = dt_old.AddHours(23);
                    if (dt1 != dt_old)
                    {
                        Score_Cell_Sum1 = 0;
                    }
                }

                if (Level == 1 && checkBox6.Checked == true)
                {
                    Score_Cell_Sum1 = Score_Cell_Sum1 + Score_Cell_Sum;
                }
                if (Level == 2 && checkBox7.Checked == true)
                {
                    Score_Cell_Sum1 = Score_Cell_Sum1 + Score_Cell_Sum;
                }
                if (Level == 3 && checkBox8.Checked == true)
                {
                    Score_Cell_Sum1 = Score_Cell_Sum1 + Score_Cell_Sum;
                }
                if (Level == 4 && checkBox9.Checked == true)
                {
                    Score_Cell_Sum1 = Score_Cell_Sum1 + Score_Cell_Sum;
                }
                if (Level == 5 && checkBox10.Checked == true)
                {
                    Score_Cell_Sum1 = Score_Cell_Sum1 + Score_Cell_Sum;
                }

                chart3.Series[0].Points.AddXY(dt1, Score_Cell_Sum1);

            }

            for (int k = 1; k <= Single_Province_Table_NearContractual1.Rows.Count; k++)
            {
                DateTime dt1 = Convert.ToDateTime((Single_Province_Table_NearContractual1.Rows[k - 1]).ItemArray[0]);
                dt1 = dt1.AddHours(23);
                double Level = 0;
                if (Technology == "2G_CS" || Technology == "2G_PS" || Technology == "3G_CS" || Technology == "3G_PS")
                {
                    Score_Cell_Sum = Convert.ToDouble((Single_Province_Table_NearContractual1.Rows[k - 1]).ItemArray[4]);
                    Level = Convert.ToDouble((Single_Province_Table_NearContractual1.Rows[k - 1]).ItemArray[3]);

                }
                if (Technology == "4G")
                {
                    Score_Cell_Sum = Convert.ToDouble((Single_Province_Table_NearContractual1.Rows[k - 1]).ItemArray[4]);
                    Level = Convert.ToDouble((Single_Province_Table_NearContractual1.Rows[k - 1]).ItemArray[3]);

                }

                if (k == 1)
                {
                    Score_Cell_Sum1 = 0;
                }
                if (k > 1)
                {
                    DateTime dt_old = Convert.ToDateTime((Single_Province_Table_NearContractual1.Rows[k - 2]).ItemArray[0]);
                    dt_old = dt_old.AddHours(23);
                    if (dt1 != dt_old)
                    {
                        Score_Cell_Sum1 = 0;
                    }
                }

                if (Level == 1 && checkBox6.Checked == true)
                {
                    Score_Cell_Sum1 = Score_Cell_Sum1 + Score_Cell_Sum;
                }
                if (Level == 2 && checkBox7.Checked == true)
                {
                    Score_Cell_Sum1 = Score_Cell_Sum1 + Score_Cell_Sum;
                }
                if (Level == 3 && checkBox8.Checked == true)
                {
                    Score_Cell_Sum1 = Score_Cell_Sum1 + Score_Cell_Sum;
                }
                if (Level == 4 && checkBox9.Checked == true)
                {
                    Score_Cell_Sum1 = Score_Cell_Sum1 + Score_Cell_Sum;
                }
                if (Level == 5 && checkBox10.Checked == true)
                {
                    Score_Cell_Sum1 = Score_Cell_Sum1 + Score_Cell_Sum;
                }

                chart4.Series[0].Points.AddXY(dt1, Score_Cell_Sum1);

            }




        }





        public double Y_max1 = 0;
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            Y_max1 = 0;
            if (textBox1.Text != "")
            {
                Y_max1 = Convert.ToDouble(textBox1.Text);
                chart1.ChartAreas[0].AxisY.Maximum = Y_max1;
            }
            if (textBox1.Text == "")
            {
                chart1.ChartAreas[0].AxisY.Maximum = Max_Y1;
            }

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            double Y_max2 = 0;
            if (textBox2.Text != "")
            {
                Y_max2 = Convert.ToDouble(textBox2.Text);
                chart2.ChartAreas[0].AxisY.Maximum = Y_max2;
            }
            if (textBox2.Text == "")
            {
                chart2.ChartAreas[0].AxisY.Maximum = Max_Y2;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            chart1.SaveImage("Image.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
            Image img = Image.FromFile("Image.jpg");
            System.Windows.Forms.Clipboard.SetImage(img);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            chart2.SaveImage("Image.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
            Image img = Image.FromFile("Image.jpg");
            System.Windows.Forms.Clipboard.SetImage(img);
        }


        private void button3_Click(object sender, EventArgs e)
        {
            chart3.SaveImage("Image.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
            Image img = Image.FromFile("Image.jpg");
            System.Windows.Forms.Clipboard.SetImage(img);
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            chart3.SaveImage("Image.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
            Image img = Image.FromFile("Image.jpg");
            System.Windows.Forms.Clipboard.SetImage(img);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            chart4.SaveImage("Image.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
            Image img = Image.FromFile("Image.jpg");
            System.Windows.Forms.Clipboard.SetImage(img);
        }


        private void checkBox10_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox10.Checked == true)
            {
                chart1.Series[4].Enabled = true;
                chart2.Series[4].Enabled = true;
                if (Region_Show == 1)
                {
                    chart3.Series[4].Enabled = true;
                    chart4.Series[4].Enabled = true;
                }
            }
            if (checkBox10.Checked == false)
            {
                chart1.Series[4].Enabled = false;
                chart2.Series[4].Enabled = false;
                if (Region_Show == 1)
                {
                    chart3.Series[4].Enabled = false;
                    chart4.Series[4].Enabled = false;
                }
            }
        }

        private void checkBox9_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox9.Checked == true)
            {
                chart1.Series[3].Enabled = true;
                chart2.Series[3].Enabled = true;
                if (Region_Show == 1)
                {
                    chart3.Series[3].Enabled = true;
                    chart4.Series[3].Enabled = true;
                }
            }
            if (checkBox9.Checked == false)
            {
                chart1.Series[3].Enabled = false;
                chart2.Series[3].Enabled = false;
                if (Region_Show == 1)
                {
                    chart3.Series[3].Enabled = false;
                    chart4.Series[3].Enabled = false;
                }
            }
        }

        private void checkBox8_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox8.Checked == true)
            {
                chart1.Series[2].Enabled = true;
                chart2.Series[2].Enabled = true;
                if (Region_Show == 1)
                {
                    chart3.Series[2].Enabled = true;
                    chart4.Series[2].Enabled = true;
                }
            }
            if (checkBox8.Checked == false)
            {
                chart1.Series[2].Enabled = false;
                chart2.Series[2].Enabled = false;
                if (Region_Show == 1)
                {
                    chart3.Series[2].Enabled = false;
                    chart4.Series[2].Enabled = false;
                }
            }
        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox7.Checked == true)
            {
                chart1.Series[1].Enabled = true;
                chart2.Series[1].Enabled = true;
                if (Region_Show == 1)
                {
                    chart3.Series[1].Enabled = true;
                    chart4.Series[1].Enabled = true;
                }
            }
            if (checkBox7.Checked == false)
            {
                chart1.Series[1].Enabled = false;
                chart2.Series[1].Enabled = false;
                if (Region_Show == 1)
                {
                    chart3.Series[1].Enabled = false;
                    chart4.Series[1].Enabled = false;
                }
            }
        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox6.Checked == true)
            {
                chart1.Series[0].Enabled = true;
                chart2.Series[0].Enabled = true;
                if (Region_Show == 1)
                {
                    chart3.Series[0].Enabled = true;
                    chart4.Series[0].Enabled = true;
                }
            }
            if (checkBox6.Checked == false)
            {
                chart1.Series[0].Enabled = false;
                chart2.Series[0].Enabled = false;
                if (Region_Show == 1)
                {
                    chart3.Series[0].Enabled = false;
                    chart4.Series[0].Enabled = false;
                }
            }
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            Min_X1 = dateTimePicker1.Value;
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            Max_X1 = dateTimePicker2.Value;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            chart1.ChartAreas[0].AxisX.Minimum = Min_X1.ToOADate();
            chart2.ChartAreas[0].AxisX.Minimum = Min_X1.ToOADate();
            chart3.ChartAreas[0].AxisX.Minimum = Min_X1.ToOADate();
            chart4.ChartAreas[0].AxisX.Minimum = Min_X1.ToOADate();

            chart1.ChartAreas[0].AxisX.Maximum = Max_X1.ToOADate();
            chart2.ChartAreas[0].AxisX.Maximum = Max_X1.ToOADate();
            chart3.ChartAreas[0].AxisX.Maximum = Max_X1.ToOADate();
            chart4.ChartAreas[0].AxisX.Maximum = Max_X1.ToOADate();

            double difference_day = (Max_X1 - Min_X1).TotalDays;
            double day_interval = Math.Round(difference_day / 20);
            if (day_interval == 0)
            {
                day_interval = 1;
            }
            chart1.ChartAreas[0].AxisX.Interval = day_interval;
            chart2.ChartAreas[0].AxisX.Interval = day_interval;
            chart3.ChartAreas[0].AxisX.Interval = day_interval;
            chart4.ChartAreas[0].AxisX.Interval = day_interval;

        }

        private void button8_Click(object sender, EventArgs e)
        {
            chart1.ChartAreas[0].AxisX.Minimum = Min_X.ToOADate() - 1;
            chart1.ChartAreas[0].AxisX.Maximum = Max_X.ToOADate() + 1;
            chart2.ChartAreas[0].AxisX.Minimum = Min_X.ToOADate() - 1;
            chart2.ChartAreas[0].AxisX.Maximum = Max_X.ToOADate() + 1;
            chart3.ChartAreas[0].AxisX.Minimum = Min_X.ToOADate() - 1;
            chart3.ChartAreas[0].AxisX.Maximum = Max_X.ToOADate() + 1;
            chart4.ChartAreas[0].AxisX.Minimum = Min_X.ToOADate() - 1;
            chart4.ChartAreas[0].AxisX.Maximum = Max_X.ToOADate() + 1;
        }



        private void dateTimePicker3_ValueChanged(object sender, EventArgs e)
        {
            checking_date = dateTimePicker3.Value;
        }


        private void comboBox8_SelectedIndexChanged(object sender, EventArgs e)
        {
            checking_day_num = Convert.ToInt16(comboBox8.SelectedItem);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DateTime Selection_Time = Node_Selection_Time;
            DateTime Selection_Time1 = Region_Selection_Time;
            string Temp_Quary = "";
            string Temp_Drop = "";


            string selection_item = "";
            if (Node_Selection_Time > Region_Selection_Time && Node_Selection_Time > Owner_Selection_Time)
            {
                selection_item = "Node";
            }
            if (Region_Selection_Time > Node_Selection_Time && Region_Selection_Time > Owner_Selection_Time)
            {
                selection_item = "Region";

                if (Technology == "2G_CS")
                {

                    int day_counter_before = 0;
                    string Dates_before = "";
                    for (int k = 1; k <= Date_Table1.Rows.Count; k++)
                    {
                        int ind = Date_Table1.Rows.Count - k + 1;
                        DateTime Date = Convert.ToDateTime((Date_Table1.Rows[ind - 1]).ItemArray[0]);
                        if (Date < checking_date.Date && day_counter_before < checking_day_num)
                        {
                            day_counter_before++;
                            if (day_counter_before == 1)
                            {
                                Dates_before = "" + Date;
                            }
                            if (day_counter_before > 1)
                            {
                                Dates_before = Dates_before + "' or Date='" + Date;
                            }
                        }
                    }


                    Solved_Quary = @"select tbl2.Date, tbl2.Cell, tbl2.BSC, tbl2.Vendor, tbl2.Contractor, tbl2.LEVEL, tbl2.Worst, tbl2.[Effeciency Index], tbl2.Status, tbl2.QIx, tbl2.QIxP, tbl2.QIxBL, tbl2.[AVG_TCH_Traffic], tbl2.[Repeated Daily at 10 days (QIx<QIb)], tbl2.[Total Days Availability > 90],
                    tbl2.[Worst(%) of CSSR], tbl2.[Worst(%) of OHSR], tbl2.[Worst(%) of CDR], tbl2.[Worst(%) of TCH_Assignment_FR], tbl2.[Worst(%) of DL Quality <=4], tbl2.[Worst(%) of UL Quality <=4],
tbl2.[Worst(%) of SDCCH_Congestion_Rate], tbl2.[Worst(%) of SDCCH_Access_Success_Rate], tbl2.[Worst(%) of SDCCH_Drop_Rate], tbl2.[Worst(%) of IHSR], tbl2.[Worst(%) of AMRHR_Usage]  from
(select[Date], [Cell], [BSC], [Vendor], [Contractor], [LEVEL], [Worst], [Effeciency Index], [Status], [QIx], [QIxP], [QIxBL], [AVG_TCH_Traffic], [Repeated Daily at 10 days (QIx<QIb)], [Total Days Availability > 90],
[Worst(%) of CSSR], [Worst(%) of OHSR], [Worst(%) of CDR], [Worst(%) of TCH_Assignment_FR], [Worst(%) of DL Quality <=4], [Worst(%) of UL Quality <=4],
[Worst(%) of SDCCH_Congestion_Rate], [Worst(%) of SDCCH_Access_Success_Rate], [Worst(%) of SDCCH_Drop_Rate], [Worst(%) of IHSR], [Worst(%) of AMRHR_Usage] from[dbo].[Contractual_WPC_2G_CS] where Contractor = '" + Region + "' and (Date ='" + Dates_before + @"')) as tbl2
left join
(select[Date], [Cell], [BSC], [Vendor], [Contractor], [LEVEL], [Worst], [Effeciency Index], [Status],  [QIxP], [QIxBL], [AVG_TCH_Traffic], [Repeated Daily at 10 days (QIx<QIb)], [Total Days Availability > 90],
[Worst(%) of CSSR], [Worst(%) of OHSR], [Worst(%) of CDR], [Worst(%) of TCH_Assignment_FR], [Worst(%) of DL Quality <=4], [Worst(%) of UL Quality <=4],
[Worst(%) of SDCCH_Congestion_Rate], [Worst(%) of SDCCH_Access_Success_Rate], [Worst(%) of SDCCH_Drop_Rate], [Worst(%) of IHSR], [Worst(%) of AMRHR_Usage] from[dbo].[Contractual_WPC_2G_CS] where Contractor = '" + Region + "' and Date ='" + checking_date.Date + @"') as tbl1
on[tbl1].Cell =[tbl2].Cell where[tbl1].Cell is Null";



                    SqlCommand Solved_Quary1 = new SqlCommand(Solved_Quary, connection);
                    Solved_Quary1.ExecuteNonQuery();
                    Solved_Data_Table = new DataTable();
                    SqlDataAdapter Solved_Date_Table1 = new SqlDataAdapter(Solved_Quary1);
                    Solved_Date_Table1.Fill(Solved_Data_Table);
                }


                if (Technology == "3G_CS" || Technology == "3G_PS")
                {

                    int day_counter_before = 0;
                    string Dates_before = "";
                    for (int k = 1; k <= Date_Table1.Rows.Count; k++)
                    {
                        int ind = Date_Table1.Rows.Count - k + 1;
                        DateTime Date = Convert.ToDateTime((Date_Table1.Rows[ind - 1]).ItemArray[0]);
                        if (Date < checking_date.Date && day_counter_before < checking_day_num)
                        {
                            day_counter_before++;
                            if (day_counter_before == 1)
                            {
                                Dates_before = "" + Date;
                            }
                            if (day_counter_before > 1)
                            {
                                Dates_before = Dates_before + "' or Date='" + Date;
                            }
                        }
                    }

                    if (Technology == "3G_CS")
                    {
                        Solved_Quary = @"select tbl2.Date, tbl2.Cell, tbl2.RNC, tbl2.Vendor, tbl2.Contractor, tbl2.LEVEL, tbl2.Worst, tbl2.[Effeciency Index], tbl2.Status, tbl2.QIx, tbl2.QIxP, tbl2.QIxBL, tbl2.[AVG_CS_Traffic)], 	tbl2.[Repeated Daily at 10 days (QIx<QIb)], 	tbl2.[Total Days Availability > 90], 	tbl2.[Worst(%) of RAB Establishment  Success Rate (CS)], 	tbl2.[Worst(%) of W2G IRAT/IF HO success rate], 	tbl2.[Worst(%) of Drop Call Rate], 	tbl2.[Worst(%) of Soft HO Success Rate], 	tbl2.[Worst(%) of CS RRC Connection Establishment SR (%)] from
(select [Date], [Cell], [RNC], [Vendor], [Contractor], [LEVEL], [Worst], [Effeciency Index], [Status], [QIx], [QIxP], [QIxBL], [AVG_CS_Traffic)], 	[Repeated Daily at 10 days (QIx<QIb)], 	[Total Days Availability > 90], 	[Worst(%) of RAB Establishment  Success Rate (CS)], 	[Worst(%) of W2G IRAT/IF HO success rate], 	[Worst(%) of Drop Call Rate], 	[Worst(%) of Soft HO Success Rate], 	[Worst(%) of CS RRC Connection Establishment SR (%)] from [dbo].[Contractual_WPC_3G_CS] where Contractor = '" + Region + "' and (Date ='" + Dates_before + @"')) as tbl2
left join
(select[Date], [Cell], [RNC], [Vendor], [Contractor], [LEVEL], [Worst], [Effeciency Index], [Status],  [QIxP], [QIxBL], [AVG_CS_Traffic)], 	[Repeated Daily at 10 days (QIx<QIb)], 	[Total Days Availability > 90], 	[Worst(%) of RAB Establishment  Success Rate (CS)], 	[Worst(%) of W2G IRAT/IF HO success rate], 	[Worst(%) of Drop Call Rate], 	[Worst(%) of Soft HO Success Rate], 	[Worst(%) of CS RRC Connection Establishment SR (%)] from [dbo].[Contractual_WPC_3G_CS] where Contractor = '" + Region + "' and Date ='" + checking_date.Date + @"') as tbl1
on[tbl1].Cell =[tbl2].Cell where[tbl1].Cell is Null";
                    }
                    if (Technology == "3G_PS")
                    {
                        Solved_Quary = @"select tbl2.Date, tbl2.ElementID1, tbl2.RNC, tbl2.Vendor, tbl2.Contractor, tbl2.LEVEL, tbl2.Worst, tbl2.[Effeciency Index], tbl2.Status, tbl2.QIx, tbl2.QIxP, tbl2.QIxBL, tbl2.[Avg Payload of Cell(GB)], 	tbl2.[Repeated Daily at 10 days (QIx<QIb)], 	tbl2.[Total Days Availability > 90], 	tbl2.[Worst(%) of RAB Establishment  Success Rate (HSDPA)], 	tbl2.[Worst(%) of RAB Establishment  Success Rate (EUL)], 	tbl2.[Worst(%) of EUL MAC User Throughput (kbps)], 	tbl2.[Worst(%) of HSDPA MAC-hs User Throughput Net (kbps)], 	tbl2.[Worst(%) of RAB Drop Rate (HSDPA)], 	tbl2.[Worst(%) of RAB Drop Rate (EUL)], 	tbl2.[Worst(%) of MultiRAB Setup Success Ratio (%)], 	tbl2.[Worst(%) of PS_RRC_Setup_Success_Rate], 	tbl2.[Worst(%) of Ps_RAB_Establish_Success_Rate], 	tbl2.[Worst(%) of PS_Multi_RAB_Establish_Success_Rate], 	tbl2.[Worst(%) of Drop_Call_Rate], 	tbl2.[Worst(%) of HSDPA_Cell_Change_Succ_Rate], 	tbl2.[Worst(%) of HS_share_PAYLOAD_%], 	tbl2.[Worst(%) of HSDPA Cell Throughput (Mbps)] from
(select [Date], [ElementID1], [RNC], [Vendor], [Contractor], [LEVEL], [Worst], [Effeciency Index], [Status], [QIx], [QIxP], [QIxBL], [Avg Payload of Cell(GB)], 	[Repeated Daily at 10 days (QIx<QIb)], 	[Total Days Availability > 90], 	[Worst(%) of RAB Establishment  Success Rate (HSDPA)], 	[Worst(%) of RAB Establishment  Success Rate (EUL)], 	[Worst(%) of EUL MAC User Throughput (kbps)], 	[Worst(%) of HSDPA MAC-hs User Throughput Net (kbps)], 	[Worst(%) of RAB Drop Rate (HSDPA)], 	[Worst(%) of RAB Drop Rate (EUL)], 	[Worst(%) of MultiRAB Setup Success Ratio (%)], 	[Worst(%) of PS_RRC_Setup_Success_Rate], 	[Worst(%) of Ps_RAB_Establish_Success_Rate], 	[Worst(%) of PS_Multi_RAB_Establish_Success_Rate], 	[Worst(%) of Drop_Call_Rate], 	[Worst(%) of HSDPA_Cell_Change_Succ_Rate], 	[Worst(%) of HS_share_PAYLOAD_%], 	[Worst(%) of HSDPA Cell Throughput (Mbps)] from [dbo].[Contractual_WPC_3G_PS] where Contractor = '" + Region + "' and (Date ='" + Dates_before + @"')) as tbl2
left join
(select[Date], [ElementID1], [RNC], [Vendor], [Contractor], [LEVEL], [Worst], [Effeciency Index], [Status],  [QIxP], [QIxBL], [Avg Payload of Cell(GB)], 	[Repeated Daily at 10 days (QIx<QIb)], 	[Total Days Availability > 90], 	[Worst(%) of RAB Establishment  Success Rate (HSDPA)], 	[Worst(%) of RAB Establishment  Success Rate (EUL)], 	[Worst(%) of EUL MAC User Throughput (kbps)], 	[Worst(%) of HSDPA MAC-hs User Throughput Net (kbps)], 	[Worst(%) of RAB Drop Rate (HSDPA)], 	[Worst(%) of RAB Drop Rate (EUL)], 	[Worst(%) of MultiRAB Setup Success Ratio (%)], 	[Worst(%) of PS_RRC_Setup_Success_Rate], 	[Worst(%) of Ps_RAB_Establish_Success_Rate], 	[Worst(%) of PS_Multi_RAB_Establish_Success_Rate], 	[Worst(%) of Drop_Call_Rate], 	[Worst(%) of HSDPA_Cell_Change_Succ_Rate], 	[Worst(%) of HS_share_PAYLOAD_%], 	[Worst(%) of HSDPA Cell Throughput (Mbps)]  from [dbo].[Contractual_WPC_3G_PS] where Contractor = '" + Region + "' and Date ='" + checking_date.Date + @"') as tbl1
on[tbl1].ElementID1 =[tbl2].ElementID1 where[tbl1].ElementID1 is Null";
                    }


                    SqlCommand Solved_Quary1 = new SqlCommand(Solved_Quary, connection);
                    Solved_Quary1.ExecuteNonQuery();
                    Solved_Data_Table = new DataTable();
                    SqlDataAdapter Solved_Date_Table1 = new SqlDataAdapter(Solved_Quary1);
                    Solved_Date_Table1.Fill(Solved_Data_Table);


                }


                if (Technology == "4G")
                {

                    int day_counter_before = 0;
                    string Dates_before = "";
                    for (int k = 1; k <= Date_Table1.Rows.Count; k++)
                    {
                        int ind = Date_Table1.Rows.Count - k + 1;
                        DateTime Date = Convert.ToDateTime((Date_Table1.Rows[ind - 1]).ItemArray[0]);
                        if (Date < checking_date.Date && day_counter_before < checking_day_num)
                        {
                            day_counter_before++;
                            if (day_counter_before == 1)
                            {
                                Dates_before = "" + Date;
                            }
                            if (day_counter_before > 1)
                            {
                                Dates_before = Dates_before + "' or Date='" + Date;
                            }
                        }
                    }

                    Solved_Quary = @"select tbl2.Date, tbl2.eNodeB, tbl2.RNC, tbl2.Vendor, tbl2.Province, tbl2.Contractor, tbl2.LEVEL, tbl2.Worst, tbl2.[Effeciency Index], tbl2.Status, tbl2.QIx, tbl2.QIxP, tbl2.QIxBL, tbl2.[Avg Payload of Cell(GB)], 	tbl2.[Repeated Daily at 10 days (QIx<QIb)], 	tbl2.[Total Days Availability > 90], 	tbl2.[Worst(%) of RRC Connection Establishment Success Rate], 	tbl2.[Worst(%) of ERAB Stablishment Success Rate (Initial)], 	tbl2.[Worst(%) of ERAB Stablishment Success Rate (Added)], 	tbl2.[Worst(%) of DL User Troughput  (Mbps)], 	tbl2.[Worst(%) of UL User Throughput (Mbps)], 	tbl2.[Worst(%) of Handover Success Rate], 	tbl2.[Worst(%) of ERAB Drop rate], 	tbl2.[Worst(%) of UE Context Drop Rate], 	tbl2.[Worst(%) of S1 Signalling Success Rate], 	tbl2.[Worst(%) of Inter Frequency Handover Execution SR (%)], 	tbl2.[Worst(%) of Intra Frequency Handover Execution SR (%)], 	tbl2.[Worst(%) of Average Ul Packet Loss Rate (%)], 	tbl2.[Worst(%) of payload per Carrier] from
(select [Date], [eNodeB], [RNC], [Vendor], [Province], [Contractor], [LEVEL], [Worst], [Effeciency Index], [Status], [QIx], [QIxP], [QIxBL], [Avg Payload of Cell(GB)], 	[Repeated Daily at 10 days (QIx<QIb)], 	[Total Days Availability > 90], 	[Worst(%) of RRC Connection Establishment Success Rate], 	[Worst(%) of ERAB Stablishment Success Rate (Initial)], 	[Worst(%) of ERAB Stablishment Success Rate (Added)], 	[Worst(%) of DL User Troughput  (Mbps)], 	[Worst(%) of UL User Throughput (Mbps)], 	[Worst(%) of Handover Success Rate], 	[Worst(%) of ERAB Drop rate], 	[Worst(%) of UE Context Drop Rate], 	[Worst(%) of S1 Signalling Success Rate], 	[Worst(%) of Inter Frequency Handover Execution SR (%)], 	[Worst(%) of Intra Frequency Handover Execution SR (%)], 	[Worst(%) of Average Ul Packet Loss Rate (%)], 	[Worst(%) of payload per Carrier] from [dbo].[Contractual_WPC_4G] where Contractor = '" + Region + "' and (Date ='" + Dates_before + @"')) as tbl2
left join
(select[Date], [eNodeB], [RNC], [Vendor], [Province], [Contractor], [LEVEL], [Worst], [Effeciency Index], [Status],  [QIxP], [QIxBL], [Avg Payload of Cell(GB)], 	[Repeated Daily at 10 days (QIx<QIb)], 	[Total Days Availability > 90], 	[Worst(%) of RRC Connection Establishment Success Rate], 	[Worst(%) of ERAB Stablishment Success Rate (Initial)], 	[Worst(%) of ERAB Stablishment Success Rate (Added)], 	[Worst(%) of DL User Troughput  (Mbps)], 	[Worst(%) of UL User Throughput (Mbps)], 	[Worst(%) of Handover Success Rate], 	[Worst(%) of ERAB Drop rate], 	[Worst(%) of UE Context Drop Rate], 	[Worst(%) of S1 Signalling Success Rate], 	[Worst(%) of Inter Frequency Handover Execution SR (%)], 	[Worst(%) of Intra Frequency Handover Execution SR (%)], 	[Worst(%) of Average Ul Packet Loss Rate (%)], 	[Worst(%) of payload per Carrier]  from [dbo].[Contractual_WPC_4G] where Contractor = '" + Region + "' and Date ='" + checking_date.Date + @"') as tbl1
on[tbl1].eNodeB =[tbl2].eNodeB where[tbl1].eNodeB is Null";

                    SqlCommand Solved_Quary1 = new SqlCommand(Solved_Quary, connection);
                    Solved_Quary1.ExecuteNonQuery();
                    Solved_Data_Table = new DataTable();
                    SqlDataAdapter Solved_Date_Table1 = new SqlDataAdapter(Solved_Quary1);
                    Solved_Date_Table1.Fill(Solved_Data_Table);



                }



            }



            if (Region_Show == 1)
            {
                label15.Text = "Waiting";
                label15.BackColor = Color.Red;

                XLWorkbook wb = new XLWorkbook();
                wb.Worksheets.Add(Solved_Data_Table, "Solved_" + Region + "_" + Technology);
                var saveFileDialog = new SaveFileDialog
                {
                    FileName = "CWA_Export_Solved_" + Region + "_" + Technology + "_" + Convert.ToString(checking_date.Month) + "." + Convert.ToString(checking_date.Day) + "." + Convert.ToString(checking_date.Year),
                    Filter = "Excel files|*.xlsx",
                    Title = "Save an Excel File"
                };

                saveFileDialog.ShowDialog();

                if (!String.IsNullOrWhiteSpace(saveFileDialog.FileName))
                    wb.SaveAs(saveFileDialog.FileName);


                label15.Text = "Finished!";
                label15.BackColor = Color.Yellow;
            }
            else
            {
                MessageBox.Show("Export Part needs to select a Region!");
            }



            if (Owner_Selection_Time > Node_Selection_Time && Owner_Selection_Time > Region_Selection_Time)
            {
                selection_item = "Owner";
            }





        }

        private void button6_Click(object sender, EventArgs e)
        {
            DateTime Selection_Time = Node_Selection_Time;
            DateTime Selection_Time1 = Region_Selection_Time;
            string Temp_Quary = "";
            string Temp_Drop = "";


            string selection_item = "";
            if (Node_Selection_Time > Region_Selection_Time && Node_Selection_Time > Owner_Selection_Time)
            {
                selection_item = "Node";
            }
            if (Region_Selection_Time > Node_Selection_Time && Region_Selection_Time > Owner_Selection_Time)
            {
                selection_item = "Region";

                if (Technology == "2G_CS")
                {

                    int day_counter_before = 0;
                    string Dates_before = "";
                    for (int k = 1; k <= Date_Table1.Rows.Count; k++)
                    {
                        int ind = Date_Table1.Rows.Count - k + 1;
                        DateTime Date = Convert.ToDateTime((Date_Table1.Rows[ind - 1]).ItemArray[0]);
                        if (Date < checking_date.Date && day_counter_before < checking_day_num)
                        {
                            day_counter_before++;
                            if (day_counter_before == 1)
                            {
                                Dates_before = "" + Date;
                            }
                            if (day_counter_before > 1)
                            {
                                Dates_before = Dates_before + "' or Date='" + Date;
                            }
                        }
                    }


                    Raised_Quary = @"select tbl1.Date, tbl1.Cell, tbl1.BSC, tbl1.Vendor, tbl1.Contractor, tbl1.LEVEL, tbl1.Worst, tbl1.[Effeciency Index], tbl1.Status, tbl1.QIx, tbl1.QIxP, tbl1.QIxBL, tbl1.[AVG_TCH_Traffic], tbl1.[Repeated Daily at 10 days (QIx<QIb)], tbl1.[Total Days Availability > 90],
                    tbl1.[Worst(%) of CSSR], tbl1.[Worst(%) of OHSR], tbl1.[Worst(%) of CDR], tbl1.[Worst(%) of TCH_Assignment_FR], tbl1.[Worst(%) of DL Quality <=4], tbl1.[Worst(%) of UL Quality <=4],
tbl1.[Worst(%) of SDCCH_Congestion_Rate], tbl1.[Worst(%) of SDCCH_Access_Success_Rate], tbl1.[Worst(%) of SDCCH_Drop_Rate], tbl1.[Worst(%) of IHSR], tbl1.[Worst(%) of AMRHR_Usage] from
(select [Date], [Cell], [BSC], [Vendor], [Contractor], [LEVEL], [Worst], [Effeciency Index], [Status], [QIx], [QIxP], [QIxBL], [AVG_TCH_Traffic], [Repeated Daily at 10 days (QIx<QIb)], [Total Days Availability > 90],
[Worst(%) of CSSR], [Worst(%) of OHSR], [Worst(%) of CDR], [Worst(%) of TCH_Assignment_FR], [Worst(%) of DL Quality <=4], [Worst(%) of UL Quality <=4],
[Worst(%) of SDCCH_Congestion_Rate], [Worst(%) of SDCCH_Access_Success_Rate], [Worst(%) of SDCCH_Drop_Rate], [Worst(%) of IHSR], [Worst(%) of AMRHR_Usage] from [dbo].[Contractual_WPC_2G_CS] where Contractor = '" + Region + "' and Date ='" + checking_date.Date + @"') as tbl1
left join
(select[Date], [Cell], [BSC], [Vendor], [Contractor], [LEVEL], [Worst], [Effeciency Index], [Status],  [QIxP], [QIxBL], [AVG_TCH_Traffic], [Repeated Daily at 10 days (QIx<QIb)], [Total Days Availability > 90],
[Worst(%) of CSSR], [Worst(%) of OHSR], [Worst(%) of CDR], [Worst(%) of TCH_Assignment_FR], [Worst(%) of DL Quality <=4], [Worst(%) of UL Quality <=4],
[Worst(%) of SDCCH_Congestion_Rate], [Worst(%) of SDCCH_Access_Success_Rate], [Worst(%) of SDCCH_Drop_Rate], [Worst(%) of IHSR], [Worst(%) of AMRHR_Usage]  from [dbo].[Contractual_WPC_2G_CS] where Contractor = '" + Region + "' and (Date ='" + Dates_before + @"')) as tbl2
on[tbl1].Cell =[tbl2].Cell where[tbl2].Cell is Null";







                    SqlCommand Raised_Quary1 = new SqlCommand(Raised_Quary, connection);
                    Raised_Quary1.ExecuteNonQuery();
                    Raised_Data_Table = new DataTable();
                    SqlDataAdapter Raised_Date_Table1 = new SqlDataAdapter(Raised_Quary1);
                    Raised_Date_Table1.Fill(Raised_Data_Table);


                }


                if (Technology == "3G_CS" || Technology == "3G_PS")
                {

                    int day_counter_before = 0;
                    string Dates_before = "";
                    for (int k = 1; k <= Date_Table1.Rows.Count; k++)
                    {
                        int ind = Date_Table1.Rows.Count - k + 1;
                        DateTime Date = Convert.ToDateTime((Date_Table1.Rows[ind - 1]).ItemArray[0]);
                        if (Date < checking_date.Date && day_counter_before < checking_day_num)
                        {
                            day_counter_before++;
                            if (day_counter_before == 1)
                            {
                                Dates_before = "" + Date;
                            }
                            if (day_counter_before > 1)
                            {
                                Dates_before = Dates_before + "' or Date='" + Date;
                            }
                        }
                    }

                    if (Technology == "3G_CS")
                    {
                        Raised_Quary = @"select tbl1.Date, tbl1.Cell, tbl1.RNC, tbl1.Vendor, tbl1.Contractor, tbl1.LEVEL, tbl1.Worst, tbl1.[Effeciency Index], tbl1.Status, tbl1.QIx, tbl1.QIxP, tbl1.QIxBL, tbl1.[AVG_CS_Traffic)], 	tbl1.[Repeated Daily at 10 days (QIx<QIb)], 	tbl1.[Total Days Availability > 90], 	tbl1.[Worst(%) of RAB Establishment  Success Rate (CS)], 	tbl1.[Worst(%) of W2G IRAT/IF HO success rate], 	tbl1.[Worst(%) of Drop Call Rate], 	tbl1.[Worst(%) of Soft HO Success Rate], 	tbl1.[Worst(%) of CS RRC Connection Establishment SR (%)] from
(select [Date], [Cell], [RNC], [Vendor], [Contractor], [LEVEL], [Worst], [Effeciency Index], [Status], [QIx], [QIxP], [QIxBL], [AVG_CS_Traffic)], 	[Repeated Daily at 10 days (QIx<QIb)], 	[Total Days Availability > 90], 	[Worst(%) of RAB Establishment  Success Rate (CS)], 	[Worst(%) of W2G IRAT/IF HO success rate], 	[Worst(%) of Drop Call Rate], 	[Worst(%) of Soft HO Success Rate], 	[Worst(%) of CS RRC Connection Establishment SR (%)] from [dbo].[Contractual_WPC_3G_CS] where Contractor = '" + Region + "' and Date ='" + checking_date.Date + @"') as tbl1
left join
(select[Date], [Cell], [RNC], [Vendor], [Contractor], [LEVEL], [Worst], [Effeciency Index], [Status],  [QIxP], [QIxBL], [AVG_CS_Traffic)], 	[Repeated Daily at 10 days (QIx<QIb)], 	[Total Days Availability > 90], 	[Worst(%) of RAB Establishment  Success Rate (CS)], 	[Worst(%) of W2G IRAT/IF HO success rate], 	[Worst(%) of Drop Call Rate], 	[Worst(%) of Soft HO Success Rate], 	[Worst(%) of CS RRC Connection Establishment SR (%)]  from [dbo].[Contractual_WPC_3G_CS] where Contractor = '" + Region + "' and (Date ='" + Dates_before + @"')) as tbl2
on[tbl1].Cell =[tbl2].Cell where[tbl2].Cell is Null";
                    }
                    if (Technology == "3G_PS")
                    {
                        Raised_Quary = @"select tbl1.Date, tbl1.ElementID1, tbl1.RNC, tbl1.Vendor, tbl1.Contractor, tbl1.LEVEL, tbl1.Worst, tbl1.[Effeciency Index], tbl1.Status, tbl1.QIx, tbl1.QIxP, tbl1.QIxBL, tbl1.[Avg Payload of Cell(GB)], 	tbl1.[Repeated Daily at 10 days (QIx<QIb)], 	tbl1.[Total Days Availability > 90], 	tbl1.[Worst(%) of RAB Establishment  Success Rate (HSDPA)], 	tbl1.[Worst(%) of RAB Establishment  Success Rate (EUL)], 	tbl1.[Worst(%) of EUL MAC User Throughput (kbps)], 	tbl1.[Worst(%) of HSDPA MAC-hs User Throughput Net (kbps)], 	tbl1.[Worst(%) of RAB Drop Rate (HSDPA)], 	tbl1.[Worst(%) of RAB Drop Rate (EUL)], 	tbl1.[Worst(%) of MultiRAB Setup Success Ratio (%)], 	tbl1.[Worst(%) of PS_RRC_Setup_Success_Rate], 	tbl1.[Worst(%) of Ps_RAB_Establish_Success_Rate], 	tbl1.[Worst(%) of PS_Multi_RAB_Establish_Success_Rate], 	tbl1.[Worst(%) of Drop_Call_Rate], 	tbl1.[Worst(%) of HSDPA_Cell_Change_Succ_Rate], 	tbl1.[Worst(%) of HS_share_PAYLOAD_%], 	tbl1.[Worst(%) of HSDPA Cell Throughput (Mbps)]  from
(select [Date], [ElementID1], [RNC], [Vendor], [Contractor], [LEVEL], [Worst], [Effeciency Index], [Status], [QIx], [QIxP], [QIxBL], [Avg Payload of Cell(GB)], 	[Repeated Daily at 10 days (QIx<QIb)], 	[Total Days Availability > 90], 	[Worst(%) of RAB Establishment  Success Rate (HSDPA)], 	[Worst(%) of RAB Establishment  Success Rate (EUL)], 	[Worst(%) of EUL MAC User Throughput (kbps)], 	[Worst(%) of HSDPA MAC-hs User Throughput Net (kbps)], 	[Worst(%) of RAB Drop Rate (HSDPA)], 	[Worst(%) of RAB Drop Rate (EUL)], 	[Worst(%) of MultiRAB Setup Success Ratio (%)], 	[Worst(%) of PS_RRC_Setup_Success_Rate], 	[Worst(%) of Ps_RAB_Establish_Success_Rate], 	[Worst(%) of PS_Multi_RAB_Establish_Success_Rate], 	[Worst(%) of Drop_Call_Rate], 	[Worst(%) of HSDPA_Cell_Change_Succ_Rate], 	[Worst(%) of HS_share_PAYLOAD_%], 	[Worst(%) of HSDPA Cell Throughput (Mbps)] from [dbo].[Contractual_WPC_3G_PS] where Contractor = '" + Region + "' and Date ='" + checking_date.Date + @"') as tbl1
left join
(select[Date], [ElementID1], [RNC], [Vendor], [Contractor], [LEVEL], [Worst], [Effeciency Index], [Status],  [QIxP], [QIxBL], [Avg Payload of Cell(GB)], 	[Repeated Daily at 10 days (QIx<QIb)], 	[Total Days Availability > 90], 	[Worst(%) of RAB Establishment  Success Rate (HSDPA)], 	[Worst(%) of RAB Establishment  Success Rate (EUL)], 	[Worst(%) of EUL MAC User Throughput (kbps)], 	[Worst(%) of HSDPA MAC-hs User Throughput Net (kbps)], 	[Worst(%) of RAB Drop Rate (HSDPA)], 	[Worst(%) of RAB Drop Rate (EUL)], 	[Worst(%) of MultiRAB Setup Success Ratio (%)], 	[Worst(%) of PS_RRC_Setup_Success_Rate], 	[Worst(%) of Ps_RAB_Establish_Success_Rate], 	[Worst(%) of PS_Multi_RAB_Establish_Success_Rate], 	[Worst(%) of Drop_Call_Rate], 	[Worst(%) of HSDPA_Cell_Change_Succ_Rate], 	[Worst(%) of HS_share_PAYLOAD_%], 	[Worst(%) of HSDPA Cell Throughput (Mbps)]  from [dbo].[Contractual_WPC_3G_PS] where Contractor = '" + Region + "' and (Date ='" + Dates_before + @"')) as tbl2
on[tbl1].ElementID1 =[tbl2].ElementID1 where[tbl2].ElementID1 is Null";
                    }


                    SqlCommand Raised_Quary1 = new SqlCommand(Raised_Quary, connection);
                    Raised_Quary1.ExecuteNonQuery();
                    Raised_Data_Table = new DataTable();
                    SqlDataAdapter Raised_Date_Table1 = new SqlDataAdapter(Raised_Quary1);
                    Raised_Date_Table1.Fill(Raised_Data_Table);


                }



                if (Technology == "4G")
                {

                    int day_counter_before = 0;
                    string Dates_before = "";
                    for (int k = 1; k <= Date_Table1.Rows.Count; k++)
                    {
                        int ind = Date_Table1.Rows.Count - k + 1;
                        DateTime Date = Convert.ToDateTime((Date_Table1.Rows[ind - 1]).ItemArray[0]);
                        if (Date < checking_date.Date && day_counter_before < checking_day_num)
                        {
                            day_counter_before++;
                            if (day_counter_before == 1)
                            {
                                Dates_before = "" + Date;
                            }
                            if (day_counter_before > 1)
                            {
                                Dates_before = Dates_before + "' or Date='" + Date;
                            }
                        }
                    }


                    Raised_Quary = @"select tbl1.Date, tbl1.eNodeB, tbl1.RNC, tbl1.Vendor, tbl1.Province, tbl1.Contractor, tbl1.LEVEL, tbl1.Worst, tbl1.[Effeciency Index], tbl1.Status, tbl1.QIx, tbl1.QIxP, tbl1.QIxBL, tbl1.[Avg Payload of Cell(GB)], 	tbl1.[Repeated Daily at 10 days (QIx<QIb)], 	tbl1.[Total Days Availability > 90], 	tbl1.[Worst(%) of RRC Connection Establishment Success Rate], 	tbl1.[Worst(%) of ERAB Stablishment Success Rate (Initial)], 	tbl1.[Worst(%) of ERAB Stablishment Success Rate (Added)], 	tbl1.[Worst(%) of DL User Troughput  (Mbps)], 	tbl1.[Worst(%) of UL User Throughput (Mbps)], 	tbl1.[Worst(%) of Handover Success Rate], 	tbl1.[Worst(%) of ERAB Drop rate], 	tbl1.[Worst(%) of UE Context Drop Rate], 	tbl1.[Worst(%) of S1 Signalling Success Rate], 	tbl1.[Worst(%) of Inter Frequency Handover Execution SR (%)], 	tbl1.[Worst(%) of Intra Frequency Handover Execution SR (%)], 	tbl1.[Worst(%) of Average Ul Packet Loss Rate (%)], 	tbl1.[Worst(%) of payload per Carrier] from
(select [Date], [eNodeB], [RNC], [Vendor], [Province], [Contractor], [LEVEL], [Worst], [Effeciency Index], [Status], [QIx], [QIxP], [QIxBL], [Avg Payload of Cell(GB)], 	[Repeated Daily at 10 days (QIx<QIb)], 	[Total Days Availability > 90], 	[Worst(%) of RRC Connection Establishment Success Rate], 	[Worst(%) of ERAB Stablishment Success Rate (Initial)], 	[Worst(%) of ERAB Stablishment Success Rate (Added)], 	[Worst(%) of DL User Troughput  (Mbps)], 	[Worst(%) of UL User Throughput (Mbps)], 	[Worst(%) of Handover Success Rate], 	[Worst(%) of ERAB Drop rate], 	[Worst(%) of UE Context Drop Rate], 	[Worst(%) of S1 Signalling Success Rate], 	[Worst(%) of Inter Frequency Handover Execution SR (%)], 	[Worst(%) of Intra Frequency Handover Execution SR (%)], 	[Worst(%) of Average Ul Packet Loss Rate (%)], 	[Worst(%) of payload per Carrier] from [dbo].[Contractual_WPC_4G] where Contractor = '" + Region + "' and Date ='" + checking_date.Date + @"') as tbl1
left join
(select[Date], [eNodeB], [RNC], [Vendor], [Province], [Contractor], [LEVEL], [Worst], [Effeciency Index], [Status],  [QIxP], [QIxBL], [Avg Payload of Cell(GB)], 	[Repeated Daily at 10 days (QIx<QIb)], 	[Total Days Availability > 90], 	[Worst(%) of RRC Connection Establishment Success Rate], 	[Worst(%) of ERAB Stablishment Success Rate (Initial)], 	[Worst(%) of ERAB Stablishment Success Rate (Added)], 	[Worst(%) of DL User Troughput  (Mbps)], 	[Worst(%) of UL User Throughput (Mbps)], 	[Worst(%) of Handover Success Rate], 	[Worst(%) of ERAB Drop rate], 	[Worst(%) of UE Context Drop Rate], 	[Worst(%) of S1 Signalling Success Rate], 	[Worst(%) of Inter Frequency Handover Execution SR (%)], 	[Worst(%) of Intra Frequency Handover Execution SR (%)], 	[Worst(%) of Average Ul Packet Loss Rate (%)], 	[Worst(%) of payload per Carrier] from [dbo].[Contractual_WPC_4G] where Contractor = '" + Region + "' and (Date ='" + Dates_before + @"')) as tbl2
on[tbl1].eNodeB =[tbl2].eNodeB where[tbl2].eNodeB is Null";


                    SqlCommand Raised_Quary1 = new SqlCommand(Raised_Quary, connection);
                    Raised_Quary1.ExecuteNonQuery();
                    Raised_Data_Table = new DataTable();
                    SqlDataAdapter Raised_Date_Table1 = new SqlDataAdapter(Raised_Quary1);
                    Raised_Date_Table1.Fill(Raised_Data_Table);


                }



            }

            if (Region_Show == 1)
            {
                label15.Text = "Waiting";
                label15.BackColor = Color.Red;

                XLWorkbook wb = new XLWorkbook();
                wb.Worksheets.Add(Raised_Data_Table, "New_" + Region + "_" + Technology);
                var saveFileDialog = new SaveFileDialog
                {
                    FileName = "CWA_Export_New_" + Region + "_" + Technology + "_" + Convert.ToString(checking_date.Month) + "." + Convert.ToString(checking_date.Day) + "." + Convert.ToString(checking_date.Year),
                    Filter = "Excel files|*.xlsx",
                    Title = "Save an Excel File"
                };

                saveFileDialog.ShowDialog();

                if (!String.IsNullOrWhiteSpace(saveFileDialog.FileName))
                    wb.SaveAs(saveFileDialog.FileName);


                label15.Text = "Finished!";
                label15.BackColor = Color.Yellow;
            }
            else
            {
                MessageBox.Show("Export Part needs to select a Region!");
            }


            if (Owner_Selection_Time > Node_Selection_Time && Owner_Selection_Time > Region_Selection_Time)
            {
                selection_item = "Owner";
            }
        }

        private void worstCellReportsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form2 newFrm = new Form2(this);                                  // Form1 for Setting
            newFrm.StartPosition = FormStartPosition.CenterScreen;
            newFrm.Text = "Worst Cell Reports";
            newFrm.Size = new Size(841, 454);
            newFrm.TopMost = true;
            newFrm.Show();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            chart1.ChartAreas[0].AxisY.Maximum = Double.NaN;
            chart2.ChartAreas[0].AxisY.Maximum = Double.NaN;

            // chart1.ChartAreas[0].AxisY.Maximum = Max_Y1;
            //chart2.ChartAreas[0].AxisY.Maximum = Max_Y2;
        }


        public string[] Heather_2G = new string[100];
        public string[] Heather_3G_CS = new string[100];
        public string[] Heather_3G_PS = new string[100];
        public string[] Heather_4G = new string[100];

        private void button10_Click(object sender, EventArgs e)
        {

            if (Technology == "2G_CS")
            {

                Repeated_Data_Table = new DataTable();


                Heather_2G[0] = "Contractor";
                Heather_2G[1] = "Province";
                Heather_2G[2] = "Vendor";
                Heather_2G[3] = "BSC";
                Heather_2G[4] = "Cell";
                Heather_2G[5] = "Count";
                int indh = 6;


                Repeated_Quary = @"select distinct([Date]) from[Contractual_WPC_2G_CS]  where Contractor='" + Region + "' and Date >= '" + checking_date.Date + @"' order by date";

                // Dare Taable
                SqlCommand Repeated_Quary1 = new SqlCommand(Repeated_Quary, connection);
                Repeated_Quary1.ExecuteNonQuery();
                Repeated_Date_Table = new DataTable();
                SqlDataAdapter Repeated_Date_Table1 = new SqlDataAdapter(Repeated_Quary1);
                Repeated_Date_Table1.Fill(Repeated_Date_Table);


                DateTime first_date_check = Convert.ToDateTime(Repeated_Date_Table.Rows[0].ItemArray[0]);
                string first_date_check_str = Convert.ToString(first_date_check.Month) + "/" + Convert.ToString(first_date_check.Day) + "/" + Convert.ToString(first_date_check.Year);


                //All Data Table
                string All_Data_Table = "select distinct [Contractor], [Province], [Vendor], [BSC], [Cell], Count(Cell) as 'Count' from[Contractual_WPC_2G_CS]  where " +
                    "Contractor = '" + Region + "' and  Status = 'Contractual WPC' and Date >= '" + checking_date.Date + "' group by[Contractor], [Province], [Vendor], [BSC], [Status], [Cell] order by Province, Vendor, BSC, Cell";
                SqlCommand All_Data_Table1 = new SqlCommand(All_Data_Table, connection);
                All_Data_Table1.ExecuteNonQuery();
                SqlDataAdapter Repeated_Data_Table1 = new SqlDataAdapter(All_Data_Table1);
                Repeated_Data_Table1.Fill(Repeated_Data_Table);


                Heather_2G[indh] = ("Level at " + first_date_check_str);
                indh++;

                DateTime other_date_check = DateTime.Today;
                string other_date_check_str = "";
                for (int k = 0; k <= Repeated_Date_Table.Rows.Count - 1; k++)
                {
                    other_date_check = Convert.ToDateTime(Repeated_Date_Table.Rows[k].ItemArray[0]);
                    other_date_check_str = Convert.ToString(other_date_check.Month) + "/" + Convert.ToString(other_date_check.Day) + "/" + Convert.ToString(other_date_check.Year);


                    string other_repeated_date = "select [t1].[Contractor], [t1].[Province], [t1].[Vendor], [t1].[BSC], [t1].[Cell], [t1].Count, [t2].Level as 'Level at " + other_date_check_str +
                        "' from (select distinct [Contractor], [Province], [Vendor], [BSC], [Status], [Cell] , count([Cell]) as 'Count' from [Contractual_WPC_2G_CS] where Date >= '" + checking_date.Date + "' and Contractor = '" + Region + "'  and status='Contractual WPC' group by [Contractor], [Province], [Vendor], [BSC], [Status], [Cell]) as t1"
            + " left join (select [Province], [Vendor], [BSC], [Cell] , [Status], [Level] from[Contractual_WPC_2G_CS] where Date= '" + other_date_check.Date + "') as t2 on t1.Cell=t2.Cell and t1.BSC=t2.BSC and t1.vendor=t2.vendor and t1.Province=t2.Province and t1.status=t2.status order by Province, Vendor, BSC,  Cell";

                    SqlCommand other_repeated_date1 = new SqlCommand(other_repeated_date, connection);
                    other_repeated_date1.ExecuteNonQuery();
                    Repeated_Data_Table_Other = new DataTable();
                    SqlDataAdapter Repeated_Data_Table2 = new SqlDataAdapter(other_repeated_date1);
                    Repeated_Data_Table2.Fill(Repeated_Data_Table_Other);

                    Repeated_Data_Table.Columns.Add("Level at " + other_date_check_str, typeof(int));

                    int C1 = Repeated_Data_Table.Rows.Count - 1;

                    for (int j = 0; j <= C1; j++)
                    {
                        Repeated_Data_Table.Rows[j][7 + k - 1] = Repeated_Data_Table_Other.Rows[j].ItemArray[6];
                    }
                    Heather_2G[indh] = ("Level at " + other_date_check_str);
                    indh++;
                }
            }




            if (Technology == "3G_CS")
            {

                Repeated_Data_Table = new DataTable();


                Heather_3G_CS[0] = "Contractor";
                Heather_3G_CS[1] = "Province";
                Heather_3G_CS[2] = "Vendor";
                Heather_3G_CS[3] = "RNC";
                Heather_3G_CS[4] = "Cell";
                Heather_3G_CS[5] = "Count";
                int indh = 6;


                Repeated_Quary = @"select distinct([Date]) from[Contractual_WPC_3G_CS]  where Contractor='" + Region + "' and Date >= '" + checking_date.Date + @"' order by date";

                // Dare Taable
                SqlCommand Repeated_Quary1 = new SqlCommand(Repeated_Quary, connection);
                Repeated_Quary1.ExecuteNonQuery();
                Repeated_Date_Table = new DataTable();
                SqlDataAdapter Repeated_Date_Table1 = new SqlDataAdapter(Repeated_Quary1);
                Repeated_Date_Table1.Fill(Repeated_Date_Table);


                DateTime first_date_check = Convert.ToDateTime(Repeated_Date_Table.Rows[0].ItemArray[0]);
                string first_date_check_str = Convert.ToString(first_date_check.Month) + "/" + Convert.ToString(first_date_check.Day) + "/" + Convert.ToString(first_date_check.Year);


                //All Data Table
                string All_Data_Table = "select distinct [Contractor], [Province], [Vendor], [RNC], [Cell], Count(Cell) as 'Count' from [Contractual_WPC_3G_CS]  where " +
                    "Contractor = '" + Region + "' and Status = 'Contractual WPC' and Date >= '" + checking_date.Date + "' group by[Contractor], [Province], [Vendor], [RNC], [Status], [Cell] order by Province, Vendor, RNC, Cell";
                SqlCommand All_Data_Table1 = new SqlCommand(All_Data_Table, connection);
                All_Data_Table1.ExecuteNonQuery();
                SqlDataAdapter Repeated_Data_Table1 = new SqlDataAdapter(All_Data_Table1);
                Repeated_Data_Table1.Fill(Repeated_Data_Table);


                Heather_3G_CS[indh] = ("Level at " + first_date_check_str);
                indh++;

                DateTime other_date_check = DateTime.Today;
                string other_date_check_str = "";
                for (int k = 0; k <= Repeated_Date_Table.Rows.Count - 1; k++)
                {
                    other_date_check = Convert.ToDateTime(Repeated_Date_Table.Rows[k].ItemArray[0]);
                    other_date_check_str = Convert.ToString(other_date_check.Month) + "/" + Convert.ToString(other_date_check.Day) + "/" + Convert.ToString(other_date_check.Year);


                    string other_repeated_date = "select [t1].[Contractor], [t1].[Province], [t1].[Vendor], [t1].[RNC], [t1].[Cell], [t1].Count, [t2].Level as 'Level at " + other_date_check_str +
                        "' from (select distinct [Contractor], [Province], [Vendor], [RNC], [Status], [Cell] , count([Cell]) as 'Count' from [Contractual_WPC_3G_CS] where Date >= '" + checking_date.Date + "' and Contractor = '" + Region + "'  and status='Contractual WPC' group by [Contractor], [Province], [Vendor], [RNC], [Status], [Cell]) as t1"
            + " left join (select [Province], [Vendor], [RNC], [Cell] , [Status], [Level] from[Contractual_WPC_3G_CS] where Date= '" + other_date_check.Date + "') as t2 on t1.Cell=t2.Cell and t1.RNC=t2.RNC and t1.vendor=t2.vendor and t1.Province=t2.Province and t1.status=t2.status order by Province, Vendor, RNC,  Cell";

                    SqlCommand other_repeated_date1 = new SqlCommand(other_repeated_date, connection);
                    other_repeated_date1.ExecuteNonQuery();
                    Repeated_Data_Table_Other = new DataTable();
                    SqlDataAdapter Repeated_Data_Table2 = new SqlDataAdapter(other_repeated_date1);
                    Repeated_Data_Table2.Fill(Repeated_Data_Table_Other);

                    Repeated_Data_Table.Columns.Add("Level at " + other_date_check_str, typeof(int));

                    int C1 = Repeated_Data_Table.Rows.Count - 1;

                    for (int j = 0; j <= C1; j++)
                    {
                        Repeated_Data_Table.Rows[j][7 + k - 1] = Repeated_Data_Table_Other.Rows[j].ItemArray[6];
                    }
                    Heather_3G_CS[indh] = ("Level at " + other_date_check_str);
                    indh++;
                }

            }






            if (Technology == "3G_PS")
            {

                Repeated_Data_Table = new DataTable();



                Heather_3G_PS[0] = "Contractor";
                Heather_3G_PS[1] = "Province";
                Heather_3G_PS[2] = "Vendor";
                Heather_3G_PS[3] = "RNC";
                Heather_3G_PS[4] = "Cell";
                Heather_3G_PS[5] = "Count";
                int indh = 6;


                Repeated_Quary = @"select distinct([Date]) from[Contractual_WPC_3G_PS]  where Contractor='" + Region + "' and Date >= '" + checking_date.Date + @"' order by date";

                // Dare Taable
                SqlCommand Repeated_Quary1 = new SqlCommand(Repeated_Quary, connection);
                Repeated_Quary1.ExecuteNonQuery();
                Repeated_Date_Table = new DataTable();
                SqlDataAdapter Repeated_Date_Table1 = new SqlDataAdapter(Repeated_Quary1);
                Repeated_Date_Table1.Fill(Repeated_Date_Table);


                DateTime first_date_check = Convert.ToDateTime(Repeated_Date_Table.Rows[0].ItemArray[0]);
                string first_date_check_str = Convert.ToString(first_date_check.Month) + "/" + Convert.ToString(first_date_check.Day) + "/" + Convert.ToString(first_date_check.Year);


                //All Data Table
                string All_Data_Table = "select distinct [Contractor], [Province], [Vendor], [RNC], [ElementID1], Count(ElementID1) as 'Count' from [Contractual_WPC_3G_PS]  where " +
                    "Contractor = '" + Region + "' and Status = 'Contractual WPC' and Date >= '" + checking_date.Date + "' group by [Contractor], [Province], [Vendor], [RNC], [Status], [ElementID1] order by Vendor, Province, RNC, ElementID1";
                SqlCommand All_Data_Table1 = new SqlCommand(All_Data_Table, connection);
                All_Data_Table1.ExecuteNonQuery();
                SqlDataAdapter Repeated_Data_Table1 = new SqlDataAdapter(All_Data_Table1);
                Repeated_Data_Table1.Fill(Repeated_Data_Table);


                Heather_3G_PS[indh] = ("Level at " + first_date_check_str);
                indh++;

                DateTime other_date_check = DateTime.Today;
                string other_date_check_str = "";
                for (int k = 0; k <= Repeated_Date_Table.Rows.Count - 1; k++)
                {
                    other_date_check = Convert.ToDateTime(Repeated_Date_Table.Rows[k].ItemArray[0]);
                    other_date_check_str = Convert.ToString(other_date_check.Month) + "/" + Convert.ToString(other_date_check.Day) + "/" + Convert.ToString(other_date_check.Year);


                    string other_repeated_date = "select [t1].[Contractor], [t1].[Province], [t1].[Vendor], [t1].[RNC], [t1].[ElementID1], [t1].Count, [t2].Level as 'Level at " + other_date_check_str +
                        "' from (select distinct [Contractor], [Province], [Vendor], [RNC], [Status], [ElementID1] , count([ElementID1]) as 'Count' from [Contractual_WPC_3G_PS] where Date >= '" + checking_date.Date + "' and Contractor = '" + Region + "'  and status='Contractual WPC' group by [Contractor], [Province], [Vendor], [RNC], [Status], [ElementID1]) as t1"
            + " left join (select [Province], [Vendor], [RNC], [ElementID1] , [Status], [Level] from[Contractual_WPC_3G_PS] where Date= '" + other_date_check.Date + "') as t2 on t1.ElementID1=t2.ElementID1 and t1.RNC=t2.RNC and t1.vendor=t2.vendor and t1.Province=t2.Province and t1.status=t2.status order by Vendor, Province, RNC,  ElementID1";

                    SqlCommand other_repeated_date1 = new SqlCommand(other_repeated_date, connection);
                    other_repeated_date1.ExecuteNonQuery();
                    Repeated_Data_Table_Other = new DataTable();
                    SqlDataAdapter Repeated_Data_Table2 = new SqlDataAdapter(other_repeated_date1);
                    Repeated_Data_Table2.Fill(Repeated_Data_Table_Other);

                    Repeated_Data_Table.Columns.Add("Level at " + other_date_check_str, typeof(int));

                    int C1 = Repeated_Data_Table.Rows.Count - 1;

                    for (int j = 0; j <= C1; j++)
                    {
                        Repeated_Data_Table.Rows[j][7 + k - 1] = Repeated_Data_Table_Other.Rows[j].ItemArray[6];
                    }
                    Heather_3G_PS[indh] = ("Level at " + other_date_check_str);
                    indh++;
                }

            }



            if (Technology == "4G")
            {
                Repeated_Data_Table = new DataTable();



                Heather_4G[0] = "Contractor";
                Heather_4G[1] = "Province";
                Heather_4G[2] = "Vendor";
                Heather_4G[3] = "RNC";
                Heather_4G[4] = "eNodeB";
                Heather_4G[5] = "Count";
                int indh = 6;


                Repeated_Quary = @"select distinct([Date]) from[Contractual_WPC_4G]  where Contractor='" + Region + "' and Date >= '" + checking_date.Date + @"' order by date";

                // Dare Taable
                SqlCommand Repeated_Quary1 = new SqlCommand(Repeated_Quary, connection);
                Repeated_Quary1.ExecuteNonQuery();
                Repeated_Date_Table = new DataTable();
                SqlDataAdapter Repeated_Date_Table1 = new SqlDataAdapter(Repeated_Quary1);
                Repeated_Date_Table1.Fill(Repeated_Date_Table);


                DateTime first_date_check = Convert.ToDateTime(Repeated_Date_Table.Rows[0].ItemArray[0]);
                string first_date_check_str = Convert.ToString(first_date_check.Month) + "/" + Convert.ToString(first_date_check.Day) + "/" + Convert.ToString(first_date_check.Year);


                //All Data Table
                string All_Data_Table = "select distinct [Contractor], [Province], [Vendor], [RNC], [eNodeB], Count(eNodeB) as 'Count' from[Contractual_WPC_4G]  where " +
                    "Contractor = '" + Region + "' and Status = 'Contractual WPC' and Date >= '" + checking_date.Date + "' group by[Contractor], [Province], [Vendor], [RNC], [Status], [eNodeB] order by Vendor, Province, RNC, eNodeB";
                SqlCommand All_Data_Table1 = new SqlCommand(All_Data_Table, connection);
                All_Data_Table1.ExecuteNonQuery();
                SqlDataAdapter Repeated_Data_Table1 = new SqlDataAdapter(All_Data_Table1);
                Repeated_Data_Table1.Fill(Repeated_Data_Table);


                Heather_4G[indh] = ("Level at " + first_date_check_str);
                indh++;

                DateTime other_date_check = DateTime.Today;
                string other_date_check_str = "";
                for (int k = 0; k <= Repeated_Date_Table.Rows.Count - 1; k++)
                {
                    other_date_check = Convert.ToDateTime(Repeated_Date_Table.Rows[k].ItemArray[0]);
                    other_date_check_str = Convert.ToString(other_date_check.Month) + "/" + Convert.ToString(other_date_check.Day) + "/" + Convert.ToString(other_date_check.Year);


                    string other_repeated_date = "select [t1].[Contractor], [t1].[Province], [t1].[Vendor], [t1].[RNC], [t1].[eNodeB], [t1].Count, [t2].Level as 'Level at " + other_date_check_str +
                        "' from (select distinct [Contractor], [Province], [Vendor], [RNC], [Status], [eNodeB] , count([eNodeB]) as 'Count' from [Contractual_WPC_4G] where Date >= '" + checking_date.Date + "' and Contractor = '" + Region + "'  and status='Contractual WPC' group by [Contractor], [Province], [Vendor], [RNC], [Status], [eNodeB]) as t1"
            + " left join (select [Province], [Vendor], [RNC], [eNodeB] , [Status], [Level] from[Contractual_WPC_4G] where Date= '" + other_date_check.Date + "') as t2 on t1.eNodeB=t2.eNodeB and t1.RNC=t2.RNC and t1.vendor=t2.vendor and t1.Province=t2.Province and t1.status=t2.status order by Province, Vendor, RNC,  eNodeB";


                    SqlCommand other_repeated_date1 = new SqlCommand(other_repeated_date, connection);
                    other_repeated_date1.ExecuteNonQuery();
                    Repeated_Data_Table_Other = new DataTable();
                    SqlDataAdapter Repeated_Data_Table2 = new SqlDataAdapter(other_repeated_date1);
                    Repeated_Data_Table2.Fill(Repeated_Data_Table_Other);

                    Repeated_Data_Table.Columns.Add("Level at " + other_date_check_str, typeof(int));

                    int C1 = Repeated_Data_Table.Rows.Count - 1;

                    for (int j = 0; j <= C1; j++)
                    {
                        Repeated_Data_Table.Rows[j][7 + k - 1] = Repeated_Data_Table_Other.Rows[j].ItemArray[6];
                    }
                    Heather_4G[indh] = ("Level at " + other_date_check_str);
                    indh++;
                }

                //MessageBox.Show("Sorry. It is under devlopement");

            }


            if (Region_Show == 1)
            {
                label15.Text = "Waiting";
                label15.BackColor = Color.Red;


                XLWorkbook wb = new XLWorkbook();
                wb.Worksheets.Add(Repeated_Data_Table, "Repeated_" + Region + "_" + Technology);
                var saveFileDialog = new SaveFileDialog
                {
                    FileName = "CWA_Export_Repeated_" + Region + "_" + Technology + "_" + Convert.ToString(checking_date.Month) + "." + Convert.ToString(checking_date.Day) + "." + Convert.ToString(checking_date.Year),
                    Filter = "Excel files|*.xlsx",
                    Title = "Save an Excel File"
                };

                saveFileDialog.ShowDialog();

                if (!String.IsNullOrWhiteSpace(saveFileDialog.FileName))
                    wb.SaveAs(saveFileDialog.FileName);


                label15.Text = "Finished!";
                label15.BackColor = Color.Yellow;

            }
            else
            {
                MessageBox.Show("Export Part needs to select a Region!");
            }


        }


        void comboBox7_MouseWheel(object sender, MouseEventArgs e)
        {
            ((HandledMouseEventArgs)e).Handled = true;
        }
        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox7.MouseWheel += new MouseEventHandler(comboBox7_MouseWheel);
            Selected_KPI = comboBox7.SelectedItem.ToString();

            //  string KPI_Score_Quary_Contractual = "";
            //  string KPI_Score_Quary_NearContractual = "";
            //  string Contractual_KPI_Text = ""; 

            //  if (checkBox3.Checked==true && Technology == "2G")
            //  {

            //      if (Selected_KPI == "CSSR")
            //      {
            //          Contractual_KPI_Text = "Worst(%) of CSSR";
            //      }
            //      if (Selected_KPI == "OHSR")
            //      {
            //          Contractual_KPI_Text = "Worst(%) of OHSR";
            //      }
            //      if (Selected_KPI == "CDR")
            //      {
            //          Contractual_KPI_Text = "Worst(%) of CDR";
            //      }
            //      if (Selected_KPI == "TCH_ASFR")
            //      {
            //          Contractual_KPI_Text = "Worst(%) of TCH_Assignment_FR";
            //      }
            //      if (Selected_KPI == "RXDL")
            //      {
            //          Contractual_KPI_Text = "Worst(%) of DL Quality <=4";
            //      }
            //      if (Selected_KPI == "RXUL")
            //      {
            //          Contractual_KPI_Text = "Worst(%) of UL Quality <=4";
            //      }
            //      if (Selected_KPI == "SDCCH_CONG")
            //      {
            //          Contractual_KPI_Text = "Worst(%) of SDCCH_Congestion_Rate";
            //      }
            //      if (Selected_KPI == "SDCCH_SR")
            //      {
            //          Contractual_KPI_Text = "Worst(%) of SDCCH_Access_Success_Rate";
            //      }
            //      if (Selected_KPI == "SDCCH_DROP")
            //      {
            //          Contractual_KPI_Text = "Worst(%) of SDCCH_Drop_Rate";
            //      }
            //      if (Selected_KPI == "IHSR")
            //      {
            //          Contractual_KPI_Text = "Worst(%) of IHSR";
            //      }

            //      if (Region_Show==1)
            //      {
            //          KPI_Score_Quary_Contractual = "select [Date], 100-Avg([" + Contractual_KPI_Text + "]) from [Contractual_WPC_2G_CS] where Contractor = '" + Region + "' and Status= 'Contractual WPC' group by Date order by Date";
            //          KPI_Score_Quary_NearContractual = "select [Date], 100-Avg([" + Contractual_KPI_Text + "]) from [Contractual_WPC_2G_CS] where Contractor = '" + Region + "' and Status= 'Near Contractual WPC' group by Date order by Date";
            //     }

            //  }


            //  // Worst Cells Count in Contractual WPC
            //  SqlCommand KPI_Score_Quary_Contractual1 = new SqlCommand(KPI_Score_Quary_Contractual, connection);
            //  KPI_Score_Quary_Contractual1.ExecuteNonQuery();
            //  DataTable KPI_Score_Contractual = new DataTable();
            //  SqlDataAdapter dataAdapter_Contractual = new SqlDataAdapter(KPI_Score_Quary_Contractual1);
            //  dataAdapter_Contractual.Fill(KPI_Score_Contractual);

            //  // Worst Cells Count in Near Contractual WPC
            //  SqlCommand KPI_Score_Quary_Near_Contractual1 = new SqlCommand(KPI_Score_Quary_NearContractual, connection);
            //  KPI_Score_Quary_Near_Contractual1.ExecuteNonQuery();
            //  DataTable KPI_Score_Near_Contractual = new DataTable();
            //  SqlDataAdapter dataAdapter_Near_Contractual = new SqlDataAdapter(KPI_Score_Quary_Near_Contractual1);
            //  dataAdapter_Contractual.Fill(KPI_Score_Near_Contractual);



            //  chart1.Series.Clear();
            //  chart1.Titles.Clear();

            //  Series newSeries1 = new Series();
            //  chart1.Series.Add(newSeries1);
            //  newSeries1.IsXValueIndexed = false;
            //  chart1.Series[0].ChartType = SeriesChartType.Line;
            //  chart1.Series[0].Color = Color.Blue;
            //  chart1.Series[0].BorderWidth = 3;
            //  chart1.ChartAreas[0].AxisX.Interval = 5;
            //  chart1.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            //  chart1.Series[0].EmptyPointStyle.Color = Color.Transparent;
            //  chart1.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";
            //  chart1.Series[0].ToolTip = "#VALX [#VALY]";
            //  chart1.Series[0].IsValueShownAsLabel = false;
            // // chart1.Series[0].LegendText = "Level 1";
            //  chart1.Series[0].IsVisibleInLegend = false;
            // // newSeries1.MarkerStyle = MarkerStyle.Circle;
            ////  newSeries1.MarkerSize = 6;




            //  chart2.Series.Clear();
            //  chart2.Titles.Clear();

            //  Series newSeries2 = new Series();
            //  chart2.Series.Add(newSeries2);
            //  newSeries2.IsXValueIndexed = false;
            //  chart2.Series[0].ChartType = SeriesChartType.Line;
            //  chart2.Series[0].Color = Color.Blue;
            //  chart2.Series[0].BorderWidth = 3;
            //  chart2.ChartAreas[0].AxisX.Interval = 5;
            //  chart2.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            //  chart2.Series[0].EmptyPointStyle.Color = Color.Transparent;
            //  chart2.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";
            //  chart2.Series[0].ToolTip = "#VALX [#VALY]";
            //  chart2.Series[0].IsValueShownAsLabel = false;
            //  //chart2.Series[0].LegendText = "Level 1";
            //  chart2.Series[0].IsVisibleInLegend = false;
            ////  newSeries2.MarkerStyle = MarkerStyle.Circle;
            //  //newSeries2.MarkerSize = 6;


            //  string score_text = "";
            //  if (Region_Show==1)
            //  {
            //      score_text =  Region +"_"+Technology+ " Score (%)";
            //  }
            //  if (Province_Show == 1)
            //  {
            //      score_text = Province + "_" + Technology + " Score (%)";
            //  }
            //  if (Node_Show == 1)
            //  {
            //      score_text = Node + " Score (%)";
            //  }

            //  Title title1 = chart1.Titles.Add(score_text+ " (Contractual)");
            //  title1.Font = new System.Drawing.Font("Arial", 12, FontStyle.Regular);
            //  Title title2 = chart2.Titles.Add(score_text + " (Near Contractual)");
            //  title2.Font = new System.Drawing.Font("Arial", 12, FontStyle.Regular);






            //  chart3.Series.Clear();
            //  chart3.Titles.Clear();



            //  chart4.Series.Clear();
            //  chart4.Titles.Clear();



            //  int y1 = 0;
            //  for (int k = 1; k <= KPI_Score_Contractual.Rows.Count; k++)
            //  {
            //      DateTime dt1 = Convert.ToDateTime((KPI_Score_Contractual.Rows[k - 1]).ItemArray[0]);
            //      dt1 = dt1.AddHours(23);
            //      double Score_Contactual = 0;
            //      Score_Contactual = Convert.ToDouble((KPI_Score_Contractual.Rows[k - 1]).ItemArray[1]);
            //      chart1.Series[0].Points.AddXY(dt1, Score_Contactual);
            //      Max_Y1 = chart1.ChartAreas[0].AxisY.Maximum;
            //      if (k == 1)
            //      {
            //          Min_X = dt1;
            //      }

            //      y1 = k;
            //  }
            //  Max_X = Convert.ToDateTime((KPI_Score_Contractual.Rows[y1 - 1]).ItemArray[0]);




            //  int y2 = 0;
            //  for (int k = 1; k <= KPI_Score_Near_Contractual.Rows.Count; k++)
            //  {
            //      DateTime dt1 = Convert.ToDateTime((KPI_Score_Near_Contractual.Rows[k - 1]).ItemArray[0]);
            //      dt1 = dt1.AddHours(23);
            //      double Score_Near_Contactual = 0;
            //      Score_Near_Contactual = Convert.ToDouble((KPI_Score_Near_Contractual.Rows[k - 1]).ItemArray[1]);
            //      chart2.Series[0].Points.AddXY(dt1, Score_Near_Contactual);
            //      Max_Y2 = chart1.ChartAreas[0].AxisY.Maximum;
            //      if (k == 1)
            //      {
            //          Min_X = dt1;
            //      }

            //      y2 = k;
            //  }
            //  Max_X = Convert.ToDateTime((KPI_Score_Near_Contractual.Rows[y2 - 1]).ItemArray[0]);










        }


        void comboBox6_MouseWheel(object sender, MouseEventArgs e)
        {
            ((HandledMouseEventArgs)e).Handled = true;
        }
        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox6.MouseWheel += new MouseEventHandler(comboBox6_MouseWheel);

            Selected_Cell = comboBox6.SelectedItem.ToString();

            //DateTime date = DateTime.Now.Date.AddDays(-90).Date;
            string date1 = "2021-08-23 00:00:00.000";

            string KPI_Text_E = "";
            string KPI_Text_H = "";
            string KPI_Text_N = "";
            if (Selected_KPI == "CSSR")
            {
                KPI_Text_E = "CSSR_MCI";
                KPI_Text_H = "CSSR3";
                KPI_Text_N = "CSSR_MCI";
            }
            if (Selected_KPI == "OHSR")
            {
                KPI_Text_E = "OHSR";
                KPI_Text_H = "OHSR2";
                KPI_Text_N = "OHSR";
            }
            if (Selected_KPI == "CDR")
            {
                KPI_Text_E = "CDR(not Affected by incoming Handovers from 3G)(Eric_CELL)";
                KPI_Text_H = "CDR3";
                KPI_Text_N = "CDR(including_CS_IRAT_handovers_3G_to2G)(Nokia_SEG)";
            }
            if (Selected_KPI == "TCH_ASFR")
            {
                KPI_Text_E = "TCH_Assign_Fail_Rate(NAK)(Eric_CELL)";
                KPI_Text_H = "TCH_Assignment_FR";
                KPI_Text_N = "TCH_Assignment_FR";
            }
            if (Selected_KPI == "RXDL")
            {
                KPI_Text_E = "RxQual_DL";
                KPI_Text_H = "RX_QUALITTY_DL_NEW";
                KPI_Text_N = "RxQuality_DL";
            }
            if (Selected_KPI == "RXUL")
            {
                KPI_Text_E = "RxQual_UL";
                KPI_Text_H = "RX_QUALITTY_UL_NEW";
                KPI_Text_N = "RxQuality_UL";
            }
            if (Selected_KPI == "SDCCH_CONG")
            {
                KPI_Text_E = "SDCCH_Congestion";
                KPI_Text_H = "SDCCH_Congestion_Rate";
                KPI_Text_N = "SDCCH_Congestion_Rate";
            }
            if (Selected_KPI == "SDCCH_SR")
            {
                KPI_Text_E = "SDCCH_Access_Succ_Rate";
                KPI_Text_H = "SDCCH_Access_Success_Rate2";
                KPI_Text_N = "SDCCH_Access_Success_Rate";
            }
            if (Selected_KPI == "SDCCH_DROP")
            {
                KPI_Text_E = "SDCCH_Drop_Rate";
                KPI_Text_H = "SDCCH_Drop_Rate";
                KPI_Text_N = "SDCCH_Drop_Rate";
            }
            if (Selected_KPI == "IHSR")
            {
                KPI_Text_E = "IHSR";
                KPI_Text_H = "IHSR2";
                KPI_Text_N = "IHSR";
            }
            if (Selected_KPI == "TCH Traffic (Erlang)")
            {
                KPI_Text_E = "TCH_Traffic";
                KPI_Text_H = "TCH_Traffic";
                KPI_Text_N = "TCH_Traffic";
            }
            if (Selected_KPI == "Availability" && Technology == "2G_CS")
            {
                KPI_Text_E = "TCH_Availability";
                KPI_Text_H = "TCH_Availability";
                KPI_Text_N = "TCH_Availability";
            }




            if (Selected_KPI == "TBF_Establish")
            {
                KPI_Text_E = "TBF_Establishment_SR(UL+DL)_New(Eric_cell)";
                KPI_Text_H = "TBF_Establishment_Success_Rate(UL+DL)(%)(HU_Cell)";
                KPI_Text_N = "TBF_Establishment_SR(DL+UL)(Nokia_SEG)";
            }
            if (Selected_KPI == "TBF_Drop")
            {
                KPI_Text_E = "TBF_Drop_R_D/UL(Eric_CELL)_Harmonized";
                KPI_Text_H = "TBF_Drop(UL+DL)(HU_Cell)";
                KPI_Text_N = "TBF_Drop_Rate(UL+DL)(Nokia_SEG)";
            }
            if (Selected_KPI == "GPRS_THR")
            {
                KPI_Text_E = "DL_GPRS_user_Throughput";
                KPI_Text_H = "Average_Throughput_of_Downlink_GPRS_LLC_per_User(kbps)";
                KPI_Text_N = "LLC_throughput_GPRS_DL(kbps)(Nokia_SEG)";
            }
            if (Selected_KPI == "EGPRS_THR")
            {
                KPI_Text_E = "LLC_Throughput_EGPRS_DL";
                KPI_Text_H = "Average_Throughput_of_Downlink_EGPRS_LLC_per_User(kbps)";
                KPI_Text_N = "LLC_throughput_EDGE_DL(kbps)(Nokia_SEG)";
            }
            if (Selected_KPI == "GPRS_THR_per_TS")
            {
                KPI_Text_E = "THR_DL_GPRS_PER_TS(Eric_CELL)";
                KPI_Text_H = "THR_DL_GPRS_PER_TS(CELL_HU)";
                KPI_Text_N = "THR_DL_GPRS_PER_TS(Nokia_SEG)";
            }
            if (Selected_KPI == "EGPRS_THR_per_TS")
            {
                KPI_Text_E = "THR_DL_EGPRS_PER_TS(Eric_CELL)";
                KPI_Text_H = "THR_DL_EGPRS_PER_TS(CELL_HU)";
                KPI_Text_N = "THR_DL_EGPRS_PER_TS(Nokia_SEG)";
            }
            if (Selected_KPI == "PS Traffic (KB)")
            {
                KPI_Text_E = "Payload_Total(KB)(Eric_CELL)";
                KPI_Text_H = "Payload_Total(CELL_HU)";
                KPI_Text_N = "Payload_Data(UL+DL)(Nokia_SEG)";
            }
            if (Selected_KPI == "Availability" && Technology == "2G_PS")
            {
                KPI_Text_E = "TCH_Availability(Eric_Cell)";
                KPI_Text_H = "TCH_Availability(HU_Cell)";
                KPI_Text_N = "TCH_Availability(Nokia_SEG)";
            }





            if (Selected_KPI == "CS_RAB_Establish")
            {
                KPI_Text_E = "Cs_RAB_Establish_Success_Rate";
                KPI_Text_H = "CS_RAB_Setup_Success_Ratio";
                KPI_Text_N = "CS_RAB_Establish_Success_Rate";
            }
            if (Selected_KPI == "CS_IRAT_HO_SR")
            {
                KPI_Text_E = "IRAT_HO_Voice_Suc_Rate";
                KPI_Text_H = "CS_IRAT_HO_SR";
                KPI_Text_N = "Inter_sys_RT_Hard_HO_SR_3Gto2G(CELL_nokia)";
            }
            if (Selected_KPI == "CS_Drop_Rate")
            {
                KPI_Text_E = "CS_Drop_Call_Rate";
                KPI_Text_H = "AMR_Call_Drop_Ratio_New(Hu_CELL)";
                KPI_Text_N = "CS_Drop_Call_Rate";
            }
            if (Selected_KPI == "Soft_HO_SR")
            {
                KPI_Text_E = "Soft_HO_Suc_Rate";
                KPI_Text_H = "Softer_Handover_Success_Ratio(Hu_Cell)";
                KPI_Text_N = "Soft_HO_Success_rate_RT";
            }
            if (Selected_KPI == "CS_RRC_SR")
            {
                KPI_Text_E = "CS_RRC_Setup_Success_Rate";
                KPI_Text_H = "CS_RRC_Connection_Establishment_SR";
                KPI_Text_N = "CS_RRC_SETUP_SR_WITHOUT_REPEAT(CELL_NOKIA)";
            }
            if (Selected_KPI == "CS Traffic (Erlang)")
            {
                KPI_Text_E = "CS_Traffic";
                KPI_Text_H = "CS_Erlang";
                KPI_Text_N = "CS_Traffic";
            }
            if (Selected_KPI == "Availability" && (Technology == "3G_CS" || Technology == "3G_PS"))
            {
                KPI_Text_E = "Cell_Availability_Rate_Exclude_Blocking(UCELL_Eric)";
                KPI_Text_H = "Radio_Network_Availability_Ratio(Hu_Cell)";
                KPI_Text_N = "Cell_Availability_excluding_blocked_by_user_state";
            }
            if (Selected_KPI == "HSDPA_SR")
            {
                KPI_Text_E = "HSDPA_RAB_Setup_Succ_Rate(UCell_Eric)";
                KPI_Text_H = "HSDPA_RAB_Setup_Success_Ratio(Hu_Cell)";
                KPI_Text_N = "HSDPA_setup_success_ratio_from_user_perspective(CELL_Nokia)";
            }
            if (Selected_KPI == "HSUPA_SR")
            {
                KPI_Text_E = "HSUPA_Setup_Success_Rate(UCell_Eric)";
                KPI_Text_H = "HSUPA_RAB_Setup_Success_Ratio(Hu_Cell)";
                KPI_Text_N = "HSUPA_Setup_Success_Ratio_from_user_perspective(CELL)";
            }
            if (Selected_KPI == "UL_User_THR")
            {
                KPI_Text_E = "HSUPA_User_Throughput_MACe(Kbps)(UCell_Eric)";
                KPI_Text_H = "hsupa_uplink_throughput_in_V16(CELL_Hu)";
                KPI_Text_N = "Average_hsupa_throughput_MACe(nokia_cell)";
            }
            if (Selected_KPI == "DL_User_THR")
            {
                KPI_Text_E = "HS_USER_Throughput_NET_PQ(Mbps)(UCell_Eric)";
                KPI_Text_H = "AVERAGE_HSDPA_USER_THROUGHPUT_DC+SC(Mbit/s)(CELL_HUAWEI)";
                KPI_Text_N = "AVERAGE_HSDPA_USER_THROUGHPUT_DC+SC(Mbit/s)(Nokia_CELL)";
            }
            if (Selected_KPI == "HSDAP_Drop_Rate")
            {
                KPI_Text_E = "HSDPA_Drop_Call_Rate(UCell_Eric)";
                KPI_Text_H = "HSDPA_cdr(%)_(Hu_Cell)_new";
                KPI_Text_N = "HSDPA_Call_Drop_Rate(Nokia_Cell)";
            }
            if (Selected_KPI == "HSUPA_Drop_Rate")
            {
                KPI_Text_E = "HSUPA_Drop_Call_Rate(UCell_Eric)";
                KPI_Text_H = "HSUPA_CDR(%)_(Hu_Cell)_new";
                KPI_Text_N = "HSUPA_Call_Drop_Rate(Nokia_CELL)";
            }
            if (Selected_KPI == "MultiRAB_SR")
            {
                KPI_Text_E = "PS_Multi_RAB_Establish_Success_Rate(without_Nas)(UCELL_Eric)";
                KPI_Text_H = "CS+PS_RAB_Setup_Success_Ratio";
                KPI_Text_N = "CSAMR+PS_MRAB_stp_SR(Nokia_CELL)";
            }
            if (Selected_KPI == "PS_RRC_SR")
            {
                KPI_Text_E = "PS_RRC_Setup_Success_Rate(UCell_Eric)";
                KPI_Text_H = "PS_RRC_Connection_success_Rate_repeatless(Hu_Cell)";
                KPI_Text_N = "PS_RRCSETUP_SR";
            }
            if (Selected_KPI == "Ps_RAB_Establish")
            {
                KPI_Text_E = "Ps_RAB_Establish_Success_Rate";
                KPI_Text_H = "PS_RAB_Setup_Success_Ratio";
                KPI_Text_N = "PS_RAB_Setup_Success_Ratio";
            }
            if (Selected_KPI == "PS_MultiRAB_Establish")
            {
                KPI_Text_E = "Ps_RAB_Establish_Success_Rate(UCell_Eric)";
                KPI_Text_H = "PS_RAB_Setup_Success_Ratio(Hu_Cell)";
                KPI_Text_N = "RAB_Setup_and_Access_Complete_Ratio_for_NRT_Service_from_User_pe";
            }
            if (Selected_KPI == "PS_Drop_Rate")
            {
                KPI_Text_E = "PS_Drop_Call_Rate(UCell_Eric)";
                KPI_Text_H = "PS_Call_Drop_Ratio";
                KPI_Text_N = "Packet_Session_Drop_Ratio_NOKIA(CELL_NOKIA)";
            }
            if (Selected_KPI == "HSDPA_Cell_Change_SR")
            {
                KPI_Text_E = "HSDPA_Cell_Change_Succ_Rate(UCell_Eric)";
                KPI_Text_H = "HSDPA_Soft_HandOver_Success_Ratio";
                KPI_Text_N = "HSDPA_Cell_Change_SR(Nokia_CELL)";
            }
            if (Selected_KPI == "HS_Share_Payload")
            {
                KPI_Text_E = "HS_share_PAYLOAD_Rate(UCell_Eric)";
                KPI_Text_H = "HS_share_PAYLOAD_%";
                KPI_Text_N = "HS_SHARE_PAYLOAD(Nokia_CELL)";
            }
            if (Selected_KPI == "DL_Cell_THR")
            {
                KPI_Text_E = "HSDPA_Cell_Scheduled_Throughput(mbps)(UCell_Eric)";
                KPI_Text_H = "HSDPA_SCHEDULING_Cell_throughput(CELL_HUAWEI)";
                KPI_Text_N = "Active_HS-DSCH_cell_throughput_mbs(CELL_nokia)";
            }
            if (Selected_KPI == "PS Traffic (GB)")
            {
                KPI_Text_E = "PS_Volume(GB)(UCell_Eric)";
                KPI_Text_H = "PAYLOAD";
                KPI_Text_N = "PS_Payload_Total(HS+R99)(Nokia_CELL)_GB";
            }
            if (Selected_KPI == "Availability" && (Technology == "3G_CS" || Technology == "3G_PS"))
            {
                KPI_Text_E = "Cell_Availability_Rate_Exclude_Blocking(UCELL_Eric)";
                KPI_Text_H = "Radio_Network_Availability_Ratio(Hu_Cell)";
                if (Technology == "3G_PS")
                {
                    KPI_Text_N = "Cell_Availability_excluding_blocked_by_user_state(Nokia_UCell)";
                }
                if (Technology == "3G_CS")
                {
                    KPI_Text_N = "Cell_Availability_excluding_blocked_by_user_state";
                }
            }

            if (Selected_KPI == "RRC_Connection_SR")
            {
                KPI_Text_E = "RRC_Estab_Success_Rate(ReAtt)(EUCell_Eric)";
                KPI_Text_H = "RRC_Connection_Setup_Success_Rate_service";
                KPI_Text_N = "RRC_Connection_Setup_Success_Ratio(Nokia_LTE_CELL)";
            }
            if (Selected_KPI == "ERAB_SR_Initial")
            {
                KPI_Text_E = "Initial_ERAB_Estab_Success_Rate(eNodeB_Eric)";
                KPI_Text_H = "E-RAB_Setup_Success_Rate";
                KPI_Text_N = "Initial_E-RAB_Setup_Success_Ratio(Nokia_LTE_CELL)";
            }
            if (Selected_KPI == "ERAB_SR_Added")
            {
                KPI_Text_E = "E-RAB_Setup_SR_incl_added_New(EUCell_Eric)";
                KPI_Text_H = "E-RAB_Setup_Success_Rate(Hu_Cell)";
                KPI_Text_N = "E-RAB_Setup_SR_incl_added(Nokia_LTE_CELL)";
            }
            if (Selected_KPI == "DL_THR")
            {
                KPI_Text_E = "Average_UE_DL_Throughput(Mbps)(eNodeB_Eric)";
                KPI_Text_H = "Average_Downlink_User_Throughput(Mbit/s)";
                KPI_Text_N = "User_Throughput_DL_mbps(Nokia_LTE_CELL)";
            }
            if (Selected_KPI == "UL_THR")
            {
                KPI_Text_E = "Average_UE_UL_Throughput(Mbps)(eNodeB_Eric)";
                KPI_Text_H = "Average_UPlink_User_Throughput(Mbit/s)";
                KPI_Text_N = "User_Throughput_UL_mbps(Nokia_LTE_CELL)";
            }
            if (Selected_KPI == "HO_SR")
            {
                KPI_Text_E = "Handover_Execution_Rate(EUCell_Eric)";
                KPI_Text_H = "Intra_RAT_Handover_SR_Intra+Inter_frequency(Huawei_LTE_Cell";
                KPI_Text_N = "Intra_RAT_Handover_SR_Intra+Inter_frequency(Nokia_LTE_CELL)";
            }
            if (Selected_KPI == "ERAB_Drop_Rate")
            {
                KPI_Text_E = "E_RAB_Drop_Rate(eNodeB_Eric)";
                KPI_Text_H = "Call_Drop_Rate";
                KPI_Text_N = "E-RAB_Drop_Ratio_RAN_View(Nokia_LTE_CELL)";
            }
            if (Selected_KPI == "S1_Signalling_SR")
            {
                KPI_Text_E = "S1Signal_Estab_Success_Rate(EUCell_Eric)";
                KPI_Text_H = "S1Signal_E-RAB_Setup_SR(Hu_Cell)";
                KPI_Text_N = "S1Signal_E-RAB_Setup_SR(Nokia_LTE_CELL)";
            }
            if (Selected_KPI == "Inter_Freq_SR")
            {
                KPI_Text_E = "InterF_Handover_Execution(eNodeB_Eric)";
                KPI_Text_H = "InterF_HOOut_SR";
                KPI_Text_N = "Inter-Freq_HO_SR(Nokia_LTE_CELL)";
            }
            if (Selected_KPI == "Intra_Freq_SR")
            {
                KPI_Text_E = "IntraF_Handover_Execution(eNodeB_Eric)";
                KPI_Text_H = "IntraF_HOOut_SR";
                KPI_Text_N = "HO_Success_Ratio_intra_eNB(Nokia_LTE_CELL)";
            }
            if (Selected_KPI == "UL_Packet_Loss")
            {
                KPI_Text_E = "Average_UE_Ul_Packet_Loss_Rate(eNodeB_Eric)";
                KPI_Text_H = "Average_UL_Packet_Loss_%(Huawei_LTE_UCell)";
                KPI_Text_N = "Packet_loss_UL(Nokia_EUCELL)";
            }
            if (Selected_KPI == "Data Traffic (GB)")
            {
                KPI_Text_E = "Total_Volume(UL+DL)(GB)(eNodeB_Eric)";
                KPI_Text_H = "Total_Traffic_Volume(GB)";
                KPI_Text_N = "Total_Payload_GB(Nokia_LTE_CELL)";
            }
            if (Selected_KPI == "Availability" && Technology == "4G")
            {
                KPI_Text_E = "Cell_Availability_Rate_Exclude_Blocking(Cell_EricLTE)";
                KPI_Text_H = "Cell_Availability_Rate_Exclude_Blocking(Cell_Hu)";
                KPI_Text_N = "cell_availability_exclude_manual_blocking(Nokia_LTE_CELL)";
            }



            if (Technology == "2G_CS")
            {
                string KPI_Quary_E = "select BSC, Date, [" + KPI_Text_E + "] from [CC2_Ericsson_Cell_Daily] where Cell='" + Selected_Cell + "' and date>='" + date1 + "' order by Date";
                SqlCommand KPI_Quary_E1 = new SqlCommand(KPI_Quary_E, connection);
                KPI_Quary_E1.ExecuteNonQuery();
                KPI_Table_E1 = new DataTable();
                SqlDataAdapter KPI_Table_E = new SqlDataAdapter(KPI_Quary_E1);
                KPI_Table_E.Fill(KPI_Table_E1);


                string KPI_Quary_H = "select BSC, Date, [" + KPI_Text_H + "] from [CC2_Huawei_Cell_Daily] where Cell='" + Selected_Cell + "' and date>='" + date1 + "' order by Date";
                SqlCommand KPI_Quary_H1 = new SqlCommand(KPI_Quary_H, connection);
                KPI_Quary_H1.ExecuteNonQuery();
                KPI_Table_H1 = new DataTable();
                SqlDataAdapter KPI_Table_H = new SqlDataAdapter(KPI_Quary_H1);
                KPI_Table_H.Fill(KPI_Table_H1);

                string KPI_Quary_N = "select BSC, Date, [" + KPI_Text_N + "] from [CC2_Nokia_Cell_Daily] where SEG='" + Selected_Cell + "' and date>='" + date1 + "'  order by Date";
                SqlCommand KPI_Quary_N1 = new SqlCommand(KPI_Quary_N, connection);
                KPI_Quary_N1.ExecuteNonQuery();
                KPI_Table_N1 = new DataTable();
                SqlDataAdapter KPI_Table_N = new SqlDataAdapter(KPI_Quary_N1);
                KPI_Table_N.Fill(KPI_Table_N1);
            }


            if (Technology == "2G_PS")
            {
                string KPI_Quary_E = "select BSC, Datetime, [" + KPI_Text_E + "] from [RD2_Ericsson_Cell_Daily] where Cell='" + Selected_Cell + "' and Datetime>='" + date1 + "' order by Datetime";
                SqlCommand KPI_Quary_E1 = new SqlCommand(KPI_Quary_E, connection);
                KPI_Quary_E1.ExecuteNonQuery();
                KPI_Table_E1 = new DataTable();
                SqlDataAdapter KPI_Table_E = new SqlDataAdapter(KPI_Quary_E1);
                KPI_Table_E.Fill(KPI_Table_E1);


                string KPI_Quary_H = "select BSC, Datetime, [" + KPI_Text_H + "] from [RD2_Huawei_Cell_Daily] where Cell='" + Selected_Cell + "' and Datetime>='" + date1 + "' order by Datetime";
                SqlCommand KPI_Quary_H1 = new SqlCommand(KPI_Quary_H, connection);
                KPI_Quary_H1.ExecuteNonQuery();
                KPI_Table_H1 = new DataTable();
                SqlDataAdapter KPI_Table_H = new SqlDataAdapter(KPI_Quary_H1);
                KPI_Table_H.Fill(KPI_Table_H1);

                string KPI_Quary_N = "select BSC, Datetime, [" + KPI_Text_N + "] from [RD2_Nokia_Cell_Daily] where SEG='" + Selected_Cell + "' and Datetime>='" + date1 + "'  order by Datetime";
                SqlCommand KPI_Quary_N1 = new SqlCommand(KPI_Quary_N, connection);
                KPI_Quary_N1.ExecuteNonQuery();
                KPI_Table_N1 = new DataTable();
                SqlDataAdapter KPI_Table_N = new SqlDataAdapter(KPI_Quary_N1);
                KPI_Table_N.Fill(KPI_Table_N1);
            }


            if (Technology == "3G_CS")
            {
                string KPI_Quary_E = "select ElementID, Date, [" + KPI_Text_E + "] from [CC3_Ericsson_Cell_Daily] where ElementID1='" + Selected_Cell + "' and date>='" + date1 + "' order by Date";
                SqlCommand KPI_Quary_E1 = new SqlCommand(KPI_Quary_E, connection);
                KPI_Quary_E1.ExecuteNonQuery();
                KPI_Table_E1 = new DataTable();
                SqlDataAdapter KPI_Table_E = new SqlDataAdapter(KPI_Quary_E1);
                KPI_Table_E.Fill(KPI_Table_E1);


                string KPI_Quary_H = "select ElementID, Date, [" + KPI_Text_H + "] from [CC3_Huawei_Cell_Daily] where ElementID1='" + Selected_Cell + "' and date>='" + date1 + "' order by Date";
                SqlCommand KPI_Quary_H1 = new SqlCommand(KPI_Quary_H, connection);
                KPI_Quary_H1.ExecuteNonQuery();
                KPI_Table_H1 = new DataTable();
                SqlDataAdapter KPI_Table_H = new SqlDataAdapter(KPI_Quary_H1);
                KPI_Table_H.Fill(KPI_Table_H1);

                string KPI_Quary_N = "select ElementID, Date, [" + KPI_Text_N + "] from [CC3_Nokia_Cell_Daily] where ElementID1='" + Selected_Cell + "' and date>='" + date1 + "'  order by Date";
                SqlCommand KPI_Quary_N1 = new SqlCommand(KPI_Quary_N, connection);
                KPI_Quary_N1.ExecuteNonQuery();
                KPI_Table_N1 = new DataTable();
                SqlDataAdapter KPI_Table_N = new SqlDataAdapter(KPI_Quary_N1);
                KPI_Table_N.Fill(KPI_Table_N1);
            }

            if (Technology == "3G_PS")
            {
                string KPI_Quary_E = "select ElementID, Date, [" + KPI_Text_E + "] from [RD3_Ericsson_Cell_Daily] where ElementID1='" + Selected_Cell + "' and date>='" + date1 + "' order by Date";
                SqlCommand KPI_Quary_E1 = new SqlCommand(KPI_Quary_E, connection);
                KPI_Quary_E1.ExecuteNonQuery();
                KPI_Table_E1 = new DataTable();
                SqlDataAdapter KPI_Table_E = new SqlDataAdapter(KPI_Quary_E1);
                KPI_Table_E.Fill(KPI_Table_E1);


                string KPI_Quary_H = "select ElementID, Date, [" + KPI_Text_H + "] from [RD3_Huawei_Cell_Daily] where ElementID1='" + Selected_Cell + "' and date>='" + date1 + "' order by Date";
                SqlCommand KPI_Quary_H1 = new SqlCommand(KPI_Quary_H, connection);
                KPI_Quary_H1.ExecuteNonQuery();
                KPI_Table_H1 = new DataTable();
                SqlDataAdapter KPI_Table_H = new SqlDataAdapter(KPI_Quary_H1);
                KPI_Table_H.Fill(KPI_Table_H1);

                string KPI_Quary_N = "select ElementID, Date, [" + KPI_Text_N + "] from [RD3_Nokia_Cell_Daily] where ElementID1='" + Selected_Cell + "' and date>='" + date1 + "'  order by Date";
                SqlCommand KPI_Quary_N1 = new SqlCommand(KPI_Quary_N, connection);
                KPI_Quary_N1.ExecuteNonQuery();
                KPI_Table_N1 = new DataTable();
                SqlDataAdapter KPI_Table_N = new SqlDataAdapter(KPI_Quary_N1);
                KPI_Table_N.Fill(KPI_Table_N1);
            }
            if (Technology == "4G")
            {
                string KPI_Quary_E = "select eNodeB, Datetime, [" + KPI_Text_E + "] from [TBL_LTE_CELL_Daily_E] where eNodeB='" + Selected_Cell + "' and Datetime>='" + date1 + "' order by Datetime";
                SqlCommand KPI_Quary_E1 = new SqlCommand(KPI_Quary_E, connection);
                KPI_Quary_E1.ExecuteNonQuery();
                KPI_Table_E1 = new DataTable();
                SqlDataAdapter KPI_Table_E = new SqlDataAdapter(KPI_Quary_E1);
                KPI_Table_E.Fill(KPI_Table_E1);


                string KPI_Quary_H = "select eNodeB, Datetime, [" + KPI_Text_H + "] from [TBL_LTE_CELL_Daily_H] where eNodeB='" + Selected_Cell + "' and Datetime>='" + date1 + "' order by Datetime";
                SqlCommand KPI_Quary_H1 = new SqlCommand(KPI_Quary_H, connection);
                KPI_Quary_H1.ExecuteNonQuery();
                KPI_Table_H1 = new DataTable();
                SqlDataAdapter KPI_Table_H = new SqlDataAdapter(KPI_Quary_H1);
                KPI_Table_H.Fill(KPI_Table_H1);

                string KPI_Quary_N = "select ElementID1, Date, [" + KPI_Text_N + "] from [TBL_LTE_CELL_Daily_N] where ElementID1='" + Selected_Cell + "' and date>='" + date1 + "'  order by Date";
                SqlCommand KPI_Quary_N1 = new SqlCommand(KPI_Quary_N, connection);
                KPI_Quary_N1.ExecuteNonQuery();
                KPI_Table_N1 = new DataTable();
                SqlDataAdapter KPI_Table_N = new SqlDataAdapter(KPI_Quary_N1);
                KPI_Table_N.Fill(KPI_Table_N1);
            }



            Form4 newFrm = new Form4(this);                                  // Form1 for Setting
            newFrm.StartPosition = FormStartPosition.Manual;
            newFrm.Location = new System.Drawing.Point(285, 180);
            newFrm.Text = "KPI Chart";
            newFrm.Size = new Size(1050, 540);
            newFrm.TopMost = true;
            newFrm.Show();


        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                checkBox2.Checked = false;
                chart_type = "SeriesChartType.StackedColumn";

            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                checkBox1.Checked = false;
                chart_type = "SeriesChartType.Line";
            }
        }

        private void mAPToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form5 newFrm = new Form5(this);                                  // Form1 for Setting
            newFrm.StartPosition = FormStartPosition.CenterScreen;
            //newFrm.Location = new FormStartPosition.CenterScreen;
            newFrm.Text = "MAP";
            newFrm.Size = new Size(1360, 760);
            newFrm.TopMost = true;
            newFrm.Show();
        }

        private void kPIZeroToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form6 newFrm = new Form6(this);                                  // Form1 for Setting
            newFrm.StartPosition = FormStartPosition.CenterScreen;
            //newFrm.Location = new FormStartPosition.CenterScreen;
            newFrm.Size = new Size(1000, 660);
            newFrm.TopMost = true;
            newFrm.Show();
        }

        private void cRToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form7 newFrm = new Form7(this);                                  // Form1 for Setting
            newFrm.StartPosition = FormStartPosition.CenterScreen;
            //newFrm.Location = new FormStartPosition.CenterScreen;
            newFrm.Size = new Size(1115, 600);
            newFrm.TopMost = true;
            newFrm.Show();
        }

        private void exportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form8 newFrm = new Form8(this);                                  // Form1 for Setting
            newFrm.StartPosition = FormStartPosition.CenterScreen;
            //newFrm.Location = new FormStartPosition.CenterScreen;
            newFrm.Size = new Size(1378, 780);
            // newFrm.AutoScroll = true;
            // newFrm.AutoSize = true;
            newFrm.TopMost = true;
            newFrm.Show();
        }

        private void availabilityToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Form9 newFrm = new Form9(this);                                  // Form1 for Setting
            newFrm.StartPosition = FormStartPosition.CenterScreen;
            //newFrm.Location = new FormStartPosition.CenterScreen;
            newFrm.Size = new Size(512, 339);
            // newFrm.AutoScroll = true;
            // newFrm.AutoSize = true;
            newFrm.TopMost = true;
            newFrm.Show();
        }

        private void integrationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form11 newFrm = new Form11(this);                                  // Form1 for Setting
            newFrm.StartPosition = FormStartPosition.CenterScreen;
            //newFrm.Location = new FormStartPosition.CenterScreen;
            newFrm.Size = new Size(1000, 475);
            // newFrm.AutoScroll = true;
            // newFrm.AutoSize = true;
            newFrm.TopMost = true;
            newFrm.Show();
        }



        private void button11_Click(object sender, EventArgs e)
        {
            

            XLWorkbook wb = new XLWorkbook();
            wb.Worksheets.Add(Last_Day_List, "Data Table");
            //wb.Worksheets.Add(Availability_Table_Result, "Site Result");
            // wb.Worksheets.Add(Single_Region_Table_Contractual, "Status");
            var saveFileDialog = new SaveFileDialog
            {
                FileName = Region + "_" + Technology + "_Contractual WPC",
                Filter = "Excel files|*.xlsx",
                Title = "Save an Excel File"
            };



            saveFileDialog.ShowDialog();

            if (!String.IsNullOrWhiteSpace(saveFileDialog.FileName))
                wb.SaveAs(saveFileDialog.FileName);

            MessageBox.Show("Finished");
        }

        private void dashboardsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form3 newFrm = new Form3(this);                                  // Form1 for Setting
            //newFrm.StartPosition = FormStartPosition.CenterScreen;
            newFrm.Text = "Dashboards";
            newFrm.AutoScroll = true;
            newFrm.AutoSize = true;
            // newFrm.Size = new Size(4000, 3000);
            newFrm.TopMost = true;
            newFrm.Show();
        }

        private void button12_Click_1(object sender, EventArgs e)
        {

                XLWorkbook wb = new XLWorkbook();
              //  wb.Worksheets.Add(Single_Region_Table_Contractual, "Region Table");


            DateTime F_Date = dateTimePicker1.Value;
            DateTime E_Date = dateTimePicker2.Value;
            if (dateTimePicker1.Value.Date==DateTime.Now.Date && dateTimePicker2.Value.Date == DateTime.Now.Date)
            {
                string All_Node_List_Quary_Contractual = "";
                if (Technology == "2G_CS")
                {
                    All_Node_List_Quary_Contractual = "select [Date], [Contractor], [BSC], [Level], [Status], count(Level) as 'Worst_Count' from [Contractual_WPC_2G_CS]  where [Contractor] = '" + Region + "' and Status = 'Contractual WPC' group by[Date], [Contractor], [BSC], [Level], [Status]  order by Date";
                }
                if (Technology == "2G_PS")
                {
                    All_Node_List_Quary_Contractual = "select [Date],  [Contractor],[BSC], [Level], [Status], count(Level) as 'Worst_Count' from [Contractual_WPC_2G_PS] where [Contractor] = '" + Region + "' and Status = 'Contractual WPC' group by[Date], [Contractor], [BSC], [Level], [Status]  order by Date";
                }
                if (Technology == "3G_CS")
                {
                    All_Node_List_Quary_Contractual = "select [Date],  [Contractor],[RNC], [Level], [Status], count(Level) as 'Worst_Count' from [Contractual_WPC_3G_CS] where [Contractor] = '" + Region + "' and Status = 'Contractual WPC' group by[Date], [Contractor], [RNC], [Level], [Status]  order by Date";
                }
                if (Technology == "3G_PS")
                {
                    All_Node_List_Quary_Contractual = "select [Date],  [Contractor],[RNC], [Level], [Status], count(Level) as 'Worst_Count' from [Contractual_WPC_3G_PS] where [Contractor] = '" + Region + "' and Status = 'Contractual WPC' group by[Date], [Contractor],[RNC],  [Level], [Status]  order by Date";
                }
                if (Technology == "4G")
                {
                    All_Node_List_Quary_Contractual = "select [Date],   [Contractor],[RNC], [Level], [Status], count(Level) as 'Worst_Count' from [Contractual_WPC_4G] where [Contractor] = '" + Region + "' and Status = 'Contractual WPC' group by [Date],  [Contractor], [RNC], [Level], [Status]  order by Date";
                }

                // Worst Cells Count in Contractual WPC
                SqlCommand All_Node_List_Quary_Contractual1 = new SqlCommand(All_Node_List_Quary_Contractual, connection);
                All_Node_List_Quary_Contractual1.ExecuteNonQuery();
                DataTable All_Node_Table_Contractual = new DataTable();
                SqlDataAdapter dataAdapter_Contractual = new SqlDataAdapter(All_Node_List_Quary_Contractual1);
                dataAdapter_Contractual.Fill(All_Node_Table_Contractual);


                wb.Worksheets.Add(All_Node_Table_Contractual, "Nodes Table");
                //wb.Worksheets.Add(Availability_Table_Result, "Site Result");
                // wb.Worksheets.Add(Single_Region_Table_Contractual, "Status");
                var saveFileDialog = new SaveFileDialog
                {
                    FileName = Region + "_" + Technology + "_Contractual WPC Results",
                    Filter = "Excel files|*.xlsx",
                    Title = "Save an Excel File"
                };



                saveFileDialog.ShowDialog();

                if (!String.IsNullOrWhiteSpace(saveFileDialog.FileName))
                    wb.SaveAs(saveFileDialog.FileName);

                MessageBox.Show("Finished");
            }
            else
            {

                string All_Node_List_Quary_Contractual = "";
                if (Technology == "2G_CS")
                {
                    All_Node_List_Quary_Contractual = "select [Date], [Contractor], [BSC], [Level], [Status], count(Level) as 'Worst_Count' from [Contractual_WPC_2G_CS]  where [Contractor] = '" + Region + "' and Status = 'Contractual WPC' and Date>='"+F_Date.Date+ "' and Date<='" + E_Date.Date + "' group by[Date], [Contractor], [BSC], [Level], [Status]  order by Date";
                }
                if (Technology == "2G_PS")
                {
                    All_Node_List_Quary_Contractual = "select [Date],  [Contractor],[BSC], [Level], [Status], count(Level) as 'Worst_Count' from [Contractual_WPC_2G_PS] where [Contractor] = '" + Region + "' and Status = 'Contractual WPC' group by[Date], [Contractor], [BSC], [Level], [Status]  order by Date";
                }
                if (Technology == "3G_CS")
                {
                    All_Node_List_Quary_Contractual = "select [Date],  [Contractor],[RNC], [Level], [Status], count(Level) as 'Worst_Count' from [Contractual_WPC_3G_CS] where [Contractor] = '" + Region + "' and Status = 'Contractual WPC' group by[Date], [Contractor], [RNC], [Level], [Status]  order by Date";
                }
                if (Technology == "3G_PS")
                {
                    All_Node_List_Quary_Contractual = "select [Date],  [Contractor],[RNC], [Level], [Status], count(Level) as 'Worst_Count' from [Contractual_WPC_3G_PS] where [Contractor] = '" + Region + "' and Status = 'Contractual WPC' group by[Date], [Contractor],[RNC],  [Level], [Status]  order by Date";
                }
                if (Technology == "4G")
                {
                    All_Node_List_Quary_Contractual = "select [Date],   [Contractor],[RNC], [Level], [Status], count(Level) as 'Worst_Count' from [Contractual_WPC_4G] where [Contractor] = '" + Region + "' and Status = 'Contractual WPC' group by [Date],  [Contractor], [RNC], [Level], [Status]  order by Date";
                }

                // Worst Cells Count in Contractual WPC
                SqlCommand All_Node_List_Quary_Contractual1 = new SqlCommand(All_Node_List_Quary_Contractual, connection);
                All_Node_List_Quary_Contractual1.ExecuteNonQuery();
                DataTable All_Node_Table_Contractual = new DataTable();
                SqlDataAdapter dataAdapter_Contractual = new SqlDataAdapter(All_Node_List_Quary_Contractual1);
                dataAdapter_Contractual.Fill(All_Node_Table_Contractual);


                wb.Worksheets.Add(All_Node_Table_Contractual, "Nodes Table");
                //wb.Worksheets.Add(Availability_Table_Result, "Site Result");
                // wb.Worksheets.Add(Single_Region_Table_Contractual, "Status");
                var saveFileDialog = new SaveFileDialog
                {
                    FileName = Region + "_" + Technology + "_Contractual WPC Results",
                    Filter = "Excel files|*.xlsx",
                    Title = "Save an Excel File"
                };



                saveFileDialog.ShowDialog();

                if (!String.IsNullOrWhiteSpace(saveFileDialog.FileName))
                    wb.SaveAs(saveFileDialog.FileName);

                MessageBox.Show("Finished");


            }






        }

        private void coreToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form12 newFrm = new Form12(this);                                  // Form1 for Setting
            newFrm.StartPosition = FormStartPosition.CenterScreen;
            //newFrm.Location = new FormStartPosition.CenterScreen;
            newFrm.Size = new Size(664, 512);
            // newFrm.AutoScroll = true;
            // newFrm.AutoSize = true;
            newFrm.TopMost = true;
            newFrm.Show();
        }

        private void customerComplainToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form13 newFrm = new Form13(this);                                  // Form1 for Setting
            newFrm.StartPosition = FormStartPosition.CenterScreen;
            //newFrm.Location = new FormStartPosition.CenterScreen;
            newFrm.Size = new Size(375, 183);
            // newFrm.AutoScroll = true;
            // newFrm.AutoSize = true;
            newFrm.TopMost = true;
            newFrm.Show();

        }

        private void hourlyCheckToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form11 newFrm = new Form11(this);                                  // Form1 for Setting
            newFrm.StartPosition = FormStartPosition.CenterScreen;
            //newFrm.Location = new FormStartPosition.CenterScreen;
            newFrm.Size = new Size(345, 218);
            // newFrm.AutoScroll = true;
            // newFrm.AutoSize = true;
            newFrm.TopMost = true;
            newFrm.Show();
        }
    }




}
