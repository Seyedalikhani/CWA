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
    public partial class Dashboards : Form
    {
        public Dashboards()
        {
            InitializeComponent();
        }


        public Main form1;


        public Dashboards(Form form)
        {
            InitializeComponent();
            form1 = (Main)form;
        }

        public string Vendor = "";
        public DateTime Last_Date = DateTime.Today;

        public DataTable TH_Hourly_Data_Table = new DataTable();
        public DataTable TH_Daily_Data_Table = new DataTable();
        int Chart_X_Date_Interval = 30;

        public double Y_min1 = 0;
        public double Y_max1 = 0;

        public string Selected_Technology = "";

        public int KPI_Availability_Daily = 0;
        public int KPI_Availability_Hourly = 0;
        public int KPI_Availability_Daily_2G = 0;
        public int KPI_Availability_Daily_3G = 0;
        public int KPI_Availability_Daily_4G = 0;

        public int KPI_Accessibility_Daily = 0;
        public int KPI_Accessibility_Hourly = 0;
        public int KPI_Accessibility_CSSR_2G = 0;
        public int KPI_Accessibility_CSSR_3G = 0;
        public int KPI_Accessibility_PSSR_3G = 0;
        public int KPI_Accessibility_ERAB_4G = 0;
        public int KPI_Accessibility_RRC_4G = 0;

        public int KPI_Retainability_Daily = 0;
        public int KPI_Retainability_Hourly = 0;
        public int KPI_Retainability_CDR_2G = 0;
        public int KPI_Retainability_CDR_3G = 0;
        public int KPI_Retainability_PDR_3G = 0;
        public int KPI_Retainability_ERABDrop_4G = 0;

        public int KPI_Mobility_Daily = 0;
        public int KPI_Mobility_Hourly = 0;
        public int KPI_Mobility_IHSR_2G = 0;
        public int KPI_Mobility_OHSR_2G = 0;
        public int KPI_Mobility_IFHO_3G = 0;
        public int KPI_Mobility_CSIRAT_3G = 0;
        public int KPI_Mobility_SOFT_3G = 0;
        public int KPI_Mobility_Inter_4G = 0;
        public int KPI_Mobility_Intra_4G = 0;

        public int KPI_Voice_Daily = 0;
        public int KPI_Voice_Hourly = 0;
        public int KPI_Voice_Daily_2G = 0;
        public int KPI_Voice_Daily_3G = 0;
        public int KPI_Voice_Daily_4G = 0;

        public int KPI_Data_Daily = 0;
        public int KPI_Data_Hourly = 0;
        public int KPI_Data_Daily_2G = 0;
        public int KPI_Data_Daily_3G = 0;
        public int KPI_Data_Daily_4G = 0;

        public int KPI_THR_Daily = 0;
        public int KPI_THR_Hourly = 0;
        public int KPI_THR_Daily_User_3G = 0;
        public int KPI_THR_Daily_Cell_3G = 0;
        public int KPI_THR_Daily_User_4G = 0;
        public int KPI_THR_Daily_Cell_4G = 0;


        public double min_value_Availability_Hourly = 0;
        public double max_value_Availability_Hourly = 0;
        public double min_value_Availability_Daily = 0;
        public double max_value_Availability_Daily = 0;
        public double min_value_Accessibility_Hourly = 0;
        public double max_value_Accessibility_Hourly = 0;
        public double min_value_Accessibility_Daily = 0;
        public double max_value_Accessibility_Daily = 0;
        public double min_value_Retainability_Hourly = 0;
        public double max_value_Retainability_Hourly = 0;
        public double min_value_Retainability_Daily = 0;
        public double max_value_Retainability_Daily = 0;
        public double min_value_Mobility_Hourly = 0;
        public double max_value_Mobility_Hourly = 0;
        public double min_value_Mobility_Daily = 0;
        public double max_value_Mobility_Daily = 0;
        public double min_value_Voice_Traffic_Hourly = 0;
        public double max_value_Voice_Traffic_Hourly = 0;
        public double min_value_Voice_Traffic_Daily = 0;
        public double max_value_Voice_Traffic_Daily = 0;
        public double min_value_Data_Traffic_Hourly = 0;
        public double max_value_Data_Traffic_Hourly = 0;
        public double min_value_Data_Traffic_Daily = 0;
        public double max_value_Data_Traffic_Daily = 0;
        public double min_value_Throughput_Hourly = 0;
        public double max_value_Throughput_Hourly = 0;
        public double min_value_Throughput_Daily = 0;
        public double max_value_Throughput_Daily = 0;


        public double MIN_Value = -100000000000;
        public double MAX_Value = 100000000000;

        public string ConnectionString = "";
        public SqlConnection connection = new SqlConnection();


        public string Connection_Type = "DataBase";
        public string Server_Name = "PERFORMANCEDB";
        public string DataBase_Name = "Performance_NAK";


        public DateTime start_date = DateTime.Today.AddDays(-30);
        public DateTime end_date = DateTime.Today.AddDays(-1);


        public string Data_Type = "";
        public DataTable Provinces_Table = new DataTable();
        public DataTable Nodes_Table = new DataTable();

        public DataTable Data_Province_2G_Table = new DataTable();

        public string[,] Node_Vendor = new string[50, 2];


        public void Form3_Load(object sender, EventArgs e)
        {

            ConnectionString = @"Server=" + Server_Name + "; Database=" + DataBase_Name + "; User ID=cwpcApp;Password=cwpcApp@830625#Ahmad";
            connection = new SqlConnection(ConnectionString);
            connection.Open();

            panel1.AutoScroll = true;
            panel1.BorderStyle = BorderStyle.FixedSingle;
            panel1.Size = new Size(1300, 1500);

            button7.Location = new Point(3, 752);
            button8.Location = new Point(582, 752);
            button9.Location = new Point(3, 1121);

            chart1.Size = new Size(540, 360);
            chart1.Location = new Point(30, 10);

            chart19.Size = new Size(540, 360);
            chart19.Location = new Point(610, 10);

            chart29.Size = new Size(540, 360);
            chart29.Location = new Point(30, 380);


            chart34.Size = new Size(540, 360);
            chart34.Location = new Point(610, 380);

            chart6.Size = new Size(540, 360);
            chart6.Location = new Point(30, 750);


            chart13.Size = new Size(540, 360);
            chart13.Location = new Point(610, 750);


            chart39.Size = new Size(540, 360);
            chart39.Location = new Point(30, 1120);


           


            chart1.ChartAreas[0].CursorX.IsUserSelectionEnabled = true;
            chart1.ChartAreas[0].CursorY.IsUserSelectionEnabled = true;

            chart6.ChartAreas[0].CursorX.IsUserSelectionEnabled = true;
            chart6.ChartAreas[0].CursorY.IsUserSelectionEnabled = true;

            chart13.ChartAreas[0].CursorX.IsUserSelectionEnabled = true;
            chart13.ChartAreas[0].CursorY.IsUserSelectionEnabled = true;

            chart19.ChartAreas[0].CursorX.IsUserSelectionEnabled = true;
            chart19.ChartAreas[0].CursorY.IsUserSelectionEnabled = true;

            chart29.ChartAreas[0].CursorX.IsUserSelectionEnabled = true;
            chart29.ChartAreas[0].CursorY.IsUserSelectionEnabled = true;


            chart34.ChartAreas[0].CursorX.IsUserSelectionEnabled = true;
            chart34.ChartAreas[0].CursorY.IsUserSelectionEnabled = true;


            chart39.ChartAreas[0].CursorX.IsUserSelectionEnabled = true;
            chart39.ChartAreas[0].CursorY.IsUserSelectionEnabled = true;



        }

        private void button1_Click(object sender, EventArgs e)
        {
            label7.Text = "Waiting...";
            label7.BackColor = Color.Yellow;



            chart1.Series.Clear();
            chart1.Titles.Clear();
            Series newSeries1 = new Series();
            chart1.Series.Add(newSeries1);
            chart1.Series[0].ChartType = SeriesChartType.Area;
            chart1.Series[0].BorderWidth = 3;
            chart1.ChartAreas[0].AxisX.Interval = 1;
            chart1.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            // chart1.Series[0].EmptyPointStyle.Color = Color.Transparent;
            // chart1.Series[0].BackImageTransparentColor = Color.Transparent;
            chart1.Series[0].YAxisType = AxisType.Secondary;
            chart1.Series[0].XValueType = ChartValueType.DateTime;
            chart1.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy";
            chart1.Series[0].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy}";
            Title title1 = chart1.Titles.Add(comboBox1.SelectedItem.ToString() + "_" + comboBox3.SelectedItem.ToString());
            title1.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            chart1.Series[0].IsValueShownAsLabel = false;




            Series newSeries1_2 = new Series();
            chart1.Series.Add(newSeries1_2);
            chart1.Series[1].ChartType = SeriesChartType.Line;
            chart1.Series[1].BorderWidth = 3;
            chart1.ChartAreas[0].AxisX.Interval = 1;
            chart1.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart1.Series[1].BackSecondaryColor = Color.Transparent;
            chart1.Series[1].BackImageTransparentColor = Color.Transparent;
            chart1.Series[1].XValueType = ChartValueType.DateTime;
            chart1.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy";
            //chart1.Series[1].YAxisType = AxisType.Secondary;
            chart1.Series[1].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy}";
            chart1.Series[1].IsValueShownAsLabel = false;




       


            chart6.Series.Clear();
            chart6.Titles.Clear();
            Series newSeries7 = new Series();
            chart6.Series.Add(newSeries7);
            chart6.Series[0].ChartType = SeriesChartType.Area;
            chart6.Series[0].BorderWidth = 3;
            chart6.ChartAreas[0].AxisX.Interval = 1;
            chart6.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            // chart6.Series[0].EmptyPointStyle.Color = Color.Transparent;
            // chart6.Series[0].BackImageTransparentColor = Color.Transparent;
            chart6.Series[0].YAxisType = AxisType.Secondary;
            chart6.Series[0].XValueType = ChartValueType.DateTime;
            chart6.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy";
            chart6.Series[0].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy}";




            Series newSeries8 = new Series();
            chart6.Series.Add(newSeries8);
            chart6.Series[1].ChartType = SeriesChartType.Line;
            chart6.Series[1].BorderWidth = 3;
            chart6.ChartAreas[0].AxisX.Interval = 1;
            chart6.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart6.Series[1].BackSecondaryColor = Color.Transparent;
            chart6.Series[1].BackImageTransparentColor = Color.Transparent;
            chart6.Series[1].XValueType = ChartValueType.DateTime;
            chart6.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy";
            //chart6.Series[1].YAxisType = AxisType.Secondary;
            chart6.Series[1].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy}";
            chart6.Series[1].IsValueShownAsLabel = false;




            chart19.Series.Clear();
            chart19.Titles.Clear();
            Series newSeries19_1 = new Series();
            chart19.Series.Add(newSeries19_1);
            chart19.Series[0].ChartType = SeriesChartType.Area;
            chart19.Series[0].BorderWidth = 3;
            chart19.ChartAreas[0].AxisX.Interval = 1;
            chart19.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            // chart19.Series[0].EmptyPointStyle.Color = Color.Transparent;
            // chart19.Series[0].BackImageTransparentColor = Color.Transparent;
            chart19.Series[0].YAxisType = AxisType.Secondary;
            chart19.Series[0].XValueType = ChartValueType.DateTime;
            chart19.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy";
            chart19.Series[0].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy}";
            Title title19_1 = chart19.Titles.Add(comboBox1.SelectedItem.ToString() + "_" + "Ericsson" + "_" + comboBox3.SelectedItem.ToString());
            title19_1.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            chart19.Series[0].IsValueShownAsLabel = false;




            Series newSeries19_2 = new Series();
            chart19.Series.Add(newSeries19_2);
            chart19.Series[1].ChartType = SeriesChartType.Line;
            chart19.Series[1].BorderWidth = 3;
            chart19.ChartAreas[0].AxisX.Interval = 1;
            chart19.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart19.Series[1].BackSecondaryColor = Color.Transparent;
            chart19.Series[1].BackImageTransparentColor = Color.Transparent;
            chart19.Series[1].XValueType = ChartValueType.DateTime;
            chart19.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy";
            //chart19.Series[1].YAxisType = AxisType.Secondary;
            chart19.Series[1].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy}";
            chart19.Series[1].IsValueShownAsLabel = false;



            chart29.Series.Clear();
            chart29.Titles.Clear();
            Series newSeries29_1 = new Series();
            chart29.Series.Add(newSeries29_1);
            chart29.Series[0].ChartType = SeriesChartType.Area;
            chart29.Series[0].BorderWidth = 3;
            chart29.ChartAreas[0].AxisX.Interval = 1;
            chart29.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            // chart29.Series[0].EmptyPointStyle.Color = Color.Transparent;
            // chart29.Series[0].BackImageTransparentColor = Color.Transparent;
            chart29.Series[0].YAxisType = AxisType.Secondary;
            chart29.Series[0].XValueType = ChartValueType.DateTime;
            chart29.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy";
            chart29.Series[0].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy}";
            Title title29_1 = chart29.Titles.Add(comboBox1.SelectedItem.ToString() + "_" + "Huawei" + "_" + comboBox3.SelectedItem.ToString());
            title29_1.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            chart29.Series[0].IsValueShownAsLabel = false;




            Series newSeries29_2 = new Series();
            chart29.Series.Add(newSeries29_2);
            chart29.Series[1].ChartType = SeriesChartType.Line;
            chart29.Series[1].BorderWidth = 3;
            chart29.ChartAreas[0].AxisX.Interval = 1;
            chart29.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart29.Series[1].BackSecondaryColor = Color.Transparent;
            chart29.Series[1].BackImageTransparentColor = Color.Transparent;
            chart29.Series[1].XValueType = ChartValueType.DateTime;
            chart29.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy";
            //chart29.Series[1].YAxisType = AxisType.Secondary;
            chart29.Series[1].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy}";
            chart29.Series[1].IsValueShownAsLabel = false;



            chart34.Series.Clear();
            chart34.Titles.Clear();
            Series newSeries34_1 = new Series();
            chart34.Series.Add(newSeries34_1);
            chart34.Series[0].ChartType = SeriesChartType.Area;
            chart34.Series[0].BorderWidth = 3;
            chart34.ChartAreas[0].AxisX.Interval = 1;
            chart34.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            // chart34.Series[0].EmptyPointStyle.Color = Color.Transparent;
            // chart34.Series[0].BackImageTransparentColor = Color.Transparent;
            chart34.Series[0].YAxisType = AxisType.Secondary;
            chart34.Series[0].XValueType = ChartValueType.DateTime;
            chart34.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy";
            chart34.Series[0].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy}";
            Title title34_1 = chart34.Titles.Add(comboBox1.SelectedItem.ToString() + "_" + "Nokia" + "_" + comboBox3.SelectedItem.ToString());
            title34_1.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            chart34.Series[0].IsValueShownAsLabel = false;




            Series newSeries34_2 = new Series();
            chart34.Series.Add(newSeries34_2);
            chart34.Series[1].ChartType = SeriesChartType.Line;
            chart34.Series[1].BorderWidth = 3;
            chart34.ChartAreas[0].AxisX.Interval = 1;
            chart34.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart34.Series[1].BackSecondaryColor = Color.Transparent;
            chart34.Series[1].BackImageTransparentColor = Color.Transparent;
            chart34.Series[1].XValueType = ChartValueType.DateTime;
            chart34.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy";
            //chart34.Series[1].YAxisType = AxisType.Secondary;
            chart34.Series[1].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy}";
            chart34.Series[1].IsValueShownAsLabel = false;






            chart39.Series.Clear();
            chart39.Titles.Clear();
            Series newSeries39_1 = new Series();
            chart39.Series.Add(newSeries39_1);
            chart39.Series[0].ChartType = SeriesChartType.Line;
            chart39.Series[0].Color = Color.Blue;
            chart39.Series[0].BorderWidth = 3;
            chart39.ChartAreas[0].AxisX.Interval = 1;
            chart39.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            // chart34.Series[0].EmptyPointStyle.Color = Color.Transparent;
            // chart34.Series[0].BackImageTransparentColor = Color.Transparent;
            chart39.Series[0].YAxisType = AxisType.Secondary;
            chart39.Series[0].XValueType = ChartValueType.DateTime;
            chart39.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy";
            chart39.Series[0].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy}";
            Title title39_1 = chart39.Titles.Add(comboBox1.SelectedItem.ToString() + "_" + "Vendors" + "_" + comboBox3.SelectedItem.ToString());
            title39_1.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            chart39.Series[0].IsValueShownAsLabel = false;



            Series newSeries39_2 = new Series();
            chart39.Series.Add(newSeries39_2);
            chart39.Series[1].ChartType = SeriesChartType.Line;
            chart39.Series[1].Color = Color.Red;
            chart39.Series[1].BorderWidth = 3;
            chart39.ChartAreas[0].AxisX.Interval = 1;
            chart39.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart39.Series[1].BackSecondaryColor = Color.Transparent;
            chart39.Series[1].BackImageTransparentColor = Color.Transparent;
            chart39.Series[1].XValueType = ChartValueType.DateTime;
            chart39.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy";
            //chart39.Series[1].YAxisType = AxisType.Secondary;
            chart39.Series[1].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy}";
            chart39.Series[1].IsValueShownAsLabel = false;



            Series newSeries39_3 = new Series();
            chart39.Series.Add(newSeries39_3);
            chart39.Series[2].ChartType = SeriesChartType.Line;
            chart39.Series[2].Color = Color.Green;
            chart39.Series[2].BorderWidth = 3;
            chart39.ChartAreas[0].AxisX.Interval = 1;
            chart39.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart39.Series[2].BackSecondaryColor = Color.Transparent;
            chart39.Series[2].BackImageTransparentColor = Color.Transparent;
            chart39.Series[2].XValueType = ChartValueType.DateTime;
            chart39.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy";
            //chart39.Series[1].YAxisType = AxisType.Secondary;
            chart39.Series[2].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy}";
            chart39.Series[2].IsValueShownAsLabel = false;








           





            Series newSeries18 = new Series();
            chart13.Series.Add(newSeries18);
            chart13.Series[1].ChartType = SeriesChartType.Line;
            chart13.Series[1].BorderWidth = 3;
            chart13.ChartAreas[0].AxisX.Interval = 1;
            chart13.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart13.Series[1].EmptyPointStyle.Color = Color.Transparent;
            chart13.Series[1].XValueType = ChartValueType.DateTime;
            chart13.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy_HH:mm";
            chart13.Series[1].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy_HH:mm}";
            chart13.Series[1].IsValueShownAsLabel = false;

            Series newSeries19 = new Series();
            chart13.Series.Add(newSeries19);
            chart13.Series[2].ChartType = SeriesChartType.Line;
            chart13.Series[2].BorderWidth = 3;
            chart13.ChartAreas[0].AxisX.Interval = 1;
            chart13.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart13.Series[2].EmptyPointStyle.Color = Color.Transparent;
            chart13.Series[2].XValueType = ChartValueType.DateTime;
            chart13.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy_HH:mm";
            chart13.Series[2].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy_HH:mm}";
            chart13.Series[2].IsValueShownAsLabel = false;


            Series newSeries20 = new Series();
            chart13.Series.Add(newSeries20);
            chart13.Series[3].ChartType = SeriesChartType.Line;
            chart13.Series[3].BorderWidth = 3;
            chart13.ChartAreas[0].AxisX.Interval = 1;
            chart13.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart13.Series[3].EmptyPointStyle.Color = Color.Transparent;
            chart13.Series[3].XValueType = ChartValueType.DateTime;
            chart13.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy_HH:mm";
            chart13.Series[3].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy_HH:mm}";
            chart13.Series[3].IsValueShownAsLabel = false;













            string Province_List = "";
            string Defalut_Name = comboBox1.SelectedItem.ToString();
            Province_List = Province_List + "Province='" + Defalut_Name + "' or ";

            if (Defalut_Name == "West Azarbaijan")
            {
                Province_List = Province_List + "Province='" + "AZARGHARBI" + "' or ";
            }
            if (Defalut_Name == "East Azarbaijan")
            {
                Province_List = Province_List + "Province='" + "AZARSHARGHI" + "' or ";
            }
            if (Defalut_Name == "Chahar Mahal Va Bakhtiari")
            {
                Province_List = Province_List + "Province='" + "Charmahal" + "' or ";
            }
            if (Defalut_Name == "Esfahan")
            {
                Province_List = Province_List + "Province='" + "ISFAHAN" + "' or ";
            }
            if (Defalut_Name == "South Khorasan")
            {
                Province_List = Province_List + "Province='" + "KHORASANJONOBI" + "' or ";
            }
            if (Defalut_Name == "Khorasan Razavi")
            {
                Province_List = Province_List + "Province='" + "KHORASANRAZAVI" + "' or ";
            }
            if (Defalut_Name == "North Khorasan")
            {
                Province_List = Province_List + "Province='" + "KHORASANSHOMALI" + "' or ";
            }
            if (Defalut_Name == "Kohgiluyeh Va Boyer Ahmad")
            {
                Province_List = Province_List + "Province='" + "KOHKILOYEH" + "' or ";
            }
            if (Defalut_Name == "Sistan Va Baluchestan")
            {
                Province_List = Province_List + "Province='" + "SISTAN" + "' or ";
            }

            if (Province_List != "")
            {
                Province_List = Province_List.Substring(0, Province_List.Length - 4);
            }
            else
            {
                MessageBox.Show("Please select form Province List");
            }






            if (Province_List != "")
            {

                string Data_Quary_2G_CS = @"select [Date], [Province], [TCH_Availability] as '2G Availability' , [TCH_Traffic_24H] as '2G Traffic (Erlang)', [CDR] as '2G Voice Drop', [CSSR_MCI] as '2G CSSR', [IHSR] as 'IHSR', [OHSR] as 'OHSR',
                                      [RxQuality_DL] as 'RxQuality_DL', [RxQuality_UL] as 'RxQuality_UL', [SDCCH_Access_Success_Rate] as 'SDCCH SR', [SDCCH_Congestion_Rate] as 'SDCCH Cong', [SDCCH_Drop_Rate] as 'SDCCH Drop', [TCH_Assignment_FR] as 'TCH ASFR', [TCH_Cong_Rate] as 'TCH Cong'
                                      from [cc2_province_NEW] where  (" + Province_List + ") and ( Date>='" + start_date + "' and Date<='" + end_date + "') order by Date";

                SqlCommand Data_Quary_2G_CS_1 = new SqlCommand(Data_Quary_2G_CS, connection);
                Data_Quary_2G_CS_1.CommandTimeout = 0;
                Data_Quary_2G_CS_1.ExecuteNonQuery();
                DataTable Data_Table_2G_CS = new DataTable();
                SqlDataAdapter Data_Table_2G_CS_1 = new SqlDataAdapter(Data_Quary_2G_CS_1);
                Data_Table_2G_CS_1.Fill(Data_Table_2G_CS);






                string Data_Quary_2G_PS = @"select [Date], [Province], [Payload_Daily-GB]/1000 as '2G Payload (TB)'
                                      from [rd2_province_NEW] where  (" + Province_List + ") and ( Date>='" + start_date + "' and Date<='" + end_date + "') order by Date";

                SqlCommand Data_Quary_2G_PS_1 = new SqlCommand(Data_Quary_2G_PS, connection);
                Data_Quary_2G_PS_1.CommandTimeout = 0;
                Data_Quary_2G_PS_1.ExecuteNonQuery();
                DataTable Data_Table_2G_PS = new DataTable();
                SqlDataAdapter Data_Table_2G_PS_1 = new SqlDataAdapter(Data_Quary_2G_PS_1);
                Data_Table_2G_PS_1.Fill(Data_Table_2G_PS);




                string Data_Quary_3G_CS = @"select [Date], [Province], [Radio Availability Ratio (%)] as '3G Availability' , [CS Traffic (24H) (Erlang)] as '3G Traffic (Erlang)', [CS Call Drop Ratio (%)] as '3G Voice Drop', [CS CSSR (%)] as '3G CSSR', [InterFrequency Hardhandover Success Ratio(%)] as 'IFHO', [CS IRAT HO SR (%)] as 'CS_IRAT_HO_SR',
                                      [Soft Handover Success Ratio (%)] as 'Soft_HO_SR' from [cc3_province_NEW] where  (" + Province_List + ") and ( Date>='" + start_date + "' and Date<='" + end_date + "') order by Date";

                SqlCommand Data_Quary_3G_CS_1 = new SqlCommand(Data_Quary_3G_CS, connection);
                Data_Quary_3G_CS_1.CommandTimeout = 0;
                Data_Quary_3G_CS_1.ExecuteNonQuery();
                DataTable Data_Table_3G_CS = new DataTable();
                SqlDataAdapter Data_Table_3G_CS_1 = new SqlDataAdapter(Data_Quary_3G_CS_1);
                Data_Table_3G_CS_1.Fill(Data_Table_3G_CS);



                string Data_Quary_3G_PS = @"select [Date], [Province],  [PAYLOAD]/1000 as '3G Payload(TB)', [PS_Call_Drop_Ratio] as 'PS_Drop', [PS_CSSR] as 'PS_Setup', [RTWP] as '3G_RSSI',
                                      [AVERAGE_HSDPA_USER_THROUGHPUT_DC+SC(Mbit/s)(RNC_HUAWEI)] as '3G_User_THR', [HSDPA_SCHEDULING_Cell_throughput(CELL_HUAWEI)] as '3G_Cell_THR' from [rd3_province_NEW] where  (" + Province_List + ") and ( Date>='" + start_date + "' and Date<='" + end_date + "') order by Date";

                SqlCommand Data_Quary_3G_PS_1 = new SqlCommand(Data_Quary_3G_PS, connection);
                Data_Quary_3G_PS_1.CommandTimeout = 0;
                Data_Quary_3G_PS_1.ExecuteNonQuery();
                DataTable Data_Table_3G_PS = new DataTable();
                SqlDataAdapter Data_Table_3G_PS_1 = new SqlDataAdapter(Data_Quary_3G_PS_1);
                Data_Table_3G_PS_1.Fill(Data_Table_3G_PS);



                string Data_Quary_4G = @"select [Date], [Province], [Cell Availability(%)] as '4G Availability', [Daily Total Payload (GB)]/1000 as '4G Payload (TB)', [DL User Throughput (Mbps)] as '4G DL User THR', [UL User Throughput (Mbps)] as '4G UL User THR',  [Average CQI] as 'CQI', [E-RAB Drop Rate] as 'ERAB Drop',
                                                                   [E-RAB Setup SR Including Added Erab] as 'ERAB Setup', [Handover Execution Rate (%)] as 'HO9 SR',  [RSSI_PUCCH] as 'PUCCH RSSI', [RSSI_PUSCH] as 'PUSCH RSSI', [RRC Establishment  SR (%)] as 'RRC SR', [S1Signal Establishment SR(%)] as 'S1 SR'  from [rd4_province_NEW_V2] where  (" + Province_List + ") and ( Date>='" + start_date + "' and Date<='" + end_date + "') order by Date";

                SqlCommand Data_Quary_4G_1 = new SqlCommand(Data_Quary_4G, connection);
                Data_Quary_4G_1.CommandTimeout = 0;
                Data_Quary_4G_1.ExecuteNonQuery();
                DataTable Data_Table_4G = new DataTable();
                SqlDataAdapter Data_Table_4G_1 = new SqlDataAdapter(Data_Quary_4G_1);
                Data_Table_4G_1.Fill(Data_Table_4G);



                string Provicne_Table = "";
                string Node_Table = "";

                if (comboBox3.SelectedItem.ToString() == "CC2")
                {
                    Provicne_Table = "CC2_Province_New";
                    Node_Table = "CC2_BSC_New";
                }
                if (comboBox3.SelectedItem.ToString() == "RD2")
                {
                    Provicne_Table = "RD2_Province_New";
                    Node_Table = "RD2_BSC_New";
                }
                if (comboBox3.SelectedItem.ToString() == "CC3")
                {
                    Provicne_Table = "CC3_Province_New";
                    Node_Table = "CC3_RNC_New";
                }
                if (comboBox3.SelectedItem.ToString() == "RD3")
                {
                    Provicne_Table = "RD3_Province_New";
                    Node_Table = "RD3_RNC_New";
                }
                if (comboBox3.SelectedItem.ToString() == "RD4")
                {
                    Provicne_Table = "RD4_Province_New_V2";
                }





                //  if (comboBox2.SelectedItem.ToString() == "All")
                // {
                string Province_Quary = @"select * from " + Provicne_Table + " where  (" + Province_List + ") and ( Date>='" + start_date + "' and Date<='" + end_date + "') order by Date";

                SqlCommand Province_Quary_1 = new SqlCommand(Province_Quary, connection);
                Province_Quary_1.CommandTimeout = 0;
                Province_Quary_1.ExecuteNonQuery();
                Provinces_Table = new DataTable();
                SqlDataAdapter Province_Table_1 = new SqlDataAdapter(Province_Quary_1);
                Province_Table_1.Fill(Provinces_Table);




                double min_value_Data_Type = 100000000000;
                double max_value_Data_Type = -100000000000;
                for (int i = 0; i < Provinces_Table.Rows.Count; i++)
                {
                    DateTime dt = DateTime.Today;
                    string Data_Value = "";
                    string Traffic_Value = "";
                    if (Data_Type == "CC2")
                    {
                        dt = Convert.ToDateTime((Provinces_Table.Rows[i]).ItemArray[6]);
                        Data_Value = (Provinces_Table.Rows[i]).ItemArray[42].ToString();
                        Traffic_Value = (Provinces_Table.Rows[i]).ItemArray[7].ToString();
                    }
                    if (Data_Type == "RD2")
                    {
                        dt = Convert.ToDateTime((Provinces_Table.Rows[i]).ItemArray[6]);
                        Data_Value = (Provinces_Table.Rows[i]).ItemArray[35].ToString();
                        Traffic_Value = (Provinces_Table.Rows[i]).ItemArray[7].ToString();
                    }
                    if (Data_Type == "CC3")
                    {
                        dt = Convert.ToDateTime((Provinces_Table.Rows[i]).ItemArray[6]);
                        Data_Value = (Provinces_Table.Rows[i]).ItemArray[35].ToString();
                        Traffic_Value = (Provinces_Table.Rows[i]).ItemArray[8].ToString();
                    }
                    if (Data_Type == "RD3")
                    {
                        dt = Convert.ToDateTime((Provinces_Table.Rows[i]).ItemArray[6]);
                        Data_Value = (Provinces_Table.Rows[i]).ItemArray[105].ToString();
                        Traffic_Value = (Provinces_Table.Rows[i]).ItemArray[8].ToString();
                    }
                    if (Data_Type == "RD4")
                    {
                        dt = Convert.ToDateTime((Provinces_Table.Rows[i]).ItemArray[6]);
                        Data_Value = (Provinces_Table.Rows[i]).ItemArray[40].ToString();
                        Traffic_Value = (Provinces_Table.Rows[i]).ItemArray[8].ToString();
                    }



                    if (Data_Value != "")
                    {
                        double Data_Type_Value = Convert.ToDouble(Data_Value);
                        double Traffic = Convert.ToDouble(Traffic_Value);

                        chart1.Series[1].Points.AddXY(dt, Data_Type_Value);
                        chart1.Series[0].Points.AddXY(dt, Traffic);



                        if (Data_Type_Value > max_value_Data_Type)
                        {
                            max_value_Data_Type = Math.Round(Data_Type_Value, MidpointRounding.AwayFromZero);
                        }
                        if (Data_Type_Value < min_value_Data_Type)
                        {
                            min_value_Data_Type = Math.Round(Data_Type_Value, MidpointRounding.AwayFromZero);
                        }
                    }
                }

                chart1.ChartAreas[0].AxisY.Maximum = max_value_Data_Type + 1;
                chart1.ChartAreas[0].AxisY.Minimum = min_value_Data_Type - 1;
                chart1.Series[1].LegendText = comboBox3.SelectedItem.ToString();
                if (Data_Type == "CC2" || Data_Type == "CC3")
                {
                    chart1.Series[0].LegendText = "Traffic (Erlang)";
                }
                else
                {
                    chart1.Series[0].LegendText = "Payload (GB)";
                }


                if (comboBox3.SelectedItem.ToString() != "RD4")
                {


                    string Vendor_List_E = "Vendor='" + "E' ";
                    string Vendor_List_N = "Vendor='" + "N' ";
                    string Vendor_List_H = "Vendor='" + "H' ";

                    // Aggrigation on Vendors
                    string Province_Vendor_Quary_E = "";
                    string Province_Vendor_Quary_N = "";
                    string Province_Vendor_Quary_H = "";
                    if (Data_Type == "CC2")
                    {
                        Province_Vendor_Quary_E = @"select Province,  Region, Vendor,  Contractor, Date ,  sum(TCH_Traffic_24H) as 'TCH_Traffic_24H', sum([TCH_Traffic_24H]*[CC2 (%)])/sum([TCH_Traffic_24H]) as 'CC2 (%)'   from " + Node_Table + " where  (" + Province_List + ") and (" + Vendor_List_E + ") and ( Date>='" + start_date + "' and Date<='" + end_date + "')  and [TCH_Traffic_24H]!=0 and [TCH_Traffic_24H] is not null group by Province,  Region, Vendor,  Contractor, Date order by Date";
                        Province_Vendor_Quary_N = @"select Province,  Region, Vendor,  Contractor, Date ,  sum(TCH_Traffic_24H) as 'TCH_Traffic_24H', sum([TCH_Traffic_24H]*[CC2 (%)])/sum([TCH_Traffic_24H]) as 'CC2 (%)'   from " + Node_Table + " where  (" + Province_List + ") and (" + Vendor_List_N + ") and ( Date>='" + start_date + "' and Date<='" + end_date + "')  and [TCH_Traffic_24H]!=0 and [TCH_Traffic_24H] is not null group by Province,  Region, Vendor,  Contractor, Date order by Date";
                        Province_Vendor_Quary_H = @"select Province,  Region, Vendor,  Contractor, Date ,  sum(TCH_Traffic_24H) as 'TCH_Traffic_24H', sum([TCH_Traffic_24H]*[CC2 (%)])/sum([TCH_Traffic_24H]) as 'CC2 (%)'   from " + Node_Table + " where  (" + Province_List + ") and (" + Vendor_List_H + ") and ( Date>='" + start_date + "' and Date<='" + end_date + "')  and [TCH_Traffic_24H]!=0 and [TCH_Traffic_24H] is not null group by Province,  Region, Vendor,  Contractor, Date order by Date";
                    }
                    if (Data_Type == "RD2")
                    {
                        Province_Vendor_Quary_E = @"select Province,  Region, Vendor,  Contractor, Date ,  sum([Payload_Daily-GB]) as 'Payload_Daily-GB', sum([Payload_Daily-GB]*[RD2 (%)])/sum([Payload_Daily-GB]) as 'RD2 (%)'   from " + Node_Table + " where  (" + Province_List + ") and (" + Vendor_List_E + ") and ( Date>='" + start_date + "' and Date<='" + end_date + "')  and [Payload_Daily-GB]!=0 and [Payload_Daily-GB] is not null group by Province,  Region, Vendor,  Contractor, Date order by Date";
                        Province_Vendor_Quary_N = @"select Province,  Region, Vendor,  Contractor, Date ,  sum([Payload_Daily-GB]) as 'Payload_Daily-GB', sum([Payload_Daily-GB]*[RD2 (%)])/sum([Payload_Daily-GB]) as 'RD2 (%)'   from " + Node_Table + " where  (" + Province_List + ") and (" + Vendor_List_N + ") and ( Date>='" + start_date + "' and Date<='" + end_date + "')  and [Payload_Daily-GB]!=0 and [Payload_Daily-GB] is not null group by Province,  Region, Vendor,  Contractor, Date order by Date";
                        Province_Vendor_Quary_H = @"select Province,  Region, Vendor,  Contractor, Date ,  sum([Payload_Daily-GB]) as 'Payload_Daily-GB', sum([Payload_Daily-GB]*[RD2 (%)])/sum([Payload_Daily-GB]) as 'RD2 (%)'   from " + Node_Table + " where  (" + Province_List + ") and (" + Vendor_List_H + ") and ( Date>='" + start_date + "' and Date<='" + end_date + "')  and [Payload_Daily-GB]!=0 and [Payload_Daily-GB] is not null group by Province,  Region, Vendor,  Contractor, Date order by Date";
                    }
                    if (Data_Type == "CC3")
                    {
                        Province_Vendor_Quary_E = @"select Province,  Region, Vendor,  Contractor, Date ,  sum([CS Traffic (24H) (Erlang)]) as 'CS Traffic (24H) (Erlang)', sum([CS Traffic (24H) (Erlang)]*[CC3 (%)])/sum([CS Traffic (24H) (Erlang)]) as 'CC3 (%)'   from " + Node_Table + " where  (" + Province_List + ") and (" + Vendor_List_E + ") and ( Date>='" + start_date + "' and Date<='" + end_date + "')  and [CS Traffic (24H) (Erlang)]!=0 and [CS Traffic (24H) (Erlang)] is not null group by Province,  Region, Vendor,  Contractor, Date order by Date";
                        Province_Vendor_Quary_N = @"select Province,  Region, Vendor,  Contractor, Date ,  sum([CS Traffic (24H) (Erlang)]) as 'CS Traffic (24H) (Erlang)', sum([CS Traffic (24H) (Erlang)]*[CC3 (%)])/sum([CS Traffic (24H) (Erlang)]) as 'CC3 (%)'   from " + Node_Table + " where  (" + Province_List + ") and (" + Vendor_List_N + ") and ( Date>='" + start_date + "' and Date<='" + end_date + "')  and [CS Traffic (24H) (Erlang)]!=0 and [CS Traffic (24H) (Erlang)] is not null group by Province,  Region, Vendor,  Contractor, Date order by Date";
                        Province_Vendor_Quary_H = @"select Province,  Region, Vendor,  Contractor, Date ,  sum([CS Traffic (24H) (Erlang)]) as 'CS Traffic (24H) (Erlang)', sum([CS Traffic (24H) (Erlang)]*[CC3 (%)])/sum([CS Traffic (24H) (Erlang)]) as 'CC3 (%)'   from " + Node_Table + " where  (" + Province_List + ") and (" + Vendor_List_H + ") and ( Date>='" + start_date + "' and Date<='" + end_date + "')  and [CS Traffic (24H) (Erlang)]!=0 and [CS Traffic (24H) (Erlang)] is not null group by Province,  Region, Vendor,  Contractor, Date order by Date";
                    }

                    if (Data_Type == "RD3")
                    {
                        Province_Vendor_Quary_E = @"select Province,  Region, Vendor,  Contractor, Date ,  sum([PAYLOAD]) as 'PAYLOAD', sum([PAYLOAD]*[RD3 (%)])/sum([PAYLOAD]) as 'RD3 (%)'   from " + Node_Table + " where  (" + Province_List + ") and (" + Vendor_List_E + ") and ( Date>='" + start_date + "' and Date<='" + end_date + "')  and [PAYLOAD]!= 0 and [PAYLOAD] is not null group by Province,  Region, Vendor,  Contractor, Date order by Date";
                        Province_Vendor_Quary_N = @"select Province,  Region, Vendor,  Contractor, Date ,  sum([PAYLOAD]) as 'PAYLOAD', sum([PAYLOAD]*[RD3 (%)])/sum([PAYLOAD]) as 'RD3 (%)'   from " + Node_Table + " where  (" + Province_List + ") and (" + Vendor_List_N + ") and ( Date>='" + start_date + "' and Date<='" + end_date + "')  and [PAYLOAD]!= 0 and [PAYLOAD] is not null group by Province,  Region, Vendor,  Contractor, Date order by Date";
                        Province_Vendor_Quary_H = @"select Province,  Region, Vendor,  Contractor, Date ,  sum([PAYLOAD]) as 'PAYLOAD', sum([PAYLOAD]*[RD3 (%)])/sum([PAYLOAD]) as 'RD3 (%)'   from " + Node_Table + " where  (" + Province_List + ") and (" + Vendor_List_H + ") and ( Date>='" + start_date + "' and Date<='" + end_date + "')  and [PAYLOAD]!= 0 and [PAYLOAD] is not null group by Province,  Region, Vendor,  Contractor, Date order by Date";
                    }


                    SqlCommand Province_Vendor_Quary_E1 = new SqlCommand(Province_Vendor_Quary_E, connection);
                    Province_Vendor_Quary_E1.CommandTimeout = 0;
                    Province_Vendor_Quary_E1.ExecuteNonQuery();
                    DataTable Provinces_Table_E = new DataTable();
                    SqlDataAdapter Provinces_Table_E_1 = new SqlDataAdapter(Province_Vendor_Quary_E1);
                    Provinces_Table_E_1.Fill(Provinces_Table_E);


                    SqlCommand Province_Vendor_Quary_H1 = new SqlCommand(Province_Vendor_Quary_H, connection);
                    Province_Vendor_Quary_H1.CommandTimeout = 0;
                    Province_Vendor_Quary_H1.ExecuteNonQuery();
                    DataTable Provinces_Table_H = new DataTable();
                    SqlDataAdapter Provinces_Table_H_1 = new SqlDataAdapter(Province_Vendor_Quary_H1);
                    Provinces_Table_H_1.Fill(Provinces_Table_H);


                    SqlCommand Province_Vendor_Quary_N1 = new SqlCommand(Province_Vendor_Quary_N, connection);
                    Province_Vendor_Quary_N1.CommandTimeout = 0;
                    Province_Vendor_Quary_N1.ExecuteNonQuery();
                    DataTable Provinces_Table_N = new DataTable();
                    SqlDataAdapter Provinces_Table_N_1 = new SqlDataAdapter(Province_Vendor_Quary_N1);
                    Provinces_Table_N_1.Fill(Provinces_Table_N);





                    double min_value_Data_Type_E = 100000000000;
                    double max_value_Data_Type_E = -100000000000;
                    for (int i = 0; i < Provinces_Table_E.Rows.Count; i++)
                    {
                        DateTime dt = DateTime.Today;
                        string Data_Value = "";
                        string Traffic_Value = "";
                        if (Data_Type == "CC2")
                        {
                            dt = Convert.ToDateTime((Provinces_Table_E.Rows[i]).ItemArray[4]);
                            Data_Value = (Provinces_Table_E.Rows[i]).ItemArray[6].ToString();
                            Traffic_Value = (Provinces_Table_E.Rows[i]).ItemArray[5].ToString();
                        }
                        if (Data_Type == "RD2")
                        {
                            dt = Convert.ToDateTime((Provinces_Table_E.Rows[i]).ItemArray[4]);
                            Data_Value = (Provinces_Table_E.Rows[i]).ItemArray[6].ToString();
                            Traffic_Value = (Provinces_Table_E.Rows[i]).ItemArray[5].ToString();
                        }
                        if (Data_Type == "CC3")
                        {
                            dt = Convert.ToDateTime((Provinces_Table_E.Rows[i]).ItemArray[4]);
                            Data_Value = (Provinces_Table_E.Rows[i]).ItemArray[6].ToString();
                            Traffic_Value = (Provinces_Table_E.Rows[i]).ItemArray[5].ToString();
                        }
                        if (Data_Type == "RD3")
                        {
                            dt = Convert.ToDateTime((Provinces_Table_E.Rows[i]).ItemArray[4]);
                            Data_Value = (Provinces_Table_E.Rows[i]).ItemArray[6].ToString();
                            Traffic_Value = (Provinces_Table_E.Rows[i]).ItemArray[5].ToString();
                        }




                        if (Data_Value != "")
                        {
                            double Data_Type_Value = Convert.ToDouble(Data_Value);
                            double Traffic = Convert.ToDouble(Traffic_Value);

                            chart19.Series[1].Points.AddXY(dt, Data_Type_Value);
                            chart19.Series[0].Points.AddXY(dt, Traffic);
                            chart39.Series[0].Points.AddXY(dt, Data_Type_Value);

                            if (Data_Type_Value > max_value_Data_Type_E)
                            {
                                max_value_Data_Type_E = Math.Round(Data_Type_Value, MidpointRounding.AwayFromZero);
                            }
                            if (Data_Type_Value < min_value_Data_Type_E)
                            {
                                min_value_Data_Type_E = Math.Round(Data_Type_Value, MidpointRounding.AwayFromZero);
                            }
                        }
                    }

                    chart19.ChartAreas[0].AxisY.Maximum = max_value_Data_Type_E + 1;
                    chart19.ChartAreas[0].AxisY.Minimum = min_value_Data_Type_E - 1;
                    chart19.Series[1].LegendText = comboBox3.SelectedItem.ToString();
                    chart39.Series[0].LegendText = "Ericsson";
                    if (Data_Type == "CC2" || Data_Type == "CC3")
                    {
                        chart19.Series[0].LegendText = "Traffic (Erlang)";
                    }
                    else
                    {
                        chart19.Series[0].LegendText = "Payload (GB)";
                    }





                    double min_value_Data_Type_H = 100000000000;
                    double max_value_Data_Type_H = -100000000000;
                    for (int i = 0; i < Provinces_Table_H.Rows.Count; i++)
                    {
                        DateTime dt = DateTime.Today;
                        string Data_Value = "";
                        string Traffic_Value = "";
                        if (Data_Type == "CC2")
                        {
                            dt = Convert.ToDateTime((Provinces_Table_H.Rows[i]).ItemArray[4]);
                            Data_Value = (Provinces_Table_H.Rows[i]).ItemArray[6].ToString();
                            Traffic_Value = (Provinces_Table_H.Rows[i]).ItemArray[5].ToString();
                        }
                        if (Data_Type == "RD2")
                        {
                            dt = Convert.ToDateTime((Provinces_Table_H.Rows[i]).ItemArray[4]);
                            Data_Value = (Provinces_Table_H.Rows[i]).ItemArray[6].ToString();
                            Traffic_Value = (Provinces_Table_H.Rows[i]).ItemArray[5].ToString();
                        }
                        if (Data_Type == "CC3")
                        {
                            dt = Convert.ToDateTime((Provinces_Table_H.Rows[i]).ItemArray[4]);
                            Data_Value = (Provinces_Table_H.Rows[i]).ItemArray[6].ToString();
                            Traffic_Value = (Provinces_Table_H.Rows[i]).ItemArray[5].ToString();
                        }
                        if (Data_Type == "RD3")
                        {
                            dt = Convert.ToDateTime((Provinces_Table_H.Rows[i]).ItemArray[4]);
                            Data_Value = (Provinces_Table_H.Rows[i]).ItemArray[6].ToString();
                            Traffic_Value = (Provinces_Table_H.Rows[i]).ItemArray[5].ToString();
                        }




                        if (Data_Value != "")
                        {
                            double Data_Type_Value = Convert.ToDouble(Data_Value);
                            double Traffic = Convert.ToDouble(Traffic_Value);

                            chart29.Series[1].Points.AddXY(dt, Data_Type_Value);
                            chart29.Series[0].Points.AddXY(dt, Traffic);
                            chart39.Series[1].Points.AddXY(dt, Data_Type_Value);


                            if (Data_Type_Value > max_value_Data_Type_H)
                            {
                                max_value_Data_Type_H = Math.Round(Data_Type_Value, MidpointRounding.AwayFromZero);
                            }
                            if (Data_Type_Value < min_value_Data_Type_H)
                            {
                                min_value_Data_Type_H = Math.Round(Data_Type_Value, MidpointRounding.AwayFromZero);
                            }
                        }
                    }

                    chart29.ChartAreas[0].AxisY.Maximum = max_value_Data_Type_H + 1;
                    chart29.ChartAreas[0].AxisY.Minimum = min_value_Data_Type_H - 1;
                    chart29.Series[1].LegendText = comboBox3.SelectedItem.ToString();
                    chart39.Series[1].LegendText = "Huawei";
                    if (Data_Type == "CC2" || Data_Type == "CC3")
                    {
                        chart29.Series[0].LegendText = "Traffic (Erlang)";
                    }
                    else
                    {
                        chart29.Series[0].LegendText = "Payload (GB)";
                    }







                    double min_value_Data_Type_N = 100000000000;
                    double max_value_Data_Type_N = -100000000000;
                    for (int i = 0; i < Provinces_Table_N.Rows.Count; i++)
                    {
                        DateTime dt = DateTime.Today;
                        string Data_Value = "";
                        string Traffic_Value = "";
                        if (Data_Type == "CC2")
                        {
                            dt = Convert.ToDateTime((Provinces_Table_N.Rows[i]).ItemArray[4]);
                            Data_Value = (Provinces_Table_N.Rows[i]).ItemArray[6].ToString();
                            Traffic_Value = (Provinces_Table_N.Rows[i]).ItemArray[5].ToString();
                        }
                        if (Data_Type == "RD2")
                        {
                            dt = Convert.ToDateTime((Provinces_Table_N.Rows[i]).ItemArray[4]);
                            Data_Value = (Provinces_Table_N.Rows[i]).ItemArray[6].ToString();
                            Traffic_Value = (Provinces_Table_N.Rows[i]).ItemArray[5].ToString();
                        }
                        if (Data_Type == "CC3")
                        {
                            dt = Convert.ToDateTime((Provinces_Table_N.Rows[i]).ItemArray[4]);
                            Data_Value = (Provinces_Table_N.Rows[i]).ItemArray[6].ToString();
                            Traffic_Value = (Provinces_Table_N.Rows[i]).ItemArray[5].ToString();
                        }
                        if (Data_Type == "RD3")
                        {
                            dt = Convert.ToDateTime((Provinces_Table_N.Rows[i]).ItemArray[4]);
                            Data_Value = (Provinces_Table_N.Rows[i]).ItemArray[6].ToString();
                            Traffic_Value = (Provinces_Table_N.Rows[i]).ItemArray[5].ToString();
                        }




                        if (Data_Value != "")
                        {
                            double Data_Type_Value = Convert.ToDouble(Data_Value);
                            double Traffic = Convert.ToDouble(Traffic_Value);

                            chart34.Series[1].Points.AddXY(dt, Data_Type_Value);
                            chart34.Series[0].Points.AddXY(dt, Traffic);
                            chart39.Series[2].Points.AddXY(dt, Data_Type_Value);

                            if (Data_Type_Value > max_value_Data_Type_N)
                            {
                                max_value_Data_Type_N = Math.Round(Data_Type_Value, MidpointRounding.AwayFromZero);
                            }
                            if (Data_Type_Value < min_value_Data_Type_N)
                            {
                                min_value_Data_Type_N = Math.Round(Data_Type_Value, MidpointRounding.AwayFromZero);
                            }
                        }
                    }

                    chart34.ChartAreas[0].AxisY.Maximum = max_value_Data_Type_N + 1;
                    chart34.ChartAreas[0].AxisY.Minimum = min_value_Data_Type_N - 1;
                    chart34.Series[1].LegendText = comboBox3.SelectedItem.ToString();
                    chart39.Series[2].LegendText = "Nokia";
                    if (Data_Type == "CC2" || Data_Type == "CC3")
                    {
                        chart34.Series[0].LegendText = "Traffic (Erlang)";
                    }
                    else
                    {
                        chart34.Series[0].LegendText = "Payload (GB)";
                    }




                }



                if (comboBox3.SelectedItem.ToString() != "RD4")
                {
                    string Node_Quary = @"select * from " + Node_Table + " where  (" + Province_List + ") and ( Date>='" + start_date + "' and Date<='" + end_date + "') order by Date";

                    SqlCommand Node_Quary_1 = new SqlCommand(Node_Quary, connection);
                    Node_Quary_1.CommandTimeout = 0;
                    Node_Quary_1.ExecuteNonQuery();
                    Nodes_Table = new DataTable();
                    SqlDataAdapter Node_Table_1 = new SqlDataAdapter(Node_Quary_1);
                    Node_Table_1.Fill(Nodes_Table);



                    comboBox4.Items.Clear();
                    Node_Vendor = new string[50, 2];
                    int node_ind = 0;
                    for (int k = 0; k < Nodes_Table.Rows.Count; k++)
                    {
                        string Node = (Nodes_Table.Rows[k]).ItemArray[0].ToString();
                        if (!comboBox4.Items.Contains(Node))
                        {

                            Node_Vendor[node_ind, 0] = Node;
                            if (Data_Type == "CC2" || Data_Type == "RD2")
                            {
                                Node_Vendor[node_ind, 1] = (Nodes_Table.Rows[k]).ItemArray[5].ToString();
                            }
                            if (Data_Type == "CC3" || Data_Type == "RD3")
                            {
                                Node_Vendor[node_ind, 1] = (Nodes_Table.Rows[k]).ItemArray[6].ToString();
                            }

                            node_ind++;

                            comboBox4.Items.Add(Node);
                        }

                    }





                }

    
     


                // Setting of Intervals
                // **************************************************************
                double difference_day = (end_date - start_date).TotalDays;
                double day_interval = Math.Round(difference_day / 20);
                if (day_interval == 0)
                {
                    day_interval = 1;
                }
                chart1.ChartAreas[0].AxisX.Interval = day_interval;
                chart6.ChartAreas[0].AxisX.Interval = day_interval;
                chart13.ChartAreas[0].AxisX.Interval = day_interval;
                chart19.ChartAreas[0].AxisX.Interval = day_interval;
                chart29.ChartAreas[0].AxisX.Interval = day_interval;
                chart34.ChartAreas[0].AxisX.Interval = day_interval;
                chart39.ChartAreas[0].AxisX.Interval = day_interval;
                // **************************************************************



                label7.Text = "Finished";
                label7.BackColor = Color.Green;



            }




        }



        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            start_date = dateTimePicker1.Value.Date;
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            end_date = dateTimePicker2.Value.Date;

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            Data_Type = comboBox3.SelectedItem.ToString();
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {

            chart6.Series.Clear();
            chart6.Titles.Clear();
            Series newSeries7 = new Series();
            chart6.Series.Add(newSeries7);
            chart6.Series[0].ChartType = SeriesChartType.Area;
            chart6.Series[0].BorderWidth = 3;
            chart6.ChartAreas[0].AxisX.Interval = 1;
            chart6.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            // chart6.Series[0].EmptyPointStyle.Color = Color.Transparent;
            // chart6.Series[0].BackImageTransparentColor = Color.Transparent;
            chart6.Series[0].YAxisType = AxisType.Secondary;
            chart6.Series[0].XValueType = ChartValueType.DateTime;
            chart6.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy";
            chart6.Series[0].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy}";
            Title title2_1 = chart6.Titles.Add(comboBox1.SelectedItem.ToString() + "_" + comboBox3.SelectedItem.ToString() + "_" + comboBox4.SelectedItem.ToString());
            title2_1.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            chart6.Series[0].IsValueShownAsLabel = false;



            Series newSeries8 = new Series();
            chart6.Series.Add(newSeries8);
            chart6.Series[1].ChartType = SeriesChartType.Line;
            chart6.Series[1].BorderWidth = 3;
            chart6.ChartAreas[0].AxisX.Interval = 1;
            chart6.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart6.Series[1].BackSecondaryColor = Color.Transparent;
            chart6.Series[1].BackImageTransparentColor = Color.Transparent;
            chart6.Series[1].XValueType = ChartValueType.DateTime;
            chart6.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy";
            //chart6.Series[1].YAxisType = AxisType.Secondary;
            chart6.Series[1].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy}";
            chart6.Series[1].IsValueShownAsLabel = false;






            chart13.Series.Clear();
            chart13.Titles.Clear();
            Series newSeries13_1 = new Series();
            chart13.Series.Add(newSeries13_1);
            chart13.Series[0].ChartType = SeriesChartType.Line;
            chart13.Series[0].BorderWidth = 3;
            chart13.Series[0].Color = Color.Red;
            chart13.ChartAreas[0].AxisX.Interval = 1;
            chart13.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart13.Series[0].EmptyPointStyle.Color = Color.Transparent;
            chart13.Series[0].XValueType = ChartValueType.DateTime;
            chart13.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy";
            chart13.Series[0].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy}";
            Title title5 = chart13.Titles.Add(comboBox3.SelectedItem.ToString() + " of All Nodes");
            title5.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            chart13.Series[0].IsValueShownAsLabel = false;


            Series newSeries13_2 = new Series();
            chart13.Series.Add(newSeries13_2);
            chart13.Series[1].ChartType = SeriesChartType.Line;
            chart13.Series[1].BorderWidth = 1;
            chart13.Series[1].Color = Color.Gray;
            chart13.ChartAreas[0].AxisX.Interval = 1;
            chart13.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart13.Series[1].EmptyPointStyle.Color = Color.Transparent;
            chart13.Series[1].XValueType = ChartValueType.DateTime;
            chart13.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy";
            chart13.Series[1].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy}";
            chart13.Series[1].IsValueShownAsLabel = false;


            Series newSeries13_3 = new Series();
            chart13.Series.Add(newSeries13_3);
            chart13.Series[2].ChartType = SeriesChartType.Line;
            chart13.Series[2].BorderWidth = 1;
            chart13.Series[2].Color = Color.Gray;
            chart13.ChartAreas[0].AxisX.Interval = 1;
            chart13.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart13.Series[2].EmptyPointStyle.Color = Color.Transparent;
            chart13.Series[2].XValueType = ChartValueType.DateTime;
            chart13.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy";
            chart13.Series[2].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy}";
            chart13.Series[2].IsValueShownAsLabel = false;



            Series newSeries13_4 = new Series();
            chart13.Series.Add(newSeries13_4);
            chart13.Series[3].ChartType = SeriesChartType.Line;
            chart13.Series[3].BorderWidth = 1;
            chart13.Series[3].Color = Color.Gray;
            chart13.ChartAreas[0].AxisX.Interval = 1;
            chart13.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart13.Series[3].EmptyPointStyle.Color = Color.Transparent;
            chart13.Series[3].XValueType = ChartValueType.DateTime;
            chart13.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy";
            chart13.Series[3].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy}";
            chart13.Series[3].IsValueShownAsLabel = false;



            Series newSeries13_5 = new Series();
            chart13.Series.Add(newSeries13_5);
            chart13.Series[4].ChartType = SeriesChartType.Line;
            chart13.Series[4].BorderWidth = 1;
            chart13.Series[4].Color = Color.Gray;
            chart13.ChartAreas[0].AxisX.Interval = 1;
            chart13.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart13.Series[4].EmptyPointStyle.Color = Color.Transparent;
            chart13.Series[4].XValueType = ChartValueType.DateTime;
            chart13.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy";
            chart13.Series[4].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy}";
            chart13.Series[4].IsValueShownAsLabel = false;


            Series newSeries13_6 = new Series();
            chart13.Series.Add(newSeries13_6);
            chart13.Series[5].ChartType = SeriesChartType.Line;
            chart13.Series[5].BorderWidth = 1;
            chart13.Series[5].Color = Color.Gray;
            chart13.ChartAreas[0].AxisX.Interval = 1;
            chart13.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart13.Series[5].EmptyPointStyle.Color = Color.Transparent;
            chart13.Series[5].XValueType = ChartValueType.DateTime;
            chart13.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy";
            chart13.Series[5].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy}";
            chart13.Series[5].IsValueShownAsLabel = false;


            Series newSeries13_7 = new Series();
            chart13.Series.Add(newSeries13_7);
            chart13.Series[6].ChartType = SeriesChartType.Line;
            chart13.Series[6].BorderWidth = 1;
            chart13.Series[6].Color = Color.Gray;
            chart13.ChartAreas[0].AxisX.Interval = 1;
            chart13.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart13.Series[6].EmptyPointStyle.Color = Color.Transparent;
            chart13.Series[6].XValueType = ChartValueType.DateTime;
            chart13.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy";
            chart13.Series[6].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy}";
            chart13.Series[6].IsValueShownAsLabel = false;



            Series newSeries13_8 = new Series();
            chart13.Series.Add(newSeries13_8);
            chart13.Series[7].ChartType = SeriesChartType.Line;
            chart13.Series[7].BorderWidth = 1;
            chart13.Series[7].Color = Color.Gray;
            chart13.ChartAreas[0].AxisX.Interval = 1;
            chart13.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart13.Series[7].EmptyPointStyle.Color = Color.Transparent;
            chart13.Series[7].XValueType = ChartValueType.DateTime;
            chart13.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy";
            chart13.Series[7].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy}";
            chart13.Series[7].IsValueShownAsLabel = false;




            Series newSeries13_9 = new Series();
            chart13.Series.Add(newSeries13_9);
            chart13.Series[8].ChartType = SeriesChartType.Line;
            chart13.Series[8].BorderWidth = 1;
            chart13.Series[8].Color = Color.Gray;
            chart13.ChartAreas[0].AxisX.Interval = 1;
            chart13.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart13.Series[8].EmptyPointStyle.Color = Color.Transparent;
            chart13.Series[8].XValueType = ChartValueType.DateTime;
            chart13.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy";
            chart13.Series[8].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy}";
            chart13.Series[8].IsValueShownAsLabel = false;


            Series newSeries13_10 = new Series();
            chart13.Series.Add(newSeries13_10);
            chart13.Series[9].ChartType = SeriesChartType.Line;
            chart13.Series[9].BorderWidth = 1;
            chart13.Series[9].Color = Color.Gray;
            chart13.ChartAreas[0].AxisX.Interval = 1;
            chart13.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart13.Series[9].EmptyPointStyle.Color = Color.Transparent;
            chart13.Series[9].XValueType = ChartValueType.DateTime;
            chart13.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy";
            chart13.Series[9].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy}";
            chart13.Series[9].IsValueShownAsLabel = false;


            Series newSeries13_11 = new Series();
            chart13.Series.Add(newSeries13_11);
            chart13.Series[10].ChartType = SeriesChartType.Line;
            chart13.Series[10].BorderWidth = 1;
            chart13.Series[10].Color = Color.Gray;
            chart13.ChartAreas[0].AxisX.Interval = 1;
            chart13.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart13.Series[10].EmptyPointStyle.Color = Color.Transparent;
            chart13.Series[10].XValueType = ChartValueType.DateTime;
            chart13.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy";
            chart13.Series[10].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy}";
            chart13.Series[10].IsValueShownAsLabel = false;


            Series newSeries13_12 = new Series();
            chart13.Series.Add(newSeries13_12);
            chart13.Series[11].ChartType = SeriesChartType.Line;
            chart13.Series[11].BorderWidth = 1;
            chart13.Series[11].Color = Color.Gray;
            chart13.ChartAreas[0].AxisX.Interval = 1;
            chart13.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart13.Series[11].EmptyPointStyle.Color = Color.Transparent;
            chart13.Series[11].XValueType = ChartValueType.DateTime;
            chart13.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy";
            chart13.Series[11].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy}";
            chart13.Series[11].IsValueShownAsLabel = false;


            Series newSeries13_13 = new Series();
            chart13.Series.Add(newSeries13_13);
            chart13.Series[12].ChartType = SeriesChartType.Line;
            chart13.Series[12].BorderWidth = 1;
            chart13.Series[12].Color = Color.Gray;
            chart13.ChartAreas[0].AxisX.Interval = 1;
            chart13.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart13.Series[12].EmptyPointStyle.Color = Color.Transparent;
            chart13.Series[12].XValueType = ChartValueType.DateTime;
            chart13.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy";
            chart13.Series[12].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy}";
            chart13.Series[12].IsValueShownAsLabel = false;


            Series newSeries13_14 = new Series();
            chart13.Series.Add(newSeries13_14);
            chart13.Series[13].ChartType = SeriesChartType.Line;
            chart13.Series[13].BorderWidth = 1;
            chart13.Series[13].Color = Color.Gray;
            chart13.ChartAreas[0].AxisX.Interval = 1;
            chart13.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart13.Series[13].EmptyPointStyle.Color = Color.Transparent;
            chart13.Series[13].XValueType = ChartValueType.DateTime;
            chart13.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy";
            chart13.Series[13].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy}";
            chart13.Series[13].IsValueShownAsLabel = false;



            Series newSeries13_15 = new Series();
            chart13.Series.Add(newSeries13_15);
            chart13.Series[14].ChartType = SeriesChartType.Line;
            chart13.Series[14].BorderWidth = 1;
            chart13.Series[14].Color = Color.Gray;
            chart13.ChartAreas[0].AxisX.Interval = 1;
            chart13.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart13.Series[14].EmptyPointStyle.Color = Color.Transparent;
            chart13.Series[14].XValueType = ChartValueType.DateTime;
            chart13.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy";
            chart13.Series[14].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy}";
            chart13.Series[14].IsValueShownAsLabel = false;


            Series newSeries13_16 = new Series();
            chart13.Series.Add(newSeries13_16);
            chart13.Series[15].ChartType = SeriesChartType.Line;
            chart13.Series[15].BorderWidth = 1;
            chart13.Series[15].Color = Color.Gray;
            chart13.ChartAreas[0].AxisX.Interval = 1;
            chart13.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart13.Series[15].EmptyPointStyle.Color = Color.Transparent;
            chart13.Series[15].XValueType = ChartValueType.DateTime;
            chart13.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy";
            chart13.Series[15].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy}";
            chart13.Series[15].IsValueShownAsLabel = false;


            Series newSeries13_17 = new Series();
            chart13.Series.Add(newSeries13_17);
            chart13.Series[16].ChartType = SeriesChartType.Line;
            chart13.Series[16].BorderWidth = 1;
            chart13.Series[16].Color = Color.Gray;
            chart13.ChartAreas[0].AxisX.Interval = 1;
            chart13.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart13.Series[16].EmptyPointStyle.Color = Color.Transparent;
            chart13.Series[16].XValueType = ChartValueType.DateTime;
            chart13.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy";
            chart13.Series[16].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy}";
            chart13.Series[16].IsValueShownAsLabel = false;



            Series newSeries13_18 = new Series();
            chart13.Series.Add(newSeries13_18);
            chart13.Series[17].ChartType = SeriesChartType.Line;
            chart13.Series[17].BorderWidth = 1;
            chart13.Series[17].Color = Color.Gray;
            chart13.ChartAreas[0].AxisX.Interval = 1;
            chart13.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart13.Series[17].EmptyPointStyle.Color = Color.Transparent;
            chart13.Series[17].XValueType = ChartValueType.DateTime;
            chart13.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy";
            chart13.Series[17].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy}";
            chart13.Series[17].IsValueShownAsLabel = false;


            Series newSeries13_19 = new Series();
            chart13.Series.Add(newSeries13_19);
            chart13.Series[18].ChartType = SeriesChartType.Line;
            chart13.Series[18].BorderWidth = 1;
            chart13.Series[18].Color = Color.Gray;
            chart13.ChartAreas[0].AxisX.Interval = 1;
            chart13.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart13.Series[18].EmptyPointStyle.Color = Color.Transparent;
            chart13.Series[18].XValueType = ChartValueType.DateTime;
            chart13.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy";
            chart13.Series[18].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy}";
            chart13.Series[18].IsValueShownAsLabel = false;

            Series newSeries13_20 = new Series();
            chart13.Series.Add(newSeries13_20);
            chart13.Series[19].ChartType = SeriesChartType.Line;
            chart13.Series[19].BorderWidth = 1;
            chart13.Series[19].Color = Color.Gray;
            chart13.ChartAreas[0].AxisX.Interval = 1;
            chart13.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart13.Series[19].EmptyPointStyle.Color = Color.Transparent;
            chart13.Series[19].XValueType = ChartValueType.DateTime;
            chart13.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy";
            chart13.Series[19].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy}";
            chart13.Series[19].IsValueShownAsLabel = false;


            Series newSeries13_21 = new Series();
            chart13.Series.Add(newSeries13_21);
            chart13.Series[20].ChartType = SeriesChartType.Line;
            chart13.Series[20].BorderWidth = 1;
            chart13.Series[20].Color = Color.Gray;
            chart13.ChartAreas[0].AxisX.Interval = 1;
            chart13.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart13.Series[20].EmptyPointStyle.Color = Color.Transparent;
            chart13.Series[20].XValueType = ChartValueType.DateTime;
            chart13.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy";
            chart13.Series[20].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy}";
            chart13.Series[20].IsValueShownAsLabel = false;


            Series newSeries13_22 = new Series();
            chart13.Series.Add(newSeries13_22);
            chart13.Series[21].ChartType = SeriesChartType.Line;
            chart13.Series[21].BorderWidth = 1;
            chart13.Series[21].Color = Color.Gray;
            chart13.ChartAreas[0].AxisX.Interval = 1;
            chart13.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart13.Series[21].EmptyPointStyle.Color = Color.Transparent;
            chart13.Series[21].XValueType = ChartValueType.DateTime;
            chart13.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy";
            chart13.Series[21].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy}";
            chart13.Series[21].IsValueShownAsLabel = false;



            Series newSeries13_23 = new Series();
            chart13.Series.Add(newSeries13_23);
            chart13.Series[22].ChartType = SeriesChartType.Line;
            chart13.Series[22].BorderWidth = 1;
            chart13.Series[22].Color = Color.Gray;
            chart13.ChartAreas[0].AxisX.Interval = 1;
            chart13.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart13.Series[22].EmptyPointStyle.Color = Color.Transparent;
            chart13.Series[22].XValueType = ChartValueType.DateTime;
            chart13.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy";
            chart13.Series[22].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy}";
            chart13.Series[22].IsValueShownAsLabel = false;


            Series newSeries13_24 = new Series();
            chart13.Series.Add(newSeries13_24);
            chart13.Series[23].ChartType = SeriesChartType.Line;
            chart13.Series[23].BorderWidth = 1;
            chart13.Series[23].Color = Color.Gray;
            chart13.ChartAreas[0].AxisX.Interval = 1;
            chart13.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart13.Series[23].EmptyPointStyle.Color = Color.Transparent;
            chart13.Series[23].XValueType = ChartValueType.DateTime;
            chart13.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy";
            chart13.Series[23].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy}";
            chart13.Series[23].IsValueShownAsLabel = false;



            Series newSeries13_25 = new Series();
            chart13.Series.Add(newSeries13_25);
            chart13.Series[24].ChartType = SeriesChartType.Line;
            chart13.Series[24].BorderWidth = 1;
            chart13.Series[24].Color = Color.Gray;
            chart13.ChartAreas[0].AxisX.Interval = 1;
            chart13.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart13.Series[24].EmptyPointStyle.Color = Color.Transparent;
            chart13.Series[24].XValueType = ChartValueType.DateTime;
            chart13.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy";
            chart13.Series[24].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy}";
            chart13.Series[24].IsValueShownAsLabel = false;


            Series newSeries13_26 = new Series();
            chart13.Series.Add(newSeries13_26);
            chart13.Series[25].ChartType = SeriesChartType.Line;
            chart13.Series[25].BorderWidth = 1;
            chart13.Series[25].Color = Color.Gray;
            chart13.ChartAreas[0].AxisX.Interval = 1;
            chart13.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart13.Series[25].EmptyPointStyle.Color = Color.Transparent;
            chart13.Series[25].XValueType = ChartValueType.DateTime;
            chart13.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy";
            chart13.Series[25].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy}";
            chart13.Series[25].IsValueShownAsLabel = false;


            Series newSeries13_27 = new Series();
            chart13.Series.Add(newSeries13_27);
            chart13.Series[26].ChartType = SeriesChartType.Line;
            chart13.Series[26].BorderWidth = 1;
            chart13.Series[26].Color = Color.Gray;
            chart13.ChartAreas[0].AxisX.Interval = 1;
            chart13.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart13.Series[26].EmptyPointStyle.Color = Color.Transparent;
            chart13.Series[26].XValueType = ChartValueType.DateTime;
            chart13.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM/yyyy";
            chart13.Series[26].ToolTip = "#VALY{F}\n#VALX{dd/MM/yyyy}";
            chart13.Series[26].IsValueShownAsLabel = false;

            chart13.Legends.Clear();


            if (comboBox3.SelectedItem.ToString() != "RD4")
            {

                string Selected_Vendor = "";
                for (int y = 1; y <= Node_Vendor.Length / 2 - 1; y++)
                {
                    if (Node_Vendor[y - 1, 0] == comboBox4.SelectedItem.ToString())
                    {
                        Selected_Vendor = Node_Vendor[y - 1, 1];
                    }
                }


                double min_value_Data_Type = 100000000000;
                double max_value_Data_Type = -100000000000;
                for (int i = 0; i < Nodes_Table.Rows.Count; i++)
                {
                    if ((Nodes_Table.Rows[i]).ItemArray[0].ToString() == comboBox4.SelectedItem.ToString())
                    {
                        DateTime dt = DateTime.Today;
                        string Data_Value = "";
                        string Traffic_Value = "";
                        if (Data_Type == "CC2")
                        {
                            dt = Convert.ToDateTime((Nodes_Table.Rows[i]).ItemArray[7]);
                            Data_Value = (Nodes_Table.Rows[i]).ItemArray[43].ToString();
                            Traffic_Value = (Nodes_Table.Rows[i]).ItemArray[8].ToString();
                        }
                        if (Data_Type == "RD2")
                        {
                            dt = Convert.ToDateTime((Nodes_Table.Rows[i]).ItemArray[7]);
                            Data_Value = (Nodes_Table.Rows[i]).ItemArray[36].ToString();
                            Traffic_Value = (Nodes_Table.Rows[i]).ItemArray[8].ToString();
                        }
                        if (Data_Type == "CC3")
                        {
                            dt = Convert.ToDateTime((Nodes_Table.Rows[i]).ItemArray[8]);
                            Data_Value = (Nodes_Table.Rows[i]).ItemArray[37].ToString();
                            Traffic_Value = (Nodes_Table.Rows[i]).ItemArray[10].ToString();
                        }
                        if (Data_Type == "RD3")
                        {
                            dt = Convert.ToDateTime((Nodes_Table.Rows[i]).ItemArray[8]);
                            Data_Value = (Nodes_Table.Rows[i]).ItemArray[107].ToString();
                            Traffic_Value = (Nodes_Table.Rows[i]).ItemArray[10].ToString();
                        }




                        if (Data_Value != "")
                        {
                            double Data_Type_Value = Convert.ToDouble(Data_Value);
                            double Traffic = Convert.ToDouble(Traffic_Value);

                            chart6.Series[1].Points.AddXY(dt, Data_Type_Value);
                            chart6.Series[0].Points.AddXY(dt, Traffic);


                            chart13.Series[0].Points.AddXY(dt, Data_Type_Value);


                            if (Data_Type_Value > max_value_Data_Type)
                            {
                                max_value_Data_Type = Math.Round(Data_Type_Value, MidpointRounding.AwayFromZero);
                            }
                            if (Data_Type_Value < min_value_Data_Type)
                            {
                                min_value_Data_Type = Math.Round(Data_Type_Value, MidpointRounding.AwayFromZero);
                            }
                        }


                        chart6.ChartAreas[0].AxisY.Maximum = max_value_Data_Type + 1;
                        chart6.ChartAreas[0].AxisY.Minimum = min_value_Data_Type - 1;
                        chart6.Series[1].LegendText = comboBox3.SelectedItem.ToString();
                        chart13.Series[0].ToolTip = comboBox4.SelectedItem.ToString() + "\n#VALY{F}\n#VALX{dd/MM/yyyy}";
                        if (Data_Type == "CC2" || Data_Type == "CC3")
                        {
                            chart6.Series[0].LegendText = "Traffic (Erlang)";
                        }
                        else
                        {
                            chart6.Series[0].LegendText = "Payload (GB)";
                        }



                    }


                }






                // Ploting all nodes
                for (int k = 1; k <= comboBox4.Items.Count; k++)
                {


                    string Node = comboBox4.Items[k - 1].ToString();

                    string Vendor_From_List = "";
                    for (int y = 1; y <= Node_Vendor.Length / 2 - 1; y++)
                    {
                        if (Node_Vendor[y - 1, 0] == Node)
                        {
                            Vendor_From_List = Node_Vendor[y - 1, 1];
                        }
                    }

                    if (Vendor_From_List == Selected_Vendor)
                    {

                        if (Node != comboBox4.SelectedItem.ToString())
                        {
                            if (comboBox3.SelectedItem.ToString() == "CC2" || comboBox3.SelectedItem.ToString() == "RD2")
                            {
                                var node_data = (from p in Nodes_Table.AsEnumerable()
                                                 where p.Field<string>("BSC Name") == Node
                                                 select p).ToList();


                                for (int i = 0; i < node_data.Count; i++)
                                {


                                    DateTime dt = DateTime.Today;
                                    string Data_Value = "";
                                    string Traffic_Value = "";

                                    if (Data_Type == "CC2")
                                    {
                                        dt = Convert.ToDateTime(node_data[i].ItemArray[7]);
                                        Data_Value = node_data[i].ItemArray[43].ToString();
                                        Traffic_Value = node_data[i].ItemArray[8].ToString();
                                    }
                                    if (Data_Type == "RD2")
                                    {
                                        dt = Convert.ToDateTime(node_data[i].ItemArray[7]);
                                        Data_Value = node_data[i].ItemArray[36].ToString();
                                        Traffic_Value = node_data[i].ItemArray[8].ToString();
                                    }


                                    if (Data_Value != "")
                                    {
                                        double Data_Type_Value = Convert.ToDouble(Data_Value);
                                        double Traffic = Convert.ToDouble(Traffic_Value);

                                        chart13.Series[k].Points.AddXY(dt, Data_Type_Value);

                                        chart13.Series[k].ToolTip = Node + "\n#VALY{F}\n#VALX{dd/MM/yyyy}";

                                        if (Data_Type_Value > max_value_Data_Type)
                                        {
                                            max_value_Data_Type = Math.Round(Data_Type_Value, MidpointRounding.AwayFromZero);
                                        }
                                        if (Data_Type_Value < min_value_Data_Type)
                                        {
                                            min_value_Data_Type = Math.Round(Data_Type_Value, MidpointRounding.AwayFromZero);
                                        }


                                    }


                                }

                            }







                            if (comboBox3.SelectedItem.ToString() == "CC3" || comboBox3.SelectedItem.ToString() == "RD3")
                            {
                                var node_data = (from p in Nodes_Table.AsEnumerable()
                                                 where p.Field<string>("RNC") == Node
                                                 select p).ToList();


                                for (int i = 0; i < node_data.Count; i++)
                                {


                                    DateTime dt = DateTime.Today;
                                    string Data_Value = "";
                                    string Traffic_Value = "";

                                    if (Data_Type == "CC3")
                                    {
                                        dt = Convert.ToDateTime(node_data[i].ItemArray[8]);
                                        Data_Value = node_data[i].ItemArray[37].ToString();
                                        Traffic_Value = node_data[i].ItemArray[10].ToString();
                                    }
                                    if (Data_Type == "RD3")
                                    {
                                        dt = Convert.ToDateTime(node_data[i].ItemArray[8]);
                                        Data_Value = node_data[i].ItemArray[107].ToString();
                                        Traffic_Value = node_data[i].ItemArray[10].ToString();
                                    }



                                    if (Data_Value != "")
                                    {
                                        double Data_Type_Value = Convert.ToDouble(Data_Value);
                                        double Traffic = Convert.ToDouble(Traffic_Value);

                                        chart13.Series[k].Points.AddXY(dt, Data_Type_Value);

                                        chart13.Series[k].ToolTip = Node + "\n#VALY{F}\n#VALX{dd/MM/yyyy}";

                                        if (Data_Type_Value > max_value_Data_Type)
                                        {
                                            max_value_Data_Type = Math.Round(Data_Type_Value, MidpointRounding.AwayFromZero);
                                        }
                                        if (Data_Type_Value < min_value_Data_Type)
                                        {
                                            min_value_Data_Type = Math.Round(Data_Type_Value, MidpointRounding.AwayFromZero);
                                        }


                                    }


                                }

                            }






                        }


                    }


                }


                chart13.ChartAreas[0].AxisY.Maximum = max_value_Data_Type + 1;
                chart13.ChartAreas[0].AxisY.Minimum = min_value_Data_Type - 1;
                chart13.Series[0].LegendText = comboBox4.SelectedItem.ToString();





                double difference_day = (end_date - start_date).TotalDays;
                double day_interval = Math.Round(difference_day / 20);
                if (day_interval == 0)
                {
                    day_interval = 1;
                }
                chart6.ChartAreas[0].AxisX.Interval = day_interval;
                chart13.ChartAreas[0].AxisX.Interval = day_interval;


            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            chart1.SaveImage("Image.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
            Image img = Image.FromFile("Image.jpg");
            System.Windows.Forms.Clipboard.SetImage(img);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            chart19.SaveImage("Image.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
            Image img = Image.FromFile("Image.jpg");
            System.Windows.Forms.Clipboard.SetImage(img);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            XLWorkbook wb = new XLWorkbook();
            if (Provinces_Table.Rows.Count != 0)
            {
                wb.Worksheets.Add(Provinces_Table, "Province Table");
            }
            if (Nodes_Table.Rows.Count != 0)
            {
                wb.Worksheets.Add(Nodes_Table, "Node Table");
            }
            //wb.Worksheets.Add(Nodes_Table, "Node Table");
            // wb.Worksheets.Add(Single_Region_Table_Contractual, "Status");
            var saveFileDialog = new SaveFileDialog
            {
                FileName = comboBox1.SelectedItem.ToString() + "_" + comboBox3.SelectedItem.ToString(),
                Filter = "Excel files|*.xlsx",
                Title = "Save an Excel File"
            };



            saveFileDialog.ShowDialog();

            if (!String.IsNullOrWhiteSpace(saveFileDialog.FileName))
                wb.SaveAs(saveFileDialog.FileName);

            MessageBox.Show("Finished");

        }

        private void button5_Click(object sender, EventArgs e)
        {
            chart29.SaveImage("Image.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
            Image img = Image.FromFile("Image.jpg");
            System.Windows.Forms.Clipboard.SetImage(img);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            chart34.SaveImage("Image.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
            Image img = Image.FromFile("Image.jpg");
            System.Windows.Forms.Clipboard.SetImage(img);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            chart6.SaveImage("Image.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
            Image img = Image.FromFile("Image.jpg");
            System.Windows.Forms.Clipboard.SetImage(img);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            chart13.SaveImage("Image.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
            Image img = Image.FromFile("Image.jpg");
            System.Windows.Forms.Clipboard.SetImage(img);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            chart39.SaveImage("Image.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
            Image img = Image.FromFile("Image.jpg");
            System.Windows.Forms.Clipboard.SetImage(img);
        }
    }
}
