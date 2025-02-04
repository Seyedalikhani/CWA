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
using ClosedXML.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace CWA
{
    public partial class LTE : Form
    {
        public LTE()
        {
            InitializeComponent();
        }


        public Main form1;


        public LTE(Form form)
        {
            InitializeComponent();
            form1 = (Main)form;
        }

        public string ConnectionString = "";
        public SqlConnection connection = new SqlConnection();


        public string Connection_Type = "DataBase";

       // public string Server_Name = "172.26.7.159";
        public string DataBase_Name = "Performance_NAK";
       public string Server_Name = "PERFORMANCEDB";

        public DataTable LTE_RNC_Data_Table = new DataTable();
        public string[,] Series_List = new string[30, 2];

        int RNC_Num = 0;

        public string[,] MIN_MAX_KPI_MAT = new string[31, 29];


        public DateTime Min_X1 = DateTime.Today;
        public DateTime Max_X1 = DateTime.Today;
        public double Min_X = 10000000000;
        public double Max_X = -10000000000;
        public DateTime Min_X_Date = DateTime.Today;
        public DateTime Max_X_Date = DateTime.Today;


        public string Interval = "Daily";





        private void chart1_MouseWheel(object sender, MouseEventArgs e)
        {
            var chart = (Chart)sender;
            var xAxis = chart.ChartAreas[0].AxisX;
            var yAxis = chart.ChartAreas[0].AxisY;

            try
            {
                if (e.Delta < 0) // Scrolled down.
                {
                    xAxis.ScaleView.ZoomReset();
                    yAxis.ScaleView.ZoomReset();
                }
                else if (e.Delta > 0) // Scrolled up.
                {
                    var xMin = xAxis.ScaleView.ViewMinimum;
                    var xMax = xAxis.ScaleView.ViewMaximum;
                    var yMin = yAxis.ScaleView.ViewMinimum;
                    var yMax = yAxis.ScaleView.ViewMaximum;

                    var posXStart = xAxis.PixelPositionToValue(e.Location.X) - (xMax - xMin) / 4;
                    var posXFinish = xAxis.PixelPositionToValue(e.Location.X) + (xMax - xMin) / 4;
                    var posYStart = yAxis.PixelPositionToValue(e.Location.Y) - (yMax - yMin) / 4;
                    var posYFinish = yAxis.PixelPositionToValue(e.Location.Y) + (yMax - yMin) / 4;

                    xAxis.ScaleView.Zoom(posXStart, posXFinish);
                    yAxis.ScaleView.Zoom(posYStart, posYFinish);
                }
            }
            catch { }
        }







        private void Form8_Load(object sender, EventArgs e)
        {



            string userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            string currentUser = userName.Substring(8, userName.Length - 8);


            string[] authorizedUsers = new string[]
    {
"Arineh.Badalians",
"arineh.badalians",
"Forough.Hosseini",
"forough.hosseini",
"Mahdi.Heidary",
"mahdi.heidary",
"Mahta.Labbafi",
"mahta.labbafi",
"Milad.Afzal",
"milad.afzal",
"Mohammadali.Amini",
"mohammadali.amini",
"Moharram.Tofighi",
"moharram.tofighi",
"Reza.Akbari",
"reza.akbari",
"Reza.Akbari",
"R.akbari",
"r.akbari",
"R-akbari",
"r-akbari",
"R.Fallah",
"r.fallah",
"rfallah",
"Robabeh.Falah",
"robabeh.falah",
"Elham.Vafaeinejad",
"elham.vafaeinejad",
"Arash.Naghdehforoushha",
"arash.naghdehforoushha",
"Majedeh.Seydi",
"majedeh.seydi",
"a.mohammadiraeisi",
"ahmad.alikhani",
"mohammad.gorji"

    };

            if (authorizedUsers.Contains(currentUser.ToLower()))
            {
                string Authorized = "OK";
            }
            else
            {
                MessageBox.Show("Limited Access! Need Authorization by Admin");
                this.Close();
            }


            // Log
            //var excelApplication = new Excel.Application();
            //var excelWorkBook = excelApplication.Application.Workbooks.Add(Type.Missing);
            //excelApplication.Cells[1, 1] = currentUser;
            //excelApplication.Cells[1, 2] = Convert.ToString(DateTime.Now);
            //string name1 = @"\\dfs\fs\NPO\6. Performance\Contractual WPC\New folder\Old\LR\LTE";
            //string CR_PATH = string.Format(name1+".xlsx");
            //excelApplication.ActiveWorkbook.SaveCopyAs(CR_PATH);
            //excelApplication.ActiveWorkbook.Saved = true;
            //excelApplication.Quit();







            ConnectionString = @"Server=" + Server_Name + "; Database=" + DataBase_Name + "; User ID=cwpcApp;Password=cwpcApp@830625#Ahmad";

           // ConnectionString = @"Server=" + Server_Name + "; Database=" + DataBase_Name + "; Trusted_Connection=True;";
            connection = new SqlConnection(ConnectionString);
            connection.Open();








            panel1.AutoScroll = true;
            panel1.BorderStyle = BorderStyle.FixedSingle;
            panel1.Size = new Size(1200, 3000);


            MIN_MAX_KPI_MAT[0, 0] = "RNC";
            MIN_MAX_KPI_MAT[0, 1] = "Payload_Max";
            MIN_MAX_KPI_MAT[0, 2] = "Payload_Min";
            MIN_MAX_KPI_MAT[0, 3] = "Availability_Max";
            MIN_MAX_KPI_MAT[0, 4] = "Availability_Min";
            MIN_MAX_KPI_MAT[0, 5] = "DL_User_THR_Max";
            MIN_MAX_KPI_MAT[0, 6] = "DL_User_THR_Min";
            MIN_MAX_KPI_MAT[0, 7] = "DL_Cell_THR_Max";
            MIN_MAX_KPI_MAT[0, 8] = "DL_Cell_THR_Min";
            MIN_MAX_KPI_MAT[0, 9] = "UL_User_THR_Max";
            MIN_MAX_KPI_MAT[0, 10] = "UL_User_THR_Min";
            MIN_MAX_KPI_MAT[0, 11] = "Latency_Max";
            MIN_MAX_KPI_MAT[0, 12] = "Latency_Min";
            MIN_MAX_KPI_MAT[0, 13] = "Service_Max";
            MIN_MAX_KPI_MAT[0, 14] = "Service_Min";
            MIN_MAX_KPI_MAT[0, 15] = "RRC_Max";
            MIN_MAX_KPI_MAT[0, 16] = "RRC_Min";
            MIN_MAX_KPI_MAT[0, 17] = "ERAB_Setup_Max";
            MIN_MAX_KPI_MAT[0, 18] = "ERAB_Setup_Min";
            MIN_MAX_KPI_MAT[0, 19] = "ERAB_Drop_Max";
            MIN_MAX_KPI_MAT[0, 20] = "ERAB_Drop_Min";
            MIN_MAX_KPI_MAT[0, 21] = "Inter_Max";
            MIN_MAX_KPI_MAT[0, 22] = "Inter_Min";
            MIN_MAX_KPI_MAT[0, 23] = "Intra_Max";
            MIN_MAX_KPI_MAT[0, 24] = "Intra_Min";
            MIN_MAX_KPI_MAT[0, 25] = "PUCCH_Max";
            MIN_MAX_KPI_MAT[0, 26] = "PUCCH_Min";
            MIN_MAX_KPI_MAT[0, 27] = "PUSCH_Max";
            MIN_MAX_KPI_MAT[0, 28] = "PUSCH_Min";




            chart1.Size = new Size(565, 400);
            chart1.Location = new Point(20, 30);
            chart2.Size = new Size(565, 400);
            chart2.Location = new Point(620, 30);

            chart3.Size = new Size(565, 400);
            chart3.Location = new Point(20, 450);
            chart4.Size = new Size(565, 400);
            chart4.Location = new Point(620, 450);

            chart5.Size = new Size(565, 400);
            chart5.Location = new Point(20, 870);
            chart6.Size = new Size(565, 400);
            chart6.Location = new Point(620, 870);

            chart7.Size = new Size(565, 400);
            chart7.Location = new Point(20, 1290);
            chart8.Size = new Size(565, 400);
            chart8.Location = new Point(620, 1290);


            chart9.Size = new Size(565, 400);
            chart9.Location = new Point(20, 1710);
            chart10.Size = new Size(565, 400);
            chart10.Location = new Point(620, 1710);

            chart11.Size = new Size(565, 400);
            chart11.Location = new Point(20, 2130);
            chart12.Size = new Size(565, 400);
            chart12.Location = new Point(620, 2130);


            chart13.Size = new Size(565, 400);
            chart13.Location = new Point(20, 2565);
            chart14.Size = new Size(565, 400);
            chart14.Location = new Point(620, 2565);



            string Date_List = "";
            if (Interval=="Daily")
            {
                Date_List = @"select distinct Datetime from [LTE_RNC_TH_Last_Day] order by Datetime";
            }
            if (Interval == "BH")
            {
                Date_List = @"select distinct Datetime from [LTE_RNC_TH_Last_Day_BH] order by Day";
            }



            SqlCommand Date_List_Quary = new SqlCommand(Date_List, connection);
            Date_List_Quary.CommandTimeout = 0;
            Date_List_Quary.ExecuteNonQuery();
            DataTable Date_List_Table = new DataTable();
            SqlDataAdapter dataAdapter_Date_List_Table = new SqlDataAdapter(Date_List_Quary);
            dataAdapter_Date_List_Table.Fill(Date_List_Table);

            for (int i = 0; i < Date_List_Table.Rows.Count; i++)
            {

                string Date = (Date_List_Table.Rows[i]).ItemArray[0].ToString();
                DateTime Day = Convert.ToDateTime((Date_List_Table.Rows[i]).ItemArray[0]);
              //  string Day1 = Date.Substring(0, 10) + " " + Day.DayOfWeek.ToString();
                string Day1 = Convert.ToString(Day.Month)+"/"+ Convert.ToString(Day.Day) + "/"+ Convert.ToString(Day.Year)+ " " + Day.DayOfWeek.ToString();
                listBox2.Items.Add(Day1);

            }


        }

        private void button5_Click(object sender, EventArgs e)
        {

            listBox1.Items.Clear();
            RNC_Num = 0;


            string LTE_RNC_Data = "";
            if (Interval=="Daily")
            {
                LTE_RNC_Data = @"select * from [LTE_RNC_TH] order by RNC, Datetime";
            }
            if (Interval == "BH")
            {
                LTE_RNC_Data = @"select * from [LTE_RNC_TH_BH] order by RNC, Day";
            }

           // LTE_RNC_Data = @"select * from [LTE_RNC_TH] order by RNC, Datetime";
            SqlCommand LTE_RNC_Data_Quary = new SqlCommand(LTE_RNC_Data, connection);
            LTE_RNC_Data_Quary.CommandTimeout = 0;
            LTE_RNC_Data_Quary.ExecuteNonQuery();
            LTE_RNC_Data_Table = new DataTable();
            SqlDataAdapter dataAdapter_LTE_RNC_Data_Table = new SqlDataAdapter(LTE_RNC_Data_Quary);
            dataAdapter_LTE_RNC_Data_Table.Fill(LTE_RNC_Data_Table);



            string RNC_List = "";
            if (Interval == "Daily")
            {
                RNC_List = @"select distinct RNC from  [LTE_RNC_TH] order by RNC";
            }
            if (Interval == "BH")
            {
                RNC_List = @"select distinct RNC from  [LTE_RNC_TH_BH] order by RNC";
            }

            //RNC_List = @"select distinct RNC from  [LTE_RNC_TH] order by RNC";
            SqlCommand RNC_List_Quary = new SqlCommand(RNC_List, connection);
            RNC_List_Quary.CommandTimeout = 0;
            RNC_List_Quary.ExecuteNonQuery();
            DataTable RNC_List_Table = new DataTable();
            SqlDataAdapter dataAdapter_RNC_List_Table = new SqlDataAdapter(RNC_List_Quary);
            dataAdapter_RNC_List_Table.Fill(RNC_List_Table);

            Series_List = new string[30, 2];

            for (int i = 0; i < RNC_List_Table.Rows.Count; i++)
            {
                RNC_Num++;
                string RNC = (RNC_List_Table.Rows[i]).ItemArray[0].ToString();
                Series_List[i, 0] = RNC;
                Series_List[i, 1] = "S" + Convert.ToString(i + 1);
                listBox1.Items.Add(RNC);

            }



            chart1.Series.Clear();
            chart1.Titles.Clear();

            Series newSeries1 = new Series(); chart1.Series.Add(newSeries1);
            Series newSeries2 = new Series(); chart1.Series.Add(newSeries2);
            Series newSeries3 = new Series(); chart1.Series.Add(newSeries3);
            Series newSeries4 = new Series(); chart1.Series.Add(newSeries4);
            Series newSeries5 = new Series(); chart1.Series.Add(newSeries5);
            Series newSeries6 = new Series(); chart1.Series.Add(newSeries6);
            Series newSeries7 = new Series(); chart1.Series.Add(newSeries7);
            Series newSeries8 = new Series(); chart1.Series.Add(newSeries8);
            Series newSeries9 = new Series(); chart1.Series.Add(newSeries9);
            Series newSeries10 = new Series(); chart1.Series.Add(newSeries10);
            Series newSeries11 = new Series(); chart1.Series.Add(newSeries11);
            Series newSeries12 = new Series(); chart1.Series.Add(newSeries12);
            Series newSeries13 = new Series(); chart1.Series.Add(newSeries13);
            Series newSeries14 = new Series(); chart1.Series.Add(newSeries14);
            Series newSeries15 = new Series(); chart1.Series.Add(newSeries15);
            Series newSeries16 = new Series(); chart1.Series.Add(newSeries16);
            Series newSeries17 = new Series(); chart1.Series.Add(newSeries17);
            Series newSeries18 = new Series(); chart1.Series.Add(newSeries18);
            Series newSeries19 = new Series(); chart1.Series.Add(newSeries19);
            Series newSeries20 = new Series(); chart1.Series.Add(newSeries20);
            Series newSeries21 = new Series(); chart1.Series.Add(newSeries21);
            Series newSeries22 = new Series(); chart1.Series.Add(newSeries22);
            Series newSeries23 = new Series(); chart1.Series.Add(newSeries23);
            Series newSeries24 = new Series(); chart1.Series.Add(newSeries24);
            Series newSeries25 = new Series(); chart1.Series.Add(newSeries25);



            Title title1 = chart1.Titles.Add("Payload (TB)");
            title1.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            chart1.Legends["Legend1"].Docking = Docking.Bottom;
            chart1.ChartAreas[0].AxisX.Interval = 5;
            chart1.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart1.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";

            for (int i = 0; i < RNC_List_Table.Rows.Count; i++)
            {
                chart1.Series[i].ChartType = SeriesChartType.Line;
                chart1.Series[i].BorderWidth = 3;
                chart1.Series[i].EmptyPointStyle.Color = Color.Transparent;
                chart1.Series[i].XValueType = ChartValueType.DateTime;
                chart1.Series[i].ToolTip = Series_List[i, 0] + "_" + "#VALX [#VALY]";
                chart1.Series[i].IsValueShownAsLabel = false;
            }


         

            chart2.Series.Clear();
            chart2.Titles.Clear();

            Series newSeries26 = new Series(); chart2.Series.Add(newSeries26);
            Series newSeries27 = new Series(); chart2.Series.Add(newSeries27);
            Series newSeries28 = new Series(); chart2.Series.Add(newSeries28);
            Series newSeries29 = new Series(); chart2.Series.Add(newSeries29);
            Series newSeries30 = new Series(); chart2.Series.Add(newSeries30);
            Series newSeries31 = new Series(); chart2.Series.Add(newSeries31);
            Series newSeries32 = new Series(); chart2.Series.Add(newSeries32);
            Series newSeries33 = new Series(); chart2.Series.Add(newSeries33);
            Series newSeries34 = new Series(); chart2.Series.Add(newSeries34);
            Series newSeries35 = new Series(); chart2.Series.Add(newSeries35);
            Series newSeries36 = new Series(); chart2.Series.Add(newSeries36);
            Series newSeries37 = new Series(); chart2.Series.Add(newSeries37);
            Series newSeries38 = new Series(); chart2.Series.Add(newSeries38);
            Series newSeries39 = new Series(); chart2.Series.Add(newSeries39);
            Series newSeries40 = new Series(); chart2.Series.Add(newSeries40);
            Series newSeries41 = new Series(); chart2.Series.Add(newSeries41);
            Series newSeries42 = new Series(); chart2.Series.Add(newSeries42);
            Series newSeries43 = new Series(); chart2.Series.Add(newSeries43);
            Series newSeries44 = new Series(); chart2.Series.Add(newSeries44);
            Series newSeries45 = new Series(); chart2.Series.Add(newSeries45);
            Series newSeries46 = new Series(); chart2.Series.Add(newSeries46);
            Series newSeries47 = new Series(); chart2.Series.Add(newSeries47);
            Series newSeries48 = new Series(); chart2.Series.Add(newSeries48);
            Series newSeries49 = new Series(); chart2.Series.Add(newSeries49);
            Series newSeries50 = new Series(); chart2.Series.Add(newSeries50);



            Title title2 = chart2.Titles.Add("Availability");
            title2.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            chart2.Legends["Legend1"].Docking = Docking.Bottom;
            chart2.ChartAreas[0].AxisX.Interval = 5;
            chart2.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart2.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";

            for (int i = 0; i < RNC_List_Table.Rows.Count; i++)
            {
                chart2.Series[i].ChartType = SeriesChartType.Line;
                chart2.Series[i].BorderWidth = 3;
                chart2.Series[i].EmptyPointStyle.Color = Color.Transparent;
                chart2.Series[i].XValueType = ChartValueType.DateTime;
                chart2.Series[i].ToolTip = Series_List[i, 0] + "_" + "#VALX [#VALY]";
                chart2.Series[i].IsValueShownAsLabel = false;
            }



            chart3.Series.Clear();
            chart3.Titles.Clear();

            Series newSeries51 = new Series(); chart3.Series.Add(newSeries51);
            Series newSeries52 = new Series(); chart3.Series.Add(newSeries52);
            Series newSeries53 = new Series(); chart3.Series.Add(newSeries53);
            Series newSeries54 = new Series(); chart3.Series.Add(newSeries54);
            Series newSeries55 = new Series(); chart3.Series.Add(newSeries55);
            Series newSeries56 = new Series(); chart3.Series.Add(newSeries56);
            Series newSeries57 = new Series(); chart3.Series.Add(newSeries57);
            Series newSeries58 = new Series(); chart3.Series.Add(newSeries58);
            Series newSeries59 = new Series(); chart3.Series.Add(newSeries59);
            Series newSeries60 = new Series(); chart3.Series.Add(newSeries60);
            Series newSeries61 = new Series(); chart3.Series.Add(newSeries61);
            Series newSeries62 = new Series(); chart3.Series.Add(newSeries62);
            Series newSeries63 = new Series(); chart3.Series.Add(newSeries63);
            Series newSeries64 = new Series(); chart3.Series.Add(newSeries64);
            Series newSeries65 = new Series(); chart3.Series.Add(newSeries65);
            Series newSeries66 = new Series(); chart3.Series.Add(newSeries66);
            Series newSeries67 = new Series(); chart3.Series.Add(newSeries67);
            Series newSeries68 = new Series(); chart3.Series.Add(newSeries68);
            Series newSeries69 = new Series(); chart3.Series.Add(newSeries69);
            Series newSeries70 = new Series(); chart3.Series.Add(newSeries70);
            Series newSeries71 = new Series(); chart3.Series.Add(newSeries71);
            Series newSeries72 = new Series(); chart3.Series.Add(newSeries72);
            Series newSeries73 = new Series(); chart3.Series.Add(newSeries73);
            Series newSeries74 = new Series(); chart3.Series.Add(newSeries74);
            Series newSeries75 = new Series(); chart3.Series.Add(newSeries75);


            Title title3 = chart3.Titles.Add("DL User Throughput (Mbps)");
            title3.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            chart3.Legends["Legend1"].Docking = Docking.Bottom;
            chart3.ChartAreas[0].AxisX.Interval = 5;
            chart3.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart3.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";

            for (int i = 0; i < RNC_List_Table.Rows.Count; i++)
            {
                chart3.Series[i].ChartType = SeriesChartType.Line;
                chart3.Series[i].BorderWidth = 3;
                chart3.Series[i].EmptyPointStyle.Color = Color.Transparent;
                chart3.Series[i].XValueType = ChartValueType.DateTime;
                chart3.Series[i].ToolTip = Series_List[i, 0] + "_" + "#VALX [#VALY]";
                chart3.Series[i].IsValueShownAsLabel = false;
            }


            chart4.Series.Clear();
            chart4.Titles.Clear();

            Series newSeries76 = new Series(); chart4.Series.Add(newSeries76);
            Series newSeries77 = new Series(); chart4.Series.Add(newSeries77);
            Series newSeries78 = new Series(); chart4.Series.Add(newSeries78);
            Series newSeries79 = new Series(); chart4.Series.Add(newSeries79);
            Series newSeries80 = new Series(); chart4.Series.Add(newSeries80);
            Series newSeries81 = new Series(); chart4.Series.Add(newSeries81);
            Series newSeries82 = new Series(); chart4.Series.Add(newSeries82);
            Series newSeries83 = new Series(); chart4.Series.Add(newSeries83);
            Series newSeries84 = new Series(); chart4.Series.Add(newSeries84);
            Series newSeries85 = new Series(); chart4.Series.Add(newSeries85);
            Series newSeries86 = new Series(); chart4.Series.Add(newSeries86);
            Series newSeries87 = new Series(); chart4.Series.Add(newSeries87);
            Series newSeries88 = new Series(); chart4.Series.Add(newSeries88);
            Series newSeries89 = new Series(); chart4.Series.Add(newSeries89);
            Series newSeries90 = new Series(); chart4.Series.Add(newSeries90);
            Series newSeries91 = new Series(); chart4.Series.Add(newSeries91);
            Series newSeries92 = new Series(); chart4.Series.Add(newSeries92);
            Series newSeries93 = new Series(); chart4.Series.Add(newSeries93);
            Series newSeries94 = new Series(); chart4.Series.Add(newSeries94);
            Series newSeries95 = new Series(); chart4.Series.Add(newSeries95);
            Series newSeries96 = new Series(); chart4.Series.Add(newSeries96);
            Series newSeries97 = new Series(); chart4.Series.Add(newSeries97);
            Series newSeries98 = new Series(); chart4.Series.Add(newSeries98);
            Series newSeries99 = new Series(); chart4.Series.Add(newSeries99);
            Series newSeries100 = new Series(); chart4.Series.Add(newSeries100);




            Title title4 = chart4.Titles.Add("DL Cell Throughput (Mbps)");
            title4.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            chart4.Legends["Legend1"].Docking = Docking.Bottom;
            chart4.ChartAreas[0].AxisX.Interval = 5;
            chart4.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart4.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";

            for (int i = 0; i < RNC_List_Table.Rows.Count; i++)
            {
                chart4.Series[i].ChartType = SeriesChartType.Line;
                chart4.Series[i].BorderWidth = 3;
                chart4.Series[i].EmptyPointStyle.Color = Color.Transparent;
                chart4.Series[i].XValueType = ChartValueType.DateTime;
                chart4.Series[i].ToolTip = Series_List[i, 0] + "_" + "#VALX [#VALY]";
                chart4.Series[i].IsValueShownAsLabel = false;
            }


            chart5.Series.Clear();
            chart5.Titles.Clear();

            Series newSeries101 = new Series(); chart5.Series.Add(newSeries101);
            Series newSeries102 = new Series(); chart5.Series.Add(newSeries102);
            Series newSeries103 = new Series(); chart5.Series.Add(newSeries103);
            Series newSeries104 = new Series(); chart5.Series.Add(newSeries104);
            Series newSeries105 = new Series(); chart5.Series.Add(newSeries105);
            Series newSeries106 = new Series(); chart5.Series.Add(newSeries106);
            Series newSeries107 = new Series(); chart5.Series.Add(newSeries107);
            Series newSeries108 = new Series(); chart5.Series.Add(newSeries108);
            Series newSeries109 = new Series(); chart5.Series.Add(newSeries109);
            Series newSeries110 = new Series(); chart5.Series.Add(newSeries110);
            Series newSeries111 = new Series(); chart5.Series.Add(newSeries111);
            Series newSeries112 = new Series(); chart5.Series.Add(newSeries112);
            Series newSeries113 = new Series(); chart5.Series.Add(newSeries113);
            Series newSeries114 = new Series(); chart5.Series.Add(newSeries114);
            Series newSeries115 = new Series(); chart5.Series.Add(newSeries115);
            Series newSeries116 = new Series(); chart5.Series.Add(newSeries116);
            Series newSeries117 = new Series(); chart5.Series.Add(newSeries117);
            Series newSeries118 = new Series(); chart5.Series.Add(newSeries118);
            Series newSeries119 = new Series(); chart5.Series.Add(newSeries119);
            Series newSeries120 = new Series(); chart5.Series.Add(newSeries120);
            Series newSeries121 = new Series(); chart5.Series.Add(newSeries121);
            Series newSeries122 = new Series(); chart5.Series.Add(newSeries122);
            Series newSeries123 = new Series(); chart5.Series.Add(newSeries123);
            Series newSeries124 = new Series(); chart5.Series.Add(newSeries124);
            Series newSeries125 = new Series(); chart5.Series.Add(newSeries125);



            Title title5 = chart5.Titles.Add("UL User Throughput (Mbps)");
            title5.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            chart5.Legends["Legend1"].Docking = Docking.Bottom;
            chart5.ChartAreas[0].AxisX.Interval = 5;
            chart5.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart5.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";

            for (int i = 0; i < RNC_List_Table.Rows.Count; i++)
            {
                chart5.Series[i].ChartType = SeriesChartType.Line;
                chart5.Series[i].BorderWidth = 3;
                chart5.Series[i].EmptyPointStyle.Color = Color.Transparent;
                chart5.Series[i].XValueType = ChartValueType.DateTime;
                chart5.Series[i].ToolTip = Series_List[i, 0] + "_" + "#VALX [#VALY]";
                chart5.Series[i].IsValueShownAsLabel = false;
            }


            chart6.Series.Clear();
            chart6.Titles.Clear();

            Series newSeries126 = new Series(); chart6.Series.Add(newSeries126);
            Series newSeries127 = new Series(); chart6.Series.Add(newSeries127);
            Series newSeries128 = new Series(); chart6.Series.Add(newSeries128);
            Series newSeries129 = new Series(); chart6.Series.Add(newSeries129);
            Series newSeries130 = new Series(); chart6.Series.Add(newSeries130);
            Series newSeries131 = new Series(); chart6.Series.Add(newSeries131);
            Series newSeries132 = new Series(); chart6.Series.Add(newSeries132);
            Series newSeries133 = new Series(); chart6.Series.Add(newSeries133);
            Series newSeries134 = new Series(); chart6.Series.Add(newSeries134);
            Series newSeries135 = new Series(); chart6.Series.Add(newSeries135);
            Series newSeries136 = new Series(); chart6.Series.Add(newSeries136);
            Series newSeries137 = new Series(); chart6.Series.Add(newSeries137);
            Series newSeries138 = new Series(); chart6.Series.Add(newSeries138);
            Series newSeries139 = new Series(); chart6.Series.Add(newSeries139);
            Series newSeries140 = new Series(); chart6.Series.Add(newSeries140);
            Series newSeries141 = new Series(); chart6.Series.Add(newSeries141);
            Series newSeries142 = new Series(); chart6.Series.Add(newSeries142);
            Series newSeries143 = new Series(); chart6.Series.Add(newSeries143);
            Series newSeries144 = new Series(); chart6.Series.Add(newSeries144);
            Series newSeries145 = new Series(); chart6.Series.Add(newSeries145);
            Series newSeries146 = new Series(); chart6.Series.Add(newSeries146);
            Series newSeries147 = new Series(); chart6.Series.Add(newSeries147);
            Series newSeries148 = new Series(); chart6.Series.Add(newSeries148);
            Series newSeries149 = new Series(); chart6.Series.Add(newSeries149);
            Series newSeries150 = new Series(); chart6.Series.Add(newSeries150);


            Title title6 = chart6.Titles.Add("DL Latency (ms)");
            title6.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            chart6.Legends["Legend1"].Docking = Docking.Bottom;
            chart6.ChartAreas[0].AxisX.Interval = 5;
            chart6.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart6.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";

            for (int i = 0; i < RNC_List_Table.Rows.Count; i++)
            {
                chart6.Series[i].ChartType = SeriesChartType.Line;
                chart6.Series[i].BorderWidth = 3;
                chart6.Series[i].EmptyPointStyle.Color = Color.Transparent;
                chart6.Series[i].XValueType = ChartValueType.DateTime;
                chart6.Series[i].ToolTip = Series_List[i, 0] + "_" + "#VALX [#VALY]";
                chart6.Series[i].IsValueShownAsLabel = false;
            }


            chart7.Series.Clear();
            chart7.Titles.Clear();
            Series newSeries151 = new Series(); chart7.Series.Add(newSeries151);
            Series newSeries152 = new Series(); chart7.Series.Add(newSeries152);
            Series newSeries153 = new Series(); chart7.Series.Add(newSeries153);
            Series newSeries154 = new Series(); chart7.Series.Add(newSeries154);
            Series newSeries155 = new Series(); chart7.Series.Add(newSeries155);
            Series newSeries156 = new Series(); chart7.Series.Add(newSeries156);
            Series newSeries157 = new Series(); chart7.Series.Add(newSeries157);
            Series newSeries158 = new Series(); chart7.Series.Add(newSeries158);
            Series newSeries159 = new Series(); chart7.Series.Add(newSeries159);
            Series newSeries160 = new Series(); chart7.Series.Add(newSeries160);
            Series newSeries161 = new Series(); chart7.Series.Add(newSeries161);
            Series newSeries162 = new Series(); chart7.Series.Add(newSeries162);
            Series newSeries163 = new Series(); chart7.Series.Add(newSeries163);
            Series newSeries164 = new Series(); chart7.Series.Add(newSeries164);
            Series newSeries165 = new Series(); chart7.Series.Add(newSeries165);
            Series newSeries166 = new Series(); chart7.Series.Add(newSeries166);
            Series newSeries167 = new Series(); chart7.Series.Add(newSeries167);
            Series newSeries168 = new Series(); chart7.Series.Add(newSeries168);
            Series newSeries169 = new Series(); chart7.Series.Add(newSeries169);
            Series newSeries170 = new Series(); chart7.Series.Add(newSeries170);
            Series newSeries171 = new Series(); chart7.Series.Add(newSeries171);
            Series newSeries172 = new Series(); chart7.Series.Add(newSeries172);
            Series newSeries173 = new Series(); chart7.Series.Add(newSeries173);
            Series newSeries174 = new Series(); chart7.Series.Add(newSeries174);
            Series newSeries175 = new Series(); chart7.Series.Add(newSeries175);


            Title title7 = chart7.Titles.Add("LTE Service SR");
            title7.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            chart7.Legends["Legend1"].Docking = Docking.Bottom;
            chart7.ChartAreas[0].AxisX.Interval = 5;
            chart7.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart7.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";

            for (int i = 0; i < RNC_List_Table.Rows.Count; i++)
            {
                chart7.Series[i].ChartType = SeriesChartType.Line;
                chart7.Series[i].BorderWidth = 3;
                chart7.Series[i].EmptyPointStyle.Color = Color.Transparent;
                chart7.Series[i].XValueType = ChartValueType.DateTime;
                chart7.Series[i].ToolTip = Series_List[i, 0] + "_" + "#VALX [#VALY]";
                chart7.Series[i].IsValueShownAsLabel = false;
            }



            chart8.Series.Clear();
            chart8.Titles.Clear();
            Series newSeries176 = new Series(); chart8.Series.Add(newSeries176);
            Series newSeries177 = new Series(); chart8.Series.Add(newSeries177);
            Series newSeries178 = new Series(); chart8.Series.Add(newSeries178);
            Series newSeries179 = new Series(); chart8.Series.Add(newSeries179);
            Series newSeries180 = new Series(); chart8.Series.Add(newSeries180);
            Series newSeries181 = new Series(); chart8.Series.Add(newSeries181);
            Series newSeries182 = new Series(); chart8.Series.Add(newSeries182);
            Series newSeries183 = new Series(); chart8.Series.Add(newSeries183);
            Series newSeries184 = new Series(); chart8.Series.Add(newSeries184);
            Series newSeries185 = new Series(); chart8.Series.Add(newSeries185);
            Series newSeries186 = new Series(); chart8.Series.Add(newSeries186);
            Series newSeries187 = new Series(); chart8.Series.Add(newSeries187);
            Series newSeries188 = new Series(); chart8.Series.Add(newSeries188);
            Series newSeries189 = new Series(); chart8.Series.Add(newSeries189);
            Series newSeries190 = new Series(); chart8.Series.Add(newSeries190);
            Series newSeries191 = new Series(); chart8.Series.Add(newSeries191);
            Series newSeries192 = new Series(); chart8.Series.Add(newSeries192);
            Series newSeries193 = new Series(); chart8.Series.Add(newSeries193);
            Series newSeries194 = new Series(); chart8.Series.Add(newSeries194);
            Series newSeries195 = new Series(); chart8.Series.Add(newSeries195);
            Series newSeries196 = new Series(); chart8.Series.Add(newSeries196);
            Series newSeries197 = new Series(); chart8.Series.Add(newSeries197);
            Series newSeries198 = new Series(); chart8.Series.Add(newSeries198);
            Series newSeries199 = new Series(); chart8.Series.Add(newSeries199);
            Series newSeries200 = new Series(); chart8.Series.Add(newSeries200);



            Title title8 = chart8.Titles.Add("RRC Connection SR");
            title8.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            chart8.Legends["Legend1"].Docking = Docking.Bottom;
            chart8.ChartAreas[0].AxisX.Interval = 5;
            chart8.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart8.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";

            for (int i = 0; i < RNC_List_Table.Rows.Count; i++)
            {
                chart8.Series[i].ChartType = SeriesChartType.Line;
                chart8.Series[i].BorderWidth = 3;
                chart8.Series[i].EmptyPointStyle.Color = Color.Transparent;
                chart8.Series[i].XValueType = ChartValueType.DateTime;
                chart8.Series[i].ToolTip = Series_List[i, 0] + "_" + "#VALX [#VALY]";
                chart8.Series[i].IsValueShownAsLabel = false;
            }



            chart9.Series.Clear();
            chart9.Titles.Clear();
            Series newSeries201 = new Series(); chart9.Series.Add(newSeries201);
            Series newSeries202 = new Series(); chart9.Series.Add(newSeries202);
            Series newSeries203 = new Series(); chart9.Series.Add(newSeries203);
            Series newSeries204 = new Series(); chart9.Series.Add(newSeries204);
            Series newSeries205 = new Series(); chart9.Series.Add(newSeries205);
            Series newSeries206 = new Series(); chart9.Series.Add(newSeries206);
            Series newSeries207 = new Series(); chart9.Series.Add(newSeries207);
            Series newSeries208 = new Series(); chart9.Series.Add(newSeries208);
            Series newSeries209 = new Series(); chart9.Series.Add(newSeries209);
            Series newSeries210 = new Series(); chart9.Series.Add(newSeries210);
            Series newSeries211 = new Series(); chart9.Series.Add(newSeries211);
            Series newSeries212 = new Series(); chart9.Series.Add(newSeries212);
            Series newSeries213 = new Series(); chart9.Series.Add(newSeries213);
            Series newSeries214 = new Series(); chart9.Series.Add(newSeries214);
            Series newSeries215 = new Series(); chart9.Series.Add(newSeries215);
            Series newSeries216 = new Series(); chart9.Series.Add(newSeries216);
            Series newSeries217 = new Series(); chart9.Series.Add(newSeries217);
            Series newSeries218 = new Series(); chart9.Series.Add(newSeries218);
            Series newSeries219 = new Series(); chart9.Series.Add(newSeries219);
            Series newSeries220 = new Series(); chart9.Series.Add(newSeries220);
            Series newSeries221 = new Series(); chart9.Series.Add(newSeries221);
            Series newSeries222 = new Series(); chart9.Series.Add(newSeries222);
            Series newSeries223 = new Series(); chart9.Series.Add(newSeries223);
            Series newSeries224 = new Series(); chart9.Series.Add(newSeries224);
            Series newSeries225 = new Series(); chart9.Series.Add(newSeries225);




            Title title9 = chart9.Titles.Add("ERAB Setup SR");
            title9.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            chart9.Legends["Legend1"].Docking = Docking.Bottom;
            chart9.ChartAreas[0].AxisX.Interval = 5;
            chart9.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart9.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";

            for (int i = 0; i < RNC_List_Table.Rows.Count; i++)
            {
                chart9.Series[i].ChartType = SeriesChartType.Line;
                chart9.Series[i].BorderWidth = 3;
                chart9.Series[i].EmptyPointStyle.Color = Color.Transparent;
                chart9.Series[i].XValueType = ChartValueType.DateTime;
                chart9.Series[i].ToolTip = Series_List[i, 0] + "_" + "#VALX [#VALY]";
                chart9.Series[i].IsValueShownAsLabel = false;
            }



            chart10.Series.Clear();
            chart10.Titles.Clear();
            Series newSeries226 = new Series(); chart10.Series.Add(newSeries226);
            Series newSeries227 = new Series(); chart10.Series.Add(newSeries227);
            Series newSeries228 = new Series(); chart10.Series.Add(newSeries228);
            Series newSeries229 = new Series(); chart10.Series.Add(newSeries229);
            Series newSeries230 = new Series(); chart10.Series.Add(newSeries230);
            Series newSeries231 = new Series(); chart10.Series.Add(newSeries231);
            Series newSeries232 = new Series(); chart10.Series.Add(newSeries232);
            Series newSeries233 = new Series(); chart10.Series.Add(newSeries233);
            Series newSeries234 = new Series(); chart10.Series.Add(newSeries234);
            Series newSeries235 = new Series(); chart10.Series.Add(newSeries235);
            Series newSeries236 = new Series(); chart10.Series.Add(newSeries236);
            Series newSeries237 = new Series(); chart10.Series.Add(newSeries237);
            Series newSeries238 = new Series(); chart10.Series.Add(newSeries238);
            Series newSeries239 = new Series(); chart10.Series.Add(newSeries239);
            Series newSeries240 = new Series(); chart10.Series.Add(newSeries240);
            Series newSeries241 = new Series(); chart10.Series.Add(newSeries241);
            Series newSeries242 = new Series(); chart10.Series.Add(newSeries242);
            Series newSeries243 = new Series(); chart10.Series.Add(newSeries243);
            Series newSeries244 = new Series(); chart10.Series.Add(newSeries244);
            Series newSeries245 = new Series(); chart10.Series.Add(newSeries245);
            Series newSeries246 = new Series(); chart10.Series.Add(newSeries246);
            Series newSeries247 = new Series(); chart10.Series.Add(newSeries247);
            Series newSeries248 = new Series(); chart10.Series.Add(newSeries248);
            Series newSeries249 = new Series(); chart10.Series.Add(newSeries249);
            Series newSeries250 = new Series(); chart10.Series.Add(newSeries250);


            Title title10 = chart10.Titles.Add("ERAB Drop Rate");
            title10.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            chart10.Legends["Legend1"].Docking = Docking.Bottom;
            chart10.ChartAreas[0].AxisX.Interval = 5;
            chart10.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart10.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";

            for (int i = 0; i < RNC_List_Table.Rows.Count; i++)
            {
                chart10.Series[i].ChartType = SeriesChartType.Line;
                chart10.Series[i].BorderWidth = 3;
                chart10.Series[i].EmptyPointStyle.Color = Color.Transparent;
                chart10.Series[i].XValueType = ChartValueType.DateTime;
                chart10.Series[i].ToolTip = Series_List[i, 0] + "_" + "#VALX [#VALY]";
                chart10.Series[i].IsValueShownAsLabel = false;
            }




            chart11.Series.Clear();
            chart11.Titles.Clear();
            Series newSeries251 = new Series(); chart11.Series.Add(newSeries251);
            Series newSeries252 = new Series(); chart11.Series.Add(newSeries252);
            Series newSeries253 = new Series(); chart11.Series.Add(newSeries253);
            Series newSeries254 = new Series(); chart11.Series.Add(newSeries254);
            Series newSeries255 = new Series(); chart11.Series.Add(newSeries255);
            Series newSeries256 = new Series(); chart11.Series.Add(newSeries256);
            Series newSeries257 = new Series(); chart11.Series.Add(newSeries257);
            Series newSeries258 = new Series(); chart11.Series.Add(newSeries258);
            Series newSeries259 = new Series(); chart11.Series.Add(newSeries259);
            Series newSeries260 = new Series(); chart11.Series.Add(newSeries260);
            Series newSeries261 = new Series(); chart11.Series.Add(newSeries261);
            Series newSeries262 = new Series(); chart11.Series.Add(newSeries262);
            Series newSeries263 = new Series(); chart11.Series.Add(newSeries263);
            Series newSeries264 = new Series(); chart11.Series.Add(newSeries264);
            Series newSeries265 = new Series(); chart11.Series.Add(newSeries265);
            Series newSeries266 = new Series(); chart11.Series.Add(newSeries266);
            Series newSeries267 = new Series(); chart11.Series.Add(newSeries267);
            Series newSeries268 = new Series(); chart11.Series.Add(newSeries268);
            Series newSeries269 = new Series(); chart11.Series.Add(newSeries269);
            Series newSeries270 = new Series(); chart11.Series.Add(newSeries270);
            Series newSeries271 = new Series(); chart11.Series.Add(newSeries271);
            Series newSeries272 = new Series(); chart11.Series.Add(newSeries272);
            Series newSeries273 = new Series(); chart11.Series.Add(newSeries273);
            Series newSeries274 = new Series(); chart11.Series.Add(newSeries274);
            Series newSeries275 = new Series(); chart11.Series.Add(newSeries275);


            Title title11 = chart11.Titles.Add("Inter Freq HO SR");
            title11.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            chart11.Legends["Legend1"].Docking = Docking.Bottom;
            chart11.ChartAreas[0].AxisX.Interval = 5;
            chart11.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart11.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";

            for (int i = 0; i < RNC_List_Table.Rows.Count; i++)
            {
                chart11.Series[i].ChartType = SeriesChartType.Line;
                chart11.Series[i].BorderWidth = 3;
                chart11.Series[i].EmptyPointStyle.Color = Color.Transparent;
                chart11.Series[i].XValueType = ChartValueType.DateTime;
                chart11.Series[i].ToolTip = Series_List[i, 0] + "_" + "#VALX [#VALY]";
                chart11.Series[i].IsValueShownAsLabel = false;
            }



            chart12.Series.Clear();
            chart12.Titles.Clear();
            Series newSeries276 = new Series(); chart12.Series.Add(newSeries276);
            Series newSeries277 = new Series(); chart12.Series.Add(newSeries277);
            Series newSeries278 = new Series(); chart12.Series.Add(newSeries278);
            Series newSeries279 = new Series(); chart12.Series.Add(newSeries279);
            Series newSeries280 = new Series(); chart12.Series.Add(newSeries280);
            Series newSeries281 = new Series(); chart12.Series.Add(newSeries281);
            Series newSeries282 = new Series(); chart12.Series.Add(newSeries282);
            Series newSeries283 = new Series(); chart12.Series.Add(newSeries283);
            Series newSeries284 = new Series(); chart12.Series.Add(newSeries284);
            Series newSeries285 = new Series(); chart12.Series.Add(newSeries285);
            Series newSeries286 = new Series(); chart12.Series.Add(newSeries286);
            Series newSeries287 = new Series(); chart12.Series.Add(newSeries287);
            Series newSeries288 = new Series(); chart12.Series.Add(newSeries288);
            Series newSeries289 = new Series(); chart12.Series.Add(newSeries289);
            Series newSeries290 = new Series(); chart12.Series.Add(newSeries290);
            Series newSeries291 = new Series(); chart12.Series.Add(newSeries291);
            Series newSeries292 = new Series(); chart12.Series.Add(newSeries292);
            Series newSeries293 = new Series(); chart12.Series.Add(newSeries293);
            Series newSeries294 = new Series(); chart12.Series.Add(newSeries294);
            Series newSeries295 = new Series(); chart12.Series.Add(newSeries295);
            Series newSeries296 = new Series(); chart12.Series.Add(newSeries296);
            Series newSeries297 = new Series(); chart12.Series.Add(newSeries297);
            Series newSeries298 = new Series(); chart12.Series.Add(newSeries298);
            Series newSeries299 = new Series(); chart12.Series.Add(newSeries299);
            Series newSeries300 = new Series(); chart12.Series.Add(newSeries300);


            Title title12 = chart12.Titles.Add("Intra Freq HO SR");
            title12.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            chart12.Legends["Legend1"].Docking = Docking.Bottom;
            chart12.ChartAreas[0].AxisX.Interval = 5;
            chart12.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart12.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";

            for (int i = 0; i < RNC_List_Table.Rows.Count; i++)
            {
                chart12.Series[i].ChartType = SeriesChartType.Line;
                chart12.Series[i].BorderWidth = 3;
                chart12.Series[i].EmptyPointStyle.Color = Color.Transparent;
                chart12.Series[i].XValueType = ChartValueType.DateTime;
                chart12.Series[i].ToolTip = Series_List[i, 0] + "_" + "#VALX [#VALY]";
                chart12.Series[i].IsValueShownAsLabel = false;
            }



            chart13.Series.Clear();
            chart13.Titles.Clear();
            Series newSeries301 = new Series(); chart13.Series.Add(newSeries301);
            Series newSeries302 = new Series(); chart13.Series.Add(newSeries302);
            Series newSeries303 = new Series(); chart13.Series.Add(newSeries303);
            Series newSeries304 = new Series(); chart13.Series.Add(newSeries304);
            Series newSeries305 = new Series(); chart13.Series.Add(newSeries305);
            Series newSeries306 = new Series(); chart13.Series.Add(newSeries306);
            Series newSeries307 = new Series(); chart13.Series.Add(newSeries307);
            Series newSeries308 = new Series(); chart13.Series.Add(newSeries308);
            Series newSeries309 = new Series(); chart13.Series.Add(newSeries309);
            Series newSeries310 = new Series(); chart13.Series.Add(newSeries310);
            Series newSeries311 = new Series(); chart13.Series.Add(newSeries311);
            Series newSeries312 = new Series(); chart13.Series.Add(newSeries312);
            Series newSeries313 = new Series(); chart13.Series.Add(newSeries313);
            Series newSeries314 = new Series(); chart13.Series.Add(newSeries314);
            Series newSeries315 = new Series(); chart13.Series.Add(newSeries315);
            Series newSeries316 = new Series(); chart13.Series.Add(newSeries316);
            Series newSeries317 = new Series(); chart13.Series.Add(newSeries317);
            Series newSeries318 = new Series(); chart13.Series.Add(newSeries318);
            Series newSeries319 = new Series(); chart13.Series.Add(newSeries319);
            Series newSeries320 = new Series(); chart13.Series.Add(newSeries320);
            Series newSeries321 = new Series(); chart13.Series.Add(newSeries321);
            Series newSeries322 = new Series(); chart13.Series.Add(newSeries322);
            Series newSeries323 = new Series(); chart13.Series.Add(newSeries323);
            Series newSeries324 = new Series(); chart13.Series.Add(newSeries324);
            Series newSeries325 = new Series(); chart13.Series.Add(newSeries325);




            Title title13 = chart13.Titles.Add("PUCCH RSSI (dBm)");
            title13.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            chart13.Legends["Legend1"].Docking = Docking.Bottom;
            chart13.ChartAreas[0].AxisX.Interval = 5;
            chart13.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart13.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";

            for (int i = 0; i < RNC_List_Table.Rows.Count; i++)
            {
                chart13.Series[i].ChartType = SeriesChartType.Line;
                chart13.Series[i].BorderWidth = 3;
                chart13.Series[i].EmptyPointStyle.Color = Color.Transparent;
                chart13.Series[i].XValueType = ChartValueType.DateTime;
                chart13.Series[i].ToolTip = Series_List[i, 0] + "_" + "#VALX [#VALY]";
                chart13.Series[i].IsValueShownAsLabel = false;
            }



            chart14.Series.Clear();
            chart14.Titles.Clear();
            Series newSeries326 = new Series(); chart14.Series.Add(newSeries326);
            Series newSeries327 = new Series(); chart14.Series.Add(newSeries327);
            Series newSeries328 = new Series(); chart14.Series.Add(newSeries328);
            Series newSeries329 = new Series(); chart14.Series.Add(newSeries329);
            Series newSeries330 = new Series(); chart14.Series.Add(newSeries330);
            Series newSeries331 = new Series(); chart14.Series.Add(newSeries331);
            Series newSeries332 = new Series(); chart14.Series.Add(newSeries332);
            Series newSeries333 = new Series(); chart14.Series.Add(newSeries333);
            Series newSeries334 = new Series(); chart14.Series.Add(newSeries334);
            Series newSeries335 = new Series(); chart14.Series.Add(newSeries335);
            Series newSeries336 = new Series(); chart14.Series.Add(newSeries336);
            Series newSeries337 = new Series(); chart14.Series.Add(newSeries337);
            Series newSeries338 = new Series(); chart14.Series.Add(newSeries338);
            Series newSeries339 = new Series(); chart14.Series.Add(newSeries339);
            Series newSeries340 = new Series(); chart14.Series.Add(newSeries340);
            Series newSeries341 = new Series(); chart14.Series.Add(newSeries341);
            Series newSeries342 = new Series(); chart14.Series.Add(newSeries342);
            Series newSeries343 = new Series(); chart14.Series.Add(newSeries343);
            Series newSeries344 = new Series(); chart14.Series.Add(newSeries344);
            Series newSeries345 = new Series(); chart14.Series.Add(newSeries345);
            Series newSeries346 = new Series(); chart14.Series.Add(newSeries346);
            Series newSeries347 = new Series(); chart14.Series.Add(newSeries347);
            Series newSeries348 = new Series(); chart14.Series.Add(newSeries348);
            Series newSeries349 = new Series(); chart14.Series.Add(newSeries349);
            Series newSeries350 = new Series(); chart14.Series.Add(newSeries350);


            Title title14 = chart14.Titles.Add("PUSCH RSSI (dBm)");
            title14.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            chart14.Legends["Legend1"].Docking = Docking.Bottom;
            chart14.ChartAreas[0].AxisX.Interval = 5;
            chart14.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart14.ChartAreas[0].AxisX.LabelStyle.Format = "dd/MM";





            newSeries1.MarkerStyle = MarkerStyle.Circle;
            newSeries2.MarkerStyle = MarkerStyle.Circle;
            newSeries3.MarkerStyle = MarkerStyle.Circle;
            newSeries4.MarkerStyle = MarkerStyle.Circle;
            newSeries5.MarkerStyle = MarkerStyle.Circle;
            newSeries6.MarkerStyle = MarkerStyle.Circle;
            newSeries7.MarkerStyle = MarkerStyle.Circle;
            newSeries8.MarkerStyle = MarkerStyle.Circle;
            newSeries9.MarkerStyle = MarkerStyle.Circle;
            newSeries10.MarkerStyle = MarkerStyle.Circle;
            newSeries11.MarkerStyle = MarkerStyle.Circle;
            newSeries12.MarkerStyle = MarkerStyle.Circle;
            newSeries13.MarkerStyle = MarkerStyle.Circle;
            newSeries14.MarkerStyle = MarkerStyle.Circle;
            newSeries15.MarkerStyle = MarkerStyle.Circle;
            newSeries16.MarkerStyle = MarkerStyle.Circle;
            newSeries17.MarkerStyle = MarkerStyle.Circle;
            newSeries18.MarkerStyle = MarkerStyle.Circle;
            newSeries19.MarkerStyle = MarkerStyle.Circle;
            newSeries20.MarkerStyle = MarkerStyle.Circle;
            newSeries21.MarkerStyle = MarkerStyle.Circle;
            newSeries22.MarkerStyle = MarkerStyle.Circle;
            newSeries23.MarkerStyle = MarkerStyle.Circle;
            newSeries24.MarkerStyle = MarkerStyle.Circle;
            newSeries25.MarkerStyle = MarkerStyle.Circle;
            newSeries26.MarkerStyle = MarkerStyle.Circle;
            newSeries27.MarkerStyle = MarkerStyle.Circle;
            newSeries28.MarkerStyle = MarkerStyle.Circle;
            newSeries29.MarkerStyle = MarkerStyle.Circle;
            newSeries30.MarkerStyle = MarkerStyle.Circle;
            newSeries31.MarkerStyle = MarkerStyle.Circle;
            newSeries32.MarkerStyle = MarkerStyle.Circle;
            newSeries33.MarkerStyle = MarkerStyle.Circle;
            newSeries34.MarkerStyle = MarkerStyle.Circle;
            newSeries35.MarkerStyle = MarkerStyle.Circle;
            newSeries36.MarkerStyle = MarkerStyle.Circle;
            newSeries37.MarkerStyle = MarkerStyle.Circle;
            newSeries38.MarkerStyle = MarkerStyle.Circle;
            newSeries39.MarkerStyle = MarkerStyle.Circle;
            newSeries40.MarkerStyle = MarkerStyle.Circle;
            newSeries41.MarkerStyle = MarkerStyle.Circle;
            newSeries42.MarkerStyle = MarkerStyle.Circle;
            newSeries43.MarkerStyle = MarkerStyle.Circle;
            newSeries44.MarkerStyle = MarkerStyle.Circle;
            newSeries45.MarkerStyle = MarkerStyle.Circle;
            newSeries46.MarkerStyle = MarkerStyle.Circle;
            newSeries47.MarkerStyle = MarkerStyle.Circle;
            newSeries48.MarkerStyle = MarkerStyle.Circle;
            newSeries49.MarkerStyle = MarkerStyle.Circle;
            newSeries50.MarkerStyle = MarkerStyle.Circle;
            newSeries51.MarkerStyle = MarkerStyle.Circle;
            newSeries52.MarkerStyle = MarkerStyle.Circle;
            newSeries53.MarkerStyle = MarkerStyle.Circle;
            newSeries54.MarkerStyle = MarkerStyle.Circle;
            newSeries55.MarkerStyle = MarkerStyle.Circle;
            newSeries56.MarkerStyle = MarkerStyle.Circle;
            newSeries57.MarkerStyle = MarkerStyle.Circle;
            newSeries58.MarkerStyle = MarkerStyle.Circle;
            newSeries59.MarkerStyle = MarkerStyle.Circle;
            newSeries60.MarkerStyle = MarkerStyle.Circle;
            newSeries61.MarkerStyle = MarkerStyle.Circle;
            newSeries62.MarkerStyle = MarkerStyle.Circle;
            newSeries63.MarkerStyle = MarkerStyle.Circle;
            newSeries64.MarkerStyle = MarkerStyle.Circle;
            newSeries65.MarkerStyle = MarkerStyle.Circle;
            newSeries66.MarkerStyle = MarkerStyle.Circle;
            newSeries67.MarkerStyle = MarkerStyle.Circle;
            newSeries68.MarkerStyle = MarkerStyle.Circle;
            newSeries69.MarkerStyle = MarkerStyle.Circle;
            newSeries70.MarkerStyle = MarkerStyle.Circle;
            newSeries71.MarkerStyle = MarkerStyle.Circle;
            newSeries72.MarkerStyle = MarkerStyle.Circle;
            newSeries73.MarkerStyle = MarkerStyle.Circle;
            newSeries74.MarkerStyle = MarkerStyle.Circle;
            newSeries75.MarkerStyle = MarkerStyle.Circle;
            newSeries76.MarkerStyle = MarkerStyle.Circle;
            newSeries77.MarkerStyle = MarkerStyle.Circle;
            newSeries78.MarkerStyle = MarkerStyle.Circle;
            newSeries79.MarkerStyle = MarkerStyle.Circle;
            newSeries80.MarkerStyle = MarkerStyle.Circle;
            newSeries81.MarkerStyle = MarkerStyle.Circle;
            newSeries82.MarkerStyle = MarkerStyle.Circle;
            newSeries83.MarkerStyle = MarkerStyle.Circle;
            newSeries84.MarkerStyle = MarkerStyle.Circle;
            newSeries85.MarkerStyle = MarkerStyle.Circle;
            newSeries86.MarkerStyle = MarkerStyle.Circle;
            newSeries87.MarkerStyle = MarkerStyle.Circle;
            newSeries88.MarkerStyle = MarkerStyle.Circle;
            newSeries89.MarkerStyle = MarkerStyle.Circle;
            newSeries90.MarkerStyle = MarkerStyle.Circle;
            newSeries91.MarkerStyle = MarkerStyle.Circle;
            newSeries92.MarkerStyle = MarkerStyle.Circle;
            newSeries93.MarkerStyle = MarkerStyle.Circle;
            newSeries94.MarkerStyle = MarkerStyle.Circle;
            newSeries95.MarkerStyle = MarkerStyle.Circle;
            newSeries96.MarkerStyle = MarkerStyle.Circle;
            newSeries97.MarkerStyle = MarkerStyle.Circle;
            newSeries98.MarkerStyle = MarkerStyle.Circle;
            newSeries99.MarkerStyle = MarkerStyle.Circle;
            newSeries100.MarkerStyle = MarkerStyle.Circle;
            newSeries101.MarkerStyle = MarkerStyle.Circle;
            newSeries102.MarkerStyle = MarkerStyle.Circle;
            newSeries103.MarkerStyle = MarkerStyle.Circle;
            newSeries104.MarkerStyle = MarkerStyle.Circle;
            newSeries105.MarkerStyle = MarkerStyle.Circle;
            newSeries106.MarkerStyle = MarkerStyle.Circle;
            newSeries107.MarkerStyle = MarkerStyle.Circle;
            newSeries108.MarkerStyle = MarkerStyle.Circle;
            newSeries109.MarkerStyle = MarkerStyle.Circle;
            newSeries110.MarkerStyle = MarkerStyle.Circle;
            newSeries111.MarkerStyle = MarkerStyle.Circle;
            newSeries112.MarkerStyle = MarkerStyle.Circle;
            newSeries113.MarkerStyle = MarkerStyle.Circle;
            newSeries114.MarkerStyle = MarkerStyle.Circle;
            newSeries115.MarkerStyle = MarkerStyle.Circle;
            newSeries116.MarkerStyle = MarkerStyle.Circle;
            newSeries117.MarkerStyle = MarkerStyle.Circle;
            newSeries118.MarkerStyle = MarkerStyle.Circle;
            newSeries119.MarkerStyle = MarkerStyle.Circle;
            newSeries120.MarkerStyle = MarkerStyle.Circle;
            newSeries121.MarkerStyle = MarkerStyle.Circle;
            newSeries122.MarkerStyle = MarkerStyle.Circle;
            newSeries123.MarkerStyle = MarkerStyle.Circle;
            newSeries124.MarkerStyle = MarkerStyle.Circle;
            newSeries125.MarkerStyle = MarkerStyle.Circle;
            newSeries126.MarkerStyle = MarkerStyle.Circle;
            newSeries127.MarkerStyle = MarkerStyle.Circle;
            newSeries128.MarkerStyle = MarkerStyle.Circle;
            newSeries129.MarkerStyle = MarkerStyle.Circle;
            newSeries130.MarkerStyle = MarkerStyle.Circle;
            newSeries131.MarkerStyle = MarkerStyle.Circle;
            newSeries132.MarkerStyle = MarkerStyle.Circle;
            newSeries133.MarkerStyle = MarkerStyle.Circle;
            newSeries134.MarkerStyle = MarkerStyle.Circle;
            newSeries135.MarkerStyle = MarkerStyle.Circle;
            newSeries136.MarkerStyle = MarkerStyle.Circle;
            newSeries137.MarkerStyle = MarkerStyle.Circle;
            newSeries138.MarkerStyle = MarkerStyle.Circle;
            newSeries139.MarkerStyle = MarkerStyle.Circle;
            newSeries140.MarkerStyle = MarkerStyle.Circle;
            newSeries141.MarkerStyle = MarkerStyle.Circle;
            newSeries142.MarkerStyle = MarkerStyle.Circle;
            newSeries143.MarkerStyle = MarkerStyle.Circle;
            newSeries144.MarkerStyle = MarkerStyle.Circle;
            newSeries145.MarkerStyle = MarkerStyle.Circle;
            newSeries146.MarkerStyle = MarkerStyle.Circle;
            newSeries147.MarkerStyle = MarkerStyle.Circle;
            newSeries148.MarkerStyle = MarkerStyle.Circle;
            newSeries149.MarkerStyle = MarkerStyle.Circle;
            newSeries150.MarkerStyle = MarkerStyle.Circle;
            newSeries151.MarkerStyle = MarkerStyle.Circle;
            newSeries152.MarkerStyle = MarkerStyle.Circle;
            newSeries153.MarkerStyle = MarkerStyle.Circle;
            newSeries154.MarkerStyle = MarkerStyle.Circle;
            newSeries155.MarkerStyle = MarkerStyle.Circle;
            newSeries156.MarkerStyle = MarkerStyle.Circle;
            newSeries157.MarkerStyle = MarkerStyle.Circle;
            newSeries158.MarkerStyle = MarkerStyle.Circle;
            newSeries159.MarkerStyle = MarkerStyle.Circle;
            newSeries160.MarkerStyle = MarkerStyle.Circle;
            newSeries161.MarkerStyle = MarkerStyle.Circle;
            newSeries162.MarkerStyle = MarkerStyle.Circle;
            newSeries163.MarkerStyle = MarkerStyle.Circle;
            newSeries164.MarkerStyle = MarkerStyle.Circle;
            newSeries165.MarkerStyle = MarkerStyle.Circle;
            newSeries166.MarkerStyle = MarkerStyle.Circle;
            newSeries167.MarkerStyle = MarkerStyle.Circle;
            newSeries168.MarkerStyle = MarkerStyle.Circle;
            newSeries169.MarkerStyle = MarkerStyle.Circle;
            newSeries170.MarkerStyle = MarkerStyle.Circle;
            newSeries171.MarkerStyle = MarkerStyle.Circle;
            newSeries172.MarkerStyle = MarkerStyle.Circle;
            newSeries173.MarkerStyle = MarkerStyle.Circle;
            newSeries174.MarkerStyle = MarkerStyle.Circle;
            newSeries175.MarkerStyle = MarkerStyle.Circle;
            newSeries176.MarkerStyle = MarkerStyle.Circle;
            newSeries177.MarkerStyle = MarkerStyle.Circle;
            newSeries178.MarkerStyle = MarkerStyle.Circle;
            newSeries179.MarkerStyle = MarkerStyle.Circle;
            newSeries180.MarkerStyle = MarkerStyle.Circle;
            newSeries181.MarkerStyle = MarkerStyle.Circle;
            newSeries182.MarkerStyle = MarkerStyle.Circle;
            newSeries183.MarkerStyle = MarkerStyle.Circle;
            newSeries184.MarkerStyle = MarkerStyle.Circle;
            newSeries185.MarkerStyle = MarkerStyle.Circle;
            newSeries186.MarkerStyle = MarkerStyle.Circle;
            newSeries187.MarkerStyle = MarkerStyle.Circle;
            newSeries188.MarkerStyle = MarkerStyle.Circle;
            newSeries189.MarkerStyle = MarkerStyle.Circle;
            newSeries190.MarkerStyle = MarkerStyle.Circle;
            newSeries191.MarkerStyle = MarkerStyle.Circle;
            newSeries192.MarkerStyle = MarkerStyle.Circle;
            newSeries193.MarkerStyle = MarkerStyle.Circle;
            newSeries194.MarkerStyle = MarkerStyle.Circle;
            newSeries195.MarkerStyle = MarkerStyle.Circle;
            newSeries196.MarkerStyle = MarkerStyle.Circle;
            newSeries197.MarkerStyle = MarkerStyle.Circle;
            newSeries198.MarkerStyle = MarkerStyle.Circle;
            newSeries199.MarkerStyle = MarkerStyle.Circle;
            newSeries200.MarkerStyle = MarkerStyle.Circle;
            newSeries201.MarkerStyle = MarkerStyle.Circle;
            newSeries202.MarkerStyle = MarkerStyle.Circle;
            newSeries203.MarkerStyle = MarkerStyle.Circle;
            newSeries204.MarkerStyle = MarkerStyle.Circle;
            newSeries205.MarkerStyle = MarkerStyle.Circle;
            newSeries206.MarkerStyle = MarkerStyle.Circle;
            newSeries207.MarkerStyle = MarkerStyle.Circle;
            newSeries208.MarkerStyle = MarkerStyle.Circle;
            newSeries209.MarkerStyle = MarkerStyle.Circle;
            newSeries210.MarkerStyle = MarkerStyle.Circle;
            newSeries211.MarkerStyle = MarkerStyle.Circle;
            newSeries212.MarkerStyle = MarkerStyle.Circle;
            newSeries213.MarkerStyle = MarkerStyle.Circle;
            newSeries214.MarkerStyle = MarkerStyle.Circle;
            newSeries215.MarkerStyle = MarkerStyle.Circle;
            newSeries216.MarkerStyle = MarkerStyle.Circle;
            newSeries217.MarkerStyle = MarkerStyle.Circle;
            newSeries218.MarkerStyle = MarkerStyle.Circle;
            newSeries219.MarkerStyle = MarkerStyle.Circle;
            newSeries220.MarkerStyle = MarkerStyle.Circle;
            newSeries221.MarkerStyle = MarkerStyle.Circle;
            newSeries222.MarkerStyle = MarkerStyle.Circle;
            newSeries223.MarkerStyle = MarkerStyle.Circle;
            newSeries224.MarkerStyle = MarkerStyle.Circle;
            newSeries225.MarkerStyle = MarkerStyle.Circle;
            newSeries226.MarkerStyle = MarkerStyle.Circle;
            newSeries227.MarkerStyle = MarkerStyle.Circle;
            newSeries228.MarkerStyle = MarkerStyle.Circle;
            newSeries229.MarkerStyle = MarkerStyle.Circle;
            newSeries230.MarkerStyle = MarkerStyle.Circle;
            newSeries231.MarkerStyle = MarkerStyle.Circle;
            newSeries232.MarkerStyle = MarkerStyle.Circle;
            newSeries233.MarkerStyle = MarkerStyle.Circle;
            newSeries234.MarkerStyle = MarkerStyle.Circle;
            newSeries235.MarkerStyle = MarkerStyle.Circle;
            newSeries236.MarkerStyle = MarkerStyle.Circle;
            newSeries237.MarkerStyle = MarkerStyle.Circle;
            newSeries238.MarkerStyle = MarkerStyle.Circle;
            newSeries239.MarkerStyle = MarkerStyle.Circle;
            newSeries240.MarkerStyle = MarkerStyle.Circle;
            newSeries241.MarkerStyle = MarkerStyle.Circle;
            newSeries242.MarkerStyle = MarkerStyle.Circle;
            newSeries243.MarkerStyle = MarkerStyle.Circle;
            newSeries244.MarkerStyle = MarkerStyle.Circle;
            newSeries245.MarkerStyle = MarkerStyle.Circle;
            newSeries246.MarkerStyle = MarkerStyle.Circle;
            newSeries247.MarkerStyle = MarkerStyle.Circle;
            newSeries248.MarkerStyle = MarkerStyle.Circle;
            newSeries249.MarkerStyle = MarkerStyle.Circle;
            newSeries250.MarkerStyle = MarkerStyle.Circle;
            newSeries251.MarkerStyle = MarkerStyle.Circle;
            newSeries252.MarkerStyle = MarkerStyle.Circle;
            newSeries253.MarkerStyle = MarkerStyle.Circle;
            newSeries254.MarkerStyle = MarkerStyle.Circle;
            newSeries255.MarkerStyle = MarkerStyle.Circle;
            newSeries256.MarkerStyle = MarkerStyle.Circle;
            newSeries257.MarkerStyle = MarkerStyle.Circle;
            newSeries258.MarkerStyle = MarkerStyle.Circle;
            newSeries259.MarkerStyle = MarkerStyle.Circle;
            newSeries260.MarkerStyle = MarkerStyle.Circle;
            newSeries261.MarkerStyle = MarkerStyle.Circle;
            newSeries262.MarkerStyle = MarkerStyle.Circle;
            newSeries263.MarkerStyle = MarkerStyle.Circle;
            newSeries264.MarkerStyle = MarkerStyle.Circle;
            newSeries265.MarkerStyle = MarkerStyle.Circle;
            newSeries266.MarkerStyle = MarkerStyle.Circle;
            newSeries267.MarkerStyle = MarkerStyle.Circle;
            newSeries268.MarkerStyle = MarkerStyle.Circle;
            newSeries269.MarkerStyle = MarkerStyle.Circle;
            newSeries270.MarkerStyle = MarkerStyle.Circle;
            newSeries271.MarkerStyle = MarkerStyle.Circle;
            newSeries272.MarkerStyle = MarkerStyle.Circle;
            newSeries273.MarkerStyle = MarkerStyle.Circle;
            newSeries274.MarkerStyle = MarkerStyle.Circle;
            newSeries275.MarkerStyle = MarkerStyle.Circle;
            newSeries276.MarkerStyle = MarkerStyle.Circle;
            newSeries277.MarkerStyle = MarkerStyle.Circle;
            newSeries278.MarkerStyle = MarkerStyle.Circle;
            newSeries279.MarkerStyle = MarkerStyle.Circle;
            newSeries280.MarkerStyle = MarkerStyle.Circle;
            newSeries281.MarkerStyle = MarkerStyle.Circle;
            newSeries282.MarkerStyle = MarkerStyle.Circle;
            newSeries283.MarkerStyle = MarkerStyle.Circle;
            newSeries284.MarkerStyle = MarkerStyle.Circle;
            newSeries285.MarkerStyle = MarkerStyle.Circle;
            newSeries286.MarkerStyle = MarkerStyle.Circle;
            newSeries287.MarkerStyle = MarkerStyle.Circle;
            newSeries288.MarkerStyle = MarkerStyle.Circle;
            newSeries289.MarkerStyle = MarkerStyle.Circle;
            newSeries290.MarkerStyle = MarkerStyle.Circle;
            newSeries291.MarkerStyle = MarkerStyle.Circle;
            newSeries292.MarkerStyle = MarkerStyle.Circle;
            newSeries293.MarkerStyle = MarkerStyle.Circle;
            newSeries294.MarkerStyle = MarkerStyle.Circle;
            newSeries295.MarkerStyle = MarkerStyle.Circle;
            newSeries296.MarkerStyle = MarkerStyle.Circle;
            newSeries297.MarkerStyle = MarkerStyle.Circle;
            newSeries298.MarkerStyle = MarkerStyle.Circle;
            newSeries299.MarkerStyle = MarkerStyle.Circle;
            newSeries300.MarkerStyle = MarkerStyle.Circle;
            newSeries301.MarkerStyle = MarkerStyle.Circle;
            newSeries302.MarkerStyle = MarkerStyle.Circle;
            newSeries303.MarkerStyle = MarkerStyle.Circle;
            newSeries304.MarkerStyle = MarkerStyle.Circle;
            newSeries305.MarkerStyle = MarkerStyle.Circle;
            newSeries306.MarkerStyle = MarkerStyle.Circle;
            newSeries307.MarkerStyle = MarkerStyle.Circle;
            newSeries308.MarkerStyle = MarkerStyle.Circle;
            newSeries309.MarkerStyle = MarkerStyle.Circle;
            newSeries310.MarkerStyle = MarkerStyle.Circle;
            newSeries311.MarkerStyle = MarkerStyle.Circle;
            newSeries312.MarkerStyle = MarkerStyle.Circle;
            newSeries313.MarkerStyle = MarkerStyle.Circle;
            newSeries314.MarkerStyle = MarkerStyle.Circle;
            newSeries315.MarkerStyle = MarkerStyle.Circle;
            newSeries316.MarkerStyle = MarkerStyle.Circle;
            newSeries317.MarkerStyle = MarkerStyle.Circle;
            newSeries318.MarkerStyle = MarkerStyle.Circle;
            newSeries319.MarkerStyle = MarkerStyle.Circle;
            newSeries320.MarkerStyle = MarkerStyle.Circle;
            newSeries321.MarkerStyle = MarkerStyle.Circle;
            newSeries322.MarkerStyle = MarkerStyle.Circle;
            newSeries323.MarkerStyle = MarkerStyle.Circle;
            newSeries324.MarkerStyle = MarkerStyle.Circle;
            newSeries325.MarkerStyle = MarkerStyle.Circle;
            newSeries326.MarkerStyle = MarkerStyle.Circle;
            newSeries327.MarkerStyle = MarkerStyle.Circle;
            newSeries328.MarkerStyle = MarkerStyle.Circle;
            newSeries329.MarkerStyle = MarkerStyle.Circle;
            newSeries330.MarkerStyle = MarkerStyle.Circle;
            newSeries331.MarkerStyle = MarkerStyle.Circle;
            newSeries332.MarkerStyle = MarkerStyle.Circle;
            newSeries333.MarkerStyle = MarkerStyle.Circle;
            newSeries334.MarkerStyle = MarkerStyle.Circle;
            newSeries335.MarkerStyle = MarkerStyle.Circle;
            newSeries336.MarkerStyle = MarkerStyle.Circle;
            newSeries337.MarkerStyle = MarkerStyle.Circle;
            newSeries338.MarkerStyle = MarkerStyle.Circle;
            newSeries339.MarkerStyle = MarkerStyle.Circle;
            newSeries340.MarkerStyle = MarkerStyle.Circle;
            newSeries341.MarkerStyle = MarkerStyle.Circle;
            newSeries342.MarkerStyle = MarkerStyle.Circle;
            newSeries343.MarkerStyle = MarkerStyle.Circle;
            newSeries344.MarkerStyle = MarkerStyle.Circle;
            newSeries345.MarkerStyle = MarkerStyle.Circle;
            newSeries346.MarkerStyle = MarkerStyle.Circle;
            newSeries347.MarkerStyle = MarkerStyle.Circle;
            newSeries348.MarkerStyle = MarkerStyle.Circle;
            newSeries349.MarkerStyle = MarkerStyle.Circle;
            newSeries350.MarkerStyle = MarkerStyle.Circle;


            newSeries1.MarkerSize = 6;
            newSeries2.MarkerSize = 6;
            newSeries3.MarkerSize = 6;
            newSeries4.MarkerSize = 6;
            newSeries5.MarkerSize = 6;
            newSeries6.MarkerSize = 6;
            newSeries7.MarkerSize = 6;
            newSeries8.MarkerSize = 6;
            newSeries9.MarkerSize = 6;
            newSeries10.MarkerSize = 6;
            newSeries11.MarkerSize = 6;
            newSeries12.MarkerSize = 6;
            newSeries13.MarkerSize = 6;
            newSeries14.MarkerSize = 6;
            newSeries15.MarkerSize = 6;
            newSeries16.MarkerSize = 6;
            newSeries17.MarkerSize = 6;
            newSeries18.MarkerSize = 6;
            newSeries19.MarkerSize = 6;
            newSeries20.MarkerSize = 6;
            newSeries21.MarkerSize = 6;
            newSeries22.MarkerSize = 6;
            newSeries23.MarkerSize = 6;
            newSeries24.MarkerSize = 6;
            newSeries25.MarkerSize = 6;
            newSeries26.MarkerSize = 6;
            newSeries27.MarkerSize = 6;
            newSeries28.MarkerSize = 6;
            newSeries29.MarkerSize = 6;
            newSeries30.MarkerSize = 6;
            newSeries31.MarkerSize = 6;
            newSeries32.MarkerSize = 6;
            newSeries33.MarkerSize = 6;
            newSeries34.MarkerSize = 6;
            newSeries35.MarkerSize = 6;
            newSeries36.MarkerSize = 6;
            newSeries37.MarkerSize = 6;
            newSeries38.MarkerSize = 6;
            newSeries39.MarkerSize = 6;
            newSeries40.MarkerSize = 6;
            newSeries41.MarkerSize = 6;
            newSeries42.MarkerSize = 6;
            newSeries43.MarkerSize = 6;
            newSeries44.MarkerSize = 6;
            newSeries45.MarkerSize = 6;
            newSeries46.MarkerSize = 6;
            newSeries47.MarkerSize = 6;
            newSeries48.MarkerSize = 6;
            newSeries49.MarkerSize = 6;
            newSeries50.MarkerSize = 6;
            newSeries51.MarkerSize = 6;
            newSeries52.MarkerSize = 6;
            newSeries53.MarkerSize = 6;
            newSeries54.MarkerSize = 6;
            newSeries55.MarkerSize = 6;
            newSeries56.MarkerSize = 6;
            newSeries57.MarkerSize = 6;
            newSeries58.MarkerSize = 6;
            newSeries59.MarkerSize = 6;
            newSeries60.MarkerSize = 6;
            newSeries61.MarkerSize = 6;
            newSeries62.MarkerSize = 6;
            newSeries63.MarkerSize = 6;
            newSeries64.MarkerSize = 6;
            newSeries65.MarkerSize = 6;
            newSeries66.MarkerSize = 6;
            newSeries67.MarkerSize = 6;
            newSeries68.MarkerSize = 6;
            newSeries69.MarkerSize = 6;
            newSeries70.MarkerSize = 6;
            newSeries71.MarkerSize = 6;
            newSeries72.MarkerSize = 6;
            newSeries73.MarkerSize = 6;
            newSeries74.MarkerSize = 6;
            newSeries75.MarkerSize = 6;
            newSeries76.MarkerSize = 6;
            newSeries77.MarkerSize = 6;
            newSeries78.MarkerSize = 6;
            newSeries79.MarkerSize = 6;
            newSeries80.MarkerSize = 6;
            newSeries81.MarkerSize = 6;
            newSeries82.MarkerSize = 6;
            newSeries83.MarkerSize = 6;
            newSeries84.MarkerSize = 6;
            newSeries85.MarkerSize = 6;
            newSeries86.MarkerSize = 6;
            newSeries87.MarkerSize = 6;
            newSeries88.MarkerSize = 6;
            newSeries89.MarkerSize = 6;
            newSeries90.MarkerSize = 6;
            newSeries91.MarkerSize = 6;
            newSeries92.MarkerSize = 6;
            newSeries93.MarkerSize = 6;
            newSeries94.MarkerSize = 6;
            newSeries95.MarkerSize = 6;
            newSeries96.MarkerSize = 6;
            newSeries97.MarkerSize = 6;
            newSeries98.MarkerSize = 6;
            newSeries99.MarkerSize = 6;
            newSeries100.MarkerSize = 6;
            newSeries101.MarkerSize = 6;
            newSeries102.MarkerSize = 6;
            newSeries103.MarkerSize = 6;
            newSeries104.MarkerSize = 6;
            newSeries105.MarkerSize = 6;
            newSeries106.MarkerSize = 6;
            newSeries107.MarkerSize = 6;
            newSeries108.MarkerSize = 6;
            newSeries109.MarkerSize = 6;
            newSeries110.MarkerSize = 6;
            newSeries111.MarkerSize = 6;
            newSeries112.MarkerSize = 6;
            newSeries113.MarkerSize = 6;
            newSeries114.MarkerSize = 6;
            newSeries115.MarkerSize = 6;
            newSeries116.MarkerSize = 6;
            newSeries117.MarkerSize = 6;
            newSeries118.MarkerSize = 6;
            newSeries119.MarkerSize = 6;
            newSeries120.MarkerSize = 6;
            newSeries121.MarkerSize = 6;
            newSeries122.MarkerSize = 6;
            newSeries123.MarkerSize = 6;
            newSeries124.MarkerSize = 6;
            newSeries125.MarkerSize = 6;
            newSeries126.MarkerSize = 6;
            newSeries127.MarkerSize = 6;
            newSeries128.MarkerSize = 6;
            newSeries129.MarkerSize = 6;
            newSeries130.MarkerSize = 6;
            newSeries131.MarkerSize = 6;
            newSeries132.MarkerSize = 6;
            newSeries133.MarkerSize = 6;
            newSeries134.MarkerSize = 6;
            newSeries135.MarkerSize = 6;
            newSeries136.MarkerSize = 6;
            newSeries137.MarkerSize = 6;
            newSeries138.MarkerSize = 6;
            newSeries139.MarkerSize = 6;
            newSeries140.MarkerSize = 6;
            newSeries141.MarkerSize = 6;
            newSeries142.MarkerSize = 6;
            newSeries143.MarkerSize = 6;
            newSeries144.MarkerSize = 6;
            newSeries145.MarkerSize = 6;
            newSeries146.MarkerSize = 6;
            newSeries147.MarkerSize = 6;
            newSeries148.MarkerSize = 6;
            newSeries149.MarkerSize = 6;
            newSeries150.MarkerSize = 6;
            newSeries151.MarkerSize = 6;
            newSeries152.MarkerSize = 6;
            newSeries153.MarkerSize = 6;
            newSeries154.MarkerSize = 6;
            newSeries155.MarkerSize = 6;
            newSeries156.MarkerSize = 6;
            newSeries157.MarkerSize = 6;
            newSeries158.MarkerSize = 6;
            newSeries159.MarkerSize = 6;
            newSeries160.MarkerSize = 6;
            newSeries161.MarkerSize = 6;
            newSeries162.MarkerSize = 6;
            newSeries163.MarkerSize = 6;
            newSeries164.MarkerSize = 6;
            newSeries165.MarkerSize = 6;
            newSeries166.MarkerSize = 6;
            newSeries167.MarkerSize = 6;
            newSeries168.MarkerSize = 6;
            newSeries169.MarkerSize = 6;
            newSeries170.MarkerSize = 6;
            newSeries171.MarkerSize = 6;
            newSeries172.MarkerSize = 6;
            newSeries173.MarkerSize = 6;
            newSeries174.MarkerSize = 6;
            newSeries175.MarkerSize = 6;
            newSeries176.MarkerSize = 6;
            newSeries177.MarkerSize = 6;
            newSeries178.MarkerSize = 6;
            newSeries179.MarkerSize = 6;
            newSeries180.MarkerSize = 6;
            newSeries181.MarkerSize = 6;
            newSeries182.MarkerSize = 6;
            newSeries183.MarkerSize = 6;
            newSeries184.MarkerSize = 6;
            newSeries185.MarkerSize = 6;
            newSeries186.MarkerSize = 6;
            newSeries187.MarkerSize = 6;
            newSeries188.MarkerSize = 6;
            newSeries189.MarkerSize = 6;
            newSeries190.MarkerSize = 6;
            newSeries191.MarkerSize = 6;
            newSeries192.MarkerSize = 6;
            newSeries193.MarkerSize = 6;
            newSeries194.MarkerSize = 6;
            newSeries195.MarkerSize = 6;
            newSeries196.MarkerSize = 6;
            newSeries197.MarkerSize = 6;
            newSeries198.MarkerSize = 6;
            newSeries199.MarkerSize = 6;
            newSeries200.MarkerSize = 6;
            newSeries201.MarkerSize = 6;
            newSeries202.MarkerSize = 6;
            newSeries203.MarkerSize = 6;
            newSeries204.MarkerSize = 6;
            newSeries205.MarkerSize = 6;
            newSeries206.MarkerSize = 6;
            newSeries207.MarkerSize = 6;
            newSeries208.MarkerSize = 6;
            newSeries209.MarkerSize = 6;
            newSeries210.MarkerSize = 6;
            newSeries211.MarkerSize = 6;
            newSeries212.MarkerSize = 6;
            newSeries213.MarkerSize = 6;
            newSeries214.MarkerSize = 6;
            newSeries215.MarkerSize = 6;
            newSeries216.MarkerSize = 6;
            newSeries217.MarkerSize = 6;
            newSeries218.MarkerSize = 6;
            newSeries219.MarkerSize = 6;
            newSeries220.MarkerSize = 6;
            newSeries221.MarkerSize = 6;
            newSeries222.MarkerSize = 6;
            newSeries223.MarkerSize = 6;
            newSeries224.MarkerSize = 6;
            newSeries225.MarkerSize = 6;
            newSeries226.MarkerSize = 6;
            newSeries227.MarkerSize = 6;
            newSeries228.MarkerSize = 6;
            newSeries229.MarkerSize = 6;
            newSeries230.MarkerSize = 6;
            newSeries231.MarkerSize = 6;
            newSeries232.MarkerSize = 6;
            newSeries233.MarkerSize = 6;
            newSeries234.MarkerSize = 6;
            newSeries235.MarkerSize = 6;
            newSeries236.MarkerSize = 6;
            newSeries237.MarkerSize = 6;
            newSeries238.MarkerSize = 6;
            newSeries239.MarkerSize = 6;
            newSeries240.MarkerSize = 6;
            newSeries241.MarkerSize = 6;
            newSeries242.MarkerSize = 6;
            newSeries243.MarkerSize = 6;
            newSeries244.MarkerSize = 6;
            newSeries245.MarkerSize = 6;
            newSeries246.MarkerSize = 6;
            newSeries247.MarkerSize = 6;
            newSeries248.MarkerSize = 6;
            newSeries249.MarkerSize = 6;
            newSeries250.MarkerSize = 6;
            newSeries251.MarkerSize = 6;
            newSeries252.MarkerSize = 6;
            newSeries253.MarkerSize = 6;
            newSeries254.MarkerSize = 6;
            newSeries255.MarkerSize = 6;
            newSeries256.MarkerSize = 6;
            newSeries257.MarkerSize = 6;
            newSeries258.MarkerSize = 6;
            newSeries259.MarkerSize = 6;
            newSeries260.MarkerSize = 6;
            newSeries261.MarkerSize = 6;
            newSeries262.MarkerSize = 6;
            newSeries263.MarkerSize = 6;
            newSeries264.MarkerSize = 6;
            newSeries265.MarkerSize = 6;
            newSeries266.MarkerSize = 6;
            newSeries267.MarkerSize = 6;
            newSeries268.MarkerSize = 6;
            newSeries269.MarkerSize = 6;
            newSeries270.MarkerSize = 6;
            newSeries271.MarkerSize = 6;
            newSeries272.MarkerSize = 6;
            newSeries273.MarkerSize = 6;
            newSeries274.MarkerSize = 6;
            newSeries275.MarkerSize = 6;
            newSeries276.MarkerSize = 6;
            newSeries277.MarkerSize = 6;
            newSeries278.MarkerSize = 6;
            newSeries279.MarkerSize = 6;
            newSeries280.MarkerSize = 6;
            newSeries281.MarkerSize = 6;
            newSeries282.MarkerSize = 6;
            newSeries283.MarkerSize = 6;
            newSeries284.MarkerSize = 6;
            newSeries285.MarkerSize = 6;
            newSeries286.MarkerSize = 6;
            newSeries287.MarkerSize = 6;
            newSeries288.MarkerSize = 6;
            newSeries289.MarkerSize = 6;
            newSeries290.MarkerSize = 6;
            newSeries291.MarkerSize = 6;
            newSeries292.MarkerSize = 6;
            newSeries293.MarkerSize = 6;
            newSeries294.MarkerSize = 6;
            newSeries295.MarkerSize = 6;
            newSeries296.MarkerSize = 6;
            newSeries297.MarkerSize = 6;
            newSeries298.MarkerSize = 6;
            newSeries299.MarkerSize = 6;
            newSeries300.MarkerSize = 6;
            newSeries301.MarkerSize = 6;
            newSeries302.MarkerSize = 6;
            newSeries303.MarkerSize = 6;
            newSeries304.MarkerSize = 6;
            newSeries305.MarkerSize = 6;
            newSeries306.MarkerSize = 6;
            newSeries307.MarkerSize = 6;
            newSeries308.MarkerSize = 6;
            newSeries309.MarkerSize = 6;
            newSeries310.MarkerSize = 6;
            newSeries311.MarkerSize = 6;
            newSeries312.MarkerSize = 6;
            newSeries313.MarkerSize = 6;
            newSeries314.MarkerSize = 6;
            newSeries315.MarkerSize = 6;
            newSeries316.MarkerSize = 6;
            newSeries317.MarkerSize = 6;
            newSeries318.MarkerSize = 6;
            newSeries319.MarkerSize = 6;
            newSeries320.MarkerSize = 6;
            newSeries321.MarkerSize = 6;
            newSeries322.MarkerSize = 6;
            newSeries323.MarkerSize = 6;
            newSeries324.MarkerSize = 6;
            newSeries325.MarkerSize = 6;
            newSeries326.MarkerSize = 6;
            newSeries327.MarkerSize = 6;
            newSeries328.MarkerSize = 6;
            newSeries329.MarkerSize = 6;
            newSeries330.MarkerSize = 6;
            newSeries331.MarkerSize = 6;
            newSeries332.MarkerSize = 6;
            newSeries333.MarkerSize = 6;
            newSeries334.MarkerSize = 6;
            newSeries335.MarkerSize = 6;
            newSeries336.MarkerSize = 6;
            newSeries337.MarkerSize = 6;
            newSeries338.MarkerSize = 6;
            newSeries339.MarkerSize = 6;
            newSeries340.MarkerSize = 6;
            newSeries341.MarkerSize = 6;
            newSeries342.MarkerSize = 6;
            newSeries343.MarkerSize = 6;
            newSeries344.MarkerSize = 6;
            newSeries345.MarkerSize = 6;
            newSeries346.MarkerSize = 6;
            newSeries347.MarkerSize = 6;
            newSeries348.MarkerSize = 6;
            newSeries349.MarkerSize = 6;
            newSeries350.MarkerSize = 6;




            chart1.ChartAreas[0].AxisY.Maximum = Double.NaN;
            chart2.ChartAreas[0].AxisY.Maximum = Double.NaN;
            chart3.ChartAreas[0].AxisY.Maximum = Double.NaN;
            chart4.ChartAreas[0].AxisY.Maximum = Double.NaN;
            chart5.ChartAreas[0].AxisY.Maximum = Double.NaN;
            chart6.ChartAreas[0].AxisY.Maximum = Double.NaN;
            chart7.ChartAreas[0].AxisY.Maximum = Double.NaN;
            chart8.ChartAreas[0].AxisY.Maximum = Double.NaN; 
            chart9.ChartAreas[0].AxisY.Maximum = Double.NaN;
            chart10.ChartAreas[0].AxisY.Maximum = Double.NaN;
            chart11.ChartAreas[0].AxisY.Maximum = Double.NaN;
            chart12.ChartAreas[0].AxisY.Maximum = Double.NaN;
            chart13.ChartAreas[0].AxisY.Maximum = Double.NaN;
            chart14.ChartAreas[0].AxisY.Maximum = Double.NaN;
            chart1.ChartAreas[0].AxisY.Minimum = Double.NaN;
            chart2.ChartAreas[0].AxisY.Minimum = Double.NaN;
            chart3.ChartAreas[0].AxisY.Minimum = Double.NaN;
            chart4.ChartAreas[0].AxisY.Minimum = Double.NaN;
            chart5.ChartAreas[0].AxisY.Minimum = Double.NaN;
            chart6.ChartAreas[0].AxisY.Minimum = Double.NaN;
            chart7.ChartAreas[0].AxisY.Minimum = Double.NaN;
            chart8.ChartAreas[0].AxisY.Minimum = Double.NaN;
            chart9.ChartAreas[0].AxisY.Minimum = Double.NaN;
            chart10.ChartAreas[0].AxisY.Minimum = Double.NaN;
            chart11.ChartAreas[0].AxisY.Minimum = Double.NaN;
            chart12.ChartAreas[0].AxisY.Minimum = Double.NaN;
            chart13.ChartAreas[0].AxisY.Minimum = Double.NaN;
            chart14.ChartAreas[0].AxisY.Minimum = Double.NaN;




            chart1.ChartAreas[0].CursorX.IsUserSelectionEnabled = true;
            chart1.ChartAreas[0].CursorY.IsUserSelectionEnabled = true;

            chart2.ChartAreas[0].CursorX.IsUserSelectionEnabled = true;
            chart2.ChartAreas[0].CursorY.IsUserSelectionEnabled = true;

            chart3.ChartAreas[0].CursorX.IsUserSelectionEnabled = true;
            chart3.ChartAreas[0].CursorY.IsUserSelectionEnabled = true;

            chart4.ChartAreas[0].CursorX.IsUserSelectionEnabled = true;
            chart4.ChartAreas[0].CursorY.IsUserSelectionEnabled = true;

            chart5.ChartAreas[0].CursorX.IsUserSelectionEnabled = true;
            chart5.ChartAreas[0].CursorY.IsUserSelectionEnabled = true;

            chart6.ChartAreas[0].CursorX.IsUserSelectionEnabled = true;
            chart6.ChartAreas[0].CursorY.IsUserSelectionEnabled = true;

            chart7.ChartAreas[0].CursorX.IsUserSelectionEnabled = true;
            chart7.ChartAreas[0].CursorY.IsUserSelectionEnabled = true;

            chart8.ChartAreas[0].CursorX.IsUserSelectionEnabled = true;
            chart8.ChartAreas[0].CursorY.IsUserSelectionEnabled = true;
            //chart8.ChartAreas[0].AxisY.ScaleView.Zoomable = true;
            //chart1.ChartAreas[0].CursorY.Interval = 0.01;


            //chart8.ChartAreas[0].CursorX.IsUserEnabled = true;
            //chart8.ChartAreas[0].CursorX.IsUserSelectionEnabled = true;
            //chart8.ChartAreas[0].AxisX.ScaleView.Zoomable = true;
            //chart8.ChartAreas[0].AxisX.ScrollBar.IsPositionedInside = true;

            //chart8.ChartAreas[0].CursorY.IsUserEnabled = true;
            //chart8.ChartAreas[0].CursorY.IsUserSelectionEnabled = true;
            //chart8.ChartAreas[0].AxisY.ScaleView.Zoomable = true;
            //chart8.ChartAreas[0].AxisY.ScrollBar.IsPositionedInside = true;



            chart9.ChartAreas[0].CursorX.IsUserSelectionEnabled = true;
            chart9.ChartAreas[0].CursorY.IsUserSelectionEnabled = true;

            chart10.ChartAreas[0].CursorX.IsUserSelectionEnabled = true;
            chart10.ChartAreas[0].CursorY.IsUserSelectionEnabled = true;

            chart11.ChartAreas[0].CursorX.IsUserSelectionEnabled = true;
            chart11.ChartAreas[0].CursorY.IsUserSelectionEnabled = true;

            chart12.ChartAreas[0].CursorX.IsUserSelectionEnabled = true;
            chart12.ChartAreas[0].CursorY.IsUserSelectionEnabled = true;

            chart13.ChartAreas[0].CursorX.IsUserSelectionEnabled = true;
            chart13.ChartAreas[0].CursorY.IsUserSelectionEnabled = true;

            chart14.ChartAreas[0].CursorX.IsUserSelectionEnabled = true;
            chart14.ChartAreas[0].CursorY.IsUserSelectionEnabled = true;


            for (int i = 0; i < RNC_List_Table.Rows.Count; i++)
            {
                chart14.Series[i].ChartType = SeriesChartType.Line;
                chart14.Series[i].BorderWidth = 3;
                chart14.Series[i].EmptyPointStyle.Color = Color.Transparent;
                chart14.Series[i].XValueType = ChartValueType.DateTime;
                chart14.Series[i].ToolTip = Series_List[i, 0] + "_" + "#VALX [#VALY]";
                chart14.Series[i].IsValueShownAsLabel = false;
            }



            string Old_RNC_Name = Series_List[0, 0].ToString();
            MIN_MAX_KPI_MAT[1, 0] = Old_RNC_Name;
            int series_index = 0;



            double Payload_Max = -10000000;
            double Payload_Min = 10000000;
            double Availability_Max = -10000000;
            double Availability_Min = 10000000;
            double DL_User_THR_Max = -10000000;
            double DL_User_THR_Min = 10000000;
            double DL_Cell_THR_Max = -10000000;
            double DL_Cell_THR_Min = 10000000;
            double UL_User_THR_Max = -10000000;
            double UL_User_THR_Min = 10000000;
            double Latency_Max = -10000000;
            double Latency_Min = 10000000;
            double Service_Max = -10000000;
            double Service_Min = 10000000;
            double RRC_Max = -10000000;
            double RRC_Min = 10000000;
            double ERAB_Setup_Max = -10000000;
            double ERAB_Setup_Min = 10000000;
            double ERAB_Drop_Max = -10000000;
            double ERAB_Drop_Min = 10000000;
            double Inter_Max = -10000000;
            double Inter_Min = 10000000;
            double Intra_Max = -10000000;
            double Intra_Min = 10000000;
            double PUCCH_Max = -10000000;
            double PUCCH_Min = 10000000;
            double PUSCH_Max = -10000000;
            double PUSCH_Min = 10000000;



            int y1 = 0;

            for (int i = 0; i < LTE_RNC_Data_Table.Rows.Count; i++)
            {
                DateTime dt = Convert.ToDateTime((LTE_RNC_Data_Table.Rows[i]).ItemArray[0]);
                DateTime dt1 = dt.AddHours(23);
                string RNC = Convert.ToString((LTE_RNC_Data_Table.Rows[i]).ItemArray[1]);
                double Payload = Convert.ToDouble((LTE_RNC_Data_Table.Rows[i]).ItemArray[16]);
                double Availability = Convert.ToDouble((LTE_RNC_Data_Table.Rows[i]).ItemArray[15]);
                double DL_User_THR = Convert.ToDouble((LTE_RNC_Data_Table.Rows[i]).ItemArray[2]);
                double DL_Cell_THR = Convert.ToDouble((LTE_RNC_Data_Table.Rows[i]).ItemArray[3]);
                double UL_User_THR = Convert.ToDouble((LTE_RNC_Data_Table.Rows[i]).ItemArray[4]);
                double Latency = Convert.ToDouble((LTE_RNC_Data_Table.Rows[i]).ItemArray[5]);
                double Service = Convert.ToDouble((LTE_RNC_Data_Table.Rows[i]).ItemArray[7]);
                double RRC = Convert.ToDouble((LTE_RNC_Data_Table.Rows[i]).ItemArray[10]);
                double ERAB_Setup = Convert.ToDouble((LTE_RNC_Data_Table.Rows[i]).ItemArray[8]);
                double ERAB_Drop = Convert.ToDouble((LTE_RNC_Data_Table.Rows[i]).ItemArray[9]);
                double Inter = Convert.ToDouble((LTE_RNC_Data_Table.Rows[i]).ItemArray[11]);
                double Intra = Convert.ToDouble((LTE_RNC_Data_Table.Rows[i]).ItemArray[12]);
                double PUCCH = Convert.ToDouble((LTE_RNC_Data_Table.Rows[i]).ItemArray[13]);
                double PUSCH = Convert.ToDouble((LTE_RNC_Data_Table.Rows[i]).ItemArray[14]);

                if (RNC == Old_RNC_Name)
                {
                    int SI = series_index;
                    chart1.Series[SI].Points.AddXY(dt, Payload);
                    chart2.Series[SI].Points.AddXY(dt, Availability);
                    chart3.Series[SI].Points.AddXY(dt, DL_User_THR);
                    chart4.Series[SI].Points.AddXY(dt, DL_Cell_THR);
                    chart5.Series[SI].Points.AddXY(dt, UL_User_THR);
                    chart6.Series[SI].Points.AddXY(dt, Latency);
                    chart7.Series[SI].Points.AddXY(dt, Service);
                    chart8.Series[SI].Points.AddXY(dt, RRC);
                    chart9.Series[SI].Points.AddXY(dt, ERAB_Setup);
                    chart10.Series[SI].Points.AddXY(dt, ERAB_Drop);
                    chart11.Series[SI].Points.AddXY(dt, Inter);
                    chart12.Series[SI].Points.AddXY(dt, Intra);
                    chart13.Series[SI].Points.AddXY(dt, PUCCH);
                    chart14.Series[SI].Points.AddXY(dt, PUSCH);

                    chart1.Series[SI].LegendText = Old_RNC_Name;
                    chart2.Series[SI].LegendText = Old_RNC_Name;
                    chart3.Series[SI].LegendText = Old_RNC_Name;
                    chart4.Series[SI].LegendText = Old_RNC_Name;
                    chart5.Series[SI].LegendText = Old_RNC_Name;
                    chart6.Series[SI].LegendText = Old_RNC_Name;
                    chart7.Series[SI].LegendText = Old_RNC_Name;
                    chart8.Series[SI].LegendText = Old_RNC_Name;
                    chart9.Series[SI].LegendText = Old_RNC_Name;
                    chart10.Series[SI].LegendText = Old_RNC_Name;
                    chart11.Series[SI].LegendText = Old_RNC_Name;
                    chart12.Series[SI].LegendText = Old_RNC_Name;
                    chart13.Series[SI].LegendText = Old_RNC_Name;
                    chart14.Series[SI].LegendText = Old_RNC_Name;



                    if (Payload > Payload_Max)
                    {
                        Payload_Max = Math.Round(Payload, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 1] = Convert.ToString(Payload_Max);
                    }
                    if (Payload < Payload_Min)
                    {
                        Payload_Min = Math.Round(Payload, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 2] = Convert.ToString(Payload_Min);
                    }
                    if (Availability > Availability_Max)
                    {
                        Availability_Max = Math.Round(Availability, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 3] = Convert.ToString(Availability_Max);
                    }
                    if (Availability < Availability_Min)
                    {
                        Availability_Min = Math.Round(Availability, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 4] = Convert.ToString(Availability_Min);
                    }
                    if (DL_User_THR > DL_User_THR_Max)
                    {
                        DL_User_THR_Max = Math.Round(DL_User_THR, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 5] = Convert.ToString(DL_User_THR_Max);
                    }
                    if (DL_User_THR < DL_User_THR_Min)
                    {
                        DL_User_THR_Min = Math.Round(DL_User_THR, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 6] = Convert.ToString(DL_User_THR_Min);
                    }
                    if (DL_Cell_THR > DL_Cell_THR_Max)
                    {
                        DL_Cell_THR_Max = Math.Round(DL_Cell_THR, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 7] = Convert.ToString(DL_Cell_THR_Max);
                    }
                    if (DL_Cell_THR < DL_Cell_THR_Min)
                    {
                        DL_Cell_THR_Min = Math.Round(DL_Cell_THR, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 8] = Convert.ToString(DL_Cell_THR_Min);
                    }
                    if (UL_User_THR > UL_User_THR_Max)
                    {
                        UL_User_THR_Max = Math.Round(UL_User_THR, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 9] = Convert.ToString(UL_User_THR_Max);
                    }
                    if (UL_User_THR < UL_User_THR_Min)
                    {
                        UL_User_THR_Min = Math.Round(UL_User_THR, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 10] = Convert.ToString(UL_User_THR_Min);
                    }
                    if (Latency > Latency_Max)
                    {
                        Latency_Max = Math.Round(Latency, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 11] = Convert.ToString(Latency_Max);
                    }
                    if (Latency < Latency_Min)
                    {
                        Latency_Min = Math.Round(Latency, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 12] = Convert.ToString(Latency_Min);
                    }
                    if (Service > Service_Max)
                    {
                        Service_Max = Math.Round(Service, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 13] = Convert.ToString(Service_Max);
                    }
                    if (Service < Service_Min)
                    {
                        Service_Min = Math.Round(Service, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 14] = Convert.ToString(Service_Min);
                    }
                    if (RRC > RRC_Max)
                    {
                        RRC_Max = Math.Round(RRC, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 15] = Convert.ToString(RRC_Max);
                    }
                    if (RRC < RRC_Min)
                    {
                        RRC_Min = Math.Round(RRC, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 16] = Convert.ToString(RRC_Min);
                    }
                    if (ERAB_Setup > ERAB_Setup_Max)
                    {
                        ERAB_Setup_Max = Math.Round(ERAB_Setup, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 17] = Convert.ToString(ERAB_Setup_Max);
                    }
                    if (ERAB_Setup < ERAB_Setup_Min)
                    {
                        ERAB_Setup_Min = Math.Round(ERAB_Setup, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 18] = Convert.ToString(ERAB_Setup_Min);
                    }
                    if (ERAB_Drop > ERAB_Drop_Max)
                    {
                        ERAB_Drop_Max = Math.Round(ERAB_Drop, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 19] = Convert.ToString(ERAB_Drop_Max);
                    }
                    if (ERAB_Drop < ERAB_Drop_Min)
                    {
                        ERAB_Drop_Min = Math.Round(ERAB_Drop, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 20] = Convert.ToString(ERAB_Drop_Min);
                    }
                    if (Inter > Inter_Max)
                    {
                        Inter_Max = Math.Round(Inter, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 21] = Convert.ToString(Inter_Max);
                    }
                    if (Inter < Inter_Min)
                    {
                        Inter_Min = Math.Round(Inter, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 22] = Convert.ToString(Inter_Min);
                    }
                    if (Intra > Intra_Max)
                    {
                        Intra_Max = Math.Round(Intra, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 23] = Convert.ToString(Intra_Max);
                    }
                    if (Intra < Intra_Min)
                    {
                        Intra_Min = Math.Round(Intra, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 24] = Convert.ToString(Intra_Min);
                    }
                    if (PUCCH > PUCCH_Max)
                    {
                        PUCCH_Max = Math.Round(PUCCH, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 25] = Convert.ToString(PUCCH_Max);
                    }
                    if (PUCCH < PUCCH_Min)
                    {
                        PUCCH_Min = Math.Round(PUCCH, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 26] = Convert.ToString(PUCCH_Min);
                    }
                    if (PUSCH > PUSCH_Max)
                    {
                        PUSCH_Max = Math.Round(PUSCH, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 27] = Convert.ToString(PUSCH_Max);
                    }
                    if (PUSCH < PUSCH_Min)
                    {
                        PUSCH_Min = Math.Round(PUSCH, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 28] = Convert.ToString(PUSCH_Min);
                    }




                }
                else
                {
                    Payload_Max = -10000000;
                    Payload_Min = 10000000;
                    Availability_Max = -10000000;
                    Availability_Min = 10000000;
                    DL_User_THR_Max = -10000000;
                    DL_User_THR_Min = 10000000;
                    DL_Cell_THR_Max = -10000000;
                    DL_Cell_THR_Min = 10000000;
                    UL_User_THR_Max = -10000000;
                    UL_User_THR_Min = 10000000;
                    Latency_Max = -10000000;
                    Latency_Min = 10000000;
                    Service_Max = -10000000;
                    Service_Min = 10000000;
                    RRC_Max = -10000000;
                    RRC_Min = 10000000;
                    ERAB_Setup_Max = -10000000;
                    ERAB_Setup_Min = 10000000;
                    ERAB_Drop_Max = -10000000;
                    ERAB_Drop_Min = 10000000;
                    Inter_Max = -10000000;
                    Inter_Min = 10000000;
                    Intra_Max = -10000000;
                    Intra_Min = 10000000;
                    PUCCH_Max = -10000000;
                    PUCCH_Min = 10000000;
                    PUSCH_Max = -10000000;
                    PUSCH_Min = 10000000;


                    series_index++;
                    MIN_MAX_KPI_MAT[series_index + 1, 0] = RNC;
                    if (Payload > Payload_Max)
                    {
                        Payload_Max = Math.Round(Payload, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 1] = Convert.ToString(Payload_Max);
                    }
                    if (Payload < Payload_Min)
                    {
                        Payload_Min = Math.Round(Payload, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 2] = Convert.ToString(Payload_Min);
                    }
                    if (Availability > Availability_Max)
                    {
                        Availability_Max = Math.Round(Availability, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 3] = Convert.ToString(Availability_Max);
                    }
                    if (Availability < Availability_Min)
                    {
                        Availability_Min = Math.Round(Availability, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 4] = Convert.ToString(Availability_Min);
                    }
                    if (DL_User_THR > DL_User_THR_Max)
                    {
                        DL_User_THR_Max = Math.Round(DL_User_THR, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 5] = Convert.ToString(DL_User_THR_Max);
                    }
                    if (DL_User_THR < DL_User_THR_Min)
                    {
                        DL_User_THR_Min = Math.Round(DL_User_THR, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 6] = Convert.ToString(DL_User_THR_Min);
                    }
                    if (DL_Cell_THR > DL_Cell_THR_Max)
                    {
                        DL_Cell_THR_Max = Math.Round(DL_Cell_THR, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 7] = Convert.ToString(DL_Cell_THR_Max);
                    }
                    if (DL_Cell_THR < DL_Cell_THR_Min)
                    {
                        DL_Cell_THR_Min = Math.Round(DL_Cell_THR, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 8] = Convert.ToString(DL_Cell_THR_Min);
                    }
                    if (UL_User_THR > UL_User_THR_Max)
                    {
                        UL_User_THR_Max = Math.Round(UL_User_THR, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 9] = Convert.ToString(UL_User_THR_Max);
                    }
                    if (UL_User_THR < UL_User_THR_Min)
                    {
                        UL_User_THR_Min = Math.Round(UL_User_THR, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 10] = Convert.ToString(UL_User_THR_Min);
                    }
                    if (Latency > Latency_Max)
                    {
                        Latency_Max = Math.Round(Latency, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 11] = Convert.ToString(Latency_Max);
                    }
                    if (Latency < Latency_Min)
                    {
                        Latency_Min = Math.Round(Latency, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 12] = Convert.ToString(Latency_Min);
                    }
                    if (Service > Service_Max)
                    {
                        Service_Max = Math.Round(Service, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 13] = Convert.ToString(Service_Max);
                    }
                    if (Service < Service_Min)
                    {
                        Service_Min = Math.Round(Service, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 14] = Convert.ToString(Service_Min);
                    }
                    if (RRC > RRC_Max)
                    {
                        RRC_Max = Math.Round(RRC, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 15] = Convert.ToString(RRC_Max);
                    }
                    if (RRC < RRC_Min)
                    {
                        RRC_Min = Math.Round(RRC, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 16] = Convert.ToString(RRC_Min);
                    }
                    if (ERAB_Setup > ERAB_Setup_Max)
                    {
                        ERAB_Setup_Max = Math.Round(ERAB_Setup, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 17] = Convert.ToString(ERAB_Setup_Max);
                    }
                    if (ERAB_Setup < ERAB_Setup_Min)
                    {
                        ERAB_Setup_Min = Math.Round(ERAB_Setup, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 18] = Convert.ToString(ERAB_Setup_Min);
                    }
                    if (ERAB_Drop > ERAB_Drop_Max)
                    {
                        ERAB_Drop_Max = Math.Round(ERAB_Drop, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 19] = Convert.ToString(ERAB_Drop_Max);
                    }
                    if (ERAB_Drop < ERAB_Drop_Min)
                    {
                        ERAB_Drop_Min = Math.Round(ERAB_Drop, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 20] = Convert.ToString(ERAB_Drop_Min);
                    }
                    if (Inter > Inter_Max)
                    {
                        Inter_Max = Math.Round(Inter, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 21] = Convert.ToString(Inter_Max);
                    }
                    if (Inter < Inter_Min)
                    {
                        Inter_Min = Math.Round(Inter, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 22] = Convert.ToString(Inter_Min);
                    }
                    if (Intra > Intra_Max)
                    {
                        Intra_Max = Math.Round(Intra, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 23] = Convert.ToString(Intra_Max);
                    }
                    if (Intra < Intra_Min)
                    {
                        Intra_Min = Math.Round(Intra, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 24] = Convert.ToString(Intra_Min);
                    }
                    if (PUCCH > PUCCH_Max)
                    {
                        PUCCH_Max = Math.Round(PUCCH, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 25] = Convert.ToString(PUCCH_Max);
                    }
                    if (PUCCH < PUCCH_Min)
                    {
                        PUCCH_Min = Math.Round(PUCCH, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 26] = Convert.ToString(PUCCH_Min);
                    }
                    if (PUSCH > PUSCH_Max)
                    {
                        PUSCH_Max = Math.Round(PUSCH, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 27] = Convert.ToString(PUSCH_Max);
                    }
                    if (PUSCH < PUSCH_Min)
                    {
                        PUSCH_Min = Math.Round(PUSCH, MidpointRounding.AwayFromZero);
                        MIN_MAX_KPI_MAT[series_index + 1, 28] = Convert.ToString(PUSCH_Min);
                    }


                    int SI = series_index;
                    chart1.Series[SI].Points.AddXY(dt, Payload);
                    chart1.Series[SI].Points.AddXY(dt, Availability);
                    chart3.Series[SI].Points.AddXY(dt, DL_User_THR);
                    chart4.Series[SI].Points.AddXY(dt, DL_Cell_THR);
                    chart5.Series[SI].Points.AddXY(dt, UL_User_THR);
                    chart6.Series[SI].Points.AddXY(dt, Latency);
                    chart7.Series[SI].Points.AddXY(dt, Service);
                    chart8.Series[SI].Points.AddXY(dt, RRC);
                    chart9.Series[SI].Points.AddXY(dt, ERAB_Setup);
                    chart10.Series[SI].Points.AddXY(dt, ERAB_Drop);
                    chart11.Series[SI].Points.AddXY(dt, Inter);
                    chart12.Series[SI].Points.AddXY(dt, Intra);
                    chart13.Series[SI].Points.AddXY(dt, PUCCH);
                    chart14.Series[SI].Points.AddXY(dt, PUSCH);

                    chart1.Series[SI].LegendText = Old_RNC_Name;
                    chart2.Series[SI].LegendText = Old_RNC_Name;
                    chart3.Series[SI].LegendText = Old_RNC_Name;
                    chart4.Series[SI].LegendText = Old_RNC_Name;
                    chart5.Series[SI].LegendText = Old_RNC_Name;
                    chart6.Series[SI].LegendText = Old_RNC_Name;
                    chart7.Series[SI].LegendText = Old_RNC_Name;
                    chart8.Series[SI].LegendText = Old_RNC_Name;
                    chart9.Series[SI].LegendText = Old_RNC_Name;
                    chart10.Series[SI].LegendText = Old_RNC_Name;
                    chart11.Series[SI].LegendText = Old_RNC_Name;
                    chart12.Series[SI].LegendText = Old_RNC_Name;
                    chart13.Series[SI].LegendText = Old_RNC_Name;
                    chart14.Series[SI].LegendText = Old_RNC_Name;
                    Old_RNC_Name = RNC;
                }


                double dt1_double = dt1.Year * 10000 + dt1.Month * 100 + dt1.Day;

                if (dt1_double > Max_X)
                {
                    Max_X = dt1_double;
                    Max_X_Date = dt1;
                }
                if (dt1_double < Min_X)
                {
                    Min_X = dt1_double;
                    Min_X_Date = dt1;
                }


            }



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
            chart5.ChartAreas[0].AxisX.Interval = day_interval;
            chart6.ChartAreas[0].AxisX.Interval = day_interval;
            chart7.ChartAreas[0].AxisX.Interval = day_interval;
            chart8.ChartAreas[0].AxisX.Interval = day_interval;
            chart9.ChartAreas[0].AxisX.Interval = day_interval;
            chart10.ChartAreas[0].AxisX.Interval = day_interval;
            chart11.ChartAreas[0].AxisX.Interval = day_interval;
            chart12.ChartAreas[0].AxisX.Interval = day_interval;
            chart13.ChartAreas[0].AxisX.Interval = day_interval;
            chart14.ChartAreas[0].AxisX.Interval = day_interval;

            MessageBox.Show("Data is Loaded");
        }



        private void button6_Click(object sender, EventArgs e)
        {

            chart1.Series[0].Enabled = false;
            chart1.Series[1].Enabled = false;
            chart1.Series[2].Enabled = false;
            chart1.Series[3].Enabled = false;
            chart1.Series[4].Enabled = false;
            chart1.Series[5].Enabled = false;
            chart1.Series[6].Enabled = false;
            chart1.Series[7].Enabled = false;
            chart1.Series[8].Enabled = false;
            chart1.Series[9].Enabled = false;
            chart1.Series[10].Enabled = false;
            chart1.Series[11].Enabled = false;
            chart1.Series[12].Enabled = false;
            chart1.Series[13].Enabled = false;
            chart1.Series[14].Enabled = false;
            chart1.Series[15].Enabled = false;
            chart1.Series[16].Enabled = false;
            chart1.Series[17].Enabled = false;
            chart1.Series[18].Enabled = false;
            chart1.Series[19].Enabled = false;
            chart1.Series[20].Enabled = false;
            chart1.Series[21].Enabled = false;
            chart1.Series[22].Enabled = false;
            chart1.Series[23].Enabled = false;
            chart1.Series[24].Enabled = false;

            chart2.Series[0].Enabled = false;
            chart2.Series[1].Enabled = false;
            chart2.Series[2].Enabled = false;
            chart2.Series[3].Enabled = false;
            chart2.Series[4].Enabled = false;
            chart2.Series[5].Enabled = false;
            chart2.Series[6].Enabled = false;
            chart2.Series[7].Enabled = false;
            chart2.Series[8].Enabled = false;
            chart2.Series[9].Enabled = false;
            chart2.Series[10].Enabled = false;
            chart2.Series[11].Enabled = false;
            chart2.Series[12].Enabled = false;
            chart2.Series[13].Enabled = false;
            chart2.Series[14].Enabled = false;
            chart2.Series[15].Enabled = false;
            chart2.Series[16].Enabled = false;
            chart2.Series[17].Enabled = false;
            chart2.Series[18].Enabled = false;
            chart2.Series[19].Enabled = false;
            chart2.Series[20].Enabled = false;
            chart2.Series[21].Enabled = false;
            chart2.Series[22].Enabled = false;
            chart2.Series[23].Enabled = false;
            chart2.Series[24].Enabled = false;

            chart3.Series[0].Enabled = false;
            chart3.Series[1].Enabled = false;
            chart3.Series[2].Enabled = false;
            chart3.Series[3].Enabled = false;
            chart3.Series[4].Enabled = false;
            chart3.Series[5].Enabled = false;
            chart3.Series[6].Enabled = false;
            chart3.Series[7].Enabled = false;
            chart3.Series[8].Enabled = false;
            chart3.Series[9].Enabled = false;
            chart3.Series[10].Enabled = false;
            chart3.Series[11].Enabled = false;
            chart3.Series[12].Enabled = false;
            chart3.Series[13].Enabled = false;
            chart3.Series[14].Enabled = false;
            chart3.Series[15].Enabled = false;
            chart3.Series[16].Enabled = false;
            chart3.Series[17].Enabled = false;
            chart3.Series[18].Enabled = false;
            chart3.Series[19].Enabled = false;
            chart3.Series[20].Enabled = false;
            chart3.Series[21].Enabled = false;
            chart3.Series[22].Enabled = false;
            chart3.Series[23].Enabled = false;
            chart3.Series[24].Enabled = false;

            chart4.Series[0].Enabled = false;
            chart4.Series[1].Enabled = false;
            chart4.Series[2].Enabled = false;
            chart4.Series[3].Enabled = false;
            chart4.Series[4].Enabled = false;
            chart4.Series[5].Enabled = false;
            chart4.Series[6].Enabled = false;
            chart4.Series[7].Enabled = false;
            chart4.Series[8].Enabled = false;
            chart4.Series[9].Enabled = false;
            chart4.Series[10].Enabled = false;
            chart4.Series[11].Enabled = false;
            chart4.Series[12].Enabled = false;
            chart4.Series[13].Enabled = false;
            chart4.Series[14].Enabled = false;
            chart4.Series[15].Enabled = false;
            chart4.Series[16].Enabled = false;
            chart4.Series[17].Enabled = false;
            chart4.Series[18].Enabled = false;
            chart4.Series[19].Enabled = false;
            chart4.Series[20].Enabled = false;
            chart4.Series[21].Enabled = false;
            chart4.Series[22].Enabled = false;
            chart4.Series[23].Enabled = false;
            chart4.Series[24].Enabled = false;


            chart5.Series[0].Enabled = false;
            chart5.Series[1].Enabled = false;
            chart5.Series[2].Enabled = false;
            chart5.Series[3].Enabled = false;
            chart5.Series[4].Enabled = false;
            chart5.Series[5].Enabled = false;
            chart5.Series[6].Enabled = false;
            chart5.Series[7].Enabled = false;
            chart5.Series[8].Enabled = false;
            chart5.Series[9].Enabled = false;
            chart5.Series[10].Enabled = false;
            chart5.Series[11].Enabled = false;
            chart5.Series[12].Enabled = false;
            chart5.Series[13].Enabled = false;
            chart5.Series[14].Enabled = false;
            chart5.Series[15].Enabled = false;
            chart5.Series[16].Enabled = false;
            chart5.Series[17].Enabled = false;
            chart5.Series[18].Enabled = false;
            chart5.Series[19].Enabled = false;
            chart5.Series[20].Enabled = false;
            chart5.Series[21].Enabled = false;
            chart5.Series[22].Enabled = false;
            chart5.Series[23].Enabled = false;
            chart5.Series[24].Enabled = false;

            chart6.Series[0].Enabled = false;
            chart6.Series[1].Enabled = false;
            chart6.Series[2].Enabled = false;
            chart6.Series[3].Enabled = false;
            chart6.Series[4].Enabled = false;
            chart6.Series[5].Enabled = false;
            chart6.Series[6].Enabled = false;
            chart6.Series[7].Enabled = false;
            chart6.Series[8].Enabled = false;
            chart6.Series[9].Enabled = false;
            chart6.Series[10].Enabled = false;
            chart6.Series[11].Enabled = false;
            chart6.Series[12].Enabled = false;
            chart6.Series[13].Enabled = false;
            chart6.Series[14].Enabled = false;
            chart6.Series[15].Enabled = false;
            chart6.Series[16].Enabled = false;
            chart6.Series[17].Enabled = false;
            chart6.Series[18].Enabled = false;
            chart6.Series[19].Enabled = false;
            chart6.Series[20].Enabled = false;
            chart6.Series[21].Enabled = false;
            chart6.Series[22].Enabled = false;
            chart6.Series[23].Enabled = false;
            chart6.Series[24].Enabled = false;

            chart7.Series[0].Enabled = false;
            chart7.Series[1].Enabled = false;
            chart7.Series[2].Enabled = false;
            chart7.Series[3].Enabled = false;
            chart7.Series[4].Enabled = false;
            chart7.Series[5].Enabled = false;
            chart7.Series[6].Enabled = false;
            chart7.Series[7].Enabled = false;
            chart7.Series[8].Enabled = false;
            chart7.Series[9].Enabled = false;
            chart7.Series[10].Enabled = false;
            chart7.Series[11].Enabled = false;
            chart7.Series[12].Enabled = false;
            chart7.Series[13].Enabled = false;
            chart7.Series[14].Enabled = false;
            chart7.Series[15].Enabled = false;
            chart7.Series[16].Enabled = false;
            chart7.Series[17].Enabled = false;
            chart7.Series[18].Enabled = false;
            chart7.Series[19].Enabled = false;
            chart7.Series[20].Enabled = false;
            chart7.Series[21].Enabled = false;
            chart7.Series[22].Enabled = false;
            chart7.Series[23].Enabled = false;
            chart7.Series[24].Enabled = false;

            chart8.Series[0].Enabled = false;
            chart8.Series[1].Enabled = false;
            chart8.Series[2].Enabled = false;
            chart8.Series[3].Enabled = false;
            chart8.Series[4].Enabled = false;
            chart8.Series[5].Enabled = false;
            chart8.Series[6].Enabled = false;
            chart8.Series[7].Enabled = false;
            chart8.Series[8].Enabled = false;
            chart8.Series[9].Enabled = false;
            chart8.Series[10].Enabled = false;
            chart8.Series[11].Enabled = false;
            chart8.Series[12].Enabled = false;
            chart8.Series[13].Enabled = false;
            chart8.Series[14].Enabled = false;
            chart8.Series[15].Enabled = false;
            chart8.Series[16].Enabled = false;
            chart8.Series[17].Enabled = false;
            chart8.Series[18].Enabled = false;
            chart8.Series[19].Enabled = false;
            chart8.Series[20].Enabled = false;
            chart8.Series[21].Enabled = false;
            chart8.Series[22].Enabled = false;
            chart8.Series[23].Enabled = false;
            chart8.Series[24].Enabled = false;


            chart9.Series[0].Enabled = false;
            chart9.Series[1].Enabled = false;
            chart9.Series[2].Enabled = false;
            chart9.Series[3].Enabled = false;
            chart9.Series[4].Enabled = false;
            chart9.Series[5].Enabled = false;
            chart9.Series[6].Enabled = false;
            chart9.Series[7].Enabled = false;
            chart9.Series[8].Enabled = false;
            chart9.Series[9].Enabled = false;
            chart9.Series[10].Enabled = false;
            chart9.Series[11].Enabled = false;
            chart9.Series[12].Enabled = false;
            chart9.Series[13].Enabled = false;
            chart9.Series[14].Enabled = false;
            chart9.Series[15].Enabled = false;
            chart9.Series[16].Enabled = false;
            chart9.Series[17].Enabled = false;
            chart9.Series[18].Enabled = false;
            chart9.Series[19].Enabled = false;
            chart9.Series[20].Enabled = false;
            chart9.Series[21].Enabled = false;
            chart9.Series[22].Enabled = false;
            chart9.Series[23].Enabled = false;
            chart9.Series[24].Enabled = false;

            chart10.Series[0].Enabled = false;
            chart10.Series[1].Enabled = false;
            chart10.Series[2].Enabled = false;
            chart10.Series[3].Enabled = false;
            chart10.Series[4].Enabled = false;
            chart10.Series[5].Enabled = false;
            chart10.Series[6].Enabled = false;
            chart10.Series[7].Enabled = false;
            chart10.Series[8].Enabled = false;
            chart10.Series[9].Enabled = false;
            chart10.Series[10].Enabled = false;
            chart10.Series[11].Enabled = false;
            chart10.Series[12].Enabled = false;
            chart10.Series[13].Enabled = false;
            chart10.Series[14].Enabled = false;
            chart10.Series[15].Enabled = false;
            chart10.Series[16].Enabled = false;
            chart10.Series[17].Enabled = false;
            chart10.Series[18].Enabled = false;
            chart10.Series[19].Enabled = false;
            chart10.Series[20].Enabled = false;
            chart10.Series[21].Enabled = false;
            chart10.Series[22].Enabled = false;
            chart10.Series[23].Enabled = false;
            chart10.Series[24].Enabled = false;


            chart11.Series[0].Enabled = false;
            chart11.Series[1].Enabled = false;
            chart11.Series[2].Enabled = false;
            chart11.Series[3].Enabled = false;
            chart11.Series[4].Enabled = false;
            chart11.Series[5].Enabled = false;
            chart11.Series[6].Enabled = false;
            chart11.Series[7].Enabled = false;
            chart11.Series[8].Enabled = false;
            chart11.Series[9].Enabled = false;
            chart11.Series[10].Enabled = false;
            chart11.Series[11].Enabled = false;
            chart11.Series[12].Enabled = false;
            chart11.Series[13].Enabled = false;
            chart11.Series[14].Enabled = false;
            chart11.Series[15].Enabled = false;
            chart11.Series[16].Enabled = false;
            chart11.Series[17].Enabled = false;
            chart11.Series[18].Enabled = false;
            chart11.Series[19].Enabled = false;
            chart11.Series[20].Enabled = false;
            chart11.Series[21].Enabled = false;
            chart11.Series[22].Enabled = false;
            chart11.Series[23].Enabled = false;
            chart11.Series[24].Enabled = false;

            chart12.Series[0].Enabled = false;
            chart12.Series[1].Enabled = false;
            chart12.Series[2].Enabled = false;
            chart12.Series[3].Enabled = false;
            chart12.Series[4].Enabled = false;
            chart12.Series[5].Enabled = false;
            chart12.Series[6].Enabled = false;
            chart12.Series[7].Enabled = false;
            chart12.Series[8].Enabled = false;
            chart12.Series[9].Enabled = false;
            chart12.Series[10].Enabled = false;
            chart12.Series[11].Enabled = false;
            chart12.Series[12].Enabled = false;
            chart12.Series[13].Enabled = false;
            chart12.Series[14].Enabled = false;
            chart12.Series[15].Enabled = false;
            chart12.Series[16].Enabled = false;
            chart12.Series[17].Enabled = false;
            chart12.Series[18].Enabled = false;
            chart12.Series[19].Enabled = false;
            chart12.Series[20].Enabled = false;
            chart12.Series[21].Enabled = false;
            chart12.Series[22].Enabled = false;
            chart12.Series[23].Enabled = false;
            chart12.Series[24].Enabled = false;

            chart13.Series[0].Enabled = false;
            chart13.Series[1].Enabled = false;
            chart13.Series[2].Enabled = false;
            chart13.Series[3].Enabled = false;
            chart13.Series[4].Enabled = false;
            chart13.Series[5].Enabled = false;
            chart13.Series[6].Enabled = false;
            chart13.Series[7].Enabled = false;
            chart13.Series[8].Enabled = false;
            chart13.Series[9].Enabled = false;
            chart13.Series[10].Enabled = false;
            chart13.Series[11].Enabled = false;
            chart13.Series[12].Enabled = false;
            chart13.Series[13].Enabled = false;
            chart13.Series[14].Enabled = false;
            chart13.Series[15].Enabled = false;
            chart13.Series[16].Enabled = false;
            chart13.Series[17].Enabled = false;
            chart13.Series[18].Enabled = false;
            chart13.Series[19].Enabled = false;
            chart13.Series[20].Enabled = false;
            chart13.Series[21].Enabled = false;
            chart13.Series[22].Enabled = false;
            chart13.Series[23].Enabled = false;
            chart13.Series[24].Enabled = false;

            chart14.Series[0].Enabled = false;
            chart14.Series[1].Enabled = false;
            chart14.Series[2].Enabled = false;
            chart14.Series[3].Enabled = false;
            chart14.Series[4].Enabled = false;
            chart14.Series[5].Enabled = false;
            chart14.Series[6].Enabled = false;
            chart14.Series[7].Enabled = false;
            chart14.Series[8].Enabled = false;
            chart14.Series[9].Enabled = false;
            chart14.Series[10].Enabled = false;
            chart14.Series[11].Enabled = false;
            chart14.Series[12].Enabled = false;
            chart14.Series[13].Enabled = false;
            chart14.Series[14].Enabled = false;
            chart14.Series[15].Enabled = false;
            chart14.Series[16].Enabled = false;
            chart14.Series[17].Enabled = false;
            chart14.Series[18].Enabled = false;
            chart14.Series[19].Enabled = false;
            chart14.Series[20].Enabled = false;
            chart14.Series[21].Enabled = false;
            chart14.Series[22].Enabled = false;
            chart14.Series[23].Enabled = false;
            chart14.Series[24].Enabled = false;

            double Payload_Max_All = -10000000;
            double Payload_Min_All = 10000000;
            double Availability_Max_All = -10000000;
            double Availability_Min_All = 10000000;
            double DL_User_THR_Max_All = -10000000;
            double DL_User_THR_Min_All = 10000000;
            double DL_Cell_THR_Max_All = -10000000;
            double DL_Cell_THR_Min_All = 10000000;
            double UL_User_THR_Max_All = -10000000;
            double UL_User_THR_Min_All = 10000000;
            double Latency_Max_All = -10000000;
            double Latency_Min_All = 10000000;
            double Service_Max_All = -10000000;
            double Service_Min_All = 10000000;
            double RRC_Max_All = -10000000;
            double RRC_Min_All = 10000000;
            double ERAB_Setup_Max_All = -10000000;
            double ERAB_Setup_Min_All = 10000000;
            double ERAB_Drop_Max_All = -10000000;
            double ERAB_Drop_Min_All = 10000000;
            double Inter_Max_All = -10000000;
            double Inter_Min_All = 10000000;
            double Intra_Max_All = -10000000;
            double Intra_Min_All = 10000000;
            double PUCCH_Max_All = -10000000;
            double PUCCH_Min_All = 10000000;
            double PUSCH_Max_All = -10000000;
            double PUSCH_Min_All = 10000000;

            for (int k = 0; k < listBox1.SelectedItems.Count; k++)
            {
                string Selected_RNC = listBox1.SelectedItems[k].ToString();
                for (int j = 0; j < RNC_Num; j++)
                {
                    if (Series_List[j, 0].ToString() != "")
                    {
                        string RNC = Series_List[j, 0].ToString();
                        if (Selected_RNC == RNC)
                        {

                            double RNC_Payload_Max = Convert.ToDouble(MIN_MAX_KPI_MAT[j + 1, 1]);
                            double RNC_Payload_Min = Convert.ToDouble(MIN_MAX_KPI_MAT[j + 1, 2]);
                            if (RNC_Payload_Max > Payload_Max_All)
                            {
                                Payload_Max_All = Math.Round(RNC_Payload_Max, MidpointRounding.AwayFromZero) + 1;
                            }
                            if (RNC_Payload_Min < Payload_Min_All)
                            {
                                Payload_Min_All = Math.Round(RNC_Payload_Min, MidpointRounding.AwayFromZero) - 1;
                            }


                            double RNC_Availability_Max = Convert.ToDouble(MIN_MAX_KPI_MAT[j + 1, 3]);
                            double RNC_Availability_Min = Convert.ToDouble(MIN_MAX_KPI_MAT[j + 1, 4]);
                            if (RNC_Availability_Max > Availability_Max_All)
                            {
                                Availability_Max_All = 100;
                            }
                            if (RNC_Availability_Min < Availability_Min_All)
                            {
                                Availability_Min_All = Math.Round(RNC_Availability_Min, MidpointRounding.AwayFromZero) - 0.5;
                            }




                            double RNC_DL_User_THR_Max = Convert.ToDouble(MIN_MAX_KPI_MAT[j + 1, 5]);
                            double RNC_DL_User_THR_Min = Convert.ToDouble(MIN_MAX_KPI_MAT[j + 1, 6]);
                            if (RNC_DL_User_THR_Max > DL_User_THR_Max_All)
                            {
                                DL_User_THR_Max_All = Math.Round(RNC_DL_User_THR_Max, MidpointRounding.AwayFromZero) + 1;
                            }
                            if (RNC_DL_User_THR_Min < DL_User_THR_Min_All)
                            {
                                DL_User_THR_Min_All = Math.Round(RNC_DL_User_THR_Min, MidpointRounding.AwayFromZero) - 1;
                            }


                            double RNC_DL_Cell_THR_Max = Convert.ToDouble(MIN_MAX_KPI_MAT[j + 1, 7]);
                            double RNC_DL_Cell_THR_Min = Convert.ToDouble(MIN_MAX_KPI_MAT[j + 1, 8]);
                            if (RNC_DL_Cell_THR_Max > DL_Cell_THR_Max_All)
                            {
                                DL_Cell_THR_Max_All = Math.Round(RNC_DL_Cell_THR_Max, MidpointRounding.AwayFromZero)+1;
                            }
                            if (RNC_DL_Cell_THR_Min < DL_Cell_THR_Min_All)
                            {
                                DL_Cell_THR_Min_All = Math.Round(RNC_DL_Cell_THR_Min, MidpointRounding.AwayFromZero)-1;
                            }

                            double RNC_UL_User_THR_Max = Convert.ToDouble(MIN_MAX_KPI_MAT[j + 1, 9]);
                            double RNC_UL_User_THR_Min = Convert.ToDouble(MIN_MAX_KPI_MAT[j + 1, 10]);
                            if (RNC_UL_User_THR_Max > UL_User_THR_Max_All)
                            {
                                UL_User_THR_Max_All = Math.Round(RNC_UL_User_THR_Max, MidpointRounding.AwayFromZero) + 1;
                            }
                            if (RNC_UL_User_THR_Min < UL_User_THR_Min_All)
                            {
                                UL_User_THR_Min_All = Math.Round(RNC_UL_User_THR_Min, MidpointRounding.AwayFromZero) - 1;
                            }

                            double RNC_Latency_Max = Convert.ToDouble(MIN_MAX_KPI_MAT[j + 1, 11]);
                            double RNC_Latency_Min = Convert.ToDouble(MIN_MAX_KPI_MAT[j + 1, 12]);
                            if (RNC_Latency_Max > Latency_Max_All)
                            {
                                Latency_Max_All = Math.Round(RNC_Latency_Max, MidpointRounding.AwayFromZero) + 1;
                            }
                            if (RNC_Latency_Min < Latency_Min_All)
                            {
                                Latency_Min_All = Math.Round(RNC_Latency_Min, MidpointRounding.AwayFromZero) - 1;
                            }

                            double RNC_Service_Max = Convert.ToDouble(MIN_MAX_KPI_MAT[j + 1, 13]);
                            double RNC_Service_Min = Convert.ToDouble(MIN_MAX_KPI_MAT[j + 1, 14]);
                            if (RNC_Service_Max > Service_Max_All)
                            {
                                Service_Max_All = 100;
                            }
                            if (RNC_Service_Min < Service_Min_All)
                            {
                                Service_Min_All = Math.Round(RNC_Service_Min, MidpointRounding.AwayFromZero) - 0.5;
                            }
                            double RNC_RRC_Max = Convert.ToDouble(MIN_MAX_KPI_MAT[j + 1, 15]);
                            double RNC_RRC_Min = Convert.ToDouble(MIN_MAX_KPI_MAT[j + 1, 16]);
                            if (RNC_RRC_Max > RRC_Max_All)
                            {
                                RRC_Max_All = 100;
                            }
                            if (RNC_RRC_Min < RRC_Min_All)
                            {
                                RRC_Min_All = Math.Round(RNC_RRC_Min, MidpointRounding.AwayFromZero) - 0.5;
                            }
                            double RNC_ERAB_Setup_Max = Convert.ToDouble(MIN_MAX_KPI_MAT[j + 1, 17]);
                            double RNC_ERAB_Setup_Min = Convert.ToDouble(MIN_MAX_KPI_MAT[j + 1, 18]);
                            if (RNC_ERAB_Setup_Max > ERAB_Setup_Max_All)
                            {
                                ERAB_Setup_Max_All = 100;
                            }
                            if (RNC_ERAB_Setup_Min < ERAB_Setup_Min_All)
                            {
                                ERAB_Setup_Min_All = Math.Round(RNC_ERAB_Setup_Min, MidpointRounding.AwayFromZero) - 0.5;
                            }
                            double RNC_ERAB_Drop_Max = Convert.ToDouble(MIN_MAX_KPI_MAT[j + 1, 19]);
                            double RNC_ERAB_Drop_Min = Convert.ToDouble(MIN_MAX_KPI_MAT[j + 1, 20]);
                            if (RNC_ERAB_Drop_Max > ERAB_Drop_Max_All)
                            {
                                ERAB_Drop_Max_All = Math.Round(RNC_ERAB_Drop_Max, MidpointRounding.AwayFromZero) + 1;
                            }
                            if (RNC_ERAB_Drop_Min < ERAB_Drop_Min_All)
                            {
                                ERAB_Drop_Min_All = 0;
                            }
                            double RNC_Inter_Max = Convert.ToDouble(MIN_MAX_KPI_MAT[j + 1, 21]);
                            double RNC_Inter_Min = Convert.ToDouble(MIN_MAX_KPI_MAT[j + 1, 22]);
                            if (RNC_Inter_Max > Inter_Max_All)
                            {
                                Inter_Max_All = 100;
                            }
                            if (RNC_Inter_Min < Inter_Min_All)
                            {
                                Inter_Min_All = Math.Round(RNC_Inter_Min, MidpointRounding.AwayFromZero) - 0.5;
                            }
                            double RNC_Intra_Max = Convert.ToDouble(MIN_MAX_KPI_MAT[j + 1, 23]);
                            double RNC_Intra_Min = Convert.ToDouble(MIN_MAX_KPI_MAT[j + 1, 24]);
                            if (RNC_Intra_Max > Intra_Max_All)
                            {
                                Intra_Max_All = 100;
                            }
                            if (RNC_Intra_Min < Intra_Min_All)
                            {
                                Intra_Min_All = Math.Round(RNC_Intra_Min, MidpointRounding.AwayFromZero) - 0.5;
                            }
                            double RNC_PUCCH_Max = Convert.ToDouble(MIN_MAX_KPI_MAT[j + 1, 25]);
                            double RNC_PUCCH_Min = Convert.ToDouble(MIN_MAX_KPI_MAT[j + 1, 26]);
                            if (RNC_PUCCH_Max > PUCCH_Max_All)
                            {
                                PUCCH_Max_All = Math.Round(RNC_PUCCH_Max, MidpointRounding.AwayFromZero) + 0.5;
                            }
                            if (RNC_PUCCH_Min < PUCCH_Min_All)
                            {
                                PUCCH_Min_All = Math.Round(RNC_PUCCH_Min, MidpointRounding.AwayFromZero) - 0.5;
                            }
                            double RNC_PUSCH_Max = Convert.ToDouble(MIN_MAX_KPI_MAT[j + 1, 27]);
                            double RNC_PUSCH_Min = Convert.ToDouble(MIN_MAX_KPI_MAT[j + 1, 28]);
                            if (RNC_PUSCH_Max > PUSCH_Max_All)
                            {
                                PUSCH_Max_All = Math.Round(RNC_PUSCH_Max, MidpointRounding.AwayFromZero) + 0.5;
                            }
                            if (RNC_PUSCH_Min < PUSCH_Min_All)
                            {
                                PUSCH_Min_All = Math.Round(RNC_PUSCH_Min, MidpointRounding.AwayFromZero) - 0.5;
                            }




                            string Series_Num = Series_List[j, 1].ToString();
                            int Series_Num1 = Convert.ToInt16(Series_Num.Substring(1, Series_Num.Length - 1));



                            chart1.Series[Series_Num1 - 1].Enabled = true;
                            chart2.Series[Series_Num1 - 1].Enabled = true;
                            chart3.Series[Series_Num1 - 1].Enabled = true;
                            chart4.Series[Series_Num1 - 1].Enabled = true;
                            chart5.Series[Series_Num1 - 1].Enabled = true;
                            chart6.Series[Series_Num1 - 1].Enabled = true;
                            chart7.Series[Series_Num1 - 1].Enabled = true;
                            chart8.Series[Series_Num1 - 1].Enabled = true;
                            chart9.Series[Series_Num1 - 1].Enabled = true;
                            chart10.Series[Series_Num1 - 1].Enabled = true;
                            chart11.Series[Series_Num1 - 1].Enabled = true;
                            chart12.Series[Series_Num1 - 1].Enabled = true;
                            chart13.Series[Series_Num1 - 1].Enabled = true;
                            chart14.Series[Series_Num1 - 1].Enabled = true;

                        }
                    }
                }
            }


            chart1.ChartAreas[0].AxisY.Maximum = Payload_Max_All;
            chart1.ChartAreas[0].AxisY.Minimum = Payload_Min_All;
            chart2.ChartAreas[0].AxisY.Maximum = Availability_Max_All;
            chart2.ChartAreas[0].AxisY.Minimum = Availability_Min_All;
            chart3.ChartAreas[0].AxisY.Maximum = DL_User_THR_Max_All;
            chart3.ChartAreas[0].AxisY.Minimum = DL_User_THR_Min_All;
            chart4.ChartAreas[0].AxisY.Maximum = DL_Cell_THR_Max_All;
            chart4.ChartAreas[0].AxisY.Minimum = DL_Cell_THR_Min_All;
            chart5.ChartAreas[0].AxisY.Maximum = UL_User_THR_Max_All;
            chart5.ChartAreas[0].AxisY.Minimum = UL_User_THR_Min_All;
            chart6.ChartAreas[0].AxisY.Maximum = Latency_Max_All;
            chart6.ChartAreas[0].AxisY.Minimum = Latency_Min_All;
            chart7.ChartAreas[0].AxisY.Maximum = Service_Max_All;
            chart7.ChartAreas[0].AxisY.Minimum = Service_Min_All;
            chart8.ChartAreas[0].AxisY.Maximum = RRC_Max_All;
            chart8.ChartAreas[0].AxisY.Minimum = RRC_Min_All;
            chart9.ChartAreas[0].AxisY.Maximum = ERAB_Setup_Max_All;
            chart9.ChartAreas[0].AxisY.Minimum = ERAB_Setup_Min_All;
            chart10.ChartAreas[0].AxisY.Maximum = ERAB_Drop_Max_All;
            chart10.ChartAreas[0].AxisY.Minimum = ERAB_Drop_Min_All;
            chart11.ChartAreas[0].AxisY.Maximum = Inter_Max_All;
            chart11.ChartAreas[0].AxisY.Minimum = Inter_Min_All;
            chart12.ChartAreas[0].AxisY.Maximum = Intra_Max_All;
            chart12.ChartAreas[0].AxisY.Minimum = Intra_Min_All;
            chart13.ChartAreas[0].AxisY.Maximum = PUCCH_Max_All;
            chart13.ChartAreas[0].AxisY.Minimum = PUCCH_Min_All;
            chart14.ChartAreas[0].AxisY.Maximum = PUSCH_Max_All;
            chart14.ChartAreas[0].AxisY.Minimum = PUSCH_Min_All;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;
                checkBox10.Checked = false;
                checkBox11.Checked = false;
                checkBox12.Checked = false;
                checkBox13.Checked = false;
                checkBox14.Checked = false;
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                checkBox1.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;
                checkBox10.Checked = false;
                checkBox11.Checked = false;
                checkBox12.Checked = false;
                checkBox13.Checked = false;
                checkBox14.Checked = false;
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked == true)
            {
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;
                checkBox10.Checked = false;
                checkBox11.Checked = false;
                checkBox12.Checked = false;
                checkBox13.Checked = false;
                checkBox14.Checked = false;
            }
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked == true)
            {
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;
                checkBox10.Checked = false;
                checkBox11.Checked = false;
                checkBox12.Checked = false;
                checkBox13.Checked = false;
                checkBox14.Checked = false;
            }
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked == true)
            {
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;
                checkBox10.Checked = false;
                checkBox11.Checked = false;
                checkBox12.Checked = false;
                checkBox13.Checked = false;
                checkBox14.Checked = false;
            }
        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox6.Checked == true)
            {
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox7.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;
                checkBox10.Checked = false;
                checkBox11.Checked = false;
                checkBox12.Checked = false;
                checkBox13.Checked = false;
                checkBox14.Checked = false;
            }
        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox7.Checked == true)
            {
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;
                checkBox10.Checked = false;
                checkBox11.Checked = false;
                checkBox12.Checked = false;
                checkBox13.Checked = false;
                checkBox14.Checked = false;
            }
        }

        private void checkBox8_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox8.Checked == true)
            {
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                checkBox9.Checked = false;
                checkBox10.Checked = false;
                checkBox11.Checked = false;
                checkBox12.Checked = false;
                checkBox13.Checked = false;
                checkBox14.Checked = false;
            }
        }

        private void checkBox9_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox9.Checked == true)
            {
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                checkBox8.Checked = false;
                checkBox10.Checked = false;
                checkBox11.Checked = false;
                checkBox12.Checked = false;
                checkBox13.Checked = false;
                checkBox14.Checked = false;
            }
        }

        private void checkBox10_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox10.Checked == true)
            {
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;
                checkBox11.Checked = false;
                checkBox12.Checked = false;
                checkBox13.Checked = false;
                checkBox14.Checked = false;
            }
        }

        private void checkBox11_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox11.Checked == true)
            {
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;
                checkBox10.Checked = false;
                checkBox12.Checked = false;
                checkBox13.Checked = false;
                checkBox14.Checked = false;
            }
        }

        private void checkBox12_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox12.Checked == true)
            {
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;
                checkBox10.Checked = false;
                checkBox11.Checked = false;
                checkBox13.Checked = false;
                checkBox14.Checked = false;
            }
        }

        private void checkBox13_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox13.Checked == true)
            {
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;
                checkBox10.Checked = false;
                checkBox11.Checked = false;
                checkBox12.Checked = false;
                checkBox14.Checked = false;
            }
        }
        private void checkBox14_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox14.Checked == true)
            {
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;
                checkBox10.Checked = false;
                checkBox11.Checked = false;
                checkBox12.Checked = false;
                checkBox13.Checked = false;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                chart1.SaveImage("Image.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
                Image img = Image.FromFile("Image.jpg");
                System.Windows.Forms.Clipboard.SetImage(img);
            }
            if (checkBox2.Checked == true)
            {
                chart2.SaveImage("Image.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
                Image img = Image.FromFile("Image.jpg");
                System.Windows.Forms.Clipboard.SetImage(img);
            }
            if (checkBox3.Checked == true)
            {
                chart3.SaveImage("Image.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
                Image img = Image.FromFile("Image.jpg");
                System.Windows.Forms.Clipboard.SetImage(img);
            }
            if (checkBox4.Checked == true)
            {
                chart4.SaveImage("Image.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
                Image img = Image.FromFile("Image.jpg");
                System.Windows.Forms.Clipboard.SetImage(img);
            }
            if (checkBox5.Checked == true)
            {
                chart5.SaveImage("Image.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
                Image img = Image.FromFile("Image.jpg");
                System.Windows.Forms.Clipboard.SetImage(img);
            }
            if (checkBox6.Checked == true)
            {
                chart6.SaveImage("Image.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
                Image img = Image.FromFile("Image.jpg");
                System.Windows.Forms.Clipboard.SetImage(img);
            }
            if (checkBox7.Checked == true)
            {
                chart7.SaveImage("Image.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
                Image img = Image.FromFile("Image.jpg");
                System.Windows.Forms.Clipboard.SetImage(img);
            }
            if (checkBox8.Checked == true)
            {
                chart8.SaveImage("Image.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
                Image img = Image.FromFile("Image.jpg");
                System.Windows.Forms.Clipboard.SetImage(img);
            }
            if (checkBox9.Checked == true)
            {
                chart9.SaveImage("Image.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
                Image img = Image.FromFile("Image.jpg");
                System.Windows.Forms.Clipboard.SetImage(img);
            }
            if (checkBox10.Checked == true)
            {
                chart10.SaveImage("Image.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
                Image img = Image.FromFile("Image.jpg");
                System.Windows.Forms.Clipboard.SetImage(img);
            }
            if (checkBox11.Checked == true)
            {
                chart11.SaveImage("Image.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
                Image img = Image.FromFile("Image.jpg");
                System.Windows.Forms.Clipboard.SetImage(img);
            }
            if (checkBox12.Checked == true)
            {
                chart12.SaveImage("Image.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
                Image img = Image.FromFile("Image.jpg");
                System.Windows.Forms.Clipboard.SetImage(img);
            }
            if (checkBox13.Checked == true)
            {
                chart13.SaveImage("Image.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
                Image img = Image.FromFile("Image.jpg");
                System.Windows.Forms.Clipboard.SetImage(img);
            }
            if (checkBox14.Checked == true)
            {
                chart14.SaveImage("Image.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
                Image img = Image.FromFile("Image.jpg");
                System.Windows.Forms.Clipboard.SetImage(img);
            }

        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            Max_X1 = dateTimePicker2.Value;
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            Min_X1 = dateTimePicker1.Value;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            chart1.ChartAreas[0].AxisX.Minimum = Min_X1.ToOADate() - 1;
            chart2.ChartAreas[0].AxisX.Minimum = Min_X1.ToOADate() - 1;
            chart3.ChartAreas[0].AxisX.Minimum = Min_X1.ToOADate() - 1;
            chart4.ChartAreas[0].AxisX.Minimum = Min_X1.ToOADate() - 1;
            chart5.ChartAreas[0].AxisX.Minimum = Min_X1.ToOADate() - 1;
            chart6.ChartAreas[0].AxisX.Minimum = Min_X1.ToOADate() - 1;
            chart7.ChartAreas[0].AxisX.Minimum = Min_X1.ToOADate() - 1;
            chart8.ChartAreas[0].AxisX.Minimum = Min_X1.ToOADate() - 1;
            chart9.ChartAreas[0].AxisX.Minimum = Min_X1.ToOADate() - 1;
            chart10.ChartAreas[0].AxisX.Minimum = Min_X1.ToOADate() - 1;
            chart11.ChartAreas[0].AxisX.Minimum = Min_X1.ToOADate() - 1;
            chart12.ChartAreas[0].AxisX.Minimum = Min_X1.ToOADate() - 1;
            chart13.ChartAreas[0].AxisX.Minimum = Min_X1.ToOADate() - 1;
            chart14.ChartAreas[0].AxisX.Minimum = Min_X1.ToOADate() - 1;


            chart1.ChartAreas[0].AxisX.Maximum = Max_X1.ToOADate() + 1;
            chart2.ChartAreas[0].AxisX.Maximum = Max_X1.ToOADate() + 1;
            chart3.ChartAreas[0].AxisX.Maximum = Max_X1.ToOADate() + 1;
            chart4.ChartAreas[0].AxisX.Maximum = Max_X1.ToOADate() + 1;
            chart5.ChartAreas[0].AxisX.Maximum = Max_X1.ToOADate() + 1;
            chart6.ChartAreas[0].AxisX.Maximum = Max_X1.ToOADate() + 1;
            chart7.ChartAreas[0].AxisX.Maximum = Max_X1.ToOADate() + 1;
            chart8.ChartAreas[0].AxisX.Maximum = Max_X1.ToOADate() + 1;
            chart9.ChartAreas[0].AxisX.Maximum = Max_X1.ToOADate() + 1;
            chart10.ChartAreas[0].AxisX.Maximum = Max_X1.ToOADate() + 1;
            chart11.ChartAreas[0].AxisX.Maximum = Max_X1.ToOADate() + 1;
            chart12.ChartAreas[0].AxisX.Maximum = Max_X1.ToOADate() + 1;
            chart13.ChartAreas[0].AxisX.Maximum = Max_X1.ToOADate() + 1;
            chart14.ChartAreas[0].AxisX.Maximum = Max_X1.ToOADate() + 1;




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
            chart5.ChartAreas[0].AxisX.Interval = day_interval;
            chart6.ChartAreas[0].AxisX.Interval = day_interval;
            chart7.ChartAreas[0].AxisX.Interval = day_interval;
            chart8.ChartAreas[0].AxisX.Interval = day_interval;
            chart9.ChartAreas[0].AxisX.Interval = day_interval;
            chart10.ChartAreas[0].AxisX.Interval = day_interval;
            chart11.ChartAreas[0].AxisX.Interval = day_interval;
            chart12.ChartAreas[0].AxisX.Interval = day_interval;
            chart13.ChartAreas[0].AxisX.Interval = day_interval;
            chart14.ChartAreas[0].AxisX.Interval = day_interval;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            chart1.ChartAreas[0].AxisX.Minimum = Min_X_Date.ToOADate() - 1;
            chart1.ChartAreas[0].AxisX.Maximum = Max_X_Date.ToOADate() + 1;
            chart2.ChartAreas[0].AxisX.Minimum = Min_X_Date.ToOADate() - 1;
            chart2.ChartAreas[0].AxisX.Maximum = Max_X_Date.ToOADate() + 1;
            chart3.ChartAreas[0].AxisX.Minimum = Min_X_Date.ToOADate() - 1;
            chart3.ChartAreas[0].AxisX.Maximum = Max_X_Date.ToOADate() + 1;
            chart4.ChartAreas[0].AxisX.Minimum = Min_X_Date.ToOADate() - 1;
            chart4.ChartAreas[0].AxisX.Maximum = Max_X_Date.ToOADate() + 1;
            chart5.ChartAreas[0].AxisX.Minimum = Min_X_Date.ToOADate() - 1;
            chart5.ChartAreas[0].AxisX.Maximum = Max_X_Date.ToOADate() + 1;
            chart6.ChartAreas[0].AxisX.Minimum = Min_X_Date.ToOADate() - 1;
            chart6.ChartAreas[0].AxisX.Maximum = Max_X_Date.ToOADate() + 1;
            chart7.ChartAreas[0].AxisX.Minimum = Min_X_Date.ToOADate() - 1;
            chart7.ChartAreas[0].AxisX.Maximum = Max_X_Date.ToOADate() + 1;
            chart8.ChartAreas[0].AxisX.Minimum = Min_X_Date.ToOADate() - 1;
            chart8.ChartAreas[0].AxisX.Maximum = Max_X_Date.ToOADate() + 1;
            chart9.ChartAreas[0].AxisX.Minimum = Min_X_Date.ToOADate() - 1;
            chart9.ChartAreas[0].AxisX.Maximum = Max_X_Date.ToOADate() + 1;
            chart10.ChartAreas[0].AxisX.Minimum = Min_X_Date.ToOADate() - 1;
            chart10.ChartAreas[0].AxisX.Maximum = Max_X_Date.ToOADate() + 1;
            chart11.ChartAreas[0].AxisX.Minimum = Min_X_Date.ToOADate() - 1;
            chart11.ChartAreas[0].AxisX.Maximum = Max_X_Date.ToOADate() + 1;
            chart12.ChartAreas[0].AxisX.Minimum = Min_X_Date.ToOADate() - 1;
            chart12.ChartAreas[0].AxisX.Maximum = Max_X_Date.ToOADate() + 1;
            chart13.ChartAreas[0].AxisX.Minimum = Min_X_Date.ToOADate() - 1;
            chart13.ChartAreas[0].AxisX.Maximum = Max_X_Date.ToOADate() + 1;
            chart14.ChartAreas[0].AxisX.Minimum = Min_X_Date.ToOADate() - 1;
            chart14.ChartAreas[0].AxisX.Maximum = Max_X_Date.ToOADate() + 1;



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
            chart5.ChartAreas[0].AxisX.Interval = day_interval;
            chart6.ChartAreas[0].AxisX.Interval = day_interval;
            chart7.ChartAreas[0].AxisX.Interval = day_interval;
            chart8.ChartAreas[0].AxisX.Interval = day_interval;
            chart9.ChartAreas[0].AxisX.Interval = day_interval;
            chart10.ChartAreas[0].AxisX.Interval = day_interval;
            chart11.ChartAreas[0].AxisX.Interval = day_interval;
            chart12.ChartAreas[0].AxisX.Interval = day_interval;
            chart13.ChartAreas[0].AxisX.Interval = day_interval;
            chart14.ChartAreas[0].AxisX.Interval = day_interval;


        }

        private void button3_Click(object sender, EventArgs e)
        {
            string Availability = "";
            if (Interval=="Daily")
            {
                Availability = "[Cell_Availability_Rate_Include_Blocking(Cell_EricLTE)]";
            }
            if (Interval == "BH")
            {
                Availability = "[Cell_Availability_Rate_Exclude_Blocking(Cell_EricLTE)]";
            }


            string Num = "";
            if (textBox1.Text != "")
            {
                Num = "TOP " + Convert.ToString(textBox1.Text);
            }

            string Availability_Condition = "";
            string Payload_Condition = "";
            if (textBox3.Text != "")
            {
                Availability_Condition = Availability+">=" + Convert.ToString(textBox3.Text);
            }

            if (textBox4.Text != "")
            {
                Payload_Condition = "[Total_Volume(UL+DL)(GB)(eNodeB_Eric)]>=" + Convert.ToString(textBox4.Text);
            }

            string Availability_Payload_Condition = "";
            if (Availability_Condition == "" && Payload_Condition != "")
            {
                Availability_Payload_Condition = Payload_Condition + " and ";
            }
            if (Availability_Condition != "" && Payload_Condition == "")
            {
                Availability_Payload_Condition = Availability_Condition + " and ";
            }
            if (Availability_Condition != "" && Payload_Condition != "")
            {
                Availability_Payload_Condition = Availability_Condition + " and " + Payload_Condition + " and ";
            }



            string KPI = "";
            string KPI_Name = "";
            string order = "";

            if (checkBox1.Checked == true)
            {
                KPI = "[Total_Volume(UL+DL)(GB)(eNodeB_Eric)]";
                KPI_Name = "Payload";
                order = "asc";
            }
            if (checkBox2.Checked == true)
            {
                KPI = Availability;
                KPI_Name = "Availability";
                order = "asc";
            }
            if (checkBox3.Checked == true)
            {
                KPI = "[Average_UE_DL_Throughput(Mbps)(eNodeB_Eric)]";
                KPI_Name = "DL_User_THR";
                order = "asc";
            }
            if (checkBox4.Checked == true)
            {
                KPI = "[Average_PDCP_Cell_Dl_Throughput(Kbps)(eNodeB_Eric)]";
                KPI_Name = "DL_Cell_THR";
                order = "asc";
            }

            if (checkBox5.Checked == true)
            {
                KPI = "[Average_UE_UL_Throughput(Mbps)(eNodeB_Eric)]";
                KPI_Name = "UL_User_THR";
                order = "asc";
            }
            if (checkBox6.Checked == true)
            {
                KPI = "[Average_UE_DL_Latency(ms)(eNodeB_Eric)]";
                KPI_Name = "Latency";
                order = "desc";
            }
            if (checkBox7.Checked == true)
            {
                KPI = "[LTE_Service_Success_Rate(eNodeB_Eric)]";
                KPI_Name = "Service_SR";
                order = "asc";
            }
            if (checkBox8.Checked == true)
            {
                KPI = "[RRC_Estab_Success_Rate(ReAtt)(EUCell_Eric)]";
                KPI_Name = "Connection_SR";
                order = "asc";
            }
            if (checkBox9.Checked == true)
            {
                KPI = "[E-RAB_Setup_SR_incl_added_New(EUCell_Eric)]";
                KPI_Name = "ERAB_Setup_SR";
                order = "asc";
            }
            if (checkBox10.Checked == true)
            {
                KPI = "[E_RAB_Drop_Rate(eNodeB_Eric)]";
                KPI_Name = "ERAB_Drop_Rate";
                order = "desc";
            }
            if (checkBox11.Checked == true)
            {
                KPI = "[InterF_Handover_Execution(eNodeB_Eric)]";
                KPI_Name = "Inter_Freq";
                order = "asc";
            }
            if (checkBox12.Checked == true)
            {
                KPI = "[IntraF_Handover_Execution(eNodeB_Eric)]";
                KPI_Name = "Intra_Freq";
                order = "asc";
            }
            if (checkBox13.Checked == true)
            {
                KPI = "[RSSI_PUCCH]";
                KPI_Name = "PUCCH_RSSI";
                order = "desc";
            }
            if (checkBox14.Checked == true)
            {
                KPI = "[RSSI_PUSCH]";
                KPI_Name = "PUSCH_RSSI";
                order = "desc";
            }



            if (checkBox1.Checked == false && checkBox2.Checked == false && checkBox3.Checked == false && checkBox4.Checked == false && checkBox5.Checked == false &&
                checkBox6.Checked == false && checkBox7.Checked == false && checkBox8.Checked == false && checkBox9.Checked == false && checkBox10.Checked == false &&
                checkBox11.Checked == false && checkBox12.Checked == false && checkBox13.Checked == false && checkBox14.Checked == false)
            {
                MessageBox.Show("Please Select a KPI");
            }





            string Date = "";
            for (int i = 0; i < listBox2.SelectedItems.Count; i++)
            {
                string Date_Name = "";
                if (Interval=="Daily")
                {
                    Date_Name = "Datetime";
                }
                if (Interval == "BH")
                {
                    Date_Name = "Day";
                }
                string Date1 = listBox2.SelectedItems[i].ToString();

                string Dtae_Disticts1 = Regex.Replace(Date1, "[^a-zA-Z0-9]", " ");
                Dtae_Disticts1 = Regex.Replace(Dtae_Disticts1, " {2,}", " ").Trim();
                string[] Dtae_Disticts = Dtae_Disticts1.Split(' ');

                string Date2 = Dtae_Disticts[2] + "-" + Dtae_Disticts[0] + "-" + Dtae_Disticts[1] + " 00:00:00.000";
             //   string Date2 = Date1.Substring(6, 4) + "-" + Date1.Substring(0, 2) + "-" + Date1.Substring(3, 2) + " 00:00:00.000";
                Date = Date + Date_Name + "='" + Date2 + "' or ";
            }

            if (Date != "")
            {
                Date = "(" + Date.Substring(0, Date.Length - 4) + ")";
            }



            if (Date == "")
            {
                MessageBox.Show("Please select the Date(s)");
            }



            string bigger_smaller = "";

            if (checkBox15.Checked == true)
            {
                bigger_smaller = ">=";
            }
            if (checkBox16.Checked == true)
            {
                bigger_smaller = "<=";
            }

            string KPI_THR = textBox5.Text;





            string Export = "";
            string RNCs = "";
            for (int i = 0; i < listBox1.SelectedItems.Count; i++)
            {
                RNCs = RNCs + "RNC='" + listBox1.SelectedItems[i].ToString() + "' or ";
            }

            if (RNCs != "")
            {
                RNCs = "(" + RNCs.Substring(0, RNCs.Length - 4) + ")";

                if (Interval=="Daily")
                {
                    Export = @" select " + Num + @" [Datetime], [Site],	[eNodeB],	[Vendor], [RNC] , [Band_Carrier], 
[Cell_Availability_Rate_Include_Blocking(Cell_EricLTE)] as 'Availability',   [Total_Volume(UL+DL)(GB)(eNodeB_Eric)] as 'Payload (GB)'," + KPI + " as '" + KPI_Name + @"'FROM
[LTE_RNC_TH_Last_Day] where " + Availability_Payload_Condition + KPI + bigger_smaller + KPI_THR + " and " + RNCs + " and " + Date + " order by " + KPI + " " + order;
                }
                if (Interval == "BH")
                {
                    Export = @" select " + Num + @" [Day], [Datetime], [Site],	[eNodeB],	[Vendor], [RNC] , [Band_Carrier], 
[Cell_Availability_Rate_Exclude_Blocking(Cell_EricLTE)] as 'Availability',   [Total_Volume(UL+DL)(GB)(eNodeB_Eric)] as 'Payload (GB)'," + KPI + " as '" + KPI_Name + @"'FROM
[LTE_RNC_TH_Last_Day_BH] where " + Availability_Payload_Condition + KPI + bigger_smaller + KPI_THR + " and " + RNCs + " and " + Date + " order by " + KPI + " " + order;
                }

            }

            if (RNCs == "")
            {
                MessageBox.Show("Please select the RNC(s)");
            }




            if (KPI_Name != "" && RNCs != "" && Date != "")
            {



                SqlCommand Data_Quary1 = new SqlCommand(Export, connection);
                Data_Quary1.CommandTimeout = 0;
                Data_Quary1.ExecuteNonQuery();
                DataTable Data_Table = new DataTable();
                SqlDataAdapter Date_Table1 = new SqlDataAdapter(Data_Quary1);
                Date_Table1.Fill(Data_Table);



                XLWorkbook wb = new XLWorkbook();
                wb.Worksheets.Add(Data_Table, "WPC");

                var saveFileDialog = new SaveFileDialog
                {
                    FileName = "WPC at " + KPI_Name,
                    Filter = "Excel files|*.xlsx",
                    Title = "Save an Excel File"
                };




                saveFileDialog.ShowDialog();

                if (!String.IsNullOrWhiteSpace(saveFileDialog.FileName))
                    wb.SaveAs(saveFileDialog.FileName);




                MessageBox.Show("Finished");

            }
        }



        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            

            if (checkBox15.Checked == false && checkBox16.Checked == false)
            {
                MessageBox.Show("Select the Signs (> or <)");
            }

        }

        private void checkBox17_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox17.Checked==true)
            {
                Interval = "Daily";
                checkBox18.Checked = false;
            }
        }

        private void checkBox18_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox18.Checked == true)
            {
                Interval = "BH";
                checkBox17.Checked = false;
            }
        }

        private void checkBox15_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox15.Checked == true)
            {
                checkBox16.Checked = false;
            }
        }

        private void checkBox16_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox16.Checked == true)
            {
                checkBox15.Checked = false;
            }
        }
    }
}
