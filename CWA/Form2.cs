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

    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }


        public Form1 form1;


        public Form2(Form form)
        {
            InitializeComponent();
            form1 = (Form1)form;
        }


        public string ConnectionString = "";
        public SqlConnection connection = new SqlConnection();
        public double[,] TH_2G = new double[6, 10];
        public double[,] TH_3G_CS = new double[6, 5];
        public double[,] TH_3G_PS = new double[6, 14];
        public double[,] TH_4G = new double[6, 11];

        //public string Server_Name = @"NAKPRG-NB1243\" + "AHMAD";
        //public string DataBase_Name = "Contract";

        public string Server_Name = "PERFORMANCEDB01";
        public string DataBase_Name = "Performance_NAK";

        public DataTable ARAS_Table = new DataTable();
        public DataTable ARAS_Data = new DataTable();
        public DataTable BASE_Table = new DataTable();
        public DataTable BASE_Data = new DataTable();
        public DataTable dtUnion_2G = new DataTable();
        public DataTable Table_2G = new DataTable();
        public DataTable dtUnion_3G_CS = new DataTable();
        public DataTable Table_3G_CS = new DataTable();
        public DataTable dtUnion_3G_PS = new DataTable();
        public DataTable Table_3G_PS = new DataTable();
        public DataTable dtUnion_4G = new DataTable();
        public DataTable Table_4G = new DataTable();

        public DateTime Date_of_WPC = DateTime.Today;
        public DateTime Date_of_WPC_7 = DateTime.Today;


        public DataTable Table_2G_WPC = new DataTable();
        public DataTable Table_3G_CS_WPC = new DataTable();
        public DataTable Table_3G_PS_WPC = new DataTable();
        public DataTable Table_4G_WPC = new DataTable();

        public string Technology = "2G";


        public DateTime Date_List = DateTime.Now;
        public string Interval = "Daily";
        public string sign = "<";
        public int Ericsson_Count = 0;
        public int Huawei_Count = 0;
        public int Nokia_Count = 0;



        private void Form2_Load(object sender, EventArgs e)
        {
            //ConnectionString = @"Server=" + Server_Name + "; Database=" + DataBase_Name + "; Trusted_Connection=True;";
            ConnectionString = @"Server=" + Server_Name + "; Database=" + DataBase_Name + "; User ID=cwpcApp;Password=cwpcApp@830625#Ahmad";
            connection = new SqlConnection(ConnectionString);
            connection.Open();


            string userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            string currentUser = userName.Substring(8, userName.Length - 8);

            string[] authorizedUsers = new string[]
{
                           "ahmad.alikhani"

};


            
            //            if (authorizedUsers.Contains(currentUser.ToLower()))
            //            {
            //                string Authorized = "OK";
            //            }
            //            else
            //            {
            //                MessageBox.Show("Limited Access! Need Authorization by Admin");
            //                this.Close();
            //            }



            //var excelApplication = new Excel.Application();

            //var excelWorkBook = excelApplication.Application.Workbooks.Add(Type.Missing);

            //excelApplication.Cells[1, 1] = currentUser;
            //excelApplication.Cells[1, 2] = DateTime.Now;

            //string CR_PATH = string.Format(@"\\172.26.7.65\Performance_Share\Old\Performance monitoring Process\Softwares\WPC_Quary.xlsx");
            //excelApplication.ActiveWorkbook.SaveCopyAs(CR_PATH);
            ////excelApplication.ActiveWorkbook.SaveCopyAs(@"C:\test.xlsx");

            //excelApplication.ActiveWorkbook.Saved = true;

            //// Close the Excel Application
            //excelApplication.Quit();




            // Table of WPC 2G
            Table_2G_WPC.Columns.Add("Contractor", typeof(String));
            Table_2G_WPC.Columns.Add("Province", typeof(String));
            Table_2G_WPC.Columns.Add("Vendor", typeof(String));
            Table_2G_WPC.Columns.Add("BSC", typeof(String));
            Table_2G_WPC.Columns.Add("Site", typeof(String));
            Table_2G_WPC.Columns.Add("Cell", typeof(String));
            Table_2G_WPC.Columns.Add("Coverage Type", typeof(String));
            Table_2G_WPC.Columns.Add("Date", typeof(DateTime));
            Table_2G_WPC.Columns.Add("Availability", typeof(double));
            Table_2G_WPC.Columns.Add("Daily TCH Traffic (Erlang)", typeof(double));
            // Table_2G_WPC.Columns.Add("Baseline KPI Value", typeof(double));
            Table_2G_WPC.Columns.Add("KPI Value", typeof(double));
            Table_2G_WPC.Columns.Add("KPI Name", typeof(String));
            Table_2G_WPC.Columns.Add("Number of Worst KPIs", typeof(string));
            Table_2G_WPC.Columns.Add("Value at Date - 1", typeof(double));
            Table_2G_WPC.Columns.Add("Value at Date - 2", typeof(double));
            Table_2G_WPC.Columns.Add("Value at Date - 3", typeof(double));
            Table_2G_WPC.Columns.Add("Value at Date - 4", typeof(double));
            Table_2G_WPC.Columns.Add("Value at Date - 5", typeof(double));
            Table_2G_WPC.Columns.Add("Value at Date - 6", typeof(double));
            Table_2G_WPC.Columns.Add("Value at Date - 7", typeof(double));




            // Settings of Thresholds
            // Ericsson - City
            // Ericsson - Rural
            //Huawei - City
            // Huawei - Rural
            //Nokia - City
            //Nokia - Rural
            //CSSR
            TH_2G[0, 0] = 93.989;
            TH_2G[1, 0] = 89.28955;
            TH_2G[2, 0] = 96.827;
            TH_2G[3, 0] = 91.98565;
            TH_2G[4, 0] = 95.034;
            TH_2G[5, 0] = 90.2823;
            //OHSR
            TH_2G[0, 1] = 93.811;
            TH_2G[1, 1] = 89.12045;
            TH_2G[2, 1] = 89.045;
            TH_2G[3, 1] = 84.59275;
            TH_2G[4, 1] = 80.511;
            TH_2G[5, 1] = 76.48545;
            //CDR
            TH_2G[0, 2] = 3.66;
            TH_2G[1, 2] = 3.843;
            TH_2G[2, 2] = 1.961;
            TH_2G[3, 2] = 2.05905;
            TH_2G[4, 2] = 5.263;
            TH_2G[5, 2] = 5.52615;
            //TCH ASFR
            TH_2G[0, 3] = 1.396;
            TH_2G[1, 3] = 1.4658;
            TH_2G[2, 3] = 2.103;
            TH_2G[3, 3] = 2.20815;
            TH_2G[4, 3] = 3.869;
            TH_2G[5, 3] = 4.06245;
            //RXDL
            TH_2G[0, 4] = 88.533;
            TH_2G[1, 4] = 84.10635;
            TH_2G[2, 4] = 91.46;
            TH_2G[3, 4] = 86.887;
            TH_2G[4, 4] = 84.535;
            TH_2G[5, 4] = 80.30825;
            //RXUL
            TH_2G[0, 5] = 92.891;
            TH_2G[1, 5] = 88.24645;
            TH_2G[2, 5] = 93.782;
            TH_2G[3, 5] = 89.0929;
            TH_2G[4, 5] = 87.287;
            TH_2G[5, 5] = 82.92265;
            //SDCCH CONG
            TH_2G[0, 6] = 0.515;
            TH_2G[1, 6] = 0.54075;
            TH_2G[2, 6] = 0.825;
            TH_2G[3, 6] = 0.86625;
            TH_2G[4, 6] = 0.313;
            TH_2G[5, 6] = 0.32865;
            //SDCCH SR
            TH_2G[0, 7] = 86.294;
            TH_2G[1, 7] = 81.9793;
            TH_2G[2, 7] = 90.151;
            TH_2G[3, 7] = 85.64345;
            TH_2G[4, 7] = 88.808;
            TH_2G[5, 7] = 84.3676;
            //SDDCCH DROP
            TH_2G[0, 8] = 4.597;
            TH_2G[1, 8] = 4.82685;
            TH_2G[2, 8] = 0.482;
            TH_2G[3, 8] = 0.5061;
            TH_2G[4, 8] = 1.351;
            TH_2G[5, 8] = 1.41855;
            //IHSR
            TH_2G[0, 9] = 92.414;
            TH_2G[1, 9] = 87.7933;
            TH_2G[2, 9] = 88.787;
            TH_2G[3, 9] = 84.34765;
            TH_2G[4, 9] = 86.047;
            TH_2G[5, 9] = 81.74465;






            // Table of WPC 3G_CS
            Table_3G_CS_WPC.Columns.Add("Contractor", typeof(String));
            Table_3G_CS_WPC.Columns.Add("Province", typeof(String));
            Table_3G_CS_WPC.Columns.Add("Vendor", typeof(String));
            Table_3G_CS_WPC.Columns.Add("RNC", typeof(String));
            Table_3G_CS_WPC.Columns.Add("Site", typeof(String));
            Table_3G_CS_WPC.Columns.Add("Cell", typeof(String));
            Table_3G_CS_WPC.Columns.Add("Coverage Type", typeof(String));
            Table_3G_CS_WPC.Columns.Add("Date", typeof(DateTime));
            Table_3G_CS_WPC.Columns.Add("Availability", typeof(double));
            Table_3G_CS_WPC.Columns.Add("Daily CS Traffic (Erlang)", typeof(double));
            // Table_3G_CS_WPC.Columns.Add("Baseline KPI Value", typeof(double));
            Table_3G_CS_WPC.Columns.Add("KPI Value", typeof(double));
            Table_3G_CS_WPC.Columns.Add("KPI Name", typeof(String));
            Table_3G_CS_WPC.Columns.Add("Number of Worst KPIs", typeof(String));
            Table_3G_CS_WPC.Columns.Add("Value at Date - 1", typeof(double));
            Table_3G_CS_WPC.Columns.Add("Value at Date - 2", typeof(double));
            Table_3G_CS_WPC.Columns.Add("Value at Date - 3", typeof(double));
            Table_3G_CS_WPC.Columns.Add("Value at Date - 4", typeof(double));
            Table_3G_CS_WPC.Columns.Add("Value at Date - 5", typeof(double));
            Table_3G_CS_WPC.Columns.Add("Value at Date - 6", typeof(double));
            Table_3G_CS_WPC.Columns.Add("Value at Date - 7", typeof(double));

            // CS_RAB_Establish
            TH_3G_CS[0, 0] = 99.725;
            TH_3G_CS[1, 0] = 94.73875;
            TH_3G_CS[2, 0] = 99.827;
            TH_3G_CS[3, 0] = 94.83565;
            TH_3G_CS[4, 0] = 99.407;
            TH_3G_CS[5, 0] = 94.43665;

            // CS_IRAT_HO_SR
            TH_3G_CS[0, 1] = 60;
            TH_3G_CS[1, 1] = 57;
            TH_3G_CS[2, 1] = 89.759;
            TH_3G_CS[3, 1] = 85.27105;
            TH_3G_CS[4, 1] = 88.889;
            TH_3G_CS[5, 1] = 84.44455;

            // CS_Drop_Rate
            TH_3G_CS[0, 2] = 0.687;
            TH_3G_CS[1, 2] = 0.72135;
            TH_3G_CS[2, 2] = 0.623;
            TH_3G_CS[3, 2] = 0.65415;
            TH_3G_CS[4, 2] = 0.84;
            TH_3G_CS[5, 2] = 0.882;

            // Soft_HO_SR
            TH_3G_CS[0, 3] = 99.752;
            TH_3G_CS[1, 3] = 94.7644;
            TH_3G_CS[2, 3] = 99.777;
            TH_3G_CS[3, 3] = 94.78815;
            TH_3G_CS[4, 3] = 99.594;
            TH_3G_CS[5, 3] = 94.6143;

            // CS_RRC_SR
            TH_3G_CS[0, 4] = 99.57;
            TH_3G_CS[1, 4] = 94.5915;
            TH_3G_CS[2, 4] = 99.521;
            TH_3G_CS[3, 4] = 94.54495;
            TH_3G_CS[4, 4] = 99.355;
            TH_3G_CS[5, 4] = 94.38725;




            // Table of WPC 3G_PS
            Table_3G_PS_WPC.Columns.Add("Contractor", typeof(String));
            Table_3G_PS_WPC.Columns.Add("Province", typeof(String));
            Table_3G_PS_WPC.Columns.Add("Vendor", typeof(String));
            Table_3G_PS_WPC.Columns.Add("RNC", typeof(String));
            Table_3G_PS_WPC.Columns.Add("Site", typeof(String));
            Table_3G_PS_WPC.Columns.Add("Cell", typeof(String));
            Table_3G_PS_WPC.Columns.Add("Coverage Type", typeof(String));
            Table_3G_PS_WPC.Columns.Add("Date", typeof(DateTime));
            Table_3G_PS_WPC.Columns.Add("Availability", typeof(double));
            Table_3G_PS_WPC.Columns.Add("Daily PS Traffic (GB)", typeof(double));
            // Table_3G_PS_WPC.Columns.Add("Baseline KPI Value", typeof(double));
            Table_3G_PS_WPC.Columns.Add("KPI Value", typeof(double));
            Table_3G_PS_WPC.Columns.Add("KPI Name", typeof(String));
            Table_3G_PS_WPC.Columns.Add("Number of Worst KPIs", typeof(String));
            Table_3G_PS_WPC.Columns.Add("Value at Date - 1", typeof(double));
            Table_3G_PS_WPC.Columns.Add("Value at Date - 2", typeof(double));
            Table_3G_PS_WPC.Columns.Add("Value at Date - 3", typeof(double));
            Table_3G_PS_WPC.Columns.Add("Value at Date - 4", typeof(double));
            Table_3G_PS_WPC.Columns.Add("Value at Date - 5", typeof(double));
            Table_3G_PS_WPC.Columns.Add("Value at Date - 6", typeof(double));
            Table_3G_PS_WPC.Columns.Add("Value at Date - 7", typeof(double));


            //HSDPA_SR
            TH_3G_PS[0, 0] = 97.812;
            TH_3G_PS[1, 0] = 92.9214;
            TH_3G_PS[2, 0] = 99.63;
            TH_3G_PS[3, 0] = 94.6485;
            TH_3G_PS[4, 0] = 97.106;
            TH_3G_PS[5, 0] = 92.2507;

            //HSUPA_SR
            TH_3G_PS[0, 1] = 97.128;
            TH_3G_PS[1, 1] = 92.2716;
            TH_3G_PS[2, 1] = 99.628;
            TH_3G_PS[3, 1] = 94.6466;
            TH_3G_PS[4, 1] = 96.325;
            TH_3G_PS[5, 1] = 91.50875;

            //UL_User_THR
            TH_3G_PS[0, 2] = 122.828;
            TH_3G_PS[1, 2] = 116.6866;
            TH_3G_PS[2, 2] = 150.299;
            TH_3G_PS[3, 2] = 142.78405;
            TH_3G_PS[4, 2] = 126.347;
            TH_3G_PS[5, 2] = 120.02965;

            //DL_User_THR
            TH_3G_PS[0, 3] = 1.056;
            TH_3G_PS[1, 3] = 1.0032;
            TH_3G_PS[2, 3] = 1.486;
            TH_3G_PS[3, 3] = 1.4117;
            TH_3G_PS[4, 3] = 0;
            TH_3G_PS[5, 3] = 0;

            //HSDPA_Drop_Rate
            TH_3G_PS[0, 4] = 2.106;
            TH_3G_PS[1, 4] = 2.2113;
            TH_3G_PS[2, 4] = 0.168;
            TH_3G_PS[3, 4] = 0.1764;
            TH_3G_PS[4, 4] = 1.234;
            TH_3G_PS[5, 4] = 1.2957;

            // HSUPA_Drop_Rate
            TH_3G_PS[0, 5] = 2.133;
            TH_3G_PS[1, 5] = 2.23965;
            TH_3G_PS[2, 5] = 0.225;
            TH_3G_PS[3, 5] = 0.23625;
            TH_3G_PS[4, 5] = 1.079;
            TH_3G_PS[5, 5] = 1.13295;

            //MultiRAB_SR
            TH_3G_PS[0, 6] = 99.153;
            TH_3G_PS[1, 6] = 94.19535;
            TH_3G_PS[2, 6] = 99.756;
            TH_3G_PS[3, 6] = 94.7682;
            TH_3G_PS[4, 6] = 98.808;
            TH_3G_PS[5, 6] = 93.8676;

            // PS_RRC_SR
            TH_3G_PS[0, 7] = 99.066;
            TH_3G_PS[1, 7] = 94.1127;
            TH_3G_PS[2, 7] = 99.514;
            TH_3G_PS[3, 7] = 94.5383;
            TH_3G_PS[4, 7] = 97.857;
            TH_3G_PS[5, 7] = 92.96415;

            //Ps_RAB_Establish
            TH_3G_PS[0, 8] = 97.805;
            TH_3G_PS[1, 8] = 92.91475;
            TH_3G_PS[2, 8] = 99.563;
            TH_3G_PS[3, 8] = 94.58485;
            TH_3G_PS[4, 8] = 98.868;
            TH_3G_PS[5, 8] = 93.9246;

            // PS_MultiRAB_Establish
            TH_3G_PS[0, 9] = 97.28;
            TH_3G_PS[1, 9] = 92.416;
            TH_3G_PS[2, 9] = 98.892;
            TH_3G_PS[3, 9] = 93.9474;
            TH_3G_PS[4, 9] = 97.282;
            TH_3G_PS[5, 9] = 92.4179;

            //PS_Drop_Rate
            TH_3G_PS[0, 10] = 2.098;
            TH_3G_PS[1, 10] = 2.2029;
            TH_3G_PS[2, 10] = 0.213;
            TH_3G_PS[3, 10] = 0.22365;
            TH_3G_PS[4, 10] = 1.613;
            TH_3G_PS[5, 10] = 1.69365;

            // HSDPA_Cell_Change_SR
            TH_3G_PS[0, 11] = 99.745;
            TH_3G_PS[1, 11] = 94.75775;
            TH_3G_PS[2, 11] = 99.619;
            TH_3G_PS[3, 11] = 94.63805;
            TH_3G_PS[4, 11] = 99.172;
            TH_3G_PS[5, 11] = 94.2134;

            // HS_Share_Payload
            TH_3G_PS[0, 12] = 99.593;
            TH_3G_PS[1, 12] = 94.61335;
            TH_3G_PS[2, 12] = 99.053;
            TH_3G_PS[3, 12] = 94.10035;
            TH_3G_PS[4, 12] = 99.617;
            TH_3G_PS[5, 12] = 94.63615;

            //DL_Cell_THR
            TH_3G_PS[0, 13] = 2.004;
            TH_3G_PS[1, 13] = 1.9038;
            TH_3G_PS[2, 13] = 2.045;
            TH_3G_PS[3, 13] = 1.94275;
            TH_3G_PS[4, 13] = 1.979;
            TH_3G_PS[5, 13] = 1.88005;






            // Table of WPC 4G
            Table_4G_WPC.Columns.Add("Contractor", typeof(String));
            Table_4G_WPC.Columns.Add("Province", typeof(String));
            Table_4G_WPC.Columns.Add("Vendor", typeof(String));
            // Table_4G_WPC.Columns.Add("BSC", typeof(String));
            Table_4G_WPC.Columns.Add("Site", typeof(String));
            Table_4G_WPC.Columns.Add("Cell", typeof(String));
            Table_4G_WPC.Columns.Add("Coverage Type", typeof(String));
            Table_4G_WPC.Columns.Add("Date", typeof(DateTime));
            Table_4G_WPC.Columns.Add("Availability", typeof(double));
            Table_4G_WPC.Columns.Add("Daily Data Traffic (GB)", typeof(double));
            // Table_4G_WPC.Columns.Add("Baseline KPI Value", typeof(double));
            Table_4G_WPC.Columns.Add("KPI Value", typeof(double));
            Table_4G_WPC.Columns.Add("KPI Name", typeof(String));
            Table_4G_WPC.Columns.Add("Number of Worst KPIs", typeof(String));
            Table_4G_WPC.Columns.Add("Value at Date - 1", typeof(double));
            Table_4G_WPC.Columns.Add("Value at Date - 2", typeof(double));
            Table_4G_WPC.Columns.Add("Value at Date - 3", typeof(double));
            Table_4G_WPC.Columns.Add("Value at Date - 4", typeof(double));
            Table_4G_WPC.Columns.Add("Value at Date - 5", typeof(double));
            Table_4G_WPC.Columns.Add("Value at Date - 6", typeof(double));
            Table_4G_WPC.Columns.Add("Value at Date - 7", typeof(double));







            //RRC_Connection_SR
            TH_4G[0, 0] = 99.758;
            TH_4G[1, 0] = 94.7701;
            TH_4G[2, 0] = 99.894;
            TH_4G[3, 0] = 94.8993;
            TH_4G[4, 0] = 99.747;
            TH_4G[5, 0] = 94.75965;
            //ERAB_SR_Initial 
            TH_4G[0, 1] = 99.192;
            TH_4G[1, 1] = 94.2324;
            TH_4G[2, 1] = 99.696;
            TH_4G[3, 1] = 94.7112;
            TH_4G[4, 1] = 99.282;
            TH_4G[5, 1] = 94.3179;
            //ERAB_SR_Added 
            TH_4G[0, 2] = 99.525;
            TH_4G[1, 2] = 94.54875;
            TH_4G[2, 2] = 99.697;
            TH_4G[3, 2] = 94.71215;
            TH_4G[4, 2] = 99.281;
            TH_4G[5, 2] = 94.31695;
            //DL_THR  
            TH_4G[0, 3] = 1.879;
            TH_4G[1, 3] = 1.78505;
            TH_4G[2, 3] = 4.057;
            TH_4G[3, 3] = 3.85415;
            TH_4G[4, 3] = 4.465;
            TH_4G[5, 3] = 4.24175;
            //UL_THR 
            TH_4G[0, 4] = 0.248;
            TH_4G[1, 4] = 0.2356;
            TH_4G[2, 4] = 0.91;
            TH_4G[3, 4] = 0.8645;
            TH_4G[4, 4] = 0.476;
            TH_4G[5, 4] = 0.4522;
            //HO_SR   
            TH_4G[0, 5] = 94.58;
            TH_4G[1, 5] = 89.851;
            TH_4G[2, 5] = 97.807;
            TH_4G[3, 5] = 92.91665;
            TH_4G[4, 5] = 91.672;
            TH_4G[5, 5] = 87.0884;
            //ERAB_Drop_Rate 
            TH_4G[0, 6] = 0.443;
            TH_4G[1, 6] = 0.46515;
            TH_4G[2, 6] = 0.491;
            TH_4G[3, 6] = 0.51555;
            TH_4G[4, 6] = 0.854;
            TH_4G[5, 6] = 0.8967;
            //S1_Signalling_SR  
            TH_4G[0, 7] = 99.831;
            TH_4G[1, 7] = 94.83945;
            TH_4G[2, 7] = 99.926;
            TH_4G[3, 7] = 94.9297;
            TH_4G[4, 7] = 99.901;
            TH_4G[5, 7] = 94.90595;
            //Inter_Freq_SR 
            TH_4G[0, 8] = 69.456;
            TH_4G[1, 8] = 65.9832;
            TH_4G[2, 8] = 94.872;
            TH_4G[3, 8] = 90.1284;
            TH_4G[4, 8] = 40;
            TH_4G[5, 8] = 38;
            //Intra_Freq_SR 
            TH_4G[0, 9] = 94.739;
            TH_4G[1, 9] = 90.00205;
            TH_4G[2, 9] = 97.638;
            TH_4G[3, 9] = 92.7561;
            TH_4G[4, 9] = 94.934;
            TH_4G[5, 9] = 90.1873;
            //UL_Packet_Loss
            TH_4G[0, 10] = 0.878;
            TH_4G[1, 10] = 0.9219;
            TH_4G[2, 10] = 0.011;
            TH_4G[3, 10] = 0.01155;
            TH_4G[4, 10] = 0.591;
            TH_4G[5, 10] = 0.62055;






        }

        private void button1_Click(object sender, EventArgs e)
        {

            //Quary of Select ARTAS File
            string ARAS_Data = @"select * from [City_NONECity]";
            SqlCommand ARAS_Quary = new SqlCommand(ARAS_Data, connection);
            ARAS_Quary.ExecuteNonQuery();
            ARAS_Table = new DataTable();
            SqlDataAdapter dataAdapter_ARAS_Table = new SqlDataAdapter(ARAS_Quary);
            dataAdapter_ARAS_Table.Fill(ARAS_Table);

            //Quary of Select Base Line
            //string BASE_Data = @"select * from [Base_2G_Daily$]";
            //SqlCommand BASE_Quary = new SqlCommand(BASE_Data, connection);
            //BASE_Quary.ExecuteNonQuery();
            //BASE_Table = new DataTable();
            //SqlDataAdapter dataAdapter_BASE_Table = new SqlDataAdapter(BASE_Quary);
            //dataAdapter_BASE_Table.Fill(BASE_Table);



            // Quary of Select Data in 2G
            if (checkBox1.Checked == true)
            {
                string Ericsson_2G_TBL = @"select [Date], [BSC], [Cell], [TCH_Traffic] as 'TCH_Traffic_Daily (Erlang)', [CSSR_MCI] as'CSSR', [TCH_Congestion] as 'TCH Congestion Rate', [SDCCH_Congestion] as 'SDCCH Congestion Rate', [TCH_Assign_Fail_Rate(NAK)(Eric_CELL)] as 'TCH Assign Fail Rate', 
[SDCCH_Access_Succ_Rate] as 'SDCCH Access Success Rate', [SDCCH_Drop_Rate] as 'SDCCH Drop Rate',  [CDR(not Affected by incoming Handovers from 3G)(Eric_CELL)] as 'Voice Drop Rate', 
[IHSR] as 'IHSR', [OHSR] as 'OHSR', [RxQual_DL] as 'RxQuality_DL', [RxQual_UL] as 'RxQuality_UL', [TCH_Availability] as 'TCH Availability', [AMRHR_USAGE] as 'AMRHR Usage'
from [dbo].[CC2_Ericsson_Cell_Daily] where 
(substring([Cell],1,2)='KJ' or substring([Cell],1,2)='CH' or substring([Cell],1,2)='AS' or substring([Cell],1,2)='GL' or 
                                                                                  substring([Cell],1,2)='GN'  or substring([Cell],1,2)='KM' or substring([Cell],1,2)='KH' or substring([Cell],1,2)='KZ' or substring([Cell],1,2)='MA'
																				  or substring([Cell],1,2)='SM' or substring([Cell],1,2)='TH' or substring([Cell],1,2)='AG' or substring([Cell],1,2)='YZ')
and ([CSSR_MCI]<93.989 or [OHSR]<93.811 or [CDR(not Affected by incoming Handovers from 3G)(Eric_CELL)]>3.66 or [TCH_Assign_Fail_Rate(NAK)(Eric_CELL)]>1.396 or [RxQual_DL]<88.533 or [RxQual_UL]<92.891 or [SDCCH_Congestion]>0.515 or [SDCCH_Access_Succ_Rate]<86.294 or [SDCCH_Drop_Rate]>4.597 or [IHSR]<92.414) and Date  ='" + Date_of_WPC + "'";



                //                string Ericsson_2G_TBL = @"select [Date], [BSC], [Cell], [TCH_Traffic] as 'TCH_Traffic_Daily (Erlang)', [CSSR_MCI] as'CSSR', [TCH_Congestion] as 'TCH Congestion Rate', [SDCCH_Congestion] as 'SDCCH Congestion Rate', [TCH_Assign_Fail_Rate(NAK)(Eric_CELL)] as 'TCH Assign Fail Rate', 
                //[SDCCH_Access_Succ_Rate] as 'SDCCH Access Success Rate', [SDCCH_Drop_Rate] as 'SDCCH Drop Rate',  [CDR(not Affected by incoming Handovers from 3G)(Eric_CELL)] as 'Voice Drop Rate', 
                //[IHSR] as 'IHSR', [OHSR] as 'OHSR', [RxQual_DL] as 'RxQuality_DL', [RxQual_UL] as 'RxQuality_UL', [TCH_Availability] as 'TCH Availability', [AMRHR_USAGE] as 'AMRHR Usage'
                //from [dbo].[CC2_Ericsson_Cell_Daily] where 
                //(substring([Cell],1,2)='TH')
                //and ([CSSR_MCI]<93.989 or [OHSR]<93.811 or [CDR(not Affected by incoming Handovers from 3G)(Eric_CELL)]>3.66 or [TCH_Assign_Fail_Rate(NAK)(Eric_CELL)]>1.396 or [RxQual_DL]<88.533 or [RxQual_UL]<92.891 or [SDCCH_Congestion]>0.515 or [SDCCH_Access_Succ_Rate]<86.294 or [SDCCH_Drop_Rate]>4.597 or [IHSR]<92.414) and Date  ='" + Date_of_WPC + "'";



                SqlCommand Ericsson_2G_TBL_Quary = new SqlCommand(Ericsson_2G_TBL, connection);
                Ericsson_2G_TBL_Quary.ExecuteNonQuery();
                DataTable Ericsson_2G_Table = new DataTable();
                SqlDataAdapter dataAdapter_Ericsson_2G = new SqlDataAdapter(Ericsson_2G_TBL_Quary);
                dataAdapter_Ericsson_2G.Fill(Ericsson_2G_Table);




                string Huawei_2G_TBL = @"select[Date], [BSC], [Cell], [TCH_Traffic], [CSSR3], [TCH_Cong], [SDCCH_Congestion_Rate], [TCH_Assignment_FR], 
            [SDCCH_Access_Success_Rate2], [SDCCH_Drop_Rate], [CDR3], 
            [IHSR2], [OHSR2], [RX_QUALITTY_DL_NEW], [RX_QUALITTY_UL_NEW], [TCH_Availability], [AMRHR_USAGE]
                    from[dbo].[CC2_Huawei_Cell_Daily] where
                   (substring([Cell],1,2)='KJ' or substring([Cell],1,2)='CH' or substring([Cell],1,2)='AS' or substring([Cell],1,2)='GL' or
                                                                                                     substring([Cell],1,2)='GN'  or substring([Cell],1,2)='KM' or substring([Cell],1,2)='KH' or substring([Cell],1,2)='KZ' or substring([Cell],1,2)='MA'
            																				  or substring([Cell],1,2)='SM' or substring([Cell],1,2)='TH' or substring([Cell],1,2)='AG' or substring([Cell],1,2)='YZ')
            and ([CSSR3]<96.827 or [OHSR2]<89.045 or [CDR3]>1.961 or [TCH_Assignment_FR]>2.103 or [RX_QUALITTY_DL_NEW]<91.46 or [RX_QUALITTY_UL_NEW]<93.782 or [SDCCH_Congestion_Rate]>0.825 or [SDCCH_Access_Success_Rate2]<90.151 or [SDCCH_Drop_Rate]>0.482 or [IHSR2]<88.787) and Date  ='" + Date_of_WPC + "'";



                //    string Huawei_2G_TBL = @"select[Date], [BSC], [Cell], [TCH_Traffic], [CSSR3], [TCH_Cong], [SDCCH_Congestion_Rate], [TCH_Assignment_FR], 
                //[SDCCH_Access_Success_Rate2], [SDCCH_Drop_Rate], [CDR3], 
                //[IHSR2], [OHSR2], [RX_QUALITTY_DL_NEW], [RX_QUALITTY_UL_NEW], [TCH_Availability], [AMRHR_USAGE]
                //        from[dbo].[CC2_Huawei_Cell_Daily] where
                //       (substring([Cell],1,2)='TH')
                //and ([CSSR3]<96.827 or [OHSR2]<89.045 or [CDR3]>1.961 or [TCH_Assignment_FR]>2.103 or [RX_QUALITTY_DL_NEW]<91.46 or [RX_QUALITTY_UL_NEW]<93.782 or [SDCCH_Congestion_Rate]>0.825 or [SDCCH_Access_Success_Rate2]<90.151 or [SDCCH_Drop_Rate]>0.482 or [IHSR2]<88.787) and Date  ='" + Date_of_WPC + "'";


                SqlCommand Huawei_2G_TBL_Quary = new SqlCommand(Huawei_2G_TBL, connection);
                Huawei_2G_TBL_Quary.ExecuteNonQuery();
                DataTable Huawei_2G_Table = new DataTable();
                SqlDataAdapter dataAdapter_Huawei_2G = new SqlDataAdapter(Huawei_2G_TBL_Quary);
                dataAdapter_Huawei_2G.Fill(Huawei_2G_Table);




                string Nokia_2G_TBL = @"select[Date], [BSC], [SEG], [TCH_Traffic], [CSSR_MCI], [TCH_Cong_Rate], [SDCCH_Congestion_Rate], [TCH_Assignment_FR], 
            [SDCCH_Access_Success_Rate], [SDCCH_Drop_Rate], [CDR(including_CS_IRAT_handovers_3G_to2G)(Nokia_SEG)], 
            [IHSR], [OHSR], [RxQuality_DL], [RxQuality_UL], [TCH_Availability], [AMRHR_USAGE]
                    from[dbo].[CC2_Nokia_Cell_Daily] where
                   (substring([SEG],1,2)='KJ' or substring([SEG],1,2)='CH' or substring([SEG],1,2)='AS' or substring([SEG],1,2)='GL' or
                                                                                                     substring([SEG],1,2)='GN'  or substring([SEG],1,2)='KM' or substring([SEG],1,2)='KH' or substring([SEG],1,2)='KZ' or substring([SEG],1,2)='MA'
            																				  or substring([SEG],1,2)='SM' or substring([SEG],1,2)='TH' or substring([SEG],1,2)='AG' or substring([SEG],1,2)='YZ')
            and ([CSSR_MCI]<95.034 or [OHSR]<80.511 or [CDR(including_CS_IRAT_handovers_3G_to2G)(Nokia_SEG)]>5.263 or [TCH_Assignment_FR]>3.869 or [RxQuality_DL]<84.535 or [RxQuality_UL]<87.287 or [SDCCH_Congestion_Rate]>0.313 or [SDCCH_Access_Success_Rate]<88.808 or [SDCCH_Drop_Rate]>1.351 or [IHSR]<86.047) and Date  ='" + Date_of_WPC + "'";




                //    string Nokia_2G_TBL = @"select[Date], [BSC], [SEG], [TCH_Traffic], [CSSR_MCI], [TCH_Cong_Rate], [SDCCH_Congestion_Rate], [TCH_Assignment_FR], 
                //[SDCCH_Access_Success_Rate], [SDCCH_Drop_Rate], [CDR(including_CS_IRAT_handovers_3G_to2G)(Nokia_SEG)], 
                //[IHSR], [OHSR], [RxQuality_DL], [RxQuality_UL], [TCH_Availability], [AMRHR_USAGE]
                //        from[dbo].[CC2_Nokia_Cell_Daily] where
                //       (substring([SEG],1,2)='TH')
                //and ([CSSR_MCI]<95.034 or [OHSR]<80.511 or [CDR(including_CS_IRAT_handovers_3G_to2G)(Nokia_SEG)]>5.263 or [TCH_Assignment_FR]>3.869 or [RxQuality_DL]<84.535 or [RxQuality_UL]<87.287 or [SDCCH_Congestion_Rate]>0.313 or [SDCCH_Access_Success_Rate]<88.808 or [SDCCH_Drop_Rate]>1.351 or [IHSR]<86.047) and Date  ='" + Date_of_WPC + "'";




                SqlCommand Nokia_2G_TBL_Quary = new SqlCommand(Nokia_2G_TBL, connection);
                Nokia_2G_TBL_Quary.ExecuteNonQuery();
                DataTable Nokia_2G_Table = new DataTable();
                SqlDataAdapter dataAdapter_Nokia_2G = new SqlDataAdapter(Nokia_2G_TBL_Quary);
                dataAdapter_Nokia_2G.Fill(Nokia_2G_Table);





                //Table of Oldest 7 Dayes
                string Ericsson_2G_TBL_7 = @"select [Date], [BSC], [Cell], [TCH_Traffic] as 'TCH_Traffic_Daily (Erlang)', [CSSR_MCI] as'CSSR', [TCH_Congestion] as 'TCH Congestion Rate', [SDCCH_Congestion] as 'SDCCH Congestion Rate', [TCH_Assign_Fail_Rate(NAK)(Eric_CELL)] as 'TCH Assign Fail Rate', 
[SDCCH_Access_Succ_Rate] as 'SDCCH Access Success Rate', [SDCCH_Drop_Rate] as 'SDCCH Drop Rate',  [CDR(not Affected by incoming Handovers from 3G)(Eric_CELL)] as 'Voice Drop Rate', 
[IHSR] as 'IHSR', [OHSR] as 'OHSR', [RxQual_DL] as 'RxQuality_DL', [RxQual_UL] as 'RxQuality_UL', [TCH_Availability] as 'TCH Availability', [AMRHR_USAGE] as 'AMRHR Usage'
from [dbo].[CC2_Ericsson_Cell_Daily] where 
(substring([Cell],1,2)='KJ' or substring([Cell],1,2)='CH' or substring([Cell],1,2)='AS' or substring([Cell],1,2)='GL' or 
                                                                                  substring([Cell],1,2)='GN'  or substring([Cell],1,2)='KM' or substring([Cell],1,2)='KH' or substring([Cell],1,2)='KZ' or substring([Cell],1,2)='MA'
																				  or substring([Cell],1,2)='SM' or substring([Cell],1,2)='TH' or substring([Cell],1,2)='AG' or substring([Cell],1,2)='YZ')
and Date  >='" + Date_of_WPC_7 + "' and Date  <'" + Date_of_WPC + "'";



                //                string Ericsson_2G_TBL_7 = @"select [Date], [BSC], [Cell], [TCH_Traffic] as 'TCH_Traffic_Daily (Erlang)', [CSSR_MCI] as'CSSR', [TCH_Congestion] as 'TCH Congestion Rate', [SDCCH_Congestion] as 'SDCCH Congestion Rate', [TCH_Assign_Fail_Rate(NAK)(Eric_CELL)] as 'TCH Assign Fail Rate', 
                //[SDCCH_Access_Succ_Rate] as 'SDCCH Access Success Rate', [SDCCH_Drop_Rate] as 'SDCCH Drop Rate',  [CDR(not Affected by incoming Handovers from 3G)(Eric_CELL)] as 'Voice Drop Rate', 
                //[IHSR] as 'IHSR', [OHSR] as 'OHSR', [RxQual_DL] as 'RxQuality_DL', [RxQual_UL] as 'RxQuality_UL', [TCH_Availability] as 'TCH Availability', [AMRHR_USAGE] as 'AMRHR Usage'
                //from [dbo].[CC2_Ericsson_Cell_Daily] where 
                //(substring([Cell],1,2)='TH')
                //and Date  >='" + Date_of_WPC_7 + "' and Date  <'" + Date_of_WPC + "'";


                SqlCommand Ericsson_2G_TBL_Quary_7 = new SqlCommand(Ericsson_2G_TBL_7, connection);
                Ericsson_2G_TBL_Quary_7.CommandTimeout = 0;
                Ericsson_2G_TBL_Quary_7.ExecuteNonQuery();
                DataTable Ericsson_2G_Table_7 = new DataTable();
                SqlDataAdapter dataAdapter_Ericsson_2G_7 = new SqlDataAdapter(Ericsson_2G_TBL_Quary_7);
                dataAdapter_Ericsson_2G_7.Fill(Ericsson_2G_Table_7);


                string Huawei_2G_TBL_7 = @"select[Date], [BSC], [Cell], [TCH_Traffic], [CSSR3], [TCH_Cong], [SDCCH_Congestion_Rate], [TCH_Assignment_FR], 
            [SDCCH_Access_Success_Rate2], [SDCCH_Drop_Rate], [CDR3], 
            [IHSR2], [OHSR2], [RX_QUALITTY_DL_NEW], [RX_QUALITTY_UL_NEW], [TCH_Availability], [AMRHR_USAGE]
                    from[dbo].[CC2_Huawei_Cell_Daily] where
                   (substring([Cell],1,2)='KJ' or substring([Cell],1,2)='CH' or substring([Cell],1,2)='AS' or substring([Cell],1,2)='GL' or
                                                                                                     substring([Cell],1,2)='GN'  or substring([Cell],1,2)='KM' or substring([Cell],1,2)='KH' or substring([Cell],1,2)='KZ' or substring([Cell],1,2)='MA'
            																				  or substring([Cell],1,2)='SM' or substring([Cell],1,2)='TH' or substring([Cell],1,2)='AG' or substring([Cell],1,2)='YZ')
            and Date  >='" + Date_of_WPC_7 + "' and Date  <'" + Date_of_WPC + "'";



                //    string Huawei_2G_TBL_7 = @"select[Date], [BSC], [Cell], [TCH_Traffic], [CSSR3], [TCH_Cong], [SDCCH_Congestion_Rate], [TCH_Assignment_FR], 
                //[SDCCH_Access_Success_Rate2], [SDCCH_Drop_Rate], [CDR3], 
                //[IHSR2], [OHSR2], [RX_QUALITTY_DL_NEW], [RX_QUALITTY_UL_NEW], [TCH_Availability], [AMRHR_USAGE]
                //        from[dbo].[CC2_Huawei_Cell_Daily] where
                //       (substring([Cell],1,2)='TH')
                //and Date  >='" + Date_of_WPC_7 + "' and Date  <'" + Date_of_WPC + "'";

                SqlCommand Huawei_2G_TBL_Quary_7 = new SqlCommand(Huawei_2G_TBL_7, connection);
                Huawei_2G_TBL_Quary_7.CommandTimeout = 0;
                Huawei_2G_TBL_Quary_7.ExecuteNonQuery();
                DataTable Huawei_2G_Table_7 = new DataTable();
                SqlDataAdapter dataAdapter_Huawei_2G_7 = new SqlDataAdapter(Huawei_2G_TBL_Quary_7);
                dataAdapter_Huawei_2G_7.Fill(Huawei_2G_Table_7);



                string Nokia_2G_TBL_7 = @"select[Date], [BSC], [SEG], [TCH_Traffic], [CSSR_MCI], [TCH_Cong_Rate], [SDCCH_Congestion_Rate], [TCH_Assignment_FR], 
            [SDCCH_Access_Success_Rate], [SDCCH_Drop_Rate], [CDR(including_CS_IRAT_handovers_3G_to2G)(Nokia_SEG)], 
            [IHSR], [OHSR], [RxQuality_DL], [RxQuality_UL], [TCH_Availability], [AMRHR_USAGE]
                    from[dbo].[CC2_Nokia_Cell_Daily] where
                   (substring([SEG],1,2)='KJ' or substring([SEG],1,2)='CH' or substring([SEG],1,2)='AS' or substring([SEG],1,2)='GL' or
                                                                                                     substring([SEG],1,2)='GN'  or substring([SEG],1,2)='KM' or substring([SEG],1,2)='KH' or substring([SEG],1,2)='KZ' or substring([SEG],1,2)='MA'
            																				  or substring([SEG],1,2)='SM' or substring([SEG],1,2)='TH' or substring([SEG],1,2)='AG' or substring([SEG],1,2)='YZ')
            and Date  >='" + Date_of_WPC_7 + "' and Date  <'" + Date_of_WPC + "'";




                //    string Nokia_2G_TBL_7 = @"select[Date], [BSC], [SEG], [TCH_Traffic], [CSSR_MCI], [TCH_Cong_Rate], [SDCCH_Congestion_Rate], [TCH_Assignment_FR], 
                //[SDCCH_Access_Success_Rate], [SDCCH_Drop_Rate], [CDR(including_CS_IRAT_handovers_3G_to2G)(Nokia_SEG)], 
                //[IHSR], [OHSR], [RxQuality_DL], [RxQuality_UL], [TCH_Availability], [AMRHR_USAGE]
                //        from[dbo].[CC2_Nokia_Cell_Daily] where
                //       (substring([SEG],1,2)='TH')
                //and Date  >='" + Date_of_WPC_7 + "' and Date  <'" + Date_of_WPC + "'";



                SqlCommand Nokia_2G_TBL_Quary_7 = new SqlCommand(Nokia_2G_TBL_7, connection);
                Nokia_2G_TBL_Quary_7.CommandTimeout = 0;
                Nokia_2G_TBL_Quary_7.ExecuteNonQuery();
                DataTable Nokia_2G_Table_7 = new DataTable();
                SqlDataAdapter dataAdapter_Nokia_2G_7 = new SqlDataAdapter(Nokia_2G_TBL_Quary_7);
                dataAdapter_Nokia_2G_7.Fill(Nokia_2G_Table_7);





                // Update Vendor an Site
                Ericsson_2G_Table.Columns.Add("Vendor", typeof(string));
                Huawei_2G_Table.Columns.Add("Vendor", typeof(string));
                Nokia_2G_Table.Columns.Add("Vendor", typeof(string));
                Ericsson_2G_Table.Columns.Add("Contractor", typeof(string));
                Huawei_2G_Table.Columns.Add("Contractor", typeof(string));
                Nokia_2G_Table.Columns.Add("Contractor", typeof(string));
                Ericsson_2G_Table.Columns.Add("Province", typeof(string));
                Huawei_2G_Table.Columns.Add("Province", typeof(string));
                Nokia_2G_Table.Columns.Add("Province", typeof(string));
                Ericsson_2G_Table.Columns.Add("Site", typeof(string));
                Huawei_2G_Table.Columns.Add("Site", typeof(string));
                Nokia_2G_Table.Columns.Add("Site", typeof(string));
                Ericsson_2G_Table.Columns.Add("Coverage Type", typeof(string));
                Huawei_2G_Table.Columns.Add("Coverage Type", typeof(string));
                Nokia_2G_Table.Columns.Add("Coverage Type", typeof(string));


                string province_letter = "";
                string cell = "";
                for (int i = 0; i < Ericsson_2G_Table.Rows.Count; i++)
                {
                    cell = Ericsson_2G_Table.Rows[i][2].ToString();
                    Ericsson_2G_Table.Rows[i][17] = "Ericsson";
                    if (cell != "")
                    {
                        Ericsson_2G_Table.Rows[i][20] = cell.Substring(0, 6);
                    }
                    else
                    {
                        Ericsson_2G_Table.Rows[i][20] = "NA";
                    }
                }
                for (int i = 0; i < Huawei_2G_Table.Rows.Count; i++)
                {
                    cell = Huawei_2G_Table.Rows[i][2].ToString();
                    Huawei_2G_Table.Rows[i][17] = "Huawei";
                    if (cell != "" && cell.Length > 7)
                    {
                        Huawei_2G_Table.Rows[i][20] = cell.Substring(0, 2) + cell.Substring(4, 4);
                    }
                    else if (cell != "" && cell.Length == 7)
                    {
                        Huawei_2G_Table.Rows[i][20] = cell.Substring(0, 6);
                    }
                    else
                    {
                        Huawei_2G_Table.Rows[i][20] = "NA";
                    }
                }
                for (int i = 0; i < Nokia_2G_Table.Rows.Count; i++)
                {
                    cell = Nokia_2G_Table.Rows[i][2].ToString();
                    Nokia_2G_Table.Rows[i][17] = "Nokia";
                    if (cell != "")
                    {
                        Nokia_2G_Table.Rows[i][20] = cell.Substring(0, 2) + cell.Substring(4, 4);
                    }
                    else
                    {
                        Nokia_2G_Table.Rows[i][20] = "NA";
                    }
                }


                //Union Tables in 3 Vendor
                dtUnion_2G = Ericsson_2G_Table.AsEnumerable().Union(Huawei_2G_Table.AsEnumerable()).CopyToDataTable<DataRow>();
                Table_2G = dtUnion_2G.AsEnumerable().Union(Nokia_2G_Table.AsEnumerable()).CopyToDataTable<DataRow>();



                // join with ARAS Table to get Coverge Type
                var JoinResult = (from p in Table_2G.AsEnumerable()
                                  join t in ARAS_Table.AsEnumerable()
                                  on p.Field<string>("Site") equals t.Field<string>("LOCATION")
                                  select new
                                  {
                                      BSC = p.Field<string>("BSC"),
                                      Site = p.Field<string>("Site"),
                                      Coverage = t.Field<string>("COVERAGE_TYPE_OPTIMIZATION")
                                  }).ToList();


                // Update Province and Contractor
                progressBar1.Minimum = 0;
                progressBar1.Maximum = Table_2G.Rows.Count;
                int Row_Count_of_WPC_Table = 0;
                //  for (int i = 0; i < 1000; i++)
                for (int i = 0; i < Table_2G.Rows.Count; i++)
                {
                    string Vendor = Table_2G.Rows[i][17].ToString();
                    string Site = Table_2G.Rows[i][20].ToString();
                    int Found = 0;
                    int ind = 0;
                    string Coverage_type = "";
                    for (int k = 0; k < JoinResult.Count; k++)
                    {
                        if (Site == JoinResult[k].Site.ToString())
                        {
                            Found = 1;
                            ind = k;
                            Coverage_type = JoinResult[ind].Coverage.ToString();
                            break;
                        }
                    }
                    if (Found == 0)
                    {
                        Coverage_type = "City";
                    }


                    cell = Table_2G.Rows[i][2].ToString();
                    if (cell != "")
                    {
                        province_letter = cell.Substring(0, 2);
                        if (province_letter == "TH")
                        {
                            Table_2G.Rows[i][18] = "NAK-Tehran"; Table_2G.Rows[i][19] = "Tehran"; Table_2G.Rows[i][21] = Coverage_type;
                        }
                        if (province_letter == "KJ")
                        {
                            Table_2G.Rows[i][18] = "NAK-Alborz"; Table_2G.Rows[i][19] = "Alborz"; Table_2G.Rows[i][21] = Coverage_type;
                        }
                        if (province_letter == "MA")
                        {
                            Table_2G.Rows[i][18] = "NAK-North"; Table_2G.Rows[i][19] = "Mazandaran"; Table_2G.Rows[i][21] = Coverage_type;
                        }
                        if (province_letter == "GN")
                        {
                            Table_2G.Rows[i][18] = "NAK-North"; Table_2G.Rows[i][19] = "Gilan"; Table_2G.Rows[i][21] = Coverage_type;
                        }
                        if (province_letter == "GL")
                        {
                            Table_2G.Rows[i][18] = "NAK-North"; Table_2G.Rows[i][19] = "Golestan"; Table_2G.Rows[i][21] = Coverage_type;
                        }
                        if (province_letter == "AS")
                        {
                            Table_2G.Rows[i][18] = "NAK-Huawei"; Table_2G.Rows[i][19] = "East Azarbaijan"; Table_2G.Rows[i][21] = Coverage_type;
                        }
                        if (province_letter == "AG")
                        {
                            Table_2G.Rows[i][18] = "NAK-Huawei"; Table_2G.Rows[i][19] = "West Azarbaijan"; Table_2G.Rows[i][21] = Coverage_type;
                        }
                        if (province_letter == "KZ")
                        {
                            Table_2G.Rows[i][18] = "NAK-Huawei"; Table_2G.Rows[i][19] = "Khuzestan"; Table_2G.Rows[i][21] = Coverage_type;
                        }
                        if (province_letter == "KH")
                        {
                            Table_2G.Rows[i][18] = "NAK-Nokia"; Table_2G.Rows[i][19] = "Khorasan Razavi"; Table_2G.Rows[i][21] = Coverage_type;
                        }
                        if (province_letter == "YZ")
                        {
                            Table_2G.Rows[i][18] = "NAK-Nokia"; Table_2G.Rows[i][19] = "Yazd"; Table_2G.Rows[i][21] = Coverage_type;
                        }
                        if (province_letter == "SM")
                        {
                            Table_2G.Rows[i][18] = "NAK-Nokia"; Table_2G.Rows[i][19] = "Semnan"; Table_2G.Rows[i][21] = Coverage_type;
                        }
                        if (province_letter == "CH")
                        {
                            Table_2G.Rows[i][18] = "NAK-Nokia"; Table_2G.Rows[i][19] = "Chahar Mahal Va Bakhtiari"; Table_2G.Rows[i][21] = Coverage_type;
                        }
                        if (province_letter == "KM")
                        {
                            Table_2G.Rows[i][18] = "NAK-Nokia"; Table_2G.Rows[i][19] = "Kerman"; Table_2G.Rows[i][21] = Coverage_type;
                        }
                    }

                    string Contractor = Table_2G.Rows[i][18].ToString();
                    string Province = Table_2G.Rows[i][19].ToString();
                    string BSC = Table_2G.Rows[i][1].ToString();
                    double Traffic = -1000;
                    double CSSR = -1000;
                    double TCH_CONG = -1000;
                    double SDCCH_CONG = -1000;
                    double TCH_ASFR = -1000;
                    double SDCCH_SR = -1000;
                    double SDCCH_DROP = -1000;
                    double CDR = -1000;
                    double IHSR = -1000;
                    double OHSR = -1000;
                    double RXDL = -1000;
                    double RXUL = -1000;
                    double Availability = -1000;
                    double AMRHR = -1000;

                    if (Table_2G.Rows[i][3].ToString() != "")
                    {
                        Traffic = Convert.ToDouble(Table_2G.Rows[i][3]);
                    }
                    if (Table_2G.Rows[i][4].ToString() != "")
                    {
                        CSSR = Convert.ToDouble(Table_2G.Rows[i][4]);
                    }
                    if (Table_2G.Rows[i][5].ToString() != "")
                    {
                        TCH_CONG = Convert.ToDouble(Table_2G.Rows[i][5]);
                    }
                    if (Table_2G.Rows[i][6].ToString() != "")
                    {
                        SDCCH_CONG = Convert.ToDouble(Table_2G.Rows[i][6]);
                    }
                    if (Table_2G.Rows[i][7].ToString() != "")
                    {
                        TCH_ASFR = Convert.ToDouble(Table_2G.Rows[i][7]);
                    }
                    if (Table_2G.Rows[i][8].ToString() != "")
                    {
                        SDCCH_SR = Convert.ToDouble(Table_2G.Rows[i][8]);
                    }
                    if (Table_2G.Rows[i][9].ToString() != "")
                    {
                        SDCCH_DROP = Convert.ToDouble(Table_2G.Rows[i][9]);
                    }
                    if (Table_2G.Rows[i][10].ToString() != "")
                    {
                        CDR = Convert.ToDouble(Table_2G.Rows[i][10]);
                    }
                    if (Table_2G.Rows[i][11].ToString() != "")
                    {
                        IHSR = Convert.ToDouble(Table_2G.Rows[i][11]);
                    }
                    if (Table_2G.Rows[i][12].ToString() != "")
                    {
                        OHSR = Convert.ToDouble(Table_2G.Rows[i][12]);
                    }
                    if (Table_2G.Rows[i][13].ToString() != "")
                    {
                        RXDL = Convert.ToDouble(Table_2G.Rows[i][13]);
                    }
                    if (Table_2G.Rows[i][14].ToString() != "")
                    {
                        RXUL = Convert.ToDouble(Table_2G.Rows[i][14]);
                    }
                    if (Table_2G.Rows[i][15].ToString() != "")
                    {
                        Availability = Convert.ToDouble(Table_2G.Rows[i][15]);
                    }
                    if (Table_2G.Rows[i][16].ToString() != "")
                    {
                        AMRHR = Convert.ToDouble(Table_2G.Rows[i][16]);
                    }

                    // Fill WPc Table
                    int TH_Index = 0;


                    var q1 = (from p in BASE_Table.AsEnumerable()
                              where p.Field<string>("Cell") == cell
                              select p).ToList();
                    var q2 = q1;

                    string CSSR_BL = "";
                    string OHSR_BL = "";
                    string CDR_BL = "";
                    string TCH_ASFR_BL = "";
                    string RXDL_BL = "";
                    string RXUL_BL = "";
                    string SDCCH_CONG_BL = "";
                    string SDCCH_SR_BL = "";
                    string SDCCH_DROP_BL = "";
                    string IHSR_BL = "";
                    if (q1.Count != 0)
                    {

                    }




                    if (Vendor == "Ericsson" && Coverage_type == "City")
                    {
                        q1 = (from p in Ericsson_2G_Table_7.AsEnumerable()
                              where p.Field<string>("Cell") == cell
                              select p).ToList();

                        TH_Index = 0;
                    }
                    if (Vendor == "Ericsson" && Coverage_type == "None_City")
                    {
                        q1 = (from p in Ericsson_2G_Table_7.AsEnumerable()
                              where p.Field<string>("Cell") == cell
                              select p).ToList();

                        TH_Index = 1;
                    }
                    if (Vendor == "Huawei" && Coverage_type == "City")
                    {
                        q1 = (from p in Huawei_2G_Table_7.AsEnumerable()
                              where p.Field<string>("Cell") == cell
                              select p).ToList();

                        TH_Index = 2;
                    }
                    if (Vendor == "Huawei" && Coverage_type == "None_City")
                    {
                        q1 = (from p in Huawei_2G_Table_7.AsEnumerable()
                              where p.Field<string>("Cell") == cell
                              select p).ToList();

                        TH_Index = 3;
                    }
                    if (Vendor == "Nokia" && Coverage_type == "City")
                    {
                        q1 = (from p in Nokia_2G_Table_7.AsEnumerable()
                              where p.Field<string>("SEG") == cell
                              select p).ToList();

                        TH_Index = 4;
                    }
                    if (Vendor == "Nokia" && Coverage_type == "None_City")
                    {
                        q1 = (from p in Nokia_2G_Table_7.AsEnumerable()
                              where p.Field<string>("SEG") == cell
                              select p).ToList();

                        TH_Index = 5;
                    }


                    DateTime[] d1_vec = new DateTime[20];
                    for (int k = 1; k <= q1.Count; k++)
                    {
                        d1_vec[k - 1] = Convert.ToDateTime(q1[q1.Count - k].ItemArray[0]);
                    }

                    int count = 0;
                    int update = 0;

                    if (CSSR > -1000 && CSSR < TH_2G[TH_Index, 0])
                    {
                        count++;
                        int KPI_Index = 4;
                        if (q1.Count == 7)
                        {
                            Table_2G_WPC.Rows.Add(Contractor, Province, Vendor, BSC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, CSSR, "CSSR", "-", q1[6].ItemArray[KPI_Index], q1[5].ItemArray[KPI_Index], q1[4].ItemArray[KPI_Index], q1[3].ItemArray[KPI_Index], q1[2].ItemArray[KPI_Index], q1[1].ItemArray[KPI_Index], q1[0].ItemArray[KPI_Index]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count != 7 && q1.Count != 0)
                        {
                            double[] KPI_of_Days = new double[7];
                            for (int j = 1; j <= 7; j++)
                            {
                                //if (d1_vec.Contains(Date_of_WPC.AddDays(-j)))
                                //{
                                //    //DateTime d1 = Convert.ToDateTime(q1[q1.Count - j].ItemArray[0]);
                                //    //if (d1 == Date_of_WPC.AddDays(-j))
                                //    //{
                                //        if (q1[q1.Count - j].ItemArray[KPI_Index].ToString() == "")
                                //        {
                                //            KPI_of_Days[j - 1] = -20000;
                                //        }
                                //        else
                                //        {
                                //            KPI_of_Days[j - 1] = Convert.ToDouble(q1[q1.Count - j].ItemArray[KPI_Index]);
                                //        }

                                //  //  }
                                //    //else
                                //    //{
                                //    //    KPI_of_Days[j - 1] = -20000;
                                //    //}
                                //}
                                //else
                                //{
                                //    KPI_of_Days[j - 1] = -20000;
                                //}
                                ////if (Convert.ToDateTime(q1[q1.Count - j].ItemArray[0]) == Date_of_WPC.AddDays(-j))
                                ////{
                                ////    KPI_of_Days[j - 1] = -20000;
                                ////}

                                if (q1.Count - j < 0)
                                {
                                    KPI_of_Days[j - 1] = -20000;
                                }
                                else
                                {
                                    DateTime d1 = Convert.ToDateTime(q1[q1.Count - j].ItemArray[0]);
                                    if (d1 == Date_of_WPC.AddDays(-j))
                                    {
                                        if (q1[q1.Count - j].ItemArray[KPI_Index].ToString() == "")
                                        {
                                            KPI_of_Days[j - 1] = -20000;
                                        }
                                        else
                                        {
                                            KPI_of_Days[j - 1] = Convert.ToDouble(q1[q1.Count - j].ItemArray[KPI_Index]);
                                        }

                                    }
                                    else
                                    {
                                        KPI_of_Days[j - 1] = -20000;
                                    }
                                }
                            }
                            Table_2G_WPC.Rows.Add(Contractor, Province, Vendor, BSC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, CSSR, "CSSR", "-", KPI_of_Days[0], KPI_of_Days[1], KPI_of_Days[2], KPI_of_Days[3], KPI_of_Days[4], KPI_of_Days[5], KPI_of_Days[6]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count == 0)
                        {
                            Table_2G_WPC.Rows.Add(Contractor, Province, Vendor, BSC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, CSSR, "CSSR", "-");
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                    }
                    if (OHSR > -1000 && OHSR < TH_2G[TH_Index, 1])
                    {
                        count++;
                        int KPI_Index = 12;
                        if (q1.Count == 7)
                        {
                            Table_2G_WPC.Rows.Add(Contractor, Province, Vendor, BSC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, OHSR, "OHSR", "-", q1[6].ItemArray[KPI_Index], q1[5].ItemArray[KPI_Index], q1[4].ItemArray[KPI_Index], q1[3].ItemArray[KPI_Index], q1[2].ItemArray[KPI_Index], q1[1].ItemArray[KPI_Index], q1[0].ItemArray[KPI_Index]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count != 7 && q1.Count != 0)
                        {
                            double[] KPI_of_Days = new double[7];
                            for (int j = 1; j <= 7; j++)
                            {
                                if (q1.Count - j < 0)
                                {
                                    KPI_of_Days[j - 1] = -20000;
                                }
                                else
                                {
                                    DateTime d1 = Convert.ToDateTime(q1[q1.Count - j].ItemArray[0]);
                                    if (d1 == Date_of_WPC.AddDays(-j))
                                    {
                                        if (q1[q1.Count - j].ItemArray[KPI_Index].ToString() == "")
                                        {
                                            KPI_of_Days[j - 1] = -20000;
                                        }
                                        else
                                        {
                                            KPI_of_Days[j - 1] = Convert.ToDouble(q1[q1.Count - j].ItemArray[KPI_Index]);
                                        }

                                    }
                                    else
                                    {
                                        KPI_of_Days[j - 1] = -20000;
                                    }
                                }
                            }
                            Table_2G_WPC.Rows.Add(Contractor, Province, Vendor, BSC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, OHSR, "OHSR", "-", KPI_of_Days[0], KPI_of_Days[1], KPI_of_Days[2], KPI_of_Days[3], KPI_of_Days[4], KPI_of_Days[5], KPI_of_Days[6]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count == 0)
                        {
                            Table_2G_WPC.Rows.Add(Contractor, Province, Vendor, BSC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, OHSR, "OHSR", "-");
                            Row_Count_of_WPC_Table++; update = 1;
                        }

                    }
                    if (CDR > -1000 && CDR > TH_2G[TH_Index, 2])
                    {
                        count++;
                        int KPI_Index = 10;
                        if (q1.Count == 7)
                        {
                            Table_2G_WPC.Rows.Add(Contractor, Province, Vendor, BSC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, CDR, "CDR", "-", q1[6].ItemArray[KPI_Index], q1[5].ItemArray[KPI_Index], q1[4].ItemArray[KPI_Index], q1[3].ItemArray[KPI_Index], q1[2].ItemArray[KPI_Index], q1[1].ItemArray[KPI_Index], q1[0].ItemArray[KPI_Index]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count != 7 && q1.Count != 0)
                        {
                            double[] KPI_of_Days = new double[7];
                            for (int j = 1; j <= 7; j++)
                            {
                                if (q1.Count - j < 0)
                                {
                                    KPI_of_Days[j - 1] = -20000;
                                }
                                else
                                {
                                    DateTime d1 = Convert.ToDateTime(q1[q1.Count - j].ItemArray[0]);
                                    if (d1 == Date_of_WPC.AddDays(-j))
                                    {
                                        if (q1[q1.Count - j].ItemArray[KPI_Index].ToString() == "")
                                        {
                                            KPI_of_Days[j - 1] = -20000;
                                        }
                                        else
                                        {
                                            KPI_of_Days[j - 1] = Convert.ToDouble(q1[q1.Count - j].ItemArray[KPI_Index]);
                                        }

                                    }
                                    else
                                    {
                                        KPI_of_Days[j - 1] = -20000;
                                    }
                                }
                            }
                            Table_2G_WPC.Rows.Add(Contractor, Province, Vendor, BSC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, CDR, "CDR", "-", KPI_of_Days[0], KPI_of_Days[1], KPI_of_Days[2], KPI_of_Days[3], KPI_of_Days[4], KPI_of_Days[5], KPI_of_Days[6]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count == 0)
                        {
                            Table_2G_WPC.Rows.Add(Contractor, Province, Vendor, BSC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, CDR, "CDR", "-");
                            Row_Count_of_WPC_Table++; update = 1;
                        }

                    }
                    if (TCH_ASFR > -1000 && TCH_ASFR > TH_2G[TH_Index, 3])
                    {
                        count++;
                        int KPI_Index = 7;
                        if (q1.Count == 7)
                        {
                            Table_2G_WPC.Rows.Add(Contractor, Province, Vendor, BSC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, TCH_ASFR, "TCH_ASFR", "-", q1[6].ItemArray[KPI_Index], q1[5].ItemArray[KPI_Index], q1[4].ItemArray[KPI_Index], q1[3].ItemArray[KPI_Index], q1[2].ItemArray[KPI_Index], q1[1].ItemArray[KPI_Index], q1[0].ItemArray[KPI_Index]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count != 7 && q1.Count != 0)
                        {
                            double[] KPI_of_Days = new double[7];
                            for (int j = 1; j <= 7; j++)
                            {
                                if (q1.Count - j < 0)
                                {
                                    KPI_of_Days[j - 1] = -20000;
                                }
                                else
                                {
                                    DateTime d1 = Convert.ToDateTime(q1[q1.Count - j].ItemArray[0]);
                                    if (d1 == Date_of_WPC.AddDays(-j))
                                    {
                                        if (q1[q1.Count - j].ItemArray[KPI_Index].ToString() == "")
                                        {
                                            KPI_of_Days[j - 1] = -20000;
                                        }
                                        else
                                        {
                                            KPI_of_Days[j - 1] = Convert.ToDouble(q1[q1.Count - j].ItemArray[KPI_Index]);
                                        }

                                    }
                                    else
                                    {
                                        KPI_of_Days[j - 1] = -20000;
                                    }
                                }
                            }
                            Table_2G_WPC.Rows.Add(Contractor, Province, Vendor, BSC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, TCH_ASFR, "TCH_ASFR", "-", KPI_of_Days[0], KPI_of_Days[1], KPI_of_Days[2], KPI_of_Days[3], KPI_of_Days[4], KPI_of_Days[5], KPI_of_Days[6]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count == 0)
                        {
                            Table_2G_WPC.Rows.Add(Contractor, Province, Vendor, BSC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, TCH_ASFR, "TCH_ASFR", "-");
                            Row_Count_of_WPC_Table++; update = 1;
                        }

                    }
                    if (RXDL > -1000 && RXDL < TH_2G[TH_Index, 4])
                    {
                        count++;
                        int KPI_Index = 13;
                        if (q1.Count == 7)
                        {
                            Table_2G_WPC.Rows.Add(Contractor, Province, Vendor, BSC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, RXDL, "RXDL", "-", q1[6].ItemArray[KPI_Index], q1[5].ItemArray[KPI_Index], q1[4].ItemArray[KPI_Index], q1[3].ItemArray[KPI_Index], q1[2].ItemArray[KPI_Index], q1[1].ItemArray[KPI_Index], q1[0].ItemArray[KPI_Index]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count != 7 && q1.Count != 0)
                        {
                            double[] KPI_of_Days = new double[7];
                            for (int j = 1; j <= 7; j++)
                            {
                                if (q1.Count - j < 0)
                                {
                                    KPI_of_Days[j - 1] = -20000;
                                }
                                else
                                {
                                    DateTime d1 = Convert.ToDateTime(q1[q1.Count - j].ItemArray[0]);
                                    if (d1 == Date_of_WPC.AddDays(-j))
                                    {
                                        if (q1[q1.Count - j].ItemArray[KPI_Index].ToString() == "")
                                        {
                                            KPI_of_Days[j - 1] = -20000;
                                        }
                                        else
                                        {
                                            KPI_of_Days[j - 1] = Convert.ToDouble(q1[q1.Count - j].ItemArray[KPI_Index]);
                                        }

                                    }
                                    else
                                    {
                                        KPI_of_Days[j - 1] = -20000;
                                    }
                                }
                            }
                            Table_2G_WPC.Rows.Add(Contractor, Province, Vendor, BSC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, RXDL, "RXDL", "-", KPI_of_Days[0], KPI_of_Days[1], KPI_of_Days[2], KPI_of_Days[3], KPI_of_Days[4], KPI_of_Days[5], KPI_of_Days[6]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count == 0)
                        {
                            Table_2G_WPC.Rows.Add(Contractor, Province, Vendor, BSC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, RXDL, "RXDL", "-");
                            Row_Count_of_WPC_Table++; update = 1;
                        }

                    }
                    if (RXUL > -1000 && RXUL < TH_2G[TH_Index, 5])
                    {
                        count++;
                        int KPI_Index = 14;
                        if (q1.Count == 7)
                        {
                            Table_2G_WPC.Rows.Add(Contractor, Province, Vendor, BSC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, RXUL, "RXUL", "-", q1[6].ItemArray[KPI_Index], q1[5].ItemArray[KPI_Index], q1[4].ItemArray[KPI_Index], q1[3].ItemArray[KPI_Index], q1[2].ItemArray[KPI_Index], q1[1].ItemArray[KPI_Index], q1[0].ItemArray[KPI_Index]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count != 7 && q1.Count != 0)
                        {
                            double[] KPI_of_Days = new double[7];
                            for (int j = 1; j <= 7; j++)
                            {
                                if (q1.Count - j < 0)
                                {
                                    KPI_of_Days[j - 1] = -20000;
                                }
                                else
                                {
                                    DateTime d1 = Convert.ToDateTime(q1[q1.Count - j].ItemArray[0]);
                                    if (d1 == Date_of_WPC.AddDays(-j))
                                    {
                                        if (q1[q1.Count - j].ItemArray[KPI_Index].ToString() == "")
                                        {
                                            KPI_of_Days[j - 1] = -20000;
                                        }
                                        else
                                        {
                                            KPI_of_Days[j - 1] = Convert.ToDouble(q1[q1.Count - j].ItemArray[KPI_Index]);
                                        }

                                    }
                                    else
                                    {
                                        KPI_of_Days[j - 1] = -20000;
                                    }
                                }
                            }
                            Table_2G_WPC.Rows.Add(Contractor, Province, Vendor, BSC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, RXUL, "RXUL", "-", KPI_of_Days[0], KPI_of_Days[1], KPI_of_Days[2], KPI_of_Days[3], KPI_of_Days[4], KPI_of_Days[5], KPI_of_Days[6]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count == 0)
                        {
                            Table_2G_WPC.Rows.Add(Contractor, Province, Vendor, BSC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, RXUL, "RXUL", "-");
                            Row_Count_of_WPC_Table++; update = 1;
                        }

                    }
                    if (SDCCH_CONG > -1000 && SDCCH_CONG > TH_2G[TH_Index, 6])
                    {
                        count++;
                        int KPI_Index = 6;
                        if (q1.Count == 7)
                        {
                            Table_2G_WPC.Rows.Add(Contractor, Province, Vendor, BSC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, SDCCH_CONG, "SDCCH_CONG", "-", q1[6].ItemArray[KPI_Index], q1[5].ItemArray[KPI_Index], q1[4].ItemArray[KPI_Index], q1[3].ItemArray[KPI_Index], q1[2].ItemArray[KPI_Index], q1[1].ItemArray[KPI_Index], q1[0].ItemArray[KPI_Index]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count != 7 && q1.Count != 0)
                        {
                            double[] KPI_of_Days = new double[7];
                            for (int j = 1; j <= 7; j++)
                            {
                                if (q1.Count - j < 0)
                                {
                                    KPI_of_Days[j - 1] = -20000;
                                }
                                else
                                {
                                    DateTime d1 = Convert.ToDateTime(q1[q1.Count - j].ItemArray[0]);
                                    if (d1 == Date_of_WPC.AddDays(-j))
                                    {
                                        if (q1[q1.Count - j].ItemArray[KPI_Index].ToString() == "")
                                        {
                                            KPI_of_Days[j - 1] = -20000;
                                        }
                                        else
                                        {
                                            KPI_of_Days[j - 1] = Convert.ToDouble(q1[q1.Count - j].ItemArray[KPI_Index]);
                                        }

                                    }
                                    else
                                    {
                                        KPI_of_Days[j - 1] = -20000;
                                    }
                                }
                            }
                            Table_2G_WPC.Rows.Add(Contractor, Province, Vendor, BSC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, SDCCH_CONG, "SDCCH_CONG", "-", KPI_of_Days[0], KPI_of_Days[1], KPI_of_Days[2], KPI_of_Days[3], KPI_of_Days[4], KPI_of_Days[5], KPI_of_Days[6]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count == 0)
                        {
                            Table_2G_WPC.Rows.Add(Contractor, Province, Vendor, BSC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, SDCCH_CONG, "SDCCH_CONG", "-");
                            Row_Count_of_WPC_Table++; update = 1;
                        }

                    }
                    if (SDCCH_SR > -1000 && SDCCH_SR < TH_2G[TH_Index, 7])
                    {
                        count++;
                        int KPI_Index = 8;
                        if (q1.Count == 7)
                        {
                            Table_2G_WPC.Rows.Add(Contractor, Province, Vendor, BSC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, SDCCH_SR, "SDCCH_SR", "-", q1[6].ItemArray[KPI_Index], q1[5].ItemArray[KPI_Index], q1[4].ItemArray[KPI_Index], q1[3].ItemArray[KPI_Index], q1[2].ItemArray[KPI_Index], q1[1].ItemArray[KPI_Index], q1[0].ItemArray[KPI_Index]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count != 7 && q1.Count != 0)
                        {
                            double[] KPI_of_Days = new double[7];
                            for (int j = 1; j <= 7; j++)
                            {
                                if (q1.Count - j < 0)
                                {
                                    KPI_of_Days[j - 1] = -20000;
                                }
                                else
                                {
                                    DateTime d1 = Convert.ToDateTime(q1[q1.Count - j].ItemArray[0]);
                                    if (d1 == Date_of_WPC.AddDays(-j))
                                    {
                                        if (q1[q1.Count - j].ItemArray[KPI_Index].ToString() == "")
                                        {
                                            KPI_of_Days[j - 1] = -20000;
                                        }
                                        else
                                        {
                                            KPI_of_Days[j - 1] = Convert.ToDouble(q1[q1.Count - j].ItemArray[KPI_Index]);
                                        }

                                    }
                                    else
                                    {
                                        KPI_of_Days[j - 1] = -20000;
                                    }
                                }
                            }
                            Table_2G_WPC.Rows.Add(Contractor, Province, Vendor, BSC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, SDCCH_SR, "SDCCH_SR", "-", KPI_of_Days[0], KPI_of_Days[1], KPI_of_Days[2], KPI_of_Days[3], KPI_of_Days[4], KPI_of_Days[5], KPI_of_Days[6]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count == 0)
                        {
                            Table_2G_WPC.Rows.Add(Contractor, Province, Vendor, BSC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, SDCCH_SR, "SDCCH_SR", "-");
                            Row_Count_of_WPC_Table++; update = 1;
                        }

                    }
                    if (SDCCH_DROP > -1000 && SDCCH_DROP > TH_2G[TH_Index, 8])
                    {
                        count++;
                        int KPI_Index = 9;

                        if (q1.Count == 7)
                        {
                            Table_2G_WPC.Rows.Add(Contractor, Province, Vendor, BSC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, SDCCH_DROP, "SDCCH_DROP", "-", q1[6].ItemArray[KPI_Index], q1[5].ItemArray[KPI_Index], q1[4].ItemArray[KPI_Index], q1[3].ItemArray[KPI_Index], q1[2].ItemArray[KPI_Index], q1[1].ItemArray[KPI_Index], q1[0].ItemArray[KPI_Index]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count != 7 && q1.Count != 0)
                        {
                            double[] KPI_of_Days = new double[7];
                            for (int j = 1; j <= 7; j++)
                            {
                                if (q1.Count - j < 0)
                                {
                                    KPI_of_Days[j - 1] = -20000;
                                }
                                else
                                {
                                    DateTime d1 = Convert.ToDateTime(q1[q1.Count - j].ItemArray[0]);
                                    if (d1 == Date_of_WPC.AddDays(-j))
                                    {
                                        if (q1[q1.Count - j].ItemArray[KPI_Index].ToString() == "")
                                        {
                                            KPI_of_Days[j - 1] = -20000;
                                        }
                                        else
                                        {
                                            KPI_of_Days[j - 1] = Convert.ToDouble(q1[q1.Count - j].ItemArray[KPI_Index]);
                                        }

                                    }
                                    else
                                    {
                                        KPI_of_Days[j - 1] = -20000;
                                    }
                                }
                            }
                            Table_2G_WPC.Rows.Add(Contractor, Province, Vendor, BSC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, SDCCH_DROP, "SDCCH_DROP", "-", KPI_of_Days[0], KPI_of_Days[1], KPI_of_Days[2], KPI_of_Days[3], KPI_of_Days[4], KPI_of_Days[5], KPI_of_Days[6]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count == 0)
                        {
                            Table_2G_WPC.Rows.Add(Contractor, Province, Vendor, BSC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, SDCCH_DROP, "SDCCH_DROP", "-");
                            Row_Count_of_WPC_Table++; update = 1;
                        }

                    }
                    if (IHSR > -1000 && IHSR < TH_2G[TH_Index, 9])
                    {
                        count++;
                        int KPI_Index = 11;
                        if (q1.Count == 7)
                        {
                            Table_2G_WPC.Rows.Add(Contractor, Province, Vendor, BSC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, IHSR, "IHSR", "-", q1[6].ItemArray[KPI_Index], q1[5].ItemArray[KPI_Index], q1[4].ItemArray[KPI_Index], q1[3].ItemArray[KPI_Index], q1[2].ItemArray[KPI_Index], q1[1].ItemArray[KPI_Index], q1[0].ItemArray[KPI_Index]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count != 7 && q1.Count != 0)
                        {
                            double[] KPI_of_Days = new double[7];
                            for (int j = 1; j <= 7; j++)
                            {
                                if (q1.Count - j < 0)
                                {
                                    KPI_of_Days[j - 1] = -20000;
                                }
                                else
                                {
                                    DateTime d1 = Convert.ToDateTime(q1[q1.Count - j].ItemArray[0]);
                                    if (d1 == Date_of_WPC.AddDays(-j))
                                    {
                                        if (q1[q1.Count - j].ItemArray[KPI_Index].ToString() == "")
                                        {
                                            KPI_of_Days[j - 1] = -20000;
                                        }
                                        else
                                        {
                                            KPI_of_Days[j - 1] = Convert.ToDouble(q1[q1.Count - j].ItemArray[KPI_Index]);
                                        }

                                    }
                                    else
                                    {
                                        KPI_of_Days[j - 1] = -20000;
                                    }
                                }
                            }
                            Table_2G_WPC.Rows.Add(Contractor, Province, Vendor, BSC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, IHSR, "IHSR", "-", KPI_of_Days[0], KPI_of_Days[1], KPI_of_Days[2], KPI_of_Days[3], KPI_of_Days[4], KPI_of_Days[5], KPI_of_Days[6]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count == 0)
                        {
                            Table_2G_WPC.Rows.Add(Contractor, Province, Vendor, BSC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, IHSR, "IHSR", "-");
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                    }

                    if (update == 1)
                    {
                        Table_2G_WPC.Rows[Row_Count_of_WPC_Table - 1][12] = Convert.ToString(count);
                    }

                    progressBar1.Value = i;

                }




                XLWorkbook wb = new XLWorkbook();
                wb.Worksheets.Add(Table_2G_WPC, "WPC_2G");
                var saveFileDialog = new SaveFileDialog
                {
                    FileName = "WPC_2G",
                    Filter = "Excel files|*.xlsx",
                    Title = "Save an Excel File"
                };

                saveFileDialog.ShowDialog();

                if (!String.IsNullOrWhiteSpace(saveFileDialog.FileName))
                    wb.SaveAs(saveFileDialog.FileName);


                label4.Text = "Finished";
                label4.BackColor = Color.Yellow;
                MessageBox.Show("FINISHED");

            }



            // Quary of Select Data in 3G_CS
            if (checkBox2.Checked == true)
            {

                string Ericsson_3G_CS_TBL = @"select [Date], [ElementID] as 'RNC', [ElementID1], [CS_Traffic] as 'CS_Traffic_Daily (Erlang)', [Cell_Availability_Rate_Exclude_Blocking(UCELL_Eric)] as 'Availability', [Cs_RAB_Establish_Success_Rate] as'CS_RAB_Establish', [IRAT_HO_Voice_Suc_Rate] as 'CS_IRAT_HO_SR', [CS_Drop_Call_Rate] as 'CS_Drop_Rate', [Soft_Handover_Succ_Rate] as 'Soft_HO_SR', 
[CS_RRC_Setup_Success_Rate] as 'CS_RRC_SR'from [dbo].[CC3_Ericsson_Cell_Daily] where 
(substring([ElementID1],1,2)='KJ' or substring([ElementID1],1,2)='CH' or substring([ElementID1],1,2)='AS' or substring([ElementID1],1,2)='GL' or 
                                                                                  substring([ElementID1],1,2)='GN'  or substring([ElementID1],1,2)='KM' or substring([ElementID1],1,2)='KH' or substring([ElementID1],1,2)='KZ' or substring([ElementID1],1,2)='MA'
																				  or substring([ElementID1],1,2)='SM' or substring([ElementID1],1,2)='TH' or substring([ElementID1],1,2)='AG' or substring([ElementID1],1,2)='YZ')
and ([Cs_RAB_Establish_Success_Rate]<99.725 or [IRAT_HO_Voice_Suc_Rate]<60 or [CS_Drop_Call_Rate]>0.687 or [Soft_Handover_Succ_Rate]<99.752 or [CS_RRC_Setup_Success_Rate]<99.57) and Date  ='" + Date_of_WPC + "'";


                SqlCommand Ericsson_3G_CS_TBL_Quary = new SqlCommand(Ericsson_3G_CS_TBL, connection);
                Ericsson_3G_CS_TBL_Quary.ExecuteNonQuery();
                DataTable Ericsson_3G_CS_Table = new DataTable();
                SqlDataAdapter dataAdapter_Ericsson_3G_CS = new SqlDataAdapter(Ericsson_3G_CS_TBL_Quary);
                dataAdapter_Ericsson_3G_CS.Fill(Ericsson_3G_CS_Table);




                string Huawei_3G_CS_TBL = @"select [Date], [ElementID] as 'RNC', [ElementID1], [CS_Erlang] as 'CS_Traffic_Daily (Erlang)', [Radio_Network_Availability_Ratio(Hu_Cell)] as 'Availability', [CS_RAB_Setup_Success_Ratio] as'CS_RAB_Establish', [CS_IRAT_HO_SR] as 'CS_IRAT_HO_SR', [AMR_Call_Drop_Ratio_New(Hu_CELL)] as 'CS_Drop_Rate', [Soft_Handover_Succ_Rate] as 'Soft_HO_SR', 
[CS_RRC_Connection_Establishment_SR] as 'CS_RRC_SR'from [dbo].[CC3_Huawei_Cell_Daily] where 
(substring([ElementID1],1,2)='KJ' or substring([ElementID1],1,2)='CH' or substring([ElementID1],1,2)='AS' or substring([ElementID1],1,2)='GL' or 
                                                                                  substring([ElementID1],1,2)='GN'  or substring([ElementID1],1,2)='KM' or substring([ElementID1],1,2)='KH' or substring([ElementID1],1,2)='KZ' or substring([ElementID1],1,2)='MA'
																				  or substring([ElementID1],1,2)='SM' or substring([ElementID1],1,2)='TH' or substring([ElementID1],1,2)='AG' or substring([ElementID1],1,2)='YZ')
and ([CS_RAB_Setup_Success_Ratio]<99.827 or [CS_IRAT_HO_SR]<89.759 or [AMR_Call_Drop_Ratio_New(Hu_CELL)]>0.623 or [Soft_Handover_Succ_Rate]<99.777 or [CS_RRC_Connection_Establishment_SR]<99.521) and Date  ='" + Date_of_WPC + "'";


                SqlCommand Huawei_3G_CS_TBL_Quary = new SqlCommand(Huawei_3G_CS_TBL, connection);
                Huawei_3G_CS_TBL_Quary.ExecuteNonQuery();
                DataTable Huawei_3G_CS_Table = new DataTable();
                SqlDataAdapter dataAdapter_Huawei_3G_CS = new SqlDataAdapter(Huawei_3G_CS_TBL_Quary);
                dataAdapter_Huawei_3G_CS.Fill(Huawei_3G_CS_Table);




                string Nokia_3G_CS_TBL = @"select [Date], [ElementID] as 'RNC', [ElementID1], [CS_Traffic] as 'CS_Traffic_Daily (Erlang)', [Cell_Availability_excluding_blocked_by_user_state] as 'Availability', [CS_RAB_Establish_Success_Rate] as'CS_RAB_Establish', [Inter_sys_RT_Hard_HO_SR_3Gto2G(CELL_nokia)] as 'CS_IRAT_HO_SR', [CS_Drop_Call_Rate] as 'CS_Drop_Rate', [Soft_Handover_Succ_Rate] as 'Soft_HO_SR', 
[CS_RRC_SETUP_SR_WITHOUT_REPEAT(CELL_NOKIA)] as 'CS_RRC_SR'from [dbo].[CC3_Nokia_Cell_Daily] where 
(substring([ElementID1],1,2)='KJ' or substring([ElementID1],1,2)='CH' or substring([ElementID1],1,2)='AS' or substring([ElementID1],1,2)='GL' or 
                                                                                  substring([ElementID1],1,2)='GN'  or substring([ElementID1],1,2)='KM' or substring([ElementID1],1,2)='KH' or substring([ElementID1],1,2)='KZ' or substring([ElementID1],1,2)='MA'
																				  or substring([ElementID1],1,2)='SM' or substring([ElementID1],1,2)='TH' or substring([ElementID1],1,2)='AG' or substring([ElementID1],1,2)='YZ')
and ([CS_RAB_Establish_Success_Rate]<99.407 or [Inter_sys_RT_Hard_HO_SR_3Gto2G(CELL_nokia)]<88.889 or [CS_Drop_Call_Rate]>0.84 or [Soft_Handover_Succ_Rate]<99.594 or [CS_RRC_SETUP_SR_WITHOUT_REPEAT(CELL_NOKIA)]<99.355) and Date  ='" + Date_of_WPC + "'";


                SqlCommand Nokia_3G_CS_TBL_Quary = new SqlCommand(Nokia_3G_CS_TBL, connection);
                Nokia_3G_CS_TBL_Quary.ExecuteNonQuery();
                DataTable Nokia_3G_CS_Table = new DataTable();
                SqlDataAdapter dataAdapter_Nokia_3G_CS = new SqlDataAdapter(Nokia_3G_CS_TBL_Quary);
                dataAdapter_Nokia_3G_CS.Fill(Nokia_3G_CS_Table);




                // Table of Oldest 7 Dayes

                string Ericsson_3G_CS_TBL_7 = @"select [Date], [ElementID] as 'RNC', [ElementID1], [CS_Traffic] as 'CS_Traffic_Daily (Erlang)', [Cell_Availability_Rate_Exclude_Blocking(UCELL_Eric)] as 'Availability', [Cs_RAB_Establish_Success_Rate] as'CS_RAB_Establish', [IRAT_HO_Voice_Suc_Rate] as 'CS_IRAT_HO_SR', [CS_Drop_Call_Rate] as 'CS_Drop_Rate', [Soft_Handover_Succ_Rate] as 'Soft_HO_SR', 
[CS_RRC_Setup_Success_Rate] as 'CS_RRC_SR'from [dbo].[CC3_Ericsson_Cell_Daily] where 
(substring([ElementID1],1,2)='KJ' or substring([ElementID1],1,2)='CH' or substring([ElementID1],1,2)='AS' or substring([ElementID1],1,2)='GL' or 
                                                                                  substring([ElementID1],1,2)='GN'  or substring([ElementID1],1,2)='KM' or substring([ElementID1],1,2)='KH' or substring([ElementID1],1,2)='KZ' or substring([ElementID1],1,2)='MA'
																				  or substring([ElementID1],1,2)='SM' or substring([ElementID1],1,2)='TH' or substring([ElementID1],1,2)='AG' or substring([ElementID1],1,2)='YZ')
and Date  >='" + Date_of_WPC_7 + "' and Date  <'" + Date_of_WPC + "'";


                SqlCommand Ericsson_3G_CS_TBL_Quary_7 = new SqlCommand(Ericsson_3G_CS_TBL_7, connection);
                Ericsson_3G_CS_TBL_Quary_7.CommandTimeout = 0;
                Ericsson_3G_CS_TBL_Quary_7.ExecuteNonQuery();
                DataTable Ericsson_3G_CS_Table_7 = new DataTable();
                SqlDataAdapter dataAdapter_Ericsson_3G_CS_7 = new SqlDataAdapter(Ericsson_3G_CS_TBL_Quary_7);
                dataAdapter_Ericsson_3G_CS_7.Fill(Ericsson_3G_CS_Table_7);




                string Huawei_3G_CS_TBL_7 = @"select [Date], [ElementID] as 'RNC', [ElementID1], [CS_Erlang] as 'CS_Traffic_Daily (Erlang)', [Radio_Network_Availability_Ratio(Hu_Cell)] as 'Availability', [CS_RAB_Setup_Success_Ratio] as'CS_RAB_Establish', [CS_IRAT_HO_SR] as 'CS_IRAT_HO_SR', [AMR_Call_Drop_Ratio_New(Hu_CELL)] as 'CS_Drop_Rate', [Soft_Handover_Succ_Rate] as 'Soft_HO_SR', 
[CS_RRC_Connection_Establishment_SR] as 'CS_RRC_SR'from [dbo].[CC3_Huawei_Cell_Daily] where 
(substring([ElementID1],1,2)='KJ' or substring([ElementID1],1,2)='CH' or substring([ElementID1],1,2)='AS' or substring([ElementID1],1,2)='GL' or 
                                                                                  substring([ElementID1],1,2)='GN'  or substring([ElementID1],1,2)='KM' or substring([ElementID1],1,2)='KH' or substring([ElementID1],1,2)='KZ' or substring([ElementID1],1,2)='MA'
																				  or substring([ElementID1],1,2)='SM' or substring([ElementID1],1,2)='TH' or substring([ElementID1],1,2)='AG' or substring([ElementID1],1,2)='YZ')
and Date  >='" + Date_of_WPC_7 + "' and Date  <'" + Date_of_WPC + "'";


                SqlCommand Huawei_3G_CS_TBL_Quary_7 = new SqlCommand(Huawei_3G_CS_TBL_7, connection);
                Huawei_3G_CS_TBL_Quary_7.CommandTimeout = 0;
                Huawei_3G_CS_TBL_Quary_7.ExecuteNonQuery();
                DataTable Huawei_3G_CS_Table_7 = new DataTable();
                SqlDataAdapter dataAdapter_Huawei_3G_CS_7 = new SqlDataAdapter(Huawei_3G_CS_TBL_Quary_7);
                dataAdapter_Huawei_3G_CS_7.Fill(Huawei_3G_CS_Table_7);




                string Nokia_3G_CS_TBL_7 = @"select [Date], [ElementID] as 'RNC', [ElementID1], [CS_Traffic] as 'CS_Traffic_Daily (Erlang)', [Cell_Availability_excluding_blocked_by_user_state] as 'Availability', [CS_RAB_Establish_Success_Rate] as'CS_RAB_Establish', [Inter_sys_RT_Hard_HO_SR_3Gto2G(CELL_nokia)] as 'CS_IRAT_HO_SR', [CS_Drop_Call_Rate] as 'CS_Drop_Rate', [Soft_Handover_Succ_Rate] as 'Soft_HO_SR', 
[CS_RRC_SETUP_SR_WITHOUT_REPEAT(CELL_NOKIA)] as 'CS_RRC_SR'from [dbo].[CC3_Nokia_Cell_Daily] where 
(substring([ElementID1],1,2)='KJ' or substring([ElementID1],1,2)='CH' or substring([ElementID1],1,2)='AS' or substring([ElementID1],1,2)='GL' or 
                                                                                  substring([ElementID1],1,2)='GN'  or substring([ElementID1],1,2)='KM' or substring([ElementID1],1,2)='KH' or substring([ElementID1],1,2)='KZ' or substring([ElementID1],1,2)='MA'
																				  or substring([ElementID1],1,2)='SM' or substring([ElementID1],1,2)='TH' or substring([ElementID1],1,2)='AG' or substring([ElementID1],1,2)='YZ')
and Date  >='" + Date_of_WPC_7 + "' and Date  <'" + Date_of_WPC + "'";


                SqlCommand Nokia_3G_CS_TBL_Quary_7 = new SqlCommand(Nokia_3G_CS_TBL_7, connection);
                Nokia_3G_CS_TBL_Quary_7.CommandTimeout = 0;
                Nokia_3G_CS_TBL_Quary_7.ExecuteNonQuery();
                DataTable Nokia_3G_CS_Table_7 = new DataTable();
                SqlDataAdapter dataAdapter_Nokia_3G_CS_7 = new SqlDataAdapter(Nokia_3G_CS_TBL_Quary_7);
                dataAdapter_Nokia_3G_CS_7.Fill(Nokia_3G_CS_Table_7);



                // Update Vendor an Site
                Ericsson_3G_CS_Table.Columns.Add("Vendor", typeof(string));
                Huawei_3G_CS_Table.Columns.Add("Vendor", typeof(string));
                Nokia_3G_CS_Table.Columns.Add("Vendor", typeof(string));
                Ericsson_3G_CS_Table.Columns.Add("Contractor", typeof(string));
                Huawei_3G_CS_Table.Columns.Add("Contractor", typeof(string));
                Nokia_3G_CS_Table.Columns.Add("Contractor", typeof(string));
                Ericsson_3G_CS_Table.Columns.Add("Province", typeof(string));
                Huawei_3G_CS_Table.Columns.Add("Province", typeof(string));
                Nokia_3G_CS_Table.Columns.Add("Province", typeof(string));
                Ericsson_3G_CS_Table.Columns.Add("Site", typeof(string));
                Huawei_3G_CS_Table.Columns.Add("Site", typeof(string));
                Nokia_3G_CS_Table.Columns.Add("Site", typeof(string));
                Ericsson_3G_CS_Table.Columns.Add("Coverage Type", typeof(string));
                Huawei_3G_CS_Table.Columns.Add("Coverage Type", typeof(string));
                Nokia_3G_CS_Table.Columns.Add("Coverage Type", typeof(string));


                string province_letter = "";
                string cell = "";
                for (int i = 0; i < Ericsson_3G_CS_Table.Rows.Count; i++)
                {
                    cell = Ericsson_3G_CS_Table.Rows[i][2].ToString();
                    Ericsson_3G_CS_Table.Rows[i][10] = "Ericsson";
                    if (cell != "" && cell.Length == 10)
                    {
                        Ericsson_3G_CS_Table.Rows[i][13] = cell.Substring(0, 2) + cell.Substring(4, 4);
                    }
                    else
                    {
                        Ericsson_3G_CS_Table.Rows[i][13] = "NA";
                    }
                }
                for (int i = 0; i < Huawei_3G_CS_Table.Rows.Count; i++)
                {
                    cell = Huawei_3G_CS_Table.Rows[i][2].ToString();
                    Huawei_3G_CS_Table.Rows[i][10] = "Huawei";
                    if (cell != "" && cell.Length == 10)
                    {
                        Huawei_3G_CS_Table.Rows[i][13] = cell.Substring(0, 2) + cell.Substring(4, 4);
                    }
                    else
                    {
                        Huawei_3G_CS_Table.Rows[i][13] = "NA";
                    }
                }
                for (int i = 0; i < Nokia_3G_CS_Table.Rows.Count; i++)
                {
                    cell = Nokia_3G_CS_Table.Rows[i][2].ToString();
                    Nokia_3G_CS_Table.Rows[i][10] = "Nokia";
                    if (cell != "" && cell.Length == 10)
                    {
                        Nokia_3G_CS_Table.Rows[i][13] = cell.Substring(0, 2) + cell.Substring(4, 4);
                    }
                    else
                    {
                        Nokia_3G_CS_Table.Rows[i][13] = "NA";
                    }
                }


                //Union Tables in 3 Vendor
                dtUnion_3G_CS = Ericsson_3G_CS_Table.AsEnumerable().Union(Huawei_3G_CS_Table.AsEnumerable()).CopyToDataTable<DataRow>();
                Table_3G_CS = dtUnion_3G_CS.AsEnumerable().Union(Nokia_3G_CS_Table.AsEnumerable()).CopyToDataTable<DataRow>();




                // join with ARAS Table to get Coverge Type
                var JoinResult = (from p in Table_3G_CS.AsEnumerable()
                                  join t in ARAS_Table.AsEnumerable()
                                  on p.Field<string>("Site") equals t.Field<string>("LOCATION")
                                  select new
                                  {
                                      RNC = p.Field<string>("RNC"),
                                      Site = p.Field<string>("Site"),
                                      Coverage = t.Field<string>("COVERAGE_TYPE_OPTIMIZATION")
                                  }).ToList();




                // Update Province and Contractor
                progressBar1.Minimum = 0;
                progressBar1.Maximum = Table_3G_CS.Rows.Count;
                int Row_Count_of_WPC_Table = 0;
                // for (int i = 0; i < 1000; i++)
                for (int i = 0; i < Table_3G_CS.Rows.Count; i++)
                {
                    string Vendor = Table_3G_CS.Rows[i][10].ToString();
                    string Site = Table_3G_CS.Rows[i][13].ToString();
                    int Found = 0;
                    int ind = 0;
                    string Coverage_type = "";
                    for (int k = 0; k < JoinResult.Count; k++)
                    {
                        if (Site == JoinResult[k].Site.ToString())
                        {
                            Found = 1;
                            ind = k;
                            Coverage_type = JoinResult[ind].Coverage.ToString();
                            break;
                        }
                    }
                    if (Found == 0)
                    {
                        Coverage_type = "City";
                    }


                    cell = Table_3G_CS.Rows[i][2].ToString();
                    if (cell != "")
                    {
                        province_letter = cell.Substring(0, 2);
                        if (province_letter == "TH")
                        {
                            Table_3G_CS.Rows[i][11] = "NAK-Tehran"; Table_3G_CS.Rows[i][12] = "Tehran"; Table_3G_CS.Rows[i][14] = Coverage_type;
                        }
                        if (province_letter == "KJ")
                        {
                            Table_3G_CS.Rows[i][11] = "NAK-Alborz"; Table_3G_CS.Rows[i][12] = "Alborz"; Table_3G_CS.Rows[i][14] = Coverage_type;
                        }
                        if (province_letter == "MA")
                        {
                            Table_3G_CS.Rows[i][11] = "NAK-North"; Table_3G_CS.Rows[i][12] = "Mazandaran"; Table_3G_CS.Rows[i][14] = Coverage_type;
                        }
                        if (province_letter == "GN")
                        {
                            Table_3G_CS.Rows[i][11] = "NAK-North"; Table_3G_CS.Rows[i][12] = "Gilan"; Table_3G_CS.Rows[i][14] = Coverage_type;
                        }
                        if (province_letter == "GL")
                        {
                            Table_3G_CS.Rows[i][11] = "NAK-North"; Table_3G_CS.Rows[i][12] = "Golestan"; Table_3G_CS.Rows[i][14] = Coverage_type;
                        }
                        if (province_letter == "AS")
                        {
                            Table_3G_CS.Rows[i][11] = "NAK-Huawei"; Table_3G_CS.Rows[i][12] = "East Azarbaijan"; Table_3G_CS.Rows[i][14] = Coverage_type;
                        }
                        if (province_letter == "AG")
                        {
                            Table_3G_CS.Rows[i][11] = "NAK-Huawei"; Table_3G_CS.Rows[i][12] = "West Azarbaijan"; Table_3G_CS.Rows[i][14] = Coverage_type;
                        }
                        if (province_letter == "KZ")
                        {
                            Table_3G_CS.Rows[i][11] = "NAK-Huawei"; Table_3G_CS.Rows[i][12] = "Khuzestan"; Table_3G_CS.Rows[i][14] = Coverage_type;
                        }
                        if (province_letter == "KH")
                        {
                            Table_3G_CS.Rows[i][11] = "NAK-Nokia"; Table_3G_CS.Rows[i][12] = "Khorasan Razavi"; Table_3G_CS.Rows[i][14] = Coverage_type;
                        }
                        if (province_letter == "YZ")
                        {
                            Table_3G_CS.Rows[i][11] = "NAK-Nokia"; Table_3G_CS.Rows[i][12] = "Yazd"; Table_3G_CS.Rows[i][14] = Coverage_type;
                        }
                        if (province_letter == "SM")
                        {
                            Table_3G_CS.Rows[i][11] = "NAK-Nokia"; Table_3G_CS.Rows[i][12] = "Semnan"; Table_3G_CS.Rows[i][14] = Coverage_type;
                        }
                        if (province_letter == "CH")
                        {
                            Table_3G_CS.Rows[i][11] = "NAK-Nokia"; Table_3G_CS.Rows[i][12] = "Chahar Mahal Va Bakhtiari"; Table_3G_CS.Rows[i][14] = Coverage_type;
                        }
                        if (province_letter == "KM")
                        {
                            Table_3G_CS.Rows[i][11] = "NAK-Nokia"; Table_3G_CS.Rows[i][12] = "Kerman"; Table_3G_CS.Rows[i][14] = Coverage_type;
                        }
                    }

                    string Contractor = Table_3G_CS.Rows[i][11].ToString();
                    string Province = Table_3G_CS.Rows[i][12].ToString();
                    string RNC = Table_3G_CS.Rows[i][1].ToString();
                    double Traffic = -1000;
                    double CS_RAB_Establish = -1000;
                    double CS_IRAT_HO_SR = -1000;
                    double CS_Drop_Rate = -1000;
                    double Soft_HO_SR = -1000;
                    double CS_RRC_SR = -1000;
                    double Availability = -1000;


                    if (Table_3G_CS.Rows[i][3].ToString() != "")
                    {
                        Traffic = Convert.ToDouble(Table_3G_CS.Rows[i][3]);
                    }
                    if (Table_3G_CS.Rows[i][5].ToString() != "")
                    {
                        CS_RAB_Establish = Convert.ToDouble(Table_3G_CS.Rows[i][5]);
                    }
                    if (Table_3G_CS.Rows[i][6].ToString() != "")
                    {
                        CS_IRAT_HO_SR = Convert.ToDouble(Table_3G_CS.Rows[i][6]);
                    }
                    if (Table_3G_CS.Rows[i][7].ToString() != "")
                    {
                        CS_Drop_Rate = Convert.ToDouble(Table_3G_CS.Rows[i][7]);
                    }
                    if (Table_3G_CS.Rows[i][8].ToString() != "")
                    {
                        Soft_HO_SR = Convert.ToDouble(Table_3G_CS.Rows[i][8]);
                    }
                    if (Table_3G_CS.Rows[i][9].ToString() != "")
                    {
                        CS_RRC_SR = Convert.ToDouble(Table_3G_CS.Rows[i][9]);
                    }
                    if (Table_3G_CS.Rows[i][4].ToString() != "")
                    {
                        Availability = Convert.ToDouble(Table_3G_CS.Rows[i][4]);
                    }

                    // Fill WPc Table
                    int TH_Index = 0;


                    var q1 = (from p in BASE_Table.AsEnumerable()
                              where p.Field<string>("Cell") == cell
                              select p).ToList();
                    var q2 = q1;

                    //string RRC_Connection_SR_BL = "";
                    //string ERAB_SR_Initial_BL = "";
                    //string ERAB_SR_Added_BL = "";
                    //string DL_THR_BL = "";
                    //string UL_THR_BL = "";
                    //string HO_SR_BL = "";
                    //string ERAB_Drop_Rate_BL = "";
                    //string S1_Signalling_SR_BL = "";
                    //string Inter_Freq_SR_BL = "";
                    //string Intra_Freq_SR_BL = "";
                    //if (q1.Count != 0)
                    //{

                    //}




                    if (Vendor == "Ericsson" && Coverage_type == "City")
                    {
                        q1 = (from p in Ericsson_3G_CS_Table_7.AsEnumerable()
                              where p.Field<string>("ElementID1") == cell
                              select p).ToList();

                        TH_Index = 0;
                    }
                    if (Vendor == "Ericsson" && Coverage_type == "None_City")
                    {
                        q1 = (from p in Ericsson_3G_CS_Table_7.AsEnumerable()
                              where p.Field<string>("ElementID1") == cell
                              select p).ToList();

                        TH_Index = 1;
                    }
                    if (Vendor == "Huawei" && Coverage_type == "City")
                    {
                        q1 = (from p in Huawei_3G_CS_Table_7.AsEnumerable()
                              where p.Field<string>("ElementID1") == cell
                              select p).ToList();

                        TH_Index = 2;
                    }
                    if (Vendor == "Huawei" && Coverage_type == "None_City")
                    {
                        q1 = (from p in Huawei_3G_CS_Table_7.AsEnumerable()
                              where p.Field<string>("ElementID1") == cell
                              select p).ToList();

                        TH_Index = 3;
                    }
                    if (Vendor == "Nokia" && Coverage_type == "City")
                    {
                        q1 = (from p in Nokia_3G_CS_Table_7.AsEnumerable()
                              where p.Field<string>("ElementID1") == cell
                              select p).ToList();

                        TH_Index = 4;
                    }
                    if (Vendor == "Nokia" && Coverage_type == "None_City")
                    {
                        q1 = (from p in Nokia_3G_CS_Table_7.AsEnumerable()
                              where p.Field<string>("ElementID1") == cell
                              select p).ToList();

                        TH_Index = 5;
                    }



                    int count = 0;
                    int update = 0;
                    if (CS_RAB_Establish > -1000 && CS_RAB_Establish < TH_3G_CS[TH_Index, 0])
                    {
                        count++;
                        int KPI_Index = 5;
                        if (q1.Count == 7)
                        {
                            Table_3G_CS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, CS_RAB_Establish, "CS_RAB_Establish", "-", q1[6].ItemArray[KPI_Index], q1[5].ItemArray[KPI_Index], q1[4].ItemArray[KPI_Index], q1[3].ItemArray[KPI_Index], q1[2].ItemArray[KPI_Index], q1[1].ItemArray[KPI_Index], q1[0].ItemArray[KPI_Index]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count != 7 && q1.Count != 0)
                        {
                            double[] KPI_of_Days = new double[7];
                            for (int j = 1; j <= 7; j++)
                            {
                                if (q1.Count - j < 0)
                                {
                                    KPI_of_Days[j - 1] = -20000;
                                }
                                else
                                {
                                    DateTime d1 = Convert.ToDateTime(q1[q1.Count - j].ItemArray[0]);
                                    if (d1 == Date_of_WPC.AddDays(-j))
                                    {
                                        if (q1[q1.Count - j].ItemArray[KPI_Index].ToString() == "")
                                        {
                                            KPI_of_Days[j - 1] = -20000;
                                        }
                                        else
                                        {
                                            KPI_of_Days[j - 1] = Convert.ToDouble(q1[q1.Count - j].ItemArray[KPI_Index]);
                                        }

                                    }
                                    else
                                    {
                                        KPI_of_Days[j - 1] = -20000;
                                    }
                                }
                            }
                            Table_3G_CS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, CS_RAB_Establish, "CS_RAB_Establish", "-", KPI_of_Days[0], KPI_of_Days[1], KPI_of_Days[2], KPI_of_Days[3], KPI_of_Days[4], KPI_of_Days[5], KPI_of_Days[6]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count == 0)
                        {
                            Table_3G_CS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, CS_RAB_Establish, "CS_RAB_Establish");
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                    }
                    if (CS_IRAT_HO_SR > -1000 && CS_IRAT_HO_SR < TH_3G_CS[TH_Index, 1])
                    {
                        count++;
                        int KPI_Index = 6;
                        if (q1.Count == 7)
                        {
                            Table_3G_CS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, CS_IRAT_HO_SR, "CS_IRAT_HO_SR", "-", q1[6].ItemArray[KPI_Index], q1[5].ItemArray[KPI_Index], q1[4].ItemArray[KPI_Index], q1[3].ItemArray[KPI_Index], q1[2].ItemArray[KPI_Index], q1[1].ItemArray[KPI_Index], q1[0].ItemArray[KPI_Index]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count != 7 && q1.Count != 0)
                        {
                            double[] KPI_of_Days = new double[7];
                            for (int j = 1; j <= 7; j++)
                            {
                                if (q1.Count - j < 0)
                                {
                                    KPI_of_Days[j - 1] = -20000;
                                }
                                else
                                {
                                    DateTime d1 = Convert.ToDateTime(q1[q1.Count - j].ItemArray[0]);
                                    if (d1 == Date_of_WPC.AddDays(-j))
                                    {
                                        if (q1[q1.Count - j].ItemArray[KPI_Index].ToString() == "")
                                        {
                                            KPI_of_Days[j - 1] = -20000;
                                        }
                                        else
                                        {
                                            KPI_of_Days[j - 1] = Convert.ToDouble(q1[q1.Count - j].ItemArray[KPI_Index]);
                                        }

                                    }
                                    else
                                    {
                                        KPI_of_Days[j - 1] = -20000;
                                    }
                                }
                            }
                            Table_3G_CS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, CS_IRAT_HO_SR, "CS_IRAT_HO_SR", "-", KPI_of_Days[0], KPI_of_Days[1], KPI_of_Days[2], KPI_of_Days[3], KPI_of_Days[4], KPI_of_Days[5], KPI_of_Days[6]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count == 0)
                        {
                            Table_3G_CS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, CS_IRAT_HO_SR, "CS_IRAT_HO_SR");
                            Row_Count_of_WPC_Table++; update = 1;
                        }

                    }
                    if (CS_Drop_Rate > -1000 && CS_Drop_Rate > TH_3G_CS[TH_Index, 2])
                    {
                        count++;
                        int KPI_Index = 7;
                        if (q1.Count == 7)
                        {
                            Table_3G_CS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, CS_Drop_Rate, "CS_Drop_Rate", "-", q1[6].ItemArray[KPI_Index], q1[5].ItemArray[KPI_Index], q1[4].ItemArray[KPI_Index], q1[3].ItemArray[KPI_Index], q1[2].ItemArray[KPI_Index], q1[1].ItemArray[KPI_Index], q1[0].ItemArray[KPI_Index]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count != 7 && q1.Count != 0)
                        {
                            double[] KPI_of_Days = new double[7];
                            for (int j = 1; j <= 7; j++)
                            {
                                if (q1.Count - j < 0)
                                {
                                    KPI_of_Days[j - 1] = -20000;
                                }
                                else
                                {
                                    DateTime d1 = Convert.ToDateTime(q1[q1.Count - j].ItemArray[0]);
                                    if (d1 == Date_of_WPC.AddDays(-j))
                                    {
                                        if (q1[q1.Count - j].ItemArray[KPI_Index].ToString() == "")
                                        {
                                            KPI_of_Days[j - 1] = -20000;
                                        }
                                        else
                                        {
                                            KPI_of_Days[j - 1] = Convert.ToDouble(q1[q1.Count - j].ItemArray[KPI_Index]);
                                        }

                                    }
                                    else
                                    {
                                        KPI_of_Days[j - 1] = -20000;
                                    }
                                }
                            }
                            Table_3G_CS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, CS_Drop_Rate, "CS_Drop_Rate", "-", KPI_of_Days[0], KPI_of_Days[1], KPI_of_Days[2], KPI_of_Days[3], KPI_of_Days[4], KPI_of_Days[5], KPI_of_Days[6]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count == 0)
                        {
                            Table_3G_CS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, CS_Drop_Rate, "CS_Drop_Rate");
                            Row_Count_of_WPC_Table++; update = 1;
                        }

                    }
                    if (Soft_HO_SR > -1000 && Soft_HO_SR < TH_3G_CS[TH_Index, 3])
                    {
                        count++;
                        int KPI_Index = 8;
                        if (q1.Count == 7)
                        {
                            Table_3G_CS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, Soft_HO_SR, "Soft_HO_SR", "-", q1[6].ItemArray[KPI_Index], q1[5].ItemArray[KPI_Index], q1[4].ItemArray[KPI_Index], q1[3].ItemArray[KPI_Index], q1[2].ItemArray[KPI_Index], q1[1].ItemArray[KPI_Index], q1[0].ItemArray[KPI_Index]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count != 7 && q1.Count != 0)
                        {
                            double[] KPI_of_Days = new double[7];
                            for (int j = 1; j <= 7; j++)
                            {
                                if (q1.Count - j < 0)
                                {
                                    KPI_of_Days[j - 1] = -20000;
                                }
                                else
                                {
                                    DateTime d1 = Convert.ToDateTime(q1[q1.Count - j].ItemArray[0]);
                                    if (d1 == Date_of_WPC.AddDays(-j))
                                    {
                                        if (q1[q1.Count - j].ItemArray[KPI_Index].ToString() == "")
                                        {
                                            KPI_of_Days[j - 1] = -20000;
                                        }
                                        else
                                        {
                                            KPI_of_Days[j - 1] = Convert.ToDouble(q1[q1.Count - j].ItemArray[KPI_Index]);
                                        }

                                    }
                                    else
                                    {
                                        KPI_of_Days[j - 1] = -20000;
                                    }
                                }
                            }
                            Table_3G_CS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, Soft_HO_SR, "Soft_HO_SR", "-", KPI_of_Days[0], KPI_of_Days[1], KPI_of_Days[2], KPI_of_Days[3], KPI_of_Days[4], KPI_of_Days[5], KPI_of_Days[6]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count == 0)
                        {
                            Table_3G_CS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, Soft_HO_SR, "Soft_HO_SR");
                            Row_Count_of_WPC_Table++; update = 1;
                        }

                    }
                    if (CS_RRC_SR > -1000 && CS_RRC_SR < TH_3G_CS[TH_Index, 4])
                    {
                        count++;
                        int KPI_Index = 9;
                        if (q1.Count == 7)
                        {
                            Table_3G_CS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, CS_RRC_SR, "CS_RRC_SR", "-", q1[6].ItemArray[KPI_Index], q1[5].ItemArray[KPI_Index], q1[4].ItemArray[KPI_Index], q1[3].ItemArray[KPI_Index], q1[2].ItemArray[KPI_Index], q1[1].ItemArray[KPI_Index], q1[0].ItemArray[KPI_Index]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count != 7 && q1.Count != 0)
                        {
                            double[] KPI_of_Days = new double[7];
                            for (int j = 1; j <= 7; j++)
                            {
                                if (q1.Count - j < 0)
                                {
                                    KPI_of_Days[j - 1] = -20000;
                                }
                                else
                                {
                                    DateTime d1 = Convert.ToDateTime(q1[q1.Count - j].ItemArray[0]);
                                    if (d1 == Date_of_WPC.AddDays(-j))
                                    {
                                        if (q1[q1.Count - j].ItemArray[KPI_Index].ToString() == "")
                                        {
                                            KPI_of_Days[j - 1] = -20000;
                                        }
                                        else
                                        {
                                            KPI_of_Days[j - 1] = Convert.ToDouble(q1[q1.Count - j].ItemArray[KPI_Index]);
                                        }

                                    }
                                    else
                                    {
                                        KPI_of_Days[j - 1] = -20000;
                                    }
                                }
                            }
                            Table_3G_CS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, CS_RRC_SR, "CS_RRC_SR", "-", KPI_of_Days[0], KPI_of_Days[1], KPI_of_Days[2], KPI_of_Days[3], KPI_of_Days[4], KPI_of_Days[5], KPI_of_Days[6]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count == 0)
                        {
                            Table_3G_CS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Traffic, CS_RRC_SR, "CS_RRC_SR");
                            Row_Count_of_WPC_Table++; update = 1;
                        }

                    }


                    if (update == 1)
                    {
                        Table_3G_CS_WPC.Rows[Row_Count_of_WPC_Table - 1][12] = Convert.ToString(count);
                    }

                    //label5.Text = Convert.ToString(Math.Round(Convert.ToDouble(i / Table_3G_CS.Rows.Count)));
                    progressBar1.Value = i;

                }




                XLWorkbook wb = new XLWorkbook();
                wb.Worksheets.Add(Table_3G_CS_WPC, "WPC_3G_CS");
                var saveFileDialog = new SaveFileDialog
                {
                    FileName = "WPC_3G_CS",
                    Filter = "Excel files|*.xlsx",
                    Title = "Save an Excel File"
                };

                saveFileDialog.ShowDialog();

                if (!String.IsNullOrWhiteSpace(saveFileDialog.FileName))
                    wb.SaveAs(saveFileDialog.FileName);




                label4.Text = "Finished";
                label4.BackColor = Color.Yellow;
                MessageBox.Show("FINISHED");


            }



            // Quary of Select Data in 3G_PS
            if (checkBox4.Checked == true)
            {
                string Ericsson_3G_PS_TBL = @"select [Date], [ElementID] as 'RNC', [ElementID1], [PS_Volume(GB)(UCell_Eric)] as 'PS_Traffic_Daily (GB)', [Cell_Availability_Rate_Exclude_Blocking(UCELL_Eric)] as 'Availability', [HSDPA_RAB_Setup_Succ_Rate(UCell_Eric)] as 'HSDPA_SR', [HSUPA_Setup_Success_Rate(UCell_Eric)] as 'HSUPA_SR', [HSUPA_User_Throughput_MACe(Kbps)(UCell_Eric)] as 'UL_User_THR', [HS_USER_Throughput_NET_PQ(Mbps)(UCell_Eric)] as 'DL_User_THR', 
[HSDPA_Drop_Call_Rate(UCell_Eric)] as 'HSDPA_Drop_Rate', [HSUPA_Drop_Call_Rate(UCell_Eric)] as 'HSUAP_Drop_Rate', [PS_Multi_RAB_Establish_Success_Rate(without_Nas)(UCELL_Eric)] as 'MultiRAB_SR' ,  [PS_RRC_Setup_Success_Rate(UCell_Eric)] as ' PS_RRC_SR',  [Ps_RAB_Establish_Success_Rate] as ' Ps_RAB_Establish', [Ps_RAB_Establish_Success_Rate(UCell_Eric)] as ' PS_MultiRAB_Establish',  [PS_Drop_Call_Rate(UCell_Eric)] as 'PS_Drop_Rate', [HSDPA_Cell_Change_Succ_Rate(UCell_Eric)] as ' HSDPA_Cell_Change_SR',  [HS_share_PAYLOAD_Rate(UCell_Eric)] as ' HS_Share_Payload', [HSDPA_Cell_Scheduled_Throughput(mbps)(UCell_Eric)] as 'DL_Cell_THR'   from [dbo].[RD3_Ericsson_Cell_Daily] where 
(substring([ElementID1],1,2)='KJ' or substring([ElementID1],1,2)='CH' or substring([ElementID1],1,2)='AS' or substring([ElementID1],1,2)='GL' or 
                                                                                  substring([ElementID1],1,2)='GN'  or substring([ElementID1],1,2)='KM' or substring([ElementID1],1,2)='KH' or substring([ElementID1],1,2)='KZ' or substring([ElementID1],1,2)='MA'
																				  or substring([ElementID1],1,2)='SM' or substring([ElementID1],1,2)='TH' or substring([ElementID1],1,2)='AG' or substring([ElementID1],1,2)='YZ')
and ([HSDPA_RAB_Setup_Succ_Rate(UCell_Eric)]<97.812 or [HSUPA_Setup_Success_Rate(UCell_Eric)]<97.128 or [HSUPA_User_Throughput_MACe(Kbps)(UCell_Eric)]<122.828 or [HS_USER_Throughput_NET_PQ(Mbps)(UCell_Eric)]<1.056 or
[HSDPA_Drop_Call_Rate(UCell_Eric)]>2.106 or [HSUPA_Drop_Call_Rate(UCell_Eric)]>2.133 or [PS_Multi_RAB_Establish_Success_Rate(without_Nas)(UCELL_Eric)]<99.153 or [PS_RRC_Setup_Success_Rate(UCell_Eric)]<99.066 or
[Ps_RAB_Establish_Success_Rate]<97.805 or [Ps_RAB_Establish_Success_Rate(UCell_Eric)]<97.28 or [PS_Drop_Call_Rate(UCell_Eric)]>2.098 or [HSDPA_Cell_Change_Succ_Rate(UCell_Eric)]<99.745 or 
[HS_share_PAYLOAD_Rate(UCell_Eric)]<99.593 or [HSDPA_Cell_Scheduled_Throughput(mbps)(UCell_Eric)]<2.004) and Date  ='" + Date_of_WPC + "'";


                SqlCommand Ericsson_3G_PS_TBL_Quary = new SqlCommand(Ericsson_3G_PS_TBL, connection);
                Ericsson_3G_PS_TBL_Quary.ExecuteNonQuery();
                DataTable Ericsson_3G_PS_Table = new DataTable();
                SqlDataAdapter dataAdapter_Ericsson_3G_PS = new SqlDataAdapter(Ericsson_3G_PS_TBL_Quary);
                dataAdapter_Ericsson_3G_PS.Fill(Ericsson_3G_PS_Table);




                string Huawei_3G_PS_TBL = @"select [Date], [ElementID] as 'RNC', [ElementID1], [PAYLOAD] as 'PS_Traffic_Daily (GB)', [Radio_Network_Availability_Ratio(Hu_Cell)] as 'Availability', [HSDPA_RAB_Setup_Success_Ratio(Hu_Cell)] as 'HSDPA_SR', [HSUPA_RAB_Setup_Success_Ratio(Hu_Cell)] as 'HSUPA_SR', [hsupa_uplink_throughput_in_V16(CELL_Hu)] as 'UL_User_THR', [AVERAGE_HSDPA_USER_THROUGHPUT_DC+SC(Mbit/s)(CELL_HUAWEI)] as 'DL_User_THR', 
[HSDPA_cdr(%)_(Hu_Cell)_new] as 'HSDPA_Drop_Rate', [HSUPA_CDR(%)_(Hu_Cell)_new] as 'HSUAP_Drop_Rate', [CS+PS_RAB_Setup_Success_Ratio] as 'MultiRAB_SR' ,  [PS_RRC_Connection_success_Rate_repeatless(Hu_Cell)] as ' PS_RRC_SR',  [PS_RAB_Setup_Success_Ratio] as ' Ps_RAB_Establish', [PS_RAB_Setup_Success_Ratio(Hu_Cell)] as ' PS_MultiRAB_Establish',  [PS_Call_Drop_Ratio] as 'PS_Drop_Rate', [HSDPA_Soft_HandOver_Success_Ratio] as ' HSDPA_Cell_Change_SR',  [HS_share_PAYLOAD_%] as ' HS_Share_Payload', [HSDPA_SCHEDULING_Cell_throughput(CELL_HUAWEI)] as 'DL_Cell_THR'   from [dbo].[RD3_Huawei_Cell_Daily] where 
(substring([ElementID1],1,2)='KJ' or substring([ElementID1],1,2)='CH' or substring([ElementID1],1,2)='AS' or substring([ElementID1],1,2)='GL' or 
                                                                                  substring([ElementID1],1,2)='GN'  or substring([ElementID1],1,2)='KM' or substring([ElementID1],1,2)='KH' or substring([ElementID1],1,2)='KZ' or substring([ElementID1],1,2)='MA'
																				  or substring([ElementID1],1,2)='SM' or substring([ElementID1],1,2)='TH' or substring([ElementID1],1,2)='AG' or substring([ElementID1],1,2)='YZ')
and ([HSDPA_RAB_Setup_Success_Ratio(Hu_Cell)]<99.63 or [HSUPA_RAB_Setup_Success_Ratio(Hu_Cell)]<99.628 or [hsupa_uplink_throughput_in_V16(CELL_Hu)]<150.299 or [AVERAGE_HSDPA_USER_THROUGHPUT_DC+SC(Mbit/s)(CELL_HUAWEI)]<1.486 or
[HSDPA_cdr(%)_(Hu_Cell)_new]>0.168 or [HSUPA_CDR(%)_(Hu_Cell)_new]>0.225 or [CS+PS_RAB_Setup_Success_Ratio]<99.756 or [PS_RRC_Connection_success_Rate_repeatless(Hu_Cell)]<99.514 or
[PS_RAB_Setup_Success_Ratio]<99.563 or [PS_RAB_Setup_Success_Ratio(Hu_Cell)]<98.892 or [PS_Call_Drop_Ratio]>0.213 or [HSDPA_Soft_HandOver_Success_Ratio]<99.619 or 
[HS_share_PAYLOAD_%]<99.053 or [HSDPA_SCHEDULING_Cell_throughput(CELL_HUAWEI)]<2.045) and Date  ='" + Date_of_WPC + "'";


                SqlCommand Huawei_3G_PS_TBL_Quary = new SqlCommand(Huawei_3G_PS_TBL, connection);
                Huawei_3G_PS_TBL_Quary.ExecuteNonQuery();
                DataTable Huawei_3G_PS_Table = new DataTable();
                SqlDataAdapter dataAdapter_Huawei_3G_PS = new SqlDataAdapter(Huawei_3G_PS_TBL_Quary);
                dataAdapter_Huawei_3G_PS.Fill(Huawei_3G_PS_Table);




                string Nokia_3G_PS_TBL = @"select [Date], [ElementID] as 'RNC', [ElementID1], [PS_Payload_Total(HS+R99)(Nokia_CELL)_GB] as 'PS_Traffic_Daily (GB)', [Cell_Availability_excluding_blocked_by_user_state(Nokia_UCell)] as 'Availability', [HSDPA_setup_success_ratio_from_user_perspective(CELL_Nokia)] as 'HSDPA_SR', [HSUPA_Setup_Success_Ratio_from_user_perspective(CELL)] as 'HSUPA_SR', [Average_hsupa_throughput_MACe(nokia_cell)] as 'UL_User_THR', [AVERAGE_HSDPA_USER_THROUGHPUT_DC+SC(Mbit/s)(Nokia_CELL)] as 'DL_User_THR', 
[HSDPA_Call_Drop_Rate(Nokia_Cell)] as 'HSDPA_Drop_Rate', [HSUPA_Call_Drop_Rate(Nokia_CELL)] as 'HSUAP_Drop_Rate', [CSAMR+PS_MRAB_stp_SR(Nokia_CELL)] as 'MultiRAB_SR' ,  [PS_RRCSETUP_SR] as ' PS_RRC_SR',  [PS_RAB_Setup_Success_Ratio] as ' Ps_RAB_Establish', [RAB_Setup_and_Access_Complete_Ratio_for_NRT_Service_from_User_pe] as ' PS_MultiRAB_Establish',  [Packet_Session_Drop_Ratio_NOKIA(CELL_NOKIA)] as 'PS_Drop_Rate', [HSDPA_Cell_Change_SR(Nokia_CELL)] as ' HSDPA_Cell_Change_SR',  [HS_SHARE_PAYLOAD(Nokia_CELL)] as ' HS_Share_Payload', [Active_HS-DSCH_cell_throughput_mbs(CELL_nokia)] as 'DL_Cell_THR'   from [dbo].[RD3_Nokia_Cell_Daily] where 
(substring([ElementID1],1,2)='KJ' or substring([ElementID1],1,2)='CH' or substring([ElementID1],1,2)='AS' or substring([ElementID1],1,2)='GL' or 
                                                                                  substring([ElementID1],1,2)='GN'  or substring([ElementID1],1,2)='KM' or substring([ElementID1],1,2)='KH' or substring([ElementID1],1,2)='KZ' or substring([ElementID1],1,2)='MA'
																				  or substring([ElementID1],1,2)='SM' or substring([ElementID1],1,2)='TH' or substring([ElementID1],1,2)='AG' or substring([ElementID1],1,2)='YZ')
and ([HSDPA_setup_success_ratio_from_user_perspective(CELL_Nokia)]<97.106 or [HSUPA_Setup_Success_Ratio_from_user_perspective(CELL)]<96.325 or [Average_hsupa_throughput_MACe(nokia_cell)]<126.347 or 
[HSDPA_Call_Drop_Rate(Nokia_Cell)]>1.234 or [HSUPA_Call_Drop_Rate(Nokia_CELL)]>1.079 or [CSAMR+PS_MRAB_stp_SR(Nokia_CELL)]<98.808 or [PS_RRCSETUP_SR]<97.857 or
[PS_RAB_Setup_Success_Ratio]<98.868 or [RAB_Setup_and_Access_Complete_Ratio_for_NRT_Service_from_User_pe]<97.282 or [Packet_Session_Drop_Ratio_NOKIA(CELL_NOKIA)]>1.613 or [HSDPA_Cell_Change_SR(Nokia_CELL)]<99.172 or 
[HS_SHARE_PAYLOAD(Nokia_CELL)]<99.617 or [Active_HS-DSCH_cell_throughput_mbs(CELL_nokia)]<1.979) and Date  ='" + Date_of_WPC + "'";


                SqlCommand Nokia_3G_PS_TBL_Quary = new SqlCommand(Nokia_3G_PS_TBL, connection);
                Nokia_3G_PS_TBL_Quary.ExecuteNonQuery();
                DataTable Nokia_3G_PS_Table = new DataTable();
                SqlDataAdapter dataAdapter_Nokia_3G_PS = new SqlDataAdapter(Nokia_3G_PS_TBL_Quary);
                dataAdapter_Nokia_3G_PS.Fill(Nokia_3G_PS_Table);



                // Table of Oldest 7 Dayes

                string Ericsson_3G_PS_TBL_7 = @"select [Date], [ElementID] as 'RNC', [ElementID1], [PS_Volume(GB)(UCell_Eric)] as 'PS_Traffic_Daily (GB)', [Cell_Availability_Rate_Exclude_Blocking(UCELL_Eric)] as 'Availability', [HSDPA_RAB_Setup_Succ_Rate(UCell_Eric)] as 'HSDPA_SR', [HSUPA_Setup_Success_Rate(UCell_Eric)] as 'HSUPA_SR', [HSUPA_User_Throughput_MACe(Kbps)(UCell_Eric)] as 'UL_User_THR', [HS_USER_Throughput_NET_PQ(Mbps)(UCell_Eric)] as 'DL_User_THR', 
[HSDPA_Drop_Call_Rate(UCell_Eric)] as 'HSDPA_Drop_Rate', [HSUPA_Drop_Call_Rate(UCell_Eric)] as 'HSUAP_Drop_Rate', [PS_Multi_RAB_Establish_Success_Rate(without_Nas)(UCELL_Eric)] as 'MultiRAB_SR' ,  [PS_RRC_Setup_Success_Rate(UCell_Eric)] as ' PS_RRC_SR',  [Ps_RAB_Establish_Success_Rate] as ' Ps_RAB_Establish', [Ps_RAB_Establish_Success_Rate(UCell_Eric)] as ' PS_MultiRAB_Establish',  [PS_Drop_Call_Rate(UCell_Eric)] as 'PS_Drop_Rate', [HSDPA_Cell_Change_Succ_Rate(UCell_Eric)] as ' HSDPA_Cell_Change_SR',  [HS_share_PAYLOAD_Rate(UCell_Eric)] as ' HS_Share_Payload', [HSDPA_Cell_Scheduled_Throughput(mbps)(UCell_Eric)] as 'DL_Cell_THR'   from [dbo].[RD3_Ericsson_Cell_Daily] where 
(substring([ElementID1],1,2)='KJ' or substring([ElementID1],1,2)='CH' or substring([ElementID1],1,2)='AS' or substring([ElementID1],1,2)='GL' or 
                                                                                  substring([ElementID1],1,2)='GN'  or substring([ElementID1],1,2)='KM' or substring([ElementID1],1,2)='KH' or substring([ElementID1],1,2)='KZ' or substring([ElementID1],1,2)='MA'
																				  or substring([ElementID1],1,2)='SM' or substring([ElementID1],1,2)='TH' or substring([ElementID1],1,2)='AG' or substring([ElementID1],1,2)='YZ')
and Date  >='" + Date_of_WPC_7 + "' and Date  <'" + Date_of_WPC + "'";


                SqlCommand Ericsson_3G_PS_TBL_Quary_7 = new SqlCommand(Ericsson_3G_PS_TBL_7, connection);
                Ericsson_3G_PS_TBL_Quary_7.CommandTimeout = 0;
                Ericsson_3G_PS_TBL_Quary_7.ExecuteNonQuery();
                DataTable Ericsson_3G_PS_Table_7 = new DataTable();
                SqlDataAdapter dataAdapter_Ericsson_3G_PS_7 = new SqlDataAdapter(Ericsson_3G_PS_TBL_Quary_7);
                dataAdapter_Ericsson_3G_PS_7.Fill(Ericsson_3G_PS_Table_7);




                string Huawei_3G_PS_TBL_7 = @"select [Date], [ElementID] as 'RNC', [ElementID1], [PAYLOAD] as 'PS_Traffic_Daily (GB)', [Radio_Network_Availability_Ratio(Hu_Cell)] as 'Availability', [HSDPA_RAB_Setup_Success_Ratio(Hu_Cell)] as 'HSDPA_SR', [HSUPA_RAB_Setup_Success_Ratio(Hu_Cell)] as 'HSUPA_SR', [hsupa_uplink_throughput_in_V16(CELL_Hu)] as 'UL_User_THR', [AVERAGE_HSDPA_USER_THROUGHPUT_DC+SC(Mbit/s)(CELL_HUAWEI)] as 'DL_User_THR', 
[HSDPA_cdr(%)_(Hu_Cell)_new] as 'HSDPA_Drop_Rate', [HSUPA_CDR(%)_(Hu_Cell)_new] as 'HSUAP_Drop_Rate', [CS+PS_RAB_Setup_Success_Ratio] as 'MultiRAB_SR' ,  [PS_RRC_Connection_success_Rate_repeatless(Hu_Cell)] as ' PS_RRC_SR',  [PS_RAB_Setup_Success_Ratio] as ' Ps_RAB_Establish', [PS_RAB_Setup_Success_Ratio(Hu_Cell)] as ' PS_MultiRAB_Establish',  [PS_Call_Drop_Ratio] as 'PS_Drop_Rate', [HSDPA_Soft_HandOver_Success_Ratio] as ' HSDPA_Cell_Change_SR',  [HS_share_PAYLOAD_%] as ' HS_Share_Payload', [HSDPA_SCHEDULING_Cell_throughput(CELL_HUAWEI)] as 'DL_Cell_THR'   from [dbo].[RD3_Huawei_Cell_Daily] where 
(substring([ElementID1],1,2)='KJ' or substring([ElementID1],1,2)='CH' or substring([ElementID1],1,2)='AS' or substring([ElementID1],1,2)='GL' or 
                                                                                  substring([ElementID1],1,2)='GN'  or substring([ElementID1],1,2)='KM' or substring([ElementID1],1,2)='KH' or substring([ElementID1],1,2)='KZ' or substring([ElementID1],1,2)='MA'
																				  or substring([ElementID1],1,2)='SM' or substring([ElementID1],1,2)='TH' or substring([ElementID1],1,2)='AG' or substring([ElementID1],1,2)='YZ')
and Date  >='" + Date_of_WPC_7 + "' and Date  <'" + Date_of_WPC + "'";


                SqlCommand Huawei_3G_PS_TBL_Quary_7 = new SqlCommand(Huawei_3G_PS_TBL_7, connection);
                Huawei_3G_PS_TBL_Quary_7.CommandTimeout = 0;
                Huawei_3G_PS_TBL_Quary_7.ExecuteNonQuery();
                DataTable Huawei_3G_PS_Table_7 = new DataTable();
                SqlDataAdapter dataAdapter_Huawei_3G_PS_7 = new SqlDataAdapter(Huawei_3G_PS_TBL_Quary_7);
                dataAdapter_Huawei_3G_PS_7.Fill(Huawei_3G_PS_Table_7);




                string Nokia_3G_PS_TBL_7 = @"select [Date], [ElementID] as 'RNC', [ElementID1], [PS_Payload_Total(HS+R99)(Nokia_CELL)_GB] as 'PS_Traffic_Daily (GB)', [Cell_Availability_excluding_blocked_by_user_state(Nokia_UCell)] as 'Availability', [HSDPA_setup_success_ratio_from_user_perspective(CELL_Nokia)] as 'HSDPA_SR', [HSUPA_Setup_Success_Ratio_from_user_perspective(CELL)] as 'HSUPA_SR', [Average_hsupa_throughput_MACe(nokia_cell)] as 'UL_User_THR', [AVERAGE_HSDPA_USER_THROUGHPUT_DC+SC(Mbit/s)(Nokia_CELL)] as 'DL_User_THR', 
[HSDPA_Call_Drop_Rate(Nokia_Cell)] as 'HSDPA_Drop_Rate', [HSUPA_Call_Drop_Rate(Nokia_CELL)] as 'HSUAP_Drop_Rate', [CSAMR+PS_MRAB_stp_SR(Nokia_CELL)] as 'MultiRAB_SR' ,  [PS_RRCSETUP_SR] as ' PS_RRC_SR',  [PS_RAB_Setup_Success_Ratio] as ' Ps_RAB_Establish', [RAB_Setup_and_Access_Complete_Ratio_for_NRT_Service_from_User_pe] as ' PS_MultiRAB_Establish',  [Packet_Session_Drop_Ratio_NOKIA(CELL_NOKIA)] as 'PS_Drop_Rate', [HSDPA_Cell_Change_SR(Nokia_CELL)] as ' HSDPA_Cell_Change_SR',  [HS_SHARE_PAYLOAD(Nokia_CELL)] as ' HS_Share_Payload', [Active_HS-DSCH_cell_throughput_mbs(CELL_nokia)] as 'DL_Cell_THR'   from [dbo].[RD3_Nokia_Cell_Daily] where 
(substring([ElementID1],1,2)='KJ' or substring([ElementID1],1,2)='CH' or substring([ElementID1],1,2)='AS' or substring([ElementID1],1,2)='GL' or 
                                                                                  substring([ElementID1],1,2)='GN'  or substring([ElementID1],1,2)='KM' or substring([ElementID1],1,2)='KH' or substring([ElementID1],1,2)='KZ' or substring([ElementID1],1,2)='MA'
																				  or substring([ElementID1],1,2)='SM' or substring([ElementID1],1,2)='TH' or substring([ElementID1],1,2)='AG' or substring([ElementID1],1,2)='YZ')
and Date  >='" + Date_of_WPC_7 + "' and Date  <'" + Date_of_WPC + "'";



                SqlCommand Nokia_3G_PS_TBL_Quary_7 = new SqlCommand(Nokia_3G_PS_TBL_7, connection);
                Nokia_3G_PS_TBL_Quary_7.CommandTimeout = 0;
                Nokia_3G_PS_TBL_Quary_7.ExecuteNonQuery();
                DataTable Nokia_3G_PS_Table_7 = new DataTable();
                SqlDataAdapter dataAdapter_Nokia_3G_PS_7 = new SqlDataAdapter(Nokia_3G_PS_TBL_Quary_7);
                dataAdapter_Nokia_3G_PS_7.Fill(Nokia_3G_PS_Table_7);




                // Update Vendor an Site
                Ericsson_3G_PS_Table.Columns.Add("Vendor", typeof(string));
                Huawei_3G_PS_Table.Columns.Add("Vendor", typeof(string));
                Nokia_3G_PS_Table.Columns.Add("Vendor", typeof(string));
                Ericsson_3G_PS_Table.Columns.Add("Contractor", typeof(string));
                Huawei_3G_PS_Table.Columns.Add("Contractor", typeof(string));
                Nokia_3G_PS_Table.Columns.Add("Contractor", typeof(string));
                Ericsson_3G_PS_Table.Columns.Add("Province", typeof(string));
                Huawei_3G_PS_Table.Columns.Add("Province", typeof(string));
                Nokia_3G_PS_Table.Columns.Add("Province", typeof(string));
                Ericsson_3G_PS_Table.Columns.Add("Site", typeof(string));
                Huawei_3G_PS_Table.Columns.Add("Site", typeof(string));
                Nokia_3G_PS_Table.Columns.Add("Site", typeof(string));
                Ericsson_3G_PS_Table.Columns.Add("Coverage Type", typeof(string));
                Huawei_3G_PS_Table.Columns.Add("Coverage Type", typeof(string));
                Nokia_3G_PS_Table.Columns.Add("Coverage Type", typeof(string));


                string province_letter = "";
                string cell = "";
                for (int i = 0; i < Ericsson_3G_PS_Table.Rows.Count; i++)
                {
                    cell = Ericsson_3G_PS_Table.Rows[i][2].ToString();
                    Ericsson_3G_PS_Table.Rows[i][19] = "Ericsson";
                    if (cell != "" && cell.Length == 10)
                    {
                        Ericsson_3G_PS_Table.Rows[i][22] = cell.Substring(0, 2) + cell.Substring(4, 4);
                    }
                    else
                    {
                        Ericsson_3G_PS_Table.Rows[i][22] = "NA";
                    }
                }
                for (int i = 0; i < Huawei_3G_PS_Table.Rows.Count; i++)
                {
                    cell = Huawei_3G_PS_Table.Rows[i][2].ToString();
                    Huawei_3G_PS_Table.Rows[i][19] = "Huawei";
                    if (cell != "" && cell.Length == 10)
                    {
                        Huawei_3G_PS_Table.Rows[i][22] = cell.Substring(0, 2) + cell.Substring(4, 4);
                    }
                    else
                    {
                        Huawei_3G_PS_Table.Rows[i][22] = "NA";
                    }
                }
                for (int i = 0; i < Nokia_3G_PS_Table.Rows.Count; i++)
                {
                    cell = Nokia_3G_PS_Table.Rows[i][2].ToString();
                    Nokia_3G_PS_Table.Rows[i][19] = "Nokia";
                    if (cell != "" && cell.Length == 10)
                    {
                        Nokia_3G_PS_Table.Rows[i][22] = cell.Substring(0, 2) + cell.Substring(4, 4);
                    }
                    else
                    {
                        Nokia_3G_PS_Table.Rows[i][22] = "NA";
                    }
                }


                //Union Tables in 3 Vendor
                dtUnion_3G_PS = Ericsson_3G_PS_Table.AsEnumerable().Union(Huawei_3G_PS_Table.AsEnumerable()).CopyToDataTable<DataRow>();
                Table_3G_PS = dtUnion_3G_PS.AsEnumerable().Union(Nokia_3G_PS_Table.AsEnumerable()).CopyToDataTable<DataRow>();




                // join with ARAS Table to get Coverge Type
                var JoinResult = (from p in Table_3G_PS.AsEnumerable()
                                  join t in ARAS_Table.AsEnumerable()
                                  on p.Field<string>("Site") equals t.Field<string>("LOCATION")
                                  select new
                                  {
                                      RNC = p.Field<string>("RNC"),
                                      Site = p.Field<string>("Site"),
                                      Coverage = t.Field<string>("COVERAGE_TYPE_OPTIMIZATION")
                                  }).ToList();




                // Update Province and Contractor
                progressBar1.Minimum = 0;
                progressBar1.Maximum = Table_3G_PS.Rows.Count;
                int Row_Count_of_WPC_Table = 0;
                // for (int i = 0; i < 1000; i++)
                for (int i = 0; i < Table_3G_PS.Rows.Count; i++)
                {
                    string Vendor = Table_3G_PS.Rows[i][19].ToString();
                    string Site = Table_3G_PS.Rows[i][22].ToString();
                    int Found = 0;
                    int ind = 0;
                    string Coverage_type = "";
                    for (int k = 0; k < JoinResult.Count; k++)
                    {
                        if (Site == JoinResult[k].Site.ToString())
                        {
                            Found = 1;
                            ind = k;
                            Coverage_type = JoinResult[ind].Coverage.ToString();
                            break;
                        }
                    }
                    if (Found == 0)
                    {
                        Coverage_type = "City";
                    }


                    cell = Table_3G_PS.Rows[i][2].ToString();
                    if (cell != "")
                    {
                        province_letter = cell.Substring(0, 2);
                        if (province_letter == "TH")
                        {
                            Table_3G_PS.Rows[i][20] = "NAK-Tehran"; Table_3G_PS.Rows[i][21] = "Tehran"; Table_3G_PS.Rows[i][23] = Coverage_type;
                        }
                        if (province_letter == "KJ")
                        {
                            Table_3G_PS.Rows[i][20] = "NAK-Alborz"; Table_3G_PS.Rows[i][21] = "Alborz"; Table_3G_PS.Rows[i][23] = Coverage_type;
                        }
                        if (province_letter == "MA")
                        {
                            Table_3G_PS.Rows[i][20] = "NAK-North"; Table_3G_PS.Rows[i][21] = "Mazandaran"; Table_3G_PS.Rows[i][23] = Coverage_type;
                        }
                        if (province_letter == "GN")
                        {
                            Table_3G_PS.Rows[i][20] = "NAK-North"; Table_3G_PS.Rows[i][21] = "Gilan"; Table_3G_PS.Rows[i][23] = Coverage_type;
                        }
                        if (province_letter == "GL")
                        {
                            Table_3G_PS.Rows[i][20] = "NAK-North"; Table_3G_PS.Rows[i][21] = "Golestan"; Table_3G_PS.Rows[i][23] = Coverage_type;
                        }
                        if (province_letter == "AS")
                        {
                            Table_3G_PS.Rows[i][20] = "NAK-Huawei"; Table_3G_PS.Rows[i][21] = "East Azarbaijan"; Table_3G_PS.Rows[i][23] = Coverage_type;
                        }
                        if (province_letter == "AG")
                        {
                            Table_3G_PS.Rows[i][20] = "NAK-Huawei"; Table_3G_PS.Rows[i][21] = "West Azarbaijan"; Table_3G_PS.Rows[i][23] = Coverage_type;
                        }
                        if (province_letter == "KZ")
                        {
                            Table_3G_PS.Rows[i][20] = "NAK-Huawei"; Table_3G_PS.Rows[i][21] = "Khuzestan"; Table_3G_PS.Rows[i][23] = Coverage_type;
                        }
                        if (province_letter == "KH")
                        {
                            Table_3G_PS.Rows[i][20] = "NAK-Nokia"; Table_3G_PS.Rows[i][21] = "Khorasan Razavi"; Table_3G_PS.Rows[i][23] = Coverage_type;
                        }
                        if (province_letter == "YZ")
                        {
                            Table_3G_PS.Rows[i][20] = "NAK-Nokia"; Table_3G_PS.Rows[i][21] = "Yazd"; Table_3G_PS.Rows[i][23] = Coverage_type;
                        }
                        if (province_letter == "SM")
                        {
                            Table_3G_PS.Rows[i][20] = "NAK-Nokia"; Table_3G_PS.Rows[i][21] = "Semnan"; Table_3G_PS.Rows[i][23] = Coverage_type;
                        }
                        if (province_letter == "CH")
                        {
                            Table_3G_PS.Rows[i][20] = "NAK-Nokia"; Table_3G_PS.Rows[i][21] = "Chahar Mahal Va Bakhtiari"; Table_3G_PS.Rows[i][23] = Coverage_type;
                        }
                        if (province_letter == "KM")
                        {
                            Table_3G_PS.Rows[i][20] = "NAK-Nokia"; Table_3G_PS.Rows[i][21] = "Kerman"; Table_3G_PS.Rows[i][23] = Coverage_type;
                        }
                    }

                    string Contractor = Table_3G_PS.Rows[i][20].ToString();
                    string Province = Table_3G_PS.Rows[i][21].ToString();
                    string RNC = Table_3G_PS.Rows[i][1].ToString();
                    double Payload = -1000;
                    double HSDPA_SR = -1000;
                    double HSUPA_SR = -1000;
                    double UL_User_THR = -1000;
                    double DL_User_THR = -1000;
                    double HSDPA_Drop_Rate = -1000;
                    double HSUPA_Drop_Rate = -1000;
                    double MultiRAB_SR = -1000;
                    double PS_RRC_SR = -1000;
                    double Ps_RAB_Establish = -1000;
                    double PS_MultiRAB_Establish = -1000;
                    double PS_Drop_Rate = -1000;
                    double HSDPA_Cell_Change_SR = -1000;
                    double HS_Share_Payload = -1000;
                    double DL_Cell_THR = -1000;
                    double Availability = -1000;




                    if (Table_3G_PS.Rows[i][3].ToString() != "")
                    {
                        Payload = Convert.ToDouble(Table_3G_PS.Rows[i][3]);
                    }
                    if (Table_3G_PS.Rows[i][5].ToString() != "")
                    {
                        HSDPA_SR = Convert.ToDouble(Table_3G_PS.Rows[i][5]);
                    }
                    if (Table_3G_PS.Rows[i][6].ToString() != "")
                    {
                        HSUPA_SR = Convert.ToDouble(Table_3G_PS.Rows[i][6]);
                    }
                    if (Table_3G_PS.Rows[i][7].ToString() != "")
                    {
                        UL_User_THR = Convert.ToDouble(Table_3G_PS.Rows[i][7]);
                    }
                    if (Table_3G_PS.Rows[i][8].ToString() != "")
                    {
                        DL_User_THR = Convert.ToDouble(Table_3G_PS.Rows[i][8]);
                    }
                    if (Table_3G_PS.Rows[i][9].ToString() != "")
                    {
                        HSDPA_Drop_Rate = Convert.ToDouble(Table_3G_PS.Rows[i][9]);
                    }
                    if (Table_3G_PS.Rows[i][10].ToString() != "")
                    {
                        HSUPA_Drop_Rate = Convert.ToDouble(Table_3G_PS.Rows[i][10]);
                    }
                    if (Table_3G_PS.Rows[i][11].ToString() != "")
                    {
                        MultiRAB_SR = Convert.ToDouble(Table_3G_PS.Rows[i][11]);
                    }
                    if (Table_3G_PS.Rows[i][12].ToString() != "")
                    {
                        PS_RRC_SR = Convert.ToDouble(Table_3G_PS.Rows[i][12]);
                    }
                    if (Table_3G_PS.Rows[i][13].ToString() != "")
                    {
                        Ps_RAB_Establish = Convert.ToDouble(Table_3G_PS.Rows[i][13]);
                    }
                    if (Table_3G_PS.Rows[i][14].ToString() != "")
                    {
                        PS_MultiRAB_Establish = Convert.ToDouble(Table_3G_PS.Rows[i][14]);
                    }
                    if (Table_3G_PS.Rows[i][15].ToString() != "")
                    {
                        PS_Drop_Rate = Convert.ToDouble(Table_3G_PS.Rows[i][15]);
                    }
                    if (Table_3G_PS.Rows[i][16].ToString() != "")
                    {
                        HSDPA_Cell_Change_SR = Convert.ToDouble(Table_3G_PS.Rows[i][16]);
                    }
                    if (Table_3G_PS.Rows[i][17].ToString() != "")
                    {
                        HS_Share_Payload = Convert.ToDouble(Table_3G_PS.Rows[i][17]);
                    }
                    if (Table_3G_PS.Rows[i][18].ToString() != "")
                    {
                        DL_Cell_THR = Convert.ToDouble(Table_3G_PS.Rows[i][18]);
                    }
                    if (Table_3G_PS.Rows[i][4].ToString() != "")
                    {
                        Availability = Convert.ToDouble(Table_3G_PS.Rows[i][4]);
                    }

                    // Fill WPc Table
                    int TH_Index = 0;


                    var q1 = (from p in BASE_Table.AsEnumerable()
                              where p.Field<string>("Cell") == cell
                              select p).ToList();
                    var q2 = q1;

                    //string RRC_Connection_SR_BL = "";
                    //string ERAB_SR_Initial_BL = "";
                    //string ERAB_SR_Added_BL = "";
                    //string DL_THR_BL = "";
                    //string UL_THR_BL = "";
                    //string HO_SR_BL = "";
                    //string ERAB_Drop_Rate_BL = "";
                    //string S1_Signalling_SR_BL = "";
                    //string Inter_Freq_SR_BL = "";
                    //string Intra_Freq_SR_BL = "";
                    //if (q1.Count != 0)
                    //{

                    //}




                    if (Vendor == "Ericsson" && Coverage_type == "City")
                    {
                        q1 = (from p in Ericsson_3G_PS_Table_7.AsEnumerable()
                              where p.Field<string>("ElementID1") == cell
                              select p).ToList();

                        TH_Index = 0;
                    }
                    if (Vendor == "Ericsson" && Coverage_type == "None_City")
                    {
                        q1 = (from p in Ericsson_3G_PS_Table_7.AsEnumerable()
                              where p.Field<string>("ElementID1") == cell
                              select p).ToList();

                        TH_Index = 1;
                    }
                    if (Vendor == "Huawei" && Coverage_type == "City")
                    {
                        q1 = (from p in Huawei_3G_PS_Table_7.AsEnumerable()
                              where p.Field<string>("ElementID1") == cell
                              select p).ToList();

                        TH_Index = 2;
                    }
                    if (Vendor == "Huawei" && Coverage_type == "None_City")
                    {
                        q1 = (from p in Huawei_3G_PS_Table_7.AsEnumerable()
                              where p.Field<string>("ElementID1") == cell
                              select p).ToList();

                        TH_Index = 3;
                    }
                    if (Vendor == "Nokia" && Coverage_type == "City")
                    {
                        q1 = (from p in Nokia_3G_PS_Table_7.AsEnumerable()
                              where p.Field<string>("ElementID1") == cell
                              select p).ToList();

                        TH_Index = 4;
                    }
                    if (Vendor == "Nokia" && Coverage_type == "None_City")
                    {
                        q1 = (from p in Nokia_3G_PS_Table_7.AsEnumerable()
                              where p.Field<string>("ElementID1") == cell
                              select p).ToList();

                        TH_Index = 5;
                    }


                    int count = 0;
                    int update = 0;
                    if (HSDPA_SR > -1000 && HSDPA_SR < TH_3G_PS[TH_Index, 0])
                    {
                        count++;
                        int KPI_Index = 5;
                        if (q1.Count == 7)
                        {
                            Table_3G_PS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, HSDPA_SR, "HSDPA_SR", "-", q1[6].ItemArray[KPI_Index], q1[5].ItemArray[KPI_Index], q1[4].ItemArray[KPI_Index], q1[3].ItemArray[KPI_Index], q1[2].ItemArray[KPI_Index], q1[1].ItemArray[KPI_Index], q1[0].ItemArray[KPI_Index]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count <= 7 && q1.Count != 7 && q1.Count != 0)
                        {
                            double[] KPI_of_Days = new double[7];
                            for (int j = 1; j <= 7; j++)
                            {
                                if (q1.Count - j < 0)
                                {
                                    KPI_of_Days[j - 1] = -20000;
                                }
                                else
                                {
                                    DateTime d1 = Convert.ToDateTime(q1[q1.Count - j].ItemArray[0]);
                                    if (d1 == Date_of_WPC.AddDays(-j))
                                    {
                                        if (q1[q1.Count - j].ItemArray[KPI_Index].ToString() == "")
                                        {
                                            KPI_of_Days[j - 1] = -20000;
                                        }
                                        else
                                        {
                                            KPI_of_Days[j - 1] = Convert.ToDouble(q1[q1.Count - j].ItemArray[KPI_Index]);
                                        }

                                    }
                                    else
                                    {
                                        KPI_of_Days[j - 1] = -20000;
                                    }
                                }
                            }
                            Table_3G_PS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, HSDPA_SR, "HSDPA_SR", "-", KPI_of_Days[0], KPI_of_Days[1], KPI_of_Days[2], KPI_of_Days[3], KPI_of_Days[4], KPI_of_Days[5], KPI_of_Days[6]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count == 0)
                        {
                            Table_3G_PS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, HSDPA_SR, "HSDPA_SR");
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                    }
                    if (HSUPA_SR > -1000 && HSUPA_SR < TH_3G_PS[TH_Index, 1])
                    {
                        count++;
                        int KPI_Index = 6;
                        if (q1.Count == 7)
                        {
                            Table_3G_PS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, HSUPA_SR, "HSUPA_SR", "-", q1[6].ItemArray[KPI_Index], q1[5].ItemArray[KPI_Index], q1[4].ItemArray[KPI_Index], q1[3].ItemArray[KPI_Index], q1[2].ItemArray[KPI_Index], q1[1].ItemArray[KPI_Index], q1[0].ItemArray[KPI_Index]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count < 7 && q1.Count != 7 && q1.Count != 0)
                        {
                            double[] KPI_of_Days = new double[7];
                            for (int j = 1; j <= 7; j++)
                            {
                                if (q1.Count - j < 0)
                                {
                                    KPI_of_Days[j - 1] = -20000;
                                }
                                else
                                {
                                    DateTime d1 = Convert.ToDateTime(q1[q1.Count - j].ItemArray[0]);
                                    if (d1 == Date_of_WPC.AddDays(-j))
                                    {
                                        if (q1[q1.Count - j].ItemArray[KPI_Index].ToString() == "")
                                        {
                                            KPI_of_Days[j - 1] = -20000;
                                        }
                                        else
                                        {
                                            KPI_of_Days[j - 1] = Convert.ToDouble(q1[q1.Count - j].ItemArray[KPI_Index]);
                                        }

                                    }
                                    else
                                    {
                                        KPI_of_Days[j - 1] = -20000;
                                    }
                                }
                            }
                            Table_3G_PS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, HSUPA_SR, "HSUPA_SR", "-", KPI_of_Days[0], KPI_of_Days[1], KPI_of_Days[2], KPI_of_Days[3], KPI_of_Days[4], KPI_of_Days[5], KPI_of_Days[6]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count == 0)
                        {
                            Table_3G_PS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, HSUPA_SR, "HSUPA_SR");
                            Row_Count_of_WPC_Table++; update = 1;
                        }

                    }
                    if (UL_User_THR > -1000 && UL_User_THR < TH_3G_PS[TH_Index, 2])
                    {
                        count++;
                        int KPI_Index = 7;
                        if (q1.Count == 7)
                        {
                            Table_3G_PS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, UL_User_THR, "UL_User_THR", "-", q1[6].ItemArray[KPI_Index], q1[5].ItemArray[KPI_Index], q1[4].ItemArray[KPI_Index], q1[3].ItemArray[KPI_Index], q1[2].ItemArray[KPI_Index], q1[1].ItemArray[KPI_Index], q1[0].ItemArray[KPI_Index]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count < 7 && q1.Count != 7 && q1.Count != 0)
                        {
                            double[] KPI_of_Days = new double[7];
                            for (int j = 1; j <= 7; j++)
                            {
                                if (q1.Count - j < 0)
                                {
                                    KPI_of_Days[j - 1] = -20000;
                                }
                                else
                                {
                                    DateTime d1 = Convert.ToDateTime(q1[q1.Count - j].ItemArray[0]);
                                    if (d1 == Date_of_WPC.AddDays(-j))
                                    {
                                        if (q1[q1.Count - j].ItemArray[KPI_Index].ToString() == "")
                                        {
                                            KPI_of_Days[j - 1] = -20000;
                                        }
                                        else
                                        {
                                            KPI_of_Days[j - 1] = Convert.ToDouble(q1[q1.Count - j].ItemArray[KPI_Index]);
                                        }

                                    }
                                    else
                                    {
                                        KPI_of_Days[j - 1] = -20000;
                                    }
                                }
                            }
                            Table_3G_PS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, UL_User_THR, "UL_User_THR", "-", KPI_of_Days[0], KPI_of_Days[1], KPI_of_Days[2], KPI_of_Days[3], KPI_of_Days[4], KPI_of_Days[5], KPI_of_Days[6]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count == 0)
                        {
                            Table_3G_PS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, UL_User_THR, "UL_User_THR");
                            Row_Count_of_WPC_Table++; update = 1;
                        }

                    }
                    if (DL_User_THR > -1000 && DL_User_THR < TH_3G_PS[TH_Index, 3])
                    {
                        count++;
                        int KPI_Index = 8;
                        if (q1.Count == 7)
                        {
                            Table_3G_PS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, DL_User_THR, "DL_User_THR", "-", q1[6].ItemArray[KPI_Index], q1[5].ItemArray[KPI_Index], q1[4].ItemArray[KPI_Index], q1[3].ItemArray[KPI_Index], q1[2].ItemArray[KPI_Index], q1[1].ItemArray[KPI_Index], q1[0].ItemArray[KPI_Index]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count < 7 && q1.Count != 7 && q1.Count != 0)
                        {
                            double[] KPI_of_Days = new double[7];
                            for (int j = 1; j <= 7; j++)
                            {
                                if (q1.Count - j < 0)
                                {
                                    KPI_of_Days[j - 1] = -20000;
                                }
                                else
                                {
                                    DateTime d1 = Convert.ToDateTime(q1[q1.Count - j].ItemArray[0]);
                                    if (d1 == Date_of_WPC.AddDays(-j))
                                    {
                                        if (q1[q1.Count - j].ItemArray[KPI_Index].ToString() == "")
                                        {
                                            KPI_of_Days[j - 1] = -20000;
                                        }
                                        else
                                        {
                                            KPI_of_Days[j - 1] = Convert.ToDouble(q1[q1.Count - j].ItemArray[KPI_Index]);
                                        }

                                    }
                                    else
                                    {
                                        KPI_of_Days[j - 1] = -20000;
                                    }
                                }
                            }
                            Table_3G_PS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, DL_User_THR, "DL_User_THR", "-", KPI_of_Days[0], KPI_of_Days[1], KPI_of_Days[2], KPI_of_Days[3], KPI_of_Days[4], KPI_of_Days[5], KPI_of_Days[6]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count == 0)
                        {
                            Table_3G_PS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, DL_User_THR, "DL_User_THR");
                            Row_Count_of_WPC_Table++; update = 1;
                        }

                    }
                    if (HSDPA_Drop_Rate > -1000 && HSDPA_Drop_Rate > TH_3G_PS[TH_Index, 4])
                    {
                        count++;
                        int KPI_Index = 9;
                        if (q1.Count == 7)
                        {
                            Table_3G_PS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, HSDPA_Drop_Rate, "HSDPA_Drop_Rate", "-", q1[6].ItemArray[KPI_Index], q1[5].ItemArray[KPI_Index], q1[4].ItemArray[KPI_Index], q1[3].ItemArray[KPI_Index], q1[2].ItemArray[KPI_Index], q1[1].ItemArray[KPI_Index], q1[0].ItemArray[KPI_Index]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count < 7 && q1.Count != 7 && q1.Count != 0)
                        {
                            double[] KPI_of_Days = new double[7];
                            for (int j = 1; j <= 7; j++)
                            {
                                if (q1.Count - j < 0)
                                {
                                    KPI_of_Days[j - 1] = -20000;
                                }
                                else
                                {
                                    DateTime d1 = Convert.ToDateTime(q1[q1.Count - j].ItemArray[0]);
                                    if (d1 == Date_of_WPC.AddDays(-j))
                                    {
                                        if (q1[q1.Count - j].ItemArray[KPI_Index].ToString() == "")
                                        {
                                            KPI_of_Days[j - 1] = -20000;
                                        }
                                        else
                                        {
                                            KPI_of_Days[j - 1] = Convert.ToDouble(q1[q1.Count - j].ItemArray[KPI_Index]);
                                        }

                                    }
                                    else
                                    {
                                        KPI_of_Days[j - 1] = -20000;
                                    }
                                }
                            }
                            Table_3G_PS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, HSDPA_Drop_Rate, "HSDPA_Drop_Rate", "-", KPI_of_Days[0], KPI_of_Days[1], KPI_of_Days[2], KPI_of_Days[3], KPI_of_Days[4], KPI_of_Days[5], KPI_of_Days[6]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count == 0)
                        {
                            Table_3G_PS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, HSDPA_Drop_Rate, "HSDPA_Drop_Rate");
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                    }

                    if (HSUPA_Drop_Rate > -1000 && HSUPA_Drop_Rate > TH_3G_PS[TH_Index, 5])
                    {
                        count++;
                        int KPI_Index = 10;
                        if (q1.Count == 7)
                        {
                            Table_3G_PS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, HSUPA_Drop_Rate, "HSUPA_Drop_Rate", "-", q1[6].ItemArray[KPI_Index], q1[5].ItemArray[KPI_Index], q1[4].ItemArray[KPI_Index], q1[3].ItemArray[KPI_Index], q1[2].ItemArray[KPI_Index], q1[1].ItemArray[KPI_Index], q1[0].ItemArray[KPI_Index]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count < 7 && q1.Count != 7 && q1.Count != 0)
                        {
                            double[] KPI_of_Days = new double[7];
                            for (int j = 1; j <= 7; j++)
                            {
                                if (q1.Count - j < 0)
                                {
                                    KPI_of_Days[j - 1] = -20000;
                                }
                                else
                                {
                                    DateTime d1 = Convert.ToDateTime(q1[q1.Count - j].ItemArray[0]);
                                    if (d1 == Date_of_WPC.AddDays(-j))
                                    {
                                        if (q1[q1.Count - j].ItemArray[KPI_Index].ToString() == "")
                                        {
                                            KPI_of_Days[j - 1] = -20000;
                                        }
                                        else
                                        {
                                            KPI_of_Days[j - 1] = Convert.ToDouble(q1[q1.Count - j].ItemArray[KPI_Index]);
                                        }

                                    }
                                    else
                                    {
                                        KPI_of_Days[j - 1] = -20000;
                                    }
                                }
                            }
                            Table_3G_PS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, HSUPA_Drop_Rate, "HSUPA_Drop_Rate", "-", KPI_of_Days[0], KPI_of_Days[1], KPI_of_Days[2], KPI_of_Days[3], KPI_of_Days[4], KPI_of_Days[5], KPI_of_Days[6]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count == 0)
                        {
                            Table_3G_PS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, HSUPA_Drop_Rate, "HSUPA_Drop_Rate");
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                    }

                    if (MultiRAB_SR > -1000 && MultiRAB_SR < TH_3G_PS[TH_Index, 6])
                    {
                        count++;
                        int KPI_Index = 11;
                        if (q1.Count == 7)
                        {
                            Table_3G_PS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, MultiRAB_SR, "MultiRAB_SR", "-", q1[6].ItemArray[KPI_Index], q1[5].ItemArray[KPI_Index], q1[4].ItemArray[KPI_Index], q1[3].ItemArray[KPI_Index], q1[2].ItemArray[KPI_Index], q1[1].ItemArray[KPI_Index], q1[0].ItemArray[KPI_Index]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count < 7 && q1.Count != 7 && q1.Count != 0)
                        {
                            double[] KPI_of_Days = new double[7];
                            for (int j = 1; j <= 7; j++)
                            {
                                if (q1.Count - j < 0)
                                {
                                    KPI_of_Days[j - 1] = -20000;
                                }
                                else
                                {
                                    DateTime d1 = Convert.ToDateTime(q1[q1.Count - j].ItemArray[0]);
                                    if (d1 == Date_of_WPC.AddDays(-j))
                                    {
                                        if (q1[q1.Count - j].ItemArray[KPI_Index].ToString() == "")
                                        {
                                            KPI_of_Days[j - 1] = -20000;
                                        }
                                        else
                                        {
                                            KPI_of_Days[j - 1] = Convert.ToDouble(q1[q1.Count - j].ItemArray[KPI_Index]);
                                        }

                                    }
                                    else
                                    {
                                        KPI_of_Days[j - 1] = -20000;
                                    }
                                }
                            }
                            Table_3G_PS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, MultiRAB_SR, "MultiRAB_SR", "-", KPI_of_Days[0], KPI_of_Days[1], KPI_of_Days[2], KPI_of_Days[3], KPI_of_Days[4], KPI_of_Days[5], KPI_of_Days[6]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count == 0)
                        {
                            Table_3G_PS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, MultiRAB_SR, "MultiRAB_SR");
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                    }

                    if (PS_RRC_SR > -1000 && PS_RRC_SR < TH_3G_PS[TH_Index, 7])
                    {
                        count++;
                        int KPI_Index = 12;
                        if (q1.Count == 7)
                        {
                            Table_3G_PS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, PS_RRC_SR, "PS_RRC_SR", "-", q1[6].ItemArray[KPI_Index], q1[5].ItemArray[KPI_Index], q1[4].ItemArray[KPI_Index], q1[3].ItemArray[KPI_Index], q1[2].ItemArray[KPI_Index], q1[1].ItemArray[KPI_Index], q1[0].ItemArray[KPI_Index]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count < 7 && q1.Count != 7 && q1.Count != 0)
                        {
                            double[] KPI_of_Days = new double[7];
                            for (int j = 1; j <= 7; j++)
                            {
                                if (q1.Count - j < 0)
                                {
                                    KPI_of_Days[j - 1] = -20000;
                                }
                                else
                                {
                                    DateTime d1 = Convert.ToDateTime(q1[q1.Count - j].ItemArray[0]);
                                    if (d1 == Date_of_WPC.AddDays(-j))
                                    {
                                        if (q1[q1.Count - j].ItemArray[KPI_Index].ToString() == "")
                                        {
                                            KPI_of_Days[j - 1] = -20000;
                                        }
                                        else
                                        {
                                            KPI_of_Days[j - 1] = Convert.ToDouble(q1[q1.Count - j].ItemArray[KPI_Index]);
                                        }

                                    }
                                    else
                                    {
                                        KPI_of_Days[j - 1] = -20000;
                                    }
                                }
                            }
                            Table_3G_PS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, PS_RRC_SR, "PS_RRC_SR", "-", KPI_of_Days[0], KPI_of_Days[1], KPI_of_Days[2], KPI_of_Days[3], KPI_of_Days[4], KPI_of_Days[5], KPI_of_Days[6]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count == 0)
                        {
                            Table_3G_PS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, PS_RRC_SR, "PS_RRC_SR");
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                    }

                    if (Ps_RAB_Establish > -1000 && Ps_RAB_Establish < TH_3G_PS[TH_Index, 8])
                    {
                        count++;
                        int KPI_Index = 13;
                        if (q1.Count == 7)
                        {
                            Table_3G_PS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, Ps_RAB_Establish, "Ps_RAB_Establish", "-", q1[6].ItemArray[KPI_Index], q1[5].ItemArray[KPI_Index], q1[4].ItemArray[KPI_Index], q1[3].ItemArray[KPI_Index], q1[2].ItemArray[KPI_Index], q1[1].ItemArray[KPI_Index], q1[0].ItemArray[KPI_Index]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count < 7 && q1.Count != 7 && q1.Count != 0)
                        {
                            double[] KPI_of_Days = new double[7];
                            for (int j = 1; j <= 7; j++)
                            {
                                if (q1.Count - j < 0)
                                {
                                    KPI_of_Days[j - 1] = -20000;
                                }
                                else
                                {
                                    DateTime d1 = Convert.ToDateTime(q1[q1.Count - j].ItemArray[0]);
                                    if (d1 == Date_of_WPC.AddDays(-j))
                                    {
                                        if (q1[q1.Count - j].ItemArray[KPI_Index].ToString() == "")
                                        {
                                            KPI_of_Days[j - 1] = -20000;
                                        }
                                        else
                                        {
                                            KPI_of_Days[j - 1] = Convert.ToDouble(q1[q1.Count - j].ItemArray[KPI_Index]);
                                        }

                                    }
                                    else
                                    {
                                        KPI_of_Days[j - 1] = -20000;
                                    }
                                }
                            }
                            Table_3G_PS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, Ps_RAB_Establish, "PS_RAB_Establish", "-", KPI_of_Days[0], KPI_of_Days[1], KPI_of_Days[2], KPI_of_Days[3], KPI_of_Days[4], KPI_of_Days[5], KPI_of_Days[6]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count == 0)
                        {
                            Table_3G_PS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, Ps_RAB_Establish, "Ps_RAB_Establish");
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                    }

                    if (PS_MultiRAB_Establish > -1000 && PS_MultiRAB_Establish < TH_3G_PS[TH_Index, 9])
                    {
                        count++;
                        int KPI_Index = 14;
                        if (q1.Count == 7)
                        {
                            Table_3G_PS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, PS_MultiRAB_Establish, "PS_MultiRAB_Establish", "-", q1[6].ItemArray[KPI_Index], q1[5].ItemArray[KPI_Index], q1[4].ItemArray[KPI_Index], q1[3].ItemArray[KPI_Index], q1[2].ItemArray[KPI_Index], q1[1].ItemArray[KPI_Index], q1[0].ItemArray[KPI_Index]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count < 7 && q1.Count != 7 && q1.Count != 0)
                        {
                            double[] KPI_of_Days = new double[7];
                            for (int j = 1; j <= 7; j++)
                            {
                                if (q1.Count - j < 0)
                                {
                                    KPI_of_Days[j - 1] = -20000;
                                }
                                else
                                {
                                    DateTime d1 = Convert.ToDateTime(q1[q1.Count - j].ItemArray[0]);
                                    if (d1 == Date_of_WPC.AddDays(-j))
                                    {
                                        if (q1[q1.Count - j].ItemArray[KPI_Index].ToString() == "")
                                        {
                                            KPI_of_Days[j - 1] = -20000;
                                        }
                                        else
                                        {
                                            KPI_of_Days[j - 1] = Convert.ToDouble(q1[q1.Count - j].ItemArray[KPI_Index]);
                                        }

                                    }
                                    else
                                    {
                                        KPI_of_Days[j - 1] = -20000;
                                    }
                                }
                            }
                            Table_3G_PS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, PS_MultiRAB_Establish, "PS_MultiRAB_Establish", "-", KPI_of_Days[0], KPI_of_Days[1], KPI_of_Days[2], KPI_of_Days[3], KPI_of_Days[4], KPI_of_Days[5], KPI_of_Days[6]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count == 0)
                        {
                            Table_3G_PS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, PS_MultiRAB_Establish, "PS_MultiRAB_Establish");
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                    }

                    if (PS_Drop_Rate > -1000 && PS_Drop_Rate > TH_3G_PS[TH_Index, 10])
                    {
                        count++;
                        int KPI_Index = 15;
                        if (q1.Count == 7)
                        {
                            Table_3G_PS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, PS_Drop_Rate, "PS_Drop_Rate", "-", q1[6].ItemArray[KPI_Index], q1[5].ItemArray[KPI_Index], q1[4].ItemArray[KPI_Index], q1[3].ItemArray[KPI_Index], q1[2].ItemArray[KPI_Index], q1[1].ItemArray[KPI_Index], q1[0].ItemArray[KPI_Index]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count < 7 && q1.Count != 7 && q1.Count != 0)
                        {
                            double[] KPI_of_Days = new double[7];
                            for (int j = 1; j <= 7; j++)
                            {
                                if (q1.Count - j < 0)
                                {
                                    KPI_of_Days[j - 1] = -20000;
                                }
                                else
                                {
                                    DateTime d1 = Convert.ToDateTime(q1[q1.Count - j].ItemArray[0]);
                                    if (d1 == Date_of_WPC.AddDays(-j))
                                    {
                                        if (q1[q1.Count - j].ItemArray[KPI_Index].ToString() == "")
                                        {
                                            KPI_of_Days[j - 1] = -20000;
                                        }
                                        else
                                        {
                                            KPI_of_Days[j - 1] = Convert.ToDouble(q1[q1.Count - j].ItemArray[KPI_Index]);
                                        }

                                    }
                                    else
                                    {
                                        KPI_of_Days[j - 1] = -20000;
                                    }
                                }
                            }
                            Table_3G_PS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, PS_Drop_Rate, "PS_Drop_Rate", "-", KPI_of_Days[0], KPI_of_Days[1], KPI_of_Days[2], KPI_of_Days[3], KPI_of_Days[4], KPI_of_Days[5], KPI_of_Days[6]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count == 0)
                        {
                            Table_3G_PS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, PS_Drop_Rate, "PS_Drop_Rate");
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                    }

                    if (HSDPA_Cell_Change_SR > -1000 && HSDPA_Cell_Change_SR < TH_3G_PS[TH_Index, 11])
                    {
                        count++;
                        int KPI_Index = 16;
                        if (q1.Count == 7)
                        {
                            Table_3G_PS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, HSDPA_Cell_Change_SR, "HSDPA_Cell_Change_SR", "-", q1[6].ItemArray[KPI_Index], q1[5].ItemArray[KPI_Index], q1[4].ItemArray[KPI_Index], q1[3].ItemArray[KPI_Index], q1[2].ItemArray[KPI_Index], q1[1].ItemArray[KPI_Index], q1[0].ItemArray[KPI_Index]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count < 7 && q1.Count != 7 && q1.Count != 0)
                        {
                            double[] KPI_of_Days = new double[7];
                            for (int j = 1; j <= 7; j++)
                            {
                                if (q1.Count - j < 0)
                                {
                                    KPI_of_Days[j - 1] = -20000;
                                }
                                else
                                {
                                    DateTime d1 = Convert.ToDateTime(q1[q1.Count - j].ItemArray[0]);
                                    if (d1 == Date_of_WPC.AddDays(-j))
                                    {
                                        if (q1[q1.Count - j].ItemArray[KPI_Index].ToString() == "")
                                        {
                                            KPI_of_Days[j - 1] = -20000;
                                        }
                                        else
                                        {
                                            KPI_of_Days[j - 1] = Convert.ToDouble(q1[q1.Count - j].ItemArray[KPI_Index]);
                                        }

                                    }
                                    else
                                    {
                                        KPI_of_Days[j - 1] = -20000;
                                    }
                                }
                            }
                            Table_3G_PS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, HSDPA_Cell_Change_SR, "HSDPA_Cell_Change_SR", "-", KPI_of_Days[0], KPI_of_Days[1], KPI_of_Days[2], KPI_of_Days[3], KPI_of_Days[4], KPI_of_Days[5], KPI_of_Days[6]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count == 0)
                        {
                            Table_3G_PS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, HSDPA_Cell_Change_SR, "HSDPA_Cell_Change_SR");
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                    }

                    if (HS_Share_Payload > -1000 && HS_Share_Payload < TH_3G_PS[TH_Index, 12])
                    {
                        count++;
                        int KPI_Index = 17;
                        if (q1.Count == 7)
                        {
                            Table_3G_PS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, HS_Share_Payload, "HS_Share_Payload", "-", q1[6].ItemArray[KPI_Index], q1[5].ItemArray[KPI_Index], q1[4].ItemArray[KPI_Index], q1[3].ItemArray[KPI_Index], q1[2].ItemArray[KPI_Index], q1[1].ItemArray[KPI_Index], q1[0].ItemArray[KPI_Index]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count < 7 && q1.Count != 7 && q1.Count != 0)
                        {
                            double[] KPI_of_Days = new double[7];
                            for (int j = 1; j <= 7; j++)
                            {
                                if (q1.Count - j < 0)
                                {
                                    KPI_of_Days[j - 1] = -20000;
                                }
                                else
                                {
                                    DateTime d1 = Convert.ToDateTime(q1[q1.Count - j].ItemArray[0]);
                                    if (d1 == Date_of_WPC.AddDays(-j))
                                    {
                                        if (q1[q1.Count - j].ItemArray[KPI_Index].ToString() == "")
                                        {
                                            KPI_of_Days[j - 1] = -20000;
                                        }
                                        else
                                        {
                                            KPI_of_Days[j - 1] = Convert.ToDouble(q1[q1.Count - j].ItemArray[KPI_Index]);
                                        }

                                    }
                                    else
                                    {
                                        KPI_of_Days[j - 1] = -20000;
                                    }
                                }
                            }
                            Table_3G_PS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, HS_Share_Payload, "HS_Share_Payload", "-", KPI_of_Days[0], KPI_of_Days[1], KPI_of_Days[2], KPI_of_Days[3], KPI_of_Days[4], KPI_of_Days[5], KPI_of_Days[6]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count == 0)
                        {
                            Table_3G_PS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, HS_Share_Payload, "HS_Share_Payload");
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                    }

                    if (DL_Cell_THR > -1000 && DL_Cell_THR < TH_3G_PS[TH_Index, 13])
                    {
                        count++;
                        int KPI_Index = 18;
                        if (q1.Count == 7)
                        {
                            Table_3G_PS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, DL_Cell_THR, "DL_Cell_THR", "-", q1[6].ItemArray[KPI_Index], q1[5].ItemArray[KPI_Index], q1[4].ItemArray[KPI_Index], q1[3].ItemArray[KPI_Index], q1[2].ItemArray[KPI_Index], q1[1].ItemArray[KPI_Index], q1[0].ItemArray[KPI_Index]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count < 7 && q1.Count != 7 && q1.Count != 0)
                        {
                            double[] KPI_of_Days = new double[7];
                            for (int j = 1; j <= 7; j++)
                            {
                                if (q1.Count - j < 0)
                                {
                                    KPI_of_Days[j - 1] = -20000;
                                }
                                else
                                {
                                    DateTime d1 = Convert.ToDateTime(q1[q1.Count - j].ItemArray[0]);
                                    if (d1 == Date_of_WPC.AddDays(-j))
                                    {
                                        if (q1[q1.Count - j].ItemArray[KPI_Index].ToString() == "")
                                        {
                                            KPI_of_Days[j - 1] = -20000;
                                        }
                                        else
                                        {
                                            KPI_of_Days[j - 1] = Convert.ToDouble(q1[q1.Count - j].ItemArray[KPI_Index]);
                                        }

                                    }
                                    else
                                    {
                                        KPI_of_Days[j - 1] = -20000;
                                    }
                                }
                            }
                            Table_3G_PS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, DL_Cell_THR, "DL_Cell_THR", "-", KPI_of_Days[0], KPI_of_Days[1], KPI_of_Days[2], KPI_of_Days[3], KPI_of_Days[4], KPI_of_Days[5], KPI_of_Days[6]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count == 0)
                        {
                            Table_3G_PS_WPC.Rows.Add(Contractor, Province, Vendor, RNC, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, DL_Cell_THR, "DL_Cell_THR");
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                    }

                    if (update == 1)
                    {
                        Table_3G_PS_WPC.Rows[Row_Count_of_WPC_Table - 1][12] = Convert.ToString(count);
                    }

                    //label5.Text = Convert.ToString(Math.Round(Convert.ToDouble(i / Table_3G_PS.Rows.Count)));
                    progressBar1.Value = i;

                }


                XLWorkbook wb = new XLWorkbook();
                wb.Worksheets.Add(Table_3G_PS_WPC, "WPC_3G_PS");
                var saveFileDialog = new SaveFileDialog
                {
                    FileName = "WPC_3G_PS",
                    Filter = "Excel files|*.xlsx",
                    Title = "Save an Excel File"
                };

                saveFileDialog.ShowDialog();

                if (!String.IsNullOrWhiteSpace(saveFileDialog.FileName))
                    wb.SaveAs(saveFileDialog.FileName);




                label4.Text = "Finished";
                label4.BackColor = Color.Yellow;
                MessageBox.Show("FINISHED");



            }




            // Quary of Select Data in 4G
            if (checkBox3.Checked == true)
            {


                string Ericsson_4G_TBL = @"select[Datetime] as 'Date', [eNodeB], [Total_Volume(UL+DL)(GB)(eNodeB_Eric)] as 'Paylaod (GB)', [Cell_Availability_Rate_Include_Blocking(Cell_EricLTE)]  as 'Availability', [RRC_Estab_Success_Rate(ReAtt)(EUCell_Eric)] as'RRC Establishmnet SR', [Initial_ERAB_Estab_Success_Rate(eNodeB_Eric)] as 'Initial ERAB Establishmnet Success Rate', [E-RAB_Setup_SR_incl_added_New(EUCell_Eric)] as 'E-RAB Stetp SR', [Average_UE_DL_Throughput(Mbps)(eNodeB_Eric)] as 'UE DL Throughput (Mbps)', 
[Average_UE_UL_Throughput(Mbps)(eNodeB_Eric)] as 'UE UL Throughput (Mbps)', [Handover_Execution_Rate(EUCell_Eric)] as 'Handover Execution Rate',  [E_RAB_Drop_Rate(eNodeB_Eric)] as 'ERAB Drop Rate', 
[S1Signal_Estab_Success_Rate(EUCell_Eric)] as 'S1 Signal Establishment Success Rate', [InterF_Handover_Execution(eNodeB_Eric)] as 'InterF Handover Execution Rate', [IntraF_Handover_Execution(eNodeB_Eric)] as 'IntraF Handover Execution Rate', [Average_UE_Ul_Packet_Loss_Rate(eNodeB_Eric)] as 'UE UL Packet Loss Rate'
from[dbo].[TBL_LTE_CELL_Daily_E] where
(substring([eNodeB],1,2)='KJ' or substring([eNodeB],1,2)='CH' or substring([eNodeB],1,2)='AS' or substring([eNodeB],1,2)='GL' or
                                                                                 substring([eNodeB],1,2)='GN'  or substring([eNodeB],1,2)='KM' or substring([eNodeB],1,2)='KH' or substring([eNodeB],1,2)='KZ' or substring([eNodeB],1,2)='MA'
																				  or substring([eNodeB],1,2)='SM' or substring([eNodeB],1,2)='TH' or substring([eNodeB],1,2)='AG' or substring([eNodeB],1,2)='YZ')
and ([RRC_Estab_Success_Rate(ReAtt)(EUCell_Eric)]<99.758 or [Initial_ERAB_Estab_Success_Rate(eNodeB_Eric)]<99.192 or [E-RAB_Setup_SR_incl_added_New(EUCell_Eric)]<99.525 or [Average_UE_DL_Throughput(Mbps)(eNodeB_Eric)]<1.879 or
[Average_UE_UL_Throughput(Mbps)(eNodeB_Eric)]<0.248 or [Handover_Execution_Rate(EUCell_Eric)]<94.58 or [E_RAB_Drop_Rate(eNodeB_Eric)]>0.443 or
[S1Signal_Estab_Success_Rate(EUCell_Eric)]<99.831 or [InterF_Handover_Execution(eNodeB_Eric)]<69.456 or [IntraF_Handover_Execution(eNodeB_Eric)]<94.739 or  [Average_UE_Ul_Packet_Loss_Rate(eNodeB_Eric)]>0.878) and Datetime  ='" + Date_of_WPC + "'";


                SqlCommand Ericsson_4G_TBL_Quary = new SqlCommand(Ericsson_4G_TBL, connection);
                Ericsson_4G_TBL_Quary.ExecuteNonQuery();
                DataTable Ericsson_4G_Table = new DataTable();
                SqlDataAdapter dataAdapter_Ericsson_4G = new SqlDataAdapter(Ericsson_4G_TBL_Quary);
                dataAdapter_Ericsson_4G.Fill(Ericsson_4G_Table);




                string Huawei_4G_TBL = @"select [Datetime] as 'Date', [eNodeB], [Total_Traffic_Volume(GB)] as 'Paylaod (GB)', [Cell_Availability_Rate_include_Blocking(Cell_Hu)]  as 'Availability', [RRC_Connection_Setup_Success_Rate_service] as'RRC Establishmnet SR', [E-RAB_Setup_Success_Rate] as 'Initial ERAB Establishmnet Success Rate', [E-RAB_Setup_Success_Rate(Hu_Cell)] as 'E-RAB Stetp SR', [Average_Downlink_User_Throughput(Mbit/s)] as 'UE DL Throughput (Mbps)', 
[Average_UPlink_User_Throughput(Mbit/s)] as 'UE UL Throughput (Mbps)', [Intra_RAT_Handover_SR_Intra+Inter_frequency(Huawei_LTE_Cell] as 'Handover Execution Rate',  [Call_Drop_Rate] as 'ERAB Drop Rate', 
[S1Signal_E-RAB_Setup_SR(Hu_Cell)] as 'S1 Signal Establishment Success Rate', [InterF_HOOut_SR] as 'InterF Handover Execution Rate', [IntraF_HOOut_SR] as 'IntraF Handover Execution Rate', [Average_UL_Packet_Loss_%(Huawei_LTE_UCell)] as 'UE UL Packet Loss Rate'
from [dbo].[TBL_LTE_CELL_Daily_H] where 
(substring([eNodeB],1,2)='KJ' or substring([eNodeB],1,2)='CH' or substring([eNodeB],1,2)='AS' or substring([eNodeB],1,2)='GL' or 
                                                                                  substring([eNodeB],1,2)='GN'  or substring([eNodeB],1,2)='KM' or substring([eNodeB],1,2)='KH' or substring([eNodeB],1,2)='KZ' or substring([eNodeB],1,2)='MA'
																				  or substring([eNodeB],1,2)='SM' or substring([eNodeB],1,2)='TH' or substring([eNodeB],1,2)='AG' or substring([eNodeB],1,2)='YZ')
and ([RRC_Connection_Setup_Success_Rate_service]<99.894 or [E-RAB_Setup_Success_Rate]<99.696 or [E-RAB_Setup_Success_Rate(Hu_Cell)]<99.697 or [Average_Downlink_User_Throughput(Mbit/s)]<4.057 or
[Average_UPlink_User_Throughput(Mbit/s)]<0.91 or [Intra_RAT_Handover_SR_Intra+Inter_frequency(Huawei_LTE_Cell]<97.807 or [Call_Drop_Rate]>0.491 or
[S1Signal_E-RAB_Setup_SR(Hu_Cell)]<99.926 or [InterF_HOOut_SR]<94.872 or [IntraF_HOOut_SR]<97.638 or [Average_UL_Packet_Loss_%(Huawei_LTE_UCell)]>0.011) and Datetime  ='" + Date_of_WPC + "'";




                SqlCommand Huawei_4G_TBL_Quary = new SqlCommand(Huawei_4G_TBL, connection);
                Huawei_4G_TBL_Quary.ExecuteNonQuery();
                DataTable Huawei_4G_Table = new DataTable();
                SqlDataAdapter dataAdapter_Huawei_4G = new SqlDataAdapter(Huawei_4G_TBL_Quary);
                dataAdapter_Huawei_4G.Fill(Huawei_4G_Table);




                string Nokia_4G_TBL = @"select [Date] as 'Date', [ElementID1], [Total_Payload_GB(Nokia_LTE_CELL)] as 'Paylaod (GB)', [cell_availability_include_manual_blocking(Nokia_LTE_CELL)]  as 'Availability', [RRC_Connection_Setup_Success_Ratio(Nokia_LTE_CELL)] as'RRC Establishmnet SR', [Initial_E-RAB_Setup_Success_Ratio(Nokia_LTE_CELL)] as 'Initial ERAB Establishmnet Success Rate', [E-RAB_Setup_SR_incl_added(Nokia_LTE_CELL)] as 'E-RAB Stetp SR', [User_Throughput_DL_mbps(Nokia_LTE_CELL)] as 'UE DL Throughput (Mbps)', 
[User_Throughput_UL_mbps(Nokia_LTE_CELL)] as 'UE UL Throughput (Mbps)', [Intra_RAT_Handover_SR_Intra+Inter_frequency(Nokia_LTE_CELL)] as 'Handover Execution Rate',  [E-RAB_Drop_Ratio_RAN_View(Nokia_LTE_CELL)] as 'ERAB Drop Rate', 
[S1Signal_E-RAB_Setup_SR(Nokia_LTE_CELL)] as 'S1 Signal Establishment Success Rate', [Inter-Freq_HO_SR(Nokia_LTE_CELL)] as 'InterF Handover Execution Rate', [HO_Success_Ratio_intra_eNB(Nokia_LTE_CELL)] as 'IntraF Handover Execution Rate', [Packet_loss_UL(Nokia_EUCELL)] as 'UE UL Packet Loss Rate'
from [dbo].[TBL_LTE_CELL_Daily_N] where 
(substring([ElementID1],1,2)='KJ' or substring([ElementID1],1,2)='CH' or substring([ElementID1],1,2)='AS' or substring([ElementID1],1,2)='GL' or 
                                                                                  substring([ElementID1],1,2)='GN'  or substring([ElementID1],1,2)='KM' or substring([ElementID1],1,2)='KH' or substring([ElementID1],1,2)='KZ' or substring([ElementID1],1,2)='MA'
																				  or substring([ElementID1],1,2)='SM' or substring([ElementID1],1,2)='TH' or substring([ElementID1],1,2)='AG' or substring([ElementID1],1,2)='YZ')
and ([RRC_Connection_Setup_Success_Ratio(Nokia_LTE_CELL)]<99.747 or [Initial_E-RAB_Setup_Success_Ratio(Nokia_LTE_CELL)]<99.282 or [E-RAB_Setup_SR_incl_added(Nokia_LTE_CELL)]<99.281 or [User_Throughput_DL_mbps(Nokia_LTE_CELL)]<4.465 or
[User_Throughput_UL_mbps(Nokia_LTE_CELL)]<0.476 or [Intra_RAT_Handover_SR_Intra+Inter_frequency(Nokia_LTE_CELL)]<91.672 or [E-RAB_Drop_Ratio_RAN_View(Nokia_LTE_CELL)]>0.854 or
[S1Signal_E-RAB_Setup_SR(Nokia_LTE_CELL)]<99.901 or [Inter-Freq_HO_SR(Nokia_LTE_CELL)]<40 or [HO_Success_Ratio_intra_eNB(Nokia_LTE_CELL)]<94.934 or [Packet_loss_UL(Nokia_EUCELL)]>0.591) and Date  ='" + Date_of_WPC + "'";


                SqlCommand Nokia_4G_TBL_Quary = new SqlCommand(Nokia_4G_TBL, connection);
                Nokia_4G_TBL_Quary.ExecuteNonQuery();
                DataTable Nokia_4G_Table = new DataTable();
                SqlDataAdapter dataAdapter_Nokia_4G = new SqlDataAdapter(Nokia_4G_TBL_Quary);
                dataAdapter_Nokia_4G.Fill(Nokia_4G_Table);






                // Table of Oldest 7 Dayes
                string Ericsson_4G_TBL_7 = @"select[Datetime] as 'Date', [eNodeB], [Total_Volume(UL+DL)(GB)(eNodeB_Eric)] as 'Paylaod (GB)', [Cell_Availability_Rate_Include_Blocking(Cell_EricLTE)]  as 'Availability', [RRC_Estab_Success_Rate(ReAtt)(EUCell_Eric)] as'RRC Establishmnet SR', [Initial_ERAB_Estab_Success_Rate(eNodeB_Eric)] as 'Initial ERAB Establishmnet Success Rate', [E-RAB_Setup_SR_incl_added_New(EUCell_Eric)] as 'E-RAB Stetp SR', [Average_UE_DL_Throughput(Mbps)(eNodeB_Eric)] as 'UE DL Throughput (Mbps)', 
[Average_UE_UL_Throughput(Mbps)(eNodeB_Eric)] as 'UE UL Throughput (Mbps)', [Handover_Execution_Rate(EUCell_Eric)] as 'Handover Execution Rate',  [E_RAB_Drop_Rate(eNodeB_Eric)] as 'ERAB Drop Rate', 
[S1Signal_Estab_Success_Rate(EUCell_Eric)] as 'S1 Signal Establishment Success Rate', [InterF_Handover_Execution(eNodeB_Eric)] as 'InterF Handover Execution Rate', [IntraF_Handover_Execution(eNodeB_Eric)] as 'IntraF Handover Execution Rate', [Average_UE_Ul_Packet_Loss_Rate(eNodeB_Eric)] as 'UE UL Packet Loss Rate'
from[dbo].[TBL_LTE_CELL_Daily_E] where
(substring([eNodeB],1,2)='KJ' or substring([eNodeB],1,2)='CH' or substring([eNodeB],1,2)='AS' or substring([eNodeB],1,2)='GL' or
                                                                                 substring([eNodeB],1,2)='GN'  or substring([eNodeB],1,2)='KM' or substring([eNodeB],1,2)='KH' or substring([eNodeB],1,2)='KZ' or substring([eNodeB],1,2)='MA'
																				  or substring([eNodeB],1,2)='SM' or substring([eNodeB],1,2)='TH' or substring([eNodeB],1,2)='AG' or substring([eNodeB],1,2)='YZ')
and Datetime  >='" + Date_of_WPC_7 + "' and Datetime  <'" + Date_of_WPC + "'";



                SqlCommand Ericsson_4G_TBL_Quary_7 = new SqlCommand(Ericsson_4G_TBL_7, connection);
                Ericsson_4G_TBL_Quary_7.CommandTimeout = 0;
                Ericsson_4G_TBL_Quary_7.ExecuteNonQuery();
                DataTable Ericsson_4G_Table_7 = new DataTable();
                SqlDataAdapter dataAdapter_Ericsson_4G_7 = new SqlDataAdapter(Ericsson_4G_TBL_Quary_7);
                dataAdapter_Ericsson_4G_7.Fill(Ericsson_4G_Table_7);


                string Huawei_4G_TBL_7 = @"select [Datetime] as 'Date', [eNodeB], [Total_Traffic_Volume(GB)] as 'Paylaod (GB)', [Cell_Availability_Rate_include_Blocking(Cell_Hu)]  as 'Availability', [RRC_Connection_Setup_Success_Rate_service] as'RRC Establishmnet SR', [E-RAB_Setup_Success_Rate] as 'Initial ERAB Establishmnet Success Rate', [E-RAB_Setup_Success_Rate(Hu_Cell)] as 'E-RAB Stetp SR', [Average_Downlink_User_Throughput(Mbit/s)] as 'UE DL Throughput (Mbps)', 
[Average_UPlink_User_Throughput(Mbit/s)] as 'UE UL Throughput (Mbps)', [Intra_RAT_Handover_SR_Intra+Inter_frequency(Huawei_LTE_Cell] as 'Handover Execution Rate',  [Call_Drop_Rate] as 'ERAB Drop Rate', 
[S1Signal_E-RAB_Setup_SR(Hu_Cell)] as 'S1 Signal Establishment Success Rate', [InterF_HOOut_SR] as 'InterF Handover Execution Rate', [IntraF_HOOut_SR] as 'IntraF Handover Execution Rate', [Average_UL_Packet_Loss_%(Huawei_LTE_UCell)] as 'UE UL Packet Loss Rate'
from [dbo].[TBL_LTE_CELL_Daily_H] where 
(substring([eNodeB],1,2)='KJ' or substring([eNodeB],1,2)='CH' or substring([eNodeB],1,2)='AS' or substring([eNodeB],1,2)='GL' or 
                                                                                  substring([eNodeB],1,2)='GN'  or substring([eNodeB],1,2)='KM' or substring([eNodeB],1,2)='KH' or substring([eNodeB],1,2)='KZ' or substring([eNodeB],1,2)='MA'
																				  or substring([eNodeB],1,2)='SM' or substring([eNodeB],1,2)='TH' or substring([eNodeB],1,2)='AG' or substring([eNodeB],1,2)='YZ')
            and Datetime  >='" + Date_of_WPC_7 + "' and Datetime  <'" + Date_of_WPC + "'";




                SqlCommand Huawei_4G_TBL_Quary_7 = new SqlCommand(Huawei_4G_TBL_7, connection);
                Huawei_4G_TBL_Quary_7.CommandTimeout = 0;
                Huawei_4G_TBL_Quary_7.ExecuteNonQuery();
                DataTable Huawei_4G_Table_7 = new DataTable();
                SqlDataAdapter dataAdapter_Huawei_4G_7 = new SqlDataAdapter(Huawei_4G_TBL_Quary_7);
                dataAdapter_Huawei_4G_7.Fill(Huawei_4G_Table_7);





                string Nokia_4G_TBL_7 = @"select [Date] as 'Date', [ElementID1], [Total_Payload_GB(Nokia_LTE_CELL)] as 'Paylaod (GB)', [cell_availability_include_manual_blocking(Nokia_LTE_CELL)]  as 'Availability', [RRC_Connection_Setup_Success_Ratio(Nokia_LTE_CELL)] as'RRC Establishmnet SR', [Initial_E-RAB_Setup_Success_Ratio(Nokia_LTE_CELL)] as 'Initial ERAB Establishmnet Success Rate', [E-RAB_Setup_SR_incl_added(Nokia_LTE_CELL)] as 'E-RAB Stetp SR', [User_Throughput_DL_mbps(Nokia_LTE_CELL)] as 'UE DL Throughput (Mbps)', 
[User_Throughput_UL_mbps(Nokia_LTE_CELL)] as 'UE UL Throughput (Mbps)', [Intra_RAT_Handover_SR_Intra+Inter_frequency(Nokia_LTE_CELL)] as 'Handover Execution Rate',  [E-RAB_Drop_Ratio_RAN_View(Nokia_LTE_CELL)] as 'ERAB Drop Rate', 
[S1Signal_E-RAB_Setup_SR(Nokia_LTE_CELL)] as 'S1 Signal Establishment Success Rate', [Inter-Freq_HO_SR(Nokia_LTE_CELL)] as 'InterF Handover Execution Rate', [HO_Success_Ratio_intra_eNB(Nokia_LTE_CELL)] as 'IntraF Handover Execution Rate', [Packet_loss_UL(Nokia_EUCELL)] as 'UE UL Packet Loss Rate'
from [dbo].[TBL_LTE_CELL_Daily_N] where 
(substring([ElementID1],1,2)='KJ' or substring([ElementID1],1,2)='CH' or substring([ElementID1],1,2)='AS' or substring([ElementID1],1,2)='GL' or 
                                                                                  substring([ElementID1],1,2)='GN'  or substring([ElementID1],1,2)='KM' or substring([ElementID1],1,2)='KH' or substring([ElementID1],1,2)='KZ' or substring([ElementID1],1,2)='MA'
																				  or substring([ElementID1],1,2)='SM' or substring([ElementID1],1,2)='TH' or substring([ElementID1],1,2)='AG' or substring([ElementID1],1,2)='YZ')
            and Date  >='" + Date_of_WPC_7 + "' and Date  <'" + Date_of_WPC + "'";




                SqlCommand Nokia_4G_TBL_Quary_7 = new SqlCommand(Nokia_4G_TBL_7, connection);
                Nokia_4G_TBL_Quary_7.CommandTimeout = 0;
                Nokia_4G_TBL_Quary_7.ExecuteNonQuery();
                DataTable Nokia_4G_Table_7 = new DataTable();
                SqlDataAdapter dataAdapter_Nokia_4G_7 = new SqlDataAdapter(Nokia_4G_TBL_Quary_7);
                dataAdapter_Nokia_4G_7.Fill(Nokia_4G_Table_7);




                // Update Vendor an Site
                Ericsson_4G_Table.Columns.Add("Vendor", typeof(string));
                Huawei_4G_Table.Columns.Add("Vendor", typeof(string));
                Nokia_4G_Table.Columns.Add("Vendor", typeof(string));
                Ericsson_4G_Table.Columns.Add("Contractor", typeof(string));
                Huawei_4G_Table.Columns.Add("Contractor", typeof(string));
                Nokia_4G_Table.Columns.Add("Contractor", typeof(string));
                Ericsson_4G_Table.Columns.Add("Province", typeof(string));
                Huawei_4G_Table.Columns.Add("Province", typeof(string));
                Nokia_4G_Table.Columns.Add("Province", typeof(string));
                Ericsson_4G_Table.Columns.Add("Site", typeof(string));
                Huawei_4G_Table.Columns.Add("Site", typeof(string));
                Nokia_4G_Table.Columns.Add("Site", typeof(string));
                Ericsson_4G_Table.Columns.Add("Coverage Type", typeof(string));
                Huawei_4G_Table.Columns.Add("Coverage Type", typeof(string));
                Nokia_4G_Table.Columns.Add("Coverage Type", typeof(string));


                string province_letter = "";
                string cell = "";
                for (int i = 0; i < Ericsson_4G_Table.Rows.Count; i++)
                {
                    cell = Ericsson_4G_Table.Rows[i][1].ToString();
                    Ericsson_4G_Table.Rows[i][15] = "Ericsson";
                    if (cell != "" && cell.Length == 10)
                    {
                        Ericsson_4G_Table.Rows[i][18] = cell.Substring(0, 2) + cell.Substring(4, 4);
                    }
                    else
                    {
                        Ericsson_4G_Table.Rows[i][18] = "NA";
                    }
                }
                for (int i = 0; i < Huawei_4G_Table.Rows.Count; i++)
                {
                    cell = Huawei_4G_Table.Rows[i][1].ToString();
                    Huawei_4G_Table.Rows[i][15] = "Huawei";
                    if (cell != "" && cell.Length == 10)
                    {
                        Huawei_4G_Table.Rows[i][18] = cell.Substring(0, 2) + cell.Substring(4, 4);
                    }
                    else
                    {
                        Huawei_4G_Table.Rows[i][18] = "NA";
                    }
                }
                for (int i = 0; i < Nokia_4G_Table.Rows.Count; i++)
                {
                    cell = Nokia_4G_Table.Rows[i][1].ToString();
                    Nokia_4G_Table.Rows[i][15] = "Nokia";
                    if (cell != "" && cell.Length == 10)
                    {
                        Nokia_4G_Table.Rows[i][18] = cell.Substring(0, 2) + cell.Substring(4, 4);
                    }
                    else
                    {
                        Nokia_4G_Table.Rows[i][18] = "NA";
                    }
                }


                //Union Tables in 3 Vendor
                dtUnion_4G = Ericsson_4G_Table.AsEnumerable().Union(Huawei_4G_Table.AsEnumerable()).CopyToDataTable<DataRow>();
                Table_4G = dtUnion_4G.AsEnumerable().Union(Nokia_4G_Table.AsEnumerable()).CopyToDataTable<DataRow>();




                // join with ARAS Table to get Coverge Type
                var JoinResult = (from p in Table_4G.AsEnumerable()
                                  join t in ARAS_Table.AsEnumerable()
                                  on p.Field<string>("Site") equals t.Field<string>("LOCATION")
                                  select new
                                  {
                                      //BSC = p.Field<string>("BSC"),
                                      Site = p.Field<string>("Site"),
                                      Coverage = t.Field<string>("COVERAGE_TYPE_OPTIMIZATION")
                                  }).ToList();


                // Update Province and Contractor
                progressBar1.Minimum = 0;
                progressBar1.Maximum = Table_4G.Rows.Count;
                int Row_Count_of_WPC_Table = 0;
                // for (int i = 0; i < 1000; i++)
                for (int i = 0; i < Table_4G.Rows.Count; i++)
                {
                    string Vendor = Table_4G.Rows[i][15].ToString();
                    string Site = Table_4G.Rows[i][18].ToString();
                    int Found = 0;
                    int ind = 0;
                    string Coverage_type = "";
                    for (int k = 0; k < JoinResult.Count; k++)
                    {
                        if (Site == JoinResult[k].Site.ToString())
                        {
                            Found = 1;
                            ind = k;
                            Coverage_type = JoinResult[ind].Coverage.ToString();
                            break;
                        }
                    }
                    if (Found == 0)
                    {
                        Coverage_type = "City";
                    }


                    cell = Table_4G.Rows[i][1].ToString();
                    if (cell != "")
                    {
                        province_letter = cell.Substring(0, 2);
                        if (province_letter == "TH")
                        {
                            Table_4G.Rows[i][16] = "NAK-Tehran"; Table_4G.Rows[i][17] = "Tehran"; Table_4G.Rows[i][19] = Coverage_type;
                        }
                        if (province_letter == "KJ")
                        {
                            Table_4G.Rows[i][16] = "NAK-Alborz"; Table_4G.Rows[i][17] = "Alborz"; Table_4G.Rows[i][19] = Coverage_type;
                        }
                        if (province_letter == "MA")
                        {
                            Table_4G.Rows[i][16] = "NAK-North"; Table_4G.Rows[i][17] = "Mazandaran"; Table_4G.Rows[i][19] = Coverage_type;
                        }
                        if (province_letter == "GN")
                        {
                            Table_4G.Rows[i][16] = "NAK-North"; Table_4G.Rows[i][17] = "Gilan"; Table_4G.Rows[i][19] = Coverage_type;
                        }
                        if (province_letter == "GL")
                        {
                            Table_4G.Rows[i][16] = "NAK-North"; Table_4G.Rows[i][17] = "Golestan"; Table_4G.Rows[i][19] = Coverage_type;
                        }
                        if (province_letter == "AS")
                        {
                            Table_4G.Rows[i][16] = "NAK-Huawei"; Table_4G.Rows[i][17] = "East Azarbaijan"; Table_4G.Rows[i][19] = Coverage_type;
                        }
                        if (province_letter == "AG")
                        {
                            Table_4G.Rows[i][16] = "NAK-Huawei"; Table_4G.Rows[i][17] = "West Azarbaijan"; Table_4G.Rows[i][19] = Coverage_type;
                        }
                        if (province_letter == "KZ")
                        {
                            Table_4G.Rows[i][16] = "NAK-Huawei"; Table_4G.Rows[i][17] = "Khuzestan"; Table_4G.Rows[i][19] = Coverage_type;
                        }
                        if (province_letter == "KH")
                        {
                            Table_4G.Rows[i][16] = "NAK-Nokia"; Table_4G.Rows[i][17] = "Khorasan Razavi"; Table_4G.Rows[i][19] = Coverage_type;
                        }
                        if (province_letter == "YZ")
                        {
                            Table_4G.Rows[i][16] = "NAK-Nokia"; Table_4G.Rows[i][17] = "Yazd"; Table_4G.Rows[i][19] = Coverage_type;
                        }
                        if (province_letter == "SM")
                        {
                            Table_4G.Rows[i][16] = "NAK-Nokia"; Table_4G.Rows[i][17] = "Semnan"; Table_4G.Rows[i][19] = Coverage_type;
                        }
                        if (province_letter == "CH")
                        {
                            Table_4G.Rows[i][16] = "NAK-Nokia"; Table_4G.Rows[i][17] = "Chahar Mahal Va Bakhtiari"; Table_4G.Rows[i][19] = Coverage_type;
                        }
                        if (province_letter == "KM")
                        {
                            Table_4G.Rows[i][16] = "NAK-Nokia"; Table_4G.Rows[i][17] = "Kerman"; Table_4G.Rows[i][19] = Coverage_type;
                        }
                    }

                    string Contractor = Table_4G.Rows[i][16].ToString();
                    string Province = Table_4G.Rows[i][17].ToString();
                    //string BSC = Table_4G.Rows[i][1].ToString();
                    double Payload = -1000;
                    double RRC_Connection_SR = -1000;
                    double ERAB_Drop_Rate = -1000;
                    double DL_THR = -1000;
                    double S1_Signalling_SR = -1000;
                    double Inter_Freq_SR = -1000;
                    double ERAB_SR_Added = -1000;
                    double Intra_Freq_SR = -1000;
                    double ERAB_SR_Initial = -1000;
                    double UL_THR = -1000;
                    double HO_SR = -1000;
                    double Availability = -1000;
                    double UL_Packet_Loss = -1000;

                    if (Table_4G.Rows[i][2].ToString() != "")
                    {
                        Payload = Convert.ToDouble(Table_4G.Rows[i][2]);
                    }
                    if (Table_4G.Rows[i][4].ToString() != "")
                    {
                        RRC_Connection_SR = Convert.ToDouble(Table_4G.Rows[i][4]);
                    }
                    if (Table_4G.Rows[i][5].ToString() != "")
                    {
                        ERAB_SR_Initial = Convert.ToDouble(Table_4G.Rows[i][5]);
                    }
                    if (Table_4G.Rows[i][6].ToString() != "")
                    {
                        ERAB_SR_Added = Convert.ToDouble(Table_4G.Rows[i][6]);
                    }
                    if (Table_4G.Rows[i][7].ToString() != "")
                    {
                        DL_THR = Convert.ToDouble(Table_4G.Rows[i][7]);
                    }
                    if (Table_4G.Rows[i][8].ToString() != "")
                    {
                        UL_THR = Convert.ToDouble(Table_4G.Rows[i][8]);
                    }
                    if (Table_4G.Rows[i][9].ToString() != "")
                    {
                        HO_SR = Convert.ToDouble(Table_4G.Rows[i][9]);
                    }
                    if (Table_4G.Rows[i][10].ToString() != "")
                    {
                        ERAB_Drop_Rate = Convert.ToDouble(Table_4G.Rows[i][10]);
                    }
                    if (Table_4G.Rows[i][11].ToString() != "")
                    {
                        S1_Signalling_SR = Convert.ToDouble(Table_4G.Rows[i][11]);
                    }
                    if (Table_4G.Rows[i][12].ToString() != "")
                    {
                        Inter_Freq_SR = Convert.ToDouble(Table_4G.Rows[i][12]);
                    }
                    if (Table_4G.Rows[i][13].ToString() != "")
                    {
                        Intra_Freq_SR = Convert.ToDouble(Table_4G.Rows[i][13]);
                    }
                    if (Table_4G.Rows[i][3].ToString() != "")
                    {
                        Availability = Convert.ToDouble(Table_4G.Rows[i][3]);
                    }
                    if (Table_4G.Rows[i][14].ToString() != "")
                    {
                        UL_Packet_Loss = Convert.ToDouble(Table_4G.Rows[i][14]);
                    }

                    // Fill WPc Table
                    int TH_Index = 0;


                    var q1 = (from p in BASE_Table.AsEnumerable()
                              where p.Field<string>("Cell") == cell
                              select p).ToList();
                    var q2 = q1;

                    //string RRC_Connection_SR_BL = "";
                    //string ERAB_SR_Initial_BL = "";
                    //string ERAB_SR_Added_BL = "";
                    //string DL_THR_BL = "";
                    //string UL_THR_BL = "";
                    //string HO_SR_BL = "";
                    //string ERAB_Drop_Rate_BL = "";
                    //string S1_Signalling_SR_BL = "";
                    //string Inter_Freq_SR_BL = "";
                    //string Intra_Freq_SR_BL = "";
                    //if (q1.Count != 0)
                    //{

                    //}




                    if (Vendor == "Ericsson" && Coverage_type == "City")
                    {
                        q1 = (from p in Ericsson_4G_Table_7.AsEnumerable()
                              where p.Field<string>("eNodeB") == cell
                              select p).ToList();

                        TH_Index = 0;
                    }
                    if (Vendor == "Ericsson" && Coverage_type == "None_City")
                    {
                        q1 = (from p in Ericsson_4G_Table_7.AsEnumerable()
                              where p.Field<string>("eNodeB") == cell
                              select p).ToList();

                        TH_Index = 1;
                    }
                    if (Vendor == "Huawei" && Coverage_type == "City")
                    {
                        q1 = (from p in Huawei_4G_Table_7.AsEnumerable()
                              where p.Field<string>("eNodeB") == cell
                              select p).ToList();

                        TH_Index = 2;
                    }
                    if (Vendor == "Huawei" && Coverage_type == "None_City")
                    {
                        q1 = (from p in Huawei_4G_Table_7.AsEnumerable()
                              where p.Field<string>("eNodeB") == cell
                              select p).ToList();

                        TH_Index = 3;
                    }
                    if (Vendor == "Nokia" && Coverage_type == "City")
                    {
                        q1 = (from p in Nokia_4G_Table_7.AsEnumerable()
                              where p.Field<string>("ElementID1") == cell
                              select p).ToList();

                        TH_Index = 4;
                    }
                    if (Vendor == "Nokia" && Coverage_type == "None_City")
                    {
                        q1 = (from p in Nokia_4G_Table_7.AsEnumerable()
                              where p.Field<string>("ElementID1") == cell
                              select p).ToList();

                        TH_Index = 5;
                    }



                    int count = 0;
                    int update = 0;
                    if (RRC_Connection_SR > -1000 && RRC_Connection_SR < TH_4G[TH_Index, 0])
                    {
                        count++;
                        int KPI_Index = 4;
                        if (q1.Count == 7)
                        {
                            Table_4G_WPC.Rows.Add(Contractor, Province, Vendor, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, RRC_Connection_SR, "RRC_Connection_SR", "-", q1[6].ItemArray[KPI_Index], q1[5].ItemArray[KPI_Index], q1[4].ItemArray[KPI_Index], q1[3].ItemArray[KPI_Index], q1[2].ItemArray[KPI_Index], q1[1].ItemArray[KPI_Index], q1[0].ItemArray[KPI_Index]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count != 7 && q1.Count != 0)
                        {
                            double[] KPI_of_Days = new double[7];
                            for (int j = 1; j <= 7; j++)
                            {
                                if (q1.Count - j < 0)
                                {
                                    KPI_of_Days[j - 1] = -20000;
                                }
                                else
                                {
                                    DateTime d1 = Convert.ToDateTime(q1[q1.Count - j].ItemArray[0]);
                                    if (d1 == Date_of_WPC.AddDays(-j))
                                    {
                                        if (q1[q1.Count - j].ItemArray[KPI_Index].ToString() == "")
                                        {
                                            KPI_of_Days[j - 1] = -20000;
                                        }
                                        else
                                        {
                                            KPI_of_Days[j - 1] = Convert.ToDouble(q1[q1.Count - j].ItemArray[KPI_Index]);
                                        }

                                    }
                                    else
                                    {
                                        KPI_of_Days[j - 1] = -20000;
                                    }
                                }
                            }
                            Table_4G_WPC.Rows.Add(Contractor, Province, Vendor, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, RRC_Connection_SR, "RRC_Connection_SR", "-", KPI_of_Days[0], KPI_of_Days[1], KPI_of_Days[2], KPI_of_Days[3], KPI_of_Days[4], KPI_of_Days[5], KPI_of_Days[6]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count == 0)
                        {
                            Table_4G_WPC.Rows.Add(Contractor, Province, Vendor, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, RRC_Connection_SR, "RRC_Connection_SR");
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                    }
                    if (ERAB_SR_Initial > -1000 && ERAB_SR_Initial < TH_4G[TH_Index, 1])
                    {
                        count++;
                        int KPI_Index = 5;
                        if (q1.Count == 7)
                        {
                            Table_4G_WPC.Rows.Add(Contractor, Province, Vendor, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, ERAB_SR_Initial, "ERAB_SR_Initial", "-", q1[6].ItemArray[KPI_Index], q1[5].ItemArray[KPI_Index], q1[4].ItemArray[KPI_Index], q1[3].ItemArray[KPI_Index], q1[2].ItemArray[KPI_Index], q1[1].ItemArray[KPI_Index], q1[0].ItemArray[KPI_Index]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count != 7 && q1.Count != 0)
                        {
                            double[] KPI_of_Days = new double[7];
                            for (int j = 1; j <= 7; j++)
                            {
                                if (q1.Count - j < 0)
                                {
                                    KPI_of_Days[j - 1] = -20000;
                                }
                                else
                                {
                                    DateTime d1 = Convert.ToDateTime(q1[q1.Count - j].ItemArray[0]);
                                    if (d1 == Date_of_WPC.AddDays(-j))
                                    {
                                        if (q1[q1.Count - j].ItemArray[KPI_Index].ToString() == "")
                                        {
                                            KPI_of_Days[j - 1] = -20000;
                                        }
                                        else
                                        {
                                            KPI_of_Days[j - 1] = Convert.ToDouble(q1[q1.Count - j].ItemArray[KPI_Index]);
                                        }

                                    }
                                    else
                                    {
                                        KPI_of_Days[j - 1] = -20000;
                                    }
                                }
                            }
                            Table_4G_WPC.Rows.Add(Contractor, Province, Vendor, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, ERAB_SR_Initial, "ERAB_SR_Initial", "-", KPI_of_Days[0], KPI_of_Days[1], KPI_of_Days[2], KPI_of_Days[3], KPI_of_Days[4], KPI_of_Days[5], KPI_of_Days[6]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count == 0)
                        {
                            Table_4G_WPC.Rows.Add(Contractor, Province, Vendor, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, ERAB_SR_Initial, "ERAB_SR_Initial");
                            Row_Count_of_WPC_Table++; update = 1;
                        }

                    }
                    if (ERAB_SR_Added > -1000 && ERAB_SR_Added < TH_4G[TH_Index, 2])
                    {
                        count++;
                        int KPI_Index = 6;
                        if (q1.Count == 7)
                        {
                            Table_4G_WPC.Rows.Add(Contractor, Province, Vendor, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, ERAB_SR_Added, "ERAB_SR_Added", "-", q1[6].ItemArray[KPI_Index], q1[5].ItemArray[KPI_Index], q1[4].ItemArray[KPI_Index], q1[3].ItemArray[KPI_Index], q1[2].ItemArray[KPI_Index], q1[1].ItemArray[KPI_Index], q1[0].ItemArray[KPI_Index]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count != 7 && q1.Count != 0)
                        {
                            double[] KPI_of_Days = new double[7];
                            for (int j = 1; j <= 7; j++)
                            {
                                if (q1.Count - j < 0)
                                {
                                    KPI_of_Days[j - 1] = -20000;
                                }
                                else
                                {
                                    DateTime d1 = Convert.ToDateTime(q1[q1.Count - j].ItemArray[0]);
                                    if (d1 == Date_of_WPC.AddDays(-j))
                                    {
                                        if (q1[q1.Count - j].ItemArray[KPI_Index].ToString() == "")
                                        {
                                            KPI_of_Days[j - 1] = -20000;
                                        }
                                        else
                                        {
                                            KPI_of_Days[j - 1] = Convert.ToDouble(q1[q1.Count - j].ItemArray[KPI_Index]);
                                        }

                                    }
                                    else
                                    {
                                        KPI_of_Days[j - 1] = -20000;
                                    }
                                }
                            }
                            Table_4G_WPC.Rows.Add(Contractor, Province, Vendor, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, ERAB_SR_Added, "ERAB_SR_Added", "-", KPI_of_Days[0], KPI_of_Days[1], KPI_of_Days[2], KPI_of_Days[3], KPI_of_Days[4], KPI_of_Days[5], KPI_of_Days[6]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count == 0)
                        {
                            Table_4G_WPC.Rows.Add(Contractor, Province, Vendor, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, ERAB_SR_Added, "ERAB_SR_Added");
                            Row_Count_of_WPC_Table++; update = 1;
                        }

                    }
                    if (DL_THR > -1000 && DL_THR < TH_4G[TH_Index, 3])
                    {
                        count++;
                        int KPI_Index = 7;
                        if (q1.Count == 7)
                        {
                            Table_4G_WPC.Rows.Add(Contractor, Province, Vendor, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, DL_THR, "DL_THR (Mbps)", "-", q1[6].ItemArray[KPI_Index], q1[5].ItemArray[KPI_Index], q1[4].ItemArray[KPI_Index], q1[3].ItemArray[KPI_Index], q1[2].ItemArray[KPI_Index], q1[1].ItemArray[KPI_Index], q1[0].ItemArray[KPI_Index]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count != 7 && q1.Count != 0)
                        {
                            double[] KPI_of_Days = new double[7];
                            for (int j = 1; j <= 7; j++)
                            {
                                if (q1.Count - j < 0)
                                {
                                    KPI_of_Days[j - 1] = -20000;
                                }
                                else
                                {
                                    DateTime d1 = Convert.ToDateTime(q1[q1.Count - j].ItemArray[0]);
                                    if (d1 == Date_of_WPC.AddDays(-j))
                                    {
                                        if (q1[q1.Count - j].ItemArray[KPI_Index].ToString() == "")
                                        {
                                            KPI_of_Days[j - 1] = -20000;
                                        }
                                        else
                                        {
                                            KPI_of_Days[j - 1] = Convert.ToDouble(q1[q1.Count - j].ItemArray[KPI_Index]);
                                        }

                                    }
                                    else
                                    {
                                        KPI_of_Days[j - 1] = -20000;
                                    }
                                }
                            }
                            Table_4G_WPC.Rows.Add(Contractor, Province, Vendor, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, DL_THR, "DL_THR (Mbps)", "-", KPI_of_Days[0], KPI_of_Days[1], KPI_of_Days[2], KPI_of_Days[3], KPI_of_Days[4], KPI_of_Days[5], KPI_of_Days[6]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count == 0)
                        {
                            Table_4G_WPC.Rows.Add(Contractor, Province, Vendor, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, DL_THR, "DL_THR (Mbps)");
                            Row_Count_of_WPC_Table++; update = 1;
                        }

                    }
                    if (UL_THR > -1000 && UL_THR < TH_4G[TH_Index, 4])
                    {
                        count++;
                        int KPI_Index = 8;
                        if (q1.Count == 7)
                        {
                            Table_4G_WPC.Rows.Add(Contractor, Province, Vendor, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, UL_THR, "UL_THR (Mbps)", "-", q1[6].ItemArray[KPI_Index], q1[5].ItemArray[KPI_Index], q1[4].ItemArray[KPI_Index], q1[3].ItemArray[KPI_Index], q1[2].ItemArray[KPI_Index], q1[1].ItemArray[KPI_Index], q1[0].ItemArray[KPI_Index]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count != 7 && q1.Count != 0)
                        {
                            double[] KPI_of_Days = new double[7];
                            for (int j = 1; j <= 7; j++)
                            {
                                if (q1.Count - j < 0)
                                {
                                    KPI_of_Days[j - 1] = -20000;
                                }
                                else
                                {
                                    DateTime d1 = Convert.ToDateTime(q1[q1.Count - j].ItemArray[0]);
                                    if (d1 == Date_of_WPC.AddDays(-j))
                                    {
                                        if (q1[q1.Count - j].ItemArray[KPI_Index].ToString() == "")
                                        {
                                            KPI_of_Days[j - 1] = -20000;
                                        }
                                        else
                                        {
                                            KPI_of_Days[j - 1] = Convert.ToDouble(q1[q1.Count - j].ItemArray[KPI_Index]);
                                        }

                                    }
                                    else
                                    {
                                        KPI_of_Days[j - 1] = -20000;
                                    }
                                }
                            }
                            Table_4G_WPC.Rows.Add(Contractor, Province, Vendor, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, UL_THR, "UL_THR (Mbps)", "-", KPI_of_Days[0], KPI_of_Days[1], KPI_of_Days[2], KPI_of_Days[3], KPI_of_Days[4], KPI_of_Days[5], KPI_of_Days[6]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count == 0)
                        {
                            Table_4G_WPC.Rows.Add(Contractor, Province, Vendor, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, UL_THR, "UL_THR (Mbps)");
                            Row_Count_of_WPC_Table++; update = 1;
                        }

                    }
                    if (HO_SR > -1000 && HO_SR < TH_4G[TH_Index, 5])
                    {
                        count++;
                        int KPI_Index = 9;
                        if (q1.Count == 7)
                        {
                            Table_4G_WPC.Rows.Add(Contractor, Province, Vendor, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, HO_SR, "HO_SR", "-", q1[6].ItemArray[KPI_Index], q1[5].ItemArray[KPI_Index], q1[4].ItemArray[KPI_Index], q1[3].ItemArray[KPI_Index], q1[2].ItemArray[KPI_Index], q1[1].ItemArray[KPI_Index], q1[0].ItemArray[KPI_Index]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count != 7 && q1.Count != 0)
                        {
                            double[] KPI_of_Days = new double[7];
                            for (int j = 1; j <= 7; j++)
                            {
                                if (q1.Count - j < 0)
                                {
                                    KPI_of_Days[j - 1] = -20000;
                                }
                                else
                                {
                                    DateTime d1 = Convert.ToDateTime(q1[q1.Count - j].ItemArray[0]);
                                    if (d1 == Date_of_WPC.AddDays(-j))
                                    {
                                        if (q1[q1.Count - j].ItemArray[KPI_Index].ToString() == "")
                                        {
                                            KPI_of_Days[j - 1] = -20000;
                                        }
                                        else
                                        {
                                            KPI_of_Days[j - 1] = Convert.ToDouble(q1[q1.Count - j].ItemArray[KPI_Index]);
                                        }

                                    }
                                    else
                                    {
                                        KPI_of_Days[j - 1] = -20000;
                                    }
                                }
                            }
                            Table_4G_WPC.Rows.Add(Contractor, Province, Vendor, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, HO_SR, "HO_SR", "-", KPI_of_Days[0], KPI_of_Days[1], KPI_of_Days[2], KPI_of_Days[3], KPI_of_Days[4], KPI_of_Days[5], KPI_of_Days[6]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count == 0)
                        {
                            Table_4G_WPC.Rows.Add(Contractor, Province, Vendor, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, HO_SR, "HO_SR");
                            Row_Count_of_WPC_Table++; update = 1;
                        }

                    }
                    if (ERAB_Drop_Rate > -1000 && ERAB_Drop_Rate > TH_4G[TH_Index, 6])
                    {
                        count++;
                        int KPI_Index = 10;
                        if (q1.Count == 7)
                        {
                            Table_4G_WPC.Rows.Add(Contractor, Province, Vendor, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, ERAB_Drop_Rate, "ERAB_Drop_Rate", "-", q1[6].ItemArray[KPI_Index], q1[5].ItemArray[KPI_Index], q1[4].ItemArray[KPI_Index], q1[3].ItemArray[KPI_Index], q1[2].ItemArray[KPI_Index], q1[1].ItemArray[KPI_Index], q1[0].ItemArray[KPI_Index]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count != 7 && q1.Count != 0)
                        {
                            double[] KPI_of_Days = new double[7];
                            for (int j = 1; j <= 7; j++)
                            {
                                if (q1.Count - j < 0)
                                {
                                    KPI_of_Days[j - 1] = -20000;
                                }
                                else
                                {
                                    DateTime d1 = Convert.ToDateTime(q1[q1.Count - j].ItemArray[0]);
                                    if (d1 == Date_of_WPC.AddDays(-j))
                                    {
                                        if (q1[q1.Count - j].ItemArray[KPI_Index].ToString() == "")
                                        {
                                            KPI_of_Days[j - 1] = -20000;
                                        }
                                        else
                                        {
                                            KPI_of_Days[j - 1] = Convert.ToDouble(q1[q1.Count - j].ItemArray[KPI_Index]);
                                        }

                                    }
                                    else
                                    {
                                        KPI_of_Days[j - 1] = -20000;
                                    }
                                }
                            }
                            Table_4G_WPC.Rows.Add(Contractor, Province, Vendor, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, ERAB_Drop_Rate, "ERAB_Drop_Rate", "-", KPI_of_Days[0], KPI_of_Days[1], KPI_of_Days[2], KPI_of_Days[3], KPI_of_Days[4], KPI_of_Days[5], KPI_of_Days[6]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count == 0)
                        {
                            Table_4G_WPC.Rows.Add(Contractor, Province, Vendor, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, ERAB_Drop_Rate, "ERAB_Drop_Rate");
                            Row_Count_of_WPC_Table++; update = 1;
                        }

                    }
                    if (S1_Signalling_SR > -1000 && S1_Signalling_SR < TH_4G[TH_Index, 7])
                    {
                        count++;
                        int KPI_Index = 11;
                        if (q1.Count == 7)
                        {
                            Table_4G_WPC.Rows.Add(Contractor, Province, Vendor, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, S1_Signalling_SR, "S1_Signalling_SR", "-", q1[6].ItemArray[KPI_Index], q1[5].ItemArray[KPI_Index], q1[4].ItemArray[KPI_Index], q1[3].ItemArray[KPI_Index], q1[2].ItemArray[KPI_Index], q1[1].ItemArray[KPI_Index], q1[0].ItemArray[KPI_Index]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count != 7 && q1.Count != 0)
                        {
                            double[] KPI_of_Days = new double[7];
                            for (int j = 1; j <= 7; j++)
                            {
                                if (q1.Count - j < 0)
                                {
                                    KPI_of_Days[j - 1] = -20000;
                                }
                                else
                                {
                                    DateTime d1 = Convert.ToDateTime(q1[q1.Count - j].ItemArray[0]);
                                    if (d1 == Date_of_WPC.AddDays(-j))
                                    {
                                        if (q1[q1.Count - j].ItemArray[KPI_Index].ToString() == "")
                                        {
                                            KPI_of_Days[j - 1] = -20000;
                                        }
                                        else
                                        {
                                            KPI_of_Days[j - 1] = Convert.ToDouble(q1[q1.Count - j].ItemArray[KPI_Index]);
                                        }

                                    }
                                    else
                                    {
                                        KPI_of_Days[j - 1] = -20000;
                                    }
                                }
                            }
                            Table_4G_WPC.Rows.Add(Contractor, Province, Vendor, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, S1_Signalling_SR, "S1_Signalling_SR", "-", KPI_of_Days[0], KPI_of_Days[1], KPI_of_Days[2], KPI_of_Days[3], KPI_of_Days[4], KPI_of_Days[5], KPI_of_Days[6]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count == 0)
                        {
                            Table_4G_WPC.Rows.Add(Contractor, Province, Vendor, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, S1_Signalling_SR, "S1_Signalling_SR");
                            Row_Count_of_WPC_Table++; update = 1;
                        }

                    }
                    if (Inter_Freq_SR > -1000 && Inter_Freq_SR < TH_4G[TH_Index, 8])
                    {
                        count++;
                        int KPI_Index = 12;

                        if (q1.Count == 7)
                        {
                            Table_4G_WPC.Rows.Add(Contractor, Province, Vendor, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, Inter_Freq_SR, "Inter_Freq_SR", "-", q1[6].ItemArray[KPI_Index], q1[5].ItemArray[KPI_Index], q1[4].ItemArray[KPI_Index], q1[3].ItemArray[KPI_Index], q1[2].ItemArray[KPI_Index], q1[1].ItemArray[KPI_Index], q1[0].ItemArray[KPI_Index]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count != 7 && q1.Count != 0)
                        {
                            double[] KPI_of_Days = new double[7];
                            for (int j = 1; j <= 7; j++)
                            {
                                if (q1.Count - j < 0)
                                {
                                    KPI_of_Days[j - 1] = -20000;
                                }
                                else
                                {
                                    DateTime d1 = Convert.ToDateTime(q1[q1.Count - j].ItemArray[0]);
                                    if (d1 == Date_of_WPC.AddDays(-j))
                                    {
                                        if (q1[q1.Count - j].ItemArray[KPI_Index].ToString() == "")
                                        {
                                            KPI_of_Days[j - 1] = -20000;
                                        }
                                        else
                                        {
                                            KPI_of_Days[j - 1] = Convert.ToDouble(q1[q1.Count - j].ItemArray[KPI_Index]);
                                        }

                                    }
                                    else
                                    {
                                        KPI_of_Days[j - 1] = -20000;
                                    }
                                }
                            }
                            Table_4G_WPC.Rows.Add(Contractor, Province, Vendor, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, Inter_Freq_SR, "Inter_Freq_SR", "-", KPI_of_Days[0], KPI_of_Days[1], KPI_of_Days[2], KPI_of_Days[3], KPI_of_Days[4], KPI_of_Days[5], KPI_of_Days[6]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count == 0)
                        {
                            Table_4G_WPC.Rows.Add(Contractor, Province, Vendor, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, Inter_Freq_SR, "Inter_Freq_SR");
                            Row_Count_of_WPC_Table++; update = 1;
                        }

                    }
                    if (Intra_Freq_SR > -1000 && Intra_Freq_SR < TH_4G[TH_Index, 9])
                    {
                        count++;
                        int KPI_Index = 13;
                        if (q1.Count == 7)
                        {
                            Table_4G_WPC.Rows.Add(Contractor, Province, Vendor, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, Intra_Freq_SR, "Intra_Freq_SR", "-", q1[6].ItemArray[KPI_Index], q1[5].ItemArray[KPI_Index], q1[4].ItemArray[KPI_Index], q1[3].ItemArray[KPI_Index], q1[2].ItemArray[KPI_Index], q1[1].ItemArray[KPI_Index], q1[0].ItemArray[KPI_Index]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count != 7 && q1.Count != 0)
                        {
                            double[] KPI_of_Days = new double[7];
                            for (int j = 1; j <= 7; j++)
                            {
                                if (q1.Count - j < 0)
                                {
                                    KPI_of_Days[j - 1] = -20000;
                                }
                                else
                                {
                                    DateTime d1 = Convert.ToDateTime(q1[q1.Count - j].ItemArray[0]);
                                    if (d1 == Date_of_WPC.AddDays(-j))
                                    {
                                        if (q1[q1.Count - j].ItemArray[KPI_Index].ToString() == "")
                                        {
                                            KPI_of_Days[j - 1] = -20000;
                                        }
                                        else
                                        {
                                            KPI_of_Days[j - 1] = Convert.ToDouble(q1[q1.Count - j].ItemArray[KPI_Index]);
                                        }

                                    }
                                    else
                                    {
                                        KPI_of_Days[j - 1] = -20000;
                                    }
                                }
                            }
                            Table_4G_WPC.Rows.Add(Contractor, Province, Vendor, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, Intra_Freq_SR, "Intra_Freq_SR", "-", KPI_of_Days[0], KPI_of_Days[1], KPI_of_Days[2], KPI_of_Days[3], KPI_of_Days[4], KPI_of_Days[5], KPI_of_Days[6]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count == 0)
                        {
                            Table_4G_WPC.Rows.Add(Contractor, Province, Vendor, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, Intra_Freq_SR, " Intra_Freq_SR");
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                    }

                    if (UL_Packet_Loss > -1000 && UL_Packet_Loss > TH_4G[TH_Index, 10])
                    {
                        count++;
                        int KPI_Index = 14;
                        if (q1.Count == 7)
                        {
                            Table_4G_WPC.Rows.Add(Contractor, Province, Vendor, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, UL_Packet_Loss, "UL_Packet_Loss", "-", q1[6].ItemArray[KPI_Index], q1[5].ItemArray[KPI_Index], q1[4].ItemArray[KPI_Index], q1[3].ItemArray[KPI_Index], q1[2].ItemArray[KPI_Index], q1[1].ItemArray[KPI_Index], q1[0].ItemArray[KPI_Index]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count != 7 && q1.Count != 0)
                        {
                            double[] KPI_of_Days = new double[7];
                            for (int j = 1; j <= 7; j++)
                            {
                                if (q1.Count - j < 0)
                                {
                                    KPI_of_Days[j - 1] = -20000;
                                }
                                else
                                {
                                    DateTime d1 = Convert.ToDateTime(q1[q1.Count - j].ItemArray[0]);
                                    if (d1 == Date_of_WPC.AddDays(-j))
                                    {
                                        if (q1[q1.Count - j].ItemArray[KPI_Index].ToString() == "")
                                        {
                                            KPI_of_Days[j - 1] = -20000;
                                        }
                                        else
                                        {
                                            KPI_of_Days[j - 1] = Convert.ToDouble(q1[q1.Count - j].ItemArray[KPI_Index]);
                                        }

                                    }
                                    else
                                    {
                                        KPI_of_Days[j - 1] = -20000;
                                    }
                                }
                            }
                            Table_4G_WPC.Rows.Add(Contractor, Province, Vendor, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, UL_Packet_Loss, "UL_Packet_Loss", "-", KPI_of_Days[0], KPI_of_Days[1], KPI_of_Days[2], KPI_of_Days[3], KPI_of_Days[4], KPI_of_Days[5], KPI_of_Days[6]);
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                        else if (q1.Count == 0)
                        {
                            Table_4G_WPC.Rows.Add(Contractor, Province, Vendor, Site, cell, Coverage_type, Date_of_WPC, Availability, Payload, UL_Packet_Loss, " UL_Packet_Loss");
                            Row_Count_of_WPC_Table++; update = 1;
                        }
                    }

                    if (update == 1)
                    {
                        Table_4G_WPC.Rows[Row_Count_of_WPC_Table - 1][11] = Convert.ToString(count);
                    }

                    progressBar1.Value = i;

                }




                XLWorkbook wb = new XLWorkbook();
                wb.Worksheets.Add(Table_4G_WPC, "WPC_4G");
                var saveFileDialog = new SaveFileDialog
                {
                    FileName = "WPC_4G",
                    Filter = "Excel files|*.xlsx",
                    Title = "Save an Excel File"
                };

                saveFileDialog.ShowDialog();

                if (!String.IsNullOrWhiteSpace(saveFileDialog.FileName))
                    wb.SaveAs(saveFileDialog.FileName);


                label4.Text = "Finished";
                label4.BackColor = Color.Yellow;
                MessageBox.Show("FINISHED");








            }




        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            Date_of_WPC = dateTimePicker1.Value.Date;
            Date_of_WPC_7 = Date_of_WPC.AddDays(-7);
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                Technology = "2G";
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                Technology = "3G";
                checkBox1.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked == true)
            {
                Technology = "4G";
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                checkBox4.Checked = false;
            }
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked == true)
            {
                Technology = "3G";
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                checkBox3.Checked = false;
            }
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked == true)
            {
                Technology = "2G";
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                checkBox10.Checked = false;
            }
        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox6.Checked == true)
            {
                Technology = "3G_CS";
                checkBox5.Checked = false;
                checkBox7.Checked = false;
                checkBox10.Checked = false;
            }
        }

        private void checkBox10_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox10.Checked == true)
            {
                Technology = "3G_PS";
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
            }
        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox7.Checked == true)
            {
                Technology = "4G";
                listBox2.Items.Clear();
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox10.Checked = false;
            }
        }

        private void checkBox8_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox8.Checked == true)
            {
                Interval = "Daily";
                checkBox9.Checked = false;
            }
        }

        private void checkBox9_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox9.Checked == true)
            {
                Interval = "BH";
                checkBox8.Checked = false;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            listBox3.Items.Clear();
            listBox2.Items.Clear();

            string Node_List_Quary = "";
            string Province_List_Quary = "";
            string p1 = "Province=";
            string Province_List = "";
            //string[] Added_Province_List = new string[5];
            //int Added_Province_Ind = 0;

            for (int k = 1; k <= listBox1.SelectedItems.Count; k++)
            {
                p1 = "Province ='" + listBox1.SelectedItems[k - 1].ToString() + "' or ";
                Province_List = Province_List + p1;

                if(listBox1.SelectedItems[k - 1].ToString()== "Chahar Mahal Va Bakhtiari")
                {
                    p1 = "Province ='Charmahal' or ";
                    Province_List = Province_List + p1;
                }
                if (listBox1.SelectedItems[k - 1].ToString() == "East Azarbaijan")
                {
                    p1 = "Province ='AZARSHARGHI' or ";
                    Province_List = Province_List + p1;
                }
                if (listBox1.SelectedItems[k - 1].ToString() == "Khorasan Razavi")
                {
                    p1 = "Province ='KHORASANRAZAVI' or ";
                    Province_List = Province_List + p1;
                }
                if (listBox1.SelectedItems[k - 1].ToString() == "West Azarbaijan")
                {
                    p1 = "Province ='AZARGHARBI' or ";
                    Province_List = Province_List + p1;
                }
            }
            if (Province_List != "")
            {
                Province_List = Province_List.Substring(0, Province_List.Length - 4);

            }

            if (checkBox5.Checked == true)
            {

                Node_List_Quary = "select * from [cc2_BSC_NEW] where Date = '" + Date_List + "' and [TCH_Traffic_24H] != 0 and (" + Province_List + ") order by Province";
                listBox3.Items.Add("CDR");
                listBox3.Items.Add("CSSR");
                listBox3.Items.Add("IHSR");
                listBox3.Items.Add("OHSR");
                listBox3.Items.Add("RxQual_DL");
                listBox3.Items.Add("RxQual_UL");
                listBox3.Items.Add("SDCCH_Access_SR");
                listBox3.Items.Add("SDCCH_Congestion");
                listBox3.Items.Add("SDCCH_Drop_Rate");
                listBox3.Items.Add("TCH_Assign_FR");
                listBox3.Items.Add("TCH_Congestion");
                listBox3.Items.Add("TCH_Traffic");
                listBox3.Items.Add("TCH_Availability");
            }
            if (checkBox6.Checked == true)
            {
                Node_List_Quary = "select * from [cc3_RNC_NEW] where Date = '" + Date_List + "' and [CS Traffic (24H) (Erlang)] != 0 and (" + Province_List + ") order by Province";

                listBox3.Items.Add("CS_RAB_Establish");
                listBox3.Items.Add("CS_IRAT_HO_SR");
                listBox3.Items.Add("CS_Drop_Rate");
                listBox3.Items.Add("Soft_HO_SR");
                listBox3.Items.Add("CS_RRC_SR");
                listBox3.Items.Add("CS_MultiRAB_SR");
                listBox3.Items.Add("Inter_Carrier_HO_SR");
                listBox3.Items.Add("CS_Traffic");
                listBox3.Items.Add("Cell_Availability");

            }
            if (checkBox10.Checked == true)
            {
                Node_List_Quary = "select * from [rd3_RNC_NEW] where Date = '" + Date_List + "' and [payload] != 0 and (" + Province_List + ") order by Province";
                listBox3.Items.Add("HSDPA_SR");
                listBox3.Items.Add("HSUPA_SR");
                listBox3.Items.Add("DL_User_THR (Mbps)");
                listBox3.Items.Add("UL_User_THR (Kbps)");
                listBox3.Items.Add("HSDAP_Drop_Rate");
                listBox3.Items.Add("HSUAP_Drop_Rate");
                listBox3.Items.Add("PS_RRC_SR");
                listBox3.Items.Add("Ps_RAB_Establish");
                listBox3.Items.Add("PS_MultiRAB_Establish");
                listBox3.Items.Add("PS_Drop_Rate");
                listBox3.Items.Add("HSDPA_Cell_Change_SR");
                listBox3.Items.Add("HS_Share_Payload");
                listBox3.Items.Add("DL_Cell_THR (Mbps)");
                listBox3.Items.Add("RSSI (dBm)");
                listBox3.Items.Add("Average CQI");
                listBox3.Items.Add("PS_Payload (GB)");
                listBox3.Items.Add("Cell_Availability");
            }
            if (checkBox7.Checked == true)
            {
                //  Node_List_Quary = "select * from [RD4_province_NEW_v2] where Date = '" + Date_List + "' and [Daily Total Payload (GB)] != 0 and (" + Province_List + ") order by Province";
                listBox3.Items.Add("RRC_Connection_SR");
                listBox3.Items.Add("ERAB_SR_Initial");
                listBox3.Items.Add("ERAB_SR_Added");
                listBox3.Items.Add("DL_THR (Mbps)");
                listBox3.Items.Add("UL_THR (Mbps)");
                listBox3.Items.Add("ERAB_Drop_Rate");
                listBox3.Items.Add("S1_Signalling_SR");
                listBox3.Items.Add("Intra_Freq_SR");
                listBox3.Items.Add("Inter_Freq_SR");
                listBox3.Items.Add("UL_Packet_Loss");
                listBox3.Items.Add("UE_DL_Latency (ms)");
                listBox3.Items.Add("Average CQI");
                listBox3.Items.Add("PUCCH_RSSI (dBm)");
                listBox3.Items.Add("PUSCH_RSSI (dBm)");
                listBox3.Items.Add("Total_Paylaod (GB)");
                listBox3.Items.Add("Cell_Availability");

            }





            if (Technology != "4G")
            {
                SqlCommand Node_List_Quary1 = new SqlCommand(Node_List_Quary, connection);
                Node_List_Quary1.ExecuteNonQuery();
                DataTable Table_Node_List_Quary = new DataTable();
                SqlDataAdapter dataAdapter_Node_List_Quary = new SqlDataAdapter(Node_List_Quary1);
                dataAdapter_Node_List_Quary.Fill(Table_Node_List_Quary);


                for (int i = 0; i < Table_Node_List_Quary.Rows.Count; i++)
                {
                    listBox2.Items.Add((Table_Node_List_Quary.Rows[i]).ItemArray[0]);
                    listBox2.SetSelected(i, true);
                    // listBox2.SelectedItems[i] = true;
                }
            }







        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            Date_List = dateTimePicker2.Value.Date;
            int rr = 0;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            
            //string ID = textBox4.Text;
            //if (ID == "56" || ID == "npo-290962" || ID == "npo-294031" || ID == "npo-294005" || ID == "npo-260346" || ID == "npo-290515" || ID == "npo-260410" || ID == "npo-298081")
            //{
                label15.Text = "Wait...";
                label15.BackColor = Color.Yellow;

                string KPI = listBox3.SelectedItem.ToString();
                //            string cell_index = "substring(Cell,1,2)=";
                //            string Province_Index_List = "";
                //            if (listBox1.SelectedItems.Count==0)
                //            {
                //                MessageBox.Show("Please select the list of Province(s)");
                //            }
                //            for (int k = 1; k <= listBox1.SelectedItems.Count; k++)
                //            {
                //                if (listBox1.SelectedItems[k-1].ToString()=="Alborz")
                //                {
                //                    Province_Index_List = Province_Index_List + cell_index + "'KJ' or ";
                //                }
                //                if (listBox1.SelectedItems[k - 1].ToString() == "AZARGHARBI")
                //                {
                //                    Province_Index_List = Province_Index_List + cell_index + "'AG' or ";
                //                }

                ////AZARSHARGHI
                ////Charmahal
                ////Gilan
                ////Golestan
                ////Kerman
                ////KHORASANRAZAVI
                ////Khuzestan
                ////Mazandaran
                ////Semnan
                ////Tehran
                ////Yazd

                //            }
                //            Province_Index_List = Province_Index_List.Substring(0, Province_Index_List.Length - 4);



                //string cell_index = "substring(Cell,1,2)=";
                string Node_Index_List = "";
                //if (listBox2.SelectedItems.Count == 0)
                //{
                //    MessageBox.Show("Please select the list of Node(s)");
                //}
                if (Technology != "4G")
                {
                    for (int k = 1; k <= listBox2.SelectedItems.Count; k++)
                    {
                        string Node = listBox2.SelectedItems[k - 1].ToString();
                        if (Node.Substring(Node.Length - 1, 1) == "E")
                        {
                            Ericsson_Count++;
                        }
                        if (Node.Substring(Node.Length - 1, 1) == "H")
                        {
                            Huawei_Count++;
                        }
                        if (Node.Substring(Node.Length - 1, 1) == "N")
                        {
                            Nokia_Count++;
                        }
                        if (Technology == "2G")
                        {
                            Node_Index_List = Node_Index_List + "BSC='" + Node + "' or ";
                        }
                        if (Technology == "3G_CS" || Technology == "3G_PS")
                        {
                            Node_Index_List = Node_Index_List + "ElementID='" + Node + "' or ";
                        }

                    }
                }
                if (Technology == "4G")
                {
                    for (int k = 1; k <= listBox1.SelectedItems.Count; k++)
                    {
                        string Node = listBox1.SelectedItems[k - 1].ToString();
                        Node_Index_List = Node_Index_List + "Province='" + Node + "' or ";

                        if (Node == "Chahar Mahal Va Bakhtiari")
                        {
                            Node_Index_List = Node_Index_List + "Province='Charmahal' or ";
                        }
                        if (Node == "East Azarbaijan")
                        {
                            Node_Index_List = Node_Index_List + "Province='AZARSHARGHI' or ";
                        }
                        if (Node == "Khorasan Razavi")
                        {
                            Node_Index_List = Node_Index_List + "Province='KHORASANRAZAVI' or ";
                        }
                        if (Node == "West Azarbaijan")
                        {
                            Node_Index_List = Node_Index_List + "Province='AZARGHARBI' or ";
                        }

                    }
                }

                if (Node_Index_List != "")
                {
                    Node_Index_List = Node_Index_List.Substring(0, Node_Index_List.Length - 4);
                }





                if (Interval == "Daily")
                {
                    string KPI_E = "";
                    string Traffic_E = "";
                    string Ave_E = "";
                    string KPI_H = "";
                    string Traffic_H = "";
                    string Ave_H = "";
                    string KPI_N = "";
                    string Traffic_N = "";
                    string Ave_N = "";
                    if (Technology == "2G")
                    {
                        Traffic_E = "[TCH_Traffic]"; Ave_E = "[TCH_Availability]";
                        Traffic_H = "[TCH_Traffic]"; Ave_H = "[TCH_Availability]";
                        Traffic_N = "[TCH_Traffic]"; Ave_N = "[TCH_Availability]";
                        if (KPI == "CSSR")
                        {
                            KPI_E = "[CSSR_MCI]";
                            KPI_H = "[CSSR3]";
                            KPI_N = "[CSSR_MCI]";
                        }
                        if (KPI == "IHSR")
                        {
                            KPI_E = "[IHSR]";
                            KPI_H = "[IHSR2]";
                            KPI_N = "[IHSR]";
                        }
                        if (KPI == "OHSR")
                        {
                            KPI_E = "[OHSR]";
                            KPI_H = "[OHSR2]";
                            KPI_N = "[OHSR]";
                        }
                        if (KPI == "CDR")
                        {
                            KPI_E = "[CDR(not Affected by incoming Handovers from 3G)(Eric_CELL)]";
                            KPI_H = "[CDR3]";
                            KPI_N = "[CDR(including_CS_IRAT_handovers_3G_to2G)(Nokia_SEG)]";
                        }
                        if (KPI == "RxQual_DL")
                        {
                            KPI_E = "[RxQual_DL]";
                            KPI_H = "[RX_QUALITTY_DL_NEW]";
                            KPI_N = "[RxQuality_DL]";
                        }
                        if (KPI == "RxQual_UL")
                        {
                            KPI_E = "[RxQual_UL]";
                            KPI_H = "[RX_QUALITTY_UL_NEW]";
                            KPI_N = "[RxQuality_UL]";
                        }
                        if (KPI == "SDCCH_Access_SR")
                        {
                            KPI_E = "[SDCCH_Access_Succ_Rate]";
                            KPI_H = "[SDCCH_Access_Success_Rate2]";
                            KPI_N = "[SDCCH_Access_Success_Rate]";
                        }
                        if (KPI == "SDCCH_Congestion")
                        {
                            KPI_E = "[SDCCH_Congestion]";
                            KPI_H = "[SDCCH_Congestion_Rate]";
                            KPI_N = "[SDCCH_Congestion_Rate]";
                        }
                        if (KPI == "SDCCH_Drop_Rate")
                        {
                            KPI_E = "[SDCCH_Drop_Rate]";
                            KPI_H = "[SDCCH_Drop_Rate]";
                            KPI_N = "[SDCCH_Drop_Rate]";
                        }
                        if (KPI == "TCH_Assign_FR")
                        {
                            KPI_E = "[TCH_Assign_Fail_Rate(NAK)(Eric_CELL)]";
                            KPI_H = "[TCH_Assignment_FR]";
                            KPI_N = "[TCH_Assignment_FR]";
                        }
                        if (KPI == "TCH_Congestion")
                        {
                            KPI_E = "[TCH_Congestion]";
                            KPI_H = "[TCH_Cong]";
                            KPI_N = "[TCH_Cong_Rate]";
                        }
                        if (KPI == "TCH_Traffic")
                        {
                            KPI_E = "[TCH_Traffic]";
                            KPI_H = "[TCH_Traffic]";
                            KPI_N = "[TCH_Traffic]";
                        }
                        if (KPI == "TCH_Availability")
                        {
                            KPI_E = "[TCH_Availability]";
                            KPI_H = "[TCH_Availability]";
                            KPI_N = "[TCH_Availability]";
                        }
                    }
                    if (Technology == "3G_CS")
                    {
                        Traffic_E = "[CS_Traffic]"; Ave_E = "[Cell_Availability_Rate_Exclude_Blocking(UCELL_Eric)]";
                        Traffic_H = "[CS_Erlang]"; Ave_H = "[Radio_Network_Availability_Ratio(Hu_Cell)]";
                        Traffic_N = "[CS_Traffic]"; Ave_N = "[Cell_Availability_excluding_blocked_by_user_state]";
                        if (KPI == "CS_RAB_Establish")
                        {
                            KPI_E = "[Cs_RAB_Establish_Success_Rate]";
                            KPI_H = "[CS_RAB_Setup_Success_Ratio]";
                            KPI_N = "[CS_RAB_Establish_Success_Rate]";
                        }
                        if (KPI == "CS_IRAT_HO_SR")
                        {
                            KPI_E = "[IRAT_HO_Voice_Suc_Rate]";
                            KPI_H = "[CS_IRAT_HO_SR]";
                            KPI_N = "[Inter_sys_RT_Hard_HO_SR_3Gto2G(CELL_nokia)]";
                        }
                        if (KPI == "CS_Drop_Rate")
                        {
                            KPI_E = "[CS_Drop_Call_Rate]";
                            KPI_H = "[AMR_Call_Drop_Ratio_New(Hu_CELL)]";
                            KPI_N = "[CS_Drop_Call_Rate]";
                        }
                        if (KPI == "Soft_HO_SR")
                        {
                            KPI_E = "[Soft_HO_Suc_Rate]";
                            KPI_H = "[Softer_Handover_Success_Ratio(Hu_Cell)]";
                            KPI_N = "[Soft_HO_Success_rate_RT]";
                        }
                        if (KPI == "CS_RRC_SR")
                        {
                            KPI_E = "[CS_RRC_Setup_Success_Rate]";
                            KPI_H = "[CS_RRC_Connection_Establishment_SR]";
                            KPI_N = "[CS_RRC_SETUP_SR_WITHOUT_REPEAT(CELL_NOKIA)]";
                        }
                        if (KPI == "CS_MultiRAB_SR")
                        {
                            KPI_E = "[CS_Multi_RAB_Establish_Success_Rate(Without_Nas)(CELL_Eric)]";
                            KPI_H = "[CSPS_RAB_Setup_Success_Ratio]";
                            KPI_N = "[CSAMR+PS_MRAB_STP_SR]";
                        }
                        if (KPI == "Inter_Carrier_HO_SR")
                        {
                            KPI_E = "[Inter_Carrier_HO_Success_Rate(UCell_Eric)]";
                            KPI_H = "[Inter_Carrier_HO_Success_Rate]";
                            KPI_N = "[Inter_Carrier_HO_Success_Rate]";
                        }
                        if (KPI == "CS_Traffic")
                        {
                            KPI_E = "[CS_Traffic]";
                            KPI_H = "[CS_Erlang]";
                            KPI_N = "[CS_Traffic]";
                        }
                        if (KPI == "Cell_Availability")
                        {
                            KPI_E = "[Cell_Availability_Rate_Exclude_Blocking(UCELL_Eric)]";
                            KPI_H = "[Radio_Network_Availability_Ratio(Hu_Cell)]";
                            KPI_N = "[Cell_Availability_excluding_blocked_by_user_state]";
                        }
                    }
                    if (Technology == "3G_PS")
                    {
                        Traffic_E = "[PS_Volume(GB)(UCell_Eric)]"; Ave_E = "[Cell_Availability_Rate_Exclude_Blocking(UCELL_Eric)]";
                        Traffic_H = "[PAYLOAD]"; Ave_H = "[Radio_Network_Availability_Ratio(Hu_Cell)]";
                        Traffic_N = "[PS_Payload_Total(HS+R99)(Nokia_CELL)_GB]"; Ave_N = "[Cell_Availability_excluding_blocked_by_user_state(Nokia_UCell)]";
                        if (KPI == "HSDPA_SR")
                        {
                            KPI_E = "[HSDPA_RAB_Setup_Succ_Rate(UCell_Eric)]";
                            KPI_H = "[HSDPA_RAB_Setup_Success_Ratio(Hu_Cell)]";
                            KPI_N = "[HSDPA_setup_success_ratio_from_user_perspective(CELL_Nokia)]";
                        }
                        if (KPI == "HSUPA_SR")
                        {
                            KPI_E = "[HSUPA_Setup_Success_Rate(UCell_Eric)]";
                            KPI_H = "[HSUPA_RAB_Setup_Success_Ratio(Hu_Cell)]";
                            KPI_N = "[HSUPA_Setup_Success_Ratio_from_user_perspective(CELL)]";
                        }
                        if (KPI == "DL_User_THR (Mbps)")
                        {
                            KPI_E = "[HS_USER_Throughput_NET_PQ(Mbps)(UCell_Eric)]";
                            KPI_H = "[AVERAGE_HSDPA_USER_THROUGHPUT_DC+SC(Mbit/s)(CELL_HUAWEI)]";
                            KPI_N = "[AVERAGE_HSDPA_USER_THROUGHPUT_DC+SC(Mbit/s)(Nokia_CELL)]";
                        }
                        if (KPI == "UL_User_THR (Kbps)")
                        {
                            KPI_E = "[HSUPA_User_Throughput_MACe(Kbps)(UCell_Eric)]";
                            KPI_H = "[hsupa_uplink_throughput_in_V16(CELL_Hu)]";
                            KPI_N = "[Average_hsupa_throughput_MACe(nokia_cell)]";
                        }
                        if (KPI == "HSDAP_Drop_Rate")
                        {
                            KPI_E = "[HSDPA_Drop_Call_Rate(UCell_Eric)]";
                            KPI_H = "[HSDPA_cdr(%)_(Hu_Cell)_new]";
                            KPI_N = "[HSDPA_Call_Drop_Rate(Nokia_Cell)]";
                        }
                        if (KPI == "HSUAP_Drop_Rate")
                        {
                            KPI_E = "[HSUPA_Drop_Call_Rate(UCell_Eric)]";
                            KPI_H = "[HSUPA_CDR(%)_(Hu_Cell)_new]";
                            KPI_N = "[HSUPA_Call_Drop_Rate(Nokia_CELL)]";
                        }
                        if (KPI == "PS_RRC_SR")
                        {
                            KPI_E = "[PS_RRC_Setup_Success_Rate(UCell_Eric)]";
                            KPI_H = "[PS_RRC_Connection_success_Rate_repeatless(Hu_Cell)]";
                            KPI_N = "[PS_RRCSETUP_SR]";
                        }
                        if (KPI == "Ps_RAB_Establish")
                        {
                            KPI_E = "[Ps_RAB_Establish_Success_Rate]";
                            KPI_H = "[PS_RAB_Setup_Success_Ratio]";
                            KPI_N = "[PS_RAB_Setup_Success_Ratio]";
                        }
                        if (KPI == "PS_MultiRAB_Establish")
                        {
                            KPI_E = "[PS_Multi_RAB_Establish_Success_Rate(without_Nas)(UCELL_Eric)]";
                            KPI_H = "[CS+PS_RAB_Setup_Success_Ratio]";
                            KPI_N = "[CSAMR+PS_MRAB_stp_SR(Nokia_CELL)]";
                        }
                        if (KPI == "PS_Drop_Rate")
                        {
                            KPI_E = "[PS_Drop_Call_Rate(UCell_Eric)]";
                            KPI_H = "[PS_Call_Drop_Ratio]";
                            KPI_N = "[Packet_Session_Drop_Ratio_NOKIA(CELL_NOKIA)]";
                        }
                        if (KPI == "HSDPA_Cell_Change_SR")
                        {
                            KPI_E = "[HSDPA_Cell_Change_Succ_Rate(UCell_Eric)]";
                            KPI_H = "[HSDPA_Soft_HandOver_Success_Ratio]";
                            KPI_N = "[HSDPA_Cell_Change_SR(Nokia_CELL)]";
                        }
                        if (KPI == "HS_Share_Payload")
                        {
                            KPI_E = "[HS_share_PAYLOAD_Rate(UCell_Eric)]";
                            KPI_H = "[HS_share_PAYLOAD_%]";
                            KPI_N = "[HS_SHARE_PAYLOAD(Nokia_CELL)]";
                        }
                        if (KPI == "DL_Cell_THR (Mbps)")
                        {
                            KPI_E = "[HSDPA_Cell_Scheduled_Throughput(mbps)(UCell_Eric)]";
                            KPI_H = "[HSDPA_SCHEDULING_Cell_throughput(CELL_HUAWEI)]";
                            KPI_N = "[Active_HS-DSCH_cell_throughput_mbs(CELL_nokia)]";
                        }
                        if (KPI == "RSSI (dBm)")
                        {
                            KPI_E = "[uplink_average_RSSI_dbm_(Eric_UCELL)]";
                            KPI_H = "[Mean_RTWP(Cell_Hu)]";
                            KPI_N = "[average_RTWP_dbm(Nokia_Cell)]";
                        }
                        if (KPI == "Average CQI")
                        {
                            KPI_E = "[Avg_CQI(UCell_Eric)]";
                            KPI_H = "[CQI_new(Hu_Cell)]";
                            KPI_N = "[AVERAGE_CQI(cell_nokia)]";
                        }
                        if (KPI == "PS_Payload (GB)")
                        {
                            KPI_E = "[PS_Volume(GB)(UCell_Eric)]";
                            KPI_H = "[PAYLOAD]";
                            KPI_N = "[PS_Payload_Total(HS+R99)(Nokia_CELL)_GB]";
                        }
                        if (KPI == "Cell_Availability")
                        {
                            KPI_E = "[Cell_Availability_Rate_Exclude_Blocking(UCELL_Eric)]";
                            KPI_H = "[Radio_Network_Availability_Ratio(Hu_Cell)]";
                            KPI_N = "[Cell_Availability_excluding_blocked_by_user_state(Nokia_UCell)]";
                        }
                    }
                    if (Technology == "4G")
                    {
                        Traffic_E = "[Total_Volume(UL+DL)(GB)(eNodeB_Eric)]"; Ave_E = "[Cell_Availability_Rate_Exclude_Blocking(Cell_EricLTE)]";
                        Traffic_H = "[Total_Traffic_Volume(GB)]"; Ave_H = "[Cell_Availability_Rate_Exclude_Blocking(Cell_Hu)]";
                        Traffic_N = "[Total_Payload_GB(Nokia_LTE_CELL)]"; Ave_N = "[cell_availability_exclude_manual_blocking(Nokia_LTE_CELL)]";
                        if (KPI == "RRC_Connection_SR")
                        {
                            KPI_E = "[RRC_Estab_Success_Rate(ReAtt)(EUCell_Eric)]";
                            KPI_H = "[RRC_Connection_Setup_Success_Rate_service]";
                            KPI_N = "[RRC_Connection_Setup_Success_Ratio(Nokia_LTE_CELL)]";
                        }
                        if (KPI == "ERAB_SR_Initial")
                        {
                            KPI_E = "[Initial_ERAB_Estab_Success_Rate(eNodeB_Eric)]";
                            KPI_H = "[E-RAB_Setup_Success_Rate]";
                            KPI_N = "[Initial_E-RAB_Setup_Success_Ratio(Nokia_LTE_CELL)]";
                        }
                        if (KPI == "ERAB_SR_Added")
                        {
                            KPI_E = "[E-RAB_Setup_SR_incl_added_New(EUCell_Eric)]";
                            KPI_H = "[E-RAB_Setup_Success_Rate(Hu_Cell)]";
                            KPI_N = "[E-RAB_Setup_SR_incl_added(Nokia_LTE_CELL)]";
                        }
                        if (KPI == "DL_THR (Mbps)")
                        {
                            KPI_E = "[Average_UE_DL_Throughput(Mbps)(eNodeB_Eric)]";
                            KPI_H = "[Average_Downlink_User_Throughput(Mbit/s)]";
                            KPI_N = "[User_Throughput_DL_mbps(Nokia_LTE_CELL)]";
                        }
                        if (KPI == "UL_THR (Mbps)")
                        {
                            KPI_E = "[Average_UE_UL_Throughput(Mbps)(eNodeB_Eric)]";
                            KPI_H = "[Average_UPlink_User_Throughput(Mbit/s)]";
                            KPI_N = "[User_Throughput_UL_mbps(Nokia_LTE_CELL)]";
                        }
                        if (KPI == "ERAB_Drop_Rate")
                        {
                            KPI_E = "[E_RAB_Drop_Rate(eNodeB_Eric)]";
                            KPI_H = "[Call_Drop_Rate]";
                            KPI_N = "[E-RAB_Drop_Ratio_RAN_View(Nokia_LTE_CELL)]";
                        }
                        if (KPI == "S1_Signalling_SR")
                        {
                            KPI_E = "[S1Signal_Estab_Success_Rate(EUCell_Eric)]";
                            KPI_H = "[S1Signal_E-RAB_Setup_SR(Hu_Cell)]";
                            KPI_N = "[S1Signal_E-RAB_Setup_SR(Nokia_LTE_CELL)]";
                        }
                        if (KPI == "Intra_Freq_SR")
                        {
                            KPI_E = "[IntraF_Handover_Execution(eNodeB_Eric)]";
                            KPI_H = "[IntraF_HOOut_SR]";
                            KPI_N = "[HO_Success_Ratio_intra_eNB(Nokia_LTE_CELL)]";
                        }
                        if (KPI == "Inter_Freq_SR")
                        {
                            KPI_E = "[InterF_Handover_Execution(eNodeB_Eric)]";
                            KPI_H = "[InterF_HOOut_SR]";
                            KPI_N = "[Inter-Freq_HO_SR(Nokia_LTE_CELL)]";
                        }
                        if (KPI == "UL_Packet_Loss")
                        {
                            KPI_E = "[Average_UE_Ul_Packet_Loss_Rate(eNodeB_Eric)]";
                            KPI_H = "[Average_UL_Packet_Loss_%(Huawei_LTE_UCell)]";
                            KPI_N = "[Packet_loss_UL(Nokia_EUCELL)]";
                        }
                        if (KPI == "UE_DL_Latency (ms)")
                        {
                            KPI_E = "[Average_UE_DL_Latency(ms)(eNodeB_Eric)]";
                            KPI_H = "[Average_DL_Latency_ms(Huawei_LTE_EUCell)]";
                            KPI_N = "[Average_Latency_DL_ms(Nokia_LTE_CELL)]";
                        }
                        if (KPI == "Average CQI")
                        {
                            KPI_E = "[CQI_(EUCell_Eric)]";
                            KPI_H = "[Average_CQI(Huawei_LTE_Cell)]";
                            KPI_N = "[Average_CQI(Nokia_LTE_CELL)]";
                        }
                        if (KPI == "PUCCH_RSSI (dBm)")
                        {
                            KPI_E = "[RSSI_PUCCH]";
                            KPI_H = "[RSSI_PUCCH]";
                            KPI_N = "[RSSI_PUCCH]";
                        }
                        if (KPI == "PUSCH_RSSI (dBm)")
                        {
                            KPI_E = "[RSSI_PUSCH]";
                            KPI_H = "[RSSI_PUSCH]";
                            KPI_N = "[RSSI_PUSCH]";
                        }
                        if (KPI == "Total_Paylaod (GB)")
                        {
                            KPI_E = "[Total_Volume(UL+DL)(GB)(eNodeB_Eric)]";
                            KPI_H = "[Total_Traffic_Volume(GB)]";
                            KPI_N = "[Total_Payload_GB(Nokia_LTE_CELL)]";
                        }
                        if (KPI == "Cell_Availability")
                        {
                            KPI_E = "[Cell_Availability_Rate_Exclude_Blocking(Cell_EricLTE)]";
                            KPI_H = "[Cell_Availability_Rate_Exclude_Blocking(Cell_Hu)]";
                            KPI_N = "[cell_availability_exclude_manual_blocking(Nokia_LTE_CELL)]";
                        }
                    }
                    string s1_E = "";
                    string s2_E = "";
                    string s1_H = "";
                    string s2_H = "";
                    string s1_N = "";
                    string s2_N = "";
                    if (textBox1.Text != "" && textBox2.Text != "" && textBox3.Text != "")
                    {
                        s2_E = KPI_E + sign + textBox3.Text + " and " + Ave_E + ">=" + textBox1.Text + " and " + Traffic_E + ">=" + textBox2.Text;
                        s2_H = KPI_H + sign + textBox3.Text + " and " + Ave_H + ">=" + textBox1.Text + " and " + Traffic_H + ">=" + textBox2.Text;
                        s2_N = KPI_N + sign + textBox3.Text + " and " + Ave_N + ">=" + textBox1.Text + " and " + Traffic_N + ">=" + textBox2.Text;
                    }
                    if (textBox1.Text == "" && textBox2.Text != "" && textBox3.Text != "")
                    {
                        s2_E = KPI_E + sign + textBox3.Text + " and " + Traffic_E + ">=" + textBox2.Text;
                        s2_H = KPI_H + sign + textBox3.Text + " and " + Traffic_H + ">=" + textBox2.Text;
                        s2_N = KPI_N + sign + textBox3.Text + " and " + Traffic_N + ">=" + textBox2.Text;
                    }
                    if (textBox1.Text != "" && textBox2.Text == "" && textBox3.Text != "")
                    {
                        s2_E = KPI_E + sign + textBox3.Text + " and " + Ave_E + ">=" + textBox1.Text;
                        s2_H = KPI_H + sign + textBox3.Text + " and " + Ave_H + ">=" + textBox1.Text;
                        s2_N = KPI_N + sign + textBox3.Text + " and " + Ave_N + ">=" + textBox1.Text;
                    }
                    if (textBox1.Text == "" && textBox2.Text == "" && textBox3.Text != "")
                    {
                        s2_E = KPI_E + sign + textBox3.Text;
                        s2_H = KPI_H + sign + textBox3.Text;
                        s2_N = KPI_N + sign + textBox3.Text;

                    }
                    if (textBox3.Text == "")
                    {
                        MessageBox.Show("Please input the KPI Threshold");
                    }
                    if (textBox3.Text != "")
                    {
                        if (Technology == "2G")
                        {
                            s1_E = KPI_E + "as'" + KPI + "' , " + Traffic_E + "as " + "'Traffic(Erlang)'" + " , " + Ave_E + "as " + "'Availability'" + " from [CC2_Ericsson_Cell_Daily]";
                            s1_H = KPI_H + "as'" + KPI + "' , " + Traffic_H + "as " + "'Traffic(Erlang)'" + " , " + Ave_H + "as " + "'Availability'" + " from [CC2_Huawei_Cell_Daily]";
                            s1_N = KPI_N + "as'" + KPI + "' , " + Traffic_N + "as " + "'Traffic(Erlang)'" + " , " + Ave_N + "as " + "'Availability'" + " from [CC2_Nokia_Cell_Daily]";
                        }
                        if (Technology == "3G_CS")
                        {
                            s1_E = KPI_E + "as'" + KPI + "' , " + Traffic_E + "as " + "'Traffic(Erlang)'" + " , " + Ave_E + "as " + "'Availability'" + " from [CC3_Ericsson_Cell_Daily]";
                            s1_H = KPI_H + "as'" + KPI + "' , " + Traffic_H + "as " + "'Traffic(Erlang)'" + " , " + Ave_H + "as " + "'Availability'" + " from [CC3_Huawei_Cell_Daily]";
                            s1_N = KPI_N + "as'" + KPI + "' , " + Traffic_N + "as " + "'Traffic(Erlang)'" + " , " + Ave_N + "as " + "'Availability'" + " from [CC3_Nokia_Cell_Daily]";
                        }
                        if (Technology == "3G_PS")
                        {
                            s1_E = KPI_E + "as'" + KPI + "' , " + Traffic_E + "as " + "'Paylaod(GB)'" + " , " + Ave_E + "as " + "'Availability'" + " from [RD3_Ericsson_Cell_Daily]";
                            s1_H = KPI_H + "as'" + KPI + "' , " + Traffic_H + "as " + "'Paylaod(GB)'" + " , " + Ave_H + "as " + "'Availability'" + " from [RD3_Huawei_Cell_Daily]";
                            s1_N = KPI_N + "as'" + KPI + "' , " + Traffic_N + "as " + "'Paylaod(GB)'" + " , " + Ave_N + "as " + "'Availability'" + " from [RD3_Nokia_Cell_Daily]";
                        }
                        if (Technology == "4G")
                        {
                            s1_E = KPI_E + "as'" + KPI + "' , " + Traffic_E + "as " + "'Paylaod(GB)'" + " , " + Ave_E + "as " + "'Availability'" + " from [TBL_LTE_CELL_Daily_E]";
                            s1_H = KPI_H + "as'" + KPI + "' , " + Traffic_H + "as " + "'Paylaod(GB)'" + " , " + Ave_H + "as " + "'Availability'" + " from [TBL_LTE_CELL_Daily_H]";
                            s1_N = KPI_N + "as'" + KPI + "' , " + Traffic_N + "as " + "'Paylaod(GB)'" + " , " + Ave_N + "as " + "'Availability'" + " from [TBL_LTE_CELL_Daily_N]";
                        }
                    }

                    if (textBox3.Text != "")
                    {
                        string KPI_Q = "";
                        if (Technology == "2G")
                        {
                            KPI_Q = "select  'Ericsson' as 'Vendor' , BSC, Cell, Province, Date, " + s1_E + " where Date = '" + Date_List + "' and " + s2_E + " and (" + Node_Index_List + ")" +
                           " union all " +
                           "select  'Huawei' as 'Vendor', BSC, Cell, Province, Date, " + s1_H + " where Date = '" + Date_List + "' and " + s2_H + " and (" + Node_Index_List + ")" +
                           " union all " +
                             "select 'Nokia' as 'Vendor', BSC, Seg, Province, Date, " + s1_N + " where Date = '" + Date_List + "' and " + s2_N + " and (" + Node_Index_List + ")";
                        }
                        if (Technology == "3G_CS")
                        {
                            KPI_Q = "select  'Ericsson' as 'Vendor' , ElementID as 'RNC', ElementID1 as 'Cell',  Date, " + s1_E + " where Date = '" + Date_List + "' and " + s2_E + " and (" + Node_Index_List + ")" +
                           " union all " +
                           "select  'Huawei' as 'Vendor',  ElementID as 'RNC', ElementID1 as 'Cell',  Date, " + s1_H + " where Date = '" + Date_List + "' and " + s2_H + " and (" + Node_Index_List + ")" +
                           " union all " +
                             "select 'Nokia' as 'Vendor', ElementID as 'RNC', ElementID1 as 'Cell',  Date, " + s1_N + " where Date = '" + Date_List + "' and " + s2_N + " and (" + Node_Index_List + ")";
                        }
                        if (Technology == "3G_PS")
                        {
                            KPI_Q = "select  'Ericsson' as 'Vendor' , ElementID as 'RNC', ElementID1 as 'Cell',  Date, " + s1_E + " where Date = '" + Date_List + "' and " + s2_E + " and (" + Node_Index_List + ")" +
                           " union all " +
                           "select  'Huawei' as 'Vendor',  ElementID as 'RNC', ElementID1 as 'Cell',  Date, " + s1_H + " where Date = '" + Date_List + "' and " + s2_H + " and (" + Node_Index_List + ")" +
                           " union all " +
                             "select 'Nokia' as 'Vendor', ElementID as 'RNC', ElementID1 as 'Cell',  Date, " + s1_N + " where Date = '" + Date_List + "' and " + s2_N + " and (" + Node_Index_List + ")";
                        }
                        if (Technology == "4G")
                        {
                            KPI_Q = "select  'Ericsson' as 'Vendor' , Province, eNodeB as 'Cell',  Datetime, " + s1_E + " where Datetime = '" + Date_List + "' and " + s2_E + " and (" + Node_Index_List + ")" +
                           " union all " +
                           "select  'Huawei' as 'Vendor',  Province, eNodeB as 'Cell',  Datetime, " + s1_H + " where Datetime = '" + Date_List + "' and " + s2_H + " and (" + Node_Index_List + ")" +
                           " union all " +
                             "select 'Nokia' as 'Vendor', Province, ElementID1 as 'Cell',  Date, " + s1_N + " where Date = '" + Date_List + "' and " + s2_N + " and (" + Node_Index_List + ")";
                        }
                        SqlCommand Node_List_Quary1 = new SqlCommand(KPI_Q, connection);
                        Node_List_Quary1.ExecuteNonQuery();
                        DataTable Table_Node_List_Quary = new DataTable();
                        SqlDataAdapter dataAdapter_Node_List_Quary = new SqlDataAdapter(Node_List_Quary1);
                        dataAdapter_Node_List_Quary.Fill(Table_Node_List_Quary);





                        XLWorkbook wb = new XLWorkbook();
                        wb.Worksheets.Add(Table_Node_List_Quary, "WPC");

                        var saveFileDialog = new SaveFileDialog
                        {
                            FileName = "WPC_" + Interval + "_" + Technology + "_" + KPI,
                            Filter = "Excel files|*.xlsx",
                            Title = "Save an Excel File"
                        };

                        saveFileDialog.ShowDialog();

                        if (!String.IsNullOrWhiteSpace(saveFileDialog.FileName))
                            wb.SaveAs(saveFileDialog.FileName);

                        label15.Text = "Finished";
                        label15.BackColor = Color.Green;

                        MessageBox.Show("Finished");





                    }
                }



                if (Interval == "BH")
                {
                    string KPI_E = "";
                    string Traffic_E = "";
                    string Ave_E = "";
                    string KPI_H = "";
                    string Traffic_H = "";
                    string Ave_H = "";
                    string KPI_N = "";
                    string Traffic_N = "";
                    string Ave_N = "";
                    if (Technology == "2G")
                    {
                        Traffic_E = "[TCH_Traffic_BH]"; Ave_E = "[TCH_Availability]";
                        Traffic_H = "[TCH_Traffic_BH]"; Ave_H = "[TCH_Availability]";
                        Traffic_N = "[TCH_Traffic_BH]"; Ave_N = "[TCH_Availability]";
                        if (KPI == "CSSR")
                        {
                            KPI_E = "[CSSR_MCI]";
                            KPI_H = "[CSSR3]";
                            KPI_N = "[CSSR_MCI]";
                        }
                        if (KPI == "IHSR")
                        {
                            KPI_E = "[IHSR]";
                            KPI_H = "[IHSR2]";
                            KPI_N = "[IHSR]";
                        }
                        if (KPI == "OHSR")
                        {
                            KPI_E = "[OHSR]";
                            KPI_H = "[OHSR2]";
                            KPI_N = "[OHSR]";
                        }
                        if (KPI == "CDR")
                        {
                            KPI_E = "[CDR(not Affected by incoming Handovers from 3G)(Eric_CELL)]";
                            KPI_H = "[CDR3]";
                            KPI_N = "[CDR(including_CS_IRAT_handovers_3G_to2G)(Nokia_SEG)]";
                        }
                        if (KPI == "RxQual_DL")
                        {
                            KPI_E = "[RxQual_DL]";
                            KPI_H = "[RX_QUALITTY_DL_NEW]";
                            KPI_N = "[RxQuality_DL]";
                        }
                        if (KPI == "RxQual_UL")
                        {
                            KPI_E = "[RxQual_UL]";
                            KPI_H = "[RX_QUALITTY_UL_NEW]";
                            KPI_N = "[RxQuality_UL]";
                        }
                        if (KPI == "SDCCH_Access_SR")
                        {
                            KPI_E = "[SDCCH_Access_Succ_Rate]";
                            KPI_H = "[SDCCH_Access_Success_Rate2]";
                            KPI_N = "[SDCCH_Access_Success_Rate]";
                        }
                        if (KPI == "SDCCH_Congestion")
                        {
                            KPI_E = "[SDCCH_Congestion]";
                            KPI_H = "[SDCCH_Congestion_Rate]";
                            KPI_N = "[SDCCH_Congestion_Rate]";
                        }
                        if (KPI == "SDCCH_Drop_Rate")
                        {
                            KPI_E = "[SDCCH_Drop_Rate]";
                            KPI_H = "[SDCCH_Drop_Rate]";
                            KPI_N = "[SDCCH_Drop_Rate]";
                        }
                        if (KPI == "TCH_Assign_FR")
                        {
                            KPI_E = "[TCH_Assign_Fail_Rate(NAK)(Eric_CELL)]";
                            KPI_H = "[TCH_Assignment_FR]";
                            KPI_N = "[TCH_Assignment_FR]";
                        }
                        if (KPI == "TCH_Congestion")
                        {
                            KPI_E = "[TCH_Congestion]";
                            KPI_H = "[TCH_Cong]";
                            KPI_N = "[TCH_Cong_Rate]";
                        }
                        if (KPI == "TCH_Traffic")
                        {
                            KPI_E = "[TCH_Traffic_BH]";
                            KPI_H = "[TCH_Traffic_BH]";
                            KPI_N = "[TCH_Traffic_BH]";
                        }
                        if (KPI == "TCH_Availability")
                        {
                            KPI_E = "[TCH_Availability]";
                            KPI_H = "[TCH_Availability]";
                            KPI_N = "[TCH_Availability]";
                        }
                    }
                    if (Technology == "3G_CS")
                    {
                        Traffic_E = "[CS_Traffic_BH]"; Ave_E = "[Cell_Availability_Rate_Exclude_Blocking(UCELL_Eric)]";
                        Traffic_H = "[CS_Erlang]"; Ave_H = "[Radio_Network_Availability_Ratio(Hu_Cell)]";
                        Traffic_N = "[CS_TrafficBH]"; Ave_N = "[Cell_Availability_excluding_blocked_by_user_state]";
                        if (KPI == "CS_RAB_Establish")
                        {
                            KPI_E = "[Cs_RAB_Establish_Success_Rate]";
                            KPI_H = "[CS_RAB_Setup_Success_Ratio]";
                            KPI_N = "[CS_RAB_Establish_Success_Rate]";
                        }
                        if (KPI == "CS_IRAT_HO_SR")
                        {
                            KPI_E = "[IRAT_HO_Voice_Suc_Rate]";
                            KPI_H = "[CS_IRAT_HO_SR]";
                            KPI_N = "[Inter_sys_RT_Hard_HO_SR_3Gto2G(CELL_nokia)]";
                        }
                        if (KPI == "CS_Drop_Rate")
                        {
                            KPI_E = "[CS_Drop_Call_Rate]";
                            KPI_H = "[AMR_Call_Drop_Ratio_New(Hu_CELL)]";
                            KPI_N = "[CS_Drop_Call_Rate]";
                        }
                        if (KPI == "Soft_HO_SR")
                        {
                            KPI_E = "[Soft_HO_Suc_Rate]";
                            KPI_H = "[Softer_Handover_Success_Ratio(Hu_Cell)]";
                            KPI_N = "[Soft_HO_Success_rate_RT]";
                        }
                        if (KPI == "CS_RRC_SR")
                        {
                            KPI_E = "[CS_RRC_Setup_Success_Rate]";
                            KPI_H = "[CS_RRC_Connection_Establishment_SR]";
                            KPI_N = "[CS_RRC_SETUP_SR_WITHOUT_REPEAT(CELL_NOKIA)]";
                        }
                        if (KPI == "CS_MultiRAB_SR")
                        {
                            KPI_E = "[CS_Multi_RAB_Establish_Success_Rate(Without_Nas)(CELL_Eric)]";
                            KPI_H = "[CSPS_RAB_Setup_Success_Ratio]";
                            KPI_N = "[CSAMR+PS_MRAB_STP_SR]";
                        }
                        if (KPI == "Inter_Carrier_HO_SR")
                        {
                            KPI_E = "[Inter_Carrier_HO_Success_Rate(UCell_Eric)]";
                            KPI_H = "[Inter_Carrier_HO_Success_Rate]";
                            KPI_N = "[Inter_Carrier_HO_Success_Rate]";
                        }
                        if (KPI == "CS_Traffic")
                        {
                            KPI_E = "[CS_Traffic]";
                            KPI_H = "[CS_Erlang]";
                            KPI_N = "[CS_Traffic]";
                        }
                        if (KPI == "Cell_Availability")
                        {
                            KPI_E = "[Cell_Availability_Rate_Exclude_Blocking(UCELL_Eric)]";
                            KPI_H = "[Radio_Network_Availability_Ratio(Hu_Cell)]";
                            KPI_N = "[Cell_Availability_excluding_blocked_by_user_state]";
                        }
                    }
                    if (Technology == "3G_PS")
                    {
                        Traffic_E = "[Payload_Total_BH]"; Ave_E = "[Cell_Availability_Rate_Exclude_Blocking(UCELL_Eric)]";
                        Traffic_H = "[Payload_Total_BH]"; Ave_H = "[Radio_Network_Availability_Ratio(Hu_Cell)]";
                        Traffic_N = "[Payload_Total_BH]"; Ave_N = "[Cell_Availability_excluding_blocked_by_user_state(Nokia_UCell)]";
                        if (KPI == "HSDPA_SR")
                        {
                            KPI_E = "[HSDPA_RAB_Setup_Succ_Rate(UCell_Eric)]";
                            KPI_H = "[HSDPA_RAB_Setup_Success_Ratio(Hu_Cell)]";
                            KPI_N = "[HSDPA_setup_success_ratio_from_user_perspective(CELL_Nokia)]";
                        }
                        if (KPI == "HSUPA_SR")
                        {
                            KPI_E = "[HSUPA_Setup_Success_Rate(UCell_Eric)]";
                            KPI_H = "[HSUPA_RAB_Setup_Success_Ratio(Hu_Cell)]";
                            KPI_N = "[HSUPA_Setup_Success_Ratio_from_user_perspective(CELL)]";
                        }
                        if (KPI == "DL_User_THR (Mbps)")
                        {
                            KPI_E = "[HS_USER_Throughput_NET_PQ(Mbps)(UCell_Eric)]";
                            KPI_H = "[AVERAGE_HSDPA_USER_THROUGHPUT_DC+SC(Mbit/s)(CELL_HUAWEI)]";
                            KPI_N = "[AVERAGE_HSDPA_USER_THROUGHPUT_DC+SC(Mbit/s)(Nokia_CELL)]";
                        }
                        if (KPI == "UL_User_THR (Kbps)")
                        {
                            KPI_E = "[HSUPA_User_Throughput_MACe(Kbps)(UCell_Eric)]";
                            KPI_H = "[hsupa_uplink_throughput_in_V16(CELL_Hu)]";
                            KPI_N = "[Average_hsupa_throughput_MACe(nokia_cell)]";
                        }
                        if (KPI == "HSDAP_Drop_Rate")
                        {
                            KPI_E = "[HSDPA_Drop_Call_Rate(UCell_Eric)]";
                            KPI_H = "[HSDPA_cdr(%)_(Hu_Cell)_new]";
                            KPI_N = "[HSDPA_Call_Drop_Rate(Nokia_Cell)]";
                        }
                        if (KPI == "HSUAP_Drop_Rate")
                        {
                            KPI_E = "[HSUPA_Drop_Call_Rate(UCell_Eric)]";
                            KPI_H = "[HSUPA_CDR(%)_(Hu_Cell)_new]";
                            KPI_N = "[HSUPA_Call_Drop_Rate(Nokia_CELL)]";
                        }
                        if (KPI == "PS_RRC_SR")
                        {
                            KPI_E = "[PS_RRC_Setup_Success_Rate(UCell_Eric)]";
                            KPI_H = "[PS_RRC_Connection_success_Rate_repeatless(Hu_Cell)]";
                            KPI_N = "[PS_RRCSETUP_SR]";
                        }
                        if (KPI == "Ps_RAB_Establish")
                        {
                            KPI_E = "[Ps_RAB_Establish_Success_Rate]";
                            KPI_H = "[PS_RAB_Setup_Success_Ratio]";
                            KPI_N = "[PS_RAB_Setup_Success_Ratio]";
                        }
                        if (KPI == "PS_MultiRAB_Establish")
                        {
                            KPI_E = "[PS_Multi_RAB_Establish_Success_Rate(without_Nas)(UCELL_Eric)]";
                            KPI_H = "[CS+PS_RAB_Setup_Success_Ratio]";
                            KPI_N = "[CSAMR+PS_MRAB_stp_SR(Nokia_CELL)]";
                        }
                        if (KPI == "PS_Drop_Rate")
                        {
                            KPI_E = "[PS_Drop_Call_Rate(UCell_Eric)]";
                            KPI_H = "[PS_Call_Drop_Ratio]";
                            KPI_N = "[Packet_Session_Drop_Ratio_NOKIA(CELL_NOKIA)]";
                        }
                        if (KPI == "HSDPA_Cell_Change_SR")
                        {
                            KPI_E = "[HSDPA_Cell_Change_Succ_Rate(UCell_Eric)]";
                            KPI_H = "[HSDPA_Soft_HandOver_Success_Ratio]";
                            KPI_N = "[HSDPA_Cell_Change_SR(Nokia_CELL)]";
                        }
                        if (KPI == "HS_Share_Payload")
                        {
                            KPI_E = "[HS_share_PAYLOAD_Rate(UCell_Eric)]";
                            KPI_H = "[HS_share_PAYLOAD_%]";
                            KPI_N = "[HS_SHARE_PAYLOAD(Nokia_CELL)]";
                        }
                        if (KPI == "DL_Cell_THR (Mbps)")
                        {
                            KPI_E = "[HSDPA_Cell_Scheduled_Throughput(mbps)(UCell_Eric)]";
                            KPI_H = "[HSDPA_SCHEDULING_Cell_throughput(CELL_HUAWEI)]";
                            KPI_N = "[Active_HS-DSCH_cell_throughput_mbs(CELL_nokia)]";
                        }
                        if (KPI == "RSSI (dBm)")
                        {
                            KPI_E = "[RTWP]";
                            KPI_H = "[RTWP]";
                            KPI_N = "[RTWP]";
                        }
                        if (KPI == "Average CQI")
                        {
                            KPI_E = "[AVERAGE_CQI]";
                            KPI_H = "[AVERAGE_CQI]";
                            KPI_N = "[AVERAGE_CQI]";
                        }
                        if (KPI == "PS_Payload (GB)")
                        {
                            KPI_E = "[Payload_Total_BH]";
                            KPI_H = "[Payload_Total_BH]";
                            KPI_N = "[Payload_Total_BH]";
                        }
                        if (KPI == "Cell_Availability")
                        {
                            KPI_E = "[Cell_Availability_Rate_Exclude_Blocking(UCELL_Eric)]";
                            KPI_H = "[Radio_Network_Availability_Ratio(Hu_Cell)]";
                            KPI_N = "[Cell_Availability_excluding_blocked_by_user_state(Nokia_UCell)]";
                        }
                    }
                    if (Technology == "4G")
                    {
                        Traffic_E = "[Total_Volume(UL+DL)(GB)(eNodeB_Eric)]"; Ave_E = "[Cell_Availability_Rate_Exclude_Blocking(Cell_EricLTE)]";
                        Traffic_H = "[Total_Traffic_Volume(GB)]"; Ave_H = "[Cell_Availability_Rate_Exclude_Blocking(Cell_Hu)]";
                        Traffic_N = "[Total_Payload_GB(Nokia_LTE_CELL)]"; Ave_N = "[cell_availability_exclude_manual_blocking(Nokia_LTE_CELL)]";
                        if (KPI == "RRC_Connection_SR")
                        {
                            KPI_E = "[RRC_Estab_Success_Rate(ReAtt)(EUCell_Eric)]";
                            KPI_H = "[RRC_Connection_Setup_Success_Rate_service]";
                            KPI_N = "[RRC_Connection_Setup_Success_Ratio(Nokia_LTE_CELL)]";
                        }
                        if (KPI == "ERAB_SR_Initial")
                        {
                            KPI_E = "[Initial_ERAB_Estab_Success_Rate(eNodeB_Eric)]";
                            KPI_H = "[E-RAB_Setup_Success_Rate]";
                            KPI_N = "[Initial_E-RAB_Setup_Success_Ratio(Nokia_LTE_CELL)]";
                        }
                        if (KPI == "ERAB_SR_Added")
                        {
                            KPI_E = "[E-RAB_Setup_SR_incl_added_New(EUCell_Eric)]";
                            KPI_H = "[E-RAB_Setup_Success_Rate(Hu_Cell)]";
                            KPI_N = "[E-RAB_Setup_SR_incl_added(Nokia_LTE_CELL)]";
                        }
                        if (KPI == "DL_THR (Mbps)")
                        {
                            KPI_E = "[Average_UE_DL_Throughput(Mbps)(eNodeB_Eric)]";
                            KPI_H = "[Average_Downlink_User_Throughput(Mbit/s)]";
                            KPI_N = "[User_Throughput_DL_mbps(Nokia_LTE_CELL)]";
                        }
                        if (KPI == "UL_THR (Mbps)")
                        {
                            KPI_E = "[Average_UE_UL_Throughput(Mbps)(eNodeB_Eric)]";
                            KPI_H = "[Average_UPlink_User_Throughput(Mbit/s)]";
                            KPI_N = "[User_Throughput_UL_mbps(Nokia_LTE_CELL)]";
                        }
                        if (KPI == "ERAB_Drop_Rate")
                        {
                            KPI_E = "[E_RAB_Drop_Rate(eNodeB_Eric)]";
                            KPI_H = "[Call_Drop_Rate]";
                            KPI_N = "[E-RAB_Drop_Ratio_RAN_View(Nokia_LTE_CELL)]";
                        }
                        if (KPI == "S1_Signalling_SR")
                        {
                            KPI_E = "[S1Signal_Estab_Success_Rate(EUCell_Eric)]";
                            KPI_H = "[S1Signal_E-RAB_Setup_SR(Hu_Cell)]";
                            KPI_N = "[S1Signal_E-RAB_Setup_SR(Nokia_LTE_CELL)]";
                        }
                        if (KPI == "Intra_Freq_SR")
                        {
                            KPI_E = "[IntraF_Handover_Execution(eNodeB_Eric)]";
                            KPI_H = "[IntraF_HOOut_SR]";
                            KPI_N = "[HO_Success_Ratio_intra_eNB(Nokia_LTE_CELL)]";
                        }
                        if (KPI == "Inter_Freq_SR")
                        {
                            KPI_E = "[InterF_Handover_Execution(eNodeB_Eric)]";
                            KPI_H = "[InterF_HOOut_SR]";
                            KPI_N = "[Inter-Freq_HO_SR(Nokia_LTE_CELL)]";
                        }
                        if (KPI == "UL_Packet_Loss")
                        {
                            KPI_E = "[Average_UE_Ul_Packet_Loss_Rate(eNodeB_Eric)]";
                            KPI_H = "[Average_UL_Packet_Loss_%(Huawei_LTE_UCell)]";
                            KPI_N = "[Packet_loss_UL(Nokia_EUCELL)]";
                        }
                        if (KPI == "UE_DL_Latency (ms)")
                        {
                            KPI_E = "[Average_UE_DL_Latency(ms)(eNodeB_Eric)]";
                            KPI_H = "[Average_DL_Latency_ms(Huawei_LTE_EUCell)]";
                            KPI_N = "[Average_Latency_DL_ms(Nokia_LTE_CELL)]";
                        }
                        if (KPI == "Average CQI")
                        {
                            KPI_E = "[CQI_(EUCell_Eric)]";
                            KPI_H = "[Average_CQI(Huawei_LTE_Cell)]";
                            KPI_N = "[Average_CQI(Nokia_LTE_CELL)]";
                        }
                        if (KPI == "PUCCH_RSSI (dBm)")
                        {
                            KPI_E = "[RSSI_PUCCH]";
                            KPI_H = "[RSSI_PUCCH]";
                            KPI_N = "[RSSI_PUCCH]";
                        }
                        if (KPI == "PUSCH_RSSI (dBm)")
                        {
                            KPI_E = "[RSSI_PUSCH]";
                            KPI_H = "[RSSI_PUSCH]";
                            KPI_N = "[RSSI_PUSCH]";
                        }
                        if (KPI == "Total_Paylaod (GB)")
                        {
                            KPI_E = "[Total_Volume(UL+DL)(GB)(eNodeB_Eric)]";
                            KPI_H = "[Total_Traffic_Volume(GB)]";
                            KPI_N = "[Total_Payload_GB(Nokia_LTE_CELL)]";
                        }
                        if (KPI == "Cell_Availability")
                        {
                            KPI_E = "[Cell_Availability_Rate_Exclude_Blocking(Cell_EricLTE)]";
                            KPI_H = "[Cell_Availability_Rate_Exclude_Blocking(Cell_Hu)]";
                            KPI_N = "[cell_availability_exclude_manual_blocking(Nokia_LTE_CELL)]";
                        }
                    }
                    string s1_E = "";
                    string s2_E = "";
                    string s1_H = "";
                    string s2_H = "";
                    string s1_N = "";
                    string s2_N = "";
                    if (textBox1.Text != "" && textBox2.Text != "" && textBox3.Text != "")
                    {
                        s2_E = KPI_E + sign + textBox3.Text + " and " + Ave_E + ">=" + textBox1.Text + " and " + Traffic_E + ">=" + textBox2.Text;
                        s2_H = KPI_H + sign + textBox3.Text + " and " + Ave_H + ">=" + textBox1.Text + " and " + Traffic_H + ">=" + textBox2.Text;
                        s2_N = KPI_N + sign + textBox3.Text + " and " + Ave_N + ">=" + textBox1.Text + " and " + Traffic_N + ">=" + textBox2.Text;
                    }
                    if (textBox1.Text == "" && textBox2.Text != "" && textBox3.Text != "")
                    {
                        s2_E = KPI_E + sign + textBox3.Text + " and " + Traffic_E + ">=" + textBox2.Text;
                        s2_H = KPI_H + sign + textBox3.Text + " and " + Traffic_H + ">=" + textBox2.Text;
                        s2_N = KPI_N + sign + textBox3.Text + " and " + Traffic_N + ">=" + textBox2.Text;
                    }
                    if (textBox1.Text != "" && textBox2.Text == "" && textBox3.Text != "")
                    {
                        s2_E = KPI_E + sign + textBox3.Text + " and " + Ave_E + ">=" + textBox1.Text;
                        s2_H = KPI_H + sign + textBox3.Text + " and " + Ave_H + ">=" + textBox1.Text;
                        s2_N = KPI_N + sign + textBox3.Text + " and " + Ave_N + ">=" + textBox1.Text;
                    }
                    if (textBox1.Text == "" && textBox2.Text == "" && textBox3.Text != "")
                    {
                        s2_E = KPI_E + sign + textBox3.Text;
                        s2_H = KPI_H + sign + textBox3.Text;
                        s2_N = KPI_N + sign + textBox3.Text;

                    }
                    if (textBox3.Text == "")
                    {
                        MessageBox.Show("Please input the KPI Threshold");
                    }
                    if (textBox3.Text != "")
                    {
                        if (Technology == "2G")
                        {
                            s1_E = KPI_E + "as'" + KPI + "' , " + Traffic_E + "as " + "'Traffic(Erlang)'" + " , " + Ave_E + "as " + "'Availability'" + " from [CC2_Ericsson_Cell_BH]";
                            s1_H = KPI_H + "as'" + KPI + "' , " + Traffic_H + "as " + "'Traffic(Erlang)'" + " , " + Ave_H + "as " + "'Availability'" + " from [CC2_Huawei_Cell_BH]";
                            s1_N = KPI_N + "as'" + KPI + "' , " + Traffic_N + "as " + "'Traffic(Erlang)'" + " , " + Ave_N + "as " + "'Availability'" + " from [CC2_Nokia_Cell_BH]";
                        }
                        if (Technology == "3G_CS")
                        {
                            s1_E = KPI_E + "as'" + KPI + "' , " + Traffic_E + "as " + "'Traffic(Erlang)'" + " , " + Ave_E + "as " + "'Availability'" + " from [CC3_Ericsson_Cell_BH]";
                            s1_H = KPI_H + "as'" + KPI + "' , " + Traffic_H + "as " + "'Traffic(Erlang)'" + " , " + Ave_H + "as " + "'Availability'" + " from [CC3_Huawei_Cell_BH]";
                            s1_N = KPI_N + "as'" + KPI + "' , " + Traffic_N + "as " + "'Traffic(Erlang)'" + " , " + Ave_N + "as " + "'Availability'" + " from [CC3_Nokia_Cell_BH]";
                        }
                        if (Technology == "3G_PS")
                        {
                            s1_E = KPI_E + "as'" + KPI + "' , " + Traffic_E + "as " + "'Paylaod(GB)'" + " , " + Ave_E + "as " + "'Availability'" + " from [RD3_Ericsson_Cell_BH]";
                            s1_H = KPI_H + "as'" + KPI + "' , " + Traffic_H + "as " + "'Paylaod(GB)'" + " , " + Ave_H + "as " + "'Availability'" + " from [RD3_Huawei_Cell_BH]";
                            s1_N = KPI_N + "as'" + KPI + "' , " + Traffic_N + "as " + "'Paylaod(GB)'" + " , " + Ave_N + "as " + "'Availability'" + " from [RD3_Nokia_Cell_BH]";
                        }
                        if (Technology == "4G")
                        {
                            s1_E = KPI_E + "as'" + KPI + "' , " + Traffic_E + "as " + "'Paylaod(GB)'" + " , " + Ave_E + "as " + "'Availability'" + " from [TBL_LTE_CELL_BH_E]";
                            s1_H = KPI_H + "as'" + KPI + "' , " + Traffic_H + "as " + "'Paylaod(GB)'" + " , " + Ave_H + "as " + "'Availability'" + " from [TBL_LTE_CELL_BH_H]";
                            s1_N = KPI_N + "as'" + KPI + "' , " + Traffic_N + "as " + "'Paylaod(GB)'" + " , " + Ave_N + "as " + "'Availability'" + " from [TBL_LTE_CELL_BH_N]";
                        }
                    }

                    if (textBox3.Text != "")
                    {
                        string KPI_Q = "";
                        if (Technology == "2G")
                        {
                            KPI_Q = "select 'Ericsson' as 'Vendor', BSC, Cell, Province, Date, " + s1_E + " where cast(Date as Date) = '" + Date_List + "' and " + s2_E + " and (" + Node_Index_List + ")" +
                           " union all " +
                           "select 'Huawei' as 'Vendor',  BSC, Cell, Province, Date, " + s1_H + " where cast(Date as Date) = '" + Date_List + "' and " + s2_H + " and (" + Node_Index_List + ")" +
                           " union all " +
                             "select 'Nokia' as 'Vendor',  BSC, Seg, Province, Date, " + s1_N + " where cast(Date as Date) = '" + Date_List + "' and " + s2_N + " and (" + Node_Index_List + ")";
                        }
                        if (Technology == "3G_CS")
                        {
                            KPI_Q = "select  'Ericsson' as 'Vendor' , ElementID as 'RNC', ElementID1 as 'Cell',  Date, " + s1_E + " where cast(Date as Date) = '" + Date_List + "' and " + s2_E + " and (" + Node_Index_List + ")" +
                           " union all " +
                           "select  'Huawei' as 'Vendor',  ElementID as 'RNC', ElementID1 as 'Cell',  Date, " + s1_H + " where cast(Date as Date) = '" + Date_List + "' and " + s2_H + " and (" + Node_Index_List + ")" +
                           " union all " +
                             "select 'Nokia' as 'Vendor', ElementID as 'RNC', ElementID1 as 'Cell',  Date, " + s1_N + " where cast(Date as Date) = '" + Date_List + "' and " + s2_N + " and (" + Node_Index_List + ")";
                        }
                        if (Technology == "3G_PS")
                        {
                            KPI_Q = "select  'Ericsson' as 'Vendor' , ElementID as 'RNC', ElementID1 as 'Cell',  Date, " + s1_E + " where cast(Date as Date) = '" + Date_List + "' and " + s2_E + " and (" + Node_Index_List + ")" +
                           " union all " +
                           "select  'Huawei' as 'Vendor',  ElementID as 'RNC', ElementID1 as 'Cell',  Date, " + s1_H + " where cast(Date as Date) = '" + Date_List + "' and " + s2_H + " and (" + Node_Index_List + ")" +
                           " union all " +
                             "select 'Nokia' as 'Vendor', ElementID as 'RNC', ElementID1 as 'Cell',  Date, " + s1_N + " where cast(Date as Date) = '" + Date_List + "' and " + s2_N + " and (" + Node_Index_List + ")";
                        }

                        if (Technology == "4G")
                        {
                            KPI_Q = "select  'Ericsson' as 'Vendor' , Province, eNodeB as 'Cell',  Datetime, " + s1_E + " where cast(Datetime as Date) = '" + Date_List + "' and " + s2_E + " and (" + Node_Index_List + ")" +
                           " union all " +
                           "select  'Huawei' as 'Vendor',  Province, eNodeB as 'Cell',  Datetime, " + s1_H + " where cast(Datetime as Date) = '" + Date_List + "' and " + s2_H + " and (" + Node_Index_List + ")" +
                           " union all " +
                             "select 'Nokia' as 'Vendor', Province, ElementID1 as 'Cell',  Date, " + s1_N + " where cast(Date as Date) = '" + Date_List + "' and " + s2_N + " and (" + Node_Index_List + ")";
                        }
                        SqlCommand Node_List_Quary1 = new SqlCommand(KPI_Q, connection);
                        Node_List_Quary1.ExecuteNonQuery();
                        DataTable Table_Node_List_Quary = new DataTable();
                        SqlDataAdapter dataAdapter_Node_List_Quary = new SqlDataAdapter(Node_List_Quary1);
                        dataAdapter_Node_List_Quary.Fill(Table_Node_List_Quary);



              

                        XLWorkbook wb = new XLWorkbook();
                        wb.Worksheets.Add(Table_Node_List_Quary, "WPC");

                        var saveFileDialog = new SaveFileDialog
                        {
                            FileName = "WPC_" + Interval + "_" + Technology + "_" + KPI,
                            Filter = "Excel files|*.xlsx",
                            Title = "Save an Excel File"
                        };

                        saveFileDialog.ShowDialog();

                        if (!String.IsNullOrWhiteSpace(saveFileDialog.FileName))
                            wb.SaveAs(saveFileDialog.FileName);

                        label15.Text = "Finished";
                        label15.BackColor = Color.Green;

                        MessageBox.Show("Finished");





                    }
                }










            //}
            //else
            //{
            //    MessageBox.Show("You should have an ID Code");
            //}




        }

        private void checkBox11_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox11.Checked == true)
            {
                sign = "<";
                checkBox12.Checked = false;
            }
        }

        private void checkBox12_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox12.Checked == true)
            {
                sign = ">";
                checkBox11.Checked = false;
            }
        }
    }
}
