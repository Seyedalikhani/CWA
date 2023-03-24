using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;
using System.Data.SqlClient;

namespace CWA
{
    public partial class CR : Form
    {
        public CR()
        {
            InitializeComponent();
        }


        public Main form1;


        public CR(Form form)
        {
            InitializeComponent();
            form1 = (Main)form;
        }



        public string FName = "";
        public int number_of_rows = 0;
        public IXLWorksheet Source_worksheet = null;
        public string ConnectionString = "";
        public SqlConnection connection = new SqlConnection();
       // public string Server_Name = "172.26.7.159";
        public string DataBase_Name = "Performance_NAK";
        public string Server_Name = "PERFORMANCEDB01";
        public DateTime Check_Day = DateTime.Now;
        public DateTime Check_Day_7 = DateTime.Now;
        public string EH_sites_list_2G = "";
        public string H_sites_lis_2G = "";
        public string N_sites_list_2G = "";
        public string sites_list_3G = "";
        public string EH_sites_list_4G = "";
        public string N_sites_list_4G = "";
        public DataTable Data_Table_2G = new DataTable();
        public DataTable Data_Table_3G_CS = new DataTable();
        public DataTable Data_Table_3G_PS = new DataTable();
        public DataTable Data_Table_4G = new DataTable();
        public DataTable CR_Output_Table = new DataTable();
        public DataTable CR_Output_Table1 = new DataTable();




        public double CSSR_2G = 97;
        public double Voice_Drop_2G = 2;
        public double IHSR_2G = 97;
        public double OHSR_2G = 97;
        public double TCH_Cong_2G = 1;
        public double TCH_ASFR_2G = 2;
        public double SDCCH_SR_2G = 97;
        public double SDCCH_Drop_2G = 2;
        public double TCH_Availability_2G = 99.5;
        public double SDCCH_Cong_2G = 1;
        public double RX_DL_2G = 97;
        public double RX_UL_2G = 97;
        public double TCH_Traffic_2G = 0;
       
    

        public double CS_RAB_3G = 98;
        public double CS_RRC_3G = 98;
        public double CS_Drop_3G = 1;
        public double PS_RAB_3G = 98;
        public double PS_RRC_3G = 98;
        public double PS_Drop_3G = 1;
        public double User_THR_3G = 1;
        public double Cell_THR_3G = 2;
        public double Soft_HO_3G = 98;
        public double Voice_Traffic_3G = 0;
        public double Data_Traffic_3G = 0;

        public double UE_DL_4G = 5;
        public double UE_UL_4G = 0.7;
        public double ERAB_Drop_4G = 1;
        public double ERAB_Setup_4G = 98;
        public double IntraFreq_HO_4G = 97;
        public double RRC_SR_4G = 98;
        public double DL_Latency_4G = 100;
        public double UL_Packet_Loss_4G = 1;
        public double CSSR_4G = 98;
        public double Payload_4G = 98;




        public int data_rows_count = 1;
        public string[] Site_list = new string[5000];

        private void button1_Click(object sender, EventArgs e)
        {


    //        string userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
    //        string currentUser = userName.Substring(8, userName.Length - 8);


    //        string[] authorizedUsers = new string[]
    //{
    //           "ahmad.alikhani",
    //           "amir.amiri",
    //           "golshan.hemati",
    //           "mona.tajbakhsh",
    //           "javad.abed",
    //           "majedeh.seydi",
    //           "neda.shafieesabet",
    //           "arash.naghdehforoushha",
    //           "parisa.pirnia"

    //};

    //        if (authorizedUsers.Contains(currentUser.ToLower()))
    //        {
    //            string Authorized = "OK";
    //        }
    //        else
    //        {
    //            MessageBox.Show("Limited Access! Need Authorization by Admin");
    //            this.Close();
    //        }



            openFileDialog1.DefaultExt = "xlsx";
            openFileDialog1.Filter = "Excel File|*.xlsx";
            DialogResult result = openFileDialog1.ShowDialog();
            string File_Name = openFileDialog1.SafeFileName.ToString();

            if (result == DialogResult.OK)
            {
                string file = openFileDialog1.FileName;
                FName = file;
                var Source_workbook = new XLWorkbook(file, XLEventTracking.Disabled);
                Source_worksheet = null;
                Source_worksheet = Source_workbook.Worksheet("CR_Country");

                label3.Text = "File's Loaded";
                label3.BackColor = Color.GreenYellow;
                number_of_rows = Source_worksheet.RowsUsed().Count();

                for (int k = 2; k <= number_of_rows; k++)
                {
                    string Performance_Engineer = Source_worksheet.Cell(k, 4).Value.ToString();
                    Site_list[k - 2] = Source_worksheet.Cell(k, 1).Value.ToString();
                    if (Performance_Engineer != "")
                    {
                        if (!comboBox1.Items.Contains(Performance_Engineer))
                        {
                            comboBox1.Items.Add(Performance_Engineer);

                        }
                    }
                }


            }

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            string Performance_Engineer = comboBox1.SelectedItem.ToString();
            for (int k = 2; k <= number_of_rows; k++)
            {
                string PE = Source_worksheet.Cell(k, 4).Value.ToString();
                string Check_Type = Source_worksheet.Cell(k, 11).Value.ToString();
                string Site = Source_worksheet.Cell(k, 1).Value.ToString();
                string End_Date = Convert.ToString(Source_worksheet.Cell(k, 7).Value.ToString());
                if (PE == "Golshan")
                {
                    listBox1.Items.Add(Site + " _ " + End_Date);
                }
            }


        }

        private void Form7_Load(object sender, EventArgs e)
        {


            //string Server_Name = @"NAKPRG-NB1243\" + "AHMAD";
            //string DataBase_Name = "Dashboards";

           //ConnectionString = @"Server=" + Server_Name + "; Database=" + DataBase_Name + "; Trusted_Connection=True;";


           ConnectionString = @"Server=" + Server_Name + "; Database=" + DataBase_Name + "; User ID=cwpcApp;Password=cwpcApp@830625#Ahmad";
            connection = new SqlConnection(ConnectionString);
            connection.Open();

            CR_Output_Table.Columns.Add("Site", typeof(string));
            CR_Output_Table.Columns.Add("Cell", typeof(string));
            CR_Output_Table.Columns.Add("Node", typeof(string));
            CR_Output_Table.Columns.Add("Vendor", typeof(string));
            CR_Output_Table.Columns.Add("Technology", typeof(string));
            CR_Output_Table.Columns.Add("KPI", typeof(string));
            CR_Output_Table.Columns.Add("Vlaue After CR", typeof(string));
            CR_Output_Table.Columns.Add("Vlaue Before CR", typeof(string));
            CR_Output_Table.Columns.Add("Performance Comment", typeof(string));
            CR_Output_Table.Columns.Add("Performance Engineer", typeof(string));
            CR_Output_Table.Columns.Add("CR Num", typeof(string));
            CR_Output_Table.Columns.Add("Start Time", typeof(DateTime));
            CR_Output_Table.Columns.Add("End Time", typeof(DateTime));
            CR_Output_Table.Columns.Add("CR Status ", typeof(string));
            CR_Output_Table.Columns.Add("CR Description", typeof(string));






            //   CR_Output_Table1 = CR_Output_Table;

            dataGridView1.ColumnCount = 15;

            //dataGridView1.Rows.Clear();
            //dataGridView1.RowCount = Data_Table_4G.Rows.Count + 1;



        }

        private void button4_Click(object sender, EventArgs e)
        {
            ConnectionString = @"Server=" + Server_Name + "; Database=" + DataBase_Name + "; User ID=cwpcApp;Password=cwpcApp@830625#Ahmad";
            connection = new SqlConnection(ConnectionString);
            connection.Open();

            EH_sites_list_2G = "";
            H_sites_lis_2G = "";
            N_sites_list_2G = "";
            sites_list_3G = "";
            EH_sites_list_4G = "";
            N_sites_list_4G = "";


            for (int i = 0; i < listBox1.SelectedItems.Count; i++)
            {
                string site_code_date = listBox1.SelectedItems[i].ToString();
                string site_code = site_code_date.Substring(0, 6);
                {
                    EH_sites_list_2G = EH_sites_list_2G + "substring([Cell],1,6)='" + site_code + "' or ";
                    H_sites_lis_2G = H_sites_lis_2G + "substring([Cell],1,2)+substring([Cell],5,4)='" + site_code + "' or ";
                    N_sites_list_2G = N_sites_list_2G + "substring([Seg],1,2)+substring([Seg],5,4)='" + site_code + "' or ";
                    sites_list_3G = sites_list_3G + "substring([ElementID1],1,2)+substring([ElementID1],5,4)='" + site_code + "' or ";
                    EH_sites_list_4G = EH_sites_list_4G + "substring([eNodeB],1,2)+substring([eNodeB],5,4)='" + site_code + "' or ";
                    N_sites_list_4G = N_sites_list_4G + "substring([ElementID1],1,2)+substring([ElementID1],5,4)='" + site_code + "' or ";
                }
            }

            EH_sites_list_2G = EH_sites_list_2G.Substring(0, EH_sites_list_2G.Length - 4);
            H_sites_lis_2G = H_sites_lis_2G.Substring(0, H_sites_lis_2G.Length - 4);
            N_sites_list_2G = N_sites_list_2G.Substring(0, N_sites_list_2G.Length - 4);
            sites_list_3G = sites_list_3G.Substring(0, sites_list_3G.Length - 4);
            EH_sites_list_4G = EH_sites_list_4G.Substring(0, EH_sites_list_4G.Length - 4);
            N_sites_list_4G = N_sites_list_4G.Substring(0, N_sites_list_4G.Length - 4);


            string Data_Quary_2G = "";
            string Data_Quary_3G_CS = "";
            string Data_Quary_3G_PS = "";
            string Data_Quary_4G = "";

            Data_Quary_2G = @"select [Date], [BSC],  substring([Cell],1,6) as 'Site' ,[Cell], [TCH_Traffic] as 'TCH_Traffic_Daily (Erlang)', [CSSR_MCI] as'CSSR', [CDR(not Affected by incoming Handovers from 3G)(Eric_CELL)] as 'Voice Drop Rate', [IHSR] as 'IHSR', [OHSR] as 'OHSR', [TCH_Congestion] as 'TCH Congestion Rate', [TCH_Assign_Fail_Rate(NAK)(Eric_CELL)] as 'TCH ASFR', [SDCCH_Access_Succ_Rate] as 'SDCCH SR', [SDCCH_Drop_Rate] as 'SDDH Drop Rate',  [TCH_Availability] as 'TCH Availability' , [SDCCH_Congestion] as 'SDCCH Cong' , [RxQual_DL] as 'RX_DL', [RxQual_UL] as 'RX_UL', 'Ericsson' as Vendor from [dbo].[CC2_Ericsson_Cell_Daily] where  (" + EH_sites_list_2G + ") and (" + "Date = '" + Check_Day + "' or Date='" + Check_Day_7 + "')" +
                        @" union all select [Date], [BSC],   substring([Cell],1,6) as 'Site', [Cell], [TCH_Traffic] as 'TCH_Traffic_Daily (Erlang)', [CSSR3] as'CSSR', [CDR3] as 'Voice Drop Rate', [IHSR2] as 'IHSR', [OHSR2] as 'OHSR', [TCH_Cong] as 'TCH Congestion Rate', [TCH_Assignment_FR] as 'TCH ASFR', [SDCCH_Access_Success_Rate2] as 'SDCCH SR', [SDCCH_Drop_Rate] as 'SDDH Drop Rate',    [TCH_Availability] as 'TCH Availability'  , [SDCCH_Congestion_Rate] as 'SDCCH Cong' , [RX_QUALITTY_DL_NEW] as 'RX_DL', [RX_QUALITTY_UL_NEW] as 'RX_UL', 'Huawei' as Vendor from [dbo].[CC2_Huawei_Cell_Daily] where (" + EH_sites_list_2G + ") and (" + "Date = '" + Check_Day + "' or Date='" + Check_Day_7 + "')" +
                        @" union all select [Date], [BSC] ,   substring([Cell],1,2)+substring([Cell],5,4) as 'Site', [Cell], [TCH_Traffic] as 'TCH_Traffic_Daily (Erlang)', [CSSR3] as'CSSR', [CDR3] as 'Voice Drop Rate', [IHSR2] as 'IHSR', [OHSR2] as 'OHSR', [TCH_Cong] as 'TCH Congestion Rate', [TCH_Assignment_FR] as 'TCH ASFR', [SDCCH_Access_Success_Rate2] as 'SDCCH SR', [SDCCH_Drop_Rate] as 'SDDH Drop Rate',    [TCH_Availability] as 'TCH Availability' , [SDCCH_Congestion_Rate] as 'SDCCH Cong' , [RX_QUALITTY_DL_NEW] as 'RX_DL', [RX_QUALITTY_UL_NEW] as 'RX_UL', 'Huawei' as Vendor from [dbo].[CC2_Huawei_Cell_Daily] where (" + H_sites_lis_2G + ") and (" + "Date = '" + Check_Day + "' or Date='" + Check_Day_7 + "')" +
                        @" union all select [Date], [BSC],   substring([SEG],1,2)+substring([SEG],5,4) as 'Site', [SEG] as 'Cell', [TCH_Traffic] as 'TCH_Traffic_Daily (Erlang)', [CSSR_MCI] as'CSSR', [CDR(including_CS_IRAT_handovers_3G_to2G)(Nokia_SEG)] as 'Voicde  Drop Rate', [IHSR] as 'IHSR', [OHSR] AS 'OHSR', [TCH_Cong_Rate] as 'TCH Congestion Rate', [TCH_Assignment_FR] as 'TCH ASFR', [SDCCH_Access_Success_Rate] as 'SDCCH SR', [SDCCH_Drop_Rate] as 'SDDH Drop Rate',     [TCH_Availability] as 'TCH Availability' , [SDCCH_Congestion_Rate] as 'SDCCH Cong' , [RxQuality_DL] as 'RX_DL', [RxQuality_UL] as 'RX_UL', 'Nokia' as Vendor from [dbo].[CC2_Nokia_Cell_Daily] where (" + N_sites_list_2G + ") and (" + "Date = '" + Check_Day + "' or Date='" + Check_Day_7 + "')";

            SqlCommand Data_Quary1 = new SqlCommand(Data_Quary_2G, connection);
            Data_Quary1.CommandTimeout = 0;
            Data_Quary1.ExecuteNonQuery();
            Data_Table_2G = new DataTable();
            SqlDataAdapter Date_Table1 = new SqlDataAdapter(Data_Quary1);
            Date_Table1.Fill(Data_Table_2G);


            Data_Quary_3G_CS = @" select [Date], [ElementID] as 'RNC',  substring([ElementID1],1,8) as 'Site', [ElementID1] as 'Cell', [CS_Traffic] as 'CS_Traffic_Daily (Erlang)', [Cs_RAB_Establish_Success_Rate] as 'CS RAB Establish', [CS_RRC_Setup_Success_Rate] as'CS RRC SR', [CS_Drop_Call_Rate] as 'Voice Drop Rate', [Cell_Availability_Rate_Exclude_Blocking(UCELL_Eric)] as 'Cell Availability' , [Soft_Handover_Succ_Rate] as 'Soft HO SR', 'Ericsson' as Vendor from [dbo].[CC3_Ericsson_Cell_Daily] where  (" + sites_list_3G + ") and (" + "Date = '" + Check_Day + "' or Date='" + Check_Day_7 + "')" +
   @" union all select [Date],  [ElementID] as 'RNC', substring([ElementID1],1,8) as 'Site', [ElementID1] as 'Cell', [CS_Erlang] as 'CS_Traffic_Daily (Erlang)', [CS_RAB_Setup_Success_Ratio] as 'CS RAB Establish', [CS_RRC_Connection_Establishment_SR] as'CS RRC SR',  [AMR_Call_Drop_Ratio_New(Hu_CELL)] as 'Voice Drop Rate',  [Radio_Network_Availability_Ratio(Hu_Cell)] as 'Cell Availability' , [Soft_Handover_Succ_Rate] as 'Soft HO SR', 'Huawei' as Vendor  from [dbo].[CC3_Huawei_Cell_Daily] where  (" + sites_list_3G + ") and (" + "Date = '" + Check_Day + "' or Date='" + Check_Day_7 + "')" +
   @" union all select [Date], [ElementID] as 'RNC',  substring([ElementID1],1,8) as 'Site', [ElementID1] as 'Cell', [CS_Traffic] as 'CS_Traffic_Daily (Erlang)', [CS_RAB_Establish_Success_Rate] as 'CS RAB Establish', [CS_RRC_SETUP_SR_WITHOUT_REPEAT(CELL_NOKIA)] as 'CS RRC SR', [CS_Drop_Call_Rate] as 'Voice Drop Rate', [Cell_Availability_excluding_blocked_by_user_state] as 'Cell Availability' , [Soft_Handover_Succ_Rate] as 'Soft HO SR', 'Nokia' as Vendor  from [dbo].[CC3_Nokia_Cell_Daily] where  (" + sites_list_3G + ") and (" + "Date = '" + Check_Day + "' or Date='" + Check_Day_7 + "')";

            SqlCommand Data_Quary2 = new SqlCommand(Data_Quary_3G_CS, connection);
            Data_Quary2.CommandTimeout = 0;
            Data_Quary2.ExecuteNonQuery();
            Data_Table_3G_CS = new DataTable();
            SqlDataAdapter Date_Table2 = new SqlDataAdapter(Data_Quary2);
            Date_Table2.Fill(Data_Table_3G_CS);


            Data_Quary_3G_PS = @" select [Date], [ElementID] as 'RNC',  substring([ElementID1],1,8) as 'Site', [ElementID1] as 'Cell', [PS_Volume(GB)(UCell_Eric)] as 'PS_Traffic_Daily (GB)', [Ps_RAB_Establish_Success_Rate] as 'PS RAB Establish', [PS_RRC_Setup_Success_Rate(UCell_Eric)] as'PS RRC SR', [PS_Drop_Call_Rate(UCell_Eric)] as 'PS Drop Rate', [HS_USER_Throughput_NET_PQ(Mbps)(UCell_Eric)] as '3G User THR' , [HSDPA_Cell_Scheduled_Throughput(mbps)(UCell_Eric)] as '3G Cell THR' , 'Ericsson' as Vendor from [dbo].[RD3_Ericsson_Cell_Daily] where  (" + sites_list_3G + ") and (" + "Date = '" + Check_Day + "' or Date='" + Check_Day_7 + "')" +
 @" union all select [Date], [ElementID] as 'RNC',    substring([ElementID1],1,8) as 'Site', [ElementID1] as 'Cell', [PAYLOAD] as 'PS_Traffic_Daily (GB)', [PS_RAB_Setup_Success_Ratio] as 'PS RAB Establish', [PS_RRC_Connection_success_Rate_repeatless(Hu_Cell)] as'PS RRC SR',  [PS_Call_Drop_Ratio] as 'PS Drop Rate',  [AVERAGE_HSDPA_USER_THROUGHPUT_DC+SC(Mbit/s)(CELL_HUAWEI)] as '3G User THR' , [HSDPA_SCHEDULING_Cell_throughput(CELL_HUAWEI)] as '3G Cell THR'  , 'Huawei' as Vendor from [dbo].[RD3_Huawei_Cell_Daily] where  (" + sites_list_3G + ") and (" + "Date = '" + Check_Day + "' or Date='" + Check_Day_7 + "')" +
 @" union all select [Date], [ElementID] as 'RNC',  substring([ElementID1],1,8) as 'Site', [ElementID1] as 'Cell', [PS_Payload_Total(HS+R99)(Nokia_CELL)_GB] as 'PS_Traffic_Daily (GB)', [RAB_Setup_and_Access_Complete_Ratio_for_NRT_Service_from_User_pe] as 'PS RAB Establish', [PS_RRCSETUP_SR] as 'PS RRC SR', [Packet_Session_Drop_Ratio_NOKIA(CELL_NOKIA)] as 'PS Drop Rate', [AVERAGE_HSDPA_USER_THROUGHPUT_DC+SC(Mbit/s)(Nokia_CELL)] as '3G User THR' , [Active_HS-DSCH_cell_throughput_mbs(CELL_nokia)] as '3G Cell THR' , 'Nokia' as Vendor from [dbo].[RD3_Nokia_Cell_Daily] where  (" + sites_list_3G + ") and (" + "Date = '" + Check_Day + "' or Date='" + Check_Day_7 + "')";


            SqlCommand Data_Quary3 = new SqlCommand(Data_Quary_3G_PS, connection);
            Data_Quary3.CommandTimeout = 0;
            Data_Quary3.ExecuteNonQuery();
            Data_Table_3G_PS = new DataTable();
            SqlDataAdapter Date_Table3 = new SqlDataAdapter(Data_Quary3);
            Date_Table3.Fill(Data_Table_3G_PS);


            Data_Quary_4G = @" select [Datetime],  substring([eNodeB],1,8) as 'Site', [eNodeB] as 'Cell', [Total_Volume(UL+DL)(GB)(eNodeB_Eric)] as 'PS_Traffic_Daily (GB)', [Average_UE_DL_Throughput(Mbps)(eNodeB_Eric)] as 'UE DL THR (Mbps)', [Average_UE_UL_Throughput(Mbps)(eNodeB_Eric)] as'UE UL THR (Mbps)', [E_RAB_Drop_Rate(eNodeB_Eric)] as 'ERAB Drop Rate', [E-RAB_Setup_SR_incl_added_New(EUCell_Eric)] as 'ERAB Setup SR', [IntraF_Handover_Execution(eNodeB_Eric)] as 'Intra Freq HO SR' , [RRC_Estab_Success_Rate(ReAtt)(EUCell_Eric)] as 'RRC Connection SR' , [Cell_Availability_Rate_Exclude_Blocking(Cell_EricLTE)] as 'Cell Availability' , [Average_UE_DL_Latency(ms)(eNodeB_Eric)] as 'DL Latency', [Average_UE_Ul_Packet_Loss_Rate(eNodeB_Eric)] as 'UL Packet Loss Rate', [LTE_Service_Success_Rate(eNodeB_Eric)] as 'Service SR' , 'Ericsson' as Vendor from [dbo].[TBL_LTE_CELL_Daily_E] where  (" + EH_sites_list_4G + ") and (" + "Datetime = '" + Check_Day + "' or Datetime='" + Check_Day_7 + "')" +
@" union all select [Datetime],    substring([eNodeB],1,8) as 'Siite', [eNodeB] as 'Cell', [Total_Traffic_Volume(GB)] as 'PS_Traffic_Daily (GB)', [Average_Downlink_User_Throughput(Mbit/s)] as 'UE DL THR (Mbps)', [Average_UPlink_User_Throughput(Mbit/s)] as'UE UL THR (Mbps)',  [Call_Drop_Rate] as 'ERAB Drop Rate',  [E-RAB_Setup_Success_Rate(Hu_Cell)] as 'ERAB Setup SR'  , [IntraF_HOOut_SR] as 'Intra Freq HO SR' , [RRC_Connection_Setup_Success_Rate_service] as 'RRC Connection SR' , [Cell_Availability_Rate_Exclude_Blocking(Cell_Hu)] as 'Cell Availability' , [Average_DL_Latency_ms(Huawei_LTE_EUCell)] as 'DL Latency', [Average_UL_Packet_Loss_%(Huawei_LTE_UCell)] as 'UL Packet Loss Rate', [CSSR(ALL)] as 'Service SR' , 'Huawei' as Vendor from [dbo].[TBL_LTE_CELL_Daily_H] where  (" + EH_sites_list_4G + ") and (" + "Datetime = '" + Check_Day + "' or Datetime='" + Check_Day_7 + "')" +
@" union all select [Date], substring([ElementID1],1,8) as 'Site', [ElementID1] as 'Cell', [Total_Payload_GB(Nokia_LTE_CELL)] as 'PS_Traffic_Daily (GB)', [User_Throughput_DL_mbps(Nokia_LTE_CELL)] as 'UE DL THR (Mbps)', [User_Throughput_UL_mbps(Nokia_LTE_CELL)] as 'UE UL THR (Mbps)', [E-RAB_Drop_Ratio_RAN_View(Nokia_LTE_CELL)] as 'ERAB Drop Rate', [E-RAB_Setup_SR_incl_added(Nokia_LTE_CELL)] as 'ERAB Setup SR' , [HO_Success_Ratio_intra_eNB(Nokia_LTE_CELL)] as 'Intra Freq HO SR' , [RRC_Connection_Setup_Success_Ratio(Nokia_LTE_CELL)] as 'RRC Connection SR' , [cell_availability_exclude_manual_blocking(Nokia_LTE_CELL)] as 'Cell Availability', [Average_Latency_DL_ms(Nokia_LTE_CELL)] as 'DL Latency', [Packet_loss_UL(Nokia_EUCELL)] as 'UL Packet Loss Rate', [Initial_E-RAB_Accessibility(Nokia_LTE_CELL)] as 'Service SR'  , 'Nokia' as Vendor from [dbo].[TBL_LTE_CELL_Daily_N] where  (" + N_sites_list_4G + ") and (" + "Date= '" + Check_Day + "' or Date='" + Check_Day_7 + "')";



            SqlCommand Data_Quary4 = new SqlCommand(Data_Quary_4G, connection);
            Data_Quary4.CommandTimeout = 0;
            Data_Quary4.ExecuteNonQuery();
            Data_Table_4G = new DataTable();
            SqlDataAdapter Date_Table4 = new SqlDataAdapter(Data_Quary4);
            Date_Table4.Fill(Data_Table_4G);



            MessageBox.Show("Loaded");


        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            Check_Day = dateTimePicker1.Value.Date;
            Check_Day_7 = dateTimePicker1.Value.Date.AddDays(-7);

        }

        private void button5_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            
            CR_Output_Table.Rows.Clear();
            CR_Output_Table1.Rows.Clear();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //dataGridView1.Rows.Clear();
            //CR_Output_Table.Rows.Clear();


            var distinctCells_2G = Data_Table_2G.AsEnumerable()
.Select(s => new
{
    Cell = s.Field<string>("Cell"),
})
.Distinct().ToList();



            var distinctCells_3G_CS = Data_Table_3G_CS.AsEnumerable()
.Select(s => new
{
    Cell = s.Field<string>("Cell"),
})
.Distinct().ToList();


            var distinctCells_3G_PS = Data_Table_3G_PS.AsEnumerable()
.Select(s => new
{
    Cell = s.Field<string>("Cell"),
})
.Distinct().ToList();


            var distinctCells_4G = Data_Table_4G.AsEnumerable()
.Select(s => new
{
    Cell = s.Field<string>("Cell"),
})
.Distinct().ToList();


            for (int j = 0; j < distinctCells_2G.Count; j++)
            {
                var cell_data = (from p in Data_Table_2G.AsEnumerable()
                                 where p.Field<string>("Cell") == distinctCells_2G[j].Cell
                                 select p).ToList();




                if (cell_data.Count == 2)
                {

                    string site_code = cell_data[0].ItemArray[2].ToString();

                    List<int> index_finder = new List<int>();
                    index_finder = Site_list.Select((s, i) => new { i, s })
    .Where(t => t.s == site_code)
    .Select(t => t.i)
    .ToList();

                    int index = index_finder[0] + 2;
                    string Eng = Source_worksheet.Cell(index, 4).Value.ToString();
                    string CR = Source_worksheet.Cell(index, 5).Value.ToString();
                    DateTime SD = Convert.ToDateTime(Source_worksheet.Cell(index, 6).Value);
                    DateTime ED = Convert.ToDateTime(Source_worksheet.Cell(index, 7).Value);
                    //SD = SD.Date;
                    //ED = ED.Date;
                    string Status = Source_worksheet.Cell(index, 8).Value.ToString();
                    string Description = Source_worksheet.Cell(index, 9).Value.ToString();

                  
                    if (cell_data[1].ItemArray[5].ToString() != "" && cell_data[0].ItemArray[5].ToString() != "")
                    {
                        string Vendor = cell_data[1].ItemArray[17].ToString();
                        string KPI_Name = "";
                        if (Vendor=="Ericsson")
                        {
                            KPI_Name = "CSSR_MCI(Eric_Cell)";
                        }
                        if (Vendor == "Huawei")
                        {
                            KPI_Name = "CSSR3(HU_Cell)";
                        }
                        if (Vendor == "Nokia")
                        {
                            KPI_Name = "CSSR_MCI(Nokia_SEG)";
                        }

                        double CSSR_2G_Day2 = Convert.ToDouble(cell_data[1].ItemArray[5]);
                        double CSSR_2G_Day1 = Convert.ToDouble(cell_data[0].ItemArray[5]);
                        if (CSSR_2G_Day2 > 0 && CSSR_2G_Day2 < CSSR_2G)
                        {
                            if (CSSR_2G_Day1 > 0)
                            {
                                if (100 * (CSSR_2G_Day2 - CSSR_2G_Day1) / CSSR_2G_Day1 < -5)
                                {
                                    CR_Output_Table.Rows.Add(cell_data[1].ItemArray[2], cell_data[1].ItemArray[3], cell_data[1].ItemArray[1], Vendor, "2G", KPI_Name, CSSR_2G_Day2, CSSR_2G_Day1, "", Eng, CR, SD, ED, Status, Description);
                                }
                            }
                            if (CSSR_2G_Day1 == 0)
                            {
                                CR_Output_Table.Rows.Add(cell_data[1].ItemArray[2], cell_data[1].ItemArray[3], cell_data[1].ItemArray[1], Vendor, "2G", KPI_Name, CSSR_2G_Day2, CSSR_2G_Day1, "", Eng, CR, SD, ED, Status, Description);
                            }
                        }
                    }

                    if (cell_data[1].ItemArray[6].ToString() != "" && cell_data[0].ItemArray[6].ToString() != "")
                    {
                        string Vendor = cell_data[1].ItemArray[17].ToString();
                        string KPI_Name = "";
                        if (Vendor == "Ericsson")
                        {
                            KPI_Name = "2G_Voice_Call_Drop_Rate(Eric_Cell)";
                        }
                        if (Vendor == "Huawei")
                        {
                            KPI_Name = "2G_Voice_Call_Drop_Rate(HU_Cell)";
                        }
                        if (Vendor == "Nokia")
                        {
                            KPI_Name = "2G_Voice_Call_Drop_Rate(Nokia_SEG)";
                        }
                        double CDR_2G_Day2 = Convert.ToDouble(cell_data[1].ItemArray[6]);
                        double CDR_2G_Day1 = Convert.ToDouble(cell_data[0].ItemArray[6]);
                        if (CDR_2G_Day2 > Voice_Drop_2G)
                        {
                            if (CDR_2G_Day1 > 0)
                            {
                                if (100 * (CDR_2G_Day2 - CDR_2G_Day1) / CDR_2G_Day1 > 30)
                                {
                                    CR_Output_Table.Rows.Add(cell_data[1].ItemArray[2], cell_data[1].ItemArray[3], cell_data[1].ItemArray[1], Vendor, "2G", KPI_Name, CDR_2G_Day2, CDR_2G_Day1, "", Eng, CR, SD, ED, Status, Description);
                                }
                            }
                            if (CDR_2G_Day1 == 0)
                            {
                                CR_Output_Table.Rows.Add(cell_data[1].ItemArray[2], cell_data[1].ItemArray[3], cell_data[1].ItemArray[1], Vendor, "2G", KPI_Name, CDR_2G_Day2, CDR_2G_Day1, "", Eng, CR, SD, ED, Status, Description);
                            }
                        }
                    }


                    if (cell_data[1].ItemArray[7].ToString() != "" && cell_data[0].ItemArray[7].ToString() != "")
                    {
                        string Vendor = cell_data[1].ItemArray[17].ToString();
                        string KPI_Name = "";
                        if (Vendor == "Ericsson")
                        {
                            KPI_Name = "IHSR(Eric_Cell)";
                        }
                        if (Vendor == "Huawei")
                        {
                            KPI_Name = "IHSR2(HU_Cell)";
                        }
                        if (Vendor == "Nokia")
                        {
                            KPI_Name = "IHSR(Nokia_SEG)";
                        }
                        double IHSR_2G_Day2 = Convert.ToDouble(cell_data[1].ItemArray[7]);
                        double IHSR_2G_Day1 = Convert.ToDouble(cell_data[0].ItemArray[7]);
                        if (IHSR_2G_Day2 > 0 && IHSR_2G_Day2 < IHSR_2G)
                        {
                            if (IHSR_2G_Day1 > 0)
                            {
                                if (100 * (IHSR_2G_Day2 - IHSR_2G_Day1) / IHSR_2G_Day1 < -5)
                                {
                                    CR_Output_Table.Rows.Add(cell_data[1].ItemArray[2], cell_data[1].ItemArray[3], cell_data[1].ItemArray[1], Vendor, "2G", KPI_Name, IHSR_2G_Day2, IHSR_2G_Day1, "", Eng, CR, SD, ED, Status, Description);
                                }
                            }
                            if (IHSR_2G_Day1 == 0)
                            {
                                CR_Output_Table.Rows.Add(cell_data[1].ItemArray[2], cell_data[1].ItemArray[3], cell_data[1].ItemArray[1], Vendor, "2G", KPI_Name, IHSR_2G_Day2, IHSR_2G_Day1, "", Eng, CR, SD, ED, Status, Description);
                            }
                        }
                    }


                    if (cell_data[1].ItemArray[8].ToString() != "" && cell_data[0].ItemArray[8].ToString() != "")
                    {
                        string Vendor = cell_data[1].ItemArray[17].ToString();
                        string KPI_Name = "";
                        if (Vendor == "Ericsson")
                        {
                            KPI_Name = "OHSR(Eric_Cell)";
                        }
                        if (Vendor == "Huawei")
                        {
                            KPI_Name = "OHSR2(HU_Cell)";
                        }
                        if (Vendor == "Nokia")
                        {
                            KPI_Name = "OHSR(Nokia_SEG)";
                        }
                        double OHSR_2G_Day2 = Convert.ToDouble(cell_data[1].ItemArray[8]);
                        double OHSR_2G_Day1 = Convert.ToDouble(cell_data[0].ItemArray[8]);
                        if (OHSR_2G_Day2 > 0 && OHSR_2G_Day2 < OHSR_2G)
                        {
                            if (OHSR_2G_Day1 > 0)
                            {
                                if (100 * (OHSR_2G_Day2 - OHSR_2G_Day1) / OHSR_2G_Day1 < -5)
                                {
                                    CR_Output_Table.Rows.Add(cell_data[1].ItemArray[2], cell_data[1].ItemArray[3], cell_data[1].ItemArray[1], Vendor, "2G", KPI_Name, OHSR_2G_Day2, OHSR_2G_Day1, "", Eng, CR, SD, ED, Status, Description);
                                }
                            }
                            if (OHSR_2G_Day1 == 0)
                            {
                                CR_Output_Table.Rows.Add(cell_data[1].ItemArray[2], cell_data[1].ItemArray[3], cell_data[1].ItemArray[1], Vendor, "2G", KPI_Name, OHSR_2G_Day2, OHSR_2G_Day1, "", Eng, CR, SD, ED, Status, Description);
                            }
                        }
                    }


                    if (cell_data[1].ItemArray[9].ToString() != "" && cell_data[0].ItemArray[9].ToString() != "")
                    {
                        string Vendor = cell_data[1].ItemArray[17].ToString();
                        string KPI_Name = "";
                        if (Vendor == "Ericsson")
                        {
                            KPI_Name = "TCH_Congestion_Rate(Eric_Cell)";
                        }
                        if (Vendor == "Huawei")
                        {
                            KPI_Name = "TCH_Cong(HU_Cell)";
                        }
                        if (Vendor == "Nokia")
                        {
                            KPI_Name = "TCH_Cong_Rate(Nokia_SEG)";
                        }
                        double TCH_Cong_2G_Day2 = Convert.ToDouble(cell_data[1].ItemArray[9]);
                        double TCH_Cong_2G_Day1 = Convert.ToDouble(cell_data[0].ItemArray[9]);
                        if (TCH_Cong_2G_Day2 > TCH_Cong_2G)
                        {
                            if (TCH_Cong_2G_Day1 > 0)
                            {
                                if (100 * (TCH_Cong_2G_Day2 - TCH_Cong_2G_Day1) / TCH_Cong_2G_Day1 > 30)
                                {
                                    CR_Output_Table.Rows.Add(cell_data[1].ItemArray[2], cell_data[1].ItemArray[3], cell_data[1].ItemArray[1], Vendor, "2G", KPI_Name, TCH_Cong_2G_Day2, TCH_Cong_2G_Day1, "", Eng, CR, SD, ED, Status, Description);
                                }
                            }
                            if (TCH_Cong_2G_Day1 == 0)
                            {
                                CR_Output_Table.Rows.Add(cell_data[1].ItemArray[2], cell_data[1].ItemArray[3], cell_data[1].ItemArray[1], Vendor, "2G", KPI_Name, TCH_Cong_2G_Day2, TCH_Cong_2G_Day1, "", Eng, CR, SD, ED, Status, Description);
                            }
                        }
                    }


                    if (cell_data[1].ItemArray[10].ToString() != "" && cell_data[0].ItemArray[10].ToString() != "")
                    {
                        string Vendor = cell_data[1].ItemArray[17].ToString();
                        string KPI_Name = "";
                        if (Vendor == "Ericsson")
                        {
                            KPI_Name = "TCH_Assign_Fail_Rate(Congestion_Excluded)(Eric_Cell)";
                        }
                        if (Vendor == "Huawei")
                        {
                            KPI_Name = "TCH_Assignment_FR(HU_Cell)";
                        }
                        if (Vendor == "Nokia")
                        {
                            KPI_Name = "TCH_Assignment_FR_New(Nokia_SEG)";
                        }
                        double TCH_ASFR_2G_Day2 = Convert.ToDouble(cell_data[1].ItemArray[10]);
                        double TCH_ASFR_2G_Day1 = Convert.ToDouble(cell_data[0].ItemArray[10]);
                        if (TCH_ASFR_2G_Day2 > TCH_ASFR_2G)
                        {
                            if (TCH_ASFR_2G_Day1 > 0)
                            {
                                if (100 * (TCH_ASFR_2G_Day2 - TCH_ASFR_2G_Day1) / TCH_ASFR_2G_Day1 > 30)
                                {
                                    CR_Output_Table.Rows.Add(cell_data[1].ItemArray[2], cell_data[1].ItemArray[3], cell_data[1].ItemArray[1], Vendor, "2G", KPI_Name, TCH_ASFR_2G_Day2, TCH_ASFR_2G_Day1, "", Eng, CR, SD, ED, Status, Description);
                                }
                            }
                            if (TCH_ASFR_2G_Day1 == 0)
                            {
                                CR_Output_Table.Rows.Add(cell_data[1].ItemArray[2], cell_data[1].ItemArray[3], cell_data[1].ItemArray[1], Vendor, "2G", KPI_Name, TCH_ASFR_2G_Day2, TCH_ASFR_2G_Day1, "", Eng, CR, SD, ED, Status, Description);
                            }
                        }
                    }

                    if (cell_data[1].ItemArray[11].ToString() != "" && cell_data[0].ItemArray[11].ToString() != "")
                    {
                        string Vendor = cell_data[1].ItemArray[17].ToString();
                        string KPI_Name = "";
                        if (Vendor == "Ericsson")
                        {
                            KPI_Name = "SDCCH_Access_Succ_Rate_New(Eric_Cell)";
                        }
                        if (Vendor == "Huawei")
                        {
                            KPI_Name = "SDCCH_Access_Success_Rate_New(Hu_Cell)";
                        }
                        if (Vendor == "Nokia")
                        {
                            KPI_Name = "SDCCH_Access_SR_NEW(Nokia_SEG)";
                        }
                        double SDCCH_SR_2G_Day2 = Convert.ToDouble(cell_data[1].ItemArray[11]);
                        double SDCCH_SR_2G_Day1 = Convert.ToDouble(cell_data[0].ItemArray[11]);
                        if (SDCCH_SR_2G_Day2 > 0 && SDCCH_SR_2G_Day2 < SDCCH_SR_2G)
                        {
                            if (SDCCH_SR_2G_Day1 > 0)
                            {
                                if (100 * (SDCCH_SR_2G_Day2 - SDCCH_SR_2G_Day1) / SDCCH_SR_2G_Day1 < -5)
                                {
                                    CR_Output_Table.Rows.Add(cell_data[1].ItemArray[2], cell_data[1].ItemArray[3], cell_data[1].ItemArray[1], Vendor, "2G", KPI_Name, SDCCH_SR_2G_Day2, SDCCH_SR_2G_Day1, "", Eng, CR, SD.Date, ED, Status, Description);
                                }
                            }
                            if (SDCCH_SR_2G_Day1 == 0)
                            {
                                CR_Output_Table.Rows.Add(cell_data[1].ItemArray[2], cell_data[1].ItemArray[3], cell_data[1].ItemArray[1], Vendor, "2G", KPI_Name, SDCCH_SR_2G_Day2, SDCCH_SR_2G_Day1, "", Eng, CR, SD, ED, Status, Description);
                            }
                        }
                    }

                    if (cell_data[1].ItemArray[12].ToString() != "" && cell_data[0].ItemArray[12].ToString() != "")
                    {
                        string Vendor = cell_data[1].ItemArray[17].ToString();
                        string KPI_Name = "";
                        if (Vendor == "Ericsson")
                        {
                            KPI_Name = "SDCCH_Drop_Rate(Eric_Cell)";
                        }
                        if (Vendor == "Huawei")
                        {
                            KPI_Name = "SDCCH_Drop_Rate(HU_Cell)";
                        }
                        if (Vendor == "Nokia")
                        {
                            KPI_Name = "SDCCH_Drop_Rate(Nokia_SEG)";
                        }
                        double SDCCH_Drop_2G_Day2 = Convert.ToDouble(cell_data[1].ItemArray[12]);
                        double SDCCH_Drop_2G_Day1 = Convert.ToDouble(cell_data[0].ItemArray[12]);
                        if (SDCCH_Drop_2G_Day2 > SDCCH_Drop_2G)
                        {
                            if (SDCCH_Drop_2G_Day1 > 0)
                            {
                                if (100 * (SDCCH_Drop_2G_Day2 - SDCCH_Drop_2G_Day1) / SDCCH_Drop_2G_Day1 > 30)
                                {
                                    CR_Output_Table.Rows.Add(cell_data[1].ItemArray[2], cell_data[1].ItemArray[3], cell_data[1].ItemArray[1], Vendor, "2G", KPI_Name, SDCCH_Drop_2G_Day2, SDCCH_Drop_2G_Day1, "", Eng, CR, SD, ED, Status, Description);
                                }
                            }
                            if (SDCCH_Drop_2G_Day1 == 0)
                            {
                                CR_Output_Table.Rows.Add(cell_data[1].ItemArray[2], cell_data[1].ItemArray[3], cell_data[1].ItemArray[1], Vendor, "2G", KPI_Name, SDCCH_Drop_2G_Day2, SDCCH_Drop_2G_Day1, "", Eng, CR, SD, ED, Status, Description);
                            }
                        }
                    }


                    if (cell_data[1].ItemArray[13].ToString() != "" && cell_data[0].ItemArray[13].ToString() != "")
                    {
                        string Vendor = cell_data[1].ItemArray[17].ToString();
                        string KPI_Name = "";
                        if (Vendor == "Ericsson")
                        {
                            KPI_Name = "TCH_Availability(Eric_Cell)";
                        }
                        if (Vendor == "Huawei")
                        {
                            KPI_Name = "TCH_Availability(HU_Cell)";
                        }
                        if (Vendor == "Nokia")
                        {
                            KPI_Name = "TCH_Availability(Nokia_SEG)";
                        }
                        double Availability_2G_Day2 = Convert.ToDouble(cell_data[1].ItemArray[13]);
                        double Availability_2G_Day1 = Convert.ToDouble(cell_data[0].ItemArray[13]);
                        if (Availability_2G_Day2 > 0 && Availability_2G_Day2 < TCH_Availability_2G)
                        {
                            if (Availability_2G_Day1 > 0)
                            {
                                if (100 * (Availability_2G_Day2 - Availability_2G_Day1) / Availability_2G_Day1 < -5)
                                {
                                    CR_Output_Table.Rows.Add(cell_data[1].ItemArray[2], cell_data[1].ItemArray[3], cell_data[1].ItemArray[1], Vendor, "2G", KPI_Name, Availability_2G_Day2, Availability_2G_Day1, "", Eng, CR, SD.Date, ED, Status, Description);
                                }
                            }
                            if (Availability_2G_Day1 == 0)
                            {
                                CR_Output_Table.Rows.Add(cell_data[1].ItemArray[2], cell_data[1].ItemArray[3], cell_data[1].ItemArray[1], Vendor, "2G", KPI_Name, Availability_2G_Day2, Availability_2G_Day1, "", Eng, CR, SD, ED, Status, Description);
                            }
                        }
                    }


                    if (cell_data[1].ItemArray[14].ToString() != "" && cell_data[0].ItemArray[14].ToString() != "")
                    {
                        string Vendor = cell_data[1].ItemArray[17].ToString();
                        string KPI_Name = "";
                        if (Vendor == "Ericsson")
                        {
                            KPI_Name = "SDCCH_Congestion_Rate(Eric_Cell)";
                        }
                        if (Vendor == "Huawei")
                        {
                            KPI_Name = "SDCCH_Congestion_Rate(HU_Cell)";
                        }
                        if (Vendor == "Nokia")
                        {
                            KPI_Name = "SDCCH_Congestion_Rate(Nokia_SEG)";
                        }
                        double SDCCH_Cong_2G_Day2 = Convert.ToDouble(cell_data[1].ItemArray[14]);
                        double SDCCH_Cong_2G_Day1 = Convert.ToDouble(cell_data[0].ItemArray[14]);
                        if (SDCCH_Cong_2G_Day2 > SDCCH_Cong_2G)
                        {
                            if (SDCCH_Cong_2G_Day1 > 0)
                            {
                                if (100 * (SDCCH_Cong_2G_Day2 - SDCCH_Cong_2G_Day1) / SDCCH_Cong_2G_Day1 > 30)
                                {
                                    CR_Output_Table.Rows.Add(cell_data[1].ItemArray[2], cell_data[1].ItemArray[3], cell_data[1].ItemArray[1], Vendor, "2G", KPI_Name, SDCCH_Cong_2G_Day2, SDCCH_Cong_2G_Day1, "", Eng, CR, SD, ED, Status, Description);
                                }
                            }
                            if (SDCCH_Cong_2G_Day1 == 0)
                            {
                                CR_Output_Table.Rows.Add(cell_data[1].ItemArray[2], cell_data[1].ItemArray[3], cell_data[1].ItemArray[1], Vendor, "2G", KPI_Name, SDCCH_Cong_2G_Day2, SDCCH_Cong_2G_Day1, "", Eng, CR, SD, ED, Status, Description);
                            }
                        }
                    }


                    if (cell_data[1].ItemArray[15].ToString() != "" && cell_data[0].ItemArray[15].ToString() != "")
                    {
                        string Vendor = cell_data[1].ItemArray[17].ToString();
                        string KPI_Name = "";
                        if (Vendor == "Ericsson")
                        {
                            KPI_Name = "RxQual_DL(Eric_cell)";
                        }
                        if (Vendor == "Huawei")
                        {
                            KPI_Name = "RX_QUAL_DL(HU_Cell)";
                        }
                        if (Vendor == "Nokia")
                        {
                            KPI_Name = "RxQuality_DL(Nokia_SEG)";
                        }
                        double RX_DL_2G_Day2 = Convert.ToDouble(cell_data[1].ItemArray[15]);
                        double RX_DL_2G_Day1 = Convert.ToDouble(cell_data[0].ItemArray[15]);
                        if (RX_DL_2G_Day2 > 0 && RX_DL_2G_Day2 < RX_DL_2G)
                        {
                            if (RX_DL_2G_Day1 > 0)
                            {
                                if (100 * (RX_DL_2G_Day2 - RX_DL_2G_Day1) / RX_DL_2G_Day1 < -5)
                                {
                                    CR_Output_Table.Rows.Add(cell_data[1].ItemArray[2], cell_data[1].ItemArray[3], cell_data[1].ItemArray[1], Vendor, "2G", KPI_Name, RX_DL_2G_Day2, RX_DL_2G_Day1, "", Eng, CR, SD, ED, Status, Description);
                                }
                            }
                            if (RX_DL_2G_Day1 == 0)
                            {
                                CR_Output_Table.Rows.Add(cell_data[1].ItemArray[2], cell_data[1].ItemArray[3], cell_data[1].ItemArray[1], Vendor, "2G", KPI_Name, RX_DL_2G_Day2, RX_DL_2G_Day1, "", Eng, CR, SD, ED, Status, Description);
                            }
                        }
                    }



                    if (cell_data[1].ItemArray[16].ToString() != "" && cell_data[0].ItemArray[16].ToString() != "")
                    {
                        string Vendor = cell_data[1].ItemArray[17].ToString();
                        string KPI_Name = "";
                        if (Vendor == "Ericsson")
                        {
                            KPI_Name = "RxQual_UL(Eric_cell)";
                        }
                        if (Vendor == "Huawei")
                        {
                            KPI_Name = "RX_QUAL_UL(HU_Cell)";
                        }
                        if (Vendor == "Nokia")
                        {
                            KPI_Name = "RxQuality_UL(Nokia_SEG)";
                        }
                        double RX_UL_2G_Day2 = Convert.ToDouble(cell_data[1].ItemArray[16]);
                        double RX_UL_2G_Day1 = Convert.ToDouble(cell_data[0].ItemArray[16]);
                        if (RX_UL_2G_Day2 > 0 && RX_UL_2G_Day2 < RX_UL_2G)
                        {
                            if (RX_UL_2G_Day1 > 0)
                            {
                                if (100 * (RX_UL_2G_Day2 - RX_UL_2G_Day1) / RX_UL_2G_Day1 < -5)
                                {
                                    CR_Output_Table.Rows.Add(cell_data[1].ItemArray[2], cell_data[1].ItemArray[3], cell_data[1].ItemArray[1], Vendor, "2G", KPI_Name, RX_UL_2G_Day2, RX_UL_2G_Day1, "", Eng, CR, SD, ED, Status, Description);
                                }
                            }
                            if (RX_UL_2G_Day1 == 0)
                            {
                                CR_Output_Table.Rows.Add(cell_data[1].ItemArray[2], cell_data[1].ItemArray[3], cell_data[1].ItemArray[1], Vendor, "2G", KPI_Name, RX_UL_2G_Day2, RX_UL_2G_Day1, "", Eng, CR, SD, ED, Status, Description);
                            }
                        }
                    }



                }
            }


            //    CR_Output_Table.Rows.Add("*****", "*****", "*****", "*****", "*****", "*****", "*****");

            for (int j = 0; j < distinctCells_3G_CS.Count; j++)
            {
                var cell_data = (from p in Data_Table_3G_CS.AsEnumerable()
                                 where p.Field<string>("Cell") == distinctCells_3G_CS[j].Cell
                                 select p).ToList();


                if (cell_data.Count == 2)
                {
                    string site_code1 = cell_data[0].ItemArray[2].ToString();
                    string site_code = site_code1.Substring(0, 2) + site_code1.Substring(4, 4);

                    List<int> index_finder = new List<int>();
                    index_finder = Site_list.Select((s, i) => new { i, s })
    .Where(t => t.s == site_code)
    .Select(t => t.i)
    .ToList();

                    int index = index_finder[0] + 2;
                    string Eng = Source_worksheet.Cell(index, 4).Value.ToString();
                    string CR = Source_worksheet.Cell(index, 5).Value.ToString();
                    DateTime SD = Convert.ToDateTime(Source_worksheet.Cell(index, 6).Value);
                    DateTime ED = Convert.ToDateTime(Source_worksheet.Cell(index, 7).Value);
                    //SD = SD.Date;
                    //ED = ED.Date;
                    string Status = Source_worksheet.Cell(index, 8).Value.ToString();
                    string Description = Source_worksheet.Cell(index, 9).Value.ToString();

                    if (cell_data[1].ItemArray[5].ToString() != "" && cell_data[0].ItemArray[5].ToString() != "")
                    {
                        string Vendor = cell_data[1].ItemArray[10].ToString();
                        string KPI_Name = "";
                        if (Vendor == "Ericsson")
                        {
                            KPI_Name = "Cs_RAB_Establish_Success_Rate(UCell_Eric)";
                        }
                        if (Vendor == "Huawei")
                        {
                            KPI_Name = "CS_RAB_Setup_Success_Ratio(Cell_Hu)";
                        }
                        if (Vendor == "Nokia")
                        {
                            KPI_Name = "CS_RAB_Establish_Success_Rate(Nokia_Cell)";
                        }

                        double CS_RAB_3G_Day2 = Convert.ToDouble(cell_data[1].ItemArray[5]);
                        double CS_RAB_3G_Day1 = Convert.ToDouble(cell_data[0].ItemArray[5]);
                        if (CS_RAB_3G_Day2 > 0 && CS_RAB_3G_Day2 < CS_RAB_3G)
                        {
                            if (CS_RAB_3G_Day1 > 0)
                            {
                                if (100 * (CS_RAB_3G_Day2 - CS_RAB_3G_Day1) / CS_RAB_3G_Day1 < -2)
                                {
                                    CR_Output_Table.Rows.Add(cell_data[1].ItemArray[2], cell_data[1].ItemArray[3], cell_data[1].ItemArray[1], Vendor, "3G", KPI_Name, CS_RAB_3G_Day2, CS_RAB_3G_Day1, "", Eng, CR, SD, ED, Status, Description);
                                }
                            }
                            if (CS_RAB_3G_Day1 == 0)
                            {
                                CR_Output_Table.Rows.Add(cell_data[1].ItemArray[2], cell_data[1].ItemArray[3], cell_data[1].ItemArray[1], Vendor, "3G", KPI_Name, CS_RAB_3G_Day2, CS_RAB_3G_Day1, "", Eng, CR, SD, ED, Status, Description);
                            }
                        }
                    }


                    if (cell_data[1].ItemArray[6].ToString() != "" && cell_data[0].ItemArray[6].ToString() != "")
                    {
                        string Vendor = cell_data[1].ItemArray[10].ToString();
                        string KPI_Name = "";
                        if (Vendor == "Ericsson")
                        {
                            KPI_Name = "CS_RRC_Setup_Success_Rate(UCell_Eric)";
                        }
                        if (Vendor == "Huawei")
                        {
                            KPI_Name = "CS_RRC_Setup_Success_Ratio(Cell.Service)(Cell_Hu)";
                        }
                        if (Vendor == "Nokia")
                        {
                            KPI_Name = "CS_RRC_SETUP_SR_WITHOUT_REPEAT(CELL_NOKIA)";
                        }
                        double CS_RRC_3G_Day2 = Convert.ToDouble(cell_data[1].ItemArray[6]);
                        double CS_RRC_3G_Day1 = Convert.ToDouble(cell_data[0].ItemArray[6]);
                        if (CS_RRC_3G_Day2 > 0 && CS_RRC_3G_Day2 < CS_RRC_3G)
                        {
                            if (CS_RRC_3G_Day1 > 0)
                            {
                                if (100 * (CS_RRC_3G_Day2 - CS_RRC_3G_Day1) / CS_RRC_3G_Day1 < -2)
                                {
                                    CR_Output_Table.Rows.Add(cell_data[1].ItemArray[2], cell_data[1].ItemArray[3], cell_data[1].ItemArray[1], Vendor, "3G", KPI_Name, CS_RRC_3G_Day2, CS_RRC_3G_Day1, "", Eng, CR, SD, ED, Status, Description);
                                }
                            }
                            if (CS_RRC_3G_Day1 == 0)
                            {
                                CR_Output_Table.Rows.Add(cell_data[1].ItemArray[2], cell_data[1].ItemArray[3], cell_data[1].ItemArray[1], Vendor, "3G", KPI_Name, CS_RRC_3G_Day2, CS_RRC_3G_Day1, "", Eng, CR, SD, ED, Status, Description);
                            }
                        }
                    }

                    if (cell_data[1].ItemArray[7].ToString() != "" && cell_data[0].ItemArray[7].ToString() != "")
                    {
                        string Vendor = cell_data[1].ItemArray[10].ToString();
                        string KPI_Name = "";
                        if (Vendor == "Ericsson")
                        {
                            KPI_Name = "Voice_Drop_Call_Rate(UCell_Eric)";
                        }
                        if (Vendor == "Huawei")
                        {
                            KPI_Name = "AMR_Call_Drop_Ratio_New(Hu_CELL)";
                        }
                        if (Vendor == "Nokia")
                        {
                            KPI_Name = "CS_Drop_Call_Rate(Nokia_CELL)";
                        }
                        double CS_Drop_3G_Day2 = Convert.ToDouble(cell_data[1].ItemArray[7]);
                        double CS_Drop_3G_Day1 = Convert.ToDouble(cell_data[0].ItemArray[7]);
                        if (CS_Drop_3G_Day2 > CS_Drop_3G)
                        {
                            if (CS_Drop_3G_Day1 > 0)
                            {
                                if (100 * (CS_Drop_3G_Day2 - CS_Drop_3G_Day1) / CS_Drop_3G_Day1 > 30)
                                {
                                    CR_Output_Table.Rows.Add(cell_data[1].ItemArray[2], cell_data[1].ItemArray[3], cell_data[1].ItemArray[1], Vendor, "3G", KPI_Name, CS_Drop_3G_Day2, CS_Drop_3G_Day1, "", Eng, CR, SD, ED, Status, Description);
                                }
                            }
                            if (CS_Drop_3G_Day1 == 0)
                            {
                                CR_Output_Table.Rows.Add(cell_data[1].ItemArray[2], cell_data[1].ItemArray[3], cell_data[1].ItemArray[1], Vendor, "3G", KPI_Name, CS_Drop_3G_Day2, CS_Drop_3G_Day1, "", Eng, CR, SD, ED, Status, Description);
                            }
                        }
                    }



                    if (cell_data[1].ItemArray[9].ToString() != "" && cell_data[0].ItemArray[9].ToString() != "")
                    {
                        string Vendor = cell_data[1].ItemArray[10].ToString();
                        string KPI_Name = "";
                        if (Vendor == "Ericsson")
                        {
                            KPI_Name = "Soft_HO_Suc_Rate(UCell_Eric)";
                        }
                        if (Vendor == "Huawei")
                        {
                            KPI_Name = "Soft_Handover_Succ_Rate(Hu_Cell)";
                        }
                        if (Vendor == "Nokia")
                        {
                            KPI_Name = "Soft_HO_SR(RT+NRT)(Nokia_Cell)";
                        }
                        double Soft_HO_3G_Day2 = Convert.ToDouble(cell_data[1].ItemArray[9]);
                        double Soft_HO_3G_Day1 = Convert.ToDouble(cell_data[0].ItemArray[9]);
                        if (Soft_HO_3G_Day2 > 0 && Soft_HO_3G_Day2 < Soft_HO_3G)
                        {
                            if (Soft_HO_3G_Day1 > 0)
                            {
                                if (100 * (Soft_HO_3G_Day2 - Soft_HO_3G_Day1) / Soft_HO_3G_Day1 < -5)
                                {
                                    CR_Output_Table.Rows.Add(cell_data[1].ItemArray[2], cell_data[1].ItemArray[3], cell_data[1].ItemArray[1], Vendor, "3G", KPI_Name, Soft_HO_3G_Day2, Soft_HO_3G_Day1, "", Eng, CR, SD, ED, Status, Description);
                                }
                            }
                            if (Soft_HO_3G_Day1 == 0)
                            {
                                CR_Output_Table.Rows.Add(cell_data[1].ItemArray[2], cell_data[1].ItemArray[3], cell_data[1].ItemArray[1], Vendor, "3G", KPI_Name, Soft_HO_3G_Day2, Soft_HO_3G_Day1, "", Eng, CR, SD, ED, Status, Description);
                            }
                        }
                    }






                }

            }




            for (int j = 0; j < distinctCells_3G_PS.Count; j++)
            {
                var cell_data = (from p in Data_Table_3G_PS.AsEnumerable()
                                 where p.Field<string>("Cell") == distinctCells_3G_PS[j].Cell
                                 select p).ToList();




                if (cell_data.Count == 2)
                {


                    string site_code1 = cell_data[0].ItemArray[2].ToString();
                    string site_code = site_code1.Substring(0, 2) + site_code1.Substring(4, 4);

                    List<int> index_finder = new List<int>();
                    index_finder = Site_list.Select((s, i) => new { i, s })
    .Where(t => t.s == site_code)
    .Select(t => t.i)
    .ToList();

                    int index = index_finder[0] + 2;
                    string Eng = Source_worksheet.Cell(index, 4).Value.ToString();
                    string CR = Source_worksheet.Cell(index, 5).Value.ToString();
                    DateTime SD = Convert.ToDateTime(Source_worksheet.Cell(index, 6).Value);
                    DateTime ED = Convert.ToDateTime(Source_worksheet.Cell(index, 7).Value);
                    //SD = SD.Date;
                    //ED = ED.Date;
                    string Status = Source_worksheet.Cell(index, 8).Value.ToString();
                    string Description = Source_worksheet.Cell(index, 9).Value.ToString();


                    if (cell_data[1].ItemArray[5].ToString() != "" && cell_data[0].ItemArray[5].ToString() != "")
                    {
                        string Vendor = cell_data[1].ItemArray[10].ToString();
                        string KPI_Name = "";
                        if (Vendor == "Ericsson")
                        {
                            KPI_Name = "Ps_RAB_Establish_Success_Rate(UCell_Eric)";
                        }
                        if (Vendor == "Huawei")
                        {
                            KPI_Name = "PS_RAB_Setup_Success_Ratio(Hu_Cell)";
                        }
                        if (Vendor == "Nokia")
                        {
                            KPI_Name = "RAB_Setup_and_Access_Failure_count_NRT_Service_from_User_perspective";
                        }

                        double PS_RAB_3G_Day2 = Convert.ToDouble(cell_data[1].ItemArray[5]);
                        double PS_RAB_3G_Day1 = Convert.ToDouble(cell_data[0].ItemArray[5]);
                        if (PS_RAB_3G_Day2 > 0 && PS_RAB_3G_Day2 < PS_RAB_3G)
                        {
                            if (PS_RAB_3G_Day1 > 0)
                            {
                                if (100 * (PS_RAB_3G_Day2 - PS_RAB_3G_Day1) / PS_RAB_3G_Day1 < -2)
                                {
                                    CR_Output_Table.Rows.Add(cell_data[1].ItemArray[2], cell_data[1].ItemArray[3], cell_data[1].ItemArray[1], Vendor, "3G", KPI_Name, PS_RAB_3G_Day2, PS_RAB_3G_Day1, "", Eng, CR, SD, ED, Status, Description);
                                }
                            }
                            if (PS_RAB_3G_Day1 == 0)
                            {
                                CR_Output_Table.Rows.Add(cell_data[1].ItemArray[2], cell_data[1].ItemArray[3], cell_data[1].ItemArray[1], Vendor, "3G", KPI_Name, PS_RAB_3G_Day2, PS_RAB_3G_Day1, "", Eng, CR, SD, ED, Status, Description);
                            }
                        }
                    }


                    if (cell_data[1].ItemArray[6].ToString() != "" && cell_data[0].ItemArray[6].ToString() != "")
                    {
                        string Vendor = cell_data[1].ItemArray[10].ToString();
                        string KPI_Name = "";
                        if (Vendor == "Ericsson")
                        {
                            KPI_Name = "PS_RRC_Setup_Success_Rate(UCell_Eric)";
                        }
                        if (Vendor == "Huawei")
                        {
                            KPI_Name = "PS_RRC_Connection_success_Rate(Hu_Cell)";
                        }
                        if (Vendor == "Nokia")
                        {
                            KPI_Name = "PS_RRCSETUP_SR(Nokia_CELL)";
                        }

                        double PS_RRC_3G_Day2 = Convert.ToDouble(cell_data[1].ItemArray[6]);
                        double PS_RRC_3G_Day1 = Convert.ToDouble(cell_data[0].ItemArray[6]);
                        if (PS_RRC_3G_Day2 > 0 && PS_RRC_3G_Day2 < PS_RRC_3G)
                        {
                            if (PS_RRC_3G_Day1 > 0)
                            {
                                if (100 * (PS_RRC_3G_Day2 - PS_RRC_3G_Day1) / PS_RRC_3G_Day1 < -2)
                                {
                                    CR_Output_Table.Rows.Add(cell_data[1].ItemArray[2], cell_data[1].ItemArray[3], cell_data[1].ItemArray[1], Vendor, "3G", KPI_Name, PS_RRC_3G_Day2, PS_RRC_3G_Day1, "", Eng, CR, SD, ED, Status, Description);
                                }
                            }
                            if (PS_RRC_3G_Day1 == 0)
                            {
                                CR_Output_Table.Rows.Add(cell_data[1].ItemArray[2], cell_data[1].ItemArray[3], cell_data[1].ItemArray[1], Vendor, "3G", KPI_Name, PS_RRC_3G_Day2, PS_RRC_3G_Day1, "", Eng, CR, SD, ED, Status, Description);
                            }
                        }
                    }


                    if (cell_data[1].ItemArray[7].ToString() != "" && cell_data[0].ItemArray[7].ToString() != "")
                    {
                        string Vendor = cell_data[1].ItemArray[10].ToString();
                        string KPI_Name = "";
                        if (Vendor == "Ericsson")
                        {
                            KPI_Name = "PacketHs_RW18(Eric_Cell)";
                        }
                        if (Vendor == "Huawei")
                        {
                            KPI_Name = "PS_Call_Drop_Ratio(Hu_Cell)";
                        }
                        if (Vendor == "Nokia")
                        {
                            KPI_Name = "PS_RAB_drop_rate_cell(Nokia_CELL)";
                        }

                        double PS_Drop_3G_Day2 = Convert.ToDouble(cell_data[1].ItemArray[7]);
                        double PS_Drop_3G_Day1 = Convert.ToDouble(cell_data[0].ItemArray[7]);
                        if (PS_Drop_3G_Day2 > PS_Drop_3G)
                        {
                            if (PS_Drop_3G_Day1 > 0)
                            {
                                if (100 * (PS_Drop_3G_Day2 - PS_Drop_3G_Day1) / PS_Drop_3G_Day1 > 30)
                                {
                                    CR_Output_Table.Rows.Add(cell_data[1].ItemArray[2], cell_data[1].ItemArray[3], cell_data[1].ItemArray[1], Vendor, "3G", KPI_Name, PS_Drop_3G_Day2, PS_Drop_3G_Day1, "", Eng, CR, SD, ED, Status, Description);
                                }
                            }
                            if (PS_Drop_3G_Day1 == 0)
                            {
                                CR_Output_Table.Rows.Add(cell_data[1].ItemArray[2], cell_data[1].ItemArray[3], cell_data[1].ItemArray[1], Vendor, "3G", KPI_Name, PS_Drop_3G_Day2, PS_Drop_3G_Day1, "", Eng, CR, SD, ED, Status, Description);
                            }
                        }
                    }


                    if (cell_data[1].ItemArray[8].ToString() != "" && cell_data[0].ItemArray[8].ToString() != "")
                    {
                        string Vendor = cell_data[1].ItemArray[10].ToString();
                        string KPI_Name = "";
                        if (Vendor == "Ericsson")
                        {
                            KPI_Name = "HS_USER_Throughput_NET_PQ(Mbps)(UCell_Eric)";
                        }
                        if (Vendor == "Huawei")
                        {
                            KPI_Name = "AVERAGE_HSDPA_USER_THROUGHPUT_SC_ONLY(Kbit/s)(CELL_HUAWEI)";
                        }
                        if (Vendor == "Nokia")
                        {
                            KPI_Name = "Mac_hs_throughput_2msec_TTI_Mbps(Nokia_CELL)";
                        }
                        double User_THR_3G_Day2 = Convert.ToDouble(cell_data[1].ItemArray[8]);
                        double User_THR_3G_Day1 = Convert.ToDouble(cell_data[0].ItemArray[8]);
                        if (User_THR_3G_Day2 > 0 && User_THR_3G_Day2 < User_THR_3G)
                        {
                            if (User_THR_3G_Day1 > 0)
                            {
                                if (100 * (User_THR_3G_Day2 - User_THR_3G_Day1) / User_THR_3G_Day1 < -30)
                                {
                                    CR_Output_Table.Rows.Add(cell_data[1].ItemArray[2], cell_data[1].ItemArray[3], cell_data[1].ItemArray[1], Vendor, "3G", KPI_Name, User_THR_3G_Day2, User_THR_3G_Day1, "", Eng, CR, SD, ED, Status, Description);
                                }
                            }
                            if (User_THR_3G_Day1 == 0)
                            {
                                CR_Output_Table.Rows.Add(cell_data[1].ItemArray[2], cell_data[1].ItemArray[3], cell_data[1].ItemArray[1], Vendor, "3G", KPI_Name, User_THR_3G_Day2, User_THR_3G_Day1, "", Eng, CR, SD, ED, Status, Description);
                            }
                        }
                    }


                    if (cell_data[1].ItemArray[9].ToString() != "" && cell_data[0].ItemArray[9].ToString() != "")
                    {
                        string Vendor = cell_data[1].ItemArray[10].ToString();
                        string KPI_Name = "";
                        if (Vendor == "Ericsson")
                        {
                            KPI_Name = "HSDPA_Cell_Scheduled_Throughput(mbps)(UCell_Eric)";
                        }
                        if (Vendor == "Huawei")
                        {
                            KPI_Name = "HSDPA_SCHEDULING_CELL_THROUGHPUT(Mbit/s)(CELL_HUAWEI)";
                        }
                        if (Vendor == "Nokia")
                        {
                            KPI_Name = "AVERAGE_HSDPA_SCHEDULING_CELL_THROUGHPUT(Mbit/s)(WCELL_nokia)";
                        }
                        double Cell_THR_3G_Day2 = Convert.ToDouble(cell_data[1].ItemArray[9]);
                        double Cell_THR_3G_Day1 = Convert.ToDouble(cell_data[0].ItemArray[9]);
                        if (Cell_THR_3G_Day2 > 0 && Cell_THR_3G_Day2 < Cell_THR_3G)
                        {
                            if (Cell_THR_3G_Day1 > 0)
                            {
                                if (100 * (Cell_THR_3G_Day2 - Cell_THR_3G_Day1) / Cell_THR_3G_Day1 < -30)
                                {
                                    CR_Output_Table.Rows.Add(cell_data[1].ItemArray[2], cell_data[1].ItemArray[3], cell_data[1].ItemArray[1], Vendor, "3G", KPI_Name, Cell_THR_3G_Day2, Cell_THR_3G_Day1, "", Eng, CR, SD, ED, Status, Description);
                                }
                            }
                            if (Cell_THR_3G_Day1 == 0)
                            {
                                CR_Output_Table.Rows.Add(cell_data[1].ItemArray[2], cell_data[1].ItemArray[3], cell_data[1].ItemArray[1], Vendor, "3G", KPI_Name, Cell_THR_3G_Day2, Cell_THR_3G_Day1, "", Eng, CR, SD, ED, Status, Description);
                            }
                        }
                    }






                }
            }


            // CR_Output_Table.Rows.Add("#####", "#####", "#####", "#####", "#####", "#####", "#####");

            for (int j = 0; j < distinctCells_4G.Count; j++)
            {
                var cell_data = (from p in Data_Table_4G.AsEnumerable()
                                 where p.Field<string>("Cell") == distinctCells_4G[j].Cell
                                 select p).ToList();


                if (cell_data.Count == 2)
                {


                    string site_code1 = cell_data[0].ItemArray[2].ToString();
                    string site_code = site_code1.Substring(0, 2) + site_code1.Substring(4, 4);

                    List<int> index_finder = new List<int>();
                    index_finder = Site_list.Select((s, i) => new { i, s })
    .Where(t => t.s == site_code)
    .Select(t => t.i)
    .ToList();

                    int index = index_finder[0] + 2;
                    string Eng = Source_worksheet.Cell(index, 4).Value.ToString();
                    string CR = Source_worksheet.Cell(index, 5).Value.ToString();
                    DateTime SD = Convert.ToDateTime(Source_worksheet.Cell(index, 6).Value);
                    DateTime ED = Convert.ToDateTime(Source_worksheet.Cell(index, 7).Value);
                    //SD = SD.Date;
                    //ED = ED.Date;
                    string Status = Source_worksheet.Cell(index, 8).Value.ToString();
                    string Description = Source_worksheet.Cell(index, 9).Value.ToString();


                    if (cell_data[1].ItemArray[4].ToString() != "" && cell_data[0].ItemArray[4].ToString() != "")
                    {
                        string Vendor = cell_data[1].ItemArray[14].ToString();
                        string KPI_Name = "";
                        if (Vendor == "Ericsson")
                        {
                            KPI_Name = "Average_UE_DL_Throughput(Mbps)(EUCell_Eric)";
                        }
                        if (Vendor == "Huawei")
                        {
                            KPI_Name = "Average_Downlink_User_Throughput(Mbit/s)(Hu_Cell)";
                        }
                        if (Vendor == "Nokia")
                        {
                            KPI_Name = "User_Throughput_DL_mbps(nokia_LTE_CELL)";
                        }
                        double UE_DL_4G_Day2 = Convert.ToDouble(cell_data[1].ItemArray[4]);
                        double UE_DL_4G_Day1 = Convert.ToDouble(cell_data[0].ItemArray[4]);
                        if (UE_DL_4G_Day2 > 0 && UE_DL_4G_Day2 < UE_DL_4G)
                        {
                            if (UE_DL_4G_Day1 > 0)
                            {
                                if (100 * (UE_DL_4G_Day2 - UE_DL_4G_Day1) / UE_DL_4G_Day1 < -30)
                                {
                                    CR_Output_Table.Rows.Add(cell_data[1].ItemArray[1], cell_data[1].ItemArray[2], "", Vendor, "4G", KPI_Name, UE_DL_4G_Day2, UE_DL_4G_Day1, "", Eng, CR, SD, ED, Status, Description);
                                }
                            }
                            if (UE_DL_4G_Day1 == 0)
                            {
                                CR_Output_Table.Rows.Add(cell_data[1].ItemArray[1], cell_data[1].ItemArray[2], "", Vendor, "4G", KPI_Name, UE_DL_4G_Day2, UE_DL_4G_Day1, "", Eng, CR, SD, ED, Status, Description);
                            }
                        }
                    }

                    if (cell_data[1].ItemArray[5].ToString() != "" && cell_data[0].ItemArray[5].ToString() != "")
                    {
                        string Vendor = cell_data[1].ItemArray[14].ToString();
                        string KPI_Name = "";
                        if (Vendor == "Ericsson")
                        {
                            KPI_Name = "Average_UE_UL_Throughput(Mbps)(EUCell_Eric)";
                        }
                        if (Vendor == "Huawei")
                        {
                            KPI_Name = "Average_UPlink_User_Throughput(Mbit/s)(Hu_Cell)";
                        }
                        if (Vendor == "Nokia")
                        {
                            KPI_Name = "User_Throughput_UL_mbps(Nokia_LTE_CELL)";
                        }
                        double UE_UL_4G_Day2 = Convert.ToDouble(cell_data[1].ItemArray[5]);
                        double UE_UL_4G_Day1 = Convert.ToDouble(cell_data[0].ItemArray[5]);
                        if (UE_UL_4G_Day2 > 0 && UE_UL_4G_Day2 < UE_UL_4G)
                        {
                            if (UE_UL_4G_Day1 > 0)
                            {
                                if (100 * (UE_UL_4G_Day2 - UE_UL_4G_Day1) / UE_UL_4G_Day1 < -50)
                                {
                                    CR_Output_Table.Rows.Add(cell_data[1].ItemArray[1], cell_data[1].ItemArray[2], "", Vendor, "4G", KPI_Name, UE_UL_4G_Day2, UE_UL_4G_Day1, "", Eng, CR, SD, ED, Status, Description);
                                }
                            }
                            if (UE_UL_4G_Day1 == 0)
                            {
                                CR_Output_Table.Rows.Add(cell_data[1].ItemArray[1], cell_data[1].ItemArray[2], "", Vendor, "4G", KPI_Name, UE_UL_4G_Day2, UE_UL_4G_Day1, "", Eng, CR, SD, ED, Status, Description);
                            }
                        }
                    }


                    if (cell_data[1].ItemArray[6].ToString() != "" && cell_data[0].ItemArray[6].ToString() != "")
                    {
                        string Vendor = cell_data[1].ItemArray[14].ToString();
                        string KPI_Name = "";
                        if (Vendor == "Ericsson")
                        {
                            KPI_Name = "E_RAB_Drop_Rate(EUCell_Eric)";
                        }
                        if (Vendor == "Huawei")
                        {
                            KPI_Name = "Call_Drop_Rate(Hu_Cell)";
                        }
                        if (Vendor == "Nokia")
                        {
                            KPI_Name = "E-RAB_Drop_New(Nokia_LTE_Cell)";
                        }
                        double ERAB_Drop_4G_Day2 = Convert.ToDouble(cell_data[1].ItemArray[6]);
                        double ERAB_Drop_4G_Day1 = Convert.ToDouble(cell_data[0].ItemArray[6]);
                        if (ERAB_Drop_4G_Day2 > ERAB_Drop_4G)
                        {
                            if (ERAB_Drop_4G_Day1 > 0)
                            {
                                if (100 * (ERAB_Drop_4G_Day2 - ERAB_Drop_4G_Day1) / ERAB_Drop_4G_Day1 > 30)
                                {
                                    CR_Output_Table.Rows.Add(cell_data[1].ItemArray[1], cell_data[1].ItemArray[2], "", Vendor, "4G", KPI_Name, ERAB_Drop_4G_Day2, ERAB_Drop_4G_Day1, "", Eng, CR, SD, ED, Status, Description);
                                }
                            }
                            if (ERAB_Drop_4G_Day1 == 0)
                            {
                                CR_Output_Table.Rows.Add(cell_data[1].ItemArray[1], cell_data[1].ItemArray[2], "", Vendor, "4G", KPI_Name, ERAB_Drop_4G_Day2, ERAB_Drop_4G_Day1, "", Eng, CR, SD, ED, Status, Description);
                            }
                        }
                    }


                    if (cell_data[1].ItemArray[7].ToString() != "" && cell_data[0].ItemArray[7].ToString() != "")
                    {
                        string Vendor = cell_data[1].ItemArray[14].ToString();
                        string KPI_Name = "";
                        if (Vendor == "Ericsson")
                        {
                            KPI_Name = "E-RAB_Setup_SR_incl_added_New(EUCell_Eric)";
                        }
                        if (Vendor == "Huawei")
                        {
                            KPI_Name = "E-RAB_Setup_Success_Rate(Hu_Cell)";
                        }
                        if (Vendor == "Nokia")
                        {
                            KPI_Name = "E-RAB_Setup_SR_incl_added(Nokia_LTE_CELL)";
                        }
                        double ERAB_Setup_4G_Day2 = Convert.ToDouble(cell_data[1].ItemArray[7]);
                        double ERAB_Setup_4G_Day1 = Convert.ToDouble(cell_data[0].ItemArray[7]);
                        if (ERAB_Setup_4G_Day2 > 0 && ERAB_Setup_4G_Day2 < ERAB_Setup_4G)
                        {
                            if (ERAB_Setup_4G_Day1 > 0)
                            {
                                if (100 * (ERAB_Setup_4G_Day2 - ERAB_Setup_4G_Day1) / ERAB_Setup_4G_Day1 < -5)
                                {
                                    CR_Output_Table.Rows.Add(cell_data[1].ItemArray[1], cell_data[1].ItemArray[2], "", Vendor, "4G", KPI_Name, ERAB_Setup_4G_Day2, ERAB_Setup_4G_Day1, "", Eng, CR, SD, ED, Status, Description);
                                }
                            }
                            if (ERAB_Setup_4G_Day1 == 0)
                            {
                                CR_Output_Table.Rows.Add(cell_data[1].ItemArray[1], cell_data[1].ItemArray[2], "", Vendor, "4G", KPI_Name, ERAB_Setup_4G_Day2, ERAB_Setup_4G_Day1, "", Eng, CR, SD, ED, Status, Description);
                            }
                        }
                    }


                    if (cell_data[1].ItemArray[8].ToString() != "" && cell_data[0].ItemArray[8].ToString() != "")
                    {
                        string Vendor = cell_data[1].ItemArray[14].ToString();
                        string KPI_Name = "";
                        if (Vendor == "Ericsson")
                        {
                            KPI_Name = "IntraF_Handover_Execution_Rate(EUCell_Eric)";
                        }
                        if (Vendor == "Huawei")
                        {
                            KPI_Name = "IntraF_HOOut_SR(Cell_HuLTE)";
                        }
                        if (Vendor == "Nokia")
                        {
                            KPI_Name = "Inter-Freq_HO_SR(Nokia_LTE_CELL)";
                        }
                        double IntraFreq_HO_4G_Day2 = Convert.ToDouble(cell_data[1].ItemArray[8]);
                        double IntraFreq_HO_4G_Day1 = Convert.ToDouble(cell_data[0].ItemArray[8]);
                        if (IntraFreq_HO_4G_Day2 > 0 && IntraFreq_HO_4G_Day2 < IntraFreq_HO_4G)
                        {
                            if (IntraFreq_HO_4G_Day1 > 0)
                            {
                                if (100 * (IntraFreq_HO_4G_Day2 - IntraFreq_HO_4G_Day1) / IntraFreq_HO_4G_Day1 < -5)
                                {
                                    CR_Output_Table.Rows.Add(cell_data[1].ItemArray[1], cell_data[1].ItemArray[2], "", Vendor, "4G", KPI_Name, IntraFreq_HO_4G_Day2, IntraFreq_HO_4G_Day1, "", Eng, CR, SD, ED, Status, Description);
                                }
                            }
                            if (IntraFreq_HO_4G_Day1 == 0)
                            {
                                CR_Output_Table.Rows.Add(cell_data[1].ItemArray[1], cell_data[1].ItemArray[2], "", Vendor, "4G", KPI_Name, IntraFreq_HO_4G_Day2, IntraFreq_HO_4G_Day1, "", Eng, CR, SD, ED, Status, Description);
                            }
                        }
                    }


                    if (cell_data[1].ItemArray[9].ToString() != "" && cell_data[0].ItemArray[9].ToString() != "")
                    {
                        string Vendor = cell_data[1].ItemArray[14].ToString();
                        string KPI_Name = "";
                        if (Vendor == "Ericsson")
                        {
                            KPI_Name = "RRC_Connection_Setup_Success_Rate(EUCell_Eric)";
                        }
                        if (Vendor == "Huawei")
                        {
                            KPI_Name = "RRC_Connection_Setup_Success_Rate_service(Hu_Cell)";
                        }
                        if (Vendor == "Nokia")
                        {
                            KPI_Name = "RRC_Connection_Setup_Success_Ratio(Nokia_LTE_CELL)";
                        }
                        double RRC_SR_4G_Day2 = Convert.ToDouble(cell_data[1].ItemArray[9]);
                        double RRC_SR_4G_Day1 = Convert.ToDouble(cell_data[0].ItemArray[9]);
                        if (RRC_SR_4G_Day2 > 0 && RRC_SR_4G_Day2 < RRC_SR_4G)
                        {
                            if (RRC_SR_4G_Day1 > 0)
                            {
                                if (100 * (RRC_SR_4G_Day2 - RRC_SR_4G_Day1) / RRC_SR_4G_Day1 < -5)
                                {
                                    CR_Output_Table.Rows.Add(cell_data[1].ItemArray[1], cell_data[1].ItemArray[2], "", Vendor, "4G", KPI_Name, RRC_SR_4G_Day2, RRC_SR_4G_Day1, "", Eng, CR, SD, ED, Status, Description);
                                }
                            }
                            if (RRC_SR_4G_Day1 == 0)
                            {
                                CR_Output_Table.Rows.Add(cell_data[1].ItemArray[1], cell_data[1].ItemArray[2], "", Vendor, "4G", KPI_Name, RRC_SR_4G_Day2, RRC_SR_4G_Day1, "", Eng, CR, SD, ED, Status, Description);
                            }
                        }
                    }


                    if (cell_data[1].ItemArray[11].ToString() != "" && cell_data[0].ItemArray[11].ToString() != "")
                    {
                        string Vendor = cell_data[1].ItemArray[14].ToString();
                        string KPI_Name = "";
                        if (Vendor == "Ericsson")
                        {
                            KPI_Name = "Average_UE_DL_Latency(ms)(EUCell_Eric)";
                        }
                        if (Vendor == "Huawei")
                        {
                            KPI_Name = "Average_DL_Latency_ms(Huawei_LTE_EUCell)";
                        }
                        if (Vendor == "Nokia")
                        {
                            KPI_Name = "Average_Latency_DL_ms(Nokia_LTE_CELL)";
                        }
                        double DL_Latency_4G_Day2 = Convert.ToDouble(cell_data[1].ItemArray[11]);
                        double DL_Latency_4G_Day1 = Convert.ToDouble(cell_data[0].ItemArray[11]);
                        if (DL_Latency_4G_Day2 > DL_Latency_4G)
                        {
                            if (DL_Latency_4G_Day1 > 0)
                            {
                                if (100 * (DL_Latency_4G_Day2 - DL_Latency_4G_Day1) / DL_Latency_4G_Day1 > 50)
                                {
                                    CR_Output_Table.Rows.Add(cell_data[1].ItemArray[1], cell_data[1].ItemArray[2], "", Vendor, "4G", KPI_Name, DL_Latency_4G_Day2, DL_Latency_4G_Day1, "", Eng, CR, SD, ED, Status, Description);
                                }
                            }
                            if (DL_Latency_4G_Day1 == 0)
                            {
                                CR_Output_Table.Rows.Add(cell_data[1].ItemArray[1], cell_data[1].ItemArray[2], "", Vendor, "4G", KPI_Name, DL_Latency_4G_Day2, DL_Latency_4G_Day1, "", Eng, CR, SD, ED, Status, Description);
                            }
                        }
                    }



                    if (cell_data[1].ItemArray[12].ToString() != "" && cell_data[0].ItemArray[12].ToString() != "")
                    {
                        string Vendor = cell_data[1].ItemArray[14].ToString();
                        string KPI_Name = "";
                        if (Vendor == "Ericsson")
                        {
                            KPI_Name = "Average_UE_Ul_Packet_Loss_Rate(EUCell_Eric)";
                        }
                        if (Vendor == "Huawei")
                        {
                            KPI_Name = "Average_UL_Packet_Loss_%(Huawei_LTE_UCell)";
                        }
                        if (Vendor == "Nokia")
                        {
                            KPI_Name = "Packet_loss_UL(Nokia_EUCELL)";
                        }
                        double UL_Packet_Loss_4G_Day2 = Convert.ToDouble(cell_data[1].ItemArray[12]);
                        double UL_Packet_Loss_4G_Day1 = Convert.ToDouble(cell_data[0].ItemArray[12]);
                        if (UL_Packet_Loss_4G_Day2 > UL_Packet_Loss_4G)
                        {
                            if (UL_Packet_Loss_4G_Day1 > 0)
                            {
                                if (100 * (UL_Packet_Loss_4G_Day2 - UL_Packet_Loss_4G_Day1) / UL_Packet_Loss_4G_Day1 > 30)
                                {
                                    CR_Output_Table.Rows.Add(cell_data[1].ItemArray[1], cell_data[1].ItemArray[2], "", Vendor, "4G", KPI_Name, UL_Packet_Loss_4G_Day2, UL_Packet_Loss_4G_Day1, "", Eng, CR, SD, ED, Status, Description);
                                }
                            }
                            if (UL_Packet_Loss_4G_Day1 == 0)
                            {
                                CR_Output_Table.Rows.Add(cell_data[1].ItemArray[1], cell_data[1].ItemArray[2], "", Vendor, "4G", KPI_Name, UL_Packet_Loss_4G_Day2, UL_Packet_Loss_4G_Day1, "", Eng, CR, SD, ED, Status, Description);
                            }
                        }
                    }



                    if (cell_data[1].ItemArray[13].ToString() != "" && cell_data[0].ItemArray[13].ToString() != "")
                    {
                        string Vendor = cell_data[1].ItemArray[14].ToString();
                        string KPI_Name = "";
                        if (Vendor == "Ericsson")
                        {
                            KPI_Name = "LTE_Service_Setup_SR(EUCell_Eric)";
                        }
                        if (Vendor == "Huawei")
                        {
                            KPI_Name = "CSSR(ALL)(Hu_Cell)";
                        }
                        if (Vendor == "Nokia")
                        {
                            KPI_Name = "Initial_E-RAB_Accessibility(Nokia_LTE_CELL)";
                        }
                        double CSSR_4G_Day2 = Convert.ToDouble(cell_data[1].ItemArray[13]);
                        double CSSR_4G_Day1 = Convert.ToDouble(cell_data[0].ItemArray[13]);
                        if (CSSR_4G_Day2 > 0 && CSSR_4G_Day2 < CSSR_4G)
                        {
                            if (CSSR_4G_Day1 > 0)
                            {
                                if (100 * (CSSR_4G_Day2 - CSSR_4G_Day1) / CSSR_4G_Day1 < -5)
                                {
                                    CR_Output_Table.Rows.Add(cell_data[1].ItemArray[1], cell_data[1].ItemArray[2], "", Vendor, "4G", KPI_Name, CSSR_4G_Day2, CSSR_4G_Day1, "", Eng, CR, SD, ED, Status, Description);
                                }
                            }
                            if (CSSR_4G_Day1 == 0)
                            {
                                CR_Output_Table.Rows.Add(cell_data[1].ItemArray[1], cell_data[1].ItemArray[2], "", Vendor, "4G", KPI_Name, CSSR_4G_Day2, CSSR_4G_Day1, "", Eng, CR, SD, ED, Status, Description);
                            }
                        }
                    }






                }
            }





            dataGridView1.RowCount = CR_Output_Table.Rows.Count + 1;

            dataGridView1.Rows[0].Cells[0].Value = "Site"; dataGridView1.Columns[0].Width = 70;
            dataGridView1.Rows[0].Cells[1].Value = "Cell"; dataGridView1.Columns[1].Width = 80;
            dataGridView1.Rows[0].Cells[2].Value = "Node"; dataGridView1.Columns[2].Width = 60;
            dataGridView1.Rows[0].Cells[3].Value = "Vendor"; dataGridView1.Columns[3].Width = 60;
            dataGridView1.Rows[0].Cells[4].Value = "Tech"; dataGridView1.Columns[4].Width = 40;
            dataGridView1.Rows[0].Cells[5].Value = "KPI"; dataGridView1.Columns[5].Width = 130;
            dataGridView1.Rows[0].Cells[6].Value = "Vlaue After CR"; dataGridView1.Columns[6].Width = 95;
            dataGridView1.Rows[0].Cells[7].Value = "Vlaue Before CR"; dataGridView1.Columns[7].Width = 95;
            dataGridView1.Rows[0].Cells[8].Value = "Performance Comment"; dataGridView1.Columns[8].Width = 250;
            dataGridView1.Rows[0].Cells[9].Value = "Owner"; dataGridView1.Columns[9].Width = 80;
            dataGridView1.Rows[0].Cells[10].Value = "CR Num"; dataGridView1.Columns[10].Width = 80;
            dataGridView1.Rows[0].Cells[11].Value = "Start Time"; dataGridView1.Columns[11].Width = 95;
            dataGridView1.Rows[0].Cells[12].Value = "End Time"; dataGridView1.Columns[12].Width = 95;
            dataGridView1.Rows[0].Cells[13].Value = "CR Statu"; dataGridView1.Columns[13].Width = 80;
            dataGridView1.Rows[0].Cells[14].Value = "CR Description"; dataGridView1.Columns[14].Width = 150;






            for (int k = 0; k <= CR_Output_Table.Rows.Count - 1; k++)
            {
                dataGridView1.Rows[k + 1].Cells[0].Value = CR_Output_Table.Rows[k][0];
                dataGridView1.Rows[k + 1].Cells[1].Value = CR_Output_Table.Rows[k][1];
                dataGridView1.Rows[k + 1].Cells[2].Value = CR_Output_Table.Rows[k][2];
                dataGridView1.Rows[k + 1].Cells[3].Value = CR_Output_Table.Rows[k][3];
                dataGridView1.Rows[k + 1].Cells[4].Value = CR_Output_Table.Rows[k][4];
                dataGridView1.Rows[k + 1].Cells[5].Value = CR_Output_Table.Rows[k][5];
                dataGridView1.Rows[k + 1].Cells[6].Value = CR_Output_Table.Rows[k][6];
                dataGridView1.Rows[k + 1].Cells[7].Value = CR_Output_Table.Rows[k][7];
                dataGridView1.Rows[k + 1].Cells[8].Value = CR_Output_Table.Rows[k][8];
                dataGridView1.Rows[k + 1].Cells[9].Value = CR_Output_Table.Rows[k][9];
                dataGridView1.Rows[k + 1].Cells[10].Value = CR_Output_Table.Rows[k][10];
                dataGridView1.Rows[k + 1].Cells[11].Value = CR_Output_Table.Rows[k][11];
                dataGridView1.Rows[k + 1].Cells[12].Value = CR_Output_Table.Rows[k][12];
                dataGridView1.Rows[k + 1].Cells[13].Value = CR_Output_Table.Rows[k][13];
                dataGridView1.Rows[k + 1].Cells[14].Value = CR_Output_Table.Rows[k][14];
            }


            data_rows_count = dataGridView1.Rows.Count;


        }

        private void button2_Click(object sender, EventArgs e)
        {

            if (dataGridView1.Rows.Count >1)
            {

                CR_Output_Table1 = CR_Output_Table;
                //CR_Output_Table1.Columns.Add("Site", typeof(string));
                //CR_Output_Table1.Columns.Add("Cell", typeof(string));
                //CR_Output_Table1.Columns.Add("Node", typeof(string));
                //CR_Output_Table1.Columns.Add("Technology", typeof(string));
                //CR_Output_Table1.Columns.Add("KPI", typeof(string));
                //CR_Output_Table1.Columns.Add("Vlaue After CR", typeof(string));
                //CR_Output_Table1.Columns.Add("Vlaue Before CR", typeof(string));
                //CR_Output_Table1.Columns.Add("Performance Comment", typeof(string));
                //CR_Output_Table1.Columns.Add("Performance Engineer", typeof(string));
                //CR_Output_Table1.Columns.Add("CR Num", typeof(string));
                //CR_Output_Table1.Columns.Add("Start Time", typeof(DateTime));
                //CR_Output_Table1.Columns.Add("End Time", typeof(DateTime));
                //CR_Output_Table1.Columns.Add("CR Status ", typeof(string));
                //CR_Output_Table1.Columns.Add("CR Description", typeof(string));


                for (int k = 0; k < data_rows_count - 1; k++)
                {
                    CR_Output_Table1.Rows[k][0] = dataGridView1.Rows[k + 1].Cells[0].Value;
                    CR_Output_Table1.Rows[k][1] = dataGridView1.Rows[k + 1].Cells[1].Value;
                    CR_Output_Table1.Rows[k][2] = dataGridView1.Rows[k + 1].Cells[2].Value;
                    CR_Output_Table1.Rows[k][3] = dataGridView1.Rows[k + 1].Cells[3].Value;
                    CR_Output_Table1.Rows[k][4] = dataGridView1.Rows[k + 1].Cells[4].Value;
                    CR_Output_Table1.Rows[k][5] = dataGridView1.Rows[k + 1].Cells[5].Value;
                    CR_Output_Table1.Rows[k][6] = dataGridView1.Rows[k + 1].Cells[6].Value;

                    CR_Output_Table1.Rows[k][7] = dataGridView1.Rows[k + 1].Cells[7].Value;

                    CR_Output_Table1.Rows[k][8] = dataGridView1.Rows[k + 1].Cells[8].Value;
                    CR_Output_Table1.Rows[k][9] = dataGridView1.Rows[k + 1].Cells[9].Value;
                    CR_Output_Table1.Rows[k][10] = dataGridView1.Rows[k + 1].Cells[10].Value;
                    CR_Output_Table1.Rows[k][11] = dataGridView1.Rows[k + 1].Cells[11].Value;
                    CR_Output_Table1.Rows[k][12] = dataGridView1.Rows[k + 1].Cells[12].Value;
                    CR_Output_Table1.Rows[k][13] = dataGridView1.Rows[k + 1].Cells[13].Value;
                    CR_Output_Table1.Rows[k][14] = dataGridView1.Rows[k + 1].Cells[14].Value;
                }

                XLWorkbook wb = new XLWorkbook();
                wb.Worksheets.Add(CR_Output_Table, "CR Data Table");
                var saveFileDialog = new SaveFileDialog
                {
                    FileName = "CR Check File",
                    Filter = "Excel files|*.xlsx",
                    Title = "Save an Excel File"
                };

                saveFileDialog.ShowDialog();

                if (!String.IsNullOrWhiteSpace(saveFileDialog.FileName))
                    wb.SaveAs(saveFileDialog.FileName);

                MessageBox.Show("Finished");
            }
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }
    }
}
