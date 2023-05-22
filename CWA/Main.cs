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
using System.Threading;

namespace CWA
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }




        private void Form1_Load(object sender, EventArgs e)
        {

            ConnectionString = @"Server=" + Server_Name + "; Database=" + DataBase_Name + "; User ID=cwpcApp; Password=cwpcApp@830625#Ahmad";
            connection = new SqlConnection(ConnectionString);
            connection.Open();


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


       
        public string ConnectionString = "";
        public SqlConnection connection = new SqlConnection();

    
        //public string Server_Name = @"NAKPRG-NB1243\" + "AHMAD";
        //public string DataBase_Name = "Contract";


        public string Server_Name = "PERFORMANCEDB01";
        public string DataBase_Name = "Performance_NAK";

        //public string Server_Name = "core";
        //public string DataBase_Name = "Core_Performance_Mohammad";


        //public string Server_Name = "172.26.7.159";
        //public string DataBase_Name = "Performance_NAK";

 


        // ****** worstCellReports ******
        private void worstCellReportsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            WorstCellsReport newFrm = new WorstCellsReport(this);                                  // Form1 for Setting
            newFrm.StartPosition = FormStartPosition.CenterScreen;
            newFrm.Text = "Worst Cell Reports";
            newFrm.Size = new Size(841, 454);
            newFrm.TopMost = true;
            newFrm.Show();
        }   


        // ****** mAP ******
        private void mAPToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MAP newFrm = new MAP(this);                                  // Form1 for Setting
            newFrm.StartPosition = FormStartPosition.CenterScreen;
            //newFrm.Location = new FormStartPosition.CenterScreen;
            newFrm.Text = "MAP";
            newFrm.Size = new Size(1360, 760);
            newFrm.TopMost = true;
            newFrm.Show();
        }

        // ****** kPIZero ******
        private void kPIZeroToolStripMenuItem_Click(object sender, EventArgs e)
        {
            KPIZero newFrm = new KPIZero(this);                                  // Form1 for Setting
            newFrm.StartPosition = FormStartPosition.CenterScreen;
            //newFrm.Location = new FormStartPosition.CenterScreen;
            newFrm.Size = new Size(1000, 660);
            newFrm.TopMost = true;
            newFrm.Show();
        }

        // ****** cR ******
        private void cRToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CR newFrm = new CR(this);                                  // Form1 for Setting
            newFrm.StartPosition = FormStartPosition.CenterScreen;
            //newFrm.Location = new FormStartPosition.CenterScreen;
            newFrm.Size = new Size(1115, 600);
            newFrm.TopMost = true;
            newFrm.Show();
        }

        // ****** LTE ******
        private void exportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LTE newFrm = new LTE(this);                                  // Form1 for Setting
            newFrm.StartPosition = FormStartPosition.CenterScreen;
            //newFrm.Location = new FormStartPosition.CenterScreen;
            newFrm.Size = new Size(1378, 780);
            // newFrm.AutoScroll = true;
            // newFrm.AutoSize = true;
            newFrm.TopMost = true;
            newFrm.Show();
        }

        // ****** availability ******
        private void availabilityToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Availability newFrm = new Availability(this);                                  // Form1 for Setting
            newFrm.StartPosition = FormStartPosition.CenterScreen;
            //newFrm.Location = new FormStartPosition.CenterScreen;
            newFrm.Size = new Size(512, 339);
            // newFrm.AutoScroll = true;
            // newFrm.AutoSize = true;
            newFrm.TopMost = true;
            newFrm.Show();
        }


        // ****** dashboards ******
        private void dashboardsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Dashboards newFrm = new Dashboards(this);                                  // Form1 for Setting
            //newFrm.StartPosition = FormStartPosition.CenterScreen;
            newFrm.Text = "Dashboards";
            newFrm.AutoScroll = true;
            newFrm.AutoSize = true;
            // newFrm.Size = new Size(4000, 3000);
            newFrm.TopMost = true;
            newFrm.Show();
        }


        // ******* core ********
        private void coreToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Core newFrm = new Core(this);                                  // Form1 for Setting
            newFrm.StartPosition = FormStartPosition.CenterScreen;
            //newFrm.Location = new FormStartPosition.CenterScreen;
            newFrm.Size = new Size(664, 512);
            // newFrm.AutoScroll = true;
            // newFrm.AutoSize = true;
            newFrm.TopMost = true;
            newFrm.Show();
        }

        // ******* customerComplain ********
        private void customerComplainToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CustomerComplaint newFrm = new CustomerComplaint(this);                                  // Form1 for Setting
            newFrm.StartPosition = FormStartPosition.CenterScreen;
            //newFrm.Location = new FormStartPosition.CenterScreen;
            newFrm.Size = new Size(375, 183);
            // newFrm.AutoScroll = true;
            // newFrm.AutoSize = true;
            newFrm.TopMost = true;
            newFrm.Show();

        }



        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            listBox2.Items.Clear();
            comboBox3.Items.Clear();
            Technology = comboBox2.SelectedItem.ToString();


            if (Technology == "2G")
            {
                listBox2.Items.Add("2G_TCH_Traffic (Erlang)");
                listBox2.Items.Add("2G_Payload (GB)");
                listBox2.Items.Add("CDR");
                listBox2.Items.Add("CSSR");
                listBox2.Items.Add("IHSR");
                listBox2.Items.Add("OHSR");
                listBox2.Items.Add("RxQual_DL");
                listBox2.Items.Add("RxQual_UL");
                listBox2.Items.Add("SDCCH_Access_SR");
                listBox2.Items.Add("SDCCH_Congestion");
                listBox2.Items.Add("SDCCH_Drop_Rate");
                listBox2.Items.Add("TCH_Assign_FR");
                listBox2.Items.Add("TCH_Congestion");
                listBox2.Items.Add("TCH_Availability");
            }

            if (Technology == "3G")
            {
                comboBox3.Items.Add("All Band");
                comboBox3.Items.Add("All Carrier");
                comboBox3.Items.Add("U900");
                comboBox3.Items.Add("U2100");
                comboBox3.Items.Add("U900F1");
                comboBox3.Items.Add("U900F2");
                comboBox3.Items.Add("U2100F1");
                comboBox3.Items.Add("U2100F2");
                comboBox3.Items.Add("U2100F3");

                listBox2.Items.Add("3G_CS_Traffic (Erlang)");
                listBox2.Items.Add("3G_PS_Payload (GB)");
                listBox2.Items.Add("CS_RAB_Establish");
                listBox2.Items.Add("CS_IRAT_HO_SR");
                listBox2.Items.Add("CS_Drop_Rate");
                listBox2.Items.Add("Soft_HO_SR");
                listBox2.Items.Add("CS_RRC_SR");
                listBox2.Items.Add("CS_MultiRAB_SR");
                listBox2.Items.Add("Inter_Carrier_HO_SR");
                listBox2.Items.Add("Cell_Availability");


                listBox2.Items.Add("HSDPA_SR");
                listBox2.Items.Add("HSUPA_SR");
                listBox2.Items.Add("DL_User_THR (Mbps)");
                listBox2.Items.Add("UL_User_THR (Kbps)");
                listBox2.Items.Add("HSDAP_Drop_Rate");
                listBox2.Items.Add("HSUAP_Drop_Rate");
                listBox2.Items.Add("PS_RRC_SR");
                listBox2.Items.Add("Ps_RAB_Establish");
                listBox2.Items.Add("PS_MultiRAB_Establish");
                listBox2.Items.Add("PS_Drop_Rate");
                listBox2.Items.Add("HSDPA_Cell_Change_SR");
                listBox2.Items.Add("HS_Share_Payload");
                listBox2.Items.Add("DL_Cell_THR (Mbps)");
                listBox2.Items.Add("RSSI (dBm)");
                listBox2.Items.Add("Average CQI");




            }
            if (Technology == "4G")
            {
                comboBox3.Items.Add("All Band");
                comboBox3.Items.Add("All Carrier");
                comboBox3.Items.Add("L1800");
                comboBox3.Items.Add("L2600");
                comboBox3.Items.Add("L2100");
                comboBox3.Items.Add("L2300");
                comboBox3.Items.Add("L900");
                comboBox3.Items.Add("L1800F1");
                comboBox3.Items.Add("L1800F2");
                comboBox3.Items.Add("L2600F1");
                comboBox3.Items.Add("L2600F2");
                comboBox3.Items.Add("L2100F1");
                comboBox3.Items.Add("L2100F2");
                comboBox3.Items.Add("L2300F1");
                comboBox3.Items.Add("L2300F2");
                comboBox3.Items.Add("L900F1");
                comboBox3.Items.Add("L900F2");
                
                listBox2.Items.Add("4G_Total_Paylaod (GB)");
                listBox2.Items.Add("Volte_Traffic (Erlang)");
                listBox2.Items.Add("RRC_Connection_SR");
                listBox2.Items.Add("ERAB_SR_Initial");
                listBox2.Items.Add("ERAB_SR_Added");
                listBox2.Items.Add("DL_THR (Mbps)");
                listBox2.Items.Add("UL_THR (Mbps)");
                listBox2.Items.Add("ERAB_Drop_Rate");
                listBox2.Items.Add("S1_Signalling_SR");
                listBox2.Items.Add("Intra_Freq_SR");
                listBox2.Items.Add("Inter_Freq_SR");
                listBox2.Items.Add("UL_Packet_Loss");
                listBox2.Items.Add("UE_DL_Latency (ms)");
                listBox2.Items.Add("Average CQI");
                listBox2.Items.Add("PUCCH_RSSI (dBm)");
                listBox2.Items.Add("PUSCH_RSSI (dBm)");
                listBox2.Items.Add("Cell_Availability");


            }
            if (Technology == "5G")
            {
                comboBox3.Items.Add("All Band");
                comboBox3.Items.Add("All Carrier");
                comboBox3.Items.Add("N2300");
                comboBox3.Items.Add("N3500");

                listBox2.Items.Add("5G_Traffic (GB)");
                listBox2.Items.Add("5G_DL_Traffic (GB)");
                listBox2.Items.Add("5G_UL_Traffic (GB)");
                listBox2.Items.Add("DL User Throughput (Mbps)");
                listBox2.Items.Add("UL User Throughput (Mbps)");
                listBox2.Items.Add("Max Number of Users");


            }
        }



        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox1.TextLength < textBox1.SelectionStart + textBox1.SelectionLength + 1)
            {
                textBox1.Text = textBox1.Text.Remove(textBox1.SelectionStart, textBox1.SelectionLength);
            }
            else
            {
                textBox1.Text = textBox1.Text.Remove(textBox1.SelectionStart, textBox1.SelectionLength + 1);
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked==true)
            {
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
            }
        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox7.Checked == true)
            {
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked == true)
            {
                checkBox2.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
            }
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked == true)
            {
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
            }
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked == true)
            {
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
            }
        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox6.Checked == true)
            {
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox7.Checked = false;
            }
        }



        // Our approach is to get data based on user selection and for and processing commands use linq instead of SQL connections.

        public bool Province_Capital_Check = false;
        public string Technology = "2G";
        public string Vendor = "All";
        public string Frequency = "All Band";
        public string Interval = "Daily";
        public string Site_list = "";

        private void button3_Click(object sender, EventArgs e)
        {
            Thread th1 = new Thread(my_thread1);
            th1.Start();
        }

        void my_thread1()
        {
            Site_list = textBox1.SelectedText.ToString();
            // Firstly we must detect the category of data 


        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            Vendor = comboBox4.SelectedItem.ToString();
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            Frequency = comboBox3.SelectedItem.ToString();
        }
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            Province_Capital_Check = checkBox1.Checked;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Interval = comboBox1.SelectedItem.ToString();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            Site_list = textBox1.SelectedText.ToString();
        }
    }




}
