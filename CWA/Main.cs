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

  
    }




}
