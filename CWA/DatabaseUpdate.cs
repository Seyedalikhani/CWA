using ClosedXML;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using Excel = Microsoft.Office.Interop.Excel;


namespace CWA
{
    public partial class DatabaseUpdate : Form
    {
        public DatabaseUpdate()
        {
            InitializeComponent();
        }


        public Main form1;


        public DatabaseUpdate(Form form)
        {
            InitializeComponent();
            form1 = (Main)form;
        }


   
        //public Excel.Application xlApp1 { get; set; }
        //public Excel.Application xlApp2 { get; set; }
        //public Excel.Workbook workbook1 { get; set; }
        //public Excel.Workbook workbook2 { get; set; }

        //public Excel.Worksheet sheet1 { get; set; }
        //public Excel.Worksheet sheet2 { get; set; }
        //public Excel.Worksheet sheet3 { get; set; }
        //public Excel.Worksheet sheet4 { get; set; }

        //public Excel.Workbook xlWorkBook { get; set; }

        public DataTable E_Table = new DataTable();
        public DataTable H_Table = new DataTable();


        public DataTable ConvertToDataTable<T>(IEnumerable<T> varlist)
        {
            DataTable dtReturn = new DataTable();

            // column names   
            PropertyInfo[] oProps = null;

            if (varlist == null) return dtReturn;

            foreach (T rec in varlist)
            {
                // Use reflection to get property names, to create table, Only first time, others will follow   
                if (oProps == null)
                {
                    oProps = ((Type)rec.GetType()).GetProperties();
                    foreach (PropertyInfo pi in oProps)
                    {
                        Type colType = pi.PropertyType;

                        if ((colType.IsGenericType) && (colType.GetGenericTypeDefinition() == typeof(Nullable<>)))
                        {
                            colType = colType.GetGenericArguments()[0];
                        }

                        dtReturn.Columns.Add(new DataColumn(pi.Name, colType));
                    }
                }

                DataRow dr = dtReturn.NewRow();

                foreach (PropertyInfo pi in oProps)
                {
                    dr[pi.Name] = pi.GetValue(rec, null) == null ? DBNull.Value : pi.GetValue
                    (rec, null);
                }

                dtReturn.Rows.Add(dr);
            }
            return dtReturn;
        }



        public DataTable ConvertWorksheetToDataTable(Excel.Worksheet sheet)
        {
            DataTable dt = new DataTable();
            Excel.Range usedRange = sheet.UsedRange;

            int rowCount = usedRange.Rows.Count;
            int colCount = usedRange.Columns.Count;

            // Read column headers
            for (int c = 1; c <= colCount; c++)
            {
                var colName = (usedRange.Cells[1, c] as Excel.Range)?.Value2?.ToString() ?? $"Column{c}";
                dt.Columns.Add(colName);
            }

            // Read rows (starting from row 2 assuming row 1 is header)
            for (int r = 2; r <= rowCount; r++)
            {
                DataRow row = dt.NewRow();
                for (int c = 1; c <= colCount; c++)
                {
                    row[c - 1] = (usedRange.Cells[r, c] as Excel.Range)?.Value2?.ToString();
                }
                dt.Rows.Add(row);
            }

            return dt;
        }


        public Excel.Application xlApp { get; set; }
        public Excel.Workbook xlWorkBook { get; set; }
        public Excel.Worksheet Sheet1 { get; set; }
        public Excel.Worksheet Sheet2 { get; set; }
        public Excel.Worksheet Sheet3 { get; set; }

        public string ConnectionString = "";
        public SqlConnection connection = new SqlConnection();
        //public string Server_Name = "PERFORMANCEDB";
        //public string DataBase_Name = "Performance_NAK";
        public string Server_Name = @"AHMAD\" + "SQLEXPRESS";
        public string DataBase_Name = "NAK";
        public string update_date_str = "";
        public DateTime update_date = DateTime.Today;






        // OPen Excel File
        private void button1_Click(object sender, EventArgs e)
        {

            openFileDialog1.DefaultExt = "xlsx";
            openFileDialog1.Filter = "Excel File|*.xlsx";
            DialogResult result = openFileDialog1.ShowDialog();
            string File_Name = openFileDialog1.SafeFileName.ToString();
            string file = openFileDialog1.FileName;
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(file);
            Sheet1 = xlWorkBook.Worksheets[1];
            Sheet2 = xlWorkBook.Worksheets[2];
            Sheet3 = xlWorkBook.Worksheets[3];

            // Convert to DataTables
            DataTable dt1 = ConvertWorksheetToDataTable(Sheet1);
            DataTable dt2 = ConvertWorksheetToDataTable(Sheet2);
            DataTable dt3 = ConvertWorksheetToDataTable(Sheet3);

            int rr = 0;

         


        }

       



        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void DatabaseUpdate_Load(object sender, EventArgs e)
        {
            ConnectionString = @"Server=" + Server_Name + "; Database=" + DataBase_Name + "; Trusted_Connection=True;";
            //ConnectionString = @"Server=" + Server_Name + "; Database=" + DataBase_Name + "; User ID=cwpcApp;Password=cwpcApp@830625#Ahmad";
            connection = new SqlConnection(ConnectionString);
            connection.Open();
        }
    }
}
