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
using System.Reflection;
using System.Threading;
using System.Text;
using System.Text.RegularExpressions;


namespace CWA
{
    public partial class AutomaticReports : Form
    {
        public AutomaticReports()
        {
            InitializeComponent();
        }


        public Main form1;


        public AutomaticReports(Form form)
        {
            InitializeComponent();
            form1 = (Main)form;
        }


        private void AutomaticReports_Load(object sender, EventArgs e)
        {

        }




        public Excel.Application xlApp1 { get; set; }
        public Excel.Application xlApp2 { get; set; }
        public Excel.Workbook workbook1 { get; set; }
        public Excel.Workbook workbook2 { get; set; }

        public Excel.Worksheet sheet1 { get; set; }
        public Excel.Worksheet sheet2 { get; set; }
        public Excel.Worksheet sheet3 { get; set; }
        public Excel.Worksheet sheet4 { get; set; }

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










        public Excel.Application xlApp { get; set; }
        public Excel.Workbook xlWorkBook { get; set; }
        public Excel.Worksheet Sheet { get; set; }

        public string ConnectionString = "";
        public SqlConnection connection = new SqlConnection();
        public string Server_Name = "PERFORMANCEDB";
        public string DataBase_Name = "Performance_NAK";
        public string update_date_str = "";
        public DateTime update_date = DateTime.Today;


        // Method of Query Execution with Output
        public DataTable Query_Execution_Table_Output(String Query)
        {
            string Quary_String = Query;
            SqlCommand Quary_Command = new SqlCommand(Quary_String, connection);
            Quary_Command.CommandTimeout = 0;
            Quary_Command.ExecuteNonQuery();
            DataTable Output_Table = new DataTable();
            SqlDataAdapter dataAdapter_Quary_Command = new SqlDataAdapter(Quary_Command);
            dataAdapter_Quary_Command.Fill(Output_Table);
            return Output_Table;
        }




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
            Sheet = xlWorkBook.Worksheets[1];



            // Fill in Datagrid list of excel sheets
            String[] excelSheets = new String[xlWorkBook.Worksheets.Count];
            int i = 0;
            dataGridView1.Rows.Clear();
            dataGridView1.RowCount = xlWorkBook.Worksheets.Count+1;
            dataGridView1.Columns[0].Width = 200;
            dataGridView1.Columns[1].Width = 50;
            foreach (Microsoft.Office.Interop.Excel.Worksheet wSheet in xlWorkBook.Worksheets)
            {
                excelSheets[i] = wSheet.Name;
                dataGridView1.Rows[i + 1].Cells[0].Value = excelSheets[i];
                i++;
            }





            if (comboBox1.SelectedItem.ToString()=="Tehran-CTO")
            {


            }




            //Excel.Range Data = Sheet.get_Range("A2", "D" + Sheet.UsedRange.Rows.Count);
            //object[,] Core_Data = (object[,])Data.Value;
            //int Count = Sheet.UsedRange.Rows.Count;

            int y = 0;



        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            update_date = dateTimePicker1.Value.Date;
            update_date_str = Convert.ToString(update_date.Year) + "-" + Convert.ToString(update_date.Month) + "-" + Convert.ToString(update_date.Day);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DataTable E_Table = new DataTable();
            DataTable H_Table = new DataTable();

            E_Table.Columns.Add("Day", typeof(DateTime));
            E_Table.Columns.Add("Date", typeof(DateTime));
            E_Table.Columns.Add("ElementID", typeof(string));
            E_Table.Columns.Add("ElementID1", typeof(string));
            E_Table.Columns.Add("Sector", typeof(string));
            E_Table.Columns.Add("Duplexing", typeof(string));
            E_Table.Columns.Add("C_pmPdcpVolDlDrb_Value", typeof(double));
            E_Table.Columns.Add("C_pmSchedActivityCellDl_Value", typeof(double));
            E_Table.Columns.Add("C_pmPdcpVolDlDrbLastTTI_Value", typeof(double));
            E_Table.Columns.Add("C_pmUeThpTimeDl_Value", typeof(double));
            E_Table.Columns.Add("C_pmPrbUsedDlDtch_Value", typeof(double));
            E_Table.Columns.Add("C_pmPrbAvailDl_Value", typeof(double));
            E_Table.Columns.Add("C_pmPrbUtilDl_000_Value", typeof(double));
            E_Table.Columns.Add("C_pmPrbUtilDl_008_Value", typeof(double));
            E_Table.Columns.Add("C_pmPrbUtilDl_009_Value", typeof(double));
            E_Table.Columns.Add("C_pmCellDowntimeAuto_Value", typeof(double));
            E_Table.Columns.Add("C_pmCellDowntimeMan_Value", typeof(double));
            E_Table.Columns.Add("C_pmRrcConnLevSum_Value", typeof(double));
            E_Table.Columns.Add("C_pmRrcConnLevSamp_Value", typeof(double));

            H_Table.Columns.Add("Day", typeof(DateTime));
            H_Table.Columns.Add("Date", typeof(DateTime));
            H_Table.Columns.Add("ElementID", typeof(string));
            H_Table.Columns.Add("ElementID1", typeof(string));
            H_Table.Columns.Add("Sector", typeof(string));
            H_Table.Columns.Add("Duplexing", typeof(string));
            H_Table.Columns.Add("LThrpbitsDL", typeof(double));
            H_Table.Columns.Add("LThrpTimeCellDLHighPrecision", typeof(double));
            H_Table.Columns.Add("LThrpbitsDLLastTTI", typeof(double));
            H_Table.Columns.Add("LThrpTimeDLRmvLastTTI", typeof(double));
            H_Table.Columns.Add("LTrafficUserAvg", typeof(double));
            H_Table.Columns.Add("LChMeasPRBDLUsedAvg", typeof(double));
            H_Table.Columns.Add("LChMeasPRBDLAvail", typeof(double));
            H_Table.Columns.Add("Counter_1526732727", typeof(double));
            H_Table.Columns.Add("Counter_1526732735", typeof(double));
            H_Table.Columns.Add("Counter_1526732736", typeof(double));
            H_Table.Columns.Add("LCellUnavailDurManual", typeof(double));
            H_Table.Columns.Add("LCellUnavailDurSys", typeof(double));







            openFileDialog1.DefaultExt = "xlsx";
            openFileDialog1.Filter = "Excel File|*.xlsx";
            DialogResult result = openFileDialog1.ShowDialog();
            string File_Name = openFileDialog1.SafeFileName.ToString();
            string Date_of_File = File_Name.Substring(12, 10);
            if (result == DialogResult.OK)
            {

                string file = openFileDialog1.FileName;
                xlApp1 = new Excel.Application();
                workbook1 = xlApp1.Workbooks.Open(file);
                sheet1 = workbook1.Worksheets[1];
                sheet2 = workbook1.Worksheets[2];
                sheet3 = workbook1.Worksheets[3];
                sheet4 = workbook1.Worksheets[4];


                Excel.Range DataRange1 = sheet1.get_Range("A2", "P" + sheet1.UsedRange.Rows.Count);
                object[,] sheet1_Data = (object[,])DataRange1.Value;
                int Count1 = sheet1.UsedRange.Rows.Count;


                Excel.Range DataRange2 = sheet2.get_Range("A2", "P" + sheet2.UsedRange.Rows.Count);
                object[,] sheet2_Data = (object[,])DataRange2.Value;
                int Count2 = sheet2.UsedRange.Rows.Count;

                Excel.Range DataRange3 = sheet3.get_Range("A2", "O" + sheet3.UsedRange.Rows.Count);
                object[,] sheet3_Data = (object[,])DataRange3.Value;
                int Count3 = sheet3.UsedRange.Rows.Count;


                Excel.Range DataRange4 = sheet4.get_Range("A2", "O" + sheet4.UsedRange.Rows.Count);
                object[,] sheet4_Data = (object[,])DataRange4.Value;
                int Count4 = sheet4.UsedRange.Rows.Count;


                for (int k = 0; k < Count1 - 1; k++)
                {

                    DateTime Date = Convert.ToDateTime(sheet1_Data[k + 1, 1]);
                    DateTime Day = Date.Date;
                    string ElementID = sheet1_Data[k + 1, 2].ToString();
                    string ElementID1 = sheet1_Data[k + 1, 3].ToString();
                    if (ElementID1.Length != 10)
                    {
                        continue;
                    }
                    string Sector = ElementID1.Substring(0, 2) + ElementID1.Substring(4, 5);
                    string Band = ElementID1.Substring(2, 2);
                    string Duplexing = "";
                    if (Band == "1L" || Band == "4L" || Band == "7L" || Band == "2L" || Band == "8L" || Band == "3L" || Band == "6L" || Band == "9L")
                    {
                        Duplexing = "FDD";
                    }
                    if (Band == "5L")
                    {
                        Duplexing = "TDD";
                    }
                    double C_pmPdcpVolDlDrb_Value = Convert.ToDouble(sheet1_Data[k + 1, 4]);
                    double C_pmSchedActivityCellDl_Value = Convert.ToDouble(sheet1_Data[k + 1, 5]);
                    double C_pmPdcpVolDlDrbLastTTI_Value = Convert.ToDouble(sheet1_Data[k + 1, 6]);
                    double C_pmUeThpTimeDl_Value = Convert.ToDouble(sheet1_Data[k + 1, 7]);
                    double C_pmPrbUsedDlDtch_Value = Convert.ToDouble(sheet1_Data[k + 1, 8]);
                    double C_pmPrbAvailDl_Value = Convert.ToDouble(sheet1_Data[k + 1, 9]);
                    double C_pmPrbUtilDl_000_Value = Convert.ToDouble(sheet1_Data[k + 1, 10]);
                    double C_pmPrbUtilDl_008_Value = Convert.ToDouble(sheet1_Data[k + 1, 11]);
                    double C_pmPrbUtilDl_009_Value = Convert.ToDouble(sheet1_Data[k + 1, 12]);
                    double C_pmCellDowntimeAuto_Value = Convert.ToDouble(sheet1_Data[k + 1, 13]);
                    double C_pmCellDowntimeMan_Value = Convert.ToDouble(sheet1_Data[k + 1, 14]);
                    double C_pmRrcConnLevSum_Value = Convert.ToDouble(sheet1_Data[k + 1, 15]);
                    double C_pmRrcConnLevSamp_Value = Convert.ToDouble(sheet1_Data[k + 1, 16]);

                    if (ElementID == "TH")
                    {
                        E_Table.Rows.Add(Day, Date, ElementID, ElementID1, Sector, Duplexing, C_pmPdcpVolDlDrb_Value, C_pmSchedActivityCellDl_Value, C_pmPdcpVolDlDrbLastTTI_Value, C_pmUeThpTimeDl_Value, C_pmPrbUsedDlDtch_Value, C_pmPrbAvailDl_Value, C_pmPrbUtilDl_000_Value, C_pmPrbUtilDl_008_Value, C_pmPrbUtilDl_009_Value, C_pmCellDowntimeAuto_Value, C_pmCellDowntimeMan_Value, C_pmRrcConnLevSum_Value, C_pmRrcConnLevSamp_Value);
                    }
                }

                for (int k = 0; k < Count2 - 1; k++)
                {
                    DateTime Date = Convert.ToDateTime(sheet2_Data[k + 1, 1]);
                    DateTime Day = Date.Date;
                    string ElementID = sheet2_Data[k + 1, 2].ToString();
                    string ElementID1 = sheet2_Data[k + 1, 3].ToString();
                    if (ElementID1.Length != 10)
                    {
                        continue;
                    }
                    string Sector = ElementID1.Substring(0, 2) + ElementID1.Substring(4, 5);
                    string Band = ElementID1.Substring(2, 2);
                    string Duplexing = "";
                    if (Band == "1L" || Band == "4L" || Band == "7L" || Band == "2L" || Band == "8L" || Band == "3L" || Band == "6L" || Band == "9L")
                    {
                        Duplexing = "FDD";
                    }
                    if (Band == "5L")
                    {
                        Duplexing = "TDD";
                    }
                    double C_pmPdcpVolDlDrb_Value = Convert.ToDouble(sheet2_Data[k + 1, 4]);
                    double C_pmSchedActivityCellDl_Value = Convert.ToDouble(sheet2_Data[k + 1, 5]);
                    double C_pmPdcpVolDlDrbLastTTI_Value = Convert.ToDouble(sheet2_Data[k + 1, 6]);
                    double C_pmUeThpTimeDl_Value = Convert.ToDouble(sheet2_Data[k + 1, 7]);
                    double C_pmPrbUsedDlDtch_Value = Convert.ToDouble(sheet2_Data[k + 1, 8]);
                    double C_pmPrbAvailDl_Value = Convert.ToDouble(sheet2_Data[k + 1, 9]);
                    double C_pmPrbUtilDl_000_Value = Convert.ToDouble(sheet2_Data[k + 1, 10]);
                    double C_pmPrbUtilDl_008_Value = Convert.ToDouble(sheet2_Data[k + 1, 11]);
                    double C_pmPrbUtilDl_009_Value = Convert.ToDouble(sheet2_Data[k + 1, 12]);
                    double C_pmCellDowntimeAuto_Value = Convert.ToDouble(sheet2_Data[k + 1, 13]);
                    double C_pmCellDowntimeMan_Value = Convert.ToDouble(sheet2_Data[k + 1, 14]);
                    double C_pmRrcConnLevSum_Value = Convert.ToDouble(sheet2_Data[k + 1, 15]);
                    double C_pmRrcConnLevSamp_Value = Convert.ToDouble(sheet2_Data[k + 1, 16]);


                    if (ElementID == "TH")
                    {
                        E_Table.Rows.Add(Day, Date, ElementID, ElementID1, Sector, Duplexing, C_pmPdcpVolDlDrb_Value, C_pmSchedActivityCellDl_Value, C_pmPdcpVolDlDrbLastTTI_Value, C_pmUeThpTimeDl_Value, C_pmPrbUsedDlDtch_Value, C_pmPrbAvailDl_Value, C_pmPrbUtilDl_000_Value, C_pmPrbUtilDl_008_Value, C_pmPrbUtilDl_009_Value, C_pmCellDowntimeAuto_Value, C_pmCellDowntimeMan_Value, C_pmRrcConnLevSum_Value, C_pmRrcConnLevSamp_Value);
                    }
                }




                for (int k = 0; k < Count3 - 1; k++)
                {

                    DateTime Date = Convert.ToDateTime(sheet3_Data[k + 1, 1]);
                    DateTime Day = Date.Date;
                    string ElementID = sheet3_Data[k + 1, 2].ToString();
                    string ElementID1 = sheet3_Data[k + 1, 3].ToString();
                    if (ElementID1.Length != 10)
                    {
                        continue;
                    }
                    string Sector = ElementID1.Substring(0, 2) + ElementID1.Substring(4, 5);
                    string Band = ElementID1.Substring(2, 2);
                    string Duplexing = "";
                    if (Band == "1L" || Band == "4L" || Band == "7L" || Band == "2L" || Band == "8L" || Band == "3L" || Band == "6L" || Band == "9L")
                    {
                        Duplexing = "FDD";
                    }
                    if (Band == "5L")
                    {
                        Duplexing = "TDD";
                    }
                    double LThrpbitsDL = Convert.ToDouble(sheet3_Data[k + 1, 4]);
                    double LThrpTimeCellDLHighPrecision = Convert.ToDouble(sheet3_Data[k + 1, 5]);
                    double LThrpbitsDLLastTTI = Convert.ToDouble(sheet3_Data[k + 1, 6]);
                    double LThrpTimeDLRmvLastTTI = Convert.ToDouble(sheet3_Data[k + 1, 7]);
                    double LTrafficUserAvg = Convert.ToDouble(sheet3_Data[k + 1, 8]);
                    double LChMeasPRBDLUsedAvg = Convert.ToDouble(sheet3_Data[k + 1, 9]);
                    double LChMeasPRBDLAvail = Convert.ToDouble(sheet3_Data[k + 1, 10]);
                    double Counter_1526732727 = Convert.ToDouble(sheet3_Data[k + 1, 11]);
                    double Counter_1526732735 = Convert.ToDouble(sheet3_Data[k + 1, 12]);
                    double Counter_1526732736 = Convert.ToDouble(sheet3_Data[k + 1, 13]);
                    double LCellUnavailDurManual = Convert.ToDouble(sheet3_Data[k + 1, 14]);
                    double LCellUnavailDurSys = Convert.ToDouble(sheet3_Data[k + 1, 15]);




                    if (ElementID == "TH")
                    {
                        H_Table.Rows.Add(Day, Date, ElementID, ElementID1, Sector, Duplexing, LThrpbitsDL, LThrpTimeCellDLHighPrecision, LThrpbitsDLLastTTI, LThrpTimeDLRmvLastTTI, LTrafficUserAvg, LChMeasPRBDLUsedAvg, LChMeasPRBDLAvail, Counter_1526732727, Counter_1526732735, Counter_1526732736, LCellUnavailDurManual, LCellUnavailDurSys);
                    }
                }




                for (int k = 0; k < Count4 - 1; k++)
                {

                    DateTime Date = Convert.ToDateTime(sheet4_Data[k + 1, 1]);
                    DateTime Day = Date.Date;
                    string ElementID = sheet4_Data[k + 1, 2].ToString();
                    string ElementID1 = sheet4_Data[k + 1, 3].ToString();
                    if (ElementID1.Length != 10)
                    {
                        continue;
                    }
                    string Sector = ElementID1.Substring(0, 2) + ElementID1.Substring(4, 5);
                    string Band = ElementID1.Substring(2, 2);
                    string Duplexing = "";
                    if (Band == "1L" || Band == "4L" || Band == "7L" || Band == "2L" || Band == "8L" || Band == "3L" || Band == "6L" || Band == "9L")
                    {
                        Duplexing = "FDD";
                    }
                    if (Band == "5L")
                    {
                        Duplexing = "TDD";
                    }
                    double LThrpbitsDL = Convert.ToDouble(sheet3_Data[k + 1, 4]);
                    double LThrpTimeCellDLHighPrecision = Convert.ToDouble(sheet4_Data[k + 1, 5]);
                    double LThrpbitsDLLastTTI = Convert.ToDouble(sheet3_Data[k + 1, 6]);
                    double LThrpTimeDLRmvLastTTI = Convert.ToDouble(sheet4_Data[k + 1, 7]);
                    double LTrafficUserAvg = Convert.ToDouble(sheet4_Data[k + 1, 8]);
                    double LChMeasPRBDLUsedAvg = Convert.ToDouble(sheet4_Data[k + 1, 9]);
                    double LChMeasPRBDLAvail = Convert.ToDouble(sheet4_Data[k + 1, 10]);
                    double Counter_1526732727 = Convert.ToDouble(sheet4_Data[k + 1, 11]);
                    double Counter_1526732735 = Convert.ToDouble(sheet4_Data[k + 1, 12]);
                    double Counter_1526732736 = Convert.ToDouble(sheet4_Data[k + 1, 13]);
                    double LCellUnavailDurManual = Convert.ToDouble(sheet4_Data[k + 1, 14]);
                    double LCellUnavailDurSys = Convert.ToDouble(sheet4_Data[k + 1, 15]);


                    if (ElementID == "TH")
                    {
                        H_Table.Rows.Add(Day, Date, ElementID, ElementID1, Sector, Duplexing, LThrpbitsDL, LThrpTimeCellDLHighPrecision, LThrpbitsDLLastTTI, LThrpTimeDLRmvLastTTI, LTrafficUserAvg, LChMeasPRBDLUsedAvg, LChMeasPRBDLAvail, Counter_1526732727, Counter_1526732735, Counter_1526732736, LCellUnavailDurManual, LCellUnavailDurSys);
                    }
                }


                var E_Results = from row in E_Table.AsEnumerable()
                                group row by new { f1 = row.Field<DateTime>("Date"), f2 = row.Field<string>("Sector"), f3 = row.Field<string>("Duplexing") } into rows
                                select new
                                {
                                    Date = rows.Key.f1,
                                    Sector = rows.Key.f2,
                                    Duplexing = rows.Key.f3,
                                    C_pmPdcpVolDlDrb_Value = rows.Sum(r => r.Field<double>("C_pmPdcpVolDlDrb_Value")),
                                    C_pmSchedActivityCellDl_Value = rows.Sum(r => r.Field<double>("C_pmSchedActivityCellDl_Value")),
                                    C_pmPdcpVolDlDrbLastTTI_Value = rows.Sum(r => r.Field<double>("C_pmPdcpVolDlDrbLastTTI_Value")),
                                    C_pmUeThpTimeDl_Value = rows.Sum(r => r.Field<double>("C_pmUeThpTimeDl_Value")),
                                    C_pmPrbUsedDlDtch_Value = rows.Sum(r => r.Field<double>("C_pmPrbUsedDlDtch_Value")),
                                    C_pmPrbAvailDl_Value = rows.Sum(r => r.Field<double>("C_pmPrbAvailDl_Value")),
                                    C_pmPrbUtilDl_000_Value = rows.Sum(r => r.Field<double>("C_pmPrbUtilDl_000_Value")),
                                    C_pmPrbUtilDl_008_Value = rows.Sum(r => r.Field<double>("C_pmPrbUtilDl_008_Value")),
                                    C_pmPrbUtilDl_009_Value = rows.Sum(r => r.Field<double>("C_pmPrbUtilDl_009_Value")),
                                    C_pmCellDowntimeAuto_Value = rows.Sum(r => r.Field<double>("C_pmCellDowntimeAuto_Value")),
                                    C_pmCellDowntimeMan_Value = rows.Sum(r => r.Field<double>("C_pmCellDowntimeMan_Value")),
                                    C_pmRrcConnLevSum_Value = rows.Sum(r => r.Field<double>("C_pmRrcConnLevSum_Value")),
                                    C_pmRrcConnLevSamp_Value = rows.Sum(r => r.Field<double>("C_pmRrcConnLevSamp_Value"))
                                };


                DataTable E_Sector_Table = new DataTable();
                E_Sector_Table = ConvertToDataTable(E_Results);




                XLWorkbook wb = new XLWorkbook();
                //var wb = new XLWorkbook();
                wb.Worksheets.Add(E_Sector_Table, "TH_E_LTE_Sec_Counter_Hourly");
                //wb.Worksheets.Add(Site_Data_Table_2G, "Result");
                var saveFileDialog = new SaveFileDialog
                {
                    FileName = "TH_LTE_Sector_Counter_Hourly_"+ Date_of_File,
                    Filter = "Excel files|*.xlsx",
                    Title = "Save an Excel File"
                };

                saveFileDialog.ShowDialog();

                if (!String.IsNullOrWhiteSpace(saveFileDialog.FileName))
                    wb.SaveAs(saveFileDialog.FileName);

                MessageBox.Show("Finished");


            }
        }
    }
}
