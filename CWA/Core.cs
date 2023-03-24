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
    public partial class Core : Form
    {


        public Core()
        {
            InitializeComponent();
        }


        public Main form1;


        public Core(Form form)
        {
            InitializeComponent();
            form1 = (Main)form;
        }



        public Excel.Application xlApp1 { get; set; }
        public Excel.Application xlApp2 { get; set; }
        public Excel.Workbook workbook1 { get; set; }
        public Excel.Workbook workbook2 { get; set; }

        public string start_date_str = "";
        public string end_date_str = "";

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




        public string Server_Name = "core";
        public string DataBase_Name = "Core_Performance_Mohammad";
        public string ConnectionString = "";
        public SqlConnection connection = new SqlConnection();
        public DataTable H_BICC_INCOMING_MSS_Table = new DataTable();
        public DataTable H_BICC_OUTGOING_MSS_Table = new DataTable();
        public DateTime start_date = DateTime.Today;
        public DateTime end_date = DateTime.Today;




        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            start_date = dateTimePicker1.Value.Date;
            start_date_str = Convert.ToString(start_date.Year) + "-" + Convert.ToString(start_date.Month) + "-" + Convert.ToString(start_date.Day);
        }

        private void Form12_Load(object sender, EventArgs e)
        {

        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            //end_date = dateTimePicker2.Value.Date.AddDays(1);
            end_date = dateTimePicker2.Value.Date.AddDays(1);
            end_date_str = Convert.ToString(end_date.Year) + "-" + Convert.ToString(end_date.Month) + "-" + Convert.ToString(end_date.Day);
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        public Excel.Worksheet sheet1 { get; set; }
        public Excel.Worksheet sheet2 { get; set; }
        public Excel.Worksheet sheet3 { get; set; }
        public Excel.Worksheet sheet4 { get; set; }
        public Excel.Worksheet sheet5 { get; set; }
        public Excel.Worksheet sheet6 { get; set; }
        public ProgressBar progfressbar1 { get; set; }

        private void button2_Click(object sender, EventArgs e)
        {

            //ConnectionString = @"Server=" + Server_Name + "; Database=" + DataBase_Name + "; Integrated Security=True";
            ConnectionString = @"Server=" + Server_Name + "; Database=" + DataBase_Name + "; User ID=Ahmad_Core; Password=cwpcApp@830625Ahmad";
            connection = new SqlConnection(ConnectionString);
            connection.Open();



            openFileDialog1.DefaultExt = "xlsx";
            openFileDialog1.Filter = "Excel File|*.xlsx";
            DialogResult result = openFileDialog1.ShowDialog();
            string File_Name = openFileDialog1.SafeFileName.ToString();
            if (result == DialogResult.OK)
            {

                string file = openFileDialog1.FileName;
                xlApp1 = new Excel.Application();
                workbook1 = xlApp1.Workbooks.Open(file);
                sheet6 = workbook1.Worksheets[6];
                sheet5 = workbook1.Worksheets[5];
                sheet4 = workbook1.Worksheets[4];
                sheet3 = workbook1.Worksheets[3];
                sheet2 = workbook1.Worksheets[2];
                sheet1 = workbook1.Worksheets[1];
            }



            Thread th1 = new Thread(my_thread1);
            th1.Start();

        }


        public void my_thread1()
        {



            //************ MSS - VLR  Huawei****************
            string MSS_VLR_Huawei_Use_Quary_String = @"select sum([Huawei MSS VLR Use]) as 'Huawei MSS VLR Use'
                            from(
                              select
                                   [NE Name],
                                   max([Total Number of Subscribers in VLR (entries)]) as 'Huawei MSS VLR Use'
                                   from(select
                                           [Start Time],
                                           [NE Name],
                                           [Total Number of Subscribers in VLR (entries)] from [H_VLR_SUBSCRIBER]
                                           WHERE CAST([Start Time] AS DATE) >='" + start_date_str + "' and CAST([Start Time] AS DATE)<'" + end_date_str + "') tble group by [NE Name]) tble";


            SqlCommand MSS_VLR_Huawei_Use_Quary = new SqlCommand(MSS_VLR_Huawei_Use_Quary_String, connection);
            MSS_VLR_Huawei_Use_Quary.CommandTimeout = 0;
            MSS_VLR_Huawei_Use_Quary.ExecuteNonQuery();
            DataTable MSS_VLR_Huawei_Use_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_MSS_VLR_Huawei_Use = new SqlDataAdapter(MSS_VLR_Huawei_Use_Quary);
            dataAdapter_Contractual_MSS_VLR_Huawei_Use.Fill(MSS_VLR_Huawei_Use_Table);


            sheet3.Cells[12, 3] = Convert.ToInt64(MSS_VLR_Huawei_Use_Table.Rows[0].ItemArray[0]);

            string MSS_VLR_Huawei_Capacity_Quary_String = @"select sum([Maximum of Users])  as 'Huawei MSS VLR Capacity' from[dbo].[H_VLR_Capacity]";
            SqlCommand MSS_VLR_Huawei_Capacity_Quary = new SqlCommand(MSS_VLR_Huawei_Capacity_Quary_String, connection);
            MSS_VLR_Huawei_Capacity_Quary.CommandTimeout = 0;
            MSS_VLR_Huawei_Capacity_Quary.ExecuteNonQuery();
            DataTable MSS_VLR_Huawei_Capacity_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_MSS_VLR_Huawei_Capacity = new SqlDataAdapter(MSS_VLR_Huawei_Capacity_Quary);
            dataAdapter_Contractual_MSS_VLR_Huawei_Capacity.Fill(MSS_VLR_Huawei_Capacity_Table);

            sheet3.Cells[10, 3] = Convert.ToInt64(MSS_VLR_Huawei_Capacity_Table.Rows[0].ItemArray[0]);





            //************MSS - VLR  Nokia****************
            string MSS_VLR_Nokia_Use_Quary_String = @"select sum([Nokia MSS VLR Use]) as 'Nokia MSS VLR Use'
                        from(
                          select
                               [NE],
                               max([Nokia MSS VLR Use Value]) as 'Nokia MSS VLR Use'
                               from(select
                                       [Date],
                                       [NE],
                                       [MSS_Peak_Lic_Utilization(Nokia(Core)]*[MSS_Licence_Capacity(Nokia_Core)] / 100 as 'Nokia MSS VLR Use Value' from [Nokia_MSS]
                                       WHERE CAST([Date] AS DATE)>='" + start_date_str + "' and CAST([Date] AS DATE) <'" + end_date_str + "' and substring([NE],9,3)= '221' ) tble group by[NE]) tble";


            SqlCommand MSS_VLR_Nokia_Use_Quary = new SqlCommand(MSS_VLR_Nokia_Use_Quary_String, connection);
            MSS_VLR_Nokia_Use_Quary.CommandTimeout = 0;
            MSS_VLR_Nokia_Use_Quary.ExecuteNonQuery();
            DataTable MSS_VLR_Nokia_Use_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_MSS_VLR_Nokia_Use = new SqlDataAdapter(MSS_VLR_Nokia_Use_Quary);
            dataAdapter_Contractual_MSS_VLR_Nokia_Use.Fill(MSS_VLR_Nokia_Use_Table);

            sheet3.Cells[12, 4] = Convert.ToInt64(MSS_VLR_Nokia_Use_Table.Rows[0].ItemArray[0]);

            string MSS_VLR_Nokia_Capacity_Quary_String = @"select sum([Nokia MSS VLR Capacity]) as 'Nokia MSS VLR Capacity'
                        from(
                          select
                               [NE],
                               max([Nokia MSS VLR Capacity Value]) as 'Nokia MSS VLR Capacity'
                               from(select
                                       [Date],
                                       [NE],
                                       [MSS_Licence_Capacity(Nokia_Core)] as 'Nokia MSS VLR Capacity Value' from [Nokia_MSS]
                                       WHERE CAST([Date] AS DATE)>='" + start_date_str + "' and CAST([Date] AS DATE) <'" + end_date_str + "' and substring([NE],9,3)= '221' ) tble group by[NE]) tble";


            SqlCommand MSS_VLR_Nokia_Capacity_Quary = new SqlCommand(MSS_VLR_Nokia_Capacity_Quary_String, connection);
            MSS_VLR_Nokia_Capacity_Quary.CommandTimeout = 0;
            MSS_VLR_Nokia_Capacity_Quary.ExecuteNonQuery();
            DataTable MSS_VLR_Nokia_Capacity_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_MSS_VLR_Nokia_Capacity = new SqlDataAdapter(MSS_VLR_Nokia_Capacity_Quary);
            dataAdapter_Contractual_MSS_VLR_Nokia_Capacity.Fill(MSS_VLR_Nokia_Capacity_Table);

            sheet3.Cells[10, 4] = Convert.ToInt64(MSS_VLR_Nokia_Capacity_Table.Rows[0].ItemArray[0]);





            //************MSS SMS - SGs Huawei************
            string MSS_SGs_Huawei_Use_Quary_String = @"select sum([Huawei MSS SGs Use]) as 'Huawei MSS SGs Use'
                        from(
                          select
                               [NE Name],
                               max([Number of SGs Subscribers in VLR (entries)]) as 'Huawei MSS SGs Use'
                               from(select
                                       [Start Time],
                                       [NE Name],
                                       [Number of SGs Subscribers in VLR (entries)] from [H_VLR_SUBSCRIBER]
                                       WHERE CAST([Start Time] AS DATE) >='" + start_date_str + "' and CAST([Start Time] AS DATE)<'" + end_date_str + "') tble group by[NE Name]) tble";


            SqlCommand MSS_SGs_Huawei_Use_Quary = new SqlCommand(MSS_SGs_Huawei_Use_Quary_String, connection);
            MSS_SGs_Huawei_Use_Quary.CommandTimeout = 0;
            MSS_SGs_Huawei_Use_Quary.ExecuteNonQuery();
            DataTable MSS_SGs_Huawei_Use_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_MSS_SGs_Huawei_Use = new SqlDataAdapter(MSS_SGs_Huawei_Use_Quary);
            dataAdapter_Contractual_MSS_SGs_Huawei_Use.Fill(MSS_SGs_Huawei_Use_Table);


            sheet3.Cells[12, 5] = Convert.ToInt64(MSS_SGs_Huawei_Use_Table.Rows[0].ItemArray[0]);

            string MSS_SGs_Huawei_Capacity_Quary_String = @"select sum([SMS over SGs])  as 'Huawei MSS SGs Capacity' from [H_VLR_Capacity]";

            SqlCommand MSS_SGs_Huawei_Capacity_Quary = new SqlCommand(MSS_SGs_Huawei_Capacity_Quary_String, connection);
            MSS_SGs_Huawei_Capacity_Quary.CommandTimeout = 0;
            MSS_SGs_Huawei_Capacity_Quary.ExecuteNonQuery();
            DataTable MSS_SGs_Huawei_Capacity_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_MSS_SGs_Huawei_Capacity = new SqlDataAdapter(MSS_SGs_Huawei_Capacity_Quary);
            dataAdapter_Contractual_MSS_SGs_Huawei_Capacity.Fill(MSS_SGs_Huawei_Capacity_Table);


            sheet3.Cells[10, 5] = Convert.ToInt64(MSS_SGs_Huawei_Capacity_Table.Rows[0].ItemArray[0]);



            //************MSS SMS - SGs Nokia************

            string MSS_SGs_Nokia_Use_Quary_String = @"select sum([Nokia MSS SGs Use]) as 'Nokia MSS SGs Use'
                        from(
                          select
                               [NE],
                               max([Nokia MSS SGs Use Value]) as 'Nokia MSS SGs Use'
                               from(select
                                       [Date],
                                       [NE],
                                       [MSS_Peak_Lic_Utilization(Nokia(Core)]*[MSS_Licence_Capacity(Nokia_Core)] / 100 as 'Nokia MSS SGs Use Value' from[dbo].[Nokia_MSS]
                                       WHERE CAST([Date] AS DATE)>='" + start_date_str + "' and CAST([Date] AS DATE) <'" + end_date_str + "' and substring([NE],9,4)= '1691' ) tble group by[NE]) tble";


            SqlCommand MSS_SGs_Nokia_Use_Quary = new SqlCommand(MSS_SGs_Nokia_Use_Quary_String, connection);
            MSS_SGs_Nokia_Use_Quary.CommandTimeout = 0;
            MSS_SGs_Nokia_Use_Quary.ExecuteNonQuery();
            DataTable MSS_SGs_Nokia_Use_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_MSS_SGs_Nokia_Use = new SqlDataAdapter(MSS_SGs_Nokia_Use_Quary);
            dataAdapter_Contractual_MSS_SGs_Nokia_Use.Fill(MSS_SGs_Nokia_Use_Table);


            sheet3.Cells[12, 6] = Convert.ToInt64(MSS_SGs_Nokia_Use_Table.Rows[0].ItemArray[0]);

            string MSS_SGs_Nokia_Capacity_Quary_String = @"select sum([Nokia MSS SGs Capacity]) as 'Nokia MSS SGs Capacity'
                        from(
                          select
                               [NE],
                               max([Nokia MSS SGs Capacity Value]) as 'Nokia MSS SGs Capacity'
                               from(select
                                       [Date],
                                       [NE],
                                       [MSS_Licence_Capacity(Nokia_Core)] as 'Nokia MSS SGs Capacity Value' from[dbo].[Nokia_MSS]
                                       WHERE CAST([Date] AS DATE)>='" + start_date_str + "' and CAST([Date] AS DATE) <'" + end_date_str + "' and substring([NE],9,4)= '1691' ) tble group by[NE]) tble";

            SqlCommand MSS_SGs_Nokia_Capacity_Quary = new SqlCommand(MSS_SGs_Nokia_Capacity_Quary_String, connection);
            MSS_SGs_Nokia_Capacity_Quary.CommandTimeout = 0;
            MSS_SGs_Nokia_Capacity_Quary.ExecuteNonQuery();
            DataTable MSS_SGs_Nokia_Capacity_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_MSS_SGs_Nokia_Capacity = new SqlDataAdapter(MSS_SGs_Nokia_Capacity_Quary);
            dataAdapter_Contractual_MSS_SGs_Nokia_Capacity.Fill(MSS_SGs_Nokia_Capacity_Table);


            sheet3.Cells[10, 6] = Convert.ToInt64(MSS_SGs_Nokia_Capacity_Table.Rows[0].ItemArray[0]);





            //************ Coonection Capacity Huawei ************

            // MGW
            string Connection_Capacity_Huawei_MGW_Use_Quary_String = @"select sum([Peak Licensed Traffic (number)]) as 'Connection Capacity Huawei Use MGW'
                                        from(
                                          select
                                               [NE Name],
                                               max([Peak Licensed Traffic (number)]) as 'Peak Licensed Traffic (number)'
                                               from(select
                                                       [start TIME],
                                                       [NE Name],
                                                       [Peak Licensed Traffic (number)] from[dbo].[H_LICENSE_MGW]
                                                       WHERE CAST([start TIME] AS DATE) >='" + start_date_str + "' and CAST([start TIME] AS DATE)<'" + end_date_str + "' and substring([NE Name],1,2)= 'MG' ) tble group by[ne Name]) tble";


            SqlCommand Connection_Capacity_Huawei_MGW_Use_Quary = new SqlCommand(Connection_Capacity_Huawei_MGW_Use_Quary_String, connection);
            Connection_Capacity_Huawei_MGW_Use_Quary.CommandTimeout = 0;
            Connection_Capacity_Huawei_MGW_Use_Quary.ExecuteNonQuery();
            DataTable Connection_Capacity_Huawei_MGW_Use_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_Connection_Capacity_Huawei_MGW_Use = new SqlDataAdapter(Connection_Capacity_Huawei_MGW_Use_Quary);
            dataAdapter_Contractual_Connection_Capacity_Huawei_MGW_Use.Fill(Connection_Capacity_Huawei_MGW_Use_Table);

            sheet3.Cells[12, 13] = Convert.ToInt64(Connection_Capacity_Huawei_MGW_Use_Table.Rows[0].ItemArray[0]);

            string Connection_Capacity_Huawei_MGW_Capacity_Quary_String = @"select sum([Authorized Traffic License]) as 'Connection Capacity Huawei Capacity MGW' from[dbo].[H_Authorized_Traffic_License] where substring([MGW Name],1,2)= 'MG'";

            SqlCommand Connection_Capacity_Huawei_MGW_Capacity_Quary = new SqlCommand(Connection_Capacity_Huawei_MGW_Capacity_Quary_String, connection);
            Connection_Capacity_Huawei_MGW_Capacity_Quary.CommandTimeout = 0;
            Connection_Capacity_Huawei_MGW_Capacity_Quary.ExecuteNonQuery();
            DataTable Connection_Capacity_Huawei_MGW_Capacity_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_Connection_Capacity_Huawei_MGW_Capacity = new SqlDataAdapter(Connection_Capacity_Huawei_MGW_Capacity_Quary);
            dataAdapter_Contractual_Connection_Capacity_Huawei_MGW_Capacity.Fill(Connection_Capacity_Huawei_MGW_Capacity_Table);

            sheet3.Cells[10, 13] = Convert.ToInt64(Connection_Capacity_Huawei_MGW_Capacity_Table.Rows[0].ItemArray[0]);






            // TGW
            string Connection_Capacity_Huawei_TGW_Use_Quary_String = @"select sum([Peak Licensed Traffic (number)]) as 'Connection Capacity Huawei Use TGW'
                                        from(
                                          select
                                               [NE Name],
                                               max([Peak Licensed Traffic (number)]) as 'Peak Licensed Traffic (number)'
                                               from(select
                                                       [start TIME],
                                                       [NE Name],
                                                       [Peak Licensed Traffic (number)] from[dbo].[H_LICENSE_MGW]
                                                       WHERE CAST([start TIME] AS DATE) >='" + start_date_str + "' and CAST([start TIME] AS DATE)<'" + end_date_str + "' and substring([NE Name],1,2)= 'TG' ) tble group by[ne Name]) tble";


            SqlCommand Connection_Capacity_Huawei_TGW_Use_Quary = new SqlCommand(Connection_Capacity_Huawei_TGW_Use_Quary_String, connection);
            Connection_Capacity_Huawei_TGW_Use_Quary.CommandTimeout = 0;
            Connection_Capacity_Huawei_TGW_Use_Quary.ExecuteNonQuery();
            DataTable Connection_Capacity_Huawei_TGW_Use_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_Connection_Capacity_Huawei_TGW_Use = new SqlDataAdapter(Connection_Capacity_Huawei_TGW_Use_Quary);
            dataAdapter_Contractual_Connection_Capacity_Huawei_TGW_Use.Fill(Connection_Capacity_Huawei_TGW_Use_Table);

            sheet3.Cells[12, 27] = Convert.ToInt64(Connection_Capacity_Huawei_TGW_Use_Table.Rows[0].ItemArray[0]);

            string Connection_Capacity_Huawei_TGW_Capacity_Quary_String = @"select sum([Authorized Traffic License]) as 'Connection Capacity Huawei Capacity TGW' from[dbo].[H_Authorized_Traffic_License] where substring([MGW Name],1,2)= 'TG'";

            SqlCommand Connection_Capacity_Huawei_TGW_Capacity_Quary = new SqlCommand(Connection_Capacity_Huawei_TGW_Capacity_Quary_String, connection);
            Connection_Capacity_Huawei_TGW_Capacity_Quary.CommandTimeout = 0;
            Connection_Capacity_Huawei_TGW_Capacity_Quary.ExecuteNonQuery();
            DataTable Connection_Capacity_Huawei_TGW_Capacity_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_Connection_Capacity_Huawei_TGW_Capacity = new SqlDataAdapter(Connection_Capacity_Huawei_TGW_Capacity_Quary);
            dataAdapter_Contractual_Connection_Capacity_Huawei_TGW_Capacity.Fill(Connection_Capacity_Huawei_TGW_Capacity_Table);

            sheet3.Cells[10, 27] = Convert.ToInt64(Connection_Capacity_Huawei_TGW_Capacity_Table.Rows[0].ItemArray[0]);





            //************ Coonection Capacity Nokia ************


            // MGW
            string Connection_Capacity_Nokia_MGW_Use_Quary_String = @"select sum([Nokia Connection Capacity]) as 'Nokia Connection Capacity Use MGW'
                        from(
                          select
                               [NE],
                               max([Nokia Connection Capacity]) as 'Nokia Connection Capacity'
                               from(select
                                       [Date],
                                       [NE],
                                       [capacity_licence_utilization_Peak] *[CC_Feature_capacity_Nokia] / 100 as 'Nokia Connection Capacity' from[dbo].[Nokia_MGW]
                                       WHERE CAST([Date] AS DATE) >='" + start_date_str + "' and CAST([Date] AS DATE)<'" + end_date_str + "' and substring([NE],1,2)= 'MG' ) tble group by[NE]) tble";


            SqlCommand Connection_Capacity_Nokia_MGW_Use_Quary = new SqlCommand(Connection_Capacity_Nokia_MGW_Use_Quary_String, connection);
            Connection_Capacity_Nokia_MGW_Use_Quary.CommandTimeout = 0;
            Connection_Capacity_Nokia_MGW_Use_Quary.ExecuteNonQuery();
            DataTable Connection_Capacity_Nokia_MGW_Use_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_Connection_Capacity_Nokia_MGW_Use = new SqlDataAdapter(Connection_Capacity_Nokia_MGW_Use_Quary);
            dataAdapter_Contractual_Connection_Capacity_Nokia_MGW_Use.Fill(Connection_Capacity_Nokia_MGW_Use_Table);

            sheet3.Cells[12, 14] = Convert.ToInt64(Connection_Capacity_Nokia_MGW_Use_Table.Rows[0].ItemArray[0]);

            string Connection_Capacity_Nokia_MGW_Capacity_Quary_String = @"select sum([Nokia Connection Capacity]) as 'Nokia Connection Capacity Licence MGW'
                        from(
                          select
                               [NE],
                               max([Nokia Connection Capacity]) as 'Nokia Connection Capacity'
                               from(select
                                       [Date],
                                       [NE],
                                       [CC_Feature_capacity_Nokia] as 'Nokia Connection Capacity' from[dbo].[Nokia_MGW]
                                       WHERE CAST([Date] AS DATE) >='" + start_date_str + "' and CAST([Date] AS DATE)<'" + end_date_str + "' and substring([NE],1,2)= 'MG' ) tble group by[NE]) tble";

            SqlCommand Connection_Capacity_Nokia_MGW_Capacity_Quary = new SqlCommand(Connection_Capacity_Nokia_MGW_Capacity_Quary_String, connection);
            Connection_Capacity_Nokia_MGW_Capacity_Quary.CommandTimeout = 0;
            Connection_Capacity_Nokia_MGW_Capacity_Quary.ExecuteNonQuery();
            DataTable Connection_Capacity_Nokia_MGW_Capacity_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_Connection_Capacity_Nokia_MGW_Capacity = new SqlDataAdapter(Connection_Capacity_Nokia_MGW_Capacity_Quary);
            dataAdapter_Contractual_Connection_Capacity_Nokia_MGW_Capacity.Fill(Connection_Capacity_Nokia_MGW_Capacity_Table);

            sheet3.Cells[10, 14] = Convert.ToInt64(Connection_Capacity_Nokia_MGW_Capacity_Table.Rows[0].ItemArray[0]);






            // TGW
            string Connection_Capacity_Nokia_TGW_Use_Quary_String = @"select sum([Nokia Connection Capacity]) as 'Nokia Connection Capacity Use TGW'
                        from(
                          select
                               [NE],
                               max([Nokia Connection Capacity]) as 'Nokia Connection Capacity'
                               from(select
                                       [Date],
                                       [NE],
                                       [capacity_licence_utilization_Peak] *[CC_Feature_capacity_Nokia] / 100 as 'Nokia Connection Capacity' from[dbo].[Nokia_MGW]
                                       WHERE CAST([Date] AS DATE) >='" + start_date_str + "' and CAST([Date] AS DATE)<'" + end_date_str + "' and substring([NE],1,2)= 'TG' ) tble group by[NE]) tble";


            SqlCommand Connection_Capacity_Nokia_TGW_Use_Quary = new SqlCommand(Connection_Capacity_Nokia_TGW_Use_Quary_String, connection);
            Connection_Capacity_Nokia_TGW_Use_Quary.CommandTimeout = 0;
            Connection_Capacity_Nokia_TGW_Use_Quary.ExecuteNonQuery();
            DataTable Connection_Capacity_Nokia_TGW_Use_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_Connection_Capacity_Nokia_TGW_Use = new SqlDataAdapter(Connection_Capacity_Nokia_TGW_Use_Quary);
            dataAdapter_Contractual_Connection_Capacity_Nokia_TGW_Use.Fill(Connection_Capacity_Nokia_TGW_Use_Table);

            sheet3.Cells[12, 28] = Convert.ToInt64(Connection_Capacity_Nokia_TGW_Use_Table.Rows[0].ItemArray[0]);

            string Connection_Capacity_Nokia_TGW_Capacity_Quary_String = @"select sum([Nokia Connection Capacity]) as 'Nokia Connection Capacity Licence TGW'
                        from(
                          select
                               [NE],
                               max([Nokia Connection Capacity]) as 'Nokia Connection Capacity'
                               from(select
                                       [Date],
                                       [NE],
                                       [CC_Feature_capacity_Nokia] as 'Nokia Connection Capacity' from[dbo].[Nokia_MGW]
                                       WHERE CAST([Date] AS DATE) >='" + start_date_str + "' and CAST([Date] AS DATE)<'" + end_date_str + "' and substring([NE],1,2)= 'TG' ) tble group by[NE]) tble";

            SqlCommand Connection_Capacity_Nokia_TGW_Capacity_Quary = new SqlCommand(Connection_Capacity_Nokia_TGW_Capacity_Quary_String, connection);
            Connection_Capacity_Nokia_TGW_Capacity_Quary.CommandTimeout = 0;
            Connection_Capacity_Nokia_TGW_Capacity_Quary.ExecuteNonQuery();
            DataTable Connection_Capacity_Nokia_TGW_Capacity_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_Connection_Capacity_Nokia_TGW_Capacity = new SqlDataAdapter(Connection_Capacity_Nokia_TGW_Capacity_Quary);
            dataAdapter_Contractual_Connection_Capacity_Nokia_TGW_Capacity.Fill(Connection_Capacity_Nokia_TGW_Capacity_Table);

            sheet3.Cells[10, 28] = Convert.ToInt64(Connection_Capacity_Nokia_TGW_Capacity_Table.Rows[0].ItemArray[0]);






            //************IUCS Nokia MGW ************
            string IUCS_Nokia_MGW_Use_Quary_String = @"select sum([Nokia IU Capacity]) as 'Nokia IU Capacity Use MGW'
                        from(
                          select
                               [NE],
                               max([Nokia IU Capacity]) as 'Nokia IU Capacity'
                               from(select
                                       [Date],
                                       [NE],
                                       [IU_IP_PEAK] *[IU_IP_feature_capacity] / 100 as 'Nokia IU Capacity' from[dbo].[Nokia_MGW]
                                       WHERE CAST([Date] AS DATE) >='" + start_date_str + "' and CAST([Date] AS DATE)<'" + end_date_str + "' and substring([NE],1,2)= 'MG' ) tble group by[NE]) tble";


            SqlCommand IUCS_Nokia_MGW_Use_Quary = new SqlCommand(IUCS_Nokia_MGW_Use_Quary_String, connection);
            IUCS_Nokia_MGW_Use_Quary.CommandTimeout = 0;
            IUCS_Nokia_MGW_Use_Quary.ExecuteNonQuery();
            DataTable IUCS_Nokia_MGW_Use_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_IUCS_Nokia_MGW_Use = new SqlDataAdapter(IUCS_Nokia_MGW_Use_Quary);
            dataAdapter_Contractual_IUCS_Nokia_MGW_Use.Fill(IUCS_Nokia_MGW_Use_Table);

            sheet3.Cells[12, 16] = Convert.ToInt64(IUCS_Nokia_MGW_Use_Table.Rows[0].ItemArray[0]);

            string IUCS_Nokia_MGW_Capacity_Quary_String = @"select sum([Nokia IU Capacity]) as 'Nokia IU Capacity Licence MGW'
                        from(
                          select
                               [NE],
                               max([Nokia IU Capacity]) as 'Nokia IU Capacity'
                               from(select
                                       [Date],
                                       [NE],
                                       [IU_IP_feature_capacity] as 'Nokia IU Capacity' from[dbo].[Nokia_MGW]
                                       WHERE CAST([Date] AS DATE) >='" + start_date_str + "' and CAST([Date] AS DATE)<'" + end_date_str + "' and substring([NE],1,2)= 'MG' ) tble group by[NE]) tble";

            SqlCommand IUCS_Nokia_MGW_Capacity_Quary = new SqlCommand(IUCS_Nokia_MGW_Capacity_Quary_String, connection);
            IUCS_Nokia_MGW_Capacity_Quary.CommandTimeout = 0;
            IUCS_Nokia_MGW_Capacity_Quary.ExecuteNonQuery();
            DataTable IUCS_Nokia_MGW_Capacity_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_IUCS_Nokia_MGW_Capacity = new SqlDataAdapter(IUCS_Nokia_MGW_Capacity_Quary);
            dataAdapter_Contractual_IUCS_Nokia_MGW_Capacity.Fill(IUCS_Nokia_MGW_Capacity_Table);

            sheet3.Cells[10, 16] = Convert.ToInt64(IUCS_Nokia_MGW_Capacity_Table.Rows[0].ItemArray[0]);






            //************NB Nokia MGW ************
            string NB_Nokia_MGW_Use_Quary_String = @"select sum([Nokia NB Capacity]) as 'Nokia NB Capacity Use MGW'
                        from(
                          select
                               [NE],
                               max([Nokia NB Capacity]) as 'Nokia NB Capacity'
                               from(select
                                       [Date],
                                       [NE],
                                       [NB_IP_FEATURE_CAPACITY]*[NB_IP_Peak_License] / 100 as 'Nokia NB Capacity' from[dbo].[Nokia_MGW]
                                       WHERE CAST([Date] AS DATE) >='" + start_date_str + "' and CAST([Date] AS DATE)<'" + end_date_str + "' and substring([NE],1,2)= 'MG' ) tble group by[NE]) tble";


            SqlCommand NB_Nokia_MGW_Use_Quary = new SqlCommand(NB_Nokia_MGW_Use_Quary_String, connection);
            NB_Nokia_MGW_Use_Quary.CommandTimeout = 0;
            NB_Nokia_MGW_Use_Quary.ExecuteNonQuery();
            DataTable NB_Nokia_MGW_Use_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_NB_Nokia_MGW_Use = new SqlDataAdapter(NB_Nokia_MGW_Use_Quary);
            dataAdapter_Contractual_NB_Nokia_MGW_Use.Fill(NB_Nokia_MGW_Use_Table);

            sheet3.Cells[12, 18] = Convert.ToInt64(NB_Nokia_MGW_Use_Table.Rows[0].ItemArray[0]);

            string NB_Nokia_MGW_Capacity_Quary_String = @"select sum([Nokia NB Capacity]) as 'Nokia NB Capacity Licence MGW'
                        from(
                          select
                               [NE],
                               max([Nokia NB Capacity]) as 'Nokia NB Capacity'
                               from(select
                                       [Date],
                                       [NE],
                                       [NB_IP_FEATURE_CAPACITY] as 'Nokia NB Capacity' from[dbo].[Nokia_MGW]
                                       WHERE CAST([Date] AS DATE) >='" + start_date_str + "' and CAST([Date] AS DATE)<'" + end_date_str + "' and substring([NE],1,2)= 'MG' ) tble group by[NE]) tble";


            SqlCommand NB_Nokia_MGW_Capacity_Quary = new SqlCommand(NB_Nokia_MGW_Capacity_Quary_String, connection);
            NB_Nokia_MGW_Capacity_Quary.CommandTimeout = 0;
            NB_Nokia_MGW_Capacity_Quary.ExecuteNonQuery();
            DataTable NB_Nokia_MGW_Capacity_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_NB_Nokia_MGW_Capacity = new SqlDataAdapter(NB_Nokia_MGW_Capacity_Quary);
            dataAdapter_Contractual_NB_Nokia_MGW_Capacity.Fill(NB_Nokia_MGW_Capacity_Table);

            sheet3.Cells[10, 18] = Convert.ToInt64(NB_Nokia_MGW_Capacity_Table.Rows[0].ItemArray[0]);




            //************AOIP Nokia MGW ************
            string AOIP_Nokia_MGW_Use_Quary_String = @"select sum([Nokia AOIP Capacity]) as 'Nokia AOIP Capacity Use MGW'
                        from(
                          select
                               [NE],
                               max([Nokia AOIP Capacity]) as 'Nokia AOIP Capacity'
                               from(select
                                       [Date],
                                       [NE],
                                       [AOIP_feature_capacity]*[AOIP_peak] / 100 as 'Nokia AOIP Capacity' from[dbo].[Nokia_MGW]
                                       WHERE CAST([Date] AS DATE) >='" + start_date_str + "' and CAST([Date] AS DATE)<'" + end_date_str + "' and substring([NE],1,2)= 'MG' ) tble group by[NE]) tble";


            SqlCommand AOIP_Nokia_MGW_Use_Quary = new SqlCommand(AOIP_Nokia_MGW_Use_Quary_String, connection);
            AOIP_Nokia_MGW_Use_Quary.CommandTimeout = 0;
            AOIP_Nokia_MGW_Use_Quary.ExecuteNonQuery();
            DataTable AOIP_Nokia_MGW_Use_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_AOIP_Nokia_MGW_Use = new SqlDataAdapter(AOIP_Nokia_MGW_Use_Quary);
            dataAdapter_Contractual_AOIP_Nokia_MGW_Use.Fill(AOIP_Nokia_MGW_Use_Table);

            sheet3.Cells[12, 20] = Convert.ToInt64(AOIP_Nokia_MGW_Use_Table.Rows[0].ItemArray[0]);

            string AOIP_Nokia_MGW_Capacity_Quary_String = @"select sum([Nokia AOIP Capacity]) as 'Nokia AOIP Capacity Licence MGW'
                        from(
                          select
                               [NE],
                               max([Nokia AOIP Capacity]) as 'Nokia AOIP Capacity'
                               from(select
                                       [Date],
                                       [NE],
                                       [AOIP_feature_capacity] as 'Nokia AOIP Capacity' from[dbo].[Nokia_MGW]
                                       WHERE CAST([Date] AS DATE) >='" + start_date_str + "' and CAST([Date] AS DATE)<'" + end_date_str + "' and substring([NE],1,2)= 'MG' ) tble group by[NE]) tble";


            SqlCommand AOIP_Nokia_MGW_Capacity_Quary = new SqlCommand(AOIP_Nokia_MGW_Capacity_Quary_String, connection);
            AOIP_Nokia_MGW_Capacity_Quary.CommandTimeout = 0;
            AOIP_Nokia_MGW_Capacity_Quary.ExecuteNonQuery();
            DataTable AOIP_Nokia_MGW_Capacity_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_AOIP_Nokia_MGW_Capacity = new SqlDataAdapter(AOIP_Nokia_MGW_Capacity_Quary);
            dataAdapter_Contractual_AOIP_Nokia_MGW_Capacity.Fill(AOIP_Nokia_MGW_Capacity_Table);

            sheet3.Cells[10, 20] = Convert.ToInt64(AOIP_Nokia_MGW_Capacity_Table.Rows[0].ItemArray[0]);






            //************MB Nokia MGW ************
            string MB_Nokia_MGW_Use_Quary_String = @"select sum([Nokia MB Capacity]) as 'Nokia MB Capacity Use MGW'
                        from(
                          select
                               [NE],
                               max([Nokia MB Capacity]) as 'Nokia MB Capacity'
                               from(select
                                       [Date],
                                       [NE],
                                       [MB_FEATURE_CAPACITY]*[MB_Peak_License] / 100 as 'Nokia MB Capacity' from[dbo].[Nokia_MGW]
                                       WHERE CAST([Date] AS DATE) >='" + start_date_str + "' and CAST([Date] AS DATE)<'" + end_date_str + "' and substring([NE],1,2)= 'MG' ) tble group by[NE]) tble";


            SqlCommand MB_Nokia_MGW_Use_Quary = new SqlCommand(MB_Nokia_MGW_Use_Quary_String, connection);
            MB_Nokia_MGW_Use_Quary.CommandTimeout = 0;
            MB_Nokia_MGW_Use_Quary.ExecuteNonQuery();
            DataTable MB_Nokia_MGW_Use_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_MB_Nokia_MGW_Use = new SqlDataAdapter(MB_Nokia_MGW_Use_Quary);
            dataAdapter_Contractual_MB_Nokia_MGW_Use.Fill(MB_Nokia_MGW_Use_Table);

            sheet3.Cells[12, 22] = Convert.ToInt64(MB_Nokia_MGW_Use_Table.Rows[0].ItemArray[0]);

            string MB_Nokia_MGW_Capacity_Quary_String = @"select sum([Nokia MB Capacity]) as 'Nokia MB Capacity Licence MGW'
                        from(
                          select
                               [NE],
                               max([Nokia MB Capacity]) as 'Nokia MB Capacity'
                               from(select
                                       [Date],
                                       [NE],
                                       [MB_FEATURE_CAPACITY] as 'Nokia MB Capacity' from[dbo].[Nokia_MGW]
                                       WHERE CAST([Date] AS DATE) >='" + start_date_str + "' and CAST([Date] AS DATE)<'" + end_date_str + "' and substring([NE],1,2)= 'MG' ) tble group by[NE]) tble";


            SqlCommand MB_Nokia_MGW_Capacity_Quary = new SqlCommand(MB_Nokia_MGW_Capacity_Quary_String, connection);
            MB_Nokia_MGW_Capacity_Quary.CommandTimeout = 0;
            MB_Nokia_MGW_Capacity_Quary.ExecuteNonQuery();
            DataTable MB_Nokia_MGW_Capacity_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_MB_Nokia_MGW_Capacity = new SqlDataAdapter(MB_Nokia_MGW_Capacity_Quary);
            dataAdapter_Contractual_MB_Nokia_MGW_Capacity.Fill(MB_Nokia_MGW_Capacity_Table);

            sheet3.Cells[10, 22] = Convert.ToInt64(MB_Nokia_MGW_Capacity_Table.Rows[0].ItemArray[0]);




            //************Ater Nokia MGW ************
            string Ater_Nokia_MGW_Use_Quary_String = @"select sum([Nokia Ater Capacity]) as 'Nokia Ater Capacity Use MGW'
                        from(
                          select
                               [NE],
                               max([Nokia Ater Capacity]) as 'Nokia Ater Capacity'
                               from(select
                                       [Date],
                                       [NE],
                                       [PC_ATER_FEATURE_CAPACITY(Nokia_Core)]*[Ater_Peak_License] / 100 as 'Nokia Ater Capacity' from[dbo].[Nokia_MGW]
                                       WHERE CAST([Date] AS DATE) >='" + start_date_str + "' and CAST([Date] AS DATE)<'" + end_date_str + "' and substring([NE],1,2)= 'MG' ) tble group by[NE]) tble";


            SqlCommand Ater_Nokia_MGW_Use_Quary = new SqlCommand(Ater_Nokia_MGW_Use_Quary_String, connection);
            Ater_Nokia_MGW_Use_Quary.CommandTimeout = 0;
            Ater_Nokia_MGW_Use_Quary.ExecuteNonQuery();
            DataTable Ater_Nokia_MGW_Use_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_Ater_Nokia_MGW_Use = new SqlDataAdapter(Ater_Nokia_MGW_Use_Quary);
            dataAdapter_Contractual_Ater_Nokia_MGW_Use.Fill(Ater_Nokia_MGW_Use_Table);

            sheet3.Cells[12, 24] = Convert.ToInt64(Ater_Nokia_MGW_Use_Table.Rows[0].ItemArray[0]);

            string Ater_Nokia_MGW_Capacity_Quary_String = @"select sum([Nokia Ater Capacity]) as 'Nokia Ater Capacity Licence MGW'
                        from(
                          select
                               [NE],
                               max([Nokia Ater Capacity]) as 'Nokia Ater Capacity'
                               from(select
                                       [Date],
                                       [NE],
                                       [PC_ATER_FEATURE_CAPACITY(Nokia_Core)] as 'Nokia Ater Capacity' from[dbo].[Nokia_MGW]
                                       WHERE CAST([Date] AS DATE) >='" + start_date_str + "' and CAST([Date] AS DATE)<'" + end_date_str + "' and substring([NE],1,2)= 'MG' ) tble group by[NE]) tble";


            SqlCommand Ater_Nokia_MGW_Capacity_Quary = new SqlCommand(Ater_Nokia_MGW_Capacity_Quary_String, connection);
            Ater_Nokia_MGW_Capacity_Quary.CommandTimeout = 0;
            Ater_Nokia_MGW_Capacity_Quary.ExecuteNonQuery();
            DataTable Ater_Nokia_MGW_Capacity_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_Ater_Nokia_MGW_Capacity = new SqlDataAdapter(Ater_Nokia_MGW_Capacity_Quary);
            dataAdapter_Contractual_Ater_Nokia_MGW_Capacity.Fill(Ater_Nokia_MGW_Capacity_Table);

            sheet3.Cells[10, 24] = Convert.ToInt64(Ater_Nokia_MGW_Capacity_Table.Rows[0].ItemArray[0]);






            //************NB Nokia TGW ************
            string NB_Nokia_TGW_Use_Quary_String = @"select sum([Nokia NB Capacity]) as 'Nokia NB Capacity Use TGW'
                        from(
                          select
                               [NE],
                               max([Nokia NB Capacity]) as 'Nokia NB Capacity'
                               from(select
                                       [Date],
                                       [NE],
                                       [NB_IP_FEATURE_CAPACITY]*[NB_IP_Peak_License] / 100 as 'Nokia NB Capacity' from[dbo].[Nokia_MGW]
                                       WHERE CAST([Date] AS DATE) >='" + start_date_str + "' and CAST([Date] AS DATE)<'" + end_date_str + "' and substring([NE],1,2)= 'TG' ) tble group by[NE]) tble";


            SqlCommand NB_Nokia_TGW_Use_Quary = new SqlCommand(NB_Nokia_TGW_Use_Quary_String, connection);
            NB_Nokia_TGW_Use_Quary.CommandTimeout = 0;
            NB_Nokia_TGW_Use_Quary.ExecuteNonQuery();
            DataTable NB_Nokia_TGW_Use_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_NB_Nokia_TGW_Use = new SqlDataAdapter(NB_Nokia_TGW_Use_Quary);
            dataAdapter_Contractual_NB_Nokia_TGW_Use.Fill(NB_Nokia_TGW_Use_Table);

            sheet3.Cells[12, 30] = Convert.ToInt64(NB_Nokia_TGW_Use_Table.Rows[0].ItemArray[0]);

            string NB_Nokia_TGW_Capacity_Quary_String = @"select sum([Nokia NB Capacity]) as 'Nokia NB Capacity Licence TGW'
                        from(
                          select
                               [NE],
                               max([Nokia NB Capacity]) as 'Nokia NB Capacity'
                               from(select
                                       [Date],
                                       [NE],
                                       [NB_IP_FEATURE_CAPACITY] as 'Nokia NB Capacity' from[dbo].[Nokia_MGW]
                                       WHERE CAST([Date] AS DATE) >='" + start_date_str + "' and CAST([Date] AS DATE)<'" + end_date_str + "' and substring([NE],1,2)= 'TG' ) tble group by[NE]) tble";


            SqlCommand NB_Nokia_TGW_Capacity_Quary = new SqlCommand(NB_Nokia_TGW_Capacity_Quary_String, connection);
            NB_Nokia_TGW_Capacity_Quary.CommandTimeout = 0;
            NB_Nokia_TGW_Capacity_Quary.ExecuteNonQuery();
            DataTable NB_Nokia_TGW_Capacity_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_NB_Nokia_TGW_Capacity = new SqlDataAdapter(NB_Nokia_TGW_Capacity_Quary);
            dataAdapter_Contractual_NB_Nokia_TGW_Capacity.Fill(NB_Nokia_TGW_Capacity_Table);

            sheet3.Cells[10, 30] = Convert.ToInt64(NB_Nokia_TGW_Capacity_Table.Rows[0].ItemArray[0]);




            //************MB Nokia TGW ************
            string MB_Nokia_TGW_Use_Quary_String = @"select sum([Nokia MB Capacity]) as 'Nokia MB Capacity Use TGW'
                        from(
                          select
                               [NE],
                               max([Nokia MB Capacity]) as 'Nokia MB Capacity'
                               from(select
                                       [Date],
                                       [NE],
                                       [MB_FEATURE_CAPACITY]*[MB_Peak_License] / 100 as 'Nokia MB Capacity' from[dbo].[Nokia_MGW]
                                       WHERE CAST([Date] AS DATE) >='" + start_date_str + "' and CAST([Date] AS DATE)<'" + end_date_str + "' and substring([NE],1,2)= 'TG' ) tble group by[NE]) tble";


            SqlCommand MB_Nokia_TGW_Use_Quary = new SqlCommand(MB_Nokia_TGW_Use_Quary_String, connection);
            MB_Nokia_TGW_Use_Quary.CommandTimeout = 0;
            MB_Nokia_TGW_Use_Quary.ExecuteNonQuery();
            DataTable MB_Nokia_TGW_Use_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_MB_Nokia_TGW_Use = new SqlDataAdapter(MB_Nokia_TGW_Use_Quary);
            dataAdapter_Contractual_MB_Nokia_TGW_Use.Fill(MB_Nokia_TGW_Use_Table);

            sheet3.Cells[12, 32] = Convert.ToInt64(MB_Nokia_TGW_Use_Table.Rows[0].ItemArray[0]);

            string MB_Nokia_TGW_Capacity_Quary_String = @"select sum([Nokia MB Capacity]) as 'Nokia MB Capacity Licence TGW'
                        from(
                          select
                               [NE],
                               max([Nokia MB Capacity]) as 'Nokia MB Capacity'
                               from(select
                                       [Date],
                                       [NE],
                                       [MB_FEATURE_CAPACITY] as 'Nokia MB Capacity' from[dbo].[Nokia_MGW]
                                       WHERE CAST([Date] AS DATE) >='" + start_date_str + "' and CAST([Date] AS DATE)<'" + end_date_str + "' and substring([NE],1,2)= 'TG' ) tble group by[NE]) tble";


            SqlCommand MB_Nokia_TGW_Capacity_Quary = new SqlCommand(MB_Nokia_TGW_Capacity_Quary_String, connection);
            MB_Nokia_TGW_Capacity_Quary.CommandTimeout = 0;
            MB_Nokia_TGW_Capacity_Quary.ExecuteNonQuery();
            DataTable MB_Nokia_TGW_Capacity_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_MB_Nokia_TGW_Capacity = new SqlDataAdapter(MB_Nokia_TGW_Capacity_Quary);
            dataAdapter_Contractual_MB_Nokia_TGW_Capacity.Fill(MB_Nokia_TGW_Capacity_Table);

            sheet3.Cells[10, 32] = Convert.ToInt64(MB_Nokia_TGW_Capacity_Table.Rows[0].ItemArray[0]);






            //-----------------CAMEL PH4 -------------------

            string CAMELPH4_Nokia_Use_Quary_String = @"select sum([Camel Use]) as 'Camel Use Nokia'
                        from(
                          select
                               [MSC name],
                               max([Camel Use]) as 'Camel Use'
                               from(select
                                       [PERIOD_START_TIME],
                                       [MSC name],
                                       [CAP_LIC_LIMIT (M406B2C4)] *[AVERAGE_CAP_LIC_USAGE_X100 (M406B2C2)] / 10000 as 'Camel Use' from N_VLR_Features
                                        WHERE CAST([PERIOD_START_TIME] AS DATE) >='" + start_date_str + "' and CAST([PERIOD_START_TIME] AS DATE)< '" + end_date_str + "' and [FEAC_ID] = '223' ) tble group by[MSC name]) tble";


            SqlCommand CAMELPH4_Nokia_Use_Quary = new SqlCommand(CAMELPH4_Nokia_Use_Quary_String, connection);
            CAMELPH4_Nokia_Use_Quary.CommandTimeout = 0;
            CAMELPH4_Nokia_Use_Quary.ExecuteNonQuery();
            DataTable CAMELPH4_Nokia_Use_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_CAMELPH4_Nokia_Use = new SqlDataAdapter(CAMELPH4_Nokia_Use_Quary);
            dataAdapter_Contractual_CAMELPH4_Nokia_Use.Fill(CAMELPH4_Nokia_Use_Table);

            sheet3.Cells[12, 8] = Convert.ToInt64(CAMELPH4_Nokia_Use_Table.Rows[0].ItemArray[0]);

            string CAMELPH4_Nokia_Capacity_Quary_String = @"select sum([Camel Capacity]) as 'Camel Capacity Nokia'
                        from(
                          select
                               [MSC name],
                               max([Camel Capacity]) as 'Camel Capacity'
                               from(select
                                       [PERIOD_START_TIME],
                                       [MSC name],
                                       [CAP_LIC_LIMIT (M406B2C4)] as 'Camel Capacity' from N_VLR_Features
                                      WHERE CAST([PERIOD_START_TIME] AS DATE) >='" + start_date_str + "' and CAST([PERIOD_START_TIME] AS DATE)< '" + end_date_str + "' and [FEAC_ID] = '223' ) tble group by[MSC name]) tble";


            SqlCommand CAMELPH4_Nokia_Capacity_Quary = new SqlCommand(CAMELPH4_Nokia_Capacity_Quary_String, connection);
            CAMELPH4_Nokia_Capacity_Quary.CommandTimeout = 0;
            CAMELPH4_Nokia_Capacity_Quary.ExecuteNonQuery();
            DataTable CAMELPH4_Nokia_Capacity_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_CAMELPH4_Nokia_Capacity = new SqlDataAdapter(CAMELPH4_Nokia_Capacity_Quary);
            dataAdapter_Contractual_CAMELPH4_Nokia_Capacity.Fill(CAMELPH4_Nokia_Capacity_Table);

            sheet3.Cells[10, 8] = Convert.ToInt64(CAMELPH4_Nokia_Capacity_Table.Rows[0].ItemArray[0]);







            // -----eSRVCC to GERAN---------

            string eSRVCC_GERAN_Huawei_Use_Quary_String = @"select sum([eSRVCC to GERAN]) as 'Use eSRVCC to GERAN'
                        from(
                          select
                               [NE_name],
                               max([eSRVCC_to_GERAN_Traffic_Erl]) as 'eSRVCC to GERAN'
                               from(select
                                       [Start Time],
                                       [NE_name],
                                       [eSRVCC_to_GERAN_Traffic_Erl] from[dbo].[eSRVCC_MSS_HOST]
                                       WHERE CAST([Start Time] AS DATE) >='" + start_date_str + "' and CAST([Start Time] AS DATE)< '" + end_date_str + "') tble group by[NE_name]) tble1";


            SqlCommand eSRVCC_GERAN_Huawei_Use_Quary = new SqlCommand(eSRVCC_GERAN_Huawei_Use_Quary_String, connection);
            eSRVCC_GERAN_Huawei_Use_Quary.CommandTimeout = 0;
            eSRVCC_GERAN_Huawei_Use_Quary.ExecuteNonQuery();
            DataTable eSRVCC_GERAN_Huawei_Use_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_eSRVCC_GERAN_Huawei_Use = new SqlDataAdapter(eSRVCC_GERAN_Huawei_Use_Quary);
            dataAdapter_Contractual_eSRVCC_GERAN_Huawei_Use.Fill(eSRVCC_GERAN_Huawei_Use_Table);

            sheet3.Cells[12, 9] = Convert.ToInt64(eSRVCC_GERAN_Huawei_Use_Table.Rows[0].ItemArray[0]);

            string eSRVCC_GERAN_Huawei_Capacity_Quary_String = @"select sum([eSRVCC to GERAN]) as 'Capacity eSRVCC to GERAN' from[dbo].[eSRVCC to GERAN license value]";


            SqlCommand eSRVCC_GERAN_Huawei_Capacity_Quary = new SqlCommand(eSRVCC_GERAN_Huawei_Capacity_Quary_String, connection);
            eSRVCC_GERAN_Huawei_Capacity_Quary.CommandTimeout = 0;
            eSRVCC_GERAN_Huawei_Capacity_Quary.ExecuteNonQuery();
            DataTable eSRVCC_GERAN_Huawei_Capacity_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_eSRVCC_GERAN_Huawei_Capacity = new SqlDataAdapter(eSRVCC_GERAN_Huawei_Capacity_Quary);
            dataAdapter_Contractual_eSRVCC_GERAN_Huawei_Capacity.Fill(eSRVCC_GERAN_Huawei_Capacity_Table);

            sheet3.Cells[10, 9] = Convert.ToInt64(eSRVCC_GERAN_Huawei_Capacity_Table.Rows[0].ItemArray[0]);





            // -----eSRVCC to UTRAN---------

            string eSRVCC_UTRAN_Huawei_Use_Quary_String = @"select sum([eSRVCC to UTRAN]) as 'Use eSRVCC to UTRAN'
                        from(
                          select
                               [NE_name],
                               max([eSRVCC_to_UTRAN_Traffic_Erl]) as 'eSRVCC to UTRAN'
                               from(select
                                       [Start Time],
                                       [NE_name],
                                       [eSRVCC_to_UTRAN_Traffic_Erl] from[dbo].[eSRVCC_MSS_HOST]
                                       WHERE CAST([Start Time] AS DATE) >='" + start_date_str + "' and CAST([Start Time] AS DATE)< '" + end_date_str + "') tble group by[NE_name]) tble1";


            SqlCommand eSRVCC_UTRAN_Huawei_Use_Quary = new SqlCommand(eSRVCC_UTRAN_Huawei_Use_Quary_String, connection);
            eSRVCC_UTRAN_Huawei_Use_Quary.CommandTimeout = 0;
            eSRVCC_UTRAN_Huawei_Use_Quary.ExecuteNonQuery();
            DataTable eSRVCC_UTRAN_Huawei_Use_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_eSRVCC_UTRAN_Huawei_Use = new SqlDataAdapter(eSRVCC_UTRAN_Huawei_Use_Quary);
            dataAdapter_Contractual_eSRVCC_UTRAN_Huawei_Use.Fill(eSRVCC_UTRAN_Huawei_Use_Table);

            sheet3.Cells[12, 11] = Convert.ToInt64(eSRVCC_UTRAN_Huawei_Use_Table.Rows[0].ItemArray[0]);

            string eSRVCC_UTRAN_Huawei_Capacity_Quary_String = @"select sum([eSRVCC to UTRAN]) as 'Capacity eSRVCC to UTRAN' from[dbo].[eSRVCC to UTRAN license value]";


            SqlCommand eSRVCC_UTRAN_Huawei_Capacity_Quary = new SqlCommand(eSRVCC_UTRAN_Huawei_Capacity_Quary_String, connection);
            eSRVCC_UTRAN_Huawei_Capacity_Quary.CommandTimeout = 0;
            eSRVCC_UTRAN_Huawei_Capacity_Quary.ExecuteNonQuery();
            DataTable eSRVCC_UTRAN_Huawei_Capacity_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_eSRVCC_UTRAN_Huawei_Capacity = new SqlDataAdapter(eSRVCC_UTRAN_Huawei_Capacity_Quary);
            dataAdapter_Contractual_eSRVCC_UTRAN_Huawei_Capacity.Fill(eSRVCC_UTRAN_Huawei_Capacity_Table);

            sheet3.Cells[10, 11] = Convert.ToInt64(eSRVCC_UTRAN_Huawei_Capacity_Table.Rows[0].ItemArray[0]);



            //@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@



            //************ MSS - VLR  Huawei Utilizaion ****************

            // find Maximum of VLR in a week
            string MSS_VLR_Huawei_Utilization_Quary_String = @"select[NE Name], max([Total Number of Subscribers]) as 'Max of Total Number of Subscribers' from(
            select[Start Time], [NE Name], sum([Total Number of Subscribers (LAI) (entries)]) as 'Total Number of Subscribers' from [H_LAC_MSS]
            WHERE CAST([Start Time] AS DATE)>='" + start_date_str + "' and CAST([Start Time] AS DATE)< '" + end_date_str + "' group by[Start Time], [NE Name]) tble group by [NE Name]";


            SqlCommand MSS_VLR_Huawei_Utilization_Quary = new SqlCommand(MSS_VLR_Huawei_Utilization_Quary_String, connection);
            MSS_VLR_Huawei_Utilization_Quary.CommandTimeout = 0;
            MSS_VLR_Huawei_Utilization_Quary.ExecuteNonQuery();
            DataTable MSS_VLR_Huawei_Utilization_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_MSS_VLR_Huawei_Utilization = new SqlDataAdapter(MSS_VLR_Huawei_Utilization_Quary);
            dataAdapter_Contractual_MSS_VLR_Huawei_Utilization.Fill(MSS_VLR_Huawei_Utilization_Table);



            string MSS_VLR_Huawei_MaxCapacity_Quary_String = @"SELECT * FROM [H_VLR_Capacity]";


            SqlCommand MSS_VLR_Huawei_MaxCapacity_Quary = new SqlCommand(MSS_VLR_Huawei_MaxCapacity_Quary_String, connection);
            MSS_VLR_Huawei_MaxCapacity_Quary.CommandTimeout = 0;
            MSS_VLR_Huawei_MaxCapacity_Quary.ExecuteNonQuery();
            DataTable MSS_VLR_Huawei_MaxCapacity_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_MSS_VLR_Huawei_MaxCapacity = new SqlDataAdapter(MSS_VLR_Huawei_MaxCapacity_Quary);
            dataAdapter_Contractual_MSS_VLR_Huawei_MaxCapacity.Fill(MSS_VLR_Huawei_MaxCapacity_Table);



            // Join with H_VLR_Capacity
            var H_VLR_Q = (from pd in MSS_VLR_Huawei_Utilization_Table.AsEnumerable()
                           join od in MSS_VLR_Huawei_MaxCapacity_Table.AsEnumerable() on new { f1 = pd.Field<string>("NE Name") } equals new { f1 = od.Field<string>("MSS Name") } into od
                           from new_od in od.DefaultIfEmpty()
                           select new
                           {
                               NE = pd.Field<string>("NE Name"),
                               Max_of_Total_Number_of_Subscribers = pd.Field<double>("Max of Total Number of Subscribers"),
                               Maximum_of_Users = new_od.Field<double>("Maximum of Users"),
                               Utilization = 100 * pd.Field<double>("Max of Total Number of Subscribers") / new_od.Field<double>("Maximum of Users"),

                           }).ToList();

            DataTable H_VLR_Table = new DataTable();
            H_VLR_Table = ConvertToDataTable(H_VLR_Q);


            for (int k = 0; k < 19; k++)
            {
                sheet3.Cells[25 + k, 3] = "";
                sheet3.Cells[25 + k, 4] = "";
                sheet3.Cells[25 + k, 5] = "";
                sheet3.Cells[25 + k, 6] = "";
            }



            int Num_H_VLR_High_Utilization = 0;

            for (int k = 0; k < H_VLR_Table.Rows.Count; k++)
            {
                if (H_VLR_Table.Rows[k].ItemArray[3].ToString() != "")
                {
                    if (H_VLR_Table.Rows[k].ItemArray[3].ToString() != "")
                    {
                        double H_VLR_Utilization = Convert.ToDouble(H_VLR_Table.Rows[k].ItemArray[3]);
                        if (H_VLR_Utilization >= 80)
                        {
                            sheet3.Cells[25 + Num_H_VLR_High_Utilization, 3] = H_VLR_Table.Rows[k].ItemArray[0];
                            sheet3.Cells[25 + Num_H_VLR_High_Utilization, 4] = Math.Round(Convert.ToDouble(H_VLR_Table.Rows[k].ItemArray[3]), 2);
                            sheet1.Cells[Num_H_VLR_High_Utilization + 3, 6] = H_VLR_Table.Rows[k].ItemArray[0];
                            sheet1.Cells[Num_H_VLR_High_Utilization + 3, 8] = Math.Round(Convert.ToDouble(H_VLR_Table.Rows[k].ItemArray[3]), 2);
                            Num_H_VLR_High_Utilization++;
                        }
                    }

                }

            }

            sheet2.Cells[2, 1] = "MSS";
            sheet2.Cells[2, 2] = "HUAWEI";
            sheet2.Cells[2, 3] = "VLR LICENSE";
            sheet2.Cells[2, 4] = 80;
            sheet2.Cells[2, 5] = Num_H_VLR_High_Utilization;
            sheet2.Cells[2, 6] = "VLR Utilization of " + Convert.ToString(Num_H_VLR_High_Utilization) + " of HUAWEI MSSs are over Threshhold";

            sheet1.Cells[2, 1] = "MSS";
            sheet1.Cells[2, 2] = "HUAWEI";
            sheet1.Cells[2, 3] = "VLR LICENSE";
            sheet1.Cells[2, 4] = 80;
            sheet1.Cells[2, 5] = Num_H_VLR_High_Utilization;
            sheet1.Cells[2, 9] = "VLR Utilization of " + Convert.ToString(Num_H_VLR_High_Utilization) + " of HUAWEI MSSs are over Threshhold";


            if (Num_H_VLR_High_Utilization >= 10)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[2, i].Interior.Color = Color.Red;
                    }
                    sheet1.Cells[2, i].Interior.Color = Color.Red;
                }
            }
            if (Num_H_VLR_High_Utilization >= 1 && Num_H_VLR_High_Utilization <= 9)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[2, i].Interior.Color = Color.Gray;
                    }
                    sheet1.Cells[2, i].Interior.Color = Color.Gray;
                }
            }
            if (Num_H_VLR_High_Utilization == 0)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[2, i].Interior.Color = Color.LightBlue;
                    }
                    sheet1.Cells[2, i].Interior.Color = Color.LightBlue;
                }
            }

            sheet3.Cells[18, 3] = Convert.ToInt64(Num_H_VLR_High_Utilization);

            checkBox1.Invoke(new Action(() => checkBox1.Checked = true));


            //************ MSS - VLR  Nokia Utilizaion ****************

            string MSS_VLR_Nokia_Utilization_Quary_String = @"select[MSC Name], max([Utilization]) as 'Max of Usilization' from(
            select[PERIOD_START_TIME], [MSC Name], max([PEAK_CAP_LIC_USAGE_X100 (M406B2C3)] / 100) as 'Utilization' from N_VLR_Features
            WHERE CAST([PERIOD_START_TIME] AS DATE)>='" + start_date_str + "' and CAST([PERIOD_START_TIME] AS DATE)< '" + end_date_str + "' and [FEAC_ID] = '221' group by[PERIOD_START_TIME], [MSC Name]) tble group by[MSC Name]";


            SqlCommand MSS_VLR_Nokia_Utilization_Quary = new SqlCommand(MSS_VLR_Nokia_Utilization_Quary_String, connection);
            MSS_VLR_Nokia_Utilization_Quary.CommandTimeout = 0;
            MSS_VLR_Nokia_Utilization_Quary.ExecuteNonQuery();
            DataTable MSS_VLR_Nokia_Utilization_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_MSS_VLR_Nokia_Utilization = new SqlDataAdapter(MSS_VLR_Nokia_Utilization_Quary);
            dataAdapter_Contractual_MSS_VLR_Nokia_Utilization.Fill(MSS_VLR_Nokia_Utilization_Table);


            int Num_N_VLR_High_Utilization = 0;

            for (int k = 0; k < MSS_VLR_Nokia_Utilization_Table.Rows.Count; k++)
            {
                if (MSS_VLR_Nokia_Utilization_Table.Rows[k].ItemArray[1].ToString() != "")
                {
                    if (MSS_VLR_Nokia_Utilization_Table.Rows[k].ItemArray[1].ToString() != "")
                    {
                        double N_VLR_Utilization = Convert.ToDouble(MSS_VLR_Nokia_Utilization_Table.Rows[k].ItemArray[1]);
                        if (N_VLR_Utilization >= 80)
                        {
                            sheet3.Cells[25 + Num_N_VLR_High_Utilization + Num_H_VLR_High_Utilization, 3] = MSS_VLR_Nokia_Utilization_Table.Rows[k].ItemArray[0];
                            sheet3.Cells[25 + Num_N_VLR_High_Utilization + Num_H_VLR_High_Utilization, 4] = Math.Round(Convert.ToDouble(MSS_VLR_Nokia_Utilization_Table.Rows[k].ItemArray[1]), 2);
                            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + 4, 6] = MSS_VLR_Nokia_Utilization_Table.Rows[k].ItemArray[0];
                            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + 4, 8] = Math.Round(Convert.ToDouble(MSS_VLR_Nokia_Utilization_Table.Rows[k].ItemArray[1]), 2);
                            Num_N_VLR_High_Utilization++;
                        }
                    }

                }

            }

            sheet2.Cells[3, 1] = "MSS";
            sheet2.Cells[3, 2] = "NOKIA";
            sheet2.Cells[3, 3] = "VLR LICENSE";
            sheet2.Cells[3, 4] = 80;
            sheet2.Cells[3, 5] = Num_N_VLR_High_Utilization;
            sheet2.Cells[3, 6] = "VLR Utilization of " + Convert.ToString(Num_N_VLR_High_Utilization) + " of NOKIA MSSs are over Threshhold";

            sheet1.Cells[Num_H_VLR_High_Utilization + 3, 1] = "MSS";
            sheet1.Cells[Num_H_VLR_High_Utilization + 3, 2] = "NOKIA";
            sheet1.Cells[Num_H_VLR_High_Utilization + 3, 3] = "VLR LICENSE";
            sheet1.Cells[Num_H_VLR_High_Utilization + 3, 4] = 80;
            sheet1.Cells[Num_H_VLR_High_Utilization + 3, 5] = Num_N_VLR_High_Utilization;
            sheet1.Cells[Num_H_VLR_High_Utilization + 3, 9] = "VLR Utilization of " + Convert.ToString(Num_N_VLR_High_Utilization) + " of NOKIA MSSs are over Threshhold";


            if (Num_N_VLR_High_Utilization >= 10)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[3, i].Interior.Color = Color.Red;
                    }
                    sheet1.Cells[3 + Num_H_VLR_High_Utilization, i].Interior.Color = Color.Red;
                }
            }
            if (Num_N_VLR_High_Utilization >= 1 && Num_N_VLR_High_Utilization <= 9)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[3, i].Interior.Color = Color.Gray;
                    }
                    sheet1.Cells[3 + Num_H_VLR_High_Utilization, i].Interior.Color = Color.Gray;
                }
            }
            if (Num_N_VLR_High_Utilization == 0)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[3, i].Interior.Color = Color.LightBlue;
                    }
                    sheet1.Cells[3 + Num_H_VLR_High_Utilization, i].Interior.Color = Color.LightBlue;
                }
            }

            sheet3.Cells[18, 4] = Convert.ToInt64(Num_N_VLR_High_Utilization);

            checkBox2.Invoke(new Action(() => checkBox2.Checked = true));

            //************ MSS - SGs  Huawei Utilizaion ****************

            // find Maximum of SGs in a week
            string MSS_SGs_Huawei_Utilization_Quary_String = @"select[NE Name], max([Total Number of Subscribers]) as 'Max of Total Number of Subscribers' from(
            select[Start Time], [NE Name], sum([Number of SGs Subscribers in VLR (entries)]) as 'Total Number of Subscribers' from [H_SG_SUB_MSS]
            WHERE (CAST([Start Time] AS DATE)>='" + start_date_str + "' and CAST([Start Time] AS DATE)< '" + end_date_str + "' and substring([NE Name],1,2)!='TS') group by[Start Time], [NE Name]) tble group by [NE Name]";


            SqlCommand MSS_SGs_Huawei_Utilization_Quary = new SqlCommand(MSS_SGs_Huawei_Utilization_Quary_String, connection);
            MSS_SGs_Huawei_Utilization_Quary.CommandTimeout = 0;
            MSS_SGs_Huawei_Utilization_Quary.ExecuteNonQuery();
            DataTable MSS_SGs_Huawei_Utilization_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_MSS_SGs_Huawei_Utilization = new SqlDataAdapter(MSS_SGs_Huawei_Utilization_Quary);
            dataAdapter_Contractual_MSS_SGs_Huawei_Utilization.Fill(MSS_SGs_Huawei_Utilization_Table);



            // Join with H_SGs_Capacity
            var H_SGs_Q = (from pd in MSS_SGs_Huawei_Utilization_Table.AsEnumerable()
                           join od in MSS_VLR_Huawei_MaxCapacity_Table.AsEnumerable() on new { f1 = pd.Field<string>("NE Name") } equals new { f1 = od.Field<string>("MSS Name") } into od
                           from new_od in od.DefaultIfEmpty()
                           select new
                           {
                               NE = pd.Field<string>("NE Name"),
                               Max_of_Total_Number_of_Subscribers = pd.Field<double>("Max of Total Number of Subscribers"),
                               Maximum_of_Users = new_od.Field<double>("SMS over SGs"),
                               Utilization = 100 * pd.Field<double>("Max of Total Number of Subscribers") / new_od.Field<double>("SMS over SGs"),

                           }).ToList();

            DataTable H_SGs_Table = new DataTable();
            H_SGs_Table = ConvertToDataTable(H_SGs_Q);

            int Num_H_SGs_High_Utilization = 0;

            for (int k = 0; k < H_SGs_Table.Rows.Count; k++)
            {
                if (H_SGs_Table.Rows[k].ItemArray[3].ToString() != "")
                {
                    if (H_SGs_Table.Rows[k].ItemArray[3].ToString() != "")
                    {
                        if (H_SGs_Table.Rows[k].ItemArray[3].ToString() != "")
                        {
                            double H_SGs_Utilization = Convert.ToDouble(H_SGs_Table.Rows[k].ItemArray[3]);
                            if (H_SGs_Utilization >= 80)
                            {
                                sheet3.Cells[25 + Num_H_SGs_High_Utilization, 5] = H_SGs_Table.Rows[k].ItemArray[0];
                                sheet3.Cells[25 + Num_H_SGs_High_Utilization, 6] = Math.Round(Convert.ToDouble(H_SGs_Table.Rows[k].ItemArray[3]), 2);
                                sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + 5, 6] = H_SGs_Table.Rows[k].ItemArray[0];
                                sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + 5, 8] = Math.Round(Convert.ToDouble(H_SGs_Table.Rows[k].ItemArray[3]), 2);
                                Num_H_SGs_High_Utilization++;
                            }
                        }

                    }

                }

            }

            sheet2.Cells[4, 1] = "MSS";
            sheet2.Cells[4, 2] = "HUAWEI";
            sheet2.Cells[4, 3] = "SGs LICENSE";
            sheet2.Cells[4, 4] = 80;
            sheet2.Cells[4, 5] = Num_H_SGs_High_Utilization;
            sheet2.Cells[4, 6] = "SMS over SGs Utilization of " + Convert.ToString(Num_H_SGs_High_Utilization) + " of HUAWEI MSSs are over Threshhold";

            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + 4, 1] = "MSS";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + 4, 2] = "HUAWEI";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + 4, 3] = "SGs LICENSE";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + 4, 4] = 80;
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + 4, 5] = Num_H_SGs_High_Utilization;
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + 4, 9] = "SMS over SGs Utilization of " + Convert.ToString(Num_H_SGs_High_Utilization) + " of HUAWEI MSSs are over Threshhold";


            if (Num_H_SGs_High_Utilization >= 10)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[4, i].Interior.Color = Color.Red;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + 4, i].Interior.Color = Color.Red;
                }
            }
            if (Num_H_SGs_High_Utilization >= 1 && Num_H_SGs_High_Utilization <= 9)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[4, i].Interior.Color = Color.Gray;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + 4, i].Interior.Color = Color.Gray;
                }
            }
            if (Num_H_SGs_High_Utilization == 0)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[4, i].Interior.Color = Color.LightBlue;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + 4, i].Interior.Color = Color.LightBlue;
                }
            }

            sheet3.Cells[18, 5] = Convert.ToInt64(Num_H_SGs_High_Utilization);

            checkBox3.Invoke(new Action(() => checkBox3.Checked = true));

            //************ MSS - SGs  Nokia Utilizaion ****************

            // find Maximum of SGs in a week
            string MSS_SGs_Nokia_Utilization_Quary_String = @"select  [MSC Name], [FEAC_ID], max(([PEAK_CAP_LIC_USAGE_X100 (M406B2C3)]/100)) as [MSS_Peak_Lic_Utilization]	FROM N_VLR_Features
            where CAST([PERIOD_START_TIME] AS DATE)>='" + start_date_str + "' and CAST([PERIOD_START_TIME] AS DATE)< '" + end_date_str + "' and [FEAC_ID]='1691' group by [MSC Name], [FEAC_ID]";


            SqlCommand MSS_SGs_Nokia_Utilization_Quary = new SqlCommand(MSS_SGs_Nokia_Utilization_Quary_String, connection);
            MSS_SGs_Nokia_Utilization_Quary.CommandTimeout = 0;
            MSS_SGs_Nokia_Utilization_Quary.ExecuteNonQuery();
            DataTable MSS_SGs_Nokia_Utilization_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_MSS_SGs_Nokia_Utilization = new SqlDataAdapter(MSS_SGs_Nokia_Utilization_Quary);
            dataAdapter_Contractual_MSS_SGs_Nokia_Utilization.Fill(MSS_SGs_Nokia_Utilization_Table);




            int Num_N_SGs_High_Utilization = 0;

            for (int k = 0; k < MSS_SGs_Nokia_Utilization_Table.Rows.Count; k++)
            {
                if (MSS_SGs_Nokia_Utilization_Table.Rows[k].ItemArray[2].ToString() != "")
                {
                    if (MSS_SGs_Nokia_Utilization_Table.Rows[k].ItemArray[2].ToString() != "")
                    {
                        double N_SGs_Utilization = Convert.ToDouble(MSS_SGs_Nokia_Utilization_Table.Rows[k].ItemArray[2]);
                        if (N_SGs_Utilization >= 80)
                        {
                            sheet3.Cells[25 + Num_N_SGs_High_Utilization + Num_H_SGs_High_Utilization, 5] = MSS_SGs_Nokia_Utilization_Table.Rows[k].ItemArray[0];
                            sheet3.Cells[25 + Num_N_SGs_High_Utilization + Num_H_SGs_High_Utilization, 6] = Math.Round(Convert.ToDouble(MSS_SGs_Nokia_Utilization_Table.Rows[k].ItemArray[2]), 2);
                            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + 6, 6] = MSS_SGs_Nokia_Utilization_Table.Rows[k].ItemArray[0];
                            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + 6, 8] = Math.Round(Convert.ToDouble(MSS_SGs_Nokia_Utilization_Table.Rows[k].ItemArray[2]), 2);
                            Num_N_SGs_High_Utilization++;
                        }
                    }

                }

            }

            sheet2.Cells[5, 1] = "MSS";
            sheet2.Cells[5, 2] = "NOKIA";
            sheet2.Cells[5, 3] = "SGs LICENSE";
            sheet2.Cells[5, 4] = 80;
            sheet2.Cells[5, 5] = Num_N_SGs_High_Utilization;
            sheet2.Cells[5, 6] = "SMS over SGs Utilization of " + Convert.ToString(Num_N_SGs_High_Utilization) + " of HUAWEI MSSs are over Threshhold";

            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + 5, 1] = "MSS";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + 5, 2] = "NOKIA";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + 5, 3] = "SGs LICENSE";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + 5, 4] = 80;
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + 5, 5] = Num_N_SGs_High_Utilization;
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + 5, 9] = "SMS over SGs Utilization of " + Convert.ToString(Num_N_SGs_High_Utilization) + " of HUAWEI MSSs are over Threshhold";


            if (Num_N_SGs_High_Utilization >= 10)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[5, i].Interior.Color = Color.Red;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + 5, i].Interior.Color = Color.Red;
                }
            }
            if (Num_N_SGs_High_Utilization >= 1 && Num_N_SGs_High_Utilization <= 9)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[5, i].Interior.Color = Color.Gray;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + 5, i].Interior.Color = Color.Gray;
                }
            }
            if (Num_N_SGs_High_Utilization == 0)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[5, i].Interior.Color = Color.LightBlue;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + 5, i].Interior.Color = Color.LightBlue;
                }
            }


            sheet3.Cells[18, 6] = Convert.ToInt64(Num_N_SGs_High_Utilization);
            checkBox4.Invoke(new Action(() => checkBox4.Checked = true));



            //************ M3UA Load Huawei ****************

            // find Maximum of M3UA in a week
            string MSS_M3UA_Huawei_Utilization_Quary_String = @"select  [NE Name], max([Receive Load (%)] + [SEND Load (%)]) as [M3UA_Utilization]	FROM H_M3UA_MSS
            where CAST([Start Time] AS DATE)>='" + start_date_str + "' and CAST([Start Time] AS DATE)< '" + end_date_str + "'  group by [NE Name]";


            SqlCommand MSS_M3UA_Huawei_Utilization_Quary = new SqlCommand(MSS_M3UA_Huawei_Utilization_Quary_String, connection);
            MSS_M3UA_Huawei_Utilization_Quary.CommandTimeout = 0;
            MSS_M3UA_Huawei_Utilization_Quary.ExecuteNonQuery();
            DataTable MSS_M3UA_Huawei_Utilization_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_MSS_M3UA_Huawei_Utilization = new SqlDataAdapter(MSS_M3UA_Huawei_Utilization_Quary);
            dataAdapter_Contractual_MSS_M3UA_Huawei_Utilization.Fill(MSS_M3UA_Huawei_Utilization_Table);




            int Num_H_M3UA_High_Utilization = 0;

            for (int k = 0; k < MSS_M3UA_Huawei_Utilization_Table.Rows.Count; k++)
            {
                if (MSS_M3UA_Huawei_Utilization_Table.Rows[k].ItemArray[1].ToString() != "")
                {
                    if (MSS_M3UA_Huawei_Utilization_Table.Rows[k].ItemArray[1].ToString() != "")
                    {
                        double H_M3UA_Utilization = Convert.ToDouble(MSS_M3UA_Huawei_Utilization_Table.Rows[k].ItemArray[1]);
                        if (H_M3UA_Utilization >= 40)
                        {
                            //sheet3.Cells[25 + Num_N_SGs_High_Utilization + Num_H_SGs_High_Utilization, 5] = MSS_SGs_Nokia_Utilization_Table.Rows[k].ItemArray[0];
                            //sheet3.Cells[25 + Num_N_SGs_High_Utilization + Num_H_SGs_High_Utilization, 6] = Math.Round(Convert.ToDouble(MSS_SGs_Nokia_Utilization_Table.Rows[k].ItemArray[2]), 2);
                            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + 7, 6] = MSS_M3UA_Huawei_Utilization_Table.Rows[k].ItemArray[0];
                            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + 7, 8] = Math.Round(Convert.ToDouble(MSS_M3UA_Huawei_Utilization_Table.Rows[k].ItemArray[1]), 2);
                            Num_H_M3UA_High_Utilization++;
                        }
                    }

                }

            }

            sheet2.Cells[6, 1] = "MSS";
            sheet2.Cells[6, 2] = "HUAWEI";
            sheet2.Cells[6, 3] = "M3UA Load";
            sheet2.Cells[6, 4] = 40;
            sheet2.Cells[6, 5] = Num_H_M3UA_High_Utilization;
            sheet2.Cells[6, 6] = "In M3UA Signaling Load Part, " + Convert.ToString(Num_H_M3UA_High_Utilization) + " of HUAWEi MSSs/TSSs are over Threshhold";

            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + 6, 1] = "MSS";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + 6, 2] = "HUAWEI";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + 6, 3] = "M3UA Load";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + 6, 4] = 40;
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + 6, 5] = Num_H_M3UA_High_Utilization;
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + 6, 9] = "In M3UA Signaling Load Part, " + Convert.ToString(Num_H_M3UA_High_Utilization) + " of HUAWEi MSSs/TSSs are over Threshhold";


            if (Num_H_M3UA_High_Utilization >= 10)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[6, i].Interior.Color = Color.Red;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + 6, i].Interior.Color = Color.Red;
                }
            }
            if (Num_H_M3UA_High_Utilization >= 1 && Num_H_M3UA_High_Utilization <= 9)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[6, i].Interior.Color = Color.Gray;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + 6, i].Interior.Color = Color.Gray;
                }
            }
            if (Num_H_M3UA_High_Utilization == 0)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[6, i].Interior.Color = Color.LightBlue;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + 6, i].Interior.Color = Color.LightBlue;
                }
            }


            checkBox5.Invoke(new Action(() => checkBox5.Checked = true));




            //************ Signaling Load Huawei ****************


            string MGW_Signaling_Huawei_Utilization_Quary_String = @"select[NE Name], max(([Octets Received Occupied Rate(%) (%)] + [Octets Sent Occupied Rate(%) (%)]) / 2) as [Signaling Utilization]  FROM H_MTP3_MGW
            where CAST([Start Time] AS DATE)>='" + start_date_str + "' and CAST([Start Time] AS DATE)<'" + end_date_str + "' group by[NE Name]";


            SqlCommand MGW_Signaling_Huawei_Utilization_Quary = new SqlCommand(MGW_Signaling_Huawei_Utilization_Quary_String, connection);
            MGW_Signaling_Huawei_Utilization_Quary.CommandTimeout = 0;
            MGW_Signaling_Huawei_Utilization_Quary.ExecuteNonQuery();
            DataTable MGW_Signaling_Huawei_Utilization_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_MGW_Signaling_Huawei_Utilization = new SqlDataAdapter(MGW_Signaling_Huawei_Utilization_Quary);
            dataAdapter_Contractual_MGW_Signaling_Huawei_Utilization.Fill(MGW_Signaling_Huawei_Utilization_Table);




            int Num_H_Signaling_High_Utilization = 0;

            for (int k = 0; k < MGW_Signaling_Huawei_Utilization_Table.Rows.Count; k++)
            {
                if (MGW_Signaling_Huawei_Utilization_Table.Rows[k].ItemArray[1].ToString() != "")
                {
                    if (MGW_Signaling_Huawei_Utilization_Table.Rows[k].ItemArray[1].ToString() != "")
                    {
                        double H_Signaling_Utilization = Convert.ToDouble(MGW_Signaling_Huawei_Utilization_Table.Rows[k].ItemArray[1]);
                        if (H_Signaling_Utilization >= 40)
                        {
                            //sheet3.Cells[25 + Num_N_SGs_High_Utilization + Num_H_SGs_High_Utilization, 5] = MSS_SGs_Nokia_Utilization_Table.Rows[k].ItemArray[0];
                            //sheet3.Cells[25 + Num_N_SGs_High_Utilization + Num_H_SGs_High_Utilization, 6] = Math.Round(Convert.ToDouble(MSS_SGs_Nokia_Utilization_Table.Rows[k].ItemArray[2]), 2);
                            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + 8, 6] = MGW_Signaling_Huawei_Utilization_Table.Rows[k].ItemArray[0];
                            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + 8, 8] = Math.Round(Convert.ToDouble(MGW_Signaling_Huawei_Utilization_Table.Rows[k].ItemArray[1]), 2);
                            Num_H_Signaling_High_Utilization++;
                        }
                    }

                }
            }

            sheet2.Cells[7, 1] = "MGW";
            sheet2.Cells[7, 2] = "HUAWEI";
            sheet2.Cells[7, 3] = "Signaling Load";
            sheet2.Cells[7, 4] = 40;
            sheet2.Cells[7, 5] = Num_H_Signaling_High_Utilization;
            sheet2.Cells[7, 6] = "In Signaling Load Part, " + Convert.ToString(Num_H_Signaling_High_Utilization) + " of HUAWEi MGWs are over Threshhold";

            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + 7, 1] = "MGW";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + 7, 2] = "HUAWEI";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + 7, 3] = "Signaling Load";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + 7, 4] = 40;
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + 7, 5] = Num_H_Signaling_High_Utilization;
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + 7, 9] = "In Signaling Load Part, " + Convert.ToString(Num_H_Signaling_High_Utilization) + " of HUAWEi MGWs are over Threshhold";


            if (Num_H_Signaling_High_Utilization >= 10)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[7, i].Interior.Color = Color.Red;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + 7, i].Interior.Color = Color.Red;
                }
            }
            if (Num_H_Signaling_High_Utilization >= 1 && Num_H_Signaling_High_Utilization <= 9)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[7, i].Interior.Color = Color.Gray;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + 7, i].Interior.Color = Color.Gray;
                }
            }
            if (Num_H_Signaling_High_Utilization == 0)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[7, i].Interior.Color = Color.LightBlue;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + 7, i].Interior.Color = Color.LightBlue;
                }
            }


            checkBox6.Invoke(new Action(() => checkBox6.Checked = true));







            //************ AInt Huawei ****************
            string MGW_AInt_Huawei_Utilization_Quary_String = @"select[NE Name], max(100 * ([Number of Busy Circuits (number)] / NULLIF([Number of Idle Circuits (number)] +[Number of Busy Circuits (number)], 0))) as 'AInt Utilization' from[H_Aint_MGW]
             where CAST([Start Time] AS DATE)>='" + start_date_str + "' and CAST([Start Time] AS DATE)<'" + end_date_str + "' group by[NE Name]";


            SqlCommand MGW_AInt_Huawei_Utilization_Quary = new SqlCommand(MGW_AInt_Huawei_Utilization_Quary_String, connection);
            MGW_AInt_Huawei_Utilization_Quary.CommandTimeout = 0;
            MGW_AInt_Huawei_Utilization_Quary.ExecuteNonQuery();
            DataTable MGW_AInt_Huawei_Utilization_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_MGW_AInt_Huawei_Utilization = new SqlDataAdapter(MGW_AInt_Huawei_Utilization_Quary);
            dataAdapter_Contractual_MGW_AInt_Huawei_Utilization.Fill(MGW_AInt_Huawei_Utilization_Table);




            int Num_H_AInt_High_Utilization = 0;

            for (int k = 0; k < MGW_AInt_Huawei_Utilization_Table.Rows.Count; k++)
            {
                if (MGW_AInt_Huawei_Utilization_Table.Rows[k].ItemArray[1].ToString() != "")
                {
                    if (MGW_AInt_Huawei_Utilization_Table.Rows[k].ItemArray[1].ToString() != "")
                    {
                        double H_AInt_Utilization = Convert.ToDouble(MGW_AInt_Huawei_Utilization_Table.Rows[k].ItemArray[1]);
                        if (H_AInt_Utilization >= 80)
                        {
                            //sheet3.Cells[25 + Num_N_SGs_High_Utilization + Num_H_SGs_High_Utilization, 5] = MSS_SGs_Nokia_Utilization_Table.Rows[k].ItemArray[0];
                            //sheet3.Cells[25 + Num_N_SGs_High_Utilization + Num_H_SGs_High_Utilization, 6] = Math.Round(Convert.ToDouble(MSS_SGs_Nokia_Utilization_Table.Rows[k].ItemArray[2]), 2);
                            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + 9, 6] = MGW_AInt_Huawei_Utilization_Table.Rows[k].ItemArray[0];
                            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + 9, 8] = Math.Round(Convert.ToDouble(MGW_AInt_Huawei_Utilization_Table.Rows[k].ItemArray[1]), 2);
                            Num_H_AInt_High_Utilization++;
                        }
                    }

                }

            }

            sheet2.Cells[8, 1] = "MGW";
            sheet2.Cells[8, 2] = "HUAWEI";
            sheet2.Cells[8, 3] = "A int. CGR Utilization";
            sheet2.Cells[8, 4] = 80;
            sheet2.Cells[8, 5] = Num_H_AInt_High_Utilization;
            sheet2.Cells[8, 6] = Convert.ToString(Num_H_AInt_High_Utilization) + " of A interface Trunk Groups Utilization in HUAWEI MGWs are over Threshhold";


            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + 8, 1] = "MGW";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + 8, 2] = "HUAWEI";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + 8, 3] = "A int. CGR Utilization";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + 8, 4] = 80;
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + 8, 5] = Num_H_AInt_High_Utilization;
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + 8, 9] = Convert.ToString(Num_H_AInt_High_Utilization) + " of A interface Trunk Groups Utilization in HUAWEI MGWs are over Threshhold";


            if (Num_H_AInt_High_Utilization >= 10)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[8, i].Interior.Color = Color.Red;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + 8, i].Interior.Color = Color.Red;
                }
            }
            if (Num_H_AInt_High_Utilization >= 1 && Num_H_AInt_High_Utilization <= 9)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[8, i].Interior.Color = Color.Gray;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + 8, i].Interior.Color = Color.Gray;
                }
            }
            if (Num_H_AInt_High_Utilization == 0)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[8, i].Interior.Color = Color.LightBlue;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + 8, i].Interior.Color = Color.LightBlue;
                }
            }


            checkBox7.Invoke(new Action(() => checkBox7.Checked = true));





            //************ MGW Utilization Huawei ****************

            string MGW_Huawei_Utilization_Quary_String = @"select  [NE Name], max([Peak Licensed Traffic (number)]) as 'MGW License UtilizatiON' from [H_LICENSE_MGW]
            WHERE CAST([Start Time] AS DATE)>='" + start_date_str + "' and CAST([Start Time] AS DATE)<'" + end_date_str + "'  group by [NE Name]";


            SqlCommand MGW_Huawei_Utilization_Quary = new SqlCommand(MGW_Huawei_Utilization_Quary_String, connection);
            MGW_Huawei_Utilization_Quary.CommandTimeout = 0;
            MGW_Huawei_Utilization_Quary.ExecuteNonQuery();
            DataTable MGW_Huawei_Utilization_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_MGW_Huawei_Utilization = new SqlDataAdapter(MGW_Huawei_Utilization_Quary);
            dataAdapter_Contractual_MGW_Huawei_Utilization.Fill(MGW_Huawei_Utilization_Table);




            string H_Authorized_Traffic_License_Quary_String = "select * from [H_Authorized_Traffic_License]";

            SqlCommand H_Authorized_Traffic_License_Quary = new SqlCommand(H_Authorized_Traffic_License_Quary_String, connection);
            H_Authorized_Traffic_License_Quary.CommandTimeout = 0;
            H_Authorized_Traffic_License_Quary.ExecuteNonQuery();
            DataTable H_Authorized_Traffic_License_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_H_Authorized_Traffic_License = new SqlDataAdapter(H_Authorized_Traffic_License_Quary);
            dataAdapter_Contractual_H_Authorized_Traffic_License.Fill(H_Authorized_Traffic_License_Table);




            // Join with H_Authorized_Traffic_License
            var H_MGW_Q = (from pd in MGW_Huawei_Utilization_Table.AsEnumerable()
                           join od in H_Authorized_Traffic_License_Table.AsEnumerable() on new { f1 = pd.Field<string>("NE Name") } equals new { f1 = od.Field<string>("MGW Name") } into od
                           from new_od in od.DefaultIfEmpty()
                           select new
                           {
                               NE = pd.Field<string>("NE Name"),
                               MGW_License = pd.Field<double>("MGW License Utilization"),
                               Authorized_Traffic_License = new_od.Field<double>("Authorized Traffic License"),
                               Utilization = 100 * pd.Field<double>("MGW License UtilizatiON") / new_od.Field<double>("Authorized Traffic License"),

                           }).ToList();

            DataTable H_MGW_Table = new DataTable();
            H_MGW_Table = ConvertToDataTable(H_MGW_Q);




            int Num_H_MGW_High_Utilization = 0;

            for (int k = 0; k < H_MGW_Table.Rows.Count; k++)
            {
                if (H_MGW_Table.Rows[k].ItemArray[3].ToString() != "")
                {
                    double H_MGW_Utilization = Convert.ToDouble(H_MGW_Table.Rows[k].ItemArray[3]);
                    if (H_MGW_Utilization >= 80)
                    {
                        //sheet3.Cells[25 + Num_N_SGs_High_Utilization + Num_H_SGs_High_Utilization, 5] = MSS_SGs_Nokia_Utilization_Table.Rows[k].ItemArray[0];
                        //sheet3.Cells[25 + Num_N_SGs_High_Utilization + Num_H_SGs_High_Utilization, 6] = Math.Round(Convert.ToDouble(MSS_SGs_Nokia_Utilization_Table.Rows[k].ItemArray[2]), 2);
                        sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + 10, 6] = H_MGW_Table.Rows[k].ItemArray[0];
                        sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + 10, 8] = Math.Round(Convert.ToDouble(H_MGW_Table.Rows[k].ItemArray[3]), 2);
                        Num_H_MGW_High_Utilization++;
                    }
                }

            }

            sheet2.Cells[9, 1] = "MGW";
            sheet2.Cells[9, 2] = "HUAWEI";
            sheet2.Cells[9, 3] = "MGW LICENSE";
            sheet2.Cells[9, 4] = 80;
            sheet2.Cells[9, 5] = Num_H_MGW_High_Utilization;
            sheet2.Cells[9, 6] = Convert.ToString(Num_H_MGW_High_Utilization) + " of HUAWEI MGWs have Traffic License Utilization over Threshhold";


            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + 9, 1] = "MGW";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + 9, 2] = "HUAWEI";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + 9, 3] = "MGW LICENSE";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + 9, 4] = 80;
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + 9, 5] = Num_H_MGW_High_Utilization;
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + 9, 9] = Convert.ToString(Num_H_MGW_High_Utilization) + " of HUAWEI MGWs have Traffic License Utilization over Threshhold";


            if (Num_H_MGW_High_Utilization >= 10)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[9, i].Interior.Color = Color.Red;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + 9, i].Interior.Color = Color.Red;
                }
            }
            if (Num_H_MGW_High_Utilization >= 1 && Num_H_MGW_High_Utilization <= 9)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[9, i].Interior.Color = Color.Gray;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + 9, i].Interior.Color = Color.Gray;
                }
            }
            if (Num_H_MGW_High_Utilization == 0)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[9, i].Interior.Color = Color.LightBlue;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + 9, i].Interior.Color = Color.LightBlue;
                }
            }


            checkBox8.Invoke(new Action(() => checkBox8.Checked = true));





            //************ BICC Utilization Huawei ****************

            // load H_BICC_INCOMING_MSS
            string H_BICC_INCOMING_MSS_Quary1_String = "Delete from [H_BICC_INCOMING_MSS] where ([BICC Utilization (%)]='NIL' or [Seizure Traffic (Erl)] is null) and CAST([Start Time] AS DATE)>='" + start_date + "' and CAST([Start Time] AS DATE)<'" + end_date + "'";
            SqlCommand H_BICC_INCOMING_MSS_Quary1 = new SqlCommand(H_BICC_INCOMING_MSS_Quary1_String, connection);
            H_BICC_INCOMING_MSS_Quary1.CommandTimeout = 0;
            H_BICC_INCOMING_MSS_Quary1.ExecuteNonQuery();
            string H_BICC_INCOMING_MSS_Quary_String = "select [Start Time], [NE Name], [BICC Office], cast([BICC Utilization (%)] as float) as 'BICC Utilization (%)' , [Number of Available Circuits (piece)],  [Seizure Traffic (Erl)] from [H_BICC_INCOMING_MSS] where CAST([Start Time] AS DATE)>='" + start_date + "' and CAST([Start Time] AS DATE)<'" + end_date + "'";
            SqlCommand H_BICC_INCOMING_MSS_Quary = new SqlCommand(H_BICC_INCOMING_MSS_Quary_String, connection);
            H_BICC_INCOMING_MSS_Quary.CommandTimeout = 0;
            H_BICC_INCOMING_MSS_Quary.ExecuteNonQuery();
            H_BICC_INCOMING_MSS_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_H_BICC_INCOMING_MSS = new SqlDataAdapter(H_BICC_INCOMING_MSS_Quary);
            dataAdapter_Contractual_H_BICC_INCOMING_MSS.Fill(H_BICC_INCOMING_MSS_Table);


            // load H_BICC_OUTGOING_MSS
            string H_BICC_OUTGOING_MSS_Quary_String = "select [Start Time], [NE Name], [BICC Office],  [Number of Available Circuits (piece)],  [Seizure Traffic (Erl)] from [H_BICC_OUTGOING_MSS] where CAST([Start Time] AS DATE)>='" + start_date + "' and CAST([Start Time] AS DATE)<'" + end_date + "'";
            SqlCommand H_BICC_OUTGOING_MSS_Quary = new SqlCommand(H_BICC_OUTGOING_MSS_Quary_String, connection);
            H_BICC_OUTGOING_MSS_Quary.CommandTimeout = 0;
            H_BICC_OUTGOING_MSS_Quary.ExecuteNonQuery();
            H_BICC_OUTGOING_MSS_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_H_BICC_OUTGOING_MSS = new SqlDataAdapter(H_BICC_OUTGOING_MSS_Quary);
            dataAdapter_Contractual_H_BICC_OUTGOING_MSS.Fill(H_BICC_OUTGOING_MSS_Table);



            // Join Using Linq (The First Step: Left Join with Null Values)
            var H_BICC_Q1 = (from pd in H_BICC_INCOMING_MSS_Table.AsEnumerable()
                             join od in H_BICC_OUTGOING_MSS_Table.AsEnumerable() on new { f1 = pd.Field<DateTime>("Start Time"), f2 = pd.Field<string>("NE Name"), f3 = pd.Field<string>("BICC Office") } equals new { f1 = od.Field<DateTime>("Start Time"), f2 = od.Field<string>("NE Name"), f3 = od.Field<string>("BICC Office") } into od
                             from new_od in od.DefaultIfEmpty()
                             select new
                             {
                                 Date = pd.Field<DateTime>("Start Time"),
                                 NE = pd.Field<string>("NE Name"),
                                 BICC_Office = pd.Field<string>("BICC Office"),
                                 //Utilization = pd.Field<string>("BICC Utilization (%)"),
                                 Utilization = (pd != null ? pd.Field<Double>("BICC Utilization (%)") : 0),
                                 TOTAL_Available_Circuits = (pd != null ? pd.Field<Double>("Number of Available Circuits (piece)") : 0),
                                 Used_Traffic = (pd != null ? pd.Field<Double>("Seizure Traffic (Erl)") : 0) + (new_od != null ? new_od.Field<Double>("Seizure Traffic (Erl)") : 0),
                             }).ToList();

            DataTable H_BICC_Table1 = new DataTable();
            H_BICC_Table1 = ConvertToDataTable(H_BICC_Q1);

            DataTable H_BICC_Table2 = new DataTable();
            H_BICC_Table2.Columns.Add("NE", typeof(string));
            H_BICC_Table2.Columns.Add("Utilization", typeof(double));
            for (int r1 = 0; r1 < H_BICC_Table1.Rows.Count; r1++)
            {
                string uti = H_BICC_Table1.Rows[r1][3].ToString();
                if (uti == "" || uti == "NIL" || uti == "NULL")
                    continue;
                H_BICC_Table2.Rows.Add(H_BICC_Table1.Rows[r1].ItemArray[1], Convert.ToDouble(uti));
            }



            // Group by Using Linq
            var H_BICC_Q2 = from row in H_BICC_Table2.AsEnumerable()
                            group row by new { f1 = row.Field<string>("NE") } into rows
                            select new
                            {
                                NE_Name = rows.Key.f1,
                                Max_Utlization = rows.Max(x => x["Utilization"])
                            };

            DataTable H_BICC_Table3 = new DataTable();
            H_BICC_Table3 = ConvertToDataTable(H_BICC_Q2);





            int Num_H_BICC_High_Utilization = 0;

            for (int k = 0; k < H_BICC_Table3.Rows.Count; k++)
            {
                if (H_BICC_Table3.Rows[k].ItemArray[1].ToString() != "")
                {
                    double H_BICC_Utilization = Convert.ToDouble(H_BICC_Table3.Rows[k].ItemArray[1]);
                    if (H_BICC_Utilization >= 80)
                    {
                        //sheet3.Cells[25 + Num_N_SGs_High_Utilization + Num_H_SGs_High_Utilization, 5] = MSS_SGs_Nokia_Utilization_Table.Rows[k].ItemArray[0];
                        //sheet3.Cells[25 + Num_N_SGs_High_Utilization + Num_H_SGs_High_Utilization, 6] = Math.Round(Convert.ToDouble(MSS_SGs_Nokia_Utilization_Table.Rows[k].ItemArray[2]), 2);
                        sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + 11, 6] = H_BICC_Table3.Rows[k].ItemArray[0];
                        sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + 11, 8] = Math.Round(Convert.ToDouble(H_BICC_Table3.Rows[k].ItemArray[1]), 2);
                        Num_H_BICC_High_Utilization++;
                    }
                }

            }

            sheet2.Cells[10, 1] = "MSS";
            sheet2.Cells[10, 2] = "HUAWEI";
            sheet2.Cells[10, 3] = "Trunk Group Utilization (BICC)";
            sheet2.Cells[10, 4] = 80;
            sheet2.Cells[10, 5] = Num_H_BICC_High_Utilization;
            sheet2.Cells[10, 6] = "HUAWEI MSSs have " + Convert.ToString(Num_H_BICC_High_Utilization) + " BICC Utilization over Threshhold";



            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + 10, 1] = "MSS";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + 10, 2] = "HUAWEI";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + 10, 3] = "Trunk Group Utilization (BICC)";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + 10, 4] = 80;
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + 10, 5] = Num_H_BICC_High_Utilization;
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + 10, 9] = "HUAWEI MSSs have " + Convert.ToString(Num_H_BICC_High_Utilization) + " BICC Utilization over Threshhold";


            if (Num_H_BICC_High_Utilization >= 10)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[10, i].Interior.Color = Color.Red;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + 10, i].Interior.Color = Color.Red;
                }
            }
            if (Num_H_BICC_High_Utilization >= 1 && Num_H_BICC_High_Utilization <= 9)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[10, i].Interior.Color = Color.Gray;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + 10, i].Interior.Color = Color.Gray;
                }
            }
            if (Num_H_BICC_High_Utilization == 0)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[10, i].Interior.Color = Color.LightBlue;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + 10, i].Interior.Color = Color.LightBlue;
                }
            }


            checkBox9.Invoke(new Action(() => checkBox9.Checked = true));



            //************ Signaling Load Nokia****************

            string MSS_Signaling_Nokia_Utilization_Quary_String = @"SELECT[MSC Name], max(([M3UA_SCTP_OCTETS_RECEIVED (M661B4C3)] + (56 *[M3UA_SCTP_PACKETS_RECEIVED (M661B4C1)]) +[M3UA_SCTP_OCTETS_SENT (M661B4C4)] + (56 *[M3UA_SCTP_PACKETS_SENT (M661B4C2)])) * 8 / (3600 * 1000)) as [Peak LOAD(KB / Sec)]
            from N_Signaling_MSS WHERE CAST([PERIOD_START_TIME] AS DATE)>= '" + start_date_str + "' and CAST([PERIOD_START_TIME] AS DATE)< '" + end_date_str + "'  group by[MSC Name]";

            SqlCommand MSS_Signaling_Nokia_Utilization_Quary = new SqlCommand(MSS_Signaling_Nokia_Utilization_Quary_String, connection);
            MSS_Signaling_Nokia_Utilization_Quary.CommandTimeout = 0;
            MSS_Signaling_Nokia_Utilization_Quary.ExecuteNonQuery();
            DataTable MSS_Signaling_Nokia_Utilization_Table = new DataTable();
            SqlDataAdapter dataAdapter_Contractual_MSS_Signaling_Nokia_Utilization = new SqlDataAdapter(MSS_Signaling_Nokia_Utilization_Quary);
            dataAdapter_Contractual_MSS_Signaling_Nokia_Utilization.Fill(MSS_Signaling_Nokia_Utilization_Table);




            int Num_N_Signaling_High_Utilization = 0;

            for (int k = 0; k < MSS_Signaling_Nokia_Utilization_Table.Rows.Count; k++)
            {
                if (MSS_Signaling_Nokia_Utilization_Table.Rows[k].ItemArray[1].ToString() != "")
                {
                    double N_Signaling_Utilization = Convert.ToDouble(MSS_Signaling_Nokia_Utilization_Table.Rows[k].ItemArray[1]);
                    if (N_Signaling_Utilization >= 1500)
                    {
                        //sheet3.Cells[25 + Num_N_SGs_High_Utilization + Num_H_SGs_High_Utilization, 5] = MSS_SGs_Nokia_Utilization_Table.Rows[k].ItemArray[0];
                        //sheet3.Cells[25 + Num_N_SGs_High_Utilization + Num_H_SGs_High_Utilization, 6] = Math.Round(Convert.ToDouble(MSS_SGs_Nokia_Utilization_Table.Rows[k].ItemArray[2]), 2);
                        sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + 12, 6] = MSS_Signaling_Nokia_Utilization_Table.Rows[k].ItemArray[0];
                        sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + 12, 8] = Math.Round(Convert.ToDouble(MSS_Signaling_Nokia_Utilization_Table.Rows[k].ItemArray[1]), 2);
                        Num_N_Signaling_High_Utilization++;
                    }
                }

            }

            sheet2.Cells[11, 1] = "MSS";
            sheet2.Cells[11, 2] = "NOKIA";
            sheet2.Cells[11, 3] = "Signaling Load";
            sheet2.Cells[11, 4] = 40;
            sheet2.Cells[11, 5] = Num_N_Signaling_High_Utilization;
            sheet2.Cells[11, 6] = Convert.ToString(Num_N_Signaling_High_Utilization) + " of Signaling Load in ALL NOKIA MSSs / TSSs are over Threshhold ";





            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + 11, 1] = "MSS";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + 11, 2] = "NOKIA";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + 11, 3] = "Signaling Load";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + 11, 4] = 40;
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + 11, 5] = Num_N_Signaling_High_Utilization;
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + 11, 9] = Convert.ToString(Num_N_Signaling_High_Utilization) + " of Signaling Load in ALL NOKIA MSSs / TSSs are over Threshhold ";


            if (Num_N_Signaling_High_Utilization >= 10)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[11, i].Interior.Color = Color.Red;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + 11, i].Interior.Color = Color.Red;
                }
            }
            if (Num_N_Signaling_High_Utilization >= 1 && Num_N_Signaling_High_Utilization <= 9)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[11, i].Interior.Color = Color.Gray;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + 11, i].Interior.Color = Color.Gray;
                }
            }
            if (Num_N_Signaling_High_Utilization == 0)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[11, i].Interior.Color = Color.LightBlue;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + 11, i].Interior.Color = Color.LightBlue;
                }
            }


            checkBox10.Invoke(new Action(() => checkBox10.Checked = true));



            //************ CGR MGW Nokia****************


            string MGW_CGR_Nokia_N_Min_Free_Circuits_Quary_String = @"select[N_Min_Free_Circuits].PERIOD_START_TIME as 'PERIOD_START_TIME1',
                   [N_Min_Free_Circuits].[MGW name] as 'MGW_name1',
            	   [N_Min_Free_Circuits].CGNU_ID as 'CGNU_ID1',
            	   [N_Min_Free_Circuits].CGNA_ID as 'CGNA_ID1',
            	   [N_Min_Free_Circuits].[CGRCGROUP_MIN_FREE (M534C4)] as 'C4' from[N_Min_Free_Circuits]
                   WHERE CAST([PERIOD_START_TIME] AS DATE)>= '" + start_date_str + "' and CAST([PERIOD_START_TIME] AS DATE)< '" + end_date_str + "'";


            SqlCommand MGW_CGR_Nokia_N_Min_Free_Circuits_Quary = new SqlCommand(MGW_CGR_Nokia_N_Min_Free_Circuits_Quary_String, connection);
            MGW_CGR_Nokia_N_Min_Free_Circuits_Quary.CommandTimeout = 0;
            MGW_CGR_Nokia_N_Min_Free_Circuits_Quary.ExecuteNonQuery();
            DataTable MGW_CGR_Nokia_N_Min_Free_Circuits_Table = new DataTable();
            SqlDataAdapter dataAdapter_MGW_CGR_Nokia_N_Min_Free_Circuits = new SqlDataAdapter(MGW_CGR_Nokia_N_Min_Free_Circuits_Quary);
            dataAdapter_MGW_CGR_Nokia_N_Min_Free_Circuits.Fill(MGW_CGR_Nokia_N_Min_Free_Circuits_Table);




            string MGW_N_CGR_Statistics_Quary_String = @"select [N_CGR_Statistics].PERIOD_START_TIME,
                   [N_CGR_Statistics].[MGW name], 
            	   [N_CGR_Statistics].CGNU_ID, 
            	   [N_CGR_Statistics].CGNA_ID, 
            	   [N_CGR_Statistics].[CGRCGROUP_NOF_WOEX_CRTS_OUT (M534C1)] as 'C1' from [N_CGR_Statistics]
                   WHERE CAST([PERIOD_START_TIME] AS DATE)>= '" + start_date_str + "' and CAST([PERIOD_START_TIME] AS DATE)< '" + end_date_str + "'";


            SqlCommand MGW_N_CGR_Statistics_Quary = new SqlCommand(MGW_N_CGR_Statistics_Quary_String, connection);
            MGW_N_CGR_Statistics_Quary.CommandTimeout = 0;
            MGW_N_CGR_Statistics_Quary.ExecuteNonQuery();
            DataTable MGW_N_CGR_Statistics_Table = new DataTable();
            SqlDataAdapter dataAdapter_MGW_N_CGR_Statistics = new SqlDataAdapter(MGW_N_CGR_Statistics_Quary);
            dataAdapter_MGW_N_CGR_Statistics.Fill(MGW_N_CGR_Statistics_Table);



            // Join Using Linq (The First Step: Left Join with Null Values)
            var N_CGR_Q1 = (from pd in MGW_CGR_Nokia_N_Min_Free_Circuits_Table.AsEnumerable()
                            join od in MGW_N_CGR_Statistics_Table.AsEnumerable() on new { f1 = pd.Field<DateTime>("PERIOD_START_TIME1"), f2 = pd.Field<string>("MGW_name1"), f3 = pd.Field<double>("CGNU_ID1"), f4 = pd.Field<string>("CGNA_ID1") } equals new { f1 = od.Field<DateTime>("PERIOD_START_TIME"), f2 = od.Field<string>("MGW name"), f3 = od.Field<double>("CGNU_ID"), f4 = od.Field<string>("CGNA_ID") } into od
                            from new_od in od.DefaultIfEmpty()
                            select new
                            {
                                Date = new_od.Field<DateTime>("PERIOD_START_TIME"),
                                NE = new_od.Field<string>("MGW name"),
                                C1 = (new_od != null ? new_od.Field<Double>("C1") : -1),
                                C4 = (pd != null ? pd.Field<Double>("C4") : -1),
                                Utilization = 100 * ((new_od != null ? new_od.Field<Double>("C1") : -1) / 100 - (pd != null ? pd.Field<Double>("C4") : -1)) / ((new_od != null ? new_od.Field<Double>("C1") : -1) / 100),
                            }).ToList();

            DataTable N_CGR_Table1 = new DataTable();
            N_CGR_Table1 = ConvertToDataTable(N_CGR_Q1);



            // Group by Using Linq
            var N_CGR_Q2 = from row in N_CGR_Table1.AsEnumerable()
                           group row by new { f1 = row.Field<string>("NE") } into rows
                           select new
                           {
                               NE_Name = rows.Key.f1,
                               Max_Utlization = rows.Max(x => x["Utilization"])
                           };

            DataTable N_CGR_Table2 = new DataTable();
            N_CGR_Table2 = ConvertToDataTable(N_CGR_Q2);





            int Num_N_CGR_High_Utilization = 0;

            for (int k = 0; k < N_CGR_Table2.Rows.Count; k++)
            {
                if (N_CGR_Table2.Rows[k].ItemArray[1].ToString() != "")
                {
                    double N_CGR_Utilization = Convert.ToDouble(N_CGR_Table2.Rows[k].ItemArray[1]);
                    if (N_CGR_Utilization >= 80)
                    {
                        //sheet3.Cells[25 + Num_N_SGs_High_Utilization + Num_H_SGs_High_Utilization, 5] = MSS_SGs_Nokia_Utilization_Table.Rows[k].ItemArray[0];
                        //sheet3.Cells[25 + Num_N_SGs_High_Utilization + Num_H_SGs_High_Utilization, 6] = Math.Round(Convert.ToDouble(MSS_SGs_Nokia_Utilization_Table.Rows[k].ItemArray[2]), 2);
                        sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + 13, 6] = N_CGR_Table2.Rows[k].ItemArray[0];
                        sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + 13, 8] = Math.Round(Convert.ToDouble(N_CGR_Table2.Rows[k].ItemArray[1]), 2);
                        Num_N_CGR_High_Utilization++;
                    }
                }

            }

            sheet2.Cells[12, 1] = "MGW";
            sheet2.Cells[12, 2] = "NOKIA";
            sheet2.Cells[12, 3] = "CGR Utilization";
            sheet2.Cells[12, 4] = 80;
            sheet2.Cells[12, 5] = Num_N_CGR_High_Utilization;
            sheet2.Cells[12, 6] = Convert.ToString(Num_N_CGR_High_Utilization) + " of All NOKIA MGWs have CGR Utilization over Threshhold";





            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + 12, 1] = "MGW";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + 12, 2] = "NOKIA";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + 12, 3] = "CGR Utilization";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + 12, 4] = 80;
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + 12, 5] = Num_N_CGR_High_Utilization;
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + 12, 9] = Convert.ToString(Num_N_CGR_High_Utilization) + " of All NOKIA MGWs have CGR Utilization over Threshhold";


            if (Num_N_CGR_High_Utilization >= 10)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[12, i].Interior.Color = Color.Red;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + 12, i].Interior.Color = Color.Red;
                }
            }
            if (Num_N_CGR_High_Utilization >= 1 && Num_N_CGR_High_Utilization <= 9)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[12, i].Interior.Color = Color.Gray;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + 12, i].Interior.Color = Color.Gray;
                }
            }
            if (Num_N_CGR_High_Utilization == 0)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[12, i].Interior.Color = Color.LightBlue;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + 12, i].Interior.Color = Color.LightBlue;
                }
            }


            checkBox11.Invoke(new Action(() => checkBox11.Checked = true));





            //************ MGW Connection Capacity Nokia****************

            string MGW_Nokia_Connection_Capacity_Quary_String = @"select Date, NE, [capacity_licence_utilization_Peak] from[Nokia_MGW]
            WHERE CAST(Date AS DATE)>= '" + start_date_str + "' and CAST(Date AS DATE)< '" + end_date_str + "' and substring(NE,1,2)= 'MG' and[capacity_licence_utilization_Peak] >= 80";


            SqlCommand MGW_Nokia_Connection_Capacity_Quary = new SqlCommand(MGW_Nokia_Connection_Capacity_Quary_String, connection);
            MGW_Nokia_Connection_Capacity_Quary.CommandTimeout = 0;
            MGW_Nokia_Connection_Capacity_Quary.ExecuteNonQuery();
            DataTable MGW_Nokia_Connection_Capacity_Table = new DataTable();
            SqlDataAdapter dataAdapter_MGW_Nokia_Connection_Capacity = new SqlDataAdapter(MGW_Nokia_Connection_Capacity_Quary);
            dataAdapter_MGW_Nokia_Connection_Capacity.Fill(MGW_Nokia_Connection_Capacity_Table);





            // Group by Using Linq
            var N_CC_Q = from row in MGW_Nokia_Connection_Capacity_Table.AsEnumerable()
                         group row by new { f1 = row.Field<string>("NE") } into rows
                         select new
                         {
                             NE_Name = rows.Key.f1,
                             Max_Utlization = rows.Max(x => x["capacity_licence_utilization_Peak"])
                         };

            DataTable N_CC_Table = new DataTable();
            N_CC_Table = ConvertToDataTable(N_CC_Q);



            int Num_N_CC_High_Utilization = 0;

            for (int k = 0; k < N_CC_Table.Rows.Count; k++)
            {
                if (N_CC_Table.Rows[k].ItemArray[1].ToString() != "")
                {
                    double N_CC_Utilization = Convert.ToDouble(N_CC_Table.Rows[k].ItemArray[1]);
                    if (N_CC_Utilization >= 80)
                    {
                        sheet3.Cells[25 + Num_N_CC_High_Utilization, 13] = N_CC_Table.Rows[k].ItemArray[0];
                        sheet3.Cells[25 + Num_N_CC_High_Utilization, 14] = Math.Round(Convert.ToDouble(N_CC_Table.Rows[k].ItemArray[1]), 2);
                        sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + 14, 6] = N_CC_Table.Rows[k].ItemArray[0];
                        sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + 14, 8] = Math.Round(Convert.ToDouble(N_CC_Table.Rows[k].ItemArray[1]), 2);
                        Num_N_CC_High_Utilization++;
                    }
                }

            }

            sheet2.Cells[13, 1] = "MGW";
            sheet2.Cells[13, 2] = "NOKIA";
            sheet2.Cells[13, 3] = "Connection Capacity";
            sheet2.Cells[13, 4] = 80;
            sheet2.Cells[13, 5] = Num_N_CC_High_Utilization;
            sheet2.Cells[13, 6] = Convert.ToString(Num_N_CC_High_Utilization) + " of ALL NOKIA MGWs have Connection Capacity License Utilization over Threshhold";





            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + 13, 1] = "MGW";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + 13, 2] = "NOKIA";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + 13, 3] = "Connection Capacity";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + 13, 4] = 80;
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + 13, 5] = Num_N_CC_High_Utilization;
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + 13, 9] = Convert.ToString(Num_N_CC_High_Utilization) + " of ALL NOKIA MGWs have Connection Capacity License Utilization over Threshhold";


            if (Num_N_CC_High_Utilization >= 10)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[13, i].Interior.Color = Color.Red;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + 13, i].Interior.Color = Color.Red;
                }
            }
            if (Num_N_CC_High_Utilization >= 1 && Num_N_CC_High_Utilization <= 9)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[13, i].Interior.Color = Color.Gray;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + 13, i].Interior.Color = Color.Gray;
                }
            }
            if (Num_N_CC_High_Utilization == 0)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[13, i].Interior.Color = Color.LightBlue;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + 13, i].Interior.Color = Color.LightBlue;
                }
            }


            sheet3.Cells[18, 14] = Convert.ToInt64(Num_N_CC_High_Utilization);


            checkBox12.Invoke(new Action(() => checkBox12.Checked = true));




            //************ MGW IUIP Nokia****************
            string MGW_Nokia_IUIP_Quary_String = @"select Date, NE, [IU_IP_Peak] from [Nokia_MGW]
            WHERE CAST(Date AS DATE)>= '" + start_date_str + "' and CAST(Date AS DATE)< '" + end_date_str + "' and substring(NE,1,2)= 'MG' and [IU_IP_Peak] >= 80";


            SqlCommand MGW_Nokia_IUIP_Quary = new SqlCommand(MGW_Nokia_IUIP_Quary_String, connection);
            MGW_Nokia_IUIP_Quary.CommandTimeout = 0;
            MGW_Nokia_IUIP_Quary.ExecuteNonQuery();
            DataTable MGW_Nokia_IUIP_Table = new DataTable();
            SqlDataAdapter dataAdapter_MGW_Nokia_IUIP = new SqlDataAdapter(MGW_Nokia_IUIP_Quary);
            dataAdapter_MGW_Nokia_IUIP.Fill(MGW_Nokia_IUIP_Table);




            // Group by Using Linq
            var N_IUIP_Q = from row in MGW_Nokia_IUIP_Table.AsEnumerable()
                           group row by new { f1 = row.Field<string>("NE") } into rows
                           select new
                           {
                               NE_Name = rows.Key.f1,
                               Max_Utlization = rows.Max(x => x["IU_IP_Peak"])
                           };

            DataTable N_IUIP_Table = new DataTable();
            N_IUIP_Table = ConvertToDataTable(N_IUIP_Q);






            int Num_N_IUIP_High_Utilization = 0;

            for (int k = 0; k < N_IUIP_Table.Rows.Count; k++)
            {
                if (N_IUIP_Table.Rows[k].ItemArray[1].ToString() != "")
                {
                    double N_IUIP_Utilization = Convert.ToDouble(N_IUIP_Table.Rows[k].ItemArray[1]);
                    if (N_IUIP_Utilization >= 80)
                    {
                        sheet3.Cells[25 + Num_N_IUIP_High_Utilization, 15] = N_IUIP_Table.Rows[k].ItemArray[0];
                        sheet3.Cells[25 + Num_N_IUIP_High_Utilization, 16] = Math.Round(Convert.ToDouble(N_IUIP_Table.Rows[k].ItemArray[1]), 2);
                        sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + 15, 6] = N_IUIP_Table.Rows[k].ItemArray[0];
                        sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + 15, 8] = Math.Round(Convert.ToDouble(N_IUIP_Table.Rows[k].ItemArray[1]), 2);
                        Num_N_IUIP_High_Utilization++;
                    }
                }

            }

            sheet2.Cells[14, 1] = "MGW";
            sheet2.Cells[14, 2] = "NOKIA";
            sheet2.Cells[14, 3] = "IUCS LICENSE";
            sheet2.Cells[14, 4] = 80;
            sheet2.Cells[14, 5] = Num_N_IUIP_High_Utilization;
            sheet2.Cells[14, 6] = Convert.ToString(Num_N_IUIP_High_Utilization) + " of ALL NOKIA MGWs have IUCS License Utilization over Threshhold";





            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + 14, 1] = "MGW";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + 14, 2] = "NOKIA";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + 14, 3] = "IUCS LICENSE";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + 14, 4] = 80;
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + 14, 5] = Num_N_IUIP_High_Utilization;
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + 14, 9] = Convert.ToString(Num_N_IUIP_High_Utilization) + " of ALL NOKIA MGWs have IUCS License Utilization over Threshhold";


            if (Num_N_IUIP_High_Utilization >= 10)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[14, i].Interior.Color = Color.Red;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + 14, i].Interior.Color = Color.Red;
                }
            }
            if (Num_N_IUIP_High_Utilization >= 1 && Num_N_IUIP_High_Utilization <= 9)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[14, i].Interior.Color = Color.Gray;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + 14, i].Interior.Color = Color.Gray;
                }
            }
            if (Num_N_IUIP_High_Utilization == 0)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[14, i].Interior.Color = Color.LightBlue;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + 14, i].Interior.Color = Color.LightBlue;
                }
            }

            sheet3.Cells[18, 16] = Convert.ToInt64(Num_N_IUIP_High_Utilization);
            checkBox13.Invoke(new Action(() => checkBox13.Checked = true));







            //************ MGW NBIP Nokia****************
            string MGW_Nokia_NBIP_Quary_String = @"select Date, NE, [NB_IP_Peak_License] from [Nokia_MGW]
            WHERE CAST(Date AS DATE)>= '" + start_date_str + "' and CAST(Date AS DATE)< '" + end_date_str + "' and substring(NE,1,2)= 'MG' and [NB_IP_Peak_License] >= 80";


            SqlCommand MGW_Nokia_NBIP_Quary = new SqlCommand(MGW_Nokia_NBIP_Quary_String, connection);
            MGW_Nokia_NBIP_Quary.CommandTimeout = 0;
            MGW_Nokia_NBIP_Quary.ExecuteNonQuery();
            DataTable MGW_Nokia_NBIP_Table = new DataTable();
            SqlDataAdapter dataAdapter_MGW_Nokia_NBIP = new SqlDataAdapter(MGW_Nokia_NBIP_Quary);
            dataAdapter_MGW_Nokia_NBIP.Fill(MGW_Nokia_NBIP_Table);



            // Group by Using Linq
            var N_NBIP_Q = from row in MGW_Nokia_NBIP_Table.AsEnumerable()
                           group row by new { f1 = row.Field<string>("NE") } into rows
                           select new
                           {
                               NE_Name = rows.Key.f1,
                               Max_Utlization = rows.Max(x => x["NB_IP_Peak_License"])
                           };

            DataTable N_NBIP_Table = new DataTable();
            N_NBIP_Table = ConvertToDataTable(N_NBIP_Q);






            int Num_N_NBIP_High_Utilization = 0;

            for (int k = 0; k < N_NBIP_Table.Rows.Count; k++)
            {
                if (N_NBIP_Table.Rows[k].ItemArray[1].ToString() != "")
                {
                    double N_NBIP_Utilization = Convert.ToDouble(N_NBIP_Table.Rows[k].ItemArray[1]);
                    if (N_NBIP_Utilization >= 80)
                    {
                        sheet3.Cells[25 + Num_N_NBIP_High_Utilization, 17] = N_NBIP_Table.Rows[k].ItemArray[0];
                        sheet3.Cells[25 + Num_N_NBIP_High_Utilization, 18] = Math.Round(Convert.ToDouble(N_NBIP_Table.Rows[k].ItemArray[1]), 2);
                        sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + 16, 6] = N_NBIP_Table.Rows[k].ItemArray[0];
                        sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + 16, 8] = Math.Round(Convert.ToDouble(N_NBIP_Table.Rows[k].ItemArray[1]), 2);
                        Num_N_NBIP_High_Utilization++;
                    }
                }

            }

            sheet2.Cells[15, 1] = "MGW";
            sheet2.Cells[15, 2] = "NOKIA";
            sheet2.Cells[15, 3] = "NBIP LICENSE";
            sheet2.Cells[15, 4] = 80;
            sheet2.Cells[15, 5] = Num_N_NBIP_High_Utilization;
            sheet2.Cells[15, 6] = Convert.ToString(Num_N_NBIP_High_Utilization) + " of ALL NOKIA MGWs have NBIP License Utilization over Threshhold";





            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + 15, 1] = "MGW";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + 15, 2] = "NOKIA";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + 15, 3] = "NBIP LICENSE";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + 15, 4] = 80;
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + 15, 5] = Num_N_NBIP_High_Utilization;
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + 15, 9] = Convert.ToString(Num_N_NBIP_High_Utilization) + " of ALL NOKIA MGWs have NBIP License Utilization over Threshhold";


            if (Num_N_NBIP_High_Utilization >= 10)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[15, i].Interior.Color = Color.Red;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + 15, i].Interior.Color = Color.Red;
                }
            }
            if (Num_N_NBIP_High_Utilization >= 1 && Num_N_NBIP_High_Utilization <= 9)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[15, i].Interior.Color = Color.Gray;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + 15, i].Interior.Color = Color.Gray;
                }
            }
            if (Num_N_NBIP_High_Utilization == 0)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[15, i].Interior.Color = Color.LightBlue;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + 15, i].Interior.Color = Color.LightBlue;
                }
            }

            sheet3.Cells[18, 18] = Convert.ToInt64(Num_N_NBIP_High_Utilization);
            checkBox14.Invoke(new Action(() => checkBox14.Checked = true));








            //************ MGW AOIP Nokia****************
            string MGW_Nokia_AOIP_Quary_String = @"select Date, NE, [AOIP_Peak] from [Nokia_MGW]
            WHERE CAST(Date AS DATE)>= '" + start_date_str + "' and CAST(Date AS DATE)< '" + end_date_str + "' and substring(NE,1,2)= 'MG' and [AOIP_Peak] >= 80";


            SqlCommand MGW_Nokia_AOIP_Quary = new SqlCommand(MGW_Nokia_AOIP_Quary_String, connection);
            MGW_Nokia_AOIP_Quary.CommandTimeout = 0;
            MGW_Nokia_AOIP_Quary.ExecuteNonQuery();
            DataTable MGW_Nokia_AOIP_Table = new DataTable();
            SqlDataAdapter dataAdapter_MGW_Nokia_AOIP = new SqlDataAdapter(MGW_Nokia_AOIP_Quary);
            dataAdapter_MGW_Nokia_AOIP.Fill(MGW_Nokia_AOIP_Table);




            // Group by Using Linq
            var N_AOIP_Q = from row in MGW_Nokia_AOIP_Table.AsEnumerable()
                           group row by new { f1 = row.Field<string>("NE") } into rows
                           select new
                           {
                               NE_Name = rows.Key.f1,
                               Max_Utlization = rows.Max(x => x["AOIP_Peak"])
                           };

            DataTable N_AOIP_Table = new DataTable();
            N_AOIP_Table = ConvertToDataTable(N_AOIP_Q);






            int Num_N_AOIP_High_Utilization = 0;

            for (int k = 0; k < N_AOIP_Table.Rows.Count; k++)
            {
                if (N_AOIP_Table.Rows[k].ItemArray[1].ToString() != "")
                {
                    double N_AOIP_Utilization = Convert.ToDouble(N_AOIP_Table.Rows[k].ItemArray[1]);
                    if (N_AOIP_Utilization >= 80)
                    {
                        sheet3.Cells[25 + Num_N_AOIP_High_Utilization, 19] = N_AOIP_Table.Rows[k].ItemArray[0];
                        sheet3.Cells[25 + Num_N_AOIP_High_Utilization, 20] = Math.Round(Convert.ToDouble(N_AOIP_Table.Rows[k].ItemArray[1]), 2);
                        sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + 17, 6] = N_AOIP_Table.Rows[k].ItemArray[0];
                        sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + 17, 8] = Math.Round(Convert.ToDouble(N_AOIP_Table.Rows[k].ItemArray[1]), 2);
                        Num_N_AOIP_High_Utilization++;
                    }
                }

            }

            sheet2.Cells[16, 1] = "MGW";
            sheet2.Cells[16, 2] = "NOKIA";
            sheet2.Cells[16, 3] = "AOIP LICENSE";
            sheet2.Cells[16, 4] = 80;
            sheet2.Cells[16, 5] = Num_N_AOIP_High_Utilization;
            sheet2.Cells[16, 6] = Convert.ToString(Num_N_AOIP_High_Utilization) + " of ALL NOKIA MGWs have AOIP License Utilization over Threshhold";





            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + 16, 1] = "MGW";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + 16, 2] = "NOKIA";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + 16, 3] = "AOIP LICENSE";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + 16, 4] = 80;
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + 16, 5] = Num_N_AOIP_High_Utilization;
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + 16, 9] = Convert.ToString(Num_N_AOIP_High_Utilization) + " of ALL NOKIA MGWs have AOIP License Utilization over Threshhold";


            if (Num_N_AOIP_High_Utilization >= 10)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[16, i].Interior.Color = Color.Red;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + 16, i].Interior.Color = Color.Red;
                }
            }
            if (Num_N_AOIP_High_Utilization >= 1 && Num_N_AOIP_High_Utilization <= 9)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[16, i].Interior.Color = Color.Gray;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + 16, i].Interior.Color = Color.Gray;
                }
            }
            if (Num_N_AOIP_High_Utilization == 0)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[16, i].Interior.Color = Color.LightBlue;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + 16, i].Interior.Color = Color.LightBlue;
                }
            }
            sheet3.Cells[18, 20] = Convert.ToInt64(Num_N_AOIP_High_Utilization);

            checkBox15.Invoke(new Action(() => checkBox15.Checked = true));




            //************ MGW MB Nokia****************
            string MGW_Nokia_MB_Quary_String = @"select Date, NE, [MB_Peak_license] from [Nokia_MGW]
            WHERE CAST(Date AS DATE)>= '" + start_date_str + "' and CAST(Date AS DATE)< '" + end_date_str + "' and substring(NE,1,2)= 'MG' and [MB_Peak_license] >= 80";


            SqlCommand MGW_Nokia_MB_Quary = new SqlCommand(MGW_Nokia_MB_Quary_String, connection);
            MGW_Nokia_MB_Quary.CommandTimeout = 0;
            MGW_Nokia_MB_Quary.ExecuteNonQuery();
            DataTable MGW_Nokia_MB_Table = new DataTable();
            SqlDataAdapter dataAdapter_MGW_Nokia_MB = new SqlDataAdapter(MGW_Nokia_MB_Quary);
            dataAdapter_MGW_Nokia_MB.Fill(MGW_Nokia_MB_Table);




            // Group by Using Linq
            var N_MB_Q = from row in MGW_Nokia_MB_Table.AsEnumerable()
                         group row by new { f1 = row.Field<string>("NE") } into rows
                         select new
                         {
                             NE_Name = rows.Key.f1,
                             Max_Utlization = rows.Max(x => x["MB_Peak_license"])
                         };

            DataTable N_MB_Table = new DataTable();
            N_MB_Table = ConvertToDataTable(N_MB_Q);






            int Num_N_MB_High_Utilization = 0;

            for (int k = 0; k < N_MB_Table.Rows.Count; k++)
            {
                if (N_MB_Table.Rows[k].ItemArray[1].ToString() != "")
                {
                    double N_MB_Utilization = Convert.ToDouble(N_MB_Table.Rows[k].ItemArray[1]);
                    if (N_MB_Utilization >= 80)
                    {
                        sheet3.Cells[25 + Num_N_MB_High_Utilization, 21] = N_MB_Table.Rows[k].ItemArray[0];
                        sheet3.Cells[25 + Num_N_MB_High_Utilization, 22] = Math.Round(Convert.ToDouble(N_MB_Table.Rows[k].ItemArray[1]), 2);
                        sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + 18, 6] = N_MB_Table.Rows[k].ItemArray[0];
                        sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + 18, 8] = Math.Round(Convert.ToDouble(N_MB_Table.Rows[k].ItemArray[1]), 2);
                        Num_N_MB_High_Utilization++;
                    }
                }

            }

            sheet2.Cells[17, 1] = "MGW";
            sheet2.Cells[17, 2] = "NOKIA";
            sheet2.Cells[17, 3] = "MB LICENSE";
            sheet2.Cells[17, 4] = 80;
            sheet2.Cells[17, 5] = Num_N_MB_High_Utilization;
            sheet2.Cells[17, 6] = Convert.ToString(Num_N_MB_High_Utilization) + " of ALL NOKIA MGWs have MB License Utilization over Threshhold";





            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + 17, 1] = "MGW";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + 17, 2] = "NOKIA";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + 17, 3] = "MB LICENSE";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + 17, 4] = 80;
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + 17, 5] = Num_N_MB_High_Utilization;
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + 17, 9] = Convert.ToString(Num_N_MB_High_Utilization) + " of ALL NOKIA MGWs have MB License Utilization over Threshhold";


            if (Num_N_MB_High_Utilization >= 10)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[17, i].Interior.Color = Color.Red;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + 17, i].Interior.Color = Color.Red;
                }
            }
            if (Num_N_MB_High_Utilization >= 1 && Num_N_MB_High_Utilization <= 9)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[17, i].Interior.Color = Color.Gray;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + 17, i].Interior.Color = Color.Gray;
                }
            }
            if (Num_N_MB_High_Utilization == 0)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[17, i].Interior.Color = Color.LightBlue;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + 17, i].Interior.Color = Color.LightBlue;
                }
            }

            sheet3.Cells[18, 22] = Convert.ToInt64(Num_N_MB_High_Utilization);
            checkBox16.Invoke(new Action(() => checkBox16.Checked = true));







            //************ MGW Ater Nokia****************
            string MGW_Nokia_Ater_Quary_String = @"select Date, NE, [Ater_Peak_license] from [Nokia_MGW]
            WHERE CAST(Date AS DATE)>= '" + start_date_str + "' and CAST(Date AS DATE)< '" + end_date_str + "' and substring(NE,1,2)= 'MG' and [Ater_Peak_license] >= 80";


            SqlCommand MGW_Nokia_Ater_Quary = new SqlCommand(MGW_Nokia_Ater_Quary_String, connection);
            MGW_Nokia_Ater_Quary.CommandTimeout = 0;
            MGW_Nokia_Ater_Quary.ExecuteNonQuery();
            DataTable MGW_Nokia_Ater_Table = new DataTable();
            SqlDataAdapter dataAdapter_MGW_Nokia_Ater = new SqlDataAdapter(MGW_Nokia_Ater_Quary);
            dataAdapter_MGW_Nokia_Ater.Fill(MGW_Nokia_Ater_Table);




            // Group by Using Linq
            var N_Ater_Q = from row in MGW_Nokia_Ater_Table.AsEnumerable()
                           group row by new { f1 = row.Field<string>("NE") } into rows
                           select new
                           {
                               NE_Name = rows.Key.f1,
                               Max_Utlization = rows.Max(x => x["Ater_Peak_license"])
                           };

            DataTable N_Ater_Table = new DataTable();
            N_Ater_Table = ConvertToDataTable(N_Ater_Q);






            int Num_N_Ater_High_Utilization = 0;

            for (int k = 0; k < N_Ater_Table.Rows.Count; k++)
            {
                if (N_Ater_Table.Rows[k].ItemArray[1].ToString() != "")
                {
                    double N_Ater_Utilization = Convert.ToDouble(N_Ater_Table.Rows[k].ItemArray[1]);
                    if (N_Ater_Utilization >= 80)
                    {
                        sheet3.Cells[25 + Num_N_Ater_High_Utilization, 23] = N_Ater_Table.Rows[k].ItemArray[0];
                        sheet3.Cells[25 + Num_N_Ater_High_Utilization, 24] = Math.Round(Convert.ToDouble(N_Ater_Table.Rows[k].ItemArray[1]), 2);
                        sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + Num_N_Ater_High_Utilization + 19, 6] = N_Ater_Table.Rows[k].ItemArray[0];
                        sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + Num_N_Ater_High_Utilization + 19, 8] = Math.Round(Convert.ToDouble(N_Ater_Table.Rows[k].ItemArray[1]), 2);
                        Num_N_Ater_High_Utilization++;
                    }
                }

            }

            sheet2.Cells[18, 1] = "MGW";
            sheet2.Cells[18, 2] = "NOKIA";
            sheet2.Cells[18, 3] = "Ater LICENSE";
            sheet2.Cells[18, 4] = 80;
            sheet2.Cells[18, 5] = Num_N_Ater_High_Utilization;
            sheet2.Cells[18, 6] = Convert.ToString(Num_N_Ater_High_Utilization) + " of ALL NOKIA MGWs have Ater License Utilization over Threshhold";





            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + 18, 1] = "MGW";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + 18, 2] = "NOKIA";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + 18, 3] = "Ater LICENSE";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + 18, 4] = 80;
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + 18, 5] = Num_N_Ater_High_Utilization;
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + 18, 9] = Convert.ToString(Num_N_Ater_High_Utilization) + " of ALL NOKIA MGWs have Ater License Utilization over Threshhold";


            if (Num_N_Ater_High_Utilization >= 10)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[18, i].Interior.Color = Color.Red;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + 18, i].Interior.Color = Color.Red;
                }
            }
            if (Num_N_Ater_High_Utilization >= 1 && Num_N_Ater_High_Utilization <= 9)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[18, i].Interior.Color = Color.Gray;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + 18, i].Interior.Color = Color.Gray;
                }
            }
            if (Num_N_Ater_High_Utilization == 0)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[18, i].Interior.Color = Color.LightBlue;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + 18, i].Interior.Color = Color.LightBlue;
                }
            }

            sheet3.Cells[18, 24] = Convert.ToInt64(Num_N_Ater_High_Utilization);
            checkBox17.Invoke(new Action(() => checkBox17.Checked = true));





            //************ MGW CAMEL-PH4 Nokia****************
            string MGW_Nokia_CAMEL_Quary_String = @"select PERIOD_START_TIME, [MSC Name], [FEAC_ID], [AVERAGE_CAP_LIC_USAGE_X100 (M406B2C2)] / 100 as 'VLR CAMEL PH4 223' from [N_VLR_Features]
            WHERE CAST(PERIOD_START_TIME AS DATE)>= '" + start_date_str + "' and CAST(PERIOD_START_TIME AS DATE)< '" + end_date_str + "' and [AVERAGE_CAP_LIC_USAGE_X100 (M406B2C2)] / 100 >= 80 and [FEAC_ID]= '223'";



            SqlCommand MGW_Nokia_CAMEL_Quary = new SqlCommand(MGW_Nokia_CAMEL_Quary_String, connection);
            MGW_Nokia_CAMEL_Quary.CommandTimeout = 0;
            MGW_Nokia_CAMEL_Quary.ExecuteNonQuery();
            DataTable MGW_Nokia_CAMEL_Table = new DataTable();
            SqlDataAdapter dataAdapter_MGW_Nokia_CAMEL = new SqlDataAdapter(MGW_Nokia_CAMEL_Quary);
            dataAdapter_MGW_Nokia_CAMEL.Fill(MGW_Nokia_CAMEL_Table);




            // Group by Using Linq
            var N_CAMEL_Q = from row in MGW_Nokia_CAMEL_Table.AsEnumerable()
                            group row by new { f1 = row.Field<string>("MSC Name") } into rows
                            select new
                            {
                                NE_Name = rows.Key.f1,
                                Max_Utlization = rows.Max(x => x["VLR CAMEL PH4 223"])
                            };

            DataTable N_CAMEL_Table = new DataTable();
            N_CAMEL_Table = ConvertToDataTable(N_CAMEL_Q);




            int Num_N_CAMEL_High_Utilization = 0;

            for (int k = 0; k < N_CAMEL_Table.Rows.Count; k++)
            {
                if (N_CAMEL_Table.Rows[k].ItemArray[1].ToString() != "")
                {
                    double N_CAMEL_Utilization = Convert.ToDouble(N_CAMEL_Table.Rows[k].ItemArray[1]);
                    if (N_CAMEL_Utilization >= 80)
                    {
                        sheet3.Cells[25 + Num_N_CAMEL_High_Utilization, 7] = N_CAMEL_Table.Rows[k].ItemArray[0];
                        sheet3.Cells[25 + Num_N_CAMEL_High_Utilization, 8] = Math.Round(Convert.ToDouble(N_CAMEL_Table.Rows[k].ItemArray[1]), 2);
                        sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + Num_N_Ater_High_Utilization + Num_N_CAMEL_High_Utilization + 20, 6] = N_CAMEL_Table.Rows[k].ItemArray[0];
                        sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + Num_N_Ater_High_Utilization + Num_N_CAMEL_High_Utilization + 20, 8] = Math.Round(Convert.ToDouble(N_CAMEL_Table.Rows[k].ItemArray[1]), 2);
                        Num_N_CAMEL_High_Utilization++;
                    }
                }


            }

            sheet2.Cells[19, 1] = "MSS";
            sheet2.Cells[19, 2] = "NOKIA";
            sheet2.Cells[19, 3] = "CAMEL PH4 LICENSE";
            sheet2.Cells[19, 4] = 80;
            sheet2.Cells[19, 5] = Num_N_CAMEL_High_Utilization;
            sheet2.Cells[19, 6] = Convert.ToString(Num_N_CAMEL_High_Utilization) + " of ALL NOKIA MGWs have CAMEL PH4 License Utilization over Threshhold";





            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + Num_N_Ater_High_Utilization + 19, 1] = "MSS";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + Num_N_Ater_High_Utilization + 19, 2] = "NOKIA";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + Num_N_Ater_High_Utilization + 19, 3] = "CAMEL PH4 LICENSE";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + Num_N_Ater_High_Utilization + 19, 4] = 80;
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + Num_N_Ater_High_Utilization + 19, 5] = Num_N_CAMEL_High_Utilization;
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + Num_N_Ater_High_Utilization + 19, 9] = Convert.ToString(Num_N_CAMEL_High_Utilization) + " of ALL NOKIA MGWs have CAMEL PH4 License Utilization over Threshhold";


            if (Num_N_CAMEL_High_Utilization >= 10)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[19, i].Interior.Color = Color.Red;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + Num_N_Ater_High_Utilization + 19, i].Interior.Color = Color.Red;
                }
            }
            if (Num_N_CAMEL_High_Utilization >= 1 && Num_N_CAMEL_High_Utilization <= 9)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[19, i].Interior.Color = Color.Gray;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + Num_N_Ater_High_Utilization + 19, i].Interior.Color = Color.Gray;
                }
            }
            if (Num_N_CAMEL_High_Utilization == 0)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[19, i].Interior.Color = Color.LightBlue;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + Num_N_Ater_High_Utilization + 19, i].Interior.Color = Color.LightBlue;
                }
            }


            sheet3.Cells[18, 8] = Convert.ToInt64(Num_N_CAMEL_High_Utilization);


            checkBox18.Invoke(new Action(() => checkBox18.Checked = true));




            //************ MSS ISUP Huawei****************
            string MSS_Huawei_ISUP_INCOMING_Quary_String = @"select[Start Time], [NE Name], [Trunk group], [Seizure Traffic (Erl)], [Number of Available Circuits (piece)] from
                        H_ISUP_INCOMING_MSS where [Number of Available Circuits (piece)] != 0 and[Number of Available Circuits (piece)] is not null and CAST([Start Time] AS DATE)>= '" + start_date_str + "' and CAST([Start Time] AS DATE)< '" + end_date_str + "'";


            SqlCommand MSS_Huawei_ISUP_INCOMING_Quary = new SqlCommand(MSS_Huawei_ISUP_INCOMING_Quary_String, connection);
            MSS_Huawei_ISUP_INCOMING_Quary.CommandTimeout = 0;
            MSS_Huawei_ISUP_INCOMING_Quary.ExecuteNonQuery();
            DataTable MSS_Huawei_ISUP_INCOMING_Table = new DataTable();
            SqlDataAdapter dataAdapter_MSS_Huawei_ISUP_INCOMING = new SqlDataAdapter(MSS_Huawei_ISUP_INCOMING_Quary);
            dataAdapter_MSS_Huawei_ISUP_INCOMING.Fill(MSS_Huawei_ISUP_INCOMING_Table);


            string MSS_Huawei_ISUP_OUTGOING_Quary_String = @" select[Start Time], [NE Name], [Trunk group], [Seizure Traffic (Erl)], [Number of Available Circuits (piece)] from
                        H_ISUP_OUTGOING_MSS where CAST([Start Time] AS DATE)>= '" + start_date_str + "' and CAST([Start Time] AS DATE)< '" + end_date_str + "'";


            SqlCommand MSS_Huawei_ISUP_OUTGOING_Quary = new SqlCommand(MSS_Huawei_ISUP_OUTGOING_Quary_String, connection);
            MSS_Huawei_ISUP_OUTGOING_Quary.CommandTimeout = 0;
            MSS_Huawei_ISUP_OUTGOING_Quary.ExecuteNonQuery();
            DataTable MSS_Huawei_ISUP_OUTGOING_Table = new DataTable();
            SqlDataAdapter dataAdapter_MSS_Huawei_ISUP_OUTGOING = new SqlDataAdapter(MSS_Huawei_ISUP_OUTGOING_Quary);
            dataAdapter_MSS_Huawei_ISUP_OUTGOING.Fill(MSS_Huawei_ISUP_OUTGOING_Table);



            // Join Using Linq (The First Step: Left Join with Null Values)
            var H_ISUP_Q = (from pd in MSS_Huawei_ISUP_INCOMING_Table.AsEnumerable()
                            join od in MSS_Huawei_ISUP_OUTGOING_Table.AsEnumerable() on new { f1 = pd.Field<DateTime>("Start Time"), f2 = pd.Field<string>("NE Name"), f3 = pd.Field<string>("Trunk group") } equals new { f1 = od.Field<DateTime>("Start Time"), f2 = od.Field<string>("NE Name"), f3 = od.Field<string>("Trunk group") } into od
                            from new_od in od.DefaultIfEmpty()
                            select new
                            {
                                Date = pd.Field<DateTime>("Start Time"),
                                NE = pd.Field<string>("NE Name"),
                                Trunk_Group = pd.Field<string>("Trunk group"),
                                ISUP_Utilization = 100 * ((pd != null ? pd.Field<Double>("Seizure Traffic (Erl)") : 0) + (new_od != null ? new_od.Field<Double>("Seizure Traffic (Erl)") : 0)) / (pd != null ? pd.Field<Double>("Number of Available Circuits (piece)") : -1),

                            }).ToList();

            DataTable H_ISUP_Table = new DataTable();
            H_ISUP_Table = ConvertToDataTable(H_ISUP_Q);



            // Group by Using Linq
            var H_ISUP_Q1 = from row in H_ISUP_Table.AsEnumerable()
                            group row by new { f1 = row.Field<string>("Trunk_Group") } into rows
                            select new
                            {
                                Trunk_Name = rows.Key.f1,
                                Max_Utilization = rows.Max(x => x["ISUP_Utilization"])
                            };

            DataTable H_ISUP_Table1 = new DataTable();
            H_ISUP_Table1 = ConvertToDataTable(H_ISUP_Q1);




            // Join Using Linq (The First Step: Left Join with Null Values)
            var H_ISUP_Q2 = (from pd in H_ISUP_Table1.AsEnumerable()
                             join od in H_ISUP_Table.AsEnumerable() on new { f1 = pd.Field<string>("Trunk_Name"), f2 = pd.Field<double>("Max_Utilization") } equals new { f1 = od.Field<string>("Trunk_Group"), f2 = od.Field<double>("ISUP_Utilization") } into od
                             from new_od in od.DefaultIfEmpty()
                             select new
                             {
                                 NE_Name = new_od.Field<string>("NE"),
                                 Trunk_Group = (pd != null ? pd.Field<string>("Trunk_Name") : ""),
                                 ISUP_Utilization = (pd != null ? pd.Field<Double>("Max_Utilization") : 0),


                             }).ToList();

            DataTable H_ISUP_Table2 = new DataTable();
            H_ISUP_Table2 = ConvertToDataTable(H_ISUP_Q2);



            int Num_H_ISUP_High_Utilization = 0;

            for (int k = 0; k < H_ISUP_Table2.Rows.Count; k++)
            {
                if (H_ISUP_Table2.Rows[k].ItemArray[2].ToString() != "")
                {
                    double H_ISUP_Utilization = Convert.ToDouble(H_ISUP_Table2.Rows[k].ItemArray[2]);
                    if (H_ISUP_Utilization >= 80)
                    {
                        //sheet3.Cells[25 + Num_N_SGs_High_Utilization + Num_H_SGs_High_Utilization, 5] = MSS_SGs_Nokia_Utilization_Table.Rows[k].ItemArray[0];
                        //sheet3.Cells[25 + Num_N_SGs_High_Utilization + Num_H_SGs_High_Utilization, 6] = Math.Round(Convert.ToDouble(MSS_SGs_Nokia_Utilization_Table.Rows[k].ItemArray[2]), 2);
                        string Trunk_Str = Convert.ToString(H_ISUP_Table2.Rows[k].ItemArray[1]);
                        if (Trunk_Str.Length >= 6)
                        {
                            Trunk_Str = Trunk_Str.Substring(6, Trunk_Str.Length - 6);

                            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + Num_N_Ater_High_Utilization + Num_N_CAMEL_High_Utilization + Num_H_ISUP_High_Utilization + 21, 6] = Convert.ToString(H_ISUP_Table2.Rows[k].ItemArray[0]);
                            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + Num_N_Ater_High_Utilization + Num_N_CAMEL_High_Utilization + Num_H_ISUP_High_Utilization + 21, 7] = Trunk_Str;
                            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + Num_N_Ater_High_Utilization + Num_N_CAMEL_High_Utilization + Num_H_ISUP_High_Utilization + 21, 8] = Math.Round(Convert.ToDouble(H_ISUP_Table2.Rows[k].ItemArray[2]), 2);
                            Num_H_ISUP_High_Utilization++;
                        }

                    }
                }

            }

            sheet2.Cells[20, 1] = "MSS";
            sheet2.Cells[20, 2] = "HUAWEI";
            sheet2.Cells[20, 3] = "Trunk Group Utilization (ISUP)";
            sheet2.Cells[20, 4] = 80;
            sheet2.Cells[20, 5] = Num_H_ISUP_High_Utilization;
            sheet2.Cells[20, 6] = "In HUAWEI MSSs / TSSs " + Convert.ToString(Num_H_ISUP_High_Utilization) + " of ISUP Trunk groups have Utilization over Threshhold";





            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + Num_N_Ater_High_Utilization + Num_N_CAMEL_High_Utilization + 20, 1] = "MSS";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + Num_N_Ater_High_Utilization + Num_N_CAMEL_High_Utilization + 20, 2] = "HUAWEI";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + Num_N_Ater_High_Utilization + Num_N_CAMEL_High_Utilization + 20, 3] = "Trunk Group Utilization (ISUP)";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + Num_N_Ater_High_Utilization + Num_N_CAMEL_High_Utilization + 20, 4] = 80;
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + Num_N_Ater_High_Utilization + Num_N_CAMEL_High_Utilization + 20, 5] = Num_H_ISUP_High_Utilization;
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + Num_N_Ater_High_Utilization + Num_N_CAMEL_High_Utilization + 20, 9] = "In HUAWEI MSSs / TSSs " + Convert.ToString(Num_H_ISUP_High_Utilization) + " of ISUP Trunk groups have Utilization over Threshhold";


            if (Num_H_ISUP_High_Utilization >= 10)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[20, i].Interior.Color = Color.Red;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + Num_N_Ater_High_Utilization + Num_N_CAMEL_High_Utilization + 20, i].Interior.Color = Color.Red;
                }
            }
            if (Num_H_ISUP_High_Utilization >= 1 && Num_H_ISUP_High_Utilization <= 9)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[20, i].Interior.Color = Color.Gray;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + Num_N_Ater_High_Utilization + Num_N_CAMEL_High_Utilization + 20, i].Interior.Color = Color.Gray;
                }
            }
            if (Num_H_ISUP_High_Utilization == 0)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[20, i].Interior.Color = Color.LightBlue;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + Num_N_Ater_High_Utilization + Num_N_CAMEL_High_Utilization + 20, i].Interior.Color = Color.LightBlue;
                }
            }


            checkBox19.Invoke(new Action(() => checkBox19.Checked = true));





            //************ MSS ISUP Nokia****************
            string MSS_Nokia_CGR_Quary_String = @"select[NE Name], max(CGR_Statistics) as 'CGR Utilization' from(
            select[PERIOD_START_TIME], [MSC Name] + '_' +[CGR Name] as 'NE Name', ROUND(100 * ([CGRCGROUP_ERLANGS_OUT_x_100 (M16B2C25)] +[CGRCGROUP_ERLANGS_IN_x_100 (M16B2C24)]) / NULLIF([CGRCGROUP_NOF_WOEX_CRTS_OUT (M16B2C6)], 0), 2) as   'CGR_Statistics'
            from N_CGR_Utilization where CAST([PERIOD_START_TIME] AS DATE)>= '" + start_date_str + "' and CAST([PERIOD_START_TIME] AS DATE)<  '" + end_date_str + "') tble group by[NE Name]";


            SqlCommand MSS_Nokia_CGR_Quary = new SqlCommand(MSS_Nokia_CGR_Quary_String, connection);
            MSS_Nokia_CGR_Quary.CommandTimeout = 0;
            MSS_Nokia_CGR_Quary.ExecuteNonQuery();
            DataTable MSS_Nokia_CGR_Quary_Table = new DataTable();
            SqlDataAdapter dataAdapter_MSS_Nokia_CGR_Quary = new SqlDataAdapter(MSS_Nokia_CGR_Quary);
            dataAdapter_MSS_Nokia_CGR_Quary.Fill(MSS_Nokia_CGR_Quary_Table);




            int Num_N_MSS_CGR_High_Utilization = 0;

            for (int k = 0; k < MSS_Nokia_CGR_Quary_Table.Rows.Count; k++)
            {

                if (MSS_Nokia_CGR_Quary_Table.Rows[k].ItemArray[1].ToString() == "")
                {
                    continue;
                }

                double N_MSS_CGR_Utilization = Convert.ToDouble(MSS_Nokia_CGR_Quary_Table.Rows[k].ItemArray[1]);
                if (N_MSS_CGR_Utilization >= 80)
                {
                    //sheet3.Cells[25 + Num_N_SGs_High_Utilization + Num_H_SGs_High_Utilization, 5] = MSS_SGs_Nokia_Utilization_Table.Rows[k].ItemArray[0];
                    //sheet3.Cells[25 + Num_N_SGs_High_Utilization + Num_H_SGs_High_Utilization, 6] = Math.Round(Convert.ToDouble(MSS_SGs_Nokia_Utilization_Table.Rows[k].ItemArray[2]), 2);
                    string Trunk_Str = Convert.ToString(MSS_Nokia_CGR_Quary_Table.Rows[k].ItemArray[0]);
                    if (Trunk_Str.Length >= 5)
                    {
                        Trunk_Str = Trunk_Str.Substring(0, 5);

                        sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + Num_N_Ater_High_Utilization + Num_N_CAMEL_High_Utilization + Num_H_ISUP_High_Utilization + Num_N_MSS_CGR_High_Utilization + 22, 6] = Trunk_Str;
                        sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + Num_N_Ater_High_Utilization + Num_N_CAMEL_High_Utilization + Num_H_ISUP_High_Utilization + Num_N_MSS_CGR_High_Utilization + 22, 7] = Convert.ToString(MSS_Nokia_CGR_Quary_Table.Rows[k].ItemArray[0]);
                        sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + Num_N_Ater_High_Utilization + Num_N_CAMEL_High_Utilization + Num_H_ISUP_High_Utilization + Num_N_MSS_CGR_High_Utilization + 22, 8] = Math.Round(Convert.ToDouble(MSS_Nokia_CGR_Quary_Table.Rows[k].ItemArray[1]), 2);
                        Num_N_MSS_CGR_High_Utilization++;
                    }

                }
            }

            sheet2.Cells[21, 1] = "MSS";
            sheet2.Cells[21, 2] = "NOKIA";
            sheet2.Cells[21, 3] = "CGR Utilization";
            sheet2.Cells[21, 4] = 80;
            sheet2.Cells[21, 5] = Num_N_MSS_CGR_High_Utilization;
            sheet2.Cells[21, 6] = "In NOKIA MSSs / TSSs  " + Convert.ToString(Num_N_MSS_CGR_High_Utilization) + " CGRs have Utilization over Threshhold";






            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + Num_N_Ater_High_Utilization + Num_N_CAMEL_High_Utilization + Num_H_ISUP_High_Utilization + 21, 1] = "MSS";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + Num_N_Ater_High_Utilization + Num_N_CAMEL_High_Utilization + Num_H_ISUP_High_Utilization + 21, 2] = "NOKIA";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + Num_N_Ater_High_Utilization + Num_N_CAMEL_High_Utilization + Num_H_ISUP_High_Utilization + 21, 3] = "CGR Utilization";
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + Num_N_Ater_High_Utilization + Num_N_CAMEL_High_Utilization + Num_H_ISUP_High_Utilization + 21, 4] = 80;
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + Num_N_Ater_High_Utilization + Num_N_CAMEL_High_Utilization + Num_H_ISUP_High_Utilization + 21, 5] = Num_N_MSS_CGR_High_Utilization;
            sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + Num_N_Ater_High_Utilization + Num_N_CAMEL_High_Utilization + Num_H_ISUP_High_Utilization + 21, 9] = "In NOKIA MSSs / TSSs  " + Convert.ToString(Num_N_MSS_CGR_High_Utilization) + " CGRs have Utilization over Threshhold";


            if (Num_N_MSS_CGR_High_Utilization >= 10)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[21, i].Interior.Color = Color.Red;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + Num_N_Ater_High_Utilization + Num_N_CAMEL_High_Utilization + Num_H_ISUP_High_Utilization + 21, i].Interior.Color = Color.Red;
                }
            }
            if (Num_N_MSS_CGR_High_Utilization >= 1 && Num_N_MSS_CGR_High_Utilization <= 9)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[21, i].Interior.Color = Color.Gray;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + Num_N_Ater_High_Utilization + Num_N_CAMEL_High_Utilization + Num_H_ISUP_High_Utilization + 21, i].Interior.Color = Color.Gray;
                }
            }
            if (Num_N_MSS_CGR_High_Utilization == 0)
            {
                for (int i = 1; i <= 9; i++)
                {
                    if (i <= 6)
                    {
                        sheet2.Cells[21, i].Interior.Color = Color.LightBlue;
                    }
                    sheet1.Cells[Num_H_VLR_High_Utilization + Num_N_VLR_High_Utilization + Num_H_SGs_High_Utilization + Num_N_SGs_High_Utilization + Num_H_M3UA_High_Utilization + Num_H_Signaling_High_Utilization + Num_H_AInt_High_Utilization + Num_H_MGW_High_Utilization + Num_H_BICC_High_Utilization + Num_N_Signaling_High_Utilization + Num_N_CGR_High_Utilization + Num_N_CC_High_Utilization + Num_N_IUIP_High_Utilization + Num_N_NBIP_High_Utilization + Num_N_AOIP_High_Utilization + Num_N_MB_High_Utilization + Num_N_Ater_High_Utilization + Num_N_CAMEL_High_Utilization + Num_H_ISUP_High_Utilization + 21, i].Interior.Color = Color.LightBlue;
                }
            }


            checkBox20.Invoke(new Action(() => checkBox20.Checked = true));






            //************ Traffic ****************


            string MSS_Traffic_Quary_String = @"select[dbo].[Traffic_Sub].Date, 'Nokia' as Vendor, [dbo].[Traffic_Sub].NE, tble.[Nokia Traffic] as 'Traffic' from(select [NE] as 'NE1', max([Traffic_Per_Sub(MErl)(Nokia_Core)]) as 'Nokia Traffic'  from[dbo].[Traffic_Sub]
           where[Traffic_Per_Sub(MErl)(Nokia_Core)] is not null  and CAST(Date AS DATE) >= '" + start_date_str + "' and CAST(Date AS DATE) < '" + end_date_str + "'  and [Traffic_Per_Sub(MErl)(Nokia_Core)] != 0 group by[NE]) tble left join [Traffic_Sub] on[Traffic_Sub].[Traffic_Per_Sub(MErl)(Nokia_Core)] = tble.[Nokia Traffic] and [Traffic_Sub].NE=tble.[NE1] and CAST(Date AS DATE) >= '" + start_date_str + "' and CAST(Date AS DATE) < '" + end_date_str + "'" +
      @" union all
select[dbo].[Traffic_Sub].Date, 'Huawei' as Vendor, [dbo].[Traffic_Sub].NE, tble.[Huawei Traffic] as 'Traffic' from(select [NE] as 'NE1', max(Traffic_Per_Sub_mErl_HU) as 'Huawei Traffic'  from[dbo].[Traffic_Sub]
where[Traffic_Per_Sub_mErl_HU] is not null  and CAST(Date AS DATE) >= '" + start_date_str + "' and CAST(Date AS DATE) < '" + end_date_str + "'  and[Traffic_Per_Sub_mErl_HU] != 0 group by[NE]) tble left join [Traffic_Sub] on[Traffic_Sub].Traffic_Per_Sub_mErl_HU = tble.[Huawei Traffic] and [Traffic_Sub].NE=tble.[NE1] and CAST(Date AS DATE) >= '" + start_date_str + "' and CAST(Date AS DATE) < '" + end_date_str + "'";



            SqlCommand MSS_Traffic_Quary = new SqlCommand(MSS_Traffic_Quary_String, connection);
            MSS_Traffic_Quary.CommandTimeout = 0;
            MSS_Traffic_Quary.ExecuteNonQuery();
            DataTable MSS_Traffic_Quary_Table = new DataTable();
            SqlDataAdapter dataAdapter_MSS_Traffic_Quary = new SqlDataAdapter(MSS_Traffic_Quary);
            dataAdapter_MSS_Traffic_Quary.Fill(MSS_Traffic_Quary_Table);


            double Max_Khoozestan = 0;
            double Max_Markazi = 0;
            double Max_Ardebil = 0;
            double Max_Hormozgan = 0;
            double Max_Mazandaran = 0;
            double Max_N_Khorasan = 0;
            double Max_Booshehr = 0;
            double Max_S_Khorasan = 0;
            double Max_Esfahan = 0;
            double Max_Golestan = 0;
            double Max_Hamedan = 0;
            double Max_Ilam = 0;
            double Max_Lorestan = 0;
            double Max_Kerman = 0;
            double Max_Kermanshah = 0;
            double Max_R_Khorasan = 0;
            double Max_Qom = 0;
            double Max_Qazvin = 0;
            double Max_Tehran = 0;
            double Max_Gilan = 0;
            double Max_Semnan = 0;
            double Max_Kordestan = 0;
            double Max_Chahahr_Mahal = 0;
            double Max_Fars = 0;
            double Max_E_Azar = 0;
            double Max_W_Azar = 0;
            double Max_Kohgilooyeh = 0;
            double Max_Yazd = 0;
            double Max_Sistan_Baloch = 0;
            double Max_Zanjan = 0;

            double Traffic_Khoozestan = 0;
            double Traffic_Markazi = 0;
            double Traffic_Ardebil = 0;
            double Traffic_Hormozgan = 0;
            double Traffic_Mazandaran = 0;
            double Traffic_N_Khorasan = 0;
            double Traffic_Booshehr = 0;
            double Traffic_S_Khorasan = 0;
            double Traffic_Esfahan = 0;
            double Traffic_Golestan = 0;
            double Traffic_Hamedan = 0;
            double Traffic_Ilam = 0;
            double Traffic_Lorestan = 0;
            double Traffic_Kerman = 0;
            double Traffic_Kermanshah = 0;
            double Traffic_R_Khorasan = 0;
            double Traffic_Qom = 0;
            double Traffic_Qazvin = 0;
            double Traffic_Tehran = 0;
            double Traffic_Gilan = 0;
            double Traffic_Semnan = 0;
            double Traffic_Kordestan = 0;
            double Traffic_Chahahr_Mahal = 0;
            double Traffic_Fars = 0;
            double Traffic_E_Azar = 0;
            double Traffic_W_Azar = 0;
            double Traffic_Kohgilooyeh = 0;
            double Traffic_Yazd = 0;
            double Traffic_Sistan_Baloch = 0;
            double Traffic_Zanjan = 0;



            for (int p = 0; p < MSS_Traffic_Quary_Table.Rows.Count; p++)
            {
                string NE = MSS_Traffic_Quary_Table.Rows[p].ItemArray[2].ToString();
                //Khoozestan
                if (NE == "MSAHA")
                {
                    Traffic_Khoozestan = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[2, 4] = Traffic_Khoozestan;
                    sheet4.Cells[2, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_Khoozestan > Max_Khoozestan)
                    {
                        Max_Khoozestan = Traffic_Khoozestan;
                        sheet4.Cells[2, 6] = Max_Khoozestan;
                        sheet5.Cells[2, 2] = Max_Khoozestan;
                    }
                }
                if (NE == "MSAHB")
                {
                    Traffic_Khoozestan = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[3, 4] = Traffic_Khoozestan;
                    sheet4.Cells[3, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_Khoozestan > Max_Khoozestan)
                    {
                        Max_Khoozestan = Traffic_Khoozestan;
                        sheet4.Cells[2, 6] = Max_Khoozestan;
                        sheet5.Cells[2, 2] = Max_Khoozestan;
                    }
                }
                if (NE == "MSAHC")
                {
                    Traffic_Khoozestan = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[4, 4] = Traffic_Khoozestan;
                    sheet4.Cells[4, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_Khoozestan > Max_Khoozestan)
                    {
                        Max_Khoozestan = Traffic_Khoozestan;
                        sheet4.Cells[2, 6] = Max_Khoozestan;
                        sheet5.Cells[2, 2] = Max_Khoozestan;
                    }
                }
                //Markazi
                if (NE == "MSAKB")
                {
                    Traffic_Markazi = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[5, 4] = Traffic_Markazi;
                    sheet4.Cells[5, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_Markazi > Max_Markazi)
                    {
                        Max_Markazi = Traffic_Markazi;
                        sheet4.Cells[5, 6] = Max_Markazi;
                        sheet5.Cells[3, 2] = Max_Markazi;
                    }
                }
                //Ardebil
                if (NE == "MSARA")
                {
                    Traffic_Ardebil = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[6, 4] = Traffic_Ardebil;
                    sheet4.Cells[6, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_Ardebil > Max_Ardebil)
                    {
                        Max_Ardebil = Traffic_Ardebil;
                        sheet4.Cells[6, 6] = Max_Ardebil;
                        sheet5.Cells[4, 2] = Max_Ardebil;
                    }
                }
                //Hormozgan
                if (NE == "MSBAA")
                {
                    Traffic_Hormozgan = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[7, 4] = Traffic_Hormozgan;
                    sheet4.Cells[7, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_Hormozgan > Max_Hormozgan)
                    {
                        Max_Hormozgan = Traffic_Hormozgan;
                        sheet4.Cells[7, 6] = Max_Hormozgan;
                        sheet5.Cells[5, 2] = Max_Hormozgan;
                    }
                }
                if (NE == "MSBAB")
                {
                    Traffic_Hormozgan = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[8, 4] = Traffic_Hormozgan;
                    sheet4.Cells[8, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_Hormozgan > Max_Hormozgan)
                    {
                        Max_Hormozgan = Traffic_Hormozgan;
                        sheet4.Cells[7, 6] = Max_Hormozgan;
                        sheet5.Cells[5, 2] = Max_Hormozgan;
                    }
                }
                //Mazandaran
                if (NE == "MSBBA")
                {
                    Traffic_Mazandaran = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[9, 4] = Traffic_Mazandaran;
                    sheet4.Cells[9, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_Mazandaran > Max_Mazandaran)
                    {
                        Max_Mazandaran = Traffic_Mazandaran;
                        sheet4.Cells[9, 6] = Max_Mazandaran;
                        sheet5.Cells[6, 2] = Max_Mazandaran;
                    }
                }
                if (NE == "MSNOA")
                {
                    Traffic_Mazandaran = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[10, 4] = Traffic_Mazandaran;
                    sheet4.Cells[10, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_Mazandaran > Max_Mazandaran)
                    {
                        Max_Mazandaran = Traffic_Mazandaran;
                        sheet4.Cells[9, 6] = Max_Mazandaran;
                        sheet5.Cells[6, 2] = Max_Mazandaran;
                    }
                }
                if (NE == "MSSRA")
                {
                    Traffic_Mazandaran = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[11, 4] = Traffic_Mazandaran;
                    sheet4.Cells[11, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_Mazandaran > Max_Mazandaran)
                    {
                        Max_Mazandaran = Traffic_Mazandaran;
                        sheet4.Cells[9, 6] = Max_Mazandaran;
                        sheet5.Cells[6, 2] = Max_Mazandaran;
                    }
                }
                //N. Khorasan
                if (NE == "MSBJA")
                {
                    Traffic_N_Khorasan = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[12, 4] = Traffic_N_Khorasan;
                    sheet4.Cells[12, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_N_Khorasan > Max_N_Khorasan)
                    {
                        Max_N_Khorasan = Traffic_N_Khorasan;
                        sheet4.Cells[12, 6] = Max_N_Khorasan;
                        sheet5.Cells[7, 2] = Max_N_Khorasan;
                    }
                }
                //Booshehr
                if (NE == "MSBOA")
                {
                    Traffic_Booshehr = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[13, 4] = Traffic_Booshehr;
                    sheet4.Cells[13, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_Booshehr > Max_Booshehr)
                    {
                        Max_Booshehr = Traffic_Booshehr;
                        sheet4.Cells[13, 6] = Max_Booshehr;
                        sheet5.Cells[8, 2] = Max_Booshehr;
                    }
                }
                //S. Khorasan
                if (NE == "MSBRA")
                {
                    Traffic_S_Khorasan = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[14, 4] = Traffic_S_Khorasan;
                    sheet4.Cells[14, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_S_Khorasan > Max_S_Khorasan)
                    {
                        Max_S_Khorasan = Traffic_S_Khorasan;
                        sheet4.Cells[14, 6] = Max_S_Khorasan;
                        sheet5.Cells[9, 2] = Max_S_Khorasan;
                    }
                }
                //Esfahan
                if (NE == "MSEFA")
                {
                    Traffic_Esfahan = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[15, 4] = Traffic_Esfahan;
                    sheet4.Cells[15, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_Esfahan > Max_Esfahan)
                    {
                        Max_Esfahan = Traffic_Esfahan;
                        sheet4.Cells[15, 6] = Max_Esfahan;
                        sheet5.Cells[10, 2] = Max_Esfahan;
                    }
                }
                if (NE == "MSEFB")
                {
                    Traffic_Esfahan = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[16, 4] = Traffic_Esfahan;
                    sheet4.Cells[16, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_Esfahan > Max_Esfahan)
                    {
                        Max_Esfahan = Traffic_Esfahan;
                        sheet4.Cells[15, 6] = Max_Esfahan;
                        sheet5.Cells[10, 2] = Max_Esfahan;
                    }
                }
                if (NE == "MSEFC")
                {
                    Traffic_Esfahan = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[17, 4] = Traffic_Esfahan;
                    sheet4.Cells[17, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_Esfahan > Max_Esfahan)
                    {
                        Max_Esfahan = Traffic_Esfahan;
                        sheet4.Cells[15, 6] = Max_Esfahan;
                        sheet5.Cells[10, 2] = Max_Esfahan;
                    }
                }
                if (NE == "MSEFD")
                {
                    Traffic_Esfahan = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[18, 4] = Traffic_Esfahan;
                    sheet4.Cells[18, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_Esfahan > Max_Esfahan)
                    {
                        Max_Esfahan = Traffic_Esfahan;
                        sheet4.Cells[15, 6] = Max_Esfahan;
                        sheet5.Cells[10, 2] = Max_Esfahan;
                    }
                }
                //Golestan
                if (NE == "MSGOA")
                {
                    Traffic_Golestan = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[19, 4] = Traffic_Golestan;
                    sheet4.Cells[19, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_Golestan > Max_Golestan)
                    {
                        Max_Golestan = Traffic_Golestan;
                        sheet4.Cells[19, 6] = Max_Golestan;
                        sheet5.Cells[11, 2] = Max_Golestan;
                    }
                }
                //Hamedan
                if (NE == "MSHNA")
                {
                    Traffic_Hamedan = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[20, 4] = Traffic_Hamedan;
                    sheet4.Cells[20, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_Hamedan > Max_Hamedan)
                    {
                        Max_Hamedan = Traffic_Hamedan;
                        sheet4.Cells[20, 6] = Max_Hamedan;
                        sheet5.Cells[12, 2] = Max_Hamedan;
                    }
                }
                if (NE == "MSHNB")
                {
                    Traffic_Hamedan = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[21, 4] = Traffic_Hamedan;
                    sheet4.Cells[21, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_Hamedan > Max_Hamedan)
                    {
                        Max_Hamedan = Traffic_Hamedan;
                        sheet4.Cells[20, 6] = Max_Hamedan;
                        sheet5.Cells[12, 2] = Max_Hamedan;
                    }
                }
                //Ilam
                if (NE == "MSILA")
                {
                    Traffic_Ilam = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[22, 4] = Traffic_Ilam;
                    sheet4.Cells[22, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_Ilam > Max_Ilam)
                    {
                        Max_Ilam = Traffic_Ilam;
                        sheet4.Cells[22, 6] = Max_Ilam;
                        sheet5.Cells[13, 2] = Max_Ilam;
                    }
                }
                //Lorestan
                if (NE == "MSKHA")
                {
                    Traffic_Lorestan = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[23, 4] = Traffic_Lorestan;
                    sheet4.Cells[23, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_Lorestan > Max_Lorestan)
                    {
                        Max_Lorestan = Traffic_Lorestan;
                        sheet4.Cells[23, 6] = Max_Lorestan;
                        sheet5.Cells[14, 2] = Max_Lorestan;
                    }
                }
                //Kerman
                if (NE == "MSKRA")
                {
                    Traffic_Kerman = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[24, 4] = Traffic_Kerman;
                    sheet4.Cells[24, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_Kerman > Max_Kerman)
                    {
                        Max_Kerman = Traffic_Kerman;
                        sheet4.Cells[24, 6] = Max_Kerman;
                        sheet5.Cells[15, 2] = Max_Kerman;
                    }
                }
                if (NE == "MSKRB")
                {
                    Traffic_Kerman = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[25, 4] = Traffic_Kerman;
                    sheet4.Cells[25, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_Kerman > Max_Kerman)
                    {
                        Max_Kerman = Traffic_Kerman;
                        sheet4.Cells[24, 6] = Max_Kerman;
                        sheet5.Cells[15, 2] = Max_Kerman;
                    }
                }
                //Kermanshah
                if (NE == "MSKSA")
                {
                    Traffic_Kermanshah = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[26, 4] = Traffic_Kermanshah;
                    sheet4.Cells[26, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_Kermanshah > Max_Kermanshah)
                    {
                        Max_Kermanshah = Traffic_Kermanshah;
                        sheet4.Cells[26, 6] = Max_Kermanshah;
                        sheet5.Cells[16, 2] = Max_Kermanshah;
                    }
                }
                //R_Khorasan
                if (NE == "MSMDA")
                {
                    Traffic_R_Khorasan = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[27, 4] = Traffic_R_Khorasan;
                    sheet4.Cells[27, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_R_Khorasan > Max_R_Khorasan)
                    {
                        Max_R_Khorasan = Traffic_R_Khorasan;
                        sheet4.Cells[27, 6] = Max_R_Khorasan;
                        sheet5.Cells[17, 2] = Max_R_Khorasan;
                    }
                }
                if (NE == "MSMDB")
                {
                    Traffic_R_Khorasan = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[28, 4] = Traffic_R_Khorasan;
                    sheet4.Cells[28, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_R_Khorasan > Max_R_Khorasan)
                    {
                        Max_R_Khorasan = Traffic_R_Khorasan;
                        sheet4.Cells[27, 6] = Max_R_Khorasan;
                        sheet5.Cells[17, 2] = Max_R_Khorasan;
                    }
                }
                if (NE == "MSMDC")
                {
                    Traffic_R_Khorasan = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[29, 4] = Traffic_R_Khorasan;
                    sheet4.Cells[29, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_R_Khorasan > Max_R_Khorasan)
                    {
                        Max_R_Khorasan = Traffic_R_Khorasan;
                        sheet4.Cells[27, 6] = Max_R_Khorasan;
                        sheet5.Cells[17, 2] = Max_R_Khorasan;
                    }
                }
                if (NE == "MSMDD")
                {
                    Traffic_R_Khorasan = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[30, 4] = Traffic_R_Khorasan;
                    sheet4.Cells[30, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_R_Khorasan > Max_R_Khorasan)
                    {
                        Max_R_Khorasan = Traffic_R_Khorasan;
                        sheet4.Cells[27, 6] = Max_R_Khorasan;
                        sheet5.Cells[17, 2] = Max_R_Khorasan;
                    }
                }
                //Qom
                if (NE == "MSQMA")
                {
                    Traffic_Qom = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[31, 4] = Traffic_Qom;
                    sheet4.Cells[31, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_Qom > Max_Qom)
                    {
                        Max_Qom = Traffic_Qom;
                        sheet4.Cells[31, 6] = Max_Qom;
                        sheet5.Cells[18, 2] = Max_Qom;
                    }
                }
                //Qazvin
                if (NE == "MSQZA")
                {
                    Traffic_Qazvin = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[32, 4] = Traffic_Qazvin;
                    sheet4.Cells[32, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_Qazvin > Max_Qazvin)
                    {
                        Max_Qazvin = Traffic_Qazvin;
                        sheet4.Cells[32, 6] = Max_Qazvin;
                        sheet5.Cells[19, 2] = Max_Qazvin;
                    }
                }


                //Tehran
                if (NE == "MSMDA")
                {
                    Traffic_Tehran = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[33, 4] = Traffic_Tehran;
                    sheet4.Cells[33, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_Tehran > Max_Tehran)
                    {
                        Max_Tehran = Traffic_Tehran;
                        sheet4.Cells[33, 6] = Max_Tehran;
                        sheet5.Cells[20, 2] = Max_Tehran;
                    }
                }
                if (NE == "MSMDB")
                {
                    Traffic_Tehran = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[34, 4] = Traffic_Tehran;
                    sheet4.Cells[34, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_Tehran > Max_Tehran)
                    {
                        Max_Tehran = Traffic_Tehran;
                        sheet4.Cells[33, 6] = Max_Tehran;
                        sheet5.Cells[20, 2] = Max_Tehran;
                    }
                }
                if (NE == "MSMDC")
                {
                    Traffic_Tehran = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[35, 4] = Traffic_Tehran;
                    sheet4.Cells[35, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_Tehran > Max_Tehran)
                    {
                        Max_Tehran = Traffic_Tehran;
                        sheet4.Cells[33, 6] = Max_Tehran;
                        sheet5.Cells[20, 2] = Max_Tehran;
                    }
                }
                if (NE == "MSMDD")
                {
                    Traffic_Tehran = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[36, 4] = Traffic_Tehran;
                    sheet4.Cells[36, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_Tehran > Max_Tehran)
                    {
                        Max_Tehran = Traffic_Tehran;
                        sheet4.Cells[33, 6] = Max_Tehran;
                        sheet5.Cells[20, 2] = Max_Tehran;
                    }
                }
                //Gilan
                if (NE == "MSRSB")
                {
                    Traffic_Gilan = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[37, 4] = Traffic_Gilan;
                    sheet4.Cells[37, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_Gilan > Max_Gilan)
                    {
                        Max_Gilan = Traffic_Gilan;
                        sheet4.Cells[37, 6] = Max_Gilan;
                        sheet5.Cells[21, 2] = Max_Gilan;
                    }
                }
                if (NE == "MSLJA")
                {
                    Traffic_Gilan = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[38, 4] = Traffic_Gilan;
                    sheet4.Cells[38, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_Gilan > Max_Gilan)
                    {
                        Max_Gilan = Traffic_Gilan;
                        sheet4.Cells[37, 6] = Max_Gilan;
                        sheet5.Cells[21, 2] = Max_Gilan;
                    }
                }
                //Semnan
                if (NE == "MSSEB")
                {
                    Traffic_Semnan = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[39, 4] = Traffic_Semnan;
                    sheet4.Cells[39, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_Semnan > Max_Semnan)
                    {
                        Max_Semnan = Traffic_Semnan;
                        sheet4.Cells[39, 6] = Max_Semnan;
                        sheet5.Cells[22, 2] = Max_Semnan;
                    }
                }
                //Kordestan
                if (NE == "MSSJA")
                {
                    Traffic_Kordestan = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[40, 4] = Traffic_Kordestan;
                    sheet4.Cells[40, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_Kordestan > Max_Kordestan)
                    {
                        Max_Kordestan = Traffic_Kordestan;
                        sheet4.Cells[40, 6] = Max_Kordestan;
                        sheet5.Cells[23, 2] = Max_Kordestan;
                    }
                }
                //Chahahr_Mahal
                if (NE == "MSSKA")
                {
                    Traffic_Chahahr_Mahal = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[41, 4] = Traffic_Chahahr_Mahal;
                    sheet4.Cells[41, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_Chahahr_Mahal > Max_Chahahr_Mahal)
                    {
                        Max_Chahahr_Mahal = Traffic_Chahahr_Mahal;
                        sheet4.Cells[41, 6] = Max_Chahahr_Mahal;
                        sheet5.Cells[24, 2] = Max_Chahahr_Mahal;
                    }
                }
                //Fars
                if (NE == "MSSZB")
                {
                    Traffic_Fars = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[42, 4] = Traffic_Fars;
                    sheet4.Cells[42, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_Fars > Max_Fars)
                    {
                        Max_Fars = Traffic_Fars;
                        sheet4.Cells[42, 6] = Max_Fars;
                        sheet5.Cells[25, 2] = Max_Fars;
                    }
                }
                if (NE == "MSSZC")
                {
                    Traffic_Fars = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[43, 4] = Traffic_Fars;
                    sheet4.Cells[43, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_Fars > Max_Fars)
                    {
                        Max_Fars = Traffic_Fars;
                        sheet4.Cells[42, 6] = Max_Fars;
                        sheet5.Cells[25, 2] = Max_Fars;
                    }
                }
                if (NE == "MSSZD")
                {
                    Traffic_Fars = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[44, 4] = Traffic_Fars;
                    sheet4.Cells[44, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_Fars > Max_Fars)
                    {
                        Max_Fars = Traffic_Fars;
                        sheet4.Cells[42, 6] = Max_Fars;
                        sheet5.Cells[25, 2] = Max_Fars;
                    }
                }
                //E_Azar
                if (NE == "MSTZA")
                {
                    Traffic_E_Azar = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[45, 4] = Traffic_E_Azar;
                    sheet4.Cells[45, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_E_Azar > Max_E_Azar)
                    {
                        Max_E_Azar = Traffic_E_Azar;
                        sheet4.Cells[45, 6] = Max_E_Azar;
                        sheet5.Cells[26, 2] = Max_E_Azar;
                    }
                }
                if (NE == "MSTZB")
                {
                    Traffic_E_Azar = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[46, 4] = Traffic_E_Azar;
                    sheet4.Cells[46, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_E_Azar > Max_E_Azar)
                    {
                        Max_E_Azar = Traffic_E_Azar;
                        sheet4.Cells[45, 6] = Max_E_Azar;
                        sheet5.Cells[26, 2] = Max_E_Azar;
                    }
                }
                if (NE == "MSTZC")
                {
                    Traffic_E_Azar = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[47, 4] = Traffic_E_Azar;
                    sheet4.Cells[47, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_E_Azar > Max_E_Azar)
                    {
                        Max_E_Azar = Traffic_E_Azar;
                        sheet4.Cells[45, 6] = Max_E_Azar;
                        sheet5.Cells[26, 2] = Max_E_Azar;
                    }
                }
                //W_Azar
                if (NE == "MSURA")
                {
                    Traffic_W_Azar = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[48, 4] = Traffic_W_Azar;
                    sheet4.Cells[48, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_W_Azar > Max_W_Azar)
                    {
                        Max_W_Azar = Traffic_W_Azar;
                        sheet4.Cells[48, 6] = Max_W_Azar;
                        sheet5.Cells[27, 2] = Max_W_Azar;
                    }
                }
                if (NE == "MSURB")
                {
                    Traffic_W_Azar = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[49, 4] = Traffic_W_Azar;
                    sheet4.Cells[49, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_W_Azar > Max_W_Azar)
                    {
                        Max_W_Azar = Traffic_W_Azar;
                        sheet4.Cells[48, 6] = Max_W_Azar;
                        sheet5.Cells[27, 2] = Max_W_Azar;
                    }
                }
                //Kohgilooyeh
                if (NE == "MSYJA")
                {
                    Traffic_Kohgilooyeh = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[50, 4] = Traffic_Kohgilooyeh;
                    sheet4.Cells[50, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_Kohgilooyeh > Max_Kohgilooyeh)
                    {
                        Max_Kohgilooyeh = Traffic_Kohgilooyeh;
                        sheet4.Cells[50, 6] = Max_Kohgilooyeh;
                        sheet5.Cells[28, 2] = Max_Kohgilooyeh;
                    }
                }
                //Yazd
                if (NE == "MSYZA")
                {
                    Traffic_Yazd = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[51, 4] = Traffic_Yazd;
                    sheet4.Cells[51, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_Yazd > Max_Yazd)
                    {
                        Max_Yazd = Traffic_Yazd;
                        sheet4.Cells[51, 6] = Max_Yazd;
                        sheet5.Cells[29, 2] = Max_Yazd;
                    }
                }
                if (NE == "MSYZB")
                {
                    Traffic_Yazd = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[52, 4] = Traffic_Yazd;
                    sheet4.Cells[52, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_Yazd > Max_Yazd)
                    {
                        Max_Yazd = Traffic_Yazd;
                        sheet4.Cells[51, 6] = Max_Yazd;
                        sheet5.Cells[29, 2] = Max_Yazd;
                    }
                }
                //Sistan_Baloch
                if (NE == "MSZHB")
                {
                    Traffic_Sistan_Baloch = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[53, 4] = Traffic_Sistan_Baloch;
                    sheet4.Cells[53, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_Sistan_Baloch > Max_Sistan_Baloch)
                    {
                        Max_Sistan_Baloch = Traffic_Sistan_Baloch;
                        sheet4.Cells[53, 6] = Max_Sistan_Baloch;
                        sheet5.Cells[30, 2] = Max_Sistan_Baloch;
                    }
                }
                //Zanjan
                if (NE == "MSZJA")
                {
                    Traffic_Zanjan = Convert.ToDouble(MSS_Traffic_Quary_Table.Rows[p].ItemArray[3]);
                    sheet4.Cells[54, 4] = Traffic_Zanjan;
                    sheet4.Cells[54, 1] = Convert.ToDateTime(MSS_Traffic_Quary_Table.Rows[p].ItemArray[0]);
                    if (Traffic_Zanjan > Max_Zanjan)
                    {
                        Max_Zanjan = Traffic_Zanjan;
                        sheet4.Cells[54, 6] = Max_Zanjan;
                        sheet5.Cells[31, 2] = Max_Zanjan;
                    }
                }

            }

            checkBox21.Invoke(new Action(() => checkBox21.Checked = true));

            MessageBox.Show("Step1 is Finished");


        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog2.DefaultExt = "txt";
            openFileDialog2.Filter = "Text File|*.txt";
            DialogResult result = openFileDialog2.ShowDialog();
            string File_Name = openFileDialog2.SafeFileName.ToString();
            if (result == DialogResult.OK)
            {
                string file = openFileDialog2.FileName;
                System.IO.StreamReader sr = new StreamReader(file);
                string line = sr.ReadLine();




                while (line != null)
                {

                    //Read the next line
                    line = sr.ReadLine();
                    if (line != null)
                    {
                        string str2 = Regex.Replace(line, "[^a-zA-Z0-9]", " ");      //هر کاراکتری که غیر از عدد و حرف بود را به کاراکتر خالی تبدیل کن
                        str2 = Regex.Replace(str2, " {2,}", " ").Trim();           //چندین کاراکتر خالی پشت سر هم را به یک کاراکتر خالی تبدیل می کند
                        string[] Split_Description = str2.Split(' ');

                        if (Split_Description.Length > 5)
                        {
                            if (Split_Description[1] == "AuC")
                            {
                                sheet6.Cells[5, 9] = Convert.ToInt64(Split_Description[5]);
                                sheet6.Cells[5, 10] = Convert.ToInt64(Split_Description[4]);
                            }
                            if (Split_Description[1] == "Number" && Split_Description[3] == "BE")
                            {
                                sheet6.Cells[6, 9] = Convert.ToInt64(Split_Description[8]);
                                sheet6.Cells[6, 10] = Convert.ToInt64(Split_Description[7]);
                            }
                            if (Split_Description[1] == "LTE" && Split_Description[2] == "BE" && Split_Description[3] == "Basic")
                            {
                                sheet6.Cells[7, 9] = Convert.ToInt64(Split_Description[6]);
                                sheet6.Cells[7, 10] = Convert.ToInt64(Split_Description[5]);
                            }
                            if (Split_Description[5] == "EAA")
                            {
                                sheet6.Cells[8, 9] = Convert.ToInt64(Split_Description[10]);
                                sheet6.Cells[8, 10] = Convert.ToInt64(Split_Description[9]);
                            }
                            if (Split_Description[1] == "Number" && Split_Description[3] == "FE")
                            {
                                sheet6.Cells[9, 9] = Convert.ToInt64(Split_Description[8]);
                                sheet6.Cells[9, 10] = Convert.ToInt64(Split_Description[7]);
                            }
                            if (Split_Description[1] == "LTE" && Split_Description[2] == "FE" && Split_Description[3] == "Basic")
                            {
                                sheet6.Cells[10, 9] = Convert.ToInt64(Split_Description[6]);
                                sheet6.Cells[10, 10] = Convert.ToInt64(Split_Description[5]);
                            }
                        }



                    }


                }




                workbook1.Save();
                workbook1.Close();
                xlApp1.Quit();
                MessageBox.Show("Report is Complete");




            }


        }

        public Excel.Application xlApp { get; set; }
        public Excel.Workbook xlWorkBook { get; set; }
        public Excel.Worksheet Sheet { get; set; }



        private void button3_Click(object sender, EventArgs e)
        {



            openFileDialog1.DefaultExt = "xlsx";
            openFileDialog1.Filter = "Excel File|*.xlsx";
            DialogResult result = openFileDialog1.ShowDialog();
            string File_Name = openFileDialog1.SafeFileName.ToString();
            string file = openFileDialog1.FileName;
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(file);
            Sheet = xlWorkBook.Worksheets[1];


//            if (Table_Name == "Tehran_CC")
//            {
//                //  ****************** CC Project Tehran  ******************

//                Server_Name = "PERFORMANCEDB01";
//                DataBase_Name = "Performance_NAK";


//                ConnectionString = @"Server=" + Server_Name + "; Database=" + DataBase_Name + "; User ID=cwpcApp;Password=cwpcApp@830625Ahmad";
//                connection = new SqlConnection(ConnectionString);
//                connection.Open();

//                Excel.Range Data = Sheet.get_Range("A2", "AJ" + Sheet.UsedRange.Rows.Count);
//                object[,] CC_Data = (object[,])Data.Value;
//                int Count = Sheet.UsedRange.Rows.Count;

//                progressBar1.Maximum = Count;

//                for (int k = 0; k < Count - 1; k++)
//                {
//                    progressBar1.Value = k+1;

//                    string TT_Code = Convert.ToString(CC_Data[k + 1, 1]);

//                    string CLARITY_TT_code = "";
//                    if (CC_Data[k + 1, 2] == null)
//                    {
//                        CLARITY_TT_code = "";
//                    }
//                    else
//                    {
//                        CLARITY_TT_code = CC_Data[k + 1, 2].ToString();
//                    }

//                    string Customer_no = "";
//                    if (CC_Data[k + 1, 3] == null)
//                    {
//                        Customer_no = "";
//                    }
//                    else
//                    {
//                        Customer_no = CC_Data[k + 1, 3].ToString();
//                    }

//                    string Service_Type_ID = "";
//                    if (CC_Data[k + 1, 4] == null)
//                    {
//                        Service_Type_ID = "";
//                    }
//                    else
//                    {
//                        Service_Type_ID = CC_Data[k + 1, 4].ToString();
//                    }


//                    string Service_Request_Title = "";
//                    if (CC_Data[k + 1, 5] == null)
//                    {
//                        Service_Request_Title = "";
//                    }
//                    else
//                    {
//                        Service_Request_Title = CC_Data[k + 1, 5].ToString();
//                    }

//                    string Service_type_full_name = "";
//                    if (CC_Data[k + 1, 6] == null)
//                    {
//                        Service_type_full_name = "";
//                    }
//                    else
//                    {
//                        Service_type_full_name = CC_Data[k + 1, 6].ToString();
//                    }

//                    string Problem_Category = "";
//                    if (CC_Data[k + 1, 7] == null)
//                    {
//                        Problem_Category = "";
//                    }
//                    else
//                    {
//                        Problem_Category = CC_Data[k + 1, 7].ToString();
//                    }

//                    string agent_Province = "";
//                    if (CC_Data[k + 1, 8] == null)
//                    {
//                        agent_Province = "";
//                    }
//                    else
//                    {
//                        agent_Province = CC_Data[k + 1, 8].ToString();
//                    }

//                    string Creation_Date_Georgian = "";
//                    if (CC_Data[k + 1, 9] == null)
//                    {
//                        Creation_Date_Georgian = "";
//                    }
//                    else
//                    {
//                        Creation_Date_Georgian = CC_Data[k + 1, 9].ToString();
//                    }

//                    string Creation_Date_Jalali = "";
//                    if (CC_Data[k + 1, 10] == null)
//                    {
//                        Creation_Date_Jalali = "";
//                    }
//                    else
//                    {
//                        Creation_Date_Jalali = CC_Data[k + 1, 10].ToString();
//                    }

//                    string Agent_ID = "";
//                    if (CC_Data[k + 1, 11] == null)
//                    {
//                        Agent_ID = "";
//                    }
//                    else
//                    {
//                        Agent_ID = CC_Data[k + 1, 11].ToString();
//                    }

//                    string Agent_Name = "";
//                    if (CC_Data[k + 1, 12] == null)
//                    {
//                        Agent_Name = "";
//                    }
//                    else
//                    {
//                        Agent_Name = CC_Data[k + 1, 12].ToString();
//                    }


//                    string Customer_segment = "";
//                    if (CC_Data[k + 1, 13] == null)
//                    {
//                        Customer_segment = "";
//                    }
//                    else
//                    {
//                        Customer_segment = CC_Data[k + 1, 13].ToString();
//                    }

//                    string channel = "";
//                    if (CC_Data[k + 1, 14] == null)
//                    {
//                        channel = "";
//                    }
//                    else
//                    {
//                        channel = CC_Data[k + 1, 14].ToString();
//                    }

//                    string Customer_Province = "";
//                    if (CC_Data[k + 1, 15] == null)
//                    {
//                        Customer_Province = "";
//                    }
//                    else
//                    {
//                        Customer_Province = CC_Data[k + 1, 15].ToString();
//                    }



//                    string Customer_City = "";
//                    if (CC_Data[k + 1, 16] == null)
//                    {
//                        Customer_City = "";
//                    }
//                    else
//                    {
//                        Customer_City = CC_Data[k + 1, 16].ToString();
//                    }


//                    string TTQL026_TTQL374_TTQL059_TTQL295 = "";
//                    if (CC_Data[k + 1, 17] == null)
//                    {
//                        TTQL026_TTQL374_TTQL059_TTQL295 = "";
//                    }
//                    else
//                    {
//                        TTQL026_TTQL374_TTQL059_TTQL295 = CC_Data[k + 1, 17].ToString();
//                        if (TTQL026_TTQL374_TTQL059_TTQL295.Length>255)
//                        {
//                            TTQL026_TTQL374_TTQL059_TTQL295 = TTQL026_TTQL374_TTQL059_TTQL295.Substring(0, 255);
//                        }

//                    }


//                    string TTQL008 = "";
//                    if (CC_Data[k + 1, 18] == null)
//                    {
//                        TTQL008 = "";
//                    }
//                    else
//                    {
//                        TTQL008 = CC_Data[k + 1, 18].ToString();
//                        if (TTQL008.Length > 255)
//                        {
//                            TTQL008 = TTQL008.Substring(0, 255);
//                        }
//                    }

//                    string TTQL006 = "";
//                    if (CC_Data[k + 1, 19] == null)
//                    {
//                        TTQL006 = "";
//                    }
//                    else
//                    {
//                        TTQL006 = CC_Data[k + 1, 19].ToString();
//                        if (TTQL006.Length > 255)
//                        {
//                            TTQL006 = TTQL006.Substring(0, 255);
//                        }

//                    }

//                    string TTQL133 = "";
//                    if (CC_Data[k + 1, 20] == null)
//                    {
//                        TTQL133 = "";
//                    }
//                    else
//                    {
//                        TTQL133 = CC_Data[k + 1, 20].ToString();
//                        if (TTQL133.Length > 255)
//                        {
//                            TTQL133 = TTQL133.Substring(0, 255);
//                        }
//                    }


//                    string TTQL551 = "";
//                    if (CC_Data[k + 1, 21] == null)
//                    {
//                        TTQL551 = "";
//                    }
//                    else
//                    {
//                        TTQL551 = CC_Data[k + 1, 21].ToString();
//                        if (TTQL551.Length > 255)
//                        {
//                            TTQL551 = TTQL551.Substring(0, 255);
//                        }
//                    }

//                    string TTQL552 = "";
//                    if (CC_Data[k + 1, 22] == null)
//                    {
//                        TTQL552 = "";
//                    }
//                    else
//                    {
//                        TTQL552 = CC_Data[k + 1, 22].ToString();
//                        if (TTQL552.Length > 255)
//                        {
//                            TTQL552 = TTQL552.Substring(0, 255);
//                        }
//                    }

//                    string TTQL553 = "";
//                    if (CC_Data[k + 1, 23] == null)
//                    {
//                        TTQL553 = "";
//                    }
//                    else
//                    {
//                        TTQL553 = CC_Data[k + 1, 23].ToString();
//                        if (TTQL553.Length > 255)
//                        {
//                            TTQL553 = TTQL553.Substring(0, 255);
//                        }
//                    }

//                    string Latitude = "";
//                    if (CC_Data[k + 1, 24] == null)
//                    {
//                        Latitude = "";
//                    }
//                    else
//                    {
//                        Latitude = CC_Data[k + 1, 24].ToString();
//                    }


//                    string Longitude = "";
//                    if (CC_Data[k + 1, 25] == null)
//                    {
//                        Longitude = "";
//                    }
//                    else
//                    {
//                        Longitude = CC_Data[k + 1, 25].ToString();
//                    }



//                    string Customer_type = "";
//                    if (CC_Data[k + 1, 26] == null)
//                    {
//                        Customer_type = "";
//                    }
//                    else
//                    {
//                        Customer_type = CC_Data[k + 1, 26].ToString();
//                    }


//                    string main_sub_TT = "";
//                    if (CC_Data[k + 1, 27] == null)
//                    {
//                        main_sub_TT = "";
//                    }
//                    else
//                    {
//                        main_sub_TT = CC_Data[k + 1, 27].ToString();
//                    }



//                    string The_Last_Agent_ID = "";
//                    if (CC_Data[k + 1, 28] == null)
//                    {
//                        The_Last_Agent_ID = "";
//                    }
//                    else
//                    {
//                        The_Last_Agent_ID = CC_Data[k + 1, 28].ToString();
//                    }


//                    string The_Last_Agent_Name = "";
//                    if (CC_Data[k + 1, 29] == null)
//                    {
//                        The_Last_Agent_Name = "";
//                    }
//                    else
//                    {
//                        The_Last_Agent_Name = CC_Data[k + 1, 29].ToString();
//                    }

//                    string Delay_Request_Count = "";
//                    if (CC_Data[k + 1, 30] == null)
//                    {
//                        Delay_Request_Count = "";
//                    }
//                    else
//                    {
//                        Delay_Request_Count = CC_Data[k + 1, 30].ToString();
//                    }


//                    string Delay_Request_Date = "";
//                    if (CC_Data[k + 1, 31] == null)
//                    {
//                        Delay_Request_Date = "";
//                    }
//                    else
//                    {
//                        Delay_Request_Date = CC_Data[k + 1, 31].ToString();
//                    }

//                    string Delay_Applied_Date = "";
//                    if (CC_Data[k + 1, 32] == null)
//                    {
//                        Delay_Applied_Date = "";
//                    }
//                    else
//                    {
//                        Delay_Applied_Date = CC_Data[k + 1, 32].ToString();
//                    }


//                    string Ticket_Status = "";
//                    if (CC_Data[k + 1, 33] == null)
//                    {
//                        Ticket_Status = "";
//                    }
//                    else
//                    {
//                        Ticket_Status = CC_Data[k + 1, 33].ToString();
//                    }


//                    string Last_Handle_Opinion = "";
//                    if (CC_Data[k + 1, 34] == null)
//                    {
//                        Last_Handle_Opinion = "";
//                    }
//                    else
//                    {
//                        Last_Handle_Opinion = CC_Data[k + 1, 34].ToString();
//                    }

//                    string Is_Rejected = "";
//                    if (CC_Data[k + 1, 35] == null)
//                    {
//                        Is_Rejected = "";
//                    }
//                    else
//                    {
//                        Is_Rejected = CC_Data[k + 1, 35].ToString();
//                    }


//                    string RNC = "";
//                    if (CC_Data[k + 1, 36] == null)
//                    {
//                        RNC = "";
//                    }
//                    else
//                    {
//                        RNC = CC_Data[k + 1, 36].ToString();
//                    }

//                    string Site =  "";

//                    DateTime Date = Convert.ToDateTime(file.Substring(file.Length - 13, 4) + "-" + file.Substring(file.Length - 9, 2) + "-" + file.Substring(file.Length - 7, 2));

//                    string DataFill ="'" + TT_Code + "','" + CLARITY_TT_code + "','" + Customer_no + "','" + Service_Type_ID + "','" +
//                                           Service_Request_Title + "','" + Service_type_full_name + "','" + Problem_Category + "','" +
//                                           agent_Province + "','" + Creation_Date_Georgian + "','" + Creation_Date_Jalali + "','" +
//                                           Agent_ID + "','" + Agent_Name + "','" + Customer_segment + "','" + channel + "','" +
//                                           Customer_Province + "','" + Customer_City + "','" + TTQL026_TTQL374_TTQL059_TTQL295 + "','" +
//                                           TTQL008 + "','" + TTQL006 + "','" + TTQL133 + "','" + TTQL551 + "','" + TTQL552 + "','" + TTQL553 + "','" +
//                                           Latitude + "','" + Longitude + "','" + Customer_type + "','" + main_sub_TT + "','" + The_Last_Agent_ID + "','" +
//                                           The_Last_Agent_Name + "','" + Delay_Request_Count + "','" + Delay_Request_Date + "','" + Delay_Applied_Date + "','" +
//                                           Ticket_Status + "','" + Last_Handle_Opinion + "','" + Is_Rejected + "','" + RNC + "','" + Site + "','" + Date + "'";


//                    string SS = @"Insert into [Tehran_CC] ([TT Code], 
//[CLARITY TT code], 
//[Customer no], 
//[Service Type ID], 
//[Service Request Title], 
//[Service type full name], 
//[Problem Category], 
//[agent Province], 
//[Creation Date_Georgian], 
//[Creation Date_Jalali], 
//[Agent ID], 
//[Agent Name], 
//[Customer segment], 
//[channel], 
//[Customer Province], 
//[Customer City], 
//[TTQL026-TTQL374-TTQL059-TTQL295], 
//[TTQL008], 
//[TTQL006], 
//[TTQL133], 
//[ TTQL551], 
//[TTQL552], 
//[TTQL553], 
//[Latitude], 
//[Longitude], 
//[Customer type], 
//[main^sub -TT ], 
//[The Last Agent ID ], 
//[The Last Agent Name], 
//[Delay Request (Count)], 
//[Delay Request Date], 
//[Delay Applied Date], 
//[Ticket Status], 
//[Last Handle Opinion], 
//[Is Rejected?], 
//[RNC], 
//[Site], 
//[Date]
//) values(" + DataFill + ")";
//                    SqlCommand SS_Query = new SqlCommand(SS, connection);
//                    SS_Query.ExecuteNonQuery();

//                }
//                connection.Close();
//                MessageBox.Show("Finished");


//            }






            //  ****************** CC Project Tehran  ******************

           


            if (Table_Name == "Traffic_Sub")
            {


                //ConnectionString = @"Server=" + Server_Name + "; Database=" + DataBase_Name + "; Integrated Security=True";
                ConnectionString = @"Server=" + Server_Name + "; Database=" + DataBase_Name + "; User ID=Ahmad_Core; Password=cwpcApp@830625Ahmad";
                connection = new SqlConnection(ConnectionString);
                connection.Open();


                //    // Other Orders to Fill Table
                //    string IMPORT_STR_1 = string.Format(@"INSERT INTO [") + Table_Name;
                //    string IMPORT_STR_2 = string.Format(@"] select [Date],
                //[NE],
                //[Traffic_Per_Sub(MErl)(Nokia_Core)],
                //[Traffic_Per_Sub_mErl_HU] 
                //from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...[", file);
                //    string IMPORT_STR_3 = Table_Name + string.Format(@"$] order by Date", file);
                //    string Insert_Table = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;


                //    SqlCommand Insert_Table_command = new SqlCommand(Insert_Table, connection);
                //    Insert_Table_command.ExecuteNonQuery();


                DataTable Traffic_Sub_Table = new DataTable();
                Traffic_Sub_Table.Columns.Add("Date", typeof(DateTime));
                Traffic_Sub_Table.Columns.Add("NE", typeof(String));
                Traffic_Sub_Table.Columns.Add("Nokia_Traffic", typeof(double));
                Traffic_Sub_Table.Columns.Add("Huawei_Traffic", typeof(double));



                Excel.Range Data = Sheet.get_Range("A2", "D" + Sheet.UsedRange.Rows.Count);
                object[,] Core_Data = (object[,])Data.Value;
                int Count = Sheet.UsedRange.Rows.Count;



                for (int k = 0; k < Count - 1; k++)
                {

                    if (Core_Data[k + 1, 3] == null && Core_Data[k + 1, 4] == null)
                    {
                        continue;
                    }

                    DateTime Date = Convert.ToDateTime(Core_Data[k + 1, 1]);
                    string NE = Core_Data[k + 1, 2].ToString();

                    double Nokia_Traffic = 0;
                    if (Core_Data[k + 1, 3] != null)
                    {
                        Nokia_Traffic = Convert.ToDouble(Core_Data[k + 1, 3]);
                    }

                    double Huawei_Traffic = 0;
                    if (Core_Data[k + 1, 4] != null)
                    {
                        Huawei_Traffic = Convert.ToDouble(Core_Data[k + 1, 4]);
                    }

                    //string DataFill = "'" + Convert.ToString(Date) + "','" + NE + "','" + Nokia_Traffic + "','" + Huawei_Traffic + "'";


                    Traffic_Sub_Table.Rows.Add(Date, NE, Nokia_Traffic, Huawei_Traffic);

                    //string SS = @"Insert into [Traffic_Sub] ([Date],NE,[Traffic_Per_Sub(MErl)(Nokia_Core)],[Traffic_Per_Sub_mErl_HU]) values(" + DataFill + ")";
                    //SqlCommand SS_Query = new SqlCommand(SS, connection);
                    //SS_Query.ExecuteNonQuery();
                }

                SqlBulkCopy objbulk_Traffic_Sub = new SqlBulkCopy(connection);
                objbulk_Traffic_Sub.DestinationTableName = "Traffic_Sub";
                objbulk_Traffic_Sub.ColumnMappings.Add("Date", "Date");
                objbulk_Traffic_Sub.ColumnMappings.Add("NE", "NE");
                objbulk_Traffic_Sub.ColumnMappings.Add("Nokia_Traffic", "Traffic_Per_Sub(MErl)(Nokia_Core)");
                objbulk_Traffic_Sub.ColumnMappings.Add("Huawei_Traffic", "Traffic_Per_Sub_mErl_HU");
                objbulk_Traffic_Sub.WriteToServer(Traffic_Sub_Table);


                connection.Close();
                MessageBox.Show("Traffic_Sub is Filled into DB");
            }


            if (Table_Name == "Nokia_MSS")
            {

                //ConnectionString = @"Server=" + Server_Name + "; Database=" + DataBase_Name + "; Integrated Security=True";
                ConnectionString = @"Server=" + Server_Name + "; Database=" + DataBase_Name + "; User ID=Ahmad_Core; Password=cwpcApp@830625Ahmad";
                connection = new SqlConnection(ConnectionString);
                connection.Open();





                DataTable Nokia_MSS_Table = new DataTable();
                Nokia_MSS_Table.Columns.Add("Date", typeof(DateTime));
                Nokia_MSS_Table.Columns.Add("NE", typeof(String));
                Nokia_MSS_Table.Columns.Add("MSS_Licence_Capacity", typeof(double));
                Nokia_MSS_Table.Columns.Add("MSS_Peak_Lic_Utilization", typeof(double));




                //    // Other Orders to Fill Table
                //    string IMPORT_STR_1 = string.Format(@"INSERT INTO [") + Table_Name;
                //    string IMPORT_STR_2 = string.Format(@"] select [Date],
                //[NE],
                //[MSS_Licence_Capacity(Nokia_Core)],
                //[MSS_Peak_Lic_Utilization(Nokia(Core)] 
                //from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...[", file);
                //    string IMPORT_STR_3 = Table_Name + string.Format(@"$] order by Date", file);
                //    string Insert_Table = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;



                Excel.Range Data = Sheet.get_Range("A2", "D" + Sheet.UsedRange.Rows.Count);
                object[,] Core_Data = (object[,])Data.Value;
                int Count = Sheet.UsedRange.Rows.Count;


                for (int k = 0; k < Count - 1; k++)
                {

                    if (Core_Data[k + 1, 3] == null && Core_Data[k + 1, 4] == null)
                    {
                        continue;
                    }

                    DateTime Date = Convert.ToDateTime(Core_Data[k + 1, 1]);
                    string NE = Core_Data[k + 1, 2].ToString();

                    double MSS_Licence_Capacity = 0;
                    if (Core_Data[k + 1, 3] != null)
                    {
                        MSS_Licence_Capacity = Convert.ToDouble(Core_Data[k + 1, 3]);
                    }

                    double MSS_Peak_Lic_Utilization = 0;
                    if (Core_Data[k + 1, 4] != null)
                    {
                        MSS_Peak_Lic_Utilization = Convert.ToDouble(Core_Data[k + 1, 4]);
                    }


                    Nokia_MSS_Table.Rows.Add(Date, NE, MSS_Licence_Capacity, MSS_Peak_Lic_Utilization);


                    //string DataFill = "'" + Convert.ToString(Date) + "','" + NE + "','" + MSS_Licence_Capacity + "','" + MSS_Peak_Lic_Utilization + "'";

                    //string SS = @"Insert into [Nokia_MSS] ([Date],NE, [MSS_Licence_Capacity(Nokia_Core)],[MSS_Peak_Lic_Utilization(Nokia(Core)] ) values(" + DataFill + ")";
                    //SqlCommand SS_Query = new SqlCommand(SS, connection);
                    //SS_Query.ExecuteNonQuery();
                }


                SqlBulkCopy objbulk_Nokia_MSS = new SqlBulkCopy(connection);
                objbulk_Nokia_MSS.DestinationTableName = "Nokia_MSS";
                objbulk_Nokia_MSS.ColumnMappings.Add("Date", "Date");
                objbulk_Nokia_MSS.ColumnMappings.Add("NE", "NE");
                objbulk_Nokia_MSS.ColumnMappings.Add("MSS_Licence_Capacity", "MSS_Licence_Capacity(Nokia_Core)");
                objbulk_Nokia_MSS.ColumnMappings.Add("MSS_Peak_Lic_Utilization", "MSS_Peak_Lic_Utilization(Nokia(Core)");
                objbulk_Nokia_MSS.WriteToServer(Nokia_MSS_Table);

                connection.Close();
                MessageBox.Show("Nokia_MSS is Filled into DB");
            }




            if (Table_Name == "Nokia_MGW")
            {
                //ConnectionString = @"Server=" + Server_Name + "; Database=" + DataBase_Name + "; Integrated Security=True";
                ConnectionString = @"Server=" + Server_Name + "; Database=" + DataBase_Name + "; User ID=Ahmad_Core; Password=cwpcApp@830625Ahmad";
                connection = new SqlConnection(ConnectionString);
                connection.Open();


                //    // Other Orders to Fill Table
                //    string IMPORT_STR_1 = string.Format(@"INSERT INTO [") + Table_Name;
                //    string IMPORT_STR_2 = string.Format(@"] select [Date],
                //[NE],
                //[MSS_Licence_Capacity(Nokia_Core)],
                //[MSS_Peak_Lic_Utilization(Nokia(Core)] 
                //from OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0','Data Source=""{0}""; Extended Properties=""Excel 12.0; HDR = Yes""')...[", file);
                //    string IMPORT_STR_3 = Table_Name + string.Format(@"$] order by Date", file);
                //    string Insert_Table = IMPORT_STR_1 + IMPORT_STR_2 + IMPORT_STR_3;



                DataTable Nokia_MGW_Table = new DataTable();
                Nokia_MGW_Table.Columns.Add("Date", typeof(DateTime));
                Nokia_MGW_Table.Columns.Add("NE", typeof(String));
                Nokia_MGW_Table.Columns.Add("capacity_licence_utilization_Peak", typeof(double));
                Nokia_MGW_Table.Columns.Add("CC_Feature_capacity_Nokia", typeof(double));
                Nokia_MGW_Table.Columns.Add("IU_IP_feature_capacity", typeof(double));
                Nokia_MGW_Table.Columns.Add("IU_IP_PEAK", typeof(double));
                Nokia_MGW_Table.Columns.Add("NB_IP_FEATURE_CAPACITY", typeof(double));
                Nokia_MGW_Table.Columns.Add("NB_IP_Peak_License", typeof(double));
                Nokia_MGW_Table.Columns.Add("AOIP_feature_capacity", typeof(double));
                Nokia_MGW_Table.Columns.Add("AOIP_peak", typeof(double));
                Nokia_MGW_Table.Columns.Add("MB_FEATURE_CAPACITY", typeof(double));
                Nokia_MGW_Table.Columns.Add("MB_Peak_License", typeof(double));
                Nokia_MGW_Table.Columns.Add("PC_ATER_FEATURE_CAPACITY", typeof(double));
                Nokia_MGW_Table.Columns.Add("Ater_Peak_License", typeof(double));


                Excel.Range Data = Sheet.get_Range("A2", "N" + Sheet.UsedRange.Rows.Count);
                object[,] Core_Data = (object[,])Data.Value;
                int Count = Sheet.UsedRange.Rows.Count;


                for (int k = 0; k < Count - 1; k++)
                {

                    if (Core_Data[k + 1, 3] == null)
                    {
                        continue;
                    }

                    DateTime Date = Convert.ToDateTime(Core_Data[k + 1, 1]);
                    string NE = Core_Data[k + 1, 2].ToString();

                    double capacity_licence_utilization_Peak = 0;
                    if (Core_Data[k + 1, 3] != null)
                    {
                        capacity_licence_utilization_Peak = Convert.ToDouble(Core_Data[k + 1, 3]);
                    }


                    double CC_Feature_capacity_Nokia = 0;
                    if (Core_Data[k + 1, 4] != null)
                    {
                        CC_Feature_capacity_Nokia = Convert.ToDouble(Core_Data[k + 1, 4]);
                    }



                    double IU_IP_feature_capacity = 0;
                    if (Core_Data[k + 1, 5] != null)
                    {
                        IU_IP_feature_capacity = Convert.ToDouble(Core_Data[k + 1, 5]);
                    }



                    double IU_IP_PEAK = 0;
                    if (Core_Data[k + 1, 6] != null)
                    {
                        IU_IP_PEAK = Convert.ToDouble(Core_Data[k + 1, 6]);
                    }



                    double NB_IP_FEATURE_CAPACITY = 0;
                    if (Core_Data[k + 1, 7] != null)
                    {
                        NB_IP_FEATURE_CAPACITY = Convert.ToDouble(Core_Data[k + 1, 7]);
                    }




                    double NB_IP_Peak_License = 0;
                    if (Core_Data[k + 1, 8] != null)
                    {
                        NB_IP_Peak_License = Convert.ToDouble(Core_Data[k + 1, 8]);
                    }


                    double AOIP_feature_capacity = 0;
                    if (Core_Data[k + 1, 9] != null)
                    {
                        AOIP_feature_capacity = Convert.ToDouble(Core_Data[k + 1, 9]);
                    }


                    double AOIP_peak = 0;
                    if (Core_Data[k + 1, 10] != null)
                    {
                        AOIP_peak = Convert.ToDouble(Core_Data[k + 1, 10]);
                    }



                    double MB_FEATURE_CAPACITY = 0;
                    if (Core_Data[k + 1, 11] != null)
                    {
                        MB_FEATURE_CAPACITY = Convert.ToDouble(Core_Data[k + 1, 11]);
                    }




                    double MB_Peak_License = 0;
                    if (Core_Data[k + 1, 12] != null)
                    {
                        MB_Peak_License = Convert.ToDouble(Core_Data[k + 1, 12]);
                    }



                    double PC_ATER_FEATURE_CAPACITY = 0;
                    if (Core_Data[k + 1, 13] != null)
                    {
                        PC_ATER_FEATURE_CAPACITY = Convert.ToDouble(Core_Data[k + 1, 13]);
                    }



                    double Ater_Peak_License = 0;
                    if (Core_Data[k + 1, 14] != null)
                    {
                        Ater_Peak_License = Convert.ToDouble(Core_Data[k + 1, 14]);
                    }


                    //                    string DataFill = "'" + Convert.ToString(Date) + "','" + NE + "','" + capacity_licence_utilization_Peak + "','" + CC_Feature_capacity_Nokia+ "','"+
                    //                                       IU_IP_feature_capacity + "','" + IU_IP_PEAK + "','" + NB_IP_FEATURE_CAPACITY +"','" + NB_IP_Peak_License + "','" +
                    //                                       AOIP_feature_capacity + "','"+ AOIP_peak + "','" + MB_FEATURE_CAPACITY + "','" + MB_Peak_License + "','" + PC_ATER_FEATURE_CAPACITY + "','" + Ater_Peak_License+"'";


                    Nokia_MGW_Table.Rows.Add(Date, NE, capacity_licence_utilization_Peak, CC_Feature_capacity_Nokia, IU_IP_feature_capacity, IU_IP_PEAK, NB_IP_FEATURE_CAPACITY, NB_IP_Peak_License, AOIP_feature_capacity, AOIP_peak, MB_FEATURE_CAPACITY, MB_Peak_License, PC_ATER_FEATURE_CAPACITY, Ater_Peak_License);


                }



                SqlBulkCopy objbulk_Nokia_MGW = new SqlBulkCopy(connection);
                objbulk_Nokia_MGW.DestinationTableName = "Nokia_MGW";
                objbulk_Nokia_MGW.ColumnMappings.Add("Date", "Date");
                objbulk_Nokia_MGW.ColumnMappings.Add("NE", "NE");
                objbulk_Nokia_MGW.ColumnMappings.Add("capacity_licence_utilization_Peak", "capacity_licence_utilization_Peak");
                objbulk_Nokia_MGW.ColumnMappings.Add("CC_Feature_capacity_Nokia", "CC_Feature_capacity_Nokia");
                objbulk_Nokia_MGW.ColumnMappings.Add("IU_IP_feature_capacity", "IU_IP_feature_capacity");
                objbulk_Nokia_MGW.ColumnMappings.Add("IU_IP_PEAK", "IU_IP_PEAK");
                objbulk_Nokia_MGW.ColumnMappings.Add("NB_IP_FEATURE_CAPACITY", "NB_IP_FEATURE_CAPACITY");
                objbulk_Nokia_MGW.ColumnMappings.Add("NB_IP_Peak_License", "NB_IP_Peak_License");
                objbulk_Nokia_MGW.ColumnMappings.Add("AOIP_feature_capacity", "AOIP_feature_capacity");
                objbulk_Nokia_MGW.ColumnMappings.Add("AOIP_peak", "AOIP_peak");
                objbulk_Nokia_MGW.ColumnMappings.Add("MB_FEATURE_CAPACITY", "MB_FEATURE_CAPACITY");
                objbulk_Nokia_MGW.ColumnMappings.Add("MB_Peak_License", "MB_Peak_License");
                objbulk_Nokia_MGW.ColumnMappings.Add("PC_ATER_FEATURE_CAPACITY", "PC_ATER_FEATURE_CAPACITY(Nokia_Core)");
                objbulk_Nokia_MGW.ColumnMappings.Add("Ater_Peak_License", "Ater_Peak_License");
                objbulk_Nokia_MGW.WriteToServer(Nokia_MGW_Table);


                connection.Close();
                MessageBox.Show("Nokia_MGW is Filled into DB");
            }


        }






        public string Table_Name = "";
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Table_Name = comboBox1.SelectedItem.ToString();
        }
    }
}
