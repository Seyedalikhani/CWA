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



//RNC Owner   Email address   RNC
//Ahmad Amiri R031E
//Ahmad Amiri R025E
//Mahya Goudarzi  R054H
//Mahya Goudarzi  R055H
//Nasrin Rezania  R024E
//Nasrin Rezania  R023E
//Newsha Sanaei   R071E
//Newsha Sanaei   R044H
//Newsha Sanaei   R042E
//Newsha Sanaei   R041E
//Rahim Habibi    R081E
//Rahim Habibi    R043E
//Sepideh Pour Ebrahim R033E
//Sepideh Pour Ebrahim R072E
//Shadi Mohabati  R021E
//Shadi Mohabati  R022E
//Zahra Bakhti    R065H
//Zahra Bakhti    R064H





namespace CWA
{
    public partial class Form13 : Form
    {
        public Form13()
        {
            InitializeComponent();
        }


        public Form1 form1;


        public Form13(Form form)
        {
            InitializeComponent();
            form1 = (Form1)form;
        }

        private void Form13_Load(object sender, EventArgs e)
        {

        }



        public Excel.Application xlApp { get; set; }
        public Excel.Workbook xlWorkBook { get; set; }
        public Excel.Worksheet Sheet { get; set; }



        public string Server_Name = "";
        public string DataBase_Name = "";
        public string ConnectionString = "";
        public SqlConnection connection = new SqlConnection();


        // Method of Convert Datetime to String
        public string Date_ToString(DateTime D1)
        {
            DateTime Last_Updated_Date = Convert.ToDateTime(D1);
            string Last_Month = "";
            if (Last_Updated_Date.Month <= 9)
            {
                Last_Month = "0" + Convert.ToString(Last_Updated_Date.Month);
            }
            else
            {
                Last_Month = Convert.ToString(Last_Updated_Date.Month);
            }
            string Last_Day = "";
            if (Last_Updated_Date.Day <= 9)
            {
                Last_Day = "0" + Convert.ToString(Last_Updated_Date.Day);
            }
            else
            {
                Last_Day = Convert.ToString(Last_Updated_Date.Day);
            }
            string Last_Updated_Date_String = Convert.ToString(Last_Updated_Date.Year) + "-" + Last_Month + "-" + Last_Day;
            return Last_Updated_Date_String;
        }


        // Method of Quaery Execution with Output
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

        // Method of Quaery Execution without Output
        void Query_Execution(String Query)
        {
            string Quary_String = Query;
            SqlCommand Quary_Command = new SqlCommand(Quary_String, connection);
            Quary_Command.CommandTimeout = 0;
            Quary_Command.ExecuteNonQuery();
        }

        private void button1_Click(object sender, EventArgs e)
        {


            Server_Name = "PERFORMANCEDB01";
            DataBase_Name = "Performance_NAK";


            ConnectionString = @"Server=" + Server_Name + "; Database=" + DataBase_Name + "; User ID=cwpcApp;Password=cwpcApp@830625#Ahmad";
            connection = new SqlConnection(ConnectionString);
            connection.Open();


            // Last Date
            string Max_Date_Quary = @"Select max(Date) from Tehran_CC_Maintable";
            DataTable Max_Date_Table = Query_Execution_Table_Output(Max_Date_Quary);
            // Convert Last Date to String
            DateTime Last_Updated_Date = Convert.ToDateTime((Max_Date_Table.Rows[Max_Date_Table.Rows.Count - 1]).ItemArray[0]);
            string Last_Updated_Date_String = Date_ToString(Last_Updated_Date);





            // Date Table 
            string Date_Quary = @"Select Distinct Date from Tehran_CC_Maintable where cast(Date as Date)>='2022-07-31' order by Date";                        // From mBegining
            //string Date_Quary = @"Select Distinct Date from Tehran_CC_Maintable where cast(Date as Date)>'"+ Last_Updated_Date_String+"' order by Date";    // From Last Date
            DataTable Date_Table= Query_Execution_Table_Output(Date_Quary);


            // RNC Table
            string RNC_Quary = @"Select Distinct RNC from Tehran_CC_Maintable where substring(RNC,1,1)='R'";
            DataTable RNC_Table= Query_Execution_Table_Output(RNC_Quary);


            // We define them beacuse we must fill 0 in cases that data is null
            DataTable RemainTTs_Table_Default = new DataTable();
            RemainTTs_Table_Default.Columns.Add("Date", typeof(DateTime));
            RemainTTs_Table_Default.Columns.Add("RNC", typeof(String));
            RemainTTs_Table_Default.Columns.Add("NewTTs", typeof(int));


            DataTable ReadyToCloseTTs_Table_Default = new DataTable();
            ReadyToCloseTTs_Table_Default.Columns.Add("Date", typeof(DateTime));
            ReadyToCloseTTs_Table_Default.Columns.Add("RNC", typeof(String));
            ReadyToCloseTTs_Table_Default.Columns.Add("ReadyToCloseTTs", typeof(int));



            DataTable CloseTTs_Table_Default = new DataTable();
            CloseTTs_Table_Default.Columns.Add("Date", typeof(DateTime));
            CloseTTs_Table_Default.Columns.Add("RNC", typeof(String));
            CloseTTs_Table_Default.Columns.Add("CloseTTs", typeof(int));

            // Remove Park contains reals close TTs that [Ticket Status] is not 'پارک'  or 'پارک سررسید گذشته'
            DataTable CloseTTsRemovePark_Table_Default = new DataTable();
            CloseTTsRemovePark_Table_Default.Columns.Add("Date", typeof(DateTime));
            CloseTTsRemovePark_Table_Default.Columns.Add("RNC", typeof(String));
            CloseTTsRemovePark_Table_Default.Columns.Add("CloseTTs", typeof(int));




            // شاخصای اینکامینگ و ردی تو کلوز چون به گذشته و حال مربوط میشوند فقط در تاریخهای جدید به روز میشوند 
            //اما در شاخص کلوز چون به زمان آینده مربوط میشود هر بار از ابتدا نیاز به انجام محاسبات است


            // Date of Incoming Table
            string Incoming_Date_Quary = @"Select Distinct Date from Tehran_CC_Remain where cast(Date as Date)>='2022-07-31' order by Date";
            DataTable Incoming_Date_Table = Query_Execution_Table_Output(Incoming_Date_Quary);
            DateTime Last_Updated_Date_Incoming = Convert.ToDateTime((Incoming_Date_Table.Rows[Incoming_Date_Table.Rows.Count-1]).ItemArray[0]);
            // اگر تاریخی در جدول اصلی باشد که در ریمن نباشد فقط برای آن جدول تمام صفر را را ایجاد میکنیم
            int New_Incoming_candidate_date = 0;
            //DateTime candidate_date_Incoming = Last_Updated_Date;
            for (int k = 1; k < Date_Table.Rows.Count; k++)
            {
                New_Incoming_candidate_date = 0;
                DateTime candidate_date = Convert.ToDateTime((Date_Table.Rows[k]).ItemArray[0]);
                for (int i = 0; i < Incoming_Date_Table.Rows.Count; i++)
                {
                    DateTime candidate_date_Incoming = Convert.ToDateTime((Incoming_Date_Table.Rows[i]).ItemArray[0]);
                    if (candidate_date== candidate_date_Incoming)
                    {
                        New_Incoming_candidate_date++;
                    }

                }
                if (New_Incoming_candidate_date==0)
                {
                    for (int j = 0; j < RNC_Table.Rows.Count; j++)
                    {
                        string RNC = Convert.ToString((RNC_Table.Rows[j]).ItemArray[0]);
                        RemainTTs_Table_Default.Rows.Add(candidate_date, RNC, 0);
                    }
                }
            }

            SqlBulkCopy objbulk_Remain = new SqlBulkCopy(connection);
            objbulk_Remain.DestinationTableName = "Tehran_CC_Remain";
            objbulk_Remain.ColumnMappings.Add("Date", "Date");
            objbulk_Remain.ColumnMappings.Add("RNC", "RNC");
            objbulk_Remain.ColumnMappings.Add("NewTTs", "NewTTs");
            objbulk_Remain.WriteToServer(RemainTTs_Table_Default);





            //Date of ReadyToClose Table
            string ReadyToClose_Date_Quary = @"Select Distinct Date from Tehran_CC_ReadyToClose where cast(Date as Date)>='2022-07-31' order by Date";
            DataTable ReadyToClose_Date_Table = Query_Execution_Table_Output(ReadyToClose_Date_Quary);
            DateTime Last_Updated_Date_ReadyToClose = Convert.ToDateTime((ReadyToClose_Date_Table.Rows[ReadyToClose_Date_Table.Rows.Count - 1]).ItemArray[0]);
            // اگر تاریخی در جدول اصلی باشد که در ریمن نباشد فقط برای آن جدول تمام صفر را را ایجاد میکنیم
            int New_ReadyToClose_candidate_date = 0;
            //DateTime candidate_date_ReadyToClose = Last_Updated_Date;
            for (int k = 1; k < Date_Table.Rows.Count; k++)
            {
                New_ReadyToClose_candidate_date = 0;
                DateTime candidate_date = Convert.ToDateTime((Date_Table.Rows[k]).ItemArray[0]);
                for (int i = 0; i < ReadyToClose_Date_Table.Rows.Count; i++)
                {
                    DateTime candidate_date_ReadyToClose = Convert.ToDateTime((ReadyToClose_Date_Table.Rows[i]).ItemArray[0]);
                    if (candidate_date == candidate_date_ReadyToClose)
                    {
                        New_ReadyToClose_candidate_date++;
                    }

                }
                if (New_ReadyToClose_candidate_date == 0)
                {
                    for (int j = 0; j < RNC_Table.Rows.Count; j++)
                    {
                        string RNC = Convert.ToString((RNC_Table.Rows[j]).ItemArray[0]);
                        ReadyToCloseTTs_Table_Default.Rows.Add(candidate_date, RNC, 0);
                    }
                }
            }

            SqlBulkCopy objbulk_ReadyToClose = new SqlBulkCopy(connection);
            objbulk_ReadyToClose.DestinationTableName = "Tehran_CC_ReadyToClose";
            objbulk_ReadyToClose.ColumnMappings.Add("Date", "Date");
            objbulk_ReadyToClose.ColumnMappings.Add("RNC", "RNC");
            objbulk_ReadyToClose.ColumnMappings.Add("ReadyToCloseTTs", "ReadyToCloseTTs");
            objbulk_ReadyToClose.WriteToServer(ReadyToCloseTTs_Table_Default);





            for (int k = 0; k < Date_Table.Rows.Count; k++)
            {
                for (int j = 0; j < RNC_Table.Rows.Count; j++)
                {
                    DateTime Date1 = Convert.ToDateTime((Date_Table.Rows[k]).ItemArray[0]);
                    string RNC = Convert.ToString((RNC_Table.Rows[j]).ItemArray[0]);
                    // RemainTTs_Table_Default.Rows.Add(Date1, RNC, 0);
                    //ReadyToCloseTTs_Table_Default.Rows.Add(Date1, RNC, 0);
                    CloseTTs_Table_Default.Rows.Add(Date1, RNC, 0);
                    CloseTTsRemovePark_Table_Default.Rows.Add(Date1, RNC, 0);
                }
            }


            //// Truncate [Tehran_CC_Remain]
            ////string Truncate_Tehran_CC_Remain_Quary = "Truncate table [Tehran_CC_Remain]";
            ////Query_Execution(Truncate_Tehran_CC_Remain_Quary);

            //// Truncate [Tehran_CC_Remain_Detail]
            ////string Truncate_Tehran_CC_Remain_Detail_Quary = "Truncate table [Tehran_CC_Remain_Detail]";
            ////Query_Execution(Truncate_Tehran_CC_Remain_Detail_Quary);

            //// **********************************

            //// Truncate [Tehran_CC_ReadyToClose]
            //string Truncate_Tehran_CC_ReadyToClose_Quary = "Truncate table [Tehran_CC_ReadyToClose]";
            //Query_Execution(Truncate_Tehran_CC_ReadyToClose_Quary);


            //// Truncate [Tehran_CC_ReadyToClose_Details]
            //string Truncate_Tehran_CC_ReadyToClose_Details_Quary = "Truncate table [Tehran_CC_ReadyToClose_Details]";
            //Query_Execution(Truncate_Tehran_CC_ReadyToClose_Details_Quary);

            //// Truncate [Tehran_CC_ReadyToClose_Detail]
            //string Truncate_Tehran_CC_ReadyToClose_Detail_Quary = "Truncate table [Tehran_CC_ReadyToClose_Detail]";
            //Query_Execution(Truncate_Tehran_CC_ReadyToClose_Detail_Quary);


            //// Truncate [ReadyToCloseTTsRemovePark_Detail_Table]
            //string Truncate_Tehran_CC_ReadyToCloseRemovePark_Detail_Quary = "Truncate table [Tehran_CC_ReadyToCloseRemovePark_Detail]";
            //Query_Execution(Truncate_Tehran_CC_ReadyToCloseRemovePark_Detail_Quary);



            ////// **********************************

            // Truncate [Tehran_CC_Close]
            string Truncate_Tehran_CC_Close_Quary = "Truncate table [Tehran_CC_Close]";
            Query_Execution(Truncate_Tehran_CC_Close_Quary);

            // Truncate [Tehran_CC_CloseRemovePark]
            string Truncate_Tehran_CC_CloseRemovePark_Quary = "Truncate table [Tehran_CC_CloseRemovePark]";
            Query_Execution(Truncate_Tehran_CC_CloseRemovePark_Quary);


            // Truncate [Tehran_CC_Close_Detail]
            string Truncate_Tehran_CC_Close_Detail_Quary = "Truncate table [Tehran_CC_Close_Detail]";
            Query_Execution(Truncate_Tehran_CC_Close_Detail_Quary);


            // Truncate [Tehran_CC_CloseNIExcluded_Detail]
            string Truncate_Tehran_CC_CloseNIExcluded_Detail_Quary = "Truncate table [Tehran_CC_CloseNIExcluded_Detail]";
            Query_Execution(Truncate_Tehran_CC_CloseNIExcluded_Detail_Quary);



            // Insert Zero Values to Defulat Tables

            SqlBulkCopy objbulk_Close = new SqlBulkCopy(connection);
            objbulk_Close.DestinationTableName = "Tehran_CC_Close";
            objbulk_Close.ColumnMappings.Add("Date", "Date");
            objbulk_Close.ColumnMappings.Add("RNC", "RNC");
            objbulk_Close.ColumnMappings.Add("CloseTTs", "CloseTTs");
            objbulk_Close.WriteToServer(CloseTTs_Table_Default);


            SqlBulkCopy objbulk_CloseRemovePark = new SqlBulkCopy(connection);
            objbulk_CloseRemovePark.DestinationTableName = "Tehran_CC_CloseRemovePark";
            objbulk_CloseRemovePark.ColumnMappings.Add("Date", "Date");
            objbulk_CloseRemovePark.ColumnMappings.Add("RNC", "RNC");
            objbulk_CloseRemovePark.ColumnMappings.Add("CloseTTs", "CloseTTs");
            objbulk_CloseRemovePark.WriteToServer(CloseTTsRemovePark_Table_Default);


            // Remain (Incoming) Tickets
            for (int i = 0; i < Date_Table.Rows.Count; i++)
            {
                DateTime Date1 = Convert.ToDateTime((Date_Table.Rows[i]).ItemArray[0]);
                string Date = Date_ToString(Date1);

                if (Date1 <= Last_Updated_Date_Incoming)
                {
                    continue;
                }

                string RemainTTs_Quary = @"select Date, RNC, count(*) as 'NewTTs' from(
select[TT Code], Date, RNC from(select New_TBL.[TT Code], New_TBL.[RNC], New_TBL.Date, Old_TBL.OldTT from(select[TT Code], RNC, Date from Tehran_CC_Maintable where cast(Date as Date) ='" +
  Date + @"') as New_TBL
  left join
  (select distinct[TT Code] as 'OldTT' from Tehran_CC_Maintable where cast(Date as Date) <'" + Date + @"')  Old_TBL
  on New_TBL.[TT Code] = Old_TBL.OldTT) tble where OldTT is null) tble group by Date, RNC";

                DataTable RemainTTs_Table = Query_Execution_Table_Output(RemainTTs_Quary);

                SqlBulkCopy objbulk_Remain1 = new SqlBulkCopy(connection);
                objbulk_Remain1.DestinationTableName = "Tehran_CC_Remain";
                objbulk_Remain1.ColumnMappings.Add("Date", "Date");
                objbulk_Remain1.ColumnMappings.Add("RNC", "RNC");
                objbulk_Remain1.ColumnMappings.Add("NewTTs", "NewTTs");
                objbulk_Remain1.WriteToServer(RemainTTs_Table);



                string RemainTTs_Detail_Quary = @"select[TT Code], Date, RNC, channel, 'Incoming TTs' as 'Status' from(
select New_TBL.[TT Code], New_TBL.[RNC], New_TBL.[channel],  New_TBL.Date, Old_TBL.OldTT from(select[TT Code], RNC, channel,   Date from Tehran_CC_Maintable where cast(Date as Date) = '" +
  Date + @"') as New_TBL
left join
(select distinct[TT Code] as 'OldTT' from Tehran_CC_Maintable where cast(Date as Date) < '" + Date + @"')  Old_TBL
on New_TBL.[TT Code] = Old_TBL.OldTT) tble where OldTT is null";



                DataTable RemainTTs_Detail_Table = Query_Execution_Table_Output(RemainTTs_Detail_Quary);

                SqlBulkCopy objbulk_Remain2 = new SqlBulkCopy(connection);
                objbulk_Remain2.DestinationTableName = "Tehran_CC_Remain_Detail";
                objbulk_Remain2.ColumnMappings.Add("TT code", "TT code");
                objbulk_Remain2.ColumnMappings.Add("Date", "Date");
                objbulk_Remain2.ColumnMappings.Add("RNC", "RNC");
                objbulk_Remain2.ColumnMappings.Add("channel", "channel");
                objbulk_Remain2.ColumnMappings.Add("Status", "Status");
                objbulk_Remain2.WriteToServer(RemainTTs_Detail_Table);
            }



            // Ready to Close Tickets
            for (int i = 0; i < Date_Table.Rows.Count - 1; i++)
            {
                DateTime Date1 = Convert.ToDateTime((Date_Table.Rows[i]).ItemArray[0]);
                string Date = Date_ToString(Date1);
                DateTime Date2 = Convert.ToDateTime((Date_Table.Rows[i + 1]).ItemArray[0]);
                string Date_After = Date_ToString(Date2);


                if (Date2 <= Last_Updated_Date_ReadyToClose)
                {
                    continue;
                }


                string ReadyToCloseTTs_Quary = @"select '" + Date_After + @"' as Date, RNC, count(*) as 'ReadyToCloseTTs' from (
select [TT Code], Date, RNC from(
select Old_TBL.[TT Code], Old_TBL.[RNC],  Old_TBL.Date, New_TBL.NewTT from (select [TT Code], RNC, Date from Tehran_CC_Maintable where cast(Date as Date)='" + Date + @"') as Old_TBL
  left join  
  (select  [TT Code] as 'NewTT'  from Tehran_CC_Maintable where cast(Date as Date)='" + Date_After + @"')  New_TBL
  on Old_TBL.[TT Code]=New_TBL.NewTT)  tble where NewTT is null) tble group by Date , RNC";

                DataTable ReadyToCloseTTs_Table = Query_Execution_Table_Output(ReadyToCloseTTs_Quary);

                SqlBulkCopy objbulk_ReadyToClose1 = new SqlBulkCopy(connection);
                objbulk_ReadyToClose1.DestinationTableName = "Tehran_CC_ReadyToClose";
                objbulk_ReadyToClose1.ColumnMappings.Add("Date", "Date");
                objbulk_ReadyToClose1.ColumnMappings.Add("RNC", "RNC");
                objbulk_ReadyToClose1.ColumnMappings.Add("ReadyToCloseTTs", "ReadyToCloseTTs");
                objbulk_ReadyToClose1.WriteToServer(ReadyToCloseTTs_Table);


                // Ready To Close Datails Table (Table between two charts)
                string ReadyToCloseTTs_Details_Quary = @"select  [TT Code], cast('" + Date_After + @"' as Date) as Date, RNC,  channel, 'Ready To Close TTs' as 'Status'   from(
select Old_TBL.[TT Code], Old_TBL.[RNC], Old_TBL.Date, Old_TBL.channel, New_TBL.NewTT from(select[TT Code], RNC, channel, Date from Tehran_CC_Maintable where cast(Date as Date) = '" + Date + @"') as Old_TBL
 left join
 (select[TT Code] as 'NewTT'  from Tehran_CC_Maintable where cast(Date as Date) = '" + Date_After + @"')  New_TBL
 on Old_TBL.[TT Code] = New_TBL.NewTT)  tble where  NewTT is null";




                DataTable ReadyToCloseTTs_Details_Table = Query_Execution_Table_Output(ReadyToCloseTTs_Details_Quary);

                SqlBulkCopy objbulk_ReadyToClose1_Details = new SqlBulkCopy(connection);
                objbulk_ReadyToClose1_Details.DestinationTableName = "Tehran_CC_ReadyToClose_Details";
                objbulk_ReadyToClose1_Details.ColumnMappings.Add("TT Code", "TT code");
                objbulk_ReadyToClose1_Details.ColumnMappings.Add("Date", "Date");
                objbulk_ReadyToClose1_Details.ColumnMappings.Add("RNC", "RNC");
                objbulk_ReadyToClose1_Details.ColumnMappings.Add("channel", "channel");
                objbulk_ReadyToClose1_Details.ColumnMappings.Add("Status", "Status");
                objbulk_ReadyToClose1_Details.WriteToServer(ReadyToCloseTTs_Details_Table);




                // Ready To Close Datail Table
                string ReadyToCloseTTs_Detail_Quary = @"select cast('" + Date_After + @"' as Date) as Date, RNC, [TT Code] as 'ReadyToCloseTTs'  from(
select Old_TBL.[TT Code], Old_TBL.[RNC],  Old_TBL.Date, New_TBL.NewTT from (select [TT Code], RNC, Date from Tehran_CC_Maintable where cast(Date as Date)='" + Date + @"' ) as Old_TBL
  left join  
  (select  [TT Code] as 'NewTT'  from Tehran_CC_Maintable where cast(Date as Date)= '" + Date_After + @"' )  New_TBL
  on Old_TBL.[TT Code]=New_TBL.NewTT)  tble  where  NewTT is null"; 




                DataTable ReadyToCloseTTs_Detail_Table = Query_Execution_Table_Output(ReadyToCloseTTs_Detail_Quary);

                SqlBulkCopy objbulk_ReadyToClose1_Detail = new SqlBulkCopy(connection);
                objbulk_ReadyToClose1_Detail.DestinationTableName = "Tehran_CC_ReadyToClose_Detail";
                objbulk_ReadyToClose1_Detail.ColumnMappings.Add("Date", "Date");
                objbulk_ReadyToClose1_Detail.ColumnMappings.Add("RNC", "RNC");
                objbulk_ReadyToClose1_Detail.ColumnMappings.Add("ReadyToCloseTTs", "ReadyToCloseTTs");
                objbulk_ReadyToClose1_Detail.WriteToServer(ReadyToCloseTTs_Detail_Table);





               // Ready To Close_Remove Park Datail Table
                string ReadyToCloseTTsRemovePark_Detail_Quary = @"select cast('" + Date_After + @"' as Date) as Date, RNC, [TT Code] as 'ReadyToCloseTTs' from(
select Old_TBL.[TT Code], Old_TBL.[RNC], Old_TBL.Date, New_TBL.NewTT from(select[TT Code], RNC, Date from Tehran_CC_Maintable where cast(Date as Date) = '" + Date + @"' and (cast([Ticket Status] as varchar)!='پارک' and cast([Ticket Status] as varchar)!='پارک سررسید گذشته')) as Old_TBL
 left join
 (select[TT Code] as 'NewTT'  from Tehran_CC_Maintable where cast(Date as Date) = '" + Date_After + @"' and (cast([Ticket Status] as varchar)!='پارک' and cast([Ticket Status] as varchar)!='پارک سررسید گذشته'))  New_TBL
 on Old_TBL.[TT Code] = New_TBL.NewTT)  tble where  NewTT is null";

                DataTable ReadyToCloseTTsRemovePark_Detail_Table = Query_Execution_Table_Output(ReadyToCloseTTsRemovePark_Detail_Quary);

                SqlBulkCopy objbulk_ReadyToCloseRemovePark_Detail = new SqlBulkCopy(connection);
                objbulk_ReadyToCloseRemovePark_Detail.DestinationTableName = "Tehran_CC_ReadyToCloseRemovePark_Detail";
                objbulk_ReadyToCloseRemovePark_Detail.ColumnMappings.Add("Date", "Date");
                objbulk_ReadyToCloseRemovePark_Detail.ColumnMappings.Add("RNC", "RNC");
                objbulk_ReadyToCloseRemovePark_Detail.ColumnMappings.Add("ReadyToCloseTTs", "ReadyToCloseTTs");
                objbulk_ReadyToCloseRemovePark_Detail.WriteToServer(ReadyToCloseTTsRemovePark_Detail_Table);


            }


            // Close Tickets
            for (int i = 0; i < Date_Table.Rows.Count - 1; i++)
            {
                DateTime Date1 = Convert.ToDateTime((Date_Table.Rows[i]).ItemArray[0]);
                string Date = Date_ToString(Date1);

                // Close TTs
                string CloseTTs_Quary = @"select '" + Date + @"' as Date, RNC, count(*) as 'CloseTTs' from (
select [ReadyToCloseTTs], Date, RNC from(
select Old_TBL.[ReadyToCloseTTs], Old_TBL.[RNC],  Old_TBL.Date, New_TBL.NewTT from (select ReadyToCloseTTs , RNC, Date from Tehran_CC_ReadyToClose_Detail where cast(Date as Date)='" + Date + @"') as Old_TBL
  left join  
  (select  distinct [TT Code] as 'NewTT'  from Tehran_CC_Maintable where cast(Date as Date)>'" + Date + @"')  New_TBL
  on Old_TBL.[ReadyToCloseTTs]=New_TBL.NewTT)  tble where NewTT is null) tble group by Date , RNC";


                DataTable CloseTTs_Table = Query_Execution_Table_Output(CloseTTs_Quary);


                SqlBulkCopy objbulk_Close1 = new SqlBulkCopy(connection);
                objbulk_Close1.DestinationTableName = "Tehran_CC_Close";
                objbulk_Close1.ColumnMappings.Add("Date", "Date");
                objbulk_Close1.ColumnMappings.Add("RNC", "RNC");
                objbulk_Close1.ColumnMappings.Add("CloseTTs", "CloseTTs");
                objbulk_Close1.WriteToServer(CloseTTs_Table);






                // Close TTs Detail
                string CloseTTs_Detail_Quary = @"select [ReadyToCloseTTs] as 'TT Code', Date, RNC, 'CloseTTs' as 'Status'  from(
select Old_TBL.[ReadyToCloseTTs], Old_TBL.[RNC],  Old_TBL.Date, New_TBL.NewTT from (select ReadyToCloseTTs , RNC, Date from Tehran_CC_ReadyToClose_Detail where cast(Date as Date)='" + Date + @"') as Old_TBL
  left join  
  (select  distinct [TT Code] as 'NewTT'  from Tehran_CC_Maintable where cast(Date as Date)>'" + Date + @"')  New_TBL
  on Old_TBL.[ReadyToCloseTTs]=New_TBL.NewTT)  tble where NewTT is null";


                DataTable CloseTTs_Detail_Table = Query_Execution_Table_Output(CloseTTs_Detail_Quary);


                SqlBulkCopy objbulk_Detail_Close1 = new SqlBulkCopy(connection);
                objbulk_Detail_Close1.DestinationTableName = "Tehran_CC_Close_Detail";
                objbulk_Detail_Close1.ColumnMappings.Add("TT Code", "TT code");
                objbulk_Detail_Close1.ColumnMappings.Add("Date", "Date");
                objbulk_Detail_Close1.ColumnMappings.Add("RNC", "RNC");
                objbulk_Detail_Close1.ColumnMappings.Add("Status", "Status");
                objbulk_Detail_Close1.WriteToServer(CloseTTs_Detail_Table);






                // Close TTs Remove Park
                string CloseTTsRemovePark_Quary = @"select '" + Date + @"' as Date, RNC, count(*) as 'CloseTTs' from (
select [ReadyToCloseTTs], Date, RNC from(
select Old_TBL.[ReadyToCloseTTs], Old_TBL.[RNC],  Old_TBL.Date, New_TBL.NewTT from (select ReadyToCloseTTs , RNC, Date from Tehran_CC_ReadyToCloseRemovePark_Detail where cast(Date as Date)='" + Date + @"') as Old_TBL
  left join  
  (select  distinct [TT Code] as 'NewTT'  from Tehran_CC_Maintable where cast(Date as Date)>'" + Date + @"')  New_TBL
  on Old_TBL.[ReadyToCloseTTs]=New_TBL.NewTT)  tble where NewTT is null) tble group by Date , RNC";


                DataTable CloseTTsRemovePark_Table = Query_Execution_Table_Output(CloseTTsRemovePark_Quary);


                // Cahnge Values bigger than 10 to 10
                for (int c = 0; c < CloseTTsRemovePark_Table.Rows.Count; c++)
                {
                    if (Convert.ToUInt16(CloseTTsRemovePark_Table.Rows[c].ItemArray[2].ToString())>10)
                    {
                        CloseTTsRemovePark_Table.Rows[c][2] = 10;
                    }
                       
                }


                SqlBulkCopy objbulk_Close1RemovePark = new SqlBulkCopy(connection);
                objbulk_Close1RemovePark.DestinationTableName = "Tehran_CC_CloseRemovePark";
                objbulk_Close1RemovePark.ColumnMappings.Add("Date", "Date");
                objbulk_Close1RemovePark.ColumnMappings.Add("RNC", "RNC");
                objbulk_Close1RemovePark.ColumnMappings.Add("CloseTTs", "CloseTTs");
                objbulk_Close1RemovePark.WriteToServer(CloseTTsRemovePark_Table);





                // Close TTs Remove Park Detail
                string CloseTTs_Detail_RemovePark_Quary = @"  select [ReadyToCloseTTs] as 'TT Code', Date, RNC, 'CloseTTs_NI-Excluded' as 'Status' from(
select Old_TBL.[ReadyToCloseTTs], Old_TBL.[RNC],  Old_TBL.Date, New_TBL.NewTT from (select ReadyToCloseTTs , RNC, Date from Tehran_CC_ReadyToCloseRemovePark_Detail where cast(Date as Date)='" + Date + @"') as Old_TBL
  left join  
  (select  distinct [TT Code] as 'NewTT'  from Tehran_CC_Maintable where cast(Date as Date)>'" + Date + @"')  New_TBL
  on Old_TBL.[ReadyToCloseTTs]=New_TBL.NewTT)  tble where NewTT is null";


                DataTable CloseTTs_Detail_RemovePark_Table = Query_Execution_Table_Output(CloseTTs_Detail_RemovePark_Quary);


                SqlBulkCopy objbulk_Detail_RemovePark_Close1 = new SqlBulkCopy(connection);
                objbulk_Detail_RemovePark_Close1.DestinationTableName = "Tehran_CC_CloseNIExcluded_Detail";
                objbulk_Detail_RemovePark_Close1.ColumnMappings.Add("TT Code", "TT code");
                objbulk_Detail_RemovePark_Close1.ColumnMappings.Add("Date", "Date");
                objbulk_Detail_RemovePark_Close1.ColumnMappings.Add("RNC", "RNC");
                objbulk_Detail_RemovePark_Close1.ColumnMappings.Add("Status", "Status");
                objbulk_Detail_RemovePark_Close1.WriteToServer(CloseTTs_Detail_RemovePark_Table);




            }



            string CloseTTs_Quary_LastDate = @"delete from[Tehran_CC_Close] where cast(Date as Date)= '" + Last_Updated_Date_String +
                "' insert into[Tehran_CC_Close] select* from(select Date, RNC, ReadyToCloseTTs as 'ClosetTTS' from[Tehran_CC_ReadyToClose] where cast(Date as Date)= '" + Last_Updated_Date_String + "') as tble";
            Query_Execution(CloseTTs_Quary_LastDate);



            string CloseTTsRemovePark_Quary_LastDate = @"delete from[Tehran_CC_CloseRemovePark] where cast(Date as Date)= '" + Last_Updated_Date_String +
                "' insert into [Tehran_CC_CloseRemovePark] select * from(select Date, RNC, ReadyToCloseTTs as 'ClosetTTS' from [Tehran_CC_ReadyToClose] where cast(Date as Date)= '" + Last_Updated_Date_String + "') as tble";
            Query_Execution(CloseTTsRemovePark_Quary_LastDate);



            string Delete_Quary = @"delete from[Tehran_CC_Remain] where cast(Date as Date)= '2022-07-31' delete from[Tehran_CC_ReadyToClose] where cast(Date as Date)= '2022-07-31' delete from[Tehran_CC_Close] where cast(Date as Date)= '2022-07-31'";
            Query_Execution(Delete_Quary);

            MessageBox.Show("Finished");

        }

        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog1.DefaultExt = "xlsx";
            openFileDialog1.Filter = "Excel File|*.xlsx";
            DialogResult result = openFileDialog1.ShowDialog();
            string File_Name = openFileDialog1.SafeFileName.ToString();
            string file = openFileDialog1.FileName;
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(file);
            Sheet = xlWorkBook.Worksheets[1];


            Server_Name = "PERFORMANCEDB01";
            DataBase_Name = "Performance_NAK";


            ConnectionString = @"Server=" + Server_Name + "; Database=" + DataBase_Name + "; User ID=cwpcApp;Password=cwpcApp@830625#Ahmad";
            connection = new SqlConnection(ConnectionString);
            connection.Open();


            //Truncate[Tehran_VIP_CC]
            string Truncate_Tehran_VIP_CC_Quary = "Truncate table [Tehran_VIP_CC]";
            Query_Execution(Truncate_Tehran_VIP_CC_Quary);



            DataTable Tehran_VIP_CC = new DataTable();
            Tehran_VIP_CC.Columns.Add("CCID", typeof(String));
            Tehran_VIP_CC.Columns.Add("Subscriber Name", typeof(String));
            Tehran_VIP_CC.Columns.Add("Phone", typeof(String));
            Tehran_VIP_CC.Columns.Add("Address", typeof(String));
            Tehran_VIP_CC.Columns.Add("latitude", typeof(double));
            Tehran_VIP_CC.Columns.Add("longitude", typeof(double));
            Tehran_VIP_CC.Columns.Add("Source Of Complaint", typeof(String));
            Tehran_VIP_CC.Columns.Add("RNC", typeof(String));
            Tehran_VIP_CC.Columns.Add("Status", typeof(String));
            Tehran_VIP_CC.Columns.Add("Created Da", typeof(DateTime));
            Tehran_VIP_CC.Columns.Add("comments", typeof(String));
            Tehran_VIP_CC.Columns.Add("SOLUTIONCATEGORY", typeof(String));
            Tehran_VIP_CC.Columns.Add("Pending Site", typeof(String));
            Tehran_VIP_CC.Columns.Add("RESPONSIBLE", typeof(String));



            Excel.Range Data = Sheet.get_Range("A2", "N" + Sheet.UsedRange.Rows.Count);
            object[,] VIP_Data = (object[,])Data.Value;
            int Count = Sheet.UsedRange.Rows.Count;


            for (int k = 0; k < Count - 1; k++)
            {

                if (VIP_Data[k + 1, 1] == null)
                {
                    continue;
                }

                DateTime Date = Convert.ToDateTime(VIP_Data[k + 1, 10]);
                if (Date == null || Date.Year<2022)
                {
                    continue;
                }

               string CCID = "";
               if (VIP_Data[k + 1, 1] != null)
               {
                   CCID = VIP_Data[k + 1, 1].ToString();
               }


                string Subscriber_Name = "";
                if (VIP_Data[k + 1, 2] != null)
                {
                    Subscriber_Name = VIP_Data[k + 1, 2].ToString();
                }

                string Phone = "";
                if (VIP_Data[k + 1, 3] != null)
                {
                    Phone = VIP_Data[k + 1, 3].ToString();
                }


                string Address = "";
                if (VIP_Data[k + 1, 4] != null)
                {
                    Address = VIP_Data[k + 1, 4].ToString();
                }


                double latitude = 0;
                if (VIP_Data[k + 1, 5] != null)
                {
                    latitude = Convert.ToDouble(VIP_Data[k + 1, 5]);
                }

                double longitude = 0;
                if (VIP_Data[k + 1, 6] != null)
                {
                    longitude = Convert.ToDouble(VIP_Data[k + 1, 6]);
                }



                string Source_Of_Complaint = "";
                if (VIP_Data[k + 1, 7] != null)
                {
                    Source_Of_Complaint = VIP_Data[k + 1, 7].ToString();
                }


                string RNC = "";
                if (VIP_Data[k + 1, 8] != null)
                {
                    RNC = VIP_Data[k + 1, 8].ToString();
                }


                string Status = "";
                if (VIP_Data[k + 1, 9] != null)
                {
                    Status = VIP_Data[k + 1, 9].ToString();
                }



                string comments = "";
                if (VIP_Data[k + 1, 11] != null)
                {
                    comments = VIP_Data[k + 1, 11].ToString();
                }



                string SOLUTIONCATEGORY = "";
                if (VIP_Data[k + 1, 12] != null)
                {
                    SOLUTIONCATEGORY = VIP_Data[k + 1, 12].ToString();
                }


                string Pending_Site = "";
                if (VIP_Data[k + 1, 13] != null)
                {
                    Pending_Site = VIP_Data[k + 1, 13].ToString();
                }


                string RESPONSIBLE = "";
                if (VIP_Data[k + 1, 14] != null)
                {
                    RESPONSIBLE = VIP_Data[k + 1, 14].ToString();
                }




                Tehran_VIP_CC.Rows.Add(CCID, Subscriber_Name, Phone,Address, latitude, longitude, Source_Of_Complaint,RNC,Status,Date, comments, SOLUTIONCATEGORY, Pending_Site, RESPONSIBLE);


            }



            SqlBulkCopy objbulk_Tehran_VIP_CC = new SqlBulkCopy(connection);
            objbulk_Tehran_VIP_CC.DestinationTableName = "Tehran_VIP_CC";
            objbulk_Tehran_VIP_CC.ColumnMappings.Add("CCID", "CCID");
            objbulk_Tehran_VIP_CC.ColumnMappings.Add("Subscriber Name", "Subscriber Name");
            objbulk_Tehran_VIP_CC.ColumnMappings.Add("Phone", "Phone");
            objbulk_Tehran_VIP_CC.ColumnMappings.Add("Address", "Address");

            objbulk_Tehran_VIP_CC.ColumnMappings.Add("latitude", "latitude");
            objbulk_Tehran_VIP_CC.ColumnMappings.Add("longitude", "longitude");
            objbulk_Tehran_VIP_CC.ColumnMappings.Add("Source Of Complaint", "Source Of Complaint");
            objbulk_Tehran_VIP_CC.ColumnMappings.Add("RNC", "RNC");

            objbulk_Tehran_VIP_CC.ColumnMappings.Add("Status", "Status");
            objbulk_Tehran_VIP_CC.ColumnMappings.Add("Created Da", "Created Da");
            objbulk_Tehran_VIP_CC.ColumnMappings.Add("comments", "comments");
            objbulk_Tehran_VIP_CC.ColumnMappings.Add("SOLUTIONCATEGORY", "SOLUTIONCATEGORY");

            objbulk_Tehran_VIP_CC.ColumnMappings.Add("Pending Site", "Pending Site");
            objbulk_Tehran_VIP_CC.ColumnMappings.Add("RESPONSIBLE", "RESPONSIBLE");


            objbulk_Tehran_VIP_CC.WriteToServer(Tehran_VIP_CC);

            connection.Close();
            MessageBox.Show("Finished");











        }
    }
}
