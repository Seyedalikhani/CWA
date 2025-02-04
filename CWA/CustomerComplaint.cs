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
    public partial class CustomerComplaint : Form
    {
        public CustomerComplaint()
        {
            InitializeComponent();
        }


        public Main form1;


        public CustomerComplaint(Form form)
        {
            InitializeComponent();
            form1 = (Main)form;
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

        // Method of Query Execution without Output
        void Query_Execution(String Query)
        {
            string Quary_String = Query;
            SqlCommand Quary_Command = new SqlCommand(Quary_String, connection);
            Quary_Command.CommandTimeout = 0;
            Quary_Command.ExecuteNonQuery();
        }

        private void button1_Click(object sender, EventArgs e)
        {


            label1.Text = "Please wait while It is running";
            label1.BackColor = Color.GreenYellow;
            Thread th1 = new Thread(my_thread1);
            th1.Start();

        }



        void my_thread1()
        {


            Server_Name = "PERFORMANCEDB";
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
            string Last_Updated_Date_String_90 = Date_ToString(Last_Updated_Date.AddDays(-90));


            // Date Table 
            string Date_Quary = @"Select Distinct Date from Tehran_CC_Maintable where cast(Date as Date)>='2022-07-31' order by Date";                        // From mBegining
            //string Date_Quary = @"Select Distinct Date from Tehran_CC_Maintable where cast(Date as Date)>'"+ Last_Updated_Date_String+"' order by Date";    // From Last Date
            DataTable Date_Table = Query_Execution_Table_Output(Date_Quary);



            // Last Date of Tehran_CC_NPOMTTR
            string Max_Date_Quary_Tehran_CC_NPOMTTR = @"Select max(Date) from Tehran_CC_NPOMTTR";
            DataTable Max_Date_Table_Tehran_CC_NPOMTTR = Query_Execution_Table_Output(Max_Date_Quary_Tehran_CC_NPOMTTR);
            // Convert Last Date to String
            DateTime Last_Updated_Date_Tehran_CC_NPOMTTR = Convert.ToDateTime((Max_Date_Table_Tehran_CC_NPOMTTR.Rows[Max_Date_Table_Tehran_CC_NPOMTTR.Rows.Count - 1]).ItemArray[0]);


            // Last Date of Tehran_CC_NPOMTTR_90
            //string Max_Date_Quary_Tehran_CC_NPOMTTR_90 = @"Select cast(max(Date) as datetime)-90 from Tehran_CC_NPOMTTR_90";
            string Max_Date_Quary_Tehran_CC_NPOMTTR_90 = @"Select cast(max(Date) as datetime)-90 from Tehran_CC_NPOMTTR";
            DataTable Max_Date_Table_Tehran_CC_NPOMTTR_90 = Query_Execution_Table_Output(Max_Date_Quary_Tehran_CC_NPOMTTR_90);
            // Convert Last Date to String
            DateTime Last_Updated_Date_Tehran_CC_NPOMTTR_90 = Convert.ToDateTime((Max_Date_Table_Tehran_CC_NPOMTTR_90.Rows[Max_Date_Table_Tehran_CC_NPOMTTR_90.Rows.Count - 1]).ItemArray[0]);



            // Last Date of Tehran_CC_FOMTTR
            string Max_Date_Quary_Tehran_CC_FOMTTR = @"Select max(Date) from Tehran_CC_FOMTTR";
            DataTable Max_Date_Table_Tehran_CC_FOMTTR = Query_Execution_Table_Output(Max_Date_Quary_Tehran_CC_FOMTTR);
            // Convert Last Date to String
            DateTime Last_Updated_Date_Tehran_CC_FOMTTR = Convert.ToDateTime((Max_Date_Table_Tehran_CC_FOMTTR.Rows[Max_Date_Table_Tehran_CC_FOMTTR.Rows.Count - 1]).ItemArray[0]);


            //Last Date of Tehran_CC_FOMTTR_90
            //string Max_Date_Quary_Tehran_CC_FOMTTR_90 = @"Select cast(max(Date) as datetime)-90 from Tehran_CC_FOMTTR_90";
            string Max_Date_Quary_Tehran_CC_FOMTTR_90 = @"Select cast(max(Date) as datetime)-90 from Tehran_CC_FOMTTR";
            DataTable Max_Date_Table_Tehran_CC_FOMTTR_90 = Query_Execution_Table_Output(Max_Date_Quary_Tehran_CC_FOMTTR_90);
            // Convert Last Date to String
            DateTime Last_Updated_Date_Tehran_CC_FOMTTR_90 = Convert.ToDateTime((Max_Date_Table_Tehran_CC_FOMTTR_90.Rows[Max_Date_Table_Tehran_CC_FOMTTR_90.Rows.Count - 1]).ItemArray[0]);


            // Last Date of Tehran_CC_Open_NotBlong_ToCQ
            string Max_Date_Quary_Tehran_CC_Open_NotBlong_ToCQ = @"Select max(Date) from Tehran_CC_Open_NotBlong_ToCQ";
            DataTable Max_Date_Table_Tehran_CC_Open_NotBlong_ToCQ = Query_Execution_Table_Output(Max_Date_Quary_Tehran_CC_Open_NotBlong_ToCQ);
            // Convert Last Date to String
            DateTime Last_Updated_Date_Tehran_CC_Open_NotBlong_ToCQ = Convert.ToDateTime((Max_Date_Table_Tehran_CC_Open_NotBlong_ToCQ.Rows[Max_Date_Table_Tehran_CC_Open_NotBlong_ToCQ.Rows.Count - 1]).ItemArray[0]);




            Query_Execution("truncate table Tehran_CC_NPOMTTR_90");
            Query_Execution("truncate table Tehran_CC_FOMTTR_90");

            // RNC Table
            string RNC_Quary = @"Select Distinct RNC from Tehran_CC_Maintable where substring(RNC,1,1)='R'";
            DataTable RNC_Table = Query_Execution_Table_Output(RNC_Quary);



            // TT Table
            string TT_Quary = @"Select Distinct [TT Code] from Tehran_CC_Maintable where cast(Date as Date)>='2022-07-31'";                        // From mBegining
            DataTable TT_Table = Query_Execution_Table_Output(TT_Quary);



            // Date of MTTR
            string Incoming_Date_Quary = @"Select Distinct Date from Tehran_CC_NPOMTTR where cast(Date as Date)>='2022-07-31' order by Date";
            DataTable Incoming_Date_Table = Query_Execution_Table_Output(Incoming_Date_Quary);
            DateTime Last_Updated_Date_Incoming = Convert.ToDateTime((Incoming_Date_Table.Rows[Incoming_Date_Table.Rows.Count - 1]).ItemArray[0]);



            // Date of QDate
            string QDate_Date_Quary = @"Select Distinct Date from Tehran_CC_Open_NotBlong_ToCQ where cast(Date as Date)>='2022-07-31' order by Date";
            DataTable QDate_Date_Table = Query_Execution_Table_Output(QDate_Date_Quary);
            DateTime Last_QDate_Date_Table = Convert.ToDateTime((QDate_Date_Table.Rows[QDate_Date_Table.Rows.Count - 1]).ItemArray[0]);


            //NPO MTTR of Tickets
            for (int i = 0; i < Date_Table.Rows.Count; i++)
            {
                DateTime Date1 = Convert.ToDateTime((Date_Table.Rows[i]).ItemArray[0]);
                string Date = Date_ToString(Date1);

                if (Date1 <= Last_Updated_Date_Tehran_CC_NPOMTTR)
                {
                    continue;
                }

                string MTTR_Quary = @"select MTTR90.Date, MTTR90.[TT Code], MTTR90.GeneratedDate, MTTR90.RNC, Tehran_CC_RNC_Owner.[RNC Owner], MTTR90.[Count of TT] from (


                                    select '" +
                                      Date + @"' as Date, TBL1.[TT Code],  TBL1.GeneratedDate, TBL1.RNC, TBL2.[Count of TT] from (

                                    select[TT Code], Date , GeneratedDate, RNC from(

                                    select 
                                      RANK() OVER(PARTITION BY  [TT Code] ORDER BY Date desc) AS ranking,
                                      [TT Code],
                                      GeneratedDate,
                                      RNC,
                                      Date from (

                                    select [TT Code], Date , GeneratedDate, RNC, DATEDIFF(day, Date,GeneratedDate) as DateDiff from (
                                    select [TT Code], Date , RNC,
                                    TRY_CAST(CAST(CAST(SUBSTRING(CONVERT (VARCHAR(50), [TT Code], 128), 1, 9)*1e7 AS INT) AS VARCHAR(8)) AS DATE) as 'GeneratedDate' 
                                    from Tehran_CC_Maintable where 
                                    (cast([Ticket Status] as varchar)='باز داخل باکس شخصی' or cast([Ticket Status] as varchar)='باز داخل باکس گروهی' ) and 
                                    (cast([The Last Agent Name] as varchar)='Delayناک تهران' or [The Last Agent Name]='Mohamadreza Kazmi')  )tbl 

                                    )tbl ) tbl where ranking=1
                                    ) TBL1

                                    left join
                                    (
                                    select [TT Code],  count(*) 'Count of TT' from Tehran_CC_Maintable where 
                                    (cast([Ticket Status] as varchar)='باز داخل باکس شخصی' or cast([Ticket Status] as varchar)='باز داخل باکس گروهی' ) and 
                                    (cast([The Last Agent Name] as varchar)='Delayناک تهران' or [The Last Agent Name]='Mohamadreza Kazmi')
                                    group by [TT Code]) TBL2

                                    on

                                    TBL1.[TT Code]=TBL2.[TT Code]

                                    ) MTTR90 left join Tehran_CC_RNC_Owner

                                    on
                                    MTTR90.RNC=Tehran_CC_RNC_Owner.[RNC Name]";




                DataTable MTTR_Table = Query_Execution_Table_Output(MTTR_Quary);

                SqlBulkCopy objbulk_MTTR = new SqlBulkCopy(connection);
                objbulk_MTTR.DestinationTableName = "Tehran_CC_NPOMTTR";
                objbulk_MTTR.ColumnMappings.Add("Date", "Date");
                objbulk_MTTR.ColumnMappings.Add("TT Code", "[TT Code]");
                objbulk_MTTR.ColumnMappings.Add("GeneratedDate", "GeneratedDate");
                objbulk_MTTR.ColumnMappings.Add("RNC", "RNC");
                objbulk_MTTR.ColumnMappings.Add("RNC Owner", "[RNC Owner]");
                objbulk_MTTR.ColumnMappings.Add("Count of TT", "[Count of TT]");
                objbulk_MTTR.WriteToServer(MTTR_Table);


            }


            //checkBox1.Checked = true;
            checkBox1.Invoke(new Action(() => checkBox1.Checked = true));

            //NPO MTTR of Tickets Last 90 days
            for (int i = 0; i < Date_Table.Rows.Count; i++)
            {
                DateTime Date1 = Convert.ToDateTime((Date_Table.Rows[i]).ItemArray[0]);
                string Date = Date_ToString(Date1);

                if (Date1 <= Last_Updated_Date_Tehran_CC_NPOMTTR_90)
                {
                    continue;
                }



                string MTTR_Quary = @"select MTTR90.Date, MTTR90.[TT Code], MTTR90.GeneratedDate, MTTR90.RNC, Tehran_CC_RNC_Owner.[RNC Owner], MTTR90.[Count of TT] from (

                                    select '" +
                                     Date + @"' as Date, TBL1.[TT Code],  TBL1.GeneratedDate, TBL1.RNC, TBL2.[Count of TT] from (

                                    select[TT Code], Date , GeneratedDate, RNC from(

                                    select 
                                      RANK() OVER(PARTITION BY  [TT Code] ORDER BY Date desc) AS ranking,
                                      [TT Code],
                                      GeneratedDate,
                                      RNC,
                                      Date from (

                                    select [TT Code], Date , GeneratedDate, RNC, DATEDIFF(day, Date,GeneratedDate) as DateDiff from (
                                    select [TT Code], Date , RNC,
                                    TRY_CAST(CAST(CAST(SUBSTRING(CONVERT (VARCHAR(50), [TT Code], 128), 1, 9)*1e7 AS INT) AS VARCHAR(8)) AS DATE) as 'GeneratedDate' 
                                    from Tehran_CC_Maintable where 
                                    (cast([Ticket Status] as varchar)='باز داخل باکس شخصی' or cast([Ticket Status] as varchar)='باز داخل باکس گروهی' ) and 
                                    (cast([The Last Agent Name] as varchar)='Delayناک تهران' or [The Last Agent Name]='Mohamadreza Kazmi')  )tbl where DATEDIFF(day, '" + Date + @"',GeneratedDate)>=-90 

                                    )tbl ) tbl where ranking=1
                                    ) TBL1

                                    left join
                                    (
                                    select [TT Code],  count(*) 'Count of TT' from Tehran_CC_Maintable where 
                                    (cast([Ticket Status] as varchar)='باز داخل باکس شخصی' or cast([Ticket Status] as varchar)='باز داخل باکس گروهی' ) and 
                                    (cast([The Last Agent Name] as varchar)='Delayناک تهران' or [The Last Agent Name]='Mohamadreza Kazmi')
                                    group by [TT Code]) TBL2

                                    on

                                    TBL1.[TT Code]=TBL2.[TT Code]

                                    ) MTTR90 left join Tehran_CC_RNC_Owner

                                    on
                                    MTTR90.RNC=Tehran_CC_RNC_Owner.[RNC Name]";




                DataTable MTTR_Table = Query_Execution_Table_Output(MTTR_Quary);

                SqlBulkCopy objbulk_MTTR = new SqlBulkCopy(connection);
                objbulk_MTTR.DestinationTableName = "Tehran_CC_NPOMTTR_90";
                objbulk_MTTR.ColumnMappings.Add("Date", "Date");
                objbulk_MTTR.ColumnMappings.Add("TT Code", "TT Code");
                objbulk_MTTR.ColumnMappings.Add("GeneratedDate", "GeneratedDate");
                objbulk_MTTR.ColumnMappings.Add("RNC", "RNC");
                objbulk_MTTR.ColumnMappings.Add("RNC Owner", "RNC Owner");
                objbulk_MTTR.ColumnMappings.Add("Count of TT", "Count of TT");
                objbulk_MTTR.WriteToServer(MTTR_Table);


            }

            checkBox2.Invoke(new Action(() => checkBox2.Checked = true));

            //NPO MTTR of Tickets Last 90 days  DailyResults

            string Last_GDate_Str = @"select max(GeneratedDate) from Tehran_CC_NPOMTTR_90";
            DataTable Last_GDate_Table = Query_Execution_Table_Output(Last_GDate_Str);
            DateTime Last_GDate = Convert.ToDateTime((Last_GDate_Table.Rows[Last_GDate_Table.Rows.Count - 1]).ItemArray[0]);
            Last_GDate = Last_GDate.AddDays(-90);
            string Last_GDate_String = Date_ToString(Last_GDate);

            string GDate_Str = @"select distinct GeneratedDate from Tehran_CC_NPOMTTR_90 where GeneratedDate>='" + Last_GDate_String + "'";
            DataTable GDate_Table = Query_Execution_Table_Output(GDate_Str);


            Query_Execution("delete from Tehran_CC_NPOMTTR_DailyResults_90 where GeneratedDate>='" + Last_GDate_String + "'");
            Query_Execution("delete from Tehran_CC_NPOMTTR_DailyResultsLess15_90 where GeneratedDate>='" + Last_GDate_String + "'");


            for (int i = 0; i < GDate_Table.Rows.Count; i++)
            {
                DateTime D1 = Convert.ToDateTime((GDate_Table.Rows[i]).ItemArray[0]);
                DateTime D2 = D1.AddDays(-90);
                string D1_String = Date_ToString(D1);
                string D2_String = Date_ToString(D2);

                string GDate_Daily_Str = @"select '" + D1_String + @"' as 'GeneratedDate', [RNC], [RNC Owner], count([TT Code]) as 'Count of TT' from(
               select TBL1.RNC, TBL1.[TT Code], Tehran_CC_RNC_Owner.[RNC Owner] from(
               select distinct[TT Code], [RNC] from Tehran_CC_NPOMTTR_90
               where [GeneratedDate] >= '" + D2_String + "' and[GeneratedDate] <= '" + D1_String + "') TBL1 left join Tehran_CC_RNC_Owner on TBL1.RNC = Tehran_CC_RNC_Owner.[RNC Name]) tbl group by [RNC], [RNC Owner]";



                DataTable GDate_Daily_Table = Query_Execution_Table_Output(GDate_Daily_Str);


                SqlBulkCopy objbulk_GDate = new SqlBulkCopy(connection);
                objbulk_GDate.DestinationTableName = "Tehran_CC_NPOMTTR_DailyResults_90";
                objbulk_GDate.ColumnMappings.Add("GeneratedDate", "GeneratedDate");
                objbulk_GDate.ColumnMappings.Add("RNC", "RNC");
                objbulk_GDate.ColumnMappings.Add("RNC Owner", "RNC Owner");
                objbulk_GDate.ColumnMappings.Add("Count of TT", "Count of TT");
                objbulk_GDate.WriteToServer(GDate_Daily_Table);




                string GDateLess15_Daily_Str = @"select '" + D1_String + @"' as 'GeneratedDate', [RNC], [RNC Owner], count([TT Code]) as 'Count of TT' from(
               select TBL1.RNC, TBL1.[TT Code], Tehran_CC_RNC_Owner.[RNC Owner] from(
               select distinct[TT Code], [RNC] from Tehran_CC_NPOMTTR_90
               where [GeneratedDate] >= '" + D2_String + "' and[GeneratedDate] <= '" + D1_String + "' and [Count of TT]<16) TBL1 left join Tehran_CC_RNC_Owner on TBL1.RNC = Tehran_CC_RNC_Owner.[RNC Name]) tbl group by [RNC], [RNC Owner]";

                DataTable GDateLess15_Daily_Table = Query_Execution_Table_Output(GDateLess15_Daily_Str);


                SqlBulkCopy objbulk_GDateLess15 = new SqlBulkCopy(connection);
                objbulk_GDateLess15.DestinationTableName = "Tehran_CC_NPOMTTR_DailyResultsLess15_90";
                objbulk_GDateLess15.ColumnMappings.Add("GeneratedDate", "GeneratedDate");
                objbulk_GDateLess15.ColumnMappings.Add("RNC", "RNC");
                objbulk_GDateLess15.ColumnMappings.Add("RNC Owner", "RNC Owner");
                objbulk_GDateLess15.ColumnMappings.Add("Count of TT", "Count of TT");
                objbulk_GDateLess15.WriteToServer(GDateLess15_Daily_Table);



            }

            checkBox3.Invoke(new Action(() => checkBox3.Checked = true));


            // delete older than 90 days
            string Last_GDate_Str2 = @"select max(GeneratedDate) from Tehran_CC_NPOMTTR_DailyResults_90";
            DataTable Last_GDate_Table2 = Query_Execution_Table_Output(Last_GDate_Str2);
            DateTime Last_GDate2 = Convert.ToDateTime((Last_GDate_Table2.Rows[Last_GDate_Table2.Rows.Count - 1]).ItemArray[0]);
            Last_GDate2 = Last_GDate2.AddDays(-90);
            string Last_GDate_String2 = Date_ToString(Last_GDate2);

            Query_Execution("delete from Tehran_CC_NPOMTTR_DailyResults_90 where GeneratedDate<'" + Last_GDate_String2 + "'");
            Query_Execution("delete from Tehran_CC_NPOMTTR_DailyResultsLess15_90 where GeneratedDate<'" + Last_GDate_String2 + "'");



            // VIPCC
            string Last_GDateVIP_Str = @"select max(cast(cast([Created] as varchar(10)) as datetime)) from Tehran_VIP_CC";
            DataTable Last_GDateVIP_Table = Query_Execution_Table_Output(Last_GDateVIP_Str);
            DateTime Last_GDateVIP = Convert.ToDateTime((Last_GDateVIP_Table.Rows[Last_GDateVIP_Table.Rows.Count - 1]).ItemArray[0]);
            DateTime First_GDateVIP = Last_GDateVIP;
            string First_GDateVIP_String = Date_ToString(First_GDateVIP);
            Last_GDateVIP = Last_GDateVIP.AddDays(-90);
            string Last_GDateVIP_String = Date_ToString(Last_GDateVIP);

            string GDateVIP_Str = @"select distinct cast(cast([Created] as varchar(10)) as datetime) as 'GeneratedDate' from Tehran_VIP_CC where cast(cast([Created] as varchar(10)) as datetime)>='" + Last_GDateVIP_String + "'";
            DataTable GDateVIP_Table = Query_Execution_Table_Output(GDateVIP_Str);
            //DateTime GDate = Convert.ToDateTime((Last_GDate_Table.Rows[Last_GDate_Table.Rows.Count - 1]).ItemArray[0]);



            string Last_GDateVIP_Str1 = @"select max(Date_of_File) from Tehran_VIP_CC";
            DataTable Last_GDateVIP_Table1 = Query_Execution_Table_Output(Last_GDateVIP_Str1);
            DateTime Last_GDateVIP1 = Convert.ToDateTime((Last_GDateVIP_Table1.Rows[Last_GDateVIP_Table1.Rows.Count - 1]).ItemArray[0]);
            DateTime First_GDateVIP1 = Last_GDateVIP1;
            string First_GDateVIP_String1 = Date_ToString(First_GDateVIP1);


            Query_Execution("truncate table Tehran_VIPCC_NPOMTTR_DailyResults_90");
            Query_Execution("truncate table Tehran_VIPCC_NPOMTTR_DailyResults_90_Less20");

            for (int i = 0; i < GDateVIP_Table.Rows.Count; i++)
            {
                DateTime D1 = Convert.ToDateTime((GDateVIP_Table.Rows[i]).ItemArray[0]);
                DateTime D2 = D1.AddDays(-90);
                string D1_String = Date_ToString(D1);
                string D2_String = Date_ToString(D2);

                string GDateVIP_Daily_Str = @"select '" + D1_String + @"' as 'GeneratedDate', [RNC Name] as 'RNC', [RNC Owner], sum(Count) as 'Count of TT' from(
               select cast(cast([Created] as varchar(10)) as date) as Date, [RNC Name], [RNC Owner], Count(*) as 'Count' from(
               select Tehran_VIP_CC.[Created], [Tehran_VIPCC_RNC_Owner].[RNC Name], [Tehran_VIPCC_RNC_Owner].[RNC Owner] from Tehran_VIP_CC left join [Tehran_VIPCC_RNC_Owner] on
               Tehran_VIP_CC.RNC = [Tehran_VIPCC_RNC_Owner].[RNC Name] where Date_Of_File ='" + First_GDateVIP_String1 + "'and cast(cast([Created] as varchar(10)) as date) >='" + D2_String + "'  and cast(cast([Created] as varchar(10)) as date)<='" + D1_String + "'  and NPOStatus!= 'Rejected'  ) tble group by cast(cast([Created] as varchar(10)) as date), [RNC Name], [RNC Owner]) tble group by[RNC Name],  [RNC Owner]";


                string GDateVIP_Daily_Str_Less20 = @"select  '" + D1_String + @"'  as 'GeneratedDate', [RNC Name] as 'RNC', [RNC Owner], sum(Count) as 'Count of TT' from(
               select cast(cast([Created] as varchar(10)) as date) as Date, [RNC Name], [RNC Owner], Count(*) as 'Count' from(
			   select Tehran_VIP_CC.[Created], [Tehran_VIPCC_RNC_Owner].[RNC Name], [Tehran_VIPCC_RNC_Owner].[RNC Owner], 
			   DATEDIFF(day, cast(cast([Created] as varchar(10)) as date),cast(cast([Modified] as varchar(10)) as date)) as 'Difference of Created and Modified'
			   from Tehran_VIP_CC left join [Tehran_VIPCC_RNC_Owner] on
               Tehran_VIP_CC.RNC = [Tehran_VIPCC_RNC_Owner].[RNC Name] where Date_Of_File ='" + First_GDateVIP_String1 + @"'and cast(cast([Created] as varchar(10)) as date) >='" + D2_String + @"'
               and cast(cast([Created] as varchar(10)) as date)<='" + D1_String + @"' and NPOStatus!= 'Rejected' and NPOStatus!= 'Open' and NPOStatus!= 'Assigned' and NPOStatus!= 'In Progress' and
               DATEDIFF(day, cast(cast([Created] as varchar(10)) as date),cast(cast([Modified] as varchar(10)) as date))<20) tble 
			   group by cast(cast([Created] as varchar(10)) as date), [RNC Name], [RNC Owner]) tble group by[RNC Name],  [RNC Owner]";



                DataTable GDateVIP_Daily_Table = Query_Execution_Table_Output(GDateVIP_Daily_Str);
                DataTable GDateVIP_DailyLess20_Table = Query_Execution_Table_Output(GDateVIP_Daily_Str_Less20);

                SqlBulkCopy objbulk_GDateVIP = new SqlBulkCopy(connection);
                objbulk_GDateVIP.DestinationTableName = "Tehran_VIPCC_NPOMTTR_DailyResults_90";
                objbulk_GDateVIP.ColumnMappings.Add("GeneratedDate", "GeneratedDate");
                objbulk_GDateVIP.ColumnMappings.Add("RNC", "RNC");
                objbulk_GDateVIP.ColumnMappings.Add("RNC Owner", "RNC Owner");
                objbulk_GDateVIP.ColumnMappings.Add("Count of TT", "Count of TT");
                objbulk_GDateVIP.WriteToServer(GDateVIP_Daily_Table);

                SqlBulkCopy objbulk_GDateVIP_Less20 = new SqlBulkCopy(connection);
                objbulk_GDateVIP_Less20.DestinationTableName = "Tehran_VIPCC_NPOMTTR_DailyResults_90_Less20";
                objbulk_GDateVIP_Less20.ColumnMappings.Add("GeneratedDate", "GeneratedDate");
                objbulk_GDateVIP_Less20.ColumnMappings.Add("RNC", "RNC");
                objbulk_GDateVIP_Less20.ColumnMappings.Add("RNC Owner", "RNC Owner");
                objbulk_GDateVIP_Less20.ColumnMappings.Add("Count of TT", "Count of TT");
                objbulk_GDateVIP_Less20.WriteToServer(GDateVIP_DailyLess20_Table);
            }


            checkBox4.Invoke(new Action(() => checkBox4.Checked = true));

            //FO MTTR of Tickets
            for (int i = 0; i < Date_Table.Rows.Count; i++)
            {
                DateTime Date1 = Convert.ToDateTime((Date_Table.Rows[i]).ItemArray[0]);
                string Date = Date_ToString(Date1);

                if (Date1 <= Last_Updated_Date_Tehran_CC_FOMTTR)
                {
                    continue;
                }

                string MTTR_Quary = @"select MTTR90.Date, MTTR90.[TT Code], MTTR90.GeneratedDate, MTTR90.RNC, Tehran_CC_RNC_Owner.[RNC Owner], MTTR90.[Count of TT] from (


                                    select '" +
                                    Date + @"' as Date, TBL1.[TT Code],  TBL1.GeneratedDate, TBL1.RNC, TBL2.[Count of TT] from (

                                    select[TT Code], Date , GeneratedDate, RNC from(

                                    select 
                                      RANK() OVER(PARTITION BY  [TT Code] ORDER BY Date desc) AS ranking,
                                      [TT Code],
                                      GeneratedDate,
                                      RNC,
                                      Date from (

                                    select [TT Code], Date , GeneratedDate, RNC, DATEDIFF(day, Date,GeneratedDate) as DateDiff from (
                                    select [TT Code], Date , RNC,
                                    TRY_CAST(CAST(CAST(SUBSTRING(CONVERT (VARCHAR(50), [TT Code], 128), 1, 9)*1e7 AS INT) AS VARCHAR(8)) AS DATE) as 'GeneratedDate' 
                                    from Tehran_CC_Maintable where 
                                    (cast([Ticket Status] as varchar)='باز داخل باکس شخصی' or cast([Ticket Status] as varchar)='باز داخل باکس گروهی' ) and 
                                    (cast([The Last Agent Name] as varchar)!='Delayناک تهران' and [The Last Agent Name]!='Mohamadreza Kazmi')  )tbl 

                                    )tbl ) tbl where ranking=1
                                    ) TBL1

                                    left join
                                    (
                                    select [TT Code],  count(*) 'Count of TT' from Tehran_CC_Maintable where 
                                    (cast([Ticket Status] as varchar)='باز داخل باکس شخصی' or cast([Ticket Status] as varchar)='باز داخل باکس گروهی' ) and 
                                    (cast([The Last Agent Name] as varchar)!='Delayناک تهران' and [The Last Agent Name]!='Mohamadreza Kazmi')
                                    group by [TT Code]) TBL2

                                    on

                                    TBL1.[TT Code]=TBL2.[TT Code]

                                    ) MTTR90 left join Tehran_CC_RNC_Owner

                                    on
                                    MTTR90.RNC=Tehran_CC_RNC_Owner.[RNC Name]";




                DataTable MTTR_Table = Query_Execution_Table_Output(MTTR_Quary);

                SqlBulkCopy objbulk_MTTR = new SqlBulkCopy(connection);
                objbulk_MTTR.DestinationTableName = "Tehran_CC_FOMTTR";
                objbulk_MTTR.ColumnMappings.Add("Date", "Date");
                objbulk_MTTR.ColumnMappings.Add("TT Code", "TT Code");
                objbulk_MTTR.ColumnMappings.Add("GeneratedDate", "GeneratedDate");
                objbulk_MTTR.ColumnMappings.Add("RNC", "RNC");
                objbulk_MTTR.ColumnMappings.Add("RNC Owner", "RNC Owner");
                objbulk_MTTR.ColumnMappings.Add("Count of TT", "Count of TT");
                objbulk_MTTR.WriteToServer(MTTR_Table);

            }


            checkBox5.Invoke(new Action(() => checkBox5.Checked = true));

            //FO MTTR of Tickets LAst 90 Days
            for (int i = 0; i < Date_Table.Rows.Count; i++)
            {
                DateTime Date1 = Convert.ToDateTime((Date_Table.Rows[i]).ItemArray[0]);
                string Date = Date_ToString(Date1);

                if (Date1 <= Last_Updated_Date_Tehran_CC_FOMTTR_90)
                {
                    continue;
                }



                string MTTR_Quary = @"select MTTR90.Date, MTTR90.[TT Code], MTTR90.GeneratedDate, MTTR90.RNC, Tehran_CC_RNC_Owner.[RNC Owner], MTTR90.[Count of TT] from (


                                    select '" +
                                    Date + @"' as Date, TBL1.[TT Code],  TBL1.GeneratedDate, TBL1.RNC, TBL2.[Count of TT] from (

                                    select[TT Code], Date , GeneratedDate, RNC from(

                                    select 
                                      RANK() OVER(PARTITION BY  [TT Code] ORDER BY Date desc) AS ranking,
                                      [TT Code],
                                      GeneratedDate,
                                      RNC,
                                      Date from (

                                    select [TT Code], Date , GeneratedDate, RNC, DATEDIFF(day, Date,GeneratedDate) as DateDiff from (
                                    select [TT Code], Date , RNC,
                                    TRY_CAST(CAST(CAST(SUBSTRING(CONVERT (VARCHAR(50), [TT Code], 128), 1, 9)*1e7 AS INT) AS VARCHAR(8)) AS DATE) as 'GeneratedDate' 
                                    from Tehran_CC_Maintable where 
                                    (cast([Ticket Status] as varchar)='باز داخل باکس شخصی' or cast([Ticket Status] as varchar)='باز داخل باکس گروهی' ) and 
                                    (cast([The Last Agent Name] as varchar)!='Delayناک تهران' and [The Last Agent Name]!='Mohamadreza Kazmi')  )tbl where DATEDIFF(day, '" + Date + @"',GeneratedDate)>=-90 

                                    )tbl ) tbl where ranking=1
                                    ) TBL1

                                    left join
                                    (
                                    select [TT Code],  count(*) 'Count of TT' from Tehran_CC_Maintable where 
                                    (cast([Ticket Status] as varchar)='باز داخل باکس شخصی' or cast([Ticket Status] as varchar)='باز داخل باکس گروهی' ) and 
                                    (cast([The Last Agent Name] as varchar)!='Delayناک تهران' and [The Last Agent Name]!='Mohamadreza Kazmi')
                                    group by [TT Code]) TBL2

                                    on

                                    TBL1.[TT Code]=TBL2.[TT Code]

                                    ) MTTR90 left join Tehran_CC_RNC_Owner

                                    on
                                    MTTR90.RNC=Tehran_CC_RNC_Owner.[RNC Name]";




                DataTable MTTR_Table = Query_Execution_Table_Output(MTTR_Quary);

                SqlBulkCopy objbulk_MTTR = new SqlBulkCopy(connection);
                objbulk_MTTR.DestinationTableName = "Tehran_CC_FOMTTR_90";
                objbulk_MTTR.ColumnMappings.Add("Date", "Date");
                objbulk_MTTR.ColumnMappings.Add("TT Code", "TT Code");
                objbulk_MTTR.ColumnMappings.Add("GeneratedDate", "GeneratedDate");
                objbulk_MTTR.ColumnMappings.Add("RNC", "RNC");
                objbulk_MTTR.ColumnMappings.Add("RNC Owner", "RNC Owner");
                objbulk_MTTR.ColumnMappings.Add("Count of TT", "Count of TT");
                objbulk_MTTR.WriteToServer(MTTR_Table);


            }

            checkBox6.Invoke(new Action(() => checkBox6.Checked = true));





            //FO MTTR of Tickets Last 90 days  DailyResults

            string Last_GDateFO_Str = @"select max(Date) from Tehran_CC_FOMTTR_90";
            DataTable Last_GDateFO_Table = Query_Execution_Table_Output(Last_GDateFO_Str);
            DateTime Last_GDateFO = Convert.ToDateTime((Last_GDateFO_Table.Rows[Last_GDateFO_Table.Rows.Count - 1]).ItemArray[0]);
            Last_GDateFO = Last_GDateFO.AddDays(-90);
            string Last_GDateFO_String = Date_ToString(Last_GDateFO);

            string GDateFO_Str = @"select distinct GeneratedDate from Tehran_CC_FOMTTR_90 where GeneratedDate>='" + Last_GDateFO_String + "'";
            DataTable GDateFO_Table = Query_Execution_Table_Output(GDateFO_Str);


            Query_Execution("delete from Tehran_CC_FOMTTR_DailyResults_90 where GeneratedDate>='" + Last_GDateFO_String + "'");


            for (int i = 0; i < GDateFO_Table.Rows.Count; i++)
            {
                DateTime D1 = Convert.ToDateTime((GDateFO_Table.Rows[i]).ItemArray[0]);
                DateTime D2 = D1.AddDays(-90);
                string D1_String = Date_ToString(D1);
                string D2_String = Date_ToString(D2);

                string GDateFO_Daily_Str = @"select '" + D1_String + @"' as 'GeneratedDate', [RNC], [RNC Owner], count([TT Code]) as 'Count of TT' from(
               select TBL1.RNC, TBL1.[TT Code], Tehran_CC_RNC_Owner.[RNC Owner] from(
               select distinct[TT Code], [RNC] from Tehran_CC_FOMTTR_90
               where [GeneratedDate] >= '" + D2_String + "' and[GeneratedDate] <= '" + D1_String + "') TBL1 left join Tehran_CC_RNC_Owner on TBL1.RNC = Tehran_CC_RNC_Owner.[RNC Name]) tbl group by [RNC], [RNC Owner]";



                DataTable GDateFO_Daily_Table = Query_Execution_Table_Output(GDateFO_Daily_Str);


                SqlBulkCopy objbulk_GDateFO = new SqlBulkCopy(connection);
                objbulk_GDateFO.DestinationTableName = "Tehran_CC_FOMTTR_DailyResults_90";
                objbulk_GDateFO.ColumnMappings.Add("GeneratedDate", "GeneratedDate");
                objbulk_GDateFO.ColumnMappings.Add("RNC", "RNC");
                objbulk_GDateFO.ColumnMappings.Add("RNC Owner", "RNC Owner");
                objbulk_GDateFO.ColumnMappings.Add("Count of TT", "Count of TT");
                objbulk_GDateFO.WriteToServer(GDateFO_Daily_Table);



            }


            checkBox7.Invoke(new Action(() => checkBox7.Checked = true));


            // delete older than 90 days
            string Last_GDate_Str3 = @"select max(GeneratedDate) from Tehran_CC_FOMTTR_DailyResults_90";
            DataTable Last_GDate_Table3 = Query_Execution_Table_Output(Last_GDate_Str3);
            DateTime Last_GDate3 = Convert.ToDateTime((Last_GDate_Table3.Rows[Last_GDate_Table3.Rows.Count - 1]).ItemArray[0]);
            Last_GDate3 = Last_GDate3.AddDays(-90);
            string Last_GDate_String3 = Date_ToString(Last_GDate3);

            Query_Execution("delete from Tehran_CC_FOMTTR_DailyResults_90 where GeneratedDate<'" + Last_GDate_String3 + "'");




            // Daily Open_Not belong to Current Quarter 
            for (int i = 0; i < Date_Table.Rows.Count; i++)
            {
                DateTime Date1 = Convert.ToDateTime((Date_Table.Rows[i]).ItemArray[0]);
                string Date = Date_ToString(Date1);

                if (Date1 <= Last_Updated_Date_Tehran_CC_Open_NotBlong_ToCQ)
                {
                    continue;
                }


                string QDate_Query = @"select * from Tehran_CC_QDate";
                DataTable QDate_Table = Query_Execution_Table_Output(QDate_Query);


                string Q_Open_Quary = @"select Date, 
                                                   [TT Code],
            	                                   RNC,
            	                                   [TT Date], 
                                                   DATEDIFF(day, Date,[TT Date]) as DateDiff, 
            	                                   [Ticket Status], 
            	                                   [The Last Agent Name], '' as QDate, '' as 'QDateDiff'
                                                   from (      
            		                                  select Date, [TT Code], RNC,
            		                                  TRY_CAST(CAST(CAST(SUBSTRING(CONVERT (VARCHAR(50), [TT Code], 128), 1, 9)*1e7 AS decimal) AS VARCHAR(8)) AS DATE) as 'TT Date', 
            		                                  [Ticket Status], [The Last Agent Name]   from Tehran_CC_Maintable 
            		                                  where (cast([Ticket Status] as varchar)='باز داخل باکس شخصی' or cast([Ticket Status] as varchar)='باز داخل باکس گروهی' ) 
                                                            and
            				                                (cast([The Last Agent Name] as varchar)='Delayناک تهران' or cast([The Last Agent Name] as varchar)='Mohamadreza Kazmi' )
            				                                and 
            				                                Date='" + Date + @"') tbl";


                DataTable Q_Open_Table = Query_Execution_Table_Output(Q_Open_Quary);

                DateTime Selected_QDate = DateTime.Now;
                // Finding QDate comparing to file date
                for (int d = 0; d < QDate_Table.Rows.Count; d++)
                {
                    int QTable_Length = QDate_Table.Rows.Count;
                    DateTime QDate = Convert.ToDateTime((QDate_Table.Rows[QTable_Length - d - 1]).ItemArray[2]);
                    if (QDate < Date1)
                    {
                        Selected_QDate = QDate;
                        break;
                    }
                }


                for (int m = 0; m < Q_Open_Table.Rows.Count; m++)
                {
                    Q_Open_Table.Rows[m][7] = Selected_QDate;
                    TimeSpan difference = Convert.ToDateTime((Q_Open_Table.Rows[m]).ItemArray[3]) - Selected_QDate;
                    Q_Open_Table.Rows[m][8] = difference.Days;
                }


                SqlBulkCopy objbulk_Open_NotBlong_ToCQ = new SqlBulkCopy(connection);
                objbulk_Open_NotBlong_ToCQ.DestinationTableName = "Tehran_CC_Open_NotBlong_ToCQ";
                objbulk_Open_NotBlong_ToCQ.ColumnMappings.Add("Date", "Date");
                objbulk_Open_NotBlong_ToCQ.ColumnMappings.Add("TT Code", "TT Code");
                objbulk_Open_NotBlong_ToCQ.ColumnMappings.Add("RNC", "RNC");
                objbulk_Open_NotBlong_ToCQ.ColumnMappings.Add("TT Date", "TT Date");
                objbulk_Open_NotBlong_ToCQ.ColumnMappings.Add("DateDiff", "DateDiff");
                objbulk_Open_NotBlong_ToCQ.ColumnMappings.Add("Ticket Status", "Ticket Status");
                objbulk_Open_NotBlong_ToCQ.ColumnMappings.Add("The Last Agent Name", "The Last Agent Name");
                objbulk_Open_NotBlong_ToCQ.ColumnMappings.Add("QDate", "QDate");
                objbulk_Open_NotBlong_ToCQ.ColumnMappings.Add("QDateDiff", "QDateDiff");
                objbulk_Open_NotBlong_ToCQ.WriteToServer(Q_Open_Table);
            }


            checkBox8.Invoke(new Action(() => checkBox8.Checked = true));

            label1.Invoke(new Action(() => label1.Text = "Finished"));
            label1.Invoke(new Action(() => label1.BackColor = Color.Green));

            //label1.Text = "Finished";
            //label1.BackColor = Color.Green;

            connection.Close();
            MessageBox.Show("Finished");


        }





    }
}
