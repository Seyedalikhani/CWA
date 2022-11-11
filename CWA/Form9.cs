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
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Threading;
using System.Data.SqlClient;
using System.Reflection;

namespace CWA
{
    public partial class Form9 : Form
    {
        public Form9()
        {
            InitializeComponent();
        }


        public Form1 form1;


        public Form9(Form form)
        {
            InitializeComponent();
            form1 = (Form1)form;


        }



        public string FName = "";
        public IXLWorksheet Source_worksheet = null;
        public Excel.Application xlApp { get; set; }
        public Excel.Workbook KPI_workbook { get; set; }
        public static string DefaultExt { get; private set; }
        public XLWorkbook Source_workbook = new XLWorkbook();
        public DataTable Availability_Table = new DataTable();
        public DataTable Availability_Table_Result = new DataTable();
        public DataTable Availability_Table_Result_Cell = new DataTable();
        public string Hour = "";
        public DateTime d3 = DateTime.Now;
        public DateTime Last_Time = DateTime.Now;
        public DateTime First_Time = DateTime.Now;
        public string Mode = "Fixed Hour";
        public string[] Days_Vec = new string[100000];
        public string[] NEs_Vec = new string[100000];
        public string[] Availabilities_Vec = new string[100000];
        public int Counts = 0;
        public string ConnectionString = "";
        public SqlConnection connection = new SqlConnection();
        //public string Server_Name = "172.26.7.159";
        public string Server_Name = "PERFORMANCEDB01";
        public string DataBase_Name = "Performance_NAK";
        public string Technology = "2G";
        public DataTable Fluctuation_Results = new DataTable();
        public DataTable Fluctuation_Results1 = new DataTable();
        public DataTable Data_Table_2G = new DataTable();
        public DataTable Data_Table_3G = new DataTable();
        public DataTable Data_Table_4G = new DataTable();

        public string Task = "Delivery";
        public string Input_Type = "DataBase";

        public string Delivery_Task_Type = "Availability";
        public string[] Traffics_Vec = new string[100000];
        public DataTable Traffic_Table = new DataTable();
        public DataTable Traffic_Table_Result_Cell = new DataTable();






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






        void my_thread1()
        {

            Availability_Table.Columns.Add("Date", typeof(DateTime));
            Availability_Table.Columns.Add("Site", typeof(string));
            Availability_Table.Columns.Add("NE", typeof(string));
            Availability_Table.Columns.Add("Cell", typeof(string));
            Availability_Table.Columns.Add("Availability", typeof(double));
            Availability_Table.Columns.Add("Cell_Status", typeof(double));

            Availability_Table_Result = new DataTable();
            Availability_Table_Result.Columns.Add("Site", typeof(string));
            Availability_Table_Result.Columns.Add("Site Availability Status", typeof(string));


            Availability_Table_Result_Cell = new DataTable();
            Availability_Table_Result_Cell.Columns.Add("Site", typeof(string));
            Availability_Table_Result_Cell.Columns.Add("NE", typeof(string));
            Availability_Table_Result_Cell.Columns.Add("Cell", typeof(string));
            Availability_Table_Result_Cell.Columns.Add("Cell Availability Status", typeof(string));
            Availability_Table_Result_Cell.Columns.Add("Site Availability Status", typeof(string));
            Availability_Table_Result_Cell.Columns.Add("N.O Delivered Time", typeof(DateTime));






            Traffic_Table.Columns.Add("Date", typeof(DateTime));
            Traffic_Table.Columns.Add("Site", typeof(string));
            Traffic_Table.Columns.Add("NE", typeof(string));
            Traffic_Table.Columns.Add("Traffic", typeof(double));
            //   Traffic_Table.Columns.Add("Cell_Status", typeof(double));



            Traffic_Table_Result_Cell = new DataTable();
            Traffic_Table_Result_Cell.Columns.Add("Site", typeof(string));
            Traffic_Table_Result_Cell.Columns.Add("Site (G1800)", typeof(string));
            Traffic_Table_Result_Cell.Columns.Add("Cell", typeof(string));
            Traffic_Table_Result_Cell.Columns.Add("Cell Traffic (G1800)", typeof(string));
            Traffic_Table_Result_Cell.Columns.Add("Cell Traffic Status", typeof(string));
            Traffic_Table_Result_Cell.Columns.Add("Site Traffic Status", typeof(string));
       //     Traffic_Table_Result_Cell.Columns.Add("N.O Delivered Time", typeof(DateTime));




            int first = 0;
            for (int k = 0; k < Counts; k++)
            {
                DateTime d1 = Convert.ToDateTime(Days_Vec[k]);
                if (k == Counts - 1)
                {
                    First_Time = d1;
                }

                string Year = Convert.ToString(d1.Year);
                string d2 = d1.DayOfWeek + "  " + d1.Day + "/" + d1.Month + "/" + Year.Substring(2, 2);
                d3 = d1;
            }
            First_Time = First_Time.AddHours(-48);


            if (Mode == "Fixed Hour")
            {

                if (Delivery_Task_Type == "Availability")
                {

                    double Availabilty_Score = -1;


                    for (int k = 0; k < Counts; k++)
                    {
                        string NE = Convert.ToString(NEs_Vec[k]);
                        string Site = "";

                        string str2 = Regex.Replace(NE, "[^a-zA-Z0-9]", " ");      //هر کاراکتری که غیر از عدد و حرف بود را به کاراکتر خالی تبدیل کن
                        str2 = Regex.Replace(str2, " {2,}", " ").Trim();           //چندین کاراکتر خالی پشت سر هم را به یک کاراکتر خالی تبدیل می کند
                        string[] Split_Description = str2.Split(' ');

                        string Tech = Split_Description[0].Substring(0, 1);
                        string Tech_Last = Split_Description[0].Substring(Split_Description[0].Length - 1, 1);

                        string CellName = "";
                        if ((Tech == "B" && (Tech_Last == "E" || Tech_Last == "H" || Tech_Last == "N")) || Split_Description[0].Length == 2)
                        {

                            CellName = Split_Description[1];
                        }
                        else
                        {
                            int c1 = Split_Description.Length;
                            CellName = Split_Description[c1 - 1];
                        }

                        if (CellName.Length == 7)
                        {
                            Site = CellName.Substring(0, 6);
                        }
                        if (CellName.Length > 7)
                        {
                            Site = CellName.Substring(0, 8);
                        }

                        double Availability_Value = 1;
                        if (Availabilities_Vec[k] != "")
                        {
                            Availability_Value = Convert.ToDouble(Availabilities_Vec[k]);
                        }
                        else if (Availabilities_Vec[k] == "")
                        {
                            Availability_Value = -1;
                        }


                        //     k++;

                        if (Availability_Value >= 99.9)
                        {
                            Availabilty_Score = 1;
                        }
                        if (Availability_Value < 99.9)
                        {
                            Availabilty_Score = 0.1;
                        }
                        if (Availability_Value == -1)
                        {
                            Availabilty_Score = 0;
                        }


                        DateTime d1 = Convert.ToDateTime(Days_Vec[k]);

                        double Difference = (Last_Time - d1).TotalHours;

                        if (Difference < 48 && Difference >= 0)
                        {
                            Availability_Table.Rows.Add(Days_Vec[k], Site, NE, CellName, Availability_Value, Availabilty_Score);
                        }

                    }



                    // delete Rehoming Cells


                    // Group by Date and Cell
                    var results = from row in Availability_Table.AsEnumerable()
                                  group row by new { f1 = row.Field<string>("Cell"), f2 = row.Field<DateTime>("Date") } into rows
                                  select new
                                  {
                                      Date1 = rows.Key.f2,
                                      Cell1 = rows.Key.f1,
                                      Cell_Count = rows.Count()
                                  };



                    DataTable Data_Table_Av = new DataTable();
                    Data_Table_Av = ConvertToDataTable(results);


                    // Join Using Linq (The First Step: Left Join with Null Values)
                    var q = (from pd in Availability_Table.AsEnumerable()
                             join od in Data_Table_Av.AsEnumerable() on new { f1 = pd.Field<string>("Cell"), f2 = pd.Field<DateTime>("Date") } equals new { f1 = od.Field<string>("Cell1"), f2 = od.Field<DateTime>("Date1") } into od
                             from new_od in od.DefaultIfEmpty()
                             select new
                             {
                                 Date = pd.Field<DateTime>("Date"),
                                 Site = pd.Field<string>("Site"),
                                 NE = pd.Field<string>("NE"),
                                 Cell = pd.Field<string>("Cell"),
                                 Availability = (pd != null ? pd.Field<Double>("Availability") : -1),
                                 Cell_Status = (pd != null ? pd.Field<Double>("Cell_Status") : -1),
                                 Cell_Count = (new_od != null ? new_od.Field<Int32>("Cell_Count") : -1),
                               }).ToList();



                    DataTable Availability_Table_Modified1 = new DataTable();
                    Availability_Table_Modified1 = ConvertToDataTable(q);




                    DataTable Availability_Table_Modified = new DataTable();

                    Availability_Table_Modified.Columns.Add("Date", typeof(DateTime));
                    Availability_Table_Modified.Columns.Add("Site", typeof(string));
                    Availability_Table_Modified.Columns.Add("NE", typeof(string));
                    Availability_Table_Modified.Columns.Add("Cell", typeof(string));
                    Availability_Table_Modified.Columns.Add("Availability", typeof(double));
                    Availability_Table_Modified.Columns.Add("Cell_Status", typeof(double));
                    Availability_Table_Modified.Columns.Add("Cell_Count", typeof(Int32));

                    for (int k=0; k< Availability_Table_Modified1.Rows.Count; k++)
                    {
                        double Avai = Convert.ToDouble(Availability_Table_Modified1.Rows[k].ItemArray[4]);
                        Int32 Count = Convert.ToInt32(Availability_Table_Modified1.Rows[k].ItemArray[6]);
                        if (Avai == -1 && Count == 2)
                        {
                            continue;
                        }

                        Availability_Table_Modified.Rows.Add(Availability_Table_Modified1.Rows[k].ItemArray[0], Availability_Table_Modified1.Rows[k].ItemArray[1], Availability_Table_Modified1.Rows[k].ItemArray[2], Availability_Table_Modified1.Rows[k].ItemArray[3], Availability_Table_Modified1.Rows[k].ItemArray[4], Availability_Table_Modified1.Rows[k].ItemArray[5], Availability_Table_Modified1.Rows[k].ItemArray[6]);
                    }




                    var distinctIds = Availability_Table_Modified.AsEnumerable()
    .Select(s => new
    {
        id = s.Field<string>("Site"),
    })
    .Distinct().ToList();



                    // Site Status
                    for (int j = 0; j < distinctIds.Count; j++)
                    {
                        var Site_Data = (from p in Availability_Table_Modified.AsEnumerable()
                                         where p.Field<string>("Site") == distinctIds[j].id
                                         select p).ToList();


                        double multiplier = 1;
                        double Availability_Sum = 0;
                        double Availability_Sum_LastDay = 0;
                        for (int h = 0; h < Site_Data.Count; h++)
                        {

                            if (Convert.ToInt16(Site_Data[h].ItemArray[4])==-1 && Convert.ToInt16(Site_Data[h].ItemArray[6])==2)
                            {
                                continue;
                            }
                            multiplier = multiplier * Convert.ToDouble(Site_Data[h].ItemArray[5]);
                            Availability_Sum = Availability_Sum + Convert.ToDouble(Site_Data[h].ItemArray[4]);

                            if (Convert.ToDateTime(Last_Time) == Convert.ToDateTime(Site_Data[h].ItemArray[0]))
                            {
                                Availability_Sum_LastDay = Availability_Sum_LastDay + Convert.ToDouble(Site_Data[h].ItemArray[4]);
                            }

                        }




                        if (Availability_Sum == 0 || Availability_Sum_LastDay == 0)
                        {
                            Availability_Table_Result.Rows.Add(distinctIds[j].id, "NI-Site Down");
                        }

                        if (Availability_Sum != 0 && Availability_Sum_LastDay != 0)
                        {
                            if (multiplier == 1)
                            {
                                Availability_Table_Result.Rows.Add(distinctIds[j].id, "N.O Delivered");
                            }
                            if (multiplier == 0)
                            {
                                Availability_Table_Result.Rows.Add(distinctIds[j].id, "NI");
                            }
                            if (multiplier < 1 && multiplier > 0)
                            {
                                Availability_Table_Result.Rows.Add(distinctIds[j].id, "NI-Cell/Site Fluctuating");
                            }
                        }

                    }


                    var distinctIds1 = Availability_Table_Modified.AsEnumerable()
    .Select(s => new
    {
        id = s.Field<string>("NE"),
    })
    .Distinct().ToList();


                    // Cell Status
                    for (int j = 0; j < distinctIds1.Count; j++)
                    {
                        var Cell_Data = (from p in Availability_Table_Modified.AsEnumerable()
                                         where p.Field<string>("NE") == distinctIds1[j].id
                                         select p).ToList();

                     
                        string Site = "";
                        string Cell = distinctIds1[j].id;

                        string str2 = Regex.Replace(Cell, "[^a-zA-Z0-9]", " ");      //هر کاراکتری که غیر از عدد و حرف بود را به کاراکتر خالی تبدیل کن
                        str2 = Regex.Replace(str2, " {2,}", " ").Trim();           //چندین کاراکتر خالی پشت سر هم را به یک کاراکتر خالی تبدیل می کند
                        string[] Split_Description = str2.Split(' ');


                        string Tech = Split_Description[0].Substring(0, 1);
                        string Tech_Last = Split_Description[0].Substring(Split_Description[0].Length - 1, 1);

                        string CellName = "";
                        if ((Tech == "B" && (Tech_Last == "E" || Tech_Last == "H" || Tech_Last == "N")) || Split_Description[0].Length == 2)
                        {
                            //int c1 = Split_Description.Length;
                            CellName = Split_Description[1];
                        }
                        else
                        {
                            int c1 = Split_Description.Length;
                            CellName = Split_Description[c1 - 1];
                        }


                        if (CellName.Length == 7)
                        {
                            Site = CellName.Substring(0, 6);
                        }
                        if (CellName.Length > 7)
                        {
                            Site = CellName.Substring(0, 8);
                        }


                        var Site_Status = (from p1 in Availability_Table_Result.AsEnumerable()
                                           where p1.Field<string>("Site") == Site
                                           select p1).ToList();


                        string Site_Status1 = Site_Status[0].ItemArray[1].ToString();




                        double multiplier = 1;
                        double Availability_Sum = 0;
                        double Availability_Sum_LastDay = 0;
                        for (int h = 0; h < Cell_Data.Count; h++)
                        {


                            //if (Convert.ToInt16(Cell_Data[h].ItemArray[4]) == -1 && Convert.ToInt16(Cell_Data[h].ItemArray[6]) == 2)
                            //{
                            //    continue;
                            //}


                            multiplier = multiplier * Convert.ToDouble(Cell_Data[h].ItemArray[5]);
                            Availability_Sum = Availability_Sum + Convert.ToDouble(Cell_Data[h].ItemArray[4]);

                            if (Convert.ToDateTime(Last_Time) == Convert.ToDateTime(Cell_Data[h].ItemArray[0]))
                            {
                                Availability_Sum_LastDay = Availability_Sum_LastDay + Convert.ToDouble(Cell_Data[h].ItemArray[4]);
                            }

                        }


                        if (Availability_Sum == 0 || Availability_Sum_LastDay == 0)
                        {
                            Availability_Table_Result_Cell.Rows.Add(Site, distinctIds1[j].id, CellName, "NI-Cell Down", Site_Status1);
                        }

                        if (Availability_Sum != 0 && Availability_Sum_LastDay != 0)
                        {
                            if (multiplier == 1)
                            {
                                if (Site_Status1 == "N.O Delivered")
                                {
                                    Availability_Table_Result_Cell.Rows.Add(Site, distinctIds1[j].id, CellName, "N.O Delivered", Site_Status1, Last_Time);
                                }
                                else
                                {
                                    Availability_Table_Result_Cell.Rows.Add(Site, distinctIds1[j].id, CellName, "N.O Delivered", Site_Status1);
                                }
                            }
                            if (multiplier == 0)
                            {
                                Availability_Table_Result_Cell.Rows.Add(Site, distinctIds1[j].id, CellName, "NI", Site_Status1);
                            }
                            if (multiplier < 1 && multiplier > 0)
                            {
                                Availability_Table_Result_Cell.Rows.Add(Site, distinctIds1[j].id, CellName, "NI-Cell/Site Fluctuating", Site_Status1);
                            }
                        }

                    }

                    //label2.Text = "File's Loaded";
                    //label2.BackColor = Color.GreenYellow;

                    // MessageBox.Show("Process Finished");
                    this.Invoke(new Action(() => { MessageBox.Show(this, "Finished"); }));
                }




                if (Delivery_Task_Type == "Traffic")
                {
                    int rr = 0;


                    for (int k = 0; k < Counts; k++)
                    {
                        string Cell = Convert.ToString(NEs_Vec[k]);
                        string Site = "";

                        string str2 = Regex.Replace(Cell, "[^a-zA-Z0-9]", " ");      //هر کاراکتری که غیر از عدد و حرف بود را به کاراکتر خالی تبدیل کن
                        str2 = Regex.Replace(str2, " {2,}", " ").Trim();           //چندین کاراکتر خالی پشت سر هم را به یک کاراکتر خالی تبدیل می کند
                        string[] Split_Description = str2.Split(' ');

                        string Tech = Split_Description[0].Substring(0, 1);
                        string Tech_Last = Split_Description[0].Substring(Split_Description[0].Length - 1, 1);

                        string CellName = "";
                        if ((Tech == "B" && (Tech_Last == "E" || Tech_Last == "H" || Tech_Last == "N")) || Split_Description[0].Length == 2)
                        {

                            CellName = Split_Description[1];
                        }
                        else
                        {
                            int c1 = Split_Description.Length;
                            CellName = Split_Description[c1 - 1];
                        }

                        if (CellName.Length == 7)
                        {
                            Site = CellName.Substring(0, 6);
                        }
                        if (CellName.Length > 7)
                        {
                            Site = CellName.Substring(0, 8);
                        }


                        double Traffic_Value = 1;
                        if (Traffics_Vec[k] != "")
                        {
                            Traffic_Value = Convert.ToDouble(Traffics_Vec[k]);
                        }
                        else if (Traffics_Vec[k] == "")
                        {
                            Traffic_Value = -1;
                        }




                        DateTime d1 = Convert.ToDateTime(Days_Vec[k]);

                        double Difference = (Last_Time - d1).TotalHours;

                        if (Difference < 48 && Difference >= 0)
                        {
                            Traffic_Table.Rows.Add(Days_Vec[k], Site, Cell, Traffic_Value);
                        }

                    }



                    var distinctIds1 = Traffic_Table.AsEnumerable()
.Select(s => new
{
id = s.Field<string>("NE"),
})
.Distinct().ToList();



                    // Cell Status
                    for (int j = 0; j < distinctIds1.Count; j++)
                    {

                        var Cell_Data = (from p in Traffic_Table.AsEnumerable()
                                         where p.Field<string>("NE") == distinctIds1[j].id
                                         select p).ToList();



                        string Site = "";
                        string Site_1800 = "";
                        string Cell = distinctIds1[j].id;

                        string str2 = Regex.Replace(Cell, "[^a-zA-Z0-9]", " ");      //هر کاراکتری که غیر از عدد و حرف بود را به کاراکتر خالی تبدیل کن
                        str2 = Regex.Replace(str2, " {2,}", " ").Trim();           //چندین کاراکتر خالی پشت سر هم را به یک کاراکتر خالی تبدیل می کند
                        string[] Split_Description = str2.Split(' ');


                        string Tech = Split_Description[0].Substring(0, 1);
                        string Tech_Last = Split_Description[0].Substring(Split_Description[0].Length - 1, 1);

                        string CellName = "";
                        if ((Tech == "B" && (Tech_Last == "E" || Tech_Last == "H" || Tech_Last == "N")) || Split_Description[0].Length == 2)
                        {
                            //int c1 = Split_Description.Length;
                            CellName = Split_Description[1];
                        }
                        else
                        {
                            int c1 = Split_Description.Length;
                            CellName = Split_Description[c1 - 1];
                        }


                        if (CellName.Length == 7)
                        {
                            Site = CellName.Substring(0, 6);
                            double Site_1800_First = Convert.ToDouble(Site.Substring(2, 1));
                            Site_1800 = Site.Substring(0, 2) + "2G" + Site.Substring(2, 4);
                        }
                        if (CellName.Length > 7)
                        {
                            Site = CellName.Substring(0, 8);
                            double Site_1800_First = Convert.ToDouble(Site.Substring(4, 1));
                            Site_1800 = Site.Substring(0, 2) + "2G" + Site.Substring(4, 4);
                        }





                        double Traffic_Sum = 0;
                        int Counter = 0;
                        double Traffic_Ave = 0;
                        for (int h = 0; h < Cell_Data.Count; h++)
                        {
                            //multiplier = multiplier * Convert.ToDouble(Cell_Data[h].ItemArray[4]);
                            if (Convert.ToDouble(Cell_Data[h].ItemArray[3]) != -1)
                            {
                                Counter++;
                                Traffic_Sum = Traffic_Sum + Convert.ToDouble(Cell_Data[h].ItemArray[3]);
                            }

                        }
                        if (Counter!=0)
                        {
                            Traffic_Ave = Traffic_Sum / Counter; 
                        }
                        else
                        {
                            Traffic_Ave = -1;
                        }


                        string Cell_Status = "";
                        if (Traffic_Ave>=0.1)
                        {
                            Cell_Status = "N.O Delivered";
                        }
                        if (Traffic_Ave>0 && Traffic_Ave < 0.1)
                        {
                            Cell_Status = "NI";
                        }
                        if (Traffic_Ave == 0)
                        {
                            Cell_Status = "NI-Cell Down";
                        }
                        if (Traffic_Ave  < 0)
                        {
                            Cell_Status = "NI";
                        }

                        Traffic_Table_Result_Cell.Rows.Add(Site, Site_1800, Cell, Traffic_Ave, Cell_Status, "");


                    }



                    var distinctIds = Traffic_Table_Result_Cell.AsEnumerable()
.Select(s => new
{
id = s.Field<string>("Site"),
})
.Distinct().ToList();


                    for (int j = 0; j < distinctIds.Count; j++)
                    {

                        var Site_Data = (from p in Traffic_Table_Result_Cell.AsEnumerable()
                                         where p.Field<string>("Site") == distinctIds[j].id
                                         select p).ToList();


                        double multiplier = 1;
                        double Traffic_Sum = 0;
                        for (int h = 0; h < Site_Data.Count; h++)
                        {
                            double Score = 0;
                            if (Site_Data[h].ItemArray[4].ToString() == "N.O Delivered")
                            {
                                Score = 1;
                            }
                            if (Site_Data[h].ItemArray[4].ToString() == "NI")
                            {
                                Score = 0.1;
                            }
                            if (Site_Data[h].ItemArray[4].ToString() == "NI-Cell Down")
                            {
                                Score = 0;
                            }

                            multiplier = multiplier * Score;
                            Traffic_Sum = Traffic_Sum + Convert.ToDouble(Site_Data[h].ItemArray[3]);
                        }



                        string Site_Status = "";

                        if (Traffic_Sum==0)
                        {
                            Site_Status = "NI-Site Down";
                        }
                        else
                        {
                            if (multiplier==1)
                            {
                                Site_Status = "N.O Delivered";
                            }
                            if (multiplier>0 && multiplier < 1)
                            {
                                Site_Status = "NI";
                            }
                            if (multiplier==0)
                            {
                                Site_Status = "NI-Cell Down";
                            }
                        }





                        for (int k1=0; k1<= Traffic_Table_Result_Cell.Rows.Count-1; k1++)
                        {
                            string site1 = (Traffic_Table_Result_Cell.Rows[k1]).ItemArray[0].ToString();
                            if (site1 == distinctIds[j].id)
                            {
                                //  (Traffic_Table_Result_Cell.Rows[k1]).ItemArray[5] = Site_Status;
                                Traffic_Table_Result_Cell.Rows[k1][5] = Site_Status;
                            }
                        }






                    }


                    this.Invoke(new Action(() => { MessageBox.Show(this, "Finished"); }));


                }





            }




            if (Mode == "Searching")
            {
                double Availabilty_Score = -1;
                int Rejected = 100000000;
                string[] Approve_Site = new string[10000];
                int app_ind = 0;
                Last_Time = Last_Time.AddHours(+1);


                while (Rejected != 0)
                {


                    Availabilty_Score = -1;

                    Last_Time = Last_Time.AddHours(-6);




                    if ((Last_Time - First_Time).TotalHours < 48)
                    {
                        break;
                    }

                    Rejected = 0;
                    Availability_Table.Rows.Clear();


                    for (int k = 0; k < Counts; k++)
                    {
                        if (Convert.ToDateTime(Days_Vec[k]) <= Last_Time)
                        {
                            DateTime d1 = Convert.ToDateTime(Days_Vec[k]);

                            double Difference = (Last_Time - d1).TotalHours;

                            if (Difference < 48 && Difference >= 0)
                            {

                                string Cell = Convert.ToString(NEs_Vec[k]);
                                string Site = "";

                                string str2 = Regex.Replace(Cell, "[^a-zA-Z0-9]", " ");      //هر کاراکتری که غیر از عدد و حرف بود را به کاراکتر خالی تبدیل کن
                                str2 = Regex.Replace(str2, " {2,}", " ").Trim();           //چندین کاراکتر خالی پشت سر هم را به یک کاراکتر خالی تبدیل می کند
                                string[] Split_Description = str2.Split(' ');

                                string Tech = Split_Description[0].Substring(0, 1);
                                string Tech_Last = Split_Description[0].Substring(Split_Description[0].Length - 1, 1);

                                string CellName = "";
                                if ((Tech == "B" && (Tech_Last == "E" || Tech_Last == "H" || Tech_Last == "N")) || Split_Description[0].Length == 2)
                                {

                                    CellName = Split_Description[1];
                                }
                                else
                                {
                                    int c1 = Split_Description.Length;
                                    CellName = Split_Description[c1 - 1];
                                }

                                if (CellName.Length == 7)
                                {
                                    Site = CellName.Substring(0, 6);
                                }
                                if (CellName.Length > 7)
                                {
                                    Site = CellName.Substring(0, 8);
                                }


                                // back to the main route if site has been approved
                                if (Approve_Site.Contains(Site))
                                {
                                    continue;
                                }



                                double Availability_Value = 1;
                                if (Availabilities_Vec[k] != "")
                                {
                                    Availability_Value = Convert.ToDouble(Availabilities_Vec[k]);
                                }
                                else if (Availabilities_Vec[k] == "")
                                {
                                    Availability_Value = -1;

                                }


                                //  k++;

                                if (Availability_Value >= 99.9)
                                {
                                    Availabilty_Score = 1;
                                }
                                if (Availability_Value < 99.9)
                                {
                                    Availabilty_Score = 0.1;

                                }
                                if (Availability_Value == -1)
                                {
                                    Availabilty_Score = 0;

                                }


                                Availability_Table.Rows.Add(Days_Vec[k], Site, Cell, Availability_Value, Availabilty_Score);
                            }
                        }
                    }


                    var distinctIds = Availability_Table.AsEnumerable()
    .Select(s => new
    {
        id = s.Field<string>("Site"),
    })
    .Distinct().ToList();



                    // Site Status
                    for (int j = 0; j < distinctIds.Count; j++)
                    {
                        var Site_Data = (from p in Availability_Table.AsEnumerable()
                                         where p.Field<string>("Site") == distinctIds[j].id
                                         select p).ToList();


                        double multiplier = 1;
                        double Availability_Sum = 0;
                        double Availability_Sum_LastDay = 0;
                        for (int h = 0; h < Site_Data.Count; h++)
                        {

                            multiplier = multiplier * Convert.ToDouble(Site_Data[h].ItemArray[4]);
                            Availability_Sum = Availability_Sum + Convert.ToDouble(Site_Data[h].ItemArray[3]);

                            if (Convert.ToDateTime(Last_Time) == Convert.ToDateTime(Site_Data[h].ItemArray[0]))
                            {
                                Availability_Sum_LastDay = Availability_Sum_LastDay + Convert.ToDouble(Site_Data[h].ItemArray[3]);
                            }

                        }




                        if (Availability_Sum == 0 || Availability_Sum_LastDay == 0)
                        {
                            Availability_Table_Result.Rows.Add(distinctIds[j].id, "NI-Site Down");
                        }

                        if (Availability_Sum != 0 && Availability_Sum_LastDay != 0)
                        {
                            if (multiplier == 1)
                            {
                                Availability_Table_Result.Rows.Add(distinctIds[j].id, "N.O Delivered");
                            }
                            if (multiplier == 0)
                            {
                                Availability_Table_Result.Rows.Add(distinctIds[j].id, "NI");
                            }
                            if (multiplier < 1 && multiplier > 0)
                            {
                                Availability_Table_Result.Rows.Add(distinctIds[j].id, "NI-Cell/Site Fluctuating");
                            }
                        }

                    }


                    var distinctIds1 = Availability_Table.AsEnumerable()
    .Select(s => new
    {
        id = s.Field<string>("NE"),
    })
    .Distinct().ToList();



                    // Cell Status
                    for (int j = 0; j < distinctIds1.Count; j++)
                    {
                        var Cell_Data = (from p in Availability_Table.AsEnumerable()
                                         where p.Field<string>("NE") == distinctIds1[j].id
                                         select p).ToList();


                        string Site = "";
                        string Cell = distinctIds1[j].id;

                        string str2 = Regex.Replace(Cell, "[^a-zA-Z0-9]", " ");      //هر کاراکتری که غیر از عدد و حرف بود را به کاراکتر خالی تبدیل کن
                        str2 = Regex.Replace(str2, " {2,}", " ").Trim();           //چندین کاراکتر خالی پشت سر هم را به یک کاراکتر خالی تبدیل می کند
                        string[] Split_Description = str2.Split(' ');


                        string Tech = Split_Description[0].Substring(0, 1);
                        string Tech_Last = Split_Description[0].Substring(Split_Description[0].Length - 1, 1);

                        string CellName = "";
                        if ((Tech == "B" && (Tech_Last == "E" || Tech_Last == "H" || Tech_Last == "N")) || Split_Description[0].Length == 2)
                        {
                            //int c1 = Split_Description.Length;
                            CellName = Split_Description[1];
                        }
                        else
                        {
                            int c1 = Split_Description.Length;
                            CellName = Split_Description[c1 - 1];
                        }


                        if (CellName.Length == 7)
                        {
                            Site = CellName.Substring(0, 6);
                        }
                        if (CellName.Length > 7)
                        {
                            Site = CellName.Substring(0, 8);
                        }


                        var Site_Status = (from p1 in Availability_Table_Result.AsEnumerable()
                                           where p1.Field<string>("Site") == Site
                                           select p1).ToList();


                        string Site_Status1 = Site_Status[Site_Status.Count - 1].ItemArray[1].ToString();




                        double multiplier = 1;
                        double Availability_Sum = 0;
                        double Availability_Sum_LastDay = 0;
                        for (int h = 0; h < Cell_Data.Count; h++)
                        {

                            multiplier = multiplier * Convert.ToDouble(Cell_Data[h].ItemArray[4]);
                            Availability_Sum = Availability_Sum + Convert.ToDouble(Cell_Data[h].ItemArray[3]);

                            if (Convert.ToDateTime(Last_Time) == Convert.ToDateTime(Cell_Data[h].ItemArray[0]))
                            {
                                Availability_Sum_LastDay = Availability_Sum_LastDay + Convert.ToDouble(Cell_Data[h].ItemArray[3]);
                            }

                        }


                        if (Availability_Sum == 0 || Availability_Sum_LastDay == 0)
                        {
                            Availability_Table_Result_Cell.Rows.Add(Site, distinctIds1[j].id, "NI-Cell Down", Site_Status1);
                            Rejected++;
                        }


                        if (Availability_Sum != 0 && Availability_Sum_LastDay != 0)
                        {
                            if (multiplier == 1)
                            {
                                if (Site_Status1 == "N.O Delivered")
                                {
                                    Availability_Table_Result_Cell.Rows.Add(Site, distinctIds1[j].id, "N.O Delivered", Site_Status1, Last_Time);
                                    Approve_Site[app_ind] = Site;
                                    app_ind++;
                                }
                                else
                                {
                                    Availability_Table_Result_Cell.Rows.Add(Site, distinctIds1[j].id, "N.O Delivered", Site_Status1);
                                }
                            }
                            if (multiplier == 0)
                            {
                                Availability_Table_Result_Cell.Rows.Add(Site, distinctIds1[j].id, "NI", Site_Status1);
                                Rejected++;
                            }
                            if (multiplier < 1 && multiplier > 0)
                            {
                                Availability_Table_Result_Cell.Rows.Add(Site, distinctIds1[j].id, "NI-Cell/Site Fluctuating", Site_Status1);
                                Rejected++;
                            }
                        }

                        //}
                    }
                    //label2.Text = "File's Loaded";
                    //label2.BackColor = Color.GreenYellow;

                }


                this.Invoke(new Action(() => { MessageBox.Show(this, "Finished"); }));


            }


        }





        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.DefaultExt = "xlsx";
            openFileDialog1.Filter = "Excel File|*.xlsx";
            DialogResult result = openFileDialog1.ShowDialog();
            string File_Name = openFileDialog1.SafeFileName.ToString();
            if (result == DialogResult.OK)
            {
                string file = openFileDialog1.FileName;
                xlApp = new Excel.Application();
                KPI_workbook = xlApp.Workbooks.Open(file);
                Excel.Worksheet sheet1 = KPI_workbook.Worksheets[1];

                string KPI_Name = sheet1.Cells[1, 3].Value;
                if (KPI_Name == "TCH_Traffic_OverL(Erlang)(Eric_Cell)" || KPI_Name == "TRVOL_TCH_OVRLAID_SUBCELL(Hu_CELL)")
                {
                    Delivery_Task_Type = "Traffic";
                }
                else
                {
                    Delivery_Task_Type = "Availability";
                }




                // Extracting Days
                Excel.Range Day = sheet1.get_Range("A2", "A" + sheet1.UsedRange.Rows.Count);
                object[,] Days = (object[,])Day.Value;

                Excel.Range NE = sheet1.get_Range("B2", "B" + sheet1.UsedRange.Rows.Count);
                object[,] NEs = (object[,])NE.Value;


                if (Delivery_Task_Type == "Availability")
                {
                    Excel.Range Availability = sheet1.get_Range("C2", "C" + sheet1.UsedRange.Rows.Count);
                    object[,] Availabilities = (object[,])Availability.Value;


                    Days_Vec = new string[sheet1.UsedRange.Rows.Count];
                    NEs_Vec = new string[sheet1.UsedRange.Rows.Count];
                    Availabilities_Vec = new string[sheet1.UsedRange.Rows.Count];
                    Counts = sheet1.UsedRange.Rows.Count - 1;
                    for (int k = 0; k < Counts; k++)
                    {
                        Days_Vec[k] = Days[k + 1, 1].ToString();
                        NEs_Vec[k] = NEs[k + 1, 1].ToString();
                        if (Availabilities[k + 1, 1] != null)
                        {
                            Availabilities_Vec[k] = Availabilities[k + 1, 1].ToString();
                        }
                        else
                        {
                            Availabilities_Vec[k] = "";
                        }

                    }


                }
                else
                {
                    Excel.Range Traffic = sheet1.get_Range("C2", "C" + sheet1.UsedRange.Rows.Count);
                    object[,] Traffics = (object[,])Traffic.Value;



                    Days_Vec = new string[sheet1.UsedRange.Rows.Count];
                    NEs_Vec = new string[sheet1.UsedRange.Rows.Count];
                    Traffics_Vec = new string[sheet1.UsedRange.Rows.Count];
                    Counts = sheet1.UsedRange.Rows.Count - 1;
                    for (int k = 0; k < Counts; k++)
                    {
                        Days_Vec[k] = Days[k + 1, 1].ToString();
                        NEs_Vec[k] = NEs[k + 1, 1].ToString();
                        if (Traffics[k + 1, 1] != null)
                        {
                            Traffics_Vec[k] = Traffics[k + 1, 1].ToString();
                        }
                        else
                        {
                            Traffics_Vec[k] = "";
                        }

                    }


                }







            }

            Thread t1 = new Thread(my_thread1);
            t1.Start();

        }




        // Export
        private void button2_Click(object sender, EventArgs e)
        {

            if (Task == "Delivery" && Delivery_Task_Type=="Availability")
            {
                XLWorkbook wb = new XLWorkbook();
                wb.Worksheets.Add(Availability_Table, "Data Table");
                //wb.Worksheets.Add(Availability_Table_Result, "Site Result");
                wb.Worksheets.Add(Availability_Table_Result_Cell, "Status");
                var saveFileDialog = new SaveFileDialog
                {
                    FileName = "Availability Delivery Report",
                    Filter = "Excel files|*.xlsx",
                    Title = "Save an Excel File"
                };



                saveFileDialog.ShowDialog();

                if (!String.IsNullOrWhiteSpace(saveFileDialog.FileName))
                    wb.SaveAs(saveFileDialog.FileName);

                MessageBox.Show("Finished");
            }


            if (Task == "Delivery" && Delivery_Task_Type == "Traffic")
            {
                XLWorkbook wb = new XLWorkbook();
                wb.Worksheets.Add(Traffic_Table, "Data Table");
                //wb.Worksheets.Add(Availability_Table_Result, "Site Result");
                wb.Worksheets.Add(Traffic_Table_Result_Cell, "Status");
                var saveFileDialog = new SaveFileDialog
                {
                    FileName = "Traffic Delivery Report",
                    Filter = "Excel files|*.xlsx",
                    Title = "Save an Excel File"
                };



                saveFileDialog.ShowDialog();

                if (!String.IsNullOrWhiteSpace(saveFileDialog.FileName))
                    wb.SaveAs(saveFileDialog.FileName);

                MessageBox.Show("Finished");
            }

            if (Task == "Fluctuation" && Input_Type == "FARAZ")
            {
                XLWorkbook wb = new XLWorkbook();
                wb.Worksheets.Add(Data_Table_2G, "Data Table");
                //wb.Worksheets.Add(Availability_Table_Result, "Site Result");
                wb.Worksheets.Add(Fluctuation_Results, "Status");
                var saveFileDialog = new SaveFileDialog
                {
                    FileName = "Availability Fluctuation Report",
                    Filter = "Excel files|*.xlsx",
                    Title = "Save an Excel File"
                };



                saveFileDialog.ShowDialog();

                if (!String.IsNullOrWhiteSpace(saveFileDialog.FileName))
                    wb.SaveAs(saveFileDialog.FileName);

                MessageBox.Show("Finished");
            }

            if (Task == "Fluctuation" && Input_Type == "DataBase")
            {
                Fluctuation_Results1 = Fluctuation_Results;
                Fluctuation_Results1.Columns.Remove("Cell Score Indicator");


                XLWorkbook wb = new XLWorkbook();
                if (Technology == "2G")
                {
                    wb.Worksheets.Add(Data_Table_2G, "Data Table");
                }
                if (Technology == "3G")
                {
                    wb.Worksheets.Add(Data_Table_3G, "Data Table");
                }
                if (Technology == "4G")
                {
                    wb.Worksheets.Add(Data_Table_4G, "Data Table");
                }
                //wb.Worksheets.Add(Availability_Table_Result, "Site Result");
                wb.Worksheets.Add(Fluctuation_Results, "Status");
                var saveFileDialog1 = new SaveFileDialog
                {
                    FileName = "Availability Fluctuation Report",
                    Filter = "Excel files|*.xlsx",
                    Title = "Save an Excel File"
                };



                saveFileDialog1.ShowDialog();

                if (!String.IsNullOrWhiteSpace(saveFileDialog1.FileName))
                    wb.SaveAs(saveFileDialog1.FileName);

                MessageBox.Show("Finished");


            }

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Hour = comboBox1.SelectedItem.ToString();
            Last_Time = dateTimePicker1.Value.Date.AddHours(Convert.ToDouble(Hour));

        }

        private void Form9_Load(object sender, EventArgs e)
        {
        }










        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked == true)
            {
                Mode = "Fixed Hour";
                checkBox4.Checked = false;
            }
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked == true)
            {
                Mode = "Searching";
                checkBox3.Checked = false;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Task = "Fluctuation";
            if (Input_Type == "DataBase")
            {
                Thread t2 = new Thread(my_thread2);
                t2.Start();
            }


        }



        void my_thread2()
        {

            ConnectionString = @"Server=" + Server_Name + "; Database=" + DataBase_Name + "; User ID=cwpcApp;Password=cwpcApp@830625#Ahmad";
            connection = new SqlConnection(ConnectionString);
            connection.Open();


            Fluctuation_Results.Columns.Add("Date", typeof(string));
            Fluctuation_Results.Columns.Add("Node", typeof(string));
            Fluctuation_Results.Columns.Add("Site", typeof(string));
            Fluctuation_Results.Columns.Add("Cell", typeof(string));
            Fluctuation_Results.Columns.Add("Zero Cell Availability (No of Days)", typeof(int));
            Fluctuation_Results.Columns.Add("Cell Availability less than 100% (No of Days)", typeof(int));
            Fluctuation_Results.Columns.Add("Cell Score", typeof(string));
            Fluctuation_Results.Columns.Add("Site Score", typeof(string));
            Fluctuation_Results.Columns.Add("Cell Score Indicator", typeof(double));




            string site_list = textBox1.Text;
            string Site_Disticnts1 = Regex.Replace(site_list, "[^a-zA-Z0-9]", " ");
            Site_Disticnts1 = Regex.Replace(Site_Disticnts1, " {2,}", " ").Trim();
            string[] Site_Code_Distincts = Site_Disticnts1.Split(' ');

            string EH_sites_list = "";
            string H_sites_list = "";
            string N_sites_list = "";
            string sites_list = "";



            for (int i = 0; i < Site_Code_Distincts.Count(); i++)
            {
                if (Technology == "2G")
                {
                    EH_sites_list = EH_sites_list + "substring([Cell],1,6)='" + Site_Code_Distincts[i] + "' or ";
                    H_sites_list = H_sites_list + "substring([Cell],1,2)+substring([Cell],5,4)='" + Site_Code_Distincts[i] + "' or ";
                    N_sites_list = N_sites_list + "substring([Seg],1,2)+substring([Seg],5,4)='" + Site_Code_Distincts[i] + "' or ";
                }
                if (Technology == "3G")
                {
                    sites_list = sites_list + "substring([ElementID1],1,8)='" + Site_Code_Distincts[i] + "' or ";
                }
                if (Technology == "4G")
                {
                    EH_sites_list = EH_sites_list + "substring([eNodeB],1,8)='" + Site_Code_Distincts[i] + "' or ";
                    N_sites_list = N_sites_list + "substring([ElementID1],1,8)='" + Site_Code_Distincts[i] + "' or ";
                }

            }
            if (Technology == "2G")
            {
                EH_sites_list = EH_sites_list.Substring(0, EH_sites_list.Length - 4);
                H_sites_list = H_sites_list.Substring(0, H_sites_list.Length - 4);
                N_sites_list = N_sites_list.Substring(0, N_sites_list.Length - 4);
            }
            if (Technology == "3G")
            {
                sites_list = sites_list.Substring(0, sites_list.Length - 4);
            }
            if (Technology == "4G")
            {
                EH_sites_list = EH_sites_list.Substring(0, EH_sites_list.Length - 4);
                N_sites_list = N_sites_list.Substring(0, N_sites_list.Length - 4);
            }




            string start_date = Convert.ToString(dateTimePicker2.Value.Date);
            string end_date = Convert.ToString(dateTimePicker3.Value.Date);
            //string start_date = Convert.ToString(dateTimePicker2.Value.Year) + "-" + Convert.ToString(dateTimePicker2.Value.Month) + "-" + Convert.ToString(dateTimePicker2.Value.Day) + " 00:00:00.000";
            //string end_date = Convert.ToString(dateTimePicker3.Value.Year) + "-" + Convert.ToString(dateTimePicker3.Value.Month) + "-" + Convert.ToString(dateTimePicker3.Value.Day) + " 00:00:00.000";


            int number_of_Days = (dateTimePicker3.Value.Date - dateTimePicker2.Value.Date).Days + 1;



            if (Technology == "2G")
            {


                string Data_Quary = @"select [Date], [BSC], substring(Cell,1,6) as 'Site', [Cell], [TCH_Availability] as 'TCH Availability' from [dbo].[CC2_Ericsson_Cell_Daily] where  (" + EH_sites_list + ") and ( Date>='" + start_date + "' and Date<='" + end_date + "')" +
                    @" union all select [Date], [BSC], substring(Cell,1,6) as 'Site', [Cell], [TCH_Availability] as 'TCH Availability' from [dbo].[CC2_Huawei_Cell_Daily] where (" + EH_sites_list + ") and ( Date>='" + start_date + "' and Date<='" + end_date + "')" +
                    @" union all select [Date], [BSC], substring(Cell,1,2)+substring(Cell,5,4) as 'Site', [Cell], [TCH_Availability] as 'TCH Availability' from [dbo].[CC2_Huawei_Cell_Daily] where (" + H_sites_list + ") and ( Date>='" + start_date + "' and Date<='" + end_date + "')" +
                    @" union all select [Date], [BSC], substring(Seg,1,2)+substring(Seg,5,4) as 'Site', [SEG] as 'Cell', [TCH_Availability] as 'TCH Availability' from [dbo].[CC2_Nokia_Cell_Daily] where (" + N_sites_list + ") and ( Date>='" + start_date + "' and Date<='" + end_date + "')";


                SqlCommand Data_Quary1 = new SqlCommand(Data_Quary, connection);
                Data_Quary1.CommandTimeout = 0;
                Data_Quary1.ExecuteNonQuery();
                //    Data_Table_2G = new DataTable();
                SqlDataAdapter Date_Table1 = new SqlDataAdapter(Data_Quary1);
                Date_Table1.Fill(Data_Table_2G);




                var distinct_CellID = Data_Table_2G.AsEnumerable()
.Select(s => new
{
    id = s.Field<string>("Cell"),
})
.Distinct().ToList();



                // Cell Status
                for (int j = 0; j < distinct_CellID.Count; j++)
                {
                    var Cell_Data = (from p in Data_Table_2G.AsEnumerable()
                                     where p.Field<string>("Cell") == distinct_CellID[j].id
                                     select p).ToList();


                    int Non_100_Count = 0;
                    // flag is Cell Status
                    string flag = "";
                    int Num_0_Count = 0;
                    double cell_indicator = 1;
                    string D1 = "";
                    for (int c = 0; c <= Cell_Data.Count - 1; c++)
                    {
                        string Availability_Value = Cell_Data[c].ItemArray[4].ToString();
                        D1 = Cell_Data[c].ItemArray[0].ToString();

                        if (Availability_Value != "")
                        {
                            double Availability_Value_Double = Convert.ToDouble(Availability_Value);
                            if (Availability_Value_Double != 100)
                            {
                                // number of days dose not equal 100
                                Non_100_Count++;
                            }
                            if (Availability_Value_Double == 0)
                            {
                                // number of days equals 0
                                Num_0_Count++;
                            }
                            if (D1 == end_date)
                            {
                                if (Availability_Value_Double == 0)
                                {
                                    flag = "Cell Down";
                                    cell_indicator = 0;
                                }
                            }
                        }
                        if (Availability_Value == "")
                        {
                            if (D1 == end_date)
                            {
                                flag = "Not Updated";
                                cell_indicator = 0.0000000001;
                            }
                        }
                    }


                    if (D1 != dateTimePicker3.Value.Date.ToString())
                    {
                        if (Cell_Data.Count < number_of_Days)
                        {
                            flag = "Not Updated";
                            cell_indicator = 0.0000000001;
                        }
                    }

                    string Date = start_date + "-" + end_date;
                    string Node = Cell_Data[0].ItemArray[1].ToString();
                    string Site = Cell_Data[0].ItemArray[2].ToString();
                    string Cell = Cell_Data[0].ItemArray[3].ToString();
                    int N_Non_100 = Non_100_Count;
                    int N_0 = Num_0_Count;

                    if (flag == "")
                    {
                        if (Non_100_Count >=5 && Non_100_Count < 9)
                        {
                            flag = "Low Fluctuation";
                            cell_indicator = 1.01;
                        }
                        if (Non_100_Count < 5)
                        {
                            flag = "Ok";
                            cell_indicator = 1;
                        }
                        if (Non_100_Count >=9)
                        {
                            flag = "High Fluctuation";
                            cell_indicator = 2;
                        }
                    }

                    Fluctuation_Results.Rows.Add(Date, Node, Site, Cell, N_0, N_Non_100, flag, "", cell_indicator);

                }



                var distinct_SiteID = Fluctuation_Results.AsEnumerable()
.Select(s => new
{
    id = s.Field<string>("Site"),
})
.Distinct().ToList();



                int rr = 1;

                //  string[,] Site_Score = new string[distinct_SiteID.Count, 2];
                double Site_Score1 = 1;  // Multiply
                double Site_Score2 = 0;  // Summation
                for (int c = 0; c <= distinct_SiteID.Count - 1; c++)
                {

                    string site = distinct_SiteID[c].id;



                    var Cell_Data = (from p in Fluctuation_Results.AsEnumerable()
                                     where p.Field<string>("Site") == site
                                     select p).ToList();

                    Site_Score1 = 1;
                    for (int k = 0; k <= Cell_Data.Count - 1; k++)
                    {
                        Site_Score1 = Site_Score1 * Convert.ToDouble(Cell_Data[k].ItemArray[8].ToString());
                        Site_Score2 = Site_Score2 + Convert.ToDouble(Cell_Data[k].ItemArray[8].ToString());
                    }


                    for (int i = 0; i <= Fluctuation_Results.Rows.Count - 1; i++)
                    {
                        string site1 = Fluctuation_Results.Rows[i].ItemArray[2].ToString();

                        if (site == site1)
                        {
                            if (Site_Score1 == 0 && Site_Score2 == 0)
                            {
                                Fluctuation_Results.Rows[i][7] = "Site Down";
                            }
                            if (Site_Score1 == 0 && Site_Score2 != 0)
                            {
                                Fluctuation_Results.Rows[i][7] = "Cell Down";
                            }
                            if (Site_Score1 > 0 && Site_Score1 < 1)
                            {
                                Fluctuation_Results.Rows[i][7] = "Not Updated";
                            }
                            if (Site_Score1 == 1)
                            {
                                Fluctuation_Results.Rows[i][7] = "Ok";
                            }
                            if (Site_Score1 > 1 && Site_Score1 < 2)
                            {
                                Fluctuation_Results.Rows[i][7] = "Low Fluctuation";
                            }
                            if (Site_Score1 >= 2)
                            {
                                Fluctuation_Results.Rows[i][7] = "High Fluctuation";
                            }
                        }


                    }


                }

                this.Invoke(new Action(() => { MessageBox.Show(this, "Finished"); }));

            }




            if (Technology == "3G")
            {


                string Data_Quary = @"select [Date],  [ElementID] as 'RNC', substring([ElementID1],1,8) as 'Site', [ElementID1] as 'Cell', [Cell_Availability_Rate_Exclude_Blocking(UCELL_Eric)] as 'Cell Availability' from [dbo].[CC3_Ericsson_Cell_Daily] where  (" + sites_list + ") and ( Date>='" + start_date + "' and Date<='" + end_date + "')" +
                    @" union all select [Date], [ElementID] as 'RNC',  substring([ElementID1],1,8) as 'Site', [ElementID1] as 'Cell', [Radio_Network_Availability_Ratio(Hu_Cell)] as 'Cell Availability' from [dbo].[CC3_Huawei_Cell_Daily] where (" + sites_list + ") and ( Date>='" + start_date + "' and Date<='" + end_date + "')" +
                    @" union all select [Date],  [ElementID] as 'RNC',  substring([ElementID1],1,8) as 'Site', [ElementID1] as 'Cell', [Cell_Availability_excluding_blocked_by_user_state] as 'Cell Availability' from [dbo].[CC3_Nokia_Cell_Daily] where (" + sites_list + ") and ( Date>='" + start_date + "' and Date<='" + end_date + "')";



                SqlCommand Data_Quary1 = new SqlCommand(Data_Quary, connection);
                Data_Quary1.CommandTimeout = 0;
                Data_Quary1.ExecuteNonQuery();
                //    Data_Table_3G = new DataTable();
                SqlDataAdapter Date_Table1 = new SqlDataAdapter(Data_Quary1);
                Date_Table1.Fill(Data_Table_3G);




                var distinct_CellID = Data_Table_3G.AsEnumerable()
.Select(s => new
{
    id = s.Field<string>("Cell"),
})
.Distinct().ToList();



                // Cell Status
                for (int j = 0; j < distinct_CellID.Count; j++)
                {
                    var Cell_Data = (from p in Data_Table_3G.AsEnumerable()
                                     where p.Field<string>("Cell") == distinct_CellID[j].id
                                     select p).ToList();


                    int Non_100_Count = 0;
                    // flag is Cell Status
                    string flag = "";
                    int Num_0_Count = 0;
                    double cell_indicator = 1;
                    string D1 = "";
                    for (int c = 0; c <= Cell_Data.Count - 1; c++)
                    {
                        string Availability_Value = Cell_Data[c].ItemArray[4].ToString();
                        D1 = Cell_Data[c].ItemArray[0].ToString();

                        if (Availability_Value != "")
                        {
                            double Availability_Value_Double = Convert.ToDouble(Availability_Value);
                            if (Availability_Value_Double != 100)
                            {
                                // number of days dose not equal 100
                                Non_100_Count++;
                            }
                            if (Availability_Value_Double == 0)
                            {
                                // number of days equals 0
                                Num_0_Count++;
                            }
                            if (D1 == end_date)
                            {
                                if (Availability_Value_Double == 0)
                                {
                                    flag = "Cell Down";
                                    cell_indicator = 0;
                                }
                            }
                        }
                        if (Availability_Value == "")
                        {
                            if (D1 == end_date)
                            {
                                flag = "Not Updated";
                                cell_indicator = 0.0000000001;
                            }
                        }
                    }


                    if (D1 != dateTimePicker3.Value.Date.ToString())
                    {
                        if (Cell_Data.Count < number_of_Days)
                        {
                            flag = "Not Updated";
                            cell_indicator = 0.0000000001;
                        }
                    }


                    string Date = start_date + "-" + end_date;
                    string Node = Cell_Data[0].ItemArray[1].ToString();
                    string Site = Cell_Data[0].ItemArray[2].ToString();
                    string Cell = Cell_Data[0].ItemArray[3].ToString();
                    int N_Non_100 = Non_100_Count;
                    int N_0 = Num_0_Count;

                    if (flag == "")
                    {
                        if (Non_100_Count >= 5 && Non_100_Count < 9)
                        {
                            flag = "Low Fluctuation";
                            cell_indicator = 1.01;
                        }
                        if (Non_100_Count < 5)
                        {
                            flag = "Ok";
                            cell_indicator = 1;
                        }
                        if (Non_100_Count >= 9)
                        {
                            flag = "High Fluctuation";
                            cell_indicator = 2;
                        }
                    }

                    Fluctuation_Results.Rows.Add(Date, Node, Site, Cell, N_0, N_Non_100, flag, "", cell_indicator);

                }



                var distinct_SiteID = Fluctuation_Results.AsEnumerable()
.Select(s => new
{
    id = s.Field<string>("Site"),
})
.Distinct().ToList();



                int rr = 1;

                //  string[,] Site_Score = new string[distinct_SiteID.Count, 2];
                double Site_Score1 = 1;  // Multiply
                double Site_Score2 = 0;  // Summation
                for (int c = 0; c <= distinct_SiteID.Count - 1; c++)
                {

                    string site = distinct_SiteID[c].id;



                    var Cell_Data = (from p in Fluctuation_Results.AsEnumerable()
                                     where p.Field<string>("Site") == site
                                     select p).ToList();

                    Site_Score1 = 1;
                    for (int k = 0; k <= Cell_Data.Count - 1; k++)
                    {
                        Site_Score1 = Site_Score1 * Convert.ToDouble(Cell_Data[k].ItemArray[8].ToString());
                        Site_Score2 = Site_Score2 + Convert.ToDouble(Cell_Data[k].ItemArray[8].ToString());
                    }


                    for (int i = 0; i <= Fluctuation_Results.Rows.Count - 1; i++)
                    {
                        string site1 = Fluctuation_Results.Rows[i].ItemArray[2].ToString();

                        if (site == site1)
                        {
                            if (Site_Score1 == 0 && Site_Score2 == 0)
                            {
                                Fluctuation_Results.Rows[i][7] = "Site Down";
                            }
                            if (Site_Score1 == 0 && Site_Score2 != 0)
                            {
                                Fluctuation_Results.Rows[i][7] = "Cell Down";
                            }
                            if (Site_Score1 > 0 && Site_Score1 < 1)
                            {
                                Fluctuation_Results.Rows[i][7] = "Not Updated";
                            }
                            if (Site_Score1 == 1)
                            {
                                Fluctuation_Results.Rows[i][7] = "Ok";
                            }
                            if (Site_Score1 > 1 && Site_Score1 < 2)
                            {
                                Fluctuation_Results.Rows[i][7] = "Low Fluctuation";
                            }
                            if (Site_Score1 >= 2)
                            {
                                Fluctuation_Results.Rows[i][7] = "High Fluctuation";
                            }
                        }


                    }


                }

                this.Invoke(new Action(() => { MessageBox.Show(this, "Finished"); }));

            }




            if (Technology == "4G")
            {


                string Data_Quary = @"select [Datetime], substring([eNodeB],1,8) as 'Site', [eNodeB] as 'Cell', [Cell_Availability_Rate_Exclude_Blocking(Cell_EricLTE)] as 'Cell Availability' from [dbo].[TBL_LTE_CELL_Daily_E] where  (" + EH_sites_list + ") and ( Datetime>='" + start_date + "' and Datetime<='" + end_date + "')" +
                    @" union all select [Datetime],  substring([eNodeB],1,8) as 'Site', [eNodeB] as 'Cell', [Cell_Availability_Rate_Exclude_Blocking(Cell_Hu)] as 'Cell Availability' from [dbo].[TBL_LTE_CELL_Daily_H]  where (" + EH_sites_list + ") and ( Datetime>='" + start_date + "' and Datetime<='" + end_date + "')" +
                    @" union all select [Date],  substring([ElementID1],1,8) as 'Site', [ElementID1] as 'Cell', [cell_availability_exclude_manual_blocking(Nokia_LTE_CELL)] as 'Cell Availability' from [dbo].[TBL_LTE_CELL_Daily_N] where (" + N_sites_list + ") and ( Date>='" + start_date + "' and Date<='" + end_date + "')";




                SqlCommand Data_Quary1 = new SqlCommand(Data_Quary, connection);
                Data_Quary1.CommandTimeout = 0;
                Data_Quary1.ExecuteNonQuery();
                //    Data_Table_4G = new DataTable();
                SqlDataAdapter Date_Table1 = new SqlDataAdapter(Data_Quary1);
                Date_Table1.Fill(Data_Table_4G);




                var distinct_CellID = Data_Table_4G.AsEnumerable()
.Select(s => new
{
    id = s.Field<string>("Cell"),
})
.Distinct().ToList();



                // Cell Status
                for (int j = 0; j < distinct_CellID.Count; j++)
                {
                    var Cell_Data = (from p in Data_Table_4G.AsEnumerable()
                                     where p.Field<string>("Cell") == distinct_CellID[j].id
                                     select p).ToList();


                    int Non_100_Count = 0;
                    // flag is Cell Status
                    string flag = "";
                    int Num_0_Count = 0;
                    double cell_indicator = 1;
                    string D1 = "";
                    for (int c = 0; c <= Cell_Data.Count - 1; c++)
                    {
                        string Availability_Value = Cell_Data[c].ItemArray[3].ToString();
                        D1 = Cell_Data[c].ItemArray[0].ToString();

                        if (Availability_Value != "")
                        {
                            double Availability_Value_Double = Convert.ToDouble(Availability_Value);
                            if (Availability_Value_Double != 100)
                            {
                                // number of days dose not equal 100
                                Non_100_Count++;
                            }
                            if (Availability_Value_Double == 0)
                            {
                                // number of days equals 0
                                Num_0_Count++;
                            }
                            if (D1 == end_date)
                            {
                                if (Availability_Value_Double == 0)
                                {
                                    flag = "Cell Down";
                                    cell_indicator = 0;
                                }
                            }
                        }
                        if (Availability_Value == "")
                        {
                            if (D1 == end_date)
                            {
                                flag = "Not Updated";
                                cell_indicator = 0.0000000001;
                            }
                        }
                    }

                    if (D1 != dateTimePicker3.Value.Date.ToString())
                    {
                        if (Cell_Data.Count < number_of_Days)
                        {
                            flag = "Not Updated";
                            cell_indicator = 0.0000000001;
                        }
                    }


                    string Date = start_date + "-" + end_date;
                    string Node = "";
                    string Site = Cell_Data[0].ItemArray[1].ToString();
                    string Cell = Cell_Data[0].ItemArray[2].ToString();
                    int N_Non_100 = Non_100_Count;
                    int N_0 = Num_0_Count;

                    if (flag == "")
                    {
                        if (Non_100_Count >= 5 && Non_100_Count < 9)
                        {
                            flag = "Low Fluctuation";
                            cell_indicator = 1.01;
                        }
                        if (Non_100_Count < 5)
                        {
                            flag = "Ok";
                            cell_indicator = 1;
                        }
                        if (Non_100_Count >= 9)
                        {
                            flag = "High Fluctuation";
                            cell_indicator = 2;
                        }
                    }

                    Fluctuation_Results.Rows.Add(Date, Node, Site, Cell, N_0, N_Non_100, flag, "", cell_indicator);

                }



                var distinct_SiteID = Fluctuation_Results.AsEnumerable()
.Select(s => new
{
    id = s.Field<string>("Site"),
})
.Distinct().ToList();



                int rr = 1;

                //  string[,] Site_Score = new string[distinct_SiteID.Count, 2];
                double Site_Score1 = 1;  // Multiply
                double Site_Score2 = 0;  // Summation
                for (int c = 0; c <= distinct_SiteID.Count - 1; c++)
                {

                    string site = distinct_SiteID[c].id;



                    var Cell_Data = (from p in Fluctuation_Results.AsEnumerable()
                                     where p.Field<string>("Site") == site
                                     select p).ToList();

                    Site_Score1 = 1;
                    for (int k = 0; k <= Cell_Data.Count - 1; k++)
                    {
                        Site_Score1 = Site_Score1 * Convert.ToDouble(Cell_Data[k].ItemArray[8].ToString());
                        Site_Score2 = Site_Score2 + Convert.ToDouble(Cell_Data[k].ItemArray[8].ToString());
                    }


                    for (int i = 0; i <= Fluctuation_Results.Rows.Count - 1; i++)
                    {
                        string site1 = Fluctuation_Results.Rows[i].ItemArray[2].ToString();

                        if (site == site1)
                        {
                            if (Site_Score1 == 0 && Site_Score2 == 0)
                            {
                                Fluctuation_Results.Rows[i][7] = "Site Down";
                            }
                            if (Site_Score1 == 0 && Site_Score2 != 0)
                            {
                                Fluctuation_Results.Rows[i][7] = "Cell Down";
                            }
                            if (Site_Score1 > 0 && Site_Score1 < 1)
                            {
                                Fluctuation_Results.Rows[i][7] = "Not Updated";
                            }
                            if (Site_Score1 == 1)
                            {
                                Fluctuation_Results.Rows[i][7] = "Ok";
                            }
                            if (Site_Score1 > 1 && Site_Score1 < 2)
                            {
                                Fluctuation_Results.Rows[i][7] = "Low Fluctuation";
                            }
                            if (Site_Score1 >= 2)
                            {
                                Fluctuation_Results.Rows[i][7] = "High Fluctuation";
                            }
                        }


                    }


                }

                this.Invoke(new Action(() => { MessageBox.Show(this, "Finished"); }));

            }
        }




        void my_thread3()
        {



        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                checkBox2.Checked = false; checkBox5.Checked = false; Technology = "2G";
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                checkBox1.Checked = false; checkBox5.Checked = false; Technology = "3G";
            }
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked == true)
            {
                checkBox1.Checked = false; checkBox2.Checked = false; Technology = "4G";
            }
        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox6.Checked == true)
            {
                Input_Type = "DataBase";
                checkBox7.Checked = false;
            }
        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox7.Checked == true)
            {
                Input_Type = "FARAZ";
                checkBox6.Checked = false;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Task = "Fluctuation";


            string start_date = Convert.ToString(dateTimePicker2.Value.Date);
            string end_date = Convert.ToString(dateTimePicker3.Value.Date);
            //string start_date = Convert.ToString(dateTimePicker2.Value.Year) + "-" + Convert.ToString(dateTimePicker2.Value.Month) + "-" + Convert.ToString(dateTimePicker2.Value.Day) + " 00:00:00.000";
            //string end_date = Convert.ToString(dateTimePicker3.Value.Year) + "-" + Convert.ToString(dateTimePicker3.Value.Month) + "-" + Convert.ToString(dateTimePicker3.Value.Day) + " 00:00:00.000";


            int number_of_Days = (dateTimePicker3.Value.Date - dateTimePicker2.Value.Date).Days + 1;




            if (Technology == "2G" && Input_Type == "FARAZ")
            {




                Fluctuation_Results.Columns.Add("Date", typeof(string));
                Fluctuation_Results.Columns.Add("Node", typeof(string));
                Fluctuation_Results.Columns.Add("Site", typeof(string));
                Fluctuation_Results.Columns.Add("Cell", typeof(string));
                Fluctuation_Results.Columns.Add("Zero Cell Availability (No of Days)", typeof(int));
                Fluctuation_Results.Columns.Add("Cell Availability less than 100% (No of Days)", typeof(int));
                Fluctuation_Results.Columns.Add("Cell Score", typeof(string));
                Fluctuation_Results.Columns.Add("Site Score", typeof(string));
                Fluctuation_Results.Columns.Add("Cell Score Indicator", typeof(double));



                Data_Table_2G.Columns.Add("Date", typeof(DateTime));
                Data_Table_2G.Columns.Add("BSC", typeof(string));
                Data_Table_2G.Columns.Add("Site", typeof(string));
                Data_Table_2G.Columns.Add("Cell", typeof(string));
                Data_Table_2G.Columns.Add("TCH Availability", typeof(string));








                openFileDialog2.DefaultExt = "xlsx";
                openFileDialog2.Filter = "Excel File|*.xlsx";
                DialogResult result = openFileDialog2.ShowDialog();
                string File_Name = openFileDialog2.SafeFileName.ToString();
                if (result == DialogResult.OK)
                {
                    string file = openFileDialog2.FileName;
                    xlApp = new Excel.Application();
                    KPI_workbook = xlApp.Workbooks.Open(file);
                    Excel.Worksheet Sheet = KPI_workbook.Worksheets[1];


                    Excel.Range Range1 = Sheet.get_Range("A2", "H" + Sheet.UsedRange.Rows.Count);
                    object[,] FARAZ_Data = (object[,])Range1.Value;
                    int Count = Sheet.UsedRange.Rows.Count;

                    for (int k = 0; k < Count - 1; k++)
                    {
                        DateTime Date = Convert.ToDateTime(FARAZ_Data[k + 1, 1]);
                        string NE = FARAZ_Data[k + 1, 2].ToString();
                        string Site = "";

                        string str2 = Regex.Replace(NE, "[^a-zA-Z0-9]", " ");      //هر کاراکتری که غیر از عدد و حرف بود را به کاراکتر خالی تبدیل کن
                        str2 = Regex.Replace(str2, " {2,}", " ").Trim();           //چندین کاراکتر خالی پشت سر هم را به یک کاراکتر خالی تبدیل می کند
                        string[] Split_Description = str2.Split(' ');
                        string BSC = Split_Description[0];

                        string Tech = Split_Description[0].Substring(0, 1);
                        string Tech_Last = Split_Description[0].Substring(Split_Description[0].Length - 1, 1);

                        string CellName = "";
                        if ((Tech == "B" && (Tech_Last == "E" || Tech_Last == "H" || Tech_Last == "N")) || Split_Description[0].Length == 2)
                        {
                            //int c1 = Split_Description.Length;
                            CellName = Split_Description[1];
                        }
                        else
                        {
                            int c1 = Split_Description.Length;
                            CellName = Split_Description[c1 - 1];
                        }


                        if (CellName.Length == 7)
                        {
                            Site = CellName.Substring(0, 6);
                        }
                        if (CellName.Length > 7)
                        {
                            Site = CellName.Substring(0, 8);
                        }



                        string Availability = "";
                        if (FARAZ_Data[k + 1, 3] != null)
                        {
                            Availability = FARAZ_Data[k + 1, 3].ToString();
                        }
                        else
                        {
                            Availability = "";
                        }



                        Data_Table_2G.Rows.Add(Date, BSC, Site, CellName, Availability);


                    }




                    var distinct_CellID = Data_Table_2G.AsEnumerable()
    .Select(s => new
    {
        id = s.Field<string>("Cell"),
    })
    .Distinct().ToList();



                    // Cell Status
                    for (int j = 0; j < distinct_CellID.Count; j++)
                    {
                        var Cell_Data = (from p in Data_Table_2G.AsEnumerable()
                                         where p.Field<string>("Cell") == distinct_CellID[j].id
                                         select p).ToList();


                        int Non_100_Count = 0;
                        // flag is Cell Status
                        string flag = "";
                        int Num_0_Count = 0;
                        double cell_indicator = 1;
                        string D1 = "";
                        for (int c = 0; c <= Cell_Data.Count - 1; c++)
                        {
                            string Availability_Value = Cell_Data[c].ItemArray[4].ToString();
                            D1 = Cell_Data[c].ItemArray[0].ToString();

                            if (Availability_Value != "")
                            {
                                double Availability_Value_Double = Convert.ToDouble(Availability_Value);
                                if (Availability_Value_Double != 100)
                                {
                                    // number of days dose not equal 100
                                    Non_100_Count++;
                                }
                                if (Availability_Value_Double == 0)
                                {
                                    // number of days equals 0
                                    Num_0_Count++;
                                }
                                if (D1 == end_date)
                                {
                                    if (Availability_Value_Double == 0)
                                    {
                                        flag = "Cell Down";
                                        cell_indicator = 0;
                                    }
                                }
                            }
                            if (Availability_Value == "")
                            {
                                if (D1 == end_date)
                                {
                                    flag = "Not Updated";
                                    cell_indicator = 0.0000000001;
                                }
                            }
                        }


                        if (D1 != dateTimePicker3.Value.Date.ToString())
                        {
                            if (Cell_Data.Count < number_of_Days)
                            {
                                flag = "Not Updated";
                                cell_indicator = 0.0000000001;
                            }
                        }

                        string Date = start_date + "-" + end_date;
                        string Node = Cell_Data[0].ItemArray[1].ToString();
                        string Site = Cell_Data[0].ItemArray[2].ToString();
                        string Cell = Cell_Data[0].ItemArray[3].ToString();
                        int N_Non_100 = Non_100_Count;
                        int N_0 = Num_0_Count;

                        if (flag == "")
                        {
                            if (Non_100_Count >= 5 && Non_100_Count < 9)
                            {
                                flag = "Low Fluctuation";
                                cell_indicator = 1.01;
                            }
                            if (Non_100_Count < 5)
                            {
                                flag = "Ok";
                                cell_indicator = 1;
                            }
                            if (Non_100_Count >= 9)
                            {
                                flag = "High Fluctuation";
                                cell_indicator = 2;
                            }

                        }

                        Fluctuation_Results.Rows.Add(Date, Node, Site, Cell, N_0, N_Non_100, flag, "", cell_indicator);

                    }



                    var distinct_SiteID = Fluctuation_Results.AsEnumerable()
    .Select(s => new
    {
        id = s.Field<string>("Site"),
    })
    .Distinct().ToList();



                    int rr = 1;

                    //  string[,] Site_Score = new string[distinct_SiteID.Count, 2];
                    double Site_Score1 = 1;  // Multiply
                    double Site_Score2 = 0;  // Summation
                    for (int c = 0; c <= distinct_SiteID.Count - 1; c++)
                    {

                        string site = distinct_SiteID[c].id;



                        var Cell_Data = (from p in Fluctuation_Results.AsEnumerable()
                                         where p.Field<string>("Site") == site
                                         select p).ToList();

                        Site_Score1 = 1;
                        for (int k = 0; k <= Cell_Data.Count - 1; k++)
                        {
                            Site_Score1 = Site_Score1 * Convert.ToDouble(Cell_Data[k].ItemArray[8].ToString());
                            Site_Score2 = Site_Score2 + Convert.ToDouble(Cell_Data[k].ItemArray[8].ToString());
                        }


                        for (int i = 0; i <= Fluctuation_Results.Rows.Count - 1; i++)
                        {
                            string site1 = Fluctuation_Results.Rows[i].ItemArray[2].ToString();

                            if (site == site1)
                            {
                                if (Site_Score1 == 0 && Site_Score2 == 0)
                                {
                                    Fluctuation_Results.Rows[i][7] = "Site Down";
                                }
                                if (Site_Score1 == 0 && Site_Score2 != 0)
                                {
                                    Fluctuation_Results.Rows[i][7] = "Cell Down";
                                }
                                if (Site_Score1 > 0 && Site_Score1 < 1)
                                {
                                    Fluctuation_Results.Rows[i][7] = "Not Updated";
                                }
                                if (Site_Score1 == 1)
                                {
                                    Fluctuation_Results.Rows[i][7] = "Ok";
                                }
                                if (Site_Score1 > 1 && Site_Score1 < 2)
                                {
                                    Fluctuation_Results.Rows[i][7] = "Low Fluctuation";
                                }
                                if (Site_Score1 >= 2)
                                {
                                    Fluctuation_Results.Rows[i][7] = "High Fluctuation";
                                }
                            }


                        }


                    }

                }


                MessageBox.Show("Finished");

            }
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {

        }
    }
}
