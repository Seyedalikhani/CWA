﻿using System;
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
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;


namespace CWA
{
    public partial class Form6 : Form
    {
        public Form6()
        {
            InitializeComponent();
        }


        public Form1 form1;


        public Form6(Form form)
        {
            InitializeComponent();
            form1 = (Form1)form;
        }


        public string ConnectionString = "";
        public SqlConnection connection = new SqlConnection();
        //public string Server_Name = "172.26.7.159";
        public string Server_Name = "PERFORMANCEDB01";
        public string DataBase_Name = "Performance_NAK";
        public string Technology = "";
        public DataTable Data_Table_2G = new DataTable();
        public DataTable Site_Data_Table_2G = new DataTable();
        public DataTable Data_Table_3G_CS = new DataTable();
        public DataTable Data_Table_3G_PS = new DataTable();
        public DataTable Data_Table_3G = new DataTable();
        public DataTable Site_Data_Table_3G = new DataTable();
        public DataTable Data_Table_4G = new DataTable();
        public DataTable Site_Data_Table_4G = new DataTable();
        public string Input_Type = "DataBase";
        public string FName = "";
        public IXLWorksheet Source_worksheet = null;
        //public DataTable Data_Table_2G = new DataTable();
        public Excel.Application xlApp { get; set; }
        public Excel.Workbook xlWorkBook { get; set; }
        public Excel.Worksheet Sheet { get; set; }




        private void Form6_Load(object sender, EventArgs e)
        {


            //string Server_Name = @"NAKPRG-NB1243\" + "AHMAD";
            //string DataBase_Name = "Dashboards";

            //ConnectionString = @"Server=" + Server_Name + "; Database=" + DataBase_Name + "; Trusted_Connection=True;";

            ConnectionString = @"Server=" + Server_Name + "; Database=" + DataBase_Name + "; User ID=cwpcApp;Password=cwpcApp@830625#Ahmad";
            connection = new SqlConnection(ConnectionString);
            connection.Open();


        }






        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.SelectAll();

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

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            string add_string_month = "";
            if (dateTimePicker1.Value.Month == 1 || dateTimePicker1.Value.Month == 2 || dateTimePicker1.Value.Month == 3 || dateTimePicker1.Value.Month == 4 || dateTimePicker1.Value.Month == 5 || dateTimePicker1.Value.Month == 6 || dateTimePicker1.Value.Month == 7 || dateTimePicker1.Value.Month == 8 || dateTimePicker1.Value.Month == 9)
            {
                add_string_month = "0";
            }
            string add_string_day = "";
            if (dateTimePicker1.Value.Day == 1 || dateTimePicker1.Value.Day == 2 || dateTimePicker1.Value.Day == 3 || dateTimePicker1.Value.Day == 4 || dateTimePicker1.Value.Day == 5 || dateTimePicker1.Value.Day == 6 || dateTimePicker1.Value.Day == 7 || dateTimePicker1.Value.Day == 8 || dateTimePicker1.Value.Day == 9)
            {
                add_string_day = "0";
            }
            listBox1.Items.Add(dateTimePicker1.Value.DayOfWeek + " " + Convert.ToString(dateTimePicker1.Value.Year) + "-" + add_string_month + Convert.ToString(dateTimePicker1.Value.Month) + "-" + add_string_day + Convert.ToString(dateTimePicker1.Value.Day));
        }

        private void button3_Click(object sender, EventArgs e)
        {
            for (int i = listBox1.SelectedItems.Count - 1; i >= 0; i--)
                listBox1.Items.Remove(listBox1.SelectedItems[i]);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            label43.Text = "Please wait while It is running";
            label43.BackColor = Color.GreenYellow;
            Thread th1 = new Thread(my_thread1);
            th1.Start();

        }


        void my_thread1()
        {

            //if (Input_Type == "DataBase")
            //{



            //    string site_list = textBox1.Text;
            //    string Site_Disticnts1 = Regex.Replace(site_list, "[^a-zA-Z0-9]", " ");
            //    Site_Disticnts1 = Regex.Replace(Site_Disticnts1, " {2,}", " ").Trim();
            //    string[] Site_Code_Distincts = Site_Disticnts1.Split(' ');

            //    string EH_sites_list_2G = "";
            //    string H_sites_list_2G = "";
            //    string N_sites_list_2G = "";
            //    string sites_list_3G = "";
            //    string EH_sites_list_4G = "";
            //    string N_sites_list_4G = "";

            //    int Data_in_2G = 0;
            //    int Data_in_3G = 0;
            //    int Data_in_4G = 0;

            //    for (int i = 0; i < Site_Code_Distincts.Count(); i++)
            //    {


            //        string Band_Letter = Site_Code_Distincts[i].Substring(3,1);

            //        if (Band_Letter=="g" || Band_Letter == "G" || Site_Code_Distincts[i].Length==6)
            //        {
            //            Technology = "2G";
            //            Data_in_2G++;
            //        }
            //        if (Band_Letter == "u" || Band_Letter == "U")
            //        {
            //            Technology = "3G";
            //            Data_in_3G++;
            //        }

            //        if (Band_Letter == "l" || Band_Letter == "L" )
            //        {
            //            Technology = "4G";
            //            Data_in_4G++;
            //        }

            //        if (Technology == "2G")
            //        {
            //            string Letter_6_Site = Site_Code_Distincts[i];
            //            if (Site_Code_Distincts[i].Length>6)
            //            {
            //                Letter_6_Site = Site_Code_Distincts[i].Substring(0, 2) + Site_Code_Distincts[i].Substring(4, 4);
            //            }
            //            EH_sites_list_2G = EH_sites_list_2G + "substring([Cell],1,6)='" + Letter_6_Site + "' or ";
            //            H_sites_list_2G = H_sites_list_2G + "substring([Cell],1,2)+substring([Cell],5,4)='" + Letter_6_Site + "' or ";
            //            N_sites_list_2G = N_sites_list_2G + "substring([Seg],1,2)+substring([Seg],5,4)='" + Letter_6_Site + "' or ";
            //        }
            //        if (Technology == "3G" || Technology == "3G-MCI")
            //        {
            //            sites_list_3G = sites_list_3G + "substring([ElementID1],1,8)='" + Site_Code_Distincts[i] + "' or ";
            //        }
            //        if (Technology == "4G")
            //        {
            //            EH_sites_list_4G = EH_sites_list_4G + "substring([eNodeB],1,8)='" + Site_Code_Distincts[i] + "' or ";
            //            N_sites_list_4G = N_sites_list_4G + "substring([ElementID1],1,8)='" + Site_Code_Distincts[i] + "' or ";
            //        }

            //    }
            //    if (Data_in_2G != 0)
            //    {
            //        EH_sites_list_2G = EH_sites_list_2G.Substring(0, EH_sites_list_2G.Length - 4);
            //        H_sites_list_2G = H_sites_list_2G.Substring(0, H_sites_list_2G.Length - 4);
            //        N_sites_list_2G = N_sites_list_2G.Substring(0, N_sites_list_2G.Length - 4);
            //    }
            //    if (Data_in_3G != 0)
            //    {
            //        sites_list_3G = sites_list_3G.Substring(0, sites_list_3G.Length - 4);
            //    }
            //    if (Data_in_4G != 0)
            //    {
            //        EH_sites_list_4G = EH_sites_list_4G.Substring(0, EH_sites_list_4G.Length - 4);
            //        N_sites_list_4G = N_sites_list_4G.Substring(0, N_sites_list_4G.Length - 4);
            //    }

            //    string date_list = "";
            //    string EH_date_list = "";
            //    string N_date_list = "";
            //    for (int i = 0; i < listBox1.Items.Count; i++)
            //    {
            //        string date_list1 = listBox1.Items[i].ToString();
            //        int space_index = 0;
            //        for (int k = 0; k < date_list1.Length; k++)
            //        {
            //            if (date_list1[k].ToString() == " ")
            //            {
            //                space_index = k;
            //                break;
            //            }
            //        }
            //        string Day = date_list1.Substring(space_index + 1, date_list1.Length - space_index - 1);
            //        //if (Technology != "4G")
            //        //{
            //            date_list = date_list + "substring(convert(varchar, Date, 23), 1, 10) = '" + Day + "' or ";
            //        //}
            //        //if (Technology == "4G")
            //        //{
            //            EH_date_list = EH_date_list + "substring(convert(varchar, Datetime, 23), 1, 10) = '" + Day + "' or ";
            //            N_date_list = N_date_list + "substring(convert(varchar, Date, 23), 1, 10) = '" + Day + "' or ";
            //       // }
            //    }
            //    //if (Technology != "4G")
            //    //{
            //        date_list = date_list.Substring(0, date_list.Length - 4);
            //    //}
            //    //if (Technology == "4G")
            //    //{
            //        EH_date_list = EH_date_list.Substring(0, EH_date_list.Length - 4);
            //        N_date_list = N_date_list.Substring(0, N_date_list.Length - 4);
            //   // }




            //    if (Data_in_2G != 0)
            //    {

            //        string Data_Quary = @"select [Date], [BSC], [Cell], [TCH_Traffic_BH] as 'TCH_Traffic_BH (Erlang)', [TCH_Availability] as 'TCH Availability' from [dbo].[CC2_Ericsson_Cell_BH] where  (" + EH_sites_list_2G + ") and (" + date_list + ")" +
            //            @" union all select [Date], [BSC], [Cell], [TCH_Traffic_BH] as 'TCH_Traffic_BH (Erlang)', [TCH_Availability] as 'TCH Availability' from [dbo].[CC2_Huawei_Cell_BH] where (" + EH_sites_list_2G + ") and (" + date_list + ")" +
            //            @" union all select [Date], [BSC], [Cell], [TCH_Traffic_BH] as 'TCH_Traffic_BH (Erlang)', [TCH_Availability] as 'TCH Availability' from [dbo].[CC2_Huawei_Cell_BH] where (" + H_sites_list_2G + ") and (" + date_list + ")" +
            //            @" union all select [Date], [BSC], [SEG] as 'Cell', [TCH_Traffic_BH] as 'TCH_Traffic_BH (Erlang)', [TCH_Availability] as 'TCH Availability' from [dbo].[CC2_Nokia_Cell_BH] where (" + N_sites_list_2G + ") and (" + date_list + ")";

            //        SqlCommand Data_Quary1 = new SqlCommand(Data_Quary, connection);
            //        Data_Quary1.CommandTimeout = 0;
            //        Data_Quary1.ExecuteNonQuery();
            //        Data_Table_2G = new DataTable();
            //        SqlDataAdapter Date_Table1 = new SqlDataAdapter(Data_Quary1);
            //        Date_Table1.Fill(Data_Table_2G);

            //        Data_Table_2G.Columns.Add("Traffic Score", typeof(int));
            //        Data_Table_2G.Columns.Add("Availability Score", typeof(int));
            //        Data_Table_2G.Columns.Add("Cell Score", typeof(double));
            //        // Data_Table_2G.Columns.Add("Site Score", typeof(double));
            //        Data_Table_2G.Columns.Add("Site", typeof(string));


            //        Site_Data_Table_2G = new DataTable();
            //        Site_Data_Table_2G.Columns.Add("Site", typeof(string));
            //        Site_Data_Table_2G.Columns.Add("Pre KPI Zero Status", typeof(string));
            //        Site_Data_Table_2G.Columns.Add("Rejected Cell List", typeof(string));






            //    }







            //    }


            if (Input_Type == "DataBase")
            {
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
                    if (Technology == "3G" || Technology == "3G-MCI")
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
                if (Technology == "3G" || Technology == "3G-MCI")
                {
                    sites_list = sites_list.Substring(0, sites_list.Length - 4);
                }
                if (Technology == "4G")
                {
                    EH_sites_list = EH_sites_list.Substring(0, EH_sites_list.Length - 4);
                    N_sites_list = N_sites_list.Substring(0, N_sites_list.Length - 4);
                }

                string date_list = "";
                string EH_date_list = "";
                string N_date_list = "";
                for (int i = 0; i < listBox1.Items.Count; i++)
                {
                    string date_list1 = listBox1.Items[i].ToString();
                    int space_index = 0;
                    for (int k = 0; k < date_list1.Length; k++)
                    {
                        if (date_list1[k].ToString() == " ")
                        {
                            space_index = k;
                            break;
                        }
                    }
                    string Day = date_list1.Substring(space_index + 1, date_list1.Length - space_index - 1);
                    if (Technology != "4G")
                    {
                        date_list = date_list + "substring(convert(varchar, Date, 23), 1, 10) = '" + Day + "' or ";
                    }
                    if (Technology == "4G")
                    {
                        EH_date_list = EH_date_list + "substring(convert(varchar, Datetime, 23), 1, 10) = '" + Day + "' or ";
                        N_date_list = N_date_list + "substring(convert(varchar, Date, 23), 1, 10) = '" + Day + "' or ";
                    }
                }
                if (Technology != "4G")
                {
                    date_list = date_list.Substring(0, date_list.Length - 4);
                }
                if (Technology == "4G")
                {
                    EH_date_list = EH_date_list.Substring(0, EH_date_list.Length - 4);
                    N_date_list = N_date_list.Substring(0, N_date_list.Length - 4);
                }


                if (Technology == "2G")
                {

                    string Data_Quary = @"select [Date], [BSC], [Cell], [TCH_Traffic_BH] as 'TCH_Traffic_BH (Erlang)', [CSSR_MCI] as'CSSR', [CDR(not Affected by incoming Handovers from 3G)(Eric_CELL)] as 'Voice Drop Rate', [IHSR] as 'IHSR', [OHSR] as 'OHSR', [TCH_Availability] as 'TCH Availability' from [dbo].[CC2_Ericsson_Cell_BH] where  (" + EH_sites_list + ") and (" + date_list + ")" +
                        @" union all select [Date], [BSC], [Cell], [TCH_Traffic_BH] as 'TCH_Traffic_BH (Erlang)', [CSSR3] as'CSSR', [CDR3] as 'Voice Drop Rate', [IHSR2] as 'IHSR', [OHSR2] as 'OHSR',  [TCH_Availability] as 'TCH Availability' from [dbo].[CC2_Huawei_Cell_BH] where (" + EH_sites_list + ") and (" + date_list + ")" +
                        @" union all select [Date], [BSC], [Cell], [TCH_Traffic_BH] as 'TCH_Traffic_BH (Erlang)', [CSSR3] as'CSSR', [CDR3] as 'Voice Drop Rate', [IHSR2] as 'IHSR', [OHSR2] as 'OHSR',  [TCH_Availability] as 'TCH Availability' from [dbo].[CC2_Huawei_Cell_BH] where (" + H_sites_list + ") and (" + date_list + ")" +
                        @" union all select [Date], [BSC], [SEG] as 'Cell', [TCH_Traffic_BH] as 'TCH_Traffic_BH (Erlang)', [CSSR_MCI] as'CSSR', [CDR(including_CS_IRAT_handovers_3G_to2G)(Nokia_SEG)] as 'Voicde  Drop Rate', [IHSR] as 'IHSR', [OHSR] AS 'OHSR', [TCH_Availability] as 'TCH Availability' from [dbo].[CC2_Nokia_Cell_BH] where (" + N_sites_list + ") and (" + date_list + ")";

                    SqlCommand Data_Quary1 = new SqlCommand(Data_Quary, connection);
                    Data_Quary1.CommandTimeout = 0;
                    Data_Quary1.ExecuteNonQuery();
                    Data_Table_2G = new DataTable();
                    SqlDataAdapter Date_Table1 = new SqlDataAdapter(Data_Quary1);
                    Date_Table1.Fill(Data_Table_2G);

                    Data_Table_2G.Columns.Add("Traffic Score", typeof(int));
                    Data_Table_2G.Columns.Add("CSSR Score", typeof(int));
                    Data_Table_2G.Columns.Add("CDR Score", typeof(int));
                    Data_Table_2G.Columns.Add("IHSR Score", typeof(int));
                    Data_Table_2G.Columns.Add("OHSR Score", typeof(int));
                    Data_Table_2G.Columns.Add("Availability Score", typeof(int));
                    Data_Table_2G.Columns.Add("Cell Score", typeof(double));
                    // Data_Table_2G.Columns.Add("Site Score", typeof(double));
                    Data_Table_2G.Columns.Add("Site", typeof(string));



                    Site_Data_Table_2G = new DataTable();
                    Site_Data_Table_2G.Columns.Add("Site", typeof(string));
                    Site_Data_Table_2G.Columns.Add("KPI Zero Status", typeof(string));
                    Site_Data_Table_2G.Columns.Add("Rejected Cell List", typeof(string));


                    // dataGridView1.ColumnCount = 11;
                    dataGridView1.Invoke(new Action(() => dataGridView1.ColumnCount = 11));
                    // dataGridView1.Rows.Clear();
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows.Clear()));
                    // dataGridView1.RowCount = Data_Table_2G.Rows.Count + 1;
                    dataGridView1.Invoke(new Action(() => dataGridView1.RowCount = Data_Table_2G.Rows.Count + 1));

                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[0].Value = "Date")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[0].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[1].Value = "BSC")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[1].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[2].Value = "Site")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[2].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[3].Value = "Cell")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[3].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[4].Value = "TCH_Traffic_BH (Erlang)")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[4].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[5].Value = "CSSR")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[5].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[6].Value = "CDR")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[6].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[7].Value = "IHSR")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[7].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[8].Value = "OHSR")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[8].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[9].Value = "Availability")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[9].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[10].Value = "Cell Status")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[10].Width = 100));


                    //dataGridView1.Rows[0].Cells[0].Value = "Date"; dataGridView1.Columns[0].Width = 100;
                    //dataGridView1.Rows[0].Cells[1].Value = "BSC"; dataGridView1.Columns[1].Width = 100;
                    //dataGridView1.Rows[0].Cells[2].Value = "Site"; dataGridView1.Columns[2].Width = 100;
                    //dataGridView1.Rows[0].Cells[3].Value = "Cell"; dataGridView1.Columns[3].Width = 100;
                    //dataGridView1.Rows[0].Cells[4].Value = "TCH_Traffic_BH (Erlang)"; dataGridView1.Columns[4].Width = 100;
                    //dataGridView1.Rows[0].Cells[5].Value = "CSSR"; dataGridView1.Columns[5].Width = 100;
                    //dataGridView1.Rows[0].Cells[6].Value = "CDR"; dataGridView1.Columns[6].Width = 100;
                    //dataGridView1.Rows[0].Cells[7].Value = "IHSR"; dataGridView1.Columns[7].Width = 100;
                    //dataGridView1.Rows[0].Cells[8].Value = "OHSR"; dataGridView1.Columns[8].Width = 100;
                    //dataGridView1.Rows[0].Cells[9].Value = "Availability"; dataGridView1.Columns[9].Width = 100;
                    //dataGridView1.Rows[0].Cells[10].Value = "Cell Status"; dataGridView1.Columns[10].Width = 100;
                    // dataGridView1.Rows[0].Cells[11].Value = "Site Status"; dataGridView1.Columns[11].Width = 100;

                    progressBar1.Minimum = 0;



                    if (Data_Table_2G.Rows.Count == 0)
                    {
                        MessageBox.Show("There is no Data in Database");
                    }

                    if (Data_Table_2G.Rows.Count != 0)
                    {
                        //progressBar1.Maximum = Data_Table_2G.Rows.Count - 1;
                        progressBar1.Invoke(new Action(() => progressBar1.Maximum = Data_Table_2G.Rows.Count - 1));
                        for (int k = 0; k < Data_Table_2G.Rows.Count; k++)
                        {

                            int result = 0;

                            // Date
                            //dataGridView1.Rows[k + 1].Cells[0].Value = Data_Table_2G.Rows[k][0];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[0].Value = Data_Table_2G.Rows[k][0]));
                            //BSC
                            //   dataGridView1.Rows[k + 1].Cells[1].Value = Data_Table_2G.Rows[k][1];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[1].Value = Data_Table_2G.Rows[k][1]));

                            // Site
                            string Cell = Data_Table_2G.Rows[k][2].ToString();
                            string Site = "";
                            if (Cell.Length == 7)
                            {
                                Site = Cell.Substring(0, 6);
                            }
                            if (Cell.Length > 7)
                            {
                                Site = Cell.Substring(0, 2) + Cell.Substring(4, 4);
                            }
                            //dataGridView1.Rows[k + 1].Cells[2].Value = Site;
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[2].Value = Site));
                            Data_Table_2G.Rows[k][16] = Site;

                            // Cell
                            // dataGridView1.Rows[k + 1].Cells[3].Value = Data_Table_2G.Rows[k][2];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[3].Value = Data_Table_2G.Rows[k][2]));

                            // Traffic

                            //  dataGridView1.Rows[k + 1].Cells[4].Value = Data_Table_2G.Rows[k][3];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[4].Value = Data_Table_2G.Rows[k][3]));
                            if (Data_Table_2G.Rows[k][3].ToString() == "")
                            {
                                Data_Table_2G.Rows[k][9] = -1;
                            }
                            else if (Convert.ToDouble(Data_Table_2G.Rows[k][3]) == Convert.ToDouble(textBox7.Text))
                            {
                                Data_Table_2G.Rows[k][9] = 1; result++; dataGridView1.Rows[k + 1].Cells[4].Style.BackColor = Color.Orange;
                            }
                            else if (Convert.ToDouble(Data_Table_2G.Rows[k][3]) > Convert.ToDouble(textBox7.Text))
                            {
                                Data_Table_2G.Rows[k][9] = 0;
                            }

                            // CSSR
                            // dataGridView1.Rows[k + 1].Cells[5].Value = Data_Table_2G.Rows[k][4];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[5].Value = Data_Table_2G.Rows[k][4]));
                            if (Data_Table_2G.Rows[k][4].ToString() == "")
                            {
                                Data_Table_2G.Rows[k][10] = -1;
                            }
                            else if (Convert.ToDouble(Data_Table_2G.Rows[k][4]) < Convert.ToDouble(textBox3.Text))
                            {
                                Data_Table_2G.Rows[k][10] = 1; result++; dataGridView1.Rows[k + 1].Cells[5].Style.BackColor = Color.Orange;
                            }
                            else if (Convert.ToDouble(Data_Table_2G.Rows[k][4]) >= Convert.ToDouble(textBox3.Text))
                            {
                                Data_Table_2G.Rows[k][10] = 0;
                            }

                            // CDR
                            // dataGridView1.Rows[k + 1].Cells[6].Value = Data_Table_2G.Rows[k][5];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[6].Value = Data_Table_2G.Rows[k][5]));
                            if (Data_Table_2G.Rows[k][5].ToString() == "")
                            {
                                Data_Table_2G.Rows[k][11] = -1;
                            }
                            else if (Convert.ToDouble(Data_Table_2G.Rows[k][5]) > Convert.ToDouble(textBox2.Text))
                            {
                                Data_Table_2G.Rows[k][11] = 1; result++; dataGridView1.Rows[k + 1].Cells[6].Style.BackColor = Color.Orange;
                            }
                            else if (Convert.ToDouble(Data_Table_2G.Rows[k][5]) <= Convert.ToDouble(textBox2.Text))
                            {
                                Data_Table_2G.Rows[k][11] = 0;
                            }

                            // IHSR
                            //  dataGridView1.Rows[k + 1].Cells[7].Value = Data_Table_2G.Rows[k][6];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[7].Value = Data_Table_2G.Rows[k][6]));
                            if (Data_Table_2G.Rows[k][6].ToString() == "")
                            {
                                Data_Table_2G.Rows[k][12] = -1;
                            }
                            else if (Convert.ToDouble(Data_Table_2G.Rows[k][6]) < Convert.ToDouble(textBox4.Text))
                            {
                                Data_Table_2G.Rows[k][12] = 1; result++; dataGridView1.Rows[k + 1].Cells[7].Style.BackColor = Color.Orange;
                            }
                            else if (Convert.ToDouble(Data_Table_2G.Rows[k][6]) >= Convert.ToDouble(textBox4.Text))
                            {
                                Data_Table_2G.Rows[k][12] = 0;
                            }

                            // OHSR
                            //  dataGridView1.Rows[k + 1].Cells[8].Value = Data_Table_2G.Rows[k][7];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[8].Value = Data_Table_2G.Rows[k][7]));
                            if (Data_Table_2G.Rows[k][7].ToString() == "")
                            {
                                Data_Table_2G.Rows[k][13] = -1;
                            }
                            else if (Convert.ToDouble(Data_Table_2G.Rows[k][7]) < Convert.ToDouble(textBox5.Text))
                            {
                                Data_Table_2G.Rows[k][13] = 1; result++; dataGridView1.Rows[k + 1].Cells[8].Style.BackColor = Color.Orange;
                            }
                            else if (Convert.ToDouble(Data_Table_2G.Rows[k][7]) >= Convert.ToDouble(textBox5.Text))
                            {
                                Data_Table_2G.Rows[k][13] = 0;
                            }

                            // Availability
                            //  dataGridView1.Rows[k + 1].Cells[9].Value = Data_Table_2G.Rows[k][8];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[9].Value = Data_Table_2G.Rows[k][8]));
                            if (Data_Table_2G.Rows[k][8].ToString() == "")
                            {
                                Data_Table_2G.Rows[k][14] = -1;
                            }
                            else if (Convert.ToDouble(Data_Table_2G.Rows[k][8]) < Convert.ToDouble(textBox6.Text) && Convert.ToDouble(Data_Table_2G.Rows[k][8]) > 0)
                            {
                                Data_Table_2G.Rows[k][14] = 1; result++; dataGridView1.Rows[k + 1].Cells[9].Style.BackColor = Color.Orange;
                            }
                            else if (Convert.ToDouble(Data_Table_2G.Rows[k][8]) >= Convert.ToDouble(textBox6.Text))
                            {
                                Data_Table_2G.Rows[k][14] = 0;
                            }


                            if (Convert.ToInt16(Data_Table_2G.Rows[k][9]) == -1)
                            {
                                Data_Table_2G.Rows[k][15] = 1; dataGridView1.Rows[k + 1].Cells[10].Value = "Not Updated";
                            }
                            else if (result == 0)
                            {
                                Data_Table_2G.Rows[k][15] = 0.1; dataGridView1.Rows[k + 1].Cells[10].Value = "Passed";
                            }
                            if (result > 0)
                            {
                                Data_Table_2G.Rows[k][15] = 0; dataGridView1.Rows[k + 1].Cells[10].Value = "Rejected";
                            }

                            //progressBar1.Value = k;
                            progressBar1.Invoke(new Action(() => progressBar1.Value = k));

                        }


                        var distinctIds = Data_Table_2G.AsEnumerable()
                           .Select(s => new
                           {
                               id = s.Field<string>("Site"),
                           })
                           .Distinct().ToList();

                        for (int j = 0; j < distinctIds.Count; j++)
                        {
                            var cell_data = (from p in Data_Table_2G.AsEnumerable()
                                             where p.Field<string>("Site") == distinctIds[j].id
                                             select p).ToList();


                            double multiplier = 1;
                            for (int h = 0; h < cell_data.Count; h++)
                            {
                                multiplier = multiplier * Convert.ToDouble(cell_data[h].ItemArray[15]);

                            }

                            if (multiplier > 0 && multiplier < 1)
                            {
                                Site_Data_Table_2G.Rows.Add(distinctIds[j].id, "Passed");
                            }
                            if (multiplier == 0)
                            {
                                Site_Data_Table_2G.Rows.Add(distinctIds[j].id, "Rejected");
                            }
                            if (multiplier == 1)
                            {
                                Site_Data_Table_2G.Rows.Add(distinctIds[j].id, "Not Updated");
                            }

                        }
                    }

                }



                if (Technology == "3G")
                {


                    string Data_Quary1 = @" select [Date], substring(convert(varchar,Date,23),1,10) as 'Day', [ElementID] as 'RNC', substring([ElementID1],1,9) as 'Sector', [ElementID1] as 'Cell', [CS_Traffic_BH] as 'CS_Traffic_BH (Erlang)', [Cs_RAB_Establish_Success_Rate] as 'CS RAB Establish', [CS_RRC_Setup_Success_Rate] as'CS RRC SR', [CS_Drop_Call_Rate] as 'Voice Drop Rate', [Cell_Availability_Rate_Exclude_Blocking(UCELL_Eric)] as 'Cell Availability' from [dbo].[CC3_Ericsson_Cell_BH] where  (" + sites_list + ") and (" + date_list + ")" +
                        @" union all select [Date], substring(convert(varchar,Date,23),1,10) as 'Day', [ElementID] as 'RNC',  substring([ElementID1],1,9) as 'Sector', [ElementID1] as 'Cell', [CS_Erlang] as 'CS_Traffic_BH (Erlang)', [CS_RAB_Setup_Success_Ratio] as 'CS RAB Establish', [CS_RRC_Connection_Establishment_SR] as'CS RRC SR',  [AMR_Call_Drop_Ratio_New(Hu_CELL)] as 'Voice Drop Rate',  [Radio_Network_Availability_Ratio(Hu_Cell)] as 'Cell Availability' from [dbo].[CC3_Huawei_Cell_BH] where (" + sites_list + ") and (" + date_list + ")" +
                        @" union all select [Date], substring(convert(varchar,Date,23),1,10) as 'Day', [ElementID] as 'RNC',  substring([ElementID1],1,9) as 'Sector', [ElementID1] as 'Cell', [CS_TrafficBH] as 'CS_Traffic_BH (Erlang)', [CS_RAB_Establish_Success_Rate] as 'CS RAB Establish', [CS_RRC_SETUP_SR_WITHOUT_REPEAT(CELL_NOKIA)] as 'CS RRC SR', [CS_Drop_Call_Rate] as 'Voice Drop Rate', [Cell_Availability_excluding_blocked_by_user_state] as 'Cell Availability' from [dbo].[CC3_Nokia_Cell_BH] where (" + sites_list + ") and (" + date_list + ")";


                    string Data_Quary2 = @"
                select T1.[Day]
      ,T1.[RNC]
      ,T1.[Sector] ,
	  
	  case when Sum_CS_RRC_SR is not null then T1.[CS RRC SR] else null end as [CS RRC SR],
	  case when Sum_CS_RAB_Establish is not null then T1.[CS RAB Establish] else null end as [CS RAB Establish],
      case when Sum_Voice_Drop_Rate is not null then T1.[Voice Drop Rate] else null end as [Voice Drop Rate],
     case when Sum_Cell_Availability is not null then T1.[Cell Availability] else null end as [Cell Availability],
      T1.[CS_Traffic_BH (Erlang)]

      from (select Day, RNC, Sector,
       sum(CS_RRC_SR_T)/case when sum([CS_RRC_SR_T1])=0 then(-1) else sum([CS_RRC_SR_T1]) end as 'CS RRC SR',

       sum(CS_RAB_Establish_T)/case when sum([CS_RAB_Establish_T1])=0 then(-1) else sum([CS_RAB_Establish_T1]) end as 'CS RAB Establish',

       sum(Voice_Drop_Rate_T)/case when sum([Voice_Drop_Rate_T1])=0 then(-1) else sum([Voice_Drop_Rate_T1]) end as 'Voice Drop Rate',

       sum(Cell_Availability_T)/case when sum([Cell_Availability_T1])=0 then(-1) else sum([Cell_Availability_T1]) end as 'Cell Availability',

       sum(cast([CS_Traffic_BH (Erlang)] as float)) as 'CS_Traffic_BH (Erlang)'


from
        (select Day, RNC, Sector,

        [CS RRC SR],
		(isnull([CS RRC SR],0)*[CS_Traffic_BH (Erlang)]) CS_RRC_SR_T,
		case when[CS RRC SR] is null then 0
		else [CS_Traffic_BH (Erlang)]
        end as CS_RRC_SR_T1,

		[CS RAB Establish],
		(isnull([CS RAB Establish],0)*[CS_Traffic_BH (Erlang)]) CS_RAB_Establish_T,
		case when[CS RAB Establish] is null then 0
		else [CS_Traffic_BH (Erlang)]
        end as CS_RAB_Establish_T1,

		[Voice Drop Rate],
		(isnull([Voice Drop Rate],0)*[CS_Traffic_BH (Erlang)]) Voice_Drop_Rate_T,
		case when[Voice Drop Rate] is null then 0
		else [CS_Traffic_BH (Erlang)]
        end as Voice_Drop_Rate_T1,

		[Cell Availability],
		(isnull([Cell Availability],0)*[CS_Traffic_BH (Erlang)]) Cell_Availability_T,
		case when[Cell Availability] is null then 0
		else [CS_Traffic_BH (Erlang)]
        end as Cell_Availability_T1,

		[CS_Traffic_BH (Erlang)]

        from(" + Data_Quary1 + @" )as tbl )tb
group by Day, RNC, Sector
		
		) T1
inner join (SELECT [Day]
                 , [RNC]
                 , [Sector]

                 , sum(cast([CS RRC SR] as float)) Sum_CS_RRC_SR
				 ,sum(cast([CS RAB Establish] as float)) Sum_CS_RAB_Establish
                 ,sum(cast([Voice Drop Rate] as float)) Sum_Voice_Drop_Rate
                 ,sum(cast([Cell Availability] as float)) Sum_Cell_Availability
            from(" + Data_Quary1 + @" )as tbl 
        group by[Day],[RNC],[Sector]	 
           )T2

  on T1.[Day]=T2.[Day] and
  T1.[RNC]=T2.[RNC] and
  T1.[Sector]=T2.[Sector]";


                    SqlCommand CS_Data_Quary = new SqlCommand(Data_Quary2, connection);
                    CS_Data_Quary.CommandTimeout = 0;
                    CS_Data_Quary.ExecuteNonQuery();
                    Data_Table_3G_CS = new DataTable();
                    SqlDataAdapter Date_Table1 = new SqlDataAdapter(CS_Data_Quary);
                    Date_Table1.Fill(Data_Table_3G_CS);




                    //string Data_Quary3 = @" select [Date], substring(convert(varchar,Date,23),1,10) as 'Day', [ElementID] as 'RNC', substring([ElementID1],1,9) as 'Sector', [ElementID1] as 'Cell', [PS_Volume(GB)(UCell_Eric)] as 'PS_Traffic_Daily (GB)', [Ps_RAB_Establish_Success_Rate] as 'PS RAB Establish', [PS_RRC_Setup_Success_Rate(UCell_Eric)] as'PS RRC SR', [HSDPA_Drop_Call_Rate(UCell_Eric)] as 'HSDPA Drop Rate', [uplink_average_RSSI_dbm_(Eric_UCELL)] as 'RSSI (dBm)' from [dbo].[RD3_Ericsson_Cell_Daily] where  (" + sites_list + ") and (" + date_list + ")" +
                    //  @" union all select [Date], substring(convert(varchar,Date,23),1,10) as 'Day', [ElementID] as 'RNC',  substring([ElementID1],1,9) as 'Sector', [ElementID1] as 'Cell', [PAYLOAD] as 'PS_Traffic_Daily (GB)', [PS_RAB_Setup_Success_Ratio] as 'PS RAB Establish', [PS_RRC_Connection_success_Rate_repeatless(Hu_Cell)] as'PS RRC SR',  [HSDPA_cdr(%)_(Hu_Cell)_new] as 'HSDPA Drop Rate',  [Mean_RTWP(Cell_Hu)] as 'RSSI (dBm)' from [dbo].[RD3_Huawei_Cell_Daily] where (" + sites_list + ") and (" + date_list + ")" +
                    //  @" union all select [Date], substring(convert(varchar,Date,23),1,10) as 'Day', [ElementID] as 'RNC',  substring([ElementID1],1,9) as 'Sector', [ElementID1] as 'Cell', [PS_Payload_Total(HS+R99)(Nokia_CELL)_GB] as 'PS_Traffic_Daily (GB)', [RAB_Setup_and_Access_Complete_Ratio_for_NRT_Service_from_User_pe] as 'PS RAB Establish', [PS_RRCSETUP_SR] as 'PS RRC SR', [HSDPA_Call_Drop_Rate(Nokia_Cell)] as 'HSDPA Drop Rate', [average_RTWP_dbm(Nokia_Cell)] as 'RSSI (dBm)' from [dbo].[RD3_Nokia_Cell_Daily] where (" + sites_list + ") and (" + date_list + ")";


                    string Data_Quary3 = @" select [Date], substring(convert(varchar,Date,23),1,10) as 'Day', [ElementID] as 'RNC', substring([ElementID1],1,9) as 'Sector', [ElementID1] as 'Cell', [PS_Volume(GB)(UCell_Eric)] as 'PS_Traffic_Daily (GB)', [Ps_RAB_Establish_Success_Rate] as 'PS RAB Establish', [PS_RRC_Setup_Success_Rate(UCell_Eric)] as'PS RRC SR', [HSDPA_Drop_Call_Rate(UCell_Eric)] as 'HSDPA Drop Rate' , [HS_USER_Throughput_NET_PQ(Mbps)(UCell_Eric)] as 'HS User THR' from [dbo].[RD3_Ericsson_Cell_Daily] where  (" + sites_list + ") and (" + date_list + ")" +
                      @" union all select [Date], substring(convert(varchar,Date,23),1,10) as 'Day', [ElementID] as 'RNC',  substring([ElementID1],1,9) as 'Sector', [ElementID1] as 'Cell', [PAYLOAD] as 'PS_Traffic_Daily (GB)', [PS_RAB_Setup_Success_Ratio] as 'PS RAB Establish', [PS_RRC_Connection_success_Rate_repeatless(Hu_Cell)] as'PS RRC SR',  [HSDPA_cdr(%)_(Hu_Cell)_new] as 'HSDPA Drop Rate',  [AVERAGE_HSDPA_USER_THROUGHPUT_DC+SC(Mbit/s)(CELL_HUAWEI)] as 'HS User THR' from [dbo].[RD3_Huawei_Cell_Daily] where (" + sites_list + ") and (" + date_list + ")" +
                      @" union all select [Date], substring(convert(varchar,Date,23),1,10) as 'Day', [ElementID] as 'RNC',  substring([ElementID1],1,9) as 'Sector', [ElementID1] as 'Cell', [PS_Payload_Total(HS+R99)(Nokia_CELL)_GB] as 'PS_Traffic_Daily (GB)', [RAB_Setup_and_Access_Complete_Ratio_for_NRT_Service_from_User_pe] as 'PS RAB Establish', [PS_RRCSETUP_SR] as 'PS RRC SR', [HSDPA_Call_Drop_Rate(Nokia_Cell)] as 'HSDPA Drop Rate', [AVERAGE_HSDPA_USER_THROUGHPUT_DC+SC(Mbit/s)(Nokia_CELL)] as  'HS User THR' from [dbo].[RD3_Nokia_Cell_Daily] where (" + sites_list + ") and (" + date_list + ")";



                    string Data_Quary4 = @"
                select T1.[Day]
      ,T1.[RNC]
      ,T1.[Sector] ,
	  
	  case when Sum_PS_RRC_SR is not null then T1.[PS RRC SR] else null end as [PS RRC SR],
	  case when Sum_PS_RAB_Establish is not null then T1.[PS RAB Establish] else null end as [PS RAB Establish],
      case when Sum_HSDPA_Drop_Rate is not null then T1.[HSDPA Drop Rate] else null end as [HSDPA Drop Rate],
      case when Sum_HS_User_THR is not null then T1.[HS User THR] else null end as [HS User THR],
      T1.[PS_Traffic_Daily (GB)]

      from (select Day, RNC, Sector,
       sum(PS_RRC_SR_T)/case when sum([PS_RRC_SR_T1])=0 then(-1) else sum([PS_RRC_SR_T1]) end as 'PS RRC SR',

       sum(PS_RAB_Establish_T)/case when sum([PS_RAB_Establish_T1])=0 then(-1) else sum([PS_RAB_Establish_T1]) end as 'PS RAB Establish',

       sum(HSDPA_Drop_Rate_T)/case when sum([HSDPA_Drop_Rate_T1])=0 then(-1) else sum([HSDPA_Drop_Rate_T1]) end as 'HSDPA Drop Rate',

        sum(HS_User_THR_T)/case when sum([HS_User_THR_T1])=0 then(-1) else sum([HS_User_THR_T1]) end as 'HS User THR',

       sum(cast([PS_Traffic_Daily (GB)] as float)) as 'PS_Traffic_Daily (GB)'


from
        (select Day, RNC, Sector,

        [PS RRC SR],
		(isnull([PS RRC SR],0)*[PS_Traffic_Daily (GB)]) PS_RRC_SR_T,
		case when[PS RRC SR] is null then 0
		else [PS_Traffic_Daily (GB)]
        end as PS_RRC_SR_T1,

		[PS RAB Establish],
		(isnull([PS RAB Establish],0)*[PS_Traffic_Daily (GB)]) PS_RAB_Establish_T,
		case when[PS RAB Establish] is null then 0
		else [PS_Traffic_Daily (GB)]
        end as PS_RAB_Establish_T1,

		[HSDPA Drop Rate],
		(isnull([HSDPA Drop Rate],0)*[PS_Traffic_Daily (GB)]) HSDPA_Drop_Rate_T,
		case when[HSDPA Drop Rate] is null then 0
		else [PS_Traffic_Daily (GB)]
        end as HSDPA_Drop_Rate_T1,


		[HS User THR],
		(isnull([HS_User_THR],0)*[PS_Traffic_Daily (GB)]) HS_User_THR_T,
		case when[HS User THR] is null then 0
		else [PS_Traffic_Daily (GB)]
        end as HS_User_THR_T1,

		[PS_Traffic_Daily (GB)]

        from(" + Data_Quary3 + @" )as tbl )tb
group by Day, RNC, Sector
		
		) T1
inner join (SELECT [Day]
                 , [RNC]
                 , [Sector]

                 , sum(cast([PS RRC SR] as float)) Sum_PS_RRC_SR
				 ,sum(cast([PS RAB Establish] as float)) Sum_PS_RAB_Establish
                 ,sum(cast([HSDPA Drop Rate] as float)) Sum_HSDPA_Drop_Rate
                ,sum(cast([HS User THR] as float)) Sum_HS_User_THR
            from(" + Data_Quary3 + @" )as tbl 
        group by[Day],[RNC],[Sector]	 
           )T2

  on T1.[Day]=T2.[Day] and
  T1.[RNC]=T2.[RNC] and
  T1.[Sector]=T2.[Sector]";


                    SqlCommand PS_Data_Quary = new SqlCommand(Data_Quary4, connection);
                    PS_Data_Quary.CommandTimeout = 0;
                    PS_Data_Quary.ExecuteNonQuery();
                    Data_Table_3G_PS = new DataTable();
                    SqlDataAdapter Date_Table2 = new SqlDataAdapter(PS_Data_Quary);
                    Date_Table2.Fill(Data_Table_3G_PS);

                    Data_Table_3G = new DataTable();

                    Data_Table_3G.Columns.Add("Day", typeof(string));
                    Data_Table_3G.Columns.Add("RNC", typeof(string));
                    Data_Table_3G.Columns.Add("Sector", typeof(string));
                    Data_Table_3G.Columns.Add("CS RRC SR", typeof(string));
                    Data_Table_3G.Columns.Add("CS RAB Establish", typeof(string));
                    Data_Table_3G.Columns.Add("Voice Drop Rate", typeof(string));
                    Data_Table_3G.Columns.Add("Cell Availability", typeof(string));
                    Data_Table_3G.Columns.Add("BH CS Traffic (Erlang)", typeof(string));
                    Data_Table_3G.Columns.Add("PS RRC SR", typeof(string));
                    Data_Table_3G.Columns.Add("PS RAB Establish", typeof(string));
                    Data_Table_3G.Columns.Add("HSDPA Drop Rate", typeof(string));
                    Data_Table_3G.Columns.Add("HS User THR (Mbps)", typeof(string));
                    Data_Table_3G.Columns.Add("Daily PS Traffic (GB)", typeof(string));

                    Data_Table_3G.Columns.Add("CS RRC SR Score", typeof(int));
                    Data_Table_3G.Columns.Add("CS RAB Establish Score", typeof(int));
                    Data_Table_3G.Columns.Add("Voice Drop Rate Score", typeof(int));
                    Data_Table_3G.Columns.Add("Cell Availability Score", typeof(int));
                    Data_Table_3G.Columns.Add("CS Traffic Score", typeof(int));

                    Data_Table_3G.Columns.Add("PS RRC SR Score", typeof(int));
                    Data_Table_3G.Columns.Add("PS RAB Establish Score", typeof(int));
                    Data_Table_3G.Columns.Add("HSDPA Drop Rate Score", typeof(int));
                    Data_Table_3G.Columns.Add("THR Score", typeof(int));
                    Data_Table_3G.Columns.Add("PS Traffic Score", typeof(int));

                    Data_Table_3G.Columns.Add("Cell Score", typeof(double));
                    Data_Table_3G.Columns.Add("Site", typeof(string));




                    Site_Data_Table_3G = new DataTable();
                    Site_Data_Table_3G.Columns.Add("Site", typeof(string));
                    Site_Data_Table_3G.Columns.Add("KPI Zero Status", typeof(string));
                    Site_Data_Table_3G.Columns.Add("Rejected Cell List", typeof(string));



                    //dataGridView1.ColumnCount = 14;
                    dataGridView1.Invoke(new Action(() => dataGridView1.ColumnCount = 15));

                    // dataGridView1.Rows.Clear();
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows.Clear()));
                    //dataGridView1.RowCount = Data_Table_3G_CS.Rows.Count + 1;
                    dataGridView1.Invoke(new Action(() => dataGridView1.RowCount = Data_Table_3G_CS.Rows.Count + 1));

                    //dataGridView1.Rows[0].Cells[0].Value = "Date"; dataGridView1.Columns[0].Width = 100;
                    //dataGridView1.Rows[0].Cells[1].Value = "RNC"; dataGridView1.Columns[1].Width = 100;
                    //dataGridView1.Rows[0].Cells[2].Value = "Site"; dataGridView1.Columns[2].Width = 100;
                    //dataGridView1.Rows[0].Cells[3].Value = "Sector"; dataGridView1.Columns[3].Width = 100;
                    //dataGridView1.Rows[0].Cells[4].Value = "CS RRC SR"; dataGridView1.Columns[4].Width = 100;
                    //dataGridView1.Rows[0].Cells[5].Value = "CS RAB Establish"; dataGridView1.Columns[5].Width = 100;
                    //dataGridView1.Rows[0].Cells[6].Value = "Voice Drop Rate"; dataGridView1.Columns[6].Width = 100;
                    //dataGridView1.Rows[0].Cells[7].Value = "Cell Availability"; dataGridView1.Columns[7].Width = 100;
                    //dataGridView1.Rows[0].Cells[8].Value = "BH CS Traffic (Erlang)"; dataGridView1.Columns[8].Width = 100;
                    //dataGridView1.Rows[0].Cells[9].Value = "PS RRC SR"; dataGridView1.Columns[9].Width = 100;
                    //dataGridView1.Rows[0].Cells[10].Value = "PS RAB Establish"; dataGridView1.Columns[10].Width = 100;
                    //dataGridView1.Rows[0].Cells[11].Value = "HSDPA Drop Rate"; dataGridView1.Columns[11].Width = 100;
                    ////dataGridView1.Rows[0].Cells[12].Value = "RSSI (dBm)"; dataGridView1.Columns[12].Width = 100;
                    //dataGridView1.Rows[0].Cells[12].Value = "Daily PS Traffic (GB)"; dataGridView1.Columns[12].Width = 100;
                    //dataGridView1.Rows[0].Cells[13].Value = "Cell Status"; dataGridView1.Columns[13].Width = 100;


                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[0].Value = "Date")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[0].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[1].Value = "RNC")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[1].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[2].Value = "Site")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[2].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[3].Value = "Sector")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[3].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[4].Value = "CS RRC SR")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[4].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[5].Value = "CS RAB Establish")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[5].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[6].Value = "Voice Drop Rate")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[6].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[7].Value = "Cell Availability")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[7].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[8].Value = "BH CS Traffic (Erlang)")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[8].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[9].Value = "PS RRC SR")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[9].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[10].Value = "PS RAB Establish")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[10].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[11].Value = "HSDPA Drop Rate")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[11].Width = 100));

                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[12].Value = "HS User THR")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[12].Width = 100));


                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[13].Value = "Daily PS Traffic (GB)")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[13].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[14].Value = "Cell Status")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[14].Width = 100));






                    progressBar1.Minimum = 0;
                    //progressBar1.Maximum = Data_Table_3G_CS.Rows.Count - 1;


                    if (Data_Table_3G_CS.Rows.Count == 0)
                    {
                        MessageBox.Show("There is no Data in Database");
                    }

                    if (Data_Table_3G_CS.Rows.Count != 0)
                    {
                        // progressBar1.Maximum = Data_Table_3G_CS.Rows.Count - 1;
                        progressBar1.Invoke(new Action(() => progressBar1.Maximum = Data_Table_3G_CS.Rows.Count - 1));
                        for (int k = 0; k < Data_Table_3G_CS.Rows.Count; k++)
                        {
                            int result = 0;

                            Data_Table_3G.Rows.Add(Data_Table_3G_CS.Rows[k][0], Data_Table_3G_CS.Rows[k][1], Data_Table_3G_CS.Rows[k][2], Data_Table_3G_CS.Rows[k][3], Data_Table_3G_CS.Rows[k][4], Data_Table_3G_CS.Rows[k][5], Data_Table_3G_CS.Rows[k][6], Data_Table_3G_CS.Rows[k][7], "", "", "", "", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "");

                            for (int j = 0; j < Data_Table_3G_PS.Rows.Count; j++)
                            {
                                if ((Data_Table_3G_CS.Rows[k][0].ToString() == Data_Table_3G_PS.Rows[j][0].ToString()) && (Data_Table_3G_CS.Rows[k][2].ToString() == Data_Table_3G_PS.Rows[j][2].ToString()))
                                {
                                    Data_Table_3G.Rows[k][8] = Data_Table_3G_PS.Rows[j][3];
                                    Data_Table_3G.Rows[k][9] = Data_Table_3G_PS.Rows[j][4];
                                    Data_Table_3G.Rows[k][10] = Data_Table_3G_PS.Rows[j][5];
                                    Data_Table_3G.Rows[k][11] = Data_Table_3G_PS.Rows[j][6];
                                    Data_Table_3G.Rows[k][12] = Data_Table_3G_PS.Rows[j][7];
                                }
                            }


                            // Date
                            //  dataGridView1.Rows[k + 1].Cells[0].Value = Data_Table_3G.Rows[k][0];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[0].Value = Data_Table_3G.Rows[k][0]));
                            //RNC
                            //  dataGridView1.Rows[k + 1].Cells[1].Value = Data_Table_3G.Rows[k][1];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[1].Value = Data_Table_3G.Rows[k][1]));
                            // Site
                            string Cell = Data_Table_3G.Rows[k][2].ToString();
                            //dataGridView1.Rows[k + 1].Cells[2].Value = Cell.Substring(0, 8);
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[2].Value = Cell.Substring(0, 8)));
                            Data_Table_3G.Rows[k][22] = Cell.Substring(0, 8);
                            // Sector
                            // dataGridView1.Rows[k + 1].Cells[3].Value = Data_Table_3G.Rows[k][2];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[3].Value = Data_Table_3G.Rows[k][2]));
                            // CS RRC SR
                            // dataGridView1.Rows[k + 1].Cells[4].Value = Data_Table_3G.Rows[k][3];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[4].Value = Data_Table_3G.Rows[k][3]));
                            if (Data_Table_3G.Rows[k][3].ToString() == "")
                            {
                                Data_Table_3G.Rows[k][12] = -1;
                            }
                            else if (Convert.ToDouble(Data_Table_3G.Rows[k][3]) < Convert.ToDouble(textBox10.Text))
                            {
                                Data_Table_3G.Rows[k][12] = 1; result++; dataGridView1.Rows[k + 1].Cells[4].Style.BackColor = Color.Orange;
                            }
                            else if (Convert.ToDouble(Data_Table_3G.Rows[k][3]) >= Convert.ToDouble(textBox10.Text))
                            {
                                Data_Table_3G.Rows[k][12] = 0;
                            }


                            // CS RAB Setablish
                            //  dataGridView1.Rows[k + 1].Cells[5].Value = Data_Table_3G.Rows[k][4];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[5].Value = Data_Table_3G.Rows[k][4]));
                            if (Data_Table_3G.Rows[k][4].ToString() == "")
                            {
                                Data_Table_3G.Rows[k][13] = -1;
                            }
                            else if (Convert.ToDouble(Data_Table_3G.Rows[k][4]) < Convert.ToDouble(textBox9.Text))
                            {
                                Data_Table_3G.Rows[k][13] = 1; result++; dataGridView1.Rows[k + 1].Cells[5].Style.BackColor = Color.Orange;
                            }
                            else if (Convert.ToDouble(Data_Table_3G.Rows[k][4]) >= Convert.ToDouble(textBox9.Text))
                            {
                                Data_Table_3G.Rows[k][13] = 0;
                            }


                            // Voice Drop Rate
                            //dataGridView1.Rows[k + 1].Cells[6].Value = Data_Table_3G.Rows[k][5];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[6].Value = Data_Table_3G.Rows[k][5]));
                            if (Data_Table_3G.Rows[k][5].ToString() == "")
                            {
                                Data_Table_3G.Rows[k][14] = -1;
                            }
                            else if (Convert.ToDouble(Data_Table_3G.Rows[k][5]) > Convert.ToDouble(textBox8.Text))
                            {
                                Data_Table_3G.Rows[k][14] = 1; result++; dataGridView1.Rows[k + 1].Cells[6].Style.BackColor = Color.Orange;
                            }
                            else if (Convert.ToDouble(Data_Table_3G.Rows[k][5]) <= Convert.ToDouble(textBox8.Text))
                            {
                                Data_Table_3G.Rows[k][14] = 0;
                            }


                            // Availability
                            //dataGridView1.Rows[k + 1].Cells[7].Value = Data_Table_3G.Rows[k][6];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[7].Value = Data_Table_3G.Rows[k][6]));
                            if (Data_Table_3G.Rows[k][6].ToString() == "")
                            {
                                Data_Table_3G.Rows[k][15] = -1;
                            }
                            else if (Convert.ToDouble(Data_Table_3G.Rows[k][6]) < Convert.ToDouble(textBox11.Text) && Convert.ToDouble(Data_Table_3G.Rows[k][6]) > 0)
                            {
                                Data_Table_3G.Rows[k][15] = 1; result++; dataGridView1.Rows[k + 1].Cells[7].Style.BackColor = Color.Orange;
                            }
                            else if (Convert.ToDouble(Data_Table_3G.Rows[k][6]) >= Convert.ToDouble(textBox11.Text))
                            {
                                Data_Table_3G.Rows[k][15] = 0;
                            }


                            // CS Traffic
                            // dataGridView1.Rows[k + 1].Cells[8].Value = Data_Table_3G.Rows[k][7];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[8].Value = Data_Table_3G.Rows[k][7]));
                            if (Data_Table_3G.Rows[k][7].ToString() == "")
                            {
                                Data_Table_3G.Rows[k][16] = -1;
                            }
                            else if (Convert.ToDouble(Data_Table_3G.Rows[k][7]) == Convert.ToDouble(textBox12.Text))
                            {
                                Data_Table_3G.Rows[k][16] = 1; result++; dataGridView1.Rows[k + 1].Cells[8].Style.BackColor = Color.Orange;
                            }
                            else if (Convert.ToDouble(Data_Table_3G.Rows[k][7]) > Convert.ToDouble(textBox12.Text))
                            {
                                Data_Table_3G.Rows[k][16] = 0;
                            }


                            // PS RRC SR
                            //dataGridView1.Rows[k + 1].Cells[9].Value = Data_Table_3G.Rows[k][8];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[9].Value = Data_Table_3G.Rows[k][8]));
                            if (Data_Table_3G.Rows[k][8].ToString() == "")
                            {
                                Data_Table_3G.Rows[k][18] = -1;
                            }
                            else if (Convert.ToDouble(Data_Table_3G.Rows[k][8]) < Convert.ToDouble(textBox15.Text))
                            {
                                Data_Table_3G.Rows[k][18] = 1; result++; dataGridView1.Rows[k + 1].Cells[9].Style.BackColor = Color.Orange;
                            }
                            else if (Convert.ToDouble(Data_Table_3G.Rows[k][8]) >= Convert.ToDouble(textBox15.Text))
                            {
                                Data_Table_3G.Rows[k][18] = 0;
                            }

                            // PS RAB Establish
                            // dataGridView1.Rows[k + 1].Cells[10].Value = Data_Table_3G.Rows[k][9];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[10].Value = Data_Table_3G.Rows[k][9]));
                            if (Data_Table_3G.Rows[k][9].ToString() == "")
                            {
                                Data_Table_3G.Rows[k][19] = -1;
                            }
                            else if (Convert.ToDouble(Data_Table_3G.Rows[k][9]) < Convert.ToDouble(textBox14.Text))
                            {
                                Data_Table_3G.Rows[k][19] = 1; result++; dataGridView1.Rows[k + 1].Cells[10].Style.BackColor = Color.Orange;
                            }
                            else if (Convert.ToDouble(Data_Table_3G.Rows[k][9]) >= Convert.ToDouble(textBox14.Text))
                            {
                                Data_Table_3G.Rows[k][19] = 0;
                            }


                            // HSDPA Drop Rate
                            // dataGridView1.Rows[k + 1].Cells[11].Value = Data_Table_3G.Rows[k][10];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[11].Value = Data_Table_3G.Rows[k][10]));
                            if (Data_Table_3G.Rows[k][10].ToString() == "")
                            {
                                Data_Table_3G.Rows[k][20] = -1;
                            }
                            else if (Convert.ToDouble(Data_Table_3G.Rows[k][10]) > Convert.ToDouble(textBox13.Text))
                            {
                                Data_Table_3G.Rows[k][20] = 1; result++; dataGridView1.Rows[k + 1].Cells[11].Style.BackColor = Color.Orange;
                            }
                            else if (Convert.ToDouble(Data_Table_3G.Rows[k][10]) <= Convert.ToDouble(textBox13.Text))
                            {
                                Data_Table_3G.Rows[k][20] = 0;
                            }



                            // HS User THR
                            // dataGridView1.Rows[k + 1].Cells[11].Value = Data_Table_3G.Rows[k][10];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[12].Value = Data_Table_3G.Rows[k][11]));
                            if (Data_Table_3G.Rows[k][11].ToString() == "")
                            {
                                Data_Table_3G.Rows[k][21] = -1;
                            }
                            else if (Convert.ToDouble(Data_Table_3G.Rows[k][11]) > 1)
                            {
                                Data_Table_3G.Rows[k][21] = 1; result++; dataGridView1.Rows[k + 1].Cells[12].Style.BackColor = Color.Orange;
                            }
                            else if (Convert.ToDouble(Data_Table_3G.Rows[k][11]) <= 1)
                            {
                                Data_Table_3G.Rows[k][21] = 0;
                            }

                            // RSSI
                            //dataGridView1.Rows[k + 1].Cells[12].Value = Data_Table_3G.Rows[k][11];
                            //string Band = Cell.Substring(2, 2);
                            //if (Data_Table_3G.Rows[k][11].ToString() == "")
                            //{
                            //    Data_Table_3G.Rows[k][21] = -1;
                            //}
                            //else if ((Convert.ToDouble(Data_Table_3G.Rows[k][11]) > Convert.ToDouble(textBox17.Text)) && (Band == "2U" || Band == "5U" || Band == "8U"))
                            //{
                            //    Data_Table_3G.Rows[k][21] = 1; result++; dataGridView1.Rows[k + 1].Cells[12].Style.BackColor = Color.Orange;
                            //}
                            //else if ((Convert.ToDouble(Data_Table_3G.Rows[k][11]) > Convert.ToDouble(textBox18.Text)) && (Band == "1U" || Band == "4U" || Band == "7U"))
                            //{
                            //    Data_Table_3G.Rows[k][21] = 1; result++; dataGridView1.Rows[k + 1].Cells[12].Style.BackColor = Color.Orange;
                            //}
                            //else if ((Convert.ToDouble(Data_Table_3G.Rows[k][11]) <= Convert.ToDouble(textBox17.Text)) && (Band == "2U" || Band == "5U" || Band == "8U"))
                            //{
                            //    Data_Table_3G.Rows[k][21] = 0;
                            //}
                            //else if ((Convert.ToDouble(Data_Table_3G.Rows[k][11]) <= Convert.ToDouble(textBox18.Text)) && (Band == "1U" || Band == "4U" || Band == "7U"))
                            //{
                            //    Data_Table_3G.Rows[k][21] = 0;
                            //}



                            // PS Traffic
                            //dataGridView1.Rows[k + 1].Cells[12].Value = Data_Table_3G.Rows[k][11];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[13].Value = Data_Table_3G.Rows[k][12]));
                            if (Data_Table_3G.Rows[k][12].ToString() == "")
                            {
                                Data_Table_3G.Rows[k][22] = -1;
                            }
                            else if (Convert.ToDouble(Data_Table_3G.Rows[k][12]) == Convert.ToDouble(textBox16.Text))
                            {
                                Data_Table_3G.Rows[k][22] = 1; result++; dataGridView1.Rows[k + 1].Cells[13].Style.BackColor = Color.Orange;
                            }
                            else if (Convert.ToDouble(Data_Table_3G.Rows[k][12]) > Convert.ToDouble(textBox16.Text))
                            {
                                Data_Table_3G.Rows[k][22] = 0;
                            }



                            if ((Convert.ToInt16(Data_Table_3G.Rows[k][17]) == -1) || (Convert.ToInt16(Data_Table_3G.Rows[k][22]) == -1))
                            {
                                //         if (Convert.ToInt16(Data_Table_3G.Rows[k][21])!= 0)
                                //  {
                                Data_Table_3G.Rows[k][23] = 1; dataGridView1.Rows[k + 1].Cells[13].Value = "Not Updated";
                                //                             }
                            }
                            else if (result == 0)
                            {
                                Data_Table_3G.Rows[k][23] = 0.1; dataGridView1.Rows[k + 1].Cells[13].Value = "Passed";
                            }
                            if (result > 0)
                            {
                                Data_Table_3G.Rows[k][23] = 0; dataGridView1.Rows[k + 1].Cells[13].Value = "Rejected";
                            }
                            progressBar1.Invoke(new Action(() => progressBar1.Value = k));
                            //   progressBar1.Value = k;
                        }



                        var distinctIds = Data_Table_3G.AsEnumerable()
           .Select(s => new
           {
               id = s.Field<string>("Site"),
           })
           .Distinct().ToList();

                        for (int j = 0; j < distinctIds.Count; j++)
                        {
                            var cell_data = (from p in Data_Table_3G.AsEnumerable()
                                             where p.Field<string>("Site") == distinctIds[j].id
                                             select p).ToList();


                            double multiplier = 1;
                            for (int h = 0; h < cell_data.Count; h++)
                            {
                                multiplier = multiplier * Convert.ToDouble(cell_data[h].ItemArray[21]);

                            }

                            if (multiplier > 0 && multiplier < 1)
                            {
                                Site_Data_Table_3G.Rows.Add(distinctIds[j].id, "Passed");
                            }
                            if (multiplier == 0)
                            {
                                Site_Data_Table_3G.Rows.Add(distinctIds[j].id, "Rejected");
                            }
                            if (multiplier == 1)
                            {
                                Site_Data_Table_3G.Rows.Add(distinctIds[j].id, "Not Updated");
                            }

                        }

                    }


                }






                if (Technology == "3G-MCI")
                {

                    string Data_Quary1 = @" select [Date], substring(convert(varchar,Date,23),1,10) as 'Day', [ElementID] as 'RNC', substring([ElementID1],1,9) as 'Sector', [ElementID1] as 'Cell', [CS_Traffic_BH] as 'CS_Traffic_BH (Erlang)', [Cs_RAB_Establish_Success_Rate] as 'CS RAB Establish', [CS_RRC_Setup_Success_Rate] as'CS RRC SR', [CS_Drop_Call_Rate] as 'Voice Drop Rate', [Cell_Availability_Rate_Exclude_Blocking(UCELL_Eric)] as 'Cell Availability' from [dbo].[CC3_Ericsson_Cell_BH] where  (" + sites_list + ") and (" + date_list + ")" +
                                @" union all select [Date], substring(convert(varchar,Date,23),1,10) as 'Day', [ElementID] as 'RNC',  substring([ElementID1],1,9) as 'Sector', [ElementID1] as 'Cell', [CS_Erlang] as 'CS_Traffic_BH (Erlang)', [CS_RAB_Setup_Success_Ratio] as 'CS RAB Establish', [CS_RRC_Connection_Establishment_SR] as'CS RRC SR',  [AMR_Call_Drop_Ratio_New(Hu_CELL)] as 'Voice Drop Rate',  [Radio_Network_Availability_Ratio(Hu_Cell)] as 'Cell Availability' from [dbo].[CC3_Huawei_Cell_BH] where (" + sites_list + ") and (" + date_list + ")" +
                                @" union all select [Date], substring(convert(varchar,Date,23),1,10) as 'Day', [ElementID] as 'RNC',  substring([ElementID1],1,9) as 'Sector', [ElementID1] as 'Cell', [CS_TrafficBH] as 'CS_Traffic_BH (Erlang)', [CS_RAB_Establish_Success_Rate] as 'CS RAB Establish', [CS_RRC_SETUP_SR_WITHOUT_REPEAT(CELL_NOKIA)] as 'CS RRC SR', [CS_Drop_Call_Rate] as 'Voice Drop Rate', [Cell_Availability_excluding_blocked_by_user_state] as 'Cell Availability' from [dbo].[CC3_Nokia_Cell_BH] where (" + sites_list + ") and (" + date_list + ")";


                    string Data_Quary2 = @"
                select T1.[Day]
      ,T1.[RNC]
      ,T1.[Sector] ,
	  
	  case when Sum_CS_RRC_SR is not null then T1.[CS RRC SR] else null end as [CS RRC SR],
	  case when Sum_CS_RAB_Establish is not null then T1.[CS RAB Establish] else null end as [CS RAB Establish],
      case when Sum_Voice_Drop_Rate is not null then T1.[Voice Drop Rate] else null end as [Voice Drop Rate],
     case when Sum_Cell_Availability is not null then T1.[Cell Availability] else null end as [Cell Availability],
      T1.[CS_Traffic_BH (Erlang)]

      from (select Day, RNC, Sector,
       sum(CS_RRC_SR_T)/case when sum([CS_RRC_SR_T1])=0 then(-1) else sum([CS_RRC_SR_T1]) end as 'CS RRC SR',

       sum(CS_RAB_Establish_T)/case when sum([CS_RAB_Establish_T1])=0 then(-1) else sum([CS_RAB_Establish_T1]) end as 'CS RAB Establish',

       sum(Voice_Drop_Rate_T)/case when sum([Voice_Drop_Rate_T1])=0 then(-1) else sum([Voice_Drop_Rate_T1]) end as 'Voice Drop Rate',

       sum(Cell_Availability_T)/case when sum([Cell_Availability_T1])=0 then(-1) else sum([Cell_Availability_T1]) end as 'Cell Availability',

       sum(cast([CS_Traffic_BH (Erlang)] as float)) as 'CS_Traffic_BH (Erlang)'


from
        (select Day, RNC, Sector,

        [CS RRC SR],
		(isnull([CS RRC SR],0)*[CS_Traffic_BH (Erlang)]) CS_RRC_SR_T,
		case when[CS RRC SR] is null then 0
		else [CS_Traffic_BH (Erlang)]
        end as CS_RRC_SR_T1,

		[CS RAB Establish],
		(isnull([CS RAB Establish],0)*[CS_Traffic_BH (Erlang)]) CS_RAB_Establish_T,
		case when[CS RAB Establish] is null then 0
		else [CS_Traffic_BH (Erlang)]
        end as CS_RAB_Establish_T1,

		[Voice Drop Rate],
		(isnull([Voice Drop Rate],0)*[CS_Traffic_BH (Erlang)]) Voice_Drop_Rate_T,
		case when[Voice Drop Rate] is null then 0
		else [CS_Traffic_BH (Erlang)]
        end as Voice_Drop_Rate_T1,

		[Cell Availability],
		(isnull([Cell Availability],0)*[CS_Traffic_BH (Erlang)]) Cell_Availability_T,
		case when[Cell Availability] is null then 0
		else [CS_Traffic_BH (Erlang)]
        end as Cell_Availability_T1,

		[CS_Traffic_BH (Erlang)]

        from(" + Data_Quary1 + @" )as tbl )tb
group by Day, RNC, Sector
		
		) T1
inner join (SELECT [Day]
                 , [RNC]
                 , [Sector]

                 , sum(cast([CS RRC SR] as float)) Sum_CS_RRC_SR
				 ,sum(cast([CS RAB Establish] as float)) Sum_CS_RAB_Establish
                 ,sum(cast([Voice Drop Rate] as float)) Sum_Voice_Drop_Rate
                 ,sum(cast([Cell Availability] as float)) Sum_Cell_Availability
            from(" + Data_Quary1 + @" )as tbl 
        group by[Day],[RNC],[Sector]	 
           )T2

  on T1.[Day]=T2.[Day] and
  T1.[RNC]=T2.[RNC] and
  T1.[Sector]=T2.[Sector]";


                    SqlCommand CS_Data_Quary = new SqlCommand(Data_Quary2, connection);
                    CS_Data_Quary.CommandTimeout = 0;
                    CS_Data_Quary.ExecuteNonQuery();
                    Data_Table_3G_CS = new DataTable();
                    SqlDataAdapter Date_Table1 = new SqlDataAdapter(CS_Data_Quary);
                    Date_Table1.Fill(Data_Table_3G_CS);




                    string Data_Quary3 = @" select [Date], substring(convert(varchar,Date,23),1,10) as 'Day', [ElementID] as 'RNC', substring([ElementID1],1,9) as 'Sector', [ElementID1] as 'Cell', [PS_Volume(GB)(UCell_Eric)] as 'PS_Traffic_Daily (GB)',  [PS_RRC_Setup_Success_Rate(UCell_Eric)] as'PS RRC SR', [PS_Drop_Call_Rate(UCell_Eric)] as 'PS Drop Rate', [uplink_average_RSSI_dbm_(Eric_UCELL)] as 'RSSI' , [HS_USER_Throughput_NET_PQ(Mbps)(UCell_Eric)] as 'HS User THR' from [dbo].[RD3_Ericsson_Cell_Daily] where  (" + sites_list + ") and (" + date_list + ")" +
                      @" union all select [Date], substring(convert(varchar,Date,23),1,10) as 'Day', [ElementID] as 'RNC',  substring([ElementID1],1,9) as 'Sector', [ElementID1] as 'Cell', [PAYLOAD] as 'PS_Traffic_Daily (GB)', [PS_RRC_Connection_success_Rate_repeatless(Hu_Cell)] as'PS RRC SR',  [PS_Call_Drop_Ratio] as 'PS Drop Rate',  [Mean_RTWP(Cell_Hu)] as 'RSSI', [AVERAGE_HSDPA_USER_THROUGHPUT_DC+SC(Mbit/s)(CELL_HUAWEI)] as 'HS User THR' from [dbo].[RD3_Huawei_Cell_Daily] where (" + sites_list + ") and (" + date_list + ")" +
                      @" union all select [Date], substring(convert(varchar,Date,23),1,10) as 'Day', [ElementID] as 'RNC',  substring([ElementID1],1,9) as 'Sector', [ElementID1] as 'Cell', [PS_Payload_Total(HS+R99)(Nokia_CELL)_GB] as 'PS_Traffic_Daily (GB)',  [PS_RRCSETUP_SR] as 'PS RRC SR', [Packet_Session_Drop_Ratio_NOKIA(CELL_NOKIA)] as 'PS Drop Rate', [average_RTWP_dbm(Nokia_Cell)] as 'RSSI', [AVERAGE_HSDPA_USER_THROUGHPUT_DC+SC(Mbit/s)(Nokia_CELL)] as 'HS User THR' from [dbo].[RD3_Nokia_Cell_Daily] where (" + sites_list + ") and (" + date_list + ")";


                    //string Data_Quary3 = @" select [Date], substring(convert(varchar,Date,23),1,10) as 'Day', [ElementID] as 'RNC', substring([ElementID1],1,9) as 'Sector', [ElementID1] as 'Cell', [PS_Volume(GB)(UCell_Eric)] as 'PS_Traffic_Daily (GB)', [Ps_RAB_Establish_Success_Rate] as 'PS RAB Establish', [PS_RRC_Setup_Success_Rate(UCell_Eric)] as'PS RRC SR', [HSDPA_Drop_Call_Rate(UCell_Eric)] as 'HSDPA Drop Rate' from [dbo].[RD3_Ericsson_Cell_Daily] where  (" + sites_list + ") and (" + date_list + ")" +
                    //  @" union all select [Date], substring(convert(varchar,Date,23),1,10) as 'Day', [ElementID] as 'RNC',  substring([ElementID1],1,9) as 'Sector', [ElementID1] as 'Cell', [PAYLOAD] as 'PS_Traffic_Daily (GB)', [PS_RAB_Setup_Success_Ratio] as 'PS RAB Establish', [PS_RRC_Connection_success_Rate_repeatless(Hu_Cell)] as'PS RRC SR',  [HSDPA_cdr(%)_(Hu_Cell)_new] as 'HSDPA Drop Rate' from [dbo].[RD3_Huawei_Cell_Daily] where (" + sites_list + ") and (" + date_list + ")" +
                    //  @" union all select [Date], substring(convert(varchar,Date,23),1,10) as 'Day', [ElementID] as 'RNC',  substring([ElementID1],1,9) as 'Sector', [ElementID1] as 'Cell', [PS_Payload_Total(HS+R99)(Nokia_CELL)_GB] as 'PS_Traffic_Daily (GB)', [RAB_Setup_and_Access_Complete_Ratio_for_NRT_Service_from_User_pe] as 'PS RAB Establish', [PS_RRCSETUP_SR] as 'PS RRC SR', [HSDPA_Call_Drop_Rate(Nokia_Cell)] as 'HSDPA Drop Rate' from [dbo].[RD3_Nokia_Cell_Daily] where (" + sites_list + ") and (" + date_list + ")";



                    string Data_Quary4 = @"
                select T1.[Day]
      ,T1.[RNC]
      ,T1.[Sector] ,
	  
	  case when Sum_PS_RRC_SR is not null then T1.[PS RRC SR] else null end as [PS RRC SR],
	  case when Sum_PS_Drop_Rate is not null then T1.[PS Drop Rate] else null end as [PS Drop Rate],
      case when Sum_RSSI is not null then T1.[RSSI] else null end as [RSSI],
      case when Sum_HS_User_THR is not null then T1.[HS User THR] else null end as [HS User THR],
      T1.[PS_Traffic_Daily (GB)]

      from (select Day, RNC, Sector,
       sum(PS_RRC_SR_T)/case when sum([PS_RRC_SR_T1])=0 then(-1) else sum([PS_RRC_SR_T1]) end as 'PS RRC SR',

       sum(PS_Drop_Rate_T)/case when sum([PS_Drop_Rate_T1])=0 then(-1) else sum([PS_Drop_Rate_T1]) end as 'PS Drop Rate',

       sum(RSSI_T)/case when sum([RSSI_T1])=0 then(-1) else sum([RSSI_T1]) end as 'RSSI',

       sum(HS_User_THR_T)/case when sum([HS_User_THR_T1])=0 then(-1) else sum([HS_User_THR_T1]) end as 'HS User THR',

       sum(cast([PS_Traffic_Daily (GB)] as float)) as 'PS_Traffic_Daily (GB)'


from
        (select Day, RNC, Sector,

        [PS RRC SR],
		(isnull([PS RRC SR],0)*[PS_Traffic_Daily (GB)]) PS_RRC_SR_T,
		case when[PS RRC SR] is null then 0
		else [PS_Traffic_Daily (GB)]
        end as PS_RRC_SR_T1,

		[PS Drop Rate],
		(isnull([PS Drop Rate],0)*[PS_Traffic_Daily (GB)]) PS_Drop_Rate_T,
		case when[PS Drop Rate] is null then 0
		else [PS_Traffic_Daily (GB)]
        end as PS_Drop_Rate_T1,


		[RSSI],
		(isnull([RSSI],0)*[PS_Traffic_Daily (GB)]) RSSI_T,
		case when[RSSI] is null then 0
		else [PS_Traffic_Daily (GB)]
        end as RSSI_T1,



		[HS User THR],
		(isnull([HS User THR],0)*[PS_Traffic_Daily (GB)]) HS_User_THR_T,
		case when[HS User THR] is null then 0
		else [PS_Traffic_Daily (GB)]
        end as HS_User_THR_T1,


		[PS_Traffic_Daily (GB)]

        from(" + Data_Quary3 + @" )as tbl )tb
group by Day, RNC, Sector
		
		) T1
inner join (SELECT [Day]
                 , [RNC]
                 , [Sector]

                 , sum(cast([PS RRC SR] as float)) Sum_PS_RRC_SR
                 ,sum(cast([PS Drop Rate] as float)) Sum_PS_Drop_Rate
                 ,sum(cast([RSSI] as float)) Sum_RSSI
                 ,sum(cast([HS User THR] as float)) Sum_HS_User_THR
            from(" + Data_Quary3 + @" )as tbl 
        group by[Day],[RNC],[Sector]	 
           )T2

  on T1.[Day]=T2.[Day] and
  T1.[RNC]=T2.[RNC] and
  T1.[Sector]=T2.[Sector]";


                    SqlCommand PS_Data_Quary = new SqlCommand(Data_Quary4, connection);
                    PS_Data_Quary.CommandTimeout = 0;
                    PS_Data_Quary.ExecuteNonQuery();
                    Data_Table_3G_PS = new DataTable();
                    SqlDataAdapter Date_Table2 = new SqlDataAdapter(PS_Data_Quary);
                    Date_Table2.Fill(Data_Table_3G_PS);

                    Data_Table_3G = new DataTable();

                    Data_Table_3G.Columns.Add("Day", typeof(string));
                    Data_Table_3G.Columns.Add("RNC", typeof(string));
                    Data_Table_3G.Columns.Add("Sector", typeof(string));
                    Data_Table_3G.Columns.Add("CS RRC SR, TH=95", typeof(string));
                    Data_Table_3G.Columns.Add("CS RAB Establish, TH=95", typeof(string));
                    Data_Table_3G.Columns.Add("Voice Drop Rate, TH=4", typeof(string));
                    Data_Table_3G.Columns.Add("Cell Availability, TH=99", typeof(string));
                    Data_Table_3G.Columns.Add("BH CS Traffic (Erlang), TH=0", typeof(string));
                    Data_Table_3G.Columns.Add("PS RRC SR, TH=95", typeof(string));
                    Data_Table_3G.Columns.Add("PS Drop Rate, TH=4", typeof(string));
                    Data_Table_3G.Columns.Add("RSSI (dBm), TH=-90", typeof(string));
                    Data_Table_3G.Columns.Add("HS User Throughput (Mbps), TH=1", typeof(string));
                    Data_Table_3G.Columns.Add("Daily PS Traffic (GB), TH=0", typeof(string));

                    Data_Table_3G.Columns.Add("CS RRC SR Score", typeof(int));
                    Data_Table_3G.Columns.Add("CS RAB Establish Score", typeof(int));
                    Data_Table_3G.Columns.Add("Voice Drop Rate Score", typeof(int));
                    Data_Table_3G.Columns.Add("Cell Availability Score", typeof(int));
                    Data_Table_3G.Columns.Add("CS Traffic Score", typeof(int));

                    Data_Table_3G.Columns.Add("PS RRC SR Score", typeof(int));
                    Data_Table_3G.Columns.Add("PS Drop Rate Score", typeof(int));
                    Data_Table_3G.Columns.Add("RSSI Score", typeof(int));
                    Data_Table_3G.Columns.Add("THR Score", typeof(int));
                    Data_Table_3G.Columns.Add("PS Traffic Score", typeof(int));

                    Data_Table_3G.Columns.Add("Cell Score", typeof(double));
                    Data_Table_3G.Columns.Add("Site", typeof(string));




                    Site_Data_Table_3G = new DataTable();
                    Site_Data_Table_3G.Columns.Add("Site", typeof(string));
                    Site_Data_Table_3G.Columns.Add("KPI Zero Status", typeof(string));
                    Site_Data_Table_3G.Columns.Add("Rejected Cell List", typeof(string));



                    //dataGridView1.ColumnCount = 14;
                    dataGridView1.Invoke(new Action(() => dataGridView1.ColumnCount = 15));

                    // dataGridView1.Rows.Clear();
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows.Clear()));
                    //dataGridView1.RowCount = Data_Table_3G_CS.Rows.Count + 1;
                    dataGridView1.Invoke(new Action(() => dataGridView1.RowCount = Data_Table_3G_CS.Rows.Count + 1));

                    //dataGridView1.Rows[0].Cells[0].Value = "Date"; dataGridView1.Columns[0].Width = 100;
                    //dataGridView1.Rows[0].Cells[1].Value = "RNC"; dataGridView1.Columns[1].Width = 100;
                    //dataGridView1.Rows[0].Cells[2].Value = "Site"; dataGridView1.Columns[2].Width = 100;
                    //dataGridView1.Rows[0].Cells[3].Value = "Sector"; dataGridView1.Columns[3].Width = 100;
                    //dataGridView1.Rows[0].Cells[4].Value = "CS RRC SR"; dataGridView1.Columns[4].Width = 100;
                    //dataGridView1.Rows[0].Cells[5].Value = "CS RAB Establish"; dataGridView1.Columns[5].Width = 100;
                    //dataGridView1.Rows[0].Cells[6].Value = "Voice Drop Rate"; dataGridView1.Columns[6].Width = 100;
                    //dataGridView1.Rows[0].Cells[7].Value = "Cell Availability"; dataGridView1.Columns[7].Width = 100;
                    //dataGridView1.Rows[0].Cells[8].Value = "BH CS Traffic (Erlang)"; dataGridView1.Columns[8].Width = 100;
                    //dataGridView1.Rows[0].Cells[9].Value = "PS RRC SR"; dataGridView1.Columns[9].Width = 100;
                    //dataGridView1.Rows[0].Cells[10].Value = "PS RAB Establish"; dataGridView1.Columns[10].Width = 100;
                    //dataGridView1.Rows[0].Cells[11].Value = "HSDPA Drop Rate"; dataGridView1.Columns[11].Width = 100;
                    //////dataGridView1.Rows[0].Cells[12].Value = "RSSI (dBm)"; dataGridView1.Columns[12].Width = 100;
                    //dataGridView1.Rows[0].Cells[12].Value = "Daily PS Traffic (GB)"; dataGridView1.Columns[12].Width = 100;
                    //dataGridView1.Rows[0].Cells[13].Value = "Cell Status"; dataGridView1.Columns[13].Width = 100;


                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[0].Value = "Date")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[0].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[1].Value = "RNC")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[1].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[2].Value = "Site")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[2].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[3].Value = "Sector")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[3].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[4].Value = "CS RRC SR")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[4].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[5].Value = "CS RAB Establish")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[5].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[6].Value = "Voice Drop Rate")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[6].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[7].Value = "Cell Availability")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[7].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[8].Value = "BH CS Traffic (Erlang)")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[8].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[9].Value = "PS RRC SR")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[9].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[10].Value = "PS Drop Rate")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[10].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[11].Value = "RSSI (dBm)")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[11].Width = 100));

                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[12].Value = "HS User Throughput (Mbps)")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[12].Width = 100));

                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[13].Value = "Daily PS Traffic (GB)")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[13].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[14].Value = "Cell Status")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[14].Width = 100));




                    progressBar1.Minimum = 0;
                    //progressBar1.Maximum = Data_Table_3G_CS.Rows.Count - 1;


                    if (Data_Table_3G_CS.Rows.Count == 0)
                    {
                        MessageBox.Show("There is no Data in Database");
                    }

                    if (Data_Table_3G_CS.Rows.Count != 0)
                    {
                        // progressBar1.Maximum = Data_Table_3G_CS.Rows.Count - 1;
                        progressBar1.Invoke(new Action(() => progressBar1.Maximum = Data_Table_3G_CS.Rows.Count - 1));
                        for (int k = 0; k < Data_Table_3G_CS.Rows.Count; k++)
                        {
                            int result = 0;

                            Data_Table_3G.Rows.Add(Data_Table_3G_CS.Rows[k][0], Data_Table_3G_CS.Rows[k][1], Data_Table_3G_CS.Rows[k][2], Data_Table_3G_CS.Rows[k][3], Data_Table_3G_CS.Rows[k][4], Data_Table_3G_CS.Rows[k][5], Data_Table_3G_CS.Rows[k][6], Data_Table_3G_CS.Rows[k][7], "", "", "", "", "", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "");

                            for (int j = 0; j < Data_Table_3G_PS.Rows.Count; j++)
                            {
                                if ((Data_Table_3G_CS.Rows[k][0].ToString() == Data_Table_3G_PS.Rows[j][0].ToString()) && (Data_Table_3G_CS.Rows[k][2].ToString() == Data_Table_3G_PS.Rows[j][2].ToString()))
                                {
                                    Data_Table_3G.Rows[k][8] = Data_Table_3G_PS.Rows[j][3];
                                    Data_Table_3G.Rows[k][9] = Data_Table_3G_PS.Rows[j][4];
                                    Data_Table_3G.Rows[k][10] = Data_Table_3G_PS.Rows[j][5];
                                    Data_Table_3G.Rows[k][11] = Data_Table_3G_PS.Rows[j][6];
                                    Data_Table_3G.Rows[k][12] = Data_Table_3G_PS.Rows[j][7];
                                }
                            }



                            // Thresholds
                            double CS_RRC_SR_TH = 95;
                            double CS_Drop_TH = 4;
                            double CS_RAB_TH = 95;
                            double CS_Traffic_TH = 0;
                            double PS_RRC_SR_TH = 95;
                            double PS_Drop_TH = 4;
                            double RSSI_TH = -90;
                            double THR_TH = 1;
                            double PS_Payload_TH = 0;
                            double Availability_TH = 99;




                            // Date
                            //  dataGridView1.Rows[k + 1].Cells[0].Value = Data_Table_3G.Rows[k][0];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[0].Value = Data_Table_3G.Rows[k][0]));
                            //RNC
                            //  dataGridView1.Rows[k + 1].Cells[1].Value = Data_Table_3G.Rows[k][1];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[1].Value = Data_Table_3G.Rows[k][1]));
                            // Site
                            string Cell = Data_Table_3G.Rows[k][2].ToString();
                            //dataGridView1.Rows[k + 1].Cells[2].Value = Cell.Substring(0, 8);
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[2].Value = Cell.Substring(0, 8)));
                            Data_Table_3G.Rows[k][24] = Cell.Substring(0, 8);
                            // Sector
                            // dataGridView1.Rows[k + 1].Cells[3].Value = Data_Table_3G.Rows[k][2];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[3].Value = Data_Table_3G.Rows[k][2]));


                            // CS RRC SR
                            // dataGridView1.Rows[k + 1].Cells[4].Value = Data_Table_3G.Rows[k][3];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[4].Value = Data_Table_3G.Rows[k][3]));
                            if (Data_Table_3G.Rows[k][3].ToString() == "")
                            {
                                Data_Table_3G.Rows[k][13] = -1;
                            }
                            else if (Convert.ToDouble(Data_Table_3G.Rows[k][3]) < CS_RRC_SR_TH)
                            {
                                Data_Table_3G.Rows[k][13] = 1; result++; dataGridView1.Rows[k + 1].Cells[4].Style.BackColor = Color.Orange;
                            }
                            else if (Convert.ToDouble(Data_Table_3G.Rows[k][3]) >= CS_RRC_SR_TH)
                            {
                                Data_Table_3G.Rows[k][13] = 0;
                            }


                            // CS RAB Setablish
                            //  dataGridView1.Rows[k + 1].Cells[5].Value = Data_Table_3G.Rows[k][4];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[5].Value = Data_Table_3G.Rows[k][4]));
                            if (Data_Table_3G.Rows[k][4].ToString() == "")
                            {
                                Data_Table_3G.Rows[k][14] = -1;
                            }
                            else if (Convert.ToDouble(Data_Table_3G.Rows[k][4]) < CS_RAB_TH)
                            {
                                Data_Table_3G.Rows[k][14] = 1; result++; dataGridView1.Rows[k + 1].Cells[5].Style.BackColor = Color.Orange;
                            }
                            else if (Convert.ToDouble(Data_Table_3G.Rows[k][4]) >= CS_RAB_TH)
                            {
                                Data_Table_3G.Rows[k][14] = 0;
                            }


                            // Voice Drop Rate
                            //dataGridView1.Rows[k + 1].Cells[6].Value = Data_Table_3G.Rows[k][5];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[6].Value = Data_Table_3G.Rows[k][5]));
                            if (Data_Table_3G.Rows[k][5].ToString() == "")
                            {
                                Data_Table_3G.Rows[k][15] = -1;
                            }
                            else if (Convert.ToDouble(Data_Table_3G.Rows[k][5]) > CS_Drop_TH)
                            {
                                Data_Table_3G.Rows[k][15] = 1; result++; dataGridView1.Rows[k + 1].Cells[6].Style.BackColor = Color.Orange;
                            }
                            else if (Convert.ToDouble(Data_Table_3G.Rows[k][5]) <= CS_Drop_TH)
                            {
                                Data_Table_3G.Rows[k][15] = 0;
                            }


                            // Availability
                            //dataGridView1.Rows[k + 1].Cells[7].Value = Data_Table_3G.Rows[k][6];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[7].Value = Data_Table_3G.Rows[k][6]));
                            if (Data_Table_3G.Rows[k][6].ToString() == "")
                            {
                                Data_Table_3G.Rows[k][16] = -1;
                            }
                            else if (Convert.ToDouble(Data_Table_3G.Rows[k][6]) < Availability_TH && Convert.ToDouble(Data_Table_3G.Rows[k][6]) > 0)
                            {
                                Data_Table_3G.Rows[k][16] = 1; result++; dataGridView1.Rows[k + 1].Cells[7].Style.BackColor = Color.Orange;
                            }
                            else if (Convert.ToDouble(Data_Table_3G.Rows[k][6]) >= Availability_TH)
                            {
                                Data_Table_3G.Rows[k][16] = 0;
                            }


                            // CS Traffic
                            // dataGridView1.Rows[k + 1].Cells[8].Value = Data_Table_3G.Rows[k][7];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[8].Value = Data_Table_3G.Rows[k][7]));
                            if (Data_Table_3G.Rows[k][7].ToString() == "")
                            {
                                Data_Table_3G.Rows[k][17] = -1;
                            }
                            else if (Convert.ToDouble(Data_Table_3G.Rows[k][7]) == CS_Traffic_TH)
                            {
                                Data_Table_3G.Rows[k][17] = 1; result++; dataGridView1.Rows[k + 1].Cells[8].Style.BackColor = Color.Orange;
                            }
                            else if (Convert.ToDouble(Data_Table_3G.Rows[k][7]) > CS_Traffic_TH)
                            {
                                Data_Table_3G.Rows[k][17] = 0;
                            }


                            // PS RRC SR
                            //dataGridView1.Rows[k + 1].Cells[9].Value = Data_Table_3G.Rows[k][8];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[9].Value = Data_Table_3G.Rows[k][8]));
                            if (Data_Table_3G.Rows[k][8].ToString() == "")
                            {
                                Data_Table_3G.Rows[k][18] = -1;
                            }
                            else if (Convert.ToDouble(Data_Table_3G.Rows[k][8]) < PS_RRC_SR_TH)
                            {
                                Data_Table_3G.Rows[k][18] = 1; result++; dataGridView1.Rows[k + 1].Cells[9].Style.BackColor = Color.Orange;
                            }
                            else if (Convert.ToDouble(Data_Table_3G.Rows[k][8]) >= PS_RRC_SR_TH)
                            {
                                Data_Table_3G.Rows[k][18] = 0;
                            }

                            // PS Drop Rate
                            // dataGridView1.Rows[k + 1].Cells[10].Value = Data_Table_3G.Rows[k][9];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[10].Value = Data_Table_3G.Rows[k][9]));
                            if (Data_Table_3G.Rows[k][9].ToString() == "")
                            {
                                Data_Table_3G.Rows[k][19] = -1;
                            }
                            else if (Convert.ToDouble(Data_Table_3G.Rows[k][9]) > PS_Drop_TH)
                            {
                                Data_Table_3G.Rows[k][19] = 1; result++; dataGridView1.Rows[k + 1].Cells[10].Style.BackColor = Color.Orange;
                            }
                            else if (Convert.ToDouble(Data_Table_3G.Rows[k][9]) <= PS_Drop_TH)
                            {
                                Data_Table_3G.Rows[k][19] = 0;
                            }


                            // RSSI
                            // dataGridView1.Rows[k + 1].Cells[11].Value = Data_Table_3G.Rows[k][10];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[11].Value = Data_Table_3G.Rows[k][10]));
                            if (Data_Table_3G.Rows[k][10].ToString() == "")
                            {
                                Data_Table_3G.Rows[k][20] = -1;
                            }
                            else if (Convert.ToDouble(Data_Table_3G.Rows[k][10]) > RSSI_TH)
                            {
                                Data_Table_3G.Rows[k][20] = 1; result++; dataGridView1.Rows[k + 1].Cells[11].Style.BackColor = Color.Orange;
                            }
                            else if (Convert.ToDouble(Data_Table_3G.Rows[k][10]) <= RSSI_TH)
                            {
                                Data_Table_3G.Rows[k][20] = 0;
                            }




                            // HS User THR
                            // dataGridView1.Rows[k + 1].Cells[11].Value = Data_Table_3G.Rows[k][10];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[12].Value = Data_Table_3G.Rows[k][11]));
                            if (Data_Table_3G.Rows[k][11].ToString() == "")
                            {
                                Data_Table_3G.Rows[k][21] = -1;
                            }
                            else if (Convert.ToDouble(Data_Table_3G.Rows[k][11]) < THR_TH)
                            {
                                Data_Table_3G.Rows[k][21] = 1; result++; dataGridView1.Rows[k + 1].Cells[12].Style.BackColor = Color.Orange;
                            }
                            else if (Convert.ToDouble(Data_Table_3G.Rows[k][11]) >= THR_TH)
                            {
                                Data_Table_3G.Rows[k][21] = 0;
                            }



                            // RSSI
                            //dataGridView1.Rows[k + 1].Cells[12].Value = Data_Table_3G.Rows[k][11];
                            //string Band = Cell.Substring(2, 2);
                            //if (Data_Table_3G.Rows[k][11].ToString() == "")
                            //{
                            //    Data_Table_3G.Rows[k][21] = -1;
                            //}
                            //else if ((Convert.ToDouble(Data_Table_3G.Rows[k][11]) > Convert.ToDouble(textBox17.Text)) && (Band == "2U" || Band == "5U" || Band == "8U"))
                            //{
                            //    Data_Table_3G.Rows[k][21] = 1; result++; dataGridView1.Rows[k + 1].Cells[12].Style.BackColor = Color.Orange;
                            //}
                            //else if ((Convert.ToDouble(Data_Table_3G.Rows[k][11]) > Convert.ToDouble(textBox18.Text)) && (Band == "1U" || Band == "4U" || Band == "7U"))
                            //{
                            //    Data_Table_3G.Rows[k][21] = 1; result++; dataGridView1.Rows[k + 1].Cells[12].Style.BackColor = Color.Orange;
                            //}
                            //else if ((Convert.ToDouble(Data_Table_3G.Rows[k][11]) <= Convert.ToDouble(textBox17.Text)) && (Band == "2U" || Band == "5U" || Band == "8U"))
                            //{
                            //    Data_Table_3G.Rows[k][21] = 0;
                            //}
                            //else if ((Convert.ToDouble(Data_Table_3G.Rows[k][11]) <= Convert.ToDouble(textBox18.Text)) && (Band == "1U" || Band == "4U" || Band == "7U"))
                            //{
                            //    Data_Table_3G.Rows[k][21] = 0;
                            //}



                            // PS Traffic
                            //dataGridView1.Rows[k + 1].Cells[12].Value = Data_Table_3G.Rows[k][11];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[13].Value = Data_Table_3G.Rows[k][12]));
                                if (Data_Table_3G.Rows[k][12].ToString() == "")
                                {
                                    Data_Table_3G.Rows[k][22] = -1;
                                }
                                else if (Convert.ToDouble(Data_Table_3G.Rows[k][12]) == PS_Payload_TH)
                                {
                                    Data_Table_3G.Rows[k][22] = 1; result++; dataGridView1.Rows[k + 1].Cells[13].Style.BackColor = Color.Orange;
                                }
                                else if (Convert.ToDouble(Data_Table_3G.Rows[k][12]) > PS_Payload_TH)
                                {
                                    Data_Table_3G.Rows[k][22] = 0;
                                }



                                if ((Convert.ToInt16(Data_Table_3G.Rows[k][17]) == -1) || (Convert.ToInt16(Data_Table_3G.Rows[k][22]) == -1))
                                {
                                    //         if (Convert.ToInt16(Data_Table_3G.Rows[k][21])!= 0)
                                    //  {
                                    Data_Table_3G.Rows[k][23] = 1; dataGridView1.Rows[k + 1].Cells[14].Value = "Not Updated";
                                    //                             }
                                }
                                else if (result == 0)
                                {
                                    Data_Table_3G.Rows[k][23] = 0.1; dataGridView1.Rows[k + 1].Cells[14].Value = "Passed";
                                }
                                if (result > 0)
                                {
                                    Data_Table_3G.Rows[k][23] = 0; dataGridView1.Rows[k + 1].Cells[14].Value = "Rejected";
                                }
                                progressBar1.Invoke(new Action(() => progressBar1.Value = k));
                                //   progressBar1.Value = k;


                        }
                            var distinctIds = Data_Table_3G.AsEnumerable()
               .Select(s => new
               {
                   id = s.Field<string>("Site"),
               })
               .Distinct().ToList();

                            for (int j = 0; j < distinctIds.Count; j++)
                            {
                                var cell_data = (from p in Data_Table_3G.AsEnumerable()
                                                 where p.Field<string>("Site") == distinctIds[j].id
                                                 select p).ToList();


                                double multiplier = 1;
                                for (int h = 0; h < cell_data.Count; h++)
                                {
                                    multiplier = multiplier * Convert.ToDouble(cell_data[h].ItemArray[23]);

                                }

                                if (multiplier > 0 && multiplier < 1)
                                {
                                    Site_Data_Table_3G.Rows.Add(distinctIds[j].id, "Passed");
                                }
                                if (multiplier == 0)
                                {
                                    Site_Data_Table_3G.Rows.Add(distinctIds[j].id, "Rejected");
                                }
                                if (multiplier == 1)
                                {
                                    Site_Data_Table_3G.Rows.Add(distinctIds[j].id, "Not Updated");
                                }

                            }

                        

                    }



                }


                if (Technology == "4G")
                {

                    string Data_Quary3 = @" select [Datetime], substring([eNodeB],1,9) as 'Sector', [eNodeB] as 'Cell', [Total_Volume(UL+DL)(GB)(eNodeB_Eric)] as 'PS_Traffic_Daily (GB)', [Average_UE_DL_Throughput(Mbps)(eNodeB_Eric)] as 'UE DL THR (Mbps)', [Average_UE_UL_Throughput(Mbps)(eNodeB_Eric)] as'UE UL THR (Mbps)', [E_RAB_Drop_Rate(eNodeB_Eric)] as 'ERAB Drop Rate', [E-RAB_Setup_SR_incl_added_New(EUCell_Eric)] as 'ERAB Setup SR', [IntraF_Handover_Execution(eNodeB_Eric)] as 'Intra Freq HO SR' , [RRC_Estab_Success_Rate(ReAtt)(EUCell_Eric)] as 'RRC Connection SR' , [Cell_Availability_Rate_Exclude_Blocking(Cell_EricLTE)] as 'Cell Availability' from [dbo].[TBL_LTE_CELL_Daily_E] where  (" + EH_sites_list + ") and (" + EH_date_list + ")" +
      @" union all select [Datetime],  substring([eNodeB],1,9) as 'Sector', [eNodeB] as 'Cell', [Total_Traffic_Volume(GB)] as 'PS_Traffic_Daily (GB)', [Average_Downlink_User_Throughput(Mbit/s)] as 'UE DL THR (Mbps)', [Average_UPlink_User_Throughput(Mbit/s)] as'UE UL THR (Mbps)',  [Call_Drop_Rate] as 'ERAB Drop Rate',  [E-RAB_Setup_Success_Rate(Hu_Cell)] as 'ERAB Setup SR'  , [IntraF_HOOut_SR] as 'Intra Freq HO SR' , [RRC_Connection_Setup_Success_Rate_service] as 'RRC Connection SR' , [Cell_Availability_Rate_Exclude_Blocking(Cell_Hu)] as 'Cell Availability' from [dbo].[TBL_LTE_CELL_Daily_H] where (" + EH_sites_list + ") and (" + EH_date_list + ")" +
      @" union all select [Date],  substring([ElementID1],1,9) as 'Sector', [ElementID1] as 'Cell', [Total_Payload_GB(Nokia_LTE_CELL)] as 'PS_Traffic_Daily (GB)', [User_Throughput_DL_mbps(Nokia_LTE_CELL)] as 'UE DL THR (Mbps)', [User_Throughput_UL_mbps(Nokia_LTE_CELL)] as 'UE UL THR (Mbps)', [E-RAB_Drop_Ratio_RAN_View(Nokia_LTE_CELL)] as 'ERAB Drop Rate', [E-RAB_Setup_SR_incl_added(Nokia_LTE_CELL)] as 'ERAB Setup SR' , [HO_Success_Ratio_intra_eNB(Nokia_LTE_CELL)] as 'Intra Freq HO SR' , [RRC_Connection_Setup_Success_Ratio(Nokia_LTE_CELL)] as 'RRC Connection SR' , [cell_availability_exclude_manual_blocking(Nokia_LTE_CELL)] as 'Cell Availability' from [dbo].[TBL_LTE_CELL_Daily_N] where (" + N_sites_list + ") and (" + N_date_list + ")";



                    string Data_Quary4 = @"
                select T1.[Datetime]
      ,T1.[Sector] ,
	  
	  case when Sum_UE_DL_THR is not null then T1.[UE DL THR (Mbps)] else null end as [UE DL THR (Mbps)],
	  case when Sum_UE_UL_THR is not null then T1.[UE UL THR (Mbps)] else null end as [UE UL THR (Mbps)],
          case when Sum_ERAB_Drop_Rate is not null then T1.[ERAB Drop Rate] else null end as [ERAB Drop Rate],
          case when Sum_ERAB_Setup_SR is not null then T1.[ERAB Setup SR] else null end as [ERAB Setup SR],
          case when Sum_Intra_Freq_HO_SR is not null then T1.[Intra Freq HO SR] else null end as [Intra Freq HO SR],
          case when Sum_RRC_Connection_SR is not null then T1.[RRC Connection SR] else null end as [RRC Connection SR],
          case when Sum_Cell_Availability is not null then T1.[Cell Availability] else null end as [Cell Availability],
          T1.[PS_Traffic_Daily (GB)]

      from (select Datetime, Sector,
       sum(UE_DL_THR_T)/case when sum([UE_DL_THR_T1])=0 then(-1) else sum([UE_DL_THR_T1]) end as 'UE DL THR (Mbps)',

       sum(UE_UL_THR_T)/case when sum([UE_UL_THR_T1])=0 then(-1) else sum([UE_UL_THR_T1]) end as 'UE UL THR (Mbps)',

       sum(ERAB_Drop_Rate_T)/case when sum([ERAB_Drop_Rate_T1])=0 then(-1) else sum([ERAB_Drop_Rate_T1]) end as 'ERAB Drop Rate',

       sum(ERAB_Setup_SR_T)/case when sum([ERAB_Setup_SR_T1])=0 then(-1) else sum([ERAB_Setup_SR_T1]) end as 'ERAB Setup SR',

       sum(Intra_Freq_HO_SR_T)/case when sum([Intra_Freq_HO_SR_T1])=0 then(-1) else sum([Intra_Freq_HO_SR_T1]) end as 'Intra Freq HO SR',

       sum(RRC_Connection_SR_T)/case when sum([RRC_Connection_SR_T1])=0 then(-1) else sum([RRC_Connection_SR_T1]) end as 'RRC Connection SR',

       sum(Cell_Availability_T)/case when sum([Cell_Availability_T1])=0 then(-1) else sum([Cell_Availability_T1]) end as 'Cell Availability',

       sum(cast([PS_Traffic_Daily (GB)] as float)) as 'PS_Traffic_Daily (GB)'


from
        (select Datetime, Sector,

        [UE DL THR (Mbps)],
		(isnull([UE DL THR (Mbps)],0)*[PS_Traffic_Daily (GB)]) UE_DL_THR_T,
		case when[UE DL THR (Mbps)] is null then 0
		else [PS_Traffic_Daily (GB)]
        end as UE_DL_THR_T1,

		[UE UL THR (Mbps)],
		(isnull([UE UL THR (Mbps)],0)*[PS_Traffic_Daily (GB)]) UE_UL_THR_T,
		case when[UE UL THR (Mbps)] is null then 0
		else [PS_Traffic_Daily (GB)]
        end as UE_UL_THR_T1,

		[ERAB Drop Rate],
		(isnull([ERAB Drop Rate],0)*[PS_Traffic_Daily (GB)]) ERAB_Drop_Rate_T,
		case when[ERAB Drop Rate] is null then 0
		else [PS_Traffic_Daily (GB)]
        end as ERAB_Drop_Rate_T1,

		[ERAB Setup SR],
		(isnull([ERAB Setup SR],0)*[PS_Traffic_Daily (GB)]) ERAB_Setup_SR_T,
		case when[ERAB Setup SR] is null then 0
		else [PS_Traffic_Daily (GB)]
        end as ERAB_Setup_SR_T1,

		[Intra Freq HO SR],
		(isnull([Intra Freq HO SR],0)*[PS_Traffic_Daily (GB)]) Intra_Freq_HO_SR_T,
		case when[Intra Freq HO SR] is null then 0
		else [PS_Traffic_Daily (GB)]
        end as Intra_Freq_HO_SR_T1,

		[RRC Connection SR],
		(isnull([RRC Connection SR],0)*[PS_Traffic_Daily (GB)]) RRC_Connection_SR_T,
		case when[RRC Connection SR] is null then 0
		else [PS_Traffic_Daily (GB)]
        end as RRC_Connection_SR_T1,

		[Cell Availability],
		(isnull([Cell Availability],0)*[PS_Traffic_Daily (GB)]) Cell_Availability_T,
		case when[Cell Availability] is null then 0
		else [PS_Traffic_Daily (GB)]
        end as Cell_Availability_T1,

		[PS_Traffic_Daily (GB)]

        from(" + Data_Quary3 + @" )as tbl )tb
group by Datetime, Sector
		
		) T1
inner join (SELECT [Datetime]
                 
                 , [Sector]

                 , sum(cast([UE DL THR (Mbps)] as float)) Sum_UE_DL_THR
		 ,sum(cast([UE UL THR (Mbps)] as float)) Sum_UE_UL_THR
                 ,sum(cast([ERAB Drop Rate] as float)) Sum_ERAB_Drop_Rate
                 ,sum(cast([ERAB Setup SR] as float)) Sum_ERAB_Setup_SR
                 ,sum(cast([Intra Freq HO SR] as float)) Sum_Intra_Freq_HO_SR
                 ,sum(cast([RRC Connection SR] as float)) Sum_RRC_Connection_SR
                 ,sum(cast([Cell Availability] as float)) Sum_Cell_Availability
            from(" + Data_Quary3 + @" )as tbl 
        group by[Datetime], [Sector]	 
           )T2

  on T1.[Datetime]=T2.[Datetime] and
  T1.[Sector]=T2.[Sector]";




                    SqlCommand PS_Data_Quary = new SqlCommand(Data_Quary4, connection);
                    PS_Data_Quary.CommandTimeout = 0;
                    PS_Data_Quary.ExecuteNonQuery();
                    Data_Table_4G = new DataTable();
                    SqlDataAdapter Date_Table2 = new SqlDataAdapter(PS_Data_Quary);
                    Date_Table2.Fill(Data_Table_4G);




                    //Data_Table_4G.Columns.Add("Date", typeof(string));
                    //Data_Table_4G.Columns.Add("Sector", typeof(string));
                    //Data_Table_4G.Columns.Add("UE DL THR (Mbps)", typeof(string));
                    //Data_Table_4G.Columns.Add("UE UL THR (Mbps)", typeof(string));
                    //Data_Table_4G.Columns.Add("ERAB Drop Rate", typeof(string));
                    //Data_Table_4G.Columns.Add("ERAB Setup SR", typeof(string));
                    //Data_Table_4G.Columns.Add("Intra Freq HO SR", typeof(string));
                    //Data_Table_4G.Columns.Add("RRC Connection SR", typeof(string));
                    //Data_Table_4G.Columns.Add("Cell Availability", typeof(string));
                    //Data_Table_4G.Columns.Add("Daily PS Traffic (GB)", typeof(string));

                    Data_Table_4G.Columns.Add("UE DL THR (Mbps) Score", typeof(int));
                    Data_Table_4G.Columns.Add("UE UL THR (Mbps) Score", typeof(int));
                    Data_Table_4G.Columns.Add("ERAB Drop Rat Score", typeof(int));
                    Data_Table_4G.Columns.Add("ERAB Setup SR Score", typeof(int));
                    Data_Table_4G.Columns.Add("Intra Freq HO SR Score", typeof(int));
                    Data_Table_4G.Columns.Add("RRC Connection SR Score", typeof(int));
                    Data_Table_4G.Columns.Add("Cell Availability Score", typeof(int));
                    Data_Table_4G.Columns.Add("Daily PS Traffic (GB) Score", typeof(int));
                    Data_Table_4G.Columns.Add("Cell Score", typeof(double));
                    Data_Table_4G.Columns.Add("Site", typeof(string));





                    Site_Data_Table_4G = new DataTable();
                    Site_Data_Table_4G.Columns.Add("Site", typeof(string));
                    Site_Data_Table_4G.Columns.Add("KPI Zero Status", typeof(string));
                    Site_Data_Table_4G.Columns.Add("Rejected Cell List", typeof(string));



                    //dataGridView1.ColumnCount = 12;
                    dataGridView1.Invoke(new Action(() => dataGridView1.ColumnCount = 12));

                    //dataGridView1.Rows.Clear();
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows.Clear()));

                    //dataGridView1.RowCount = Data_Table_4G.Rows.Count + 1;
                    dataGridView1.Invoke(new Action(() => dataGridView1.RowCount = Data_Table_4G.Rows.Count + 1));

                    //dataGridView1.Rows[0].Cells[0].Value = "Date"; dataGridView1.Columns[0].Width = 100;
                    //dataGridView1.Rows[0].Cells[1].Value = "Site"; dataGridView1.Columns[1].Width = 100;
                    //dataGridView1.Rows[0].Cells[2].Value = "Sector"; dataGridView1.Columns[2].Width = 100;
                    //dataGridView1.Rows[0].Cells[3].Value = "UE DL THR (Mbps)"; dataGridView1.Columns[3].Width = 100;
                    //dataGridView1.Rows[0].Cells[4].Value = "UE UL THR (Mbps)"; dataGridView1.Columns[4].Width = 100;
                    //dataGridView1.Rows[0].Cells[5].Value = "ERAB Drop Rate"; dataGridView1.Columns[5].Width = 100;
                    //dataGridView1.Rows[0].Cells[6].Value = "ERAB Setup SR"; dataGridView1.Columns[6].Width = 100;
                    //dataGridView1.Rows[0].Cells[7].Value = "Intra Freq HO SR"; dataGridView1.Columns[7].Width = 100;
                    //dataGridView1.Rows[0].Cells[8].Value = "RRC Connection SR"; dataGridView1.Columns[8].Width = 100;
                    //dataGridView1.Rows[0].Cells[9].Value = "Cell Availability"; dataGridView1.Columns[9].Width = 100;
                    //dataGridView1.Rows[0].Cells[10].Value = "Daily PS Traffic (GB)"; dataGridView1.Columns[10].Width = 100;
                    //dataGridView1.Rows[0].Cells[11].Value = "Cell Status"; dataGridView1.Columns[11].Width = 100;




                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[0].Value = "Date")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[0].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[1].Value = "Site")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[1].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[2].Value = "Sector")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[2].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[3].Value = "UE DL THR (Mbps)")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[3].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[4].Value = "UE UL THR (Mbps)")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[4].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[5].Value = "ERAB Drop Rate")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[5].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[6].Value = "ERAB Setup SR")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[6].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[7].Value = "Intra Freq HO SR")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[7].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[8].Value = "RRC Connection SR")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[8].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[9].Value = "Cell Availability")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[9].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[10].Value = "Daily PS Traffic (GB)")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[10].Width = 100));
                    dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[11].Value = "Cell Status")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[11].Width = 100));


                    progressBar1.Minimum = 0;


                    if (Data_Table_4G.Rows.Count == 0)
                    {
                        MessageBox.Show("There is no Data in Database");
                    }

                    if (Data_Table_4G.Rows.Count != 0)
                    {
                        //  progressBar1.Maximum = Data_Table_4G.Rows.Count - 1;
                        progressBar1.Invoke(new Action(() => progressBar1.Maximum = Data_Table_4G.Rows.Count - 1));
                        for (int k = 0; k < Data_Table_4G.Rows.Count; k++)
                        {
                            int result = 0;

                            // Date
                            //dataGridView1.Rows[k + 1].Cells[0].Value = Data_Table_4G.Rows[k][0];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[0].Value = Data_Table_4G.Rows[k][0]));


                            // Site
                            string Cell = Data_Table_4G.Rows[k][1].ToString();
                            //  dataGridView1.Rows[k + 1].Cells[1].Value = Cell.Substring(0, 8);
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[1].Value = Cell.Substring(0, 8)));
                            Data_Table_4G.Rows[k][19] = Cell.Substring(0, 8);


                            // Cell
                            //dataGridView1.Rows[k + 1].Cells[2].Value = Data_Table_4G.Rows[k][1];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[2].Value = Data_Table_4G.Rows[k][1]));


                            // UE DL THR
                            //  dataGridView1.Rows[k + 1].Cells[3].Value = Data_Table_4G.Rows[k][2];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[3].Value = Data_Table_4G.Rows[k][2]));
                            if (Data_Table_4G.Rows[k][2].ToString() == "")
                            {
                                Data_Table_4G.Rows[k][10] = -1;
                            }
                            else if (Convert.ToDouble(Data_Table_4G.Rows[k][2]) < Convert.ToDouble(textBox19.Text))
                            {
                                Data_Table_4G.Rows[k][10] = 1; result++; dataGridView1.Rows[k + 1].Cells[3].Style.BackColor = Color.Orange;
                            }
                            else if (Convert.ToDouble(Data_Table_4G.Rows[k][2]) >= Convert.ToDouble(textBox19.Text))
                            {
                                Data_Table_4G.Rows[k][10] = 0;
                            }

                            // UE UL THR
                            //dataGridView1.Rows[k + 1].Cells[4].Value = Data_Table_4G.Rows[k][3];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[4].Value = Data_Table_4G.Rows[k][3]));
                            if (Data_Table_4G.Rows[k][3].ToString() == "")
                            {
                                Data_Table_4G.Rows[k][11] = -1;
                            }
                            else if (Convert.ToDouble(Data_Table_4G.Rows[k][3]) < Convert.ToDouble(textBox20.Text))
                            {
                                Data_Table_4G.Rows[k][11] = 1; result++; dataGridView1.Rows[k + 1].Cells[4].Style.BackColor = Color.Orange;
                            }
                            else if (Convert.ToDouble(Data_Table_4G.Rows[k][3]) >= Convert.ToDouble(textBox20.Text))
                            {
                                Data_Table_4G.Rows[k][11] = 0;
                            }


                            // ERAB Drop Rate
                            // dataGridView1.Rows[k + 1].Cells[5].Value = Data_Table_4G.Rows[k][4];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[5].Value = Data_Table_4G.Rows[k][4]));
                            if (Data_Table_4G.Rows[k][4].ToString() == "")
                            {
                                Data_Table_4G.Rows[k][12] = -1;
                            }
                            else if (Convert.ToDouble(Data_Table_4G.Rows[k][4]) > Convert.ToDouble(textBox21.Text))
                            {
                                Data_Table_4G.Rows[k][12] = 1; result++; dataGridView1.Rows[k + 1].Cells[5].Style.BackColor = Color.Orange;
                            }
                            else if (Convert.ToDouble(Data_Table_4G.Rows[k][4]) <= Convert.ToDouble(textBox21.Text))
                            {
                                Data_Table_4G.Rows[k][12] = 0;
                            }



                            // ERAB Setup SR
                            // dataGridView1.Rows[k + 1].Cells[6].Value = Data_Table_4G.Rows[k][3];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[6].Value = Data_Table_4G.Rows[k][5]));
                            if (Data_Table_4G.Rows[k][5].ToString() == "")
                            {
                                Data_Table_4G.Rows[k][13] = -1;
                            }
                            else if (Convert.ToDouble(Data_Table_4G.Rows[k][5]) < Convert.ToDouble(textBox22.Text))
                            {
                                Data_Table_4G.Rows[k][13] = 1; result++; dataGridView1.Rows[k + 1].Cells[6].Style.BackColor = Color.Orange;
                            }
                            else if (Convert.ToDouble(Data_Table_4G.Rows[k][5]) >= Convert.ToDouble(textBox22.Text))
                            {
                                Data_Table_4G.Rows[k][13] = 0;
                            }


                            // Intra Freq HO SR
                            //  dataGridView1.Rows[k + 1].Cells[7].Value = Data_Table_4G.Rows[k][3];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[7].Value = Data_Table_4G.Rows[k][6]));
                            if (Data_Table_4G.Rows[k][6].ToString() == "")
                            {
                                Data_Table_4G.Rows[k][14] = -1;
                            }
                            else if (Convert.ToDouble(Data_Table_4G.Rows[k][6]) < Convert.ToDouble(textBox23.Text))
                            {
                                Data_Table_4G.Rows[k][14] = 1; result++; dataGridView1.Rows[k + 1].Cells[7].Style.BackColor = Color.Orange;
                            }
                            else if (Convert.ToDouble(Data_Table_4G.Rows[k][6]) >= Convert.ToDouble(textBox23.Text))
                            {
                                Data_Table_4G.Rows[k][14] = 0;
                            }


                            // RRC Connection SR
                            //dataGridView1.Rows[k + 1].Cells[8].Value = Data_Table_4G.Rows[k][3];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[8].Value = Data_Table_4G.Rows[k][7]));
                            if (Data_Table_4G.Rows[k][7].ToString() == "")
                            {
                                Data_Table_4G.Rows[k][15] = -1;
                            }
                            else if (Convert.ToDouble(Data_Table_4G.Rows[k][7]) < Convert.ToDouble(textBox24.Text))
                            {
                                Data_Table_4G.Rows[k][15] = 1; result++; dataGridView1.Rows[k + 1].Cells[8].Style.BackColor = Color.Orange;
                            }
                            else if (Convert.ToDouble(Data_Table_4G.Rows[k][7]) >= Convert.ToDouble(textBox24.Text))
                            {
                                Data_Table_4G.Rows[k][15] = 0;
                            }


                            // Cell Availability
                            // dataGridView1.Rows[k + 1].Cells[9].Value = Data_Table_4G.Rows[k][3];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[9].Value = Data_Table_4G.Rows[k][8]));
                            if (Data_Table_4G.Rows[k][8].ToString() == "")
                            {
                                Data_Table_4G.Rows[k][16] = -1;
                            }
                            else if (Convert.ToDouble(Data_Table_4G.Rows[k][8]) < Convert.ToDouble(textBox25.Text) && Convert.ToDouble(Data_Table_4G.Rows[k][8]) > 0)
                            {
                                Data_Table_4G.Rows[k][16] = 1; result++; dataGridView1.Rows[k + 1].Cells[9].Style.BackColor = Color.Orange;
                            }
                            else if (Convert.ToDouble(Data_Table_4G.Rows[k][8]) >= Convert.ToDouble(textBox25.Text))
                            {
                                Data_Table_4G.Rows[k][16] = 0;
                            }


                            // Daily PS Traffic (GB)
                            // dataGridView1.Rows[k + 1].Cells[10].Value = Data_Table_4G.Rows[k][3];
                            dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[10].Value = Data_Table_4G.Rows[k][9]));
                            if (Data_Table_4G.Rows[k][9].ToString() == "")
                            {
                                Data_Table_4G.Rows[k][17] = -1;
                            }
                            else if (Convert.ToDouble(Data_Table_4G.Rows[k][9]) == Convert.ToDouble(textBox26.Text))
                            {
                                Data_Table_4G.Rows[k][17] = 1; result++; dataGridView1.Rows[k + 1].Cells[10].Style.BackColor = Color.Orange;
                            }
                            else if (Convert.ToDouble(Data_Table_4G.Rows[k][9]) > Convert.ToDouble(textBox26.Text))
                            {
                                Data_Table_4G.Rows[k][17] = 0;
                            }



                            if (Convert.ToInt16(Data_Table_4G.Rows[k][17]) == -1)
                            {
                                Data_Table_4G.Rows[k][18] = 1; dataGridView1.Rows[k + 1].Cells[11].Value = "Not Updated";
                            }
                            else if (result == 0)
                            {
                                Data_Table_4G.Rows[k][18] = 0.1; dataGridView1.Rows[k + 1].Cells[11].Value = "Passed";
                            }
                            if (result > 0)
                            {
                                Data_Table_4G.Rows[k][18] = 0; dataGridView1.Rows[k + 1].Cells[11].Value = "Rejected";
                            }




                            progressBar1.Invoke(new Action(() => progressBar1.Value = k));
                            //progressBar1.Value = k;
                        }


                        var distinctIds = Data_Table_4G.AsEnumerable()
        .Select(s => new
        {
            id = s.Field<string>("Site"),
        })
        .Distinct().ToList();

                        for (int j = 0; j < distinctIds.Count; j++)
                        {
                            var cell_data = (from p in Data_Table_4G.AsEnumerable()
                                             where p.Field<string>("Site") == distinctIds[j].id
                                             select p).ToList();


                            double multiplier = 1;
                            for (int h = 0; h < cell_data.Count; h++)
                            {
                                multiplier = multiplier * Convert.ToDouble(cell_data[h].ItemArray[18]);

                            }

                            if (multiplier > 0 && multiplier < 1)
                            {
                                Site_Data_Table_4G.Rows.Add(distinctIds[j].id, "Passed");
                            }
                            if (multiplier == 0)
                            {
                                Site_Data_Table_4G.Rows.Add(distinctIds[j].id, "Rejected");
                            }
                            if (multiplier == 1)
                            {
                                Site_Data_Table_4G.Rows.Add(distinctIds[j].id, "Not Updated");
                            }

                        }





                    }


                }


                label43.Invoke(new Action(() => label43.Text = "Finished"));
                label43.Invoke(new Action(() => label43.BackColor = Color.LightGreen));

            }



            //private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
            //{

            //}

            //private void textBox27_TextChanged(object sender, EventArgs e)
            //{

            //    //if (textBox27.Text=="performance-nak")
            //    //{
            //    //    this.Close();
            //    //}
            //}

            //private void checkBox1_CheckedChanged(object sender, EventArgs e)
            //{

            //    if (checkBox1.Checked == true)
            //    {
            //        Input_Type = "DataBase";
            //        checkBox2.Checked = false;
            //    }
            //}

            //private void checkBox2_CheckedChanged(object sender, EventArgs e)
            //{

            //    if (checkBox2.Checked == true)
            //    {
            //        Input_Type = "FARAZ";
            //        checkBox1.Checked = false;
            //    }
            //}







        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                Input_Type = "DataBase";
                checkBox2.Checked = false;
            }
            Technology = comboBox1.SelectedItem.ToString();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            openFileDialog1.DefaultExt = "xlsx";
            openFileDialog1.Filter = "Excel File|*.xlsx";
            DialogResult result1 = openFileDialog1.ShowDialog();
            string File_Name = openFileDialog1.SafeFileName.ToString();


            if (Input_Type == "FARAZ" && Technology == "2G")
            {


                //string Data_Quary = @"select [Date], [BSC], [Cell], [TCH_Traffic_BH] as 'TCH_Traffic_BH (Erlang)', [CSSR_MCI] as'CSSR', [CDR(not Affected by incoming Handovers from 3G)(Eric_CELL)] as 'Voice Drop Rate', [IHSR] as 'IHSR', [OHSR] as 'OHSR', [TCH_Availability] as 'TCH Availability' from [dbo].[CC2_Ericsson_Cell_BH] where  (" + EH_sites_list + ") and (" + date_list + ")" +
                //       @" union all select [Date], [BSC], [Cell], [TCH_Traffic_BH] as 'TCH_Traffic_BH (Erlang)', [CSSR3] as'CSSR', [CDR3] as 'Voice Drop Rate', [IHSR2] as 'IHSR', [OHSR2] as 'OHSR',  [TCH_Availability] as 'TCH Availability' from [dbo].[CC2_Huawei_Cell_BH] where (" + EH_sites_list + ") and (" + date_list + ")" +
                //       @" union all select [Date], [BSC], [Cell], [TCH_Traffic_BH] as 'TCH_Traffic_BH (Erlang)', [CSSR3] as'CSSR', [CDR3] as 'Voice Drop Rate', [IHSR2] as 'IHSR', [OHSR2] as 'OHSR',  [TCH_Availability] as 'TCH Availability' from [dbo].[CC2_Huawei_Cell_BH] where (" + H_sites_list + ") and (" + date_list + ")" +
                //       @" union all select [Date], [BSC], [SEG] as 'Cell', [TCH_Traffic_BH] as 'TCH_Traffic_BH (Erlang)', [CSSR_MCI] as'CSSR', [CDR(including_CS_IRAT_handovers_3G_to2G)(Nokia_SEG)] as 'Voicde  Drop Rate', [IHSR] as 'IHSR', [OHSR] AS 'OHSR', [TCH_Availability] as 'TCH Availability' from [dbo].[CC2_Nokia_Cell_BH] where (" + N_sites_list + ") and (" + date_list + ")";

                //SqlCommand Data_Quary1 = new SqlCommand(Data_Quary, connection);
                //Data_Quary1.CommandTimeout = 0;
                //Data_Quary1.ExecuteNonQuery();
                //Data_Table_2G = new DataTable();
                //SqlDataAdapter Date_Table1 = new SqlDataAdapter(Data_Quary1);
                //Date_Table1.Fill(Data_Table_2G);

                Data_Table_2G.Columns.Add("Date", typeof(DateTime));
                Data_Table_2G.Columns.Add("BSC", typeof(string));
                Data_Table_2G.Columns.Add("Cell", typeof(string));
                Data_Table_2G.Columns.Add("TCH_Traffic_BH (Erlang)", typeof(string));
                Data_Table_2G.Columns.Add("CSSR", typeof(string));
                Data_Table_2G.Columns.Add("Voice Drop Rate", typeof(string));
                Data_Table_2G.Columns.Add("IHSR", typeof(string));
                Data_Table_2G.Columns.Add("OHSR", typeof(string));
                Data_Table_2G.Columns.Add("TCH Availability", typeof(string));

                Data_Table_2G.Columns.Add("Traffic Score", typeof(int));
                Data_Table_2G.Columns.Add("CSSR Score", typeof(int));
                Data_Table_2G.Columns.Add("CDR Score", typeof(int));
                Data_Table_2G.Columns.Add("IHSR Score", typeof(int));
                Data_Table_2G.Columns.Add("OHSR Score", typeof(int));
                Data_Table_2G.Columns.Add("Availability Score", typeof(int));
                Data_Table_2G.Columns.Add("Cell Score", typeof(double));
                // Data_Table_2G.Columns.Add("Site Score", typeof(double));
                Data_Table_2G.Columns.Add("Site", typeof(string));


                if (result1 == DialogResult.OK)
                {
                    string file = openFileDialog1.FileName;
                    FName = file;
                    xlApp = new Excel.Application();
                    xlWorkBook = xlApp.Workbooks.Open(file);
                    Sheet = xlWorkBook.Worksheets[1];

                    Excel.Range History_TT = Sheet.get_Range("A2", "H" + Sheet.UsedRange.Rows.Count);
                    object[,] FARAZ_Data = (object[,])History_TT.Value;
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



                        string Traffic = "";
                        string CSSR = "";
                        string CDR = "";
                        string IHSR = "";
                        string OHSR = "";
                        string Availability = "";
                        if (FARAZ_Data[k + 1, 8] != null)
                        {
                            Traffic = FARAZ_Data[k + 1, 8].ToString();
                        }
                        else
                        {
                            Traffic = "";
                        }

                        if (FARAZ_Data[k + 1, 4] != null)
                        {
                            CSSR = FARAZ_Data[k + 1, 4].ToString();
                        }
                        else
                        {
                            CSSR = "";
                        }

                        if (FARAZ_Data[k + 1, 3] != null)
                        {
                            CDR = FARAZ_Data[k + 1, 3].ToString();
                        }
                        else
                        {
                            CDR = "";
                        }

                        if (FARAZ_Data[k + 1, 5] != null)
                        {
                            IHSR = FARAZ_Data[k + 1, 5].ToString();
                        }
                        else
                        {
                            IHSR = "";
                        }

                        if (FARAZ_Data[k + 1, 6] != null)
                        {
                            OHSR = FARAZ_Data[k + 1, 6].ToString();
                        }
                        else
                        {
                            OHSR = "";
                        }

                        if (FARAZ_Data[k + 1, 7] != null)
                        {
                            Availability = FARAZ_Data[k + 1, 7].ToString();
                        }
                        else
                        {
                            Availability = "";
                        }


                        Data_Table_2G.Rows.Add(Date, BSC, CellName, Traffic, CSSR, CDR, IHSR, OHSR, Availability, -1, -1, -1, -1, -1, -1, -1, Site);

                    }

                }





                Site_Data_Table_2G = new DataTable();
                Site_Data_Table_2G.Columns.Add("Site", typeof(string));
                Site_Data_Table_2G.Columns.Add("KPI Zero Status", typeof(string));
                Site_Data_Table_2G.Columns.Add("Rejected Cell List", typeof(string));


                dataGridView1.ColumnCount = 11;

                dataGridView1.Rows.Clear();
                dataGridView1.RowCount = Data_Table_2G.Rows.Count + 1;
                dataGridView1.Rows[0].Cells[0].Value = "Date"; dataGridView1.Columns[0].Width = 100;
                dataGridView1.Rows[0].Cells[1].Value = "BSC"; dataGridView1.Columns[1].Width = 100;
                dataGridView1.Rows[0].Cells[2].Value = "Site"; dataGridView1.Columns[2].Width = 100;
                dataGridView1.Rows[0].Cells[3].Value = "Cell"; dataGridView1.Columns[3].Width = 100;
                dataGridView1.Rows[0].Cells[4].Value = "TCH_Traffic_BH (Erlang)"; dataGridView1.Columns[4].Width = 100;
                dataGridView1.Rows[0].Cells[5].Value = "CSSR"; dataGridView1.Columns[5].Width = 100;
                dataGridView1.Rows[0].Cells[6].Value = "CDR"; dataGridView1.Columns[6].Width = 100;
                dataGridView1.Rows[0].Cells[7].Value = "IHSR"; dataGridView1.Columns[7].Width = 100;
                dataGridView1.Rows[0].Cells[8].Value = "OHSR"; dataGridView1.Columns[8].Width = 100;
                dataGridView1.Rows[0].Cells[9].Value = "Availability"; dataGridView1.Columns[9].Width = 100;
                dataGridView1.Rows[0].Cells[10].Value = "Cell Status"; dataGridView1.Columns[10].Width = 100;
                // dataGridView1.Rows[0].Cells[11].Value = "Site Status"; dataGridView1.Columns[11].Width = 100;

                progressBar1.Minimum = 0;



                if (Data_Table_2G.Rows.Count == 0)
                {
                    MessageBox.Show("There is no Data in Database");
                }


                if (Data_Table_2G.Rows.Count != 0)
                {
                    progressBar1.Maximum = Data_Table_2G.Rows.Count - 1;
                    for (int k = 0; k < Data_Table_2G.Rows.Count; k++)
                    {

                        int result = 0;

                        // Date
                        dataGridView1.Rows[k + 1].Cells[0].Value = Data_Table_2G.Rows[k][0];
                        //BSC
                        dataGridView1.Rows[k + 1].Cells[1].Value = Data_Table_2G.Rows[k][1];
                        // Site
                        string Cell = Data_Table_2G.Rows[k][2].ToString();
                        string Site = Data_Table_2G.Rows[k][16].ToString();
                        //if (Cell.Length == 7)
                        //{
                        //    Site = Cell.Substring(0, 6);
                        //}
                        //if (Cell.Length > 7)
                        //{
                        //    Site = Cell.Substring(0, 2) + Cell.Substring(4, 4);
                        //}
                        dataGridView1.Rows[k + 1].Cells[2].Value = Site;
                        Data_Table_2G.Rows[k][16] = Site;

                        // Cell
                        dataGridView1.Rows[k + 1].Cells[3].Value = Data_Table_2G.Rows[k][2];

                        // Traffic

                        dataGridView1.Rows[k + 1].Cells[4].Value = Data_Table_2G.Rows[k][3];
                        if (Data_Table_2G.Rows[k][3].ToString() == "")
                        {
                            Data_Table_2G.Rows[k][9] = -1;
                        }
                        else if (Convert.ToDouble(Data_Table_2G.Rows[k][3]) == Convert.ToDouble(textBox7.Text))
                        {
                            Data_Table_2G.Rows[k][9] = 1; result++; dataGridView1.Rows[k + 1].Cells[4].Style.BackColor = Color.Orange;
                        }
                        else if (Convert.ToDouble(Data_Table_2G.Rows[k][3]) > Convert.ToDouble(textBox7.Text))
                        {
                            Data_Table_2G.Rows[k][9] = 0;
                        }

                        // CSSR
                        dataGridView1.Rows[k + 1].Cells[5].Value = Data_Table_2G.Rows[k][4];
                        if (Data_Table_2G.Rows[k][4].ToString() == "")
                        {
                            Data_Table_2G.Rows[k][10] = -1;
                        }
                        else if (Convert.ToDouble(Data_Table_2G.Rows[k][4]) < Convert.ToDouble(textBox3.Text))
                        {
                            Data_Table_2G.Rows[k][10] = 1; result++; dataGridView1.Rows[k + 1].Cells[5].Style.BackColor = Color.Orange;
                        }
                        else if (Convert.ToDouble(Data_Table_2G.Rows[k][4]) >= Convert.ToDouble(textBox3.Text))
                        {
                            Data_Table_2G.Rows[k][10] = 0;
                        }

                        // CDR
                        dataGridView1.Rows[k + 1].Cells[6].Value = Data_Table_2G.Rows[k][5];
                        if (Data_Table_2G.Rows[k][5].ToString() == "")
                        {
                            Data_Table_2G.Rows[k][11] = -1;
                        }
                        else if (Convert.ToDouble(Data_Table_2G.Rows[k][5]) > Convert.ToDouble(textBox2.Text))
                        {
                            Data_Table_2G.Rows[k][11] = 1; result++; dataGridView1.Rows[k + 1].Cells[6].Style.BackColor = Color.Orange;
                        }
                        else if (Convert.ToDouble(Data_Table_2G.Rows[k][5]) <= Convert.ToDouble(textBox2.Text))
                        {
                            Data_Table_2G.Rows[k][11] = 0;
                        }

                        // IHSR
                        dataGridView1.Rows[k + 1].Cells[7].Value = Data_Table_2G.Rows[k][6];
                        if (Data_Table_2G.Rows[k][6].ToString() == "")
                        {
                            Data_Table_2G.Rows[k][12] = -1;
                        }
                        else if (Convert.ToDouble(Data_Table_2G.Rows[k][6]) < Convert.ToDouble(textBox4.Text))
                        {
                            Data_Table_2G.Rows[k][12] = 1; result++; dataGridView1.Rows[k + 1].Cells[7].Style.BackColor = Color.Orange;
                        }
                        else if (Convert.ToDouble(Data_Table_2G.Rows[k][6]) >= Convert.ToDouble(textBox4.Text))
                        {
                            Data_Table_2G.Rows[k][12] = 0;
                        }

                        // OHSR
                        dataGridView1.Rows[k + 1].Cells[8].Value = Data_Table_2G.Rows[k][7];
                        if (Data_Table_2G.Rows[k][7].ToString() == "")
                        {
                            Data_Table_2G.Rows[k][13] = -1;
                        }
                        else if (Convert.ToDouble(Data_Table_2G.Rows[k][7]) < Convert.ToDouble(textBox5.Text))
                        {
                            Data_Table_2G.Rows[k][13] = 1; result++; dataGridView1.Rows[k + 1].Cells[8].Style.BackColor = Color.Orange;
                        }
                        else if (Convert.ToDouble(Data_Table_2G.Rows[k][7]) >= Convert.ToDouble(textBox5.Text))
                        {
                            Data_Table_2G.Rows[k][13] = 0;
                        }

                        // Availability
                        dataGridView1.Rows[k + 1].Cells[9].Value = Data_Table_2G.Rows[k][8];
                        if (Data_Table_2G.Rows[k][8].ToString() == "")
                        {
                            Data_Table_2G.Rows[k][14] = -1;
                        }
                        else if (Convert.ToDouble(Data_Table_2G.Rows[k][8]) < Convert.ToDouble(textBox6.Text) && Convert.ToDouble(Data_Table_2G.Rows[k][8]) > 0)
                        {
                            Data_Table_2G.Rows[k][14] = 1; result++; dataGridView1.Rows[k + 1].Cells[9].Style.BackColor = Color.Orange;
                        }
                        else if (Convert.ToDouble(Data_Table_2G.Rows[k][8]) >= Convert.ToDouble(textBox6.Text))
                        {
                            Data_Table_2G.Rows[k][14] = 0;
                        }


                        if (Convert.ToInt16(Data_Table_2G.Rows[k][9]) == -1)
                        {
                            Data_Table_2G.Rows[k][15] = 1; dataGridView1.Rows[k + 1].Cells[10].Value = "Not Updated";
                        }
                        else if (result == 0)
                        {
                            Data_Table_2G.Rows[k][15] = 0.1; dataGridView1.Rows[k + 1].Cells[10].Value = "Passed";
                        }
                        if (result > 0)
                        {
                            Data_Table_2G.Rows[k][15] = 0; dataGridView1.Rows[k + 1].Cells[10].Value = "Rejected";
                        }

                        progressBar1.Value = k;

                    }


                    var distinctIds = Data_Table_2G.AsEnumerable()
                       .Select(s => new
                       {
                           id = s.Field<string>("Site"),
                       })
                       .Distinct().ToList();

                    for (int j = 0; j < distinctIds.Count; j++)
                    {
                        var cell_data = (from p in Data_Table_2G.AsEnumerable()
                                         where p.Field<string>("Site") == distinctIds[j].id
                                         select p).ToList();


                        double multiplier = 1;
                        for (int h = 0; h < cell_data.Count; h++)
                        {
                            multiplier = multiplier * Convert.ToDouble(cell_data[h].ItemArray[15]);

                        }

                        if (multiplier > 0 && multiplier < 1)
                        {
                            Site_Data_Table_2G.Rows.Add(distinctIds[j].id, "Passed");
                        }
                        if (multiplier == 0)
                        {
                            Site_Data_Table_2G.Rows.Add(distinctIds[j].id, "Rejected");
                        }
                        if (multiplier == 1)
                        {
                            Site_Data_Table_2G.Rows.Add(distinctIds[j].id, "Not Updated");
                        }

                    }
                }


            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            openFileDialog1.DefaultExt = "xlsx";
            openFileDialog1.Filter = "Excel File|*.xlsx";
            DialogResult result1 = openFileDialog1.ShowDialog();
            string File_Name = openFileDialog1.SafeFileName.ToString();
            if (Input_Type == "FARAZ" && Technology == "2G")
            {
                Technology = "2G-MCI";

                //string Data_Quary = @"select [Date], [BSC], [Cell], [TCH_Traffic_BH] as 'TCH_Traffic_BH (Erlang)', [CSSR_MCI] as'CSSR', [CDR(not Affected by incoming Handovers from 3G)(Eric_CELL)] as 'Voice Drop Rate', [IHSR] as 'IHSR', [OHSR] as 'OHSR', [TCH_Availability] as 'TCH Availability' from [dbo].[CC2_Ericsson_Cell_BH] where  (" + EH_sites_list + ") and (" + date_list + ")" +
                //       @" union all select [Date], [BSC], [Cell], [TCH_Traffic_BH] as 'TCH_Traffic_BH (Erlang)', [CSSR3] as'CSSR', [CDR3] as 'Voice Drop Rate', [IHSR2] as 'IHSR', [OHSR2] as 'OHSR',  [TCH_Availability] as 'TCH Availability' from [dbo].[CC2_Huawei_Cell_BH] where (" + EH_sites_list + ") and (" + date_list + ")" +
                //       @" union all select [Date], [BSC], [Cell], [TCH_Traffic_BH] as 'TCH_Traffic_BH (Erlang)', [CSSR3] as'CSSR', [CDR3] as 'Voice Drop Rate', [IHSR2] as 'IHSR', [OHSR2] as 'OHSR',  [TCH_Availability] as 'TCH Availability' from [dbo].[CC2_Huawei_Cell_BH] where (" + H_sites_list + ") and (" + date_list + ")" +
                //       @" union all select [Date], [BSC], [SEG] as 'Cell', [TCH_Traffic_BH] as 'TCH_Traffic_BH (Erlang)', [CSSR_MCI] as'CSSR', [CDR(including_CS_IRAT_handovers_3G_to2G)(Nokia_SEG)] as 'Voicde  Drop Rate', [IHSR] as 'IHSR', [OHSR] AS 'OHSR', [TCH_Availability] as 'TCH Availability' from [dbo].[CC2_Nokia_Cell_BH] where (" + N_sites_list + ") and (" + date_list + ")";

                //SqlCommand Data_Quary1 = new SqlCommand(Data_Quary, connection);
                //Data_Quary1.CommandTimeout = 0;
                //Data_Quary1.ExecuteNonQuery();
                //Data_Table_2G = new DataTable();
                //SqlDataAdapter Date_Table1 = new SqlDataAdapter(Data_Quary1);
                //Date_Table1.Fill(Data_Table_2G);

                Data_Table_2G.Columns.Add("CS Datetime", typeof(DateTime));
                // Data_Table_2G.Columns.Add("PS Datetime", typeof(DateTime));
                Data_Table_2G.Columns.Add("BSC", typeof(string));
                Data_Table_2G.Columns.Add("Cell", typeof(string));
                Data_Table_2G.Columns.Add("TCH_Traffic_BH (Erlang), TH=0", typeof(string));
                Data_Table_2G.Columns.Add("CSSR, TH=92", typeof(string));
                Data_Table_2G.Columns.Add("Voice Drop Rate, TH=4", typeof(string));
                Data_Table_2G.Columns.Add("IHSR, TH=92", typeof(string));
                Data_Table_2G.Columns.Add("OHSR, TH=92", typeof(string));
                Data_Table_2G.Columns.Add("TCH Availability, TH=99", typeof(string));
                Data_Table_2G.Columns.Add("TBF Establishmnet SR, TH=80", typeof(string));

                Data_Table_2G.Columns.Add("Traffic Score", typeof(int));
                Data_Table_2G.Columns.Add("CSSR Score", typeof(int));
                Data_Table_2G.Columns.Add("CDR Score", typeof(int));
                Data_Table_2G.Columns.Add("IHSR Score", typeof(int));
                Data_Table_2G.Columns.Add("OHSR Score", typeof(int));
                Data_Table_2G.Columns.Add("Availability Score", typeof(int));
                Data_Table_2G.Columns.Add("TBF Score", typeof(int));
                Data_Table_2G.Columns.Add("Cell Score", typeof(double));
                // Data_Table_2G.Columns.Add("Site Score", typeof(double));
                Data_Table_2G.Columns.Add("Site", typeof(string));


                string Vendor = "";

                if (result1 == DialogResult.OK)
                {
                    string file = openFileDialog1.FileName;
                    FName = file;
                    xlApp = new Excel.Application();
                    xlWorkBook = xlApp.Workbooks.Open(file);
                    Sheet = xlWorkBook.Worksheets[1];




                    Excel.Range History_TT = Sheet.get_Range("A2", "I" + Sheet.UsedRange.Rows.Count);
                    object[,] FARAZ_Data = (object[,])History_TT.Value;
                    int Count = Sheet.UsedRange.Rows.Count;

                    for (int k = 0; k < Count - 1; k++)
                    {

                        if (FARAZ_Data[k + 1, 8] == null && FARAZ_Data[k + 1, 9] == null)
                        {
                            continue;
                        }

                        DateTime Date = Convert.ToDateTime(FARAZ_Data[k + 1, 1]);
                        string NE = FARAZ_Data[k + 1, 2].ToString();
                        string Site = "";


                        // Filling TBF Establishment SR
                        for (int k1 = 0; k1 < Count - 1; k1++)
                        {
                            DateTime Date1 = Convert.ToDateTime(FARAZ_Data[k1 + 1, 1]);
                            string NE1 = FARAZ_Data[k1 + 1, 2].ToString();
                            if (Date.Date == Date1.Date && NE == NE1 && FARAZ_Data[k1 + 1, 7] != null)
                            {
                                FARAZ_Data[k + 1, 7] = FARAZ_Data[k1 + 1, 7];
                                break;
                            }
                        }






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


                        if (Tech_Last == "E")
                        {
                            Vendor = "Ericsson";
                        }
                        if (Tech_Last == "H")
                        {
                            Vendor = "Huawei";
                        }
                        if (Tech_Last == "N")
                        {
                            Vendor = "Nokia";
                        }

                        string Traffic = "";
                        string CSSR = "";
                        string CDR = "";
                        string IHSR = "";
                        string OHSR = "";
                        string Availability = "";
                        string TBF = "";
                        if (FARAZ_Data[k + 1, 9] != null)
                        {
                            Traffic = FARAZ_Data[k + 1, 9].ToString();
                        }
                        else
                        {
                            Traffic = "";
                        }

                        if (FARAZ_Data[k + 1, 4] != null)
                        {
                            CSSR = FARAZ_Data[k + 1, 4].ToString();
                        }
                        else
                        {
                            CSSR = "";
                        }

                        if (FARAZ_Data[k + 1, 3] != null)
                        {
                            CDR = FARAZ_Data[k + 1, 3].ToString();
                        }
                        else
                        {
                            CDR = "";
                        }

                        if (FARAZ_Data[k + 1, 5] != null)
                        {
                            IHSR = FARAZ_Data[k + 1, 5].ToString();
                        }
                        else
                        {
                            IHSR = "";
                        }

                        if (FARAZ_Data[k + 1, 6] != null)
                        {
                            OHSR = FARAZ_Data[k + 1, 6].ToString();
                        }
                        else
                        {
                            OHSR = "";
                        }

                        if (FARAZ_Data[k + 1, 7] != null)
                        {
                            TBF = FARAZ_Data[k + 1, 7].ToString();
                        }
                        else
                        {
                            TBF = "";
                        }

                        if (FARAZ_Data[k + 1, 8] != null)
                        {
                            Availability = FARAZ_Data[k + 1, 8].ToString();
                        }
                        else
                        {
                            Availability = "";
                        }


                        Data_Table_2G.Rows.Add(Date, BSC, CellName, Traffic, CSSR, CDR, IHSR, OHSR, Availability, TBF, -1, -1, -1, -1, -1, -1, -1, -1, Site);

                    }

                }





                Site_Data_Table_2G = new DataTable();
                Site_Data_Table_2G.Columns.Add("Site", typeof(string));
                Site_Data_Table_2G.Columns.Add("KPI Zero Status", typeof(string));
                Site_Data_Table_2G.Columns.Add("Rejected Cell List", typeof(string));


                dataGridView1.ColumnCount = 12;

                dataGridView1.Rows.Clear();
                dataGridView1.RowCount = Data_Table_2G.Rows.Count + 1;
                dataGridView1.Rows[0].Cells[0].Value = "Date"; dataGridView1.Columns[0].Width = 100;
                dataGridView1.Rows[0].Cells[1].Value = "BSC"; dataGridView1.Columns[1].Width = 100;
                dataGridView1.Rows[0].Cells[2].Value = "Site"; dataGridView1.Columns[2].Width = 100;
                dataGridView1.Rows[0].Cells[3].Value = "Cell"; dataGridView1.Columns[3].Width = 100;
                dataGridView1.Rows[0].Cells[4].Value = "TCH_Traffic_BH (Erlang)"; dataGridView1.Columns[4].Width = 100;
                dataGridView1.Rows[0].Cells[5].Value = "CSSR"; dataGridView1.Columns[5].Width = 100;
                dataGridView1.Rows[0].Cells[6].Value = "CDR"; dataGridView1.Columns[6].Width = 100;
                dataGridView1.Rows[0].Cells[7].Value = "IHSR"; dataGridView1.Columns[7].Width = 100;
                dataGridView1.Rows[0].Cells[8].Value = "OHSR"; dataGridView1.Columns[8].Width = 100;
                dataGridView1.Rows[0].Cells[9].Value = "Availability"; dataGridView1.Columns[9].Width = 100;
                dataGridView1.Rows[0].Cells[10].Value = "TBF Estblishment SR"; dataGridView1.Columns[10].Width = 100;
                dataGridView1.Rows[0].Cells[11].Value = "Cell Status"; dataGridView1.Columns[11].Width = 100;
                // dataGridView1.Rows[0].Cells[11].Value = "Site Status"; dataGridView1.Columns[11].Width = 100;

                progressBar1.Minimum = 0;



                if (Data_Table_2G.Rows.Count == 0)
                {
                    MessageBox.Show("There is no Data in Database");
                }


                if (Data_Table_2G.Rows.Count != 0)
                {



                    double Traffic_TH = 0;
                    double CSSR_TH = 92;
                    double CDR_TH = 4;
                    double IHSR_TH = 92;
                    double OHSR_TH = 92;
                    double Availability_TH = 99;
                    double TBF_TH = 80;




                    progressBar1.Maximum = Data_Table_2G.Rows.Count - 1;
                    for (int k = 0; k < Data_Table_2G.Rows.Count; k++)
                    {

                        int result = 0;

                        // Date
                        dataGridView1.Rows[k + 1].Cells[0].Value = Data_Table_2G.Rows[k][0];
                        //BSC
                        dataGridView1.Rows[k + 1].Cells[1].Value = Data_Table_2G.Rows[k][1];
                        // Site
                        string Cell = Data_Table_2G.Rows[k][2].ToString();
                        string Site = Data_Table_2G.Rows[k][18].ToString();
                        //if (Cell.Length == 7)
                        //{
                        //    Site = Cell.Substring(0, 6);
                        //}
                        //if (Cell.Length > 7)
                        //{
                        //    Site = Cell.Substring(0, 2) + Cell.Substring(4, 4);
                        //}
                        dataGridView1.Rows[k + 1].Cells[2].Value = Site;
                        Data_Table_2G.Rows[k][18] = Site;

                        // Cell
                        dataGridView1.Rows[k + 1].Cells[3].Value = Data_Table_2G.Rows[k][2];

                        // Traffic

                        dataGridView1.Rows[k + 1].Cells[4].Value = Data_Table_2G.Rows[k][3];
                        if (Data_Table_2G.Rows[k][3].ToString() == "")
                        {
                            Data_Table_2G.Rows[k][10] = -1;
                        }
                        else if (Convert.ToDouble(Data_Table_2G.Rows[k][3]) == Traffic_TH)
                        {
                            Data_Table_2G.Rows[k][10] = 1; result++; dataGridView1.Rows[k + 1].Cells[4].Style.BackColor = Color.Orange;
                        }
                        else if (Convert.ToDouble(Data_Table_2G.Rows[k][3]) > Traffic_TH)
                        {
                            Data_Table_2G.Rows[k][10] = 0;
                        }

                        // CSSR
                        dataGridView1.Rows[k + 1].Cells[5].Value = Data_Table_2G.Rows[k][4];
                        if (Data_Table_2G.Rows[k][4].ToString() == "")
                        {
                            Data_Table_2G.Rows[k][11] = -1;
                        }
                        else if (Convert.ToDouble(Data_Table_2G.Rows[k][4]) < CSSR_TH)
                        {
                            Data_Table_2G.Rows[k][11] = 1; result++; dataGridView1.Rows[k + 1].Cells[5].Style.BackColor = Color.Orange;
                        }
                        else if (Convert.ToDouble(Data_Table_2G.Rows[k][4]) >= CSSR_TH)
                        {
                            Data_Table_2G.Rows[k][11] = 0;
                        }

                        // CDR
                        dataGridView1.Rows[k + 1].Cells[6].Value = Data_Table_2G.Rows[k][5];
                        if (Data_Table_2G.Rows[k][5].ToString() == "")
                        {
                            Data_Table_2G.Rows[k][12] = -1;
                        }
                        else if (Convert.ToDouble(Data_Table_2G.Rows[k][5]) > CDR_TH)
                        {
                            Data_Table_2G.Rows[k][12] = 1; result++; dataGridView1.Rows[k + 1].Cells[6].Style.BackColor = Color.Orange;
                        }
                        else if (Convert.ToDouble(Data_Table_2G.Rows[k][5]) <= CDR_TH)
                        {
                            Data_Table_2G.Rows[k][12] = 0;
                        }

                        // IHSR
                        dataGridView1.Rows[k + 1].Cells[7].Value = Data_Table_2G.Rows[k][6];
                        if (Data_Table_2G.Rows[k][6].ToString() == "")
                        {
                            Data_Table_2G.Rows[k][13] = -1;
                        }
                        else if (Convert.ToDouble(Data_Table_2G.Rows[k][6]) < IHSR_TH)
                        {
                            Data_Table_2G.Rows[k][13] = 1; result++; dataGridView1.Rows[k + 1].Cells[7].Style.BackColor = Color.Orange;
                        }
                        else if (Convert.ToDouble(Data_Table_2G.Rows[k][6]) >= IHSR_TH)
                        {
                            Data_Table_2G.Rows[k][13] = 0;
                        }

                        // OHSR
                        dataGridView1.Rows[k + 1].Cells[8].Value = Data_Table_2G.Rows[k][7];
                        if (Data_Table_2G.Rows[k][7].ToString() == "")
                        {
                            Data_Table_2G.Rows[k][14] = -1;
                        }
                        else if (Convert.ToDouble(Data_Table_2G.Rows[k][7]) < OHSR_TH)
                        {
                            Data_Table_2G.Rows[k][14] = 1; result++; dataGridView1.Rows[k + 1].Cells[8].Style.BackColor = Color.Orange;
                        }
                        else if (Convert.ToDouble(Data_Table_2G.Rows[k][7]) >= OHSR_TH)
                        {
                            Data_Table_2G.Rows[k][14] = 0;
                        }

                        // Availability
                        dataGridView1.Rows[k + 1].Cells[9].Value = Data_Table_2G.Rows[k][8];
                        if (Data_Table_2G.Rows[k][8].ToString() == "")
                        {
                            Data_Table_2G.Rows[k][15] = -1;
                        }
                        else if (Convert.ToDouble(Data_Table_2G.Rows[k][8]) < Availability_TH && Convert.ToDouble(Data_Table_2G.Rows[k][8]) > 0)
                        {
                            Data_Table_2G.Rows[k][15] = 1; result++; dataGridView1.Rows[k + 1].Cells[9].Style.BackColor = Color.Orange;
                        }
                        else if (Convert.ToDouble(Data_Table_2G.Rows[k][8]) >= Availability_TH)
                        {
                            Data_Table_2G.Rows[k][15] = 0;
                        }



                        // TBF
                        dataGridView1.Rows[k + 1].Cells[10].Value = Data_Table_2G.Rows[k][9];
                        if (Data_Table_2G.Rows[k][9].ToString() == "")
                        {
                            Data_Table_2G.Rows[k][16] = -1;
                        }
                        else if (Convert.ToDouble(Data_Table_2G.Rows[k][9]) < TBF_TH)
                        {
                            Data_Table_2G.Rows[k][16] = 1; result++; dataGridView1.Rows[k + 1].Cells[10].Style.BackColor = Color.Orange;
                        }
                        else if (Convert.ToDouble(Data_Table_2G.Rows[k][9]) >= TBF_TH)
                        {
                            Data_Table_2G.Rows[k][16] = 0;
                        }




                        // Cell Score
                        if (Convert.ToInt16(Data_Table_2G.Rows[k][10]) == -1)
                        {
                            Data_Table_2G.Rows[k][17] = 1; dataGridView1.Rows[k + 1].Cells[11].Value = "Not Updated";
                        }
                        else if (result == 0)
                        {
                            Data_Table_2G.Rows[k][17] = 0.1; dataGridView1.Rows[k + 1].Cells[11].Value = "Passed";
                        }
                        if (result > 0)
                        {
                            Data_Table_2G.Rows[k][17] = 0; dataGridView1.Rows[k + 1].Cells[11].Value = "Rejected";
                        }

                        progressBar1.Value = k;

                    }


                    var distinctIds = Data_Table_2G.AsEnumerable()
                       .Select(s => new
                       {
                           id = s.Field<string>("Site"),
                       })
                       .Distinct().ToList();

                    for (int j = 0; j < distinctIds.Count; j++)
                    {
                        var cell_data = (from p in Data_Table_2G.AsEnumerable()
                                         where p.Field<string>("Site") == distinctIds[j].id
                                         select p).ToList();


                        double multiplier = 1;
                        for (int h = 0; h < cell_data.Count; h++)
                        {
                            multiplier = multiplier * Convert.ToDouble(cell_data[h].ItemArray[17]);

                        }

                        if (multiplier > 0 && multiplier < 1)
                        {
                            Site_Data_Table_2G.Rows.Add(distinctIds[j].id, "Passed");
                        }
                        if (multiplier == 0)
                        {
                            Site_Data_Table_2G.Rows.Add(distinctIds[j].id, "Rejected");
                        }
                        if (multiplier == 1)
                        {
                            Site_Data_Table_2G.Rows.Add(distinctIds[j].id, "Not Updated");
                        }

                    }
                }

            }



            if (Input_Type == "FARAZ" && Technology == "4G")
            {
                Technology = "4G-MCI";
                //string Data_Quary = @"select [Date], [BSC], [Cell], [TCH_Traffic_BH] as 'TCH_Traffic_BH (Erlang)', [CSSR_MCI] as'CSSR', [CDR(not Affected by incoming Handovers from 3G)(Eric_CELL)] as 'Voice Drop Rate', [IHSR] as 'IHSR', [OHSR] as 'OHSR', [TCH_Availability] as 'TCH Availability' from [dbo].[CC2_Ericsson_Cell_BH] where  (" + EH_sites_list + ") and (" + date_list + ")" +
                //       @" union all select [Date], [BSC], [Cell], [TCH_Traffic_BH] as 'TCH_Traffic_BH (Erlang)', [CSSR3] as'CSSR', [CDR3] as 'Voice Drop Rate', [IHSR2] as 'IHSR', [OHSR2] as 'OHSR',  [TCH_Availability] as 'TCH Availability' from [dbo].[CC2_Huawei_Cell_BH] where (" + EH_sites_list + ") and (" + date_list + ")" +
                //       @" union all select [Date], [BSC], [Cell], [TCH_Traffic_BH] as 'TCH_Traffic_BH (Erlang)', [CSSR3] as'CSSR', [CDR3] as 'Voice Drop Rate', [IHSR2] as 'IHSR', [OHSR2] as 'OHSR',  [TCH_Availability] as 'TCH Availability' from [dbo].[CC2_Huawei_Cell_BH] where (" + H_sites_list + ") and (" + date_list + ")" +
                //       @" union all select [Date], [BSC], [SEG] as 'Cell', [TCH_Traffic_BH] as 'TCH_Traffic_BH (Erlang)', [CSSR_MCI] as'CSSR', [CDR(including_CS_IRAT_handovers_3G_to2G)(Nokia_SEG)] as 'Voicde  Drop Rate', [IHSR] as 'IHSR', [OHSR] AS 'OHSR', [TCH_Availability] as 'TCH Availability' from [dbo].[CC2_Nokia_Cell_BH] where (" + N_sites_list + ") and (" + date_list + ")";

                //SqlCommand Data_Quary1 = new SqlCommand(Data_Quary, connection);
                //Data_Quary1.CommandTimeout = 0;
                //Data_Quary1.ExecuteNonQuery();
                //Data_Table_2G = new DataTable();
                //SqlDataAdapter Date_Table1 = new SqlDataAdapter(Data_Quary1);
                //Date_Table1.Fill(Data_Table_2G);

                DataTable Data_Table_4G1 = new DataTable();

                Data_Table_4G1.Columns.Add("Date", typeof(DateTime));
                Data_Table_4G1.Columns.Add("Site", typeof(string));
                Data_Table_4G1.Columns.Add("Sector", typeof(string));
                Data_Table_4G1.Columns.Add("Cell", typeof(string));
                Data_Table_4G1.Columns.Add("ERAB Setup SR, TH=96", typeof(string));
                Data_Table_4G1.Columns.Add("UE DL THR (Mbps), TH=4", typeof(string));
                Data_Table_4G1.Columns.Add("Cell Availability, TH=99", typeof(string));
                Data_Table_4G1.Columns.Add("CSFB_Success_Rate, TH=95", typeof(string));
                Data_Table_4G1.Columns.Add("Volte_Traffic, TH=0", typeof(string));
                Data_Table_4G1.Columns.Add("ERAB Drop Rate, TH=3", typeof(string));
                Data_Table_4G1.Columns.Add("Intra Freq HO SR, TH=95", typeof(string));
                Data_Table_4G1.Columns.Add("RRC Connection SR, TH=96", typeof(string));
                Data_Table_4G1.Columns.Add("PS_Traffic_Daily (GB), TH=0", typeof(string));


                Data_Table_4G.Columns.Add("Date", typeof(DateTime));
                Data_Table_4G.Columns.Add("Cell", typeof(string));
                Data_Table_4G.Columns.Add("ERAB Setup SR, TH=96", typeof(string));
                Data_Table_4G.Columns.Add("UE DL THR (Mbps), TH=4", typeof(string));
                Data_Table_4G.Columns.Add("Cell Availability, TH=99", typeof(string));
                Data_Table_4G.Columns.Add("CSFB_Success_Rate, TH=95", typeof(string));
                Data_Table_4G.Columns.Add("Volte_Traffic, TH=0", typeof(string));
                Data_Table_4G.Columns.Add("ERAB Drop Rate, TH=3", typeof(string));
                Data_Table_4G.Columns.Add("Intra Freq HO SR, TH=95", typeof(string));
                Data_Table_4G.Columns.Add("RRC Connection SR, TH=96", typeof(string));
                Data_Table_4G.Columns.Add("PS_Traffic_Daily (GB), TH=0", typeof(string));

                Data_Table_4G.Columns.Add("ERAB Setup SR Score", typeof(int));
                Data_Table_4G.Columns.Add("UE DL THR (Mbps) Score", typeof(int));
                Data_Table_4G.Columns.Add("Cell Availability Score", typeof(int));
                Data_Table_4G.Columns.Add("CSFB_Success_Rate Score", typeof(int));
                Data_Table_4G.Columns.Add("Volte_Traffic Score", typeof(int));
                Data_Table_4G.Columns.Add("ERAB Drop Rate Score", typeof(int));
                Data_Table_4G.Columns.Add("Intra Freq HO SR Score", typeof(int));
                Data_Table_4G.Columns.Add("RRC_Connection_Setup_Success Score", typeof(int));
                Data_Table_4G.Columns.Add("PS_Traffic_Daily (GB) Score", typeof(int));

                Data_Table_4G.Columns.Add("Cell Score", typeof(double));
                // Data_Table_4G.Columns.Add("Site Score", typeof(double));
                Data_Table_4G.Columns.Add("Site", typeof(string));



                string Vendor = "";

                if (result1 == DialogResult.OK)
                {
                    string file = openFileDialog1.FileName;
                    FName = file;
                    xlApp = new Excel.Application();
                    xlWorkBook = xlApp.Workbooks.Open(file);
                    Sheet = xlWorkBook.Worksheets[1];




                    Excel.Range History_TT = Sheet.get_Range("A2", "K" + Sheet.UsedRange.Rows.Count);
                    object[,] FARAZ_Data = (object[,])History_TT.Value;
                    int Count = Sheet.UsedRange.Rows.Count;


                    Excel.Range History_TT1 = Sheet.get_Range("A1", "K" + Sheet.UsedRange.Rows.Count);
                    object[,] FARAZ_Data1 = (object[,])History_TT1.Value;


                    if (FARAZ_Data1[1, 3].ToString() == "cell_availability_exclude_manual_blocking(Nokia_LTE_CELL)")
                    {
                        Vendor = "Nokia";
                    }

                    if (FARAZ_Data1[1, 3].ToString() == "Average_Downlink_User_Throughput(Mbit/s)(Hu_Cell)")
                    {
                        Vendor = "Huawei";
                    }

                    if (FARAZ_Data1[1, 3].ToString() == "Average_UE_DL_Throughput(Mbps)(EUCell_Eric)")
                    {
                        Vendor = "Ericsson";
                    }



                    for (int k = 0; k < Count - 1; k++)
                    {
                        DateTime Date = Convert.ToDateTime(FARAZ_Data[k + 1, 1]);
                        string Cell_Name = FARAZ_Data[k + 1, 2].ToString();
                        string Sector = Cell_Name.Substring(5, 9);
                        string Site = Sector.Substring(0, 8);





                        //string str2 = Regex.Replace(NE, "[^a-zA-Z0-9]", " ");      //?? ???????? ?? ??? ?? ??? ? ??? ??? ?? ?? ??????? ???? ????? ??
                        //str2 = Regex.Replace(str2, " {2,}", " ").Trim();           //????? ??????? ???? ??? ?? ?? ?? ?? ?? ??????? ???? ????? ?? ???
                        //string[] Split_Description = str2.Split(' ');
                        //string BSC = Split_Description[0];

                        //string Tech = Split_Description[0].Substring(0, 1);
                        //string Tech_Last = Split_Description[0].Substring(Split_Description[0].Length - 1, 1);

                        //string CellName = "";
                        //if ((Tech == "B" && (Tech_Last == "E" || Tech_Last == "H" || Tech_Last == "N")) || Split_Description[0].Length == 2)
                        //{
                        //    //int c1 = Split_Description.Length;
                        //    CellName = Split_Description[1];
                        //}
                        //else
                        //{
                        //    int c1 = Split_Description.Length;
                        //    CellName = Split_Description[c1 - 1];
                        //}


                        //if (CellName.Length == 7)
                        //{
                        //    Site = CellName.Substring(0, 6);
                        //}
                        //if (CellName.Length > 7)
                        //{
                        //    Site = CellName.Substring(0, 8);
                        //}



                        string ERAB_Setup_SR = "";
                        string UE_DL_THR = "";
                        string Cell_Availability = "";
                        string CSFB_Success_Rate = "";
                        string Volte_Traffic = "";
                        string ERAB_Drop_Rate = "";
                        string Intra_Freq_HO_SR = "";
                        string RRC_Connection_SR = "";
                        string PS_Traffic_Daily = "";

                        if (Vendor == "Ericsson")
                        {
                            if (FARAZ_Data[k + 1, 8] != null)
                            {
                                ERAB_Setup_SR = FARAZ_Data[k + 1, 8].ToString();
                            }
                            else
                            {
                                ERAB_Setup_SR = "";
                            }
                            if (FARAZ_Data[k + 1, 3] != null)
                            {
                                UE_DL_THR = FARAZ_Data[k + 1, 3].ToString();
                            }
                            else
                            {
                                UE_DL_THR = "";
                            }
                            if (FARAZ_Data[k + 1, 4] != null)
                            {
                                Cell_Availability = FARAZ_Data[k + 1, 4].ToString();
                            }
                            else
                            {
                                Cell_Availability = "";
                            }
                            if (FARAZ_Data[k + 1, 5] != null)
                            {
                                CSFB_Success_Rate = FARAZ_Data[k + 1, 5].ToString();
                            }
                            else
                            {
                                CSFB_Success_Rate = "";
                            }
                            if (FARAZ_Data[k + 1, 6] != null)
                            {
                                Volte_Traffic = FARAZ_Data[k + 1, 6].ToString();
                            }
                            else
                            {
                                Volte_Traffic = "";
                            }
                            if (FARAZ_Data[k + 1, 7] != null)
                            {
                                ERAB_Drop_Rate = FARAZ_Data[k + 1, 7].ToString();
                            }
                            else
                            {
                                ERAB_Drop_Rate = "";
                            }

                            if (FARAZ_Data[k + 1, 9] != null)
                            {
                                Intra_Freq_HO_SR = FARAZ_Data[k + 1, 9].ToString();
                            }
                            else
                            {
                                Intra_Freq_HO_SR = "";
                            }
                            if (FARAZ_Data[k + 1, 10] != null)
                            {
                                RRC_Connection_SR = FARAZ_Data[k + 1, 10].ToString();
                            }
                            else
                            {
                                RRC_Connection_SR = "";
                            }

                            if (FARAZ_Data[k + 1, 11] != null)
                            {
                                PS_Traffic_Daily = FARAZ_Data[k + 1, 11].ToString();
                            }
                            else
                            {
                                PS_Traffic_Daily = "";
                            }
                        }


                        if (Vendor == "Huawei")
                        {

                            if (FARAZ_Data[k + 1, 7] != null)
                            {
                                ERAB_Setup_SR = FARAZ_Data[k + 1, 7].ToString();
                            }
                            else
                            {
                                ERAB_Setup_SR = "";
                            }
                            if (FARAZ_Data[k + 1, 3] != null)
                            {
                                UE_DL_THR = FARAZ_Data[k + 1, 3].ToString();
                            }
                            else
                            {
                                UE_DL_THR = "";
                            }
                            if (FARAZ_Data[k + 1, 4] != null)
                            {
                                Cell_Availability = FARAZ_Data[k + 1, 4].ToString();
                            }
                            else
                            {
                                Cell_Availability = "";
                            }
                            if (FARAZ_Data[k + 1, 5] != null)
                            {
                                CSFB_Success_Rate = FARAZ_Data[k + 1, 5].ToString();
                            }
                            else
                            {
                                CSFB_Success_Rate = "";
                            }
                            if (FARAZ_Data[k + 1, 11] != null)
                            {
                                Volte_Traffic = FARAZ_Data[k + 1, 11].ToString();
                            }
                            else
                            {
                                Volte_Traffic = "";
                            }
                            if (FARAZ_Data[k + 1, 6] != null)
                            {
                                ERAB_Drop_Rate = FARAZ_Data[k + 1, 6].ToString();
                            }
                            else
                            {
                                ERAB_Drop_Rate = "";
                            }

                            if (FARAZ_Data[k + 1, 8] != null)
                            {
                                Intra_Freq_HO_SR = FARAZ_Data[k + 1, 8].ToString();
                            }
                            else
                            {
                                Intra_Freq_HO_SR = "";
                            }
                            if (FARAZ_Data[k + 1, 9] != null)
                            {
                                RRC_Connection_SR = FARAZ_Data[k + 1, 9].ToString();
                            }
                            else
                            {
                                RRC_Connection_SR = "";
                            }

                            if (FARAZ_Data[k + 1, 10] != null)
                            {
                                PS_Traffic_Daily = FARAZ_Data[k + 1, 10].ToString();
                            }
                            else
                            {
                                PS_Traffic_Daily = "";
                            }

                        }
                        if (Vendor == "Nokia")
                        {
                            if (FARAZ_Data[k + 1, 5] != null)
                            {
                                ERAB_Setup_SR = FARAZ_Data[k + 1, 5].ToString();
                            }
                            else
                            {
                                ERAB_Setup_SR = "";
                            }
                            if (FARAZ_Data[k + 1, 11] != null)
                            {
                                UE_DL_THR = FARAZ_Data[k + 1, 11].ToString();
                            }
                            else
                            {
                                UE_DL_THR = "";
                            }
                            if (FARAZ_Data[k + 1, 3] != null)
                            {
                                Cell_Availability = FARAZ_Data[k + 1, 3].ToString();
                            }
                            else
                            {
                                Cell_Availability = "";
                            }
                            if (FARAZ_Data[k + 1, 6] != null)
                            {
                                CSFB_Success_Rate = FARAZ_Data[k + 1, 6].ToString();
                            }
                            else
                            {
                                CSFB_Success_Rate = "";
                            }
                            if (FARAZ_Data[k + 1, 10] != null)
                            {
                                Volte_Traffic = FARAZ_Data[k + 1, 10].ToString();
                            }
                            else
                            {
                                Volte_Traffic = "";
                            }
                            if (FARAZ_Data[k + 1, 4] != null)
                            {
                                ERAB_Drop_Rate = FARAZ_Data[k + 1, 4].ToString();
                            }
                            else
                            {
                                ERAB_Drop_Rate = "";
                            }

                            if (FARAZ_Data[k + 1, 7] != null)
                            {
                                Intra_Freq_HO_SR = FARAZ_Data[k + 1, 7].ToString();
                            }
                            else
                            {
                                Intra_Freq_HO_SR = "";
                            }
                            if (FARAZ_Data[k + 1, 8] != null)
                            {
                                RRC_Connection_SR = FARAZ_Data[k + 1, 8].ToString();
                            }
                            else
                            {
                                RRC_Connection_SR = "";
                            }

                            if (FARAZ_Data[k + 1, 9] != null)
                            {
                                PS_Traffic_Daily = FARAZ_Data[k + 1, 9].ToString();
                            }
                            else
                            {
                                PS_Traffic_Daily = "";
                            }
                        }

                        //  Data_Table_4G1.Rows.Add(Date, Site, Sector, Cell_Name, ERAB_Setup_SR, UE_DL_THR, Cell_Availability, CSFB_Success_Rate, Volte_Traffic, ERAB_Drop_Rate, Intra_Freq_HO_SR, RRC_Connection_SR, PS_Traffic_Daily);


                        Data_Table_4G.Rows.Add(Date, Sector, ERAB_Setup_SR, UE_DL_THR, Cell_Availability, CSFB_Success_Rate, Volte_Traffic, ERAB_Drop_Rate, Intra_Freq_HO_SR, RRC_Connection_SR, PS_Traffic_Daily, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, Site);

                    }

                }




                // Group By 
                //      var groupedData = (from b in Data_Table_4G1.AsEnumerable()
                //                        group b by new
                //                        {
                //                            Date=b.Field<DateTime>("Date"),
                //                            Sector=b.Field<string>("Sector")
                //                        } into g
                //                        select new
                //                        {
                //                            Date = g.Key.Date,
                //                            Sector = g.Key.Sector,

                //                            //ChargeTag = g.Key,

                //                            ERAB_Setup_SR = g.Select(x => x.Field<string>("ERAB Setup SR, TH=96"))
                //.Where(s => !String.IsNullOrEmpty(s))
                //.Select(Convert.ToDouble)
                //.Average(),
                //                            UE_DL_THR = g.Select(x => x.Field<string>("UE DL THR (Mbps), TH=4"))
                //.Where(s => !String.IsNullOrEmpty(s))
                //.Select(Convert.ToDouble)
                //.Average(),

                //                            Cell_Availability = g.Select(x => x.Field<string>("Cell Availability, TH=99"))
                //.Where(s => !String.IsNullOrEmpty(s))
                //.Select(Convert.ToDouble)
                //.Average(),

                //                            CSFB_Success_Rate = g.Select(x => x.Field<string>("CSFB_Success_Rate, TH=95"))
                //.Where(s => !String.IsNullOrEmpty(s))
                //.Select(Convert.ToDouble)
                //.Average(),

                //                            Volte_Traffic = g.Select(x => x.Field<string>("Volte_Traffic, TH=0"))
                //.Where(s => !String.IsNullOrEmpty(s))
                //.Select(Convert.ToDouble)
                //.Sum(),

                //                            ERAB_Drop_Rate = g.Select(x => x.Field<string>("ERAB Drop Rate, TH=3"))
                //.Where(s => !String.IsNullOrEmpty(s))
                //.Select(Convert.ToDouble)
                //.Average(),

                //                            Intra_Freq_HO_SR = g.Select(x => x.Field<string>("Intra Freq HO SR, TH=95"))
                //.Where(s => !String.IsNullOrEmpty(s))
                //.Select(Convert.ToDouble)
                //.Average(),

                //                            RRC_Connection_SR = g.Select(x => x.Field<string>("RRC Connection SR, TH=96"))
                //.Where(s => !String.IsNullOrEmpty(s))
                //.Select(Convert.ToDouble)
                //.Average(),

                //                            PS_Traffic_Daily = g.Select(x => x.Field<string>("PS_Traffic_Daily (GB), TH=0"))
                //.Where(s => !String.IsNullOrEmpty(s))
                //.Select(Convert.ToDouble)
                //.Sum()
                //                        }).ToList();



                //      for (int i = 0; i <= groupedData.Count-1; i++)
                //      {
                //          DateTime Date1 = Convert.ToDateTime(groupedData[i].Date.ToString());
                //          string Cell_Name1 = groupedData[i].Sector.ToString();
                //          string Site1 = Cell_Name1.Substring(0, 8);
                //          string ERAB_Setup_SR1 = groupedData[i].ERAB_Setup_SR.ToString();
                //          string UE_DL_THR1 = groupedData[i].UE_DL_THR.ToString();
                //          string Cell_Availability1 = groupedData[i].Cell_Availability.ToString();
                //          string CSFB_Success_Rate1 = groupedData[i].CSFB_Success_Rate.ToString();
                //          string Volte_Traffic1 = groupedData[i].Volte_Traffic.ToString();
                //          string ERAB_Drop_Rate1 = groupedData[i].ERAB_Drop_Rate.ToString();
                //          string Intra_Freq_HO_SR1 = groupedData[i].Intra_Freq_HO_SR.ToString();
                //          string RRC_Connection_SR1 = groupedData[i].RRC_Connection_SR.ToString();
                //          string PS_Traffic_Daily1 = groupedData[i].PS_Traffic_Daily.ToString();

                //          Data_Table_4G.Rows.Add(Date1, Cell_Name1, ERAB_Setup_SR1, UE_DL_THR1, Cell_Availability1, CSFB_Success_Rate1, Volte_Traffic1, ERAB_Drop_Rate1, Intra_Freq_HO_SR1, RRC_Connection_SR1, PS_Traffic_Daily1,  -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, Site1);


                //      }






                Site_Data_Table_4G = new DataTable();
                Site_Data_Table_4G.Columns.Add("Site", typeof(string));
                Site_Data_Table_4G.Columns.Add("KPI Zero Status", typeof(string));
                Site_Data_Table_4G.Columns.Add("Rejected Cell List", typeof(string));



                //dataGridView1.ColumnCount = 12;
                dataGridView1.Invoke(new Action(() => dataGridView1.ColumnCount = 13));

                //dataGridView1.Rows.Clear();
                dataGridView1.Invoke(new Action(() => dataGridView1.Rows.Clear()));

                //dataGridView1.RowCount = Data_Table_4G.Rows.Count + 1;
                dataGridView1.Invoke(new Action(() => dataGridView1.RowCount = Data_Table_4G.Rows.Count + 1));

                //dataGridView1.Rows[0].Cells[0].Value = "Date"; dataGridView1.Columns[0].Width = 100;
                //dataGridView1.Rows[0].Cells[1].Value = "Site"; dataGridView1.Columns[1].Width = 100;
                //dataGridView1.Rows[0].Cells[2].Value = "Sector"; dataGridView1.Columns[2].Width = 100;
                //dataGridView1.Rows[0].Cells[3].Value = "UE DL THR (Mbps)"; dataGridView1.Columns[3].Width = 100;
                //dataGridView1.Rows[0].Cells[4].Value = "UE UL THR (Mbps)"; dataGridView1.Columns[4].Width = 100;
                //dataGridView1.Rows[0].Cells[5].Value = "ERAB Drop Rate"; dataGridView1.Columns[5].Width = 100;
                //dataGridView1.Rows[0].Cells[6].Value = "ERAB Setup SR"; dataGridView1.Columns[6].Width = 100;
                //dataGridView1.Rows[0].Cells[7].Value = "Intra Freq HO SR"; dataGridView1.Columns[7].Width = 100;
                //dataGridView1.Rows[0].Cells[8].Value = "RRC Connection SR"; dataGridView1.Columns[8].Width = 100;
                //dataGridView1.Rows[0].Cells[9].Value = "Cell Availability"; dataGridView1.Columns[9].Width = 100;
                //dataGridView1.Rows[0].Cells[10].Value = "Daily PS Traffic (GB)"; dataGridView1.Columns[10].Width = 100;
                //dataGridView1.Rows[0].Cells[11].Value = "Cell Status"; dataGridView1.Columns[11].Width = 100;




                dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[0].Value = "Date")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[0].Width = 100));
                dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[1].Value = "Site")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[1].Width = 100));
                dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[2].Value = "Sector")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[2].Width = 100));
                dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[3].Value = "ERAB Setup SR")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[3].Width = 100));
                dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[4].Value = "UE DL THR (Mbps)")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[4].Width = 100));
                dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[5].Value = "Cell Availability")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[5].Width = 100));
                dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[6].Value = "CSFB SR")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[6].Width = 100));
                dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[7].Value = "Volte Traffic (Erlang)")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[7].Width = 100));
                dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[8].Value = "ERAB Drop Rate")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[8].Width = 100));
                dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[9].Value = "Intra Freq HO SR")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[9].Width = 100));
                dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[10].Value = "RRC Connection SR")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[10].Width = 100));
                dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[11].Value = "Daily PS Traffic (GB)")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[11].Width = 100));
                dataGridView1.Invoke(new Action(() => dataGridView1.Rows[0].Cells[12].Value = "Cell Status")); dataGridView1.Invoke(new Action(() => dataGridView1.Columns[12].Width = 100));


                progressBar1.Minimum = 0;



                double ERAB_Setup_SR_TH = 96;
                double UE_DL_THR_TH = 8;
                double Cell_Availability_TH = 99;
                double CSFB_Success_Rate_TH = 95;
                double Volte_Traffic_TH = 0;
                double ERAB_Drop_Rate_TH = 3;
                double Intra_Freq_HO_SR_TH = 95;
                double RRC_Connection_SR_TH = 96;
                double PS_Traffic_Daily_TH = 0;





                if (Data_Table_4G.Rows.Count == 0)
                {
                    MessageBox.Show("There is no Data in Database");
                }

                if (Data_Table_4G.Rows.Count != 0)
                {
                    //  progressBar1.Maximum = Data_Table_4G.Rows.Count - 1;
                    progressBar1.Invoke(new Action(() => progressBar1.Maximum = Data_Table_4G.Rows.Count - 1));
                    for (int k = 0; k < Data_Table_4G.Rows.Count; k++)
                    {
                        int result = 0;

                        // Date
                        //dataGridView1.Rows[k + 1].Cells[0].Value = Data_Table_4G.Rows[k][0];
                        dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[0].Value = Data_Table_4G.Rows[k][0]));


                        // Site
                        string Cell = Data_Table_4G.Rows[k][1].ToString();
                        //  dataGridView1.Rows[k + 1].Cells[1].Value = Cell.Substring(0, 8);
                        dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[1].Value = Cell.Substring(0, 8)));
                        Data_Table_4G.Rows[k][21] = Cell.Substring(0, 8);


                        // Cell
                        //dataGridView1.Rows[k + 1].Cells[2].Value = Data_Table_4G.Rows[k][1];
                        dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[2].Value = Data_Table_4G.Rows[k][1]));




                        // ERAB Setup SR
                        // dataGridView1.Rows[k + 1].Cells[6].Value = Data_Table_4G.Rows[k][3];
                        dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[3].Value = Data_Table_4G.Rows[k][2]));
                        if (Data_Table_4G.Rows[k][2].ToString() == "")
                        {
                            Data_Table_4G.Rows[k][11] = -1;
                        }
                        else if (Convert.ToDouble(Data_Table_4G.Rows[k][2]) < ERAB_Setup_SR_TH)
                        {
                            Data_Table_4G.Rows[k][11] = 1; result++; dataGridView1.Rows[k + 1].Cells[3].Style.BackColor = Color.Orange;
                        }
                        else if (Convert.ToDouble(Data_Table_4G.Rows[k][2]) >= ERAB_Setup_SR_TH)
                        {
                            Data_Table_4G.Rows[k][11] = 0;
                        }



                        // UE DL THR
                        //  dataGridView1.Rows[k + 1].Cells[3].Value = Data_Table_4G.Rows[k][2];
                        dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[4].Value = Data_Table_4G.Rows[k][3]));
                        if (Data_Table_4G.Rows[k][3].ToString() == "")
                        {
                            Data_Table_4G.Rows[k][12] = -1;
                        }
                        else if (Convert.ToDouble(Data_Table_4G.Rows[k][3]) < UE_DL_THR_TH)
                        {
                            Data_Table_4G.Rows[k][12] = 1; result++; dataGridView1.Rows[k + 1].Cells[4].Style.BackColor = Color.Orange;
                        }
                        else if (Convert.ToDouble(Data_Table_4G.Rows[k][3]) >= UE_DL_THR_TH)
                        {
                            Data_Table_4G.Rows[k][12] = 0;
                        }


                        // Cell Availability
                        // dataGridView1.Rows[k + 1].Cells[9].Value = Data_Table_4G.Rows[k][3];
                        dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[5].Value = Data_Table_4G.Rows[k][4]));
                        if (Data_Table_4G.Rows[k][4].ToString() == "")
                        {
                            Data_Table_4G.Rows[k][13] = -1;
                        }
                        else if (Convert.ToDouble(Data_Table_4G.Rows[k][4]) < Cell_Availability_TH && Convert.ToDouble(Data_Table_4G.Rows[k][4]) > 0)
                        {
                            Data_Table_4G.Rows[k][13] = 1; result++; dataGridView1.Rows[k + 1].Cells[5].Style.BackColor = Color.Orange;
                        }
                        else if (Convert.ToDouble(Data_Table_4G.Rows[k][4]) >= Cell_Availability_TH)
                        {
                            Data_Table_4G.Rows[k][13] = 0;
                        }



                        // CSFB SR
                        // dataGridView1.Rows[k + 1].Cells[5].Value = Data_Table_4G.Rows[k][4];
                        dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[6].Value = Data_Table_4G.Rows[k][5]));
                        if (Data_Table_4G.Rows[k][5].ToString() == "")
                        {
                            Data_Table_4G.Rows[k][14] = -1;
                        }
                        else if (Convert.ToDouble(Data_Table_4G.Rows[k][5]) < CSFB_Success_Rate_TH)
                        {
                            Data_Table_4G.Rows[k][14] = 1; result++; dataGridView1.Rows[k + 1].Cells[6].Style.BackColor = Color.Orange;
                        }
                        else if (Convert.ToDouble(Data_Table_4G.Rows[k][5]) >= CSFB_Success_Rate_TH)
                        {
                            Data_Table_4G.Rows[k][14] = 0;
                        }




                        // Volte
                        // dataGridView1.Rows[k + 1].Cells[5].Value = Data_Table_4G.Rows[k][4];
                        dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[7].Value = Data_Table_4G.Rows[k][6]));
                        if (Data_Table_4G.Rows[k][6].ToString() == "")
                        {
                            Data_Table_4G.Rows[k][15] = -1;
                        }
                        else if (Convert.ToDouble(Data_Table_4G.Rows[k][6]) == Volte_Traffic_TH)
                        {
                            Data_Table_4G.Rows[k][15] = 1; result++; dataGridView1.Rows[k + 1].Cells[7].Style.BackColor = Color.Orange;
                        }
                        else if (Convert.ToDouble(Data_Table_4G.Rows[k][6]) > Volte_Traffic_TH)
                        {
                            Data_Table_4G.Rows[k][15] = 0;
                        }




                        // ERAB Drop Rate
                        // dataGridView1.Rows[k + 1].Cells[5].Value = Data_Table_4G.Rows[k][4];
                        dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[8].Value = Data_Table_4G.Rows[k][7]));
                        if (Data_Table_4G.Rows[k][7].ToString() == "")
                        {
                            Data_Table_4G.Rows[k][16] = -1;
                        }
                        else if (Convert.ToDouble(Data_Table_4G.Rows[k][7]) > ERAB_Drop_Rate_TH)
                        {
                            Data_Table_4G.Rows[k][16] = 1; result++; dataGridView1.Rows[k + 1].Cells[8].Style.BackColor = Color.Orange;
                        }
                        else if (Convert.ToDouble(Data_Table_4G.Rows[k][7]) <= ERAB_Drop_Rate_TH)
                        {
                            Data_Table_4G.Rows[k][16] = 0;
                        }





                        // Intra Freq HO SR
                        //  dataGridView1.Rows[k + 1].Cells[7].Value = Data_Table_4G.Rows[k][3];
                        dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[9].Value = Data_Table_4G.Rows[k][8]));
                        if (Data_Table_4G.Rows[k][8].ToString() == "")
                        {
                            Data_Table_4G.Rows[k][17] = -1;
                        }
                        else if (Convert.ToDouble(Data_Table_4G.Rows[k][8]) < Intra_Freq_HO_SR_TH)
                        {
                            Data_Table_4G.Rows[k][17] = 1; result++; dataGridView1.Rows[k + 1].Cells[9].Style.BackColor = Color.Orange;
                        }
                        else if (Convert.ToDouble(Data_Table_4G.Rows[k][8]) >= Intra_Freq_HO_SR_TH)
                        {
                            Data_Table_4G.Rows[k][17] = 0;
                        }


                        // RRC Connection SR
                        //dataGridView1.Rows[k + 1].Cells[8].Value = Data_Table_4G.Rows[k][3];
                        dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[10].Value = Data_Table_4G.Rows[k][9]));
                        if (Data_Table_4G.Rows[k][9].ToString() == "")
                        {
                            Data_Table_4G.Rows[k][18] = -1;
                        }
                        else if (Convert.ToDouble(Data_Table_4G.Rows[k][9]) < RRC_Connection_SR_TH)
                        {
                            Data_Table_4G.Rows[k][18] = 1; result++; dataGridView1.Rows[k + 1].Cells[10].Style.BackColor = Color.Orange;
                        }
                        else if (Convert.ToDouble(Data_Table_4G.Rows[k][9]) >= RRC_Connection_SR_TH)
                        {
                            Data_Table_4G.Rows[k][18] = 0;
                        }





                        // Daily PS Traffic (GB)
                        // dataGridView1.Rows[k + 1].Cells[10].Value = Data_Table_4G.Rows[k][3];
                        dataGridView1.Invoke(new Action(() => dataGridView1.Rows[k + 1].Cells[11].Value = Data_Table_4G.Rows[k][10]));
                        if (Data_Table_4G.Rows[k][10].ToString() == "")
                        {
                            Data_Table_4G.Rows[k][19] = -1;
                        }
                        else if (Convert.ToDouble(Data_Table_4G.Rows[k][10]) == PS_Traffic_Daily_TH)
                        {
                            Data_Table_4G.Rows[k][19] = 1; result++; dataGridView1.Rows[k + 1].Cells[11].Style.BackColor = Color.Orange;
                        }
                        else if (Convert.ToDouble(Data_Table_4G.Rows[k][10]) > PS_Traffic_Daily_TH)
                        {
                            Data_Table_4G.Rows[k][19] = 0;
                        }


                        // Cell Score
                        if (Convert.ToInt16(Data_Table_4G.Rows[k][19]) == -1)  // Traffic
                        {
                            Data_Table_4G.Rows[k][20] = 1; dataGridView1.Rows[k + 1].Cells[12].Value = "Not Updated";
                        }
                        else if (result == 0)
                        {
                            Data_Table_4G.Rows[k][20] = 0.1; dataGridView1.Rows[k + 1].Cells[12].Value = "Passed";
                        }
                        if (result > 0)
                        {
                            Data_Table_4G.Rows[k][20] = 0; dataGridView1.Rows[k + 1].Cells[12].Value = "Rejected";
                        }




                        progressBar1.Invoke(new Action(() => progressBar1.Value = k));
                        //progressBar1.Value = k;
                    }


                    var distinctIds = Data_Table_4G.AsEnumerable()
    .Select(s => new
    {
        id = s.Field<string>("Site"),
    })
    .Distinct().ToList();

                    for (int j = 0; j < distinctIds.Count; j++)
                    {
                        var cell_data = (from p in Data_Table_4G.AsEnumerable()
                                         where p.Field<string>("Site") == distinctIds[j].id
                                         select p).ToList();


                        double multiplier = 1;
                        for (int h = 0; h < cell_data.Count; h++)
                        {
                            multiplier = multiplier * Convert.ToDouble(cell_data[h].ItemArray[20]);

                        }

                        if (multiplier > 0 && multiplier < 1)
                        {
                            Site_Data_Table_4G.Rows.Add(distinctIds[j].id, "Passed");
                        }
                        if (multiplier == 0)
                        {
                            Site_Data_Table_4G.Rows.Add(distinctIds[j].id, "Rejected");
                        }
                        if (multiplier == 1)
                        {
                            Site_Data_Table_4G.Rows.Add(distinctIds[j].id, "Not Updated");
                        }

                    }



                }


            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (Technology == "2G")
            {
                XLWorkbook wb = new XLWorkbook();
                wb.Worksheets.Add(Data_Table_2G, "Data Table");
                wb.Worksheets.Add(Site_Data_Table_2G, "Result");
                var saveFileDialog = new SaveFileDialog
                {
                    FileName = Technology + "_KPI_Zero_Check",
                    Filter = "Excel files|*.xlsx",
                    Title = "Save an Excel File"
                };

                IXLWorksheet Source_worksheet = wb.Worksheet("Data Table");
                int number_of_rows = Source_worksheet.RowsUsed().Count();

                string[] Rejected_Cell = new string[10000];
                string[] Rejected_Site = new string[10000];
                int r1 = 0;

                for (int k = 1; k <= number_of_rows; k++)
                {
                    for (int j = 1; j <= 6; j++)
                    {
                        string val = Source_worksheet.Cell(k, j + 9).Value.ToString();
                        if (val == "1")
                        {
                            Source_worksheet.Cell(k, j + 3).Style.Fill.BackgroundColor = XLColor.Red;
                            Source_worksheet.Cell(k, 16).Style.Fill.BackgroundColor = XLColor.Red;
                        }
                    }

                }

                IXLWorksheet Source_worksheet1 = wb.Worksheet("Result");
                int number_of_rows1 = Source_worksheet1.RowsUsed().Count();
                ///Source_worksheet1.Cell(0, 3).Value = "Rejected Cell List";
                for (int k = 1; k <= number_of_rows1; k++)  // Sheet of Results (Source_worksheet1)
                {
                    string site = Source_worksheet1.Cell(k, 1).Value.ToString();
                    int cell_indexer = 1;
                    string[] cell_list = new string[10];
                    string Cells = "";
                    for (int i = 1; i <= number_of_rows; i++)     // Sheet if Data  (Source_worksheet)
                    {
                        string site1 = Source_worksheet.Cell(i, 17).Value.ToString();
                        string Cell = Source_worksheet.Cell(i, 3).Value.ToString();
                        string val = Source_worksheet.Cell(i, 16).Value.ToString();

                        if (site == site1 && val == "0")
                        {
                            if (!cell_list.Contains(Cell))
                            {
                                //Source_worksheet1.Cell(k, cell_indexer + 2).Value = Cell;
                                Cells = Cells + Cell + ", ";
                                Source_worksheet1.Cell(k, 3).Value = Cells;
                                cell_list[cell_indexer - 1] = Cell;
                                cell_indexer++;
                            }
                        }
                    }

                }



                saveFileDialog.ShowDialog();

                if (!String.IsNullOrWhiteSpace(saveFileDialog.FileName))
                    wb.SaveAs(saveFileDialog.FileName);

                MessageBox.Show("Finished");
            }




            if (Technology == "2G-MCI")
            {
                XLWorkbook wb = new XLWorkbook();
                wb.Worksheets.Add(Data_Table_2G, "Data Table");
                wb.Worksheets.Add(Site_Data_Table_2G, "Result");
                var saveFileDialog = new SaveFileDialog
                {
                    FileName = "2G_KPI_Zero_Check",
                    Filter = "Excel files|*.xlsx",
                    Title = "Save an Excel File"
                };

                IXLWorksheet Source_worksheet = wb.Worksheet("Data Table");
                int number_of_rows = Source_worksheet.RowsUsed().Count();

                string[] Rejected_Cell = new string[10000];
                string[] Rejected_Site = new string[10000];
                int r1 = 0;

                for (int k = 1; k <= number_of_rows; k++)
                {
                    for (int j = 1; j <= 7; j++)
                    {
                        string val = Source_worksheet.Cell(k, j + 10).Value.ToString();
                        if (val == "1")
                        {
                            Source_worksheet.Cell(k, j + 3).Style.Fill.BackgroundColor = XLColor.Red;
                            Source_worksheet.Cell(k, 18).Style.Fill.BackgroundColor = XLColor.Red;
                            r1++;
                        }
                    }

                }




                IXLWorksheet Source_worksheet1 = wb.Worksheet("Result");
                int number_of_rows1 = Source_worksheet1.RowsUsed().Count();
                ///Source_worksheet1.Cell(0, 3).Value = "Rejected Cell List";
                for (int k = 1; k <= number_of_rows1; k++)  // Sheet of Results (Source_worksheet1)
                {
                    string site = Source_worksheet1.Cell(k, 1).Value.ToString();
                    int cell_indexer = 1;
                    string[] cell_list = new string[10];
                    string Cells = "";
                    for (int i = 1; i <= number_of_rows; i++)     // Sheet if Data  (Source_worksheet)
                    {
                        string site1 = Source_worksheet.Cell(i, 19).Value.ToString();
                        string Cell = Source_worksheet.Cell(i, 3).Value.ToString();
                        string val = Source_worksheet.Cell(i, 18).Value.ToString();

                        if (site == site1 && val == "0")
                        {
                            if (!cell_list.Contains(Cell))
                            {
                                //Source_worksheet1.Cell(k, cell_indexer + 2).Value = Cell;
                                Cells = Cells + Cell + ", ";
                                Source_worksheet1.Cell(k, 3).Value = Cells;
                                cell_list[cell_indexer - 1] = Cell;
                                cell_indexer++;
                            }
                        }
                    }

                }



                saveFileDialog.ShowDialog();

                if (!String.IsNullOrWhiteSpace(saveFileDialog.FileName))
                    wb.SaveAs(saveFileDialog.FileName);

                MessageBox.Show("Finished");
            }





            if (Technology == "3G" || Technology == "3G-MCI")
            {
                XLWorkbook wb = new XLWorkbook();
                wb.Worksheets.Add(Data_Table_3G, "Data Table");
                wb.Worksheets.Add(Site_Data_Table_3G, "Result");
                var saveFileDialog = new SaveFileDialog
                {
                    FileName = "3G_KPI_Zero_Check",
                    Filter = "Excel files|*.xlsx",
                    Title = "Save an Excel File"
                };





                IXLWorksheet Source_worksheet = wb.Worksheet("Data Table");
                int number_of_rows = Source_worksheet.RowsUsed().Count();

                for (int k = 1; k <= number_of_rows; k++)
                {
                    for (int j = 1; j <= 10; j++)
                    {
                        string val = Source_worksheet.Cell(k, j + 13).Value.ToString();
                        if (val == "1")
                        {
                            Source_worksheet.Cell(k, j + 3).Style.Fill.BackgroundColor = XLColor.Red;
                            Source_worksheet.Cell(k, 24).Style.Fill.BackgroundColor = XLColor.Red;
                        }
                    }

                }

                IXLWorksheet Source_worksheet1 = wb.Worksheet("Result");
                int number_of_rows1 = Source_worksheet1.RowsUsed().Count();
                for (int k = 1; k <= number_of_rows1; k++)  // Sheet of Results (Source_worksheet1)
                {
                    string site = Source_worksheet1.Cell(k, 1).Value.ToString();
                    int cell_indexer = 1;
                    string[] cell_list = new string[10];
                    string Cells = "";
                    for (int i = 1; i <= number_of_rows; i++)     // Sheet if Data  (Source_worksheet)
                    {
                        string site1 = Source_worksheet.Cell(i, 25).Value.ToString();
                        string Cell = Source_worksheet.Cell(i, 3).Value.ToString();
                        string val = Source_worksheet.Cell(i, 24).Value.ToString();

                        if (site == site1 && val == "0")
                        {
                            if (!cell_list.Contains(Cell))
                            {
                                //Source_worksheet1.Cell(k, cell_indexer + 2).Value = Cell;
                                Cells = Cells + Cell + ", ";
                                Source_worksheet1.Cell(k, 3).Value = Cells;
                                cell_list[cell_indexer - 1] = Cell;
                                cell_indexer++;
                            }
                        }
                    }

                }



                saveFileDialog.ShowDialog();

                if (!String.IsNullOrWhiteSpace(saveFileDialog.FileName))
                    wb.SaveAs(saveFileDialog.FileName);

                MessageBox.Show("Finished");
            }

            if (Technology == "4G-MCI")
            {
                XLWorkbook wb = new XLWorkbook();
                wb.Worksheets.Add(Data_Table_4G, "Data Table");
                wb.Worksheets.Add(Site_Data_Table_4G, "Result");
                var saveFileDialog = new SaveFileDialog
                {
                    FileName = "4G_KPI_Zero_Check",
                    Filter = "Excel files|*.xlsx",
                    Title = "Save an Excel File"
                };

                IXLWorksheet Source_worksheet = wb.Worksheet("Data Table");
                int number_of_rows = Source_worksheet.RowsUsed().Count();

                for (int k = 1; k <= number_of_rows; k++)
                {
                    for (int j = 1; j <= 9; j++)
                    {
                        string val = Source_worksheet.Cell(k, j + 12).Value.ToString();
                        if (val == "1")
                        {
                            Source_worksheet.Cell(k, j + 3).Style.Fill.BackgroundColor = XLColor.Red;
                            Source_worksheet.Cell(k, 21).Style.Fill.BackgroundColor = XLColor.Red;
                        }
                    }

                }

                IXLWorksheet Source_worksheet1 = wb.Worksheet("Result");
                int number_of_rows1 = Source_worksheet1.RowsUsed().Count();
                for (int k = 1; k <= number_of_rows1; k++)  // Sheet of Results (Source_worksheet1)
                {
                    string site = Source_worksheet1.Cell(k, 1).Value.ToString();
                    int cell_indexer = 1;
                    string[] cell_list = new string[10];
                    string Cells = "";
                    for (int i = 1; i <= number_of_rows; i++)     // Sheet if Data  (Source_worksheet)
                    {

                        string site1 = Source_worksheet.Cell(i, 22).Value.ToString();
                        string Cell = Source_worksheet.Cell(i, 2).Value.ToString();
                        string val = Source_worksheet.Cell(i, 21).Value.ToString();


                        if (site == site1 && val == "0")
                        {
                            if (!cell_list.Contains(Cell))
                            {
                                //Source_worksheet1.Cell(k, cell_indexer + 2).Value = Cell;
                                Cells = Cells + Cell + ", ";
                                Source_worksheet1.Cell(k, 3).Value = Cells;
                                cell_list[cell_indexer - 1] = Cell;
                                cell_indexer++;
                            }
                        }
                    }

                }


                saveFileDialog.ShowDialog();

                if (!String.IsNullOrWhiteSpace(saveFileDialog.FileName))
                    wb.SaveAs(saveFileDialog.FileName);

                MessageBox.Show("Finished");
            }






            if (Technology == "4G")
            {
                XLWorkbook wb = new XLWorkbook();
                wb.Worksheets.Add(Data_Table_4G, "Data Table");
                wb.Worksheets.Add(Site_Data_Table_4G, "Result");
                var saveFileDialog = new SaveFileDialog
                {
                    FileName = "4G_KPI_Zero_Check",
                    Filter = "Excel files|*.xlsx",
                    Title = "Save an Excel File"
                };

                IXLWorksheet Source_worksheet = wb.Worksheet("Data Table");
                int number_of_rows = Source_worksheet.RowsUsed().Count();

                for (int k = 1; k <= number_of_rows; k++)
                {
                    for (int j = 1; j <= 8; j++)
                    {
                        string val = Source_worksheet.Cell(k, j + 10).Value.ToString();
                        if (val == "1")
                        {
                            Source_worksheet.Cell(k, j + 2).Style.Fill.BackgroundColor = XLColor.Red;
                            Source_worksheet.Cell(k, 19).Style.Fill.BackgroundColor = XLColor.Red;
                        }
                    }

                }

                IXLWorksheet Source_worksheet1 = wb.Worksheet("Result");
                int number_of_rows1 = Source_worksheet1.RowsUsed().Count();
                for (int k = 1; k <= number_of_rows1; k++)  // Sheet of Results (Source_worksheet1)
                {
                    string site = Source_worksheet1.Cell(k, 1).Value.ToString();
                    int cell_indexer = 1;
                    string[] cell_list = new string[10];
                    string Cells = "";
                    for (int i = 1; i <= number_of_rows; i++)     // Sheet if Data  (Source_worksheet)
                    {
                        string site1 = Source_worksheet.Cell(i, 20).Value.ToString();
                        string Cell = Source_worksheet.Cell(i, 2).Value.ToString();
                        string val = Source_worksheet.Cell(i, 19).Value.ToString();

                        if (site == site1 && val == "0")
                        {
                            if (!cell_list.Contains(Cell))
                            {
                                //Source_worksheet1.Cell(k, cell_indexer + 2).Value = Cell;
                                Cells = Cells + Cell + ", ";
                                Source_worksheet1.Cell(k, 3).Value = Cells;
                                cell_list[cell_indexer - 1] = Cell;
                                cell_indexer++;
                            }
                        }
                    }

                }


                saveFileDialog.ShowDialog();

                if (!String.IsNullOrWhiteSpace(saveFileDialog.FileName))
                    wb.SaveAs(saveFileDialog.FileName);

                MessageBox.Show("Finished");
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                Input_Type = "FARAZ";
                checkBox1.Checked = false;
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                Input_Type = "DataBase";
                checkBox2.Checked = false;
            }
        }

        private void textBox20_TextChanged(object sender, EventArgs e)
        {

        }
    }
}