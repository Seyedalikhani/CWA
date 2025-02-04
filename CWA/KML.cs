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
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;
using System.Reflection;
using System.Xml.Linq;
using System.Xml;
using System.Text.RegularExpressions;
using System.IO;

namespace CWA
{
    public partial class KML : Form
    {
        public KML()
        {
            InitializeComponent();
        }


        public Main form1;


        public KML(Form form)
        {
            InitializeComponent();
            form1 = (Main)form;
        }

        private void KML_Load(object sender, EventArgs e)
        {
            ConnectionString = @"Server=" + Server_Name + "; Database=" + DataBase_Name + "; User ID=cwpcApp;Password=cwpcApp@830625#Ahmad";
            connection = new SqlConnection(ConnectionString);
            connection.Open();


            KMLTable.Columns.Add("Location", typeof(string));
            //KMLTable.Columns.Add("LAT", typeof(string));
            //KMLTable.Columns.Add("LON", typeof(string));
            KMLTable.Columns.Add("SiteCode_Technology", typeof(string));
            KMLTable.Columns.Add("Location_Band", typeof(string));
            KMLTable.Columns.Add("Project", typeof(string));
            KMLTable.Columns.Add("Vendor", typeof(string));
            KMLTable.Columns.Add("City", typeof(string));
            KMLTable.Columns.Add("RSSV_Date_in_Spring", typeof(string));
            KMLTable.Columns.Add("Status", typeof(string));
            KMLTable.Columns.Add("Pass_Date", typeof(string));
            KMLTable.Columns.Add("Traffic", typeof(string));
            KMLTable.Columns.Add("Assignment_to_SC_Date", typeof(string));
            KMLTable.Columns.Add("SSV_Owner", typeof(string));
            KMLTable.Columns.Add("Submit_to_OPT_Team_Date", typeof(string));
            KMLTable.Columns.Add("Priorities", typeof(string));
            KMLTable.Columns.Add("Comment_External", typeof(string));
            KMLTable.Columns.Add("NAK_Team", typeof(string));
            KMLTable.Columns.Add("NAK_SCOPE", typeof(string));


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
        public void Query_Execution(String Query)
        {
            string Quary_String = Query;
            SqlCommand Quary_Command = new SqlCommand(Quary_String, connection);
            Quary_Command.CommandTimeout = 0;
            Quary_Command.ExecuteNonQuery();
        }




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




        public string ConnectionString = "";
        public SqlConnection connection = new SqlConnection();
        public string Server_Name = "PERFORMANCEDB";
        public string DataBase_Name = "Performance_NAK";
        public DataTable KMLTable = new DataTable();
        public DataTable ARASTable = new DataTable();
        public DataTable SSV_KMLTable = new DataTable();
        public string Site = "";
        public string Technology = "";
        public string Subcontractor = "";
        public string Projct = "";
        public string Periority = "";
        public string KML_Type = "Subcontractor_Project";



        public IXLWorksheet Source_worksheet = null;
        public Excel.Application xlApp { get; set; }
        public Excel.Workbook Template_workbook { get; set; }


        // Defining structure
        public struct Contractor
        {
            public string[] Projects;
            public string[] KMLText;
        }

        public struct Periority_Contractor
        {
            public string[] Projects;
            public string[] KMLText;
        }


        public string Fixed_part = @"<?xml version='1.0' encoding='UTF-8'?>
<kml xmlns='http://www.opengis.net/kml/2.2'
xmlns:gx='http://www.google.com/kml/ext/2.2'
xmlns:kml='http://www.opengis.net/kml/2.2'
xmlns:atom='http://www.w3.org/2005/Atom'>
<Document>
<name>
KML_Export_20240208-Subcons.kml
</name>
    <open>1</open>
    <Style id='s_hollow_hl'>
        <IconStyle>
            <color>ff000000</color>
            <scale>1.3</scale>
            <Icon>
                <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
            </Icon>
            <hotSpot x='0.5' y='0.5' xunits='fraction' yunits='fraction'/>
        </IconStyle>
        <LineStyle>
            <color>ff000000</color>
            <width>1</width>
        </LineStyle>
        <PolyStyle>
            <color>00000000</color>
        </PolyStyle>
    </Style>
    <StyleMap id='m_hollow'>
        <Pair>
            <key>normal</key>
            <styleUrl>#s_hollow</styleUrl>
        </Pair>
        <Pair>
            <key>highlight</key>
            <styleUrl>#s_hollow_hl</styleUrl>
        </Pair>
    </StyleMap>
    <Style id='s_hollow'>
        <IconStyle>
            <color>ff000000</color>
            <scale>1.1</scale>
            <Icon>
                <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
            </Icon>
            <hotSpot x='0.5' y='0.5' xunits='fraction' yunits='fraction'/>
        </IconStyle>
        <LineStyle>
            <color>ff000000</color>
            <width>1</width>
        </LineStyle>
        <PolyStyle>
            <color>00000000</color>
        </PolyStyle>
    </Style>
    <Style id='s_white_hl'>
        <IconStyle>
            <color>ffffffff</color>
            <scale>1.3</scale>
            <Icon>
                <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
            </Icon>
            <hotSpot x='0.5' y='0.5' xunits='fraction' yunits='fraction'/>
        </IconStyle>
        <LineStyle>
            <color>ff000000</color>
            <width>1</width>
        </LineStyle>
        <PolyStyle>
            <color>ffffffff</color>
        </PolyStyle>
    </Style>
    <StyleMap id='m_white'>
        <Pair>
            <key>normal</key>
            <styleUrl>#s_white</styleUrl>
        </Pair>
        <Pair>
            <key>highlight</key>
            <styleUrl>#s_white_hl</styleUrl>
        </Pair>
    </StyleMap>
    <Style id='s_white'>
        <IconStyle>
            <color>ffffffff</color>
            <scale>1.1</scale>
            <Icon>
                <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
            </Icon>
            <hotSpot x='0.5' y='0.5' xunits='fraction' yunits='fraction'/>
        </IconStyle>
        <LineStyle>
            <color>ff000000</color>
            <width>1</width>
        </LineStyle>
        <PolyStyle>
            <color>ffffffff</color>
        </PolyStyle>
    </Style>
    <Style id='s_red_hl'>
        <IconStyle>
            <color>ff0000ff</color>
            <scale>1.3</scale>
            <Icon>
                <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
            </Icon>
            <hotSpot x='0.5' y='0.5' xunits='fraction' yunits='fraction'/>
        </IconStyle>
        <LineStyle>
            <color>ff000000</color>
            <width>1</width>
        </LineStyle>
        <PolyStyle>
            <color>ff0000ff</color>
        </PolyStyle>
    </Style>
    <StyleMap id='m_red'>
        <Pair>
            <key>normal</key>
            <styleUrl>#s_red</styleUrl>
        </Pair>
        <Pair>
            <key>highlight</key>
            <styleUrl>#s_red_hl</styleUrl>
        </Pair>
    </StyleMap>
    <Style id='s_red'>
        <IconStyle>
            <color>ff0000ff</color>
            <scale>1.1</scale>
            <Icon>
                <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
            </Icon>
            <hotSpot x='0.5' y='0.5' xunits='fraction' yunits='fraction'/>
        </IconStyle>
        <LineStyle>
            <color>ff000000</color>
            <width>1</width>
        </LineStyle>
        <PolyStyle>
            <color>ff0000ff</color>
        </PolyStyle>
    </Style>
    <Style id='s_black_hl'>
        <IconStyle>
            <color>ff000000</color>
            <scale>1.3</scale>
            <Icon>
                <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
            </Icon>
            <hotSpot x='0.5' y='0.5' xunits='fraction' yunits='fraction'/>
        </IconStyle>
        <LineStyle>
            <color>ff000000</color>
            <width>1</width>
        </LineStyle>
        <PolyStyle>
            <color>ff000000</color>
        </PolyStyle>
    </Style>
    <StyleMap id='m_black'>
        <Pair>
            <key>normal</key>
            <styleUrl>#s_black</styleUrl>
        </Pair>
        <Pair>
            <key>highlight</key>
            <styleUrl>#s_black_hl</styleUrl>
        </Pair>
    </StyleMap>
    <Style id='s_black'>
        <IconStyle>
            <color>ff000000</color>
            <scale>1.1</scale>
            <Icon>
                <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
            </Icon>
            <hotSpot x='0.5' y='0.5' xunits='fraction' yunits='fraction'/>
        </IconStyle>
        <LineStyle>
            <color>ff000000</color>
            <width>1</width>
        </LineStyle>
        <PolyStyle>
            <color>ff000000</color>
        </PolyStyle>
    </Style>
    <Style id='s_green_hl'>
        <IconStyle>
            <color>ff00ff00</color>
            <scale>1.3</scale>
            <Icon>
                <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
            </Icon>
            <hotSpot x='0.5' y='0.5' xunits='fraction' yunits='fraction'/>
        </IconStyle>
        <LineStyle>
            <color>ff000000</color>
            <width>1</width>
        </LineStyle>
        <PolyStyle>
            <color>ff00ff00</color>
        </PolyStyle>
    </Style>
    <StyleMap id='m_green'>
        <Pair>
            <key>normal</key>
            <styleUrl>#s_green</styleUrl>
        </Pair>
        <Pair>
            <key>highlight</key>
            <styleUrl>#s_green_hl</styleUrl>
        </Pair>
    </StyleMap>
    <Style id='s_green'>
        <IconStyle>
            <color>ff00ff00</color>
            <scale>1.1</scale>
            <Icon>
                <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
            </Icon>
            <hotSpot x='0.5' y='0.5' xunits='fraction' yunits='fraction'/>
        </IconStyle>
        <LineStyle>
            <color>ff000000</color>
            <width>1</width>
        </LineStyle>
        <PolyStyle>
            <color>ff00ff00</color>
        </PolyStyle>
    </Style>
    <Style id='s_blue_hl'>
        <IconStyle>
            <color>ffff0000</color>
            <scale>1.3</scale>
            <Icon>
                <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
            </Icon>
            <hotSpot x='0.5' y='0.5' xunits='fraction' yunits='fraction'/>
        </IconStyle>
        <LineStyle>
            <color>ff000000</color>
            <width>1</width>
        </LineStyle>
        <PolyStyle>
            <color>ffff0000</color>
        </PolyStyle>
    </Style>
    <StyleMap id='m_blue'>
        <Pair>
            <key>normal</key>
            <styleUrl>#s_blue</styleUrl>
        </Pair>
        <Pair>
            <key>highlight</key>
            <styleUrl>#s_blue_hl</styleUrl>
        </Pair>
    </StyleMap>
    <Style id='s_blue'>
        <IconStyle>
            <color>ffff0000</color>
            <scale>1.1</scale>
            <Icon>
                <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
            </Icon>
            <hotSpot x='0.5' y='0.5' xunits='fraction' yunits='fraction'/>
        </IconStyle>
        <LineStyle>
            <color>ff000000</color>
            <width>1</width>
        </LineStyle>
        <PolyStyle>
            <color>ffff0000</color>
        </PolyStyle>
    </Style>
    <Style id='s_yellow_hl'>
        <IconStyle>
            <color>ff00ffff</color>
            <scale>1.3</scale>
            <Icon>
                <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
            </Icon>
            <hotSpot x='0.5' y='0.5' xunits='fraction' yunits='fraction'/>
        </IconStyle>
        <LineStyle>
            <color>ff000000</color>
            <width>1</width>
        </LineStyle>
        <PolyStyle>
            <color>ff00ffff</color>
        </PolyStyle>
    </Style>
    <StyleMap id='m_yellow'>
        <Pair>
            <key>normal</key>
            <styleUrl>#s_yellow</styleUrl>
        </Pair>
        <Pair>
            <key>highlight</key>
            <styleUrl>#s_yellow_hl</styleUrl>
        </Pair>
    </StyleMap>
    <Style id='s_yellow'>
        <IconStyle>
            <color>ff00ffff</color>
            <scale>1.1</scale>
            <Icon>
                <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
            </Icon>
            <hotSpot x='0.5' y='0.5' xunits='fraction' yunits='fraction'/>
        </IconStyle>
        <LineStyle>
            <color>ff000000</color>
            <width>1</width>
        </LineStyle>
        <PolyStyle>
            <color>ff00ffff</color>
        </PolyStyle>
    </Style>
    <Style id='s_orange_hl'>
        <IconStyle>
            <color>ff00aaff</color>
            <scale>1.3</scale>
            <Icon>
                <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
            </Icon>
            <hotSpot x='0.5' y='0.5' xunits='fraction' yunits='fraction'/>
        </IconStyle>
        <LineStyle>
            <color>ff000000</color>
            <width>1</width>
        </LineStyle>
        <PolyStyle>
            <color>ff00aaff</color>
        </PolyStyle>
    </Style>
    <StyleMap id='m_orange'>
        <Pair>
            <key>normal</key>
            <styleUrl>#s_orange</styleUrl>
        </Pair>
        <Pair>
            <key>highlight</key>
            <styleUrl>#s_orange_hl</styleUrl>
        </Pair>
    </StyleMap>
    <Style id='s_orange'>
        <IconStyle>
            <color>ff00aaff</color>
            <scale>1.1</scale>
            <Icon>
                <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
            </Icon>
            <hotSpot x='0.5' y='0.5' xunits='fraction' yunits='fraction'/>
        </IconStyle>
        <LineStyle>
            <color>ff000000</color>
            <width>1</width>
        </LineStyle>
        <PolyStyle>
            <color>ff00aaff</color>
        </PolyStyle>
    </Style>
    <Style id='s_tiad_hl'>
        <IconStyle>
            <color>ff5f1f9f</color>
            <scale>1.3</scale>
            <Icon>
                <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
            </Icon>
            <hotSpot x='0.5' y='0.5' xunits='fraction' yunits='fraction'/>
        </IconStyle>
        <LineStyle>
            <color>ff000000</color>
            <width>1</width>
        </LineStyle>
        <PolyStyle>
            <color>ff5f1f9f</color>
        </PolyStyle>
    </Style>
    <StyleMap id='m_tiad'>
        <Pair>
            <key>normal</key>
            <styleUrl>#s_tiad</styleUrl>
        </Pair>
        <Pair>
            <key>highlight</key>
            <styleUrl>#s_tiad_hl</styleUrl>
        </Pair>
    </StyleMap>
    <Style id='s_tiad'>
        <IconStyle>
            <color>ff5f1f9f</color>
            <scale>1.1</scale>
            <Icon>
                <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
            </Icon>
            <hotSpot x='0.5' y='0.5' xunits='fraction' yunits='fraction'/>
        </IconStyle>
        <LineStyle>
            <color>ff000000</color>
            <width>1</width>
        </LineStyle>
        <PolyStyle>
            <color>ff5f1f9f</color>
        </PolyStyle>
    </Style>
    <Style id='s_tial_hl'>
        <IconStyle>
            <color>ffac8acd</color>
            <scale>1.3</scale>
            <Icon>
                <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
            </Icon>
            <hotSpot x='0.5' y='0.5' xunits='fraction' yunits='fraction'/>
        </IconStyle>
        <LineStyle>
            <color>ff000000</color>
            <width>1</width>
        </LineStyle>
        <PolyStyle>
            <color>ffac8acd</color>
        </PolyStyle>
    </Style>
    <StyleMap id='m_tial'>
        <Pair>
            <key>normal</key>
            <styleUrl>#s_tial</styleUrl>
        </Pair>
        <Pair>
            <key>highlight</key>
            <styleUrl>#s_tial_hl</styleUrl>
        </Pair>
    </StyleMap>
    <Style id='s_tial'>
        <IconStyle>
            <color>ffac8acd</color>
            <scale>1.1</scale>
            <Icon>
                <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
            </Icon>
            <hotSpot x='0.5' y='0.5' xunits='fraction' yunits='fraction'/>
        </IconStyle>
        <LineStyle>
            <color>ff000000</color>
            <width>1</width>
        </LineStyle>
        <PolyStyle>
            <color>ffac8acd</color>
        </PolyStyle>
    </Style>
    <Style id='l_hollow_hl'>
        <IconStyle>
            <color>ff000000</color>
            <scale>1.3</scale>
            <Icon>
                <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
            </Icon>
            <hotSpot x='0.5' y='0.5' xunits='fraction' yunits='fraction'/>
        </IconStyle>
        <LineStyle>
            <color>ff000000</color>
            <width>1</width>
        </LineStyle>
        <PolyStyle>
            <color>00000000</color>
        </PolyStyle>
    </Style>
    <StyleMap id='ml_hollow'>
        <Pair>
            <key>normal</key>
            <styleUrl>#l_hollow</styleUrl>
        </Pair>
        <Pair>
            <key>highlight</key>
            <styleUrl>#l_hollow_hl</styleUrl>
        </Pair>
    </StyleMap>
    <Style id='l_hollow'>
        <IconStyle>
            <color>ff000000</color>
            <scale>1.1</scale>
            <Icon>
                <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
            </Icon>
            <hotSpot x='0.5' y='0.5' xunits='fraction' yunits='fraction'/>
        </IconStyle>
        <LineStyle>
            <color>ff000000</color>
            <width>1</width>
        </LineStyle>
        <PolyStyle>
            <color>00000000</color>
        </PolyStyle>
    </Style>
    <Style id='l_white_hl'>
        <IconStyle>
            <color>ffffffff</color>
            <scale>1.3</scale>
            <Icon>
                <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
            </Icon>
            <hotSpot x='0.5' y='0.5' xunits='fraction' yunits='fraction'/>
        </IconStyle>
        <LineStyle>
            <color>ffffffff</color>
            <width>1</width>
        </LineStyle>
        <PolyStyle>
            <color>ff000000</color>
        </PolyStyle>
    </Style>
    <StyleMap id='ml_white'>
        <Pair>
            <key>normal</key>
            <styleUrl>#l_white</styleUrl>
        </Pair>
        <Pair>
            <key>highlight</key>
            <styleUrl>#l_white_hl</styleUrl>
        </Pair>
    </StyleMap>
    <Style id='l_white'>
        <IconStyle>
            <color>ffffffff</color>
            <scale>1.1</scale>
            <Icon>
                <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
            </Icon>
            <hotSpot x='0.5' y='0.5' xunits='fraction' yunits='fraction'/>
        </IconStyle>
        <LineStyle>
            <color>ffffffff</color>
            <width>1</width>
        </LineStyle>
        <PolyStyle>
            <color>ff000000</color>
        </PolyStyle>
    </Style>
    <Style id='l_red_hl'>
        <IconStyle>
            <color>ff0000ff</color>
            <scale>1.3</scale>
            <Icon>
                <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
            </Icon>
            <hotSpot x='0.5' y='0.5' xunits='fraction' yunits='fraction'/>
        </IconStyle>
        <LineStyle>
            <color>ff0000ff</color>
            <width>1</width>
        </LineStyle>
        <PolyStyle>
            <color>ffffffff</color>
        </PolyStyle>
    </Style>
    <StyleMap id='ml_red'>
        <Pair>
            <key>normal</key>
            <styleUrl>#l_red</styleUrl>
        </Pair>
        <Pair>
            <key>highlight</key>
            <styleUrl>#l_red_hl</styleUrl>
        </Pair>
    </StyleMap>
    <Style id='l_red'>
        <IconStyle>
            <color>ff0000ff</color>
            <scale>1.1</scale>
            <Icon>
                <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
            </Icon>
            <hotSpot x='0.5' y='0.5' xunits='fraction' yunits='fraction'/>
        </IconStyle>
        <LineStyle>
            <color>ff0000ff</color>
            <width>1</width>
        </LineStyle>
        <PolyStyle>
            <color>ffffffff</color>
        </PolyStyle>
    </Style>
    <Style id='l_black_hl'>
        <IconStyle>
            <color>ff000000</color>
            <scale>1.3</scale>
            <Icon>
                <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
            </Icon>
            <hotSpot x='0.5' y='0.5' xunits='fraction' yunits='fraction'/>
        </IconStyle>
        <LineStyle>
            <color>ff000000</color>
            <width>1</width>
        </LineStyle>
        <PolyStyle>
            <color>ffffffff</color>
        </PolyStyle>
    </Style>
    <StyleMap id='ml_black'>
        <Pair>
            <key>normal</key>
            <styleUrl>#l_black</styleUrl>
        </Pair>
        <Pair>
            <key>highlight</key>
            <styleUrl>#l_black_hl</styleUrl>
        </Pair>
    </StyleMap>
    <Style id='l_black'>
        <IconStyle>
            <color>ff000000</color>
            <scale>1.1</scale>
            <Icon>
                <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
            </Icon>
            <hotSpot x='0.5' y='0.5' xunits='fraction' yunits='fraction'/>
        </IconStyle>
        <LineStyle>
            <color>ff000000</color>
            <width>1</width>
        </LineStyle>
        <PolyStyle>
            <color>ffffffff</color>
        </PolyStyle>
    </Style>
    <Style id='l_green_hl'>
        <IconStyle>
            <color>ff00ff00</color>
            <scale>1.3</scale>
            <Icon>
                <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
            </Icon>
            <hotSpot x='0.5' y='0.5' xunits='fraction' yunits='fraction'/>
        </IconStyle>
        <LineStyle>
            <color>ff00ff00</color>
            <width>1</width>
        </LineStyle>
        <PolyStyle>
            <color>ffffffff</color>
        </PolyStyle>
    </Style>
    <StyleMap id='ml_green'>
        <Pair>
            <key>normal</key>
            <styleUrl>#l_green</styleUrl>
        </Pair>
        <Pair>
            <key>highlight</key>
            <styleUrl>#l_green_hl</styleUrl>
        </Pair>
    </StyleMap>
    <Style id='l_green'>
        <IconStyle>
            <color>ff00ff00</color>
            <scale>1.1</scale>
            <Icon>
                <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
            </Icon>
            <hotSpot x='0.5' y='0.5' xunits='fraction' yunits='fraction'/>
        </IconStyle>
        <LineStyle>
            <color>ff00ff00</color>
            <width>1</width>
        </LineStyle>
        <PolyStyle>
            <color>ffffffff</color>
        </PolyStyle>
    </Style>
    <Style id='l_blue_hl'>
        <IconStyle>
            <color>ffff0000</color>
            <scale>1.3</scale>
            <Icon>
                <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
            </Icon>
            <hotSpot x='0.5' y='0.5' xunits='fraction' yunits='fraction'/>
        </IconStyle>
        <LineStyle>
            <color>ffff0000</color>
            <width>1</width>
        </LineStyle>
        <PolyStyle>
            <color>ffffffff</color>
        </PolyStyle>
    </Style>
    <StyleMap id='ml_blue'>
        <Pair>
            <key>normal</key>
            <styleUrl>#l_blue</styleUrl>
        </Pair>
        <Pair>
            <key>highlight</key>
            <styleUrl>#l_blue_hl</styleUrl>
        </Pair>
    </StyleMap>
    <Style id='l_blue'>
        <IconStyle>
            <color>ffff0000</color>
            <scale>1.1</scale>
            <Icon>
                <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
            </Icon>
            <hotSpot x='0.5' y='0.5' xunits='fraction' yunits='fraction'/>
        </IconStyle>
        <LineStyle>
            <color>ffff0000</color>
            <width>1</width>
        </LineStyle>
        <PolyStyle>
            <color>ffffffff</color>
        </PolyStyle>
    </Style>
    <Style id='l_yellow_hl'>
        <IconStyle>
            <color>ff00ffff</color>
            <scale>1.3</scale>
            <Icon>
                <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
            </Icon>
            <hotSpot x='0.5' y='0.5' xunits='fraction' yunits='fraction'/>
        </IconStyle>
        <LineStyle>
            <color>ff00ffff</color>
            <width>1</width>
        </LineStyle>
        <PolyStyle>
            <color>ffffffff</color>
        </PolyStyle>
    </Style>
    <StyleMap id='ml_yellow'>
        <Pair>
            <key>normal</key>
            <styleUrl>#l_yellow</styleUrl>
        </Pair>
        <Pair>
            <key>highlight</key>
            <styleUrl>#l_yellow_hl</styleUrl>
        </Pair>
    </StyleMap>
    <Style id='l_yellow'>
        <IconStyle>
            <color>ff00ffff</color>
            <scale>1.1</scale>
            <Icon>
                <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
            </Icon>
            <hotSpot x='0.5' y='0.5' xunits='fraction' yunits='fraction'/>
        </IconStyle>
        <LineStyle>
            <color>ff00ffff</color>
            <width>1</width>
        </LineStyle>
        <PolyStyle>
            <color>ffffffff</color>
        </PolyStyle>
    </Style>
    <Style id='l_orange_hl'>
        <IconStyle>
            <color>ff00aaff</color>
            <scale>1.3</scale>
            <Icon>
                <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
            </Icon>
            <hotSpot x='0.5' y='0.5' xunits='fraction' yunits='fraction'/>
        </IconStyle>
        <LineStyle>
            <color>ff00aaff</color>
            <width>1</width>
        </LineStyle>
        <PolyStyle>
            <color>ffffffff</color>
        </PolyStyle>
    </Style>
    <StyleMap id='ml_orange'>
        <Pair>
            <key>normal</key>
            <styleUrl>#l_orange</styleUrl>
        </Pair>
        <Pair>
            <key>highlight</key>
            <styleUrl>#l_orange_hl</styleUrl>
        </Pair>
    </StyleMap>
    <Style id='l_orange'>
        <IconStyle>
            <color>ff00aaff</color>
            <scale>1.1</scale>
            <Icon>
                <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
            </Icon>
            <hotSpot x='0.5' y='0.5' xunits='fraction' yunits='fraction'/>
        </IconStyle>
        <LineStyle>
            <color>ff00aaff</color>
            <width>1</width>
        </LineStyle>
        <PolyStyle>
            <color>ffffffff</color>
        </PolyStyle>
    </Style>
    <Style id='l_tiad_hl'>
        <IconStyle>
            <color>ff5f1f9f</color>
            <scale>1.3</scale>
            <Icon>
                <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
            </Icon>
            <hotSpot x='0.5' y='0.5' xunits='fraction' yunits='fraction'/>
        </IconStyle>
        <LineStyle>
            <color>ff5f1f9f</color>
            <width>1</width>
        </LineStyle>
        <PolyStyle>
            <color>ffffffff</color>
        </PolyStyle>
    </Style>
    <StyleMap id='ml_tiad'>
        <Pair>
            <key>normal</key>
            <styleUrl>#l_tiad</styleUrl>
        </Pair>
        <Pair>
            <key>highlight</key>
            <styleUrl>#l_tiad_hl</styleUrl>
        </Pair>
    </StyleMap>
    <Style id='l_tiad'>
        <IconStyle>
            <color>ff5f1f9f</color>
            <scale>1.1</scale>
            <Icon>
                <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
            </Icon>
            <hotSpot x='0.5' y='0.5' xunits='fraction' yunits='fraction'/>
        </IconStyle>
        <LineStyle>
            <color>ff5f1f9f</color>
            <width>1</width>
        </LineStyle>
        <PolyStyle>
            <color>ffffffff</color>
        </PolyStyle>
    </Style>
    <Style id='l_tial_hl'>
        <IconStyle>
            <color>ffac8acd</color>
            <scale>1.3</scale>
            <Icon>
                <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
            </Icon>
            <hotSpot x='0.5' y='0.5' xunits='fraction' yunits='fraction'/>
        </IconStyle>
        <LineStyle>
            <color>ffac8acd</color>
            <width>1</width>
        </LineStyle>
        <PolyStyle>
            <color>ffffffff</color>
        </PolyStyle>
    </Style>
    <StyleMap id='ml_tial'>
        <Pair>
            <key>normal</key>
            <styleUrl>#l_tial</styleUrl>
        </Pair>
        <Pair>
            <key>highlight</key>
            <styleUrl>#l_tial_hl</styleUrl>
        </Pair>
    </StyleMap>
    <Style id='l_tial'>
        <IconStyle>
            <color>ffac8acd</color>
            <scale>1.1</scale>
            <Icon>
                <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
            </Icon>
            <hotSpot x='0.5' y='0.5' xunits='fraction' yunits='fraction'/>
        </IconStyle>
        <LineStyle>
            <color>ffac8acd</color>
            <width>1</width>
        </LineStyle>
        <PolyStyle>
            <color>ffffffff</color>
        </PolyStyle>
    </Style>";




        private void button3_Click(object sender, EventArgs e)
        {


            openFileDialog2.DefaultExt = "kml";
            openFileDialog2.Filter = "KML File|*.kml";
            DialogResult result = openFileDialog2.ShowDialog();
            string File_Name = openFileDialog2.SafeFileName.ToString();
            string file = openFileDialog2.FileName;


            XmlTextWriter textWriter = new XmlTextWriter(file, null);


            if (KML_Type == "Subcontractor_Project")
            {


                Contractor FaraCell;
                Contractor Figetel;
                Contractor Vihan;
                Contractor Mahna;
                Contractor PCOM;
                Contractor EMCI;
                Contractor Tebyan;
                Contractor Nak_Team;
                Contractor Farhikhtegan;
                Contractor Nak_ICT;



                FaraCell.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                FaraCell.KMLText = new string[] { "", "", "", "", "", "", "" };

                Figetel.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                Figetel.KMLText = new string[] { "", "", "", "", "", "", "" };

                Vihan.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                Vihan.KMLText = new string[] { "", "", "", "", "", "", "" };

                Mahna.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                Mahna.KMLText = new string[] { "", "", "", "", "", "", "" };

                PCOM.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                PCOM.KMLText = new string[] { "", "", "", "", "", "", "" };

                EMCI.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                EMCI.KMLText = new string[] { "", "", "", "", "", "", "" };

                Tebyan.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                Tebyan.KMLText = new string[] { "", "", "", "", "", "", "" };

                Nak_Team.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                Nak_Team.KMLText = new string[] { "", "", "", "", "", "", "" };

                Farhikhtegan.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                Farhikhtegan.KMLText = new string[] { "", "", "", "", "", "", "" };

                Nak_ICT.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                Nak_ICT.KMLText = new string[] { "", "", "", "", "", "", "" };



                for (int k = 0; k < SSV_KMLTable.Rows.Count; k++)
                {
                    string Location = SSV_KMLTable.Rows[k].ItemArray[0].ToString();
                    string SiteCode_Technology = SSV_KMLTable.Rows[k].ItemArray[3].ToString();
                    string LAT = SSV_KMLTable.Rows[k].ItemArray[1].ToString();
                    string LON = SSV_KMLTable.Rows[k].ItemArray[2].ToString();
                    string Project = SSV_KMLTable.Rows[k].ItemArray[5].ToString();
                    string Vendor = SSV_KMLTable.Rows[k].ItemArray[6].ToString();
                    string City = SSV_KMLTable.Rows[k].ItemArray[7].ToString();
                    string Assignment_to_SC_Date = SSV_KMLTable.Rows[k].ItemArray[12].ToString();
                    string Contractor = SSV_KMLTable.Rows[k].ItemArray[13].ToString();
                    string Periority = SSV_KMLTable.Rows[k].ItemArray[15].ToString();
                    string Submit_to_OPT_Team_Date = SSV_KMLTable.Rows[k].ItemArray[14].ToString();
                    string Comment_External = SSV_KMLTable.Rows[k].ItemArray[16].ToString();



                    string Point_Periority = "";
                    if (Periority == "Priority 0")
                    {
                        Point_Periority = "P0";
                    }
                    if (Periority == "Priority 1")
                    {
                        Point_Periority = "P1";
                    }
                    if (Periority == "Priority 2")
                    {
                        Point_Periority = "P2";
                    }
                    if (Periority == "Priority 3")
                    {
                        Point_Periority = "P3";
                    }
                    if (Periority == "Priority 4")
                    {
                        Point_Periority = "P4";
                    }


                    string Point_Color = "";
                    string Point_Name = "";
                    if (Submit_to_OPT_Team_Date != "")
                    {
                        Point_Color = "#m_green";
                        Point_Name = Location;
                    }
                    else
                    { 
                        if (Point_Periority == "P0")
                        //if ((Project == "PH8" || Project == "PH9" || Project == "TDD") && Point_Periority == "P0")
                        {
                            Point_Color = "#m_tiad";
                            Point_Name = "P0";
                        }
                        if ((Project == "PH8" || Project == "PH9" || Project == "TDD") && Point_Periority == "P1")
                        {
                            Point_Color = "#m_red";
                            Point_Name = "P1";
                        }
                        if (Project == "5G" && Point_Periority == "P1")
                        {
                            Point_Color = "#m_tiad";
                            Point_Name = "P1";
                        }
                        if ((Project == "PH8" || Project == "PH9" || Project == "TDD") && Point_Periority == "P2")
                        {
                            Point_Color = "#m_orange";
                            Point_Name = "P2";
                        }
                        if ((Project == "PH8" || Project == "PH9" || Project == "TDD") && Point_Periority == "P3")
                        {
                            Point_Color = "#m_yellow";
                            Point_Name = "P3";
                        }
                        if ((Project == "PH8" || Project == "PH9" || Project == "TDD") && Point_Periority == "P4")
                        {
                            Point_Color = "#m_white";
                            Point_Name = "P4";
                        }
                        if (Project == "Tower Change")
                        {
                            Point_Color = "#m_tial";
                            Point_Name = Location;
                        }
                        if (Comment_External != "")
                        {
                            Point_Color = "#m_blue";
                            Point_Name = Location;
                        }
                        if (Project == "USO" || Project == "USO-ReDrive")
                        {
                            Point_Color = "#m_black";
                            Point_Name = Location;
                        }
                    }


                    if (Contractor == "FaraCell")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        FaraCell.KMLText[Project_index] = FaraCell.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }




                    if (Contractor == "Figetel")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        Figetel.KMLText[Project_index] = Figetel.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }



                    if (Contractor == "Vihan")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        Vihan.KMLText[Project_index] = Vihan.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }




                    if (Contractor == "Mahna")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        Mahna.KMLText[Project_index] = Mahna.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }





                    if (Contractor == "PCOM")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        PCOM.KMLText[Project_index] = PCOM.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }




                    if (Contractor == "EMCI")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        EMCI.KMLText[Project_index] = EMCI.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }






                    if (Contractor == "Tebyan")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        Tebyan.KMLText[Project_index] = Tebyan.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }





                    if (Contractor == "Nak_Team")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        Nak_Team.KMLText[Project_index] = Nak_Team.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }






                    if (Contractor == "Farhikhtegan")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        Farhikhtegan.KMLText[Project_index] = Farhikhtegan.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }





                    if (Contractor == "Nak_ICT")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        Nak_ICT.KMLText[Project_index] = Nak_ICT.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }





                }


                string FaraCell_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    FaraCell_KMLText = FaraCell_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + FaraCell.KMLText[j] +
          @"</Folder>";

                }


                FaraCell_KMLText = @"
            <Folder>
            <name> " + "FaraCell" + @" </name>"
    + FaraCell_KMLText +
    @" </Folder>";





                string Figetel_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    Figetel_KMLText = Figetel_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + Figetel.KMLText[j] +
          @"</Folder>";

                }


                Figetel_KMLText = @"
            <Folder>
            <name> " + "Figetel" + @" </name>"
    + Figetel_KMLText +
    @" </Folder>";







                string Vihan_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    Vihan_KMLText = Vihan_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + Vihan.KMLText[j] +
          @"</Folder>";

                }


                Vihan_KMLText = @"
            <Folder>
            <name> " + "Vihan" + @" </name>"
    + Vihan_KMLText +
    @" </Folder>";




                string Mahna_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    Mahna_KMLText = Mahna_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + Mahna.KMLText[j] +
          @"</Folder>";

                }


                Mahna_KMLText = @"
            <Folder>
            <name> " + "Mahna" + @" </name>"
    + Mahna_KMLText +
    @" </Folder>";





                string PCOM_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    PCOM_KMLText = PCOM_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + PCOM.KMLText[j] +
          @"</Folder>";

                }


                PCOM_KMLText = @"
            <Folder>
            <name> " + "PCOM" + @" </name>"
    + PCOM_KMLText +
    @" </Folder>";






                string EMCI_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    EMCI_KMLText = EMCI_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + EMCI.KMLText[j] +
          @"</Folder>";

                }


                EMCI_KMLText = @"
            <Folder>
            <name> " + "EMCI" + @" </name>"
    + EMCI_KMLText +
    @" </Folder>";



                string Tebyan_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    Tebyan_KMLText = Tebyan_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + Tebyan.KMLText[j] +
          @"</Folder>";

                }


                Tebyan_KMLText = @"
            <Folder>
            <name> " + "Tebyan" + @" </name>"
    + Tebyan_KMLText +
    @" </Folder>";



                string Nak_Team_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    Nak_Team_KMLText = Nak_Team_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + Nak_Team.KMLText[j] +
          @"</Folder>";

                }


                Nak_Team_KMLText = @"
            <Folder>
            <name> " + "Nak_Team" + @" </name>"
    + Nak_Team_KMLText +
    @" </Folder>";






                string Farhikhtegan_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    Farhikhtegan_KMLText = Farhikhtegan_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + Farhikhtegan.KMLText[j] +
          @"</Folder>";

                }


                Farhikhtegan_KMLText = @"
            <Folder>
            <name> " + "Farhikhtegan" + @" </name>"
    + Farhikhtegan_KMLText +
    @" </Folder>";




                string Nak_ICT_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    Nak_ICT_KMLText = Nak_ICT_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + Nak_ICT.KMLText[j] +
          @"</Folder>";

                }


                Nak_ICT_KMLText = @"
            <Folder>
            <name> " + "Nak_ICT" + @" </name>"
    + Nak_ICT_KMLText +
    @" </Folder>";





                textWriter.WriteStartElement(Fixed_part + FaraCell_KMLText + Figetel_KMLText + Vihan_KMLText + Mahna_KMLText + PCOM_KMLText + EMCI_KMLText + Tebyan_KMLText + Nak_Team_KMLText + Farhikhtegan_KMLText + Nak_ICT_KMLText + @"</Document>" +
                                                                               @"</kml>");


                textWriter.Close();

                string contents = File.ReadAllText(file);
                string output = contents.Substring(1, contents.Length - 3);
                File.WriteAllText(file, output);






                MessageBox.Show("Finished");




            }




            if (KML_Type == "Periority_Subcontractor_Project")
            {



                Periority_Contractor P0_FaraCell;
                Periority_Contractor P0_Figetel;
                Periority_Contractor P0_Vihan;
                Periority_Contractor P0_Mahna;
                Periority_Contractor P0_PCOM;
                Periority_Contractor P0_EMCI;
                Periority_Contractor P0_Tebyan;
                Periority_Contractor P0_Nak_Team;
                Periority_Contractor P0_Farhikhtegan;
                Periority_Contractor P0_Nak_ICT;


                Periority_Contractor P1_FaraCell;
                Periority_Contractor P1_Figetel;
                Periority_Contractor P1_Vihan;
                Periority_Contractor P1_Mahna;
                Periority_Contractor P1_PCOM;
                Periority_Contractor P1_EMCI;
                Periority_Contractor P1_Tebyan;
                Periority_Contractor P1_Nak_Team;
                Periority_Contractor P1_Farhikhtegan;
                Periority_Contractor P1_Nak_ICT;

                Periority_Contractor P2_FaraCell;
                Periority_Contractor P2_Figetel;
                Periority_Contractor P2_Vihan;
                Periority_Contractor P2_Mahna;
                Periority_Contractor P2_PCOM;
                Periority_Contractor P2_EMCI;
                Periority_Contractor P2_Tebyan;
                Periority_Contractor P2_Nak_Team;
                Periority_Contractor P2_Farhikhtegan;
                Periority_Contractor P2_Nak_ICT;


                Periority_Contractor P3_FaraCell;
                Periority_Contractor P3_Figetel;
                Periority_Contractor P3_Vihan;
                Periority_Contractor P3_Mahna;
                Periority_Contractor P3_PCOM;
                Periority_Contractor P3_EMCI;
                Periority_Contractor P3_Tebyan;
                Periority_Contractor P3_Nak_Team;
                Periority_Contractor P3_Farhikhtegan;
                Periority_Contractor P3_Nak_ICT;


                Periority_Contractor P4_FaraCell;
                Periority_Contractor P4_Figetel;
                Periority_Contractor P4_Vihan;
                Periority_Contractor P4_Mahna;
                Periority_Contractor P4_PCOM;
                Periority_Contractor P4_EMCI;
                Periority_Contractor P4_Tebyan;
                Periority_Contractor P4_Nak_Team;
                Periority_Contractor P4_Farhikhtegan;
                Periority_Contractor P4_Nak_ICT;




                P0_FaraCell.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P0_FaraCell.KMLText = new string[] { "", "", "", "", "", "", "" };

                P0_Figetel.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P0_Figetel.KMLText = new string[] { "", "", "", "", "", "", "" };

                P0_Vihan.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P0_Vihan.KMLText = new string[] { "", "", "", "", "", "", "" };

                P0_Mahna.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P0_Mahna.KMLText = new string[] { "", "", "", "", "", "", "" };

                P0_PCOM.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P0_PCOM.KMLText = new string[] { "", "", "", "", "", "", "" };

                P0_EMCI.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P0_EMCI.KMLText = new string[] { "", "", "", "", "", "", "" };

                P0_Tebyan.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P0_Tebyan.KMLText = new string[] { "", "", "", "", "", "", "" };

                P0_Nak_Team.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P0_Nak_Team.KMLText = new string[] { "", "", "", "", "", "", "" };

                P0_Farhikhtegan.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P0_Farhikhtegan.KMLText = new string[] { "", "", "", "", "", "", "" };

                P0_Nak_ICT.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P0_Nak_ICT.KMLText = new string[] { "", "", "", "", "", "", "" };





                P1_FaraCell.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P1_FaraCell.KMLText = new string[] { "", "", "", "", "", "", "" };

                P1_Figetel.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P1_Figetel.KMLText = new string[] { "", "", "", "", "", "", "" };

                P1_Vihan.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P1_Vihan.KMLText = new string[] { "", "", "", "", "", "", "" };

                P1_Mahna.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P1_Mahna.KMLText = new string[] { "", "", "", "", "", "", "" };

                P1_PCOM.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P1_PCOM.KMLText = new string[] { "", "", "", "", "", "", "" };

                P1_EMCI.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P1_EMCI.KMLText = new string[] { "", "", "", "", "", "", "" };

                P1_Tebyan.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P1_Tebyan.KMLText = new string[] { "", "", "", "", "", "", "" };

                P1_Nak_Team.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P1_Nak_Team.KMLText = new string[] { "", "", "", "", "", "", "" };

                P1_Farhikhtegan.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P1_Farhikhtegan.KMLText = new string[] { "", "", "", "", "", "", "" };

                P1_Nak_ICT.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P1_Nak_ICT.KMLText = new string[] { "", "", "", "", "", "", "" };




                P2_FaraCell.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P2_FaraCell.KMLText = new string[] { "", "", "", "", "", "", "" };

                P2_Figetel.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P2_Figetel.KMLText = new string[] { "", "", "", "", "", "", "" };

                P2_Vihan.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P2_Vihan.KMLText = new string[] { "", "", "", "", "", "", "" };

                P2_Mahna.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P2_Mahna.KMLText = new string[] { "", "", "", "", "", "", "" };

                P2_PCOM.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P2_PCOM.KMLText = new string[] { "", "", "", "", "", "", "" };

                P2_EMCI.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P2_EMCI.KMLText = new string[] { "", "", "", "", "", "", "" };

                P2_Tebyan.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P2_Tebyan.KMLText = new string[] { "", "", "", "", "", "", "" };

                P2_Nak_Team.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P2_Nak_Team.KMLText = new string[] { "", "", "", "", "", "", "" };

                P2_Farhikhtegan.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P2_Farhikhtegan.KMLText = new string[] { "", "", "", "", "", "", "" };

                P2_Nak_ICT.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P2_Nak_ICT.KMLText = new string[] { "", "", "", "", "", "", "" };



                P3_FaraCell.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P3_FaraCell.KMLText = new string[] { "", "", "", "", "", "", "" };

                P3_Figetel.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P3_Figetel.KMLText = new string[] { "", "", "", "", "", "", "" };

                P3_Vihan.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P3_Vihan.KMLText = new string[] { "", "", "", "", "", "", "" };

                P3_Mahna.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P3_Mahna.KMLText = new string[] { "", "", "", "", "", "", "" };

                P3_PCOM.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P3_PCOM.KMLText = new string[] { "", "", "", "", "", "", "" };

                P3_EMCI.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P3_EMCI.KMLText = new string[] { "", "", "", "", "", "", "" };

                P3_Tebyan.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P3_Tebyan.KMLText = new string[] { "", "", "", "", "", "", "" };

                P3_Nak_Team.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P3_Nak_Team.KMLText = new string[] { "", "", "", "", "", "", "" };

                P3_Farhikhtegan.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P3_Farhikhtegan.KMLText = new string[] { "", "", "", "", "", "", "" };

                P3_Nak_ICT.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P3_Nak_ICT.KMLText = new string[] { "", "", "", "", "", "", "" };




                P4_FaraCell.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P4_FaraCell.KMLText = new string[] { "", "", "", "", "", "", "" };

                P4_Figetel.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P4_Figetel.KMLText = new string[] { "", "", "", "", "", "", "" };

                P4_Vihan.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P4_Vihan.KMLText = new string[] { "", "", "", "", "", "", "" };

                P4_Mahna.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P4_Mahna.KMLText = new string[] { "", "", "", "", "", "", "" };

                P4_PCOM.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P4_PCOM.KMLText = new string[] { "", "", "", "", "", "", "" };

                P4_EMCI.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P4_EMCI.KMLText = new string[] { "", "", "", "", "", "", "" };

                P4_Tebyan.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P4_Tebyan.KMLText = new string[] { "", "", "", "", "", "", "" };

                P4_Nak_Team.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P4_Nak_Team.KMLText = new string[] { "", "", "", "", "", "", "" };

                P4_Farhikhtegan.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P4_Farhikhtegan.KMLText = new string[] { "", "", "", "", "", "", "" };

                P4_Nak_ICT.Projects = new string[] { "PH8", "PH9", "TDD", "USO-ReDrive", "USO", "Tower Change", "5G" };
                P4_Nak_ICT.KMLText = new string[] { "", "", "", "", "", "", "" };



                for (int k = 0; k < SSV_KMLTable.Rows.Count; k++)
                {
                    string Location = SSV_KMLTable.Rows[k].ItemArray[0].ToString();
                    string SiteCode_Technology = SSV_KMLTable.Rows[k].ItemArray[3].ToString();
                    string LAT = SSV_KMLTable.Rows[k].ItemArray[1].ToString();
                    string LON = SSV_KMLTable.Rows[k].ItemArray[2].ToString();
                    string Project = SSV_KMLTable.Rows[k].ItemArray[5].ToString();
                    string Vendor = SSV_KMLTable.Rows[k].ItemArray[6].ToString();
                    string City = SSV_KMLTable.Rows[k].ItemArray[7].ToString();
                    string Assignment_to_SC_Date = SSV_KMLTable.Rows[k].ItemArray[12].ToString();
                    string Contractor = SSV_KMLTable.Rows[k].ItemArray[13].ToString();
                    string Periority = SSV_KMLTable.Rows[k].ItemArray[15].ToString();
                    string Submit_to_OPT_Team_Date = SSV_KMLTable.Rows[k].ItemArray[14].ToString();
                    string Comment_External = SSV_KMLTable.Rows[k].ItemArray[16].ToString();



                    string Point_Periority = "";
                    if (Periority == "Priority 0")
                    {
                        Point_Periority = "P0";
                    }
                    if (Periority == "Priority 1")
                    {
                        Point_Periority = "P1";
                    }
                    if (Periority == "Priority 2")
                    {
                        Point_Periority = "P2";
                    }
                    if (Periority == "Priority 3")
                    {
                        Point_Periority = "P3";
                    }
                    if (Periority == "Priority 4")
                    {
                        Point_Periority = "P4";
                    }


                    string Point_Color = "";
                    string Point_Name = "";
                    if (Submit_to_OPT_Team_Date != "")
                    {
                        Point_Color = "#m_green";
                        Point_Name = Location;
                    }
                    else
                    {
                        if (Point_Periority == "P0")
                        //if ((Project == "PH8" || Project == "PH9" || Project == "TDD") && Point_Periority == "P0")
                        {
                            Point_Color = "#m_tiad";
                            Point_Name = "P0";
                        }
                        if ((Project == "PH8" || Project == "PH9" || Project == "TDD") && Point_Periority == "P1")
                        {
                            Point_Color = "#m_red";
                            Point_Name = "P1";
                        }
                        if (Project == "5G" && Point_Periority == "P1")
                        {
                            Point_Color = "#m_tiad";
                            Point_Name = "P1";
                        }
                        if ((Project == "PH8" || Project == "PH9" || Project == "TDD") && Point_Periority == "P2")
                        {
                            Point_Color = "#m_orange";
                            Point_Name = "P2";
                        }
                        if ((Project == "PH8" || Project == "PH9" || Project == "TDD") && Point_Periority == "P3")
                        {
                            Point_Color = "#m_yellow";
                            Point_Name = "P3";
                        }
                        if ((Project == "PH8" || Project == "PH9" || Project == "TDD") && Point_Periority == "P4")
                        {
                            Point_Color = "#m_white";
                            Point_Name = "P4";
                        }
                        if (Project == "Tower Change")
                        {
                            Point_Color = "#m_tial";
                            Point_Name = Location;
                        }
                        if (Comment_External != "")
                        {
                            Point_Color = "#m_blue";
                            Point_Name = Location;
                        }
                        if (Project == "USO" || Project == "USO-ReDrive")
                        {
                            Point_Color = "#m_black";
                            Point_Name = Location;
                        }
                    }






                    if (Contractor == "FaraCell" && Periority == "Priority 0")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P0_FaraCell.KMLText[Project_index] = P0_FaraCell.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }







                    if (Contractor == "FaraCell" && Periority == "Priority 1")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P1_FaraCell.KMLText[Project_index] = P1_FaraCell.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }




                    if (Contractor == "FaraCell" && Periority == "Priority 2")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P2_FaraCell.KMLText[Project_index] = P2_FaraCell.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }


                    if (Contractor == "FaraCell" && Periority == "Priority 3")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P3_FaraCell.KMLText[Project_index] = P3_FaraCell.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }







                    if (Contractor == "FaraCell" && Periority == "Priority 4")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P4_FaraCell.KMLText[Project_index] = P4_FaraCell.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }






                    if (Contractor == "Figetel" && Periority == "Priority 0")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P0_Figetel.KMLText[Project_index] = P0_Figetel.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }






                    if (Contractor == "Figetel" && Periority == "Priority 1")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P1_Figetel.KMLText[Project_index] = P1_Figetel.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }




                    if (Contractor == "Figetel" && Periority == "Priority 2")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P2_Figetel.KMLText[Project_index] = P2_Figetel.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }


                    if (Contractor == "Figetel" && Periority == "Priority 3")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P3_Figetel.KMLText[Project_index] = P3_Figetel.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }







                    if (Contractor == "Figetel" && Periority == "Priority 4")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P4_Figetel.KMLText[Project_index] = P4_Figetel.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }







                    if (Contractor == "Vihan" && Periority == "Priority 0")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P0_Vihan.KMLText[Project_index] = P0_Vihan.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }







                    if (Contractor == "Vihan" && Periority == "Priority 1")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P1_Vihan.KMLText[Project_index] = P1_Vihan.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }




                    if (Contractor == "Vihan" && Periority == "Priority 2")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P2_Vihan.KMLText[Project_index] = P2_Vihan.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }


                    if (Contractor == "Vihan" && Periority == "Priority 3")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P3_Vihan.KMLText[Project_index] = P3_Vihan.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }







                    if (Contractor == "Vihan" && Periority == "Priority 4")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P4_Vihan.KMLText[Project_index] = P4_Vihan.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }






                    if (Contractor == "Mahna" && Periority == "Priority 0")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P0_Mahna.KMLText[Project_index] = P0_Mahna.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }





                    if (Contractor == "Mahna" && Periority == "Priority 1")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P1_Mahna.KMLText[Project_index] = P1_Mahna.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }




                    if (Contractor == "Mahna" && Periority == "Priority 2")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P2_Mahna.KMLText[Project_index] = P2_Mahna.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }


                    if (Contractor == "Mahna" && Periority == "Priority 3")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P3_Mahna.KMLText[Project_index] = P3_Mahna.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }







                    if (Contractor == "Mahna" && Periority == "Priority 4")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P4_Mahna.KMLText[Project_index] = P4_Mahna.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }



                    if (Contractor == "PCOM" && Periority == "Priority 0")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P0_PCOM.KMLText[Project_index] = P0_PCOM.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }






                    if (Contractor == "PCOM" && Periority == "Priority 1")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P1_PCOM.KMLText[Project_index] = P1_PCOM.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }




                    if (Contractor == "PCOM" && Periority == "Priority 2")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P2_PCOM.KMLText[Project_index] = P2_PCOM.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }


                    if (Contractor == "PCOM" && Periority == "Priority 3")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P3_PCOM.KMLText[Project_index] = P3_PCOM.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }







                    if (Contractor == "PCOM" && Periority == "Priority 4")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P4_PCOM.KMLText[Project_index] = P4_PCOM.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }





                    if (Contractor == "EMCI" && Periority == "Priority 0")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P0_EMCI.KMLText[Project_index] = P0_EMCI.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }





                    if (Contractor == "EMCI" && Periority == "Priority 1")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P1_EMCI.KMLText[Project_index] = P1_EMCI.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }




                    if (Contractor == "EMCI" && Periority == "Priority 2")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P2_EMCI.KMLText[Project_index] = P2_EMCI.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }


                    if (Contractor == "EMCI" && Periority == "Priority 3")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P3_EMCI.KMLText[Project_index] = P3_EMCI.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }







                    if (Contractor == "EMCI" && Periority == "Priority 4")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P4_EMCI.KMLText[Project_index] = P4_EMCI.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }





                    if (Contractor == "Tebyan" && Periority == "Priority 0")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P0_Tebyan.KMLText[Project_index] = P0_Tebyan.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }





                    if (Contractor == "Tebyan" && Periority == "Priority 1")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P1_Tebyan.KMLText[Project_index] = P1_Tebyan.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }




                    if (Contractor == "Tebyan" && Periority == "Priority 2")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P2_Tebyan.KMLText[Project_index] = P2_Tebyan.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }


                    if (Contractor == "Tebyan" && Periority == "Priority 3")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P3_Tebyan.KMLText[Project_index] = P3_Tebyan.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }







                    if (Contractor == "Tebyan" && Periority == "Priority 4")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P4_Tebyan.KMLText[Project_index] = P4_Tebyan.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }




                    if (Contractor == "Nak_Team" && Periority == "Priority 0")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P0_Nak_Team.KMLText[Project_index] = P0_Nak_Team.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }





                    if (Contractor == "Nak_Team" && Periority == "Priority 1")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P1_Nak_Team.KMLText[Project_index] = P1_Nak_Team.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }









                    if (Contractor == "Nak_Team" && Periority == "Priority 2")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P2_Nak_Team.KMLText[Project_index] = P2_Nak_Team.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }


                    if (Contractor == "Nak_Team" && Periority == "Priority 3")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P3_Nak_Team.KMLText[Project_index] = P3_Nak_Team.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }







                    if (Contractor == "Nak_Team" && Periority == "Priority 4")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P4_Nak_Team.KMLText[Project_index] = P4_Nak_Team.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }




                    if (Contractor == "Farhikhtegan" && Periority == "Priority 0")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P0_Farhikhtegan.KMLText[Project_index] = P0_Farhikhtegan.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }





                    if (Contractor == "Farhikhtegan" && Periority == "Priority 1")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P1_Farhikhtegan.KMLText[Project_index] = P1_Farhikhtegan.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }




                    if (Contractor == "Farhikhtegan" && Periority == "Priority 2")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P2_Farhikhtegan.KMLText[Project_index] = P2_Farhikhtegan.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }


                    if (Contractor == "Farhikhtegan" && Periority == "Priority 3")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P3_Farhikhtegan.KMLText[Project_index] = P3_Farhikhtegan.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }







                    if (Contractor == "Farhikhtegan" && Periority == "Priority 4")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P4_Farhikhtegan.KMLText[Project_index] = P4_Farhikhtegan.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }






                    if (Contractor == "Nak_ICT" && Periority == "Priority 0")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P0_Nak_ICT.KMLText[Project_index] = P0_Nak_ICT.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }




                    if (Contractor == "Nak_ICT" && Periority == "Priority 1")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P1_Nak_ICT.KMLText[Project_index] = P1_Nak_ICT.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }




                    if (Contractor == "Nak_ICT" && Periority == "Priority 2")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P2_Nak_ICT.KMLText[Project_index] = P2_Nak_ICT.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }


                    if (Contractor == "Nak_ICT" && Periority == "Priority 3")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P3_Nak_ICT.KMLText[Project_index] = P3_Nak_ICT.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }







                    if (Contractor == "Nak_ICT" && Periority == "Priority 4")
                    {
                        int Project_index = 0;
                        if (Project == "PH8")
                        {
                            Project_index = 0;
                        }
                        if (Project == "PH9")
                        {
                            Project_index = 1;
                        }
                        if (Project == "TDD")
                        {
                            Project_index = 2;
                        }
                        if (Project == "USO-ReDrive")
                        {
                            Project_index = 3;
                        }
                        if (Project == "USO")
                        {
                            Project_index = 4;
                        }
                        if (Project == "Tower Change")
                        {
                            Project_index = 5;
                        }
                        if (Project == "5G")
                        {
                            Project_index = 6;
                        }
                        P4_Nak_ICT.KMLText[Project_index] = P4_Nak_ICT.KMLText[Project_index] + @"
                         <open>1</open>
                              <gx:balloonVisibility>1</gx:balloonVisibility>
                                   <Placemark>
                                      <name>" + Point_Name + @"</name>     
                                           <description><![CDATA[
                                  <b>" + SiteCode_Technology + @"</b>
                                  <hr>
                                  Project = " + Project + @" <br>
                                  Priority = " + Periority + @" <br>
                                  Vendor =" + Vendor + @" <br>
                                  City =" + City + @" <br>
                                  Location =" + Location + @" <br>
                                  Site Name =" + SiteCode_Technology + @"<br>
                                  GPS Lat =" + LAT + @" <br>
                                  GPS Lon =" + LON + @" <br>
                                  Subcon =" + Contractor + @" <br>
                                  Assignment to SC Date = " + Assignment_to_SC_Date + @" <br>
                                  Submit to OPT Team Date = " + Submit_to_OPT_Team_Date + @" <br>
                                  Field Issie External = " + Comment_External + @" <br>
                                  <hr>]]>
                                      </description>
                                                <LookAt>
                                                   <longitude> " + LON + @" </longitude>
                                                        <latitude>" + LAT + @"</latitude>
                                                             <altitude> 0 </altitude>                                              
                                                                  <heading> -1.539914092246387e-008 </heading>
                                                                       <tilt> 0 </tilt>
                                                                            <range> 640383.0131348133 </range>
                                                                                <gx:altitudeMode> relativeToSeaFloor </gx:altitudeMode>
                                                                                  </LookAt>
                                                                                     <styleUrl>" + Point_Color + @"</styleUrl>
                <Point>
                   <gx:drawOrder>1</gx:drawOrder>
                        <coordinates>" + LON + "," + LAT + @",0</coordinates>
                           </Point>
                          </Placemark>";
                    }








                }



                string P0_FaraCell_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P0_FaraCell_KMLText = P0_FaraCell_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P0_FaraCell.KMLText[j] +
          @"</Folder>";

                }



                string P1_FaraCell_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P1_FaraCell_KMLText = P1_FaraCell_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P1_FaraCell.KMLText[j] +
          @"</Folder>";

                }




                string P2_FaraCell_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P2_FaraCell_KMLText = P2_FaraCell_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P2_FaraCell.KMLText[j] +
          @"</Folder>";

                }





                string P3_FaraCell_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P3_FaraCell_KMLText = P3_FaraCell_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P3_FaraCell.KMLText[j] +
          @"</Folder>";

                }






                string P4_FaraCell_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P4_FaraCell_KMLText = P4_FaraCell_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P4_FaraCell.KMLText[j] +
          @"</Folder>";

                }



                string P0_Figetel_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P0_Figetel_KMLText = P0_Figetel_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P0_Figetel.KMLText[j] +
          @"</Folder>";

                }



                string P1_Figetel_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P1_Figetel_KMLText = P1_Figetel_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P1_Figetel.KMLText[j] +
          @"</Folder>";

                }



                string P2_Figetel_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P2_Figetel_KMLText = P2_Figetel_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P2_Figetel.KMLText[j] +
          @"</Folder>";

                }





                string P3_Figetel_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P3_Figetel_KMLText = P3_Figetel_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P3_Figetel.KMLText[j] +
          @"</Folder>";

                }




                string P4_Figetel_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P4_Figetel_KMLText = P4_Figetel_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P4_Figetel.KMLText[j] +
          @"</Folder>";

                }


                string P0_Vihan_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P0_Vihan_KMLText = P0_Vihan_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P0_Vihan.KMLText[j] +
          @"</Folder>";

                }



                string P1_Vihan_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P1_Vihan_KMLText = P1_Vihan_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P1_Vihan.KMLText[j] +
          @"</Folder>";

                }



                string P2_Vihan_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P2_Vihan_KMLText = P2_Vihan_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P2_Vihan.KMLText[j] +
          @"</Folder>";

                }





                string P3_Vihan_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P3_Vihan_KMLText = P3_Vihan_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P3_Vihan.KMLText[j] +
          @"</Folder>";

                }




                string P4_Vihan_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P4_Vihan_KMLText = P4_Vihan_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P4_Vihan.KMLText[j] +
          @"</Folder>";

                }


                string P0_Mahna_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P0_Mahna_KMLText = P0_Mahna_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P0_Mahna.KMLText[j] +
          @"</Folder>";

                }




                string P1_Mahna_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P1_Mahna_KMLText = P1_Mahna_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P1_Mahna.KMLText[j] +
          @"</Folder>";

                }



                string P2_Mahna_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P2_Mahna_KMLText = P2_Mahna_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P2_Mahna.KMLText[j] +
          @"</Folder>";

                }





                string P3_Mahna_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P3_Mahna_KMLText = P3_Mahna_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P3_Mahna.KMLText[j] +
          @"</Folder>";

                }




                string P4_Mahna_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P4_Mahna_KMLText = P4_Mahna_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P4_Mahna.KMLText[j] +
          @"</Folder>";

                }


                string P0_PCOM_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P0_PCOM_KMLText = P0_PCOM_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P0_PCOM.KMLText[j] +
          @"</Folder>";

                }





                string P1_PCOM_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P1_PCOM_KMLText = P1_PCOM_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P1_PCOM.KMLText[j] +
          @"</Folder>";

                }



                string P2_PCOM_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P2_PCOM_KMLText = P2_PCOM_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P2_PCOM.KMLText[j] +
          @"</Folder>";

                }





                string P3_PCOM_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P3_PCOM_KMLText = P3_PCOM_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P3_PCOM.KMLText[j] +
          @"</Folder>";

                }




                string P4_PCOM_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P4_PCOM_KMLText = P4_PCOM_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P4_PCOM.KMLText[j] +
          @"</Folder>";

                }



                string P0_EMCI_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P0_EMCI_KMLText = P0_EMCI_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P0_EMCI.KMLText[j] +
          @"</Folder>";

                }




                string P1_EMCI_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P1_EMCI_KMLText = P1_EMCI_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P1_EMCI.KMLText[j] +
          @"</Folder>";

                }



                string P2_EMCI_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P2_EMCI_KMLText = P2_EMCI_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P2_EMCI.KMLText[j] +
          @"</Folder>";

                }





                string P3_EMCI_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P3_EMCI_KMLText = P3_EMCI_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P3_EMCI.KMLText[j] +
          @"</Folder>";

                }




                string P4_EMCI_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P4_EMCI_KMLText = P4_EMCI_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P4_EMCI.KMLText[j] +
          @"</Folder>";

                }

                string P0_Tebyan_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P0_Tebyan_KMLText = P0_Tebyan_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P0_Tebyan.KMLText[j] +
          @"</Folder>";

                }





                string P1_Tebyan_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P1_Tebyan_KMLText = P1_Tebyan_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P1_Tebyan.KMLText[j] +
          @"</Folder>";

                }



                string P2_Tebyan_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P2_Tebyan_KMLText = P2_Tebyan_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P2_Tebyan.KMLText[j] +
          @"</Folder>";

                }





                string P3_Tebyan_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P3_Tebyan_KMLText = P3_Tebyan_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P3_Tebyan.KMLText[j] +
          @"</Folder>";

                }




                string P4_Tebyan_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P4_Tebyan_KMLText = P4_Tebyan_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P4_Tebyan.KMLText[j] +
          @"</Folder>";

                }




                string P0_Nak_Team_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P0_Nak_Team_KMLText = P0_Nak_Team_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P0_Nak_Team.KMLText[j] +
          @"</Folder>";

                }



                string P1_Nak_Team_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P1_Nak_Team_KMLText = P1_Nak_Team_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P1_Nak_Team.KMLText[j] +
          @"</Folder>";

                }



                string P2_Nak_Team_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P2_Nak_Team_KMLText = P2_Nak_Team_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P2_Nak_Team.KMLText[j] +
          @"</Folder>";

                }





                string P3_Nak_Team_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P3_Nak_Team_KMLText = P3_Nak_Team_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P3_Nak_Team.KMLText[j] +
          @"</Folder>";

                }




                string P4_Nak_Team_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P4_Nak_Team_KMLText = P4_Nak_Team_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P4_Nak_Team.KMLText[j] +
          @"</Folder>";

                }




                string P0_Farhikhtegan_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P0_Farhikhtegan_KMLText = P0_Farhikhtegan_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P0_Farhikhtegan.KMLText[j] +
          @"</Folder>";

                }




                string P1_Farhikhtegan_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P1_Farhikhtegan_KMLText = P1_Farhikhtegan_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P1_Farhikhtegan.KMLText[j] +
          @"</Folder>";

                }



                string P2_Farhikhtegan_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P2_Farhikhtegan_KMLText = P2_Farhikhtegan_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P2_Farhikhtegan.KMLText[j] +
          @"</Folder>";

                }





                string P3_Farhikhtegan_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P3_Farhikhtegan_KMLText = P3_Farhikhtegan_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P3_Farhikhtegan.KMLText[j] +
          @"</Folder>";

                }




                string P4_Farhikhtegan_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P4_Farhikhtegan_KMLText = P4_Farhikhtegan_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P4_Farhikhtegan.KMLText[j] +
          @"</Folder>";

                }

                string P0_Nak_ICT_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P0_Nak_ICT_KMLText = P0_Nak_ICT_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P0_Nak_ICT.KMLText[j] +
          @"</Folder>";

                }


                string P1_Nak_ICT_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P1_Nak_ICT_KMLText = P1_Nak_ICT_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P1_Nak_ICT.KMLText[j] +
          @"</Folder>";

                }



                string P2_Nak_ICT_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P2_Nak_ICT_KMLText = P2_Nak_ICT_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P2_Nak_ICT.KMLText[j] +
          @"</Folder>";

                }





                string P3_Nak_ICT_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P3_Nak_ICT_KMLText = P3_Nak_ICT_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P3_Nak_ICT.KMLText[j] +
          @"</Folder>";

                }




                string P4_Nak_ICT_KMLText = "";
                for (int j = 0; j <= 6; j++)
                {
                    string Prj = "";
                    if (j == 0)
                    {
                        Prj = "PH8";
                    }
                    if (j == 1)
                    {
                        Prj = "PH9";
                    }
                    if (j == 2)
                    {
                        Prj = "TDD";
                    }
                    if (j == 3)
                    {
                        Prj = "USO-ReDrive";
                    }
                    if (j == 4)
                    {
                        Prj = "USO";
                    }
                    if (j == 5)
                    {
                        Prj = "Tower Change";
                    }
                    if (j == 6)
                    {
                        Prj = "5G";
                    }


                    P4_Nak_ICT_KMLText = P4_Nak_ICT_KMLText +
                @"
            <Folder>
            <name> " + Prj + @" </name>"
    + P4_Nak_ICT.KMLText[j] +
          @"</Folder>";

                }



                string P0_KMLText = "";
                string P1_KMLText = "";
                string P2_KMLText = "";
                string P3_KMLText = "";
                string P4_KMLText = "";

                P0_KMLText = P0_FaraCell_KMLText + P0_Figetel_KMLText + P0_Vihan_KMLText + P0_Mahna_KMLText + P0_PCOM_KMLText + P0_EMCI_KMLText + P0_Tebyan_KMLText + P0_Nak_Team_KMLText + P0_Farhikhtegan_KMLText + P0_Nak_ICT_KMLText;
                P1_KMLText = P1_FaraCell_KMLText + P1_Figetel_KMLText + P1_Vihan_KMLText + P1_Mahna_KMLText + P1_PCOM_KMLText + P1_EMCI_KMLText + P1_Tebyan_KMLText + P1_Nak_Team_KMLText+ P1_Farhikhtegan_KMLText + P1_Nak_ICT_KMLText;
                P2_KMLText = P2_FaraCell_KMLText + P2_Figetel_KMLText + P2_Vihan_KMLText + P2_Mahna_KMLText + P2_PCOM_KMLText + P2_EMCI_KMLText + P2_Tebyan_KMLText + P2_Nak_Team_KMLText + P2_Farhikhtegan_KMLText + P2_Nak_ICT_KMLText;
                P3_KMLText = P3_FaraCell_KMLText + P3_Figetel_KMLText + P3_Vihan_KMLText + P3_Mahna_KMLText + P3_PCOM_KMLText + P3_EMCI_KMLText + P3_Tebyan_KMLText + P3_Nak_Team_KMLText + P3_Farhikhtegan_KMLText + P3_Nak_ICT_KMLText;
                P4_KMLText = P4_FaraCell_KMLText + P4_Figetel_KMLText + P4_Vihan_KMLText + P4_Mahna_KMLText + P4_PCOM_KMLText + P4_EMCI_KMLText + P4_Tebyan_KMLText + P4_Nak_Team_KMLText + P4_Farhikhtegan_KMLText + P4_Nak_ICT_KMLText;




                P0_KMLText = @"
            <Folder>
            <name> " + "P0" + @" </name>" + @"
            <Folder>
            <name> " + "FaraCell" + @" </name>"
            + P0_FaraCell_KMLText +
            @" </Folder>" + @"
            <Folder>
            <name> " + "Figetel" + @" </name> "
              + P0_Figetel_KMLText +
              @" </Folder>" + @"
            <Folder>
            <name> " + "Vihan" + @" </name> "
             + P0_Vihan_KMLText +
             @" </Folder>" + @"
            <Folder>
            <name> " + "Mahna" + @" </name> "
              + P0_Mahna_KMLText +
              @" </Folder>" + @"
            <Folder>
            <name> " + "PCOM" + @" </name> "
              + P0_PCOM_KMLText +
              @" </Folder>" + @"
            <Folder>
            <name> " + "EMCI" + @" </name> "
              + P0_EMCI_KMLText +
              @" </Folder>" + @"
            <Folder>
            <name> " + "Tebyan" + @" </name> "
              + P0_Tebyan_KMLText +
              @" </Folder>" + @"
            <Folder>
            <name> " + "Nak_Team" + @" </name> "
              + P0_Nak_Team_KMLText +
              @" </Folder>" + @"
            <Folder>
            <name> " + "Farhikhtegan" + @" </name> "
              + P0_Farhikhtegan_KMLText +
              @" </Folder>" + @"
            <Folder>
            <name> " + "Nak_ICT" + @" </name> "
              + P0_Nak_ICT_KMLText +
              @" </Folder>" +
             @" </Folder>";








                P1_KMLText = @"
            <Folder>
            <name> " + "P1" + @" </name>" + @"
            <Folder>
            <name> " + "FaraCell" + @" </name>"
+ P1_FaraCell_KMLText +
@" </Folder>" + @"
            <Folder>
            <name> " + "Figetel" + @" </name> "
  + P1_Figetel_KMLText +
  @" </Folder>"+ @"
            <Folder>
            <name> " + "Vihan" + @" </name> "
 + P1_Vihan_KMLText +
 @" </Folder>" + @"
            <Folder>
            <name> " + "Mahna" + @" </name> "
  + P1_Mahna_KMLText +
  @" </Folder>" + @"
            <Folder>
            <name> " + "PCOM" + @" </name> "
  + P1_PCOM_KMLText +
  @" </Folder>" + @"
            <Folder>
            <name> " + "EMCI" + @" </name> "
  + P1_EMCI_KMLText +
  @" </Folder>" + @"
            <Folder>
            <name> " + "Tebyan" + @" </name> "
  + P1_Tebyan_KMLText +
  @" </Folder>" + @"
            <Folder>
            <name> " + "Nak_Team" + @" </name> "
  + P1_Nak_Team_KMLText +
  @" </Folder>" + @"
            <Folder>
            <name> " + "Farhikhtegan" + @" </name> "
  + P1_Farhikhtegan_KMLText +
  @" </Folder>" + @"
            <Folder>
            <name> " + "Nak_ICT" + @" </name> "
  + P1_Nak_ICT_KMLText +
  @" </Folder>" +
 @" </Folder>";


                P2_KMLText = @"
            <Folder>
            <name> " + "P2" + @" </name>" + @"
            <Folder>
            <name> " + "FaraCell" + @" </name>"
+ P2_FaraCell_KMLText +
@" </Folder>" + @"
            <Folder>
            <name> " + "Figetel" + @" </name> "
+ P2_Figetel_KMLText +
@" </Folder>" + @"
            <Folder>
            <name> " + "Vihan" + @" </name> "
+ P2_Vihan_KMLText +
@" </Folder>" + @"
            <Folder>
            <name> " + "Mahna" + @" </name> "
+ P2_Mahna_KMLText +
@" </Folder>" + @"
            <Folder>
            <name> " + "PCOM" + @" </name> "
+ P2_PCOM_KMLText +
@" </Folder>" + @"
            <Folder>
            <name> " + "EMCI" + @" </name> "
+ P2_EMCI_KMLText +
@" </Folder>" + @"
            <Folder>
            <name> " + "Tebyan" + @" </name> "
+ P2_Tebyan_KMLText +
@" </Folder>" + @"
            <Folder>
            <name> " + "Nak_Team" + @" </name> "
+ P2_Nak_Team_KMLText +
@" </Folder>" + @"
            <Folder>
            <name> " + "Farhikhtegan" + @" </name> "
+ P2_Farhikhtegan_KMLText +
@" </Folder>" + @"
            <Folder>
            <name> " + "Nak_ICT" + @" </name> "
+ P2_Nak_ICT_KMLText +
@" </Folder>" +
@" </Folder>";



                P3_KMLText = @"
            <Folder>
            <name> " + "P3" + @" </name>" + @"
            <Folder>
            <name> " + "FaraCell" + @" </name>"
+ P3_FaraCell_KMLText +
@" </Folder>" + @"
            <Folder>
            <name> " + "Figetel" + @" </name> "
  + P3_Figetel_KMLText +
  @" </Folder>" + @"
            <Folder>
            <name> " + "Vihan" + @" </name> "
 + P3_Vihan_KMLText +
 @" </Folder>" + @"
            <Folder>
            <name> " + "Mahna" + @" </name> "
  + P3_Mahna_KMLText +
  @" </Folder>" + @"
            <Folder>
            <name> " + "PCOM" + @" </name> "
  + P3_PCOM_KMLText +
  @" </Folder>" + @"
            <Folder>
            <name> " + "EMCI" + @" </name> "
  + P3_EMCI_KMLText +
  @" </Folder>" + @"
            <Folder>
            <name> " + "Tebyan" + @" </name> "
  + P3_Tebyan_KMLText +
  @" </Folder>" + @"
            <Folder>
            <name> " + "Nak_Team" + @" </name> "
  + P3_Nak_Team_KMLText +
  @" </Folder>" + @"
            <Folder>
            <name> " + "Farhikhtegan" + @" </name> "
  + P3_Farhikhtegan_KMLText +
  @" </Folder>" + @"
            <Folder>
            <name> " + "Nak_ICT" + @" </name> "
  + P3_Nak_ICT_KMLText +
  @" </Folder>" +
 @" </Folder>";



                P4_KMLText = @"
            <Folder>
            <name> " + "P4" + @" </name>" + @"
            <Folder>
            <name> " + "FaraCell" + @" </name>"
+ P4_FaraCell_KMLText +
@" </Folder>" + @"
            <Folder>
            <name> " + "Figetel" + @" </name> "
  + P4_Figetel_KMLText +
  @" </Folder>" + @"
            <Folder>
            <name> " + "Vihan" + @" </name> "
 + P4_Vihan_KMLText +
 @" </Folder>" + @"
            <Folder>
            <name> " + "Mahna" + @" </name> "
  + P4_Mahna_KMLText +
  @" </Folder>" + @"
            <Folder>
            <name> " + "PCOM" + @" </name> "
  + P4_PCOM_KMLText +
  @" </Folder>" + @"
            <Folder>
            <name> " + "EMCI" + @" </name> "
  + P4_EMCI_KMLText +
  @" </Folder>" + @"
            <Folder>
            <name> " + "Tebyan" + @" </name> "
  + P4_Tebyan_KMLText +
  @" </Folder>" + @"
            <Folder>
            <name> " + "Nak_Team" + @" </name> "
  + P4_Nak_Team_KMLText +
  @" </Folder>" + @"
            <Folder>
            <name> " + "Farhikhtegan" + @" </name> "
  + P4_Farhikhtegan_KMLText +
  @" </Folder>" + @"
            <Folder>
            <name> " + "Nak_ICT" + @" </name> "
  + P4_Nak_ICT_KMLText +
  @" </Folder>" +
 @" </Folder>";



                textWriter.WriteStartElement(Fixed_part + P0_KMLText + P1_KMLText + P2_KMLText + P3_KMLText+ P4_KMLText+ @"</Document>" +
                                                               @"</kml>");




                textWriter.Close();

                string contents = File.ReadAllText(file);
                string output = contents.Substring(1, contents.Length - 3);
                File.WriteAllText(file, output);






                MessageBox.Show("Finished");




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
                Template_workbook = xlApp.Workbooks.Open(file);
                Excel.Worksheet sheet1 = Template_workbook.Worksheets[1];


                Excel.Range KMLData = sheet1.get_Range("A2", "Q" + sheet1.UsedRange.Rows.Count);
                object[,] KMLs = (object[,])KMLData.Value;

                for (int k = 0; k < sheet1.UsedRange.Rows.Count - 1; k++)
                {
                    string Location = "";
                    if (KMLs[k + 1, 1] != null)
                    {
                        Location = KMLs[k + 1, 1].ToString();
                    }
                    string SiteCode_Technology = "";
                    if (KMLs[k + 1, 2] != null)
                    {
                        SiteCode_Technology = KMLs[k + 1, 2].ToString();
                    }
                    string Location_Band = "";
                    if (KMLs[k + 1, 3] != null)
                    {
                        Location_Band = KMLs[k + 1, 3].ToString();
                    }
                    string Project = "";
                    if (KMLs[k + 1, 4] != null)
                    {
                        Project = KMLs[k + 1, 4].ToString();
                    }
                    string Vendor = "";
                    if (KMLs[k + 1, 5] != null)
                    {
                        Vendor = KMLs[k + 1, 5].ToString();
                    }
                    string City = "";
                    if (KMLs[k + 1, 6] != null)
                    {
                        City = KMLs[k + 1, 6].ToString();
                    }
                    string RSSV_Date_in_Spring = "";
                    if (KMLs[k + 1, 7] != null)
                    {
                        RSSV_Date_in_Spring = KMLs[k + 1, 7].ToString();
                    }
                    string Status = "";
                    if (KMLs[k + 1, 8] != null)
                    {
                        Status = KMLs[k + 1, 8].ToString();
                    }
                    string Pass_Date = "";
                    if (KMLs[k + 1, 9] != null)
                    {
                        Pass_Date = KMLs[k + 1, 9].ToString();
                    }
                    string Traffic = "";
                    if (KMLs[k + 1, 10] != null)
                    {
                        Traffic = KMLs[k + 1, 10].ToString();
                    }
                    string Assignment_to_SC_Date = "";
                    if (KMLs[k + 1, 11] != null)
                    {
                        Assignment_to_SC_Date = KMLs[k + 1, 11].ToString();
                    }
                    string SSV_Owner = "";
                    if (KMLs[k + 1, 12] != null)
                    {
                        SSV_Owner = KMLs[k + 1, 12].ToString();
                    }
                    string Submit_to_OPT_Team_Date = "";
                    if (KMLs[k + 1, 13] != null)
                    {
                        Submit_to_OPT_Team_Date = KMLs[k + 1, 13].ToString();
                    }
                    string Priorities = "";
                    if (KMLs[k + 1, 14] != null)
                    {
                        Priorities = KMLs[k + 1, 14].ToString();
                    }
                    string Comment_External = "";
                    if (KMLs[k + 1, 15] != null)
                    {
                        Comment_External = KMLs[k + 1, 15].ToString();
                    }
                    string NAK_Team = "";
                    if (KMLs[k + 1, 16] != null)
                    {
                        NAK_Team = KMLs[k + 1, 16].ToString();
                    }
                    string NAK_SCOPE = "";
                    if (KMLs[k + 1, 17] != null)
                    {
                        NAK_SCOPE = KMLs[k + 1, 17].ToString();
                    }


                    KMLTable.Rows.Add(Location, SiteCode_Technology, Location_Band, Project, Vendor, City, RSSV_Date_in_Spring, Status, Pass_Date, Traffic, Assignment_to_SC_Date, SSV_Owner, Submit_to_OPT_Team_Date, Priorities, Comment_External, NAK_Team, NAK_SCOPE);

                }


                string Query = @"select substring([کد سایت],1,6) as 'Site', [عرض جغرافیایی] as 'Lat', [طول جغرافیایی] as 'Lon'  from ARAS_DIA ";
                ARASTable = Query_Execution_Table_Output(Query);



                // Join to find LAT and LON
                var SSVTable = (from pd in KMLTable.AsEnumerable()
                                join od in ARASTable.AsEnumerable() on new { f1 = pd.Field<string>("Location") } equals new { f1 = od.Field<string>("Site") } into od
                                from new_od in od.DefaultIfEmpty()
                                select new
                                {
                                    Location = pd.Field<string>("Location"),
                                    LAT = (new_od != null ? new_od.Field<double>("LAT") : -1),
                                    LON = (new_od != null ? new_od.Field<double>("LON") : -1),
                                    SiteCode_Technology = pd.Field<string>("SiteCode_Technology"),
                                    Location_Band = pd.Field<string>("Location_Band"),
                                    Project = pd.Field<string>("Project"),
                                    Vendor = pd.Field<string>("Vendor"),
                                    City = pd.Field<string>("City"),
                                    RSSV_Date_in_Spring = pd.Field<string>("RSSV_Date_in_Spring"),
                                    Status = pd.Field<string>("Status"),
                                    Pass_Date = pd.Field<string>("Pass_Date"),
                                    Traffic = pd.Field<string>("Traffic"),
                                    Assignment_to_SC_Date = pd.Field<string>("Assignment_to_SC_Date"),
                                    SSV_Owner = pd.Field<string>("SSV_Owner"),
                                    Submit_to_OPT_Team_Date = pd.Field<string>("Submit_to_OPT_Team_Date"),
                                    Priorities = pd.Field<string>("Priorities"),
                                    Comment_External = pd.Field<string>("Comment_External"),
                                    NAK_Team = pd.Field<string>("NAK_Team"),
                                    NAK_SCOPE = pd.Field<string>("NAK_SCOPE")

                                }).ToList();


                SSV_KMLTable = ConvertToDataTable(SSVTable);

                label1.Text = "File is Loaded";
                label1.BackColor = Color.GreenYellow;

            }



        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked==true)
            {
                KML_Type = "Subcontractor_Project";
                checkBox2.Checked = false;
                checkBox3.Checked = false;
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                KML_Type = "Periority_Subcontractor_Project";
                checkBox1.Checked = false;
                checkBox3.Checked = false;
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked == true)
            {
                KML_Type = "Subcontractor_Periority_Project_Issue";
                checkBox1.Checked = false;
                checkBox2.Checked = false;
            }
        }
    }


}

