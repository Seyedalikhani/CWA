using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using GMap.NET;
using GMap.NET.WindowsForms;
using GMap.NET.WindowsForms.Markers;
using GMap.NET.MapProviders;
using ClosedXML.Excel;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Threading;
using System.Data.SqlClient;
using System.Collections;



namespace CWA
{
    public partial class MAP : Form
    {

        public MAP()
        {
            InitializeComponent();
        }


        public Main form1;


        public string ConnectionString = "";
        public SqlConnection connection = new SqlConnection();

        public string Server_Name = @"AHMAD\" + "SQLEXPRESS";
        public string DataBase_Name = "NAK";


        //public string Server_Name = "PERFORMANCEDB";
        //public string DataBase_Name = "Performance_NAK";


        public DataTable ARAS_DATA_Table = new DataTable();

        public GMapOverlay Layer1;
        public GMapOverlay Layer2;
        public GMapOverlay Layer3;
        public GMapOverlay Layer4;
        public GMapOverlay Layer5;
        public GMapOverlay Layer6;
        public GMapOverlay Layer7;
        public GMapOverlay Layer8;
        public GMapOverlay Layer9;
        public GMapOverlay Layer10;
        public GMapOverlay Layer11;
        public GMapOverlay Layer12;
        public GMapOverlay Layer13;
        public GMapOverlay Layer14;
        public GMapOverlay Layer15;

        public GMapOverlay Layer16;
        public GMapOverlay Layer17;

        public GMapOverlay Layer20;


        public class TNodeItem
        {
            public string CellName { get; set; }
            public string Azimuth { get; set; }
            public string Band { get; set; }
            public string M_Tilt { get; set; }
            public string E_Tilt { get; set; }
        }

        public class TNode
        {
            public double Lat { get; set; }
            public double Lon { get; set; }
            public string Site { get; set; }
            public string Name { get; set; }
            public string BSC { get; set; }
            public string RNC { get; set; }
            public string LAC_Name { get; set; }
            public string Code_Name { get; set; }
            public string Band_Name { get; set; }
            public string Site_Type_1 { get; set; }
            public string Tower_Type { get; set; }
            public string Tower_Heihgt { get; set; }
            public string Site_Type_2 { get; set; }
            public string BTS_Type { get; set; }
            public string Antenna_Type { get; set; }
            public string Covergae { get; set; }
            public string Link { get; set; }
            public string Province { get; set; }
            public string City { get; set; }

            public int Site_Flag_in_DB { get; set; }
            public List<TNodeItem> Items { get; set; }
        }

        public List<TNode> Nodes;


        public string[] Site_List = new string[1000000];
        public string[] All_Site_List = new string[1000000];

        public string Selected_Province = "";

        public DataTable Code_Sites_2G_Table = new DataTable();
        public DataTable Code_Sites_3G_Table = new DataTable();


        public List<PointLatLng> points1 = new List<PointLatLng>();
        public List<PointLatLng> points2 = new List<PointLatLng>();

        public GMapOverlay polyOverlay1 = new GMapOverlay("polygons");
        public GMapOverlay polyOverlay2 = new GMapOverlay("polygons");

        public int Polygoner = 0;
        public int Polygoner1 = 0;


        double lat_max = 0;
        double lng_max = 0;

        double lat_min = 100;
        double lng_min = 100;


        public double Mod360(double a)
        {
            while (a >= 360) a = a - 360;
            while (a < 360) a = a + 360;
            return a;
        }


        public DataTable ConvertToDataTable<T>(IList<T> data)
        {
            PropertyDescriptorCollection properties =
               TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();
            foreach (PropertyDescriptor prop in properties)
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            foreach (T item in data)
            {
                DataRow row = table.NewRow();
                foreach (PropertyDescriptor prop in properties)
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                table.Rows.Add(row);
            }
            return table;

        }






        public MAP(Form form)
        {
            InitializeComponent();
            form1 = (Main)form;
        }


        void comboBox1_MouseWheel(object sender, MouseEventArgs e)
        {
            ((HandledMouseEventArgs)e).Handled = true;
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {





            comboBox4.Items.Clear();
            comboBox1.MouseWheel += new MouseEventHandler(comboBox1_MouseWheel);
            gMapControl1.Zoom = 5;
            comboBox2.Items.Clear();
            Layer1.Clear();
            Layer2.Clear();
            Layer3.Clear();
            Layer4.Clear();
            Layer5.Clear();
            Layer6.Clear();
            Layer7.Clear();
            Layer8.Clear();
            Layer9.Clear();
            Layer10.Clear();
            Layer11.Clear();
            Layer12.Clear();
            Layer13.Clear();
            Layer14.Clear();
            Layer15.Clear();

            Layer16.Clear();
            Layer17.Clear();

            Layer20.Clear();
            Nodes = new List<TNode>();
            Site_List = new string[1000000];

            Selected_Province = comboBox1.SelectedItem.ToString();



            string MAP_DATA = @"select * from [ARAS_DB] WHERE [PROVINCE_EN]='" + Selected_Province + "' order by LOCATION ,Technology";
            SqlCommand MAP_DATA_Quary = new SqlCommand(MAP_DATA, connection);
            MAP_DATA_Quary.CommandTimeout = 0;
            MAP_DATA_Quary.ExecuteNonQuery();
            ARAS_DATA_Table = new DataTable();
            SqlDataAdapter dataAdapter_MAP_DATA_Table = new SqlDataAdapter(MAP_DATA_Quary);
            dataAdapter_MAP_DATA_Table.Fill(ARAS_DATA_Table);




            //DataTable MAP_DATA_Table = new DataTable();
            //MAP_DATA_Table.Columns.Add("Contractor", typeof(string));
            //MAP_DATA_Table.Columns.Add("Province", typeof(string));
            //MAP_DATA_Table.Columns.Add("City", typeof(string));
            //MAP_DATA_Table.Columns.Add("Node", typeof(string));
            //MAP_DATA_Table.Columns.Add("Technology", typeof(string));
            //MAP_DATA_Table.Columns.Add("Band", typeof(string));
            //MAP_DATA_Table.Columns.Add("Site Name", typeof(string));
            //MAP_DATA_Table.Columns.Add("Site", typeof(string));
            //MAP_DATA_Table.Columns.Add("Cell", typeof(string));
            //MAP_DATA_Table.Columns.Add("Lat", typeof(double));
            //MAP_DATA_Table.Columns.Add("Long", typeof(double));
            //MAP_DATA_Table.Columns.Add("Azimuth", typeof(double));
            //MAP_DATA_Table.Columns.Add("Height", typeof(double));
            //MAP_DATA_Table.Columns.Add("LAC", typeof(int));
            //MAP_DATA_Table.Columns.Add("WLL MCI", typeof(string));
            //MAP_DATA_Table.Columns.Add("Site Type", typeof(string));
            //MAP_DATA_Table.Columns.Add("Mechanical Tilt", typeof(double));
            //MAP_DATA_Table.Columns.Add("Electrical Tilt", typeof(double));
            //MAP_DATA_Table.Columns.Add("Antenna Type", typeof(string));
            //MAP_DATA_Table.Columns.Add("Coverage Type", typeof(string));
            //MAP_DATA_Table.Columns.Add("Segmentation Optimization", typeof(string));
            //MAP_DATA_Table.Columns.Add("Address", typeof(string));


            int site_ind = 0;
            int node_counter = 0;
            int k1 = 0;
            string Band_Text = "";

            string BSC_Name = "";
            //string RNC_Name = "";

            for (int k = 0; k < ARAS_DATA_Table.Rows.Count; k++)
            {

                //string BSC_Name = "";
                string RNC_Name = "";


                string Province = ARAS_DATA_Table.Rows[k][0].ToString();
                if (Province == Selected_Province)
                {
                    k1++;
                    // string Province = "";
                    string Contractor = "";
                    string Technology = ARAS_DATA_Table.Rows[k][7].ToString();
                    string Band = ARAS_DATA_Table.Rows[k][8].ToString();
                    string Site_Type = ARAS_DATA_Table.Rows[k][10].ToString();
                    string Node = ARAS_DATA_Table.Rows[k][9].ToString();
                    string Site = ARAS_DATA_Table.Rows[k][3].ToString();
                    string Province_Index = Site.Substring(0, 2);
                    string Cell = ARAS_DATA_Table.Rows[k][2].ToString();
                    string LAT = ARAS_DATA_Table.Rows[k][4].ToString();
                    string LONG = ARAS_DATA_Table.Rows[k][5].ToString();
                    string Azimuth = ARAS_DATA_Table.Rows[k][6].ToString();
                    string Band_Index = Cell.Substring(2, 2);
                    string Vendor_Index = "";

                    if (Node != "")
                    {
                        Vendor_Index = Node.Substring(Node.Length - 1, 1);
                        if (!comboBox4.Items.Contains(Node))
                        {
                            comboBox4.Items.Add(Node);

                        }
                    }





                    if (Province_Index == "AG")
                    {
                        Province = "WEST AZERBAIJAN"; Contractor = "NAK-Huawei";
                    }
                    if (Province_Index == "AR")
                    {
                        Province = "ARDABIL"; Contractor = "FARAFAN";
                    }
                    if (Province_Index == "AS")
                    {
                        Province = "EAST AZERBAIJAN"; Contractor = "NAK-Huawei";
                    }
                    if (Province_Index == "BU")
                    {
                        Province = "BUSHEHR"; Contractor = "FARAFAN";
                    }
                    if (Province_Index == "CH")
                    {
                        Province = "CHAHARMAHAL BAKHTIARI"; Contractor = "NAK-Nokia";
                    }
                    if (Province_Index == "ES")
                    {
                        Province = "ESFAHAN"; Contractor = "Huawei";
                    }
                    if (Province_Index == "FS")
                    {
                        Province = "FARS"; Contractor = "BR-TEL";
                    }
                    if (Province_Index == "GL")
                    {
                        Province = "GILAN"; Contractor = "NAK-North";
                    }
                    if (Province_Index == "GN")
                    {
                        Province = "GOLESTAN"; Contractor = "NAK-North";
                    }
                    if (Province_Index == "HN")
                    {
                        Province = "HAMEDAN"; Contractor = "FARAFAN";
                    }
                    if (Province_Index == "HZ")
                    {
                        Province = "HORMOZGAN"; Contractor = "FARAFAN";
                    }
                    if (Province_Index == "IL")
                    {
                        Province = "ILAM"; Contractor = "BR-TEL";
                    }
                    if (Province_Index == "KB")
                    {
                        Province = "KOHGILUYEH AND BOYER AHMAD"; Contractor = "BR-TEL";
                    }
                    if (Province_Index == "KD")
                    {
                        Province = "KORDESTAN"; Contractor = "FARAFAN";
                    }
                    if (Province_Index == "KH")
                    {
                        Province = "KHORASAN RAZAVI"; Contractor = "NAK-Nokia";
                    }
                    if (Province_Index == "KM")
                    {
                        Province = "KERMAN"; Contractor = "NAK-Nokia";
                    }
                    if (Province_Index == "KS")
                    {
                        Province = "KERMANSHAH"; Contractor = "FARAFAN";
                    }
                    if (Province_Index == "KZ")
                    {
                        Province = "KHOUZESTAN"; Contractor = "NAK-Huawei";
                    }
                    if (Province_Index == "LN")
                    {
                        Province = "LORESTAN"; Contractor = "FARAFAN";
                    }
                    if (Province_Index == "MA")
                    {
                        Province = "MAZANDARAN"; Contractor = "NAK-North";
                    }
                    if (Province_Index == "MK")
                    {
                        Province = "MARKAZI"; Contractor = "BR-TEL";
                    }
                    if (Province_Index == "NK")
                    {
                        Province = "NORTH KHORASAN"; Contractor = "BR-TEL";
                    }
                    if (Province_Index == "QM")
                    {
                        Province = "QOM"; Contractor = "Huawei";
                    }
                    if (Province_Index == "QN")
                    {
                        Province = "QAZVIN"; Contractor = "FARAFAN";
                    }
                    if (Province_Index == "SB")
                    {
                        Province = "SISTAN VA BALUCHESTAN"; Contractor = "BR-TEL";
                    }
                    if (Province_Index == "SK")
                    {
                        Province = "SOUTH KHORASAN"; Contractor = "BR-TEL";
                    }
                    if (Province_Index == "SM")
                    {
                        Province = "SEMNAN"; Contractor = "NAK-Nokia";
                    }
                    if (Province_Index == "YZ")
                    {
                        Province = "YAZD"; Contractor = "NAK-Nokia";
                    }
                    if (Province_Index == "ZN")
                    {
                        Province = "ZANJAN"; Contractor = "FARAFAN";
                    }
                    if (Province_Index == "TH")
                    {
                        Province = "TEHRAN"; Contractor = "NAK-Tehran";
                    }
                    if (Province_Index == "KJ")
                    {
                        Province = "ALBORZ"; Contractor = "NAK-Alborz";
                    }



                    All_Site_List[k1 - 1] = Site;



                    if (!Site_List.Contains(Site))
                    {

                        //var result = (from myrow in Code_Sites_2G_Table.AsEnumerable()
                        //              where myrow.Field<String>("Site") == Site
                        //              select myrow).ToList();
                        //string BSC = "";
                        //if (result.Count!=0)
                        //{
                        //    BSC = result[0].ItemArray[0].ToString();
                        //}



                        //var result1 = (from myrow in Code_Sites_3G_Table.AsEnumerable()
                        //              where myrow.Field<String>("Site") == Site
                        //              select myrow).ToList();
                        //string RNC = "";
                        //if (result1.Count != 0)
                        //{
                        //    RNC = result1[0].ItemArray[0].ToString();
                        //}



                        BSC_Name = "";
                        RNC_Name = "";
                        if (Technology == "2G")
                        {
                            BSC_Name = Node;
                        }
                        if (Technology == "3G")
                        {
                            RNC_Name = Node;
                        }


                        comboBox2.Items.Add(Site);
                        Site_List[site_ind] = Site;
                        site_ind++;
                        Nodes.Add(new TNode()
                        {
                            Site = Site,
                            Lat = Convert.ToDouble(LAT),
                            Lon = Convert.ToDouble(LONG),
                            BSC = BSC_Name,
                            RNC = RNC_Name
                        });
                        node_counter++;

                        Nodes[node_counter - 1].Items = new List<TNodeItem>();
                        Nodes[node_counter - 1].Items.Add(new TNodeItem
                        {
                            CellName = Cell,
                            Azimuth = Azimuth,
                            Band = Band

                        });


                        //GMapMarker marker = new GMarkerGoogle(new PointLatLng(Nodes[site_ind - 1].Lat, Nodes[site_ind - 1].Lon), GMarkerGoogleType.yellow);
                        //marker.ToolTipText =
                        //    "Site= " + Site + '\n';
                        ////  "BSC= " + BSC+ '\n'+
                        ////   "RNC= " + RNC + '\n'; 
                        //Layer1.Markers.Add(marker);



                        if (Site_Type == "Macro")
                        {
                            if (Vendor_Index == "E")
                            {
                                GMapMarker marker1 = new GMarkerGoogle(new PointLatLng(Nodes[site_ind - 1].Lat, Nodes[site_ind - 1].Lon), GMarkerGoogleType.yellow);
                                marker1.ToolTipText =
                                    "Site= " + Site + '\n';
                                    //"BSC= " + BSC_Name + '\n' +
                                    //"RNC= " + RNC_Name + '\n';

                                Layer9.Markers.Add(marker1);
                            }
                            if (Vendor_Index == "H")
                            {
                                GMapMarker marker1 = new GMarkerGoogle(new PointLatLng(Nodes[site_ind - 1].Lat, Nodes[site_ind - 1].Lon), GMarkerGoogleType.green);
                                marker1.ToolTipText =
                                       "Site= " + Site + '\n';
                                    //"BSC= " + BSC_Name + '\n' +
                                    //"RNC= " + RNC_Name + '\n';


                                Layer9.Markers.Add(marker1);
                            }
                            if (Vendor_Index == "N")
                            {
                                GMapMarker marker1 = new GMarkerGoogle(new PointLatLng(Nodes[site_ind - 1].Lat, Nodes[site_ind - 1].Lon), GMarkerGoogleType.red);
                                marker1.ToolTipText =
                                        "Site= " + Site + '\n';
                                    //"BSC= " + BSC_Name + '\n' +
                                    //"RNC= " + RNC_Name + '\n';


                                Layer9.Markers.Add(marker1);
                            }

                        }
                        if (Site_Type == "Micro")
                        {

                            if (Vendor_Index == "E")
                            {
                                GMapMarker marker1 = new GMarkerGoogle(new PointLatLng(Nodes[site_ind - 1].Lat, Nodes[site_ind - 1].Lon), GMarkerGoogleType.yellow);
                                marker1.ToolTipText =
                                        "Site= " + Site + '\n' +
                                    "BSC= " + BSC_Name + '\n' +
                                    "RNC= " + RNC_Name + '\n';


                                Layer10.Markers.Add(marker1);
                            }
                            if (Vendor_Index == "H")
                            {
                                GMapMarker marker1 = new GMarkerGoogle(new PointLatLng(Nodes[site_ind - 1].Lat, Nodes[site_ind - 1].Lon), GMarkerGoogleType.green);
                                marker1.ToolTipText =
                                        "Site= " + Site + '\n' +
                                    "BSC= " + BSC_Name + '\n' +
                                    "RNC= " + RNC_Name + '\n';


                                Layer10.Markers.Add(marker1);
                            }
                            if (Vendor_Index == "N")
                            {
                                GMapMarker marker1 = new GMarkerGoogle(new PointLatLng(Nodes[site_ind - 1].Lat, Nodes[site_ind - 1].Lon), GMarkerGoogleType.red);
                                marker1.ToolTipText =
                                        "Site= " + Site + '\n' +
                                    "BSC= " + BSC_Name + '\n' +
                                    "RNC= " + RNC_Name + '\n';


                                Layer10.Markers.Add(marker1);
                            }

                        }
                        if (Site_Type == "Pico")
                        {

                            if (Vendor_Index == "E")
                            {
                                GMapMarker marker1 = new GMarkerGoogle(new PointLatLng(Nodes[site_ind - 1].Lat, Nodes[site_ind - 1].Lon), GMarkerGoogleType.yellow);
                                marker1.ToolTipText =
                                        "Site= " + Site + '\n' +
                                    "BSC= " + BSC_Name + '\n' +
                                    "RNC= " + RNC_Name + '\n';


                                Layer11.Markers.Add(marker1);
                            }
                            if (Vendor_Index == "H")
                            {
                                GMapMarker marker1 = new GMarkerGoogle(new PointLatLng(Nodes[site_ind - 1].Lat, Nodes[site_ind - 1].Lon), GMarkerGoogleType.green);
                                marker1.ToolTipText =
                                       "Site= " + Site + '\n' +
                                    "BSC= " + BSC_Name + '\n' +
                                    "RNC= " + RNC_Name + '\n';


                                Layer11.Markers.Add(marker1);
                            }
                            if (Vendor_Index == "N")
                            {
                                GMapMarker marker1 = new GMarkerGoogle(new PointLatLng(Nodes[site_ind - 1].Lat, Nodes[site_ind - 1].Lon), GMarkerGoogleType.red);
                                marker1.ToolTipText =
                                       "Site= " + Site + '\n' +
                                    "BSC= " + BSC_Name + '\n' +
                                    "RNC= " + RNC_Name + '\n';


                                Layer11.Markers.Add(marker1);
                            }

                        }




                    }
                    else
                    {
                        var node = Nodes.FirstOrDefault(x => x.Site == Site);   // from list Nodes the first element whch site property of it equals Node_Name is found and set in node variable

                        if (Technology == "3G" && Nodes[site_ind - 1].RNC == "")
                        {
                            if (Site_Type == "Macro")
                            {
                                if (Vendor_Index == "E")
                                {
                                    RNC_Name = Node;
                                    Nodes[site_ind - 1].RNC = Node;
                                    GMapMarker marker1 = new GMarkerGoogle(new PointLatLng(Nodes[site_ind - 1].Lat, Nodes[site_ind - 1].Lon), GMarkerGoogleType.yellow);

                                    marker1.ToolTipText =
                                                "Site= " + Site + '\n' +
                                            "BSC= " + BSC_Name + '\n' +
                                            "RNC= " + RNC_Name + '\n';

                                    Layer12.Markers.Add(marker1);
                                }
                                if (Vendor_Index == "H")
                                {
                                    RNC_Name = Node;
                                    Nodes[site_ind - 1].RNC = Node;
                                    GMapMarker marker1 = new GMarkerGoogle(new PointLatLng(Nodes[site_ind - 1].Lat, Nodes[site_ind - 1].Lon), GMarkerGoogleType.green);

                                    marker1.ToolTipText =
                                                "Site= " + Site + '\n' +
                                            "BSC= " + BSC_Name + '\n' +
                                            "RNC= " + RNC_Name + '\n';

                                    Layer12.Markers.Add(marker1);
                                }
                                if (Vendor_Index == "N")
                                {
                                    RNC_Name = Node;
                                    Nodes[site_ind - 1].RNC = Node;
                                    GMapMarker marker1 = new GMarkerGoogle(new PointLatLng(Nodes[site_ind - 1].Lat, Nodes[site_ind - 1].Lon), GMarkerGoogleType.red);

                                    marker1.ToolTipText =
                                                "Site= " + Site + '\n' +
                                            "BSC= " + BSC_Name + '\n' +
                                            "RNC= " + RNC_Name + '\n';

                                    Layer12.Markers.Add(marker1);
                                }
                            }


                            if (Site_Type == "Micro")
                            {
                                if (Vendor_Index == "E")
                                {
                                    RNC_Name = Node;
                                    Nodes[site_ind - 1].RNC = Node;
                                    GMapMarker marker1 = new GMarkerGoogle(new PointLatLng(Nodes[site_ind - 1].Lat, Nodes[site_ind - 1].Lon), GMarkerGoogleType.yellow);

                                    marker1.ToolTipText =
                                                "Site= " + Site + '\n' +
                                            "BSC= " + BSC_Name + '\n' +
                                            "RNC= " + RNC_Name + '\n';

                                    Layer13.Markers.Add(marker1);
                                }
                                if (Vendor_Index == "H")
                                {
                                    RNC_Name = Node;
                                    Nodes[site_ind - 1].RNC = Node;
                                    GMapMarker marker1 = new GMarkerGoogle(new PointLatLng(Nodes[site_ind - 1].Lat, Nodes[site_ind - 1].Lon), GMarkerGoogleType.green);

                                    marker1.ToolTipText =
                                                "Site= " + Site + '\n' +
                                            "BSC= " + BSC_Name + '\n' +
                                            "RNC= " + RNC_Name + '\n';

                                    Layer13.Markers.Add(marker1);
                                }
                                if (Vendor_Index == "N")
                                {
                                    RNC_Name = Node;
                                    Nodes[site_ind - 1].RNC = Node;
                                    GMapMarker marker1 = new GMarkerGoogle(new PointLatLng(Nodes[site_ind - 1].Lat, Nodes[site_ind - 1].Lon), GMarkerGoogleType.red);

                                    marker1.ToolTipText =
                                                "Site= " + Site + '\n' +
                                            "BSC= " + BSC_Name + '\n' +
                                            "RNC= " + RNC_Name + '\n';

                                    Layer13.Markers.Add(marker1);
                                }
                            }


                            if (Site_Type == "Pico")
                            {
                                if (Vendor_Index == "E")
                                {
                                    RNC_Name = Node;
                                    Nodes[site_ind - 1].RNC = Node;
                                    GMapMarker marker1 = new GMarkerGoogle(new PointLatLng(Nodes[site_ind - 1].Lat, Nodes[site_ind - 1].Lon), GMarkerGoogleType.yellow);

                                    marker1.ToolTipText =
                                                "Site= " + Site + '\n' +
                                            "BSC= " + BSC_Name + '\n' +
                                            "RNC= " + RNC_Name + '\n';

                                    Layer14.Markers.Add(marker1);
                                }
                                if (Vendor_Index == "H")
                                {
                                    RNC_Name = Node;
                                    Nodes[site_ind - 1].RNC = Node;
                                    GMapMarker marker1 = new GMarkerGoogle(new PointLatLng(Nodes[site_ind - 1].Lat, Nodes[site_ind - 1].Lon), GMarkerGoogleType.green);

                                    marker1.ToolTipText =
                                                "Site= " + Site + '\n' +
                                            "BSC= " + BSC_Name + '\n' +
                                            "RNC= " + RNC_Name + '\n';

                                    Layer14.Markers.Add(marker1);
                                }
                                if (Vendor_Index == "N")
                                {
                                    RNC_Name = Node;
                                    Nodes[site_ind - 1].RNC = Node;
                                    GMapMarker marker1 = new GMarkerGoogle(new PointLatLng(Nodes[site_ind - 1].Lat, Nodes[site_ind - 1].Lon), GMarkerGoogleType.red);

                                    marker1.ToolTipText =
                                                "Site= " + Site + '\n' +
                                            "BSC= " + BSC_Name + '\n' +
                                            "RNC= " + RNC_Name + '\n';

                                    Layer14.Markers.Add(marker1);
                                }
                            }


                        }

                        if (node != null)
                        {
                            node.Items.Add(new TNodeItem
                            {
                                CellName = Cell,
                                Azimuth = Azimuth,
                                Band = Band
                            });


                            double r1 = 0.0005;
                            double angle = 5;
                            int ind = 0;
                            foreach (TNodeItem item in node.Items)
                            {

                                if (node.Items[ind].Azimuth != "")
                                {
                                    double AZ1 = Convert.ToDouble(node.Items[ind].Azimuth);
                                    // double AZ1 = Convert.ToDouble(Azimuth);

                                    if (node.Items[ind].Band == Band && Band == "G900")
                                    {
                                        r1 = 0.0005;
                                        angle = 53;
                                        List<PointLatLng> points = new List<PointLatLng>();
                                        points.Add(new PointLatLng(node.Lat, node.Lon));
                                        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 - angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 - angle / 2) * Math.PI / 180)));
                                        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 + angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 + angle / 2) * Math.PI / 180)));

                                        GMapPolygon polygon = new GMapPolygon(points, item.CellName + " " + item.Azimuth.ToString());

                                        polygon.Fill = new SolidBrush(Color.FromArgb(200, Color.Red));
                                        polygon.Stroke = new Pen(Color.Red, 1);
                                        Layer2.Polygons.Add(polygon);
                                    }

                                    if (node.Items[ind].Band == Band && Band == "G1800")
                                    {
                                        r1 = 0.0007;
                                        angle = 45;
                                        List<PointLatLng> points = new List<PointLatLng>();
                                        points.Add(new PointLatLng(node.Lat, node.Lon));
                                        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 - angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 - angle / 2) * Math.PI / 180)));
                                        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 + angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 + angle / 2) * Math.PI / 180)));

                                        GMapPolygon polygon = new GMapPolygon(points, item.CellName + " " + item.Azimuth.ToString());

                                        polygon.Fill = new SolidBrush(Color.FromArgb(200, Color.Green));
                                        polygon.Stroke = new Pen(Color.Green, 1);
                                        Layer3.Polygons.Add(polygon);
                                    }


                                    if (node.Items[ind].Band == Band && Band == "U900")
                                    {
                                        r1 = 0.0009;
                                        angle = 37;
                                        List<PointLatLng> points = new List<PointLatLng>();
                                        points.Add(new PointLatLng(node.Lat, node.Lon));
                                        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 - angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 - angle / 2) * Math.PI / 180)));
                                        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 + angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 + angle / 2) * Math.PI / 180)));

                                        GMapPolygon polygon = new GMapPolygon(points, item.CellName + " " + item.Azimuth.ToString());

                                        polygon.Fill = new SolidBrush(Color.FromArgb(200, Color.Orange));
                                        polygon.Stroke = new Pen(Color.Orange, 1);
                                        Layer4.Polygons.Add(polygon);
                                    }
                                    if (node.Items[ind].Band == Band && Band == "U2100")
                                    {
                                        r1 = 0.0011;
                                        angle = 29;
                                        List<PointLatLng> points = new List<PointLatLng>();
                                        points.Add(new PointLatLng(node.Lat, node.Lon));
                                        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 - angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 - angle / 2) * Math.PI / 180)));
                                        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 + angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 + angle / 2) * Math.PI / 180)));

                                        GMapPolygon polygon = new GMapPolygon(points, item.CellName + " " + item.Azimuth.ToString());

                                        polygon.Fill = new SolidBrush(Color.FromArgb(200, Color.Blue));
                                        polygon.Stroke = new Pen(Color.Blue, 1);
                                        Layer5.Polygons.Add(polygon);
                                    }
                                    if (node.Items[ind].Band == Band && Band == "L1800")
                                    {
                                        r1 = 0.0013;
                                        angle = 21;
                                        List<PointLatLng> points = new List<PointLatLng>();
                                        points.Add(new PointLatLng(node.Lat, node.Lon));
                                        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 - angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 - angle / 2) * Math.PI / 180)));
                                        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 + angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 + angle / 2) * Math.PI / 180)));

                                        GMapPolygon polygon = new GMapPolygon(points, item.CellName + " " + item.Azimuth.ToString());

                                        polygon.Fill = new SolidBrush(Color.FromArgb(200, Color.Gray));
                                        polygon.Stroke = new Pen(Color.Gray, 1);
                                        Layer6.Polygons.Add(polygon);
                                    }
                                    if (node.Items[ind].Band == Band && Band == "L2100")
                                    {
                                        r1 = 0.0015;
                                        angle = 13;
                                        List<PointLatLng> points = new List<PointLatLng>();
                                        points.Add(new PointLatLng(node.Lat, node.Lon));
                                        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 - angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 - angle / 2) * Math.PI / 180)));
                                        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 + angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 + angle / 2) * Math.PI / 180)));

                                        GMapPolygon polygon = new GMapPolygon(points, item.CellName + " " + item.Azimuth.ToString());

                                        polygon.Fill = new SolidBrush(Color.FromArgb(200, Color.Black));
                                        polygon.Stroke = new Pen(Color.Black, 1);
                                        Layer7.Polygons.Add(polygon);

                                    }
                                    if (node.Items[ind].Band == Band && Band == "L2600")
                                    {
                                        r1 = 0.0017;
                                        angle = 5;
                                        List<PointLatLng> points = new List<PointLatLng>();
                                        points.Add(new PointLatLng(node.Lat, node.Lon));
                                        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 - angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 - angle / 2) * Math.PI / 180)));
                                        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 + angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 + angle / 2) * Math.PI / 180)));

                                        GMapPolygon polygon = new GMapPolygon(points, item.CellName + " " + item.Azimuth.ToString());

                                        polygon.Fill = new SolidBrush(Color.FromArgb(200, Color.Brown));
                                        polygon.Stroke = new Pen(Color.Brown, 1);
                                        Layer8.Polygons.Add(polygon);

                                    }
                                }


                                ind++;


                                //    // double AZ1 = 0;
                                //    if (item.Azimuth == "")
                                //{
                                //    r1 = 0;
                                //    string setctor_letter = item.CellName.Substring(item.CellName.Length - 1, 1);
                                //    if (setctor_letter == "A")
                                //    {
                                //        AZ1 = 0;
                                //    }
                                //    else if (setctor_letter == "B")
                                //    {
                                //        AZ1 = 120;
                                //    }
                                //    else if (setctor_letter == "C")
                                //    {
                                //        AZ1 = 240;
                                //    }
                                //    else if (setctor_letter == "D")
                                //    {
                                //        AZ1 = 300;
                                //    }
                                //    else
                                //    {
                                //        AZ1 = 100;
                                //    }


                                //    if (Band == "G900")
                                //    {
                                //        //r1 = 0.0005;
                                //        angle = 53;
                                //        List<PointLatLng> points = new List<PointLatLng>();
                                //        points.Add(new PointLatLng(node.Lat, node.Lon));
                                //        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 - angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 - angle / 2) * Math.PI / 180)));
                                //        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 + angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 + angle / 2) * Math.PI / 180)));

                                //        GMapPolygon polygon = new GMapPolygon(points, item.CellName + " " + item.Azimuth.ToString());

                                //        polygon.Fill = new SolidBrush(Color.FromArgb(200, Color.Red));
                                //        polygon.Stroke = new Pen(Color.Red, 1);
                                //        Layer2.Polygons.Add(polygon);
                                //    }

                                //    if (Band == "G1800")
                                //    {
                                //        //r1 = 0.0007;
                                //        angle = 45;
                                //        List<PointLatLng> points = new List<PointLatLng>();
                                //        points.Add(new PointLatLng(node.Lat, node.Lon));
                                //        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 - angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 - angle / 2) * Math.PI / 180)));
                                //        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 + angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 + angle / 2) * Math.PI / 180)));

                                //        GMapPolygon polygon = new GMapPolygon(points, item.CellName + " " + item.Azimuth.ToString());

                                //        polygon.Fill = new SolidBrush(Color.FromArgb(200, Color.Green));
                                //        polygon.Stroke = new Pen(Color.Green, 1);
                                //        Layer3.Polygons.Add(polygon);
                                //    }


                                //    if (Band == "U900")
                                //    {
                                //       // r1 = 0.0009;
                                //        angle = 37;
                                //        List<PointLatLng> points = new List<PointLatLng>();
                                //        points.Add(new PointLatLng(node.Lat, node.Lon));
                                //        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 - angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 - angle / 2) * Math.PI / 180)));
                                //        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 + angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 + angle / 2) * Math.PI / 180)));

                                //        GMapPolygon polygon = new GMapPolygon(points, item.CellName + " " + item.Azimuth.ToString());

                                //        polygon.Fill = new SolidBrush(Color.FromArgb(200, Color.Orange));
                                //        polygon.Stroke = new Pen(Color.Orange, 1);
                                //        Layer4.Polygons.Add(polygon);
                                //    }
                                //    if (Band == "U2100")
                                //    {
                                //       // r1 = 0.0011;
                                //        angle = 29;
                                //        List<PointLatLng> points = new List<PointLatLng>();
                                //        points.Add(new PointLatLng(node.Lat, node.Lon));
                                //        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 - angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 - angle / 2) * Math.PI / 180)));
                                //        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 + angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 + angle / 2) * Math.PI / 180)));

                                //        GMapPolygon polygon = new GMapPolygon(points, item.CellName + " " + item.Azimuth.ToString());

                                //        polygon.Fill = new SolidBrush(Color.FromArgb(200, Color.Blue));
                                //        polygon.Stroke = new Pen(Color.Blue, 1);
                                //        Layer5.Polygons.Add(polygon);
                                //    }
                                //    if (Band == "L1800")
                                //    {
                                //       // r1 = 0.0013;
                                //        angle = 21;
                                //        List<PointLatLng> points = new List<PointLatLng>();
                                //        points.Add(new PointLatLng(node.Lat, node.Lon));
                                //        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 - angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 - angle / 2) * Math.PI / 180)));
                                //        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 + angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 + angle / 2) * Math.PI / 180)));

                                //        GMapPolygon polygon = new GMapPolygon(points, item.CellName + " " + item.Azimuth.ToString());

                                //        polygon.Fill = new SolidBrush(Color.FromArgb(200, Color.Gray));
                                //        polygon.Stroke = new Pen(Color.Gray, 1);
                                //        Layer6.Polygons.Add(polygon);
                                //    }
                                //    if (Band == "L2100")
                                //    {
                                //      //  r1 = 0.0015;
                                //        angle = 13;
                                //        List<PointLatLng> points = new List<PointLatLng>();
                                //        points.Add(new PointLatLng(node.Lat, node.Lon));
                                //        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 - angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 - angle / 2) * Math.PI / 180)));
                                //        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 + angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 + angle / 2) * Math.PI / 180)));

                                //        GMapPolygon polygon = new GMapPolygon(points, item.CellName + " " + item.Azimuth.ToString());

                                //        polygon.Fill = new SolidBrush(Color.FromArgb(200, Color.Black));
                                //        polygon.Stroke = new Pen(Color.Black, 1);
                                //        Layer7.Polygons.Add(polygon);

                                //    }
                                //    if (Band == "L2600")
                                //    {
                                //       // r1 = 0.0017;
                                //        angle = 5;
                                //        List<PointLatLng> points = new List<PointLatLng>();
                                //        points.Add(new PointLatLng(node.Lat, node.Lon));
                                //        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 - angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 - angle / 2) * Math.PI / 180)));
                                //        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 + angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 + angle / 2) * Math.PI / 180)));

                                //        GMapPolygon polygon = new GMapPolygon(points, item.CellName + " " + item.Azimuth.ToString());

                                //        polygon.Fill = new SolidBrush(Color.FromArgb(200, Color.Brown));
                                //        polygon.Stroke = new Pen(Color.Brown, 1);
                                //        Layer8.Polygons.Add(polygon);

                                //    }


                                //}
                                //else
                                //{
                                //     AZ1 = Convert.ToDouble(item.Azimuth);
                                //    //AZ1 = Convert.ToDouble(Azimuth);

                                //    if (Band == "G900")
                                //    {
                                //        r1 = 0.0005;
                                //        angle = 53;
                                //        List<PointLatLng> points = new List<PointLatLng>();
                                //        points.Add(new PointLatLng(node.Lat, node.Lon));
                                //        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 - angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 - angle / 2) * Math.PI / 180)));
                                //        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 + angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 + angle / 2) * Math.PI / 180)));

                                //        GMapPolygon polygon = new GMapPolygon(points, item.CellName + " " + item.Azimuth.ToString());

                                //        polygon.Fill = new SolidBrush(Color.FromArgb(200, Color.Red));
                                //        polygon.Stroke = new Pen(Color.Red, 1);
                                //        Layer2.Polygons.Add(polygon);
                                //    }

                                //    if (Band == "G1800")
                                //    {
                                //        r1 = 0.0007;
                                //        angle = 45;
                                //        List<PointLatLng> points = new List<PointLatLng>();
                                //        points.Add(new PointLatLng(node.Lat, node.Lon));
                                //        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 - angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 - angle / 2) * Math.PI / 180)));
                                //        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 + angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 + angle / 2) * Math.PI / 180)));

                                //        GMapPolygon polygon = new GMapPolygon(points, item.CellName + " " + item.Azimuth.ToString());

                                //        polygon.Fill = new SolidBrush(Color.FromArgb(200, Color.Green));
                                //        polygon.Stroke = new Pen(Color.Green, 1);
                                //        Layer3.Polygons.Add(polygon);
                                //    }


                                //    if (Band == "U900")
                                //    {
                                //        r1 = 0.0009;
                                //        angle = 37;
                                //        List<PointLatLng> points = new List<PointLatLng>();
                                //        points.Add(new PointLatLng(node.Lat, node.Lon));
                                //        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 - angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 - angle / 2) * Math.PI / 180)));
                                //        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 + angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 + angle / 2) * Math.PI / 180)));

                                //        GMapPolygon polygon = new GMapPolygon(points, item.CellName + " " + item.Azimuth.ToString());

                                //        polygon.Fill = new SolidBrush(Color.FromArgb(200, Color.Orange));
                                //        polygon.Stroke = new Pen(Color.Orange, 1);
                                //        Layer4.Polygons.Add(polygon);
                                //    }
                                //    if (Band == "U2100")
                                //    {
                                //        r1 = 0.0011;
                                //        angle = 29;
                                //        List<PointLatLng> points = new List<PointLatLng>();
                                //        points.Add(new PointLatLng(node.Lat, node.Lon));
                                //        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 - angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 - angle / 2) * Math.PI / 180)));
                                //        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 + angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 + angle / 2) * Math.PI / 180)));

                                //        GMapPolygon polygon = new GMapPolygon(points, item.CellName + " " + item.Azimuth.ToString());

                                //        polygon.Fill = new SolidBrush(Color.FromArgb(200, Color.Blue));
                                //        polygon.Stroke = new Pen(Color.Blue, 1);
                                //        Layer5.Polygons.Add(polygon);
                                //    }
                                //    if (Band == "L1800")
                                //    {
                                //        r1 = 0.0013;
                                //        angle = 21;
                                //        List<PointLatLng> points = new List<PointLatLng>();
                                //        points.Add(new PointLatLng(node.Lat, node.Lon));
                                //        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 - angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 - angle / 2) * Math.PI / 180)));
                                //        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 + angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 + angle / 2) * Math.PI / 180)));

                                //        GMapPolygon polygon = new GMapPolygon(points, item.CellName + " " + item.Azimuth.ToString());

                                //        polygon.Fill = new SolidBrush(Color.FromArgb(200, Color.Gray));
                                //        polygon.Stroke = new Pen(Color.Gray, 1);
                                //        Layer6.Polygons.Add(polygon);
                                //    }
                                //    if (Band == "L2100")
                                //    {
                                //        r1 = 0.0015;
                                //        angle = 13;
                                //        List<PointLatLng> points = new List<PointLatLng>();
                                //        points.Add(new PointLatLng(node.Lat, node.Lon));
                                //        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 - angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 - angle / 2) * Math.PI / 180)));
                                //        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 + angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 + angle / 2) * Math.PI / 180)));

                                //        GMapPolygon polygon = new GMapPolygon(points, item.CellName + " " + item.Azimuth.ToString());

                                //        polygon.Fill = new SolidBrush(Color.FromArgb(200, Color.Black));
                                //        polygon.Stroke = new Pen(Color.Black, 1);
                                //        Layer7.Polygons.Add(polygon);

                                //    }
                                //    if (Band == "L2600")
                                //    {
                                //        r1 = 0.0017;
                                //        angle = 5;
                                //        List<PointLatLng> points = new List<PointLatLng>();
                                //        points.Add(new PointLatLng(node.Lat, node.Lon));
                                //        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 - angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 - angle / 2) * Math.PI / 180)));
                                //        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 + angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 + angle / 2) * Math.PI / 180)));

                                //        GMapPolygon polygon = new GMapPolygon(points, item.CellName + " " + item.Azimuth.ToString());

                                //        polygon.Fill = new SolidBrush(Color.FromArgb(200, Color.Brown));
                                //        polygon.Stroke = new Pen(Color.Brown, 1);
                                //        Layer8.Polygons.Add(polygon);

                                //    }


                                //}
                                //if (AZ1 < 0 || AZ1 > 360)
                                //{
                                //    string setctor_letter = item.CellName.Substring(item.CellName.Length - 1, 1);
                                //    if (setctor_letter == "A")
                                //    {
                                //        AZ1 = 0;
                                //    }
                                //    else if (setctor_letter == "B")
                                //    {
                                //        AZ1 = 120;
                                //    }
                                //    else if (setctor_letter == "C")
                                //    {
                                //        AZ1 = 240;
                                //    }
                                //    else if (setctor_letter == "D")
                                //    {
                                //        AZ1 = 300;
                                //    }
                                //    else
                                //    {
                                //        AZ1 = 100;
                                //    }


                                //    r1 = 0;
                                //    if (Band == "G900")
                                //    {
                                //     //   r1 = 0.0005;
                                //        angle = 53;
                                //        List<PointLatLng> points = new List<PointLatLng>();
                                //        points.Add(new PointLatLng(node.Lat, node.Lon));
                                //        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 - angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 - angle / 2) * Math.PI / 180)));
                                //        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 + angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 + angle / 2) * Math.PI / 180)));

                                //        GMapPolygon polygon = new GMapPolygon(points, item.CellName + " " + item.Azimuth.ToString());

                                //        polygon.Fill = new SolidBrush(Color.FromArgb(200, Color.Red));
                                //        polygon.Stroke = new Pen(Color.Red, 1);
                                //        Layer2.Polygons.Add(polygon);
                                //    }

                                //    if (Band == "G1800")
                                //    {
                                //        //r1 = 0.0007;
                                //        angle = 45;
                                //        List<PointLatLng> points = new List<PointLatLng>();
                                //        points.Add(new PointLatLng(node.Lat, node.Lon));
                                //        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 - angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 - angle / 2) * Math.PI / 180)));
                                //        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 + angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 + angle / 2) * Math.PI / 180)));

                                //        GMapPolygon polygon = new GMapPolygon(points, item.CellName + " " + item.Azimuth.ToString());

                                //        polygon.Fill = new SolidBrush(Color.FromArgb(200, Color.Green));
                                //        polygon.Stroke = new Pen(Color.Green, 1);
                                //        Layer3.Polygons.Add(polygon);
                                //    }


                                //    if (Band == "U900")
                                //    {
                                //       // r1 = 0.0009;
                                //        angle = 37;
                                //        List<PointLatLng> points = new List<PointLatLng>();
                                //        points.Add(new PointLatLng(node.Lat, node.Lon));
                                //        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 - angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 - angle / 2) * Math.PI / 180)));
                                //        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 + angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 + angle / 2) * Math.PI / 180)));

                                //        GMapPolygon polygon = new GMapPolygon(points, item.CellName + " " + item.Azimuth.ToString());

                                //        polygon.Fill = new SolidBrush(Color.FromArgb(200, Color.Orange));
                                //        polygon.Stroke = new Pen(Color.Orange, 1);
                                //        Layer4.Polygons.Add(polygon);
                                //    }
                                //    if (Band == "U2100")
                                //    {
                                //       // r1 = 0.0011;
                                //        angle = 29;
                                //        List<PointLatLng> points = new List<PointLatLng>();
                                //        points.Add(new PointLatLng(node.Lat, node.Lon));
                                //        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 - angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 - angle / 2) * Math.PI / 180)));
                                //        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 + angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 + angle / 2) * Math.PI / 180)));

                                //        GMapPolygon polygon = new GMapPolygon(points, item.CellName + " " + item.Azimuth.ToString());

                                //        polygon.Fill = new SolidBrush(Color.FromArgb(200, Color.Blue));
                                //        polygon.Stroke = new Pen(Color.Blue, 1);
                                //        Layer5.Polygons.Add(polygon);
                                //    }
                                //    if (Band == "L1800")
                                //    {
                                //      //  r1 = 0.0013;
                                //        angle = 21;
                                //        List<PointLatLng> points = new List<PointLatLng>();
                                //        points.Add(new PointLatLng(node.Lat, node.Lon));
                                //        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 - angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 - angle / 2) * Math.PI / 180)));
                                //        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 + angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 + angle / 2) * Math.PI / 180)));

                                //        GMapPolygon polygon = new GMapPolygon(points, item.CellName + " " + item.Azimuth.ToString());

                                //        polygon.Fill = new SolidBrush(Color.FromArgb(200, Color.Gray));
                                //        polygon.Stroke = new Pen(Color.Gray, 1);
                                //        Layer6.Polygons.Add(polygon);
                                //    }
                                //    if (Band == "L2100")
                                //    {
                                //      //  r1 = 0.0015;
                                //        angle = 13;
                                //        List<PointLatLng> points = new List<PointLatLng>();
                                //        points.Add(new PointLatLng(node.Lat, node.Lon));
                                //        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 - angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 - angle / 2) * Math.PI / 180)));
                                //        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 + angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 + angle / 2) * Math.PI / 180)));

                                //        GMapPolygon polygon = new GMapPolygon(points, item.CellName + " " + item.Azimuth.ToString());

                                //        polygon.Fill = new SolidBrush(Color.FromArgb(200, Color.Black));
                                //        polygon.Stroke = new Pen(Color.Black, 1);
                                //        Layer7.Polygons.Add(polygon);

                                //    }
                                //    if (Band == "L2600")
                                //    {
                                //      //  r1 = 0.0017;
                                //        angle = 5;
                                //        List<PointLatLng> points = new List<PointLatLng>();
                                //        points.Add(new PointLatLng(node.Lat, node.Lon));
                                //        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 - angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 - angle / 2) * Math.PI / 180)));
                                //        points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 + angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 + angle / 2) * Math.PI / 180)));

                                //        GMapPolygon polygon = new GMapPolygon(points, item.CellName + " " + item.Azimuth.ToString());

                                //        polygon.Fill = new SolidBrush(Color.FromArgb(200, Color.Brown));
                                //        polygon.Stroke = new Pen(Color.Brown, 1);
                                //        Layer8.Polygons.Add(polygon);

                                //    }


                                //}



                                //if (Band == "G900")
                                //{
                                //    r1 = 0.0005;
                                //    angle = 53;
                                //    List<PointLatLng> points = new List<PointLatLng>();
                                //    points.Add(new PointLatLng(node.Lat, node.Lon));
                                //    points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 - angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 - angle / 2) * Math.PI / 180)));
                                //    points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 + angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 + angle / 2) * Math.PI / 180)));

                                //    GMapPolygon polygon = new GMapPolygon(points, item.CellName + " " + item.Azimuth.ToString());

                                //    polygon.Fill = new SolidBrush(Color.FromArgb(200, Color.Red));
                                //    polygon.Stroke = new Pen(Color.Red, 1);
                                //    Layer2.Polygons.Add(polygon);
                                //}

                                //if (Band == "G1800")
                                //{
                                //    r1 = 0.0007;
                                //    angle = 45;
                                //    List<PointLatLng> points = new List<PointLatLng>();
                                //    points.Add(new PointLatLng(node.Lat, node.Lon));
                                //    points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 - angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 - angle / 2) * Math.PI / 180)));
                                //    points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 + angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 + angle / 2) * Math.PI / 180)));

                                //    GMapPolygon polygon = new GMapPolygon(points, item.CellName + " " + item.Azimuth.ToString());

                                //    polygon.Fill = new SolidBrush(Color.FromArgb(200, Color.Green));
                                //    polygon.Stroke = new Pen(Color.Green, 1);
                                //    Layer3.Polygons.Add(polygon);
                                //}


                                //if (Band == "U900")
                                //{
                                //    r1 = 0.0009;
                                //    angle = 37;
                                //    List<PointLatLng> points = new List<PointLatLng>();
                                //    points.Add(new PointLatLng(node.Lat, node.Lon));
                                //    points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 - angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 - angle / 2) * Math.PI / 180)));
                                //    points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 + angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 + angle / 2) * Math.PI / 180)));

                                //    GMapPolygon polygon = new GMapPolygon(points, item.CellName + " " + item.Azimuth.ToString());

                                //    polygon.Fill = new SolidBrush(Color.FromArgb(200, Color.Orange));
                                //    polygon.Stroke = new Pen(Color.Orange, 1);
                                //    Layer4.Polygons.Add(polygon);
                                //}
                                //if (Band == "U2100")
                                //{
                                //    r1 = 0.0011;
                                //    angle = 29;
                                //    List<PointLatLng> points = new List<PointLatLng>();
                                //    points.Add(new PointLatLng(node.Lat, node.Lon));
                                //    points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 - angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 - angle / 2) * Math.PI / 180)));
                                //    points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 + angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 + angle / 2) * Math.PI / 180)));

                                //    GMapPolygon polygon = new GMapPolygon(points, item.CellName + " " + item.Azimuth.ToString());

                                //    polygon.Fill = new SolidBrush(Color.FromArgb(200, Color.Blue));
                                //    polygon.Stroke = new Pen(Color.Blue, 1);
                                //    Layer5.Polygons.Add(polygon);
                                //}
                                //if (Band == "L1800")
                                //{
                                //    r1 = 0.0013;
                                //    angle = 21;
                                //    List<PointLatLng> points = new List<PointLatLng>();
                                //    points.Add(new PointLatLng(node.Lat, node.Lon));
                                //    points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 - angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 - angle / 2) * Math.PI / 180)));
                                //    points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 + angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 + angle / 2) * Math.PI / 180)));

                                //    GMapPolygon polygon = new GMapPolygon(points, item.CellName + " " + item.Azimuth.ToString());

                                //    polygon.Fill = new SolidBrush(Color.FromArgb(200, Color.Gray));
                                //    polygon.Stroke = new Pen(Color.Gray, 1);
                                //    Layer6.Polygons.Add(polygon);
                                //}
                                //if (Band == "L2100")
                                //{
                                //    r1 = 0.0015;
                                //    angle = 13;
                                //    List<PointLatLng> points = new List<PointLatLng>();
                                //    points.Add(new PointLatLng(node.Lat, node.Lon));
                                //    points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 - angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 - angle / 2) * Math.PI / 180)));
                                //    points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 + angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 + angle / 2) * Math.PI / 180)));

                                //    GMapPolygon polygon = new GMapPolygon(points, item.CellName + " " + item.Azimuth.ToString());

                                //    polygon.Fill = new SolidBrush(Color.FromArgb(200, Color.Black));
                                //    polygon.Stroke = new Pen(Color.Black, 1);
                                //    Layer7.Polygons.Add(polygon);

                                //}
                                //if (Band == "L2600")
                                //{
                                //    r1 = 0.0017;
                                //    angle = 5;
                                //    List<PointLatLng> points = new List<PointLatLng>();
                                //    points.Add(new PointLatLng(node.Lat, node.Lon));
                                //    points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 - angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 - angle / 2) * Math.PI / 180)));
                                //    points.Add(new PointLatLng(node.Lat + r1 * Math.Cos(Mod360(AZ1 + angle / 2) * Math.PI / 180), node.Lon + r1 * Math.Sin(Mod360(AZ1 + angle / 2) * Math.PI / 180)));

                                //    GMapPolygon polygon = new GMapPolygon(points, item.CellName + " " + item.Azimuth.ToString());

                                //    polygon.Fill = new SolidBrush(Color.FromArgb(200, Color.Brown));
                                //    polygon.Stroke = new Pen(Color.Brown, 1);
                                //    Layer8.Polygons.Add(polygon);

                                //}


                            }

                        }

                    }

                }
            }




        }


        void comboBox2_MouseWheel(object sender, MouseEventArgs e)
        {
            ((HandledMouseEventArgs)e).Handled = true;
        }
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox2.MouseWheel += new MouseEventHandler(comboBox2_MouseWheel);
            Layer20.Markers.Clear();
            var node = Nodes.FirstOrDefault(x => x.Site == comboBox2.SelectedItem.ToString());
            GMapMarker marker = new GMarkerGoogle(new PointLatLng(node.Lat, node.Lon), GMarkerGoogleType.blue);
            gMapControl1.Position = new PointLatLng(node.Lat, node.Lon);
            gMapControl1.Zoom = 15;
            Layer20.Markers.Add(marker);
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

            Layer2.IsVisibile = checkBox1.Checked;
            gMapControl1.Update();
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {

            Layer3.IsVisibile = checkBox2.Checked;
            gMapControl1.Update();
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {

            Layer4.IsVisibile = checkBox3.Checked;
            gMapControl1.Update();
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {

            Layer5.IsVisibile = checkBox4.Checked;
            gMapControl1.Update();
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {

            Layer6.IsVisibile = checkBox5.Checked;
            gMapControl1.Update();
        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {

            Layer7.IsVisibile = checkBox6.Checked;
            gMapControl1.Update();
        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {

            Layer8.IsVisibile = checkBox7.Checked;
            gMapControl1.Update();
        }


        private void checkBox8_CheckedChanged_1(object sender, EventArgs e)
        {
            Layer1.IsVisibile = false;
            // Layer2.IsVisibile = false;
            //Layer3.IsVisibile = false;
            //Layer4.IsVisibile = false;
            //Layer5.IsVisibile = false;
            //Layer6.IsVisibile = false;
            //Layer7.IsVisibile = false;
            //Layer8.IsVisibile = false;
            Layer9.IsVisibile = checkBox8.Checked;
            Layer12.IsVisibile = checkBox8.Checked;
            Layer2.IsVisibile = true;
            Layer3.IsVisibile = true;
            Layer4.IsVisibile = true;
            Layer5.IsVisibile = true;
            Layer6.IsVisibile = true;
            Layer7.IsVisibile = true;
            Layer8.IsVisibile = true;
            gMapControl1.Update();
        }

        private void checkBox9_CheckedChanged_1(object sender, EventArgs e)
        {
            Layer1.IsVisibile = false;
            // Layer2.IsVisibile = false;
            //Layer3.IsVisibile = false;
            //Layer4.IsVisibile = false;
            //Layer5.IsVisibile = false;
            //Layer6.IsVisibile = false;
            //Layer7.IsVisibile = false;
            //Layer8.IsVisibile = false;
            Layer10.IsVisibile = checkBox9.Checked;
            Layer13.IsVisibile = checkBox9.Checked;
            Layer2.IsVisibile = true;
            Layer3.IsVisibile = true;
            Layer4.IsVisibile = true;
            Layer5.IsVisibile = true;
            Layer6.IsVisibile = true;
            Layer7.IsVisibile = true;
            Layer8.IsVisibile = true;
            gMapControl1.Update();
        }

        private void checkBox10_CheckedChanged_1(object sender, EventArgs e)
        {
            Layer1.IsVisibile = false;
            //Layer2.IsVisibile = false;
            //Layer3.IsVisibile = false;
            //Layer4.IsVisibile = false;
            //Layer5.IsVisibile = false;
            //Layer6.IsVisibile = false;
            //Layer7.IsVisibile = false;
            //Layer8.IsVisibile = false;
            Layer11.IsVisibile = checkBox10.Checked;
            Layer14.IsVisibile = checkBox9.Checked;
            Layer2.IsVisibile = true;
            Layer3.IsVisibile = true;
            Layer4.IsVisibile = true;
            Layer5.IsVisibile = true;
            Layer6.IsVisibile = true;
            Layer7.IsVisibile = true;
            Layer8.IsVisibile = true;
            gMapControl1.Update();
        }



        private void Form5_Load(object sender, EventArgs e)
        {

            //string userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            //string currentUser = userName.Substring(8, userName.Length - 8);

            //string[] authorizedUsers = new string[]
            //{
            //           "ahmad.alikhani",
            //           "elham.vafaeinejad",
            //           "alireza.aghagoli",
            //          "amir.moshfeghkia",
            //          "mohammadali.amini",
            //          "arash.naghdehforoushha",
            //           "masoud.zaerin"


            //};

            //if (authorizedUsers.Contains(currentUser.ToLower()))
            //{
            //    string Authorized = "OK";
            //}
            //else
            //{
            //    MessageBox.Show("Limited Access! Need Authorization by Admin");
            //    this.Close();
            //}






            gMapControl1.MapProvider = GoogleMapProvider.Instance;
            GMaps.Instance.Mode = AccessMode.ServerAndCache;
            Directory.CreateDirectory(Application.StartupPath + @"\Cache");
            gMapControl1.DragButton = MouseButtons.Left;
            gMapControl1.Position = new PointLatLng(31.9118, 54.33629);
            gMapControl1.Zoom = 5;
            Layer1 = new GMapOverlay("Layer1");
            gMapControl1.Overlays.Add(Layer1);
            Layer2 = new GMapOverlay("Layer2");
            gMapControl1.Overlays.Add(Layer2);
            Layer3 = new GMapOverlay("Layer3");
            gMapControl1.Overlays.Add(Layer3);
            Layer4 = new GMapOverlay("Layer4");
            gMapControl1.Overlays.Add(Layer4);
            Layer5 = new GMapOverlay("Layer5");
            gMapControl1.Overlays.Add(Layer5);

            Layer6 = new GMapOverlay("Layer6");
            gMapControl1.Overlays.Add(Layer6);
            Layer7 = new GMapOverlay("Layer7");
            gMapControl1.Overlays.Add(Layer7);
            Layer8 = new GMapOverlay("Layer8");
            gMapControl1.Overlays.Add(Layer8);

            Layer9 = new GMapOverlay("Layer9");
            gMapControl1.Overlays.Add(Layer9);
            Layer10 = new GMapOverlay("Layer10");
            gMapControl1.Overlays.Add(Layer10);
            Layer11 = new GMapOverlay("Layer11");
            gMapControl1.Overlays.Add(Layer11);
            Layer12 = new GMapOverlay("Layer12");
            gMapControl1.Overlays.Add(Layer12);
            Layer13 = new GMapOverlay("Layer13");
            gMapControl1.Overlays.Add(Layer13);
            Layer14 = new GMapOverlay("Layer14");
            gMapControl1.Overlays.Add(Layer14);
            Layer15 = new GMapOverlay("Layer15");
            gMapControl1.Overlays.Add(Layer15);

            Layer16 = new GMapOverlay("Layer16");
            gMapControl1.Overlays.Add(Layer16);
            Layer17 = new GMapOverlay("Layer17");
            gMapControl1.Overlays.Add(Layer17);

            Layer20 = new GMapOverlay("Layer20");
            gMapControl1.Overlays.Add(Layer20);
            Nodes = new List<TNode>();

            ConnectionString = @"Server=" + Server_Name + "; Database=" + DataBase_Name + "; Trusted_Connection=True;";
            //ConnectionString = @"Server=" + Server_Name + "; Database=" + DataBase_Name + "; User ID=cwpcApp;Password=cwpcApp@830625#Ahmad; Max Pool Size=10;";
            connection = new SqlConnection(ConnectionString);
            connection.Open();


            //string MAP_DATA = @"select * from ARAS_SMART_MAP WHERE [PROVINCE EN]='ILAM'";
            //SqlCommand MAP_DATA_Quary = new SqlCommand(MAP_DATA, connection);
            //MAP_DATA_Quary.CommandTimeout = 0;
            //MAP_DATA_Quary.ExecuteNonQuery();
            //ARAS_DATA_Table = new DataTable();
            //SqlDataAdapter dataAdapter_MAP_DATA_Table = new SqlDataAdapter(MAP_DATA_Quary);
            //dataAdapter_MAP_DATA_Table.Fill(ARAS_DATA_Table);



            //var db = new DataClasses1DataContext();
            //var Q = db.ARAS_SMART_MAPs;

            //  var ARAS_DATA_Table1 = (from C in Q.AsEnumerable() select C).OrderBy(x => x.CELLNAME).ToList();
            //// var ARAS_DATA_Table1 = (from C1 in Q.AsEnumerable() select C1.LOCATION).ToList();
            //ARAS_DATA_Table = ConvertToDataTable(ARAS_DATA_Table1);




            //var query = (from C2 in db.ARAS_SMART_MAPs.AsEnumerable()
            //             join TBL_2G in Code_Sites_2G_Table.AsEnumerable() on C2.LOCATION equals TBL_2G.Field<string>("Site")
            //             select new
            //             {
            //                 C2.LOCATION,
            //                 BSC = TBL_2G.Field<string>("BSC")  //wait
            //             }).ToList();


            int tt = 0;






        }





        private void gMapControl1_MouseDoubleClick(object sender, MouseEventArgs e)
        {


            // Make Line
            if (checkBox12.Checked == true)
            {
                if (e.Button == MouseButtons.Left)
                {
                    double lat = gMapControl1.FromLocalToLatLng(e.X, e.Y).Lat;
                    double lng = gMapControl1.FromLocalToLatLng(e.X, e.Y).Lng;

                    if (points1.Count == 2)
                    {
                        points1.Clear();
                        Layer16.Markers.Clear();
                        polyOverlay1.Polygons.Clear();
                        gMapControl1.Overlays.Add(polyOverlay1);

                    }

                    points1.Add(new PointLatLng(lat, lng));

                    GMapMarker marker1 = new GMarkerGoogle(new PointLatLng(lat, lng), GMarkerGoogleType.pink_dot);
                    Layer16.Markers.Add(marker1);
                }

            }




            // Make Polygon
            if (checkBox13.Checked == true)
            {
                if (e.Button == MouseButtons.Left)
                {
                    double lat = gMapControl1.FromLocalToLatLng(e.X, e.Y).Lat;
                    double lng = gMapControl1.FromLocalToLatLng(e.X, e.Y).Lng;


                    if (Polygoner == points2.Count)
                    {
                        points2.Clear();
                        Layer17.Markers.Clear();
                        polyOverlay2.Polygons.Clear();
                        gMapControl1.Overlays.Add(polyOverlay2);
                    }

                    points2.Add(new PointLatLng(lat, lng));

                    GMapMarker marker2 = new GMarkerGoogle(new PointLatLng(lat, lng), GMarkerGoogleType.pink_dot);
                    Layer17.Markers.Add(marker2);


                    if (lat >= lat_max)
                    {
                        lat_max = lat;
                    }
                    if (lng >= lng_max)
                    {
                        lng_max = lng;
                    }
                    if (lat <= lat_min)
                    {
                        lat_min = lat;
                    }
                    if (lng <= lng_min)
                    {
                        lng_min = lng;
                    }

                }

            }


        }



        // Delete Line
        private void checkBox12_CheckedChanged(object sender, EventArgs e)
        {
            points1 = new List<PointLatLng>();
            if (checkBox12.Checked == true)
            {
                groupBox3.BackColor = Color.Green;
                checkBox13.Checked = false;
            }
            if (checkBox12.Checked == false)
            {
                textBox1.Clear();
                GMapPolygon polygon1 = new GMapPolygon(points1, "mypolygon1");
                groupBox3.BackColor = Color.DarkOrange;
                Layer16.Markers.Clear();
                polyOverlay1.Polygons.Clear();
                gMapControl1.Overlays.Add(polyOverlay1);
                points1.Clear();
            }
        }


        // Delete Polygon
        private void checkBox13_CheckedChanged(object sender, EventArgs e)
        {
            points2 = new List<PointLatLng>();
            if (checkBox13.Checked == true)
            {
                groupBox4.BackColor = Color.Green;
                checkBox12.Checked = false;
            }
            if (checkBox13.Checked == false)
            {
                GMapPolygon polygon2 = new GMapPolygon(points2, "mypolygon2");
                groupBox4.BackColor = Color.DarkOrange;
                Layer17.Markers.Clear();
                polyOverlay2.Polygons.Clear();
                gMapControl1.Overlays.Add(polyOverlay2);
                points2.Clear();


                lat_max = 0;
                lng_max = 0;

                lat_min = 100;
                lng_min = 100;

            }
        }


        // Draw Line
        private void button7_Click(object sender, EventArgs e)
        {


            polyOverlay1 = new GMapOverlay("polygons1");
            GMapPolygon polygon1 = new GMapPolygon(points1, "mypolygon1");
            polygon1.Fill = new SolidBrush(Color.FromArgb(50, Color.Red));
            polygon1.Stroke = new Pen(Color.Red, 1);
            polyOverlay1.Polygons.Add(polygon1);
            gMapControl1.Overlays.Add(polyOverlay1);


            double Lat1 = points1[0].Lat;
            double Long1 = points1[0].Lng;
            double Lat2 = points1[1].Lat;
            double Long2 = points1[1].Lng;


            double dLat = (Lat2 - Lat1) * Math.PI / 180;
            double dLong = (Long2 - Long1) * Math.PI / 180;

            double a = Math.Sin(dLat / 2) * Math.Sin(dLat / 2) + Math.Cos(Lat1 * Math.PI / 180) * Math.Cos(Lat2 * Math.PI / 180) * Math.Sin(dLong / 2) * Math.Sin(dLong / 2);
            double c = 2 * Math.Atan2(Math.Sqrt(a), Math.Sqrt(1 - a));
            double Dist = 6371 * c;
            string Dist1 = String.Format("{0:0.00}", Dist);
            textBox1.Text = Dist1 + " km";

            //gMapControl1.Zoom = gMapControl1.Zoom + 1;
            //gMapControl1.Zoom = gMapControl1.Zoom - 1;
        }


        // Draw Polygon
        private void button8_Click(object sender, EventArgs e)
        {
            polyOverlay2 = new GMapOverlay("polygons2");
            GMapPolygon polygon2 = new GMapPolygon(points2, "mypolygon2");
            polygon2.Fill = new SolidBrush(Color.FromArgb(50, Color.Red));
            polygon2.Stroke = new Pen(Color.Red, 1);
            polyOverlay2.Polygons.Add(polygon2);
            gMapControl1.Overlays.Add(polyOverlay2);
            Polygoner = points2.Count();

            points2 = new List<PointLatLng>();
            Polygoner = 0;
        }




        private void button1_Click(object sender, EventArgs e)
        {
            gMapControl1.Zoom = gMapControl1.Zoom + 1;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            gMapControl1.Zoom = gMapControl1.Zoom - 1;

        }

        private void button3_Click(object sender, EventArgs e)
        {
            double panFactor = 15 / (gMapControl1.Zoom * 10);
            gMapControl1.Position = new PointLatLng(gMapControl1.Position.Lat, gMapControl1.Position.Lng - panFactor);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            double panFactor = 15 / (gMapControl1.Zoom * 10);
            gMapControl1.Position = new PointLatLng(gMapControl1.Position.Lat + panFactor, gMapControl1.Position.Lng);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            double panFactor = 15 / (gMapControl1.Zoom * 10);
            gMapControl1.Position = new PointLatLng(gMapControl1.Position.Lat - panFactor, gMapControl1.Position.Lng);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            double panFactor = 15 / (gMapControl1.Zoom * 10);
            gMapControl1.Position = new PointLatLng(gMapControl1.Position.Lat, gMapControl1.Position.Lng + panFactor);
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            checkBox11.Checked = true;
            Layer15.Markers.Clear();
            string Node = comboBox4.SelectedItem.ToString();
            string tech = "";
            if (Node.Substring(0, 1) == "B")
            {
                tech = "2G";
            }
            if (Node.Substring(0, 1) != "B")
            {
                tech = "3G";
            }
            for (int t = 0; t < Nodes.Count; t++)
            {
                int RAN_Finder = 0;
                if (tech == "2G")
                {
                    RAN_Finder = Convert.ToInt16(string.Compare(Nodes[t].BSC.ToString(), comboBox4.SelectedItem.ToString()));
                }
                if (tech == "3G")
                {
                    RAN_Finder = Convert.ToInt16(string.Compare(Nodes[t].RNC.ToString(), comboBox4.SelectedItem.ToString()));
                }
                if (RAN_Finder == 0)
                {
                    GMapMarker marker = new GMarkerGoogle(new PointLatLng(Nodes[t].Lat, Nodes[t].Lon), GMarkerGoogleType.orange);
                    marker.ToolTipText = Nodes[t].Site;
                    Layer15.Markers.Add(marker);
                }
            }
        }

        private void gMapControl1_Load(object sender, EventArgs e)
        {

        }

        private void checkBox11_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox11.Checked == false)
            {
                Layer15.Markers.Clear();
            }
        }


        // Export 
        private void button9_Click(object sender, EventArgs e)
        {

            MessageBox.Show("This Part needs to be completed...");

            DataTable Export_DATA_Table = new DataTable();

            Export_DATA_Table.Columns.Add("Province", typeof(string));
            Export_DATA_Table.Columns.Add("City", typeof(string));
            Export_DATA_Table.Columns.Add("Site", typeof(string));
            Export_DATA_Table.Columns.Add("Cell", typeof(string));
            Export_DATA_Table.Columns.Add("Lat", typeof(string));
            Export_DATA_Table.Columns.Add("Long", typeof(string));
            Export_DATA_Table.Columns.Add("Azimuth", typeof(string));
            Export_DATA_Table.Columns.Add("Technology", typeof(string));
            Export_DATA_Table.Columns.Add("Band", typeof(string));
            Export_DATA_Table.Columns.Add("Node", typeof(string));
            Export_DATA_Table.Columns.Add("Site Type", typeof(string));




            for (int k = 0; k < ARAS_DATA_Table.Rows.Count; k++)
            {
                double LAT = 0;
                double LONG = 0;
                if (ARAS_DATA_Table.Rows[k][4].ToString() != "")
                {
                    LAT = Convert.ToDouble(ARAS_DATA_Table.Rows[k][4]);
                }
                if (ARAS_DATA_Table.Rows[k][5].ToString() != "")
                {
                    LONG = Convert.ToDouble(ARAS_DATA_Table.Rows[k][5]);
                }
                if (LAT >= lat_min && LAT <= lat_max && LONG >= lng_min && LONG <= lng_max)
                {
                    string Province = ARAS_DATA_Table.Rows[k][0].ToString();
                    string City = ARAS_DATA_Table.Rows[k][1].ToString();
                    string Site = ARAS_DATA_Table.Rows[k][3].ToString();
                    string Cell = ARAS_DATA_Table.Rows[k][2].ToString();
                    string LAT1 = ARAS_DATA_Table.Rows[k][4].ToString();
                    string LONG1 = ARAS_DATA_Table.Rows[k][5].ToString();
                    string Azimuth = ARAS_DATA_Table.Rows[k][6].ToString();
                    string Technology = ARAS_DATA_Table.Rows[k][7].ToString();
                    string Band = ARAS_DATA_Table.Rows[k][8].ToString();
                    string Node = ARAS_DATA_Table.Rows[k][9].ToString();
                    string Site_Type = ARAS_DATA_Table.Rows[k][10].ToString();

                    Export_DATA_Table.Rows.Add(Province, City, Site, Cell, LAT1, LONG1, Azimuth, Technology, Band, Node, Site_Type);
                }

             
            }

            XLWorkbook wb = new XLWorkbook();
            wb.Worksheets.Add(Export_DATA_Table, "NPM MAP Export");
            var saveFileDialog = new SaveFileDialog
            {
                FileName = Selected_Province + "_Export",
                Filter = "Excel files|*.xlsx",
                Title = "Save an Excel File"
            };

            saveFileDialog.ShowDialog();

            if (!String.IsNullOrWhiteSpace(saveFileDialog.FileName))
                wb.SaveAs(saveFileDialog.FileName);

            MessageBox.Show("Finished");

        }


        // Upload and Isert ARAS Database

        public string FName = "";
        public IXLWorksheet Source_worksheet = null;
        //public DataTable Data_Table_2G = new DataTable();
        public Excel.Application xlApp { get; set; }
        public Excel.Workbook xlWorkBook { get; set; }
        public Excel.Worksheet Sheet1 { get; set; }
        public Excel.Worksheet Sheet2 { get; set; }
        public Excel.Worksheet Sheet3 { get; set; }
        public Excel.Worksheet Sheet4 { get; set; }
        public Excel.Worksheet Sheet5 { get; set; }
        public Excel.Worksheet Sheet6 { get; set; }
        public Excel.Worksheet Sheet7 { get; set; }
        public Excel.Worksheet Sheet8 { get; set; }

        public DataTable ARAS_Table = new DataTable();
        
        private void button10_Click(object sender, EventArgs e)
        {

            ARAS_Table.Columns.Add("PROVINCE EN", typeof(string));
            ARAS_Table.Columns.Add("CITY EN", typeof(string));
            ARAS_Table.Columns.Add("CELLNAME", typeof(string));
            ARAS_Table.Columns.Add("LOCATION", typeof(string));
            ARAS_Table.Columns.Add("LATITUDE", typeof(float));
            ARAS_Table.Columns.Add("LONGITUDE", typeof(float));
            ARAS_Table.Columns.Add("AZIMUTH", typeof(float));
            ARAS_Table.Columns.Add("Technology", typeof(string));
            ARAS_Table.Columns.Add("Band", typeof(string));
            ARAS_Table.Columns.Add("Node", typeof(string));
            ARAS_Table.Columns.Add("Site_Type", typeof(string));
            ARAS_Table.Columns.Add("COVERAGE_TYPE_OPTIMIZATION", typeof(string));
            ARAS_Table.Columns.Add("SEGMENTATION_OPTIMIZATION", typeof(string));
            ARAS_Table.Columns.Add("Site_Type2", typeof(string));
            ARAS_Table.Columns.Add("BTS Type", typeof(string));
            ARAS_Table.Columns.Add("TOWER TYPE", typeof(string));
            ARAS_Table.Columns.Add("TOWER HEIGHT", typeof(float));
            ARAS_Table.Columns.Add("ANTENNA HEIGHT", typeof(float));
            ARAS_Table.Columns.Add("Antenna Type", typeof(string));



            openFileDialog1.DefaultExt = "xlsx";
            openFileDialog1.Filter = "Excel File|*.xlsx";
            DialogResult result1 = openFileDialog1.ShowDialog();
            string File_Name = openFileDialog1.SafeFileName.ToString();



            if (result1 == DialogResult.OK)
            {
                string file = openFileDialog1.FileName;
                FName = file;
                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(file);


                // Location
                Sheet1 = xlWorkBook.Worksheets[1];
                Excel.Range Location = Sheet1.get_Range("A2", "GB" + Sheet1.UsedRange.Rows.Count);
                object[,] ARAS_Location = (object[,])Location.Value;

                // G900
                Sheet2 = xlWorkBook.Worksheets[3];
                Excel.Range G900 = Sheet2.get_Range("A2", "AJ" + Sheet2.UsedRange.Rows.Count);
                object[,] ARAS_G900 = (object[,])Location.Value;

                // G1800
                Sheet3 = xlWorkBook.Worksheets[4];
                Excel.Range G1800 = Sheet3.get_Range("A2", "AJ" + Sheet3.UsedRange.Rows.Count);
                object[,] ARAS_G1800 = (object[,])Location.Value;

                // U900
                Sheet4 = xlWorkBook.Worksheets[5];
                Excel.Range U900 = Sheet4.get_Range("A2", "AJ" + Sheet4.UsedRange.Rows.Count);
                object[,] ARAS_U900 = (object[,])Location.Value;

                // U2100
                Sheet5 = xlWorkBook.Worksheets[6];
                Excel.Range U2100 = Sheet5.get_Range("A2", "AJ" + Sheet5.UsedRange.Rows.Count);
                object[,] ARAS_U2100 = (object[,])Location.Value;

                // L1800
                Sheet6 = xlWorkBook.Worksheets[7];
                Excel.Range L1800 = Sheet6.get_Range("A2", "AJ" + Sheet6.UsedRange.Rows.Count);
                object[,] ARAS_L1800 = (object[,])Location.Value;

                // L2100
                Sheet7 = xlWorkBook.Worksheets[8];
                Excel.Range L2100 = Sheet7.get_Range("A2", "AJ" + Sheet7.UsedRange.Rows.Count);
                object[,] ARAS_L2100 = (object[,])Location.Value;

                // L2600
                Sheet8 = xlWorkBook.Worksheets[9];
                Excel.Range L2600 = Sheet8.get_Range("A2", "AJ" + Sheet8.UsedRange.Rows.Count);
                object[,] ARAS_L2600 = (object[,])Location.Value;



                int Count = Sheet1.UsedRange.Rows.Count;

                for (int k = 0; k < Count - 1; k++)
                {
                    string Province_Persian = ARAS_Location[k + 1, 2].ToString();
                    string Province_En = "";
                    string City_En = ARAS_Location[k + 1, 10].ToString();
                    string Code_Site= ARAS_Location[k + 1, 4].ToString();



                    int Count_G900 = Sheet2.UsedRange.Rows.Count;

                    for (int i = 0; i < Count_G900 - 1; i++)
                    {
                        string Code_Site_G900= ARAS_G900[i + 1, 4].ToString();
                        if (Code_Site_G900== Code_Site)
                        {

                        }
                    }



                        if (Province_Persian == "آذربایجان غربی") { Province_En = "WEST AZERBAIJAN"; }

//                    ALBORZ
//ARDABIL
//BUSHEHR
//CHAHARMAHAL BAKHTIARI
//EAST AZERBAIJAN
//ESFAHAN
//FARS
//GILAN
//GOLESTAN
//HAMEDAN
//HORMOZGAN
//ILAM
//KERMAN
//KERMANSHAH
//KHORASAN RAZAVI
//KHOUZESTAN
//KOHGILUYEH AND BOYER AHMAD
//KORDESTAN
//LORESTAN
//MARKAZI
//MAZANDARAN
//NORTH KHORASAN
//QAZVIN
//QOM
//SEMNAN
//SISTAN VA BALUCHESTAN
//SOUTH KHORASAN
//TEHRAN

//YAZD
//ZANJAN






                    ARAS_Table.Rows.Add(Province_En, City_En, "", "", 0,0,0, "", "", "", "", "", "", "", "", "",0, 0, "");


                }




                MessageBox.Show("ARAS is Loaded!");






            }





            }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }


}

