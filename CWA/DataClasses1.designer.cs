﻿#pragma warning disable 1591
//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace CWA
{
	using System.Data.Linq;
	using System.Data.Linq.Mapping;
	using System.Data;
	using System.Collections.Generic;
	using System.Reflection;
	using System.Linq;
	using System.Linq.Expressions;
	using System.ComponentModel;
	using System;
	
	
	[global::System.Data.Linq.Mapping.DatabaseAttribute(Name="Performance_NAK")]
	public partial class DataClasses1DataContext : System.Data.Linq.DataContext
	{
		
		private static System.Data.Linq.Mapping.MappingSource mappingSource = new AttributeMappingSource();
		
    #region Extensibility Method Definitions
    partial void OnCreated();
    #endregion
		
		public DataClasses1DataContext() : 
				base(global::CWA.Properties.Settings.Default.Performance_NAKConnectionString1, mappingSource)
		{
			OnCreated();
		}
		
		public DataClasses1DataContext(string connection) : 
				base(connection, mappingSource)
		{
			OnCreated();
		}
		
		public DataClasses1DataContext(System.Data.IDbConnection connection) : 
				base(connection, mappingSource)
		{
			OnCreated();
		}
		
		public DataClasses1DataContext(string connection, System.Data.Linq.Mapping.MappingSource mappingSource) : 
				base(connection, mappingSource)
		{
			OnCreated();
		}
		
		public DataClasses1DataContext(System.Data.IDbConnection connection, System.Data.Linq.Mapping.MappingSource mappingSource) : 
				base(connection, mappingSource)
		{
			OnCreated();
		}
		
		public System.Data.Linq.Table<ARAS_SMART_MAP> ARAS_SMART_MAPs
		{
			get
			{
				return this.GetTable<ARAS_SMART_MAP>();
			}
		}

        public object ARAS_DATA_Table { get; internal set; }
    }
	
	[global::System.Data.Linq.Mapping.TableAttribute(Name="dbo.ARAS_SMART_MAP")]
	public partial class ARAS_SMART_MAP
	{
		
		private string _REGION;
		
		private string _REGION_EN;
		
		private string _PROVINCE;
		
		private string _PROVINCE_EN;
		
		private string _CITY_ON_COUNTRY_DIVISIONS;
		
		private string _CITY_EN;
		
		private string _SITENAME;
		
		private string _CELLNAME;
		
		private string _LOCATION;
		
		private System.Nullable<double> _LAC;
		
		private string _WLL_MCI;
		
		private System.Nullable<double> _CELLID;
		
		private System.Nullable<double> _LATITUDE;
		
		private System.Nullable<double> _LONGITUDE;
		
		private string _GSM_CELL_TYPE;
		
		private System.Nullable<double> _AZIMUTH;
		
		private string _ADDRESS;
		
		private System.Nullable<double> _ANTENNAHEIGHT;
		
		private System.Nullable<double> _MECHANICALTILT;
		
		private System.Nullable<double> _ELECTRICALTILT;
		
		private string _ANTENNATYPEID;
		
		private string _GAIN;
		
		private string _SECTOR_COVERAGE;
		
		private string _LOSS;
		
		private string _H_BEAMWIDTH;
		
		private string _V_BEAMWIDTH;
		
		private string _FRONTBACKRATIO;
		
		private string _COVERAGE_TYPE_OPTIMIZATION;
		
		private string _SEGMENTATION_OPTIMIZATION;
		
		public ARAS_SMART_MAP()
		{
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_REGION", DbType="NVarChar(255)")]
		public string REGION
		{
			get
			{
				return this._REGION;
			}
			set
			{
				if ((this._REGION != value))
				{
					this._REGION = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Name="[REGION EN]", Storage="_REGION_EN", DbType="NVarChar(255)")]
		public string REGION_EN
		{
			get
			{
				return this._REGION_EN;
			}
			set
			{
				if ((this._REGION_EN != value))
				{
					this._REGION_EN = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_PROVINCE", DbType="NVarChar(255)")]
		public string PROVINCE
		{
			get
			{
				return this._PROVINCE;
			}
			set
			{
				if ((this._PROVINCE != value))
				{
					this._PROVINCE = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Name="[PROVINCE EN]", Storage="_PROVINCE_EN", DbType="NVarChar(255)")]
		public string PROVINCE_EN
		{
			get
			{
				return this._PROVINCE_EN;
			}
			set
			{
				if ((this._PROVINCE_EN != value))
				{
					this._PROVINCE_EN = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Name="[CITY ON COUNTRY DIVISIONS]", Storage="_CITY_ON_COUNTRY_DIVISIONS", DbType="NVarChar(255)")]
		public string CITY_ON_COUNTRY_DIVISIONS
		{
			get
			{
				return this._CITY_ON_COUNTRY_DIVISIONS;
			}
			set
			{
				if ((this._CITY_ON_COUNTRY_DIVISIONS != value))
				{
					this._CITY_ON_COUNTRY_DIVISIONS = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Name="[CITY EN]", Storage="_CITY_EN", DbType="NVarChar(255)")]
		public string CITY_EN
		{
			get
			{
				return this._CITY_EN;
			}
			set
			{
				if ((this._CITY_EN != value))
				{
					this._CITY_EN = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_SITENAME", DbType="NVarChar(255)")]
		public string SITENAME
		{
			get
			{
				return this._SITENAME;
			}
			set
			{
				if ((this._SITENAME != value))
				{
					this._SITENAME = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_CELLNAME", DbType="NVarChar(255)")]
		public string CELLNAME
		{
			get
			{
				return this._CELLNAME;
			}
			set
			{
				if ((this._CELLNAME != value))
				{
					this._CELLNAME = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_LOCATION", DbType="NVarChar(255)")]
		public string LOCATION
		{
			get
			{
				return this._LOCATION;
			}
			set
			{
				if ((this._LOCATION != value))
				{
					this._LOCATION = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_LAC", DbType="Float")]
		public System.Nullable<double> LAC
		{
			get
			{
				return this._LAC;
			}
			set
			{
				if ((this._LAC != value))
				{
					this._LAC = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Name="[WLL/MCI]", Storage="_WLL_MCI", DbType="NVarChar(255)")]
		public string WLL_MCI
		{
			get
			{
				return this._WLL_MCI;
			}
			set
			{
				if ((this._WLL_MCI != value))
				{
					this._WLL_MCI = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_CELLID", DbType="Float")]
		public System.Nullable<double> CELLID
		{
			get
			{
				return this._CELLID;
			}
			set
			{
				if ((this._CELLID != value))
				{
					this._CELLID = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_LATITUDE", DbType="Float")]
		public System.Nullable<double> LATITUDE
		{
			get
			{
				return this._LATITUDE;
			}
			set
			{
				if ((this._LATITUDE != value))
				{
					this._LATITUDE = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_LONGITUDE", DbType="Float")]
		public System.Nullable<double> LONGITUDE
		{
			get
			{
				return this._LONGITUDE;
			}
			set
			{
				if ((this._LONGITUDE != value))
				{
					this._LONGITUDE = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_GSM_CELL_TYPE", DbType="NVarChar(255)")]
		public string GSM_CELL_TYPE
		{
			get
			{
				return this._GSM_CELL_TYPE;
			}
			set
			{
				if ((this._GSM_CELL_TYPE != value))
				{
					this._GSM_CELL_TYPE = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_AZIMUTH", DbType="Float")]
		public System.Nullable<double> AZIMUTH
		{
			get
			{
				return this._AZIMUTH;
			}
			set
			{
				if ((this._AZIMUTH != value))
				{
					this._AZIMUTH = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_ADDRESS", DbType="NVarChar(255)")]
		public string ADDRESS
		{
			get
			{
				return this._ADDRESS;
			}
			set
			{
				if ((this._ADDRESS != value))
				{
					this._ADDRESS = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_ANTENNAHEIGHT", DbType="Float")]
		public System.Nullable<double> ANTENNAHEIGHT
		{
			get
			{
				return this._ANTENNAHEIGHT;
			}
			set
			{
				if ((this._ANTENNAHEIGHT != value))
				{
					this._ANTENNAHEIGHT = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_MECHANICALTILT", DbType="Float")]
		public System.Nullable<double> MECHANICALTILT
		{
			get
			{
				return this._MECHANICALTILT;
			}
			set
			{
				if ((this._MECHANICALTILT != value))
				{
					this._MECHANICALTILT = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_ELECTRICALTILT", DbType="Float")]
		public System.Nullable<double> ELECTRICALTILT
		{
			get
			{
				return this._ELECTRICALTILT;
			}
			set
			{
				if ((this._ELECTRICALTILT != value))
				{
					this._ELECTRICALTILT = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_ANTENNATYPEID", DbType="NVarChar(255)")]
		public string ANTENNATYPEID
		{
			get
			{
				return this._ANTENNATYPEID;
			}
			set
			{
				if ((this._ANTENNATYPEID != value))
				{
					this._ANTENNATYPEID = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_GAIN", DbType="NVarChar(255)")]
		public string GAIN
		{
			get
			{
				return this._GAIN;
			}
			set
			{
				if ((this._GAIN != value))
				{
					this._GAIN = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Name="[SECTOR COVERAGE]", Storage="_SECTOR_COVERAGE", DbType="NVarChar(255)")]
		public string SECTOR_COVERAGE
		{
			get
			{
				return this._SECTOR_COVERAGE;
			}
			set
			{
				if ((this._SECTOR_COVERAGE != value))
				{
					this._SECTOR_COVERAGE = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_LOSS", DbType="NVarChar(255)")]
		public string LOSS
		{
			get
			{
				return this._LOSS;
			}
			set
			{
				if ((this._LOSS != value))
				{
					this._LOSS = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_H_BEAMWIDTH", DbType="NVarChar(255)")]
		public string H_BEAMWIDTH
		{
			get
			{
				return this._H_BEAMWIDTH;
			}
			set
			{
				if ((this._H_BEAMWIDTH != value))
				{
					this._H_BEAMWIDTH = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_V_BEAMWIDTH", DbType="NVarChar(255)")]
		public string V_BEAMWIDTH
		{
			get
			{
				return this._V_BEAMWIDTH;
			}
			set
			{
				if ((this._V_BEAMWIDTH != value))
				{
					this._V_BEAMWIDTH = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_FRONTBACKRATIO", DbType="NVarChar(255)")]
		public string FRONTBACKRATIO
		{
			get
			{
				return this._FRONTBACKRATIO;
			}
			set
			{
				if ((this._FRONTBACKRATIO != value))
				{
					this._FRONTBACKRATIO = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_COVERAGE_TYPE_OPTIMIZATION", DbType="NVarChar(255)")]
		public string COVERAGE_TYPE_OPTIMIZATION
		{
			get
			{
				return this._COVERAGE_TYPE_OPTIMIZATION;
			}
			set
			{
				if ((this._COVERAGE_TYPE_OPTIMIZATION != value))
				{
					this._COVERAGE_TYPE_OPTIMIZATION = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_SEGMENTATION_OPTIMIZATION", DbType="NVarChar(255)")]
		public string SEGMENTATION_OPTIMIZATION
		{
			get
			{
				return this._SEGMENTATION_OPTIMIZATION;
			}
			set
			{
				if ((this._SEGMENTATION_OPTIMIZATION != value))
				{
					this._SEGMENTATION_OPTIMIZATION = value;
				}
			}
		}

        public object DataRowExtensions { get; internal set; }
    }
}
#pragma warning restore 1591
