// VBConversions Note: VB project level imports
using System.Collections.Generic;
using System;
using Office = Microsoft.Office.Core;
using Microsoft.VisualBasic;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text;
using System.Linq;
// End of VB project level imports

using System.Windows.Forms;


namespace ListeDeGarde
{
	public sealed class PublicConstants
	{
		
		// --------------------DB Address
		//Public const CONSTFILEADDRESS As String = "C:\Users\Martin\Documents\Scheduling Mira\ListesDeGarde.accdb"
		public static string CONSTFILEADDRESS;
		//----------------------Table Name Mapping
		
		//DB Access constants for Table SShiftType
		public const string TABLE_shiftType = "TABLE_shiftType";
		public const string SQLShiftStart = "ShiftStart";
		public const string SQLShiftStop = "ShiftStop";
		public const string SQLShiftType = "ShiftType";
		public const string SQLActive = "Active";
		public const string SQLDescription = "Description";
		public const string SQLLundi = "Lundi";
		public const string SQLMardi = "Mardi";
		public const string SQLMercredi = "Mercredi";
		public const string SQLJeudi = "Jeudi";
		public const string SQLVendredi = "Vendredi";
		public const string SQLSamedi = "Samedi";
		public const string SQLDimanche = "Dimanche";
		public const string SQLFerie = "Ferie";
		public const string SQLCompilation = "Compilation";
		public const string SQLOrder = "The_Order";
		
		
		
		//Public Const SQLEffectiveStart = "EffectiveStart"
		//Public Const SQLEffectiveEnd = "EffectiveEnd"
		
		//DB Access constants for Table SDoc
		public const string TABLE_Doc = "TABLE_Doc";
		public const string SQLFirstName = "FirstName";
		public const string SQLLastName = "LastName";
		public const string SQLInitials = "Initials";
		//Public Const SQLActive As String = "Active"
		public const string SQLEffectiveStart = "EffectiveStart";
		public const string SQLEffectiveEnd = "EffectiveEnd";
		public const string SQLShift1 = "Shift1";
		public const string SQLShift2 = "Shift2";
		public const string SQLShift3 = "Shift3";
		public const string SQLShift4 = "Shift4";
		public const string SQLShift5 = "Shift5";
		public const string SQLUrgenceTog = "Urgence";
		public const string SQLHospitTog = "Hospit";
		public const string SQLSoinsTog = "Soins";
		public const string SQLNuitsTog = "Nuits";
		public const string SQLVersion = "Version";
		
		//DB Access constants for Table ScheduleData
		public const string TABLE_ScheduleData = "TABLE_ScheduleData";
		public const string SQLDate = "aDate";
		//Public Const SQLShiftType As String = "ShiftType"
		//Public Const SQLInitials As String = "Initials"
		
		//DB Access constants for Table_NonDispo
		public const string Table_NonDispo = "Table_NonDispo";
		public const string SQLDateStart = "aDateStart";
		public const string SQLTimeStart = "aTimeStart";
		public const string SQLDateStop = "aDateStop";
		public const string SQLTimeStop = "aTimeStop";
		//Public Const SQLInitials As String = "Initials"
		
		//-----------------------Default Values
		public const long DEFAULTDATE = 29221;
		public const long kTicksToDays = 864000000000;
	}
	
	public sealed class PublicStructures
	{
		//------------------------------------------------------------------------------
		//                   =========== Structures ===========
		//------------------------------------------------------------------------------
		
		public struct T_DBRefTypeS
		{
			public string theSQLName;
			public string theValue;
		}
		
		public struct T_DBRefTypeA
		{
			public string theSQLName;
			public PublicEnums.Availability theValue;
		}
		
		public struct T_DBRefTypeL
		{
			public string theSQLName;
			public long theValue;
		}
		
		public struct T_DBRefTypeI
		{
			public string theSQLName;
			public int theValue;
		}
		
		public struct T_DBRefTypeB
		{
			public string theSQLName;
			public bool theValue;
		}
		
		public struct T_DBRefTypeD
		{
			public string theSQLName;
			public DateTime theValue;
		}
	}
	
	
	public sealed class PublicEnums
	{
		public enum Availability
		{
			Assigne = 0,
			Dispo = 1,
			NonDispoPermanente = 2,
			NonDispoTemporaire = 3,
			SurUtilise = 4,
			AssigneSpecial = 5
		}
		
		public enum Weekdays
		{
			monday = 1,
			tuesday = 2,
			wednesday = 3,
			Thursay = 4,
			Friday = 5,
			Saturday = 6,
			Sunday = 7
		}
		
		public enum EnumWhereSubClause
		{
			EW_None,
			EW_begin,
			EW_end
		}
		
	}
	
	public sealed class MyGlobals
	{
		
		public static SShiftType globalShiftTypes;
		//Public theNonDispoList As Collection
		public static string theList;
		public static Collection theRangeCollection;
		
		// Conenction Global variables
		public static ADODB.Connection cnn = new ADODB.Connection(); //Connection object definition
		//Public rs As New ADODB.Recordset    'recordset object definition
		//Public strSQL As String             'Query String
		public static string FileAddress; //String to contain Excell database path and filename
		//Public GlobalDBAccessClass As DBAccessClass1
		public static Settings1 MySettingsGlobal;
		public static string[] daystrings = new string[] {"dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"};
		
		public static string[] monthstrings = new string[] {"janvier", "février", "mars", "avril", "mai", "juin", "juillet", "aout", "septembre", "octobre", "novembre", "décembre"};
		public static string[] yearstrings = new string[] {"2014", "2015", "2016", "2017"};
		
		public static string[] HoursStrings = new string[] {"00", "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23"};
		
		public static string[] MinutesStrings = new string[] {"00", "05", "10", "15", "20", "25", "30", "35", "40", "45", "50", "55"};
		
		
	}
	
	public sealed class GlobalFunctions
	{
		
		public static void LoadDatabaseFileLocation()
		{
			OpenFileDialog filedialog = new OpenFileDialog();
			filedialog.Title = "Select Location of database file";
			filedialog.InitialDirectory = "";
			filedialog.Filter = "Access DB files (*.accdb)|*.accdb";
			
			filedialog.RestoreDirectory = true;
			if (filedialog.ShowDialog() == DialogResult.OK)
			{
				MyGlobals.MySettingsGlobal.DataBaseLocation = filedialog.FileName;
			}
			MyGlobals.MySettingsGlobal.Save();
			
		}
		
		public static string cAccessDateStr(DateTime theDate)
		{
			
			return "#" + theDate.ToString("yyyy-M-d") + "#";
			
		}
		
	}
}
