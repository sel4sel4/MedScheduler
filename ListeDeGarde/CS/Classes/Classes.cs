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

using System.Diagnostics;
using System.Windows.Forms;
using System.Configuration;


namespace ListeDeGarde
{
	public class SYear
	{
		private int pYear;
		private Collection pMonths;
		
		public int Year
		{
			get
			{
				return pYear;
			}
		}
		
		public Collection Months
		{
			get
			{
				return pMonths;
			}
		}
		
		public SYear(int aYear)
		{
			pYear = aYear;
			pMonths = new Collection();
			for (var x = 1; x <= 12; x++)
			{
				SMonth theMonth = default(SMonth);
				theMonth = new SMonth(System.Convert.ToInt32(x), aYear);
				pMonths.Add(theMonth, x.ToString(), null, null);
			}
		}
		
	}
	
	public class SMonth
	{
		private int pYear;
		private int pMonth;
		private Collection pDays;
		private Collection pShiftypes;
		private Collection pDocList;
		
		public int Year
		{
			get
			{
				return pYear;
			}
		}
		public int Month
		{
			get
			{
				return pMonth;
			}
		}
		public Collection Days
		{
			get
			{
				return pDays;
			}
		}
		public Collection ShiftTypes
		{
			get
			{
				return pShiftypes;
			}
		}
		public Collection DocList
		{
			get
			{
				return pDocList;
			}
		}
		public SMonth(int aMonth, int aYear)
		{
			pShiftypes = SShiftType.loadShiftTypesFromDBPerMonth(aMonth, aYear);
			pDocList = SDoc.LoadAllDocsPerMonth(aYear, aMonth);
			int theDaysInMonth = DateTime.DaysInMonth(aYear, aMonth);
			pYear = aYear;
			pMonth = aMonth;
			pDays = new Collection();
			for (var x = 1; x <= theDaysInMonth; x++)
			{
				SDay theDay = default(SDay);
				theDay = new SDay(System.Convert.ToInt32(x), aMonth, aYear, this);
				pDays.Add(theDay, x.ToString(), null, null);
			}
		}
		
	}
	
	public class SDay
	{
		private DateTime pDate; //uniqueID
		private Collection pShifts;
		private SMonth pMonth;
		
		public Collection Shifts
		{
			get
			{
				return pShifts;
			}
		}
		public DateTime theDate
		{
			get
			{
				return pDate;
			}
		}
		public SMonth Month
		{
			get
			{
				return pMonth;
			}
		}
		
		public SDay(int aDay, int aMonth, int aYear, SMonth CMonth)
		{
			pDate = new DateTime(aYear, aMonth, aDay);
			pMonth = CMonth;
			pShifts = new Collection();
			bool addShift = false;
			//populate the shift collection by cycling through
			//the active SShiftTypes collection
			SShiftType aShiftType = default(SShiftType);
			int theCounter = 1;
			foreach (SShiftType tempLoopVar_aShiftType in pMonth.ShiftTypes)
			{
				aShiftType = tempLoopVar_aShiftType;
				if (aShiftType.Active)
				{
					switch (pDate.DayOfWeek)
					{
						case DayOfWeek.Monday:
							if (aShiftType.Lundi == true)
							{
								addShift = true;
							}
							else
							{
								addShift = false;
							}
							break;
						case DayOfWeek.Tuesday:
							if (aShiftType.Mardi == true)
							{
								addShift = true;
							}
							else
							{
								addShift = false;
							}
							break;
						case DayOfWeek.Wednesday:
							if (aShiftType.Mercredi == true)
							{
								addShift = true;
							}
							else
							{
								addShift = false;
							}
							break;
						case DayOfWeek.Thursday:
							if (aShiftType.Jeudi == true)
							{
								addShift = true;
							}
							else
							{
								addShift = false;
							}
							break;
						case DayOfWeek.Friday:
							if (aShiftType.Vendredi == true)
							{
								addShift = true;
							}
							else
							{
								addShift = false;
							}
							break;
						case DayOfWeek.Saturday:
							if (aShiftType.Samedi == true)
							{
								addShift = true;
							}
							else
							{
								addShift = false;
							}
							break;
						case DayOfWeek.Sunday:
							if (aShiftType.Dimanche == true)
							{
								addShift = true;
							}
							else
							{
								addShift = false;
							}
							break;
					}
					if (addShift == true)
					{
						SShift theShift = new SShift(aShiftType.ShiftType, pDate, aShiftType.ShiftStart, aShiftType.ShiftStop, aShiftType.Description, this);
						
						pShifts.Add(theShift, aShiftType.ShiftType.ToString(), null, null);
					}
				}
			}
			
		}
		
	}
	
	public class SShift
	{
		private int pShiftStart;
		private int pShiftStop;
		private int pShiftType;
		private string pDescription;
		private string pDoc;
		private Collection pDocAvailabilities;
		private DateTime pDate;
		private int pStatus;
		private Excel.Range pRange;
		private SDay pDay;
		
		public string Doc
		{
			get
			{
				return pDoc;
			}
			set
			{
				pDoc = value;
			}
		}
		public int Status
		{
			get
			{
				return pStatus;
			}
			set
			{
				pStatus = value;
			}
		}
		public string Description
		{
			get
			{
				return pDescription;
			}
		}
		public Excel.Range aRange
		{
			get
			{
				return pRange;
			}
			set
			{
				pRange = value;
			}
		}
		public DateTime aDate
		{
			get
			{
				return pDate;
			}
		}
		public int ShiftType
		{
			get
			{
				return pShiftType;
			}
		}
		public int ShiftStart
		{
			get
			{
				return pShiftStart;
			}
		}
		public int ShiftStop
		{
			get
			{
				return pShiftStop;
			}
		}
		public Collection DocAvailabilities
		{
			get
			{
				return pDocAvailabilities;
			}
			set
			{
				pDocAvailabilities = value;
			}
		}
		
		public SShift(int aShiftType, DateTime aDate, int aShiftStart, int aShiftStop, string aDescription, SDay aDay)
		{
			pDate = aDate;
			pShiftType = aShiftType;
			pShiftStart = aShiftStart;
			pShiftStop = aShiftStop;
			pStatus = 0; // for empty
			pDescription = aDescription;
			pDay = aDay;
			
			pDocAvailabilities = new Collection();
			SDocAvailable theSDocAvailable = default(SDocAvailable);
			SDoc aSDoc = default(SDoc);
			PublicEnums.Availability theDispo = default(PublicEnums.Availability);
			foreach (SDoc tempLoopVar_aSDoc in pDay.Month.DocList)
			{
				aSDoc = tempLoopVar_aSDoc;
				//conditional code to make doc unavailable if shift is not active for the doc
				switch (aShiftType)
				{
					case 1: //urgence
					case 2:
					case 3:
					case 4:
						if (aSDoc.UrgenceTog == false)
						{
							theDispo = PublicEnums.Availability.NonDispoPermanente;
						}
						else
						{
							theDispo = PublicEnums.Availability.Dispo;
						}
						break;
					case 5: //urgence nuit
						if (aSDoc.UrgenceTog == false || aSDoc.NuitsTog == false)
						{
							theDispo = PublicEnums.Availability.NonDispoPermanente;
						}
						else
						{
							theDispo = PublicEnums.Availability.Dispo;
						}
						break;
						
					case 6: //hospit
						if (aSDoc.HospitTog == false)
						{
							theDispo = PublicEnums.Availability.NonDispoPermanente;
						}
						else
						{
							theDispo = PublicEnums.Availability.Dispo;
						}
						break;
					case 7: //soins
						if (aSDoc.SoinsTog == false)
						{
							theDispo = PublicEnums.Availability.NonDispoPermanente;
						}
						else
						{
							theDispo = PublicEnums.Availability.Dispo;
						}
						break;
					default:
						theDispo = PublicEnums.Availability.Dispo;
						break;
				}
				theSDocAvailable = new SDocAvailable(aSDoc.Initials, theDispo, pDate, pShiftType);
				pDocAvailabilities.Add(theSDocAvailable, aSDoc.Initials, null, null);
			}
			
		}
		
	}
	
	public class SShiftType
	{
		private PublicStructures.T_DBRefTypeI pShiftStart;
		private PublicStructures.T_DBRefTypeI pShiftStop;
		private PublicStructures.T_DBRefTypeI pShiftType;
		private PublicStructures.T_DBRefTypeB pActive;
		private PublicStructures.T_DBRefTypeB pCompilation;
		private PublicStructures.T_DBRefTypeS pDescription;
		private PublicStructures.T_DBRefTypeI pVersion;
		private PublicStructures.T_DBRefTypeB pLundi;
		private PublicStructures.T_DBRefTypeB pMardi;
		private PublicStructures.T_DBRefTypeB pMercredi;
		private PublicStructures.T_DBRefTypeB pJeudi;
		private PublicStructures.T_DBRefTypeB pVendredi;
		private PublicStructures.T_DBRefTypeB pSamedi;
		private PublicStructures.T_DBRefTypeB pDimanche;
		private PublicStructures.T_DBRefTypeB pFerie;
		private PublicStructures.T_DBRefTypeI pOrder;
		
		
		
		
		public int Version
		{
			get
			{
				return pVersion.theValue;
			}
			set
			{
				pVersion.theValue = value;
			}
		}
		public int ShiftStart
		{
			get
			{
				return pShiftStart.theValue;
			}
			set
			{
				pShiftStart.theValue = value;
			}
		}
		public int ShiftStop
		{
			get
			{
				return pShiftStop.theValue;
			}
			set
			{
				pShiftStop.theValue = value;
			}
		}
		public int ShiftType
		{
			get
			{
				return pShiftType.theValue;
			}
			set
			{
				pShiftType.theValue = value;
			}
		}
		public bool Active
		{
			get
			{
				return pActive.theValue;
			}
			set
			{
				pActive.theValue = value;
			}
		}
		public bool Compilation
		{
			get
			{
				return pCompilation.theValue;
			}
			set
			{
				pCompilation.theValue = value;
			}
		}
		public string Description
		{
			get
			{
				return pDescription.theValue;
			}
			set
			{
				pDescription.theValue = value;
			}
		}
		public bool Lundi
		{
			get
			{
				return pLundi.theValue;
			}
			set
			{
				pLundi.theValue = value;
			}
		}
		public bool Mardi
		{
			get
			{
				return pMardi.theValue;
			}
			set
			{
				pMardi.theValue = value;
			}
		}
		public bool Mercredi
		{
			get
			{
				return pMercredi.theValue;
			}
			set
			{
				pMercredi.theValue = value;
			}
		}
		public bool Jeudi
		{
			get
			{
				return pJeudi.theValue;
			}
			set
			{
				pJeudi.theValue = value;
			}
		}
		public bool Vendredi
		{
			get
			{
				return pVendredi.theValue;
			}
			set
			{
				pVendredi.theValue = value;
			}
		}
		public bool Samedi
		{
			get
			{
				return pSamedi.theValue;
			}
			set
			{
				pSamedi.theValue = value;
			}
		}
		public bool Dimanche
		{
			get
			{
				return pDimanche.theValue;
			}
			set
			{
				pDimanche.theValue = value;
			}
		}
		public bool Ferie
		{
			get
			{
				return pFerie.theValue;
			}
			set
			{
				pFerie.theValue = value;
			}
		}
		public int Order
		{
			get
			{
				return pOrder.theValue;
			}
			set
			{
				pOrder.theValue = value;
			}
		}
		
		
		public SShiftType()
		{
			pShiftStart.theSQLName = PublicConstants.SQLShiftStart;
			pShiftStop.theSQLName = PublicConstants.SQLShiftStop;
			pShiftType.theSQLName = PublicConstants.SQLShiftType;
			pActive.theSQLName = PublicConstants.SQLActive;
			pDescription.theSQLName = PublicConstants.SQLDescription;
			pVersion.theSQLName = PublicConstants.SQLVersion;
			pLundi.theSQLName = PublicConstants.SQLLundi;
			pMardi.theSQLName = PublicConstants.SQLMardi;
			pMercredi.theSQLName = PublicConstants.SQLMercredi;
			pJeudi.theSQLName = PublicConstants.SQLJeudi;
			pVendredi.theSQLName = PublicConstants.SQLVendredi;
			pSamedi.theSQLName = PublicConstants.SQLSamedi;
			pDimanche.theSQLName = PublicConstants.SQLDimanche;
			pFerie.theSQLName = PublicConstants.SQLFerie;
			pCompilation.theSQLName = PublicConstants.SQLCompilation;
			pOrder.theSQLName = PublicConstants.SQLOrder;
		}
		public static Collection loadShiftTypesFromDBPerMonth(int aMonth, int aYear)
		{
			SQLStrBuilder theBuiltSql = new SQLStrBuilder();
			ADODB.Recordset theRS = new ADODB.Recordset();
			DBAC theDBAC = new DBAC();
			SShiftType aShifttype = default(SShiftType);
			Collection theShiftTypeCollection = default(Collection);
			theShiftTypeCollection = new Collection();
			int theVersion = default(int);
			theVersion = ((aYear - 2000) * 100) + aMonth;
			
			//check if a version exists for the month
			
			theBuiltSql.SQL_Select("*");
			theBuiltSql.SQL_From(PublicConstants.TABLE_shiftType);
			theBuiltSql.SQL_Where(PublicConstants.SQLVersion, "=", theVersion);
			theBuiltSql.SQL_Order_By(PublicConstants.SQLOrder);
			theDBAC.COpenDB(theBuiltSql.SQLStringSelect, theRS);
			
			if (theRS.RecordCount > 0) //if a version exists load it
			{
				theRS.MoveFirst();
				for (int x = 1; x <= theRS.RecordCount; x++)
				{
					aShifttype = new SShiftType();
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLShiftStart].Value))
					{
						aShifttype.ShiftStart = System.Convert.ToInt32(theRS.Fields[PublicConstants.SQLShiftStart].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLShiftStop].Value))
					{
						aShifttype.ShiftStop = System.Convert.ToInt32(theRS.Fields[PublicConstants.SQLShiftStop].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLShiftType].Value))
					{
						aShifttype.ShiftType = System.Convert.ToInt32(theRS.Fields[PublicConstants.SQLShiftType].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLActive].Value))
					{
						aShifttype.Active = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLActive].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLVersion].Value))
					{
						aShifttype.Version = System.Convert.ToInt32(theRS.Fields[PublicConstants.SQLVersion].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLDescription].Value))
					{
						aShifttype.Description = (theRS.Fields[PublicConstants.SQLDescription].Value).ToString();
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLLundi].Value))
					{
						aShifttype.Lundi = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLLundi].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLMardi].Value))
					{
						aShifttype.Mardi = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLMardi].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLMercredi].Value))
					{
						aShifttype.Mercredi = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLMercredi].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLJeudi].Value))
					{
						aShifttype.Jeudi = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLJeudi].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLVendredi].Value))
					{
						aShifttype.Vendredi = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLVendredi].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLSamedi].Value))
					{
						aShifttype.Samedi = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLSamedi].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLDimanche].Value))
					{
						aShifttype.Dimanche = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLDimanche].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLFerie].Value))
					{
						aShifttype.Ferie = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLFerie].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLCompilation].Value))
					{
						aShifttype.Compilation = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLCompilation].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLOrder].Value))
					{
						aShifttype.Order = System.Convert.ToInt32(theRS.Fields[PublicConstants.SQLOrder].Value);
					}
					
					
					theShiftTypeCollection.Add(aShifttype, null, null, null);
					theRS.MoveNext();
				}
			}
			else //if no version exists, load the template version (0)
			{
				theBuiltSql.SQLClear();
				theBuiltSql.SQL_Select("*");
				theBuiltSql.SQL_From(PublicConstants.TABLE_shiftType);
				theBuiltSql.SQL_Where(PublicConstants.SQLVersion, "=", 0);
				theBuiltSql.SQL_Order_By(PublicConstants.SQLOrder);
				theDBAC.COpenDB(theBuiltSql.SQLStringSelect, theRS);
				
				if (theRS.RecordCount > 0) //if at least one template shifttype exists load it as a collection
				{
					
					theRS.MoveFirst();
					for (int x = 1; x <= theRS.RecordCount; x++)
					{
						aShifttype = new SShiftType();
						if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLShiftStart].Value))
						{
							aShifttype.ShiftStart = System.Convert.ToInt32(theRS.Fields[PublicConstants.SQLShiftStart].Value);
						}
						if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLShiftStop].Value))
						{
							aShifttype.ShiftStop = System.Convert.ToInt32(theRS.Fields[PublicConstants.SQLShiftStop].Value);
						}
						if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLShiftType].Value))
						{
							aShifttype.ShiftType = System.Convert.ToInt32(theRS.Fields[PublicConstants.SQLShiftType].Value);
						}
						if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLActive].Value))
						{
							aShifttype.Active = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLActive].Value);
						}
						aShifttype.Version = theVersion; //change version to YYYYMM integer
						if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLDescription].Value))
						{
							aShifttype.Description = (theRS.Fields[PublicConstants.SQLDescription].Value).ToString();
						}
						if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLLundi].Value))
						{
							aShifttype.Lundi = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLLundi].Value);
						}
						if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLMardi].Value))
						{
							aShifttype.Mardi = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLMardi].Value);
						}
						if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLMercredi].Value))
						{
							aShifttype.Mercredi = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLMercredi].Value);
						}
						if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLJeudi].Value))
						{
							aShifttype.Jeudi = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLJeudi].Value);
						}
						if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLVendredi].Value))
						{
							aShifttype.Vendredi = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLVendredi].Value);
						}
						if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLSamedi].Value))
						{
							aShifttype.Samedi = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLSamedi].Value);
						}
						if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLDimanche].Value))
						{
							aShifttype.Dimanche = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLDimanche].Value);
						}
						if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLFerie].Value))
						{
							aShifttype.Ferie = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLFerie].Value);
						}
						if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLCompilation].Value))
						{
							aShifttype.Compilation = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLCompilation].Value);
						}
						if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLOrder].Value))
						{
							aShifttype.Order = System.Convert.ToInt32(theRS.Fields[PublicConstants.SQLOrder].Value);
						}
						aShifttype.Save(); //save the shifttype version to DB
						theShiftTypeCollection.Add(aShifttype, null, null, null);
						theRS.MoveNext();
					}
					
				}
			}
			return theShiftTypeCollection;
		}
		public static Collection loadTemplateShiftTypesFromDB()
		{
			SQLStrBuilder theBuiltSql = new SQLStrBuilder();
			ADODB.Recordset theRS = new ADODB.Recordset();
			DBAC theDBAC = new DBAC();
			SShiftType aShifttype = default(SShiftType);
			Collection theShiftTypeCollection = default(Collection);
			theShiftTypeCollection = new Collection();
			theBuiltSql.SQL_Select("*");
			theBuiltSql.SQL_From(PublicConstants.TABLE_shiftType);
			theBuiltSql.SQL_Where(PublicConstants.SQLVersion, "=", 0);
			theBuiltSql.SQL_Order_By(PublicConstants.SQLOrder);
			theDBAC.COpenDB(theBuiltSql.SQLStringSelect, theRS);
			
			if (theRS.RecordCount > 0)
			{
				
				theRS.MoveFirst();
				for (int x = 1; x <= theRS.RecordCount; x++)
				{
					aShifttype = new SShiftType();
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLShiftStart].Value))
					{
						aShifttype.ShiftStart = System.Convert.ToInt32(theRS.Fields[PublicConstants.SQLShiftStart].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLShiftStop].Value))
					{
						aShifttype.ShiftStop = System.Convert.ToInt32(theRS.Fields[PublicConstants.SQLShiftStop].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLShiftType].Value))
					{
						aShifttype.ShiftType = System.Convert.ToInt32(theRS.Fields[PublicConstants.SQLShiftType].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLActive].Value))
					{
						aShifttype.Active = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLActive].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLDescription].Value))
					{
						aShifttype.Description = (theRS.Fields[PublicConstants.SQLDescription].Value).ToString();
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLLundi].Value))
					{
						aShifttype.Lundi = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLLundi].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLMardi].Value))
					{
						aShifttype.Mardi = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLMardi].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLMercredi].Value))
					{
						aShifttype.Mercredi = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLMercredi].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLJeudi].Value))
					{
						aShifttype.Jeudi = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLJeudi].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLVendredi].Value))
					{
						aShifttype.Vendredi = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLVendredi].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLSamedi].Value))
					{
						aShifttype.Samedi = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLSamedi].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLDimanche].Value))
					{
						aShifttype.Dimanche = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLDimanche].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLFerie].Value))
					{
						aShifttype.Ferie = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLFerie].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLCompilation].Value))
					{
						aShifttype.Compilation = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLCompilation].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLOrder].Value))
					{
						aShifttype.Order = System.Convert.ToInt32(theRS.Fields[PublicConstants.SQLOrder].Value);
					}
					
					theShiftTypeCollection.Add(aShifttype, null, null, null);
					theRS.MoveNext();
				}
			}
			return theShiftTypeCollection;
		}
		public static int ActiveShiftTypesCountPerMonth(int aMonth, int aYear)
		{
			SQLStrBuilder theBuiltSql = new SQLStrBuilder();
			ADODB.Recordset theRS = new ADODB.Recordset();
			DBAC theDBAC = new DBAC();
			int theVersion = default(int);
			theVersion = ((aYear - 2000) * 100) + aMonth;
			
			//check if a version exists for the month
			
			theBuiltSql.SQL_Select("*");
			theBuiltSql.SQL_From(PublicConstants.TABLE_shiftType);
			theBuiltSql.SQL_Where(PublicConstants.SQLVersion, "=", theVersion);
			theBuiltSql.SQL_Where(PublicConstants.SQLActive, "=", true);
			theDBAC.COpenDB(theBuiltSql.SQLStringSelect, theRS);
			
			return theRS.RecordCount;
		}
		public void Copy(SShiftType TheInstanceToBeCopied)
		{
			
			SShiftType with_1 = TheInstanceToBeCopied;
			
			//Me.pCollection = .ShiftCollection
			this.ShiftStart = with_1.ShiftStart;
			this.ShiftStop = with_1.ShiftStop;
			this.ShiftType = with_1.ShiftType;
			this.Version = with_1.Version;
			this.Active = with_1.Active;
			this.Description = with_1.Description;
			this.Lundi = with_1.Lundi;
			this.Mardi = with_1.Mardi;
			this.Mercredi = with_1.Mercredi;
			this.Jeudi = with_1.Jeudi;
			this.Vendredi = with_1.Vendredi;
			this.Samedi = with_1.Samedi;
			this.Dimanche = with_1.Dimanche;
			this.Ferie = with_1.Ferie;
			this.Compilation = with_1.Compilation;
			this.Order = with_1.Order;
			
		}
		public void Save()
		{
			SQLStrBuilder theBuiltSql = new SQLStrBuilder();
			ADODB.Recordset theRS = new ADODB.Recordset();
			DBAC theDBAC = new DBAC();
			
			theBuiltSql.SQL_Select("*");
			theBuiltSql.SQL_From(PublicConstants.TABLE_shiftType);
			theBuiltSql.SQL_Where(pShiftType.theSQLName, "=", this.ShiftType);
			theBuiltSql.SQL_Where(pVersion.theSQLName, "=", this.Version);
			theDBAC.COpenDB(theBuiltSql.SQLStringSelect, theRS);
			
			int theCount = theRS.RecordCount;
			
			switch (theCount)
			{
				case 0: //if not create a new entry
					theBuiltSql.SQLClear();
					theBuiltSql.SQL_Insert(PublicConstants.TABLE_shiftType);
					theBuiltSql.SQL_Values(pShiftStart.theSQLName, ShiftStart);
					theBuiltSql.SQL_Values(pShiftStop.theSQLName, ShiftStop);
					theBuiltSql.SQL_Values(pVersion.theSQLName, Version);
					theBuiltSql.SQL_Values(pShiftType.theSQLName, ShiftType);
					theBuiltSql.SQL_Values(pActive.theSQLName, Active);
					theBuiltSql.SQL_Values(pDescription.theSQLName, Description);
					theBuiltSql.SQL_Values(pLundi.theSQLName, Lundi);
					theBuiltSql.SQL_Values(pMardi.theSQLName, Mardi);
					theBuiltSql.SQL_Values(pMercredi.theSQLName, Mercredi);
					theBuiltSql.SQL_Values(pJeudi.theSQLName, Jeudi);
					theBuiltSql.SQL_Values(pVendredi.theSQLName, Vendredi);
					theBuiltSql.SQL_Values(pSamedi.theSQLName, Samedi);
					theBuiltSql.SQL_Values(pDimanche.theSQLName, Dimanche);
					theBuiltSql.SQL_Values(pFerie.theSQLName, Ferie);
					theBuiltSql.SQL_Values(pCompilation.theSQLName, Compilation);
					theBuiltSql.SQL_Values(pOrder.theSQLName, Order);
					
					int numaffected = default(int);
					theDBAC.CExecuteDB(theBuiltSql.SQLStringInsert, numaffected);
					break;
				default:
					Debug.WriteLine("there is already an existing instance with this version number ... this is bad");
					break;
			}
		}
		public void Update()
		{
			SQLStrBuilder theBuiltSql = new SQLStrBuilder();
			ADODB.Recordset theRS = new ADODB.Recordset();
			DBAC theDBAC = new DBAC();
			
			theBuiltSql.SQL_Select("*");
			theBuiltSql.SQL_From(PublicConstants.TABLE_shiftType);
			theBuiltSql.SQL_Where(pShiftType.theSQLName, "=", this.ShiftType);
			theBuiltSql.SQL_Where(pVersion.theSQLName, "=", this.Version);
			theDBAC.COpenDB(theBuiltSql.SQLStringSelect, theRS);
			
			int theCount = theRS.RecordCount;
			
			switch (theCount)
			{
				case 0:
					Debug.WriteLine("there is nothing to update ... this is bad");
					break;
					
				case 1: //if yes update it with the new value
					theRS.Fields[pShiftStart.theSQLName].Value = ShiftStart;
					theRS.Fields[pShiftStop.theSQLName].Value = ShiftStop;
					theRS.Fields[pVersion.theSQLName].Value = Version;
					theRS.Fields[pActive.theSQLName].Value = Active;
					theRS.Fields[pShiftType.theSQLName].Value = ShiftType;
					theRS.Fields[pDescription.theSQLName].Value = Description;
					theRS.Fields[pLundi.theSQLName].Value = Lundi;
					theRS.Fields[pMardi.theSQLName].Value = Mardi;
					theRS.Fields[pMercredi.theSQLName].Value = Mercredi;
					theRS.Fields[pJeudi.theSQLName].Value = Jeudi;
					theRS.Fields[pVendredi.theSQLName].Value = Vendredi;
					theRS.Fields[pSamedi.theSQLName].Value = Samedi;
					theRS.Fields[pDimanche.theSQLName].Value = Dimanche;
					theRS.Fields[pFerie.theSQLName].Value = Ferie;
					theRS.Fields[pCompilation.theSQLName].Value = Compilation;
					theRS.Fields[pOrder.theSQLName].Value = Order;
					theRS.ActiveConnection = theDBAC.aConnection;
					theRS.UpdateBatch((ADODB.AffectEnum) 3);
					theRS.Close();
					break;
				default:
					Debug.WriteLine("there is more than one copy of this entry ... this is bad");
					break;
			}
		}
		
	}
	
	public class SDoc
	{
		private PublicStructures.T_DBRefTypeS pFirstName;
		private PublicStructures.T_DBRefTypeS pLastName;
		private PublicStructures.T_DBRefTypeS pInitials;
		private PublicStructures.T_DBRefTypeB pActive;
		private PublicStructures.T_DBRefTypeI pVersion;
		private PublicStructures.T_DBRefTypeI pShift1;
		private PublicStructures.T_DBRefTypeI pShift2;
		private PublicStructures.T_DBRefTypeI pShift3;
		private PublicStructures.T_DBRefTypeI pShift4;
		private PublicStructures.T_DBRefTypeI pShift5;
		private PublicStructures.T_DBRefTypeB pUrgenceTog;
		private PublicStructures.T_DBRefTypeB pHospitTog;
		private PublicStructures.T_DBRefTypeB pSoinsTog;
		private PublicStructures.T_DBRefTypeB pNuitsTog;
		private int pYear;
		private int pMonth;
		
		public string FirstName
		{
			get
			{
				return pFirstName.theValue;
			}
			set
			{
				pFirstName.theValue = value;
			}
		}
		public string LastName
		{
			get
			{
				return pLastName.theValue;
			}
			set
			{
				pLastName.theValue = value;
			}
		}
		public string Initials
		{
			get
			{
				return pInitials.theValue;
			}
			set
			{
				pInitials.theValue = value;
			}
		}
		public bool Active
		{
			get
			{
				return pActive.theValue;
			}
			set
			{
				pActive.theValue = value;
			}
		}
		public int Version
		{
			get
			{
				return pVersion.theValue;
			}
			set
			{
				pVersion.theValue = value;
			}
		}
		public int Shift1
		{
			get
			{
				return pShift1.theValue;
			}
			set
			{
				pShift1.theValue = value;
			}
		}
		public int Shift2
		{
			get
			{
				return pShift2.theValue;
			}
			set
			{
				pShift2.theValue = value;
			}
		}
		public int Shift3
		{
			get
			{
				return pShift3.theValue;
			}
			set
			{
				pShift3.theValue = value;
			}
		}
		public int Shift4
		{
			get
			{
				return pShift4.theValue;
			}
			set
			{
				pShift4.theValue = value;
			}
		}
		public int Shift5
		{
			get
			{
				return pShift5.theValue;
			}
			set
			{
				pShift5.theValue = value;
			}
		}
		public bool UrgenceTog
		{
			get
			{
				return pUrgenceTog.theValue;
			}
			set
			{
				pUrgenceTog.theValue = value;
			}
		}
		public bool HospitTog
		{
			get
			{
				return pHospitTog.theValue;
			}
			set
			{
				pHospitTog.theValue = value;
			}
		}
		public bool SoinsTog
		{
			get
			{
				return pSoinsTog.theValue;
			}
			set
			{
				pSoinsTog.theValue = value;
			}
		}
		public bool NuitsTog
		{
			get
			{
				return pNuitsTog.theValue;
			}
			set
			{
				pNuitsTog.theValue = value;
			}
		}
		public string FistAndLastName
		{
			get
			{
				return FirstName + " " + LastName;
			}
		}
		public SDoc()
		{
			pFirstName.theSQLName = PublicConstants.SQLFirstName;
			pLastName.theSQLName = PublicConstants.SQLLastName;
			pInitials.theSQLName = PublicConstants.SQLInitials;
			pActive.theSQLName = PublicConstants.SQLActive;
			pVersion.theSQLName = PublicConstants.SQLVersion;
			pShift1.theSQLName = PublicConstants.SQLShift1;
			pShift2.theSQLName = PublicConstants.SQLShift2;
			pShift3.theSQLName = PublicConstants.SQLShift3;
			pShift4.theSQLName = PublicConstants.SQLShift4;
			pShift5.theSQLName = PublicConstants.SQLShift5;
			pUrgenceTog.theSQLName = PublicConstants.SQLUrgenceTog;
			pHospitTog.theSQLName = PublicConstants.SQLHospitTog;
			pSoinsTog.theSQLName = PublicConstants.SQLSoinsTog;
			pNuitsTog.theSQLName = PublicConstants.SQLNuitsTog;
			
			FirstName = "FirstName";
			LastName = "LastName";
			Initials = "Initialles";
			Active = true;
			Version = 1;
			Shift1 = 0;
			Shift2 = 0;
			Shift3 = 0;
			Shift4 = 0;
			Shift5 = 0;
			UrgenceTog = true;
			HospitTog = true;
			SoinsTog = true;
			NuitsTog = true;
		}
		public void Delete()
		{
			SQLStrBuilder theBuiltSql = new SQLStrBuilder();
			ADODB.Recordset theRS = new ADODB.Recordset();
			DBAC theDBAC = new DBAC();
			int numaffected = default(int);
			theBuiltSql.SQL_From(PublicConstants.TABLE_Doc);
			theBuiltSql.SQL_Where(pFirstName.theSQLName, "=", FirstName);
			theBuiltSql.SQL_Where(pLastName.theSQLName, "=", LastName);
			theBuiltSql.SQL_Where(pInitials.theSQLName, "=", Initials);
			theBuiltSql.SQL_Where(pActive.theSQLName, "=", Active);
			theBuiltSql.SQL_Where(pVersion.theSQLName, "=", Version);
			theBuiltSql.SQL_Where(pShift1.theSQLName, "=", Shift1);
			theBuiltSql.SQL_Where(pShift2.theSQLName, "=", Shift2);
			theBuiltSql.SQL_Where(pShift3.theSQLName, "=", Shift3);
			theBuiltSql.SQL_Where(pShift4.theSQLName, "=", Shift4);
			theBuiltSql.SQL_Where(pShift5.theSQLName, "=", Shift5);
			theBuiltSql.SQL_Where(pUrgenceTog.theSQLName, "=", UrgenceTog);
			theBuiltSql.SQL_Where(pHospitTog.theSQLName, "=", HospitTog);
			theBuiltSql.SQL_Where(pSoinsTog.theSQLName, "=", SoinsTog);
			theBuiltSql.SQL_Where(pNuitsTog.theSQLName, "=", NuitsTog);
			theDBAC.CExecuteDB(theBuiltSql.SQLStringDelete, numaffected);
			if (numaffected != 1)
			{
				Debug.WriteLine("there is more than one copy of this entry ... this is bad");
			}
		}
		public SDoc(int aYear, int aMonth)
		{
			pFirstName.theSQLName = PublicConstants.SQLFirstName;
			pLastName.theSQLName = PublicConstants.SQLLastName;
			pInitials.theSQLName = PublicConstants.SQLInitials;
			pActive.theSQLName = PublicConstants.SQLActive;
			pVersion.theSQLName = PublicConstants.SQLVersion;
			pShift1.theSQLName = PublicConstants.SQLShift1;
			pShift2.theSQLName = PublicConstants.SQLShift2;
			pShift3.theSQLName = PublicConstants.SQLShift3;
			pShift4.theSQLName = PublicConstants.SQLShift4;
			pShift5.theSQLName = PublicConstants.SQLShift5;
			pUrgenceTog.theSQLName = PublicConstants.SQLUrgenceTog;
			pHospitTog.theSQLName = PublicConstants.SQLHospitTog;
			pSoinsTog.theSQLName = PublicConstants.SQLSoinsTog;
			pNuitsTog.theSQLName = PublicConstants.SQLNuitsTog;
			
			FirstName = "FirstName";
			LastName = "LastName";
			Initials = "Initialles";
			Active = true;
			Version = 1;
			Shift1 = 0;
			Shift2 = 0;
			Shift3 = 0;
			Shift4 = 0;
			Shift5 = 0;
			UrgenceTog = true;
			HospitTog = true;
			SoinsTog = true;
			NuitsTog = true;
			
			
			//pYear = aYear
			//pMonth = aMonth
			
			//If pDocList Is Nothing Then
			//    pDocList = New Collection
			//    LoadAllDocs(aYear, aMonth)
			//End If
		}
		public SDoc(string aFirstName, string aLastName, string aInitials, bool aActive, int aVersion, int aShift1, int aShift2, int aShift3, int aShift4, int aShift5, bool aUrgenceTog, bool aHospitTog, bool aSoinsTog, bool aNuitsTog)
		{
			
			pFirstName.theSQLName = PublicConstants.SQLFirstName;
			pLastName.theSQLName = PublicConstants.SQLLastName;
			pInitials.theSQLName = PublicConstants.SQLInitials;
			pActive.theSQLName = PublicConstants.SQLActive;
			pVersion.theSQLName = PublicConstants.SQLVersion;
			pShift1.theSQLName = PublicConstants.SQLShift1;
			pShift2.theSQLName = PublicConstants.SQLShift2;
			pShift3.theSQLName = PublicConstants.SQLShift3;
			pShift4.theSQLName = PublicConstants.SQLShift4;
			pShift5.theSQLName = PublicConstants.SQLShift5;
			pUrgenceTog.theSQLName = PublicConstants.SQLUrgenceTog;
			pHospitTog.theSQLName = PublicConstants.SQLHospitTog;
			pSoinsTog.theSQLName = PublicConstants.SQLSoinsTog;
			pNuitsTog.theSQLName = PublicConstants.SQLNuitsTog;
			
			FirstName = aFirstName;
			LastName = aLastName;
			Initials = aInitials;
			Active = aActive;
			Version = aVersion;
			Shift1 = aShift1;
			Shift2 = aShift2;
			Shift3 = aShift3;
			Shift4 = aShift4;
			Shift5 = aShift5;
			
			UrgenceTog = aUrgenceTog;
			HospitTog = aHospitTog;
			SoinsTog = aSoinsTog;
			NuitsTog = aNuitsTog;
			
		}
		public void save()
		{
			SQLStrBuilder theBuiltSql = new SQLStrBuilder();
			ADODB.Recordset theRS = new ADODB.Recordset();
			DBAC theDBAC = new DBAC();
			
			theBuiltSql.SQL_Select("*");
			theBuiltSql.SQL_From(PublicConstants.TABLE_Doc);
			theBuiltSql.SQL_Where(pInitials.theSQLName, "=", Initials);
			theBuiltSql.SQL_Where(pVersion.theSQLName, "=", Version);
			theDBAC.COpenDB(theBuiltSql.SQLStringSelect, theRS);
			
			int theCount = theRS.RecordCount;
			
			switch (theCount)
			{
				case 0: //if not create a new entry
					theBuiltSql.SQLClear();
					theBuiltSql.SQL_Insert(PublicConstants.TABLE_Doc);
					theBuiltSql.SQL_Values(pFirstName.theSQLName, FirstName);
					theBuiltSql.SQL_Values(pLastName.theSQLName, LastName);
					theBuiltSql.SQL_Values(pInitials.theSQLName, Initials);
					theBuiltSql.SQL_Values(pActive.theSQLName, Active);
					theBuiltSql.SQL_Values(pVersion.theSQLName, Version);
					theBuiltSql.SQL_Values(pShift1.theSQLName, Shift1);
					theBuiltSql.SQL_Values(pShift2.theSQLName, Shift2);
					theBuiltSql.SQL_Values(pShift3.theSQLName, Shift3);
					theBuiltSql.SQL_Values(pShift4.theSQLName, Shift4);
					theBuiltSql.SQL_Values(pShift5.theSQLName, Shift5);
					theBuiltSql.SQL_Values(pUrgenceTog.theSQLName, UrgenceTog);
					theBuiltSql.SQL_Values(pHospitTog.theSQLName, HospitTog);
					theBuiltSql.SQL_Values(pSoinsTog.theSQLName, SoinsTog);
					theBuiltSql.SQL_Values(pNuitsTog.theSQLName, NuitsTog);
					
					int numaffected = default(int);
					theDBAC.CExecuteDB(theBuiltSql.SQLStringInsert, numaffected);
					//Debug.WriteLine(.SQLStringInsert)
					//Debug.WriteLine("Number of databaseentries" + numaffected.ToString())
					break;
					
				case 1: //if yes update it with the new value
					theRS.Fields[pFirstName.theSQLName].Value = FirstName;
					theRS.Fields[pLastName.theSQLName].Value = LastName;
					theRS.Fields[pInitials.theSQLName].Value = Initials;
					theRS.Fields[pActive.theSQLName].Value = Active;
					theRS.Fields[pVersion.theSQLName].Value = Version;
					theRS.Fields[pShift1.theSQLName].Value = Shift1;
					theRS.Fields[pShift2.theSQLName].Value = Shift2;
					theRS.Fields[pShift3.theSQLName].Value = Shift3;
					theRS.Fields[pShift4.theSQLName].Value = Shift4;
					theRS.Fields[pShift5.theSQLName].Value = Shift5;
					theRS.Fields[pUrgenceTog.theSQLName].Value = UrgenceTog;
					theRS.Fields[pHospitTog.theSQLName].Value = HospitTog;
					theRS.Fields[pSoinsTog.theSQLName].Value = SoinsTog;
					theRS.Fields[pNuitsTog.theSQLName].Value = NuitsTog;
					
					theRS.ActiveConnection = theDBAC.aConnection;
					theRS.UpdateBatch((ADODB.AffectEnum) 3);
					theRS.Close();
					break;
				default:
					Debug.WriteLine("there is more than one copy of this entry ... this is bad");
					break;
			}
		}
		public static Collection LoadAllDocsPerMonth(int aYear, int aMonth)
		{
			SQLStrBuilder theBuiltSql = new SQLStrBuilder();
			ADODB.Recordset theRS = new ADODB.Recordset();
			DBAC theDBAC = new DBAC();
			DateTime theCurrentMonthDate = DateAndTime.DateSerial(aYear, aMonth, 1);
			Collection aCollection = default(Collection);
			aCollection = new Collection();
			int theVersion = default(int);
			theVersion = ((aYear - 2000) * 100) + aMonth;
			
			//check if a version exists for the month
			theBuiltSql.SQL_Select("*");
			theBuiltSql.SQL_From(PublicConstants.TABLE_Doc);
			theBuiltSql.SQL_Where(PublicConstants.SQLVersion, "=", theVersion);
			theBuiltSql.SQL_Order_By(PublicConstants.SQLLastName);
			theDBAC.COpenDB(theBuiltSql.SQLStringSelect, theRS);
			
			if (theRS.RecordCount > 0) //if a version exists load it
			{
				theRS.MoveFirst();
				for (int x = 1; x <= theRS.RecordCount; x++)
				{
					SDoc aSDoc = new SDoc();
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLFirstName].Value))
					{
						aSDoc.FirstName = (theRS.Fields[PublicConstants.SQLFirstName].Value).ToString();
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLLastName].Value))
					{
						aSDoc.LastName = (theRS.Fields[PublicConstants.SQLLastName].Value).ToString();
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLInitials].Value))
					{
						aSDoc.Initials = (theRS.Fields[PublicConstants.SQLInitials].Value).ToString();
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLActive].Value))
					{
						aSDoc.Active = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLActive].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLVersion].Value))
					{
						aSDoc.Version = System.Convert.ToInt32(theRS.Fields[PublicConstants.SQLVersion].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLShift1].Value))
					{
						aSDoc.Shift1 = System.Convert.ToInt32(theRS.Fields[PublicConstants.SQLShift1].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLShift2].Value))
					{
						aSDoc.Shift2 = System.Convert.ToInt32(theRS.Fields[PublicConstants.SQLShift2].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLShift3].Value))
					{
						aSDoc.Shift3 = System.Convert.ToInt32(theRS.Fields[PublicConstants.SQLShift3].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLShift4].Value))
					{
						aSDoc.Shift4 = System.Convert.ToInt32(theRS.Fields[PublicConstants.SQLShift4].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLShift5].Value))
					{
						aSDoc.Shift5 = System.Convert.ToInt32(theRS.Fields[PublicConstants.SQLShift5].Value);
					}
					
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLUrgenceTog].Value))
					{
						aSDoc.UrgenceTog = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLUrgenceTog].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLHospitTog].Value))
					{
						aSDoc.HospitTog = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLHospitTog].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLSoinsTog].Value))
					{
						aSDoc.SoinsTog = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLSoinsTog].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLNuitsTog].Value))
					{
						aSDoc.NuitsTog = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLNuitsTog].Value);
					}
					
					aCollection.Add(aSDoc, aSDoc.Initials, null, null);
					theRS.MoveNext();
				}
			}
			else //if no version exists, load the template version (0)
			{
				theBuiltSql.SQLClear();
				theBuiltSql.SQL_Select("*");
				theBuiltSql.SQL_From(PublicConstants.TABLE_Doc);
				theBuiltSql.SQL_Where(PublicConstants.SQLVersion, "=", 0);
				theBuiltSql.SQL_Order_By(PublicConstants.SQLLastName);
				theDBAC.COpenDB(theBuiltSql.SQLStringSelect, theRS);
				
				if (theRS.RecordCount > 0) //if at least one template shifttype exists load it as a collection
				{
					theRS.MoveFirst();
					for (int x = 1; x <= theRS.RecordCount; x++)
					{
						SDoc aSDoc = new SDoc();
						if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLFirstName].Value))
						{
							aSDoc.FirstName = (theRS.Fields[PublicConstants.SQLFirstName].Value).ToString();
						}
						if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLLastName].Value))
						{
							aSDoc.LastName = (theRS.Fields[PublicConstants.SQLLastName].Value).ToString();
						}
						if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLInitials].Value))
						{
							aSDoc.Initials = (theRS.Fields[PublicConstants.SQLInitials].Value).ToString();
						}
						if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLActive].Value))
						{
							aSDoc.Active = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLActive].Value);
						}
						aSDoc.Version = theVersion; //change version to YYYYMM integer
						if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLShift1].Value))
						{
							aSDoc.Shift1 = System.Convert.ToInt32(theRS.Fields[PublicConstants.SQLShift1].Value);
						}
						if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLShift2].Value))
						{
							aSDoc.Shift2 = System.Convert.ToInt32(theRS.Fields[PublicConstants.SQLShift2].Value);
						}
						if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLShift3].Value))
						{
							aSDoc.Shift3 = System.Convert.ToInt32(theRS.Fields[PublicConstants.SQLShift3].Value);
						}
						if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLShift4].Value))
						{
							aSDoc.Shift4 = System.Convert.ToInt32(theRS.Fields[PublicConstants.SQLShift4].Value);
						}
						if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLShift5].Value))
						{
							aSDoc.Shift5 = System.Convert.ToInt32(theRS.Fields[PublicConstants.SQLShift5].Value);
						}
						if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLUrgenceTog].Value))
						{
							aSDoc.UrgenceTog = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLUrgenceTog].Value);
						}
						if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLHospitTog].Value))
						{
							aSDoc.HospitTog = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLHospitTog].Value);
						}
						if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLSoinsTog].Value))
						{
							aSDoc.SoinsTog = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLSoinsTog].Value);
						}
						if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLNuitsTog].Value))
						{
							aSDoc.NuitsTog = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLNuitsTog].Value);
						}
						aSDoc.save();
						aCollection.Add(aSDoc, aSDoc.Initials, null, null);
						theRS.MoveNext();
					}
				}
			}
			
			return aCollection;
		}
		public static Collection LoadTempateDocsFromDB()
		{
			SQLStrBuilder theBuiltSql = new SQLStrBuilder();
			ADODB.Recordset theRS = new ADODB.Recordset();
			DBAC theDBAC = new DBAC();
			Collection aCollection = default(Collection);
			aCollection = new Collection();
			
			//check if a version exists for the month
			theBuiltSql.SQL_Select("*");
			theBuiltSql.SQL_From(PublicConstants.TABLE_Doc);
			theBuiltSql.SQL_Where(PublicConstants.SQLVersion, "=", 0);
			theBuiltSql.SQL_Order_By(PublicConstants.SQLLastName);
			theDBAC.COpenDB(theBuiltSql.SQLStringSelect, theRS);
			
			if (theRS.RecordCount > 0) //if a version exists load it
			{
				theRS.MoveFirst();
				for (int x = 1; x <= theRS.RecordCount; x++)
				{
					SDoc aSDoc = new SDoc();
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLFirstName].Value))
					{
						aSDoc.FirstName = (theRS.Fields[PublicConstants.SQLFirstName].Value).ToString();
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLLastName].Value))
					{
						aSDoc.LastName = (theRS.Fields[PublicConstants.SQLLastName].Value).ToString();
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLInitials].Value))
					{
						aSDoc.Initials = (theRS.Fields[PublicConstants.SQLInitials].Value).ToString();
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLActive].Value))
					{
						aSDoc.Active = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLActive].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLVersion].Value))
					{
						aSDoc.Version = System.Convert.ToInt32(theRS.Fields[PublicConstants.SQLVersion].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLShift1].Value))
					{
						aSDoc.Shift1 = System.Convert.ToInt32(theRS.Fields[PublicConstants.SQLShift1].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLShift2].Value))
					{
						aSDoc.Shift2 = System.Convert.ToInt32(theRS.Fields[PublicConstants.SQLShift2].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLShift3].Value))
					{
						aSDoc.Shift3 = System.Convert.ToInt32(theRS.Fields[PublicConstants.SQLShift3].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLShift4].Value))
					{
						aSDoc.Shift4 = System.Convert.ToInt32(theRS.Fields[PublicConstants.SQLShift4].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLShift5].Value))
					{
						aSDoc.Shift5 = System.Convert.ToInt32(theRS.Fields[PublicConstants.SQLShift5].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLUrgenceTog].Value))
					{
						aSDoc.UrgenceTog = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLUrgenceTog].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLHospitTog].Value))
					{
						aSDoc.HospitTog = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLHospitTog].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLSoinsTog].Value))
					{
						aSDoc.SoinsTog = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLSoinsTog].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[PublicConstants.SQLNuitsTog].Value))
					{
						aSDoc.NuitsTog = System.Convert.ToBoolean(theRS.Fields[PublicConstants.SQLNuitsTog].Value);
					}
					
					aCollection.Add(aSDoc, aSDoc.Initials, null, null);
					theRS.MoveNext();
				}
			}
			return aCollection;
		}
		
	}
	
	public class SDocStats
	{
		private string pInitials;
		private int pShift1;
		private int pShift2;
		private int pShift3;
		private int pShift4;
		private int pShift5;
		private int pShift1E;
		private int pShift2E;
		private int pShift3E;
		private int pShift4E;
		private int pShift5E;
		
		public string Initials
		{
			get
			{
				return pInitials;
			}
			set
			{
				pInitials = value;
			}
		}
		public int shift1
		{
			get
			{
				return pShift1;
			}
			set
			{
				pShift1 = value;
			}
		}
		public int shift2
		{
			get
			{
				return pShift2;
			}
			set
			{
				pShift2 = value;
			}
		}
		public int shift3
		{
			get
			{
				return pShift3;
			}
			set
			{
				pShift3 = value;
			}
		}
		public int shift4
		{
			get
			{
				return pShift4;
			}
			set
			{
				pShift4 = value;
			}
		}
		public int shift5
		{
			get
			{
				return pShift5;
			}
			set
			{
				pShift5 = value;
			}
		}
		public int shift1E
		{
			get
			{
				return pShift1E;
			}
			set
			{
				pShift1E = value;
			}
		}
		public int shift2E
		{
			get
			{
				return pShift2E;
			}
			set
			{
				pShift2E = value;
			}
		}
		public int shift3E
		{
			get
			{
				return pShift3E;
			}
			set
			{
				pShift3E = value;
			}
		}
		public int shift4E
		{
			get
			{
				return pShift4E;
			}
			set
			{
				pShift4E = value;
			}
		}
		public int shift5E
		{
			get
			{
				return pShift5E;
			}
			set
			{
				pShift5E = value;
			}
		}
		
		public SDocStats(string aInitials, int aShift1, int aShift2, int aShift3, int aShift4, int aShift5)
		{
			Initials = aInitials;
			shift1 = aShift1;
			shift2 = aShift2;
			shift3 = aShift3;
			shift4 = aShift4;
			shift5 = aShift5;
			shift1E = aShift1;
			shift2E = aShift2;
			shift3E = aShift3;
			shift4E = aShift4;
			shift5E = aShift5;
			
		}
		
	}
	
	public class SDocAvailable
	{
		private PublicStructures.T_DBRefTypeS pDocInitial;
		private PublicEnums.Availability pAvailability;
		private PublicStructures.T_DBRefTypeD pDate;
		private PublicStructures.T_DBRefTypeI pShiftType;
		
		public string DocInitial
		{
			get
			{
				return pDocInitial.theValue;
			}
			set
			{
				pDocInitial.theValue = value;
			}
		}
		public PublicEnums.Availability Availability
		{
			get
			{
				return pAvailability;
			}
			set
			{
				if (value == PublicEnums.Availability.Assigne)
				{
					UpdateScheduleDataTable(value);
				}
				else if (pAvailability == PublicEnums.Availability.Assigne & value == PublicEnums.Availability.Dispo)
				{
					DeleteScheduleDataEntry();
				}
				pAvailability = value;
			}
		}
		public PublicEnums.Availability SetAvailabilityfromDB
		{
			set
			{
				pAvailability = value;
			}
		}
		public DateTime Date_
		{
			get
			{
				return pDate.theValue;
			}
			set
			{
				pDate.theValue = value;
			}
		}
		public int ShiftType
		{
			get
			{
				return pShiftType.theValue;
			}
			set
			{
				pShiftType.theValue = value;
			}
		}
		
		public SDocAvailable(string aDocInitial, int aAvailability, DateTime aDate, int aShiftType)
		{
			pAvailability = PublicEnums.Availability.Dispo;
			pDocInitial.theSQLName = PublicConstants.SQLInitials;
			pDate.theSQLName = PublicConstants.SQLDate;
			pShiftType.theSQLName = PublicConstants.SQLShiftType;
			
			DocInitial = aDocInitial;
			Availability = (PublicEnums.Availability) aAvailability;
			Date_ = aDate;
			ShiftType = aShiftType;
		}
		public SDocAvailable(DateTime aDate)
		{
			pDocInitial.theSQLName = PublicConstants.SQLInitials;
			pDate.theSQLName = PublicConstants.SQLDate;
			pShiftType.theSQLName = PublicConstants.SQLShiftType;
			pAvailability = PublicEnums.Availability.Dispo;
			Date_ = aDate;
		}
		public SDocAvailable()
		{
			pDocInitial.theSQLName = PublicConstants.SQLInitials;
			pDate.theSQLName = PublicConstants.SQLDate;
			pShiftType.theSQLName = PublicConstants.SQLShiftType;
			pAvailability = PublicEnums.Availability.Dispo;
		}
		public void UpdateScheduleDataTable(int theAvail)
		{
			//check if an entry already exists for this date and shift
			SQLStrBuilder theBuiltSql = new SQLStrBuilder();
			ADODB.Recordset theRS = new ADODB.Recordset();
			DBAC theDBAC = new DBAC();
			
			theBuiltSql.SQL_Select("*");
			theBuiltSql.SQL_From(PublicConstants.TABLE_ScheduleData);
			theBuiltSql.SQL_Where(pDate.theSQLName, "=", Date_);
			theBuiltSql.SQL_Where(pShiftType.theSQLName, "=", ShiftType);
			theDBAC.COpenDB(theBuiltSql.SQLStringSelect, theRS);
			
			int theCount = theRS.RecordCount;
			
			switch (theCount)
			{
				case 0: //if not create a new entry
					theBuiltSql.SQLClear();
					theBuiltSql.SQL_Insert(PublicConstants.TABLE_ScheduleData);
					theBuiltSql.SQL_Values(pDate.theSQLName, Date_);
					theBuiltSql.SQL_Values(pShiftType.theSQLName, ShiftType);
					theBuiltSql.SQL_Values(pDocInitial.theSQLName, DocInitial);
					int numaffected = default(int);
					theDBAC.CExecuteDB(theBuiltSql.SQLStringInsert, numaffected);
					//Debug.WriteLine(.SQLStringInsert)
					//Debug.WriteLine("Number of databaseentries" + numaffected.ToString())
					break;
					
				case 1: //if yes update it with the new value
					theRS.Fields[pDocInitial.theSQLName].Value = pDocInitial.theValue;
					theRS.ActiveConnection = theDBAC.aConnection;
					theRS.UpdateBatch((ADODB.AffectEnum) 3);
					theRS.Close();
					break;
				default:
					break;
					//Debug.WriteLine("there is more than one copy of this entry ... this is bad")
					
					
			}
		}
		public void DeleteScheduleDataEntry()
		{
			//check if an entry already exists for this date and shift
			SQLStrBuilder theBuiltSql = new SQLStrBuilder();
			ADODB.Recordset theRS = new ADODB.Recordset();
			DBAC theDBAC = new DBAC();
			
			
			theBuiltSql.SQL_From(PublicConstants.TABLE_ScheduleData);
			theBuiltSql.SQL_Where(pDate.theSQLName, "=", Date_);
			theBuiltSql.SQL_Where(pShiftType.theSQLName, "=", ShiftType);
			int numaffected = default(int);
			theDBAC.CExecuteDB(theBuiltSql.SQLStringDelete, numaffected);
		}
		public Collection doesDataExistForThisMonth()
		{
			
			SQLStrBuilder theBuiltSql = new SQLStrBuilder();
			ADODB.Recordset theRS = new ADODB.Recordset();
			DBAC theDBAC = new DBAC();
			DateTime theStartdate = DateAndTime.DateSerial(Date_.Year, Date_.Month, 1);
			DateTime theStopdate = DateAndTime.DateSerial(Date_.Year, Date_.Month + 1, 1);
			theBuiltSql.SQL_Select("*");
			theBuiltSql.SQL_From(PublicConstants.TABLE_ScheduleData);
			theBuiltSql.SQL_Where(pDate.theSQLName, ">=", theStartdate);
			theBuiltSql.SQL_Where(pDate.theSQLName, "<", theStopdate);
			theDBAC.COpenDB(theBuiltSql.SQLStringSelect, theRS);
			
			if (theRS.RecordCount > 0)
			{
				SDocAvailable aSDocAvailable = default(SDocAvailable);
				Collection aCollection = new Collection();
				theRS.MoveFirst();
				for (int x = 1; x <= theRS.RecordCount; x++)
				{
					aSDocAvailable = new SDocAvailable();
					aSDocAvailable.DocInitial = (theRS.Fields[this.pDocInitial.theSQLName].Value).ToString();
					aSDocAvailable.Date_ = System.Convert.ToDateTime(theRS.Fields[this.pDate.theSQLName].Value);
					aSDocAvailable.ShiftType = System.Convert.ToInt32(theRS.Fields[this.pShiftType.theSQLName].Value);
					aCollection.Add(aSDocAvailable, null, null, null);
					theRS.MoveNext();
				}
				return aCollection;
			}
			return null;
		}
		
	}
	
	public class SNonDispo
	{
		private PublicStructures.T_DBRefTypeS pDocInitial;
		private PublicStructures.T_DBRefTypeD pDateStart;
		private PublicStructures.T_DBRefTypeD pDateStop;
		private PublicStructures.T_DBRefTypeI pTimeStart;
		private PublicStructures.T_DBRefTypeI pTimeStop;
		private string pDu;
		private string pAu;
		
		public string du
		{
			get
			{
				int myhours = (int) (pTimeStart.theValue / 60);
				int myminutes = pTimeStart.theValue - (myhours * 60);
				DateTime atime = new DateTime(1, 1, 1, myhours, myminutes, 0);
				string astr = MyGlobals.daystrings[pDateStart.theValue.DayOfWeek] + " le " + pDateStart.theValue.Day.ToString() + " " + MyGlobals.monthstrings[pDateStart.theValue.Month - 1] + " " + pDateStart.theValue.Year.ToString() + " à " + "0" + atime.Hour.ToString().Substring("0" + atime.Hour.ToString().Length - 2, 2) + ":" + "0" + atime.Minute.ToString().Substring("0" + atime.Minute.ToString().Length - 2, 2);
				
				return astr;
			}
		}
		public string au
		{
			get
			{
				int myhours = (int) (pTimeStop.theValue / 60);
				int myminutes = pTimeStop.theValue - (myhours * 60);
				DateTime atime = new DateTime(1, 1, 1, myhours, myminutes, 0);
				string astr = MyGlobals.daystrings[pDateStop.theValue.DayOfWeek] + " le " + pDateStop.theValue.Day.ToString() + " " + MyGlobals.monthstrings[pDateStop.theValue.Month - 1] + " " + pDateStop.theValue.Year.ToString() + " à " + "0" + atime.Hour.ToString().Substring("0" + atime.Hour.ToString().Length - 2, 2) + ":" + "0" + atime.Minute.ToString().Substring("0" + atime.Minute.ToString().Length - 2, 2);
				return astr;
			}
		}
		public string DocInitial
		{
			get
			{
				return pDocInitial.theValue;
			}
			set
			{
				pDocInitial.theValue = value;
			}
		}
		public DateTime DateStart
		{
			get
			{
				return pDateStart.theValue;
			}
			set
			{
				pDateStart.theValue = value;
			}
		}
		public DateTime DateStop
		{
			get
			{
				return pDateStop.theValue;
			}
			set
			{
				pDateStop.theValue = value;
			}
		}
		public int TimeStart
		{
			get
			{
				return pTimeStart.theValue;
			}
			set
			{
				pTimeStart.theValue = value;
			}
		}
		public int TimeStop
		{
			get
			{
				return pTimeStop.theValue;
			}
			set
			{
				pTimeStop.theValue = value;
			}
		}
		
		public SNonDispo(string aDocInitial, DateTime aDateStart, DateTime aDateStop, int aTimeStart, int aTimeStop)
		{
			
			pDocInitial.theSQLName = PublicConstants.SQLInitials;
			pDateStart.theSQLName = PublicConstants.SQLDateStart;
			pDateStop.theSQLName = PublicConstants.SQLDateStop;
			pTimeStart.theSQLName = PublicConstants.SQLTimeStart;
			pTimeStop.theSQLName = PublicConstants.SQLTimeStop;
			
			DocInitial = aDocInitial;
			DateStart = aDateStart;
			DateStop = aDateStop;
			TimeStart = aTimeStart;
			TimeStop = aTimeStop;
			if (IsUnique())
			{
				
				SQLStrBuilder theBuiltSql = new SQLStrBuilder();
				DBAC theDBAC = new DBAC();
				
				theBuiltSql.SQLClear();
				theBuiltSql.SQL_Insert(PublicConstants.Table_NonDispo);
				theBuiltSql.SQL_Values(pDocInitial.theSQLName, DocInitial);
				theBuiltSql.SQL_Values(pDateStart.theSQLName, DateStart);
				theBuiltSql.SQL_Values(pTimeStart.theSQLName, TimeStart);
				theBuiltSql.SQL_Values(pDateStop.theSQLName, DateStop);
				theBuiltSql.SQL_Values(pTimeStop.theSQLName, TimeStop);
				
				int numaffected = default(int);
				theDBAC.CExecuteDB(theBuiltSql.SQLStringInsert, numaffected);
			}
		}
		public SNonDispo()
		{
			pDocInitial.theSQLName = PublicConstants.SQLInitials;
			pDateStart.theSQLName = PublicConstants.SQLDateStart;
			pDateStop.theSQLName = PublicConstants.SQLDateStop;
			pTimeStart.theSQLName = PublicConstants.SQLTimeStart;
			pTimeStop.theSQLName = PublicConstants.SQLTimeStop;
		}
		private bool IsUnique()
		{
			SQLStrBuilder theBuiltSql = new SQLStrBuilder();
			ADODB.Recordset theRS = new ADODB.Recordset();
			DBAC theDBAC = new DBAC();
			theBuiltSql.SQLClear();
			theBuiltSql.SQL_Select(pDocInitial.theSQLName);
			theBuiltSql.SQL_From(PublicConstants.Table_NonDispo);
			theBuiltSql.SQL_Where(pDocInitial.theSQLName, "=", DocInitial);
			theBuiltSql.SQL_Where(pDateStart.theSQLName, "=", DateStart);
			theBuiltSql.SQL_Where(pTimeStart.theSQLName, "=", TimeStart);
			theBuiltSql.SQL_Where(pDateStop.theSQLName, "=", DateStop);
			theBuiltSql.SQL_Where(pTimeStop.theSQLName, "=", TimeStop);
			theDBAC.COpenDB(theBuiltSql.SQLStringSelect, theRS);
			if (theRS.RecordCount > 0)
			{
				return false;
			}
			else
			{
				return true;
			}
		}
		public Collection GetNonDispoListForDoc(string aDocInitials, int aYear, int aMonth)
		{
			SQLStrBuilder theBuiltSql = new SQLStrBuilder();
			ADODB.Recordset theRS = new ADODB.Recordset();
			DBAC theDBAC = new DBAC();
			DateTime theStartdate = DateAndTime.DateSerial(aYear, aMonth, 1);
			DateTime theStopdate = DateAndTime.DateSerial(aYear, aMonth + 1, 1);
			theBuiltSql.SQL_Select("*");
			theBuiltSql.SQL_From(PublicConstants.Table_NonDispo);
			theBuiltSql.SQL_Where(pDocInitial.theSQLName, "=", aDocInitials);
			theBuiltSql.SQL_Where(pDateStop.theSQLName, ">=", theStartdate);
			theBuiltSql.SQL_Where(pDateStart.theSQLName, "<", theStopdate);
			theBuiltSql.SQL_Order_By(pDateStart.theSQLName);
			theBuiltSql.SQL_Order_By(pTimeStart.theSQLName);
			theDBAC.COpenDB(theBuiltSql.SQLStringSelect, theRS);
			SNonDispo aSNonDispo = default(SNonDispo);
			int theCount = theRS.RecordCount;
			if (theCount > 0)
			{
				Collection aCollection = new Collection();
				theRS.MoveFirst();
				for (int x = 1; x <= theCount; x++)
				{
					aSNonDispo = new SNonDispo();
					if (!Information.IsDBNull(theRS.Fields[pDocInitial.theSQLName].Value))
					{
						aSNonDispo.DocInitial = (theRS.Fields[pDocInitial.theSQLName].Value).ToString();
					}
					if (!Information.IsDBNull(theRS.Fields[pDateStart.theSQLName].Value))
					{
						aSNonDispo.DateStart = System.Convert.ToDateTime(theRS.Fields[pDateStart.theSQLName].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[pTimeStart.theSQLName].Value))
					{
						aSNonDispo.TimeStart = System.Convert.ToInt32(theRS.Fields[pTimeStart.theSQLName].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[pDateStop.theSQLName].Value))
					{
						aSNonDispo.DateStop = System.Convert.ToDateTime(theRS.Fields[pDateStop.theSQLName].Value);
					}
					if (!Information.IsDBNull(theRS.Fields[pTimeStop.theSQLName].Value))
					{
						aSNonDispo.TimeStop = System.Convert.ToInt32(theRS.Fields[pTimeStop.theSQLName].Value);
					}
					aCollection.Add(aSNonDispo, x.ToString(), null, null);
					theRS.MoveNext();
				}
				return aCollection;
			}
			else
			{
				return null;
			}
		}
		public void Delete()
		{
			SQLStrBuilder theBuiltSql = new SQLStrBuilder();
			DBAC theDBAC = new DBAC();
			int numaffected = default(int);
			theBuiltSql.SQL_From(PublicConstants.Table_NonDispo);
			theBuiltSql.SQL_Where(pDocInitial.theSQLName, "=", DocInitial);
			theBuiltSql.SQL_Where(pDateStop.theSQLName, "=", DateStop);
			theBuiltSql.SQL_Where(pDateStart.theSQLName, "=", DateStart);
			theBuiltSql.SQL_Where(pTimeStop.theSQLName, "=", TimeStop);
			theBuiltSql.SQL_Where(pTimeStart.theSQLName, "=", TimeStart);
			theDBAC.CExecuteDB(theBuiltSql.SQLStringDelete, numaffected);
		}
		
	}
	
	public class DBAC
	{
		
		//Public cnn As New ADODB.Connection  'Connection object definition
		//Public rs As New ADODB.Recordset    'recordset object definition
		
		const string Provider = "Provider=Microsoft.ACE.OLEDB.12.0;";
		//Const DBpassword = "Jet OLEDB:Database Password=plasma;"
		
		private long theConnectionState;
		private string mConnectionString;
		private ADODB.Connection mConnection;
		public ADODB.Connection aConnection
		{
			get
			{
				return mConnection;
			}
		}
		
		public DBAC()
		{
			//On Error GoTo errhandler
			if (MyGlobals.cnn.State == (int) ADODB.ObjectStateEnum.adStateClosed)
			{
				
				
				if (PublicConstants.CONSTFILEADDRESS == "")
				{
					if (MyGlobals.MySettingsGlobal.DataBaseLocation == "")
					{
						GlobalFunctions.LoadDatabaseFileLocation();
					}
					PublicConstants.CONSTFILEADDRESS = MyGlobals.MySettingsGlobal.DataBaseLocation;
				}
				mConnectionString = Provider + "Data Source=" + PublicConstants.CONSTFILEADDRESS; //_
				//+ ";" & DBpassword
				MyGlobals.cnn.ConnectionString = mConnectionString;
				MyGlobals.cnn.Open("", "", "", -1);
			}
			
			mConnection = MyGlobals.cnn;
			//        On Error GoTo 0
			//        Exit Sub
			//errhandler:
			//        MsgBox("An error occurred during initial connection to DB: " + _
			//               CStr(Err.Number) + "  :  " + _
			//               CStr(Err.Description))
			
			//        'add code to select current location for the database !!FEATURE!!
			
		}
		public void COpenDB(string theSQLstr, ADODB.Recordset theRS)
		{
			
			//if myReadonly is true file is open as readonly, otherwise it is modifiable
			
			//Current DB address is hardcoded, there is
			//code in ThisWorkbook to allow DB selection at workbook open
			//DB address can be changed by simply changing the FileAddress assignment below
			
			if (theRS.State == (int) ADODB.ObjectStateEnum.adStateOpen)
			{
				theRS.Close();
			}
			theRS.CursorLocation = ADODB.CursorLocationEnum.adUseClient;
			
			//On Error GoTo errhandler
			//Debug.Print(theSQLstr)
			theRS.Open(theSQLstr, mConnection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, -1);
			theRS.ActiveConnection = null;
			
			// On Error GoTo 0
			//Exit Sub 'to not run errhandler for nothing
			//errhandler:
			
			//        MsgBox("An error occurred during SELECT query execution: " + _
			//               CStr(Err.Number) & "  :  " + _
			//               CStr(Err.Description) + _
			//               "   SQL TEXT:   " + _
			//               theSQLstr)
			
		}
		public void CExecuteDB(string theSQLstr, long numAffected)
		{
			
			//On Error GoTo errhandler
			object theNumAffectedObj = default(object);
			mConnection.Execute(theSQLstr, out theNumAffectedObj, -1);
			//StoreToAuditFile theSQLstr
			//On Error GoTo 0
			numAffected = System.Convert.ToInt32(theNumAffectedObj);
			//        Exit Sub 'to not run errhandler for nothing
			
			
			//errhandler:
			
			//        Dim theError As String
			//        theError = CStr(Err.Number) & "  :  " & CStr(Err.Description) & "   SQL TEXT:   " & theSQLstr
			//        theError = Replace(theError, "'", "''")
			//        MsgBox("An error occurred during execution of an INSERT or UPDATE query: " & theError)
			
			//        ''LogError theError
			
		}
		
		//    Private Sub StoreToAuditFile(theSQLstr As String)
		
		//        Dim theBuiltSql As New SQLStrBuilder
		//        Dim numaffected2
		//        theSQLstr = Replace(theSQLstr, "'", "''")
		//        With theBuiltSql
		
		//            .SQL_Insert(AUDITTABLE)
		//            .SQL_Values("TimeDateStamp", TimeToMillisecond)
		
		//            .SQL_Values("UserName", theUserName)
		//            .SQL_Values("Query", theSQLstr)
		
		//            On Error GoTo errhandler
		//            Dim theSQLString As String
		//            theSQLString = .SQLStringInsert
		//            mConnection.Execute(theSQLString, numaffected2)
		//        End With
		
		//        Exit Sub
		
		//errhandler:
		
		//        Dim theError As String
		//        theError = CStr(Err.Number) & "  :  " & CStr(Err.Description) & "   SQL TEXT:   " & theSQLString
		//        theError = Replace(theError, "'", "''")
		//        MsgBox("An error occurred during connection to DB: " & theError)
		
		//        LogError(theError)
		
		//    End Sub
		
		
		
		//Private Sub LogError(theError As String)
		
		//    Dim theBuiltSql As New SQLStrBuilder
		//    Dim numaffected2
		//    theError = Replace(theError, "'", "''")
		//    With theBuiltSql
		
		//        .SQL_Insert(AUDITTABLE)
		//        .SQL_Values("TimeDateStamp", TimeToMillisecond)
		
		//        .SQL_Values("UserName", "ERRORLOG")
		//        .SQL_Values("Query", theError)
		
		
		//        mConnection.Execute.SQLStringInsert, numaffected2
		//    End With
		
		//End Sub
		
		
		//Public Sub GetSchema(theREC As ADODB.Recordset)
		
		//    theREC = mConnection.OpenSchema(adSchemaColumns)
		//    ' theREC.ActiveConnection = Nothing
		
		//End Sub
		
	}
	
	public class SQLStrBuilder
	{
		
		private string theSelect;
		private int theSelectCounter;
		private string theFrom;
		private int theFromCounter;
		private string theWhere;
		private int theWhereCounter;
		private string theGroupBy;
		private int theGroupByCounter;
		private string theOrderBy;
		private int theOrderByCounter;
		private string theSet;
		private int theSetCounter;
		private string theUpdate;
		private int theUpdateCounter;
		private string theInsert;
		private int theInsertCounter;
		private string theValueS;
		private int theValuesCounter;
		private string theInto;
		private string theStrSQL;
		
		public string SQLStringSelect
		{
			get
			{
				if (theSelectCounter < 1 | theFromCounter < 1)
				{
					return "";
				}
				theStrSQL = theSelect + theFrom + theWhere + theGroupBy + theOrderBy;
				return theStrSQL;
			}
		}
		public string SQLStringUpdate
		{
			get
			{
				if (theUpdateCounter < 1 | theSetCounter < 1)
				{
					return "";
				}
				theStrSQL = theUpdate + theSet + theWhere + ";";
				return theStrSQL;
			}
		}
		public string SQLStringInsert
		{
			get
			{
				if (theInsertCounter < 1 | theValuesCounter < 1)
				{
					return "";
				}
				theStrSQL = theInsert + theInto + ") " + theValueS + ");";
				return theStrSQL;
			}
		}
		public string SQLStringDelete
		{
			get
			{
				theStrSQL = "DELETE" + theFrom + theWhere;
				return theStrSQL;
			}
		}
		
		public SQLStrBuilder()
		{
			
			theStrSQL = "";
			theSelect = "";
			theFrom = "";
			theWhere = "";
			theGroupBy = "";
			theOrderBy = "";
			theSet = "";
			theUpdate = "";
			theInsert = "";
			theValueS = "";
			theInto = "";
			
			theSelectCounter = 0;
			theFromCounter = 0;
			theWhereCounter = 0;
			theGroupByCounter = 0;
			theOrderByCounter = 0;
			theSetCounter = 0;
			theUpdateCounter = 0;
			theInsertCounter = 0;
			theValuesCounter = 0;
			
		}
		public void SQL_Select(string theColumnName)
		{
			
			if (theSelectCounter == 0)
			{
				theSelect = "SELECT " + theColumnName;
				theSelectCounter++;
			}
			else
			{
				theSelect = theSelect + ", " + theColumnName;
				theSelectCounter++;
			}
			
		}
		public void SQL_From(string theTableName)
		{
			
			if (theFromCounter == 0)
			{
				theFrom = " FROM " + theTableName;
				theFromCounter++;
			}
			else
			{
				theFrom = theFrom + ", " + theTableName;
				theFromCounter++;
			}
			
		}
		public void SQL_Where(string theColumnName, string theCondition, object theValue, string theOperator = "AND", PublicEnums.EnumWhereSubClause theSubclause = PublicEnums.EnumWhereSubClause.EW_None, int theParenthesesCount = 1, bool isFieldName = false)
		{
			
			string theValueStr = default(string);
			string theSubClauseStr_Begin = default(string);
			string theSubClauseStr_End = default(string);
			int theCounter = default(int);
			
			switch (Information.TypeName(theValue))
			{
				case "String":
					if (isFieldName == false)
					{
						theValueStr = "\'" + (theValue).ToString() + "\'";
					}
					else
					{
						theValueStr = (theValue).ToString();
					}
					break;
					
				case "Date":
					theValueStr = GlobalFunctions.cAccessDateStr(System.Convert.ToDateTime(theValue));
					break;
				case "Boolean":
					if (System.Convert.ToBoolean(theValue) == true)
					{
						theValueStr = "true";
					}
					else
					{
						theValueStr = "false";
					}
					break;
				default:
					theValueStr = (theValue).ToString();
					break;
			}
			theSubClauseStr_Begin = "";
			theSubClauseStr_End = "";
			switch (theSubclause)
			{
				case PublicEnums.EnumWhereSubClause.EW_None:
					theSubClauseStr_Begin = "";
					theSubClauseStr_End = "";
					break;
				case PublicEnums.EnumWhereSubClause.EW_begin:
					for (theCounter = 1; theCounter <= theParenthesesCount; theCounter++)
					{
						theSubClauseStr_Begin = theSubClauseStr_Begin + "(";
					}
					break;
					
				case PublicEnums.EnumWhereSubClause.EW_end:
					for (theCounter = 1; theCounter <= theParenthesesCount; theCounter++)
					{
						theSubClauseStr_End = theSubClauseStr_End + ")";
					}
					break;
			}
			
			if (theWhereCounter == 0)
			{
				theWhere = " WHERE " + theSubClauseStr_Begin + theColumnName + theCondition + theValueStr + theSubClauseStr_End;
				theWhereCounter++;
			}
			else
			{
				theWhere = theWhere + " " + theOperator + " " + theSubClauseStr_Begin + theColumnName + " " + theCondition + " " + theValueStr + theSubClauseStr_End;
				theWhereCounter++;
			}
			
		}
		public void SQL_Where_in(string theColumnName, string theItems)
		{
			
			if (theWhereCounter == 0)
			{
				theWhere = " WHERE " + theColumnName + theItems;
				theWhereCounter++;
			}
			else
			{
				theWhere = theWhere + " AND " + theColumnName + theItems;
				theWhereCounter++;
			}
			
		}
		public void SQL_Group_By(string theColumnName)
		{
			
			if (theGroupByCounter == 0)
			{
				theGroupBy = " GROUP BY " + theColumnName;
				theGroupByCounter++;
			}
			else
			{
				theGroupBy = theGroupBy + ", " + theColumnName;
				theGroupByCounter++;
			}
			
		}
		public void SQL_Order_By(string theColumnName, string SortOrder = "ASC")
		{
			
			if (theOrderByCounter == 0)
			{
				theOrderBy = " ORDER BY " + theColumnName + " " + SortOrder;
				theOrderByCounter++;
			}
			else
			{
				theOrderBy = theOrderBy + ", " + theColumnName + " " + SortOrder;
				theOrderByCounter++;
			}
			
		}
		public void SQL_Set(string theColumnName, object theValue)
		{
			
			string theValueStr = default(string);
			switch (Information.TypeName(theValue))
			{
				case "String":
					theValueStr = "\'" + (theValue).ToString() + "\'";
					break;
				case "Date":
					theValueStr = GlobalFunctions.cAccessDateStr(System.Convert.ToDateTime(theValue));
					break;
				default:
					theValueStr = (theValue).ToString();
					break;
			}
			
			if (theSetCounter == 0)
			{
				theSet = " SET " + theColumnName + "=" + theValueStr;
				theSetCounter++;
			}
			else
			{
				theSet = theSet + ", " + theColumnName + "=" + theValueStr;
				theSetCounter++;
			}
			
		}
		public void SQL_Update(string theTableName)
		{
			
			theUpdate = "UPDATE " + theTableName;
			theUpdateCounter++;
			
		}
		public void SQL_Insert(string theTableName)
		{
			
			theInsert = "INSERT INTO " + theTableName;
			theInsertCounter++;
			
		}
		public void SQL_Values(string theColumnName, object theValue)
		{
			
			string theValueStr = default(string);
			switch (Information.TypeName(theValue))
			{
				case "String":
					theValueStr = "\'" + (theValue).ToString() + "\'";
					break;
				case "Date":
					theValueStr = GlobalFunctions.cAccessDateStr(System.Convert.ToDateTime(theValue));
					break;
				default:
					theValueStr = (theValue).ToString();
					break;
			}
			
			if (theValuesCounter == 0)
			{
				theInto = " (" + theColumnName;
				theValueS = "VALUES (" + theValueStr;
				theValuesCounter++;
			}
			else
			{
				theInto = theInto + ", " + theColumnName;
				theValueS = theValueS + ", " + theValueStr;
				theValuesCounter++;
			}
			
		}
		public void SQLClear()
		{
			
			theStrSQL = "";
			theSelect = "";
			theFrom = "";
			theWhere = "";
			theGroupBy = "";
			theOrderBy = "";
			theSet = "";
			theUpdate = "";
			theInsert = "";
			theValueS = "";
			theInto = "";
			
			theSelectCounter = 0;
			theFromCounter = 0;
			theWhereCounter = 0;
			theGroupByCounter = 0;
			theOrderByCounter = 0;
			theSetCounter = 0;
			theUpdateCounter = 0;
			theInsertCounter = 0;
			theValuesCounter = 0;
			
		}
		
		
	}
	
	
	
	
	
	
}
