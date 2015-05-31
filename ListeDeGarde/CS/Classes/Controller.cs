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

using Microsoft.VisualBasic.CompilerServices;

namespace ListeDeGarde
{
	public class Controller
	{
		private Excel.Worksheet controlledExcelSheet;
		private SMonth controlledMonth;
		private bool monthloaded = false;
		private UserControl4 monthlystats;
		private Form2 theMonthlyStatsForm;
		private Collection SDocStatsCollection;
		private const long theRestTime = 432000000000;
		private string theHighlightedDoc;
		private Microsoft.Office.Tools.CustomTaskPane theCustomTaskPane;
		
		public SMonth aControlledMonth
		{
			get
			{
				return controlledMonth;
			}
		}
		public string pHighlightedDoc
		{
			get
			{
				return theHighlightedDoc;
			}
		}
		
		public Controller(Excel.Worksheet aSheet, int aYear, int aMonth, string aMonthString)
		{
			
			//load the sheet
			controlledExcelSheet = aSheet;
			controlledExcelSheet.Change += new System.EventHandler(this.controlledExcelSheet_Change);
			controlledExcelSheet.BeforeDelete += new System.EventHandler(this.controlledExcelSheet_BeforeDelete);
			
			//create a month
			controlledMonth = new SMonth(aMonth, aYear);
			
			//Load shift types collection into global
			//controlledShiftTypes = controlledMonth.ShiftTypes
			theHighlightedDoc = "";
			Globals.ThisAddIn.theCurrentController = this;
			resetSheet();
			
			
		}
		public void resetSheetExt()
		{
			//clear the sheet
			controlledExcelSheet.Unprotect();
			controlledExcelSheet.Cells.Clear();
			//create a month
			controlledMonth = new SMonth(controlledMonth.Month, controlledMonth.Year);
			theHighlightedDoc = "";
			Globals.ThisAddIn.theCurrentController = this;
			//Load shift types collection into global
			//controlledShiftTypes = controlledMonth.ShiftTypes
			resetSheet();
		}
		public void statsMensuelles()
		{
			
			//If theMonthlyStatsForm Is Nothing Then
			//    theMonthlyStatsForm = New Form2
			//Else
			//    theMonthlyStatsForm.Dispose()
			//    theMonthlyStatsForm = New Form2
			//End If
			//theMonthlyStatsForm.TopMost = True
			//theMonthlyStatsForm.Show()
			
			
			MonthlyDocStatsTPF MyTaskPaneView = default(MonthlyDocStatsTPF);
			MyTaskPaneView = new MonthlyDocStatsTPF();
			theCustomTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(MyTaskPaneView, "Statistiques");
			theCustomTaskPane.Visible = true;
			statsMensuellesUpdate();
			
			
		}
		public void HighLightDocAvailablilities(string Initials)
		{
			//cycle through the month and highlight everywhere theDoc is available.
			SDay aday = default(SDay);
			SShift aShift = default(SShift);
			SDocAvailable adocAvail = default(SDocAvailable);
			//Globals.ThisAddIn.Application.ScreenUpdating = False
			controlledExcelSheet.Unprotect();
			foreach (SDay tempLoopVar_aday in controlledMonth.Days)
			{
				aday = tempLoopVar_aday;
				foreach (SShift tempLoopVar_aShift in aday.Shifts)
				{
					aShift = tempLoopVar_aShift;
					foreach (SDocAvailable tempLoopVar_adocAvail in aShift.DocAvailabilities)
					{
						adocAvail = tempLoopVar_adocAvail;
						if (adocAvail.DocInitial == Initials)
						{
							if (adocAvail.Availability == PublicEnums.Availability.Dispo)
							{
								aShift.aRange.Interior.Color = Information.RGB(0, 233, 118);
							}
							else if (adocAvail.Availability == PublicEnums.Availability.Assigne)
							{
								aShift.aRange.Interior.Color = Information.RGB(0, 255, 255);
							}
							else if (adocAvail.Availability == PublicEnums.Availability.NonDispoPermanente)
							{
								aShift.aRange.Interior.Color = Information.RGB(220, 20, 60);
							}
							else if (adocAvail.Availability == PublicEnums.Availability.NonDispoTemporaire)
							{
								aShift.aRange.Interior.Color = Information.RGB(219, 112, 147);
							}
							else if (adocAvail.Availability == PublicEnums.Availability.SurUtilise)
							{
								aShift.aRange.Interior.Color = Information.RGB(209, 95, 238);
							}
							else
							{
							}
						}
						
					}
				}
			}
			theHighlightedDoc = Initials;
			//Globals.ThisAddIn.Application.ScreenUpdating = True
			statsMensuellesUpdate();
			controlledExcelSheet.Protect();
		}
		public void HighLightDocAvailSingleCell(SShift theShift, string Initials)
		{
			//cycle through the month and highlight everywhere theDoc is available.
			
			SDocAvailable adocAvail = default(SDocAvailable);
			foreach (SDocAvailable tempLoopVar_adocAvail in theShift.DocAvailabilities)
			{
				adocAvail = tempLoopVar_adocAvail;
				if (adocAvail.DocInitial == Initials)
				{
					if (adocAvail.Availability == PublicEnums.Availability.Dispo)
					{
						theShift.aRange.Interior.Color = Information.RGB(0, 233, 118);
					}
					else if (adocAvail.Availability == PublicEnums.Availability.Assigne)
					{
						theShift.aRange.Interior.Color = Information.RGB(0, 255, 255);
					}
					else if (adocAvail.Availability == PublicEnums.Availability.NonDispoPermanente)
					{
						theShift.aRange.Interior.Color = Information.RGB(220, 20, 60);
					}
					else if (adocAvail.Availability == PublicEnums.Availability.NonDispoTemporaire)
					{
						theShift.aRange.Interior.Color = Information.RGB(219, 112, 147);
					}
					else if (adocAvail.Availability == PublicEnums.Availability.SurUtilise)
					{
						theShift.aRange.Interior.Color = Information.RGB(209, 95, 238);
					}
					else
					{
					}
				}
			}
		}
		public void fixlist(SShift theShift)
		{
			string theSetValue = "";
			SDocAvailable theDocAvailable = default(SDocAvailable);
			string thelist = "";
			foreach (SDocAvailable tempLoopVar_theDocAvailable in theShift.DocAvailabilities)
			{
				theDocAvailable = tempLoopVar_theDocAvailable;
				if (theDocAvailable.Availability == PublicEnums.Availability.Dispo)
				{
					thelist = thelist + theDocAvailable.DocInitial + ",";
				}
				else if (theDocAvailable.Availability == PublicEnums.Availability.Assigne)
				{
					thelist = thelist + theDocAvailable.DocInitial + ",";
					theSetValue = theDocAvailable.DocInitial;
				}
				else
				{
				}
			}
			if (thelist.Length > 0)
			{
				thelist = thelist.Substring(0, thelist.Length - 1);
			}
			controlledExcelSheet.Unprotect();
			object with_1 = theShift.aRange.Validation;
			with_1.Delete();
			if (thelist != "")
			{
				with_1.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, thelist);
				with_1.IgnoreBlank = true;
				with_1.InCellDropdown = true;
				with_1.InputTitle = "";
				with_1.ErrorTitle = "";
				with_1.InputMessage = "";
				with_1.ErrorMessage = "";
				with_1.ShowInput = true;
				with_1.ShowError = true;
			}
			theShift.aRange.Locked = false;
			
		}
		
		private void controlledExcelSheet_Change(Excel.Range Target)
		{
			
			if (monthloaded == false)
			{
				return;
			}
			controlledExcelSheet.Unprotect();
			//System.Diagnostics.Debug.WriteLine("WithEvents: You Changed Cells " + Target.Address + " " + controlledExcelSheet.Name)
			SDay aday = default(SDay);
			SShift aShift = default(SShift);
			SDocAvailable adocAvail;
			bool anExitNotice = false;
			string firstDoc = "";
			
			foreach (SDay tempLoopVar_aday in controlledMonth.Days)
			{
				aday = tempLoopVar_aday;
				foreach (SShift tempLoopVar_aShift in aday.Shifts)
				{
					aShift = tempLoopVar_aShift;
					if (aShift.aRange.Address == Target.Address)
					{
						//make current Doc dispo again
						if (aShift.Doc != null)
						{
							if (aShift.DocAvailabilities.Contains(aShift.Doc))
							{
								adocAvail = (SDocAvailable) (aShift.DocAvailabilities[aShift.Doc]);
								adocAvail.Availability = PublicEnums.Availability.Dispo;
								firstDoc = aShift.Doc;
								anExitNotice = true;
							}
						}
						
						//'assign new doc
						if (Target.Value == null)
						{
							//adocAvail = aShift.DocAvailabilities.Item(firstDoc)
							//adocAvail.Availability = PublicEnums.Availability.Dispo
							fixAvailability(firstDoc, controlledMonth, aShift, firstDoc);
							aShift.Doc = "";
						}
						else
						{
							if (aShift.DocAvailabilities.Contains((Target.Value).ToString()))
							{
								adocAvail = (SDocAvailable) (aShift.DocAvailabilities[Target.Value]);
								adocAvail.Availability = PublicEnums.Availability.Assigne;
								fixAvailability((Target.Value).ToString(), controlledMonth, aShift, firstDoc);
								aShift.Doc = (Target.Value).ToString();
								anExitNotice = true;
							}
						}
					}
					if (anExitNotice == true)
					{
						break;
					}
				}
				if (anExitNotice == true)
				{
					break;
				}
			}
			if (anExitNotice == true)
			{
				// resetSheet()
				if (theCustomTaskPane != null)
				{
					if (theCustomTaskPane.Visible)
					{
						statsMensuellesUpdate();
					}
				}
			}
			controlledExcelSheet.Protect();
		}
		private void controlledExcelSheet_BeforeDelete()
		{
			
			Globals.ThisAddIn.theControllerCollection.Remove((string) controlledExcelSheet.Name);
			
		}
		private void theMonthlyStatsForm_close(System.Object sender, System.Windows.Forms.FormClosingEventArgs e)
		{
			SDocStatsCollection = null;
		}
		
		private void fixAvailability(string aDoc, SMonth aMonth, SShift ashift, string firstDoc = "")
		{
			DateTime theDate = ashift.aDate;
			int theShift = ashift.ShiftType;
			int theshiftStart = ashift.ShiftStart;
			int theshiftStop = ashift.ShiftStop;
			int theStartDay = theDate.Day - 1;
			int theStopDay = theDate.Day + 1;
			SShift myShift = default(SShift);
			SDay aDay = (SDay) (aMonth.Days[ashift.aDate.Day]);
			long nonDispoStart = default(long);
			long nonDispoStop = default(long);
			long shftStop = default(long);
			long shftStart = default(long);
			Collection RecheckCollection = new Collection();
			SShift RecheckShift = default(SShift);
			
			for (int x = ashift.aDate.Day - 1; x <= ashift.aDate.Day + 1; x++)
			{
				if (aMonth.Days.Contains(x.ToString()))
				{
					aDay = (SDay) (aMonth.Days[x.ToString()]);
					foreach (SShift tempLoopVar_myShift in aDay.Shifts)
					{
						myShift = tempLoopVar_myShift;
						
						nonDispoStart = ashift.aDate.Ticks + (ashift.ShiftStart) * 600000000 - theRestTime;
						nonDispoStop = ashift.aDate.Ticks + (ashift.ShiftStop) * 600000000 + theRestTime;
						shftStop = myShift.aDate.Ticks + (myShift.ShiftStop) * 600000000;
						shftStart = myShift.aDate.Ticks + (myShift.ShiftStart) * 600000000;
						SDocAvailable thedocAvail = default(SDocAvailable);
						
						if (firstDoc != "") //do opposite of the top one
						{
							//then check if this doc is assigned in prevous or next day
							//if yes redo fixavailability on either or both of those if not leave as is
							if (myShift.DocAvailabilities.Contains(firstDoc))
							{
								thedocAvail = (SDocAvailable) (myShift.DocAvailabilities[firstDoc]);
								if (thedocAvail.Availability == PublicEnums.Availability.Assigne)
								{
									RecheckCollection.Add(myShift, null, null, null);
								}
							}
						}
						
						if ((shftStart > nonDispoStart & shftStart < nonDispoStop) || (shftStop > nonDispoStart & shftStop < nonDispoStop) || (shftStart > nonDispoStart & shftStop < nonDispoStop))
						{
							
							thedocAvail = (SDocAvailable) (myShift.DocAvailabilities[aDoc]);
							if (thedocAvail.Availability != PublicEnums.Availability.NonDispoPermanente & thedocAvail.Availability != PublicEnums.Availability.Assigne)
							{
								thedocAvail.Availability = PublicEnums.Availability.NonDispoTemporaire;
								fixlist(myShift);
							}
							
							if (firstDoc != "") //do opposite of the top one
							{
								//then check if this doc is assigned in prevous or next day
								//if yes redo fixavailability on either or both of those if not leave as is
								if (myShift.DocAvailabilities.Contains(firstDoc))
								{
									thedocAvail = (SDocAvailable) (myShift.DocAvailabilities[firstDoc]);
									if (thedocAvail.Availability != PublicEnums.Availability.NonDispoPermanente & thedocAvail.Availability != PublicEnums.Availability.Assigne)
									{
										thedocAvail.Availability = PublicEnums.Availability.Dispo;
										
										fixlist(myShift);
										
									}
								}
							}
							
							
						}
						if (theHighlightedDoc != "")
						{
							HighLightDocAvailSingleCell(myShift, theHighlightedDoc);
						}
					}
				}
			}
			
			if (RecheckCollection.Count > 0)
			{
				foreach (SShift tempLoopVar_RecheckShift in RecheckCollection)
				{
					RecheckShift = tempLoopVar_RecheckShift;
					fixAvailability(firstDoc, aMonth, RecheckShift);
				}
			}
			
		}
		private void addBordersAroundRange(Excel.Range aRange)
		{
			
			object with_1 = aRange.Borders(Excel.XlBordersIndex.xlEdgeBottom);
			with_1.LineStyle = Excel.XlLineStyle.xlContinuous;
			with_1.Weight = Excel.XlBorderWeight.xlThin;
			with_1.ColorIndex = Excel.Constants.xlAutomatic;
			object with_2 = aRange.Borders(Excel.XlBordersIndex.xlEdgeTop);
			with_2.LineStyle = Excel.XlLineStyle.xlContinuous;
			with_2.Weight = Excel.XlBorderWeight.xlThin;
			with_2.ColorIndex = Excel.Constants.xlAutomatic;
			object with_3 = aRange.Borders(Excel.XlBordersIndex.xlEdgeLeft);
			with_3.LineStyle = Excel.XlLineStyle.xlContinuous;
			with_3.Weight = Excel.XlBorderWeight.xlThin;
			with_3.ColorIndex = Excel.Constants.xlAutomatic;
			object with_4 = aRange.Borders(Excel.XlBordersIndex.xlEdgeRight);
			with_4.LineStyle = Excel.XlLineStyle.xlContinuous;
			with_4.Weight = Excel.XlBorderWeight.xlThin;
			with_4.ColorIndex = Excel.Constants.xlAutomatic;
			
		}
		private void statsMensuellesUpdate()
		{
			//pour chaque medecin compter chaque type de shift
			
			if (theCustomTaskPane != null)
			{
				if (theCustomTaskPane.Visible == true)
				{
					
					Collection theDocCollection = SDoc.LoadAllDocsPerMonth(controlledMonth.Year, controlledMonth.Month);
					SDoc aSDoc = default(SDoc);
					SShift ashift = default(SShift);
					SDay aDay = default(SDay);
					SDocAvailable aDOcAvail;
					SDocStats theSDocStats = default(SDocStats);
					if (SDocStatsCollection == null)
					{
						SDocStatsCollection = new Collection();
						foreach (SDoc tempLoopVar_aSDoc in theDocCollection)
						{
							aSDoc = tempLoopVar_aSDoc;
							theSDocStats = new SDocStats(aSDoc.Initials, aSDoc.Shift1, aSDoc.Shift2, aSDoc.Shift3, aSDoc.Shift4, aSDoc.Shift5);
							SDocStatsCollection.Add(theSDocStats, aSDoc.Initials, null, null);
							
						}
					}
					else
					{
						foreach (SDocStats tempLoopVar_theSDocStats in SDocStatsCollection)
						{
							theSDocStats = tempLoopVar_theSDocStats;
							theSDocStats.shift1 = theSDocStats.shift1E;
							theSDocStats.shift2 = theSDocStats.shift2E;
							theSDocStats.shift3 = theSDocStats.shift3E;
							theSDocStats.shift4 = theSDocStats.shift4E;
							theSDocStats.shift5 = theSDocStats.shift5E;
						}
						
					}
					
					int docCount = 0;
					int shiftCount = 0;
					foreach (SDocStats tempLoopVar_theSDocStats in SDocStatsCollection)
					{
						theSDocStats = tempLoopVar_theSDocStats;
						foreach (SDay tempLoopVar_aDay in controlledMonth.Days)
						{
							aDay = tempLoopVar_aDay;
							shiftCount = 0;
							foreach (SShift tempLoopVar_ashift in aDay.Shifts)
							{
								ashift = tempLoopVar_ashift;
								if (ashift.ShiftType > 5)
								{
									break;
								}
								aDOcAvail = (SDocAvailable) (ashift.DocAvailabilities[theSDocStats.Initials]);
								if (aDOcAvail.Availability == PublicEnums.Availability.Assigne)
								{
									switch (ashift.ShiftType)
									{
										case 1:
											theSDocStats.shift1--;
											break;
										case 2:
											theSDocStats.shift2--;
											break;
										case 3:
											theSDocStats.shift3--;
											break;
										case 4:
											theSDocStats.shift4--;
											break;
										case 5:
											theSDocStats.shift5--;
											break;
									}
								}
								shiftCount++;
							}
						}
						docCount++;
					}
					
					//Dim bCollection As System.Windows.Forms.Control.ControlCollection = theMonthlyStatsForm.Controls
					//Dim aElementHost As System.Windows.Forms.Integration.ElementHost = bCollection(0)
					//monthlystats = aElementHost.Child
					//monthlystats.loadarray(SDocStatsCollection)
					int[] theArray = null;
					if (theHighlightedDoc != "")
					{
						
						Array.Resize(ref theArray, 4);
						int weekCount = 0;
						bool firstday = true;
						
						//go through each day of month
						foreach (SDay tempLoopVar_aDay in controlledMonth.Days)
						{
							aDay = tempLoopVar_aDay;
							
							//update counter on week change
							
							
							if ((int) aDay.theDate.DayOfWeek == 1 && firstday == false)
							{
								weekCount++;
								if (weekCount > 3)
								{
									Array.Resize(ref theArray, weekCount + 1);
								}
							}
							
							firstday = false;
							
							foreach (SShift tempLoopVar_ashift in aDay.Shifts)
							{
								ashift = tempLoopVar_ashift;
								if (ashift.ShiftType > 5)
								{
									break;
								}
								aDOcAvail = (SDocAvailable) (ashift.DocAvailabilities[theHighlightedDoc]);
								if (aDOcAvail.Availability == PublicEnums.Availability.Assigne)
								{
									//populate simple array of week counts
									theArray[weekCount] = theArray[weekCount] + 1;
								}
							}
						}
					}
					else
					{
						theArray = null;
					}
					
					
					
					System.Windows.Forms.Control.ControlCollection aCollection = theCustomTaskPane.Control.Controls;
					System.Windows.Forms.Integration.ElementHost bElementHost = (Windows.Forms.Integration.ElementHost) (aCollection[0]);
					MonthlyDocStatsTP theMonthlyDocStatsTP = (MonthlyDocStatsTP) bElementHost.Child;
					theMonthlyDocStatsTP.loadarray(SDocStatsCollection, theArray);
				}
			}
			
			
			
		}
		private void SetUpPermNonDispos()
		{
			SNonDispo theSNonDispo = new SNonDispo();
			SNonDispo aSNonDispo = default(SNonDispo);
			Collection aCollection = default(Collection);
			SDay aDay = default(SDay);
			SShift ashift = default(SShift);
			SDoc theSDoc = new SDoc(controlledMonth.Year, controlledMonth.Month);
			Collection docCollection = controlledMonth.DocList;
			SDoc aSDoc = default(SDoc);
			long nonDispoStart = default(long);
			long nonDispoStop = default(long);
			long shftStop = default(long);
			long shftStart = default(long);
			
			//For Each doc in the total collection of doctors
			foreach (SDoc tempLoopVar_aSDoc in docCollection)
			{
				aSDoc = tempLoopVar_aSDoc;
				
				//get the unavailability list for one doctor
				aCollection = theSNonDispo.GetNonDispoListForDoc(aSDoc.Initials, controlledMonth.Year, controlledMonth.Month);
				if (aCollection != null)
				{
					//iterate through the doctors list of unavailabilities
					foreach (SNonDispo tempLoopVar_aSNonDispo in aCollection)
					{
						aSNonDispo = tempLoopVar_aSNonDispo;
						int stopDay = default(int);
						int startday = default(int);
						if (aSNonDispo.DateStart.Month == controlledMonth.Month)
						{
							startday = aSNonDispo.DateStart.Day;
						}
						else if (aSNonDispo.DateStart.Month < controlledMonth.Month)
						{
							startday = 1;
						}
						if (aSNonDispo.DateStop.Month == controlledMonth.Month)
						{
							stopDay = aSNonDispo.DateStop.Day;
						}
						else if (aSNonDispo.DateStop.Month > controlledMonth.Month)
						{
							stopDay = System.DateTime.DaysInMonth(controlledMonth.Year, controlledMonth.Month);
						}
						
						for (int y = startday - 1; y <= stopDay; y++)
						{
							if (controlledMonth.Days.Contains((y).ToString()))
							{
								aDay = (SDay) (controlledMonth.Days[y]);
								foreach (SShift tempLoopVar_ashift in aDay.Shifts)
								{
									ashift = tempLoopVar_ashift;
									nonDispoStart = aSNonDispo.DateStart.Ticks + (aSNonDispo.TimeStart) * 600000000;
									nonDispoStop = aSNonDispo.DateStop.Ticks + (aSNonDispo.TimeStop) * 600000000;
									shftStop = ashift.aDate.Ticks + (ashift.ShiftStop) * 600000000;
									shftStart = ashift.aDate.Ticks + (ashift.ShiftStart) * 600000000;
									
									if ((shftStart > nonDispoStart & shftStart < nonDispoStop) || (shftStop > nonDispoStart & shftStop < nonDispoStop) || (shftStart > nonDispoStart & shftStop < nonDispoStop))
									{
										
										SDocAvailable thedocAvail = default(SDocAvailable);
										thedocAvail = (SDocAvailable) (ashift.DocAvailabilities[aSDoc.Initials]);
										//check if doc is assigned and ask to clear (provide some info.. make surutlisé
										if (thedocAvail.Availability != PublicEnums.Availability.Assigne)
										{
											thedocAvail.Availability = PublicEnums.Availability.NonDispoPermanente;
										}
										fixlist(ashift);
									}
								}
							}
						}
					}
				}
			}
		}
		private void resetSheet()
		{
			monthloaded = false; //set boolean toggle to false to stop event triggers
			controlledExcelSheet.Unprotect();
			string amonthstring = MyGlobals.monthstrings[aControlledMonth.Month - 1];
			// Globals.ThisAddIn.Application.ScreenUpdating = False
			controlledExcelSheet.Cells.Clear(); //clear the worksheet
			SDay theDay = default(SDay);
			int row = default(int);
			int col = 0;
			
			//get number of shifts
			int rowheight1 = SShiftType.ActiveShiftTypesCountPerMonth(aControlledMonth.Month, aControlledMonth.Year) + 1;
			//assign colwidth as 2
			int colwidth1 = 2;
			
			
			//populate the top left corner of sheet with year and month strings
			controlledExcelSheet.Range("A1").Value = amonthstring;
			controlledExcelSheet.Range("B1").Value = aControlledMonth.Year.ToString();
			
			//set top left corner of calendar
			Excel.Range theRangeA3 = controlledExcelSheet.Range("A3");
			Excel.Range theRange = default(Excel.Range);
			
			//create the month to display in worksheet
			foreach (SDay tempLoopVar_theDay in controlledMonth.Days)
			{
				theDay = tempLoopVar_theDay;
				if ((theDay.theDate.DayOfWeek) == 0)
				{
					col = 6;
				}
				else
				{
					col = (int) ((int) theDay.theDate.DayOfWeek - 1);
				}
				theRange = theRangeA3.Offset(row * rowheight1, col * colwidth1);
				
				Excel.Range theRangeForShiftType;
				Excel.Range TheRAngeForDocLists = default(Excel.Range);
				SShift theShift = default(SShift);
				
				foreach (SShift tempLoopVar_theShift in theDay.Shifts)
				{
					theShift = tempLoopVar_theShift;
					theRangeForShiftType = theRange.Offset(theShift.ShiftType, 0);
					theRangeForShiftType.Value2 = "\'" + theShift.Description;
					TheRAngeForDocLists = theRange.Offset(theShift.ShiftType, 1);
					theShift.aRange = TheRAngeForDocLists;
					
					fixlist(theShift);
				}
				
				theRange.Offset(0, colwidth1 - 1).Value = theDay.theDate.Day;
				theRange.Offset(0, colwidth1 - 1).Interior.Color = Information.RGB(160, 160, 160);
				theRange.Offset(0, colwidth1 - 2).Value = MyGlobals.daystrings[theDay.theDate.DayOfWeek];
				theRange.Offset(0, colwidth1 - 2).Interior.Color = Information.RGB(160, 160, 160);
				theRange = theRange.Resize(rowheight1, colwidth1);
				addBordersAroundRange(theRange);
				if (col == 6)
				{
					row++;
				}
			}
			SetupAssignedDocs();
			SetUpPermNonDispos();
			//Globals.ThisAddIn.Application.ScreenUpdating = True
			
			monthloaded = true;
			controlledExcelSheet.Protect(true, true, true, false, false, false, false, AllowInsertingRows false, false, false, true, false, false, false);
			controlledExcelSheet.EnableSelection = Microsoft.Office.Interop.Excel.XlEnableSelection.xlUnlockedCells;
			
		}
		private void SetupAssignedDocs()
		{
			SDocAvailable aTest = new SDocAvailable(DateAndTime.DateSerial(aControlledMonth.Year, aControlledMonth.Month, 1));
			Collection aCollection = default(Collection);
			SDay theDay2 = default(SDay);
			SShift theShift2 = default(SShift);
			SDocAvailable theDocAvailble;
			aCollection = aTest.doesDataExistForThisMonth();
			if (aCollection != null)
			{
				SDocAvailable theAssignedDocs = default(SDocAvailable);
				foreach (SDocAvailable tempLoopVar_theAssignedDocs in aCollection)
				{
					theAssignedDocs = tempLoopVar_theAssignedDocs;
					theDay2 = (SDay) (controlledMonth.Days[theAssignedDocs.Date_.Day]);
					if (theDay2.Shifts.Contains(theAssignedDocs.ShiftType.ToString()))
					{
						theShift2 = (SShift) (theDay2.Shifts[theAssignedDocs.ShiftType.ToString()]);
						theShift2.Doc = theAssignedDocs.DocInitial;
						if (theShift2.DocAvailabilities.Contains(theAssignedDocs.DocInitial))
						{
							theDocAvailble = (SDocAvailable) (theShift2.DocAvailabilities[theAssignedDocs.DocInitial]);
							theDocAvailble.SetAvailabilityfromDB = PublicEnums.Availability.Assigne;
							theShift2.aRange.Value = theAssignedDocs.DocInitial;
							fixAvailability(theShift2.Doc, controlledMonth, theShift2);
						}
						else
						{
							Windows.MessageBox.Show("Un medecin avec les initialles " + theAssignedDocs.DocInitial + " Etait assigné au quart de travail " + theShift2.Description.ToString() + " le " + theDay2.theDate.Day.ToString() + ", mais le medecin a été retiré de la liste des médecins. Son assignation au quart de travail a été retiré.");
							SDocAvailable aSDocAvailable = new SDocAvailable(" ", PublicEnums.Availability.Assigne, theDay2.theDate, theShift2.ShiftType);
							aSDocAvailable.DeleteScheduleDataEntry();
						}
					}
				}
			}
		}
		private void ClearAvailability()
		{
			SDay aDay = default(SDay);
			SShift ashift = default(SShift);
			SDocAvailable aDocAvail = default(SDocAvailable);
			
			foreach (SDay tempLoopVar_aDay in aControlledMonth.Days)
			{
				aDay = tempLoopVar_aDay;
				foreach (SShift tempLoopVar_ashift in aDay.Shifts)
				{
					ashift = tempLoopVar_ashift;
					foreach (SDocAvailable tempLoopVar_aDocAvail in ashift.DocAvailabilities)
					{
						aDocAvail = tempLoopVar_aDocAvail;
						aDocAvail.Availability = PublicEnums.Availability.Dispo;
					}
					fixlist(ashift);
				}
			}
		}
	}
	
}
