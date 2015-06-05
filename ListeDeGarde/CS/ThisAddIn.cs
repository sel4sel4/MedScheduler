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
	public partial class ThisAddIn
	{
		//--------------------------------FIELDS-------------------------------------
		public Excel.Application xlApp;
		public Excel.Workbook xlBook;
		public Excel.Worksheet xlSheet1;
		public List<Controller> theControllerCollection;
		public Controller theCurrentController;
		public Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;
		private ThisAddinHelper myThisAddinHelper;
		//--------------------------------PROPERTIES-------------------------------------
		public Microsoft.Office.Tools.CustomTaskPane taskpane
		{
			get
			{
				return myCustomTaskPane;
			}
		}
		
		
		//--------------------------------EVENTS-------------------------------------
		
		private void ThisAddIn_Startup(System.Object sender, System.EventArgs e)
		{
			//Create the task pane to create monthly Calendar
			MyGlobals.MyAddin = this;
			YearMonthPicker MyTaskPaneView = default(YearMonthPicker);
			MyTaskPaneView = new YearMonthPicker();
			myCustomTaskPane = this.CustomTaskPanes.Add(MyTaskPaneView, "Liste de Garde");
			myCustomTaskPane.Visible = true;
			//Load xlApp into the global variable.
			xlApp = Globals.ThisAddIn.Application;
			
			//create a new Controller collection
			theControllerCollection = new List<Controller>();
			
			//Initialize the persistent settings (stores database location)
			MyGlobals.MySettingsGlobal = new Settings1();
			MyGlobals.MySettingsGlobal.SettingsLoaded += new System.Configuration.SettingsLoadedEventHandler(MyGlobals.MySettingsGlobal.Settings1_SettingsLoaded);
			
			
		}
		private void ThisAddIn_Shutdown(System.Object sender, System.EventArgs e)
		{
			
		}
		private void xlApp_Workbookopen(Excel.Workbook Wb)
		{
			xlBook = Globals.ThisAddIn.Application.ActiveWorkbook;
			xlSheet1 = (global::Microsoft.Office.Interop.Excel.Worksheet) Globals.ThisAddIn.Application.ActiveSheet;
			
		}
		private void xlApp_WorkbookBeforeClose(Excel.Workbook Wb, ref bool Cancel)
		{
			Wb.Saved = true; //Set the dirty flag to true so there is no prompt to save. all Data is kept in th access DB anyway
		}
		private void xlApp_SheetActivate(object Obb)
		{
			
			//need to rebuild the taskpane on the basis of the currentlyselected month
			//code below retreives the handle to the UserControl to trigger redraw() public function
			System.Windows.Forms.Control.ControlCollection aCollection = myCustomTaskPane.Control.Controls;
			System.Windows.Forms.Integration.ElementHost aElementHost = (System.Windows.Forms.Integration.ElementHost) (aCollection[0]);
			YearMonthPickerC theYearMonthPickerC = (YearMonthPickerC) aElementHost.Child;
			theYearMonthPickerC.redraw();
			Excel.Worksheet theActivatedSheet = (Excel.Worksheet) Obb;
			
			if (!Globals.ThisAddIn.theControllerCollection.Exists(xy => theActivatedSheet.Name == xy.aControlledExcelSheet.Name))
			{
				return;
			}
			theCurrentController = Globals.ThisAddIn.theControllerCollection.Find(xy => theActivatedSheet.Name == xy.aControlledExcelSheet.Name);
			
		}
		protected override object RequestComAddInAutomationService()
		{
			if (myThisAddinHelper == null)
			{
				myThisAddinHelper = new ThisAddinHelper();
			}
			return myThisAddinHelper;
		}
		
		//--------------------------------PUBLIC METHODS-------------------------------------
	}
	
	[InteropServices.ComVisible(true)][Runtime.InteropServices.InterfaceType(Runtime.InteropServices.ComInterfaceType.InterfaceIsDual)][Guid("de4491a3-4ada-485a-a0cb-bb67f15d6e00")]public interface IThisAddinHelper
	{
		void Launch();
		void testclick();
	}
	
	
	[InteropServices.ComVisible(true)][InteropServices.ClassInterface(Runtime.InteropServices.ClassInterfaceType.None)][Guid("9ED54F84-A85D-4fcd-A854-44251E925F09")]public class ThisAddinHelper : System.Runtime.InteropServices.StandardOleMarshalObject, IThisAddinHelper
	{
		
		
		public void Launch()
		{
			
			Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
			if (MyGlobals.MySettingsGlobal.DataBaseLocation == "")
			{
				GlobalFunctions.LoadDatabaseFileLocation();
			}
			else
			{
				PublicConstants.CONSTFILEADDRESS = MyGlobals.MySettingsGlobal.DataBaseLocation;
			}
			
			//if sheet already exists exit
			//If Globals.ThisAddIn.theControllerCollection.Exists("Avril" + "-" + "2010") Then Exit Sub
			if (Globals.ThisAddIn.theControllerCollection.Exists(xy => xy.aControlledMonth.Month == 4 && xy.aControlledMonth.Year == 2010))
			{
				return;
			}
			
			//create a new sheet
			Globals.ThisAddIn.xlSheet1 = (Excel.Worksheet) (wb.Sheets.Add(wb.Sheets(wb.Sheets.Count), 1, Excel.XlSheetType.xlWorksheet));
			
			//rename the new sheet
			Globals.ThisAddIn.xlSheet1.Name = "Avril" + "-" + "2010";
			
			Controller theController = new Controller(Globals.ThisAddIn.xlSheet1, int.Parse("2010"), 4, "avril");
			
			Globals.ThisAddIn.theControllerCollection.Add(theController);
			
		}
		
		public void testclick()
		{
			System.Windows.Forms.Control.ControlCollection aCollection = MyGlobals.MyAddin.myCustomTaskPane.Control.Controls;
			System.Windows.Forms.Integration.ElementHost bElementHost = (Windows.Forms.Integration.ElementHost) (aCollection[0]);
			YearMonthPickerC theYearMonthPickerC = (YearMonthPickerC) bElementHost.Child;
			theYearMonthPickerC.TestClick();
		}
	}
	
	
}
