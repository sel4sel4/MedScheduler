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
		public Excel.Application xlApp;
		public Excel.Workbook xlBook;
		public Excel.Worksheet xlSheet1;
		public Collection theControllerCollection;
		public Controller theCurrentController;
		private Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;
		
		public Microsoft.Office.Tools.CustomTaskPane taskpane
		{
			get
			{
				return myCustomTaskPane;
			}
		}
		
		private void ThisAddIn_Startup(System.Object sender, System.EventArgs e)
		{
			//Create the task pane to create monthly Calendar
			YearMonthPicker MyTaskPaneView = default(YearMonthPicker);
			MyTaskPaneView = new YearMonthPicker();
			myCustomTaskPane = this.CustomTaskPanes.Add(MyTaskPaneView, "Liste de Garde");
			myCustomTaskPane.Visible = true;
			//Load xlApp into the global variable.
			xlApp = Globals.ThisAddIn.Application;
			
			//create a new Controller collection
			theControllerCollection = new Collection();
			
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
			UserControl2 aUserControl2 = (UserControl2) aElementHost.Child;
			aUserControl2.redraw();
		}
		
		
		
	}
	
}
