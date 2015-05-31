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





namespace ListeDeGarde
{
	sealed class UseDelegateSM
	{
		
		//==================================================================
		//Demonstrates Using a Delegate for Event Handling
		//==================================================================
		
		private static Excel.Application xlApp;
		private static Excel.Workbook xlBook;
		private static Excel.Worksheet xlSheet1;
		private static Excel.AppEvents_WorkbookBeforeCloseEventHandler EventDel_BeforeBookClose;
		private static Excel.DocEvents_ChangeEventHandler EventDel_CellsChange;
		
		public static void UseDelegate()
		{
			
			
			
			//Start Excel and create a new workbook.
			xlApp = Globals.ThisAddIn.Application;
			UseDelegateSM.xlApp.WorkbookOpen += new System.EventHandler(this.xlApp_Workbookopen);
			UseDelegateSM.xlApp.WorkbookBeforeClose += new System.EventHandler(this.xlApp_WorkbookBeforeClose);
			UseDelegateSM.xlApp.SheetActivate += new System.EventHandler(this.xlApp_SheetActivate);
			xlBook = Globals.ThisAddIn.Application.ActiveWorkbook;
			xlBook.Windows(1).Caption = "Uses UseDelegate";
			
			//Get references to the three worksheets.
			xlSheet1 = (global::Microsoft.Office.Interop.Excel.Worksheet) Globals.ThisAddIn.Application.ActiveSheet;
			
			//Add an event handler for the WorkbookBeforeClose Event of the
			//Application object.
			EventDel_BeforeBookClose = new Excel.AppEvents_WorkbookBeforeCloseEventHandler(BeforeBookClose);
			xlApp.WorkbookBeforeClose += new System.EventHandler(EventDel_BeforeBookClose);
			
			//Add an event handler for the Change event of both Worksheet
			//objects.
			EventDel_CellsChange = new Excel.DocEvents_ChangeEventHandler(CellsChange);
			xlSheet1.Change += new System.EventHandler(EventDel_CellsChange);
			
			//Make Excel visible and give the user control.
			//xlApp.Visible = True
			//xlApp.UserControl = True
		}
		
		private static void CellsChange(Excel.Range Target)
		{
			//This is called when a cell or cells on a worksheet are changed.
			//System.Diagnostics.Debug.WriteLine("Delegate: You Changed Cells " + Target.Address + " on " + _
			// Target.Worksheet.Name())
		}
		
		private static void BeforeBookClose(Excel.Workbook Wb, bool Cancel)
		{
			//This is called when you choose to close the workbook in Excel.
			//The event handlers are removed, and then the workbook is closed
			//without saving changes.
			//System.Diagnostics.Debug.WriteLine("Delegate: Closing the workbook and removing event handlers.")
			xlSheet1.Change -= new System.EventHandler(EventDel_CellsChange);
			xlApp.WorkbookBeforeClose -= new System.EventHandler(EventDel_BeforeBookClose);
			Wb.Saved = true; //Set the dirty flag to true so there is no prompt to save.
		}
		
		
		//    Dim xlApp As Excel.Application
		//2.         Dim xlBook As Excel.Workbook
		//3.         Dim xlSheet As Excel.Worksheet
		//4.         Dim xlButton As Excel.OLEObject
		//5.         Dim iStartLine As Long
		//6.         xlApp = New Excel.Application
		//7.         xlApp.Visible = True
		//8.         xlBook = xlApp.Workbooks.Add
		//9.         xlSheet = xlBook.ActiveSheet
		//10.         xlButton = xlSheet.OLEObjects.Add(ClassType:="Forms.CommandButton.1", _
		//11.             Link:=False, DisplayAsIcon:=False, Left:=30, Top:=20, Width:=72, Height:=24)
		//12.         xlButton.Name = "BtnTest"
		//13.         xlButton.Object.Caption = "Press"
		//14.         With xlBook.VBProject.VBComponents.Item(xlSheet.CodeName).CodeModule
		//15.             iStartLine = .CreateEventProc("Click", "BtnTest") + 1
		//16.             .InsertLines(iStartLine, "Msgbox ""Hi""")
		//17.         End With
		//18.         xlApp.VBE.MainWindow.Visible = False
		
		
		
	}
	
}
