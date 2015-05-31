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
	public partial class Form1
	{
		public Form1()
		{
			
			// This call is required by the designer.
			InitializeComponent();
			this.Text = "Veuillez enter les non-disponibilitées";
			// Add any initialization after the InitializeComponent() call.
			
		}
		
		//Protected Overrides Sub Finalize()
		//    If Not Globals.ThisAddIn.theControllerCollection.Contains(Globals.ThisAddIn.Application.ActiveSheet.name) Then Exit Sub
		//    Dim aController As Controller = Globals.ThisAddIn.theControllerCollection.Item(Globals.ThisAddIn.Application.ActiveSheet.name)
		
		//    aController.resetSheetExt()
		//    MyBase.Finalize()
		//End Sub
	}
}
