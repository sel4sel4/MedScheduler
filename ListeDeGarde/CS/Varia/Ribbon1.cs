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

using Microsoft.Office.Tools.Ribbon;


namespace ListeDeGarde
{
	public partial class Ribbon1
	{
		private Form1 aform1;
		public void Ribbon1_Load(System.Object sender, RibbonUIEventArgs e)
		{
			
		}
		
		public void Button1_Click(object sender, RibbonControlEventArgs e)
		{
			// UseWithEvents()
			Globals.ThisAddIn.taskpane.Visible = true;
		}
		
		public void Button2_Click(object sender, RibbonControlEventArgs e)
		{
			//UseDelegate()
			aform1 = new Form1();
			aform1.FormClosing += new System.Windows.Forms.FormClosingEventHandler(aForm1_close);
			aform1.ShowDialog();
		}
		
		public void Button3_Click(object sender, RibbonControlEventArgs e)
		{
			GlobalFunctions.LoadDatabaseFileLocation();
		}
		
		public void Button4_Click(object sender, RibbonControlEventArgs e)
		{
			DrInterfaceForm aform1 = new DrInterfaceForm();
			aform1.ShowDialog();
		}
		
		public void ShiftButton_Click(object sender, RibbonControlEventArgs e)
		{
			ShiftInterfaceF aform1 = new ShiftInterfaceF();
			aform1.ShowDialog();
		}
		
		private void aForm1_close(System.Object sender, System.Windows.Forms.FormClosingEventArgs e)
		{
			if (!Globals.ThisAddIn.theControllerCollection.Contains((string) Globals.ThisAddIn.Application.ActiveSheet.name))
			{
				return;
			}
			Controller aController = Globals.ThisAddIn.theControllerCollection[Globals.ThisAddIn.Application.ActiveSheet.name];
			aController.resetSheetExt();
		}
		
		
		public void ExpectDoc_Click(object sender, RibbonControlEventArgs e)
		{
			Controller theController;
			if (Globals.ThisAddIn.theControllerCollection.Count < 1)
			{
				return;
			}
			if (!Globals.ThisAddIn.theControllerCollection.Contains((string) Globals.ThisAddIn.Application.ActiveSheet.name))
			{
				return;
			}
			theController = Globals.ThisAddIn.theControllerCollection[Globals.ThisAddIn.Application.ActiveSheet.name];
			
			DocExpectationsF aDocExpecationF = default(DocExpectationsF);
			aDocExpecationF = new DocExpectationsF();
			aDocExpecationF.Show();
		}
	}
	
}
