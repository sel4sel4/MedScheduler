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

using System.Windows.Controls;
using System.Diagnostics;



namespace ListeDeGarde
{
	public partial class UserControl2
	{
		
		private Button newBtn;
		private Controller theController;
		
		private void Button_Click(object sender, Windows.RoutedEventArgs e)
		{
			//get references to workbook
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
			if (Globals.ThisAddIn.theControllerCollection.Contains(this.combo2.Text + "-" + this.combo1.Text))
			{
				return;
			}
			
			//create a new sheet
			Globals.ThisAddIn.xlSheet1 = (Excel.Worksheet) (wb.Sheets.Add(wb.Sheets(wb.Sheets.Count), 1, Excel.XlSheetType.xlWorksheet));
			
			//rename the new sheet
			Globals.ThisAddIn.xlSheet1.Name = this.combo2.Text + "-" + this.combo1.Text;
			
			
			
			//Get list of doc names and create a button for each. Assign initals as button name
			//Dim theSDoc As New SDoc(CInt(Me.combo1.Text), Me.combo2.SelectedIndex + 1)
			//Dim aSDoc As SDoc
			//RemoveDocButtons()
			//For Each aSDoc In theSDoc.DocList
			//    Me.addDocButton(aSDoc.FirstName + " " + aSDoc.LastName, aSDoc.Initials)
			//Next
			Collection theDocCollection = SDoc.LoadAllDocsPerMonth(System.Convert.ToInt32(this.combo1.Text), System.Convert.ToInt32(this.combo2.SelectedIndex + 1));
			RemoveDocButtons();
			SDoc aSDoc = default(SDoc);
			foreach (SDoc tempLoopVar_aSDoc in theDocCollection)
			{
				aSDoc = tempLoopVar_aSDoc;
				this.addDocButton(aSDoc.FirstName + " " + aSDoc.LastName, aSDoc.Initials);
			}
			
			this.MoisAnnee.Content = this.combo1.Text + "-" + this.combo2.Text;
			
			//create a controller instance and add it to the global collection
			theController = new Controller(Globals.ThisAddIn.xlSheet1, System.Convert.ToInt32(this.combo1.Text), System.Convert.ToInt32(this.combo2.SelectedIndex + 1), (string) this.combo2.Text);
			
			Globals.ThisAddIn.theControllerCollection.Add(theController, (string) Globals.ThisAddIn.xlSheet1.Name, null, null);
			Initialles_Load();
			
		}
		
		private void combo1_Loaded(object sender, Windows.RoutedEventArgs e)
		{
			ComboBox theComboBox;
			theComboBox = (ComboBox) sender;
			theComboBox.ItemsSource = MyGlobals.yearstrings;
			theComboBox.SelectedIndex = 0;
		}
		
		private void addDocButton(string docName, string initials)
		{
			var newBtn = new Button();
			newBtn.Content = docName;
			newBtn.Name = initials;
			newBtn.Click += new System.Windows.RoutedEventHandler(onBtnClick);
			SplMain.Children.Add(newBtn);
			
		}
		
		private void RemoveDocButtons()
		{
			Button aButton = default(Button);
			for (int x = 0; x <= SplMain.Children.Count - 1; x++)
			{
				aButton = (Button) (SplMain.Children[0]);
				aButton.Click -= new System.Windows.RoutedEventHandler(onBtnClick);
				SplMain.Children.RemoveAt(0);
			}
		}
		
		private void onBtnClick(object sender, Windows.RoutedEventArgs e)
		{
			Button aButton = (Button) sender;
			//Debug.WriteLine(aButton.Content)
			if (!Globals.ThisAddIn.theControllerCollection.Contains((string) Globals.ThisAddIn.Application.ActiveSheet.name))
			{
				return;
			}
			Controller aController = Globals.ThisAddIn.theControllerCollection[Globals.ThisAddIn.Application.ActiveSheet.name];
			aController.HighLightDocAvailablilities(aButton.Name);
			
		}
		
		private void ComboBox_Loaded_1(object sender, Windows.RoutedEventArgs e)
		{
			ComboBox theComboBox;
			theComboBox = (ComboBox) sender;
			theComboBox.ItemsSource = MyGlobals.monthstrings;
			theComboBox.SelectedIndex = 0;
		}
		
		public void redraw()
		{
			//Dim theSDoc As New SDoc(CInt(Me.combo1.Text), Me.combo2.SelectedIndex + 1)
			
			//RemoveDocButtons()
			//For Each aSDoc In theSDoc.DocList
			//    Me.addDocButton(aSDoc.FirstName + " " + aSDoc.LastName, aSDoc.Initials)
			//Next
			if (Globals.ThisAddIn.theControllerCollection.Count < 1)
			{
				return;
			}
			if (Globals.ThisAddIn.theControllerCollection.Contains((string) Globals.ThisAddIn.Application.ActiveSheet.name))
			{
				theController = Globals.ThisAddIn.theControllerCollection[Globals.ThisAddIn.Application.ActiveSheet.name];
				
				if (theController != null)
				{
					this.MoisAnnee.Content = theController.aControlledMonth.Year.ToString() + "-" + MyGlobals.monthstrings[theController.aControlledMonth.Month - 1];
					Collection theDocCollection = SDoc.LoadAllDocsPerMonth(System.Convert.ToInt32(this.combo1.Text), System.Convert.ToInt32(this.combo2.SelectedIndex + 1));
					RemoveDocButtons();
					SDoc aSDoc = default(SDoc);
					foreach (SDoc tempLoopVar_aSDoc in theDocCollection)
					{
						aSDoc = tempLoopVar_aSDoc;
						this.addDocButton(aSDoc.FirstName + " " + aSDoc.LastName, aSDoc.Initials);
					}
				}
				else
				{
					this.MoisAnnee.Content = "";
				}
			}
			Initialles_Load();
			
			
		}
		
		public UserControl2()
		{
			
			// This call is required by the designer.
			InitializeComponent();
			
			// Add any initialization after the InitializeComponent() call.
			this.MoisAnnee.Content = "";
		}
		
		public void StatsBtn_Click(object sender, Windows.RoutedEventArgs e)
		{
			//lauch action in controller
			if (!Globals.ThisAddIn.theControllerCollection.Contains((string) Globals.ThisAddIn.Application.ActiveSheet.name))
			{
				return;
			}
			theController = Globals.ThisAddIn.theControllerCollection[Globals.ThisAddIn.Application.ActiveSheet.name];
			theController.statsMensuelles();
			
			
			
		}
		
		private void Button_Click_1(object sender, Windows.RoutedEventArgs e)
		{
			if (!Globals.ThisAddIn.theControllerCollection.Contains((string) Globals.ThisAddIn.Application.ActiveSheet.name))
			{
				return;
			}
			theController = Globals.ThisAddIn.theControllerCollection[Globals.ThisAddIn.Application.ActiveSheet.name];
			Excel.Range myRange = (global::Microsoft.Office.Interop.Excel.Range) Globals.ThisAddIn.Application.Selection;
			SDay aDAy = default(SDay);
			SShift aShift = default(SShift);
			SDocAvailable aDocAvail;
			if (myRange.Count == 1)
			{
				foreach (SDay tempLoopVar_aDAy in theController.aControlledMonth.Days)
				{
					aDAy = tempLoopVar_aDAy;
					foreach (SShift tempLoopVar_aShift in aDAy.Shifts)
					{
						aShift = tempLoopVar_aShift;
						if (myRange.Address == aShift.aRange.Address)
						{
							aDocAvail = (SDocAvailable) (aShift.DocAvailabilities[this.Initialles.SelectedValue]);
							aDocAvail.Availability = PublicEnums.Availability.Assigne;
							theController.fixlist(aShift);
							myRange.Value = this.Initialles.SelectedValue;
						}
					}
				}
			}
		}
		
		private void Initialles_Load()
		{
			Collection aCollection = SDoc.LoadAllDocsPerMonth(theController.aControlledMonth.Year, theController.aControlledMonth.Month);
			this.Initialles.ItemsSource = aCollection;
			this.Initialles.DisplayMemberPath = "Initials";
			this.Initialles.SelectedValuePath = "Initials";
			this.Initialles.SelectedIndex = 0;
		}
		
		private void Button_Click_2(object sender, Windows.RoutedEventArgs e)
		{
			if (theController != null)
			{
				theController.resetSheetExt();
			}
		}
		
		public void TestClick()
		{
			this.combo1.SelectedIndex = 3; //year (2017)
			this.combo2.SelectedIndex = 3; //month (april)
			Button_Click(new object(), new Windows.RoutedEventArgs());
		}
	}
	
	
	
}
