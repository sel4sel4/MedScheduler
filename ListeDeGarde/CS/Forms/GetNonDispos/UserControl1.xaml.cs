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
using System.Windows.Controls;
using System.Collections.ObjectModel;

namespace ListeDeGarde
{
	public partial class UserControl1
	{
		//Private theDocList() As String
		private Collection theDocList2;
		//Private theInitialsList() As String
		private string[] TimesList = new string[] {"0:00", "1:00", "2:00", "3:00", "4:00", "5:00", "6:00", "7:00", "8:00", "9:00", "10:00", "11:00", "12:00", "13:00", "14:00", "15:00", "16:00", "17:00", "18:00", "19:00", "20:00", "21:00", "22:00", "23:00"};
		private int aMonthP = 0;
		private int aYearP = 0;
		private bool changesOngoing = false;
		private Collection theNonDispoCollection;
		
		public void AddNonDispo_Click(object sender, Windows.RoutedEventArgs e)
		{
			
			if (this.DocList.SelectedIndex == -1 || this.StartTime.SelectedIndex == -1 || this.StopTime.SelectedIndex == -1 || this.StopDate.Text == "" || this.StartDate.Text == "")
			{
				return;
			}
			
			SNonDispo aSNonDispo = new SNonDispo((this.DocList.SelectedValue).ToString(), StartDate.SelectedDate.Value, StopDate.SelectedDate.Value, System.Convert.ToInt32(this.StartTime.SelectedIndex * 60), System.Convert.ToInt32(this.StopTime.SelectedIndex * 60));
			
			updateListview();
			//If Not Globals.ThisAddIn.theControllerCollection.Contains(Globals.ThisAddIn.Application.ActiveSheet.name) Then Exit Sub
			//Dim aController As Controller = Globals.ThisAddIn.theControllerCollection.Item(Globals.ThisAddIn.Application.ActiveSheet.name)
			
			
			//aController.resetSheetExt()
			
			
			
		}
		
		public UserControl1()
		{
			
			// This call is required by the designer.
			InitializeComponent();
			
			// Add any initialization after the InitializeComponent() call.
			this.StartTime.ItemsSource = TimesList;
			this.StopTime.ItemsSource = TimesList;
			
			this.StartTime.SelectedIndex = 7;
			this.StopTime.SelectedIndex = 7;
			
			int x = 0;
			
			if (MyGlobals.MySettingsGlobal.DataBaseLocation == "")
			{
				GlobalFunctions.LoadDatabaseFileLocation();
			}
			else
			{
				PublicConstants.CONSTFILEADDRESS = MyGlobals.MySettingsGlobal.DataBaseLocation;
			}
			
			if (Globals.ThisAddIn.theControllerCollection.Contains((string) Globals.ThisAddIn.Application.ActiveSheet.name))
			{
				Controller aController = Globals.ThisAddIn.theControllerCollection[Globals.ThisAddIn.Application.ActiveSheet.name];
				aYearP = aController.aControlledMonth.Year;
				aMonthP = aController.aControlledMonth.Month;
			}
			else
			{
				DateTime aDate = new DateTime();
				aDate = DateTime.Now;
				aYearP = aDate.Year;
				aMonthP = aDate.Month;
			}
			
			changesOngoing = true;
			this.aMonth.SelectedIndex = aMonthP - 1;
			this.aYear.SelectedItem = aYearP.ToString();
			changesOngoing = false;
			
			LoadDocList();
			
		}
		
		public void DocList_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			if (changesOngoing == true)
			{
				return;
			}
			updateListview();
		}
		
		public void aMonth_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			if (changesOngoing == true)
			{
				return;
			}
			aMonthP = aMonth.SelectedIndex + 1;
			updateListview();
		}
		public void aYear_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			if (changesOngoing == true)
			{
				return;
			}
			aYearP = System.Convert.ToInt32(aYear.SelectedItem);
			updateListview();
		}
		
		public void StartDate_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			StopDate.SelectedDate = DateAndTime.DateSerial(StartDate.SelectedDate.Value.Year, StartDate.SelectedDate.Value.Month, StartDate.SelectedDate.Value.Day + 1);
		}
		
		private void updateListview()
		{
			
			NonDispoList.ItemsSource = null;
			//get nondispolist
			SNonDispo theSNonDispo = new SNonDispo();
			int x = 0;
			if (DocList.SelectedIndex != -1)
			{
				theNonDispoCollection = theSNonDispo.GetNonDispoListForDoc((DocList.SelectedValue).ToString(), aYearP, aMonthP);
				if (theNonDispoCollection != null)
				{
					
					ContextMenu theContextMenu = new ContextMenu();
					MenuItem theMenuItem1 = new MenuItem();
					theMenuItem1.Header = "Delete";
					theContextMenu.DataContext = NonDispoList;
					theMenuItem1.Click += new System.Windows.RoutedEventHandler(this.MenuItem1Clicked);
					theContextMenu.Items.Add(theMenuItem1);
					this.NonDispoList.ContextMenu = theContextMenu;
					NonDispoList.ItemsSource = theNonDispoCollection;
				}
			}
			StartDate.SelectedDate = DateAndTime.DateSerial(aYearP, aMonthP, 1);
			
		}
		
		private void MenuItem1Clicked(object sender, System.Windows.RoutedEventArgs e)
		{
			SNonDispo theNonDispo = default(SNonDispo);
			if (NonDispoList.SelectedIndex >= 0)
			{
				theNonDispo = (SNonDispo) NonDispoList.SelectedItem;
				theNonDispo.Delete();
				updateListview();
			}
			if (Globals.ThisAddIn.theControllerCollection.Contains((string) Globals.ThisAddIn.Application.ActiveSheet.name))
			{
				Controller aController = Globals.ThisAddIn.theControllerCollection[Globals.ThisAddIn.Application.ActiveSheet.name];
				aController.resetSheetExt();
			}
		}
		
		private void aMonth_Loaded(object sender, Windows.RoutedEventArgs e)
		{
			changesOngoing = true;
			ComboBox theComboBox;
			theComboBox = (ComboBox) sender;
			theComboBox.ItemsSource = MyGlobals.monthstrings;
			changesOngoing = false;
		}
		
		private void aYear_Loaded(object sender, Windows.RoutedEventArgs e)
		{
			changesOngoing = true;
			ComboBox theComboBox;
			theComboBox = (ComboBox) sender;
			theComboBox.ItemsSource = MyGlobals.yearstrings;
			changesOngoing = false;
		}
		
		private void LoadDocList()
		{
			Collection theSDocCollection = new Collection();
			theSDocCollection = SDoc.LoadAllDocsPerMonth(aYearP, aMonthP);
			changesOngoing = true;
			if (theSDocCollection.Count > 0)
			{
				this.DocList.ItemsSource = theSDocCollection;
				this.DocList.DisplayMemberPath = "FistAndLastName";
				this.DocList.SelectedValuePath = "Initials";
				this.DocList.SelectedIndex = 0;
				updateListview();
			}
			changesOngoing = false;
		}
		
		
	}
	
}
