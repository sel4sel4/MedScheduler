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

namespace ListeDeGarde
{
	public partial class DrInterface
	{
		private SDoc waitingForNewSave;
		private List<SDoc> myDocCollection;
		private bool changesongoing = false;
		private int aYearP;
		private int aMonthP;
		
		public DrInterface()
		{
			// This call is required by the designer.
			InitializeComponent();
			ContextMenu theContextMenu = new ContextMenu();
			MenuItem theMenuItem1 = new MenuItem();
			theMenuItem1.Header = "Delete";
			theContextMenu.DataContext = DocListView;
			theMenuItem1.Click += new System.Windows.RoutedEventHandler(this.MenuItem1Clicked);
			theContextMenu.Items.Add(theMenuItem1);
			this.DocListView.ContextMenu = theContextMenu;
			GetYearMonth();
			initializeDocList();
			Lock(true);
		}
		private void GetYearMonth()
		{
			if (Globals.ThisAddIn.theControllerCollection.Exists(xy => xy.aControlledExcelSheet.Name == Globals.ThisAddIn.Application.ActiveSheet.name))
			{
				Controller aController = Globals.ThisAddIn.theControllerCollection.Find(xy => xy.aControlledExcelSheet.Name == Globals.ThisAddIn.Application.ActiveSheet.name);
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
			changesongoing = true;
			Month.SelectedIndex = aMonthP - 1;
			Year.SelectedValue = (aYearP).ToString();
			changesongoing = false;
		}
		public void DocListView_selectionChanged(object sender, System.Windows.RoutedEventArgs e)
		{
			if (changesongoing == true)
			{
				return;
			}
			SDoc aSDoc = default(SDoc);
			aSDoc = (SDoc) DocListView.SelectedItem;
			this.initials1.Text = aSDoc.Initials;
			this.initials1.IsReadOnly = true;
			this.firstName1.Text = aSDoc.FirstName;
			this.lastName1.Text = aSDoc.LastName;
			this.version1.Text = (aSDoc.Version).ToString();
			this.Soins.IsChecked = aSDoc.SoinsTog;
			this.Active.IsChecked = aSDoc.Active;
			this.Hospit.IsChecked = aSDoc.HospitTog;
			this.Nuits.IsChecked = aSDoc.NuitsTog;
			this.Urgence.IsChecked = aSDoc.UrgenceTog;
			Lock(true);
			waitingForNewSave = null;
		}
		private void MenuItem1Clicked(object sender, System.Windows.RoutedEventArgs e)
		{
			SDoc aSDoc = default(SDoc);
			aSDoc = (SDoc) DocListView.SelectedItem;
			aSDoc.Delete();
			changesongoing = true;
			initializeDocList(System.Convert.ToBoolean(Edit_Template.IsChecked));
			changesongoing = false;
		}
		private void EraseBtn_Click(object sender, system.Windows.RoutedEventArgs e) //erase doc button
		{
			SDoc aSDoc = default(SDoc);
			aSDoc = (SDoc) DocListView.SelectedItem;
			aSDoc.Delete();
			changesongoing = true;
			initializeDocList(System.Convert.ToBoolean(Edit_Template.IsChecked));
			changesongoing = false;
		}
		private void NewBtn_Click(object sender, Windows.RoutedEventArgs e) // new doc button
		{
			waitingForNewSave = new SDoc();
			this.initials1.IsReadOnly = false;
			int theVersion = default(int);
			if (Edit_Template.IsChecked)
			{
				theVersion = 0;
			}
			else
			{
				theVersion = System.Convert.ToInt32((((int.Parse(Year.Text)) - 2000) * 100) + (this.Month.SelectedIndex + 1));
			}
			Lock(false);
			this.initials1.Text = waitingForNewSave.Initials;
			this.firstName1.Text = waitingForNewSave.FirstName;
			this.lastName1.Text = waitingForNewSave.LastName;
			this.version1.Text = (theVersion).ToString();
			this.Soins.IsChecked = waitingForNewSave.SoinsTog;
			this.Active.IsChecked = waitingForNewSave.Active;
			this.Hospit.IsChecked = waitingForNewSave.HospitTog;
			this.Nuits.IsChecked = waitingForNewSave.NuitsTog;
			this.Urgence.IsChecked = waitingForNewSave.UrgenceTog;
		}
		private void ModifyBtn_Click(object sender, System.Windows.RoutedEventArgs e) //modify doc button
		{
			Lock(!this.firstName1.IsReadOnly);
		}
        private void SaveBtn_Click(object sender, System.Windows.RoutedEventArgs e) //save doc button
		{
			SDoc aSDoc = default(SDoc);
			if (waitingForNewSave != null)
			{
				aSDoc = waitingForNewSave;
			}
			else
			{
				aSDoc = (SDoc) DocListView.SelectedItem;
			}
			aSDoc.Initials = (string) this.initials1.Text;
			aSDoc.FirstName = (string) this.firstName1.Text;
			aSDoc.LastName = (string) this.lastName1.Text;
			aSDoc.Version = System.Convert.ToInt32(this.version1.Text);
			aSDoc.SoinsTog = System.Convert.ToBoolean(this.Soins.IsChecked);
			aSDoc.Active = System.Convert.ToBoolean(this.Active.IsChecked);
			aSDoc.HospitTog = System.Convert.ToBoolean(this.Hospit.IsChecked);
			aSDoc.NuitsTog = System.Convert.ToBoolean(this.Nuits.IsChecked);
			aSDoc.UrgenceTog = System.Convert.ToBoolean(this.Urgence.IsChecked);
			aSDoc.save();
			changesongoing = true;
			bool isTemplate = default(bool);
			if (this.version1.Text == (0).ToString())
			{
				isTemplate = true;
			}
			else
			{
				isTemplate = false;
			}
			initializeDocList(isTemplate);
			changesongoing = false;
			this.initials1.IsReadOnly = true;
		}
		private void initializeDocList(bool getTemplate = false)
		{
			if (getTemplate == true)
			{
				myDocCollection = SDoc.LoadTempateDocsFromDB();
			}
			else
			{
				myDocCollection = SDoc.LoadAllDocsPerMonth(aYearP, aMonthP);
			}
			changesongoing = true;
			this.DocListView.ItemsSource = myDocCollection;
			changesongoing = false;
			this.DocListView.SelectedIndex = 0;
		}
		private void Lock(bool locked)
		{
			this.firstName1.IsReadOnly = locked;
			this.lastName1.IsReadOnly = locked;
			this.version1.IsReadOnly = locked;
			this.Soins.IsEnabled = !locked;
			this.Active.IsEnabled = !locked;
			this.Hospit.IsEnabled = !locked;
			this.Nuits.IsEnabled = !locked;
			this.Urgence.IsEnabled = !locked;
		}
		
		private void aMonth_Loaded(object sender, Windows.RoutedEventArgs e)
		{
			changesongoing = true;
			ComboBox theComboBox;
			theComboBox = (ComboBox) sender;
			theComboBox.ItemsSource = MyGlobals.monthstrings;
			changesongoing = false;
		}
		private void aYear_Loaded(object sender, Windows.RoutedEventArgs e)
		{
			changesongoing = true;
			ComboBox theComboBox;
			theComboBox = (ComboBox) sender;
			theComboBox.ItemsSource = MyGlobals.yearstrings;
			changesongoing = false;
		}
		
		public void Edit_Template_Checked(object sender, Windows.RoutedEventArgs e)
		{
			changesongoing = true;
			Month.IsEnabled = false;
			Year.IsEnabled = false;
			initializeDocList(true);
			changesongoing = false;
		}
		public void Edit_Template_Unchecked(object sender, Windows.RoutedEventArgs e)
		{
			changesongoing = true;
			Month.IsEnabled = true;
			Year.IsEnabled = true;
			initializeDocList();
			changesongoing = false;
		}
		public void Year_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			if (changesongoing)
			{
				return;
			}
			aYearP = System.Convert.ToInt32(Year.SelectedValue);
			initializeDocList();
			
		}
		public void Month_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			if (changesongoing)
			{
				return;
			}
			aMonthP = Month.SelectedIndex + 1;
			initializeDocList();
			
		}
		
		
	}
	
}
