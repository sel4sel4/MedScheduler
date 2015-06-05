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
	public partial class ShiftInterfaceC
	{
		private int aYearP;
		private int aMonthP;
		private bool changesOngoing = false;
		private List<SShiftType> myShiftTypeCollection;
		private SShiftType aSShiftType;
		private SShiftType aNewSShiftType;
		
		public ShiftInterfaceC()
		{
			
			// This call is required by the designer.
			InitializeComponent();
			
			// Add any initialization after the InitializeComponent() call.
			ContextMenu theContextMenu = new ContextMenu();
			MenuItem theMenuItem1 = new MenuItem();
			theMenuItem1.Header = "inActivate";
			theContextMenu.DataContext = this.ShiftListView;
			theMenuItem1.Click += new System.Windows.RoutedEventHandler(this.MenuItem1Clicked);
			theContextMenu.Items.Add(theMenuItem1);
			this.ShiftListView.ContextMenu = theContextMenu;
			GetYearMonth();
			initializeShiftList();
			Lock(true);
		}
		private void MenuItem1Clicked(object sender, System.Windows.RoutedEventArgs e)
		{
			SShiftType aSShift;
			aSShift = (SShiftType) (ShiftListView.Items[ShiftListView.SelectedIndex]);
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
			changesOngoing = true;
			Month.SelectedIndex = aMonthP - 1;
			Year.SelectedValue = (aYearP).ToString();
			changesOngoing = false;
		}
		private void initializeShiftList(bool getTemplate = false)
		{
			if (getTemplate == true)
			{
				myShiftTypeCollection = SShiftType.loadTemplateShiftTypesFromDB();
			}
			else
			{
				myShiftTypeCollection = SShiftType.loadShiftTypesFromDBPerMonth(aMonthP, aYearP);
			}
			changesOngoing = true;
			this.ShiftListView.ItemsSource = myShiftTypeCollection;
			changesOngoing = false;
			this.ShiftListView.SelectedIndex = 0;
			
		}
		private void Lock(bool locked)
		{
			this.Description.IsReadOnly = locked;
			this.VersionNo.IsReadOnly = true;
			this.StartHour.IsEnabled = !locked;
			this.StartMin.IsEnabled = !locked;
			this.StopHour.IsEnabled = !locked;
			this.StopMin.IsEnabled = !locked;
			this.lundi.IsEnabled = !locked;
			this.mardi.IsEnabled = !locked;
			this.mercredi.IsEnabled = !locked;
			this.jeudi.IsEnabled = !locked;
			this.vendredi.IsEnabled = !locked;
			this.samedi.IsEnabled = !locked;
			this.dimache.IsEnabled = !locked;
			this.férié.IsEnabled = !locked;
			this.CompilerCB.IsEnabled = !locked;
			
		}
		public void ShiftListView_selectionChanged(object sender, System.Windows.RoutedEventArgs e)
		{
			UpdateListValues();
		}
        private void Hours_Loaded(object sender, System.Windows.RoutedEventArgs e)
		{
			ComboBox theComboBox;
			theComboBox = (ComboBox) sender;
			theComboBox.ItemsSource = MyGlobals.HoursStrings;
			theComboBox.SelectedIndex = 0;
			UpdateListValues();
		}
        private void Mins_Loaded(object sender, System.Windows.RoutedEventArgs e)
		{
			ComboBox theComboBox;
			theComboBox = (ComboBox) sender;
			theComboBox.ItemsSource = MyGlobals.MinutesStrings;
			theComboBox.SelectedIndex = 0;
			UpdateListValues();
		}
        public void EditBtn_Click(object sender, System.Windows.RoutedEventArgs e)
		{
			Lock(false);
		}
		private void UpdateListValues()
		{
			//If IsDBNull(ShiftListView.SelectedItem) Then Exit Sub
			if (changesOngoing)
			{
				return;
			}
			aSShiftType = (SShiftType) ShiftListView.SelectedItem;
			this.Description.Text = aSShiftType.Description;
			this.VersionNo.Text = (aSShiftType.Version).ToString();
			this.StartHour.SelectedIndex = aSShiftType.ShiftStart / 60;
			this.StartMin.SelectedIndex = (aSShiftType.ShiftStart % 60) / 5;
			int theStopInMinutes = default(int);
			if (aSShiftType.ShiftStop >= 1440)
			{
				theStopInMinutes = aSShiftType.ShiftStop - 1440;
			}
			else
			{
				theStopInMinutes = aSShiftType.ShiftStop;
			}
			this.StopHour.SelectedIndex = theStopInMinutes / 60;
			this.StopMin.SelectedIndex = (aSShiftType.ShiftStop % 60) / 5;
			this.ActiveCB.IsChecked = aSShiftType.Active;
			
			this.lundi.IsChecked = aSShiftType.Lundi;
			this.mardi.IsChecked = aSShiftType.Mardi;
			this.mercredi.IsChecked = aSShiftType.Mercredi;
			this.jeudi.IsChecked = aSShiftType.Jeudi;
			this.vendredi.IsChecked = aSShiftType.Vendredi;
			this.samedi.IsChecked = aSShiftType.Samedi;
			this.dimache.IsChecked = aSShiftType.Dimanche;
			this.férié.IsChecked = aSShiftType.Ferie;
			
			this.CompilerCB.IsChecked = aSShiftType.Compilation;
			
			
			Lock(true);
		}
        private void aMonth_Loaded(object sender, System.Windows.RoutedEventArgs e)
		{
			changesOngoing = true;
			ComboBox theComboBox;
			theComboBox = (ComboBox) sender;
			theComboBox.ItemsSource = MyGlobals.monthstrings;
			changesOngoing = false;
		}
        private void aYear_Loaded(object sender, System.Windows.RoutedEventArgs e)
		{
			changesOngoing = true;
			ComboBox theComboBox;
			theComboBox = (ComboBox) sender;
			theComboBox.ItemsSource = MyGlobals.yearstrings;
			changesOngoing = false;
		}
		public void Year_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			if (changesOngoing)
			{
				return;
			}
			aYearP = System.Convert.ToInt32(Year.SelectedValue);
			initializeShiftList();
			
		}
		public void Month_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			if (changesOngoing)
			{
				return;
			}
			aMonthP = Month.SelectedIndex + 1;
			initializeShiftList();
			
		}
        public void Edit_Template_Checked(object sender, System.Windows.RoutedEventArgs e)
		{
			Month.IsEnabled = false;
			Year.IsEnabled = false;
			initializeShiftList(true);
		}
        public void Edit_Template_Unchecked(object sender, System.Windows.RoutedEventArgs e)
		{
			Month.IsEnabled = true;
			Year.IsEnabled = true;
			initializeShiftList();
		}
        public void SaveBtn_Click(object sender, System.Windows.RoutedEventArgs e)
		{
			aSShiftType.Description = (string) this.Description.Text;
			aSShiftType.ShiftStart = System.Convert.ToInt32(this.StartHour.SelectedIndex * 60 + this.StartMin.SelectedIndex * 5);
			aSShiftType.ShiftStop = System.Convert.ToInt32(this.StopHour.SelectedIndex * 60 + this.StopMin.SelectedIndex * 5);
			if (aSShiftType.ShiftStart > aSShiftType.ShiftStop)
			{
				aSShiftType.ShiftStop = aSShiftType.ShiftStop + 1440;
			}
			aSShiftType.Version = System.Convert.ToInt32(this.VersionNo.Text);
			aSShiftType.Active = System.Convert.ToBoolean(this.ActiveCB.IsChecked);
			aSShiftType.Lundi = System.Convert.ToBoolean(this.lundi.IsChecked);
			aSShiftType.Mardi = System.Convert.ToBoolean(this.mardi.IsChecked);
			aSShiftType.Mercredi = System.Convert.ToBoolean(this.mercredi.IsChecked);
			aSShiftType.Jeudi = System.Convert.ToBoolean(this.jeudi.IsChecked);
			aSShiftType.Vendredi = System.Convert.ToBoolean(this.vendredi.IsChecked);
			aSShiftType.Samedi = System.Convert.ToBoolean(this.samedi.IsChecked);
			aSShiftType.Dimanche = System.Convert.ToBoolean(this.dimache.IsChecked);
			aSShiftType.Ferie = System.Convert.ToBoolean(this.férié.IsChecked);
			aSShiftType.Compilation = System.Convert.ToBoolean(this.CompilerCB.IsChecked);
			aSShiftType.Update();
			Windows.MessageBox.Show("Le quart de travail a été mis a jour.");
			Globals.ThisAddIn.theCurrentController.resetSheetExt();
			
			
		}
	}
	
}
