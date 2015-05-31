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
using System.Windows.Media;



namespace ListeDeGarde
{
	public partial class DocExpectations
	{
		
		private Collection theDocCollection;
		public DocExpectations()
		{
			
			// This call is required by the designer.
			InitializeComponent();
			DrawGrid();
			
			
			
		}
		private void DrawGrid()
		{
			
			StackPanel aHorizStackPanel = default(StackPanel);
			Label aLabel = default(Label);
			TextBox aTextBox = default(TextBox);
			SShiftType aShift = default(SShiftType);
			// Add any initialization after the InitializeComponent() call.
			if (!Globals.ThisAddIn.theControllerCollection.Contains((string) Globals.ThisAddIn.Application.ActiveSheet.name))
			{
				return;
			}
			Controller aController = Globals.ThisAddIn.theControllerCollection[Globals.ThisAddIn.Application.ActiveSheet.name];
			SDoc aSDoc = default(SDoc);
			
			MyPanel.Children.Clear();
			
			aLabel = (Label) (new Label());
			aLabel.Content = "";
			aLabel.Width = 120;
			aLabel.Height = 70;
			RotateTransform aRotateTransform = new RotateTransform();
			aRotateTransform.Angle = 270;
			
			aHorizStackPanel = new StackPanel();
			aHorizStackPanel.Orientation = Orientation.Horizontal;
			aHorizStackPanel.Height = 100;
			aHorizStackPanel.Children.Add(aLabel);
			
			
			foreach (SShiftType tempLoopVar_aShift in aController.aControlledMonth.ShiftTypes)
			{
				aShift = tempLoopVar_aShift;
				if (aShift.ShiftType > 5)
				{
					break;
				}
				aLabel = (Label) (new Label());
				aLabel.Content = aShift.Description;
				aLabel.Width = 70;
				aLabel.Height = 30;
				aLabel.LayoutTransform = aRotateTransform;
				aHorizStackPanel.Children.Add(aLabel);
			}
			aLabel = (Label) (new Label());
			aLabel.Content = "Total";
			aLabel.Width = 70;
			aLabel.Height = 30;
			aLabel.LayoutTransform = aRotateTransform;
			aHorizStackPanel.Children.Add(aLabel);
			
			this.MyPanel.Children.Add(aHorizStackPanel);
			aHorizStackPanel.Name = "Header";
			
			if (!(MyPanel.FindName(aHorizStackPanel.Name) == null))
			{
				MyPanel.UnregisterName(aHorizStackPanel.Name);
			}
			MyPanel.RegisterName(aHorizStackPanel.Name, aHorizStackPanel);
			
			
			if (this.Edit_Template.IsChecked == false)
			{
				theDocCollection = aController.aControlledMonth.DocList;
			}
			else
			{
				theDocCollection = SDoc.LoadTempateDocsFromDB();
			}
			
			foreach (SDoc tempLoopVar_aSDoc in theDocCollection)
			{
				aSDoc = tempLoopVar_aSDoc;
				aHorizStackPanel = new StackPanel();
				aLabel = (Label) (new Label());
				aLabel.Content = aSDoc.FistAndLastName;
				aLabel.Width = 120;
				aHorizStackPanel.Height = 21;
				aHorizStackPanel.Orientation = Orientation.Horizontal;
				this.MyPanel.Children.Add(aHorizStackPanel);
				aHorizStackPanel.Name = aSDoc.Initials;
				if (!(MyPanel.FindName(aHorizStackPanel.Name) == null))
				{
					MyPanel.UnregisterName(aHorizStackPanel.Name);
				}
				MyPanel.RegisterName(aHorizStackPanel.Name, aHorizStackPanel);
				aHorizStackPanel.Children.Add(aLabel);
				
				foreach (SShiftType tempLoopVar_aShift in aController.aControlledMonth.ShiftTypes)
				{
					aShift = tempLoopVar_aShift;
					if (aShift.ShiftType > 5)
					{
						break;
					}
					aTextBox = new TextBox();
					switch (aShift.ShiftType)
					{
						case 1:
							aTextBox.Text = (aSDoc.Shift1).ToString();
							break;
						case 2:
							aTextBox.Text = (aSDoc.Shift2).ToString();
							break;
						case 3:
							aTextBox.Text = (aSDoc.Shift3).ToString();
							break;
						case 4:
							aTextBox.Text = (aSDoc.Shift4).ToString();
							break;
						case 5:
							aTextBox.Text = (aSDoc.Shift5).ToString();
							break;
					}
					aTextBox.Width = 30;
					aTextBox.Name = (string) (aSDoc.Initials + "_" + (aShift.ShiftType).ToString());
					aTextBox.TextChanged += new System.Windows.Controls.TextChangedEventHandler(TextHasChanged);
					aHorizStackPanel.Children.Add(aTextBox);
					if (!(MyPanel.FindName(aTextBox.Name) == null))
					{
						MyPanel.UnregisterName(aTextBox.Name);
					}
					MyPanel.RegisterName(aTextBox.Name, aTextBox);
					
				}
				aTextBox = new TextBox();
				aTextBox.Text = "0";
				aTextBox.Width = 30;
				aTextBox.IsEnabled = false;
				aTextBox.Name = "Total_" + aSDoc.Initials;
				aHorizStackPanel.Children.Add(aTextBox);
				if (!(MyPanel.FindName(aTextBox.Name) == null))
				{
					MyPanel.UnregisterName(aTextBox.Name);
				}
				MyPanel.RegisterName(aTextBox.Name, aTextBox);
			}
			
			aHorizStackPanel = new StackPanel();
			aLabel = (Label) (new Label());
			aLabel.Content = "Total:";
			aLabel.Width = 120;
			aHorizStackPanel.Height = 21;
			aHorizStackPanel.Orientation = Orientation.Horizontal;
			this.MyPanel.Children.Add(aHorizStackPanel);
			aHorizStackPanel.Name = "Total";
			if (!(MyPanel.FindName(aHorizStackPanel.Name) == null))
			{
				MyPanel.UnregisterName(aHorizStackPanel.Name);
			}
			MyPanel.RegisterName(aHorizStackPanel.Name, aHorizStackPanel);
			aHorizStackPanel.Children.Add(aLabel);
			
			for (var x = 1; x <= 5; x++)
			{
				aTextBox = new TextBox();
				aTextBox.Text = "0";
				aTextBox.Width = 30;
				aTextBox.IsEnabled = false;
				aTextBox.Name = (string) ("Total_" + (x).ToString());
				aHorizStackPanel.Children.Add(aTextBox);
				if (!(MyPanel.FindName(aTextBox.Name) == null))
				{
					MyPanel.UnregisterName(aTextBox.Name);
				}
				MyPanel.RegisterName(aTextBox.Name, aTextBox);
				
			}
			
			aHorizStackPanel = new StackPanel();
			aLabel = (Label) (new Label());
			aLabel.Content = "Expected:";
			aLabel.Width = 120;
			aHorizStackPanel.Height = 21;
			aHorizStackPanel.Orientation = Orientation.Horizontal;
			this.MyPanel.Children.Add(aHorizStackPanel);
			aHorizStackPanel.Children.Add(aLabel);
			int[] theArray = null;
			theArray = CountExpectedShiftsPerMonth();
			for (var x = 0; x <= 4; x++)
			{
				aLabel = (Label) (new Label());
				aLabel.Content = theArray[(int) x];
				aLabel.Width = 30;
				aHorizStackPanel.Children.Add(aLabel);
			}
			CalculateTotals();
			
		}
		
		
		private void TextHasChanged(object sender, Windows.RoutedEventArgs e)
		{
			TextBox myTextBox = default(TextBox);
			myTextBox = (TextBox) sender;
			if (myTextBox.Name.Substring(0, 5) == "Total")
			{
				return;
			}
			if (!Information.IsNumeric(myTextBox.Text))
			{
				myTextBox.Text = "0";
				return;
			}
			
			CalculateTotals();
			//Dim mySplit As String() = myTextBox.Name.Split(New Char() {"_"c})
			//Dim x As Integer
			//Dim aObject As Object
			//Dim aTextBox As TextBox
			//Dim aTotal As Integer = 0
			//For x = 1 To 5
			//    aObject = MyPanel.FindName(mySplit(0) + "_" + x.ToString())
			//    aTextBox = CType(aObject, TextBox)
			//    aTotal = aTotal + CInt(aTextBox.Text)
			//Next
			//aObject = MyPanel.FindName("Total_" + mySplit(0))
			//aTextBox = CType(aObject, TextBox)
			//aTextBox.Text = CStr(aTotal)
			
			//Dim aSDoc As SDoc
			//aTotal = 0
			//For Each aSDoc In theDocCollection
			//    aObject = MyPanel.FindName(aSDoc.Initials + "_" + mySplit(1))
			//    aTextBox = CType(aObject, TextBox)
			//    aTotal = aTotal + CInt(aTextBox.Text)
			//Next
			//aObject = MyPanel.FindName("Total_" + mySplit(1))
			//aTextBox = CType(aObject, TextBox)
			//aTextBox.Text = CStr(aTotal)
		}
		
		public void Edit_Template_Checked(object sender, Windows.RoutedEventArgs e)
		{
			
			DrawGrid();
		}
		public void Edit_Template_Unchecked(object sender, Windows.RoutedEventArgs e)
		{
			DrawGrid();
		}
		
		public void SaveBtn_Click(object sender, Windows.RoutedEventArgs e)
		{
			//cycle through all doctors, load the shift numbers from the grid
			//apply them to each doc and save them either to the template or to the specific month.
			
			SDoc aSDoc = default(SDoc);
			
			foreach (SDoc tempLoopVar_aSDoc in theDocCollection)
			{
				aSDoc = tempLoopVar_aSDoc;
				aSDoc.Shift1 = System.Convert.ToInt32(MyPanel.FindName(aSDoc.Initials + "_1").text);
				aSDoc.Shift2 = System.Convert.ToInt32(MyPanel.FindName(aSDoc.Initials + "_2").text);
				aSDoc.Shift3 = System.Convert.ToInt32(MyPanel.FindName(aSDoc.Initials + "_3").text);
				aSDoc.Shift4 = System.Convert.ToInt32(MyPanel.FindName(aSDoc.Initials + "_4").text);
				aSDoc.Shift5 = System.Convert.ToInt32(MyPanel.FindName(aSDoc.Initials + "_5").text);
				aSDoc.save();
			}
			Globals.ThisAddIn.theCurrentController.resetSheetExt();
			
		}
		
		private void CalculateTotals()
		{
			
			int x = default(int);
			object aObject = default(object);
			TextBox aTextBox = default(TextBox);
			SDoc aSDoc = default(SDoc);
			int horizTotal = 0;
			int vert1Total = 0;
			int vert2Total = 0;
			int vert3Total = 0;
			int vert4Total = 0;
			int vert5Total = 0;
			
			
			foreach (SDoc tempLoopVar_aSDoc in theDocCollection)
			{
				aSDoc = tempLoopVar_aSDoc;
				for (x = 1; x <= 5; x++)
				{
					aObject = MyPanel.FindName(aSDoc.Initials + "_" + x.ToString());
					aTextBox = (TextBox) aObject;
					horizTotal = horizTotal + int.Parse(aTextBox.Text);
					switch (x)
					{
						case 1:
							vert1Total = vert1Total + int.Parse(aTextBox.Text);
							break;
						case 2:
							vert2Total = vert2Total + int.Parse(aTextBox.Text);
							break;
						case 3:
							vert3Total = vert3Total + int.Parse(aTextBox.Text);
							break;
						case 4:
							vert4Total = vert4Total + int.Parse(aTextBox.Text);
							break;
						case 5:
							vert5Total = vert5Total + int.Parse(aTextBox.Text);
							break;
							
					}
					
					
				}
				aObject = MyPanel.FindName("Total_" + aSDoc.Initials);
				aTextBox = (TextBox) aObject;
				aTextBox.Text = (horizTotal).ToString();
				horizTotal = 0;
			}
			MyPanel.FindName("Total_1").text = (vert1Total).ToString();
			MyPanel.FindName("Total_2").text = (vert2Total).ToString();
			MyPanel.FindName("Total_3").text = (vert3Total).ToString();
			MyPanel.FindName("Total_4").text = (vert4Total).ToString();
			MyPanel.FindName("Total_5").text = (vert5Total).ToString();
		}
		
		private int[] CountExpectedShiftsPerMonth()
		{
			int[] theArray = null;
			theArray = new int[5];
			
			SMonth theControlledMonth = default(SMonth);
			theControlledMonth = Globals.ThisAddIn.theCurrentController.aControlledMonth;
			SDay aDay = default(SDay);
			SShift ashift = default(SShift);
			foreach (SDay tempLoopVar_aDay in theControlledMonth.Days)
			{
				aDay = tempLoopVar_aDay;
				foreach (SShift tempLoopVar_ashift in aDay.Shifts)
				{
					ashift = tempLoopVar_ashift;
					if (ashift.ShiftType <= 5)
					{
						theArray[ashift.ShiftType - 1] = theArray[ashift.ShiftType - 1] + 1;
					}
				}
			}
			return theArray;
		}
	}
	
}
