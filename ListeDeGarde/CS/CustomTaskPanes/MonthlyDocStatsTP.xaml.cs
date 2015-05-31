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
using System.Windows.Data;
using System.Windows.Media;


namespace ListeDeGarde
{
	public partial class MonthlyDocStatsTP
	{
		private Collection aCollection;
		private int[] aArray;
		
		public void loadarray(Collection theCollection, int[] theArray)
		{
			aCollection = theCollection;
			aArray = theArray;
			DrawGrid();
		}
		
		public MonthlyDocStatsTP()
		{
			// This call is required by the designer.
			InitializeComponent();
			
		}
		
		private void DrawGrid()
		{
			
			if (Globals.ThisAddIn.theControllerCollection.Count < 1)
			{
				return;
			}
			if (!Globals.ThisAddIn.theControllerCollection.Contains((Globals.ThisAddIn.Application.ActiveSheet.name).ToString()))
			{
				return;
			}
			Controller aController = Globals.ThisAddIn.theControllerCollection[Globals.ThisAddIn.Application.ActiveSheet.name];
			
			//clear everything
			MyPanel.Children.Clear();
			SDocStats theStats = default(SDocStats);
			StackPanel aHorizStackPanel = default(StackPanel);
			Label aLabel = default(Label);
			
			//create empty placeholder top left
			aLabel = (Label) (new Label());
			aLabel.Content = "";
			aLabel.Width = 50;
			aLabel.Height = 70;
			aHorizStackPanel = new StackPanel();
			aHorizStackPanel.Orientation = Orientation.Horizontal;
			aHorizStackPanel.Height = 96;
			aHorizStackPanel.Children.Add(aLabel);
			
			//create shift headers
			foreach (var aShift in aController.aControlledMonth.ShiftTypes)
			{
				if (aShift.ShiftType > 5)
				{
					break;
				}
				aLabel = (Label) (new Label());
				aLabel.Content = aShift.Description;
				aLabel.Width = 70;
				aLabel.Height = 25;
				aLabel.VerticalContentAlignment = Windows.VerticalAlignment.Center;
				RotateTransform aRotateTransform = new RotateTransform();
				aRotateTransform.Angle = 270;
				aLabel.LayoutTransform = aRotateTransform;
				aHorizStackPanel.Children.Add(aLabel);
			}
			this.MyPanel.Children.Add(aHorizStackPanel);
			aHorizStackPanel.Name = "Header";
			int theLoopCounter = 1;
			//create doc list with shifts counts
			foreach (SDocStats tempLoopVar_theStats in aCollection)
			{
				theStats = tempLoopVar_theStats;
				aHorizStackPanel = new StackPanel();
				aLabel = (Label) (new Label());
				aLabel.Content = theStats.Initials;
				aLabel.Width = 50;
				aLabel.Height = 18.5;
				aLabel.Padding = new Windows.Thickness(4);
				
				aHorizStackPanel.Height = 18.5;
				aHorizStackPanel.Orientation = Orientation.Horizontal;
				this.MyPanel.Children.Add(aHorizStackPanel);
				aHorizStackPanel.Children.Add(aLabel);
				
				foreach (var aShift in aController.aControlledMonth.ShiftTypes)
				{
					if (aShift.ShiftType > 5)
					{
						break;
					}
					aLabel = (Label) (new Label());
					if ((int) aShift.ShiftType == 1)
					{
						aLabel.Content = (theStats.shift1).ToString();
					}
					else if ((int) aShift.ShiftType == 2)
					{
						aLabel.Content = (theStats.shift2).ToString();
					}
					else if ((int) aShift.ShiftType == 3)
					{
						aLabel.Content = (theStats.shift3).ToString();
					}
					else if ((int) aShift.ShiftType == 4)
					{
						aLabel.Content = (theStats.shift4).ToString();
					}
					else if ((int) aShift.ShiftType == 5)
					{
						aLabel.Content = (theStats.shift5).ToString();
					}
					if (theStats.Initials == Globals.ThisAddIn.theCurrentController.pHighlightedDoc)
					{
						aLabel.Background = new SolidColorBrush(Color.FromRgb((byte) 150, (byte) 100, (byte) 150));
					}
					aLabel.Padding = new Windows.Thickness(3);
					aLabel.HorizontalContentAlignment = Windows.HorizontalAlignment.Center;
					aLabel.BorderBrush = System.Windows.Media.Brushes.Black;
					if (theLoopCounter == aCollection.Count)
					{
						aLabel.BorderThickness = new Windows.Thickness(1, 1, 0, 1);
					}
					else
					{
						aLabel.BorderThickness = new Windows.Thickness(1, 1, 0, 0);
					}
					
					aLabel.Width = 25;
					aLabel.Height = 18.5;
					aHorizStackPanel.Children.Add(aLabel);
				}
				if (theLoopCounter == aCollection.Count)
				{
					aLabel.BorderThickness = new Windows.Thickness(1, 1, 1, 1);
				}
				else
				{
					aLabel.BorderThickness = new Windows.Thickness(1, 1, 1, 0);
				}
				theLoopCounter++;
			}
			
			
			//clear everything
			MyPanel2.Children.Clear();
			aLabel = (Label) (new Label());
			aLabel.Content = Globals.ThisAddIn.theCurrentController.pHighlightedDoc;
			MyPanel2.Children.Add(aLabel);
			
			int y = default(int);
			if (aArray != null)
			{
				for (y = 0; y <= (aArray.Length - 1); y++)
				{
					aLabel = (Label) (new Label());
					aLabel.Content = (aArray[y]).ToString();
					MyPanel2.Children.Add(aLabel);
					
				}
			}
			
		}
	}
	
}
