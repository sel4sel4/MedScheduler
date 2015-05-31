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
	public partial class UserControl4
	{
		
		private Collection aCollection;
		
		public void loadarray(Collection theCollection)
		{
			aCollection = theCollection;
			DrawGrid();
		}
		
		public UserControl4()
		{
			// This call is required by the designer.
			InitializeComponent();
			
		}
		
		private void DrawGrid()
		{
			
			SDocStats theStats = default(SDocStats);
			
			StackPanel aHorizStackPanel = default(StackPanel);
			Label aLabel = default(Label);
			// Add any initialization after the InitializeComponent() call.
			MyPanel.Children.Clear();
			
			aLabel = (Label) (new Label());
			aLabel.Content = "";
			aLabel.Width = 50;
			aLabel.Height = 70;
			RotateTransform aRotateTransform = new RotateTransform();
			aRotateTransform.Angle = 270;
			
			aHorizStackPanel = new StackPanel();
			aHorizStackPanel.Orientation = Orientation.Horizontal;
			aHorizStackPanel.Height = 100;
			aHorizStackPanel.Children.Add(aLabel);
			
			if (Globals.ThisAddIn.theControllerCollection.Count < 1)
			{
				return;
			}
			if (!Globals.ThisAddIn.theControllerCollection.Contains((string) Globals.ThisAddIn.Application.ActiveSheet.name))
			{
				return;
			}
			Controller aController = Globals.ThisAddIn.theControllerCollection[Globals.ThisAddIn.Application.ActiveSheet.name];
			
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
				aLabel.LayoutTransform = aRotateTransform;
				aHorizStackPanel.Children.Add(aLabel);
			}
			
			this.MyPanel.Children.Add(aHorizStackPanel);
			aHorizStackPanel.Name = "Header";
			
			foreach (SDocStats tempLoopVar_theStats in aCollection)
			{
				theStats = tempLoopVar_theStats;
				aHorizStackPanel = new StackPanel();
				aLabel = (Label) (new Label());
				aLabel.Content = theStats.Initials;
				aLabel.Width = 50;
				aHorizStackPanel.Height = 21;
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
					aLabel.Width = 25;
					aHorizStackPanel.Children.Add(aLabel);
				}
			}
		}
		
	}
	
}
