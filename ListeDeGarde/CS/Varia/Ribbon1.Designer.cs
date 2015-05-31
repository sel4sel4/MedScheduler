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
	public partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
	{
		
		[System.Diagnostics.DebuggerNonUserCode()]public Ribbon1(System.ComponentModel.IContainer container) : this()
		{
			
			//Required for Windows.Forms Class Composition Designer support
			if (container != null)
			{
				container.Add(this);
			}
			
		}
		
		[System.Diagnostics.DebuggerNonUserCode()]public Ribbon1() : base(Globals.Factory.GetRibbonFactory())
		{
			
			//This call is required by the Component Designer.
			InitializeComponent();
			
		}
		
		//Component overrides dispose to clean up the component list.
		[System.Diagnostics.DebuggerNonUserCode()]protected override void Dispose(bool disposing)
		{
			try
			{
				if (disposing && components != null)
				{
					components.Dispose();
				}
			}
			finally
			{
				base.Dispose(disposing);
			}
		}
		
		//Required by the Component Designer
		private System.ComponentModel.Container components = null;
		
		//NOTE: The following procedure is required by the Component Designer
		//It can be modified using the Component Designer.
		//Do not modify it using the code editor.
		[System.Diagnostics.DebuggerStepThrough()]private void InitializeComponent()
		{
			this.Tab1 = this.Factory.CreateRibbonTab;
			base.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(Ribbon1_Load);
			this.Group1 = this.Factory.CreateRibbonGroup;
			this.Button1 = this.Factory.CreateRibbonButton;
			this.Button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Button1_Click);
			this.Button2 = this.Factory.CreateRibbonButton;
			this.Button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Button2_Click);
			this.Button3 = this.Factory.CreateRibbonButton;
			this.Button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Button3_Click);
			this.Group2 = this.Factory.CreateRibbonGroup;
			this.Button4 = this.Factory.CreateRibbonButton;
			this.Button4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Button4_Click);
			this.ShiftButton = this.Factory.CreateRibbonButton;
			this.ShiftButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ShiftButton_Click);
			this.ExpectDoc = this.Factory.CreateRibbonButton;
			this.ExpectDoc.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ExpectDoc_Click);
			this.Group3 = this.Factory.CreateRibbonGroup;
			this.Tab1.SuspendLayout();
			this.Group1.SuspendLayout();
			this.Group2.SuspendLayout();
			this.Group3.SuspendLayout();
			//
			//Tab1
			//
			this.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
			this.Tab1.Groups.Add(this.Group2);
			this.Tab1.Groups.Add(this.Group3);
			this.Tab1.Groups.Add(this.Group1);
			this.Tab1.Label = "Liste de Garde";
			this.Tab1.Name = "Tab1";
			//
			//Group1
			//
			this.Group1.Items.Add(this.Button3);
			this.Group1.Label = "Base de donnés";
			this.Group1.Name = "Group1";
			//
			//Button1
			//
			this.Button1.Label = "Créer un mois";
			this.Button1.Name = "Button1";
			//
			//Button2
			//
			this.Button2.Label = "Non Disponibilités";
			this.Button2.Name = "Button2";
			//
			//Button3
			//
			this.Button3.Label = "Adresse de BD";
			this.Button3.Name = "Button3";
			//
			//Group2
			//
			this.Group2.Items.Add(this.Button4);
			this.Group2.Items.Add(this.Button2);
			this.Group2.Items.Add(this.ExpectDoc);
			this.Group2.Label = "Médecins";
			this.Group2.Name = "Group2";
			//
			//Button4
			//
			this.Button4.Label = "Medecins";
			this.Button4.Name = "Button4";
			//
			//ShiftButton
			//
			this.ShiftButton.Label = "Définition des quarts";
			this.ShiftButton.Name = "ShiftButton";
			//
			//ExpectDoc
			//
			this.ExpectDoc.Label = "Répartition Mensuelle des quarts";
			this.ExpectDoc.Name = "ExpectDoc";
			//
			//Group3
			//
			this.Group3.Items.Add(this.ShiftButton);
			this.Group3.Items.Add(this.Button1);
			this.Group3.Label = "Quarts de travail";
			this.Group3.Name = "Group3";
			//
			//Ribbon1
			//
			this.Name = "Ribbon1";
			this.RibbonType = "Microsoft.Excel.Workbook";
			this.Tabs.Add(this.Tab1);
			this.Tab1.ResumeLayout(false);
			this.Tab1.PerformLayout();
			this.Group1.ResumeLayout(false);
			this.Group1.PerformLayout();
			this.Group2.ResumeLayout(false);
			this.Group2.PerformLayout();
			this.Group3.ResumeLayout(false);
			this.Group3.PerformLayout();
			
		}
		
		internal Microsoft.Office.Tools.Ribbon.RibbonTab Tab1;
		internal Microsoft.Office.Tools.Ribbon.RibbonGroup Group1;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton Button1;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton Button2;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton Button3;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton Button4;
		internal Microsoft.Office.Tools.Ribbon.RibbonGroup Group2;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton ShiftButton;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton ExpectDoc;
		internal Microsoft.Office.Tools.Ribbon.RibbonGroup Group3;
	}
	
	public partial class ThisRibbonCollection
	{
		
		[System.Diagnostics.DebuggerNonUserCode()]internal Ribbon1 Ribbon1
		{
			get
			{
				return this.GetRibbon<Ribbon1>();
			}
		}
	}
	
}
