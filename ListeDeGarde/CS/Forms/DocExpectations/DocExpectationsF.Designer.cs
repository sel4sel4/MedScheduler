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
	[global::Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]public partial class DocExpectationsF : System.Windows.Forms.Form
	{
		
		//Form overrides dispose to clean up the component list.
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
		
		//Required by the Windows Form Designer
		private System.ComponentModel.Container components = null;
		
		//NOTE: The following procedure is required by the Windows Form Designer
		//It can be modified using the Windows Form Designer.
		//Do not modify it using the code editor.
		[System.Diagnostics.DebuggerStepThrough()]private void InitializeComponent()
		{
			this.ElementHost1 = new System.Windows.Forms.Integration.ElementHost();
			this.DocExpectations1 = new ListeDeGarde.DocExpectations();
			this.SuspendLayout();
			//
			//ElementHost1
			//
			this.ElementHost1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.ElementHost1.Location = new System.Drawing.Point(0, 0);
			this.ElementHost1.Name = "ElementHost1";
			this.ElementHost1.Size = new System.Drawing.Size(357, 451);
			this.ElementHost1.TabIndex = 0;
			this.ElementHost1.Text = "ElementHost1";
			this.ElementHost1.Child = this.DocExpectations1;
			//
			//DocExpectationsF
			//
			this.AutoScaleDimensions = new System.Drawing.SizeF((float) (6.0F), (float) (13.0F));
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(357, 451);
			this.Controls.Add(this.ElementHost1);
			this.Name = "DocExpectationsF";
			this.Text = "DocExpectationsF";
			this.ResumeLayout(false);
			
		}
		internal System.Windows.Forms.Integration.ElementHost ElementHost1;
		internal ListeDeGarde.DocExpectations DocExpectations1;
	}
	
}
