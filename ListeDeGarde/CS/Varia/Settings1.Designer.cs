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


//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.34014
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------




namespace ListeDeGarde
{
	[global::System.Runtime.CompilerServices.CompilerGeneratedAttribute(), global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "12.0.0.0")]public sealed partial class Settings1 : global::System.Configuration.ApplicationSettingsBase
	{
		
		private static Settings1 defaultInstance = (Settings1) (global::System.Configuration.ApplicationSettingsBase.Synchronized(new Settings1()));
		
		public static Settings1 Default
		{
			get
			{
				return defaultInstance;
			}
		}
		
		[global::System.Configuration.UserScopedSettingAttribute(), global::System.Diagnostics.DebuggerNonUserCodeAttribute(), global::System.Configuration.DefaultSettingValueAttribute("F:\\Users\\Martin\\Documents\\Scheduling Mira\\ListesDeGarde.accdb"), global::System.Configuration.SettingsManageabilityAttribute(global::System.Configuration.SettingsManageability.Roaming)]public string DataBaseLocation
		{
			get
			{
				return System.Convert.ToString(this["DataBaseLocation"]);
			}
			set
			{
				this["DataBaseLocation"] = value;
			}
		}
	}
	
}
