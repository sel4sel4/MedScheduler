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


//This class allows you to handle specific events on the settings class:
// The SettingChanging event is raised before a setting's value is changed.
// The PropertyChanged event is raised after a setting's value is changed.
// The SettingsLoaded event is raised after the setting values are loaded.
// The SettingsSaving event is raised before the setting values are saved.
namespace ListeDeGarde
{
	public sealed partial class Settings1
	{
		
		public void Settings1_SettingsLoaded(object sender, System.Configuration.SettingsLoadedEventArgs e)
		{
			PublicConstants.CONSTFILEADDRESS = Settings1.Default.DataBaseLocation;
		}
	}
	
}
