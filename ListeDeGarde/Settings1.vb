'This class allows you to handle specific events on the settings class:
' The SettingChanging event is raised before a setting's value is changed.
' The PropertyChanged event is raised after a setting's value is changed.
' The SettingsLoaded event is raised after the setting values are loaded.
' The SettingsSaving event is raised before the setting values are saved.
Partial Public NotInheritable Class Settings1

    Private Sub Settings1_SettingsLoaded(sender As Object, e As Configuration.SettingsLoadedEventArgs) Handles Me.SettingsLoaded
        CONSTFILEADDRESS = Settings1.Default.DataBaseLocation
    End Sub
End Class
