Imports System.Windows.Forms

Public Module PublicConstants

    ' --------------------DB Address
    'Public const CONSTFILEADDRESS As String = "C:\Users\Martin\Documents\Scheduling Mira\ListesDeGarde.accdb"
    Public CONSTFILEADDRESS As String
    '----------------------Table Name Mapping

    'DB Access constants for Table SShiftType
    Public Const TABLE_shiftType As String = "TABLE_shiftType"
    Public Const SQLShiftStart As String = "ShiftStart"
    Public Const SQLShiftStop As String = "ShiftStop"
    Public Const SQLShiftType As String = "ShiftType"
    Public Const SQLActive As String = "Active"
    Public Const SQLDescription As String = "Description"
    Public Const SQLLundi As String = "Lundi"
    Public Const SQLMardi As String = "Mardi"
    Public Const SQLMercredi As String = "Mercredi"
    Public Const SQLJeudi As String = "Jeudi"
    Public Const SQLVendredi As String = "Vendredi"
    Public Const SQLSamedi As String = "Samedi"
    Public Const SQLDimanche As String = "Dimanche"
    Public Const SQLFerie As String = "Ferie"
    Public Const SQLCompilation As String = "Compilation"
    Public Const SQLOrder As String = "The_Order"



    'Public Const SQLEffectiveStart = "EffectiveStart"
    'Public Const SQLEffectiveEnd = "EffectiveEnd"

    'DB Access constants for Table SDoc
    Public Const TABLE_Doc As String = "TABLE_Doc"
    Public Const SQLFirstName As String = "FirstName"
    Public Const SQLLastName As String = "LastName"
    Public Const SQLInitials As String = "Initials"
    'Public Const SQLActive As String = "Active"
    Public Const SQLEffectiveStart As String = "EffectiveStart"
    Public Const SQLEffectiveEnd As String = "EffectiveEnd"
    Public Const SQLShift1 As String = "Shift1"
    Public Const SQLShift2 As String = "Shift2"
    Public Const SQLShift3 As String = "Shift3"
    Public Const SQLShift4 As String = "Shift4"
    Public Const SQLShift5 As String = "Shift5"
    Public Const SQLUrgenceTog As String = "Urgence"
    Public Const SQLHospitTog As String = "Hospit"
    Public Const SQLSoinsTog As String = "Soins"
    Public Const SQLNuitsTog As String = "Nuits"
    Public Const SQLVersion As String = "Version"

    'DB Access constants for Table ScheduleData
    Public Const TABLE_ScheduleData As String = "TABLE_ScheduleData"
    Public Const SQLDate As String = "aDate"
    'Public Const SQLShiftType As String = "ShiftType"
    'Public Const SQLInitials As String = "Initials"

    'DB Access constants for Table_NonDispo
    Public Const Table_NonDispo As String = "Table_NonDispo"
    Public Const SQLDateStart As String = "aDateStart"
    Public Const SQLTimeStart As String = "aTimeStart"
    Public Const SQLDateStop As String = "aDateStop"
    Public Const SQLTimeStop As String = "aTimeStop"
    'Public Const SQLInitials As String = "Initials"

    '-----------------------Default Values
    Public Const DEFAULTDATE As Long = 29221
    Public Const kTicksToDays As Long = 864000000000
End Module

Public Module PublicStructures
    '------------------------------------------------------------------------------
    '                   =========== Structures ===========
    '------------------------------------------------------------------------------

    Public Structure T_DBRefTypeS
        Dim theSQLName As String
        Dim theValue As String
    End Structure

    Public Structure T_DBRefTypeA
        Dim theSQLName As String
        Dim theValue As Availability
    End Structure

    Public Structure T_DBRefTypeL
        Dim theSQLName As String
        Dim theValue As Long
    End Structure

    Public Structure T_DBRefTypeI
        Dim theSQLName As String
        Dim theValue As Integer
    End Structure

    Public Structure T_DBRefTypeB
        Dim theSQLName As String
        Dim theValue As Boolean
    End Structure

    Public Structure T_DBRefTypeD
        Dim theSQLName As String
        Dim theValue As Date
    End Structure
End Module


Public Module PublicEnums
    Public Enum Availability
        Assigne = 0
        Dispo = 1
        NonDispoPermanente = 2
        NonDispoTemporaire = 3
        SurUtilise = 4
        AssigneSpecial = 5
    End Enum

    Public Enum Weekdays
        monday = 1
        tuesday = 2
        wednesday = 3
        Thursay = 4
        Friday = 5
        Saturday = 6
        Sunday = 7
    End Enum

    Public Enum EnumWhereSubClause
        EW_None
        EW_begin
        EW_end
    End Enum

End Module

Public Module MyGlobals

    Public globalShiftTypes As SShiftType
    'Public theNonDispoList As Collection
    Public theList As String
    Public theRangeCollection As Collection
    Public MyAddin As ThisAddIn

    ' Conenction Global variables
    Public cnn As New ADODB.Connection  'Connection object definition
    'Public rs As New ADODB.Recordset    'recordset object definition
    'Public strSQL As String             'Query String
    Public FileAddress As String        'String to contain Excell database path and filename
    'Public GlobalDBAccessClass As DBAccessClass1
    Public MySettingsGlobal As Settings1
    Public daystrings() As String = {"dimanche", "lundi", "mardi", _
                                   "mercredi", "jeudi", "vendredi", _
                                   "samedi"}

    Public monthstrings() As String = {"janvier", "février", "mars", _
                                    "avril", "mai", "juin", _
                                    "juillet", "aout", "septembre", _
                                    "octobre", "novembre", "décembre"}
    Public yearstrings() As String = {"2014", "2015", "2016", "2017"}

    Public HoursStrings() As String = {"00", "01", "02", _
                                    "03", "04", "05", _
                                    "06", "07", "08", _
                                    "09", "10", "11", _
                                    "12", "13", "14", _
                                    "15", "16", "17", _
                                    "18", "19", "20", _
                                    "21", "22", "23"}

    Public MinutesStrings() As String = {"00", "05", "10", _
                                "15", "20", "25", _
                                "30", "35", "40", _
                                "45", "50", "55"}


End Module

Public Module GlobalFunctions

    Public Sub LoadDatabaseFileLocation()
        Dim filedialog As OpenFileDialog = New OpenFileDialog()
        filedialog.Title = "Select Location of database file"
        filedialog.InitialDirectory = ""
        filedialog.Filter = "Access DB files (*.accdb)|*.accdb"

        filedialog.RestoreDirectory = True
        If filedialog.ShowDialog() = DialogResult.OK Then
            MySettingsGlobal.DataBaseLocation = filedialog.FileName
        End If
        MySettingsGlobal.Save()

    End Sub

    Public Function cAccessDateStr(theDate As Date) As String

        Return "#" + theDate.ToString("yyyy-M-d") + "#"

    End Function

End Module