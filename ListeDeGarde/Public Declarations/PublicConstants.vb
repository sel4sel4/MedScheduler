Imports System.Windows.Forms

Public Module PublicConstants

    ' --------------------DB Address
    'Public const CONSTFILEADDRESS As String = "C:\Users\Martin\Documents\Scheduling Mira\ListesDeGarde.accdb"
    Public CONSTFILEADDRESS As String
    '----------------------Table Name Mapping

    'DB Access constants for Table ScheduleShiftType
    Public Const TABLE_shiftType = "TABLE_shiftType"
    Public Const SQLShiftStart = "ShiftStart"
    Public Const SQLShiftStop = "ShiftStop"
    Public Const SQLShiftType = "ShiftType"
    Public Const SQLActive = "Active"
    Public Const SQLDescription = "Description"

    'DB Access constants for Table ScheduleDoc
    Public Const TABLE_Doc = "TABLE_Doc"
    Public Const SQLFirstName = "FirstName"
    Public Const SQLLastName = "LastName"
    Public Const SQLInitials = "Initials"
    'Public Const SQLActive = "Active"

    'DB Access constants for Table ScheduleData
    Public Const TABLE_ScheduleData = "TABLE_ScheduleData"
    Public Const SQLDate = "aDate"
    'Public Const SQLShiftType = "ShiftType"
    'Public Const SQLInitials = "Initials"

    'DB Access constants for Table_NonDispo
    Public Const Table_NonDispo = "Table_NonDispo"
    Public Const SQLDateStart = "aDateStart"
    Public Const SQLTimeStart = "aTimeStart"
    Public Const SQLDateStop = "aDateStop"
    Public Const SQLTimeStop = "aTimeStop"
    'Public Const SQLInitials = "Initials"

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

    Public globalShiftTypes As ScheduleShiftType
    'Public theNonDispoList As Collection
    Public theList As String
    Public theRangeCollection As Collection

    ' Conenction Global variables
    'Public cnn As New ADODB.Connection  'Connection object definition
    'Public rs As New ADODB.Recordset    'recordset object definition
    'Public strSQL As String             'Query String
    Public FileAddress As String        'String to contain Excell database path and filename
    'Public GlobalDBAccessClass As DBAccessClass1
End Module

'Public Module GlobalFunctions

'    Public Sub LoadDatabaseFileLocation()
'        If My.SettingsSettings.DataBaseLocation = "" Then
'            Dim filedialog As OpenFileDialog = New OpenFileDialog()
'            filedialog.Title = "Select Location of database file"
'            filedialog.InitialDirectory = ""
'            filedialog.Filter = "*.accdb"
'            filedialog.RestoreDirectory = True
'            If filedialog.ShowDialog() = DialogResult.OK Then
'                Mysettings1.DataBaseLocation = filedialog.FileName
'            End If
'            Mysettings1.Save()
'        End If
'    End Sub


'End Module