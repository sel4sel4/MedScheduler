Imports System.Windows.Forms

Public Class ThisAddIn
    '--------------------------------FIELDS-------------------------------------
    Public WithEvents xlApp As Excel.Application
    Public WithEvents xlBook As Excel.Workbook
    Public WithEvents xlSheet1 As Excel.Worksheet
    Public theControllerCollection As List(Of Controller)
    Public theCurrentController As Controller
    Public myCustomTaskPane As Microsoft.Office.Tools.CustomTaskPane
    Private myThisAddinHelper As ThisAddinHelper
    '--------------------------------PROPERTIES-------------------------------------
    ReadOnly Property taskpane() As Microsoft.Office.Tools.CustomTaskPane
        Get
            Return myCustomTaskPane
        End Get
    End Property
 

    '--------------------------------EVENTS-------------------------------------

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        'Create the task pane to create monthly Calendar
        MyAddin = Me
        Dim MyTaskPaneView As YearMonthPicker
        MyTaskPaneView = New YearMonthPicker
        myCustomTaskPane = Me.CustomTaskPanes.Add(MyTaskPaneView, "Liste de Garde")
        myCustomTaskPane.Visible = True
        'Load xlApp into the global variable.
        xlApp = Globals.ThisAddIn.Application

        'create a new Controller collection 
        theControllerCollection = New List(Of Controller)

        'Initialize the persistent settings (stores database location) 
        MySettingsGlobal = New Settings1


    End Sub
    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub
    Private Sub xlApp_Workbookopen(ByVal Wb As Excel.Workbook) Handles xlApp.WorkbookOpen
        xlBook = Globals.ThisAddIn.Application.ActiveWorkbook
        xlSheet1 = CType(Globals.ThisAddIn.Application.ActiveSheet, Global.Microsoft.Office.Interop.Excel.Worksheet)

    End Sub
    Private Sub xlApp_WorkbookBeforeClose(ByVal Wb As Excel.Workbook, _
  ByRef Cancel As Boolean) Handles xlApp.WorkbookBeforeClose
        Wb.Saved = True 'Set the dirty flag to true so there is no prompt to save. all Data is kept in th access DB anyway
    End Sub
    Private Sub xlApp_SheetActivate(ByVal Obb As Object) Handles xlApp.SheetActivate

        'need to rebuild the taskpane on the basis of the currentlyselected month
        'code below retreives the handle to the UserControl to trigger redraw() public function
        Dim aCollection As System.Windows.Forms.Control.ControlCollection = myCustomTaskPane.Control.Controls
        Dim aElementHost As System.Windows.Forms.Integration.ElementHost = CType(aCollection(0), Integration.ElementHost)
        Dim theYearMonthPickerC As YearMonthPickerC = CType(aElementHost.Child, YearMonthPickerC)
        theYearMonthPickerC.redraw()
        Dim theActivatedSheet As Excel.Worksheet = CType(Obb, Excel.Worksheet)

        If Not Globals.ThisAddIn.theControllerCollection.Exists(Function(xy) theActivatedSheet.Name = xy.aControlledExcelSheet.Name) Then Exit Sub
        theCurrentController = Globals.ThisAddIn.theControllerCollection.Find(Function(xy) theActivatedSheet.Name = xy.aControlledExcelSheet.Name)

    End Sub
    Protected Overrides Function RequestComAddInAutomationService() As Object
        If myThisAddinHelper Is Nothing Then myThisAddinHelper = New ThisAddinHelper
        Return myThisAddinHelper
    End Function

    '--------------------------------PUBLIC METHODS-------------------------------------
End Class

<Runtime.InteropServices.ComVisible(True)> _
<Runtime.InteropServices.InterfaceType(Runtime.InteropServices.ComInterfaceType.InterfaceIsDual)> _
<Runtime.InteropServices.Guid("de4491a3-4ada-485a-a0cb-bb67f15d6e00")> _
Public Interface IThisAddinHelper
    Sub Launch()
    Sub testclick()
End Interface


<Runtime.InteropServices.ComVisible(True)> _
<Runtime.InteropServices.ClassInterface(Runtime.InteropServices.ClassInterfaceType.None)> _
<Runtime.InteropServices.Guid("9ED54F84-A85D-4fcd-A854-44251E925F09")> _
Public Class ThisAddinHelper
    Inherits System.Runtime.InteropServices.StandardOleMarshalObject
    Implements IThisAddinHelper


    Public Sub Launch() Implements IThisAddinHelper.Launch

        Dim wb As Excel.Workbook = Globals.ThisAddIn.Application.ActiveWorkbook
        If MySettingsGlobal.DataBaseLocation = "" Then
            LoadDatabaseFileLocation()
        Else : CONSTFILEADDRESS = MySettingsGlobal.DataBaseLocation
        End If

        'if sheet already exists exit
        'If Globals.ThisAddIn.theControllerCollection.Exists("Avril" + "-" + "2010") Then Exit Sub
        If Globals.ThisAddIn.theControllerCollection.Exists(Function(xy) xy.aControlledMonth.Month = 4 And xy.aControlledMonth.Year = 2010) Then Exit Sub

        'create a new sheet
        Globals.ThisAddIn.xlSheet1 = _
            DirectCast(wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count), _
                        Count:=1, Type:=Excel.XlSheetType.xlWorksheet), Excel.Worksheet)

        'rename the new sheet
        Globals.ThisAddIn.xlSheet1.Name = "Avril" + "-" + "2010"

        Dim theController As Controller = New Controller(Globals.ThisAddIn.xlSheet1, CInt("2010"), 4, "avril")

        Globals.ThisAddIn.theControllerCollection.Add(theController)

    End Sub

    Public Sub testclick() Implements IThisAddinHelper.testclick
        Dim aCollection As System.Windows.Forms.Control.ControlCollection = MyGlobals.MyAddin.myCustomTaskPane.Control.Controls
        Dim bElementHost As System.Windows.Forms.Integration.ElementHost = CType(aCollection(0), Windows.Forms.Integration.ElementHost)
        Dim theYearMonthPickerC As YearMonthPickerC = CType(bElementHost.Child, YearMonthPickerC)
        theYearMonthPickerC.TestClick()
    End Sub
End Class

