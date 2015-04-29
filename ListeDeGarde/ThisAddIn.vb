Imports System.Windows.Forms

Public Class ThisAddIn
    Public WithEvents xlApp As Excel.Application
    Public WithEvents xlBook As Excel.Workbook
    Public WithEvents xlSheet1 As Excel.Worksheet
    Public theControllerCollection As Collection
    Public theCurrentController As Controller
    Private myCustomTaskPane As Microsoft.Office.Tools.CustomTaskPane

    ReadOnly Property taskpane()
        Get
            Return myCustomTaskPane
        End Get
    End Property

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        'Create the task pane to create monthly Calendar
        Dim MyTaskPaneView As YearMonthPicker
        MyTaskPaneView = New YearMonthPicker
        myCustomTaskPane = Me.CustomTaskPanes.Add(MyTaskPaneView, "Liste de Garde")
        myCustomTaskPane.Visible = True
        'Load xlApp into the global variable.
        xlApp = Globals.ThisAddIn.Application

        'create a new Controller collection 
        theControllerCollection = New Collection

        'Initialize the persistent settings (stores database location) 
        MySettingsGlobal = New Settings1

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    Private Sub xlApp_Workbookopen(ByVal Wb As Excel.Workbook) Handles xlApp.WorkbookOpen
        xlBook = Globals.ThisAddIn.Application.ActiveWorkbook
        xlSheet1 = Globals.ThisAddIn.Application.ActiveSheet

    End Sub

    Private Sub xlApp_WorkbookBeforeClose(ByVal Wb As Excel.Workbook, _
  ByRef Cancel As Boolean) Handles xlApp.WorkbookBeforeClose
        Wb.Saved = True 'Set the dirty flag to true so there is no prompt to save. all Data is kept in th access DB anyway
    End Sub

    Private Sub xlApp_SheetActivate(ByVal Obb As Object) Handles xlApp.SheetActivate

        'need to rebuild the taskpane on the basis of the currentlyselected month
        'code below retreives the handle to the UserControl to trigger redraw() public function
        Dim aCollection As System.Windows.Forms.Control.ControlCollection = myCustomTaskPane.Control.Controls
        Dim aElementHost As System.Windows.Forms.Integration.ElementHost = aCollection(0)
        Dim aUserControl2 As UserControl2 = aElementHost.Child
        aUserControl2.redraw()
    End Sub

End Class
