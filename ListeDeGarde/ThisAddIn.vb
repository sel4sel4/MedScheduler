Public Class ThisAddIn
    Public WithEvents xlApp As Excel.Application
    Public WithEvents xlBook As Excel.Workbook
    Public WithEvents xlSheet1 As Excel.Worksheet
    Private myCustomTaskPane As Microsoft.Office.Tools.CustomTaskPane


    ReadOnly Property taskpane()
        Get
            Return myCustomTaskPane
        End Get
    End Property

    Private Sub ThisAddIn_Startup() Handles Me.Startup

        Dim MyTaskPaneView As YearMonthPicker
        MyTaskPaneView = New YearMonthPicker
        myCustomTaskPane = Me.CustomTaskPanes.Add(MyTaskPaneView, "Liste de Garde")
        myCustomTaskPane.Visible = True
        xlApp = Globals.ThisAddIn.Application

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    Private Sub xlApp_Workbookopen(ByVal Wb As Excel.Workbook) Handles xlApp.WorkbookOpen
        xlBook = Globals.ThisAddIn.Application.ActiveWorkbook
        xlSheet1 = Globals.ThisAddIn.Application.ActiveSheet
        Dim theRange As Excel.Range
        theRange = xlSheet1.Range("a1")
        theRange.Value2 = "I have just opened"
        System.Diagnostics.Debug.WriteLine("ThisAddin: opening the workbook")
        xlBook.Saved = True 'Set the dirty flag to true so there is no prompt to save
    End Sub

    Private Sub xlApp_WorkbookBeforeClose(ByVal Wb As Excel.Workbook, _
  ByRef Cancel As Boolean) Handles xlApp.WorkbookBeforeClose
        System.Diagnostics.Debug.WriteLine("WithEvents: Closing the workbook.")
        Wb.Saved = True 'Set the dirty flag to true so there is no prompt to save
    End Sub

    Private Sub xlApp_SheetActivate(ByVal Obb As Object) Handles xlApp.SheetActivate
        System.Diagnostics.Debug.WriteLine("WithEvents: switching activeSheet.")
        Me.xlSheet1 = CType(Obb, Excel.Worksheet)


        Dim mycontrol As System.Windows.Forms.UserControl = Globals.ThisAddIn.myCustomTaskPane.Control
        Dim aCollection As System.Windows.Forms.Control.ControlCollection = mycontrol.Controls
        Dim aEH As System.Windows.Forms.Integration.ElementHost = aCollection(0)
        Dim aUC2 As UserControl2 = aEH.Child
        aUC2.xlSheet2 = CType(Obb, Excel.Worksheet)
    End Sub






End Class
