


Module UseDelegateSM

    '==================================================================
    'Demonstrates Using a Delegate for Event Handling
    '==================================================================

    Private xlApp As Excel.Application
    Private xlBook As Excel.Workbook
    Private xlSheet1 As Excel.Worksheet
    Private EventDel_BeforeBookClose As Excel.AppEvents_WorkbookBeforeCloseEventHandler
    Private EventDel_CellsChange As Excel.DocEvents_ChangeEventHandler

    Public Sub UseDelegate()




        'Start Excel and create a new workbook.
        xlApp = Globals.ThisAddIn.Application
        xlBook = Globals.ThisAddIn.Application.ActiveWorkbook
        xlBook.Windows(1).Caption = "Uses UseDelegate"

        'Get references to the three worksheets.
        xlSheet1 = Globals.ThisAddIn.Application.ActiveSheet

        'Add an event handler for the WorkbookBeforeClose Event of the
        'Application object.
        EventDel_BeforeBookClose = New Excel.AppEvents_WorkbookBeforeCloseEventHandler( _
              AddressOf BeforeBookClose)
        AddHandler xlApp.WorkbookBeforeClose, EventDel_BeforeBookClose

        'Add an event handler for the Change event of both Worksheet 
        'objects.
        EventDel_CellsChange = New Excel.DocEvents_ChangeEventHandler( _
              AddressOf CellsChange)
        AddHandler xlSheet1.Change, EventDel_CellsChange

        'Make Excel visible and give the user control.
        'xlApp.Visible = True
        'xlApp.UserControl = True
    End Sub

    Private Sub CellsChange(ByVal Target As Excel.Range)
        'This is called when a cell or cells on a worksheet are changed.
        System.Diagnostics.Debug.WriteLine("Delegate: You Changed Cells " + Target.Address + " on " + _
                          Target.Worksheet.Name())
    End Sub

    Private Sub BeforeBookClose(ByVal Wb As Excel.Workbook, ByRef Cancel As Boolean)
        'This is called when you choose to close the workbook in Excel.
        'The event handlers are removed, and then the workbook is closed 
        'without saving changes.
        System.Diagnostics.Debug.WriteLine("Delegate: Closing the workbook and removing event handlers.")
        RemoveHandler xlSheet1.Change, EventDel_CellsChange
        RemoveHandler xlApp.WorkbookBeforeClose, EventDel_BeforeBookClose
        Wb.Saved = True 'Set the dirty flag to true so there is no prompt to save.
    End Sub


    '    Dim xlApp As Excel.Application
    '2.         Dim xlBook As Excel.Workbook
    '3.         Dim xlSheet As Excel.Worksheet
    '4.         Dim xlButton As Excel.OLEObject
    '5.         Dim iStartLine As Long
    '6.         xlApp = New Excel.Application
    '7.         xlApp.Visible = True
    '8.         xlBook = xlApp.Workbooks.Add
    '9.         xlSheet = xlBook.ActiveSheet
    '10.         xlButton = xlSheet.OLEObjects.Add(ClassType:="Forms.CommandButton.1", _
    '11.             Link:=False, DisplayAsIcon:=False, Left:=30, Top:=20, Width:=72, Height:=24)
    '12.         xlButton.Name = "BtnTest"
    '13.         xlButton.Object.Caption = "Press"
    '14.         With xlBook.VBProject.VBComponents.Item(xlSheet.CodeName).CodeModule
    '15.             iStartLine = .CreateEventProc("Click", "BtnTest") + 1
    '16.             .InsertLines(iStartLine, "Msgbox ""Hi""")
    '17.         End With
    '18.         xlApp.VBE.MainWindow.Visible = False



End Module
