Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1
    Private WithEvents aform1 As Form1
    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        ' UseWithEvents()
        Globals.ThisAddIn.taskpane.visible = True
    End Sub

    Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs) Handles Button2.Click
        'UseDelegate()
        aform1 = New Form1
        aform1.ShowDialog()
    End Sub

    Private Sub Button3_Click(sender As Object, e As RibbonControlEventArgs) Handles Button3.Click
        LoadDatabaseFileLocation()
    End Sub

    Private Sub Button4_Click(sender As Object, e As RibbonControlEventArgs) Handles Button4.Click
        Dim aform1 As New DrInterfaceForm
        aform1.ShowDialog()
    End Sub

    Private Sub ShiftButton_Click(sender As Object, e As RibbonControlEventArgs) Handles ShiftButton.Click
        Dim aform1 As New ShiftInterfaceF
        aform1.ShowDialog()
    End Sub

    Private Sub aForm1_close() Handles aform1.FormClosing
        If Not Globals.ThisAddIn.theControllerCollection.Contains(Globals.ThisAddIn.Application.ActiveSheet.name) Then Exit Sub
        Dim aController As Controller = Globals.ThisAddIn.theControllerCollection.Item(Globals.ThisAddIn.Application.ActiveSheet.name)
        aController.resetSheetExt()
    End Sub

 
    Private Sub ExpectDoc_Click(sender As Object, e As RibbonControlEventArgs) Handles ExpectDoc.Click
        Dim theController As Controller
        If Globals.ThisAddIn.theControllerCollection.Count < 1 Then Exit Sub
        If Not Globals.ThisAddIn.theControllerCollection.Contains(Globals.ThisAddIn.Application.ActiveSheet.name) Then Exit Sub
        theController = Globals.ThisAddIn.theControllerCollection(Globals.ThisAddIn.Application.ActiveSheet.name)

        Dim aDocExpecationF As DocExpectationsF
        aDocExpecationF = New DocExpectationsF
        aDocExpecationF.Show()
    End Sub
End Class
