
Public Class Form1
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        Me.Text = "Veuillez enter les non-disponibilitées"
        ' Add any initialization after the InitializeComponent() call.

    End Sub

    'Protected Overrides Sub Finalize()
    '    If Not Globals.ThisAddIn.theControllerCollection.Contains(Globals.ThisAddIn.Application.ActiveSheet.name) Then Exit Sub
    '    Dim aController As Controller = Globals.ThisAddIn.theControllerCollection.Item(Globals.ThisAddIn.Application.ActiveSheet.name)

    '    aController.resetSheetExt()
    '    MyBase.Finalize()
    'End Sub
End Class