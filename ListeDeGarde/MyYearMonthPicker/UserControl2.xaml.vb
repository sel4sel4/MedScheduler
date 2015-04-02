Imports System.Windows.Controls
Imports System.Diagnostics


Public Class UserControl2

    Private WithEvents newBtn As Button

    Private Sub Button_Click(sender As Object, e As Windows.RoutedEventArgs)
        'get references to workbook
        Dim wb As Excel.Workbook = Globals.ThisAddIn.Application.ActiveWorkbook

        'create a new sheet
        Globals.ThisAddIn.xlSheet1 = _
            DirectCast(wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count), _
                        Count:=1, Type:=Excel.XlSheetType.xlWorksheet), Excel.Worksheet)

        'rename the new sheet
        Globals.ThisAddIn.xlSheet1.Name = Me.combo2.Text + "-" + Me.combo1.Text

        'Get list of doc names and create a button for each. Assign initals as button name
        Dim theScheduleDoc As New ScheduleDoc(CInt(Me.combo1.Text), Me.combo2.SelectedIndex + 1)
        Dim aScheduleDoc As ScheduleDoc
        RemoveDocButtons()
        For Each aScheduleDoc In theScheduleDoc.DocList
            Me.addDocButton(aScheduleDoc.FirstName + " " + aScheduleDoc.LastName, aScheduleDoc.Initials)
        Next

        Me.MoisAnnee.Content = Me.combo1.Text + "-" + Me.combo2.Text

        'create a controller instance and add it to the global collection
        Dim aController As New Controller(Globals.ThisAddIn.xlSheet1, CInt(Me.combo1.Text), Me.combo2.SelectedIndex + 1, Me.combo2.Text)
        Globals.ThisAddIn.theControllerCollection.Add(aController, Globals.ThisAddIn.xlSheet1.Name)
    End Sub

    Private Sub combo1_Loaded(sender As Object, e As Windows.RoutedEventArgs)
        Dim theList As New List(Of String)
        theList.Add("2014")
        theList.Add("2015")
        theList.Add("2016")
        theList.Add("2017")
        Dim theComboBox As ComboBox
        theComboBox = CType(sender, ComboBox)
        theComboBox.ItemsSource = theList
        theComboBox.SelectedIndex = 0
    End Sub

    Private Sub addDocButton(docName As String, initials As String)
        Dim newBtn = New Button()
        newBtn.Content = docName
        newBtn.Name = initials
        AddHandler newBtn.Click, AddressOf onBtnClick
        SplMain.Children.Add(newBtn)

    End Sub

    Private Sub RemoveDocButtons()
        Dim aButton As Button
        For x As Integer = 0 To SplMain.Children.Count - 1
            aButton = SplMain.Children.Item(0)
            RemoveHandler aButton.Click, AddressOf onBtnClick
            SplMain.Children.RemoveAt(0)
        Next
    End Sub

    Private Sub onBtnClick(ByVal sender As Object, ByVal e As Windows.RoutedEventArgs)
        Dim aButton As Button = CType(sender, Button)
        Debug.WriteLine(aButton.Content)

        Dim aController As Controller = Globals.ThisAddIn.theControllerCollection.Item(Globals.ThisAddIn.Application.ActiveSheet.name)
        aController.HighLightDocAvailablilities(aButton.Name)

    End Sub

    Private Sub ComboBox_Loaded_1(sender As Object, e As Windows.RoutedEventArgs)
        Dim theList As New List(Of String)
        theList.Add("Janvier")
        theList.Add("Fevrier")
        theList.Add("Mars")
        theList.Add("Avril")
        theList.Add("May")
        theList.Add("Juin")
        theList.Add("Juillet")
        theList.Add("Aout")
        theList.Add("Septembre")
        theList.Add("Octobre")
        theList.Add("Novembre")
        theList.Add("Decembre")
        Dim theComboBox As ComboBox
        theComboBox = CType(sender, ComboBox)
        theComboBox.ItemsSource = theList
        theComboBox.SelectedIndex = 0
    End Sub

    Public Sub redraw()
        Dim theScheduleDoc As New ScheduleDoc(CInt(Me.combo1.Text), Me.combo2.SelectedIndex + 1)
        Dim aScheduleDoc As ScheduleDoc
        RemoveDocButtons()
        For Each aScheduleDoc In theScheduleDoc.DocList
            Me.addDocButton(aScheduleDoc.FirstName + " " + aScheduleDoc.LastName, aScheduleDoc.Initials)
        Next
        If Globals.ThisAddIn.theControllerCollection.Count < 1 Then Exit Sub
        Dim theController As Controller
        Try
            theController = Globals.ThisAddIn.theControllerCollection(Globals.ThisAddIn.Application.ActiveSheet.name)
        Catch ex As Exception
        End Try
        If Not IsNothing(theController) Then
            Dim monthstrings() As String = {"Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "juillet", "Aout", "Septembre", "Octobre", "Novembre", "Decembre"}
            Me.MoisAnnee.Content = theController.aControlledMonth.Year.ToString + "-" + monthstrings(theController.aControlledMonth.Month - 1)
        Else : Me.MoisAnnee.Content = ""
        End If
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        Me.MoisAnnee.Content = ""
        ' Add any initialization after the InitializeComponent() call.

    End Sub
End Class


