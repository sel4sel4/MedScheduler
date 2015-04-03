Imports System.Windows.Controls
Imports System.Diagnostics


Public Class UserControl2

    Private WithEvents newBtn As Button
    Private monthstrings() As String = {"Janvier", "Février", "Mars", _
                                        "Avril", "Mai", "Juin", _
                                        "juillet", "Aout", "Septembre", _
                                        "Octobre", "Novembre", "Décembre"}

    Private Sub Button_Click(sender As Object, e As Windows.RoutedEventArgs)
        'get references to workbook
        Dim wb As Excel.Workbook = Globals.ThisAddIn.Application.ActiveWorkbook
        If MySettingsGlobal.DataBaseLocation = "" Then
            LoadDatabaseFileLocation()
        Else : CONSTFILEADDRESS = MySettingsGlobal.DataBaseLocation
        End If

        'if sheet already exists exit
        If Globals.ThisAddIn.theControllerCollection.Contains(Me.combo1.Text + "-" + Me.combo2.Text) Then Exit Sub

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
        If Not Globals.ThisAddIn.theControllerCollection.Contains(Globals.ThisAddIn.Application.ActiveSheet.name) Then Exit Sub
        Dim aController As Controller = Globals.ThisAddIn.theControllerCollection.Item(Globals.ThisAddIn.Application.ActiveSheet.name)
        aController.HighLightDocAvailablilities(aButton.Name)

    End Sub

    Private Sub ComboBox_Loaded_1(sender As Object, e As Windows.RoutedEventArgs)
        Dim theComboBox As ComboBox
        theComboBox = CType(sender, ComboBox)
        theComboBox.ItemsSource = monthstrings
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
        If Globals.ThisAddIn.theControllerCollection.Contains(Globals.ThisAddIn.Application.ActiveSheet.name) Then
            theController = Globals.ThisAddIn.theControllerCollection(Globals.ThisAddIn.Application.ActiveSheet.name)

            If Not IsNothing(theController) Then
                Me.MoisAnnee.Content = theController.aControlledMonth.Year.ToString + "-" + monthstrings(theController.aControlledMonth.Month - 1)
            Else : Me.MoisAnnee.Content = ""
            End If
        End If

    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.MoisAnnee.Content = ""
    End Sub

    Private Sub StatsBtn_Click(sender As Object, e As Windows.RoutedEventArgs) Handles StatsBtn.Click
        'lauch action in controller
        If Not Globals.ThisAddIn.theControllerCollection.Contains(Globals.ThisAddIn.Application.ActiveSheet.name) Then Exit Sub
        Dim aController As Controller = Globals.ThisAddIn.theControllerCollection.Item(Globals.ThisAddIn.Application.ActiveSheet.name)
        aController.statsMensuelles()
    End Sub
End Class


