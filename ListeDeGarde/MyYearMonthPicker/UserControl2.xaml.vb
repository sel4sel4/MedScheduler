Imports System.Windows.Controls
Imports System.Diagnostics


Public Class UserControl2

    Private WithEvents newBtn As Button
    Private theController As Controller

    Private Sub Button_Click(sender As Object, e As Windows.RoutedEventArgs)
        'get references to workbook
        Dim wb As Excel.Workbook = Globals.ThisAddIn.Application.ActiveWorkbook
        If MySettingsGlobal.DataBaseLocation = "" Then
            LoadDatabaseFileLocation()
        Else : CONSTFILEADDRESS = MySettingsGlobal.DataBaseLocation
        End If

        'if sheet already exists exit
        If Globals.ThisAddIn.theControllerCollection.Contains(Me.combo2.Text + "-" + Me.combo1.Text) Then Exit Sub

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
        theController = New Controller(Globals.ThisAddIn.xlSheet1, CInt(Me.combo1.Text), Me.combo2.SelectedIndex + 1, Me.combo2.Text)

        Globals.ThisAddIn.theControllerCollection.Add(theController, Globals.ThisAddIn.xlSheet1.Name)
        Initialles_Load()

    End Sub

    Private Sub combo1_Loaded(sender As Object, e As Windows.RoutedEventArgs)
        Dim theComboBox As ComboBox
        theComboBox = CType(sender, ComboBox)
        theComboBox.ItemsSource = yearstrings
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
        'Debug.WriteLine(aButton.Content)
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
        If Globals.ThisAddIn.theControllerCollection.Contains(Globals.ThisAddIn.Application.ActiveSheet.name) Then
            theController = Globals.ThisAddIn.theControllerCollection(Globals.ThisAddIn.Application.ActiveSheet.name)

            If Not IsNothing(theController) Then
                Me.MoisAnnee.Content = theController.aControlledMonth.Year.ToString + "-" + monthstrings(theController.aControlledMonth.Month - 1)
            Else : Me.MoisAnnee.Content = ""
            End If
        End If
        Initialles_Load()


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
        theController = Globals.ThisAddIn.theControllerCollection.Item(Globals.ThisAddIn.Application.ActiveSheet.name)
        theController.statsMensuelles()
       

    End Sub

    Private Sub Button_Click_1(sender As Object, e As Windows.RoutedEventArgs)
        If Not Globals.ThisAddIn.theControllerCollection.Contains(Globals.ThisAddIn.Application.ActiveSheet.name) Then Exit Sub
        theController = Globals.ThisAddIn.theControllerCollection.Item(Globals.ThisAddIn.Application.ActiveSheet.name)
        Dim myRange As Excel.Range = Globals.ThisAddIn.Application.Selection
        Dim aDAy As ScheduleDay
        Dim aShift As ScheduleShift
        Dim aDocAvail As scheduleDocAvailable
        If myRange.Count = 1 Then
            For Each aDAy In theController.aControlledMonth.Days
                For Each aShift In aDAy.Shifts
                    If myRange.Address = aShift.aRange.Address Then
                        aDocAvail = aShift.DocAvailabilities.Item(Me.Initialles.SelectedValue)
                        aDocAvail.Availability = Availability.Assigne
                        theController.fixlist(aShift)
                        myRange.Value = Me.Initialles.SelectedValue
                    End If
                Next
            Next
        End If
    End Sub

    Private Sub Initialles_Load()
        Dim aCollection As Collection = ScheduleDoc.LoadAllDocsPerMonth(theController.aControlledMonth.Year, theController.aControlledMonth.Month)
        Me.Initialles.ItemsSource = aCollection
        Me.Initialles.DisplayMemberPath = "Initials"
        Me.Initialles.SelectedValuePath = "Initials"
        Me.Initialles.SelectedIndex = 0
    End Sub

    Private Sub Button_Click_2(sender As Object, e As Windows.RoutedEventArgs)
        If Not theController Is Nothing Then
            theController.resetSheetExt()
        End If
    End Sub
End Class


