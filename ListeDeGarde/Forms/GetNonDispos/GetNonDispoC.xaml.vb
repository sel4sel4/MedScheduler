Imports System.Diagnostics
Imports System.Windows.Controls
Imports System.Collections.ObjectModel
Public Class GetNonDispoC
    'Private theDocList() As String
    Private theDocList2 As Collection
    'Private theInitialsList() As String
    Private TimesList() As String = {"0:00", "1:00", "2:00", _
                                    "3:00", "4:00", "5:00", _
                                    "6:00", "7:00", "8:00", _
                                    "9:00", "10:00", "11:00", _
                                    "12:00", "13:00", "14:00", _
                                    "15:00", "16:00", "17:00", _
                                    "18:00", "19:00", "20:00", _
                                    "21:00", "22:00", "23:00"}
    Private aMonthP As Integer = 0
    Private aYearP As Integer = 0
    Private changesOngoing As Boolean = False
    Private theNonDispoCollection As Collection

    Private Sub AddNonDispo_Click(sender As Object, e As Windows.RoutedEventArgs) Handles AddNonDispo.Click

        If Me.DocList.SelectedIndex = -1 Or _
            Me.StartTime.SelectedIndex = -1 Or _
            Me.StopTime.SelectedIndex = -1 Or _
            Me.StopDate.Text = "" Or _
            Me.StartDate.Text = "" Then Exit Sub

        Dim aSNonDispo As New SNonDispo(CStr(Me.DocList.SelectedValue), _
                                                       StartDate.SelectedDate.Value, _
                                                      StopDate.SelectedDate.Value, _
                                                        Me.StartTime.SelectedIndex * 60, _
                                                        Me.StopTime.SelectedIndex * 60)

        updateListview()
        'If Not Globals.ThisAddIn.theControllerCollection.Contains(Globals.ThisAddIn.Application.ActiveSheet.name) Then Exit Sub
        'Dim aController As Controller = Globals.ThisAddIn.theControllerCollection.Item(Globals.ThisAddIn.Application.ActiveSheet.name)


        'aController.resetSheetExt()



    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.StartTime.ItemsSource = TimesList
        Me.StopTime.ItemsSource = TimesList

        Me.StartTime.SelectedIndex = 7
        Me.StopTime.SelectedIndex = 7

        Dim x As Integer = 0

        If MySettingsGlobal.DataBaseLocation = "" Then
            LoadDatabaseFileLocation()
        Else : CONSTFILEADDRESS = MySettingsGlobal.DataBaseLocation
        End If

        If Globals.ThisAddIn.theControllerCollection.Contains(Globals.ThisAddIn.Application.ActiveSheet.name) Then
            Dim aController As Controller = Globals.ThisAddIn.theControllerCollection.Item(Globals.ThisAddIn.Application.ActiveSheet.name)
            aYearP = aController.aControlledMonth.Year
            aMonthP = aController.aControlledMonth.Month
        Else
            Dim aDate As New DateTime
            aDate = DateTime.Now
            aYearP = aDate.Year
            aMonthP = aDate.Month
        End If

        changesOngoing = True
        Me.aMonth.SelectedIndex = aMonthP - 1
        Me.aYear.SelectedItem = aYearP.ToString()
        changesOngoing = False

        LoadDocList()

    End Sub

    Private Sub DocList_SelectionChanged(sender As Object, e As Windows.Controls.SelectionChangedEventArgs) Handles DocList.SelectionChanged
        If changesOngoing = True Then Exit Sub
        updateListview()
    End Sub

    Private Sub aMonth_SelectionChanged(sender As Object, e As Windows.Controls.SelectionChangedEventArgs) Handles aMonth.SelectionChanged
        If changesOngoing = True Then Exit Sub
        aMonthP = aMonth.SelectedIndex + 1
        updateListview()
    End Sub
    Private Sub aYear_SelectionChanged(sender As Object, e As Windows.Controls.SelectionChangedEventArgs) Handles aYear.SelectionChanged
        If changesOngoing = True Then Exit Sub
        aYearP = CInt(aYear.SelectedItem)
        updateListview()
    End Sub

    Private Sub StartDate_SelectionChanged(sender As Object, e As Windows.Controls.SelectionChangedEventArgs) Handles StartDate.SelectedDateChanged
        StopDate.SelectedDate = DateSerial(StartDate.SelectedDate.Value.Year, StartDate.SelectedDate.Value.Month, StartDate.SelectedDate.Value.Day + 1)
    End Sub

    Private Sub updateListview()

        NonDispoList.ItemsSource = Nothing
        'get nondispolist
        Dim theSNonDispo As New SNonDispo
        Dim x As Integer = 0
        If DocList.SelectedIndex <> -1 Then
            theNonDispoCollection = theSNonDispo.GetNonDispoListForDoc(CStr(DocList.SelectedValue), aYearP, aMonthP)
            If Not IsNothing(theNonDispoCollection) Then

                Dim theContextMenu As New ContextMenu()
                Dim theMenuItem1 As New MenuItem()
                theMenuItem1.Header = "Delete"
                theContextMenu.DataContext = NonDispoList
                AddHandler theMenuItem1.Click, AddressOf Me.MenuItem1Clicked
                theContextMenu.Items.Add(theMenuItem1)
                Me.NonDispoList.ContextMenu = theContextMenu
                NonDispoList.ItemsSource = theNonDispoCollection
            End If
        End If
        StartDate.SelectedDate = DateSerial(aYearP, aMonthP, 1)

    End Sub

    Private Sub MenuItem1Clicked(sender As Object, e As System.Windows.RoutedEventArgs)
        Dim theNonDispo As SNonDispo
        If NonDispoList.SelectedIndex >= 0 Then
            theNonDispo = CType(NonDispoList.SelectedItem, SNonDispo)
            theNonDispo.Delete()
            updateListview()
        End If
        'If Globals.ThisAddIn.theControllerCollection.Contains(Globals.ThisAddIn.Application.ActiveSheet.name) Then
        '    Dim aController As Controller = Globals.ThisAddIn.theControllerCollection.Item(Globals.ThisAddIn.Application.ActiveSheet.name)
        '    aController.resetSheetExt()
        'End If
    End Sub

    Private Sub aMonth_Loaded(sender As Object, e As Windows.RoutedEventArgs)
        changesOngoing = True
        Dim theComboBox As ComboBox
        theComboBox = CType(sender, ComboBox)
        theComboBox.ItemsSource = monthstrings
        changesOngoing = False
    End Sub

    Private Sub aYear_Loaded(sender As Object, e As Windows.RoutedEventArgs)
        changesOngoing = True
        Dim theComboBox As ComboBox
        theComboBox = CType(sender, ComboBox)
        theComboBox.ItemsSource = yearstrings
        changesOngoing = False
    End Sub

    Private Sub LoadDocList()
        Dim theSDocCollection As New Collection
        theSDocCollection = SDoc.LoadAllDocsPerMonth(aYearP, aMonthP)
        changesOngoing = True
        If theSDocCollection.Count > 0 Then
            Me.DocList.ItemsSource = theSDocCollection
            Me.DocList.DisplayMemberPath = "FistAndLastName"
            Me.DocList.SelectedValuePath = "Initials"
            Me.DocList.SelectedIndex = 0
            updateListview()
        End If
        changesOngoing = False
    End Sub


End Class
