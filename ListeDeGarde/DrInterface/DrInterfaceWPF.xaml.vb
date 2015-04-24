Imports System.Windows.Controls
Public Class DrInterface
    Private waitingForNewSave As ScheduleDoc
    Private myDocCollection As Collection
    Private changesongoing As Boolean = False
    Private aYearP As Integer
    Private aMonthP As Integer

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        Dim theContextMenu As New ContextMenu()
        Dim theMenuItem1 As New MenuItem()
        theMenuItem1.Header = "Delete"
        theContextMenu.DataContext = DocListView
        AddHandler theMenuItem1.Click, AddressOf Me.MenuItem1Clicked
        theContextMenu.Items.Add(theMenuItem1)
        Me.DocListView.ContextMenu = theContextMenu
        GetYearMonth()
        initializeDocList()
        Lock(True)
    End Sub
    Private Sub GetYearMonth()
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
        Month.SelectedIndex = aMonthP - 1
        Year.SelectedValue = CStr(aYearP)
        changesOngoing = False
    End Sub
    Private Sub DocListView_selectionChanged(sender As Object, e As System.Windows.RoutedEventArgs) Handles DocListView.SelectionChanged
        If changesongoing = True Then Exit Sub
        Dim ascheduleDoc As ScheduleDoc
        ascheduleDoc = DocListView.SelectedItem
        Me.initials1.Text = ascheduleDoc.Initials
        Me.initials1.IsReadOnly = True
        Me.firstName1.Text = ascheduleDoc.FirstName
        Me.lastName1.Text = ascheduleDoc.LastName
        Me.version1.Text = ascheduleDoc.Version
        Me.Soins.IsChecked = ascheduleDoc.SoinsTog
        Me.Active.IsChecked = ascheduleDoc.Active
        Me.Hospit.IsChecked = ascheduleDoc.HospitTog
        Me.Nuits.IsChecked = ascheduleDoc.NuitsTog
        Me.Urgence.IsChecked = ascheduleDoc.UrgenceTog
        Lock(True)
        waitingForNewSave = Nothing
    End Sub
    Private Sub MenuItem1Clicked(sender As Object, e As System.Windows.RoutedEventArgs)
        Dim ascheduleDoc As ScheduleDoc
        ascheduleDoc = DocListView.Items(DocListView.SelectedIndex)
        System.Diagnostics.Debug.WriteLine(ascheduleDoc.FirstName + " " + ascheduleDoc.LastName)
    End Sub
    Private Sub EraseBtn_Click(sender As Object, e As Windows.RoutedEventArgs) 'erase doc button
        Dim ascheduleDoc As ScheduleDoc
        ascheduleDoc = DocListView.SelectedItem
        ascheduleDoc.Delete()
        changesongoing = True
        initializeDocList(Edit_Template.IsChecked)
        changesongoing = False
    End Sub
    Private Sub NewBtn_Click(sender As Object, e As Windows.RoutedEventArgs) ' new doc button
        waitingForNewSave = New ScheduleDoc()
        Me.initials1.IsReadOnly = False
        Dim theVersion As Integer
        If Edit_Template.IsChecked Then
            theVersion = 0
        Else
            theVersion = ((CInt(Year.Text) - 2000) * 100) + (Me.Month.SelectedIndex + 1)
        End If
        Lock(False)
        Me.initials1.Text = waitingForNewSave.Initials
        Me.firstName1.Text = waitingForNewSave.FirstName
        Me.lastName1.Text = waitingForNewSave.LastName
        Me.version1.Text = theVersion
        Me.Soins.IsChecked = waitingForNewSave.SoinsTog
        Me.Active.IsChecked = waitingForNewSave.Active
        Me.Hospit.IsChecked = waitingForNewSave.HospitTog
        Me.Nuits.IsChecked = waitingForNewSave.NuitsTog
        Me.Urgence.IsChecked = waitingForNewSave.UrgenceTog
    End Sub
    Private Sub ModifyBtn_Click(sender As Object, e As Windows.RoutedEventArgs) 'modify doc button
        Lock(Not Me.firstName1.IsReadOnly)
    End Sub
    Private Sub SaveBtn_Click(sender As Object, e As Windows.RoutedEventArgs) 'save doc button
        Dim ascheduleDoc As ScheduleDoc
        If Not IsNothing(waitingForNewSave) Then
            ascheduleDoc = waitingForNewSave
        Else
            ascheduleDoc = DocListView.SelectedItem
        End If
        ascheduleDoc.Initials = Me.initials1.Text
        ascheduleDoc.FirstName = Me.firstName1.Text
        ascheduleDoc.LastName = Me.lastName1.Text
        ascheduleDoc.Version = Me.version1.Text
        ascheduleDoc.SoinsTog = Me.Soins.IsChecked
        ascheduleDoc.Active = Me.Active.IsChecked
        ascheduleDoc.HospitTog = Me.Hospit.IsChecked
        ascheduleDoc.NuitsTog = Me.Nuits.IsChecked
        ascheduleDoc.UrgenceTog = Me.Urgence.IsChecked
        ascheduleDoc.save()
        changesongoing = True
        initializeDocList()
        changesongoing = False
        Me.initials1.IsReadOnly = True
    End Sub
    Private Sub initializeDocList(Optional getTemplate As Boolean = False)
        If getTemplate = True Then
            myDocCollection = ScheduleDoc.LoadTempateDocsFromDB()
        Else
            myDocCollection = ScheduleDoc.LoadAllDocsPerMonth(aYearP, aMonthP)
        End If
        changesongoing = True
        Me.DocListView.ItemsSource = myDocCollection
        changesongoing = False
        Me.DocListView.SelectedIndex = 0
    End Sub
    Private Sub Lock(locked As Boolean)
        Me.firstName1.IsReadOnly = locked
        Me.lastName1.IsReadOnly = locked
        Me.version1.IsReadOnly = locked
        Me.Soins.IsEnabled = Not locked
        Me.Active.IsEnabled = Not locked
        Me.Hospit.IsEnabled = Not locked
        Me.Nuits.IsEnabled = Not locked
        Me.Urgence.IsEnabled = Not locked
    End Sub

    Private Sub aMonth_Loaded(sender As Object, e As Windows.RoutedEventArgs)
        changesongoing = True
        Dim theComboBox As ComboBox
        theComboBox = CType(sender, ComboBox)
        theComboBox.ItemsSource = monthstrings
        changesongoing = False
    End Sub
    Private Sub aYear_Loaded(sender As Object, e As Windows.RoutedEventArgs)
        changesongoing = True
        Dim theComboBox As ComboBox
        theComboBox = CType(sender, ComboBox)
        theComboBox.ItemsSource = yearstrings
        changesongoing = False
    End Sub

    Private Sub Edit_Template_Checked(sender As Object, e As Windows.RoutedEventArgs) Handles Edit_Template.Checked
        changesongoing = True
        Month.IsEnabled = False
        Year.IsEnabled = False
        initializeDocList(True)
        changesongoing = False
    End Sub
    Private Sub Edit_Template_Unchecked(sender As Object, e As Windows.RoutedEventArgs) Handles Edit_Template.Unchecked
        changesongoing = True
        Month.IsEnabled = True
        Year.IsEnabled = True
        initializeDocList()
        changesongoing = False
    End Sub
    Private Sub Year_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles Year.SelectionChanged
        If changesongoing Then Exit Sub
        aYearP = CInt(Year.SelectedValue)
        initializeDocList()

    End Sub
    Private Sub Month_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles Month.SelectionChanged
        If changesongoing Then Exit Sub
        aMonthP = Month.SelectedIndex + 1
        initializeDocList()

    End Sub


End Class
