Imports System.Windows.Controls
Public Class DrInterface


    Private waitingForNewSave As ScheduleDoc
    Private changesongoing As Boolean = False


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
        initializeDocList()
        Lock(True)

        'Me.DataContext = CType(Me.DocListView.SelectedItem, ScheduleDoc)
        ''initials1.Text = "{Binding Path=Initials}"
        ''firstName1.Text = "{Binding Path=FirstName}"
        ''lastName1.Text = "{Binding Path=LastName}"
        '' Add any initialization after the InitializeComponent() call.

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
        Me.Du1.SelectedDate = ascheduleDoc.EffectiveStart
        Me.Au1.SelectedDate = ascheduleDoc.EffectiveEnd
        Lock(True)
        waitingForNewSave = Nothing

    End Sub

    Private Sub MenuItem1Clicked(sender As Object, e As System.Windows.RoutedEventArgs)
        Dim ascheduleDoc As ScheduleDoc
        ascheduleDoc = DocListView.Items(DocListView.SelectedIndex)
        System.Diagnostics.Debug.WriteLine(ascheduleDoc.FirstName + " " + ascheduleDoc.LastName)
    End Sub

    Private Sub Button_Click(sender As Object, e As Windows.RoutedEventArgs)
        Dim ascheduleDoc As ScheduleDoc
        If Not IsNothing(waitingForNewSave) Then
            ascheduleDoc = waitingForNewSave
        Else
            ascheduleDoc = DocListView.SelectedItem
        End If
        'ascheduleDoc.Initials = Me.initials1.Text
        ascheduleDoc.FirstName = Me.firstName1.Text
        ascheduleDoc.LastName = Me.lastName1.Text
        ascheduleDoc.Version = Me.version1.Text
        ascheduleDoc.SoinsTog = Me.Soins.IsChecked
        ascheduleDoc.Active = Me.Active.IsChecked
        ascheduleDoc.HospitTog = Me.Hospit.IsChecked
        ascheduleDoc.NuitsTog = Me.Nuits.IsChecked
        ascheduleDoc.UrgenceTog = Me.Urgence.IsChecked
        ascheduleDoc.EffectiveStart = Me.Du1.SelectedDate
        ascheduleDoc.EffectiveEnd = Me.Au1.SelectedDate
        ascheduleDoc.save()
        changesongoing = True
        initializeDocList()
        changesongoing = False
        Me.initials1.IsReadOnly = True
    End Sub

    Private Sub Button_Click_1(sender As Object, e As Windows.RoutedEventArgs)

        Dim aDate As DateTime = DateTime.Today
        waitingForNewSave = New ScheduleDoc(aDate.Year, aDate.Month)
        Me.initials1.IsReadOnly = False
        Lock(False)
        Me.initials1.Text = waitingForNewSave.Initials
        Me.firstName1.Text = waitingForNewSave.FirstName
        Me.lastName1.Text = waitingForNewSave.LastName
        Me.version1.Text = waitingForNewSave.Version
        Me.Soins.IsChecked = waitingForNewSave.SoinsTog
        Me.Active.IsChecked = waitingForNewSave.Active
        Me.Hospit.IsChecked = waitingForNewSave.HospitTog
        Me.Nuits.IsChecked = waitingForNewSave.NuitsTog
        Me.Urgence.IsChecked = waitingForNewSave.UrgenceTog
        Me.Du1.SelectedDate = waitingForNewSave.EffectiveStart
        Me.Au1.SelectedDate = waitingForNewSave.EffectiveEnd

    End Sub

    Private Sub initializeDocList()
        Dim aYearP As Integer
        Dim aMonthP As Integer
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

        Dim aDocList As Collection
        aDocList = ScheduleDoc.LoadAllDocs2(aYearP, aMonthP)
        Me.DocListView.ItemsSource = aDocList
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
        Me.Du1.IsEnabled = Not locked
        Me.Au1.IsEnabled = Not locked
    End Sub

    Private Sub Button_Click_2(sender As Object, e As Windows.RoutedEventArgs)
        Lock(Not Me.firstName1.IsReadOnly)
    End Sub
End Class
