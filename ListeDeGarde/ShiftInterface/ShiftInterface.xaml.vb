Imports System.Windows.Controls
Public Class UserControl3
    Private aYearP As Integer
    Private aMonthP As Integer
    Private changesOngoing As Boolean = False
    Private myShiftTypeCollection As Collection
    Private aScheduleshiftType As ScheduleShiftType
    Private aNewScheduleshiftType As ScheduleShiftType

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Dim theContextMenu As New ContextMenu()
        Dim theMenuItem1 As New MenuItem()
        theMenuItem1.Header = "inActivate"
        theContextMenu.DataContext = Me.ShiftListView
        AddHandler theMenuItem1.Click, AddressOf Me.MenuItem1Clicked
        theContextMenu.Items.Add(theMenuItem1)
        Me.ShiftListView.ContextMenu = theContextMenu
        GetYearMonth()
        initializeShiftList()
        Lock(True)
    End Sub
    Private Sub MenuItem1Clicked(sender As Object, e As System.Windows.RoutedEventArgs)
        Dim ascheduleshift As ScheduleShiftType
        ascheduleshift = ShiftListView.Items(ShiftListView.SelectedIndex)
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
    Private Sub initializeShiftList(Optional getTemplate As Boolean = False)
        If getTemplate = True Then
            myShiftTypeCollection = ScheduleShiftType.loadTemplateShiftTypesFromDB()
        Else
            myShiftTypeCollection = ScheduleShiftType.loadShiftTypesFromDBPerMonth(aMonthP, aYearP)
        End If
        changesOngoing = True
        Me.ShiftListView.ItemsSource = myShiftTypeCollection
        changesOngoing = False
        Me.ShiftListView.SelectedIndex = 0

    End Sub
    Private Sub Lock(locked As Boolean)
        Me.Description.IsReadOnly = locked
        Me.VersionNo.IsReadOnly = True
        Me.StartHour.IsEnabled = Not locked
        Me.StartMin.IsEnabled = Not locked
        Me.StopHour.IsEnabled = Not locked
        Me.StopMin.IsEnabled = Not locked
        Me.EffectiveStartDate.IsEnabled = Not locked
        Me.EffectiveStopDate.IsEnabled = Not locked
    End Sub
    Private Sub ShiftListView_selectionChanged(sender As Object, e As System.Windows.RoutedEventArgs) Handles ShiftListView.SelectionChanged
        UpdateListValues()
    End Sub
    Private Sub Hours_Loaded(sender As Object, e As Windows.RoutedEventArgs)
        Dim theComboBox As ComboBox
        theComboBox = CType(sender, ComboBox)
        theComboBox.ItemsSource = HoursStrings
        theComboBox.SelectedIndex = 0
        UpdateListValues()
    End Sub
    Private Sub Mins_Loaded(sender As Object, e As Windows.RoutedEventArgs)
        Dim theComboBox As ComboBox
        theComboBox = CType(sender, ComboBox)
        theComboBox.ItemsSource = MinutesStrings
        theComboBox.SelectedIndex = 0
        UpdateListValues()
    End Sub
    Private Sub EditBtn_Click(sender As Object, e As Windows.RoutedEventArgs) Handles EditBtn.Click
        Lock(False)
    End Sub
    Private Sub UpdateListValues()
        'If IsDBNull(ShiftListView.SelectedItem) Then Exit Sub
        If changesOngoing Then Exit Sub
        ascheduleshiftType = ShiftListView.SelectedItem
        Me.Description.Text = ascheduleshiftType.Description
        Me.VersionNo.Text = ascheduleshiftType.Version
        Me.StartHour.SelectedIndex = ascheduleshiftType.ShiftStart \ 60
        Me.StartMin.SelectedIndex = (ascheduleshiftType.ShiftStart Mod 60) \ 5
        Dim theStopInMinutes As Integer
        If ascheduleshiftType.ShiftStop >= 1440 Then
            theStopInMinutes = ascheduleshiftType.ShiftStop - 1440
        Else
            theStopInMinutes = ascheduleshiftType.ShiftStop
        End If
        Me.StopHour.SelectedIndex = theStopInMinutes \ 60
        Me.StopMin.SelectedIndex = (ascheduleshiftType.ShiftStop Mod 60) \ 5
        Me.EffectiveStartDate.SelectedDate = ascheduleshiftType.EffectiveDateStart
        Me.EffectiveStopDate.SelectedDate = aScheduleshiftType.EffectiveDateStop
        Me.ActiveCB.IsChecked = aScheduleshiftType.Active

        Lock(True)
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
    Private Sub Year_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles Year.SelectionChanged
        If changesOngoing Then Exit Sub
        aYearP = CInt(Year.SelectedValue)
        initializeShiftList()

    End Sub
    Private Sub Month_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles Month.SelectionChanged
        If changesOngoing Then Exit Sub
        aMonthP = Month.SelectedIndex + 1
        initializeShiftList()

    End Sub
    Private Sub Edit_Template_Checked(sender As Object, e As Windows.RoutedEventArgs) Handles Edit_Template.Checked
        Month.IsEnabled = False
        Year.IsEnabled = False
        initializeShiftList(True)
    End Sub
    Private Sub Edit_Template_Unchecked(sender As Object, e As Windows.RoutedEventArgs) Handles Edit_Template.Unchecked
        Month.IsEnabled = True
        Year.IsEnabled = True
        initializeShiftList()
    End Sub
    Private Sub SaveBtn_Click(sender As Object, e As Windows.RoutedEventArgs) Handles SaveBtn.Click
        aScheduleshiftType.Description = Me.Description.Text
        aScheduleshiftType.ShiftStart = Me.StartHour.SelectedIndex * 60 + Me.StartMin.SelectedIndex * 5
        aScheduleshiftType.ShiftStop = Me.StopHour.SelectedIndex * 60 + Me.StopMin.SelectedIndex * 5
        If aScheduleshiftType.ShiftStart > aScheduleshiftType.ShiftStop Then
            aScheduleshiftType.ShiftStop = aScheduleshiftType.ShiftStop + 1440
        End If
        aScheduleshiftType.Version = CInt(Me.VersionNo.Text)
        aScheduleshiftType.Active = Me.ActiveCB.IsChecked
        aScheduleshiftType.Update()
    End Sub
End Class
