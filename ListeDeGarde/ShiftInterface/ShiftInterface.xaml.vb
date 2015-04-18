Imports System.Windows.Controls
Public Class UserControl3
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
        initializeShiftList()
        Lock(True)
    End Sub


    Private Sub MenuItem1Clicked(sender As Object, e As System.Windows.RoutedEventArgs)
        Dim ascheduleshift As ScheduleShiftType
        ascheduleshift = ShiftListView.Items(ShiftListView.SelectedIndex)
    End Sub


    Private Sub initializeShiftList()
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

        Dim myShiftType As New ScheduleShiftType(aMonthP, aYearP, True)
        Me.ShiftListView.ItemsSource = myShiftType.ShiftCollection
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
        Dim ascheduleshiftType As ScheduleShiftType
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
        Me.EffectiveStopDate.SelectedDate = ascheduleshiftType.EffectiveDateStop

        Lock(True)
    End Sub


End Class
