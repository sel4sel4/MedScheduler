Imports System.Windows.Controls
Public Class ShiftInterfaceC
    Private aYearP As Integer
    Private aMonthP As Integer
    Private changesOngoing As Boolean = False
    Private myShiftTypeCollection As Collection
    Private aSShiftType As SShiftType
    Private aNewSShiftType As SShiftType

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
        Dim aSShift As SShiftType
        aSShift = CType(ShiftListView.Items(ShiftListView.SelectedIndex), SShiftType)
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
            myShiftTypeCollection = SShiftType.loadTemplateShiftTypesFromDB()
        Else
            myShiftTypeCollection = SShiftType.loadShiftTypesFromDBPerMonth(aMonthP, aYearP)
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
        Me.lundi.IsEnabled = Not locked
        Me.mardi.IsEnabled = Not locked
        Me.mercredi.IsEnabled = Not locked
        Me.jeudi.IsEnabled = Not locked
        Me.vendredi.IsEnabled = Not locked
        Me.samedi.IsEnabled = Not locked
        Me.dimache.IsEnabled = Not locked
        Me.férié.IsEnabled = Not locked
        Me.CompilerCB.IsEnabled = Not locked

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
        aSShiftType = CType(ShiftListView.SelectedItem, SShiftType)
        Me.Description.Text = aSShiftType.Description
        Me.VersionNo.Text = CStr(aSShiftType.Version)
        Me.StartHour.SelectedIndex = aSShiftType.ShiftStart \ 60
        Me.StartMin.SelectedIndex = (aSShiftType.ShiftStart Mod 60) \ 5
        Dim theStopInMinutes As Integer
        If aSShiftType.ShiftStop >= 1440 Then
            theStopInMinutes = aSShiftType.ShiftStop - 1440
        Else
            theStopInMinutes = aSShiftType.ShiftStop
        End If
        Me.StopHour.SelectedIndex = theStopInMinutes \ 60
        Me.StopMin.SelectedIndex = (aSShiftType.ShiftStop Mod 60) \ 5
        Me.ActiveCB.IsChecked = aSShiftType.Active

        Me.lundi.IsChecked = aSShiftType.Lundi
        Me.mardi.IsChecked = aSShiftType.Mardi
        Me.mercredi.IsChecked = aSShiftType.Mercredi
        Me.jeudi.IsChecked = aSShiftType.Jeudi
        Me.vendredi.IsChecked = aSShiftType.Vendredi
        Me.samedi.IsChecked = aSShiftType.Samedi
        Me.dimache.IsChecked = aSShiftType.Dimanche
        Me.férié.IsChecked = aSShiftType.Ferie

        Me.CompilerCB.IsChecked = aSShiftType.Compilation


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
        aSShiftType.Description = Me.Description.Text
        aSShiftType.ShiftStart = Me.StartHour.SelectedIndex * 60 + Me.StartMin.SelectedIndex * 5
        aSShiftType.ShiftStop = Me.StopHour.SelectedIndex * 60 + Me.StopMin.SelectedIndex * 5
        If aSShiftType.ShiftStart > aSShiftType.ShiftStop Then
            aSShiftType.ShiftStop = aSShiftType.ShiftStop + 1440
        End If
        aSShiftType.Version = CInt(Me.VersionNo.Text)
        aSShiftType.Active = CBool(Me.ActiveCB.IsChecked)
        aSShiftType.Lundi = CBool(Me.lundi.IsChecked)
        aSShiftType.Mardi = CBool(Me.mardi.IsChecked)
        aSShiftType.Mercredi = CBool(Me.mercredi.IsChecked)
        aSShiftType.Jeudi = CBool(Me.jeudi.IsChecked)
        aSShiftType.Vendredi = CBool(Me.vendredi.IsChecked)
        aSShiftType.Samedi = CBool(Me.samedi.IsChecked)
        aSShiftType.Dimanche = CBool(Me.dimache.IsChecked)
        aSShiftType.Ferie = CBool(Me.férié.IsChecked)
        aSShiftType.Compilation = CBool(Me.CompilerCB.IsChecked)
        aSShiftType.Update()
        Windows.MessageBox.Show("Le quart de travail a été mis a jour.")
        Globals.ThisAddIn.theCurrentController.resetSheetExt()


    End Sub
End Class
