Imports System.Windows.Controls
Public Class DrInterface




    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
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
        Dim aDOcInstance As New ScheduleDoc(aYearP, aMonthP)
        'Dim theDocInstance As ScheduleDoc
        aDocList = aDOcInstance.DocList

        Dim theContextMenu As New ContextMenu()
        Dim theMenuItem1 As New MenuItem()
        theMenuItem1.Header = "Delete"
        theContextMenu.DataContext = DocListView
        AddHandler theMenuItem1.Click, AddressOf Me.MenuItem1Clicked
        theContextMenu.Items.Add(theMenuItem1)
        Me.DocListView.ContextMenu = theContextMenu
        Me.DocListView.ItemsSource = aDocList
        Me.DocListView.SelectedIndex = 0


        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub DocListView_selectionChanged(sender As Object, e As System.Windows.RoutedEventArgs) Handles DocListView.SelectionChanged
        Dim ascheduleDoc As ScheduleDoc
        ascheduleDoc = DocListView.Items(DocListView.SelectedIndex)
        Me.initials1.Text = ascheduleDoc.Initials
        Me.firstName1.Text = ascheduleDoc.FirstName
        Me.lastName1.Text = ascheduleDoc.LastName
    End Sub

    Private Sub MenuItem1Clicked(sender As Object, e As System.Windows.RoutedEventArgs)
        Dim ascheduleDoc As ScheduleDoc
        ascheduleDoc = DocListView.Items(DocListView.SelectedIndex)
        System.Diagnostics.Debug.WriteLine(ascheduleDoc.FirstName + " " + ascheduleDoc.LastName)
    End Sub
End Class
