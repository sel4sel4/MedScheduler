Imports System.Diagnostics
Imports System.Windows.Controls
Public Class UserControl1
    Private theDocList() As String
    Private theInitialsList() As String
    Private TimesList() As String = {"0:00", "1:00", "2:00", _
                                    "3:00", "4:00", "5:00", _
                                    "6:00", "7:00", "8:00", _
                                    "9:00", "10:00", "11:00", _
                                    "12:00", "13:00", "14:00", _
                                    "15:00", "16:00", "17:00", _
                                    "18:00", "19:00", "20:00", _
                                    "21:00", "22:00", "23:00"}
    Private theNonDispoCollection As Collection

    Private Sub AddNonDispo_Click(sender As Object, e As Windows.RoutedEventArgs) Handles AddNonDispo.Click
        
        If Me.DocList.SelectedIndex = -1 Or _
            Me.StartTime.SelectedIndex = -1 Or _
            Me.StopTime.SelectedIndex = -1 Or _
            Me.StopDate.Text = "" Or _
            Me.StartDate.Text = "" Then Exit Sub

        Dim aScheduleNonDispo As New ScheduleNonDispo(theInitialsList(Me.DocList.SelectedIndex), _
                                                        Me.StartDate.SelectedDate, _
                                                        Me.StopDate.SelectedDate, _
                                                        Me.StartTime.SelectedIndex * 60, _
                                                        Me.StopTime.SelectedIndex * 60)

        If Not Globals.ThisAddIn.theControllerCollection.Contains(Globals.ThisAddIn.Application.ActiveSheet.name) Then Exit Sub
        Dim aController As Controller = Globals.ThisAddIn.theControllerCollection.Item(Globals.ThisAddIn.Application.ActiveSheet.name)
        updateListview()
        aController.resetSheet()


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
        Dim aYear As Integer, aMonth As Integer

        If Globals.ThisAddIn.theControllerCollection.Contains(Globals.ThisAddIn.Application.ActiveSheet.name) Then
            Dim aController As Controller = Globals.ThisAddIn.theControllerCollection.Item(Globals.ThisAddIn.Application.ActiveSheet.name)
            aYear = aController.aControlledMonth.Year
            aMonth = aController.aControlledMonth.Month
        Else
            Dim aDate As New DateTime
            aDate = DateTime.Now
            aYear = aDate.Year
            aMonth = aDate.Month
        End If
        Dim theScheduleDoc As New ScheduleDoc(aYear, aMonth)
        Dim aScheduleDoc As ScheduleDoc
        ReDim theDocList(0 To theScheduleDoc.DocList.Count - 1)
        ReDim theInitialsList(0 To theScheduleDoc.DocList.Count - 1)
        For Each aScheduleDoc In theScheduleDoc.DocList
            theDocList(x) = aScheduleDoc.FirstName + " " + aScheduleDoc.LastName
            theInitialsList(x) = aScheduleDoc.Initials
            x = x + 1
        Next
        Me.DocList.ItemsSource = theDocList


    End Sub

    Private Sub DocList_SelectionChanged(sender As Object, e As Windows.Controls.SelectionChangedEventArgs) Handles DocList.SelectionChanged
        updateListview()
    End Sub

    Private Sub NonDispoList_SelectionChanged(sender As Object, e As Windows.Controls.SelectionChangedEventArgs) Handles NonDispoList.SelectionChanged
        'Debug.WriteLine("the selected index is" + NonDispoList.SelectedIndex.ToString())
    End Sub


    Private Sub updateListview()

        'get year and month
        Dim aYear As Integer, aMonth As Integer
        If Globals.ThisAddIn.theControllerCollection.Contains(Globals.ThisAddIn.Application.ActiveSheet.name) Then
            Dim aController As Controller = Globals.ThisAddIn.theControllerCollection.Item(Globals.ThisAddIn.Application.ActiveSheet.name)
            aYear = aController.aControlledMonth.Year
            aMonth = aController.aControlledMonth.Month
        Else
            Dim aDate As New DateTime
            aDate = DateTime.Now
            aYear = aDate.Year
            aMonth = aDate.Month
        End If

        Dim theListMenuItem As ListViewItem


        'get nondispolist
        Dim theSchedulenondispo As New ScheduleNonDispo
        Dim aSchedulenondispo As ScheduleNonDispo
        Dim x As Integer = 0
        theNonDispoCollection = theSchedulenondispo.GetNonDispoListForDoc(theInitialsList(DocList.SelectedIndex), aYear, aMonth)
        NonDispoList.Items.Clear()
        If Not IsNothing(theNonDispoCollection) Then

            Dim theContextMenu As New ContextMenu()
            Dim theMenuItem1 As New MenuItem()
            theMenuItem1.Header = "Delete"
            Dim theMenuItem2 As New MenuItem()
            theMenuItem2.Header = "Edit"
            theContextMenu.DataContext = NonDispoList
            AddHandler theMenuItem1.Click, AddressOf Me.MenuItem1Clicked
            AddHandler theMenuItem2.Click, AddressOf Me.MenuItem2Clicked
            theContextMenu.Items.Add(theMenuItem1)
            theContextMenu.Items.Add(theMenuItem2)
            For Each aSchedulenondispo In theNonDispoCollection

                Dim myhours As Integer = aSchedulenondispo.TimeStart / 60
                Dim myminutes As Integer = aSchedulenondispo.TimeStart - (myhours * 60)
                Dim atime As New DateTime(1, 1, 1, myhours, myminutes, 0)
                myhours = aSchedulenondispo.TimeStop / 60
                myminutes = aSchedulenondispo.TimeStop - (myhours * 60)
                Dim atime2 As New DateTime(1, 1, 1, myhours, myminutes, 0)

                theListMenuItem = New ListViewItem()

                theListMenuItem.Content = "Du " + Right("0" + aSchedulenondispo.DateStart.Day.ToString(), 2) + "/" + _
                Right("0" + aSchedulenondispo.DateStart.Month.ToString(), 2) + "/" + _
                aSchedulenondispo.DateStart.Year.ToString() + "  " + _
                Right("0" + atime.Hour.ToString(), 2) + ":" + Right("0" + atime.Minute.ToString(), 2) + " Au " + _
                Right("0" + aSchedulenondispo.DateStop.Day.ToString(), 2) + "/" + _
                Right("0" + aSchedulenondispo.DateStop.Month.ToString(), 2) + "/" + _
                aSchedulenondispo.DateStop.Year.ToString() + "  " + _
                Right("0" + atime2.Hour.ToString(), 2) + ":" + Right("0" + atime2.Minute.ToString(), 2)


                theListMenuItem.ContextMenu = theContextMenu
                NonDispoList.Items.Add(theListMenuItem)
            Next

        End If

    End Sub

    Private Sub MenuItem1Clicked(sender As Object, e As System.Windows.RoutedEventArgs)
        Debug.WriteLine("MenuItem1Clicked")
        Dim theMenuItem1 As MenuItem
        theMenuItem1 = CType(sender, MenuItem)
        Dim theContextmenu As ContextMenu
        theContextmenu = theMenuItem1.Parent
        Dim theListview As ListView
        theListview = CType(theContextmenu.DataContext, ListView)
        Debug.WriteLine("selcdted item is:" + theListview.SelectedIndex.ToString())
        If theNonDispoCollection.Contains((theListview.SelectedIndex + 1).ToString()) Then
            Dim theNonDispo As ScheduleNonDispo
            theNonDispo = theNonDispoCollection(theListview.SelectedIndex + 1)
            theNonDispo.delete()
        End If
        updateListview()
    End Sub

    Private Sub MenuItem2Clicked(sender As Object, e As System.Windows.RoutedEventArgs)
        Debug.WriteLine("MenuItem2Clicked")
    End Sub
End Class
