Imports System.Windows.Controls
Imports System.Windows.Data
Imports System.Windows.Media

Public Class MonthlyDocStatsTP
    Private aCollection As Collection

    Public Sub loadarray(theCollection As Collection)
        aCollection = theCollection
        DrawGrid()
    End Sub

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()

    End Sub

    Private Sub DrawGrid()

        Dim theStats As ScheduleDocStats

        Dim aHorizStackPanel As StackPanel
        Dim aLabel As Label
        ' Add any initialization after the InitializeComponent() call.
        MyPanel.Children.Clear()

        aLabel = New Label
        aLabel.Content = ""
        aLabel.Width = 50
        aLabel.Height = 70
        Dim aRotateTransform As New RotateTransform()
        aRotateTransform.Angle = 270

        aHorizStackPanel = New StackPanel
        aHorizStackPanel.Orientation = Orientation.Horizontal
        aHorizStackPanel.Height = 100
        aHorizStackPanel.Children.Add(aLabel)

        If Globals.ThisAddIn.theControllerCollection.Count < 1 Then Exit Sub
        If Not Globals.ThisAddIn.theControllerCollection.Contains(Globals.ThisAddIn.Application.ActiveSheet.name) Then Exit Sub
        Dim aController As Controller = Globals.ThisAddIn.theControllerCollection.Item(Globals.ThisAddIn.Application.ActiveSheet.name)

        For Each aShift In aController.aControlledMonth.ShiftTypes
            If aShift.ShiftType > 5 Then Exit For
            aLabel = New Label
            aLabel.Content = aShift.Description
            aLabel.Width = 70
            aLabel.Height = 25
            aLabel.LayoutTransform = aRotateTransform
            aHorizStackPanel.Children.Add(aLabel)
        Next

        Me.MyPanel.Children.Add(aHorizStackPanel)
        aHorizStackPanel.Name = "Header"

        For Each theStats In aCollection
            aHorizStackPanel = New StackPanel
            aLabel = New Label
            aLabel.Content = theStats.Initials
            aLabel.Width = 50
            aHorizStackPanel.Height = 21
            aHorizStackPanel.Orientation = Orientation.Horizontal
            Me.MyPanel.Children.Add(aHorizStackPanel)
            aHorizStackPanel.Children.Add(aLabel)

            For Each aShift In aController.aControlledMonth.ShiftTypes
                If aShift.ShiftType > 5 Then Exit For
                aLabel = New Label
                Select Case aShift.ShiftType
                    Case 1
                        aLabel.Content = CStr(theStats.shift1)
                    Case 2
                        aLabel.Content = CStr(theStats.shift2)
                    Case 3
                        aLabel.Content = CStr(theStats.shift3)
                    Case 4
                        aLabel.Content = CStr(theStats.shift4)
                    Case 5
                        aLabel.Content = CStr(theStats.shift5)
                End Select
                aLabel.Width = 25
                aHorizStackPanel.Children.Add(aLabel)
            Next
        Next
    End Sub
End Class
