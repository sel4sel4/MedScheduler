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

        If Globals.ThisAddIn.theControllerCollection.Count < 1 Then Exit Sub
        If Not Globals.ThisAddIn.theControllerCollection.Contains(Globals.ThisAddIn.Application.ActiveSheet.name) Then Exit Sub
        Dim aController As Controller = Globals.ThisAddIn.theControllerCollection.Item(Globals.ThisAddIn.Application.ActiveSheet.name)

        'clear everything
        MyPanel.Children.Clear()
        Dim theStats As ScheduleDocStats
        Dim aHorizStackPanel As StackPanel
        Dim aLabel As Label

        'create empty placeholder top left
        aLabel = New Label
        aLabel.Content = ""
        aLabel.Width = 50
        aLabel.Height = 70
        aHorizStackPanel = New StackPanel
        aHorizStackPanel.Orientation = Orientation.Horizontal
        aHorizStackPanel.Height = 96
        aHorizStackPanel.Children.Add(aLabel)

        'create shift headers
        For Each aShift In aController.aControlledMonth.ShiftTypes
            If aShift.ShiftType > 5 Then Exit For
            aLabel = New Label
            aLabel.Content = aShift.Description
            aLabel.Width = 70
            aLabel.Height = 25
            Dim aRotateTransform As New RotateTransform()
            aRotateTransform.Angle = 270
            aLabel.LayoutTransform = aRotateTransform
            aHorizStackPanel.Children.Add(aLabel)
        Next
        Me.MyPanel.Children.Add(aHorizStackPanel)
        aHorizStackPanel.Name = "Header"

        'create doc list with shifts counts
        For Each theStats In aCollection
            aHorizStackPanel = New StackPanel
            aLabel = New Label
            aLabel.Content = theStats.Initials
            aLabel.Width = 50
            aLabel.Height = 18.5
            aLabel.Padding = New Windows.Thickness(4)

            aHorizStackPanel.Height = 18.5
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
                If theStats.Initials = Globals.ThisAddIn.theCurrentController.pHighlightedDoc Then
                    aLabel.Background = New SolidColorBrush(Color.FromRgb(150, 100, 150))
                End If
                aLabel.Padding = New Windows.Thickness(4)
                aLabel.Width = 25
                aLabel.Height = 18.5
                aHorizStackPanel.Children.Add(aLabel)
            Next
        Next
    End Sub
End Class
