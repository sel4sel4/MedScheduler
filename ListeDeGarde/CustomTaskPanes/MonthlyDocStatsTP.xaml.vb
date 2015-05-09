Imports System.Windows.Controls
Imports System.Windows.Data
Imports System.Windows.Media

Public Class MonthlyDocStatsTP
    Private aCollection As Collection
    Private aArray As Integer()

    Public Sub loadarray(theCollection As Collection, theArray As Integer())
        aCollection = theCollection
        aArray = theArray
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
        Dim theStats As SDocStats
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
            aLabel.VerticalContentAlignment = Windows.VerticalAlignment.Center
            Dim aRotateTransform As New RotateTransform()
            aRotateTransform.Angle = 270
            aLabel.LayoutTransform = aRotateTransform
            aHorizStackPanel.Children.Add(aLabel)
        Next
        Me.MyPanel.Children.Add(aHorizStackPanel)
        aHorizStackPanel.Name = "Header"
        Dim theLoopCounter As Integer = 1
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
                aLabel.Padding = New Windows.Thickness(3)
                aLabel.HorizontalContentAlignment = Windows.HorizontalAlignment.Center
                aLabel.BorderBrush = System.Windows.Media.Brushes.Black
                If theLoopCounter = aCollection.Count Then
                    aLabel.BorderThickness = New Windows.Thickness(1, 1, 0, 1)
                Else
                    aLabel.BorderThickness = New Windows.Thickness(1, 1, 0, 0)
                End If

                aLabel.Width = 25
                aLabel.Height = 18.5
                aHorizStackPanel.Children.Add(aLabel)
            Next
            If theLoopCounter = aCollection.Count Then
                aLabel.BorderThickness = New Windows.Thickness(1, 1, 1, 1)
            Else
                aLabel.BorderThickness = New Windows.Thickness(1, 1, 1, 0)
            End If
            theLoopCounter = theLoopCounter + 1
        Next


        'clear everything
        MyPanel2.Children.Clear()
        aLabel = New Label
        aLabel.Content = Globals.ThisAddIn.theCurrentController.pHighlightedDoc
        MyPanel2.Children.Add(aLabel)

        Dim y As Integer
        If Not aArray Is Nothing Then
            For y = 0 To UBound(aArray, 1)
                aLabel = New Label
                aLabel.Content = CStr(aArray(y))
                MyPanel2.Children.Add(aLabel)

            Next
        End If

    End Sub
End Class
