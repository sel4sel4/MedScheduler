Imports System.Windows.Controls
Imports System.Windows.Media


Public Class DocExpectations





    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        Dim aHorizStackPanel As StackPanel
        Dim aLabel As Label
        Dim aTextBox As TextBox
        Dim aShift As ScheduleShiftType
        ' Add any initialization after the InitializeComponent() call.
        If Not Globals.ThisAddIn.theControllerCollection.Contains(Globals.ThisAddIn.Application.ActiveSheet.name) Then Exit Sub
        Dim aController As Controller = Globals.ThisAddIn.theControllerCollection.Item(Globals.ThisAddIn.Application.ActiveSheet.name)
        Dim aScheduleDoc As ScheduleDoc

        aLabel = New Label
        aLabel.Content = ""
        aLabel.Width = 120
        aLabel.Height = 70
        Dim aRotateTransform As New RotateTransform()
        aRotateTransform.Angle = 270

        aHorizStackPanel = New StackPanel
        aHorizStackPanel.Orientation = Orientation.Horizontal
        aHorizStackPanel.Height = 100
        aHorizStackPanel.Children.Add(aLabel)


        For Each aShift In aController.aControlledMonth.ShiftTypes
            If aShift.ShiftType > 5 Then Exit For
            aLabel = New Label
            aLabel.Content = aShift.Description
            aLabel.Width = 70
            aLabel.Height = 30
            aLabel.LayoutTransform = aRotateTransform
            aHorizStackPanel.Children.Add(aLabel)
        Next
        aLabel = New Label
        aLabel.Content = "Total"
        aLabel.Width = 70
        aLabel.Height = 30
        aLabel.LayoutTransform = aRotateTransform
        aHorizStackPanel.Children.Add(aLabel)

        Me.MyPanel.Children.Add(aHorizStackPanel)
        aHorizStackPanel.Name = "Header"
        MyPanel.RegisterName(aHorizStackPanel.Name, aHorizStackPanel)

        For Each aScheduleDoc In aController.aControlledMonth.DocList
            aHorizStackPanel = New StackPanel
            aLabel = New Label
            aLabel.Content = aScheduleDoc.FistAndLastName
            aLabel.Width = 120
            aHorizStackPanel.Height = 21
            aHorizStackPanel.Orientation = Orientation.Horizontal
            Me.MyPanel.Children.Add(aHorizStackPanel)
            aHorizStackPanel.Name = aScheduleDoc.Initials
            MyPanel.RegisterName(aHorizStackPanel.Name, aHorizStackPanel)
            aHorizStackPanel.Children.Add(aLabel)

            For Each aShift In aController.aControlledMonth.ShiftTypes
                If aShift.ShiftType > 5 Then Exit For
                aTextBox = New TextBox
                aTextBox.Text = "0"
                aTextBox.Width = 30
                aTextBox.Name = aScheduleDoc.Initials + "_" + CStr(aShift.ShiftType)
                AddHandler aTextBox.TextChanged, AddressOf TextHasChanged
                aHorizStackPanel.Children.Add(aTextBox)
                MyPanel.RegisterName(aTextBox.Name, aTextBox)

            Next
            aTextBox = New TextBox
            aTextBox.Text = "0"
            aTextBox.Width = 30
            aTextBox.IsEnabled = False
            aTextBox.Name = "Total_" + aScheduleDoc.Initials
            aHorizStackPanel.Children.Add(aTextBox)
            MyPanel.RegisterName(aTextBox.Name, aTextBox)
        Next

        aHorizStackPanel = New StackPanel
        aLabel = New Label
        aLabel.Content = "Total:"
        aLabel.Width = 120
        aHorizStackPanel.Height = 21
        aHorizStackPanel.Orientation = Orientation.Horizontal
        Me.MyPanel.Children.Add(aHorizStackPanel)
        aHorizStackPanel.Name = "Total"
        MyPanel.RegisterName(aHorizStackPanel.Name, aHorizStackPanel)
        aHorizStackPanel.Children.Add(aLabel)

        For x = 1 To 5
            aTextBox = New TextBox
            aTextBox.Text = "0"
            aTextBox.Width = 30
            aTextBox.Name = "Total_" + CStr(x)
            AddHandler aTextBox.TextChanged, AddressOf TextHasChanged
            aHorizStackPanel.Children.Add(aTextBox)
            MyPanel.RegisterName(aTextBox.Name, aTextBox)

        Next

    End Sub
    Private Sub TextHasChanged(ByVal sender As Object, ByVal e As Windows.RoutedEventArgs)
        Dim myTextBox As TextBox
        'Dim theDocInitials As String
        myTextBox = CType(sender, TextBox)
        If Left(myTextBox.Name, 5) = "Total" Then Exit Sub
        Dim mySplit As String() = myTextBox.Name.Split(New Char() {"_"c})
        'Dim aDependencyObject As Windows.DependencyObject
        'aDependencyObject = VisualTreeHelper.GetParent(myTextBox)
        'Dim aStackPanel As StackPanel
        'aStackPanel = CType(aDependencyObject, StackPanel)
        'theDocInitials = aStackPanel.Name
        'Windows.MessageBox.Show(mySplit(0) + ":" + mySplit(1))
        Dim x As Integer
        Dim aObject As Object
        Dim aTextBox As TextBox
        Dim aTotal As Integer = 0
        For x = 1 To 5
            aObject = MyPanel.FindName(mySplit(0) + "_" + x.ToString())
            aTextBox = CType(aObject, TextBox)
            aTotal = aTotal + CInt(aTextBox.Text)
        Next
        aObject = MyPanel.FindName("Total_" + mySplit(0))
        aTextBox = CType(aObject, TextBox)
        aTextBox.Text = CStr(aTotal)

        If Not Globals.ThisAddIn.theControllerCollection.Contains(Globals.ThisAddIn.Application.ActiveSheet.name) Then Exit Sub
        Dim aController As Controller = Globals.ThisAddIn.theControllerCollection.Item(Globals.ThisAddIn.Application.ActiveSheet.name)
        Dim aScheduleDoc As ScheduleDoc
        aTotal = 0
        For Each aScheduleDoc In aController.aControlledMonth.DocList
            aObject = MyPanel.FindName(aScheduleDoc.Initials + "_" + mySplit(1))
            aTextBox = CType(aObject, TextBox)
            aTotal = aTotal + CInt(aTextBox.Text)
        Next
        aObject = MyPanel.FindName("Total_" + mySplit(1))
        aTextBox = CType(aObject, TextBox)
        aTextBox.Text = CStr(aTotal)
    End Sub
End Class
