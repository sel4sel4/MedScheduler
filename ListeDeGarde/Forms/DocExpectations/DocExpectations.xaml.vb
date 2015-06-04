Imports System.Windows.Controls
Imports System.Windows.Media


Public Class DocExpectations

    Private theDocCollection As List(Of SDoc)
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        DrawGrid()



    End Sub

    Private Sub DrawGrid()

        Dim aHorizStackPanel As StackPanel
        Dim aLabel As Label
        Dim aTextBox As TextBox
        Dim aShift As SShiftType
        Dim aSDoc As SDoc

        'If controller exists, get it
        If Not Globals.ThisAddIn.theControllerCollection.Exists(Function(xy) xy.aControlledExcelSheet.Name = Globals.ThisAddIn.Application.ActiveSheet.name) Then Exit Sub
        Dim aController As Controller = Globals.ThisAddIn.theControllerCollection.Find(Function(xy) xy.aControlledExcelSheet.Name = Globals.ThisAddIn.Application.ActiveSheet.name)


        MyPanel.Children.Clear()

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

        If Not MyPanel.FindName(aHorizStackPanel.Name) Is Nothing Then MyPanel.UnregisterName(aHorizStackPanel.Name)
        MyPanel.RegisterName(aHorizStackPanel.Name, aHorizStackPanel)


        If Me.Edit_Template.IsChecked = False Then
            theDocCollection = aController.aControlledMonth.DocList
        Else
            theDocCollection = SDoc.LoadTempateDocsFromDB()
        End If

        For Each aSDoc In theDocCollection
            aHorizStackPanel = New StackPanel
            aLabel = New Label
            aLabel.Content = aSDoc.FistAndLastName
            aLabel.Width = 120
            aHorizStackPanel.Height = 21
            aHorizStackPanel.Orientation = Orientation.Horizontal
            Me.MyPanel.Children.Add(aHorizStackPanel)
            aHorizStackPanel.Name = aSDoc.Initials
            If Not MyPanel.FindName(aHorizStackPanel.Name) Is Nothing Then MyPanel.UnregisterName(aHorizStackPanel.Name)
            MyPanel.RegisterName(aHorizStackPanel.Name, aHorizStackPanel)
            aHorizStackPanel.Children.Add(aLabel)

            For Each aShift In aController.aControlledMonth.ShiftTypes
                If aShift.ShiftType > 5 Then Exit For
                aTextBox = New TextBox
                Select Case aShift.ShiftType
                    Case 1
                        aTextBox.Text = CStr(aSDoc.Shift1)
                    Case 2
                        aTextBox.Text = CStr(aSDoc.Shift2)
                    Case 3
                        aTextBox.Text = CStr(aSDoc.Shift3)
                    Case 4
                        aTextBox.Text = CStr(aSDoc.Shift4)
                    Case 5
                        aTextBox.Text = CStr(aSDoc.Shift5)
                End Select
                aTextBox.Width = 30
                aTextBox.Name = aSDoc.Initials + "_" + CStr(aShift.ShiftType)
                AddHandler aTextBox.TextChanged, AddressOf TextHasChanged
                aHorizStackPanel.Children.Add(aTextBox)
                If Not MyPanel.FindName(aTextBox.Name) Is Nothing Then MyPanel.UnregisterName(aTextBox.Name)
                MyPanel.RegisterName(aTextBox.Name, aTextBox)

            Next
            aTextBox = New TextBox
            aTextBox.Text = "0"
            aTextBox.Width = 30
            aTextBox.IsEnabled = False
            aTextBox.Name = "Total_" + aSDoc.Initials
            aHorizStackPanel.Children.Add(aTextBox)
            If Not MyPanel.FindName(aTextBox.Name) Is Nothing Then MyPanel.UnregisterName(aTextBox.Name)
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
        If Not MyPanel.FindName(aHorizStackPanel.Name) Is Nothing Then MyPanel.UnregisterName(aHorizStackPanel.Name)
        MyPanel.RegisterName(aHorizStackPanel.Name, aHorizStackPanel)
        aHorizStackPanel.Children.Add(aLabel)

        For x = 1 To 5
            aTextBox = New TextBox
            aTextBox.Text = "0"
            aTextBox.Width = 30
            aTextBox.IsEnabled = False
            aTextBox.Name = "Total_" + CStr(x)
            aHorizStackPanel.Children.Add(aTextBox)
            If Not MyPanel.FindName(aTextBox.Name) Is Nothing Then MyPanel.UnregisterName(aTextBox.Name)
            MyPanel.RegisterName(aTextBox.Name, aTextBox)

        Next

        aHorizStackPanel = New StackPanel
        aLabel = New Label
        aLabel.Content = "Expected:"
        aLabel.Width = 120
        aHorizStackPanel.Height = 21
        aHorizStackPanel.Orientation = Orientation.Horizontal
        Me.MyPanel.Children.Add(aHorizStackPanel)
        aHorizStackPanel.Children.Add(aLabel)
        Dim theArray As Integer()
        theArray = CountExpectedShiftsPerMonth()
        For x = 0 To 4
            aLabel = New Label
            aLabel.Content = theArray(x)
            aLabel.Width = 30
            aHorizStackPanel.Children.Add(aLabel)
        Next
        CalculateTotals()

    End Sub

    Private Sub TextHasChanged(ByVal sender As Object, ByVal e As Windows.RoutedEventArgs)
        Dim myTextBox As TextBox
        myTextBox = CType(sender, TextBox)
        If Left(myTextBox.Name, 5) = "Total" Then Exit Sub
        If Not IsNumeric(myTextBox.Text) Then
            myTextBox.Text = "0"
            Exit Sub
        End If

        CalculateTotals()

    End Sub

    Private Sub Edit_Template_Checked(sender As Object, e As Windows.RoutedEventArgs) Handles Edit_Template.Checked

        DrawGrid()
    End Sub
    Private Sub Edit_Template_Unchecked(sender As Object, e As Windows.RoutedEventArgs) Handles Edit_Template.Unchecked
        DrawGrid()
    End Sub

    Private Sub SaveBtn_Click(sender As Object, e As Windows.RoutedEventArgs) Handles SaveBtn.Click
        'cycle through all doctors, load the shift numbers from the grid
        'apply them to each doc and save them either to the template or to the specific month.

        Dim aSDoc As SDoc

        For Each aSDoc In theDocCollection
            aSDoc.Shift1 = CInt(MyPanel.FindName(aSDoc.Initials + "_1").text)
            aSDoc.Shift2 = CInt(MyPanel.FindName(aSDoc.Initials + "_2").text)
            aSDoc.Shift3 = CInt(MyPanel.FindName(aSDoc.Initials + "_3").text)
            aSDoc.Shift4 = CInt(MyPanel.FindName(aSDoc.Initials + "_4").text)
            aSDoc.Shift5 = CInt(MyPanel.FindName(aSDoc.Initials + "_5").text)
            aSDoc.save()
        Next
        Globals.ThisAddIn.theCurrentController.resetSheetExt()

    End Sub

    Private Sub CalculateTotals()

        Dim x As Integer
        Dim aObject As Object
        Dim aTextBox As TextBox
        Dim aSDoc As SDoc
        Dim horizTotal As Integer = 0
        Dim vert1Total As Integer = 0
        Dim vert2Total As Integer = 0
        Dim vert3Total As Integer = 0
        Dim vert4Total As Integer = 0
        Dim vert5Total As Integer = 0


        For Each aSDoc In theDocCollection
            For x = 1 To 5
                aObject = MyPanel.FindName(aSDoc.Initials + "_" + x.ToString())
                aTextBox = CType(aObject, TextBox)
                horizTotal = horizTotal + CInt(aTextBox.Text)
                Select Case x
                    Case 1
                        vert1Total = vert1Total + CInt(aTextBox.Text)
                    Case 2
                        vert2Total = vert2Total + CInt(aTextBox.Text)
                    Case 3
                        vert3Total = vert3Total + CInt(aTextBox.Text)
                    Case 4
                        vert4Total = vert4Total + CInt(aTextBox.Text)
                    Case 5
                        vert5Total = vert5Total + CInt(aTextBox.Text)

                End Select


            Next
            aObject = MyPanel.FindName("Total_" + aSDoc.Initials)
            aTextBox = CType(aObject, TextBox)
            aTextBox.Text = CStr(horizTotal)
            horizTotal = 0
        Next
        MyPanel.FindName("Total_1").text = CStr(vert1Total)
        MyPanel.FindName("Total_2").text = CStr(vert2Total)
        MyPanel.FindName("Total_3").text = CStr(vert3Total)
        MyPanel.FindName("Total_4").text = CStr(vert4Total)
        MyPanel.FindName("Total_5").text = CStr(vert5Total)
    End Sub

    Private Function CountExpectedShiftsPerMonth() As Integer()
        Dim theArray As Integer()
        ReDim theArray(4)

        Dim theControlledMonth As SMonth
        theControlledMonth = Globals.ThisAddIn.theCurrentController.aControlledMonth
        Dim aDay As SDay
        Dim ashift As SShift
        For Each aDay In theControlledMonth.Days
            For Each ashift In aDay.Shifts
                If ashift.ShiftType <= 5 Then theArray(ashift.ShiftType - 1) = theArray(ashift.ShiftType - 1) + 1
            Next
        Next
        Return theArray
    End Function
End Class
