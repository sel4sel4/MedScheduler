Imports System.Windows.Controls
Imports System.Diagnostics


Public Class UserControl2
    ' Private WithEvents xlButton As Excel.OLEObject

    'Structure MonthToSheet
    '    Public pScheduleMonth As ScheduleMonth
    '    Public pSheet As Excel.Worksheet
    '    Public pController As Controller
    'End Structure

    Private aMonthToSheetColl As Collection
    Private aMonthToSheetCounter As Integer = -1
    Private atLeastOneMonthSetup As Boolean = False
    Public WithEvents xlSheet2 As Excel.Worksheet = Globals.ThisAddIn.xlSheet1
    Private WithEvents newBtn As Button
    Public WithEvents xlApp As Excel.Application


    Private Sub Button_Click(sender As Object, e As Windows.RoutedEventArgs)
        atLeastOneMonthSetup = False
        If aMonthToSheetColl Is Nothing Then aMonthToSheetColl = New Collection
        Dim aMonthToSheet As MonthToSheet
        'get references to app, workbook
        Dim xlApp As Excel.Application = Globals.ThisAddIn.Application
        Dim wb As Excel.Workbook = Globals.ThisAddIn.Application.ActiveWorkbook

        'create a new sheet
        With wb
            Globals.ThisAddIn.xlSheet1 = DirectCast(.Sheets.Add(After:=.Sheets(.Sheets.Count), Count:=1, Type:=Excel.XlSheetType.xlWorksheet), Excel.Worksheet)
        End With
        aMonthToSheet.pSheet = Globals.ThisAddIn.xlSheet1
        'rename the new sheet
        Globals.ThisAddIn.xlSheet1.Name = Me.combo2.Text + "-" + Me.combo1.Text
        Dim theCodename As String = Globals.ThisAddIn.xlSheet1.Name
        xlSheet2 = Globals.ThisAddIn.xlSheet1
        'pull year and month data from the user control
        Dim aMonth As Integer = Me.combo2.SelectedIndex + 1
        Dim aYear As Integer = CInt(Me.combo1.Text)

        'pull colwidth and rowheight from user control
        'Dim colwidth1 As Integer = CInt(Me.tb1.Text)
        Dim colwidth1 As Integer = 2
        'Dim rowheight1 As Integer = CInt(Me.tb2.Text)

        'populate the top left corner of sheet with year and month strings
        Dim theRangeA1 As Excel.Range = Globals.ThisAddIn.xlSheet1.Range("A1")
        Dim theRangeB1 As Excel.Range = Globals.ThisAddIn.xlSheet1.Range("B1")
        theRangeA1.Value = Me.combo2.Text
        theRangeB1.Value = Me.combo1.Text

        Dim theRangeA3 As Excel.Range = Globals.ThisAddIn.xlSheet1.Range("A3")
        Dim theRange As Excel.Range

        Dim theScheduleDoc As New ScheduleDoc(aYear, aMonth)
        Dim aScheduleDoc As ScheduleDoc
        RemoveDocButtons()
        For Each aScheduleDoc In theScheduleDoc.DocList
            Me.addDocButton(aScheduleDoc.FirstName + " " + aScheduleDoc.LastName, aScheduleDoc.Initials)
        Next



        'Load shift types collection into global
        globalShiftTypes = New ScheduleShiftType


        Dim theMonth As New ScheduleMonth(aMonth, aYear)
        aMonthToSheet.pScheduleMonth = theMonth
        aMonthToSheetColl.Add(aMonthToSheet, theCodename)

        Dim theDay As ScheduleDay

        Dim row As Integer, col As Integer

        col = 0

        'set starting point for generating Calendar

        'get number of shifts
        Dim rowheight1 As Integer = globalShiftTypes.ShiftCollection.Count + 1

        For Each theDay In theMonth.Days
            col = CInt(theDay.theDate.DayOfWeek)
            theRange = theRangeA3.Offset(row * rowheight1, col * colwidth1)

            Dim theRangeForShiftType As Excel.Range
            Dim TheRAngeForDocLists As Excel.Range
            Dim theShift As ScheduleShift

            Dim theCounter1 As Integer = 1
            For Each theShift In theDay.Shifts
                ' Dim theSetValue As String = ""
                theRangeForShiftType = theRange.Offset(theCounter1, 0)
                theRangeForShiftType.Value2 = "'" + theShift.Description
                TheRAngeForDocLists = theRange.Offset(theCounter1, 1)
                theShift.aRange = TheRAngeForDocLists
   
                fixlist(theShift)
                theCounter1 = theCounter1 + 1
            Next

            theRange.Offset(0, colwidth1 - 1).Value = theDay.theDate.Day
            theRange = theRange.Resize(rowheight1, colwidth1)
            addBordersAroundRange(theRange)
            If col = 6 Then row = row + 1
        Next
        atLeastOneMonthSetup = True



    End Sub

    Private Sub addBordersAroundRange(aRange As Excel.Range)

        With aRange.Borders(Excel.XlBordersIndex.xlEdgeBottom)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .Weight = Excel.XlBorderWeight.xlThin
            .ColorIndex = Excel.Constants.xlAutomatic
        End With
        With aRange.Borders(Excel.XlBordersIndex.xlEdgeTop)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .Weight = Excel.XlBorderWeight.xlThin
            .ColorIndex = Excel.Constants.xlAutomatic
        End With
        With aRange.Borders(Excel.XlBordersIndex.xlEdgeLeft)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .Weight = Excel.XlBorderWeight.xlThin
            .ColorIndex = Excel.Constants.xlAutomatic
        End With
        With aRange.Borders(Excel.XlBordersIndex.xlEdgeRight)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .Weight = Excel.XlBorderWeight.xlThin
            .ColorIndex = Excel.Constants.xlAutomatic
        End With

    End Sub

    Private Sub combo1_Loaded(sender As Object, e As Windows.RoutedEventArgs)
        Dim theList As New List(Of String)
        theList.Add("2014")
        theList.Add("2015")
        theList.Add("2016")
        theList.Add("2017")
        Dim theComboBox As ComboBox
        theComboBox = CType(sender, ComboBox)
        theComboBox.ItemsSource = theList
        theComboBox.SelectedIndex = 0
    End Sub

    Public Sub addDocButton(docName As String, initials As String)
        Dim newBtn = New Button()
        newBtn.Content = docName
        newBtn.Name = initials
        AddHandler newBtn.Click, AddressOf onBtnClick
        SplMain.Children.Add(newBtn)

    End Sub

    Public Sub RemoveDocButtons()
        Dim aButton As Button
        For x As Integer = 0 To SplMain.Children.Count - 1
            aButton = SplMain.Children.Item(0)
            RemoveHandler aButton.Click, AddressOf onBtnClick
            SplMain.Children.RemoveAt(0)
        Next
    End Sub

    Private Sub onBtnClick(ByVal sender As Object, ByVal e As Windows.RoutedEventArgs)
        Dim aButton As Button = CType(sender, Button)
        Debug.WriteLine(aButton.Content)
        HighLightDocAvailablilities(aButton.Name)

    End Sub

    Private Sub HighLightDocAvailablilities(Initials As String)
        'cycle through the month and highlight everywhere theDoc is available.
        Dim aMonth As ScheduleMonth
        Dim aday As ScheduleDay
        Dim aShift As ScheduleShift
        Dim adocAvail As scheduleDocAvailable
        Dim aMontoShet As MonthToSheet = aMonthToSheetColl(Globals.ThisAddIn.xlSheet1.Name)
        aMonth = aMontoShet.pScheduleMonth

        If Not IsNothing(aMonth) Then
            For Each aday In aMonth.Days
                For Each aShift In aday.Shifts
                    For Each adocAvail In aShift.DocAvailabilities
                        If adocAvail.DocInitial = Initials Then
                            Select Case adocAvail.Availability
                                Case PublicEnums.Availability.Dispo
                                    aShift.aRange.Interior.Color = RGB(0, 233, 118)
                                Case PublicEnums.Availability.Assigne
                                    aShift.aRange.Interior.Color = RGB(0, 255, 255)
                                Case PublicEnums.Availability.NonDispoPermanente
                                    aShift.aRange.Interior.Color = RGB(220, 20, 60)
                                Case PublicEnums.Availability.NonDispoTemporaire
                                    aShift.aRange.Interior.Color = RGB(219, 112, 147)
                                Case PublicEnums.Availability.SurUtilise
                                    aShift.aRange.Interior.Color = RGB(209, 95, 238)
                                Case Else

                            End Select
                        End If

                    Next
                Next
            Next
        End If
    End Sub

    Private Sub xlSheet2_Change(ByVal Target As Excel.Range) Handles xlSheet2.Change

        If atLeastOneMonthSetup = False Then Exit Sub
        System.Diagnostics.Debug.WriteLine("WithEvents: You Changed Cells " + Target.Address + " " + xlSheet2.Name)
        Dim aMonth As ScheduleMonth
        Dim aday As ScheduleDay
        Dim aShift As ScheduleShift
        Dim adocAvail As scheduleDocAvailable
        Dim anExitNotice As Boolean = False
        Dim aMontoShet As MonthToSheet = aMonthToSheetColl(Globals.ThisAddIn.xlSheet1.Name)
        aMonth = aMontoShet.pScheduleMonth
        If Not IsNothing(aMonth) Then
            For Each aday In aMonth.Days
                For Each aShift In aday.Shifts
                    If aShift.aRange.Address = Target.Address Then
                        For Each adocAvail In aShift.DocAvailabilities
                            If adocAvail.DocInitial = Target.Value Then
                                adocAvail.Availability = PublicEnums.Availability.Assigne
                                fixAvailability(Target.Value, aMonth, aShift)
                                anExitNotice = True
                                Exit For
                            End If
                            If anExitNotice = True Then Exit For
                        Next
                    End If
                    If anExitNotice = True Then Exit For
                Next
                If anExitNotice = True Then Exit For
            Next
        End If
    End Sub

    Private Sub fixAvailability(aDoc As String, aMonth As ScheduleMonth, ashift As ScheduleShift)
        Dim theDate As Date = ashift.aDate
        Dim theShift As Integer = ashift.ShiftType
        Dim theshiftStart As Integer = ashift.ShiftStart
        Dim theshiftStop As Integer = ashift.ShiftStop
        Dim theStartDay As Integer = theDate.Day - 1
        Dim theStopDay As Integer = theDate.Day + 1
        Dim myShift As ScheduleShift
        Dim aDay As ScheduleDay = aMonth.Days(ashift.aDate.Day)
        For Each myShift In aDay.Shifts
            If myShift.ShiftStop > ashift.ShiftStart - 920 Or ashift.ShiftStop + 920 > myShift.ShiftStart Then
                Dim thedocAvail As scheduleDocAvailable
                For Each thedocAvail In myShift.DocAvailabilities
                    If thedocAvail.DocInitial = aDoc And _
                            thedocAvail.Availability <> PublicEnums.Availability.NonDispoPermanente And _
                            thedocAvail.Availability <> PublicEnums.Availability.Assigne Then
                        thedocAvail.Availability = PublicEnums.Availability.NonDispoTemporaire
                        fixlist(myShift)
                    End If
                Next
            End If
        Next
        aDay = aMonth.Days(ashift.aDate.Day - 1) 'FIX: check if first day of month
        For Each myShift In aDay.Shifts
            If myShift.ShiftStop > ashift.ShiftStart - 920 + 1440 Then
                Dim thedocAvail As scheduleDocAvailable
                For Each thedocAvail In myShift.DocAvailabilities
                    If thedocAvail.DocInitial = aDoc And _
                        thedocAvail.Availability <> PublicEnums.Availability.NonDispoPermanente And _
                        thedocAvail.Availability <> PublicEnums.Availability.Assigne Then
                        thedocAvail.Availability = PublicEnums.Availability.NonDispoTemporaire
                        fixlist(myShift)
                    End If
                Next
            End If
        Next
        aDay = aMonth.Days(ashift.aDate.Day + 1) 'FIX: check if last day of month
        For Each myShift In aDay.Shifts
            If ashift.ShiftStop + 920 - 1440 > myShift.ShiftStart Then
                Dim thedocAvail As scheduleDocAvailable
                For Each thedocAvail In myShift.DocAvailabilities
                    If thedocAvail.DocInitial = aDoc And _
                        thedocAvail.Availability <> PublicEnums.Availability.NonDispoPermanente And _
                        thedocAvail.Availability <> PublicEnums.Availability.Assigne Then
                        thedocAvail.Availability = PublicEnums.Availability.NonDispoTemporaire
                        fixlist(myShift)
                    End If
                Next
            End If
        Next

    End Sub

    Public Sub fixlist(theShift As ScheduleShift)
        Dim theSetValue As String = ""
        Dim theDocAvailable As scheduleDocAvailable
        Dim thelist As String = ""
        For Each theDocAvailable In theShift.DocAvailabilities
            Select Case theDocAvailable.Availability
                Case PublicEnums.Availability.Dispo
                    thelist = thelist + theDocAvailable.DocInitial + ","
                Case PublicEnums.Availability.Assigne
                    theSetValue = theDocAvailable.DocInitial
                Case Else

            End Select
        Next
        If thelist.Length > 0 Then thelist = Left(thelist, thelist.Length - 1)

        With theShift.aRange.Validation
            .Delete()
            If thelist <> "" Then
                .Add(Type:=Excel.XlDVType.xlValidateList, _
                     AlertStyle:=Excel.XlDVAlertStyle.xlValidAlertStop, _
                     Operator:=Excel.XlFormatConditionOperator.xlBetween, _
                     Formula1:=thelist)
                .IgnoreBlank = True
                .InCellDropdown = True
                .InputTitle = ""
                .ErrorTitle = ""
                .InputMessage = ""
                .ErrorMessage = ""
                .ShowInput = True
                .ShowError = True
            End If
        End With
        ' If theSetValue <> "" Then theShift.aRange.Value = theSetValue



    End Sub

    Private Sub xlSheet2_BeforeDelete() Handles xlSheet2.BeforeDelete
        aMonthToSheetColl.Remove(xlSheet2.Name)

    End Sub

    Private Sub ComboBox_Loaded_1(sender As Object, e As Windows.RoutedEventArgs)
        Dim theList As New List(Of String)
        theList.Add("Janvier")
        theList.Add("Fevrier")
        theList.Add("Mars")
        theList.Add("Avril")
        theList.Add("May")
        theList.Add("Juin")
        theList.Add("Juillet")
        theList.Add("Aout")
        theList.Add("Septembre")
        theList.Add("Octobre")
        theList.Add("Novembre")
        theList.Add("Decembre")
        Dim theComboBox As ComboBox
        theComboBox = CType(sender, ComboBox)
        theComboBox.ItemsSource = theList
        theComboBox.SelectedIndex = 0
    End Sub

End Class


