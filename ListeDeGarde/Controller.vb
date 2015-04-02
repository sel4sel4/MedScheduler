Public Class Controller
    Private WithEvents controlledExcelSheet As Excel.Worksheet
    Private controlledMonth As ScheduleMonth
    Private controlledShiftTypes As ScheduleShiftType
    Private monthloaded As Boolean = False

    Public ReadOnly Property aControlledMonth() As ScheduleMonth
        Get
            Return controlledMonth
        End Get
    End Property

    Public Sub New(aSheet As Excel.Worksheet, aYear As Integer, aMonth As Integer, aMonthString As String)

        'load the sheet
        controlledExcelSheet = aSheet

        'create a month
        controlledMonth = New ScheduleMonth(aMonth, aYear)

        'Load shift types collection into global
        controlledShiftTypes = New ScheduleShiftType

        Dim theDay As ScheduleDay
        Dim row As Integer
        Dim col As Integer = 0

        'get number of shifts
        Dim rowheight1 As Integer = controlledShiftTypes.ShiftCollection.Count + 1
        'assign colwidth as 2
        Dim colwidth1 As Integer = 2

        'populate the top left corner of sheet with year and month strings
        controlledExcelSheet.Range("A1").Value = aMonthString
        controlledExcelSheet.Range("B1").Value = aYear.ToString()

        'set top left corner of calendar
        Dim theRangeA3 As Excel.Range = controlledExcelSheet.Range("A3")
        Dim theRange As Excel.Range


        For Each theDay In controlledMonth.Days
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

        monthloaded = True

    End Sub

    Private Sub controlledExcelSheet_Change(ByVal Target As Excel.Range) Handles controlledExcelSheet.Change

        If monthloaded = False Then Exit Sub

        System.Diagnostics.Debug.WriteLine("WithEvents: You Changed Cells " + Target.Address + " " + controlledExcelSheet.Name)
        Dim aday As ScheduleDay
        Dim aShift As ScheduleShift
        Dim adocAvail As scheduleDocAvailable
        Dim anExitNotice As Boolean = False

        For Each aday In controlledMonth.Days
            For Each aShift In aday.Shifts
                If aShift.aRange.Address = Target.Address Then
                    For Each adocAvail In aShift.DocAvailabilities
                        If adocAvail.DocInitial = Target.Value Then
                            adocAvail.Availability = PublicEnums.Availability.Assigne
                            fixAvailability(Target.Value, controlledMonth, aShift)
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
    End Sub

    Public Sub HighLightDocAvailablilities(Initials As String)
        'cycle through the month and highlight everywhere theDoc is available.
        Dim aday As ScheduleDay
        Dim aShift As ScheduleShift
        Dim adocAvail As scheduleDocAvailable

        For Each aday In controlledMonth.Days
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

End Class
