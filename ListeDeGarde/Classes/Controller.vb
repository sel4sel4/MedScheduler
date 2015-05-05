Public Class Controller
    Private WithEvents controlledExcelSheet As Excel.Worksheet
    Private controlledMonth As ScheduleMonth
    Private monthloaded As Boolean = False
    Private monthlystats As UserControl4
    Private WithEvents theMonthlyStatsForm As Form2
    Private ScheduleDocStatsCollection As Collection
    Private Const theRestTime As Long = 432000000000
    Private theHighlightedDoc As String
    Private theCustomTaskPane As Microsoft.Office.Tools.CustomTaskPane

    Public ReadOnly Property aControlledMonth() As ScheduleMonth
        Get
            Return controlledMonth
        End Get
    End Property

    Public ReadOnly Property pHighlightedDoc() As String
        Get
            Return theHighlightedDoc
        End Get
    End Property

    Public Sub New(aSheet As Excel.Worksheet, aYear As Integer, aMonth As Integer, aMonthString As String)

        'load the sheet
        controlledExcelSheet = aSheet

        'create a month
        controlledMonth = New ScheduleMonth(aMonth, aYear)

        'Load shift types collection into global
        'controlledShiftTypes = controlledMonth.ShiftTypes
        theHighlightedDoc = ""
        Globals.ThisAddIn.theCurrentController = Me
        resetSheet()
       

    End Sub

    Private Sub controlledExcelSheet_Change(ByVal Target As Excel.Range) Handles controlledExcelSheet.Change

        If monthloaded = False Then Exit Sub
        controlledExcelSheet.Unprotect()
        'System.Diagnostics.Debug.WriteLine("WithEvents: You Changed Cells " + Target.Address + " " + controlledExcelSheet.Name)
        Dim aday As ScheduleDay
        Dim aShift As ScheduleShift
        Dim adocAvail As scheduleDocAvailable
        Dim anExitNotice As Boolean = False
        Dim firstDoc As String = ""

        For Each aday In controlledMonth.Days
            For Each aShift In aday.Shifts
                If aShift.aRange.Address = Target.Address Then
                    'make current Doc dispo again
                    If Not IsNothing(aShift.Doc) Then
                        If aShift.DocAvailabilities.Contains(aShift.Doc) Then
                            adocAvail = aShift.DocAvailabilities.Item(aShift.Doc)
                            adocAvail.Availability = PublicEnums.Availability.Dispo
                            firstDoc = aShift.Doc
                            anExitNotice = True
                        End If
                    End If

                    ''assign new doc
                    If IsNothing(Target.Value) Then
                        'adocAvail = aShift.DocAvailabilities.Item(firstDoc)
                        'adocAvail.Availability = PublicEnums.Availability.Dispo
                        fixAvailability(firstDoc, controlledMonth, aShift, firstDoc)
                        aShift.Doc = ""
                    Else
                        If aShift.DocAvailabilities.Contains(Target.Value) Then
                            adocAvail = aShift.DocAvailabilities.Item(Target.Value)
                            adocAvail.Availability = PublicEnums.Availability.Assigne
                            fixAvailability(Target.Value, controlledMonth, aShift, firstDoc)
                            aShift.Doc = Target.Value
                            anExitNotice = True
                        End If
                    End If
                End If
                If anExitNotice = True Then Exit For
            Next
            If anExitNotice = True Then Exit For
        Next
        If anExitNotice = True Then
            ' resetSheet()
            If Not theCustomTaskPane Is Nothing Then
                If theCustomTaskPane.Visible Then
                    statsMensuellesUpdate()
                End If
            End If
        End If
        controlledExcelSheet.Protect()
    End Sub

    Public Sub HighLightDocAvailablilities(Initials As String)
        'cycle through the month and highlight everywhere theDoc is available.
        Dim aday As ScheduleDay
        Dim aShift As ScheduleShift
        Dim adocAvail As scheduleDocAvailable
        'Globals.ThisAddIn.Application.ScreenUpdating = False
        controlledExcelSheet.Unprotect()
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
        theHighlightedDoc = Initials
        'Globals.ThisAddIn.Application.ScreenUpdating = True
        statsMensuellesUpdate()
        controlledExcelSheet.Protect()
    End Sub

    Public Sub HighLightDocAvailSingleCell(theShift As ScheduleShift, Initials As String)
        'cycle through the month and highlight everywhere theDoc is available.

        Dim adocAvail As scheduleDocAvailable
        For Each adocAvail In theShift.DocAvailabilities
            If adocAvail.DocInitial = Initials Then
                Select Case adocAvail.Availability
                    Case PublicEnums.Availability.Dispo
                        theShift.aRange.Interior.Color = RGB(0, 233, 118)
                    Case PublicEnums.Availability.Assigne
                        theShift.aRange.Interior.Color = RGB(0, 255, 255)
                    Case PublicEnums.Availability.NonDispoPermanente
                        theShift.aRange.Interior.Color = RGB(220, 20, 60)
                    Case PublicEnums.Availability.NonDispoTemporaire
                        theShift.aRange.Interior.Color = RGB(219, 112, 147)
                    Case PublicEnums.Availability.SurUtilise
                        theShift.aRange.Interior.Color = RGB(209, 95, 238)
                    Case Else
                End Select
            End If
        Next
    End Sub


    Private Sub fixAvailability(aDoc As String, aMonth As ScheduleMonth, ashift As ScheduleShift, Optional firstDoc As String = "")
        Dim theDate As Date = ashift.aDate
        Dim theShift As Integer = ashift.ShiftType
        Dim theshiftStart As Integer = ashift.ShiftStart
        Dim theshiftStop As Integer = ashift.ShiftStop
        Dim theStartDay As Integer = theDate.Day - 1
        Dim theStopDay As Integer = theDate.Day + 1
        Dim myShift As ScheduleShift
        Dim aDay As ScheduleDay = aMonth.Days(ashift.aDate.Day)
        Dim nonDispoStart As Long
        Dim nonDispoStop As Long
        Dim shftStop As Long
        Dim shftStart As Long
        Dim RecheckCollection As New Collection
        Dim RecheckShift As ScheduleShift

        For x As Integer = ashift.aDate.Day - 1 To ashift.aDate.Day + 1
            If aMonth.Days.Contains(x.ToString()) Then
                aDay = aMonth.Days.Item(x.ToString())
                For Each myShift In aDay.Shifts
   
                    nonDispoStart = ashift.aDate.Ticks + CLng(ashift.ShiftStart) * 600000000 - theRestTime
                    nonDispoStop = ashift.aDate.Ticks + CLng(ashift.ShiftStop) * 600000000 + theRestTime
                    shftStop = myShift.aDate.Ticks + CLng(myShift.ShiftStop) * 600000000
                    shftStart = myShift.aDate.Ticks + CLng(myShift.ShiftStart) * 600000000
                    Dim thedocAvail As scheduleDocAvailable

                    If firstDoc <> "" Then 'do opposite of the top one
                        'then check if this doc is assigned in prevous or next day
                        'if yes redo fixavailability on either or both of those if not leave as is
                        If myShift.DocAvailabilities.Contains(firstDoc) Then
                            thedocAvail = myShift.DocAvailabilities.Item(firstDoc)
                            If thedocAvail.Availability = PublicEnums.Availability.Assigne Then
                                RecheckCollection.Add(myShift)
                            End If
                        End If
                    End If

                    If (shftStart > nonDispoStart And shftStart < nonDispoStop) Or _
                                       (shftStop > nonDispoStart And shftStop < nonDispoStop) Or _
                                       (shftStart > nonDispoStart And shftStop < nonDispoStop) Then

                        thedocAvail = myShift.DocAvailabilities.Item(aDoc)
                        If thedocAvail.Availability <> PublicEnums.Availability.NonDispoPermanente And _
                                thedocAvail.Availability <> PublicEnums.Availability.Assigne Then
                            thedocAvail.Availability = PublicEnums.Availability.NonDispoTemporaire
                            fixlist(myShift)
                        End If

                        If firstDoc <> "" Then 'do opposite of the top one
                            'then check if this doc is assigned in prevous or next day
                            'if yes redo fixavailability on either or both of those if not leave as is
                            If myShift.DocAvailabilities.Contains(firstDoc) Then
                                thedocAvail = myShift.DocAvailabilities.Item(firstDoc)
                                If thedocAvail.Availability <> PublicEnums.Availability.NonDispoPermanente And _
                                thedocAvail.Availability <> PublicEnums.Availability.Assigne Then
                                    thedocAvail.Availability = PublicEnums.Availability.Dispo

                                    fixlist(myShift)

                                End If
                            End If
                        End If

                        
                    End If
                    If theHighlightedDoc <> "" Then
                        HighLightDocAvailSingleCell(myShift, theHighlightedDoc)
                    End If
                Next
            End If
        Next

        If RecheckCollection.Count > 0 Then
            For Each RecheckShift In RecheckCollection
                fixAvailability(firstDoc, aMonth, RecheckShift)
            Next
        End If

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
                    thelist = thelist + theDocAvailable.DocInitial + ","
                    theSetValue = theDocAvailable.DocInitial
                Case Else

            End Select
        Next
        If thelist.Length > 0 Then thelist = Left(thelist, thelist.Length - 1)
        controlledExcelSheet.Unprotect()
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
        theShift.aRange.Locked = False

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

    Private Sub controlledExcelSheet_BeforeDelete() Handles controlledExcelSheet.BeforeDelete

        Globals.ThisAddIn.theControllerCollection.Remove(controlledExcelSheet.Name)

    End Sub

    Private Sub theMonthlyStatsForm_close() Handles theMonthlyStatsForm.FormClosing
        ScheduleDocStatsCollection = Nothing
    End Sub

    Public Sub statsMensuelles()

        'If theMonthlyStatsForm Is Nothing Then
        '    theMonthlyStatsForm = New Form2
        'Else
        '    theMonthlyStatsForm.Dispose()
        '    theMonthlyStatsForm = New Form2
        'End If
        'theMonthlyStatsForm.TopMost = True
        'theMonthlyStatsForm.Show()


        Dim MyTaskPaneView As MonthlyDocStatsTPF
        MyTaskPaneView = New MonthlyDocStatsTPF
        theCustomTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(MyTaskPaneView, "Statistiques")
        theCustomTaskPane.Visible = True
        statsMensuellesUpdate()


    End Sub

    Private Sub statsMensuellesUpdate()
        'pour chaque medecin compter chaque type de shift

        If Not theCustomTaskPane Is Nothing Then
            If theCustomTaskPane.Visible = True Then

                Dim theDocCollection As Collection = ScheduleDoc.LoadAllDocsPerMonth(controlledMonth.Year, controlledMonth.Month)
                Dim aScheduleDoc As ScheduleDoc
                Dim ashift As ScheduleShift
                Dim aDay As ScheduleDay
                Dim aDOcAvail As scheduleDocAvailable
                Dim theScheduleDocStats As ScheduleDocStats
                If ScheduleDocStatsCollection Is Nothing Then
                    ScheduleDocStatsCollection = New Collection
                    For Each aScheduleDoc In theDocCollection
                        theScheduleDocStats = New ScheduleDocStats(aScheduleDoc.Initials, _
                                                                   aScheduleDoc.Shift1, _
                                                                  aScheduleDoc.Shift2, _
                                                                  aScheduleDoc.Shift3, _
                                                                  aScheduleDoc.Shift4, _
                                                                  aScheduleDoc.Shift5)
                        ScheduleDocStatsCollection.Add(theScheduleDocStats, aScheduleDoc.Initials)

                    Next
                Else
                    For Each theScheduleDocStats In ScheduleDocStatsCollection
                        theScheduleDocStats.shift1 = theScheduleDocStats.shift1E
                        theScheduleDocStats.shift2 = theScheduleDocStats.shift2E
                        theScheduleDocStats.shift3 = theScheduleDocStats.shift3E
                        theScheduleDocStats.shift4 = theScheduleDocStats.shift4E
                        theScheduleDocStats.shift5 = theScheduleDocStats.shift5E
                    Next

                End If

                Dim docCount As Integer = 0
                Dim shiftCount As Integer = 0
                For Each theScheduleDocStats In ScheduleDocStatsCollection
                    For Each aDay In controlledMonth.Days
                        shiftCount = 0
                        For Each ashift In aDay.Shifts
                            If ashift.ShiftType > 5 Then Exit For
                            aDOcAvail = ashift.DocAvailabilities(theScheduleDocStats.Initials)
                            If aDOcAvail.Availability = PublicEnums.Availability.Assigne Then
                                Select Case ashift.ShiftType
                                    Case 1
                                        theScheduleDocStats.shift1 = theScheduleDocStats.shift1 - 1
                                    Case 2
                                        theScheduleDocStats.shift2 = theScheduleDocStats.shift2 - 1
                                    Case 3
                                        theScheduleDocStats.shift3 = theScheduleDocStats.shift3 - 1
                                    Case 4
                                        theScheduleDocStats.shift4 = theScheduleDocStats.shift4 - 1
                                    Case 5
                                        theScheduleDocStats.shift5 = theScheduleDocStats.shift5 - 1
                                End Select
                            End If
                            shiftCount = shiftCount + 1
                        Next
                    Next
                    docCount = docCount + 1
                Next

                'Dim bCollection As System.Windows.Forms.Control.ControlCollection = theMonthlyStatsForm.Controls
                'Dim aElementHost As System.Windows.Forms.Integration.ElementHost = bCollection(0)
                'monthlystats = aElementHost.Child
                'monthlystats.loadarray(ScheduleDocStatsCollection)
                Dim theArray As Integer()
                If theHighlightedDoc <> "" Then

                    ReDim Preserve theArray(3)
                    Dim weekCount As Integer = 0
                    Dim firstday As Boolean = True

                    'go through each day of month
                    For Each aDay In controlledMonth.Days

                        'update counter on week change


                        If aDay.theDate.DayOfWeek = 1 And firstday = False Then
                            weekCount = weekCount + 1
                            If weekCount > 3 Then ReDim Preserve theArray(weekCount)
                        End If

                        firstday = False

                        For Each ashift In aDay.Shifts
                            If ashift.ShiftType > 5 Then Exit For
                            aDOcAvail = ashift.DocAvailabilities(theHighlightedDoc)
                            If aDOcAvail.Availability = PublicEnums.Availability.Assigne Then
                                'populate simple array of week counts
                                theArray(weekCount) = theArray(weekCount) + 1
                            End If
                        Next
                    Next
                Else : theArray = Nothing
                End If



                Dim aCollection As System.Windows.Forms.Control.ControlCollection = theCustomTaskPane.Control.Controls
                Dim bElementHost As System.Windows.Forms.Integration.ElementHost = aCollection(0)
                Dim theMonthlyDocStatsTP As MonthlyDocStatsTP = bElementHost.Child
                theMonthlyDocStatsTP.loadarray(ScheduleDocStatsCollection, theArray)
            End If
        End If



    End Sub

    Private Sub SetUpPermNonDispos()
        Dim theSchedulenondispo As New ScheduleNonDispo
        Dim aSchedulenondispo As ScheduleNonDispo
        Dim aCollection As Collection
        Dim aDay As ScheduleDay
        Dim ashift As ScheduleShift
        Dim theScheduledoc As New ScheduleDoc(controlledMonth.Year, controlledMonth.Month)
        Dim docCollection As Collection = controlledMonth.DocList
        Dim ascheduleDoc As ScheduleDoc
        Dim nonDispoStart As Long
        Dim nonDispoStop As Long
        Dim shftStop As Long
        Dim shftStart As Long

        'For Each doc in the total collection of doctors
        For Each ascheduleDoc In docCollection

            'get the unavailability list for one doctor
            aCollection = theSchedulenondispo.GetNonDispoListForDoc(ascheduleDoc.Initials, controlledMonth.Year, controlledMonth.Month)
            If Not IsNothing(aCollection) Then
                'iterate through the doctors list of unavailabilities
                For Each aSchedulenondispo In aCollection
                    Dim stopDay As Integer
                    Dim startday As Integer
                    Select Case aSchedulenondispo.DateStart.Month
                        Case controlledMonth.Month
                            startday = aSchedulenondispo.DateStart.Day
                        Case Is < controlledMonth.Month
                            startday = 1
                    End Select
                    Select Case aSchedulenondispo.DateStop.Month
                        Case controlledMonth.Month
                            stopDay = aSchedulenondispo.DateStop.Day
                        Case Is > controlledMonth.Month
                            stopDay = System.DateTime.DaysInMonth(controlledMonth.Year, controlledMonth.Month)
                    End Select

                    For y As Integer = startday - 1 To stopDay
                        If controlledMonth.Days.Contains(y) Then
                            aDay = controlledMonth.Days.Item(y)
                            For Each ashift In aDay.Shifts
                                nonDispoStart = aSchedulenondispo.DateStart.Ticks + CLng(aSchedulenondispo.TimeStart) * 600000000
                                nonDispoStop = aSchedulenondispo.DateStop.Ticks + CLng(aSchedulenondispo.TimeStop) * 600000000
                                shftStop = ashift.aDate.Ticks + CLng(ashift.ShiftStop) * 600000000
                                shftStart = ashift.aDate.Ticks + CLng(ashift.ShiftStart) * 600000000

                                If (shftStart > nonDispoStart And shftStart < nonDispoStop) Or _
                                    (shftStop > nonDispoStart And shftStop < nonDispoStop) Or _
                                    (shftStart > nonDispoStart And shftStop < nonDispoStop) Then

                                    Dim thedocAvail As scheduleDocAvailable
                                    thedocAvail = ashift.DocAvailabilities.Item(ascheduleDoc.Initials)
                                    'check if doc is assigned and ask to clear (provide some info.. make surutlisé
                                    If thedocAvail.Availability <> PublicEnums.Availability.Assigne Then
                                        thedocAvail.Availability = PublicEnums.Availability.NonDispoPermanente
                                    End If
                                    fixlist(ashift)
                                End If
                            Next
                        End If
                    Next
                Next
            End If
        Next
    End Sub

    Private Sub resetSheet()
        monthloaded = False 'set boolean toggle to false to stop event triggers
        controlledExcelSheet.Unprotect()
        Dim amonthstring As String = monthstrings(aControlledMonth.Month - 1)
        ' Globals.ThisAddIn.Application.ScreenUpdating = False
        controlledExcelSheet.Cells.Clear() 'clear the worksheet
        Dim theDay As ScheduleDay
        Dim row As Integer
        Dim col As Integer = 0

        'get number of shifts
        Dim rowheight1 As Integer = ScheduleShiftType.ActiveShiftTypesCountPerMonth(aControlledMonth.Month, aControlledMonth.Year) + 1
        'assign colwidth as 2
        Dim colwidth1 As Integer = 2


        'populate the top left corner of sheet with year and month strings
        controlledExcelSheet.Range("A1").Value = amonthstring
        controlledExcelSheet.Range("B1").Value = aControlledMonth.Year.ToString()

        'set top left corner of calendar
        Dim theRangeA3 As Excel.Range = controlledExcelSheet.Range("A3")
        Dim theRange As Excel.Range

        'create the month to display in worksheet
        For Each theDay In controlledMonth.Days
            Select Case (theDay.theDate.DayOfWeek)
                Case 0
                    col = 6
                Case Else
                    col = theDay.theDate.DayOfWeek - 1
            End Select
            theRange = theRangeA3.Offset(row * rowheight1, col * colwidth1)

            Dim theRangeForShiftType As Excel.Range
            Dim TheRAngeForDocLists As Excel.Range
            Dim theShift As ScheduleShift

            For Each theShift In theDay.Shifts
                theRangeForShiftType = theRange.Offset(theShift.ShiftType, 0)
                theRangeForShiftType.Value2 = "'" + theShift.Description
                TheRAngeForDocLists = theRange.Offset(theShift.ShiftType, 1)
                theShift.aRange = TheRAngeForDocLists

                fixlist(theShift)
            Next

            theRange.Offset(0, colwidth1 - 1).Value = theDay.theDate.Day
            theRange.Offset(0, colwidth1 - 1).Interior.Color = RGB(160, 160, 160)
            theRange.Offset(0, colwidth1 - 2).Value = daystrings(theDay.theDate.DayOfWeek)
            theRange.Offset(0, colwidth1 - 2).Interior.Color = RGB(160, 160, 160)
            theRange = theRange.Resize(rowheight1, colwidth1)
            addBordersAroundRange(theRange)
            If col = 6 Then row = row + 1
        Next
        SetupAssignedDocs()
        SetUpPermNonDispos()
        'Globals.ThisAddIn.Application.ScreenUpdating = True

        monthloaded = True
        controlledExcelSheet.Protect(DrawingObjects:=True, Contents:=True, Scenarios:= _
        True, AllowFormattingCells:=False, AllowFormattingColumns:=False, _
        AllowFormattingRows:=False, AllowInsertingColumns:=False, AllowInsertingRows _
        :=False, AllowInsertingHyperlinks:=False, AllowDeletingColumns:=False, _
        AllowDeletingRows:=True, AllowSorting:=False, AllowFiltering:=False, _
        AllowUsingPivotTables:=False)
        controlledExcelSheet.EnableSelection = Microsoft.Office.Interop.Excel.XlEnableSelection.xlUnlockedCells

    End Sub

    Public Sub resetSheetExt()
        'clear the sheet
        controlledExcelSheet.Unprotect()
        controlledExcelSheet.Cells.Clear()
        'create a month
        controlledMonth = New ScheduleMonth(controlledMonth.Month, controlledMonth.Year)
        theHighlightedDoc = ""
        Globals.ThisAddIn.theCurrentController = Me
        'Load shift types collection into global
        'controlledShiftTypes = controlledMonth.ShiftTypes
        resetSheet()
    End Sub

    Private Sub SetupAssignedDocs()
        Dim aTest As New scheduleDocAvailable(DateSerial(aControlledMonth.Year, aControlledMonth.Month, 1))
        Dim aCollection As Collection
        Dim theDay2 As ScheduleDay
        Dim theShift2 As ScheduleShift
        Dim theDocAvailble As scheduleDocAvailable
        aCollection = aTest.doesDataExistForThisMonth()
        If Not IsNothing(aCollection) Then
            Dim theAssignedDocs As scheduleDocAvailable
            For Each theAssignedDocs In aCollection
                theDay2 = controlledMonth.Days.Item(theAssignedDocs.Date_.Day)
                If theDay2.Shifts.Contains(theAssignedDocs.ShiftType.ToString()) Then
                    theShift2 = theDay2.Shifts.Item(theAssignedDocs.ShiftType.ToString())
                    theShift2.Doc = theAssignedDocs.DocInitial
                    If theShift2.DocAvailabilities.Contains(theAssignedDocs.DocInitial) Then
                        theDocAvailble = theShift2.DocAvailabilities(theAssignedDocs.DocInitial)
                        theDocAvailble.SetAvailabilityfromDB = PublicEnums.Availability.Assigne
                        theShift2.aRange.Value = theAssignedDocs.DocInitial
                        fixAvailability(theShift2.Doc, controlledMonth, theShift2)
                    Else
                        Windows.MessageBox.Show("Un medecin avec les initialles " _
                                                + theAssignedDocs.DocInitial + " Etait assigné au quart de travail " _
                                                + theShift2.Description.ToString() + " le " + theDay2.theDate.Day.ToString() _
                                                + ", mais le medecin a été retiré de la liste des médecins. Son assignation au quart de travail a été retiré.")
                        Dim aScheduleDocAvailable As New scheduleDocAvailable(" ", PublicEnums.Availability.Assigne, theDay2.theDate, theShift2.ShiftType)
                        aScheduleDocAvailable.DeleteScheduleDataEntry()
                    End If
                End If
            Next
        End If
    End Sub

    Private Sub ClearAvailability()
        Dim aDay As ScheduleDay
        Dim ashift As ScheduleShift
        Dim aDocAvail As scheduleDocAvailable

        For Each aDay In aControlledMonth.Days
            For Each ashift In aDay.Shifts
                For Each aDocAvail In ashift.DocAvailabilities
                    aDocAvail.Availability = Availability.Dispo
                Next
                fixlist(ashift)
            Next
        Next
    End Sub
End Class
