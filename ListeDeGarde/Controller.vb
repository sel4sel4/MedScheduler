Public Class Controller
    Private WithEvents controlledExcelSheet As Excel.Worksheet
    Private controlledMonth As ScheduleMonth
    Private monthloaded As Boolean = False
    Private Const theRestTime As Long = 432000000000

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
        'controlledShiftTypes = controlledMonth.ShiftTypes
        resetSheet()
       

    End Sub

    Private Sub controlledExcelSheet_Change(ByVal Target As Excel.Range) Handles controlledExcelSheet.Change

        If monthloaded = False Then Exit Sub

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
            'statsMensuelles()
        End If

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

    Private Sub controlledExcelSheet_BeforeDelete() Handles controlledExcelSheet.BeforeDelete

        Globals.ThisAddIn.theControllerCollection.Remove(controlledExcelSheet.Name)

    End Sub

    Public Sub statsMensuelles()
        'Dim theShiftTypeCounts As Collection
        'Dim ascheduleshifttype As ScheduleShiftType
        'pour chaque medecin compter chaque type de shift
        Dim theDocCollection As Collection = ScheduleDoc.LoadAllDocsPerMonth(controlledMonth.Year, controlledMonth.Month)
        Dim aScheduleDoc As ScheduleDoc
        Dim ashift As ScheduleShift
        Dim aDay As ScheduleDay
        'Dim StartingRange As Excel.Range
        'StartingRange = controlledExcelSheet.Range("p3")

        'For Each ascheduleshifttype In controlledMonth.ShiftTypes
        '    StartingRange.Offset(-1, ascheduleshifttype.ShiftType).Value = "'" + ascheduleshifttype.Description
        'Next

        Dim aDOcAvail As scheduleDocAvailable

        'For Each aScheduleDoc In theDocCollection
        '    StartingRange.Value = aScheduleDoc.Initials
        '    theShiftTypeCounts = New Collection

        '    For Each ascheduleshifttype In globalShiftTypes.ShiftCollection
        '        StartingRange.Offset(0, ascheduleshifttype.ShiftType).Value = 0
        '    Next


        '    For Each aDay In controlledMonth.Days
        '        For Each ashift In aDay.Shifts
        '            aDOcAvail = ashift.DocAvailabilities(aScheduleDoc.Initials)
        '            If aDOcAvail.Availability = PublicEnums.Availability.Assigne Then
        '                StartingRange.Offset(0, aDOcAvail.ShiftType).Value = StartingRange.Offset(0, aDOcAvail.ShiftType).Value + 1
        '            End If
        '        Next

        '    Next


        '    StartingRange = StartingRange.Offset(1, 0)
        'Next
        Dim anArray As Integer(,)

        ReDim anArray(theDocCollection.Count - 1, controlledMonth.ShiftTypes.Count - 1)
        Dim docCount As Integer = 0
        Dim shiftCount As Integer = 0
        For Each aScheduleDoc In theDocCollection
            For Each aDay In controlledMonth.Days
                shiftCount = 0
                For Each ashift In aDay.Shifts
                    aDOcAvail = ashift.DocAvailabilities(aScheduleDoc.Initials)
                    If aDOcAvail.Availability = PublicEnums.Availability.Assigne Then
                        anArray(docCount, shiftCount) = anArray(docCount, shiftCount) + 1
                    End If
                    shiftCount = shiftCount + 1
                Next
            Next
            docCount = docCount + 1
        Next
        docCount = 0
        Dim aCollection As New Collection
        Dim theScheduleDocStats As ScheduleDocStats
        For Each aScheduleDoc In theDocCollection
            theScheduleDocStats = New ScheduleDocStats(aScheduleDoc.Initials, _
                                                       anArray(docCount, 0), _
                                                       anArray(docCount, 1), _
                                                       anArray(docCount, 2), _
                                                       anArray(docCount, 3), _
                                                       anArray(docCount, 4), _
                                                       anArray(docCount, 5), _
                                                       anArray(docCount, 6), _
                                                       anArray(docCount, 7))
            aCollection.Add(theScheduleDocStats)
            docCount = docCount + 1
        Next






        'noter le medecin sur le WS
        'dans un array de dimension n= types de shifts
        'compter les assignations
        'transferer les donnees sur la page

        'Dim StartingRange As Excel.Range = controlledExcelSheet.Range("p3")
        'StartingRange = StartingRange.Resize(8, 8)
        'StartingRange.Value = anArray

        Dim theFOrm As New Form2
        theFOrm.Show()

        'need to rebuild the taskpane on the basis of the currentlyselected month
        'code below retreives the handle to the UserControl to trigger redraw() public function
        Dim bCollection As System.Windows.Forms.Control.ControlCollection = theFOrm.Controls
        Dim aElementHost As System.Windows.Forms.Integration.ElementHost = bCollection(0)
        Dim aUserControl4 As UserControl4 = aElementHost.Child
        aUserControl4.loadarray(aCollection)

    End Sub

    Private Sub SetUpPermNonDispos()
        Dim theSchedulenondispo As New ScheduleNonDispo
        Dim aSchedulenondispo As ScheduleNonDispo
        Dim aCollection As Collection
        Dim aDay As ScheduleDay
        Dim ashift As ScheduleShift
        Dim theScheduledoc As New ScheduleDoc(controlledMonth.Year, controlledMonth.Month)
        Dim docCollection As Collection = theScheduledoc.DocList
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

                    'start with the day prior to the start of the unavailability (to cover the 0-8 shift)
                    'If controlledMonth.Days.Contains(aSchedulenondispo.DateStart.Day - 1) Then
                    '    aDay = controlledMonth.Days.Item(aSchedulenondispo.DateStart.Day - 1)
                    '    For Each ashift In aDay.Shifts
                    '        If aSchedulenondispo.TimeStart + 1440 < ashift.ShiftStop Then
                    '            Dim thedocAvail As scheduleDocAvailable
                    '            If ashift.DocAvailabilities.Contains(ascheduleDoc.Initials) Then
                    '                thedocAvail = ashift.DocAvailabilities.Item(ascheduleDoc.Initials)
                    '                thedocAvail.Availability = PublicEnums.Availability.NonDispoPermanente
                    '                fixlist(ashift)
                    '            End If
                    '        End If
                    '    Next
                    'End If
                    'FIX: non-dispos spanning more than one month
                    'cycle through the days included in the non dispo

                    For y As Integer = aSchedulenondispo.DateStart.Day - 1 To aSchedulenondispo.DateStop.Day
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

        Dim amonthstring As String = monthstrings(aControlledMonth.Month - 1)

        controlledExcelSheet.Cells.Clear() 'clear the worksheet
        Dim theDay As ScheduleDay
        Dim row As Integer
        Dim col As Integer = 0

        'get number of shifts
        Dim rowheight1 As Integer = controlledMonth.Days(1).shifts.Count + 1
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
            col = CInt(theDay.theDate.DayOfWeek)
            theRange = theRangeA3.Offset(row * rowheight1, col * colwidth1)

            Dim theRangeForShiftType As Excel.Range
            Dim TheRAngeForDocLists As Excel.Range
            Dim theShift As ScheduleShift

            Dim theCounter1 As Integer = 1
            For Each theShift In theDay.Shifts
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

        SetupAssignedDocs()
        SetUpPermNonDispos()
        monthloaded = True

    End Sub

    Public Sub resetSheetExt()
        'clear the sheet
        controlledExcelSheet.Cells.Clear()
        'create a month
        controlledMonth = New ScheduleMonth(controlledMonth.Month, controlledMonth.Year)

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
                    theDocAvailble = theShift2.DocAvailabilities(theAssignedDocs.DocInitial)
                    theDocAvailble.SetAvailabilityfromDB = PublicEnums.Availability.Assigne
                    theShift2.aRange.Value = theAssignedDocs.DocInitial
                    fixAvailability(theShift2.Doc, controlledMonth, theShift2)
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
