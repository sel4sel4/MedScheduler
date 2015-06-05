Public Class Controller
    Private WithEvents controlledExcelSheet As Excel.Worksheet
    Private controlledMonth As SMonth
    Private monthloaded As Boolean = False
    'Private monthlystats As MonthlyStatsC
    'Private WithEvents theMonthlyStatsForm As MonthlyStats
    Private SDocStatsCollection As List(Of SDocStats)
    Private Const theRestTime As Long = 432000000000
    Private theHighlightedDoc As String
    Private theCustomTaskPane As Microsoft.Office.Tools.CustomTaskPane

    Public ReadOnly Property aControlledMonth() As SMonth
        Get
            Return controlledMonth
        End Get
    End Property

    Public ReadOnly Property aControlledExcelSheet() As Excel.Worksheet
        Get
            Return controlledExcelSheet
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
        controlledMonth = New SMonth(aMonth, aYear)

        'Load shift types collection into global
        'controlledShiftTypes = controlledMonth.ShiftTypes
        theHighlightedDoc = ""
        Globals.ThisAddIn.theCurrentController = Me
        resetSheet()


    End Sub
    Public Sub resetSheetExt()
        'clear the sheet
        controlledExcelSheet.Unprotect()
        controlledExcelSheet.Cells.Clear()
        'create a month
        controlledMonth = New SMonth(controlledMonth.Month, controlledMonth.Year)
        theHighlightedDoc = ""
        Globals.ThisAddIn.theCurrentController = Me
        'Load shift types collection into global
        'controlledShiftTypes = controlledMonth.ShiftTypes
        resetSheet()
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
    Public Sub HighLightDocAvailablilities(Initials As String)
        'cycle through the month and highlight everywhere theDoc is available.
        Dim aday As SDay
        Dim aShift As SShift
        Dim adocAvail As SDocAvailable
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
    Public Sub HighLightDocAvailSingleCell(theShift As SShift, Initials As String)
        'cycle through the month and highlight everywhere theDoc is available.

        Dim adocAvail As SDocAvailable
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
    Public Sub fixlist(theShift As SShift)
        Dim theSetValue As String = ""
        Dim theDocAvailable As SDocAvailable
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

    Private Sub controlledExcelSheet_Change(ByVal Target As Excel.Range) Handles controlledExcelSheet.Change

        If monthloaded = False Then Exit Sub
        controlledExcelSheet.Unprotect()
        'System.Diagnostics.Debug.WriteLine("WithEvents: You Changed Cells " + Target.Address + " " + controlledExcelSheet.Name)
        Dim aday As SDay
        Dim aShift As SShift
        Dim adocAvail As SDocAvailable
        Dim anExitNotice As Boolean = False
        Dim firstDoc As String = ""

        For Each aday In controlledMonth.Days
            For Each aShift In aday.Shifts
                If aShift.aRange.Address = Target.Address Then
                    'make current Doc dispo again
                    If Not IsNothing(aShift.Doc) Then
                        If aShift.DocAvailabilities.Exists(Function(xy) xy.DocInitial = aShift.Doc) Then
                            'adocAvail = CType(aShift.DocAvailabilities.Find(Function(xy) xy.DocInitial = aShift.Doc), SDocAvailable)
                            adocAvail = aShift.DocAvailabilities.Find(Function(xy) xy.DocInitial = aShift.Doc)
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
                        If aShift.DocAvailabilities.Exists(Function(xy) xy.DocInitial = CStr(Target.Value)) Then
                            adocAvail = CType(aShift.DocAvailabilities.Find(Function(xy) xy.DocInitial = CStr(Target.Value)), SDocAvailable)
                            adocAvail.Availability = PublicEnums.Availability.Assigne
                            fixAvailability(CStr(Target.Value), controlledMonth, aShift, firstDoc)
                            aShift.Doc = CStr(Target.Value)
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
    Private Sub controlledExcelSheet_BeforeDelete() Handles controlledExcelSheet.BeforeDelete

        'Globals.ThisAddIn.theControllerCollection.Remove(controlledExcelSheet.Name)
        Globals.ThisAddIn.theControllerCollection.RemoveAll(Function(xy) xy.aControlledMonth.Month = Me.aControlledMonth.Month And _
                                                                        xy.aControlledMonth.Year = Me.aControlledMonth.Year)

    End Sub
    'Private Sub theMonthlyStatsForm_close() Handles theMonthlyStatsForm.FormClosing
    '   SDocStatsCollection = Nothing
    'End Sub

    Private Sub fixAvailability(aDoc As String, aMonth As SMonth, ashift As SShift, Optional firstDoc As String = "")
        Dim theDate As Date = ashift.aDate
        Dim theShift As Integer = ashift.ShiftType
        Dim theshiftStart As Integer = ashift.ShiftStart
        Dim theshiftStop As Integer = ashift.ShiftStop
        Dim theStartDay As Integer = theDate.Day - 1
        Dim theStopDay As Integer = theDate.Day + 1
        Dim myShift As SShift
        Dim aDay As SDay = CType(aMonth.Days.Find(Function(xY) ashift.aDate.Day = xY.theDate.Day), SDay)
        Dim nonDispoStart As Long
        Dim nonDispoStop As Long
        Dim shftStop As Long
        Dim shftStart As Long
        Dim RecheckCollection As New List(Of SShift)
        Dim RecheckShift As SShift

        For x As Integer = ashift.aDate.Day - 1 To ashift.aDate.Day + 1
            Dim yx As Integer = x
            If aMonth.Days.Exists(Function(theDay) theDay.theDate.Day = yx) Then
                'If aMonth.Days.Contains(x.ToString()) Then
                aDay = CType(aMonth.Days.Find(Function(theDay) theDay.theDate.Day = yx), SDay)
                For Each myShift In aDay.Shifts

                    nonDispoStart = ashift.aDate.Ticks + CLng(ashift.ShiftStart) * 600000000 - theRestTime
                    nonDispoStop = ashift.aDate.Ticks + CLng(ashift.ShiftStop) * 600000000 + theRestTime
                    shftStop = myShift.aDate.Ticks + CLng(myShift.ShiftStop) * 600000000
                    shftStart = myShift.aDate.Ticks + CLng(myShift.ShiftStart) * 600000000
                    Dim thedocAvail As SDocAvailable

                    If firstDoc <> "" Then 'do opposite of the top one
                        'then check if this doc is assigned in prevous or next day
                        'if yes redo fixavailability on either or both of those if not leave as is
                        If myShift.DocAvailabilities.Exists(Function(xy) xy.DocInitial = firstDoc) Then
                            thedocAvail = CType(myShift.DocAvailabilities.Find(Function(xy) xy.DocInitial = firstDoc), SDocAvailable)
                            If thedocAvail.Availability = PublicEnums.Availability.Assigne Then
                                RecheckCollection.Add(myShift)
                            End If
                        End If
                    End If

                    If (shftStart > nonDispoStart And shftStart < nonDispoStop) Or _
                                       (shftStop > nonDispoStart And shftStop < nonDispoStop) Or _
                                       (shftStart > nonDispoStart And shftStop < nonDispoStop) Then

                        thedocAvail = CType(myShift.DocAvailabilities.Find(Function(xy) xy.DocInitial = aDoc), SDocAvailable)
                        If thedocAvail.Availability <> PublicEnums.Availability.NonDispoPermanente And _
                                thedocAvail.Availability <> PublicEnums.Availability.Assigne Then
                            thedocAvail.Availability = PublicEnums.Availability.NonDispoTemporaire
                            fixlist(myShift)
                        End If

                        If firstDoc <> "" Then 'do opposite of the top one
                            'then check if this doc is assigned in prevous or next day
                            'if yes redo fixavailability on either or both of those if not leave as is
                            If myShift.DocAvailabilities.Exists(Function(xy) xy.DocInitial = firstDoc) Then
                                thedocAvail = CType(myShift.DocAvailabilities.Find(Function(xy) xy.DocInitial = firstDoc), SDocAvailable)
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

    Private Shared Function FindCompiledShifts(xy As SShiftType) As Boolean

        Return xy.Compilation

    End Function

    Private Sub statsMensuellesUpdate()
        'GetMonthlyCountsFromDB()
        GetMonthlyCountsFromInstances()

        Dim thetime3 As DateTime
        Dim thetime4 As DateTime
        thetime3 = DateTime.Now
        If Not theCustomTaskPane Is Nothing Then
            If theCustomTaskPane.Visible = True Then

                Dim theDocCollection As List(Of SDoc) = SDoc.LoadAllDocsPerMonth(controlledMonth.Year, controlledMonth.Month)
                Dim aSDoc As SDoc
                Dim ashift As SShift
                Dim aDay As SDay
                Dim aDOcAvail As SDocAvailable
                Dim theSDocStats As SDocStats
                If SDocStatsCollection Is Nothing Then
                    SDocStatsCollection = New List(Of SDocStats)
                    For Each aSDoc In theDocCollection
                        theSDocStats = New SDocStats(aSDoc.Initials, _
                                                                    aSDoc.Shift1, _
                                                                    aSDoc.Shift2, _
                                                                    aSDoc.Shift3, _
                                                                    aSDoc.Shift4, _
                                                                    aSDoc.Shift5)
                        SDocStatsCollection.Add(theSDocStats)

                    Next
                Else
                    For Each theSDocStats In SDocStatsCollection
                        theSDocStats.shift1 = theSDocStats.shift1E
                        theSDocStats.shift2 = theSDocStats.shift2E
                        theSDocStats.shift3 = theSDocStats.shift3E
                        theSDocStats.shift4 = theSDocStats.shift4E
                        theSDocStats.shift5 = theSDocStats.shift5E
                    Next

                End If

                Dim docCount As Integer = 0
                Dim shiftCount As Integer = 0
                For Each theSDocStats In SDocStatsCollection
                    For Each aDay In controlledMonth.Days
                        shiftCount = 0
                        For Each ashift In aDay.Shifts
                            If ashift.ShiftType > 5 Then Exit For
                            aDOcAvail = CType(ashift.DocAvailabilities.Find(Function(xy) xy.DocInitial = theSDocStats.Initials), SDocAvailable)
                            If aDOcAvail.Availability = PublicEnums.Availability.Assigne Then
                                Select Case ashift.ShiftType
                                    Case 1
                                        theSDocStats.shift1 = theSDocStats.shift1 - 1
                                    Case 2
                                        theSDocStats.shift2 = theSDocStats.shift2 - 1
                                    Case 3
                                        theSDocStats.shift3 = theSDocStats.shift3 - 1
                                    Case 4
                                        theSDocStats.shift4 = theSDocStats.shift4 - 1
                                    Case 5
                                        theSDocStats.shift5 = theSDocStats.shift5 - 1
                                End Select
                            End If
                            shiftCount = shiftCount + 1
                        Next
                    Next
                    docCount = docCount + 1
                Next

                thetime4 = DateTime.Now
                Dim theticks As Long = thetime4.Ticks - thetime3.Ticks
                Dim span As TimeSpan = TimeSpan.FromTicks(theticks)
                System.Diagnostics.Debug.WriteLine(span)

                'Dim bCollection As System.Windows.Forms.Control.ControlCollection = theMonthlyStatsForm.Controls
                'Dim aElementHost As System.Windows.Forms.Integration.ElementHost = bCollection(0)
                'monthlystats = aElementHost.Child
                'monthlystats.loadarray(SDocStatsCollection)
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
                            aDOcAvail = CType(ashift.DocAvailabilities.Find(Function(xy) xy.DocInitial = theHighlightedDoc), SDocAvailable)
                            If aDOcAvail.Availability = PublicEnums.Availability.Assigne Then
                                'populate simple array of week counts
                                theArray(weekCount) = theArray(weekCount) + 1
                            End If
                        Next
                    Next
                Else : theArray = Nothing
                End If



                Dim aCollection As System.Windows.Forms.Control.ControlCollection = theCustomTaskPane.Control.Controls
                Dim bElementHost As System.Windows.Forms.Integration.ElementHost = CType(aCollection(0), Windows.Forms.Integration.ElementHost)
                Dim theMonthlyDocStatsTP As MonthlyDocStatsTP = CType(bElementHost.Child, MonthlyDocStatsTP)
                theMonthlyDocStatsTP.loadarray(SDocStatsCollection, theArray)
            End If
        End If



    End Sub
    Private Sub SetUpPermNonDispos()
        Dim theSNonDispo As New SNonDispo
        Dim aSNonDispo As SNonDispo
        Dim aCollection As List(Of SNonDispo)
        Dim aDay As SDay
        Dim ashift As SShift
        Dim theSDoc As New SDoc(controlledMonth.Year, controlledMonth.Month)
        Dim docCollection As List(Of SDoc) = controlledMonth.DocList
        Dim aSDoc As SDoc
        Dim nonDispoStart As Long
        Dim nonDispoStop As Long
        Dim shftStop As Long
        Dim shftStart As Long

        'For Each doc in the total collection of doctors
        For Each aSDoc In docCollection

            'get the unavailability list for one doctor
            aCollection = theSNonDispo.GetNonDispoListForDoc(aSDoc.Initials, controlledMonth.Year, controlledMonth.Month)
            If Not IsNothing(aCollection) Then
                'iterate through the doctors list of unavailabilities
                For Each aSNonDispo In aCollection
                    Dim stopDay As Integer
                    Dim startday As Integer
                    nonDispoStart = aSNonDispo.DateStart.Ticks + CLng(aSNonDispo.TimeStart) * 600000000
                    nonDispoStop = aSNonDispo.DateStop.Ticks + CLng(aSNonDispo.TimeStop) * 600000000

                    Select Case aSNonDispo.DateStart.Month
                        Case controlledMonth.Month
                            startday = aSNonDispo.DateStart.Day
                        Case Is < controlledMonth.Month
                            startday = 1
                    End Select
                    Select Case aSNonDispo.DateStop.Month
                        Case controlledMonth.Month
                            stopDay = aSNonDispo.DateStop.Day
                        Case Is > controlledMonth.Month
                            stopDay = System.DateTime.DaysInMonth(controlledMonth.Year, controlledMonth.Month)
                    End Select
                    If (controlledMonth.Month = 1 And aSNonDispo.DateStart.Day = 15 And aSDoc.Initials = "DG" And controlledMonth.Year = 2014) Then
                        Dim test As Integer = 1
                    End If


                    For y As Integer = startday - 1 To stopDay
                        Dim yx As Integer = y
                        If controlledMonth.Days.Exists(Function(theDay) theDay.theDate.Day = yx) Then
                            aDay = CType(controlledMonth.Days.Find(Function(theDay) theDay.theDate.Day = yx), SDay)
                            For Each ashift In aDay.Shifts

                                shftStop = ashift.aDate.Ticks + CLng(ashift.ShiftStop) * 600000000
                                shftStart = ashift.aDate.Ticks + CLng(ashift.ShiftStart) * 600000000

                                If ((nonDispoStart > shftStart And nonDispoStart < shftStop) Or _
                                    (nonDispoStop > shftStart And nonDispoStop < shftStop) Or _
                                    (nonDispoStart < shftStart And nonDispoStop > shftStop)) Then


                                    Dim thedocAvail As SDocAvailable
                                    thedocAvail = CType(ashift.DocAvailabilities.Find(Function(xy) xy.DocInitial = aSDoc.Initials), SDocAvailable)
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
        Globals.ThisAddIn.Application.ScreenUpdating = False
        controlledExcelSheet.Cells.Clear() 'clear the worksheet
        Dim theDay As SDay
        Dim row As Integer
        Dim col As Integer = 0

        'get number of shifts
        Dim rowheight1 As Integer = SShiftType.ActiveShiftTypesCountPerMonth(aControlledMonth.Month, aControlledMonth.Year) + 1
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
            Dim theShift As SShift

            For Each theShift In theDay.Shifts
                theRangeForShiftType = theRange.Offset(theShift.AdjustedOrder, 0)
                theRangeForShiftType.Value2 = "'" + theShift.Description
                TheRAngeForDocLists = theRange.Offset(theShift.AdjustedOrder, 1)
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
        Globals.ThisAddIn.Application.ScreenUpdating = True

        monthloaded = True
        controlledExcelSheet.Protect(DrawingObjects:=True, Contents:=True, Scenarios:= _
        True, AllowFormattingCells:=False, AllowFormattingColumns:=False, _
        AllowFormattingRows:=False, AllowInsertingColumns:=False, AllowInsertingRows _
        :=False, AllowInsertingHyperlinks:=False, AllowDeletingColumns:=False, _
        AllowDeletingRows:=True, AllowSorting:=False, AllowFiltering:=False, _
        AllowUsingPivotTables:=False)
        controlledExcelSheet.EnableSelection = Microsoft.Office.Interop.Excel.XlEnableSelection.xlUnlockedCells

    End Sub
    Private Sub SetupAssignedDocs()
        Dim aTest As New SDocAvailable(DateSerial(aControlledMonth.Year, aControlledMonth.Month, 1))
        Dim aCollection As List(Of SDocAvailable)
        Dim theDay2 As SDay
        Dim theShift2 As SShift
        Dim theDocAvailble As SDocAvailable
        aCollection = aTest.doesDataExistForThisMonth()
        If Not IsNothing(aCollection) Then
            Dim theAssignedDocs As SDocAvailable
            For Each theAssignedDocs In aCollection
                theDay2 = CType(controlledMonth.Days.Find(Function(xy) theAssignedDocs.Date_.Day = xy.theDate.Day), SDay)
                If theDay2.Shifts.Exists(Function(xy) xy.ShiftType = theAssignedDocs.ShiftType) Then
                    theShift2 = CType(theDay2.Shifts.Find((Function(xy) xy.ShiftType = theAssignedDocs.ShiftType)), SShift)
                    theShift2.Doc = theAssignedDocs.DocInitial
                    If theShift2.DocAvailabilities.Exists(Function(xy) xy.DocInitial = theAssignedDocs.DocInitial) Then
                        theDocAvailble = CType(theShift2.DocAvailabilities.Find(Function(xy) xy.DocInitial = theAssignedDocs.DocInitial), SDocAvailable)
                        theDocAvailble.SetAvailabilityfromDB = PublicEnums.Availability.Assigne
                        theShift2.aRange.Value = theAssignedDocs.DocInitial
                        fixAvailability(theShift2.Doc, controlledMonth, theShift2)
                    Else
                        Windows.MessageBox.Show("Un medecin avec les initialles " _
                                                + theAssignedDocs.DocInitial + " Etait assigné au quart de travail " _
                                                + theShift2.Description.ToString() + " le " + theDay2.theDate.Day.ToString() _
                                                + ", mais le medecin a été retiré de la liste des médecins. Son assignation au quart de travail a été retiré.")
                        Dim aSDocAvailable As New SDocAvailable(" ", PublicEnums.Availability.Assigne, theDay2.theDate, theShift2.ShiftType)
                        aSDocAvailable.DeleteScheduleDataEntry()
                    End If
                End If
            Next
        End If
    End Sub
    Private Sub ClearAvailability()
        Dim aDay As SDay
        Dim ashift As SShift
        Dim aDocAvail As SDocAvailable

        For Each aDay In aControlledMonth.Days
            For Each ashift In aDay.Shifts
                For Each aDocAvail In ashift.DocAvailabilities
                    aDocAvail.Availability = Availability.Dispo
                Next
                fixlist(ashift)
            Next
        Next
    End Sub

    Private Sub GetMonthlyCountsFromDB()
        'pour chaque medecin compter chaque type de shift
        Dim theTime As DateTime
        Dim theTime2 As DateTime
        theTime = DateTime.Now
        Dim aShiftType As SShiftType
        Dim theBuiltSql As New SQLStrBuilder
        Dim theRS As New ADODB.Recordset
        Dim theDBAC As New DBAC
        Dim theList As New List(Of String)
        Dim theCompiledShifts As List(Of SShiftType)
        Dim theDateB = DateSerial(Globals.ThisAddIn.theCurrentController.aControlledMonth.Year, Globals.ThisAddIn.theCurrentController.aControlledMonth.Month, 1)
        Dim theDateE = DateSerial(Globals.ThisAddIn.theCurrentController.aControlledMonth.Year, Globals.ThisAddIn.theCurrentController.aControlledMonth.Month + 1, 1)

        theCompiledShifts = Globals.ThisAddIn.theCurrentController.aControlledMonth.ShiftTypes.FindAll(AddressOf FindCompiledShifts)

        For Each aShiftType In theCompiledShifts

            With theBuiltSql
                .SQLClear()
                .SQL_Select("count(ID)")
                .SQL_From(TABLE_ScheduleData + " sd2")
                .SQL_Where("sd2." + SQLInitials, "=", "sd." + SQLInitials, "AND", EnumWhereSubClause.EW_None, 1, True)
                .SQL_Where("sd2." + SQLDate, ">=", theDateB)
                .SQL_Where("sd2." + SQLDate, "<", theDateE)
                .SQL_Where("sd2." + SQLShiftType, "=", aShiftType.ShiftType, "AND")

                theList.Add(.SQLStringSelect)
            End With

        Next

        With theBuiltSql
            .SQLClear()
            .SQL_Select("distinct " + SQLInitials)
            For Each astring As String In theList
                .SQL_Select("(" + astring + ")")
            Next

            .SQL_From(TABLE_ScheduleData + " sd")
            .SQL_Order_By(SQLInitials)

            theDBAC.COpenDB(.SQLStringSelect, theRS)
        End With
        theRS.MoveFirst()
        Dim theDocInitials(theRS.RecordCount - 1) As String
        Dim theCounts(theRS.RecordCount - 1, theRS.Fields.Count - 2) As Integer
        Dim theShiftDescs(theRS.Fields.Count - 1) As String

        For y As Integer = 0 To theRS.RecordCount - 1
            theDocInitials(y) = theRS.Fields(0).Value
            For x As Integer = 1 To theRS.Fields.Count - 1
                If y = 0 Then
                    theShiftDescs(x) = theCompiledShifts.Item(x - 1).Description
                End If
                theCounts(y, x - 1) = theRS.Fields(x).Value
            Next
            theRS.MoveNext()
        Next
        theTime2 = DateTime.Now
        Dim theticks As Long = theTime2.Ticks - theTime.Ticks
        Dim span As TimeSpan = TimeSpan.FromTicks(theticks)
        System.Diagnostics.Debug.WriteLine(span)
    End Sub

    Private Sub GetMonthlyCountsFromInstances()
        'pour chaque medecin compter chaque type de shift
        Dim theTime5 As DateTime
        Dim theTime6 As DateTime
        theTime5 = DateTime.Now

        Dim theDateB = DateSerial(Globals.ThisAddIn.theCurrentController.aControlledMonth.Year, Globals.ThisAddIn.theCurrentController.aControlledMonth.Month, 1)
        Dim theDateE = DateSerial(Globals.ThisAddIn.theCurrentController.aControlledMonth.Year, Globals.ThisAddIn.theCurrentController.aControlledMonth.Month + 1, 1)

        Dim theDocList As List(Of SDoc) = Globals.ThisAddIn.theCurrentController.aControlledMonth.DocList
        Dim theCompiledShifts As List(Of SShiftType) = Globals.ThisAddIn.theCurrentController.aControlledMonth.ShiftTypes.FindAll(AddressOf FindCompiledShifts)

        Dim theDocInitials(theDocList.Count - 1) As String
        Dim theCounts(theDocList.Count - 1, theCompiledShifts.Count - 1) As Integer
        Dim theShiftDescs(theCompiledShifts.Count - 1) As String

        Dim x As Integer = 0

        For Each aDoc As SDoc In theDocList
            theDocInitials(x) = aDoc.Initials
            Dim y As Integer = 0
            For Each aShiftType As SShiftType In theCompiledShifts
                If x = 0 Then
                    theShiftDescs(y) = aShiftType.Description
                End If
                For Each aDay As SDay In Globals.ThisAddIn.theCurrentController.aControlledMonth.Days
                    If aDay.Shifts.Exists(Function(xy) xy.ShiftType = aShiftType.ShiftType) Then
                        Dim aShift As SShift = aDay.Shifts.Find(Function(xy) xy.ShiftType = aShiftType.ShiftType)
                        If aShift.Doc = aDoc.Initials Then
                            theCounts(x, y) = theCounts(x, y) + 1
                        End If
                    End If

                Next
                y = y + 1
            Next
            x = x + 1
        Next
        theTime6 = DateTime.Now
        Dim theticks As Long = theTime6.Ticks - theTime5.Ticks
        Dim span As TimeSpan = TimeSpan.FromTicks(theticks)
        System.Diagnostics.Debug.WriteLine(span)
    End Sub
End Class
