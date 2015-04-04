Imports System.Diagnostics
Imports System.Windows.Forms
Imports System.Configuration

Public Class ScheduleYear
    Private pYear As Integer
    Private pMonths As Collection

    ReadOnly Property Year() As Integer
        Get
            Return pYear
        End Get
    End Property

    ReadOnly Property Months() As Collection
        Get
            Return pMonths
        End Get
    End Property

    Public Sub New(aYear As Integer)
        pYear = aYear
        pMonths = New Collection
        For x = 1 To 12
            Dim theMonth As ScheduleMonth
            theMonth = New ScheduleMonth(x, aYear)
            pMonths.Add(theMonth, x.ToString())
        Next
    End Sub

End Class

Public Class ScheduleMonth
    Private pYear As Integer
    Private pMonth As Integer
    Private pDays As Collection

    ReadOnly Property Year() As Integer
        Get
            Return pYear
        End Get
    End Property

    ReadOnly Property Month() As Integer
        Get
            Return pMonth
        End Get
    End Property

    ReadOnly Property Days() As Collection
        Get
            Return pDays
        End Get
    End Property

    Public Sub New(aMonth As Integer, aYear As Integer)
        globalShiftTypes = New ScheduleShiftType
        Dim theDaysInMonth As Integer = DateTime.DaysInMonth(aYear, aMonth)
        pYear = aYear
        pMonth = aMonth
        pDays = New Collection
        For x = 1 To theDaysInMonth
            Dim theDay As ScheduleDay
            theDay = New ScheduleDay(x, aMonth, aYear)
            pDays.Add(theDay, x.ToString())
        Next


    End Sub

End Class

Public Class ScheduleDay
    Private pDate As DateTime 'uniqueID
    Private pShifts As Collection

    ReadOnly Property Shifts() As Collection
        Get
            Return pShifts
        End Get
    End Property

    ReadOnly Property theDate() As DateTime
        Get
            Return pDate
        End Get
    End Property

    Public Sub New(aDay As Integer, aMonth As Integer, aYear As Integer)
        pDate = New DateTime(aYear, aMonth, aDay)
        pShifts = New Collection

        'populate the shift collection by cycling through 
        'the active ScheduleShiftTypes collection
        Dim aShiftType As ScheduleShiftType
        Dim theCounter As Integer = 1
        For Each aShiftType In globalShiftTypes.ShiftCollection
            Dim theShift As New ScheduleShift(aShiftType.ShiftType, _
                                              pDate, _
                                              aShiftType.ShiftStart, _
                                              aShiftType.ShiftStop, _
                                              aShiftType.Description)
            pShifts.Add(theShift, aShiftType.ShiftType.ToString())
        Next

    End Sub

End Class

Public Class ScheduleShift
    Private pShiftStart As Integer
    Private pShiftStop As Integer
    Private pShiftType As Integer
    Private pDescription As String
    Private pDoc As String
    Private pDocAvailabilities As Collection
    Private pDate As DateTime
    Private pStatus As Integer
    Private pRange As Excel.Range
    Private Shared DocList As Collection

    Public Property Doc() As String
        Get
            Return pDoc
        End Get
        Set(ByVal value As String)
            pDoc = value
        End Set
    End Property
    Public Property Status() As Integer
        Get
            Return pStatus
        End Get
        Set(ByVal value As Integer)
            pStatus = value
        End Set
    End Property
    Public ReadOnly Property Description() As String
        Get
            Return pDescription
        End Get
    End Property
    Public Property aRange() As Excel.Range
        Get
            Return pRange
        End Get
        Set(ByVal value As Excel.Range)
            pRange = value
        End Set
    End Property
    Public ReadOnly Property aDate() As Date
        Get
            Return pDate
        End Get
    End Property
    Public ReadOnly Property ShiftType() As Integer
        Get
            Return pShiftType
        End Get
    End Property
    Public ReadOnly Property ShiftStart() As Integer
        Get
            Return pShiftStart
        End Get
    End Property
    Public ReadOnly Property ShiftStop() As Integer
        Get
            Return pShiftStop
        End Get
    End Property
    Public Property DocAvailabilities() As Collection
        Get
            Return pDocAvailabilities
        End Get
        Set(ByVal value As Collection)
            pDocAvailabilities = value
        End Set
    End Property

    Public Sub New(aShiftType As Integer, _
                   aDate As DateTime, _
                   aShiftStart As Integer, _
                   aShiftStop As Integer, _
                   aDescription As String)
        pDate = aDate
        pShiftType = aShiftType
        pShiftStart = aShiftStart
        pShiftStop = aShiftStop
        pStatus = 0 ' for empty
        pDescription = aDescription

        If IsNothing(DocList) Then
            DocList = New Collection
            Dim theScheduleDoc As New ScheduleDoc(pDate.Year, pDate.Month)
            DocList = theScheduleDoc.DocList
        End If
        pDocAvailabilities = New Collection
        Dim theScheduleDocAvailable As scheduleDocAvailable
        Dim aScheduleDoc As ScheduleDoc
        For Each aScheduleDoc In DocList
            theScheduleDocAvailable = New scheduleDocAvailable(aScheduleDoc.Initials, _
                                                               PublicEnums.Availability.Dispo, _
                                                               pDate, _
                                                               pShiftType)
            pDocAvailabilities.Add(theScheduleDocAvailable, aScheduleDoc.Initials)
        Next

    End Sub
    Public Sub addDoc(aDoc As String)
        'modify tally for the doc
        'change docs schedule
    End Sub
    Public Sub clearDoc()
        'modify tally for the doc
        'Change docs schedule
    End Sub


End Class

Public Class ScheduleShiftType
    Private pShiftStart As T_DBRefTypeI
    Private pShiftStop As T_DBRefTypeI
    Private pShiftType As T_DBRefTypeI
    Private pActive As T_DBRefTypeB
    Private pDescription As T_DBRefTypeS
    Shared pCollection As Collection

    Public Property ShiftStart() As Integer
        Get
            Return pShiftStart.theValue
        End Get
        Set(ByVal value As Integer)
            pShiftStart.theValue = value
        End Set
    End Property
    Public Property ShiftStop() As Integer
        Get
            Return pShiftStop.theValue
        End Get
        Set(ByVal value As Integer)
            pShiftStop.theValue = value
        End Set
    End Property
    Public Property ShiftType() As Integer
        Get
            Return pShiftType.theValue
        End Get
        Set(ByVal value As Integer)
            pShiftType.theValue = value
        End Set
    End Property
    Public Property Active() As Boolean
        Get
            Return pActive.theValue
        End Get
        Set(ByVal value As Boolean)
            pActive.theValue = value
        End Set
    End Property
    Public Property Description() As String
        Get
            Return pDescription.theValue
        End Get
        Set(ByVal value As String)
            pDescription.theValue = value
        End Set
    End Property
    Public ReadOnly Property ShiftCollection() As Collection
        Get
            Return pCollection
        End Get
    End Property

    Public Sub New()
        pShiftStart.theSQLName = SQLShiftStart
        pShiftStop.theSQLName = SQLShiftStop
        pShiftType.theSQLName = SQLShiftType
        pActive.theSQLName = SQLActive
        pDescription.theSQLName = SQLDescription

        If pCollection Is Nothing Then
            pCollection = New Collection
            loadActiveShiftTypesFromDB()
        End If

    End Sub
    Public Sub loadActiveShiftTypesFromDB()
        Dim theBuiltSql As New SQLStrBuilder
        Dim theRS As New ADODB.Recordset
        Dim theDBAC As New DBAC

        With theBuiltSql
            .SQL_Select("*")
            .SQL_From(TABLE_shiftType)
            .SQL_Where(pActive.theSQLName, "=", True)
            .SQL_Order_By(pShiftType.theSQLName)

            theDBAC.COpenDB(.SQLStringSelect, theRS)
        End With

        Dim theCount As Integer = theRS.RecordCount
        If theCount > 0 Then
            theRS.MoveFirst()
            For x As Integer = 1 To theCount
                Dim aShifttype As New ScheduleShiftType
                aShifttype.ShiftStart = theRS.Fields(pShiftStart.theSQLName).Value
                aShifttype.ShiftStop = theRS.Fields(pShiftStop.theSQLName).Value
                aShifttype.ShiftType = theRS.Fields(pShiftType.theSQLName).Value
                aShifttype.Active = theRS.Fields(pActive.theSQLName).Value
                aShifttype.Description = theRS.Fields(pDescription.theSQLName).Value

                pCollection.Add(aShifttype)
                theRS.MoveNext()
            Next

        End If

    End Sub

End Class

Public Class ScheduleDoc
    Private pFirstName As T_DBRefTypeS
    Private pLastName As T_DBRefTypeS
    Private pInitials As T_DBRefTypeS
    Private pActive As T_DBRefTypeB
    Private pYear As Integer
    Private pMonth As Integer
    Private pDays As Collection
    Private Shared pDocList As Collection

    Public ReadOnly Property DocList() As Collection
        Get
            Return pDocList
        End Get
    End Property

    Public ReadOnly Property FirstName() As String
        Get
            Return pFirstName.theValue
        End Get
    End Property
    Public ReadOnly Property LastName() As String
        Get
            Return pLastName.theValue
        End Get
    End Property

    Public ReadOnly Property Initials() As String
        Get
            Return pInitials.theValue
        End Get
    End Property


    Public Sub New(aYear As Integer, aMonth As Integer)
        pFirstName.theSQLName = SQLFirstName
        pLastName.theSQLName = SQLLastName
        pInitials.theSQLName = SQLInitials
        pActive.theSQLName = SQLActive
        pYear = aYear
        pMonth = aMonth

        If pDocList Is Nothing Then
            pDocList = New Collection
            LoadAllDocs(aYear, aMonth)
        End If
    End Sub

    Private Sub LoadAllDocs(aYear As Integer, aMonth As Integer)
        Dim theBuiltSql As New SQLStrBuilder
        Dim theRS As New ADODB.Recordset
        Dim theDBAC As New DBAC

        With theBuiltSql
            .SQL_Select("*")
            .SQL_From(TABLE_Doc)
            .SQL_Where(pActive.theSQLName, "=", True)
            .SQL_Order_By(pLastName.theSQLName)

            theDBAC.COpenDB(.SQLStringSelect, theRS)
        End With

        If theRS.RecordCount > 0 Then
            theRS.MoveFirst()
            For x As Integer = 1 To theRS.RecordCount
                Dim aScheduleDoc As New ScheduleDoc(aYear, aMonth)
                aScheduleDoc.pFirstName.theValue = theRS.Fields(pFirstName.theSQLName).Value
                aScheduleDoc.pLastName.theValue = theRS.Fields(pLastName.theSQLName).Value
                aScheduleDoc.pInitials.theValue = theRS.Fields(pInitials.theSQLName).Value
                aScheduleDoc.pActive.theValue = theRS.Fields(pActive.theSQLName).Value
                pDocList.Add(aScheduleDoc)
                theRS.MoveNext()
            Next

        End If
    End Sub

End Class

Public Class scheduleDocAvailable

    Private pDocInitial As T_DBRefTypeS
    Private pAvailability As PublicEnums.Availability
    Private pDate As T_DBRefTypeD
    Private pShiftType As T_DBRefTypeI

    Public Property DocInitial() As String
        Get
            Return pDocInitial.theValue
        End Get
        Set(ByVal value As String)
            pDocInitial.theValue = value
        End Set
    End Property
    Public Property Availability() As PublicEnums.Availability
        Get
            Return pAvailability
        End Get
        Set(ByVal value As PublicEnums.Availability)
            If value = PublicEnums.Availability.Assigne Then UpdateScheduleDataTable(value)
            pAvailability = value
        End Set
    End Property

    Public WriteOnly Property SetAvailabilityfromDB() As PublicEnums.Availability
        Set(ByVal value As PublicEnums.Availability)
            pAvailability = value
        End Set
    End Property

    Public Property Date_() As Date
        Get
            Return pDate.theValue
        End Get
        Set(ByVal value As Date)
            pDate.theValue = value
        End Set
    End Property

    Public Property DateL() As Long
        Get
            Return pDate.theValue.Ticks / kTicksToDays
        End Get
        Set(ByVal value As Long)
            Dim aDateType As DateTime
            aDateType = New DateTime(value * kTicksToDays)
            pDate.theValue = DateSerial(aDateType.Year, aDateType.Month, aDateType.Day)
        End Set
    End Property

    Public Property ShiftType() As Integer
        Get
            Return pShiftType.theValue
        End Get
        Set(ByVal value As Integer)
            pShiftType.theValue = value
        End Set
    End Property

    Public Sub New(aDocInitial As String, _
                   aAvailability As Integer, _
                   aDate As Date, _
                   aShiftType As Integer)

        pDocInitial.theSQLName = SQLInitials
        pDate.theSQLName = SQLDate
        pShiftType.theSQLName = SQLShiftType

        pDocInitial.theValue = aDocInitial
        pAvailability = aAvailability
        pDate.theValue = aDate
        pShiftType.theValue = aShiftType
    End Sub

    Public Sub New(aDate As Date)
        pDocInitial.theSQLName = SQLInitials
        pDate.theSQLName = SQLDate
        pShiftType.theSQLName = SQLShiftType

        pDate.theValue = aDate
    End Sub

    Public Sub New()
        pDocInitial.theSQLName = SQLInitials
        pDate.theSQLName = SQLDate
        pShiftType.theSQLName = SQLShiftType
    End Sub

    Public Sub UpdateScheduleDataTable(theAvail As Integer)
        'check if an entry already exists for this date and shift
        Dim theBuiltSql As New SQLStrBuilder
        Dim theRS As New ADODB.Recordset
        Dim theDBAC As New DBAC

        With theBuiltSql
            .SQL_Select("*")
            .SQL_From(TABLE_ScheduleData)
            .SQL_Where(pDate.theSQLName, "=", DateL)
            .SQL_Where(pShiftType.theSQLName, "=", pShiftType.theValue)
            theDBAC.COpenDB(.SQLStringSelect, theRS)
        End With

        Dim theCount As Integer = theRS.RecordCount

        Select Case theCount
            Case 0  'if not create a new entry
                With theBuiltSql
                    .SQLClear()
                    .SQL_Insert(TABLE_ScheduleData)
                    .SQL_Values(pDate.theSQLName, DateL)
                    .SQL_Values(pShiftType.theSQLName, pShiftType.theValue)
                    .SQL_Values(pDocInitial.theSQLName, pDocInitial.theValue)
                    Dim numaffected As Integer
                    theDBAC.CExecuteDB(.SQLStringInsert, numaffected)
                    'Debug.WriteLine(.SQLStringInsert)
                    'Debug.WriteLine("Number of databaseentries" + numaffected.ToString())
                End With

            Case 1 'if yes update it with the new value
                theRS.Fields(pDocInitial.theSQLName).Value = pDocInitial.theValue
                theRS.ActiveConnection = theDBAC.aConnection
                theRS.UpdateBatch()
                theRS.Close()
            Case Else
                'Debug.WriteLine("there is more than one copy of this entry ... this is bad")


        End Select
    End Sub

    Public Function doesDataExistForThisMonth() As Collection

        Dim theBuiltSql As New SQLStrBuilder
        Dim theRS As New ADODB.Recordset
        Dim theDBAC As New DBAC
        Dim startdate As Date
        Dim stopdate As Date
        startdate = DateSerial(pDate.theValue.Year, pDate.theValue.Month, 1)
        stopdate = DateSerial(pDate.theValue.Year, pDate.theValue.Month + 1, 1)
        With theBuiltSql
            .SQL_Select("*")
            .SQL_From(TABLE_ScheduleData)
            .SQL_Where(pDate.theSQLName, ">=", startdate.Ticks / kTicksToDays)
            .SQL_Where(pDate.theSQLName, "<", stopdate.Ticks / kTicksToDays)
            theDBAC.COpenDB(.SQLStringSelect, theRS)
        End With

        If theRS.RecordCount > 0 Then
            Dim aScheduleDocAvailable As scheduleDocAvailable
            Dim aCollection As New Collection
            theRS.MoveFirst()
            For x As Integer = 1 To theRS.RecordCount
                aScheduleDocAvailable = New scheduleDocAvailable
                aScheduleDocAvailable.DocInitial = theRS.Fields(Me.pDocInitial.theSQLName).Value
                aScheduleDocAvailable.DateL = theRS.Fields(Me.pDate.theSQLName).Value
                aScheduleDocAvailable.ShiftType = theRS.Fields(Me.pShiftType.theSQLName).Value
                aCollection.Add(aScheduleDocAvailable)
                theRS.MoveNext()
            Next
            Return aCollection
        End If
        Return Nothing
    End Function

End Class

Public Class ScheduleNonDispo

    Private pDocInitial As T_DBRefTypeS
    Private pDateStart As T_DBRefTypeD
    Private pDateStop As T_DBRefTypeD
    Private pTimeStart As T_DBRefTypeI
    Private pTimeStop As T_DBRefTypeI

    Public Property DocInitial() As String
        Get
            Return pDocInitial.theValue
        End Get
        Set(ByVal value As String)
            pDocInitial.theValue = value
        End Set
    End Property

    Public Property DateStartL() As Long
        Get
            Return pDateStart.theValue.Ticks / kTicksToDays
        End Get
        Set(ByVal value As Long)
            Dim aDateType As DateTime
            aDateType = New DateTime(value * kTicksToDays)
            pDateStart.theValue = DateSerial(aDateType.Year, aDateType.Month, aDateType.Day)
        End Set
    End Property

    Public Property DateStart() As Date
        Get
            Return pDateStart.theValue
        End Get
        Set(ByVal value As Date)
            pDateStart.theValue = value
        End Set
    End Property

    Public Property DateStopL() As Long
        Get
            Return pDateStop.theValue.Ticks / kTicksToDays
        End Get
        Set(ByVal value As Long)
            Dim aDateType As DateTime
            aDateType = New DateTime(value * kTicksToDays)
            pDateStop.theValue = DateSerial(aDateType.Year, aDateType.Month, aDateType.Day)
        End Set
    End Property

    Public Property DateStop() As Date
        Get
            Return pDateStop.theValue
        End Get
        Set(ByVal value As Date)
            pDateStop.theValue = value
        End Set
    End Property

    Public Property TimeStart() As Integer
        Get
            Return pTimeStart.theValue
        End Get
        Set(ByVal value As Integer)
            pTimeStart.theValue = value
        End Set
    End Property

    Public Property TimeStop() As Integer
        Get
            Return pTimeStop.theValue
        End Get
        Set(ByVal value As Integer)
            pTimeStop.theValue = value
        End Set
    End Property

    Public Sub New(aDocInitial As String, _
                   aDateStart As Date, _
                   aDateStop As Date, _
                   aTimeStart As Integer, _
                   aTimeStop As Integer)

        pDocInitial.theSQLName = SQLInitials
        pDateStart.theSQLName = SQLDateStart
        pDateStop.theSQLName = SQLDateStop
        pTimeStart.theSQLName = SQLTimeStart
        pTimeStop.theSQLName = SQLTimeStop

        pDocInitial.theValue = aDocInitial
        pDateStart.theValue = aDateStart
        pDateStop.theValue = aDateStop
        pTimeStart.theValue = aTimeStart
        pTimeStop.theValue = aTimeStop

        Dim theBuiltSql As New SQLStrBuilder
        Dim theRS As New ADODB.Recordset
        Dim theDBAC As New DBAC

        With theBuiltSql
            .SQLClear()
            .SQL_Insert(Table_NonDispo)
            .SQL_Values(pDocInitial.theSQLName, DocInitial)
            .SQL_Values(pDateStart.theSQLName, DateStartL)
            .SQL_Values(pTimeStart.theSQLName, TimeStart)
            .SQL_Values(pDateStop.theSQLName, DateStopL)
            .SQL_Values(pTimeStop.theSQLName, TimeStop)

            Dim numaffected As Integer
            theDBAC.CExecuteDB(.SQLStringInsert, numaffected)
        End With

    End Sub
    Public Sub New()

        pDocInitial.theSQLName = SQLInitials
        pDateStart.theSQLName = SQLDateStart
        pDateStop.theSQLName = SQLDateStop
        pTimeStart.theSQLName = SQLTimeStart
        pTimeStop.theSQLName = SQLTimeStop

    End Sub

    Public Function GetNonDispoListForDoc(aDocInitials As String, aYear As Integer, aMonth As Integer) As Collection
        Dim theBuiltSql As New SQLStrBuilder
        Dim theRS As New ADODB.Recordset
        Dim theDBAC As New DBAC
        Dim theStartdate As Date = DateSerial(aYear, aMonth, 1)
        Dim theStopdate As Date = DateSerial(aYear, aMonth + 1, 1)
        With theBuiltSql
            .SQL_Select("*")
            .SQL_From(Table_NonDispo)
            .SQL_Where(pDocInitial.theSQLName, "=", aDocInitials)
            .SQL_Where(pDateStop.theSQLName, ">=", theStartdate.Ticks / kTicksToDays)
            .SQL_Where(pDateStart.theSQLName, "<", theStopdate.Ticks / kTicksToDays)
            .SQL_Order_By(pDateStart.theSQLName)
            .SQL_Order_By(pTimeStart.theSQLName)
            theDBAC.COpenDB(.SQLStringSelect, theRS)
        End With
        Dim ascheduleNonDispo As ScheduleNonDispo
        Dim theCount As Integer = theRS.RecordCount
        If theCount > 0 Then
            Dim aCollection As New Collection
            theRS.MoveFirst()
            For x As Integer = 1 To theCount
                ascheduleNonDispo = New ScheduleNonDispo
                If Not IsDBNull(theRS.Fields(pDocInitial.theSQLName).Value) Then ascheduleNonDispo.DocInitial = theRS.Fields(pDocInitial.theSQLName).Value
                If Not IsDBNull(theRS.Fields(pDateStart.theSQLName).Value) Then ascheduleNonDispo.DateStartL = theRS.Fields(pDateStart.theSQLName).Value
                If Not IsDBNull(theRS.Fields(pTimeStart.theSQLName).Value) Then ascheduleNonDispo.TimeStart = theRS.Fields(pTimeStart.theSQLName).Value
                If Not IsDBNull(theRS.Fields(pDateStop.theSQLName).Value) Then ascheduleNonDispo.DateStopL = theRS.Fields(pDateStop.theSQLName).Value
                If Not IsDBNull(theRS.Fields(pTimeStop.theSQLName).Value) Then ascheduleNonDispo.TimeStop = theRS.Fields(pTimeStop.theSQLName).Value
                aCollection.Add(ascheduleNonDispo)
                theRS.MoveNext()
            Next
            Return aCollection
        Else : Return Nothing
        End If
    End Function

    


End Class

Public Class DBAC

    'Public cnn As New ADODB.Connection  'Connection object definition
    'Public rs As New ADODB.Recordset    'recordset object definition

    Const Provider = "Provider=Microsoft.ACE.OLEDB.12.0;"
    Const DBpassword = "Jet OLEDB:Database Password=plasma;"

    Private theConnectionState As Long
    Private mConnectionString As String
    Private mConnection As ADODB.Connection
    Public ReadOnly Property aConnection() As ADODB.Connection
        Get
            Return mConnection
        End Get
    End Property


    Public Sub New()
        On Error GoTo errhandler
        mConnection = New ADODB.Connection
        mConnectionString = Provider + "Data Source=" _
            + CONSTFILEADDRESS _
            + ";" & DBpassword
        mConnection.ConnectionString = mConnectionString
        mConnection.Open()
        On Error GoTo 0
        Exit Sub
errhandler:
        MsgBox("An error occurred during initial connection to DB: " + _
               CStr(Err.Number) + "  :  " + _
               CStr(Err.Description) + _
               "  Most likely cause is database is not where it is " _
               + "supposed to be or there are coonnection issues to the N: Drive")

        'add code to select current location for the database !!FEATURE!!

    End Sub


    Public Sub COpenDB(theSQLstr As String, theRS As ADODB.Recordset)

        'if myReadonly is true file is open as readonly, otherwise it is modifiable

        'Current DB address is hardcoded, there is 
        'code in ThisWorkbook to allow DB selection at workbook open
        'DB address can be changed by simply changing the FileAddress assignment below

        If theRS.State = ADODB.ObjectStateEnum.adStateOpen Then theRS.Close()
        theRS.CursorLocation = ADODB.CursorLocationEnum.adUseClient

        On Error GoTo errhandler
        'Debug.Print(theSQLstr)
        theRS.Open(theSQLstr, _
                   mConnection, _
                   ADODB.CursorTypeEnum.adOpenKeyset, _
                   ADODB.LockTypeEnum.adLockOptimistic)
        theRS.ActiveConnection = Nothing

        On Error GoTo 0
        Exit Sub 'to not run errhandler for nothing
errhandler:

        MsgBox("An error occurred during connection to DB: " + _
               CStr(Err.Number) & "  :  " + _
               CStr(Err.Description) + _
               "   SQL TEXT:   " + _
               theSQLstr)

    End Sub


    Public Sub CExecuteDB(theSQLstr As String, numAffected As Long)

        On Error GoTo errhandler
        mConnection.Execute(theSQLstr, numAffected)
        'StoreToAuditFile theSQLstr
        On Error GoTo 0

        Exit Sub 'to not run errhandler for nothing


errhandler:

        'Dim theError As String
        'theError = CStr(Err.Number) & "  :  " & CStr(Err.Description) & "   SQL TEXT:   " & theSQLstr
        'theError = Replace(theError, "'", "''")
        'MsgBox("An error occurred during connection to DB: " & theError)

        ''LogError theError

    End Sub

    Protected Overrides Sub finalize()

        mConnection.Close()
        mConnection = Nothing

    End Sub


    '    Private Sub StoreToAuditFile(theSQLstr As String)

    '        Dim theBuiltSql As New SQLStrBuilder
    '        Dim numaffected2
    '        theSQLstr = Replace(theSQLstr, "'", "''")
    '        With theBuiltSql

    '            .SQL_Insert(AUDITTABLE)
    '            .SQL_Values("TimeDateStamp", TimeToMillisecond)

    '            .SQL_Values("UserName", theUserName)
    '            .SQL_Values("Query", theSQLstr)

    '            On Error GoTo errhandler
    '            Dim theSQLString As String
    '            theSQLString = .SQLStringInsert
    '            mConnection.Execute(theSQLString, numaffected2)
    '        End With

    '        Exit Sub

    'errhandler:

    '        Dim theError As String
    '        theError = CStr(Err.Number) & "  :  " & CStr(Err.Description) & "   SQL TEXT:   " & theSQLString
    '        theError = Replace(theError, "'", "''")
    '        MsgBox("An error occurred during connection to DB: " & theError)

    '        LogError(theError)

    '    End Sub



    'Private Sub LogError(theError As String)

    '    Dim theBuiltSql As New SQLStrBuilder
    '    Dim numaffected2
    '    theError = Replace(theError, "'", "''")
    '    With theBuiltSql

    '        .SQL_Insert(AUDITTABLE)
    '        .SQL_Values("TimeDateStamp", TimeToMillisecond)

    '        .SQL_Values("UserName", "ERRORLOG")
    '        .SQL_Values("Query", theError)


    '        mConnection.Execute.SQLStringInsert, numaffected2
    '    End With

    'End Sub


    'Public Sub GetSchema(theREC As ADODB.Recordset)

    '    theREC = mConnection.OpenSchema(adSchemaColumns)
    '    ' theREC.ActiveConnection = Nothing

    'End Sub

End Class

Public Class SQLStrBuilder

    Private theSelect As String, theSelectCounter As Integer
    Private theFrom As String, theFromCounter As Integer
    Private theWhere As String, theWhereCounter As Integer
    Private theGroupBy As String, theGroupByCounter As Integer
    Private theOrderBy As String, theOrderByCounter As Integer
    Private theSet As String, theSetCounter As Integer
    Private theUpdate As String, theUpdateCounter As Integer
    Private theInsert As String, theInsertCounter As Integer
    Private theValueS As String, theValuesCounter As Integer
    Private theInto As String
    Private theStrSQL As String

    ReadOnly Property SQLStringSelect() As String
        Get
            If theSelectCounter < 1 Or theFromCounter < 1 Then Return ""
            theStrSQL = theSelect & theFrom & theWhere & theGroupBy & theOrderBy
            Return theStrSQL
        End Get
    End Property

    ReadOnly Property SQLStringUpdate() As String
        Get
            If theUpdateCounter < 1 Or theSetCounter < 1 Then Return ""
            theStrSQL = theUpdate & theSet & theWhere & ";"
            Return theStrSQL
        End Get
    End Property

    ReadOnly Property SQLStringInsert() As String
        Get
            If theInsertCounter < 1 Or theValuesCounter < 1 Then Return ""
            theStrSQL = theInsert & theInto & ") " & theValueS & ");"
            Return theStrSQL
        End Get
    End Property

    ReadOnly Property SQLStringDelete() As String
        Get
            theStrSQL = "DELETE" & theFrom & theWhere
            Return theStrSQL
        End Get
    End Property

    Public Sub New()

        theStrSQL = ""
        theSelect = ""
        theFrom = ""
        theWhere = ""
        theGroupBy = ""
        theOrderBy = ""
        theSet = ""
        theUpdate = ""
        theInsert = ""
        theValueS = ""
        theInto = ""

        theSelectCounter = 0
        theFromCounter = 0
        theWhereCounter = 0
        theGroupByCounter = 0
        theOrderByCounter = 0
        theSetCounter = 0
        theUpdateCounter = 0
        theInsertCounter = 0
        theValuesCounter = 0

    End Sub

    Public Sub SQL_Select(theColumnName As String)

        If theSelectCounter = 0 Then
            theSelect = "SELECT " & theColumnName
            theSelectCounter = theSelectCounter + 1
        Else
            theSelect = theSelect & ", " & theColumnName
            theSelectCounter = theSelectCounter + 1
        End If

    End Sub

    Public Sub SQL_From(theTableName As String)

        If theFromCounter = 0 Then
            theFrom = " FROM " & theTableName
            theFromCounter = theFromCounter + 1
        Else
            theFrom = theFrom & ", " & theTableName
            theFromCounter = theFromCounter + 1
        End If

    End Sub

    Public Sub SQL_Where(theColumnName As String, _
                            theCondition As String, _
                            theValue As Object, _
                            Optional theOperator As String = "AND", _
                            Optional theSubclause As EnumWhereSubClause = EnumWhereSubClause.EW_None, _
                            Optional theParenthesesCount As Integer = 1, _
                            Optional isFieldName As Boolean = False)

        Dim theValueStr As String
        Dim theSubClauseStr_Begin As String
        Dim theSubClauseStr_End As String
        Dim theCounter As Integer

        Select Case TypeName(theValue)
            Case "String"
                If isFieldName = False Then
                    theValueStr = "'" & theValue & "'"
                Else
                    theValueStr = theValue
                End If
            Case "Date"
                If isFieldName = False Then
                    theValueStr = "#" & theValue & "#"
                Else
                    theValueStr = theValue
                End If
            Case "Boolean"
                If theValue = True Then
                    theValueStr = "true"
                Else
                    theValueStr = "false"
                End If
            Case Else
                theValueStr = CStr(theValue)
        End Select
        theSubClauseStr_Begin = ""
        theSubClauseStr_End = ""
        Select Case theSubclause
            Case EnumWhereSubClause.EW_None
                theSubClauseStr_Begin = ""
                theSubClauseStr_End = ""
            Case EnumWhereSubClause.EW_begin
                For theCounter = 1 To theParenthesesCount
                    theSubClauseStr_Begin = theSubClauseStr_Begin & "("
                Next

            Case EnumWhereSubClause.EW_end
                For theCounter = 1 To theParenthesesCount
                    theSubClauseStr_End = theSubClauseStr_End & ")"
                Next
        End Select

        If theWhereCounter = 0 Then
            theWhere = " WHERE " & theSubClauseStr_Begin & theColumnName _
                & theCondition & theValueStr & theSubClauseStr_End
            theWhereCounter = theWhereCounter + 1
        Else
            theWhere = theWhere & " " & theOperator & " " & theSubClauseStr_Begin & _
                theColumnName & " " & theCondition & " " & theValueStr & theSubClauseStr_End
            theWhereCounter = theWhereCounter + 1
        End If

    End Sub

    Public Sub SQL_Where_in(theColumnName As String, theItems As String)

        If theWhereCounter = 0 Then
            theWhere = " WHERE " & theColumnName & theItems
            theWhereCounter = theWhereCounter + 1
        Else
            theWhere = theWhere & " AND " & theColumnName & theItems
            theWhereCounter = theWhereCounter + 1
        End If

    End Sub

    Public Sub SQL_Group_By(theColumnName As String)

        If theGroupByCounter = 0 Then
            theGroupBy = " GROUP BY " & theColumnName
            theGroupByCounter = theGroupByCounter + 1
        Else
            theGroupBy = theGroupBy & ", " & theColumnName
            theGroupByCounter = theGroupByCounter + 1
        End If

    End Sub

    Public Sub SQL_Order_By(theColumnName As String, Optional SortOrder As String = "ASC")

        If theOrderByCounter = 0 Then
            theOrderBy = " ORDER BY " & theColumnName & " " & SortOrder
            theOrderByCounter = theOrderByCounter + 1
        Else
            theOrderBy = theOrderBy & ", " & theColumnName & " " & SortOrder
            theOrderByCounter = theOrderByCounter + 1
        End If

    End Sub

    Public Sub SQL_Set(theColumnName As String, theValue As Object)

        Dim theValueStr As String
        Select Case TypeName(theValue)
            Case "String"
                theValueStr = "'" & theValue & "'"
            Case Else
                theValueStr = CStr(theValue)
        End Select

        If theSetCounter = 0 Then
            theSet = " SET " & theColumnName & "=" & theValueStr
            theSetCounter = theSetCounter + 1
        Else
            theSet = theSet & ", " & theColumnName & "=" & theValueStr
            theSetCounter = theSetCounter + 1
        End If

    End Sub

    Public Sub SQL_Update(theTableName As String)

        theUpdate = "UPDATE " & theTableName
        theUpdateCounter = theUpdateCounter + 1

    End Sub

    Public Sub SQL_Insert(theTableName As String)

        theInsert = "INSERT INTO " & theTableName
        theInsertCounter = theInsertCounter + 1

    End Sub

    Public Sub SQL_Values(theColumnName As String, theValue As Object)

        Dim theValueStr As String
        Select Case TypeName(theValue)
            Case "String"
                theValueStr = "'" & theValue & "'"
            Case "Date"
                theValueStr = "#" & theValue & "#"
            Case Else
                theValueStr = CStr(theValue)
        End Select

        If theValuesCounter = 0 Then
            theInto = " (" & theColumnName
            theValueS = "VALUES (" & theValueStr
            theValuesCounter = theValuesCounter + 1
        Else
            theInto = theInto & ", " & theColumnName
            theValueS = theValueS & ", " & theValueStr
            theValuesCounter = theValuesCounter + 1
        End If

    End Sub

    Public Sub SQLClear()

        theStrSQL = ""
        theSelect = ""
        theFrom = ""
        theWhere = ""
        theGroupBy = ""
        theOrderBy = ""
        theSet = ""
        theUpdate = ""
        theInsert = ""
        theValueS = ""
        theInto = ""

        theSelectCounter = 0
        theFromCounter = 0
        theWhereCounter = 0
        theGroupByCounter = 0
        theOrderByCounter = 0
        theSetCounter = 0
        theUpdateCounter = 0
        theInsertCounter = 0
        theValuesCounter = 0

    End Sub


End Class





