﻿Imports System.Diagnostics
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
    Private pShiftypes As Collection
    Private pDocList As Collection

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
    ReadOnly Property ShiftTypes() As Collection
        Get
            Return pShiftypes
        End Get
    End Property
    ReadOnly Property DocList() As Collection
        Get
            Return pDocList
        End Get
    End Property
    Public Sub New(aMonth As Integer, aYear As Integer)
        pShiftypes = ScheduleShiftType.loadShiftTypesFromDBPerMonth(aMonth, aYear)
        pDocList = ScheduleDoc.LoadAllDocsPerMonth(aYear, aMonth)
        Dim theDaysInMonth As Integer = DateTime.DaysInMonth(aYear, aMonth)
        pYear = aYear
        pMonth = aMonth
        pDays = New Collection
        For x = 1 To theDaysInMonth
            Dim theDay As ScheduleDay
            theDay = New ScheduleDay(x, aMonth, aYear, Me)
            pDays.Add(theDay, x.ToString())
        Next
    End Sub

End Class

Public Class ScheduleDay
    Private pDate As DateTime 'uniqueID
    Private pShifts As Collection
    Private pMonth As ScheduleMonth

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
    ReadOnly Property Month() As ScheduleMonth
        Get
            Return pMonth
        End Get
    End Property

    Public Sub New(aDay As Integer, aMonth As Integer, aYear As Integer, ByRef CMonth As ScheduleMonth)
        pDate = New DateTime(aYear, aMonth, aDay)
        pMonth = CMonth
        pShifts = New Collection

        'populate the shift collection by cycling through 
        'the active ScheduleShiftTypes collection
        Dim aShiftType As ScheduleShiftType
        Dim theCounter As Integer = 1
        For Each aShiftType In pMonth.ShiftTypes
            If aShiftType.Active Then
                Dim theShift As New ScheduleShift(aShiftType.ShiftType, _
                                                  pDate, _
                                                  aShiftType.ShiftStart, _
                                                  aShiftType.ShiftStop, _
                                                  aShiftType.Description, _
                                                  Me)
                pShifts.Add(theShift, aShiftType.ShiftType.ToString())
            End If
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
    Private pDay As ScheduleDay

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
                   aDescription As String, _
                   ByRef aDay As ScheduleDay)
        pDate = aDate
        pShiftType = aShiftType
        pShiftStart = aShiftStart
        pShiftStop = aShiftStop
        pStatus = 0 ' for empty
        pDescription = aDescription
        pDay = aDay

        pDocAvailabilities = New Collection
        Dim theScheduleDocAvailable As scheduleDocAvailable
        Dim aScheduleDoc As ScheduleDoc
        Dim theDispo As PublicEnums.Availability
        For Each aScheduleDoc In pDay.Month.DocList
            'conditional code to make doc unavailable if shift is not active for the doc
            Select Case (aShiftType)
                Case 1, 2, 3, 4 'urgence
                    If aScheduleDoc.UrgenceTog = False Then theDispo = Availability.NonDispoPermanente _
                        Else theDispo = Availability.Dispo
                Case 5 'urgence nuit
                    If aScheduleDoc.UrgenceTog = False Or aScheduleDoc.NuitsTog = False Then _
                        theDispo = Availability.NonDispoPermanente Else theDispo = Availability.Dispo

                Case 6 'hospit
                    If aScheduleDoc.HospitTog = False Then theDispo = Availability.NonDispoPermanente _
                        Else theDispo = Availability.Dispo
                Case 7 'soins
                    If aScheduleDoc.SoinsTog = False Then theDispo = Availability.NonDispoPermanente _
                        Else theDispo = Availability.Dispo
                Case Else
                    theDispo = Availability.Dispo
            End Select
            theScheduleDocAvailable = New scheduleDocAvailable(aScheduleDoc.Initials, _
                                                               theDispo, _
                                                               pDate, _
                                                               pShiftType)
            pDocAvailabilities.Add(theScheduleDocAvailable, aScheduleDoc.Initials)
        Next

    End Sub

End Class

Public Class ScheduleShiftType
    Private pShiftStart As T_DBRefTypeI
    Private pShiftStop As T_DBRefTypeI
    Private pShiftType As T_DBRefTypeI
    Private pActive As T_DBRefTypeB
    Private pDescription As T_DBRefTypeS
    Private pVersion As T_DBRefTypeI


    Public Property Version() As Integer
        Get
            Return pVersion.theValue
        End Get
        Set(ByVal value As Integer)
            pVersion.theValue = value
        End Set
    End Property
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

    Public Sub New()
        pShiftStart.theSQLName = SQLShiftStart
        pShiftStop.theSQLName = SQLShiftStop
        pShiftType.theSQLName = SQLShiftType
        pActive.theSQLName = SQLActive
        pDescription.theSQLName = SQLDescription
        pVersion.theSQLName = SQLVersion
    End Sub
    Public Shared Function loadShiftTypesFromDBPerMonth(aMonth As Integer, aYear As Integer) As Collection
        Dim theBuiltSql As New SQLStrBuilder
        Dim theRS As New ADODB.Recordset
        Dim theDBAC As New DBAC
        Dim aShifttype As ScheduleShiftType
        Dim theShiftTypeCollection As Collection
        theShiftTypeCollection = New Collection
        Dim theVersion As Integer : theVersion = ((aYear - 2000) * 100) + aMonth

        'check if a version exists for the month

        With theBuiltSql
            .SQL_Select("*")
            .SQL_From(TABLE_shiftType)
            .SQL_Where(SQLVersion, "=", theVersion)
            .SQL_Order_By(SQLShiftType)
            theDBAC.COpenDB(.SQLStringSelect, theRS)
        End With

        If theRS.RecordCount > 0 Then 'if a version exists load it
            theRS.MoveFirst()
            For x As Integer = 1 To theRS.RecordCount
                aShifttype = New ScheduleShiftType()
                If Not IsDBNull(theRS.Fields(SQLShiftStart).Value) Then _
                    aShifttype.ShiftStart = theRS.Fields(SQLShiftStart).Value
                If Not IsDBNull(theRS.Fields(SQLShiftStop).Value) Then _
                    aShifttype.ShiftStop = theRS.Fields(SQLShiftStop).Value
                If Not IsDBNull(theRS.Fields(SQLShiftType).Value) Then _
                    aShifttype.ShiftType = theRS.Fields(SQLShiftType).Value
                If Not IsDBNull(theRS.Fields(SQLActive).Value) Then _
                    aShifttype.Active = theRS.Fields(SQLActive).Value
                If Not IsDBNull(theRS.Fields(SQLVersion).Value) Then _
                    aShifttype.Version = theRS.Fields(SQLVersion).Value
                If Not IsDBNull(theRS.Fields(SQLDescription).Value) Then _
                    aShifttype.Description = theRS.Fields(SQLDescription).Value

                theShiftTypeCollection.Add(aShifttype)
                theRS.MoveNext()
            Next
        Else 'if no version exists, load the template version (0)
            With theBuiltSql
                .SQLClear()
                .SQL_Select("*")
                .SQL_From(TABLE_shiftType)
                .SQL_Where(SQLVersion, "=", 0)
                .SQL_Order_By(SQLShiftType)
                theDBAC.COpenDB(.SQLStringSelect, theRS)
            End With

            If theRS.RecordCount > 0 Then 'if at least one template shifttype exists load it as a collection

                theRS.MoveFirst()
                For x As Integer = 1 To theRS.RecordCount
                    aShifttype = New ScheduleShiftType()
                    If Not IsDBNull(theRS.Fields(SQLShiftStart).Value) Then _
                        aShifttype.ShiftStart = theRS.Fields(SQLShiftStart).Value
                    If Not IsDBNull(theRS.Fields(SQLShiftStop).Value) Then _
                        aShifttype.ShiftStop = theRS.Fields(SQLShiftStop).Value
                    If Not IsDBNull(theRS.Fields(SQLShiftType).Value) Then _
                        aShifttype.ShiftType = theRS.Fields(SQLShiftType).Value
                    If Not IsDBNull(theRS.Fields(SQLActive).Value) Then _
                        aShifttype.Active = theRS.Fields(SQLActive).Value
                    aShifttype.Version = theVersion 'change version to YYYYMM integer
                    If Not IsDBNull(theRS.Fields(SQLDescription).Value) Then _
                        aShifttype.Description = theRS.Fields(SQLDescription).Value
                    aShifttype.Save() 'save the shifttype version to DB
                    theShiftTypeCollection.Add(aShifttype)
                    theRS.MoveNext()
                Next

            End If
        End If
        Return theShiftTypeCollection
    End Function
    Public Shared Function loadTemplateShiftTypesFromDB() As Collection
        Dim theBuiltSql As New SQLStrBuilder
        Dim theRS As New ADODB.Recordset
        Dim theDBAC As New DBAC
        Dim aShifttype As ScheduleShiftType
        Dim theShiftTypeCollection As Collection
        theShiftTypeCollection = New Collection
        With theBuiltSql
            .SQL_Select("*")
            .SQL_From(TABLE_shiftType)
            .SQL_Where(SQLVersion, "=", 0)
            .SQL_Order_By(SQLShiftType)
            theDBAC.COpenDB(.SQLStringSelect, theRS)
        End With

        If theRS.RecordCount > 0 Then

            theRS.MoveFirst()
            For x As Integer = 1 To theRS.RecordCount
                aShifttype = New ScheduleShiftType()
                If Not IsDBNull(theRS.Fields(SQLShiftStart).Value) Then _
                    aShifttype.ShiftStart = theRS.Fields(SQLShiftStart).Value
                If Not IsDBNull(theRS.Fields(SQLShiftStop).Value) Then _
                    aShifttype.ShiftStop = theRS.Fields(SQLShiftStop).Value
                If Not IsDBNull(theRS.Fields(SQLShiftType).Value) Then _
                    aShifttype.ShiftType = theRS.Fields(SQLShiftType).Value
                If Not IsDBNull(theRS.Fields(SQLActive).Value) Then _
                    aShifttype.Active = theRS.Fields(SQLActive).Value
                If Not IsDBNull(theRS.Fields(SQLDescription).Value) Then _
                    aShifttype.Description = theRS.Fields(SQLDescription).Value

                theShiftTypeCollection.Add(aShifttype)
                theRS.MoveNext()
            Next
        End If
        Return theShiftTypeCollection
    End Function
    Public Sub Copy(TheInstanceToBeCopied As ScheduleShiftType)

        With TheInstanceToBeCopied

            'Me.pCollection = .ShiftCollection
            Me.ShiftStart = .ShiftStart
            Me.ShiftStop = .ShiftStop
            Me.ShiftType = .ShiftType
            Me.Version = .Version
            Me.Active = .Active
            Me.Description = .Description

        End With
    End Sub
    Public Sub Save()
        Dim theBuiltSql As New SQLStrBuilder
        Dim theRS As New ADODB.Recordset
        Dim theDBAC As New DBAC

        With theBuiltSql
            .SQL_Select("*")
            .SQL_From(TABLE_shiftType)
            .SQL_Where(pShiftType.theSQLName, "=", Me.ShiftType)
            .SQL_Where(pVersion.theSQLName, "=", Me.Version)
            theDBAC.COpenDB(.SQLStringSelect, theRS)
        End With

        Dim theCount As Integer = theRS.RecordCount

        Select Case theCount
            Case 0  'if not create a new entry
                With theBuiltSql
                    .SQLClear()
                    .SQL_Insert(TABLE_shiftType)
                    .SQL_Values(pShiftStart.theSQLName, ShiftStart)
                    .SQL_Values(pShiftStop.theSQLName, ShiftStop)
                    .SQL_Values(pVersion.theSQLName, Version)
                    .SQL_Values(pShiftType.theSQLName, ShiftType)
                    .SQL_Values(pActive.theSQLName, Active)
                    .SQL_Values(pDescription.theSQLName, Description)

                    Dim numaffected As Integer
                    theDBAC.CExecuteDB(.SQLStringInsert, numaffected)
                End With
            Case Else
                Debug.WriteLine("there is already an existing instance with this version number ... this is bad")
        End Select
    End Sub
    Public Sub Update()
        Dim theBuiltSql As New SQLStrBuilder
        Dim theRS As New ADODB.Recordset
        Dim theDBAC As New DBAC

        With theBuiltSql
            .SQL_Select("*")
            .SQL_From(TABLE_shiftType)
            .SQL_Where(pShiftType.theSQLName, "=", Me.ShiftType)
            .SQL_Where(pVersion.theSQLName, "=", Me.Version)
            theDBAC.COpenDB(.SQLStringSelect, theRS)
        End With

        Dim theCount As Integer = theRS.RecordCount

        Select Case theCount
            Case 0
                Debug.WriteLine("there is nothing to update ... this is bad")

            Case 1 'if yes update it with the new value
                theRS.Fields(pShiftStart.theSQLName).Value = ShiftStart
                theRS.Fields(pShiftStop.theSQLName).Value = ShiftStop
                theRS.Fields(pVersion.theSQLName).Value = Version
                theRS.Fields(pActive.theSQLName).Value = Active
                theRS.Fields(pShiftType.theSQLName).Value = ShiftType
                theRS.Fields(pDescription.theSQLName).Value = Description
                theRS.ActiveConnection = theDBAC.aConnection
                theRS.UpdateBatch()
                theRS.Close()
            Case Else
                Debug.WriteLine("there is more than one copy of this entry ... this is bad")
        End Select
    End Sub

End Class

Public Class ScheduleDoc
    Private pFirstName As T_DBRefTypeS
    Private pLastName As T_DBRefTypeS
    Private pInitials As T_DBRefTypeS
    Private pActive As T_DBRefTypeB
    Private pVersion As T_DBRefTypeI
    Private pMinShift As T_DBRefTypeI
    Private pMaxShift As T_DBRefTypeI
    Private pUrgenceTog As T_DBRefTypeB
    Private pHospitTog As T_DBRefTypeB
    Private pSoinsTog As T_DBRefTypeB
    Private pNuitsTog As T_DBRefTypeB
    Private pYear As Integer
    Private pMonth As Integer
    'Private Shared pDocList As Collection

    'Public ReadOnly Property DocList() As Collection
    '    Get
    '        Return pDocList
    '    End Get
    'End Property
    Public Property FirstName() As String
        Get
            Return pFirstName.theValue
        End Get
        Set(value As String)
            pFirstName.theValue = value
        End Set
    End Property
    Public Property LastName() As String
        Get
            Return pLastName.theValue
        End Get
        Set(value As String)
            pLastName.theValue = value
        End Set
    End Property
    Public Property Initials() As String
        Get
            Return pInitials.theValue
        End Get
        Set(value As String)
            pInitials.theValue = value
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
    Public Property Version() As Integer
        Get
            Return pVersion.theValue
        End Get
        Set(ByVal value As Integer)
            pVersion.theValue = value
        End Set
    End Property
    Public Property MinShift() As Integer
        Get
            Return pMinShift.theValue
        End Get
        Set(ByVal value As Integer)
            pMinShift.theValue = value
        End Set
    End Property
    Public Property MaxShift() As Integer
        Get
            Return pMaxShift.theValue
        End Get
        Set(ByVal value As Integer)
            pMaxShift.theValue = value
        End Set
    End Property
    Public Property UrgenceTog() As Boolean
        Get
            Return pUrgenceTog.theValue
        End Get
        Set(ByVal value As Boolean)
            pUrgenceTog.theValue = value
        End Set
    End Property
    Public Property HospitTog() As Boolean
        Get
            Return pHospitTog.theValue
        End Get
        Set(ByVal value As Boolean)
            pHospitTog.theValue = value
        End Set
    End Property
    Public Property SoinsTog() As Boolean
        Get
            Return pSoinsTog.theValue
        End Get
        Set(ByVal value As Boolean)
            pSoinsTog.theValue = value
        End Set
    End Property
    Public Property NuitsTog() As Boolean
        Get
            Return pNuitsTog.theValue
        End Get
        Set(ByVal value As Boolean)
            pNuitsTog.theValue = value
        End Set
    End Property
    Public ReadOnly Property FistAndLastName() As String
        Get
            Return FirstName + " " + LastName
        End Get
    End Property
    Public Sub New()
        pFirstName.theSQLName = SQLFirstName
        pLastName.theSQLName = SQLLastName
        pInitials.theSQLName = SQLInitials
        pActive.theSQLName = SQLActive
        pVersion.theSQLName = SQLVersion
        pMinShift.theSQLName = SQLMinShift
        pMaxShift.theSQLName = SQLMaxShift
        pUrgenceTog.theSQLName = SQLUrgenceTog
        pHospitTog.theSQLName = SQLHospitTog
        pSoinsTog.theSQLName = SQLSoinsTog
        pNuitsTog.theSQLName = SQLNuitsTog

        FirstName = "FirstName"
        LastName = "LastName"
        Initials = "Initialles"
        Active = True
        Version = 1
        MinShift = 0
        MaxShift = 99
        UrgenceTog = True
        HospitTog = True
        SoinsTog = True
        NuitsTog = True
    End Sub
    Public Sub Delete()
        Dim theBuiltSql As New SQLStrBuilder
        Dim theRS As New ADODB.Recordset
        Dim theDBAC As New DBAC
        Dim numaffected As Integer
        With theBuiltSql
            .SQL_From(TABLE_Doc)
            .SQL_Where(pFirstName.theSQLName, "=", FirstName)
            .SQL_Where(pLastName.theSQLName, "=", LastName)
            .SQL_Where(pInitials.theSQLName, "=", Initials)
            .SQL_Where(pActive.theSQLName, "=", Active)
            .SQL_Where(pVersion.theSQLName, "=", Version)
            .SQL_Where(pMinShift.theSQLName, "=", MinShift)
            .SQL_Where(pMaxShift.theSQLName, "=", MaxShift)
            .SQL_Where(pUrgenceTog.theSQLName, "=", UrgenceTog)
            .SQL_Where(pHospitTog.theSQLName, "=", HospitTog)
            .SQL_Where(pSoinsTog.theSQLName, "=", SoinsTog)
            .SQL_Where(pNuitsTog.theSQLName, "=", NuitsTog)
            theDBAC.CExecuteDB(.SQLStringDelete, numaffected)
        End With
        If numaffected <> 1 Then
            Debug.WriteLine("there is more than one copy of this entry ... this is bad")
        End If
    End Sub
    Public Sub New(aYear As Integer, aMonth As Integer)
        pFirstName.theSQLName = SQLFirstName
        pLastName.theSQLName = SQLLastName
        pInitials.theSQLName = SQLInitials
        pActive.theSQLName = SQLActive
        pVersion.theSQLName = SQLVersion
        pMinShift.theSQLName = SQLMinShift
        pMaxShift.theSQLName = SQLMaxShift
        pUrgenceTog.theSQLName = SQLUrgenceTog
        pHospitTog.theSQLName = SQLHospitTog
        pSoinsTog.theSQLName = SQLSoinsTog
        pNuitsTog.theSQLName = SQLNuitsTog

        FirstName = "FirstName"
        LastName = "LastName"
        Initials = "Initialles"
        Active = True
        Version = 1
        MinShift = 0
        MaxShift = 99
        UrgenceTog = True
        HospitTog = True
        SoinsTog = True
        NuitsTog = True


        'pYear = aYear
        'pMonth = aMonth

        'If pDocList Is Nothing Then
        '    pDocList = New Collection
        '    LoadAllDocs(aYear, aMonth)
        'End If
    End Sub
    Public Sub New(aFirstName As String, _
                    aLastName As String, _
                    aInitials As String, _
                    aActive As Boolean, _
                    aVersion As Integer, _
                    aMinShift As Integer, _
                    aMaxShift As Integer, _
                    aUrgenceTog As Boolean, _
                    aHospitTog As Boolean, _
                    aSoinsTog As Boolean, _
                    aNuitsTog As Boolean)

        pFirstName.theSQLName = SQLFirstName
        pLastName.theSQLName = SQLLastName
        pInitials.theSQLName = SQLInitials
        pActive.theSQLName = SQLActive
        pVersion.theSQLName = SQLVersion
        pMinShift.theSQLName = SQLMinShift
        pMaxShift.theSQLName = SQLMaxShift
        pUrgenceTog.theSQLName = SQLUrgenceTog
        pHospitTog.theSQLName = SQLHospitTog
        pSoinsTog.theSQLName = SQLSoinsTog
        pNuitsTog.theSQLName = SQLNuitsTog

        FirstName = aFirstName
        LastName = aLastName
        Initials = aInitials
        Active = aActive
        Version = aVersion
        MinShift = aMinShift
        MaxShift = aMaxShift
        UrgenceTog = aUrgenceTog
        HospitTog = aHospitTog
        SoinsTog = aSoinsTog
        NuitsTog = aNuitsTog

    End Sub
    Public Sub save()
        Dim theBuiltSql As New SQLStrBuilder
        Dim theRS As New ADODB.Recordset
        Dim theDBAC As New DBAC

        With theBuiltSql
            .SQL_Select("*")
            .SQL_From(TABLE_Doc)
            .SQL_Where(pInitials.theSQLName, "=", Initials)
            .SQL_Where(pVersion.theSQLName, "=", Version)
            theDBAC.COpenDB(.SQLStringSelect, theRS)
        End With

        Dim theCount As Integer = theRS.RecordCount

        Select Case theCount
            Case 0  'if not create a new entry
                With theBuiltSql
                    .SQLClear()
                    .SQL_Insert(TABLE_Doc)
                    .SQL_Values(pFirstName.theSQLName, FirstName)
                    .SQL_Values(pLastName.theSQLName, LastName)
                    .SQL_Values(pInitials.theSQLName, Initials)
                    .SQL_Values(pActive.theSQLName, Active)
                    .SQL_Values(pVersion.theSQLName, Version)
                    .SQL_Values(pMinShift.theSQLName, MinShift)
                    .SQL_Values(pMaxShift.theSQLName, MaxShift)
                    .SQL_Values(pUrgenceTog.theSQLName, UrgenceTog)
                    .SQL_Values(pHospitTog.theSQLName, HospitTog)
                    .SQL_Values(pSoinsTog.theSQLName, SoinsTog)
                    .SQL_Values(pNuitsTog.theSQLName, NuitsTog)

                    Dim numaffected As Integer
                    theDBAC.CExecuteDB(.SQLStringInsert, numaffected)
                    'Debug.WriteLine(.SQLStringInsert)
                    'Debug.WriteLine("Number of databaseentries" + numaffected.ToString())
                End With

            Case 1 'if yes update it with the new value
                theRS.Fields(pFirstName.theSQLName).Value = FirstName
                theRS.Fields(pLastName.theSQLName).Value = LastName
                theRS.Fields(pInitials.theSQLName).Value = Initials
                theRS.Fields(pActive.theSQLName).Value = Active
                theRS.Fields(pVersion.theSQLName).Value = Version
                theRS.Fields(pMinShift.theSQLName).Value = MinShift
                theRS.Fields(pMaxShift.theSQLName).Value = MaxShift
                theRS.Fields(pUrgenceTog.theSQLName).Value = UrgenceTog
                theRS.Fields(pHospitTog.theSQLName).Value = HospitTog
                theRS.Fields(pSoinsTog.theSQLName).Value = SoinsTog
                theRS.Fields(pNuitsTog.theSQLName).Value = NuitsTog

                theRS.ActiveConnection = theDBAC.aConnection
                theRS.UpdateBatch()
                theRS.Close()
            Case Else
                Debug.WriteLine("there is more than one copy of this entry ... this is bad")
        End Select
    End Sub
    Public Shared Function LoadAllDocsPerMonth(aYear As Integer, aMonth As Integer) As Collection
        Dim theBuiltSql As New SQLStrBuilder
        Dim theRS As New ADODB.Recordset
        Dim theDBAC As New DBAC
        Dim theCurrentMonthDate As Date = DateSerial(aYear, aMonth, 1)
        Dim aCollection As Collection
        aCollection = New Collection
        Dim theVersion As Integer : theVersion = ((aYear - 2000) * 100) + aMonth

        'check if a version exists for the month
        With theBuiltSql
            .SQL_Select("*")
            .SQL_From(TABLE_Doc)
            .SQL_Where(SQLVersion, "=", theVersion)
            .SQL_Order_By(SQLLastName)
            theDBAC.COpenDB(.SQLStringSelect, theRS)
        End With

        If theRS.RecordCount > 0 Then 'if a version exists load it
            theRS.MoveFirst()
            For x As Integer = 1 To theRS.RecordCount
                Dim aScheduleDoc As New ScheduleDoc()
                If Not IsDBNull(theRS.Fields(SQLFirstName).Value) Then _
                aScheduleDoc.FirstName = theRS.Fields(SQLFirstName).Value
                If Not IsDBNull(theRS.Fields(SQLLastName).Value) Then _
                aScheduleDoc.LastName = theRS.Fields(SQLLastName).Value
                If Not IsDBNull(theRS.Fields(SQLInitials).Value) Then _
                aScheduleDoc.Initials = theRS.Fields(SQLInitials).Value
                If Not IsDBNull(theRS.Fields(SQLActive).Value) Then _
                aScheduleDoc.Active = theRS.Fields(SQLActive).Value
                If Not IsDBNull(theRS.Fields(SQLVersion).Value) Then _
                aScheduleDoc.Version = theRS.Fields(SQLVersion).Value
                If Not IsDBNull(theRS.Fields(SQLMinShift).Value) Then _
                    aScheduleDoc.MinShift = theRS.Fields(SQLMinShift).Value
                If Not IsDBNull(theRS.Fields(SQLMaxShift).Value) Then _
                    aScheduleDoc.MaxShift = theRS.Fields(SQLMaxShift).Value
                If Not IsDBNull(theRS.Fields(SQLUrgenceTog).Value) Then _
                    aScheduleDoc.UrgenceTog = theRS.Fields(SQLUrgenceTog).Value
                If Not IsDBNull(theRS.Fields(SQLHospitTog).Value) Then _
                    aScheduleDoc.HospitTog = theRS.Fields(SQLHospitTog).Value
                If Not IsDBNull(theRS.Fields(SQLSoinsTog).Value) Then _
                    aScheduleDoc.SoinsTog = theRS.Fields(SQLSoinsTog).Value
                If Not IsDBNull(theRS.Fields(SQLNuitsTog).Value) Then _
                    aScheduleDoc.NuitsTog = theRS.Fields(SQLNuitsTog).Value

                aCollection.Add(aScheduleDoc, aScheduleDoc.Initials)
                theRS.MoveNext()
            Next
        Else 'if no version exists, load the template version (0)
            With theBuiltSql
                .SQLClear()
                .SQL_Select("*")
                .SQL_From(TABLE_Doc)
                .SQL_Where(SQLVersion, "=", 0)
                .SQL_Order_By(SQLLastName)
                theDBAC.COpenDB(.SQLStringSelect, theRS)
            End With

            If theRS.RecordCount > 0 Then 'if at least one template shifttype exists load it as a collection
                theRS.MoveFirst()
                For x As Integer = 1 To theRS.RecordCount
                    Dim aScheduleDoc As New ScheduleDoc()
                    If Not IsDBNull(theRS.Fields(SQLFirstName).Value) Then _
                    aScheduleDoc.FirstName = theRS.Fields(SQLFirstName).Value
                    If Not IsDBNull(theRS.Fields(SQLLastName).Value) Then _
                    aScheduleDoc.LastName = theRS.Fields(SQLLastName).Value
                    If Not IsDBNull(theRS.Fields(SQLInitials).Value) Then _
                    aScheduleDoc.Initials = theRS.Fields(SQLInitials).Value
                    If Not IsDBNull(theRS.Fields(SQLActive).Value) Then _
                    aScheduleDoc.Active = theRS.Fields(SQLActive).Value
                    aScheduleDoc.Version = theVersion 'change version to YYYYMM integer
                    If Not IsDBNull(theRS.Fields(SQLMinShift).Value) Then _
                        aScheduleDoc.MinShift = theRS.Fields(SQLMinShift).Value
                    If Not IsDBNull(theRS.Fields(SQLMaxShift).Value) Then _
                        aScheduleDoc.MaxShift = theRS.Fields(SQLMaxShift).Value
                    If Not IsDBNull(theRS.Fields(SQLUrgenceTog).Value) Then _
                        aScheduleDoc.UrgenceTog = theRS.Fields(SQLUrgenceTog).Value
                    If Not IsDBNull(theRS.Fields(SQLHospitTog).Value) Then _
                        aScheduleDoc.HospitTog = theRS.Fields(SQLHospitTog).Value
                    If Not IsDBNull(theRS.Fields(SQLSoinsTog).Value) Then _
                        aScheduleDoc.SoinsTog = theRS.Fields(SQLSoinsTog).Value
                    If Not IsDBNull(theRS.Fields(SQLNuitsTog).Value) Then _
                        aScheduleDoc.NuitsTog = theRS.Fields(SQLNuitsTog).Value
                    aScheduleDoc.save()
                    aCollection.Add(aScheduleDoc, aScheduleDoc.Initials)
                    theRS.MoveNext()
                Next
            End If
        End If

        Return aCollection
    End Function
    Public Shared Function LoadTempateDocsFromDB() As Collection
        Dim theBuiltSql As New SQLStrBuilder
        Dim theRS As New ADODB.Recordset
        Dim theDBAC As New DBAC
        Dim aCollection As Collection
        aCollection = New Collection

        'check if a version exists for the month
        With theBuiltSql
            .SQL_Select("*")
            .SQL_From(TABLE_Doc)
            .SQL_Where(SQLVersion, "=", 0)
            .SQL_Order_By(SQLLastName)
            theDBAC.COpenDB(.SQLStringSelect, theRS)
        End With

        If theRS.RecordCount > 0 Then 'if a version exists load it
            theRS.MoveFirst()
            For x As Integer = 1 To theRS.RecordCount
                Dim aScheduleDoc As New ScheduleDoc()
                If Not IsDBNull(theRS.Fields(SQLFirstName).Value) Then _
                aScheduleDoc.FirstName = theRS.Fields(SQLFirstName).Value
                If Not IsDBNull(theRS.Fields(SQLLastName).Value) Then _
                aScheduleDoc.LastName = theRS.Fields(SQLLastName).Value
                If Not IsDBNull(theRS.Fields(SQLInitials).Value) Then _
                aScheduleDoc.Initials = theRS.Fields(SQLInitials).Value
                If Not IsDBNull(theRS.Fields(SQLActive).Value) Then _
                aScheduleDoc.Active = theRS.Fields(SQLActive).Value
                If Not IsDBNull(theRS.Fields(SQLVersion).Value) Then _
                aScheduleDoc.Version = theRS.Fields(SQLVersion).Value
                If Not IsDBNull(theRS.Fields(SQLMinShift).Value) Then _
                    aScheduleDoc.MinShift = theRS.Fields(SQLMinShift).Value
                If Not IsDBNull(theRS.Fields(SQLMaxShift).Value) Then _
                    aScheduleDoc.MaxShift = theRS.Fields(SQLMaxShift).Value
                If Not IsDBNull(theRS.Fields(SQLUrgenceTog).Value) Then _
                    aScheduleDoc.UrgenceTog = theRS.Fields(SQLUrgenceTog).Value
                If Not IsDBNull(theRS.Fields(SQLHospitTog).Value) Then _
                    aScheduleDoc.HospitTog = theRS.Fields(SQLHospitTog).Value
                If Not IsDBNull(theRS.Fields(SQLSoinsTog).Value) Then _
                    aScheduleDoc.SoinsTog = theRS.Fields(SQLSoinsTog).Value
                If Not IsDBNull(theRS.Fields(SQLNuitsTog).Value) Then _
                    aScheduleDoc.NuitsTog = theRS.Fields(SQLNuitsTog).Value

                aCollection.Add(aScheduleDoc, aScheduleDoc.Initials)
                theRS.MoveNext()
            Next
        End If
        Return aCollection
    End Function

End Class
Public Class ScheduleDocStats
    Private pInitials As String
    Private pShift1 As Integer
    Private pShift2 As Integer
    Private pShift3 As Integer
    Private pShift4 As Integer
    Private pShift5 As Integer
    Private pShift6 As Integer
    Private pShift7 As Integer
    Private pShift8 As Integer


    Public Property Initials() As String
        Get
            Return pInitials
        End Get
        Set(value As String)
            pInitials = value
        End Set
    End Property

    Public Property shift1() As Integer
        Get
            Return pShift1
        End Get
        Set(ByVal value As Integer)
            pShift1 = value
        End Set
    End Property
    Public Property shift2() As Integer
        Get
            Return pShift2
        End Get
        Set(ByVal value As Integer)
            pShift2 = value
        End Set
    End Property
    Public Property shift3() As Integer
        Get
            Return pShift3
        End Get
        Set(ByVal value As Integer)
            pShift3 = value
        End Set
    End Property
    Public Property shift4() As Integer
        Get
            Return pShift4
        End Get
        Set(ByVal value As Integer)
            pShift4 = value
        End Set
    End Property
    Public Property shift5() As Integer
        Get
            Return pShift5
        End Get
        Set(ByVal value As Integer)
            pShift5 = value
        End Set
    End Property
    Public Property shift6() As Integer
        Get
            Return pShift6
        End Get
        Set(ByVal value As Integer)
            pShift6 = value
        End Set
    End Property
    Public Property shift7() As Integer
        Get
            Return pShift7
        End Get
        Set(ByVal value As Integer)
            pShift7 = value
        End Set
    End Property
    Public Property shift8() As Integer
        Get
            Return pShift8
        End Get
        Set(ByVal value As Integer)
            pShift8 = value
        End Set
    End Property

    Public Sub New(aInitials As String, aShift1 As Integer, _
                    aShift2 As Integer, _
                     aShift3 As Integer, _
                      aShift4 As Integer, _
                       aShift5 As Integer, _
                        aShift6 As Integer, _
                         aShift7 As Integer, _
                          aShift8 As Integer)
        Initials = aInitials
        shift1 = aShift1
        shift2 = aShift2
        shift3 = aShift3
        shift4 = aShift4
        shift5 = aShift5
        shift6 = aShift6
        shift7 = aShift7
        shift8 = aShift8


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
            If value = PublicEnums.Availability.Assigne Then
                UpdateScheduleDataTable(value)
            ElseIf pAvailability = PublicEnums.Availability.Assigne And _
                value = PublicEnums.Availability.Dispo Then
                DeleteScheduleDataEntry()
            End If
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
        pAvailability = PublicEnums.Availability.Dispo
        pDocInitial.theSQLName = SQLInitials
        pDate.theSQLName = SQLDate
        pShiftType.theSQLName = SQLShiftType

        DocInitial = aDocInitial
        Availability = aAvailability
        Date_ = aDate
        ShiftType = aShiftType
    End Sub
    Public Sub New(aDate As Date)
        pDocInitial.theSQLName = SQLInitials
        pDate.theSQLName = SQLDate
        pShiftType.theSQLName = SQLShiftType
        pAvailability = PublicEnums.Availability.Dispo
        Date_ = aDate
    End Sub
    Public Sub New()
        pDocInitial.theSQLName = SQLInitials
        pDate.theSQLName = SQLDate
        pShiftType.theSQLName = SQLShiftType
        pAvailability = PublicEnums.Availability.Dispo
    End Sub
    Public Sub UpdateScheduleDataTable(theAvail As Integer)
        'check if an entry already exists for this date and shift
        Dim theBuiltSql As New SQLStrBuilder
        Dim theRS As New ADODB.Recordset
        Dim theDBAC As New DBAC

        With theBuiltSql
            .SQL_Select("*")
            .SQL_From(TABLE_ScheduleData)
            .SQL_Where(pDate.theSQLName, "=", Date_)
            .SQL_Where(pShiftType.theSQLName, "=", ShiftType)
            theDBAC.COpenDB(.SQLStringSelect, theRS)
        End With

        Dim theCount As Integer = theRS.RecordCount

        Select Case theCount
            Case 0  'if not create a new entry
                With theBuiltSql
                    .SQLClear()
                    .SQL_Insert(TABLE_ScheduleData)
                    .SQL_Values(pDate.theSQLName, Date_)
                    .SQL_Values(pShiftType.theSQLName, ShiftType)
                    .SQL_Values(pDocInitial.theSQLName, DocInitial)
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
    Public Sub DeleteScheduleDataEntry()
        'check if an entry already exists for this date and shift
        Dim theBuiltSql As New SQLStrBuilder
        Dim theRS As New ADODB.Recordset
        Dim theDBAC As New DBAC

        With theBuiltSql

            .SQL_From(TABLE_ScheduleData)
            .SQL_Where(pDate.theSQLName, "=", Date_)
            .SQL_Where(pShiftType.theSQLName, "=", ShiftType)
            Dim numaffected As Integer
            theDBAC.CExecuteDB(.SQLStringDelete, numaffected)
        End With
    End Sub
    Public Function doesDataExistForThisMonth() As Collection

        Dim theBuiltSql As New SQLStrBuilder
        Dim theRS As New ADODB.Recordset
        Dim theDBAC As New DBAC
        Dim theStartdate As Date = DateSerial(Date_.Year, Date_.Month, 1)
        Dim theStopdate As Date = DateSerial(Date_.Year, Date_.Month + 1, 1)
        With theBuiltSql
            .SQL_Select("*")
            .SQL_From(TABLE_ScheduleData)
            .SQL_Where(pDate.theSQLName, ">=", theStartdate)
            .SQL_Where(pDate.theSQLName, "<", theStopdate)
            theDBAC.COpenDB(.SQLStringSelect, theRS)
        End With

        If theRS.RecordCount > 0 Then
            Dim aScheduleDocAvailable As scheduleDocAvailable
            Dim aCollection As New Collection
            theRS.MoveFirst()
            For x As Integer = 1 To theRS.RecordCount
                aScheduleDocAvailable = New scheduleDocAvailable
                aScheduleDocAvailable.DocInitial = theRS.Fields(Me.pDocInitial.theSQLName).Value
                aScheduleDocAvailable.Date_ = theRS.Fields(Me.pDate.theSQLName).Value
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
    Private pDu As String
    Private pAu As String

    Public ReadOnly Property du() As String
        Get
            Dim myhours As Integer = pTimeStart.theValue / 60
            Dim myminutes As Integer = pTimeStart.theValue - (myhours * 60)
            Dim atime As New DateTime(1, 1, 1, myhours, myminutes, 0)
            Dim astr As String = daystrings(pDateStart.theValue.DayOfWeek) + " le " + pDateStart.theValue.Day.ToString() _
                                 + " " + monthstrings(pDateStart.theValue.Month - 1) _
                                 + " " + pDateStart.theValue.Year.ToString() _
                                  + " à " + Right("0" + atime.Hour.ToString(), 2) _
                                  + ":" + Right("0" + atime.Minute.ToString(), 2)

            Return astr
        End Get
    End Property
    Public ReadOnly Property au() As String
        Get
            Dim myhours As Integer = pTimeStop.theValue / 60
            Dim myminutes As Integer = pTimeStop.theValue - (myhours * 60)
            Dim atime As New DateTime(1, 1, 1, myhours, myminutes, 0)
            Dim astr As String = daystrings(pDateStop.theValue.DayOfWeek) + " le " + pDateStop.theValue.Day.ToString() _
                                 + " " + monthstrings(pDateStop.theValue.Month - 1) _
                                 + " " + pDateStop.theValue.Year.ToString() _
                                  + " à " + Right("0" + atime.Hour.ToString(), 2) _
                                  + ":" + Right("0" + atime.Minute.ToString(), 2)
            Return astr
        End Get
    End Property
    Public Property DocInitial() As String
        Get
            Return pDocInitial.theValue
        End Get
        Set(ByVal value As String)
            pDocInitial.theValue = value
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

        DocInitial = aDocInitial
        DateStart = aDateStart
        DateStop = aDateStop
        TimeStart = aTimeStart
        TimeStop = aTimeStop
        If IsUnique() Then

            Dim theBuiltSql As New SQLStrBuilder
            Dim theDBAC As New DBAC

            With theBuiltSql
                .SQLClear()
                .SQL_Insert(Table_NonDispo)
                .SQL_Values(pDocInitial.theSQLName, DocInitial)
                .SQL_Values(pDateStart.theSQLName, DateStart)
                .SQL_Values(pTimeStart.theSQLName, TimeStart)
                .SQL_Values(pDateStop.theSQLName, DateStop)
                .SQL_Values(pTimeStop.theSQLName, TimeStop)

                Dim numaffected As Integer
                theDBAC.CExecuteDB(.SQLStringInsert, numaffected)
            End With
        End If
    End Sub
    Public Sub New()
        pDocInitial.theSQLName = SQLInitials
        pDateStart.theSQLName = SQLDateStart
        pDateStop.theSQLName = SQLDateStop
        pTimeStart.theSQLName = SQLTimeStart
        pTimeStop.theSQLName = SQLTimeStop
    End Sub
    Private Function IsUnique() As Boolean
        Dim theBuiltSql As New SQLStrBuilder
        Dim theRS As New ADODB.Recordset
        Dim theDBAC As New DBAC
        With theBuiltSql
            .SQLClear()
            .SQL_Select(pDocInitial.theSQLName)
            .SQL_From(Table_NonDispo)
            .SQL_Where(pDocInitial.theSQLName, "=", DocInitial)
            .SQL_Where(pDateStart.theSQLName, "=", DateStart)
            .SQL_Where(pTimeStart.theSQLName, "=", TimeStart)
            .SQL_Where(pDateStop.theSQLName, "=", DateStop)
            .SQL_Where(pTimeStop.theSQLName, "=", TimeStop)
            theDBAC.COpenDB(.SQLStringSelect, theRS)
        End With
        If theRS.RecordCount > 0 Then Return False Else Return True
    End Function
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
            .SQL_Where(pDateStop.theSQLName, ">=", theStartdate)
            .SQL_Where(pDateStart.theSQLName, "<", theStopdate)
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
                If Not IsDBNull(theRS.Fields(pDateStart.theSQLName).Value) Then ascheduleNonDispo.DateStart = theRS.Fields(pDateStart.theSQLName).Value
                If Not IsDBNull(theRS.Fields(pTimeStart.theSQLName).Value) Then ascheduleNonDispo.TimeStart = theRS.Fields(pTimeStart.theSQLName).Value
                If Not IsDBNull(theRS.Fields(pDateStop.theSQLName).Value) Then ascheduleNonDispo.DateStop = theRS.Fields(pDateStop.theSQLName).Value
                If Not IsDBNull(theRS.Fields(pTimeStop.theSQLName).Value) Then ascheduleNonDispo.TimeStop = theRS.Fields(pTimeStop.theSQLName).Value
                aCollection.Add(ascheduleNonDispo, x.ToString())
                theRS.MoveNext()
            Next
            Return aCollection
        Else : Return Nothing
        End If
    End Function
    Public Sub Delete()
        Dim theBuiltSql As New SQLStrBuilder
        Dim theDBAC As New DBAC
        Dim numaffected As Integer
        With theBuiltSql
            .SQL_From(Table_NonDispo)
            .SQL_Where(pDocInitial.theSQLName, "=", DocInitial)
            .SQL_Where(pDateStop.theSQLName, "=", DateStop)
            .SQL_Where(pDateStart.theSQLName, "=", DateStart)
            .SQL_Where(pTimeStop.theSQLName, "=", TimeStop)
            .SQL_Where(pTimeStart.theSQLName, "=", TimeStart)
            theDBAC.CExecuteDB(.SQLStringDelete, numaffected)
        End With
    End Sub

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
        If cnn.State = ADODB.ObjectStateEnum.adStateClosed Then

            If MySettingsGlobal.DataBaseLocation = "" Then LoadDatabaseFileLocation()
            If CONSTFILEADDRESS = "" Then CONSTFILEADDRESS = MySettingsGlobal.DataBaseLocation
            mConnectionString = Provider + "Data Source=" _
            + CONSTFILEADDRESS _
            + ";" & DBpassword
            cnn.ConnectionString = mConnectionString
            cnn.Open()
        End If

        mConnection = cnn
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
                    theValueStr = "'" + theValue + "'"
                Else
                    theValueStr = theValue
                End If

            Case "Date"
                theValueStr = cAccessDateStr(theValue)
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
            Case "Date"
                theValueStr = cAccessDateStr(theValue)
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
                theValueStr = cAccessDateStr(theValue)
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





