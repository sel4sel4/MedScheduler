Imports System.Diagnostics
Imports System.Windows.Forms
Imports System.Configuration

Public Class SYear
    Private pYear As Integer
    Private pMonths As List(Of SMonth)

    ReadOnly Property Year() As Integer
        Get
            Return pYear
        End Get
    End Property

    ReadOnly Property Months() As List(Of SMonth)
        Get
            Return pMonths
        End Get
    End Property

    Public Sub New(aYear As Integer)
        pYear = aYear
        pMonths = New List(Of SMonth)
        For x = 1 To 12
            Dim theMonth As SMonth
            theMonth = New SMonth(x, aYear)
            pMonths.Add(theMonth)
        Next
    End Sub

End Class

Public Class SMonth
    Private pYear As Integer
    Private pMonth As Integer
    Private pDays As List(Of SDay)
    Private pShiftypes As List(Of SShiftType)
    Private pDocList As List(Of SDoc)

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
    ReadOnly Property Days() As List(Of SDay)
        Get
            Return pDays
        End Get
    End Property
    ReadOnly Property ShiftTypes() As List(Of SShiftType)
        Get
            Return pShiftypes
        End Get
    End Property
    ReadOnly Property DocList() As List(Of SDoc)
        Get
            Return pDocList
        End Get
    End Property
    Public Sub New(aMonth As Integer, aYear As Integer)
        pShiftypes = SShiftType.loadShiftTypesFromDBPerMonth(aMonth, aYear)
        pDocList = SDoc.LoadAllDocsPerMonth(aYear, aMonth)
        Dim theDaysInMonth As Integer = DateTime.DaysInMonth(aYear, aMonth)
        pYear = aYear
        pMonth = aMonth
        pDays = New List(Of SDay)
        For x = 1 To theDaysInMonth
            Dim theDay As SDay
            theDay = New SDay(x, aMonth, aYear, Me)
            pDays.Add(theDay)
        Next
    End Sub

End Class

Public Class SDay
    Private pDate As DateTime 'uniqueID
    Private pShifts As Collection
    Private pMonth As SMonth

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
    ReadOnly Property Month() As SMonth
        Get
            Return pMonth
        End Get
    End Property

    Public Sub New(aDay As Integer, aMonth As Integer, aYear As Integer, ByRef CMonth As SMonth)
        pDate = New DateTime(aYear, aMonth, aDay)
        pMonth = CMonth
        pShifts = New Collection
        Dim addShift As Boolean = False
        'populate the shift collection by cycling through 
        'the active SShiftTypes collection
        Dim aShiftType As SShiftType
        Dim theCounter As Integer = 1
        For Each aShiftType In pMonth.ShiftTypes
            If aShiftType.Active Then
                Select Case pDate.DayOfWeek
                    Case DayOfWeek.Monday
                        If aShiftType.Lundi = True Then addShift = True Else addShift = False
                    Case DayOfWeek.Tuesday
                        If aShiftType.Mardi = True Then addShift = True Else addShift = False
                    Case DayOfWeek.Wednesday
                        If aShiftType.Mercredi = True Then addShift = True Else addShift = False
                    Case DayOfWeek.Thursday
                        If aShiftType.Jeudi = True Then addShift = True Else addShift = False
                    Case DayOfWeek.Friday
                        If aShiftType.Vendredi = True Then addShift = True Else addShift = False
                    Case DayOfWeek.Saturday
                        If aShiftType.Samedi = True Then addShift = True Else addShift = False
                    Case DayOfWeek.Sunday
                        If aShiftType.Dimanche = True Then addShift = True Else addShift = False
                End Select
                If addShift = True Then
                    Dim theShift As New SShift(aShiftType.ShiftType, _
                                                      pDate, _
                                                      aShiftType.ShiftStart, _
                                                      aShiftType.ShiftStop, _
                                                      aShiftType.Description, _
                                                      Me)

                    pShifts.Add(theShift, aShiftType.ShiftType.ToString())
                End If
            End If
        Next

    End Sub

End Class

Public Class SShift
    Private pShiftStart As Integer
    Private pShiftStop As Integer
    Private pShiftType As Integer
    Private pDescription As String
    Private pDoc As String
    Private pDocAvailabilities As Collection
    Private pDate As DateTime
    Private pStatus As Integer
    Private pRange As Excel.Range
    Private pDay As SDay

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
                   ByRef aDay As SDay)
        pDate = aDate
        pShiftType = aShiftType
        pShiftStart = aShiftStart
        pShiftStop = aShiftStop
        pStatus = 0 ' for empty
        pDescription = aDescription
        pDay = aDay

        pDocAvailabilities = New Collection
        Dim theSDocAvailable As SDocAvailable
        Dim aSDoc As SDoc
        Dim theDispo As PublicEnums.Availability
        For Each aSDoc In pDay.Month.DocList
            'conditional code to make doc unavailable if shift is not active for the doc
            Select Case (aShiftType)
                Case 1, 2, 3, 4 'urgence
                    If aSDoc.UrgenceTog = False Then theDispo = Availability.NonDispoPermanente _
                        Else theDispo = Availability.Dispo
                Case 5 'urgence nuit
                    If aSDoc.UrgenceTog = False Or aSDoc.NuitsTog = False Then _
                        theDispo = Availability.NonDispoPermanente Else theDispo = Availability.Dispo

                Case 6 'hospit
                    If aSDoc.HospitTog = False Then theDispo = Availability.NonDispoPermanente _
                        Else theDispo = Availability.Dispo
                Case 7 'soins
                    If aSDoc.SoinsTog = False Then theDispo = Availability.NonDispoPermanente _
                        Else theDispo = Availability.Dispo
                Case Else
                    theDispo = Availability.Dispo
            End Select
            theSDocAvailable = New SDocAvailable(aSDoc.Initials, _
                                                               theDispo, _
                                                               pDate, _
                                                               pShiftType)
            pDocAvailabilities.Add(theSDocAvailable, aSDoc.Initials)
        Next

    End Sub

End Class

Public Class SShiftType
    Private pShiftStart As T_DBRefTypeI
    Private pShiftStop As T_DBRefTypeI
    Private pShiftType As T_DBRefTypeI
    Private pActive As T_DBRefTypeB
    Private pCompilation As T_DBRefTypeB
    Private pDescription As T_DBRefTypeS
    Private pVersion As T_DBRefTypeI
    Private pLundi As T_DBRefTypeB
    Private pMardi As T_DBRefTypeB
    Private pMercredi As T_DBRefTypeB
    Private pJeudi As T_DBRefTypeB
    Private pVendredi As T_DBRefTypeB
    Private pSamedi As T_DBRefTypeB
    Private pDimanche As T_DBRefTypeB
    Private pFerie As T_DBRefTypeB
    Private pOrder As T_DBRefTypeI




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
    Public Property Compilation() As Boolean
        Get
            Return pCompilation.theValue
        End Get
        Set(ByVal value As Boolean)
            pCompilation.theValue = value
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
    Public Property Lundi() As Boolean
        Get
            Return pLundi.theValue
        End Get
        Set(ByVal value As Boolean)
            pLundi.theValue = value
        End Set
    End Property
    Public Property Mardi() As Boolean
        Get
            Return pMardi.theValue
        End Get
        Set(ByVal value As Boolean)
            pMardi.theValue = value
        End Set
    End Property
    Public Property Mercredi() As Boolean
        Get
            Return pMercredi.theValue
        End Get
        Set(ByVal value As Boolean)
            pMercredi.theValue = value
        End Set
    End Property
    Public Property Jeudi() As Boolean
        Get
            Return pJeudi.theValue
        End Get
        Set(ByVal value As Boolean)
            pJeudi.theValue = value
        End Set
    End Property
    Public Property Vendredi() As Boolean
        Get
            Return pVendredi.theValue
        End Get
        Set(ByVal value As Boolean)
            pVendredi.theValue = value
        End Set
    End Property
    Public Property Samedi() As Boolean
        Get
            Return pSamedi.theValue
        End Get
        Set(ByVal value As Boolean)
            pSamedi.theValue = value
        End Set
    End Property
    Public Property Dimanche() As Boolean
        Get
            Return pDimanche.theValue
        End Get
        Set(ByVal value As Boolean)
            pDimanche.theValue = value
        End Set
    End Property
    Public Property Ferie() As Boolean
        Get
            Return pFerie.theValue
        End Get
        Set(ByVal value As Boolean)
            pFerie.theValue = value
        End Set
    End Property
    Public Property Order() As Integer
        Get
            Return pOrder.theValue
        End Get
        Set(ByVal value As Integer)
            pOrder.theValue = value
        End Set
    End Property


    Public Sub New()
        pShiftStart.theSQLName = SQLShiftStart
        pShiftStop.theSQLName = SQLShiftStop
        pShiftType.theSQLName = SQLShiftType
        pActive.theSQLName = SQLActive
        pDescription.theSQLName = SQLDescription
        pVersion.theSQLName = SQLVersion
        pLundi.theSQLName = SQLLundi
        pMardi.theSQLName = SQLMardi
        pMercredi.theSQLName = SQLMercredi
        pJeudi.theSQLName = SQLJeudi
        pVendredi.theSQLName = SQLVendredi
        pSamedi.theSQLName = SQLSamedi
        pDimanche.theSQLName = SQLDimanche
        pFerie.theSQLName = SQLFerie
        pCompilation.theSQLName = SQLCompilation
        pOrder.theSQLName = SQLOrder
    End Sub
    Public Shared Function loadShiftTypesFromDBPerMonth(aMonth As Integer, aYear As Integer) As List(Of SShiftType)
        Dim theBuiltSql As New SQLStrBuilder
        Dim theRS As New ADODB.Recordset
        Dim theDBAC As New DBAC
        Dim aShifttype As SShiftType
        Dim theShiftTypeCollection As List(Of SShiftType)
        theShiftTypeCollection = New List(Of SShiftType)
        Dim theVersion As Integer : theVersion = ((aYear - 2000) * 100) + aMonth

        'check if a version exists for the month

        With theBuiltSql
            .SQL_Select("*")
            .SQL_From(TABLE_shiftType)
            .SQL_Where(SQLVersion, "=", theVersion)
            .SQL_Order_By(SQLOrder)
            theDBAC.COpenDB(.SQLStringSelect, theRS)
        End With

        If theRS.RecordCount > 0 Then 'if a version exists load it
            theRS.MoveFirst()
            For x As Integer = 1 To theRS.RecordCount
                aShifttype = New SShiftType()
                If Not IsDBNull(theRS.Fields(SQLShiftStart).Value) Then _
                    aShifttype.ShiftStart = CInt(theRS.Fields(SQLShiftStart).Value)
                If Not IsDBNull(theRS.Fields(SQLShiftStop).Value) Then _
                    aShifttype.ShiftStop = CInt(theRS.Fields(SQLShiftStop).Value)
                If Not IsDBNull(theRS.Fields(SQLShiftType).Value) Then _
                    aShifttype.ShiftType = CInt(theRS.Fields(SQLShiftType).Value)
                If Not IsDBNull(theRS.Fields(SQLActive).Value) Then _
                    aShifttype.Active = CBool(theRS.Fields(SQLActive).Value)
                If Not IsDBNull(theRS.Fields(SQLVersion).Value) Then _
                    aShifttype.Version = CInt(theRS.Fields(SQLVersion).Value)
                If Not IsDBNull(theRS.Fields(SQLDescription).Value) Then _
                    aShifttype.Description = CStr(theRS.Fields(SQLDescription).Value)
                If Not IsDBNull(theRS.Fields(SQLLundi).Value) Then _
                    aShifttype.Lundi = CBool(theRS.Fields(SQLLundi).Value)
                If Not IsDBNull(theRS.Fields(SQLMardi).Value) Then _
                    aShifttype.Mardi = CBool(theRS.Fields(SQLMardi).Value)
                If Not IsDBNull(theRS.Fields(SQLMercredi).Value) Then _
                    aShifttype.Mercredi = CBool(theRS.Fields(SQLMercredi).Value)
                If Not IsDBNull(theRS.Fields(SQLJeudi).Value) Then _
                    aShifttype.Jeudi = CBool(theRS.Fields(SQLJeudi).Value)
                If Not IsDBNull(theRS.Fields(SQLVendredi).Value) Then _
                    aShifttype.Vendredi = CBool(theRS.Fields(SQLVendredi).Value)
                If Not IsDBNull(theRS.Fields(SQLSamedi).Value) Then _
                    aShifttype.Samedi = CBool(theRS.Fields(SQLSamedi).Value)
                If Not IsDBNull(theRS.Fields(SQLDimanche).Value) Then _
                    aShifttype.Dimanche = CBool(theRS.Fields(SQLDimanche).Value)
                If Not IsDBNull(theRS.Fields(SQLFerie).Value) Then _
                    aShifttype.Ferie = CBool(theRS.Fields(SQLFerie).Value)
                If Not IsDBNull(theRS.Fields(SQLCompilation).Value) Then _
                    aShifttype.Compilation = CBool(theRS.Fields(SQLCompilation).Value)
                If Not IsDBNull(theRS.Fields(SQLOrder).Value) Then _
                    aShifttype.Order = CInt(theRS.Fields(SQLOrder).Value)


                theShiftTypeCollection.Add(aShifttype)
                theRS.MoveNext()
            Next
        Else 'if no version exists, load the template version (0)
            With theBuiltSql
                .SQLClear()
                .SQL_Select("*")
                .SQL_From(TABLE_shiftType)
                .SQL_Where(SQLVersion, "=", 0)
                .SQL_Order_By(SQLOrder)
                theDBAC.COpenDB(.SQLStringSelect, theRS)
            End With

            If theRS.RecordCount > 0 Then 'if at least one template shifttype exists load it as a collection

                theRS.MoveFirst()
                For x As Integer = 1 To theRS.RecordCount
                    aShifttype = New SShiftType()
                    If Not IsDBNull(theRS.Fields(SQLShiftStart).Value) Then _
                        aShifttype.ShiftStart = CInt(theRS.Fields(SQLShiftStart).Value)
                    If Not IsDBNull(theRS.Fields(SQLShiftStop).Value) Then _
                        aShifttype.ShiftStop = CInt(theRS.Fields(SQLShiftStop).Value)
                    If Not IsDBNull(theRS.Fields(SQLShiftType).Value) Then _
                        aShifttype.ShiftType = CInt(theRS.Fields(SQLShiftType).Value)
                    If Not IsDBNull(theRS.Fields(SQLActive).Value) Then _
                        aShifttype.Active = CBool(theRS.Fields(SQLActive).Value)
                    aShifttype.Version = theVersion 'change version to YYYYMM integer
                    If Not IsDBNull(theRS.Fields(SQLDescription).Value) Then _
                        aShifttype.Description = CStr(theRS.Fields(SQLDescription).Value)
                    If Not IsDBNull(theRS.Fields(SQLLundi).Value) Then _
                        aShifttype.Lundi = CBool(theRS.Fields(SQLLundi).Value)
                    If Not IsDBNull(theRS.Fields(SQLMardi).Value) Then _
                        aShifttype.Mardi = CBool(theRS.Fields(SQLMardi).Value)
                    If Not IsDBNull(theRS.Fields(SQLMercredi).Value) Then _
                        aShifttype.Mercredi = CBool(theRS.Fields(SQLMercredi).Value)
                    If Not IsDBNull(theRS.Fields(SQLJeudi).Value) Then _
                        aShifttype.Jeudi = CBool(theRS.Fields(SQLJeudi).Value)
                    If Not IsDBNull(theRS.Fields(SQLVendredi).Value) Then _
                        aShifttype.Vendredi = CBool(theRS.Fields(SQLVendredi).Value)
                    If Not IsDBNull(theRS.Fields(SQLSamedi).Value) Then _
                        aShifttype.Samedi = CBool(theRS.Fields(SQLSamedi).Value)
                    If Not IsDBNull(theRS.Fields(SQLDimanche).Value) Then _
                        aShifttype.Dimanche = CBool(theRS.Fields(SQLDimanche).Value)
                    If Not IsDBNull(theRS.Fields(SQLFerie).Value) Then _
                        aShifttype.Ferie = CBool(theRS.Fields(SQLFerie).Value)
                    If Not IsDBNull(theRS.Fields(SQLCompilation).Value) Then _
                        aShifttype.Compilation = CBool(theRS.Fields(SQLCompilation).Value)
                    If Not IsDBNull(theRS.Fields(SQLOrder).Value) Then _
                        aShifttype.Order = CInt(theRS.Fields(SQLOrder).Value)
                    aShifttype.Save() 'save the shifttype version to DB
                    theShiftTypeCollection.Add(aShifttype)
                    theRS.MoveNext()
                Next

            End If
        End If
        Return theShiftTypeCollection
    End Function
    Public Shared Function loadTemplateShiftTypesFromDB() As List(Of SShiftType)
        Dim theBuiltSql As New SQLStrBuilder
        Dim theRS As New ADODB.Recordset
        Dim theDBAC As New DBAC
        Dim aShifttype As SShiftType
        Dim theShiftTypeCollection As List(Of SShiftType)
        theShiftTypeCollection = New List(Of SShiftType)
        With theBuiltSql
            .SQL_Select("*")
            .SQL_From(TABLE_shiftType)
            .SQL_Where(SQLVersion, "=", 0)
            .SQL_Order_By(SQLOrder)
            theDBAC.COpenDB(.SQLStringSelect, theRS)
        End With

        If theRS.RecordCount > 0 Then

            theRS.MoveFirst()
            For x As Integer = 1 To theRS.RecordCount
                aShifttype = New SShiftType()
                If Not IsDBNull(theRS.Fields(SQLShiftStart).Value) Then _
                    aShifttype.ShiftStart = CInt(theRS.Fields(SQLShiftStart).Value)
                If Not IsDBNull(theRS.Fields(SQLShiftStop).Value) Then _
                    aShifttype.ShiftStop = CInt(theRS.Fields(SQLShiftStop).Value)
                If Not IsDBNull(theRS.Fields(SQLShiftType).Value) Then _
                    aShifttype.ShiftType = CInt(theRS.Fields(SQLShiftType).Value)
                If Not IsDBNull(theRS.Fields(SQLActive).Value) Then _
                    aShifttype.Active = CBool(theRS.Fields(SQLActive).Value)
                If Not IsDBNull(theRS.Fields(SQLDescription).Value) Then _
                    aShifttype.Description = CStr(theRS.Fields(SQLDescription).Value)
                If Not IsDBNull(theRS.Fields(SQLLundi).Value) Then _
                    aShifttype.Lundi = CBool(theRS.Fields(SQLLundi).Value)
                If Not IsDBNull(theRS.Fields(SQLMardi).Value) Then _
                    aShifttype.Mardi = CBool(theRS.Fields(SQLMardi).Value)
                If Not IsDBNull(theRS.Fields(SQLMercredi).Value) Then _
                    aShifttype.Mercredi = CBool(theRS.Fields(SQLMercredi).Value)
                If Not IsDBNull(theRS.Fields(SQLJeudi).Value) Then _
                    aShifttype.Jeudi = CBool(theRS.Fields(SQLJeudi).Value)
                If Not IsDBNull(theRS.Fields(SQLVendredi).Value) Then _
                    aShifttype.Vendredi = CBool(theRS.Fields(SQLVendredi).Value)
                If Not IsDBNull(theRS.Fields(SQLSamedi).Value) Then _
                    aShifttype.Samedi = CBool(theRS.Fields(SQLSamedi).Value)
                If Not IsDBNull(theRS.Fields(SQLDimanche).Value) Then _
                    aShifttype.Dimanche = CBool(theRS.Fields(SQLDimanche).Value)
                If Not IsDBNull(theRS.Fields(SQLFerie).Value) Then _
                    aShifttype.Ferie = CBool(theRS.Fields(SQLFerie).Value)
                If Not IsDBNull(theRS.Fields(SQLCompilation).Value) Then _
                    aShifttype.Compilation = CBool(theRS.Fields(SQLCompilation).Value)
                If Not IsDBNull(theRS.Fields(SQLOrder).Value) Then _
                    aShifttype.Order = CInt(theRS.Fields(SQLOrder).Value)

                theShiftTypeCollection.Add(aShifttype)
                theRS.MoveNext()
            Next
        End If
        Return theShiftTypeCollection
    End Function
    Public Shared Function ActiveShiftTypesCountPerMonth(aMonth As Integer, aYear As Integer) As Integer
        Dim theBuiltSql As New SQLStrBuilder
        Dim theRS As New ADODB.Recordset
        Dim theDBAC As New DBAC
        Dim theVersion As Integer : theVersion = ((aYear - 2000) * 100) + aMonth

        'check if a version exists for the month

        With theBuiltSql
            .SQL_Select("*")
            .SQL_From(TABLE_shiftType)
            .SQL_Where(SQLVersion, "=", theVersion)
            .SQL_Where(SQLActive, "=", True)
            theDBAC.COpenDB(.SQLStringSelect, theRS)
        End With

        Return theRS.RecordCount
    End Function
    Public Sub Copy(TheInstanceToBeCopied As SShiftType)

        With TheInstanceToBeCopied

            'Me.pCollection = .ShiftCollection
            Me.ShiftStart = .ShiftStart
            Me.ShiftStop = .ShiftStop
            Me.ShiftType = .ShiftType
            Me.Version = .Version
            Me.Active = .Active
            Me.Description = .Description
            Me.Lundi = .Lundi
            Me.Mardi = .Mardi
            Me.Mercredi = .Mercredi
            Me.Jeudi = .Jeudi
            Me.Vendredi = .Vendredi
            Me.Samedi = .Samedi
            Me.Dimanche = .Dimanche
            Me.Ferie = .Ferie
            Me.Compilation = .Compilation
            Me.Order = .Order

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
                    .SQL_Values(pLundi.theSQLName, Lundi)
                    .SQL_Values(pMardi.theSQLName, Mardi)
                    .SQL_Values(pMercredi.theSQLName, Mercredi)
                    .SQL_Values(pJeudi.theSQLName, Jeudi)
                    .SQL_Values(pVendredi.theSQLName, Vendredi)
                    .SQL_Values(pSamedi.theSQLName, Samedi)
                    .SQL_Values(pDimanche.theSQLName, Dimanche)
                    .SQL_Values(pFerie.theSQLName, Ferie)
                    .SQL_Values(pCompilation.theSQLName, Compilation)
                    .SQL_Values(pOrder.theSQLName, Order)

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
                theRS.Fields(pLundi.theSQLName).Value = Lundi
                theRS.Fields(pMardi.theSQLName).Value = Mardi
                theRS.Fields(pMercredi.theSQLName).Value = Mercredi
                theRS.Fields(pJeudi.theSQLName).Value = Jeudi
                theRS.Fields(pVendredi.theSQLName).Value = Vendredi
                theRS.Fields(pSamedi.theSQLName).Value = Samedi
                theRS.Fields(pDimanche.theSQLName).Value = Dimanche
                theRS.Fields(pFerie.theSQLName).Value = Ferie
                theRS.Fields(pCompilation.theSQLName).Value = Compilation
                theRS.Fields(pOrder.theSQLName).Value = Order
                theRS.ActiveConnection = theDBAC.aConnection
                theRS.UpdateBatch()
                theRS.Close()
            Case Else
                Debug.WriteLine("there is more than one copy of this entry ... this is bad")
        End Select
    End Sub

End Class

Public Class SDoc
    Private pFirstName As T_DBRefTypeS
    Private pLastName As T_DBRefTypeS
    Private pInitials As T_DBRefTypeS
    Private pActive As T_DBRefTypeB
    Private pVersion As T_DBRefTypeI
    Private pShift1 As T_DBRefTypeI
    Private pShift2 As T_DBRefTypeI
    Private pShift3 As T_DBRefTypeI
    Private pShift4 As T_DBRefTypeI
    Private pShift5 As T_DBRefTypeI
    Private pUrgenceTog As T_DBRefTypeB
    Private pHospitTog As T_DBRefTypeB
    Private pSoinsTog As T_DBRefTypeB
    Private pNuitsTog As T_DBRefTypeB
    Private pYear As Integer
    Private pMonth As Integer

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
    Public Property Shift1() As Integer
        Get
            Return pShift1.theValue
        End Get
        Set(ByVal value As Integer)
            pShift1.theValue = value
        End Set
    End Property
    Public Property Shift2() As Integer
        Get
            Return pShift2.theValue
        End Get
        Set(ByVal value As Integer)
            pShift2.theValue = value
        End Set
    End Property
    Public Property Shift3() As Integer
        Get
            Return pShift3.theValue
        End Get
        Set(ByVal value As Integer)
            pShift3.theValue = value
        End Set
    End Property
    Public Property Shift4() As Integer
        Get
            Return pShift4.theValue
        End Get
        Set(ByVal value As Integer)
            pShift4.theValue = value
        End Set
    End Property
    Public Property Shift5() As Integer
        Get
            Return pShift5.theValue
        End Get
        Set(ByVal value As Integer)
            pShift5.theValue = value
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
        pShift1.theSQLName = SQLShift1
        pShift2.theSQLName = SQLShift2
        pShift3.theSQLName = SQLShift3
        pShift4.theSQLName = SQLShift4
        pShift5.theSQLName = SQLShift5
        pUrgenceTog.theSQLName = SQLUrgenceTog
        pHospitTog.theSQLName = SQLHospitTog
        pSoinsTog.theSQLName = SQLSoinsTog
        pNuitsTog.theSQLName = SQLNuitsTog

        FirstName = "FirstName"
        LastName = "LastName"
        Initials = "Initialles"
        Active = True
        Version = 1
        Shift1 = 0
        Shift2 = 0
        Shift3 = 0
        Shift4 = 0
        Shift5 = 0
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
            .SQL_Where(pShift1.theSQLName, "=", Shift1)
            .SQL_Where(pShift2.theSQLName, "=", Shift2)
            .SQL_Where(pShift3.theSQLName, "=", Shift3)
            .SQL_Where(pShift4.theSQLName, "=", Shift4)
            .SQL_Where(pShift5.theSQLName, "=", Shift5)
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
        pShift1.theSQLName = SQLShift1
        pShift2.theSQLName = SQLShift2
        pShift3.theSQLName = SQLShift3
        pShift4.theSQLName = SQLShift4
        pShift5.theSQLName = SQLShift5
        pUrgenceTog.theSQLName = SQLUrgenceTog
        pHospitTog.theSQLName = SQLHospitTog
        pSoinsTog.theSQLName = SQLSoinsTog
        pNuitsTog.theSQLName = SQLNuitsTog

        FirstName = "FirstName"
        LastName = "LastName"
        Initials = "Initialles"
        Active = True
        Version = 1
        Shift1 = 0
        Shift2 = 0
        Shift3 = 0
        Shift4 = 0
        Shift5 = 0
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
                    aShift1 As Integer, _
                    aShift2 As Integer, _
                    aShift3 As Integer, _
                    aShift4 As Integer, _
                    aShift5 As Integer, _
                    aUrgenceTog As Boolean, _
                    aHospitTog As Boolean, _
                    aSoinsTog As Boolean, _
                    aNuitsTog As Boolean)

        pFirstName.theSQLName = SQLFirstName
        pLastName.theSQLName = SQLLastName
        pInitials.theSQLName = SQLInitials
        pActive.theSQLName = SQLActive
        pVersion.theSQLName = SQLVersion
        pShift1.theSQLName = SQLShift1
        pShift2.theSQLName = SQLShift2
        pShift3.theSQLName = SQLShift3
        pShift4.theSQLName = SQLShift4
        pShift5.theSQLName = SQLShift5
        pUrgenceTog.theSQLName = SQLUrgenceTog
        pHospitTog.theSQLName = SQLHospitTog
        pSoinsTog.theSQLName = SQLSoinsTog
        pNuitsTog.theSQLName = SQLNuitsTog

        FirstName = aFirstName
        LastName = aLastName
        Initials = aInitials
        Active = aActive
        Version = aVersion
        Shift1 = aShift1
        Shift2 = aShift2
        Shift3 = aShift3
        Shift4 = aShift4
        Shift5 = aShift5

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
                    .SQL_Values(pShift1.theSQLName, Shift1)
                    .SQL_Values(pShift2.theSQLName, Shift2)
                    .SQL_Values(pShift3.theSQLName, Shift3)
                    .SQL_Values(pShift4.theSQLName, Shift4)
                    .SQL_Values(pShift5.theSQLName, Shift5)
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
                theRS.Fields(pShift1.theSQLName).Value = Shift1
                theRS.Fields(pShift2.theSQLName).Value = Shift2
                theRS.Fields(pShift3.theSQLName).Value = Shift3
                theRS.Fields(pShift4.theSQLName).Value = Shift4
                theRS.Fields(pShift5.theSQLName).Value = Shift5
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
    Public Shared Function LoadAllDocsPerMonth(aYear As Integer, aMonth As Integer) As List(Of SDoc)
        Dim theBuiltSql As New SQLStrBuilder
        Dim theRS As New ADODB.Recordset
        Dim theDBAC As New DBAC
        Dim theCurrentMonthDate As Date = DateSerial(aYear, aMonth, 1)
        Dim aCollection As List(Of SDoc)
        aCollection = New List(Of SDoc)
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
                Dim aSDoc As New SDoc()
                If Not IsDBNull(theRS.Fields(SQLFirstName).Value) Then _
                aSDoc.FirstName = CStr(theRS.Fields(SQLFirstName).Value)
                If Not IsDBNull(theRS.Fields(SQLLastName).Value) Then _
                aSDoc.LastName = CStr(theRS.Fields(SQLLastName).Value)
                If Not IsDBNull(theRS.Fields(SQLInitials).Value) Then _
                aSDoc.Initials = CStr(theRS.Fields(SQLInitials).Value)
                If Not IsDBNull(theRS.Fields(SQLActive).Value) Then _
                aSDoc.Active = CBool(theRS.Fields(SQLActive).Value)
                If Not IsDBNull(theRS.Fields(SQLVersion).Value) Then _
                aSDoc.Version = CInt(theRS.Fields(SQLVersion).Value)
                If Not IsDBNull(theRS.Fields(SQLShift1).Value) Then _
                aSDoc.Shift1 = CInt(theRS.Fields(SQLShift1).Value)
                If Not IsDBNull(theRS.Fields(SQLShift2).Value) Then _
                aSDoc.Shift2 = CInt(theRS.Fields(SQLShift2).Value)
                If Not IsDBNull(theRS.Fields(SQLShift3).Value) Then _
                aSDoc.Shift3 = CInt(theRS.Fields(SQLShift3).Value)
                If Not IsDBNull(theRS.Fields(SQLShift4).Value) Then _
                aSDoc.Shift4 = CInt(theRS.Fields(SQLShift4).Value)
                If Not IsDBNull(theRS.Fields(SQLShift5).Value) Then _
                aSDoc.Shift5 = CInt(theRS.Fields(SQLShift5).Value)

                If Not IsDBNull(theRS.Fields(SQLUrgenceTog).Value) Then _
                    aSDoc.UrgenceTog = CBool(theRS.Fields(SQLUrgenceTog).Value)
                If Not IsDBNull(theRS.Fields(SQLHospitTog).Value) Then _
                    aSDoc.HospitTog = CBool(theRS.Fields(SQLHospitTog).Value)
                If Not IsDBNull(theRS.Fields(SQLSoinsTog).Value) Then _
                    aSDoc.SoinsTog = CBool(theRS.Fields(SQLSoinsTog).Value)
                If Not IsDBNull(theRS.Fields(SQLNuitsTog).Value) Then _
                    aSDoc.NuitsTog = CBool(theRS.Fields(SQLNuitsTog).Value)

                aCollection.Add(aSDoc)
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
                    Dim aSDoc As New SDoc()
                    If Not IsDBNull(theRS.Fields(SQLFirstName).Value) Then _
                    aSDoc.FirstName = CStr(theRS.Fields(SQLFirstName).Value)
                    If Not IsDBNull(theRS.Fields(SQLLastName).Value) Then _
                    aSDoc.LastName = CStr(theRS.Fields(SQLLastName).Value)
                    If Not IsDBNull(theRS.Fields(SQLInitials).Value) Then _
                    aSDoc.Initials = CStr(theRS.Fields(SQLInitials).Value)
                    If Not IsDBNull(theRS.Fields(SQLActive).Value) Then _
                    aSDoc.Active = CBool(theRS.Fields(SQLActive).Value)
                    aSDoc.Version = theVersion 'change version to YYYYMM integer
                    If Not IsDBNull(theRS.Fields(SQLShift1).Value) Then _
                    aSDoc.Shift1 = CInt(theRS.Fields(SQLShift1).Value)
                    If Not IsDBNull(theRS.Fields(SQLShift2).Value) Then _
                    aSDoc.Shift2 = CInt(theRS.Fields(SQLShift2).Value)
                    If Not IsDBNull(theRS.Fields(SQLShift3).Value) Then _
                    aSDoc.Shift3 = CInt(theRS.Fields(SQLShift3).Value)
                    If Not IsDBNull(theRS.Fields(SQLShift4).Value) Then _
                    aSDoc.Shift4 = CInt(theRS.Fields(SQLShift4).Value)
                    If Not IsDBNull(theRS.Fields(SQLShift5).Value) Then _
                    aSDoc.Shift5 = CInt(theRS.Fields(SQLShift5).Value)
                    If Not IsDBNull(theRS.Fields(SQLUrgenceTog).Value) Then _
                        aSDoc.UrgenceTog = CBool(theRS.Fields(SQLUrgenceTog).Value)
                    If Not IsDBNull(theRS.Fields(SQLHospitTog).Value) Then _
                        aSDoc.HospitTog = CBool(theRS.Fields(SQLHospitTog).Value)
                    If Not IsDBNull(theRS.Fields(SQLSoinsTog).Value) Then _
                        aSDoc.SoinsTog = CBool(theRS.Fields(SQLSoinsTog).Value)
                    If Not IsDBNull(theRS.Fields(SQLNuitsTog).Value) Then _
                        aSDoc.NuitsTog = CBool(theRS.Fields(SQLNuitsTog).Value)
                    aSDoc.save()
                    aCollection.Add(aSDoc)
                    theRS.MoveNext()
                Next
            End If
        End If

        Return aCollection
    End Function
    Public Shared Function LoadTempateDocsFromDB() As List(Of SDoc)
        Dim theBuiltSql As New SQLStrBuilder
        Dim theRS As New ADODB.Recordset
        Dim theDBAC As New DBAC
        Dim aCollection As List(Of SDoc)
        aCollection = New List(Of SDoc)

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
                Dim aSDoc As New SDoc()
                If Not IsDBNull(theRS.Fields(SQLFirstName).Value) Then _
                aSDoc.FirstName = CStr(theRS.Fields(SQLFirstName).Value)
                If Not IsDBNull(theRS.Fields(SQLLastName).Value) Then _
                aSDoc.LastName = CStr(theRS.Fields(SQLLastName).Value)
                If Not IsDBNull(theRS.Fields(SQLInitials).Value) Then _
                aSDoc.Initials = CStr(theRS.Fields(SQLInitials).Value)
                If Not IsDBNull(theRS.Fields(SQLActive).Value) Then _
                aSDoc.Active = CBool(theRS.Fields(SQLActive).Value)
                If Not IsDBNull(theRS.Fields(SQLVersion).Value) Then _
                aSDoc.Version = CInt(theRS.Fields(SQLVersion).Value)
                If Not IsDBNull(theRS.Fields(SQLShift1).Value) Then _
                aSDoc.Shift1 = CInt(theRS.Fields(SQLShift1).Value)
                If Not IsDBNull(theRS.Fields(SQLShift2).Value) Then _
                aSDoc.Shift2 = CInt(theRS.Fields(SQLShift2).Value)
                If Not IsDBNull(theRS.Fields(SQLShift3).Value) Then _
                aSDoc.Shift3 = CInt(theRS.Fields(SQLShift3).Value)
                If Not IsDBNull(theRS.Fields(SQLShift4).Value) Then _
                aSDoc.Shift4 = CInt(theRS.Fields(SQLShift4).Value)
                If Not IsDBNull(theRS.Fields(SQLShift5).Value) Then _
                aSDoc.Shift5 = CInt(theRS.Fields(SQLShift5).Value)
                If Not IsDBNull(theRS.Fields(SQLUrgenceTog).Value) Then _
                    aSDoc.UrgenceTog = CBool(theRS.Fields(SQLUrgenceTog).Value)
                If Not IsDBNull(theRS.Fields(SQLHospitTog).Value) Then _
                    aSDoc.HospitTog = CBool(theRS.Fields(SQLHospitTog).Value)
                If Not IsDBNull(theRS.Fields(SQLSoinsTog).Value) Then _
                    aSDoc.SoinsTog = CBool(theRS.Fields(SQLSoinsTog).Value)
                If Not IsDBNull(theRS.Fields(SQLNuitsTog).Value) Then _
                    aSDoc.NuitsTog = CBool(theRS.Fields(SQLNuitsTog).Value)

                aCollection.Add(aSDoc)
                theRS.MoveNext()
            Next
        End If
        Return aCollection
    End Function

End Class

Public Class SDocStats
    Private pInitials As String
    Private pShift1 As Integer
    Private pShift2 As Integer
    Private pShift3 As Integer
    Private pShift4 As Integer
    Private pShift5 As Integer
    Private pShift1E As Integer
    Private pShift2E As Integer
    Private pShift3E As Integer
    Private pShift4E As Integer
    Private pShift5E As Integer

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
    Public Property shift1E() As Integer
        Get
            Return pShift1E
        End Get
        Set(ByVal value As Integer)
            pShift1E = value
        End Set
    End Property
    Public Property shift2E() As Integer
        Get
            Return pShift2E
        End Get
        Set(ByVal value As Integer)
            pShift2E = value
        End Set
    End Property
    Public Property shift3E() As Integer
        Get
            Return pShift3E
        End Get
        Set(ByVal value As Integer)
            pShift3E = value
        End Set
    End Property
    Public Property shift4E() As Integer
        Get
            Return pShift4E
        End Get
        Set(ByVal value As Integer)
            pShift4E = value
        End Set
    End Property
    Public Property shift5E() As Integer
        Get
            Return pShift5E
        End Get
        Set(ByVal value As Integer)
            pShift5E = value
        End Set
    End Property

    Public Sub New(aInitials As String,
                        aShift1 As Integer, _
                        aShift2 As Integer, _
                        aShift3 As Integer, _
                        aShift4 As Integer, _
                        aShift5 As Integer)
        Initials = aInitials
        shift1 = aShift1
        shift2 = aShift2
        shift3 = aShift3
        shift4 = aShift4
        shift5 = aShift5
        shift1E = aShift1
        shift2E = aShift2
        shift3E = aShift3
        shift4E = aShift4
        shift5E = aShift5

    End Sub

End Class

Public Class SDocAvailable
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
        Availability = CType(aAvailability, PublicEnums.Availability)
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
            Dim aSDocAvailable As SDocAvailable
            Dim aCollection As New Collection
            theRS.MoveFirst()
            For x As Integer = 1 To theRS.RecordCount
                aSDocAvailable = New SDocAvailable
                aSDocAvailable.DocInitial = CStr(theRS.Fields(Me.pDocInitial.theSQLName).Value)
                aSDocAvailable.Date_ = CDate(theRS.Fields(Me.pDate.theSQLName).Value)
                aSDocAvailable.ShiftType = CInt(theRS.Fields(Me.pShiftType.theSQLName).Value)
                aCollection.Add(aSDocAvailable)
                theRS.MoveNext()
            Next
            Return aCollection
        End If
        Return Nothing
    End Function

End Class

Public Class SNonDispo
    Private pDocInitial As T_DBRefTypeS
    Private pDateStart As T_DBRefTypeD
    Private pDateStop As T_DBRefTypeD
    Private pTimeStart As T_DBRefTypeI
    Private pTimeStop As T_DBRefTypeI
    Private pDu As String
    Private pAu As String

    Public ReadOnly Property du() As String
        Get
            Dim myhours As Integer = CInt(pTimeStart.theValue / 60)
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
            Dim myhours As Integer = CInt(pTimeStop.theValue / 60)
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
        Dim aSNonDispo As SNonDispo
        Dim theCount As Integer = theRS.RecordCount
        If theCount > 0 Then
            Dim aCollection As New Collection
            theRS.MoveFirst()
            For x As Integer = 1 To theCount
                aSNonDispo = New SNonDispo
                If Not IsDBNull(theRS.Fields(pDocInitial.theSQLName).Value) Then aSNonDispo.DocInitial = CStr(theRS.Fields(pDocInitial.theSQLName).Value)
                If Not IsDBNull(theRS.Fields(pDateStart.theSQLName).Value) Then aSNonDispo.DateStart = CDate(theRS.Fields(pDateStart.theSQLName).Value)
                If Not IsDBNull(theRS.Fields(pTimeStart.theSQLName).Value) Then aSNonDispo.TimeStart = CInt(theRS.Fields(pTimeStart.theSQLName).Value)
                If Not IsDBNull(theRS.Fields(pDateStop.theSQLName).Value) Then aSNonDispo.DateStop = CDate(theRS.Fields(pDateStop.theSQLName).Value)
                If Not IsDBNull(theRS.Fields(pTimeStop.theSQLName).Value) Then aSNonDispo.TimeStop = CInt(theRS.Fields(pTimeStop.theSQLName).Value)
                aCollection.Add(aSNonDispo, x.ToString())
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

    Const Provider As String = "Provider=Microsoft.ACE.OLEDB.12.0;"
    'Const DBpassword = "Jet OLEDB:Database Password=plasma;"

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


            If CONSTFILEADDRESS = "" Then
                If MySettingsGlobal.DataBaseLocation = "" Then
                    LoadDatabaseFileLocation()
                End If
                CONSTFILEADDRESS = MySettingsGlobal.DataBaseLocation
            End If
            mConnectionString = Provider + "Data Source=" _
            + CONSTFILEADDRESS '_
            '+ ";" & DBpassword
            cnn.ConnectionString = mConnectionString
            cnn.Open()
        End If

        mConnection = cnn
        On Error GoTo 0
        Exit Sub
errhandler:
        MsgBox("An error occurred during initial connection to DB: " + _
               CStr(Err.Number) + "  :  " + _
               CStr(Err.Description))

        '        'add code to select current location for the database !!FEATURE!!

    End Sub
    Public Sub New(fileAddress As String)
        On Error GoTo errhandler
        If cnn.State = ADODB.ObjectStateEnum.adStateClosed Then

            mConnectionString = Provider + "Data Source=" + fileAddress
            cnn.ConnectionString = mConnectionString
            cnn.Open()
        End If

        mConnection = cnn
        On Error GoTo 0
        Exit Sub
errhandler:
        MsgBox("An error occurred during initial connection to DB: " + _
               CStr(Err.Number) + "  :  " + _
               CStr(Err.Description))

        '        'add code to select current location for the database !!FEATURE!!

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

        MsgBox("An error occurred during SELECT query execution: " + _
               CStr(Err.Number) & "  :  " + _
               CStr(Err.Description) + _
               "   SQL TEXT:   " + _
               theSQLstr)

    End Sub
    Public Sub CExecuteDB(theSQLstr As String, numAffected As Long)

        On Error GoTo errhandler
        Dim theNumAffectedObj As Object
        mConnection.Execute(theSQLstr, theNumAffectedObj)
        'StoreToAuditFile theSQLstr
        On Error GoTo 0
        numAffected = CLng(theNumAffectedObj)
        Exit Sub 'to not run errhandler for nothing


errhandler:

        Dim theError As String
        theError = CStr(Err.Number) & "  :  " & CStr(Err.Description) & "   SQL TEXT:   " & theSQLstr
        theError = Replace(theError, "'", "''")
        MsgBox("An error occurred during execution of an INSERT or UPDATE query: " & theError)

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
                    theValueStr = "'" + CStr(theValue) + "'"
                Else
                    theValueStr = CStr(theValue)
                End If

            Case "Date"
                theValueStr = cAccessDateStr(CDate(theValue))
            Case "Boolean"
                If CBool(theValue) = True Then
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
                theValueStr = "'" & CStr(theValue) & "'"
            Case "Date"
                theValueStr = cAccessDateStr(CDate(theValue))
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
                theValueStr = "'" & CStr(theValue) & "'"
            Case "Date"
                theValueStr = cAccessDateStr(CDate(theValue))
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





