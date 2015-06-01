Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports ListeDeGarde
Imports System
Imports System.Collections.Generic
Imports System.Runtime.Remoting
Imports System.Runtime.Remoting.Channels
Imports System.Runtime.Remoting.Channels.Tcp
Imports Microsoft.Office.Interop


<TestClass()> Public Class BaseClassTests
    <ClassInitialize()> _
    Public Shared Sub ClassInit(ByVal context As TestContext)
        ListeDeGarde.CONSTFILEADDRESS = "C:\Users\sel4_000\Documents\Scheduling Mira\listesdegarde.accdb"
    End Sub

    <TestMethod()>
    Public Sub BaseClassTests_Month()
        Dim aMonth As SMonth
        aMonth = New SMonth(1, 2010)
        Assert.IsTrue(Not aMonth Is Nothing)
    End Sub

    <TestMethod()>
    Public Sub BaseClassTests_Days()
        Dim aMonth As SMonth
        aMonth = New SMonth(1, 2010)
        Assert.IsTrue(aMonth.Days.Count > 0)
    End Sub
    <TestMethod()>
    Public Sub BaseClassTests_Shifts()
        Dim aMonth As SMonth
        Dim aDay As SDay
        aMonth = New SMonth(1, 2010)
        aDay = aMonth.Days(1)
        Assert.IsTrue(aDay.Shifts.Count > 0)
    End Sub
    <TestMethod()>
    Public Sub BaseClassTests_Docs()
        Dim aMonth As SMonth
        aMonth = New SMonth(1, 2010)
        Assert.IsTrue(aMonth.DocList.Count > 0)
    End Sub
    <TestMethod()>
    Public Sub BaseClassTests_ShiftTypes()
        Dim aMonth As SMonth
        aMonth = New SMonth(1, 2010)
        Assert.IsTrue(aMonth.ShiftTypes.Count > 0)
    End Sub
    <TestMethod()>
    Public Sub BaseClassTests_DocAvail()
        Dim aMonth As SMonth
        Dim aDay As SDay
        Dim aShift As SShift
        aMonth = New SMonth(1, 2010)
        aDay = aMonth.Days(1)
        aShift = aDay.Shifts(1)
        Assert.IsTrue(aShift.DocAvailabilities.Count > 0)
    End Sub


End Class

<TestClass()> Public Class DB_Connection
    <TestMethod()> Public Sub DB_Connection_Connects()

        Dim adbac As New DBAC()
        Assert.IsTrue(Not adbac.aConnection Is Nothing)
    End Sub
End Class

