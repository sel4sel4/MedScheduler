Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports ListeDeGarde
Imports System
Imports System.Collections.Generic
Imports System.Runtime.Remoting
Imports System.Runtime.Remoting.Channels
Imports System.Runtime.Remoting.Channels.Tcp
Imports Microsoft.Office.Interop


<TestClass()> Public Class UnitTest1
    <ClassInitialize()> _
    Public Shared Sub ClassInit(ByVal context As TestContext)
        ListeDeGarde.CONSTFILEADDRESS = "C:\Users\sel4_000\Documents\Scheduling Mira\listesdegarde.accdb"
    End Sub

    <TestMethod()>
    Public Sub CanCreateMonth()
        Dim aMonth As SMonth
        aMonth = New SMonth(1, 2010)
        Assert.IsTrue(Not aMonth Is Nothing)
    End Sub

    <TestMethod()>
    Public Sub CanCreateDays()
        Dim aMonth As SMonth
        aMonth = New SMonth(1, 2010)
        Assert.IsTrue(aMonth.Days.Count > 0)
    End Sub
    <TestMethod()>
    Public Sub CanCreateShifts()
        Dim aMonth As SMonth
        Dim aDay As SDay
        aMonth = New SMonth(1, 2010)
        aDay = aMonth.Days(1)
        Assert.IsTrue(aDay.Shifts.Count > 0)
    End Sub
    <TestMethod()>
    Public Sub CanLoadDocs()
        Dim aMonth As SMonth
        aMonth = New SMonth(1, 2010)
        Assert.IsTrue(aMonth.DocList.Count > 0)
    End Sub
    <TestMethod()>
    Public Sub CanLoadShiftTypes()
        Dim aMonth As SMonth
        aMonth = New SMonth(1, 2010)
        Assert.IsTrue(aMonth.ShiftTypes.Count > 0)
    End Sub
    <TestMethod()>
    Public Sub CanLoadDocAvail()
        Dim aMonth As SMonth
        Dim aDay As SDay
        Dim aShift As SShift
        aMonth = New SMonth(1, 2010)
        aDay = aMonth.Days(1)
        aShift = aDay.Shifts(1)
        Assert.IsTrue(aShift.DocAvailabilities.Count > 0)
    End Sub


End Class

<TestClass()> Public Class UnitTest2
    <TestMethod()> Public Sub Database_connects()

        Dim adbac As New DBAC()
        Assert.IsTrue(Not adbac.aConnection Is Nothing)
    End Sub
End Class

<TestClass()> Public Class UnitTest3
    Public Shared testMonth As SMonth
    Public Shared theDay As SDay
    Public Shared xlsApp As Excel.Application
    Public Shared xlsWB As Excel.Workbook
    Public Shared xlsSheet As Excel.Worksheet
    Public Shared aController As Controller

    <TestMethod()> Public Sub openup()
        xlsApp = New Excel.Application
        xlsApp.Visible = True
        xlsWB = xlsApp.Workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet)
        xlsApp.ScreenUpdating = False
        xlsSheet = xlsApp.ActiveSheet
        xlsApp.ScreenUpdating = True
        xlsApp.Visible = True
        'Assert.IsTrue(Not ListeDeGarde.MyAddin.myCustomTaskPane Is Nothing)
        Assert.IsTrue(True)
    End Sub



End Class
