Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports ListeDeGarde
Imports System.Diagnostics
Imports Microsoft.Office.Interop

<TestClass()> Public Class classes
    Public Shared testMonth As ScheduleMonth
    Public Shared theDay As ScheduleDay
    Public Shared xlsApp As Excel.Application
    Public Shared xlsWB As Excel.Workbook
    Public Shared xlsSheet As Excel.Worksheet
    Public Shared aController As Controller

    <ClassInitialize()> Public Shared Sub MySetup(sTestContext As TestContext)

        MySettingsGlobal = New Settings1
        testMonth = New ScheduleMonth(4, 2018)

        xlsApp = New Excel.Application
        xlsApp.Visible = True
        xlsWB = xlsApp.Workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet)
        xlsApp.ScreenUpdating = False
        xlsSheet = xlsApp.ActiveSheet
        xlsApp.ScreenUpdating = True
    End Sub
    <TestMethod()> Public Sub DaysCreated()
        'arrange
        'act
        'assert
        Assert.IsTrue(testMonth.Days.Count > 0)
    End Sub

    <TestMethod()> Public Sub ShiftsCreated()
        'arrange
        'act
        theDay = testMonth.Days(1)
        'assert
        Assert.IsTrue(theDay.Shifts.Count > 0)
    End Sub

    <TestMethod()> Public Sub DocAvailCreated()
        'arrange
        'act
        Dim theshift As ScheduleShift
        theshift = theDay.Shifts(1)
        'assert
        Assert.IsTrue(theshift.DocAvailabilities.Count > 0)
    End Sub

    <TestMethod()> Public Sub CreateExcelCalendar()
        aController = New Controller(xlsSheet, 2017, 3, "Mars")
        Assert.IsTrue(xlsSheet.Range("g4").Value = "8-16 Urg.")
    End Sub

    <TestMethod()> Public Sub HighlightingWorks()
        xlsSheet.Range("h4").Value = ""
        aController.HighLightDocAvailablilities("MM")
        Assert.IsTrue(xlsSheet.Range("h4").Interior.Color = RGB(0, 233, 118))
    End Sub

    <TestMethod()> Public Sub AddDoc()
        xlsSheet.Range("h4").Value = "MM"
        Dim aDay As ScheduleDay = aController.aControlledMonth.Days(1)
        Dim aShift As ScheduleShift = aDay.Shifts(1)
        Dim acollection As Collection = aShift.DocAvailabilities
        Dim aDocAvail As scheduleDocAvailable = acollection.Item("MM")
        Assert.IsTrue(aDocAvail.Availability = PublicEnums.Availability.Assigne)
        xlsSheet.Range("h4").Value = ""
    End Sub

    <TestMethod()> Public Sub HighlightingRepeated()
        xlsSheet.Range("h4").Value = "MM"
        aController.HighLightDocAvailablilities("MM")
        Assert.IsTrue(xlsSheet.Range("h4").Interior.Color = RGB(0, 255, 255))
        xlsSheet.Range("h4").Value = ""
    End Sub

    <TestMethod()> Public Sub HighlightingTempNonDispos()
        xlsSheet.Range("h4").Value = "MM"
        aController.HighLightDocAvailablilities("MM")
        Assert.IsTrue(xlsSheet.Range("h5").Interior.Color = RGB(219, 112, 147))
        xlsSheet.Range("h4").Value = ""
    End Sub
End Class