Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports ListeDeGarde
Imports System.Diagnostics

<TestClass()> Public Class classes

    <TestMethod()> Public Sub DaysCreated()
        'arrange
        'act
        MySettingsGlobal = New Settings1
        Dim testMonth As ScheduleMonth
        testMonth = New ScheduleMonth(4, 2015)
        'assert
        Assert.IsTrue(testMonth.Days.Count > 0)
    End Sub

    <TestMethod()> Public Sub ShiftsCreated()
        'arrange
        'act
        Dim theDay As ScheduleDay
        MySettingsGlobal = New Settings1
        Dim testMonth As ScheduleMonth
        testMonth = New ScheduleMonth(4, 2015)

        theDay = testMonth.Days(1)
        'assert
        Assert.IsTrue(theDay.Shifts.Count > 0)
    End Sub

    <TestMethod()> Public Sub DocAvailCreated()
        'arrange
        'act
        Dim theshift As ScheduleShift
        Dim theDay As ScheduleDay
        MySettingsGlobal = New Settings1
        Dim testMonth As ScheduleMonth
        testMonth = New ScheduleMonth(1, 2014)

        theDay = testMonth.Days(1)
        theshift = theDay.Shifts(1)
        'assert
        Assert.IsTrue(theshift.DocAvailabilities.Count > 0)
    End Sub

End Class