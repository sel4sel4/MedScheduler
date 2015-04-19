Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports ListeDeGarde
Imports System.Diagnostics

<TestClass()> Public Class classes

    <TestMethod()> Public Sub MonthCreated()
        'arrange
        'act
        MySettingsGlobal = New Settings1
        Dim testMonth As ScheduleMonth
        testMonth = New ScheduleMonth(1, 2014)
        'assert
        Assert.IsTrue(testMonth.Days.Count > 0)
    End Sub

    <TestMethod()> Public Sub DaysCreated()
        'arrange
        'act
        Dim theDay As ScheduleDay
        MySettingsGlobal = New Settings1
        Dim testMonth As ScheduleMonth
        testMonth = New ScheduleMonth(1, 2014)

        theDay = testMonth.Days(1)
        'assert
        Assert.IsTrue(theDay.Shifts.Count > 0)
    End Sub

End Class