Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports ListeDeGarde

<TestClass()> Public Class classes

    <TestMethod()> Public Sub TestMethod1()
        'arrange
        'act
        Dim testMonth As New ScheduleMonth(1, 2014)
        'assert
        Assert.IsTrue(testMonth.Days.Count > 0)
    End Sub

End Class