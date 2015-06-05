Imports System
Imports System.Text
Imports System.Collections.Generic
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports ListeDeGarde
Imports System.Runtime.Remoting
Imports System.Runtime.Remoting.Channels
Imports System.Runtime.Remoting.Channels.Tcp
Imports Microsoft.Office.Interop


<TestClass()> Public NotInheritable Class VSTOFunctions
    Public Shared xlsApp As Excel.Application
    Public Shared theCOMConnection As Object

    <TestInitialize()> _
    Public Sub init()
        'System.Diagnostics.Debug.WriteLine("initfired")
        Dim workbookToTest As String = "C:\ExcelWorkbook1\ExcelWorkbook1\bin\Debug\ExcelWorkbook1.xls"
        xlsApp = New Excel.Application
        xlsApp.Visible = True
        xlsApp.Workbooks.Open(workbookToTest)
        Dim thecomaddins As Microsoft.Office.Core.COMAddIns = xlsApp.COMAddIns
        Dim theaddin As Microsoft.Office.Core.COMAddIn = thecomaddins.Item("ListeDeGarde")
        theCOMConnection = theaddin.Object
    End Sub

    <TestMethod()> _
    Public Sub VSTOFunctions_UseTaskPane()

        theCOMConnection.testclick()
        'System.Diagnostics.Debug.WriteLine("runtest")
        Assert.IsTrue(True)
    End Sub
    <TestCleanup()> _
    Public Sub clean()
        'System.Diagnostics.Debug.WriteLine("quit")
        xlsApp.Quit()
    End Sub

End Class
