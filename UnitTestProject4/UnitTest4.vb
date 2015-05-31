Imports System
Imports System.Text
Imports System.Collections.Generic
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports ListeDeGarde
Imports System.Runtime.Remoting
Imports System.Runtime.Remoting.Channels
Imports System.Runtime.Remoting.Channels.Tcp
Imports Microsoft.Office.Interop
Imports System.Diagnostics.Debug

<TestClass()> Public Class UnitTest4

    Public Shared aWorkbook As Excel.Workbook

    Public Shared workbookToTest As String = "C:\ExcelWorkbook1\ExcelWorkbook1\bin\Debug\ExcelWorkbook1.xls"
    Public Shared testMonth As SMonth
    Public Shared theDay As SDay
    Public Shared xlsApp As Excel.Application
    Public Shared xlsWB As Excel.Workbook
    Public Shared xlsSheet As Excel.Worksheet
    Public Shared aController As Controller

    <ClassInitialize()> Public Shared Sub MyClassInitialize(ByVal testContext As TestContext)
        

    End Sub
    <TestMethod()> Public Sub openup()


        xlsApp = New Excel.Application
        xlsApp.Visible = True
        xlsApp.Workbooks.Open(workbookToTest)
        Dim thecomaddins As Microsoft.Office.Core.COMAddIns = xlsApp.COMAddIns
        Dim theaddin As Microsoft.Office.Core.COMAddIn = thecomaddins.Item("ListeDeGarde")
        Dim theA As Object = theaddin.Object
        Dim astring As String = theA.getaddin()
        theA.launch()
        'Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("this has passed the test", thestring)
        'xlsApp.Quit()
        'Dim aCollection As System.Windows.Forms.Control.ControlCollection = ThisWorkbook.myCustomTaskPane.Control.Controls
        'Dim bElementHost As System.Windows.Forms.Integration.ElementHost = CType(aCollection(0), Windows.Forms.Integration.ElementHost)
        'Dim theUserCOntrol2 As UserControl2 = CType(bElementHost.Child, UserControl2)
        'theUserCOntrol2.TestClick()

        'System.Diagnostics.Debug.WriteLine("a test")
        ' Assert.IsTrue(Not ThisWorkbook.myCustomTaskPane Is Nothing)
        'Assert.IsTrue(False)

    End Sub



End Class
