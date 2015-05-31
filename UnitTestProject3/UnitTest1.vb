Imports System.Text
Imports ListeDeGarde

Imports System.Data.OleDb
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass()> Public Class UnitTest1

    <TestMethod()> Public Sub TestMethod1()
        Const Provider As String = "Provider=Microsoft.ACE.OLEDB.12.0;"
        'Const DBpassword = "Jet OLEDB:Database Password=plasma;"

        Dim theConnectionState As Long
        Dim mConnectionString As String
        Dim mConnection As ADODB.Connection

        If cnn.State = ADODB.ObjectStateEnum.adStateClosed Then

            mConnectionString = Provider + "Data Source=C:\Users\sel4_000\Documents\Scheduling Mira\Listedegarde.accdb"
            cnn.ConnectionString = mConnectionString
            cnn.Open()
        End If

        mConnection = cnn
        On Error GoTo 0
        Exit Sub


        'Dim adbac As DBAC
        'adbac = New DBAC()
    End Sub

End Class