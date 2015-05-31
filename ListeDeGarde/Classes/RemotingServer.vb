Imports System.Runtime.Remoting

Imports System.Runtime.Remoting.Channels

Imports System.Runtime.Remoting.Channels.Tcp



'This class is to enable unit testing

Public Class RemotingServer



    Shared channel As TcpChannel



    Public Shared Sub Start()

        'pick any open channel number

        channel = New TcpChannel(8085)

        ChannelServices.RegisterChannel(channel, False)

        RemotingServices.Marshal(Globals.ThisAddIn.xlApp.ActiveWorkbook, "myCom")

    End Sub



    Public Shared Sub Unregister()

        ChannelServices.UnregisterChannel(channel)

    End Sub

End Class


