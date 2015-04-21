Public Class UserControl4


    Public Sub loadarray(anArray As List(Of List(Of Integer)))

        MyData.ItemsSource = anArray
        MyData2.ItemsSource = anArray

    End Sub
End Class
