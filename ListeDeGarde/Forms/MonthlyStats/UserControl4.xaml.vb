
Imports System.Windows.Controls
Imports System.Windows.Data
Public Class UserControl4


    Public Sub loadarray(aCollection As Collection)
        MyData.ItemsSource = Nothing
        MyData.ItemsSource = aCollection


    End Sub
End Class
