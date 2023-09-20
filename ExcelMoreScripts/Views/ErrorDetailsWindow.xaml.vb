Imports System.Windows

Public Class ErrorDetailsWindow
    Private Sub ErrorDetailsWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        TxtCode.Text = TryCast(DataContext, ErrorInformation)?.Code
    End Sub
End Class
