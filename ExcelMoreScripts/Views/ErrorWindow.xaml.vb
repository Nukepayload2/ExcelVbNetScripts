Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media

Public Class ErrorWindow
    Private Sub ErrorWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        DataContext = ErrorViewModel.Instance
    End Sub

    Private Sub BtnClearError_Click(sender As Object, e As RoutedEventArgs) Handles BtnClearError.Click
        ErrorViewModel.Instance.Errors.Clear()
    End Sub

    Public Sub ScrollToBottom()
        If VisualTreeHelper.GetChildrenCount(LstErrors) > 0 Then
            Dim border = DirectCast(VisualTreeHelper.GetChild(LstErrors, 0), Border)
            Dim scrollViewer = DirectCast(VisualTreeHelper.GetChild(border, 0), ScrollViewer)
            scrollViewer.ScrollToBottom()
        End If
    End Sub
End Class
