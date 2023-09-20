Imports System.Windows.Input
Imports System.Windows.Interop
Imports ExcelDna.Integration

Public Class ViewErrorDetailsCommand
    Implements ICommand

    Public Shared ReadOnly Property Instance As New ViewErrorDetailsCommand

    Public Event CanExecuteChanged As EventHandler Implements ICommand.CanExecuteChanged

    Public Sub Execute(parameter As Object) Implements ICommand.Execute
        Dim errInfo = TryCast(parameter, ErrorInformation)
        If errInfo Is Nothing Then Return
        Dim detailsWindow = My.Windows.ErrorDetailsWindow
        detailsWindow.DataContext = errInfo
        detailsWindow.Show()
        DialogParentHelper.ChangeParentWindow(New WindowInteropHelper(detailsWindow).Handle, ExcelDnaUtil.WindowHandle)
    End Sub

    Public Function CanExecute(parameter As Object) As Boolean Implements ICommand.CanExecute
        Return True
    End Function
End Class
