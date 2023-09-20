Imports System.Reflection
Imports System.Windows.Interop
Imports ExcelDna.Integration
Imports ExcelMoreScripts.NetDesktopExtension

Module FormulaErrorHandler
    Sub HandleCompilationError(title As String, information As String, code As String)
        Dim errInfo As New ErrorInformation(ErrorStages.Compilation, title, information, information, code)
        ShowErrorWindow(errInfo)
    End Sub

    Sub HandleRuntimeError(ex As Exception, code As String)
        If TypeOf ex Is TargetInvocationException Then
            ex = ex.InnerException
        End If
        Dim errInfo As New ErrorInformation(ErrorStages.Runtime, ex.GetType.FullName, ex.Message, ex.ToString, code)
        ShowErrorWindow(errInfo)
    End Sub

    Private Async Sub ShowErrorWindow(errInfo As ErrorInformation)
        Await JoinableTaskFactory.SwitchToUIThreadAsync
        ErrorViewModel.Instance.Errors.Add(errInfo)
        Dim errWindow = My.Windows.ErrorWindow
        errWindow.Show()
        errWindow.ScrollToBottom()
        DialogParentHelper.ChangeParentWindow(New WindowInteropHelper(errWindow).Handle, ExcelDnaUtil.WindowHandle)
    End Sub

End Module
