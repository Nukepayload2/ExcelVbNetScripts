Imports ExcelMoreScripts.NetDesktopExtension.Threading

Public Class JoinableTaskFactory
    Public Shared Function SwitchToUIThreadAsync() As STAThreadSwitchAwaitable
        Return New STAThreadSwitchAwaitable()
    End Function

    Public Shared Function SwitchToWorkerThreadAsync() As WorkerThreadSwitchAwaitable
        Return New WorkerThreadSwitchAwaitable()
    End Function
End Class
