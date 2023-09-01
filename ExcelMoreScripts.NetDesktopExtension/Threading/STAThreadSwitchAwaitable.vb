Namespace Threading
    Public Structure STAThreadSwitchAwaitable
        Public Function GetAwaiter() As STAThreadSwitchAwaiter
            Return New STAThreadSwitchAwaiter()
        End Function
    End Structure
End Namespace
