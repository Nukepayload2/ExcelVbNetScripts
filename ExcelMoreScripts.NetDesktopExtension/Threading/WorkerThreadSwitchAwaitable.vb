Namespace Threading
    Public Structure WorkerThreadSwitchAwaitable
        Public Function GetAwaiter() As WorkerThreadSwitchAwaiter
            Return New WorkerThreadSwitchAwaiter()
        End Function
    End Structure
End Namespace
