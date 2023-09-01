Imports System.Runtime.CompilerServices
Imports System.Threading

Namespace Threading
    Public Class WorkerThreadSwitchAwaiter
        Implements INotifyCompletion

        Private _completed As Boolean

        Private _exception As Exception

        Public ReadOnly Property IsCompleted As Boolean
            Get
                Return Volatile.Read(_completed) OrElse
                    Thread.CurrentThread.IsThreadPoolThread
            End Get
        End Property

        Public Sub GetResult()
            If _exception IsNot Nothing Then
                Throw _exception
            End If
            If Not IsCompleted Then
                Throw New InvalidOperationException("Unable to switch to a worker thread.")
            End If
        End Sub

        Public Sub OnCompleted(continuation As Action) Implements INotifyCompletion.OnCompleted
            ThreadPool.QueueUserWorkItem(
            Sub()
                Try
                    continuation()
                Catch ex As Exception
                    _exception = ex
                End Try
                Volatile.Write(_completed, True)
            End Sub)
        End Sub
    End Class
End Namespace
