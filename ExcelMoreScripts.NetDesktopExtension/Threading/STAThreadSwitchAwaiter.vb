Imports System.Runtime.CompilerServices
Imports System.Threading

Namespace Threading
    Public Class STAThreadSwitchAwaiter
        Implements INotifyCompletion

        Private _completed As Boolean

        Private _exception As Exception

        Public ReadOnly Property IsCompleted As Boolean
            Get
                Return Volatile.Read(_completed) OrElse
                    Thread.CurrentThread.GetApartmentState() = ApartmentState.STA
            End Get
        End Property

        Public Sub GetResult()
            If _exception IsNot Nothing Then
                Throw _exception
            End If
            If Not IsCompleted Then
                Throw New InvalidOperationException("Unable to switch to a STA thread.")
            End If
        End Sub

        Public Sub OnCompleted(continuation As Action) Implements INotifyCompletion.OnCompleted
            STAThreadCache.DispatcherForSTAThread.BeginInvoke(
            Sub()
                Try
                    continuation?.Invoke()
                Catch ex As Exception
                    _exception = ex
                End Try
                Volatile.Write(_completed, value:=True)
            End Sub)
        End Sub
    End Class
End Namespace
