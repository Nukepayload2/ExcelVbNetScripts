Imports System.Threading
Imports System.Windows
Imports System.Windows.Threading

Namespace Threading
    Public Class STAThreadCache
        Private Shared _wpfApp As Application
        Private Shared _staThread As Thread

        Public Shared Sub CreateSTAThread()
            Using evt = New AutoResetEvent(False)
                _staThread = New Thread(
                        Sub()
                            _wpfApp = New Application With {.ShutdownMode = ShutdownMode.OnExplicitShutdown}
                            evt.Set()
                            _wpfApp.Run()
                        End Sub)
                _staThread.SetApartmentState(ApartmentState.STA)
                _staThread.Start()
                evt.WaitOne()
            End Using
        End Sub ' CreateSTAThread

        Public Shared ReadOnly Property DispatcherForSTAThread() As Dispatcher
            Get
                Return _wpfApp?.Dispatcher
            End Get
        End Property

        Public Shared Sub DisposeSTAThread()
            DispatcherForSTAThread?.Invoke(AddressOf _wpfApp.Shutdown)
        End Sub
    End Class  ' STAThreadCache
End Namespace
