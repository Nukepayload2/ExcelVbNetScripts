Imports ExcelDna.Integration
Imports ExcelDna.IntelliSense
Imports ExcelDna.Registration
Imports ExcelMoreScripts.NetDesktopExtension.Threading

Public Class MyAddIn
    Implements IExcelAddIn

    Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen
        ExcelRegistration.GetExcelFunctions().
            ProcessParamsRegistrations().
            ProcessAsyncRegistrations().
            RegisterFunctions()
        IntelliSenseServer.Install()
        STAThreadCache.CreateSTAThread()
        My.AddIn = Me
        CompilerHelper.BeginPreheating()
    End Sub

    Public Sub AutoClose() Implements IExcelAddIn.AutoClose
        My.AddIn = Nothing
        IntelliSenseServer.Uninstall()
        STAThreadCache.DisposeSTAThread()
    End Sub
End Class

Namespace My
    Module MyAddInGlobal
        Public Property AddIn As MyAddIn
    End Module

End Namespace
