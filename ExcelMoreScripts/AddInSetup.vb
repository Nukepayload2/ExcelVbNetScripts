Imports ExcelDna.Integration
Imports ExcelDna.IntelliSense
Imports ExcelDna.Registration

Public Class IntelliSenseAddIn
    Implements IExcelAddIn

    Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen
        ExcelRegistration.GetExcelFunctions().
            ProcessParamsRegistrations().
            ProcessAsyncRegistrations().
            RegisterFunctions()
        IntelliSenseServer.Install()
    End Sub

    Public Sub AutoClose() Implements IExcelAddIn.AutoClose
        IntelliSenseServer.Uninstall()
    End Sub
End Class
