Imports ExcelDna.Integration

Public Module MyFunctions

    <ExcelFunction(Description:="Run Visual Basic .NET code snippet", Name:="VB.NET.TOPLEVEL")>
    Public Function RunVbNetTopLevel(code As String) As Object
        Dim vbCodeFull As String = $"
    Imports System

    Public Class Program
        Public Shared Function Main() As Object
{code}
        End Function
    End Class
"

        Dim returnValue = CompilerHelper.CompileAndRunVbCode(vbCodeFull, "Main", Array.Empty(Of Object)())
        Return returnValue
    End Function
End Module
