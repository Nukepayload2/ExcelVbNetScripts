Imports ExcelDna.Integration
Imports ExcelDna.Registration.Utils

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

        Dim returnValue = AsyncTaskUtil.RunAsTask(NameOf(RunVbNetTopLevel), {code},
             Function() CompilerHelper.CompileAndRunVbCode(vbCodeFull, "Main", Array.Empty(Of Object)()))

        Return returnValue
    End Function

    <ExcelFunction(Description:="Run Visual Basic .NET function", Name:="VB.NET.FUNCTION")>
    Public Function RunVbNetFunction(
        <ExcelArgument(Name:="code", Description:="Code that accepts function")>
        code As String,
        <ExcelArgument(Name:="parameterNameAndValuePair", Description:="paramName1, paramValue1, paramName2, paramValue2, ...")>
        ParamArray parameterNameAndValuePair As Object()) As Object

        If parameterNameAndValuePair.Length Mod 2 = 1 Then
            Return "ERROR: Incorrect parameter count"
        End If

        Dim argNameList = From arg In parameterNameAndValuePair.Zip(Enumerable.Range(1, parameterNameAndValuePair.Length))
                          Where arg.Second Mod 2 = 1 Select $"{arg.First} As Object"

        Dim argValueList = From arg In parameterNameAndValuePair.Zip(Enumerable.Range(1, parameterNameAndValuePair.Length))
                           Where arg.Second Mod 2 = 0 Select arg.First

        Dim vbCodeFull As String = $"
    Imports System

    Public Class Program
        Public Shared Function Main({String.Join(",", argNameList)}) As Object
{code}
        End Function
    End Class
"

        ' TODO: https://github.com/Excel-DNA/Registration/issues/25
        Dim returnValue = AsyncTaskUtil.RunAsTask(NameOf(RunVbNetFunction), {CObj(code), parameterNameAndValuePair},
            Function() CompilerHelper.CompileAndRunVbCode(vbCodeFull, "Main", argValueList.ToArray))

        Return returnValue
    End Function
End Module
