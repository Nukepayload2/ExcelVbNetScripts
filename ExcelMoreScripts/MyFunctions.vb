Imports ExcelDna.Integration
Imports ExcelDna.Registration.Utils

Public Module MyFunctions

    ' TODO: Make a project UI
    Private ReadOnly s_defaultImports As String = "Imports Microsoft.VisualBasic
Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Linq
Imports System.Xml.Linq
Imports System.Threading.Tasks
' Added some additional imports based on Excel built-in functions
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Numerics"

    ' TODO: Automatically convert pointed cell to rich text and colorize with Roslyn
    <ExcelFunction(Description:="Run Visual Basic .NET function", Name:="VB.NET.FUNCTION")>
    Public Function RunVbNetFunction(
        <ExcelArgument(Name:="code", Description:="Code that is a function body. It must have a return statement.")>
        code As String,
        <ExcelArgument(Name:="parameterNameAndValuePair",
                       Description:="Optionally populates parameters for this function. Order: paramName1, paramValue1, paramName2, paramValue2, ...")>
        ParamArray parameterNameAndValuePairs As Object()) As Object

        If parameterNameAndValuePairs.Length Mod 2 = 1 Then
            Return "ERROR: Incorrect parameter count"
        End If

        Dim splittedCode = SplitTopAndBodyLines(code)

        Dim argList = ExtractArgList(parameterNameAndValuePairs)

        Dim vbCodeFull As String = $"{splittedCode.topLines}
{s_defaultImports}
    Public Class Program
        Public Shared Function Main({String.Join(",", argList.argNameList)}) As Object
{splittedCode.bodyLines}
        End Function
    End Class
"

        Dim returnValue = AsyncTaskUtil.RunAsTask(NameOf(RunVbNetFunction), {CObj(code), parameterNameAndValuePairs},
            Function() CompilerHelper.CompileAndRunVbCode(vbCodeFull, "Main", argList.argValueList.ToArray))

        Return returnValue
    End Function

    Private Function ExtractArgList(parameterNameAndValuePairs() As Object) As (argNameList As IEnumerable(Of String), argValueList As IEnumerable(Of Object))
        Dim argNameList = From arg In parameterNameAndValuePairs.Zip(Enumerable.Range(1, parameterNameAndValuePairs.Length))
                          Where arg.Second Mod 2 = 1 Select $"{arg.First} As Object"

        Dim argValueList = From arg In parameterNameAndValuePairs.Zip(Enumerable.Range(1, parameterNameAndValuePairs.Length))
                           Where arg.Second Mod 2 = 0 Select arg.First

        Return (argNameList, argValueList)
    End Function

    <ExcelFunction(Description:="Run Visual Basic .NET async function", Name:="VB.NET.ASYNC.FUNCTION")>
    Public Function RunVbNetAsyncFunction(
        <ExcelArgument(Name:="code", Description:="Code that is a async function body. It must have a return statement.")>
        code As String,
        <ExcelArgument(Name:="parameterNameAndValuePair",
                       Description:="Optionally populates parameters for this function. Order: paramName1, paramValue1, paramName2, paramValue2, ...")>
        ParamArray parameterNameAndValuePairs As Object()) As Object

        If parameterNameAndValuePairs.Length Mod 2 = 1 Then
            Return "ERROR: Incorrect parameter count"
        End If

        Dim splittedCode = SplitTopAndBodyLines(code)

        Dim argList = ExtractArgList(parameterNameAndValuePairs)

        Dim vbCodeFull As String = $"{splittedCode.topLines}
{s_defaultImports}
    Public Class Program
        Public Shared Async Function MainAsync({String.Join(",", argList.argNameList)}) As Task(Of Object)
{splittedCode.bodyLines}
        End Function
    End Class
"

        Dim returnValue = AsyncTaskUtil.RunTask(NameOf(RunVbNetFunction), {CObj(code), parameterNameAndValuePairs},
            Async Function()
                Dim runResult = CompilerHelper.CompileAndRunVbCode(vbCodeFull, "MainAsync", argList.argValueList.ToArray)
                Dim taskResult = TryCast(runResult, Task(Of Object))
                If taskResult IsNot Nothing Then
                    Return Await taskResult
                Else
                    Return runResult
                End If
            End Function)

        Return returnValue
    End Function

    Private Function SplitTopAndBodyLines(code As String) As (topLines As String, bodyLines As String)
        Dim splitted = code.Split(vbLf)
        Dim topLines =
            (From ln In splitted
             Let trimmed = ln.TrimStart
             Take While String.IsNullOrEmpty(trimmed) OrElse ' Has nothing
                        trimmed.StartsWith("'"c) OrElse ' Comment
                        trimmed.StartsWith("REM ", StringComparison.OrdinalIgnoreCase) OrElse ' Comment
                        trimmed.StartsWith("Imports ", StringComparison.OrdinalIgnoreCase) OrElse
                        trimmed.StartsWith("Option ", StringComparison.OrdinalIgnoreCase)
             Select ln).ToArray

        Dim bodyLines = String.Join(Environment.NewLine, splitted.Skip(topLines.Length))
        Return (String.Join(Environment.NewLine, topLines), bodyLines)
    End Function

End Module
