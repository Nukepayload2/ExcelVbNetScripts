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
Imports System.Text.Json.Nodes
Imports System.Numerics"

    ' TODO: Automatically convert pointed cell to rich text and colorize with Roslyn
    <ExcelFunction(Description:="Run Visual Basic .NET function", Name:="VB.NET.FUNCTION")>
    Public Function RunVbNetFunction(
        <ExcelArgument(Name:="code", Description:="Code that is a function body. It must have a return statement.")>
        code As String,
        <ExcelArgument(Name:="parameterNameAndValuePair",
                       Description:="Optionally populates parameters for this function. Order: paramName1, paramValue1, paramName2, paramValue2, ...")>
        ParamArray parameterNameAndValuePairs As Object()) As Object

        If Not AskOrReuseCanRunCode() Then
            Return "#BLOCKED!"
        End If

        If parameterNameAndValuePairs.Length Mod 2 = 1 Then
            Return My.Resources.Resources.Error_IncorrectParameterCount
        End If

        Dim splittedCode = TopLevelCodeHelper.SplitTopAndBodyLines(code)

        Dim argList = TopLevelCodeHelper.ExtractArgList(parameterNameAndValuePairs)

        Dim vbCodeFull As String = $"{splittedCode.topLines}
{s_defaultImports}
Public Module {CompilerHelper.ScriptClassName}
    Public Function Main({String.Join(",", argList.argNameList)}) As Object
{splittedCode.bodyLines}
    End Function
{s_xlTypeHelper}
End Module"

        Dim returnValue = AsyncTaskUtil.RunAsTask(NameOf(RunVbNetFunction), {CObj(code), parameterNameAndValuePairs},
            Function() CompilerHelper.CompileAndRunVbCode(vbCodeFull, "Main", argList.argValueList.ToArray))

        Return returnValue
    End Function

    <ExcelFunction(Description:="Run Visual Basic .NET async function", Name:="VB.NET.ASYNC.FUNCTION")>
    Public Function RunVbNetAsyncFunction(
        <ExcelArgument(Name:="code", Description:="Code that is a async function body. It must have a return statement.")>
        code As String,
        <ExcelArgument(Name:="parameterNameAndValuePair",
                       Description:="Optionally populates parameters for this function. Order: paramName1, paramValue1, paramName2, paramValue2, ...")>
        ParamArray parameterNameAndValuePairs As Object()) As Object

        If Not AskOrReuseCanRunCode() Then
            Return "#BLOCKED!"
        End If

        If parameterNameAndValuePairs.Length Mod 2 = 1 Then
            Return My.Resources.Resources.Error_IncorrectParameterCount
        End If

        Dim splittedCode = TopLevelCodeHelper.SplitTopAndBodyLines(code)

        Dim argList = TopLevelCodeHelper.ExtractArgList(parameterNameAndValuePairs)

        Dim vbCodeFull As String = $"{splittedCode.topLines}
{s_defaultImports}
Public Module {CompilerHelper.ScriptClassName}
    Public Async Function MainAsync({String.Join(",", argList.argNameList)}) As Task(Of Object)
{splittedCode.bodyLines}
    End Function
{s_xlTypeHelper}
End Module
"

        Dim returnValue = AsyncTaskUtil.RunTask(NameOf(RunVbNetFunction), {CObj(code), parameterNameAndValuePairs},
            Async Function()
                Dim runResult = CompilerHelper.CompileAndRunVbCode(vbCodeFull, "MainAsync", argList.argValueList.ToArray)
                Dim taskResult = TryCast(runResult, Task(Of Object))
                If taskResult IsNot Nothing Then
                    Try
                        Return Await taskResult
                    Catch ex As Exception
                        HandleRuntimeError(ex, vbCodeFull)
                        Return CompilerHelper.FormatRuntimeException(ex)
                    End Try
                Else
                    Return runResult
                End If
            End Function)

        Return returnValue
    End Function

    Private ReadOnly s_xlTypeHelper As String = "    Private s_xlEmptyType As Type
    Private s_xlErrorType As Type

    Private Function CachedInstanceOf(value As Object, typeName As String, ByRef typeCache As Type) As Boolean
        If value Is Nothing Then Return False
        If typeCache Is Nothing Then
            Dim valType = value.GetType
            If valType.FullName = typeName Then
                typeCache = valType
                Return True
            End If
            Return False
        End If
        Return typeCache.IsInstanceOfType(value)
    End Function

    Private Function IsXlEmpty(value As Object) As Boolean
        Return CachedInstanceOf(value, ""ExcelDna.Integration.ExcelEmpty"", s_xlEmptyType)
    End Function

    Private Function IsXlError(value As Object) As Boolean
        Return CachedInstanceOf(value, ""ExcelDna.Integration.ExcelError"", s_xlErrorType)
    End Function"

    Private s_xlEmptyType As Type
    Private s_xlErrorType As Type

    Private Function CachedInstanceOf(value As Object, typeName As String, ByRef typeCache As Type) As Boolean
        If value Is Nothing Then Return False
        If typeCache Is Nothing Then
            Dim valType = value.GetType
            If valType.FullName = typeName Then
                typeCache = valType
                Return True
            End If
            Return False
        End If
        Return typeCache.IsInstanceOfType(value)
    End Function

    Private Function IsXlEmpty(value As Object) As Boolean
        Return CachedInstanceOf(value, "ExcelDna.Integration.ExcelEmpty", s_xlEmptyType)
    End Function

    Private Function IsXlError(value As Object) As Boolean
        Return CachedInstanceOf(value, "ExcelDna.Integration.ExcelError", s_xlErrorType)
    End Function
End Module
