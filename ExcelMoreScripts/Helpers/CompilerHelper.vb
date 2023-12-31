﻿Imports System.IO
Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Runtime.Loader
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.VisualBasic

Public Class CompilerHelper
    Private Shared s_commonCompilation As VisualBasicCompilation

    ' TODO: Set max size from settings UI.
    Private Shared WithEvents CodeToAssemblyCache As New LruCache(Of String, Assembly)(100)

    Private Shared Sub CodeToAssemblyCache_ItemAutoRemoved(sender As LruCache(Of String, Assembly), e As Assembly) Handles CodeToAssemblyCache.ItemAutoRemoved
        Dim context = AssemblyLoadContext.GetLoadContext(e)
        If context.IsCollectible Then
            Debug.WriteLine("Unload assembly " & context.Name)
            context.Unload()
        End If
    End Sub

    Public Shared Function CompileAndRunVbCode(codeSnippet As String, methodName As String, args As Object()) As Object
        Dim failMessage As String = Nothing
        Dim cachedAssembly = CompileOrGetCachedAssembly(codeSnippet, failMessage)
        If cachedAssembly Is Nothing Then
            Return failMessage
        End If

        Dim entryPoint As MethodInfo = GetScriptMethodInAssembly(cachedAssembly, methodName)

        If entryPoint IsNot Nothing Then
            Try
                Dim returnValue = entryPoint.Invoke(Nothing, args)
                Return returnValue
            Catch ex As Exception
                HandleRuntimeError(ex, codeSnippet)
                Return FormatRuntimeException(ex)
            End Try
        Else
            Return $"Method '{methodName}' not found in the compiled assembly."
        End If
    End Function

    Public Shared Sub BeginPreheating()
        Task.Run(Sub()
                     Try
                         CompileAndRunVbCode("Module Program
Public Function Main() As Object
Return 0
End Function
End Module", "Main", Array.Empty(Of Object))
                     Catch ex As Exception
                     End Try
                 End Sub)
    End Sub

    Public Shared Function FormatRuntimeException(ex As Exception) As String
        If TypeOf ex Is TargetInvocationException AndAlso ex.InnerException IsNot Nothing Then
            ex = ex.InnerException
        End If
        Return $"{My.Resources.Resources.Error_RuntimeError} {ex.GetType.FullName}:{Environment.NewLine}{ex.Message}"
    End Function

    <MethodImpl(MethodImplOptions.Synchronized)>
    Private Shared Function CompileOrGetCachedAssembly(codeSnippet As String, ByRef failMessage As String) As Assembly
        Dim cachedAssembly As Assembly = Nothing
        If Not CodeToAssemblyCache.TryGetValue(codeSnippet, cachedAssembly) Then
            cachedAssembly = CompileAssembly(codeSnippet, failMessage)
            If cachedAssembly Is Nothing Then
                Return Nothing
            Else
                CodeToAssemblyCache(codeSnippet) = cachedAssembly
                Debug.WriteLine($"New assembly ({CodeToAssemblyCache.Count}/{CodeToAssemblyCache.Capacity}) " &
                                AssemblyLoadContext.GetLoadContext(cachedAssembly).Name)
            End If
        Else
            Debug.WriteLine($"Use cached assembly ({CodeToAssemblyCache.Count}/{CodeToAssemblyCache.Capacity}) " &
                            AssemblyLoadContext.GetLoadContext(cachedAssembly).Name)
        End If

        Return cachedAssembly
    End Function

    Private Shared Function CreateCompilation() As VisualBasicCompilation
        Dim commonAssemblyLocations =
            From dll In Directory.GetFiles(Path.GetDirectoryName(GetType(Object).Assembly.Location), "*.dll")
            Let fn = Path.GetFileNameWithoutExtension(dll)
            Where fn.Contains("System") OrElse fn.Contains("Microsoft") OrElse fn = "WindowsBase"
            Where Not fn.Contains("Native")
            Select dll

        Dim references = From loc In commonAssemblyLocations
                         Select MetadataReference.CreateFromFile(loc)
        Dim compilation = VisualBasicCompilation.Create($"ExcelVbNetDynamicAssembly{Guid.NewGuid:N}").
            WithOptions(New VisualBasicCompilationOptions(OutputKind.DynamicallyLinkedLibrary)).
            AddReferences(references)
        Return compilation
    End Function

    <MethodImpl(MethodImplOptions.Synchronized)>
    Private Shared Function CompileAssembly(codeSnippet As String, ByRef compilationFailedMessage As String) As Assembly
        Dim syntaxTree = VisualBasicSyntaxTree.ParseText(codeSnippet)

        If s_commonCompilation Is Nothing Then
            s_commonCompilation = CreateCompilation()
        Else
            s_commonCompilation = s_commonCompilation.RemoveAllSyntaxTrees()
        End If
        s_commonCompilation = s_commonCompilation.AddSyntaxTrees(syntaxTree)

        Dim memoryStream = New MemoryStream()
        Dim emitResult = s_commonCompilation.Emit(memoryStream)

        If emitResult.Success Then
            memoryStream.Position = 0
            Dim collectableAssemblyLoadContext As New AssemblyLoadContext($"script {Now:yyyy-MM-dd hh:mm:ss.FFFFFFF}", True)
            Return collectableAssemblyLoadContext.LoadFromStream(memoryStream)
        Else
            Dim errors = From diag In emitResult.Diagnostics
                         Where diag.Severity = DiagnosticSeverity.Error
            compilationFailedMessage =
                String.Join(Environment.NewLine,
                            From err In errors Select x = err.ToString)

            Dim firstError = errors.FirstOrDefault
            Dim errDesc = firstError?.Descriptor
            HandleCompilationError($"{firstError?.Id}: {errDesc?.Title}", compilationFailedMessage, codeSnippet)
        End If

        Return Nothing
    End Function

    Public Const ScriptClassName = "Program"

    Private Shared Function GetScriptMethodInAssembly(assembly As Assembly, methodName As String) As MethodInfo
        Dim t = assembly.GetType(ScriptClassName)
        Return t?.GetMethod(methodName)

        Return Nothing
    End Function
End Class
