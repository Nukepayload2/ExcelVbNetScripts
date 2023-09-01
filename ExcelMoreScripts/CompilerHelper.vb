Imports System.IO
Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.VisualBasic

Public Class CompilerHelper
    Private Shared s_commonCompilation As VisualBasicCompilation

    ' TODO: Clean this cache when the key is not in use.
    Private Shared ReadOnly s_codeToAssemblyCache As New Dictionary(Of String, Assembly)

    Public Shared Function CompileAndRunVbCode(codeSnippet As String, methodName As String, args As Object()) As Object
        Dim failMessage As String = Nothing
        Dim cachedAssembly = CompileOrGetCachedAssembly(codeSnippet, failMessage)
        If cachedAssembly Is Nothing Then
            Return failMessage
        End If

        Dim entryPoint As MethodInfo = FindMethodInAssembly(cachedAssembly, methodName)

        If entryPoint IsNot Nothing Then
            Dim declaringTypeInstance = Activator.CreateInstance(entryPoint.DeclaringType)
            Try
                Dim returnValue = entryPoint.Invoke(declaringTypeInstance, args)
                Return returnValue
            Catch ex As Exception
                Return FormatRuntimeException(ex)
            End Try
        Else
            Return $"Method '{methodName}' not found in the compiled assembly."
        End If
    End Function

    Public Shared Function FormatRuntimeException(ex As Exception) As String
        If TypeOf ex Is TargetInvocationException AndAlso ex.InnerException IsNot Nothing Then
            ex = ex.InnerException
        End If
        Return $"Runtime Error {ex.GetType.FullName}:{Environment.NewLine}{ex.Message}"
    End Function

    <MethodImpl(MethodImplOptions.Synchronized)>
    Private Shared Function CompileOrGetCachedAssembly(codeSnippet As String, ByRef failMessage As String) As Assembly
        Dim cachedAssembly As Assembly = Nothing
        If Not s_codeToAssemblyCache.TryGetValue(codeSnippet, cachedAssembly) Then
            cachedAssembly = CompileAssembly(codeSnippet, failMessage)
            If cachedAssembly Is Nothing Then
                Return Nothing
            Else
                s_codeToAssemblyCache(codeSnippet) = cachedAssembly
            End If
        End If

        Return cachedAssembly
    End Function

    Private Shared Function CreateCompilation() As VisualBasicCompilation
        LoadCommonlyUsedLibs()
        Dim systemReferences =
            From asm In AppDomain.CurrentDomain.GetAssemblies()
            Where Not asm.IsDynamic
            Let loc = asm.Location
            Where File.Exists(loc)
            Select MetadataReference.CreateFromFile(asm.Location)
        Dim references = systemReferences
        Dim compilation = VisualBasicCompilation.Create($"ExcelVbNetDynamicAssembly{Guid.NewGuid:N}").
            WithOptions(New VisualBasicCompilationOptions(OutputKind.DynamicallyLinkedLibrary)).
            AddReferences(references)
        Return compilation
    End Function

    Private Shared Sub LoadCommonlyUsedLibs()
        ' XLinq
        RuntimeHelpers.RunClassConstructor(GetType(System.Xml.Linq.XElement).TypeHandle)
        ' Newtonsoft.Json
        RuntimeHelpers.RunClassConstructor(GetType(Newtonsoft.Json.JsonConvert).TypeHandle)
    End Sub

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
            Return Assembly.Load(memoryStream.ToArray())
        Else
            compilationFailedMessage = String.Join(Environment.NewLine,
                                                   From diag In emitResult.Diagnostics Select x = diag.ToString)
        End If

        Return Nothing
    End Function

    Private Shared Function FindMethodInAssembly(assembly As Assembly, methodName As String) As MethodInfo
        For Each t In assembly.GetTypes()
            Dim method = t.GetMethod(methodName)
            If method IsNot Nothing Then
                Return method
            End If
        Next

        Return Nothing
    End Function
End Class