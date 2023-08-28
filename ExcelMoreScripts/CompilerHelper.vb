Imports System.IO
Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.VisualBasic

Public Class CompilerHelper
    Private Shared cachedAssembly As Assembly = Nothing
    Private Shared s_lastCodeSnippet As String

    <MethodImpl(MethodImplOptions.Synchronized)>
    Public Shared Function CompileAndRunVbCode(codeSnippet As String, methodName As String, args As Object()) As Object
        If cachedAssembly Is Nothing OrElse s_lastCodeSnippet <> codeSnippet Then
            Dim failMessage As String = Nothing
            cachedAssembly = CompileAssembly(codeSnippet, failMessage)
            If failMessage IsNot Nothing Then
                Return failMessage
            End If
            s_lastCodeSnippet = codeSnippet
        End If

        If cachedAssembly IsNot Nothing Then
            Dim entryPoint As MethodInfo = FindMethodInAssembly(cachedAssembly, methodName)

            If entryPoint IsNot Nothing Then
                Dim declaringTypeInstance = Activator.CreateInstance(entryPoint.DeclaringType)
                Dim returnValue = entryPoint.Invoke(declaringTypeInstance, args)
                Return returnValue
            Else
                Console.WriteLine($"Method '{methodName}' not found in the compiled assembly.")
            End If
        Else
            Console.WriteLine("Compilation failed.")
        End If

        Return Nothing
    End Function

    Private Shared s_commonCompilation As VisualBasicCompilation
    Private Shared Function CreateCompilation() As VisualBasicCompilation
        Dim systemReferences = AppDomain.CurrentDomain.GetAssemblies() _
                            .Where(Function(assembly) assembly.FullName.StartsWith("System") OrElse assembly.FullName.StartsWith("Microsoft")) _
                            .Select(Function(assembly) MetadataReference.CreateFromFile(assembly.Location))
        Dim references = systemReferences
        Dim compilation = VisualBasicCompilation.Create("ExcelVbNetDynamicAssembly") _
                                                .WithOptions(New VisualBasicCompilationOptions(OutputKind.DynamicallyLinkedLibrary)) _
                                                .AddReferences(references)
        Return compilation
    End Function

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