Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass>
Public Class CompilationTest
    <TestMethod>
    Sub FullSourceFile()
        Dim vbCodeSnippet As String = "
    Imports System

    Public Class Program
        Public Shared Function GetGreeting() As String
            Return ""Hello, World!""
        End Function

        Public Shared Function AddNumbers(a As Integer, b As Integer) As Integer
            Return a + b
        End Function
    End Class
"

        Dim returnValue1 = CompilerHelper.CompileAndRunVbCode(vbCodeSnippet, "GetGreeting", {})
        Console.WriteLine("Returned value (GetGreeting): " & returnValue1)

        Dim returnValue2 = CompilerHelper.CompileAndRunVbCode(vbCodeSnippet, "AddNumbers", {1, 2})
        Console.WriteLine("Returned value (AddNumbers): " & returnValue2)

        Dim vbCodeSnippet2 As String = "
    Imports System

    Public Class Program
        Public Shared Function GetGreeting() As String
            Return ""Hello, World!!!!""
        End Function

        Public Shared Function MulNumbers(a As Integer, b As Integer) As Integer
            Return a * b
        End Function
    End Class
"

        Dim returnValue3 = CompilerHelper.CompileAndRunVbCode(vbCodeSnippet2, "GetGreeting", {})
        Console.WriteLine("Returned value (GetGreeting): " & returnValue3)

        Dim returnValue4 = CompilerHelper.CompileAndRunVbCode(vbCodeSnippet2, "MulNumbers", {1, 2})
        Console.WriteLine("Returned value (AddNumbers): " & returnValue4)
    End Sub
End Class
