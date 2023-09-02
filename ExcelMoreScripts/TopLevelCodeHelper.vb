Public Class TopLevelCodeHelper

    Public Shared Function ExtractArgList(parameterNameAndValuePairs() As Object) As (argNameList As IEnumerable(Of String), argValueList As IEnumerable(Of Object))
        Dim argNameList = From arg In parameterNameAndValuePairs.Zip(Enumerable.Range(1, parameterNameAndValuePairs.Length))
                          Where arg.Second Mod 2 = 1 Select $"{arg.First} As Object"

        Dim argValueList = From arg In parameterNameAndValuePairs.Zip(Enumerable.Range(1, parameterNameAndValuePairs.Length))
                           Where arg.Second Mod 2 = 0 Select arg.First

        Return (argNameList, argValueList)
    End Function

    Public Shared Function SplitTopAndBodyLines(code As String) As (topLines As String, bodyLines As String)
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

End Class
