Imports System.Text.Json.Nodes
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass>
Public Class DemoTest
    Private Shared ReadOnly s_testJsonData As String = <![CDATA[
{
 "person" : {
  "name": "John Doe",
  "age": 30,
  "email": "johndoe@example.com",
  "interests": [
    "reading",
    "gaming",
    "cooking"
  ],
  "is_student": true
 }
}
]]>.Value

    Private Shared Function JsonGetProperty(jsonText As Object, propQuery As Object) As Object
        Dim json = JsonNode.Parse(CStr(jsonText))
        Dim queryString = CStr(propQuery)
        Dim curTokens As New List(Of JsonNode) From {json}
        For Each section In queryString.Split("."c)
            Dim flattenCurTokens = curTokens.OfType(Of JsonNode).ToList
            Dim flattenItems As Action(Of IEnumerable(Of JsonArray)) =
            Sub(arrs)
                For Each arr In arrs
                    flattenItems(arr.OfType(Of JsonArray))
                    For Each item In arr
                        If TypeOf item IsNot JsonArray Then
                            flattenCurTokens.Add(item)
                        End If
                    Next
                Next
            End Sub
            flattenItems(curTokens.OfType(Of JsonArray))
            Dim resultOfCurrentLayer As New List(Of JsonNode)
            For Each curToken In flattenCurTokens
                If TypeOf curToken Is JsonObject Then
                    For Each prop In DirectCast(curToken, JsonObject)
                        If prop.Key Like section Then
                            resultOfCurrentLayer.Add(prop.Value)
                        End If
                    Next
                End If
            Next
            curTokens = resultOfCurrentLayer
        Next

        Dim arrResult = Aggregate jVal In curTokens.OfType(Of JsonValue)
                        Select CObj(jVal.ToString) Into ToArray

        If arrResult.Length = 0 Then
            Dim unwrappedArrayResult =
                From jArr In curTokens.OfType(Of JsonArray), jVal In jArr.OfType(Of JsonValue)
                Select CObj(jVal.ToString)
            arrResult = arrResult.Concat(unwrappedArrayResult).ToArray
        End If

        Return arrResult
    End Function

    <TestMethod>
    Public Sub TestJsonGetProperty()
        Dim query = "person.*"
        Dim result = JsonGetProperty(s_testJsonData, query)
        Stop
    End Sub
End Class
