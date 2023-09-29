Imports System.Windows.Input

Public Class ErrorInformation
    Public Sub New(errorStage As String, title As String, information As String, longInformation As String, code As String)
        Me.ErrorStage = errorStage
        Me.Title = title
        Me.Information = information
        Me.LongInformation = longInformation
        Me.Code = code
    End Sub

    Public ReadOnly Property ErrorStage As String
    Public ReadOnly Property Title As String
    Public ReadOnly Property Information As String
    Public ReadOnly Property LongInformation As String
    Public ReadOnly Property Code As String
    Public ReadOnly Property Timestamp As Date = Now
End Class

Public Class ErrorStages
    Public Shared ReadOnly Property Compilation As String = "Compilation"
    Public Shared ReadOnly Property Runtime As String = "Runtime"
End Class
