Imports System.Collections.ObjectModel

Public Class ErrorViewModel
    Public Shared ReadOnly Property Instance As New ErrorViewModel

    Public ReadOnly Property Errors As New ObservableCollection(Of ErrorInformation)
End Class
