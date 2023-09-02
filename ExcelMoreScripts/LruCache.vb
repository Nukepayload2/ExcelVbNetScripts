Imports System.Runtime.InteropServices

Public Class LruCache(Of TKey, TValue)
    Private ReadOnly _capacity As Integer
    Private ReadOnly _cache As New Dictionary(Of TKey, LinkedListNode(Of KeyValuePair(Of TKey, TValue)))()
    Private ReadOnly _lruList As New LinkedList(Of KeyValuePair(Of TKey, TValue))()

    Public Sub New(capacity As Integer)
        _capacity = capacity
    End Sub

    Public ReadOnly Property Capacity As Integer
        Get
            Return _capacity
        End Get
    End Property

    Public ReadOnly Property Count As Integer
        Get
            Return _lruList.Count
        End Get
    End Property

    Public Event ItemAutoRemoved(sender As LruCache(Of TKey, TValue), e As TValue)

    Default Public Property Item(key As TKey) As TValue
        Get
            Dim node As LinkedListNode(Of KeyValuePair(Of TKey, TValue)) = Nothing
            If Not _cache.TryGetValue(key, node) Then Return Nothing
            MoveToFront(node)
            Return node.Value.Value
        End Get
        Set(value As TValue)
            Dim node As LinkedListNode(Of KeyValuePair(Of TKey, TValue)) = Nothing
            If Not _cache.TryGetValue(key, node) Then
                If _lruList.Count >= _capacity Then
                    Dim removedItem = RemoveLast()
                    RaiseEvent ItemAutoRemoved(Me, removedItem.Value.Value)
                End If
                node = New LinkedListNode(Of KeyValuePair(Of TKey, TValue))(New KeyValuePair(Of TKey, TValue)(key, value))
                _cache.Add(key, node)
                _lruList.AddFirst(node)
            Else
                MoveToFront(node)
                node.Value = New KeyValuePair(Of TKey, TValue)(key, value)
            End If
        End Set
    End Property

    Private Sub MoveToFront(node As LinkedListNode(Of KeyValuePair(Of TKey, TValue)))
        _lruList.Remove(node)
        _lruList.AddFirst(node)
    End Sub

    Private Function RemoveLast() As LinkedListNode(Of KeyValuePair(Of TKey, TValue))
        Dim last As LinkedListNode(Of KeyValuePair(Of TKey, TValue)) = _lruList.Last
        _cache.Remove(last.Value.Key)
        _lruList.RemoveLast()
        Return last
    End Function

    Public Function TryGetValue(key As TKey, <Out> ByRef value As TValue) As Boolean
        Dim node As LinkedListNode(Of KeyValuePair(Of TKey, TValue)) = Nothing

        If _cache.TryGetValue(key, node) Then
            MoveToFront(node)
            value = node.Value.Value
            Return True
        End If

        value = Nothing
        Return False
    End Function
End Class
