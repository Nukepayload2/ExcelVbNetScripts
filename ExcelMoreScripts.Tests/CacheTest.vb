Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass>
Public Class CacheTest

    <TestMethod>
    Public Sub TestLruCacheAccessThroughItem()
        Dim cache As New LruCache(Of String, String)(2)

        cache("key1") = "value1"
        cache("key2") = "value2"

        Assert.AreEqual("value1", cache("key1"))
        Assert.AreEqual("value2", cache("key2"))
        Assert.IsNull(cache("key3"))

        ' Remove key1
        cache("key3") = "value3"
        Assert.IsNull(cache("key1"))
        Assert.AreEqual("value2", cache("key2"))
        Assert.AreEqual("value3", cache("key3"))

    End Sub

    <TestMethod>
    Public Sub TestLruCacheTryGetValue()
        Dim cache As New LruCache(Of String, String)(2)

        cache("key1") = "value1"
        cache("key2") = "value2"

        Dim value1 As String = Nothing
        Assert.IsTrue(cache.TryGetValue("key1", value1))
        Assert.AreEqual("value1", value1)

        Dim value2 As String = Nothing
        Assert.IsTrue(cache.TryGetValue("key2", value2))
        Assert.AreEqual("value2", value2)

        Dim value3 As String = Nothing
        Assert.IsFalse(cache.TryGetValue("key3", value3))
        Assert.IsNull(value3)

        ' Remove key1
        cache("key3") = "value3"

        Assert.IsFalse(cache.TryGetValue("key1", value1))
        Assert.IsNull(value1)

        value2 = Nothing
        Assert.IsTrue(cache.TryGetValue("key2", value2))
        Assert.AreEqual("value2", value2)

        value3 = Nothing
        Assert.IsTrue(cache.TryGetValue("key3", value3))
        Assert.AreEqual("value3", value3)

    End Sub

End Class
