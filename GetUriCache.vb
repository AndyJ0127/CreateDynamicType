Imports System.Runtime.Caching
Imports Newtonsoft.Json
Imports System.Reflection
Imports System.Reflection.Emit
Imports System.Collections.Generic
Imports System.Threading

Public Class GetUriCache

    '非同步 IEnumerable(Of T) Method=========================================================
    Public Async Function GetJsonCacheAsync(Of T)(cacheName As String, targetURI As String, cacheMinTime As Integer) As Tasks.Task(Of Object)

        Dim cache As ObjectCache = MemoryCache.Default

        Dim cacheContents As CacheItem = Nothing

        Try
            cacheContents = cache.GetCacheItem(cacheName)

            Try
                If (cacheContents Is Nothing) = True Then
                    Return Await Me.GetJsonAsync(Of T)(cacheName:=cacheName, targetURI:=targetURI, cacheMinTime:=cacheMinTime)
                Else
                    Return cacheContents.Value
                End If
            Catch ex As Exception
                HttpContext.Current.Response.Write("GetJsonCacheAsync =" & ex.Message & "<BR>")
            End Try

        Catch ex As Exception
            cache.Remove(key:=cacheName)

            HttpContext.Current.Response.Write("GetJsonCacheAsync cacheContents=" & ex.Message & "<BR>")
        Finally

        End Try

        Return Await Me.GetJsonAsync(Of T)(cacheName:=cacheName, targetURI:=targetURI, cacheMinTime:=cacheMinTime)

    End Function

    Private Async Function GetJsonAsync(Of T)(cacheName As String, targetURI As String, cacheMinTime As Integer) As Tasks.Task(Of T)

        Dim httpClient As New System.Net.Http.HttpClient()

        httpClient.MaxResponseContentBufferSize = Int32.MaxValue

        Dim response = Nothing

        Dim collection As T = Nothing

        Dim policy As New CacheItemPolicy()

        Dim cacheItem As ObjectCache = MemoryCache.Default

        Try
            httpClient.Timeout = TimeSpan.FromSeconds(5)
            response = Await httpClient.GetStringAsync(targetURI)

            collection = JsonConvert.DeserializeObject(Of T)(response)

            policy.AbsoluteExpiration = DateTime.Now.AddMinutes(cacheMinTime)
            'RunTime Cache的保留時間為(分鐘)

            cacheItem.Add(key:=cacheName, value:=collection, policy:=policy)

        Catch ex As Exception
            cacheItem.Remove(key:=cacheName)

            collection = JsonConvert.DeserializeObject(Of T)(response)

            HttpContext.Current.Response.Write("GetJsonAsync=" & ex.Message & "<BR>")
        Finally

        End Try

        httpClient.Dispose()

        Return collection

    End Function
    '非同步 IEnumerable(Of T) Method=========================================================



    '同步 IEnumerable(Of T) Method=========================================================
    Public Function GetJsonCache(Of T)(cacheName As String, targetURI As String, cacheMinTime As Integer) As Object

        Dim cache As ObjectCache = MemoryCache.Default

        Dim cacheContents As CacheItem = Nothing

        Try
            cacheContents = cache.GetCacheItem(cacheName)

            Try
                If (cacheContents Is Nothing) = True Then
                    Return Me.GetJson(Of T)(cacheName:=cacheName, targetURI:=targetURI, cacheMinTime:=cacheMinTime)
                Else
                    Return cacheContents.Value
                End If
            Catch ex As Exception
                HttpContext.Current.Response.Write("GetJsonCache =" & ex.Message & "<BR>")
            End Try

        Catch ex As Exception
            HttpContext.Current.Response.Write("GetJsonCache cacheContents=" & ex.Message & "<BR>")
        End Try

        Return Me.GetJson(Of T)(cacheName:=cacheName, targetURI:=targetURI, cacheMinTime:=cacheMinTime)

    End Function

    Private Function GetJson(Of T)(cacheName As String, targetURI As String, cacheMinTime As Integer) As T

        Dim webClient As New System.Net.WebClient()

        webClient.Encoding = System.Text.Encoding.UTF8


        Dim response As String = webClient.DownloadString(targetURI)

        Dim collection As T = JsonConvert.DeserializeObject(Of T)(response)

        Dim policy As New CacheItemPolicy()

        Dim cacheItem As ObjectCache = MemoryCache.Default

        Try
            policy.AbsoluteExpiration = DateTime.Now.AddMinutes(cacheMinTime)
            'RunTime Cache的保留時間為(分鐘)
            cacheItem.Add(key:=cacheName, value:=collection, policy:=policy)

        Catch ex As Exception
            cacheItem.Remove(key:=cacheName)

            collection = JsonConvert.DeserializeObject(Of T)(response)

            HttpContext.Current.Response.Write("GetJson=" & ex.Message & "<BR>")
        End Try

        webClient.Dispose()

        Return collection

    End Function
    '同步 IEnumerable(Of T) Method=========================================================

End Class
