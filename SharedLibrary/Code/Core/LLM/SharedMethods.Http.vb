' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: SharedMethods.Http.vb
' Purpose:
'   Shared HTTP transport abstraction used by higher-level Red Ink services for text and
'   binary requests without coupling callers to one .NET networking stack.
'
' Architecture / Function:
'   - Defines normalized request/response DTOs and HttpStackPreference routing.
'   - Supports HttpClient and WinHTTP execution with configured ordering/fallback,
'     headers, request bodies, compression handling, timeouts and response decoding.
'   - Keeps transport selection and compatibility fallback in one place; authentication,
'     API-specific payloads and semantic retry policy remain with the calling service.
'   - Callers should treat returned bytes/text as untrusted remote content.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports System.IO.Compression
Imports System.Net
Imports System.Net.Http
Imports System.Text
Imports System.Threading.Tasks

Namespace SharedLibrary
    Partial Public Class SharedMethods

        Public Shared INI_HttpStack_Cached As String = ""

        Public Const Disable_HttpClient As Boolean = False

        Public Enum HttpStackPreference
            PreferConfiguredDefault = 0
            HttpClientOnly = 1
            WinHttpOnly = 2
            PreferHttpClientWithWinHttpFallback = 3
            PreferWinHttpWithHttpClientFallback = 4
        End Enum

        Public Class SharedHttpRequest
            Public Property Url As String = ""
            Public Property Method As String = "GET"
            Public Property TimeoutMs As Integer = 30000
            Public Property UserAgent As String = ""
            Public Property Accept As String = ""
            Public Property ContentType As String = ""
            Public Property BodyText As String = ""
            Public Property BodyBytes As Byte()
            Public Property Headers As Dictionary(Of String, String)
            Public Property StackPreference As HttpStackPreference = HttpStackPreference.PreferConfiguredDefault
        End Class

        Public Class SharedHttpResponse
            Public Property StatusCode As Integer
            Public Property ReasonPhrase As String = ""
            Public Property ContentType As String = ""
            Public Property CharSet As String = ""
            Public Property BodyBytes As Byte()
            Public Property UsedStack As String = ""

            Public Function GetBodyAsString() As String
                Dim data = If(BodyBytes, New Byte() {})
                Dim enc As Encoding = Encoding.UTF8

                If Not String.IsNullOrWhiteSpace(CharSet) Then
                    Try
                        enc = Encoding.GetEncoding(CharSet)
                    Catch
                        enc = Encoding.UTF8
                    End Try
                End If

                Return enc.GetString(data)
            End Function
        End Class

        Public Shared Async Function SendHttpRequestAsync(request As SharedHttpRequest) As Task(Of SharedHttpResponse)
            If request Is Nothing Then
                Throw New ArgumentNullException(NameOf(request))
            End If

            If String.IsNullOrWhiteSpace(request.Url) Then
                Throw New ArgumentException("Request URL must not be empty.", NameOf(request))
            End If

            EnsureModernTlsForSharedHttp()

            Dim stackOrder = GetSharedHttpStackOrder(request.StackPreference)
            Dim lastException As Exception = Nothing

            For Each stack As String In stackOrder
                Try
                    Select Case stack
                        Case "HttpClient"
                            Return Await SendHttpRequestWithHttpClientAsync(request).ConfigureAwait(False)

                        Case "WinHTTP"
                            Return Await Task.Run(Function() SendHttpRequestWithWinHttp(request)).ConfigureAwait(False)

                        Case Else
                            Throw New InvalidOperationException($"Unsupported HTTP stack: {stack}")
                    End Select
                Catch ex As Exception
                    lastException = New Exception($"{stack}: {ex.Message}", ex)
                End Try
            Next

            Throw New Exception(
                $"All configured HTTP stacks failed. Last error: {If(lastException IsNot Nothing, lastException.Message, "Unknown error")}",
                lastException)
        End Function

        Private Shared Function GetSharedHttpStackOrder(preference As HttpStackPreference) As List(Of String)
            If Disable_HttpClient Then
                Return New List(Of String) From {"WinHTTP"}
            End If

            Dim effectivePreference = preference
            If effectivePreference = HttpStackPreference.PreferConfiguredDefault Then
                effectivePreference = GetConfiguredSharedHttpStackPreference()
            End If

            Select Case effectivePreference
                Case HttpStackPreference.HttpClientOnly
                    Return New List(Of String) From {"HttpClient"}

                Case HttpStackPreference.WinHttpOnly
                    Return New List(Of String) From {"WinHTTP"}

                Case HttpStackPreference.PreferWinHttpWithHttpClientFallback
                    Return New List(Of String) From {"WinHTTP", "HttpClient"}

                Case Else
                    Return New List(Of String) From {"HttpClient", "WinHTTP"}
            End Select
        End Function

        Private Shared Function GetConfiguredSharedHttpStackPreference() As HttpStackPreference
            Dim configuredValue = SharedMethods.INI_HttpStack_Cached

            Debug.WriteLine($"Configured HTTP stack preference: {If(configuredValue, "Not set")}")

            If String.IsNullOrWhiteSpace(configuredValue) Then
                Return HttpStackPreference.PreferHttpClientWithWinHttpFallback
            End If

            Select Case configuredValue.Trim().ToLowerInvariant()
                Case "httpclient"
                    Return HttpStackPreference.HttpClientOnly

                Case "winhttp"
                    Return HttpStackPreference.WinHttpOnly

                Case "prefer-winhttp", "fallback"
                    Return HttpStackPreference.PreferWinHttpWithHttpClientFallback

                Case "prefer-httpclient", "default"
                    Return HttpStackPreference.PreferHttpClientWithWinHttpFallback

                Case Else
                    Return HttpStackPreference.PreferHttpClientWithWinHttpFallback
            End Select
        End Function


        Private Shared Sub EnsureModernTlsForSharedHttp()
            Try
                ServicePointManager.SecurityProtocol =
                    ServicePointManager.SecurityProtocol Or
                    SecurityProtocolType.Tls12 Or
                    CType(12288, SecurityProtocolType)
            Catch
                Try
                    ServicePointManager.SecurityProtocol =
                        ServicePointManager.SecurityProtocol Or
                        SecurityProtocolType.Tls12
                Catch
                End Try
            End Try
        End Sub

        Private Shared Async Function SendHttpRequestWithHttpClientAsync(request As SharedHttpRequest) As Task(Of SharedHttpResponse)
            Dim handler As New HttpClientHandler() With {
                .AllowAutoRedirect = True,
                .AutomaticDecompression = DecompressionMethods.GZip Or DecompressionMethods.Deflate
            }

            Using client As New HttpClient(handler)
                client.Timeout = TimeSpan.FromMilliseconds(request.TimeoutMs)

                If Not String.IsNullOrWhiteSpace(request.UserAgent) Then
                    client.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", request.UserAgent)
                End If

                If Not String.IsNullOrWhiteSpace(request.Accept) Then
                    client.DefaultRequestHeaders.TryAddWithoutValidation("Accept", request.Accept)
                End If

                If request.Headers IsNot Nothing Then
                    For Each kvp In request.Headers
                        client.DefaultRequestHeaders.TryAddWithoutValidation(kvp.Key, kvp.Value)
                    Next
                End If

                Using message As New HttpRequestMessage(New HttpMethod(request.Method), request.Url)
                    Dim content = CreateSharedHttpContent(request)
                    If content IsNot Nothing Then
                        message.Content = content
                    End If

                    Using response = Await client.SendAsync(message, HttpCompletionOption.ResponseContentRead).ConfigureAwait(False)
                        Dim result As New SharedHttpResponse() With {
                            .StatusCode = CInt(response.StatusCode),
                            .ReasonPhrase = If(response.ReasonPhrase, ""),
                            .BodyBytes = Await response.Content.ReadAsByteArrayAsync().ConfigureAwait(False),
                            .UsedStack = "HttpClient"
                        }

                        If response.Content IsNot Nothing AndAlso
                           response.Content.Headers IsNot Nothing AndAlso
                           response.Content.Headers.ContentType IsNot Nothing Then

                            result.ContentType = If(response.Content.Headers.ContentType.MediaType, "")
                            result.CharSet = If(response.Content.Headers.ContentType.CharSet, "")
                        End If

                        Return result
                    End Using
                End Using
            End Using
        End Function

        Private Shared Function SendHttpRequestWithWinHttp(request As SharedHttpRequest) As SharedHttpResponse
            Dim requestType = Type.GetTypeFromProgID("WinHttp.WinHttpRequest.5.1")
            If requestType Is Nothing Then
                Throw New InvalidOperationException("WinHTTP COM component 'WinHttp.WinHttpRequest.5.1' is not available.")
            End If

            Dim comRequest = Activator.CreateInstance(requestType)
            If comRequest Is Nothing Then
                Throw New InvalidOperationException("Failed to create WinHTTP request instance.")
            End If

            Try
                requestType.InvokeMember(
                    "Open",
                    System.Reflection.BindingFlags.InvokeMethod,
                    Nothing,
                    comRequest,
                    New Object() {request.Method, request.Url, False})

                requestType.InvokeMember(
                    "SetTimeouts",
                    System.Reflection.BindingFlags.InvokeMethod,
                    Nothing,
                    comRequest,
                    New Object() {request.TimeoutMs, request.TimeoutMs, request.TimeoutMs, request.TimeoutMs})

                If Not String.IsNullOrWhiteSpace(request.UserAgent) Then
                    requestType.InvokeMember(
                        "SetRequestHeader",
                        System.Reflection.BindingFlags.InvokeMethod,
                        Nothing,
                        comRequest,
                        New Object() {"User-Agent", request.UserAgent})
                End If

                If Not String.IsNullOrWhiteSpace(request.Accept) Then
                    requestType.InvokeMember(
                        "SetRequestHeader",
                        System.Reflection.BindingFlags.InvokeMethod,
                        Nothing,
                        comRequest,
                        New Object() {"Accept", request.Accept})
                End If

                If request.Headers IsNot Nothing Then
                    For Each kvp In request.Headers
                        requestType.InvokeMember(
                            "SetRequestHeader",
                            System.Reflection.BindingFlags.InvokeMethod,
                            Nothing,
                            comRequest,
                            New Object() {kvp.Key, kvp.Value})
                    Next
                End If

                If Not String.IsNullOrWhiteSpace(request.ContentType) Then
                    requestType.InvokeMember(
                        "SetRequestHeader",
                        System.Reflection.BindingFlags.InvokeMethod,
                        Nothing,
                        comRequest,
                        New Object() {"Content-Type", request.ContentType})
                End If

                Try
                    requestType.InvokeMember(
                        "Option",
                        System.Reflection.BindingFlags.SetProperty,
                        Nothing,
                        comRequest,
                        New Object() {9, 512 Or 2048})
                Catch
                End Try

                Dim sendBody As Object = Nothing
                If request.BodyBytes IsNot Nothing AndAlso request.BodyBytes.Length > 0 Then
                    sendBody = request.BodyBytes
                ElseIf Not String.IsNullOrEmpty(request.BodyText) Then
                    sendBody = request.BodyText
                End If

                requestType.InvokeMember(
                    "Send",
                    System.Reflection.BindingFlags.InvokeMethod,
                    Nothing,
                    comRequest,
                    If(sendBody Is Nothing, Nothing, New Object() {sendBody}))

                Dim statusCode = System.Convert.ToInt32(
                    requestType.InvokeMember(
                        "Status",
                        System.Reflection.BindingFlags.GetProperty,
                        Nothing,
                        comRequest,
                        Nothing),
                    System.Globalization.CultureInfo.InvariantCulture)

                Dim reasonPhrase = System.Convert.ToString(
                    requestType.InvokeMember(
                        "StatusText",
                        System.Reflection.BindingFlags.GetProperty,
                        Nothing,
                        comRequest,
                        Nothing),
                    System.Globalization.CultureInfo.InvariantCulture)

                Dim rawHeaders = System.Convert.ToString(
                    requestType.InvokeMember(
                        "GetAllResponseHeaders",
                        System.Reflection.BindingFlags.InvokeMethod,
                        Nothing,
                        comRequest,
                        Nothing),
                    System.Globalization.CultureInfo.InvariantCulture)

                Dim headers = ParseSharedHttpHeaders(rawHeaders)
                Dim bodyBytes = ConvertWinHttpResponseBodyToByteArray(
                    requestType.InvokeMember(
                        "ResponseBody",
                        System.Reflection.BindingFlags.GetProperty,
                        Nothing,
                        comRequest,
                        Nothing))

                bodyBytes = DecompressSharedHttpBody(bodyBytes, GetSharedHttpHeaderValue(headers, "Content-Encoding"))

                Dim result As New SharedHttpResponse() With {
                    .StatusCode = statusCode,
                    .ReasonPhrase = If(reasonPhrase, ""),
                    .BodyBytes = bodyBytes,
                    .UsedStack = "WinHTTP"
                }

                ApplySharedHttpContentType(result, GetSharedHttpHeaderValue(headers, "Content-Type"))

                Return result

            Finally
                Try
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(comRequest)
                Catch
                End Try
            End Try
        End Function

        Private Shared Function CreateSharedHttpContent(request As SharedHttpRequest) As HttpContent
            If request.BodyBytes IsNot Nothing AndAlso request.BodyBytes.Length > 0 Then
                Dim content As New ByteArrayContent(request.BodyBytes)
                If Not String.IsNullOrWhiteSpace(request.ContentType) Then
                    content.Headers.ContentType = System.Net.Http.Headers.MediaTypeHeaderValue.Parse(request.ContentType)
                End If
                Return content
            End If

            If Not String.IsNullOrEmpty(request.BodyText) Then
                Dim content As New StringContent(request.BodyText, Encoding.UTF8)
                If Not String.IsNullOrWhiteSpace(request.ContentType) Then
                    content.Headers.ContentType = System.Net.Http.Headers.MediaTypeHeaderValue.Parse(request.ContentType)
                End If
                Return content
            End If

            Return Nothing
        End Function

        Private Shared Function ParseSharedHttpHeaders(rawHeaders As String) As Dictionary(Of String, String)
            Dim result As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

            If String.IsNullOrWhiteSpace(rawHeaders) Then
                Return result
            End If

            Dim lines = rawHeaders.Split(New String() {vbCrLf}, StringSplitOptions.RemoveEmptyEntries)
            For Each line As String In lines
                Dim colonIndex = line.IndexOf(":"c)
                If colonIndex <= 0 Then
                    Continue For
                End If

                Dim name = line.Substring(0, colonIndex).Trim()
                Dim value = line.Substring(colonIndex + 1).Trim()

                If result.ContainsKey(name) Then
                    result(name) = result(name) & ", " & value
                Else
                    result.Add(name, value)
                End If
            Next

            Return result
        End Function

        Private Shared Function GetSharedHttpHeaderValue(headers As Dictionary(Of String, String), headerName As String) As String
            If headers Is Nothing OrElse String.IsNullOrWhiteSpace(headerName) Then
                Return ""
            End If

            Dim value As String = Nothing
            If headers.TryGetValue(headerName, value) Then
                Return value
            End If

            Return ""
        End Function

        Private Shared Sub ApplySharedHttpContentType(response As SharedHttpResponse, contentTypeHeader As String)
            If response Is Nothing OrElse String.IsNullOrWhiteSpace(contentTypeHeader) Then
                Return
            End If

            Dim parts = contentTypeHeader.Split(";"c)
            If parts.Length > 0 Then
                response.ContentType = parts(0).Trim()
            End If

            For i As Integer = 1 To parts.Length - 1
                Dim part = parts(i).Trim()
                If part.StartsWith("charset=", StringComparison.OrdinalIgnoreCase) Then
                    response.CharSet = part.Substring(8).Trim().Trim(""""c)
                    Exit For
                End If
            Next
        End Sub

        Private Shared Function ConvertWinHttpResponseBodyToByteArray(rawBody As Object) As Byte()
            If rawBody Is Nothing Then
                Return New Byte() {}
            End If

            If TypeOf rawBody Is Byte() Then
                Return CType(rawBody, Byte())
            End If

            If TypeOf rawBody Is Object() Then
                Dim values = CType(rawBody, Object())
                Dim bytes(values.Length - 1) As Byte
                For i As Integer = 0 To values.Length - 1
                    bytes(i) = System.Convert.ToByte(values(i), System.Globalization.CultureInfo.InvariantCulture)
                Next
                Return bytes
            End If

            Dim text = System.Convert.ToString(rawBody, System.Globalization.CultureInfo.InvariantCulture)
            If String.IsNullOrEmpty(text) Then
                Return New Byte() {}
            End If

            Return Encoding.UTF8.GetBytes(text)
        End Function

        Private Shared Function DecompressSharedHttpBody(bodyBytes As Byte(), contentEncoding As String) As Byte()
            If bodyBytes Is Nothing OrElse bodyBytes.Length = 0 Then
                Return If(bodyBytes, New Byte() {})
            End If

            If String.IsNullOrWhiteSpace(contentEncoding) Then
                Return bodyBytes
            End If

            If contentEncoding.IndexOf("gzip", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Using input As New MemoryStream(bodyBytes)
                    Using zip As New GZipStream(input, CompressionMode.Decompress)
                        Using output As New MemoryStream()
                            zip.CopyTo(output)
                            Return output.ToArray()
                        End Using
                    End Using
                End Using
            End If

            If contentEncoding.IndexOf("deflate", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Using input As New MemoryStream(bodyBytes)
                    Using zip As New DeflateStream(input, CompressionMode.Decompress)
                        Using output As New MemoryStream()
                            zip.CopyTo(output)
                            Return output.ToArray()
                        End Using
                    End Using
                End Using
            End If

            Return bodyBytes
        End Function

    End Class
End Namespace
