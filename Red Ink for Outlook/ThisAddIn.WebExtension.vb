' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved.
' For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.WebExtension.vb
' Purpose:
'   Provides the local web extension surface (HTML UI + JSON API) for chat and
'   document-related LLM operations. Hosts a local `HttpListener` to serve the
'   single-file HTML UI and a JSON command API; dispatches commands to the add-in;
'   schedules/cancels LLM jobs; manages model selection (primary/secondary/alternate);
'   persists chat state; supports optional file uploads and inline extraction; and
'   converts markdown responses to HTML for browser rendering.
'
' Key Responsibilities:
'   - Serve UI (GET `InkyUiRoute`) and JSON API (POST `InkyApiRoute`) under `/inky`.
'   - Route commands (`inky_*`) and fall back to legacy dispatcher commands.
'   - Persist per-chat state (chat 1/2) via `InkyState`/`ChatTurn` in application settings.
'   - Run LLM requests as background jobs (`LlmJob`) with cancellation + TTL cleanup.
'   - Support optional file uploads (DataURL) and inline extraction for known types.
'   - Render markdown to HTML via Markdig (inline CSS/JS; no external assets).
'   - Marshal Office/UI work back onto the UI thread where required.
'   - Guard operations during suspend/resume and update watchdog progress counters.
'   - Sanitize model output (strip role markers) prior to browser display.
'
' Architecture:
'   - HTTP Layer: one `HttpListener` (prefix `/inky`) serving:
'       * UI HTML (GET `/inky`)
'       * JSON API (POST `/inky/api`)
'       * CORS preflight (OPTIONS)
'   - Routing Constants: `InkyBasePath`, `InkyUiRoute`, `InkyApiRoute`.
'   - State: `InkyState` + `ChatTurn`, stored per chat id in `My.Settings`.
'   - Jobs: `ConcurrentDictionary(Of String, LlmJob)` + TTL (`JobTtlMinutes`).
'   - Cancellation: per-job `CancellationTokenSource` + legacy `llmOperationCts`.
'   - Alternate Models: optional apply/restore guarded by `AlternateModelLock`.
'   - Tooling: optional model tooling selection + enablement persisted in chat state.
'
' Notes:
'   - This file is a partial class of `ThisAddIn`; shared flags/fields (e.g.
'     `activeRequests`, `powerChanging`, `lastListenerProgressUtc`) are defined in
'     other partials.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Markdig
Imports Microsoft.Office.Interop.Word
Imports SharedLibrary
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods
Imports System.Linq

Partial Public Class ThisAddIn

    Private Const InkyBasePath As String = "/inky"
    Private Const InkyUiRoute As String = "/inky"          ' GET → serves HTML
    Private Const InkyApiRoute As String = "/inky/api"     ' POST (JSON) → commands
    Private Const InkyName As String = "Inky"              ' Fallback; AN6 preferred

    Private Const LegacyWebExtensionRoute As String = "/redink"

    Private Function IsLocalChatAndInkyPlayEnabled() As System.Boolean
        Return INI_WebServerBlock <> 2
    End Function

    Private Function IsLegacyEdgeExtensionReceiverEnabled() As System.Boolean
        Return INI_WebServerBlock <> 3
    End Function

    Private Sub SendFeatureBlockedResponse(
        ByVal res As System.Net.HttpListenerResponse,
        ByVal requestId As System.String,
        ByVal message As System.String,
        Optional ByVal addCors As System.Boolean = False
    )
        Dim payload() As System.Byte = System.Text.Encoding.UTF8.GetBytes(message)

        SendBufferedHttpResponse(
            res,
            403,
            "text/plain; charset=utf-8",
            payload,
            requestId,
            "feature-blocked",
            addCors:=addCors)
    End Sub

    Private Const AllToolUse As String = "Advanced tools"
    Private Const AllToolUseDescription As String =
        "Turns selected advanced tools on or off. Configure which advanced tools are callable through the Agents button."

    Private activeChatId As Integer = 1   ' 1 or 2 – in‑memory only (not persisted)

    Private ReadOnly OriginalSecondaryModelName As String = INI_Model_2

    ''' <summary>Selected tools for the current chat session.</summary>
    Private _selectedToolsForChat As List(Of ModelConfig) = Nothing

    ''' <summary>Whether tooling is enabled for the chat.</summary>
    Private _chatToolingEnabled As Boolean = False

    ''' <summary>Whether agent mode (document tools) is enabled for the chat.</summary>
    Private _chatAdvancedToolsEnabled As Boolean = False

    ''' <summary>Returns the persisted effective agent-mode state and keeps the in-memory flag synchronized.</summary>
    Private Function GetEffectiveAdvancedToolsEnabled() As Boolean
        Try
            Dim st = LoadInkyState()
            Dim enabled As Boolean = (st IsNot Nothing AndAlso st.AgentModeEnabled)
            _chatAdvancedToolsEnabled = enabled
            Return enabled
        Catch
            Return _chatAdvancedToolsEnabled
        End Try
    End Function

    ''' <summary>
    ''' Represents a single LLM background job (request/response lifecycle, cancellation, file context).
    ''' </summary>
    Private Class LlmJob
        Implements IDisposable

        Public Property Id As String
        Public Property CreatedUtc As DateTime
        Public Property Tcs As TaskCompletionSource(Of String)
        Public Property Cts As CancellationTokenSource
        Public Property UseSecond As Boolean
        Public Property FileObject As String
        Public Property ChatId As Integer = 1
        Public Property ScheduledTaskId As String = ""

        Public Sub Dispose() Implements IDisposable.Dispose
            Try
                Cts?.Cancel()
                Cts?.Dispose()
                Tcs?.TrySetCanceled()
            Catch
            End Try
        End Sub
    End Class

    Private Class PendingScheduledBrowserTask
        Public Property TaskId As String = ""
        Public Property Prompt As String = ""
        Public Property ChatId As Integer = 2
        Public Property CreatedUtc As DateTime = DateTime.UtcNow
    End Class

    Private ReadOnly jobMap As New System.Collections.Concurrent.ConcurrentDictionary(Of String, LlmJob)()
    Private activeJobs As Integer = 0
    Private Const JobTtlMinutes As Integer = 45
    Private ReadOnly _pendingScheduledBrowserTaskGate As New Object()
    Private _pendingScheduledBrowserTask As PendingScheduledBrowserTask = Nothing

    Friend Sub QueuePendingScheduledBrowserTask(taskId As String, prompt As String, Optional chatId As Integer = 2)
        If String.IsNullOrWhiteSpace(taskId) OrElse String.IsNullOrWhiteSpace(prompt) Then Return

        SyncLock _pendingScheduledBrowserTaskGate
            _pendingScheduledBrowserTask = New PendingScheduledBrowserTask() With {
                .TaskId = taskId,
                .Prompt = prompt,
                .ChatId = chatId,
                .CreatedUtc = DateTime.UtcNow
            }
        End SyncLock
    End Sub

    Friend Function HasPendingScheduledBrowserTask() As Boolean
        SyncLock _pendingScheduledBrowserTaskGate
            Return _pendingScheduledBrowserTask IsNot Nothing
        End SyncLock
    End Function

    Friend Sub ClearPendingScheduledBrowserTask(Optional taskId As String = Nothing)
        SyncLock _pendingScheduledBrowserTaskGate
            If _pendingScheduledBrowserTask Is Nothing Then Return

            If String.IsNullOrWhiteSpace(taskId) OrElse
               _pendingScheduledBrowserTask.TaskId.Equals(taskId, StringComparison.OrdinalIgnoreCase) Then
                _pendingScheduledBrowserTask = Nothing
            End If
        End SyncLock
    End Sub

    Private Function ClaimPendingScheduledBrowserTask() As PendingScheduledBrowserTask
        SyncLock _pendingScheduledBrowserTaskGate
            Dim claimed = _pendingScheduledBrowserTask
            _pendingScheduledBrowserTask = Nothing
            Return claimed
        End SyncLock
    End Function

    ''' <summary>
    ''' Handles a single HTTP request (routing: favicon, UI HTML, API POST, CORS preflight).
    ''' Dispatches commands to ProcessRequestInAddIn and writes response bytes.
    ''' </summary>
    Private Async Function HandleHttpRequest(
        ByVal ctx As System.Net.HttpListenerContext,
        Optional ByVal requestId As System.String = Nothing
    ) As System.Threading.Tasks.Task

        If System.String.IsNullOrWhiteSpace(requestId) Then
            requestId = NewInkyRequestId()
        End If

        Dim req As System.Net.HttpListenerRequest = ctx.Request
        Dim res As System.Net.HttpListenerResponse = ctx.Response
        Dim requestPath As System.String = GetSafeRequestPath(req)
        Dim requestMethod As System.String = If(req.HttpMethod, System.String.Empty)
        Dim debugEnabled As System.Boolean = INI_APIDebug

        Dim isLegacyWebExtensionRequest As System.Boolean =
            System.String.Equals(requestPath, LegacyWebExtensionRoute, System.StringComparison.OrdinalIgnoreCase) OrElse
            System.String.Equals(requestPath, LegacyWebExtensionRoute & "/", System.StringComparison.OrdinalIgnoreCase)

        Dim isLocalChatRequest As System.Boolean =
            System.String.Equals(requestPath, "/", System.StringComparison.OrdinalIgnoreCase) OrElse
            System.String.Equals(requestPath, "/inky/ping", System.StringComparison.OrdinalIgnoreCase) OrElse
            System.String.Equals(requestPath, InkyUiRoute, System.StringComparison.OrdinalIgnoreCase) OrElse
            System.String.Equals(requestPath, InkyUiRoute & "/", System.StringComparison.OrdinalIgnoreCase) OrElse
            System.String.Equals(requestPath, InkyPlayRoute, System.StringComparison.OrdinalIgnoreCase) OrElse
            System.String.Equals(requestPath, InkyPlayRoute & "/", System.StringComparison.OrdinalIgnoreCase) OrElse
            System.String.Equals(requestPath, InkyApiRoute, System.StringComparison.OrdinalIgnoreCase) OrElse
            System.String.Equals(requestPath, InkyApiRoute & "/", System.StringComparison.OrdinalIgnoreCase)

        System.Threading.Interlocked.Increment(activeRequests)

        Dim hb As System.Threading.Timer = Nothing

        Try
            If isLocalChatRequest AndAlso Not IsLocalChatAndInkyPlayEnabled() Then
                SendFeatureBlockedResponse(
                    res,
                    requestId,
                    "Local chat and InkyPlay are disabled by WebServerBlock.",
                    addCors:=True)
                Return
            End If

            If isLegacyWebExtensionRequest AndAlso Not IsLegacyEdgeExtensionReceiverEnabled() Then
                SendFeatureBlockedResponse(
                    res,
                    requestId,
                    "The Edge extension receiver is disabled by WebServerBlock.",
                    addCors:=True)
                Return
            End If

            If debugEnabled Then
                Dim rawUrl As System.String = ""
                Dim hostHeader As System.String = ""
                Dim userHostName As System.String = ""
                Dim userAgent As System.String = ""
                Dim contentTypeReq As System.String = ""
                Dim contentLengthReq As System.String = ""
                Dim localEndPointText As System.String = ""
                Dim remoteEndPointText As System.String = ""

                Try
                    rawUrl = If(req.RawUrl, "")
                    hostHeader = If(req.Headers("Host"), "")
                    userHostName = If(req.UserHostName, "")
                    userAgent = If(req.UserAgent, "")
                    contentTypeReq = If(req.ContentType, "")
                    contentLengthReq = req.ContentLength64.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    localEndPointText = GetSafeEndpointText(req.LocalEndPoint)
                    remoteEndPointText = GetSafeEndpointText(req.RemoteEndPoint)
                Catch
                End Try

                AppendInkyServerLog(
                    "REQ " & requestId &
                    " BEGIN Method=" & requestMethod &
                    "; Path=" & requestPath &
                    "; RawUrl=" & rawUrl &
                    "; HostHeader=" & hostHeader &
                    "; UserHostName=" & userHostName &
                    "; UserAgent=" & ClipForInkyServerLog(userAgent, 400) &
                    "; ContentType=" & contentTypeReq &
                    "; ContentLength=" & contentLengthReq &
                    "; LocalEndPoint=" & localEndPointText &
                    "; RemoteEndPoint=" & remoteEndPointText)
            End If

            hb = New System.Threading.Timer(
                Sub(stateObj As System.Object)
                    Try
                        lastListenerProgressUtc = System.DateTime.UtcNow
                    Catch
                    End Try
                End Sub,
                state:=Nothing,
                dueTime:=System.TimeSpan.FromSeconds(5),
                period:=System.TimeSpan.FromSeconds(5))

            If System.Threading.Interlocked.CompareExchange(powerChanging, 0, 0) <> 0 Then
                Dim msgBytes() As System.Byte =
                    System.Text.Encoding.UTF8.GetBytes("Temporarily unavailable during power transition.")

                res.StatusDescription = "Service Unavailable (suspend/resume)"
                res.AddHeader("Retry-After", "2")

                SendBufferedHttpResponse(
                    res,
                    503,
                    "text/plain; charset=utf-8",
                    msgBytes,
                    requestId,
                    "powerChanging",
                    addCors:=False)

                Return
            End If

            If IsInResumeCooldown() Then
                Dim msgBytes() As System.Byte =
                    System.Text.Encoding.UTF8.GetBytes("Resuming from sleep; please retry in a few seconds.")

                res.StatusDescription = "Service Unavailable (resume cooldown)"
                res.AddHeader("Retry-After", "5")

                SendBufferedHttpResponse(
                    res,
                    503,
                    "text/plain; charset=utf-8",
                    msgBytes,
                    requestId,
                    "resumeCooldown",
                    addCors:=True)

                Return
            End If

            If requestMethod.Equals("OPTIONS", System.StringComparison.OrdinalIgnoreCase) Then
                res.AddHeader("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS")
                res.AddHeader("Access-Control-Allow-Headers", "Content-Type, Authorization")

                SendBufferedHttpResponse(
                    res,
                    204,
                    "text/plain; charset=utf-8",
                    New System.Byte() {},
                    requestId,
                    "options",
                    addCors:=True)

                Return
            End If

            If requestMethod.Equals("GET", System.StringComparison.OrdinalIgnoreCase) AndAlso
               System.String.Equals(requestPath, "/favicon.ico", System.StringComparison.OrdinalIgnoreCase) Then

                If debugEnabled Then
                    AppendInkyServerLog("REQ " & requestId & " favicon begin.")
                End If

                Dim png() As System.Byte = GetLogoPngBytes()

                SendBufferedHttpResponse(
                    res,
                    200,
                    "image/png",
                    png,
                    requestId,
                    "favicon",
                    addCors:=False)

                Return
            End If

            If requestMethod.Equals("GET", System.StringComparison.OrdinalIgnoreCase) AndAlso
               System.String.Equals(requestPath, "/inky/ping", System.StringComparison.OrdinalIgnoreCase) Then

                Dim pingBytes() As System.Byte = System.Text.Encoding.UTF8.GetBytes("ok")

                SendBufferedHttpResponse(
                    res,
                    200,
                    "text/plain; charset=utf-8",
                    pingBytes,
                    requestId,
                    "ping",
                    addCors:=True)

                Return
            End If

            If requestMethod.Equals("GET", System.StringComparison.OrdinalIgnoreCase) AndAlso
               (System.String.Equals(requestPath, "/", System.StringComparison.OrdinalIgnoreCase) OrElse
                System.String.Equals(requestPath, InkyUiRoute, System.StringComparison.OrdinalIgnoreCase) OrElse
                System.String.Equals(requestPath, InkyUiRoute & "/", System.StringComparison.OrdinalIgnoreCase)) Then

                If debugEnabled Then
                    AppendInkyServerLog("REQ " & requestId & " BuildInkyHtmlPage begin.")
                End If

                Dim html As System.String = BuildInkyHtmlPage()
                Dim bufUi() As System.Byte = System.Text.Encoding.UTF8.GetBytes(html)

                If debugEnabled Then
                    AppendInkyServerLog(
                        "REQ " & requestId &
                        " BuildInkyHtmlPage end. htmlChars=" &
                        html.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                        "; htmlBytes=" &
                        bufUi.LongLength.ToString(System.Globalization.CultureInfo.InvariantCulture))
                End If

                SendBufferedHttpResponse(
                    res,
                    200,
                    "text/html; charset=utf-8",
                    bufUi,
                    requestId,
                    "inky-ui",
                    addCors:=False)

                Return
            End If

            If requestMethod.Equals("GET", System.StringComparison.OrdinalIgnoreCase) AndAlso
               (System.String.Equals(requestPath, InkyPlayRoute, System.StringComparison.OrdinalIgnoreCase) OrElse
                System.String.Equals(requestPath, InkyPlayRoute & "/", System.StringComparison.OrdinalIgnoreCase)) Then

                If debugEnabled Then
                    AppendInkyServerLog("REQ " & requestId & " BuildInkyPlayHtmlPage begin.")
                End If

                Dim html As System.String = BuildInkyPlayHtmlPage()
                Dim bufPlay() As System.Byte = System.Text.Encoding.UTF8.GetBytes(html)

                If debugEnabled Then
                    AppendInkyServerLog(
                        "REQ " & requestId &
                        " BuildInkyPlayHtmlPage end. htmlChars=" &
                        html.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                        "; htmlBytes=" &
                        bufPlay.LongLength.ToString(System.Globalization.CultureInfo.InvariantCulture))
                End If

                SendBufferedHttpResponse(
                    res,
                    200,
                    "text/html; charset=utf-8",
                    bufPlay,
                    requestId,
                    "inky-play",
                    addCors:=False)

                Return
            End If

            If requestMethod.Equals("POST", System.StringComparison.OrdinalIgnoreCase) AndAlso
               (System.String.Equals(requestPath, LegacyWebExtensionRoute, System.StringComparison.OrdinalIgnoreCase) OrElse
                System.String.Equals(requestPath, LegacyWebExtensionRoute & "/", System.StringComparison.OrdinalIgnoreCase)) Then

                Dim body As System.String = System.String.Empty
                Dim cmd As System.String = ""

                If req.HasEntityBody Then
                    If debugEnabled Then
                        AppendInkyServerLog("REQ " & requestId & " reading legacy POST body.")
                    End If

                    Using rdr As New System.IO.StreamReader(
                        req.InputStream,
                        System.Text.Encoding.UTF8,
                        detectEncodingFromByteOrderMarks:=False,
                        bufferSize:=8192,
                        leaveOpen:=False)

                        body = Await rdr.ReadToEndAsync().ConfigureAwait(False)
                    End Using
                End If

                If debugEnabled Then
                    cmd = TryGetJsonCommandForInkyServerLog(body)

                    AppendInkyServerLog(
                        "REQ " & requestId &
                        " legacy body read complete. bodyLen=" &
                        body.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                        "; cmd=" & cmd &
                        "; bodyExcerpt=" & ClipForInkyServerLog(body, 1200))

                    AppendInkyServerLog("REQ " & requestId & " legacy ProcessRequestInAddIn begin. cmd=" & cmd)
                End If

                Dim responseText As System.String =
                    Await ProcessRequestInAddIn(body, requestPath).ConfigureAwait(False)

                If responseText Is Nothing Then responseText = System.String.Empty

                If debugEnabled Then
                    AppendInkyServerLog(
                        "REQ " & requestId &
                        " legacy ProcessRequestInAddIn end. cmd=" & cmd &
                        "; responseLen=" &
                        responseText.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                        "; responsePrefix=" &
                        ClipForInkyServerLog(responseText, 160))
                End If

                Dim contentType As System.String = "text/plain; charset=utf-8"

                If responseText.StartsWith("CT:html" & vbLf, System.StringComparison.Ordinal) Then
                    contentType = "text/html; charset=utf-8"
                    responseText = responseText.Substring(("CT:html" & vbLf).Length)
                ElseIf responseText.StartsWith("CT:json" & vbLf, System.StringComparison.Ordinal) Then
                    contentType = "application/json; charset=utf-8"
                    responseText = responseText.Substring(("CT:json" & vbLf).Length)
                End If

                Dim buf() As System.Byte = System.Text.Encoding.UTF8.GetBytes(responseText)

                SendBufferedHttpResponse(
                    res,
                    200,
                    contentType,
                    buf,
                    requestId,
                    "legacy-redink",
                    addCors:=True)

                Return
            End If

            If requestMethod.Equals("POST", System.StringComparison.OrdinalIgnoreCase) AndAlso
               (System.String.Equals(requestPath, InkyApiRoute, System.StringComparison.OrdinalIgnoreCase) OrElse
                System.String.Equals(requestPath, InkyApiRoute & "/", System.StringComparison.OrdinalIgnoreCase)) Then

                Dim body As System.String = System.String.Empty
                Dim cmd As System.String = ""

                If req.HasEntityBody Then
                    If debugEnabled Then
                        AppendInkyServerLog("REQ " & requestId & " reading POST body.")
                    End If

                    Using rdr As New System.IO.StreamReader(
                        req.InputStream,
                        System.Text.Encoding.UTF8,
                        detectEncodingFromByteOrderMarks:=False,
                        bufferSize:=8192,
                        leaveOpen:=False)

                        body = Await rdr.ReadToEndAsync().ConfigureAwait(False)
                    End Using
                End If

                If debugEnabled Then
                    cmd = TryGetJsonCommandForInkyServerLog(body)

                    AppendInkyServerLog(
                        "REQ " & requestId &
                        " body read complete. bodyLen=" &
                        body.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                        "; cmd=" & cmd &
                        "; bodyExcerpt=" & ClipForInkyServerLog(body, 1200))

                    AppendInkyServerLog("REQ " & requestId & " ProcessRequestInAddIn begin. cmd=" & cmd)
                End If

                Dim responseText As System.String =
                    Await ProcessRequestInAddIn(body, requestPath).ConfigureAwait(False)

                If responseText Is Nothing Then responseText = System.String.Empty

                If debugEnabled Then
                    AppendInkyServerLog(
                        "REQ " & requestId &
                        " ProcessRequestInAddIn end. cmd=" & cmd &
                        "; responseLen=" &
                        responseText.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                        "; responsePrefix=" &
                        ClipForInkyServerLog(responseText, 160))
                End If

                Dim contentType As System.String = "text/plain; charset=utf-8"

                If responseText.StartsWith("CT:html" & vbLf, System.StringComparison.Ordinal) Then
                    contentType = "text/html; charset=utf-8"
                    responseText = responseText.Substring(("CT:html" & vbLf).Length)
                ElseIf responseText.StartsWith("CT:json" & vbLf, System.StringComparison.Ordinal) Then
                    contentType = "application/json; charset=utf-8"
                    responseText = responseText.Substring(("CT:json" & vbLf).Length)
                End If

                Dim buf() As System.Byte = System.Text.Encoding.UTF8.GetBytes(responseText)

                SendBufferedHttpResponse(
                    res,
                    200,
                    contentType,
                    buf,
                    requestId,
                    "inky-api",
                    addCors:=True)

                Return
            End If

            If System.String.Equals(requestPath, InkyApiRoute, System.StringComparison.OrdinalIgnoreCase) OrElse
               System.String.Equals(requestPath, InkyApiRoute & "/", System.StringComparison.OrdinalIgnoreCase) Then

                Dim msg405 As System.String = "Method not allowed."
                Dim buf405() As System.Byte = System.Text.Encoding.UTF8.GetBytes(msg405)

                SendBufferedHttpResponse(
                    res,
                    405,
                    "text/plain; charset=utf-8",
                    buf405,
                    requestId,
                    "405",
                    addCors:=True)

                Return
            End If

            Dim msg404 As System.String = "Not found."
            Dim buf404() As System.Byte = System.Text.Encoding.UTF8.GetBytes(msg404)

            SendBufferedHttpResponse(
                res,
                404,
                "text/plain; charset=utf-8",
                buf404,
                requestId,
                "404",
                addCors:=True)

            Return

        Catch ex As System.Exception
            If debugEnabled Then
                AppendInkyServerLog(
                    "REQ " & requestId &
                    " HandleHttpRequest failed for " &
                    requestMethod & " " & requestPath & ".",
                    ex)
            End If

            If IsInkyTransportWriteFailure(ex) Then
                If debugEnabled Then
                    AppendInkyServerLog(
                        "REQ " & requestId &
                        " transport-level response failure detected. Skipping fallback 500.")
                End If
                Return
            End If

            Try
                Dim json As System.String = "{""ok"":false,""error"":""Internal server error.""}"
                Dim bufErr() As System.Byte = System.Text.Encoding.UTF8.GetBytes(json)

                SendBufferedHttpResponse(
                    res,
                    500,
                    "application/json; charset=utf-8",
                    bufErr,
                    requestId,
                    "fallback-500",
                    addCors:=True)

                If debugEnabled Then
                    AppendInkyServerLog("REQ " & requestId & " fallback 500 written from HandleHttpRequest catch.")
                End If
            Catch exWrite As System.Exception
                If debugEnabled Then
                    AppendInkyServerLog("REQ " & requestId & " fallback 500 write failed in HandleHttpRequest catch.", exWrite)
                End If
            End Try
        Finally
            Try
                If hb IsNot Nothing Then hb.Dispose()
            Catch exDispose As System.Exception
                If debugEnabled Then
                    AppendInkyServerLog("REQ " & requestId & " heartbeat timer dispose failed.", exDispose)
                End If
            End Try

            System.Threading.Interlocked.Decrement(activeRequests)
            lastListenerProgressUtc = System.DateTime.UtcNow

            If debugEnabled Then
                AppendInkyServerLog(
                    "REQ " & requestId &
                    " END. activeRequests=" &
                    System.Threading.Interlocked.CompareExchange(activeRequests, 0, 0).ToString(System.Globalization.CultureInfo.InvariantCulture))
            End If
        End Try
    End Function


    Private Function IsInkyTransportWriteFailure(ByVal ex As System.Exception) As System.Boolean
        Dim current As System.Exception = ex

        While current IsNot Nothing
            If TypeOf current Is System.Net.HttpListenerException Then Return True
            If TypeOf current Is System.IO.IOException Then Return True
            If TypeOf current Is System.ObjectDisposedException Then Return True
            current = current.InnerException
        End While

        Return False
    End Function

    Private Sub SendBufferedHttpResponse(
        ByVal res As System.Net.HttpListenerResponse,
        ByVal statusCode As System.Int32,
        ByVal contentType As System.String,
        ByVal payload As System.Byte(),
        ByVal requestId As System.String,
        ByVal label As System.String,
        Optional ByVal addCors As System.Boolean = False
    )
        If payload Is Nothing Then
            payload = New System.Byte() {}
        End If

        Dim closeAttempted As System.Boolean = False
        Dim debugEnabled As System.Boolean = INI_APIDebug

        Try
            res.StatusCode = statusCode

            If addCors Then
                res.AddHeader("Access-Control-Allow-Origin", "*")
            End If

            If Not System.String.IsNullOrWhiteSpace(contentType) Then
                res.ContentType = contentType
            End If

            res.SendChunked = False
            res.ContentLength64 = payload.LongLength

            If debugEnabled Then
                AppendInkyServerLog(
                    "REQ " & requestId &
                    " response write begin. label=" & label &
                    "; status=" & statusCode.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                    "; contentType=" & If(contentType, "") &
                    "; bytes=" & payload.LongLength.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                    "; writer=OutputStreamWrite-v2")
            End If

            If payload.Length > 0 Then
                res.OutputStream.Write(payload, 0, payload.Length)
            End If

            res.OutputStream.Flush()

            If debugEnabled Then
                AppendInkyServerLog(
                    "REQ " & requestId &
                    " response write flushed. label=" & label &
                    "; writer=OutputStreamWrite-v2")
            End If

        Catch ex As System.Exception
            If debugEnabled Then
                AppendInkyServerLog(
                    "REQ " & requestId &
                    " response write failed. label=" & label &
                    "; writer=OutputStreamWrite-v2",
                    ex)
            End If

            Throw

        Finally
            Try
                closeAttempted = True
                res.Close()

                If debugEnabled Then
                    AppendInkyServerLog(
                        "REQ " & requestId &
                        " response close complete. label=" & label &
                        "; writer=OutputStreamWrite-v2")
                End If

            Catch exClose As System.Exception
                If debugEnabled Then
                    AppendInkyServerLog(
                        "REQ " & requestId &
                        " response close failed/ignored. label=" & label &
                        "; closeAttempted=" & closeAttempted.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                        "; writer=OutputStreamWrite-v2",
                        exClose)
                End If
            End Try
        End Try

        If debugEnabled Then
            AppendInkyServerLog(
                "REQ " & requestId &
                " response write complete. label=" & label &
                "; writer=OutputStreamWrite-v2")
        End If
    End Sub


    ''' <summary>
    ''' Returns logo PNG bytes; falls back to 1x1 transparent if resource load fails.
    ''' </summary>
    Private Function GetLogoPngBytes() As System.Byte()
        Try
            Using src As System.Drawing.Bitmap = CType(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard).Clone(), System.Drawing.Bitmap)
                Using ms As New System.IO.MemoryStream()
                    src.Save(ms, System.Drawing.Imaging.ImageFormat.Png)
                    Return ms.ToArray()
                End Using
            End Using
        Catch
            ' 1x1 transparent PNG fallback
            Return System.Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQImWNgYGD4DwABdQF+8m3rXQAAAABJRU5ErkJggg==")
        End Try
    End Function


    ' LLM helper (runs off the UI thread) 


    ' 1) Field for scheduler (class/module level)
    Private Shared llmScheduler As System.Threading.Tasks.TaskScheduler


    ' 2) Initialize STA thread with dedicated WinForms message loop


    ''' <summary>
    ''' Ensures a dedicated STA thread with WindowsFormsSynchronizationContext is running for UI-bound operations if required.
    ''' </summary>
    Private Sub EnsureLlmUiThread()
        If llmThread IsNot Nothing AndAlso llmThread.IsAlive Then Return
        If llmScheduler IsNot Nothing Then Return

        Dim tcs As New System.Threading.Tasks.TaskCompletionSource(Of System.Threading.Tasks.TaskScheduler)()

        llmThread = New System.Threading.Thread(
            Sub()
                System.Threading.SynchronizationContext.SetSynchronizationContext(
                    New System.Windows.Forms.WindowsFormsSynchronizationContext())

                llmSyncContext = System.Threading.SynchronizationContext.Current
                tcs.SetResult(System.Threading.Tasks.TaskScheduler.FromCurrentSynchronizationContext())

                System.Windows.Forms.Application.Run()
            End Sub)

        llmThread.SetApartmentState(System.Threading.ApartmentState.STA)
        llmThread.IsBackground = True
        llmThread.Start()

        llmScheduler = tcs.Task.Result
    End Sub

    ''' <summary>
    ''' Signals STA thread to exit and joins with timeout; clears scheduler and context references.
    ''' </summary>
    Private Sub StopLlmUiThread()
        Try
            If llmSyncContext IsNot Nothing Then
                llmSyncContext.Post(
                    Sub() System.Windows.Forms.Application.ExitThread(),
                    Nothing)
            End If
        Catch
        End Try
        Try
            If llmThread IsNot Nothing AndAlso llmThread.IsAlive Then
                If Not llmThread.Join(2000) Then
                    ' Optional: Log that the thread did not terminate in time.
                End If
            End If
        Catch
        End Try

        llmScheduler = Nothing
        llmSyncContext = Nothing
        llmThread = Nothing
    End Sub

    ''' <summary>
    ''' Runs an LLM operation asynchronously on thread pool with timeout, cancellation propagation, and optional secondary model use.
    ''' Returns model output or status text.
    ''' </summary>
    ''' <summary>
    ''' Runs an LLM operation asynchronously on thread pool with timeout, cancellation propagation, and optional secondary model use.
    ''' Returns model output or status text.
    ''' </summary>
    Public Function RunLlmAsync(
        ByVal sysPrompt As System.String,
        ByVal userPrompt As System.String,
        Optional ByVal UseSecondAPI As System.Boolean = False,
        Optional ByVal ShowTimer As System.Boolean = True,
        Optional ByVal FileObject As System.String = "",
        Optional ByVal cancellationToken As System.Threading.CancellationToken = Nothing
    ) As System.Threading.Tasks.Task(Of System.String)

        ' Use the thread pool – no STA message loop
        Dim effectiveTimeout As Integer = CInt(If(UseSecondAPI, INI_Timeout_2, INI_Timeout))
        ModelTimeout = effectiveTimeout

        Return System.Threading.Tasks.Task.Run(
            Async Function() As System.Threading.Tasks.Task(Of System.String)
                ' Check if we're in a power transition BEFORE starting
                If System.Threading.Interlocked.CompareExchange(powerChanging, 0, 0) <> 0 Then
                    Return "Operation cancelled due to power transition."
                End If

                Using linkedCts As System.Threading.CancellationTokenSource =
                    System.Threading.CancellationTokenSource.CreateLinkedTokenSource(cancellationToken)

                    ' Create a separate timeout CTS that we can control
                    Using timeoutCts As New System.Threading.CancellationTokenSource()
                        ' Add a generous buffer to the model timeout
                        Dim totalTimeout = effectiveTimeout + 60 ' 60 second buffer
                        timeoutCts.CancelAfter(TimeSpan.FromSeconds(totalTimeout))

                        Using combinedCts As System.Threading.CancellationTokenSource =
                            System.Threading.CancellationTokenSource.CreateLinkedTokenSource(linkedCts.Token, timeoutCts.Token)

                            ' Publish this operation's CTS so HandlePowerSuspendAsync can cancel
                            ' the in-flight LLM call before the system suspends. Capture any
                            ' previously published CTS so we can restore it on completion.
                            Dim previousLlmCts As System.Threading.CancellationTokenSource =
                                System.Threading.Interlocked.Exchange(llmOperationCts, combinedCts)

                            ' If a suspend already started while we were setting up, honor it now.
                            If System.Threading.Interlocked.CompareExchange(powerChanging, 0, 0) <> 0 Then
                                Try : combinedCts.Cancel() : Catch : End Try
                            End If

                            Try
                                Dim llmTask = LLM(sysPrompt, userPrompt, "", "", 0, UseSecondAPI,
                                                  Not ShowTimer, "", FileObject, combinedCts.Token, EnsureUI:=False)

                                ' Wait for either completion or cancellation
                                Dim llmOut As String = Await llmTask.ConfigureAwait(False)

                                ' Mark successful completion
                                lastSuccessfulOperationUtc = System.DateTime.UtcNow

                                If UseSecondAPI AndAlso originalConfigLoaded Then
                                    RestoreDefaults(_context, originalConfig)
                                    originalConfigLoaded = False
                                End If

                                Return If(llmOut, String.Empty)

                            Catch ex As OperationCanceledException When timeoutCts.IsCancellationRequested
                                ' This was a timeout, not user cancellation
                                System.Diagnostics.Debug.WriteLine($"LLM operation timed out after {totalTimeout}s")
                                Return "Operation timed out. Please try again with a shorter prompt or different model."

                            Catch ex As OperationCanceledException
                                ' Check if this is due to power transition
                                If System.Threading.Interlocked.CompareExchange(powerChanging, 0, 0) <> 0 Then
                                    Return "Operation cancelled due to power transition."
                                End If
                                Return "Operation was canceled by the user."

                            Catch ex As System.Exception
                                ' Log the exception for debugging
                                System.Diagnostics.Debug.WriteLine($"LLM Error: {ex.Message}")
                                Return "An error occurred during processing."

                            Finally
                                ' Stop publishing this operation's CTS, but only if it is still
                                ' the current one (a concurrent op or a suspend may have replaced
                                ' or cleared it already). Restore the previously published CTS.
                                System.Threading.Interlocked.CompareExchange(llmOperationCts, previousLlmCts, combinedCts)

                                ' Ensure cleanup happens even during power transitions
                                Try
                                    If UseSecondAPI AndAlso originalConfigLoaded Then
                                        RestoreDefaults(_context, originalConfig)
                                        originalConfigLoaded = False
                                    End If
                                Catch
                                End Try
                            End Try
                        End Using
                    End Using
                End Using
            End Function,
            cancellationToken)
    End Function

    ''' <summary>
    ''' Presents comparison UI and returns user acceptance Boolean (True accepted, False canceled).
    ''' </summary>
    Private Async Function CompareAndInsertSyncConfirm(
        originalText As String,
        llmResult As String
    ) As System.Threading.Tasks.Task(Of Boolean)

        ' Show compare on UI thread
        Await SwitchToUi(Sub() CompareAndInsertText(originalText, llmResult, True)).ConfigureAwait(False)

        ' Await user decision (Esc = False, OK/close = True)
        Dim accepted As Boolean = Await WaitForPreviewDecisionAsync().ConfigureAwait(False)
        Return accepted
    End Function


    ''' <summary>
    ''' Waits for preview dialog dismissal (OK = True, Escape = False). Attaches handlers via polling timer.
    ''' </summary>
    Private Async Function WaitForPreviewDecisionAsync() As System.Threading.Tasks.Task(Of System.Boolean)
        Dim tcs As New System.Threading.Tasks.TaskCompletionSource(Of System.Boolean)()

        ' Handler attachment once on UI thread
        Await SwitchToUi(Sub()

                             Dim previewForm As System.Windows.Forms.Form = Nothing
                             Dim searchTimer As New System.Windows.Forms.Timer() With {.Interval = 100}

                             AddHandler searchTimer.Tick,
                                 Sub()
                                     If previewForm Is Nothing OrElse previewForm.IsDisposed Then
                                         previewForm = System.Windows.Forms.Application.OpenForms _
                                             .Cast(Of System.Windows.Forms.Form)() _
                                             .FirstOrDefault(Function(f As System.Windows.Forms.Form) f.Name = "ShowRTFCustomMessageBox" _
                                                                 OrElse f.Text.StartsWith(AN))

                                         If previewForm Is Nothing Then Return

                                         previewForm.KeyPreview = True

                                         AddHandler previewForm.KeyDown,
                                             Sub(_s As System.Object, e As System.Windows.Forms.KeyEventArgs)
                                                 If e.KeyCode = System.Windows.Forms.Keys.Escape Then
                                                     tcs.TrySetResult(False)
                                                 End If
                                             End Sub

                                         AddHandler previewForm.FormClosed,
                                             Sub(_s As System.Object, _e As System.Windows.Forms.FormClosedEventArgs)
                                                 tcs.TrySetResult(True)
                                                 ' Failsafe: dispose timer
                                                 If searchTimer.Enabled Then
                                                     searchTimer.Stop()
                                                     searchTimer.Dispose()
                                                 End If
                                             End Sub
                                     End If

                                     If tcs.Task.IsCompleted Then
                                         searchTimer.Stop()
                                         searchTimer.Dispose() ' Patch C
                                     End If
                                 End Sub

                             searchTimer.Start()
                         End Sub).ConfigureAwait(False)

        ' IMPORTANT: In async function Task(Of Boolean) → Await tcs.Task
        Return Await tcs.Task.ConfigureAwait(False)
    End Function


    ' Chatbot 


    ''' <summary>
    ''' Attempts to read a setting value by reflection; returns fallback if missing or invalid.
    ''' </summary>
    Private Function TryGetAppSetting(Of T)(ByVal key As System.String, ByVal fallback As T) As T
        Try
            Dim p = GetType(My.MySettings).GetProperty(key, System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.Instance)
            If p IsNot Nothing Then
                Dim v = DirectCast(p.GetValue(My.Settings, Nothing), Object)
                If v IsNot Nothing Then Return DirectCast(v, T)
            End If
        Catch
        End Try
        Return fallback
    End Function

    ''' <summary>
    ''' Returns bot name based on AN6 setting or default fallback.
    ''' </summary>
    Private Function GetBotName() As System.String
        ' Try: My.Settings("AN6") → else "Inky"
        Dim v As System.String = TryGetAppSetting(Of System.String)("AN6", Nothing)
        If Not System.String.IsNullOrWhiteSpace(v) Then Return v
        Return "Inky"
    End Function

    ''' <summary>
    ''' Returns a data URL (PNG base64) for the logo or empty string on failure.
    ''' </summary>
    Private Function GetLogoDataUrl() As System.String
        Try
            Using src As System.Drawing.Bitmap = CType(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard).Clone(), System.Drawing.Bitmap)
                Using ms As New System.IO.MemoryStream()
                    src.Save(ms, System.Drawing.Imaging.ImageFormat.Png)
                    Dim b64 As System.String = System.Convert.ToBase64String(ms.ToArray())
                    Return "data:image/png;base64," & b64
                End Using
            End Using
        Catch
            Return ""
        End Try
    End Function


    ''' <summary>
    ''' Returns display key for model (prefers ModelDescription, falls back to Model).
    ''' </summary>
    Private Function GetModelDisplayKey(ByVal model As ModelConfig) As System.String
        If model Is Nothing Then Return ""
        ' Prefer descriptive label:
        If Not System.String.IsNullOrWhiteSpace(model.ModelDescription) Then
            Return model.ModelDescription
        End If
        ' Fallback: internal model name
        If Not System.String.IsNullOrWhiteSpace(model.Model) Then
            Return model.Model
        End If
        Return "Model"
    End Function

    ''' <summary>
    ''' Returns the host application's UI culture two-letter ISO code or "en" fallback.
    ''' </summary>
    Private Function GetHostUiLanguageTwoLetter() As System.String
        Try
            Dim app As Microsoft.Office.Interop.Outlook.Application = Nothing
            Dim languageSettings As Microsoft.Office.Core.LanguageSettings = Nothing

            Try
                app = Me.Application
            Catch
                app = Nothing
            End Try

            If app IsNot Nothing Then
                Try
                    languageSettings = app.LanguageSettings
                Catch
                    languageSettings = Nothing
                End Try
            End If

            If languageSettings IsNot Nothing Then
                Dim lcid As System.Int32 =
                    languageSettings.LanguageID(Microsoft.Office.Core.MsoAppLanguageID.msoLanguageIDUI)

                Return New System.Globalization.CultureInfo(lcid).TwoLetterISOLanguageName
            End If
        Catch
        End Try

        Try
            Return System.Globalization.CultureInfo.CurrentUICulture.TwoLetterISOLanguageName
        Catch
            Return "en"
        End Try
    End Function

    ''' <summary>
    ''' Returns localized greeting based on the Outlook UI language.
    ''' </summary>
    Private Function GetFriendlyGreeting() As System.String
        Dim name As System.String = GetBotName()
        Dim tl As System.String = GetHostUiLanguageTwoLetter()

        Select Case tl
            Case "de" : Return $"Hallo! Ich bin {name}. Wie kann ich helfen?"
            Case "fr" : Return $"Salut ! Je suis {name}. Comment puis-je aider ?"
            Case "it" : Return $"Ciao! Sono {name}. In cosa posso aiutarti?"
            Case "es" : Return $"¡Hola! Soy {name}. ¿En qué puedo ayudarte?"
            Case Else : Return $"Hi! I’m {name}. How can I help?"
        End Select
    End Function

    Dim botName As System.String = GetBotName()
    Dim brandName As System.String = AN
    Dim logoUrl As System.String = GetLogoDataUrl()
    Dim greet As System.String = GetFriendlyGreeting()

    ''' <summary>
    ''' Simple persisted state container for chat.
    ''' </summary>
    <Serializable>
    Private Class InkyState
        Public History As System.Collections.Generic.List(Of ChatTurn) = New System.Collections.Generic.List(Of ChatTurn)()
        Public SelectedModelKey As System.String = ""
        Public UseSecondApi As System.Boolean = False
        Public LastAssistantText As System.String = ""
        Public DarkMode As System.Boolean = False
        Public SupportsFileUploads As System.Boolean = False
        Public ToolingEnabled As System.Boolean = False
        Public SelectedToolNames As System.Collections.Generic.List(Of String) = New System.Collections.Generic.List(Of String)()
        ' Opt-out persistence: store only the tools the user explicitly turned OFF.
        ' Anything not listed is on by default (incl. newly added tools). Renamed from the
        ' legacy "Selected..." fields so old opt-in values are dropped on first load.
        Public DeselectedMainToolNames As System.Collections.Generic.List(Of String) = New System.Collections.Generic.List(Of String)()
        Public DeselectedAdvancedToolNames As System.Collections.Generic.List(Of String) = New System.Collections.Generic.List(Of String)()
        Public AgentModeEnabled As System.Boolean = False
        Public PreAgentModelKey As System.String = ""
        Public PreAgentUseSecondApi As System.Boolean = False
        Public AgentModelActive As System.Boolean = False
    End Class

    ''' <summary>
    ''' Represents a single conversational turn (user or assistant) with markdown/HTML payload and timestamp.
    ''' </summary>
    <Serializable>
    Private Class ChatTurn
        Public Role As System.String   ' "user" or "assistant"
        Public Markdown As System.String
        Public Html As System.String
        Public Utc As System.DateTime
    End Class

    ''' <summary>
    ''' Returns the Outlook UI language two-letter ISO code or "en" fallback.
    ''' </summary>
    Private Function GetUserLanguageTwoLetter() As System.String
        Return GetHostUiLanguageTwoLetter()
    End Function

    ''' <summary>
    ''' Determines whether file uploads are supported for current model selection.
    ''' </summary>
    Private Function ComputeSupportsFiles(ByVal useSecond As System.Boolean,
                                          ByVal selectedKey As System.String) As System.Boolean
        Try
            ' Primary API (no alternate selected)
            If Not useSecond Then
                Return Not System.String.IsNullOrWhiteSpace(INI_APICall_Object)
            End If

            ' Second API, default model
            If System.String.IsNullOrWhiteSpace(selectedKey) Then
                Return Not System.String.IsNullOrWhiteSpace(INI_APICall_Object_2)
            End If

            ' Second API, alternate model -> read APICall_Object
            Dim alts As System.Collections.Generic.List(Of ModelConfig) = Nothing
            Try
                alts = LoadAlternativeModels(INI_AlternateModelPath, _context)
            Catch
                alts = Nothing
            End Try
            If alts Is Nothing Then Return False

            Dim sel As ModelConfig =
                alts.FirstOrDefault(Function(m As ModelConfig)
                                        If m Is Nothing Then Return False
                                        If Not System.String.IsNullOrWhiteSpace(m.ModelDescription) AndAlso
                                           System.String.Equals(m.ModelDescription, selectedKey, System.StringComparison.OrdinalIgnoreCase) Then
                                            Return True
                                        End If
                                        If Not System.String.IsNullOrWhiteSpace(m.Model) AndAlso
                                           System.String.Equals(m.Model, selectedKey, System.StringComparison.OrdinalIgnoreCase) Then
                                            Return True
                                        End If
                                        Return False
                                    End Function)

            If sel Is Nothing Then Return False

            ' Direct access – fallback via reflection if property missing
            Dim v As System.String = Nothing
            Try
                v = sel.APICall_Object
            Catch
                Try
                    Dim p As System.Reflection.PropertyInfo =
                        GetType(ModelConfig).GetProperty("APICall_Object",
                            System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.Instance)
                    If p IsNot Nothing Then
                        Dim o As System.Object = p.GetValue(sel, Nothing)
                        If o IsNot Nothing Then v = System.Convert.ToString(o, System.Globalization.CultureInfo.InvariantCulture)
                    End If
                Catch
                End Try
            End Try

            Return Not System.String.IsNullOrWhiteSpace(v)
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Loads chat state for given (or active) chat id; initializes defaults if empty.
    ''' </summary>
    Private Function LoadInkyState(Optional chatId As Integer = -1) As InkyState
        If chatId = -1 Then chatId = activeChatId
        Dim settingKey As String = If(chatId = 2, "ChatHistory_Inky2", "ChatHistory_Inky")
        Try
            Dim raw As String = ""
            Try : raw = DirectCast(My.Settings.[GetType]().GetProperty(settingKey).GetValue(My.Settings, Nothing), String) : Catch : raw = "" : End Try
            If String.IsNullOrWhiteSpace(raw) Then
                Dim st As New InkyState()
                ' Default dark mode on first empty chat
                st.DarkMode = True
                Return st
            End If
            Dim stLoaded = Newtonsoft.Json.JsonConvert.DeserializeObject(Of InkyState)(raw)
            If stLoaded Is Nothing Then stLoaded = New InkyState()
            Return stLoaded
        Catch
            Dim st As New InkyState() : st.DarkMode = True
            Return st
        End Try
    End Function

    ''' <summary>
    ''' Persists chat state for given (or active) chat id into application settings.
    ''' </summary>
    Private Sub SaveInkyState(st As InkyState, Optional chatId As Integer = -1)
        TrySaveInkyState(st, chatId, "SaveInkyState")
    End Sub

    Private Function TrySaveInkyState(st As InkyState,
                                  Optional chatId As Integer = -1,
                                  Optional operationName As String = "") As Boolean
        If chatId = -1 Then chatId = activeChatId

        Dim settingKey As String = If(chatId = 2, "ChatHistory_Inky2", "ChatHistory_Inky")
        Dim opName As String = If(String.IsNullOrWhiteSpace(operationName), "SaveInkyState", operationName)

        Try
            Dim json = Newtonsoft.Json.JsonConvert.SerializeObject(st)

            My.Settings.[GetType]().
            GetProperty(settingKey).
            SetValue(My.Settings, json, Nothing)

            My.Settings.Save()

            Debug.WriteLine(
            $"[Inky] {opName}: saved chat={chatId}; turns={If(st?.History?.Count, 0)}; lastAssistantLen={If(If(st?.LastAssistantText, ""), "").Length}")

            Return True
        Catch ex As Exception
            Debug.WriteLine($"[Inky] {opName} failed: chat={chatId}; {ex.Message}")
            Return False
        End Try
    End Function

    Private Function PersistAssistantTurnForJob(job As LlmJob,
                                            assistantText As String,
                                            assistantHtml As String) As Boolean
        If job Is Nothing Then Return False

        Dim targetChatId As Integer = If(job.ChatId = 2, 2, 1)

        Try
            Dim stJob As InkyState = LoadInkyState(targetChatId)

            stJob.History.Add(New ChatTurn With {
            .Role = "assistant",
            .Markdown = assistantText,
            .Html = assistantHtml,
            .Utc = Date.UtcNow
        })

            stJob.LastAssistantText = assistantText

            Dim saved As Boolean = TrySaveInkyState(
            stJob,
            targetChatId,
            "PersistAssistantTurnForJob")

            Debug.WriteLine(
            $"[Inky] PersistAssistantTurnForJob: job={If(job.Id, "")}; targetChat={targetChatId}; activeChat={activeChatId}; saved={saved}; assistantLen={If(assistantText, "").Length}")

            Return saved
        Catch ex As Exception
            Debug.WriteLine(
            $"[Inky] PersistAssistantTurnForJob failed: job={If(job.Id, "")}; targetChat={targetChatId}; activeChat={activeChatId}; {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Converts markdown to HTML using Markdig advanced pipeline; falls back to HTML-encoded text on error.
    ''' </summary>
    Private Function MarkdownToHtml(ByVal md As System.String) As System.String
        Try
            ' Maximum markdown functionality + soft line breaks as <br/>
            Dim pipeline As Markdig.MarkdownPipeline =
                New Markdig.MarkdownPipelineBuilder().
                    UseAdvancedExtensions().
                    UseSoftlineBreakAsHardlineBreak().
                    UsePipeTables().
                    UseGridTables().
                    UseListExtras().
                    UseFootnotes().
                    UseDefinitionLists().
                    UseAbbreviations().
                    UseAutoLinks().
                    UseTaskLists().
                    UseMathematics().
                    UseFigures().
                    UseGenericAttributes().
                    Build()

            Return Markdig.Markdown.ToHtml(md, pipeline)
        Catch ex As System.Exception
            ' Fallback: safely encode and preserve line breaks
            Return System.Net.WebUtility.HtmlEncode(md).Replace(vbLf, "<br>")
        End Try
    End Function

    ''' <summary>
    ''' Returns a clipped copy of history limited to the specified character cap (most recent retained).
    ''' </summary>
    Private Function CapHistoryToChars(ByVal st As InkyState, ByVal maxChars As System.Int32) As System.Collections.Generic.List(Of ChatTurn)
        If maxChars <= 0 Then Return New System.Collections.Generic.List(Of ChatTurn)(st.History)
        Dim acc As New System.Text.StringBuilder()
        Dim clipped As New System.Collections.Generic.List(Of ChatTurn)()
        ' iterate from the end (most recent) backwards until cap reached
        For i As System.Int32 = st.History.Count - 1 To 0 Step -1
            Dim turn As ChatTurn = st.History(i)
            Dim piece As System.String = $"[{turn.Role}]{turn.Markdown}" & vbLf
            If acc.Length + piece.Length > maxChars Then Exit For
            acc.Insert(0, piece)
            clipped.Insert(0, turn)
        Next
        Return clipped
    End Function

    ''' <summary>
    ''' Returns user-friendly label for currently selected model (primary or secondary).
    ''' </summary>
    Private Function GetSelectedModelLabel(ByVal useSecond As System.Boolean, ByVal selectedKey As System.String) As System.String
        If Not useSecond Then
            Return If(System.String.IsNullOrWhiteSpace(INI_Model), "Default model", INI_Model)
        End If
        If Not System.String.IsNullOrWhiteSpace(selectedKey) Then
            Return selectedKey
        End If
        Return If(System.String.IsNullOrWhiteSpace(INI_Model_2), "Second API model", INI_Model_2)
    End Function

    ''' <summary>
    ''' Ensures tooling selections are loaded and the tooling-enabled flag is consistent with model support.
    ''' </summary>
    ''' <summary>
    ''' Ensures tooling selections are loaded and the tooling-enabled flag is consistent with model support.
    ''' </summary>
    Private Function SyncToolingState(ByVal st As InkyState, ByRef supportsTooling As Boolean) As Boolean
        _selectedToolsForChat = GetLocalChatEffectiveSelection(st, includeInteractiveM365Tools:=True)

        supportsTooling = CurrentModelSupportsTooling(st)

        If _apActive Then
            supportsTooling = False
        End If

        Dim hasTools As Boolean =
            (_selectedToolsForChat IsNot Nothing AndAlso _selectedToolsForChat.Count > 0) OrElse
            IsChatAgentWorkspaceConnected()

        Dim enabled As Boolean = st.ToolingEnabled AndAlso hasTools AndAlso supportsTooling

        If st.ToolingEnabled <> enabled Then
            st.ToolingEnabled = enabled
            SaveInkyState(st)
        End If

        _chatToolingEnabled = enabled
        _chatAdvancedToolsEnabled = st.AgentModeEnabled AndAlso enabled AndAlso supportsTooling AndAlso Not _apActive

        Return enabled
    End Function

    ''' <summary>
    ''' Determines whether the currently selected model supports tooling.
    ''' </summary>
    Private Function CurrentModelSupportsTooling(st As InkyState) As Boolean
        Try
            If Not st.UseSecondApi Then
                ' Check primary model - need to get its config and check APICall_ToolInstructions
                Dim primaryConfig = GetCurrentConfig(_context)
                Return primaryConfig IsNot Nothing AndAlso ModelSupportsTooling(primaryConfig)
            End If

            If String.IsNullOrWhiteSpace(st.SelectedModelKey) Then
                ' Default secondary model - check INI_APICall_ToolInstructions_2
                Return Not String.IsNullOrWhiteSpace(INI_APICall_ToolInstructions_2)
            End If

            ' Check selected alternate model
            Dim alts As List(Of ModelConfig) = Nothing
            Try
                alts = LoadAlternativeModels(INI_AlternateModelPath, _context)
            Catch
                Return False
            End Try

            If alts Is Nothing Then Return False

            Dim sel = alts.FirstOrDefault(Function(m)
                                              If m Is Nothing Then Return False
                                              If Not String.IsNullOrWhiteSpace(m.ModelDescription) AndAlso
                                                 String.Equals(m.ModelDescription, st.SelectedModelKey, StringComparison.OrdinalIgnoreCase) Then Return True
                                              If Not String.IsNullOrWhiteSpace(m.Model) AndAlso
                                                 String.Equals(m.Model, st.SelectedModelKey, StringComparison.OrdinalIgnoreCase) Then Return True
                                              Return False
                                          End Function)

            Return sel IsNot Nothing AndAlso ModelSupportsTooling(sel)
        Catch
            Return False
        End Try
    End Function


    ''' <summary>
    ''' Scans the alternate models INI for an entry with AgentDefaultModel=True.
    ''' Returns the ModelConfig and its display key without applying it to context.
    ''' </summary>
    ''' <param name="displayKey">Output: the display label for the found model.</param>
    ''' <returns>The ModelConfig if found; Nothing otherwise.</returns>
    Private Function FindAgentDefaultModel(ByRef displayKey As String) As ModelConfig
        displayKey = Nothing
        Try
            If String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then Return Nothing

            Dim alts As List(Of ModelConfig) = Nothing
            Try
                alts = LoadAlternativeModels(INI_AlternateModelPath, _context, includeToolOnly:=False, toolsOnly:=False)
            Catch
                Return Nothing
            End Try
            If alts Is Nothing OrElse alts.Count = 0 Then Return Nothing

            ' GetSpecialTaskModel searches for a key with a truthy value.
            ' We replicate the search here WITHOUT applying to context.
            Dim iniPath As String = ExpandEnvironmentVariables(INI_AlternateModelPath)
            If Not IO.File.Exists(iniPath) Then Return Nothing

            Dim truthy As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
                "true", "yes", "wahr", "ja", "on", "1"
            }

            Dim currentDict As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            Dim sectionName As String = ""

            Dim checkSection As Func(Of String) =
                Function()
                    If currentDict.Count = 0 Then Return Nothing
                    If Not currentDict.ContainsKey("AgentDefaultModel") Then Return Nothing
                    Dim raw As String = currentDict("AgentDefaultModel")
                    If raw Is Nothing Then Return Nothing

                    ' Strip inline comments, quotes
                    Dim scIdx = raw.IndexOf(";"c) : If scIdx >= 0 Then raw = raw.Substring(0, scIdx)
                    Dim hashIdx = raw.IndexOf("#"c) : If hashIdx >= 0 Then raw = raw.Substring(0, hashIdx)
                    raw = raw.Trim()
                    If raw.Length >= 2 AndAlso ((raw.StartsWith("""") AndAlso raw.EndsWith("""")) OrElse
                                                (raw.StartsWith("'") AndAlso raw.EndsWith("'"))) Then
                        raw = raw.Substring(1, raw.Length - 2).Trim()
                    End If
                    If truthy.Contains(raw.ToLowerInvariant()) Then Return sectionName
                    Return Nothing
                End Function

            ' Helper: match a raw INI section name to a loaded ModelConfig.
            ' ModelDescription may be decorated with ModelNote and/or ToolingSuffix,
            ' so we use StartsWith for ModelDescription and Equals for Model.
            Dim matchToConfig As Func(Of String, ModelConfig) =
                Function(section As String)
                    Return alts.FirstOrDefault(Function(m)
                                                   If m Is Nothing Then Return False
                                                   If Not String.IsNullOrWhiteSpace(m.ModelDescription) AndAlso
                                                      m.ModelDescription.StartsWith(section, StringComparison.OrdinalIgnoreCase) Then Return True
                                                   If Not String.IsNullOrWhiteSpace(m.Model) AndAlso
                                                      String.Equals(m.Model, section, StringComparison.OrdinalIgnoreCase) Then Return True
                                                   Return False
                                               End Function)
                End Function

            For Each rawLine In IO.File.ReadAllLines(iniPath)
                Dim line = rawLine.Trim()
                If line.Length = 0 OrElse line.StartsWith(";") OrElse line.StartsWith("#") Then Continue For

                If line.StartsWith("[") AndAlso line.EndsWith("]") Then
                    Dim matchedSection = checkSection()
                    If matchedSection IsNot Nothing Then
                        Dim mc = matchToConfig(matchedSection)
                        If mc IsNot Nothing Then
                            displayKey = If(Not String.IsNullOrWhiteSpace(mc.ModelDescription), mc.ModelDescription, mc.Model)
                            Return mc
                        End If
                    End If
                    currentDict.Clear()
                    sectionName = line.Substring(1, line.Length - 2).Trim()
                    Continue For
                End If

                Dim tokens = line.Split(New Char() {"="c}, 2)
                If tokens.Length = 2 Then currentDict(tokens(0).Trim()) = tokens(1).Trim()
            Next

            ' Check final section
            Dim finalMatch = checkSection()
            If finalMatch IsNot Nothing Then
                Dim mc = matchToConfig(finalMatch)
                If mc IsNot Nothing Then
                    displayKey = If(Not String.IsNullOrWhiteSpace(mc.ModelDescription), mc.ModelDescription, mc.Model)
                    Return mc
                End If
            End If

            Return Nothing
        Catch
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Returns True if an AgentDefaultModel is defined in the alternate models INI.
    ''' Lightweight check — does not load model configs.
    ''' </summary>
    Private Function IsAgentDefaultModelAvailable() As Boolean
        Dim dummy As String = Nothing
        Return FindAgentDefaultModel(dummy) IsNot Nothing
    End Function


    ''' <summary>
    ''' Checks if tooling should be used based on current settings and model capability.
    ''' </summary>
    Private Function ShouldUseTooling(st As InkyState) As Boolean
        If Not st.ToolingEnabled Then Return False

        _selectedToolsForChat = GetLocalChatEffectiveSelection(st, includeInteractiveM365Tools:=True)

        If activeChatId = AP_SchedulerLocalChatId AndAlso HasActiveScheduledTaskWorkspace() Then
            _selectedToolsForChat =
                MergeScheduledTaskExecutionTools(
                    _selectedToolsForChat,
                    includeAllAdvancedTools:=True,
                    executionMode:=AP_TaskExecutionModeBrowserPrompt)
        End If

        If (_selectedToolsForChat Is Nothing OrElse _selectedToolsForChat.Count = 0) AndAlso
           Not IsChatAgentWorkspaceConnected() AndAlso
           Not HasActiveScheduledTaskWorkspace() Then
            Return False
        End If

        Return CurrentModelSupportsTooling(st)
    End Function

    ''' <summary>
    ''' Ensures tools are selected, prompting user if necessary.
    ''' </summary>
    Private Function EnsureToolsSelected(st As InkyState,
                                         Optional includeInteractiveM365Tools As Boolean = False) As Boolean
        _selectedToolsForChat = GetLocalChatEffectiveSelection(st, includeInteractiveM365Tools)

        If activeChatId = AP_SchedulerLocalChatId AndAlso HasActiveScheduledTaskWorkspace() Then
            _selectedToolsForChat =
                MergeScheduledTaskExecutionTools(
                    _selectedToolsForChat,
                    includeAllAdvancedTools:=True,
                    executionMode:=AP_TaskExecutionModeBrowserPrompt)
        End If

        Return (_selectedToolsForChat IsNot Nothing AndAlso _selectedToolsForChat.Count > 0) OrElse
               IsChatAgentWorkspaceConnected() OrElse
               HasActiveScheduledTaskWorkspace()
    End Function

    ''' <summary>
    ''' Gets the list of available tools for browser display.
    ''' </summary>
    Private Function GetToolListForBrowser(Optional includeInteractiveM365Tools As Boolean = False) As List(Of Object)
        Dim result As New List(Of Object)()

        Try
            Dim st = LoadInkyState()

            Dim selectedNameSet As New HashSet(Of String)(
                ResolveLocalChatSelectedMainToolNames(st, includeInteractiveM365Tools),
                StringComparer.OrdinalIgnoreCase)

            For Each t In GetLocalChatMainSelectableTools(includeInteractiveM365Tools)
                If t Is Nothing OrElse String.IsNullOrWhiteSpace(t.ToolName) Then Continue For

                result.Add(New With {
                    .name = t.ToolName,
                    .description = If(t.ModelDescription, t.ToolInstructionsPrompt),
                    .selected = selectedNameSet.Contains(t.ToolName)
                })
            Next
        Catch
        End Try

        Return result
    End Function

    ' Builds the entire HTML UI (single file; no external assets)
    ''' <summary>
    ''' Generates full HTML (inline CSS + JS) for the chat UI page.
    ''' </summary>
    Private Function BuildInkyHtmlPage() As System.String
        Dim botName As String = GetBotName()
        Dim brandName As String = If(Not String.IsNullOrWhiteSpace(AN), AN, botName)
        Dim logoUrl As String = GetLogoDataUrl()
        Dim greet As String = GetFriendlyGreeting()

        Dim html As New System.Text.StringBuilder()

        html.AppendLine("<!doctype html>")
        html.AppendLine("<html lang=""en""><head><meta charset=""utf-8"">")
        html.AppendLine("<meta name=""viewport"" content=""width=device-width, initial-scale=1"">")
        html.AppendLine("<link rel=""shortcut icon"" type=""image/png"" href=""" & System.Net.WebUtility.HtmlEncode(logoUrl) & """>")
        html.AppendLine("<link rel=""icon"" type=""image/png"" href=""" & System.Net.WebUtility.HtmlEncode(logoUrl) & """>")
        html.AppendLine("<title>" & System.Net.WebUtility.HtmlEncode(brandName) & " — Local Chat</title>")

        ' CSS
        html.AppendLine("<style>")
        html.AppendLine(":root{--bg:#0b0f14;--card:#11161d;--fg:#e8eef6;--muted:#9aa8b7;--border:#1b2430;--border-strong:#2d3744;--elev:#1a222c;--press-shadow:inset 0 2px 6px rgba(0,0,0,.45);} ")
        html.AppendLine(":root.light{--bg:#f6f7f9;--card:#ffffff;--fg:#0e1116;--muted:#5d6a77;--border:#e2e5e9;--border-strong:#c9cfd6;--elev:#eef1f4;--press-shadow:inset 0 2px 5px rgba(0,0,0,.08);} ")
        html.AppendLine("html,body{height:100%;margin:0;font-family:system-ui,Segoe UI,Roboto,Arial,sans-serif;background:var(--bg);color:var(--fg);} ")
        html.AppendLine(".wrap{display:flex;flex-direction:column;height:100%;} ")
        html.AppendLine(".topbar{display:flex;gap:.5rem;align-items:center;padding:.75rem 1rem;border-bottom:1px solid var(--border);background:var(--card);position:sticky;top:0;z-index:5;flex-wrap:nowrap;overflow:hidden;} ")
        html.AppendLine(".topline{display:flex;align-items:center;gap:.6rem;min-width:0;} ")
        html.AppendLine(".topline img.logo{width:24px;height:24px;border-radius:6px;display:block} ")
        html.AppendLine(".topline .brandbig{font-weight:700;white-space:nowrap;} ")
        html.AppendLine(".topline .sub{color:var(--muted);font-size:.9rem;white-space:nowrap;} ")
        html.AppendLine(".spacer{flex:0 0 0;} ")
        html.AppendLine("select,button,input,textarea{background:var(--card);color:var(--fg);border:1px solid var(--border);border-radius:.6rem;font:inherit;} ")
        html.AppendLine("select,button,input{padding:.5rem .7rem;} ")
        html.AppendLine("button{cursor:pointer;transition:background .16s,filter .12s,transform .08s,box-shadow .18s;} ")
        html.AppendLine("button:hover{filter:brightness(1.07)} ")
        html.AppendLine("button:disabled{opacity:.5;cursor:not-allowed} ")
        html.AppendLine("button.is-pressed,.chatTab.is-pressed{transform:translateY(1px);box-shadow:var(--press-shadow);filter:brightness(.92);} ")
        html.AppendLine("button:active:not(:disabled){transform:translateY(1px);box-shadow:var(--press-shadow);filter:brightness(.9);} ")
        html.AppendLine(".chat{flex:1;overflow-y:auto;overflow-x:hidden;padding:1rem;box-sizing:border-box;} ")
        html.AppendLine(".row{display:flex;margin:0 auto 1rem auto;max-width:100%;width:100%;padding:0 .25rem;box-sizing:border-box;min-width:0;} ")
        html.AppendLine(".row.bot{justify-content:flex-start} ")
        html.AppendLine(".row.user{justify-content:flex-end} ")
        html.AppendLine(".bubble{max-width:85%;min-width:0;box-sizing:border-box;padding:1rem;border:1px solid var(--border);background:var(--card);border-radius:1rem;box-shadow:0 1px 3px rgba(0,0,0,.25);overflow-wrap:anywhere;word-break:break-word;} ")
        html.AppendLine(".bot .bubble{border-top-right-radius:.35rem} ")
        html.AppendLine(".user .bubble{border-top-left-radius:.35rem} ")
        html.AppendLine(".role{font-size:.75rem;color:var(--muted);margin-bottom:.25rem} ")
        html.AppendLine(".inputbar{display:flex;gap:.5rem;padding:1rem;border-top:1px solid var(--border);background:var(--card)} ")
        html.AppendLine("textarea{flex:1;resize:vertical;min-height:52px;max-height:220px;border-radius:.8rem;padding:.75rem;line-height:1.25;} ")
        html.AppendLine(".hint{font-size:.7rem;letter-spacing:.3px;color:var(--muted);padding:.25rem 1rem 1rem} ")
        html.AppendLine("a{color:inherit;text-decoration:underline;text-decoration-color:rgba(255,255,255,.35)} ")
        html.AppendLine(":root.light a{text-decoration-color:rgba(0,0,0,.4)} ")
        html.AppendLine("a:hover{filter:brightness(1.15)} ")
        html.AppendLine("code,pre{font-family:ui-monospace,Consolas,monospace;font-size:.85rem} ")
        html.AppendLine("pre{overflow-x:auto;max-width:100%;padding:.75rem;border:1px solid var(--border);border-radius:.6rem;position:relative;background:var(--elev);box-sizing:border-box;} ")
        html.AppendLine("table{display:block;max-width:100%;overflow-x:auto;border-collapse:collapse;margin:6px 0;} ")
        html.AppendLine(".code-copy-btn{position:absolute;top:6px;right:6px;padding:4px 8px;font-size:.65rem;line-height:1;border:1px solid var(--border);border-radius:4px;background:rgba(0,0,0,.45);backdrop-filter:blur(3px);cursor:pointer;display:flex;align-items:center;gap:6px;color:var(--fg);opacity:0;transition:opacity .18s,background .18s;} ")
        html.AppendLine("pre:hover .code-copy-btn{opacity:1} ")
        html.AppendLine(".code-copy-btn svg{width:16px;height:16px;display:block} ")
        html.AppendLine(".code-copy-btn.copied{background:#2c3440;color:#fff} ")
        html.AppendLine(":root.light .code-copy-btn.copied{background:#d5d9dd;color:#111} ")
        html.AppendLine(".code-copy-btn:focus{outline:2px solid var(--border-strong);} ")
        html.AppendLine(".chatTab{padding:.45rem .55rem;min-width:32px;font-size:.7rem;font-weight:600;line-height:1;border:1px solid var(--border);background:var(--card);color:var(--muted);transition:background .18s,border-color .18s,color .18s,transform .08s,box-shadow .18s;flex-shrink:0;} ")
        html.AppendLine(".chatTab:hover:not(:disabled){background:var(--elev);color:var(--fg);} ")
        html.AppendLine(".chatTab.active{background:#222b35;border-color:var(--border-strong);color:#fff;box-shadow:inset 0 0 0 1px #303c46;} ")
        html.AppendLine(":root.light .chatTab.active{background:#e2e5e9;border-color:var(--border-strong);color:#0e1116;box-shadow:inset 0 0 0 1px #c9cfd6;} ")
        html.AppendLine(".chatTab:focus{outline:2px solid var(--border-strong);outline-offset:1px;} ")

        html.AppendLine(".typing-dots{display:inline-flex;gap:6px;align-items:center;} ")
        html.AppendLine(".typing-dots span{width:7px;height:7px;border-radius:50%;background:currentColor;opacity:.35;animation:tdots 1.2s infinite ease-in-out;} ")
        html.AppendLine(".typing-dots span:nth-child(2){animation-delay:.2s} ")
        html.AppendLine(".typing-dots span:nth-child(3){animation-delay:.4s} ")
        html.AppendLine("@keyframes tdots{0%,80%,100%{transform:translateY(0);opacity:.3}40%{transform:translateY(-5px);opacity:.85}} ")
        html.AppendLine(".typing-elapsed{margin-left:8px;font-size:.65rem;color:var(--muted);font-family:ui-monospace,monospace;opacity:.8;} ")
        html.AppendLine(".actions{display:flex;flex-direction:row;gap:.5rem;align-items:stretch;} ")
        html.AppendLine(".actions .stack{display:flex;flex-direction:column;gap:.5rem;} ")
        html.AppendLine("#cancelBtn{display:none;align-self:stretch;height:auto;} ")

        html.AppendLine("#modelSel{flex:1 1 260px;max-width:420px;min-width:110px;overflow:hidden;white-space:nowrap;text-overflow:ellipsis;box-sizing:border-box;} ")
        html.AppendLine("#modelSel.squeezed{max-width:55vw;} ")
        html.AppendLine(".topbar button,#modelSel{flex-shrink:0;} ")
        html.AppendLine("@media (max-width:1000px){#modelSel{max-width:360px;}} ")
        html.AppendLine("@media (max-width:880px){.topline .sub{display:none;}} ")
        html.AppendLine("@media (max-width:760px){.topline .brandbig{max-width:140px;overflow:hidden;text-overflow:ellipsis;}#modelSel{max-width:300px;}} ")
        html.AppendLine("@media (max-width:640px){#modelSel{max-width:55vw;} .topline .brandbig{max-width:110px;}} ")
        html.AppendLine("@media (max-width:640px){.actions{flex-direction:column;}.actions .stack{flex-direction:row;}.actions .stack button{flex:1;}#cancelBtn{align-self:auto;height:auto;}} ")

        ' stack "Use sources" + "Tooling log" under the sources button without growing the overall topbar height
        html.AppendLine(".toolingSlot{display:none;flex-direction:column;justify-content:center;gap:2px;line-height:1;flex-shrink:0;min-width:max-content;} ")
        html.AppendLine(".toolingRow{display:flex;align-items:center;gap:6px;white-space:nowrap;} ")
        html.AppendLine(".toolingRow label{font-size:.8rem;color:var(--muted);cursor:pointer;user-select:none;} ")
        html.AppendLine(".toolingRow input{margin:0;} ")
        html.AppendLine("#toolLogBtn{flex-shrink:0;} ")
        html.AppendLine("#toolLogBtn.active{background:#222b35;border-color:var(--border-strong);color:#fff;box-shadow:inset 0 0 0 1px #303c46;} ")
        html.AppendLine(":root.light #toolLogBtn.active{background:#e2e5e9;border-color:var(--border-strong);color:#0e1116;box-shadow:inset 0 0 0 1px #c9cfd6;} ")
        html.AppendLine("#agentModelBtn{flex-shrink:0;display:none;line-height:1;align-items:center;justify-content:center;padding:.45rem .5rem;transition:background .18s,border-color .18s,color .18s,transform .08s,box-shadow .18s;} ")
        html.AppendLine("#agentModelBtn.active{background:#1a5276;border-color:#2980b9;color:#fff;box-shadow:inset 0 0 0 1px #2980b9;} ")
        html.AppendLine(":root.light #agentModelBtn.active{background:#d4e6f1;border-color:#2980b9;color:#1a5276;box-shadow:inset 0 0 0 1px #2980b9;} ")
        html.AppendLine("#agentFiles{font-size:.7rem;color:var(--muted);padding:0 1rem .5rem;display:none;} ")
        html.AppendLine("#agentFiles .file-tag{display:inline-flex;align-items:center;gap:6px;background:var(--elev);border:1px solid var(--border);border-radius:999px;padding:2px 4px 2px 8px;margin:2px;font-size:.65rem;} ")
        html.AppendLine("#agentFiles .file-tag-name{display:inline-block;} ")
        html.AppendLine("#agentFiles .file-tag-remove{display:inline-flex;align-items:center;justify-content:center;width:18px;height:18px;padding:0;border:1px solid var(--border);border-radius:999px;background:transparent;color:var(--muted);font-size:.9rem;line-height:1;cursor:pointer;} ")
        html.AppendLine("#agentFiles .file-tag-remove:hover{background:var(--card);color:var(--fg);} ")
        html.AppendLine("#agentWorkspace{font-size:.7rem;color:var(--muted);padding:0 1rem .35rem;display:none;border-top:1px solid var(--border);background:var(--bg);} ")
        html.AppendLine("#agentWorkspace .ws-title{font-weight:600;color:var(--fg);margin-right:6px;} ")
        html.AppendLine("#agentWorkspace .ws-path{opacity:.8;word-break:break-all;} ")
        html.AppendLine("#agentWorkspace button{font-size:.65rem;padding:2px 6px;margin-left:4px;} ")
        html.AppendLine(".inky-modal-backdrop{position:fixed;inset:0;background:rgba(0,0,0,.45);display:none;align-items:center;justify-content:center;z-index:50;padding:20px;} ")
        html.AppendLine(".inky-modal-backdrop.show{display:flex;} ")
        html.AppendLine(".inky-modal{width:min(560px,96vw);max-height:90vh;overflow:auto;background:var(--card);color:var(--fg);border:1px solid var(--border-strong);border-radius:14px;box-shadow:0 18px 48px rgba(0,0,0,.35);} ")
        html.AppendLine(".inky-modal-hd{display:flex;align-items:center;justify-content:space-between;padding:14px 16px;border-bottom:1px solid var(--border);} ")
        html.AppendLine(".inky-modal-ttl{font-weight:700;font-size:1rem;} ")
        html.AppendLine(".inky-modal-close{padding:.35rem .55rem;line-height:1;} ")
        html.AppendLine(".inky-modal-bd{padding:16px;} ")
        html.AppendLine(".inky-modal-ft{display:flex;justify-content:flex-end;gap:8px;padding:14px 16px;border-top:1px solid var(--border);} ")
        html.AppendLine(".inky-form-row{margin-bottom:14px;} ")
        html.AppendLine(".inky-form-row:last-child{margin-bottom:0;} ")
        html.AppendLine(".inky-form-row label.hd{display:block;font-size:.82rem;font-weight:600;color:var(--fg);margin-bottom:6px;} ")
        html.AppendLine(".inky-help{font-size:.74rem;color:var(--muted);line-height:1.35;} ")
        html.AppendLine(".inky-radio-list,.inky-check-list{display:flex;flex-direction:column;gap:8px;} ")
        html.AppendLine(".inky-radio-list label,.inky-check-list label{display:flex;align-items:flex-start;gap:8px;font-size:.85rem;color:var(--fg);} ")
        html.AppendLine(".inky-radio-list input,.inky-check-list input{margin-top:2px;} ")
        html.AppendLine(".inky-pathbox{padding:10px 12px;border:1px solid var(--border);border-radius:10px;background:var(--elev);font-size:.8rem;word-break:break-all;} ")
        html.AppendLine(".inky-danger{color:#d35454;} ")
        html.AppendLine(".inky-inline-actions{display:flex;gap:8px;flex-wrap:wrap;margin-top:10px;} ")
        html.AppendLine(".inky-message-box{padding:10px 12px;border:1px solid var(--border);border-radius:10px;background:var(--elev);font-size:.82rem;line-height:1.45;white-space:pre-wrap;word-break:break-word;} ")
        html.AppendLine(".inky-message-box.error{border-color:#8f3d3d;background:rgba(170,56,56,.12);color:#ffd7d7;} ")
        html.AppendLine(":root.light .inky-message-box.error{border-color:#d7a2a2;background:#fff1f1;color:#7a1f1f;} ")
        html.AppendLine("</style>")
        html.AppendLine("</head><body>")
        html.AppendLine("<div class=""wrap"">")

        ' Top bar
        html.AppendLine("  <div class=""topbar"">")
        html.AppendLine("    <div class=""topline"">")
        If Not String.IsNullOrWhiteSpace(logoUrl) Then
            html.AppendLine("      <img class=""logo"" src=""" & System.Net.WebUtility.HtmlEncode(logoUrl) & """ alt=""logo"">")
        End If
        html.AppendLine("      <div class=""brandbig"">" & System.Net.WebUtility.HtmlEncode(brandName) & "</div>")
        html.AppendLine("      <div class=""sub"">Local Chat</div>")
        html.AppendLine("    </div>")
        html.AppendLine("    <div class=""spacer""></div>")
        html.AppendLine("    <select id=""modelSel"" title=""Model""></select>")
        html.AppendLine("    <button id=""agentModelBtn"" title=""Toggle agent model (AgentDefaultModel) — auto-switches model, enables agent mode with all tools"" style=""display:none;line-height:1;align-items:center;justify-content:center;padding:.45rem .5rem;""><svg viewBox=""0 0 24 24"" width=""18"" height=""18"" fill=""none"" stroke=""currentColor"" stroke-width=""2"" stroke-linecap=""round"" stroke-linejoin=""round""><rect x=""3"" y=""4"" width=""18"" height=""12"" rx=""2""/><line x1=""7"" y1=""20"" x2=""17"" y2=""20""/><line x1=""12"" y1=""16"" x2=""12"" y2=""20""/><circle cx=""9"" cy=""10"" r=""1""/><circle cx=""15"" cy=""10"" r=""1""/></svg></button>")
        html.AppendLine("    <button id=""copyBtn"" title=""Copy last answer to clipboard"">Copy last</button>")
        html.AppendLine("    <button id=""toWordBtn"" title=""Move this chat thread into a new Word document"">To Word</button>")
        html.AppendLine("    <button id=""clearBtn"" title=""Clear current conversation"">Clear</button>")
        html.AppendLine("    <button id=""chat1Btn"" class=""chatTab"" data-chat=""1"" title=""Chat 1"">1</button>")
        html.AppendLine("    <button id=""chat2Btn"" class=""chatTab"" data-chat=""2"" title=""Chat 2"">2</button>")
        html.AppendLine("    <button id=""toolsBtn"" title=""Select " & System.Net.WebUtility.HtmlEncode(ToolFriendlyName) & """ style=""display:none;"">" & System.Net.WebUtility.HtmlEncode(ToolFriendlyName) & "</button>")

        html.AppendLine("    <div id=""toolingSlot"" class=""toolingSlot"">")
        html.AppendLine("      <div class=""toolingRow"">")
        html.AppendLine("        <input type=""checkbox"" id=""toolingChk"" title=""Enable the selected " & System.Net.WebUtility.HtmlEncode(ToolFriendlyName.ToLower()) & """>")
        html.AppendLine("        <label for=""toolingChk"" id=""toolingLbl"">Use " & System.Net.WebUtility.HtmlEncode(ToolFriendlyName.ToLower()) & "</label>")
        html.AppendLine("      </div>")
        html.AppendLine("      <div class=""toolingRow"">")
        html.AppendLine("        <input type=""checkbox"" id=""advancedToolsChk""title=""" & System.Net.WebUtility.HtmlEncode(AllToolUseDescription) & """>")
        html.AppendLine("        <label for=""advancedToolsChk"" id=""advancedToolsLbl"">" & System.Net.WebUtility.HtmlEncode(AllToolUse) & "</label>")
        html.AppendLine("      </div>")
        html.AppendLine("    </div>")
        html.AppendLine("    <button id=""toolLogBtn"" title=""Toggle tooling log window"" style=""display:none;line-height:1;align-items:center;justify-content:center;padding:.45rem .5rem;""><svg viewBox=""0 0 24 24"" width=""16"" height=""16"" fill=""none"" stroke=""currentColor"" stroke-width=""2"" stroke-linecap=""round"" stroke-linejoin=""round""><path d=""M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z""/><polyline points=""14 2 14 8 20 8""/><line x1=""16"" y1=""13"" x2=""8"" y2=""13""/><line x1=""16"" y1=""17"" x2=""8"" y2=""17""/><polyline points=""10 9 9 9 8 9""/></svg></button>")

        html.AppendLine("    <div class=""toolingRow"" id=""memorySlot"" style=""display:inline-flex;align-items:center;gap:6px;flex-shrink:0;"">")
        html.AppendLine("      <input type=""checkbox"" id=""memoryChk"" title=""Enable Inky Memory — persistent cross-session learning"">")
        html.AppendLine("      <label for=""memoryChk"" style=""font-size:.8rem;color:var(--muted);cursor:pointer;user-select:none;white-space:nowrap;"">Memory</label>")
        html.AppendLine("      <a href=""#"" id=""memoryEditLnk"" title=""Edit the Inky Memory file"" style=""font-size:.7rem;color:var(--muted);text-decoration:underline;cursor:pointer;white-space:nowrap;display:none;"">Edit</a>")
        html.AppendLine("    </div>")
        html.AppendLine("    <button id=""playBtn"" title=""Open mini games"" style=""line-height:1;display:flex;align-items:center;justify-content:center;""><svg viewBox=""0 0 24 24"" width=""18"" height=""18"" fill=""none"" stroke=""currentColor"" stroke-width=""2"" stroke-linecap=""round"" stroke-linejoin=""round""><rect x=""3"" y=""8"" width=""18"" height=""8"" rx=""4"" ry=""4""/><circle cx=""8"" cy=""12"" r=""1""/><circle cx=""12"" cy=""12"" r=""1""/><circle cx=""16"" cy=""12"" r=""1""/></svg></button>")
        html.AppendLine("    <button id=""themeBtn"" title=""Toggle theme"" style=""line-height:1;display:flex;align-items:center;justify-content:center;""><svg id=""themeIcon"" viewBox=""0 0 24 24"" width=""18"" height=""18"" fill=""none"" stroke=""currentColor"" stroke-width=""2"" stroke-linecap=""round"" stroke-linejoin=""round""><circle cx=""12"" cy=""12"" r=""5""/><line x1=""12"" y1=""1"" x2=""12"" y2=""3""/><line x1=""12"" y1=""21"" x2=""12"" y2=""23""/><line x1=""4.22"" y1=""4.22"" x2=""5.64"" y2=""5.64""/><line x1=""18.36"" y1=""18.36"" x2=""19.78"" y2=""19.78""/><line x1=""1"" y1=""12"" x2=""3"" y2=""12""/><line x1=""21"" y1=""12"" x2=""23"" y2=""12""/><line x1=""4.22"" y1=""19.78"" x2=""5.64"" y2=""18.36""/><line x1=""18.36"" y1=""5.64"" x2=""19.78"" y2=""4.22""/></svg></button>")
        html.AppendLine("  </div>")

        html.AppendLine("  <div id=""chat"" class=""chat""></div>")

        html.AppendLine("  <div class=""inputbar"">")
        html.AppendLine("    <textarea id=""msg"" placeholder=""" & System.Net.WebUtility.HtmlEncode(greet) & """ autofocus></textarea>")
        html.AppendLine("    <div class=""actions"">" &
                    "<div class=""stack"">" &
                        "<button id=""sendBtn"">Send</button>" &
                        "<button id=""pureBtn"" title=""Send only this raw text (no system prompt, no history)"">Pure</button>" &
                    "</div>" &
                    "<button id=""cancelBtn"" style=""display:none;"">Cancel</button>" &
                "</div>")
        html.AppendLine("  </div>")
        html.AppendLine("  <div id=""agentWorkspace""></div>")
        html.AppendLine("  <div id=""agentFiles""></div>")
        html.AppendLine("  <div id=""workspaceModalBackdrop"" class=""inky-modal-backdrop"">")
        html.AppendLine("    <div class=""inky-modal"" role=""dialog"" aria-modal=""true"" aria-labelledby=""workspaceModalTitle"">")
        html.AppendLine("      <div class=""inky-modal-hd"">")
        html.AppendLine("        <div id=""workspaceModalTitle"" class=""inky-modal-ttl"">Agent Workspace</div>")
        html.AppendLine("        <button id=""workspaceModalCloseBtn"" class=""inky-modal-close"" type=""button"" title=""Close"">✕</button>")
        html.AppendLine("      </div>")
        html.AppendLine("      <div id=""workspaceModalBody"" class=""inky-modal-bd""></div>")
        html.AppendLine("      <div class=""inky-modal-ft"">")
        html.AppendLine("        <button id=""workspaceModalCancelBtn"" type=""button"">Cancel</button>")
        html.AppendLine("        <button id=""workspaceModalOkBtn"" type=""button"">Save</button>")
        html.AppendLine("      </div>")
        html.AppendLine("    </div>")
        html.AppendLine("  </div>")

        ' Build dynamic hint text
        Dim hintParts As New System.Collections.Generic.List(Of String)()
        hintParts.Add("Drag &amp; drop a file (only visible to the chatbot for the current prompt)")
        hintParts.Add("Enter=send")
        hintParts.Add("Shift+Enter=newline")
        hintParts.Add("Ctrl+L=clear")
        If INI_PromptLib Then
            hintParts.Add("/=insert prompt library entry")
        End If

        ' Check if (t) trigger is available (ToolDefaultModel defined in INI)
        Dim toolTriggerAvailable As Boolean = False
        Try
            If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                toolTriggerAvailable = GetSpecialTaskModel(_context, INI_AlternateModelPath, "ToolDefaultModel")
                If toolTriggerAvailable Then
                    ' Immediately restore — we only wanted to check availability
                    If originalConfigLoaded Then
                        RestoreDefaults(_context, originalConfig)
                    End If
                    originalConfigLoaded = False
                End If
            End If
        Catch
        End Try

        If toolTriggerAvailable Then
            hintParts.Add(ToolTrigger & "=process this in agent mode")
        End If

        html.AppendLine("  <div class=""hint"">" & String.Join(" • ", hintParts) & "</div>")
        html.AppendLine("</div>")

        ' JS
        html.AppendLine("<script>")
        html.AppendLine("window.__botName=" & Newtonsoft.Json.JsonConvert.SerializeObject(botName) & ";")
        html.AppendLine("let __supportsFiles=false;")
        html.AppendLine("let __pendingFilePath='';")
        html.AppendLine("let dark=false;")
        html.AppendLine("let __currentJobId=null;")
        html.AppendLine("let __jobCanceled=false;")
        html.AppendLine("let __typingBubbleId=null;")
        html.AppendLine("let __jobStartTs=0;")
        html.AppendLine("let __elapsedTimer=null;")
        html.AppendLine("let __lastPrompt=" & Newtonsoft.Json.JsonConvert.SerializeObject(My.Settings.Inky_LastPrompt) & ";")
        html.AppendLine("let __pendingScheduledTaskId='';")

        ' Press feedback
        html.AppendLine("(function(){const pressOn=e=>{const b=e.target.closest('button');if(!b||b.disabled)return;b.classList.add('is-pressed');};const pressOff=()=>{document.querySelectorAll('button.is-pressed').forEach(b=>b.classList.remove('is-pressed'));};['mousedown','touchstart'].forEach(ev=>document.addEventListener(ev,pressOn,{passive:true}));['mouseup','mouseleave','blur'].forEach(ev=>document.addEventListener(ev,pressOff));document.addEventListener('keydown',e=>{if((e.key===' '||e.key==='Enter')){const b=e.target.closest('button');if(b&&!b.disabled)b.classList.add('is-pressed');}});document.addEventListener('keyup',e=>{if(e.key===' '||e.key==='Enter')pressOff();});})();")

        ' Helpers
        html.AppendLine("function copyText(t){if(navigator.clipboard){return navigator.clipboard.writeText(t);}return new Promise((res,rej)=>{try{const ta=document.createElement('textarea');ta.value=t;ta.style.position='fixed';ta.style.left='-9999px';document.body.appendChild(ta);ta.select();document.execCommand('copy');ta.remove();res();}catch(e){rej(e);}});}")
        html.AppendLine("function enhanceCodeBlocks(scope){(scope||document).querySelectorAll('pre').forEach(pre=>{if(pre.dataset.enhanced==='1')return;const btn=document.createElement('button');btn.type='button';btn.className='code-copy-btn';btn.innerHTML='<svg viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" stroke-width=""2"" stroke-linecap=""round"" stroke-linejoin=""round""><rect x=""9"" y=""9"" width=""13"" height=""13"" rx=""2"" ry=""2""/><path d=""M5 15H4a2 2 0 0 1-2-2V4c0-1.1.9-2 2-2h9a2 2 0 0 1 2 2v1""/></svg>';btn.addEventListener('click',()=>{const code=pre.querySelector('code');const txt=code?code.innerText:pre.innerText;copyText(txt).then(()=>{btn.classList.add('copied');setTimeout(()=>btn.classList.remove('copied'),1500);});});pre.appendChild(btn);pre.dataset.enhanced='1';});}")
        html.AppendLine("const api=async(cmd,data={})=>{try{const r=await fetch('/inky/api',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(Object.assign({Command:cmd},data))});const txt=await r.text();try{return JSON.parse(txt);}catch{return{ok:false,error:txt}}}catch(e){return{ok:false,error:e.message||'Network error'}}};")
        html.AppendLine("function isPromptLibrarySlashTrigger(){const pos=msgEl.selectionStart||0;if(pos<=0)return true;const prev=msgEl.value.charAt(pos-1);return /\s/.test(prev);} ")
        html.AppendLine("function insertPromptIntoMessage(text){const start=msgEl.selectionStart||0;const end=msgEl.selectionEnd||start;msgEl.setRangeText(String(text||''),start,end,'end');msgEl.focus();} ")
        html.AppendLine("async function openPromptLibrary(){const r=await api('inky_promptlibpick');if(!r||!r.ok){if(r&&r.error)alert(r.error||'Prompt library failed');return;}if(r.prompt){insertPromptIntoMessage(r.prompt);}};")

        html.AppendLine("const chatEl=document.getElementById('chat');")
        html.AppendLine("const msgEl=document.getElementById('msg');")
        html.AppendLine("const modelSel=document.getElementById('modelSel');")
        html.AppendLine("const copyBtn=document.getElementById('copyBtn');")
        html.AppendLine("const toWordBtn=document.getElementById('toWordBtn');")
        html.AppendLine("const clearBtn=document.getElementById('clearBtn');")
        html.AppendLine("const themeBtn=document.getElementById('themeBtn');")
        html.AppendLine("const playBtn=document.getElementById('playBtn');")
        html.AppendLine("const cancelBtn=document.getElementById('cancelBtn');")
        html.AppendLine("const chat1Btn=document.getElementById('chat1Btn');")
        html.AppendLine("const chat2Btn=document.getElementById('chat2Btn');")
        html.AppendLine("const sendBtn=document.getElementById('sendBtn');")
        html.AppendLine("const pureBtn=document.getElementById('pureBtn');")
        html.AppendLine("const topbar=document.querySelector('.topbar');")
        html.AppendLine("const toolsBtn=document.getElementById('toolsBtn');")
        html.AppendLine("const toolingSlot=document.getElementById('toolingSlot');")
        html.AppendLine("const toolingChk=document.getElementById('toolingChk');")
        html.AppendLine("const toolingLbl=document.getElementById('toolingLbl');")
        html.AppendLine("const advancedToolsChk=document.getElementById('advancedToolsChk');")
        html.AppendLine("const agentLbl=document.getElementById('agentLbl');")
        html.AppendLine("const toolLogBtn=document.getElementById('toolLogBtn');")
        html.AppendLine("const agentFilesEl=document.getElementById('agentFiles');")
        html.AppendLine("const agentWorkspaceEl=document.getElementById('agentWorkspace');")
        html.AppendLine("const agentModelBtn=document.getElementById('agentModelBtn');")
        html.AppendLine("const workspaceModalBackdrop=document.getElementById('workspaceModalBackdrop');")
        html.AppendLine("const workspaceModalTitle=document.getElementById('workspaceModalTitle');")
        html.AppendLine("const workspaceModalBody=document.getElementById('workspaceModalBody');")
        html.AppendLine("const workspaceModalOkBtn=document.getElementById('workspaceModalOkBtn');")
        html.AppendLine("const workspaceModalCancelBtn=document.getElementById('workspaceModalCancelBtn');")
        html.AppendLine("const workspaceModalCloseBtn=document.getElementById('workspaceModalCloseBtn');")
        html.AppendLine("let __advancedToolsEnabled=false;")
        html.AppendLine("let __agentModelActive=false;")
        html.AppendLine("let __agentModelAvailable=false;")
        html.AppendLine("let __toolLogEnabled=true;")
        html.AppendLine("let __toolingEnabled=false;")
        html.AppendLine("let __modelSupportsTooling=false;")
        html.AppendLine("let __workspaceDialogMode='';")
        html.AppendLine("function syncAdvancedToolsUi(state){state=state||{};if(typeof state.advancedToolsEnabled==='boolean'){__advancedToolsEnabled=!!state.advancedToolsEnabled;advancedToolsChk.checked=__advancedToolsEnabled;}if(typeof state.agentModelAvailable==='boolean'){__agentModelAvailable=!!state.agentModelAvailable;}if(typeof state.agentModelActive==='boolean'){__agentModelActive=!!state.agentModelActive;}updateToolingVisibility();updateAgentModelBtn();const hasWorkspace=Object.prototype.hasOwnProperty.call(state,'agentWorkspace');updateAgentWorkspaceDisplay(hasWorkspace?state.agentWorkspace:null);updateAgentFilesDisplay(Array.isArray(state.agentFiles)?state.agentFiles:[]);applyCoupling();}")

        html.AppendLine("function setTheme(isDark){dark=!!isDark;document.documentElement.classList.toggle('light',!dark);var icon=document.getElementById('themeIcon');if(icon){if(dark){icon.innerHTML='<path d=""M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z""/>';}else{icon.innerHTML='<circle cx=""12"" cy=""12"" r=""5""/><line x1=""12"" y1=""1"" x2=""12"" y2=""3""/><line x1=""12"" y1=""21"" x2=""12"" y2=""23""/><line x1=""4.22"" y1=""4.22"" x2=""5.64"" y2=""5.64""/><line x1=""18.36"" y1=""18.36"" x2=""19.78"" y2=""19.78""/><line x1=""1"" y1=""12"" x2=""3"" y2=""12""/><line x1=""21"" y1=""12"" x2=""23"" y2=""12""/><line x1=""4.22"" y1=""19.78"" x2=""5.64"" y2=""18.36""/><line x1=""18.36"" y1=""5.64"" x2=""19.78"" y2=""4.22""/>';}}} ")
        html.AppendLine("function forceExternalLinks(scope){try{(scope||document).querySelectorAll('a[href]').forEach(a=>{a.target='_blank';a.rel='noopener noreferrer';});}catch{}}")
        html.AppendLine("function setActiveChatBtn(id){document.querySelectorAll('.chatTab').forEach(b=>b.classList.toggle('active',b.dataset.chat==String(id)));}")
        html.AppendLine("function disableChatSwitch(dis){chat1Btn.disabled=dis;chat2Btn.disabled=dis;}")

        ' Tooltip updates & responsive adjustment
        html.AppendLine("function updateModelTooltip(){try{if(!modelSel) return;const opt=modelSel.options[modelSel.selectedIndex];if(opt){modelSel.title=opt.textContent||'Model';}}catch{}}")
        html.AppendLine("function adjustModelSel(){if(!topbar) return;requestAnimationFrame(()=>{if(topbar.scrollWidth>topbar.clientWidth){modelSel.classList.add('squeezed');}else{modelSel.classList.remove('squeezed');}});}")
        html.AppendLine("window.addEventListener('resize',adjustModelSel);")

        html.AppendLine("function render(turns){chatEl.innerHTML='';for(const t of (turns||[])){const row=document.createElement('div');row.className='row '+(t.role==='user'?'user':'bot');const bub=document.createElement('div');bub.className='bubble';const rl=document.createElement('div');rl.className='role';rl.textContent=(t.role==='user'?'You':(window.__botName||'Bot'));bub.appendChild(rl);const cont=document.createElement('div');if(t && t.html){cont.innerHTML=t.html;forceExternalLinks(cont);}else if(t && t.markdown){const safe=t.markdown.replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('\n','<br>');cont.innerHTML=safe;}bub.appendChild(cont);row.appendChild(bub);chatEl.appendChild(row);}chatEl.scrollTop=chatEl.scrollHeight;enhanceCodeBlocks(chatEl);} ")
        html.AppendLine("function addTempAssistantBubble(html){const id='tmp-'+Math.random().toString(36).slice(2);chatEl.insertAdjacentHTML('beforeend',`<div class=""row bot"" id=""${id}""><div class=""bubble""><div class=""role"">${window.__botName||'Bot'}</div><div class=""tmpContent"">${html}</div></div></div>`);chatEl.scrollTop=chatEl.scrollHeight;return id;}")
        html.AppendLine("function removeTempBubble(id){const el=document.getElementById(id);if(el)el.remove();}")
        html.AppendLine("function replaceAssistantBubble(id,html){const row=document.getElementById(id);if(!row)return;const c=row.querySelector('.tmpContent');if(c){c.innerHTML=html;forceExternalLinks(row);enhanceCodeBlocks(row);}}")

        ' Typing + elapsed
        html.AppendLine("function ensureTypingBubble(){if(__typingBubbleId)return;const content='<div class=""typing-container""><span class=""typing-dots""><span></span><span></span><span></span></span><span id=""typingElapsed"" class=""typing-elapsed"" style=""display:none;"">(0s)</span></div>';__typingBubbleId=addTempAssistantBubble(content);}")
        html.AppendLine("function updateElapsed(){if(!__typingBubbleId)return;const el=document.getElementById('typingElapsed');if(!el)return;const sec=Math.floor((Date.now()-__jobStartTs)/1000);if(sec>=10){el.style.display='inline-block';el.textContent='(' + sec + 's)';}}")
        html.AppendLine("function startElapsedTimer(){stopElapsedTimer();__jobStartTs=Date.now();__elapsedTimer=setInterval(updateElapsed,1000);}")
        html.AppendLine("function stopElapsedTimer(){if(__elapsedTimer){clearInterval(__elapsedTimer);__elapsedTimer=null;}const el=document.getElementById('typingElapsed');if(el)el.style.display='none';}")
        html.AppendLine("function removeTypingBubble(){if(__typingBubbleId){removeTempBubble(__typingBubbleId);__typingBubbleId=null;}stopElapsedTimer();}")

        ' Boot        
        html.AppendLine("async function claimScheduledTask(){const r=await api('inky_claimscheduledtask');if(!r||!r.ok)return null;return r.task||null;}")
        html.AppendLine("async function boot(){const st=await api('inky_getstate');if(!st.ok){alert(st.error||'Init failed');return;}__supportsFiles=(st.supportsFiles===true);setTheme(st.darkMode!==false);render(st.history||[]);modelSel.innerHTML='';for(const m of (st.models||[])){const o=document.createElement('option');o.value=m.key||'';o.textContent=m.label||'';o.disabled=!!m.disabled;o.title=o.textContent;if(m.selected&&!o.disabled)o.selected=true;modelSel.appendChild(o);}if(!modelSel.value){const fe=[...modelSel.options].find(o=>!o.disabled&&o.value);if(fe)fe.selected=true;}updateModelTooltip();if(st.greeting&&(!Array.isArray(st.history)||st.history.length===0)){msgEl.placeholder=st.greeting;}setActiveChatBtn(st.activeChat||1);__modelSupportsTooling=(st.supportsTooling===true);__toolingEnabled=(st.toolingEnabled===true);toolingChk.checked=__toolingEnabled;__toolLogEnabled=(st.toolingLogEnabled!==false);toolLogBtn.classList.toggle('active',__toolLogEnabled);__memoryEnabled=(st.inkyMemoryEnabled===true);memoryChk.checked=__memoryEnabled;memoryEditLnk.style.display=__memoryEnabled?'inline':'none';syncAdvancedToolsUi({advancedToolsEnabled:st.advancedToolsEnabled===true,agentWorkspace:st.agentWorkspace,agentFiles:st.agentFiles||[],agentModelAvailable:st.agentModelAvailable===true,agentModelActive:st.agentModelActive===true});adjustModelSel();const scheduledTask=await claimScheduledTask();if(scheduledTask&&scheduledTask.prompt&&!__currentJobId){__pendingScheduledTaskId=scheduledTask.taskId||'';msgEl.value=scheduledTask.prompt;await send();}}")

        ' Poll job
        html.AppendLine("function buildAssistantTurnFromJobResult(r){const md=String((r&&r.result)||'').trim();if(!md)return null;const html=String((r&&r.resultHtml)||'').trim()||md.replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('\n','<br>');return {role:'assistant',markdown:md,html:html,utc:new Date().toISOString()};}")
        html.AppendLine("function ensureJobResultVisible(st,r){const hist=(st&&Array.isArray(st.history))?st.history.slice():[];const turn=buildAssistantTurnFromJobResult(r);if(!turn)return hist;const activeChat=Number((st&&st.activeChat)||1);const resultChat=Number((r&&r.chat)||activeChat);if(activeChat!==resultChat)return hist;const last=hist.length?hist[hist.length-1]:null;const lastMd=String((last&&last.markdown)||'').trim();if(last&&last.role==='assistant'&&lastMd===turn.markdown)return hist;hist.push(turn);return hist;}")
        html.AppendLine("async function pollJob(jobId){if(!jobId)return;__currentJobId=jobId;__jobCanceled=false;ensureTypingBubble();startElapsedTimer();cancelBtn.style.display='inline-block';disableChatSwitch(true);try{for(;;){await new Promise(r=>setTimeout(r,2000));if(__jobCanceled)break;const s=await api('inky_jobstatus',{Job:jobId});if(!s.ok){console.warn('job status error',s.error);break;}if(s.status==='running'){continue;}if(s.status==='done'){if(Array.isArray(s.history)){render(s.history);}else{const st=await api('inky_getstate');if(st.ok){const hist=ensureJobResultVisible(st,s);render(hist);}else{render(ensureJobResultVisible({history:[]},s));console.warn('state sync error',st&&st.error);}}const stSync=await api('inky_getstate');if(stSync&&stSync.ok){if(stSync.agentFiles)updateAgentFilesDisplay(stSync.agentFiles);syncAdvancedToolsUi({advancedToolsEnabled:stSync.advancedToolsEnabled===true,agentWorkspace:stSync.agentWorkspace,agentFiles:stSync.agentFiles||[],agentModelAvailable:stSync.agentModelAvailable===true,agentModelActive:stSync.agentModelActive===true});}break;}if(s.status==='canceled'){const st=await api('inky_getstate');if(st.ok){render(st.history||[]);if(st.agentFiles)updateAgentFilesDisplay(st.agentFiles);}break;}if(s.status==='error'){console.warn('job failed',s.error);break;}const st=await api('inky_getstate');if(st.ok){const hist=ensureJobResultVisible(st,s);render(hist);if(st.agentFiles)updateAgentFilesDisplay(st.agentFiles);}break;}}finally{cancelBtn.style.display='none';removeTypingBubble();sendBtn.disabled=false;pureBtn.disabled=false;disableChatSwitch(false);__currentJobId=null;adjustModelSel();}}")

        ' Send (normal)
        html.AppendLine("async function send(){if(__currentJobId){return;}const t=msgEl.value.trim();if(!t)return;const scheduledTaskId=__pendingScheduledTaskId||'';__pendingScheduledTaskId='';__lastPrompt=t;msgEl.value='';sendBtn.disabled=true;pureBtn.disabled=true;chatEl.insertAdjacentHTML('beforeend',`<div class=""row user""><div class=""bubble""><div class=""role"">You</div><div>${t.replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('\n','<br>')}</div></div></div>`);let typingId=addTempAssistantBubble('<span class=""typing-dots""><span></span><span></span><span></span></span>');const payload={Text:t};if(scheduledTaskId)payload.ScheduledTaskId=scheduledTaskId;if(__pendingFilePath)payload.FileObject=__pendingFilePath;let r;try{r=await api('inky_send',payload);}catch(e){r={ok:false,error:e.message||'Network error'};}if(!r||!r.ok){removeTempBubble(typingId);sendBtn.disabled=false;pureBtn.disabled=false;alert(r&&r.error||'Error');__pendingFilePath='';adjustModelSel();return;}__pendingFilePath='';if(r.job){if(r.history){render(r.history||[]);}removeTempBubble(typingId);__typingBubbleId=null;ensureTypingBubble();startElapsedTimer();cancelBtn.style.display='inline-block';disableChatSwitch(true);pollJob(r.job);}else{removeTempBubble(typingId);sendBtn.disabled=false;pureBtn.disabled=false;if(r.history){render(r.history||[]);}adjustModelSel();}}")

        ' PureSend
        html.AppendLine("async function pureSend(){if(__currentJobId){return;}const t=msgEl.value.trim();if(!t)return;__lastPrompt=t;msgEl.value='';sendBtn.disabled=true;pureBtn.disabled=true;chatEl.insertAdjacentHTML('beforeend',`<div class=""row user""><div class=""bubble""><div class=""role"">You</div><div>${('Pure: '+t).replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('\n','<br>')}</div></div></div>`);let typingId=addTempAssistantBubble('<span class=""typing-dots""><span></span><span></span><span></span></span>');const payload={Text:t};if(__pendingFilePath)payload.FileObject=__pendingFilePath;let r;try{r=await api('inky_pure',payload);}catch(e){r={ok:false,error:e.message||'Network error'};}if(!r||!r.ok){removeTempBubble(typingId);sendBtn.disabled=false;pureBtn.disabled=false;alert(r&&r.error||'Error');__pendingFilePath='';adjustModelSel();return;}__pendingFilePath='';if(r.job){if(r.history){render(r.history||[]);}removeTempBubble(typingId);__typingBubbleId=null;ensureTypingBubble();startElapsedTimer();cancelBtn.style.display='inline-block';disableChatSwitch(true);pollJob(r.job);}else{removeTempBubble(typingId);sendBtn.disabled=false;pureBtn.disabled=false;if(r.history){render(r.history||[]);}adjustModelSel();}}")

        ' drag/drop
        html.AppendLine("(function(){const stop=e=>{e.preventDefault();e.stopPropagation();};['dragenter','dragover','dragleave','drop'].forEach(ev=>document.addEventListener(ev,stop,false));document.addEventListener('drop',async e=>{const files=[...(e.dataTransfer&&e.dataTransfer.files)||[]];if(!files.length)return;")
        html.AppendLine("if(__advancedToolsEnabled){")
        html.AppendLine("  const tempId=addTempAssistantBubble(`Loading ${files.length} file(s) for agent mode…`);")
        html.AppendLine("  try{const paths=[];for(const f of files){const fr=new FileReader();const dataUrl=await new Promise((res,rej)=>{fr.onerror=()=>rej(new Error('read error'));fr.onload=()=>res(fr.result);fr.readAsDataURL(f);});const up=await api('inky_upload',{Name:f.name,DataUrl:String(dataUrl||'')});if(up.ok&&up.path)paths.push(up.path);}if(paths.length>0){const r=await api('inky_agentaddfiles',{Paths:paths});if(r.ok){replaceAssistantBubble(tempId,`Loaded ${r.addedCount} file(s) for agent mode.`);updateAgentFilesDisplay(r.files||[]);}else{replaceAssistantBubble(tempId,'Failed to add files: '+(r.error||'unknown'));}}else{replaceAssistantBubble(tempId,'No files could be uploaded.');}}catch(err){replaceAssistantBubble(tempId,'Upload failed: '+(err&&err.message?err.message:'unknown'));}return;}")
        html.AppendLine("const f=files[0];if(!__supportsFiles){addTempAssistantBubble('File uploads are not supported for the current model.');return;}const tempId=addTempAssistantBubble(`Uploading <b>${f.name.replaceAll('&','&amp;')}</b> (${(f.size/1024).toFixed(1)} KB)…`);try{const fr=new FileReader();const dataUrl=await new Promise((res,rej)=>{fr.onerror=()=>rej(new Error('read error'));fr.onload=()=>res(fr.result);fr.readAsDataURL(f);});const r=await api('inky_upload',{Name:f.name,DataUrl:String(dataUrl||'')});if(!r.ok){replaceAssistantBubble(tempId,'Upload failed: '+(r.error||'unknown'));return;}if(r.supported===false){replaceAssistantBubble(tempId,'File uploads are not supported for this model.');return;}__pendingFilePath=r.path||'';replaceAssistantBubble(tempId,`Added file: <b>${(r.name||f.name).replaceAll('&','&amp;')}</b>`);}catch(err){replaceAssistantBubble(tempId,'Upload failed: '+(err&&err.message?err.message:'unknown'));}} ,false);})();")

        ' events
        html.AppendLine("modelSel.addEventListener('change',async()=>{if(__currentJobId)return;const opt=modelSel.options[modelSel.selectedIndex];if(!opt||opt.disabled||!opt.value){const fe=[...modelSel.options].find(o=>!o.disabled&&o.value);if(fe)fe.selected=true;}const r=await api('inky_setmodel',{Key:opt.value});updateModelTooltip();adjustModelSel();if(!r.ok){alert(r.error||'Failed to set model');return;}if(typeof r.supportsFiles==='boolean')__supportsFiles=r.supportsFiles;if(typeof r.supportsTooling==='boolean'){__modelSupportsTooling=!!r.supportsTooling;}if(typeof r.toolingEnabled==='boolean'){__toolingEnabled=!!r.toolingEnabled;}toolingChk.checked=__toolingEnabled;syncAdvancedToolsUi({advancedToolsEnabled:typeof r.advancedToolsEnabled==='boolean'?r.advancedToolsEnabled:false,agentWorkspace:Object.prototype.hasOwnProperty.call(r,'agentWorkspace')?r.agentWorkspace:null,agentFiles:Array.isArray(r.agentFiles)?r.agentFiles:[],agentModelAvailable:typeof r.agentModelAvailable==='boolean'?r.agentModelAvailable:__agentModelAvailable,agentModelActive:typeof r.agentModelActive==='boolean'?r.agentModelActive:false});});")
        html.AppendLine("clearBtn.addEventListener('click',async()=>{if(__currentJobId)return;const r=await api('inky_clear');if(r.ok){render([]);__pendingFilePath='';if(r.greeting)msgEl.placeholder=r.greeting;if(typeof r.toolingEnabled==='boolean'){__toolingEnabled=!!r.toolingEnabled;toolingChk.checked=__toolingEnabled;}if(typeof r.supportsTooling==='boolean'){__modelSupportsTooling=!!r.supportsTooling;updateToolingVisibility();}const st=await api('inky_getstate');if(st&&st.ok){syncAdvancedToolsUi({advancedToolsEnabled:st.advancedToolsEnabled===true,agentWorkspace:st.agentWorkspace,agentFiles:st.agentFiles||[],agentModelAvailable:st.agentModelAvailable===true,agentModelActive:st.agentModelActive===true});if(st.greeting)msgEl.placeholder=st.greeting;}else{updateAgentFilesDisplay([]);}applyCoupling();}else{alert(r.error||'Failed to clear');}adjustModelSel();});")
        html.AppendLine("copyBtn.addEventListener('click',async()=>{const r=await api('inky_copylast');if(!r.ok){alert(r.error||'Nothing to copy')}});")
        html.AppendLine("toWordBtn.addEventListener('click',async()=>{if(__currentJobId)return;const r=await api('inky_toword');if(!r.ok){alert(r.error||'Failed to create Word document')}});")
        html.AppendLine("playBtn.addEventListener('click',()=>{if(__currentJobId)return;const w=window.open('/inky/play','_blank');if(w){w.opener=null;}});")
        html.AppendLine("themeBtn.addEventListener('click',async()=>{if(__currentJobId)return;const target=!dark;setTheme(target);const r=await api('inky_toggletheme');if(!r.ok){setTheme(!target);alert(r.error||'Theme switch failed');return;}if(typeof r.darkMode==='boolean')setTheme(r.darkMode===true);adjustModelSel();});")
        html.AppendLine("msgEl.addEventListener('keydown',async e=>{if(e.ctrlKey&&e.key.toLowerCase()==='p'){e.preventDefault();if(__lastPrompt){insertPromptIntoMessage(__lastPrompt);}return;}if(e.key==='/'&&!e.ctrlKey&&!e.altKey&&!e.metaKey&&isPromptLibrarySlashTrigger()){e.preventDefault();if(__currentJobId)return;await openPromptLibrary();return;}if(e.key==='Enter'&&!e.shiftKey){e.preventDefault();send();return;}if(e.ctrlKey&&e.key.toLowerCase()==='l'){e.preventDefault();clearBtn.click();}});")
        html.AppendLine("sendBtn.addEventListener('click',send);")
        html.AppendLine("pureBtn.addEventListener('click',pureSend);")
        html.AppendLine("cancelBtn.addEventListener('click',async()=>{if(!__currentJobId)return;__jobCanceled=true;await api('inky_cancel',{Job:__currentJobId});});")
        html.AppendLine("chatEl.addEventListener('click',async e=>{const a=e.target&&e.target.closest&&e.target.closest('a[href]');if(!a)return;const href=String(a.getAttribute('href')||'').trim();if(!href)return;if(/^file:\/\//i.test(href)||/^[A-Za-z]:[\\/]/.test(href)){e.preventDefault();const r=await api('inky_openpath',{Path:href});if(!r||!r.ok)alert((r&&r.error)||'Could not open file link');return;}if(a.target!=='_blank'){a.target='_blank';a.rel='noopener noreferrer';}});")
        html.AppendLine("async function switchChat(n){if(__currentJobId)return;const r=await api('inky_switch',{Chat:String(n)});if(!r.ok){alert(r.error||'Switch failed');return;}setActiveChatBtn(r.activeChat||n);render(r.history||[]);if(r.greeting){msgEl.placeholder=r.greeting;}if(r.models&&r.models.length){modelSel.innerHTML='';for(const m of r.models){const o=document.createElement('option');o.value=m.key||'';o.textContent=m.label||'';o.disabled=!!m.disabled;o.title=o.textContent;if(m.selected&&!o.disabled)o.selected=true;modelSel.appendChild(o);}if(!modelSel.value){const fe=[...modelSel.options].find(o=>!o.disabled&&o.value);if(fe)fe.selected=true;}}if(typeof r.supportsFiles==='boolean')__supportsFiles=r.supportsFiles;if(typeof r.toolingEnabled==='boolean'){__toolingEnabled=!!r.toolingEnabled;toolingChk.checked=__toolingEnabled;}if(typeof r.supportsTooling==='boolean'){__modelSupportsTooling=!!r.supportsTooling;}syncAdvancedToolsUi({advancedToolsEnabled:r.advancedToolsEnabled===true,agentWorkspace:r.agentWorkspace,agentFiles:r.agentFiles||[],agentModelAvailable:r.agentModelAvailable===true,agentModelActive:r.agentModelActive===true});updateModelTooltip();adjustModelSel();}")
        html.AppendLine("chat1Btn.addEventListener('click',()=>switchChat(1));")
        html.AppendLine("chat2Btn.addEventListener('click',()=>switchChat(2));")

        ' Tooling UI visibility + coupling logic
        html.AppendLine("function updateToolingVisibility(){const show=__modelSupportsTooling===true;toolsBtn.style.display=show?'inline-block':'none';toolingSlot.style.display=show?'flex':'none';toolLogBtn.style.display=show?'flex':'none';if(!show){__toolingEnabled=false;toolingChk.checked=false;__advancedToolsEnabled=false;advancedToolsChk.checked=false;__agentModelActive=false;updateAgentWorkspaceDisplay(null);updateAgentFilesDisplay([]);}updateAgentModelBtn();applyCoupling();}")
        html.AppendLine("function applyCoupling(){advancedToolsChk.disabled=!__toolingEnabled;}")
        html.AppendLine("toolsBtn.addEventListener('click',async()=>{if(__currentJobId)return;const r=await api('inky_selecttools',{IncludeInteractiveM365Tools:true});if(!r.ok){alert(r.error||'Failed to select tools');return;}if(typeof r.toolingEnabled==='boolean'){__toolingEnabled=!!r.toolingEnabled;toolingChk.checked=__toolingEnabled;}if(typeof r.advancedToolsEnabled==='boolean'){__advancedToolsEnabled=!!r.advancedToolsEnabled;advancedToolsChk.checked=__advancedToolsEnabled;}applyCoupling();});")
        html.AppendLine("toolingChk.addEventListener('change',async()=>{if(toolingChk.checked){const r=await api('inky_settooling',{Enabled:true});if(!r.ok){toolingChk.checked=false;if(r.openSources){await api('inky_selecttools',{IncludeInteractiveM365Tools:true});}else{alert(r.error||'Failed to toggle tooling');}applyCoupling();return;}__toolingEnabled=!!r.enabled;if(typeof r.advancedToolsEnabled==='boolean'){__advancedToolsEnabled=!!r.advancedToolsEnabled;advancedToolsChk.checked=__advancedToolsEnabled;}syncAdvancedToolsUi({advancedToolsEnabled:__advancedToolsEnabled,agentWorkspace:r.agentWorkspace,agentFiles:r.agentFiles||[],agentModelAvailable:__agentModelAvailable,agentModelActive:__agentModelActive});}else{const r=await api('inky_settooling',{Enabled:false});if(!r.ok){toolingChk.checked=true;alert(r.error||'Failed to toggle tooling');applyCoupling();return;}__toolingEnabled=!!r.enabled;syncAdvancedToolsUi({advancedToolsEnabled:__advancedToolsEnabled,agentWorkspace:r.agentWorkspace,agentFiles:r.agentFiles||[],agentModelAvailable:__agentModelAvailable,agentModelActive:__agentModelActive});}});")
        html.AppendLine("advancedToolsChk.addEventListener('change',async()=>{const desired=advancedToolsChk.checked;const r=await api('inky_setagent',{Enabled:desired});if(!r||!r.ok){advancedToolsChk.checked=!desired;alert(r&&r.error||'Failed to toggle advanced tools');applyCoupling();return;}syncAdvancedToolsUi({advancedToolsEnabled:r.advancedToolsEnabled===true,agentWorkspace:r.agentWorkspace,agentFiles:r.files||[],agentModelAvailable:__agentModelAvailable,agentModelActive:__agentModelActive});});")

        ' Tooling log button (toggle)
        html.AppendLine("toolLogBtn.addEventListener('click',async()=>{__toolLogEnabled=!__toolLogEnabled;toolLogBtn.classList.toggle('active',__toolLogEnabled);const r=await api('inky_settoolinglog',{Enabled:__toolLogEnabled});if(!r.ok){__toolLogEnabled=!__toolLogEnabled;toolLogBtn.classList.toggle('active',__toolLogEnabled);}});")

        ' Memory toggle + editor
        html.AppendLine("memoryChk.addEventListener('change',async()=>{const r=await api('inky_setmemory',{Enabled:memoryChk.checked});if(!r.ok){memoryChk.checked=!memoryChk.checked;alert(r.error||'Failed to toggle memory');}else{__memoryEnabled=!!r.enabled;}memoryEditLnk.style.display=__memoryEnabled?'inline':'none';});")
        html.AppendLine("memoryEditLnk.addEventListener('click',async(e)=>{e.preventDefault();const r=await api('inky_editmemory');if(!r.ok)alert(r.error||'Could not open memory editor');});")

        ' Agent files display
        html.AppendLine("function updateAgentFilesDisplay(files){if(!agentFilesEl)return;if(!files||files.length===0||!__advancedToolsEnabled){agentFilesEl.style.display='none';agentFilesEl.innerHTML='';return;}agentFilesEl.style.display='block';let h='📎 Agent files: ';for(const f of files){const size=Number(f&&f.size||0);const kb=(size/1024).toFixed(1);const name=String(f&&f.name||'');h+=`<span class=""file-tag""><span class=""file-tag-name"" data-open-file=""${esc(name)}"" title=""Open this file"" style=""cursor:pointer"">${esc(name)} (${kb} KB)</span><button type=""button"" class=""file-tag-download"" data-download-file=""${esc(name)}"" title=""Save a copy to the Desktop Inky folder"">⤓</button><button type=""button"" class=""file-tag-remove"" data-remove-file=""${esc(name)}"" title=""Remove this file"">×</button></span> `;}agentFilesEl.innerHTML=h;}")
        html.AppendLine("agentFilesEl.addEventListener('click',async e=>{const openEl=e.target&&e.target.closest&&e.target.closest('[data-open-file]');if(openEl){e.preventDefault();const fileName=openEl.getAttribute('data-open-file')||'';if(!fileName)return;const r=await api('inky_agentopenfile',{Name:fileName});if(!r||!r.ok)alert(r&&r.error||'Failed to open file');return;}const dl=e.target&&e.target.closest&&e.target.closest('[data-download-file]');if(dl){e.preventDefault();const fileName=dl.getAttribute('data-download-file')||'';if(!fileName)return;dl.disabled=true;try{const r=await api('inky_agentdownloadfile',{Name:fileName});if(!r||!r.ok){alert(r&&r.error||'Failed to download file');return;}addTempAssistantBubble('Saved a copy to: '+esc(r.savedPath||''));}finally{dl.disabled=false;}return;}const btn=e.target&&e.target.closest&&e.target.closest('[data-remove-file]');if(!btn)return;e.preventDefault();if(__currentJobId)return;const fileName=btn.getAttribute('data-remove-file')||'';if(!fileName)return;btn.disabled=true;const r=await api('inky_agentremovefile',{Name:fileName});if(!r||!r.ok){btn.disabled=false;alert(r&&r.error||'Failed to remove file');return;}updateAgentFilesDisplay(r.files||[]);});")
        html.AppendLine("function updateAgentModelBtn(){if(!agentModelBtn)return;agentModelBtn.style.display=__agentModelAvailable?'flex':'none';agentModelBtn.classList.toggle('active',__agentModelActive);}")
        html.AppendLine("agentModelBtn.addEventListener('click',async()=>{if(__currentJobId)return;agentModelBtn.disabled=true;try{const r=await api('inky_toggleagentmodel');if(!r.ok){alert(r.error||'Failed to toggle agent model');return;}if(r.models&&r.models.length){modelSel.innerHTML='';for(const m of r.models){const o=document.createElement('option');o.value=m.key||'';o.textContent=m.label||'';o.disabled=!!m.disabled;o.title=o.textContent;if(m.selected&&!o.disabled)o.selected=true;modelSel.appendChild(o);}if(!modelSel.value){const fe=[...modelSel.options].find(o=>!o.disabled&&o.value);if(fe)fe.selected=true;}}updateModelTooltip();if(typeof r.supportsFiles==='boolean')__supportsFiles=r.supportsFiles;if(typeof r.supportsTooling==='boolean'){__modelSupportsTooling=!!r.supportsTooling;}if(typeof r.toolingEnabled==='boolean'){__toolingEnabled=!!r.toolingEnabled;toolingChk.checked=__toolingEnabled;}syncAdvancedToolsUi({advancedToolsEnabled:r.advancedToolsEnabled===true,agentWorkspace:Object.prototype.hasOwnProperty.call(r,'agentWorkspace')?r.agentWorkspace:null,agentFiles:r.agentFiles||[],agentModelAvailable:__agentModelAvailable,agentModelActive:r.active===true});adjustModelSel();}finally{agentModelBtn.disabled=false;}});")
        html.AppendLine("function esc(s){return String(s||'').replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('""','&quot;');}")
        html.AppendLine("function updateAgentWorkspaceDisplay(ws){if(!agentWorkspaceEl)return;if(!__advancedToolsEnabled){agentWorkspaceEl.style.display='none';agentWorkspaceEl.innerHTML='';return;}agentWorkspaceEl.style.display='block';if(!ws||!ws.connected){agentWorkspaceEl.innerHTML='<span class=""ws-title"">📁 Workspace:</span> Not connected <button id=""wsConnectBtn"">Connect folder</button>';bindWorkspaceButtons();return;}agentWorkspaceEl.innerHTML='<span class=""ws-title"">📁 Workspace:</span> '+esc(ws.name||'Workspace')+' <span class=""ws-path"">'+esc(ws.rootPath||'')+'</span> <button id=""wsOpenBtn"">Open</button><button id=""wsChangeBtn"">Change</button><button id=""wsPermBtn"">Permissions</button><button id=""wsRevokeBtn"">Revoke</button>';bindWorkspaceButtons();}")
        html.AppendLine("async function refreshWorkspace(){const r=await api('inky_agentworkspace_get');if(r&&r.ok)updateAgentWorkspaceDisplay(r.workspace);}")

        html.AppendLine("function closeWorkspaceModal(){workspaceModalBackdrop.classList.remove('show');workspaceModalBody.innerHTML='';workspaceModalOkBtn.textContent='Save';workspaceModalCancelBtn.style.display='';__workspaceDialogMode='';}")
        html.AppendLine("function openWorkspaceModal(title,bodyHtml,okText,mode,showCancel){workspaceModalTitle.textContent=title||'Agent Workspace';workspaceModalBody.innerHTML=bodyHtml||'';workspaceModalOkBtn.textContent=okText||'Save';workspaceModalCancelBtn.style.display=(showCancel===false)?'none':'';workspaceModalBackdrop.classList.add('show');__workspaceDialogMode=mode||'';}")
        html.AppendLine("function showWorkspaceModalMessage(title,message,isError){const html='<div class=""inky-form-row""><div class=""inky-message-box'+(isError===true?' error':'')+'"">'+esc(message||'Unexpected error.')+'</div></div>';openWorkspaceModal(title||'Agent Workspace',html,'OK','message',false);}")
        html.AppendLine("function showWorkspaceError(message,title){showWorkspaceModalMessage(title||'Workspace Error',message||'Unexpected error.',true);}")
        html.AppendLine("function getWorkspaceDialogValues(){return{persistUntilRevoked:!!document.getElementById('wsPersistUntilRevoked')?.checked,allowRead:!!document.getElementById('wsAllowRead')?.checked,allowWrite:!!document.getElementById('wsAllowWrite')?.checked,allowMoveCopyRename:!!document.getElementById('wsAllowMoveCopyRename')?.checked,allowDelete:!!document.getElementById('wsAllowDelete')?.checked,saveDroppedFilesToWorkspace:!!document.getElementById('wsSaveDroppedFiles')?.checked,includeHiddenSystem:!!document.getElementById('wsIncludeHiddenSystem')?.checked};}")
        html.AppendLine("function buildWorkspaceDialogHtml(ws,mode){const connected=!!(ws&&ws.connected);const root=connected?('<div class=""inky-form-row""><label class=""hd"">Current workspace</label><div class=""inky-pathbox"">'+esc(ws.rootPath||'')+'</div></div>'):'';const persist=!!(ws&&ws.persistUntilRevoked);const allowRead=ws&&typeof ws.allowRead==='boolean'?!!ws.allowRead:true;const allowWrite=ws&&typeof ws.allowWrite==='boolean'?!!ws.allowWrite:true;const allowMoveCopyRename=ws&&typeof ws.allowMoveCopyRename==='boolean'?!!ws.allowMoveCopyRename:true;const allowDelete=ws&&typeof ws.allowDelete==='boolean'?!!ws.allowDelete:false;const saveDrops=ws&&typeof ws.saveDroppedFilesToWorkspace==='boolean'?!!ws.saveDroppedFilesToWorkspace:false;const includeHidden=ws&&typeof ws.includeHiddenSystem==='boolean'?!!ws.includeHiddenSystem:false;let intro='';if(mode==='connect'){intro='<div class=""inky-form-row inky-help"">Select how the Local Agent should use the workspace after you choose the folder.</div>';}else if(mode==='revoke'){return '<div class=""inky-form-row""><div class=""inky-help"">Revoke access to the current Local Agent workspace?</div></div>'+(connected?('<div class=""inky-form-row""><div class=""inky-pathbox"">'+esc(ws.rootPath||'')+'</div></div>'):'');}return intro+root+'<div class=""inky-form-row""><label class=""hd"">Access duration</label><div class=""inky-radio-list""><label><input type=""radio"" name=""wsDuration"" id=""wsPersistSessionOnly"" '+(!persist?'checked':'')+'> <span>This session only</span></label><label><input type=""radio"" name=""wsDuration"" id=""wsPersistUntilRevoked"" '+(persist?'checked':'')+'> <span>Until revoked</span></label></div><div class=""inky-help"">Session access disappears when Outlook is closed. Persistent access remains until the user revokes it.</div></div><div class=""inky-form-row""><label class=""hd"">Allowed operations</label><div class=""inky-check-list""><label><input type=""checkbox"" id=""wsAllowRead"" '+(allowRead?'checked':'')+'> <span>Read and list workspace files</span></label><label><input type=""checkbox"" id=""wsAllowWrite"" '+(allowWrite?'checked':'')+'> <span>Create and overwrite workspace files</span></label><label><input type=""checkbox"" id=""wsAllowMoveCopyRename"" '+(allowMoveCopyRename?'checked':'')+'> <span>Copy, move, and rename files</span></label><label><input type=""checkbox"" id=""wsAllowDelete"" '+(allowDelete?'checked':'')+'> <span class=""inky-danger"">Delete files and remove folders</span></label></div></div><div class=""inky-form-row""><label class=""hd"">Additional options</label><div class=""inky-check-list""><label><input type=""checkbox"" id=""wsSaveDroppedFiles"" '+(saveDrops?'checked':'')+'> <span>Also save drag-and-drop uploads into the workspace</span></label><label><input type=""checkbox"" id=""wsIncludeHiddenSystem"" '+(includeHidden?'checked':'')+'> <span>Include hidden/system files</span></label></div></div>'; }")
        html.AppendLine("async function connectWorkspace(){const st=await api('inky_agentworkspace_get');if(!st||!st.ok){showWorkspaceError(st&&st.error||'Could not load workspace state','Connect Agent Workspace');return;}openWorkspaceModal('Connect Agent Workspace',buildWorkspaceDialogHtml(st.workspace||{},'connect'),'Choose folder…','connect');}")
        html.AppendLine("async function changeWorkspaceDirect(){const st=await api('inky_agentworkspace_get');if(!st||!st.ok){showWorkspaceError(st&&st.error||'Could not load workspace state','Change Agent Workspace');return;}const ws=st.workspace||{};const r=await api('inky_agentworkspace_select',{PersistUntilRevoked:!!ws.persistUntilRevoked});if(!r||!r.ok){if(r&&r.error==='No workspace folder selected.')return;showWorkspaceError(r&&r.error||'Failed to select workspace','Change Agent Workspace');return;}updateAgentWorkspaceDisplay(r.workspace);}")
        html.AppendLine("async function editWorkspacePermissions(){const st=await api('inky_agentworkspace_get');if(!st||!st.ok){showWorkspaceError(st&&st.error||'Could not load workspace permissions','Workspace Permissions');return;}openWorkspaceModal('Workspace Permissions',buildWorkspaceDialogHtml(st.workspace||{},'permissions'),'Save','permissions');}")
        html.AppendLine("async function confirmWorkspaceRevoke(){const st=await api('inky_agentworkspace_get');if(!st||!st.ok){showWorkspaceError(st&&st.error||'Could not load workspace state','Revoke Workspace Access');return;}openWorkspaceModal('Revoke Workspace Access',buildWorkspaceDialogHtml(st.workspace||{},'revoke'),'Revoke','revoke');}")
        html.AppendLine("async function saveWorkspaceDialog(){if(__workspaceDialogMode==='message'){closeWorkspaceModal();return;}if(__workspaceDialogMode==='revoke'){const r=await api('inky_agentworkspace_revoke');if(r&&r.ok){closeWorkspaceModal();updateAgentWorkspaceDisplay(r.workspace);}else{showWorkspaceError(r&&r.error||'Failed to revoke workspace','Revoke Workspace Access');}return;}const values=getWorkspaceDialogValues();if(!values.allowRead){showWorkspaceError('Workspace read/list permission must remain enabled.','Workspace Permissions');return;}if(__workspaceDialogMode==='connect'){const r=await api('inky_agentworkspace_select',{PersistUntilRevoked:values.persistUntilRevoked});if(!r||!r.ok){if(r&&r.error==='No workspace folder selected.')return;showWorkspaceError(r&&r.error||'Failed to select workspace','Connect Agent Workspace');return;}const p=await api('inky_agentworkspace_permissions',{PersistUntilRevoked:values.persistUntilRevoked,AllowRead:values.allowRead,AllowWrite:values.allowWrite,AllowMoveCopyRename:values.allowMoveCopyRename,AllowDelete:values.allowDelete,SaveDroppedFilesToWorkspace:values.saveDroppedFilesToWorkspace,IncludeHiddenSystem:values.includeHiddenSystem});if(!p||!p.ok){showWorkspaceError(p&&p.error||'Failed to save workspace permissions','Connect Agent Workspace');return;}closeWorkspaceModal();updateAgentWorkspaceDisplay(p.workspace);}else if(__workspaceDialogMode==='permissions'){const r=await api('inky_agentworkspace_permissions',{PersistUntilRevoked:values.persistUntilRevoked,AllowRead:values.allowRead,AllowWrite:values.allowWrite,AllowMoveCopyRename:values.allowMoveCopyRename,AllowDelete:values.allowDelete,SaveDroppedFilesToWorkspace:values.saveDroppedFilesToWorkspace,IncludeHiddenSystem:values.includeHiddenSystem});if(r&&r.ok){closeWorkspaceModal();updateAgentWorkspaceDisplay(r.workspace);}else{showWorkspaceError(r&&r.error||'Failed to update permissions','Workspace Permissions');}}}")
        html.AppendLine("function bindWorkspaceButtons(){const c=document.getElementById('wsConnectBtn');if(c)c.onclick=connectWorkspace;const ch=document.getElementById('wsChangeBtn');if(ch)ch.onclick=changeWorkspaceDirect;const o=document.getElementById('wsOpenBtn');if(o)o.onclick=async()=>{const r=await api('inky_agentworkspace_open');if(!r.ok)showWorkspaceError(r.error||'Could not open workspace','Agent Workspace');};const rv=document.getElementById('wsRevokeBtn');if(rv)rv.onclick=confirmWorkspaceRevoke;const p=document.getElementById('wsPermBtn');if(p)p.onclick=editWorkspacePermissions;}")
        html.AppendLine("workspaceModalCancelBtn.addEventListener('click',closeWorkspaceModal);")
        html.AppendLine("workspaceModalCloseBtn.addEventListener('click',closeWorkspaceModal);")
        html.AppendLine("workspaceModalOkBtn.addEventListener('click',saveWorkspaceDialog);")
        html.AppendLine("workspaceModalBackdrop.addEventListener('click',e=>{if(e.target===workspaceModalBackdrop)closeWorkspaceModal();});")

        html.AppendLine("document.addEventListener('keydown',e=>{if(e.key==='Escape'&&workspaceModalBackdrop.classList.contains('show')){e.preventDefault();closeWorkspaceModal();}});")
        html.AppendLine("boot();")
        html.AppendLine("</script>")
        html.AppendLine("</body></html>")

        Return html.ToString()
    End Function

    ''' <summary>
    ''' Wraps object into JSON response with ok=True plus optional fields.
    ''' </summary>
    Private Function JsonOk(o As Object) As System.String
        Return "CT:json" & vbLf & Newtonsoft.Json.JsonConvert.SerializeObject(o)
    End Function

    ''' <summary>
    ''' Wraps error message into JSON response with ok=False.
    ''' </summary>
    Private Function JsonErr(msg As System.String) As System.String
        Return JsonOk(New With {.ok = False, .error = msg})
    End Function

    Private Function TryGetJsonCommandForInkyServerLog(ByVal body As System.String) As System.String
        If System.String.IsNullOrWhiteSpace(body) Then Return ""

        Try
            Dim j As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(body)
            Return If(j("Command")?.ToString(), "")
        Catch ex As System.Exception
            AppendInkyServerLog("TryGetJsonCommandForInkyServerLog failed.", ex)
            Return ""
        End Try
    End Function

    ' ===== (D) EXTEND ProcessRequestInAddIn with the Inky commands =====

    Private Shared ReadOnly AlternateModelLock As New Object

    ''' <summary>
    ''' Central dispatcher for browser API commands under InkyApiRoute; handles chat operations,
    ''' job control, model selection, theme toggle, file upload, and falls back to legacy commands.
    ''' </summary>
    Private Async Function ProcessRequestInAddIn(
        body As System.String,
        rawUrl As System.String) As System.Threading.Tasks.Task(Of System.String)

        ' If this is a browser POST to our Inky API, j may be JSON; otherwise keep your existing flow
        If rawUrl IsNot Nothing AndAlso rawUrl.StartsWith(InkyApiRoute, System.StringComparison.OrdinalIgnoreCase) Then
            Try
                Dim j As Newtonsoft.Json.Linq.JObject = If(
                    Not System.String.IsNullOrWhiteSpace(body),
                    Newtonsoft.Json.Linq.JObject.Parse(body),
                    New Newtonsoft.Json.Linq.JObject())
                Dim cmd As System.String = j("Command")?.ToString()

                If INI_APIDebug Then
                    AppendInkyServerLog(
                        "ProcessRequestInAddIn begin. rawUrl=" &
                        If(rawUrl, "") &
                        "; cmd=" &
                        If(cmd, "") &
                        "; bodyLen=" &
                        If(body, "").Length.ToString(System.Globalization.CultureInfo.InvariantCulture))
                End If

                Select Case cmd

                                        ' ---- InkyPlay commands ----
                    Case "inkyplay_generate", "inkyplay_gethighscores", "inkyplay_savescore", "inkyplay_clearhighscores"
                        Dim playResult As String = Await ProcessInkyPlayCommand(cmd, j).ConfigureAwait(False)
                        If playResult IsNot Nothing Then Return playResult


                    Case "inky_setagent"
                        Try
                            If IsChatAgentBlocked() Then
                                Return JsonErr("Agent mode is not available while AutoPilot is running.")
                            End If

                            Dim enabled As Boolean = CBool(j("Enabled"))
                            Dim st = LoadInkyState()

                            If enabled AndAlso Not CurrentModelSupportsTooling(st) Then
                                Return JsonErr("The selected model does not support tool calling.")
                            End If

                            If enabled Then
                                Dim effectiveAgentTools =
                GetLocalChatEffectiveTools(
                    ResolveLocalChatSelectedMainToolNames(st, includeInteractiveM365Tools:=True),
                    ResolveLocalChatSelectedAdvancedToolNames(st, includeInteractiveM365Tools:=True),
                    True,
                    includeInteractiveM365Tools:=True)

                                If effectiveAgentTools.Count = 0 AndAlso Not IsChatAgentWorkspaceConnected() Then
                                    ' Opt-out defaults keep all tools on unless the user deselected them all.
                                    ' Do not auto-open the selection dialog here; report a plain error instead.
                                    Return JsonErr("Agent mode requires at least one tool or a connected workspace. All tools are on by default — use the Tools button to re-enable some (or connect a workspace), then try again.")
                                End If

                                st.ToolingEnabled = True
                                _chatToolingEnabled = True
                            Else
                                ChatAgentClearFiles()
                            End If

                            st.AgentModeEnabled = enabled
                            _selectedToolsForChat = GetLocalChatEffectiveSelection(st, includeInteractiveM365Tools:=True)

                            If enabled AndAlso _selectedToolsForChat.Count = 0 AndAlso Not IsChatAgentWorkspaceConnected() Then
                                st.AgentModeEnabled = False
                                Return JsonErr("Agent mode requires at least one tool or a connected workspace.")
                            End If

                            SaveInkyState(st)

                            Return JsonOk(New With {
                                    .ok = True,
                                    .advancedToolsEnabled = enabled,
                                    .agentWorkspace = GetAgentWorkspaceForBrowser(),
                                    .files = GetAgentFileListForBrowser()
                                })
                        Catch ex As Exception
                            Return JsonErr("Failed to toggle advanced tools: " & ex.Message)
                        End Try


                    Case "inky_editmemory"
                        Try
                            Dim ownerHandle As IntPtr = IntPtr.Zero
                            Try
                                ownerHandle = NativeMethods.GetForegroundWindow()
                            Catch
                            End Try

                            Await SwitchToUi(Sub()
                                                 SharedMethods.EditInkyMemoryFile(_context, ownerHandle)
                                             End Sub).ConfigureAwait(False)
                            Return JsonOk(New With {.ok = True})
                        Catch ex As Exception
                            Return JsonErr("Failed to open memory editor: " & ex.Message)
                        End Try

                    Case "inky_toggleagentmodel"
                        ' Toggle the AgentDefaultModel on/off with one click.
                        ' When toggling ON:  find AgentDefaultModel in INI, switch model selector to it,
                        '                    enable agent mode + tooling, return updated model list.
                        ' When toggling OFF: restore the model that was active before, disable agent model flag.
                        Try
                            Dim st = LoadInkyState()

                            If st.AgentModelActive Then
                                ' ── TOGGLE OFF: restore previous model ──
                                st.AgentModelActive = False
                                st.SelectedModelKey = st.PreAgentModelKey
                                st.UseSecondApi = st.PreAgentUseSecondApi
                                st.PreAgentModelKey = ""
                                st.PreAgentUseSecondApi = False

                                ' Also turn off agent mode (tools)
                                st.AgentModeEnabled = False
                                _chatAdvancedToolsEnabled = False
                                ChatAgentClearFiles()

                                ' Re-evaluate tooling/files for restored model
                                Dim supportsToolingOff = CurrentModelSupportsTooling(st)
                                If Not supportsToolingOff AndAlso st.ToolingEnabled Then
                                    st.ToolingEnabled = False
                                    _chatToolingEnabled = False
                                End If

                                Try
                                    st.SupportsFileUploads = ComputeSupportsFiles(st.UseSecondApi, st.SelectedModelKey)
                                Catch
                                    st.SupportsFileUploads = False
                                End Try

                                SaveInkyState(st)

                                Dim models = Await GetModelListForBrowserAsync(st)
                                Dim supTooling As Boolean = False
                                Dim effTooling As Boolean = SyncToolingState(st, supTooling)

                                Return JsonOk(New With {
                                    .ok = True,
                                    .active = False,
                                    .models = models,
                                    .supportsFiles = st.SupportsFileUploads,
                                    .supportsTooling = supTooling,
                                    .toolingEnabled = effTooling,
                                    .advancedToolsEnabled = False,
                                    .agentFiles = GetAgentFileListForBrowser(),
                                    .agentWorkspace = GetAgentWorkspaceForBrowser()
                                })
                            Else
                                ' ── TOGGLE ON: find and apply AgentDefaultModel ──
                                If IsChatAgentBlocked() Then
                                    Return JsonErr("Agent mode is not available while AutoPilot is running.")
                                End If

                                Dim agentDisplayKey As String = Nothing
                                Dim agentConfig As ModelConfig = FindAgentDefaultModel(agentDisplayKey)

                                If agentConfig Is Nothing OrElse String.IsNullOrWhiteSpace(agentDisplayKey) Then
                                    Return JsonErr("No model with 'AgentDefaultModel=True' found in the alternate model configuration.")
                                End If

                                ' Verify the agent model supports tooling
                                If Not ModelSupportsTooling(agentConfig) Then
                                    Return JsonErr("The AgentDefaultModel does not support tool calling. Check its APICall_ToolInstructions setting.")
                                End If

                                ' Save current model selection so we can restore on toggle-off
                                st.PreAgentModelKey = st.SelectedModelKey
                                st.PreAgentUseSecondApi = st.UseSecondApi
                                st.AgentModelActive = True

                                ' Switch to the agent model first, but do not mark agent mode as active yet.
                                ' The source-selection flow reads AgentModeEnabled while building the
                                ' effective tool list, so setting it too early causes the wrong UI/sequence.
                                st.UseSecondApi = True
                                st.SelectedModelKey = agentDisplayKey
                                st.AgentModeEnabled = False
                                _chatAdvancedToolsEnabled = False

                                ' Ensure sources are actually selected (prompt user if necessary)
                                Dim effectiveAgentTools =
                                            GetLocalChatEffectiveTools(
                                                ResolveLocalChatSelectedMainToolNames(st, includeInteractiveM365Tools:=True),
                                                ResolveLocalChatSelectedAdvancedToolNames(st, includeInteractiveM365Tools:=True),
                                                True,
                                                includeInteractiveM365Tools:=True)

                                If effectiveAgentTools.Count = 0 AndAlso Not IsChatAgentWorkspaceConnected() Then
                                    ' Opt-out defaults keep all tools on unless the user deselected them all.
                                    ' Abort the toggle and restore the previous model instead of auto-opening a dialog.
                                    st.AgentModelActive = False
                                    st.SelectedModelKey = st.PreAgentModelKey
                                    st.UseSecondApi = st.PreAgentUseSecondApi
                                    st.PreAgentModelKey = ""
                                    st.PreAgentUseSecondApi = False
                                    st.AgentModeEnabled = False
                                    _chatAdvancedToolsEnabled = False
                                    SaveInkyState(st)
                                    Return JsonErr($"The agent model requires at least one {ToolFriendlyName.ToLower()} or a connected workspace. All tools are on by default — use the Tools button to re-enable some, then try again.")
                                End If

                                ' Only now is agent mode actually active.
                                st.AgentModeEnabled = True
                                _chatAdvancedToolsEnabled = True

                                ' Force-enable tooling since we verified tools exist above
                                st.ToolingEnabled = True
                                _chatToolingEnabled = True
                                _selectedToolsForChat = GetLocalChatEffectiveSelection(st, includeInteractiveM365Tools:=True)

                                Try
                                    st.SupportsFileUploads = ComputeSupportsFiles(st.UseSecondApi, st.SelectedModelKey)
                                Catch
                                    st.SupportsFileUploads = False
                                End Try

                                Try
                                    My.Settings.Inky_UseSecondApiSelected = st.UseSecondApi
                                    My.Settings.Inky_SelectedModelKey = st.SelectedModelKey
                                    My.Settings.Save()
                                Catch
                                End Try

                                SaveInkyState(st)

                                Dim models = Await GetModelListForBrowserAsync(st)

                                ' Do NOT call SyncToolingState here — it wipes _selectedToolsForChat
                                ' and may fail to reload it, undoing the agent model activation.
                                ' We already know the agent model supports tooling (verified above)
                                ' and tools are selected (EnsureToolsSelected passed).
                                Dim supTooling As Boolean = True
                                Dim effTooling As Boolean = True

                                Return JsonOk(New With {
                                    .ok = True,
                                    .active = True,
                                    .agentModelLabel = agentDisplayKey,
                                    .models = models,
                                    .supportsFiles = st.SupportsFileUploads,
                                    .supportsTooling = supTooling,
                                    .toolingEnabled = effTooling,
                                    .advancedToolsEnabled = True,
                                    .agentFiles = GetAgentFileListForBrowser(),
                                    .agentWorkspace = GetAgentWorkspaceForBrowser()
                                })
                            End If
                        Catch ex As Exception
                            Return JsonErr("Agent model toggle failed: " & ex.Message)
                        End Try

                    Case "inky_agentaddfiles"
                        ' Called after drag-and-drop when advanced tools mode is on
                        Try
                            If Not GetEffectiveAdvancedToolsEnabled() Then
                                Return JsonErr("Advanced tools mode is not enabled.")
                            End If
                            If IsChatAgentBlocked() Then
                                Return JsonErr("Agent mode is not available while AutoPilot is running.")
                            End If

                            Dim paths = j("Paths")
                            If paths Is Nothing Then Return JsonErr("No file paths provided.")

                            Dim addedFiles As New List(Of String)()
                            For Each p In paths
                                Dim filePath = p?.ToString()
                                If String.IsNullOrWhiteSpace(filePath) Then Continue For
                                Dim att = ChatAgentAddFile(filePath)
                                If att IsNot Nothing Then addedFiles.Add(att.OriginalFileName)
                            Next

                            Return JsonOk(New With {
                                .ok = True,
                                .addedCount = addedFiles.Count,
                                .files = GetAgentFileListForBrowser()
                            })
                        Catch ex As Exception
                            Return JsonErr("Failed to add files: " & ex.Message)
                        End Try

                    Case "inky_agentdownloadfile"
                        Try
                            Dim fileName As String = j("Name")?.ToString()
                            If String.IsNullOrWhiteSpace(fileName) Then
                                Return JsonErr("No file name provided.")
                            End If

                            Dim savedPath As String = ChatAgentDownloadFileToDesktop(fileName)
                            If String.IsNullOrWhiteSpace(savedPath) Then
                                Return JsonErr("The file could not be found in the current agent session.")
                            End If

                            Return JsonOk(New With {
                                .ok = True,
                                .savedPath = savedPath
                            })
                        Catch ex As Exception
                            Return JsonErr("Failed to download file: " & ex.Message)
                        End Try

                    Case "inky_agentopenfile"
                        Try
                            Dim fileName As String = j("Name")?.ToString()
                            If String.IsNullOrWhiteSpace(fileName) Then
                                Return JsonErr("No file name provided.")
                            End If

                            If Not ChatAgentOpenFileInDefaultApp(fileName) Then
                                Return JsonErr("The file could not be found in the current agent session.")
                            End If

                            Return JsonOk(New With {.ok = True})
                        Catch ex As Exception
                            Return JsonErr("Failed to open file: " & ex.Message)
                        End Try

                    Case "inky_agentremovefile"
                        Try
                            If _chatAgentActive Then
                                Return JsonErr("Cannot remove files while agent processing is running.")
                            End If

                            Dim fileName As String = j("Name")?.ToString()
                            If String.IsNullOrWhiteSpace(fileName) Then
                                Return JsonErr("No file name provided.")
                            End If

                            If Not ChatAgentRemoveFile(fileName) Then
                                Return JsonErr("The file could not be found in the current agent session.")
                            End If

                            Return JsonOk(New With {
                                .ok = True,
                                .files = GetAgentFileListForBrowser()
                            })
                        Catch ex As Exception
                            Return JsonErr("Failed to remove file: " & ex.Message)
                        End Try

                    Case "inky_agentclearfiles"
                        Try
                            ChatAgentClearFiles()
                            Return JsonOk(New With {
                                .ok = True,
                                .files = GetAgentFileListForBrowser()
                            })
                        Catch ex As Exception
                            Return JsonErr("Failed to clear files: " & ex.Message)
                        End Try

                    Case "inky_agentworkspace_get"
                        Try
                            Return JsonOk(New With {
                                .ok = True,
                                .workspace = GetAgentWorkspaceForBrowser()
                            })
                        Catch ex As Exception
                            Return JsonErr("Failed to get workspace state: " & ex.Message)
                        End Try

                    Case "inky_agentworkspace_select"
                        Try
                            If Not GetEffectiveToolingEnabled() Then Return JsonErr("Agents are not enabled.")
                            If IsChatAgentBlocked() Then Return JsonErr("Agent mode is not available while AutoPilot is running.")

                            Dim persistUntilRevoked As Boolean = CBool(If(j("PersistUntilRevoked"), False))
                            Dim selectedPath As String = Nothing

                            Await SwitchToUi(Sub()
                                                 Using dlg As New DragDropForm(DragDropMode.DirectoryOnly)
                                                     dlg.Text = "Select Agent Workspace Folder"
                                                     dlg.SetInstructionText("Drag and drop the workspace folder here, or click Browse. The Local Agent will only be allowed to access this folder and its subfolders.")
                                                     If dlg.ShowDialog() = DialogResult.OK Then
                                                         selectedPath = dlg.SelectedFilePath
                                                     End If
                                                 End Using
                                             End Sub).ConfigureAwait(False)

                            If String.IsNullOrWhiteSpace(selectedPath) Then
                                Return JsonErr("No workspace folder selected.")
                            End If

                            If Not ChatAgentWorkspaceSetRoot(selectedPath, persistUntilRevoked) Then
                                Return JsonErr("The selected workspace folder is not valid.")
                            End If

                            Return JsonOk(New With {
                                .ok = True,
                                .workspace = GetAgentWorkspaceForBrowser()
                            })
                        Catch ex As Exception
                            Return JsonErr("Failed to select workspace: " & ex.Message)
                        End Try

                    Case "inky_agentworkspace_permissions"
                        Try
                            ChatAgentWorkspaceSetPermissions(
                                CBool(If(j("PersistUntilRevoked"), False)),
                                CBool(If(j("AllowRead"), True)),
                                CBool(If(j("AllowWrite"), True)),
                                CBool(If(j("AllowMoveCopyRename"), True)),
                                CBool(If(j("AllowDelete"), False)),
                                CBool(If(j("SaveDroppedFilesToWorkspace"), False)),
                                CBool(If(j("IncludeHiddenSystem"), False)))

                            Return JsonOk(New With {
                                .ok = True,
                                .workspace = GetAgentWorkspaceForBrowser()
                            })
                        Catch ex As Exception
                            Return JsonErr("Failed to update workspace permissions: " & ex.Message)
                        End Try

                    Case "inky_agentworkspace_revoke"
                        Try
                            ChatAgentWorkspaceRevoke()
                            Return JsonOk(New With {
                                .ok = True,
                                .workspace = GetAgentWorkspaceForBrowser()
                            })
                        Catch ex As Exception
                            Return JsonErr("Failed to revoke workspace: " & ex.Message)
                        End Try

                    Case "inky_agentworkspace_open"
                        Try
                            Dim ws = GetAgentWorkspaceForBrowser()
                            Dim rootPath = TryCast(ws.GetType().GetProperty("rootPath")?.GetValue(ws, Nothing), String)
                            If String.IsNullOrWhiteSpace(rootPath) OrElse Not Directory.Exists(rootPath) Then
                                Return JsonErr("No workspace is connected.")
                            End If

                            Process.Start("explorer.exe", rootPath)

                            Return JsonOk(New With {.ok = True})
                        Catch ex As Exception
                            Return JsonErr("Failed to open workspace: " & ex.Message)
                        End Try

                    Case "inky_openpath"
                        Try
                            Dim rawPath As String = j("Path")?.ToString()
                            If String.IsNullOrWhiteSpace(rawPath) Then
                                Return JsonErr("Missing path.")
                            End If

                            Dim fullPath As String = rawPath.Trim()

                            If fullPath.StartsWith("file://", StringComparison.OrdinalIgnoreCase) Then
                                fullPath = New Uri(fullPath).LocalPath
                            End If

                            fullPath = Path.GetFullPath(fullPath)

                            If File.Exists(fullPath) Then
                                Process.Start("explorer.exe", "/select,""" & fullPath & """")
                            ElseIf Directory.Exists(fullPath) Then
                                Process.Start("explorer.exe", """" & fullPath & """")
                            Else
                                Return JsonErr("The referenced path does not exist.")
                            End If

                            Return JsonOk(New With {.ok = True})
                        Catch ex As Exception
                            Return JsonErr("Failed to open path: " & ex.Message)
                        End Try

                    Case "inky_settoolinglog"
                        Try
                            Dim enabled As Boolean = CBool(j("Enabled"))
                            INI_ToolingLogWindow = enabled
                            Return JsonOk(New With {.ok = True, .enabled = enabled})
                        Catch ex As Exception
                            Return JsonErr("Failed to toggle tooling log: " & ex.Message)
                        End Try

                    Case "inky_setmemory"
                        Try
                            Dim enabled As Boolean = CBool(j("Enabled"))
                            My.Settings.Inky_InkyMemory = enabled
                            My.Settings.Save()
                            Return JsonOk(New With {.ok = True, .enabled = enabled})
                        Catch ex As Exception
                            Return JsonErr("Failed to toggle memory: " & ex.Message)
                        End Try

                    Case "inky_selecttools"

                        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "LocalChat_SelectTools invoked")

                        Try
                            Dim st = LoadInkyState()
                            If Not CurrentModelSupportsTooling(st) Then
                                Return JsonErr($"The selected model does not support {ToolFriendlyName.ToLower()}.")
                            End If

                            Dim includeInteractiveM365Tools As Boolean = True

                            If j("IncludeInteractiveM365Tools") IsNot Nothing Then
                                Try
                                    includeInteractiveM365Tools = includeInteractiveM365Tools OrElse CBool(j("IncludeInteractiveM365Tools"))
                                Catch
                                End Try
                            End If

                            Dim selectedMainToolNames As List(Of String) = Nothing
                            Dim selectedAdvancedToolNames As List(Of String) = Nothing

                            Await SwitchToUi(
                                Sub()
                                    selectedMainToolNames =
                                        ShowLocalChatToolSelectionDialog(
                                            ResolveLocalChatSelectedMainToolNames(st, includeInteractiveM365Tools),
                                            ResolveLocalChatSelectedAdvancedToolNames(st, includeInteractiveM365Tools),
                                            selectedAdvancedToolNames,
                                            includeInteractiveM365Tools)
                                End Sub).ConfigureAwait(False)

                            If selectedMainToolNames Is Nothing Then
                                Return JsonOk(New With {.ok = True, .cancelled = True})
                            End If

                            st.DeselectedMainToolNames =
                                ComputeLocalChatDeselectedMainToolNames(selectedMainToolNames, includeInteractiveM365Tools)
                            st.DeselectedAdvancedToolNames =
                                ComputeLocalChatDeselectedAdvancedToolNames(If(selectedAdvancedToolNames, New List(Of String)()), includeInteractiveM365Tools)
                            _selectedToolsForChat = GetLocalChatEffectiveSelection(st, includeInteractiveM365Tools)

                            If HasLocalChatAnyCallableTools(st, includeInteractiveM365Tools) Then
                                st.ToolingEnabled = True
                                _chatToolingEnabled = True
                            End If

                            SaveInkyState(st)

                            Return JsonOk(New With {
                                .ok = True,
                                .tools = GetToolListForBrowser(includeInteractiveM365Tools:=True),
                                .count = ResolveLocalChatSelectedMainToolNames(st, includeInteractiveM365Tools).Count,
                                .advancedCount = ResolveLocalChatSelectedAdvancedToolNames(st, includeInteractiveM365Tools).Count,
                                .toolingEnabled = st.ToolingEnabled,
                                .advancedToolsEnabled = st.AgentModeEnabled
                            })
                        Catch ex As Exception
                            Return JsonErr("Tool selection failed: " & ex.Message)
                        End Try

                    Case "inky_settooling"

                        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "LocalChat_SetTooling invoked")

                        Try
                            Dim enabled As Boolean = CBool(j("Enabled"))
                            Dim st = LoadInkyState()

                            If enabled AndAlso Not CurrentModelSupportsTooling(st) Then
                                Return JsonErr($"The selected model does not support {ToolFriendlyName.ToLower()}.")
                            End If

                            Dim includeInteractiveM365Tools As Boolean = True

                            If enabled AndAlso Not EnsureToolsSelected(st, includeInteractiveM365Tools:=includeInteractiveM365Tools) Then
                                ' Opt-out defaults keep all tools on unless the user deselected them all.
                                ' Report a plain error rather than signalling the UI to auto-open the selection dialog.
                                Return JsonErr("Tooling requires at least one tool or a connected workspace. All tools are on by default — use the Tools button to re-enable some, then try again.")
                            End If

                            st.ToolingEnabled = enabled
                            _chatToolingEnabled = enabled

                            If Not enabled Then
                                ChatAgentClearFiles()
                            End If

                            SaveInkyState(st)

                            Return JsonOk(New With {
                                .ok = True,
                                .enabled = enabled,
                                .advancedToolsEnabled = st.AgentModeEnabled,
                                .agentWorkspace = GetAgentWorkspaceForBrowser(),
                                .agentFiles = GetAgentFileListForBrowser()
                            })
                        Catch ex As Exception
                            Return JsonErr("Failed to toggle tooling: " & ex.Message)
                        End Try

                    Case "inky_gettoolingstate"
                        Try
                            Dim st = LoadInkyState()

                            Return JsonOk(New With {
                                .ok = True,
                                .enabled = st.ToolingEnabled,
                                .supportsTooling = CurrentModelSupportsTooling(st),
                                .tools = GetToolListForBrowser(includeInteractiveM365Tools:=True),
                                .selectedCount = ResolveLocalChatSelectedMainToolNames(st, includeInteractiveM365Tools:=True).Count,
                                .advancedToolsEnabled = st.AgentModeEnabled,
                                .advancedSelectedCount = ResolveLocalChatSelectedAdvancedToolNames(st, includeInteractiveM365Tools:=True).Count
                            })
                        Catch ex As Exception
                            Return JsonErr("Failed to get tooling state: " & ex.Message)
                        End Try

                    Case "inky_getstate"
                        Try
                            Dim debugEnabled As System.Boolean = INI_APIDebug

                            If debugEnabled Then
                                AppendInkyServerLog("CMD inky_getstate begin.")
                            End If

                            Dim st As InkyState = LoadInkyState()
                            _chatAdvancedToolsEnabled = st.AgentModeEnabled

                            If debugEnabled Then
                                AppendInkyServerLog(
                                    "CMD inky_getstate state loaded. HistoryCount=" &
                                    If(st?.History?.Count, 0).ToString(System.Globalization.CultureInfo.InvariantCulture) &
                                    "; UseSecondApi=" &
                                    st.UseSecondApi.ToString() &
                                    "; SelectedModelKey=" &
                                    If(st.SelectedModelKey, "") &
                                    "; AgentModeEnabled=" &
                                    st.AgentModeEnabled.ToString())
                            End If

                            Try
                                st.DarkMode = My.Settings.Inky_DarkMode
                            Catch exDark As System.Exception
                                If debugEnabled Then
                                    AppendInkyServerLog("CMD inky_getstate failed while reading My.Settings.Inky_DarkMode.", exDark)
                                End If
                            End Try

                            Try
                                st.SupportsFileUploads = ComputeSupportsFiles(st.UseSecondApi, st.SelectedModelKey)
                                SaveInkyState(st)
                            Catch exSupports As System.Exception
                                st.SupportsFileUploads = False
                                If debugEnabled Then
                                    AppendInkyServerLog("CMD inky_getstate ComputeSupportsFiles/SaveInkyState failed.", exSupports)
                                End If
                            End Try

                            Dim greeting As System.String = Nothing
                            If st.History.Count = 0 Then greeting = GetFriendlyGreeting()

                            If debugEnabled Then
                                AppendInkyServerLog("CMD inky_getstate before GetModelListForBrowserAsync.")
                            End If

                            Dim models As System.Collections.Generic.List(Of System.Object) =
                                Await GetModelListForBrowserAsync(st)

                            If debugEnabled Then
                                AppendInkyServerLog(
                                    "CMD inky_getstate after GetModelListForBrowserAsync. ModelCount=" &
                                    If(models?.Count, 0).ToString(System.Globalization.CultureInfo.InvariantCulture))
                            End If

                            Dim supportsTooling As Boolean = False
                            Dim toolingEnabled As Boolean = SyncToolingState(st, supportsTooling)

                            If debugEnabled Then
                                AppendInkyServerLog(
                                    "CMD inky_getstate after SyncToolingState. toolingEnabled=" &
                                    toolingEnabled.ToString() &
                                    "; supportsTooling=" &
                                    supportsTooling.ToString())
                            End If

                            Dim result As System.String =
                                JsonOk(New With {
                                    .ok = True,
                                    .history = ToBrowserTurns(st.History),
                                    .greeting = greeting,
                                    .models = models,
                                    .modelLabel = GetSelectedModelLabel(st.UseSecondApi, st.SelectedModelKey),
                                    .darkMode = st.DarkMode,
                                    .supportsFiles = st.SupportsFileUploads,
                                    .activeChat = activeChatId,
                                    .toolingEnabled = toolingEnabled,
                                    .supportsTooling = supportsTooling,
                                    .toolingLogEnabled = INI_ToolingLogWindow,
                                    .tools = GetToolListForBrowser(includeInteractiveM365Tools:=True),
                                    .advancedToolsEnabled = st.AgentModeEnabled,
                                    .agentFiles = GetAgentFileListForBrowser(),
                                    .agentWorkspace = GetAgentWorkspaceForBrowser(),
                                    .agentModelActive = st.AgentModelActive,
                                    .agentModelAvailable = IsAgentDefaultModelAvailable(),
                                    .inkyMemoryEnabled = My.Settings.Inky_InkyMemory
                                })

                            If debugEnabled Then
                                AppendInkyServerLog(
                                    "CMD inky_getstate end. ResultLen=" &
                                    result.Length.ToString(System.Globalization.CultureInfo.InvariantCulture))
                            End If

                            Return result
                        Catch ex As System.Exception
                            If INI_APIDebug Then
                                AppendInkyServerLog("CMD inky_getstate failed.", ex)
                            End If
                            Throw
                        End Try

                     ' Remaining cases kept intact (command-specific logic)
                    ' ------------------------------------------------------------------
                    Case "inky_claimscheduledtask"
                        Dim claimed = ClaimPendingScheduledBrowserTask()

                        If claimed IsNot Nothing Then
                            activeChatId = claimed.ChatId
                            Try
                                My.Settings.Inky_LastChat = activeChatId
                                My.Settings.Save()
                            Catch
                            End Try
                        End If

                        Return JsonOk(New With {
                                .ok = True,
                                .task = If(claimed Is Nothing,
                                           Nothing,
                                           New With {
                                               .taskId = claimed.TaskId,
                                               .prompt = claimed.Prompt,
                                               .chatId = claimed.ChatId
                                           })
                            })

                    Case "inky_switch"
                        Dim which As String = j("Chat")?.ToString()
                        activeChatId = If(which = "2", 2, 1)
                        Dim stSw = LoadInkyState()
                        _chatAdvancedToolsEnabled = stSw.AgentModeEnabled
                        If stSw.History.Count = 0 AndAlso Not stSw.DarkMode Then stSw.DarkMode = True
                        Dim greetingSwitch As String = If(stSw.History.Count = 0, GetFriendlyGreeting(), Nothing)

                        Try
                            My.Settings.Inky_LastChat = activeChatId
                            My.Settings.Save()
                        Catch
                        End Try

                        Dim supportsTooling As Boolean = False
                        Dim effectiveToolingEnabled As Boolean = SyncToolingState(stSw, supportsTooling)

                        ' Build model list for the switched chat's state
                        Dim models As List(Of Object) = Await GetModelListForBrowserAsync(stSw)

                        Return JsonOk(New With {
                                .ok = True,
                                .history = ToBrowserTurns(stSw.History),
                                .activeChat = activeChatId,
                                .darkMode = stSw.DarkMode,
                                .supportsFiles = ComputeSupportsFiles(stSw.UseSecondApi, stSw.SelectedModelKey),
                                .greeting = greetingSwitch,
                                .toolingEnabled = effectiveToolingEnabled,
                                .supportsTooling = supportsTooling,
                                .toolingLogEnabled = INI_ToolingLogWindow,
                                .models = models,
                                .advancedToolsEnabled = stSw.AgentModeEnabled,
                                .agentFiles = GetAgentFileListForBrowser(),
                                .agentWorkspace = GetAgentWorkspaceForBrowser(),
                                .agentModelActive = stSw.AgentModelActive,
                                .agentModelAvailable = IsAgentDefaultModelAvailable()
                            })

                    Case "inky_upload"

                        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "LocalChat_Upload invoked")

                        ' (Upload logic preserved)
                        ' ------------------------------------------------------------------
                        Try
                            Dim stU As InkyState = LoadInkyState()
                            Dim supports As System.Boolean = False
                            Try
                                supports = ComputeSupportsFiles(stU.UseSecondApi, stU.SelectedModelKey)
                            Catch
                                supports = False
                            End Try
                            If Not supports Then
                                Return JsonOk(New With {.ok = True, .supported = False})
                            End If
                            Dim name As System.String = j("Name")?.ToString()
                            Dim dataUrl As System.String = j("DataUrl")?.ToString()
                            If System.String.IsNullOrWhiteSpace(name) OrElse System.String.IsNullOrWhiteSpace(dataUrl) Then
                                Return JsonErr("Missing file data.")
                            End If
                            Dim commaIx As System.Int32 = dataUrl.IndexOf(","c)
                            If commaIx < 0 Then Return JsonErr("Bad DataURL.")
                            Dim b64 As System.String = dataUrl.Substring(commaIx + 1)
                            Dim bytes() As System.Byte
                            Try
                                bytes = System.Convert.FromBase64String(b64)
                            Catch exB64 As System.Exception
                                Return JsonErr("Invalid base64: " & exB64.Message)
                            End Try
                            Dim dir As System.String = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "InkyUploads")
                            If Not System.IO.Directory.Exists(dir) Then System.IO.Directory.CreateDirectory(dir)
                            Dim safeName As System.String = System.IO.Path.GetFileName(name)
                            For Each c As System.Char In System.IO.Path.GetInvalidFileNameChars()
                                safeName = safeName.Replace(c, "_"c)
                            Next
                            Dim unique As System.String = System.Guid.NewGuid().ToString("N")
                            Dim target As System.String = System.IO.Path.Combine(dir, unique & "_" & safeName)
                            System.IO.File.WriteAllBytes(target, bytes)
                            Return JsonOk(New With {.ok = True, .supported = True, .path = target, .name = safeName, .size = bytes.LongLength})
                        Catch exUp As System.Exception
                            Return JsonErr("Upload failed: " & exUp.Message)
                        End Try

                    Case "inky_cancel"

                        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "LocalChat_Cancel invoked")

                        ' (Cancellation logic preserved)
                        ' ------------------------------------------------------------------
                        ' Optional job id (preferred modern path)
                        Dim jobId As String = j("Job")?.ToString()

                        If Not String.IsNullOrWhiteSpace(jobId) Then
                            Dim job As LlmJob = Nothing
                            If Not jobMap.TryGetValue(jobId, job) Then
                                Return JsonErr("Unknown job.")
                            End If
                            Try
                                If job Is Nothing OrElse job.Cts Is Nothing Then
                                    Return JsonErr("Job has no cancellation handle.")
                                End If
                                If Not job.Cts.IsCancellationRequested Then
                                    job.Cts.Cancel()
                                    Return JsonOk(New With {
                                        .ok = True,
                                        .job = jobId,
                                        .status = "cancelRequested"
                                    })
                                Else
                                    Return JsonOk(New With {
                                        .ok = True,
                                        .job = jobId,
                                        .status = "alreadyCanceled"
                                    })
                                End If
                            Catch ex As System.Exception
                                Return JsonErr("Cancel failed: " & ex.Message)
                            End Try
                        End If
                        ' Legacy path
                        If llmOperationCts IsNot Nothing AndAlso Not llmOperationCts.IsCancellationRequested Then
                            Try
                                llmOperationCts.Cancel()
                                Return JsonOk(New With {
                                    .ok = True,
                                    .status = "cancelRequestedLegacy"
                                })
                            Catch ex As System.Exception
                                Return JsonErr("Legacy cancel failed: " & ex.Message)
                            End Try
                        End If
                        ' Fallback
                        Dim fallbackJob As LlmJob =
                            jobMap.Values _
                                  .Where(Function(x) x IsNot Nothing AndAlso
                                                     x.Tcs IsNot Nothing AndAlso
                                                     Not x.Tcs.Task.IsCompleted) _
                                  .OrderByDescending(Function(x) x.CreatedUtc) _
                                  .FirstOrDefault()
                        If fallbackJob IsNot Nothing Then
                            Try
                                If Not fallbackJob.Cts.IsCancellationRequested Then
                                    fallbackJob.Cts.Cancel()
                                    Return JsonOk(New With {
                                        .ok = True,
                                        .job = fallbackJob.Id,
                                        .status = "cancelRequestedFallback"
                                    })
                                Else
                                    Return JsonOk(New With {
                                        .ok = True,
                                        .job = fallbackJob.Id,
                                        .status = "alreadyCanceledFallback"
                                    })
                                End If
                            Catch ex As System.Exception
                                Return JsonErr("Fallback cancel failed: " & ex.Message)
                            End Try
                        End If
                        Return JsonErr("No active operation to cancel.")

                    Case "inky_send"
                        ' ------------------ (A) Read request & validate ------------------
                        Dim fileObject As System.String = j("FileObject")?.ToString()
                        Dim uploadedTempPath As System.String = fileObject
                        Dim textBody As System.String = j("Text")?.ToString()
                        Dim scheduledTaskId As System.String = j("ScheduledTaskId")?.ToString()
                        If System.String.IsNullOrWhiteSpace(textBody) Then
                            Return JsonErr("Please enter a message.")
                        End If
                        Try
                            My.Settings.Inky_LastPrompt = textBody
                            My.Settings.Save()
                        Catch
                        End Try
                        Dim st As InkyState = LoadInkyState()
                        ' Recompute upload capability (client may be stale)
                        Dim supportsFilesNow As System.Boolean = False
                        Try
                            supportsFilesNow = ComputeSupportsFiles(st.UseSecondApi, st.SelectedModelKey)
                        Catch
                            supportsFilesNow = False
                        End Try

                        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "LocalChat_Send invoked")

                        ' ------------------ (A2) Detect and strip (t) trigger ------------------
                        ' If the user includes "(t)" anywhere in the prompt, switch to the
                        ' ToolDefaultModel for this single request (same pattern as Form1.vb).
                        Dim toolTriggerDetected As Boolean = False
                        Dim toolTriggerConfig As ModelConfig = Nothing
                        Dim toolTriggerDisplayKey As String = Nothing

                        If textBody.IndexOf(ToolTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                            textBody = textBody.Replace(ToolTrigger, "").Trim()
                            If Not String.IsNullOrWhiteSpace(textBody) Then
                                Try
                                    If String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                                        Return JsonErr($"The {ToolTrigger} trigger requires an alternate model configuration file, but none is configured.")
                                    End If

                                    ' Snapshot → apply ToolDefaultModel → capture → restore
                                    Dim preConfig As ModelConfig = GetCurrentConfig(_context)
                                    Dim found As Boolean = GetSpecialTaskModel(
                                        _context, INI_AlternateModelPath, "ToolDefaultModel")

                                    If found Then
                                        toolTriggerConfig = GetCurrentConfig(_context)
                                        toolTriggerDisplayKey = If(Not String.IsNullOrWhiteSpace(toolTriggerConfig.ModelDescription),
                                                                   toolTriggerConfig.ModelDescription, toolTriggerConfig.Model)

                                        ' Immediately restore original config
                                        If originalConfigLoaded Then
                                            RestoreDefaults(_context, originalConfig)
                                        End If
                                        originalConfigLoaded = False

                                        ' Verify tooling support
                                        If ModelSupportsTooling(toolTriggerConfig) Then
                                            toolTriggerDetected = True

                                            ' Ensure sources are selected (same logic as Form1.vb)
                                            If _selectedToolsForChat Is Nothing OrElse _selectedToolsForChat.Count = 0 Then
                                                If Not EnsureToolsSelected(st, includeInteractiveM365Tools:=True) Then
                                                    ' No persisted tools — prompt user to select
                                                    Dim selectedTools As List(Of ModelConfig) = Nothing
                                                    Try
                                                        Await SwitchToUi(Sub()
                                                                             selectedTools = SelectToolsForSession(True, ToolFriendlyName, includeInteractiveM365Tools:=True)
                                                                         End Sub).ConfigureAwait(False)
                                                    Catch
                                                    End Try

                                                    If selectedTools IsNot Nothing AndAlso selectedTools.Count > 0 Then
                                                        _selectedToolsForChat = selectedTools
                                                        st.SelectedToolNames = selectedTools.Select(Function(tl) tl.ToolName).ToList()
                                                        st.ToolingEnabled = True
                                                        _chatToolingEnabled = True
                                                        SaveInkyState(st)
                                                    Else
                                                        ' User cancelled — abort
                                                        toolTriggerDetected = False
                                                        Return JsonErr($"The {ToolTrigger} trigger requires {ToolFriendlyName.ToLower} to be selected. Please select at least one source and try again.")
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        ' Restore if GetSpecialTaskModel partially modified state
                                        If originalConfigLoaded Then
                                            RestoreDefaults(_context, originalConfig)
                                        End If
                                        originalConfigLoaded = False
                                    End If
                                Catch
                                    ' Swallow — fall back to normal send
                                End Try
                            Else
                                ' Prompt was only "(t)" with nothing else
                                Return JsonErr($"The {ToolTrigger} trigger was used but no prompt text was provided.")
                            End If
                        End If

                        ' ------------------ (B) File / clipboard object extraction (unchanged logic) ------------
                        Dim extractedDoc As System.String = Nothing
                        Dim extractedLabel As System.String = Nothing
                        Dim attachedType As System.String = Nothing
                        Dim hadInlineExtraction As System.Boolean = False
                        If Not System.String.IsNullOrWhiteSpace(fileObject) Then
                            Dim okOffice As System.Boolean = False
                            Try
                                okOffice = TryExtractOfficeText(fileObject, extractedDoc, extractedLabel)
                            Catch
                                okOffice = False
                            End Try
                            If okOffice Then
                                hadInlineExtraction = True
                                attachedType = "office"
                                fileObject = Nothing    ' Do NOT pass a path to the model
                            Else
                                Dim okText As System.Boolean = False
                                Try
                                    okText = TryExtractTextLike(fileObject, extractedDoc, extractedLabel)
                                Catch
                                    okText = False
                                End Try
                                If okText Then
                                    hadInlineExtraction = True
                                    attachedType = "text"
                                    fileObject = Nothing
                                Else
                                    If (Not supportsFilesNow) AndAlso (Not System.String.IsNullOrWhiteSpace(fileObject)) Then
                                        ' Model does not support file objects
                                        st.History.Add(New ChatTurn With {
                                            .Role = "assistant",
                                            .Markdown = "This model does not support file attachments.",
                                            .Html = MarkdownToHtml("This model does not support file attachments."),
                                            .Utc = System.DateTime.UtcNow
                                        })
                                        SaveInkyState(st)
                                        ' Cleanup temp
                                        Try
                                            If Not System.String.IsNullOrWhiteSpace(uploadedTempPath) AndAlso IO.File.Exists(uploadedTempPath) Then
                                                IO.File.Delete(uploadedTempPath)
                                            End If
                                        Catch
                                        End Try
                                        Return JsonOk(New With {.ok = True, .history = ToBrowserTurns(st.History)})
                                    End If
                                    ' Else: keep fileObject (e.g. PDF) for API that supports raw path/object
                                End If
                            End If
                        End If
                        ' ------------------ (C) Append user turn immediately ------------------
                        Dim userTurn As New ChatTurn With {
                            .Role = "user",
                            .Markdown = textBody,
                            .Html = MarkdownToHtml(textBody),
                            .Utc = Date.UtcNow
                        }
                        st.History.Add(userTurn)
                        ' Cap history (unchanged)
                        Dim cap As Integer = 0
                        Try : cap = INI_ChatCap : Catch : cap = 4000 : End Try
                        Dim clipped As List(Of ChatTurn) = CapHistoryToChars(st, cap)
                        ' Build dialog prompt now (passed into job)
                        Dim sbDialog As New System.Text.StringBuilder()
                        sbDialog.AppendLine("<DIALOG>")
                        For Each t In clipped
                            If t.Role = "user" Then
                                sbDialog.AppendLine("[USER] " & t.Markdown)
                            Else
                                sbDialog.AppendLine("[ASSISTANT] " & t.Markdown)
                            End If
                        Next
                        sbDialog.AppendLine("</DIALOG>")
                        ' Add inline extracted document block if any (unchanged)
                        If hadInlineExtraction AndAlso Not String.IsNullOrWhiteSpace(extractedDoc) Then
                            sbDialog.AppendLine()
                            Dim lbl As String = EscapeForXml(If(extractedLabel, "Attached document"))
                            Dim typ As String = If(String.IsNullOrWhiteSpace(attachedType), "text", attachedType)
                            sbDialog.AppendLine("<ATTACHED_DOCUMENT type=""" & typ & """ label=""" & lbl & """>")
                            sbDialog.AppendLine(extractedDoc)
                            sbDialog.AppendLine("</ATTACHED_DOCUMENT>")
                        End If

                        ' Add agent file listing so the model knows which files are already loaded/staged
                        If _chatAdvancedToolsEnabled AndAlso _chatAgentFiles IsNot Nothing AndAlso _chatAgentFiles.Count > 0 Then
                            sbDialog.AppendLine()
                            sbDialog.AppendLine("[LOADED FILES]")
                            For i As Integer = 0 To _chatAgentFiles.Count - 1
                                Dim att = _chatAgentFiles(i)
                                Dim sizeStr = If(att.SizeBytes > 0, $" ({att.SizeBytes / 1024:F0} KB)", "")
                                Dim statusStr = If(att.IsOverSizeLimit, " [OVER SIZE LIMIT - cannot process]", "")
                                sbDialog.AppendLine($"  {i + 1}. {att.OriginalFileName}{sizeStr}{statusStr}")
                            Next
                            sbDialog.AppendLine("[/LOADED FILES]")
                            sbDialog.AppendLine()
                            sbDialog.AppendLine("These are the files already loaded into the current agent session. Tools such as list_attachments, read_attachment, process_word_document, extract_pdf_text, compare_word_documents, and similar session tools operate on these loaded files.")
                        End If

                        ' Add workspace context independently of whether any files are already staged
                        If _chatAdvancedToolsEnabled Then
                            Dim workspacePromptBlock As String = BuildAgentWorkspacePromptBlock()
                            If Not String.IsNullOrWhiteSpace(workspacePromptBlock) Then
                                sbDialog.AppendLine()
                                sbDialog.AppendLine(workspacePromptBlock)
                            End If
                        End If

                        If _chatAdvancedToolsEnabled Then
                            Dim sessionFilesPromptBlock As String = BuildAgentSessionFilesPromptBlock()
                            If Not String.IsNullOrWhiteSpace(sessionFilesPromptBlock) Then
                                sbDialog.AppendLine()
                                sbDialog.AppendLine(sessionFilesPromptBlock)
                            End If

                            Dim workspaceHistoryPromptBlock As String = BuildAgentWorkspaceHistoryPromptBlock()
                            If Not String.IsNullOrWhiteSpace(workspaceHistoryPromptBlock) Then
                                sbDialog.AppendLine()
                                sbDialog.AppendLine(workspaceHistoryPromptBlock)
                            End If
                        End If

                        ' Persist state with user turn BEFORE returning (important)
                        SaveInkyState(st)
                        ' ------------------ (D) Prepare system prompt (same logic) ------------------
                        Dim sysPromptBase As String = _context.SP_Chat.Replace("{Location}", Location)
                        Dim nowLocal As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss zzz", Globalization.CultureInfo.InvariantCulture)
                        sysPromptBase &= Environment.NewLine & "Current local date/time: " & nowLocal
                        sysPromptBase &= Environment.NewLine & $"Your name is '{AN6}'. "

                        If _chatAdvancedToolsEnabled Then
                            sysPromptBase &= Environment.NewLine &
                                "Local Chat Agent behavior: Tools are optional. If the user's latest request can be answered directly, especially for creative writing, drafting, rewriting, brainstorming, summarizing, explanation, or ordinary chat, answer directly without using tools. Do not claim that you are unable to perform a normal language-generation task merely because no tool, attachment, or workspace action is required. Use tools only when they are actually needed for files, workspace operations, document processing, or external/source-backed work."
                        End If

                        ' Inject InkyMemory into system prompt if enabled
                        If My.Settings.Inky_InkyMemory Then
                            Dim memoryContent = ReadInkyMemory(_context.INI_InkyMemoryCap)
                            sysPromptBase &= Environment.NewLine & _context.SP_Add_InkyMemory
                            If Not String.IsNullOrWhiteSpace(memoryContent) Then
                                sysPromptBase &= Environment.NewLine & "<INKY_MEMORY_CURRENT>" & Environment.NewLine & memoryContent & Environment.NewLine & "</INKY_MEMORY_CURRENT>"
                            End If
                        End If

                        Dim useSecondApiLocal As Boolean = st.UseSecondApi
                        Dim selectedModelKeyLocal As String = st.SelectedModelKey
                        Dim supportsToolingForJob As Boolean = CurrentModelSupportsTooling(st)
                        Dim toolingEnabledForJob As Boolean = st.ToolingEnabled AndAlso supportsToolingForJob
                        Dim agentModeEnabledForJob As Boolean = st.AgentModeEnabled AndAlso toolingEnabledForJob
                        Dim selectedToolsForJob As List(Of ModelConfig) =
                            GetLocalChatEffectiveSelection(st, includeInteractiveM365Tools:=True)
                        ' Capture file object (may be Nothing after extraction)
                        Dim finalFileObject As String = fileObject
                        Dim tempUploadPathCopy As String = uploadedTempPath
                        ' ------------------ (E) Create background job ------------------
                        Dim jobId As String = Guid.NewGuid().ToString("N")
                        Dim jobCts As New CancellationTokenSource()
                        Dim tcs As New TaskCompletionSource(Of String)(TaskCreationOptions.RunContinuationsAsynchronously)
                        Dim originatingChatId As Integer = activeChatId

                        Dim job As New LlmJob With {
                                        .Id = jobId,
                                        .CreatedUtc = Date.UtcNow,
                                        .Tcs = tcs,
                                        .Cts = jobCts,
                                        .UseSecond = useSecondApiLocal,
                                        .FileObject = finalFileObject,
                                        .ChatId = originatingChatId,
                                        .ScheduledTaskId = If(scheduledTaskId, "")
                                    }
                        If Not jobMap.TryAdd(jobId, job) Then
                            jobCts.Dispose()
                            Return JsonErr("Failed to register job.")
                        End If
                        Threading.Interlocked.Increment(activeJobs)
                        ' Keep old global CTS reference for legacy inky_cancel support
                        Try
                            If llmOperationCts IsNot Nothing Then
                                llmOperationCts.Cancel()
                                llmOperationCts.Dispose()
                            End If
                        Catch
                        End Try
                        llmOperationCts = jobCts

                        If Not String.IsNullOrWhiteSpace(scheduledTaskId) Then
                            SchedulerMarkLocalTaskRunning(scheduledTaskId)
                        End If

                        ' Replace the original System.Threading.Tasks.Task.Run(Sub() ... End Sub) block with this:
                        System.Threading.Tasks.Task.Run(
                            Sub()
                                Dim configSnapshot As ModelConfig = Nothing
                                Dim snapshotTaken As Boolean = False
                                Dim alternateApplied As Boolean = False
                                Dim localOutput As String = ""
                                Dim agentAbortDetected As Boolean = False
                                Dim agentToolCallLogSnapshot As List(Of AutoPilotToolCallEntry) = Nothing
                                Dim agentOutputFiles As List(Of String) = Nothing
                                Dim scheduledTaskFinalized As Boolean = False
                                Try
                                    ' (1) Alternate model application (safer pattern)
                                    If useSecondApiLocal AndAlso Not String.IsNullOrWhiteSpace(selectedModelKeyLocal) Then
                                        Try
                                            SyncLock AlternateModelLock
                                                configSnapshot = GetCurrentConfig(_context)
                                                snapshotTaken = True
                                                Dim alts = LoadAlternativeModels(INI_AlternateModelPath, _context)
                                                Dim sel = alts?.FirstOrDefault(
                                                    Function(m)
                                                        If m Is Nothing Then Return False
                                                        If Not String.IsNullOrWhiteSpace(m.ModelDescription) AndAlso
                                                           String.Equals(m.ModelDescription, selectedModelKeyLocal, StringComparison.OrdinalIgnoreCase) Then Return True
                                                        If Not String.IsNullOrWhiteSpace(m.Model) AndAlso
                                                           String.Equals(m.Model, selectedModelKeyLocal, StringComparison.OrdinalIgnoreCase) Then Return True
                                                        Return False
                                                    End Function)
                                                If sel IsNot Nothing Then
                                                    ApplyModelConfig(_context, sel)
                                                    alternateApplied = True
                                                End If
                                            End SyncLock
                                        Catch
                                            ' Swallow – we will still attempt restore if snapshotTaken
                                        End Try
                                    End If
                                    ' (2) Run LLM - with or without tooling
                                    Dim useAgentMode As Boolean =
                                        agentModeEnabledForJob AndAlso
                                        supportsToolingForJob AndAlso
                                        Not _apActive

                                    Dim agentToolsForJob As List(Of ModelConfig) = Nothing

                                    ' (t) trigger: override model selection for this single call
                                    Dim useToolTrigger As Boolean = toolTriggerDetected AndAlso toolTriggerConfig IsNot Nothing
                                    If useToolTrigger Then
                                        ' Apply the ToolDefaultModel config for this request
                                        Try
                                            SyncLock AlternateModelLock
                                                If Not snapshotTaken Then
                                                    configSnapshot = GetCurrentConfig(_context)
                                                    snapshotTaken = True
                                                End If
                                                ApplyModelConfig(_context, toolTriggerConfig)
                                            End SyncLock
                                        Catch
                                        End Try
                                    End If

                                    If useAgentMode Then
                                        agentToolsForJob = ChatAgentSetupToolContext(selectedToolsForJob)
                                    End If

                                    Try
                                        If useToolTrigger AndAlso selectedToolsForJob IsNot Nothing AndAlso selectedToolsForJob.Count > 0 Then
                                            ' (t) trigger path: use ToolDefaultModel with selected sources
                                            ' Respects INI_ToolingLogWindow (same as Form1.vb chkShowToolingLog)
                                            localOutput = ExecuteToolingLoop(
                                                    sysPromptBase,
                                                    "",
                                                    selectedToolsForJob,
                                                    True,
                                                    finalFileObject,
                                                    False, "",
                                                    False, False, "",
                                                    False, "",
                                                    False, "", "", "",
                                                    sbDialog.ToString(),
                                                    True,
                                                    Not INI_ToolingLogWindow,
                                                    False,
                                                    jobCts.Token
                                                ).GetAwaiter().GetResult()

                                            ' Restore config AFTER tooling completes
                                            If snapshotTaken Then
                                                SyncLock AlternateModelLock
                                                    RestoreDefaults(_context, configSnapshot)
                                                End SyncLock
                                                snapshotTaken = False
                                            End If
                                        ElseIf useAgentMode AndAlso agentToolsForJob IsNot Nothing AndAlso agentToolsForJob.Count > 0 Then
                                            ' Match AutoPilot's iteration limit for agent/tool mode
                                            Dim previousMaxIterations = INI_ToolingMaximumIterations
                                            INI_ToolingMaximumIterations = AP_MaxToolIterations
                                            Try
                                                localOutput = ExecuteToolingLoop(
                                                    sysPromptBase,
                                                    "",
                                                    agentToolsForJob,
                                                    useSecondApiLocal,
                                                    finalFileObject,
                                                    False, "",
                                                    False, False, "",
                                                    False, "",
                                                    False, "", "", "",
                                                    sbDialog.ToString(),
                                                    True,
                                                    Not INI_ToolingLogWindow,
                                                    False,
                                                    jobCts.Token,
                                                    _chatAgentTempDir
                                                ).GetAwaiter().GetResult()
                                            Finally
                                                INI_ToolingMaximumIterations = previousMaxIterations
                                            End Try

                                            ' Restore config AFTER tooling completes
                                            If snapshotTaken Then
                                                SyncLock AlternateModelLock
                                                    RestoreDefaults(_context, configSnapshot)
                                                End SyncLock
                                                snapshotTaken = False
                                            End If

                                            agentAbortDetected = ShouldBuildChatAgentAbortReport(localOutput, jobCts.IsCancellationRequested)

                                            ' Snapshot before teardown clears the AutoPilot shared state
                                            agentToolCallLogSnapshot = CloneChatAgentToolCallLog()

                                            ' Remove knowledge-store source files not cited by the LLM
                                            RemoveUncitedKnowledgeSourceCopies(localOutput)

                                            ' Collect outputs to Desktop\Inky\...
                                            Dim citedExternalFiles =
                                                ExtractCitedLocalFilePaths(localOutput).
                                                    Where(Function(p)
                                                              If String.IsNullOrWhiteSpace(_chatAgentTempDir) Then Return True
                                                              Return Not IsPathWithinOrEqual(p, _chatAgentTempDir)
                                                          End Function).
                                                    ToList()

                                            agentOutputFiles = ChatAgentCollectAndCopyOutputs(citedExternalFiles)

                                            If Not String.IsNullOrWhiteSpace(job.ScheduledTaskId) AndAlso
                                               agentOutputFiles IsNot Nothing AndAlso
                                               agentOutputFiles.Count > 0 Then
                                                SchedulerPersistLocalBrowserOutputs(job.ScheduledTaskId, agentOutputFiles)
                                            End If

                                            If agentAbortDetected Then
                                                localOutput = BuildChatAgentAbortReport(
                                                    agentToolCallLogSnapshot,
                                                    agentOutputFiles,
                                                    localOutput)
                                            ElseIf agentOutputFiles IsNot Nothing AndAlso agentOutputFiles.Count > 0 Then
                                                localOutput = RemoveGeneratedOutputFilesSections(If(localOutput, String.Empty))

                                                Dim outputFilesMarkdown As String = BuildOutputFilesMarkdown(agentOutputFiles)
                                                If Not String.IsNullOrWhiteSpace(outputFilesMarkdown) Then
                                                    localOutput = localOutput.TrimEnd() & vbCrLf & vbCrLf & outputFilesMarkdown
                                                End If
                                            End If

                                        ElseIf toolingEnabledForJob AndAlso selectedToolsForJob IsNot Nothing AndAlso selectedToolsForJob.Count > 0 Then
                                            localOutput = ExecuteToolingLoop(
                                                    sysPromptBase,
                                                    "",
                                                    selectedToolsForJob,
                                                    useSecondApiLocal,
                                                    finalFileObject,
                                                    False, "",
                                                    False, False, "",
                                                    False, "",
                                                    False, "", "", "",
                                                    sbDialog.ToString(),
                                                    True,
                                                    Not INI_ToolingLogWindow,
                                                    False,
                                                    jobCts.Token
                                                ).GetAwaiter().GetResult()

                                            ' Restore config AFTER tooling completes
                                            If snapshotTaken Then
                                                SyncLock AlternateModelLock
                                                    RestoreDefaults(_context, configSnapshot)
                                                End SyncLock
                                                snapshotTaken = False
                                            End If
                                        Else
                                            ' Standard LLM call
                                            localOutput = RunLlmAsync(
                                                    sysPromptBase,
                                                    sbDialog.ToString(),
                                                    useSecondApiLocal,
                                                    False,
                                                    finalFileObject,
                                                    jobCts.Token
                                                ).GetAwaiter().GetResult()
                                        End If
                                    Finally
                                        If useAgentMode Then
                                            ChatAgentTeardownToolContext()
                                        End If
                                    End Try
                                    If localOutput Is Nothing Then localOutput = String.Empty
                                    localOutput = SanitizeModelOutputForBrowser(localOutput).Trim()
                                    If localOutput.Length > 0 AndAlso
                                       localOutput.Equals("Operation was canceled by the user.", StringComparison.OrdinalIgnoreCase) Then
                                        localOutput = "Aborted by user."
                                    End If

                                    ' Process InkyMemory updates from LLM response (if enabled)
                                    If My.Settings.Inky_InkyMemory Then
                                        localOutput = ProcessInkyMemoryResponse(localOutput, _context.INI_InkyMemoryCap)
                                    End If

                                    ' (3) Build assistant turn or error turn
                                    Dim assistantText As String = localOutput
                                    Dim wasCanceled As Boolean = jobCts.IsCancellationRequested

                                    If assistantText.Length = 0 Then
                                        assistantText = If(wasCanceled,
                                                   "Aborted by user.",
                                                   "Error: The model did not provide a response.")
                                    End If

                                    Dim htmlOut As String = MarkdownToHtml(assistantText)
                                    Dim persisted As Boolean = PersistAssistantTurnForJob(job, assistantText, htmlOut)

                                    If Not persisted Then
                                        Debug.WriteLine(
                                    $"[Inky] Assistant result was returned but could not be persisted for job={job.Id}; chat={job.ChatId}")
                                    End If

                                    If Not String.IsNullOrWhiteSpace(job.ScheduledTaskId) AndAlso Not scheduledTaskFinalized Then
                                        SchedulerCompleteLocalBrowserTask(job.ScheduledTaskId, assistantText)
                                        scheduledTaskFinalized = True
                                    End If

                                    If wasCanceled AndAlso localOutput.Length = 0 Then
                                        tcs.TrySetCanceled()
                                    Else
                                        tcs.TrySetResult(assistantText)
                                    End If
                                Catch exOp As OperationCanceledException
                                    If Not String.IsNullOrWhiteSpace(job.ScheduledTaskId) AndAlso Not scheduledTaskFinalized Then
                                        SchedulerFailLocalBrowserTask(job.ScheduledTaskId, "Aborted by user.")
                                        scheduledTaskFinalized = True
                                    End If
                                    tcs.TrySetCanceled()
                                Catch ex As System.Exception
                                    If Not String.IsNullOrWhiteSpace(job.ScheduledTaskId) AndAlso Not scheduledTaskFinalized Then
                                        SchedulerFailLocalBrowserTask(job.ScheduledTaskId, ex.Message)
                                        scheduledTaskFinalized = True
                                    End If
                                    tcs.TrySetException(ex)
                                Finally
                                    If Not String.IsNullOrWhiteSpace(job.ScheduledTaskId) AndAlso
                                       Not scheduledTaskFinalized AndAlso
                                       jobCts.IsCancellationRequested Then
                                        SchedulerFailLocalBrowserTask(job.ScheduledTaskId, "Aborted by user.")
                                    End If

                                    ' (4) Cleanup temp upload
                                    Try
                                        If Not String.IsNullOrWhiteSpace(tempUploadPathCopy) AndAlso IO.File.Exists(tempUploadPathCopy) Then
                                            IO.File.Delete(tempUploadPathCopy)
                                        End If
                                    Catch
                                    End Try
                                    ' (5) Restore original config unconditionally if snapshot was taken
                                    Try
                                        If snapshotTaken Then
                                            SyncLock AlternateModelLock
                                                RestoreDefaults(_context, configSnapshot)
                                            End SyncLock
                                        End If
                                    Catch
                                        ' Ignore restore errors
                                    End Try
                                    Threading.Interlocked.Decrement(activeJobs)
                                End Try
                            End Sub)
                        ' ------------------ (F) Immediate response with job id ------------------
                        Return JsonOk(New With {
                            .ok = True,
                            .job = jobId,
                            .status = "running",
                            .history = ToBrowserTurns(st.History)
                        })

                    Case "inky_pure"

                        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "LocalChat_Pure invoked")

                        ' (Pure mode logic preserved)
                        ' ------------------------------------------------------------------
                        Dim fileObject As String = j("FileObject")?.ToString()
                        Dim textBody As String = j("Text")?.ToString()
                        If String.IsNullOrWhiteSpace(textBody) Then
                            Return JsonErr("Please enter a message.")
                        End If
                        Try
                            My.Settings.Inky_LastPrompt = textBody
                            My.Settings.Save()
                        Catch
                        End Try
                        Dim stPure As InkyState = LoadInkyState()
                        ' Decide model context (still respect selected model choice)
                        Dim useSecondApiLocal As Boolean = stPure.UseSecondApi
                        Dim selectedModelKeyLocal As String = stPure.SelectedModelKey
                        ' Validate file object support (only pass through if supported; no extraction)
                        Dim supportsFilesNow As Boolean = False
                        Try
                            supportsFilesNow = ComputeSupportsFiles(useSecondApiLocal, selectedModelKeyLocal)
                        Catch
                            supportsFilesNow = False
                        End Try
                        If Not supportsFilesNow Then
                            fileObject = Nothing
                        End If
                        ' Record user turn (prefixed)
                        stPure.History.Add(New ChatTurn With {
                            .Role = "user",
                            .Markdown = "Pure: " & textBody,
                            .Html = MarkdownToHtml("Pure: " & textBody),
                            .Utc = Date.UtcNow
                        })
                        SaveInkyState(stPure)
                        ' Prepare background job (mirrors inky_send logic but no dialog/system prompt)
                        Dim jobIdP As String = Guid.NewGuid().ToString("N")
                        Dim jobCtsP As New CancellationTokenSource()
                        Dim tcsP As New TaskCompletionSource(Of String)(TaskCreationOptions.RunContinuationsAsynchronously)
                        Dim originatingChatIdP As Integer = activeChatId

                        Dim jobP As New LlmJob With {
                                    .Id = jobIdP,
                                    .CreatedUtc = Date.UtcNow,
                                    .Tcs = tcsP,
                                    .Cts = jobCtsP,
                                    .UseSecond = useSecondApiLocal,
                                    .FileObject = fileObject,
                                    .ChatId = originatingChatIdP
                                }
                        If Not jobMap.TryAdd(jobIdP, jobP) Then
                            jobCtsP.Dispose()
                            Return JsonErr("Failed to register job.")
                        End If
                        Threading.Interlocked.Increment(activeJobs)
                        ' Maintain legacy cancellation
                        Try
                            If llmOperationCts IsNot Nothing Then
                                llmOperationCts.Cancel()
                                llmOperationCts.Dispose()
                            End If
                        Catch
                        End Try
                        llmOperationCts = jobCtsP
                        System.Threading.Tasks.Task.Run(
                            Sub()
                                Dim originalCfgLoadedP As Boolean = False
                                Dim usedAlternate As Boolean = False
                                Try
                                    ' Apply alternate model if a specific alternate is selected
                                    If useSecondApiLocal AndAlso Not String.IsNullOrWhiteSpace(selectedModelKeyLocal) Then
                                        Try
                                            originalConfig = GetCurrentConfig(_context) : originalCfgLoadedP = True
                                            Dim alts = LoadAlternativeModels(INI_AlternateModelPath, _context)
                                            Dim sel = alts?.FirstOrDefault(
                                                Function(m)
                                                    If m Is Nothing Then Return False
                                                    If Not String.IsNullOrWhiteSpace(m.ModelDescription) AndAlso
                                                       String.Equals(m.ModelDescription, selectedModelKeyLocal, StringComparison.OrdinalIgnoreCase) Then Return True
                                                    If Not String.IsNullOrWhiteSpace(m.Model) AndAlso
                                                       String.Equals(m.Model, selectedModelKeyLocal, StringComparison.OrdinalIgnoreCase) Then Return True
                                                    Return False
                                                End Function)
                                            If sel IsNot Nothing Then
                                                ApplyModelConfig(_context, sel)
                                                usedAlternate = True
                                            End If
                                        Catch
                                        End Try
                                    End If
                                    ' Raw call: NO system prompt, NO history packaging
                                    Dim output = RunLlmAsync("", textBody, useSecondApiLocal, False, fileObject, jobCtsP.Token).GetAwaiter().GetResult()
                                    If output Is Nothing Then output = String.Empty
                                    output = SanitizeModelOutputForBrowser(output).Trim()

                                    Dim assistantText As String = output
                                    Dim wasCanceled As Boolean = jobCtsP.IsCancellationRequested

                                    If assistantText.Length = 0 Then
                                        assistantText = If(wasCanceled,
                                                           "Aborted by user.",
                                                           "Error: The model returned no content.")
                                    End If

                                    Dim htmlOut As String = MarkdownToHtml(assistantText)
                                    Dim persisted As Boolean = PersistAssistantTurnForJob(jobP, assistantText, htmlOut)

                                    If Not persisted Then
                                        Debug.WriteLine(
                                            $"[Inky] Pure result was returned but could not be persisted for job={jobP.Id}; chat={jobP.ChatId}")
                                    End If

                                    If wasCanceled AndAlso output.Length = 0 Then
                                        tcsP.TrySetCanceled()
                                    Else
                                        tcsP.TrySetResult(assistantText)
                                    End If
                                Catch exOp As OperationCanceledException
                                    tcsP.TrySetCanceled()
                                Catch ex As System.Exception
                                    tcsP.TrySetException(ex)
                                Finally
                                    ' Restore config if alternate used
                                    Try
                                        If useSecondApiLocal AndAlso (usedAlternate OrElse originalCfgLoadedP) AndAlso originalConfigLoaded Then
                                            RestoreDefaults(_context, originalConfig)
                                            originalConfigLoaded = False
                                        End If
                                    Catch
                                    End Try
                                    Threading.Interlocked.Decrement(activeJobs)
                                End Try
                            End Sub)
                        Return JsonOk(New With {
                            .ok = True,
                            .job = jobIdP,
                            .status = "running",
                            .history = ToBrowserTurns(stPure.History)
                        })


                    Case "inky_jobstatus"
                        ' (Job status check preserved)
                        ' ------------------------------------------------------------------
                        Dim jobId = j("Job")?.ToString()
                        If String.IsNullOrWhiteSpace(jobId) Then Return JsonErr("Missing Job id.")
                        Dim job As LlmJob = Nothing
                        If Not jobMap.TryGetValue(jobId, job) Then
                            Return JsonErr("Unknown job.")
                        End If
                        ' Cleanup old jobs (lazy TTL)
                        If (Date.UtcNow - job.CreatedUtc).TotalMinutes > JobTtlMinutes Then
                            Dim dump As LlmJob = Nothing
                            jobMap.TryRemove(jobId, dump)
                            Return JsonErr("Job expired.")
                        End If
                        Dim t = job.Tcs.Task
                        If Not t.IsCompleted Then
                            Return JsonOk(New With {.ok = True, .job = jobId, .status = "running"})
                        End If
                        If t.IsCanceled Then
                            Return JsonOk(New With {.ok = True, .job = jobId, .status = "canceled"})
                        ElseIf t.IsFaulted Then
                            Return JsonOk(New With {.ok = False, .job = jobId, .status = "error", .error = t.Exception.GetBaseException().Message})
                        Else
                            Dim resultText As String = If(t.Result, "")
                            Dim stJob As InkyState = LoadInkyState(job.ChatId)

                            Return JsonOk(New With {
                                    .ok = True,
                                    .job = jobId,
                                    .status = "done",
                                    .result = resultText,
                                    .resultHtml = MarkdownToHtml(resultText),
                                    .chat = job.ChatId,
                                    .history = ToBrowserTurns(If(stJob?.History, New System.Collections.Generic.List(Of ChatTurn)()))
                                })
                        End If

                    Case "inky_canceljob"

                        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "LocalChat_CancelJob invoked")

                        ' (Direct cancel by id)
                        ' ------------------------------------------------------------------
                        Dim jobId = j("Job")?.ToString()
                        If String.IsNullOrWhiteSpace(jobId) Then Return JsonErr("Missing Job id.")
                        Dim job As LlmJob = Nothing
                        If Not jobMap.TryGetValue(jobId, job) Then Return JsonErr("Unknown job.")
                        Try
                            job.Cts.Cancel()
                            Return JsonOk(New With {.ok = True, .job = jobId, .status = "cancelRequested"})
                        Catch ex As System.Exception
                            Return JsonErr("Cancel failed: " & ex.Message)
                        End Try

                    Case "inky_promptlibpick"
                        Try
                            If Not INI_PromptLib Then
                                Return JsonOk(New With {.ok = True, .prompt = ""})
                            End If

                            Dim ownerHandle As IntPtr = IntPtr.Zero
                            Try
                                ownerHandle = NativeMethods.GetForegroundWindow()
                            Catch
                            End Try

                            Dim selectedPrompt As String =
                                Await SwitchToUi(
                                    Function()
                                        Return SharedMethods.ShowPromptInsertionSelector(
                                            INI_PromptLibPath,
                                            INI_PromptLibPathLocal,
                                            _context,
                                            Nothing,
                                            My.Settings.Inky_LastPrompt,
                                            ownerHandle
                                        )
                                    End Function).ConfigureAwait(False)

                            Return JsonOk(New With {
                                .ok = True,
                                .prompt = If(selectedPrompt, "")
                            })
                        Catch ex As Exception
                            Return JsonErr("Failed to open prompt library: " & ex.Message)
                        End Try

                    Case "inky_toword"

                        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "LocalChat_ToWord invoked")

                        Try
                            Dim st As InkyState = LoadInkyState()
                            Dim turns = If(st?.History, New System.Collections.Generic.List(Of ChatTurn)())

                            Dim sb As New System.Text.StringBuilder()
                            sb.AppendLine("# " & GetBotName() & " — Chat thread")
                            sb.AppendLine()
                            sb.AppendLine("Generated: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss zzz", Globalization.CultureInfo.InvariantCulture))
                            sb.AppendLine()

                            For Each t In turns
                                If t Is Nothing Then Continue For
                                Dim who As String = If(String.Equals(t.Role, "user", StringComparison.OrdinalIgnoreCase), "You", GetBotName())
                                sb.AppendLine("## " & who)
                                sb.AppendLine(t.Markdown)
                                sb.AppendLine()
                            Next

                            Dim markdownThread As String = sb.ToString()

                            Dim ok As Boolean = Await TryCreateWordDocFromMarkdown(markdownThread).ConfigureAwait(False)

                            If Not ok Then Return JsonErr("Could not create a Word document.")

                            Return JsonOk(New With {.ok = True})
                        Catch ex As System.Exception
                            Return JsonErr("To Word failed: " & ex.Message)
                        End Try

                    Case "inky_clear"
                        ' Clear history ONLY; preserve model/tooling selections.
                        Dim stClear As InkyState = LoadInkyState()

                        stClear.History = New System.Collections.Generic.List(Of ChatTurn)()
                        stClear.LastAssistantText = ""

                        ' Clear agent files on chat clear
                        ChatAgentClearFiles()

                        ' Keep ToolingEnabled / SelectedToolNames / model selection / dark mode as-is
                        SaveInkyState(stClear)

                        Dim supportsTooling As Boolean = False
                        Dim effectiveToolingEnabled As Boolean = SyncToolingState(stClear, supportsTooling)

                        Return JsonOk(New With {
                            .ok = True,
                            .activeChat = activeChatId,
                            .greeting = GetFriendlyGreeting(),
                            .toolingEnabled = effectiveToolingEnabled,
                            .supportsTooling = supportsTooling
                        })

                    Case "inky_copylast"

                        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "LocalChat_CopyLast invoked")

                        ' (Copy last assistant response)
                        ' ------------------------------------------------------------------
                        Dim stCopy As InkyState = LoadInkyState()
                        If System.String.IsNullOrWhiteSpace(stCopy.LastAssistantText) Then
                            Return JsonErr("No assistant response available to copy.")
                        End If
                        Await SwitchToUi(Sub()
                                             SLib.PutInClipboard(MarkdownToRtfConverter.Convert(stCopy.LastAssistantText))
                                         End Sub).ConfigureAwait(False)
                        Return JsonOk(New With {.ok = True})

                    Case "inky_setmodel"

                        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "LocalChat_SetModel invoked")

                        ' (Model selection)
                        ' ------------------------------------------------------------------
                        Dim key As System.String = j("Key")?.ToString()
                        Dim st As InkyState = LoadInkyState()
                        If System.String.IsNullOrWhiteSpace(key) OrElse System.String.Equals(key, "default", System.StringComparison.OrdinalIgnoreCase) Then
                            st.UseSecondApi = False
                            st.SelectedModelKey = ""
                        ElseIf System.String.Equals(key, "__second__", System.StringComparison.OrdinalIgnoreCase) Then
                            st.UseSecondApi = True
                            st.SelectedModelKey = ""
                        Else
                            st.UseSecondApi = True
                            st.SelectedModelKey = key
                        End If
                        ' Reset agent model toggle when user manually selects a different model
                        If st.AgentModelActive Then
                            st.AgentModelActive = False
                            st.PreAgentModelKey = ""
                            st.PreAgentUseSecondApi = False
                        End If

                        Try
                            My.Settings.Inky_UseSecondApiSelected = st.UseSecondApi
                            My.Settings.Inky_SelectedModelKey = st.SelectedModelKey
                            My.Settings.Save()
                        Catch
                        End Try
                        ' Re-evaluate upload capability for selected model
                        Try
                            st.SupportsFileUploads = ComputeSupportsFiles(st.UseSecondApi, st.SelectedModelKey)
                        Catch
                            st.SupportsFileUploads = False
                        End Try
                        ' Check if new model supports tooling
                        Dim supportsTooling = CurrentModelSupportsTooling(st)

                        If Not supportsTooling Then
                            Dim hadAnyToolingState As Boolean =
                                st.ToolingEnabled OrElse
                                st.AgentModeEnabled OrElse
                                _chatToolingEnabled OrElse
                                _chatAdvancedToolsEnabled

                            st.ToolingEnabled = False
                            st.AgentModeEnabled = False
                            _chatToolingEnabled = False
                            _chatAdvancedToolsEnabled = False

                            If hadAnyToolingState Then
                                ChatAgentClearFiles()
                            End If
                        End If

                        ' Auto-enable tooling when switching TO a tooling-capable model
                        ' if sources are already selected (or were previously).
                        If supportsTooling Then
                            If Not st.ToolingEnabled Then
                                Dim hasTools As Boolean = (st.SelectedToolNames IsNot Nothing AndAlso st.SelectedToolNames.Count > 0) OrElse
                                                          (_selectedToolsForChat IsNot Nothing AndAlso _selectedToolsForChat.Count > 0)
                                If hasTools Then
                                    st.ToolingEnabled = True
                                    _chatToolingEnabled = True
                                End If
                            End If

                            ' IMPORTANT:
                            ' Do NOT auto-enable AgentModeEnabled ("Use all tools") here.
                            ' The checkbox must remain a deliberate user choice.
                            ' When it is unchecked, Local Chat should use only the selected
                            ' Sources from the normal tooling pipeline.
                        End If

                        SaveInkyState(st)

                        Return JsonOk(New With {
                                .ok = True,
                                .supportsFiles = st.SupportsFileUploads,
                                .activeChat = activeChatId,
                                .supportsTooling = supportsTooling,
                                .toolingEnabled = st.ToolingEnabled,
                                .advancedToolsEnabled = st.AgentModeEnabled,
                                .agentFiles = GetAgentFileListForBrowser(),
                                .agentWorkspace = GetAgentWorkspaceForBrowser(),
                                .agentModelActive = st.AgentModelActive,
                                .agentModelAvailable = IsAgentDefaultModelAvailable()
                            })

                    Case "inky_toggletheme"
                        ' (Theme toggle)
                        ' ------------------------------------------------------------------
                        Dim st As InkyState = LoadInkyState()
                        st.DarkMode = Not st.DarkMode
                        SaveInkyState(st)
                        Try
                            My.Settings.Inky_DarkMode = st.DarkMode
                            My.Settings.Save()
                        Catch
                        End Try
                        Return JsonOk(New With {.ok = True, .darkMode = st.DarkMode, .activeChat = activeChatId})

                    Case Else
                        Return JsonErr("Unknown command.")
                End Select

            Catch ex As System.Exception
                If INI_APIDebug Then
                    AppendInkyServerLog(
                        "ProcessRequestInAddIn failed. rawUrl=" &
                        If(rawUrl, "") &
                        "; bodyLen=" &
                        If(body, "").Length.ToString(System.Globalization.CultureInfo.InvariantCulture),
                        ex)
                End If

                Return JsonErr("Bad request: " & ex.Message)
            End Try
        End If

        ' ---- FALLBACK to existing command dispatcher (legacy) ----
        Dim j0 = Newtonsoft.Json.Linq.JObject.Parse(If(body, "{}"))
        Dim cmd0 = j0("Command")?.ToString()
        Dim textBody0 = j0("Text")?.ToString()
        Dim sourceUrl = j0("URL")?.ToString()

        Select Case cmd0
            Case "redink_sendtooutlook"

                SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Browser_SentToOutlook invoked")

                ' (Legacy Outlook insert)
                If String.IsNullOrWhiteSpace(textBody0) Then Return ""
                ' All Outlook automation on UI thread
                Await SwitchToUi(Sub()
                                     Dim olApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
                                     Dim insp As Microsoft.Office.Interop.Outlook.Inspector = ComRetry(Function() olApp.ActiveInspector())
                                     If insp Is Nothing Then Exit Sub
                                     ' Guard CurrentItem (never access inline)
                                     Dim curr As Object = Nothing
                                     Try
                                         curr = ComRetry(Function() insp.CurrentItem)
                                     Catch
                                         curr = Nothing
                                     End Try
                                     If curr Is Nothing OrElse Not TypeOf curr Is Microsoft.Office.Interop.Outlook.MailItem Then
                                         Exit Sub
                                     End If
                                     Dim mail As Microsoft.Office.Interop.Outlook.MailItem =
                                         CType(curr, Microsoft.Office.Interop.Outlook.MailItem)
                                     ' Guard Sent property
                                     Try
                                         If ComRetry(Function() mail.Sent) Then
                                             If insp IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(insp) : insp = Nothing
                                             Exit Sub
                                         End If
                                     Catch
                                         If insp IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(insp) : insp = Nothing
                                         Exit Sub
                                     End Try
                                     ' Guard WordEditor and selection
                                     Dim doc As Microsoft.Office.Interop.Word.Document = Nothing
                                     Try
                                         doc = ComRetry(Function() CType(insp.WordEditor, Microsoft.Office.Interop.Word.Document))
                                         If doc Is Nothing Then Exit Sub
                                         Dim rng As Microsoft.Office.Interop.Word.Range = Nothing
                                         Try
                                             rng = doc.Application.Selection.Range
                                             doc.Application.ScreenUpdating = False
                                             rng.Text = textBody0 & " (" & sourceUrl & ")"
                                             doc.Application.ScreenUpdating = True
                                         Finally
                                             If rng IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(rng) : rng = Nothing
                                         End Try
                                     Finally
                                         If doc IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(doc) : doc = Nothing
                                         If insp IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(insp) : insp = Nothing
                                     End Try
                                 End Sub)
                Return ""

            Case "redink_translate"

                SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Browser_Translate invoked")

                ' ─── 1  guard clauses ─────────────────────────────────────────
                If String.IsNullOrWhiteSpace(textBody0) Then Return ""
                Dim targetLang As String = Await SwitchToUi(Function()
                                                                Return Global.SharedLibrary.SharedLibrary.SharedMethods.PromptForTargetLanguage(
                                                                    "Enter your target language:",
                                                                    AN & " Translate (for Browser)",
                                                                    INI_Language1,
                                                                    _context)
                                                            End Function)
                If String.IsNullOrWhiteSpace(targetLang) OrElse targetLang = "ESC" Then
                    Return ""
                End If
                TranslateLanguage = targetLang.Trim()
                ' ─── 2  call the LLM on the UI thread, get Task(Of String) ─────
                Dim llmOut As String = Await RunLlmAsync(
                    InterpolateAtRuntime(SP_Translate),
                    $"<TEXTTOPROCESS>{textBody0}</TEXTTOPROCESS>")
                ' ─── 3  clean up the wrapper tags / markdown ──────────────────
                llmOut = llmOut.Replace("<TEXTTOPROCESS>", "") _
                       .Replace("</TEXTTOPROCESS>", "") _
                       .Replace("**", "").Trim()
                If llmOut = "" Then Return ""
                ' Optional: copy to clipboard so the user can paste manually
                Await SwitchToUi(Sub() SLib.PutInClipboard(llmOut))
                ' ─── 4  SEND the translation back to the caller ───────────────
                Return llmOut

            Case "redink_correct"

                SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Browser_Correct invoked")

                If String.IsNullOrWhiteSpace(textBody0) Then Return ""
                ' 1)  Run the LLM on the UI thread
                Dim llmOut As String = Await RunLlmAsync(
                    InterpolateAtRuntime(SP_Correct),
                    $"<TEXTTOPROCESS>{textBody0}</TEXTTOPROCESS>")
                llmOut = llmOut.Replace("<TEXTTOPROCESS>", "").Replace("</TEXTTOPROCESS>", "")
                If llmOut = "" Then Return ""
                ' 2)  Show the compare / preview window (synchronous)
                Await SwitchToUi(Sub()
                                     CompareAndInsertText(textBody0, llmOut, True)
                                 End Sub)
                ' 3)  
                Dim accepted As Boolean = Await WaitForPreviewDecisionAsync()
                If Not accepted Then Return ""
                Return llmOut

            Case "redink_freestyle"

                SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Browser_Freestyle invoked")

                ' (Freestyle flow preserved)
                '─── A  gather prompt on UI thread ──────────────────────────────
                ' ---------------------------------------------------------------
                Dim noText As Boolean = String.IsNullOrWhiteSpace(textBody0)
                Dim promptCaption As String = AN & " Freestyle (for Browser)"
                Dim wordInstalled As Boolean = False
                Try
                    Dim wordApp As Object = CreateObject("Word.Application")
                    wordInstalled = True
                    wordApp.Quit()
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp)
                Catch ex As System.Exception
                    wordInstalled = False
                End Try
                Dim sb As New System.Text.StringBuilder()
                If noText Then
                    sb.Append("Please provide the prompt you wish to execute ")
                Else
                    sb.Append("Please provide the prompt you wish to execute using the selected text ")
                End If
                sb.Append("('" & MarkupPrefix & "' for markups, '" & InsertPrefix & "' for direct insert" & If(wordInstalled, " and '" & NewDocPrefix & "' to put the output in a new Word document)", ")"))
                If INI_PromptLib Then sb.Append(" or press 'OK' for the prompt library")
                If INI_SecondAPI Then sb.Append($"; add '{SecondAPICode}' to use {If(String.IsNullOrWhiteSpace(INI_AlternateModelPath), $"the secondary model ({INI_Model_2})", "one of the other models")}")
                If Not String.IsNullOrWhiteSpace(My.Settings.LastPrompt) Then sb.Append("; ctrl-p for your last prompt")
                sb.Append(":")
                Dim promptMsg As String = sb.ToString()
                Dim OptionalButtons As System.Tuple(Of String, String, String)()
                If wordInstalled Then
                    OptionalButtons = {
                        System.Tuple.Create("OK, do a new doc", $"Use this to automatically insert '{NewDocPrefix}' as a prefix.", NewDocPrefix)
                    }
                End If
                OtherPrompt = Await SwitchToUi(Function()
                                                   Return SLib.ShowCustomInputBox(promptMsg, promptCaption, False, "", My.Settings.LastPrompt, If(wordInstalled, OptionalButtons, Nothing), Context:=_context)
                                               End Function)
                Dim doMarkupFlag As Boolean = False
                Dim doInsertFlag As Boolean = False
                Dim UseSecondAPI As Boolean = False
                Dim DoNewDoc As Boolean = False
                If String.IsNullOrEmpty(OtherPrompt) AndAlso OtherPrompt <> "ESC" AndAlso INI_PromptLib Then
                    Dim sel = Await SwitchToUi(Function()
                                                   Return ShowPromptSelector(INI_PromptLibPath, INI_PromptLibPathLocal, Not noText, Nothing)
                                               End Function)
                    OtherPrompt = sel.Item1
                    doMarkupFlag = sel.Item2
                    doInsertFlag = Not sel.Item4
                End If
                If String.IsNullOrWhiteSpace(OtherPrompt) OrElse OtherPrompt = "ESC" Then
                    Return ""
                End If
                My.Settings.LastPrompt = OtherPrompt
                My.Settings.Save()
                If OtherPrompt.StartsWith(InsertPrefix, StringComparison.OrdinalIgnoreCase) Then
                    OtherPrompt = OtherPrompt.Substring(InsertPrefix.Length).Trim()
                    doInsertFlag = True
                ElseIf OtherPrompt.StartsWith(MarkupPrefix, StringComparison.OrdinalIgnoreCase) AndAlso Not noText Then
                    OtherPrompt = OtherPrompt.Substring(MarkupPrefix.Length).Trim()
                    doMarkupFlag = True
                    doInsertFlag = True
                ElseIf OtherPrompt.StartsWith(NewDocPrefix, StringComparison.OrdinalIgnoreCase) AndAlso Not noText Then
                    OtherPrompt = OtherPrompt.Substring(NewDocPrefix.Length).Trim()
                    DoNewDoc = True
                    doMarkupFlag = False
                End If
                If INI_SecondAPI AndAlso OtherPrompt.Contains(SecondAPICode) Then
                    UseSecondAPI = True
                    OtherPrompt = OtherPrompt.Replace(SecondAPICode, "").Trim()
                    If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                        Dim sel = Await SwitchToUi(Function()
                                                       Return Not ShowModelSelection(_context, INI_AlternateModelPath)
                                                   End Function)
                        If sel Then
                            originalConfigLoaded = False
                            Return ""
                        End If
                    End If
                End If
                '─── B  call the LLM on UI thread (async) ──────────────────────
                Dim llmResult As String
                If noText Then
                    llmResult = Await RunLlmAsync(InterpolateAtRuntime(SP_FreestyleNoText), "", UseSecondAPI)
                Else
                    llmResult = Await RunLlmAsync(InterpolateAtRuntime(SP_FreestyleText), $"<TEXTTOPROCESS>{textBody0}</TEXTTOPROCESS>", UseSecondAPI)
                End If
                llmResult = llmResult.Replace("<TEXTTOPROCESS>", "") _
                             .Replace("</TEXTTOPROCESS>", "") _
                             .Trim()
                If String.IsNullOrEmpty(llmResult) Then Return ""
                '─── C  present / insert / clipboard exactly like old code ─────
                ' A) markup path (implies insert)  -----------------------------
                If doMarkupFlag Then
                    Await SwitchToUi(Sub()
                                         CompareAndInsertText(textBody0, llmResult, True)
                                     End Sub)
                    Dim acceptedFs As Boolean = Await WaitForPreviewDecisionAsync()
                    If Not acceptedFs Then Return ""
                    Return llmResult
                End If
                ' B) plain insert path  ----------------------------------------
                If doInsertFlag Then
                    Return llmResult
                End If
                If DoNewDoc And wordInstalled Then
                    If Await TryCreateWordDocFromMarkdown(llmResult) Then
                        Return ""
                    End If
                    Await SwitchToUi(Sub()
                                         ShowCustomMessageBox("Could not create new Word document and insert the LLM output; providing your output to a separate window.")
                                     End Sub)
                End If
                ' C) clipboard-only path  --------------------------------------
                Dim finalTxt As String = Await SwitchToUi(Function()
                                                              Return SLib.ShowCustomWindow(
                                                                  "The LLM has provided the following result (you can edit it):",
                                                                  llmResult,
                                                                  "You can choose whether you want to have the original text put into the clipboard or your text with any changes you have made (without formatting)." & If(wordInstalled, " If you choose to insert the original text with formatting, a new Word document will be created with it. ", " ") & "If you select Cancel, nothing will be put into the clipboard (you can yourself copy it to the clipboard).",
                                                                  AN, False, True, If(wordInstalled, True, False))
                                                          End Function)
                If String.Equals(finalTxt, "Markdown", StringComparison.OrdinalIgnoreCase) Then
                    If wordInstalled AndAlso Await TryCreateWordDocFromMarkdown(llmResult) Then
                        Return ""
                    Else
                        Await SwitchToUi(Sub()
                                             ShowCustomMessageBox("Could not create new Word document and insert the LLM output (however, it will be in the clipboard).")
                                             finalTxt = MarkdownToRtfConverter.Convert(llmResult)
                                         End Sub)
                    End If
                End If
                If Not String.IsNullOrWhiteSpace(finalTxt) Then
                    Await SwitchToUi(Sub() SLib.PutInClipboard(finalTxt))
                End If

                Dim plainText As String

                ' Convert RTF to plain text using RichTextBox
                Using rtb As New RichTextBox()
                    rtb.Rtf = finalTxt
                    plainText = rtb.Text
                End Using

                Return plainText
        End Select

        Return ""
    End Function

    ''' <summary>
    ''' Creates a new Word document, inserts markdown-converted content, shows user message. Returns True on success.
    ''' </summary>
    Private Async Function TryCreateWordDocFromMarkdown(markdown As String) As Task(Of Boolean)
        Try
            Dim wordApp As New Microsoft.Office.Interop.Word.Application()
            wordApp.Visible = True
            Dim newDoc As Microsoft.Office.Interop.Word.Document = wordApp.Documents.Add()
            Dim docSelection As Microsoft.Office.Interop.Word.Selection = wordApp.Selection
            InsertTextWithMarkdown(docSelection, markdown, True)
            Await SwitchToUi(Sub()
                                 ShowCustomMessageBox("Your Word document has been created. It may be hidden behind the other windows.")
                             End Sub)
            Return True
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Builds selectable model list (primary, secondary, alternatives) based on current configuration and persisted state.
    ''' </summary>
    Private Async Function GetModelListForBrowserAsync(ByVal st As InkyState) _
        As System.Threading.Tasks.Task(Of System.Collections.Generic.List(Of Object))

        Dim list As New System.Collections.Generic.List(Of Object)()

        ' 0) If config is not loaded or primary name is blank, try to (re)load once
        If (Not INIloaded) OrElse (String.IsNullOrWhiteSpace(INI_Model) AndAlso Not INI_SecondAPI) Then
            Try
                InitializeConfig(False, False)
            Catch
            End Try
        End If

        ' NOTE: Do NOT reconcile from My.Settings here - the passed-in st already has
        ' the per-chat state. Overwriting it would break independent per-chat model selection.

        ' 1) Availability
        Dim hasPrimary As Boolean = Not String.IsNullOrWhiteSpace(INI_Model)
        Dim hasSecondApi As Boolean = INI_SecondAPI
        Dim hasSecondModelName As Boolean = Not String.IsNullOrWhiteSpace(INI_Model_2)
        Dim hasSecondary As Boolean = hasSecondApi AndAlso hasSecondModelName

        ' Alt list
        Dim alts As System.Collections.Generic.List(Of ModelConfig) = Nothing
        Dim altCount As Integer = 0
        Try
            If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                alts = LoadAlternativeModels(INI_AlternateModelPath, _context)
                If alts IsNot Nothing Then altCount = alts.Count
            End If
        Catch
            altCount = 0
            alts = Nothing
        End Try

        ' 2) Normalize stale alternate selection EVEN if alts could not be loaded
        If st.UseSecondApi AndAlso Not String.IsNullOrWhiteSpace(st.SelectedModelKey) Then
            Dim exists As Boolean = False
            If alts IsNot Nothing Then
                exists = alts.Any(Function(m)
                                      If m Is Nothing Then Return False
                                      Dim label = If(Not String.IsNullOrWhiteSpace(m.ModelDescription), m.ModelDescription, m.Model)
                                      Return String.Equals(label, st.SelectedModelKey, StringComparison.OrdinalIgnoreCase)
                                  End Function)
            Else
                exists = False
            End If
            If Not exists Then
                st.SelectedModelKey = ""
                SaveInkyState(st)
            End If
        End If

        ' 3) Only primary
        If hasPrimary AndAlso Not hasSecondary AndAlso altCount = 0 Then
            list.Add(New With {
                .key = "default",
                .label = INI_Model,
                .selected = (Not st.UseSecondApi),
                .disabled = False,
                .isSeparator = False
            })
            Return list
        End If

        ' 4) Build list
        If hasPrimary Then
            list.Add(New With {.key = "__hdr_primary__", .label = "Primary model:", .selected = False, .disabled = True, .isSeparator = False})
            list.Add(New With {
                .key = "default",
                .label = INI_Model,
                .selected = (Not st.UseSecondApi),
                .disabled = False,
                .isSeparator = False
            })
        End If

        If hasSecondary Then
            list.Add(New With {.key = "__hdr_secondary__", .label = "Secondary model:", .selected = False, .disabled = True, .isSeparator = False})
            list.Add(New With {
                .key = "__second__",
                .label = INI_Model_2,
                .selected = (st.UseSecondApi AndAlso String.IsNullOrWhiteSpace(st.SelectedModelKey)),
                .disabled = False,
                .isSeparator = False
            })
        End If

        If altCount > 0 AndAlso alts IsNot Nothing Then
            list.Add(New With {.key = "__sep__", .label = "Alternative models:", .selected = False, .disabled = True, .isSeparator = True})
            For Each m In alts
                If m Is Nothing Then Continue For
                Dim label = If(Not String.IsNullOrWhiteSpace(m.ModelDescription), m.ModelDescription, m.Model)
                If String.IsNullOrWhiteSpace(label) Then label = "Model"
                list.Add(New With {
                    .key = label,
                    .label = label,
                    .selected = (st.UseSecondApi AndAlso String.Equals(st.SelectedModelKey, label, StringComparison.OrdinalIgnoreCase)),
                    .disabled = False,
                    .isSeparator = False
                })
            Next
        End If

        ' 5) Never return an empty list → synthesize a safe primary entry
        If list.Count = 0 Then
            Dim lbl = If(String.IsNullOrWhiteSpace(INI_Model), "Primary model (not configured)", INI_Model)
            list.Add(New With {
                .key = "default",
                .label = lbl,
                .selected = True,
                .disabled = String.IsNullOrWhiteSpace(INI_Model),
                .isSeparator = False
            })
        End If

        Return list
    End Function

    Private Function RemoveGeneratedOutputFilesSections(ByVal markdown As String) As String
        If String.IsNullOrWhiteSpace(markdown) Then Return If(markdown, String.Empty)

        Dim pattern As String =
        "(?ims)^\s*(?:\*\*)?Output files\s*\(\d+\)\s*:?(?:\*\*)?\s*\r?\n(?:\s*[-*]?\s*[^\r\n]+\r?\n?)+"

        Dim cleaned As String = System.Text.RegularExpressions.Regex.Replace(markdown, pattern, "").TrimEnd()
        Return cleaned
    End Function

    Private Function BuildOutputFilesMarkdown(ByVal outputFiles As List(Of String)) As String
        If outputFiles Is Nothing OrElse outputFiles.Count = 0 Then Return String.Empty

        Dim distinctFiles = outputFiles.
        Where(Function(f) Not String.IsNullOrWhiteSpace(f)).
        Select(Function(f) Path.GetFileName(f)).
        Where(Function(f) Not String.IsNullOrWhiteSpace(f)).
        Distinct(StringComparer.OrdinalIgnoreCase).
        ToList()

        If distinctFiles.Count = 0 Then Return String.Empty

        Dim fileList = String.Join(vbCrLf, distinctFiles.Select(Function(f) "- " & f))
        Return "**Output files (" & distinctFiles.Count.ToString() & "):**" & vbCrLf & fileList
    End Function


    ''' <summary>
    ''' Sanitizes raw model output by removing role markers and normalizing excessive blank lines.
    ''' </summary>
    Private Function SanitizeModelOutputForBrowser(ByVal raw As System.String) As System.String
        If raw Is Nothing Then Return System.String.Empty

        Dim s As System.String = raw

        ' 1) Remove full role-only lines (e.g. "[ASSISTANT]" or "[USER]:")
        s = System.Text.RegularExpressions.Regex.Replace(
            s,
            "(?im)^\s*\[(?:assistant|user)\]\s*:?\s*$",
            "",
            System.Text.RegularExpressions.RegexOptions.None)

        ' 2) Remove role markers at line start (preserve text)
        s = System.Text.RegularExpressions.Regex.Replace(
            s,
            "(?im)^\s*\[(?:assistant|user)\]\s*",
            "",
            System.Text.RegularExpressions.RegexOptions.None)

        ' 3) Remove alternative "ASSISTANT:" / "USER:" prefixes at line start
        s = System.Text.RegularExpressions.Regex.Replace(
            s,
            "(?im)^\s*(?:assistant|user)\s*:\s*",
            "",
            System.Text.RegularExpressions.RegexOptions.None)

        ' 4) Normalize excessive blank lines
        s = System.Text.RegularExpressions.Regex.Replace(
            s,
            "(\r?\n){3,}",
            System.Environment.NewLine & System.Environment.NewLine,
            System.Text.RegularExpressions.RegexOptions.None)

        Return s
    End Function

    ''' <summary>
    ''' Maps internal ChatTurn list to browser DTO objects (camelCase).
    ''' </summary>
    Private Function ToBrowserTurns(list As System.Collections.Generic.List(Of ChatTurn)) _
        As System.Collections.Generic.List(Of Object)

        Dim out As New System.Collections.Generic.List(Of Object)()
        For Each t In list
            out.Add(New With {
                .role = t.Role,
                .markdown = t.Markdown,
                .html = t.Html,
                .utc = t.Utc
            })
        Next
        Return out
    End Function

    ''' <summary>
    ''' Invokes user32.dll to obtain the current process GDI (uiFlags=0) or USER (uiFlags=1) handle count.
    ''' </summary>
    ''' <param name="hProcess">Handle of the process being queried.</param>
    ''' <param name="uiFlags">0 for GDI handles, 1 for USER handles.</param>
    ''' <returns>The handle count reported by the OS for the specified resource type.</returns>

    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function GetGuiResources(hProcess As System.IntPtr, uiFlags As System.Int32) As System.UInt32
    End Function

    ''' <summary>
    ''' Returns current process GDI handle count.
    ''' </summary>
    Private Shared Function GetGdiCount() As System.UInt32
        Return GetGuiResources(System.Diagnostics.Process.GetCurrentProcess().Handle, 0UI)
    End Function

    ''' <summary>
    ''' Returns current process USER handle count.
    ''' </summary>
    Private Shared Function GetUserCount() As System.UInt32
        Return GetGuiResources(System.Diagnostics.Process.GetCurrentProcess().Handle, 1UI)
    End Function


    ' Opt-out migration is handled by renaming the persisted fields (old "Selected..." values
    ' simply do not bind to the new "Deselected..." fields), so no in-place migration is needed.
    ' Retained as a no-op to keep existing call sites stable.
    Private Sub MigrateLocalChatToolSelections(st As InkyState,
                                           Optional includeInteractiveM365Tools As Boolean = True)
        ' Intentionally left blank.
    End Sub

    Private Function GetLocalChatEffectiveSelection(st As InkyState,
                                                    Optional includeInteractiveM365Tools As Boolean = True) As List(Of ModelConfig)
        If st Is Nothing Then
            Return New List(Of ModelConfig)()
        End If

        ' Opt-out defaults: resolve selected = available-for-state minus the user's deselections.
        Dim effective =
            GetLocalChatEffectiveTools(
                ResolveLocalChatSelectedMainToolNames(st, includeInteractiveM365Tools),
                ResolveLocalChatSelectedAdvancedToolNames(st, includeInteractiveM365Tools),
                st.AgentModeEnabled,
                includeInteractiveM365Tools)

        st.SelectedToolNames = effective.
            Select(Function(t) t.ToolName).
            Distinct(StringComparer.OrdinalIgnoreCase).
            ToList()

        Return effective
    End Function

    Private Function HasLocalChatAnyCallableTools(st As InkyState,
                                                  Optional includeInteractiveM365Tools As Boolean = True) As Boolean
        Dim effective = GetLocalChatEffectiveSelection(st, includeInteractiveM365Tools)
        If effective.Count > 0 Then Return True
        Return IsChatAgentWorkspaceConnected()
    End Function

    Private Function GetEffectiveToolingEnabled() As Boolean
        Try
            Dim st = LoadInkyState()
            Dim enabled As Boolean = (st IsNot Nothing AndAlso st.ToolingEnabled)
            _chatToolingEnabled = enabled
            Return enabled
        Catch
            Return _chatToolingEnabled
        End Try
    End Function

End Class
