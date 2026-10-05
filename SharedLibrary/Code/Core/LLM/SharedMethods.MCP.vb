' Part of "SharedLibrary"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: SharedMethods.MCP.vb
' Purpose: Provides shared runtime helpers for executing MCP tool calls over
'          SSE transport, including session endpoint discovery, initialize
'          handshake, request dispatch, and SSE response collection.
'
' Contains:
'  - Transport constants:
'      - `MCP_SSE_PREFIX`
'      - `MCP_PROTOCOL_VERSION`
'  - Session bootstrap:
'      - `AcquireMCPSSESessionEndpoint` opens the SSE stream, reads the
'        advertised POST endpoint, and performs the MCP initialize handshake.
'  - Tool execution:
'      - `ExecuteMCPSSEToolCall` performs a complete SSE-based MCP round-trip.
'  - SSE helpers:
'      - Reads endpoint and response events from the event stream.
'      - Resolves relative SSE endpoints and validates transport responses.
'
' Notes:
'  - These helpers bypass `LLM(...)` because SSE transport returns the effective
'    tool result on the GET stream rather than in the POST response body.
' =============================================================================

Option Explicit On
Option Strict On

Imports System.IO
Imports System.Net.Http
Imports System.Text
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary

    Partial Public Class SharedMethods

        ''' <summary>
        ''' Marker prefix for INI Endpoint values that require an SSE session handshake.
        ''' When the Endpoint starts with "sse:", the caller must strip this prefix and
        ''' call <see cref="AcquireMCPSSESessionEndpoint"/> to obtain the actual POST URL
        ''' before passing it to LLM().
        ''' </summary>
        Public Const MCP_SSE_PREFIX As String = "sse:"

        ''' <summary>MCP protocol version used in the initialize handshake.</summary>
        Public Const MCP_PROTOCOL_VERSION As String = "2025-03-26"

        ''' <summary>
        ''' Performs an SSE handshake against the given base URL:
        '''   1. GET the SSE URL with Accept: text/event-stream
        '''   2. Read SSE events until "event: endpoint" + "data: ..." arrives
        '''   3. Resolve the relative endpoint to an absolute URL
        '''   4. POST the MCP "initialize" + "notifications/initialized" handshake
        '''   5. Return the session POST endpoint (e.g. https://server/messages?sessionId=abc)
        '''
        ''' The returned URL is ephemeral — it is valid only for the current session.
        ''' </summary>
        ''' <param name="sseBaseUrl">The base SSE URL (e.g. https://server/sse).</param>
        ''' <param name="headerA">Auth header name (e.g. "Authorization"), or empty.</param>
        ''' <param name="headerB">Auth header value with {apikey} already resolved, or empty.</param>
        ''' <param name="timeoutMs">Timeout in milliseconds for the handshake.</param>
        ''' <returns>The fully-qualified session POST endpoint URL.</returns>
        ''' <exception cref="Exception">Thrown when the handshake fails or times out.</exception>
        Public Shared Async Function AcquireMCPSSESessionEndpoint(
                sseBaseUrl As String,
                headerA As String,
                headerB As String,
                Optional timeoutMs As Integer = 30000) As Task(Of String)

            EnsureTls12()

            Dim handler As New HttpClientHandler()
            ConfigureMCPHandler(handler)

            Dim client As New HttpClient(handler)
            client.Timeout = TimeSpan.FromMilliseconds(timeoutMs)
            client.DefaultRequestHeaders.Accept.Clear()
            client.DefaultRequestHeaders.Accept.Add(
                New Headers.MediaTypeWithQualityHeaderValue("text/event-stream"))

            If Not String.IsNullOrWhiteSpace(headerA) AndAlso Not String.IsNullOrWhiteSpace(headerB) Then
                client.DefaultRequestHeaders.TryAddWithoutValidation(headerA, headerB)
            End If

            Dim sseResponse As HttpResponseMessage = Nothing
            Dim reader As StreamReader = Nothing
            Dim postEndpoint As String = ""

            Try
                ' ── Step 1: GET the SSE stream ───────────────────────────────
                Dim request As New HttpRequestMessage(HttpMethod.Get, sseBaseUrl)
                sseResponse = Await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead).ConfigureAwait(False)

                If Not sseResponse.IsSuccessStatusCode Then
                    Dim preview = Await sseResponse.Content.ReadAsStringAsync().ConfigureAwait(False)
                    Throw New Exception(
                        $"SSE handshake GET {CInt(sseResponse.StatusCode)} ({sseResponse.ReasonPhrase}): " &
                        If(preview.Length > 300, preview.Substring(0, 300), preview))
                End If

                Dim stream = Await sseResponse.Content.ReadAsStreamAsync().ConfigureAwait(False)
                reader = New StreamReader(stream, Encoding.UTF8)

                ' ── Step 2: Read until "event: endpoint" ─────────────────────
                postEndpoint = Await ReadSSEEndpointEvent(reader, timeoutMs).ConfigureAwait(False)

                If String.IsNullOrWhiteSpace(postEndpoint) Then
                    Throw New Exception($"SSE stream at {sseBaseUrl} did not return an endpoint event within {timeoutMs}ms.")
                End If

                ' ── Step 3: Resolve relative URL ─────────────────────────────
                postEndpoint = ResolveSSEEndpoint(sseBaseUrl, postEndpoint)

                ' ── Step 4: POST initialize handshake ────────────────────────
                Await PostMCPInitializeHandshake(postEndpoint, headerA, headerB, timeoutMs).ConfigureAwait(False)

                Return postEndpoint

            Finally
                Try : reader?.Dispose() : Catch : End Try
                Try : sseResponse?.Dispose() : Catch : End Try
                Try : client.Dispose() : Catch : End Try
            End Try
        End Function


        Private Shared Sub WriteMCPDebugFile(fileName As String, content As String)
            Try
                Dim desktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
                Dim debugFilePath As String = Path.Combine(desktopPath, fileName)
                File.WriteAllText(debugFilePath, RedactSensitiveDebugText(If(content, "")))
            Catch
            End Try
        End Sub

        ''' <summary>
        ''' Performs a complete MCP tool call over SSE transport:
        '''   1. GET the SSE base URL → read "event: endpoint" → obtain session POST URL
        '''   2. POST "initialize" + "notifications/initialized" handshake
        '''   3. POST the actual JSON-RPC request (the tool call)
        '''   4. Read the JSON-RPC response from the SSE GET stream
        '''   5. Return the raw JSON response string
        '''
        ''' This bypasses LLM() entirely because SSE transport returns the real response
        ''' on the GET stream, not in the POST response body (which is just "Accepted").
        ''' </summary>
        ''' <param name="sseBaseUrl">The base SSE URL (e.g. https://server/sse).</param>
        ''' <param name="requestBody">The JSON-RPC request body to send (tools/call payload).</param>
        ''' <param name="headerA">Auth header name (e.g. "Authorization"), or empty.</param>
        ''' <param name="headerB">Auth header value with {apikey} already resolved, or empty.</param>
        ''' <param name="timeoutMs">Timeout in milliseconds for the entire operation.</param>
        ''' <returns>The raw JSON response string from the MCP server.</returns>
        ''' <exception cref="Exception">Thrown when the handshake, request, or response reading fails.</exception>
        Public Shared Async Function ExecuteMCPSSEToolCall(
                context As ISharedContext,
                sseBaseUrl As String,
                requestBody As String,
                headerA As String,
                headerB As String,
                Optional timeoutMs As Integer = 60000) As Task(Of String)

            EnsureTls12()
            Await Agents.AgentGate.EnterAsync().ConfigureAwait(False)
            Try

                EnsureTls12()

                Dim handler As New HttpClientHandler()
                ConfigureMCPHandler(handler)

                Dim client As New HttpClient(handler)
                client.Timeout = TimeSpan.FromMilliseconds(timeoutMs)
                client.DefaultRequestHeaders.Accept.Clear()
                client.DefaultRequestHeaders.Accept.Add(
                    New Headers.MediaTypeWithQualityHeaderValue("text/event-stream"))

                If Not String.IsNullOrWhiteSpace(headerA) AndAlso Not String.IsNullOrWhiteSpace(headerB) Then
                    client.DefaultRequestHeaders.TryAddWithoutValidation(headerA, headerB)
                End If

                Dim sseResponse As HttpResponseMessage = Nothing
                Dim reader As StreamReader = Nothing

                Try
                    ' ── Step 1: GET the SSE stream → obtain session endpoint ──────
                    Dim request As New HttpRequestMessage(HttpMethod.Get, sseBaseUrl)
                    sseResponse = Await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead).ConfigureAwait(False)

                    If Not sseResponse.IsSuccessStatusCode Then
                        Dim preview = Await sseResponse.Content.ReadAsStringAsync().ConfigureAwait(False)
                        Throw New Exception(
                            $"SSE GET {CInt(sseResponse.StatusCode)} ({sseResponse.ReasonPhrase}): " &
                            If(preview.Length > 300, preview.Substring(0, 300), preview))
                    End If

                    Dim stream = Await sseResponse.Content.ReadAsStreamAsync().ConfigureAwait(False)
                    reader = New StreamReader(stream, Encoding.UTF8)

                    Dim postEndpoint As String = Await ReadSSEEndpointEvent(reader, timeoutMs).ConfigureAwait(False)

                    If String.IsNullOrWhiteSpace(postEndpoint) Then
                        Throw New Exception($"SSE stream at {sseBaseUrl} did not return an endpoint event.")
                    End If

                    postEndpoint = ResolveSSEEndpoint(sseBaseUrl, postEndpoint)

                    If context IsNot Nothing AndAlso context.INI_APIDebug Then
                        Dim sentDebug As New StringBuilder()
                        sentDebug.AppendLine($"SENT TO API (SSE GET): {sseBaseUrl}")
                        sentDebug.AppendLine($"SSE POST ENDPOINT: {postEndpoint}")
                        sentDebug.AppendLine($"HeaderA: {headerA}")
                        sentDebug.AppendLine($"HeaderB: {headerB}")
                        sentDebug.AppendLine("SSE TOOL REQUEST:")
                        sentDebug.AppendLine(requestBody)

                        Debug.WriteLine(sentDebug.ToString())
                        WriteMCPDebugFile("RI_Debug_Sent.json", sentDebug.ToString())
                    End If

                    ' ── Step 2: POST initialize handshake ────────────────────────
                    Using postHandler As New HttpClientHandler()
                        ConfigureMCPHandler(postHandler)

                        Using postClient As New HttpClient(postHandler)
                            postClient.Timeout = TimeSpan.FromMilliseconds(timeoutMs)
                            If Not String.IsNullOrWhiteSpace(headerA) AndAlso Not String.IsNullOrWhiteSpace(headerB) Then
                                postClient.DefaultRequestHeaders.TryAddWithoutValidation(headerA, headerB)
                            End If

                            ' initialize
                            Dim initPayload As New JObject From {
                                {"jsonrpc", "2.0"}, {"id", 1}, {"method", "initialize"},
                                {"params", New JObject From {
                                    {"protocolVersion", MCP_PROTOCOL_VERSION},
                                    {"capabilities", New JObject()},
                                    {"clientInfo", New JObject From {{"name", "RedInk"}, {"version", "1.0"}}}
                                }}
                            }
                            Dim initContent As New StringContent(initPayload.ToString(Formatting.None), Encoding.UTF8, "application/json")
                            Dim initResp = Await postClient.PostAsync(postEndpoint, initContent).ConfigureAwait(False)
                            If Not initResp.IsSuccessStatusCode Then
                                Dim err = Await initResp.Content.ReadAsStringAsync().ConfigureAwait(False)
                                Throw New Exception($"SSE initialize POST {CInt(initResp.StatusCode)}: {If(err.Length > 300, err.Substring(0, 300), err)}")
                            End If

                            ' Wait for initialize response on SSE stream (consume it)
                            Await WaitForSSEResponse(reader, "1", timeoutMs).ConfigureAwait(False)

                            ' notifications/initialized
                            Dim notifPayload As New JObject From {{"jsonrpc", "2.0"}, {"method", "notifications/initialized"}}
                            Dim notifContent As New StringContent(notifPayload.ToString(Formatting.None), Encoding.UTF8, "application/json")
                            Try
                                Await postClient.PostAsync(postEndpoint, notifContent).ConfigureAwait(False)
                            Catch
                            End Try

                            ' ── Step 3: POST the actual tool call ────────────────────
                            Dim toolContent As New StringContent(requestBody, Encoding.UTF8, "application/json")
                            Dim toolResp = Await postClient.PostAsync(postEndpoint, toolContent).ConfigureAwait(False)

                            If Not toolResp.IsSuccessStatusCode Then
                                Dim err = Await toolResp.Content.ReadAsStringAsync().ConfigureAwait(False)
                                Throw New Exception($"SSE tool POST {CInt(toolResp.StatusCode)}: {If(err.Length > 300, err.Substring(0, 300), err)}")
                            End If

                            ' Check if the POST response itself contains a JSON-RPC result
                            Dim postBody = Await toolResp.Content.ReadAsStringAsync().ConfigureAwait(False)
                            If Not String.IsNullOrWhiteSpace(postBody) AndAlso postBody.TrimStart().StartsWith("{") Then
                                Try
                                    Dim directResult = JObject.Parse(postBody)
                                    If directResult("result") IsNot Nothing OrElse directResult("error") IsNot Nothing Then
                                        Dim directResultText As String = directResult.ToString(Formatting.None)

                                        If context IsNot Nothing AndAlso context.INI_APIDebug Then
                                            Debug.WriteLine($"RECEIVED FROM API (SSE):{Environment.NewLine}{directResultText}")
                                            WriteMCPDebugFile("RI_Debug_Received.json", directResultText)
                                        End If

                                        Return directResultText
                                    End If
                                Catch
                                End Try
                            End If
                        End Using
                    End Using

                    ' ── Step 4: Read the tool call response from the SSE stream ──
                    Dim expectedId As String = ""
                    Try
                        Dim reqObj = JObject.Parse(requestBody)
                        expectedId = If(reqObj("id")?.ToString(), "")
                    Catch
                    End Try

                    Dim responseJson = Await WaitForSSEResponse(reader, expectedId, timeoutMs).ConfigureAwait(False)
                    Dim responseJsonText As String = responseJson.ToString(Formatting.None)

                    If context IsNot Nothing AndAlso context.INI_APIDebug Then
                        Debug.WriteLine($"RECEIVED FROM API (SSE):{Environment.NewLine}{responseJsonText}")
                        WriteMCPDebugFile("RI_Debug_Received.json", responseJsonText)
                    End If

                    Return responseJsonText

                Finally
                    Try : reader?.Dispose() : Catch : End Try
                    Try : sseResponse?.Dispose() : Catch : End Try
                    Try : client.Dispose() : Catch : End Try
                End Try
            Finally
                Agents.AgentGate.Release()
            End Try
        End Function

        ''' <summary>
        ''' Reads SSE events from the stream until a JSON-RPC response with a matching id arrives.
        ''' </summary>
        Private Shared Async Function WaitForSSEResponse(reader As StreamReader, expectedId As String, timeoutMs As Integer) As Task(Of JObject)
            Dim currentEvent As String = ""
            Dim deadline = DateTime.UtcNow.AddMilliseconds(timeoutMs)

            While DateTime.UtcNow < deadline
                Dim line = Await ReadLineWithTimeout(reader, Math.Min(timeoutMs, 30000)).ConfigureAwait(False)
                If line Is Nothing Then Exit While

                If line.StartsWith("event:", StringComparison.OrdinalIgnoreCase) Then
                    currentEvent = line.Substring(6).Trim()
                ElseIf line.StartsWith("data:", StringComparison.OrdinalIgnoreCase) Then
                    Dim data = line.Substring(5).Trim()
                    If String.IsNullOrWhiteSpace(data) Then Continue While
                    If currentEvent.Equals("endpoint", StringComparison.OrdinalIgnoreCase) Then Continue While

                    If data.TrimStart().StartsWith("{") Then
                        Try
                            Dim msg = JObject.Parse(data)
                            If Not String.IsNullOrWhiteSpace(expectedId) Then
                                Dim msgId = msg("id")?.ToString()
                                If msgId IsNot Nothing AndAlso msgId = expectedId Then
                                    CheckJsonRpcError(msg)
                                    Return msg
                                End If
                            End If
                            If msg("result") IsNot Nothing OrElse msg("error") IsNot Nothing Then
                                CheckJsonRpcError(msg)
                                Return msg
                            End If
                        Catch
                        End Try
                    End If
                End If
            End While

            Throw New Exception("Timeout waiting for SSE response.")
        End Function

        ' ─────────────────────────────────────────────────────────────────
        '  MCP SSE private helpers
        ' ─────────────────────────────────────────────────────────────────

        ''' <summary>Ensures TLS 1.2 is enabled.</summary>
        Private Shared Sub EnsureTls12()
            If (System.Net.ServicePointManager.SecurityProtocol And System.Net.SecurityProtocolType.Tls12) = 0 Then
                System.Net.ServicePointManager.SecurityProtocol =
                    System.Net.ServicePointManager.SecurityProtocol Or System.Net.SecurityProtocolType.Tls12
            End If
        End Sub

        ''' <summary>Configures an HttpClientHandler with proxy and decompression settings.</summary>
        Private Shared Sub ConfigureMCPHandler(handler As HttpClientHandler)
            handler.AutomaticDecompression = System.Net.DecompressionMethods.GZip Or System.Net.DecompressionMethods.Deflate
            handler.UseProxy = True
            handler.Proxy = System.Net.WebRequest.DefaultWebProxy
            If handler.Proxy IsNot Nothing Then
                handler.Proxy.Credentials = System.Net.CredentialCache.DefaultCredentials
            End If
        End Sub

        ''' <summary>Reads SSE events until an "endpoint" event arrives; returns the data value.</summary>
        Private Shared Async Function ReadSSEEndpointEvent(reader As StreamReader, timeoutMs As Integer) As Task(Of String)
            Dim currentEvent As String = ""
            Dim deadline = DateTime.UtcNow.AddMilliseconds(timeoutMs)

            While DateTime.UtcNow < deadline
                Dim line = Await ReadLineWithTimeout(reader, Math.Min(timeoutMs, 15000)).ConfigureAwait(False)
                If line Is Nothing Then Exit While

                If line.StartsWith("event:", StringComparison.OrdinalIgnoreCase) Then
                    currentEvent = line.Substring(6).Trim()
                ElseIf line.StartsWith("data:", StringComparison.OrdinalIgnoreCase) Then
                    Dim data = line.Substring(5).Trim()
                    If currentEvent.Equals("endpoint", StringComparison.OrdinalIgnoreCase) AndAlso
                       Not String.IsNullOrWhiteSpace(data) Then
                        Return data
                    End If
                End If
            End While

            Return Nothing
        End Function

        ''' <summary>Resolves a potentially relative SSE endpoint URL against the base SSE URL.</summary>
        Private Shared Function ResolveSSEEndpoint(sseBaseUrl As String, postEndpoint As String) As String
            If postEndpoint.StartsWith("/") Then
                Dim uri As New Uri(sseBaseUrl)
                Return $"{uri.Scheme}://{uri.Authority}{postEndpoint}"
            ElseIf Not postEndpoint.StartsWith("http", StringComparison.OrdinalIgnoreCase) Then
                Return sseBaseUrl.TrimEnd("/"c) & "/" & postEndpoint
            End If
            Return postEndpoint
        End Function

        ''' <summary>POSTs the MCP initialize + notifications/initialized handshake.</summary>
        Private Shared Async Function PostMCPInitializeHandshake(
                postEndpoint As String, headerA As String, headerB As String,
                timeoutMs As Integer) As Task

            Using postHandler As New HttpClientHandler()
                ConfigureMCPHandler(postHandler)

                Using postClient As New HttpClient(postHandler)
                    postClient.Timeout = TimeSpan.FromMilliseconds(timeoutMs)
                    If Not String.IsNullOrWhiteSpace(headerA) AndAlso Not String.IsNullOrWhiteSpace(headerB) Then
                        postClient.DefaultRequestHeaders.TryAddWithoutValidation(headerA, headerB)
                    End If

                    Dim initPayload As New JObject From {
                        {"jsonrpc", "2.0"}, {"id", 1}, {"method", "initialize"},
                        {"params", New JObject From {
                            {"protocolVersion", MCP_PROTOCOL_VERSION},
                            {"capabilities", New JObject()},
                            {"clientInfo", New JObject From {{"name", "RedInk"}, {"version", "1.0"}}}
                        }}
                    }
                    Dim initContent As New StringContent(initPayload.ToString(Formatting.None), Encoding.UTF8, "application/json")
                    Dim initResp = Await postClient.PostAsync(postEndpoint, initContent).ConfigureAwait(False)
                    If Not initResp.IsSuccessStatusCode Then
                        Dim err = Await initResp.Content.ReadAsStringAsync().ConfigureAwait(False)
                        Throw New Exception($"SSE initialize POST {CInt(initResp.StatusCode)}: {If(err.Length > 300, err.Substring(0, 300), err)}")
                    End If

                    Dim notifPayload As New JObject From {{"jsonrpc", "2.0"}, {"method", "notifications/initialized"}}
                    Dim notifContent As New StringContent(notifPayload.ToString(Formatting.None), Encoding.UTF8, "application/json")
                    Try
                        Await postClient.PostAsync(postEndpoint, notifContent).ConfigureAwait(False)
                    Catch
                    End Try
                End Using
            End Using
        End Function

        ''' <summary>
        ''' Validates a JSON-RPC response and throws if it contains an error object.
        ''' </summary>
        Private Shared Sub CheckJsonRpcError(result As JObject)
            Dim err = result("error")
            If err IsNot Nothing Then
                Throw New Exception($"JSON-RPC error {If(err("code")?.ToString(), "")}: {If(err("message")?.ToString(), "Unknown error")}")
            End If
        End Sub

        ''' <summary>Reads a single line from a StreamReader with a timeout.</summary>
        Private Shared Async Function ReadLineWithTimeout(reader As StreamReader, timeoutMs As Integer) As Task(Of String)
            Using cts As New Threading.CancellationTokenSource(timeoutMs)
                Try
                    Dim lineTask = reader.ReadLineAsync()
                    Dim completed = Await Task.WhenAny(lineTask, Task.Delay(timeoutMs, cts.Token)).ConfigureAwait(False)
                    If completed Is lineTask Then
                        cts.Cancel()
                        Return Await lineTask
                    End If
                    Return Nothing
                Catch ex As OperationCanceledException
                    Return Nothing
                Catch
                    Return Nothing
                End Try
            End Using
        End Function

        ''' <summary>
        ''' Marker prefix for INI Endpoint values that require MCP Streamable HTTP
        ''' initialization before the actual tools/call request is sent.
        ''' </summary>
        Public Const MCP_STREAMABLE_PREFIX As String = "mcp:"

        ''' <summary>
        ''' Performs a complete MCP tool call over Streamable HTTP transport:
        '''   1. POST "initialize"
        '''   2. POST "notifications/initialized"
        '''   3. POST the actual JSON-RPC tool request
        '''   4. Return the raw JSON-RPC response string
        ''' </summary>
        Public Shared Async Function ExecuteMCPStreamableToolCall(
                mcpUrl As String,
                requestBody As String,
                headerA As String,
                headerB As String,
                Optional timeoutMs As Integer = 60000) As Task(Of String)

            EnsureTls12()

            Await Agents.AgentGate.EnterAsync().ConfigureAwait(False)
            Try

                Using handler As New HttpClientHandler()
                    ConfigureMCPHandler(handler)

                    Using client As New HttpClient(handler)
                        client.Timeout = TimeSpan.FromMilliseconds(timeoutMs)
                        client.DefaultRequestHeaders.Accept.Clear()
                        client.DefaultRequestHeaders.Accept.Add(
                            New Headers.MediaTypeWithQualityHeaderValue("application/json"))
                        client.DefaultRequestHeaders.Accept.Add(
                            New Headers.MediaTypeWithQualityHeaderValue("text/event-stream"))

                        If Not String.IsNullOrWhiteSpace(headerA) AndAlso Not String.IsNullOrWhiteSpace(headerB) Then
                            client.DefaultRequestHeaders.TryAddWithoutValidation(headerA, headerB)
                        End If

                        Dim initPayload As New JObject From {
                            {"jsonrpc", "2.0"},
                            {"id", 1},
                            {"method", "initialize"},
                            {"params", New JObject From {
                                {"protocolVersion", MCP_PROTOCOL_VERSION},
                                {"capabilities", New JObject()},
                                {"clientInfo", New JObject From {
                                    {"name", "RedInk"},
                                    {"version", "1.0"}
                                }}
                            }}
                        }

                        Await PostMCPStreamableRequest(
                            client,
                            mcpUrl,
                            initPayload,
                            timeoutMs,
                            expectResponse:=True).ConfigureAwait(False)

                        Dim notifPayload As New JObject From {
                            {"jsonrpc", "2.0"},
                            {"method", "notifications/initialized"}
                        }

                        Try
                            Await PostMCPStreamableRequest(
                                client,
                                mcpUrl,
                                notifPayload,
                                timeoutMs,
                                expectResponse:=False).ConfigureAwait(False)
                        Catch
                        End Try

                        Dim toolPayload As JObject = JObject.Parse(requestBody)
                        Dim toolResult As JObject =
                            Await PostMCPStreamableRequest(
                                client,
                                mcpUrl,
                                toolPayload,
                                timeoutMs,
                                expectResponse:=True).ConfigureAwait(False)

                        Return toolResult.ToString(Formatting.None)
                    End Using
                End Using
            Finally
                Agents.AgentGate.Release()
            End Try
        End Function

        ''' <summary>
        ''' Sends a single JSON-RPC POST to a Streamable HTTP MCP endpoint and returns the parsed JSON response.
        ''' </summary>
        Private Shared Async Function PostMCPStreamableRequest(
                client As HttpClient,
                mcpUrl As String,
                payload As JObject,
                timeoutMs As Integer,
                Optional expectResponse As Boolean = True) As Task(Of JObject)

            Dim content As New StringContent(payload.ToString(Formatting.None), Encoding.UTF8, "application/json")
            Dim response As HttpResponseMessage = Await client.PostAsync(mcpUrl, content).ConfigureAwait(False)
            Dim responseText As String = Await response.Content.ReadAsStringAsync().ConfigureAwait(False)

            If Not response.IsSuccessStatusCode Then
                Throw New Exception(
                    $"MCP Streamable HTTP POST {CInt(response.StatusCode)} ({response.ReasonPhrase}): " &
                    If(responseText.Length > 300, responseText.Substring(0, 300), responseText))
            End If

            If Not expectResponse Then
                Return New JObject()
            End If

            responseText = ExtractJsonFromMCPResponse(
                responseText,
                If(response.Content.Headers.ContentType?.MediaType, ""))

            If String.IsNullOrWhiteSpace(responseText) Then
                Return New JObject()
            End If

            Dim result As JObject = JObject.Parse(responseText)
            CheckJsonRpcError(result)
            Return result
        End Function

        ''' <summary>
        ''' Extracts a JSON-RPC message from a body that may be plain JSON or SSE-framed text.
        ''' </summary>
        Private Shared Function ExtractJsonFromMCPResponse(responseText As String, contentType As String) As String
            If String.IsNullOrWhiteSpace(responseText) Then
                Return responseText
            End If

            Dim trimmed As String = responseText.TrimStart()
            Dim mediaType As String = If(contentType, "")

            Dim looksLikeSse As Boolean =
                mediaType.IndexOf("text/event-stream", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                trimmed.StartsWith("data:", StringComparison.OrdinalIgnoreCase) OrElse
                trimmed.StartsWith("event:", StringComparison.OrdinalIgnoreCase) OrElse
                trimmed.StartsWith(":", StringComparison.Ordinal)

            If Not looksLikeSse Then
                Return responseText
            End If

            Dim lastJson As String = Nothing

            For Each rawLine As String In responseText.Split(
                New String() {vbCrLf, vbLf, vbCr},
                StringSplitOptions.None)

                Dim line As String = If(rawLine, "").Trim()

                If line.StartsWith("data:", StringComparison.OrdinalIgnoreCase) Then
                    Dim data As String = line.Substring(5).Trim()

                    If data.Length > 0 AndAlso
                       Not data.Equals("[DONE]", StringComparison.OrdinalIgnoreCase) AndAlso
                       data.TrimStart().StartsWith("{") Then

                        Try
                            Dim candidate As JObject = JObject.Parse(data)
                            If candidate("result") IsNot Nothing OrElse candidate("error") IsNot Nothing Then
                                lastJson = data
                            End If
                        Catch
                        End Try
                    End If
                End If
            Next

            Return If(lastJson, responseText)
        End Function


    End Class

End Namespace