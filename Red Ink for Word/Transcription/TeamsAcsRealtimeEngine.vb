' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved.
'
' =============================================================================
' File: TeamsAcsRealtimeEngine.vb
' Purpose: Word-side realtime Teams/ACS transcription engine.
'
' Architecture:
'   This class is the Word/VSTO-side client for realtime Teams transcription.
'   It does not join Teams calls directly and it does not call Azure
'   Communication Services directly. Instead, it connects to a Red Ink bridge
'   WebSocket and receives already-transcribed, speaker-labelled transcript
'   events from that bridge.
'
'   Intended runtime flow:
'
'       Microsoft Word / VSTO Transcriptor
'           |
'           |  local/private WebSocket
'           |  ws://127.0.0.1:<port>/redink/transcriptor/client/{sessionId}
'           v
'       Red Ink Teams/ACS Bridge
'           |
'           |  ACS Call Automation callback endpoint
'           |  https://<public-url>/redink/acs/callbacks
'           |
'           |  ACS realtime transcription WebSocket endpoint
'           |  wss://<public-url>/redink/acs/transcription/{sessionId}
'           v
'       Azure Communication Services / Teams meeting
'
'   ACS must be able to reach the bridge over a public HTTPS/WSS URL.
'   Word may connect to the bridge locally. Therefore, if the bridge is hosted
'   inside the Word process, a public tunnel or reverse proxy is still required
'   for ACS, for example ngrok, Azure Dev Tunnels, Cloudflare Tunnel, Azure App
'   Service reverse proxy, or a customer-controlled public HTTPS endpoint.
'
' What this engine does:
'   - Implements ITranscriptionEngine.
'   - Connects to the bridge WebSocket configured by _bridgeWebSocketUri.
'   - Sends a "start" JSON message containing the Teams meeting join URL and
'     selected locale.
'   - Sends a "stop" JSON message when stopped.
'   - Receives JSON messages from the bridge.
'   - Accepts either simplified Red Ink bridge messages or raw ACS-style
'     TranscriptionMetadata / TranscriptionData messages.
'   - Raises PartialResult and FinalResult.
'   - Sets TranscriptionEventArgs.Speaker by reflection when a speaker value is
'     available.
'
' What this engine intentionally does not do:
'   - It does not capture local audio.
'   - PushAudioAsync is intentionally a no-op.
'   - It does not use the Azure Speech SDK.
'   - It does not use the Azure Communication Services SDK directly.
'   - It does not perform batch or file transcription.
'   - It does not perform speech diarization from mixed audio.
'   - It expects the bridge/ACS layer to provide participant attribution.
'
' Required Transcriptor integration:
'   - Add EngineKind.TeamsAcsRealtime.
'   - Register this engine in LoadEngines().
'   - Populate language choices from TeamsAcsRealtimeEngine.SupportedLanguages.
'   - Create the engine in CreateEngineAsync().
'   - Treat it as a live-only engine.
'   - Do not start AudioCaptureService for this engine.
'
'   The transcriptor should have logic equivalent to:
'
'       Private Shared Function EngineNeedsLocalAudioCapture(kind As EngineKind) As Boolean
'           Select Case kind
'               Case EngineKind.TeamsAcsRealtime
'                   Return False
'               Case Else
'                   Return True
'           End Select
'       End Function
'
'   In OnStart(), create and start AudioCaptureService only if
'   EngineNeedsLocalAudioCapture(d.Kind) returns True.
'
' Expected configuration:
'   - _bridgeWebSocketUri:
'       The WebSocket URI used by this engine to connect to the Red Ink bridge.
'       For an integrated Word-hosted bridge this can be local, for example:
'
'           ws://127.0.0.1:47883/redink/transcriptor/client/default
'
'   - _meetingJoinUrl:
'       The Teams meeting join URL pasted/configured by the user, for example:
'
'           https://teams.microsoft.com/l/meetup-join/...
'
'   - _bridgeBearerToken:
'       Optional bearer token used only to authenticate this engine to the
'       bridge WebSocket. It is not an Azure token.
'
' Suggested INI mapping:
'   - INI_Model_Parameter1 = bridge WebSocket URI
'   - INI_Model_Parameter2 = Teams meeting join URL
'   - INI_Model_Parameter3 = optional encrypted bridge bearer token
'
' Additional bridge configuration, not consumed directly by this class, must
' exist wherever the bridge is implemented:
'   - ACS connection string or equivalent Azure credential.
'   - Public base URL reachable by ACS.
'   - ACS callback path.
'   - ACS transcription WebSocket path.
'
' Bridge responsibilities:
'   The bridge is the missing server-side component. It may be implemented
'   inside the existing Word-hosted web server, as a companion executable, or as
'   a cloud service. If hosted inside Word, it still needs a public HTTPS/WSS
'   tunnel or reverse proxy so ACS can reach it.
'
'   The bridge must expose at least:
'
'       WS   /redink/transcriptor/client/{sessionId}
'       WS   /redink/acs/transcription/{sessionId}
'       POST /redink/acs/callbacks
'
'   /redink/transcriptor/client/{sessionId}
'       Used by this engine.
'       Accepts the engine WebSocket connection.
'       Receives "start" and "stop" messages.
'       Sends status/error/transcript messages back to this engine.
'
'   /redink/acs/transcription/{sessionId}
'       Used by Azure Communication Services.
'       Must be publicly reachable as WSS.
'       Receives ACS realtime transcription WebSocket messages.
'       Parses ACS TranscriptionMetadata and TranscriptionData packets.
'       Forwards normalized transcript messages to the engine client socket.
'
'   /redink/acs/callbacks
'       Used by Azure Communication Services Call Automation.
'       Must be publicly reachable as HTTPS.
'       Receives Call Automation lifecycle events such as call connected,
'       call disconnected, transcription started, transcription stopped, and
'       transcription failed.
'       Forwards useful status/error messages to the engine client socket.
'
' Start message sent by this engine to the bridge:
'
'       {
'         "type": "start",
'         "meetingJoinUrl": "https://teams.microsoft.com/l/meetup-join/...",
'         "locale": "de-DE",
'         "source": "RedInkWordTranscriptor"
'       }
'
' Stop message sent by this engine to the bridge:
'
'       {
'         "type": "stop",
'         "source": "RedInkWordTranscriptor"
'       }
'
' Simplified status message expected from the bridge:
'
'       {
'         "type": "status",
'         "message": "Joining Teams meeting..."
'       }
'
' Simplified error message expected from the bridge:
'
'       {
'         "type": "error",
'         "message": "Human-readable error text."
'       }
'
' Simplified final transcript message expected from the bridge:
'
'       {
'         "type": "final",
'         "speaker": "Participant name or participantRawID",
'         "text": "Recognized transcript text.",
'         "resultStatus": "Final"
'       }
'
' Simplified partial transcript message expected from the bridge:
'
'       {
'         "type": "partial",
'         "speaker": "Participant name or participantRawID",
'         "text": "Interim transcript text.",
'         "resultStatus": "Intermediate"
'       }
'
' Raw ACS message shape accepted by this engine:
'   This engine also accepts ACS-style messages where:
'
'       kind = "TranscriptionMetadata"
'
'   or:
'
'       kind = "TranscriptionData"
'       transcriptionData.text = transcript text
'       transcriptionData.participantRawID = participant identifier
'       transcriptionData.resultStatus = "Final" or intermediate status
'
'   If transcriptionData.participantRawID is missing, this engine also checks:
'
'       transcriptionData.participant.rawId
'
' Speaker handling:
'   The bridge should prefer a stable human-readable speaker label if it can map
'   the ACS participant identifier to a display name. Otherwise it should send
'   participantRawID.
'
'   Recommended bridge-side mapping:
'
'       participantRawID -> display name
'
'   If no display name is available, emit the raw participant ID. This engine
'   will pass the value into TranscriptionEventArgs.Speaker, and the existing
'   transcriptor UI can render:
'
'       Speaker: transcript text
'
' ACS implementation notes for the bridge:
'   The bridge must use Azure Communication Services Call Automation or another
'   documented ACS Teams interoperability path to join the Teams meeting as a
'   visible participant, for example "Red Ink Transcriptor".
'
'   The bridge must start realtime transcription for the ACS call and configure
'   a WebSocket transport URL pointing to:
'
'       wss://<public-url>/redink/acs/transcription/{sessionId}
'
'   The bridge must also configure a Call Automation callback URL pointing to:
'
'       https://<public-url>/redink/acs/callbacks
'
'   Required ACS realtime transcription settings conceptually include:
'       - transcription transport = WebSocket
'       - transport URI = public ACS transcription WebSocket URI
'       - locale or locales = selected language, e.g. de-DE
'       - start transcription = true, or start transcription after call connect
'       - enable intermediate results = true if partial results are desired
'
'   These names correspond to the documented ACS Call Automation
'   TranscriptionOptions concepts. Exact property names depend on the SDK used.
'
' Python bridge implementation option:
'   A bridge can be implemented entirely in Python with open-source web
'   components plus Microsoft Azure SDK packages.
'
'   Typical stack:
'       fastapi
'       uvicorn[standard]
'       websockets
'       azure-communication-callautomation
'       azure-communication-identity
'       python-dotenv
'
'   Python bridge endpoints should match the route contract above.
'
'   The Python bridge should:
'       1. Accept this engine on /redink/transcriptor/client/{sessionId}.
'       2. Receive the start message.
'       3. Join the Teams meeting as the Red Ink Transcriptor ACS participant.
'       4. Configure ACS realtime transcription to stream to
'          /redink/acs/transcription/{sessionId}.
'       5. Receive ACS transcription packets.
'       6. Convert them to simplified Red Ink messages.
'       7. Send those messages back to this engine.
'       8. On stop, stop transcription and hang up/leave the ACS call.
'
' Lifecycle requirements:
'   - On engine stop, the bridge must stop transcription and leave the Teams
'     meeting/call.
'   - On Word shutdown, all active bridge sessions must be stopped.
'   - If the engine WebSocket disconnects unexpectedly, the bridge should stop
'     the corresponding ACS transcription/call after a short grace period.
'   - If ACS disconnects, the bridge should send an error/status message to the
'     engine and close or mark the session as ended.
'
' Security requirements:
'   - Use HTTPS/WSS for all ACS-facing public endpoints.
'   - Authenticate the engine-to-bridge WebSocket if the bridge is not strictly
'     local.
'   - Validate session IDs.
'   - Do not expose ACS connection strings or Azure credentials to the Word
'     engine.
'   - Treat Teams meeting links as sensitive.
'   - Log only what is necessary for diagnostics.
'
' Privacy/compliance expectations:
'   The ACS participant is expected to join the Teams meeting visibly, for
'   example as "Red Ink Transcriptor". This must not be implemented as a hidden
'   listener. Tenant policy, consent, recording/transcription notifications, and
'   applicable legal requirements must be respected by the bridge/product.
'
' =============================================================================

Option Explicit On
Option Strict Off

Namespace Transcription

    Public Class TeamsAcsRealtimeEngine
        Implements ITranscriptionEngine

        Public Const DisplayNameValue As String = "Red Ink Teams Transcriptor Bot"

        Private ReadOnly _bridgeWebSocketUri As String
        Private ReadOnly _meetingJoinUrl As String
        Private ReadOnly _bridgeBearerToken As String

        Private _ws As System.Net.WebSockets.ClientWebSocket
        Private _readerTask As System.Threading.Tasks.Task
        Private _cts As System.Threading.CancellationTokenSource
        Private ReadOnly _sendLock As New System.Threading.SemaphoreSlim(1, 1)
        Private _stopStarted As Integer = 0

        Private Shared ReadOnly _supportedLanguages As String() = {
            "de-DE", "en-US", "en-GB", "fr-FR", "it-IT", "es-ES", "nl-NL"
        }

        Public Shared ReadOnly Property SupportedLanguages As String()
            Get
                Return _supportedLanguages.OrderBy(Function(x) x, System.StringComparer.OrdinalIgnoreCase).ToArray()
            End Get
        End Property

        Public Event PartialResult As System.EventHandler(Of TranscriptionEventArgs) Implements ITranscriptionEngine.PartialResult
        Public Event FinalResult As System.EventHandler(Of TranscriptionEventArgs) Implements ITranscriptionEngine.FinalResult
        Public Event EngineError As System.EventHandler(Of TranscriptionErrorEventArgs) Implements ITranscriptionEngine.EngineError
        Public Event Status As System.EventHandler(Of TranscriptionStatusEventArgs) Implements ITranscriptionEngine.Status

        Public Sub New(bridgeWebSocketUri As String, meetingJoinUrl As String, Optional bridgeBearerToken As String = "")
            _bridgeWebSocketUri = If(bridgeWebSocketUri, "").Trim()
            _meetingJoinUrl = If(meetingJoinUrl, "").Trim()
            _bridgeBearerToken = If(bridgeBearerToken, "").Trim()
        End Sub

        Public ReadOnly Property Name As String Implements ITranscriptionEngine.Name
            Get
                Return "Red Ink Teams Transcriptor Bot"
            End Get
        End Property

        Public ReadOnly Property Kind As EngineKind Implements ITranscriptionEngine.Kind
            Get
                Return EngineKind.TeamsAcsRealtime
            End Get
        End Property

        Public ReadOnly Property SupportsLiveStreaming As Boolean Implements ITranscriptionEngine.SupportsLiveStreaming
            Get
                Return True
            End Get
        End Property

        Public ReadOnly Property SupportsFileTranscription As Boolean Implements ITranscriptionEngine.SupportsFileTranscription
            Get
                Return False
            End Get
        End Property

        Public ReadOnly Property SupportsDiarization As Boolean Implements ITranscriptionEngine.SupportsDiarization
            Get
                Return True
            End Get
        End Property

        Public ReadOnly Property SupportsMultiChannelDiarization As Boolean Implements ITranscriptionEngine.SupportsMultiChannelDiarization
            Get
                Return False
            End Get
        End Property

        Public Async Function StartLiveAsync(opts As TranscriptionOptions, ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task Implements ITranscriptionEngine.StartLiveAsync
            If System.String.IsNullOrWhiteSpace(_bridgeWebSocketUri) Then
                Throw New System.InvalidOperationException("Teams ACS bridge WebSocket URI is missing.")
            End If

            If System.String.IsNullOrWhiteSpace(_meetingJoinUrl) Then
                Throw New System.InvalidOperationException("Teams meeting join URL is missing.")
            End If

            System.Threading.Interlocked.Exchange(_stopStarted, 0)
            _cts = System.Threading.CancellationTokenSource.CreateLinkedTokenSource(ct)
            _ws = New System.Net.WebSockets.ClientWebSocket()
            _ws.Options.KeepAliveInterval = System.TimeSpan.FromSeconds(20)

            If Not System.String.IsNullOrWhiteSpace(_bridgeBearerToken) Then
                _ws.Options.SetRequestHeader("Authorization", "Bearer " & _bridgeBearerToken)
            End If

            RaiseStatusMessage("Connecting to Teams ACS bridge…")

            Try
                Await _ws.ConnectAsync(New System.Uri(_bridgeWebSocketUri), _cts.Token).ConfigureAwait(False)
            Catch ex As System.Exception
                Dim detail As String = "Teams ACS bridge connection failed: " & GetDetailedExceptionMessage(ex)
                RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs(detail, ex, True))
                Throw New System.InvalidOperationException(detail, ex)
            End Try

            _readerTask = System.Threading.Tasks.Task.Run(Function() ReadLoopAsync(_cts.Token))

            Dim locale As String = NormalizeLanguageCode(opts)
            Dim startJson As Newtonsoft.Json.Linq.JObject = New Newtonsoft.Json.Linq.JObject From {
                {"type", "start"},
                {"meetingJoinUrl", _meetingJoinUrl},
                {"locale", locale},
                {"source", "RedInkWordTranscriptor"}
            }

            Await SendJsonAsync(startJson, _cts.Token).ConfigureAwait(False)
            RaiseStatusMessage("Teams ACS transcription requested (" & locale & ").")
        End Function

        Public Function PushAudioAsync(pcm As Byte(), bytesValid As Integer, ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task Implements ITranscriptionEngine.PushAudioAsync
            Return System.Threading.Tasks.Task.CompletedTask
        End Function

        Public Async Function StopLiveAsync() As System.Threading.Tasks.Task Implements ITranscriptionEngine.StopLiveAsync
            If System.Threading.Interlocked.Exchange(_stopStarted, 1) <> 0 Then
                Return
            End If

            Dim wsToClose As System.Net.WebSockets.ClientWebSocket = _ws
            Dim readerToAwait As System.Threading.Tasks.Task = _readerTask
            Dim ctsToCancel As System.Threading.CancellationTokenSource = _cts

            Try
                If wsToClose IsNot Nothing AndAlso wsToClose.State = System.Net.WebSockets.WebSocketState.Open Then
                    Dim stopJson As Newtonsoft.Json.Linq.JObject = New Newtonsoft.Json.Linq.JObject From {
                        {"type", "stop"},
                        {"source", "RedInkWordTranscriptor"}
                    }

                    Try
                        Await SendJsonAsync(stopJson, System.Threading.CancellationToken.None).ConfigureAwait(False)
                    Catch ex As System.Exception
                    End Try

                    Try
                        Await wsToClose.CloseOutputAsync(
                            System.Net.WebSockets.WebSocketCloseStatus.NormalClosure,
                            "stop",
                            System.Threading.CancellationToken.None).ConfigureAwait(False)
                    Catch ex As System.Exception
                    End Try
                End If
            Finally
                If ctsToCancel IsNot Nothing Then
                    Try
                        ctsToCancel.Cancel()
                    Catch ex As System.Exception
                    End Try
                End If
            End Try

            If readerToAwait IsNot Nothing Then
                Try
                    Await System.Threading.Tasks.Task.WhenAny(
                        readerToAwait,
                        System.Threading.Tasks.Task.Delay(1500)).ConfigureAwait(False)
                Catch ex As System.Exception
                End Try
            End If

            Try
                If wsToClose IsNot Nothing Then
                    wsToClose.Dispose()
                End If
            Catch ex As System.Exception
            End Try

            Try
                If ctsToCancel IsNot Nothing Then
                    ctsToCancel.Dispose()
                End If
            Catch ex As System.Exception
            End Try

            _ws = Nothing
            _readerTask = Nothing
            _cts = Nothing

            RaiseStatusMessage("Teams ACS transcription stopped.")
        End Function

        Public Function TranscribeFileAsync(filePath As String, opts As TranscriptionOptions, ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task Implements ITranscriptionEngine.TranscribeFileAsync
            Throw New System.NotSupportedException("Teams ACS realtime engine is live-only.")
        End Function

        Public Function DisposeAsync() As System.Threading.Tasks.ValueTask Implements System.IAsyncDisposable.DisposeAsync
            Return New System.Threading.Tasks.ValueTask(DisposeAsyncCoreAsync())
        End Function

        Private Async Function DisposeAsyncCoreAsync() As System.Threading.Tasks.Task
            Await StopLiveAsync().ConfigureAwait(False)

            Try
                _sendLock.Dispose()
            Catch ex As System.Exception
            End Try
        End Function

        Private Async Function ReadLoopAsync(ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task
            Dim buffer(64 * 1024 - 1) As Byte

            Try
                While _ws IsNot Nothing AndAlso _ws.State = System.Net.WebSockets.WebSocketState.Open AndAlso Not ct.IsCancellationRequested
                    Using ms As New System.IO.MemoryStream()
                        Dim result As System.Net.WebSockets.WebSocketReceiveResult = Nothing

                        Do
                            result = Await _ws.ReceiveAsync(New System.ArraySegment(Of Byte)(buffer), ct).ConfigureAwait(False)

                            If result.MessageType = System.Net.WebSockets.WebSocketMessageType.Close Then
                                Return
                            End If

                            ms.Write(buffer, 0, result.Count)
                        Loop While Not result.EndOfMessage

                        Dim text As String = System.Text.Encoding.UTF8.GetString(ms.ToArray())
                        HandleBridgeMessage(text)
                    End Using
                End While
            Catch ex As System.OperationCanceledException
            Catch ex As System.Exception
                If System.Threading.Interlocked.CompareExchange(_stopStarted, 0, 0) = 0 Then
                    RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("Teams ACS bridge read failed: " & GetDetailedExceptionMessage(ex), ex, False))
                End If
            End Try
        End Function

        Private Async Function SendJsonAsync(jo As Newtonsoft.Json.Linq.JObject, ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task
            If _ws Is Nothing OrElse _ws.State <> System.Net.WebSockets.WebSocketState.Open Then
                Return
            End If

            Dim json As String = jo.ToString(Newtonsoft.Json.Formatting.None)
            Dim bytes As Byte() = System.Text.Encoding.UTF8.GetBytes(json)

            Await _sendLock.WaitAsync(ct).ConfigureAwait(False)

            Try
                Await _ws.SendAsync(
                    New System.ArraySegment(Of Byte)(bytes),
                    System.Net.WebSockets.WebSocketMessageType.Text,
                    True,
                    ct).ConfigureAwait(False)
            Catch ex As System.Exception
                Dim detail As String = "Teams ACS bridge send failed: " & GetDetailedExceptionMessage(ex)
                RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs(detail, ex, False))
                Throw New System.InvalidOperationException(detail, ex)
            Finally
                _sendLock.Release()
            End Try
        End Function

        Private Sub HandleBridgeMessage(rawMessage As String)
            If System.String.IsNullOrWhiteSpace(rawMessage) Then
                Return
            End If

            Try
                Dim jo As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(rawMessage)

                Dim kind As String = If(jo("kind")?.ToString(), "").Trim()
                If System.String.Equals(kind, "TranscriptionMetadata", System.StringComparison.OrdinalIgnoreCase) Then
                    RaiseStatusMessage("Teams ACS transcription stream opened.")
                    Return
                End If

                If System.String.Equals(kind, "TranscriptionData", System.StringComparison.OrdinalIgnoreCase) Then
                    HandleAcsTranscriptionData(jo)
                    Return
                End If

                HandleSimplifiedBridgeMessage(jo)
            Catch ex As Newtonsoft.Json.JsonException
                RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("Teams ACS bridge sent invalid JSON: " & Truncate(rawMessage, 1000), ex, False))
            Catch ex As System.Exception
                RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("Teams ACS message handling failed: " & GetDetailedExceptionMessage(ex), ex, False))
            End Try
        End Sub

        Private Sub HandleAcsTranscriptionData(jo As Newtonsoft.Json.Linq.JObject)
            Dim data As Newtonsoft.Json.Linq.JObject = TryCast(jo("transcriptionData"), Newtonsoft.Json.Linq.JObject)
            If data Is Nothing Then
                Return
            End If

            Dim text As String = If(data("text")?.ToString(), "").Trim()
            If text.Length = 0 Then
                Return
            End If

            Dim participantRawId As String = If(data("participantRawID")?.ToString(), "").Trim()
            If participantRawId.Length = 0 Then
                Dim participant As Newtonsoft.Json.Linq.JObject = TryCast(data("participant"), Newtonsoft.Json.Linq.JObject)
                If participant IsNot Nothing Then
                    participantRawId = If(participant("rawId")?.ToString(), "").Trim()
                End If
            End If

            Dim resultStatus As String = If(data("resultStatus")?.ToString(), "").Trim()
            Dim isFinal As Boolean = System.String.Equals(resultStatus, "Final", System.StringComparison.OrdinalIgnoreCase)

            RaiseTranscript(text, participantRawId, isFinal)
        End Sub

        Private Sub HandleSimplifiedBridgeMessage(jo As Newtonsoft.Json.Linq.JObject)
            Dim typeValue As String = If(jo("type")?.ToString(), "").Trim()

            If System.String.Equals(typeValue, "status", System.StringComparison.OrdinalIgnoreCase) Then
                RaiseStatusMessage(If(jo("message")?.ToString(), "Teams ACS bridge status."))
                Return
            End If

            If System.String.Equals(typeValue, "error", System.StringComparison.OrdinalIgnoreCase) Then
                Dim errorMessage As String = If(jo("message")?.ToString(), "Teams ACS bridge error.")
                RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs(errorMessage, Nothing, False))
                Return
            End If

            Dim text As String = If(jo("text")?.ToString(), "").Trim()
            If text.Length = 0 Then
                Return
            End If

            Dim speaker As String = If(jo("speaker")?.ToString(), "").Trim()
            If speaker.Length = 0 Then
                speaker = If(jo("displayName")?.ToString(), "").Trim()
            End If
            If speaker.Length = 0 Then
                speaker = If(jo("participantRawID")?.ToString(), "").Trim()
            End If

            Dim resultStatus As String = If(jo("resultStatus")?.ToString(), "").Trim()
            Dim isFinal As Boolean =
                System.String.Equals(typeValue, "final", System.StringComparison.OrdinalIgnoreCase) OrElse
                System.String.Equals(resultStatus, "Final", System.StringComparison.OrdinalIgnoreCase)

            RaiseTranscript(text, speaker, isFinal)
        End Sub

        Private Sub RaiseTranscript(text As String, speaker As String, isFinal As Boolean)
            Dim ev As TranscriptionEventArgs = New TranscriptionEventArgs(text, isFinal)
            SetSpeakerIfAvailable(ev, speaker)

            If isFinal Then
                RaiseEvent FinalResult(Me, ev)
            Else
                RaiseEvent PartialResult(Me, ev)
            End If
        End Sub

        Private Shared Sub SetSpeakerIfAvailable(ev As TranscriptionEventArgs, speaker As String)
            If ev Is Nothing OrElse System.String.IsNullOrWhiteSpace(speaker) Then
                Return
            End If

            Try
                Dim p As System.Reflection.PropertyInfo = ev.GetType().GetProperty("Speaker")
                If p IsNot Nothing AndAlso p.CanWrite Then
                    p.SetValue(ev, speaker.Trim(), Nothing)
                End If
            Catch ex As System.Exception
            End Try
        End Sub

        Private Sub RaiseStatusMessage(message As String)
            RaiseEvent Status(Me, New TranscriptionStatusEventArgs(message))
        End Sub

        Private Shared Function NormalizeLanguageCode(opts As TranscriptionOptions) As String
            Dim raw As String = ""

            If opts IsNot Nothing Then
                raw = If(opts.LanguageCode, "").Trim()
            End If

            If System.String.IsNullOrWhiteSpace(raw) OrElse System.String.Equals(raw, "auto", System.StringComparison.OrdinalIgnoreCase) Then
                Return "de-DE"
            End If

            Select Case raw.Trim().ToLowerInvariant()
                Case "de" : Return "de-DE"
                Case "en" : Return "en-US"
                Case "fr" : Return "fr-FR"
                Case "it" : Return "it-IT"
                Case "es" : Return "es-ES"
                Case "nl" : Return "nl-NL"
            End Select

            Return raw
        End Function

        Private Shared Function GetDetailedExceptionMessage(ex As System.Exception) As String
            If ex Is Nothing Then
                Return ""
            End If

            Dim sb As New System.Text.StringBuilder()
            Dim current As System.Exception = ex

            While current IsNot Nothing
                If sb.Length > 0 Then
                    sb.Append(" -> ")
                End If

                sb.Append(current.Message)
                current = current.InnerException
            End While

            Return sb.ToString()
        End Function

        Private Shared Function Truncate(value As String, maxLength As Integer) As String
            Dim text As String = If(value, "")

            If maxLength <= 0 OrElse text.Length <= maxLength Then
                Return text
            End If

            Return text.Substring(0, maxLength) & "…"
        End Function

    End Class

End Namespace
