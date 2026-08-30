' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: GeminiTranscribeLiveEngine.vb
' Purpose: Live transcription using Google Gemini 3.5 Transcribe Live over the
'          documented BidiGenerateContent WebSocket API. Uses only .NET/BCL
'          networking and the existing Red Ink audio capture contract.
' =============================================================================

Option Explicit On
Option Strict On

Imports System.Collections.Concurrent
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Net.WebSockets
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports Newtonsoft.Json.Linq

Namespace Transcription

    Public Class GeminiTranscribeLiveEngine
        Implements ITranscriptionEngine

        Public Const DisplayNameValue As System.String = "Google Gemini 3.5 Transcribe (live)"
        Public Const DefaultVertexModel As System.String = "gemini-3.5-transcribe-live-preview"
        Public Const DefaultGeminiApiModel As System.String = "gemini-3.5-transcribe-live"
        Public Const DefaultLocation As System.String = "global"

        ' Google documents a 10 minute live-transcription/session ceiling and an
        ' approximately 10 minute WebSocket connection lifetime. Rotate early so
        ' the next capture frame always has a healthy session available.
        Private Shared ReadOnly SessionRotationAge As System.TimeSpan = System.TimeSpan.FromMinutes(8.75R)
        Private Shared ReadOnly SetupTimeout As System.TimeSpan = System.TimeSpan.FromSeconds(20)
        Private Shared ReadOnly FlushTimeout As System.TimeSpan = System.TimeSpan.FromSeconds(1.5R)

        Public Event PartialResult As System.EventHandler(Of TranscriptionEventArgs) Implements ITranscriptionEngine.PartialResult
        Public Event FinalResult As System.EventHandler(Of TranscriptionEventArgs) Implements ITranscriptionEngine.FinalResult
        Public Event EngineError As System.EventHandler(Of TranscriptionErrorEventArgs) Implements ITranscriptionEngine.EngineError
        Public Event Status As System.EventHandler(Of TranscriptionStatusEventArgs) Implements ITranscriptionEngine.Status

        Private ReadOnly _projectId As System.String
        Private ReadOnly _location As System.String
        Private ReadOnly _model As System.String
        Private ReadOnly _tokenFactory As System.Func(Of System.Threading.Tasks.Task(Of System.String))
        Private ReadOnly _apiKey As System.String
        Private ReadOnly _useVertex As System.Boolean
        Private ReadOnly _mode As System.String
        Private ReadOnly _customVocabulary As System.Collections.Generic.List(Of System.String)
        Private ReadOnly _endpointOverride As System.String
        Private ReadOnly _sendGate As New System.Threading.SemaphoreSlim(1, 1)
        Private ReadOnly _lifecycleGate As New System.Threading.SemaphoreSlim(1, 1)
        Private ReadOnly _audioQueue As New System.Collections.Concurrent.ConcurrentQueue(Of System.Byte())()
        Private ReadOnly _audioSignal As New System.Threading.SemaphoreSlim(0)
        Private ReadOnly _queueStateLock As New System.Object()

        Private _socket As System.Net.WebSockets.ClientWebSocket
        Private _senderCts As System.Threading.CancellationTokenSource
        Private _senderTask As System.Threading.Tasks.Task
        Private _drainCompletion As New System.Threading.Tasks.TaskCompletionSource(Of System.Boolean)(System.Threading.Tasks.TaskCreationOptions.RunContinuationsAsynchronously)
        Private _queuedFrameCount As System.Int32
        Private _acceptingAudio As System.Boolean
        Private _receiveCts As System.Threading.CancellationTokenSource
        Private _receiveTask As System.Threading.Tasks.Task
        Private _sessionStartedUtc As System.DateTime = System.DateTime.MinValue
        Private _rotationRequested As System.Boolean
        Private _stopping As System.Boolean
        Private _opts As TranscriptionOptions
        Private _setupCompletion As System.Threading.Tasks.TaskCompletionSource(Of System.Boolean)
        Private _flushCompletion As System.Threading.Tasks.TaskCompletionSource(Of System.Boolean)
        Private _lastEmittedText As System.String = System.String.Empty
        Private _firstAudioFrameLogged As System.Boolean
        Private _setupCompletedForCurrentSession As System.Boolean

        Public Sub New(projectId As System.String,
                       tokenFactory As System.Func(Of System.Threading.Tasks.Task(Of System.String)),
                       Optional model As System.String = DefaultVertexModel,
                       Optional location As System.String = DefaultLocation,
                       Optional mode As System.String = "VERBATIM",
                       Optional customVocabulary As System.Collections.Generic.IEnumerable(Of System.String) = Nothing,
                       Optional endpointOverride As System.String = "")

            _projectId = If(projectId, System.String.Empty).Trim()
            _tokenFactory = tokenFactory
            _apiKey = System.String.Empty
            _useVertex = True
            _model = If(System.String.IsNullOrWhiteSpace(model), DefaultVertexModel, model.Trim())
            _location = NormalizeLocation(location)
            _endpointOverride = If(endpointOverride, System.String.Empty).Trim()
            _mode = NormalizeMode(mode)
            _customVocabulary = NormalizeVocabulary(customVocabulary)

            If System.String.IsNullOrWhiteSpace(_projectId) Then
                Throw New System.ArgumentException("Google Cloud project id is required.", NameOf(projectId))
            End If
            If _tokenFactory Is Nothing Then
                Throw New System.ArgumentNullException(NameOf(tokenFactory))
            End If
        End Sub

        Public Sub New(apiKey As System.String,
                       Optional model As System.String = DefaultGeminiApiModel,
                       Optional mode As System.String = "VERBATIM",
                       Optional customVocabulary As System.Collections.Generic.IEnumerable(Of System.String) = Nothing,
                       Optional endpointOverride As System.String = "")

            _projectId = System.String.Empty
            _tokenFactory = Nothing
            _apiKey = If(apiKey, System.String.Empty).Trim()
            _useVertex = False
            _model = If(System.String.IsNullOrWhiteSpace(model), DefaultGeminiApiModel, model.Trim())
            _location = DefaultLocation
            _endpointOverride = If(endpointOverride, System.String.Empty).Trim()
            _mode = NormalizeMode(mode)
            _customVocabulary = NormalizeVocabulary(customVocabulary)

            If System.String.IsNullOrWhiteSpace(_apiKey) Then
                Throw New System.ArgumentException("Gemini API key is required.", NameOf(apiKey))
            End If
        End Sub

        Public ReadOnly Property Name As System.String Implements ITranscriptionEngine.Name
            Get
                Return DisplayNameValue
            End Get
        End Property

        Public ReadOnly Property Kind As EngineKind Implements ITranscriptionEngine.Kind
            Get
                Return EngineKind.GeminiTranscribeLive
            End Get
        End Property

        Public ReadOnly Property SupportsLiveStreaming As System.Boolean Implements ITranscriptionEngine.SupportsLiveStreaming
            Get
                Return True
            End Get
        End Property

        Public ReadOnly Property SupportsFileTranscription As System.Boolean Implements ITranscriptionEngine.SupportsFileTranscription
            Get
                Return False
            End Get
        End Property

        Public ReadOnly Property SupportsDiarization As System.Boolean Implements ITranscriptionEngine.SupportsDiarization
            Get
                Return False
            End Get
        End Property

        Public ReadOnly Property SupportsMultiChannelDiarization As System.Boolean Implements ITranscriptionEngine.SupportsMultiChannelDiarization
            Get
                Return False
            End Get
        End Property

        Public Async Function StartLiveAsync(opts As TranscriptionOptions, ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task Implements ITranscriptionEngine.StartLiveAsync
            _opts = If(opts, New TranscriptionOptions())
            _stopping = False
            _rotationRequested = False
            _lastEmittedText = System.String.Empty
            _firstAudioFrameLogged = False
            _setupCompletedForCurrentSession = False
            DebugLog("StartLiveAsync entered. auth=" & If(_useVertex, "VertexOAuth2", "GeminiApiKey") &
                     ", model=" & _model &
                     ", location=" & _location &
                     ", mode=" & _mode &
                     ", language=" & If(_opts Is Nothing, "<none>", If(_opts.LanguageCode, System.String.Empty)) &
                     ", customVocabularyCount=" & _customVocabulary.Count.ToString(System.Globalization.CultureInfo.InvariantCulture))

            Dim discardedFrame As System.Byte() = Nothing
            Do While _audioQueue.TryDequeue(discardedFrame)
            Loop
            System.Threading.Interlocked.Exchange(_queuedFrameCount, 0)

            Dim previousSenderCts As System.Threading.CancellationTokenSource = _senderCts
            If previousSenderCts IsNot Nothing Then
                Try
                    previousSenderCts.Cancel()
                Catch
                End Try
                previousSenderCts.Dispose()
            End If

            _senderCts = System.Threading.CancellationTokenSource.CreateLinkedTokenSource(ct)
            _drainCompletion = New System.Threading.Tasks.TaskCompletionSource(Of System.Boolean)(System.Threading.Tasks.TaskCreationOptions.RunContinuationsAsynchronously)

            If _opts.EnableDiarization Then
                RaiseStatusMessage("Gemini 3.5 Transcribe Live does not support speaker diarization; continuing without diarization.")
            End If

            Try
                Await _lifecycleGate.WaitAsync(ct)
                Try
                    Await OpenConnectionAsync(ct)
                Finally
                    _lifecycleGate.Release()
                End Try

                SyncLock _queueStateLock
                    _acceptingAudio = True
                End SyncLock
                _senderTask = AudioSenderLoopAsync(_senderCts.Token)
            Catch ex As System.Exception
                SyncLock _queueStateLock
                    _acceptingAudio = False
                End SyncLock
                DebugException("StartLiveAsync failed", ex)
                Throw
            End Try
        End Function

        Public Function PushAudioAsync(pcm As System.Byte(), bytesValid As System.Int32, ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task Implements ITranscriptionEngine.PushAudioAsync
            ct.ThrowIfCancellationRequested()

            If pcm Is Nothing OrElse bytesValid <= 0 Then
                Return System.Threading.Tasks.Task.CompletedTask
            End If
            If bytesValid > pcm.Length Then
                bytesValid = pcm.Length
            End If

            Dim audioBytes(bytesValid - 1) As System.Byte
            System.Buffer.BlockCopy(pcm, 0, audioBytes, 0, bytesValid)

            SyncLock _queueStateLock
                If Not _acceptingAudio OrElse _stopping Then
                    Return System.Threading.Tasks.Task.CompletedTask
                End If

                _audioQueue.Enqueue(audioBytes)
                System.Threading.Interlocked.Increment(_queuedFrameCount)
                _audioSignal.Release()
            End SyncLock

            Return System.Threading.Tasks.Task.CompletedTask
        End Function

        Public Async Function StopLiveAsync() As System.Threading.Tasks.Task Implements ITranscriptionEngine.StopLiveAsync
            _stopping = True

            Dim drainTask As System.Threading.Tasks.Task
            SyncLock _queueStateLock
                _acceptingAudio = False
                If System.Threading.Interlocked.CompareExchange(_queuedFrameCount, 0, 0) = 0 Then
                    _drainCompletion.TrySetResult(True)
                End If
                drainTask = _drainCompletion.Task
            End SyncLock

            ' The host stops capture before calling StopLiveAsync. Drain every
            ' frame accepted before that point before the final stream flush.
            _audioSignal.Release()
            Await drainTask

            Dim senderCts As System.Threading.CancellationTokenSource = _senderCts
            If senderCts IsNot Nothing Then
                Try
                    senderCts.Cancel()
                Catch
                End Try
            End If

            Dim senderTask As System.Threading.Tasks.Task = _senderTask
            If senderTask IsNot Nothing Then
                Try
                    Await senderTask
                Catch ex As System.OperationCanceledException
                End Try
            End If

            Await _sendGate.WaitAsync()
            Try
                Await FlushAndCloseCurrentSocketAsync(System.Threading.CancellationToken.None)
            Finally
                _sendGate.Release()
            End Try
        End Function

        Private Async Function AudioSenderLoopAsync(ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task
            Try
                Do While Not ct.IsCancellationRequested
                    Await _audioSignal.WaitAsync(ct)

                    Dim audioBytes As System.Byte() = Nothing
                    Do While _audioQueue.TryDequeue(audioBytes)
                        Await SendAcceptedAudioFrameAsync(audioBytes, ct)

                        Dim remaining As System.Int32 = System.Threading.Interlocked.Decrement(_queuedFrameCount)
                        If remaining = 0 Then
                            SyncLock _queueStateLock
                                If Not _acceptingAudio Then
                                    _drainCompletion.TrySetResult(True)
                                End If
                            End SyncLock
                        End If
                    Loop
                Loop
            Catch ex As System.OperationCanceledException When ct.IsCancellationRequested
            Catch ex As System.Exception
                If Not _stopping Then
                    RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("Gemini 3.5 Transcribe Live audio sender stopped: " & ex.Message, ex, False))
                End If
                Throw
            End Try
        End Function

        Private Async Function SendAcceptedAudioFrameAsync(audioBytes As System.Byte(), ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task
            If audioBytes Is Nothing OrElse audioBytes.Length = 0 Then
                Return
            End If

            Do
                ct.ThrowIfCancellationRequested()

                Try
                    Await _sendGate.WaitAsync(ct)
                    Try
                        If NeedsRotation() Then
                            Await RotateSessionAsync(ct)
                        End If

                        Dim mediaChunk As New Newtonsoft.Json.Linq.JObject()
                        mediaChunk("mimeType") = "audio/pcm;rate=16000"
                        mediaChunk("data") = System.Convert.ToBase64String(audioBytes)

                        Dim realtimeInput As New Newtonsoft.Json.Linq.JObject()
                        realtimeInput("audio") = mediaChunk

                        Dim root As New Newtonsoft.Json.Linq.JObject()
                        root("realtimeInput") = realtimeInput

                        If Not _firstAudioFrameLogged Then
                            _firstAudioFrameLogged = True
                            DebugLog("Sending first audio frame. bytes=" & audioBytes.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                                     ", socketState=" & If(_socket Is Nothing, "<null>", _socket.State.ToString()))
                        End If

                        Await SendJsonAsync(_socket, root, ct)
                        Return
                    Finally
                        _sendGate.Release()
                    End Try
                Catch ex As System.OperationCanceledException When ct.IsCancellationRequested
                    Throw
                Catch ex As System.Exception
                    DebugException("SendAcceptedAudioFrameAsync transport failure", ex)
                    ' Preserve the accepted frame across a reconnect. A transport
                    ' failure can make server receipt ambiguous, so a retry can
                    ' theoretically duplicate a tiny boundary frame, but it
                    ' avoids a known audio gap.
                    _rotationRequested = True
                    If Not _stopping Then
                        RaiseStatusMessage("Gemini 3.5 Transcribe Live connection interrupted; retaining queued audio and reconnecting automatically…")
                    End If
                End Try

                ' VB.NET does not permit Await inside Catch/Finally/SyncLock.
                ' Back off only after the exception handler has completed.
                Await System.Threading.Tasks.Task.Delay(250, ct)
            Loop
        End Function

        Public Function TranscribeFileAsync(filePath As System.String, opts As TranscriptionOptions, ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task Implements ITranscriptionEngine.TranscribeFileAsync
            Throw New System.NotSupportedException("Gemini 3.5 Transcribe Live is a live-only engine. Use the Gemini file engine for pre-recorded audio.")
        End Function

        Private Function NeedsRotation() As System.Boolean
            Dim ws As System.Net.WebSockets.ClientWebSocket = _socket
            If ws Is Nothing OrElse ws.State <> System.Net.WebSockets.WebSocketState.Open Then
                Return True
            End If
            If _rotationRequested Then
                Return True
            End If
            If _sessionStartedUtc = System.DateTime.MinValue Then
                Return True
            End If
            Return (System.DateTime.UtcNow - _sessionStartedUtc) >= SessionRotationAge
        End Function

        Private Async Function RotateSessionAsync(ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task
            Await _lifecycleGate.WaitAsync(ct)
            Try
                If Not NeedsRotation() Then
                    Return
                End If

                RaiseStatusMessage("Gemini 3.5 Transcribe Live session limit approaching; starting the next session automatically…")
                Await FlushAndCloseCurrentSocketAsync(ct)
                Await OpenConnectionAsync(ct)
                _rotationRequested = False
                RaiseStatusMessage("Gemini 3.5 Transcribe Live session continued automatically.")
            Finally
                _lifecycleGate.Release()
            End Try
        End Function

        Private Async Function OpenConnectionAsync(ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task
            ct.ThrowIfCancellationRequested()
            _setupCompletedForCurrentSession = False
            DebugLog("OpenConnectionAsync starting. auth=" & If(_useVertex, "VertexOAuth2", "GeminiApiKey") &
                     ", modelResource=" & BuildModelResourceName())

            Dim ws As New System.Net.WebSockets.ClientWebSocket()
            ws.Options.KeepAliveInterval = System.TimeSpan.FromSeconds(20)

            If _useVertex Then
                DebugLog("Requesting Vertex OAuth2 access token.")
                Dim token As System.String = Await _tokenFactory()
                If System.String.IsNullOrWhiteSpace(token) Then
                    ws.Dispose()
                    Throw New System.InvalidOperationException("Google OAuth2 access token is empty.")
                End If
                DebugLog("Vertex OAuth2 token obtained. tokenLength=" & token.Length.ToString(System.Globalization.CultureInfo.InvariantCulture))
                ws.Options.SetRequestHeader("Authorization", "Bearer " & token)
            End If

            Dim endpoint As New System.Uri(BuildWebSocketEndpoint())
            DebugLog("Connecting WebSocket to " & GetSafeEndpointForDebug(endpoint))
            Try
                Await ws.ConnectAsync(endpoint, ct)
            Catch ex As System.Exception
                DebugException("WebSocket ConnectAsync failed for " & GetSafeEndpointForDebug(endpoint), ex)
                ws.Dispose()
                Throw
            End Try
            DebugLog("WebSocket connected. state=" & ws.State.ToString())

            Dim oldReceiveCts As System.Threading.CancellationTokenSource = _receiveCts
            If oldReceiveCts IsNot Nothing Then
                Try
                    oldReceiveCts.Cancel()
                Catch
                End Try
                oldReceiveCts.Dispose()
            End If

            _socket = ws
            _receiveCts = System.Threading.CancellationTokenSource.CreateLinkedTokenSource(ct)
            _setupCompletion = New System.Threading.Tasks.TaskCompletionSource(Of System.Boolean)(System.Threading.Tasks.TaskCreationOptions.RunContinuationsAsynchronously)
            _flushCompletion = New System.Threading.Tasks.TaskCompletionSource(Of System.Boolean)(System.Threading.Tasks.TaskCreationOptions.RunContinuationsAsynchronously)
            _receiveTask = ReceiveLoopAsync(ws, _receiveCts.Token)

            Dim setup As Newtonsoft.Json.Linq.JObject = BuildSetupMessage()
            DebugLog("Sending setup. model=" & BuildModelResourceName() &
                     ", language=" & If(_opts Is Nothing, "<none>", If(_opts.LanguageCode, System.String.Empty)) &
                     ", mode=" & _mode &
                     ", customVocabularyCount=" & _customVocabulary.Count.ToString(System.Globalization.CultureInfo.InvariantCulture))
            Await SendJsonAsync(ws, setup, ct)
            DebugLog("Setup payload sent; waiting for setupComplete.")

            Dim completed As System.Threading.Tasks.Task = Await System.Threading.Tasks.Task.WhenAny(_setupCompletion.Task, System.Threading.Tasks.Task.Delay(SetupTimeout, ct))
            If completed IsNot _setupCompletion.Task Then
                DebugLog("setupComplete timeout. socketState=" & ws.State.ToString())
                Throw New System.TimeoutException("Gemini 3.5 Transcribe Live setup did not complete within the expected time.")
            End If

            Await _setupCompletion.Task
            _setupCompletedForCurrentSession = True
            _sessionStartedUtc = System.DateTime.UtcNow
            DebugLog("setupComplete received. Live session is ready.")
        End Function

        Private Function BuildSetupMessage() As Newtonsoft.Json.Linq.JObject
            Dim setup As New Newtonsoft.Json.Linq.JObject()
            setup("model") = BuildModelResourceName()

            Dim generationConfig As New Newtonsoft.Json.Linq.JObject()
            Dim modalities As New Newtonsoft.Json.Linq.JArray()
            modalities.Add("TEXT")
            generationConfig("responseModalities") = modalities
            setup("generationConfig") = generationConfig

            Dim transcriptionConfig As New Newtonsoft.Json.Linq.JObject()
            Dim languageCodes As New Newtonsoft.Json.Linq.JArray()
            Dim language As System.String = If(_opts Is Nothing, System.String.Empty, If(_opts.LanguageCode, System.String.Empty)).Trim()
            If Not System.String.IsNullOrWhiteSpace(language) AndAlso Not System.String.Equals(language, "auto", System.StringComparison.OrdinalIgnoreCase) Then
                languageCodes.Add(language)
            End If
            transcriptionConfig("languageCodes") = languageCodes
            transcriptionConfig("mode") = _mode

            If _customVocabulary.Count > 0 Then
                Dim vocabulary As New Newtonsoft.Json.Linq.JArray()
                For Each term As System.String In _customVocabulary
                    vocabulary.Add(term)
                Next
                transcriptionConfig("customVocabulary") = vocabulary
            End If

            setup("inputAudioTranscription") = transcriptionConfig

            Dim root As New Newtonsoft.Json.Linq.JObject()
            root("setup") = setup
            Return root
        End Function

        Private Async Function FlushAndCloseCurrentSocketAsync(ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task
            Dim ws As System.Net.WebSockets.ClientWebSocket = _socket
            If ws Is Nothing Then
                Return
            End If

            If ws.State = System.Net.WebSockets.WebSocketState.Open Then
                Try
                    _flushCompletion = New System.Threading.Tasks.TaskCompletionSource(Of System.Boolean)(System.Threading.Tasks.TaskCreationOptions.RunContinuationsAsynchronously)

                    Dim realtimeInput As New Newtonsoft.Json.Linq.JObject()
                    realtimeInput("audioStreamEnd") = True
                    Dim root As New Newtonsoft.Json.Linq.JObject()
                    root("realtimeInput") = realtimeInput
                    Await SendJsonAsync(ws, root, ct)

                    Dim delayTask As System.Threading.Tasks.Task = System.Threading.Tasks.Task.Delay(FlushTimeout, ct)
                    Await System.Threading.Tasks.Task.WhenAny(_flushCompletion.Task, delayTask)
                Catch ex As System.OperationCanceledException When ct.IsCancellationRequested
                    Throw
                Catch
                End Try
            End If

            Dim receiveCts As System.Threading.CancellationTokenSource = _receiveCts
            _receiveCts = Nothing
            If receiveCts IsNot Nothing Then
                Try
                    receiveCts.Cancel()
                Catch
                End Try
            End If

            Try
                If ws.State = System.Net.WebSockets.WebSocketState.Open OrElse ws.State = System.Net.WebSockets.WebSocketState.CloseReceived Then
                    Await ws.CloseOutputAsync(System.Net.WebSockets.WebSocketCloseStatus.NormalClosure, "session rotation", System.Threading.CancellationToken.None)
                End If
            Catch
            End Try

            If System.Object.ReferenceEquals(_socket, ws) Then
                _socket = Nothing
            End If

            Try
                ws.Dispose()
            Catch
            End Try

            If receiveCts IsNot Nothing Then
                receiveCts.Dispose()
            End If
        End Function

        Private Async Function ReceiveLoopAsync(ws As System.Net.WebSockets.ClientWebSocket, ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task
            Dim buffer(16383) As System.Byte

            Try
                Do While Not ct.IsCancellationRequested AndAlso ws.State = System.Net.WebSockets.WebSocketState.Open
                    Using ms As New System.IO.MemoryStream()
                        Dim result As System.Net.WebSockets.WebSocketReceiveResult = Nothing
                        Do
                            result = Await ws.ReceiveAsync(New System.ArraySegment(Of System.Byte)(buffer), ct)
                            If result.MessageType = System.Net.WebSockets.WebSocketMessageType.Close Then
                                DebugLog("WebSocket close received. status=" & If(result.CloseStatus.HasValue, result.CloseStatus.Value.ToString(), "<none>") &
                                         ", description=" & If(result.CloseStatusDescription, System.String.Empty))
                                Exit Do
                            End If
                            If result.Count > 0 Then
                                ms.Write(buffer, 0, result.Count)
                            End If
                        Loop Until result.EndOfMessage

                        If result Is Nothing OrElse result.MessageType = System.Net.WebSockets.WebSocketMessageType.Close Then
                            Exit Do
                        End If

                        Dim json As System.String = System.Text.Encoding.UTF8.GetString(ms.ToArray())
                        If Not _setupCompletedForCurrentSession Then
                            DebugLog("Pre-setup server message: " & TruncateForDebug(json, 4000))
                        End If
                        HandleServerMessage(json)
                    End Using
                Loop
            Catch ex As System.OperationCanceledException When ct.IsCancellationRequested
            Catch ex As System.Exception
                DebugException("ReceiveLoopAsync failed", ex)
                If Not _stopping Then
                    _rotationRequested = True
                    RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("Gemini 3.5 Transcribe Live receive loop interrupted: " & ex.Message, ex, False))
                End If
            Finally
                If Not _stopping AndAlso System.Object.ReferenceEquals(_socket, ws) Then
                    _rotationRequested = True
                End If
            End Try
        End Function

        Private Sub HandleServerMessage(json As System.String)
            If System.String.IsNullOrWhiteSpace(json) Then
                Return
            End If

            Dim root As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(json)

            If root("setupComplete") IsNot Nothing Then
                DebugLog("Server message contains setupComplete.")
                Dim setupCompletion As System.Threading.Tasks.TaskCompletionSource(Of System.Boolean) = _setupCompletion
                If setupCompletion IsNot Nothing Then
                    setupCompletion.TrySetResult(True)
                End If
            End If

            If root("goAway") IsNot Nothing Then
                _rotationRequested = True
                RaiseStatusMessage("Gemini Live requested connection rotation; the next audio frame will start a fresh session automatically.")
            End If

            EmitTranscriptionToken(root.SelectToken("serverContent.interimInputTranscription"), False)
            EmitTranscriptionToken(root.SelectToken("serverContent.inputTranscription"), True)
            EmitTranscriptionToken(root("inputTranscription"), True)

            If root.SelectToken("serverContent.turnComplete") IsNot Nothing AndAlso
               root.SelectToken("serverContent.turnComplete").Type = Newtonsoft.Json.Linq.JTokenType.Boolean AndAlso
               root.SelectToken("serverContent.turnComplete").Value(Of System.Boolean)() Then
                Dim flushCompletion As System.Threading.Tasks.TaskCompletionSource(Of System.Boolean) = _flushCompletion
                If flushCompletion IsNot Nothing Then
                    flushCompletion.TrySetResult(True)
                End If
            End If
        End Sub

        Private Sub EmitTranscriptionToken(token As Newtonsoft.Json.Linq.JToken, defaultFinal As System.Boolean)
            If token Is Nothing Then
                Return
            End If

            Dim textToken As Newtonsoft.Json.Linq.JToken = token("text")
            If textToken Is Nothing Then
                Return
            End If

            Dim text As System.String = textToken.ToString().Trim()
            If System.String.IsNullOrWhiteSpace(text) Then
                Return
            End If

            Dim isFinal As System.Boolean = defaultFinal
            Dim finishedToken As Newtonsoft.Json.Linq.JToken = token("finished")
            If finishedToken IsNot Nothing AndAlso finishedToken.Type = Newtonsoft.Json.Linq.JTokenType.Boolean Then
                isFinal = finishedToken.Value(Of System.Boolean)()
            End If

            If isFinal AndAlso System.String.Equals(text, _lastEmittedText, System.StringComparison.Ordinal) Then
                Return
            End If

            If isFinal Then
                _lastEmittedText = text
                RaiseEvent FinalResult(Me, New TranscriptionEventArgs(text, True))
                Dim flushCompletion As System.Threading.Tasks.TaskCompletionSource(Of System.Boolean) = _flushCompletion
                If flushCompletion IsNot Nothing Then
                    flushCompletion.TrySetResult(True)
                End If
            Else
                RaiseEvent PartialResult(Me, New TranscriptionEventArgs(text, False))
            End If
        End Sub

        Private Shared Async Function SendJsonAsync(ws As System.Net.WebSockets.ClientWebSocket,
                                                    payload As Newtonsoft.Json.Linq.JObject,
                                                    ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task
            If ws Is Nothing OrElse ws.State <> System.Net.WebSockets.WebSocketState.Open Then
                Throw New System.InvalidOperationException("Gemini Live WebSocket is not open.")
            End If

            Dim bytes As System.Byte() = System.Text.Encoding.UTF8.GetBytes(payload.ToString(Newtonsoft.Json.Formatting.None))
            Await ws.SendAsync(New System.ArraySegment(Of System.Byte)(bytes), System.Net.WebSockets.WebSocketMessageType.Text, True, ct)
        End Function

        Private Shared Function NormalizeLocation(requestedLocation As System.String) As System.String
            Dim normalizedLocation As System.String = If(requestedLocation, System.String.Empty).Trim()
            If System.String.IsNullOrWhiteSpace(normalizedLocation) Then
                Return DefaultLocation
            End If
            Return normalizedLocation
        End Function

        Private Function BuildWebSocketEndpoint() As System.String
            If Not System.String.IsNullOrWhiteSpace(_endpointOverride) Then
                Return ExpandEndpointTemplate(_endpointOverride)
            End If

            If Not _useVertex Then
                Return "wss://generativelanguage.googleapis.com/ws/google.ai.generativelanguage.v1beta.GenerativeService.BidiGenerateContent?key=" & System.Uri.EscapeDataString(_apiKey)
            End If

            Dim host As System.String
            If System.String.Equals(_location, "global", System.StringComparison.OrdinalIgnoreCase) Then
                host = "aiplatform.googleapis.com"
            Else
                host = _location & "-aiplatform.googleapis.com"
            End If

            Return "wss://" & host & "/ws/google.cloud.aiplatform.v1beta1.LlmBidiService/BidiGenerateContent"
        End Function

        Private Function ExpandEndpointTemplate(template As System.String) As System.String
            Dim result As System.String = If(template, System.String.Empty).Trim()
            result = result.Replace("{project}", System.Uri.EscapeDataString(_projectId))
            result = result.Replace("{location}", System.Uri.EscapeDataString(_location))
            result = result.Replace("{model}", System.Uri.EscapeDataString(_model))
            If Not _useVertex AndAlso result.IndexOf("{api_key}", System.StringComparison.OrdinalIgnoreCase) >= 0 Then
                result = result.Replace("{api_key}", System.Uri.EscapeDataString(_apiKey))
            End If
            Return result
        End Function

        Private Function BuildModelResourceName() As System.String
            If Not _useVertex Then
                Return "models/" & _model
            End If

            Return "projects/" & _projectId & "/locations/" & _location & "/publishers/google/models/" & _model
        End Function

        Private Shared Function NormalizeMode(value As System.String) As System.String
            Dim normalized As System.String = If(value, System.String.Empty).Trim().ToUpperInvariant()
            If System.String.Equals(normalized, "SMART", System.StringComparison.Ordinal) Then
                Return "SMART"
            End If
            Return "VERBATIM"
        End Function

        Private Shared Function NormalizeVocabulary(values As System.Collections.Generic.IEnumerable(Of System.String)) As System.Collections.Generic.List(Of System.String)
            Return If(values, System.Linq.Enumerable.Empty(Of System.String)()).
                Where(Function(x As System.String) Not System.String.IsNullOrWhiteSpace(x)).
                Select(Function(x As System.String) x.Trim()).
                Distinct(System.StringComparer.OrdinalIgnoreCase).
                Take(1000).
                ToList()
        End Function

        Private Shared Sub DebugLog(message As System.String)
            Try
                System.Diagnostics.Debug.WriteLine("[GeminiTranscribeLive] " & System.DateTime.Now.ToString("HH:mm:ss.fff", System.Globalization.CultureInfo.InvariantCulture) & " " & If(message, System.String.Empty))
            Catch
            End Try
        End Sub

        Private Shared Sub DebugException(context As System.String, ex As System.Exception)
            If ex Is Nothing Then
                DebugLog(context & ": <no exception>")
                Return
            End If

            DebugLog(context & ": " & ex.GetType().FullName & ": " & ex.Message)
            If ex.InnerException IsNot Nothing Then
                DebugLog(context & " inner: " & ex.InnerException.GetType().FullName & ": " & ex.InnerException.Message)
            End If
            If Not System.String.IsNullOrWhiteSpace(ex.StackTrace) Then
                DebugLog(context & " stack: " & TruncateForDebug(ex.StackTrace, 8000))
            End If
        End Sub

        Private Shared Function GetSafeEndpointForDebug(endpoint As System.Uri) As System.String
            If endpoint Is Nothing Then Return "<null>"
            Return endpoint.GetLeftPart(System.UriPartial.Path)
        End Function

        Private Shared Function TruncateForDebug(value As System.String, maxLength As System.Int32) As System.String
            Dim text As System.String = If(value, System.String.Empty)
            If maxLength <= 0 OrElse text.Length <= maxLength Then Return text
            Return text.Substring(0, maxLength) & "…[truncated]"
        End Function

        Private Sub RaiseStatusMessage(message As System.String)
            RaiseEvent Status(Me, New TranscriptionStatusEventArgs(message))
        End Sub

        Public Function DisposeAsync() As System.Threading.Tasks.ValueTask Implements System.IAsyncDisposable.DisposeAsync
            Return New System.Threading.Tasks.ValueTask(DisposeCoreAsync())
        End Function

        Private Async Function DisposeCoreAsync() As System.Threading.Tasks.Task
            Try
                Await StopLiveAsync()
            Catch
            End Try

            Dim senderCts As System.Threading.CancellationTokenSource = _senderCts
            If senderCts IsNot Nothing Then
                senderCts.Dispose()
                _senderCts = Nothing
            End If

            _audioSignal.Dispose()
            _sendGate.Dispose()
            _lifecycleGate.Dispose()
        End Function

    End Class

End Namespace
