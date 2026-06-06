' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: OpenAiRealtimeEngine.vb
' Purpose: Implements the ITranscriptionEngine interface for real-time
'          transcription using the OpenAI API, likely leveraging a streaming
'          or WebSocket-based connection for low-latency results.
'
' Architecture:
'  - ITranscriptionEngine Implementation: Provides concrete methods for
'    initializing, starting, and stopping the transcription process.
'  - Real-time Communication: Manages a persistent connection to the OpenAI
'    service to send audio data and receive transcription hypotheses.
'  - Audio Processing: Handles the encoding and chunking of audio data into
'    the format required by the OpenAI real-time transcription endpoint.
'  - Event Handling: Raises events for partial and final transcription results,
'    allowing the UI to update in real-time.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Net
Imports System.Net.WebSockets
Imports System.Text
Imports System.Threading
Imports Newtonsoft.Json.Linq

Namespace Transcription

    Public Class OpenAiRealtimeEngine
        Implements ITranscriptionEngine

        Public Const DisplayNameValue As String = "OpenAI Realtime Whisper (streaming)"

        Private Const DefaultRealtimeSessionModel As String = "gpt-realtime-2"
        Private Const DefaultRealtimeTranscriptionModel As String = "gpt-realtime-whisper"

        Private Const DefaultOpenAiRealtimeUrl As String = "wss://api.openai.com/v1/realtime?intent=transcription"

        Private Const InputSampleRate As Integer = 16000
        Private Const RealtimeSampleRate As Integer = 24000

        Private _stopStarted As Integer = 0

        Private Const SpeechRmsThreshold As Double = 350.0R
        Private Const MaximumUtteranceMilliseconds As Integer = 8000
        Private Const MinimumSpeechMillisecondsBeforeCommit As Integer = 500
        Private Const MinimumTotalMillisecondsBeforeCommit As Integer = 1200
        Private Const SilenceAfterSpeechMilliseconds As Integer = 900

        Private _speechBytesSinceCommit As Integer = 0
        Private _silenceBytesAfterSpeech As Integer = 0
        Private _utteranceStartedUtc As System.DateTime = System.DateTime.MinValue
        Private _sawInputTranscriptionForCurrentCommit As Boolean = False

        Private _finalTranscriptTcs As System.Threading.Tasks.TaskCompletionSource(Of Boolean)

        Private _sessionConfiguredTcs As System.Threading.Tasks.TaskCompletionSource(Of Boolean)
        Private ReadOnly _completedTranscriptKeys As New System.Collections.Generic.HashSet(Of String)(System.StringComparer.Ordinal)


        Private Shared ReadOnly _supportedLanguages As String() = {
            "auto", "ar", "bg", "ca", "cs", "da", "de", "el", "en", "es", "et", "fi", "fr",
            "he", "hi", "hr", "hu", "id", "it", "ja", "ko", "lt", "lv", "ms", "nl",
            "no", "pl", "pt", "ro", "ru", "sk", "sl", "sv", "th", "tr", "uk", "vi", "zh"
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

        Public ReadOnly Property Name As String Implements ITranscriptionEngine.Name
            Get
                Return "OpenAI Realtime"
            End Get
        End Property

        Public ReadOnly Property Kind As EngineKind Implements ITranscriptionEngine.Kind
            Get
                Return EngineKind.OpenAiRealtime
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
                Return False
            End Get
        End Property

        Public ReadOnly Property SupportsMultiChannelDiarization As Boolean Implements ITranscriptionEngine.SupportsMultiChannelDiarization
            Get
                Return False
            End Get
        End Property

        Private ReadOnly _apiKey As String
        Private _ws As System.Net.WebSockets.ClientWebSocket
        Private _readerTask As System.Threading.Tasks.Task
        Private _cts As System.Threading.CancellationTokenSource
        Private ReadOnly _sendLock As New System.Threading.SemaphoreSlim(1, 1)
        Private _lastCommitUtc As System.DateTime = System.DateTime.MinValue
        Private _bytesSinceCommit As Integer = 0

        Public Sub New(apiKey As String)
            _apiKey = apiKey
        End Sub

        Private Sub RaiseStatusMessage(message As String)
            RaiseEvent Status(Me, New TranscriptionStatusEventArgs(message))
        End Sub

        Private Shared Function NormalizeRealtimeTranscriptionModel(opts As TranscriptionOptions) As String
            Return DefaultRealtimeTranscriptionModel
        End Function

        Private Shared Function BuildRealtimeUri() As System.Uri
            Return New System.Uri(DefaultOpenAiRealtimeUrl)
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

        Private Shared Sub EnsureTls12()
            Try
                AppContext.SetSwitch("Switch.System.Net.DontEnableSchUseStrongCrypto", False)
            Catch
            End Try

            Try
                AppContext.SetSwitch("Switch.System.Net.DontEnableSystemDefaultTlsVersions", False)
            Catch
            End Try

            Try
                ServicePointManager.Expect100Continue = False
            Catch
            End Try

            Try
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
            Catch
                Try
                    ServicePointManager.SecurityProtocol = CType(3072, SecurityProtocolType)
                Catch
                End Try
            End Try
        End Sub

        Private Function CreateWebSocket() As System.Net.WebSockets.ClientWebSocket
            Dim ws As New System.Net.WebSockets.ClientWebSocket()
            ws.Options.SetRequestHeader("Authorization", "Bearer " & _apiKey)
            ws.Options.KeepAliveInterval = System.TimeSpan.FromSeconds(20)
            Return ws
        End Function

        Public Async Function StartLiveAsync(opts As TranscriptionOptions, ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task Implements ITranscriptionEngine.StartLiveAsync
            EnsureTls12()

            System.Threading.Interlocked.Exchange(_stopStarted, 0)
            ResetFinalTranscriptWait()

            _sessionConfiguredTcs = New System.Threading.Tasks.TaskCompletionSource(Of Boolean)()

            SyncLock _completedTranscriptKeys
                _completedTranscriptKeys.Clear()
            End SyncLock

            Dim transcriptionModel As String = NormalizeRealtimeTranscriptionModel(opts)

            _cts = System.Threading.CancellationTokenSource.CreateLinkedTokenSource(ct)
            _ws = CreateWebSocket()

            Try
                RaiseStatusMessage("Connecting to OpenAI Realtime transcription…")
                Await _ws.ConnectAsync(BuildRealtimeUri(), _cts.Token).ConfigureAwait(False)
            Catch ex As System.Exception
                Dim detail As String = "OpenAI Realtime connection failed: " & GetDetailedExceptionMessage(ex)

                Try
                    _ws.Dispose()
                Catch
                End Try

                _ws = Nothing
                RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs(detail, ex, True))
                Throw New System.InvalidOperationException(detail, ex)
            End Try

            RaiseStatusMessage("Transcription model: " & transcriptionModel)

            _readerTask = System.Threading.Tasks.Task.Run(Function() ReadLoop(_cts.Token))

            Dim configuredLanguage As String = If(opts IsNot Nothing, If(opts.LanguageCode, "").Trim(), "")

            Dim transcriptionSettings As New Newtonsoft.Json.Linq.JObject From {
    {"model", transcriptionModel},
    {"delay", "low"}
}

            If Not String.IsNullOrWhiteSpace(configuredLanguage) AndAlso
   Not String.Equals(configuredLanguage, "auto", StringComparison.OrdinalIgnoreCase) Then
                transcriptionSettings("language") = configuredLanguage
            End If

            Dim inputAudio As New Newtonsoft.Json.Linq.JObject From {
    {"format", New Newtonsoft.Json.Linq.JObject From {
        {"type", "audio/pcm"},
        {"rate", RealtimeSampleRate}
    }},
    {"transcription", transcriptionSettings}
}

            inputAudio("turn_detection") = Newtonsoft.Json.Linq.JValue.CreateNull()

            Dim sessionUpdate As New Newtonsoft.Json.Linq.JObject From {
        {"type", "session.update"},
        {"session", New Newtonsoft.Json.Linq.JObject From {
            {"type", "transcription"},
            {"audio", New Newtonsoft.Json.Linq.JObject From {
                {"input", inputAudio}
            }}
        }}
    }

            RaiseStatusMessage("Configuring OpenAI Realtime transcription session…")
            Await SendJsonAsync(sessionUpdate, _cts.Token).ConfigureAwait(False)

            Try
                Dim configuredTask As System.Threading.Tasks.Task = _sessionConfiguredTcs.Task

                Dim completedTask As System.Threading.Tasks.Task =
            Await System.Threading.Tasks.Task.WhenAny(
                configuredTask,
                System.Threading.Tasks.Task.Delay(3000, _cts.Token)).ConfigureAwait(False)

                If completedTask IsNot configuredTask Then
                    RaiseStatusMessage("OpenAI Realtime transcription session was not confirmed within 3 seconds.")
                End If
            Catch ex As System.OperationCanceledException
            End Try

            _lastCommitUtc = System.DateTime.UtcNow
            _bytesSinceCommit = 0
            _speechBytesSinceCommit = 0
            _silenceBytesAfterSpeech = 0
            _utteranceStartedUtc = System.DateTime.MinValue
            _sawInputTranscriptionForCurrentCommit = False
        End Function

        Public Async Function PushAudioAsync(pcm As Byte(), bytesValid As Integer, ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task Implements ITranscriptionEngine.PushAudioAsync
            If _ws Is Nothing OrElse _ws.State <> System.Net.WebSockets.WebSocketState.Open OrElse pcm Is Nothing OrElse bytesValid <= 0 Then
                Return
            End If

            Dim evenBytesValid As Integer = bytesValid - (bytesValid Mod 2)
            If evenBytesValid <= 0 Then
                Return
            End If

            Dim audio24k As Byte() = ResamplePcm16Mono16kTo24k(pcm, evenBytesValid)
            If audio24k Is Nothing OrElse audio24k.Length = 0 Then
                Return
            End If

            Dim hasSpeech As Boolean = Pcm16HasSpeech(audio24k, audio24k.Length, SpeechRmsThreshold)

            If hasSpeech Then
                If _speechBytesSinceCommit <= 0 Then
                    _utteranceStartedUtc = System.DateTime.UtcNow
                End If

                Await AppendAudioAsync(audio24k, ct).ConfigureAwait(False)

                _bytesSinceCommit += audio24k.Length
                _speechBytesSinceCommit += audio24k.Length
                _silenceBytesAfterSpeech = 0
            Else
                ' Ignore leading silence completely. This prevents "empty/no speech" commits.
                If _speechBytesSinceCommit <= 0 Then
                    Return
                End If

                ' Keep trailing silence after speech so the model gets a natural utterance ending.
                Await AppendAudioAsync(audio24k, ct).ConfigureAwait(False)

                _bytesSinceCommit += audio24k.Length
                _silenceBytesAfterSpeech += audio24k.Length
            End If

            If ShouldCommitCurrentUtterance() Then
                Await CommitBufferAsync().ConfigureAwait(False)
            End If
        End Function

        Private Async Function AppendAudioAsync(audio24k As Byte(), ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task
            If audio24k Is Nothing OrElse audio24k.Length = 0 Then
                Return
            End If

            If _ws Is Nothing OrElse _ws.State <> System.Net.WebSockets.WebSocketState.Open Then
                Return
            End If

            Dim b64 As String = System.Convert.ToBase64String(audio24k)

            Dim payload As New Newtonsoft.Json.Linq.JObject From {
        {"type", "input_audio_buffer.append"},
        {"audio", b64}
    }

            Await SendJsonAsync(payload, ct).ConfigureAwait(False)
        End Function

        Private Function ShouldCommitCurrentUtterance() As Boolean
            If _bytesSinceCommit <= 0 Then
                Return False
            End If

            If _speechBytesSinceCommit <= 0 Then
                Return False
            End If

            Dim minimumSpeechBytes As Integer =
        CInt((CDbl(RealtimeSampleRate) * 2.0R) * (CDbl(MinimumSpeechMillisecondsBeforeCommit) / 1000.0R))

            If _speechBytesSinceCommit < minimumSpeechBytes Then
                Return False
            End If

            Dim minimumTotalBytes As Integer =
        CInt((CDbl(RealtimeSampleRate) * 2.0R) * (CDbl(MinimumTotalMillisecondsBeforeCommit) / 1000.0R))

            If _bytesSinceCommit < minimumTotalBytes Then
                Return False
            End If

            Dim silenceAfterSpeechBytes As Integer =
        CInt((CDbl(RealtimeSampleRate) * 2.0R) * (CDbl(SilenceAfterSpeechMilliseconds) / 1000.0R))

            If _silenceBytesAfterSpeech >= silenceAfterSpeechBytes Then
                Return True
            End If

            If _utteranceStartedUtc <> System.DateTime.MinValue AndAlso
       (System.DateTime.UtcNow - _utteranceStartedUtc).TotalMilliseconds >= MaximumUtteranceMilliseconds Then
                Return True
            End If

            Return False
        End Function



        Private Shared Function Pcm16HasSpeech(pcm As Byte(), bytesValid As Integer, rmsThreshold As Double) As Boolean
            If pcm Is Nothing OrElse bytesValid < 2 Then
                Return False
            End If

            Dim evenBytesValid As Integer = bytesValid - (bytesValid Mod 2)
            If evenBytesValid <= 0 Then
                Return False
            End If

            Dim sampleCount As Integer = evenBytesValid \ 2
            If sampleCount <= 0 Then
                Return False
            End If

            Dim sumSquares As Double = 0.0R

            For i As Integer = 0 To sampleCount - 1
                Dim sample As Short = System.BitConverter.ToInt16(pcm, i * 2)
                Dim sampleDouble As Double = CDbl(sample)
                sumSquares += sampleDouble * sampleDouble
            Next

            Dim rms As Double = System.Math.Sqrt(sumSquares / CDbl(sampleCount))

            Return rms >= rmsThreshold
        End Function

        Public Async Function StopLiveAsync() As System.Threading.Tasks.Task Implements ITranscriptionEngine.StopLiveAsync
            If System.Threading.Interlocked.Exchange(_stopStarted, 1) <> 0 Then
                Return
            End If

            Dim wsToClose As System.Net.WebSockets.ClientWebSocket = _ws
            Dim readerToAwait As System.Threading.Tasks.Task = _readerTask
            Dim ctsToCancel As System.Threading.CancellationTokenSource = _cts
            Dim finalTranscriptTask As System.Threading.Tasks.Task = Nothing

            Try
                Dim minimumTotalBytes As Integer =
    CInt((CDbl(RealtimeSampleRate) * 2.0R) * (CDbl(MinimumTotalMillisecondsBeforeCommit) / 1000.0R))

                If _bytesSinceCommit >= minimumTotalBytes AndAlso _speechBytesSinceCommit > 0 Then
                    finalTranscriptTask = BeginFinalTranscriptWait()
                    Await CommitBufferAsync().ConfigureAwait(False)

                    Try
                        Await System.Threading.Tasks.Task.WhenAny(
                    finalTranscriptTask,
                    System.Threading.Tasks.Task.Delay(2500)).ConfigureAwait(False)
                    Catch
                    End Try
                End If
            Catch
            End Try

            Try
                If wsToClose IsNot Nothing AndAlso wsToClose.State = System.Net.WebSockets.WebSocketState.Open Then
                    Try
                        Await wsToClose.CloseOutputAsync(
                    System.Net.WebSockets.WebSocketCloseStatus.NormalClosure,
                    "stop",
                    System.Threading.CancellationToken.None).ConfigureAwait(False)
                    Catch
                    End Try
                End If
            Catch
            End Try

            If readerToAwait IsNot Nothing Then
                Try
                    Await System.Threading.Tasks.Task.WhenAny(
                readerToAwait,
                System.Threading.Tasks.Task.Delay(1500)).ConfigureAwait(False)
                Catch
                End Try
            End If

            Try
                If ctsToCancel IsNot Nothing Then
                    ctsToCancel.Cancel()
                End If
            Catch
            End Try

            If readerToAwait IsNot Nothing Then
                Try
                    Await readerToAwait.ConfigureAwait(False)
                Catch
                End Try
            End If

            Try
                If wsToClose IsNot Nothing Then
                    wsToClose.Dispose()
                End If
            Catch
            End Try

            If ctsToCancel IsNot Nothing Then
                Try
                    ctsToCancel.Dispose()
                Catch
                End Try
            End If

            _ws = Nothing
            _readerTask = Nothing
            _cts = Nothing
            _bytesSinceCommit = 0
            _speechBytesSinceCommit = 0
            _silenceBytesAfterSpeech = 0
            _lastCommitUtc = System.DateTime.MinValue
            _utteranceStartedUtc = System.DateTime.MinValue
            _finalTranscriptTcs = Nothing

            RaiseStatusMessage("OpenAI Realtime stopped.")
        End Function


        Public Function TranscribeFileAsync(filePath As String, opts As TranscriptionOptions, ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task Implements ITranscriptionEngine.TranscribeFileAsync
            Throw New System.NotSupportedException("OpenAI Realtime is live-only. Use OpenAI REST for files.")
        End Function

        Private Async Function CommitBufferAsync() As System.Threading.Tasks.Task
            If _ws Is Nothing OrElse _ws.State <> System.Net.WebSockets.WebSocketState.Open Then
                Return
            End If

            If _bytesSinceCommit <= 0 OrElse _speechBytesSinceCommit <= 0 Then
                ResetAudioCommitState()
                Return
            End If

            Dim minimumTotalBytes As Integer =
        CInt((CDbl(RealtimeSampleRate) * 2.0R) * (CDbl(MinimumTotalMillisecondsBeforeCommit) / 1000.0R))

            If _bytesSinceCommit < minimumTotalBytes Then
                RaiseStatusMessage("OpenAI Realtime: audio buffer not committed because it is too short.")
                ResetAudioCommitState()
                Return
            End If

            _sawInputTranscriptionForCurrentCommit = False

            Dim payload As New Newtonsoft.Json.Linq.JObject From {
        {"type", "input_audio_buffer.commit"}
    }

            Await SendJsonAsync(payload, GetSendCancellationToken(System.Threading.CancellationToken.None)).ConfigureAwait(False)

            _lastCommitUtc = System.DateTime.UtcNow

            ResetAudioCommitState()

            RaiseStatusMessage("OpenAI Realtime speech audio committed.")
        End Function

        Private Sub ResetAudioCommitState()
            _bytesSinceCommit = 0
            _speechBytesSinceCommit = 0
            _silenceBytesAfterSpeech = 0
            _utteranceStartedUtc = System.DateTime.MinValue
        End Sub


        Private Function GetSendCancellationToken(fallback As System.Threading.CancellationToken) As System.Threading.CancellationToken
            If _cts IsNot Nothing Then
                Return _cts.Token
            End If

            Return fallback
        End Function

        Private Function IsStoppingOrCanceled(ct As System.Threading.CancellationToken) As Boolean
            If ct.IsCancellationRequested Then
                Return True
            End If

            If _cts IsNot Nothing AndAlso _cts.IsCancellationRequested Then
                Return True
            End If

            Return System.Threading.Interlocked.CompareExchange(_stopStarted, 0, 0) <> 0
        End Function

        Private Async Function SendJsonAsync(payload As Newtonsoft.Json.Linq.JObject, ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task
            If payload Is Nothing Then
                Return
            End If

            Await SendJsonAsync(payload.ToString(Newtonsoft.Json.Formatting.None), ct).ConfigureAwait(False)
        End Function

        Private Async Function SendJsonAsync(json As String, ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task
            If _ws Is Nothing OrElse _ws.State <> System.Net.WebSockets.WebSocketState.Open Then
                Return
            End If

            Try
                Dim outgoingType As String = ""

                Try
                    Dim outgoing As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(json)
                    outgoingType = If(outgoing("type")?.ToString(), "")
                Catch
                End Try

                If outgoingType.Length > 0 Then
                    System.Diagnostics.Debug.WriteLine("[OpenAI Realtime OUT] " & outgoingType & ": " & RedactRealtimeJsonForLog(json))
                Else
                    System.Diagnostics.Debug.WriteLine("[OpenAI Realtime OUT] " & RedactRealtimeJsonForLog(json))
                End If
            Catch
            End Try

            Dim sendCt As System.Threading.CancellationToken = GetSendCancellationToken(ct)
            Dim lockTaken As Boolean = False

            Try
                Await _sendLock.WaitAsync(sendCt).ConfigureAwait(False)
                lockTaken = True

                If _ws Is Nothing OrElse _ws.State <> System.Net.WebSockets.WebSocketState.Open Then
                    Return
                End If

                Dim bytes As Byte() = System.Text.Encoding.UTF8.GetBytes(json)

                Await _ws.SendAsync(
            New System.ArraySegment(Of Byte)(bytes),
            System.Net.WebSockets.WebSocketMessageType.Text,
            True,
            sendCt).ConfigureAwait(False)

            Catch ex As System.OperationCanceledException
                If IsStoppingOrCanceled(sendCt) Then
                    Return
                End If

                Dim detail As String = "OpenAI Realtime WS send failed: " & GetDetailedExceptionMessage(ex)
                RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs(detail, ex, False))
                Throw New System.InvalidOperationException(detail, ex)

            Catch ex As System.ObjectDisposedException
                If IsStoppingOrCanceled(sendCt) Then
                    Return
                End If

                Dim detail As String = "OpenAI Realtime WS send failed: " & GetDetailedExceptionMessage(ex)
                RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs(detail, ex, False))
                Throw New System.InvalidOperationException(detail, ex)

            Catch ex As System.Exception
                If IsStoppingOrCanceled(sendCt) OrElse
           _ws Is Nothing OrElse
           _ws.State = System.Net.WebSockets.WebSocketState.Aborted OrElse
           _ws.State = System.Net.WebSockets.WebSocketState.CloseReceived OrElse
           _ws.State = System.Net.WebSockets.WebSocketState.CloseSent OrElse
           _ws.State = System.Net.WebSockets.WebSocketState.Closed Then
                    Return
                End If

                Dim detail As String = "OpenAI Realtime WS send failed: " & GetDetailedExceptionMessage(ex)
                RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs(detail, ex, False))
                Throw New System.InvalidOperationException(detail, ex)
            Finally
                If lockTaken Then
                    _sendLock.Release()
                End If
            End Try
        End Function


        Private Shared Function RedactRealtimeJsonForLog(json As String) As String
            If System.String.IsNullOrWhiteSpace(json) Then
                Return json
            End If

            Try
                Dim jo As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(json)

                RedactJsonTokenForLog(jo)

                Return jo.ToString(Newtonsoft.Json.Formatting.None)
            Catch ex As Newtonsoft.Json.JsonException
                Return json
            End Try
        End Function

        Private Shared Sub RedactJsonTokenForLog(token As Newtonsoft.Json.Linq.JToken)
            If token Is Nothing Then
                Return
            End If

            If token.Type = Newtonsoft.Json.Linq.JTokenType.Object Then
                Dim obj As Newtonsoft.Json.Linq.JObject = DirectCast(token, Newtonsoft.Json.Linq.JObject)

                For Each prop As Newtonsoft.Json.Linq.JProperty In obj.Properties().ToList()
                    Dim name As String = prop.Name

                    If System.String.Equals(name, "audio", System.StringComparison.OrdinalIgnoreCase) OrElse
               System.String.Equals(name, "delta", System.StringComparison.OrdinalIgnoreCase) Then

                        Dim valueText As String = If(prop.Value IsNot Nothing, prop.Value.ToString(), "")
                        Dim valueLength As Integer = valueText.Length

                        If valueLength > 80 Then
                            prop.Value = "[redacted " & valueLength.ToString(System.Globalization.CultureInfo.InvariantCulture) & " chars]"
                        End If
                    Else
                        RedactJsonTokenForLog(prop.Value)
                    End If
                Next

                Return
            End If

            If token.Type = Newtonsoft.Json.Linq.JTokenType.Array Then
                For Each child As Newtonsoft.Json.Linq.JToken In token.Children()
                    RedactJsonTokenForLog(child)
                Next
            End If
        End Sub


        Private Async Function ReadLoop(ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task
            Dim buf(64 * 1024 - 1) As Byte

            Try
                While _ws IsNot Nothing AndAlso _ws.State = System.Net.WebSockets.WebSocketState.Open AndAlso Not ct.IsCancellationRequested
                    Dim sb As New System.Text.StringBuilder()
                    Dim r As System.Net.WebSockets.WebSocketReceiveResult

                    Do
                        r = Await _ws.ReceiveAsync(New System.ArraySegment(Of Byte)(buf), ct).ConfigureAwait(False)

                        If r.MessageType = System.Net.WebSockets.WebSocketMessageType.Close Then
                            Return
                        End If

                        sb.Append(System.Text.Encoding.UTF8.GetString(buf, 0, r.Count))
                    Loop While Not r.EndOfMessage

                    HandleServerMessage(sb.ToString())
                End While
            Catch ex As System.OperationCanceledException
            Catch ex As System.Exception
                RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("OpenAI Realtime WS read failed: " & GetDetailedExceptionMessage(ex), ex, False))
            End Try
        End Function


        Private Sub ResetFinalTranscriptWait()
            _finalTranscriptTcs = New System.Threading.Tasks.TaskCompletionSource(Of Boolean)()
        End Sub

        Private Function BeginFinalTranscriptWait() As System.Threading.Tasks.Task
            ResetFinalTranscriptWait()
            Return _finalTranscriptTcs.Task
        End Function

        Private Sub CompleteFinalTranscriptWait()
            If _finalTranscriptTcs IsNot Nothing Then
                Try
                    _finalTranscriptTcs.TrySetResult(True)
                Catch
                End Try
            End If
        End Sub

        Private Shared Function IsIgnorableRealtimeError(detail As String) As Boolean
            Dim normalized As String = If(detail, "").Trim().ToLowerInvariant()

            If normalized.Length = 0 Then
                Return False
            End If

            Return normalized.Contains("buffer is empty") OrElse
           normalized.Contains("audio buffer is empty") OrElse
           normalized.Contains("input_audio_buffer_commit_empty") OrElse
           normalized.Contains("buffer too small") OrElse
           normalized.Contains("buffer is too small") OrElse
           normalized.Contains("audio buffer too small") OrElse
           normalized.Contains("audio buffer is too small") OrElse
           normalized.Contains("input_audio_buffer_commit_too_small") OrElse
           normalized.Contains("no audio") OrElse
           normalized.Contains("no speech") OrElse
           normalized.Contains("too short") OrElse
           normalized.Contains("silence")
        End Function

        Private Sub HandleServerMessage(msg As String)
            If System.String.IsNullOrWhiteSpace(msg) Then
                Return
            End If

            System.Diagnostics.Debug.WriteLine("[OpenAI Realtime IN] " & RedactRealtimeJsonForLog(msg))

            Try
                Dim jo As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(msg)
                Dim eventType As String = If(jo("type")?.ToString(), "")

                Select Case eventType
                    Case "conversation.item.input_audio_transcription.delta"
                        _sawInputTranscriptionForCurrentCommit = True

                        Dim delta As String = If(jo("delta")?.ToString(), "")
                        If delta.Length > 0 Then
                            RaiseEvent PartialResult(Me, New TranscriptionEventArgs(delta, False))
                        End If

                    Case "conversation.item.input_audio_transcription.completed"
                        _sawInputTranscriptionForCurrentCommit = True

                        Dim itemId As String = If(jo("item_id")?.ToString(), "")
                        Dim contentIndex As String = If(jo("content_index")?.ToString(), "")
                        Dim transcript As String = If(jo("transcript")?.ToString(), "")

                        CompleteFinalTranscriptWait()

                        If transcript.Length > 0 Then
                            Dim key As String = itemId & "|" & contentIndex & "|" & transcript

                            Dim shouldRaise As Boolean = False

                            SyncLock _completedTranscriptKeys
                                If Not _completedTranscriptKeys.Contains(key) Then
                                    _completedTranscriptKeys.Add(key)
                                    shouldRaise = True
                                End If
                            End SyncLock

                            If shouldRaise Then
                                RaiseEvent FinalResult(Me, New TranscriptionEventArgs(transcript, True))
                            Else
                                System.Diagnostics.Debug.WriteLine("[OpenAI Realtime] Duplicate input transcript ignored: " & transcript)
                            End If
                        Else
                            RaiseStatusMessage("OpenAI Realtime: transcription completed without text.")
                        End If

                    Case "conversation.item.input_audio_transcription.failed"
                        CompleteFinalTranscriptWait()

                        Dim detail As String = ""
                        Dim errorToken As Newtonsoft.Json.Linq.JToken = jo("error")

                        If errorToken IsNot Nothing AndAlso errorToken.Type = Newtonsoft.Json.Linq.JTokenType.Object Then
                            detail = If(errorToken("message")?.ToString(), msg)
                        Else
                            detail = msg
                        End If

                        If IsIgnorableRealtimeError(detail) Then
                            RaiseStatusMessage("OpenAI Realtime: no usable speech detected.")
                            Return
                        End If

                        RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("OpenAI Realtime transcription failed: " & detail, Nothing, False))

                    Case "session.created", "session.updated"
                        Dim sessionObject As Newtonsoft.Json.Linq.JObject = TryCast(jo("session"), Newtonsoft.Json.Linq.JObject)
                        Dim sessionType As String = If(sessionObject IsNot Nothing, If(sessionObject("type")?.ToString(), ""), "")

                        If System.String.Equals(sessionType, "transcription", System.StringComparison.OrdinalIgnoreCase) Then
                            If _sessionConfiguredTcs IsNot Nothing Then
                                Try
                                    _sessionConfiguredTcs.TrySetResult(True)
                                Catch
                                End Try
                            End If

                            RaiseStatusMessage("OpenAI Realtime transcription session configured.")
                        ElseIf System.String.IsNullOrWhiteSpace(sessionType) Then
                            RaiseStatusMessage("OpenAI Realtime session configured.")
                        Else
                            RaiseStatusMessage("OpenAI Realtime session configured (" & sessionType & ").")
                        End If

                    Case "input_audio_buffer.committed"
                        RaiseStatusMessage("OpenAI Realtime audio buffer committed.")

                    Case "input_audio_buffer.cleared"
                        RaiseStatusMessage("OpenAI Realtime audio buffer cleared.")

                    Case "input_audio_buffer.speech_started"
                        RaiseStatusMessage("OpenAI Realtime server speech started.")

                    Case "input_audio_buffer.speech_stopped"
                        RaiseStatusMessage("OpenAI Realtime server speech stopped.")

                    Case "response.created",
                 "response.output_item.added",
                 "response.content_part.added",
                 "response.content_part.done",
                 "response.output_audio.delta",
                 "response.output_audio.done",
                 "response.output_audio_transcript.delta",
                 "response.output_audio_transcript.done",
                 "response.done"
                        ' Do not raise PartialResult or FinalResult here.
                        ' These are assistant-output events, not dictation transcript events.
                        System.Diagnostics.Debug.WriteLine("[OpenAI Realtime] Unexpected assistant-output event ignored: " & eventType)

                    Case "error"
                        Dim errToken As Newtonsoft.Json.Linq.JToken = jo("error")
                        Dim detail As String = ""

                        If errToken IsNot Nothing AndAlso errToken.Type = Newtonsoft.Json.Linq.JTokenType.Object Then
                            detail = If(errToken("message")?.ToString(), msg)
                        Else
                            detail = msg
                        End If

                        If IsIgnorableRealtimeError(detail) Then
                            CompleteFinalTranscriptWait()
                            RaiseStatusMessage("OpenAI Realtime: audio buffer was empty or too small; ignored.")
                            Return
                        End If

                        RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("OpenAI Realtime error: " & detail, Nothing, False))

                    Case Else
                        System.Diagnostics.Debug.WriteLine("[OpenAI Realtime] Unhandled event type: " & eventType)
                End Select
            Catch ex As Newtonsoft.Json.JsonException
                RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("OpenAI Realtime invalid JSON: " & msg, ex, False))
            End Try
        End Sub

        Private Shared Function ResamplePcm16Mono16kTo24k(pcm As Byte(), bytesValid As Integer) As Byte()
            Dim inputSampleCount As Integer = bytesValid \ 2
            If inputSampleCount <= 1 Then
                Return New Byte() {}
            End If

            Dim outputSampleCount As Integer =
                CInt(System.Math.Floor(inputSampleCount * (CDbl(RealtimeSampleRate) / CDbl(InputSampleRate))))

            Dim outputBytes(outputSampleCount * 2 - 1) As Byte

            For i As Integer = 0 To outputSampleCount - 1
                Dim sourcePosition As Double = CDbl(i) * CDbl(InputSampleRate) / CDbl(RealtimeSampleRate)
                Dim sourceIndex As Integer = CInt(System.Math.Floor(sourcePosition))
                Dim fraction As Double = sourcePosition - CDbl(sourceIndex)

                If sourceIndex >= inputSampleCount - 1 Then
                    sourceIndex = inputSampleCount - 2
                    fraction = 1.0R
                End If

                Dim sample1 As Short = System.BitConverter.ToInt16(pcm, sourceIndex * 2)
                Dim sample2 As Short = System.BitConverter.ToInt16(pcm, (sourceIndex + 1) * 2)

                Dim interpolated As Integer =
                    CInt(System.Math.Round(CDbl(sample1) + (CDbl(sample2) - CDbl(sample1)) * fraction))

                If interpolated > Short.MaxValue Then
                    interpolated = Short.MaxValue
                ElseIf interpolated < Short.MinValue Then
                    interpolated = Short.MinValue
                End If

                Dim bytes As Byte() = System.BitConverter.GetBytes(CShort(interpolated))
                outputBytes(i * 2) = bytes(0)
                outputBytes(i * 2 + 1) = bytes(1)
            Next

            Return outputBytes
        End Function

        Public Function DisposeAsync() As System.Threading.Tasks.ValueTask Implements System.IAsyncDisposable.DisposeAsync
            Return New System.Threading.Tasks.ValueTask(DisposeAsyncCore())
        End Function

        Private Async Function DisposeAsyncCore() As System.Threading.Tasks.Task
            Await StopLiveAsync().ConfigureAwait(False)

            Try
                _sendLock.Dispose()
            Catch
            End Try
        End Function
    End Class

End Namespace