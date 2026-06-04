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

        Public Const DisplayNameValue As String = "OpenAI Realtime gpt-realtime-2 / gpt-realtime-whisper (streaming)"

        Private Const DefaultRealtimeSessionModel As String = "gpt-realtime-2"
        Private Const DefaultRealtimeTranscriptionModel As String = "gpt-realtime-whisper"

        Private Const DefaultOpenAiRealtimeUrl As String =
            "wss://api.openai.com/v1/realtime?model=" & DefaultRealtimeSessionModel

        Private Const InputSampleRate As Integer = 16000
        Private Const RealtimeSampleRate As Integer = 24000
        Private Const CommitIntervalMilliseconds As Integer = 1000
        Private Const MinimumCommitBytes As Integer = RealtimeSampleRate * 2

        Private _stopStarted As Integer = 0

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

            Dim transcriptionModel As String = NormalizeRealtimeTranscriptionModel(opts)

            _cts = System.Threading.CancellationTokenSource.CreateLinkedTokenSource(ct)
            _ws = CreateWebSocket()

            Try
                RaiseStatusMessage("Connecting to OpenAI Realtime (" & DefaultRealtimeSessionModel & ")…")
                RaiseStatusMessage("Using TLS 1.2 for OpenAI Realtime connection…")
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

            RaiseStatusMessage("OpenAI Realtime WebSocket connected.")
            RaiseStatusMessage("OpenAI Realtime session model: " & DefaultRealtimeSessionModel)

            _readerTask = System.Threading.Tasks.Task.Run(Function() ReadLoop(_cts.Token))

            Dim inputAudio As New Newtonsoft.Json.Linq.JObject From {
                {"format", New Newtonsoft.Json.Linq.JObject From {
                    {"type", "audio/pcm"},
                    {"rate", 24000}
                }},
                {"transcription", New Newtonsoft.Json.Linq.JObject From {
                    {"model", DefaultRealtimeTranscriptionModel},
                    {"delay", "low"}
                }}
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

            _lastCommitUtc = System.DateTime.UtcNow
            _bytesSinceCommit = 0
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

            Dim b64 As String = System.Convert.ToBase64String(audio24k)

            Dim payload As New Newtonsoft.Json.Linq.JObject From {
                {"type", "input_audio_buffer.append"},
                {"audio", b64}
            }

            Await SendJsonAsync(payload, ct).ConfigureAwait(False)

            _bytesSinceCommit += audio24k.Length

            If _bytesSinceCommit >= MinimumCommitBytes AndAlso
               (System.DateTime.UtcNow - _lastCommitUtc).TotalMilliseconds >= CommitIntervalMilliseconds Then
                Await CommitBufferAsync().ConfigureAwait(False)
            End If
        End Function

        Public Async Function StopLiveAsync() As System.Threading.Tasks.Task Implements ITranscriptionEngine.StopLiveAsync
            If System.Threading.Interlocked.Exchange(_stopStarted, 1) <> 0 Then
                Return
            End If

            Dim wsToClose As System.Net.WebSockets.ClientWebSocket = _ws
            Dim readerToAwait As System.Threading.Tasks.Task = _readerTask
            Dim ctsToCancel As System.Threading.CancellationTokenSource = _cts

            Try
                Await CommitBufferAsync().ConfigureAwait(False)
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
            _lastCommitUtc = System.DateTime.MinValue

            RaiseStatusMessage("OpenAI Realtime stopped.")
        End Function

        Public Function TranscribeFileAsync(filePath As String, opts As TranscriptionOptions, ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task Implements ITranscriptionEngine.TranscribeFileAsync
            Throw New System.NotSupportedException("OpenAI Realtime is live-only. Use OpenAI REST for files.")
        End Function

        Private Async Function CommitBufferAsync() As System.Threading.Tasks.Task
            If _ws Is Nothing OrElse _ws.State <> System.Net.WebSockets.WebSocketState.Open Then
                Return
            End If

            If _bytesSinceCommit <= 0 Then
                Return
            End If

            Dim payload As New Newtonsoft.Json.Linq.JObject From {
                {"type", "input_audio_buffer.commit"}
            }

            Await SendJsonAsync(payload, GetSendCancellationToken(System.Threading.CancellationToken.None)).ConfigureAwait(False)

            _lastCommitUtc = System.DateTime.UtcNow
            _bytesSinceCommit = 0
            RaiseStatusMessage("OpenAI Realtime audio committed.")
        End Function

        Private Function GetSendCancellationToken(fallback As System.Threading.CancellationToken) As System.Threading.CancellationToken
            If _cts IsNot Nothing Then
                Return _cts.Token
            End If

            Return fallback
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

            Dim sendCt As System.Threading.CancellationToken = GetSendCancellationToken(ct)

            Await _sendLock.WaitAsync(sendCt).ConfigureAwait(False)

            Try
                Dim bytes As Byte() = System.Text.Encoding.UTF8.GetBytes(json)
                Await _ws.SendAsync(
                    New System.ArraySegment(Of Byte)(bytes),
                    System.Net.WebSockets.WebSocketMessageType.Text,
                    True,
                    sendCt).ConfigureAwait(False)
            Catch ex As System.Exception
                Dim detail As String = "OpenAI Realtime WS send failed: " & GetDetailedExceptionMessage(ex)
                RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs(detail, ex, False))
                Throw New System.InvalidOperationException(detail, ex)
            Finally
                _sendLock.Release()
            End Try
        End Function

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

        Private Sub HandleServerMessage(msg As String)
            If System.String.IsNullOrWhiteSpace(msg) Then
                Return
            End If

            System.Diagnostics.Debug.WriteLine("[OpenAI Realtime] " & msg)

            Try
                Dim jo As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(msg)
                Dim eventType As String = If(jo("type")?.ToString(), "")

                Select Case eventType
                    Case "conversation.item.input_audio_transcription.delta"
                        Dim delta As String = If(jo("delta")?.ToString(), "")
                        If delta.Length > 0 Then
                            RaiseEvent PartialResult(Me, New TranscriptionEventArgs(delta, False))
                        End If

                    Case "conversation.item.input_audio_transcription.completed"
                        Dim transcript As String = If(jo("transcript")?.ToString(), "")
                        If transcript.Length > 0 Then
                            RaiseEvent FinalResult(Me, New TranscriptionEventArgs(transcript, True))
                        End If

                    Case "session.created", "session.updated"
                        Dim sessionObject As Newtonsoft.Json.Linq.JObject = TryCast(jo("session"), Newtonsoft.Json.Linq.JObject)
                        Dim sessionType As String = If(sessionObject IsNot Nothing, If(sessionObject("type")?.ToString(), ""), "")

                        If System.String.IsNullOrWhiteSpace(sessionType) Then
                            RaiseStatusMessage("OpenAI Realtime session configured.")
                        Else
                            RaiseStatusMessage("OpenAI Realtime session configured (" & sessionType & ").")
                        End If

                    Case "input_audio_buffer.committed"
                        RaiseStatusMessage("OpenAI Realtime audio committed.")

                    Case "response.created",
                         "response.output_item.added",
                         "response.content_part.added",
                         "response.output_audio.delta",
                         "response.output_audio_transcript.delta",
                         "response.done"
                        ' Ignore assistant/voice-output events for dictation.

                    Case "error"
                        Dim errToken As Newtonsoft.Json.Linq.JToken = jo("error")
                        Dim detail As String = ""

                        If errToken IsNot Nothing AndAlso errToken.Type = Newtonsoft.Json.Linq.JTokenType.Object Then
                            detail = If(errToken("message")?.ToString(), msg)
                        Else
                            detail = msg
                        End If

                        RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("OpenAI Realtime error: " & detail, Nothing, False))
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