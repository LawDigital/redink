Option Explicit On
Option Strict Off

Imports System.Net.WebSockets
Imports System.Text
Imports System.Threading
Imports Newtonsoft.Json.Linq

Namespace Transcription

    Public Class OpenAiRealtimeEngine
        Implements ITranscriptionEngine

        Public Const DisplayNameValue As String = "OpenAI Realtime gpt-realtime-whisper (streaming)"

        Private Const DefaultRealtimeModel As String = "gpt-realtime-whisper"
        Private Const DefaultOpenAiRealtimeUrl As String = "wss://api.openai.com/v1/realtime?model=" & DefaultRealtimeModel
        Private Const InputSampleRate As Integer = 16000
        Private Const RealtimeSampleRate As Integer = 24000
        Private Const CommitIntervalMilliseconds As Integer = 3000
        Private Const MinimumCommitBytes As Integer = RealtimeSampleRate * 2

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

        Private ReadOnly _wsUrl As String
        Private ReadOnly _apiKey As String
        Private _ws As System.Net.WebSockets.ClientWebSocket
        Private _readerTask As System.Threading.Tasks.Task
        Private _cts As System.Threading.CancellationTokenSource
        Private ReadOnly _sendLock As New System.Threading.SemaphoreSlim(1, 1)
        Private _lastCommitUtc As System.DateTime = System.DateTime.MinValue
        Private _bytesSinceCommit As Integer = 0

        Public Sub New(apiKey As String)
            _wsUrl = DefaultOpenAiRealtimeUrl
            _apiKey = apiKey
        End Sub

        Public Async Function StartLiveAsync(opts As TranscriptionOptions, ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task Implements ITranscriptionEngine.StartLiveAsync
            _cts = System.Threading.CancellationTokenSource.CreateLinkedTokenSource(ct)
            _ws = New System.Net.WebSockets.ClientWebSocket()
            _ws.Options.SetRequestHeader("Authorization", "Bearer " & _apiKey)

            Await _ws.ConnectAsync(New System.Uri(_wsUrl), _cts.Token)

            Dim lang As String = If(opts Is Nothing OrElse System.String.IsNullOrWhiteSpace(opts.LanguageCode) OrElse opts.LanguageCode.Equals("auto", System.StringComparison.OrdinalIgnoreCase), "", opts.LanguageCode.Trim())

            Dim transcription As New Newtonsoft.Json.Linq.JObject From {
                {"model", DefaultRealtimeModel},
                {"delay", "low"}
            }

            If Not System.String.IsNullOrWhiteSpace(lang) Then
                transcription("language") = lang
            End If

            Dim sessUpdate As New Newtonsoft.Json.Linq.JObject From {
                {"type", "session.update"},
                {"session", New Newtonsoft.Json.Linq.JObject From {
                    {"type", "transcription"},
                    {"audio", New Newtonsoft.Json.Linq.JObject From {
                        {"input", New Newtonsoft.Json.Linq.JObject From {
                            {"format", New Newtonsoft.Json.Linq.JObject From {
                                {"type", "audio/pcm"},
                                {"rate", RealtimeSampleRate}
                            }},
                            {"transcription", transcription},
                            {"turn_detection", Newtonsoft.Json.Linq.JValue.CreateNull()}
                        }}
                    }}
                }}
            }

            Await SendJsonAsync(sessUpdate.ToString(Newtonsoft.Json.Formatting.None))
            _lastCommitUtc = System.DateTime.UtcNow
            _bytesSinceCommit = 0
            _readerTask = System.Threading.Tasks.Task.Run(Function() ReadLoop(_cts.Token))
        End Function

        Public Async Function PushAudioAsync(pcm As Byte(), bytesValid As Integer, ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task Implements ITranscriptionEngine.PushAudioAsync
            If _ws Is Nothing OrElse _ws.State <> System.Net.WebSockets.WebSocketState.Open OrElse pcm Is Nothing OrElse bytesValid <= 0 Then
                Return
            End If

            Dim audio24k As Byte() = ResamplePcm16Mono16kTo24k(pcm, bytesValid)
            Dim b64 As String = System.Convert.ToBase64String(audio24k)
            Dim payload As New Newtonsoft.Json.Linq.JObject From {
                {"type", "input_audio_buffer.append"},
                {"audio", b64}
            }

            Await SendJsonAsync(payload.ToString(Newtonsoft.Json.Formatting.None))

            _bytesSinceCommit += audio24k.Length

            If _bytesSinceCommit >= MinimumCommitBytes AndAlso (System.DateTime.UtcNow - _lastCommitUtc).TotalMilliseconds >= CommitIntervalMilliseconds Then
                Await CommitBufferAsync()
            End If
        End Function

        Public Async Function StopLiveAsync() As System.Threading.Tasks.Task Implements ITranscriptionEngine.StopLiveAsync
            Try
                Await CommitBufferAsync()
            Catch ex As System.Exception
            End Try

            Try
                If _cts IsNot Nothing Then
                    _cts.Cancel()
                End If
            Catch ex As System.Exception
            End Try

            Try
                If _ws IsNot Nothing AndAlso _ws.State = System.Net.WebSockets.WebSocketState.Open Then
                    Await _ws.CloseAsync(System.Net.WebSockets.WebSocketCloseStatus.NormalClosure, "stop", System.Threading.CancellationToken.None)
                End If
            Catch ex As System.Exception
            End Try

            If _readerTask IsNot Nothing Then
                Try
                    Await _readerTask
                Catch ex As System.Exception
                End Try
            End If

            Try
                If _ws IsNot Nothing Then
                    _ws.Dispose()
                End If
            Catch ex As System.Exception
            End Try

            _ws = Nothing
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

            Await SendJsonAsync("{""type"":""input_audio_buffer.commit""}")
            _lastCommitUtc = System.DateTime.UtcNow
            _bytesSinceCommit = 0
        End Function

        Private Async Function SendJsonAsync(json As String) As System.Threading.Tasks.Task
            If _ws Is Nothing Then
                Return
            End If

            Await _sendLock.WaitAsync()

            Dim sendErr As System.Exception = Nothing

            Try
                Dim bytes As Byte() = System.Text.Encoding.UTF8.GetBytes(json)
                Await _ws.SendAsync(New System.ArraySegment(Of Byte)(bytes), System.Net.WebSockets.WebSocketMessageType.Text, True, _cts.Token)
            Catch ex As System.Exception
                sendErr = ex
            Finally
                _sendLock.Release()
            End Try

            If sendErr IsNot Nothing Then
                RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("OpenAI Realtime WS send: " & sendErr.Message, sendErr, False))
            End If
        End Function

        Private Async Function ReadLoop(ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task
            Dim buf(64 * 1024 - 1) As Byte

            Try
                While _ws IsNot Nothing AndAlso _ws.State = System.Net.WebSockets.WebSocketState.Open AndAlso Not ct.IsCancellationRequested
                    Dim sb As New System.Text.StringBuilder()
                    Dim r As System.Net.WebSockets.WebSocketReceiveResult

                    Do
                        r = Await _ws.ReceiveAsync(New System.ArraySegment(Of Byte)(buf), ct)
                        If r.MessageType = System.Net.WebSockets.WebSocketMessageType.Close Then
                            Return
                        End If
                        sb.Append(System.Text.Encoding.UTF8.GetString(buf, 0, r.Count))
                    Loop While Not r.EndOfMessage

                    HandleServerMessage(sb.ToString())
                End While
            Catch ex As System.OperationCanceledException
            Catch ex As System.Exception
                RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("OpenAI Realtime WS read: " & ex.Message, ex, False))
            End Try
        End Function

        Private Sub HandleServerMessage(msg As String)
            If System.String.IsNullOrWhiteSpace(msg) Then
                Return
            End If

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
                        Dim full As String = If(jo("transcript")?.ToString(), "")
                        If full.Length > 0 Then
                            RaiseEvent FinalResult(Me, New TranscriptionEventArgs(full.Trim(), True))
                        End If

                    Case "error"
                        RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("OpenAI Realtime error: " & msg, Nothing, False))
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

            Dim outputSampleCount As Integer = CInt(System.Math.Floor(inputSampleCount * (CDbl(RealtimeSampleRate) / CDbl(InputSampleRate))))
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
                Dim interpolated As Integer = CInt(System.Math.Round(CDbl(sample1) + (CDbl(sample2) - CDbl(sample1)) * fraction))

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
            Await StopLiveAsync()
            Try
                _sendLock.Dispose()
            Catch ex As System.Exception
            End Try
        End Function
    End Class

End Namespace
