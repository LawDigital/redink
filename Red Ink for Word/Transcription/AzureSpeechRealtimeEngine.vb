' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: AzureSpeechRealtimeEngine.vb
' Purpose: Implements a transcription engine using Azure's real-time
'          Speech-to-Text service for continuous transcription.
'
' Architecture:
'  - Real-time Communication: Utilizes the Azure Speech SDK to establish a
'    persistent connection for real-time audio streaming.
'  - Event-Driven Processing: Subscribes to events from the Speech SDK to
'    receive partial and final transcription results as they are generated.
'  - Audio Stream Management: Pushes audio data from the AudioCaptureService
'    into the Speech SDK's audio stream.
'  - Lifecycle Control: Manages the start and stop of the continuous
'    recognition session.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Net
Imports System.Net.WebSockets
Imports System.Text
Imports System.Threading
Imports System.Linq
Imports Newtonsoft.Json.Linq

Namespace Transcription

    ''' <summary>
    ''' Azure Speech-to-Text live streaming engine.
    ''' 
    ''' This is intentionally shaped like OpenAiRealtimeEngine:
    ''' - StartLiveAsync opens one live WebSocket recognition session.
    ''' - PushAudioAsync accepts PCM16 mono 16 kHz frames from AudioCaptureService.
    ''' - PartialResult is raised for speech.hypothesis.
    ''' - FinalResult is raised for speech.phrase.
    ''' - TranscribeFileAsync is not implemented; use Azure fast/batch REST for file mode.
    ''' 
    ''' Required project integration:
    ''' 1. Add EngineKind.AzureSpeechRealtime to your EngineKind enum.
    ''' 2. Add this engine to LoadEngines().
    ''' 3. In CreateEngineAsync(), instantiate New AzureSpeechRealtimeEngine(subscriptionKey, regionOrEndpoint).
    ''' 
    ''' The regionOrEndpoint argument may be:
    ''' - "westeurope"
    ''' - "https://westeurope.api.cognitive.microsoft.com"
    ''' - "https://your-resource-name.cognitiveservices.azure.com"
    ''' - "wss://westeurope.stt.speech.microsoft.com"
    ''' </summary>
    Public Class AzureSpeechRealtimeEngine
        Implements ITranscriptionEngine

        Public Const DisplayNameValue As String = "Azure Speech-to-Text (streaming)"

        Private Const InputSampleRate As Integer = 16000
        Private Const BytesPerSample As Integer = 2
        Private Const ChannelCount As Integer = 1

        Private Const DefaultLanguage As String = "de-DE"

        Private Const SpeechRmsThreshold As Double = 350.0R
        Private Const MaximumUtteranceMilliseconds As Integer = 15000
        Private Const MinimumSpeechMillisecondsBeforeSend As Integer = 120
        Private Const EndSilenceBeforeStatusMilliseconds As Integer = 900

        Private ReadOnly _subscriptionKey As String
        Private ReadOnly _regionOrEndpoint As String
        Private ReadOnly _region As String

        Private _ws As System.Net.WebSockets.ClientWebSocket
        Private _readerTask As System.Threading.Tasks.Task
        Private _cts As System.Threading.CancellationTokenSource
        Private ReadOnly _sendLock As New System.Threading.SemaphoreSlim(1, 1)

        Private _connectionId As String = ""
        Private _requestId As String = ""
        Private _stopStarted As Integer = 0

        Private _bytesSinceSpeechStart As Integer = 0
        Private _speechBytesSinceStart As Integer = 0
        Private _silenceBytesAfterSpeech As Integer = 0
        Private _utteranceStartedUtc As System.DateTime = System.DateTime.MinValue

        Private _turnFinishedTcs As System.Threading.Tasks.TaskCompletionSource(Of Boolean)
        Private ReadOnly _completedPhraseKeys As New System.Collections.Generic.HashSet(Of String)(System.StringComparer.Ordinal)

        Private Shared ReadOnly _supportedLanguages As String() = {
            "auto",
            "ar-SA", "bg-BG", "ca-ES", "cs-CZ", "da-DK", "de-DE", "el-GR", "en-US", "en-GB",
            "es-ES", "es-MX", "et-EE", "fi-FI", "fr-FR", "he-IL", "hi-IN", "hr-HR", "hu-HU",
            "id-ID", "it-IT", "ja-JP", "ko-KR", "lt-LT", "lv-LV", "ms-MY", "nl-NL", "nb-NO",
            "pl-PL", "pt-PT", "pt-BR", "ro-RO", "ru-RU", "sk-SK", "sl-SI", "sv-SE", "th-TH",
            "tr-TR", "uk-UA", "vi-VN", "zh-CN", "zh-HK", "zh-TW"
        }

        Private Const DefaultRegionOrEndpoint As String = "westeurope"

        Private Function GetEffectiveRegionOrEndpoint() As String
            If Not System.String.IsNullOrWhiteSpace(_regionOrEndpoint) Then
                Return _regionOrEndpoint
            End If

            Return DefaultRegionOrEndpoint
        End Function

        Private Function GetEffectiveRealtimeRegion() As String
            If Not System.String.IsNullOrWhiteSpace(_region) Then
                Return _region
            End If

            Dim raw As String = GetEffectiveRegionOrEndpoint()

            If raw.StartsWith("https://", System.StringComparison.OrdinalIgnoreCase) OrElse
               raw.StartsWith("http://", System.StringComparison.OrdinalIgnoreCase) OrElse
               raw.StartsWith("wss://", System.StringComparison.OrdinalIgnoreCase) Then

                Dim u As New System.Uri(raw)
                Dim host As String = u.Host

                If host.EndsWith(".api.cognitive.microsoft.com", System.StringComparison.OrdinalIgnoreCase) Then
                    Return host.Substring(0, host.Length - ".api.cognitive.microsoft.com".Length)
                End If

                Throw New System.InvalidOperationException("Azure realtime streaming requires STT_Azure default.region when a custom endpoint is used.")
            End If

            If raw.IndexOf("."c) < 0 Then
                Return raw
            End If

            Throw New System.InvalidOperationException("Azure realtime streaming requires a region such as 'westeurope'.")
        End Function

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
                Return "Azure Speech"
            End Get
        End Property

        Public ReadOnly Property Kind As EngineKind Implements ITranscriptionEngine.Kind
            Get
                Try
                    Return CType(System.Enum.Parse(GetType(EngineKind), "AzureSpeechRealtime"), EngineKind)
                Catch
                    ' Keeps this single file compilable before the enum is extended.
                    ' Add EngineKind.AzureSpeechRealtime and this fallback will no longer be used.
                    Return EngineKind.OpenAiRealtime
                End Try
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

        Public Sub New(subscriptionKey As String, regionOrEndpoint As String, Optional region As String = "")
            _subscriptionKey = If(subscriptionKey, "").Trim()
            _regionOrEndpoint = If(regionOrEndpoint, "").Trim()
            _region = If(region, "").Trim()
        End Sub


        Private Sub RaiseStatusMessage(message As String)
            RaiseEvent Status(Me, New TranscriptionStatusEventArgs(message))
        End Sub

        Private Shared Sub EnsureTls12()
            Try
                System.AppContext.SetSwitch("Switch.System.Net.DontEnableSchUseStrongCrypto", False)
            Catch
            End Try

            Try
                System.AppContext.SetSwitch("Switch.System.Net.DontEnableSystemDefaultTlsVersions", False)
            Catch
            End Try

            Try
                System.Net.ServicePointManager.Expect100Continue = False
            Catch
            End Try

            Try
                System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12
            Catch
                Try
                    System.Net.ServicePointManager.SecurityProtocol = CType(3072, System.Net.SecurityProtocolType)
                Catch
                End Try
            End Try
        End Sub

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

        Private Async Function GetAccessTokenAsync(ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of String)
            Dim raw As String = GetEffectiveRegionOrEndpoint()
            Dim region As String = GetEffectiveRealtimeRegion()
            Dim tokenUri As String

            If raw.StartsWith("https://", System.StringComparison.OrdinalIgnoreCase) OrElse
               raw.StartsWith("http://", System.StringComparison.OrdinalIgnoreCase) Then
                tokenUri = raw.TrimEnd("/"c) & "/sts/v1.0/issueToken"
            Else
                tokenUri = "https://" & region & ".api.cognitive.microsoft.com/sts/v1.0/issueToken"
            End If

            Using http As New System.Net.Http.HttpClient()
                http.Timeout = System.TimeSpan.FromSeconds(20)

                Using req As New System.Net.Http.HttpRequestMessage(System.Net.Http.HttpMethod.Post, tokenUri)
                    req.Headers.TryAddWithoutValidation("Ocp-Apim-Subscription-Key", _subscriptionKey)

                    If Not System.String.IsNullOrWhiteSpace(_region) Then
                        req.Headers.TryAddWithoutValidation("Ocp-Apim-Subscription-Region", _region)
                    End If

                    req.Content = New System.Net.Http.StringContent("")

                    Using resp As System.Net.Http.HttpResponseMessage = Await http.SendAsync(req, ct).ConfigureAwait(False)
                        Dim body As String = Await resp.Content.ReadAsStringAsync().ConfigureAwait(False)

                        If Not resp.IsSuccessStatusCode Then
                            Throw New System.InvalidOperationException(
                                "Azure Speech token request failed: HTTP " &
                                CInt(resp.StatusCode).ToString(System.Globalization.CultureInfo.InvariantCulture) &
                                " " & resp.StatusCode.ToString() & ": " & body)
                        End If

                        Dim token As String = If(body, "").Trim()
                        If token.Length = 0 Then
                            Throw New System.InvalidOperationException("Azure Speech token request returned an empty token.")
                        End If

                        Return token
                    End Using
                End Using
            End Using
        End Function

        Private Function CreateWebSocket(accessToken As String) As System.Net.WebSockets.ClientWebSocket
            Dim ws As New System.Net.WebSockets.ClientWebSocket()
            ws.Options.SetRequestHeader("Authorization", "Bearer " & accessToken)
            ws.Options.SetRequestHeader("X-ConnectionId", _connectionId)
            ws.Options.KeepAliveInterval = System.TimeSpan.FromSeconds(20)
            Return ws
        End Function

        Private Shared Function NormalizeLanguageCode(opts As TranscriptionOptions) As String
            Dim raw As String = ""

            If opts IsNot Nothing Then
                raw = If(opts.LanguageCode, "").Trim()
            End If

            If System.String.IsNullOrWhiteSpace(raw) OrElse
               System.String.Equals(raw, "auto", System.StringComparison.OrdinalIgnoreCase) Then
                Return DefaultLanguage
            End If

            Select Case raw.Trim().ToLowerInvariant()
                Case "ar" : Return "ar-SA"
                Case "bg" : Return "bg-BG"
                Case "ca" : Return "ca-ES"
                Case "cs" : Return "cs-CZ"
                Case "da" : Return "da-DK"
                Case "de" : Return "de-DE"
                Case "el" : Return "el-GR"
                Case "en" : Return "en-US"
                Case "es" : Return "es-ES"
                Case "et" : Return "et-EE"
                Case "fi" : Return "fi-FI"
                Case "fr" : Return "fr-FR"
                Case "he" : Return "he-IL"
                Case "hi" : Return "hi-IN"
                Case "hr" : Return "hr-HR"
                Case "hu" : Return "hu-HU"
                Case "id" : Return "id-ID"
                Case "it" : Return "it-IT"
                Case "ja" : Return "ja-JP"
                Case "ko" : Return "ko-KR"
                Case "lt" : Return "lt-LT"
                Case "lv" : Return "lv-LV"
                Case "ms" : Return "ms-MY"
                Case "nl" : Return "nl-NL"
                Case "no", "nb" : Return "nb-NO"
                Case "pl" : Return "pl-PL"
                Case "pt" : Return "pt-PT"
                Case "ro" : Return "ro-RO"
                Case "ru" : Return "ru-RU"
                Case "sk" : Return "sk-SK"
                Case "sl" : Return "sl-SI"
                Case "sv" : Return "sv-SE"
                Case "th" : Return "th-TH"
                Case "tr" : Return "tr-TR"
                Case "uk" : Return "uk-UA"
                Case "vi" : Return "vi-VN"
                Case "zh" : Return "zh-CN"
            End Select

            Return raw
        End Function

        Private Shared Function GetEndpointHost(regionOrEndpoint As String) As String
            Dim raw As String = If(regionOrEndpoint, "").Trim()

            If raw.Length = 0 Then
                Throw New System.InvalidOperationException("Azure Speech region or endpoint is missing.")
            End If

            If raw.StartsWith("wss://", System.StringComparison.OrdinalIgnoreCase) OrElse
               raw.StartsWith("https://", System.StringComparison.OrdinalIgnoreCase) OrElse
               raw.StartsWith("http://", System.StringComparison.OrdinalIgnoreCase) Then

                Dim u As New System.Uri(raw)
                Dim host As String = u.Host

                If host.EndsWith(".api.cognitive.microsoft.com", System.StringComparison.OrdinalIgnoreCase) Then
                    Dim region As String = host.Substring(0, host.Length - ".api.cognitive.microsoft.com".Length)
                    Return region & ".stt.speech.microsoft.com"
                End If

                If host.EndsWith(".cognitiveservices.azure.com", System.StringComparison.OrdinalIgnoreCase) Then
                    Return host
                End If

                If host.EndsWith(".stt.speech.microsoft.com", System.StringComparison.OrdinalIgnoreCase) Then
                    Return host
                End If

                Return host
            End If

            If raw.IndexOf("."c) < 0 Then
                Return raw & ".stt.speech.microsoft.com"
            End If

            Return raw
        End Function

        Private Function BuildRealtimeUri(languageCode As String) As System.Uri
            Dim region As String = GetEffectiveRealtimeRegion()

            Dim query As String =
                "language=" & System.Uri.EscapeDataString(languageCode) &
                "&format=detailed"

            Return New System.Uri("wss://" & region & ".stt.speech.microsoft.com/speech/recognition/conversation/cognitiveservices/v1?" & query)
        End Function

        Public Async Function StartLiveAsync(opts As TranscriptionOptions, ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task Implements ITranscriptionEngine.StartLiveAsync
            EnsureTls12()

            If System.String.IsNullOrWhiteSpace(_subscriptionKey) Then
                Throw New System.InvalidOperationException("Azure Speech subscription key is missing.")
            End If

            System.Threading.Interlocked.Exchange(_stopStarted, 0)
            _connectionId = NewGuidN()
            _requestId = NewGuidN()
            _turnFinishedTcs = New System.Threading.Tasks.TaskCompletionSource(Of Boolean)()

            SyncLock _completedPhraseKeys
                _completedPhraseKeys.Clear()
            End SyncLock

            ResetAudioState()

            Dim languageCode As String = NormalizeLanguageCode(opts)
            Dim accessToken As String = Await GetAccessTokenAsync(ct).ConfigureAwait(False)

            _cts = System.Threading.CancellationTokenSource.CreateLinkedTokenSource(ct)
            _ws = CreateWebSocket(accessToken)

            Try
                RaiseStatusMessage("Connecting to Azure Speech streaming transcription…")
                Await _ws.ConnectAsync(BuildRealtimeUri(languageCode), _cts.Token).ConfigureAwait(False)
            Catch ex As System.Exception
                Dim detail As String = "Azure Speech connection failed: " & GetDetailedExceptionMessage(ex)

                Try
                    _ws.Dispose()
                Catch
                End Try

                _ws = Nothing
                RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs(detail, ex, True))
                Throw New System.InvalidOperationException(detail, ex)
            End Try

            _readerTask = System.Threading.Tasks.Task.Run(Function() ReadLoop(_cts.Token))

            Try
                Await SendSpeechConfigAsync(_cts.Token).ConfigureAwait(False)
            Catch ex As System.Exception
                Dim detail As String = "Azure Speech stream configuration failed: " & GetDetailedExceptionMessage(ex)
                RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs(detail, ex, True))
                Throw New System.InvalidOperationException(detail, ex)
            End Try

            RaiseStatusMessage("Azure Speech streaming session opened (" & languageCode & ").")
        End Function

        Public Async Function PushAudioAsync(pcm As Byte(), bytesValid As Integer, ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task Implements ITranscriptionEngine.PushAudioAsync
            If _ws Is Nothing OrElse _ws.State <> System.Net.WebSockets.WebSocketState.Open OrElse pcm Is Nothing OrElse bytesValid <= 0 Then
                Return
            End If

            Dim evenBytesValid As Integer = bytesValid - (bytesValid Mod 2)
            If evenBytesValid <= 0 Then
                Return
            End If

            Dim sendBytes(evenBytesValid - 1) As Byte
            System.Buffer.BlockCopy(pcm, 0, sendBytes, 0, evenBytesValid)

            Dim hasSpeech As Boolean = Pcm16HasSpeech(sendBytes, sendBytes.Length, SpeechRmsThreshold)

            If hasSpeech Then
                If _speechBytesSinceStart <= 0 Then
                    _utteranceStartedUtc = System.DateTime.UtcNow
                End If

                _speechBytesSinceStart += sendBytes.Length
                _silenceBytesAfterSpeech = 0
            Else
                If _speechBytesSinceStart > 0 Then
                    _silenceBytesAfterSpeech += sendBytes.Length
                End If
            End If

            _bytesSinceSpeechStart += sendBytes.Length

            Dim minimumSpeechBytes As Integer =
                CInt((CDbl(InputSampleRate) * CDbl(BytesPerSample)) * (CDbl(MinimumSpeechMillisecondsBeforeSend) / 1000.0R))

            If _speechBytesSinceStart <= 0 OrElse _speechBytesSinceStart >= minimumSpeechBytes Then
                Await SendAudioAsync(sendBytes, ct).ConfigureAwait(False)
            End If

            If ShouldResetLocalSpeechWindow() Then
                ResetAudioState()
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
                If wsToClose IsNot Nothing AndAlso wsToClose.State = System.Net.WebSockets.WebSocketState.Open Then
                    Try
                        Await SendAudioEndAsync(System.Threading.CancellationToken.None).ConfigureAwait(False)
                    Catch
                    End Try

                    Try
                        If _turnFinishedTcs IsNot Nothing Then
                            Await System.Threading.Tasks.Task.WhenAny(
                                _turnFinishedTcs.Task,
                                System.Threading.Tasks.Task.Delay(1500)).ConfigureAwait(False)
                        End If
                    Catch
                    End Try

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

            Try
                If ctsToCancel IsNot Nothing Then
                    ctsToCancel.Cancel()
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
            _connectionId = ""
            _requestId = ""
            _turnFinishedTcs = Nothing
            ResetAudioState()

            RaiseStatusMessage("Azure Speech stopped.")
        End Function

        Public Function TranscribeFileAsync(filePath As String, opts As TranscriptionOptions, ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task Implements ITranscriptionEngine.TranscribeFileAsync
            Throw New System.NotSupportedException("Azure Speech streaming is live-only. Use Azure fast transcription or batch transcription REST for files.")
        End Function

        Private Function ShouldResetLocalSpeechWindow() As Boolean
            If _speechBytesSinceStart <= 0 Then
                Return False
            End If

            Dim silenceBytes As Integer =
                CInt((CDbl(InputSampleRate) * CDbl(BytesPerSample)) * (CDbl(EndSilenceBeforeStatusMilliseconds) / 1000.0R))

            If _silenceBytesAfterSpeech >= silenceBytes Then
                Return True
            End If

            If _utteranceStartedUtc <> System.DateTime.MinValue AndAlso
               (System.DateTime.UtcNow - _utteranceStartedUtc).TotalMilliseconds >= MaximumUtteranceMilliseconds Then
                Return True
            End If

            Return False
        End Function

        Private Sub ResetAudioState()
            _bytesSinceSpeechStart = 0
            _speechBytesSinceStart = 0
            _silenceBytesAfterSpeech = 0
            _utteranceStartedUtc = System.DateTime.MinValue
        End Sub

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

        Private Shared Function NewGuidN() As String
            Return System.Guid.NewGuid().ToString("N", System.Globalization.CultureInfo.InvariantCulture)
        End Function

        Private Shared Function IsoUtcNow() As String
            Return System.DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ", System.Globalization.CultureInfo.InvariantCulture)
        End Function

        Private Shared Function BuildTextMessage(path As String, requestId As String, contentType As String, body As String) As String
            Dim sb As New System.Text.StringBuilder()

            sb.Append("Path: ").Append(path).Append(vbCrLf)
            sb.Append("X-RequestId: ").Append(requestId).Append(vbCrLf)
            sb.Append("X-Timestamp: ").Append(IsoUtcNow()).Append(vbCrLf)

            If Not System.String.IsNullOrWhiteSpace(contentType) Then
                sb.Append("Content-Type: ").Append(contentType).Append(vbCrLf)
            End If

            sb.Append(vbCrLf)

            If body IsNot Nothing Then
                sb.Append(body)
            End If

            Return sb.ToString()
        End Function

        Private Shared Function BuildBinaryMessage(path As String, requestId As String, contentType As String, body As Byte()) As Byte()
            Dim sb As New System.Text.StringBuilder()

            sb.Append("Path: ").Append(path).Append(vbCrLf)
            sb.Append("X-RequestId: ").Append(requestId).Append(vbCrLf)
            sb.Append("X-Timestamp: ").Append(IsoUtcNow()).Append(vbCrLf)

            If Not System.String.IsNullOrWhiteSpace(contentType) Then
                sb.Append("Content-Type: ").Append(contentType).Append(vbCrLf)
            End If

            sb.Append(vbCrLf)

            Dim headerBytes As Byte() = System.Text.Encoding.UTF8.GetBytes(sb.ToString())
            Dim payloadBytes As Byte() = If(body, System.Array.Empty(Of Byte)())

            Dim headerLength As UShort = CUShort(headerBytes.Length)
            Dim prefix As Byte() = System.BitConverter.GetBytes(headerLength)

            If System.BitConverter.IsLittleEndian Then
                System.Array.Reverse(prefix)
            End If

            Dim result(2 + headerBytes.Length + payloadBytes.Length - 1) As Byte

            System.Buffer.BlockCopy(prefix, 0, result, 0, 2)
            System.Buffer.BlockCopy(headerBytes, 0, result, 2, headerBytes.Length)

            If payloadBytes.Length > 0 Then
                System.Buffer.BlockCopy(payloadBytes, 0, result, 2 + headerBytes.Length, payloadBytes.Length)
            End If

            Return result
        End Function

        Private Async Function SendSpeechConfigAsync(ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task
            Dim speechConfig As New Newtonsoft.Json.Linq.JObject From {
                {"context", New Newtonsoft.Json.Linq.JObject From {
                    {"system", New Newtonsoft.Json.Linq.JObject From {
                        {"name", "RedInk.Transcription"},
                        {"version", "1.0.0"},
                        {"build", "drop-in-vbnet"},
                        {"lang", "VB.NET"}
                    }},
                    {"os", New Newtonsoft.Json.Linq.JObject From {
                        {"platform", "Windows"},
                        {"name", "Windows"},
                        {"version", System.Environment.OSVersion.VersionString}
                    }}
                }}
            }

            Dim message As String = BuildTextMessage(
                "speech.config",
                _requestId,
                "application/json; charset=utf-8",
                speechConfig.ToString(Newtonsoft.Json.Formatting.None))

            Await SendTextAsync(message, ct).ConfigureAwait(False)
            RaiseStatusMessage("Azure Speech stream configured.")
        End Function

        Private Async Function SendAudioAsync(audio As Byte(), ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task
            If audio Is Nothing OrElse audio.Length = 0 Then
                Return
            End If

            Dim message As Byte() = BuildBinaryMessage(
                "audio",
                _requestId,
                "audio/x-wav;codec=audio/pcm;samplerate=16000",
                audio)

            Await SendBinaryAsync(message, ct).ConfigureAwait(False)
        End Function

        Private Async Function SendAudioEndAsync(ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task
            Dim empty As Byte() = New Byte() {}
            Dim message As Byte() = BuildBinaryMessage(
                "audio",
                _requestId,
                "audio/x-wav;codec=audio/pcm;samplerate=16000",
                empty)

            Await SendBinaryAsync(message, ct).ConfigureAwait(False)
        End Function

        Private Function GetSendCancellationToken(fallback As System.Threading.CancellationToken) As System.Threading.CancellationToken
            If _cts IsNot Nothing Then
                Return _cts.Token
            End If

            Return fallback
        End Function

        Private Async Function SendTextAsync(text As String, ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task
            If _ws Is Nothing OrElse _ws.State <> System.Net.WebSockets.WebSocketState.Open Then
                Return
            End If

            Dim sendCt As System.Threading.CancellationToken = GetSendCancellationToken(ct)

            Await _sendLock.WaitAsync(sendCt).ConfigureAwait(False)

            Try
                Dim bytes As Byte() = System.Text.Encoding.UTF8.GetBytes(text)
                Await _ws.SendAsync(
                    New System.ArraySegment(Of Byte)(bytes),
                    System.Net.WebSockets.WebSocketMessageType.Text,
                    True,
                    sendCt).ConfigureAwait(False)
            Catch ex As System.Exception
                Dim detail As String = "Azure Speech WS text send failed: " & GetDetailedExceptionMessage(ex)
                RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs(detail, ex, False))
                Throw New System.InvalidOperationException(detail, ex)
            Finally
                _sendLock.Release()
            End Try
        End Function

        Private Async Function SendBinaryAsync(bytes As Byte(), ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task
            If _ws Is Nothing OrElse _ws.State <> System.Net.WebSockets.WebSocketState.Open Then
                Return
            End If

            If bytes Is Nothing Then
                Return
            End If

            Dim sendCt As System.Threading.CancellationToken = GetSendCancellationToken(ct)

            Await _sendLock.WaitAsync(sendCt).ConfigureAwait(False)

            Try
                Await _ws.SendAsync(
                    New System.ArraySegment(Of Byte)(bytes),
                    System.Net.WebSockets.WebSocketMessageType.Binary,
                    True,
                    sendCt).ConfigureAwait(False)
            Catch ex As System.Exception
                Dim detail As String = "Azure Speech WS audio send failed: " & GetDetailedExceptionMessage(ex)
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
                    Dim ms As New System.IO.MemoryStream()
                    Dim r As System.Net.WebSockets.WebSocketReceiveResult

                    Do
                        r = Await _ws.ReceiveAsync(New System.ArraySegment(Of Byte)(buf), ct).ConfigureAwait(False)

                        If r.MessageType = System.Net.WebSockets.WebSocketMessageType.Close Then
                            Return
                        End If

                        ms.Write(buf, 0, r.Count)
                    Loop While Not r.EndOfMessage

                    If r.MessageType = System.Net.WebSockets.WebSocketMessageType.Text Then
                        Dim msg As String = System.Text.Encoding.UTF8.GetString(ms.ToArray())
                        HandleServerTextMessage(msg)
                    ElseIf r.MessageType = System.Net.WebSockets.WebSocketMessageType.Binary Then
                        Dim msg As String = System.Text.Encoding.UTF8.GetString(ms.ToArray())
                        HandleServerTextMessage(msg)
                    End If
                End While
            Catch ex As System.OperationCanceledException
            Catch ex As System.Exception
                RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("Azure Speech WS read failed: " & GetDetailedExceptionMessage(ex), ex, False))
            End Try
        End Function

        Private Sub HandleServerTextMessage(message As String)
            If System.String.IsNullOrWhiteSpace(message) Then
                Return
            End If

            Try
                Dim headerEnd As Integer = message.IndexOf(vbCrLf & vbCrLf, System.StringComparison.Ordinal)
                Dim headerText As String = ""
                Dim bodyText As String = ""

                If headerEnd >= 0 Then
                    headerText = message.Substring(0, headerEnd)
                    bodyText = message.Substring(headerEnd + 4)
                Else
                    headerText = message
                    bodyText = ""
                End If

                Dim path As String = GetHeaderValue(headerText, "Path")

                Select Case path
                    Case "turn.start"
                        RaiseStatusMessage("Azure Speech turn started.")

                    Case "speech.startDetected"
                        RaiseStatusMessage("Azure Speech detected speech.")

                    Case "speech.endDetected"
                        RaiseStatusMessage("Azure Speech detected speech end.")

                    Case "speech.hypothesis"
                        HandleHypothesis(bodyText)

                    Case "speech.phrase"
                        HandlePhrase(bodyText)

                    Case "turn.end"
                        If _turnFinishedTcs IsNot Nothing Then
                            Try
                                _turnFinishedTcs.TrySetResult(True)
                            Catch
                            End Try
                        End If
                        RaiseStatusMessage("Azure Speech turn ended.")

                    Case Else
                        System.Diagnostics.Debug.WriteLine("[Azure Speech] Unhandled WS path: " & path & " " & TruncateForLog(bodyText, 1000))
                End Select
            Catch ex As Newtonsoft.Json.JsonException
                RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("Azure Speech invalid JSON message: " & TruncateForLog(message, 2000), ex, False))
            Catch ex As System.Exception
                RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("Azure Speech message handling failed: " & GetDetailedExceptionMessage(ex), ex, False))
            End Try
        End Sub

        Private Shared Function GetHeaderValue(headerText As String, headerName As String) As String
            If headerText Is Nothing OrElse headerName Is Nothing Then
                Return ""
            End If

            Dim lines As String() = headerText.Replace(vbCrLf, vbLf).Split(Microsoft.VisualBasic.ControlChars.Lf)
            Dim prefix As String = headerName & ":"

            For Each rawLine As String In lines
                Dim line As String = If(rawLine, "").Trim()

                If line.StartsWith(prefix, System.StringComparison.OrdinalIgnoreCase) Then
                    Return line.Substring(prefix.Length).Trim()
                End If
            Next

            Return ""
        End Function

        Private Sub HandleHypothesis(bodyText As String)
            If System.String.IsNullOrWhiteSpace(bodyText) Then
                Return
            End If

            Dim jo As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(bodyText)
            Dim text As String = If(jo("Text")?.ToString(), "")

            If text.Length > 0 Then
                RaiseEvent PartialResult(Me, New TranscriptionEventArgs(text, False))
            End If
        End Sub

        Private Sub HandlePhrase(bodyText As String)
            If System.String.IsNullOrWhiteSpace(bodyText) Then
                Return
            End If

            Dim jo As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(bodyText)

            Dim status As String = If(jo("RecognitionStatus")?.ToString(), "")
            If status.Length > 0 AndAlso
               Not System.String.Equals(status, "Success", System.StringComparison.OrdinalIgnoreCase) Then

                If IsIgnorableAzureStatus(status) Then
                    RaiseStatusMessage("Azure Speech: no usable speech detected.")
                    Return
                End If

                RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("Azure Speech recognition status: " & status, Nothing, False))
                Return
            End If

            Dim transcript As String = If(jo("DisplayText")?.ToString(), "")

            If System.String.IsNullOrWhiteSpace(transcript) Then
                Dim nbest As Newtonsoft.Json.Linq.JArray = TryCast(jo("NBest"), Newtonsoft.Json.Linq.JArray)

                If nbest IsNot Nothing AndAlso nbest.Count > 0 Then
                    Dim first As Newtonsoft.Json.Linq.JObject = TryCast(nbest(0), Newtonsoft.Json.Linq.JObject)

                    If first IsNot Nothing Then
                        transcript = If(first("Display")?.ToString(), "")
                        If System.String.IsNullOrWhiteSpace(transcript) Then
                            transcript = If(first("Lexical")?.ToString(), "")
                        End If
                    End If
                End If
            End If

            transcript = If(transcript, "").Trim()

            If transcript.Length = 0 Then
                RaiseStatusMessage("Azure Speech: recognition completed without text.")
                Return
            End If

            Dim offset As String = If(jo("Offset")?.ToString(), "")
            Dim duration As String = If(jo("Duration")?.ToString(), "")
            Dim key As String = offset & "|" & duration & "|" & transcript

            Dim shouldRaise As Boolean = False

            SyncLock _completedPhraseKeys
                If Not _completedPhraseKeys.Contains(key) Then
                    _completedPhraseKeys.Add(key)
                    shouldRaise = True
                End If
            End SyncLock

            If shouldRaise Then
                RaiseEvent FinalResult(Me, New TranscriptionEventArgs(transcript, True))
            Else
                System.Diagnostics.Debug.WriteLine("[Azure Speech] Duplicate final phrase ignored: " & transcript)
            End If
        End Sub

        Private Shared Function IsIgnorableAzureStatus(status As String) As Boolean
            Dim normalized As String = If(status, "").Trim().ToLowerInvariant()

            Return normalized.Length = 0 OrElse
                   normalized.Contains("no match") OrElse
                   normalized.Contains("nomatch") OrElse
                   normalized.Contains("initialsilencetimeout") OrElse
                   normalized.Contains("babbletimeout") OrElse
                   normalized.Contains("endofspeech")
        End Function

        Private Shared Function TruncateForLog(value As String, maxLength As Integer) As String
            Dim text As String = If(value, "")

            If maxLength <= 0 OrElse text.Length <= maxLength Then
                Return text
            End If

            Return text.Substring(0, maxLength) & "…"
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
