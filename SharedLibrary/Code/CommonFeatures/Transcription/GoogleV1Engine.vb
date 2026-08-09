' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: GoogleV1Engine.vb
' Purpose: Implements the ITranscriptionEngine interface using the Google
'          Cloud Speech-to-Text API v1. This engine is responsible for sending
'          audio data to Google's service and processing the transcription results.
'
' Architecture:
'  - ITranscriptionEngine Implementation: Fulfills the contract for starting
'    and stopping transcription.
'  - API Integration: Interacts with the Google Cloud Speech-to-Text v1 API,
'    handling authentication, request formation, and response parsing.
'  - Audio Handling: Prepares audio data in a format compatible with the API,
'    potentially handling different encodings and sample rates.
'  - Asynchronous Operations: Manages asynchronous calls to the Google API to
'    avoid blocking the UI thread during transcription.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Collections.Concurrent
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports Google.Cloud.Speech.V1
Imports Google.Protobuf
Imports Grpc.Core

Namespace Transcription

    Public Class GoogleV1Engine
        Implements ITranscriptionEngine

        Public Const DisplayName As String = "Google STT (V1)"

        Private Const DefaultGoogleV1Endpoint As String = "eu-speech.googleapis.com:443"

        Private Shared ReadOnly _supportedLanguages As String() = {
            "af-ZA", "am-ET", "ar-BH", "ar-DZ", "ar-EG", "ar-IQ", "ar-IL", "ar-JO", "ar-KW", "ar-LB", "ar-MA",
            "ar-MR", "ar-OM", "ar-PS", "ar-QA", "ar-SA", "ar-SY", "ar-TN", "ar-AE", "ar-YE", "az-AZ", "bg-BG",
            "bn-BD", "bn-IN", "bs-BA", "ca-ES", "cmn-Hans-CN", "cmn-Hans-HK", "cmn-Hant-TW", "cs-CZ", "da-DK",
            "de-AT", "de-CH", "de-DE", "el-GR", "en-AU", "en-CA", "en-GH", "en-HK", "en-IE", "en-IN", "en-KE",
            "en-NG", "en-NZ", "en-PH", "en-PK", "en-SG", "en-TZ", "en-US", "en-ZA", "es-AR", "es-BO", "es-CL",
            "es-CO", "es-CR", "es-DO", "es-EC", "es-ES", "es-GT", "es-HN", "es-MX", "es-NI", "es-PA", "es-PE",
            "es-PR", "es-PY", "es-SV", "es-UY", "es-VE", "et-EE", "eu-ES", "fa-IR", "fi-FI", "fil-PH", "fr-BE",
            "fr-CA", "fr-CH", "fr-FR", "gl-ES", "gu-IN", "hi-IN", "hr-HR", "hu-HU", "hy-AM", "id-ID", "is-IS",
            "it-CH", "it-IT", "iw-IL", "ja-JP", "jv-ID", "ka-GE", "kk-KZ", "km-KH", "kn-IN", "ko-KR", "lo-LA",
            "lt-LT", "lv-LV", "ml-IN", "mn-MN", "mr-IN", "ms-MY", "my-MM", "ne-NP", "nl-BE", "nl-NL", "no-NO",
            "pa-Guru-IN", "pl-PL", "pt-BR", "pt-PT", "ro-RO", "ru-RU", "rw-RW", "si-LK", "sk-SK", "sl-SI", "sr-RS",
            "ss-Latn-ZA", "st-ZA", "su-ID", "sv-SE", "sw-KE", "sw-TZ", "ta-IN", "ta-LK", "ta-MY", "ta-SG", "te-IN",
            "th-TH", "tn-Latn-ZA", "tr-TR", "uk-UA", "ur-IN", "ur-PK", "uz-UZ", "ve-ZA", "vi-VN", "xh-ZA",
            "yue-Hant-HK", "zu-ZA"
        }

        Public Shared ReadOnly Property SupportedLanguages As String()
            Get
                Return _supportedLanguages.OrderBy(Function(x) x, StringComparer.OrdinalIgnoreCase).ToArray()
            End Get
        End Property


        Public Event PartialResult As EventHandler(Of TranscriptionEventArgs) Implements ITranscriptionEngine.PartialResult
        Public Event FinalResult As EventHandler(Of TranscriptionEventArgs) Implements ITranscriptionEngine.FinalResult
        Public Event EngineError As EventHandler(Of TranscriptionErrorEventArgs) Implements ITranscriptionEngine.EngineError
        Public Event Status As EventHandler(Of TranscriptionStatusEventArgs) Implements ITranscriptionEngine.Status

        Public ReadOnly Property Name As String Implements ITranscriptionEngine.Name
            Get
                Return "Google STT (V1)"
            End Get
        End Property

        Public ReadOnly Property Kind As EngineKind Implements ITranscriptionEngine.Kind
            Get
                Return EngineKind.GoogleV1
            End Get
        End Property

        Public ReadOnly Property SupportsLiveStreaming As Boolean Implements ITranscriptionEngine.SupportsLiveStreaming
            Get
                Return True
            End Get
        End Property

        Public ReadOnly Property SupportsFileTranscription As Boolean Implements ITranscriptionEngine.SupportsFileTranscription
            Get
                Return True
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

        Private Sub RaiseStatusMessage(message As String, Optional progressPercent As System.Nullable(Of Integer) = Nothing)
            RaiseEvent Status(Me, New TranscriptionStatusEventArgs(message, progressPercent))
        End Sub

        Private ReadOnly _endpoint As String
        Private ReadOnly _tokenFactory As Func(Of Task(Of String))

        Private _opts As TranscriptionOptions
        Private _client As SpeechClient
        Private _stream As SpeechClient.StreamingRecognizeStream
        Private _audioQueue As BlockingCollection(Of ByteString)
        Private _writerTask As Task
        Private _readerTask As Task
        Private _readerCts As CancellationTokenSource
        Private _streamStarted As DateTime
        Private _completed As Boolean
        Private _lastPartial As String = String.Empty

        Private Const STREAMING_LIMIT_MS As Integer = 290000
        Private Const RING_BUFFER_SIZE As Integer = 50

        Private ReadOnly _ringBuffer As New Queue(Of ByteString)()
        Private ReadOnly _recoverySem As New SemaphoreSlim(1, 1)
        Private ReadOnly _speakerMap As New Dictionary(Of Integer, String)()
        Private _nextSpeaker As Integer = 1

        Public Sub New(endpoint As String, tokenFactory As Func(Of Task(Of String)))
            If String.IsNullOrWhiteSpace(endpoint) OrElse Not endpoint.ToLowerInvariant().Contains("speech.googleapis.com") Then
                _endpoint = DefaultGoogleV1Endpoint
            Else
                _endpoint = endpoint.Trim()
                If Not _endpoint.Contains(":") Then
                    _endpoint &= ":443"
                End If
            End If

            _tokenFactory = tokenFactory
        End Sub

        Public Async Function StartLiveAsync(opts As TranscriptionOptions, ct As CancellationToken) As Task Implements ITranscriptionEngine.StartLiveAsync
            _opts = opts
            _audioQueue = New BlockingCollection(Of ByteString)()
            _readerCts = New CancellationTokenSource()
            _speakerMap.Clear()
            _nextSpeaker = 1

            Await BuildClientAndStream()
            StartWriter()
            _readerTask = StartReader()
        End Function

        Public Function PushAudioAsync(pcm As Byte(), bytesValid As Integer, ct As CancellationToken) As Task Implements ITranscriptionEngine.PushAudioAsync
            If _completed OrElse _audioQueue Is Nothing OrElse _audioQueue.IsAddingCompleted Then
                Return Task.CompletedTask
            End If

            Dim chunk As ByteString = ByteString.CopyFrom(pcm, 0, bytesValid)

            SyncLock _ringBuffer
                _ringBuffer.Enqueue(chunk)
                If _ringBuffer.Count > RING_BUFFER_SIZE Then
                    _ringBuffer.Dequeue()
                End If
            End SyncLock

            _audioQueue.Add(chunk)

            If (DateTime.UtcNow - _streamStarted).TotalMilliseconds > STREAMING_LIMIT_MS Then
                _streamStarted = DateTime.UtcNow
                Task.Run(
                    Async Function()
                        Await TryRecoverAsync()
                    End Function)
            End If

            Return Task.CompletedTask
        End Function

        Public Async Function StopLiveAsync() As Task Implements ITranscriptionEngine.StopLiveAsync
            Try
                If _readerCts IsNot Nothing Then
                    _readerCts.Cancel()
                End If
            Catch
            End Try

            Await SafeCompleteStream()

            If _readerTask IsNot Nothing Then
                Try
                    Await _readerTask
                Catch
                End Try
            End If
        End Function

        Public Async Function TranscribeFileAsync(filePath As String, opts As TranscriptionOptions, ct As CancellationToken) As Task Implements ITranscriptionEngine.TranscribeFileAsync
            _opts = opts
            RaiseStatusMessage("Preparing file…", 0)
            Await BuildClientOnly()

            Dim pcm As Byte() = VoskEngine.LoadAudioToPcm16Mono16k(filePath)
            Dim bytesPerSec As Integer = 16000 * 2
            Dim sliceSize As Integer = 50 * bytesPerSec
            Dim overlap As Integer = 2 * bytesPerSec
            Dim offset As Integer = 0
            Dim lastProgressPercent As Integer = -1

            RaiseStatusMessage("Transcribing file…", 0)

            While offset < pcm.Length AndAlso Not ct.IsCancellationRequested
                Dim endPos As Integer = Math.Min(offset + sliceSize, pcm.Length)
                Dim slice(endPos - offset - 1) As Byte
                Buffer.BlockCopy(pcm, offset, slice, 0, endPos - offset)

                Dim cfg As RecognitionConfig = BuildConfig(opts)
                Dim audio As RecognitionAudio = RecognitionAudio.FromBytes(slice)
                Dim resp As RecognizeResponse = Await _client.RecognizeAsync(cfg, audio)

                For Each r As SpeechRecognitionResult In resp.Results
                    If r.Alternatives.Count > 0 Then
                        RaiseEvent FinalResult(Me, New TranscriptionEventArgs(r.Alternatives(0).Transcript.Trim(), True))
                    End If
                Next

                Dim progressPercent As Integer =
                    CInt(Math.Truncate((CDbl(endPos) / Math.Max(1.0R, CDbl(pcm.Length))) * 100.0R))

                If progressPercent <> lastProgressPercent Then
                    lastProgressPercent = progressPercent
                    RaiseStatusMessage("Transcribing file…", progressPercent)
                End If

                If endPos >= pcm.Length Then
                    Exit While
                End If

                offset = endPos - overlap
                If offset < 0 Then
                    offset = 0
                End If
            End While

            If ct.IsCancellationRequested Then
                RaiseStatusMessage("File transcription canceled.")
            Else
                RaiseStatusMessage("File transcription completed.", 100)
            End If
        End Function

        Private Function BuildConfig(opts As TranscriptionOptions) As RecognitionConfig
            Dim cfg As New RecognitionConfig With {
                .Encoding = RecognitionConfig.Types.AudioEncoding.Linear16,
                .SampleRateHertz = 16000,
                .LanguageCode = opts.LanguageCode,
                .EnableAutomaticPunctuation = True,
                .EnableSpokenPunctuation = True,
                .Model = "latest_long",
                .UseEnhanced = True
            }

            If opts.EnableDiarization Then
                cfg.DiarizationConfig = New SpeakerDiarizationConfig With {
                    .EnableSpeakerDiarization = True,
                    .MinSpeakerCount = Math.Max(2, opts.MinSpeakers),
                    .MaxSpeakerCount = Math.Max(opts.MinSpeakers, opts.MaxSpeakers)
                }
            End If

            Return cfg
        End Function

        Private Async Function BuildClientOnly() As Task
            Dim token As String = Await _tokenFactory()

            Dim callCreds As CallCredentials =
                CallCredentials.FromInterceptor(
                    Async Function(c, md)
                        md.Add("Authorization", "Bearer " & token)
                        Await Task.CompletedTask
                    End Function)

            Dim chCreds As ChannelCredentials = ChannelCredentials.Create(ChannelCredentials.SecureSsl, callCreds)

            _client = New SpeechClientBuilder() With {
                .Endpoint = _endpoint,
                .ChannelCredentials = chCreds
            }.Build()
        End Function

        Private Async Function BuildClientAndStream() As Task
            Await BuildClientOnly()

            _streamStarted = DateTime.UtcNow
            _completed = False
            _stream = _client.StreamingRecognize()

            Dim primaryConfig As New StreamingRecognitionConfig With {
                .Config = BuildConfig(_opts),
                .InterimResults = True,
                .SingleUtterance = False
            }

            Dim fallbackNeeded As Boolean = False

            Try
                Await _stream.WriteAsync(New StreamingRecognizeRequest With {.StreamingConfig = primaryConfig})
            Catch
                fallbackNeeded = True
            End Try

            If fallbackNeeded Then
                Dim cfg2 As New RecognitionConfig With {
                    .Encoding = RecognitionConfig.Types.AudioEncoding.Linear16,
                    .SampleRateHertz = 16000,
                    .LanguageCode = _opts.LanguageCode
                }

                Dim sc As New StreamingRecognitionConfig With {
                    .Config = cfg2,
                    .InterimResults = True
                }

                _stream = _client.StreamingRecognize()
                Await _stream.WriteAsync(New StreamingRecognizeRequest With {.StreamingConfig = sc})
            End If

        End Function

        Private Sub StartWriter()
            _writerTask =
                Task.Run(
                    Async Function()
                        Try
                            For Each chunk As ByteString In _audioQueue.GetConsumingEnumerable()
                                Dim shouldExit As Boolean = False

                                Try
                                    If _stream Is Nothing Then
                                        shouldExit = True
                                    Else
                                        Await _stream.WriteAsync(New StreamingRecognizeRequest With {.AudioContent = chunk})
                                    End If
                                Catch
                                    shouldExit = True
                                End Try

                                If shouldExit Then
                                    Exit For
                                End If
                            Next
                        Catch
                        End Try
                    End Function)
        End Sub

        Private Function StartReader() As Task
            Dim token As CancellationToken = _readerCts.Token

            Return Task.Run(
                Async Function()
                    Try
                        Dim it = _stream.GetResponseStream().GetAsyncEnumerator(token)

                        While Await it.MoveNextAsync()
                            For Each result In it.Current.Results
                                If result.Alternatives.Count = 0 Then
                                    Continue For
                                End If

                                If result.IsFinal Then
                                    Dim alt = result.Alternatives(0)
                                    Dim txt As String = alt.Transcript.Trim()
                                    _lastPartial = String.Empty

                                    If _opts.EnableDiarization AndAlso alt.Words.Count > 0 Then
                                        Dim sb As New StringBuilder()
                                        Dim curSpeaker As String = LabelFor(alt.Words(0).SpeakerTag)

                                        For Each w In alt.Words
                                            Dim lab As String = LabelFor(w.SpeakerTag)

                                            If lab <> curSpeaker Then
                                                RaiseEvent FinalResult(Me, New TranscriptionEventArgs(sb.ToString().Trim(), True, curSpeaker))
                                                sb.Clear()
                                                curSpeaker = lab
                                            End If

                                            sb.Append(w.Word & " ")
                                        Next

                                        If sb.Length > 0 Then
                                            RaiseEvent FinalResult(Me, New TranscriptionEventArgs(sb.ToString().Trim(), True, curSpeaker))
                                        End If
                                    Else
                                        RaiseEvent FinalResult(Me, New TranscriptionEventArgs(txt, True))
                                    End If
                                Else
                                    _lastPartial = result.Alternatives(0).Transcript
                                    RaiseEvent PartialResult(Me, New TranscriptionEventArgs(_lastPartial, False))
                                End If
                            Next
                        End While
                    Catch ocex As OperationCanceledException
                    Catch rex As RpcException
                        If Not token.IsCancellationRequested AndAlso rex.StatusCode <> StatusCode.Cancelled Then
                            If Not String.IsNullOrWhiteSpace(_lastPartial) Then
                                RaiseEvent FinalResult(Me, New TranscriptionEventArgs(_lastPartial.Trim(), True))
                                _lastPartial = String.Empty
                            End If

                            RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("Google V1 stream interrupted.", rex, False))
                        End If
                    Catch ex As Exception
                        RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("Reader: " & ex.Message, ex, False))
                    End Try
                End Function)
        End Function

        Private Function LabelFor(tag As Integer) As String
            If _speakerMap.ContainsKey(tag) Then
                Return _speakerMap(tag)
            End If

            Dim lab As String = "Speaker " & _nextSpeaker.ToString()
            _nextSpeaker += 1
            _speakerMap(tag) = lab
            Return lab
        End Function

        Private Async Function TryRecoverAsync() As Task
            If Not Await _recoverySem.WaitAsync(0) Then
                Return
            End If

            Try
                If Not String.IsNullOrWhiteSpace(_lastPartial) Then
                    RaiseEvent FinalResult(Me, New TranscriptionEventArgs(_lastPartial.Trim(), True))
                    _lastPartial = String.Empty
                End If

                Dim oldReader As Task = _readerTask

                Try
                    If _readerCts IsNot Nothing Then
                        _readerCts.Cancel()
                    End If
                Catch
                End Try

                Await SafeCompleteStream()

                If oldReader IsNot Nothing Then
                    Try
                        Await oldReader
                    Catch
                    End Try
                End If

                _readerCts = New CancellationTokenSource()
                Await BuildClientAndStream()
                _audioQueue = New BlockingCollection(Of ByteString)()
                StartWriter()
                _readerTask = StartReader()
            Finally
                _recoverySem.Release()
            End Try
        End Function

        Private Async Function SafeCompleteStream() As Task
            _completed = True

            If _audioQueue IsNot Nothing AndAlso Not _audioQueue.IsAddingCompleted Then
                Try
                    _audioQueue.CompleteAdding()
                Catch
                End Try
            End If

            If _writerTask IsNot Nothing AndAlso Not _writerTask.IsCompleted Then
                Try
                    Await _writerTask
                Catch
                End Try
            End If

            Try
                If _stream IsNot Nothing Then
                    Await _stream.WriteCompleteAsync()
                End If
            Catch
            End Try

            Try
                If _stream IsNot Nothing Then
                    _stream.Dispose()
                End If
            Catch
            End Try

            _stream = Nothing
        End Function

        Public Function DisposeAsync() As ValueTask Implements IAsyncDisposable.DisposeAsync
            Return New ValueTask(DisposeAsyncCore())
        End Function

        Private Async Function DisposeAsyncCore() As Task
            Await StopLiveAsync()
        End Function
    End Class

End Namespace
