Option Explicit On
Option Strict Off

Imports System.Collections.Concurrent
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports Google.Api.Gax.Grpc
Imports Google.Protobuf
Imports Grpc.Core
Imports Newtonsoft.Json.Linq
Imports SpeechV2 = Google.Cloud.Speech.V2

Namespace Transcription

    Public Class GoogleV2Engine
        Implements ITranscriptionEngine

        ' =========================================================================
        ' Google STT V2 / Chirp 2 constants        
        ' =========================================================================
        Public Const DisplayName As String = "Google Chirp 2 (V2)"
        Private Const GoogleV2Endpoint As String = "europe-west4-speech.googleapis.com:443"
        Private Const GoogleV2Location As String = "europe-west4"
        Private Const GoogleV2ProjectNumber As String = "1092163392692"
        Private Const GoogleV2RecognizerId As String = "_"   ' Replace with a real recognizer id if "_" still fails
        Private Const GoogleV2DefaultModel As String = "chirp_2"
        Private Const GoogleV2DefaultLanguage As String = "de-DE"

        Public Shared ReadOnly Property SupportedLanguages As String()
            Get
                Return GoogleV1Engine.SupportedLanguages
            End Get
        End Property

        Public Event PartialResult As EventHandler(Of TranscriptionEventArgs) Implements ITranscriptionEngine.PartialResult
        Public Event FinalResult As EventHandler(Of TranscriptionEventArgs) Implements ITranscriptionEngine.FinalResult
        Public Event EngineError As EventHandler(Of TranscriptionErrorEventArgs) Implements ITranscriptionEngine.EngineError
        Public Event Status As EventHandler(Of TranscriptionStatusEventArgs) Implements ITranscriptionEngine.Status

        Public ReadOnly Property Name As String Implements ITranscriptionEngine.Name
            Get
                Return DisplayName
            End Get
        End Property

        Public ReadOnly Property Kind As EngineKind Implements ITranscriptionEngine.Kind
            Get
                Return EngineKind.GoogleV2
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
                Return False
            End Get
        End Property

        Public ReadOnly Property SupportsMultiChannelDiarization As Boolean Implements ITranscriptionEngine.SupportsMultiChannelDiarization
            Get
                Return True
            End Get
        End Property

        Private ReadOnly _serviceAccountEmail As String
        Private ReadOnly _serviceAccountPrivateKeyRaw As String
        Private ReadOnly _serviceAccountTokenUri As String

        Private _client As SpeechV2.SpeechClient
        Private _stream As SpeechV2.SpeechClient.StreamingRecognizeStream
        Private _audioQueue As BlockingCollection(Of ByteString)
        Private _writerTask As Task
        Private _readerTask As Task
        Private _readerCts As CancellationTokenSource
        Private _opts As TranscriptionOptions
        Private _firstResponseSeen As Boolean

        Public Sub New(serviceAccountEmail As String, serviceAccountPrivateKeyRaw As String, serviceAccountTokenUri As String)
            _serviceAccountEmail = serviceAccountEmail
            _serviceAccountPrivateKeyRaw = serviceAccountPrivateKeyRaw
            _serviceAccountTokenUri = serviceAccountTokenUri
        End Sub

        Private Sub RaiseStatusMessage(message As String)
            RaiseEvent Status(Me, New TranscriptionStatusEventArgs(message))
        End Sub

        Private Function GetRecognizerName() As String
            Return "projects/" & GoogleV2ProjectNumber & "/locations/" & GoogleV2Location & "/recognizers/" & GoogleV2RecognizerId
        End Function

        Private Shared Function FormatPrivateKey(rawKey As String) As String
            Dim noEscapes As String = rawKey.Replace("\n", "")
            Dim sb As New StringBuilder()

            For i As Integer = 0 To noEscapes.Length - 1 Step 64
                Dim chunk As String
                If i + 64 <= noEscapes.Length Then
                    chunk = noEscapes.Substring(i, 64)
                Else
                    chunk = noEscapes.Substring(i)
                End If
                sb.AppendLine(chunk)
            Next

            Return "-----BEGIN PRIVATE KEY-----" & vbLf &
                   sb.ToString() &
                   "-----END PRIVATE KEY-----" & vbLf
        End Function

        Private Shared Function ExtractProjectIdFromServiceAccountEmail(email As String) As String
            If String.IsNullOrWhiteSpace(email) Then
                Return ""
            End If

            Dim atPos As Integer = email.IndexOf("@"c)
            If atPos < 0 Then
                Return ""
            End If

            Dim domain As String = email.Substring(atPos + 1)
            Dim suffix As String = ".iam.gserviceaccount.com"

            If domain.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) Then
                Return domain.Substring(0, domain.Length - suffix.Length)
            End If

            Dim firstDot As Integer = domain.IndexOf("."c)
            If firstDot > 0 Then
                Return domain.Substring(0, firstDot)
            End If

            Return domain
        End Function

        Private Function BuildServiceAccountJson() As String
            If String.IsNullOrWhiteSpace(_serviceAccountEmail) Then
                Throw New InvalidOperationException("Google V2 service-account email is empty.")
            End If

            If String.IsNullOrWhiteSpace(_serviceAccountPrivateKeyRaw) Then
                Throw New InvalidOperationException("Google V2 service-account private key is empty.")
            End If

            Dim tokenUri As String = If(String.IsNullOrWhiteSpace(_serviceAccountTokenUri),
                                        "https://oauth2.googleapis.com/token",
                                        _serviceAccountTokenUri)

            Dim escapedMail As String = Uri.EscapeDataString(_serviceAccountEmail)
            Dim projectId As String = ExtractProjectIdFromServiceAccountEmail(_serviceAccountEmail)

            Dim json As New JObject From {
                {"type", "service_account"},
                {"project_id", projectId},
                {"private_key_id", ""},
                {"private_key", FormatPrivateKey(_serviceAccountPrivateKeyRaw)},
                {"client_email", _serviceAccountEmail},
                {"client_id", ""},
                {"auth_uri", "https://accounts.google.com/o/oauth2/auth"},
                {"token_uri", tokenUri},
                {"auth_provider_x509_cert_url", "https://www.googleapis.com/oauth2/v1/certs"},
                {"client_x509_cert_url", "https://www.googleapis.com/robot/v1/metadata/x509/" & escapedMail}
            }

            Return json.ToString()
        End Function

        Private Async Function BuildClient() As Task
            Dim builder As New SpeechV2.SpeechClientBuilder() With {
                .Endpoint = GoogleV2Endpoint,
                .JsonCredentials = BuildServiceAccountJson(),
                .GrpcAdapter = GrpcCoreAdapter.Instance
            }

            _client = builder.Build()

            RaiseStatusMessage("Google V2 client built with GrpcCoreAdapter.")
            RaiseStatusMessage("Google V2 service account: " & _serviceAccountEmail)

            Await Task.CompletedTask
        End Function

        Private Function BuildRecognitionConfig(opts As TranscriptionOptions) As SpeechV2.RecognitionConfig
            Dim effectiveModel As String = If(String.IsNullOrWhiteSpace(opts.Model), GoogleV2DefaultModel, opts.Model)
            Dim effectiveLanguage As String = If(String.IsNullOrWhiteSpace(opts.LanguageCode), GoogleV2DefaultLanguage, opts.LanguageCode)

            Dim cfg As New SpeechV2.RecognitionConfig With {
                .Model = effectiveModel,
                .ExplicitDecodingConfig = New SpeechV2.ExplicitDecodingConfig With {
                    .Encoding = SpeechV2.ExplicitDecodingConfig.Types.AudioEncoding.Linear16,
                    .SampleRateHertz = 16000,
                    .AudioChannelCount = If(opts.MultiChannelDiarization, 2, 1)
                },
                .Features = New SpeechV2.RecognitionFeatures With {
                    .EnableAutomaticPunctuation = True,
                    .EnableWordTimeOffsets = False,
                    .MultiChannelMode = If(
                        opts.MultiChannelDiarization,
                        SpeechV2.RecognitionFeatures.Types.MultiChannelMode.SeparateRecognitionPerChannel,
                        SpeechV2.RecognitionFeatures.Types.MultiChannelMode.Unspecified)
                }
            }

            cfg.LanguageCodes.Add(effectiveLanguage)

            Return cfg
        End Function

        Public Async Function StartLiveAsync(opts As TranscriptionOptions, ct As CancellationToken) As Task Implements ITranscriptionEngine.StartLiveAsync
            _opts = opts
            _audioQueue = New BlockingCollection(Of ByteString)()
            _readerCts = New CancellationTokenSource()
            _firstResponseSeen = False

            RaiseStatusMessage("Starting Google V2…")
            RaiseStatusMessage("Project=" & GoogleV2ProjectNumber & ", Location=" & GoogleV2Location & ", Endpoint=" & GoogleV2Endpoint)

            Try
                Await BuildClient()

                _stream = _client.StreamingRecognize()

                Dim req As New SpeechV2.StreamingRecognizeRequest With {
                    .Recognizer = GetRecognizerName(),
                    .StreamingConfig = New SpeechV2.StreamingRecognitionConfig With {
                        .Config = BuildRecognitionConfig(opts),
                        .StreamingFeatures = New SpeechV2.StreamingRecognitionFeatures With {
                            .InterimResults = True
                        }
                    }
                }

                RaiseStatusMessage("Opening Google V2 streaming session…")
                RaiseStatusMessage("Recognizer=" & req.Recognizer)
                RaiseStatusMessage("Language=" & If(String.IsNullOrWhiteSpace(opts.LanguageCode), GoogleV2DefaultLanguage & "(default)", opts.LanguageCode))
                RaiseStatusMessage("Model=" & If(String.IsNullOrWhiteSpace(opts.Model), GoogleV2DefaultModel & "(default)", opts.Model))

                Await _stream.WriteAsync(req)
                RaiseStatusMessage("Google V2 stream configured.")

                StartWriter()
                _readerTask = StartReader()
            Catch ex As Exception
                RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("Google V2 start failed: " & ex.Message, ex, True))
                Throw
            End Try
        End Function

        Public Function PushAudioAsync(pcm As Byte(), bytesValid As Integer, ct As CancellationToken) As Task Implements ITranscriptionEngine.PushAudioAsync
            If _audioQueue Is Nothing OrElse _audioQueue.IsAddingCompleted Then
                Return Task.CompletedTask
            End If

            _audioQueue.Add(ByteString.CopyFrom(pcm, 0, bytesValid))
            Return Task.CompletedTask
        End Function

        Public Async Function StopLiveAsync() As Task Implements ITranscriptionEngine.StopLiveAsync
            Try
                If _readerCts IsNot Nothing Then
                    _readerCts.Cancel()
                End If
            Catch
            End Try

            Try
                If _stream IsNot Nothing Then
                    Await _stream.WriteCompleteAsync()
                End If
            Catch ex As Exception
                RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("Google V2 WriteComplete failed: " & ex.Message, ex, False))
            End Try

            Try
                If _audioQueue IsNot Nothing AndAlso Not _audioQueue.IsAddingCompleted Then
                    _audioQueue.CompleteAdding()
                End If

                If _writerTask IsNot Nothing Then
                    Await _writerTask
                End If
            Catch ex As Exception
                RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("Google V2 writer shutdown failed: " & ex.Message, ex, False))
            End Try

            If _readerTask IsNot Nothing Then
                Try
                    Await _readerTask
                Catch ex As Exception
                    RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("Google V2 reader shutdown failed: " & ex.Message, ex, False))
                End Try
            End If

            Try
                If _stream IsNot Nothing Then
                    _stream.Dispose()
                End If
            Catch
            End Try

            _stream = Nothing
            RaiseStatusMessage("Google V2 stopped.")
        End Function

        Public Async Function TranscribeFileAsync(filePath As String, opts As TranscriptionOptions, ct As CancellationToken) As Task Implements ITranscriptionEngine.TranscribeFileAsync
            RaiseStatusMessage("Streaming file to Google V2…")
            Await StartLiveAsync(opts, ct)

            Dim pcm As Byte() = VoskEngine.LoadAudioToPcm16Mono16k(filePath)
            Const chunkSize As Integer = 4096
            Dim bytesPerSec As Integer = 16000 * 2
            Dim pos As Integer = 0

            While pos < pcm.Length AndAlso Not ct.IsCancellationRequested
                Dim n As Integer = Math.Min(chunkSize, pcm.Length - pos)
                Dim slice(n - 1) As Byte
                Buffer.BlockCopy(pcm, pos, slice, 0, n)
                Await PushAudioAsync(slice, n, ct)
                Await Task.Delay(CInt(1000.0 * n / bytesPerSec), ct)
                pos += n
            End While

            Await StopLiveAsync()
        End Function

        Private Sub StartWriter()
            _writerTask =
                Task.Run(
                    Async Function()
                        Try
                            For Each chunk As ByteString In _audioQueue.GetConsumingEnumerable()
                                Try
                                    If _stream Is Nothing Then
                                        RaiseStatusMessage("Google V2 writer stopped because stream is nothing.")
                                        Exit For
                                    End If

                                    Await _stream.WriteAsync(New SpeechV2.StreamingRecognizeRequest With {
                                        .Audio = chunk
                                    })
                                Catch ex As Exception
                                    RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("Google V2 audio write failed: " & ex.Message, ex, False))
                                    Exit For
                                End Try
                            Next
                        Catch ex As Exception
                            RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("Google V2 writer failed: " & ex.Message, ex, False))
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
                            If Not _firstResponseSeen Then
                                _firstResponseSeen = True
                                RaiseStatusMessage("Google V2 returned its first response.")
                            End If

                            For Each r In it.Current.Results
                                If r.Alternatives.Count = 0 Then
                                    Continue For
                                End If

                                Dim txt As String = r.Alternatives(0).Transcript.Trim()
                                Dim spk As String = String.Empty

                                If r.ChannelTag > 0 Then
                                    spk = "Speaker " & r.ChannelTag.ToString()
                                End If

                                If r.IsFinal Then
                                    RaiseEvent FinalResult(Me, New TranscriptionEventArgs(txt, True, spk))
                                Else
                                    RaiseEvent PartialResult(Me, New TranscriptionEventArgs(txt, False, spk))
                                End If
                            Next
                        End While
                    Catch ocex As OperationCanceledException
                    Catch rex As RpcException
                        RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("Google V2 RPC error (" & rex.StatusCode.ToString() & "): " & rex.Status.Detail, rex, False))
                    Catch ex As Exception
                        RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("Google V2 reader failed: " & ex.Message, ex, False))
                    End Try
                End Function)
        End Function

        Public Function DisposeAsync() As ValueTask Implements IAsyncDisposable.DisposeAsync
            Return New ValueTask(DisposeAsyncCore())
        End Function

        Private Async Function DisposeAsyncCore() As Task
            Await StopLiveAsync()
        End Function
    End Class

End Namespace