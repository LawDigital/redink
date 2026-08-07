' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: GoogleV2Engine.vb
' Purpose: Implements the ITranscriptionEngine interface using Google's Speech-to-Text
'          API (V2). It handles audio streaming, recognition, and result processing.
'
' Architecture:
'  - ITranscriptionEngine: Adheres to the common transcription engine contract
'    defined in the SharedLibrary.
'  - Google Cloud Client: Uses the `Google.Cloud.Speech.V2.SpeechClient` for
'    communication with the Google Cloud Speech-to-Text service.
'  - Asynchronous Streaming: Manages a bidirectional streaming call to send
'    audio data and receive transcription results asynchronously.
'  - State Management: Tracks the transcription state (e.g., active, stopping)
'    and handles events for transcription results and errors.
'  - Configuration: Configured with Google Cloud project details and recognition
'    settings (e.g., language, model).
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Collections.Concurrent
Imports System.Collections.Generic
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports Google.Api.Gax.Grpc
Imports Google.Protobuf
Imports Grpc.Core
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SpeechV2 = Google.Cloud.Speech.V2

Namespace Transcription

    Public Class GoogleV2Engine
        Implements ITranscriptionEngine

        Public Const DisplayName As String = "Google Chirp 2 (V2)"

        Private Const GoogleV2Endpoint As String = "europe-west4-speech.googleapis.com:443"
        Private Const GoogleV2Location As String = "europe-west4"
        Private Const GoogleV2RecognizerId As String = "_"
        Private Const GoogleV2DefaultModel As String = "chirp_2"
        Private Const GoogleV2DefaultLanguage As String = "de-DE"
        Private Const DefaultGoogleOAuthScope As String = "https://www.googleapis.com/auth/cloud-platform"

        Private Const StreamingLimitMs As Integer = 290000
        Private Const RingBufferSize As Integer = 50
        Private Const OAuthTokenLifetimeSeconds As Integer = 3600
        Private Const OAuthRefreshSkewSeconds As Integer = 300
        Private Const DuplicateFinalWindowSeconds As Integer = 15

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
        Private ReadOnly _projectId As String
        Private ReadOnly _endpoint As String
        Private ReadOnly _location As String
        Private ReadOnly _recognizerId As String
        Private ReadOnly _defaultModel As String
        Private ReadOnly _defaultLanguage As String
        Private ReadOnly _oauthScopes As String
        Private ReadOnly _debugSource As String

        Private ReadOnly _ringBuffer As New Queue(Of ByteString)()
        Private ReadOnly _recoverySem As New SemaphoreSlim(1, 1)
        Private ReadOnly _stateSyncRoot As New Object()

        Private Shared ReadOnly _googleOAuthSem As New SemaphoreSlim(1, 1)

        Private _client As SpeechV2.SpeechClient
        Private _stream As SpeechV2.SpeechClient.StreamingRecognizeStream
        Private _audioQueue As BlockingCollection(Of ByteString)
        Private _writerTask As Task
        Private _readerTask As Task
        Private _readerCts As CancellationTokenSource
        Private _opts As TranscriptionOptions
        Private _firstResponseSeen As Boolean
        Private _streamStarted As DateTime
        Private _completed As Boolean
        Private _stopping As Boolean
        Private _lastPartial As String = String.Empty
        Private _accessToken As String = String.Empty
        Private _accessTokenExpiresUtc As DateTime = DateTime.MinValue
        Private _lastFinalResultKey As String = String.Empty
        Private _lastFinalResultUtc As DateTime = DateTime.MinValue

        Public Sub New(serviceAccountEmail As String, serviceAccountPrivateKeyRaw As String, serviceAccountTokenUri As String)
            Me.New(serviceAccountEmail, serviceAccountPrivateKeyRaw, serviceAccountTokenUri, "", "", "", "", "", "", "", "")
        End Sub

        Public Sub New(serviceAccountEmail As String,
                       serviceAccountPrivateKeyRaw As String,
                       serviceAccountTokenUri As String,
                       projectId As String,
                       endpoint As String,
                       location As String,
                       recognizerId As String,
                       defaultModel As String,
                       defaultLanguage As String,
                       Optional debugSource As String = "")
            Me.New(serviceAccountEmail,
                   serviceAccountPrivateKeyRaw,
                   serviceAccountTokenUri,
                   projectId,
                   endpoint,
                   location,
                   recognizerId,
                   defaultModel,
                   defaultLanguage,
                   "",
                   debugSource)
        End Sub

        Public Sub New(serviceAccountEmail As String,
                       serviceAccountPrivateKeyRaw As String,
                       serviceAccountTokenUri As String,
                       projectId As String,
                       endpoint As String,
                       location As String,
                       recognizerId As String,
                       defaultModel As String,
                       defaultLanguage As String,
                       oauthScopes As String,
                       Optional debugSource As String = "")

            _serviceAccountEmail = serviceAccountEmail
            _serviceAccountPrivateKeyRaw = serviceAccountPrivateKeyRaw
            _serviceAccountTokenUri = serviceAccountTokenUri
            _projectId = If(projectId, "").Trim()
            _endpoint = If(endpoint, "").Trim()
            _location = If(location, "").Trim()
            _recognizerId = If(recognizerId, "").Trim()
            _defaultModel = If(defaultModel, "").Trim()
            _defaultLanguage = If(defaultLanguage, "").Trim()
            _oauthScopes = If(oauthScopes, "").Trim()
            _debugSource = If(debugSource, "").Trim()
        End Sub

        Private Function GetDebugPrefix() As String
            If String.IsNullOrWhiteSpace(_debugSource) Then
                Return "[GoogleV2]"
            End If

            Return "[GoogleV2|" & _debugSource & "]"
        End Function

        Private Function GetEffectiveProjectId() As String
            If Not String.IsNullOrWhiteSpace(_projectId) Then
                Return _projectId
            End If

            Throw New InvalidOperationException("Google V2 project id is missing.")
        End Function

        Private Function GetEffectiveEndpoint() As String
            If Not String.IsNullOrWhiteSpace(_endpoint) Then
                Return _endpoint
            End If

            Return GoogleV2Endpoint
        End Function

        Private Function GetEffectiveLocation() As String
            If Not String.IsNullOrWhiteSpace(_location) Then
                Return _location
            End If

            Return GoogleV2Location
        End Function

        Private Function GetEffectiveRecognizerId() As String
            If Not String.IsNullOrWhiteSpace(_recognizerId) Then
                Return _recognizerId
            End If

            Return GoogleV2RecognizerId
        End Function

        Private Function GetEffectiveDefaultModel() As String
            If Not String.IsNullOrWhiteSpace(_defaultModel) Then
                Return _defaultModel
            End If

            Return GoogleV2DefaultModel
        End Function

        Private Function GetEffectiveDefaultLanguage() As String
            If Not String.IsNullOrWhiteSpace(_defaultLanguage) Then
                Return _defaultLanguage
            End If

            Return GoogleV2DefaultLanguage
        End Function

        Private Function GetEffectiveTokenUri() As String
            If Not String.IsNullOrWhiteSpace(_serviceAccountTokenUri) Then
                Return _serviceAccountTokenUri
            End If

            Return "https://oauth2.googleapis.com/token"
        End Function

        Private Function GetEffectiveOAuthScopes() As String
            If Not String.IsNullOrWhiteSpace(_oauthScopes) Then
                Return _oauthScopes
            End If

            Return DefaultGoogleOAuthScope
        End Function

        Private Function NormalizeConfiguredLanguage(rawLanguage As String) As String
            Dim normalized As String = If(rawLanguage, "").Trim()

            If normalized.Length = 0 OrElse String.Equals(normalized, "auto", StringComparison.OrdinalIgnoreCase) Then
                Return GetEffectiveDefaultLanguage()
            End If

            Return normalized
        End Function

        Private Sub RaiseStatusMessage(message As String, Optional progressPercent As System.Nullable(Of Integer) = Nothing)
            RaiseEvent Status(Me, New TranscriptionStatusEventArgs(message, progressPercent))
        End Sub

        Private Function GetRecognizerName() As String
            Return "projects/" & GetEffectiveProjectId() & "/locations/" & GetEffectiveLocation() & "/recognizers/" & GetEffectiveRecognizerId()
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

        Private Async Function GetFreshAccessTokenAsync() As Task(Of String)
            If Not String.IsNullOrWhiteSpace(_accessToken) AndAlso DateTime.UtcNow < _accessTokenExpiresUtc Then
                Return _accessToken
            End If

            Await _googleOAuthSem.WaitAsync()

            Try
                If Not String.IsNullOrWhiteSpace(_accessToken) AndAlso DateTime.UtcNow < _accessTokenExpiresUtc Then
                    Return _accessToken
                End If

                GoogleOAuthHelper.client_email = _serviceAccountEmail
                GoogleOAuthHelper.private_key = FormatPrivateKey(_serviceAccountPrivateKeyRaw)
                GoogleOAuthHelper.scopes = GetEffectiveOAuthScopes()
                GoogleOAuthHelper.token_uri = GetEffectiveTokenUri()
                GoogleOAuthHelper.token_life = OAuthTokenLifetimeSeconds

                Dim token As String = Await GoogleOAuthHelper.GetAccessToken()

                If String.IsNullOrWhiteSpace(token) Then
                    Throw New InvalidOperationException("Google V2 access token could not be obtained.")
                End If

                _accessToken = token
                _accessTokenExpiresUtc = DateTime.UtcNow.AddSeconds(Math.Max(60, OAuthTokenLifetimeSeconds - OAuthRefreshSkewSeconds))

                RaiseStatusMessage("Google V2 OAuth2 access token refreshed.")
                System.Diagnostics.Debug.WriteLine(GetDebugPrefix() & " OAuth token refreshed. Expires=" & _accessTokenExpiresUtc.ToString("u"))

                Return _accessToken
            Finally
                _googleOAuthSem.Release()
            End Try
        End Function

        Private Async Function BuildClient() As Task
            Dim token As String = Await GetFreshAccessTokenAsync()

            Dim callCreds As CallCredentials =
                CallCredentials.FromInterceptor(
                    Async Function(c, md)
                        md.Add("Authorization", "Bearer " & token)
                        Await Task.CompletedTask
                    End Function)

            Dim chCreds As ChannelCredentials = ChannelCredentials.Create(ChannelCredentials.SecureSsl, callCreds)

            Dim builder As New SpeechV2.SpeechClientBuilder() With {
                .Endpoint = GetEffectiveEndpoint(),
                .ChannelCredentials = chCreds,
                .GrpcAdapter = GrpcCoreAdapter.Instance
            }

            _client = builder.Build()

            RaiseStatusMessage("Google V2 client built with refreshed OAuth2 token.")
            RaiseStatusMessage("Google V2 service account: " & _serviceAccountEmail)

            System.Diagnostics.Debug.WriteLine(GetDebugPrefix() & " Client built with bearer token.")
            System.Diagnostics.Debug.WriteLine(GetDebugPrefix() & " ServiceAccount=" & _serviceAccountEmail)
            System.Diagnostics.Debug.WriteLine(GetDebugPrefix() & " TokenUri=" & GetEffectiveTokenUri())
            System.Diagnostics.Debug.WriteLine(GetDebugPrefix() & " EffectiveEndpoint=" & GetEffectiveEndpoint())
            System.Diagnostics.Debug.WriteLine(GetDebugPrefix() & " EffectiveProjectId=" & GetEffectiveProjectId())
            System.Diagnostics.Debug.WriteLine(GetDebugPrefix() & " EffectiveLocation=" & GetEffectiveLocation())
            System.Diagnostics.Debug.WriteLine(GetDebugPrefix() & " EffectiveRecognizerName=" & GetRecognizerName())
            System.Diagnostics.Debug.WriteLine(GetDebugPrefix() & " EffectiveScopes=" & GetEffectiveOAuthScopes())
        End Function

        Private Function BuildRecognitionConfig(opts As TranscriptionOptions) As SpeechV2.RecognitionConfig
            Dim effectiveOpts As TranscriptionOptions = If(opts, New TranscriptionOptions())
            Dim effectiveModel As String = GetEffectiveDefaultModel()
            Dim effectiveLanguage As String = GetEffectiveDefaultLanguage()
            Dim multiChannel As Boolean = effectiveOpts.MultiChannelDiarization

            If Not String.IsNullOrWhiteSpace(effectiveOpts.Model) Then
                effectiveModel = effectiveOpts.Model.Trim()
            End If

            effectiveLanguage = NormalizeConfiguredLanguage(effectiveOpts.LanguageCode)

            Dim cfg As New SpeechV2.RecognitionConfig With {
                .Model = effectiveModel,
                .ExplicitDecodingConfig = New SpeechV2.ExplicitDecodingConfig With {
                    .Encoding = SpeechV2.ExplicitDecodingConfig.Types.AudioEncoding.Linear16,
                    .SampleRateHertz = 16000,
                    .AudioChannelCount = If(multiChannel, 2, 1)
                },
                .Features = New SpeechV2.RecognitionFeatures With {
                    .EnableAutomaticPunctuation = True,
                    .EnableWordTimeOffsets = False,
                    .MultiChannelMode = If(
                        multiChannel,
                        SpeechV2.RecognitionFeatures.Types.MultiChannelMode.SeparateRecognitionPerChannel,
                        SpeechV2.RecognitionFeatures.Types.MultiChannelMode.Unspecified)
                }
            }

            cfg.LanguageCodes.Add(effectiveLanguage)

            System.Diagnostics.Debug.WriteLine(
                GetDebugPrefix() & " BuildRecognitionConfig " &
                "ConfiguredLanguage=" & If(String.IsNullOrWhiteSpace(effectiveOpts.LanguageCode), "(empty)", effectiveOpts.LanguageCode) &
                "; EffectiveLanguage=" & effectiveLanguage &
                "; ConfiguredModel=" & If(String.IsNullOrWhiteSpace(effectiveOpts.Model), "(empty)", effectiveOpts.Model) &
                "; EffectiveModel=" & effectiveModel &
                "; MultiChannelDiarization=" & If(multiChannel, "True", "False") &
                "; AudioChannelCount=" & cfg.ExplicitDecodingConfig.AudioChannelCount.ToString() &
                "; MultiChannelMode=" & cfg.Features.MultiChannelMode.ToString())

            Return cfg
        End Function

        Private Async Function BuildAndConfigureStreamAsync(opts As TranscriptionOptions) As Task(Of SpeechV2.SpeechClient.StreamingRecognizeStream)
            Dim effectiveOpts As TranscriptionOptions = If(opts, New TranscriptionOptions())

            Await BuildClient()

            Dim stream As SpeechV2.SpeechClient.StreamingRecognizeStream = _client.StreamingRecognize()

            Dim req As New SpeechV2.StreamingRecognizeRequest With {
                .Recognizer = GetRecognizerName(),
                .StreamingConfig = New SpeechV2.StreamingRecognitionConfig With {
                    .Config = BuildRecognitionConfig(effectiveOpts),
                    .StreamingFeatures = New SpeechV2.StreamingRecognitionFeatures With {
                        .InterimResults = True
                    }
                }
            }

            System.Diagnostics.Debug.WriteLine(GetDebugPrefix() & " Request.Recognizer=" & req.Recognizer)
            System.Diagnostics.Debug.WriteLine(GetDebugPrefix() & " Request.Model=" & req.StreamingConfig.Config.Model)
            System.Diagnostics.Debug.WriteLine(GetDebugPrefix() & " Request.LanguageCodes=" & String.Join(",", req.StreamingConfig.Config.LanguageCodes))
            System.Diagnostics.Debug.WriteLine(GetDebugPrefix() & " Request.AudioChannelCount=" & req.StreamingConfig.Config.ExplicitDecodingConfig.AudioChannelCount.ToString())
            System.Diagnostics.Debug.WriteLine(GetDebugPrefix() & " Request.MultiChannelMode=" & req.StreamingConfig.Config.Features.MultiChannelMode.ToString())

            RaiseStatusMessage("Opening Google V2 streaming session…")
            RaiseStatusMessage("Recognizer=" & req.Recognizer)
            RaiseStatusMessage("Language=" & If(String.IsNullOrWhiteSpace(effectiveOpts.LanguageCode), GoogleV2DefaultLanguage & "(default)", effectiveOpts.LanguageCode))
            RaiseStatusMessage("Model=" & If(String.IsNullOrWhiteSpace(effectiveOpts.Model), GoogleV2DefaultModel & "(default)", effectiveOpts.Model))

            Await stream.WriteAsync(req)

            _streamStarted = DateTime.UtcNow
            _completed = False

            RaiseStatusMessage("Google V2 stream configured.")

            Return stream
        End Function

        Public Async Function StartLiveAsync(opts As TranscriptionOptions, ct As CancellationToken) As Task Implements ITranscriptionEngine.StartLiveAsync
            _opts = If(opts, New TranscriptionOptions())
            _firstResponseSeen = False
            _completed = False
            _stopping = False
            _lastPartial = String.Empty
            _lastFinalResultKey = String.Empty
            _lastFinalResultUtc = DateTime.MinValue

            SyncLock _ringBuffer
                _ringBuffer.Clear()
            End SyncLock

            Dim initialQueue As New BlockingCollection(Of ByteString)()
            Dim initialReaderCts As New CancellationTokenSource()

            SyncLock _stateSyncRoot
                _audioQueue = initialQueue
                _readerCts = initialReaderCts
                _writerTask = Nothing
                _readerTask = Nothing
                _stream = Nothing
            End SyncLock

            RaiseStatusMessage("Starting Google V2…")
            RaiseStatusMessage("Project=" & GetEffectiveProjectId() & ", Location=" & GetEffectiveLocation() & ", Endpoint=" & GetEffectiveEndpoint())

            System.Diagnostics.Debug.WriteLine(GetDebugPrefix() & " StartLiveAsync entered.")
            System.Diagnostics.Debug.WriteLine(GetDebugPrefix() & " CancellationRequested=" & ct.IsCancellationRequested.ToString())
            System.Diagnostics.Debug.WriteLine(
                GetDebugPrefix() & " IncomingOpts " &
                "LanguageCode=" & If(String.IsNullOrWhiteSpace(_opts.LanguageCode), "(empty)", _opts.LanguageCode) &
                "; Model=" & If(String.IsNullOrWhiteSpace(_opts.Model), "(empty)", _opts.Model) &
                "; MultiChannelDiarization=" & If(_opts.MultiChannelDiarization, "True", "False"))

            Try
                Dim stream As SpeechV2.SpeechClient.StreamingRecognizeStream = Await BuildAndConfigureStreamAsync(_opts)
                Dim writer As Task = StartWriter(initialQueue, stream)
                Dim reader As Task = StartReader(stream, initialReaderCts.Token)

                SyncLock _stateSyncRoot
                    _stream = stream
                    _writerTask = writer
                    _readerTask = reader
                End SyncLock
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine(GetDebugPrefix() & " StartLiveAsync EXCEPTION: " & ex.ToString())
                RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("Google V2 start failed: " & ex.Message, ex, True))
                Throw
            End Try
        End Function

        Public Function PushAudioAsync(pcm As Byte(), bytesValid As Integer, ct As CancellationToken) As Task Implements ITranscriptionEngine.PushAudioAsync
            If _stopping OrElse _completed OrElse pcm Is Nothing OrElse bytesValid <= 0 Then
                Return Task.CompletedTask
            End If

            Dim chunk As ByteString = ByteString.CopyFrom(pcm, 0, bytesValid)

            SyncLock _ringBuffer
                _ringBuffer.Enqueue(chunk)

                If _ringBuffer.Count > RingBufferSize Then
                    _ringBuffer.Dequeue()
                End If
            End SyncLock

            For attempt As Integer = 0 To 1
                Dim targetQueue As BlockingCollection(Of ByteString) = Nothing

                SyncLock _stateSyncRoot
                    targetQueue = _audioQueue
                End SyncLock

                If targetQueue Is Nothing OrElse targetQueue.IsAddingCompleted Then
                    Exit For
                End If

                Try
                    targetQueue.Add(chunk, ct)
                    Exit For
                Catch ex As InvalidOperationException
                    If attempt = 1 Then
                        Exit For
                    End If
                Catch ex As OperationCanceledException
                    Exit For
                End Try
            Next

            If (DateTime.UtcNow - _streamStarted).TotalMilliseconds > StreamingLimitMs Then
                _streamStarted = DateTime.UtcNow

                Task.Run(
                    Async Function()
                        Await TryRecoverAsync()
                    End Function)
            End If

            Return Task.CompletedTask
        End Function

        Public Async Function StopLiveAsync() As Task Implements ITranscriptionEngine.StopLiveAsync
            Await _recoverySem.WaitAsync()

            Try
                _stopping = True
                _completed = True

                Dim queueToStop As BlockingCollection(Of ByteString) = Nothing
                Dim writerToStop As Task = Nothing
                Dim readerToStop As Task = Nothing
                Dim streamToStop As SpeechV2.SpeechClient.StreamingRecognizeStream = Nothing
                Dim readerCtsToStop As CancellationTokenSource = Nothing

                SyncLock _stateSyncRoot
                    queueToStop = _audioQueue
                    writerToStop = _writerTask
                    readerToStop = _readerTask
                    streamToStop = _stream
                    readerCtsToStop = _readerCts

                    _audioQueue = Nothing
                    _writerTask = Nothing
                    _readerTask = Nothing
                    _stream = Nothing
                    _readerCts = Nothing
                End SyncLock

                Try
                    If readerCtsToStop IsNot Nothing Then
                        readerCtsToStop.Cancel()
                    End If
                Catch
                End Try

                Await SafeCompleteStreamAsync(queueToStop, writerToStop, streamToStop)

                If readerToStop IsNot Nothing Then
                    Try
                        Await readerToStop
                    Catch ex As Exception
                        RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("Google V2 reader shutdown failed: " & ex.Message, ex, False))
                    End Try
                End If
            Finally
                _recoverySem.Release()
            End Try

            RaiseStatusMessage("Google V2 stopped.")
        End Function

        Public Async Function TranscribeFileAsync(filePath As String, opts As TranscriptionOptions, ct As CancellationToken) As Task Implements ITranscriptionEngine.TranscribeFileAsync
            RaiseStatusMessage("Preparing file for Google V2…", 0)
            Await StartLiveAsync(opts, ct)

            Dim pcm As Byte() = VoskEngine.LoadAudioToPcm16Mono16k(filePath)
            Const chunkSize As Integer = 4096
            Dim bytesPerSec As Integer = 16000 * 2
            Dim pos As Integer = 0
            Dim lastProgressPercent As Integer = -1

            RaiseStatusMessage("Streaming file to Google V2…", 0)

            While pos < pcm.Length AndAlso Not ct.IsCancellationRequested
                Dim n As Integer = Math.Min(chunkSize, pcm.Length - pos)
                Dim slice(n - 1) As Byte
                Buffer.BlockCopy(pcm, pos, slice, 0, n)
                Await PushAudioAsync(slice, n, ct)
                pos += n

                Dim progressPercent As Integer =
                    CInt(Math.Truncate((CDbl(pos) / Math.Max(1.0R, CDbl(pcm.Length))) * 100.0R))

                If progressPercent <> lastProgressPercent Then
                    lastProgressPercent = progressPercent
                    RaiseStatusMessage("Streaming file to Google V2…", progressPercent)
                End If

                Await Task.Delay(CInt(1000.0 * n / bytesPerSec), ct)
            End While

            If ct.IsCancellationRequested Then
                RaiseStatusMessage("Google V2 file transcription canceled.")
            Else
                RaiseStatusMessage("Finalizing Google V2 file transcription…", 100)
            End If

            Await StopLiveAsync()

            If Not ct.IsCancellationRequested Then
                RaiseStatusMessage("Google V2 file transcription completed.", 100)
            End If
        End Function

        Private Function StartWriter(queue As BlockingCollection(Of ByteString),
                                     stream As SpeechV2.SpeechClient.StreamingRecognizeStream) As Task
            Return Task.Run(
                Async Function()
                    Try
                        For Each chunk As ByteString In queue.GetConsumingEnumerable()
                            If stream Is Nothing Then
                                Exit For
                            End If

                            Try
                                Await stream.WriteAsync(New SpeechV2.StreamingRecognizeRequest With {
                                    .Audio = chunk
                                })
                            Catch ex As Exception
                                If Not _stopping Then
                                    RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("Google V2 audio write failed: " & ex.Message, ex, False))
                                End If

                                Exit For
                            End Try
                        Next
                    Catch ex As Exception
                        If Not _stopping Then
                            RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("Google V2 writer failed: " & ex.Message, ex, False))
                        End If
                    End Try
                End Function)
        End Function

        Private Function StartReader(stream As SpeechV2.SpeechClient.StreamingRecognizeStream,
                                     token As CancellationToken) As Task
            Return Task.Run(
                Async Function()
                    Try
                        Dim it = stream.GetResponseStream().GetAsyncEnumerator(token)

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

                                If txt.Length = 0 Then
                                    Continue For
                                End If

                                Dim spk As String = String.Empty

                                If r.ChannelTag > 0 Then
                                    spk = "Speaker " & r.ChannelTag.ToString()
                                End If

                                If r.IsFinal Then
                                    _lastPartial = String.Empty

                                    If ShouldEmitFinalResult(txt, spk) Then
                                        RaiseEvent FinalResult(Me, New TranscriptionEventArgs(txt, True, spk))
                                    Else
                                        System.Diagnostics.Debug.WriteLine(GetDebugPrefix() & " Duplicate replayed final ignored: " & txt)
                                    End If
                                Else
                                    _lastPartial = txt
                                    RaiseEvent PartialResult(Me, New TranscriptionEventArgs(txt, False, spk))
                                End If
                            Next
                        End While
                    Catch ocex As OperationCanceledException
                    Catch rex As RpcException
                        If Not token.IsCancellationRequested AndAlso Not _stopping AndAlso rex.StatusCode <> StatusCode.Cancelled Then
                            RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("Google V2 RPC error (" & rex.StatusCode.ToString() & "): " & rex.Status.Detail, rex, False))

                            Task.Run(
                                Async Function()
                                    Await TryRecoverAsync()
                                End Function)
                        End If
                    Catch ex As Exception
                        If Not _stopping Then
                            RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("Google V2 reader failed: " & ex.Message, ex, False))

                            Task.Run(
                                Async Function()
                                    Await TryRecoverAsync()
                                End Function)
                        End If
                    End Try
                End Function)
        End Function

        Private Async Function TryRecoverAsync() As Task
            If _stopping Then
                Return
            End If

            If Not Await _recoverySem.WaitAsync(0) Then
                Return
            End If

            Try
                If _stopping Then
                    Return
                End If

                RaiseStatusMessage("Google V2 restarting live stream…")

                Dim replayChunks As List(Of ByteString) = GetRingBufferSnapshot()
                Dim newBackingQueue As New ConcurrentQueue(Of ByteString)()

                For Each chunk As ByteString In replayChunks
                    newBackingQueue.Enqueue(chunk)
                Next

                Dim newAudioQueue As New BlockingCollection(Of ByteString)(newBackingQueue)
                Dim newReaderCts As New CancellationTokenSource()

                Dim oldQueue As BlockingCollection(Of ByteString) = Nothing
                Dim oldWriter As Task = Nothing
                Dim oldReader As Task = Nothing
                Dim oldStream As SpeechV2.SpeechClient.StreamingRecognizeStream = Nothing
                Dim oldReaderCts As CancellationTokenSource = Nothing

                SyncLock _stateSyncRoot
                    oldQueue = _audioQueue
                    oldWriter = _writerTask
                    oldReader = _readerTask
                    oldStream = _stream
                    oldReaderCts = _readerCts

                    _audioQueue = newAudioQueue
                    _writerTask = Nothing
                    _readerTask = Nothing
                    _stream = Nothing
                    _readerCts = newReaderCts
                End SyncLock

                Try
                    If oldReaderCts IsNot Nothing Then
                        oldReaderCts.Cancel()
                    End If
                Catch
                End Try

                Await SafeCompleteStreamAsync(oldQueue, oldWriter, oldStream)

                If oldReader IsNot Nothing Then
                    Try
                        Await oldReader
                    Catch
                    End Try
                End If

                If _stopping Then
                    Try
                        newAudioQueue.CompleteAdding()
                    Catch
                    End Try

                    Return
                End If

                _firstResponseSeen = False

                Dim newStream As SpeechV2.SpeechClient.StreamingRecognizeStream = Await BuildAndConfigureStreamAsync(_opts)
                Dim newWriter As Task = StartWriter(newAudioQueue, newStream)
                Dim newReader As Task = StartReader(newStream, newReaderCts.Token)

                SyncLock _stateSyncRoot
                    _stream = newStream
                    _writerTask = newWriter
                    _readerTask = newReader
                End SyncLock

                RaiseStatusMessage("Google V2 live stream restarted.")
                System.Diagnostics.Debug.WriteLine(GetDebugPrefix() & " Stream restarted. ReplayChunks=" & replayChunks.Count.ToString())
            Catch ex As Exception
                RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("Google V2 recovery failed: " & ex.Message, ex, False))
            Finally
                _recoverySem.Release()
            End Try
        End Function

        Private Function GetRingBufferSnapshot() As List(Of ByteString)
            Dim snapshot As New List(Of ByteString)()

            SyncLock _ringBuffer
                snapshot.AddRange(_ringBuffer)
            End SyncLock

            Return snapshot
        End Function

        Private Shared Function BuildFinalResultKey(text As String, speaker As String) As String
            Return If(speaker, "").Trim() & "|" & If(text, "").Trim().ToLowerInvariant()
        End Function

        Private Function ShouldEmitFinalResult(text As String, speaker As String) As Boolean
            Dim key As String = BuildFinalResultKey(text, speaker)
            Dim now As DateTime = DateTime.UtcNow

            SyncLock _stateSyncRoot
                If key = _lastFinalResultKey AndAlso (now - _lastFinalResultUtc).TotalSeconds <= DuplicateFinalWindowSeconds Then
                    Return False
                End If

                _lastFinalResultKey = key
                _lastFinalResultUtc = now
            End SyncLock

            Return True
        End Function

        Private Async Function SafeCompleteStreamAsync(targetQueue As BlockingCollection(Of ByteString),
                                                       targetWriter As Task,
                                                       targetStream As SpeechV2.SpeechClient.StreamingRecognizeStream) As Task
            _completed = True

            If targetQueue IsNot Nothing AndAlso Not targetQueue.IsAddingCompleted Then
                Try
                    targetQueue.CompleteAdding()
                Catch
                End Try
            End If

            If targetWriter IsNot Nothing Then
                Try
                    Await targetWriter
                Catch
                End Try
            End If

            Try
                If targetStream IsNot Nothing Then
                    Await targetStream.WriteCompleteAsync()
                End If
            Catch
            End Try

            Try
                If targetStream IsNot Nothing Then
                    targetStream.Dispose()
                End If
            Catch
            End Try
        End Function

        Public Function DisposeAsync() As ValueTask Implements IAsyncDisposable.DisposeAsync
            Return New ValueTask(DisposeAsyncCore())
        End Function

        Private Async Function DisposeAsyncCore() As Task
            Await StopLiveAsync()
        End Function
    End Class

End Namespace
