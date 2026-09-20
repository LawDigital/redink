' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: GeminiTranscribeEngine.vb
' Purpose: File-based transcription using Google Gemini 3.5 Transcribe on
'          Vertex AI / Gemini Enterprise Agent Platform.
' =============================================================================

Option Explicit On
Option Strict On

Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports Newtonsoft.Json.Linq
Imports NAudio.Wave

Namespace Transcription

    Public Class GeminiTranscribeEngine
        Implements ITranscriptionEngine

        Public Const DisplayNameValue As System.String = "Google Gemini 3.5 Transcribe (file)"
        Public Const DefaultVertexModel As System.String = "gemini-3.5-transcribe-preview"
        Public Const DefaultGeminiApiModel As System.String = "gemini-3.5-transcribe"
        Public Const DefaultLocation As System.String = "global"

        Private Shared ReadOnly _supportedLanguages As System.String() = {
            "auto", "af-ZA", "am-ET", "ar-EG", "hy-AM", "as-IN", "az-AZ", "be-BY", "bn-BD", "bn-IN",
            "bs-BA", "bg-BG", "rup-BG", "my-MM", "yue-Hant-HK", "ca-ES", "ceb", "km-KH", "hr-HR", "cs-CZ",
            "da-DK", "nl-NL", "en-AU", "en-GB", "en-IN", "en-US", "et-EE", "fa-IR", "fil-PH", "fi-FI",
            "fr-FR", "fr-CA", "gl-ES", "ka-GE", "de-DE", "el-GR", "gu-IN", "ha-NG", "he-IL", "hi-IN",
            "hu-HU", "is-IS", "id-ID", "it-IT", "ja-JP", "jv-ID", "kea-CV", "kn-IN", "kk-KZ", "ko-KR",
            "ky-KG", "lv-LV", "ln-CD", "lt-LT", "mk-MK", "ms-MY", "ml-IN", "mt-MT", "cmn-Hans-CN", "mr-IN",
            "mn-MN", "ne-NP", "nb-NO", "or-IN", "pl-PL", "pt-BR", "pt-PT", "pa-IN", "pa-Guru-IN", "ro-RO",
            "ru-RU", "sr-RS", "sd-Arab-IN", "sk-SK", "sl-SI", "es-419", "es-ES", "es-US", "sw-KE", "sv-SE",
            "tg-TJ", "te-IN", "th-TH", "tr-TR", "uk-UA", "uz-UZ", "vi-VN"
        }

        Public Shared ReadOnly Property SupportedLanguages As System.String()
            Get
                Return _supportedLanguages.OrderBy(Function(x As System.String) If(System.String.Equals(x, "auto", System.StringComparison.OrdinalIgnoreCase), "", x), System.StringComparer.OrdinalIgnoreCase).ToArray()
            End Get
        End Property

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
        Private ReadOnly _wordTimestamp As System.Boolean
        Private ReadOnly _customVocabulary As System.Collections.Generic.List(Of System.String)
        Private ReadOnly _endpointOverride As System.String
        Private ReadOnly _http As System.Net.Http.HttpClient

        Public Sub New(projectId As System.String,
                       tokenFactory As System.Func(Of System.Threading.Tasks.Task(Of System.String)),
                       Optional model As System.String = DefaultVertexModel,
                       Optional location As System.String = DefaultLocation,
                       Optional mode As System.String = "VERBATIM",
                       Optional wordTimestamp As System.Boolean = False,
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
            _wordTimestamp = wordTimestamp
            _customVocabulary = If(customVocabulary, System.Linq.Enumerable.Empty(Of System.String)()).
                Where(Function(x As System.String) Not System.String.IsNullOrWhiteSpace(x)).
                Select(Function(x As System.String) x.Trim()).
                Distinct(System.StringComparer.OrdinalIgnoreCase).
                Take(1000).
                ToList()

            If System.String.IsNullOrWhiteSpace(_projectId) Then
                Throw New System.ArgumentException("Google Cloud project id is required.", NameOf(projectId))
            End If

            If _tokenFactory Is Nothing Then
                Throw New System.ArgumentNullException(NameOf(tokenFactory))
            End If

            EnsureTls12()
            Dim handler As New System.Net.Http.HttpClientHandler() With {
                .AllowAutoRedirect = True,
                .AutomaticDecompression = System.Net.DecompressionMethods.GZip Or System.Net.DecompressionMethods.Deflate
            }
            _http = New System.Net.Http.HttpClient(handler) With {.Timeout = System.TimeSpan.FromMinutes(30)}
        End Sub

        Public Sub New(apiKey As System.String,
                       Optional model As System.String = DefaultGeminiApiModel,
                       Optional mode As System.String = "VERBATIM",
                       Optional wordTimestamp As System.Boolean = False,
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
            _wordTimestamp = wordTimestamp
            _customVocabulary = If(customVocabulary, System.Linq.Enumerable.Empty(Of System.String)()).
                Where(Function(x As System.String) Not System.String.IsNullOrWhiteSpace(x)).
                Select(Function(x As System.String) x.Trim()).
                Distinct(System.StringComparer.OrdinalIgnoreCase).
                Take(1000).
                ToList()

            If System.String.IsNullOrWhiteSpace(_apiKey) Then
                Throw New System.ArgumentException("Gemini API key is required.", NameOf(apiKey))
            End If

            EnsureTls12()
            Dim handler As New System.Net.Http.HttpClientHandler() With {
                .AllowAutoRedirect = True,
                .AutomaticDecompression = System.Net.DecompressionMethods.GZip Or System.Net.DecompressionMethods.Deflate
            }
            _http = New System.Net.Http.HttpClient(handler) With {.Timeout = System.TimeSpan.FromMinutes(30)}
        End Sub

        Public ReadOnly Property Name As System.String Implements ITranscriptionEngine.Name
            Get
                Return DisplayNameValue
            End Get
        End Property

        Public ReadOnly Property Kind As EngineKind Implements ITranscriptionEngine.Kind
            Get
                Return EngineKind.GeminiTranscribe
            End Get
        End Property

        Public ReadOnly Property SupportsLiveStreaming As System.Boolean Implements ITranscriptionEngine.SupportsLiveStreaming
            Get
                Return False
            End Get
        End Property

        Public ReadOnly Property SupportsFileTranscription As System.Boolean Implements ITranscriptionEngine.SupportsFileTranscription
            Get
                Return True
            End Get
        End Property

        Public ReadOnly Property SupportsDiarization As System.Boolean Implements ITranscriptionEngine.SupportsDiarization
            Get
                Return True
            End Get
        End Property

        Public ReadOnly Property SupportsMultiChannelDiarization As System.Boolean Implements ITranscriptionEngine.SupportsMultiChannelDiarization
            Get
                Return False
            End Get
        End Property

        Public Function StartLiveAsync(opts As TranscriptionOptions, ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task Implements ITranscriptionEngine.StartLiveAsync
            Throw New System.NotSupportedException("Gemini 3.5 Transcribe file mode is request/response only. The Google live-preview model uses a separate bidirectional Live API contract.")
        End Function

        Public Function PushAudioAsync(pcm As System.Byte(), bytesValid As System.Int32, ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task Implements ITranscriptionEngine.PushAudioAsync
            Return System.Threading.Tasks.Task.CompletedTask
        End Function

        Public Function StopLiveAsync() As System.Threading.Tasks.Task Implements ITranscriptionEngine.StopLiveAsync
            Return System.Threading.Tasks.Task.CompletedTask
        End Function

        Private Const MaximumRequestAudioSeconds As System.Double = 15.0R * 60.0R
        Private Const MaximumDecodedChunkBytes As System.Int64 = 48L * 1024L * 1024L

        Public Async Function TranscribeFileAsync(filePath As System.String, opts As TranscriptionOptions, ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task Implements ITranscriptionEngine.TranscribeFileAsync
            If System.String.IsNullOrWhiteSpace(filePath) OrElse Not System.IO.File.Exists(filePath) Then
                Throw New System.IO.FileNotFoundException("Audio file not found.", filePath)
            End If

            ct.ThrowIfCancellationRequested()

            Dim duration As System.TimeSpan = GetAudioDuration(filePath)
            Dim mustSplitByDuration As System.Boolean = duration.TotalSeconds > MaximumRequestAudioSeconds
            Dim mustSplitByPayload As System.Boolean = New System.IO.FileInfo(filePath).Length > MaximumDecodedChunkBytes

            If Not mustSplitByDuration AndAlso Not mustSplitByPayload Then
                RaiseStatusMessage("Preparing audio for Gemini 3.5 Transcribe…")
                Dim originalBytes As System.Byte() = System.IO.File.ReadAllBytes(filePath)
                Await TranscribePayloadAsync(originalBytes, ResolveMimeType(filePath), opts, ct, Nothing, Nothing)
                RaiseStatusMessage("Gemini 3.5 Transcribe file transcription completed.")
                Return
            End If

            RaiseStatusMessage("Audio exceeds the safe single-request duration/payload. Splitting into PCM/WAV request chunks without lossy re-encoding…")

            Using reader As New NAudio.Wave.MediaFoundationReader(filePath)
                Dim waveFormat As NAudio.Wave.WaveFormat = reader.WaveFormat
                If waveFormat Is Nothing OrElse waveFormat.AverageBytesPerSecond <= 0 OrElse waveFormat.BlockAlign <= 0 Then
                    Throw New System.InvalidOperationException("Unable to determine a stable decoded audio format for Gemini chunking.")
                End If

                Dim maxBytesByTime As System.Int64 = CLng(System.Math.Floor(MaximumRequestAudioSeconds * CDbl(waveFormat.AverageBytesPerSecond)))
                Dim chunkDataBytes As System.Int64 = System.Math.Min(maxBytesByTime, MaximumDecodedChunkBytes)
                chunkDataBytes -= (chunkDataBytes Mod CLng(waveFormat.BlockAlign))
                If chunkDataBytes <= 0L Then
                    Throw New System.InvalidOperationException("Unable to calculate a valid Gemini transcription chunk size.")
                End If

                Dim estimatedDecodedLength As System.Int64 = System.Math.Max(0L, reader.Length)
                Dim estimatedChunks As System.Int32 = If(estimatedDecodedLength > 0L,
                    CInt(System.Math.Ceiling(CDbl(estimatedDecodedLength) / CDbl(chunkDataBytes))),
                    CInt(System.Math.Max(1.0R, System.Math.Ceiling(duration.TotalSeconds / MaximumRequestAudioSeconds))))

                Dim chunkIndex As System.Int32 = 0
                Do While reader.Position < reader.Length
                    ct.ThrowIfCancellationRequested()
                    chunkIndex += 1

                    Dim wavBytes As System.Byte() = ReadNextPcmWaveChunk(reader, waveFormat, chunkDataBytes, ct)
                    If wavBytes.Length <= 44 Then
                        Exit Do
                    End If

                    RaiseStatusMessage("Gemini 3.5 Transcribe chunk " & chunkIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                                       "/" & estimatedChunks.ToString(System.Globalization.CultureInfo.InvariantCulture) & "…")

                    Await TranscribePayloadAsync(wavBytes,
                                                 "audio/wav",
                                                 opts,
                                                 ct,
                                                 chunkIndex,
                                                 estimatedChunks)
                Loop
            End Using

            RaiseStatusMessage("Gemini 3.5 Transcribe multi-request file transcription completed.")
        End Function

        Private Shared Function GetAudioDuration(filePath As System.String) As System.TimeSpan
            Try
                Using reader As New NAudio.Wave.MediaFoundationReader(filePath)
                    Return reader.TotalTime
                End Using
            Catch ex As System.Exception
                Throw New System.InvalidOperationException("Unable to read audio duration for Gemini transcription: " & ex.Message, ex)
            End Try
        End Function

        Private Shared Function ReadNextPcmWaveChunk(reader As NAudio.Wave.WaveStream,
                                                       waveFormat As NAudio.Wave.WaveFormat,
                                                       maximumDataBytes As System.Int64,
                                                       ct As System.Threading.CancellationToken) As System.Byte()
            Using output As New System.IO.MemoryStream()
                Using writer As New NAudio.Wave.WaveFileWriter(output, waveFormat)
                    Dim bufferSize As System.Int32 = 64 * 1024
                    bufferSize -= (bufferSize Mod waveFormat.BlockAlign)
                    If bufferSize <= 0 Then bufferSize = waveFormat.BlockAlign
                    Dim buffer(bufferSize - 1) As System.Byte
                    Dim written As System.Int64 = 0L

                    Do While written < maximumDataBytes
                        ct.ThrowIfCancellationRequested()

                        Dim remaining As System.Int64 = maximumDataBytes - written
                        Dim requested As System.Int32 = CInt(System.Math.Min(CLng(buffer.Length), remaining))
                        requested -= (requested Mod waveFormat.BlockAlign)
                        If requested <= 0 Then Exit Do

                        Dim read As System.Int32 = reader.Read(buffer, 0, requested)
                        If read <= 0 Then Exit Do

                        If (read Mod waveFormat.BlockAlign) <> 0 Then
                            Throw New System.InvalidOperationException("Decoded audio reader returned a non-block-aligned frame; refusing to drop bytes during Gemini chunking.")
                        End If

                        writer.Write(buffer, 0, read)
                        written += read
                    Loop

                    writer.Flush()
                End Using
                Return output.ToArray()
            End Using
        End Function

        Private Async Function TranscribePayloadAsync(audioBytes As System.Byte(),
                                                       mimeType As System.String,
                                                       opts As TranscriptionOptions,
                                                       ct As System.Threading.CancellationToken,
                                                       chunkIndex As System.Nullable(Of System.Int32),
                                                       chunkCount As System.Nullable(Of System.Int32)) As System.Threading.Tasks.Task
            ct.ThrowIfCancellationRequested()

            Dim accessToken As System.String = System.String.Empty
            If _useVertex Then
                accessToken = Await _tokenFactory()
                If System.String.IsNullOrWhiteSpace(accessToken) Then
                    Throw New System.InvalidOperationException("Google OAuth2 access token is empty.")
                End If
            End If

            Dim requestBody As Newtonsoft.Json.Linq.JObject = BuildRequest(audioBytes, mimeType, opts)
            Dim endpoint As System.String = BuildEndpoint()

            Using request As New System.Net.Http.HttpRequestMessage(System.Net.Http.HttpMethod.Post, endpoint)
                If _useVertex Then
                    request.Headers.Authorization = New System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken)
                Else
                    request.Headers.Add("x-goog-api-key", _apiKey)
                End If
                request.Headers.ExpectContinue = False
                request.Content = New System.Net.Http.StringContent(requestBody.ToString(Newtonsoft.Json.Formatting.None), System.Text.Encoding.UTF8, "application/json")

                If chunkIndex.HasValue AndAlso chunkCount.HasValue Then
                    RaiseStatusMessage("Uploading Gemini transcription chunk " & chunkIndex.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                                       "/" & chunkCount.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) & "…")
                Else
                    RaiseStatusMessage("Uploading audio to Gemini 3.5 Transcribe…")
                End If

                Try
                    Using response As System.Net.Http.HttpResponseMessage = Await _http.SendAsync(request, System.Net.Http.HttpCompletionOption.ResponseContentRead, ct)
                        Dim responseText As System.String = Await response.Content.ReadAsStringAsync()
                        If Not response.IsSuccessStatusCode Then
                            Dim detail As System.String = "Gemini 3.5 Transcribe HTTP " & CInt(response.StatusCode).ToString(System.Globalization.CultureInfo.InvariantCulture) & " " & response.StatusCode.ToString() & ": " & responseText
                            RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs(detail, Nothing, False))
                            Throw New System.InvalidOperationException(detail)
                        End If

                        HandleResponse(responseText)
                    End Using
                Catch ex As System.OperationCanceledException When ct.IsCancellationRequested
                    RaiseStatusMessage("Gemini 3.5 Transcribe canceled.")
                    Throw
                Catch ex As System.Exception
                    If TypeOf ex Is System.InvalidOperationException Then
                        Throw
                    End If
                    Dim detail As System.String = "Gemini 3.5 Transcribe request failed: " & ex.Message
                    RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs(detail, ex, False))
                    Throw New System.InvalidOperationException(detail, ex)
                End Try
            End Using
        End Function

        Private Function BuildRequest(audioBytes As System.Byte(), mimeType As System.String, opts As TranscriptionOptions) As Newtonsoft.Json.Linq.JObject
            Dim transcriptionConfig As New Newtonsoft.Json.Linq.JObject()

            Dim language As System.String = If(opts Is Nothing, System.String.Empty, If(opts.LanguageCode, System.String.Empty)).Trim()
            Dim languageCodes As New Newtonsoft.Json.Linq.JArray()
            If Not System.String.IsNullOrWhiteSpace(language) AndAlso Not System.String.Equals(language, "auto", System.StringComparison.OrdinalIgnoreCase) Then
                languageCodes.Add(language)
            End If
            transcriptionConfig("languageCodes") = languageCodes

            Dim diarization As System.Boolean = opts IsNot Nothing AndAlso opts.EnableDiarization
            Dim effectiveMode As System.String = _mode
            If diarization OrElse _wordTimestamp Then effectiveMode = "VERBATIM"

            transcriptionConfig("mode") = effectiveMode
            If diarization Then transcriptionConfig("diarization") = True
            If _wordTimestamp Then transcriptionConfig("wordTimestamp") = True

            If _customVocabulary.Count > 0 Then
                Dim vocabulary As New Newtonsoft.Json.Linq.JArray()
                For Each term As System.String In _customVocabulary
                    vocabulary.Add(term)
                Next
                transcriptionConfig("customVocabulary") = vocabulary
            End If

            Dim inlineData As New Newtonsoft.Json.Linq.JObject()
            inlineData("mimeType") = mimeType
            inlineData("data") = System.Convert.ToBase64String(audioBytes)

            Dim audioPart As New Newtonsoft.Json.Linq.JObject()
            audioPart("inlineData") = inlineData

            Dim parts As New Newtonsoft.Json.Linq.JArray()
            parts.Add(audioPart)

            Dim content As New Newtonsoft.Json.Linq.JObject()
            content("role") = "user"
            content("parts") = parts

            Dim contents As New Newtonsoft.Json.Linq.JArray()
            contents.Add(content)

            Dim generationConfig As New Newtonsoft.Json.Linq.JObject()
            generationConfig("audioTranscriptionConfig") = transcriptionConfig

            Dim root As New Newtonsoft.Json.Linq.JObject()
            root("contents") = contents
            root("generationConfig") = generationConfig
            Return root
        End Function

        Private Sub HandleResponse(responseText As System.String)
            Dim root As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(responseText)
            Dim parts As Newtonsoft.Json.Linq.JToken = root.SelectToken("candidates[0].content.parts")
            If parts Is Nothing OrElse parts.Type <> Newtonsoft.Json.Linq.JTokenType.Array Then
                Throw New System.InvalidOperationException("Gemini 3.5 Transcribe returned no transcript parts.")
            End If

            Dim emitted As System.Boolean = False
            Dim combined As New System.Text.StringBuilder()

            For Each part As Newtonsoft.Json.Linq.JToken In parts.Children()
                Dim audioTx As Newtonsoft.Json.Linq.JToken = part("audioTranscription")
                Dim speaker As System.String = System.String.Empty
                If audioTx IsNot Nothing AndAlso audioTx("speakerLabel") IsNot Nothing Then speaker = audioTx("speakerLabel").ToString()
                Dim text As System.String = System.String.Empty
                If part("text") IsNot Nothing Then text = part("text").ToString()
                If System.String.IsNullOrWhiteSpace(text) AndAlso audioTx IsNot Nothing Then
                    If audioTx("text") IsNot Nothing Then text = audioTx("text").ToString()
                End If

                If Not System.String.IsNullOrWhiteSpace(text) Then
                    If Not System.String.IsNullOrWhiteSpace(speaker) Then
                        RaiseEvent FinalResult(Me, New TranscriptionEventArgs(text.Trim(), True, speaker.Trim()))
                        emitted = True
                    Else
                        If combined.Length > 0 Then
                            combined.AppendLine()
                        End If
                        combined.Append(text.Trim())
                    End If
                End If
            Next

            If combined.Length > 0 Then
                RaiseEvent FinalResult(Me, New TranscriptionEventArgs(combined.ToString(), True))
                emitted = True
            End If

            If Not emitted Then
                Throw New System.InvalidOperationException("Gemini 3.5 Transcribe response did not contain transcription text.")
            End If
        End Sub

        Private Shared Function NormalizeLocation(requestedLocation As System.String) As System.String
            Dim normalizedLocation As System.String = If(requestedLocation, System.String.Empty).Trim()
            If System.String.IsNullOrWhiteSpace(normalizedLocation) Then
                Return DefaultLocation
            End If
            Return normalizedLocation
        End Function

        Private Function BuildEndpoint() As System.String
            If Not System.String.IsNullOrWhiteSpace(_endpointOverride) Then
                Return ExpandEndpointTemplate(_endpointOverride)
            End If

            If Not _useVertex Then
                Return "https://generativelanguage.googleapis.com/v1beta/models/" & System.Uri.EscapeDataString(_model) & ":generateContent"
            End If

            Dim host As System.String =
                If(System.String.Equals(_location, "global", System.StringComparison.OrdinalIgnoreCase),
                   "aiplatform.googleapis.com",
                   _location & "-aiplatform.googleapis.com")

            Return "https://" & host & "/v1/projects/" & System.Uri.EscapeDataString(_projectId) &
                   "/locations/" & System.Uri.EscapeDataString(_location) &
                   "/publishers/google/models/" & System.Uri.EscapeDataString(_model) & ":generateContent"
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

        Private Shared Function NormalizeMode(value As System.String) As System.String
            Dim normalized As System.String = If(value, System.String.Empty).Trim().ToUpperInvariant()
            If normalized = "SMART" Then
                Return "SMART"
            End If
            Return "VERBATIM"
        End Function

        Private Shared Function ResolveMimeType(filePath As System.String) As System.String
            Select Case System.IO.Path.GetExtension(filePath).ToLowerInvariant()
                Case ".wav"
                    Return "audio/wav"
                Case ".mp3"
                    Return "audio/mpeg"
                Case ".m4a", ".mp4"
                    Return "audio/mp4"
                Case ".flac"
                    Return "audio/flac"
                Case ".ogg", ".oga", ".opus"
                    Return "audio/ogg"
                Case ".aac"
                    Return "audio/aac"
                Case Else
                    Return "application/octet-stream"
            End Select
        End Function

        Private Shared Sub EnsureTls12()
            Try
                System.AppContext.SetSwitch("Switch.System.Net.DontEnableSchUseStrongCrypto", False)
                System.AppContext.SetSwitch("Switch.System.Net.DontEnableSystemDefaultTlsVersions", False)
                System.Net.ServicePointManager.Expect100Continue = False
                System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12
            Catch
            End Try
        End Sub

        Private Sub RaiseStatusMessage(message As System.String)
            RaiseEvent Status(Me, New TranscriptionStatusEventArgs(message))
        End Sub

        Public Function DisposeAsync() As System.Threading.Tasks.ValueTask Implements System.IAsyncDisposable.DisposeAsync
            _http.Dispose()
            Return New System.Threading.Tasks.ValueTask(System.Threading.Tasks.Task.CompletedTask)
        End Function

    End Class

End Namespace
