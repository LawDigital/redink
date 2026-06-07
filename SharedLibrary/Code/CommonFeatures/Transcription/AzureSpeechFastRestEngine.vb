' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: AzureSpeechFastRestEngine.vb
' Purpose: Implements a transcription engine using Azure's Speech-to-Text
'          REST API for fast, short audio transcriptions.
'
' Architecture:
'  - API Interaction: Communicates with the Azure Speech-to-Text REST API
'    endpoint.
'  - Authentication: Manages API keys and authentication with Azure services.
'  - Data Formatting: Prepares audio data to be sent in the format required
'    by the REST API.
'  - Response Handling: Parses the JSON response from the API to extract the
'    transcribed text.
' =============================================================================


Option Explicit On
Option Strict Off

Imports System.IO
Imports System.Linq
Imports System.Net
Imports System.Net.Http
Imports System.Text
Imports NAudio.Wave
Imports Newtonsoft.Json.Linq

Namespace Transcription

    ''' <summary>
    ''' Azure Speech-to-Text file transcription engine using the Fast Transcription REST API.
    '''
    ''' API target:
    '''     POST {endpoint}/speechtotext/transcriptions:transcribe?api-version=2025-10-15
    '''
    ''' Required integration outside this file:
    '''     1. Add EngineKind.AzureSpeechFastRest to EngineKind.
    '''     2. Add AzureSpeechFastRestEngine.DisplayNameValue in LoadEngines().
    '''     3. In CreateEngineAsync(), return New AzureSpeechFastRestEngine(key, endpointOrRegion).
    '''
    ''' endpointOrRegion may be:
    '''     - "westeurope"
    '''     - "https://westeurope.api.cognitive.microsoft.com"
    '''     - "https://your-speech-resource.cognitiveservices.azure.com"
    '''
    ''' This engine mirrors OpenAiRestEngine: it converts the selected audio file to PCM16 mono
    ''' 16 kHz WAV in memory and uploads it as multipart/form-data.
    ''' </summary>
    Public Class AzureSpeechFastRestEngine
        Implements ITranscriptionEngine

        Public Const DisplayNameValue As String = "Azure Speech-to-Text Fast REST (files)"

        Private Const DefaultLanguage As String = "de-DE"
        Private Const DefaultApiVersion As String = "2025-10-15"
        Private Const DefaultEndpointOrRegion As String = "westeurope"

        Private ReadOnly _subscriptionKey As String
        Private ReadOnly _endpointOrRegion As String
        Private ReadOnly _apiVersion As String
        Private ReadOnly _region As String
        Private ReadOnly _http As System.Net.Http.HttpClient

        Public Sub New(subscriptionKey As String, endpointOrRegion As String, Optional apiVersion As String = "", Optional region As String = "")
            _subscriptionKey = If(subscriptionKey, "").Trim()
            _endpointOrRegion = If(endpointOrRegion, "").Trim()
            _apiVersion = If(apiVersion, "").Trim()
            _region = If(region, "").Trim()

            EnsureTls12()

            Dim handler As New System.Net.Http.HttpClientHandler() With {
                .AllowAutoRedirect = True,
                .AutomaticDecompression = System.Net.DecompressionMethods.GZip Or System.Net.DecompressionMethods.Deflate
            }

            Try
                handler.SslProtocols = System.Security.Authentication.SslProtocols.Tls12
            Catch
            End Try

            _http = New System.Net.Http.HttpClient(handler) With {
                .Timeout = System.TimeSpan.FromMinutes(30)
            }
        End Sub

        Private Function GetEffectiveEndpointOrRegion() As String
            If Not System.String.IsNullOrWhiteSpace(_endpointOrRegion) Then
                Return _endpointOrRegion
            End If

            Return DefaultEndpointOrRegion
        End Function

        Private Function GetEffectiveApiVersion() As String
            If Not System.String.IsNullOrWhiteSpace(_apiVersion) Then
                Return _apiVersion
            End If

            Return DefaultApiVersion
        End Function

        Private Shared ReadOnly _supportedLanguages As String() = {
            "auto",
            "ar-SA", "bg-BG", "ca-ES", "cs-CZ", "da-DK", "de-DE", "el-GR", "en-US", "en-GB",
            "es-ES", "es-MX", "et-EE", "fi-FI", "fr-FR", "he-IL", "hi-IN", "hr-HR", "hu-HU",
            "id-ID", "it-IT", "ja-JP", "ko-KR", "lt-LT", "lv-LV", "ms-MY", "nl-NL", "nb-NO",
            "pl-PL", "pt-PT", "pt-BR", "ro-RO", "ru-RU", "sk-SK", "sl-SI", "sv-SE", "th-TH",
            "tr-TR", "uk-UA", "vi-VN", "zh-CN", "zh-HK", "zh-TW"
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
                Return "Azure Speech Fast REST"
            End Get
        End Property

        Public ReadOnly Property Kind As EngineKind Implements ITranscriptionEngine.Kind
            Get
                Return CType(System.Enum.Parse(GetType(EngineKind), "AzureSpeechFastRest"), EngineKind)
            End Get
        End Property

        Public ReadOnly Property SupportsLiveStreaming As Boolean Implements ITranscriptionEngine.SupportsLiveStreaming
            Get
                Return False
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

        Private Sub RaiseStatusMessage(message As String)
            RaiseEvent Status(Me, New TranscriptionStatusEventArgs(message))
        End Sub

        Public Function StartLiveAsync(opts As TranscriptionOptions, ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task Implements ITranscriptionEngine.StartLiveAsync
            Throw New System.NotSupportedException("Azure Speech Fast REST is file/request-response transcription only. Use AzureSpeechSdkRealtimeEngine for live microphone transcription.")
        End Function

        Public Function PushAudioAsync(pcm As Byte(), bytesValid As Integer, ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task Implements ITranscriptionEngine.PushAudioAsync
            Return System.Threading.Tasks.Task.CompletedTask
        End Function

        Public Function StopLiveAsync() As System.Threading.Tasks.Task Implements ITranscriptionEngine.StopLiveAsync
            Return System.Threading.Tasks.Task.CompletedTask
        End Function

        Public Async Function TranscribeFileAsync(filePath As String, opts As TranscriptionOptions, ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task Implements ITranscriptionEngine.TranscribeFileAsync
            If System.String.IsNullOrWhiteSpace(_subscriptionKey) Then
                Throw New System.InvalidOperationException("Azure Speech subscription key is missing.")
            End If

            If System.String.IsNullOrWhiteSpace(filePath) OrElse Not System.IO.File.Exists(filePath) Then
                Throw New System.IO.FileNotFoundException("Audio file not found.", filePath)
            End If

            RaiseStatusMessage("Preparing file for Azure Speech Fast REST…")

            Using payload As System.IO.Stream = CreateUploadStream(filePath)
                Await PostMultipartAsync(payload, System.IO.Path.GetFileNameWithoutExtension(filePath) & ".wav", opts, ct).ConfigureAwait(False)
            End Using
        End Function


        Private Function GetEffectiveRequestEndpoint() As String
            If Not System.String.IsNullOrWhiteSpace(_region) Then
                Return NormalizeEndpoint(_region)
            End If

            Return NormalizeEndpoint(GetEffectiveEndpointOrRegion())
        End Function


        Private Async Function PostMultipartAsync(payload As System.IO.Stream, fileName As String, opts As TranscriptionOptions, ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task
            Dim endpoint As String = GetEffectiveRequestEndpoint()
            Dim requestUri As String = endpoint.TrimEnd("/"c) & "/speechtotext/transcriptions:transcribe?api-version=" & System.Uri.EscapeDataString(GetEffectiveApiVersion())
            Dim definitionJson As String = BuildDefinitionJson(opts)

            EnsureTls12()

            Using req As New System.Net.Http.HttpRequestMessage(System.Net.Http.HttpMethod.Post, requestUri)
                req.Headers.TryAddWithoutValidation("Ocp-Apim-Subscription-Key", _subscriptionKey)
                req.Headers.TryAddWithoutValidation("api-key", _subscriptionKey)

                If Not System.String.IsNullOrWhiteSpace(_region) Then
                    req.Headers.TryAddWithoutValidation("Ocp-Apim-Subscription-Region", _region)
                End If

                req.Headers.ExpectContinue = False
                req.Version = New System.Version(1, 1)

                Using form As New System.Net.Http.MultipartFormDataContent()
                    Dim audioContent As New System.Net.Http.StreamContent(payload)
                    audioContent.Headers.ContentType = New System.Net.Http.Headers.MediaTypeHeaderValue("audio/wav")
                    form.Add(audioContent, "audio", fileName)
                    form.Add(New System.Net.Http.StringContent(definitionJson, System.Text.Encoding.UTF8, "application/json"), "definition")

                    req.Content = form

                    RaiseStatusMessage("Uploading audio to Azure Speech Fast REST…")

                    Try
                        Using resp As System.Net.Http.HttpResponseMessage = Await _http.SendAsync(req, System.Net.Http.HttpCompletionOption.ResponseContentRead, ct).ConfigureAwait(False)
                            Dim body As String = Await resp.Content.ReadAsStringAsync().ConfigureAwait(False)

                            If Not resp.IsSuccessStatusCode Then
                                Dim detail As String = "Azure Speech Fast REST HTTP " &
                                    CInt(resp.StatusCode).ToString(System.Globalization.CultureInfo.InvariantCulture) &
                                    " " & resp.StatusCode.ToString() & ": " & body

                                RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs(detail, Nothing, False))
                                Throw New System.InvalidOperationException(detail)
                            End If

                            RaiseStatusMessage("Parsing Azure Speech Fast REST response…")
                            HandleResponseBody(body, opts)
                            RaiseStatusMessage("Azure Speech Fast REST file transcription completed.")
                        End Using
                    Catch ex As System.OperationCanceledException When ct.IsCancellationRequested
                        RaiseStatusMessage("Azure Speech Fast REST canceled.")
                        Throw
                    Catch ex As System.Net.Http.HttpRequestException
                        Dim detail As String = "Azure Speech Fast REST request failed: " & GetDetailedExceptionMessage(ex)
                        RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs(detail, ex, False))
                        Throw New System.InvalidOperationException(detail, ex)
                    End Try
                End Using
            End Using
        End Function

        Private Sub HandleResponseBody(body As String, opts As TranscriptionOptions)
            Try
                Dim jo As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(body)
                Dim phrases As Newtonsoft.Json.Linq.JArray = TryCast(jo("phrases"), Newtonsoft.Json.Linq.JArray)
                Dim combinedPhrases As Newtonsoft.Json.Linq.JArray = TryCast(jo("combinedPhrases"), Newtonsoft.Json.Linq.JArray)
                Dim diarizationRequested As Boolean = WantsDiarization(opts)

                If diarizationRequested Then
                    Dim diarizedText As String = BuildPhraseText(phrases)
                    If diarizedText.Length > 0 Then
                        RaiseEvent FinalResult(Me, New TranscriptionEventArgs(diarizedText, True))
                        Return
                    End If
                End If

                Dim combinedText As String = BuildCombinedPhraseText(combinedPhrases)
                If combinedText.Length > 0 Then
                    RaiseEvent FinalResult(Me, New TranscriptionEventArgs(combinedText, True))
                    Return
                End If

                Dim phraseText As String = BuildPhraseText(phrases)
                If phraseText.Length > 0 Then
                    RaiseEvent FinalResult(Me, New TranscriptionEventArgs(phraseText, True))
                    Return
                End If

                Throw New System.InvalidOperationException("Azure Speech Fast REST returned no transcription text.")
            Catch ex As Newtonsoft.Json.JsonException
                If Not System.String.IsNullOrWhiteSpace(body) Then
                    RaiseEvent FinalResult(Me, New TranscriptionEventArgs(body.Trim(), True))
                    Return
                End If

                Throw
            End Try
        End Sub

        Private Shared Function BuildCombinedPhraseText(combinedPhrases As Newtonsoft.Json.Linq.JArray) As String
            If combinedPhrases Is Nothing OrElse combinedPhrases.Count = 0 Then
                Return ""
            End If

            Dim sb As New System.Text.StringBuilder()

            For Each phrase As Newtonsoft.Json.Linq.JToken In combinedPhrases
                Dim textValue As String = If(phrase("text")?.ToString(), "").Trim()
                Dim channelValue As String = If(phrase("channel")?.ToString(), "").Trim()

                If textValue.Length > 0 Then
                    If channelValue.Length > 0 Then
                        sb.AppendLine("[Channel " & channelValue & "] " & textValue)
                    Else
                        sb.AppendLine(textValue)
                    End If
                End If
            Next

            Return sb.ToString().Trim()
        End Function

        Private Shared Function BuildPhraseText(phrases As Newtonsoft.Json.Linq.JArray) As String
            If phrases Is Nothing OrElse phrases.Count = 0 Then
                Return ""
            End If

            Dim sb As New System.Text.StringBuilder()

            For Each phrase As Newtonsoft.Json.Linq.JToken In phrases
                Dim textValue As String = If(phrase("text")?.ToString(), "").Trim()
                Dim speakerValue As String = If(phrase("speaker")?.ToString(), "").Trim()

                If textValue.Length > 0 Then
                    If speakerValue.Length > 0 Then
                        sb.AppendLine("[Speaker " & speakerValue & "] " & textValue)
                    Else
                        sb.AppendLine(textValue)
                    End If
                End If
            Next

            Return sb.ToString().Trim()
        End Function





        Private Shared Function GetDiarizationMaxSpeakers(opts As TranscriptionOptions) As Integer
            Dim value As Integer = 2

            If opts IsNot Nothing Then
                Try
                    Dim prop As System.Reflection.PropertyInfo = opts.GetType().GetProperty("DiarizationMaxSpeakers")
                    If prop IsNot Nothing Then
                        Dim raw As Object = prop.GetValue(opts, Nothing)
                        If raw IsNot Nothing Then
                            value = System.Convert.ToInt32(raw, System.Globalization.CultureInfo.InvariantCulture)
                        End If
                    End If
                Catch ex As System.Exception
                End Try
            End If

            If value < 2 Then
                value = 2
            End If

            If value > 35 Then
                value = 35
            End If

            Return value
        End Function
        Private Sub HandleResponseBody(body As String)
            Try
                Dim jo As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(body)
                Dim combinedPhrases As Newtonsoft.Json.Linq.JArray = TryCast(jo("combinedPhrases"), Newtonsoft.Json.Linq.JArray)

                If combinedPhrases IsNot Nothing AndAlso combinedPhrases.Count > 0 Then
                    Dim sb As New System.Text.StringBuilder()

                    For Each phrase As Newtonsoft.Json.Linq.JToken In combinedPhrases
                        Dim textValue As String = If(phrase("text")?.ToString(), "").Trim()
                        Dim channelValue As String = If(phrase("channel")?.ToString(), "").Trim()

                        If textValue.Length > 0 Then
                            If channelValue.Length > 0 Then
                                sb.AppendLine("[Channel " & channelValue & "] " & textValue)
                            Else
                                sb.AppendLine(textValue)
                            End If
                        End If
                    Next

                    Dim combinedText As String = sb.ToString().Trim()
                    If combinedText.Length > 0 Then
                        RaiseEvent FinalResult(Me, New TranscriptionEventArgs(combinedText, True))
                        Return
                    End If
                End If

                Dim phrases As Newtonsoft.Json.Linq.JArray = TryCast(jo("phrases"), Newtonsoft.Json.Linq.JArray)
                If phrases IsNot Nothing AndAlso phrases.Count > 0 Then
                    Dim sb As New System.Text.StringBuilder()

                    For Each phrase As Newtonsoft.Json.Linq.JToken In phrases
                        Dim textValue As String = If(phrase("text")?.ToString(), "").Trim()
                        Dim speakerValue As String = If(phrase("speaker")?.ToString(), "").Trim()

                        If textValue.Length > 0 Then
                            If speakerValue.Length > 0 Then
                                sb.AppendLine("[Speaker " & speakerValue & "] " & textValue)
                            Else
                                sb.AppendLine(textValue)
                            End If
                        End If
                    Next

                    Dim phraseText As String = sb.ToString().Trim()
                    If phraseText.Length > 0 Then
                        RaiseEvent FinalResult(Me, New TranscriptionEventArgs(phraseText, True))
                        Return
                    End If
                End If

                Throw New System.InvalidOperationException("Azure Speech Fast REST returned no transcription text.")
            Catch ex As Newtonsoft.Json.JsonException
                If Not System.String.IsNullOrWhiteSpace(body) Then
                    RaiseEvent FinalResult(Me, New TranscriptionEventArgs(body.Trim(), True))
                    Return
                End If

                Throw
            End Try
        End Sub

        Private Shared Function BuildDefinitionJson(opts As TranscriptionOptions) As String
            Dim languageCode As String = NormalizeLanguageCode(opts)

            Dim definition As New Newtonsoft.Json.Linq.JObject()
            definition("locales") = New Newtonsoft.Json.Linq.JArray(languageCode)

            If WantsDiarization(opts) Then
                Dim maxSpeakers As Integer = GetDiarizationMaxSpeakers(opts)

                definition("diarization") = New Newtonsoft.Json.Linq.JObject From {
                    {"enabled", True},
                    {"maxSpeakers", maxSpeakers}
                }
            ElseIf opts IsNot Nothing AndAlso opts.MultiChannelDiarization Then
                definition("channels") = New Newtonsoft.Json.Linq.JArray(0, 1)
            End If

            Return definition.ToString(Newtonsoft.Json.Formatting.None)
        End Function

        Private Shared Function WantsDiarization(opts As TranscriptionOptions) As Boolean
            If opts Is Nothing Then
                Return False
            End If

            Try
                Dim modelText As String = If(opts.Model, "").Trim()
                If modelText.IndexOf("diar", System.StringComparison.OrdinalIgnoreCase) >= 0 Then
                    Return True
                End If
            Catch
            End Try

            Try
                Dim prop As System.Reflection.PropertyInfo = opts.GetType().GetProperty("Diarization")
                If prop IsNot Nothing Then
                    Dim value As Object = prop.GetValue(opts, Nothing)
                    If value IsNot Nothing AndAlso System.Convert.ToBoolean(value, System.Globalization.CultureInfo.InvariantCulture) Then
                        Return True
                    End If
                End If
            Catch
            End Try

            Try
                Dim prop As System.Reflection.PropertyInfo = opts.GetType().GetProperty("EnableDiarization")
                If prop IsNot Nothing Then
                    Dim value As Object = prop.GetValue(opts, Nothing)
                    If value IsNot Nothing AndAlso System.Convert.ToBoolean(value, System.Globalization.CultureInfo.InvariantCulture) Then
                        Return True
                    End If
                End If
            Catch
            End Try

            Return False
        End Function

        Private Shared Function CreateUploadStream(filePath As String) As System.IO.Stream
            Dim pcm As Byte() = VoskEngine.LoadAudioToPcm16Mono16k(filePath)
            Dim wavStream As New System.IO.MemoryStream()

            Using raw As New NAudio.Wave.RawSourceWaveStream(New System.IO.MemoryStream(pcm, False), New NAudio.Wave.WaveFormat(16000, 16, 1))
                NAudio.Wave.WaveFileWriter.WriteWavFileToStream(wavStream, raw)
            End Using

            wavStream.Position = 0
            Return wavStream
        End Function

        Private Shared Function NormalizeEndpoint(endpointOrRegion As String) As String
            Dim raw As String = If(endpointOrRegion, "").Trim()

            If raw.StartsWith("https://", System.StringComparison.OrdinalIgnoreCase) OrElse raw.StartsWith("http://", System.StringComparison.OrdinalIgnoreCase) Then
                Return raw.TrimEnd("/"c)
            End If

            Return "https://" & raw & ".api.cognitive.microsoft.com"
        End Function

        Private Shared Function NormalizeLanguageCode(opts As TranscriptionOptions) As String
            Dim raw As String = ""

            If opts IsNot Nothing Then
                raw = If(opts.LanguageCode, "").Trim()
            End If

            If System.String.IsNullOrWhiteSpace(raw) OrElse System.String.Equals(raw, "auto", System.StringComparison.OrdinalIgnoreCase) Then
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

        Public Function DisposeAsync() As System.Threading.Tasks.ValueTask Implements System.IAsyncDisposable.DisposeAsync
            Try
                _http.Dispose()
            Catch ex As System.Exception
            End Try

            Return New System.Threading.Tasks.ValueTask()
        End Function
    End Class

End Namespace
