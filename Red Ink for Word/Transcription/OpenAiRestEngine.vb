Option Explicit On
Option Strict Off

Imports System.IO
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Threading
Imports NAudio.Wave
Imports Newtonsoft.Json.Linq
Imports System.Text
Imports System.Net
Imports System.Security.Authentication

Namespace Transcription

    Public Class OpenAiRestEngine
        Implements ITranscriptionEngine

        Public Const DisplayNameWhisper1 As String = "OpenAI whisper-1 (REST)"
        Public Const DisplayNameGpt4o As String = "OpenAI gpt-4o-transcribe (REST)"
        Public Const DisplayNameGpt4oMini As String = "OpenAI gpt-4o-mini-transcribe (REST)"
        Public Const DisplayNameGpt4oDiarize As String = "OpenAI gpt-4o-transcribe-diarize (REST diarization)"

        Private Const DefaultOpenAiRestUrl As String = "https://api.openai.com/v1/audio/transcriptions"
        Private Const DefaultRestModel As String = "gpt-4o-mini-transcribe"
        Private Const DiarizeModel As String = "gpt-4o-transcribe-diarize"

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
                Return "OpenAI REST"
            End Get
        End Property

        Public ReadOnly Property Kind As EngineKind Implements ITranscriptionEngine.Kind
            Get
                Return EngineKind.OpenAiRest
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

        Private ReadOnly _endpoint As String
        Private ReadOnly _apiKey As String
        Private ReadOnly _http As System.Net.Http.HttpClient
        Private _opts As TranscriptionOptions


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

        Public Sub New(apiKey As String)
            _endpoint = DefaultOpenAiRestUrl
            _apiKey = apiKey

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

        Private Sub RaiseStatusMessage(message As String)
            RaiseEvent Status(Me, New TranscriptionStatusEventArgs(message))
        End Sub

        Public Function StartLiveAsync(opts As TranscriptionOptions, ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task Implements ITranscriptionEngine.StartLiveAsync
            Throw New System.NotSupportedException("OpenAI REST is file/request-response transcription only. Use OpenAI Realtime for live microphone transcription.")
        End Function

        Public Function PushAudioAsync(pcm As Byte(), bytesValid As Integer, ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task Implements ITranscriptionEngine.PushAudioAsync
            Return System.Threading.Tasks.Task.CompletedTask
        End Function

        Public Function StopLiveAsync() As System.Threading.Tasks.Task Implements ITranscriptionEngine.StopLiveAsync
            Return System.Threading.Tasks.Task.CompletedTask
        End Function

        Public Async Function TranscribeFileAsync(filePath As String, opts As TranscriptionOptions, ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task Implements ITranscriptionEngine.TranscribeFileAsync
            _opts = opts

            RaiseStatusMessage("Preparing file for OpenAI REST…")

            Using payload As System.IO.Stream = CreateUploadStream(filePath)
                Await PostMultipartAsync(payload, System.IO.Path.GetFileNameWithoutExtension(filePath) & ".wav", opts, ct)
            End Using
        End Function

        Private Shared Function CreateUploadStream(filePath As String) As System.IO.Stream
            Dim pcm As Byte() = VoskEngine.LoadAudioToPcm16Mono16k(filePath)
            Dim wavStream As New System.IO.MemoryStream()

            Using raw As New RawSourceWaveStream(New System.IO.MemoryStream(pcm, False), New WaveFormat(16000, 16, 1))
                WaveFileWriter.WriteWavFileToStream(wavStream, raw)
            End Using

            wavStream.Position = 0
            Return wavStream
        End Function

        Private Shared Function GetDetailedExceptionMessage(ex As System.Exception) As String
            If ex Is Nothing Then
                Return ""
            End If

            Dim sb As New StringBuilder()
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

        Private Async Function PostMultipartAsync(payload As System.IO.Stream, fileName As String, opts As TranscriptionOptions, ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task
            Dim model As String = NormalizeModel(opts)
            ValidateOptions(model, opts)

            EnsureTls12()

            Using req As New System.Net.Http.HttpRequestMessage(System.Net.Http.HttpMethod.Post, _endpoint)
                req.Headers.Authorization = New System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", _apiKey)
                req.Headers.ExpectContinue = False
                req.Version = New System.Version(1, 1)

                Using form As New System.Net.Http.MultipartFormDataContent()
                    Dim sc As New System.Net.Http.StreamContent(payload)
                    sc.Headers.ContentType = New System.Net.Http.Headers.MediaTypeHeaderValue("audio/wav")
                    form.Add(sc, "file", fileName)
                    form.Add(New System.Net.Http.StringContent(model), "model")

                    If IsDiarizeModel(model) Then
                        form.Add(New System.Net.Http.StringContent("diarized_json"), "response_format")
                        form.Add(New System.Net.Http.StringContent("auto"), "chunking_strategy")
                    Else
                        form.Add(New System.Net.Http.StringContent("json"), "response_format")
                    End If

                    If Not System.String.IsNullOrWhiteSpace(opts.LanguageCode) AndAlso Not opts.LanguageCode.Equals("auto", System.StringComparison.OrdinalIgnoreCase) Then
                        form.Add(New System.Net.Http.StringContent(opts.LanguageCode), "language")
                    End If

                    If Not IsDiarizeModel(model) AndAlso Not System.String.IsNullOrWhiteSpace(opts.Prompt) Then
                        form.Add(New System.Net.Http.StringContent(opts.Prompt), "prompt")
                    End If

                    req.Content = form

                    RaiseStatusMessage("Uploading audio to OpenAI REST…")

                    Try
                        Using resp As System.Net.Http.HttpResponseMessage = Await _http.SendAsync(req, System.Net.Http.HttpCompletionOption.ResponseContentRead, ct)
                            Dim body As String = Await resp.Content.ReadAsStringAsync()

                            If Not resp.IsSuccessStatusCode Then
                                Dim detail As String = "OpenAI REST HTTP " &
                            CInt(resp.StatusCode).ToString(System.Globalization.CultureInfo.InvariantCulture) &
                            " " & resp.StatusCode.ToString() &
                            ": " & body

                                RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs(detail, Nothing, False))
                                Throw New System.InvalidOperationException(detail)
                            End If

                            RaiseStatusMessage("Parsing OpenAI REST response…")
                            HandleResponseBody(model, body)
                            RaiseStatusMessage("OpenAI REST file transcription completed.")
                        End Using
                    Catch ex As System.OperationCanceledException When ct.IsCancellationRequested
                        RaiseStatusMessage("OpenAI REST canceled.")
                        Throw
                    Catch ex As System.Net.Http.HttpRequestException
                        Dim detail As String = "OpenAI REST request failed: " & GetDetailedExceptionMessage(ex)
                        RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs(detail, ex, False))
                        Throw New System.InvalidOperationException(detail, ex)
                    End Try
                End Using
            End Using
        End Function

        Private Shared Function NormalizeModel(opts As TranscriptionOptions) As String
            If opts Is Nothing OrElse System.String.IsNullOrWhiteSpace(opts.Model) Then
                Return DefaultRestModel
            End If

            Return opts.Model.Trim()
        End Function

        Private Shared Function IsDiarizeModel(model As String) As Boolean
            Return System.String.Equals(model, DiarizeModel, System.StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Sub ValidateOptions(model As String, opts As TranscriptionOptions)
            If IsDiarizeModel(model) AndAlso opts IsNot Nothing AndAlso Not System.String.IsNullOrWhiteSpace(opts.Prompt) Then
                Throw New System.InvalidOperationException("The OpenAI gpt-4o-transcribe-diarize model does not support the prompt parameter.")
            End If
        End Sub

        Private Sub HandleResponseBody(model As String, body As String)
            Try
                Dim jo As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(body)

                If IsDiarizeModel(model) Then
                    Dim diarizedText As String = FormatDiarizedSegments(jo)
                    If Not System.String.IsNullOrWhiteSpace(diarizedText) Then
                        RaiseEvent FinalResult(Me, New TranscriptionEventArgs(diarizedText.Trim(), True))
                        Return
                    End If
                End If

                Dim textValue As String = If(jo("text")?.ToString(), "")
                If Not System.String.IsNullOrWhiteSpace(textValue) Then
                    RaiseEvent FinalResult(Me, New TranscriptionEventArgs(textValue.Trim(), True))
                    Return
                End If

                Throw New System.InvalidOperationException("OpenAI REST returned no transcription text.")
            Catch ex As Newtonsoft.Json.JsonException
                If Not System.String.IsNullOrWhiteSpace(body) Then
                    RaiseEvent FinalResult(Me, New TranscriptionEventArgs(body.Trim(), True))
                    Return
                End If

                Throw
            End Try
        End Sub

        Private Shared Function FormatDiarizedSegments(jo As Newtonsoft.Json.Linq.JObject) As String
            Dim segments As Newtonsoft.Json.Linq.JArray = TryCast(jo("segments"), Newtonsoft.Json.Linq.JArray)

            If segments Is Nothing OrElse segments.Count = 0 Then
                Return ""
            End If

            Dim sb As New System.Text.StringBuilder()

            For Each seg As Newtonsoft.Json.Linq.JToken In segments
                Dim speaker As String = If(seg("speaker")?.ToString(), "?").Trim()
                Dim textValue As String = If(seg("text")?.ToString(), "").Trim()

                If textValue.Length > 0 Then
                    sb.AppendLine("[" & speaker & "] " & textValue)
                End If
            Next

            Return sb.ToString().Trim()
        End Function

        Private Shared Function GetContentTypeForFileName(fileName As String) As String
            Dim ext As String = System.IO.Path.GetExtension(If(fileName, "")).ToLowerInvariant()

            Select Case ext
                Case ".wav"
                    Return "audio/wav"
                Case ".mp3", ".mpeg", ".mpga"
                    Return "audio/mpeg"
                Case ".m4a", ".mp4"
                    Return "audio/mp4"
                Case ".ogg"
                    Return "audio/ogg"
                Case ".flac"
                    Return "audio/flac"
                Case ".webm"
                    Return "audio/webm"
                Case Else
                    Return "application/octet-stream"
            End Select
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
