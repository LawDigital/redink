Option Explicit On
Option Strict Off

Imports System.IO
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Threading
Imports NAudio.Wave
Imports Newtonsoft.Json.Linq

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
        Private ReadOnly _http As New System.Net.Http.HttpClient() With {.Timeout = System.TimeSpan.FromMinutes(10)}
        Private _opts As TranscriptionOptions

        Public Sub New(apiKey As String)
            _endpoint = DefaultOpenAiRestUrl
            _apiKey = apiKey
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

            Using fs As New System.IO.FileStream(filePath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read)
                Await PostMultipartAsync(fs, System.IO.Path.GetFileName(filePath), opts, ct)
            End Using
        End Function

        Private Async Function PostMultipartAsync(payload As System.IO.Stream, fileName As String, opts As TranscriptionOptions, ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task
            Dim model As String = NormalizeModel(opts)
            ValidateOptions(model, opts)

            Using req As New System.Net.Http.HttpRequestMessage(System.Net.Http.HttpMethod.Post, _endpoint)
                req.Headers.Authorization = New System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", _apiKey)

                Using form As New System.Net.Http.MultipartFormDataContent()
                    Dim sc As New System.Net.Http.StreamContent(payload)
                    sc.Headers.ContentType = New System.Net.Http.Headers.MediaTypeHeaderValue(GetContentTypeForFileName(fileName))
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

                    Using resp As System.Net.Http.HttpResponseMessage = Await _http.SendAsync(req, ct)
                        Dim body As String = Await resp.Content.ReadAsStringAsync()

                        If Not resp.IsSuccessStatusCode Then
                            RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("OpenAI REST HTTP " & CInt(resp.StatusCode).ToString(System.Globalization.CultureInfo.InvariantCulture) & " " & resp.StatusCode.ToString() & ": " & body, Nothing, False))
                            Return
                        End If

                        HandleResponseBody(model, body)
                    End Using
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
                End If
            Catch ex As Newtonsoft.Json.JsonException
                If Not System.String.IsNullOrWhiteSpace(body) Then
                    RaiseEvent FinalResult(Me, New TranscriptionEventArgs(body.Trim(), True))
                End If
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
