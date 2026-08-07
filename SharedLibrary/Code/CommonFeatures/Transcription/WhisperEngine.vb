' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: WhisperEngine.vb
' Purpose: Implements the ITranscriptionEngine interface using a local
'          instance of OpenAI's Whisper model. This provides high-quality,
'          on-device transcription.
'
' Architecture:
'  - ITranscriptionEngine Implementation: Fulfills the transcription contract
'    for use with the Whisper model.
'  - Local Whisper Integration: Manages the interaction with a local Whisper
'    implementation (e.g., whisper.cpp), including model loading and execution.
'  - Audio Processing: Prepares audio data in the format required by the
'    Whisper model (e.g., 16kHz, 16-bit mono PCM).
'  - Resource Management: Handles the potentially significant memory and CPU/GPU
'    resources required to run the Whisper model locally.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports NAudio.Wave
Imports Whisper.net
Imports Whisper.net.LibraryLoader

Namespace Transcription

    Public Class WhisperEngine
        Implements ITranscriptionEngine

        Private Shared ReadOnly _supportedLanguages As String() = {
            "auto", "af", "am", "ar", "as", "az", "ba", "be", "bg", "bn", "bo", "br", "bs", "ca", "cs", "cy",
            "da", "de", "el", "en", "es", "et", "eu", "fa", "fi", "fo", "fr", "gl", "gu", "ha", "haw", "he",
            "hi", "hr", "ht", "hu", "hy", "id", "is", "it", "ja", "jw", "ka", "kk", "km", "kn", "ko", "la",
            "lb", "ln", "lo", "lt", "lv", "mg", "mi", "mk", "ml", "mn", "mr", "ms", "mt", "my", "ne", "nl",
            "nn", "no", "oc", "pa", "pl", "ps", "pt", "ro", "ru", "sa", "sd", "si", "sk", "sl", "sn", "so",
            "sq", "sr", "su", "sv", "sw", "ta", "te", "tg", "th", "tk", "tl", "tr", "tt", "uk", "ur", "uz",
            "vi", "yi", "yo", "zh", "zu"
        }

        Public Shared ReadOnly Property SupportedLanguages As String()
            Get
                Return _supportedLanguages.
                    OrderBy(Function(x) If(String.Equals(x, "auto", StringComparison.OrdinalIgnoreCase), "", x), StringComparer.OrdinalIgnoreCase).
                    ToArray()
            End Get
        End Property


        Public Event PartialResult As EventHandler(Of TranscriptionEventArgs) Implements ITranscriptionEngine.PartialResult
        Public Event FinalResult As EventHandler(Of TranscriptionEventArgs) Implements ITranscriptionEngine.FinalResult
        Public Event EngineError As EventHandler(Of TranscriptionErrorEventArgs) Implements ITranscriptionEngine.EngineError
        Public Event Status As EventHandler(Of TranscriptionStatusEventArgs) Implements ITranscriptionEngine.Status

        Public ReadOnly Property Name As String Implements ITranscriptionEngine.Name
            Get
                Return _modelFile
            End Get
        End Property

        Public ReadOnly Property Kind As EngineKind Implements ITranscriptionEngine.Kind
            Get
                Return EngineKind.WhisperLocal
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
                Return False
            End Get
        End Property

        Private ReadOnly _modelRoot As String
        Private ReadOnly _modelFile As String
        Private _factory As WhisperFactory
        Private _processor As WhisperProcessor
        Private _audioBuffer As New List(Of Single)()
        Private Const ROLLING_OVERLAP_SAMPLES As Integer = 16000 \ 2
        Private Const PROCESS_THRESHOLD_SAMPLES As Integer = 16000 * 2
        Private _cancelled As Boolean

        Private Shared ReadOnly _runtimeInitLock As New Object()
        Private Shared _runtimeConfigured As Boolean

        Public Sub New(modelRoot As String, modelFileName As String)
            _modelRoot = modelRoot
            _modelFile = modelFileName
        End Sub

        Private Sub RaiseStatusMessage(message As String, Optional progressPercent As System.Nullable(Of Integer) = Nothing)
            RaiseEvent Status(Me, New TranscriptionStatusEventArgs(message, progressPercent))
        End Sub

        Private Sub Init(opts As TranscriptionOptions)
            EnsureRuntimeConfigured()

            Dim modelPath As String = Path.Combine(_modelRoot, _modelFile)
            _factory = WhisperFactory.FromPath(modelPath)

            Dim vad As Single = 0.6F
            If opts.VadThreshold > 0.0F AndAlso opts.VadThreshold < 1.0F Then
                vad = opts.VadThreshold
            End If

            Dim lang As String = If(String.IsNullOrWhiteSpace(opts.LanguageCode), "auto", opts.LanguageCode)

            Dim builder = _factory.CreateBuilder() _
                .WithLanguage(lang) _
                .WithThreads(Environment.ProcessorCount) _
                .WithNoSpeechThreshold(vad) _
                .WithTemperature(0.3)

            If opts.Translate Then
                builder = builder.WithTranslate()
            End If

            _processor = builder.Build()
            _cancelled = False
        End Sub

        Private Sub EnsureRuntimeConfigured()
            If _runtimeConfigured Then
                Return
            End If

            SyncLock _runtimeInitLock
                If _runtimeConfigured Then
                    Return
                End If

                Dim speechPath As String = _modelRoot

                If String.IsNullOrWhiteSpace(speechPath) Then
                    _runtimeConfigured = True
                    Return
                End If

                speechPath = Path.GetFullPath(speechPath.Trim())

                If Not speechPath.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal) Then
                    speechPath &= Path.DirectorySeparatorChar
                End If

                Dim currentPath As String = Environment.GetEnvironmentVariable("PATH")

                If String.IsNullOrEmpty(currentPath) Then
                    Environment.SetEnvironmentVariable("PATH", speechPath)
                ElseIf Not PathContainsDirectory(currentPath, speechPath) Then
                    Environment.SetEnvironmentVariable("PATH", currentPath & Path.PathSeparator & speechPath)
                End If

                RuntimeOptions.LibraryPath = speechPath
                'RuntimeOptions.RuntimeLibraryOrder = New List(Of RuntimeLibrary) From {RuntimeLibrary.Cuda, RuntimeLibrary.Cpu}

                _runtimeConfigured = True
            End SyncLock
        End Sub

        Private Shared Function PathContainsDirectory(pathValue As String, directoryPath As String) As Boolean
            If String.IsNullOrWhiteSpace(pathValue) OrElse String.IsNullOrWhiteSpace(directoryPath) Then
                Return False
            End If

            For Each part As String In pathValue.Split(Path.PathSeparator)
                Dim candidate As String = part.Trim()

                If candidate.Length = 0 Then
                    Continue For
                End If

                If Not candidate.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal) Then
                    candidate &= Path.DirectorySeparatorChar
                End If

                If String.Equals(candidate, directoryPath, StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If
            Next

            Return False
        End Function

        Public Function StartLiveAsync(opts As TranscriptionOptions, ct As CancellationToken) As Task Implements ITranscriptionEngine.StartLiveAsync
            Init(opts)
            Return Task.CompletedTask
        End Function

        Public Async Function PushAudioAsync(pcm As Byte(), bytesValid As Integer, ct As CancellationToken) As Task Implements ITranscriptionEngine.PushAudioAsync
            If _processor Is Nothing OrElse _cancelled Then
                Return
            End If

            For i As Integer = 0 To bytesValid - 2 Step 2
                Dim s As Short = BitConverter.ToInt16(pcm, i)
                _audioBuffer.Add(s / 32768.0F)
            Next

            If _audioBuffer.Count < PROCESS_THRESHOLD_SAMPLES Then
                Return
            End If

            Dim batch As Single() = _audioBuffer.ToArray()

            If _audioBuffer.Count > ROLLING_OVERLAP_SAMPLES Then
                _audioBuffer.RemoveRange(0, _audioBuffer.Count - ROLLING_OVERLAP_SAMPLES)
            End If

            Await ProcessSegmentsAsync(batch, ct)
        End Function

        Public Async Function StopLiveAsync() As Task Implements ITranscriptionEngine.StopLiveAsync
            _cancelled = True

            If _audioBuffer.Count > 0 AndAlso _processor IsNot Nothing Then
                Await ProcessSegmentsAsync(_audioBuffer.ToArray(), CancellationToken.None)
                _audioBuffer.Clear()
            End If
        End Function

        Public Async Function TranscribeFileAsync(filePath As String, opts As TranscriptionOptions, ct As CancellationToken) As Task Implements ITranscriptionEngine.TranscribeFileAsync
            Init(opts)
            RaiseStatusMessage("Preparing file…")
            Dim samples As Single() = LoadAudioToFloat16k(filePath)
            RaiseStatusMessage("Transcribing file…")
            Await ProcessSegmentsAsync(samples, ct)

            If ct.IsCancellationRequested OrElse _cancelled Then
                RaiseStatusMessage("File transcription canceled.")
            Else
                RaiseStatusMessage("File transcription completed.", 100)
            End If
        End Function

        Private Async Function ProcessSegmentsAsync(samples As Single(), ct As CancellationToken) As Task
            If _processor Is Nothing OrElse samples Is Nothing OrElse samples.Length = 0 Then
                Return
            End If

            Dim enumerator As IAsyncEnumerator(Of SegmentData) = Nothing

            Try
                Dim segs = _processor.ProcessAsync(samples)
                enumerator = segs.GetAsyncEnumerator(ct)

                Dim hasNext As Boolean = True

                While hasNext
                    Try
                        hasNext = Await enumerator.MoveNextAsync()
                    Catch ex As Exception
                        RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("Whisper iter: " & ex.Message, ex, False))
                        Exit While
                    End Try

                    If Not hasNext OrElse ct.IsCancellationRequested OrElse _cancelled Then
                        Exit While
                    End If

                    Dim seg As SegmentData = enumerator.Current
                    Dim text As String = seg.Text

                    text = Regex.Replace(text, "\[.*?\]", String.Empty)
                    text = Regex.Replace(text, "\*.*?\*", String.Empty)

                    If Not String.IsNullOrWhiteSpace(text) Then
                        RaiseEvent FinalResult(Me, New TranscriptionEventArgs(text.Trim(), True))
                    End If
                End While
            Catch ex As Exception
                RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("Whisper proc: " & ex.Message, ex, False))
            End Try

            If enumerator IsNot Nothing Then
                Try
                    Await enumerator.DisposeAsync()
                Catch ex As Exception
                    RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("Whisper dispose: " & ex.Message, ex, False))
                End Try
            End If
        End Function

        Friend Shared Function LoadAudioToFloat16k(filePath As String) As Single()
            Using r As New MediaFoundationReader(filePath)
                Dim fmt As New WaveFormat(16000, 1)

                Using rs As New MediaFoundationResampler(r, fmt)
                    rs.ResamplerQuality = 60

                    Dim sp = rs.ToSampleProvider()
                    Dim acc As New List(Of Single)()
                    Dim buf(1023) As Single
                    Dim n As Integer

                    Do
                        n = sp.Read(buf, 0, buf.Length)
                        If n > 0 Then
                            acc.AddRange(buf.Take(n))
                        End If
                    Loop While n > 0

                    Return acc.ToArray()
                End Using
            End Using
        End Function

        Public Function DisposeAsync() As ValueTask Implements IAsyncDisposable.DisposeAsync
            Return New ValueTask(DisposeAsyncCore())
        End Function

        Private Async Function DisposeAsyncCore() As Task
            If _processor IsNot Nothing Then
                Try
                    Await _processor.DisposeAsync()
                Catch
                End Try
                _processor = Nothing
            End If

            Try
                If _factory IsNot Nothing Then
                    _factory.Dispose()
                End If
            Catch
            End Try
            _factory = Nothing
        End Function
    End Class

End Namespace
