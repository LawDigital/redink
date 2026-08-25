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
        Private Const LIVE_DEDUP_MAX_TOKENS As Integer = 12
        Private Const LIVE_RECENT_TEXT_MAX_CHARS As Integer = 1200
        Private Const LIVE_SILENCE_RMS_THRESHOLD As Double = 0.0025R
        Private Const LIVE_SILENCE_PEAK_THRESHOLD As Double = 0.012R
        Private Const LIVE_NO_SPEECH_REJECT_FLOOR As Single = 0.65F
        Private _cancelled As Boolean
        Private _recentLiveFinalText As String = String.Empty
        Private _liveNoSpeechRejectThreshold As Single = 0.65F
        Private _liveBatchProcessed As Boolean

        Private Shared ReadOnly _runtimeInitLock As New Object()
        Private Shared _runtimeConfigured As Boolean

        Public Sub New(modelRoot As String, modelFileName As String)
            _modelRoot = modelRoot
            _modelFile = modelFileName
        End Sub

        Private Sub RaiseStatusMessage(message As String, Optional progressPercent As System.Nullable(Of Integer) = Nothing)
            RaiseEvent Status(Me, New TranscriptionStatusEventArgs(message, progressPercent))
        End Sub

        Private Sub Init(opts As TranscriptionOptions, isLive As Boolean)
            EnsureRuntimeConfigured()

            Dim modelPath As String = Path.Combine(_modelRoot, _modelFile)
            _factory = WhisperFactory.FromPath(modelPath)
            RaiseStatusMessage("Whisper runtime: " & GetLoadedRuntimeDisplayText())

            Dim vad As Single = 0.6F
            If opts.VadThreshold > 0.0F AndAlso opts.VadThreshold < 1.0F Then
                vad = opts.VadThreshold
            End If

            Dim lang As String = If(String.IsNullOrWhiteSpace(opts.LanguageCode), "auto", opts.LanguageCode).Trim()

            _liveNoSpeechRejectThreshold = Math.Max(LIVE_NO_SPEECH_REJECT_FLOOR, vad)

            Dim builder = _factory.CreateBuilder() _
                .WithThreads(Environment.ProcessorCount) _
                .WithNoSpeechThreshold(vad) _
                .WithTemperature(0.0)

            If String.Equals(lang, "auto", StringComparison.OrdinalIgnoreCase) Then
                builder = builder.WithLanguageDetection()
            Else
                builder = builder.WithLanguage(lang)
            End If

            ' Live audio is deliberately processed in short overlapping batches. Carrying
            ' Whisper's previous decoded text into the next batch can amplify the overlap
            ' into repetition loops. The application already reconciles the resulting live
            ' transcript downstream, so make each live decode independent. For full-file
            ' transcription, retain Whisper's normal previous-text conditioning because it
            ' improves continuity across the model's internal long-form windows.
            If isLive Then
                builder = builder.WithNoContext()
            End If

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

                ' Do not set RuntimeOptions.LibraryPath to the model directory. LibraryPath is a
                ' custom native-library override and would bypass Whisper.net's multi-runtime probing.
                ' Keep the model directory on PATH for backwards-compatible discovery of optional
                ' native dependencies that administrators may already deploy next to the models.
                RuntimeOptions.RuntimeLibraryOrder = New List(Of RuntimeLibrary) From {
                    RuntimeLibrary.Cuda,
                    RuntimeLibrary.Cuda12,
                    RuntimeLibrary.Vulkan,
                    RuntimeLibrary.OpenVino,
                    RuntimeLibrary.Cpu,
                    RuntimeLibrary.CpuNoAvx
                }

                _runtimeConfigured = True
            End SyncLock
        End Sub

        Public Shared Function GetLoadedRuntimeDisplayText() As String
            Dim loaded As System.Nullable(Of RuntimeLibrary) = RuntimeOptions.LoadedLibrary
            If Not loaded.HasValue Then
                Return "Automatic (selected when Whisper is first initialized)"
            End If

            Select Case loaded.Value
                Case RuntimeLibrary.Cuda
                    Return "CUDA 13 (NVIDIA GPU)"
                Case RuntimeLibrary.Cuda12
                    Return "CUDA 12 (NVIDIA GPU)"
                Case RuntimeLibrary.Vulkan
                    Return "Vulkan (GPU)"
                Case RuntimeLibrary.OpenVino
                    Return "OpenVINO (accelerated runtime)"
                Case RuntimeLibrary.Cpu
                    Return "CPU"
                Case RuntimeLibrary.CpuNoAvx
                    Return "CPU (NoAVX compatibility)"
                Case Else
                    Return loaded.Value.ToString()
            End Select
        End Function

        Public Shared Function GetRuntimeOptimizationHint() As String
            Dim loaded As System.Nullable(Of RuntimeLibrary) = RuntimeOptions.LoadedLibrary

            If loaded.HasValue Then
                Select Case loaded.Value
                    Case RuntimeLibrary.Cuda, RuntimeLibrary.Cuda12, RuntimeLibrary.Vulkan
                        Return "GPU acceleration is active. No additional Whisper accelerator setup is required."
                    Case RuntimeLibrary.OpenVino
                        Return "OpenVINO acceleration is active. A compatible GPU runtime may still be faster where supported."
                    Case RuntimeLibrary.Cpu, RuntimeLibrary.CpuNoAvx
                        Return "CPU fallback is active. For GPU acceleration, install a supported accelerator: NVIDIA CUDA Toolkit 13.0.1+ or 12.4.1+, Vulkan Toolkit 1.4.321.1+, or Intel OpenVINO 2024.4+. Whisper will select the best available runtime automatically on the next application start."
                End Select
            End If

            Return "Whisper selects the best available runtime automatically on first use. Optional GPU acceleration: NVIDIA CUDA Toolkit 13.0.1+ or 12.4.1+, Vulkan Toolkit 1.4.321.1+, or Intel OpenVINO 2024.4+."
        End Function

        Public Shared Function IsRuntimeLoaded() As Boolean
            Return RuntimeOptions.LoadedLibrary.HasValue
        End Function

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
            Init(opts, True)
            _audioBuffer.Clear()
            _recentLiveFinalText = String.Empty
            _liveBatchProcessed = False
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

            Await ProcessSegmentsAsync(batch, ct, True)
            _liveBatchProcessed = True
        End Function

        Public Async Function StopLiveAsync() As Task Implements ITranscriptionEngine.StopLiveAsync
            ' After every normal live batch the buffer contains only the rolling overlap.
            ' Process on stop only when audio arrived after that overlap; otherwise the last
            ' batch would be decoded twice. Do this before setting _cancelled, because the
            ' segment loop intentionally stops immediately once cancellation is active.
            Dim hasUnprocessedAudio As Boolean = _audioBuffer.Count > 0 AndAlso
                                                (Not _liveBatchProcessed OrElse _audioBuffer.Count > ROLLING_OVERLAP_SAMPLES)

            If _processor IsNot Nothing AndAlso hasUnprocessedAudio Then
                Await ProcessSegmentsAsync(_audioBuffer.ToArray(), CancellationToken.None, True)
            End If

            _audioBuffer.Clear()
            _cancelled = True
        End Function

        Public Async Function TranscribeFileAsync(filePath As String, opts As TranscriptionOptions, ct As CancellationToken) As Task Implements ITranscriptionEngine.TranscribeFileAsync
            Init(opts, False)
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

        Private Async Function ProcessSegmentsAsync(samples As Single(), ct As CancellationToken, Optional liveBoundaryDeduplication As Boolean = False) As Task
            If _processor Is Nothing OrElse samples Is Nothing OrElse samples.Length = 0 Then
                Return
            End If

            Dim enumerator As IAsyncEnumerator(Of SegmentData) = Nothing

            Try
                Dim segs = _processor.ProcessAsync(samples)
                enumerator = segs.GetAsyncEnumerator(ct)

                Dim hasNext As Boolean = True
                Dim atLiveBatchBoundary As Boolean = liveBoundaryDeduplication AndAlso Not String.IsNullOrWhiteSpace(_recentLiveFinalText)

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

                    If liveBoundaryDeduplication AndAlso Not IsLiveSegmentAcousticallySupported(samples, seg) Then
                        System.Diagnostics.Debug.WriteLine(
                            "[Whisper.Live] Suppressed low-speech segment. NoSpeechProbability=" &
                            seg.NoSpeechProbability.ToString("0.000", System.Globalization.CultureInfo.InvariantCulture) &
                            "; Text=" & If(text, String.Empty))
                        Continue While
                    End If

                    text = Regex.Replace(text, "\[.*?\]", String.Empty)
                    text = Regex.Replace(text, "\*.*?\*", String.Empty)

                    If Not String.IsNullOrWhiteSpace(text) Then
                        text = text.Trim()

                        If liveBoundaryDeduplication AndAlso atLiveBatchBoundary Then
                            text = RemoveRepeatedLivePrefix(_recentLiveFinalText, text)

                            If String.IsNullOrWhiteSpace(text) Then
                                Continue While
                            End If

                            atLiveBatchBoundary = False
                        End If

                        If liveBoundaryDeduplication Then
                            RememberLiveFinalText(text)
                        End If

                        RaiseEvent FinalResult(Me, New TranscriptionEventArgs(text, True))
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


        Private Function IsLiveSegmentAcousticallySupported(samples As Single(), seg As SegmentData) As Boolean
            If samples Is Nothing OrElse samples.Length = 0 OrElse seg Is Nothing Then
                Return False
            End If

            If seg.NoSpeechProbability >= _liveNoSpeechRejectThreshold Then
                Return False
            End If

            Const sampleRate As Integer = 16000
            Const marginSamples As Integer = sampleRate \ 10

            Dim startSample As Integer = Math.Max(0, CInt(Math.Floor(seg.Start.TotalSeconds * sampleRate)) - marginSamples)
            Dim endSample As Integer = Math.Min(samples.Length, CInt(Math.Ceiling(seg.End.TotalSeconds * sampleRate)) + marginSamples)

            If endSample <= startSample Then
                startSample = 0
                endSample = samples.Length
            End If

            Dim sumSquares As Double = 0.0R
            Dim peak As Double = 0.0R
            Dim count As Integer = 0

            For i As Integer = startSample To endSample - 1
                Dim amplitude As Double = Math.Abs(CDbl(samples(i)))
                If amplitude > peak Then
                    peak = amplitude
                End If

                sumSquares += amplitude * amplitude
                count += 1
            Next

            If count = 0 Then
                Return False
            End If

            Dim rms As Double = Math.Sqrt(sumSquares / count)

            ' Reject only when both measures indicate near-silence. Using both avoids
            ' discarding quiet speech while suppressing Whisper's common silence tails.
            Return rms >= LIVE_SILENCE_RMS_THRESHOLD OrElse peak >= LIVE_SILENCE_PEAK_THRESHOLD
        End Function

        Private Shared Function RemoveRepeatedLivePrefix(previousText As String, currentText As String) As String
            If String.IsNullOrWhiteSpace(previousText) OrElse String.IsNullOrWhiteSpace(currentText) Then
                Return currentText
            End If

            Dim previousTokens As MatchCollection = Regex.Matches(previousText, "[\p{L}\p{N}]+(?:['’\-][\p{L}\p{N}]+)*")
            Dim currentTokens As MatchCollection = Regex.Matches(currentText, "[\p{L}\p{N}]+(?:['’\-][\p{L}\p{N}]+)*")

            If previousTokens.Count < 2 OrElse currentTokens.Count < 2 Then
                Return currentText
            End If

            Dim maximumOverlap As Integer = Math.Min(LIVE_DEDUP_MAX_TOKENS, Math.Min(previousTokens.Count, currentTokens.Count))
            Dim overlapCount As Integer = 0

            For candidateCount As Integer = maximumOverlap To 2 Step -1
                Dim matches As Boolean = True
                Dim previousStart As Integer = previousTokens.Count - candidateCount

                For tokenIndex As Integer = 0 To candidateCount - 1
                    If Not String.Equals(previousTokens(previousStart + tokenIndex).Value,
                                         currentTokens(tokenIndex).Value,
                                         StringComparison.OrdinalIgnoreCase) Then
                        matches = False
                        Exit For
                    End If
                Next

                If matches Then
                    overlapCount = candidateCount
                    Exit For
                End If
            Next

            If overlapCount = 0 Then
                Return currentText
            End If

            If overlapCount >= currentTokens.Count Then
                Return String.Empty
            End If

            Dim firstNewToken As Match = currentTokens(overlapCount)
            Return currentText.Substring(firstNewToken.Index).Trim()
        End Function

        Private Sub RememberLiveFinalText(text As String)
            If String.IsNullOrWhiteSpace(text) Then
                Return
            End If

            If String.IsNullOrWhiteSpace(_recentLiveFinalText) Then
                _recentLiveFinalText = text.Trim()
            Else
                _recentLiveFinalText = (_recentLiveFinalText & " " & text.Trim()).Trim()
            End If

            If _recentLiveFinalText.Length > LIVE_RECENT_TEXT_MAX_CHARS Then
                _recentLiveFinalText = _recentLiveFinalText.Substring(_recentLiveFinalText.Length - LIVE_RECENT_TEXT_MAX_CHARS)
            End If
        End Sub

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
