' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ITranscriptionEngine.vb
' Purpose: Defines the common interface for all transcription engines. This
'          ensures that different transcription services (e.g., OpenAI, Google,
'          Vosk) can be used interchangeably by the application.
'
' Architecture:
'  - Contract Definition: Specifies the methods, properties, and events that
'    each transcription engine must implement.
'  - Core Methods: Includes methods for initializing the engine, starting and
'    stopping the transcription process, and processing audio data.
'  - Events: Defines events for communicating transcription results (e.g.,
'    partial, final) and errors back to the consuming components.
'  - Abstraction: Decouples the main application logic from the specific
'    details of any single transcription service.
' =============================================================================


Option Explicit On
Option Strict Off

Imports System.Threading
Imports System.Threading.Tasks

Namespace Transcription

    Public Enum AudioSourceMode
        MicrophoneOnly
        SystemAudioOnly
        MicrophoneAndSystem
    End Enum

    Public Enum EngineKind
        Vosk
        WhisperLocal
        GoogleV1
        GoogleV2
        GeminiTranscribe
        OpenAiRest
        OpenAiRealtime
        AzureSpeechRealtime
        AzureSpeechFastRest
        TeamsAcsRealtime
        GeminiTranscribeLive
    End Enum

    Public Class TranscriptionOptions
        Private _diarization As Boolean = False
        Private _diarizationMaxSpeakers As Integer = 2

        Public Property LanguageCode As String = "auto"

        Public Property EnableDiarization As Boolean
            Get
                Return _diarization
            End Get
            Set(value As Boolean)
                _diarization = value
            End Set
        End Property

        Public Property Diarization As Boolean
            Get
                Return _diarization
            End Get
            Set(value As Boolean)
                _diarization = value
            End Set
        End Property

        Public Property MinSpeakers As Integer = 2

        Public Property MaxSpeakers As Integer
            Get
                Return _diarizationMaxSpeakers
            End Get
            Set(value As Integer)
                _diarizationMaxSpeakers = value
            End Set
        End Property

        Public Property DiarizationMaxSpeakers As Integer
            Get
                Return _diarizationMaxSpeakers
            End Get
            Set(value As Integer)
                _diarizationMaxSpeakers = value
            End Set
        End Property

        Public Property Translate As Boolean = False
        Public Property Model As String = ""
        Public Property VadThreshold As Single = 0.6F
        Public Property VoskSimilarityThreshold As Double = 1.0
        Public Property MultiChannelDiarization As Boolean = False
        Public Property AudioDebugDump As Boolean = False
        Public Property PreferredMicrophoneDeviceId As String = ""
        Public Property PreferredMicrophoneDisplayName As String = ""
        Public Property TurnDetection As String = "server_vad"
        Public Property Prompt As String = ""
    End Class

    Public Class TranscriptionEventArgs
        Inherits EventArgs
        Public Property Text As String = ""
        Public Property Speaker As String = ""    ' "" = unknown
        Public Property IsFinal As Boolean
        Public Sub New(text As String, isFinal As Boolean, Optional speaker As String = "")
            Me.Text = text
            Me.IsFinal = isFinal
            Me.Speaker = speaker
        End Sub
    End Class

    Public Class TranscriptionErrorEventArgs
        Inherits EventArgs
        Public Property Message As String
        Public Property [Exception] As Exception
        Public Property Fatal As Boolean
        Public Sub New(msg As String, ex As Exception, fatal As Boolean)
            Me.Message = msg
            Me.Exception = ex
            Me.Fatal = fatal
        End Sub
    End Class

    Public Class TranscriptionStatusEventArgs
        Inherits EventArgs
        Public Property Message As String
        Public Property ProgressPercent As System.Nullable(Of Integer)

        Public Sub New(msg As String, Optional progressPercent As System.Nullable(Of Integer) = Nothing)
            Me.Message = msg
            Me.ProgressPercent = progressPercent
        End Sub
    End Class

    ''' <summary>
    ''' All engines push 16 kHz, mono, 16-bit signed PCM. Multi-channel diarization engines
    ''' (Google V2) can accept interleaved stereo when <see cref="TranscriptionOptions.MultiChannelDiarization"/>
    ''' is true; in that case the AudioCaptureService delivers stereo PCM (mic = L, system = R).
    ''' </summary>
    Public Interface ITranscriptionEngine
        Inherits IAsyncDisposable

        ReadOnly Property Name As String
        ReadOnly Property Kind As EngineKind
        ReadOnly Property SupportsLiveStreaming As Boolean
        ReadOnly Property SupportsFileTranscription As Boolean
        ReadOnly Property SupportsDiarization As Boolean
        ReadOnly Property SupportsMultiChannelDiarization As Boolean

        Event PartialResult As EventHandler(Of TranscriptionEventArgs)
        Event FinalResult As EventHandler(Of TranscriptionEventArgs)
        Event EngineError As EventHandler(Of TranscriptionErrorEventArgs)
        Event Status As EventHandler(Of TranscriptionStatusEventArgs)

        Function StartLiveAsync(opts As TranscriptionOptions, ct As CancellationToken) As Task
        Function PushAudioAsync(pcm As Byte(), bytesValid As Integer, ct As CancellationToken) As Task
        Function StopLiveAsync() As Task

        ''' <summary>Transcribes a local audio file. Engine is responsible for any decoding/resampling.</summary>
        Function TranscribeFileAsync(filePath As String, opts As TranscriptionOptions, ct As CancellationToken) As Task

    End Interface

End Namespace
