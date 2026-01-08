' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.Transcriptor.vb
' Purpose: Provides real-time speech-to-text transcription using Vosk, Whisper, 
'          or Google Cloud Speech-to-Text APIs with speaker identification support.
'
' Architecture:
'  - Multi-Engine Support: Vosk (local), Whisper (local), Google STT (cloud)
'  - Audio Input: Microphone capture via NAudio with optional system audio mixing
'  - Real-Time Processing: Streaming transcription with partial/final results
'  - Speaker Identification: Vosk embeddings (cosine similarity), Google diarization,
'    Whisper translation option
'  - File Transcription: Batch processing of audio files (WAV, MP3, AAC, M4A, WMA)
'  - Stream Management: Google STT uses resilient streaming with automatic recovery,
'    watchdog timer, and ring buffer for seamless reconnection
'  - OAuth2 Authentication: Token caching and refresh for Google Cloud APIs
'  - Sleep Prevention: Prevents system sleep during active transcription
'  - Post-Processing: Optional LLM-based text processing via prompt library
'
' External Dependencies:
'  - Vosk: Local speech recognition models
'  - Whisper.net: Local GGML models  
'  - Google.Cloud.Speech.V1: Cloud-based STT with diarization
'  - NAudio: Audio capture and format conversion
'  - SharedLibrary.SharedMethods: UI helpers, LLM integration, OAuth2
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Collections.Concurrent
Imports System.Data
Imports System.Diagnostics
Imports System.Drawing
Imports System.Globalization
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text.Json.Serialization
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Windows
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports DiffPlex
Imports DiffPlex.DiffBuilder
Imports DocumentFormat.OpenXml
Imports Google.Cloud.Speech.V1
Imports Google.Protobuf
Imports Grpc.Core
Imports Microsoft.Office.Interop.Word
Imports NAudio.CoreAudioApi
Imports NAudio.Wave
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports Vosk
Imports Whisper.net
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods
Imports SharedLibrary.SharedLibrary


Partial Public Class ThisAddIn

    ''' <summary>
    ''' Main form for speech-to-text transcription supporting multiple engines (Vosk, Whisper, Google STT).
    ''' Provides real-time audio capture, speaker identification, and optional post-processing.
    ''' </summary>
    Public Class TranscriptionForm

        Inherits Form

        ' ============================================================================
        ' P/Invoke: Sleep Prevention
        ' Prevents Windows from entering sleep mode during active transcription
        ' ============================================================================

        ''' <summary>
        ''' Win32 API function to prevent system sleep during transcription.
        ''' </summary>
        <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
        Private Shared Function SetThreadExecutionState(ByVal esFlags As UInteger) As UInteger
        End Function

        ' Constants for sleep prevention
        Private Const ES_CONTINUOUS As UInteger = &H80000000UI
        Private Const ES_SYSTEM_REQUIRED As UInteger = &H1UI
        Private Const ES_DISPLAY_REQUIRED As UInteger = &H2UI ' Optional: keeps the display on too

        ''' <summary>
        ''' Tracks whether this form instance set the system sleep lock.
        ''' True = this form is responsible for releasing; False = another component owns it.
        ''' </summary>
        Private _iSetTheSleepLock As Boolean = False

        ' ============================================================================
        ' UI Components
        ' ============================================================================
        Private RichTextBox1 As Forms.RichTextBox
        Private StartButton As Forms.Button
        Private StopButton As Forms.Button
        Private ClearButton As Forms.Button
        Private LoadButton As Forms.Button
        Private AudioButton As Forms.Button
        Private QuitButton As Forms.Button
        Private ProcessButton As Forms.Button
        Private cultureComboBox As Forms.ComboBox
        Private deviceComboBox As Forms.ComboBox
        Private processCombobox As Forms.ComboBox
        Private SpeakerIdent As System.Windows.Forms.CheckBox
        Private SpeakerDistance As Forms.TextBox
        Private Label1 As Label
        Private Label2 As Label
        Private StatusLabel As Label
        Private PartialTextLabel As Label
        Private ButtonPanel As Panel

        ' ============================================================================
        ' Prompt Library for Post-Processing
        ' ============================================================================
        ''' <summary>
        ''' Stores titles of available transcript processing prompts.
        ''' </summary>
        Private TranscriptPromptsTitles As New List(Of String)

        ''' <summary>
        ''' Stores the actual prompt content for transcript processing.
        ''' </summary>
        Private TranscriptPromptsLibrary As New List(Of String)

        ' ============================================================================
        ' Vosk STT Engine Variables
        ' ============================================================================
        ''' <summary>
        ''' Vosk speech recognizer instance for local speech recognition.
        ''' </summary>
        Private recognizer As VoskRecognizer

        ''' <summary>
        ''' Tooltip text for Vosk speaker identification feature.
        ''' </summary>
        Private Const VoskTooltip = "Only for Vosk: Set similarity threshold for speaker identification (0.5-0.7 for real-time speaker tracking, 1.0-1.5 for meetings/interviews)"

        ''' <summary>
        ''' Toggle button text for Vosk speaker identification.
        ''' </summary>
        Private Const VoskToggle = "Iden"

        ' ============================================================================
        ' Whisper STT Engine Variables
        ' ============================================================================
        ''' <summary>
        ''' Whisper speech processor instance for local speech recognition with translation support.
        ''' </summary>
        Private WhisperRecognizer As WhisperProcessor

        ''' <summary>
        ''' Audio buffer for accumulating samples before Whisper processing.
        ''' </summary>
        Private audioBuffer As New List(Of Single)

        ''' <summary>
        ''' Tooltip text for Whisper translation and voice detection threshold.
        ''' </summary>
        Private Const WhisperTooltip = "Only for Whisper: Select if text shall be translated to English and the threshold for detecting voice (default = 0.6, increase for noisy environments)"

        ''' <summary>
        ''' Toggle button text for Whisper translation feature.
        ''' </summary>
        Private Const WhisperToggle = "Trans"

        ' ============================================================================
        ' Google Cloud STT Engine Variables
        ' ============================================================================
        ''' <summary>
        ''' Indicates whether Google Cloud Speech-to-Text is available and configured.
        ''' </summary>
        Private GoogleSpeech As Boolean = False

        ''' <summary>
        ''' Indicates whether to use secondary API credentials for Google STT.
        ''' </summary>
        Private STTSecondAPI As Boolean = False

        ''' <summary>
        ''' Indicates whether Google STT is the selected engine.
        ''' </summary>
        Private IsGoogle As Boolean = False

        ''' <summary>
        ''' Tooltip text for Google speaker diarization feature.
        ''' </summary>
        Private Const GoogleTooltip = "Only for Google: Set the maximum number of speakers expected for diarization (speaker tracking)"

        ''' <summary>
        ''' Toggle button text for Google speaker diarization.
        ''' </summary>
        Private Const GoogleToggle = "Iden"

        ''' <summary>
        ''' Background task that reads responses from Google's bidirectional gRPC stream.
        ''' </summary>
        Private googleReaderTask As System.Threading.Tasks.Task

        ''' <summary>
        ''' Cancellation token source for the Google reader task.
        ''' </summary>
        Private readerCts As CancellationTokenSource = New CancellationTokenSource()

        ''' <summary>
        ''' Bidirectional gRPC stream for Google Cloud Speech-to-Text.
        ''' </summary>
        Private _stream As SpeechClient.StreamingRecognizeStream

        ''' <summary>
        ''' Character position in RichTextBox where current Google transcription session started.
        ''' </summary>
        Private googleTranscriptStart As Integer = 0

        ''' <summary>
        ''' Google Cloud Speech client instance.
        ''' </summary>
        Private client As SpeechClient

        ''' <summary>
        ''' Language code for Google STT (e.g., "en-US", "de-DE").
        ''' </summary>
        Private GoogleLanguageCode As String = ""

        ''' <summary>
        ''' Thread-safe queue for audio chunks sent to Google STT stream.
        ''' </summary>
        Private audioQueue As New System.Collections.Concurrent.BlockingCollection(Of ByteString)()

        ''' <summary>
        ''' Flag indicating whether Google stream has completed writing.
        ''' </summary>
        Private _googleStreamCompleted As Boolean = False

        ''' <summary>
        ''' Maximum streaming duration in milliseconds (4 minutes 50 seconds) before forcing reconnection.
        ''' </summary>
        Private Const STREAMING_LIMIT_MS As Integer = 290000

        ''' <summary>
        ''' Timestamp when current Google streaming session started.
        ''' </summary>
        Private streamingStartTime As DateTime

        ' ============================================================================
        ' Google STT Stream Recovery Components
        ' ============================================================================
        ''' <summary>
        ''' Ring buffer holding last 50 audio chunks for stream recovery after timeout or error.
        ''' </summary>
        Private ReadOnly ringBuffer As New Queue(Of Google.Protobuf.ByteString)()

        ''' <summary>
        ''' Maximum number of audio chunks stored in ring buffer.
        ''' </summary>
        Private Const RING_BUFFER_SIZE As Integer = 50

        ''' <summary>
        ''' Semaphore ensuring only one stream recovery operation executes at a time.
        ''' </summary>
        Private ReadOnly recoverySemaphore As New System.Threading.SemaphoreSlim(1, 1)

        ''' <summary>
        ''' Background task that writes audio chunks from queue to Google stream.
        ''' </summary>
        Private writerTask As System.Threading.Tasks.Task

        ' ============================================================================
        ' Google STT Watchdog Timer Components
        ' ============================================================================
        ''' <summary>
        ''' Watchdog timer that monitors API responsiveness and triggers recovery if needed.
        ''' </summary>
        Private _apiWatchdogTimer As System.Threading.Timer

        ''' <summary>
        ''' Tracks last API response time (Ticks) for watchdog monitoring.
        ''' Thread-safe via Interlocked operations.
        ''' </summary>
        Private _lastApiResponseTicks As Long

        ''' <summary>
        ''' Number of seconds of API silence before triggering stream recovery.
        ''' </summary>
        Private Const API_RESPONSE_TIMEOUT_SECONDS As Integer = 3

        ''' <summary>
        ''' Last known partial result from Google API, committed during recovery if stream fails.
        ''' </summary>
        Private _lastKnownPartialResult As String = ""

        ''' <summary>
        ''' Partial text that was just committed during recovery, used to detect duplicate final results.
        ''' </summary>
        Private _justCommittedPartialText As String = ""

        ' ============================================================================
        ' Google STT Speaker Diarization Components
        ' ============================================================================
        ''' <summary>
        ''' Maps temporary Google speaker tags to consistent human-readable labels (Speaker 1, Speaker 2, etc.).
        ''' Maintains session-wide speaker identity across stream recoveries.
        ''' </summary>
        Private _speakerTagToLabelMap As New Dictionary(Of Integer, String)

        ''' <summary>
        ''' Counter for assigning unique speaker numbers.
        ''' </summary>
        Private _nextSpeakerNumber As Integer = 1

        ' ============================================================================
        ' Audio Capture & Processing
        ' ============================================================================
        ''' <summary>
        ''' NAudio microphone capture instance.
        ''' </summary>
        Private waveIn As WaveInEvent

        ''' <summary>
        ''' Flag indicating whether transcription is currently active.
        ''' </summary>
        Private capturing As Boolean = False

        ''' <summary>
        ''' Current partial transcription text (before finalization).
        ''' </summary>
        Private partialText As String = ""

        ''' <summary>
        ''' Accumulates all finalized transcription text.
        ''' </summary>
        Private finalText As New StringBuilder()

        ''' <summary>
        ''' Flag indicating whether transcription was canceled by user.
        ''' </summary>
        Private STTCanceled As Boolean = False

        ''' <summary>
        ''' Cancellation token source for async operations.
        ''' </summary>
        Private cts As CancellationTokenSource = New CancellationTokenSource()

        ''' <summary>
        ''' Currently selected STT model ("vosk", "whisper", or "google").
        ''' </summary>
        Private STTModel As String = "whisper"

        ' ============================================================================
        ' System Audio Loopback Capture (for mixing mic + system audio)
        ' ============================================================================
        ''' <summary>
        ''' Unused placeholder for loopback capture.
        ''' </summary>
        Private loopback As WasapiLoopbackCapture

        ''' <summary>
        ''' Unused placeholder for loopback buffer.
        ''' </summary>
        Private loopbackBuffer As BufferedWaveProvider

        ''' <summary>
        ''' WASAPI loopback capture instance for system audio.
        ''' </summary>
        Private loopbackCapture As WasapiLoopbackCapture

        ''' <summary>
        ''' Buffer provider for raw loopback audio in native format.
        ''' </summary>
        Private loopbackRawProvider As BufferedWaveProvider

        ''' <summary>
        ''' Resampler converting loopback audio to microphone format (16 kHz mono).
        ''' </summary>
        Private loopbackResampler As MediaFoundationResampler

        ''' <summary>
        ''' Flag indicating whether user selected multi-source capture (mic + system audio).
        ''' </summary>
        Private _multiSourceSelected As Boolean = False

        ''' <summary>
        ''' Gets whether multi-source audio capture (mic + system audio) is enabled.
        ''' </summary>
        Private ReadOnly Property MultiSourceEnabled As Boolean
            Get
                Return _multiSourceSelected
            End Get
        End Property

        ' ============================================================================
        ' OAuth2 Token Management for Google STT
        ' ============================================================================
        ''' <summary>
        ''' Cached OAuth2 access token for primary Google STT API.
        ''' </summary>
        Private sttAccessToken1 As String = String.Empty

        ''' <summary>
        ''' Expiry time for primary OAuth2 token.
        ''' </summary>
        Private sttTokenExpiry1 As DateTime = DateTime.MinValue

        ''' <summary>
        ''' Cached OAuth2 access token for secondary Google STT API.
        ''' </summary>
        Private sttAccessToken2 As String = String.Empty

        ''' <summary>
        ''' Expiry time for secondary OAuth2 token.
        ''' </summary>
        Private sttTokenExpiry2 As DateTime = DateTime.MinValue

        ''' <summary>
        ''' Formats a raw private key string into PEM format with 64-character line breaks.
        ''' </summary>
        ''' <param name="rawKey">The raw private key string (may contain \n escapes)</param>
        ''' <returns>Properly formatted PEM private key with header and footer</returns>
        Public Shared Function FormatPrivateKey(rawKey As String) As String
            Dim noEscapes = rawKey.Replace("\n", "")
            Dim sb As New System.Text.StringBuilder()
            For i As Integer = 0 To noEscapes.Length - 1 Step 64
                Dim chunk = If(i + 64 <= noEscapes.Length,
                  noEscapes.Substring(i, 64),
                  noEscapes.Substring(i))
                sb.AppendLine(chunk)
            Next
            Return "-----BEGIN PRIVATE KEY-----" & vbLf &
       sb.ToString() &
       "-----END PRIVATE KEY-----" & vbLf
        End Function

        ''' <summary>
        ''' Retrieves a fresh OAuth2 access token for STT API, using cached token if still valid.
        ''' </summary>
        ''' <param name="useSecond">True to use secondary API credentials, False for primary</param>
        ''' <returns>Valid access token or empty string on error</returns>
        Private Async Function GetFreshSTTToken(useSecond As Boolean) As System.Threading.Tasks.Task(Of String)

            Try
                Dim token As String
                Dim expiry As DateTime

                If useSecond Then
                    token = sttAccessToken2
                    expiry = sttTokenExpiry2
                Else
                    token = sttAccessToken1
                    expiry = sttTokenExpiry1
                End If

                If String.IsNullOrEmpty(token) OrElse DateTime.UtcNow >= expiry Then
                    ' Select API parameters based on which credential set to use
                    Dim clientEmail = If(useSecond, INI_OAuth2ClientMail_2, INI_OAuth2ClientMail)
                    Dim scopes = If(useSecond, INI_OAuth2Scopes_2, INI_OAuth2Scopes)
                    Dim rawKey = If(useSecond, INI_APIKey_2, INI_APIKey)
                    Dim authServer = If(useSecond, INI_OAuth2Endpoint_2, INI_OAuth2Endpoint)
                    Dim life = If(useSecond, INI_OAuth2ATExpiry_2, INI_OAuth2ATExpiry)

                    ' Configure GoogleOAuthHelper
                    GoogleOAuthHelper.client_email = clientEmail
                    GoogleOAuthHelper.private_key = FormatPrivateKey(rawKey)
                    GoogleOAuthHelper.scopes = scopes
                    GoogleOAuthHelper.token_uri = authServer
                    GoogleOAuthHelper.token_life = life

                    ' Fetch new token
                    Dim newToken As String = Await GoogleOAuthHelper.GetAccessToken()
                    Dim newExpiry As DateTime = DateTime.UtcNow.AddSeconds(life - 300)

                    If useSecond Then
                        sttAccessToken2 = newToken
                        sttTokenExpiry2 = newExpiry
                    Else
                        sttAccessToken1 = newToken
                        sttTokenExpiry1 = newExpiry
                    End If

                    token = newToken
                End If

                Return token

            Catch ex As System.Exception
                System.Windows.Forms.MessageBox.Show(
        $"Error fetching STT token: {ex.Message}",
        "Transcription Error",
        System.Windows.Forms.MessageBoxButtons.OK,
        System.Windows.Forms.MessageBoxIcon.Error)
                Return String.Empty
            End Try
        End Function

        ''' <summary>
        ''' Initializes the transcription form UI and loads available speech recognition models.
        ''' Detects Google STT availability based on OAuth2 configuration.
        ''' </summary>
        Public Sub New()
            ' Initialize UI Components
            InitializeComponents()

            Me.AutoScaleMode = AutoScaleMode.Dpi

            ' Load available speech recognition models
            Dim modelPath As String = Globals.ThisAddIn.INI_SpeechModelPath
            Dim modelsexist As Boolean = False

            Dim Endpoint As String = INI_Endpoint
            Dim Endpoint_2 As String = INI_Endpoint_2

            ' Detect Google STT configuration
            If Endpoint.Contains(GoogleIdentifier) And INI_OAuth2 Then
                STTSecondAPI = False
                IsGoogle = True
            ElseIf Endpoint_2.Contains(GoogleIdentifier) And INI_OAuth2_2 Then
                STTSecondAPI = True
                IsGoogle = True
            End If

            If IsGoogle And Not String.IsNullOrWhiteSpace(STTEndpoint) Then
                GoogleSpeech = True
                cultureComboBox.Items.Add(GoogleSTT_Desc)
                modelsexist = True
            End If

            ' Enumerate Vosk and Whisper models from configured path
            If Directory.Exists(modelPath) Then
                For Each dir As String In Directory.GetDirectories(modelPath)
                    Dim dirName As String = System.IO.Path.GetFileName(dir)
                    If dirName.StartsWith("vosk-model") Then
                        cultureComboBox.Items.Add(dirName)
                        modelsexist = True
                    End If
                Next

                For Each file As String In Directory.GetFiles(modelPath)
                    Dim fileName As String = System.IO.Path.GetFileName(file)
                    If fileName.StartsWith("ggml") Then
                        cultureComboBox.Items.Add(fileName)
                        modelsexist = True
                    End If
                Next

            End If

            ' Pre-select the last used model if it exists in the list
            Dim lastModel As String = My.Settings.LastSpeechModel
            If Not String.IsNullOrEmpty(lastModel) AndAlso cultureComboBox.Items.Contains(lastModel) Then
                cultureComboBox.SelectedItem = lastModel
            End If

            ' Wire up event handlers for dynamic UI updates
            AddHandler Me.cultureComboBox.MouseMove, AddressOf cultureComboBox_MouseMove

            LoadAudioDevices()

            AddHandler Me.deviceComboBox.MouseMove, AddressOf deviceComboBox_MouseMove

            AddHandler Me.deviceComboBox.SelectedIndexChanged, AddressOf Me.deviceComboBox_SelectedIndexChanged

            LoadAndPopulateProcessComboBox(Globals.ThisAddIn.INI_PromptLibPath_Transcript, processCombobox)

            ' Update tooltips based on selected model
            Dim index As Integer = Me.cultureComboBox.SelectedIndex
            If index >= 0 Then
                If Me.cultureComboBox.Items(index).startswith(GoogleSTT_Desc) Then
                    Me.SpeakerIdent.Text = GoogleToggle
                    ToolTip.SetToolTip(Me.SpeakerDistance, GoogleTooltip)
                    ToolTip.SetToolTip(Me.SpeakerIdent, GoogleTooltip)

                ElseIf Me.cultureComboBox.Items(index).startswith("ggml") Then
                    Me.SpeakerIdent.Text = WhisperToggle
                    ToolTip.SetToolTip(Me.SpeakerDistance, WhisperTooltip)
                    ToolTip.SetToolTip(Me.SpeakerIdent, WhisperTooltip)
                Else
                    Me.SpeakerIdent.Text = VoskToggle
                    ToolTip.SetToolTip(Me.SpeakerDistance, VoskTooltip)
                    ToolTip.SetToolTip(Me.SpeakerIdent, VoskTooltip)
                End If
            End If

            ' Wire up button event handlers
            AddHandler StartButton.Click, AddressOf StartButton_Click
            AddHandler StopButton.Click, AddressOf StopButton_Click
            AddHandler ClearButton.Click, AddressOf ClearButton_Click
            AddHandler LoadButton.Click, AddressOf LoadButton_Click
            AddHandler AudioButton.Click, AddressOf AudioButton_Click
            AddHandler QuitButton.Click, AddressOf QuitButton_Click
            AddHandler ProcessButton.Click, AddressOf ProcessButton_Click

            ' Make window resizable
            Me.MinimumSize = New System.Drawing.Size(800, 440)

            If Not modelsexist Then
                ShowCustomMessageBox($"No Vosk or Whisper models have been found at the configured path ('{modelPath}'). A model is necessary for transcribing. You can download models for free at {VoskSource} and {WhisperSource}.", $"{AN} Transcriptor")
                Me.Close()
            End If
        End Sub

        ''' <summary>
        ''' Tooltip control for displaying context-sensitive help.
        ''' </summary>
        Private ToolTip As New Forms.ToolTip()

        ''' <summary>
        ''' Updates multi-source capture flag when device selection changes.
        ''' Runs on UI thread—safe to read SelectedItem directly.
        ''' </summary>
        Private Sub deviceComboBox_SelectedIndexChanged(sender As Object, e As EventArgs)
            Dim s As String = TryCast(Me.deviceComboBox.SelectedItem, String)
            _multiSourceSelected = Not String.IsNullOrEmpty(s) _
                        AndAlso s.EndsWith("(plus audio output)")
        End Sub

        ''' <summary>
        ''' Updates tooltip and speaker identification controls based on selected model.
        ''' Handles mouse move over culture (model) combo box.
        ''' </summary>
        Private Sub cultureComboBox_MouseMove(sender As Object, e As MouseEventArgs)
            Dim index As Integer = Me.cultureComboBox.SelectedIndex
            If index >= 0 Then
                ToolTip.SetToolTip(Me.cultureComboBox, Me.cultureComboBox.Items(index).ToString())
                If Me.cultureComboBox.Items(index).startswith(GoogleSTT_Desc) Then
                    Me.SpeakerIdent.Text = GoogleToggle
                    ToolTip.SetToolTip(Me.SpeakerDistance, GoogleTooltip)
                    ToolTip.SetToolTip(Me.SpeakerIdent, GoogleTooltip)
                ElseIf Me.cultureComboBox.Items(index).startswith("ggml") Then
                    Me.SpeakerIdent.Text = WhisperToggle
                    ToolTip.SetToolTip(Me.SpeakerDistance, WhisperTooltip)
                    ToolTip.SetToolTip(Me.SpeakerIdent, WhisperTooltip)
                Else
                    Me.SpeakerIdent.Text = VoskToggle
                    ToolTip.SetToolTip(Me.SpeakerDistance, VoskTooltip)
                    ToolTip.SetToolTip(Me.SpeakerIdent, VoskTooltip)
                End If
            End If
        End Sub

        ''' <summary>
        ''' Displays device name tooltip on mouse hover.
        ''' </summary>
        Private Sub deviceComboBox_MouseMove(sender As Object, e As MouseEventArgs)
            Dim index As Integer = Me.deviceComboBox.SelectedIndex
            If index >= 0 Then
                ToolTip.SetToolTip(Me.deviceComboBox, Me.deviceComboBox.Items(index).ToString())
            End If
        End Sub

        ''' <summary>
        ''' Configures the audio output device for loopback capture (system audio mixing).
        ''' Displays device selection dialog and saves choice to settings.
        ''' </summary>
        Public Sub ConfigureAudioOutputDevice()
            ' 1) Enumerate all active audio render endpoints (output devices)
            Dim enumerator As New MMDeviceEnumerator()
            Dim devices As MMDeviceCollection =
enumerator.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active)

            ' 2) Pack FriendlyNames and IDs into parallel arrays, including Default at index 0
            Dim totalCount As Integer = devices.Count + 1
            Dim deviceNames(totalCount - 1) As String
            Dim deviceIds(totalCount - 1) As String

            ' 2a) Default Audio Output Device (as used by WasapiLoopbackCapture)
            deviceNames(0) = "Default Audio Output Device"
            deviceIds(0) = String.Empty

            ' 2b) All other devices starting at index 1
            For i As Integer = 0 To devices.Count - 1
                deviceNames(i + 1) = devices(i).FriendlyName
                deviceIds(i + 1) = devices(i).ID
            Next

            ' 3) Determine currently saved device from settings (empty ID → Default)
            Dim currentDeviceId As String = My.Settings.AudioOutputDevice
            Dim currentDeviceName As String = String.Empty
            Dim idxSaved As Integer = Array.IndexOf(deviceIds, currentDeviceId)
            If idxSaved >= 0 Then
                currentDeviceName = deviceNames(idxSaved)
            End If

            ' 4) Build prompt for selection dialog
            Dim prompt As String = "Choose the audio output device for capturing"
            If Not String.IsNullOrEmpty(currentDeviceName) Then
                prompt &= $" (currently: {currentDeviceName})"
            End If
            prompt &= ":"

            ' 5) Show selection dialog
            Dim selection As String = ShowSelectionForm(
prompt,
$"{AN} Transcriptor",
deviceNames)

            ' 6) If selection valid, determine index and set/clear settings
            If Not String.IsNullOrEmpty(selection) AndAlso selection <> "esc" Then
                Dim chosenIndex As Integer = Array.IndexOf(deviceNames, selection)
                If chosenIndex >= 0 Then
                    If chosenIndex = 0 Then
                        ' Default selected → clear setting
                        My.Settings.AudioOutputDevice = String.Empty
                    Else
                        ' Save specific device ID
                        My.Settings.AudioOutputDevice = deviceIds(chosenIndex)
                    End If

                    Try
                        My.Settings.Save()
                    Catch ex As System.Exception
                        ShowCustomMessageBox($"Error saving audio output device setting: {ex.Message}")
                    End Try
                End If
            End If
        End Sub

        ''' <summary>
        ''' Initializes all UI controls and builds the form layout using TableLayoutPanel.
        ''' Sets up a DPI-aware form with three rows: model/source selectors, transcript area, and action buttons.
        ''' </summary>
        Private Sub InitializeComponents()
            ' DPI-aware form setup
            Me.Font = New System.Drawing.Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
            Me.AutoScaleMode = AutoScaleMode.Font
            Me.Text = $"{AN} Transcriptor (editable text, audio will not be stored)"
            Me.FormBorderStyle = FormBorderStyle.Sizable

            ' Create controls

            ' Transcript area
            Me.RichTextBox1 = New RichTextBox() With {
.Font = New System.Drawing.Font("Segoe UI", 10.0F, FontStyle.Regular, GraphicsUnit.Point),
.Multiline = True,
.ScrollBars = RichTextBoxScrollBars.Vertical,
.Dock = DockStyle.Fill
}

            ' Selector labels
            Me.Label1 = New Label() With {.Text = "Model:", .AutoSize = True}
            Me.Label2 = New Label() With {.Text = "Source:", .AutoSize = True}

            ' Model / source dropdowns (start 50px wider)
            Me.cultureComboBox = New System.Windows.Forms.ComboBox() With {
.DropDownStyle = ComboBoxStyle.DropDownList,
.Width = 250
}
            Me.deviceComboBox = New System.Windows.Forms.ComboBox() With {
.DropDownStyle = ComboBoxStyle.DropDownList,
.Width = 450
}

            ' Speaker toggle + threshold
            Me.SpeakerIdent = New System.Windows.Forms.CheckBox() With {.Text = VoskToggle, .AutoSize = True}
            Me.SpeakerDistance = New System.Windows.Forms.TextBox() With {
.Text = If(My.Settings.LastSpeakerDistance <= 0, "1.0", My.Settings.LastSpeakerDistance.ToString()),
.Width = 50,
.AutoSize = False
}

            ' Status + partial text
            Me.StatusLabel = New Label() With {
.Text = "Transcribing:",
.AutoSize = True,
.Dock = DockStyle.Top
}
            Me.PartialTextLabel = New Label() With {
.Text = "...",
.AutoSize = True,
.MinimumSize = New System.Drawing.Size(0, 70),
.Dock = DockStyle.Top
}

            ' Action buttons + bottom combobox
            Me.StartButton = New System.Windows.Forms.Button() With {.Text = "Start", .AutoSize = True}
            Me.StopButton = New System.Windows.Forms.Button() With {.Text = "Stop", .AutoSize = True, .Enabled = False}
            Me.ClearButton = New System.Windows.Forms.Button() With {.Text = "Clear", .AutoSize = True}
            Me.LoadButton = New System.Windows.Forms.Button() With {.Text = "Load", .AutoSize = True}
            Me.AudioButton = New System.Windows.Forms.Button() With {.Text = "Dev", .AutoSize = True}
            Me.QuitButton = New System.Windows.Forms.Button() With {.Text = "Quit", .AutoSize = True}
            Me.ProcessButton = New System.Windows.Forms.Button() With {.Text = "Process:", .AutoSize = True}
            Me.processCombobox = New System.Windows.Forms.ComboBox() With {
.DropDownStyle = ComboBoxStyle.DropDownList,
.Width = 250
}

            ' Add right margin to prevent control crowding
            Dim pad As New Padding(0, 0, 10, 0)
            For Each ctl In {Label1, cultureComboBox, Label2, deviceComboBox, SpeakerIdent, SpeakerDistance,
             StartButton, StopButton, ClearButton, LoadButton, AudioButton, QuitButton, ProcessButton}
                ctl.Margin = pad
            Next
            processCombobox.Margin = pad

            ' Build layout

            ' Root: 3 rows—top selectors, middle transcript, bottom actions
            Dim root As New TableLayoutPanel() With {
.Dock = DockStyle.Fill,
.AutoSize = True,
.AutoSizeMode = AutoSizeMode.GrowAndShrink,
.ColumnCount = 1,
.RowCount = 3,
.Padding = New Padding(10)
}
            root.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            root.RowStyles.Add(New RowStyle(SizeType.AutoSize))    ' row0: selectors
            root.RowStyles.Add(New RowStyle(SizeType.Percent, 100)) ' row1: transcript
            root.RowStyles.Add(New RowStyle(SizeType.AutoSize))    ' row2: actions

            ' Row 0: selectors laid out in a TableLayoutPanel so combos stretch
            Dim topRow As New TableLayoutPanel() With {
.Dock = DockStyle.Top,
.AutoSize = False,
.Height = cultureComboBox.PreferredHeight + 10,
.ColumnCount = 6,
.RowCount = 1,
.Padding = New Padding(0, 0, 0, 10)
}
            topRow.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            topRow.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))
            topRow.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            topRow.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))
            topRow.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            topRow.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))

            cultureComboBox.Dock = DockStyle.Fill
            deviceComboBox.Dock = DockStyle.Fill

            topRow.Controls.Add(Label1, 0, 0)
            topRow.Controls.Add(cultureComboBox, 1, 0)
            topRow.Controls.Add(Label2, 2, 0)
            topRow.Controls.Add(deviceComboBox, 3, 0)
            topRow.Controls.Add(SpeakerIdent, 4, 0)
            topRow.Controls.Add(SpeakerDistance, 5, 0)

            root.Controls.Add(topRow, 0, 0)

            ' Row 1: status, partial, then main RichTextBox
            Dim mid As New TableLayoutPanel() With {
.Dock = DockStyle.Fill,
.AutoSize = True,
.AutoSizeMode = AutoSizeMode.GrowAndShrink,
.ColumnCount = 1,
.RowCount = 3
}
            mid.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            mid.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            mid.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            mid.RowStyles.Add(New RowStyle(SizeType.Percent, 100))

            mid.Controls.Add(StatusLabel, 0, 0)
            mid.Controls.Add(PartialTextLabel, 0, 1)
            mid.Controls.Add(RichTextBox1, 0, 2)

            root.Controls.Add(mid, 0, 1)

            ' Row 2: bottom actions in a stretchy TableLayoutPanel
            Dim bottomRow As New TableLayoutPanel() With {
.Dock = DockStyle.Bottom,
.AutoSize = False,
.Height = StartButton.PreferredSize.Height + 20,
.ColumnCount = 8,
.RowCount = 1,
.Padding = New Padding(0, 10, 0, 0)
}
            ' First seven columns auto-size, last column (processCombobox) fills
            For i = 1 To 7
                bottomRow.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            Next
            bottomRow.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))

            processCombobox.Dock = DockStyle.Fill

            bottomRow.Controls.Add(StartButton, 0, 0)
            bottomRow.Controls.Add(StopButton, 1, 0)
            bottomRow.Controls.Add(ClearButton, 2, 0)
            bottomRow.Controls.Add(LoadButton, 3, 0)
            bottomRow.Controls.Add(AudioButton, 4, 0)
            bottomRow.Controls.Add(QuitButton, 5, 0)
            bottomRow.Controls.Add(ProcessButton, 6, 0)
            bottomRow.Controls.Add(processCombobox, 7, 0)

            root.Controls.Add(bottomRow, 0, 2)

            ' Swap in our root layout
            Me.Controls.Clear()
            Me.Controls.Add(root)

            ' Freeze minimum size once first shown
            AddHandler Me.Shown, Sub() Me.MinimumSize = Me.Size

            ' Set icon
            Dim bmp As New Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
            Me.Icon = Icon.FromHandle(bmp.GetHicon())
        End Sub

        ''' <summary>
        ''' Handles form resize events to enforce minimum width based on rightmost control.
        ''' </summary>
        Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles Me.Resize
            Dim minWidth As Integer = SpeakerDistance.Left + SpeakerDistance.Width + 40
            If Me.Width < minWidth Then
                Me.Width = minWidth ' Force minimum width dynamically
            End If
        End Sub

        ''' <summary>
        ''' Stops all active audio recording and transcription engines.
        ''' Gracefully shuts down loopback capture, microphone input, and selected STT engine.
        ''' Releases system sleep lock if this form instance was responsible for setting it.
        ''' </summary>
        Private Async Function StopRecording() As System.Threading.Tasks.Task

            If loopbackCapture IsNot Nothing Then
                RemoveHandler loopbackCapture.DataAvailable, AddressOf OnLoopbackDataAvailable
                loopbackCapture.StopRecording()
                loopbackCapture.Dispose()
                loopbackCapture = Nothing
            End If

            If loopbackResampler IsNot Nothing Then
                loopbackResampler.Dispose()
                loopbackResampler = Nothing
                loopbackRawProvider = Nothing
            End If

            If waveIn IsNot Nothing Then
                RemoveHandler waveIn.DataAvailable, AddressOf OnGoogleDataAvailable
                RemoveHandler waveIn.DataAvailable, AddressOf OnAudioDataAvailable
                waveIn.StopRecording()
                waveIn.Dispose()
                waveIn = Nothing
            End If

            CancelTranscription()

            If STTModel = "google" AndAlso _stream IsNot Nothing Then
                Await SafeCompleteAndDisposeGoogleStreamAsync(readerCts.Token)
            End If

            If WhisperRecognizer IsNot Nothing Then
                PartialTextLabel.Invoke(Sub() PartialTextLabel.Text = "Whisper stopped...")
                Await WhisperRecognizer.DisposeAsync()
                WhisperRecognizer = Nothing
            End If

            ' Only release the sleep lock IF we were the ones who set it
            If _iSetTheSleepLock Then
                ' We are responsible, so we release the lock
                SetThreadExecutionState(ES_CONTINUOUS)
                _iSetTheSleepLock = False
                Debug.WriteLine("This form released the sleep lock.")
            Else
                ' We are not responsible, so we do nothing to the execution state
                Debug.WriteLine("Another component is managing the sleep lock. This form took no action.")
            End If

        End Function

        ''' <summary>
        ''' Handles Stop button click. Disables UI, stops recording asynchronously, and re-enables controls.
        ''' For Vosk, commits any pending partial text before clearing.
        ''' </summary>
        Private Sub StopButton_Click(sender As Object, e As EventArgs)

            If Not capturing Then Return

            STTCanceled = True

            ' Prevent multiple clicks
            Me.StopButton.Enabled = False
            If STTModel <> "vosk" Then
                PartialTextLabel.Text = "Stopping…"
            End If

            System.Threading.Tasks.Task.Run(Async Function()
                                                Try
                                                    Await StopRecording()
                                                    If STTModel = "google" Then StopApiWatchdogTimer()
                                                Catch ex As System.Exception
                                                    ' Silently handle exceptions during cleanup
                                                End Try

                                                Me.Invoke(Sub()
                                                              Me.StartButton.Enabled = True
                                                              Me.LoadButton.Enabled = True
                                                              Me.AudioButton.Enabled = True
                                                              Me.cultureComboBox.Enabled = True
                                                              Me.deviceComboBox.Enabled = True
                                                              Me.SpeakerIdent.Enabled = True
                                                              Me.SpeakerDistance.Enabled = True

                                                              If STTModel = "vosk" Then
                                                                  Addline(PartialTextLabel.Text)
                                                              End If
                                                              PartialTextLabel.Text = String.Empty
                                                          End Sub)
                                            End Function)

            capturing = False

        End Sub

        ''' <summary>
        ''' Handles Clear button click. Clears all transcript text from RichTextBox.
        ''' </summary>
        Private Sub ClearButton_Click(sender As Object, e As EventArgs)
            RichTextBox1.Invoke(Sub()
                                    RichTextBox1.Text = ""
                                    RichTextBox1.SelectionStart = RichTextBox1.Text.Length
                                    RichTextBox1.ScrollToCaret()
                                End Sub)
        End Sub

        ''' <summary>
        ''' Handles form closing event. If transcription is active, stops recording before closing.
        ''' </summary>
        Private Sub FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.Closing

            If e.CloseReason = CloseReason.UserClosing Then
                If capturing Then

                    STTCanceled = True

                    Me.StopButton.Enabled = False
                    Me.AudioButton.Enabled = False
                    Me.QuitButton.Enabled = False
                    If STTModel <> "vosk" Then
                        PartialTextLabel.Text = "Stopping…"
                    End If

                    System.Threading.Tasks.Task.Run(Async Function()
                                                        Try
                                                            Await StopRecording()
                                                            If STTModel = "google" Then StopApiWatchdogTimer()
                                                        Catch ex As System.Exception
                                                            ' Silently handle exceptions during cleanup
                                                        End Try

                                                        Me.Invoke(Sub()
                                                                      Me.StartButton.Enabled = False
                                                                      Me.LoadButton.Enabled = False

                                                                      If STTModel = "vosk" Then
                                                                          Addline(PartialTextLabel.Text)
                                                                      End If
                                                                      PartialTextLabel.Text = String.Empty
                                                                  End Sub)
                                                    End Function)

                    capturing = False

                End If
            End If
        End Sub

        ''' <summary>
        ''' Handles Audio button click. Opens audio output device configuration dialog.
        ''' </summary>
        Private Sub AudioButton_Click(sender As Object, e As EventArgs)
            ConfigureAudioOutputDevice()
        End Sub

        ''' <summary>
        ''' Handles Quit button click. Stops recording if active and closes the form.
        ''' </summary>
        Private Sub QuitButton_Click(sender As Object, e As EventArgs)

            If capturing Then

                STTCanceled = True

                Me.StopButton.Enabled = False
                Me.AudioButton.Enabled = False
                Me.QuitButton.Enabled = False
                If STTModel <> "vosk" Then
                    PartialTextLabel.Text = "Stopping…"
                End If

                System.Threading.Tasks.Task.Run(Async Function()
                                                    Try
                                                        Await StopRecording()
                                                    Catch ex As System.Exception
                                                        ' Silently handle exceptions during cleanup
                                                    End Try

                                                    Me.Invoke(Sub()
                                                                  Me.StartButton.Enabled = False
                                                                  Me.LoadButton.Enabled = False

                                                                  If STTModel = "vosk" Then
                                                                      Addline(PartialTextLabel.Text)
                                                                  End If
                                                                  PartialTextLabel.Text = String.Empty
                                                              End Sub)
                                                End Function)

                capturing = False

            End If
            Me.Close()
        End Sub

        ''' <summary>
        ''' Handles Load button click. Prompts user to select an audio file and transcribes it using selected STT engine.
        ''' Supports WAV, MP3, AAC, M4A, MP4, and WMA formats. For Google STT, offers choice between chunked and streaming modes.
        ''' </summary>
        Private Async Sub LoadButton_Click(sender As Object, e As EventArgs)
            If capturing Then Return

            Dim filepath As String = ""

            DragDropFormLabel = "Supported are audio files (*.wav, *.mp3, *.aac, *.m4a, *.mp4 and *.wma)"
            DragDropFormFilter = "Supported Files|*.wav;*.mp3;*.aac;*.m4a;*.mp4;*.wma|" &
                     "Wave files (*.wav)|*.wav|" &
                     "MP3 files (*.mp3)|*.mp3|" &
                     "AAC files (*.aac, *.m4a, *.mp4)|*.aac;*.m4a;*.mp4|" &
                     "WMA files (*.wma)|*.wma|" &
                     "All files|*.*"

            Using form As New DragDropForm()

                If form.ShowDialog() = DialogResult.OK Then
                    DragDropFormLabel = ""
                    DragDropFormFilter = ""
                    filepath = form.SelectedFilePath
                    If Not File.Exists(filepath) Then
                        ShowCustomMessageBox("The selected file was not found.")
                        Return
                    End If
                Else
                    DragDropFormLabel = ""
                    DragDropFormFilter = ""
                    Return
                End If
            End Using
            DragDropFormLabel = ""
            DragDropFormFilter = ""

            Dim splash As New SLib.SplashScreen($"Loading model...")
            splash.Show()
            splash.Refresh()

            cts = New CancellationTokenSource()
            STTCanceled = False
            audioBuffer.Clear()

            Try
                ' Determine STT model from selected combo box item
                If Me.cultureComboBox.SelectedItem.ToString().StartsWith(GoogleSTT_Desc) Then
                    STTModel = "google"
                ElseIf Me.cultureComboBox.SelectedItem.ToString().StartsWith("ggml") Then
                    STTModel = "whisper"
                Else
                    STTModel = "vosk"
                End If

                Select Case STTModel

                    Case "google"

                        readerCts = New CancellationTokenSource()

                        ' Ask user for language code
                        Dim language As String = ShowSelectionForm("Select the language code you want to transcribe in:", $"{GoogleSTT_Desc}", GoogleSTTsupportedLanguages)

                        language = language.Trim()

                        If String.IsNullOrWhiteSpace(language) OrElse String.Equals(language, "ESC", StringComparison.OrdinalIgnoreCase) Then
                            splash.Close()
                            Return
                        End If

                        If Not GoogleSTTsupportedLanguages.Any(Function(code) code.Trim().Normalize().IndexOf(language, StringComparison.OrdinalIgnoreCase) = 0) Then
                            splash.Close()
                            ShowCustomMessageBox("This language code is not supported. Supported are: " & String.Join(", ", GoogleSTTsupportedLanguages))
                            Return
                        End If

                        GoogleLanguageCode = language

                    Case "vosk"
                        StartVosk()

                    Case "whisper"

                        Dim language As String = ShowCustomInputBox("Enter the language ISO code you want Whisper to transcribe (e.g. en, de, fr, etc.) or go with 'auto':", "Whisper Language Code", True, "auto")

                        language = language.ToLower()

                        If String.IsNullOrWhiteSpace(language) Or language = "esc" Or Not WhisperSupportedLanguages.Contains(language.ToLower()) Then
                            splash.Close()
                            If Not WhisperSupportedLanguages.Contains(language.ToLower()) And language <> "esc" Then
                                ShowCustomMessageBox("This language code is not supported. Supported are: Afrikaans (af), Albanian (sq), Amharic (am), Arabic (ar), Armenian (hy), Assamese (as), Azerbaijani (az), Bashkir (ba), Basque (eu), Belarusian (be), Bengali (bn), Bosnian (bs), Breton (br), Bulgarian (bg), Catalan (ca), Chinese (zh), Croatian (hr), Czech (cs), Danish (da), Dutch (nl), English (en), Estonian (et), Faroese (fo), Finnish (fi), French (fr), Galician (gl), Georgian (ka), German (de), Greek (el), Gujarati (gu), Haitian Creole (ht), Hausa (ha), Hebrew (he), Hindi (hi), Hungarian (hu), Icelandic (is), Indonesian (id), Italian (it), Japanese (ja), Javanese (jv), Kannada (kn), Kazakh (kk), Khmer (km), Kinyarwanda (rw), Kirghiz (ky), Korean (ko), Latvian (lv), Lithuanian (lt), Luxembourgish (lb), Macedonian (mk), Malagasy (mg), Malay (ms), Malayalam (ml), Maltese (mt), Maori (mi), Marathi (mr), Mongolian (mn), Myanmar (my), Nepali (ne), Norwegian (no), Occitan (oc), Pashto (ps), Persian (fa), Polish (pl), Portuguese (pt), Punjabi (pa), Romanian (ro), Russian (ru), Sanskrit (sa), Serbian (sr), Sindhi (sd), Sinhala (si), Slovak (sk), Slovenian (sl), Somali (so), Spanish (es), Sundanese (su), Swahili (sw), Swedish (sv), Tagalog (tl), Tajik (tg), Tamil (ta), Tatar (tt), Telugu (te), Thai (th), Turkish (tr), Ukrainian (uk), Urdu (ur), Uzbek (uz), Vietnamese (vi), Welsh (cy), Yiddish (yi), Yoruba (yo), Zulu (zu)")
                            End If
                            STTCanceled = True
                            Return
                        End If

                        StartWhisper(language)
                        STTCanceled = False
                        PartialTextLabel.Invoke(Sub() PartialTextLabel.Text = "Whisper is listening and working... (no partial results shown, please wait)")

                    Case Else
                        splash.Close()
                        ShowCustomMessageBox($"No valid model selected. Please select a model.")
                        Return

                End Select

                ' Save user preferences
                My.Settings.LastAudioSource = Me.deviceComboBox.SelectedItem.ToString()
                My.Settings.LastSpeechModel = Me.cultureComboBox.SelectedItem.ToString()
                My.Settings.LastSpeakerEnabled = Me.SpeakerIdent.Checked
                similarityThreshold = Double.Parse(Me.SpeakerDistance.Text)
                If STTModel = "google" Then
                    If similarityThreshold < 1 Then similarityThreshold = 1.0
                Else
                    If similarityThreshold = 0 Then similarityThreshold = 1.0
                    If similarityThreshold < 0.2 Then similarityThreshold = 0.2
                    If similarityThreshold > 2.5 Then similarityThreshold = 2.5
                End If
                My.Settings.LastSpeakerDistance = similarityThreshold

                My.Settings.Save()

                capturing = True
                Me.StartButton.Enabled = False
                Me.cultureComboBox.Enabled = False
                Me.deviceComboBox.Enabled = False
                Me.SpeakerIdent.Enabled = False
                Me.SpeakerDistance.Enabled = False
                Me.StopButton.Enabled = True
                Me.LoadButton.Enabled = False
                Me.AudioButton.Enabled = False
                splash.Close()

                Select Case STTModel
                    Case "google"
                        googleTranscriptStart = RichTextBox1.TextLength
                        Dim methodChoice As Integer = ShowCustomYesNoBox("Select your Google transcription method (you may have to try which one works better):", "Send chunks (faster)", "Stream (less gaps)")

                        Debug.WriteLine("Choice = " & methodChoice)

                        If methodChoice = 0 Then
                            splash.Close()
                            Return
                        End If

                        ' Close splash, UI is already disabled
                        splash.Close()

                        splash = New SLib.SplashScreen($"Transcribing file ...")
                        splash.Show()
                        splash.Refresh()

                        Try

                            ' Call chunking vs. streaming method
                            If methodChoice = 1 Then
                                Await GoogleChunkedTranscribeAudioFile(filepath)
                            Else
                                Await GoogleFileStreamTranscription(filepath)
                            End If

                        Catch ex As Exception
                            splash.Close()
                            ShowCustomMessageBox($"Error in Transcribing File using Google: {ex.Message}")
                        Finally
                            splash.Close()
                            Me.Invoke(Sub()
                                          capturing = False
                                          StartButton.Enabled = True
                                          StopButton.Enabled = False
                                          LoadButton.Enabled = True
                                          AudioButton.Enabled = True
                                          cultureComboBox.Enabled = True
                                          deviceComboBox.Enabled = True
                                          SpeakerIdent.Enabled = True
                                          SpeakerDistance.Enabled = True
                                      End Sub)
                        End Try

                    Case "vosk"
                        VoskTranscribeAudioFile(filepath)
                    Case "whisper"
                        WhisperTranscribeAudioFile(filepath)
                        ShowCustomMessageBox($"Transcription using Whisper has started In the background. You can continue working. Do not quit Word. Press 'Stop' to stop transcription.")
                End Select

            Catch ex As Exception
                splash.Close()
                ShowCustomMessageBox($"There has been an Error starting the transcription engine (Error: {ex.Message}).")

            End Try

        End Sub

        ''' <summary>
        ''' Handles Process button click. Applies selected LLM prompt to transcript text (selection or entire document).
        ''' Inserts processed result into Word document with markdown formatting.
        ''' </summary>
        Private Async Sub ProcessButton_Click(sender As Object, e As EventArgs)
            If processCombobox.SelectedIndex >= 0 Then
                Dim selectedIndex As Integer = processCombobox.SelectedIndex
                If selectedIndex < TranscriptPromptsLibrary.Count Then
                    Dim OtherPrompt As String = TranscriptPromptsLibrary(selectedIndex)
                    Dim SelectedText As String = ""
                    If String.IsNullOrWhiteSpace(RichTextBox1.SelectedText) Then
                        SelectedText = RichTextBox1.Text
                    Else
                        SelectedText = RichTextBox1.SelectedText
                    End If

                    Dim LLMResult As String = Await LLM(OtherPrompt & " (Current Date: " & DateTime.Now.ToString("dd MMM yyyy", CultureInfo.GetCultureInfo("en-US")) & ")", SelectedText, "", "", 0, False)

                    Dim wordApp As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application
                    Dim selection As Microsoft.Office.Interop.Word.Selection = wordApp.Selection

                    If wordApp.Documents.Count > 0 Then
                        ' Collapse any existing selection towards the end
                        selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                        ' Insert the markdown text
                        InsertTextWithMarkdown(selection, LLMResult, True)
                    End If
                End If
            End If
        End Sub

        ''' <summary>
        ''' Enumerates available audio input devices and populates device combo box.
        ''' Each device appears twice: once for microphone only, once for mic + system audio.
        ''' Pre-selects last used device from settings.
        ''' </summary>
        Private Sub LoadAudioDevices()
            deviceComboBox.Items.Clear()
            Dim i As Integer = 0
            For i = 0 To WaveInEvent.DeviceCount - 1
                Dim capabilities = WaveInEvent.GetCapabilities(i)
                Dim micName As String = $"{i}: {capabilities.ProductName}"
                ' a) plain mic
                deviceComboBox.Items.Add(micName)
                ' b) mic + system audio
                deviceComboBox.Items.Add($"{micName} (plus audio output)")
            Next

            ' Select default device (if available)
            Dim lastAudioSource As String = My.Settings.LastAudioSource
            If Not String.IsNullOrEmpty(lastAudioSource) AndAlso deviceComboBox.Items.Contains(lastAudioSource) Then
                deviceComboBox.SelectedItem = lastAudioSource
            ElseIf deviceComboBox.Items.Count > 0 Then
                deviceComboBox.SelectedIndex = 0
            End If
            Dim sel = TryCast(deviceComboBox.SelectedItem, String)
            _multiSourceSelected = (sel IsNot Nothing AndAlso sel.EndsWith("(plus audio output)"))
        End Sub

        ''' <summary>
        ''' Initializes Vosk speech recognizer with selected model and optional speaker recognition.
        ''' Configures recognizer for optimal real-time performance with word timestamps and partial results.
        ''' </summary>
        Private Sub StartVosk()
            Dim modelpath As String = System.IO.Path.Combine(ExpandEnvironmentVariables(Globals.ThisAddIn.INI_SpeechModelPath), Me.cultureComboBox.SelectedItem.ToString())
            Dim model As New Model(modelpath)
            recognizer = New VoskRecognizer(model, 16000.0F)
            If Me.SpeakerIdent.Checked Then
                ' Get the first available speaker model in the directory
                Dim speakerModelPath As String = System.IO.Directory.GetDirectories(System.IO.Path.Combine(ExpandEnvironmentVariables(Globals.ThisAddIn.INI_SpeechModelPath), "Speaker\"), "vosk-model*").FirstOrDefault()
                If String.IsNullOrEmpty(speakerModelPath) Then
                    ShowCustomMessageBox($"No speaker model found (at {System.IO.Path.Combine(ExpandEnvironmentVariables(Globals.ThisAddIn.INI_SpeechModelPath), "Speaker\")}. Speaker recognition will be disabled.")
                    Me.SpeakerIdent.Checked = False
                Else
                    Dim speakerModel As SpkModel = New SpkModel(speakerModelPath)
                    recognizer.SetSpkModel(speakerModel)
                End If

                Debug.WriteLine("Vosk recognizer initialized")
            End If

            recognizer.SetMaxAlternatives(0) ' Forces earlier finalization
            recognizer.SetWords(True) ' Enable word timestamps
            recognizer.SetPartialWords(True) ' Partial words emitted faster
        End Sub

        ''' <summary>
        ''' Initializes Whisper speech processor with selected GGML model and optional translation.
        ''' Configures processor with language, thread count, voice detection threshold, and temperature settings.
        ''' </summary>
        ''' <param name="language">ISO 639-1 language code or "auto" for automatic detection</param>
        Private Sub StartWhisper(Optional language As String = "auto")
            Dim modelpath As String = System.IO.Path.Combine(ExpandEnvironmentVariables(Globals.ThisAddIn.INI_SpeechModelPath), Me.cultureComboBox.SelectedItem.ToString())

            ' Load the model using WhisperFactory with the specified runtime options
            Dim factory As WhisperFactory = WhisperFactory.FromPath(modelpath)

            ' Configure the builder with language, threads, etc.
            If Me.SpeakerIdent.Checked Then
                Dim builder = factory.CreateBuilder() _
            .WithLanguage(language) _
            .WithThreads(Environment.ProcessorCount) _
            .WithNoSpeechThreshold(Single.Parse(Me.SpeakerDistance.Text)) _
            .WithTemperature(0.3) _
            .WithTranslate()

                ' Build the recognizer
                WhisperRecognizer = builder.Build()
            Else
                Dim builder = factory.CreateBuilder() _
            .WithLanguage(language) _
            .WithThreads(Environment.ProcessorCount) _
            .WithNoSpeechThreshold(Single.Parse(Me.SpeakerDistance.Text)) _
            .WithTemperature(0.3)

                ' Build the recognizer
                WhisperRecognizer = builder.Build()
            End If
        End Sub

        ''' <summary>
        ''' Initializes Google Cloud Speech-to-Text streaming session with OAuth2 authentication.
        ''' Creates bidirectional gRPC stream with token refresh interceptor and initializes writer task.
        ''' </summary>
        Private Async Function StartGoogleSTT() As System.Threading.Tasks.Task
            ' 1) Define interceptor that fetches a fresh token for each new streaming call

            Dim callCreds As Grpc.Core.CallCredentials = Grpc.Core.CallCredentials.FromInterceptor(
            Async Function(contextCall, metadata)
                ' Use our local helper instead of context.GetFresh...
                Dim tokenToSend As String = Await GetFreshSTTToken(STTSecondAPI)
                metadata.Add("Authorization", $"Bearer {tokenToSend}")
                Await System.Threading.Tasks.Task.CompletedTask
            End Function
        )

            ' 2) Build ChannelCredentials with Secure SSL + our interceptor
            Dim channelCreds As Grpc.Core.ChannelCredentials = Grpc.Core.ChannelCredentials.Create(
                    Grpc.Core.ChannelCredentials.SecureSsl,
                    callCreds
                )

            ' 3) Create a brand new SpeechClient using the above channelCreds
            Dim builder As New Google.Cloud.Speech.V1.SpeechClientBuilder() With {
                    .Endpoint = STTEndpoint,
                    .ChannelCredentials = channelCreds
                }
            client = builder.Build()

            ' 4) Open streaming connection with InitializeGoogleStream()
            ' This calls "_stream = client.StreamingRecognize() ..." in the background and sends
            ' StreamingConfig via WriteAsync. The interceptor becomes active on first WriteAsync.
            Await InitializeGoogleStream()

            SyncLock ringBuffer
                ringBuffer.Clear()
            End SyncLock

            StartAudioQueueWriter()

        End Function

        ''' <summary>
        ''' Resets the Google stream completion flag to allow new streaming sessions.
        ''' </summary>
        Private Sub ResetGoogleStreamFlag()
            _googleStreamCompleted = False
        End Sub

        ''' <summary>
        ''' Initializes bidirectional Google STT streaming with recognition configuration.
        ''' Configures audio encoding, language, punctuation, and optional speaker diarization.
        ''' Falls back to basic config if diarization fails for selected language.
        ''' </summary>
        Private Async Function InitializeGoogleStream() As System.Threading.Tasks.Task

            streamingStartTime = DateTime.UtcNow
            ResetGoogleStreamFlag()

            Try
                ' Open bidirectional streaming

                If Me.SpeakerIdent.Checked Then

                    Dim minSpeakers As Integer = 2
                    Dim maxSpeakers As Integer = 6 ' Default maximum, adjust if needed

                    ' Try to read values from UI with safe defaults
                    Try
                        ' Assumption: SpeakerDistance is now MaxCount and a new TextField is MinCount
                        maxSpeakers = CInt(Double.Parse(Me.SpeakerDistance.Text))
                    Catch
                        ' Use default values on error
                    End Try

                    ' Limit values to Google-supported range
                    minSpeakers = System.Math.Max(2, minSpeakers)
                    maxSpeakers = System.Math.Max(minSpeakers, maxSpeakers)

                    _stream = client.StreamingRecognize()
                    Dim streamingConfig As New StreamingRecognitionConfig With {
                .Config = New RecognitionConfig With {
                    .Encoding = RecognitionConfig.Types.AudioEncoding.Linear16,
                    .SampleRateHertz = 16000,
                    .LanguageCode = GoogleLanguageCode,
                    .EnableAutomaticPunctuation = True,
                    .EnableSpokenPunctuation = True,
                    .EnableWordTimeOffsets = False,
                    .EnableWordConfidence = False,
                    .Model = "latest_long",
                    .UseEnhanced = True,
                    .DiarizationConfig = New SpeakerDiarizationConfig With {
                        .EnableSpeakerDiarization = Me.SpeakerIdent.Checked,
                                .MinSpeakerCount = minSpeakers,
                            .MaxSpeakerCount = maxSpeakers
                                                    }
                },
                .InterimResults = True,
                .SingleUtterance = False
            }
                    Await _stream.WriteAsync(New StreamingRecognizeRequest With {.StreamingConfig = streamingConfig})

                Else
                    _stream = client.StreamingRecognize()
                    Dim streamingConfig As New StreamingRecognitionConfig With {
            .Config = New RecognitionConfig With {
                .Encoding = RecognitionConfig.Types.AudioEncoding.Linear16,
                .SampleRateHertz = 16000,
                .LanguageCode = GoogleLanguageCode,
                .EnableAutomaticPunctuation = True,
                    .EnableSpokenPunctuation = True,
                    .EnableWordTimeOffsets = False,
                    .EnableWordConfidence = False,
                    .Model = "latest_long",
                    .UseEnhanced = True
                        },
                        .InterimResults = True,
                        .SingleUtterance = False
                    }
                    Await _stream.WriteAsync(New StreamingRecognizeRequest With {.StreamingConfig = streamingConfig})
                End If

            Catch ex As System.Exception

                ShowCustomMessageBox("No speaker diarization available for this language (or other error).", $"{GoogleSTT_Desc} Language Code")
                _stream = client.StreamingRecognize()
                Dim streamingConfig As New StreamingRecognitionConfig With {
        .Config = New RecognitionConfig With {
            .Encoding = RecognitionConfig.Types.AudioEncoding.Linear16,
            .SampleRateHertz = 16000,
            .LanguageCode = GoogleLanguageCode
                    },
                    .InterimResults = True
                }
                _stream.WriteAsync(New StreamingRecognizeRequest With {.StreamingConfig = streamingConfig}).Wait()

            End Try

        End Function

        ''' <summary>
        ''' Handles Start button click for real-time transcription.
        ''' Loads selected STT model, prompts for language (if needed), starts recording, and begins transcription.
        ''' Validates language codes and prevents duplicate starts. Enables watchdog timer for Google STT.
        ''' </summary>
        Private Async Sub StartButton_Click(sender As Object, e As EventArgs)

            If capturing Then
                Return
            End If

            Dim splash As New SLib.SplashScreen($"Loading model...")
            splash.Show()
            splash.Refresh()

            cts = New CancellationTokenSource()
            STTCanceled = False
            audioBuffer.Clear()

            Try
                ' Determine STT model from selected combo box item
                If Me.cultureComboBox.SelectedItem.ToString().StartsWith(GoogleSTT_Desc) Then
                    STTModel = "google"
                ElseIf Me.cultureComboBox.SelectedItem.ToString().StartsWith("ggml") Then
                    STTModel = "whisper"
                Else
                    STTModel = "vosk"
                End If

                Select Case STTModel

                    Case "google"

                        readerCts = New CancellationTokenSource()

                        Dim language As String = ShowSelectionForm("Select the language code you want to transcribe in:", $"{GoogleSTT_Desc}", GoogleSTTsupportedLanguages)

                        language = language.Trim()

                        ' First handle empty or escape
                        If String.IsNullOrWhiteSpace(language) OrElse String.Equals(language, "esc", StringComparison.OrdinalIgnoreCase) Then
                            splash.Close()
                            STTCanceled = True
                            Return
                        End If

                        ' Now do a true case-insensitive lookup
                        If Not GoogleSTTsupportedLanguages.Any(
                        Function(code)
                            Return String.Equals(code, language, StringComparison.OrdinalIgnoreCase)
                        End Function) Then
                            splash.Close()
                            ShowCustomMessageBox("This language code is not supported. Supported are: " & String.Join(", ", GoogleSTTsupportedLanguages), $"{GoogleSTT_Desc} Language Code")
                            STTCanceled = True
                            Return
                        End If

                        Try
                            GoogleLanguageCode = language
                            Await StartGoogleSTT()
                        Catch ex As System.Exception
                            ShowCustomMessageBox("Error starting transcription service: {ex.Message}", $"{GoogleSTT_Desc}")
                            STTCanceled = True
                            Return
                        End Try

                        If Not StartRecording() Then
                            splash.Close()
                            Return
                        End If

                        googleTranscriptStart = RichTextBox1.TextLength

                        _speakerTagToLabelMap.Clear()
                        _nextSpeakerNumber = 1

                        Me.googleReaderTask = StartGoogleReaderTask()

                    Case "vosk"

                        StartVosk()

                        If Not StartRecording() Then
                            splash.Close()
                            Return
                        End If

                    Case "whisper"

                        Dim language As String = ShowCustomInputBox("Enter the language ISO code you want Whisper to transcribe (e.g. en, de, fr, etc.) or go with 'auto':", "Whisper Language Code", True, "auto")

                        language = language.ToLower()

                        If String.IsNullOrWhiteSpace(language) Or language = "esc" Or Not WhisperSupportedLanguages.Contains(language.ToLower()) Then
                            splash.Close()
                            If Not WhisperSupportedLanguages.Contains(language.ToLower()) And language <> "esc" Then
                                ShowCustomMessageBox("This language code is not supported. Supported are: Afrikaans (af), Albanian (sq), Amharic (am), Arabic (ar), Armenian (hy), Assamese (as), Azerbaijani (az), Bashkir (ba), Basque (eu), Belarusian (be), Bengali (bn), Bosnian (bs), Breton (br), Bulgarian (bg), Catalan (ca), Chinese (zh), Croatian (hr), Czech (cs), Danish (da), Dutch (nl), English (en), Estonian (et), Faroese (fo), Finnish (fi), French (fr), Galician (gl), Georgian (ka), German (de), Greek (el), Gujarati (gu), Haitian Creole (ht), Hausa (ha), Hebrew (he), Hindi (hi), Hungarian (hu), Icelandic (is), Indonesian (id), Italian (it), Japanese (ja), Javanese (jv), Kannada (kn), Kazakh (kk), Khmer (km), Kinyarwanda (rw), Kirghiz (ky), Korean (ko), Latvian (lv), Lithuanian (lt), Luxembourgish (lb), Macedonian (mk), Malagasy (mg), Malay (ms), Malayalam (ml), Maltese (mt), Maori (mi), Marathi (mr), Mongolian (mn), Myanmar (my), Nepali (ne), Norwegian (no), Occitan (oc), Pashto (ps), Persian (fa), Polish (pl), Portuguese (pt), Punjabi (pa), Romanian (ro), Russian (ru), Sanskrit (sa), Serbian (sr), Sindhi (sd), Sinhala (si), Slovak (sk), Slovenian (sl), Somali (so), Spanish (es), Sundanese (su), Swahili (sw), Swedish (sv), Tagalog (tl), Tajik (tg), Tamil (ta), Tatar (tt), Telugu (te), Thai (th), Turkish (tr), Ukrainian (uk), Urdu (ur), Uzbek (uz), Vietnamese (vi), Welsh (cy), Yiddish (yi), Yoruba (yo), Zulu (zu)")
                            End If
                            STTCanceled = True
                            Return
                        End If
                        StartWhisper(language)

                        If Not StartRecording() Then
                            splash.Close()
                            STTCanceled = True
                            Return
                        End If
                        STTCanceled = False

                        PartialTextLabel.Invoke(Sub() PartialTextLabel.Text = "Whisper is listening and working... (no partial results shown, please wait)")
                    Case Else
                        splash.Close()
                        ShowCustomMessageBox($"No valid model selected. Please select a model.")
                        Return

                End Select

                ' Save user preferences
                My.Settings.LastAudioSource = Me.deviceComboBox.SelectedItem.ToString()
                My.Settings.LastSpeechModel = Me.cultureComboBox.SelectedItem.ToString()
                My.Settings.LastSpeakerEnabled = Me.SpeakerIdent.Checked
                similarityThreshold = Double.Parse(Me.SpeakerDistance.Text)
                If STTModel = "google" Then
                    If similarityThreshold < 1 Then similarityThreshold = 1.0
                Else
                    If similarityThreshold = 0 Then similarityThreshold = 1.0
                    If similarityThreshold < 0.2 Then similarityThreshold = 0.2
                    If similarityThreshold > 2.5 Then similarityThreshold = 2.5
                End If
                My.Settings.LastSpeakerDistance = similarityThreshold

                My.Settings.Save()

                If STTModel = "google" Then StartApiWatchdogTimer()

                capturing = True
                Me.StartButton.Enabled = False
                Me.cultureComboBox.Enabled = False
                Me.deviceComboBox.Enabled = False
                Me.SpeakerIdent.Enabled = False
                Me.SpeakerDistance.Enabled = False
                Me.StopButton.Enabled = True
                Me.LoadButton.Enabled = False
                Me.AudioButton.Enabled = False
                splash.Close()

            Catch ex As Exception
                splash.Close()
                ShowCustomMessageBox($"There has been an error starting the transcription engine (Error: {ex.Message}).")

            End Try
        End Sub

        ''' <summary>
        ''' Starts audio recording from selected microphone with optional system audio mixing.
        ''' Configures WaveInEvent and optional WasapiLoopbackCapture with resampling.
        ''' Prevents system sleep during recording and assigns appropriate event handlers based on STT engine.
        ''' </summary>
        ''' <returns>True if recording started successfully, False if device selection is invalid</returns>
        Private Function StartRecording() As Boolean

            Dim ss As String = TryCast(Me.deviceComboBox.SelectedItem, String)
            Dim deviceIndex As Integer

            ' Parse device index from selection string (format: "0: Device Name")
            Dim pos As Integer = If(ss?.IndexOf(":"c), -1)
            If pos < 0 OrElse Not Integer.TryParse(ss.Substring(0, pos), deviceIndex) Then
                ShowCustomMessageBox($"Invalid device selection: '{ss}'")
                Return False
            End If

            ' Configure microphone capture with 16 kHz mono format (required by all STT engines)
            waveIn = New WaveInEvent() With {
            .DeviceNumber = deviceIndex,
            .WaveFormat = New WaveFormat(16000, 1)
        }

            ' Set up system audio loopback capture if multi-source is enabled
            If MultiSourceEnabled Then

                ' Attempt to use the output device configured in settings
                Dim audioOutputDeviceId As String = My.Settings.AudioOutputDevice
                Dim chosenDevice As MMDevice = Nothing

                If Not String.IsNullOrEmpty(audioOutputDeviceId) Then
                    Try
                        Dim enumerator As New MMDeviceEnumerator()
                        chosenDevice = enumerator.GetDevice(audioOutputDeviceId)
                    Catch ex As System.Exception
                        ' Invalid ID or device not found → fallback to default
                        chosenDevice = Nothing
                    End Try
                End If

                ' 1) Create LoopbackCapture with specific device or default
                If chosenDevice IsNot Nothing Then
                    loopbackCapture = New WasapiLoopbackCapture(chosenDevice)
                Else
                    loopbackCapture = New WasapiLoopbackCapture()
                End If

                ' 2) Raw provider in native format (will be resampled later)
                loopbackRawProvider = New BufferedWaveProvider(loopbackCapture.WaveFormat) With {
                        .DiscardOnBufferOverflow = True
                    }
                AddHandler loopbackCapture.DataAvailable, Sub(s, ev)
                                                              loopbackRawProvider.AddSamples(ev.Buffer, 0, ev.BytesRecorded)
                                                          End Sub

                ' 3) Resample from native format → microphone format (16 kHz mono 16-bit)
                loopbackResampler = New MediaFoundationResampler(loopbackRawProvider, waveIn.WaveFormat) With {
                    .ResamplerQuality = 60
                }

                ' 4) Start loopback recording
                Try
                    loopbackCapture.StartRecording()
                Catch ex As System.Exception
                    ' Device possibly in exclusive use → fallback to mic-only mode
                    ShowCustomMessageBox("Cannot capture system audio: Device is in exclusive use or invalid. Continuing with mic only.")
                    loopbackCapture.Dispose()
                    loopbackCapture = Nothing
                    loopbackResampler?.Dispose()
                    loopbackResampler = Nothing
                    loopbackRawProvider = Nothing
                End Try
            End If

            ' Assign appropriate data handler based on selected STT engine
            If STTModel = "google" Then
                AddHandler waveIn.DataAvailable, AddressOf OnGoogleDataAvailable
            Else
                AddHandler waveIn.DataAvailable, AddressOf OnAudioDataAvailable
            End If
            waveIn.StartRecording()

            ' Request system stay awake during transcription
            ' The function returns the PREVIOUS execution state
            Dim previousState As UInteger = SetThreadExecutionState(ES_CONTINUOUS Or ES_SYSTEM_REQUIRED)

            ' Check if the SYSTEM_REQUIRED flag was already set before our call
            ' Using bitwise AND: if result is 0, the flag was NOT set
            If (previousState And ES_SYSTEM_REQUIRED) = 0 Then
                ' Lock was NOT active before, therefore WE are responsible for releasing it
                _iSetTheSleepLock = True
                Debug.WriteLine("Sleep lock was not active. This form has now acquired it.")
            Else
                ' Lock was ALREADY active, we are not responsible for releasing it
                _iSetTheSleepLock = False
                Debug.WriteLine("Sleep lock was already active. This form will not release it.")
            End If

            Return True

        End Function

        ''' <summary>
        ''' Starts the Google STT API watchdog timer that monitors for unresponsive streams.
        ''' Timer fires every 1 second and triggers recovery if no API response received within timeout period.
        ''' </summary>
        Private Sub StartApiWatchdogTimer()
            ' Initialize the last response time to now
            System.Threading.Interlocked.Exchange(_lastApiResponseTicks, DateTime.UtcNow.Ticks)

            ' Dispose of any existing timer to prevent orphans
            _apiWatchdogTimer?.Dispose()

            ' Create a new timer that calls CheckApiResponse every 1000ms (1 second)
            _apiWatchdogTimer = New System.Threading.Timer(
                                AddressOf CheckApiResponse,
                                Nothing,
                                TimeSpan.FromSeconds(1),
                                TimeSpan.FromSeconds(1)
                            )
        End Sub

        ''' <summary>
        ''' Stops and disposes the Google STT API watchdog timer.
        ''' </summary>
        Private Sub StopApiWatchdogTimer()
            _apiWatchdogTimer?.Dispose()
            _apiWatchdogTimer = Nothing
        End Sub

        ''' <summary>
        ''' Watchdog timer callback that detects Google API unresponsiveness and triggers recovery.
        ''' Fires every 1 second and checks if API has responded within the configured timeout period.
        ''' If timeout exceeded, stops timer and initiates stream recovery.
        ''' </summary>
        ''' <param name="state">Timer state (unused)</param>
        Private Sub CheckApiResponse(state As Object)
            ' If not capturing or recovery already in progress, do nothing
            If Not capturing OrElse recoverySemaphore.CurrentCount = 0 Then
                Return
            End If

            ' Atomically read the last response time
            Dim lastResponseTime As New DateTime(System.Threading.Interlocked.Read(_lastApiResponseTicks))

            ' Check if elapsed time has exceeded our timeout
            If (DateTime.UtcNow - lastResponseTime).TotalSeconds > API_RESPONSE_TIMEOUT_SECONDS Then
                ' API has not responded in time - stream is likely hung
                Debug.WriteLine($"[ApiWatchdog] No API response for >{API_RESPONSE_TIMEOUT_SECONDS}s. Forcing stream recovery.")

                ' Stop the timer to prevent re-triggering during recovery
                StopApiWatchdogTimer()

                ' Use existing thread-safe recovery method to restart the stream
                ' Watchdog will be restarted after recovery completes
                System.Threading.Tasks.Task.Run(Async Sub() Await TryRecoverGoogleStreamAsync())
            End If
        End Sub

        ''' <summary>
        ''' Starts Google STT streaming reader task that processes transcription responses.
        ''' Handles both final and interim results, applies speaker diarization when enabled,
        ''' and manages duplicate detection after stream recovery.
        ''' Updates watchdog timestamp on every API response.
        ''' </summary>
        ''' <returns>Background task that reads from Google's response stream</returns>
        Private Function StartGoogleReaderTask() As System.Threading.Tasks.Task
            ' Cancel any existing reader task
            If readerCts IsNot Nothing Then
                Try
                    readerCts.Cancel()
                Catch
                    ' Silently ignore cancellation errors
                End Try
            End If

            readerCts = New CancellationTokenSource()

            Dim newTask = System.Threading.Tasks.Task.Run(
                            Async Sub()
                                Dim token = readerCts.Token
                                Try
                                    Dim enumerator = _stream.GetResponseStream().GetAsyncEnumerator(token)

                                    While Await enumerator.MoveNextAsync()
                                        ' Update watchdog timestamp for every API response
                                        System.Threading.Interlocked.Exchange(_lastApiResponseTicks, DateTime.UtcNow.Ticks)

                                        For Each result In enumerator.Current.Results
                                            If result.IsFinal Then
                                                ' Process final transcription result
                                                If result.Alternatives.Count > 0 Then
                                                    Dim bestAlternative = result.Alternatives(0)
                                                    Dim finalTranscript As String = bestAlternative.Transcript.Trim()
                                                    _lastKnownPartialResult = ""

                                                    ' Check if this is a duplicate of a partial result committed during recovery
                                                    If Not String.IsNullOrEmpty(_justCommittedPartialText) AndAlso
                                                       String.Equals(finalTranscript, _justCommittedPartialText.Trim(), StringComparison.OrdinalIgnoreCase) Then

                                                        ' This is a duplicate - ignore it
                                                        Debug.WriteLine($"[ReaderTask] Ignoring duplicate final result: '{finalTranscript}'")
                                                        _justCommittedPartialText = ""

                                                    Else
                                                        ' This is a new, valid final result
                                                        _justCommittedPartialText = ""

                                                        ' Apply speaker diarization formatting if enabled
                                                        If Me.SpeakerIdent.Checked AndAlso bestAlternative.Words.Count > 0 Then

                                                            Dim currentSegment As New System.Text.StringBuilder()
                                                            ' Get label for the first word's speaker
                                                            Dim currentSpeakerLabel As String = GetSpeakerLabel(bestAlternative.Words(0).SpeakerTag)

                                                            For Each wordInfo In bestAlternative.Words
                                                                Dim wordSpeakerLabel As String = GetSpeakerLabel(wordInfo.SpeakerTag)

                                                                If wordSpeakerLabel <> currentSpeakerLabel Then
                                                                    ' Speaker changed - commit previous speaker's segment
                                                                    Dim segmentToCommit As String = $"{currentSpeakerLabel}: {currentSegment.ToString().Trim()}"
                                                                    Addline(segmentToCommit)

                                                                    ' Start new segment for new speaker
                                                                    currentSegment.Clear()
                                                                    currentSpeakerLabel = wordSpeakerLabel
                                                                End If

                                                                ' Append current word to segment
                                                                currentSegment.Append(wordInfo.Word & " ")
                                                            Next

                                                            ' Commit final segment after loop
                                                            If currentSegment.Length > 0 Then
                                                                Dim finalSegmentToCommit As String = $"{currentSpeakerLabel}: {currentSegment.ToString().Trim()}"
                                                                Addline(finalSegmentToCommit)
                                                            End If
                                                        Else
                                                            ' Standard non-diarization logic
                                                            Addline(finalTranscript)
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                ' Process interim (partial) result
                                                If result.Alternatives.Count > 0 Then
                                                    Dim partialTranscript = result.Alternatives(0).Transcript
                                                    PartialTextLabel.Invoke(Sub() PartialTextLabel.Text = partialTranscript)
                                                    _lastKnownPartialResult = partialTranscript
                                                End If
                                            End If
                                        Next
                                    End While

                                Catch ex As OperationCanceledException
                                    Debug.WriteLine($"[ReaderTask] Gracefully cancelled via OperationCanceledException.")
                                Catch rex As RpcException
                                    If token.IsCancellationRequested OrElse rex.StatusCode = StatusCode.Cancelled Then
                                        Debug.WriteLine($"[ReaderTask] Gracefully cancelled via RpcException (Status: {rex.StatusCode}).")
                                    Else
                                        Debug.WriteLine($"[ReaderTask] Unexpected RpcException (Status: {rex.StatusCode}). Requesting recovery...")
                                        System.Threading.Tasks.Task.Run(Async Sub() Await TryRecoverGoogleStreamAsync())
                                    End If
                                Catch ex As Exception
                                    Debug.WriteLine($"[ReaderTask] UNEXPECTED FATAL ERROR: {ex.ToString()}")
                                End Try
                            End Sub)
            Return newTask
        End Function

        ''' <summary>
        ''' Maps temporary Google speaker tags to consistent human-readable labels (Speaker 1, Speaker 2, etc.).
        ''' Maintains session-wide speaker identity across stream recoveries.
        ''' Automatically assigns new speaker numbers for previously unseen tags.
        ''' </summary>
        ''' <param name="speakerTag">Temporary speaker tag from Google API</param>
        ''' <returns>Consistent speaker label for display</returns>
        Private Function GetSpeakerLabel(speakerTag As Integer) As String
            ' Check if we've already seen this tag in this session
            If _speakerTagToLabelMap.ContainsKey(speakerTag) Then
                ' Return the consistent label we already assigned
                Return _speakerTagToLabelMap(speakerTag)
            Else
                ' This is a new speaker tag - assign it a new label
                Dim newLabel As String = $"Speaker {_nextSpeakerNumber}"
                _nextSpeakerNumber += 1

                ' Store the mapping for future use
                _speakerTagToLabelMap.Add(speakerTag, newLabel)

                Return newLabel
            End If
        End Function

        ''' <summary>
        ''' Thread-safe wrapper for Google STT stream recovery.
        ''' Commits any pending partial results before recovery, acquires recovery semaphore to prevent
        ''' concurrent recovery attempts, and restarts watchdog timer after successful recovery.
        ''' </summary>
        Private Async Function TryRecoverGoogleStreamAsync() As System.Threading.Tasks.Task

            ' Check if there is a pending partial result that needs to be committed
            If Not String.IsNullOrWhiteSpace(_lastKnownPartialResult) Then
                ' Create a copy to avoid race conditions
                Dim partialToCommit As String = _lastKnownPartialResult

                ' Reset the class-level variable immediately
                _lastKnownPartialResult = ""
                _justCommittedPartialText = partialToCommit

                ' Use thread-safe Addline method to append to RichTextBox
                Debug.WriteLine($"[TryRecover] Committing lost partial result: '{partialToCommit}'")
                Addline(partialToCommit)
            End If

            ' Asynchronously wait to acquire the semaphore
            ' If another thread already has it, this thread waits without blocking thread-pool
            Await recoverySemaphore.WaitAsync()
            Try
                ' Now that we have the lock, perform actual recovery
                ' Any other threads calling this method will wait at WaitAsync above
                Debug.WriteLine($"[TryRecover] Acquired semaphore. Starting recovery... ts={DateTime.UtcNow:HH:mm:ss.fff}")
                Await RecoverGoogleStream()
                streamingStartTime = DateTime.UtcNow ' Reset timer after successful recovery
                Me.Invoke(Sub() StartApiWatchdogTimer())
            Finally
                ' CRITICAL: Always release semaphore in Finally block to prevent deadlocks
                recoverySemaphore.Release()
                Debug.WriteLine($"[TryRecover] Released semaphore. ts={DateTime.UtcNow:HH:mm:ss.fff}")
            End Try
        End Function

        ''' <summary>
        ''' Starts background writer task that consumes audio chunks from queue and sends to Google STT stream.
        ''' Handles graceful shutdown when queue is completed and catches expected exceptions during recovery.
        ''' Exits cleanly on stream disposal or cancellation.
        ''' </summary>
        Private Sub StartAudioQueueWriter()
            writerTask = System.Threading.Tasks.Task.Run(
                                Async Sub()
                                    Try
                                        ' Loop automatically exits when queue is completed by SafeCompleteAndDisposeGoogleStreamAsync
                                        For Each chunk As Google.Protobuf.ByteString In audioQueue.GetConsumingEnumerable()
                                            Try
                                                ' If stream was disposed during recovery, exit immediately
                                                If _stream Is Nothing Then
                                                    Debug.WriteLine("[Writer] Stream is null. Exiting task.")
                                                    Return
                                                End If

                                                ' Send the audio chunk to Google STT
                                                Await _stream.WriteAsync(New StreamingRecognizeRequest With {.AudioContent = chunk})

                                            Catch ex As RpcException
                                                ' gRPC error (e.g., stream cancelled) - expected during shutdown/recovery
                                                Debug.WriteLine($"[Writer] RpcException (Status: {ex.StatusCode}). Exiting writer task.")
                                                Return

                                            Catch ex As NullReferenceException
                                                ' Stream set to Nothing by another thread
                                                Debug.WriteLine("[Writer] Stream became null. Exiting writer task.")
                                                Return

                                            Catch ex As InvalidOperationException
                                                ' Stream used after being closed
                                                Debug.WriteLine($"[Writer] InvalidOperationException (likely closed stream). Exiting writer task.")
                                                Return

                                            Catch ex As Exception
                                                ' Unexpected error
                                                Debug.WriteLine($"[Writer] Unhandled exception in write loop: {ex.GetType().Name}. Exiting writer task.")
                                                Return
                                            End Try
                                        Next

                                    Catch ex As InvalidOperationException
                                        ' GetConsumingEnumerable called on completed/disposed collection - expected during recovery
                                        Debug.WriteLine("[Writer] Task ending gracefully due to completed or disposed audio queue.")

                                    Catch ex As Exception
                                        ' Truly unexpected error at task level
                                        Debug.WriteLine($"[Writer] UNEXPECTED FATAL ERROR in writer task: {ex.ToString()}")
                                    End Try
                                End Sub)
        End Sub

        ''' <summary>
        ''' Handles incoming audio data from microphone for Google STT.
        ''' Mixes in system audio if multi-source enabled, adds to ring buffer for recovery,
        ''' enqueues for streaming, and monitors for timeout to trigger recovery.
        ''' </summary>
        ''' <param name="sender">WaveInEvent source</param>
        ''' <param name="e">Audio data event args with buffer and byte count</param>
        Private Async Sub OnGoogleDataAvailable(sender As Object, e As WaveInEventArgs)
            If _googleStreamCompleted Then Return

            Dim now = DateTime.UtcNow
            Dim elapsed = (now - streamingStartTime).TotalMilliseconds

            ' 1) Mix in system audio (loopback) if enabled
            If MultiSourceEnabled AndAlso loopbackCapture IsNot Nothing AndAlso loopbackResampler IsNot Nothing Then
                Dim mixBuf(e.BytesRecorded - 1) As Byte
                Dim bytesRead = loopbackResampler.Read(mixBuf, 0, e.BytesRecorded)
                If bytesRead > 0 Then
                    ' Mix microphone and system audio samples (16-bit PCM)
                    For i As Integer = 0 To bytesRead - 1 Step 2
                        Dim micSample As Integer = BitConverter.ToInt16(e.Buffer, i)
                        Dim outSample As Integer = BitConverter.ToInt16(mixBuf, i)
                        Dim summedSample As Integer = micSample + outSample

                        ' Clamp to Int16 range to prevent overflow
                        If summedSample > Short.MaxValue Then summedSample = Short.MaxValue
                        If summedSample < Short.MinValue Then summedSample = Short.MinValue

                        Dim ba() As Byte = BitConverter.GetBytes(CShort(summedSample))
                        e.Buffer(i) = ba(0)
                        e.Buffer(i + 1) = ba(1)
                    Next
                End If
            End If

            Dim chunk As Google.Protobuf.ByteString =
        Google.Protobuf.ByteString.CopyFrom(e.Buffer, 0, e.BytesRecorded)

            ' 2) Write to ring buffer (maximum 50 chunks for recovery)
            SyncLock ringBuffer
                ringBuffer.Enqueue(chunk)
                If ringBuffer.Count > RING_BUFFER_SIZE Then ringBuffer.Dequeue()
            End SyncLock

            ' 3) Write to queue for streaming (if still open)
            If Not audioQueue.IsAddingCompleted Then
                audioQueue.Add(chunk)
            End If

            ' 4) Check timeout and trigger recovery if needed
            If elapsed > STREAMING_LIMIT_MS Then
                Debug.WriteLine($"[OnGoogleDataAvailable] Timeout detected. Requesting recovery... ts={DateTime.UtcNow:HH:mm:ss.fff}")
                ' Fire-and-forget recovery task to avoid blocking audio processing
                System.Threading.Tasks.Task.Run(Async Sub() Await TryRecoverGoogleStreamAsync())
                ' Reset timer here to prevent immediate re-triggering
                streamingStartTime = DateTime.UtcNow
                Return
            End If

        End Sub

        ''' <summary>
        ''' Recovers Google STT stream after timeout or error by shutting down old components,
        ''' waiting for old reader task completion, initializing new stream with fresh token,
        ''' and restarting writer and reader tasks with new audio queue.
        ''' </summary>
        Private Async Function RecoverGoogleStream() As System.Threading.Tasks.Task
            Debug.WriteLine($"[RecoverGoogleStream] Starting...")

            ' --- 1. SHUTDOWN OLD COMPONENTS ---
            ' Store old task before overwriting class-level variable
            Dim oldReaderTask As System.Threading.Tasks.Task = Me.googleReaderTask

            ' Cancel the old reader's token source
            If readerCts IsNot Nothing Then
                Try
                    readerCts.Cancel()
                    Debug.WriteLine($"[RecoverGoogleStream] Old CancellationTokenSource cancelled.")
                Catch ex As Exception
                    ' Ignore cancellation errors
                End Try
            End If

            ' Gracefully complete and dispose old stream object
            Await SafeCompleteAndDisposeGoogleStreamAsync(readerCts.Token)
            Debug.WriteLine($"[RecoverGoogleStream] Old stream disposed.")

            ' Explicitly wait for old reader task to finish
            ' This is KEY to preventing race conditions
            If oldReaderTask IsNot Nothing Then
                Try
                    Await oldReaderTask
                    Debug.WriteLine($"[RecoverGoogleStream] Old reader task has completed.")
                Catch ex As Exception
                    ' Expected exceptions (like TaskCanceled) - just log and continue
                    Debug.WriteLine($"[RecoverGoogleStream] Awaiting old reader task threw: {ex.GetType().Name}")
                End Try
            End If

            ' --- 2. INITIALIZE NEW COMPONENTS ---

            ' Create new client and stream with fresh OAuth2 token
            Dim newToken As String = Await GetFreshSTTToken(STTSecondAPI)
            Dim callCreds = Grpc.Core.CallCredentials.FromInterceptor(
    Async Function(contextCall, metadata)
        metadata.Add("Authorization", $"Bearer {newToken}")
        Await System.Threading.Tasks.Task.CompletedTask
    End Function)
            Dim channelCreds = Grpc.Core.ChannelCredentials.Create(
    Grpc.Core.ChannelCredentials.SecureSsl, callCreds)
            Dim builder = New Google.Cloud.Speech.V1.SpeechClientBuilder() With {
                        .Endpoint = STTEndpoint,
                        .ChannelCredentials = channelCreds
                    }
            client = builder.Build()

            ' Initialize new stream (must happen after old stream and reader are fully dead)
            ResetGoogleStreamFlag()
            Await InitializeGoogleStream()
            Debug.WriteLine($"[RecoverGoogleStream] New stream initialized.")

            ' --- 3. START NEW TASKS ---

            ' Create fresh audio queue and start writer task
            audioQueue = New System.Collections.Concurrent.BlockingCollection(Of ByteString)()
            StartAudioQueueWriter()
            Debug.WriteLine($"[RecoverGoogleStream] New writer task started.")

            ' Start new reader task and assign to class-level variable
            Me.googleReaderTask = StartGoogleReaderTask()
            Debug.WriteLine($"[RecoverGoogleStream] New reader task started.")
        End Function

        ''' <summary>
        ''' OBSOLETE: Old stream recovery implementation - retained for reference only.
        ''' This function is prefixed with 'x' indicating it should not be called.
        ''' Use RecoverGoogleStream instead.
        ''' </summary>
        Private Async Function xRecoverGoogleStream() As System.Threading.Tasks.Task

            Try
                ' 1) Create new token and client
                Dim newToken As String = Await GetFreshSTTToken(STTSecondAPI)
                Dim callCreds = Grpc.Core.CallCredentials.FromInterceptor(
    Async Function(contextCall, metadata)
        metadata.Add("Authorization", $"Bearer {newToken}")
        Await System.Threading.Tasks.Task.CompletedTask
    End Function)
                Dim channelCreds = Grpc.Core.ChannelCredentials.Create(
    Grpc.Core.ChannelCredentials.SecureSsl, callCreds)
                Dim builder = New Google.Cloud.Speech.V1.SpeechClientBuilder() With {
    .Endpoint = STTEndpoint,
    .ChannelCredentials = channelCreds
}
                client = builder.Build()

                ' 2) Reinitialize stream
                streamingStartTime = DateTime.UtcNow
                ResetGoogleStreamFlag()
                Await InitializeGoogleStream()

                ' 3) Reset offset
                Dim offset As Integer = 0
                Me.Invoke(Sub() offset = RichTextBox1.TextLength)
                googleTranscriptStart = offset

                ' 4) Replay ring buffer into queue
                SyncLock ringBuffer
                    For Each oldChunk In ringBuffer
                        audioQueue.Add(oldChunk)
                    Next
                End SyncLock

                ' 5) Restart reader
                StartGoogleReaderTask()

                SyncLock ringBuffer
                    ringBuffer.Clear()
                End SyncLock

            Catch ex As System.Exception
                ' Silently handle recovery errors
            End Try
        End Function

        ''' <summary>
        ''' OBSOLETE: Old stream disposal implementation - retained for reference only.
        ''' This function is prefixed with 'xx' indicating it should not be called.
        ''' Use SafeCompleteAndDisposeGoogleStreamAsync instead.
        ''' </summary>
        Private Async Function xxSafeCompleteAndDisposeGoogleStreamAsync(token As CancellationToken) As System.Threading.Tasks.Task
            ' 1) Complete stream cleanly
            Try
                If _stream IsNot Nothing AndAlso Not _googleStreamCompleted Then
                    Await _stream.WriteCompleteAsync()
                    _googleStreamCompleted = True
                    ' NOTE: Do NOT call CompleteAdding here
                End If
            Catch ex As System.Exception
                Debug.WriteLine($"Error in SafeComplete…: {ex.Message}")
            End Try

            ' 2) Close queue and wait for writer task
            audioQueue.CompleteAdding()
            Await writerTask

            ' 3) Release stream object
            If _stream IsNot Nothing Then
                _stream.Dispose()
                _stream = Nothing
            End If
        End Function

        ''' <summary>
        ''' Gracefully completes and disposes Google STT stream, handling cancellation scenarios.
        ''' Waits for writer task completion before disposing stream resources.
        ''' Ignores expected exceptions during cancellation and handles queue completion safely.
        ''' </summary>
        ''' <param name="token">Cancellation token to check for forced cancellation</param>
        Private Async Function SafeCompleteAndDisposeGoogleStreamAsync(token As CancellationToken) As System.Threading.Tasks.Task
            ' 1) Complete stream cleanly if still valid and not cancelled
            Try
                If _stream IsNot Nothing AndAlso Not _googleStreamCompleted AndAlso Not token.IsCancellationRequested Then
                    Await _stream.WriteCompleteAsync()
                End If
            Catch ex As RpcException When ex.StatusCode = StatusCode.Cancelled
                ' Expected exception if stream was cancelled - safely ignore and proceed with cleanup
                Debug.WriteLine($"[SafeComplete] Ignored expected RpcException (Cancelled).")
            Catch ex As Exception
                ' Catch other errors but don't let them stop cleanup
                Debug.WriteLine($"[SafeComplete] Error during WriteCompleteAsync: {ex.Message}")
            End Try

            _googleStreamCompleted = True

            ' 2) Wait for writer task to finish (exits when queue completed or exception hit)
            If writerTask IsNot Nothing AndAlso Not writerTask.IsCompleted Then
                Try
                    ' Don't try to complete queue if already done
                    If Not audioQueue.IsAddingCompleted Then
                        audioQueue.CompleteAdding()
                    End If
                    Await writerTask
                Catch ex As Exception
                    Debug.WriteLine($"[SafeComplete] Error while awaiting writerTask: {ex.Message}")
                End Try
            End If

            ' 3) Dispose stream object (this is a local method call, safe to use null-conditional)
            _stream?.Dispose()
            _stream = Nothing
        End Function

        ''' <summary>
        ''' Converts raw 16-bit PCM audio buffer to normalized float array (-1.0 to 1.0).
        ''' Used for Whisper processing which requires floating-point audio samples.
        ''' </summary>
        ''' <param name="buffer">Raw byte buffer containing 16-bit PCM audio data</param>
        ''' <returns>Float array with normalized samples in range -1.0 to 1.0</returns>
        Private Function ConvertAudioToFloat(buffer As Byte()) As Single()
            ' Each sample = 2 bytes (16-bit), so half as many float samples
            Dim floatArray As Single() = New Single((buffer.Length \ 2) - 1) {}

            ' Convert raw 16-bit PCM to normalized float (-1.0 to 1.0)
            For i As Integer = 0 To buffer.Length - 2 Step 2
                Dim sample As Short = BitConverter.ToInt16(buffer, i)
                floatArray(i \ 2) = sample / 32768.0F
            Next

            Return floatArray
        End Function

        ''' <summary>
        ''' Handles loopback (system audio) data available event.
        ''' Buffers system audio for later mixing with microphone input.
        ''' NOTE: Currently unused - audio mixing is done inline in OnAudioDataAvailable.
        ''' </summary>
        ''' <param name="sender">WasapiLoopbackCapture source</param>
        ''' <param name="e">Audio data event args</param>
        Private Sub OnLoopbackDataAvailable(sender As Object, e As WaveInEventArgs)
            ' Buffer the system audio for later mixing
            loopbackBuffer.AddSamples(e.Buffer, 0, e.BytesRecorded)
        End Sub

        ''' <summary>
        ''' Handles incoming audio data from microphone for Vosk and Whisper STT engines.
        ''' Mixes in system audio if multi-source enabled, normalizes audio levels,
        ''' and dispatches to appropriate STT engine for processing.
        ''' </summary>
        ''' <param name="sender">WaveInEvent source</param>
        ''' <param name="e">Audio data event args with buffer and byte count</param>
        Private Async Sub OnAudioDataAvailable(sender As Object, e As WaveInEventArgs)

            ' Mix in system audio (loopback) if multi-source enabled
            If MultiSourceEnabled AndAlso loopbackCapture IsNot Nothing AndAlso loopbackResampler IsNot Nothing Then
                Dim mixBuf(e.BytesRecorded - 1) As Byte
                ' Read same number of bytes from resampler (16kHz mono 16-bit)
                Dim bytesRead = loopbackResampler.Read(mixBuf, 0, e.BytesRecorded)
                If bytesRead > 0 Then
                    ' Mix microphone and system audio samples
                    For i As Integer = 0 To bytesRead - 1 Step 2
                        Dim micSample As Integer = BitConverter.ToInt16(e.Buffer, i)
                        Dim outSample As Integer = BitConverter.ToInt16(mixBuf, i)
                        Dim summedSample As Integer = micSample + outSample

                        ' Clamp to Int16 range
                        If summedSample > Short.MaxValue Then summedSample = Short.MaxValue
                        If summedSample < Short.MinValue Then summedSample = Short.MinValue

                        Dim ba() As Byte = BitConverter.GetBytes(CShort(summedSample))
                        e.Buffer(i) = ba(0)
                        e.Buffer(i + 1) = ba(1)
                    Next
                End If
            End If

            Dim buffer As Byte() = e.Buffer
            Dim bytesRecorded As Integer = e.BytesRecorded

            ' Convert to 16-bit PCM samples and normalize
            Dim sampleCount As Integer = CInt(bytesRecorded / 2)
            Dim samples(sampleCount - 1) As Single ' Float array for normalized audio

            For i As Integer = 0 To sampleCount - 1
                ' Convert 16-bit PCM to float (-1.0 to 1.0)
                Dim sample As Short = BitConverter.ToInt16(buffer, i * 2)
                Dim floatSample As Single = sample / 32768.0F
                samples(i) = floatSample
            Next

            ' Normalize samples (equalize audio levels)
            Dim maxSample As Single = samples.Max(Function(x) System.Math.Abs(x))
            If maxSample > 0 Then
                Dim gain As Single = 1.0F / maxSample ' Compute normalization factor
                For i As Integer = 0 To sampleCount - 1
                    samples(i) *= gain ' Apply normalization
                Next
            End If

            ' Convert back to 16-bit PCM after normalization
            For i As Integer = 0 To sampleCount - 1
                Dim normalizedSample As Short = CShort(samples(i) * 32767)
                Dim bytes As Byte() = BitConverter.GetBytes(normalizedSample)
                buffer(i * 2) = bytes(0)
                buffer(i * 2 + 1) = bytes(1)
            Next

            ' Dispatch to appropriate STT engine
            Select Case STTModel
                Case "vosk"
                    If recognizer IsNot Nothing AndAlso capturing Then
                        Dim jsonResult As String = ""
                        jsonResult = If(recognizer.AcceptWaveform(e.Buffer, e.BytesRecorded),
                                        recognizer.Result, recognizer.PartialResult)
                        ProcessTranscriptionJson(jsonResult)
                    End If

                Case "whisper"

                    If WhisperRecognizer Is Nothing Then Return

                    Try
                        ' Convert audio buffer to float array for Whisper
                        Dim whispersamples As Single() = ConvertAudioToFloat(e.Buffer)

                        ' Append to buffer and wait for enough data
                        audioBuffer.AddRange(whispersamples)
                        If audioBuffer.Count < 32000 Then Return ' Wait for ~2 seconds of audio at 16kHz

                        ' Copy buffered audio and clear buffer for next batch
                        Dim processSamples = audioBuffer.ToArray()
                        audioBuffer.Clear()
                        e.Buffer.Initialize() ' Clear the input buffer

                        ' Process transcription asynchronously
                        Await ProcessWhisper(processSamples)
                    Catch ex As Exception
                        Debug.WriteLine($"Error in OnAudioDataAvailable: {ex.Message}")
                    End Try
            End Select

        End Sub

        ''' <summary>
        ''' Processes Whisper transcription results from audio samples.
        ''' Filters out bracketed and asterisk-enclosed text patterns, then appends valid text to RichTextBox.
        ''' Handles single result batch to prevent duplicate processing.
        ''' </summary>
        ''' <param name="samples">Float array of normalized audio samples from Whisper</param>
        Private Async Function ProcessWhisper(samples As Single()) As System.Threading.Tasks.Task
            Try
                If STTCanceled Then Return

                Dim segments As IAsyncEnumerable(Of SegmentData) = WhisperRecognizer.ProcessAsync(samples)

                ' Iterate over the transcription results (only once)
                Dim enumerator = segments.GetAsyncEnumerator()

                If Await enumerator.MoveNextAsync() Then ' Only process the first result batch
                    Dim result As SegmentData = enumerator.Current
                    Dim text As String = result.Text

                    ' Remove bracketed content like [MUSIC] and asterisk-enclosed content like *applause*
                    text = Regex.Replace(text, "\[.*?\]", String.Empty)
                    text = Regex.Replace(text, "\*.*?\*", String.Empty)

                    If Not String.IsNullOrWhiteSpace(text) And Not STTCanceled Then
                        Me.Invoke(Sub()
                                      RichTextBox1.AppendText(text & vbCrLf)
                                      RichTextBox1.ScrollToCaret()
                                  End Sub)
                    End If
                End If

                Await enumerator.DisposeAsync()

            Catch ex As Exception
                Debug.WriteLine($"Error in ProcessWhisper: {ex.Message}")
            End Try
        End Function

        ''' <summary>
        ''' Transcribes an entire audio file using Whisper STT engine.
        ''' Loads audio file, converts to float array, processes asynchronously with cancellation support,
        ''' and updates UI with transcription progress. Resets UI state and shows completion message when done.
        ''' </summary>
        ''' <param name="filepath">Path to audio file (WAV, MP3, AAC, M4A, WMA)</param>
        Public Async Function WhisperTranscribeAudioFile(filepath As String) As System.Threading.Tasks.Task

            Try
                PartialTextLabel.Invoke(Sub() PartialTextLabel.Text = "Whisper is reading and transcribing your file...")

                ' Load and convert audio file to float samples required by Whisper
                Dim samples As Single() = LoadAudioToFloatArray(filepath)

                Dim segments As IAsyncEnumerable(Of SegmentData) = WhisperRecognizer.ProcessAsync(samples)

                Dim enumerator = segments.GetAsyncEnumerator()

                Dim Exited As Boolean = False

                ' Process all segments until completion or cancellation
                While Await enumerator.MoveNextAsync()

                    If cts.Token.IsCancellationRequested Then
                        Exited = True
                        Exit While
                    End If

                    Dim result As SegmentData = enumerator.Current
                    Dim Text = result.Text
                    If Not String.IsNullOrWhiteSpace(Text) And Not STTCanceled Then
                        Me.Invoke(Sub()
                                      RichTextBox1.AppendText(Text & vbCrLf)
                                      RichTextBox1.ScrollToCaret()
                                  End Sub)
                    End If

                End While
                Await enumerator.DisposeAsync()

                ' Cleanup and reset UI state
                STTCanceled = True
                Await StopRecording()
                capturing = False
                Me.StartButton.Enabled = True
                Me.StopButton.Enabled = False
                Me.AudioButton.Enabled = True
                Me.LoadButton.Enabled = True
                Me.cultureComboBox.Enabled = True
                Me.deviceComboBox.Enabled = True
                Me.SpeakerIdent.Enabled = True
                Me.SpeakerDistance.Enabled = True
                PartialTextLabel.Invoke(Sub() PartialTextLabel.Text = "")

                If Exited Then
                    ShowCustomMessageBox("Transcription aborted.")
                Else
                    ShowCustomMessageBox("The transcription of your file is complete.")
                End If

            Catch ex As Exception
                Debug.WriteLine($"Error in WhisperTranscribeAudioFile: {ex.Message}")
            End Try
        End Function

        ''' <summary>
        ''' Cancels ongoing transcription by triggering the cancellation token.
        ''' Safe to call even if no transcription is active.
        ''' </summary>
        Public Sub CancelTranscription()
            If cts IsNot Nothing Then
                cts.Cancel()
            End If
        End Sub

        ''' <summary>
        ''' Loads an audio file and converts it to a float array suitable for Whisper processing.
        ''' Supports multiple formats (MP3, WAV, FLAC, etc.) via MediaFoundationReader.
        ''' Resamples audio to 16 kHz mono and converts to normalized float samples (-1.0 to 1.0).
        ''' </summary>
        ''' <param name="filepath">Path to audio file</param>
        ''' <returns>Float array of normalized audio samples at 16 kHz mono</returns>
        Public Function LoadAudioToFloatArray(filepath As String) As Single()
            Using reader As New MediaFoundationReader(filepath) ' Supports MP3, WAV, FLAC, etc.
                ' Convert audio to 16kHz Mono (Whisper requires this format)
                Dim waveFormat = New WaveFormat(16000, 1) ' 16kHz, Mono
                Using resampler As New MediaFoundationResampler(reader, waveFormat)
                    resampler.ResamplerQuality = 60

                    ' Convert to floating point explicitly
                    Dim floatProvider As ISampleProvider = resampler.ToSampleProvider()

                    ' Read audio data into a floating-point array
                    Dim samples As New List(Of Single)()
                    Dim buffer As Single() = New Single(1024 - 1) {} ' Buffer for PCM float samples
                    Dim samplesRead As Integer

                    Do
                        samplesRead = floatProvider.Read(buffer, 0, buffer.Length)
                        If samplesRead > 0 Then
                            samples.AddRange(buffer.Take(samplesRead))
                        End If
                    Loop While samplesRead > 0

                    Return samples.ToArray()
                End Using
            End Using
        End Function

        ''' <summary>
        ''' Transcribes an audio file using Google Cloud Speech-to-Text in chunked mode.
        ''' Splits audio into 50-second slices with 2-second overlap, processes each via RecognizeAsync,
        ''' and appends results to transcript. Self-terminating after completion.
        ''' Ensures Google client is initialized with fresh OAuth2 token before processing.
        ''' </summary>
        ''' <param name="filepath">Path to audio file to transcribe</param>
        Public Async Function GoogleChunkedTranscribeAudioFile(filepath As String) _
As System.Threading.Tasks.Task

            ' 0) Ensure client is initialized with fresh OAuth2 credentials
            If client Is Nothing Then
                Dim tokenToSend As String = Await GetFreshSTTToken(STTSecondAPI)
                Dim callCreds As Grpc.Core.CallCredentials = Grpc.Core.CallCredentials.FromInterceptor(
            Async Function(contextCall, metadata)
                metadata.Add("Authorization", $"Bearer {tokenToSend}")
                Await System.Threading.Tasks.Task.CompletedTask
            End Function
        )
                Dim channelCreds As Grpc.Core.ChannelCredentials = Grpc.Core.ChannelCredentials.Create(
            Grpc.Core.ChannelCredentials.SecureSsl,
            callCreds
        )
                Dim builder As New Google.Cloud.Speech.V1.SpeechClientBuilder() With {
            .Endpoint = STTEndpoint,
            .ChannelCredentials = channelCreds
        }
                client = builder.Build()
            End If

            ' 1) Load PCM data (16 kHz, mono, 16-bit)
            Dim pcmData As Byte() = LoadAudioToPCM(filepath)

            ' 2) Chunk parameters
            Dim bytesPerSec As Integer = 16000 * 2      ' 32,000 bytes/sec (16 kHz * 16-bit)
            Dim sliceLenSec As Integer = 50              ' 50-second chunks
            Dim overlapSec As Integer = 2                ' 2-second overlap between chunks
            Dim sliceSize As Integer = sliceLenSec * bytesPerSec
            Dim overlapSize As Integer = overlapSec * bytesPerSec
            Dim offset As Integer = 0

            ' 3) Loop through all audio slices
            While offset < pcmData.Length AndAlso Not STTCanceled
                Dim endPos = System.Math.Min(offset + sliceSize, pcmData.Length)
                Dim slice(endPos - offset - 1) As Byte
                Array.Copy(pcmData, offset, slice, 0, endPos - offset)

                ' 4) Build RecognitionConfig for this chunk
                Dim config As New Google.Cloud.Speech.V1.RecognitionConfig With {
            .Encoding = Google.Cloud.Speech.V1.RecognitionConfig.Types.AudioEncoding.Linear16,
            .SampleRateHertz = 16000,
            .LanguageCode = GoogleLanguageCode,
            .EnableAutomaticPunctuation = True,
            .Model = "latest_long",
            .UseEnhanced = True
        }
                Dim audio As Google.Cloud.Speech.V1.RecognitionAudio =
            Google.Cloud.Speech.V1.RecognitionAudio.FromBytes(slice)

                ' 5) Synchronous API call (non-streaming)
                Dim response As Google.Cloud.Speech.V1.RecognizeResponse =
            Await client.RecognizeAsync(config, audio)

                ' 6) Append results to transcript
                For Each result As Google.Cloud.Speech.V1.SpeechRecognitionResult In response.Results
                    If result.Alternatives.Count > 0 Then
                        Addline(result.Alternatives(0).Transcript)
                    End If
                Next

                ' 7) Exit if this was the last slice
                If endPos >= pcmData.Length Then
                    Exit While
                End If

                ' Otherwise, advance offset with overlap
                offset = endPos - overlapSize
                If offset < 0 Then offset = 0
            End While

            ' 8) Show completion message
            ShowCustomMessageBox("Chunked transcription complete.", $"{AN} Transcriptor")
        End Function

        ''' <summary>
        ''' Transcribes a local audio file via Google STT StreamingRecognize.
        ''' Feeds audio through existing audioQueue → writerTask → readerTask pipeline at file tempo (16 kHz),
        ''' then cleanly shuts down stream. Displays streaming progress and completion message.
        ''' </summary>
        ''' <param name="filepath">Path to audio file to transcribe</param>
        Public Async Function GoogleFileStreamTranscription(filepath As String) As System.Threading.Tasks.Task
            ' 1) Update UI status
            PartialTextLabel.Invoke(Sub() PartialTextLabel.Text = $"{GoogleSTT_Desc} streaming file…")

            ' 2) Initialize stream and reader task
            readerCts = New CancellationTokenSource()
            Await StartGoogleSTT()                             ' Opens _stream & writes config
            googleTranscriptStart = RichTextBox1.TextLength
            googleReaderTask = StartGoogleReaderTask()         ' Starts reading responses

            ' 3) Reset queue and writer
            audioQueue = New BlockingCollection(Of Google.Protobuf.ByteString)()
            StartAudioQueueWriter()                            ' Writes from audioQueue to _stream

            ' 4) Load PCM data from file
            Dim pcmFull As Byte() = LoadAudioToPCM(filepath)

            ' 5) Remove RIFF header if WAV format
            Dim pcmData = If(
        pcmFull.Length > 44 AndAlso
        System.Text.Encoding.ASCII.GetString(pcmFull, 0, 4) = "RIFF",
        pcmFull.Skip(44).ToArray(),
        pcmFull
        )

            ' 6) Enqueue audio at file tempo (16 kHz)
            Const chunkSize As Integer = 4096
            Dim bytesPerSec As Integer = 16000 * 2  ' 16 kHz × 16-bit Mono = 32,000 bytes/sec
            Dim pos As Integer = 0

            While pos < pcmData.Length AndAlso Not STTCanceled
                Dim len = System.Math.Min(chunkSize, pcmData.Length - pos)
                Dim chunk = Google.Protobuf.ByteString.CopyFrom(pcmData, pos, len)
                audioQueue.Add(chunk)

                ' Throttle to file tempo (simulates real-time streaming)
                Dim delayMs = CInt(1000.0 * len / bytesPerSec)
                Await System.Threading.Tasks.Task.Delay(delayMs)

                pos += len
            End While

            ' 7) Close queue → writer knows no more audio is coming
            audioQueue.CompleteAdding()

            ' 8) Cleanly complete stream and wait for reader task
            Await SafeCompleteAndDisposeGoogleStreamAsync(readerCts.Token)
            Await googleReaderTask

            ' 9) Cleanup and re-enable UI
            StopApiWatchdogTimer()
            PartialTextLabel.Invoke(Sub() PartialTextLabel.Text = "")
            ShowCustomMessageBox("Streaming transcription complete.", $"{AN} Transcriptor")
            Me.Invoke(Sub()
                          capturing = False
                          StartButton.Enabled = True
                          StopButton.Enabled = False
                          LoadButton.Enabled = True
                          AudioButton.Enabled = True
                          cultureComboBox.Enabled = True
                          deviceComboBox.Enabled = True
                          SpeakerIdent.Enabled = True
                          SpeakerDistance.Enabled = True
                      End Sub)
        End Function

        ''' <summary>
        ''' Transcribes an audio file using Vosk STT engine with Escape key cancellation support.
        ''' Processes audio in 4KB chunks, extracts text from JSON results, And updates UI in real-time.
        ''' Resets UI state and displays completion/cancellation message when done.
        ''' </summary>
        ''' <param name="filepath">Path to audio file to transcribe</param>
        Public Async Function VoskTranscribeAudioFile(filepath As String) As System.Threading.Tasks.Task
            Try
                PartialTextLabel.Invoke(Sub() PartialTextLabel.Text = "Vosk is reading and transcribing your file... press 'Esc' to abort")

                Dim Exited As Boolean = False

                ' Load PCM audio directly (no float conversion needed for Vosk)
                Dim pcmData As Byte() = LoadAudioToPCM(filepath)

                ' Initialize Vosk recognizer
                recognizer.Reset()

                ' Stream PCM data to Vosk recognizer in small chunks
                Dim chunkSize As Integer = 4096
                Dim offset As Integer = 0

                While offset < pcmData.Length

                    ' Allow UI to process events (for Escape key detection)
                    System.Windows.Forms.Application.DoEvents()

                    ' Check for Escape key press (both current state and toggle state)
                    If (GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0 Then
                        Exited = True
                        Exit While
                    End If

                    If (GetAsyncKeyState(VK_ESCAPE) And 1) <> 0 Then
                        Exited = True
                        Exit While
                    End If

                    ' Extract chunk from audio data
                    Dim chunkLength As Integer = System.Math.Min(chunkSize, pcmData.Length - offset)
                    Dim chunk As Byte() = pcmData.Skip(offset).Take(chunkLength).ToArray()

                    ' Feed the chunk into the recognizer
                    Dim resultAvailable As Boolean = recognizer.AcceptWaveform(chunk, chunk.Length)

                    ' Retrieve transcription (final or partial)
                    Dim resultText As String
                    If resultAvailable Then
                        Dim resultJson As String = recognizer.Result()
                        resultText = ExtractTextFromJson(resultJson)
                    Else
                        Dim partialJson As String = recognizer.PartialResult()
                        resultText = ExtractTextFromJson(partialJson)
                    End If

                    ' Update UI with transcribed text
                    If Not String.IsNullOrWhiteSpace(resultText) And Not STTCanceled Then
                        Me.Invoke(Sub()
                                      RichTextBox1.AppendText(resultText & vbCrLf)
                                      RichTextBox1.ScrollToCaret()
                                  End Sub)
                    End If

                    offset += chunkLength
                End While

                ' Get final result from recognizer
                Dim finalResultJson As String = recognizer.FinalResult()
                Dim finalText As String = ExtractTextFromJson(finalResultJson)

                If Not String.IsNullOrWhiteSpace(finalText) Then
                    Me.Invoke(Sub()
                                  RichTextBox1.AppendText(finalText & vbCrLf)
                                  RichTextBox1.ScrollToCaret()
                              End Sub)
                End If

                ' Reset flags and UI state
                PartialTextLabel.Invoke(Sub() PartialTextLabel.Text = "")
                STTCanceled = True
                Await StopRecording()
                capturing = False
                Me.StartButton.Enabled = True
                Me.StopButton.Enabled = False
                Me.LoadButton.Enabled = True
                Me.AudioButton.Enabled = True
                Me.cultureComboBox.Enabled = True
                Me.deviceComboBox.Enabled = True
                Me.SpeakerIdent.Enabled = True
                Me.SpeakerDistance.Enabled = True
                PartialTextLabel.Invoke(Sub() PartialTextLabel.Text = "")

                If Exited Then
                    ShowCustomMessageBox("Transcription aborted.")
                Else
                    ShowCustomMessageBox("The transcription of your file is complete.")
                End If
            Catch ex As Exception
                Debug.WriteLine($"Error in VoskTranscribeAudioFile: {ex.Message}")
            End Try
        End Function

        ''' <summary>
        ''' Extracts transcribed text from Vosk JSON response.
        ''' Parses JSON and returns the "text" field value, or empty string if not found.
        ''' </summary>
        ''' <param name="jsonString">JSON string from Vosk recognizer</param>
        ''' <returns>Extracted text or empty string</returns>
        Private Function ExtractTextFromJson(jsonString As String) As String
            Try
                Dim json As JObject = JObject.Parse(jsonString)
                If json.ContainsKey("text") Then
                    Return json("text").ToString()
                Else
                    Return String.Empty
                End If
            Catch ex As Exception
                Debug.WriteLine($"JSON Parsing Error: {ex.Message}")
                Return String.Empty
            End Try
        End Function

        ''' <summary>
        ''' Loads an audio file and converts it to 16 kHz mono PCM byte array.
        ''' Supports multiple formats (MP3, WAV, FLAC, etc.) via MediaFoundationReader.
        ''' Returns raw PCM data suitable for Vosk or Google STT processing.
        ''' </summary>
        ''' <param name="filepath">Path to audio file</param>
        ''' <returns>Byte array containing 16 kHz, 16-bit, mono PCM audio data</returns>
        Public Function LoadAudioToPCM(filepath As String) As Byte()
            Using reader As New MediaFoundationReader(filepath) ' Supports MP3, WAV, FLAC, etc.
                ' Convert audio to 16kHz Mono PCM (required by Vosk and Google STT)
                Dim waveFormat = New WaveFormat(16000, 16, 1) ' 16kHz, 16-bit, Mono

                Using resampler As New MediaFoundationResampler(reader, waveFormat)
                    resampler.ResamplerQuality = 60

                    ' Use MemoryStream to store PCM data
                    Using memoryStream As New MemoryStream()
                        Using pcmWriter As New WaveFileWriter(memoryStream, waveFormat)
                            Dim buffer(4096 - 1) As Byte
                            Dim bytesRead As Integer

                            Do
                                bytesRead = resampler.Read(buffer, 0, buffer.Length)
                                If bytesRead > 0 Then
                                    pcmWriter.Write(buffer, 0, bytesRead)
                                End If
                            Loop While bytesRead > 0

                            pcmWriter.Flush()
                        End Using

                        ' Return raw PCM byte array
                        Return memoryStream.ToArray()
                    End Using
                End Using
            End Using
        End Function

        ''' <summary>
        ''' Processes Vosk JSON transcription results with optional speaker identification.
        ''' Extracts text and speaker embeddings from JSON, identifies speakers via Euclidean distance,
        ''' and adds formatted text to transcript. Handles both final and partial results.
        ''' </summary>
        ''' <param name="jsonString">JSON string from Vosk recognizer containing text and optional speaker embeddings</param>
        Private Sub ProcessTranscriptionJson(jsonString As String)
            Try
                Dim jsonObject As JObject = JObject.Parse(jsonString)

                If jsonObject.ContainsKey("text") AndAlso jsonObject("text") IsNot Nothing Then
                    Dim completedLine As String = jsonObject("text").ToString()
                    If Not String.IsNullOrWhiteSpace(completedLine) Then

                        ' Check if speaker embeddings are available
                        If jsonObject.ContainsKey("spk") AndAlso jsonObject("spk").Type = JTokenType.Array Then
                            Dim speakerArray As JArray = jsonObject("spk")
                            Dim speakerEmbedding As List(Of Double) = speakerArray.Select(Function(x) CDbl(x)).ToList()

                            ' Identify the speaker using Euclidean distance
                            Dim speakerID As String = IdentifySpeaker(speakerEmbedding)
                            completedLine = $"{speakerID}: " & completedLine
                        End If

                        ' Add line to transcript
                        Addline(completedLine)
                    End If
                ElseIf jsonObject.ContainsKey("partial") AndAlso jsonObject("partial") IsNot Nothing Then
                    ' Update partial text label with interim result
                    partialText = jsonObject("partial").ToString()
                    PartialTextLabel.Invoke(Sub() PartialTextLabel.Text = partialText)
                End If

            Catch ex As Exception
                MessageBox.Show("Error in ProcessTranscriptionJson: " & ex.Message, "Error")
            End Try
        End Sub

        ' ============================================================================
        ' Vosk Speaker Identification via Euclidean Distance
        ' ============================================================================

        ''' <summary>
        ''' Dictionary storing multiple embeddings per speaker for improved matching accuracy.
        ''' Key = Speaker ID (e.g., "Speaker 1"), Value = List of normalized embedding vectors.
        ''' </summary>
        Dim knownSpeakers As New Dictionary(Of String, List(Of List(Of Double)))

        ''' <summary>
        ''' Euclidean distance threshold for speaker matching.
        ''' Lower values = stricter matching (0.5-0.7 for real-time, 1.0-1.5 for meetings).
        ''' </summary>
        Dim similarityThreshold As Double = 1.0

        ''' <summary>
        ''' Identifies a speaker from their voice embedding using Euclidean distance.
        ''' Compares new embedding against average embeddings of known speakers.
        ''' Assigns new speaker ID if no match found within threshold.
        ''' Stores embedding history (last 5) to stabilize detection across variations.
        ''' </summary>
        ''' <param name="newEmbedding">Speaker embedding vector from Vosk</param>
        ''' <returns>Speaker ID string (e.g., "Speaker 1", "Speaker 2")</returns>
        Private Function IdentifySpeaker(newEmbedding As List(Of Double)) As String
            ' Normalize new embedding for consistent comparison
            newEmbedding = NormalizeEmbedding(newEmbedding)

            Dim bestMatch As String = "Unknown"
            Dim bestDistance As Double = Double.MaxValue

            For Each kvp In knownSpeakers
                Dim existingEmbeddings As List(Of List(Of Double)) = kvp.Value

                ' Compute similarity with the average embedding of the stored speaker
                Dim avgEmbedding As List(Of Double) = GetAverageEmbedding(existingEmbeddings)
                Dim distance As Double = EuclideanDistance(avgEmbedding, newEmbedding)

                ' Consider as the same speaker if distance is below threshold
                If distance < bestDistance AndAlso distance < similarityThreshold Then
                    bestMatch = kvp.Key
                    bestDistance = distance
                End If
            Next

            ' If no match, assign a new speaker ID
            If bestMatch = "Unknown" Then
                Dim newSpeakerID As String = "Speaker " & (knownSpeakers.Count + 1).ToString()
                knownSpeakers(newSpeakerID) = New List(Of List(Of Double)) From {newEmbedding}
                Return newSpeakerID
            Else
                ' Store the new embedding for future matches (stabilizes detection)
                knownSpeakers(bestMatch).Add(newEmbedding)

                ' Limit stored embeddings to the last 5 to prevent memory overuse
                If knownSpeakers(bestMatch).Count > 5 Then
                    knownSpeakers(bestMatch).RemoveAt(0)
                End If

                Return bestMatch
            End If
        End Function

        ''' <summary>
        ''' Normalizes an embedding vector to unit length using L2 normalization.
        ''' Ensures embeddings are comparable regardless of magnitude.
        ''' </summary>
        ''' <param name="embedding">Raw embedding vector</param>
        ''' <returns>Normalized embedding vector with L2 norm = 1</returns>
        Private Function NormalizeEmbedding(embedding As List(Of Double)) As List(Of Double)
            Dim norm As Double = System.Math.Sqrt(embedding.Sum(Function(x) x * x))
            If norm = 0 Then Return embedding
            Return embedding.Select(Function(x) x / norm).ToList()
        End Function

        ''' <summary>
        ''' Computes the average embedding from multiple stored embeddings for a speaker.
        ''' Used to create a stable reference embedding that accounts for voice variations.
        ''' </summary>
        ''' <param name="embeddings">List of embedding vectors for a single speaker</param>
        ''' <returns>Average embedding vector</returns>
        Private Function GetAverageEmbedding(embeddings As List(Of List(Of Double))) As List(Of Double)
            Dim embeddingSize As Integer = embeddings(0).Count
            Dim avgEmbedding As New List(Of Double)(New Double(embeddingSize - 1) {})

            ' Sum up all embeddings
            For Each emb In embeddings
                For i As Integer = 0 To embeddingSize - 1
                    avgEmbedding(i) += emb(i)
                Next
            Next

            ' Divide by the number of stored embeddings
            For i As Integer = 0 To embeddingSize - 1
                avgEmbedding(i) /= embeddings.Count
            Next

            Return avgEmbedding
        End Function

''' <summary>
''' Computes Euclidean distance between two speaker embedding vectors.
''' Lower distance indicates more similar speakers.
/// Formula sqrt(sum((vec1[i] - vec2[i])^2))
''' </summary>
''' <param name="vec1">First embedding vector</param>
''' <param name="vec2">Second embedding vector</param>
''' <returns>Euclidean distance (0 = identical, larger = more different)</returns>
Private Function EuclideanDistance(vec1 As List(Of Double), vec2 As List(Of Double)) As Double
            Dim sum As Double = 0
            For i As Integer = 0 To vec1.Count - 1
                sum += (vec1(i) - vec2(i)) ^ 2
            Next
            Return System.Math.Sqrt(sum)
        End Function

        ''' <summary>
        ''' Computes cosine similarity between two speaker embedding vectors.
        ''' Alternative to Euclidean distance (currently unused in favor of Euclidean).
        ''' Formula: dot(vec1, vec2) / (||vec1|| * ||vec2||)
        ''' </summary>
        ''' <param name="vec1">First embedding vector</param>
        ''' <param name="vec2">Second embedding vector</param>
        ''' <returns>Cosine similarity (1 = identical, 0 = orthogonal, -1 = opposite)</returns>
        Private Function CosineSimilarity(vec1 As List(Of Double), vec2 As List(Of Double)) As Double
            Dim dotProduct As Double = vec1.Zip(vec2, Function(a, b) a * b).Sum()
            Dim magnitude1 As Double = System.Math.Sqrt(vec1.Sum(Function(a) a * a))
            Dim magnitude2 As Double = System.Math.Sqrt(vec2.Sum(Function(b) b * b))

            If magnitude1 = 0 OrElse magnitude2 = 0 Then
                Return 0
            End If

            Return dotProduct / (magnitude1 * magnitude2)
        End Function

        ''' <summary>
        ''' Thread-safely appends a completed transcription line to the RichTextBox and internal buffer.
        ''' Clears partial text label, scrolls to caret, and updates Google transcript start position if applicable.
        ''' </summary>
        ''' <param name="completedline">Finalized transcription text to add</param>
        Private Sub Addline(completedline As String)
            completedline = completedline.Trim()

            ' Thread-safe append to internal text buffer
            SyncLock finalText
                finalText.AppendLine(completedline)
            End SyncLock

            ' Thread-safe UI update (deadlock-safe as it only writes)
            RichTextBox1.Invoke(Sub()
                                    ' Clear the partial text label
                                    PartialTextLabel.Text = ""

                                    ' Append the new completed line
                                    RichTextBox1.AppendText(completedline & vbCrLf)

                                    RichTextBox1.SelectionStart = RichTextBox1.Text.Length
                                    RichTextBox1.ScrollToCaret()
                                    If STTModel = "google" Then googleTranscriptStart = RichTextBox1.TextLength
                                End Sub)
        End Sub

        ''' <summary>
        ''' Replaces Google STT transcript text from session start position to end with new full transcript.
        ''' Used for updating entire Google transcription session when receiving complete results.
        ''' Thread-safe via Invoke.
        ''' </summary>
        ''' <param name="fullTranscript">Complete transcript text to replace existing Google session text</param>
        Private Sub ReplaceAndAddLine(fullTranscript As String)
            RichTextBox1.Invoke(Sub()
                                    ' 1) Select everything from the start index to the end
                                    RichTextBox1.Select(googleTranscriptStart, RichTextBox1.TextLength - googleTranscriptStart)
                                    ' 2) Replace it with the entire new transcript
                                    RichTextBox1.SelectedText = fullTranscript & Environment.NewLine
                                    ' 3) Reset the caret to the end
                                    RichTextBox1.SelectionStart = RichTextBox1.Text.Length
                                    RichTextBox1.ScrollToCaret()
                                    If STTModel = "google" Then googleTranscriptStart = RichTextBox1.TextLength
                                End Sub)
        End Sub

        ''' <summary>
        ''' Loads transcript processing prompts from file and populates the process ComboBox.
        ''' Parses pipe-delimited prompt library file, extracts titles and prompt text.
        ''' Displays error messages for missing or malformed files.
        ''' </summary>
        ''' <param name="filePath">Path to transcript prompt library file</param>
        ''' <param name="processComboBox">ComboBox to populate with prompt titles</param>
        Public Sub LoadAndPopulateProcessComboBox(filePath As String, processComboBox As Forms.ComboBox)
            ' Execute LoadPrompts function
            Dim resultCode As Integer = LoadTranscriptPrompts(ExpandEnvironmentVariables(filePath))

            ' Clear the combo box before populating
            processComboBox.Items.Clear()

            ' Check if prompts were successfully loaded
            If resultCode = 0 AndAlso TranscriptPromptsTitles.Count > 0 Then
                ' Add the titles to the combo box
                For Each title As String In TranscriptPromptsTitles
                    processComboBox.Items.Add(title)
                Next
            End If
        End Sub

''' <summary>
''' Loads transcript processing prompts from a pipe-delimited file.
''' File format: Title|Prompt text (one per line, lines starting with ';' are comments).
''' Returns error codes for various failure conditions.
/// </summary>
''' <param name="filePath">Path to prompt library file (environment variables expanded)</param>
''' <returns>0=success, 1=file not found, 2=format error, 3=no prompts found, 99=unexpected error</returns>
Private Function LoadTranscriptPrompts(filePath As String) As Integer

            ' Initialize the return code to 0 (no error)
            Dim returnCode As Integer = 0

            filePath = ExpandEnvironmentVariables(filePath)

            Try
                ' Verify the file exists
                If Not System.IO.File.Exists(filePath) Then
                    ShowCustomMessageBox("The transcript prompt library file was not found.")
                    Return 1
                End If

                TranscriptPromptsTitles.Clear()
                TranscriptPromptsLibrary.Clear()

                ' Read all lines from the file
                Dim lines = System.IO.File.ReadAllLines(filePath)

                For Each line As String In lines
                    ' Trim leading and trailing spaces
                    Dim trimmedLine = line.Trim()

                    ' Ignore empty lines and lines starting with ';' (comments)
                    If Not String.IsNullOrEmpty(trimmedLine) AndAlso Not trimmedLine.StartsWith(";") Then
                        ' Split the line by the delimiter '|'
                        Dim promptData = trimmedLine.Split("|"c)

                        ' Ensure there are at least two parts (title and prompt)
                        If promptData.Length >= 2 Then
                            Dim title = promptData(0).Trim()
                            Dim prompt = String.Join("|", promptData.Skip(1)).Trim()

                            ' Add title and prompt to the respective lists
                            TranscriptPromptsTitles.Add(title)
                            TranscriptPromptsLibrary.Add(prompt)
                        End If
                    End If
                Next

                ' Check if no prompts were found
                If TranscriptPromptsLibrary.Count = 0 Then
                    returnCode = 3
                    ShowCustomMessageBox("No prompts have been found in the configured transcript prompt library file.")
                End If

            Catch ex As System.IO.FileNotFoundException
                returnCode = 1
                ShowCustomMessageBox("The transcript prompt library file was not found: " & ex.Message)

            Catch ex As IndexOutOfRangeException
                returnCode = 2
                ShowCustomMessageBox("The format of the transcript prompt library file is not correct (is a '|' or text thereafter missing?): " & ex.Message)

            Catch ex As Exception
                returnCode = 99
                ShowCustomMessageBox("An unexpected error occurred while loading transcript prompts: " & ex.Message)
            End Try

            Return returnCode
        End Function

    End Class

    ''' <summary>
    ''' Simple dialog form for stopping transcription operations.
    ''' Displays a centered dialog with a Stop button that sets StopRequested flag.
    ''' Used for user-initiated transcription cancellation.

    Public Class StopForm
        Inherits Form

        ''' <summary>
        ''' Flag indicating whether user clicked the Stop button.
        ''' </summary>
        Public Property StopRequested As Boolean = False

        ''' <summary>
        ''' Initializes the StopForm with German title and centered Stop button.
        ''' </summary>
        Public Sub New()
            Me.Text = "Stop Transcription"
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.Width = 200
            Me.Height = 100

            Dim btnStop As New System.Windows.Forms.Button() With {
        .Text = "Stop",
        .Dock = DockStyle.Fill
    }
            AddHandler btnStop.Click, Sub(s, e)
                                          Me.StopRequested = True
                                          Me.Close()
                                      End Sub

            Me.Controls.Add(btnStop)
        End Sub
    End Class


End Class
