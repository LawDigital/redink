' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.TextToSpeech.vb
' Purpose: Provides text-to-speech (TTS) functionality for Word documents using
'          Google Cloud TTS, OpenAI TTS, and legacy Windows SAPI engines.
'
' Architecture:
'  - Multi-Engine Support: Detects and manages Google TTS (OAuth2-based) and 
'    OpenAI TTS (API key-based) endpoints from INI configuration. Falls back
'    to legacy Windows SpeechSynthesizer when cloud engines unavailable.
'  - Token Management: Caches OAuth2 tokens (ttsAccessToken1/2) with expiry tracking;
'    refreshes via GetFreshTTSToken when needed.
'  - Sleep Lock Management: Cooperative system sleep prevention during TTS operations
'    via AcquireTTSSleepLock/ReleaseTTSSleepLock to avoid interruptions.
'  - Podcast Generation: Parses dialogue text (host/guest tags), generates multi-speaker
'    audio with voice alternation, merges segments via NAudio, encodes to MP3.
'  - Paragraph Processing: Iterates Word selection paragraphs, handles titles/bullets,
'    optional LLM-based text cleaning, generates per-paragraph audio, inserts silence
'    between segments, merges final output.
'  - Audio Processing: Uses NAudio for MP3 reading/writing, MediaFoundation for encoding,
'    resampling to uniform PCM format (44.1kHz, 16-bit, stereo) before merging.
'  - SSML Support: Conditionally wraps text in <speak> tags unless NoSSML flag set.
'  - Cancellation: Monitors VK_ESCAPE key state and ProgressBarModule.CancelOperation
'    flag for user-initiated abort.
'  - External Dependencies: SharedLibrary.SharedMethods (UI dialogs, progress tracking,
'    clipboard), NAudio (audio I/O), GoogleOAuthHelper (token acquisition),
'    Microsoft.Office.Interop.Word (document access).
' =============================================================================

Option Explicit On
Option Strict On

Imports System.Diagnostics
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Runtime.InteropServices
Imports System.Speech.Synthesis
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports DocumentFormat.OpenXml
Imports Microsoft.Office.Interop.Word
Imports NAudio.Wave
Imports NetOffice.PowerPointApi
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ' ==================== PODCAST SCRIPT GENERATION ====================

    ''' <summary>
    ''' Generates a podcast script from selected text using LLM with configurable parameters
    ''' (host/guest names, target audience, duration, language, context).
    ''' Displays parameter input form, saves settings, and processes text via InterpolateAtRuntime(SP_Podcast).
    ''' </summary>
    Public Async Sub CreatePodcast()
        If INILoadFail() Then Return
        Dim application As Word.Application = Globals.ThisAddIn.Application
        Dim selection As Microsoft.Office.Interop.Word.Selection = application.Selection

        If selection.Type = WdSelectionType.wdSelectionIP Then
            ShowCustomMessageBox("Please select the text to be processed.")
            Return
        End If

        HostName = My.Settings.Hostname
        GuestName = My.Settings.Guestname
        TargetAudience = My.Settings.TargetAudience
        Duration = My.Settings.Duration
        Language = My.Settings.Language
        DialogueContext = My.Settings.DialogueContext
        ExtraInstructions = My.Settings.ExtraInstructions

        Dim params() As SLib.InputParameter = {
                    New SLib.InputParameter("Host name", HostName),
                    New SLib.InputParameter("Guest name", GuestName),
                    New SLib.InputParameter("Target audience", TargetAudience),
                    New SLib.InputParameter("Context, background info", DialogueContext),
                    New SLib.InputParameter("Target length", Duration),
                    New SLib.InputParameter("Language of dialogue", Language),
                    New SLib.InputParameter("Extra instructions", ExtraInstructions)
                    }

        If ShowCustomVariableInputForm("Please enter the following parameters to take into account when creating your podcast script:", $"Create Podcast Script", params) Then

            HostName = params(0).Value.ToString()
            GuestName = params(1).Value.ToString()
            TargetAudience = params(2).Value.ToString()
            DialogueContext = params(3).Value.ToString()
            Duration = params(4).Value.ToString()
            Language = params(5).Value.ToString()
            ExtraInstructions = params(6).Value.ToString()

            My.Settings.Hostname = HostName
            My.Settings.Guestname = GuestName
            My.Settings.TargetAudience = TargetAudience
            My.Settings.DialogueContext = DialogueContext
            My.Settings.Duration = Duration
            My.Settings.Language = Language
            My.Settings.ExtraInstructions = ExtraInstructions
            My.Settings.Save()

            Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Podcast), True, False, False, False, False, 3, True, False, True, False, 0, False, "", True)

        End If

    End Sub

    ' ==================== SYSTEM SLEEP LOCK MANAGEMENT ====================

    ''' <summary>
    ''' Windows API function to set thread execution state for sleep prevention.
    ''' </summary>
    ''' <param name="esFlags">Execution state flags (ES_CONTINUOUS, ES_SYSTEM_REQUIRED).</param>
    ''' <returns>Previous execution state flags.</returns>
    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function SetThreadExecutionState(ByVal esFlags As UInteger) As UInteger
    End Function

    ''' <summary>Windows API constant: Maintains execution state until explicitly reset.</summary>
    Private Const ES_CONTINUOUS As UInteger = &H80000000UI

    ''' <summary>Windows API constant: Prevents system from entering sleep (must be combined with ES_CONTINUOUS).</summary>
    Private Const ES_SYSTEM_REQUIRED As UInteger = &H1UI

    ''' <summary>Tracks whether this TTS module acquired the system sleep lock (vs. inherited from another component).</summary>
    Private Shared _ttsAcquiredTheSleepLock As Boolean = False

    ''' <summary>
    ''' Cooperatively acquires a system sleep lock for TTS operations.
    ''' Checks if a lock is already active before taking responsibility for it.
    ''' Sets _ttsAcquiredTheSleepLock flag to indicate ownership.
    ''' </summary>
    Public Shared Sub AcquireTTSSleepLock()
        ' Always request that the system stay awake.
        ' The function returns the PREVIOUS state.
        Dim previousState As UInteger = SetThreadExecutionState(ES_CONTINUOUS Or ES_SYSTEM_REQUIRED)

        ' Check if the SYSTEM_REQUIRED flag was already set in the previous state.
        If (previousState And ES_SYSTEM_REQUIRED) = 0 Then
            ' The lock was NOT active before. Therefore, the TTS engine is now responsible.
            _ttsAcquiredTheSleepLock = True
            Debug.WriteLine("[TTS] Sleep lock was not active. TTS has now acquired it.")
        Else
            ' The lock was ALREADY active. The TTS engine is not responsible for releasing it.
            _ttsAcquiredTheSleepLock = False
            Debug.WriteLine("[TTS] Sleep lock was already active. TTS will not release it.")
        End If
    End Sub

    ''' <summary>
    ''' Cooperatively releases the system sleep lock, but only if the TTS
    ''' engine was the component that originally acquired it (_ttsAcquiredTheSleepLock = True).
    ''' </summary>
    Public Shared Sub ReleaseTTSSleepLock()
        ' Only release the sleep lock IF we were the ones who set it.
        If _ttsAcquiredTheSleepLock Then
            ' We are responsible, so we release the lock.
            SetThreadExecutionState(ES_CONTINUOUS)
            _ttsAcquiredTheSleepLock = False
            Debug.WriteLine("[TTS] TTS has released the sleep lock.")
        Else
            ' We are not responsible, so we do nothing.
            Debug.WriteLine("[TTS] Another component is managing the sleep lock. TTS took no action.")
        End If
    End Sub

    ''' <summary>Enumeration of supported TTS engines.</summary>
    Public Enum TTSEngine
        Google = 0
        OpenAI = 1
    End Enum

    ''' <summary>Currently selected TTS engine for audio generation.</summary>
    Public Shared TTS_SelectedEngine As TTSEngine = TTSEngine.Google

    ''' <summary>
    ''' Detects available TTS engines by parsing INI_Endpoint and INI_TTSEndpoint configurations.
    ''' Sets availability flags (TTS_googleAvailable, TTS_openAIAvailable) and endpoint URIs
    ''' based on GoogleIdentifier/OpenAIIdentifier matches and OAuth2 flags.
    ''' </summary>
    Public Sub DetectTTSEngines()
        ' Split authentication endpoints from INI configuration
        Dim auth1 As String = ThisAddIn.INI_Endpoint
        Dim auth2 As String = ThisAddIn.INI_Endpoint_2

        ' Split TTS endpoints from INI configuration
        Dim ttsEps = If(String.IsNullOrEmpty(ThisAddIn.INI_TTSEndpoint),
                     Array.Empty(Of String)(),
                     INI_TTSEndpoint.Split("¦"c))
        Dim tts1 As String = If(ttsEps.Length > 0, ttsEps(0), "")
        Dim tts2 As String = If(ttsEps.Length > 1, ttsEps(1), "")

        ' Reset availability flags and endpoint URIs
        TTS_googleAvailable = False : TTS_googleSecondary = False
        TTS_openAIAvailable = False : TTS_openAISecondary = False
        TTS_GoogleEndpoint = "" : TTS_OpenAIEndpoint = ""

        ' Check if Google TTS is configured with OAuth2 enabled
        If auth1.Contains(GoogleIdentifier) AndAlso ThisAddIn.INI_OAuth2 Then
            TTS_googleAvailable = True
            TTS_googleSecondary = False
        End If
        If auth2.Contains(GoogleIdentifier) AndAlso ThisAddIn.INI_OAuth2_2 Then
            TTS_googleAvailable = True
            TTS_googleSecondary = True
        End If

        ' Check if OpenAI TTS is configured (no OAuth2 required)
        If auth1.Contains(OpenAIIdentifier) Then
            TTS_openAIAvailable = True
            TTS_openAISecondary = False
        End If
        If auth2.Contains(OpenAIIdentifier) Then
            TTS_openAIAvailable = True
            TTS_openAISecondary = True
        End If

        ' Assign TTS endpoint URIs based on identifier match
        If tts1.Contains(GoogleIdentifier) Then TTS_GoogleEndpoint = tts1
        If tts2.Contains(GoogleIdentifier) Then TTS_GoogleEndpoint = tts2

        If tts1.Contains(OpenAIIdentifier) Then TTS_OpenAIEndpoint = tts1
        If tts2.Contains(OpenAIIdentifier) Then TTS_OpenAIEndpoint = tts2

        ' Exit early if neither engine is properly configured
        If Not TTS_googleAvailable AndAlso Not TTS_openAIAvailable Then
            Return
        End If
    End Sub

    ''' <summary>
    ''' Determines whether to use secondary API configuration for the specified engine.
    ''' </summary>
    ''' <param name="engine">TTS engine to check.</param>
    ''' <returns>True if secondary configuration should be used.</returns>
    Private Shared Function UseSecondaryFor(engine As TTSEngine) As Boolean
        If engine = TTSEngine.Google Then
            Return TTS_googleSecondary
        Else
            Return TTS_openAISecondary
        End If
    End Function

    ' ==================== TOKEN MANAGEMENT ====================

    ' Token cache for Google OAuth2 authentication (primary API)
    Private Shared ttsAccessToken1 As String = String.Empty
    Private Shared ttsTokenExpiry1 As DateTime = DateTime.MinValue

    ' Token cache for Google OAuth2 authentication (secondary API)
    Private Shared ttsAccessToken2 As String = String.Empty
    Private Shared ttsTokenExpiry2 As DateTime = DateTime.MinValue

    ''' <summary>
    ''' Retrieves a cached OAuth2 access token or fetches a new one if expired.
    ''' </summary>
    ''' <param name="useSecond">True to use secondary API configuration (INI_*_2 settings).</param>
    ''' <returns>Valid OAuth2 bearer token, or empty string on failure.</returns>
    Private Shared Async Function GetFreshTTSToken(useSecond As Boolean) _
    As System.Threading.Tasks.Task(Of String)

        Try
            Dim token As String
            Dim expiry As DateTime

            If useSecond Then
                token = ttsAccessToken2
                expiry = ttsTokenExpiry2
            Else
                token = ttsAccessToken1
                expiry = ttsTokenExpiry1
            End If

            ' If token is missing or expired, fetch a new one
            If String.IsNullOrEmpty(token) OrElse DateTime.UtcNow >= expiry Then
                ' Select parameters based on chosen API configuration
                Dim clientEmail = If(useSecond, INI_OAuth2ClientMail_2, INI_OAuth2ClientMail)
                Dim scopes = If(useSecond, INI_OAuth2Scopes_2, INI_OAuth2Scopes)
                Dim rawKey = If(useSecond, INI_APIKey_2, INI_APIKey)
                Dim authServer = If(useSecond, INI_OAuth2Endpoint_2, INI_OAuth2Endpoint)
                Dim life = If(useSecond, INI_OAuth2ATExpiry_2, INI_OAuth2ATExpiry)

                ' Configure GoogleOAuthHelper with selected API settings
                GoogleOAuthHelper.client_email = clientEmail
                GoogleOAuthHelper.private_key = TranscriptionForm.FormatPrivateKey(rawKey)
                GoogleOAuthHelper.scopes = scopes
                GoogleOAuthHelper.token_uri = authServer
                GoogleOAuthHelper.token_life = life

                ' Fetch new OAuth2 token
                Dim newToken As String = Await GoogleOAuthHelper.GetAccessToken()
                Dim newExpiry = DateTime.UtcNow.AddSeconds(life - 300)

                If useSecond Then
                    ttsAccessToken2 = newToken
                    ttsTokenExpiry2 = newExpiry
                Else
                    ttsAccessToken1 = newToken
                    ttsTokenExpiry1 = newExpiry
                End If

                token = newToken
            End If

            Return token

        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(
            $"Error fetching TTS token: {ex.Message}",
            "TTS Error",
            System.Windows.Forms.MessageBoxButtons.OK,
            System.Windows.Forms.MessageBoxIcon.Error)
            Return String.Empty
        End Try
    End Function

    ''' <summary>Cancellation token source for aborting TTS operations.</summary>
    Public Shared cts As New CancellationTokenSource()

    ' ==================== AUDIO GENERATION - OPENAI ====================

    ''' <summary>
    ''' Generates audio via OpenAI TTS API.
    ''' </summary>
    ''' <param name="input">Text to synthesize (SSML not supported by OpenAI).</param>
    ''' <param name="languageCode">Language code (not used by OpenAI, kept for signature compatibility).</param>
    ''' <param name="voiceName">OpenAI voice name (e.g., "alloy", "echo").</param>
    ''' <param name="pitch">Pitch adjustment (not supported by OpenAI, ignored).</param>
    ''' <param name="speakingRate">Speaking rate (not supported by OpenAI, ignored).</param>
    ''' <returns>MP3 audio bytes, or Nothing on error.</returns>
    Private Shared Async Function GenerateOpenAITTSAsync(
        input As String,
        languageCode As String,
        voiceName As String,
        pitch As Double,
        speakingRate As Double
    ) As Task(Of Byte())

        Try

            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

            Dim apiKey = If(TTS_openAISecondary, DecodedAPI_2, DecodedAPI)

            Debug.WriteLine($"[TTS] OpenAI endpoint = '{TTS_OpenAIEndpoint}'")
            Debug.WriteLine($"[TTS] OpenAI model    = '{TTS_OpenAI_Model}'")
            Debug.WriteLine($"[TTS] OpenAI voice    = '{voiceName}'")
            Debug.WriteLine($"[TTS] OpenAI input    = '{Left(input, 200)}'")
            Debug.WriteLine($"[TTS] OpenAI API Key  = '{If(String.IsNullOrEmpty(apiKey), "(empty)", Left(apiKey, 8) & "...")}'")

            Using client As New System.Net.Http.HttpClient()
                client.DefaultRequestHeaders.Authorization =
                New Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiKey)

                ' Build JSON request payload for OpenAI TTS API
                Dim j = New JObject From {
                {"model", TTS_OpenAI_Model},
                {"input", input},
                {"voice", voiceName},
                {"response_format", "mp3"}
            }

                Debug.WriteLine($"[TTS] Request JSON = {j.ToString()}")

                Dim content = New StringContent(j.ToString(), Encoding.UTF8, "application/json")

                ' Send POST request to detected OpenAI TTS endpoint
                Dim resp = Await client.PostAsync(TTS_OpenAIEndpoint, content).ConfigureAwait(False)
                If resp.IsSuccessStatusCode Then
                    Return Await resp.Content.ReadAsByteArrayAsync().ConfigureAwait(False)
                Else
                    Dim err = Await resp.Content.ReadAsStringAsync().ConfigureAwait(False)
                    Debug.WriteLine($"[TTS] OpenAI error: Status={resp.StatusCode}, Headers={resp.Headers}, Body='{err}'")
                    Throw New System.Exception($"OpenAI TTS Error {resp.StatusCode}: {err}")
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show($"Error in GenerateOpenAITTSAsync: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return System.Array.Empty(Of Byte)()
    End Function

    ' ==================== AUDIO GENERATION - GOOGLE/OPENAI UNIFIED ====================

    ''' <summary>
    ''' Generates audio from text using the selected TTS engine (Google or OpenAI).
    ''' Manages authentication (OAuth2 for Google, API key for OpenAI), handles SSML wrapping,
    ''' displays progress messages for large text, and supports cancellation via cts token.
    ''' </summary>
    ''' <param name="input">Text or JSON payload to synthesize.</param>
    ''' <param name="languageCode">Language code (e.g., "en-US").</param>
    ''' <param name="voiceName">Voice identifier specific to the selected engine.</param>
    ''' <param name="nossml">True to strip SSML tags before processing.</param>
    ''' <param name="Pitch">Voice pitch adjustment (-20.0 to 20.0 for Google).</param>
    ''' <param name="SpeakingRate">Speech rate multiplier (0.25 to 4.0 for Google).</param>
    ''' <param name="CurrentPara">Current paragraph text for error reporting (optional).</param>
    ''' <returns>MP3 audio bytes, or Nothing on error/cancellation.</returns>
    Public Shared Async Function GenerateAudioFromText(input As String, Optional languageCode As String = "en-US", Optional voiceName As String = "en-US-Studio-O", Optional nossml As Boolean = False, Optional Pitch As Double = 0, Optional SpeakingRate As Double = 1, Optional CurrentPara As String = "") As Task(Of Byte())

        AcquireTTSSleepLock()

        Try

            Dim eng = TTS_SelectedEngine

            If eng = TTSEngine.OpenAI Then
                ' Extract voice identifier by removing description suffix
                Dim rawVoice = voiceName.Split(" "c)(0)
                Return Await GenerateOpenAITTSAsync(input,
                                       languageCode,
                                       rawVoice,
                                       Pitch,
                                       SpeakingRate)
            End If

            Using httpClient As New HttpClient()

                Dim AccessToken As String = Await GetFreshTTSToken(UseSecondaryFor(TTSEngine.Google))
                If String.IsNullOrEmpty(AccessToken) Then
                    ShowCustomMessageBox("Error generating audio - authentication failed (no token).")
                    Return Nothing
                End If

                httpClient.DefaultRequestHeaders.Authorization = New Net.Http.Headers.AuthenticationHeaderValue("Bearer", AccessToken)

                Dim requestBody As JObject
                Dim jsonPayload As String

                If input.Trim().StartsWith("{") Then
                    jsonPayload = input
                Else

                    Dim textlabel As String = "text"
                    Dim ssmlPattern As String = "<[^>]+>"  ' Matches any tag-like structure <...>

                    If nossml Then
                        input = Regex.Replace(input, ssmlPattern, String.Empty)
                    Else
                        If Regex.IsMatch(input, ssmlPattern) Then
                            If Not input.Trim().StartsWith("<speak>") Then
                                input = "<speak>" & input & "</speak>"
                            End If
                            textlabel = "ssml"
                        End If
                    End If

                    ' Build Google TTS request body for single-speaker synthesis
                    requestBody = New JObject From {
                    {"input", New JObject From {{$"{textlabel}", input}}},
                    {"voice", New JObject From {
                        {"languageCode", languageCode},
                        {"name", voiceName}
                    }},
                    {"audioConfig", New JObject From {
                        {"audioEncoding", "MP3"},
                        {"pitch", Pitch},
                        {"speakingRate", SpeakingRate},
                        {"effectsProfileId", New JArray("small-bluetooth-speaker-class-device")}
                    }}
                }
                    jsonPayload = requestBody.ToString()
                End If
                ' Serialize request payload to JSON
                Dim content As New StringContent(jsonPayload, Encoding.UTF8, "application/json")

                Try
                    ' Send TTS generation request to Google API

                    If Len(input) > TTSLargeText Then
                        Dim t As New Thread(Sub()
                                                ShowCustomMessageBox("Audio generation has started and runs in the background. Press 'Esc' to abort.).", "", 3, "", True)
                                            End Sub)
                        t.SetApartmentState(ApartmentState.STA)
                        t.Start()
                    End If

                    Dim response As HttpResponseMessage = Await httpClient.PostAsync(TTS_GoogleEndpoint & "text:synthesize", content, cts.Token).ConfigureAwait(False)

                    ' Validate API response
                    If response Is Nothing Then
                        ShowCustomMessageBox("Error generating audio: No response from Google TTS API.")
                        Return Nothing
                    End If

                    Dim responseString As String = Await response.Content.ReadAsStringAsync()

                    ' Log API response for debugging purposes
                    Debug.WriteLine($"Google TTS API Response: {responseString}")

                    If response.IsSuccessStatusCode Then
                        Dim responseJson As JObject = JObject.Parse(responseString)

                        ' Verify that response contains audioContent field
                        If responseJson.ContainsKey("audioContent") Then
                            Dim audioBase64 As String = responseJson("audioContent").ToString()
                            Return System.Convert.FromBase64String(audioBase64)
                        Else
                            ShowCustomMessageBox("Error generating audio: 'audioContent' not found in response.")
                            Return Nothing
                        End If
                    Else
                        ShowCustomMessageBox($"Error generating audio: API returned status {response.StatusCode}. Response: {responseString}{If(String.IsNullOrEmpty(CurrentPara), "", "Text: " & CurrentPara) & " [in clipboard]"}).")
                        If Not String.IsNullOrEmpty(CurrentPara) Then SLib.PutInClipboard(response.StatusCode & vbCrLf & vbCrLf & responseString & vbCrLf & vbCrLf & CurrentPara)
                        Return Nothing
                    End If
                Catch ex As TaskCanceledException
                    ShowCustomMessageBox("Audio generation aborted.")
                    Return Nothing
                Catch ex As Exception
                    MessageBox.Show($"Error in GenerateAudioFromText (HTTP): {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return Nothing
                End Try

            End Using
        Catch ex As Exception
            MessageBox.Show($"Error in GenerateAudioFromText: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        Finally
            ' Release sleep lock if this module acquired it
            ReleaseTTSSleepLock()
        End Try

    End Function

    ' ==================== CONVERSATION PARSING ====================

    ''' <summary>
    ''' Parses plain text dialogue into speaker/text tuples.
    ''' Recognizes host tags (hostTags) and guest tags (guestTags) at paragraph start.
    ''' Combines consecutive paragraphs from the same speaker.
    ''' </summary>
    ''' <param name="text">Multi-line dialogue text with speaker tags.</param>
    ''' <returns>List of (speaker code, text) tuples: "H" for host, "G" for guest.</returns>
    Public Function ParseTextToConversation(text As String) As List(Of Tuple(Of String, String))
        Dim conversation As New List(Of Tuple(Of String, String))
        Dim currentSpeaker As String = ""
        Dim currentText As String = ""

        Dim paragraphs As String() = text.Split({vbCrLf, vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)

        For Each para As String In paragraphs
            Dim trimmedText As String = para.Trim()
            If String.IsNullOrEmpty(trimmedText) Then Continue For

            ' Detect speaker tag at start of paragraph
            Dim newSpeaker As String = ""
            If hostTags.Any(Function(tag) trimmedText.StartsWith(tag, StringComparison.OrdinalIgnoreCase)) Then
                newSpeaker = "H"
                trimmedText = trimmedText.Substring(trimmedText.IndexOf(":"c) + 1).Trim()
            ElseIf guestTags.Any(Function(tag) trimmedText.StartsWith(tag, StringComparison.OrdinalIgnoreCase)) Then
                newSpeaker = "G"
                trimmedText = trimmedText.Substring(trimmedText.IndexOf(":"c) + 1).Trim()
            End If

            ' Store previous speaker segment and start new one when speaker changes
            If newSpeaker <> "" Then
                If Not String.IsNullOrEmpty(currentSpeaker) Then
                    conversation.Add(Tuple.Create(currentSpeaker, currentText.Trim()))
                End If
                currentSpeaker = newSpeaker
                currentText = trimmedText
            Else
                ' Append text to current speaker's dialogue
                If Not String.IsNullOrEmpty(currentSpeaker) Then
                    currentText &= " " & trimmedText
                End If
            End If
        Next

        ' Add final speaker segment to conversation list
        If Not String.IsNullOrEmpty(currentSpeaker) Then
            conversation.Add(Tuple.Create(currentSpeaker, currentText.Trim()))
        End If

        Return conversation
    End Function

    ' ==================== MULTI-SPEAKER PODCAST GENERATION ====================

    ''' <summary>
    ''' Generates multi-speaker podcast audio from parsed conversation segments.
    ''' Alternates between host and guest voices, generates audio per segment,
    ''' merges all segments into a single output file, and optionally plays the result.
    ''' </summary>
    ''' <param name="conversation">List of (speaker, text) tuples from ParseTextToConversation.</param>
    ''' <param name="filepath">Output file path for merged audio.</param>
    ''' <param name="languagecode">Language code for TTS.</param>
    ''' <param name="hostVoice">Voice identifier for host segments.</param>
    ''' <param name="guestVoice">Voice identifier for guest segments.</param>
    ''' <param name="pitch">Voice pitch adjustment.</param>
    ''' <param name="speakingrate">Speech rate multiplier.</param>
    ''' <param name="nossml">True to disable SSML processing.</param>
    Async Sub GenerateAndPlayPodcastAudio(
        conversation As List(Of Tuple(Of String, String)),
        filepath As String,
        languagecode As String,
        hostVoice As String,
        guestVoice As String,
        pitch As Double,
        speakingrate As Double,
        nossml As Boolean
    )

        Try

            Dim outputFiles As New List(Of String)

            ' Ensure valid output file path; use default if empty
            If String.IsNullOrWhiteSpace(filepath) Then
                filepath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), TTSDefaultFile)
            End If

            ' Apply default values for missing parameters
            If String.IsNullOrEmpty(languagecode) Then languagecode = "en-US"
            If String.IsNullOrEmpty(hostVoice) Then hostVoice = "en-US-Studio-O"
            If String.IsNullOrEmpty(guestVoice) Then guestVoice = "en-US-Casual-K"

            Dim Exited As Boolean = False
            Dim eng = TTS_SelectedEngine

            Using httpClient As New HttpClient()
                ' Configure HTTP client authorization based on selected TTS engine
                If eng = TTSEngine.Google Then
                    Debug.WriteLine($"[TTS] Using Google TTS engine with endpoint '{TTS_GoogleEndpoint}'")
                    ' Google: fetch OAuth token
                    Dim token = Await GetFreshTTSToken(TTS_googleSecondary)
                    If String.IsNullOrEmpty(token) Then
                        ShowCustomMessageBox("Error generating audio - authentication failed (no token).")
                        Return
                    End If
                    httpClient.DefaultRequestHeaders.Authorization =
                    New Net.Http.Headers.AuthenticationHeaderValue("Bearer", token)
                Else
                    Debug.WriteLine($"[TTS] Using OpenAI TTS engine with endpoint '{TTS_OpenAIEndpoint}'")
                    ' OpenAI: use API key
                    Dim key = If(TTS_openAISecondary, INI_APIKey_2, INI_APIKey)
                    httpClient.DefaultRequestHeaders.Authorization =
                    New Net.Http.Headers.AuthenticationHeaderValue("Bearer", key)
                End If

                ' Display background operation notification in separate thread
                Dim t As New Thread(Sub()
                                        ShowCustomMessageBox(
                                        "Audio generation has started and runs in the background. Press 'Esc' to abort.",
                                        "", 3, "", True)
                                    End Sub)
                t.SetApartmentState(ApartmentState.STA)
                t.Start()

                ' Generate audio for each speaker segment in conversation
                ' Check ESC key state: &H8000 = currently pressed, 1 = state toggled since last check
                For i = 0 To conversation.Count - 1

                    If (GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0 Then Exited = True : Exit For
                    If (GetAsyncKeyState(VK_ESCAPE) And 1) <> 0 Then Exited = True : Exit For

                    Dim speaker = conversation(i).Item1
                    Dim text = conversation(i).Item2
                    Dim voice = If(speaker = "H", hostVoice, guestVoice)

                    ' Process SSML: strip tags if disabled, wrap in <speak> if needed
                    Dim textlabel = "text"
                    If Not nossml Then
                        If Regex.IsMatch(text, "<[^>]+>") AndAlso Not text.Trim().StartsWith("<speak>") Then
                            text = $"<speak>{text}</speak>"
                            textlabel = "ssml"
                        End If
                    Else
                        text = Regex.Replace(text, "<[^>]+>", "")
                    End If

                    Dim audioBytes As Byte()

                    If eng = TTSEngine.Google Then
                        ' Generate audio via Google TTS API
                        Dim requestBody = New JObject From {
                        {"input", New JObject From {{textlabel, text}}},
                        {"voice", New JObject From {
                            {"languageCode", languagecode},
                            {"name", voice}
                        }},
                        {"audioConfig", New JObject From {
                            {"audioEncoding", "MP3"},
                            {"pitch", pitch},
                            {"speakingRate", speakingrate},
                            {"effectsProfileId", New JArray("small-bluetooth-speaker-class-device")}
                        }}
                    }

                        Dim content = New StringContent(requestBody.ToString(), Encoding.UTF8, "application/json")
                        Dim resp = Await httpClient.PostAsync(TTS_GoogleEndpoint & "text:synthesize", content)
                        Dim respStr = Await resp.Content.ReadAsStringAsync()
                        Dim respJson = JObject.Parse(respStr)

                        If respJson.ContainsKey("audioContent") Then
                            audioBytes = System.Convert.FromBase64String(respJson("audioContent").ToString())
                        Else
                            ShowCustomMessageBox("Error: no audioContent in Google response.")
                            Continue For
                        End If

                    Else
                        ' Generate audio via OpenAI TTS API
                        ' Extract voice identifier by removing description suffix
                        Dim rawVoice = voice.Split(" "c)(0)
                        audioBytes = Await GenerateOpenAITTSAsync(text, languagecode, rawVoice, pitch, speakingrate)
                    End If

                    Debug.WriteLine($"Generated audio of {audioBytes.Length} for speaker {speaker} ({voice}) with text length {text.Length} characters.")

                    ' Save generated audio segment to temporary file
                    Dim tempFile = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{AN2}_podcast_temp_{i}.mp3")
                    File.WriteAllBytes(tempFile, audioBytes)
                    outputFiles.Add(tempFile)

                    ' Rate-limit API requests to avoid overwhelming TTS service
                    Await System.Threading.Tasks.Task.Delay(1000)
                Next

                ' Merge audio segments and delete temporary files
                If Not Exited Then MergeAudioFiles(outputFiles, filepath)
                For Each f In outputFiles : File.Delete(f) : Next
            End Using

            If Exited Then
                ShowCustomMessageBox("Multi-speaker audio generation aborted.")
            Else
                If ShowCustomYesNoBox(
                    $"Your multi-speaker audio sequence has been generated ('{filepath}') and is ready to be played. Play it?",
                    "Yes", "No (file remains available)") = 1 Then
                    PlayAudio(filepath)
                End If
            End If

        Catch ex As Exception
            Debug.WriteLine($"Error generating podcast audio: {ex.Message}")
        Finally

        End Try
    End Sub


    ' ==================== AUDIO FILE MERGING ====================

    ''' <summary>
    ''' Merges multiple audio files (MP3/WAV) into a single output file.
    ''' Resamples all inputs to uniform PCM format (44.1kHz, 16-bit, stereo),
    ''' concatenates to WAV, then optionally encodes to MP3 via MediaFoundation.
    ''' </summary>
    ''' <param name="inputFiles">List of input audio file paths.</param>
    ''' <param name="outputFile">Output file path (.mp3 or .wav extension).</param>
    Public Sub MergeAudioFiles(inputFiles As System.Collections.Generic.List(Of System.String), outputFile As System.String)
        If inputFiles Is Nothing OrElse inputFiles.Count = 0 Then Throw New ArgumentException("No input files.")
        Dim take As Integer = inputFiles.Count

        ' Step 1: Concatenate all inputs to a temporary WAV file with uniform PCM format
        Dim tempWav As String = System.IO.Path.ChangeExtension(System.IO.Path.GetTempFileName(), ".wav")
        Dim targetFormat As New NAudio.Wave.WaveFormat(44100, 16, 2) ' 44.1kHz, 16-bit, stereo

        Using writer As New NAudio.Wave.WaveFileWriter(tempWav, targetFormat)
            For i = 0 To take - 1
                Dim inPath = inputFiles(i)
                If Not System.IO.File.Exists(inPath) Then Continue For

                Debug.WriteLine(inPath)

                ' AudioFileReader decodes MP3/WAV/etc. to 32-bit float
                Using src As New NAudio.Wave.AudioFileReader(inPath)
                    ' Resample/convert to the target PCM format
                    Using resampler As New NAudio.Wave.MediaFoundationResampler(src, targetFormat)
                        resampler.ResamplerQuality = 60
                        Dim buffer(8192 - 1) As Byte
                        While True
                            Dim read = resampler.Read(buffer, 0, buffer.Length)
                            If read = 0 Then Exit While
                            writer.Write(buffer, 0, read)
                        End While
                    End Using
                End Using
            Next
        End Using

        ' Step 2: Encode to MP3 if requested, otherwise deliver WAV
        Dim ext = System.IO.Path.GetExtension(outputFile)
        If String.Equals(ext, ".mp3", StringComparison.OrdinalIgnoreCase) Then
            Try
                NAudio.MediaFoundation.MediaFoundationApi.Startup()
                Using wavReader As New NAudio.Wave.WaveFileReader(tempWav)
                    ' Encode to MP3 at 192 kbps bitrate
                    NAudio.Wave.MediaFoundationEncoder.EncodeToMp3(wavReader, outputFile, 192000)
                End Using
            Catch ex As Exception
                ' Fallback: Save as WAV if Media Foundation encoder unavailable (e.g., Windows N editions)
                Dim wavFallback = System.IO.Path.ChangeExtension(outputFile, ".wav")
                System.IO.File.Copy(tempWav, wavFallback, True)
                Throw New InvalidOperationException("Media Foundation MP3 encoder unavailable. Wrote WAV instead: " & wavFallback, ex)
            Finally
                NAudio.MediaFoundation.MediaFoundationApi.Shutdown()
                Try : System.IO.File.Delete(tempWav) : Catch : End Try
            End Try
        Else
            ' Caller requested a non-MP3 extension; deliver the WAV file
            System.IO.File.Copy(tempWav, outputFile, True)
            Try : System.IO.File.Delete(tempWav) : Catch : End Try
        End If
    End Sub

    ''' <summary>
    ''' Writes audio bytes to a file.
    ''' </summary>
    ''' <param name="audioData">Audio data (typically MP3 format).</param>
    ''' <param name="filePath">Target file path.</param>
    Public Shared Sub SaveAudioToFile(audioData As Byte(), filePath As String)
        Try
            If audioData IsNot Nothing AndAlso audioData.Length > 0 Then
                File.WriteAllBytes(filePath, audioData)
                Debug.WriteLine($"Audio file saved: {filePath}")
            Else
                Debug.WriteLine("No audio received.")
            End If
        Catch ex As Exception
            Debug.WriteLine($"Error saving file: {ex.Message}")
        End Try
    End Sub

    Public Shared Sub ConfigurePlaybackOutput(waveOut As WaveOutEvent)
        If waveOut Is Nothing Then
            Return
        End If

        waveOut.DesiredLatency = 300
        waveOut.NumberOfBuffers = 4
    End Sub

    Private NotInheritable Class TimedSilenceSampleProvider
        Implements NAudio.Wave.ISampleProvider

        Private ReadOnly _waveFormat As NAudio.Wave.WaveFormat
        Private _remainingSamples As Integer

        Public Sub New(waveFormat As NAudio.Wave.WaveFormat, duration As TimeSpan)
            _waveFormat = waveFormat

            Dim totalSamples As Double =
                duration.TotalSeconds *
                CDbl(waveFormat.SampleRate) *
                CDbl(waveFormat.Channels)

            _remainingSamples = System.Math.Max(0, CInt(System.Math.Ceiling(totalSamples)))
        End Sub

        Public ReadOnly Property WaveFormat As NAudio.Wave.WaveFormat Implements NAudio.Wave.ISampleProvider.WaveFormat
            Get
                Return _waveFormat
            End Get
        End Property

        Public Function Read(buffer As Single(), offset As Integer, count As Integer) As Integer Implements NAudio.Wave.ISampleProvider.Read
            If _remainingSamples <= 0 Then
                Return 0
            End If

            Dim samplesToWrite As Integer = System.Math.Min(count, _remainingSamples)
            Array.Clear(buffer, offset, samplesToWrite)
            _remainingSamples -= samplesToWrite
            Return samplesToWrite
        End Function
    End Class

    Public Shared Function CreatePlaybackWaveProvider(audioReader As AudioFileReader,
                                                      Optional includeLeadingSilence As Boolean = True,
                                                      Optional leadingSilenceMilliseconds As Integer = 900) As IWaveProvider
        Dim sampleProvider As NAudio.Wave.ISampleProvider = audioReader

        If Not includeLeadingSilence OrElse leadingSilenceMilliseconds <= 0 Then
            Return New NAudio.Wave.SampleProviders.SampleToWaveProvider(sampleProvider)
        End If

        Dim leadingSilence As NAudio.Wave.ISampleProvider =
            New TimedSilenceSampleProvider(
                sampleProvider.WaveFormat,
                TimeSpan.FromMilliseconds(leadingSilenceMilliseconds))

        Dim playbackSequence As New NAudio.Wave.SampleProviders.ConcatenatingSampleProvider(
            New System.Collections.Generic.List(Of NAudio.Wave.ISampleProvider) From {
                leadingSilence,
                sampleProvider
            })

        Return New NAudio.Wave.SampleProviders.SampleToWaveProvider(playbackSequence)
    End Function

    ' ==================== AUDIO PLAYBACK ====================
    ''' <summary>
    ''' Plays an MP3 audio file using NAudio WaveOutEvent.
    ''' Displays splash screen with "press Esc to abort" message during playback.
    ''' </summary>
    ''' <param name="filePath">Path to MP3 file.</param>
    Public Shared Sub PlayAudio(filePath As String,
                                Optional includeLeadingSilence As Boolean = True,
                                Optional leadingSilenceMilliseconds As Integer = 900)

        Dim splash As New SLib.SplashScreen($"Playing MP3... press 'Esc' to abort")
        If File.Exists(filePath) Then
            splash.Show()
            splash.Refresh()
        End If

        Try

            If File.Exists(filePath) Then

                Using audioReader As New AudioFileReader(filePath)
                    Using waveOut As New WaveOutEvent()
                        ConfigurePlaybackOutput(waveOut)
                        waveOut.Init(CreatePlaybackWaveProvider(audioReader, includeLeadingSilence, leadingSilenceMilliseconds))
                        waveOut.Play()

                        While waveOut.PlaybackState <> PlaybackState.Stopped
                            Thread.Sleep(100)
                            System.Windows.Forms.Application.DoEvents()

                            If (GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0 Then
                                Exit While
                            End If

                            If (GetAsyncKeyState(VK_ESCAPE) And 1) <> 0 Then
                                Exit While
                            End If
                        End While

                        waveOut.Stop()
                    End Using
                End Using

                splash.Close()

            Else
                splash.Close()
                ShowCustomMessageBox("Audio file not found.")
            End If
        Catch ex As Exception
            splash.Close()
            ShowCustomMessageBox($"Error playing audio: {ex.Message}")
        End Try
    End Sub

    ' ==================== SINGLE-SPEAKER AUDIO GENERATION ====================

    ''' <summary>
    ''' Generates audio from text and optionally plays it immediately.
    ''' Creates temporary file if filepath is empty; deletes temporary file after playback.
    ''' Prompts user before playback if text length exceeds TTSLargeText threshold.
    ''' </summary>
    ''' <param name="textToSpeak">Text to synthesize.</param>
    ''' <param name="filepath">Output file path (empty for temporary file).</param>
    ''' <param name="languageCode">Language code.</param>
    ''' <param name="voiceName">Voice identifier.</param>
    Shared Async Sub GenerateAndPlayAudio(textToSpeak As String, filepath As String, Optional languageCode As String = "en-US", Optional voiceName As String = "en-US-Studio-O")

        Dim Temporary As Boolean = (filepath = "")

        Dim audioBytes As Byte() = Await System.Threading.Tasks.Task.Run(Function() GenerateAudioFromText(textToSpeak, languageCode, voiceName).Result)

        Try
            If audioBytes IsNot Nothing Then
                If Temporary Then
                    filepath = System.IO.Path.Combine(ExpandEnvironmentVariables("%TEMP%"), $"{AN2}_temp.mp3")
                End If
                SaveAudioToFile(audioBytes, filepath)
                Dim Result As Integer = 1
                If Len(textToSpeak) > TTSLargeText Then
                    Result = ShowCustomYesNoBox("Your audio sequence has been generated " & If(Temporary, "", $"('{filepath}') ") & "and is ready to be played. Play it?", "Yes", If(Temporary, "No", "No (file remains available)"))
                End If
                If Result = 1 Then
                    PlayAudio(filepath)
                End If
                If Temporary Then
                    System.IO.File.Delete(filepath)
                End If
            End If
        Catch ex As System.Exception
            ' Suppress errors to allow workflow to continue
        End Try
    End Sub

    ' ==================== PODCAST READING WITH VOICE SELECTION ====================

    ''' <summary>
    ''' Reads a podcast dialogue by parsing speaker tags, prompting for voice selection,
    ''' collecting TTS parameters (pitch, rate, SSML), and generating multi-speaker audio.
    ''' Requires both host and guest tags present in text.
    ''' </summary>
    ''' <param name="Text">Dialogue text with host/guest speaker tags.</param>
    Public Sub ReadPodcast(Text As String)

        Dim NoSSML As Boolean = My.Settings.NoSSML
        Dim Pitch As Double = My.Settings.Pitch
        Dim SpeakingRate As Double = My.Settings.Speakingrate

        ' Define TTS parameter collection for user input form
        Dim params() As SLib.InputParameter = {
                    New SLib.InputParameter("Pitch", Pitch),
                    New SLib.InputParameter("Speaking Rate", SpeakingRate),
                    New SLib.InputParameter("No SSML", NoSSML)
                    }

        Dim conversation As List(Of Tuple(Of String, String)) = ParseTextToConversation(Text)
        Dim hasHost As Boolean = conversation.Any(Function(t) t.Item1 = "H")
        Dim hasGuest As Boolean = conversation.Any(Function(t) t.Item1 = "G")

        If hasHost AndAlso hasGuest Then
            Using frm As New TTSSelectionForm("Select the voice you wish to use for creating your audio file and configure where to save it.", $"{AN} Text-to-Speech - Select Voices", True)
                If frm.ShowDialog() = DialogResult.OK Then
                    Dim selectedVoices As List(Of String) = frm.SelectedVoices
                    Dim selectedLanguage As String = frm.SelectedLanguage
                    Dim outputPath As String = frm.SelectedOutputPath

                    Debug.WriteLine("Voices=" & selectedVoices(0))
                    Debug.WriteLine("TTS_SelectedEngine=" & TTS_SelectedEngine)

                    ' Display parameter input form; update variables if user confirms
                    If ShowCustomVariableInputForm("Please enter the following parameters to apply when creating your podcast audio file:", $"Create Podcast Audio", params) Then

                        ' Update settings with user-provided values
                        Pitch = CDbl(params(0).Value)
                        SpeakingRate = CDbl(params(1).Value)
                        NoSSML = CBool(params(2).Value)

                        My.Settings.NoSSML = NoSSML
                        My.Settings.Pitch = Pitch
                        My.Settings.Speakingrate = SpeakingRate
                        My.Settings.Save()

                        GenerateAndPlayPodcastAudio(conversation, outputPath, selectedLanguage, selectedVoices(0).Replace(" (male)", "").Replace(" (female)", ""), selectedVoices(1).Replace(" (male)", "").Replace(" (female)", ""), Pitch, SpeakingRate, NoSSML)
                    End If
                End If
            End Using
        Else
            ' Validation failed: conversation must contain both host and guest segments
            ShowCustomMessageBox($"No conversation was found. Use '{hostTags(0)}' and '{guestTags(0)}' to dedicate content to the host and guest.")
        End If

    End Sub

    ' ==================== PARAGRAPH-BASED AUDIO GENERATION ====================

    ''' <summary>
    ''' Generates audio from selected Word paragraphs with advanced features:
    ''' voice alternation for titles, bullet detection, silence insertion, optional
    ''' LLM-based text cleaning, progress tracking, and ESC-based cancellation.
    ''' Saves cleaned text to .txt file when CleanText option enabled.
    ''' </summary>
    ''' <param name="filepath">Output audio file path (empty for temporary file).</param>
    ''' <param name="languageCode">Language code for TTS.</param>
    ''' <param name="voiceName">Primary voice identifier.</param>
    ''' <param name="voiceNameAlt">Alternate voice for title paragraphs (optional).</param>
    Public Async Sub GenerateAndPlayAudioFromSelectionParagraphs(filepath As String, Optional languageCode As String = "en-US", Optional voiceName As String = "en-US-Studio-O", Optional voiceNameAlt As String = "")

        Dim CurrentPara As String = ""

        Try

            Dim Temporary As Boolean = (filepath = "")
            Dim Alternate As Boolean = True

            If Temporary Then
                filepath = System.IO.Path.Combine(ExpandEnvironmentVariables("%TEMP%"), $"{AN2}_temp.mp3")
            End If

            If voiceNameAlt = "" Then Alternate = False

            ' Retrieve currently selected text from Word document
            Dim app As Word.Application = Globals.ThisAddIn.Application
            Dim selection As Microsoft.Office.Interop.Word.Selection = app.Selection
            If selection Is Nothing OrElse selection.Paragraphs.Count = 0 Then
                ShowCustomMessageBox("No text selected.")
                Return
            End If

            Dim NoSSML As Boolean = My.Settings.NoSSML
            Dim Pitch As Double = My.Settings.Pitch
            Dim SpeakingRate As Double = My.Settings.Speakingrate
            Dim ReadTitleNumbers As Boolean = False
            Dim CleanText As Boolean = False
            Dim CleanTextPrompt As String = My.Settings.CleanTextPrompt
            If String.IsNullOrWhiteSpace(CleanTextPrompt) Then CleanTextPrompt = SP_CleanTextPrompt

            ' Define TTS configuration parameters for user input
            Dim params() As SLib.InputParameter = {
                    New SLib.InputParameter("Pitch", Pitch),
                    New SLib.InputParameter("Speaking Rate", SpeakingRate),
                    New SLib.InputParameter("No SSML", NoSSML),
                    New SLib.InputParameter("Title Numbers", ReadTitleNumbers),
                    New SLib.InputParameter("Clean text", CleanText)
                    }

            ' Display parameter input form; exit if user cancels
            If Not ShowCustomVariableInputForm("Please enter the following parameters to apply when creating your audio file based on your text:", $"Create Audio", params) Then Return

            Pitch = CDbl(params(0).Value)
            SpeakingRate = CDbl(params(1).Value)
            NoSSML = CBool(params(2).Value)
            ReadTitleNumbers = CBool(params(3).Value)
            CleanText = CBool(params(4).Value)

            My.Settings.NoSSML = NoSSML
            My.Settings.Pitch = Pitch
            My.Settings.Speakingrate = SpeakingRate
            My.Settings.Save()

            If CleanText Then
                CleanTextPrompt = ShowCustomInputBox("Please enter the prompt to 'clean' the text with (each paragraph will be submitted to this prompt)", "Create Audio", False, CleanTextPrompt, Context:=_context).Trim()
                If CleanTextPrompt = "ESC" Then Return
                If CleanTextPrompt = "" Then
                    CleanText = False
                Else
                    My.Settings.CleanTextPrompt = CleanTextPrompt
                    My.Settings.Save()
                End If
            End If

            Dim totalParagraphs As Integer = selection.Paragraphs.Count
            Dim tempFiles As New List(Of String)
            Dim paragraphIndex As Integer = 0
            Dim sentenceEndPunctuation As String() = {".", "!", "?", ";", ":", ",", ")", "]", "}"}
            Dim bracketedTextPattern As String = "^\s*[\(\[\{][^\)\]\}]*[\)\]\}]\s*$"

            Dim voiceName1 As String = voiceName
            Dim voiceName2 As String = voiceNameAlt
            Dim currentVoiceName As String = voiceName1
            Dim firstTitleEncountered As Boolean = False
            Dim LastTextWasTitle As Boolean = False

            Dim cleanedTextBuilder As New System.Text.StringBuilder()

            ShowProgressBarInSeparateThread($"{AN} Audio Generation", "Starting audio generation...")
            ProgressBarModule.CancelOperation = False

            Dim silenceFileAfterBullet As String = Await GenerateSilenceAudioFileAsync(0.3)
            Dim silenceFileTitle As String = Await GenerateSilenceAudioFileAsync(0.7)
            Dim silenceFileRegular As String = Await GenerateSilenceAudioFileAsync(0.3)

            ' Iterate through selected Word paragraphs and generate audio for each
            ' Check for ESC key or cancel flag; abort and cleanup if triggered
            For Each para As Microsoft.Office.Interop.Word.Paragraph In selection.Paragraphs
                If (GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0 Or (GetAsyncKeyState(VK_ESCAPE) And 1) <> 0 Or ProgressBarModule.CancelOperation Then
                    For Each file In tempFiles
                        Try
                            If IO.File.Exists(file) Then IO.File.Delete(file)
                        Catch ex As Exception
                            Debug.WriteLine($"Error deleting temp file {file}: {ex.Message}")
                        End Try
                    Next
                    ShowCustomMessageBox("Audio generation aborted by user.")
                    ProgressBarModule.CancelOperation = True
                    Return
                End If

                ' Extract paragraph text with optional numbering prefix
                Dim paraText As String

                ' Include list numbering prefix if present and ReadTitleNumbers enabled
                If Not String.IsNullOrEmpty(para.Range.ListFormat.ListString) And ReadTitleNumbers Then
                    paraText = para.Range.ListFormat.ListString.Trim("."c) & vbCrLf & para.Range.Text.Trim()
                Else
                    paraText = para.Range.Text.Trim()
                End If

                ' Skip empty paragraphs and bracketed-only text
                If String.IsNullOrWhiteSpace(paraText) Or Regex.IsMatch(paraText, bracketedTextPattern) Then Continue For
                ' Skip paragraphs containing only digits, whitespace, or control characters
                If Regex.IsMatch(paraText, "^[\d\p{C}\s]+$") Then Continue For

                Dim lastChar As String = paraText.Substring(paraText.Length - 1)

                ' Append period if paragraph doesn't end with sentence-ending punctuation
                If Not sentenceEndPunctuation.Contains(lastChar) Then
                    paraText = paraText & "."
                End If

                ' Detect if paragraph belongs to a numbered or bulleted list
                Dim isBullet As Boolean = False
                If para.Range.ListFormat IsNot Nothing AndAlso para.Range.ListFormat.ListType <> WdListType.wdListNoNumbering Then
                    isBullet = True
                End If

                ' Detect title paragraphs based on style, line count, and punctuation
                Dim isTitle As Boolean = False
                Dim styleName As String = ""
                Try
                    Dim styleObj As Word.Style = TryCast(para.Range.Style, Word.Style)
                    If styleObj IsNot Nothing Then
                        styleName = styleObj.NameLocal.ToLowerInvariant()
                    Else
                        styleName = String.Empty
                    End If
                Catch ex As Exception
                    Debug.WriteLine("Error retrieving style: " & ex.Message)
                End Try
                If styleName.Contains("heading") Then
                    isTitle = True
                Else
                    Dim lineCount As Long = para.Range.ComputeStatistics(WdStatistic.wdStatisticLines)
                    If lineCount <= 2 Then
                        isTitle = True
                    End If
                    If Not paraText.EndsWith(".") Then
                        isTitle = True
                    End If
                End If

                Debug.WriteLine("Para = " & paraText & vbCrLf & vbCrLf)
                Debug.WriteLine("IsTitle = " & isTitle & vbCrLf)
                CurrentPara = Left(paraText, 400) & "..."

                If isTitle AndAlso Alternate Then
                    If Not firstTitleEncountered Then
                        firstTitleEncountered = True
                        ' Keep current voice for first title
                    Else
                        If Not LastTextWasTitle Then
                            ' Alternate voice for subsequent titles
                            Debug.WriteLine("Switching ...")
                            If currentVoiceName = voiceName1 Then
                                currentVoiceName = voiceName2
                            Else
                                currentVoiceName = voiceName1
                            End If
                        End If
                    End If
                    LastTextWasTitle = True
                Else
                    LastTextWasTitle = False
                End If

                ' Configure progress bar with total paragraph count
                GlobalProgressMax = totalParagraphs

                ' Update progress indicator
                GlobalProgressValue = paragraphIndex + 1
                GlobalProgressLabel = $"Paragraph {paragraphIndex + 1} of {totalParagraphs} (some may be skipped)"

                ' Insert brief silence before bullet list items
                If isBullet Then
                    Dim silenceFileBefore As String = Await GenerateSilenceAudioFileAsync(0.1)
                    If Not String.IsNullOrEmpty(silenceFileBefore) Then tempFiles.Add(silenceFileBefore)
                End If

                If CleanText Then
                    ' Clean paragraph text using LLM with user-defined prompt
                    paraText = Await LLM(CleanTextPrompt, "<TEXTTOPROCESS>" & paraText & "</TEXTTOPROCESS>", "", "", 0, False, True)
                    paraText = paraText.Trim().Replace("<TEXTTOPROCESS>", "").Replace("</TEXTTOPROCESS>", "").Trim()
                    CurrentPara = Left(CurrentPara, 100) & $"... [cleaned: {Left(paraText, 400)}...]"
                    Debug.WriteLine("Cleaned Para = " & paraText & vbCrLf & vbCrLf)

                End If

                ' Generate audio for paragraph using configured TTS engine
                Dim paragraphAudioBytes As Byte() = Await GenerateAudioFromText(paraText, languageCode, currentVoiceName, NoSSML, Pitch, SpeakingRate, CurrentPara)

                CurrentPara = ""

                If paragraphAudioBytes IsNot Nothing Then
                    If CleanText Then
                        cleanedTextBuilder.AppendLine(paraText)
                        cleanedTextBuilder.AppendLine() ' Empty line between paragraphs
                    End If
                    Dim tempParaFile As String = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{AN2}_temp_para_{paragraphIndex}.mp3")
                    File.WriteAllBytes(tempParaFile, paragraphAudioBytes)
                    tempFiles.Add(tempParaFile)
                    Debug.WriteLine("Created " & tempParaFile)
                Else
                    ' Skip paragraph if audio generation failed
                    Debug.WriteLine("Creation failed")
                    Continue For
                End If

                ' Insert brief silence after bullet list items
                If isBullet Then
                    If Not String.IsNullOrEmpty(silenceFileAfterBullet) Then tempFiles.Add(silenceFileAfterBullet)
                End If

                ' Insert context-appropriate silence after each paragraph:
                ' Titles receive 0.7 sec pause; regular text receives 0.3 sec pause
                If isTitle Then
                    If Not String.IsNullOrEmpty(silenceFileTitle) Then tempFiles.Add(silenceFileTitle)
                Else
                    If Not String.IsNullOrEmpty(silenceFileRegular) Then tempFiles.Add(silenceFileRegular)
                End If

                ' Rate-limit API requests to avoid overwhelming TTS service
                Await System.Threading.Tasks.Task.Delay(1000)

                paragraphIndex += 1
            Next

            ' Exit if no paragraphs generated audio (all skipped)
            If tempFiles.Count = 0 Then
                ShowCustomMessageBox("No valid paragraphs found For audio generation; skipping empty ones and {...}, [...] and (...).")
                Return
            End If

            If Not ProgressBarModule.CancelOperation Then
                ' Merge all audio segments into final output file
                GlobalProgressLabel = $"Merging audio {totalParagraphs} snippets..."
                MergeAudioFiles(tempFiles, filepath)
            End If

            If Not ProgressBarModule.CancelOperation AndAlso CleanText Then
                Try
                    Dim txtPath As String = System.IO.Path.ChangeExtension(filepath, ".txt")
                    System.IO.File.WriteAllText(txtPath, cleanedTextBuilder.ToString(), System.Text.Encoding.UTF8)
                Catch ex As System.Exception
                    ' Suppress errors to allow workflow to continue
                    Debug.WriteLine("Error writing cleaned text file: " & ex.Message)
                End Try
            End If

            ' Delete all temporary audio segment files
            For Each file In tempFiles
                Try
                    If IO.File.Exists(file) Then IO.File.Delete(file)
                Catch ex As Exception
                    Debug.WriteLine($"Error deleting temp file {file}: {ex.Message}")
                End Try
            Next

            If Not ProgressBarModule.CancelOperation Then
                ProgressBarModule.CancelOperation = True
                ' Play final merged audio and cleanup temporary file if applicable
                PlayAudio(filepath)
                If Temporary Then
                    System.IO.File.Delete(filepath)
                End If
            Else
                ProgressBarModule.CancelOperation = True
                ShowCustomMessageBox("Audio generation aborted by user.")
            End If

        Catch ex As Exception
            ShowCustomMessageBox($"Error generating audio from selected paragraphs ({ex.Message}{If(String.IsNullOrEmpty(CurrentPara), "", "; Text: " & CurrentPara) & " [in clipboard]"}).")
            If Not String.IsNullOrEmpty(CurrentPara) Then SLib.PutInClipboard(ex.Message & vbCrLf & vbCrLf & CurrentPara)
        End Try
    End Sub

    ' ==================== SILENCE GENERATION ====================

    ''' <summary>
    ''' Asynchronously generates a silence audio file of specified duration.
    ''' Wrapper for synchronous GenerateSilenceAudioFile.
    ''' </summary>
    ''' <param name="durationSeconds">Duration of silence in seconds.</param>
    ''' <returns>Path to generated MP3 silence file, or Nothing on error.</returns>
    Private Async Function GenerateSilenceAudioFileAsync(durationSeconds As Double) As Task(Of String)
        Return Await System.Threading.Tasks.Task.Run(Function() GenerateSilenceAudioFile(durationSeconds))
    End Function

    ''' <summary>
    ''' Creates a PCM buffer filled with zeros (silence) and encodes it to MP3.
    ''' Uses 24kHz, 16-bit, mono format by default.
    ''' </summary>
    ''' <param name="durationSeconds">Duration of silence in seconds.</param>
    ''' <returns>Path to generated MP3 file, or Nothing on error.</returns>
    Private Function GenerateSilenceAudioFile(durationSeconds As Double) As String
        Try
            ' Configure audio format for silence generation (24kHz, 16-bit, mono)
            Dim sampleRate As Integer = 24000
            Dim channels As Integer = 1
            Dim bitsPerSample As Integer = 16
            Dim blockAlign As Integer = channels * (bitsPerSample \ 8)
            Dim totalSamples As Integer = CInt(sampleRate * durationSeconds)
            Dim totalBytes As Integer = totalSamples * blockAlign

            ' Allocate byte buffer initialized to zeros (silence)
            Dim silenceBytes(totalBytes - 1) As Byte

            ' Generate temporary file path with duration-based naming
            Dim tempFile As String = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{AN2}_silence_{CInt(durationSeconds * 1000)}ms.mp3")

            ' Create WAV stream from silence buffer and encode to MP3
            Using ms As New MemoryStream(silenceBytes)
                Dim waveFormat As New WaveFormat(sampleRate, bitsPerSample, channels)
                Using waveStream As New RawSourceWaveStream(ms, waveFormat)
                    MediaFoundationEncoder.EncodeToMp3(waveStream, tempFile)
                End Using
            End Using

            Return tempFile
        Catch ex As Exception
            Debug.WriteLine($"Error generating silence audio: {ex.Message}")
            Return Nothing
        End Try
    End Function

    ' ==================== LEGACY WINDOWS TTS (SAPI) ====================

    ''' <summary>Legacy Windows SAPI speech synthesizer for fallback TTS.</summary>
    Private synth As New SpeechSynthesizer()

    ''' <summary>
    ''' Displays list of installed Windows SAPI voices, prompts user to select by number,
    ''' speaks a confirmation message using the selected voice, and saves selection to settings.
    ''' </summary>
    Public Shared Sub SelectVoiceByNumber()
        ' Initialize SpeechSynthesizer instance
        Dim synth As New SpeechSynthesizer()

        ' Enumerate all installed SAPI voices on the system
        Dim installedVoices As List(Of InstalledVoice) = synth.GetInstalledVoices().ToList()
        Dim voiceNames As New List(Of String)()

        ' Build voice selection menu with numeric indices
        Dim sb As New StringBuilder()
        sb.AppendLine("Available voices for Text-to-Speech:" & vbCrLf)

        For i As Integer = 0 To installedVoices.Count - 1
            Dim voiceInfo As VoiceInfo = installedVoices(i).VoiceInfo
            voiceNames.Add(voiceInfo.Name)
            sb.AppendLine($"{i}: {voiceInfo.Name}")
        Next

        If voiceNames.Count = 0 Then
            ShowCustomMessageBox("No voices available on this system.", "Text-to-Speech")
            Return
        End If

        Dim UserInput As String = ShowCustomInputBox(sb.ToString(), "Select Voice for Text Reader", True)

        If String.IsNullOrWhiteSpace(UserInput) Then Return

        Dim selectedIndex As Integer
        If Integer.TryParse(UserInput, selectedIndex) AndAlso selectedIndex >= 0 AndAlso selectedIndex < voiceNames.Count Then
            ' Apply selected voice and speak confirmation message
            Dim chosenVoice As String = voiceNames(selectedIndex)
            Try
                synth.SelectVoice(chosenVoice)
                My.Settings.LastVoice = chosenVoice
                My.Settings.Save()

                synth.Speak($"Hello! I am now using the voice: {chosenVoice}")
            Catch ex As Exception
                ShowCustomMessageBox("Error selecting voice: " & ex.Message)
            End Try
        Else
            ShowCustomMessageBox("Invalid voice number entered.", "Text-to-Speech")
        End If
    End Sub

    ''' <summary>
    ''' Reads selected Word text aloud using legacy Windows SpeechSynthesizer.
    ''' Cancels ongoing speech if already speaking. Uses voice from My.Settings.LastVoice.
    ''' </summary>
    Public Sub SpeakSelectedText()

        Debug.WriteLine("Status: " & synth.State.ToString())

        If synth.State = SynthesizerState.Speaking Then
            synth.SpeakAsyncCancelAll()
            ShowCustomMessageBox("Reading out aborted.", "Text-to-Speech")
            Return
        End If

        Try
            ' Retrieve active Word application instance
            Dim wordApp As Word.Application = Globals.ThisAddIn.Application

            ' Extract selected text from document
            Dim selectedText As String = wordApp.Selection.Text.Trim()

            If String.IsNullOrEmpty(selectedText) Then
                ShowCustomMessageBox("No text selected in Word.", "Text-to-Speech")
                Return
            End If

            ' Initiate asynchronous speech synthesis with saved voice
            synth.SelectVoice(My.Settings.LastVoice)
            synth.SpeakAsync(selectedText)

            ShowCustomMessageBox($"Reading out the selected text (using {My.Settings.LastVoice}). You can stop this by again calling this function.", "Text-to-Speech")

        Catch ex As Exception
            MessageBox.Show("Error in SpeakSelectedText: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


End Class
