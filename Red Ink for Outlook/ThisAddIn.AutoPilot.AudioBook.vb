' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.AutoPilot.AudioBook.vb
' Purpose: AutoPilot tool that generates audio files (podcast or audiobook)
'          from email body text or document attachments using Google Cloud TTS
'          or OpenAI TTS. Supports multi-speaker podcast dialogues (via LLM
'          script generation) and single/alternating-voice audiobook narration.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TTS — CONSTANTS & STATE
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Const AB_DefaultFile As String = "audiobook.mp3"
    Private Const AB_GoogleIdentifier As String = "googleapis.com"
    Private Const AB_OpenAIIdentifier As String = "openai.com"
    Private Const AB_OpenAI_Model As String = "tts-1-hd"
    Private Const AB_OpenAI_FallbackEndpoint As String = "https://api.openai.com/v1/audio/speech"

    ''' <summary>Max characters per TTS API call to avoid timeouts.</summary>
    Private Const AB_MaxCharsPerSegment As Integer = 4000

    ''' <summary>Duration of silence inserted between same-speaker paragraphs in audiobook mode (seconds).</summary>
    Private Const AB_SilenceBetweenParagraphs As Double = 0.9

    ''' <summary>Duration of silence inserted when the narrator voice changes in audiobook mode (seconds).</summary>
    Private Const AB_SilenceOnSpeakerChange As Double = 1.6

    ''' <summary>Duration of silence inserted between speaker turns in podcast mode (seconds).</summary>
    Private Const AB_SilenceBetweenPodcastTurns As Double = 0.8

    ''' <summary>Duration of silence inserted after title/heading segments (seconds).</summary>
    Private Const AB_SilenceAfterTitle As Double = 0.7

    ''' <summary>Maximum number of retries when a TTS segment returns no audio.</summary>
    Private Const AB_MaxSegmentRetries As Integer = 3

    Private Shared AB_googleAvailable As Boolean = False
    Private Shared AB_googleSecondary As Boolean = False
    Private Shared AB_openAIAvailable As Boolean = False
    Private Shared AB_openAISecondary As Boolean = False
    Private Shared AB_GoogleEndpoint As String = ""
    Private Shared AB_OpenAIEndpoint As String = ""

    Private Enum ABEngine
        Google = 0
        OpenAI = 1
    End Enum

    Private Shared AB_SelectedEngine As ABEngine = ABEngine.OpenAI

    Private Shared ReadOnly AB_DefaultVoiceA_OpenAI As String = "alloy"
    Private Shared ReadOnly AB_DefaultVoiceB_OpenAI As String = "nova"
    Private Shared ReadOnly AB_DefaultLanguage As String = "en-US"

    Private Shared ReadOnly AB_HostTags As String() = {"H:", "Host:", "A:", "1:"}
    Private Shared ReadOnly AB_GuestTags As String() = {"G:", "Guest:", "Gast:", "B:", "2:"}


    Private Const AB_GoogleFemaleVoiceStem As String = "Chirp3-HD-Achernar"
    Private Const AB_GoogleMaleVoiceStem As String = "Chirp3-HD-Achird"
    Private Const AB_OpenAIFemaleVoice As String = "nova"
    Private Const AB_OpenAIMaleVoice As String = "ash"


    Private Shared AB_AccessToken1 As String = String.Empty
    Private Shared AB_TokenExpiry1 As DateTime = DateTime.MinValue
    Private Shared AB_AccessToken2 As String = String.Empty
    Private Shared AB_TokenExpiry2 As DateTime = DateTime.MinValue

    ''' <summary>
    ''' Cached API key captured at TTS engine detection time.
    ''' Avoids reading DecodedAPI/DecodedAPI_2 at execution time when the
    ''' tooling loop may have overwritten them with a different model's key.
    ''' </summary>
    Private Shared AB_CachedApiKey As String = ""

    ''' <summary>
    ''' Cached Google OAuth2 parameters captured at TTS engine detection time.
    ''' The tooling loop's ApplyModelConfig overwrites the _context INI_OAuth2*
    ''' properties with the LLM model's values (e.g. Gemini), wiping the TTS
    ''' service account credentials. These cached copies survive that mutation.
    ''' </summary>
    Private Shared AB_CachedOAuth2ClientMail As String = ""
    Private Shared AB_CachedOAuth2Scopes As String = ""
    Private Shared AB_CachedOAuth2RawKey As String = ""
    Private Shared AB_CachedOAuth2Endpoint As String = ""
    Private Shared AB_CachedOAuth2ATExpiry As Long = 3600

    ''' <summary>
    ''' A minimal valid MP3 frame encoding silence (MPEG1 Layer 3, 128 kbps,
    ''' 44100 Hz, mono). Duration ≈ 26 ms per frame. We repeat this to build
    ''' arbitrary-length silence without requiring NAudio or MediaFoundation.
    ''' </summary>
    Private Shared ReadOnly AB_SilentMp3Frame As Byte() = {
        &HFF, &HFB, &H90, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
        &H0, &H0, &H0
    }


    Private Shared Function AB_NormalizeLanguageForStandardAudio(languageCode As String) As String
        Dim lang = If(String.IsNullOrWhiteSpace(languageCode), AB_DefaultLanguage, languageCode.Trim())

        If lang.StartsWith("de", StringComparison.OrdinalIgnoreCase) Then
            Return "de-DE"
        End If

        Return lang
    End Function

    Private Shared Function AB_BuildStandardLanguageInstruction(languageCode As String) As String
        Dim lang = If(String.IsNullOrWhiteSpace(languageCode), AB_DefaultLanguage, languageCode.Trim())
        Dim twoLetter = If(lang.Length > 2, lang.Substring(0, 2), lang)

        If twoLetter.Equals("de", StringComparison.OrdinalIgnoreCase) Then
            Return "[IMPORTANT: Write the script in Standard German (Hochdeutsch). Never use Swiss German, Alemannic, Austrian dialect, regional dialect, dialect spelling, or dialect vocabulary, even if the locale is Switzerland, Austria, Liechtenstein, or any other German-speaking region.]"
        End If

        Return $"[IMPORTANT: Write the script in standard {twoLetter} language. Do not use {twoLetter} dialects.]"
    End Function


    Private Shared Function AB_NormalizeSpeakerGender(gender As String, fallbackGender As String) As String
        Dim g = If(gender, "").Trim().ToLowerInvariant()

        Select Case g
            Case "f", "female", "woman", "w", "weiblich", "frau"
                Return "female"
            Case "m", "male", "man", "mann", "männlich"
                Return "male"
            Case Else
                Return fallbackGender
        End Select
    End Function

    Private Shared Function AB_BuildGoogleVoiceNameForGender(languageCode As String, gender As String) As String
        Dim lang = AB_NormalizeLanguageForStandardAudio(languageCode)
        Dim normalizedGender = AB_NormalizeSpeakerGender(gender, "female")

        If normalizedGender.Equals("male", StringComparison.OrdinalIgnoreCase) Then
            Return lang & "-" & AB_GoogleMaleVoiceStem
        End If

        Return lang & "-" & AB_GoogleFemaleVoiceStem
    End Function

    Private Shared Function AB_IsLikelyGoogleVoiceName(voiceName As String) As Boolean
        If String.IsNullOrWhiteSpace(voiceName) Then Return False

        Return Regex.IsMatch(
        voiceName.Trim(),
        "^[a-z]{2,3}-[A-Z]{2}-[A-Za-z0-9]+(?:-[A-Za-z0-9]+)*$",
        RegexOptions.CultureInvariant)
    End Function

    Private Shared Function AB_RewriteGoogleVoiceLanguagePrefix(voiceName As String, languageCode As String) As String
        If String.IsNullOrWhiteSpace(voiceName) Then Return ""

        Dim normalizedLanguage = AB_NormalizeLanguageForStandardAudio(languageCode)
        Dim m = Regex.Match(voiceName.Trim(), "^[a-z]{2,3}-[A-Z]{2}-(.+)$", RegexOptions.CultureInvariant)

        If m.Success Then
            Return normalizedLanguage & "-" & m.Groups(1).Value
        End If

        Return voiceName.Trim()
    End Function

    Private Shared Function AB_IsOpenAIVoiceName(voiceName As String) As Boolean
        Dim v = If(voiceName, "").Trim().ToLowerInvariant()
        If v = "" Then Return False

        Dim validVoices As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
        "alloy", "ash", "ballad", "coral", "echo", "fable", "nova", "onyx", "sage", "shimmer", "verse"
    }

        Return validVoices.Contains(v)
    End Function

    Private Shared Function AB_ResolveVoiceNameForEngine(voiceName As String, languageCode As String, gender As String) As String
        If AB_SelectedEngine = ABEngine.Google Then
            If AB_IsLikelyGoogleVoiceName(voiceName) Then
                Return AB_RewriteGoogleVoiceLanguagePrefix(voiceName, languageCode)
            End If

            Return AB_BuildGoogleVoiceNameForGender(languageCode, gender)
        End If

        If AB_IsOpenAIVoiceName(voiceName) Then
            Return voiceName.Trim().ToLowerInvariant()
        End If

        If AB_NormalizeSpeakerGender(gender, "female").Equals("male", StringComparison.OrdinalIgnoreCase) Then
            Return AB_OpenAIMaleVoice
        End If

        Return AB_OpenAIFemaleVoice
    End Function

    Private Shared Sub AB_ResolveAudioSpeakerAndVoiceInputs(
    mode As String,
    languageCode As String,
    ByRef hostName As String,
    ByRef guestName As String,
    ByRef hostGender As String,
    ByRef guestGender As String,
    ByRef voiceA As String,
    ByRef voiceB As String,
    Optional singleVoice As Boolean = False,
    Optional narratorGender As String = Nothing)

        Dim isPodcast = String.Equals(mode, "podcast", StringComparison.OrdinalIgnoreCase)

        If isPodcast Then
            If String.IsNullOrWhiteSpace(hostName) Then hostName = "Marie"
            If String.IsNullOrWhiteSpace(guestName) Then guestName = "Marc"

            hostGender = AB_NormalizeSpeakerGender(hostGender, "female")
            guestGender = AB_NormalizeSpeakerGender(guestGender, "male")

            voiceA = AB_ResolveVoiceNameForEngine(voiceA, languageCode, hostGender)
            voiceB = AB_ResolveVoiceNameForEngine(voiceB, languageCode, guestGender)
            Return
        End If

        Dim effectiveNarratorGender = AB_NormalizeSpeakerGender(narratorGender, "female")

        If singleVoice Then
            voiceA = AB_ResolveVoiceNameForEngine(voiceA, languageCode, effectiveNarratorGender)
            voiceB = voiceA
            Return
        End If

        hostGender = AB_NormalizeSpeakerGender(hostGender, "female")
        guestGender = AB_NormalizeSpeakerGender(guestGender, "male")

        voiceA = AB_ResolveVoiceNameForEngine(voiceA, languageCode, hostGender)
        voiceB = AB_ResolveVoiceNameForEngine(voiceB, languageCode, guestGender)
    End Sub


    ' ═══════════════════════════════════════════════════════════════════════════
    '  TTS — ENGINE DETECTION
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Sub AB_DetectTTSEngines()
        Dim auth1 As String = If(INI_Endpoint, "")
        Dim auth2 As String = If(INI_Endpoint_2, "")

        Dim rawTts As String = If(INI_TTSEndpoint, "")
        Dim ttsEps = If(String.IsNullOrEmpty(rawTts),
                     Array.Empty(Of String)(),
                     rawTts.Split("¦"c))
        Dim tts1 As String = If(ttsEps.Length > 0, ttsEps(0), "")
        Dim tts2 As String = If(ttsEps.Length > 1, ttsEps(1), "")

        AB_googleAvailable = False : AB_googleSecondary = False
        AB_openAIAvailable = False : AB_openAISecondary = False
        AB_GoogleEndpoint = "" : AB_OpenAIEndpoint = ""

        If auth1.Contains(AB_GoogleIdentifier) AndAlso INI_OAuth2 Then
            AB_googleAvailable = True : AB_googleSecondary = False
        End If
        If auth2.Contains(AB_GoogleIdentifier) AndAlso INI_OAuth2_2 Then
            AB_googleAvailable = True : AB_googleSecondary = True
        End If
        If auth1.Contains(AB_OpenAIIdentifier) Then
            AB_openAIAvailable = True : AB_openAISecondary = False
        End If
        If auth2.Contains(AB_OpenAIIdentifier) Then
            AB_openAIAvailable = True : AB_openAISecondary = True
        End If

        If tts1.Contains(AB_GoogleIdentifier) Then AB_GoogleEndpoint = tts1
        If tts2.Contains(AB_GoogleIdentifier) Then AB_GoogleEndpoint = tts2
        If tts1.Contains(AB_OpenAIIdentifier) Then AB_OpenAIEndpoint = tts1
        If tts2.Contains(AB_OpenAIIdentifier) Then AB_OpenAIEndpoint = tts2

        ' Fallback: if OpenAI is available (API key present) but no dedicated
        ' TTS endpoint was configured in INI_TTSEndpoint, use the standard URL.
        If AB_openAIAvailable AndAlso String.IsNullOrWhiteSpace(AB_OpenAIEndpoint) Then
            AB_OpenAIEndpoint = AB_OpenAI_FallbackEndpoint
        End If

        ' Default engine selection: prefer Google for non-English languages
        ' because Google TTS produces native-sounding voices per language,
        ' whereas OpenAI voices have an English accent regardless of input.
        ' The caller can override AB_SelectedEngine after detection.
        If AB_openAIAvailable Then
            AB_SelectedEngine = ABEngine.OpenAI
        ElseIf AB_googleAvailable Then
            AB_SelectedEngine = ABEngine.Google
        End If

        ' Cache the correct API key at detection time so it survives
        ' ApplyModelConfig mutations during the tooling loop.
        If AB_openAIAvailable Then
            AB_CachedApiKey = If(AB_openAISecondary, DecodedAPI_2, DecodedAPI)
        ElseIf AB_googleAvailable Then
            AB_CachedApiKey = If(AB_googleSecondary, DecodedAPI_2, DecodedAPI)
        Else
            AB_CachedApiKey = ""
        End If

        ' Cache Google OAuth2 parameters at detection time. The tooling loop's
        ' ApplyModelConfig overwrites the _context INI_OAuth2* properties with
        ' the LLM model's values (e.g. Gemini), wiping the TTS service account
        ' credentials that AB_GetFreshToken needs.
        If AB_googleAvailable Then
            If AB_googleSecondary Then
                AB_CachedOAuth2ClientMail = If(INI_OAuth2ClientMail_2, "")
                AB_CachedOAuth2Scopes = If(INI_OAuth2Scopes_2, "")
                AB_CachedOAuth2RawKey = If(INI_APIKey_2, "")
                AB_CachedOAuth2Endpoint = If(INI_OAuth2Endpoint_2, "")
                AB_CachedOAuth2ATExpiry = INI_OAuth2ATExpiry_2
            Else
                AB_CachedOAuth2ClientMail = If(INI_OAuth2ClientMail, "")
                AB_CachedOAuth2Scopes = If(INI_OAuth2Scopes, "")
                AB_CachedOAuth2RawKey = If(INI_APIKey, "")
                AB_CachedOAuth2Endpoint = If(INI_OAuth2Endpoint, "")
                AB_CachedOAuth2ATExpiry = INI_OAuth2ATExpiry
            End If
        Else
            AB_CachedOAuth2ClientMail = ""
            AB_CachedOAuth2Scopes = ""
            AB_CachedOAuth2RawKey = ""
            AB_CachedOAuth2Endpoint = ""
            AB_CachedOAuth2ATExpiry = 3600
        End If
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TTS — ENGINE SELECTION HELPER
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Selects the best TTS engine based on the user/LLM request and language.
    ''' When both engines are available:
    '''   - An explicit "google"/"openai" preference is honoured.
    '''   - For non-English languages, Google is preferred (native-accent voices).
    '''   - For English (or unknown language), OpenAI is preferred.
    ''' When only one engine is available it is always used.
    ''' </summary>
    Private Shared Sub AB_SelectBestEngine(enginePreference As String, language As String)
        ' If only one engine is available, there is nothing to choose.
        If AB_openAIAvailable AndAlso Not AB_googleAvailable Then
            AB_SelectedEngine = ABEngine.OpenAI
            Return
        End If
        If AB_googleAvailable AndAlso Not AB_openAIAvailable Then
            AB_SelectedEngine = ABEngine.Google
            Return
        End If

        ' Both engines are available — check for explicit preference first.
        If Not String.IsNullOrWhiteSpace(enginePreference) Then
            Dim pref = enginePreference.Trim().ToLowerInvariant()
            If pref.Contains("google") Then
                AB_SelectedEngine = ABEngine.Google
                Return
            End If
            If pref.Contains("openai") Then
                AB_SelectedEngine = ABEngine.OpenAI
                Return
            End If
        End If

        ' Auto-select based on language: Google produces better non-English accents.
        Dim lang = If(String.IsNullOrWhiteSpace(language), "en", language.Trim().ToLowerInvariant())
        If lang.StartsWith("en") Then
            AB_SelectedEngine = ABEngine.OpenAI
        Else
            AB_SelectedEngine = ABEngine.Google
        End If
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TTS — PRIVATE KEY FORMATTING (mirrors Word TranscriptionForm helper)
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Formats a raw private key string into PEM format with proper headers
    ''' and 64-character line breaks required by BouncyCastle.
    ''' </summary>
    Private Shared Function AB_FormatPrivateKey(rawKey As String) As String
        Dim noEscapes = rawKey.Replace("\n", "")
        Dim sb As New StringBuilder()
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

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TTS — TOKEN MANAGEMENT (Google OAuth2)
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Obtains a fresh Google OAuth2 access token for TTS, using the cached
    ''' OAuth2 parameters that were captured at engine detection time.
    ''' This avoids reading the live INI_OAuth2* properties which may have been
    ''' overwritten by ApplyModelConfig during the tooling loop.
    ''' </summary>
    Private Shared Async Function AB_GetFreshToken(useSecond As Boolean) As Task(Of String)
        Try
            Dim token = If(useSecond, AB_AccessToken2, AB_AccessToken1)
            Dim expiry = If(useSecond, AB_TokenExpiry2, AB_TokenExpiry1)

            If String.IsNullOrEmpty(token) OrElse DateTime.UtcNow >= expiry Then
                ' Use the cached OAuth2 parameters instead of the live INI_* properties.
                ' The live properties are overwritten by the tooling loop's ApplyModelConfig
                ' with the LLM model's config (e.g. Gemini), wiping the TTS service account.
                Dim clientEmail = AB_CachedOAuth2ClientMail
                Dim scopes = AB_CachedOAuth2Scopes
                Dim rawKey = AB_CachedOAuth2RawKey
                Dim authServer = AB_CachedOAuth2Endpoint
                Dim life = AB_CachedOAuth2ATExpiry

                If String.IsNullOrWhiteSpace(clientEmail) OrElse String.IsNullOrWhiteSpace(rawKey) OrElse
                   String.IsNullOrWhiteSpace(authServer) Then
                    Debug.WriteLine($"[AB-TTS] Google OAuth2 cached credentials are empty — cannot obtain token. " &
                                    $"clientEmail='{If(Not String.IsNullOrWhiteSpace(clientEmail), "(set)", "(empty)")}' " &
                                    $"rawKey='{If(Not String.IsNullOrWhiteSpace(rawKey), "(set)", "(empty)")}' " &
                                    $"authServer='{If(Not String.IsNullOrWhiteSpace(authServer), "(set)", "(empty)")}'")
                    Return String.Empty
                End If

                GoogleOAuthHelper.client_email = clientEmail
                GoogleOAuthHelper.private_key = AB_FormatPrivateKey(rawKey)
                GoogleOAuthHelper.scopes = scopes
                GoogleOAuthHelper.token_uri = authServer
                GoogleOAuthHelper.token_life = life

                Dim newToken = Await GoogleOAuthHelper.GetAccessToken()
                Dim newExpiry = DateTime.UtcNow.AddSeconds(life - 300)

                If useSecond Then
                    AB_AccessToken2 = newToken : AB_TokenExpiry2 = newExpiry
                Else
                    AB_AccessToken1 = newToken : AB_TokenExpiry1 = newExpiry
                End If
                token = newToken
            End If
            Return token
        Catch ex As Exception
            Debug.WriteLine($"[AB-TTS] Token error: {ex.Message}")
            Return String.Empty
        End Try
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TTS — SILENCE GENERATION (lightweight, no NAudio dependency)
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Creates a short silent MP3 file by repeating a minimal valid MP3 frame.
    ''' Each frame ≈ 26 ms at 128 kbps / 44100 Hz mono.
    ''' </summary>
    Private Shared Function AB_GenerateSilenceFile(durationSeconds As Double, tempDir As String) As String
        Try
            ' Each frame is ~26.12 ms; compute how many frames we need.
            Dim frameCount = CInt(Math.Ceiling(durationSeconds / 0.02612))
            If frameCount < 1 Then frameCount = 1

            Dim filePath = Path.Combine(tempDir, $"ab_silence_{CInt(durationSeconds * 1000)}ms.mp3")

            Using fs As New FileStream(filePath, FileMode.Create, FileAccess.Write)
                For i = 0 To frameCount - 1
                    fs.Write(AB_SilentMp3Frame, 0, AB_SilentMp3Frame.Length)
                Next
            End Using

            Return filePath
        Catch ex As Exception
            Debug.WriteLine($"[AB-TTS] Silence generation error: {ex.Message}")
            Return Nothing
        End Try
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TTS — SINGLE SEGMENT AUDIO GENERATION
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Shared Async Function AB_GenerateSegmentAsync(
        text As String,
        languageCode As String,
        voiceName As String,
        Optional ct As System.Threading.CancellationToken = Nothing) As System.Threading.Tasks.Task(Of Byte())

        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

        If AB_SelectedEngine = ABEngine.OpenAI Then
            Dim endpoint = AB_OpenAIEndpoint
            If String.IsNullOrWhiteSpace(endpoint) Then
                Debug.WriteLine("[AB-TTS] OpenAI endpoint is empty — cannot generate segment.")
                Return Nothing
            End If

            ' Use the cached API key that was captured at TTS engine detection time
            ' (before the tooling loop's ApplyModelConfig could overwrite DecodedAPI/DecodedAPI_2
            ' with the LLM model's key). Only fall back to the live properties if the cache is empty.
            Dim apiKey = AB_CachedApiKey
            If String.IsNullOrWhiteSpace(apiKey) Then
                apiKey = If(AB_openAISecondary, DecodedAPI_2, DecodedAPI)
            End If
            If String.IsNullOrWhiteSpace(apiKey) Then
                Debug.WriteLine("[AB-TTS] OpenAI API key is empty (both cached and live).")
                Return Nothing
            End If

            If String.IsNullOrWhiteSpace(text) Then
                Debug.WriteLine("[AB-TTS] Input text is empty.")
                Return Nothing
            End If

            If text.Length > 4096 Then
                Debug.WriteLine($"[AB-TTS] Input too long for OpenAI speech API: {text.Length} chars.")
                Return Nothing
            End If

            Dim safeVoice As String = voiceName.Split(" "c)(0).Trim().ToLowerInvariant()

            Using client As New System.Net.Http.HttpClient()
                client.Timeout = TimeSpan.FromSeconds(120)
                client.DefaultRequestHeaders.Authorization =
                New Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiKey)

                Dim j = New JObject From {
                {"model", AB_OpenAI_Model},
                {"input", text},
                {"voice", safeVoice},
                {"response_format", "mp3"}
            }

                Dim content = New System.Net.Http.StringContent(j.ToString(), Encoding.UTF8, "application/json")

                Debug.WriteLine($"[AB-TTS] POST {endpoint} model={AB_OpenAI_Model} voice={safeVoice} len={text.Length}")

                Dim resp = Await client.PostAsync(endpoint, content, ct).ConfigureAwait(False)

                If resp.IsSuccessStatusCode Then
                    Return Await resp.Content.ReadAsByteArrayAsync().ConfigureAwait(False)
                Else
                    Dim err = Await resp.Content.ReadAsStringAsync().ConfigureAwait(False)
                    Debug.WriteLine($"[AB-TTS] OpenAI error {(CInt(resp.StatusCode)).ToString()} {resp.StatusCode}: {err}")
                    Return Nothing
                End If
            End Using
        Else
            Dim token = Await AB_GetFreshToken(AB_googleSecondary)
            If String.IsNullOrEmpty(token) Then Return Nothing

            Dim endpoint = AB_GoogleEndpoint
            If String.IsNullOrWhiteSpace(endpoint) Then
                Debug.WriteLine("[AB-TTS] Google endpoint is empty — cannot generate segment.")
                Return Nothing
            End If

            Using client As New System.Net.Http.HttpClient()
                client.Timeout = TimeSpan.FromSeconds(120)
                client.DefaultRequestHeaders.Authorization =
                New Net.Http.Headers.AuthenticationHeaderValue("Bearer", token)

                Dim textLabel = "text"
                Dim processedText = text
                If Regex.IsMatch(processedText, "<[^>]+>") Then
                    If Not processedText.Trim().StartsWith("<speak>") Then
                        processedText = "<speak>" & processedText & "</speak>"
                    End If
                    textLabel = "ssml"
                End If

                Dim requestBody = New JObject From {
                {"input", New JObject From {{textLabel, processedText}}},
                {"voice", New JObject From {
                    {"languageCode", languageCode},
                    {"name", voiceName}
                }},
                {"audioConfig", New JObject From {
                    {"audioEncoding", "MP3"},
                    {"pitch", 0},
                    {"speakingRate", 1},
                    {"effectsProfileId", New JArray("small-bluetooth-speaker-class-device")}
                }}
            }

                Dim content = New System.Net.Http.StringContent(requestBody.ToString(), Encoding.UTF8, "application/json")
                Debug.WriteLine($"[AB-TTS] POST {endpoint}text:synthesize lang={languageCode} voice={voiceName} len={text.Length}")

                Dim resp = Await client.PostAsync(endpoint & "text:synthesize", content, ct).ConfigureAwait(False)
                If resp.IsSuccessStatusCode Then
                    Dim respStr = Await resp.Content.ReadAsStringAsync().ConfigureAwait(False)
                    Dim respJson = JObject.Parse(respStr)
                    If respJson.ContainsKey("audioContent") Then
                        Return System.Convert.FromBase64String(respJson("audioContent").ToString())
                    Else
                        Debug.WriteLine($"[AB-TTS] Google response missing 'audioContent'. Keys: {String.Join(", ", respJson.Properties().Select(Function(p) p.Name))}")
                    End If
                Else
                    Dim err = Await resp.Content.ReadAsStringAsync().ConfigureAwait(False)
                    Debug.WriteLine($"[AB-TTS] Google error {(CInt(resp.StatusCode)).ToString()} {resp.StatusCode}: {err}")
                End If

                Return Nothing
            End Using
        End If
    End Function



    ' ═══════════════════════════════════════════════════════════════════════════
    '  TTS — AUDIO FILE MERGING (MP3 byte concatenation — no NAudio needed)
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Merges multiple MP3 files into a single output MP3 by direct byte
    ''' concatenation. This works reliably for MP3 files from the same TTS
    ''' provider since they share identical encoding parameters.
    ''' </summary>
    Private Shared Sub AB_MergeAudioFiles(inputFiles As List(Of String), outputFile As String)
        If inputFiles Is Nothing OrElse inputFiles.Count = 0 Then Return

        Using outStream As New FileStream(outputFile, FileMode.Create, FileAccess.Write)
            For Each inPath In inputFiles
                If Not File.Exists(inPath) Then Continue For
                Dim bytes = File.ReadAllBytes(inPath)
                outStream.Write(bytes, 0, bytes.Length)
            Next
        End Using
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TTS — CONVERSATION PARSER
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Shared Function AB_ParseConversation(text As String) As List(Of Tuple(Of String, String))
        Dim conversation As New List(Of Tuple(Of String, String))()
        Dim currentSpeaker As String = ""
        Dim currentText As String = ""

        ' Normalise literal \n escape sequences (from JSON string values) to real newlines
        Dim normalised = text.Replace("\r\n", vbCrLf).Replace("\r", vbCr).Replace("\n", vbLf)

        Dim paragraphs = normalised.Split({vbCrLf, vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)
        For Each para In paragraphs
            Dim trimmed = para.Trim()
            If String.IsNullOrEmpty(trimmed) Then Continue For

            ' Clean Markdown formatting from the start of the line (e.g. "**Host:**" -> "Host:")
            ' preventing first-segment loss if LLM bolds the first speaker tag.
            Dim cleanPara = Regex.Replace(trimmed, "^[\*#_]+", "")

            Dim newSpeaker As String = ""
            If AB_HostTags.Any(Function(tag) cleanPara.StartsWith(tag, StringComparison.OrdinalIgnoreCase)) Then
                newSpeaker = "H"
                trimmed = cleanPara.Substring(cleanPara.IndexOf(":"c) + 1).Trim()
            ElseIf AB_GuestTags.Any(Function(tag) cleanPara.StartsWith(tag, StringComparison.OrdinalIgnoreCase)) Then
                newSpeaker = "G"
                trimmed = cleanPara.Substring(cleanPara.IndexOf(":"c) + 1).Trim()
            End If

            If newSpeaker <> "" Then
                If Not String.IsNullOrEmpty(currentSpeaker) Then
                    conversation.Add(Tuple.Create(currentSpeaker, AB_StripSsml(currentText.Trim())))
                End If
                currentSpeaker = newSpeaker
                currentText = trimmed
            Else
                If Not String.IsNullOrEmpty(currentSpeaker) Then
                    currentText &= " " & trimmed
                End If
            End If
        Next

        If Not String.IsNullOrEmpty(currentSpeaker) Then
            conversation.Add(Tuple.Create(currentSpeaker, AB_StripSsml(currentText.Trim())))
        End If
        Return conversation
    End Function

    ''' <summary>
    ''' Strips SSML / XML tags from text produced by the LLM podcast script.
    ''' The podcast system prompt instructs the LLM to embed SSML tags like
    ''' &lt;emphasis&gt;, &lt;prosody&gt;, &lt;lang&gt;, &lt;break&gt;, &lt;say-as&gt;.
    ''' LLM output is rarely well-formed XML, and malformed SSML causes Google
    ''' TTS to reject the entire segment.  Stripping the tags and sending as
    ''' plain text is safer and still produces correct pronunciation.
    ''' </summary>
    Private Shared Function AB_StripSsml(text As String) As String
        If String.IsNullOrWhiteSpace(text) Then Return text
        ' Remove all XML/SSML tags (opening, closing, self-closing)
        Dim stripped = Regex.Replace(text, "<[^>]+>", "")

        ' Collapse runs of horizontal whitespace (spaces/tabs) but PRESERVE NEWLINES.
        ' The original \s{2,} regex matched newlines, merging "H: ...\nG: ..." into "H: ... G: ...",
        ' breaking the dialogue parser which relies on line breaks.
        stripped = Regex.Replace(stripped, "[ \t]{2,}", " ")

        ' Unescape common XML entities the LLM may have emitted
        stripped = stripped.Replace("&amp;", "&").
                           Replace("&lt;", "<").
                           Replace("&gt;", ">").
                           Replace("&quot;", """").
                           Replace("&apos;", "'")
        Return stripped.Trim()
    End Function

    ''' <summary>
    ''' Strips narrator-style labels ("Host:", "Segment 1:", etc.) that the
    ''' outer tooling LLM sometimes adds when it pre-formats the text before
    ''' passing it to create_audio_file.  These labels confuse the podcast
    ''' dialogue detector (hasHostTag is True but hasGuestTag is False) and
    ''' pollute the source material sent to the inner LLM for script generation.
    ''' </summary>
    Private Shared Function AB_StripNarratorLabels(text As String) As String
        If String.IsNullOrWhiteSpace(text) Then Return text
        ' Remove "Host:" / "Narrator:" / "Segment N:" labels at line starts
        Dim cleaned = Regex.Replace(text, "(?mi)^\s*(Host|Narrator|Segment\s*\d+)\s*:\s*", "")
        Return cleaned.Trim()
    End Function


    ''' <summary>
    ''' Strips the Xing/Info VBR header frame from the beginning of an MP3 byte
    ''' array. TTS providers often prepend this frame, and when multiple segments
    ''' are concatenated the stale header causes players to report wrong duration.
    ''' </summary>
    Private Shared Function AB_StripXingHeader(mp3 As Byte()) As Byte()
        If mp3 Is Nothing OrElse mp3.Length < 4 Then Return mp3

        ' An MP3 frame starts with the sync word &HFF &HE0 (top 11 bits set).
        ' Check if the first frame contains a Xing or Info tag.
        If mp3(0) <> &HFF OrElse (mp3(1) And &HE0) <> &HE0 Then Return mp3

        ' Determine the MPEG version and layer to calculate the side-info length,
        ' which tells us where the Xing/Info tag sits inside the frame.
        Dim mpegV1 = (mp3(1) And &H8) = &H8  ' MPEG1 vs MPEG2/2.5
        Dim stereo = (mp3(3) And &HC0) <> &HC0 ' not mono
        Dim sideInfoLen As Integer
        If mpegV1 Then
            sideInfoLen = If(stereo, 32, 17)
        Else
            sideInfoLen = If(stereo, 17, 9)
        End If

        Dim tagOffset = 4 + sideInfoLen ' past the 4-byte header + side information
        If tagOffset + 4 > mp3.Length Then Return mp3

        ' Look for "Xing" (&H58 &H69 &H6E &H67) or "Info" (&H49 &H6E &H66 &H6F)
        Dim isXing = (mp3(tagOffset) = &H58 AndAlso mp3(tagOffset + 1) = &H69 AndAlso
                      mp3(tagOffset + 2) = &H6E AndAlso mp3(tagOffset + 3) = &H67)
        Dim isInfo = (mp3(tagOffset) = &H49 AndAlso mp3(tagOffset + 1) = &H6E AndAlso
                      mp3(tagOffset + 2) = &H66 AndAlso mp3(tagOffset + 3) = &H6F)

        If Not isXing AndAlso Not isInfo Then Return mp3

        ' Calculate the frame length from the header to skip exactly one frame.
        ' Bitrate index is bits 15-12 of the header, sample rate bits 11-10.
        Dim bitrateIndex = (mp3(2) >> 4) And &HF
        Dim sampleRateIndex = (mp3(2) >> 2) And &H3
        Dim padding = (mp3(2) >> 1) And &H1

        ' MPEG1 Layer 3 bitrate table (kbps). Index 0 and 15 are invalid.
        Dim bitrates = {0, 32, 40, 48, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320, 0}
        Dim sampleRates = If(mpegV1, {44100, 48000, 32000}, {22050, 24000, 16000})

        If bitrateIndex < 1 OrElse bitrateIndex > 14 Then Return mp3
        If sampleRateIndex > 2 Then Return mp3

        Dim bitrate = bitrates(bitrateIndex) * 1000
        Dim sampleRate = sampleRates(sampleRateIndex)
        Dim frameLen = (144 * bitrate \ sampleRate) + padding

        If frameLen <= 0 OrElse frameLen >= mp3.Length Then Return mp3

        ' Return everything after the first (Xing/Info) frame.
        Dim result(mp3.Length - frameLen - 1) As Byte
        Buffer.BlockCopy(mp3, frameLen, result, 0, result.Length)
        Return result
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TTS — PARAGRAPH SPLITTER FOR AUDIOBOOK MODE
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Shared Function AB_SplitIntoParagraphs(text As String) As List(Of Tuple(Of String, String))
        Dim segments As New List(Of Tuple(Of String, String))()

        ' Normalise literal \n escape sequences (from JSON string values) to real newlines
        Dim normalised = text.Replace("\r\n", vbCrLf).Replace("\r", vbCr).Replace("\n", vbLf)

        ' Split on double newlines first (blank-line paragraph breaks).
        ' If that produces only one segment (e.g. text uses single-line breaks
        ' or the LLM script fallback path lands here), re-split on single newlines
        ' so short texts and LLM-generated scripts always produce multiple segments.
        Dim paragraphs As String()
        paragraphs = normalised.Split({vbCrLf & vbCrLf, vbCr & vbCr, vbLf & vbLf}, StringSplitOptions.RemoveEmptyEntries)
        If paragraphs.Length <= 1 Then
            paragraphs = normalised.Split({vbCrLf, vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)
        End If

        Dim speaker = "A"

        For Each para In paragraphs
            Dim trimmed = para.Trim()
            If String.IsNullOrEmpty(trimmed) Then Continue For

            If trimmed.Length <= AB_MaxCharsPerSegment Then
                segments.Add(Tuple.Create(speaker, trimmed))
                speaker = If(speaker = "A", "B", "A")
            Else
                Dim sentences = Regex.Split(trimmed, "(?<=[.!?])\s+")
                Dim chunk As New StringBuilder()
                For Each sentence In sentences
                    If chunk.Length + sentence.Length + 1 > AB_MaxCharsPerSegment AndAlso chunk.Length > 0 Then
                        segments.Add(Tuple.Create(speaker, chunk.ToString().Trim()))
                        speaker = If(speaker = "A", "B", "A")
                        chunk.Clear()
                    End If
                    If chunk.Length > 0 Then chunk.Append(" ")
                    chunk.Append(sentence)
                Next
                If chunk.Length > 0 Then
                    segments.Add(Tuple.Create(speaker, chunk.ToString().Trim()))
                    speaker = If(speaker = "A", "B", "A")
                End If
            End If
        Next
        Return segments
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TTS — MAIN ORCHESTRATOR
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Async Function AB_GenerateAudioFileAsync(
        text As String,
        outputPath As String,
        mode As String,
        language As String,
        voiceA As String,
        voiceB As String,
        podcastDuration As String,
        podcastContext As String,
        podcastInstructions As String,
        podcastHostName As String,
        podcastGuestName As String,
        podcastHostGender As String,
        podcastGuestGender As String,
        singleVoice As Boolean,
        narratorGender As String,
        ct As CancellationToken) As Task(Of Boolean)

        If String.IsNullOrWhiteSpace(language) Then language = AB_DefaultLanguage
        language = AB_NormalizeLanguageForStandardAudio(language)

        If String.IsNullOrWhiteSpace(mode) Then mode = "audiobook"
        If String.IsNullOrWhiteSpace(podcastDuration) Then podcastDuration = "5 minutes"

        Dim effectiveHostName As String = podcastHostName
        Dim effectiveGuestName As String = podcastGuestName
        Dim effectiveHostGender As String = podcastHostGender
        Dim effectiveGuestGender As String = podcastGuestGender

        AB_ResolveAudioSpeakerAndVoiceInputs(
                        mode,
                        language,
                        effectiveHostName,
                        effectiveGuestName,
                        effectiveHostGender,
                        effectiveGuestGender,
                        voiceA,
                        voiceB,
                        singleVoice,
                        narratorGender)

        Debug.WriteLine($"[AB-TTS] Orchestrator: engine={AB_SelectedEngine} lang={language} voiceA={voiceA} voiceB={voiceB} mode={mode} singleVoice={singleVoice} narratorGender={narratorGender} host={effectiveHostName}/{effectiveHostGender} guest={effectiveGuestName}/{effectiveGuestGender}")
        Dim segments As List(Of Tuple(Of String, String))

        ' Normalise literal \n escape sequences that arrive from JSON-encoded
        ' tool parameters so that tag detection and parsing work consistently.
        text = text.Replace("\r\n", vbCrLf).Replace("\r", vbCr).Replace("\n", vbLf)

        If mode.Equals("podcast", StringComparison.OrdinalIgnoreCase) Then
            ' Check whether the text already contains a dialogue script with
            ' speaker tags at the START of lines.  We require BOTH a host tag
            ' and a guest tag to treat the text as a pre-formed dialogue.
            Dim lines = text.Split({vbCrLf, vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)
            Dim hasHostTag As Boolean = lines.Any(Function(l)
                                                      Dim tr = l.TrimStart()
                                                      Return AB_HostTags.Any(Function(tag) tr.StartsWith(tag, StringComparison.OrdinalIgnoreCase))
                                                  End Function)
            Dim hasGuestTag As Boolean = lines.Any(Function(l)
                                                       Dim tr = l.TrimStart()
                                                       Return AB_GuestTags.Any(Function(tag) tr.StartsWith(tag, StringComparison.OrdinalIgnoreCase))
                                                   End Function)

            If hasHostTag AndAlso hasGuestTag Then
                ' The text already has both H: and G: tags — parse it directly.
                segments = AB_ParseConversation(text)
                ApDashboardLog("🎙 Parsed existing dialogue script for podcast audio.", "step")
            Else
                ' The text does NOT have a proper two-speaker dialogue.
                ' Generate one via the LLM using the podcast system prompt.
                ' IMPORTANT: Use a separate variable for the truncated language code
                ' so the full BCP-47 code (e.g. "de-DE") remains intact for TTS.
                ApDashboardLog("🎙 Generating podcast dialogue script via LLM...", "step")

                HostName = effectiveHostName
                GuestName = effectiveGuestName
                TargetAudience = "General audience"
                Duration = podcastDuration
                DialogueContext = If(String.IsNullOrWhiteSpace(podcastContext), "None", podcastContext)
                ExtraInstructions = If(String.IsNullOrWhiteSpace(podcastInstructions), "None", podcastInstructions)

                Dim llmLanguage = If(language.Length > 2, language.Substring(0, 2), language)

                ' NOTE: Do NOT overwrite the main 'language' variable here. The full BCP-47 code
                ' (e.g. "en-US") must be preserved for the TTS engine call later.
                ' language = llmLanguage 

                ' Strip any pre-existing Host:/Segment: labels from the text so
                ' the LLM receives clean source material, not a half-formatted
                ' narration that the outer LLM may have produced.
                Dim cleanText = AB_StripNarratorLabels(text)

                Debug.WriteLine($"[AB-TTS] Generating podcast currently. Input text len: {cleanText.Length}. Language: {language}")

                Dim systemPrompt = InterpolateAtRuntime(SP_Podcast)
                ' Explicitly instruct the LLM to use the generic Standard language (e.g. "de")
                ' to prevent it from writing in dialects (like Swiss German) when the locale is specific (e.g. "de-CH").
                Dim langInstruction = AB_BuildStandardLanguageInstruction(language)
                Dim userPrompt = langInstruction & vbCrLf & vbCrLf & "<TEXTTOPROCESS>" & cleanText & "</TEXTTOPROCESS>"

                Dim script = Await LLM(systemPrompt, userPrompt,
                                       UseSecondAPI:=_apUseSecondApi,
                                       HideSplash:=True, EnsureUI:=False,
                                       cancellationToken:=ct)

                If String.IsNullOrWhiteSpace(script) Then
                    ApDashboardLog("⚠ LLM returned empty podcast script.", "warn")
                    Return False
                End If

                Debug.WriteLine($"[AB-TTS] LLM returned script (len={script.Length}). Preview: {script.Substring(0, Math.Min(script.Length, 200)).Replace(vbCr, "").Replace(vbLf, " ")}...")

                ct.ThrowIfCancellationRequested()

                ' Strip SSML tags from the LLM-generated script. The podcast
                ' system prompt instructs the LLM to include SSML (<emphasis>,
                ' <prosody>, <lang>, <break>, <say-as>), but LLM output is
                ' rarely well-formed XML.  Malformed SSML causes Google TTS to
                ' reject the entire segment.  Plain text is safer and still
                ' produces correct pronunciation for the target language.
                script = AB_StripSsml(script)

                segments = AB_ParseConversation(script)
                ApDashboardLog($"🎙 Podcast script generated: {segments.Count} dialogue segments.", "step")
                Debug.WriteLine($"[AB-TTS] Parsed {segments.Count} segments from script.")

                If segments.Count < 2 Then
                    ApDashboardLog("⚠ Script had no H:/G: tags — falling back to audiobook mode.", "warn")
                    segments = AB_SplitIntoParagraphs(script)
                End If
            End If
        Else
            segments = AB_SplitIntoParagraphs(text)
            ApDashboardLog($"📖 Audiobook mode: {segments.Count} text segments.", "step")
        End If

        If segments.Count = 0 Then
            ApDashboardLog("⚠ No segments to process.", "warn")
            Return False
        End If

        ' Pre-generate silence files at each duration tier
        Dim tempDir = Path.GetDirectoryName(outputPath)
        Dim silenceFileParagraph = AB_GenerateSilenceFile(AB_SilenceBetweenParagraphs, tempDir)
        Dim silenceFileSpeakerChange = AB_GenerateSilenceFile(AB_SilenceOnSpeakerChange, tempDir)
        Dim silenceFilePodcastTurn = AB_GenerateSilenceFile(AB_SilenceBetweenPodcastTurns, tempDir)

        Dim tempFiles As New List(Of String)()
        Dim silenceFiles As New List(Of String)()
        If Not String.IsNullOrEmpty(silenceFileParagraph) Then silenceFiles.Add(silenceFileParagraph)
        If Not String.IsNullOrEmpty(silenceFileSpeakerChange) AndAlso Not silenceFiles.Contains(silenceFileSpeakerChange) Then silenceFiles.Add(silenceFileSpeakerChange)
        If Not String.IsNullOrEmpty(silenceFilePodcastTurn) AndAlso Not silenceFiles.Contains(silenceFilePodcastTurn) Then silenceFiles.Add(silenceFilePodcastTurn)

        Try
            Dim previousSpeaker As String = ""

            For i = 0 To segments.Count - 1
                ct.ThrowIfCancellationRequested()

                Dim speaker = segments(i).Item1
                Dim segText = segments(i).Item2

                Dim voice As String
                If mode.Equals("podcast", StringComparison.OrdinalIgnoreCase) Then
                    voice = If(speaker = "H", voiceA, voiceB)
                Else
                    voice = If(speaker = "A", voiceA, voiceB)
                End If

                Debug.WriteLine($"[AB-TTS] Processing Segment {i + 1}: Speaker={speaker}, Voice={voice}, Lang={language}, Len={segText.Length}")

                If String.IsNullOrWhiteSpace(segText) Then
                    Debug.WriteLine($"[AB-TTS] Segment {i + 1} text is empty, skipping.")
                    Continue For
                End If

                Dim audioBytes As Byte() = Nothing
                For attempt = 1 To AB_MaxSegmentRetries
                    ct.ThrowIfCancellationRequested()
                    audioBytes = Await AB_GenerateSegmentAsync(segText, language, voice, ct)
                    If audioBytes IsNot Nothing AndAlso audioBytes.Length > 0 Then Exit For
                    Debug.WriteLine($"[AB-TTS] Segment {i + 1} attempt {attempt}/{AB_MaxSegmentRetries} returned no audio.")
                    If attempt < AB_MaxSegmentRetries Then
                        ApDashboardLog($"⚠ Segment {i + 1} attempt {attempt} failed, retrying...", "warn")
                        Await Task.Delay(1000 * attempt, ct) ' exponential back-off: 1s, 2s
                    End If
                Next

                If audioBytes Is Nothing OrElse audioBytes.Length = 0 Then
                    ApDashboardLog($"⚠ Segment {i + 1}/{segments.Count} returned no audio after {AB_MaxSegmentRetries} attempts, skipping.", "warn")
                    Continue For
                End If

                ' Strip any Xing/Info VBR header from the TTS response so that
                ' byte-concatenated MP3s don't confuse players into reporting
                ' an incorrect total duration.
                audioBytes = AB_StripXingHeader(audioBytes)

                Dim tempFile = Path.Combine(tempDir, $"ab_seg_{i:D4}.mp3")
                File.WriteAllBytes(tempFile, audioBytes)
                tempFiles.Add(tempFile)

                ' Select the appropriate silence gap based on context:
                '  - Podcast mode: always use the podcast turn pause between segments
                '  - Audiobook mode: use the longer speaker-change pause when the
                '    narrator voice switches, otherwise the regular paragraph pause
                Dim pauseFile As String = Nothing
                If mode.Equals("podcast", StringComparison.OrdinalIgnoreCase) Then
                    pauseFile = silenceFilePodcastTurn
                Else
                    Dim speakerChanged = (previousSpeaker <> "" AndAlso speaker <> previousSpeaker)
                    pauseFile = If(speakerChanged, silenceFileSpeakerChange, silenceFileParagraph)
                End If
                If Not String.IsNullOrEmpty(pauseFile) AndAlso File.Exists(pauseFile) Then
                    tempFiles.Add(pauseFile)
                End If

                previousSpeaker = speaker
                ApDashboardLog($"🔊 Segment {i + 1}/{segments.Count} generated ({audioBytes.Length / 1024:F0} KB).", "step")
                Await Task.Delay(500, ct)
            Next

            If tempFiles.Count = 0 Then
                ApDashboardLog("⚠ No audio segments were generated.", "warn")
                Return False
            End If

            ct.ThrowIfCancellationRequested()
            AB_MergeAudioFiles(tempFiles, outputPath)
            ApDashboardLog($"✅ Audio file created: {Path.GetFileName(outputPath)} ({segments.Count} segments merged).", "step")
            Return True

        Finally
            For Each f In tempFiles
                ' Don't delete the shared silence files yet — they may appear
                ' multiple times in the list.  Clean them up separately.
                If silenceFiles.Contains(f) Then Continue For
                Try : File.Delete(f) : Catch : End Try
            Next
            For Each f In silenceFiles
                Try : File.Delete(f) : Catch : End Try
            Next
        End Try
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TTS — TOOL EXECUTOR
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Async Function ExecuteCreateAudioFileTool(
            toolCall As ToolCall,
            context As ToolExecutionContext,
            ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId,
            .ToolName = toolCall.ToolName
        }

        Try
            Dim text = GetArgString(toolCall.Arguments, "text")
            If String.IsNullOrWhiteSpace(text) Then
                Dim errMsg = "Missing required parameter: text"
                response.Success = False
                response.ErrorMessage = errMsg
                response.Response = errMsg
                Return response
            End If

            Dim mode = If(GetArgString(toolCall.Arguments, "mode"), "audiobook")
            Dim language = If(GetArgString(toolCall.Arguments, "language"), AB_DefaultLanguage)
            Dim voiceA = GetArgString(toolCall.Arguments, "voice_a")
            Dim voiceB = GetArgString(toolCall.Arguments, "voice_b")
            Dim hostName = GetArgString(toolCall.Arguments, "host_name")
            Dim guestName = GetArgString(toolCall.Arguments, "guest_name")
            Dim hostGender = GetArgString(toolCall.Arguments, "host_gender")
            Dim guestGender = GetArgString(toolCall.Arguments, "guest_gender")
            Dim singleVoice = GetArgBool(toolCall.Arguments, "single_voice", False)
            Dim narratorGender = GetArgString(toolCall.Arguments, "narrator_gender")
            Dim duration = If(GetArgString(toolCall.Arguments, "duration"), "5 minutes")
            Dim instructions = GetArgString(toolCall.Arguments, "instructions")
            Dim topicContext = GetArgString(toolCall.Arguments, "context")
            Dim outputFilename = If(GetArgString(toolCall.Arguments, "output_filename"), AB_DefaultFile)
            Dim enginePref = GetArgString(toolCall.Arguments, "engine")

            If mode.Equals("audiobook", StringComparison.OrdinalIgnoreCase) AndAlso
   (toolCall.Arguments Is Nothing OrElse Not toolCall.Arguments.ContainsKey("single_voice")) Then
                singleVoice = True
            End If

            For Each c In Path.GetInvalidFileNameChars()
                outputFilename = outputFilename.Replace(c, "_"c)
            Next
            If Not outputFilename.EndsWith(".mp3", StringComparison.OrdinalIgnoreCase) Then
                outputFilename = Path.GetFileNameWithoutExtension(outputFilename) & ".mp3"
            End If

            Dim outputPath = Path.Combine(_apCurrentTempDir, outputFilename)
            Dim counter = 1
            While File.Exists(outputPath)
                Dim baseName = Path.GetFileNameWithoutExtension(outputFilename)
                outputPath = Path.Combine(_apCurrentTempDir, baseName & $"_{counter}.mp3")
                counter += 1
            End While

            ' Select the best engine for this request (considers language + preference).
            AB_SelectBestEngine(enginePref, language)

            Dim textLen = text.Length
            context.Log($"Generating {mode} audio: {textLen} chars, language={language}, engine={AB_SelectedEngine}, single_voice={singleVoice}, narrator_gender={If(narratorGender, "default")}, speakers={If(hostName, "default")}/{If(guestName, "default")}, genders={If(hostGender, "default")}/{If(guestGender, "default")}, voices={If(voiceA, "auto")}/{If(voiceB, "auto")}")
            ApDashboardLog($"🔊 Creating {mode} audio ({textLen:N0} chars, {language}, {AB_SelectedEngine})...", "step")

            ' Do NOT call AB_DetectTTSEngines() here. The detection already ran
            ' at tool registration time (in GetAutoPilotInternalTools) and set
            ' the shared static fields. Re-detecting here would read INI_Endpoint
            ' and INI_Endpoint_2 which may have been overwritten by the tooling
            ' loop's ApplyModelConfig, causing the TTS engine to appear unavailable.

            Debug.WriteLine($"[AB-TTS] Engine state: google={AB_googleAvailable} openai={AB_openAIAvailable} " &
                            $"selected={AB_SelectedEngine} googleEP='{AB_GoogleEndpoint}' openaiEP='{AB_OpenAIEndpoint}'")

            If Not AB_googleAvailable AndAlso Not AB_openAIAvailable Then
                Dim errMsg = "No TTS engine is configured. Cannot generate audio. " &
                    "Ensure INI_Endpoint contains an OpenAI or Google endpoint, " &
                    "and INI_TTSEndpoint contains the TTS-specific URL."
                response.Success = False
                response.ErrorMessage = errMsg
                response.Response = errMsg
                Return response
            End If

            Dim success = Await AB_GenerateAudioFileAsync(
              text, outputPath, mode, language,
              voiceA, voiceB, duration, topicContext, instructions,
              hostName, guestName, hostGender, guestGender,
              singleVoice, narratorGender, ct)

            If success AndAlso File.Exists(outputPath) Then

                Dim fileSize = New FileInfo(outputPath).Length
                response.Success = True
                response.Response = $"Audio file '{Path.GetFileName(outputPath)}' generated successfully " &
                    $"({fileSize / 1024:F0} KB, {mode} mode, engine={AB_SelectedEngine}). The file is attached to the reply."

                ' Register explicitly even when the current mail/session has no inbound attachment.
                ' The generated carrier is marked IsToolOutput, so it is never mistaken for
                ' an original inbound file by the bounded compatibility collector.
                RegisterAutoPilotGeneratedOutputFile(outputPath)
            Else
                Dim errMsg = "Audio generation failed. No audio segments could be produced. " &
                    $"Engine={AB_SelectedEngine}, Endpoint='{If(AB_SelectedEngine = ABEngine.OpenAI, AB_OpenAIEndpoint, AB_GoogleEndpoint)}'. " &
                    "Check the Output window for [AB-TTS] diagnostic messages."
                response.Success = False
                response.ErrorMessage = errMsg
                response.Response = errMsg
            End If

        Catch ex As OperationCanceledException
            Dim errMsg = "Audio generation was cancelled."
            response.Success = False
            response.ErrorMessage = errMsg
            response.Response = errMsg
        Catch ex As Exception
            Dim errMsg = If(String.IsNullOrWhiteSpace(ex.Message), ex.GetType().Name, ex.Message)
            Debug.WriteLine($"[AB-TTS] Tool error: {ex}")
            response.Success = False
            response.ErrorMessage = errMsg
            response.Response = $"Error generating audio: {errMsg}"
        End Try

        Return response
    End Function

End Class