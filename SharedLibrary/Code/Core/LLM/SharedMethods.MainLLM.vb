' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: SharedMethods.MainLLM.vb
' Purpose: Contains the main shared LLM invocation pipeline, including optional post-correction,
'          anonymization/re-identification, request construction (GET/POST), optional multipart
'          object upload, retry/timeout/cancellation handling, response parsing, and token usage logging.
'
' Architecture:
'  - PostCorrection: Optional second-stage call to `LLM` driven by `context.INI_PostCorrection`.
'  - LLM (core pipeline):
'      - Cancellation: Returns empty string early if `cancellationToken` is already canceled; links a local
'        token source to support both caller cancellation and splash-screen cancellation.
'      - Anonymization: Applies `AnonymizeText` / `ReidentifyText` based on model-specific settings and
'        preserves optional `<TEXTTOPROCESS>` wrapping.
'      - OAuth2 API key refresh (optional): Uses `GetFreshAccessToken` to refresh `context.DecodedAPI`
'        or `context.DecodedAPI_2` and sets token expiry timestamps.
'      - Endpoint/APICall templating: Replaces placeholders (model, prompts, temperature, API key, session id).
'      - Dual-call mode (optional): Splits endpoint/body/response key on the Unicode "¦" separator to run
'        a POST followed by a GET whose URL/body is filled from selected POST response tokens.
'      - HTTP: Supports GET via `get:` endpoint prefix; otherwise POST. Retries 429 responses with backoff.
'      - Response normalization: Detects simple SSE framing and extracts "data:" payloads, then validates JSON.
'      - Response extraction: Uses `HandleObject` and `JsonTemplateFormatter.FormatJsonWithTemplate` and may
'        append citations via `ExtractCitations`.
'      - Output cleanup: Optional replacement of ß, dash normalization, whitespace cleanup, hidden marker removal,
'        and optional removal of content up to the last `</THINK>` tag.
'  - Prompt log and token spending log: `LogTokenSpending` persists recent prompts in settings and optionally
'    appends a cost/token line to a desktop log file with retry/exclusive-lock writing.
'  - JSON helpers: `HandleObject` and `FindJsonProperty` read selected values and detect error payloads.
'  - OAuth2 helper: `GetFreshAccessToken` prepares a PEM key for `GoogleOAuthHelper` and refreshes access tokens.
'  - Utility helpers: `CleanString` JSON-escapes prompt strings, `EstimateTokenCount` approximates token usage,
'    `FixMimeType` normalizes MIME aliases used for object uploads.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Globalization
Imports System.IO
Imports System.Security.Cryptography.X509Certificates
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Windows.Forms
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary

    Partial Public Class SharedMethods

        ''' <summary>
        ''' Optionally performs a post-processing LLM call as configured by <paramref name="context"/> and returns the resulting text.
        ''' </summary>
        ''' <param name="context">Shared configuration context used to read post-correction settings and model/API selection.</param>
        ''' <param name="inputText">Input text to be passed through the post-correction prompt.</param>
        ''' <param name="UseSecondAPI">If <c>True</c>, uses the secondary API/model configuration from <paramref name="context"/>.</param>
        ''' <returns>The original <paramref name="inputText"/> if post-correction is not configured; otherwise the post-corrected text.</returns>
        Public Shared Async Function PostCorrection(context As ISharedContext, inputText As String, Optional ByVal UseSecondAPI As Boolean = False) As Task(Of String)
            Dim OutputText As String = inputText
            If Not String.IsNullOrEmpty(context.INI_PostCorrection) Then

                ' Wait not to overload the API
                Await System.Threading.Tasks.Task.Delay(500)

                OutputText = Await LLM(context, context.INI_PostCorrection, "<TEXTTOPROCESS>" & inputText & "</TEXTTOPROCESS>", "", "", 0, UseSecondAPI)
            End If
            Return OutputText
        End Function

        ''' <summary>
        ''' Calls a configured LLM endpoint and returns extracted response text based on the configured response template/key.
        ''' </summary>
        ''' <param name="context">Shared configuration context used for endpoint templates, headers, model selection, and feature flags.</param>
        ''' <param name="promptSystem">System prompt content used to build the request (may be templated into endpoint/body).</param>
        ''' <param name="promptUser">User prompt content used to build the request (may be templated into endpoint/body).</param>
        ''' <param name="Model">Optional model override; "Default" or empty defers to configured model.</param>
        ''' <param name="Temperature">Optional temperature override; "Default" or empty defers to configured temperature.</param>
        ''' <param name="Timeout">Optional timeout override in milliseconds; 0 defers to configured timeout.</param>
        ''' <param name="UseSecondAPI">If <c>True</c>, uses the secondary API/model configuration from <paramref name="context"/>.</param>
        ''' <param name="Hidesplash">If <c>True</c>, suppresses the countdown splash UI.</param>
        ''' <param name="AddUserPrompt">Optional additional user instruction used for placeholder replacement and logging.</param>
        ''' <param name="FileObject">Optional file/clipboard object reference used for object upload features.</param>
        ''' <param name="cancellationToken">Cancellation token propagated to network calls and linked to splash cancellation.</param>
        ''' <param name="ToolExecution">If <c>True</c> then LLM expects to be in the tooling execution mode when calling an LLM (necessary for building APICall).</c>.</param>
        ''' <returns>Extracted text from the JSON response; returns an empty string on cancellation or on handled errors.</returns>
        Public Shared Async Function LLM(context As ISharedContext, ByVal promptSystem As String, ByVal promptUser As String, Optional ByVal Model As String = "", Optional ByVal Temperature As String = "", Optional ByVal Timeout As Long = 0, Optional ByVal UseSecondAPI As Boolean = False, Optional ByVal Hidesplash As Boolean = False, Optional ByVal AddUserPrompt As String = "", Optional FileObject As String = "", Optional cancellationToken As Threading.CancellationToken = Nothing, Optional ToolExecution As Boolean = False) As Task(Of String)

            If cancellationToken.IsCancellationRequested Then
                Return ""
            End If

            ' Anonymization features

            Dim ModelName As String = If(UseSecondAPI, context.INI_Model_2, context.INI_Model)
            Dim AnonSetting As String = If(UseSecondAPI, context.INI_Anon_2, context.INI_Anon)
            Dim OverrideAnonSetting As String = LoadAnonSettingsForModel(ModelName)
            Dim AnonActive As Boolean = False
            If Not String.IsNullOrWhiteSpace(OverrideAnonSetting) Then AnonSetting = OverrideAnonSetting
            If Not String.IsNullOrWhiteSpace(AnonSetting) Then
                Dim AnonType As Integer = GetTypeFromSettings(AnonSetting)
                If AnonType > 0 And Not String.IsNullOrWhiteSpace(promptUser) Then
                    Dim AnonMode As String = GetModeFromSettings(AnonSetting)

                    Dim TTPPrefix As Boolean = False
                    Dim TTPSuffix As Boolean = False
                    If promptUser.TrimStart().StartsWith("<TEXTTOPROCESS>", StringComparison.OrdinalIgnoreCase) Then
                        TTPPrefix = True
                        promptUser = promptUser.TrimStart()
                        promptUser = promptUser.Substring("<TEXTTOPROCESS>".Length)
                    End If
                    If promptUser.TrimEnd().EndsWith("</TEXTTOPROCESS>", StringComparison.OrdinalIgnoreCase) Then
                        TTPSuffix = True
                        promptUser = promptUser.TrimEnd()
                        promptUser = promptUser.Substring(0, promptUser.Length - "</TEXTTOPROCESS>".Length)
                    End If

                    promptUser = AnonymizeText(promptUser, ModelName, AnonMode, AnonType)

                    If String.IsNullOrWhiteSpace(promptUser) Then Return ""

                    If TTPPrefix Then promptUser = "<TEXTTOPROCESS>" & promptUser
                    If TTPSuffix Then promptUser = promptUser & "</TEXTTOPROCESS>"

                    AnonActive = True
                End If
            End If

            If cancellationToken.IsCancellationRequested Then
                Return ""
            End If

            Dim splash As SplashScreenCountDown = Nothing
            Dim cts As System.Threading.CancellationTokenSource = Nothing

            Dim TokenCountString As String = ""

            Try

                ' Configure TLS
                If (System.Net.ServicePointManager.SecurityProtocol And System.Net.SecurityProtocolType.Tls12) = 0 Then
                    System.Net.ServicePointManager.SecurityProtocol = System.Net.ServicePointManager.SecurityProtocol Or System.Net.SecurityProtocolType.Tls12
                End If
                System.Net.ServicePointManager.DefaultConnectionLimit = 100 ' Adjust based on expected load

                ' Initialize API variables
                Dim Endpoint As String
                Dim HeaderA As String
                Dim HeaderB As String
                Dim APICall As String
                Dim TemperatureValue As String
                Dim ModelValue As String
                Dim TimeoutValue As Long
                Dim ResponseKey As String
                Dim DoubleS As Boolean
                Dim NoThink As Boolean

                ''' <summary>
                ''' Per-request unique session identifier used for endpoint/body placeholder replacement.
                ''' </summary>
                Dim OwnSessionID As String = GenerateUniqueId()

                ' === Support for two calls in a single run ===
                Dim sep As String = "¦" ' Unicode broken bar; rarely occurs in URLs/JSON.
                Dim sep2 As String = ";" ' Separator for multiple parameters in the POST response key list.
                Dim postEndpoint As String
                Dim getEndpointTemplate As String = ""
                Dim postAPICall As String
                Dim getAPICallTemplate As String = ""
                Dim postResponseKey As String
                Dim getResponseKey As String = ""

                Dim multiCall As Boolean = False

                If UseSecondAPI Then

                    If context.INI_OAuth2_2 Then
                        context.DecodedAPI_2 = Await GetFreshAccessToken(context, context.INI_OAuth2ClientMail_2, context.INI_OAuth2Scopes_2, context.INI_APIKey_2, context.INI_OAuth2Endpoint_2, context.INI_OAuth2ATExpiry_2, True)
                        If context.DecodedAPI_2 = "" Then Exit Function
                    End If

                Else
                    If context.INI_OAuth2 Then
                        context.DecodedAPI = Await GetFreshAccessToken(context, context.INI_OAuth2ClientMail, context.INI_OAuth2Scopes, context.INI_APIKey, context.INI_OAuth2Endpoint, context.INI_OAuth2ATExpiry, False)
                        If context.DecodedAPI = "" Then Exit Function
                    End If
                End If

                If UseSecondAPI Then

                    Dim ModelPlaceholder As String = context.INI_Model_2

                    If Not SharedMethods.ProcessParameterPlaceholders(ModelPlaceholder) Then
                        ShowCustomMessageBox("Aborted by user.")
                        Return ""
                    End If

                    Endpoint = Replace(Replace(Replace(context.INI_Endpoint_2, "{model}", ModelPlaceholder), "{apikey}", context.DecodedAPI_2), "{ownsessionid}", OwnSessionID)
                    HeaderA = Replace(Replace(context.INI_HeaderA_2, "{model}", ModelPlaceholder), "{apikey}", context.DecodedAPI_2)
                    HeaderB = Replace(Replace(context.INI_HeaderB_2, "{model}", ModelPlaceholder), "{apikey}", context.DecodedAPI_2)
                    APICall = context.INI_APICall_2
                    ResponseKey = context.INI_Response_2
                    DoubleS = context.INI_DoubleS

                    TemperatureValue = If(String.IsNullOrEmpty(Temperature) OrElse Temperature = "Default", context.INI_Temperature_2, Temperature)
                    ModelValue = If(String.IsNullOrEmpty(Model) OrElse Model = "Default", ModelPlaceholder, Model)
                    TimeoutValue = If(Timeout = 0, context.INI_Timeout_2, Timeout)
                    TokenCountString = context.INI_TokenCount_2
                Else

                    Dim ModelPlaceholder As String = context.INI_Model

                    If Not SharedMethods.ProcessParameterPlaceholders(ModelPlaceholder) Then
                        ShowCustomMessageBox("Aborted by user.")
                        Return ""
                    End If

                    Endpoint = Replace(Replace(Replace(context.INI_Endpoint, "{model}", ModelPlaceholder), "{apikey}", context.DecodedAPI), "{ownsessionid}", OwnSessionID)
                    HeaderA = Replace(Replace(context.INI_HeaderA, "{model}", ModelPlaceholder), "{apikey}", context.DecodedAPI)
                    HeaderB = Replace(Replace(context.INI_HeaderB, "{model}", ModelPlaceholder), "{apikey}", context.DecodedAPI)
                    APICall = context.INI_APICall
                    ResponseKey = context.INI_Response
                    DoubleS = context.INI_DoubleS
                    TemperatureValue = If(String.IsNullOrEmpty(Temperature) OrElse Temperature = "Default", context.INI_Temperature, Temperature)
                    ModelValue = If(String.IsNullOrEmpty(Model) OrElse Model = "Default", ModelPlaceholder, Model)
                    TimeoutValue = If(Timeout = 0, context.INI_Timeout, Timeout)
                    TokenCountString = context.INI_TokenCount

                    ' Tooling not supported for primary model by definition
                    ToolExecution = False

                End If

                If TimeoutValue = 0 Then TimeoutValue = 30000

                Dim timeoutSeconds = CInt(TimeoutValue \ 1000)

                If Not SharedMethods.ProcessParameterPlaceholders(APICall) Then
                    ShowCustomMessageBox("Aborted by user.")
                    Return ""
                End If

                NoThink = False
                If Not String.IsNullOrEmpty(ResponseKey) Then
                    Dim trigger As String = If(TryCast(NoThinkTrigger, String), String.Empty)
                    If Not String.IsNullOrEmpty(trigger) Then

                        Dim idx As Integer = ResponseKey.LastIndexOf(trigger, StringComparison.OrdinalIgnoreCase)
                        If idx >= 0 Then
                            NoThink = True
                            ' Remove ALL occurrences (case-insensitive) and trim.
                            ResponseKey = Regex.Replace(ResponseKey,
                                                        Regex.Escape(trigger),
                                                        String.Empty,
                                                        RegexOptions.IgnoreCase).Trim()
                        End If

                    End If
                End If

                ' --- ToolCallMatching trigger inside ResponseKey ---
                ' Expected form: "(toolcall:<pattern>)" but we remove it even if malformed (e.g. missing ")", missing "<>").
                ' Define the markers: ToolCallMatchingStart = "(toolcall:", ToolCallMatchingEnd = ")", ToolCallMatchingMiddle = ":"

                Dim DetectToolCall As String = String.Empty

                If Not String.IsNullOrEmpty(ResponseKey) Then

                    Dim startMarker As String = ToolCallMatchingStart
                    Dim startIdx As Integer = ResponseKey.IndexOf(startMarker, StringComparison.OrdinalIgnoreCase)

                    If startIdx >= 0 Then
                        ' Find the closing ")" after the marker; if missing, treat as malformed and consume to end.
                        Dim endIdx As Integer = ResponseKey.IndexOf(ToolCallMatchingEnd, startIdx, StringComparison.OrdinalIgnoreCase)
                        Dim triggerLen As Integer = If(endIdx >= 0,
                               (endIdx - startIdx + ToolCallMatchingEnd.Length),
                               (ResponseKey.Length - startIdx))

                        Dim triggerText As String = ResponseKey.Substring(startIdx, triggerLen)

                        ' 1) Extract pattern (prefer <...> if present)
                        Dim lt As Integer = triggerText.IndexOf("<"c)
                        Dim gt As Integer = triggerText.LastIndexOf(">"c)

                        If lt >= 0 AndAlso gt > lt Then
                            DetectToolCall = triggerText.Substring(lt + 1, gt - lt - 1).Trim()
                        Else
                            ' 2) Fallback: pattern starts after ":" and ends before ")" (or end of triggerText)
                            Dim colonIdx As Integer = triggerText.IndexOf(ToolCallMatchingMiddle, StringComparison.OrdinalIgnoreCase)
                            If colonIdx >= 0 Then
                                Dim raw As String = triggerText.Substring(colonIdx + ToolCallMatchingMiddle.Length)
                                Dim paren As Integer = raw.LastIndexOf(ToolCallMatchingEnd, StringComparison.OrdinalIgnoreCase)
                                If paren >= 0 Then raw = raw.Substring(0, paren)
                                DetectToolCall = raw.Trim()
                            End If
                        End If

                        ' Validate .NET regex; if invalid, wipe it (but still remove trigger from ResponseKey)
                        If Not String.IsNullOrWhiteSpace(DetectToolCall) Then
                            Try
                                Dim rx As New Regex(DetectToolCall)
                            Catch ex As ArgumentException
                                DetectToolCall = String.Empty
                            End Try
                        End If

                        ' Remove the trigger block from ResponseKey (even if malformed)
                        ResponseKey = (ResponseKey.Substring(0, startIdx) &
                                       ResponseKey.Substring(startIdx + triggerLen)).Trim()
                    End If

                End If

                ' Determine RKMode (default = 2) based on trigger markers optionally embedded in ResponseKey.
                ' (rkmode_all)      -> 1
                ' (rkmode_longest)  -> 2
                ' (rkmode_first)    -> 3

                ' If multiple markers somehow appear, the first one found in the list below (priority order) wins.
                Dim RKMode As Integer = 2 ' default
                If Not String.IsNullOrEmpty(ResponseKey) Then
                    Dim modeTriggers = {
                        New With {.Trigger = RKModeTrigger1, .Mode = 1},
                        New With {.Trigger = RKModeTrigger2, .Mode = 2},
                        New With {.Trigger = RKModeTrigger3, .Mode = 3}
                    }

                    For Each mt In modeTriggers
                        Dim idx As Integer = ResponseKey.LastIndexOf(mt.Trigger, StringComparison.OrdinalIgnoreCase)
                        If idx >= 0 Then
                            RKMode = mt.Mode
                            ' Remove the trigger occurrence
                            ResponseKey = (ResponseKey.Substring(0, idx) &
                                           ResponseKey.Substring(idx + mt.Trigger.Length)).Trim()
                            Exit For
                        End If
                    Next
                End If

                ' Create splash & CTS once:
                splash = New SplashScreenCountDown("Waiting for the AI to respond...", 0, 0, timeoutSeconds)

                'cts = New System.Threading.CancellationTokenSource()
                'AddHandler splash.CancelRequested, Sub() cts.Cancel()
                'Dim ct As System.Threading.CancellationToken = cts.Token

                ' Link a local CTS with the external token so both caller cancellation and the splash cancel button apply.
                cts = System.Threading.CancellationTokenSource.CreateLinkedTokenSource(cancellationToken)
                AddHandler splash.CancelRequested, Sub() cts.Cancel()
                Dim ct As System.Threading.CancellationToken = cts.Token

                If Not Hidesplash Then
                    splash.Show()
                    splash.RestartCountdown(timeoutSeconds)
                End If

                Endpoint = Endpoint.Replace("{promptsystem}", CleanString(Left(promptSystem, 32000)))
                Endpoint = Endpoint.Replace("{promptuser}", CleanString(Left(promptUser, 32000).Replace("<TEXTTOPROCESS>", "").Replace("</TEXTTOPROCESS>", "").Trim()))
                Endpoint = Endpoint.Replace("{userinstruction}", CleanString(AddUserPrompt))
                Endpoint = Endpoint.Trim().Replace(" ", "+")

                Dim epParts() As String = Endpoint.Split(New String() {sep}, StringSplitOptions.None)
                Dim apiParts() As String = APICall.Split(New String() {sep}, StringSplitOptions.None)
                Dim respParts() As String = ResponseKey.Split(New String() {sep}, StringSplitOptions.None)

                If epParts.Length = 2 AndAlso apiParts.Length = 2 AndAlso respParts.Length = 2 Then
                    postEndpoint = epParts(0)
                    getEndpointTemplate = epParts(1)
                    postAPICall = apiParts(0)
                    getAPICallTemplate = apiParts(1)
                    postResponseKey = respParts(0)
                    getResponseKey = respParts(1)
                    multiCall = True
                Else
                    postEndpoint = Endpoint
                    postAPICall = APICall
                    postResponseKey = ResponseKey
                End If

                Endpoint = postEndpoint
                APICall = postAPICall
                ResponseKey = postResponseKey

                Dim useGetMethod As Boolean = False
                If Endpoint.StartsWith("get:", System.StringComparison.OrdinalIgnoreCase) Then
                    useGetMethod = True
                    ' Remove "get:" prefix.
                    Endpoint = Endpoint.Substring(4)
                End If

                ' Replace placeholders in the request body
                Dim requestBody As String = APICall
                requestBody = requestBody.Replace("{model}", ModelValue)
                requestBody = requestBody.Replace("{ownsessionid}", OwnSessionID)
                requestBody = requestBody.Replace("{promptsystem}", CleanString(promptSystem))
                requestBody = requestBody.Replace("{promptuser}", CleanString(promptUser))
                requestBody = requestBody.Replace("{userinstruction}", CleanString(AddUserPrompt))
                requestBody = requestBody.Replace("{temperature}", TemperatureValue)

                ' Handle Tooling instructions if applicable

                If ToolExecution Then
                    requestBody = requestBody.Replace(LLM_APICall_Placeholder_ToolDefinitions, context.INI_APICall_ToolInstructions_2)
                    requestBody = requestBody.Replace(LLM_APICall_Placeholder_ToolResponses, context.INI_APICall_ToolResponses_2)
                Else
                    ' Remove tool placeholders 
                    requestBody = requestBody.Replace(LLM_APICall_Placeholder_ToolDefinitions, "")
                    requestBody = requestBody.Replace(LLM_APICall_Placeholder_ToolResponses, "")
                End If

                ' Handle object upload if configured

                Dim ObjectCall As String = If(UseSecondAPI, context.INI_APICall_Object_2, context.INI_APICall_Object)
                Dim requiresMultipart As Boolean = ObjectCall.ToLowerInvariant().Trim().StartsWith("multipart:")

                Dim fileName As String = ""
                Dim fileBytes() As Byte = Nothing
                Dim mimeType As String = ""
                Dim multipart As New System.Net.Http.MultipartFormDataContent()
                Dim fileFieldName As String = "file" ' Default if not specified

                If Not String.IsNullOrWhiteSpace(ObjectCall) AndAlso Not String.IsNullOrWhiteSpace(FileObject) Then

                    Try
                        Dim encodedData As String = ""

                        If FileObject.Equals("clipboard", StringComparison.OrdinalIgnoreCase) Then
                            Dim mime As String = Nothing, data As String = Nothing
                            If Not TryGetClipboardObject(mime, data) Then
                                If Not Hidesplash Then ShowCustomMessageBox("No supported data found in the clipboard.") Else Return "No supported data found in the clipboard."
                                Return ""
                            End If
                            mimeType = FixMimeType(mime)
                            If Not requiresMultipart Then
                                encodedData = data
                            Else
                                fileBytes = System.Convert.FromBase64String(data)
                                fileName = "clipboard.png"
                            End If
                        Else
                            ' Standard case: file processed via MimeHelper.
                            Dim mresult = MimeHelper.GetFileMimeTypeAndBase64(FileObject)
                            mimeType = FixMimeType(mresult.MimeType.Trim())
                            If Not requiresMultipart Then
                                encodedData = mresult.EncodedData.Trim()
                            Else
                                fileBytes = System.IO.File.ReadAllBytes(FileObject)
                                fileName = System.IO.Path.GetFileName(FileObject)
                            End If
                        End If

                        ' --- Handle multiple ObjectCall variants with MIME-type filters ---
                        ' Format: [mime1,mime2]variant1¦[mime3]variant2¦variant3 (unfiltered fallback)
                        Dim variantSep As String = "¦"
                        Dim objectCallVariants() As String = ObjectCall.Split(New String() {variantSep}, StringSplitOptions.None)

                        Dim selectedObjectCall As String = Nothing
                        Dim unfilteredObjectCall As String = Nothing

                        For Each objCallEntry As String In objectCallVariants
                            Dim trimmedEntry As String = objCallEntry.Trim()
                            If String.IsNullOrEmpty(trimmedEntry) Then Continue For

                            ' Check if entry starts with a MIME filter [...]
                            If trimmedEntry.StartsWith("[") Then
                                Dim closeBracketIdx As Integer = trimmedEntry.IndexOf("]"c)
                                If closeBracketIdx > 0 Then
                                    ' Extract MIME types from filter
                                    Dim filterPart As String = trimmedEntry.Substring(1, closeBracketIdx - 1)
                                    Dim allowedMimes() As String = filterPart.Split(","c)

                                    ' Check if current mimeType matches any allowed MIME type
                                    For Each allowedMime As String In allowedMimes
                                        Dim normalizedAllowed As String = FixMimeType(allowedMime.Trim()).ToLowerInvariant()
                                        If mimeType.ToLowerInvariant() = normalizedAllowed Then
                                            ' Match found - use the entry content after the filter
                                            selectedObjectCall = trimmedEntry.Substring(closeBracketIdx + 1).TrimStart()
                                            Exit For
                                        End If
                                    Next

                                    If selectedObjectCall IsNot Nothing Then Exit For
                                Else
                                    ' Malformed filter (no closing bracket) - treat as unfiltered
                                    If unfilteredObjectCall Is Nothing Then unfilteredObjectCall = trimmedEntry
                                End If
                            Else
                                ' No filter - this is an unfiltered fallback entry
                                If unfilteredObjectCall Is Nothing Then unfilteredObjectCall = trimmedEntry
                            End If
                        Next

                        ' Use selected entry, or fall back to unfiltered entry
                        If selectedObjectCall Is Nothing Then
                            selectedObjectCall = unfilteredObjectCall
                        End If

                        ' If no matching entry found, show error
                        If selectedObjectCall Is Nothing Then
                            Dim errorMsg As String = $"Error: The file/object provided is of an unsupported MIME type ({mimeType}). None of the configured variants accept this type."
                            If Hidesplash Then
                                Return errorMsg
                            Else
                                ShowCustomMessageBox(errorMsg)
                                Return ""
                            End If
                        End If

                        ' Use the selected entry as the effective ObjectCall
                        ObjectCall = selectedObjectCall
                        requiresMultipart = ObjectCall.ToLowerInvariant().Trim().StartsWith("multipart:")

                        requestBody = requestBody.Replace(LLM_APICall_Placeholder_Objectcall, ObjectCall)

                        If Not requiresMultipart Then
                            requestBody = requestBody.Replace("{mimetype}", mimeType)
                            requestBody = requestBody.Replace("{encodeddata}", encodedData)
                        Else
                            ' Prepare variables

                            Dim config As String

                            ' Remove "multipart:" prefix.
                            config = ObjectCall.Substring("multipart:".Length)

                            ' Split on unescaped semicolons (support ;; as escape for ;).
                            Dim parts As New List(Of String)()
                            Dim current As String = ""
                            Dim i As Integer = 0
                            While i < config.Length
                                If config(i) = ";"c Then
                                    If i + 1 < config.Length AndAlso config(i + 1) = ";"c Then
                                        current &= ";"c  ' Escaped semicolon
                                        i += 1
                                    Else
                                        parts.Add(current)
                                        current = ""
                                    End If
                                Else
                                    current &= config(i)
                                End If
                                i += 1
                            End While
                            If current.Length > 0 Then parts.Add(current)

                            ' Parse fields and add to multipart.
                            For Each part In parts
                                Dim idx As Integer = part.IndexOf(":")
                                If idx > 0 Then
                                    Dim fieldName As String = part.Substring(0, idx).Trim()
                                    Dim fieldValue As String = part.Substring(idx + 1).Trim()
                                    If fieldName.Equals("filefield", StringComparison.OrdinalIgnoreCase) Then
                                        fileFieldName = fieldValue
                                    Else
                                        ' Replace placeholders as needed.
                                        fieldValue = fieldValue.Replace("{model}", ModelValue) _
                                                               .Replace("{promptsystem}", CleanString(promptSystem)) _
                                                               .Replace("{promptuser}", CleanString(promptUser)) _
                                                               .Replace("{temperature}", TemperatureValue) _
                                                               .Replace("{ownsessionid}", OwnSessionID) _
                                                               .Replace("{userinstruction}", CleanString(AddUserPrompt))

                                        multipart.Add(New System.Net.Http.StringContent(fieldValue, System.Text.Encoding.UTF8), fieldName)
                                    End If
                                End If
                            Next
                        End If

                    Catch ex As System.Exception
                        If Not Hidesplash Then ShowCustomMessageBox($"Error encoding '{FileObject}': {ex.Message}") Else Return $"Error encoding '{FileObject}': {ex.Message}"
                        Return ""
                    End Try

                End If

                requestBody = requestBody.Replace("{objectcall}", "")

                Dim Returnvalue As String = ""

                Try

                    ' Configure HttpClient with timeout.
                    Using handler As New System.Net.Http.HttpClientHandler()
                        handler.UseProxy = True
                        handler.Proxy = System.Net.WebRequest.DefaultWebProxy
                        handler.Proxy.Credentials = System.Net.CredentialCache.DefaultCredentials
                        Using client As New System.Net.Http.HttpClient(handler)
                            client.Timeout = TimeSpan.FromMilliseconds(TimeoutValue)

                            ' Send the request.
                            Try

                                Dim maxRetries As Integer = 3
                                Dim delayIntervals As Integer() = {5000, 10000, 30000} ' delays in milliseconds
                                Dim responseText As String = ""
                                Dim lastContentType As String = Nothing

                                For attempt As Integer = 0 To maxRetries
                                    ' On retries, wait the specified delay before sending a new request.
                                    If attempt > 0 Then
                                        If Not Hidesplash Then
                                            splash.RestartCountdown(timeoutSeconds, "Slowing down due to AI...")
                                        End If
                                        Await System.Threading.Tasks.Task.Delay(delayIntervals(attempt - 1), ct)
                                    End If

                                    Dim requestContent As System.Net.Http.HttpContent
                                    If requiresMultipart Then
                                        ' Recreate MultipartFormDataContent on each attempt to avoid duplicate parts.
                                        Dim attemptMultipart As New System.Net.Http.MultipartFormDataContent()

                                        ' Re-add cached string fields (parsed earlier outside the loop).
                                        For Each part In multipart
                                            If TypeOf part Is System.Net.Http.StringContent Then
                                                Dim name As String = part.Headers.ContentDisposition?.Name?.Trim(""""c)
                                                If Not String.IsNullOrEmpty(name) Then
                                                    Dim val As String = Await part.ReadAsStringAsync()
                                                    attemptMultipart.Add(New System.Net.Http.StringContent(val, System.Text.Encoding.UTF8), name)
                                                End If
                                            End If
                                        Next

                                        ' Add file content.
                                        Dim fileContent As New System.Net.Http.ByteArrayContent(fileBytes)
                                        fileContent.Headers.ContentType = New System.Net.Http.Headers.MediaTypeHeaderValue(mimeType)
                                        attemptMultipart.Add(fileContent, fileFieldName, fileName)
                                        requestContent = attemptMultipart
                                    Else
                                        requestContent = New System.Net.Http.StringContent(requestBody, System.Text.Encoding.UTF8, "application/json")
                                    End If

                                    Dim response As System.Net.Http.HttpResponseMessage

                                    splash.RestartCountdown(timeoutSeconds)

                                    If useGetMethod Then
                                        ' Build a GET request (no request body is sent here).
                                        Dim getReq As New System.Net.Http.HttpRequestMessage(System.Net.Http.HttpMethod.Get, Endpoint)
                                        If Not String.IsNullOrEmpty(HeaderA) AndAlso Not String.IsNullOrEmpty(HeaderB) Then
                                            If Not getReq.Headers.Contains(HeaderA) Then
                                                getReq.Headers.Add(HeaderA, HeaderB)
                                            End If
                                        End If
                                        If context.INI_APIDebug Then
                                            Debug.WriteLine($"SENT TO API as GET ({Endpoint}):{Environment.NewLine}{String.Empty}") ' No body.
                                        End If
                                        response = Await client.SendAsync(getReq, System.Net.Http.HttpCompletionOption.ResponseContentRead, ct).ConfigureAwait(False)
                                    Else
                                        ' Standard POST request.
                                        If Not String.IsNullOrEmpty(HeaderA) AndAlso Not String.IsNullOrEmpty(HeaderB) Then
                                            If Not client.DefaultRequestHeaders.Contains(HeaderA) Then
                                                client.DefaultRequestHeaders.Add(HeaderA, HeaderB)
                                            End If
                                        End If
                                        If context.INI_APIDebug Then
                                            If requiresMultipart Then
                                                Dim multipartInfo As New System.Text.StringBuilder()
                                                multipartInfo.AppendLine($"SENT TO API ({Endpoint}) as multipart:")
                                                ' Cast requestContent back to MultipartFormDataContent for enumeration.
                                                For Each content As System.Net.Http.HttpContent In DirectCast(requestContent, System.Net.Http.MultipartFormDataContent)
                                                    Dim contentName As String = ""
                                                    If content.Headers.ContentDisposition IsNot Nothing Then
                                                        contentName = content.Headers.ContentDisposition.Name
                                                        If Not String.IsNullOrEmpty(contentName) Then
                                                            contentName = contentName.Trim(""""c)
                                                        End If
                                                    End If
                                                    If TypeOf content Is System.Net.Http.StringContent Then
                                                        Dim val As String = Await content.ReadAsStringAsync()
                                                        multipartInfo.AppendLine($" - {contentName}: '{val}'")
                                                    ElseIf TypeOf content Is System.Net.Http.ByteArrayContent Then
                                                        Dim fileNamex As String = ""
                                                        If content.Headers.ContentDisposition IsNot Nothing Then
                                                            fileNamex = content.Headers.ContentDisposition.FileName?.Trim(""""c)
                                                        End If
                                                        multipartInfo.AppendLine($" - {contentName}: <file: '{fileNamex}', type: {content.Headers.ContentType}>")
                                                    Else
                                                        multipartInfo.AppendLine($" - {contentName}: <unknown part type>")
                                                    End If
                                                Next
                                                Debug.WriteLine(multipartInfo.ToString())
                                                Try
                                                    Dim desktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
                                                    Dim debugFilePath As String = System.IO.Path.Combine(desktopPath, "RI_Debug_Sent.json")
                                                    System.IO.File.WriteAllText(debugFilePath, multipartInfo.ToString())
                                                Catch
                                                    ' Silent fail
                                                End Try
                                            Else
                                                Debug.WriteLine($"SENT TO API ({Endpoint}):{Environment.NewLine}{requestBody}")
                                                Try
                                                    Dim desktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
                                                    Dim debugFilePath As String = System.IO.Path.Combine(desktopPath, "RI_Debug_Sent.json")
                                                    System.IO.File.WriteAllText(debugFilePath, requestBody)
                                                Catch
                                                    ' Silent fail
                                                End Try

                                            End If
                                        End If
                                        response = Await client.PostAsync(Endpoint, requestContent, ct).ConfigureAwait(False)
                                    End If

                                    lastContentType = If(response.Content IsNot Nothing AndAlso response.Content.Headers IsNot Nothing AndAlso response.Content.Headers.ContentType IsNot Nothing,
                                             response.Content.Headers.ContentType.ToString(),
                                             Nothing)

                                    If response.IsSuccessStatusCode Then
                                        responseText = Await response.Content.ReadAsStringAsync().ConfigureAwait(False)
                                        Exit For

                                    ElseIf response.StatusCode = 429 Then
                                        If attempt = maxRetries Then
                                            If Not Hidesplash Then ShowCustomMessageBox($"HTTP Error {response.StatusCode} when accessing the LLM endpoint: This error is typically either because (1) the server resource is exhausted on the side of the provider (retry it later or reduce the work load; {AN} already tried to slow down and wait, but this could not overcome the overload condition), or (2) you have not yet correctly configured your service account (e.g., with OpenAI API, you have to have a credit card registered and an amount entered before you generate your API key).") Else Return $"HTTP Error {response.StatusCode} when accessing the LLM endpoint: This error is typically either because (1) the server resource is exhausted on the side of the provider (retry it later or reduce the work load; {AN} already tried to slow down and wait, but this could not overcome the overload condition), or (2) you have not yet correctly configured your service account (e.g., with OpenAI API, you have to have a credit card registered and an amount entered before you generate your API key)."
                                            Return ""
                                        End If
                                        Continue For
                                    Else
                                        Dim errorContent As String = Await response.Content.ReadAsStringAsync().ConfigureAwait(False)
                                        If Not Hidesplash Then ShowCustomMessageBox($"HTTP Error {response.StatusCode} when accessing the LLM endpoint: {errorContent}")
                                        Return ""
                                    End If
                                Next

                                If Not Hidesplash Then
                                    splash.RestartCountdown(timeoutSeconds, "Waiting for the AI to respond...")
                                End If

                                If context.INI_APIDebug Then
                                    Debug.WriteLine($"RECEIVED FROM API:{Environment.NewLine}{responseText}")
                                    Try
                                        Dim desktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
                                        Dim debugFilePath As String = System.IO.Path.Combine(desktopPath, "RI_Debug_Received.json")
                                        System.IO.File.WriteAllText(debugFilePath, responseText)
                                    Catch
                                        ' Silent fail
                                    End Try
                                End If

                                ' Normalize/validate response BEFORE JSON parse.
                                Dim respTrim As String = If(responseText, String.Empty)

                                ' SSE normalization: drop keepalive/comment lines starting with ":" and unwrap "data:" frames.
                                Dim sseProbe As String = respTrim.TrimStart()
                                If sseProbe.StartsWith(":", StringComparison.Ordinal) OrElse sseProbe.StartsWith("data:", StringComparison.OrdinalIgnoreCase) Then
                                    Dim sb As New System.Text.StringBuilder()
                                    Dim lfNormalized As String = respTrim.Replace(vbCrLf, vbLf)
                                    Dim parts As String() = lfNormalized.Split(New String() {vbLf}, StringSplitOptions.None)

                                    For Each raw In parts
                                        Dim t As String = If(raw, String.Empty).Trim()
                                        If t.Length = 0 Then Continue For

                                        If t.StartsWith(":", StringComparison.Ordinal) Then
                                            Continue For
                                        End If

                                        If t.StartsWith("data:", StringComparison.OrdinalIgnoreCase) Then
                                            Dim payload As String = t.Substring(5).Trim()
                                            If payload.Length > 0 AndAlso Not payload.Equals("[DONE]", StringComparison.OrdinalIgnoreCase) Then
                                                sb.Append(payload)
                                            End If
                                            Continue For
                                        End If

                                        sb.AppendLine(raw)
                                    Next

                                    responseText = sb.ToString().Trim()
                                End If

                                respTrim = If(responseText, String.Empty).TrimStart()

                                If String.IsNullOrWhiteSpace(respTrim) Then
                                    If Not Hidesplash Then ShowCustomMessageBox("Empty response from the endpoint.")
                                    Return ""
                                End If

                                If Not (respTrim.StartsWith("{") OrElse respTrim.StartsWith("[")) Then
                                    Dim preview As String = If(respTrim.Length > 400, respTrim.Substring(0, 400) & "...", respTrim)
                                    If Not Hidesplash Then ShowCustomMessageBox($"Endpoint returned non‑JSON (Content-Type={lastContentType}). First bytes: {preview}")
                                    Return ""
                                End If

                                ' Process the response.
                                Dim root As Newtonsoft.Json.Linq.JToken = Newtonsoft.Json.Linq.JToken.Parse(responseText)
                                LogTokenSpending(root, TokenCountString, AddUserPrompt)

                                If multiCall Then

                                    ' 1) Split all keys and extract values from the POST response.
                                    Dim keys() As String = postResponseKey.Split(New String() {sep2}, StringSplitOptions.None)
                                    Dim extracted As New Dictionary(Of String, String)

                                    For Each key As String In keys
                                        Dim val As String = CType(root, Newtonsoft.Json.Linq.JObject).SelectToken(key)?.ToString()
                                        If String.IsNullOrEmpty(val) Then
                                            Throw New System.Exception($"POST response contains no value for '{key}'.")
                                        End If
                                        extracted(key) = val
                                    Next

                                    ' 2) Fill placeholders in GET endpoint.
                                    Dim rawGetEndpoint As String = getEndpointTemplate
                                    rawGetEndpoint = rawGetEndpoint.Replace("{model}", ModelValue)
                                    rawGetEndpoint = rawGetEndpoint.Replace("{ownsessionid}", OwnSessionID)
                                    rawGetEndpoint = rawGetEndpoint.Replace("{apikey}", If(UseSecondAPI, context.DecodedAPI_2, context.DecodedAPI))
                                    For Each kvp As KeyValuePair(Of String, String) In extracted
                                        rawGetEndpoint = rawGetEndpoint.Replace("{" & kvp.Key & "}", kvp.Value)
                                    Next

                                    ' 3) Fill placeholders in optional GET body.
                                    Dim rawGetBody As String = getAPICallTemplate
                                    If Not String.IsNullOrWhiteSpace(rawGetBody) Then
                                        rawGetBody = rawGetBody.Replace("{model}", ModelValue)
                                        rawGetBody = rawGetBody.Replace("{ownsessionid}", OwnSessionID)
                                        rawGetBody = rawGetBody.Replace("{apikey}", If(UseSecondAPI, context.DecodedAPI_2, context.DecodedAPI))
                                        For Each kvp As KeyValuePair(Of String, String) In extracted
                                            rawGetBody = rawGetBody.Replace("{" & kvp.Key & "}", kvp.Value)
                                        Next
                                    End If

                                    ' 4) Prepare GET request and set header.
                                    Dim getReq As New System.Net.Http.HttpRequestMessage(System.Net.Http.HttpMethod.Get, rawGetEndpoint)

                                    If Not String.IsNullOrEmpty(HeaderA) AndAlso Not String.IsNullOrEmpty(HeaderB) Then
                                        getReq.Headers.Add(HeaderA, HeaderB)
                                    End If

                                    If Not String.IsNullOrWhiteSpace(rawGetBody) Then
                                        getReq.Content = New System.Net.Http.StringContent(rawGetBody, System.Text.Encoding.UTF8, "application/json")
                                    End If

                                    If context.INI_APIDebug Then
                                        Debug.WriteLine($"SENT TO API as GET ({rawGetEndpoint}):{Environment.NewLine}{rawGetBody}")
                                        Try
                                            Dim desktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
                                            Dim debugFilePath As String = System.IO.Path.Combine(desktopPath, "RI_Debug_Sent_Get.json")
                                            System.IO.File.WriteAllText(debugFilePath, rawGetBody)
                                        Catch
                                            ' Silent fail
                                        End Try
                                    End If

                                    splash.RestartCountdown(timeoutSeconds)

                                    ' 5) Send GET request.
                                    Dim getResponseText As String
                                    Using getResp = Await client.SendAsync(getReq, System.Net.Http.HttpCompletionOption.ResponseContentRead, ct).ConfigureAwait(False)
                                        If getResp.IsSuccessStatusCode Then
                                            getResponseText = Await getResp.Content.ReadAsStringAsync().ConfigureAwait(False)
                                        Else
                                            Dim err = Await getResp.Content.ReadAsStringAsync().ConfigureAwait(False)
                                            Throw New System.Exception($"HTTP GET Error {getResp.StatusCode}: {err}")
                                        End If
                                    End Using

                                    If context.INI_APIDebug Then
                                        Debug.WriteLine($"RECEIVED FROM API (GET):{Environment.NewLine}{getResponseText}")
                                        Try
                                            Dim desktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
                                            Dim debugFilePath As String = System.IO.Path.Combine(desktopPath, "RI_Debug_Received_GET.json")
                                            System.IO.File.WriteAllText(debugFilePath, responseText)
                                        Catch
                                            ' Silent fail
                                        End Try
                                    End If

                                    ' 6) Process GET response using the same extraction logic as POST-only mode.
                                    Dim root2 As Newtonsoft.Json.Linq.JToken = Newtonsoft.Json.Linq.JToken.Parse(getResponseText)

                                    Select Case root2.Type
                                        Case Newtonsoft.Json.Linq.JTokenType.Object
                                            Dim obj2 As Newtonsoft.Json.Linq.JObject = CType(root2, Newtonsoft.Json.Linq.JObject)
                                            Returnvalue = HandleObject(obj2, getResponseKey, getResponseText, RKMode, DetectToolCall)

                                        Case Newtonsoft.Json.Linq.JTokenType.Array
                                            ' If template has a loop, process entire array at once
                                            Dim hasLoop = Regex.IsMatch(ResponseKey, "\{\%\s*for\s+", RegexOptions.Singleline)
                                            If hasLoop Then
                                                ' Pass entire response to template formatter which will handle the array
                                                Returnvalue = JsonTemplateFormatter.FormatJsonWithTemplate(responseText, ResponseKey)
                                            Else
                                                ' Legacy behavior: process each item separately
                                                For Each item As Newtonsoft.Json.Linq.JToken In CType(root, Newtonsoft.Json.Linq.JArray)
                                                    If item.Type = Newtonsoft.Json.Linq.JTokenType.Object Then
                                                        Returnvalue &= HandleObject(CType(item, Newtonsoft.Json.Linq.JObject),
                                                        ResponseKey, responseText, RKMode, DetectToolCall)
                                                    End If
                                                Next
                                            End If

                                        Case Else
                                            If Not Hidesplash Then ShowCustomMessageBox($"Unexpected JSON root type: {root2.Type} ({getResponseText})")
                                    End Select

                                Else
                                    ' POST-only processing.
                                    Select Case root.Type
                                        Case Newtonsoft.Json.Linq.JTokenType.Object
                                            Dim jsonObject As Newtonsoft.Json.Linq.JObject = CType(root, Newtonsoft.Json.Linq.JObject)
                                            Returnvalue = HandleObject(jsonObject, ResponseKey, responseText, RKMode, DetectToolCall)

                                        Case Newtonsoft.Json.Linq.JTokenType.Array
                                            ' If template has a loop, process entire array at once
                                            Dim hasLoop = Regex.IsMatch(ResponseKey, "\{\%\s*for\s+", RegexOptions.Singleline)
                                            If hasLoop Then
                                                ' Pass entire response to template formatter which will handle the array
                                                Returnvalue = JsonTemplateFormatter.FormatJsonWithTemplate(responseText, ResponseKey)
                                            Else
                                                ' Legacy behavior: process each item separately
                                                For Each item As Newtonsoft.Json.Linq.JToken In CType(root, Newtonsoft.Json.Linq.JArray)
                                                    If item.Type = Newtonsoft.Json.Linq.JTokenType.Object Then
                                                        Returnvalue &= HandleObject(CType(item, Newtonsoft.Json.Linq.JObject),
                                                            ResponseKey, responseText, RKMode, DetectToolCall)
                                                    End If
                                                Next
                                            End If

                                        Case Else
                                            If Not Hidesplash Then ShowCustomMessageBox($"Unexpected JSON root type: {root.Type} ({responseText})")
                                    End Select
                                End If

                            Catch ex As System.Net.Http.HttpRequestException When Not ct.IsCancellationRequested
                                If Not Hidesplash Then ShowCustomMessageBox($"An HTTP request exception occurred: {ex.Message} when accessing the LLM endpoint (2).")
                            Catch ex As TaskCanceledException When ct.IsCancellationRequested
                                Throw New OperationCanceledException(ct)
                            Catch ex As TaskCanceledException When Not ct.IsCancellationRequested
                                If Not Hidesplash Then splash.Close()
                                If Not Hidesplash Then ShowCustomMessageBox($"The request to the endpoint timed out. Please try again or increase the timeout setting.")
                            Catch ex As System.Exception When Not ct.IsCancellationRequested
                                If Not Hidesplash Then splash.Close()
                                If Not Hidesplash Then ShowCustomMessageBox($"The response from the endpoint resulted in an error: {ex.Message}")
                            End Try
                        End Using ' Dispose HttpClient
                    End Using ' Dispose HttpClientHandler
                Catch ex As OperationCanceledException
                    If Not Hidesplash Then ShowCustomMessageBox("Request canceled.")
                    Return ""
                Finally
                    cts.Dispose()
                    If Not Hidesplash Then splash.Close()
                End Try
                If DoubleS Then
                    Returnvalue = Returnvalue.Replace(ChrW(223), "ss") ' Replace German sharp-S if needed
                End If
                If context.INI_NoDash Then
                    Returnvalue = System.Text.RegularExpressions.Regex.Replace(
                                        Returnvalue,
                                        "([A-Za-z0-9,])\s*—\s*([A-Za-z0-9])",
                                        "$1 – $2"
                                    )
                End If
                If context.INI_Clean Then
                    Returnvalue = System.Text.RegularExpressions.Regex.Replace(
                                    Returnvalue,
                                    "(?<=\S) {2,}",
                                    " "
                                )
                    Returnvalue = RemoveHiddenMarkers(Returnvalue)
                End If

                If NoThink Then

                    If Not String.IsNullOrEmpty(Returnvalue) Then
                        Dim tag As String = "</THINK>"
                        Dim idx As Integer = Returnvalue.LastIndexOf(tag, StringComparison.OrdinalIgnoreCase)
                        If idx >= 0 Then
                            Dim startPos As Integer = idx + tag.Length
                            If startPos >= Returnvalue.Length Then
                                Returnvalue = String.Empty
                            Else
                                ' Remove everything up to and including the last </THINK>,
                                ' then strip any leading CR/LF/whitespace before the real text.
                                Returnvalue = Returnvalue.Substring(startPos).TrimStart()
                            End If
                        End If
                    End If

                End If

                If AnonActive Then Returnvalue = ReidentifyText(Returnvalue)

                Return Returnvalue

            Catch ex As System.Exception

#If DEBUG Then
                Debug.WriteLine("Error: " & ex.Message)
                Debug.WriteLine("Stacktrace: " & ex.StackTrace)

                System.Diagnostics.Debugger.Break()
#End If


                If Not Hidesplash Then ShowCustomMessageBox($"An unexpected error occurred when accessing the LLM endpoint: {ex.Message}") Else Return $"An unexpected error occurred when accessing the LLM endpoint: {ex.Message}"
                Return ""
            Finally
                If Not Hidesplash Then
                    splash.Close()
                End If
            End Try
        End Function

        ''' <summary>
        ''' Generates a GUID-based unique id string without separators (32 hex characters).
        ''' </summary>
        ''' <returns>A GUID string in "N" format; returns XdatetimeX if GUID generation fails.</returns>
        Public Shared Function GenerateUniqueId() As String
            Try
                Return System.Guid.NewGuid().ToString("N")
            Catch ex As System.Exception
                ' Fallback: timestamp + random (extremely unlikely path)
                Return DateTime.UtcNow.Ticks.ToString("X") & (New Random()).Next().ToString("X")
            End Try
        End Function

        ''' <summary>
        ''' Opens an editor window for the prompt log stored in settings and saves edits back to settings.
        ''' </summary>
        Public Shared Sub ShowAndEditPromptLog()
            Const MaxItems As Integer = PromptLogCap
            Const SepLine As String = "----- Prompt Entry Separator -----"

            Try
                ' 1) Get current log (ensure not Nothing)
                Dim log = My.Settings.PromptLog
                If log Is Nothing Then
                    log = New System.Collections.Specialized.StringCollection()
                    My.Settings.PromptLog = log
                End If

                ' 2) Build editable body (preserve multi-line entries; separate by a clear separator)
                Dim body As New StringBuilder()
                For i As Integer = 0 To log.Count - 1
                    If i > 0 Then
                        body.AppendLine()
                        body.AppendLine(SepLine)
                        body.AppendLine()
                    End If
                    body.Append(log(i))
                Next

                ' 3) Show editor; if user cancels, ShowCustomWindow should return Nothing
                Dim intro As String = $"Prompt Log (most recent first). Edit entries freely. Keep the line '{SepLine}' between items."
                Dim finalRemark As String = $"OK saves changes (at least one entry must exist). Cancel aborts. Limit is {MaxItems} items (older ones are dropped)."
                Dim result As String = ShowCustomWindow(intro, body.ToString(), finalRemark, AN, NoRTF:=True, Getfocus:=True)

                If result Is Nothing OrElse result = "" Then
                    Exit Sub ' canceled
                End If

                ' 4) Parse back using the separator (preserve internal newlines in each entry)
                Dim parts = result.Split(New String() {vbCrLf & SepLine & vbCrLf}, StringSplitOptions.None)

                Dim updated = New System.Collections.Specialized.StringCollection()
                For Each part In parts
                    Dim trimmed = part.Trim()
                    If trimmed.Length > 0 Then
                        updated.Add(trimmed)
                    End If
                Next

                ' 5) Enforce LIFO cap (keep first MaxItems; they’re most recent-first)
                While updated.Count > MaxItems
                    updated.RemoveAt(updated.Count - 1)
                End While

                ' 6) Save
                My.Settings.PromptLog = updated
                My.Settings.Save()

            Catch
                ' swallow to avoid UI disruption
            End Try
        End Sub

        ''' <summary>
        ''' Logs prompt text to settings and extracts token usage metrics from the JSON response using <paramref name="tokenCountString"/>.
        ''' </summary>
        ''' <param name="root">Parsed JSON response token used for SelectToken reads.</param>
        ''' <param name="tokenCountString">
        ''' Semicolon-separated list of token usage segment names, optionally extended with a numeric multiplier and currency code.
        ''' </param>
        ''' <param name="prompt">Prompt/instruction text recorded in the prompt log and in the cost log (when configured).</param>
        Private Shared Sub LogTokenSpending(ByRef root As JToken, tokenCountString As String, prompt As String)

            ' 0) only run if there's something to log

            If String.IsNullOrWhiteSpace(prompt) Then
                Return
            End If

            ' A) Log latest prompt in My.Settings (LIFO, keep last 25)
            Try
                Dim log = My.Settings.PromptLog
                If log Is Nothing Then
                    log = New System.Collections.Specialized.StringCollection()
                    My.Settings.PromptLog = log
                End If

                Dim ts = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)
                log.Insert(0, $"{ts} | {prompt}")

                While log.Count > PromptLogCap
                    log.RemoveAt(log.Count - 1)
                End While

                My.Settings.Save()
            Catch
                ' ignore settings errors
            End Try

            If String.IsNullOrWhiteSpace(tokenCountString) Then
                Return
            End If

            ' 1) split & trim all parts
            Dim parts() As String
            Try
                parts = tokenCountString _
            .Split(";"c) _
            .Select(Function(p) p.Trim()) _
            .ToArray()
            Catch
                Return
            End Try

            ' 2) determine which parts are segment names vs multiplier & currency
            Dim segmentNames As String()
            Dim multiplier As Double? = Nothing
            Dim currencyCode As String = String.Empty

            If parts.Length >= 3 Then
                segmentNames = parts.Take(parts.Length - 2).ToArray()

                Dim rawMult = parts(parts.Length - 2)
                Dim parsedMult As Double = 0
                If Double.TryParse(rawMult, NumberStyles.Float, CultureInfo.InvariantCulture, parsedMult) Then
                    multiplier = parsedMult
                    currencyCode = parts(parts.Length - 1)
                Else
                    ' invalid multiplier → skip cost line
                    multiplier = Nothing
                End If
            Else
                segmentNames = parts
            End If

            ' 3) extract each token value, auto‑prefix usageMetadata if needed
            Dim segmentValues As New Dictionary(Of String, Long)()
            Dim totalTokens As Long = 0

            For Each name In segmentNames
                Dim path = If(name.Contains("."), name, $"usageMetadata.{name}")
                Dim tok As String = Nothing

                Try
                    tok = root.SelectToken(path)?.ToString()
                Catch
                    Return  ' silent exit on any JSON path error
                End Try

                If String.IsNullOrEmpty(tok) Then
                    Return
                End If

                Dim n As Long = 0
                If Not Long.TryParse(tok, n) Then
                    Return
                End If

                segmentValues(name) = n
                totalTokens += n
            Next

            ' 4) build the log entry
            Dim nowStamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)
            Dim sb As New StringBuilder()
            sb.AppendLine(nowStamp)

            ' truncate prompt + ellipsis if needed
            Dim promptText = prompt
            If promptText.Length > 2048 Then
                promptText = promptText.Substring(0, 2048) & "…"
            End If
            sb.AppendLine("Prompt: " & promptText)

            sb.Append("Token counts: ")
            For Each kvp In segmentValues
                sb.Append($"{kvp.Value} ({kvp.Key}), ")
            Next
            If sb.Length >= 2 Then sb.Length -= 2  ' remove trailing comma+space
            sb.AppendLine()

            sb.AppendLine($"Total tokens: {totalTokens} (total)")

            If multiplier.HasValue Then
                Dim costValue = Math.Round(totalTokens * multiplier.Value / 1000, 2)
                sb.AppendLine($"Value: {currencyCode} {costValue} ({currencyCode} {multiplier.Value}/1000 tokens)")
            End If

            sb.AppendLine()  ' blank line separator            
            Dim entryText = sb.ToString()

            ' 5) determine file path on Desktop
            Dim desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
            Dim filePath = Path.Combine(desktop, "redink-cost.txt")

            ' 6) write with exclusive lock & retry
            Const maxRetries As Integer = 5
            Const delayMs As Integer = 100
            Dim written As Boolean = False

            For attempt As Integer = 1 To maxRetries
                Try
                    Using fs As New FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None)
                        ' write header if file is new
                        If fs.Length = 0 Then
                            Dim header = $"RED INK FREESTYLE TOKEN SPENDING LOG (USER: {Environment.UserName})" & Environment.NewLine & Environment.NewLine
                            Dim hb() = Encoding.UTF8.GetBytes(header)
                            fs.Write(hb, 0, hb.Length)
                        End If

                        ' append entry
                        fs.Seek(0, SeekOrigin.End)
                        Dim eb() = Encoding.UTF8.GetBytes(entryText)
                        fs.Write(eb, 0, eb.Length)
                        fs.Flush()
                    End Using

                    written = True
                    Exit For

                Catch
                    ' wait a bit then retry
                    Thread.Sleep(delayMs)
                End Try
            Next

            ' 7) if all attempts fail, show error dialog
            If Not written Then
                ShowCustomMessageBox(
            $"Error writing log file '{filePath}'.{Environment.NewLine}" &
            $"Entry was:{Environment.NewLine}{entryText}"
        )
            End If
        End Sub

        ''' <summary>
        ''' Extracts response content from a JSON object using the configured response key/template and handles error payloads.
        ''' </summary>
        ''' <param name="jsonObject">Parsed JSON object returned by the endpoint.</param>
        ''' <param name="ResponseKey">Response key/template used by <c>JsonTemplateFormatter</c> or the literal string "JSON".</param>
        ''' <param name="ResponseText">Original response text used for error messages and "JSON" passthrough.</param>
        ''' <param name="RKMode">Response extraction mode passed through to <c>JsonTemplateFormatter</c>.</param>
        ''' <returns>Extracted response text; empty string on handled error.</returns>
        Private Shared Function HandleObject(jsonObject As Newtonsoft.Json.Linq.JObject, ResponseKey As String, ResponseText As String, RKMode As Integer, DetectToolCall As String) As String

            ' Extract the "error" segment
            Dim text As String = FindJsonProperty(jsonObject, "error")

            If Not String.IsNullOrEmpty(text) Then
                text = FindJsonProperty(jsonObject, "message")
                ShowCustomMessageBox($"The LLM API generated the following error message: {Environment.NewLine}{text}{Environment.NewLine}{ResponseText}")
                Return ""
            Else

                text = ""

                Dim ImageFile As String = ImageDecoder.DecodeAndSaveImage(jsonObject)
                If Not String.IsNullOrWhiteSpace(ImageFile) Then
                    text = vbCrLf & "Image saved to: " & ImageFile & vbCrLf
                    text = text.Replace("\", "\\")
                End If

                If ResponseKey = "JSON" OrElse
                  (Not String.IsNullOrWhiteSpace(DetectToolCall) AndAlso
                   Regex.IsMatch(ResponseText, DetectToolCall, RegexOptions.Singleline Or RegexOptions.CultureInvariant)) Then

                    text = ResponseText
                Else
                    text = text & JsonTemplateFormatter.FormatJsonWithTemplate(jsonObject, ResponseKey, RKMode)
                    Dim hasLoop = Regex.IsMatch(ResponseKey, "\{\%\s*for\s+([^\s\%]+)\s*\%\}", RegexOptions.Singleline)
                    Dim hasPh = Regex.IsMatch(ResponseKey, "\{([^}]+)\}")
                    If Not hasLoop AndAlso Not hasPh Then text = text & ExtractCitations(jsonObject)
                End If

                Return text
            End If
        End Function

        ''' <summary>
        ''' Removes selected Unicode control/format/separator characters from <paramref name="text"/> while preserving spaces and CR/LF.
        ''' </summary>
        ''' <param name="text">Input text.</param>
        ''' <returns>Text with hidden marker categories removed.</returns>
        ''' <exception cref="System.Exception">Thrown when <paramref name="text"/> is <c>Nothing</c>.</exception>
        Public Shared Function RemoveHiddenMarkers(text As String) As String
            If text Is Nothing Then
                Throw New System.Exception("Cannot remove hidden markers from a null string.")
            End If

            Dim sb As New StringBuilder(text.Length)
            For Each ch As Char In text
                Dim uc As UnicodeCategory = Char.GetUnicodeCategory(ch)

                ' Allow ordinary space plus CR (U+000D) and LF (U+000A).
                If ch = " "c OrElse
                   ch = ChrW(13) OrElse      ' Carriage Return
                   ch = ChrW(10) OrElse      ' Line Feed
                   (uc <> UnicodeCategory.Control AndAlso
                    uc <> UnicodeCategory.Format AndAlso
                    uc <> UnicodeCategory.LineSeparator AndAlso
                    uc <> UnicodeCategory.ParagraphSeparator AndAlso
                    uc <> UnicodeCategory.SpaceSeparator) Then

                    sb.Append(ch)
                End If
            Next

            Return sb.ToString()
        End Function

        ''' <summary>
        ''' Extracts citations/references from known JSON response shapes and formats them into a human-readable list.
        ''' </summary>
        ''' <param name="jsonObj">JSON object to inspect.</param>
        ''' <returns>Formatted citations text; returns an empty string if no citations are found.</returns>
        Public Shared Function ExtractCitations(ByRef jsonObj As JObject) As String
            Try
                Dim OriginalJsonObj As JObject = CType(jsonObj.DeepClone(), JObject)
                Dim citationList As New List(Of String)
                Dim sourceUris As New HashSet(Of String)

                ' 1. Attempt extraction from candidates path (if present)
                Dim candidateCitations As JToken = jsonObj.SelectToken("candidates[0].content.parts[0].citations")
                If candidateCitations IsNot Nothing Then
                    If candidateCitations.Type = JTokenType.Array Then
                        For Each citation As JObject In candidateCitations
                            ProcessCitationObject(citation, citationList, sourceUris)
                        Next
                    ElseIf candidateCitations.Type = JTokenType.Object Then
                        ProcessCitationObject(CType(candidateCitations, JObject), citationList, sourceUris)
                    End If
                End If

                ' 2. Check for top-level citations (outside of candidates)
                Dim topLevelCitations As JToken = jsonObj.SelectToken("citations")
                If topLevelCitations IsNot Nothing Then
                    If topLevelCitations.Type = JTokenType.Array Then
                        For Each citation As JToken In topLevelCitations
                            If citation.Type = JTokenType.String Then
                                citationList.Add(citation.ToString())
                            ElseIf citation.Type = JTokenType.Object Then
                                ProcessCitationObject(CType(citation, JObject), citationList, sourceUris)
                            End If
                        Next
                    ElseIf topLevelCitations.Type = JTokenType.Object Then
                        ' Handle Format 2 (fullNote/shortNote) in a top-level object
                        Dim fullNote As String = topLevelCitations("fullNote")?.ToString()
                        If Not String.IsNullOrEmpty(fullNote) Then
                            citationList.Add(fullNote)
                        End If
                        Dim shortNote As String = topLevelCitations("shortNote")?.ToString()
                        If Not String.IsNullOrEmpty(shortNote) Then
                            citationList.Add(shortNote)
                        End If
                        ' In case no fullNote exists, fallback to checking for a URL
                        Dim url As String = topLevelCitations("url")?.ToString()
                        If Not String.IsNullOrEmpty(url) Then
                            citationList.Add(url)
                        End If
                    End If
                End If

                ' 3. Check citation metadata sources
                Dim metadataSources As JToken = jsonObj.SelectToken("citationMetadata.citationSources")
                If metadataSources IsNot Nothing AndAlso metadataSources.Type = JTokenType.Array Then
                    For Each source As JObject In metadataSources
                        ProcessMetadataSource(source, citationList, sourceUris)
                    Next
                End If

                ' 3b. Evaluate Google grounding supports.
                ProcessGroundingSupports(jsonObj, citationList, sourceUris)

                ' 4. Check legacy formats
                ExtractLegacyCitations(jsonObj, citationList, sourceUris)

                Debug.WriteLine("Total citations count: " & citationList.Count.ToString())

                ' 5. Build output: if any citation was found, format them; otherwise fall back.
                If citationList.Count > 0 Then
                    Debug.WriteLine("Citations: " & String.Join(", ", citationList))
                    Return FormatCitations(citationList)
                Else
                    Dim result As String = ExtractSimpleCitations(OriginalJsonObj)
                    Debug.WriteLine("Fallback Result = " & result)
                    Return result
                End If

            Catch ex As Exception
                Debug.WriteLine("Error parsing citations: " & ex.Message)
            End Try

            Return String.Empty
        End Function

        ''' <summary>
        ''' Interprets a citation object and appends extracted citation text/URLs to the output lists.
        ''' </summary>
        ''' <param name="citation">Citation object.</param>
        ''' <param name="citationList">List that receives formatted citation entries.</param>
        ''' <param name="sourceUris">Set used to deduplicate source URLs.</param>
        Private Shared Sub ProcessCitationObject(citation As JObject, ByRef citationList As List(Of String), ByRef sourceUris As HashSet(Of String))
            Try
                ' Format 1: Check for a "source" property (MLA/Chicago style)
                Dim source = citation.SelectToken("source")
                If source IsNot Nothing Then
                    AddSource(source, citationList, sourceUris)
                    ' Optionally include an inline citation if available
                    Dim inlineCitation = citation("inlineCitation")?.ToString()
                    If Not String.IsNullOrEmpty(inlineCitation) Then
                        citationList.Add("Inline: " & inlineCitation)
                    End If
                    Return
                End If

                ' Format 2: Check for a "fullNote" property (full note/short note format)
                Dim fullNote As String = citation("fullNote")?.ToString()
                If Not String.IsNullOrEmpty(fullNote) Then
                    citationList.Add(fullNote)
                    Return
                End If

                ' Format 3: IEEE style with "referenceEntry"
                Dim refEntry As String = citation("referenceEntry")?.ToString()
                If Not String.IsNullOrEmpty(refEntry) Then
                    Dim ieeeUri As String = ExtractIeeeUri(refEntry)
                    If Not String.IsNullOrEmpty(ieeeUri) AndAlso sourceUris.Add(ieeeUri) Then
                        citationList.Add($"{refEntry} | Source: {ieeeUri}")
                    Else
                        citationList.Add(refEntry)
                    End If
                    Return
                End If

                ' Format 4: Harvard style with "referenceList.entry" and optionally "textualCitation"
                Dim refListToken As JToken = citation.SelectToken("referenceList.entry")
                If refListToken IsNot Nothing Then
                    Dim refList As String = refListToken.ToString()
                    Dim harvardUri As String = ExtractHarvardUri(refList)
                    Dim textualCitation As String = citation("textualCitation")?.ToString()
                    Dim formattedCitation As String = (If(Not String.IsNullOrEmpty(textualCitation), textualCitation, "") & " | " & refList).Trim(" "c, "|"c)
                    citationList.Add(formattedCitation)
                    Return
                End If

                ' Fallback: If the citation object has a "url" property directly, extract it.
                Dim url As String = citation("url")?.ToString()
                If Not String.IsNullOrEmpty(url) AndAlso sourceUris.Add(url) Then
                    citationList.Add(url)
                End If

            Catch ex As Exception
                Debug.WriteLine("Error processing citation object: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Extracts citation metadata from a <c>citationSources</c> entry and appends a formatted line to <paramref name="citationList"/>.
        ''' </summary>
        ''' <param name="source">Citation source JSON object.</param>
        ''' <param name="citationList">List that receives formatted citation entries.</param>
        ''' <param name="sourceUris">Set used to deduplicate source URLs.</param>
        Private Shared Sub ProcessMetadataSource(source As JObject, ByRef citationList As List(Of String), ByRef sourceUris As HashSet(Of String))
            Try
                Dim uri As String = source("uri")?.ToString()
                If Not String.IsNullOrEmpty(uri) AndAlso sourceUris.Add(uri) Then
                    Dim title As String = source("title")?.ToString()
                    If String.IsNullOrWhiteSpace(title) Then title = "No title"
                    Dim authors As String = String.Join(", ", source.SelectTokens("authors[*].given").Select(Function(t) t.ToString()))
                    Dim doi As String = source("doi")?.ToString()
                    citationList.Add($"Source: {title} | Authors: {If(authors, "Unknown")} | DOI: {If(doi, "N/A")} | URL: {uri}")
                End If
            Catch ex As Exception
                Debug.WriteLine("Error processing metadata source: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Extracts and formats citation details from a <c>source</c> token and appends it to <paramref name="citationList"/>.
        ''' </summary>
        ''' <param name="source">Source token containing citation details.</param>
        ''' <param name="citationList">List that receives formatted citation entries.</param>
        ''' <param name="sourceUris">Set used to deduplicate source URLs.</param>
        Private Shared Sub AddSource(source As JToken, ByRef citationList As List(Of String), ByRef sourceUris As HashSet(Of String))
            Try
                Dim uri As String = source("uri")?.ToString()
                If String.IsNullOrEmpty(uri) OrElse sourceUris.Contains(uri) Then Return

                Dim sb As New StringBuilder()
                sb.Append("Source: ")

                ' Build title with container if available
                Dim title As String = source("title")?.ToString()
                Dim container As String = source("containerTitle")?.ToString()
                If Not String.IsNullOrEmpty(container) Then
                    sb.Append($"{title}. In: {container}")
                Else
                    sb.Append(title)
                End If

                ' Add authors
                Dim authors = source.SelectTokens("authors[*]")
                If authors IsNot Nothing AndAlso authors.Any() Then
                    sb.Append(" | Authors: ")
                    For Each author In authors
                        Dim given As String = author("given")?.ToString()
                        Dim family As String = author("family")?.ToString()
                        If Not String.IsNullOrEmpty(family) Then
                            sb.Append($"{family}, {given}; ")
                        End If
                    Next
                    If sb.Length > 2 Then
                        sb.Length -= 2 ' Remove last semicolon and space
                    End If
                End If

                ' Add publication info
                Dim pubDate As String = source("publicationDate")?.ToString()
                If Not String.IsNullOrEmpty(pubDate) Then
                    sb.Append($" | Published: {pubDate}")
                End If

                ' Add DOI if available
                Dim doi As String = source("doi")?.ToString()
                If Not String.IsNullOrEmpty(doi) Then
                    sb.Append($" | DOI: {doi}")
                End If
                sb.Append($" | URL: {uri}")

                citationList.Add(sb.ToString())
                sourceUris.Add(uri)
            Catch ex As Exception
                Debug.WriteLine("Error adding source: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Extracts legacy citation URLs from older JSON response shapes.
        ''' </summary>
        ''' <param name="jsonObj">JSON object to inspect.</param>
        ''' <param name="citationList">List that receives formatted citation entries.</param>
        ''' <param name="sourceUris">Set used to deduplicate source URLs.</param>
        Private Shared Sub ExtractLegacyCitations(jsonObj As JObject, ByRef citationList As List(Of String), ByRef sourceUris As HashSet(Of String))
            Try
                ' Old format v0.9 compatibility: look for any "sources" with a URL.
                Dim legacyCitations = jsonObj.SelectTokens("$..sources[?(@.url)]")
                For Each legacySource In legacyCitations
                    Dim url As String = legacySource("url")?.ToString()
                    If Not String.IsNullOrEmpty(url) AndAlso sourceUris.Add(url) Then
                        citationList.Add($"Legacy source: {url}")
                    End If
                Next
            Catch ex As Exception
                Debug.WriteLine("Error processing legacy citations: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Extracts a DOI from an IEEE reference entry and returns a DOI URL when present.
        ''' </summary>
        ''' <param name="refEntry">Reference entry string.</param>
        ''' <returns>DOI URL when detected; otherwise an empty string.</returns>
        Private Shared Function ExtractIeeeUri(refEntry As String) As String
            Try
                Dim doiMatch = Regex.Match(refEntry, "doi:\s*(\S+)")
                If doiMatch.Success Then
                    ' Trim any trailing punctuation
                    Return $"https://doi.org/{doiMatch.Groups(1).Value.TrimEnd("."c)}"
                End If
            Catch ex As Exception
                Debug.WriteLine("DOI extraction error: " & ex.Message)
            End Try
            Return String.Empty
        End Function

        ''' <summary>
        ''' Extracts a URL from a Harvard-style reference entry when it contains an "Available at:" segment.
        ''' </summary>
        ''' <param name="refEntry">Reference entry string.</param>
        ''' <returns>Extracted URL when detected; otherwise an empty string.</returns>
        Private Shared Function ExtractHarvardUri(refEntry As String) As String
            Try
                Dim uriMatch = Regex.Match(refEntry, "Available at:\s*(\S+)\s*\(")
                If uriMatch.Success Then
                    Return uriMatch.Groups(1).Value
                End If
            Catch ex As Exception
                Debug.WriteLine("Harvard URI extraction error: " & ex.Message)
            End Try
            Return String.Empty
        End Function

        ''' <summary>
        ''' Formats a list of citation strings into a numbered "References:" section and wraps URLs as Markdown links.
        ''' </summary>
        ''' <param name="citationList">List of citations to format.</param>
        ''' <returns>Formatted citations section.</returns>
        Private Shared Function FormatCitations(citationList As List(Of String)) As String
            Dim sb As New StringBuilder()
            sb.AppendLine()
            sb.AppendLine()
            sb.AppendLine(vbCrLf & vbCrLf & "References:")
            For i As Integer = 0 To citationList.Count - 1
                Dim text As String = System.Text.RegularExpressions.Regex.Replace(
                                        citationList(i),
                                        "(?<!\]\()https?://\S+",
                                        Function(m) $"[{m.Value}]({m.Value})")
                sb.AppendLine($"[{i + 1}] {text}")
            Next
            Return sb.ToString()
        End Function

        ''' <summary>
        ''' Extracts citations from the simple <c>citations</c> JSON property and formats them as a numbered list.
        ''' </summary>
        ''' <param name="jsonObj">JSON object passed by reference.</param>
        ''' <returns>Formatted citation output; empty string on parse errors or if no citations are present.</returns>
        Private Shared Function ExtractSimpleCitations(ByRef jsonObj As JObject) As String
            Try
                Dim citations As JToken = jsonObj.SelectToken("citations")
                Dim citationList As New List(Of String)

                If citations IsNot Nothing Then
                    If citations.Type = JTokenType.Array Then
                        For Each citation As JToken In citations
                            If citation.Type = JTokenType.String Then
                                citationList.Add(citation.ToString())
                            ElseIf citation.Type = JTokenType.Object Then
                                Dim url As JToken = citation.SelectToken("url")
                                If url IsNot Nothing Then
                                    citationList.Add(url.ToString())
                                Else
                                    Dim fullNote As String = citation("fullNote")?.ToString()
                                    If Not String.IsNullOrEmpty(fullNote) Then
                                        citationList.Add(fullNote)
                                    End If
                                End If
                            End If
                        Next
                    ElseIf citations.Type = JTokenType.Object Then
                        Dim fullNote As String = citations("fullNote")?.ToString()
                        If Not String.IsNullOrEmpty(fullNote) Then
                            citationList.Add(fullNote)
                        Else
                            Dim url As String = citations("url")?.ToString()
                            If Not String.IsNullOrEmpty(url) Then
                                citationList.Add(url)
                            End If
                        End If
                    End If
                End If

                Dim simpleCitationOutput As New StringBuilder()
                simpleCitationOutput.AppendLine(vbCrLf)
                For i As Integer = 0 To citationList.Count - 1
                    simpleCitationOutput.AppendLine("[" & (i + 1).ToString() & "] " & citationList(i))
                Next

                Return simpleCitationOutput.ToString()

            Catch ex As Exception
                Debug.WriteLine("Error parsing JSON for simple citations: " & ex.Message)
            End Try

            Return String.Empty
        End Function

        ''' <summary>
        ''' Extracts and formats citations from Google grounding metadata (<c>groundingSupports</c>/<c>groundingChunks</c>).
        ''' </summary>
        ''' <param name="jsonObj">JSON response object.</param>
        ''' <param name="citationList">List that receives formatted citation entries.</param>
        ''' <param name="sourceUris">Set used to deduplicate source URLs.</param>
        Private Shared Sub ProcessGroundingSupports(jsonObj As Newtonsoft.Json.Linq.JObject,
                                            ByRef citationList As System.Collections.Generic.List(Of String),
                                            ByRef sourceUris As System.Collections.Generic.HashSet(Of String))

            Try
                ' Check expected paths.
                Dim supports As Newtonsoft.Json.Linq.JToken =
            jsonObj.SelectToken("candidates[0].groundingMetadata.groundingSupports")
                Dim chunks As Newtonsoft.Json.Linq.JToken =
            jsonObj.SelectToken("candidates[0].groundingMetadata.groundingChunks")

                If supports Is Nothing OrElse chunks Is Nothing _
           OrElse supports.Type <> Newtonsoft.Json.Linq.JTokenType.Array _
           OrElse chunks.Type <> Newtonsoft.Json.Linq.JTokenType.Array Then Exit Sub

                ' Process each support segment.
                For Each support As Newtonsoft.Json.Linq.JObject In supports

                    Dim segText As String = support.SelectToken("segment.text")?.ToString()
                    Dim idxTokens As Newtonsoft.Json.Linq.JToken =
                support.SelectToken("groundingChunkIndices")

                    If System.String.IsNullOrWhiteSpace(segText) _
               OrElse idxTokens Is Nothing _
               OrElse idxTokens.Type <> Newtonsoft.Json.Linq.JTokenType.Array Then Continue For

                    segText = RemoveMarkdownFormatting(segText)
                    Dim sb As New System.Text.StringBuilder()
                    sb.Append("... " &
                      segText.Replace(vbCrLf, " ") _
                             .Replace(vbCr, " ") _
                             .Replace(vbLf, " ") _
                             .Trim() &
                      " ...")

                    ' Segment-level deduplication.
                    Dim localUris As New System.Collections.Generic.HashSet(Of String)(
                                    System.StringComparer.OrdinalIgnoreCase)

                    For Each idxTok In idxTokens
                        Dim idx As Integer
                        If Not Integer.TryParse(idxTok.ToString(), idx) Then Continue For
                        If idx < 0 OrElse idx >= chunks.Count Then Continue For

                        Dim webObj As Newtonsoft.Json.Linq.JObject = TryCast(chunks(idx).SelectToken("web"), JObject)
                        If webObj Is Nothing Then Continue For

                        Dim uri As String = webObj("uri")?.ToString()
                        Dim title As String = webObj("title")?.ToString()

                        If System.String.IsNullOrWhiteSpace(uri) _
                   OrElse Not localUris.Add(uri) Then Continue For

                        If System.String.IsNullOrWhiteSpace(title) Then title = "No title"

                        sb.Append(" ([" & title & "](" & uri & "))")
                    Next

                    citationList.Add(sb.ToString())
                Next

            Catch ex As System.Exception
                System.Diagnostics.Debug.WriteLine("Error processing groundingSupports: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Returns a standards-compliant MIME type for a given (possibly legacy or vendor-prefixed) MIME type.
        ''' </summary>
        ''' <param name="legacyType">The input MIME type (may be legacy, nonstandard, or correct).</param>
        ''' <returns>The corrected, standard MIME type (or original if not matched).</returns>
        Public Shared Function FixMimeType(legacyType As String) As String
            If String.IsNullOrWhiteSpace(legacyType) Then Return "application/octet-stream"
            Select Case legacyType.Trim.ToLowerInvariant()
                ' --- Images ---
                Case "image/x-png", "image/x-citrix-png" : Return "image/png"
                Case "image/x-jpeg", "image/pjpeg", "image/pjepg", "image/x-pjpeg", "image/x-citrix-jpeg" : Return "image/jpeg"
                Case "image/jpg" : Return "image/jpeg"
                Case "image/x-bmp", "image/x-ms-bmp" : Return "image/bmp"
                Case "image/x-tiff" : Return "image/tiff"
                Case "image/x-emf" : Return "image/emf"
                Case "image/x-wmf" : Return "image/wmf"
                Case "image/x-icon" : Return "image/vnd.microsoft.icon"
                Case "image/ico" : Return "image/vnd.microsoft.icon"
                Case "image/svg" : Return "image/svg+xml"
                Case "image/x-svg" : Return "image/svg+xml"
                ' --- Audio/Video ---
                Case "audio/x-wav" : Return "audio/wav"
                Case "audio/x-mp3", "audio/mpeg3" : Return "audio/mpeg"
                Case "audio/x-midi", "audio/midi" : Return "audio/midi"
                Case "video/x-msvideo" : Return "video/x-msvideo"
                ' --- Documents ---
                Case "application/x-pdf", "application/pdfx" : Return "application/pdf"
                Case "application/x-rtf" : Return "application/rtf"
                Case "application/x-msword" : Return "application/msword"
                Case "application/x-msexcel" : Return "application/vnd.ms-excel"
                Case "application/x-mspowerpoint" : Return "application/vnd.ms-powerpoint"
                Case "application/vnd.ms-word.document.macroenabled.12" : Return "application/msword"
                Case "application/vnd.ms-excel.sheet.macroenabled.12" : Return "application/vnd.ms-excel"
                Case "application/vnd.ms-powerpoint.presentation.macroenabled.12" : Return "application/vnd.ms-powerpoint"
                ' --- Office Open XML ---
                Case "application/x-docx" : Return "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                Case "application/x-xlsx" : Return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                Case "application/x-pptx" : Return "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                ' --- Archives/Compression ---
                Case "application/x-zip-compressed", "application/x-zip" : Return "application/zip"
                Case "application/x-gzip" : Return "application/gzip"
                Case "application/x-tar" : Return "application/x-tar"
                Case "application/x-7z-compressed" : Return "application/x-7z-compressed"
                ' --- Text/CSV ---
                Case "text/x-csv" : Return "text/csv"
                Case "text/x-log" : Return "text/plain"
                Case "text/x-ini" : Return "text/plain"
                ' --- Misc ---
                Case "application/x-shockwave-flash" : Return "application/vnd.adobe.flash.movie"
                Case "application/x-msdownload" : Return "application/octet-stream"
                Case "application/x-bittorrent" : Return "application/x-bittorrent"
                Case "application/x-iso9660-image" : Return "application/x-iso9660-image"
                ' --- Defaults and unknowns ---
                Case "" : Return "application/octet-stream"
                Case Else
                    ' Special handling for some popular typos and aliases:
                    If legacyType.ToLowerInvariant() = "image/jpg" Then Return "image/jpeg"
                    If legacyType.ToLowerInvariant() = "image/tif" Then Return "image/tiff"
                    Return legacyType
            End Select
        End Function

        ''' <summary>
        ''' Recursively searches a JSON token for a property name and returns its first string value.
        ''' </summary>
        ''' <param name="token">Token to search.</param>
        ''' <param name="searchtext">Property name to match.</param>
        ''' <returns>Property value as string when found; otherwise <c>Nothing</c>.</returns>
        Public Shared Function FindJsonProperty(token As JToken, searchtext As String) As String
            If token.Type = JTokenType.Object Then
                For Each prop As JProperty In CType(token, JObject).Properties()
                    If prop.Name = searchtext Then
                        Return prop.Value.ToString()
                    End If
                    Dim result As String = FindJsonProperty(prop.Value, searchtext)
                    If Not String.IsNullOrEmpty(result) Then Return result
                Next
            ElseIf token.Type = JTokenType.Array Then
                For Each item As JToken In CType(token, JArray)
                    Dim result As String = FindJsonProperty(item, searchtext)
                    If Not String.IsNullOrEmpty(result) Then Return result
                Next
            End If
            Return Nothing
        End Function

        ''' <summary>
        ''' Ensures an OAuth2 access token is available and not expired, refreshing it when needed via <c>GoogleOAuthHelper</c>.
        ''' </summary>
        ''' <param name="context">Shared configuration context where decoded API keys and expiry timestamps are stored.</param>
        ''' <param name="clientEmail">OAuth2 client email passed to <c>GoogleOAuthHelper</c>.</param>
        ''' <param name="ClientScopes">OAuth2 scopes passed to <c>GoogleOAuthHelper</c>.</param>
        ''' <param name="PrivateKey">Private key string used to construct a PEM key for <c>GoogleOAuthHelper</c>.</param>
        ''' <param name="AuthServer">Token endpoint URI passed to <c>GoogleOAuthHelper</c>.</param>
        ''' <param name="TLife">Lifetime in seconds used to compute expiry timestamps in <paramref name="context"/>.</param>
        ''' <param name="SecondAPI">If <c>True</c>, updates the secondary token fields in <paramref name="context"/>.</param>
        ''' <returns>Access token string; returns an empty string on errors.</returns>
        Public Shared Async Function GetFreshAccessToken(context As ISharedContext, ByVal clientEmail As String, ByVal ClientScopes As String, ByVal PrivateKey As String, ByVal AuthServer As String, ByVal TLife As Long, ByVal SecondAPI As Boolean) As Task(Of String)
            Try

                Dim accessToken As String = String.Empty
                Dim currentexpiry As DateTime
                If SecondAPI Then
                    accessToken = context.DecodedAPI_2
                    currentexpiry = context.TokenExpiry_2
                Else
                    accessToken = context.DecodedAPI
                    currentexpiry = context.TokenExpiry
                End If

                PrivateKey = PrivateKey.Replace("\n", "")

                Dim formattedKey As String = String.Empty

                For i As Integer = 0 To PrivateKey.Length - 1 Step 64
                    If i + 64 <= PrivateKey.Length Then
                        formattedKey &= PrivateKey.Substring(i, 64) & vbLf
                    Else
                        formattedKey &= PrivateKey.Substring(i) & vbLf
                    End If
                Next

                GoogleOAuthHelper.client_email = clientEmail
                GoogleOAuthHelper.private_key = "-----BEGIN PRIVATE KEY-----" & vbLf & formattedKey & "-----END PRIVATE KEY-----" & vbLf
                GoogleOAuthHelper.scopes = ClientScopes
                GoogleOAuthHelper.token_uri = AuthServer
                GoogleOAuthHelper.token_life = TLife

                If String.IsNullOrEmpty(accessToken) OrElse DateTime.UtcNow >= currentexpiry Then
                    ' Token is missing or expired, fetch a new one.
                    accessToken = Await GoogleOAuthHelper.GetAccessToken()
                    If SecondAPI Then
                        context.TokenExpiry_2 = DateTime.UtcNow.AddSeconds(GoogleOAuthHelper.token_life - 300) ' Set expiry 5 minutes before actual
                        context.DecodedAPI_2 = accessToken
                    Else
                        context.TokenExpiry = DateTime.UtcNow.AddSeconds(GoogleOAuthHelper.token_life - 300) ' Set expiry 5 minutes before actual
                        context.DecodedAPI = accessToken
                    End If
                End If
                Return accessToken

            Catch ex As System.Exception
                ' Handle exceptions explicitly with System.Exception
                MessageBox.Show("Error while fetching an access token: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                If SecondAPI Then
                    context.DecodedAPI_2 = String.Empty
                Else
                    context.DecodedAPI = String.Empty
                End If
                Return String.Empty
            End Try

        End Function

        ''' <summary>
        ''' Escapes a string for safe injection into JSON or URL templates by encoding control characters and quotes/backslashes.
        ''' </summary>
        ''' <param name="input">Input string to escape.</param>
        ''' <param name="collapseSpaces">If <c>True</c>, collapses repeated spaces within each line while preserving indentation.</param>
        ''' <returns>Escaped string, or an empty string when <paramref name="input"/> is null/whitespace.</returns>
        Public Shared Function CleanString(ByVal input As String, Optional ByVal collapseSpaces As Boolean = True) As String
            ' If empty or whitespace only, return an empty string.
            If System.String.IsNullOrWhiteSpace(input) Then
                Return ""
            End If

            ' 1) First pass: escape into sbEscaped.
            Dim sbEscaped As New System.Text.StringBuilder(input.Length * 2)
            For Each c As Char In input
                Select Case AscW(c)
                    Case 8      ' backspace
                        sbEscaped.Append("\b")
                    Case 9      ' tab
                        sbEscaped.Append("\t")
                    Case 10     ' line feed
                        sbEscaped.Append("\n")
                    Case 12     ' form feed
                        sbEscaped.Append("\f")
                    Case 13     ' carriage return → normalized to "\n"
                        sbEscaped.Append("\n")
                    Case 34     ' double-quote → must become "\""
                        sbEscaped.Append("\""")
                    Case 92     ' backslash → "\\"
                        sbEscaped.Append("\\")
                    Case 0 To 31 ' other control codes → "\uXXXX"
                        sbEscaped.Append("\u" & AscW(c).ToString("X4"))
                    Case Else
                        sbEscaped.Append(c)
                End Select
            Next

            ' 2) Second pass: collapse spaces only when collapseSpaces = True.
            If collapseSpaces Then
                Dim raw As String = sbEscaped.ToString()
                Dim lines As String() = raw.Split(New String() {"\n"}, System.StringSplitOptions.None)
                Dim sbResult As New System.Text.StringBuilder(raw.Length)

                For i As Integer = 0 To lines.Length - 1
                    Dim line As String = lines(i)
                    ' Preserve leading spaces.
                    Dim indentLen As Integer = 0
                    While indentLen < line.Length AndAlso line(indentLen) = " "c
                        indentLen += 1
                    End While
                    Dim prefix As String = line.Substring(0, indentLen)
                    Dim rest As String = line.Substring(indentLen)

                    ' Collapse multiple spaces in the remainder.
                    Dim sbLine As New System.Text.StringBuilder(rest.Length)
                    Dim lastWasSpaceInner As Boolean = False
                    For Each c2 As Char In rest
                        If c2 = " "c Then
                            If Not lastWasSpaceInner Then
                                sbLine.Append(" "c)
                                lastWasSpaceInner = True
                            End If
                        Else
                            sbLine.Append(c2)
                            lastWasSpaceInner = False
                        End If
                    Next

                    sbResult.Append(prefix).Append(sbLine.ToString())
                    If i < lines.Length - 1 Then sbResult.Append("\n")
                Next

                Return sbResult.ToString()
            End If

            ' If collapseSpaces = False, return the escaped string.
            Return sbEscaped.ToString()
        End Function

        ''' <summary>
        ''' Estimates token usage by dividing character count by 4 and rounding up.
        ''' </summary>
        ''' <param name="text">Input text.</param>
        ''' <returns>Estimated token count.</returns>
        Public Shared Function EstimateTokenCount(text As String) As Integer
            ' Trim the text and handle edge cases
            If String.IsNullOrWhiteSpace(text) Then Return 0

            ' Estimate tokens: Average of 4 characters per token for English Language
            Dim charCount As Integer = text.Length
            Dim estimatedTokens As Integer = CInt(Math.Ceiling(charCount / 4.0))

            Return estimatedTokens
        End Function

    End Class
End Namespace