' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.Tooling.ToolExecution.vb
' Purpose: External tool execution, transport handling, and tool service error extraction.
'          Routes tool calls through HTTP/SSE/MCP transports or standard LLM() bridge.
'
' Architecture:
'  - Transport Modes:
'      - SSE Transport: Bypasses LLM() for "sse:" prefixed endpoints.
'          - Uses ExecuteMCPSSEToolCall() for MCP Server-Sent Events protocol.
'          - Supports OAuth2 token refresh on 401 Unauthorized responses.
'      - MCP Streamable HTTP: Bypasses LLM() for "mcp_streamable:" endpoints.
'          - Uses ExecuteMCPStreamableToolCall() for direct HTTP streaming.
'      - Standard Transport: Routes through LLM() with APICall injection.
'          - Applies toolConfig to _context, forces JSON response mode.
'          - Restores original configuration via backupConfig.
'  - Tool Configuration & Preparation:
'      - Resolves ToolAPICall or APICall template from ModelConfig.
'      - Replaces {param} placeholders with actual tool call arguments.
'      - Applies type-specific formatting via FormatToolValueForPlaceholder().
'      - Handles unreplaced placeholders via ToolParameterDefaults fallback.
'  - OAuth2 & Authentication:
'      - Supports OAuth2 flow via GetFreshAccessToken() when toolConfig.OAuth2=true.
'      - Replaces {apikey} in HeaderB with DecodedAPI token.
'      - Forces token refresh on Unauthorized responses via ForceRefreshToolOAuthToken().
'  - Error Extraction:
'      - TryExtractToolServiceErrorMessage(): Parses structured error responses.
'      - Detects {"error": {...}} or {"result": {"isError": true}} patterns.
'  - Tool Call Detection & Extraction:
'      - ContainsToolCalls(): Regex-based detection using ToolCallDetectionPattern.
'      - ExtractToolCalls(): JSON-path-based extraction using ToolCallExtractionMap.
'      - ExtractToolCallPatternFromResponse(): Extracts embedded regex from INI_Response_2.
'  - Utility Helpers:
'      - IsValidJson(): Validates JSON object/array structure.
'      - EscapeJsonString(): Escapes control characters for JSON string literals.
'      - GetToolParameterSchemas(): Extracts parameter schema from ToolDefinition.
'      - GetToolRequiredParameters(): Extracts required parameter names.
'      - GetToolParameterType(): Resolves parameter type from schema.
'      - GetToolParameterEnumValues(): Extracts enum constraints.
'
' Key Functions:
'  - ExecuteExternalTool(): Primary external tool execution entry point.
'  - TryExtractToolServiceErrorMessage(): Error response parser.
'  - ContainsToolCalls(): Tool call detection.
'  - ExtractToolCalls(): Tool call extraction.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.Reflection
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Microsoft.Web.WebView2.Core
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedContext
Imports SharedLibrary.SharedLibrary.SharedMethods

''' <summary>
''' Provides tooling support helpers for model-agnostic tool/function calling in LLM interactions.
''' </summary>
Partial Public Class ThisAddIn


    ''' <summary>
    ''' Executes an external tool by applying its <see cref="ModelConfig"/> to <c>_context</c>, preparing
    ''' the tool API call payload, and invoking <c>LLM</c> in JSON response mode.
    ''' For SSE/MCP endpoints (prefixed with "sse:"), bypasses LLM() and calls the MCP server directly.
    ''' </summary>
    ''' <param name="toolCall">Tool call extracted from the LLM response.</param>
    ''' <param name="toolConfig">Tool configuration to apply for this call.</param>
    ''' <param name="context">Tool execution context for logging and diagnostics.</param>
    ''' <param name="cancellationToken">Optional cancellation token to abort execution.</param>
    ''' <returns>Tool response containing the tool service result or an error.</returns>
    Private Async Function ExecuteExternalTool(toolCall As ToolCall, toolConfig As ModelConfig, context As ToolExecutionContext, Optional cancellationToken As System.Threading.CancellationToken = Nothing) As Task(Of ToolResponse)
        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId,
            .ToolName = toolCall.ToolName
        }

        Try
            cancellationToken.ThrowIfCancellationRequested()

            Dim apiCallTemplate = toolConfig.ToolAPICall
            If String.IsNullOrWhiteSpace(apiCallTemplate) Then
                apiCallTemplate = toolConfig.APICall
            End If

            If String.IsNullOrWhiteSpace(apiCallTemplate) Then
                ' The tool is advertised but no transport (HTTP/MCP/SSE) is wired for it
                ' on this host. We return a STRUCTURED error payload (resultKind="error",
                ' error.code="tool_not_executable") so the orchestrator's deliverable-
                ' fallback gate can force the model to try a fallback tool instead of
                ' accepting "blocked" on the first hit.
                response.Success = False
                response.ResultKind = "error"
                response.ErrorCode = "tool_not_executable"
                response.ErrorMessage =
                "The tool '" & If(toolCall.ToolName, "") &
                "' is advertised but not executable on this host (no transport template). " &
                "Choose a different tool. If the user authorized a fallback, attempt it before declaring blocked."
                response.Response = Agents.ToolExecutorRegistry.BuildNotExecutablePayload(
                toolCall.ToolName,
                If(context IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(context.HostKind),
                   context.HostKind,
                   "Outlook"))
                ToolingFileLogger.LogError(
                "Tool advertised but not executable on this host.",
                details:=$"ToolName='{toolCall.ToolName}'")
                Return response
            End If

            Dim apiCall = apiCallTemplate
            Dim parameterSchemas = GetToolParameterSchemas(toolConfig)
            Dim requiredParameters = GetToolRequiredParameters(toolConfig)

            For Each kvp In toolCall.Arguments
                Dim placeholder = "{" & kvp.Key & "}"
                Dim schemaToken As JToken = Nothing
                parameterSchemas.TryGetValue(kvp.Key, schemaToken)

                Dim value As String
                If kvp.Value Is Nothing Then
                    value = ""
                ElseIf TypeOf kvp.Value Is JToken Then
                    Dim jt = DirectCast(kvp.Value, JToken)
                    If jt.Type = JTokenType.String Then
                        value = FormatToolValueForPlaceholder(jt.Value(Of String)(), schemaToken)
                    Else
                        value = jt.ToString(Formatting.None)
                    End If
                Else
                    value = FormatToolValueForPlaceholder(kvp.Value.ToString(), schemaToken)
                End If

                apiCall = apiCall.Replace(placeholder, value)
            Next

            Dim unreplacedPattern As New Regex("\{([a-zA-Z_][a-zA-Z0-9_]*)\}")
            Dim unreplacedMatches = unreplacedPattern.Matches(apiCall)

            If unreplacedMatches.Count > 0 Then
                ToolingFileLogger.LogWarn(
                    "Unreplaced placeholders found in tool APICall (defaults will be applied if available).",
                    details:=$"ToolName='{toolCall.ToolName}'; Count={unreplacedMatches.Count}; APICall='{apiCall}'")

                Dim toolDefaults As Dictionary(Of String, String) = Nothing
                If Not String.IsNullOrWhiteSpace(toolConfig.ToolParameterDefaults) Then
                    Try
                        toolDefaults = JsonConvert.DeserializeObject(Of Dictionary(Of String, String))(toolConfig.ToolParameterDefaults)
                    Catch ex As Exception
                        ToolingFileLogger.LogWarn(
                            "ToolParameterDefaults parse failed.",
                            details:=$"ToolName='{toolCall.ToolName}'; ToolParameterDefaults='{toolConfig.ToolParameterDefaults}'",
                            ex:=ex)
                    End Try
                End If

                For Each m As Match In unreplacedMatches
                    Dim placeholderName = m.Groups(1).Value
                    Dim schemaToken As JToken = Nothing
                    parameterSchemas.TryGetValue(placeholderName, schemaToken)

                    Dim shouldRemoveProperty As Boolean = False
                    Dim replacement As String = ResolveToolDefaultValue(
                        placeholderName,
                        toolDefaults,
                        schemaToken,
                        requiredParameters.Contains(placeholderName),
                        shouldRemoveProperty)

                    If shouldRemoveProperty Then
                        apiCall = RemoveToolArgumentPlaceholderProperty(apiCall, placeholderName)
                    Else
                        apiCall = apiCall.Replace(m.Value, replacement)
                    End If
                Next

                Dim remainingMatches = unreplacedPattern.Matches(apiCall)
                If remainingMatches.Count > 0 Then
                    Dim remainingNames = remainingMatches.
                        Cast(Of Match)().
                        Select(Function(m) m.Groups(1).Value).
                        Distinct(StringComparer.OrdinalIgnoreCase).
                        ToList()

                    response.Success = False
                    response.ErrorMessage = $"Unreplaced placeholders remain after applying defaults: {String.Join(", ", remainingNames)}"

                    ToolingFileLogger.LogError(
                        "Unreplaced placeholders remain in tool APICall after applying defaults.",
                        details:=$"ToolName='{toolCall.ToolName}'; RemainingCount={remainingMatches.Count}; Remaining='{String.Join(", ", remainingNames)}'; APICall='{apiCall}'")

                    Return response
                End If
            End If

            If toolConfig.OAuth2 Then
                toolConfig.DecodedAPI = Await SharedMethods.GetFreshAccessToken(
                    _context,
                    toolConfig.OAuth2ClientMail,
                    toolConfig.OAuth2Scopes,
                    toolConfig.APIKey,
                    toolConfig.OAuth2Endpoint,
                    toolConfig.OAuth2ATExpiry,
                    True,
                    False).ConfigureAwait(False)

                If String.IsNullOrWhiteSpace(toolConfig.DecodedAPI) Then
                    response.Success = False
                    response.ErrorMessage = "OAuth2 authentication failed."
                    ToolingFileLogger.LogError(
                        "OAuth2 authentication failed before MCP tool execution.",
                        details:=$"ToolName='{toolCall.ToolName}'")
                    Return response
                End If
            End If

            ' ── SSE transport: full round-trip bypassing LLM() ───────────
            If Not String.IsNullOrWhiteSpace(toolConfig.Endpoint) AndAlso
               toolConfig.Endpoint.StartsWith(SharedMethods.MCP_SSE_PREFIX, StringComparison.OrdinalIgnoreCase) Then

                Dim sseBase = toolConfig.Endpoint.Substring(SharedMethods.MCP_SSE_PREFIX.Length)
                Dim resolvedHeaderB = If(toolConfig.HeaderB, "").Replace("{apikey}", If(toolConfig.DecodedAPI, ""))

                context.Log($"SSE transport: executing tool {toolCall.ToolName} via {sseBase}")
                ToolingFileLogger.LogStep($"SSE round-trip for {toolCall.ToolName} at {sseBase}")
                ToolingFileLogger.LogStep($"SSE request body: {apiCall}")

                Dim sseAttemptedRefresh As Boolean = False
                Dim sseEx As Exception = Nothing
                Dim sseDone As Boolean = False
                Dim sseCancelled As Boolean = False

                Do
                    sseEx = Nothing

                    Try
                        cancellationToken.ThrowIfCancellationRequested()

                        resolvedHeaderB = If(toolConfig.HeaderB, "").Replace("{apikey}", If(toolConfig.DecodedAPI, ""))

                        Dim rawResult = Await SharedMethods.ExecuteMCPSSEToolCall(
                            _context,
                            sseBase, apiCall,
                            If(toolConfig.HeaderA, ""), resolvedHeaderB,
                            CInt(Math.Min(If(toolConfig.Timeout > 0, toolConfig.Timeout, 60000L), Integer.MaxValue)))

                        ToolingFileLogger.LogRawResponseStub($"SSE tool result ({toolCall.ToolName})", rawResult)

                        If Not String.IsNullOrWhiteSpace(rawResult) Then
                            Dim toolErrorMessage As String = ""
                            response.Response = rawResult

                            If TryExtractToolServiceErrorMessage(rawResult, toolErrorMessage) Then
                                response.Success = False
                                response.ErrorMessage = toolErrorMessage
                                ToolingFileLogger.LogWarn(
                                    "SSE tool service returned a logical error.",
                                    details:=$"ToolName='{toolCall.ToolName}'; Error='{toolErrorMessage}'")
                            Else
                                response.Success = True
                            End If
                        Else
                            response.Success = False
                            response.ErrorMessage = "Empty response from SSE tool service"
                            ToolingFileLogger.LogError("Empty SSE response.", details:=$"ToolName='{toolCall.ToolName}'")
                        End If

                        sseDone = True

                    Catch ex As OperationCanceledException
                        sseCancelled = True
                        response.Success = False
                        response.ErrorMessage = "Operation was cancelled"
                        ToolingFileLogger.LogWarn($"Tool {toolCall.ToolName} cancelled during SSE execution.")
                    Catch ex As Exception
                        sseEx = ex
                    End Try

                    If sseDone OrElse sseCancelled Then Exit Do
                    If sseAttemptedRefresh Then Exit Do
                    If Not ShouldRetryMCPAfterUnauthorized(toolConfig, sseEx) Then Exit Do

                    sseAttemptedRefresh = True
                    ToolingFileLogger.LogWarn(
                        "SSE tool call returned Unauthorized. Forcing MCP OAuth refresh and retrying once.",
                        details:=$"ToolName='{toolCall.ToolName}'; SseBase='{sseBase}'")

                    Dim sseRefreshOk As Boolean = Await ForceRefreshToolOAuthToken(toolConfig, toolCall.ToolName).ConfigureAwait(False)
                    If Not sseRefreshOk Then Exit Do
                Loop

                If Not sseDone AndAlso Not sseCancelled AndAlso sseEx IsNot Nothing Then
                    response.Success = False
                    response.ErrorMessage = $"SSE tool call failed: {sseEx.Message}"
                    ToolingFileLogger.LogError("SSE tool call failed.",
                        details:=$"ToolName='{toolCall.ToolName}'; SseBase='{sseBase}'", ex:=sseEx)
                End If

                Return response
            End If

            ' ── MCP Streamable HTTP transport: full round-trip bypassing LLM() ─
            If IsMCPStreamableToolCall(toolConfig.Endpoint, apiCall) Then
                Dim mcpUrl As String = If(toolConfig.Endpoint, "")
                If mcpUrl.StartsWith(SharedMethods.MCP_STREAMABLE_PREFIX, StringComparison.OrdinalIgnoreCase) Then
                    mcpUrl = mcpUrl.Substring(SharedMethods.MCP_STREAMABLE_PREFIX.Length)
                End If

                Dim resolvedHeaderB = If(toolConfig.HeaderB, "").Replace("{apikey}", If(toolConfig.DecodedAPI, ""))

                context.Log($"MCP Streamable HTTP: executing tool {toolCall.ToolName} via {mcpUrl}")
                ToolingFileLogger.LogStep($"MCP Streamable HTTP round-trip for {toolCall.ToolName} at {mcpUrl}")
                ToolingFileLogger.LogStep($"MCP Streamable HTTP request body: {apiCall}")

                Dim streamAttemptedRefresh As Boolean = False
                Dim streamEx As Exception = Nothing
                Dim streamDone As Boolean = False
                Dim streamCancelled As Boolean = False

                Do
                    streamEx = Nothing

                    Try
                        cancellationToken.ThrowIfCancellationRequested()

                        resolvedHeaderB = If(toolConfig.HeaderB, "").Replace("{apikey}", If(toolConfig.DecodedAPI, ""))

                        Dim rawResult = Await SharedMethods.ExecuteMCPStreamableToolCall(
                            mcpUrl,
                            apiCall,
                            If(toolConfig.HeaderA, ""),
                            resolvedHeaderB,
                            CInt(Math.Min(If(toolConfig.Timeout > 0, toolConfig.Timeout, 60000L), Integer.MaxValue)))

                        ToolingFileLogger.LogRawResponseStub($"MCP Streamable HTTP tool result ({toolCall.ToolName})", rawResult)

                        If Not String.IsNullOrWhiteSpace(rawResult) Then
                            Dim toolErrorMessage As String = ""
                            response.Response = rawResult

                            If TryExtractToolServiceErrorMessage(rawResult, toolErrorMessage) Then
                                response.Success = False
                                response.ErrorMessage = toolErrorMessage
                                ToolingFileLogger.LogWarn(
                                    "MCP Streamable HTTP tool service returned a logical error.",
                                    details:=$"ToolName='{toolCall.ToolName}'; Error='{toolErrorMessage}'")
                            Else
                                response.Success = True
                            End If
                        Else
                            response.Success = False
                            response.ErrorMessage = "Empty response from MCP Streamable HTTP tool service"
                            ToolingFileLogger.LogError(
                                "Empty MCP Streamable HTTP response.",
                                details:=$"ToolName='{toolCall.ToolName}'; Endpoint='{mcpUrl}'")
                        End If

                        streamDone = True

                    Catch ex As OperationCanceledException
                        streamCancelled = True
                        response.Success = False
                        response.ErrorMessage = "Operation was cancelled"
                        ToolingFileLogger.LogWarn(
                            $"Tool {toolCall.ToolName} cancelled during MCP Streamable HTTP execution.")
                    Catch ex As Exception
                        streamEx = ex
                    End Try

                    If streamDone OrElse streamCancelled Then Exit Do
                    If streamAttemptedRefresh Then Exit Do
                    If Not ShouldRetryMCPAfterUnauthorized(toolConfig, streamEx) Then Exit Do

                    streamAttemptedRefresh = True
                    ToolingFileLogger.LogWarn(
                        "MCP Streamable HTTP tool call returned Unauthorized. Forcing MCP OAuth refresh and retrying once.",
                        details:=$"ToolName='{toolCall.ToolName}'; Endpoint='{mcpUrl}'")

                    Dim streamRefreshOk As Boolean = Await ForceRefreshToolOAuthToken(toolConfig, toolCall.ToolName).ConfigureAwait(False)
                    If Not streamRefreshOk Then Exit Do
                Loop

                If Not streamDone AndAlso Not streamCancelled AndAlso streamEx IsNot Nothing Then
                    response.Success = False
                    response.ErrorMessage = $"MCP Streamable HTTP tool call failed: {streamEx.Message}"
                    ToolingFileLogger.LogError(
                        "MCP Streamable HTTP tool call failed.",
                        details:=$"ToolName='{toolCall.ToolName}'; Endpoint='{mcpUrl}'",
                        ex:=streamEx)
                End If

                Return response
            End If

            ' ── Standard transport: route through LLM() ──────────────────
            Dim backupConfig = GetCurrentConfig(_context)

            Try
                cancellationToken.ThrowIfCancellationRequested()

                Dim errorFlag As Boolean = False
                ApplyModelConfig(_context, toolConfig, errorFlag)
                If errorFlag Then
                    response.Success = False
                    response.ErrorMessage = "Failed to apply tool configuration"
                    ToolingFileLogger.LogError("Failed to apply tool configuration.", details:=$"ToolName='{toolCall.ToolName}'")
                    Return response
                End If

                _context.INI_APICall_2 = apiCall

                Dim originalResponse = _context.INI_Response_2
                _context.INI_Response_2 = "JSON"

                context.Log($"Calling external service for tool: {toolCall.ToolName}")
                ToolingFileLogger.LogPreToolLlmCallSnapshot(_context)

                Dim configuredToolTimeoutMs As Long = If(_context.INI_Timeout_2 > 0, _context.INI_Timeout_2, 60000)
                Dim perCallToolTimeoutMs As Integer = SharedLibrary.Agents.HostToolRegistration.GetPerCallLlmTimeoutMs(
                    configuredToolTimeoutMs,
                    New String() {toolCall.ToolName},
                    If(_context.INI_APICall_ToolInstructions_2, "").Length,
                    If(_context.INI_APICall_ToolResponses_2, "").Length)
                If perCallToolTimeoutMs <> configuredToolTimeoutMs Then
                    context.Log($"[PERF] Tool LLM timeout elevated: tool={toolCall.ToolName}; baseMs={configuredToolTimeoutMs}; effectiveMs={perCallToolTimeoutMs}.", "diag")
                End If

                Dim result = Await LLM("", "", "", "", perCallToolTimeoutMs, True, True, "", "", cancellationToken, EnsureUI:=False)

                ToolingFileLogger.LogRawResponseStub($"Tool LLM() result ({toolCall.ToolName})", result)

                _context.INI_Response_2 = originalResponse

                If Not String.IsNullOrWhiteSpace(result) Then
                    Dim toolErrorMessage As String = ""

                    response.Response = result

                    If TryExtractToolServiceErrorMessage(result, toolErrorMessage) Then
                        response.Success = False
                        response.ErrorMessage = toolErrorMessage
                        ToolingFileLogger.LogWarn(
                            "Tool service returned a logical error.",
                            details:=$"ToolName='{toolCall.ToolName}'; Error='{toolErrorMessage}'")
                    Else
                        response.Success = True
                    End If
                Else
                    response.Success = False
                    response.ErrorMessage = "Empty response from tool service"
                    ToolingFileLogger.LogError("Empty response from tool service.", details:=$"ToolName='{toolCall.ToolName}'; APICall='{apiCall}'")
                End If

            Finally
                RestoreDefaults(_context, backupConfig)
            End Try

        Catch ex As OperationCanceledException
            response.Success = False
            response.ErrorMessage = "Operation was cancelled"
            ToolingFileLogger.LogWarn($"Tool {toolCall.ToolName} cancelled during external execution.")

        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            ToolingFileLogger.LogError("Tool execution error.", details:=$"ToolName='{toolCall.ToolName}'", ex:=ex)
        End Try

        Return response
    End Function



    Private Function TryExtractToolServiceErrorMessage(rawResponse As String, ByRef errorMessage As String) As Boolean
        errorMessage = ""

        If String.IsNullOrWhiteSpace(rawResponse) Then
            Return False
        End If

        Try
            Dim root As JObject = JObject.Parse(rawResponse)

            Dim errorToken As JToken = root("error")
            If errorToken IsNot Nothing Then
                Dim message As String = If(errorToken("message"), "").ToString().Trim()
                Dim code As String = If(errorToken("code"), "").ToString().Trim()

                If message = "" Then
                    message = errorToken.ToString(Formatting.None)
                End If

                errorMessage = If(code <> "", $"{code}: {message}", message)
                Return True
            End If

            Dim isErrorToken As JToken = root.SelectToken("result.isError")
            Dim isError As Boolean = False

            If isErrorToken IsNot Nothing Then
                If isErrorToken.Type = JTokenType.Boolean Then
                    isError = isErrorToken.Value(Of Boolean)()
                Else
                    Boolean.TryParse(isErrorToken.ToString(), isError)
                End If
            End If

            If Not isError Then
                Return False
            End If

            Dim messages As New List(Of String)()
            Dim contentArray As JArray = TryCast(root.SelectToken("result.content"), JArray)

            If contentArray IsNot Nothing Then
                For Each item As JToken In contentArray
                    Dim text As String = If(item("text"), "").ToString().Trim()
                    If text <> "" Then
                        messages.Add(text)
                    End If
                Next
            End If

            If messages.Count > 0 Then
                errorMessage = String.Join(" ", messages)
            Else
                Dim resultToken As JToken = root("result")
                errorMessage = If(resultToken Is Nothing, "Tool service returned an error.", resultToken.ToString(Formatting.None))
            End If

            Return True
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Determines whether a response contains tool calls by applying a detection regex pattern.
    ''' If <paramref name="detectionPattern"/> is empty, the pattern is derived from <c>INI_Response_2</c>.
    ''' </summary>
    ''' <param name="response">LLM response text.</param>
    ''' <param name="detectionPattern">Regex pattern used for detection.</param>
    ''' <returns>True if tool calls are detected; otherwise False.</returns>
    Public Function ContainsToolCalls(response As String, detectionPattern As String) As Boolean
        If String.IsNullOrWhiteSpace(response) Then Return False

        Dim pattern As String = detectionPattern
        If String.IsNullOrWhiteSpace(pattern) Then
            pattern = ExtractToolCallPatternFromResponse(INI_Response_2)
        End If

        If String.IsNullOrWhiteSpace(pattern) Then Return False

        Try
            Return Regex.IsMatch(response, pattern, RegexOptions.Singleline Or RegexOptions.CultureInvariant)
        Catch ex As Exception
            ToolingFileLogger.LogError("Regex match error.", details:=$"pattern='{pattern}'", ex:=ex)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Extracts a regex pattern embedded in a response key using the <c>ToolCallMatching*</c> markers.
    ''' </summary>
    ''' <param name="responseKey">Response configuration key string (e.g., <c>INI_Response_2</c>).</param>
    ''' <returns>Extracted regex pattern, or an empty string if not available/invalid.</returns>
    Private Function ExtractToolCallPatternFromResponse(responseKey As String) As String
        If String.IsNullOrEmpty(responseKey) Then
            Return String.Empty
        End If

        Dim startMarker As String = ToolCallMatchingStart
        Dim startIdx As Integer = responseKey.IndexOf(startMarker, StringComparison.OrdinalIgnoreCase)
        If startIdx < 0 Then Return String.Empty

        Dim endIdx As Integer = responseKey.IndexOf(ToolCallMatchingEnd, startIdx, StringComparison.OrdinalIgnoreCase)
        Dim triggerLen As Integer = If(endIdx >= 0,
               (endIdx - startIdx + ToolCallMatchingEnd.Length),
               (responseKey.Length - startIdx))

        Dim triggerText As String = responseKey.Substring(startIdx, triggerLen)

        Dim lt As Integer = triggerText.IndexOf("<"c)
        Dim gt As Integer = triggerText.LastIndexOf(">"c)

        Dim detectedPattern As String = String.Empty

        If lt >= 0 AndAlso gt > lt Then
            detectedPattern = triggerText.Substring(lt + 1, gt - lt - 1).Trim()
        Else
            Dim colonIdx As Integer = triggerText.IndexOf(ToolCallMatchingMiddle, StringComparison.OrdinalIgnoreCase)
            If colonIdx >= 0 Then
                Dim raw As String = triggerText.Substring(colonIdx + ToolCallMatchingMiddle.Length)
                Dim paren As Integer = raw.LastIndexOf(ToolCallMatchingEnd, StringComparison.OrdinalIgnoreCase)
                If paren >= 0 Then raw = raw.Substring(0, paren)
                detectedPattern = raw.Trim()
            End If
        End If

        If Not String.IsNullOrWhiteSpace(detectedPattern) Then
            Try
                Dim rx As New Regex(detectedPattern)
            Catch ex As ArgumentException
                ToolingFileLogger.LogError("Invalid regex pattern.", details:=$"pattern='{detectedPattern}'", ex:=ex)
                Return String.Empty
            End Try
        End If

        Return detectedPattern
    End Function

    ''' <summary>
    ''' Extracts tool calls from a JSON response according to a JSON "extraction map".
    ''' </summary>
    ''' <param name="response">Response text expected to parse as JSON.</param>
    ''' <param name="extractionMap">JSON map specifying paths for tool call array/id/name/arguments.</param>
    ''' <returns>List of extracted tool calls (may be empty).</returns>
    Public Function ExtractToolCalls(response As String, extractionMap As String) As List(Of ToolCall)
        Dim calls As New List(Of ToolCall)()

        If String.IsNullOrWhiteSpace(response) OrElse String.IsNullOrWhiteSpace(extractionMap) Then
            ToolingFileLogger.LogWarn(
                "ExtractToolCalls: Missing response or extractionMap.",
                details:=$"responseEmpty={String.IsNullOrWhiteSpace(response)}; extractionMapEmpty={String.IsNullOrWhiteSpace(extractionMap)}")
            Return calls
        End If

        Try
            Dim jResponse As JToken = JToken.Parse(response)
            Dim jMap As JObject = JObject.Parse(extractionMap)

            Dim arrayPath = If(jMap("array_path")?.ToString(), "")
            Dim callIdPath = If(jMap("call_id_path")?.ToString(), "id")
            Dim namePath = If(jMap("name_path")?.ToString(), "name")
            Dim argsPath = If(jMap("arguments_path")?.ToString(), "arguments")

            Dim toolCallTokens As IEnumerable(Of JToken)

            If Not String.IsNullOrWhiteSpace(arrayPath) Then
                toolCallTokens = jResponse.SelectTokens(arrayPath).ToList()
            Else
                toolCallTokens = {jResponse}
            End If

            For Each tcToken In toolCallTokens
                Try
                    Dim tc As New ToolCall() With {
                        .CallId = If(tcToken.SelectToken(callIdPath)?.ToString(), Guid.NewGuid().ToString("N")),
                        .ToolName = If(tcToken.SelectToken(namePath)?.ToString(), ""),
                        .RawJson = tcToken.ToString()
                    }

                    Dim argsToken = tcToken.SelectToken(argsPath)
                    If argsToken IsNot Nothing Then
                        If argsToken.Type = JTokenType.String Then
                            Try
                                tc.Arguments = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(argsToken.ToString())
                            Catch ex As Exception
                                tc.Arguments = New Dictionary(Of String, Object) From {{"raw", argsToken.ToString()}}
                                ToolingFileLogger.LogWarn(
                                    "Arguments JSON string could not be deserialized; stored in 'raw'.",
                                    details:=$"ToolName='{tc.ToolName}'; CallId={tc.CallId}",
                                    ex:=ex)
                            End Try
                        Else
                            tc.Arguments = argsToken.ToObject(Of Dictionary(Of String, Object))()
                        End If
                    End If

                    If Not String.IsNullOrWhiteSpace(tc.ToolName) Then
                        calls.Add(tc)
                    Else
                        ToolingFileLogger.LogWarn("ExtractToolCalls: Skipped tool call with empty ToolName.", details:=$"Raw={tc.RawJson}")
                    End If
                Catch ex As Exception
                    ToolingFileLogger.LogError("Error parsing individual tool call.", ex:=ex)
                    Debug.WriteLine($"Error parsing individual tool call: {ex.Message}")
                End Try
            Next

        Catch ex As Exception
            ToolingFileLogger.LogError("ExtractToolCalls error.", details:=$"extractionMap='{extractionMap}'", ex:=ex)
            Debug.WriteLine($"ExtractToolCalls error: {ex.Message}")
        End Try

        Return calls
    End Function

    ''' <summary>
    ''' Determines whether a string represents a JSON object or array.
    ''' </summary>
    ''' <param name="str">Candidate JSON string.</param>
    ''' <returns>True if valid JSON object/array; otherwise False.</returns>
    Private Function IsValidJson(str As String) As Boolean
        If String.IsNullOrWhiteSpace(str) Then Return False
        str = str.Trim()
        If (str.StartsWith("{") AndAlso str.EndsWith("}")) OrElse
           (str.StartsWith("[") AndAlso str.EndsWith("]")) Then
            Try
                JToken.Parse(str)
                Return True
            Catch
                Return False
            End Try
        End If
        Return False
    End Function

    ''' <summary>
    ''' Escapes a string for safe embedding into a JSON string literal.
    ''' </summary>
    ''' <param name="str">Input string.</param>
    ''' <returns>Escaped string content (without surrounding quotes).</returns>
    Private Function EscapeJsonString(str As String) As String
        If String.IsNullOrEmpty(str) Then Return ""

        Dim sb As New StringBuilder()
        For Each c As Char In str
            Select Case c
                Case """"c : sb.Append("\""")
                Case "\"c : sb.Append("\\")
                Case "/"c : sb.Append("\/")
                Case ChrW(8) : sb.Append("\b")   ' Backspace
                Case ChrW(12) : sb.Append("\f")  ' Form feed
                Case vbLf(0) : sb.Append("\n")
                Case vbCr(0) : sb.Append("\r")
                Case vbTab(0) : sb.Append("\t")
                Case Else
                    If AscW(c) < 32 Then
                        ' Other control characters
                        sb.Append("\u" & AscW(c).ToString("X4"))
                    Else
                        sb.Append(c)
                    End If
            End Select
        Next
        Return sb.ToString()
    End Function



End Class