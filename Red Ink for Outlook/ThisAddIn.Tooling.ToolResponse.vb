' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.Tooling.ToolResponse.vb
' Purpose: Tool response building, serialization, and formatting utilities.
'          Prepares tool execution results for model injection and user display.
'
' Architecture:
'  - Model Response Serialization:
'      - BuildToolResponsesForModel(): Converts List(Of ToolResponse) to model-specific JSON payload.
'      - BuildToolResponseContentForModel(): Formats individual response content.
'      - Supports compaction for sub-agent contexts via CompactToolResponseContentForSubAgent().
'      - Handles both success and structured error responses (resultKind="error").
'  - Response Formatting & Display:
'      - BuildResultExcerpt(): Creates brief summaries for log window display.
'      - BuildCondensedParamSummary(): Formats tool call parameters for diagnostics.
'      - BuildToolReplaySummary(): Extracts summary field from structured responses.
'  - Sub-Agent Response Handling:
'      - CompactToolResponseContentForSubAgent(): Truncates large responses with continuation hints.
'      - Tracks compaction state via WasCompactedForModelReplay and ModelReplayContent.
'  - Recovery Prompts:
'      - BuildSubAgentEmptyResponseRecoveryPrompt(): Generates repair prompt for empty sub-agent turns.
'      - GetLastSuccessfulToolResponse(): Retrieves most recent successful tool execution.
'  - Template-Based Serialization:
'      - Uses APICall_ToolResponses, APICall_ToolResponses_Template, APICall_ToolCallPart_Template.
'      - Supports model-agnostic placeholders: {call_id}, {name}, {arguments}, {response}.
'      - Handles quoted vs. raw JSON injection based on template structure.
'      - Special handling for Gemini-style functionResponse/function_response payloads.
'
' Key Functions:
'  - BuildToolResponsesForModel(): Primary serialization entry point.
'  - BuildResultExcerpt(): User-friendly result summaries.
'  - BuildCondensedParamSummary(): Parameter display formatting.
'  - CompactToolResponseContentForSubAgent(): Large response truncation.
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
    ''' Builds the model-specific tool response payload to inject into the next iteration of the tooling loop.
    ''' </summary>
    ''' <param name="responses">Tool execution outcomes to serialize.</param>
    ''' <param name="toolingModel">Tooling model that defines response templates and container structure.</param>
    ''' <returns>Serialized tool response payload.</returns>
    Public Function BuildToolResponsesForModel(responses As List(Of ToolResponse),
                                           toolingModel As ModelConfig,
                                           Optional compactForSubAgent As Boolean = False,
                                           Optional compactStaleLargeResponses As Boolean = False,
                                           Optional keepRecentFullCount As Integer = 2,
                                           Optional staleCompactionThresholdChars As Integer = -1,
                                           Optional staleCompactionPreviewChars As Integer = -1) As String
        If toolingModel Is Nothing Then
            ToolingFileLogger.LogWarn("BuildToolResponsesForModel: toolingModel is Nothing.")
            Return ""
        End If

        If String.IsNullOrWhiteSpace(toolingModel.APICall_ToolResponses) Then
            ToolingFileLogger.LogWarn("BuildToolResponsesForModel: toolingModel.APICall_ToolResponses is empty.")
            Return ""
        End If

        Dim responsePartTemplate As String = toolingModel.APICall_ToolResponses_Template
        If String.IsNullOrWhiteSpace(responsePartTemplate) Then
            ToolingFileLogger.LogWarn("BuildToolResponsesForModel: toolingModel.APICall_ToolResponses_Template is empty.")
            Return ""
        End If

        Dim callPartTemplate As String = If(toolingModel.APICall_ToolCallPart_Template, "")
        Dim useCallParts As Boolean = Not String.IsNullOrWhiteSpace(callPartTemplate)

        Dim callParts As New StringBuilder()
        Dim responseParts As New StringBuilder()
        Dim firstCall As Boolean = True
        Dim firstResp As Boolean = True

        Dim responseCount As Integer = If(responses Is Nothing, 0, responses.Count)
        Dim respIndex As Integer = -1
        For Each resp In responses
            respIndex += 1
            ' A response whose own body exceeds the large-result threshold is always
            ' compaction-eligible, regardless of recency: it is stored by reference and
            ' remains fully retrievable via context_expand, so keeping it "recent-full"
            ' would only bloat the payload. Smaller responses keep the recency exemption.
            Dim respIsLarge As Boolean =
                resp IsNot Nothing AndAlso
                Not String.IsNullOrEmpty(resp.Response) AndAlso
                resp.Response.Length > SubAgentLargeToolResponseThresholdChars
            Dim isStaleForCompaction As Boolean =
                compactStaleLargeResponses AndAlso
                (respIsLarge OrElse (respIndex < responseCount - keepRecentFullCount))
            If useCallParts Then
                ' Extract the original arguments from the parsed tool call JSON
                Dim argsJson As String = "{}"
                Try
                    Dim jCall = JObject.Parse(resp.OriginalCallJson)
                    Dim argsToken = jCall("arguments")
                    If argsToken IsNot Nothing Then
                        If argsToken.Type = JTokenType.String Then
                            argsJson = argsToken.ToString()
                        Else
                            argsJson = argsToken.ToString(Formatting.None)
                        End If
                    End If
                Catch
                    argsJson = "{}"
                End Try

                ' Determine if arguments should be escaped (template has quoted placeholder)
                Dim escapedArgsJson As String
                If callPartTemplate.Contains("""{arguments}""") Then
                    escapedArgsJson = EscapeJsonString(argsJson)
                Else
                    escapedArgsJson = argsJson
                End If

                ' Build the call part, also support {call} placeholder for raw call JSON
                Dim callPart As String = callPartTemplate _
                    .Replace("{call_id}", If(resp.CallId, "")) _
                    .Replace("{name}", If(resp.ToolName, "")) _
                    .Replace("{arguments}", escapedArgsJson) _
                    .Replace("{call}", resp.OriginalCallJson)

                If Not firstCall Then callParts.Append(",")
                callParts.Append(callPart)
                firstCall = False
            End If

            ' Build response content. Under budget pressure, older stale results may be
            ' reference-compacted using a lower threshold/preview so medium-sized results
            ' also move into the drawer; everything stays retrievable via context_expand.
            Dim effStaleThresholdChars As Integer = -1
            Dim effStalePreviewChars As Integer = -1
            If isStaleForCompaction AndAlso staleCompactionThresholdChars > 0 Then
                effStaleThresholdChars = staleCompactionThresholdChars
                effStalePreviewChars = staleCompactionPreviewChars
            End If
            Dim responseContent As String = BuildToolResponseContentForModel(resp, compactForSubAgent OrElse isStaleForCompaction, effStaleThresholdChars, effStalePreviewChars)

            ' Model-agnostic handling:
            ' - If the response placeholder is quoted, emit an escaped string.
            ' - If the template is a Gemini-style functionResponse/function_response payload,
            '   force the inserted response to be a JSON object (arrays/scalars wrapped).
            ' - Otherwise preserve raw valid JSON for providers that accept arrays/scalars.
            Dim finalResponseContent As String
            Dim templateRequiresQuotedString As Boolean = responsePartTemplate.Contains("""{response}""")
            Dim templateLooksLikeGeminiFunctionResponse As Boolean =
                    responsePartTemplate.IndexOf("functionResponse", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                    responsePartTemplate.IndexOf("function_response", StringComparison.OrdinalIgnoreCase) >= 0

            If templateRequiresQuotedString Then
                finalResponseContent = EscapeJsonString(responseContent)
            ElseIf responsePartTemplate.Contains("{response}") Then
                Try
                    Dim parsed As JToken = JToken.Parse(responseContent)

                    If templateLooksLikeGeminiFunctionResponse Then
                        If TypeOf parsed Is JObject Then
                            finalResponseContent = parsed.ToString(Formatting.None)
                        ElseIf TypeOf parsed Is JArray Then
                            finalResponseContent = New JObject(
                                    New JProperty("items", parsed)
                                ).ToString(Formatting.None)
                        Else
                            finalResponseContent = New JObject(
                                    New JProperty("result", parsed)
                                ).ToString(Formatting.None)
                        End If
                    Else
                        finalResponseContent = parsed.ToString(Formatting.None)
                    End If
                Catch
                    finalResponseContent = New JObject(
                            New JProperty("result", responseContent)
                        ).ToString(Formatting.None)
                End Try
            Else
                finalResponseContent = EscapeJsonString(responseContent)
            End If

            Dim respPart As String = responsePartTemplate _
                .Replace("{call_id}", If(resp.CallId, "")) _
                .Replace("{name}", If(resp.ToolName, "")) _
                .Replace("{response}", finalResponseContent)

            If Not firstResp Then responseParts.Append(",")
            responseParts.Append(respPart)
            firstResp = False
        Next

        Dim functionCallsOutput As String = callParts.ToString()
        Dim responsesOutput As String = responseParts.ToString()

        ' Replace placeholders - NO comma manipulation by code
        ' Templates are responsible for their own structure
        Dim result As String = toolingModel.APICall_ToolResponses

        ' Simple replacement - if content exists, replace; if empty, remove placeholder
        result = result.Replace("{functioncalls}", functionCallsOutput)
        result = result.Replace("{responses}", responsesOutput)

        ' Clean up any empty structural remnants (empty arrays, double commas, etc.)
        ' This handles cases where one placeholder was empty
        result = Regex.Replace(result, "\[\s*\]", "[]")           ' Normalize empty arrays
        result = Regex.Replace(result, ",\s*,", ",")              ' Remove double commas
        result = Regex.Replace(result, "\[\s*,", "[")             ' Remove leading comma in array
        result = Regex.Replace(result, ",\s*\]", "]")             ' Remove trailing comma in array

        Return result
    End Function

    ''' <summary>
    ''' Wraps <see cref="BuildToolResponsesForModel"/> with a payload-size budget. The first
    ''' pass keeps the most recent results fully visible. Only when the overall payload grows
    ''' beyond the budget does it progressively shrink the recent-full window and then
    ''' reference-compact older medium-sized results using lower thresholds. Everything moved
    ''' this way stays fully retrievable via context_expand, so compaction is lossless. The
    ''' model can also voluntarily tighten this via the context_compact tool.
    ''' Capability-driven: no tool-name or content-type heuristics.
    ''' </summary>
    Public Function BuildToolResponsesForModelBudgeted(responses As List(Of ToolResponse),
                                                       toolingModel As ModelConfig,
                                                       Optional compactForSubAgent As Boolean = False) As String
        Dim keepRecentFullCount As Integer = 2

        Dim requestedKeep As Integer
        If SharedLibrary.Agents.ToolResultStore.TryGetRequestedKeepRecent(
                SharedLibrary.Agents.WorkflowContinuity.CurrentWorkflowId, requestedKeep) Then
            keepRecentFullCount = Math.Min(keepRecentFullCount, requestedKeep)
        End If

        Dim payload As String = BuildToolResponsesForModel(
            responses,
            toolingModel,
            compactForSubAgent:=compactForSubAgent,
            compactStaleLargeResponses:=True,
            keepRecentFullCount:=keepRecentFullCount)

        Dim budget As Integer =
            If(ThisAddIn.INI_ToolResponsePayloadBudgetChars > 0,
               ThisAddIn.INI_ToolResponsePayloadBudgetChars,
               SharedLibrary.Agents.ToolingConstants.ToolResponsePayloadBudgetChars)
        Dim mediumThreshold As Integer =
            If(ThisAddIn.INI_BudgetMediumCompactionThresholdChars > 0,
               ThisAddIn.INI_BudgetMediumCompactionThresholdChars,
               SharedLibrary.Agents.ToolingConstants.BudgetMediumCompactionThresholdChars)
        Dim aggressiveThreshold As Integer =
            If(ThisAddIn.INI_BudgetAggressiveCompactionThresholdChars > 0,
               ThisAddIn.INI_BudgetAggressiveCompactionThresholdChars,
               SharedLibrary.Agents.ToolingConstants.BudgetAggressiveCompactionThresholdChars)
        Dim previewChars As Integer =
            If(ThisAddIn.INI_BudgetCompactionPreviewChars > 0,
               ThisAddIn.INI_BudgetCompactionPreviewChars,
               SharedLibrary.Agents.ToolingConstants.BudgetCompactionPreviewChars)
        If budget <= 0 OrElse String.IsNullOrEmpty(payload) OrElse payload.Length <= budget Then
            Return payload
        End If

        ' Stage 1: shrink the recent-full window (down to 0). Lossless.
        While payload.Length > budget AndAlso keepRecentFullCount > 0
            keepRecentFullCount -= 1
            payload = BuildToolResponsesForModel(
                responses,
                toolingModel,
                compactForSubAgent:=compactForSubAgent,
                compactStaleLargeResponses:=True,
                keepRecentFullCount:=keepRecentFullCount)
        End While

        If payload.Length <= budget Then
            Return payload
        End If

        ' Stage 2: reference-compact older medium-sized results using progressively
        ' lower thresholds until the payload fits or the floor is reached.
        For Each mediumThresholdChars As Integer In New Integer() {mediumThreshold, aggressiveThreshold}
            payload = BuildToolResponsesForModel(
                responses,
                toolingModel,
                compactForSubAgent:=compactForSubAgent,
                compactStaleLargeResponses:=True,
                keepRecentFullCount:=0,
                staleCompactionThresholdChars:=mediumThresholdChars,
                staleCompactionPreviewChars:=previewChars)
            If payload.Length <= budget Then
                Exit For
            End If
        Next

        If payload.Length > budget Then
            ToolingFileLogger.LogWarn(
                "Tool response payload still exceeds budget after progressive compaction.",
                details:=$"payloadChars={payload.Length}; budgetChars={budget}")
        End If

        Return payload
    End Function

    Private Function BuildToolResponseContentForModel(resp As ToolResponse,
                                                   Optional compactForSubAgent As Boolean = False,
                                                   Optional overrideThresholdChars As Integer = -1,
                                                   Optional overridePreviewChars As Integer = -1) As String
        If resp Is Nothing Then Return ""

        Dim rawContent As String

        If resp.Success Then
            rawContent = If(resp.Response, "")
        ElseIf IsStructuredErrorToolResponse(resp) Then
            rawContent = If(resp.Response, "")
        Else
            rawContent = $"Error: {If(resp.ErrorMessage, "Tool failed.")}"
        End If

        If Not compactForSubAgent Then
            Return rawContent
        End If

        Return CompactToolResponseContentForSubAgent(resp, rawContent, overrideThresholdChars, overridePreviewChars)
    End Function

    ''' <summary>
    ''' Returns True when a large tool response must be replayed in full rather than
    ''' compacted, because truncation could drop deliverable/M365 reference fields
    ''' (path, saved_path, output_reference, memory_key, reference) that downstream
    ''' logic relies on. Capability-driven: no tool-name-specific heuristics.
    ''' </summary>
    Private Function MustPreserveFullResponseForReplay(resp As ToolResponse, rawContent As String) As Boolean
        If resp Is Nothing Then Return False

        Dim toolName As String = If(resp.ToolName, "").Trim()
        If toolName <> "" Then
            Try
                Dim deliverableTools = SharedLibrary.Agents.HostToolRegistration.GetDeliverableCapableToolNames(
                    SharedLibrary.Agents.ToolingHostKind.Outlook)
                If deliverableTools IsNot Nothing AndAlso deliverableTools.Contains(toolName) Then
                    Return True
                End If
            Catch
            End Try
        End If

        Dim raw As String = If(rawContent, "")
        If raw <> "" Then
            Try
                If Not String.IsNullOrWhiteSpace(
                    SharedLibrary.Agents.WorkflowContinuity.ExtractStructuredResultReference(raw)) Then
                    Return True
                End If

                If Not String.IsNullOrWhiteSpace(
                    SharedLibrary.Agents.WorkflowContinuity.ExtractOutputReference(raw)) Then
                    Return True
                End If
            Catch
            End Try
        End If

        Return False
    End Function

    ''' <summary>
    ''' When a stale response is a previously expanded window (carries a result_ref plus a
    ''' content_window), returns a compact stub that keeps the navigation pointer but drops
    ''' the window body. The full result remains retrievable via context_expand, so this is
    ''' lossless. Returns Nothing when the response is not a windowed reference.
    ''' </summary>
    Private Function TryBuildReferencedWindowStub(rawContent As String) As String
        Dim raw As String = If(rawContent, "")
        If raw = "" Then Return Nothing

        Dim obj As JObject
        Try
            obj = JObject.Parse(raw)
        Catch
            Return Nothing
        End Try

        Dim refToken = obj("result_ref")
        Dim windowToken = obj("content_window")
        If refToken Is Nothing OrElse windowToken Is Nothing Then Return Nothing

        Dim refValue As String = refToken.ToString()
        If String.IsNullOrWhiteSpace(refValue) Then Return Nothing

        Dim windowLength As Integer = windowToken.ToString().Length
        If windowLength <= 1024 Then Return Nothing

        Dim toolValue As String = If(obj("tool") IsNot Nothing, obj("tool").ToString(), "")

        Dim stub As New JObject(
            New JProperty("ok", True),
            New JProperty("tool", toolValue),
            New JProperty("result_ref", refValue),
            New JProperty("start_char", obj("start_char")),
            New JProperty("returned_chars", obj("returned_chars")),
            New JProperty("total_chars", obj("total_chars")),
            New JProperty("next_offset", obj("next_offset")),
            New JProperty("truncated", obj("truncated")),
            New JProperty("omitted_window_chars", windowLength),
            New JProperty("note", "A previously expanded window was omitted from context to save space. Call context_expand with this result_ref and the offsets to re-read it."))

        Return stub.ToString(Formatting.None)
    End Function

    Private Function CompactToolResponseContentForSubAgent(resp As ToolResponse, rawContent As String,
                                                           Optional overrideThresholdChars As Integer = -1,
                                                           Optional overridePreviewChars As Integer = -1) As String
        If resp Is Nothing Then Return If(rawContent, "")

        Dim raw As String = If(rawContent, "")

        Dim thresholdChars As Integer =
            If(overrideThresholdChars > 0, overrideThresholdChars, SubAgentLargeToolResponseThresholdChars)
        Dim previewChars As Integer =
            If(overridePreviewChars > 0, overridePreviewChars, SubAgentLargeToolResponseExcerptChars)

        Dim windowStub As String = TryBuildReferencedWindowStub(raw)
        If windowStub IsNot Nothing Then
            resp.ModelReplayContent = windowStub
            resp.ModelReplaySummary = BuildToolReplaySummary(resp)
            resp.WasCompactedForModelReplay = True
            Return windowStub
        End If

        If raw.Length <= thresholdChars Then
            resp.ModelReplayContent = raw
            resp.ModelReplaySummary = BuildToolReplaySummary(resp)
            resp.WasCompactedForModelReplay = False
            Return raw
        End If

        If MustPreserveFullResponseForReplay(resp, raw) Then
            resp.ModelReplayContent = raw
            resp.ModelReplaySummary = BuildToolReplaySummary(resp)
            resp.WasCompactedForModelReplay = False
            Return raw
        End If

        Dim excerptLength As Integer = Math.Min(previewChars, raw.Length)
        Dim excerpt As String = raw.Substring(0, excerptLength)
        Dim summary As String = BuildToolReplaySummary(resp)

        Dim stored As SharedLibrary.Agents.ToolResultStore.StoredResult =
            SharedLibrary.Agents.ToolResultStore.Put(
                SharedLibrary.Agents.WorkflowContinuity.CurrentWorkflowId,
                If(resp.ToolName, ""),
                raw)

        Dim compactObj As New JObject(
        New JProperty("ok", resp.Success),
        New JProperty("tool", If(resp.ToolName, "")),
        New JProperty("summary", summary),
        New JProperty("result_ref", stored.Ref),
        New JProperty("preview", excerpt),
        New JProperty("total_chars", raw.Length),
        New JProperty("returned_chars", excerptLength),
        New JProperty("truncated", True),
        New JProperty("next_offset", excerptLength),
        New JProperty("continuation", "The full result is stored. To read more, call context_expand with this result_ref, using start_char and max_chars to page through the full content."))

        resp.ModelReplayContent = compactObj.ToString(Formatting.None)
        resp.ModelReplaySummary = summary
        resp.WasCompactedForModelReplay = True
        Return resp.ModelReplayContent
    End Function

    Private Function BuildToolReplaySummary(resp As ToolResponse) As String
        If resp Is Nothing Then Return ""

        If Not String.IsNullOrWhiteSpace(resp.ModelReplaySummary) Then
            Return resp.ModelReplaySummary
        End If

        Dim summary As String = ""

        If Not String.IsNullOrWhiteSpace(resp.Response) Then
            Try
                Dim tok As JToken = JToken.Parse(resp.Response)
                If TypeOf tok Is JObject Then
                    summary = DirectCast(tok, JObject).Value(Of String)("summary")
                End If
            Catch
            End Try
        End If

        If String.IsNullOrWhiteSpace(summary) Then
            summary = $"{If(resp.ToolName, "tool")} succeeded. {BuildResultExcerpt(If(resp.Response, ""), 280)}"
        End If

        resp.ModelReplaySummary = summary
        Return summary
    End Function



    ''' <summary>
    ''' Builds a brief excerpt of the tool result for display in the log window.
    ''' </summary>
    ''' <param name="result">Full tool response text.</param>
    ''' <param name="maxExcerptLength">Maximum length for the excerpt portion.</param>
    ''' <returns>Formatted string like "12,345 chars: 'The quick brown fox...'".</returns>
    Private Function BuildResultExcerpt(result As String, Optional maxExcerptLength As Integer = 80) As String
        If String.IsNullOrEmpty(result) Then
            Return "0 chars (empty)"
        End If

        Dim charCount As Integer = result.Length
        Dim formattedCount As String = charCount.ToString("N0")

        ' Clean up the result for excerpt (remove excessive whitespace/newlines)
        Dim cleaned As String = Regex.Replace(result, "\s+", " ").Trim()

        If cleaned.Length <= maxExcerptLength Then
            Return $"{formattedCount} chars: '{cleaned}'"
        End If

        ' Truncate and add ellipsis
        Dim excerpt As String = cleaned.Substring(0, maxExcerptLength - 3) & "..."
        Return $"{formattedCount} chars: '{excerpt}'"
    End Function

    ''' <summary>
    ''' Builds a condensed parameter summary for display in the log window.
    ''' </summary>
    ''' <param name="arguments">Tool call arguments dictionary.</param>
    ''' <param name="maxLength">Maximum length for each parameter value display.</param>
    ''' <returns>Formatted parameter string like " (query: 'search term', count: 10)".</returns>
    Private Function BuildCondensedParamSummary(arguments As Dictionary(Of String, Object), Optional maxLength As Integer = 50) As String
        If arguments Is Nothing OrElse arguments.Count = 0 Then
            Return ""
        End If

        Dim parts As New List(Of String)()

        For Each kvp In arguments
            Dim valueStr As String = ""
            If kvp.Value IsNot Nothing Then
                If TypeOf kvp.Value Is JArray Then
                    Dim arr = DirectCast(kvp.Value, JArray)
                    valueStr = $"[{arr.Count} items]"
                ElseIf TypeOf kvp.Value Is IEnumerable(Of Object) AndAlso Not TypeOf kvp.Value Is String Then
                    valueStr = $"[{DirectCast(kvp.Value, IEnumerable(Of Object)).Count()} items]"
                Else
                    valueStr = kvp.Value.ToString()
                    ' Use shorter limit for long text parameters like "instruction"
                    Dim effectiveMax = If(valueStr.Length > 200, Math.Min(maxLength, 80), maxLength)
                    If valueStr.Length > effectiveMax Then
                        valueStr = valueStr.Substring(0, effectiveMax - 3) & "..."
                    End If
                End If
            End If

            parts.Add($"{kvp.Key}: '{valueStr}'")
        Next

        Return $" ({String.Join(", ", parts)})"
    End Function


    Private Function GetLastSuccessfulToolResponse(context As ToolExecutionContext) As ToolResponse
        If context Is Nothing OrElse context.AllToolResponses Is Nothing Then Return Nothing

        For i As Integer = context.AllToolResponses.Count - 1 To 0 Step -1
            Dim resp = context.AllToolResponses(i)
            If resp IsNot Nothing AndAlso resp.Success Then
                Return resp
            End If
        Next

        Return Nothing
    End Function



End Class
