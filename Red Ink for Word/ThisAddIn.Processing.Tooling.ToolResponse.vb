' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.Processing.Tooling.ToolResponse.vb
' Purpose: Tool response data model and response formatting for model replay.
'
' Responsibilities:
'  - Define ToolResponse class (call ID, tool name, response content, success/error state, timestamps).
'  - Build response content for model injection (success vs. error payload formatting).
'  - Compact large tool responses for sub-agent context efficiency.
'  - Generate tool response summaries (excerpts for display).
'  - Extract summary fields from structured JSON responses.
'  - Track compaction state for model replay (full vs. excerpt mode).
'  - Extract tool service error messages (structured vs. unstructured).
'  - Retrieve last successful tool response from session history.
'  - Build sub-agent empty-response recovery prompts.
'  - Format tool responses for continuation guards.
'
' Architecture:
'  - ToolResponse as value object holding execution outcome.
'  - Support both structured (JSON) and unstructured (text) responses.
'  - Adaptive formatting based on template requirements (quoted string vs. raw JSON).
'  - Threshold-based compaction to avoid token bloat in sub-agent contexts.
'
' External Dependencies:
'  - Newtonsoft.Json for JSON parsing and compaction.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods


Partial Public Class ThisAddIn

    Public Class ToolResponse

        ''' <summary>Tool call identifier used to correlate call and response objects.</summary>
        Public Property CallId As String

        ''' <summary>Name of the tool that was executed.</summary>
        Public Property ToolName As String

        ''' <summary>Raw response returned by the tool execution.</summary>
        Public Property Response As String

        ''' <summary>True if the tool execution completed successfully; otherwise False.</summary>
        Public Property Success As Boolean

        ''' <summary>Error message populated when <see cref="Success"/> is False.</summary>
        Public Property ErrorMessage As String

        ''' <summary>Timestamp captured at response creation time.</summary>
        Public Property Timestamp As DateTime

        ''' <summary>Original tool call JSON as extracted from the LLM response.</summary>
        Public Property OriginalCallJson As String

        Public Property ResultKind As String
        Public Property ErrorCode As String

        Public Property ModelReplayContent As String
        Public Property ModelReplaySummary As String
        Public Property WasCompactedForModelReplay As Boolean
        Public Property NormalizedCallSignature As String
        Public Property WasDuplicateReplay As Boolean

        ''' <summary>True when a repair-loop advisor determined the failure is terminal (budget exhausted or non-recoverable).</summary>
        Public Property RepairLoopTerminal As Boolean

        ''' <summary>Human-readable reason for <see cref="RepairLoopTerminal"/>, surfaced to abort/finalization handling.</summary>
        Public Property RepairLoopTerminalReason As String

        ''' <summary>
        ''' Initializes a new tool response instance with default success state.
        ''' </summary>
        Public Sub New()
            Timestamp = DateTime.Now
            Success = True
        End Sub
    End Class


    Private Function BuildToolResponseContentForModel(resp As ToolResponse,
                                                  Optional compactForSubAgent As Boolean = False) As String
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

        Return CompactToolResponseContentForSubAgent(resp, rawContent)
    End Function

    Private Function CompactToolResponseContentForSubAgent(resp As ToolResponse, rawContent As String) As String
        If resp Is Nothing Then Return If(rawContent, "")

        Dim raw As String = If(rawContent, "")
        If raw.Length <= SubAgentLargeToolResponseThresholdChars Then
            resp.ModelReplayContent = raw
            resp.ModelReplaySummary = BuildToolReplaySummary(resp)
            resp.WasCompactedForModelReplay = False
            Return raw
        End If

        Dim excerptLength As Integer = Math.Min(SubAgentLargeToolResponseExcerptChars, raw.Length)
        Dim excerpt As String = raw.Substring(0, excerptLength)
        Dim summary As String = BuildToolReplaySummary(resp)

        Dim compactObj As New JObject(
        New JProperty("ok", resp.Success),
        New JProperty("tool", If(resp.ToolName, "")),
        New JProperty("summary", summary),
        New JProperty("content_excerpt", excerpt),
        New JProperty("total_chars", raw.Length),
        New JProperty("returned_chars", excerptLength),
        New JProperty("truncated", True),
        New JProperty("next_offset", excerptLength),
        New JProperty("continuation", "If more content is needed, call the same tool again with a smaller window using max_chars and start_char/offset."))

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

    Private Function BuildSubAgentEmptyResponseRecoveryPrompt(context As ToolExecutionContext) As String
        Dim lastSuccess = GetLastSuccessfulToolResponse(context)
        Dim summary As String = BuildToolReplaySummary(lastSuccess)

        Dim sb As New System.Text.StringBuilder()
        sb.Append("SUB-AGENT EMPTY-RESPONSE RECOVERY: The previous model turn was empty after a successful tool call. ")
        sb.Append("Do not repeat a large raw tool response. ")
        sb.Append("In THIS turn you must either return the required final JSON object, or call one smaller follow-up tool. ")
        sb.Append("If more source text is needed, request a smaller window using max_chars and start_char/offset.")

        If Not String.IsNullOrWhiteSpace(summary) Then
            sb.AppendLine()
            sb.Append("Last successful tool result: ")
            sb.Append(summary)
        End If

        Return sb.ToString()
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



End Class
