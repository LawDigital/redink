' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.Tooling.PromptBuilding.vb
' Purpose: Constructs prompt augmentations and context blocks for tooling loop execution.
'          Builds authoritative user request metadata, completed facts, and continuation blocks.
'
' Architecture:
'  - User Request Metadata:
'      - BuildLatestUserRequestMetadataBlock(): Creates [CURRENT_USER_REQUEST] XML block.
'          - Wraps LatestUserRequestRaw and optional HostTaskSummary.
'          - Marked as authoritative to prevent LLM reinterpretation.
'      - BuildPromptWithAuthoritativeLatestUserRequest(): Injects metadata block into prompt body.
'  - Completed Facts Block:
'      - BuildCompletedFactsPromptBlock(): Extracts and formats completed tool responses.
'          - Collects successful responses up to max items limit.
'          - Uses BuildToolReplaySummary() to humanize each fact.
'          - Wraps in <COMPLETED_FACTS> XML tags for model consumption.
'  - Continuation Blocks:
'      - BuildPostToolContinuationBlock(): Routes to DeliverableCompletionContinuationBlock.
'      - Prepares model for final response after tool execution completes.
'  - Diagnostics:
'      - BuildPromptDiagnosticStub(): SHA256-fingerprints prompts for logging.
'          - Records length, hash (first 16 hex chars), and excerpt (120 chars).
'          - Used by LogLatestUserRequestDiagnostic() at various pipeline stages.
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

Partial Public Class ThisAddIn


    Private Function BuildLatestUserRequestMetadataBlock(context As ToolExecutionContext) As String
        If context Is Nothing OrElse String.IsNullOrWhiteSpace(context.LatestUserRequestRaw) Then
            Return ""
        End If

        Dim sb As New System.Text.StringBuilder()

        sb.AppendLine("[CURRENT_USER_REQUEST]")
        sb.AppendLine("LATEST_USER_REQUEST_RAW is authoritative for this run.")
        sb.AppendLine("Do not replace, narrow, or reinterpret it based on prior context, memory stubs, workflow summaries, host status text, or completed subtasks.")
        sb.AppendLine("<LATEST_USER_REQUEST_RAW>")
        sb.AppendLine(context.LatestUserRequestRaw)
        sb.AppendLine("</LATEST_USER_REQUEST_RAW>")

        If Not String.IsNullOrWhiteSpace(context.HostTaskSummary) Then
            sb.AppendLine("<HOST_TASK_SUMMARY>")
            sb.AppendLine(context.HostTaskSummary)
            sb.AppendLine("</HOST_TASK_SUMMARY>")
        End If

        If context.SequencingState IsNot Nothing Then sb.AppendLine(context.SequencingState.BuildOutputContractPrompt())
        sb.AppendLine("[/CURRENT_USER_REQUEST]")
        Return sb.ToString().TrimEnd()
    End Function

    Private Function BuildPromptWithAuthoritativeLatestUserRequest(context As ToolExecutionContext,
                                                                   promptBody As String) As String
        Dim requestBlock As String = BuildLatestUserRequestMetadataBlock(context)

        If String.IsNullOrWhiteSpace(requestBlock) Then
            Return If(promptBody, "")
        End If

        If String.IsNullOrWhiteSpace(promptBody) Then
            Return requestBlock
        End If

        Return requestBlock & Environment.NewLine & Environment.NewLine & promptBody
    End Function

    Private Function BuildCompletedFactsPromptBlock(context As ToolExecutionContext,
                                                    Optional maxItems As Integer = 3) As String
        If context Is Nothing OrElse context.AllToolResponses Is Nothing OrElse context.AllToolResponses.Count = 0 Then
            Return ""
        End If

        Dim facts As New List(Of String)()

        For Each resp In context.AllToolResponses
            If resp Is Nothing OrElse Not resp.Success Then Continue For

            Dim summary As String = Regex.Replace(If(BuildToolReplaySummary(resp), ""), "\s+", " ").Trim()
            If summary = "" Then Continue For

            facts.Add("- " & summary)

            If facts.Count >= maxItems Then
                Exit For
            End If
        Next

        If facts.Count = 0 Then
            Return ""
        End If

        Dim sb As New System.Text.StringBuilder()
        sb.AppendLine("<COMPLETED_FACTS>")
        For Each fact In facts
            sb.AppendLine(fact)
        Next
        sb.AppendLine("</COMPLETED_FACTS>")
        Return sb.ToString().TrimEnd()
    End Function

    Private Function BuildPostToolContinuationBlock(context As ToolExecutionContext) As String
        Return BuildDeliverableCompletionContinuationBlock(context)
    End Function



End Class