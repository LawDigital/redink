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
        Public Property ReplayRetention As SharedLibrary.Agents.ToolReplayRetentionKind
        Public Property ProducedIteration As Integer
        Public Property ControlPlanePayloadKey As String
        Public Property NormalizedCallSignature As String
        Public Property WasDuplicateReplay As Boolean

        ''' <summary>True when a tool classifies a failed result as recoverable planning/input feedback that should not count toward the generic host failure circuit breaker.</summary>
        Public Property RepairLoopRecoverable As Boolean

        ''' <summary>True when a repair-loop advisor determined the failure is terminal (budget exhausted or non-recoverable).</summary>
        Public Property RepairLoopTerminal As Boolean

        ''' <summary>Human-readable reason for <see cref="RepairLoopTerminal"/>, surfaced to abort/finalization handling.</summary>
        Public Property RepairLoopTerminalReason As String

        ''' <summary>
        ''' Host-internal verified effects produced by this concrete successful execution.
        ''' These values are never accepted from model-supplied artifact metadata.
        ''' </summary>
        Public Property VerifiedArtifactEffects As System.Collections.Generic.List(Of System.String)

        ''' <summary>
        ''' Initializes a new tool response instance with default success state.
        ''' </summary>
        Public Sub New()
            Timestamp = DateTime.Now
            Success = True
            VerifiedArtifactEffects = New System.Collections.Generic.List(Of System.String)()
        End Sub
    End Class


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
                                                       Optional compactForSubAgent As Boolean = False,
                                                       Optional currentIteration As Integer = -1) As String
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
            keepRecentFullCount:=keepRecentFullCount,
            currentIteration:=currentIteration)

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
            LogToolReplayRetentionDiagnostics(responses, currentIteration, payload.Length, budget)
            Return payload
        End If

        ' Stage 1: shrink the ordinary recent-full window. Current-turn and control-plane semantics are protected independently.
        While payload.Length > budget AndAlso keepRecentFullCount > 0
            keepRecentFullCount -= 1
            payload = BuildToolResponsesForModel(
                responses,
                toolingModel,
                compactForSubAgent:=compactForSubAgent,
                compactStaleLargeResponses:=True,
                keepRecentFullCount:=keepRecentFullCount,
                currentIteration:=currentIteration)
        End While

        If payload.Length <= budget Then
            LogToolReplayRetentionDiagnostics(responses, currentIteration, payload.Length, budget)
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
                staleCompactionPreviewChars:=previewChars,
                currentIteration:=currentIteration)
            If payload.Length <= budget Then
                Exit For
            End If
        Next

        If payload.Length > budget Then
            ToolingFileLogger.LogWarn(
                "Tool response payload still exceeds budget after progressive compaction.",
                details:=$"payloadChars={payload.Length}; budgetChars={budget}")
        End If

        LogToolReplayRetentionDiagnostics(responses, currentIteration, payload.Length, budget)
        Return payload
    End Function

    Private Sub LogToolReplayRetentionDiagnostics(responses As List(Of ToolResponse),
                                                   currentIteration As Integer,
                                                   payloadChars As Integer,
                                                   budgetChars As Integer)
        If responses Is Nothing Then Return

        For responseIndex As Integer = 0 To responses.Count - 1
            Dim resp As ToolResponse = responses(responseIndex)
            If resp Is Nothing Then Continue For

            Dim retention As SharedLibrary.Agents.ToolReplayRetentionKind =
                ResolveEffectiveReplayRetention(resp, currentIteration)

            If retention = SharedLibrary.Agents.ToolReplayRetentionKind.NormalHistorical AndAlso
               Not resp.WasCompactedForModelReplay Then
                Continue For
            End If

            Dim originalChars As Integer = If(resp.Response, "").Length
            Dim replayChars As Integer =
                If(resp.WasCompactedForModelReplay,
                   If(resp.ModelReplayContent, "").Length,
                   originalChars)
            Dim replayReference As String = ""

            If resp.WasCompactedForModelReplay AndAlso Not String.IsNullOrWhiteSpace(resp.ModelReplayContent) Then
                Try
                    Dim replayObject As JObject = JObject.Parse(resp.ModelReplayContent)
                    replayReference = If(replayObject.Value(Of String)("result_ref"), "")
                    If String.IsNullOrWhiteSpace(replayReference) Then
                        replayReference = If(replayObject.Value(Of String)("control_plane_key"), "")
                    End If
                Catch
                End Try
            End If

            ToolingFileLogger.LogDiag(
                $"Tool replay retention: index={responseIndex}; tool={If(resp.ToolName, "")}; retention={retention}; producedIteration={resp.ProducedIteration}; currentIteration={currentIteration}; compacted={If(resp.WasCompactedForModelReplay, "true", "false")}; originalChars={originalChars}; replayChars={replayChars}; replayRef={replayReference}; payloadChars={payloadChars}; budgetChars={budgetChars}")
        Next
    End Sub

    Private Function BuildToolResponseContentForModel(resp As ToolResponse,
                                                  Optional compactForSubAgent As Boolean = False,
                                                  Optional overrideThresholdChars As Integer = -1,
                                                  Optional overridePreviewChars As Integer = -1,
                                                  Optional replayRetention As SharedLibrary.Agents.ToolReplayRetentionKind = SharedLibrary.Agents.ToolReplayRetentionKind.NormalHistorical) As String
        If resp Is Nothing Then Return ""

        Dim rawContent As String

        If resp.Success Then
            rawContent = If(resp.Response, "")
        ElseIf IsStructuredErrorToolResponse(resp) Then
            rawContent = If(resp.Response, "")
        Else
            rawContent = $"Error: {If(resp.ErrorMessage, "Tool failed.")}"
        End If

        If replayRetention = SharedLibrary.Agents.ToolReplayRetentionKind.ControlPlanePinned Then
            Dim controlPlaneEnvelope As String = TryBuildControlPlaneReplayEnvelope(resp, rawContent)
            If controlPlaneEnvelope IsNot Nothing Then
                resp.ModelReplayContent = controlPlaneEnvelope
                resp.ModelReplaySummary = BuildToolReplaySummary(resp)
                resp.WasCompactedForModelReplay = True
                Return controlPlaneEnvelope
            End If

            ' Fail safe: if the host could not prove that the control payload is pinned
            ' elsewhere, preserve the full response rather than silently degrading it.
            resp.ModelReplayContent = rawContent
            resp.ModelReplaySummary = BuildToolReplaySummary(resp)
            resp.WasCompactedForModelReplay = False
            Return rawContent
        End If

        If Not compactForSubAgent Then
            ' Large deliverable/reference-bearing responses are expensive to echo on every
            ' model turn. Replay a compact envelope proactively while retaining resp.Response
            ' losslessly for artifact registration, promotion and verification.
            If rawContent.Length > SharedLibrary.Agents.ToolingConstants.DeliverableSafeReplayEnvelopeThresholdChars Then
                Dim safeEnvelope As String = TryBuildDeliverableSafeReplayEnvelope(resp, rawContent)
                If safeEnvelope IsNot Nothing Then
                    resp.ModelReplayContent = safeEnvelope
                    resp.ModelReplaySummary = BuildToolReplaySummary(resp)
                    resp.WasCompactedForModelReplay = True
                    Return safeEnvelope
                End If
            End If
            Return rawContent
        End If

        Return CompactToolResponseContentForSubAgent(resp, rawContent, overrideThresholdChars, overridePreviewChars)
    End Function

    Private Function ResolveEffectiveReplayRetention(resp As ToolResponse,
                                                     currentIteration As Integer) As SharedLibrary.Agents.ToolReplayRetentionKind
        If resp Is Nothing Then Return SharedLibrary.Agents.ToolReplayRetentionKind.NormalHistorical
        If resp.ReplayRetention = SharedLibrary.Agents.ToolReplayRetentionKind.ControlPlanePinned Then
            Return SharedLibrary.Agents.ToolReplayRetentionKind.ControlPlanePinned
        End If
        If currentIteration >= 0 AndAlso resp.ProducedIteration = currentIteration Then
            Return SharedLibrary.Agents.ToolReplayRetentionKind.CurrentTurnCritical
        End If
        Return SharedLibrary.Agents.ToolReplayRetentionKind.NormalHistorical
    End Function

    Private Function TryBuildControlPlaneReplayEnvelope(resp As ToolResponse,
                                                        rawContent As String) As String
        If resp Is Nothing OrElse String.IsNullOrWhiteSpace(resp.ControlPlanePayloadKey) Then Return Nothing

        Dim envelope As New JObject(
            New JProperty("ok", resp.Success),
            New JProperty("tool", If(resp.ToolName, "")),
            New JProperty("control_plane_pinned", True),
            New JProperty("control_plane_key", resp.ControlPlanePayloadKey),
            New JProperty("total_chars", If(rawContent, "").Length),
            New JProperty("summary", "The full authoritative control payload is pinned by the host and replayed separately on every parent turn."))

        Try
            Dim source As JObject = JObject.Parse(If(rawContent, ""))
            For Each fieldName As String In New String() {
                "name",
                "description",
                "origin",
                "network_allowed",
                "allowed_tools",
                "declared_deliverable_count",
                "declared_deliverable_required_effects",
                "declared_required_successful_tools",
                "declared_required_successful_tools_before_final_mutation"
            }
                Dim token As JToken = source(fieldName)
                If token IsNot Nothing AndAlso token.Type <> JTokenType.Null Then
                    envelope(fieldName) = token.DeepClone()
                End If
            Next
        Catch
            ' The payload itself remains pinned losslessly. Envelope enrichment is optional.
        End Try

        Return envelope.ToString(Formatting.None)
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
                    SharedLibrary.Agents.ToolingHostKind.Word)
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


    Private Sub CopyReplayFieldIfPresent(source As JObject, target As JObject, fieldName As String)
        If source Is Nothing OrElse target Is Nothing OrElse String.IsNullOrWhiteSpace(fieldName) Then Return
        Dim token As JToken = source(fieldName)
        If token Is Nothing OrElse token.Type = JTokenType.Null Then Return
        target(fieldName) = token.DeepClone()
    End Sub

    ''' <summary>
    ''' Builds a compact replay envelope for large deliverable/reference-bearing tool
    ''' results. The full raw response is stored losslessly in ToolResultStore, while
    ''' artifact identity/reference fields and mutation counts remain directly visible
    ''' to the model. Artifact registration itself always uses ToolResponse.Response and
    ''' therefore occurs independently of this replay-only compaction.
    ''' </summary>
    Private Function TryBuildDeliverableSafeReplayEnvelope(resp As ToolResponse, rawContent As String) As String
        If resp Is Nothing Then Return Nothing

        Dim raw As String = If(rawContent, "")
        If raw = "" OrElse Not MustPreserveFullResponseForReplay(resp, raw) Then Return Nothing

        ' Reuse an existing lossless replay envelope for this immutable ToolResponse.
        ' This keeps result_ref stable across iterations and avoids storing the same
        ' large result repeatedly merely because the parent prompt is rebuilt.
        If resp.WasCompactedForModelReplay AndAlso Not String.IsNullOrWhiteSpace(resp.ModelReplayContent) Then
            Try
                Dim existing As JObject = JObject.Parse(resp.ModelReplayContent)
                Dim compactedToken As JToken = existing("compacted_for_model_replay")
                If compactedToken IsNot Nothing AndAlso
                   compactedToken.Type = JTokenType.Boolean AndAlso
                   compactedToken.Value(Of Boolean)() AndAlso
                   Not String.IsNullOrWhiteSpace(existing.Value(Of String)("result_ref")) Then
                    Return resp.ModelReplayContent
                End If
            Catch
            End Try
        End If

        Dim source As JObject
        Try
            source = JObject.Parse(raw)
        Catch
            Return Nothing
        End Try

        Dim stored As SharedLibrary.Agents.ToolResultStore.StoredResult =
            SharedLibrary.Agents.ToolResultStore.Put(
                SharedLibrary.Agents.WorkflowContinuity.CurrentWorkflowId,
                If(resp.ToolName, ""),
                raw)

        Dim compact As New JObject(
            New JProperty("ok", resp.Success),
            New JProperty("tool", If(resp.ToolName, "")),
            New JProperty("summary", BuildToolReplaySummary(resp)),
            New JProperty("result_ref", stored.Ref),
            New JProperty("total_chars", raw.Length),
            New JProperty("compacted_for_model_replay", True))

        Dim replayFields As String() = {
            "artifact_id",
            "logical_deliverable_id",
            "output_slot_id",
            "artifact_state",
            "storage_kind",
            "output_reference",
            "output_file",
            "output_filename",
            "output_path",
            "saved_path",
            "path",
            "attachment_name",
            "memory_key",
            "reference",
            "applied_update_count",
            "content_mutation_update_count",
            "failed_update_count",
            "skipped_non_writable_count",
            "partial_success",
            "write_blocked_by_protection",
            "edited_in_place",
            "mutated_worksheets"
        }

        For Each fieldName As String In replayFields
            CopyReplayFieldIfPresent(source, compact, fieldName)
        Next

        ' Preserve generic structured/output references even when the producing tool
        ' nests them instead of exposing a conventional top-level property.
        Try
            Dim structuredRef As String = SharedLibrary.Agents.WorkflowContinuity.ExtractStructuredResultReference(raw)
            If Not String.IsNullOrWhiteSpace(structuredRef) AndAlso compact("structured_result_reference") Is Nothing Then
                compact("structured_result_reference") = structuredRef
            End If

            Dim outputRef As String = SharedLibrary.Agents.WorkflowContinuity.ExtractOutputReference(raw)
            If Not String.IsNullOrWhiteSpace(outputRef) AndAlso compact("output_reference") Is Nothing Then
                compact("output_reference") = outputRef
            End If
        Catch
        End Try

        Dim issues As JArray = TryCast(source("issues"), JArray)
        If issues IsNot Nothing Then
            compact("issue_count") = issues.Count

            Dim nonAppliedIssues As New JArray()
            For Each issueToken As JToken In issues
                Dim issueObject As JObject = TryCast(issueToken, JObject)
                If issueObject Is Nothing Then Continue For

                Dim status As String = If(issueObject.Value(Of String)("status"), "").Trim()
                If String.Equals(status, "applied", StringComparison.OrdinalIgnoreCase) Then Continue For

                nonAppliedIssues.Add(issueObject.DeepClone())
                If nonAppliedIssues.Count >= 5 Then Exit For
            Next

            If nonAppliedIssues.Count > 0 Then
                compact("issues") = nonAppliedIssues
                If nonAppliedIssues.Count < issues.Count Then
                    compact("issues_truncated") = True
                End If
            End If
        End If

        compact("continuation") =
            "The full tool result is stored losslessly. Use context_expand with result_ref only if additional detail is required."

        Return compact.ToString(Formatting.None)
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

        Dim deliverableSafeEnvelope As String = TryBuildDeliverableSafeReplayEnvelope(resp, raw)
        If deliverableSafeEnvelope IsNot Nothing Then
            resp.ModelReplayContent = deliverableSafeEnvelope
            resp.ModelReplaySummary = BuildToolReplaySummary(resp)
            resp.WasCompactedForModelReplay = True
            Return deliverableSafeEnvelope
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
