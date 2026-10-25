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

        Public Property PreservedResultRef As System.String
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
        ''' True only when the concrete tool implementation explicitly declares that a materially
        ''' different successful path may supersede this failed step even when opaque operation ids
        ''' differ. The shared runtime still requires substantive success and, for deliverable tasks,
        ''' a host-validated deliverable before clearing the prior failure.
        ''' </summary>
        Public Property AllowCrossScopeAlternativeRecovery As Boolean

        ''' <summary>Optional artifact extension (for example .docx) that a replacement deliverable must satisfy before cross-scope recovery can clear this failure.</summary>
        Public Property AlternativeRecoveryRequiredArtifactExtension As System.String

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
    ''' pass keeps current-turn results and recent history fully visible. Only when the overall
    ''' payload grows beyond the budget does it progressively shrink historical replay first;
    ''' only a final budget stage may reference-compact current-turn results. Everything moved
    ''' this way stays fully retrievable via context_expand, so compaction is lossless. The
    ''' model can also voluntarily tighten this via the context_compact tool.
    ''' Capability-driven: no tool-name or content-type heuristics.
    ''' </summary>
    Public Function BuildToolResponsesForModelBudgeted(responses As List(Of ToolResponse),
                                                       toolingModel As ModelConfig,
                                                       Optional compactForSubAgent As Boolean = False,
                                                       Optional currentIteration As Integer = -1) As String
        Dim keepRecentFullCount As System.Int32 = System.Math.Max(0, INI_ToolResponseKeepRecentFullCount)

        Dim requestedKeep As Integer
        Dim requestedCompaction As System.Boolean = SharedLibrary.Agents.ToolResultStore.TryTakeRequestedKeepRecent(
                SharedLibrary.Agents.WorkflowContinuity.CurrentWorkflowId, requestedKeep)
        If requestedCompaction Then
            keepRecentFullCount = requestedKeep
        End If

        ' A configured protection count is an upper bound, not a retry-loop count.
        ' Clamp to actual historical responses before the progressive budget passes.
        Dim historicalResponseCount As System.Int32 = 0
        If responses IsNot Nothing Then
            For Each historicalResponse As ToolResponse In responses
                If historicalResponse IsNot Nothing AndAlso
                   ResolveEffectiveReplayRetention(historicalResponse, currentIteration) =
                   SharedLibrary.Agents.ToolReplayRetentionKind.NormalHistorical Then historicalResponseCount += 1
            Next
        End If
        keepRecentFullCount = System.Math.Min(keepRecentFullCount, historicalResponseCount)

        Dim recoveryRequested As System.Boolean = SharedLibrary.Agents.ToolResultStore.IsContextRecoveryRequested(
            SharedLibrary.Agents.WorkflowContinuity.CurrentWorkflowId)

        Dim payload As String = BuildToolResponsesForModel(
            responses,
            toolingModel,
            compactForSubAgent:=False,
            compactStaleLargeResponses:=requestedCompaction,
            keepRecentFullCount:=keepRecentFullCount,
            currentIteration:=currentIteration,
            allowCurrentTurnStructuredCompaction:=False,
            allowCurrentTurnReferenceCompaction:=False)

        Dim budget As Integer =
            If(ThisAddIn.INI_ToolResponsePayloadBudgetChars >= 0,
               ThisAddIn.INI_ToolResponsePayloadBudgetChars,
               SharedLibrary.Agents.ToolingConstants.ToolResponsePayloadBudgetChars)
        Dim budgetPolicyActive As System.Boolean = requestedCompaction OrElse recoveryRequested OrElse
            (Not INI_ToolResponseCompactOnlyAfterEmpty AndAlso budget > 0)
        If budget <= 0 AndAlso (requestedCompaction OrElse recoveryRequested) Then
            budget = INI_ToolResponseRecoveryBudgetChars
        End If
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

        If Not budgetPolicyActive OrElse budget <= 0 OrElse String.IsNullOrEmpty(payload) OrElse payload.Length <= budget Then
            LogToolReplayRetentionDiagnostics(responses, currentIteration, payload.Length, budget)
            Return payload
        End If

        ' Stage 0: before sacrificing historical evidence, compact newly produced large
        ' structured results to deterministic evidence cores. Exact retained scalar values
        ' remain visible while the full original stays available by result_ref.
        payload = BuildToolResponsesForModel(
            responses,
            toolingModel,
            compactForSubAgent:=False,
            compactStaleLargeResponses:=True,
            keepRecentFullCount:=keepRecentFullCount,
            currentIteration:=currentIteration,
            allowCurrentTurnStructuredCompaction:=True,
            allowCurrentTurnReferenceCompaction:=False)

        If payload.Length <= budget Then
            LogToolReplayRetentionDiagnostics(responses, currentIteration, payload.Length, budget)
            Return payload
        End If

        ' Stage 1: shrink only the historical recent-full window. Current-turn results may
        ' continue using evidence-core replay, and current results never displace historical
        ' responses merely because they were emitted in the same new batch.
        While payload.Length > budget AndAlso keepRecentFullCount > 0
            keepRecentFullCount -= 1
            payload = BuildToolResponsesForModel(
                responses,
                toolingModel,
                compactForSubAgent:=False,
                compactStaleLargeResponses:=True,
                keepRecentFullCount:=keepRecentFullCount,
                currentIteration:=currentIteration,
                allowCurrentTurnStructuredCompaction:=True,
                allowCurrentTurnReferenceCompaction:=False)
        End While

        If payload.Length <= budget Then
            LogToolReplayRetentionDiagnostics(responses, currentIteration, payload.Length, budget)
            Return payload
        End If

        ' Stage 2: progressively compact historical medium-sized results. Structured JSON
        ' is reduced to an exact evidence core before falling back to preview-only replay.
        For Each mediumThresholdChars As Integer In New Integer() {mediumThreshold, aggressiveThreshold}
            payload = BuildToolResponsesForModel(
                responses,
                toolingModel,
                compactForSubAgent:=False,
                compactStaleLargeResponses:=True,
                keepRecentFullCount:=0,
                staleCompactionThresholdChars:=mediumThresholdChars,
                staleCompactionPreviewChars:=previewChars,
                currentIteration:=currentIteration,
                allowCurrentTurnStructuredCompaction:=True,
                allowCurrentTurnReferenceCompaction:=False)
            If payload.Length <= budget Then Exit For
        Next

        ' Stage 3: only after evidence-preserving compaction is exhausted may opaque
        ' current-turn results fall back to reference/preview compaction. context_expand
        ' itself remains lossless on its current turn.
        If payload.Length > budget Then
            payload = BuildToolResponsesForModel(
                responses,
                toolingModel,
                compactForSubAgent:=False,
                compactStaleLargeResponses:=True,
                keepRecentFullCount:=0,
                staleCompactionThresholdChars:=aggressiveThreshold,
                staleCompactionPreviewChars:=previewChars,
                currentIteration:=currentIteration,
                allowCurrentTurnStructuredCompaction:=True,
                allowCurrentTurnReferenceCompaction:=True)
        End If

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
            Dim replayMode As String = If(resp.WasCompactedForModelReplay, "compacted", "full")

            If resp.WasCompactedForModelReplay AndAlso Not String.IsNullOrWhiteSpace(resp.ModelReplayContent) Then
                Try
                    Dim replayObject As JObject = JObject.Parse(resp.ModelReplayContent)
                    replayReference = If(replayObject.Value(Of String)("result_ref"), "")
                    If String.IsNullOrWhiteSpace(replayReference) Then
                        replayReference = If(replayObject.Value(Of String)("control_plane_key"), "")
                    End If

                    If replayObject("evidence_core") IsNot Nothing Then
                        replayMode = "evidence_core"
                    ElseIf replayObject("preview") IsNot Nothing Then
                        replayMode = "preview"
                    ElseIf replayObject("omitted_window_chars") IsNot Nothing Then
                        replayMode = "window_stub"
                    ElseIf Not String.IsNullOrWhiteSpace(replayObject.Value(Of String)("control_plane_key")) Then
                        replayMode = "control_plane"
                    End If
                Catch
                    replayMode = "compacted_unclassified"
                End Try
            End If

            ToolingFileLogger.LogDiag(
                $"Tool replay retention: index={responseIndex}; tool={If(resp.ToolName, "")}; retention={retention}; producedIteration={resp.ProducedIteration}; currentIteration={currentIteration}; compacted={If(resp.WasCompactedForModelReplay, "true", "false")}; replayMode={replayMode}; originalChars={originalChars}; replayChars={replayChars}; replayRef={replayReference}; payloadChars={payloadChars}; budgetChars={budgetChars}")
        Next
    End Sub

    Private Function BuildToolResponseContentForModel(resp As ToolResponse,
                                                  Optional compactForSubAgent As Boolean = False,
                                                  Optional overrideThresholdChars As Integer = -1,
                                                  Optional overridePreviewChars As Integer = -1,
                                                  Optional replayRetention As SharedLibrary.Agents.ToolReplayRetentionKind = SharedLibrary.Agents.ToolReplayRetentionKind.NormalHistorical,
                                                  Optional allowCurrentTurnStructuredCompaction As System.Boolean = False,
                                                  Optional allowCurrentTurnReferenceCompaction As System.Boolean = False) As String
        If resp Is Nothing Then Return ""

        Dim rawContent As String

        If resp.Success Then
            rawContent = If(resp.Response, "")
        ElseIf IsStructuredErrorToolResponse(resp) Then
            rawContent = If(resp.Response, "")
        ElseIf Not System.String.IsNullOrEmpty(resp.Response) Then
            ' Failed tools may contain completed work or partial source evidence.
            ' Keep the complete body and label the failed execution explicitly.
            rawContent = New Newtonsoft.Json.Linq.JObject(
                New Newtonsoft.Json.Linq.JProperty("ok", False),
                New Newtonsoft.Json.Linq.JProperty("execution_succeeded", False),
                New Newtonsoft.Json.Linq.JProperty("error", If(resp.ErrorMessage, "Tool failed.")),
                New Newtonsoft.Json.Linq.JProperty("raw_result", resp.Response),
                New Newtonsoft.Json.Linq.JProperty("result_contract", "Unabridged output of a failed tool. Preserve useful evidence, but do not infer successful completion from this body.")
            ).ToString(Newtonsoft.Json.Formatting.None)
        Else
            rawContent = $"Error: {If(resp.ErrorMessage, "Tool failed.")}"
        End If

        If replayRetention = SharedLibrary.Agents.ToolReplayRetentionKind.CurrentTurnCritical AndAlso
           SharedLibrary.Agents.ToolingRuntimePrimitives.RequiresLosslessCurrentTurnReplay(resp.ToolName) Then

            ' A model-requested expansion must reach the model losslessly at least once.
            ' Historical compaction may replace it with a navigation stub on later turns.
            resp.ModelReplayContent = rawContent
            resp.ModelReplaySummary = BuildToolReplaySummary(resp)
            resp.WasCompactedForModelReplay = False
            Return rawContent
        End If

        If replayRetention = SharedLibrary.Agents.ToolReplayRetentionKind.NormalHistorical AndAlso
           Not compactForSubAgent AndAlso resp.WasCompactedForModelReplay AndAlso
           SharedLibrary.Agents.ContextProgressPolicy.CanRetainHistoricalReplay(resp.ModelReplayContent) Then
            Return resp.ModelReplayContent
        End If

        If resp.Success AndAlso
           replayRetention <> SharedLibrary.Agents.ToolReplayRetentionKind.ControlPlanePinned AndAlso
           INI_ToolResultReferenceThresholdChars > 0 AndAlso rawContent.Length > INI_ToolResultReferenceThresholdChars Then
            Dim sourceReference As System.String = TryBuildLargeResultReferenceReplay(resp, rawContent)
            If sourceReference IsNot Nothing Then
                resp.ModelReplayContent = sourceReference
                resp.ModelReplaySummary = BuildToolReplaySummary(resp)
                resp.WasCompactedForModelReplay = True
                Return sourceReference
            End If
        End If

        If replayRetention = SharedLibrary.Agents.ToolReplayRetentionKind.CurrentTurnCritical AndAlso
           allowCurrentTurnStructuredCompaction Then

            ' Under budget pressure, a newly produced structured result may be replayed as
            ' an exact evidence core before older source evidence is sacrificed. Values that
            ' remain in evidence_core are verbatim; omitted fields stay retrievable by ref.
            Dim evidenceEnvelope As String = TryBuildEvidenceReplayEnvelope(resp, rawContent)
            If evidenceEnvelope IsNot Nothing AndAlso evidenceEnvelope.Length < rawContent.Length Then
                resp.ModelReplayContent = evidenceEnvelope
                resp.ModelReplaySummary = BuildToolReplaySummary(resp)
                resp.WasCompactedForModelReplay = True
                Return evidenceEnvelope
            End If
        End If

        If replayRetention = SharedLibrary.Agents.ToolReplayRetentionKind.CurrentTurnCritical AndAlso
           Not allowCurrentTurnReferenceCompaction Then

            ' A current-turn result that cannot be reduced safely to a structured evidence
            ' core remains full until historical/evidence-preserving compaction is exhausted.
            resp.ModelReplayContent = rawContent
            resp.ModelReplaySummary = BuildToolReplaySummary(resp)
            resp.WasCompactedForModelReplay = False
            Return rawContent
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
            ' Keep replay-state metadata synchronized with what the model actually receives.
            ' A response that is full again must not remain marked as compacted from an
            ' earlier budget pass, otherwise final evidence review would act on stale state.
            resp.ModelReplayContent = rawContent
            resp.ModelReplaySummary = BuildToolReplaySummary(resp)
            resp.WasCompactedForModelReplay = False
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
        Dim pinned As SharedLibrary.Agents.ToolingControlPlaneStore.PinnedPayload = Nothing
        If Not SharedLibrary.Agents.ToolingControlPlaneStore.TryGetActiveSkillPayload(
            SharedLibrary.Agents.WorkflowContinuity.CurrentWorkflowId, pinned) OrElse pinned Is Nothing OrElse
           Not System.String.Equals(pinned.Key, resp.ControlPlanePayloadKey, System.StringComparison.Ordinal) OrElse
           Not System.String.Equals(pinned.FullPayload, If(rawContent, ""), System.StringComparison.Ordinal) Then
            Return Nothing
        End If

        Dim envelope As New JObject(
            New JProperty("ok", resp.Success),
            New JProperty("tool", If(resp.ToolName, "")),
            New JProperty("control_plane_pinned", True),
            New JProperty("content_complete", False),
            New JProperty("representation", "control_plane_stub"),
            New JProperty("is_semantic_summary", False),
            New JProperty("retrieval_required", "This is an incomplete pointer. Use the full authoritative control payload replayed separately by the host on this turn; do not infer missing instructions from this stub."),
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

        Dim referenced As SharedLibrary.Agents.ToolResultStore.StoredResult = Nothing
        If Not SharedLibrary.Agents.ToolResultStore.TryGetForWorkflow(
            refValue, SharedLibrary.Agents.WorkflowContinuity.CurrentWorkflowId, referenced) OrElse
           Not referenced.IsDurable Then Return Nothing

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

        Return SharedLibrary.Agents.ToolResultStore.AnnotateReferenceEnvelope(stub).ToString(Formatting.None)
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
                    Dim reusable As SharedLibrary.Agents.ToolResultStore.StoredResult = Nothing
                    If TryResolveReusableReplayStoredResult(resp, raw, reusable) AndAlso reusable.IsDurable Then
                        Return SharedLibrary.Agents.ToolResultStore.AnnotateReferenceEnvelope(existing).ToString(Formatting.None)
                    End If
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

        Dim stored As SharedLibrary.Agents.ToolResultStore.StoredResult = Nothing
        If Not TryResolveReusableReplayStoredResult(resp, raw, stored) Then
            stored = SharedLibrary.Agents.ToolResultStore.Put(
                SharedLibrary.Agents.WorkflowContinuity.CurrentWorkflowId, If(resp.ToolName, ""), raw)
        End If
        resp.PreservedResultRef = stored.Ref
        If Not stored.IsDurable Then Return Nothing

        Dim compact As New JObject(
            New JProperty("ok", resp.Success),
            New JProperty("tool", If(resp.ToolName, "")),
            New JProperty("summary", BuildToolReplaySummary(resp)),
            New JProperty("result_ref", stored.Ref),
            New JProperty("total_chars", raw.Length),
            New JProperty("compacted_for_model_replay", True))

        Dim replayFields As String() = {
            "resource_id",
            "extraction_complete",
            "processed_ranges",
            "coverage",
            "source_hash",
            "source_sha256",
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

        If source("result_ref") IsNot Nothing Then compact("source_result_ref") = source("result_ref").DeepClone()
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
            "This envelope omits tool-result content. Retrieve every omitted part needed for the task through context_expand before using it; whole-source conclusions require complete coverage."

        Return SharedLibrary.Agents.ToolResultStore.AnnotateReferenceEnvelope(compact).ToString(Formatting.None)
    End Function

    ''' <summary>
    ''' Builds a compact evidence-preserving replay for structured JSON. The full result is
    ''' stored losslessly by reference; every value retained in evidence_core is copied exactly.
    ''' </summary>
    Private Function TryBuildEvidenceReplayEnvelope(resp As ToolResponse, rawContent As System.String) As System.String
        If resp Is Nothing OrElse Not resp.Success Then Return Nothing

        Dim raw As System.String = If(rawContent, System.String.Empty)
        If raw = System.String.Empty Then Return Nothing
        If MustPreserveFullResponseForReplay(resp, raw) Then Return Nothing

        Dim evidenceCoreJson As System.String = System.String.Empty
        Dim omittedPropertyCount As System.Int32 = 0
        Dim exactScalarCount As System.Int32 = 0
        Dim arrayItemPropertyLimit As System.Int32 = 0

        If Not SharedLibrary.Agents.ToolingRuntimePrimitives.TryBuildEvidenceCore(
                raw,
                INI_EvidenceReplayCoreMaxChars,
                evidenceCoreJson,
                omittedPropertyCount,
                exactScalarCount,
                arrayItemPropertyLimit,
                INI_EvidenceReplayPreferredArrayItemProperties,
                INI_EvidenceReplayMaxExactStringChars) Then
            Return Nothing
        End If

        Dim stored As SharedLibrary.Agents.ToolResultStore.StoredResult = Nothing
        If Not TryResolveReusableReplayStoredResult(resp, raw, stored) Then
            stored = SharedLibrary.Agents.ToolResultStore.Put(
                SharedLibrary.Agents.WorkflowContinuity.CurrentWorkflowId,
                If(resp.ToolName, System.String.Empty),
                raw)
        End If

        resp.PreservedResultRef = stored.Ref
        If Not stored.IsDurable Then Return Nothing

        Dim evidenceCore As JToken
        Try
            evidenceCore = JToken.Parse(evidenceCoreJson)
        Catch ex As System.Exception
            Return Nothing
        End Try

        Dim compact As New JObject(
            New JProperty("ok", True),
            New JProperty("tool", If(resp.ToolName, System.String.Empty)),
            New JProperty("summary", BuildToolReplaySummary(resp)),
            New JProperty("result_ref", stored.Ref),
            New JProperty("total_chars", raw.Length),
            New JProperty("compacted_for_model_replay", True),
            New JProperty("evidence_core_exact_values", True),
            New JProperty("evidence_core_omitted_property_count", omittedPropertyCount),
            New JProperty("evidence_core_exact_scalar_count", exactScalarCount),
            New JProperty("evidence_core_array_item_property_limit", arrayItemPropertyLimit),
            New JProperty("evidence_core", evidenceCore),
            New JProperty("evidence_contract",
                "Every scalar value present in evidence_core is verbatim from the original tool result and is authoritative for that value. Omitted properties/text are not evidence. If a required fact is omitted, call context_expand with result_ref before relying on it. Never reconstruct or alter exact factual values."),
            New JProperty("continuation",
                "The full original result is stored losslessly. Use context_expand with result_ref only for facts or text omitted from evidence_core."))

        Return SharedLibrary.Agents.ToolResultStore.AnnotateReferenceEnvelope(compact).ToString(Formatting.None)
    End Function

    Private Function TryResolveReusableReplayStoredResult(resp As ToolResponse,
                                                           rawContent As System.String,
                                                           ByRef stored As SharedLibrary.Agents.ToolResultStore.StoredResult) As System.Boolean
        stored = Nothing
        If resp Is Nothing Then Return False
        If SharedLibrary.Agents.ToolResultStore.TryReuseReference(
            resp.PreservedResultRef, SharedLibrary.Agents.WorkflowContinuity.CurrentWorkflowId,
            If(resp.ToolName, System.String.Empty), If(rawContent, System.String.Empty), stored) Then Return True
        If Not resp.WasCompactedForModelReplay OrElse
           System.String.IsNullOrWhiteSpace(resp.ModelReplayContent) Then
            Return False
        End If

        Dim existingRef As System.String = System.String.Empty
        Try
            Dim existingEnvelope As Newtonsoft.Json.Linq.JObject =
                Newtonsoft.Json.Linq.JObject.Parse(resp.ModelReplayContent)
            existingRef = If(existingEnvelope.Value(Of System.String)("result_ref"), System.String.Empty).Trim()
        Catch ex As System.Exception
            Return False
        End Try

        If existingRef = System.String.Empty Then Return False

        Return SharedLibrary.Agents.ToolResultStore.TryReuseReference(
            existingRef,
            SharedLibrary.Agents.WorkflowContinuity.CurrentWorkflowId,
            If(resp.ToolName, System.String.Empty),
            If(rawContent, System.String.Empty),
            stored)
    End Function

    Private Function TryBuildLargeResultReferenceReplay(resp As ToolResponse, rawContent As System.String) As System.String
        If MustPreserveFullResponseForReplay(resp, rawContent) Then
            Dim safeEnvelope As System.String = TryBuildDeliverableSafeReplayEnvelope(resp, rawContent)
            Return If(safeEnvelope IsNot Nothing AndAlso safeEnvelope.Length < rawContent.Length, safeEnvelope, Nothing)
        End If
        Dim stored As SharedLibrary.Agents.ToolResultStore.StoredResult = Nothing
        If Not TryResolveReusableReplayStoredResult(resp, rawContent, stored) Then
            stored = SharedLibrary.Agents.ToolResultStore.Put(
                SharedLibrary.Agents.WorkflowContinuity.CurrentWorkflowId, resp.ToolName, rawContent)
        End If
        resp.PreservedResultRef = stored.Ref
        If Not stored.IsDurable Then Return Nothing
        Dim previewLength As System.Int32 = SharedLibrary.Agents.ToolResultStore.GetUnicodeSafePrefixLength(rawContent, INI_ToolResultReferencePreviewChars)
        Dim envelope As Newtonsoft.Json.Linq.JObject = SharedLibrary.Agents.ToolResultStore.DescribeReference(stored)
        envelope("tool") = resp.ToolName
        envelope("preview") = rawContent.Substring(0, previewLength)
        envelope("returned_chars") = previewLength
        envelope("next_offset") = previewLength
        envelope("truncated") = previewLength < rawContent.Length
        envelope("context_handling_reason") = "Large-result reference replay, not semantic compaction. Prefer a complete source file and isolated file analysis; otherwise load every task-required range through context_expand."
        Dim serialized As System.String = envelope.ToString(Newtonsoft.Json.Formatting.None)
        Return If(serialized.Length < rawContent.Length, serialized, Nothing)
    End Function

    Private Function CompactToolResponseContentForSubAgent(resp As ToolResponse, rawContent As String,
                                                           Optional overrideThresholdChars As Integer = -1,
                                                           Optional overridePreviewChars As Integer = -1) As String
        If resp Is Nothing Then Return If(rawContent, "")

        Dim raw As String = If(rawContent, "")

        Dim thresholdChars As Integer =
            If(overrideThresholdChars > 0, overrideThresholdChars, INI_SubAgentLargeToolResponseThresholdChars)
        Dim previewChars As Integer =
            If(overridePreviewChars >= 0, overridePreviewChars, INI_SubAgentLargeToolResponseExcerptChars)

        Dim windowStub As String = TryBuildReferencedWindowStub(raw)
        If windowStub IsNot Nothing AndAlso windowStub.Length < raw.Length Then
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
        If deliverableSafeEnvelope IsNot Nothing AndAlso deliverableSafeEnvelope.Length < raw.Length Then
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

        ' Prefer deterministic structured evidence replay over a positional text preview.
        ' This preserves exact scalar facts across all array items whenever the result is JSON.
        Dim evidenceEnvelope As String = TryBuildEvidenceReplayEnvelope(resp, raw)
        If evidenceEnvelope IsNot Nothing AndAlso evidenceEnvelope.Length < raw.Length Then
            resp.ModelReplayContent = evidenceEnvelope
            resp.ModelReplaySummary = BuildToolReplaySummary(resp)
            resp.WasCompactedForModelReplay = True
            Return evidenceEnvelope
        End If

        Dim excerptLength As System.Int32 = SharedLibrary.Agents.ToolResultStore.GetUnicodeSafePrefixLength(raw, previewChars)
        Dim excerpt As String = raw.Substring(0, excerptLength)
        Dim summary As String = BuildToolReplaySummary(resp)

        Dim stored As SharedLibrary.Agents.ToolResultStore.StoredResult = Nothing
        If Not TryResolveReusableReplayStoredResult(resp, raw, stored) Then
            stored = SharedLibrary.Agents.ToolResultStore.Put(
                SharedLibrary.Agents.WorkflowContinuity.CurrentWorkflowId,
                If(resp.ToolName, ""),
                raw)
        End If

        resp.PreservedResultRef = stored.Ref
        If Not stored.IsDurable Then
            resp.ModelReplayContent = raw
            resp.WasCompactedForModelReplay = False
            Return raw
        End If

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

        Dim serializedPreview As System.String = SharedLibrary.Agents.ToolResultStore.AnnotateReferenceEnvelope(compactObj).ToString(Formatting.None)
        If serializedPreview.Length >= raw.Length Then
            resp.ModelReplayContent = raw
            resp.WasCompactedForModelReplay = False
            Return raw
        End If
        resp.ModelReplayContent = serializedPreview
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
            summary = $"{If(resp.ToolName, "tool")} {If(resp.Success, "succeeded", "failed")}. Incomplete display excerpt: {BuildResultExcerpt(If(resp.Response, ""), 280)}"
        End If

        resp.ModelReplaySummary = summary
        Return summary
    End Function

    Private Function PreserveSuccessfulToolResults(context As ToolExecutionContext) As Newtonsoft.Json.Linq.JArray
        Dim results As New Newtonsoft.Json.Linq.JArray()
        If context Is Nothing OrElse context.AllToolResponses Is Nothing Then Return results
        For Each response As ToolResponse In context.AllToolResponses
            If response Is Nothing Then Continue For
            Dim retainedBody As System.String = If(response.Response, "")
            If Not response.Success Then
                retainedBody = New Newtonsoft.Json.Linq.JObject(
                    New Newtonsoft.Json.Linq.JProperty("execution_succeeded", False),
                    New Newtonsoft.Json.Linq.JProperty("raw_result", retainedBody),
                    New Newtonsoft.Json.Linq.JProperty("error", If(response.ErrorMessage, "Tool failed."))
                ).ToString(Newtonsoft.Json.Formatting.None)
            End If
            Dim stored As SharedLibrary.Agents.ToolResultStore.StoredResult = Nothing
            If Not SharedLibrary.Agents.ToolResultStore.TryReuseReference(
                response.PreservedResultRef, context.WorkflowId, response.ToolName, retainedBody, stored) OrElse Not stored.IsDurable Then
                stored = SharedLibrary.Agents.ToolResultStore.Put(context.WorkflowId, response.ToolName, retainedBody)
                response.PreservedResultRef = stored.Ref
            End If
            ' Failed outputs stay in the continuation archive, never in the success manifest.
            If Not response.Success Then Continue For
            Dim item As Newtonsoft.Json.Linq.JObject = SharedLibrary.Agents.ToolResultStore.DescribeReference(stored)
            item("tool") = response.ToolName
            item("call_id") = response.CallId
            item("execution_succeeded") = True
            If Not stored.IsDurable Then
                item("content_complete") = True
                item("representation") = "full_tool_result"
                item("content") = response.Response
            End If
            results.Add(item)
        Next
        Return results
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

    ''' <summary>
    ''' Schedules one bounded evidence-grounding review after the latest set of tool
    ''' responses whenever source evidence is currently compacted by reference.
    ''' </summary>
    Private Function TryScheduleEvidenceFinalizationReview(context As ToolExecutionContext,
                                                           proposedFinalResponse As System.String) As System.Boolean
        If context Is Nothing OrElse context.AllToolResponses Is Nothing Then Return False

        Dim responseCount As System.Int32 = context.AllToolResponses.Count
        If context.EvidenceFinalizationReviewResponseCount = responseCount Then Return False

        Dim compactedEvidenceCount As System.Int32 = 0
        For Each response As ToolResponse In context.AllToolResponses
            If response Is Nothing OrElse Not response.Success Then Continue For
            Dim visibleResult As System.String = If(System.String.IsNullOrEmpty(response.ModelReplayContent), response.Response, response.ModelReplayContent)
            If SharedLibrary.Agents.ToolingRuntimePrimitives.IsEvidenceCompactedReplay(visibleResult) Then
                compactedEvidenceCount += 1
            End If
        Next

        If compactedEvidenceCount = 0 Then Return False

        context.EvidenceFinalizationReviewResponseCount = responseCount
        context.PendingContinuationGuardPrompt = SharedLibrary.Agents.ToolCallSequencing.EvidenceFinalizationReviewInstruction
        context.PendingGuardTitle = "HOST EVIDENCE FINALIZATION REVIEW"
        context.PendingRejectedTurnExplanation =
            $"The draft final response must be re-checked against {compactedEvidenceCount} compacted evidence result(s) before acceptance."
        context.PendingRejectedAssistantTurn = If(proposedFinalResponse, System.String.Empty)
        context.LogWarn(
            "Final response deferred for evidence-grounding review.",
            details:=$"compactedEvidenceResults={compactedEvidenceCount}; toolResponseCount={responseCount}")
        Return True
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




    Private Function TryExtractToolServiceErrorMessage(rawResponse As System.String, ByRef errorMessage As System.String) As System.Boolean
        errorMessage = System.String.Empty
        If System.String.IsNullOrWhiteSpace(rawResponse) Then Return False
        Dim errorCode As System.String = System.String.Empty
        Dim resultKind As System.String = System.String.Empty
        If SharedLibrary.Agents.SubAgentRuntimeHardening.TryGetEnvelopeErrorInfo(rawResponse, errorCode, resultKind) Then
            If Not SharedLibrary.Agents.SubAgentRuntimeHardening.TryGetEnvelopeErrorMessage(rawResponse, errorMessage) Then
                errorMessage = "Tool service reported failure: " & errorCode
            End If
            Return True
        End If
        Try
            Dim root As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(rawResponse)
            Dim result As Newtonsoft.Json.Linq.JObject = TryCast(root("result"), Newtonsoft.Json.Linq.JObject)
            If root("isError") IsNot Nothing OrElse result Is Nothing Then result = root
            Dim isError As System.Boolean = False
            Dim isErrorToken As Newtonsoft.Json.Linq.JToken = result("isError")
            If isErrorToken Is Nothing OrElse Not System.Boolean.TryParse(isErrorToken.ToString().Trim(), isError) OrElse Not isError Then Return False
            ' Recognition is authoritative even when the optional human-readable content is malformed.
            errorMessage = "Tool service returned an error."
            Dim messages As New System.Collections.Generic.List(Of System.String)()
            Dim content As Newtonsoft.Json.Linq.JArray = TryCast(result("content"), Newtonsoft.Json.Linq.JArray)
            If content IsNot Nothing Then
                For Each item As Newtonsoft.Json.Linq.JToken In content
                    Dim entry As Newtonsoft.Json.Linq.JObject = TryCast(item, Newtonsoft.Json.Linq.JObject)
                    Dim textToken As Newtonsoft.Json.Linq.JToken = If(entry Is Nothing, Nothing, entry("text"))
                    If textToken IsNot Nothing AndAlso textToken.Type = Newtonsoft.Json.Linq.JTokenType.String Then
                        Dim message As System.String = textToken.Value(Of System.String)().Trim()
                        If message <> System.String.Empty Then messages.Add(message)
                    End If
                Next
            End If
            If messages.Count > 0 Then errorMessage = System.String.Join(" ", messages)
            Return True
        Catch ex As System.Exception
            Return errorMessage <> System.String.Empty
        End Try
    End Function



End Class
