' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ToolCallSequencing.vb
' Purpose: Validates tool call sequences and final turn acceptance:
'           - Blocks dependent batches (ensure tool call ordering).
'           - Enforces <TASK_STATUS> footer contract (Q13).
'           - Guards against action promises without invocation (Q10).
'           - Manages memory grounding modes (required/optional/none).
'           - Carries bootstrap-classified source-format authority for deterministic host validation.
'           - Detects unresolved tool failures and orchestrates repair prompts.
'
' Architecture:
'  - Validates ActiveToolingTurn sequences (tool calls vs. finals).
'  - TaskStatusKind: Complete, Blocked, ContinueTurn, or Missing.
'  - MemoryGroundingMode: None, OptionalMode, Required.
'  - MemoryGroundingStage: progression from ListRequired through FullMemoryAvailable.
' =============================================================================

Option Strict On
Option Explicit On


Imports System.Collections
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Agents

    Public NotInheritable Class ToolCallSequencing


        Public Shared Function TryPrepareToolCallArgumentsForPreflight(
            toolConfig As SharedLibrary.ModelConfig,
            arguments As System.Collections.Generic.Dictionary(Of System.String, System.Object),
            ByRef preparedArguments As System.Collections.Generic.Dictionary(Of System.String, System.Object),
            ByRef failureMessage As System.String) As System.Boolean

            preparedArguments = arguments
            failureMessage = System.String.Empty

            If toolConfig Is Nothing OrElse toolConfig.ToolCallArgumentNormalizer Is Nothing Then
                Return True
            End If

            Try
                preparedArguments = toolConfig.ToolCallArgumentNormalizer.Invoke(arguments)
            Catch ex As System.Exception
                preparedArguments = arguments
                failureMessage = "Tool arguments could not be normalized for validation."
                Return False
            End Try

            If preparedArguments Is Nothing Then
                preparedArguments = arguments
                failureMessage = "Tool argument normalization returned no argument object."
                Return False
            End If

            Return True
        End Function

        Public Const TaskStatusReasonMaxChars As Integer = 160

        Private Sub New()
        End Sub

        Public Const DependentBatchingInstruction As String =
            "When using tools, the host executes multiple tool calls from one response strictly in the order emitted." & vbCrLf &
            "If several independent calls are already known and none requires inspecting another call's result, prefer emitting them together in the same response instead of spending separate model turns on each call. This is especially appropriate for independent read-only lookups." & vbCrLf &
            "This reduces model round-trips only; tool execution remains ordered and you must not assume parallel execution. Do not batch mutations merely for speed." & vbCrLf &
            "Only emit multiple tool calls in one response when every later call's arguments are already fully known at emission time." & vbCrLf &
            "If a later step depends on inspecting the result of an earlier tool call, emit only the earlier call and wait for its result before deciding the next call." & vbCrLf &
            "Do not rely on the host to rewrite, infer, defer, queue, or replay omitted tool calls."

        Public Const ConsolidatableToolConsolidationInstruction As String =
            "A tool designed to complete an entire task in a single call has already run successfully in this session." & vbCrLf &
            "Do not issue additional calls to that tool for work that could have been included in the earlier call." & vbCrLf &
            "Only call it again if you genuinely must inspect the earlier result before deciding the next step; otherwise consolidate all remaining deterministic processing into one call."

        ''' <summary>
        ''' Explains the large-result 'drawer' to the model: big results are replaced by a short
        ''' result_ref plus a preview, are re-readable via context_expand, and can be voluntarily
        ''' shelved via context_compact when no longer needed in full. Appended only when at least
        ''' one of those tools is advertised.
        ''' </summary>
        Public Const ContextDrawerInstruction As String =
            "CONTEXT MANAGEMENT: Large tool results are not kept in full in the conversation. Each is replaced by a short 'result_ref' plus a preview, and the full text stays available." & vbCrLf &
            "To read more of a stored result, call context_expand with its result_ref (optionally start_char and max_chars) to page through the full content." & vbCrLf &
            "When you no longer need older results in full, you may call context_compact to move them out of the active context and free space; they remain retrievable via context_expand. Prefer letting the host manage this automatically, and use context_compact only when you know earlier results are no longer needed."

        Public Const UnresolvedToolFailureCode As String = "unresolved_tool_failure"
        Public Const InvalidTextOnlyFinalizationCode As String = "invalid_text_only_finalization"
        Public Const MissingRequiredMemoryAccessCode As String = "missing_required_memory_access"
        Public Const MemoryListDoneButMemoryGetRequiredCode As String = "memory_list_done_but_memory_get_required"
        Public Const MemoryGetFailedCode As String = "memory_get_failed"
        Public Const NoRelevantMemoryAvailableCode As String = "no_relevant_memory_available"
        Public Const PartialMemoryRetrievalRequiresSubsetDisclosureCode As String = "partial_memory_retrieval_requires_subset_disclosure"
        Public Const RequestedDeliverableNotCreatedCode As String = "requested_deliverable_not_created"
        Public Const RequestedDeliverableSlotsIncompleteCode As String = "requested_deliverable_slots_incomplete"

        Public Const RequiredMemoryGetAllThreshold As Integer = 10

        Public Const ToolNotExposedInCurrentTurnCode As String = "tool_not_exposed_in_current_turn"


        Public Enum TaskStatusKind
            None
            Complete
            Blocked
            ContinueTurn
        End Enum

        Public Enum ActiveToolingTurnKind
            InvalidTurn
            ToolCallTurn
            FinalCompleteTurn
            FinalBlockedTurn
        End Enum

        Public Enum MemoryGroundingMode
            None
            OptionalMode
            Required
        End Enum

        Public Enum MemoryGroundingStage
            NotStarted
            ListRequired
            GetRequired
            FullMemoryAvailable
            NoRelevantMemory
            Blocked
        End Enum

        Public Enum MemoryGroundingAuthority
            None
            Classifier
            ExplicitOverride
        End Enum

        Public NotInheritable Class TaskStatusParseResult
            Public Property IsPresent As Boolean
            Public Property IsValid As Boolean
            Public Property Status As TaskStatusKind
            Public Property Reason As String
            Public Property FooterCount As Integer
            Public Property FailureReason As String
            Public Property FooterJson As String
            Public Property TextBeforeFooter As String
            Public Property MemoryGroundingScope As String

            Public ReadOnly Property MemoryGroundingScopeIsSubset As Boolean
                Get
                    Return String.Equals(
                        If(MemoryGroundingScope, ""),
                        "subset",
                        StringComparison.OrdinalIgnoreCase)
                End Get
            End Property

            Public ReadOnly Property Summary As String
                Get
                    If Not IsPresent Then Return "missing"
                    If Not IsValid Then Return "invalid:" & If(FailureReason, "")
                    Return Status.ToString().ToLowerInvariant()
                End Get
            End Property
        End Class

        Public NotInheritable Class ActiveToolingTurnValidationResult
            Public Property TurnKind As ActiveToolingTurnKind
            Public Property InvalidReason As String
            Public Property TaskStatus As TaskStatusParseResult

            Public ReadOnly Property TaskStatusSummary As String
                Get
                    If TaskStatus Is Nothing Then Return "missing"
                    Return TaskStatus.Summary
                End Get
            End Property
        End Class

        Public Enum ToolCallClassification
            ReadOnlyIndependent
            Mutating
            Stateful
            Skill
            Agent
            Unknown
        End Enum

        ''' <summary>
        ''' Defines how an unresolved tool failure may be cleared by later successful work.
        ''' The policy is host-agnostic and deliberately conservative: simply issuing another
        ''' tool call never recovers a failure; recovery is recorded only after a successful
        ''' substantive tool result.
        ''' </summary>
        Public Enum ToolFailureRecoveryPolicy
            SameToolSuccessOnly
            CompatibleAlternativeSuccessAllowed
            DifferentAlternativeSuccessOnly
            NoAutomaticRecovery
        End Enum

        Public Enum ToolFailureCategory
            Validation
            ArtifactContract
            DocumentProcessing
            Transport
            Model
            Unknown
        End Enum

        Public NotInheritable Class ToolFailureRecord
            Public Property Sequence As Long
            Public Property ToolName As String = String.Empty
            Public Property ErrorCode As String = String.Empty
            Public Property ErrorMessage As String = String.Empty
            Public Property Category As ToolFailureCategory = ToolFailureCategory.Unknown
            Public Property SkippedByPolicy As Boolean
            Public Property ReturnedToParent As Boolean
            Public Property Terminal As Boolean
            Public Property ToolErrorHandling As String = String.Empty
            Public Property ToolClassification As ToolCallClassification = ToolCallClassification.Unknown
            Public Property RecoveryScopeKey As String = String.Empty
            ' Logical operation / concrete step / physical attempt are tracked separately.
            ' RecoveryScopeKey remains the compatibility-facing opaque scope; LogicalOperationKey
            ' groups related steps, while StepKey identifies the exact retry unit. Attempt ids are
            ' host-generated and never supplied by the model.
            Public Property LogicalOperationKey As String = String.Empty
            Public Property StepKey As String = String.Empty
            Public Property FirstAttemptSequence As Long
            Public Property LastAttemptSequence As Long
            Public Property AttemptCount As Integer
            Public Property RecoveryEvidenceStepKey As String = String.Empty
            Public Property RecoveryPolicy As ToolFailureRecoveryPolicy = ToolFailureRecoveryPolicy.SameToolSuccessOnly
            ' True only for a pre-execution validation failure whose invalid/incomplete input
            ' could not yet supply a trustworthy recovery scope. A corrected retry may therefore
            ' bind a real scope, but only when the host-owned repair-correlation signature below
            ' proves that all non-repairable arguments are unchanged.
            Public Property AllowRetryScopeRebinding As System.Boolean = False
            Public Property RetryCorrelationSignature As System.String = System.String.Empty
            Public Property RetryCorrelationRepairablePaths As System.Collections.Generic.List(Of System.String) =
                New System.Collections.Generic.List(Of System.String)()
            ' An alternative-path success is not always cleared immediately. When causal identity
            ' cannot be proven by an explicit scope (or the failure was delegated to the parent),
            ' success is recorded as recovery evidence and committed only on an accepted final turn.
            Public Property RecoveryEvidenceObserved As Boolean
            Public Property RecoveryEvidenceToolName As String = String.Empty
            ' Host-owned bounded whole-path recovery may deliberately cross an opaque
            ' operation/task scope. Normal alternative recovery remains scope-exact.
            Public Property CrossScopeAlternativeRecoveryAllowed As Boolean = False
            ' Optional tool-declared output-type contract for cross-scope fallback recovery.
            ' Empty preserves generic behavior; e.g. .docx requires a replacement DOCX.
            Public Property CrossScopeAlternativeRecoveryRequiredArtifactExtension As System.String = System.String.Empty
            ' Substantive progress epoch at which this failure occurred. Consecutive fallback
            ' failures created after the same prior substantive success share an epoch and may
            ' be superseded together by one later successful alternative path.
            Public Property ProgressEpoch As Long
        End Class

        Public NotInheritable Class PlannedToolCall
            Public Property Index As Integer
            Public Property ToolName As String
            Public Property Classification As ToolCallClassification
            Public Property IsBarrier As Boolean
            Public Property WillExecute As Boolean
            Public Property SkipReason As String
        End Class

        Public NotInheritable Class ToolBatchPlan

            Public Sub New()
                Calls = New List(Of PlannedToolCall)()
            End Sub

            Public Property Calls As List(Of PlannedToolCall)

            Public ReadOnly Property TotalCallCount As Integer
                Get
                    Return Calls.Count
                End Get
            End Property

            Public ReadOnly Property ExecutedCount As Integer
                Get
                    Dim count As Integer = 0

                    For Each item In Calls
                        If item IsNot Nothing AndAlso item.WillExecute Then
                            count += 1
                        End If
                    Next

                    Return count
                End Get
            End Property

            Public ReadOnly Property DeferredCount As Integer
                Get
                    Dim count As Integer = 0

                    For Each item In Calls
                        If item IsNot Nothing AndAlso Not item.WillExecute Then
                            count += 1
                        End If
                    Next

                    Return count
                End Get
            End Property

            Public ReadOnly Property IsFullyBatchSafe As Boolean
                Get
                    If Calls.Count = 0 Then Return False

                    For Each item In Calls
                        If item Is Nothing Then Return False
                        If item.IsBarrier Then Return False
                        If Not item.WillExecute Then Return False
                    Next

                    Return True
                End Get
            End Property

        End Class

        ''' <summary>
        ''' A single host-agnostic deliverable artifact that was verified to exist on
        ''' disk when it was registered. Used as the source of truth for the completion
        ''' gate and for host-side delivery (Outlook attachment / Word output copy).
        ''' </summary>
        Public NotInheritable Class DeliverableArtifact
            Public Property ArtifactId As String = ""
            Public Property LogicalDeliverableId As String = ""
            Public Property OutputSlotId As String = ""
            Public Property SessionPath As String = ""
            Public Property SourceTool As String = ""
            Public Property LegacyCompatibilityEligible As Boolean
            Public Property WasObservedLegacyFileDelta As Boolean
            Public Property IsFinalDeliverable As Boolean
            Public Property LifecycleState As ArtifactLifecycleState = ArtifactLifecycleState.Intermediate
            Public Property DeliveryIntent As ArtifactDeliveryIntent = ArtifactDeliveryIntent.None
            Public Property StorageKind As ArtifactStorageKind = ArtifactStorageKind.Unknown
            Public Property SupersedesArtifactId As String = ""
            Public Property IsExplicitContract As Boolean
            Public Property RegisteredUtc As DateTime
            ' Trusted host/tool evidence about what happened to this physical artifact.
            ' "materialized" is added when an existing file is registered. Additional
            ' opaque effects are added only by host-authored tool capabilities after
            ' successful execution; model-supplied artifact JSON cannot forge them.
            Public Property VerifiedEffects As System.Collections.Generic.HashSet(Of System.String) =
                New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
        End Class

        ''' <summary>
        ''' Exact caller-declared identity of one expected user-facing output slot.
        ''' Both values are opaque. No filename, path, extension, prompt, or semantic
        ''' inference participates in expected-output matching.
        ''' </summary>
        Public NotInheritable Class ExpectedDeliverableSlot
            Public Property LogicalDeliverableId As String = ""
            Public Property OutputSlotId As String = ""
            ' Opaque completion effects required in addition to physical existence.
            Public Property RequiredEffects As System.Collections.Generic.HashSet(Of System.String) =
                New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
        End Class

        Public NotInheritable Class ToolingRunState
            Public Property HasUnresolvedToolFailure As Boolean
            Public Property LastToolName As String
            Public Property LastErrorCode As String
            Public Property LastErrorMessage As String
            Public Property LastFailureSkippedByPolicy As Boolean
            Public Property LastFailureReturnedToParent As Boolean
            Public Property LastFailureRecoveredByToolCall As Boolean
            Public Property LastFailureHandledByBlockedFinal As Boolean
            Public Property LastFailureUltimatelyFatal As Boolean
            Public Property RecoveryToolName As String
            Public Property LastRecoveredFailureSummary As String = String.Empty
            Public Property LastFailureTerminal As Boolean
            Public Property LastFailureRecoveryPolicy As ToolFailureRecoveryPolicy = ToolFailureRecoveryPolicy.SameToolSuccessOnly
            Public Property LastFailureToolClassification As ToolCallClassification = ToolCallClassification.Unknown
            Public Property LastFailureToolErrorHandling As String = String.Empty
            Public Property UnresolvedToolFailures As New List(Of ToolFailureRecord)()
            Private _failureSequence As Long
            Private _attemptSequence As Long
            Private _substantiveProgressEpoch As Long
            Private _artifactRevisionSequence As Long

            ' Tool-agnostic retry fidelity state. When a failed tool call carried a named
            ' design/template constraint, a retry of that same tool may not silently drop or
            ' replace it. The dispatcher uses the shared helper below in both Outlook and Word.
            Public Property RetryInvariantArgumentsByTool As New System.Collections.Generic.Dictionary(Of String, System.Collections.Generic.Dictionary(Of String, String))(System.StringComparer.Ordinal)
            Public Property RetryInvariantPendingFailureTools As New System.Collections.Generic.HashSet(Of String)(System.StringComparer.Ordinal)

            Public Property ActiveToolingSession As Boolean
            Public Property HasOpenToolWorkflow As Boolean
            Public Property LastStateFilePath As String
            Public Property LastOutputPath As String
            Public Property LastCollectionSize As Integer?
            Public Property LastProcessedItemCount As Integer?
            Public Property LastSuccessfulToolCall As String
            Public Property RequiredSuccessfulTools As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
            Public Property RequiredSuccessfulToolsBeforeFinalMutation As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
            Public Property SuccessfulToolsThisRun As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
            Public Property LastMutationToolCall As String
            Public Property LastAgentToolCall As String
            Public Property LastReadOnlyStateToolCall As String
            Public Property LastDetectedTurnType As String
            Public Property LastInvalidTurnReason As String
            Public Property FinalResponseOrigin As String
            Public Property ToolRequiredModeUsed As Boolean

            Public Property UserLanguage As String
            Public Property UserSuppliedSourceFormatAuthority As System.Boolean = False
            Public Property UserSuppliedSourceFormatAuthorityReason As System.String = System.String.Empty
            Public Property LastStructuredToolResult As String
            Public Property LastStructuredToolResultKind As String
            Public Property LastStructuredToolName As String
            Public Property LastKnownOutputReference As String

            Public Property RequestRequiresCreatedDeliverable As Boolean
            Public Property RequestDeliverableSummary As String
            Public Property LastToolProducesIntermediateData As Boolean
            Public Property LastToolProducesUserDeliverable As Boolean
            Public Property LastToolOutputArtifactRef As String
            Public Property LastToolOutputFilePath As String
            Public Property LastToolOutputMimeType As String
            Public Property LastToolOutputKind As String
            Public Property AnyUserDeliverableProducedThisRun As Boolean
            Public Property OperationRegistry As ExplicitOperationRegistry =
                New ExplicitOperationRegistry()

            Public Property SubAgentTaskRegistry As ExplicitSubAgentTaskRegistry =
                New ExplicitSubAgentTaskRegistry()

            ''' <summary>
            ''' Per-run allow-list of tool names that are capable of producing a user
            ''' deliverable (host-provided from HostToolRegistration.GetDeliverableCapableToolNames).
            ''' Used to prevent read-only tools (e.g. text extract/search) from registering or
            ''' promoting a deliverable just because their result echoes a generic 'path' field.
            ''' When empty/unpopulated, deliverable inference falls back to the prior behavior so
            ''' hosts that do not set this cannot regress.
            ''' </summary>
            Public Property DeliverableCapableToolNames As HashSet(Of String) =
                New System.Collections.Generic.HashSet(Of String)(System.StringComparer.OrdinalIgnoreCase)

            ''' <summary>
            ''' True when the given tool name is allowed to produce a deliverable. Fails open
            ''' (returns True) when the capability set is empty so unpopulated hosts keep prior behavior.
            ''' </summary>
            Public Function IsDeliverableCapableTool(toolName As String) As Boolean
                If DeliverableCapableToolNames Is Nothing OrElse DeliverableCapableToolNames.Count = 0 Then
                    Return True
                End If
                Dim name As String = If(toolName, "").Trim()
                Return name <> "" AndAlso DeliverableCapableToolNames.Contains(name)
            End Function

            ''' <summary>
            ''' Authoritative per-run registry of deliverable artifacts that were verified
            ''' to exist on disk at registration time. Shared by all tooling hosts
            ''' (Outlook AutoPilot, Outlook Local Agent, Word) as the single source of
            ''' truth for the completion gate and for host-side delivery.
            ''' </summary>
            Public Property RegisteredDeliverableArtifacts As List(Of DeliverableArtifact) =
                New List(Of DeliverableArtifact)()

            ''' <summary>
            ''' Physical paths claimed by a tool call that explicitly declared artifacts[].
            ''' This set is suppression-only compatibility telemetry: it never establishes
            ''' artifact identity, lifecycle, finality, supersession, or delivery intent. It
            ''' prevents malformed/conflicting explicit artifact payloads from being silently
            ''' reintroduced later through older host-side path-only compatibility channels.
            ''' </summary>
            Public Property ExplicitArtifactProtocolOwnedPaths As HashSet(Of String) =
                New System.Collections.Generic.HashSet(Of String)(System.StringComparer.OrdinalIgnoreCase)

            Public Sub RegisterExplicitArtifactProtocolOwnedPath(candidatePath As String)
                Dim rawPath As String = If(candidatePath, "").Trim()
                If rawPath = "" Then Return

                If ExplicitArtifactProtocolOwnedPaths Is Nothing Then
                    ExplicitArtifactProtocolOwnedPaths =
                        New System.Collections.Generic.HashSet(Of String)(System.StringComparer.OrdinalIgnoreCase)
                End If

                Try
                    ExplicitArtifactProtocolOwnedPaths.Add(System.IO.Path.GetFullPath(rawPath))
                Catch ex As System.Exception
                    ' Invalid path text remains unresolved and cannot participate in
                    ' compatibility suppression or delivery.
                End Try
            End Sub

            Public Property ExpectedDeliverableSlots As List(Of ExpectedDeliverableSlot) =
                New List(Of ExpectedDeliverableSlot)()

            Public Property ExpectedDeliverableContractLocked As Boolean = False

            ' True once a syntactically valid expected_artifacts array has been declared,
            ' including the intentional empty contract []. This is distinct from Locked:
            ' top-level runs may declare an authoritative contract without being delegated.
            Public Property ExpectedDeliverableContractDeclared As Boolean = False

            Public ReadOnly Property HasExpectedDeliverableContract As Boolean
                Get
                    Return ExpectedDeliverableContractDeclared OrElse
                           ExpectedDeliverableContractLocked OrElse
                           (ExpectedDeliverableSlots IsNot Nothing AndAlso
                            ExpectedDeliverableSlots.Count > 0)
                End Get
            End Property

            Public Sub RegisterRequiredSuccessfulTools(toolNames As System.Collections.Generic.IEnumerable(Of System.String))
                If toolNames Is Nothing Then Return
                If RequiredSuccessfulTools Is Nothing Then
                    RequiredSuccessfulTools = New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
                End If

                For Each rawName As System.String In toolNames
                    Dim toolName As System.String = If(rawName, System.String.Empty).Trim()
                    If toolName = System.String.Empty Then Continue For
                    RequiredSuccessfulTools.Add(toolName)
                Next
            End Sub

            Public Sub RegisterRequiredSuccessfulToolsBeforeFinalMutation(toolNames As System.Collections.Generic.IEnumerable(Of System.String))
                If toolNames Is Nothing Then Return
                If RequiredSuccessfulToolsBeforeFinalMutation Is Nothing Then
                    RequiredSuccessfulToolsBeforeFinalMutation =
                        New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
                End If

                For Each rawName As System.String In toolNames
                    Dim toolName As System.String = If(rawName, System.String.Empty).Trim()
                    If toolName = System.String.Empty Then Continue For
                    RequiredSuccessfulToolsBeforeFinalMutation.Add(toolName)
                Next
            End Sub

            Public Sub RegisterSuccessfulTool(toolName As System.String)
                Dim normalizedToolName As System.String = If(toolName, System.String.Empty).Trim()
                If normalizedToolName = System.String.Empty Then Return
                If SuccessfulToolsThisRun Is Nothing Then
                    SuccessfulToolsThisRun = New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
                End If
                SuccessfulToolsThisRun.Add(normalizedToolName)
            End Sub

            Public Function GetMissingRequiredSuccessfulTools() As System.Collections.Generic.List(Of System.String)
                Dim missing As New System.Collections.Generic.List(Of System.String)()
                If RequiredSuccessfulTools Is Nothing OrElse RequiredSuccessfulTools.Count = 0 Then Return missing

                For Each requiredTool As System.String In RequiredSuccessfulTools
                    If SuccessfulToolsThisRun Is Nothing OrElse Not SuccessfulToolsThisRun.Contains(requiredTool) Then
                        missing.Add(requiredTool)
                    End If
                Next

                missing.Sort(System.StringComparer.OrdinalIgnoreCase)
                Return missing
            End Function

            Public Function GetMissingRequiredSuccessfulToolsBeforeFinalMutation() As System.Collections.Generic.List(Of System.String)
                Dim missing As New System.Collections.Generic.List(Of System.String)()
                If RequiredSuccessfulToolsBeforeFinalMutation Is Nothing OrElse
                   RequiredSuccessfulToolsBeforeFinalMutation.Count = 0 Then
                    Return missing
                End If

                For Each requiredTool As System.String In RequiredSuccessfulToolsBeforeFinalMutation
                    If SuccessfulToolsThisRun Is Nothing OrElse Not SuccessfulToolsThisRun.Contains(requiredTool) Then
                        missing.Add(requiredTool)
                    End If
                Next

                missing.Sort(System.StringComparer.OrdinalIgnoreCase)
                Return missing
            End Function

            Public Sub RegisterExpectedDeliverableSlot(
                logicalDeliverableId As String,
                outputSlotId As String,
                Optional requiredEffects As System.Collections.Generic.IEnumerable(Of System.String) = Nothing)

                Dim logicalId As String = If(logicalDeliverableId, "").Trim()
                Dim slotId As String = If(outputSlotId, "").Trim()

                If logicalId = "" OrElse slotId = "" Then Return

                If ExpectedDeliverableSlots Is Nothing Then
                    ExpectedDeliverableSlots = New List(Of ExpectedDeliverableSlot)()
                End If

                Dim normalizedEffects As System.Collections.Generic.HashSet(Of System.String) =
                    NormalizeArtifactEffectSet(requiredEffects, includeMaterialized:=True)

                For Each existing As ExpectedDeliverableSlot In ExpectedDeliverableSlots
                    If existing Is Nothing Then Continue For

                    If String.Equals(existing.LogicalDeliverableId,
                                     logicalId,
                                     StringComparison.Ordinal) AndAlso
                       String.Equals(existing.OutputSlotId,
                                     slotId,
                                     StringComparison.Ordinal) Then
                        If existing.RequiredEffects Is Nothing Then
                            existing.RequiredEffects = New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
                        End If
                        existing.RequiredEffects.UnionWith(normalizedEffects)
                        Return
                    End If
                Next

                ExpectedDeliverableSlots.Add(
                    New ExpectedDeliverableSlot With {
                        .LogicalDeliverableId = logicalId,
                        .OutputSlotId = slotId,
                        .RequiredEffects = normalizedEffects
                    })

                RequestRequiresCreatedDeliverable = True
            End Sub

            Private Shared Function NormalizeArtifactEffectSet(
                effects As System.Collections.Generic.IEnumerable(Of System.String),
                Optional includeMaterialized As System.Boolean = False) As System.Collections.Generic.HashSet(Of System.String)

                Dim result As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
                If includeMaterialized Then result.Add("materialized")
                If effects Is Nothing Then Return result

                For Each rawEffect As System.String In effects
                    Dim effect As System.String = If(rawEffect, System.String.Empty).Trim().ToLowerInvariant()
                    If effect = System.String.Empty Then Continue For
                    If Not System.Text.RegularExpressions.Regex.IsMatch(effect, "^[a-z][a-z0-9_.-]{0,63}$") Then Continue For
                    result.Add(effect)
                Next

                Return result
            End Function

            Private Shared Function ParseRequiredEffects(
                token As Newtonsoft.Json.Linq.JToken) As System.Collections.Generic.HashSet(Of System.String)

                If token Is Nothing OrElse token.Type = Newtonsoft.Json.Linq.JTokenType.Null Then
                    Return NormalizeArtifactEffectSet(Nothing, includeMaterialized:=True)
                End If

                Dim values As New System.Collections.Generic.List(Of System.String)()
                If token.Type = Newtonsoft.Json.Linq.JTokenType.Array Then
                    For Each item As Newtonsoft.Json.Linq.JToken In DirectCast(token, Newtonsoft.Json.Linq.JArray)
                        If item Is Nothing OrElse item.Type = Newtonsoft.Json.Linq.JTokenType.Null Then Continue For
                        values.Add(item.ToString())
                    Next
                Else
                    values.AddRange(token.ToString().Split(New System.Char() {","c, ";"c, " "c}, System.StringSplitOptions.RemoveEmptyEntries))
                End If

                Return NormalizeArtifactEffectSet(values, includeMaterialized:=True)
            End Function

            Public Sub ApplyRequiredEffectsToExpectedDeliverables(
                effects As System.Collections.Generic.IEnumerable(Of System.String))

                If ExpectedDeliverableSlots Is Nothing OrElse ExpectedDeliverableSlots.Count = 0 Then Return
                Dim normalized As System.Collections.Generic.HashSet(Of System.String) =
                    NormalizeArtifactEffectSet(effects, includeMaterialized:=True)

                For Each expected As ExpectedDeliverableSlot In ExpectedDeliverableSlots
                    If expected Is Nothing Then Continue For
                    If expected.RequiredEffects Is Nothing Then
                        expected.RequiredEffects = New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
                    End If
                    expected.RequiredEffects.UnionWith(normalized)
                Next
            End Sub

            Public Shared Function IsExplicitSubAgentDelegationCall(
                toolName As String,
                arguments As IDictionary(Of String, Object)) As Boolean

                If String.IsNullOrWhiteSpace(toolName) OrElse arguments Is Nothing Then Return False
                If Not toolName.StartsWith("agent_", StringComparison.OrdinalIgnoreCase) Then Return False

                Dim rawTaskId As Object = Nothing
                If Not arguments.TryGetValue("subagent_task_id", rawTaskId) OrElse rawTaskId Is Nothing Then
                    Return False
                End If

                Return Not String.IsNullOrWhiteSpace(System.Convert.ToString(rawTaskId))
            End Function

            Public Function ValidateExpectedArtifactArguments(
                arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object),
                ByRef failureMessage As System.String) As System.Boolean

                failureMessage = System.String.Empty
                If arguments Is Nothing Then Return True

                Dim raw As System.Object = Nothing
                If Not arguments.TryGetValue("expected_artifacts", raw) Then Return True

                If raw Is Nothing Then
                    failureMessage = "expected_artifacts must be an array; use [] for an explicitly empty final-artifact contract."
                    Return False
                End If

                Dim token As Newtonsoft.Json.Linq.JToken = Nothing
                Try
                    token = Newtonsoft.Json.Linq.JToken.FromObject(raw)
                Catch ex As System.Exception
                    failureMessage = "expected_artifacts could not be parsed as an array."
                    Return False
                End Try

                If token Is Nothing OrElse token.Type <> Newtonsoft.Json.Linq.JTokenType.Array Then
                    failureMessage = "expected_artifacts must be an array; use [] for an explicitly empty final-artifact contract."
                    Return False
                End If

                For Each item As Newtonsoft.Json.Linq.JToken In DirectCast(token, Newtonsoft.Json.Linq.JArray)
                    Dim obj As Newtonsoft.Json.Linq.JObject = TryCast(item, Newtonsoft.Json.Linq.JObject)
                    If obj Is Nothing Then
                        failureMessage = "Every expected_artifacts item must be an object."
                        Return False
                    End If

                    Dim logicalId As System.String = System.String.Empty
                    If Not TryGetExpectedArtifactIdentityValue(
                        obj,
                        "logical_deliverable_id",
                        logicalId,
                        failureMessage) Then
                        Return False
                    End If

                    Dim slotId As System.String = System.String.Empty
                    If Not TryGetExpectedArtifactIdentityValue(
                        obj,
                        "output_slot_id",
                        slotId,
                        failureMessage) Then
                        Return False
                    End If
                Next

                Return True
            End Function

            Private Shared Function TryGetExpectedArtifactIdentityValue(
                obj As Newtonsoft.Json.Linq.JObject,
                fieldName As System.String,
                ByRef value As System.String,
                ByRef failureMessage As System.String) As System.Boolean

                value = System.String.Empty

                If obj Is Nothing OrElse System.String.IsNullOrWhiteSpace(fieldName) Then
                    failureMessage = "expected_artifacts contains an invalid identity field."
                    Return False
                End If

                Dim fieldToken As Newtonsoft.Json.Linq.JToken = obj(fieldName)
                If fieldToken Is Nothing OrElse
                   fieldToken.Type = Newtonsoft.Json.Linq.JTokenType.Null OrElse
                   fieldToken.Type = Newtonsoft.Json.Linq.JTokenType.Undefined Then
                    failureMessage =
                        "expected_artifacts item field '" & fieldName & "' must be a non-empty scalar value."
                    Return False
                End If

                If Not TypeOf fieldToken Is Newtonsoft.Json.Linq.JValue Then
                    failureMessage =
                        "expected_artifacts item field '" & fieldName & "' must be a scalar value, not an object or array."
                    Return False
                End If

                Try
                    value = If(fieldToken.Value(Of System.String)(), System.String.Empty).Trim()
                Catch ex As System.Exception
                    failureMessage =
                        "expected_artifacts item field '" & fieldName & "' could not be converted to a scalar string value."
                    Return False
                End Try

                If value = System.String.Empty Then
                    failureMessage =
                        "expected_artifacts item field '" & fieldName & "' must be a non-empty scalar value."
                    Return False
                End If

                Return True
            End Function

            Public Sub RegisterExpectedDeliverablesFromArguments(
                arguments As IDictionary(Of String, Object))

                If ExpectedDeliverableContractLocked Then Return
                If arguments Is Nothing Then Return

                Dim raw As Object = Nothing
                If Not arguments.TryGetValue("expected_artifacts", raw) OrElse
                   raw Is Nothing Then
                    Return
                End If

                Dim token As JToken = Nothing

                Try
                    token = JToken.FromObject(raw)
                Catch ex As System.Exception
                    Return
                End Try

                If token Is Nothing OrElse token.Type <> JTokenType.Array Then Return

                ' Parse and validate the entire explicit contract before mutating run state.
                ' A malformed item must never leave a partially registered expected-slot set.
                Dim pendingSlots As New List(Of ExpectedDeliverableSlot)()

                For Each item As JToken In DirectCast(token, JArray)
                    Dim obj As JObject = TryCast(item, JObject)
                    If obj Is Nothing Then Return

                    Dim logicalId As String =
                        If(obj.Value(Of String)("logical_deliverable_id"), "").Trim()

                    Dim slotId As String =
                        If(obj.Value(Of String)("output_slot_id"), "").Trim()

                    If logicalId = "" OrElse slotId = "" Then Return

                    Dim duplicatePending As Boolean =
                        pendingSlots.Any(
                            Function(existing)
                                Return existing IsNot Nothing AndAlso
                                       String.Equals(
                                           If(existing.LogicalDeliverableId, ""),
                                           logicalId,
                                           StringComparison.Ordinal) AndAlso
                                       String.Equals(
                                           If(existing.OutputSlotId, ""),
                                           slotId,
                                           StringComparison.Ordinal)
                            End Function)

                    If Not duplicatePending Then
                        pendingSlots.Add(
                            New ExpectedDeliverableSlot With {
                                .LogicalDeliverableId = logicalId,
                                .OutputSlotId = slotId,
                                .RequiredEffects = ParseRequiredEffects(obj("required_effects"))
                            })
                    End If
                Next

                ExpectedDeliverableContractDeclared = True

                For Each pending As ExpectedDeliverableSlot In pendingSlots
                    RegisterExpectedDeliverableSlot(
                        pending.LogicalDeliverableId,
                        pending.OutputSlotId,
                        pending.RequiredEffects)
                Next
            End Sub

            Public Function ValidateLockedExpectedArtifactArguments(
                arguments As IDictionary(Of String, Object),
                ByRef failureReason As String) As Boolean

                Return ValidateLockedExpectedArtifactArguments(
                    arguments,
                    failureReason,
                    "")
            End Function

            Public Function ValidateLockedExpectedArtifactArguments(
                arguments As IDictionary(Of String, Object),
                ByRef failureReason As String,
                toolName As System.String) As Boolean

                failureReason = ""

                If Not ExpectedDeliverableContractLocked Then Return True
                If arguments Is Nothing Then Return True

                Dim hasArtifactId As Boolean =
                    arguments.ContainsKey("artifact_id") AndAlso
                    arguments("artifact_id") IsNot Nothing

                Dim hasLogicalId As Boolean =
                    arguments.ContainsKey("logical_deliverable_id") AndAlso
                    arguments("logical_deliverable_id") IsNot Nothing

                Dim hasSlotId As Boolean =
                    arguments.ContainsKey("output_slot_id") AndAlso
                    arguments("output_slot_id") IsNot Nothing

                Dim hasSupersedesArtifactId As Boolean =
                    arguments.ContainsKey("supersedes_artifact_id") AndAlso
                    arguments("supersedes_artifact_id") IsNot Nothing

                Dim hasDirectArtifactIdentity As Boolean =
                    hasArtifactId OrElse
                    hasLogicalId OrElse
                    hasSlotId OrElse
                    hasSupersedesArtifactId

                Dim hasExpectedArtifacts As System.Boolean =
                    arguments.ContainsKey("expected_artifacts")

                ' expected_artifacts on agent_* is the CHILD delegation contract and is
                ' intentionally independent from this run's locked contract. For a
                ' deliverable-producing call in THIS run, however, a locked multi-slot
                ' contract must never permit an ambiguous physical side effect. If the
                ' call clearly declares a final/deliverable output but no opaque slot
                ' pair is supplied while multiple slots remain unresolved, reject it
                ' before tool execution rather than inferring from file names or types.
                If Not hasDirectArtifactIdentity Then
                    If RequiresExplicitLockedProducerSlotSelection(toolName, arguments) Then
                        failureReason = "locked_expected_artifact_slot_required"
                        Return False
                    End If

                    If Not hasExpectedArtifacts OrElse
                       IsExplicitSubAgentDelegationCall(toolName, arguments) Then
                        Return True
                    End If
                End If

                If hasLogicalId OrElse hasSlotId OrElse hasArtifactId OrElse hasSupersedesArtifactId Then
                    Dim logicalId As String =
                        If(If(hasLogicalId, System.Convert.ToString(arguments("logical_deliverable_id")), ""), "").Trim()

                    Dim slotId As String =
                        If(If(hasSlotId, System.Convert.ToString(arguments("output_slot_id")), ""), "").Trim()

                    If logicalId = "" OrElse slotId = "" Then
                        failureReason = "locked_expected_artifact_identity_incomplete"
                        Return False
                    End If

                    If Not IsExpectedDeliverableSlot(logicalId, slotId) Then
                        failureReason = "locked_expected_artifact_slot_mismatch"
                        Return False
                    End If
                End If

                Dim rawExpected As Object = Nothing

                If Not arguments.TryGetValue("expected_artifacts", rawExpected) Then
                    Return True
                End If

                If rawExpected Is Nothing Then
                    failureReason = "locked_expected_artifact_contract_invalid"
                    Return False
                End If

                Dim token As JToken = Nothing

                Try
                    token = JToken.FromObject(rawExpected)
                Catch ex As System.Exception
                    failureReason = "locked_expected_artifact_contract_invalid"
                    Return False
                End Try

                If token Is Nothing OrElse token.Type <> JTokenType.Array Then
                    failureReason = "locked_expected_artifact_contract_invalid"
                    Return False
                End If

                Dim suppliedSlots As New List(Of ExpectedDeliverableSlot)()

                For Each item As JToken In DirectCast(token, JArray)
                    Dim obj As JObject = TryCast(item, JObject)
                    If obj Is Nothing Then
                        failureReason = "locked_expected_artifact_contract_invalid"
                        Return False
                    End If

                    Dim identityFailure As System.String = System.String.Empty
                    Dim logicalId As System.String = System.String.Empty
                    If Not TryGetExpectedArtifactIdentityValue(
                        obj,
                        "logical_deliverable_id",
                        logicalId,
                        identityFailure) Then

                        failureReason = "locked_expected_artifact_contract_invalid"
                        Return False
                    End If

                    Dim slotId As System.String = System.String.Empty
                    If Not TryGetExpectedArtifactIdentityValue(
                        obj,
                        "output_slot_id",
                        slotId,
                        identityFailure) Then

                        failureReason = "locked_expected_artifact_contract_invalid"
                        Return False
                    End If

                    Dim duplicateSupplied As Boolean =
                        suppliedSlots.Any(
                            Function(existing)
                                Return existing IsNot Nothing AndAlso
                                       String.Equals(
                                           If(existing.LogicalDeliverableId, ""),
                                           logicalId,
                                           StringComparison.Ordinal) AndAlso
                                       String.Equals(
                                           If(existing.OutputSlotId, ""),
                                           slotId,
                                           StringComparison.Ordinal)
                            End Function)

                    If Not duplicateSupplied Then
                        suppliedSlots.Add(
                            New ExpectedDeliverableSlot With {
                                .LogicalDeliverableId = logicalId,
                                .OutputSlotId = slotId,
                                .RequiredEffects = ParseRequiredEffects(obj("required_effects"))
                            })
                    End If
                Next

                Dim lockedCount As Integer =
                    If(ExpectedDeliverableSlots Is Nothing, 0, ExpectedDeliverableSlots.Count)

                If suppliedSlots.Count <> lockedCount Then
                    failureReason = "locked_expected_artifact_contract_mismatch"
                    Return False
                End If

                For Each expected As ExpectedDeliverableSlot In ExpectedDeliverableSlots
                    If expected Is Nothing Then
                        failureReason = "locked_expected_artifact_contract_invalid"
                        Return False
                    End If

                    Dim found As Boolean =
                        suppliedSlots.Any(
                            Function(supplied)
                                Return supplied IsNot Nothing AndAlso
                                       String.Equals(
                                           If(supplied.LogicalDeliverableId, ""),
                                           If(expected.LogicalDeliverableId, ""),
                                           StringComparison.Ordinal) AndAlso
                                       String.Equals(
                                           If(supplied.OutputSlotId, ""),
                                           If(expected.OutputSlotId, ""),
                                           StringComparison.Ordinal)
                            End Function)

                    If Not found Then
                        failureReason = "locked_expected_artifact_contract_mismatch"
                        Return False
                    End If
                Next

                Return True
            End Function

            Private Function RequiresExplicitLockedProducerSlotSelection(
                toolName As System.String,
                arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object)) As System.Boolean

                If arguments Is Nothing OrElse Not ExpectedDeliverableContractLocked Then Return False
                If ExpectedDeliverableSlots Is Nothing OrElse ExpectedDeliverableSlots.Count <= 1 Then Return False
                If Not IsDeliverableCapableTool(toolName) Then Return False

                Dim unresolvedCount As System.Int32 = 0
                For Each expected As ExpectedDeliverableSlot In ExpectedDeliverableSlots
                    If expected Is Nothing Then Continue For
                    If Not IsExpectedDeliverableSlotSatisfied(expected.LogicalDeliverableId, expected.OutputSlotId) Then
                        unresolvedCount += 1
                        If unresolvedCount > 1 Then Exit For
                    End If
                Next
                If unresolvedCount <= 1 Then Return False

                Dim stateText As System.String = GetArgumentText(arguments, "artifact_state")
                Dim intentText As System.String = GetArgumentText(arguments, "artifact_delivery_intent")
                If System.String.Equals(stateText, "final", System.StringComparison.OrdinalIgnoreCase) OrElse
                   System.String.Equals(intentText, "deliver_to_user", System.StringComparison.OrdinalIgnoreCase) OrElse
                   System.String.Equals(intentText, "deliver_and_persist", System.StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If

                If GetArgumentText(arguments, "output_filename") <> System.String.Empty Then Return True

                Dim normalizedToolName As System.String = If(toolName, System.String.Empty).Trim()
                If System.String.Equals(normalizedToolName, "python_execute", System.StringComparison.OrdinalIgnoreCase) Then
                    Dim code As System.String = GetArgumentText(arguments, "code")
                    If code.IndexOf("agent_api.output_path", System.StringComparison.OrdinalIgnoreCase) >= 0 Then Return True
                End If

                Return False
            End Function

            Private Function ResolveLockedProducerArtifactSlot(
                arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object)) As ExpectedDeliverableSlot

                If arguments Is Nothing OrElse Not ExpectedDeliverableContractLocked Then Return Nothing
                If ExpectedDeliverableSlots Is Nothing OrElse ExpectedDeliverableSlots.Count = 0 Then Return Nothing

                Dim suppliedLogicalId As System.String = GetArgumentText(arguments, "logical_deliverable_id")
                Dim suppliedSlotId As System.String = GetArgumentText(arguments, "output_slot_id")

                If suppliedLogicalId <> System.String.Empty OrElse suppliedSlotId <> System.String.Empty Then
                    If suppliedLogicalId = System.String.Empty OrElse suppliedSlotId = System.String.Empty Then Return Nothing
                    Return GetExpectedDeliverableSlot(suppliedLogicalId, suppliedSlotId)
                End If

                Dim unresolved As New System.Collections.Generic.List(Of ExpectedDeliverableSlot)()
                For Each expected As ExpectedDeliverableSlot In ExpectedDeliverableSlots
                    If expected Is Nothing Then Continue For
                    If Not IsExpectedDeliverableSlotSatisfied(expected.LogicalDeliverableId, expected.OutputSlotId) Then
                        unresolved.Add(expected)
                    End If
                Next

                If unresolved.Count = 1 Then Return unresolved(0)
                Return Nothing
            End Function

            Private Function BuildLockedExpectedArtifactArguments() As System.Collections.Generic.List(Of System.Object)
                Dim result As New System.Collections.Generic.List(Of System.Object)()
                If ExpectedDeliverableSlots Is Nothing Then Return result

                For Each expected As ExpectedDeliverableSlot In ExpectedDeliverableSlots
                    If expected Is Nothing Then Continue For
                    result.Add(
                        New System.Collections.Generic.Dictionary(Of System.String, System.Object)(System.StringComparer.Ordinal) From {
                            {"logical_deliverable_id", If(expected.LogicalDeliverableId, System.String.Empty)},
                            {"output_slot_id", If(expected.OutputSlotId, System.String.Empty)}
                        })
                Next

                Return result
            End Function

            ''' <summary>
            ''' Completes host-owned artifact metadata for an output-producing call whenever
            ''' its locked expected slot can be identified without heuristic inference. For
            ''' multiple-slot contracts the caller must name the exact host-issued logical/slot
            ''' pair while more than one slot remains open. Once exactly one unresolved slot
            ''' remains, the host may bind that slot automatically.
            ''' </summary>
            Public Function ShouldNormalizeLockedProducerArtifactArguments(
                toolName As System.String,
                arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object)) As System.Boolean

                Dim expected As ExpectedDeliverableSlot = ResolveLockedProducerArtifactSlot(arguments)
                If expected Is Nothing Then Return False

                If ArtifactDelivery.HasExplicitArtifactIdentityArguments(arguments) Then Return True

                Dim stateText As System.String = GetArgumentText(arguments, "artifact_state")
                Dim intentText As System.String = GetArgumentText(arguments, "artifact_delivery_intent")
                If System.String.Equals(stateText, "final", System.StringComparison.OrdinalIgnoreCase) OrElse
                   System.String.Equals(intentText, "deliver_to_user", System.StringComparison.OrdinalIgnoreCase) OrElse
                   System.String.Equals(intentText, "deliver_and_persist", System.StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If

                If GetArgumentText(arguments, "output_filename") <> System.String.Empty Then Return True

                Dim normalizedToolName As System.String = If(toolName, System.String.Empty).Trim()
                If System.String.Equals(normalizedToolName, "python_execute", System.StringComparison.OrdinalIgnoreCase) Then
                    Dim code As System.String = GetArgumentText(arguments, "code")
                    Return code.IndexOf("agent_api.output_path", System.StringComparison.OrdinalIgnoreCase) >= 0
                End If

                If System.String.Equals(normalizedToolName, "excel_complete_live_workbook", System.StringComparison.OrdinalIgnoreCase) AndAlso
                   arguments.ContainsKey("updates") AndAlso arguments("updates") IsNot Nothing Then
                    Dim attachmentName As System.String = GetArgumentText(arguments, "attachment_name")
                    If attachmentName <> System.String.Empty AndAlso RegisteredDeliverableArtifacts IsNot Nothing Then
                        For i As System.Int32 = RegisteredDeliverableArtifacts.Count - 1 To 0 Step -1
                            Dim current As DeliverableArtifact = RegisteredDeliverableArtifacts(i)
                            If current Is Nothing OrElse current.LifecycleState <> ArtifactLifecycleState.Final Then Continue For
                            If System.String.IsNullOrWhiteSpace(current.SessionPath) Then Continue For
                            Dim currentName As System.String = System.IO.Path.GetFileName(current.SessionPath)
                            If System.String.Equals(currentName, attachmentName, System.StringComparison.OrdinalIgnoreCase) Then Return True
                        Next
                    End If
                End If

                Return False
            End Function

            Public Sub NormalizeLockedProducerArtifactArguments(
                arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object),
                toolCallId As System.String)

                Dim expected As ExpectedDeliverableSlot = ResolveLockedProducerArtifactSlot(arguments)
                If expected Is Nothing Then Return

                Dim expectedLogicalId As System.String = If(expected.LogicalDeliverableId, System.String.Empty).Trim()
                Dim expectedSlotId As System.String = If(expected.OutputSlotId, System.String.Empty).Trim()
                If expectedLogicalId = System.String.Empty OrElse expectedSlotId = System.String.Empty Then Return

                Dim suppliedLogicalId As System.String = GetArgumentText(arguments, "logical_deliverable_id")
                Dim suppliedSlotId As System.String = GetArgumentText(arguments, "output_slot_id")

                If suppliedLogicalId <> System.String.Empty AndAlso
                   Not System.String.Equals(suppliedLogicalId, expectedLogicalId, System.StringComparison.Ordinal) Then Return
                If suppliedSlotId <> System.String.Empty AndAlso
                   Not System.String.Equals(suppliedSlotId, expectedSlotId, System.StringComparison.Ordinal) Then Return

                If suppliedLogicalId = System.String.Empty Then arguments("logical_deliverable_id") = expectedLogicalId
                If suppliedSlotId = System.String.Empty Then arguments("output_slot_id") = expectedSlotId

                Dim artifactId As System.String = BuildHostArtifactRevisionId(toolCallId, expectedLogicalId, expectedSlotId)
                arguments("artifact_id") = artifactId
                arguments.Remove("supersedes_artifact_id")

                If RegisteredDeliverableArtifacts IsNot Nothing Then
                    For i As System.Int32 = RegisteredDeliverableArtifacts.Count - 1 To 0 Step -1
                        Dim prior As DeliverableArtifact = RegisteredDeliverableArtifacts(i)
                        If prior Is Nothing Then Continue For
                        If prior.LifecycleState = ArtifactLifecycleState.Superseded Then Continue For
                        If Not System.String.Equals(If(prior.LogicalDeliverableId, System.String.Empty).Trim(), expectedLogicalId, System.StringComparison.Ordinal) Then Continue For
                        If Not System.String.Equals(If(prior.OutputSlotId, System.String.Empty).Trim(), expectedSlotId, System.StringComparison.Ordinal) Then Continue For
                        If Not System.String.IsNullOrWhiteSpace(prior.ArtifactId) Then
                            arguments("supersedes_artifact_id") = prior.ArtifactId.Trim()
                        End If
                        Exit For
                    Next
                End If

                If GetArgumentText(arguments, "artifact_state") = System.String.Empty Then
                    arguments("artifact_state") = "final"
                End If
                If GetArgumentText(arguments, "artifact_delivery_intent") = System.String.Empty Then
                    arguments("artifact_delivery_intent") = "deliver_to_user"
                End If

                Dim expectedRaw As System.Object = Nothing
                If Not arguments.TryGetValue("expected_artifacts", expectedRaw) OrElse expectedRaw Is Nothing Then
                    arguments("expected_artifacts") = BuildLockedExpectedArtifactArguments()
                End If
            End Sub

            Private Shared Function GetArgumentText(arguments As IDictionary(Of String, Object), key As String) As String
                If arguments Is Nothing OrElse System.String.IsNullOrWhiteSpace(key) Then Return ""
                Dim raw As Object = Nothing
                If Not arguments.TryGetValue(key, raw) OrElse raw Is Nothing Then Return ""
                Return If(System.Convert.ToString(raw), "").Trim()
            End Function

            Private Function BuildHostArtifactRevisionId(toolCallId As String, logicalId As String, slotId As String) As String
                _artifactRevisionSequence += 1
                Dim seed As String =
                    _artifactRevisionSequence.ToString(System.Globalization.CultureInfo.InvariantCulture) & "|" &
                    If(toolCallId, "").Trim() & "|" &
                    If(logicalId, "").Trim() & "|" &
                    If(slotId, "").Trim()

                Using sha As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
                    Dim bytes As Byte() = System.Text.Encoding.UTF8.GetBytes(seed)
                    Dim hash As Byte() = sha.ComputeHash(bytes)
                    Dim builder As New System.Text.StringBuilder(24)
                    For i As Integer = 0 To System.Math.Min(11, hash.Length - 1)
                        builder.Append(hash(i).ToString("x2", System.Globalization.CultureInfo.InvariantCulture))
                    Next
                    Return "host_art_" & builder.ToString()
                End Using
            End Function

            Public Function ValidateExplicitArtifactIdentityArguments(
                arguments As IDictionary(Of String, Object),
                ByRef failureReason As String,
                ByRef failureMessage As String) As Boolean

                failureReason = System.String.Empty
                failureMessage = System.String.Empty
                If arguments Is Nothing Then Return True

                Dim artifactId As System.String = GetArgumentText(arguments, "artifact_id")
                Dim logicalId As System.String = GetArgumentText(arguments, "logical_deliverable_id")
                Dim slotId As System.String = GetArgumentText(arguments, "output_slot_id")
                Dim supersedesId As System.String = GetArgumentText(arguments, "supersedes_artifact_id")

                Dim hasAnyExplicitArtifactIdentity As System.Boolean =
                    artifactId <> System.String.Empty OrElse
                    logicalId <> System.String.Empty OrElse
                    slotId <> System.String.Empty OrElse
                    supersedesId <> System.String.Empty

                If Not hasAnyExplicitArtifactIdentity Then Return True

                If artifactId = System.String.Empty OrElse
                   logicalId = System.String.Empty OrElse
                   slotId = System.String.Empty Then

                    Dim missingFields As New System.Collections.Generic.List(Of System.String)()
                    If artifactId = System.String.Empty Then missingFields.Add("artifact_id")
                    If logicalId = System.String.Empty Then missingFields.Add("logical_deliverable_id")
                    If slotId = System.String.Empty Then missingFields.Add("output_slot_id")

                    failureReason = "explicit_artifact_identity_incomplete"
                    failureMessage =
                        "Explicit artifact identity is incomplete. Missing non-empty field(s): " &
                        System.String.Join(", ", missingFields) &
                        ". artifact_id, logical_deliverable_id, and output_slot_id are required together when any explicit artifact identity value is supplied."
                    Return False
                End If

                If supersedesId <> System.String.Empty AndAlso
                   System.String.Equals(artifactId, supersedesId, System.StringComparison.Ordinal) Then
                    failureReason = "explicit_artifact_self_supersession"
                    failureMessage = "supersedes_artifact_id must not equal artifact_id."
                    Return False
                End If

                If RegisteredDeliverableArtifacts Is Nothing Then Return True

                Dim existingArtifact As DeliverableArtifact =
                    RegisteredDeliverableArtifacts.FirstOrDefault(
                        Function(existing)
                            Return existing IsNot Nothing AndAlso
                                   System.String.Equals(
                                       If(existing.ArtifactId, System.String.Empty),
                                       artifactId,
                                       System.StringComparison.Ordinal)
                        End Function)

                If existingArtifact IsNot Nothing Then
                    If Not System.String.Equals(
                        If(existingArtifact.LogicalDeliverableId, System.String.Empty).Trim(),
                        logicalId,
                        System.StringComparison.Ordinal) OrElse
                       Not System.String.Equals(
                        If(existingArtifact.OutputSlotId, System.String.Empty).Trim(),
                        slotId,
                        System.StringComparison.Ordinal) OrElse
                       Not System.String.Equals(
                        If(existingArtifact.SupersedesArtifactId, System.String.Empty).Trim(),
                        supersedesId,
                        System.StringComparison.Ordinal) Then

                        failureReason = "explicit_artifact_id_conflict"
                        failureMessage =
                            "artifact_id is already registered with different logical_deliverable_id, output_slot_id, or supersedes_artifact_id values."
                        Return False
                    End If

                    If existingArtifact.LifecycleState = ArtifactLifecycleState.Final OrElse
                       existingArtifact.LifecycleState = ArtifactLifecycleState.Superseded Then
                        failureReason = "explicit_artifact_id_terminal"
                        failureMessage = "artifact_id already identifies a Final or Superseded artifact and cannot be executed again."
                        Return False
                    End If
                End If

                If supersedesId <> System.String.Empty Then
                    Dim supersededArtifact As DeliverableArtifact =
                        RegisteredDeliverableArtifacts.FirstOrDefault(
                            Function(existing)
                                Return existing IsNot Nothing AndAlso
                                       System.String.Equals(
                                           If(existing.ArtifactId, System.String.Empty),
                                           supersedesId,
                                           System.StringComparison.Ordinal)
                            End Function)

                    If supersededArtifact Is Nothing OrElse
                       Not System.String.Equals(
                           If(supersededArtifact.LogicalDeliverableId, System.String.Empty).Trim(),
                           logicalId,
                           System.StringComparison.Ordinal) OrElse
                       Not System.String.Equals(
                           If(supersededArtifact.OutputSlotId, System.String.Empty).Trim(),
                           slotId,
                           System.StringComparison.Ordinal) Then

                        failureReason = "explicit_artifact_supersession_slot_mismatch"
                        failureMessage =
                            "supersedes_artifact_id must identify an existing artifact in the same logical_deliverable_id and output_slot_id."
                        Return False
                    End If
                End If

                Return True
            End Function

            Public Sub LockExpectedDeliverableContractFromJson(expectedArtifactsJson As String)
                ExpectedDeliverableContractLocked = False
                ExpectedDeliverableContractDeclared = False
                ExpectedDeliverableSlots = New List(Of ExpectedDeliverableSlot)()

                Dim json As String = If(expectedArtifactsJson, "").Trim()
                If json = "" Then
                    Throw New System.ArgumentException(
                        "expectedArtifactsJson is required; use [] explicitly for no expected final artifacts.")
                End If

                Dim token As JToken = JToken.Parse(json)
                If token.Type <> JTokenType.Array Then Throw New ArgumentException("expectedArtifactsJson must be a JSON array.")

                For Each item As JToken In DirectCast(token, JArray)
                    Dim obj As JObject = TryCast(item, JObject)
                    If obj Is Nothing Then Throw New ArgumentException("Each expected_artifacts item must be an object.")
                    Dim logicalId As String = If(obj.Value(Of String)("logical_deliverable_id"), "").Trim()
                    Dim slotId As String = If(obj.Value(Of String)("output_slot_id"), "").Trim()
                    If logicalId = "" OrElse slotId = "" Then Throw New ArgumentException("Each expected_artifacts item requires logical_deliverable_id and output_slot_id.")
                    RegisterExpectedDeliverableSlot(logicalId, slotId)
                Next

                ExpectedDeliverableContractDeclared = True
                ExpectedDeliverableContractLocked = True
            End Sub

            Public Function GetExpectedDeliverableSlot(
                logicalDeliverableId As System.String,
                outputSlotId As System.String) As ExpectedDeliverableSlot

                Dim logicalId As System.String = If(logicalDeliverableId, System.String.Empty).Trim()
                Dim slotId As System.String = If(outputSlotId, System.String.Empty).Trim()
                If logicalId = System.String.Empty OrElse slotId = System.String.Empty OrElse ExpectedDeliverableSlots Is Nothing Then Return Nothing

                For Each expected As ExpectedDeliverableSlot In ExpectedDeliverableSlots
                    If expected Is Nothing Then Continue For
                    If System.String.Equals(If(expected.LogicalDeliverableId, System.String.Empty), logicalId, System.StringComparison.Ordinal) AndAlso
                       System.String.Equals(If(expected.OutputSlotId, System.String.Empty), slotId, System.StringComparison.Ordinal) Then
                        Return expected
                    End If
                Next

                Return Nothing
            End Function

            Public Function ArtifactSatisfiesExpectedEffects(
                artifact As DeliverableArtifact,
                expected As ExpectedDeliverableSlot) As System.Boolean

                If artifact Is Nothing OrElse expected Is Nothing Then Return False
                Dim required As System.Collections.Generic.HashSet(Of System.String) =
                    NormalizeArtifactEffectSet(expected.RequiredEffects, includeMaterialized:=True)
                Dim verified As System.Collections.Generic.HashSet(Of System.String) =
                    NormalizeArtifactEffectSet(artifact.VerifiedEffects, includeMaterialized:=False)

                For Each effect As System.String In required
                    If Not verified.Contains(effect) Then Return False
                Next
                Return True
            End Function

            Public Function IsQualifiedFinalArtifactForDelivery(artifact As DeliverableArtifact) As System.Boolean
                If artifact Is Nothing Then Return False
                If Not HasExpectedDeliverableContract Then Return True

                Dim expected As ExpectedDeliverableSlot = GetExpectedDeliverableSlot(
                    artifact.LogicalDeliverableId,
                    artifact.OutputSlotId)
                If expected Is Nothing Then Return False
                Return ArtifactSatisfiesExpectedEffects(artifact, expected)
            End Function

            Public Sub NoteVerifiedArtifactEffects(
                effects As System.Collections.Generic.IEnumerable(Of System.String),
                candidatePaths As System.Collections.Generic.IEnumerable(Of System.String),
                Optional logicalDeliverableId As System.String = "",
                Optional outputSlotId As System.String = "")

                If RegisteredDeliverableArtifacts Is Nothing Then Return
                Dim normalizedEffects As System.Collections.Generic.HashSet(Of System.String) =
                    NormalizeArtifactEffectSet(effects, includeMaterialized:=False)
                If normalizedEffects.Count = 0 Then Return

                Dim normalizedPaths As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
                If candidatePaths IsNot Nothing Then
                    For Each rawPath As System.String In candidatePaths
                        If System.String.IsNullOrWhiteSpace(rawPath) Then Continue For
                        Try
                            normalizedPaths.Add(System.IO.Path.GetFullPath(rawPath))
                        Catch ex As System.Exception
                        End Try
                    Next
                End If

                Dim logicalId As System.String = If(logicalDeliverableId, System.String.Empty).Trim()
                Dim slotId As System.String = If(outputSlotId, System.String.Empty).Trim()

                For Each artifact As DeliverableArtifact In RegisteredDeliverableArtifacts
                    If artifact Is Nothing Then Continue For
                    Dim pathMatches As System.Boolean = False
                    If Not System.String.IsNullOrWhiteSpace(artifact.SessionPath) AndAlso normalizedPaths.Count > 0 Then
                        Try
                            pathMatches = normalizedPaths.Contains(System.IO.Path.GetFullPath(artifact.SessionPath))
                        Catch ex As System.Exception
                        End Try
                    End If

                    Dim slotMatches As System.Boolean =
                        logicalId <> System.String.Empty AndAlso slotId <> System.String.Empty AndAlso
                        System.String.Equals(If(artifact.LogicalDeliverableId, System.String.Empty), logicalId, System.StringComparison.Ordinal) AndAlso
                        System.String.Equals(If(artifact.OutputSlotId, System.String.Empty), slotId, System.StringComparison.Ordinal)

                    If Not pathMatches AndAlso Not slotMatches Then Continue For
                    If artifact.VerifiedEffects Is Nothing Then
                        artifact.VerifiedEffects = New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
                    End If
                    artifact.VerifiedEffects.UnionWith(normalizedEffects)
                Next
            End Sub

            Public Sub NoteVerifiedArtifactEffectsFromSuccessfulTool(
                toolConfig As SharedLibrary.ModelConfig,
                arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object),
                responseText As System.String,
                Optional runtimeVerifiedEffects As System.Collections.Generic.IEnumerable(Of System.String) = Nothing)

                Dim effects As New System.Collections.Generic.List(Of System.String)()
                If toolConfig IsNot Nothing AndAlso Not System.String.IsNullOrWhiteSpace(toolConfig.VerifiedArtifactEffects) Then
                    effects.AddRange(
                        toolConfig.VerifiedArtifactEffects.Split(
                            New System.Char() {","c, ";"c, " "c},
                            System.StringSplitOptions.RemoveEmptyEntries))
                End If
                If runtimeVerifiedEffects IsNot Nothing Then
                    effects.AddRange(runtimeVerifiedEffects)
                End If
                If effects.Count = 0 Then Return

                Dim paths As New System.Collections.Generic.List(Of System.String)()

                Try
                    If Not System.String.IsNullOrWhiteSpace(responseText) Then
                        Dim root As Newtonsoft.Json.Linq.JObject = TryCast(Newtonsoft.Json.Linq.JToken.Parse(responseText), Newtonsoft.Json.Linq.JObject)
                        If root IsNot Nothing Then
                            CollectArtifactEffectPaths(root, paths)
                            Dim result As Newtonsoft.Json.Linq.JObject = TryCast(root("result"), Newtonsoft.Json.Linq.JObject)
                            If result IsNot Nothing Then CollectArtifactEffectPaths(result, paths)
                        End If
                    End If
                Catch ex As System.Exception
                End Try

                Dim logicalId As System.String = System.String.Empty
                Dim slotId As System.String = System.String.Empty
                Dim raw As System.Object = Nothing
                If arguments IsNot Nothing AndAlso arguments.TryGetValue("logical_deliverable_id", raw) AndAlso raw IsNot Nothing Then
                    logicalId = System.Convert.ToString(raw).Trim()
                End If
                raw = Nothing
                If arguments IsNot Nothing AndAlso arguments.TryGetValue("output_slot_id", raw) AndAlso raw IsNot Nothing Then
                    slotId = System.Convert.ToString(raw).Trim()
                End If

                NoteVerifiedArtifactEffects(effects, paths, logicalId, slotId)
            End Sub

            Private Shared Sub CollectArtifactEffectPaths(
                obj As Newtonsoft.Json.Linq.JObject,
                target As System.Collections.Generic.List(Of System.String))

                If obj Is Nothing OrElse target Is Nothing Then Return
                For Each key As System.String In New System.String() {"output_path", "outputFilePath", "output_file_path", "file_path", "path"}
                    Dim token As Newtonsoft.Json.Linq.JToken = obj(key)
                    If token IsNot Nothing AndAlso token.Type = Newtonsoft.Json.Linq.JTokenType.String Then
                        target.Add(token.ToString())
                    End If
                Next

                Dim artifacts As Newtonsoft.Json.Linq.JToken = obj("artifacts")
                If artifacts IsNot Nothing AndAlso artifacts.Type = Newtonsoft.Json.Linq.JTokenType.Array Then
                    For Each item As Newtonsoft.Json.Linq.JToken In DirectCast(artifacts, Newtonsoft.Json.Linq.JArray)
                        Dim artifactObj As Newtonsoft.Json.Linq.JObject = TryCast(item, Newtonsoft.Json.Linq.JObject)
                        If artifactObj Is Nothing Then Continue For
                        Dim pathToken As Newtonsoft.Json.Linq.JToken = artifactObj("path")
                        If pathToken IsNot Nothing AndAlso pathToken.Type = Newtonsoft.Json.Linq.JTokenType.String Then
                            target.Add(pathToken.ToString())
                        End If
                    Next
                End If
            End Sub

            Public Function IsExpectedDeliverableSlot(logicalDeliverableId As String, outputSlotId As String) As Boolean
                Dim logicalId As String = If(logicalDeliverableId, "").Trim()
                Dim slotId As String = If(outputSlotId, "").Trim()
                If logicalId = "" OrElse slotId = "" OrElse ExpectedDeliverableSlots Is Nothing Then Return False
                For Each expected As ExpectedDeliverableSlot In ExpectedDeliverableSlots
                    If expected Is Nothing Then Continue For
                    If String.Equals(If(expected.LogicalDeliverableId, ""), logicalId, StringComparison.Ordinal) AndAlso String.Equals(If(expected.OutputSlotId, ""), slotId, StringComparison.Ordinal) Then Return True
                Next
                Return False
            End Function

            Public Function IsExpectedDeliverableSlotSatisfied(
                logicalDeliverableId As System.String,
                outputSlotId As System.String) As System.Boolean

                Dim expected As ExpectedDeliverableSlot =
                    GetExpectedDeliverableSlot(logicalDeliverableId, outputSlotId)
                If expected Is Nothing OrElse RegisteredDeliverableArtifacts Is Nothing Then Return False

                Dim currentFinalCount As System.Int32 = 0
                For Each artifact As DeliverableArtifact In RegisteredDeliverableArtifacts
                    If artifact Is Nothing Then Continue For
                    If artifact.LifecycleState <> ArtifactLifecycleState.Final Then Continue For
                    If Not artifact.IsFinalDeliverable Then Continue For
                    If Not artifact.IsExplicitContract Then Continue For
                    If System.String.IsNullOrWhiteSpace(artifact.ArtifactId) Then Continue For
                    If System.String.IsNullOrWhiteSpace(artifact.LogicalDeliverableId) Then Continue For
                    If System.String.IsNullOrWhiteSpace(artifact.OutputSlotId) Then Continue For

                    If artifact.DeliveryIntent <> ArtifactDeliveryIntent.DeliverToUser AndAlso
                       artifact.DeliveryIntent <> ArtifactDeliveryIntent.DeliverAndPersist Then
                        Continue For
                    End If

                    If Not System.String.Equals(
                        If(artifact.LogicalDeliverableId, System.String.Empty),
                        If(expected.LogicalDeliverableId, System.String.Empty),
                        System.StringComparison.Ordinal) Then
                        Continue For
                    End If

                    If Not System.String.Equals(
                        If(artifact.OutputSlotId, System.String.Empty),
                        If(expected.OutputSlotId, System.String.Empty),
                        System.StringComparison.Ordinal) Then
                        Continue For
                    End If

                    If System.String.IsNullOrWhiteSpace(artifact.SessionPath) Then Continue For
                    If Not ArtifactSatisfiesExpectedEffects(artifact, expected) Then Continue For

                    Try
                        If System.IO.File.Exists(artifact.SessionPath) Then
                            currentFinalCount += 1
                        End If
                    Catch ex As System.Exception
                    End Try
                Next

                Return currentFinalCount = 1
            End Function

            Public ReadOnly Property HasAllExpectedDeliverableSlots As Boolean
                Get
                    If ExpectedDeliverableSlots Is Nothing OrElse
                       ExpectedDeliverableSlots.Count = 0 Then
                        Return HasExpectedDeliverableContract
                    End If

                    If RegisteredDeliverableArtifacts Is Nothing Then Return False

                    For Each expected As ExpectedDeliverableSlot In ExpectedDeliverableSlots
                        If expected Is Nothing Then Return False

                        ' Exactly one current, existing user-facing Final must satisfy each
                        ' expected slot. Zero is incomplete; more than one is ambiguous/corrupt.
                        If Not IsExpectedDeliverableSlotSatisfied(
                            expected.LogicalDeliverableId,
                            expected.OutputSlotId) Then
                            Return False
                        End If
                    Next

                    Return True
                End Get
            End Property

            ''' <summary>
            ''' Registers a produced artifact path, but ONLY if the file actually exists on
            ''' disk. Paths that cannot be verified are ignored so a model can never satisfy
            ''' the completion gate with an unbacked path string. Path identity never establishes artifact identity or finality.
            ''' </summary>
            Public Sub RegisterExistingDeliverableArtifact(candidatePath As String,
                                                           sourceTool As String,
                                                           isFinalDeliverable As Boolean)
                ArtifactDelivery.RegisterLegacyPath(Me, candidatePath, sourceTool, isFinalDeliverable)
            End Sub

            ''' <summary>
            ''' Returns True only if at least one registered deliverable artifact still
            ''' exists on disk. This is the authoritative completion condition for
            ''' file-required tasks and must not rely on unverified metadata strings.
            ''' </summary>
            Public ReadOnly Property HasValidatedFinalDeliverable As Boolean
                Get
                    Return ArtifactDelivery.HasValidatedFinalDeliverable(Me)
                End Get
            End Property

            ''' <summary>
            ''' Completion-safe deliverable check. Explicit expected-artifact contracts remain
            ''' authoritative. Without such a contract, a bounded Legacy compatibility output may
            ''' satisfy completion without being promoted to an explicit Registry Final.
            ''' </summary>
            Public ReadOnly Property HasValidatedDeliverableForCompletion As Boolean
                Get
                    If HasExpectedDeliverableContract Then
                        If ExpectedDeliverableSlots Is Nothing OrElse ExpectedDeliverableSlots.Count = 0 Then
                            Return False
                        End If
                        Return HasAllExpectedDeliverableSlots
                    End If

                    If HasValidatedFinalDeliverable Then Return True

                    Try
                        Dim legacyPaths As System.Collections.Generic.List(Of String) =
                            ArtifactDelivery.ResolveLegacyCompatibilityPaths(Me)
                        Return legacyPaths IsNot Nothing AndAlso legacyPaths.Count > 0
                    Catch ex As System.Exception
                        Return False
                    End Try
                End Get
            End Property

            Public Property ConsolidatableToolSuccessCounts As Dictionary(Of String, Integer)
            Public Property LastConsolidatableToolName As String

            Public Function NoteConsolidatableToolSuccess(toolName As String) As Integer
                If String.IsNullOrWhiteSpace(toolName) Then Return 0

                If ConsolidatableToolSuccessCounts Is Nothing Then
                    ConsolidatableToolSuccessCounts =
                        New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
                End If

                Dim current As Integer = 0
                ConsolidatableToolSuccessCounts.TryGetValue(toolName, current)
                current += 1
                ConsolidatableToolSuccessCounts(toolName) = current
                LastConsolidatableToolName = toolName
                Return current
            End Function

            Public ReadOnly Property HasRepeatedConsolidatableToolCalls As Boolean
                Get
                    If ConsolidatableToolSuccessCounts Is Nothing Then Return False
                    For Each pair In ConsolidatableToolSuccessCounts
                        If pair.Value > 1 Then Return True
                    Next
                    Return False
                End Get
            End Property

            Public Property MemoryGroundingMode As MemoryGroundingMode
            Public Property MemoryGroundingAuthority As MemoryGroundingAuthority
            Public Property MemoryGroundingStage As MemoryGroundingStage
            Public Property ShouldExposeRecentMemoryStubs As Boolean
            Public Property MemoryListCalledThisTurn As Boolean
            Public Property MemoryGetCalledThisTurn As Boolean
            Public Property FullMemoryValueAvailableThisTurn As Boolean
            Public Property MemoryListReturnedNoEntriesThisTurn As Boolean
            Public Property MemoryListEntryCount As Integer
            Public Property MemoryGetCountThisTurn As Integer
            Public Property MemoryGetRequiredAfterList As Boolean
            Public Property MemoryKeysSuggestedForGet As List(Of String)
            Public Property FinalCompleteRejectedForMissingMemoryAccess As Boolean
            Public Property FinalCompleteRejectedForPartialMemoryRetrieval As Boolean
            Public Property MemoryKeysRetrievedThisTurn As List(Of String)
            Public Property FinalAnswerBasedOnSubset As Boolean

            Public ReadOnly Property IsRequiredMemoryGroundingEnforced As Boolean
                Get
                    Return MemoryGroundingMode = MemoryGroundingMode.Required AndAlso
               MemoryGroundingAuthority = MemoryGroundingAuthority.ExplicitOverride
                End Get
            End Property


            Public ReadOnly Property RequiresParentRecovery As Boolean
                Get
                    Return HasUnresolvedToolFailure AndAlso
                   LastFailureSkippedByPolicy AndAlso
                   LastFailureReturnedToParent
                End Get
            End Property

            Public ReadOnly Property HasTerminalUnresolvedToolFailure As Boolean
                Get
                    If UnresolvedToolFailures Is Nothing OrElse UnresolvedToolFailures.Count = 0 Then
                        Return False
                    End If

                    For Each failure As ToolFailureRecord In UnresolvedToolFailures
                        If failure IsNot Nothing AndAlso failure.Terminal Then
                            Return True
                        End If
                    Next

                    Return False
                End Get
            End Property

            Public Sub NoteToolFailure(toolName As String,
                               Optional errorCode As String = "",
                               Optional errorMessage As String = "",
                               Optional skippedByPolicy As Boolean = False,
                               Optional returnedToParent As Boolean = False,
                               Optional toolErrorHandling As String = "retry",
                               Optional terminal As Boolean = False,
                               Optional recoveryScopeKey As String = "",
                               Optional allowRetryScopeRebinding As System.Boolean = False,
                               Optional retryCorrelationArguments As System.Collections.Generic.IDictionary(Of System.String, System.Object) = Nothing)
                Dim normalizedToolName As String = If(toolName, "").Trim()
                Dim normalizedHandling As String = If(toolErrorHandling, "").Trim()
                Dim normalizedRecoveryScopeKey As String = If(recoveryScopeKey, "").Trim()
                Dim logicalOperationKey As String = ResolveLogicalOperationKey(normalizedRecoveryScopeKey)
                Dim stepKey As String = BuildFailureStepKey(normalizedToolName, normalizedRecoveryScopeKey)
                Dim failurePolicy As ToolFailureRecoveryPolicy = ResolveFailureRecoveryPolicy(
                    normalizedHandling,
                    terminal,
                    skippedByPolicy,
                    returnedToParent)

                _failureSequence += 1
                _attemptSequence += 1

                If UnresolvedToolFailures Is Nothing Then
                    UnresolvedToolFailures = New List(Of ToolFailureRecord)()
                End If

                Dim retryCorrelationSignature As System.String = System.String.Empty
                Dim retryCorrelationRepairablePaths As System.Collections.Generic.List(Of System.String) = Nothing
                Dim hasExactRetryCorrelation As System.Boolean = False

                If normalizedRecoveryScopeKey = System.String.Empty AndAlso
                   retryCorrelationArguments IsNot Nothing Then

                    hasExactRetryCorrelation = TryBuildPreflightRetryCorrelation(
                        retryCorrelationArguments,
                        If(errorCode, System.String.Empty),
                        retryCorrelationSignature,
                        retryCorrelationRepairablePaths)
                End If

                ' Only the exact same concrete step is a retry. Scope-less preflight failures
                ' with host-owned repair correlation are additionally keyed by that correlation
                ' so two different rejected calls using the same tool cannot collapse together.
                Dim record As ToolFailureRecord = Nothing
                For i As Integer = UnresolvedToolFailures.Count - 1 To 0 Step -1
                    Dim candidate As ToolFailureRecord = UnresolvedToolFailures(i)
                    If candidate Is Nothing Then Continue For
                    If Not System.String.Equals(If(candidate.StepKey, ""), stepKey, System.StringComparison.Ordinal) Then Continue For

                    Dim candidateHasCorrelation As System.Boolean =
                        Not System.String.IsNullOrWhiteSpace(candidate.RetryCorrelationSignature)

                    If candidateHasCorrelation <> hasExactRetryCorrelation Then Continue For
                    If hasExactRetryCorrelation AndAlso
                       (Not System.String.Equals(
                           candidate.RetryCorrelationSignature,
                           retryCorrelationSignature,
                           System.StringComparison.Ordinal) OrElse
                        Not RetryCorrelationPathsMatch(
                            candidate.RetryCorrelationRepairablePaths,
                            retryCorrelationRepairablePaths)) Then
                        Continue For
                    End If

                    record = candidate
                    Exit For
                Next

                If record Is Nothing Then
                    record = New ToolFailureRecord With {
                        .FirstAttemptSequence = _attemptSequence,
                        .AttemptCount = 0
                    }
                    UnresolvedToolFailures.Add(record)
                End If

                record.Sequence = _failureSequence
                record.ToolName = normalizedToolName
                record.ErrorCode = If(errorCode, "")
                record.ErrorMessage = If(errorMessage, "")
                record.Category = ClassifyToolFailureCategory(normalizedToolName, record.ErrorCode, record.ErrorMessage)
                record.SkippedByPolicy = skippedByPolicy
                record.ReturnedToParent = returnedToParent
                record.Terminal = terminal
                record.ToolErrorHandling = normalizedHandling
                record.ToolClassification = ToolCallSequencing.ClassifyToolNameForRecovery(normalizedToolName)
                record.RecoveryScopeKey = normalizedRecoveryScopeKey
                record.LogicalOperationKey = logicalOperationKey
                record.StepKey = stepKey
                record.LastAttemptSequence = _attemptSequence
                record.AttemptCount += 1
                record.RecoveryPolicy = failurePolicy
                record.AllowRetryScopeRebinding = allowRetryScopeRebinding AndAlso hasExactRetryCorrelation
                record.RetryCorrelationSignature = If(hasExactRetryCorrelation, retryCorrelationSignature, System.String.Empty)
                record.RetryCorrelationRepairablePaths =
                    If(hasExactRetryCorrelation,
                       New System.Collections.Generic.List(Of System.String)(retryCorrelationRepairablePaths),
                       New System.Collections.Generic.List(Of System.String)())
                ' A new/failed retry invalidates prior replacement evidence for this exact step.
                record.RecoveryEvidenceObserved = False
                record.RecoveryEvidenceToolName = String.Empty
                record.RecoveryEvidenceStepKey = String.Empty
                record.CrossScopeAlternativeRecoveryAllowed = False
                record.CrossScopeAlternativeRecoveryRequiredArtifactExtension = System.String.Empty
                record.ProgressEpoch = _substantiveProgressEpoch

                ProjectLatestUnresolvedFailure()

                Dim retryInvariantKey As String = BuildRetryInvariantKey(normalizedToolName, normalizedRecoveryScopeKey)
                If retryInvariantKey <> "" AndAlso
                   RetryInvariantArgumentsByTool IsNot Nothing AndAlso
                   RetryInvariantArgumentsByTool.ContainsKey(retryInvariantKey) Then

                    If RetryInvariantPendingFailureTools Is Nothing Then
                        RetryInvariantPendingFailureTools = New System.Collections.Generic.HashSet(Of String)(System.StringComparer.Ordinal)
                    End If
                    RetryInvariantPendingFailureTools.Add(retryInvariantKey)
                End If
            End Sub

            ''' <summary>
            ''' Legacy compatibility hook. Merely issuing a later tool call is not evidence of
            ''' recovery, so this method intentionally never clears the unresolved failure.
            ''' Actual recovery is evaluated by NoteSuccessfulProgress after a successful result.
            ''' </summary>
            Public Sub NoteRecoveryByLaterToolCall(toolName As String)
                If Not HasUnresolvedToolFailure Then Return
                RecoveryToolName = ""
            End Sub

            Public Sub NoteBlockedFinalHandled()
                If Not HasUnresolvedToolFailure Then Return

                If UnresolvedToolFailures IsNot Nothing Then
                    UnresolvedToolFailures.Clear()
                End If

                HasUnresolvedToolFailure = False
                LastFailureRecoveredByToolCall = False
                LastFailureHandledByBlockedFinal = True
                LastFailureUltimatelyFatal = False
                RecoveryToolName = ""
                LastToolName = ""
                LastErrorCode = ""
                LastErrorMessage = ""
                LastFailureSkippedByPolicy = False
                LastFailureReturnedToParent = False
                LastFailureTerminal = False
                LastFailureRecoveryPolicy = ToolFailureRecoveryPolicy.SameToolSuccessOnly
                LastFailureToolClassification = ToolCallClassification.Unknown
                LastFailureToolErrorHandling = ""
                If RetryInvariantPendingFailureTools IsNot Nothing Then RetryInvariantPendingFailureTools.Clear()
                If RetryInvariantArgumentsByTool IsNot Nothing Then RetryInvariantArgumentsByTool.Clear()
            End Sub

            ''' <summary>
            ''' Allows one concrete failed tool to declare that a different successful recovery
            ''' path may cross an opaque operation/task scope. This does not itself clear anything.
            ''' The normal recovery policy, substantive-success checks and deliverable validation
            ''' remain authoritative. Exact same-tool retries continue to require exact scope identity.
            ''' </summary>
            Public Sub AllowCrossScopeAlternativeRecovery(failedToolName As System.String,
                                                           recoveryScopeKey As System.String,
                                                           Optional requiredArtifactExtension As System.String = "",
                                                           Optional recoveryLabel As String = "declared_alternative_path")
                If Not HasUnresolvedToolFailure OrElse UnresolvedToolFailures Is Nothing OrElse UnresolvedToolFailures.Count = 0 Then
                    Return
                End If

                Dim normalizedToolName As System.String = If(failedToolName, System.String.Empty).Trim()
                Dim normalizedScope As System.String = If(recoveryScopeKey, System.String.Empty).Trim()
                Dim normalizedRequiredExtension As System.String = If(requiredArtifactExtension, System.String.Empty).Trim()
                If normalizedRequiredExtension <> System.String.Empty AndAlso Not normalizedRequiredExtension.StartsWith(".", System.StringComparison.Ordinal) Then
                    normalizedRequiredExtension = "." & normalizedRequiredExtension
                End If
                If normalizedToolName = System.String.Empty Then Return

                For i As System.Int32 = UnresolvedToolFailures.Count - 1 To 0 Step -1
                    Dim candidate As ToolFailureRecord = UnresolvedToolFailures(i)
                    If candidate Is Nothing Then Continue For
                    If Not System.String.Equals(candidate.ToolName, normalizedToolName, System.StringComparison.OrdinalIgnoreCase) Then Continue For
                    If Not RecoveryScopeKeysMatch(candidate.RecoveryScopeKey, normalizedScope) Then Continue For
                    If candidate.Terminal Then Continue For
                    If candidate.RecoveryPolicy <> ToolFailureRecoveryPolicy.CompatibleAlternativeSuccessAllowed Then Continue For

                    candidate.CrossScopeAlternativeRecoveryAllowed = True
                    candidate.CrossScopeAlternativeRecoveryRequiredArtifactExtension = normalizedRequiredExtension
                    candidate.RecoveryEvidenceObserved = False
                    candidate.RecoveryEvidenceToolName = String.Empty
                    candidate.RecoveryEvidenceStepKey = String.Empty
                    RecoveryToolName = If(recoveryLabel, System.String.Empty)
                    Exit For
                Next

                ProjectLatestUnresolvedFailure()
            End Sub

            Public Sub BeginBoundedAlternativeRecovery(failedToolName As String,
                                                       Optional recoveryLabel As String = "full_tool_path_recovery",
                                                       Optional recoveryScopeKey As String = "")
                If Not HasUnresolvedToolFailure OrElse UnresolvedToolFailures Is Nothing OrElse UnresolvedToolFailures.Count = 0 Then
                    Return
                End If

                Dim normalizedToolName As String = If(failedToolName, String.Empty).Trim()
                Dim normalizedRecoveryScopeKey As String = If(recoveryScopeKey, String.Empty).Trim()
                If normalizedToolName = String.Empty Then Return

                Dim matchedFailure As Boolean = False
                For i As Integer = UnresolvedToolFailures.Count - 1 To 0 Step -1
                    Dim candidate As ToolFailureRecord = UnresolvedToolFailures(i)
                    If candidate Is Nothing Then Continue For
                    If Not System.String.Equals(candidate.ToolName, normalizedToolName, System.StringComparison.OrdinalIgnoreCase) Then Continue For
                    If normalizedRecoveryScopeKey <> String.Empty AndAlso
                       Not RecoveryScopeKeysMatch(candidate.RecoveryScopeKey, normalizedRecoveryScopeKey) Then Continue For

                    ' The local retry/circuit-breaker is exhausted, but the host has explicitly
                    ' opened one bounded whole-workflow recovery pass. The failed tool itself
                    ' remains exhausted; a materially different compatible substantive tool may
                    ' now supersede this failure.
                    candidate.Terminal = False
                    candidate.RecoveryPolicy = ToolFailureRecoveryPolicy.DifferentAlternativeSuccessOnly
                    candidate.RecoveryEvidenceObserved = False
                    candidate.RecoveryEvidenceToolName = String.Empty
                    candidate.RecoveryEvidenceStepKey = String.Empty
                    candidate.CrossScopeAlternativeRecoveryAllowed = True
                    candidate.CrossScopeAlternativeRecoveryRequiredArtifactExtension = System.String.Empty
                    candidate.ProgressEpoch = _substantiveProgressEpoch
                    matchedFailure = True
                Next

                If Not matchedFailure Then Return

                LastFailureUltimatelyFatal = False
                LastFailureTerminal = False
                LastFailureRecoveryPolicy = ToolFailureRecoveryPolicy.DifferentAlternativeSuccessOnly
                RecoveryToolName = If(recoveryLabel, String.Empty)
                ProjectLatestUnresolvedFailure()
            End Sub

            Public Sub NoteFailureFatal()
                If Not HasUnresolvedToolFailure Then Return

                Dim latest As ToolFailureRecord = GetLatestUnresolvedFailure()
                If latest IsNot Nothing Then
                    latest.Terminal = True
                    latest.RecoveryPolicy = ToolFailureRecoveryPolicy.NoAutomaticRecovery
                End If

                LastFailureUltimatelyFatal = True
                LastFailureTerminal = True
                LastFailureRecoveryPolicy = ToolFailureRecoveryPolicy.NoAutomaticRecovery
                If RetryInvariantPendingFailureTools IsNot Nothing Then RetryInvariantPendingFailureTools.Clear()
                If RetryInvariantArgumentsByTool IsNot Nothing Then RetryInvariantArgumentsByTool.Clear()
            End Sub

            ''' <summary>
            ''' Records successful substantive progress using Operation -> Step -> Attempt semantics.
            ''' An exact successful retry resolves only the matching failed StepKey immediately.
            ''' A successful different step is never treated as a retry; it may only record
            ''' replacement/alternative recovery evidence, which is committed after an accepted
            ''' final turn. Multiple successful executions of distinct step_ids inside the same
            ''' operation therefore remain valid continuations rather than accidental retries.
            ''' </summary>
            Public Sub NoteSuccessfulProgress(Optional toolName As String = "",
                                              Optional recoveryScopeKey As String = "",
                                              Optional successfulArguments As System.Collections.Generic.IDictionary(Of System.String, System.Object) = Nothing)
                Dim normalizedToolName As String = If(toolName, "").Trim()
                Dim normalizedRecoveryScopeKey As String = If(recoveryScopeKey, "").Trim()

                If IsRecoveryNeutralAdministrativeTool(normalizedToolName) Then
                    Return
                End If

                _attemptSequence += 1
                Dim currentProgressEpoch As Long = _substantiveProgressEpoch
                Dim successfulStepKey As String = BuildFailureStepKey(normalizedToolName, normalizedRecoveryScopeKey)
                Dim successfulLogicalOperationKey As String = ResolveLogicalOperationKey(normalizedRecoveryScopeKey)

                If Not HasUnresolvedToolFailure OrElse UnresolvedToolFailures Is Nothing OrElse UnresolvedToolFailures.Count = 0 Then
                    _substantiveProgressEpoch += 1
                    Return
                End If

                Dim recoveryIndex As Integer = -1
                Dim correlatedRetryMatchCount As System.Int32 = 0
                Dim correlatedRetryRecoveryIndex As System.Int32 = -1

                ' Exact retry: same concrete StepKey. A different step_id or operation_id is
                ' deliberately not a retry, even if it calls the same tool. Explicit sub-agent
                ' task recovery retains its historical cross-agent special case. Scope-less
                ' host-correlated preflight failures are cleared only when exactly one unresolved
                ' candidate can explain the successful call; ambiguous matches remain visible.
                For i As Integer = UnresolvedToolFailures.Count - 1 To 0 Step -1
                    Dim candidate As ToolFailureRecord = UnresolvedToolFailures(i)
                    If candidate Is Nothing Then Continue For
                    Dim exactRetryCorrelationMatches As System.Boolean =
                        DoesPreflightRetryCorrelationMatch(candidate, successfulArguments)
                    Dim sameStep As System.Boolean =
                        System.String.Equals(If(candidate.StepKey, ""), successfulStepKey, System.StringComparison.Ordinal) AndAlso
                        exactRetryCorrelationMatches
                    Dim sameExplicitSubAgentTask As System.Boolean =
                        IsSameExplicitSubAgentTaskRecovery(candidate, normalizedToolName, normalizedRecoveryScopeKey)
                    Dim sameToolScopeReboundRetry As System.Boolean =
                        candidate.AllowRetryScopeRebinding AndAlso
                        exactRetryCorrelationMatches AndAlso
                        candidate.ProgressEpoch = currentProgressEpoch AndAlso
                        System.String.Equals(
                            If(candidate.ToolName, System.String.Empty),
                            normalizedToolName,
                            System.StringComparison.OrdinalIgnoreCase)
                    If Not sameStep AndAlso Not sameExplicitSubAgentTask AndAlso Not sameToolScopeReboundRetry Then Continue For

                    If candidate.RecoveryPolicy <> ToolFailureRecoveryPolicy.SameToolSuccessOnly AndAlso
                       candidate.RecoveryPolicy <> ToolFailureRecoveryPolicy.CompatibleAlternativeSuccessAllowed Then
                        Continue For
                    End If

                    If Not System.String.IsNullOrWhiteSpace(candidate.RetryCorrelationSignature) Then
                        correlatedRetryMatchCount += 1
                        correlatedRetryRecoveryIndex = i
                        Continue For
                    End If

                    recoveryIndex = i
                    Exit For
                Next

                If recoveryIndex < 0 AndAlso correlatedRetryMatchCount = 1 Then
                    recoveryIndex = correlatedRetryRecoveryIndex
                End If

                If recoveryIndex >= 0 Then
                    Dim recovered As ToolFailureRecord = UnresolvedToolFailures(recoveryIndex)
                    LastRecoveredFailureSummary = BuildFailureRecoverySummary(recovered, normalizedToolName)
                    UnresolvedToolFailures.RemoveAt(recoveryIndex)

                    If recovered IsNot Nothing AndAlso RetryInvariantPendingFailureTools IsNot Nothing Then
                        RetryInvariantPendingFailureTools.Remove(
                            BuildRetryInvariantKey(recovered.ToolName, recovered.RecoveryScopeKey))
                    End If

                    If UnresolvedToolFailures.Count > 0 Then
                        ProjectLatestUnresolvedFailure()
                    Else
                        HasUnresolvedToolFailure = False
                        LastToolName = ""
                        LastErrorCode = ""
                        LastErrorMessage = ""
                        LastFailureSkippedByPolicy = False
                        LastFailureReturnedToParent = False
                        LastFailureTerminal = False
                        LastFailureRecoveryPolicy = ToolFailureRecoveryPolicy.SameToolSuccessOnly
                        LastFailureToolClassification = ToolCallClassification.Unknown
                        LastFailureToolErrorHandling = ""
                        LastFailureRecoveredByToolCall = True
                        LastFailureHandledByBlockedFinal = False
                        LastFailureUltimatelyFatal = False
                        RecoveryToolName = normalizedToolName
                    End If

                    _substantiveProgressEpoch += 1
                    Return
                End If

                ' Different successful steps are alternative/replacement candidates, not retries.
                ' Record evidence only. HasBlockingUnresolvedToolFailure will permit finalization
                ' when that evidence is sufficient, and FinalizeObservedAlternativeRecoveries
                ' commits it only after the final turn is actually accepted.
                For i As Integer = UnresolvedToolFailures.Count - 1 To 0 Step -1
                    Dim candidate As ToolFailureRecord = UnresolvedToolFailures(i)
                    If candidate Is Nothing Then Continue For

                    If System.String.IsNullOrWhiteSpace(candidate.RecoveryScopeKey) AndAlso
                       candidate.ProgressEpoch <> currentProgressEpoch Then
                        Exit For
                    End If

                    If Not CanSuccessfulStepProvideReplacementEvidence(
                        candidate,
                        normalizedToolName,
                        normalizedRecoveryScopeKey,
                        successfulLogicalOperationKey,
                        successfulStepKey) Then
                        Continue For
                    End If

                    candidate.RecoveryEvidenceObserved = True
                    candidate.RecoveryEvidenceToolName = normalizedToolName
                    candidate.RecoveryEvidenceStepKey = successfulStepKey
                Next

                If UnresolvedToolFailures.Count > 0 Then
                    ProjectLatestUnresolvedFailure()

                    Dim boundedAlternativeStillPending As Boolean =
                        UnresolvedToolFailures.Any(
                            Function(candidate As ToolFailureRecord)
                                Return candidate IsNot Nothing AndAlso
                                       candidate.ProgressEpoch = currentProgressEpoch AndAlso
                                       candidate.RecoveryPolicy = ToolFailureRecoveryPolicy.DifferentAlternativeSuccessOnly AndAlso
                                       Not candidate.RecoveryEvidenceObserved
                            End Function)
                    If boundedAlternativeStillPending Then Return
                End If

                ' Evidence does not clear the failure here. It only means a later accepted final
                ' turn may commit the replacement. Progress still advances so unrelated earlier
                ' unscoped failures cannot be accidentally recovered by much later work.
                _substantiveProgressEpoch += 1
            End Sub

            Private Function CanSuccessfulStepProvideReplacementEvidence(
                failure As ToolFailureRecord,
                successfulToolName As String,
                successfulRecoveryScopeKey As String,
                successfulLogicalOperationKey As String,
                successfulStepKey As String) As Boolean

                If failure Is Nothing Then Return False
                If failure.Terminal AndAlso Not failure.ReturnedToParent Then Return False
                If IsRecoveryNeutralAdministrativeTool(successfulToolName) Then Return False
                If System.String.Equals(If(failure.StepKey, ""), If(successfulStepKey, ""), System.StringComparison.Ordinal) Then Return False

                If failure.RecoveryPolicy <> ToolFailureRecoveryPolicy.CompatibleAlternativeSuccessAllowed AndAlso
                   failure.RecoveryPolicy <> ToolFailureRecoveryPolicy.DifferentAlternativeSuccessOnly Then
                    Return False
                End If

                If failure.RecoveryPolicy = ToolFailureRecoveryPolicy.DifferentAlternativeSuccessOnly AndAlso
                   System.String.Equals(If(failure.ToolName, ""), If(successfulToolName, ""), System.StringComparison.OrdinalIgnoreCase) Then
                    Return False
                End If

                If failure.ReturnedToParent Then Return True

                Dim sameLogicalOperation As Boolean =
                    Not System.String.IsNullOrWhiteSpace(failure.LogicalOperationKey) AndAlso
                    System.String.Equals(
                        failure.LogicalOperationKey,
                        If(successfulLogicalOperationKey, ""),
                        System.StringComparison.Ordinal)

                ' A different step in the same logical operation is a continuation, not
                ' recovery. It must not hide a failed sibling step unless the host explicitly
                ' opened an alternative-recovery path for that failure.
                If sameLogicalOperation AndAlso Not failure.CrossScopeAlternativeRecoveryAllowed Then
                    Return False
                End If

                If failure.CrossScopeAlternativeRecoveryAllowed Then
                    If RequestRequiresCreatedDeliverable Then
                        If Not IsDeliverableCapableTool(successfulToolName) OrElse
                           Not HasValidatedDeliverableForCompletion Then
                            Return False
                        End If
                        Return HasValidatedDeliverableForRecoveryExtension(
                            failure.CrossScopeAlternativeRecoveryRequiredArtifactExtension)
                    End If
                    Return True
                End If

                ' A different logical operation can supersede an earlier skipped producer only
                ' when the host has concrete outcome evidence: both steps are deliverable-capable
                ' and a validated current deliverable exists. This is the generic replacement
                ' case that fixes stale producer failures without relabeling the later call as a retry.
                If RequestRequiresCreatedDeliverable AndAlso
                   IsDeliverableCapableTool(failure.ToolName) AndAlso
                   IsDeliverableCapableTool(successfulToolName) AndAlso
                   HasValidatedDeliverableForCompletion Then
                    Return True
                End If

                ' Legacy unscoped skip failures retain the prior bounded-epoch behavior.
                If System.String.IsNullOrWhiteSpace(failure.RecoveryScopeKey) AndAlso
                   System.String.IsNullOrWhiteSpace(successfulRecoveryScopeKey) Then
                    Return True
                End If

                Return False
            End Function

            Private Function HasValidatedDeliverableForRecoveryExtension(requiredExtension As System.String) As System.Boolean
                Dim normalizedExtension As System.String = If(requiredExtension, System.String.Empty).Trim()
                If normalizedExtension = System.String.Empty Then Return True
                If Not normalizedExtension.StartsWith(".", System.StringComparison.Ordinal) Then
                    normalizedExtension = "." & normalizedExtension
                End If

                If RegisteredDeliverableArtifacts IsNot Nothing Then
                    For Each artifact As DeliverableArtifact In RegisteredDeliverableArtifacts
                        If artifact Is Nothing OrElse System.String.IsNullOrWhiteSpace(artifact.SessionPath) Then Continue For
                        If Not System.IO.File.Exists(artifact.SessionPath) Then Continue For
                        If Not System.String.Equals(
                            System.IO.Path.GetExtension(artifact.SessionPath),
                            normalizedExtension,
                            System.StringComparison.OrdinalIgnoreCase) Then Continue For

                        If artifact.LifecycleState = ArtifactLifecycleState.Final OrElse
                           artifact.IsFinalDeliverable OrElse
                           artifact.LegacyCompatibilityEligible Then
                            Return True
                        End If
                    Next
                End If

                Try
                    Dim legacyPaths As System.Collections.Generic.List(Of System.String) =
                        ArtifactDelivery.ResolveLegacyCompatibilityPaths(Me)
                    If legacyPaths Is Nothing Then Return False
                    For Each legacyPath As System.String In legacyPaths
                        If System.String.IsNullOrWhiteSpace(legacyPath) OrElse Not System.IO.File.Exists(legacyPath) Then Continue For
                        If System.String.Equals(
                            System.IO.Path.GetExtension(legacyPath),
                            normalizedExtension,
                            System.StringComparison.OrdinalIgnoreCase) Then
                            Return True
                        End If
                    Next
                Catch ex As System.Exception
                    Return False
                End Try

                Return False
            End Function

            ''' <summary>
            ''' Commits two-phase alternative recovery only after the host has accepted a final
            ''' turn. This prevents unrelated intermediate successes from prematurely erasing an
            ''' unresolved failure while still allowing generic fallback paths to complete.
            ''' </summary>
            Public Sub FinalizeObservedAlternativeRecoveries(Optional recoveryLabel As String = "alternative_recovery_finalized")
                If Not HasUnresolvedToolFailure OrElse UnresolvedToolFailures Is Nothing OrElse UnresolvedToolFailures.Count = 0 Then
                    Return
                End If

                For i As Integer = UnresolvedToolFailures.Count - 1 To 0 Step -1
                    Dim candidate As ToolFailureRecord = UnresolvedToolFailures(i)
                    If candidate Is Nothing OrElse Not candidate.RecoveryEvidenceObserved Then
                        Continue For
                    End If

                    LastRecoveredFailureSummary = BuildFailureRecoverySummary(
                        candidate,
                        candidate.RecoveryEvidenceToolName)
                    If RetryInvariantPendingFailureTools IsNot Nothing Then
                        RetryInvariantPendingFailureTools.Remove(
                            BuildRetryInvariantKey(candidate.ToolName, candidate.RecoveryScopeKey))
                    End If
                    UnresolvedToolFailures.RemoveAt(i)
                Next

                If UnresolvedToolFailures.Count > 0 Then
                    ProjectLatestUnresolvedFailure()
                Else
                    HasUnresolvedToolFailure = False
                    LastToolName = ""
                    LastErrorCode = ""
                    LastErrorMessage = ""
                    LastFailureSkippedByPolicy = False
                    LastFailureReturnedToParent = False
                    LastFailureTerminal = False
                    LastFailureRecoveryPolicy = ToolFailureRecoveryPolicy.SameToolSuccessOnly
                    LastFailureToolClassification = ToolCallClassification.Unknown
                    LastFailureToolErrorHandling = ""
                    LastFailureRecoveredByToolCall = True
                    LastFailureHandledByBlockedFinal = False
                    LastFailureUltimatelyFatal = False
                    RecoveryToolName = If(recoveryLabel, "")
                End If
            End Sub

            Public Function ClearLatestFailureByCode(errorCode As String,
                                                     recoveryLabel As String) As Boolean
                If UnresolvedToolFailures Is Nothing OrElse UnresolvedToolFailures.Count = 0 Then
                    Return False
                End If

                Dim normalizedErrorCode As String = If(errorCode, "").Trim()
                For i As Integer = UnresolvedToolFailures.Count - 1 To 0 Step -1
                    Dim candidate As ToolFailureRecord = UnresolvedToolFailures(i)
                    If candidate Is Nothing Then Continue For
                    If Not System.String.Equals(candidate.ErrorCode, normalizedErrorCode, System.StringComparison.OrdinalIgnoreCase) Then Continue For

                    UnresolvedToolFailures.RemoveAt(i)
                    If UnresolvedToolFailures.Count > 0 Then
                        ProjectLatestUnresolvedFailure()
                    Else
                        HasUnresolvedToolFailure = False
                        LastToolName = ""
                        LastErrorCode = ""
                        LastErrorMessage = ""
                        LastFailureSkippedByPolicy = False
                        LastFailureReturnedToParent = False
                        LastFailureTerminal = False
                        LastFailureRecoveryPolicy = ToolFailureRecoveryPolicy.SameToolSuccessOnly
                        LastFailureToolClassification = ToolCallClassification.Unknown
                        LastFailureToolErrorHandling = ""
                        LastFailureRecoveredByToolCall = True
                        LastFailureHandledByBlockedFinal = False
                        LastFailureUltimatelyFatal = False
                        RecoveryToolName = If(recoveryLabel, "")
                    End If
                    Return True
                Next

                Return False
            End Function

            Private Shared Function ResolveLogicalOperationKey(recoveryScopeKey As String) As String
                Dim normalized As String = If(recoveryScopeKey, "").Trim()
                If normalized = "" Then Return ""

                Const stepMarker As String = "|step:"
                If normalized.StartsWith("operation:", System.StringComparison.Ordinal) Then
                    Dim stepIndex As Integer = normalized.IndexOf(stepMarker, System.StringComparison.Ordinal)
                    If stepIndex > 0 Then Return normalized.Substring(0, stepIndex)
                End If

                Return normalized
            End Function

            Private Shared Function BuildFailureStepKey(toolName As String, recoveryScopeKey As String) As String
                Dim normalizedTool As String = If(toolName, "").Trim().ToLowerInvariant()
                Dim normalizedScope As String = If(recoveryScopeKey, "").Trim()
                Return "tool:" & normalizedTool & "|scope:" & normalizedScope
            End Function

            Private Shared Function IsSameExplicitSubAgentTaskRecovery(failure As ToolFailureRecord,
                                                                          successfulToolName As System.String,
                                                                          successfulRecoveryScopeKey As System.String) As System.Boolean
                If failure Is Nothing Then Return False
                Dim failureScope As System.String = If(failure.RecoveryScopeKey, System.String.Empty).Trim()
                If Not failureScope.StartsWith("subagent:", System.StringComparison.Ordinal) Then Return False
                If Not RecoveryScopeKeysMatch(failureScope, successfulRecoveryScopeKey) Then Return False
                Return Not System.String.IsNullOrWhiteSpace(successfulToolName) AndAlso
                       successfulToolName.StartsWith(AgentToolRouter.AgentToolPrefix, System.StringComparison.OrdinalIgnoreCase)
            End Function

            Private Shared Function RecoveryScopeKeysMatch(left As String,
                                                           right As String) As Boolean
                Return System.String.Equals(
                    If(left, "").Trim(),
                    If(right, "").Trim(),
                    System.StringComparison.Ordinal)
            End Function

            Private Const RetryCorrelationMaskValue As System.String = "__REDINK_HOST_REPAIRABLE_FIELD__"

            Private Shared Function RetryCorrelationPathsMatch(
                left As System.Collections.Generic.IList(Of System.String),
                right As System.Collections.Generic.IList(Of System.String)) As System.Boolean

                Dim leftCount As System.Int32 = If(left Is Nothing, 0, left.Count)
                Dim rightCount As System.Int32 = If(right Is Nothing, 0, right.Count)
                If leftCount <> rightCount Then Return False

                For index As System.Int32 = 0 To leftCount - 1
                    If Not System.String.Equals(
                        If(left(index), System.String.Empty),
                        If(right(index), System.String.Empty),
                        System.StringComparison.Ordinal) Then
                        Return False
                    End If
                Next

                Return True
            End Function

            Private Shared Function TryBuildPreflightRetryCorrelation(
                arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object),
                errorCode As System.String,
                ByRef signature As System.String,
                ByRef repairablePaths As System.Collections.Generic.List(Of System.String)) As System.Boolean

                signature = System.String.Empty
                repairablePaths = New System.Collections.Generic.List(Of System.String)()
                If arguments Is Nothing Then Return False

                Dim normalizedErrorCode As System.String = If(errorCode, System.String.Empty).Trim()
                Dim root As Newtonsoft.Json.Linq.JObject = Nothing

                Try
                    root = TryCast(Newtonsoft.Json.Linq.JToken.FromObject(arguments), Newtonsoft.Json.Linq.JObject)
                Catch ex As System.Exception
                    Return False
                End Try

                If root Is Nothing Then Return False

                If System.String.Equals(
                    normalizedErrorCode,
                    "explicit_artifact_identity_incomplete",
                    System.StringComparison.Ordinal) Then

                    Dim requiredIdentityFields As System.String() = {
                        "artifact_id",
                        "logical_deliverable_id",
                        "output_slot_id"
                    }

                    For Each fieldName As System.String In requiredIdentityFields
                        Dim rawValue As System.Object = Nothing
                        Dim valueText As System.String = System.String.Empty
                        If arguments.TryGetValue(fieldName, rawValue) AndAlso rawValue IsNot Nothing Then
                            valueText = If(System.Convert.ToString(rawValue), System.String.Empty).Trim()
                        End If
                        If valueText = System.String.Empty Then
                            repairablePaths.Add("field:" & fieldName)
                        End If
                    Next

                ElseIf System.String.Equals(
                    normalizedErrorCode,
                    "invalid_expected_artifacts",
                    System.StringComparison.Ordinal) Then

                    Dim expectedToken As Newtonsoft.Json.Linq.JToken = root("expected_artifacts")
                    If expectedToken Is Nothing OrElse
                       expectedToken.Type = Newtonsoft.Json.Linq.JTokenType.Null OrElse
                       expectedToken.Type <> Newtonsoft.Json.Linq.JTokenType.Array Then

                        repairablePaths.Add("field:expected_artifacts")
                    Else
                        Dim expectedArray As Newtonsoft.Json.Linq.JArray = DirectCast(expectedToken, Newtonsoft.Json.Linq.JArray)
                        For index As System.Int32 = 0 To expectedArray.Count - 1
                            Dim itemObject As Newtonsoft.Json.Linq.JObject = TryCast(expectedArray(index), Newtonsoft.Json.Linq.JObject)
                            If itemObject Is Nothing Then
                                repairablePaths.Add("expected_item:" & index.ToString(System.Globalization.CultureInfo.InvariantCulture))
                                Continue For
                            End If

                            For Each fieldName As System.String In New System.String() {"logical_deliverable_id", "output_slot_id"}
                                Dim ignoredValue As System.String = System.String.Empty
                                Dim ignoredFailure As System.String = System.String.Empty
                                If Not TryGetExpectedArtifactIdentityValue(
                                    itemObject,
                                    fieldName,
                                    ignoredValue,
                                    ignoredFailure) Then

                                    repairablePaths.Add(
                                        "expected_field:" &
                                        index.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                                        ":" & fieldName)
                                End If
                            Next
                        Next
                    End If
                Else
                    Return False
                End If

                If repairablePaths.Count = 0 Then Return False
                repairablePaths.Sort(System.StringComparer.Ordinal)

                Dim masked As Newtonsoft.Json.Linq.JObject = DirectCast(root.DeepClone(), Newtonsoft.Json.Linq.JObject)
                If Not ApplyRetryCorrelationMasks(masked, repairablePaths) Then Return False

                Dim canonical As Newtonsoft.Json.Linq.JToken = CanonicalizeRetryCorrelationToken(masked)
                signature = canonical.ToString(Newtonsoft.Json.Formatting.None)
                Return Not System.String.IsNullOrWhiteSpace(signature)
            End Function

            Private Shared Function DoesPreflightRetryCorrelationMatch(
                failure As ToolFailureRecord,
                successfulArguments As System.Collections.Generic.IDictionary(Of System.String, System.Object)) As System.Boolean

                If failure Is Nothing Then Return False
                If System.String.IsNullOrWhiteSpace(failure.RetryCorrelationSignature) Then Return True
                If successfulArguments Is Nothing OrElse
                   failure.RetryCorrelationRepairablePaths Is Nothing OrElse
                   failure.RetryCorrelationRepairablePaths.Count = 0 Then
                    Return False
                End If

                Dim root As Newtonsoft.Json.Linq.JObject = Nothing
                Try
                    root = TryCast(Newtonsoft.Json.Linq.JToken.FromObject(successfulArguments), Newtonsoft.Json.Linq.JObject)
                Catch ex As System.Exception
                    Return False
                End Try
                If root Is Nothing Then Return False

                Dim masked As Newtonsoft.Json.Linq.JObject = DirectCast(root.DeepClone(), Newtonsoft.Json.Linq.JObject)
                If Not ApplyRetryCorrelationMasks(masked, failure.RetryCorrelationRepairablePaths) Then Return False

                Dim canonical As Newtonsoft.Json.Linq.JToken = CanonicalizeRetryCorrelationToken(masked)
                Dim successfulSignature As System.String = canonical.ToString(Newtonsoft.Json.Formatting.None)
                Return System.String.Equals(
                    failure.RetryCorrelationSignature,
                    successfulSignature,
                    System.StringComparison.Ordinal)
            End Function

            Private Shared Function ApplyRetryCorrelationMasks(
                root As Newtonsoft.Json.Linq.JObject,
                repairablePaths As System.Collections.Generic.IList(Of System.String)) As System.Boolean

                If root Is Nothing OrElse repairablePaths Is Nothing Then Return False

                For Each path As System.String In repairablePaths
                    Dim normalizedPath As System.String = If(path, System.String.Empty)
                    If normalizedPath.StartsWith("field:", System.StringComparison.Ordinal) Then
                        Dim fieldName As System.String = normalizedPath.Substring("field:".Length)
                        If System.String.IsNullOrWhiteSpace(fieldName) Then Return False
                        root(fieldName) = New Newtonsoft.Json.Linq.JValue(RetryCorrelationMaskValue)
                        Continue For
                    End If

                    If normalizedPath.StartsWith("expected_item:", System.StringComparison.Ordinal) Then
                        Dim indexText As System.String = normalizedPath.Substring("expected_item:".Length)
                        Dim index As System.Int32
                        If Not System.Int32.TryParse(
                            indexText,
                            System.Globalization.NumberStyles.None,
                            System.Globalization.CultureInfo.InvariantCulture,
                            index) Then
                            Return False
                        End If

                        Dim expectedArray As Newtonsoft.Json.Linq.JArray = TryCast(root("expected_artifacts"), Newtonsoft.Json.Linq.JArray)
                        If expectedArray Is Nothing OrElse index < 0 OrElse index >= expectedArray.Count Then Return False
                        expectedArray(index) = New Newtonsoft.Json.Linq.JValue(RetryCorrelationMaskValue)
                        Continue For
                    End If

                    If normalizedPath.StartsWith("expected_field:", System.StringComparison.Ordinal) Then
                        Dim remainder As System.String = normalizedPath.Substring("expected_field:".Length)
                        Dim separatorIndex As System.Int32 = remainder.IndexOf(":"c)
                        If separatorIndex <= 0 OrElse separatorIndex >= remainder.Length - 1 Then Return False

                        Dim indexText As System.String = remainder.Substring(0, separatorIndex)
                        Dim fieldName As System.String = remainder.Substring(separatorIndex + 1)
                        Dim index As System.Int32
                        If Not System.Int32.TryParse(
                            indexText,
                            System.Globalization.NumberStyles.None,
                            System.Globalization.CultureInfo.InvariantCulture,
                            index) Then
                            Return False
                        End If

                        Dim expectedArray As Newtonsoft.Json.Linq.JArray = TryCast(root("expected_artifacts"), Newtonsoft.Json.Linq.JArray)
                        If expectedArray Is Nothing OrElse index < 0 OrElse index >= expectedArray.Count Then Return False
                        Dim itemObject As Newtonsoft.Json.Linq.JObject = TryCast(expectedArray(index), Newtonsoft.Json.Linq.JObject)
                        If itemObject Is Nothing OrElse System.String.IsNullOrWhiteSpace(fieldName) Then Return False
                        itemObject(fieldName) = New Newtonsoft.Json.Linq.JValue(RetryCorrelationMaskValue)
                        Continue For
                    End If

                    Return False
                Next

                Return True
            End Function

            Private Shared Function CanonicalizeRetryCorrelationToken(
                token As Newtonsoft.Json.Linq.JToken) As Newtonsoft.Json.Linq.JToken

                If token Is Nothing Then Return Newtonsoft.Json.Linq.JValue.CreateNull()

                Dim objectToken As Newtonsoft.Json.Linq.JObject = TryCast(token, Newtonsoft.Json.Linq.JObject)
                If objectToken IsNot Nothing Then
                    Dim properties As New System.Collections.Generic.List(Of Newtonsoft.Json.Linq.JProperty)(objectToken.Properties())
                    properties.Sort(
                        Function(left As Newtonsoft.Json.Linq.JProperty, right As Newtonsoft.Json.Linq.JProperty) As System.Int32
                            Return System.StringComparer.Ordinal.Compare(left.Name, right.Name)
                        End Function)

                    Dim canonicalObject As New Newtonsoft.Json.Linq.JObject()
                    For Each propertyToken As Newtonsoft.Json.Linq.JProperty In properties
                        canonicalObject.Add(
                            propertyToken.Name,
                            CanonicalizeRetryCorrelationToken(propertyToken.Value))
                    Next
                    Return canonicalObject
                End If

                Dim arrayToken As Newtonsoft.Json.Linq.JArray = TryCast(token, Newtonsoft.Json.Linq.JArray)
                If arrayToken IsNot Nothing Then
                    Dim canonicalArray As New Newtonsoft.Json.Linq.JArray()
                    For Each item As Newtonsoft.Json.Linq.JToken In arrayToken
                        canonicalArray.Add(CanonicalizeRetryCorrelationToken(item))
                    Next
                    Return canonicalArray
                End If

                Return token.DeepClone()
            End Function

            Private Shared Function ResolveFailureRecoveryPolicy(toolErrorHandling As String,
                                                                 terminal As Boolean,
                                                                 skippedByPolicy As Boolean,
                                                                 returnedToParent As Boolean) As ToolFailureRecoveryPolicy
                Dim normalizedHandling As String = If(toolErrorHandling, "").Trim().ToLowerInvariant()
                Dim parentAlternativeAllowed As Boolean = skippedByPolicy AndAlso returnedToParent

                If terminal Then
                    If parentAlternativeAllowed Then
                        Return ToolFailureRecoveryPolicy.DifferentAlternativeSuccessOnly
                    End If
                    Return ToolFailureRecoveryPolicy.NoAutomaticRecovery
                End If

                If parentAlternativeAllowed Then
                    Return ToolFailureRecoveryPolicy.CompatibleAlternativeSuccessAllowed
                End If

                Select Case normalizedHandling
                    Case "retry"
                        Return ToolFailureRecoveryPolicy.SameToolSuccessOnly
                    Case "abort"
                        Return ToolFailureRecoveryPolicy.NoAutomaticRecovery
                    Case Else
                        ' The host treats an empty/unknown ToolErrorHandling value as skip.
                        ' Mirror that behavior here so the sequencing state machine and physical
                        ' dispatcher cannot disagree about whether an alternative path is allowed.
                        Return ToolFailureRecoveryPolicy.CompatibleAlternativeSuccessAllowed
                End Select
            End Function

            Private Function GetLatestUnresolvedFailure() As ToolFailureRecord
                Dim index As Integer = GetLatestUnresolvedFailureIndex()
                If index < 0 Then
                    Return Nothing
                End If
                Return UnresolvedToolFailures(index)
            End Function

            Private Function GetLatestUnresolvedFailureIndex() As Integer
                If UnresolvedToolFailures Is Nothing OrElse UnresolvedToolFailures.Count = 0 Then
                    Return -1
                End If

                Dim bestIndex As Integer = -1
                Dim bestSequence As Long = Long.MinValue
                For i As Integer = 0 To UnresolvedToolFailures.Count - 1
                    Dim candidate As ToolFailureRecord = UnresolvedToolFailures(i)
                    If candidate Is Nothing Then Continue For
                    If bestIndex < 0 OrElse candidate.Sequence > bestSequence Then
                        bestIndex = i
                        bestSequence = candidate.Sequence
                    End If
                Next
                Return bestIndex
            End Function

            Friend Sub ProjectLatestUnresolvedFailure()
                Dim latest As ToolFailureRecord = GetLatestUnresolvedFailure()
                If latest Is Nothing Then
                    HasUnresolvedToolFailure = False
                    Return
                End If

                HasUnresolvedToolFailure = True
                LastToolName = If(latest.ToolName, "")
                LastErrorCode = If(latest.ErrorCode, "")
                LastErrorMessage = If(latest.ErrorMessage, "")
                LastFailureSkippedByPolicy = latest.SkippedByPolicy
                LastFailureReturnedToParent = latest.ReturnedToParent
                LastFailureTerminal = latest.Terminal
                LastFailureRecoveryPolicy = latest.RecoveryPolicy
                LastFailureToolClassification = latest.ToolClassification
                LastFailureToolErrorHandling = If(latest.ToolErrorHandling, "")
                LastFailureRecoveredByToolCall = False
                LastFailureHandledByBlockedFinal = False
                LastFailureUltimatelyFatal = False
                RecoveryToolName = ""
            End Sub

            Private Shared Function IsRecoveryNeutralAdministrativeTool(toolName As String) As Boolean
                If System.String.IsNullOrWhiteSpace(toolName) Then
                    Return True
                End If

                Dim normalizedToolName As String = toolName.Trim()

                If normalizedToolName.StartsWith("memory_", System.StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If

                Return System.String.Equals(normalizedToolName, ToolLoaderTool.LoaderToolName, System.StringComparison.OrdinalIgnoreCase) OrElse
                       System.String.Equals(normalizedToolName, CapabilityRoutingTool.ResolverToolName, System.StringComparison.OrdinalIgnoreCase) OrElse
                       System.String.Equals(normalizedToolName, "report_progress", System.StringComparison.OrdinalIgnoreCase) OrElse
                       System.String.Equals(normalizedToolName, "tool_describe", System.StringComparison.OrdinalIgnoreCase) OrElse
                       System.String.Equals(normalizedToolName, "context_compact", System.StringComparison.OrdinalIgnoreCase) OrElse
                       System.String.Equals(normalizedToolName, "context_expand", System.StringComparison.OrdinalIgnoreCase)
            End Function
        End Class

        Private Shared Function BuildRetryInvariantKey(toolName As String, recoveryScopeKey As String) As String
            Dim normalizedToolName As String = If(toolName, "").Trim().ToLowerInvariant()
            Dim normalizedScopeKey As String = If(recoveryScopeKey, "").Trim()
            If normalizedScopeKey = "" Then
                Return normalizedToolName
            End If
            Return normalizedToolName & ChrW(&H1F) & normalizedScopeKey
        End Function

        Private Shared ReadOnly RetryInvariantArgumentNames As String() = {"design_name", "template_attachment_name", "document_type", "document_language", "organization"}

        ''' <summary>
        ''' Preserves named artifact-fidelity constraints after a failed tool call until that same
        ''' tool later succeeds. This logic is deliberately tool-agnostic and is consumed by both
        ''' Outlook and Word dispatchers. A missing invariant is restored; an attempted replacement
        ''' is rejected. Intervening recovery/helper tools do not silently release the constraint.
        ''' </summary>
        Public Shared Function EnforceRetryInvariantArguments(toolName As String,
                                                               arguments As System.Collections.Generic.IDictionary(Of String, Object),
                                                               runState As ToolingRunState,
                                                               ByRef restoredSummary As String,
                                                               ByRef validationError As String) As Boolean
            restoredSummary = ""
            validationError = ""
            If runState Is Nothing OrElse arguments Is Nothing Then Return True

            Dim normalizedToolName As String = If(toolName, "").Trim()
            If normalizedToolName = "" Then Return True

            If runState.RetryInvariantArgumentsByTool Is Nothing Then
                runState.RetryInvariantArgumentsByTool =
                    New System.Collections.Generic.Dictionary(Of String, System.Collections.Generic.Dictionary(Of String, String))(System.StringComparer.Ordinal)
            End If

            Dim recoveryScopeKey As String = ResolveExplicitRecoveryScopeKey(arguments)

            ' Once the host has explicitly opened a bounded alternative/full-path recovery,
            ' the exhausted tool itself is no longer a valid recovery action. This prohibition
            ' is deliberately scope-agnostic because callers may issue a fresh operation/task id
            ' on every attempt. Allowing the same failed tool under a new scope would bypass the
            ' circuit breaker and turn a bounded alternative recovery into another retry loop.
            If runState.HasUnresolvedToolFailure AndAlso runState.UnresolvedToolFailures IsNot Nothing Then
                For Each failure As ToolFailureRecord In runState.UnresolvedToolFailures
                    If failure Is Nothing Then Continue For
                    If failure.CrossScopeAlternativeRecoveryAllowed AndAlso
                       failure.RecoveryPolicy = ToolFailureRecoveryPolicy.DifferentAlternativeSuccessOnly AndAlso
                       System.String.Equals(failure.ToolName, normalizedToolName, System.StringComparison.OrdinalIgnoreCase) Then

                        validationError =
                            "The previous execution path for tool '" & normalizedToolName &
                            "' exhausted its retry budget. Bounded alternative recovery requires a materially different tool/path; " &
                            "changing operation/task identifiers does not make the exhausted tool eligible again."
                        Return False
                    End If
                Next
            End If

            Dim retryInvariantKey As String = BuildRetryInvariantKey(normalizedToolName, recoveryScopeKey)
            Dim captured As System.Collections.Generic.Dictionary(Of String, String) = Nothing
            runState.RetryInvariantArgumentsByTool.TryGetValue(retryInvariantKey, captured)

            Dim isRetryOfFailedTool As Boolean = False
            If runState.HasUnresolvedToolFailure AndAlso runState.UnresolvedToolFailures IsNot Nothing Then
                For Each failure As ToolFailureRecord In runState.UnresolvedToolFailures
                    If failure Is Nothing Then Continue For
                    If System.String.Equals(failure.ToolName, normalizedToolName, System.StringComparison.OrdinalIgnoreCase) AndAlso
                       RecoveryScopeKeysMatch(failure.RecoveryScopeKey, recoveryScopeKey) Then
                        isRetryOfFailedTool = True
                        Exit For
                    End If
                Next
            End If

            Dim hasPendingRetryFidelity As Boolean =
                runState.RetryInvariantPendingFailureTools IsNot Nothing AndAlso
                runState.RetryInvariantPendingFailureTools.Contains(retryInvariantKey)

            If (isRetryOfFailedTool OrElse hasPendingRetryFidelity) AndAlso captured IsNot Nothing AndAlso captured.Count > 0 Then
                Dim restored As New System.Collections.Generic.List(Of String)()
                For Each pair As System.Collections.Generic.KeyValuePair(Of String, String) In captured
                    If System.String.IsNullOrWhiteSpace(pair.Value) Then Continue For

                    Dim currentValue As String = GetRetryInvariantArgumentValue(arguments, pair.Key)
                    If System.String.IsNullOrWhiteSpace(currentValue) Then
                        arguments(pair.Key) = pair.Value
                        restored.Add(pair.Key & "=" & pair.Value)
                        Continue For
                    End If

                    If Not System.String.Equals(currentValue, pair.Value, System.StringComparison.OrdinalIgnoreCase) Then
                        validationError = "Retry attempted to replace required artifact-fidelity argument '" & pair.Key &
                                          "' value '" & pair.Value & "' with '" & currentValue &
                                          "'. The original value remains binding after the failed tool call."
                        Return False
                    End If
                Next

                If restored.Count > 0 Then restoredSummary = System.String.Join(", ", restored)
            End If

            ' Capture the invariant values actually carried by this attempt. The dictionary is
            ' per tool, so progress/reporting or a different recovery tool cannot erase a failed
            ' artifact tool's design/template constraint before its retry.
            Dim currentCapture As New System.Collections.Generic.Dictionary(Of String, String)(System.StringComparer.OrdinalIgnoreCase)
            For Each argumentName As String In RetryInvariantArgumentNames
                Dim value As String = GetRetryInvariantArgumentValue(arguments, argumentName)
                If value <> "" Then currentCapture(argumentName) = value
            Next

            If currentCapture.Count > 0 Then
                runState.RetryInvariantArgumentsByTool(retryInvariantKey) = currentCapture
            ElseIf Not isRetryOfFailedTool AndAlso Not hasPendingRetryFidelity Then
                runState.RetryInvariantArgumentsByTool.Remove(retryInvariantKey)
            End If

            Return True
        End Function

        ''' <summary>
        ''' Refreshes retry-fidelity capture after deterministic host-side argument resolution.
        ''' Use this when the host supplies a default design/template after the initial dispatcher
        ''' gate, so a later retry cannot silently replace that effective artifact choice.
        ''' </summary>
        Public Shared Sub CaptureRetryInvariantArguments(toolName As System.String,
                                                         arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object),
                                                         runState As ToolingRunState)
            If runState Is Nothing OrElse arguments Is Nothing Then Return
            Dim normalizedToolName As System.String = If(toolName, System.String.Empty).Trim()
            If normalizedToolName = "" Then Return

            If runState.RetryInvariantArgumentsByTool Is Nothing Then
                runState.RetryInvariantArgumentsByTool =
                    New System.Collections.Generic.Dictionary(Of String, System.Collections.Generic.Dictionary(Of String, String))(System.StringComparer.Ordinal)
            End If

            Dim recoveryScopeKey As System.String = ResolveExplicitRecoveryScopeKey(arguments)
            Dim retryInvariantKey As System.String = BuildRetryInvariantKey(normalizedToolName, recoveryScopeKey)
            Dim capture As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.OrdinalIgnoreCase)
            For Each argumentName As System.String In RetryInvariantArgumentNames
                Dim value As System.String = GetRetryInvariantArgumentValue(arguments, argumentName)
                If value <> "" Then capture(argumentName) = value
            Next

            If capture.Count > 0 Then
                runState.RetryInvariantArgumentsByTool(retryInvariantKey) = capture
            End If
        End Sub

        Private Shared Function GetRetryInvariantArgumentValue(arguments As System.Collections.Generic.IDictionary(Of String, Object),
                                                               argumentName As String) As String
            If arguments Is Nothing OrElse System.String.IsNullOrWhiteSpace(argumentName) Then Return ""

            For Each pair As System.Collections.Generic.KeyValuePair(Of String, Object) In arguments
                If Not System.String.Equals(pair.Key, argumentName, System.StringComparison.OrdinalIgnoreCase) Then Continue For
                If pair.Value Is Nothing Then Return ""
                Return pair.Value.ToString().Trim()
            Next
            Return ""
        End Function

        Public Shared Function FormatMemoryGroundingMode(mode As MemoryGroundingMode) As String
            Select Case mode
                Case MemoryGroundingMode.Required
                    Return "required"
                Case MemoryGroundingMode.OptionalMode
                    Return "optional"
                Case Else
                    Return "none"
            End Select
        End Function



        Public Shared Function BuildExecutionPlan(toolNames As IEnumerable(Of String)) As ToolBatchPlan
            Dim plan As New ToolBatchPlan()

            If toolNames Is Nothing Then
                Return plan
            End If

            Dim barrierReached As Boolean = False
            Dim index As Integer = 0

            For Each rawName In toolNames
                Dim toolName As String = If(rawName, "").Trim()
                Dim classification = ClassifyToolName(toolName)
                Dim isBarrier As Boolean = IsBarrierClassification(classification)

                Dim item As New PlannedToolCall With {
                    .Index = index,
                    .ToolName = toolName,
                    .Classification = classification,
                    .IsBarrier = isBarrier,
                    .WillExecute = Not barrierReached,
                    .SkipReason = ""
                }

                If barrierReached Then
                    item.SkipReason = "deferred_after_sequencing_barrier"
                End If

                plan.Calls.Add(item)

                If isBarrier Then
                    barrierReached = True
                End If

                index += 1
            Next

            Return plan
        End Function

        Public Shared Function ClassifyToolName(toolName As String) As ToolCallClassification
            If String.IsNullOrWhiteSpace(toolName) Then
                Return ToolCallClassification.Unknown
            End If

            Dim name As String = toolName.Trim().ToLowerInvariant()

            If name.StartsWith("agent_", StringComparison.Ordinal) Then
                Return ToolCallClassification.Agent
            End If

            If name.StartsWith("skill_", StringComparison.Ordinal) OrElse
               name.Equals("skill_use", StringComparison.Ordinal) Then
                Return ToolCallClassification.Skill
            End If

            If name.Equals("tool_loader", StringComparison.Ordinal) OrElse
               name.StartsWith("memory_", StringComparison.Ordinal) Then
                Return ToolCallClassification.Stateful
            End If

            If HasAnyPhrase(name, "make_dir", "mkdir", "rmdir") Then
                Return ToolCallClassification.Mutating
            End If

            If HasAnyToken(name,
                           "state", "session", "cursor", "next", "queue", "loader") Then
                Return ToolCallClassification.Stateful
            End If

            If HasAnyToken(name,
                           "write", "save", "create", "delete", "remove", "move", "rename", "copy",
                           "append", "insert", "update", "set", "put", "apply", "stage",
                           "download", "upload", "commit", "send", "post") Then
                Return ToolCallClassification.Mutating
            End If

            If HasAnyToken(name,
                           "read", "get", "list", "inventory", "search", "find",
                           "extract", "query", "lookup", "retrieve", "fetch", "inspect") Then
                Return ToolCallClassification.ReadOnlyIndependent
            End If

            Return ToolCallClassification.Unknown
        End Function

        Private Shared Function RecoveryScopeKeysMatch(left As String,
                                                       right As String) As Boolean
            Return System.String.Equals(
                If(left, "").Trim(),
                If(right, "").Trim(),
                System.StringComparison.Ordinal)
        End Function

        Public Shared Function ResolveExplicitRecoveryScopeKey(arguments As IDictionary(Of String, Object)) As String
            If arguments Is Nothing Then
                Return ""
            End If

            Dim value As Object = Nothing
            Dim operationId As String = ""
            If arguments.TryGetValue("operation_id", value) AndAlso value IsNot Nothing Then
                operationId = If(System.Convert.ToString(value), "").Trim()
                If operationId <> "" Then
                    Dim stepId As String = ""
                    value = Nothing
                    If arguments.TryGetValue("step_id", value) AndAlso value IsNot Nothing Then
                        stepId = If(System.Convert.ToString(value), "").Trim()
                    End If
                    If stepId = "" Then stepId = "__default__"
                    Return "operation:" & operationId & "|step:" & stepId
                End If
            End If

            value = Nothing
            Dim subAgentTaskId As String = ""
            If arguments.TryGetValue("subagent_task_id", value) AndAlso value IsNot Nothing Then
                subAgentTaskId = If(System.Convert.ToString(value), "").Trim()
                If subAgentTaskId <> "" Then
                    Return "subagent:" & subAgentTaskId
                End If
            End If

            Dim logicalDeliverableId As String = ""
            Dim outputSlotId As String = ""

            value = Nothing
            If arguments.TryGetValue("logical_deliverable_id", value) AndAlso value IsNot Nothing Then
                logicalDeliverableId = If(System.Convert.ToString(value), "").Trim()
            End If

            value = Nothing
            If arguments.TryGetValue("output_slot_id", value) AndAlso value IsNot Nothing Then
                outputSlotId = If(System.Convert.ToString(value), "").Trim()
            End If

            If logicalDeliverableId <> "" AndAlso outputSlotId <> "" Then
                Return "artifact:" & logicalDeliverableId & "|" & outputSlotId
            End If

            Return ""
        End Function

        Friend Shared Function ClassifyToolNameForRecovery(toolName As String) As ToolCallClassification
            Dim classification As ToolCallClassification = ClassifyToolName(toolName)
            If classification <> ToolCallClassification.Unknown Then
                Return classification
            End If

            Dim name As String = If(toolName, "").Trim().ToLowerInvariant()
            If name = "" Then
                Return ToolCallClassification.Unknown
            End If

            ' Recovery classification is deliberately broader than batching classification.
            ' It is used only to decide whether a successful *different* tool may satisfy a
            ' skipped failure; it never changes execution ordering or mutation barriers.
            If HasAnyToken(name,
                           "edit", "markup", "comment", "redact", "watermark", "merge", "split",
                           "convert", "fill", "complete", "execute", "interact", "publish", "respond", "reply") Then
                Return ToolCallClassification.Mutating
            End If

            If HasAnyToken(name,
                           "open", "browse", "navigate") Then
                Return ToolCallClassification.Stateful
            End If

            If HasAnyToken(name,
                           "snapshot", "observe", "describe", "preview", "summarize", "analyze",
                           "compare", "validate", "verify", "check", "cite") Then
                Return ToolCallClassification.ReadOnlyIndependent
            End If

            Return ToolCallClassification.Unknown
        End Function

        Public Shared Function IsBarrierClassification(classification As ToolCallClassification) As Boolean
            Select Case classification
                Case ToolCallClassification.ReadOnlyIndependent
                    Return False
                Case Else
                    Return True
            End Select
        End Function

        Public Shared Function ShouldBlockTextOnlyFinalization(runState As ToolingRunState,
                                                               retryCount As Integer,
                                                               maxRetryCount As Integer,
                                                               hasValidFinalAnswer As Boolean) As Boolean
            If retryCount < maxRetryCount Then
                Return False
            End If

            If runState IsNot Nothing AndAlso runState.HasUnresolvedToolFailure Then
                Return True
            End If

            Return Not hasValidFinalAnswer
        End Function

        Public Shared Function BuildBlockedResultPayload(errorCode As String,
                                                         phase As String,
                                                         message As String,
                                                         Optional lastToolName As String = "",
                                                         Optional lastToolErrorCode As String = "",
                                                         Optional lastToolErrorMessage As String = "",
                                                         Optional retryable As System.Nullable(Of Boolean) = Nothing) As String
            Dim errorObject As New JObject(
                New JProperty("code", If(errorCode, "")),
                New JProperty("phase", If(phase, "")),
                New JProperty("message", If(message, "")))

            If retryable.HasValue Then
                errorObject("retryable") = retryable.Value
            End If

            Dim obj As New JObject(
                New JProperty("status", "blocked"),
                New JProperty("error", errorObject))

            If Not String.IsNullOrWhiteSpace(lastToolName) OrElse
               Not String.IsNullOrWhiteSpace(lastToolErrorCode) OrElse
               Not String.IsNullOrWhiteSpace(lastToolErrorMessage) Then

                Dim lastTool As New JObject()

                If Not String.IsNullOrWhiteSpace(lastToolName) Then
                    lastTool("name") = lastToolName
                End If

                If Not String.IsNullOrWhiteSpace(lastToolErrorCode) Then
                    lastTool("errorCode") = lastToolErrorCode
                End If

                If Not String.IsNullOrWhiteSpace(lastToolErrorMessage) Then
                    lastTool("message") = lastToolErrorMessage
                End If

                obj("lastToolFailure") = lastTool
                errorObject("lastToolFailure") = lastTool.DeepClone()
            End If

            Return obj.ToString(Formatting.None)
        End Function


        Public Shared Function StripTaskStatusBlocksFromUserFacingText(text As String) As String
            Dim raw As String = If(text, "")
            If raw = "" Then
                Return ""
            End If

            ' TASK_STATUS is an egress contract for the parent answer, not a global
            ' scrub token. Preserve literal tags inside JSON/tool/agent payload data.
            Return TaskStatusFooterParser.Strip(raw).Trim()
        End Function

        Public Shared Function ExtractVisibleUserFacingText(text As String) As String
            Dim raw As String = If(text, "")
            If raw = "" Then
                Return ""
            End If

            Dim visible As String =
                Regex.Replace(
                    raw,
                    "<[^>]+>",
                    " ",
                    RegexOptions.IgnoreCase Or RegexOptions.Singleline Or RegexOptions.CultureInvariant)

            visible =
                Regex.Replace(
                    visible,
                    "\s+",
                    " ",
                    RegexOptions.CultureInvariant)

            Return visible.Trim()
        End Function

        Public Shared Function HasSubstantiveUserFacingText(text As String) As Boolean
            Dim visible As String = ExtractVisibleUserFacingText(text)

            If visible = "" Then
                Return False
            End If

            Return Regex.IsMatch(
                visible,
                "\p{L}",
                RegexOptions.CultureInvariant)
        End Function

        ''' <summary>
        ''' Determines whether a text is (in whole) a raw structured payload such as a JSON object
        ''' or array. Used to prevent raw tool/protocol content from being surfaced to the user as a
        ''' final answer. A leading Markdown code fence is tolerated. Returns True only when the entire
        ''' remaining content parses as JSON, so ordinary prose that merely contains braces is not
        ''' misclassified.
        ''' </summary>
        Public Shared Function LooksLikeRawStructuredPayload(text As String) As Boolean
            Dim raw As String = If(text, "").Trim()
            If raw = "" Then
                Return False
            End If

            If raw.StartsWith("```", StringComparison.Ordinal) Then
                Dim firstBreak As Integer = raw.IndexOf(vbLf, StringComparison.Ordinal)
                If firstBreak >= 0 Then
                    raw = raw.Substring(firstBreak + 1)
                End If
                If raw.EndsWith("```", StringComparison.Ordinal) Then
                    raw = raw.Substring(0, raw.Length - 3)
                End If
                raw = raw.Trim()
                If raw = "" Then
                    Return False
                End If
            End If

            Dim firstChar As Char = raw(0)
            If firstChar <> "{"c AndAlso firstChar <> "["c Then
                Return False
            End If

            Try
                JToken.Parse(raw)
                Return True
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Host-agnostic gate deciding whether a final response is safe to present to the end user.
        ''' A response is presentable only when, after stripping the TASK_STATUS footer, it contains
        ''' substantive natural-language text and is not merely a raw structured (JSON) payload.
        ''' </summary>
        Public Shared Function IsUserPresentableFinalText(text As String) As Boolean
            Dim stripped As String = StripTaskStatusBlocksFromUserFacingText(If(text, ""))

            If Not HasSubstantiveUserFacingText(stripped) Then
                Return False
            End If

            If LooksLikeRawStructuredPayload(stripped) Then
                Return False
            End If

            Return True
        End Function


        ''' <summary>
        ''' Host-agnostic detector for provider/tool envelopes that must NEVER be surfaced as a
        ''' user-facing final answer. Returns True when the text is (in whole) a raw JSON payload,
        ''' or contains an embedded provider tool-call / function-call / function-response envelope.
        ''' Used by the forced-final and max-iteration acceptance gates so an envelope forces a
        ''' host-generated blocked result instead of being accepted as final text.
        ''' </summary>
        Public Shared Function ContainsProviderToolEnvelope(text As String) As Boolean
            Dim raw As String = If(text, "").Trim()
            If raw = "" Then
                Return False
            End If

            If raw.StartsWith("```", StringComparison.Ordinal) Then
                Dim firstBreak As Integer = raw.IndexOf(vbLf, StringComparison.Ordinal)
                If firstBreak >= 0 Then
                    raw = raw.Substring(firstBreak + 1)
                End If
                If raw.EndsWith("```", StringComparison.Ordinal) Then
                    raw = raw.Substring(0, raw.Length - 3)
                End If
                raw = raw.Trim()
                If raw = "" Then
                    Return False
                End If
            End If

            ' A whole-payload JSON object/array is a provider envelope, never user-facing.
            If LooksLikeRawStructuredPayload(raw) Then
                Return True
            End If

            ' Embedded provider tool-call / function-call / function-response markers.
            Return Regex.IsMatch(
                raw,
                "(""functionCall""|""function_call""|""functionResponse""|""function_response""|""tool_calls""|""tool_call""|""toolUse""|""tool_use"")",
                RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant)
        End Function


        Public Shared Function ParseStrictTaskStatus(text As String) As TaskStatusParseResult
            Dim result As New TaskStatusParseResult() With {
        .Status = TaskStatusKind.None
    }

            Dim trimmedText As String = If(text, "")
            Dim trimmedEnd As String = trimmedText.TrimEnd()

            If trimmedEnd = "" Then
                result.FailureReason = "empty_response"
                Return result
            End If

            Dim location As TaskStatusFooterParser.TaskStatusEnvelopeLocation =
                TaskStatusFooterParser.LocateTrailingEnvelope(trimmedEnd)

            result.FooterCount = If(location Is Nothing, 0, location.FooterCount)

            If location Is Nothing OrElse Not location.IsPresent Then
                result.FailureReason = "missing_task_status"
                Return result
            End If

            result.IsPresent = True

            If location.FooterCount <> 1 Then
                result.FailureReason = "multiple_task_status"
                Return result
            End If

            Dim jsonText As String = If(location.RawBody, System.String.Empty).Trim()
            result.FooterJson = jsonText
            result.TextBeforeFooter = trimmedEnd.Substring(0, location.StartIndex).TrimEnd()

            Try
                Dim obj As JObject = JObject.Parse(jsonText)
                Dim statusText As String = If(obj.Value(Of String)("status"), "").Trim().ToLowerInvariant()
                Dim rawReasonText As String = If(obj.Value(Of String)("reason"), "")
                Dim memoryGroundingScopeText As String = If(obj.Value(Of String)("memoryGroundingScope"), "").Trim().ToLowerInvariant()

                If rawReasonText.Trim() = "" Then
                    result.FailureReason = "task_status_missing_reason"
                    Return result
                End If

                If memoryGroundingScopeText <> "" AndAlso memoryGroundingScopeText <> "subset" Then
                    result.FailureReason = "task_status_invalid_memory_grounding_scope"
                    Return result
                End If

                Select Case statusText
                    Case "complete"
                        result.Status = TaskStatusKind.Complete
                    Case "blocked"
                        result.Status = TaskStatusKind.Blocked
                    Case "continue"
                        result.Status = TaskStatusKind.ContinueTurn
                    Case Else
                        result.FailureReason = "task_status_invalid_status"
                        Return result
                End Select

                result.Reason = NormalizeFooterReason(rawReasonText, statusText)
                result.MemoryGroundingScope = memoryGroundingScopeText
                result.IsValid = True
                Return result
            Catch
                result.FailureReason = "malformed_task_status"
                Return result
            End Try
        End Function


        Public NotInheritable Class MemoryGroundingIntentDecision
            Public Property MemoryGroundingMode As MemoryGroundingMode = MemoryGroundingMode.None
            Public Property Reason As String = "invalid_classifier_output"
            Public Property ShouldExposeRecentMemoryStubs As Boolean
            Public Property ExplicitStoredMemoryRequired As Boolean
            Public Property IsValid As Boolean
        End Class

        Public Shared Function BuildMemoryGroundingIntentClassifierSystemPrompt() As String
            Return "Classify whether the assistant's next answer should be grounded in session memory or prior stored workflow results. " &
                "Decide ONLY the memory-grounding mode for the current task. " &
                "Do NOT rewrite, replace, narrow, reinterpret, or summarize away the current task. " &
                "Treat <LATEST_USER_REQUEST_RAW> as the authoritative latest user request. " &
                "<HOST_TASK_SUMMARY> is secondary host metadata only and must never replace or narrow <LATEST_USER_REQUEST_RAW>. " &
                "The ""reason"" field must explain only the memory-grounding decision, not restate or rewrite the task. " &
                "Return EXACTLY one raw JSON object and nothing else. " &
                "Do NOT use Markdown. Do NOT use code fences. Do NOT add explanations. Do NOT add surrounding text. " &
                "The output must be exactly one JSON object with exactly these fields: " &
                "{""memoryGroundingMode"":""none|optional|required"",""reason"":""short reason"",""shouldExposeRecentMemoryStubs"":true|false,""explicitStoredMemoryRequired"":true|false}. " &
                "Use ""required"" ONLY when the user's latest request explicitly requires an answer based on stored Memory, remembered stored content, prior saved results, or previous saved workflow outputs. " &
                "Do NOT use ""required"" merely because earlier stored context may be helpful, relevant, or convenient. " &
                "If stored Memory could help but is not explicitly demanded by the user, use ""optional"" instead. " &
                "If the request is a new task that does not explicitly require saved Memory or prior saved results, do not use ""required"". " &
                "Set ""explicitStoredMemoryRequired"":true ONLY when that explicit user demand is present. Otherwise set it to false. " &
                "Base the decision on semantic meaning, not on language-specific keywords or surface wording."
        End Function

        Public Shared Function BuildMemoryGroundingIntentClassifierUserPrompt(latestUserRequestRaw As String,
                                                                              Optional hostTaskSummary As String = "") As String
            Dim sb As New System.Text.StringBuilder()

            sb.AppendLine("[CLASSIFIER_TASK_CONTEXT]")
            sb.AppendLine("LATEST_USER_REQUEST_RAW is authoritative for this classification.")
            sb.AppendLine("<LATEST_USER_REQUEST_RAW>")
            sb.AppendLine(If(latestUserRequestRaw, ""))
            sb.AppendLine("</LATEST_USER_REQUEST_RAW>")

            If Not String.IsNullOrWhiteSpace(hostTaskSummary) Then
                sb.AppendLine("<HOST_TASK_SUMMARY>")
                sb.AppendLine(hostTaskSummary.Trim())
                sb.AppendLine("</HOST_TASK_SUMMARY>")
            End If

            sb.AppendLine("[/CLASSIFIER_TASK_CONTEXT]")
            Return sb.ToString().TrimEnd()
        End Function

        Public Shared Function ParseMemoryGroundingIntentClassifierDecision(raw As String) As MemoryGroundingIntentDecision
            Dim normalizedOutput As String = ""
            Dim parseError As String = ""
            Return ParseMemoryGroundingIntentClassifierDecision(raw, normalizedOutput, parseError)
        End Function

        Public Shared Function ParseMemoryGroundingIntentClassifierDecision(raw As String,
                                                                            ByRef normalizedOutput As String,
                                                                            ByRef parseError As String) As MemoryGroundingIntentDecision
            Dim result As New MemoryGroundingIntentDecision()

            normalizedOutput = NormalizeMemoryGroundingIntentClassifierOutput(raw)
            parseError = ""

            If String.IsNullOrWhiteSpace(normalizedOutput) Then
                parseError = "empty_classifier_output"
                Return result
            End If

            Try
                Dim obj As JObject = JObject.Parse(normalizedOutput)

                Dim parsedMode As MemoryGroundingMode
                If Not TryParseMemoryGroundingModeText(
                    If(obj.Value(Of String)("memoryGroundingMode"), ""),
                    parsedMode) Then
                    parseError = "invalid_memory_grounding_mode"
                    Return result
                End If

                Dim reasonToken As JToken = obj("reason")
                If reasonToken Is Nothing OrElse reasonToken.Type <> JTokenType.String Then
                    parseError = "missing_or_invalid_reason"
                    Return result
                End If

                Dim exposeToken As JToken = obj("shouldExposeRecentMemoryStubs")
                If exposeToken Is Nothing OrElse exposeToken.Type <> JTokenType.Boolean Then
                    parseError = "missing_or_invalid_shouldExposeRecentMemoryStubs"
                    Return result
                End If

                Dim explicitRequiredToken As JToken = obj("explicitStoredMemoryRequired")
                If explicitRequiredToken Is Nothing OrElse explicitRequiredToken.Type <> JTokenType.Boolean Then
                    parseError = "missing_or_invalid_explicitStoredMemoryRequired"
                    Return result
                End If

                result.MemoryGroundingMode = parsedMode
                result.Reason = reasonToken.Value(Of String)().Trim()
                If result.Reason = "" Then
                    result.Reason = "parsed_classifier_output"
                End If

                result.ShouldExposeRecentMemoryStubs = exposeToken.Value(Of Boolean)()
                result.ExplicitStoredMemoryRequired = explicitRequiredToken.Value(Of Boolean)()
                result.IsValid = True
                Return result
            Catch ex As Exception
                parseError = ex.Message
                Return result
            End Try
        End Function


        Private Shared Function NormalizeMemoryGroundingIntentClassifierOutput(raw As String) As String
            Dim trimmed As String = If(raw, "").Trim()
            If trimmed = "" Then
                Return ""
            End If

            Dim fencedMatch As Match =
                Regex.Match(
                    trimmed,
                    "^\s*```(?:[A-Za-z0-9_-]+)?\s*\r?\n(?<body>[\s\S]*?)\r?\n```\s*$",
                    RegexOptions.CultureInvariant)

            If fencedMatch.Success Then
                Return fencedMatch.Groups("body").Value.Trim()
            End If

            Return trimmed
        End Function

        Public Shared Function FormatMemoryGroundingStage(stage As MemoryGroundingStage) As String
            Select Case stage
                Case MemoryGroundingStage.ListRequired
                    Return "list_required"
                Case MemoryGroundingStage.GetRequired
                    Return "get_required"
                Case MemoryGroundingStage.FullMemoryAvailable
                    Return "full_memory_available"
                Case MemoryGroundingStage.NoRelevantMemory
                    Return "no_relevant_memory"
                Case MemoryGroundingStage.Blocked
                    Return "blocked"
                Case Else
                    Return "not_started"
            End Select
        End Function

        Public Shared Function IsMemoryGroundingRejectionReason(reason As String) As Boolean
            Select Case If(reason, "").Trim().ToLowerInvariant()
                Case MissingRequiredMemoryAccessCode,
                     MemoryListDoneButMemoryGetRequiredCode,
                     MemoryGetFailedCode,
                     NoRelevantMemoryAvailableCode,
                     PartialMemoryRetrievalRequiresSubsetDisclosureCode
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        Private Shared Function BuildDistinctMemoryKeyList(keys As IEnumerable(Of String)) As List(Of String)
            Dim result As New List(Of String)()

            If keys Is Nothing Then
                Return result
            End If

            For Each key In keys
                Dim normalized As String = If(key, "").Trim()
                If normalized = "" Then Continue For
                If Not result.Contains(normalized, StringComparer.OrdinalIgnoreCase) Then
                    result.Add(normalized)
                End If
            Next

            Return result
        End Function

        Private Shared Function TryParseMemoryGetKey(rawResponse As String, ByRef memoryKey As String) As Boolean
            memoryKey = ""

            Dim trimmed As String = If(rawResponse, "").Trim()
            If trimmed = "" Then
                Return False
            End If

            Try
                Dim obj As JObject = JObject.Parse(trimmed)
                memoryKey = If(obj.Value(Of String)("key"), "").Trim()
                Return memoryKey <> ""
            Catch
                Return False
            End Try
        End Function

        Private Shared Function GetMemoryKeysStillUnretrieved(runState As ToolingRunState) As List(Of String)
            If runState Is Nothing Then
                Return New List(Of String)()
            End If

            Dim suggested = BuildDistinctMemoryKeyList(runState.MemoryKeysSuggestedForGet)
            Dim retrieved = BuildDistinctMemoryKeyList(runState.MemoryKeysRetrievedThisTurn)

            Return suggested.
                Where(Function(key) Not retrieved.Contains(key, StringComparer.OrdinalIgnoreCase)).
                ToList()
        End Function

        Private Shared Function ShouldRecommendRetrievingAllListedKeys(runState As ToolingRunState) As Boolean
            If runState Is Nothing Then
                Return False
            End If

            Return runState.MemoryListEntryCount > 0 AndAlso
                   runState.MemoryListEntryCount <= RequiredMemoryGetAllThreshold
        End Function

        Private Shared Sub UpdateFinalAnswerSubsetState(runState As ToolingRunState)
            If runState Is Nothing Then
                Return
            End If

            Dim unretrieved = GetMemoryKeysStillUnretrieved(runState)

            runState.FinalAnswerBasedOnSubset =
                runState.MemoryGetCountThisTurn > 0 AndAlso
                unretrieved.Count > 0
        End Sub

        Private NotInheritable Class MemoryListEntryDescriptor
            Public Property Key As String
            Public Property Summary As String
            Public Property WorkflowId As String
            Public Property TrustedForRuntime As Boolean
            Public Property UpdatedAt As DateTime
            Public Property Tags As List(Of String)
        End Class

        Private Shared Function TokenizeMemorySelectionText(text As String) As List(Of String)
            If String.IsNullOrWhiteSpace(text) Then
                Return New List(Of String)()
            End If

            Return Regex.Matches(text.ToLowerInvariant(), "[\p{L}\p{Nd}_-]{3,}").
                Cast(Of Match)().
                Select(Function(m) m.Value.Trim()).
                Where(Function(s) s <> "").
                Distinct(StringComparer.OrdinalIgnoreCase).
                ToList()
        End Function

        Private Shared Function ScoreMemoryListEntry(entry As MemoryListEntryDescriptor,
                                                     currentWorkflowId As String,
                                                     latestUserRequestRaw As String) As Integer
            If entry Is Nothing Then
                Return Integer.MinValue
            End If

            Dim score As Integer = 0
            Dim normalizedWorkflowId As String = If(currentWorkflowId, "").Trim()

            If normalizedWorkflowId <> "" AndAlso
               If(entry.WorkflowId, "").Trim().Equals(normalizedWorkflowId, StringComparison.OrdinalIgnoreCase) Then
                score += 1000000
            End If

            If entry.TrustedForRuntime Then
                score += 10000
            End If

            Dim haystack As String =
                ((If(entry.Summary, "") & " " & String.Join(" ", If(entry.Tags, New List(Of String)()))).Trim()).
                ToLowerInvariant()

            For Each token As String In TokenizeMemorySelectionText(latestUserRequestRaw)
                If haystack.Contains(token) Then
                    score += 100
                End If
            Next

            Return score
        End Function

        Public Shared Function SelectDeterministicMemoryKeysForHostFollowUp(rawMemoryListResponse As String,
                                                                            currentWorkflowId As String,
                                                                            latestUserRequestRaw As String,
                                                                            Optional maxKeys As Integer = 3) As List(Of String)
            Dim result As New List(Of String)()
            Dim descriptors As New List(Of MemoryListEntryDescriptor)()

            Dim trimmed As String = If(rawMemoryListResponse, "").Trim()
            If trimmed = "" Then
                Return result
            End If

            Try
                Dim token As JToken = JToken.Parse(trimmed)
                Dim arr As JArray = TryCast(token, JArray)
                If arr Is Nothing Then
                    Return result
                End If

                For Each item As JToken In arr
                    Dim obj As JObject = TryCast(item, JObject)
                    If obj Is Nothing Then Continue For

                    Dim key As String = If(obj.Value(Of String)("key"), "").Trim()
                    If key = "" Then Continue For

                    Dim metadata As JObject = TryCast(obj("metadata"), JObject)
                    Dim tagsArray As JArray = TryCast(obj("tags"), JArray)
                    Dim tags As New List(Of String)()

                    If tagsArray IsNot Nothing Then
                        tags = tagsArray.
                            Select(Function(t As JToken) t.ToString().Trim()).
                            Where(Function(t As String) t <> "").
                            ToList()
                    End If

                    Dim workflowId As String = ""
                    Dim trustedForRuntime As Boolean = False

                    If metadata IsNot Nothing Then
                        workflowId = If(metadata.Value(Of String)("workflowId"), "").Trim()
                        trustedForRuntime = If(metadata.Value(Of Boolean?)("trustedForRuntime"), False)
                    End If

                    descriptors.Add(New MemoryListEntryDescriptor With {
                        .Key = key,
                        .Summary = If(obj.Value(Of String)("summary"), "").Trim(),
                        .WorkflowId = workflowId,
                        .TrustedForRuntime = trustedForRuntime,
                        .UpdatedAt = If(obj.Value(Of DateTime?)("updatedAt"), DateTime.MinValue),
                        .Tags = tags
                    })
                Next
            Catch
                Return result
            End Try

            If descriptors.Count = 0 Then
                Return result
            End If

            Dim ordered As List(Of MemoryListEntryDescriptor) =
                descriptors.
                    OrderByDescending(Function(entry) ScoreMemoryListEntry(entry, currentWorkflowId, latestUserRequestRaw)).
                    ThenByDescending(Function(entry) entry.UpdatedAt).
                    ThenBy(Function(entry) entry.Key, StringComparer.OrdinalIgnoreCase).
                    ToList()

            If ordered.Count <= RequiredMemoryGetAllThreshold Then
                Return ordered.Select(Function(entry) entry.Key).ToList()
            End If

            Dim normalizedWorkflowId As String = If(currentWorkflowId, "").Trim()

            If normalizedWorkflowId <> "" Then
                Dim workflowMatches As List(Of String) =
                    ordered.
                        Where(
                            Function(entry)
                                Return If(entry.WorkflowId, "").Trim().Equals(normalizedWorkflowId, StringComparison.OrdinalIgnoreCase)
                            End Function).
                        Select(Function(entry) entry.Key).
                        ToList()

                If workflowMatches.Count > 0 Then
                    Return workflowMatches
                End If
            End If

            Return ordered.
                Take(Math.Max(1, maxKeys)).
                Select(Function(entry) entry.Key).
                ToList()
        End Function

        Private Shared Function HasExplicitSubsetDisclosure(taskStatus As TaskStatusParseResult) As Boolean
            If taskStatus Is Nothing OrElse Not taskStatus.IsValid Then
                Return False
            End If

            Return taskStatus.MemoryGroundingScopeIsSubset
        End Function

        Private Shared Function TryParseMemoryGroundingModeText(value As String,
                                                                ByRef mode As MemoryGroundingMode) As Boolean
            Select Case If(value, "").Trim().ToLowerInvariant()
                Case "required"
                    mode = MemoryGroundingMode.Required
                    Return True
                Case "optional"
                    mode = MemoryGroundingMode.OptionalMode
                    Return True
                Case "none"
                    mode = MemoryGroundingMode.None
                    Return True
                Case Else
                    mode = MemoryGroundingMode.None
                    Return False
            End Select
        End Function


        Public Shared Function GetFinalMutationPrerequisiteFailureReason(
            runState As ToolingRunState,
            toolName As System.String,
            arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object)) As System.String

            If runState Is Nothing Then Return System.String.Empty
            If Not IsFinalArtifactMutationCall(arguments) Then Return System.String.Empty

            Dim missing As System.Collections.Generic.List(Of System.String) =
                runState.GetMissingRequiredSuccessfulToolsBeforeFinalMutation()

            If missing.Count = 0 Then Return System.String.Empty

            Return "final_mutation_missing_required_successful_tools:" &
                   System.String.Join(",", missing)
        End Function

        Private Shared Function IsFinalArtifactMutationCall(
            arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object)) As System.Boolean

            If arguments Is Nothing Then Return False

            Dim rawValue As System.Object = Nothing

            If TryGetArgumentValue(arguments, "artifact_state", rawValue) AndAlso rawValue IsNot Nothing Then
                If System.String.Equals(
                    rawValue.ToString().Trim(),
                    "final",
                    System.StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If
            End If

            rawValue = Nothing
            If TryGetArgumentValue(arguments, "artifact_delivery_intent", rawValue) AndAlso rawValue IsNot Nothing Then
                Dim deliveryIntent As System.String = rawValue.ToString().Trim()
                If System.String.Equals(deliveryIntent, "deliver_to_user", System.StringComparison.OrdinalIgnoreCase) OrElse
                   System.String.Equals(deliveryIntent, "deliver_and_persist", System.StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If
            End If

            rawValue = Nothing
            If TryGetArgumentValue(arguments, "expected_artifacts", rawValue) AndAlso rawValue IsNot Nothing Then
                Dim token As Newtonsoft.Json.Linq.JToken = TryCast(rawValue, Newtonsoft.Json.Linq.JToken)
                If token IsNot Nothing Then
                    If token.Type = Newtonsoft.Json.Linq.JTokenType.Array Then
                        Return DirectCast(token, Newtonsoft.Json.Linq.JArray).Count > 0
                    End If
                End If

                Dim enumerable As System.Collections.IEnumerable =
                    TryCast(rawValue, System.Collections.IEnumerable)
                If enumerable IsNot Nothing AndAlso Not TypeOf rawValue Is System.String Then
                    For Each ignored As System.Object In enumerable
                        Return True
                    Next
                End If
            End If

            Return False
        End Function

        Private Shared Function TryGetArgumentValue(
            arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object),
            key As System.String,
            ByRef value As System.Object) As System.Boolean

            value = Nothing
            If arguments Is Nothing OrElse System.String.IsNullOrWhiteSpace(key) Then Return False

            If arguments.TryGetValue(key, value) Then Return True

            For Each pair As System.Collections.Generic.KeyValuePair(Of System.String, System.Object) In arguments
                If System.String.Equals(pair.Key, key, System.StringComparison.OrdinalIgnoreCase) Then
                    value = pair.Value
                    Return True
                End If
            Next

            Return False
        End Function

        Public Shared Function ValidateActiveToolingTurn(responseText As String,
                                                         hasToolCalls As Boolean,
                                                         hasUnresolvedToolFailure As Boolean,
                                                         Optional runState As ToolingRunState = Nothing) As ActiveToolingTurnValidationResult
            Dim result As New ActiveToolingTurnValidationResult() With {
                .TurnKind = ActiveToolingTurnKind.InvalidTurn,
                .InvalidReason = "",
                .TaskStatus = Nothing
            }

            If hasToolCalls Then
                result.TurnKind = ActiveToolingTurnKind.ToolCallTurn
                Return result
            End If

            If String.IsNullOrWhiteSpace(responseText) Then
                result.InvalidReason = "empty_response"
                Return result
            End If

            If IsRawInternalJsonResponse(responseText) Then
                result.InvalidReason = "raw_internal_json"
                Return result
            End If

            Dim parsed As TaskStatusParseResult = ParseStrictTaskStatus(responseText)
            result.TaskStatus = parsed

            If Not parsed.IsPresent Then
                result.InvalidReason = parsed.FailureReason
                Return result
            End If

            If Not parsed.IsValid Then
                result.InvalidReason = parsed.FailureReason
                Return result
            End If

            If String.IsNullOrWhiteSpace(parsed.TextBeforeFooter) Then
                result.InvalidReason = "missing_user_facing_text"
                Return result
            End If

            If Not HasSubstantiveUserFacingText(parsed.TextBeforeFooter) Then
                result.InvalidReason = "non_user_facing_final_text"
                Return result
            End If

            Select Case parsed.Status
                Case TaskStatusKind.Complete
                    If hasUnresolvedToolFailure Then
                        result.InvalidReason = "complete_with_unresolved_tool_failure"
                        Return result
                    End If

                    Dim missingRequiredTools As System.Collections.Generic.List(Of System.String) =
                        If(runState Is Nothing,
                           New System.Collections.Generic.List(Of System.String)(),
                           runState.GetMissingRequiredSuccessfulTools())

                    If missingRequiredTools.Count > 0 Then
                        result.InvalidReason = "complete_missing_required_successful_tools:" & System.String.Join(",", missingRequiredTools)
                        Return result
                    End If

                    Dim memoryGroundingFailureReason As String =
                        GetRequiredMemoryGroundingFailureReason(
                            runState,
                            ActiveToolingTurnKind.FinalCompleteTurn,
                            parsed)

                    If runState IsNot Nothing Then
                        runState.FinalCompleteRejectedForMissingMemoryAccess = False
                        runState.FinalCompleteRejectedForPartialMemoryRetrieval = False
                    End If

                    If memoryGroundingFailureReason <> "" Then
                        If runState IsNot Nothing Then
                            runState.FinalCompleteRejectedForMissingMemoryAccess =
                                IsMemoryGroundingRejectionReason(memoryGroundingFailureReason)

                            runState.FinalCompleteRejectedForPartialMemoryRetrieval =
                                String.Equals(
                                    memoryGroundingFailureReason,
                                    PartialMemoryRetrievalRequiresSubsetDisclosureCode,
                                    StringComparison.OrdinalIgnoreCase)
                        End If

                        result.InvalidReason = memoryGroundingFailureReason
                        Return result
                    End If

                    Dim requestedDeliverableFailureReason As String =
                        GetRequestedDeliverableFailureReason(
                            runState,
                            ActiveToolingTurnKind.FinalCompleteTurn,
                            parsed)

                    If requestedDeliverableFailureReason <> "" Then
                        result.InvalidReason = requestedDeliverableFailureReason
                        Return result
                    End If

                    result.TurnKind = ActiveToolingTurnKind.FinalCompleteTurn
                    Return result

                Case TaskStatusKind.Blocked
                    result.TurnKind = ActiveToolingTurnKind.FinalBlockedTurn
                    Return result

                Case TaskStatusKind.ContinueTurn
                    result.InvalidReason = "task_status_continue_not_final"
                    Return result

                Case Else
                    result.InvalidReason = "invalid_turn"
                    Return result
            End Select
        End Function

        Public Shared Function RequiresToolEnabledRepair(invalidReason As String) As Boolean
            Return System.String.Equals(
                If(invalidReason, System.String.Empty).Trim(),
                "complete_with_unresolved_tool_failure",
                System.StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function GetLatestUnresolvedToolFailure(runState As ToolingRunState) As ToolFailureRecord
            If runState Is Nothing OrElse
               Not runState.HasUnresolvedToolFailure OrElse
               runState.UnresolvedToolFailures Is Nothing OrElse
               runState.UnresolvedToolFailures.Count = 0 Then

                Return Nothing
            End If

            Dim latest As ToolFailureRecord = Nothing
            For Each candidate As ToolFailureRecord In runState.UnresolvedToolFailures
                If candidate Is Nothing Then Continue For
                If latest Is Nothing OrElse candidate.Sequence > latest.Sequence Then
                    latest = candidate
                End If
            Next

            Return latest
        End Function

        Private Shared Function FormatRepairDiagnosticValue(value As String, Optional maxLength As Integer = 160) As String
            Dim normalized As String = If(value, System.String.Empty).Replace(ChrW(13), " ").Replace(ChrW(10), " ").Trim()
            If maxLength > 0 AndAlso normalized.Length > maxLength Then
                normalized = normalized.Substring(0, maxLength) & "..."
            End If
            Return normalized
        End Function

        Private Shared Function BuildUnresolvedToolFailureRepairPrompt(runState As ToolingRunState) As String
            Dim failure As ToolFailureRecord = GetLatestUnresolvedToolFailure(runState)

            Dim prompt As String =
                "REPAIR: Completion is currently not permitted because a prior tool failure remains unresolved. " &
                "Do not return status complete until the unresolved failure has actually been resolved by successful tool work. " &
                "A tool call may be retried, including with the same arguments, when another attempt may succeed. " &
                "If no permitted recovery path can succeed, return a truthful blocked explanation ending with exactly one valid " &
                "<TASK_STATUS>{""status"":""blocked"",""reason"":""no safe completion path""}</TASK_STATUS>."

            If failure Is Nothing Then
                Return prompt
            End If

            prompt &= " Current unresolved failure: tool=" & FormatRepairDiagnosticValue(failure.ToolName) &
                      "; errorCode=" & FormatRepairDiagnosticValue(failure.ErrorCode) &
                      "; recoveryPolicy=" & failure.RecoveryPolicy.ToString() &
                      "; terminal=" & failure.Terminal.ToString().ToLowerInvariant() & "."

            Dim recoveryScopeKey As String = FormatRepairDiagnosticValue(failure.RecoveryScopeKey)
            If recoveryScopeKey <> System.String.Empty Then
                prompt &= " recoveryScope=" & recoveryScopeKey & "."
            End If

            Select Case failure.RecoveryPolicy
                Case ToolFailureRecoveryPolicy.SameToolSuccessOnly
                    prompt &= " This failure is resolved automatically only by a later successful retry of the exact same operation step. Reuse the same operation_id and step_id when present; a new step_id is a continuation and does not clear this failure. Do not switch tools merely to clear this failure."
                Case ToolFailureRecoveryPolicy.CompatibleAlternativeSuccessAllowed
                    prompt &= " A later successful retry of the exact same operation step may resolve this failure immediately. A different step is a continuation, not a retry; only a host-recognized replacement path with verified outcome evidence may supersede the failed step."
                Case ToolFailureRecoveryPolicy.DifferentAlternativeSuccessOnly
                    prompt &= " The failed tool itself is not the permitted automatic recovery path; use only a materially different host-recognized compatible alternative if one is available."
                Case ToolFailureRecoveryPolicy.NoAutomaticRecovery
                    prompt &= " This failure has no automatic successful-tool recovery path. Do not claim completion while it remains unresolved; use a host-authorized recovery mechanism if one exists, otherwise return blocked."
            End Select

            Return prompt
        End Function

        Public Shared Function BuildActiveToolingRepairPrompt(Optional runState As ToolingRunState = Nothing,
                                                      Optional invalidReason As String = "") As String
            Dim normalizedInvalidReason As String = If(invalidReason, "").Trim().ToLowerInvariant()
            Dim prompt As String

            If normalizedInvalidReason.StartsWith("complete_missing_required_successful_tools:", System.StringComparison.Ordinal) Then
                Dim missingTools As System.String = normalizedInvalidReason.Substring("complete_missing_required_successful_tools:".Length)
                Return "The selected skill declares mandatory successful verification/tool steps that have not yet succeeded: " & missingTools & ". Do not finalize. Load/call those exact tools as prescribed by the skill, use their results, then continue the workflow. If a required tool genuinely cannot run, return a truthful blocked result rather than inventing the missing verification."
            End If

            Select Case normalizedInvalidReason
                Case "complete_with_unresolved_tool_failure"
                    Return BuildUnresolvedToolFailureRepairPrompt(runState)
                Case RequestedDeliverableSlotsIncompleteCode
                    Dim missingSlots As System.Collections.Generic.List(Of System.String) =
                        GetMissingExpectedDeliverableSlotKeys(runState)
                    Dim missingSummary As System.String =
                        If(missingSlots.Count > 0, System.String.Join(", ", missingSlots), "one or more declared output slots")
                    Return "REPAIR: The final response was rejected because these expected output slots are not yet satisfied: " &
                           missingSummary & ". Keep the COMPLETE expected_artifacts contract unchanged. Create or correctly bind only the missing final artifacts, then finalize again. Do not shrink, replace, or invent slot identifiers merely to pass finalization."
                Case RequestedDeliverableNotCreatedCode
                    Return "REPAIR: The task requires a created user deliverable, but no completion-safe final artifact is registered. Produce the requested artifact through an authorized producer tool, then finalize again. Do not claim completion from a path string or inspection result alone."
                Case "task_status_reason_too_long",
             "task_status_missing_reason",
             "malformed_task_status"
                    prompt =
                "REPAIR: Your previous TASK_STATUS footer was malformed. " &
                "Your next response must be EXACTLY one of: " &
                "(1) the next required tool call and nothing else; " &
                "(2) a user-facing final prose answer ending with exactly one valid <TASK_STATUS>{""status"":""complete"",""reason"":""answer ready""}</TASK_STATUS>; or " &
                "(3) a user-facing blocked explanation ending with exactly one valid <TASK_STATUS>{""status"":""blocked"",""reason"":""no safe completion path""}</TASK_STATUS>. " &
                "The reason must be a single very short plain phrase, ideally 2-6 words, with no line breaks, and no more than " &
                TaskStatusReasonMaxChars.ToString() & " characters."
                Case "non_user_facing_final_text"
                    prompt =
                "REPAIR: Your previous final text was not valid user-facing prose. " &
                "Your next response must be EXACTLY one of: " &
                "(1) the next required tool call and nothing else; " &
                "(2) a user-facing final prose answer ending with exactly one valid <TASK_STATUS>{""status"":""complete"",""reason"":""answer ready""}</TASK_STATUS>; or " &
                "(3) a user-facing blocked explanation ending with exactly one valid <TASK_STATUS>{""status"":""blocked"",""reason"":""no safe completion path""}</TASK_STATUS>. " &
                "The reason must be a single very short plain phrase, ideally 2-6 words, with no line breaks, and no more than " &
                TaskStatusReasonMaxChars.ToString() & " characters."
                Case Else
                    prompt =
                "REPAIR: Your previous turn was not valid for the active tooling contract. " &
                "Your next response must be EXACTLY one of: " &
                "(1) the next required tool call and nothing else; " &
                "(2) a user-facing final prose answer ending with exactly one valid <TASK_STATUS>{""status"":""complete"",""reason"":""answer ready""}</TASK_STATUS>; or " &
                "(3) a user-facing blocked explanation ending with exactly one valid <TASK_STATUS>{""status"":""blocked"",""reason"":""no safe completion path""}</TASK_STATUS>. " &
                "The reason must be a single very short plain phrase, ideally 2-6 words, with no line breaks, and no more than " &
                TaskStatusReasonMaxChars.ToString() & " characters."
            End Select

            If runState IsNot Nothing AndAlso
       runState.MemoryGroundingMode = MemoryGroundingMode.Required Then

                prompt &= " If the final answer relies only on a retrieved subset of listed Memory entries, include ""memoryGroundingScope"":""subset"" inside the TASK_STATUS JSON footer."
            End If

            Return prompt
        End Function

        Public Shared Sub NoteMemoryGroundingToolResult(runState As ToolingRunState,
                                                        toolName As String,
                                                        rawResponse As String,
                                                        succeeded As Boolean)
            If runState Is Nothing OrElse String.IsNullOrWhiteSpace(toolName) Then
                Return
            End If

            If runState.MemoryKeysSuggestedForGet Is Nothing Then
                runState.MemoryKeysSuggestedForGet = New List(Of String)()
            End If

            If runState.MemoryKeysRetrievedThisTurn Is Nothing Then
                runState.MemoryKeysRetrievedThisTurn = New List(Of String)()
            End If

            Select Case toolName.Trim().ToLowerInvariant()
                Case MemoryTools.ToolList
                    runState.MemoryListCalledThisTurn = True

                    Dim entryCount As Integer = 0
                    Dim memoryKeys As List(Of String) = Nothing

                    If succeeded AndAlso TryParseMemoryListMetadata(rawResponse, entryCount, memoryKeys) Then
                        runState.MemoryListEntryCount = entryCount
                        runState.MemoryKeysSuggestedForGet = BuildDistinctMemoryKeyList(memoryKeys)
                        runState.MemoryListReturnedNoEntriesThisTurn = (entryCount = 0)

                        If entryCount = 0 Then
                            runState.MemoryGroundingStage = MemoryGroundingStage.NoRelevantMemory
                            runState.MemoryGetRequiredAfterList = False
                        Else
                            runState.MemoryGroundingStage = MemoryGroundingStage.GetRequired
                            runState.MemoryGetRequiredAfterList = True
                        End If
                    Else
                        runState.MemoryListEntryCount = 0
                        runState.MemoryListReturnedNoEntriesThisTurn = False
                        runState.MemoryGroundingStage = MemoryGroundingStage.ListRequired
                        runState.MemoryGetRequiredAfterList = False
                    End If

                Case MemoryTools.ToolGet
                    runState.MemoryGetCalledThisTurn = True
                    runState.MemoryGetCountThisTurn += 1

                    Dim retrievedKey As String = ""
                    If succeeded AndAlso TryParseMemoryGetKey(rawResponse, retrievedKey) Then
                        If retrievedKey <> "" AndAlso
                           Not runState.MemoryKeysRetrievedThisTurn.Contains(retrievedKey, StringComparer.OrdinalIgnoreCase) Then
                            runState.MemoryKeysRetrievedThisTurn.Add(retrievedKey)
                        End If
                    End If

                    If succeeded AndAlso MemoryGetReturnedFullValue(rawResponse) Then
                        runState.FullMemoryValueAvailableThisTurn = True

                        Dim unretrieved = GetMemoryKeysStillUnretrieved(runState)

                        If unretrieved.Count = 0 Then
                            runState.MemoryGroundingStage = MemoryGroundingStage.FullMemoryAvailable
                            runState.MemoryGetRequiredAfterList = False
                        Else
                            runState.MemoryGroundingStage = MemoryGroundingStage.GetRequired
                            runState.MemoryGetRequiredAfterList = True
                        End If
                    Else
                        runState.MemoryGroundingStage = MemoryGroundingStage.Blocked
                    End If
            End Select

            UpdateFinalAnswerSubsetState(runState)
        End Sub

        Public Shared Function GetRequiredMemoryGroundingFailureReason(runState As ToolingRunState,
                                                               proposedTurnKind As ActiveToolingTurnKind,
                                                               Optional taskStatus As TaskStatusParseResult = Nothing) As String
            If proposedTurnKind <> ActiveToolingTurnKind.FinalCompleteTurn Then
                Return ""
            End If

            If runState Is Nothing Then
                Return ""
            End If

            If Not runState.IsRequiredMemoryGroundingEnforced Then
                Return ""
            End If

            If runState.MemoryListCalledThisTurn AndAlso runState.MemoryListReturnedNoEntriesThisTurn Then
                Return ""
            End If

            If runState.MemoryGetCalledThisTurn AndAlso
       runState.MemoryGroundingStage = MemoryGroundingStage.Blocked Then
                Return MemoryGetFailedCode
            End If

            If runState.MemoryListCalledThisTurn AndAlso runState.MemoryListEntryCount > 0 Then
                Dim unretrieved As List(Of String) = GetMemoryKeysStillUnretrieved(runState)

                If runState.MemoryGetCountThisTurn = 0 Then
                    Return MemoryListDoneButMemoryGetRequiredCode
                End If

                If unretrieved.Count = 0 Then
                    Return ""
                End If

                If runState.MemoryGetCountThisTurn > 0 Then
                    Return ""
                End If
            End If

            If runState.FullMemoryValueAvailableThisTurn Then
                Return ""
            End If

            Return MissingRequiredMemoryAccessCode
        End Function


        Public Shared Function IsRequiredMemoryGroundingSatisfied(runState As ToolingRunState,
                                                                  proposedTurnKind As ActiveToolingTurnKind) As Boolean
            Return GetRequiredMemoryGroundingFailureReason(runState, proposedTurnKind) = ""
        End Function

        Public Shared Function RequiresRequiredMemoryGroundingBeforeNonMemoryTool(runState As ToolingRunState,
                                                                         toolName As String) As Boolean
            If runState Is Nothing Then
                Return False
            End If

            If Not runState.IsRequiredMemoryGroundingEnforced Then
                Return False
            End If

            If MemoryTools.IsMemoryTool(toolName) Then
                Return False
            End If

            If runState.MemoryListCalledThisTurn AndAlso runState.MemoryListReturnedNoEntriesThisTurn Then
                Return False
            End If

            If runState.FullMemoryValueAvailableThisTurn Then
                Return False
            End If

            If Not runState.MemoryListCalledThisTurn Then
                Return True
            End If

            If runState.MemoryGroundingStage = MemoryGroundingStage.ListRequired OrElse
       runState.MemoryGroundingStage = MemoryGroundingStage.Blocked Then
                Return True
            End If

            If runState.MemoryListEntryCount > 0 AndAlso runState.MemoryGetCountThisTurn = 0 Then
                Return True
            End If

            Return False
        End Function


        Public Shared Function BuildRequiredMemoryGroundingRepairPrompt(Optional runState As ToolingRunState = Nothing) As String
            Dim genericPrompt As String =
        "Memory grounding is explicitly required for this run. Use memory_list and memory_get before any non-memory tool or final answer. If no relevant stored entries exist, you may continue without Memory."

            If runState Is Nothing OrElse Not runState.IsRequiredMemoryGroundingEnforced Then
                Return genericPrompt
            End If

            If Not runState.MemoryListCalledThisTurn OrElse
       runState.MemoryGroundingStage = MemoryGroundingStage.ListRequired Then
                Return "Memory grounding is explicitly required for this run. In THIS turn, call exactly one tool: memory_list. Do not call any non-memory tool. Do not finalize yet."
            End If

            If runState.MemoryListCalledThisTurn AndAlso runState.MemoryListReturnedNoEntriesThisTurn Then
                Return "No stored entries were available. You may continue without Memory, or return a short, understandable blocked message if the task still cannot be completed reliably."
            End If

            Dim unretrieved As List(Of String) = GetMemoryKeysStillUnretrieved(runState)
            Dim keyHints As IList(Of String) = runState.MemoryKeysSuggestedForGet

            If unretrieved IsNot Nothing AndAlso unretrieved.Count > 0 Then
                keyHints = unretrieved
            End If

            Dim keyPromptSuffix As String = BuildMemoryKeysPromptSuffix(keyHints)

            If runState.MemoryListEntryCount > 0 AndAlso runState.MemoryGetCountThisTurn = 0 Then
                Return "Memory grounding is explicitly required for this run. In THIS turn, call exactly one tool: memory_get for the most relevant stored entry before any other tool or final answer. Do not call any non-memory tool. Do not finalize yet." & keyPromptSuffix
            End If

            If runState.MemoryGroundingStage = MemoryGroundingStage.Blocked Then
                Return "Memory grounding is explicitly required for this run, but the stored content could not be loaded successfully. Retry with memory_get if a relevant key is available, or return a short, understandable blocked message. Do not call any non-memory tool until Memory is resolved." & keyPromptSuffix
            End If

            If unretrieved.Count > 0 Then
                Return "At least one stored entry was loaded. You may continue using the loaded Memory. If you finalize based on only part of the stored content, say clearly that the answer may be incomplete." & keyPromptSuffix
            End If

            Return genericPrompt
        End Function

        Public Shared Function BuildMemoryGroundingStateSummary(runState As ToolingRunState) As String
            If runState Is Nothing Then
                Return "memoryGroundingMode=none; memoryGroundingAuthority=none; memoryGroundingStage=not_started; shouldExposeRecentMemoryStubs=false; memoryListEntryCount=0; memoryGetCountThisTurn=0; memoryGetRequiredAfterList=false; memoryKeysSuggestedForGet=(none); memoryKeysRetrievedThisTurn=(none); memoryKeysStillUnretrieved=(none); memoryListCalledThisTurn=false; memoryGetCalledThisTurn=false; fullMemoryValueAvailableThisTurn=false; finalAnswerBasedOnSubset=false; FinalCompleteRejectedForMissingMemoryAccess=false; FinalCompleteRejectedForPartialMemoryRetrieval=false"
            End If

            Dim unretrieved As List(Of String) = GetMemoryKeysStillUnretrieved(runState)

            Return "memoryGroundingMode=" & FormatMemoryGroundingMode(runState.MemoryGroundingMode) & ";" &
                    " memoryGroundingAuthority=" & runState.MemoryGroundingAuthority.ToString().ToLowerInvariant() & ";" &
                    " memoryGroundingStage=" & FormatMemoryGroundingStage(runState.MemoryGroundingStage) & ";" &
                    " shouldExposeRecentMemoryStubs=" & If(runState.ShouldExposeRecentMemoryStubs, "true", "false") & ";" &
                    " memoryListEntryCount=" & runState.MemoryListEntryCount.ToString(Globalization.CultureInfo.InvariantCulture) & ";" &
                    " memoryGetCountThisTurn=" & runState.MemoryGetCountThisTurn.ToString(Globalization.CultureInfo.InvariantCulture) & ";" &
                    " memoryGetRequiredAfterList=" & If(runState.MemoryGetRequiredAfterList, "true", "false") & ";" &
                    " memoryKeysSuggestedForGet=" & BuildMemoryKeysSummary(runState.MemoryKeysSuggestedForGet) & ";" &
                    " memoryKeysRetrievedThisTurn=" & BuildMemoryKeysSummary(runState.MemoryKeysRetrievedThisTurn) & ";" &
                    " memoryKeysStillUnretrieved=" & BuildMemoryKeysSummary(unretrieved) & ";" &
                    " memoryListCalledThisTurn=" & If(runState.MemoryListCalledThisTurn, "true", "false") & ";" &
                    " memoryGetCalledThisTurn=" & If(runState.MemoryGetCalledThisTurn, "true", "false") & ";" &
                    " fullMemoryValueAvailableThisTurn=" & If(runState.FullMemoryValueAvailableThisTurn, "true", "false") & ";" &
                    " finalAnswerBasedOnSubset=" & If(runState.FinalAnswerBasedOnSubset, "true", "false") & ";" &
                    " FinalCompleteRejectedForMissingMemoryAccess=" & If(runState.FinalCompleteRejectedForMissingMemoryAccess, "true", "false") & ";" &
                    " FinalCompleteRejectedForPartialMemoryRetrieval=" & If(runState.FinalCompleteRejectedForPartialMemoryRetrieval, "true", "false")
        End Function

        Private Shared Function TryParseMemoryListMetadata(rawResponse As String,
                                                           ByRef entryCount As Integer,
                                                           ByRef memoryKeys As List(Of String)) As Boolean
            entryCount = 0
            memoryKeys = New List(Of String)()

            Dim trimmed As String = If(rawResponse, "").Trim()
            If trimmed = "" Then
                Return False
            End If

            Try
                Dim token As JToken = JToken.Parse(trimmed)
                Dim arr As JArray = TryCast(token, JArray)
                If arr Is Nothing Then
                    Return False
                End If

                entryCount = arr.Count

                For Each item As JToken In arr
                    Dim obj As JObject = TryCast(item, JObject)
                    If obj Is Nothing Then Continue For

                    Dim key As String = If(obj.Value(Of String)("key"), "").Trim()
                    If key <> "" Then
                        memoryKeys.Add(key)
                    End If
                Next

                Return True
            Catch
                Return False
            End Try
        End Function

        Private Shared Function BuildMemoryKeysSummary(memoryKeys As IList(Of String)) As String
            If memoryKeys Is Nothing OrElse memoryKeys.Count = 0 Then
                Return "(none)"
            End If

            Return String.Join(", ", memoryKeys)
        End Function

        Private Shared Function BuildMemoryKeysPromptSuffix(memoryKeys As IList(Of String)) As String
            If memoryKeys Is Nothing OrElse memoryKeys.Count = 0 Then
                Return ""
            End If

            Return " Available keys: " & String.Join(", ", memoryKeys) & "."
        End Function

        Private Shared Function MemoryListHasNoEntries(rawResponse As String) As Boolean
            Dim trimmed As String = If(rawResponse, "").Trim()

            If trimmed = "" Then
                Return False
            End If

            If trimmed = "[]" Then
                Return True
            End If

            Try
                Dim token As JToken = JToken.Parse(trimmed)

                If TypeOf token Is JArray Then
                    Return DirectCast(token, JArray).Count = 0
                End If

                Dim obj As JObject = TryCast(token, JObject)
                If obj Is Nothing Then
                    Return False
                End If

                For Each propertyName In New String() {"items", "entries", "results"}
                    Dim child As JToken = obj(propertyName)

                    If TypeOf child Is JArray Then
                        Return DirectCast(child, JArray).Count = 0
                    End If
                Next
            Catch
            End Try

            Return False
        End Function

        Private Shared Function MemoryGetReturnedFullValue(rawResponse As String) As Boolean
            Dim trimmed As String = If(rawResponse, "").Trim()

            If trimmed = "" Then
                Return False
            End If

            Try
                Dim obj As JObject = TryCast(JToken.Parse(trimmed), JObject)
                If obj Is Nothing Then
                    Return False
                End If

                Dim valueToken As JToken = obj("value")
                Return valueToken IsNot Nothing AndAlso valueToken.Type <> JTokenType.Null
            Catch
                Return False
            End Try
        End Function


        Public Shared Function ClassifyToolFailureCategory(toolName As System.String,
                                                           errorCode As System.String,
                                                           errorMessage As System.String) As ToolFailureCategory
            Dim code As System.String = If(errorCode, System.String.Empty).Trim().ToLowerInvariant()
            Dim message As System.String = If(errorMessage, System.String.Empty).Trim().ToLowerInvariant()
            Dim tool As System.String = If(toolName, System.String.Empty).Trim().ToLowerInvariant()

            If code.Contains("artifact") OrElse code.Contains("deliverable") OrElse code.Contains("slot") Then
                Return ToolFailureCategory.ArtifactContract
            End If

            If code.Contains("schema") OrElse code.Contains("argument") OrElse code.Contains("validation") OrElse
               code = "invalid_tool_arguments" OrElse code = "no_change_applied" Then
                Return ToolFailureCategory.Validation
            End If

            If code.Contains("timeout") OrElse code.Contains("transport") OrElse code.Contains("http") OrElse
               code.Contains("network") OrElse message.Contains("timed out") OrElse message.Contains("connection") Then
                Return ToolFailureCategory.Transport
            End If

            If code.Contains("model") OrElse code.Contains("llm") OrElse code.Contains("empty_response") Then
                Return ToolFailureCategory.Model
            End If

            If tool.StartsWith("word_", System.StringComparison.Ordinal) OrElse
               code.Contains("openxml") OrElse code.Contains("document") OrElse message.Contains("word") Then
                Return ToolFailureCategory.DocumentProcessing
            End If

            Return ToolFailureCategory.Unknown
        End Function

        ''' <summary>
        ''' Builds one bounded metadata-only audit record for a tool-call phase or outcome.
        ''' The record intentionally excludes free-form content arguments and is safe to emit
        ''' before host preflight gates so rejected calls remain correlatable.
        ''' </summary>
        Private Const ToolCallAuditMaxDiagnosticCharacters As System.Int32 = 16384
        Private Const ToolCallAuditMaxIdentityCharacters As System.Int32 = 512
        Private Const ToolCallAuditMaxArgumentKeys As System.Int32 = 64
        Private Const ToolCallAuditMaxArgumentKeyCharacters As System.Int32 = 256
        Private Const ToolCallAuditMaxReferenceEntries As System.Int32 = 32
        Private Const ToolCallAuditMaxReferenceDepth As System.Int32 = 16
        Private Const ToolCallAuditMaxReferenceTraversalNodes As System.Int32 = 2048
        Private Const ToolCallAuditMaxReferenceCharacters As System.Int32 = 2048
        Private Const ToolCallAuditMaxFullPathCharacters As System.Int32 = 8192

        Public Shared Function BuildToolCallAuditDiagnostic(
            runId As System.String,
            callId As System.String,
            toolName As System.String,
            arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object),
            Optional recoveryScopeKey As System.String = "",
            Optional callStage As System.String = "",
            Optional outcome As System.String = "",
            Optional errorCode As System.String = "") As System.String

            Dim normalizedRunId As System.String = If(runId, System.String.Empty).Trim()
            Dim normalizedCallId As System.String = If(callId, System.String.Empty).Trim()
            Dim normalizedToolName As System.String = If(toolName, System.String.Empty).Trim()
            Dim normalizedCallStage As System.String = If(callStage, System.String.Empty).Trim()
            Dim normalizedOutcome As System.String = If(outcome, System.String.Empty).Trim()
            Dim normalizedErrorCode As System.String = If(errorCode, System.String.Empty).Trim()

            Try
                Dim parts As New System.Collections.Generic.List(Of System.String)()
                parts.Add("runId=" & FormatToolCallAuditBoundedText(normalizedRunId, ToolCallAuditMaxIdentityCharacters))
                parts.Add("callId=" & FormatToolCallAuditBoundedText(normalizedCallId, ToolCallAuditMaxIdentityCharacters))
                parts.Add("tool=" & FormatToolCallAuditBoundedText(normalizedToolName, ToolCallAuditMaxIdentityCharacters))
                If normalizedCallStage <> System.String.Empty Then
                    parts.Add("callStage=" & FormatToolCallAuditBoundedText(normalizedCallStage, ToolCallAuditMaxIdentityCharacters))
                End If
                If normalizedOutcome <> System.String.Empty Then
                    parts.Add("outcome=" & FormatToolCallAuditBoundedText(normalizedOutcome, ToolCallAuditMaxIdentityCharacters))
                End If
                If normalizedErrorCode <> System.String.Empty Then
                    parts.Add("errorCode=" & FormatToolCallAuditBoundedText(normalizedErrorCode, ToolCallAuditMaxIdentityCharacters))
                End If

                Dim normalizedRecoveryScope As System.String = If(recoveryScopeKey, System.String.Empty).Trim()
                If normalizedRecoveryScope = System.String.Empty Then
                    normalizedRecoveryScope = ResolveExplicitRecoveryScopeKey(arguments)
                End If
                If normalizedRecoveryScope <> System.String.Empty Then
                    parts.Add("recoveryScope=" & FormatToolCallAuditBoundedText(normalizedRecoveryScope, ToolCallAuditMaxIdentityCharacters))
                End If

                Dim operationId As System.String = System.String.Empty
                Dim stepId As System.String = System.String.Empty
                If arguments IsNot Nothing Then
                    Dim rawOperationId As System.Object = Nothing
                    If arguments.TryGetValue("operation_id", rawOperationId) AndAlso rawOperationId IsNot Nothing Then
                        operationId = rawOperationId.ToString().Trim()
                    End If
                    Dim rawStepId As System.Object = Nothing
                    If arguments.TryGetValue("step_id", rawStepId) AndAlso rawStepId IsNot Nothing Then
                        stepId = rawStepId.ToString().Trim()
                    End If
                End If
                If operationId <> System.String.Empty Then
                    parts.Add("operationId=" & FormatToolCallAuditBoundedText(operationId, ToolCallAuditMaxIdentityCharacters))
                    parts.Add("stepId=" & FormatToolCallAuditBoundedText(stepId, ToolCallAuditMaxIdentityCharacters))
                End If

                Dim operationIdentities As System.Collections.Generic.List(Of ExplicitOperationIdentity) =
                    ExplicitOperationRegistry.ExtractOperationIdentities(arguments)
                If operationIdentities IsNot Nothing AndAlso operationIdentities.Count > 0 Then
                    Dim operationSteps As New Newtonsoft.Json.Linq.JArray()
                    Dim operationLimit As System.Int32 = System.Math.Min(operationIdentities.Count, ToolCallAuditMaxReferenceEntries)
                    For index As System.Int32 = 0 To operationLimit - 1
                        Dim identity As ExplicitOperationIdentity = operationIdentities(index)
                        If identity Is Nothing Then Continue For
                        operationSteps.Add(New Newtonsoft.Json.Linq.JObject(
                            New Newtonsoft.Json.Linq.JProperty(
                                "operation_id",
                                BuildToolCallAuditBoundedToken(If(identity.OperationId, System.String.Empty), ToolCallAuditMaxIdentityCharacters)),
                            New Newtonsoft.Json.Linq.JProperty(
                                "step_id",
                                BuildToolCallAuditBoundedToken(If(identity.StepId, System.String.Empty), ToolCallAuditMaxIdentityCharacters))))
                    Next
                    If operationSteps.Count > 0 Then
                        parts.Add("operationSteps=" & operationSteps.ToString(Newtonsoft.Json.Formatting.None))
                    End If
                    If operationIdentities.Count > operationLimit Then
                        parts.Add("operationStepsOmitted=" &
                                  (operationIdentities.Count - operationLimit).ToString(System.Globalization.CultureInfo.InvariantCulture))
                    End If
                End If

                If arguments IsNot Nothing AndAlso arguments.Count > 0 Then
                    Dim argumentKeys As New System.Collections.Generic.List(Of System.String)(arguments.Keys)
                    argumentKeys.Sort(System.StringComparer.OrdinalIgnoreCase)
                    Dim argumentKeyArray As New Newtonsoft.Json.Linq.JArray()
                    Dim argumentKeyLimit As System.Int32 = System.Math.Min(argumentKeys.Count, ToolCallAuditMaxArgumentKeys)
                    For index As System.Int32 = 0 To argumentKeyLimit - 1
                        argumentKeyArray.Add(BuildToolCallAuditBoundedToken(argumentKeys(index), ToolCallAuditMaxArgumentKeyCharacters))
                    Next
                    parts.Add("argumentKeys=" & argumentKeyArray.ToString(Newtonsoft.Json.Formatting.None))
                    If argumentKeys.Count > argumentKeyLimit Then
                        parts.Add("argumentKeysOmitted=" &
                                  (argumentKeys.Count - argumentKeyLimit).ToString(System.Globalization.CultureInfo.InvariantCulture))
                    End If

                    Dim targetReferences As New System.Collections.Generic.List(Of System.String)()
                    Dim targetReferencesTruncated As System.Boolean = False
                    Dim targetReferencesDepthLimited As System.Boolean = False
                    Dim targetReferencesTraversalLimited As System.Boolean = False
                    Dim targetReferenceTraversalNodes As System.Int32 = 0
                    CollectToolCallAuditReferences(
                        Newtonsoft.Json.Linq.JToken.FromObject(arguments),
                        System.String.Empty,
                        targetReferences,
                        targetReferencesTruncated,
                        targetReferencesDepthLimited,
                        targetReferencesTraversalLimited,
                        targetReferenceTraversalNodes)
                    targetReferences.Sort(System.StringComparer.OrdinalIgnoreCase)
                    If targetReferences.Count > 0 Then
                        parts.Add("targetRefs=" & System.String.Join(",", targetReferences))
                    End If
                    If targetReferencesTruncated Then
                        parts.Add("targetRefsTruncated=true")
                    End If
                    If targetReferencesDepthLimited Then
                        parts.Add("targetRefsDepthLimited=true")
                    End If
                    If targetReferencesTraversalLimited Then
                        parts.Add("targetRefsTraversalLimited=true")
                    End If
                End If

                Dim diagnostic As System.String = "Tool call audit: " & System.String.Join("; ", parts)
                If diagnostic.Length <= ToolCallAuditMaxDiagnosticCharacters Then Return diagnostic

                Return BuildToolCallAuditOverflowDiagnostic(
                    normalizedRunId,
                    normalizedCallId,
                    normalizedToolName,
                    normalizedCallStage,
                    normalizedOutcome,
                    normalizedErrorCode,
                    diagnostic)
            Catch ex As System.Exception
                Try
                    Return "Tool call audit: runId=" & FormatToolCallAuditBoundedText(normalizedRunId, ToolCallAuditMaxIdentityCharacters) &
                           "; callId=" & FormatToolCallAuditBoundedText(normalizedCallId, ToolCallAuditMaxIdentityCharacters) &
                           "; tool=" & FormatToolCallAuditBoundedText(normalizedToolName, ToolCallAuditMaxIdentityCharacters) &
                           "; auditError=" & FormatToolCallAuditBoundedText(ex.GetType().Name, ToolCallAuditMaxIdentityCharacters)
                Catch fallbackEx As System.Exception
                    Return "Tool call audit: auditError=" & fallbackEx.GetType().Name
                End Try
            End Try
        End Function

        Private Shared Function BuildToolCallAuditOverflowDiagnostic(
            runId As System.String,
            callId As System.String,
            toolName As System.String,
            callStage As System.String,
            outcome As System.String,
            errorCode As System.String,
            fullDiagnostic As System.String) As System.String

            Dim parts As New System.Collections.Generic.List(Of System.String)()
            parts.Add("runId=" & FormatToolCallAuditBoundedText(runId, ToolCallAuditMaxIdentityCharacters))
            parts.Add("callId=" & FormatToolCallAuditBoundedText(callId, ToolCallAuditMaxIdentityCharacters))
            parts.Add("tool=" & FormatToolCallAuditBoundedText(toolName, ToolCallAuditMaxIdentityCharacters))
            If Not System.String.IsNullOrWhiteSpace(callStage) Then
                parts.Add("callStage=" & FormatToolCallAuditBoundedText(callStage, ToolCallAuditMaxIdentityCharacters))
            End If
            If Not System.String.IsNullOrWhiteSpace(outcome) Then
                parts.Add("outcome=" & FormatToolCallAuditBoundedText(outcome, ToolCallAuditMaxIdentityCharacters))
            End If
            If Not System.String.IsNullOrWhiteSpace(errorCode) Then
                parts.Add("errorCode=" & FormatToolCallAuditBoundedText(errorCode, ToolCallAuditMaxIdentityCharacters))
            End If
            parts.Add("detailsOmitted=true")
            parts.Add("fullChars=" & If(fullDiagnostic, System.String.Empty).Length.ToString(System.Globalization.CultureInfo.InvariantCulture))
            parts.Add("fullSha256=" & ComputeToolCallAuditSha256(If(fullDiagnostic, System.String.Empty)))
            Return "Tool call audit: " & System.String.Join("; ", parts)
        End Function

        Private Shared Function BuildToolCallAuditBoundedToken(
            value As System.String,
            maxCharacters As System.Int32) As Newtonsoft.Json.Linq.JToken

            Dim normalized As System.String = If(value, System.String.Empty)
            Dim effectiveLimit As System.Int32 = System.Math.Max(1, maxCharacters)
            If normalized.Length <= effectiveLimit Then
                Return New Newtonsoft.Json.Linq.JValue(normalized)
            End If

            Return New Newtonsoft.Json.Linq.JObject(
                New Newtonsoft.Json.Linq.JProperty("omitted", True),
                New Newtonsoft.Json.Linq.JProperty("chars", normalized.Length),
                New Newtonsoft.Json.Linq.JProperty("sha256", ComputeToolCallAuditSha256(normalized)))
        End Function

        Private Shared Function FormatToolCallAuditBoundedText(
            value As System.String,
            maxCharacters As System.Int32) As System.String

            Return BuildToolCallAuditBoundedToken(value, maxCharacters).ToString(Newtonsoft.Json.Formatting.None)
        End Function

        Private Shared Function IsToolCallAuditReferenceKey(argumentName As System.String) As System.Boolean
            Dim normalized As System.String = If(argumentName, System.String.Empty).Trim().ToLowerInvariant()
            If normalized = System.String.Empty Then Return False

            Select Case normalized
                Case "path", "paths", "target", "filename", "file_name", "output_filename",
                     "artifact_id", "logical_deliverable_id", "output_slot_id",
                     "supersedes_artifact_id", "result_ref"
                    Return True
            End Select

            Return normalized.EndsWith("_path", System.StringComparison.Ordinal) OrElse
                   normalized.EndsWith("_paths", System.StringComparison.Ordinal) OrElse
                   normalized.EndsWith("_filename", System.StringComparison.Ordinal) OrElse
                   normalized.EndsWith("_filenames", System.StringComparison.Ordinal) OrElse
                   normalized.EndsWith("_file", System.StringComparison.Ordinal) OrElse
                   normalized.EndsWith("_files", System.StringComparison.Ordinal) OrElse
                   normalized.EndsWith("_ref", System.StringComparison.Ordinal) OrElse
                   normalized.EndsWith("_refs", System.StringComparison.Ordinal)
        End Function

        Private Shared Sub CollectToolCallAuditReferences(
            token As Newtonsoft.Json.Linq.JToken,
            prefix As System.String,
            result As System.Collections.Generic.List(Of System.String),
            ByRef truncated As System.Boolean,
            ByRef depthLimited As System.Boolean,
            ByRef traversalLimited As System.Boolean,
            ByRef traversalNodes As System.Int32,
            Optional depth As System.Int32 = 0)

            If token Is Nothing OrElse result Is Nothing OrElse truncated OrElse traversalLimited Then Return
            If depth > ToolCallAuditMaxReferenceDepth Then
                depthLimited = True
                Return
            End If

            traversalNodes += 1
            If traversalNodes > ToolCallAuditMaxReferenceTraversalNodes Then
                traversalLimited = True
                Return
            End If

            Dim obj As Newtonsoft.Json.Linq.JObject = TryCast(token, Newtonsoft.Json.Linq.JObject)
            If obj IsNot Nothing Then
                For Each prop As Newtonsoft.Json.Linq.JProperty In obj.Properties()
                    traversalNodes += 1
                    If traversalNodes > ToolCallAuditMaxReferenceTraversalNodes Then
                        traversalLimited = True
                        Return
                    End If

                    Dim currentPath As System.String =
                        If(System.String.IsNullOrWhiteSpace(prefix), prop.Name, prefix & "." & prop.Name)

                    If IsToolCallAuditReferenceKey(prop.Name) Then
                        AddToolCallAuditReferenceValues(currentPath, prop.Name, prop.Value, result, truncated)
                        If truncated Then Return
                    End If

                    If prop.Value IsNot Nothing AndAlso
                       (prop.Value.Type = Newtonsoft.Json.Linq.JTokenType.Object OrElse
                        prop.Value.Type = Newtonsoft.Json.Linq.JTokenType.Array) Then
                        If depth >= ToolCallAuditMaxReferenceDepth Then
                            depthLimited = True
                        Else
                            CollectToolCallAuditReferences(
                                prop.Value,
                                currentPath,
                                result,
                                truncated,
                                depthLimited,
                                traversalLimited,
                                traversalNodes,
                                depth + 1)
                            If truncated OrElse traversalLimited Then Return
                        End If
                    End If
                Next
                Return
            End If

            Dim arr As Newtonsoft.Json.Linq.JArray = TryCast(token, Newtonsoft.Json.Linq.JArray)
            If arr Is Nothing Then Return

            For index As System.Int32 = 0 To arr.Count - 1
                traversalNodes += 1
                If traversalNodes > ToolCallAuditMaxReferenceTraversalNodes Then
                    traversalLimited = True
                    Return
                End If
                Dim currentPath As System.String = prefix & "[" &
                    index.ToString(System.Globalization.CultureInfo.InvariantCulture) & "]"
                If depth >= ToolCallAuditMaxReferenceDepth Then
                    If arr(index) IsNot Nothing AndAlso
                       (arr(index).Type = Newtonsoft.Json.Linq.JTokenType.Object OrElse
                        arr(index).Type = Newtonsoft.Json.Linq.JTokenType.Array) Then
                        depthLimited = True
                    End If
                Else
                    CollectToolCallAuditReferences(
                        arr(index),
                        currentPath,
                        result,
                        truncated,
                        depthLimited,
                        traversalLimited,
                        traversalNodes,
                        depth + 1)
                    If truncated OrElse traversalLimited Then Return
                End If
            Next
        End Sub

        Private Shared Sub AddToolCallAuditReferenceValues(
            path As System.String,
            key As System.String,
            value As Newtonsoft.Json.Linq.JToken,
            result As System.Collections.Generic.List(Of System.String),
            ByRef truncated As System.Boolean)

            If value Is Nothing OrElse result Is Nothing OrElse truncated Then Return

            Dim arr As Newtonsoft.Json.Linq.JArray = TryCast(value, Newtonsoft.Json.Linq.JArray)
            If arr IsNot Nothing Then
                For index As System.Int32 = 0 To arr.Count - 1
                    If Not TypeOf arr(index) Is Newtonsoft.Json.Linq.JValue Then Continue For
                    If result.Count >= ToolCallAuditMaxReferenceEntries Then
                        truncated = True
                        Return
                    End If
                    Dim itemPath As System.String = path & "[" &
                        index.ToString(System.Globalization.CultureInfo.InvariantCulture) & "]"
                    result.Add(
                        FormatToolCallAuditBoundedText(itemPath, ToolCallAuditMaxArgumentKeyCharacters) & "=" &
                        FormatToolCallAuditReferenceToken(arr(index), IsToolCallAuditPathKey(key)))
                Next
                Return
            End If

            If Not TypeOf value Is Newtonsoft.Json.Linq.JValue Then Return
            If result.Count >= ToolCallAuditMaxReferenceEntries Then
                truncated = True
                Return
            End If
            result.Add(
                FormatToolCallAuditBoundedText(path, ToolCallAuditMaxArgumentKeyCharacters) & "=" &
                FormatToolCallAuditReferenceToken(value, IsToolCallAuditPathKey(key)))
        End Sub

        Private Shared Function IsToolCallAuditPathKey(argumentName As System.String) As System.Boolean
            Dim normalized As System.String = If(argumentName, System.String.Empty).Trim().ToLowerInvariant()
            If normalized = "path" OrElse normalized = "paths" OrElse
               normalized = "filename" OrElse normalized = "file_name" OrElse
               normalized = "output_filename" Then
                Return True
            End If

            Return normalized.EndsWith("_path", System.StringComparison.Ordinal) OrElse
                   normalized.EndsWith("_paths", System.StringComparison.Ordinal) OrElse
                   normalized.EndsWith("_filename", System.StringComparison.Ordinal) OrElse
                   normalized.EndsWith("_filenames", System.StringComparison.Ordinal) OrElse
                   normalized.EndsWith("_file", System.StringComparison.Ordinal) OrElse
                   normalized.EndsWith("_files", System.StringComparison.Ordinal)
        End Function

        Private Shared Function FormatToolCallAuditReferenceToken(
            value As Newtonsoft.Json.Linq.JToken,
            preserveFullValue As System.Boolean) As System.String

            Dim serialized As System.String =
                If(value Is Nothing, "null", value.ToString(Newtonsoft.Json.Formatting.None))
            Dim maxCharacters As System.Int32 =
                If(preserveFullValue, ToolCallAuditMaxFullPathCharacters, ToolCallAuditMaxReferenceCharacters)
            If serialized.Length <= maxCharacters Then Return serialized

            Dim omission As New Newtonsoft.Json.Linq.JObject(
                New Newtonsoft.Json.Linq.JProperty("omitted", True),
                New Newtonsoft.Json.Linq.JProperty("chars", serialized.Length),
                New Newtonsoft.Json.Linq.JProperty("sha256", ComputeToolCallAuditSha256(serialized)))
            Return omission.ToString(Newtonsoft.Json.Formatting.None)
        End Function

        Private Shared Function ComputeToolCallAuditSha256(value As System.String) As System.String
            Using sha As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
                Dim bytes As System.Byte() = System.Text.Encoding.UTF8.GetBytes(If(value, System.String.Empty))
                Return System.BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", System.String.Empty).ToLowerInvariant()
            End Using
        End Function

        Public Shared Function BuildToolFailureDiagnostic(toolName As System.String,
                                                          errorCode As System.String,
                                                          errorMessage As System.String,
                                                          Optional recoveryScopeKey As System.String = "") As System.String
            Dim category As ToolFailureCategory = ClassifyToolFailureCategory(toolName, errorCode, errorMessage)
            Dim parts As New System.Collections.Generic.List(Of System.String)()
            parts.Add("category=" & category.ToString())
            parts.Add("tool=" & If(toolName, System.String.Empty).Trim())
            parts.Add("errorCode=" & If(errorCode, System.String.Empty).Trim())
            If Not System.String.IsNullOrWhiteSpace(recoveryScopeKey) Then
                parts.Add("recoveryScope=" & recoveryScopeKey.Trim())
            End If
            Return System.String.Join("; ", parts)
        End Function

        Private Shared Function BuildFailureRecoverySummary(failure As ToolFailureRecord,
                                                            recoveryToolName As System.String) As System.String
            If failure Is Nothing Then Return System.String.Empty

            Return "category=" & failure.Category.ToString() &
                   "; failedTool=" & If(failure.ToolName, System.String.Empty) &
                   "; errorCode=" & If(failure.ErrorCode, System.String.Empty) &
                   "; operation=" & If(failure.LogicalOperationKey, System.String.Empty) &
                   "; step=" & If(failure.StepKey, System.String.Empty) &
                   "; attempts=" & failure.AttemptCount.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                   "; lastAttempt=" & failure.LastAttemptSequence.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                   "; recoveryScope=" & If(failure.RecoveryScopeKey, System.String.Empty) &
                   "; recoveredBy=" & If(recoveryToolName, System.String.Empty)
        End Function

        Public Shared Function ConsumeRecoveredFailureSummary(runState As ToolingRunState) As System.String
            If runState Is Nothing Then Return System.String.Empty
            Dim value As System.String = If(runState.LastRecoveredFailureSummary, System.String.Empty)
            runState.LastRecoveredFailureSummary = System.String.Empty
            Return value
        End Function

        Public Shared Function GetMissingExpectedDeliverableSlotKeys(runState As ToolingRunState) As System.Collections.Generic.List(Of System.String)
            Dim result As New System.Collections.Generic.List(Of System.String)()
            If runState Is Nothing OrElse runState.ExpectedDeliverableSlots Is Nothing Then Return result

            For Each expected As ExpectedDeliverableSlot In runState.ExpectedDeliverableSlots
                If expected Is Nothing Then Continue For
                If runState.IsExpectedDeliverableSlotSatisfied(expected.LogicalDeliverableId, expected.OutputSlotId) Then Continue For
                result.Add(If(expected.LogicalDeliverableId, System.String.Empty) & "/" & If(expected.OutputSlotId, System.String.Empty))
            Next

            Return result
        End Function

        Public Shared Function BuildFinalizationDiagnostic(runState As ToolingRunState, invalidReason As System.String) As System.String
            Dim reason As System.String = If(invalidReason, System.String.Empty).Trim()
            If runState Is Nothing Then Return "phase=finalization; invalidReason=" & reason

            Dim parts As New System.Collections.Generic.List(Of System.String)()
            parts.Add("phase=finalization")
            parts.Add("invalidReason=" & reason)

            If System.String.Equals(reason, RequestedDeliverableSlotsIncompleteCode, System.StringComparison.OrdinalIgnoreCase) Then
                Dim missing As System.Collections.Generic.List(Of System.String) = GetMissingExpectedDeliverableSlotKeys(runState)
                If missing.Count > 0 Then parts.Add("missingSlots=" & System.String.Join(",", missing))
            End If

            If runState.HasUnresolvedToolFailure AndAlso runState.UnresolvedToolFailures IsNot Nothing Then
                Dim latest As ToolFailureRecord = GetLatestUnresolvedToolFailure(runState)
                If latest IsNot Nothing Then
                    parts.Add("failureCategory=" & latest.Category.ToString())
                    parts.Add("failedTool=" & If(latest.ToolName, System.String.Empty))
                    parts.Add("errorCode=" & If(latest.ErrorCode, System.String.Empty))
                    If Not System.String.IsNullOrWhiteSpace(latest.LogicalOperationKey) Then
                        parts.Add("operation=" & latest.LogicalOperationKey)
                    End If
                    If Not System.String.IsNullOrWhiteSpace(latest.StepKey) Then
                        parts.Add("step=" & latest.StepKey)
                    End If
                    If latest.AttemptCount > 0 Then
                        parts.Add("attempts=" & latest.AttemptCount.ToString(System.Globalization.CultureInfo.InvariantCulture))
                        parts.Add("lastAttempt=" & latest.LastAttemptSequence.ToString(System.Globalization.CultureInfo.InvariantCulture))
                    End If
                    If Not System.String.IsNullOrWhiteSpace(latest.RecoveryScopeKey) Then
                        parts.Add("recoveryScope=" & latest.RecoveryScopeKey)
                    End If
                End If
            End If

            Return System.String.Join("; ", parts)
        End Function

        Public Shared Function HasProducedUserDeliverable(runState As ToolingRunState) As Boolean
            If runState Is Nothing Then
                Return False
            End If

            ' Only authoritative, current, explicitly registered Finals count as a
            ' produced user deliverable. Legacy output_file_path/artifact-ref metadata,
            ' staging location, filenames, and other weak signals must never satisfy
            ' deliverable progress/completion logic.
            If runState.HasExpectedDeliverableContract Then
                If runState.ExpectedDeliverableSlots Is Nothing OrElse
                   runState.ExpectedDeliverableSlots.Count = 0 Then
                    Return False
                End If

                Return runState.HasAllExpectedDeliverableSlots
            End If

            Return runState.HasValidatedDeliverableForCompletion
        End Function

        Public Shared Function GetRequestedDeliverableFailureReason(runState As ToolingRunState,
                                                                   proposedTurnKind As ActiveToolingTurnKind,
                                                                   Optional taskStatus As TaskStatusParseResult = Nothing) As String
            If proposedTurnKind <> ActiveToolingTurnKind.FinalCompleteTurn Then
                Return ""
            End If

            If runState Is Nothing OrElse Not runState.RequestRequiresCreatedDeliverable Then
                Return ""
            End If

            If runState.ExpectedDeliverableSlots IsNot Nothing AndAlso
               runState.ExpectedDeliverableSlots.Count > 0 Then

                If runState.HasAllExpectedDeliverableSlots Then
                    Return ""
                End If

                Return RequestedDeliverableSlotsIncompleteCode
            End If

            If runState.HasValidatedDeliverableForCompletion Then
                Return ""
            End If

            Return RequestedDeliverableNotCreatedCode
        End Function

        Private Shared Sub ResetLastToolOutputMetadata(runState As ToolingRunState)
            If runState Is Nothing Then
                Return
            End If

            runState.LastToolProducesIntermediateData = False
            runState.LastToolProducesUserDeliverable = False
            runState.LastToolOutputArtifactRef = ""
            runState.LastToolOutputFilePath = ""
            runState.LastToolOutputMimeType = ""
            runState.LastToolOutputKind = ""
        End Sub

        Private Shared Sub NoteExplicitArtifactProtocolOwnedPaths(
            runState As ToolingRunState,
            rootObject As JObject,
            resultObject As JObject,
            outputFilePath As String,
            outputFiles As System.Collections.Generic.IEnumerable(Of String))

            If runState Is Nothing Then Return

            runState.RegisterExplicitArtifactProtocolOwnedPath(outputFilePath)

            If outputFiles IsNot Nothing Then
                For Each outputPath As String In outputFiles
                    runState.RegisterExplicitArtifactProtocolOwnedPath(outputPath)
                Next
            End If

            For Each container As JObject In New JObject() {rootObject, resultObject}
                If container Is Nothing Then Continue For

                Dim artifactsToken As JToken = container("artifacts")
                If artifactsToken Is Nothing OrElse artifactsToken.Type = JTokenType.Null Then Continue For

                If artifactsToken.Type = JTokenType.Array Then
                    For Each artifactToken As JToken In DirectCast(artifactsToken, JArray)
                        Dim artifactObject As JObject = TryCast(artifactToken, JObject)
                        If artifactObject IsNot Nothing Then
                            runState.RegisterExplicitArtifactProtocolOwnedPath(
                                If(artifactObject.Value(Of String)("path"), ""))
                        ElseIf artifactToken IsNot Nothing AndAlso artifactToken.Type = JTokenType.String Then
                            runState.RegisterExplicitArtifactProtocolOwnedPath(artifactToken.ToString())
                        End If
                    Next
                Else
                    Dim artifactObject As JObject = TryCast(artifactsToken, JObject)
                    If artifactObject IsNot Nothing Then
                        runState.RegisterExplicitArtifactProtocolOwnedPath(
                            If(artifactObject.Value(Of String)("path"), ""))
                    ElseIf artifactsToken.Type = JTokenType.String Then
                        runState.RegisterExplicitArtifactProtocolOwnedPath(artifactsToken.ToString())
                    End If
                End If
            Next
        End Sub

        Private Shared Function ExtractFirstBooleanValue(payload As JObject,
                                                        ParamArray keys() As String) As Boolean?
            If payload Is Nothing OrElse keys Is Nothing Then
                Return Nothing
            End If

            For Each key In keys
                If String.IsNullOrWhiteSpace(key) Then Continue For

                Dim token As JToken = payload(key)
                If token Is Nothing OrElse token.Type = JTokenType.Null Then Continue For

                If token.Type = JTokenType.Boolean Then
                    Return token.Value(Of Boolean)()
                End If

                Dim parsed As Boolean
                If Boolean.TryParse(token.ToString().Trim(), parsed) Then
                    Return parsed
                End If
            Next

            Return Nothing
        End Function

        Private Shared Sub NoteStructuredToolOutputMetadata(runState As ToolingRunState,
                                                            payload As JToken,
                                                            normalizedKind As String)
            If runState Is Nothing Then
                Return
            End If

            ResetLastToolOutputMetadata(runState)

            If payload Is Nothing Then
                Return
            End If

            Dim rootObject As JObject = TryCast(payload, JObject)
            Dim resultObject As JObject = Nothing

            If rootObject IsNot Nothing Then
                resultObject = TryCast(rootObject("result"), JObject)
            End If

            Dim explicitIntermediate As Boolean? =
                ExtractFirstBooleanValue(
                    rootObject,
                    "producesIntermediateData",
                    "produces_intermediate_data")

            If Not explicitIntermediate.HasValue Then
                explicitIntermediate =
                    ExtractFirstBooleanValue(
                        resultObject,
                        "producesIntermediateData",
                        "produces_intermediate_data")
            End If

            Dim explicitDeliverable As Boolean? =
                ExtractFirstBooleanValue(
                    rootObject,
                    "producesUserDeliverable",
                    "produces_user_deliverable")

            If Not explicitDeliverable.HasValue Then
                explicitDeliverable =
                    ExtractFirstBooleanValue(
                        resultObject,
                        "producesUserDeliverable",
                        "produces_user_deliverable")
            End If

            Dim createdStatus As Boolean? =
                ExtractFirstBooleanValue(
                    rootObject,
                    "created",
                    "saved",
                    "exported")

            If Not createdStatus.HasValue Then
                createdStatus =
                    ExtractFirstBooleanValue(
                        resultObject,
                        "created",
                        "saved",
                        "exported")
            End If

            Dim artifactRef As String =
                ExtractFirstStringValue(
                    rootObject,
                    "outputArtifactRef",
                    "output_artifact_ref",
                    "artifact_ref",
                    "output_reference",
                    "reference",
                    "state_reference")

            If String.IsNullOrWhiteSpace(artifactRef) Then
                artifactRef =
                    ExtractFirstStringValue(
                        resultObject,
                        "outputArtifactRef",
                        "output_artifact_ref",
                        "artifact_ref",
                        "output_reference",
                        "reference",
                        "state_reference")
            End If

            Dim explicitOutputFilePath As String =
                ExtractFirstStringValue(
                    rootObject,
                    "outputFilePath",
                    "output_file_path",
                    "output_path",
                    "file_path")

            If String.IsNullOrWhiteSpace(explicitOutputFilePath) Then
                explicitOutputFilePath =
                    ExtractFirstStringValue(
                        resultObject,
                        "outputFilePath",
                        "output_file_path",
                        "output_path",
                        "file_path")
            End If

            Dim genericPath As String =
                ExtractFirstStringValue(
                    rootObject,
                    "path")

            If String.IsNullOrWhiteSpace(genericPath) Then
                genericPath =
                    ExtractFirstStringValue(
                        resultObject,
                        "path")
            End If

            Dim outputFilePath As String =
                If(Not System.String.IsNullOrWhiteSpace(explicitOutputFilePath),
                   explicitOutputFilePath,
                   genericPath)

            Dim outputFileName As String =
                ExtractFirstStringValue(
                    rootObject,
                    "outputFileName",
                    "output_file_name",
                    "output_filename",
                    "file_name",
                    "filename")

            If String.IsNullOrWhiteSpace(outputFileName) Then
                outputFileName =
                    ExtractFirstStringValue(
                        resultObject,
                        "outputFileName",
                        "output_file_name",
                        "output_filename",
                        "file_name",
                        "filename")
            End If

            Dim outputFiles As New List(Of String)()

            For Each value As String In ExtractStringListValues(rootObject, "outputFiles", "output_files")
                AddDistinctString(outputFiles, value)
            Next

            For Each value As String In ExtractStringListValues(resultObject, "outputFiles", "output_files")
                AddDistinctString(outputFiles, value)
            Next

            If String.IsNullOrWhiteSpace(outputFilePath) AndAlso outputFiles.Count > 0 Then
                outputFilePath = outputFiles(0)
            End If

            If String.IsNullOrWhiteSpace(outputFilePath) AndAlso
               Not String.IsNullOrWhiteSpace(outputFileName) Then
                outputFilePath = outputFileName
            End If

            If String.IsNullOrWhiteSpace(artifactRef) AndAlso outputFiles.Count > 0 Then
                artifactRef = outputFiles(0)
            End If

            If String.IsNullOrWhiteSpace(artifactRef) AndAlso
               Not String.IsNullOrWhiteSpace(outputFileName) Then
                artifactRef = outputFileName
            End If

            Dim outputMimeType As String =
                ExtractFirstStringValue(
                    rootObject,
                    "outputMimeType",
                    "output_mime_type",
                    "mime_type",
                    "mime",
                    "content_type")

            If String.IsNullOrWhiteSpace(outputMimeType) Then
                outputMimeType =
                    ExtractFirstStringValue(
                        resultObject,
                        "outputMimeType",
                        "output_mime_type",
                        "mime_type",
                        "mime",
                        "content_type")
            End If

            Dim outputKind As String =
                ExtractFirstStringValue(
                    rootObject,
                    "outputKind",
                    "output_kind",
                    "kind",
                    "result_kind")

            If String.IsNullOrWhiteSpace(outputKind) Then
                outputKind =
                    ExtractFirstStringValue(
                        resultObject,
                        "outputKind",
                        "output_kind",
                        "kind",
                        "result_kind")
            End If

            If String.IsNullOrWhiteSpace(outputKind) Then
                outputKind = If(normalizedKind, "").Trim()
            End If

            ' A transport-successful mutation that applied zero changes is NOT a deliverable:
            ' it must neither flag deliverable production nor register an artifact, otherwise the
            ' completion gate (HasValidatedFinalDeliverable) would be falsely satisfied without a
            ' real file having been produced.
            Dim isZeroChangeResult As Boolean = IsZeroChangeOperationToken(rootObject, resultObject)

            ' Explicit producesUserDeliverable is authoritative and always honored. All other
            ' signals are "weak" (a generic created/saved/exported flag, an artifact reference, or
            ' an output path that may just echo a source 'path'). A weak signal may only infer a
            ' deliverable when the producing tool is actually capable of producing one. This stops
            ' read-only tools (e.g. text extract/search) from registering or being promoted to a
            ' forced deliverable merely because their result echoed the input file path.
            Dim hasExplicitDeliverableSignal As Boolean = explicitDeliverable.GetValueOrDefault(False)

            Dim toolClassification As ToolCallClassification =
                ClassifyToolName(runState.LastStructuredToolName)

            Dim genericPathMayInferDeliverable As Boolean =
                toolClassification <> ToolCallClassification.Mutating

            Dim hasWeakDeliverableSignal As Boolean =
                createdStatus.GetValueOrDefault(False) OrElse
                Not String.IsNullOrWhiteSpace(artifactRef) OrElse
                Not String.IsNullOrWhiteSpace(explicitOutputFilePath) OrElse
                outputFiles.Count > 0 OrElse
                (genericPathMayInferDeliverable AndAlso Not String.IsNullOrWhiteSpace(genericPath))

            Dim producingToolIsDeliverableCapable As Boolean =
                runState.IsDeliverableCapableTool(runState.LastStructuredToolName)

            Dim hasExplicitIntermediateSignal As Boolean =
                explicitIntermediate.GetValueOrDefault(False)

            Dim inferredDeliverable As Boolean =
                Not isZeroChangeResult AndAlso
                (hasExplicitDeliverableSignal OrElse
                 (Not hasExplicitIntermediateSignal AndAlso
                  hasWeakDeliverableSignal AndAlso
                  producingToolIsDeliverableCapable))

            Dim inferredIntermediate As Boolean =
                hasExplicitIntermediateSignal OrElse
                ((TypeOf payload Is JObject OrElse TypeOf payload Is JArray) AndAlso
                 Not inferredDeliverable)

            runState.LastToolProducesUserDeliverable = inferredDeliverable
            runState.LastToolProducesIntermediateData = inferredIntermediate
            runState.LastToolOutputArtifactRef = If(artifactRef, "")
            runState.LastToolOutputFilePath = If(outputFilePath, "")
            runState.LastToolOutputMimeType = If(outputMimeType, "")
            runState.LastToolOutputKind = If(outputKind, "")

            If inferredDeliverable Then
                runState.AnyUserDeliverableProducedThisRun = True
            End If

            ' Host-agnostic artifact registry.
            ' Explicit artifacts[] is authoritative for relationships such as
            ' logical output slots and supersession. Legacy output paths remain
            ' supported and are never heuristically merged.
            If Not isZeroChangeResult Then
                Dim explicitArtifactsDeclared As Boolean =
                    ArtifactDelivery.DeclaresExplicitArtifacts(rootObject, resultObject)

                If explicitArtifactsDeclared Then
                    ' Once a tool declares artifacts[], that protocol is authoritative.
                    ' A malformed/conflicting payload must remain unresolved; never turn
                    ' the same physical side effect into a path-only legacy deliverable.
                    NoteExplicitArtifactProtocolOwnedPaths(
                        runState,
                        rootObject,
                        resultObject,
                        outputFilePath,
                        outputFiles)

                    ArtifactDelivery.RegisterExplicitArtifacts(
                        runState,
                        rootObject,
                        resultObject,
                        runState.LastStructuredToolName)
                Else
                    runState.RegisterExistingDeliverableArtifact(
                        outputFilePath,
                        runState.LastStructuredToolName,
                        inferredDeliverable)

                    For Each producedPath As String In outputFiles
                        runState.RegisterExistingDeliverableArtifact(
                            producedPath,
                            runState.LastStructuredToolName,
                            inferredDeliverable)
                    Next
                End If
            End If
            If Not String.IsNullOrWhiteSpace(outputFilePath) Then
                runState.LastKnownOutputReference = outputFilePath
                runState.LastOutputPath = outputFilePath
                runState.LastStateFilePath = outputFilePath
            ElseIf Not String.IsNullOrWhiteSpace(artifactRef) Then
                runState.LastKnownOutputReference = artifactRef
            End If
        End Sub


        Private Shared Sub AddDistinctString(results As List(Of String), value As String)
            If results Is Nothing Then
                Return
            End If

            Dim normalized As String = If(value, "").Trim()
            If normalized = "" Then
                Return
            End If

            For Each existing As String In results
                If String.Equals(existing, normalized, StringComparison.OrdinalIgnoreCase) Then
                    Return
                End If
            Next

            results.Add(normalized)
        End Sub

        Private Shared Function ExtractStringListValues(payload As JObject,
                                                        ParamArray keys() As String) As List(Of String)
            Dim results As New List(Of String)()

            If payload Is Nothing OrElse keys Is Nothing Then
                Return results
            End If

            For Each key As String In keys
                If String.IsNullOrWhiteSpace(key) Then Continue For

                Dim token As JToken = payload(key)
                If token Is Nothing OrElse token.Type = JTokenType.Null Then Continue For

                If token.Type = JTokenType.String Then
                    AddDistinctString(results, token.ToString())
                    Continue For
                End If

                Dim arr As JArray = TryCast(token, JArray)
                If arr Is Nothing Then Continue For

                For Each item As JToken In arr
                    If item Is Nothing OrElse item.Type = JTokenType.Null Then Continue For
                    AddDistinctString(results, item.ToString())
                Next
            Next

            Return results
        End Function

        Private Shared Function TryGetStructuredDeliverableResult(responseText As String,
                                                                  ByRef rootObject As JObject,
                                                                  ByRef resultObject As JObject) As Boolean
            rootObject = Nothing
            resultObject = Nothing

            Dim raw As String = If(responseText, "").Trim()
            If raw = "" Then
                Return False
            End If

            Try
                rootObject = TryCast(JToken.Parse(raw), JObject)
                If rootObject Is Nothing Then
                    Return False
                End If

                resultObject = TryCast(rootObject("result"), JObject)
                Return True
            Catch
                Return False
            End Try
        End Function

        Public Shared Function ExtractCreatedDeliverableReferences(responseText As String) As List(Of String)
            Dim references As New List(Of String)()
            Dim rootObject As JObject = Nothing
            Dim resultObject As JObject = Nothing

            If Not TryGetStructuredDeliverableResult(responseText, rootObject, resultObject) Then
                Return references
            End If

            AddDistinctString(references,
                ExtractFirstStringValue(
                    rootObject,
                    "outputArtifactRef",
                    "output_artifact_ref",
                    "artifact_ref",
                    "output_reference",
                    "reference"))

            AddDistinctString(references,
                ExtractFirstStringValue(
                    resultObject,
                    "outputArtifactRef",
                    "output_artifact_ref",
                    "artifact_ref",
                    "output_reference",
                    "reference"))

            AddDistinctString(references,
                ExtractFirstStringValue(
                    rootObject,
                    "outputFilePath",
                    "output_file_path",
                    "output_path",
                    "file_path",
                    "path"))

            AddDistinctString(references,
                ExtractFirstStringValue(
                    resultObject,
                    "outputFilePath",
                    "output_file_path",
                    "output_path",
                    "file_path",
                    "path"))

            AddDistinctString(references,
                ExtractFirstStringValue(
                    rootObject,
                    "outputFileName",
                    "output_file_name",
                    "output_filename",
                    "file_name",
                    "filename"))

            AddDistinctString(references,
                ExtractFirstStringValue(
                    resultObject,
                    "outputFileName",
                    "output_file_name",
                    "output_filename",
                    "file_name",
                    "filename"))

            For Each value As String In ExtractStringListValues(rootObject, "outputFiles", "output_files")
                AddDistinctString(references, value)
            Next

            For Each value As String In ExtractStringListValues(resultObject, "outputFiles", "output_files")
                AddDistinctString(references, value)
            Next

            Return references
        End Function

        Public Shared Function IsSuccessfulDeliverableResult(responseText As String) As Boolean
            Dim rootObject As JObject = Nothing
            Dim resultObject As JObject = Nothing

            If Not TryGetStructuredDeliverableResult(responseText, rootObject, resultObject) Then
                Return False
            End If

            Dim producesUserDeliverable As Boolean =
                ExtractFirstBooleanValue(
                    rootObject,
                    "producesUserDeliverable",
                    "produces_user_deliverable").GetValueOrDefault(False)

            If Not producesUserDeliverable Then
                producesUserDeliverable =
                    ExtractFirstBooleanValue(
                        resultObject,
                        "producesUserDeliverable",
                        "produces_user_deliverable").GetValueOrDefault(False)
            End If

            Dim created As Boolean =
                ExtractFirstBooleanValue(
                    rootObject,
                    "created",
                    "saved",
                    "exported").GetValueOrDefault(False)

            If Not created Then
                created =
                    ExtractFirstBooleanValue(
                        resultObject,
                        "created",
                        "saved",
                        "exported").GetValueOrDefault(False)
            End If

            Dim references As List(Of String) = ExtractCreatedDeliverableReferences(responseText)

            Return (producesUserDeliverable AndAlso created) OrElse references.Count > 0
        End Function

        ''' <summary>
        ''' Classifies a transport-successful tool result as an operation no-op when it
        ''' reports that zero changes were applied. Mutation tools (Word write/markup/comment)
        ''' return status='none'/'no_match' and/or applied_count=0 while still returning valid
        ''' (non-error) JSON, i.e. transport success. Such a result must NOT count as workflow
        ''' progress. Returns False for results without these fields (non-mutation tools are
        ''' unaffected) and for any result that applied at least one change.
        ''' </summary>
        Public Shared Function IsZeroChangeOperationResult(responseText As String) As Boolean
            Dim raw As String = If(responseText, "").Trim()
            If raw = "" Then Return False

            Dim obj As JObject
            Try
                obj = TryCast(JToken.Parse(raw), JObject)
            Catch
                Return False
            End Try
            If obj Is Nothing Then Return False

            ' applied_count is the authoritative signal when present.
            Dim appliedToken As JToken = obj("applied_count")
            If appliedToken IsNot Nothing AndAlso appliedToken.Type <> JTokenType.Null Then
                Dim appliedCount As Integer
                If Integer.TryParse(appliedToken.ToString().Trim(), appliedCount) Then
                    Return appliedCount <= 0
                End If
            End If

            ' Fall back to an explicit no-op status only when applied_count is absent.
            Dim statusValue As String = If(obj.Value(Of String)("status"), "").Trim().ToLowerInvariant()
            Return statusValue = "none" OrElse statusValue = "no_match"
        End Function

        ''' <summary>
        ''' Builds a stable per-anchor breaker key from the tool name and the file it targets,
        ''' independent of the exact 'find' text. This makes reworded retries against the same
        ''' file collapse onto one counter so a repeated no-op edit can be bounded.
        ''' </summary>
        Public Shared Function BuildOperationTargetKey(toolName As String, responseText As String) As String
            Dim name As String = If(toolName, "").Trim().ToLowerInvariant()
            Dim path As String = ""
            Try
                Dim obj As JObject = TryCast(JToken.Parse(If(responseText, "").Trim()), JObject)
                If obj IsNot Nothing Then
                    path = If(obj.Value(Of String)("path"), "").Trim().ToLowerInvariant()
                End If
            Catch
            End Try
            Return name & "|" & path
        End Function

        ''' <summary>
        ''' Builds the same per-target breaker key as <see cref="BuildOperationTargetKey"/> but from a
        ''' known tool name and file path (e.g. the tool call arguments), so the key can be computed
        ''' BEFORE the tool executes. Used by the pre-execution no-op circuit breaker.
        ''' </summary>
        Public Shared Function BuildOperationTargetKeyFromPath(toolName As String, path As String) As String
            Return If(toolName, "").Trim().ToLowerInvariant() & "|" & If(path, "").Trim().ToLowerInvariant()
        End Function

        ''' <summary>
        ''' Builds a stable per-run key for a read/expansion request (e.g. context_expand) from the
        ''' stored reference plus the requested window, so a repeated expansion of an already-read
        ''' (ref + range) can be detected as no-progress and suppressed. Generalizes the no-op circuit
        ''' breaker beyond Word mutations to any repeated no-progress call.
        ''' </summary>
        Public Shared Function BuildExpandedRefRangeKey(refId As String, rangeStart As String, rangeEnd As String) As String
            Return If(refId, "").Trim().ToLowerInvariant() & "|" &
                   If(rangeStart, "").Trim() & "|" &
                   If(rangeEnd, "").Trim()
        End Function

        ''' <summary>
        ''' Token overload of the zero-change classifier for callers that already parsed the result.
        ''' Checks applied_count (authoritative) first on the root and result objects, then falls back
        ''' to an explicit no-op status. Returns False when neither signal is present.
        ''' </summary>
        Private Shared Function IsZeroChangeOperationToken(root As JObject, result As JObject) As Boolean
            For Each obj As JObject In New JObject() {root, result}
                If obj Is Nothing Then Continue For
                Dim ac As JToken = obj("applied_count")
                If ac IsNot Nothing AndAlso ac.Type <> JTokenType.Null Then
                    Dim n As Integer
                    If Integer.TryParse(ac.ToString().Trim(), n) Then
                        Return n <= 0
                    End If
                End If
            Next

            For Each obj As JObject In New JObject() {root, result}
                If obj Is Nothing Then Continue For
                Dim st As String = If(obj.Value(Of String)("status"), "").Trim().ToLowerInvariant()
                If st = "none" OrElse st = "no_match" Then Return True
            Next

            Return False
        End Function

        Public Shared Sub NoteToolResultForRepair(runState As ToolingRunState,
                                                  toolName As String,
                                                  responseText As String,
                                                  Optional resultKind As String = "")
            If runState Is Nothing Then Return

            ResetLastToolOutputMetadata(runState)

            Dim raw As String = If(responseText, "").Trim()
            If raw = "" Then Return

            Dim normalizedKind As String = If(resultKind, "").Trim()
            If String.Equals(normalizedKind, "error", StringComparison.OrdinalIgnoreCase) Then
                Return
            End If

            Try
                Dim token As JToken = JToken.Parse(raw)

                If Not TypeOf token Is JObject AndAlso Not TypeOf token Is JArray Then
                    Return
                End If

                runState.LastStructuredToolResult = raw
                runState.LastStructuredToolName = If(toolName, "")

                If normalizedKind = "" Then
                    normalizedKind = If(TypeOf token Is JObject, "json_object", "json_array")
                End If

                runState.LastStructuredToolResultKind = normalizedKind
                NoteStructuredToolOutputMetadata(runState, token, normalizedKind)

                If TypeOf token Is JObject Then
                    TryNoteStructuredOutputReference(runState, DirectCast(token, JObject))
                End If
            Catch
                If normalizedKind <> "" AndAlso
                   Not String.Equals(normalizedKind, "text", StringComparison.OrdinalIgnoreCase) Then

                    runState.LastStructuredToolResult = raw
                    runState.LastStructuredToolName = If(toolName, "")
                    runState.LastStructuredToolResultKind = normalizedKind

                    If String.Equals(normalizedKind, "json_object", StringComparison.OrdinalIgnoreCase) OrElse
                       String.Equals(normalizedKind, "json_array", StringComparison.OrdinalIgnoreCase) Then
                        runState.LastToolProducesIntermediateData = True
                    End If
                End If
            End Try
        End Sub

        Private Shared Sub TryNoteStructuredOutputReference(runState As ToolingRunState,
                                                            payload As JObject)
            If runState Is Nothing OrElse payload Is Nothing Then Return

            Dim reference As String =
                ExtractFirstStringValue(
                    payload,
                    "output_reference",
                    "state_reference",
                    "reference",
                    "output_path",
                    "state_path",
                    "path",
                    "file_path",
                    "workspace_path",
                    "outputFilePath",
                    "output_file_path",
                    "outputFileName",
                    "output_file_name",
                    "output_filename",
                    "file_name",
                    "filename",
                    "memory_key",
                    "stub")

            If String.IsNullOrWhiteSpace(reference) Then
                reference =
                    ExtractFirstStringValue(
                        TryCast(payload("result"), JObject),
                        "output_reference",
                        "state_reference",
                        "reference",
                        "output_path",
                        "state_path",
                        "path",
                        "file_path",
                        "workspace_path",
                        "outputFilePath",
                        "output_file_path",
                        "outputFileName",
                        "output_file_name",
                        "output_filename",
                        "file_name",
                        "filename",
                        "memory_key",
                        "stub")
            End If

            If Not String.IsNullOrWhiteSpace(reference) Then
                runState.LastKnownOutputReference = reference
            End If
        End Sub

        Private Shared Function ExtractFirstStringValue(payload As JObject,
                                                        ParamArray keys() As String) As String
            If payload Is Nothing OrElse keys Is Nothing Then Return ""

            For Each key In keys
                If String.IsNullOrWhiteSpace(key) Then Continue For

                Dim token As JToken = payload(key)
                If token Is Nothing Then Continue For

                Dim value As String = token.ToString().Trim()
                If value <> "" Then
                    Return value
                End If
            Next

            Return ""
        End Function


        Public Shared Sub NoteToolExecutionMetadata(runState As ToolingRunState,
                                                    toolName As String,
                                                    arguments As IDictionary(Of String, Object),
                                                    success As Boolean)
            If runState Is Nothing Then Return

            runState.ActiveToolingSession = True
            runState.HasOpenToolWorkflow = True

            Dim classification As ToolCallClassification = ClassifyToolName(toolName)

            If success Then
                runState.LastSuccessfulToolCall = If(toolName, "")
                runState.RegisterSuccessfulTool(toolName)
            End If

            Select Case classification
                Case ToolCallClassification.Mutating
                    runState.LastMutationToolCall = If(toolName, "")
                Case ToolCallClassification.Agent
                    runState.LastAgentToolCall = If(toolName, "")
                Case ToolCallClassification.ReadOnlyIndependent, ToolCallClassification.Stateful
                    runState.LastReadOnlyStateToolCall = If(toolName, "")
            End Select

            Dim knownPath As String = ExtractFirstPathArgument(arguments)
            If Not String.IsNullOrWhiteSpace(knownPath) Then
                runState.LastKnownOutputReference = knownPath
                runState.LastStateFilePath = knownPath

                If classification = ToolCallClassification.Mutating Then
                    runState.LastOutputPath = knownPath
                End If
            End If

            Dim collectionSize As Integer? = InferCollectionSize(arguments)
            If collectionSize.HasValue Then
                runState.LastCollectionSize = collectionSize
            End If

            If success AndAlso runState.LastCollectionSize.HasValue AndAlso runState.LastCollectionSize.Value > 1 Then
                runState.LastProcessedItemCount = If(runState.LastProcessedItemCount, 0) + 1
            End If
        End Sub

        ''' <summary>
        ''' Returns corrective guidance when a tool flagged as single-invocation-preferring has already run
        ''' successfully in the session, so the model consolidates remaining work instead of issuing repeated,
        ''' expensive re-invocations. Returns an empty string when no such repetition risk exists.
        ''' </summary>
        Public Shared Function BuildConsolidatableToolGuidance(runState As ToolingRunState) As String
            If runState Is Nothing Then Return ""
            If String.IsNullOrWhiteSpace(runState.LastConsolidatableToolName) Then Return ""
            Return ConsolidatableToolConsolidationInstruction
        End Function

        Public Shared Function BuildTaskStatusFooter(status As String, reason As String) As String
            Dim normalizedStatus As String = If(status, "").Trim().ToLowerInvariant()
            If normalizedStatus = "" Then
                normalizedStatus = "blocked"
            End If

            Dim footerObject As New JObject(
        New JProperty("status", normalizedStatus),
        New JProperty("reason", NormalizeFooterReason(reason, normalizedStatus)))

            Return "<TASK_STATUS>" & footerObject.ToString(Formatting.None) & "</TASK_STATUS>"
        End Function

        Public Enum HostFailureCategory
            Unknown = 0
            TransportRateLimited = 1
            TransportUnavailable = 2
            TransportTimeout = 3
            RequiredCapabilityUnavailable = 4
            DeliverableValidationFailed = 5
            InternalExecutionFailure = 6
        End Enum

        Public Shared Function ClassifyHostFailure(errorCode As System.String,
                                                   message As System.String) As HostFailureCategory
            Dim code As System.String = If(errorCode, System.String.Empty).Trim().ToLowerInvariant()
            Dim detail As System.String = If(message, System.String.Empty).Trim().ToLowerInvariant()

            If code = "llm_transport_timeout" OrElse code.Contains("timeout") Then
                Return HostFailureCategory.TransportTimeout
            End If

            If code = "llm_transport_retry_exhausted" OrElse code.Contains("transport") Then
                If detail.Contains("429") OrElse detail.Contains("rate limit") OrElse detail.Contains("rate-limit") OrElse detail.Contains("quota") Then
                    Return HostFailureCategory.TransportRateLimited
                End If
                Return HostFailureCategory.TransportUnavailable
            End If

            If code.Contains("not_available") OrElse code.Contains("unavailable_tool") OrElse
               code.Contains("required_capability") OrElse code.Contains("missing_required_tool") Then
                Return HostFailureCategory.RequiredCapabilityUnavailable
            End If

            If code.Contains("deliverable") OrElse code.Contains("artifact") OrElse
               code.Contains("validation") Then
                Return HostFailureCategory.DeliverableValidationFailed
            End If

            If code = "invalid_text_only_finalization" OrElse
               code = "complete_missing_required_successful_tools" OrElse
               code = "host_generated_blocked" Then
                Return HostFailureCategory.InternalExecutionFailure
            End If

            Return HostFailureCategory.Unknown
        End Function

        Public Shared Function TryBuildDeterministicHostFailureMessage(errorCode As System.String,
                                                                       message As System.String,
                                                                       userLanguage As System.String,
                                                                       ByRef result As System.String) As System.Boolean
            result = System.String.Empty
            Dim category As HostFailureCategory = ClassifyHostFailure(errorCode, message)
            If category = HostFailureCategory.Unknown Then Return False

            Dim lang As System.String = If(userLanguage, System.String.Empty).Trim().ToLowerInvariant()
            Dim sep As System.Int32 = lang.IndexOfAny(New System.Char() {"-"c, "_"c})
            If sep > 0 Then lang = lang.Substring(0, sep)

            Select Case lang
                Case "de"
                    Select Case category
                        Case HostFailureCategory.TransportRateLimited
                            result = "Der KI-Dienst ist vorübergehend ausgelastet oder rate-limitiert. Der Vorgang konnte deshalb nicht zuverlässig abgeschlossen werden. Bitte versuchen Sie es später erneut."
                        Case HostFailureCategory.TransportUnavailable
                            result = "Der KI-Dienst ist vorübergehend nicht erreichbar. Der Vorgang konnte deshalb nicht zuverlässig abgeschlossen werden. Bitte versuchen Sie es später erneut."
                        Case HostFailureCategory.TransportTimeout
                            result = "Die Anfrage an den KI-Dienst hat das zulässige Zeitlimit überschritten. Bitte versuchen Sie es erneut."
                        Case HostFailureCategory.RequiredCapabilityUnavailable
                            result = "Eine für diesen Vorgang erforderliche Funktion ist derzeit nicht verfügbar. Der Vorgang konnte deshalb nicht zuverlässig abgeschlossen werden."
                        Case HostFailureCategory.DeliverableValidationFailed
                            result = "Das angeforderte Ergebnis konnte nicht vollständig und zuverlässig validiert werden. Es wurde deshalb nicht als abgeschlossen ausgeliefert."
                        Case Else
                            result = "Der Vorgang konnte aufgrund eines internen Ausführungsfehlers nicht zuverlässig abgeschlossen werden. Bitte versuchen Sie es erneut."
                    End Select
                Case "fr"
                    Select Case category
                        Case HostFailureCategory.TransportRateLimited
                            result = "Le service d’IA est temporairement saturé ou limité en débit. La tâche n’a donc pas pu être terminée de manière fiable. Veuillez réessayer plus tard."
                        Case HostFailureCategory.TransportUnavailable
                            result = "Le service d’IA est temporairement indisponible. La tâche n’a donc pas pu être terminée de manière fiable. Veuillez réessayer plus tard."
                        Case HostFailureCategory.TransportTimeout
                            result = "La requête adressée au service d’IA a dépassé le délai autorisé. Veuillez réessayer."
                        Case HostFailureCategory.RequiredCapabilityUnavailable
                            result = "Une fonction requise pour cette tâche n’est actuellement pas disponible. La tâche n’a donc pas pu être terminée de manière fiable."
                        Case HostFailureCategory.DeliverableValidationFailed
                            result = "Le résultat demandé n’a pas pu être validé complètement et de manière fiable. Il n’a donc pas été remis comme résultat final."
                        Case Else
                            result = "La tâche n’a pas pu être terminée de manière fiable en raison d’une erreur d’exécution interne. Veuillez réessayer."
                    End Select
                Case "it"
                    Select Case category
                        Case HostFailureCategory.TransportRateLimited
                            result = "Il servizio di IA è temporaneamente sovraccarico o soggetto a limiti di frequenza. L’operazione non ha quindi potuto essere completata in modo affidabile. Riprova più tardi."
                        Case HostFailureCategory.TransportUnavailable
                            result = "Il servizio di IA è temporaneamente non disponibile. L’operazione non ha quindi potuto essere completata in modo affidabile. Riprova più tardi."
                        Case HostFailureCategory.TransportTimeout
                            result = "La richiesta al servizio di IA ha superato il tempo massimo consentito. Riprova."
                        Case HostFailureCategory.RequiredCapabilityUnavailable
                            result = "Una funzione necessaria per questa operazione non è attualmente disponibile. L’operazione non ha quindi potuto essere completata in modo affidabile."
                        Case HostFailureCategory.DeliverableValidationFailed
                            result = "Il risultato richiesto non ha potuto essere convalidato in modo completo e affidabile e pertanto non è stato consegnato come risultato finale."
                        Case Else
                            result = "L’operazione non ha potuto essere completata in modo affidabile a causa di un errore interno di esecuzione. Riprova."
                    End Select
                Case Else
                    Select Case category
                        Case HostFailureCategory.TransportRateLimited
                            result = "The AI service is temporarily busy or rate-limited, so the task could not be completed reliably. Please try again later."
                        Case HostFailureCategory.TransportUnavailable
                            result = "The AI service is temporarily unavailable, so the task could not be completed reliably. Please try again later."
                        Case HostFailureCategory.TransportTimeout
                            result = "The request to the AI service exceeded the allowed time limit. Please try again."
                        Case HostFailureCategory.RequiredCapabilityUnavailable
                            result = "A capability required for this task is currently unavailable, so the task could not be completed reliably."
                        Case HostFailureCategory.DeliverableValidationFailed
                            result = "The requested result could not be fully and reliably validated, so it was not delivered as complete."
                        Case Else
                            result = "The task could not be completed reliably because of an internal execution error. Please try again."
                    End Select
            End Select

            Return result <> System.String.Empty
        End Function

        Public Shared Function BuildUserSafeBlockedFinalMessage(runState As ToolingRunState,
                                                                errorCode As String,
                                                                message As String,
                                                                successCount As Integer,
                                                                failedCount As Integer,
                                                                Optional userLanguage As String = "",
                                                                Optional appendTaskStatusFooter As Boolean = True) As String
            Dim deterministicMessage As System.String = System.String.Empty
            If TryBuildDeterministicHostFailureMessage(errorCode, message, userLanguage, deterministicMessage) Then
                If appendTaskStatusFooter Then
                    deterministicMessage &= " " & BuildTaskStatusFooter("blocked", If(errorCode, "host_generated_blocked"))
                End If
                Return deterministicMessage.Trim()
            End If

            Dim useMemoryMessage As Boolean =
                String.Equals(errorCode, MissingRequiredMemoryAccessCode, StringComparison.OrdinalIgnoreCase) OrElse
                String.Equals(errorCode, MemoryListDoneButMemoryGetRequiredCode, StringComparison.OrdinalIgnoreCase) OrElse
                String.Equals(errorCode, MemoryGetFailedCode, StringComparison.OrdinalIgnoreCase) OrElse
                String.Equals(errorCode, NoRelevantMemoryAvailableCode, StringComparison.OrdinalIgnoreCase) OrElse
                String.Equals(errorCode, PartialMemoryRetrievalRequiresSubsetDisclosureCode, StringComparison.OrdinalIgnoreCase) OrElse
                (runState IsNot Nothing AndAlso
                 runState.MemoryGroundingMode = MemoryGroundingMode.Required AndAlso
                 (runState.MemoryListCalledThisTurn OrElse
                  runState.MemoryGetCalledThisTurn OrElse
                  runState.FullMemoryValueAvailableThisTurn OrElse
                  runState.MemoryGetCountThisTurn > 0))

            Dim finalText As String

            If String.Equals(errorCode, RequestedDeliverableNotCreatedCode, StringComparison.OrdinalIgnoreCase) Then
                If HasProducedUserDeliverable(runState) Then
                    finalText = "Something went wrong after the requested deliverable was created. Please review the created result and try again if needed."
                Else
                    finalText = "Something went wrong. I could not reliably create the requested deliverable. Please try again or narrow the request."
                End If
            Else
                finalText =
                    If(
                        useMemoryMessage,
                        "Something went wrong. I could not fully load or evaluate the stored content. Please try again or narrow the request.",
                        "Something went wrong. I could not finish the task reliably. Please try again or narrow the request.")
            End If

            If appendTaskStatusFooter Then
                finalText &= " " & BuildTaskStatusFooter("blocked", If(errorCode, "host_generated_blocked"))
            End If

            Return finalText.Trim()
        End Function

        Private Shared Function IsRawInternalJsonResponse(text As String) As Boolean
            Dim trimmed As String = If(text, "").Trim()
            If trimmed = "" Then Return False

            Try
                Dim token As JToken = JToken.Parse(trimmed)
                Dim obj As JObject = TryCast(token, JObject)
                If obj Is Nothing Then Return False

                Return obj("status") IsNot Nothing OrElse
                       obj("error") IsNot Nothing OrElse
                       obj("resultKind") IsNot Nothing
            Catch
                Return False
            End Try
        End Function

        Private Shared Function NormalizeFooterReason(reason As String, Optional status As String = "") As String
            Dim normalized As String =
        Regex.Replace(
            If(reason, ""),
            "\s+",
            " ",
            RegexOptions.CultureInvariant).Trim()

            If normalized = "" Then
                Select Case If(status, "").Trim().ToLowerInvariant()
                    Case "complete"
                        normalized = "answer ready"
                    Case "blocked"
                        normalized = "no safe completion path"
                    Case Else
                        normalized = "status recorded"
                End Select
            End If

            If normalized.Length > TaskStatusReasonMaxChars Then
                normalized = normalized.Substring(0, TaskStatusReasonMaxChars).Trim()
            End If

            Return normalized
        End Function


        Private Shared Function ExtractFirstPathArgument(arguments As IDictionary(Of String, Object)) As String
            If arguments Is Nothing Then Return ""

            Dim keys As String() = {
                "path",
                "file_path",
                "source_path",
                "target_path",
                "output_path",
                "state_path",
                "workspace_path"
            }

            For Each key In keys
                If Not arguments.ContainsKey(key) OrElse arguments(key) Is Nothing Then Continue For

                Dim value As String = TryGetScalarString(arguments(key))
                If Not String.IsNullOrWhiteSpace(value) Then
                    Return value.Trim()
                End If
            Next

            Return ""
        End Function

        Private Shared Function InferCollectionSize(arguments As IDictionary(Of String, Object)) As Integer?
            If arguments Is Nothing Then Return Nothing

            For Each pair In arguments
                If pair.Value Is Nothing OrElse TypeOf pair.Value Is String Then Continue For

                If TypeOf pair.Value Is JArray Then
                    Return DirectCast(pair.Value, JArray).Count
                End If

                If TypeOf pair.Value Is IEnumerable Then
                    Dim count As Integer = 0
                    For Each item In DirectCast(pair.Value, IEnumerable)
                        count += 1
                    Next
                    Return count
                End If
            Next

            Return Nothing
        End Function

        Private Shared Function TryGetScalarString(value As Object) As String
            If value Is Nothing Then Return ""

            If TypeOf value Is JValue Then
                Return DirectCast(value, JValue).ToString()
            End If

            If TypeOf value Is String Then
                Return CStr(value)
            End If

            Return value.ToString()
        End Function

        Private Shared Function HasAnyPhrase(name As String, ParamArray phrases() As String) As Boolean
            If String.IsNullOrWhiteSpace(name) Then Return False
            If phrases Is Nothing Then Return False

            For Each phrase In phrases
                If String.IsNullOrWhiteSpace(phrase) Then Continue For

                If name.IndexOf(phrase, StringComparison.OrdinalIgnoreCase) >= 0 Then
                    Return True
                End If
            Next

            Return False
        End Function

        Private Shared Function HasAnyToken(name As String, ParamArray expectedTokens() As String) As Boolean
            If String.IsNullOrWhiteSpace(name) Then Return False
            If expectedTokens Is Nothing Then Return False

            Dim tokens = name.Split(New Char() {"_"c, "-"c, "."c}, StringSplitOptions.RemoveEmptyEntries)

            For Each token In tokens
                For Each expected In expectedTokens
                    If String.IsNullOrWhiteSpace(expected) Then Continue For

                    If token.Equals(expected, StringComparison.OrdinalIgnoreCase) Then
                        Return True
                    End If
                Next
            Next

            Return False
        End Function

        Public Shared Function HasBlockingUnresolvedToolFailure(runState As ToolingRunState) As Boolean
            If runState Is Nothing OrElse Not runState.HasUnresolvedToolFailure Then
                Return False
            End If

            If runState.UnresolvedToolFailures IsNot Nothing AndAlso runState.UnresolvedToolFailures.Count > 0 Then
                For Each failure As ToolFailureRecord In runState.UnresolvedToolFailures
                    If failure Is Nothing Then Continue For
                    If String.Equals(
                        If(failure.ErrorCode, "").Trim(),
                        ToolNotExposedInCurrentTurnCode,
                        StringComparison.OrdinalIgnoreCase) Then
                        Continue For
                    End If

                    If failure.RecoveryEvidenceObserved Then
                        Continue For
                    End If

                    ' Defensive finalization rule: an earlier artifact-scoped producer/inspection
                    ' failure is no longer blocking when that exact expected logical slot is now
                    ' fully qualified by one existing Final with all host-verified required effects.
                    ' This is deliberately narrow: unscoped failures, operation/subagent scopes,
                    ' different slots and incomplete deliverable contracts remain blocking.
                    If IsArtifactScopedFailureSupersededByValidatedDeliverable(runState, failure) Then
                        Continue For
                    End If

                    Return True
                Next
                Return False
            End If

            ' Compatibility fallback for states created before the unresolved-failure registry
            ' was introduced or for external callers that only populated the legacy projection.
            Return Not String.Equals(
                If(runState.LastErrorCode, "").Trim(),
                ToolNotExposedInCurrentTurnCode,
                StringComparison.OrdinalIgnoreCase)
        End Function

        Public Shared Function HasDeclaredTerminalOutcomeFailure(runState As ToolingRunState) As System.Boolean
            If runState Is Nothing OrElse Not runState.HasUnresolvedToolFailure Then Return False

            If runState.UnresolvedToolFailures IsNot Nothing AndAlso runState.UnresolvedToolFailures.Count > 0 Then
                For Each failure As ToolFailureRecord In runState.UnresolvedToolFailures
                    If failure Is Nothing OrElse Not failure.Terminal Then Continue For
                    If failure.RecoveryEvidenceObserved Then Continue For
                    If SubAgentRuntimeHardening.IsDeclaredTerminalOutcomeErrorCode(failure.ErrorCode) Then
                        Return True
                    End If
                Next
                Return False
            End If

            Return runState.LastFailureTerminal AndAlso
                   SubAgentRuntimeHardening.IsDeclaredTerminalOutcomeErrorCode(runState.LastErrorCode)
        End Function

        Private Shared Function IsArtifactScopedFailureSupersededByValidatedDeliverable(
            runState As ToolingRunState,
            failure As ToolFailureRecord) As System.Boolean

            If runState Is Nothing OrElse failure Is Nothing Then Return False
            If Not runState.HasExpectedDeliverableContract Then Return False

            Dim scope As System.String = If(failure.RecoveryScopeKey, System.String.Empty).Trim()
            Const artifactPrefix As System.String = "artifact:"
            If Not scope.StartsWith(artifactPrefix, System.StringComparison.Ordinal) Then Return False

            Dim identity As System.String = scope.Substring(artifactPrefix.Length)
            Dim separatorIndex As System.Int32 = identity.IndexOf("|"c)
            If separatorIndex <= 0 OrElse separatorIndex >= identity.Length - 1 Then Return False
            If identity.IndexOf("|"c, separatorIndex + 1) >= 0 Then Return False

            Dim logicalId As System.String = identity.Substring(0, separatorIndex).Trim()
            Dim slotId As System.String = identity.Substring(separatorIndex + 1).Trim()
            If logicalId = System.String.Empty OrElse slotId = System.String.Empty Then Return False
            If Not runState.IsExpectedDeliverableSlot(logicalId, slotId) Then Return False

            Return runState.IsExpectedDeliverableSlotSatisfied(logicalId, slotId)
        End Function

        Public Shared Sub ClearNonBlockingUnresolvedToolFailure(runState As ToolingRunState,
                                                         recoveryLabel As String)
            If runState Is Nothing OrElse Not runState.HasUnresolvedToolFailure Then
                Return
            End If

            Dim clearedRegistryFailure As Boolean = False
            While runState.ClearLatestFailureByCode(ToolNotExposedInCurrentTurnCode, recoveryLabel)
                clearedRegistryFailure = True
            End While

            If runState.UnresolvedToolFailures IsNot Nothing AndAlso runState.UnresolvedToolFailures.Count > 0 Then
                For i As System.Int32 = runState.UnresolvedToolFailures.Count - 1 To 0 Step -1
                    Dim failure As ToolFailureRecord = runState.UnresolvedToolFailures(i)
                    If Not IsArtifactScopedFailureSupersededByValidatedDeliverable(runState, failure) Then Continue For
                    runState.UnresolvedToolFailures.RemoveAt(i)
                    clearedRegistryFailure = True
                Next

                If clearedRegistryFailure Then
                    If runState.UnresolvedToolFailures.Count > 0 Then
                        runState.ProjectLatestUnresolvedFailure()
                    Else
                        runState.HasUnresolvedToolFailure = False
                        runState.LastToolName = System.String.Empty
                        runState.LastErrorCode = System.String.Empty
                        runState.LastErrorMessage = System.String.Empty
                        runState.LastFailureSkippedByPolicy = False
                        runState.LastFailureReturnedToParent = False
                        runState.LastFailureTerminal = False
                        runState.LastFailureRecoveryPolicy = ToolFailureRecoveryPolicy.SameToolSuccessOnly
                        runState.LastFailureToolClassification = ToolCallClassification.Unknown
                        runState.LastFailureToolErrorHandling = System.String.Empty
                        runState.LastFailureRecoveredByToolCall = True
                        runState.LastFailureHandledByBlockedFinal = False
                        runState.LastFailureUltimatelyFatal = False
                        runState.RecoveryToolName = If(recoveryLabel, System.String.Empty)
                    End If
                End If
            End If

            If clearedRegistryFailure Then
                Return
            End If

            ' Compatibility fallback for legacy states without a registry entry.
            If Not String.Equals(
                If(runState.LastErrorCode, "").Trim(),
                ToolNotExposedInCurrentTurnCode,
                StringComparison.OrdinalIgnoreCase) Then
                Return
            End If

            runState.HasUnresolvedToolFailure = False
            runState.LastFailureRecoveredByToolCall = True
            runState.LastFailureHandledByBlockedFinal = False
            runState.LastFailureUltimatelyFatal = False
            runState.RecoveryToolName = If(recoveryLabel, "")
        End Sub

    End Class

End Namespace
