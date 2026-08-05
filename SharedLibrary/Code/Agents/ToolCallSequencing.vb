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


        Public Const TaskStatusReasonMaxChars As Integer = 160

        Private Sub New()
        End Sub

        Public Const DependentBatchingInstruction As String =
            "When using tools, the host executes multiple tool calls from one response strictly in the order emitted." & vbCrLf &
            "Only emit multiple tool calls in one response when every later call's arguments are already fully known at emission time." & vbCrLf &
            "If a later step depends on inspecting the result of an earlier tool call, emit only the earlier call and wait for its result before deciding the next call." & vbCrLf &
            "Do not rely on the host to rewrite, infer, defer, queue, or replay omitted tool calls."

        Public Const ConsolidatableToolConsolidationInstruction As String =
            "A tool designed to complete an entire task in a single call has already run successfully in this session." & vbCrLf &
            "Do not issue additional calls to that tool for work that could have been included in the earlier call." & vbCrLf &
            "Only call it again if you genuinely must inspect the earlier result before deciding the next step; otherwise consolidate all remaining deterministic processing into one call."

        Public Const UnresolvedToolFailureCode As String = "unresolved_tool_failure"
        Public Const InvalidTextOnlyFinalizationCode As String = "invalid_text_only_finalization"
        Public Const MissingRequiredMemoryAccessCode As String = "missing_required_memory_access"
        Public Const MemoryListDoneButMemoryGetRequiredCode As String = "memory_list_done_but_memory_get_required"
        Public Const MemoryGetFailedCode As String = "memory_get_failed"
        Public Const NoRelevantMemoryAvailableCode As String = "no_relevant_memory_available"
        Public Const PartialMemoryRetrievalRequiresSubsetDisclosureCode As String = "partial_memory_retrieval_requires_subset_disclosure"
        Public Const RequestedDeliverableNotCreatedCode As String = "requested_deliverable_not_created"

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

            Public Property ActiveToolingSession As Boolean
            Public Property HasOpenToolWorkflow As Boolean
            Public Property LastStateFilePath As String
            Public Property LastOutputPath As String
            Public Property LastCollectionSize As Integer?
            Public Property LastProcessedItemCount As Integer?
            Public Property LastSuccessfulToolCall As String
            Public Property LastMutationToolCall As String
            Public Property LastAgentToolCall As String
            Public Property LastReadOnlyStateToolCall As String
            Public Property LastDetectedTurnType As String
            Public Property LastInvalidTurnReason As String
            Public Property FinalResponseOrigin As String
            Public Property ToolRequiredModeUsed As Boolean

            Public Property UserLanguage As String
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

            Public Sub NoteToolFailure(toolName As String,
                               Optional errorCode As String = "",
                               Optional errorMessage As String = "",
                               Optional skippedByPolicy As Boolean = False,
                               Optional returnedToParent As Boolean = False)
                HasUnresolvedToolFailure = True
                LastToolName = If(toolName, "")
                LastErrorCode = If(errorCode, "")
                LastErrorMessage = If(errorMessage, "")
                LastFailureSkippedByPolicy = skippedByPolicy
                LastFailureReturnedToParent = returnedToParent
                LastFailureRecoveredByToolCall = False
                LastFailureHandledByBlockedFinal = False
                LastFailureUltimatelyFatal = False
                RecoveryToolName = ""
            End Sub

            Public Sub NoteRecoveryByLaterToolCall(toolName As String)
                If Not HasUnresolvedToolFailure Then Return

                HasUnresolvedToolFailure = False
                LastFailureRecoveredByToolCall = True
                LastFailureHandledByBlockedFinal = False
                LastFailureUltimatelyFatal = False
                RecoveryToolName = If(toolName, "")
            End Sub

            Public Sub NoteBlockedFinalHandled()
                If Not HasUnresolvedToolFailure Then Return

                HasUnresolvedToolFailure = False
                LastFailureRecoveredByToolCall = False
                LastFailureHandledByBlockedFinal = True
                LastFailureUltimatelyFatal = False
                RecoveryToolName = ""
            End Sub

            Public Sub NoteFailureFatal()
                If Not HasUnresolvedToolFailure Then Return
                LastFailureUltimatelyFatal = True
            End Sub

            Public Sub NoteSuccessfulProgress()
                HasUnresolvedToolFailure = False
                LastToolName = ""
                LastErrorCode = ""
                LastErrorMessage = ""
                LastFailureSkippedByPolicy = False
                LastFailureReturnedToParent = False
                LastFailureRecoveredByToolCall = False
                LastFailureHandledByBlockedFinal = False
                LastFailureUltimatelyFatal = False
                RecoveryToolName = ""
            End Sub
        End Class

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
                                                         Optional lastToolErrorMessage As String = "") As String
            Dim obj As New JObject(
                New JProperty("status", "blocked"),
                New JProperty("error", New JObject(
                    New JProperty("code", If(errorCode, "")),
                    New JProperty("phase", If(phase, "")),
                    New JProperty("message", If(message, "")))))

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
            End If

            Return obj.ToString(Formatting.None)
        End Function


        Public Shared Function StripTaskStatusBlocksFromUserFacingText(text As String) As String
            Dim raw As String = If(text, "")
            If raw = "" Then
                Return ""
            End If

            Dim stripped As String =
                Regex.Replace(
                    raw,
                    "\s*<TASK_STATUS>\s*\{.*?\}\s*</TASK_STATUS>\s*",
                    "",
                    RegexOptions.IgnoreCase Or RegexOptions.Singleline Or RegexOptions.CultureInvariant)

            Return stripped.Trim()
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

            Dim matches As MatchCollection =
        Regex.Matches(
            trimmedEnd,
            "<TASK_STATUS>\s*(\{.*?\})\s*</TASK_STATUS>",
            RegexOptions.IgnoreCase Or RegexOptions.Singleline Or RegexOptions.CultureInvariant)

            result.FooterCount = matches.Count

            If matches.Count = 0 Then
                result.FailureReason = "missing_task_status"
                Return result
            End If

            result.IsPresent = True

            If matches.Count <> 1 Then
                result.FailureReason = "multiple_task_status"
                Return result
            End If

            Dim match As Match = matches(0)

            If match.Index + match.Length <> trimmedEnd.Length Then
                result.FailureReason = "task_status_not_at_end"
                Return result
            End If

            Dim jsonText As String = match.Groups(1).Value.Trim()
            result.FooterJson = jsonText
            result.TextBeforeFooter = trimmedEnd.Substring(0, match.Index).TrimEnd()

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

        Public Shared Function BuildActiveToolingRepairPrompt(Optional runState As ToolingRunState = Nothing,
                                                      Optional invalidReason As String = "") As String
            Dim normalizedInvalidReason As String = If(invalidReason, "").Trim().ToLowerInvariant()
            Dim prompt As String

            Select Case normalizedInvalidReason
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


        Public Shared Function HasProducedUserDeliverable(runState As ToolingRunState) As Boolean
            If runState Is Nothing Then
                Return False
            End If

            Return runState.AnyUserDeliverableProducedThisRun OrElse
                   runState.LastToolProducesUserDeliverable OrElse
                   Not String.IsNullOrWhiteSpace(runState.LastToolOutputArtifactRef) OrElse
                   Not String.IsNullOrWhiteSpace(runState.LastToolOutputFilePath)
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

            If HasProducedUserDeliverable(runState) Then
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

            Dim outputFilePath As String =
                ExtractFirstStringValue(
                    rootObject,
                    "outputFilePath",
                    "output_file_path",
                    "output_path",
                    "file_path",
                    "path")

            If String.IsNullOrWhiteSpace(outputFilePath) Then
                outputFilePath =
                    ExtractFirstStringValue(
                        resultObject,
                        "outputFilePath",
                        "output_file_path",
                        "output_path",
                        "file_path",
                        "path")
            End If

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

            Dim inferredDeliverable As Boolean =
                explicitDeliverable.GetValueOrDefault(False) OrElse
                createdStatus.GetValueOrDefault(False) OrElse
                Not String.IsNullOrWhiteSpace(artifactRef) OrElse
                Not String.IsNullOrWhiteSpace(outputFilePath)

            Dim inferredIntermediate As Boolean =
                explicitIntermediate.GetValueOrDefault(False) OrElse
                ((TypeOf payload Is JObject OrElse TypeOf payload Is JArray) AndAlso Not inferredDeliverable)

            runState.LastToolProducesUserDeliverable = inferredDeliverable
            runState.LastToolProducesIntermediateData = inferredIntermediate
            runState.LastToolOutputArtifactRef = If(artifactRef, "")
            runState.LastToolOutputFilePath = If(outputFilePath, "")
            runState.LastToolOutputMimeType = If(outputMimeType, "")
            runState.LastToolOutputKind = If(outputKind, "")

            If inferredDeliverable Then
                runState.AnyUserDeliverableProducedThisRun = True
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

        Public Shared Function BuildUserSafeBlockedFinalMessage(runState As ToolingRunState,
                                                                errorCode As String,
                                                                message As String,
                                                                successCount As Integer,
                                                                failedCount As Integer,
                                                                Optional userLanguage As String = "",
                                                                Optional appendTaskStatusFooter As Boolean = True) As String
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

            Return Not String.Equals(
                If(runState.LastErrorCode, "").Trim(),
                ToolNotExposedInCurrentTurnCode,
                StringComparison.OrdinalIgnoreCase)
        End Function

        Public Shared Sub ClearNonBlockingUnresolvedToolFailure(runState As ToolingRunState,
                                                         recoveryLabel As String)
            If runState Is Nothing OrElse Not runState.HasUnresolvedToolFailure Then
                Return
            End If

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
