' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.Tooling.ToolExecutionContext.vb
' Purpose: Encapsulates per-run tooling state, execution tracking, logging, and workflow metadata.
'          Serves as central context object passed through entire tooling loop lifecycle.
'
' Data Structure:
'  - Tool & Registry State:
'      - SelectedTools: ModelConfig instances for this session.
'      - AllowedToolRegistry/AuthoritativeToolRegistry: Tool availability constraints.
'      - AuthoritativeToolRegistrySnapshot: Parent registry snapshot for sub-agents.
'      - LazyToolLoadingEnabled: Lightweight index initially exposed to model.
'  - Execution Tracking:
'      - AllToolResponses: All ToolResponse objects (success and failure) for session.
'      - CurrentIteration/MaxIterations: Loop counters and abort thresholds.
'      - IsCancelled: UI-driven cancellation flag.
'  - Failure Management:
'      - FailedToolCallCounts: Per-tool failure counter (OrdinalIgnoreCase).
'      - DuplicateFailureAbortThreshold: Abort when same tool fails N times.
'      - ConsecutiveFailedToolName/Count/AbortThreshold: Abort on N consecutive failures.
'      - LastToolExecutionSignature/RepeatCount/DuplicateToolExecutionAbortThreshold: Detect duplicate-tool loops.
'  - Logging & UI:
'      - LogEntries: In-memory timestamped log lines.
'      - LogWindowForm: Optional UI window for user-visible progress.
'      - ExternalLogSink: Optional callback for external log forwarding.
'      - Log(message, level): Central log method routing to file, UI, and external sinks.
'  - Workflow & Sequencing:
'      - WorkflowId/RuntimeState: Workflow metadata for continuity logging.
'      - LatestUserRequestRaw/HostTaskSummary: User request and task context.
'      - SequencingState: Tool call sequencing decisions (memory grounding, tool priority).
'      - RunId: Unique run identifier (GUID without hyphens).
'  - Sub-Agent Tracking:
'      - SubAgentInvocationCount: Total sub-agent invocation counter.
'      - SubAgentInvocationCountsByAgent: Per-agent invocation counter dict.
'  - State Finalization:
'      - FinalizationBlocked: Prevents early completion.
'      - FinalizationBlockedReason: Human-readable explanation for block.
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

    ''' <summary>
    ''' Holds per-run state for the tooling loop, including selected tools, iteration counters, and logging.
    ''' </summary>
    Public Class ToolExecutionContext

        ''' <summary>Tools selected for this session.</summary>
        Public Property SelectedTools As List(Of ModelConfig)

        ''' <summary>Allow-listed tools selected by the user, available for on-demand loading.</summary>
        Public Property AllowedToolRegistry As SharedLibrary.Agents.ToolRegistry

        Public Property AuthoritativeToolRegistry As SharedLibrary.Agents.ToolRegistry

        ''' <summary>True when only a lightweight tool index is initially exposed to the model.</summary>
        Public Property LazyToolLoadingEnabled As Boolean

        ''' <summary>All responses generated during this session (successful and failed).</summary>
        Public Property AllToolResponses As List(Of ToolResponse)

        ''' <summary>Current iteration counter within <see cref="ExecuteToolingLoop"/>.</summary>
        Public Property CurrentIteration As Integer

        ''' <summary>Maximum permitted number of iterations.</summary>
        Public Property MaxIterations As Integer

        ''' <summary>Cancellation flag set by UI event handler.</summary>
        Public Property IsCancelled As Boolean

        ''' <summary>In-memory log entries appended during session execution.</summary>
        Public Property LogEntries As List(Of String)

        ''' <summary>Snapshot of the LLM/tooling model config used for tool call detection/extraction formats.</summary>
        Public Property ToolingModel As ModelConfig

        ''' <summary>Optional UI log window instance used for user-visible progress logging.</summary>
        Public Property LogWindowForm As LogWindow

        Public Property LastToolExecutionSignature As String
        Public Property LastToolExecutionRepeatCount As Integer
        Public Property DuplicateToolExecutionAbortThreshold As Integer

        Public Property FailedToolCallCounts As Dictionary(Of String, Integer)
        Public Property DuplicateFailureAbortThreshold As Integer

        Public Property ConsecutiveFailedToolName As String
        Public Property ConsecutiveFailedToolCount As Integer
        Public Property ConsecutiveToolFailureAbortThreshold As Integer

        Public Property PrematureTextRetryCount As Integer = 0

        Public Const MaxContinuationRetries As Integer = 5
        Public Const MaxEmptyResponseRetries As Integer = 1

        Public Property PendingRejectedAssistantTurn As String = ""
        Public Property LastInvalidAssistantTurnSignature As String = ""
        Public Property LastInvalidTurnReason As String = ""
        Public Property LastInvalidTurnRepeatCount As Integer = 0
        Public Property ForceNoToolFinalizationRequested As Boolean
        Public Property ForceNoToolFinalizationReason As String = ""
        Public Property PendingContinuationGuardPrompt As String
        Public Property PendingGuardTitle As String = ""
        Public Property PendingRejectedTurnExplanation As String = ""

        Public Property HostKind As String
        Public Property AllowedToolNames As HashSet(Of String)
        Public Property EnforceAllowedToolScope As Boolean
        Public Property EmptyMainModelResponse As Boolean

        Public Property SubAgentEmptyResponseRetryCount As Integer

        Public Property SequencingState As SharedLibrary.Agents.ToolCallSequencing.ToolingRunState
        Public Property FinalizationBlocked As Boolean
        Public Property FinalizationBlockedReason As String

        Public Property RunId As String
        Public Property SubAgentInvocationCount As Integer
        Public Property SubAgentInvocationCountsByAgent As Dictionary(Of String, Integer)

        Public Property AuthoritativeToolRegistrySnapshot As SharedLibrary.Agents.ToolRegistry

        Public Property WorkflowId As String
        Public Property RuntimeState As SharedLibrary.Agents.WorkflowRuntimeState
        Public Property LatestUserRequestRaw As String
        Public Property HostTaskSummary As String

        Public Property FinalResponseContract As SharedLibrary.Agents.ToolingFinalResponseContract =
            SharedLibrary.Agents.ToolingFinalResponseContract.UserFacingTaskStatus

        ''' <summary>
        ''' Initializes a new tool execution context with default collections and limits.
        ''' </summary>
        Public Sub New()
            SelectedTools = New List(Of ModelConfig)()
            AllToolResponses = New List(Of ToolResponse)()
            LogEntries = New List(Of String)()
            FailedToolCallCounts = New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
            DuplicateFailureAbortThreshold = 2
            CurrentIteration = 0
            MaxIterations = INI_ToolingMaximumIterations
            IsCancelled = False
            LastToolExecutionSignature = ""
            LastToolExecutionRepeatCount = 0
            DuplicateToolExecutionAbortThreshold = 3
            ConsecutiveFailedToolName = ""
            ConsecutiveFailedToolCount = 0
            ConsecutiveToolFailureAbortThreshold = 3
            PendingContinuationGuardPrompt = ""
            ZeroChangeOperationCounts = New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
            AllowedToolNames = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            EnforceAllowedToolScope = False
            EmptyMainModelResponse = False
            SubAgentEmptyResponseRetryCount = 0
            SequencingState = New SharedLibrary.Agents.ToolCallSequencing.ToolingRunState()
            FinalizationBlocked = False
            FinalizationBlockedReason = ""
            RunId = Guid.NewGuid().ToString("N")
            SubAgentInvocationCount = 0
            SubAgentInvocationCountsByAgent = New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
            AuthoritativeToolRegistry = Nothing
            AuthoritativeToolRegistrySnapshot = Nothing
            LastInvalidAssistantTurnSignature = ""
            LastInvalidTurnReason = ""
            LastInvalidTurnRepeatCount = 0
            ForceNoToolFinalizationRequested = False
            ForceNoToolFinalizationReason = ""
        End Sub

        ''' <summary>Per-target counts of transport-successful but zero-change (no-op) tool results this run.</summary>
        Public Property ZeroChangeOperationCounts As Dictionary(Of String, Integer)

        ''' <summary>Bounded number of no-op attempts against the same target before the edit is marked unresolved.</summary>
        Public Const ZeroChangeOperationAbortThreshold As Integer = 3

        ''' <summary>Per-run set of already-expanded (ref + range) keys used to suppress repeated no-progress reads.</summary>
        Public Property ExpandedContextKeys As HashSet(Of String) =
            New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        Public Property LogPrefix As String
        Public Property ExternalLogSink As Action(Of String, String)
        Public Property ProgressSink As Action(Of String)

        Private _lastProgressText As String = ""
        Private _lastProgressHeading As String = ""
        Private _lastEmittedProgress As String = ""

        ''' <summary>
        ''' Forwards the host-derived per-step progress note (A) to the optional progress sink.
        ''' This channel is always emitted on its own and never depends on a major-step heading (B1)
        ''' ever being present, so A still reaches the user even if no report_progress call occurs.
        ''' </summary>
        Public Sub ReportProgress(text As String)
            Dim a As String = If(text, "").Trim()
            If a = "" Then Return
            _lastProgressText = a
            EmitProgress()
        End Sub

        ''' <summary>
        ''' Records the model-authored major-step heading (B1) emitted via the report_progress tool.
        ''' The heading is optional: when present it prefixes the current A note; when absent the A note
        ''' is shown on its own. This method never suppresses the A channel.
        ''' </summary>
        Public Sub ReportMajorProgress(heading As String)
            Dim h As String = If(heading, "").Trim()
            If h = "" Then Return
            _lastProgressHeading = h
            EmitProgress()
        End Sub

        ''' <summary>
        ''' Composes and forwards the combined progress line, skipping empty text and consecutive
        ''' duplicates. When a B1 heading exists it becomes the lead statement and the English A note
        ''' is shown on the next indented line; otherwise the A note is emitted standalone. Never throws.
        ''' </summary>
        Private Sub EmitProgress()
            If ProgressSink Is Nothing Then Return

            Dim msg As String
            If _lastProgressHeading <> "" AndAlso _lastProgressText <> "" Then
                msg = _lastProgressHeading & vbCrLf & "    " & _lastProgressText
            ElseIf _lastProgressHeading <> "" Then
                msg = _lastProgressHeading
            Else
                msg = _lastProgressText
            End If

            msg = msg.Trim()
            If msg = "" Then Return
            If String.Equals(_lastEmittedProgress, msg, StringComparison.Ordinal) Then Return
            _lastEmittedProgress = msg

            Try
                ProgressSink.Invoke(msg)
            Catch ex As Exception
                ToolingFileLogger.LogWarn("Failed to forward progress entry.", ex:=ex)
            End Try
        End Sub

        Private _lastVisibleLogMessage As String = ""
        Private _lastVisibleLogLevel As String = ""

        Private Function TryRegisterVisibleLog(message As String, level As String) As Boolean
            Dim normalizedMessage As String = If(message, "").Trim()
            Dim normalizedLevel As String = If(level, "step").Trim().ToLowerInvariant()

            If normalizedMessage = "" Then
                Return False
            End If

            If String.Equals(_lastVisibleLogMessage, normalizedMessage, StringComparison.Ordinal) AndAlso
               String.Equals(_lastVisibleLogLevel, normalizedLevel, StringComparison.Ordinal) Then
                Return False
            End If

            _lastVisibleLogMessage = normalizedMessage
            _lastVisibleLogLevel = normalizedLevel
            Return True
        End Function

        Private Sub AppendVisibleLogInternal(message As String, level As String)
            If Not TryRegisterVisibleLog(message, level) Then
                Return
            End If

            If LogWindowForm IsNot Nothing AndAlso Not LogWindowForm.IsDisposed Then
                Try
                    LogWindowForm.AppendLog(message, level)
                Catch ex As Exception
                    ToolingFileLogger.LogWarn("Failed to append to LogWindow.", ex:=ex)
                End Try
            End If

            If ExternalLogSink IsNot Nothing Then
                Try
                    ExternalLogSink.Invoke(message, level)
                Catch ex As Exception
                    ToolingFileLogger.LogWarn("Failed to forward log entry.", ex:=ex)
                End Try
            End If
        End Sub

        Public Sub MirrorVisibleLog(message As String, Optional level As String = "step")
            Dim normalizedLevel As String = If(level, "step").Trim().ToLowerInvariant()
            Dim visibleMessage As String = If(message, "")

            AppendVisibleLogInternal(visibleMessage, normalizedLevel)
        End Sub

        Public Sub Log(message As String, Optional level As String = "step")
            Dim normalizedPrefix As String = If(LogPrefix, "").Trim()
            Dim isSubAgent As Boolean =
                normalizedPrefix.StartsWith("[subagent]", StringComparison.OrdinalIgnoreCase)
            Dim leadingMarker As String = If(isSubAgent, "[subagent]", "")
            Dim humanMessage As String = If(message, "").Trim()

            If Not isSubAgent AndAlso normalizedPrefix <> "" Then
                humanMessage = (normalizedPrefix & " " & humanMessage).Trim()
            End If

            Dim fullMessage As String =
                SharedLibrary.Agents.WorkflowContinuity.ComposeWorkflowLogMessage(
                    humanMessage,
                    WorkflowId,
                    If(RuntimeState?.CurrentPhase, ""),
                    hostName:=HostKind,
                    leadingMarker:=leadingMarker)

            Dim entry = $"[{DateTime.Now:HH:mm:ss}] {fullMessage}"
            LogEntries.Add(entry)
            Debug.WriteLine($"[Tooling] {entry}")

            Dim normalizedLevel As String = If(level, "step").Trim().ToLowerInvariant()

            Select Case normalizedLevel
                Case "diag"
                    ToolingFileLogger.LogDiag(fullMessage)
                Case "warn", "warning"
                    ToolingFileLogger.LogWarn(fullMessage)
                Case "error", "err", "fail"
                    ToolingFileLogger.LogError(fullMessage)
                Case Else
                    ToolingFileLogger.LogStep(fullMessage)
            End Select

            If normalizedLevel = "diag" Then Return

            Dim visibleMessage As String =
                If(isSubAgent, "  [sub-agent] " & humanMessage, humanMessage)

            AppendVisibleLogInternal(visibleMessage, normalizedLevel)
        End Sub

        Public Sub LogWarn(message As String,
                           Optional details As String = "",
                           Optional ex As Exception = Nothing,
                           Optional visibleToUser As Boolean = True,
                           Optional userMessage As String = "")
            Dim visibleMessage As String = If(String.IsNullOrWhiteSpace(userMessage), message, userMessage)

            If visibleToUser AndAlso Not String.IsNullOrWhiteSpace(visibleMessage) Then
                Log(visibleMessage, "warn")
            End If

            Dim filePrimaryMessage As String = message

            If visibleToUser AndAlso String.Equals(visibleMessage, message, StringComparison.Ordinal) Then
                filePrimaryMessage = ""
            End If

            If Not String.IsNullOrWhiteSpace(filePrimaryMessage) OrElse
               Not String.IsNullOrWhiteSpace(details) OrElse
               ex IsNot Nothing Then

                ToolingFileLogger.LogWarn(filePrimaryMessage, details, ex)
            End If
        End Sub

        Public Sub LogError(message As String,
                            Optional details As String = "",
                            Optional ex As Exception = Nothing,
                            Optional visibleToUser As Boolean = True,
                            Optional userMessage As String = "")
            Dim visibleMessage As String = If(String.IsNullOrWhiteSpace(userMessage), message, userMessage)

            If visibleToUser AndAlso Not String.IsNullOrWhiteSpace(visibleMessage) Then
                Log(visibleMessage, "error")
            End If

            Dim filePrimaryMessage As String = message

            If visibleToUser AndAlso String.Equals(visibleMessage, message, StringComparison.Ordinal) Then
                filePrimaryMessage = ""
            End If

            If Not String.IsNullOrWhiteSpace(filePrimaryMessage) OrElse
               Not String.IsNullOrWhiteSpace(details) OrElse
               ex IsNot Nothing Then

                ToolingFileLogger.LogError(filePrimaryMessage, details, ex)
            End If
        End Sub

        ''' <summary>
        ''' Placeholder method (intentionally empty) retained by the caller surface.
        ''' </summary>
        Public Sub WriteDebugLog()
            ' Intentionally empty (avoid multiple log files).
        End Sub
    End Class



End Class
