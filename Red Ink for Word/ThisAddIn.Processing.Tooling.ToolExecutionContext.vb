' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.Processing.Tooling.ToolExecutionContext.vb
' Purpose: Execution state container for tool session lifecycle and diagnostics.
'
' Responsibilities:
'  - Hold per-session state: selected tools, allowed registries, iteration counts.
'  - Track all tool responses (successful and failed) during session.
'  - Manage continuation guards and retry prompts for failed turns.
'  - Log session events with optional UI window display.
'  - Track duplicate tool execution detection (signature + repeat count).
'  - Track consecutive tool failures for abort thresholds.
'  - Hold premature text response retry state.
'  - Maintain sequencing state (tool ordering, deliverable requirements, memory grounding).
'  - Expose finalization blocked state and reason codes.
'  - Integrate with workflow continuity (WorkflowId, RuntimeState).
'  - Provide diagnostic snapshots for logs.
'
' Architecture:
'  - Mutable state container passed through ExecuteToolingLoop iterations.
'  - Nested logging via optional LogWindow (WinForms Form).
'  - External log sink support for sub-agent integration.
'  - Exposes SequencingState for complex orchestration decisions.
'
' External Dependencies:
'  - SharedLibrary.Agents.ToolCallSequencing for sequencing state.
'  - SharedLibrary.Agents.WorkflowContinuity for workflow state.
'  - LogWindow for optional UI logging display.
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

        Public Property ConsecutiveFailedToolName As String
        Public Property ConsecutiveFailedToolCount As Integer
        Public Property ConsecutiveToolFailureAbortThreshold As Integer

        Public Property PrematureTextRetryCount As Integer = 0
        Public Property PendingContinuationGuardPrompt As String = ""
        Public Property PendingRejectedAssistantTurn As String = ""
        Public Property LastInvalidAssistantTurnSignature As String = ""
        Public Property LastInvalidTurnReason As String = ""
        Public Property LastInvalidTurnRepeatCount As Integer = 0
        Public Property ForceNoToolFinalizationRequested As Boolean
        Public Property ForceNoToolFinalizationReason As String = ""
        Public Property PendingGuardTitle As String = ""
        Public Property PendingRejectedTurnExplanation As String = ""

        Public Const MaxContinuationRetries As Integer = 5

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
            CurrentIteration = 0
            MaxIterations = INI_ToolingMaximumIterations
            IsCancelled = False
            LastToolExecutionSignature = ""
            LastToolExecutionRepeatCount = 0
            DuplicateToolExecutionAbortThreshold = 3
            ConsecutiveFailedToolName = ""
            ConsecutiveFailedToolCount = 0
            ConsecutiveToolFailureAbortThreshold = 3
            AllowedToolNames = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            EnforceAllowedToolScope = False
            EmptyMainModelResponse = False
            SequencingState = New SharedLibrary.Agents.ToolCallSequencing.ToolingRunState()
            FinalizationBlocked = False
            FinalizationBlockedReason = ""
            RunId = Guid.NewGuid().ToString("N")
            SubAgentInvocationCount = 0
            SubAgentInvocationCountsByAgent = New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
            AuthoritativeToolRegistry = Nothing
            AuthoritativeToolRegistrySnapshot = Nothing
            SubAgentEmptyResponseRetryCount = 0
            LastInvalidAssistantTurnSignature = ""
            LastInvalidTurnReason = ""
            LastInvalidTurnRepeatCount = 0
            ForceNoToolFinalizationRequested = False
            ForceNoToolFinalizationReason = ""
        End Sub

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
                msg = _lastProgressHeading & " [" & _lastProgressText & "]"
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

        ''' <summary>
        ''' Appends a message to the in-memory log, debugger output, optional UI log window, and the tooling file log.
        ''' </summary>
        ''' <param name="message">Log message.</param>
        ''' <param name="level">Log level passed to the UI log window.</param>

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

            If LogWindowForm IsNot Nothing AndAlso Not LogWindowForm.IsDisposed Then
                Try
                    LogWindowForm.AppendLog(visibleMessage, normalizedLevel)
                Catch ex As Exception
                    ToolingFileLogger.LogWarn("Failed to append to LogWindow.", ex:=ex)
                End Try
            End If

            If ExternalLogSink IsNot Nothing Then
                Try
                    ExternalLogSink.Invoke(visibleMessage, normalizedLevel)
                Catch ex As Exception
                    ToolingFileLogger.LogWarn("Failed to forward log entry.", ex:=ex)
                End Try
            End If
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
