' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved.
' For license to use see https://redink.ai.
'
' =============================================================================
' File: ExplicitSubAgentTaskRegistry.vb
'
' Purpose:
'   Tracks explicitly identified logical sub-agent tasks within one tooling run.
'   It tracks bounded execution attempts for one logical sub-agent task. A failed
'   attempt may be retried under the SAME exact task id; completion and exhausted
'   retry budgets are terminal.
'
' Identity contract:
'   - Task identity comes ONLY from an explicit caller-supplied SubAgentTaskId.
'   - No prompt, filename, path, agent-input, anchor, or semantic similarity is
'     used to infer whether two sub-agent calls are the same task.
'   - The same logical sub-agent task must reuse the same SubAgentTaskId.
'   - A genuinely new independent sub-agent task must use a new SubAgentTaskId.
'
' Scope:
'   This registry guards sub-agent INVOCATIONS. Per-operation retry state inside
'   an agent remains handled separately by ExplicitOperationRegistry.
'
' Lifecycle:
'       Active attempt
'         |
'         +--> Completed
'         +--> RetryableFailure --> Active next attempt
'         +--> TerminalUnresolved / TerminalBlocked (retry budget exhausted)
'
'   Runner-internal continuation retries do not consume another logical attempt.
'   Parent retries do consume the bounded per-task attempt budget.
'
' =============================================================================

Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic

Namespace Agents

    Public Enum ExplicitSubAgentTaskStatus
        Active = 0
        RetryableFailure = 1
        Completed = 2
        TerminalUnresolved = 3
        TerminalBlocked = 4
    End Enum

    Public NotInheritable Class ExplicitSubAgentTaskRecord
        Public Property TaskId As String = ""
        Public Property AgentName As String = ""
        Public Property Status As ExplicitSubAgentTaskStatus =
            ExplicitSubAgentTaskStatus.Active
        Public Property TerminalReason As String = ""
        Public Property AttemptCount As System.Int32 = 0
        Public Property UpdatedUtc As DateTime = System.DateTime.UtcNow
    End Class

    Public NotInheritable Class ExplicitSubAgentTaskRegistry

        Public Const MaxAttemptsPerTask As System.Int32 = 3

        Private ReadOnly _records As New Dictionary(Of String, ExplicitSubAgentTaskRecord)(
            System.StringComparer.Ordinal)
        Private ReadOnly _syncRoot As New Object()

        Public Function IsTerminal(agentName As String, taskId As String) As Boolean
            Dim key As String = BuildKey(agentName, taskId)
            If key = "" Then Return False

            SyncLock _syncRoot
                Dim record As ExplicitSubAgentTaskRecord = Nothing
                If Not _records.TryGetValue(key, record) OrElse record Is Nothing Then
                    Return False
                End If

                Return record.Status = ExplicitSubAgentTaskStatus.Completed OrElse
                       record.Status = ExplicitSubAgentTaskStatus.TerminalUnresolved OrElse
                       record.Status = ExplicitSubAgentTaskStatus.TerminalBlocked
            End SyncLock
        End Function

        Public Function TryBegin(agentName As String, taskId As String, Optional allowActiveContinuation As Boolean = False) As Boolean
            Dim key As String = BuildKey(agentName, taskId)
            If key = "" Then Return False

            SyncLock _syncRoot
                Dim record As ExplicitSubAgentTaskRecord = Nothing
                If Not _records.TryGetValue(key, record) OrElse record Is Nothing Then
                    _records(key) =
                        New ExplicitSubAgentTaskRecord With {
                            .TaskId = If(taskId, "").Trim(),
                            .AgentName = If(agentName, "").Trim(),
                            .Status = ExplicitSubAgentTaskStatus.Active,
                            .AttemptCount = 1,
                            .UpdatedUtc = System.DateTime.UtcNow
                        }
                    Return True
                End If

                If record.Status = ExplicitSubAgentTaskStatus.Active AndAlso allowActiveContinuation Then
                    ' SubAgentRunner owns its internal continuation retry. It is still
                    ' part of the same parent-level logical attempt and consumes no new slot.
                    record.UpdatedUtc = System.DateTime.UtcNow
                    Return True
                End If

                If record.Status = ExplicitSubAgentTaskStatus.RetryableFailure AndAlso
                   record.AttemptCount < MaxAttemptsPerTask Then

                    record.Status = ExplicitSubAgentTaskStatus.Active
                    record.AttemptCount += 1
                    record.TerminalReason = ""
                    record.UpdatedUtc = System.DateTime.UtcNow
                    Return True
                End If

                Return False
            End SyncLock
        End Function

        Public Function IsRetryable(agentName As String, taskId As String) As Boolean
            Dim key As String = BuildKey(agentName, taskId)
            If key = "" Then Return False

            SyncLock _syncRoot
                Dim record As ExplicitSubAgentTaskRecord = Nothing
                If Not _records.TryGetValue(key, record) OrElse record Is Nothing Then Return False
                Return record.Status = ExplicitSubAgentTaskStatus.RetryableFailure AndAlso
                       record.AttemptCount < MaxAttemptsPerTask
            End SyncLock
        End Function

        Public Function GetRemainingAttempts(agentName As String, taskId As String) As System.Int32
            Dim key As String = BuildKey(agentName, taskId)
            If key = "" Then Return 0

            SyncLock _syncRoot
                Dim record As ExplicitSubAgentTaskRecord = Nothing
                If Not _records.TryGetValue(key, record) OrElse record Is Nothing Then Return MaxAttemptsPerTask
                If record.Status = ExplicitSubAgentTaskStatus.Completed OrElse
                   record.Status = ExplicitSubAgentTaskStatus.TerminalUnresolved OrElse
                   record.Status = ExplicitSubAgentTaskStatus.TerminalBlocked Then Return 0
                Return System.Math.Max(0, MaxAttemptsPerTask - record.AttemptCount)
            End SyncLock
        End Function

        Public Sub MarkActive(agentName As String, taskId As String)
            Dim key As String = BuildKey(agentName, taskId)
            If key = "" Then Return

            SyncLock _syncRoot
                If Not _records.ContainsKey(key) Then
                    _records(key) =
                        New ExplicitSubAgentTaskRecord With {
                            .TaskId = If(taskId, "").Trim(),
                            .AgentName = If(agentName, "").Trim(),
                            .Status = ExplicitSubAgentTaskStatus.Active,
                            .AttemptCount = 1,
                            .UpdatedUtc = System.DateTime.UtcNow
                        }
                End If
            End SyncLock
        End Sub

        Public Sub MarkCompleted(agentName As String, taskId As String)
            SetTerminal(
                agentName,
                taskId,
                ExplicitSubAgentTaskStatus.Completed,
                "")
        End Sub

        Public Sub MarkUnresolved(agentName As String,
                                  taskId As String,
                                  reason As String,
                                  Optional forceTerminal As Boolean = False)
            SetFailure(
                agentName,
                taskId,
                ExplicitSubAgentTaskStatus.TerminalUnresolved,
                reason,
                forceTerminal)
        End Sub

        Public Sub MarkBlocked(agentName As String,
                               taskId As String,
                               reason As String,
                               Optional forceTerminal As Boolean = False)
            SetFailure(
                agentName,
                taskId,
                ExplicitSubAgentTaskStatus.TerminalBlocked,
                reason,
                forceTerminal)
        End Sub

        Private Sub SetFailure(agentName As String,
                               taskId As String,
                               exhaustedStatus As ExplicitSubAgentTaskStatus,
                               reason As String,
                               forceTerminal As Boolean)
            Dim key As String = BuildKey(agentName, taskId)
            If key = "" Then Return

            SyncLock _syncRoot
                Dim record As ExplicitSubAgentTaskRecord = Nothing
                If Not _records.TryGetValue(key, record) OrElse record Is Nothing Then
                    record =
                        New ExplicitSubAgentTaskRecord With {
                            .TaskId = If(taskId, "").Trim(),
                            .AgentName = If(agentName, "").Trim(),
                            .Status = ExplicitSubAgentTaskStatus.Active,
                            .AttemptCount = 1
                        }
                    _records(key) = record
                End If

                If record.Status = ExplicitSubAgentTaskStatus.Completed OrElse
                   record.Status = ExplicitSubAgentTaskStatus.TerminalUnresolved OrElse
                   record.Status = ExplicitSubAgentTaskStatus.TerminalBlocked Then Return

                If forceTerminal OrElse record.AttemptCount >= MaxAttemptsPerTask Then
                    record.Status = exhaustedStatus
                Else
                    record.Status = ExplicitSubAgentTaskStatus.RetryableFailure
                End If

                record.TerminalReason = If(reason, "")
                record.UpdatedUtc = System.DateTime.UtcNow
            End SyncLock
        End Sub

        Private Sub SetTerminal(
            agentName As String,
            taskId As String,
            status As ExplicitSubAgentTaskStatus,
            reason As String)

            Dim key As String = BuildKey(agentName, taskId)
            If key = "" Then Return

            SyncLock _syncRoot
                Dim record As ExplicitSubAgentTaskRecord = Nothing

                If Not _records.TryGetValue(key, record) OrElse record Is Nothing Then
                    record =
                        New ExplicitSubAgentTaskRecord With {
                            .TaskId = If(taskId, "").Trim(),
                            .AgentName = If(agentName, "").Trim()
                        }

                    _records(key) = record
                End If

                ' True terminal state is monotonic. Retryable failures are handled by SetFailure.
                If record.Status = ExplicitSubAgentTaskStatus.Completed OrElse
                   record.Status = ExplicitSubAgentTaskStatus.TerminalUnresolved OrElse
                   record.Status = ExplicitSubAgentTaskStatus.TerminalBlocked Then
                    Return
                End If

                record.Status = status
                record.TerminalReason = If(reason, "")
                record.UpdatedUtc = System.DateTime.UtcNow
            End SyncLock
        End Sub

        Private Shared Function BuildKey(agentName As String, taskId As String) As String
            Dim id As String = If(taskId, "").Trim()
            If id = "" Then Return ""

            ' subagent_task_id is the complete opaque identity of the delegated task.
            ' Agent name is metadata only and must not create another identity dimension.
            Return id
        End Function

    End Class

End Namespace
