' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved.
' For license to use see https://redink.ai.
'
' =============================================================================
' File: ExplicitSubAgentTaskRegistry.vb
'
' Purpose:
'   Tracks explicitly identified logical sub-agent tasks within one tooling run.
'   It prevents a completed, unresolved, or blocked sub-agent task from starting
'   another isolated model run under the same exact task id.
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
'       Active
'         |
'         +--> Completed
'         +--> TerminalUnresolved
'         +--> TerminalBlocked
'
'   All three terminal states prevent another isolated sub-agent model run for
'   the same exact SubAgentTaskId in the same ToolingRunState.
'
' =============================================================================

Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic

Namespace Agents

    Public Enum ExplicitSubAgentTaskStatus
        Active = 0
        Completed = 1
        TerminalUnresolved = 2
        TerminalBlocked = 3
    End Enum

    Public NotInheritable Class ExplicitSubAgentTaskRecord
        Public Property TaskId As String = ""
        Public Property AgentName As String = ""
        Public Property Status As ExplicitSubAgentTaskStatus =
            ExplicitSubAgentTaskStatus.Active
        Public Property TerminalReason As String = ""
        Public Property UpdatedUtc As DateTime = System.DateTime.UtcNow
    End Class

    Public NotInheritable Class ExplicitSubAgentTaskRegistry

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
                    _records(key) = New ExplicitSubAgentTaskRecord With {.TaskId = If(taskId, "").Trim(), .AgentName = If(agentName, "").Trim(), .Status = ExplicitSubAgentTaskStatus.Active, .UpdatedUtc = System.DateTime.UtcNow}
                    Return True
                End If
                If record.Status = ExplicitSubAgentTaskStatus.Active AndAlso allowActiveContinuation Then
                    record.UpdatedUtc = System.DateTime.UtcNow
                    Return True
                End If
                Return False
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

        Public Sub MarkUnresolved(agentName As String, taskId As String, reason As String)
            SetTerminal(
                agentName,
                taskId,
                ExplicitSubAgentTaskStatus.TerminalUnresolved,
                reason)
        End Sub

        Public Sub MarkBlocked(agentName As String, taskId As String, reason As String)
            SetTerminal(
                agentName,
                taskId,
                ExplicitSubAgentTaskStatus.TerminalBlocked,
                reason)
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

                ' Terminal state is monotonic and cannot be replaced by a later result.
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
