' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved.
' For license to use see https://redink.ai.
'
' =============================================================================
' File: ExplicitOperationRegistry.vb
'
' Purpose:
'   Provides shared, host-agnostic lifecycle tracking for explicitly identified
'   logical tool operations across a tooling run and nested sub-agent executions.
'
'   Its primary purpose is to prevent repeated execution of an operation that has
'   already reached a terminal unresolved or blocked state, while allowing other
'   independent operations to continue.
'
' Operation identity:
'   Logical operations are identified ONLY by an explicit caller-supplied
'   `operation_id`.
'
'   The registry MUST NOT infer that two calls represent the same operation from:
'
'     - tool name
'     - file path
'     - anchor/find text
'     - replacement/comment text
'     - call JSON similarity
'     - prompt similarity
'     - semantic similarity
'     - filename or document identity
'
'   If no operation_id is supplied, this registry does not participate and the
'   existing legacy no-progress/circuit-breaker behavior remains available.
'
' Intended lifecycle:
'
'     Pending
'       |
'       +---- successful structured result ----> Succeeded
'       |
'       +---- repeated no-progress ------------> TerminalUnresolved
'       |
'       +---- explicit non-recoverable block --> TerminalBlocked
'
' Terminal behavior:
'   Once an exact operation_id is terminal, the same operation_id must not be
'   executed again in the same shared run state.
'
'   A changed anchor does NOT create a new logical operation when it is merely a
'   recovery attempt for the same requested edit. The caller must reuse the same
'   operation_id.
'
'   A genuinely new requested operation must receive a different explicit id.
'
' Shared-run behavior:
'   Parent and nested sub-agent tooling loops should reference the same
'   ExplicitOperationRegistry instance. This ensures that an operation exhausted
'   inside an isolated editor/sub-agent remains terminal when control returns to
'   the parent or when another nested invocation is attempted.
'
' Tool contract:
'   Tools that participate SHOULD accept `operation_id`:
'
'     - at top level for a single operation; and/or
'     - inside each item of a batched `tasks` array.
'
'   Structured tool results SHOULD return the same operation_id unchanged for
'   each per-task result.
'
'   Example input:
'
'     {
'       "tasks": [
'         {
'           "operation_id": "edit-17",
'           "find": "...",
'           "text": "..."
'         }
'       ]
'     }
'
'   Example output:
'
'     {
'       "status": "partial",
'       "tasks": [
'         {
'           "operation_id": "edit-17",
'           "applied": false,
'           "reason": "no_match"
'         }
'       ]
'     }
'
' Retry semantics:
'   - The registry counts attempts only for explicit operation ids.
'   - Successful application marks that exact operation as Succeeded.
'   - Repeated structured no-progress may mark it TerminalUnresolved once the
'     configured attempt limit is reached.
'   - A terminal operation is rejected before physical tool execution.
'   - Independent operation ids remain executable.
'
' Relationship to legacy circuit breakers:
'   This class supplements, rather than replaces, existing path-/call-based
'   duplicate and zero-change guards.
'
'   Legacy guards remain useful for tools that have not yet adopted explicit
'   operation ids. ExplicitOperationRegistry is the authoritative mechanism where
'   an operation_id is present.
'
' Scope:
'   This registry tracks logical operation execution only.
'   It does NOT:
'
'     - identify deliverable artifacts
'     - decide file finality
'     - decide storage location
'     - perform delivery
'     - infer sub-agent task similarity
'
'   Artifact identity and delivery are handled separately by ArtifactDelivery.
'
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports System.Linq
Imports Newtonsoft.Json.Linq

Namespace Agents
    Public Enum ExplicitOperationStatus
        Pending = 0
        Succeeded = 1
        TerminalUnresolved = 2
        TerminalBlocked = 3
    End Enum

    Public NotInheritable Class ExplicitOperationRecord
        Public Property OperationId As String = ""
        Public Property AttemptCount As Integer
        Public Property Status As ExplicitOperationStatus = ExplicitOperationStatus.Pending
        Public Property TerminalReason As String = ""
        Public Property UpdatedUtc As DateTime = DateTime.UtcNow
    End Class

    Public NotInheritable Class ExplicitOperationRegistry
        Private ReadOnly _records As New System.Collections.Generic.Dictionary(Of String, ExplicitOperationRecord)(System.StringComparer.Ordinal)
        Private ReadOnly _syncRoot As New Object()

        Public Function IsTerminal(operationId As String) As Boolean
            Dim id As String = If(operationId, "").Trim()
            If id = "" Then Return False

            SyncLock _syncRoot
                Dim record As ExplicitOperationRecord = Nothing
                If Not _records.TryGetValue(id, record) OrElse record Is Nothing Then Return False
                Return IsTerminalStatus(record.Status)
            End SyncLock
        End Function

        Public Function TryGetFirstTerminalOperationId(arguments As System.Collections.Generic.IDictionary(Of String, Object),
                                                       ByRef terminalOperationId As String) As Boolean
            terminalOperationId = ""
            For Each id As String In ExtractOperationIds(arguments)
                If IsTerminal(id) Then
                    terminalOperationId = id
                    Return True
                End If
            Next
            Return False
        End Function

        ''' <summary>
        ''' Returns True when at least one explicit operation id in the current call is not
        ''' terminal. Exact explicit operation lifecycle then takes precedence over the
        ''' generic duplicate-success replay guard.
        ''' </summary>
        Public Function HasAnyNonTerminalOperationId(
            arguments As System.Collections.Generic.IDictionary(Of String, Object)) As Boolean

            For Each id As String In ExtractOperationIds(arguments)
                If Not IsTerminal(id) Then
                    Return True
                End If
            Next

            Return False
        End Function

        ''' <summary>
        ''' For a batched tasks[] call with no top-level operation_id, removes only task
        ''' items whose exact operation_id is already terminal when at least one independent
        ''' task remains executable. If every identified task is terminal, the arguments are
        ''' left untouched so the normal terminal-operation guard rejects the whole call.
        ''' </summary>
        Public Function TryFilterTerminalTaskOperations(
            arguments As System.Collections.Generic.IDictionary(Of String, Object),
            ByRef skippedTerminalOperationIds As System.Collections.Generic.List(Of String)) As Boolean

            skippedTerminalOperationIds = New System.Collections.Generic.List(Of String)()

            If arguments Is Nothing Then Return False

            Dim directOperationId As Object = Nothing
            If arguments.TryGetValue("operation_id", directOperationId) AndAlso
               directOperationId IsNot Nothing AndAlso
               Not System.String.IsNullOrWhiteSpace(directOperationId.ToString()) Then

                Return False
            End If

            Dim tasksValue As Object = Nothing
            If Not arguments.TryGetValue("tasks", tasksValue) OrElse tasksValue Is Nothing Then
                Return False
            End If

            Dim sourceTasks As Newtonsoft.Json.Linq.JArray

            Try
                sourceTasks =
                    TryCast(
                        Newtonsoft.Json.Linq.JToken.FromObject(tasksValue),
                        Newtonsoft.Json.Linq.JArray)
            Catch ex As System.Exception
                Return False
            End Try

            If sourceTasks Is Nothing OrElse sourceTasks.Count = 0 Then
                Return False
            End If

            Dim filteredTasks As New Newtonsoft.Json.Linq.JArray()
            Dim keptCount As Integer = 0

            For Each token As Newtonsoft.Json.Linq.JToken In sourceTasks
                Dim taskObject As Newtonsoft.Json.Linq.JObject =
                    TryCast(token, Newtonsoft.Json.Linq.JObject)

                If taskObject Is Nothing Then
                    filteredTasks.Add(token.DeepClone())
                    keptCount += 1
                    Continue For
                End If

                Dim operationId As String =
                    If(taskObject.Value(Of String)("operation_id"), "").Trim()

                If operationId <> "" AndAlso IsTerminal(operationId) Then
                    AddDistinct(skippedTerminalOperationIds, operationId)
                    Continue For
                End If

                filteredTasks.Add(token.DeepClone())
                keptCount += 1
            Next

            If skippedTerminalOperationIds.Count = 0 OrElse keptCount = 0 Then
                skippedTerminalOperationIds.Clear()
                Return False
            End If

            arguments("tasks") = filteredTasks
            Return True
        End Function

        ''' <summary>
        ''' Adds host diagnostics for terminal task items that were deterministically
        ''' omitted from a partial batch. This does not alter operation state.
        ''' </summary>
        Public Function AnnotateSkippedTerminalOperations(
            responseText As String,
            skippedTerminalOperationIds As System.Collections.Generic.IEnumerable(Of String)) As String

            If skippedTerminalOperationIds Is Nothing Then
                Return If(responseText, "")
            End If

            Dim ids As New System.Collections.Generic.List(Of String)()

            For Each id As String In skippedTerminalOperationIds
                AddDistinct(ids, id)
            Next

            If ids.Count = 0 Then
                Return If(responseText, "")
            End If

            Try
                Dim root As Newtonsoft.Json.Linq.JObject =
                    TryCast(
                        Newtonsoft.Json.Linq.JToken.Parse(If(responseText, "").Trim()),
                        Newtonsoft.Json.Linq.JObject)

                If root Is Nothing Then
                    Return If(responseText, "")
                End If

                Dim skippedIdsJson As New Newtonsoft.Json.Linq.JArray()
                For Each id As String In ids
                    skippedIdsJson.Add(id)
                Next

                root("host_skipped_terminal_operation_ids") = skippedIdsJson

                Return root.ToString(Newtonsoft.Json.Formatting.None)
            Catch ex As System.Exception
                Return If(responseText, "")
            End Try
        End Function

        Public Sub NoteAttempt(operationId As String)
            Dim id As String = If(operationId, "").Trim()
            If id = "" Then Return

            SyncLock _syncRoot
                Dim record As ExplicitOperationRecord = GetOrCreateLocked(id)
                If record Is Nothing OrElse IsTerminalStatus(record.Status) Then Return
                record.AttemptCount += 1
                record.UpdatedUtc = System.DateTime.UtcNow
            End SyncLock
        End Sub

        Public Sub MarkSucceeded(operationId As String)
            SetStatus(operationId, ExplicitOperationStatus.Succeeded, "")
        End Sub

        Public Sub MarkTerminalUnresolved(operationId As String, reason As String)
            SetStatus(operationId, ExplicitOperationStatus.TerminalUnresolved, If(reason, ""))
        End Sub

        Public Sub MarkTerminalBlocked(operationId As String, reason As String)
            SetStatus(operationId, ExplicitOperationStatus.TerminalBlocked, If(reason, ""))
        End Sub

        Public Sub ApplyToolResult(arguments As System.Collections.Generic.IDictionary(Of String, Object),
                                   responseText As String,
                                   maxAttempts As Integer)
            Dim inputIds As System.Collections.Generic.List(Of String) = ExtractOperationIds(arguments)
            If inputIds.Count = 0 Then Return

            Dim allowedIds As New System.Collections.Generic.HashSet(Of String)(inputIds, System.StringComparer.Ordinal)
            Dim handledAnyPerTask As Boolean = False

            Try
                Dim root As Newtonsoft.Json.Linq.JObject = TryCast(Newtonsoft.Json.Linq.JToken.Parse(If(responseText, "").Trim()), Newtonsoft.Json.Linq.JObject)
                If root IsNot Nothing Then
                    Dim tasks As Newtonsoft.Json.Linq.JArray = TryCast(root("tasks"), Newtonsoft.Json.Linq.JArray)
                    If tasks IsNot Nothing Then
                        For Each token As Newtonsoft.Json.Linq.JToken In tasks
                            Dim obj As Newtonsoft.Json.Linq.JObject = TryCast(token, Newtonsoft.Json.Linq.JObject)
                            If obj Is Nothing Then Continue For

                            Dim id As String = If(obj.Value(Of String)("operation_id"), "").Trim()
                            If id = "" OrElse Not allowedIds.Contains(id) Then Continue For

                            handledAnyPerTask = True
                            NoteAttempt(id)

                            Dim applied As Boolean = False
                            Dim appliedToken As Newtonsoft.Json.Linq.JToken = obj("applied")
                            Dim hasApplied As Boolean =
                                appliedToken IsNot Nothing AndAlso
                                appliedToken.Type <> Newtonsoft.Json.Linq.JTokenType.Null AndAlso
                                Boolean.TryParse(appliedToken.ToString(), applied)

                            If hasApplied AndAlso applied Then
                                MarkSucceeded(id)
                            Else
                                Dim attemptCount As Integer = GetAttemptCount(id)
                                If attemptCount >= System.Math.Max(1, maxAttempts) Then
                                    Dim reason As String = If(obj.Value(Of String)("reason"), obj.Value(Of String)("error"))
                                    MarkTerminalUnresolved(id, If(reason, "explicit_operation_no_progress"))
                                End If
                            End If
                        Next
                    End If
                End If
            Catch ex As System.Exception
            End Try

            If handledAnyPerTask Then Return

            If ToolCallSequencing.IsZeroChangeOperationResult(responseText) Then
                For Each id As String In inputIds
                    NoteAttempt(id)
                    If GetAttemptCount(id) >= System.Math.Max(1, maxAttempts) Then
                        MarkTerminalUnresolved(id, "explicit_operation_no_progress")
                    End If
                Next
            End If
        End Sub

        Public Shared Function ExtractOperationIds(arguments As System.Collections.Generic.IDictionary(Of String, Object)) As System.Collections.Generic.List(Of String)
            Dim result As New System.Collections.Generic.List(Of String)()
            If arguments Is Nothing Then Return result

            Dim direct As Object = Nothing
            If arguments.TryGetValue("operation_id", direct) AndAlso direct IsNot Nothing Then
                AddDistinct(result, direct.ToString())
            End If

            Dim tasksValue As Object = Nothing
            If arguments.TryGetValue("tasks", tasksValue) AndAlso tasksValue IsNot Nothing Then
                Try
                    Dim arr As Newtonsoft.Json.Linq.JArray = TryCast(Newtonsoft.Json.Linq.JToken.FromObject(tasksValue), Newtonsoft.Json.Linq.JArray)
                    If arr IsNot Nothing Then
                        For Each token As Newtonsoft.Json.Linq.JToken In arr
                            Dim obj As Newtonsoft.Json.Linq.JObject = TryCast(token, Newtonsoft.Json.Linq.JObject)
                            If obj Is Nothing Then Continue For
                            AddDistinct(result, If(obj.Value(Of String)("operation_id"), ""))
                        Next
                    End If
                Catch ex As System.Exception
                End Try
            End If

            Return result
        End Function

        Private Sub SetStatus(operationId As String, status As ExplicitOperationStatus, reason As String)
            Dim id As String = If(operationId, "").Trim()
            If id = "" Then Return

            SyncLock _syncRoot
                Dim record As ExplicitOperationRecord = GetOrCreateLocked(id)
                If record Is Nothing Then Return

                ' Terminal status is monotonic. A later result may not reopen or replace it.
                If IsTerminalStatus(record.Status) Then Return

                record.Status = status
                record.TerminalReason = If(reason, "")
                record.UpdatedUtc = System.DateTime.UtcNow
            End SyncLock
        End Sub

        Private Function GetAttemptCount(operationId As String) As Integer
            Dim id As String = If(operationId, "").Trim()
            If id = "" Then Return 0

            SyncLock _syncRoot
                Dim record As ExplicitOperationRecord = Nothing
                If Not _records.TryGetValue(id, record) OrElse record Is Nothing Then Return 0
                Return record.AttemptCount
            End SyncLock
        End Function

        Private Function GetOrCreateLocked(operationId As String) As ExplicitOperationRecord
            Dim id As String = If(operationId, "").Trim()
            If id = "" Then Return Nothing

            Dim record As ExplicitOperationRecord = Nothing
            If Not _records.TryGetValue(id, record) OrElse record Is Nothing Then
                record = New ExplicitOperationRecord With {.OperationId = id}
                _records(id) = record
            End If
            Return record
        End Function

        Private Shared Function IsTerminalStatus(status As ExplicitOperationStatus) As Boolean
            Return status = ExplicitOperationStatus.Succeeded OrElse
                   status = ExplicitOperationStatus.TerminalUnresolved OrElse
                   status = ExplicitOperationStatus.TerminalBlocked
        End Function

        Private Shared Sub AddDistinct(target As System.Collections.Generic.List(Of String), value As String)
            Dim id As String = If(value, "").Trim()
            If id = "" Then Return
            If Not target.Any(Function(x As String) System.String.Equals(x, id, System.StringComparison.Ordinal)) Then target.Add(id)
        End Sub
    End Class
End Namespace
