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
'   Logical operations are identified by an explicit caller-supplied `operation_id`.
'   A logical operation may contain multiple concrete steps identified by optional
'   `step_id`. Retries reuse the same operation_id + step_id. Distinct continuations
'   inside the same logical operation use the same operation_id and a new step_id.
'   If step_id is omitted, the operation is treated as one legacy single step.
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
'   Terminal state belongs to one exact operation_id + step_id pair. A terminal
'   step may not be executed again, but another step_id in the same operation remains
'   executable. A changed anchor or reformulation used only to retry the same failed
'   step must reuse both ids; it must not manufacture a new step_id to bypass the guard.
'
'   A genuinely new logical operation receives a new operation_id.
'
' Shared-run behavior:
'   Parent and nested sub-agent tooling loops should reference the same
'   ExplicitOperationRegistry instance. This ensures that an operation exhausted
'   inside an isolated editor/sub-agent remains terminal when control returns to
'   the parent or when another nested invocation is attempted.
'
' Tool contract:
'   Tools that participate SHOULD accept `operation_id` and may accept `step_id`:
'
'     - at top level for a single operation; and/or
'     - inside each item of a batched `tasks` array.
'
'   Structured tool results SHOULD return the same operation_id unchanged for
'   each per-task result and, when they expose step_id, return that unchanged as well.
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
'   - The registry counts attempts per explicit operation_id + step_id.
'   - Successful application marks that exact step as Succeeded.
'   - Repeated structured no-progress may mark that step TerminalUnresolved once the
'     configured attempt limit is reached.
'   - A terminal step is rejected before physical tool execution.
'   - Other step_ids in the same operation and independent operation ids remain executable.
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
        Public Property StepId As String = ""
        Public Property AttemptCount As Integer
        Public Property Status As ExplicitOperationStatus = ExplicitOperationStatus.Pending
        Public Property TerminalReason As String = ""
        Public Property UpdatedUtc As DateTime = DateTime.UtcNow
    End Class


    Public NotInheritable Class ExplicitOperationIdentity
        Public Property OperationId As String = ""
        Public Property StepId As String = ""
    End Class

    Public NotInheritable Class ExplicitOperationRegistry
        Private ReadOnly _records As New System.Collections.Generic.Dictionary(Of String, ExplicitOperationRecord)(System.StringComparer.Ordinal)
        Private ReadOnly _syncRoot As New Object()

        Public Function IsTerminal(operationId As String, Optional stepId As String = "") As Boolean
            Dim id As String = If(operationId, "").Trim()
            If id = "" Then Return False

            SyncLock _syncRoot
                Dim record As ExplicitOperationRecord = Nothing
                If Not _records.TryGetValue(BuildRecordKey(id, stepId), record) OrElse record Is Nothing Then Return False
                Return IsTerminalStatus(record.Status)
            End SyncLock
        End Function

        Public Function IsSucceeded(operationId As String, Optional stepId As String = "") As Boolean
            Dim id As String = If(operationId, "").Trim()
            If id = "" Then Return False

            SyncLock _syncRoot
                Dim record As ExplicitOperationRecord = Nothing
                If Not _records.TryGetValue(BuildRecordKey(id, stepId), record) OrElse record Is Nothing Then Return False
                Return record.Status = ExplicitOperationStatus.Succeeded
            End SyncLock
        End Function

        Public Function TryGetFirstSucceededOperationId(arguments As System.Collections.Generic.IDictionary(Of String, Object),
                                                        ByRef succeededOperationId As String) As Boolean
            succeededOperationId = ""
            For Each identity As ExplicitOperationIdentity In ExtractOperationIdentities(arguments)
                If identity Is Nothing Then Continue For
                If IsSucceeded(identity.OperationId, identity.StepId) Then
                    succeededOperationId = identity.OperationId
                    Return True
                End If
            Next
            Return False
        End Function

        Public Function TryGetFirstTerminalOperationId(arguments As System.Collections.Generic.IDictionary(Of String, Object),
                                                       ByRef terminalOperationId As String) As Boolean
            terminalOperationId = ""
            For Each identity As ExplicitOperationIdentity In ExtractOperationIdentities(arguments)
                If identity Is Nothing Then Continue For
                If IsTerminal(identity.OperationId, identity.StepId) Then
                    terminalOperationId = identity.OperationId
                    Return True
                End If
            Next
            Return False
        End Function

        ''' <summary>
        ''' Returns True when at least one explicit operation step in the current call is not
        ''' terminal. An operation may contain multiple independently identified steps; a
        ''' successful/terminal step therefore does not make the whole operation terminal.
        ''' </summary>
        Public Function HasAnyNonTerminalOperationId(
            arguments As System.Collections.Generic.IDictionary(Of String, Object)) As Boolean

            For Each identity As ExplicitOperationIdentity In ExtractOperationIdentities(arguments)
                If identity Is Nothing Then Continue For
                If Not IsTerminal(identity.OperationId, identity.StepId) Then
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
                Dim stepId As String =
                    If(taskObject.Value(Of String)("step_id"), "").Trim()

                If operationId <> "" AndAlso IsTerminal(operationId, stepId) Then
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

        Public Sub NoteAttempt(operationId As String, Optional stepId As String = "")
            Dim id As String = If(operationId, "").Trim()
            If id = "" Then Return

            SyncLock _syncRoot
                Dim record As ExplicitOperationRecord = GetOrCreateLocked(id, stepId)
                If record Is Nothing OrElse IsTerminalStatus(record.Status) Then Return
                record.AttemptCount += 1
                record.UpdatedUtc = System.DateTime.UtcNow
            End SyncLock
        End Sub

        Public Sub MarkSucceeded(operationId As String, Optional stepId As String = "")
            SetStatus(operationId, stepId, ExplicitOperationStatus.Succeeded, "")
        End Sub

        Public Sub MarkTerminalUnresolved(operationId As String, reason As String, Optional stepId As String = "")
            SetStatus(operationId, stepId, ExplicitOperationStatus.TerminalUnresolved, If(reason, ""))
        End Sub

        Public Sub MarkTerminalBlocked(operationId As String, reason As String, Optional stepId As String = "")
            SetStatus(operationId, stepId, ExplicitOperationStatus.TerminalBlocked, If(reason, ""))
        End Sub

        Public Sub ApplyToolResult(arguments As System.Collections.Generic.IDictionary(Of String, Object),
                                   responseText As String,
                                   maxAttempts As Integer)
            Dim inputIdentities As System.Collections.Generic.List(Of ExplicitOperationIdentity) = ExtractOperationIdentities(arguments)
            If inputIdentities.Count = 0 Then Return

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
                            If id = "" Then Continue For

                            Dim resultStepId As String = If(obj.Value(Of String)("step_id"), "").Trim()
                            Dim identity As ExplicitOperationIdentity = ResolveInputIdentity(inputIdentities, id, resultStepId)
                            If identity Is Nothing Then Continue For

                            handledAnyPerTask = True
                            NoteAttempt(identity.OperationId, identity.StepId)

                            Dim applied As Boolean = False
                            Dim appliedToken As Newtonsoft.Json.Linq.JToken = obj("applied")
                            Dim hasApplied As Boolean =
                                appliedToken IsNot Nothing AndAlso
                                appliedToken.Type <> Newtonsoft.Json.Linq.JTokenType.Null AndAlso
                                Boolean.TryParse(appliedToken.ToString(), applied)

                            If hasApplied AndAlso applied Then
                                MarkSucceeded(identity.OperationId, identity.StepId)
                            Else
                                Dim attemptCount As Integer = GetAttemptCount(identity.OperationId, identity.StepId)
                                If attemptCount >= System.Math.Max(1, maxAttempts) Then
                                    Dim reason As String = If(obj.Value(Of String)("reason"), obj.Value(Of String)("error"))
                                    MarkTerminalUnresolved(identity.OperationId, If(reason, "explicit_operation_no_progress"), identity.StepId)
                                End If
                            End If
                        Next
                    End If
                End If
            Catch ex As System.Exception
            End Try

            If handledAnyPerTask Then Return

            If ToolCallSequencing.IsZeroChangeOperationResult(responseText) Then
                For Each identity As ExplicitOperationIdentity In inputIdentities
                    If identity Is Nothing Then Continue For
                    NoteAttempt(identity.OperationId, identity.StepId)
                    If GetAttemptCount(identity.OperationId, identity.StepId) >= System.Math.Max(1, maxAttempts) Then
                        MarkTerminalUnresolved(identity.OperationId, "explicit_operation_no_progress", identity.StepId)
                    End If
                Next
            End If
        End Sub

        Public Shared Function ExtractOperationIds(arguments As System.Collections.Generic.IDictionary(Of String, Object)) As System.Collections.Generic.List(Of String)
            Dim result As New System.Collections.Generic.List(Of String)()
            For Each identity As ExplicitOperationIdentity In ExtractOperationIdentities(arguments)
                If identity Is Nothing Then Continue For
                AddDistinct(result, identity.OperationId)
            Next
            Return result
        End Function

        Public Shared Function ExtractOperationIdentities(arguments As System.Collections.Generic.IDictionary(Of String, Object)) As System.Collections.Generic.List(Of ExplicitOperationIdentity)
            Dim result As New System.Collections.Generic.List(Of ExplicitOperationIdentity)()
            If arguments Is Nothing Then Return result

            Dim direct As Object = Nothing
            If arguments.TryGetValue("operation_id", direct) AndAlso direct IsNot Nothing Then
                Dim operationId As String = direct.ToString().Trim()
                If operationId <> "" Then
                    Dim stepValue As Object = Nothing
                    Dim stepId As String = ""
                    If arguments.TryGetValue("step_id", stepValue) AndAlso stepValue IsNot Nothing Then
                        stepId = stepValue.ToString().Trim()
                    End If
                    AddIdentityDistinct(result, operationId, stepId)
                End If
            End If

            Dim tasksValue As Object = Nothing
            If arguments.TryGetValue("tasks", tasksValue) AndAlso tasksValue IsNot Nothing Then
                Try
                    Dim arr As Newtonsoft.Json.Linq.JArray = TryCast(Newtonsoft.Json.Linq.JToken.FromObject(tasksValue), Newtonsoft.Json.Linq.JArray)
                    If arr IsNot Nothing Then
                        For Each token As Newtonsoft.Json.Linq.JToken In arr
                            Dim obj As Newtonsoft.Json.Linq.JObject = TryCast(token, Newtonsoft.Json.Linq.JObject)
                            If obj Is Nothing Then Continue For
                            AddIdentityDistinct(
                                result,
                                If(obj.Value(Of String)("operation_id"), ""),
                                If(obj.Value(Of String)("step_id"), ""))
                        Next
                    End If
                Catch ex As System.Exception
                End Try
            End If

            Return result
        End Function

        Private Sub SetStatus(operationId As String,
                              stepId As String,
                              status As ExplicitOperationStatus,
                              reason As String)
            Dim id As String = If(operationId, "").Trim()
            If id = "" Then Return

            SyncLock _syncRoot
                Dim record As ExplicitOperationRecord = GetOrCreateLocked(id, stepId)
                If record Is Nothing Then Return

                ' A concrete step is monotonic once it succeeds or becomes terminal. A
                ' different step_id remains independently executable inside the same
                ' logical operation_id.
                If IsTerminalStatus(record.Status) Then Return

                record.Status = status
                record.TerminalReason = If(reason, "")
                record.UpdatedUtc = System.DateTime.UtcNow
            End SyncLock
        End Sub

        Private Function GetAttemptCount(operationId As String, Optional stepId As String = "") As Integer
            Dim id As String = If(operationId, "").Trim()
            If id = "" Then Return 0

            SyncLock _syncRoot
                Dim record As ExplicitOperationRecord = Nothing
                If Not _records.TryGetValue(BuildRecordKey(id, stepId), record) OrElse record Is Nothing Then Return 0
                Return record.AttemptCount
            End SyncLock
        End Function

        Private Function GetOrCreateLocked(operationId As String, Optional stepId As String = "") As ExplicitOperationRecord
            Dim id As String = If(operationId, "").Trim()
            If id = "" Then Return Nothing
            Dim normalizedStepId As String = NormalizeStepId(stepId)
            Dim recordKey As String = BuildRecordKey(id, normalizedStepId)

            Dim record As ExplicitOperationRecord = Nothing
            If Not _records.TryGetValue(recordKey, record) OrElse record Is Nothing Then
                record = New ExplicitOperationRecord With {
                    .OperationId = id,
                    .StepId = normalizedStepId
                }
                _records(recordKey) = record
            End If
            Return record
        End Function

        Private Shared Function IsTerminalStatus(status As ExplicitOperationStatus) As Boolean
            Return status = ExplicitOperationStatus.Succeeded OrElse
                   status = ExplicitOperationStatus.TerminalUnresolved OrElse
                   status = ExplicitOperationStatus.TerminalBlocked
        End Function

        Private Shared Function NormalizeStepId(stepId As String) As String
            Dim normalized As String = If(stepId, "").Trim()
            If normalized = "" Then Return "__default__"
            Return normalized
        End Function

        Private Shared Function BuildRecordKey(operationId As String, Optional stepId As String = "") As String
            Return If(operationId, "").Trim() & ChrW(31) & NormalizeStepId(stepId)
        End Function

        Private Shared Function ResolveInputIdentity(inputIdentities As System.Collections.Generic.IEnumerable(Of ExplicitOperationIdentity),
                                                     operationId As String,
                                                     resultStepId As String) As ExplicitOperationIdentity
            Dim id As String = If(operationId, "").Trim()
            Dim stepId As String = If(resultStepId, "").Trim()
            If id = "" OrElse inputIdentities Is Nothing Then Return Nothing

            Dim candidates As System.Collections.Generic.List(Of ExplicitOperationIdentity) =
                inputIdentities.Where(
                    Function(item As ExplicitOperationIdentity)
                        Return item IsNot Nothing AndAlso
                               System.String.Equals(If(item.OperationId, ""), id, System.StringComparison.Ordinal)
                    End Function).ToList()

            If stepId <> "" Then
                For Each candidate As ExplicitOperationIdentity In candidates
                    If System.String.Equals(NormalizeStepId(candidate.StepId), NormalizeStepId(stepId), System.StringComparison.Ordinal) Then
                        Return candidate
                    End If
                Next
                Return Nothing
            End If

            If candidates.Count = 1 Then Return candidates(0)

            For Each candidate As ExplicitOperationIdentity In candidates
                If NormalizeStepId(candidate.StepId) = "__default__" Then Return candidate
            Next

            Return Nothing
        End Function

        Private Shared Sub AddIdentityDistinct(target As System.Collections.Generic.List(Of ExplicitOperationIdentity),
                                               operationId As String,
                                               stepId As String)
            If target Is Nothing Then Return
            Dim id As String = If(operationId, "").Trim()
            If id = "" Then Return
            Dim normalizedStepId As String = If(stepId, "").Trim()
            Dim recordKey As String = BuildRecordKey(id, normalizedStepId)

            If target.Any(
                Function(existing As ExplicitOperationIdentity)
                    Return existing IsNot Nothing AndAlso
                           System.String.Equals(
                               BuildRecordKey(existing.OperationId, existing.StepId),
                               recordKey,
                               System.StringComparison.Ordinal)
                End Function) Then
                Return
            End If

            target.Add(New ExplicitOperationIdentity With {
                .OperationId = id,
                .StepId = normalizedStepId
            })
        End Sub

        Private Shared Sub AddDistinct(target As System.Collections.Generic.List(Of String), value As String)
            Dim id As String = If(value, "").Trim()
            If id = "" Then Return
            If Not target.Any(Function(x As String) System.String.Equals(x, id, System.StringComparison.Ordinal)) Then target.Add(id)
        End Sub
    End Class

    ''' <summary>
    ''' Capability-driven orchestration contract for tools whose successful physical
    ''' execution completes one explicitly identified logical operation.
    '''
    ''' operation_id and optional step_id are model/host orchestration metadata only. They are
    ''' injected into the model-facing schema for tools carrying <see cref="CapabilityTag"/>,
    ''' but are removed again before the underlying tool implementation is called. This keeps
    ''' existing tool transports and implementations unchanged.
    ''' </summary>
    Public NotInheritable Class ExplicitOperationToolContract
        Public Const CapabilityTag As System.String = "explicit_operation"

        Private Const GuidanceMarker As System.String = "EXPLICIT OPERATION RULE:"

        Private Sub New()
        End Sub

        Public Shared Function HasCapability(tool As SharedLibrary.ModelConfig) As Boolean
            If tool Is Nothing Then Return False
            Return HasCapabilityTag(tool.CapabilityTags, CapabilityTag)
        End Function

        Public Shared Function HasCapabilityTag(capabilityTags As System.String,
                                                requiredTag As System.String) As Boolean
            Dim wanted As System.String = If(requiredTag, "").Trim()
            If wanted = "" OrElse System.String.IsNullOrWhiteSpace(capabilityTags) Then Return False

            Dim normalized As System.String = capabilityTags.Replace(";", ",").Replace(" ", ",")
            For Each rawTag As System.String In normalized.Split(New System.Char() {","c}, System.StringSplitOptions.RemoveEmptyEntries)
                If System.String.Equals(rawTag.Trim(), wanted, System.StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If
            Next

            Return False
        End Function

        ''' <summary>
        ''' Adds orchestration-only operation_id and optional step_id to the model-facing canonical schema.
        ''' Idempotent and intentionally driven only by CapabilityTags.
        ''' </summary>
        Public Shared Function ApplyToModelConfig(tool As SharedLibrary.ModelConfig) As SharedLibrary.ModelConfig
            If tool Is Nothing OrElse Not HasCapability(tool) Then Return tool
            If System.String.IsNullOrWhiteSpace(tool.ToolDefinition) Then Return tool

            Try
                Dim root As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(tool.ToolDefinition)
                Dim parameters As Newtonsoft.Json.Linq.JObject = TryCast(root("parameters"), Newtonsoft.Json.Linq.JObject)
                If parameters Is Nothing Then
                    parameters = New Newtonsoft.Json.Linq.JObject()
                    parameters("type") = "object"
                    root("parameters") = parameters
                End If

                Dim properties As Newtonsoft.Json.Linq.JObject = TryCast(parameters("properties"), Newtonsoft.Json.Linq.JObject)
                If properties Is Nothing Then
                    properties = New Newtonsoft.Json.Linq.JObject()
                    parameters("properties") = properties
                End If

                If properties("operation_id") Is Nothing Then
                    properties("operation_id") =
                        New Newtonsoft.Json.Linq.JObject(
                            New Newtonsoft.Json.Linq.JProperty("type", "string"),
                            New Newtonsoft.Json.Linq.JProperty(
                                "description",
                                "Opaque stable id for the logical operation. Keep it stable across related steps; use a new id only for a genuinely distinct logical operation."))
                End If

                If properties("step_id") Is Nothing Then
                    properties("step_id") =
                        New Newtonsoft.Json.Linq.JObject(
                            New Newtonsoft.Json.Linq.JProperty("type", "string"),
                            New Newtonsoft.Json.Linq.JProperty(
                                "description",
                                "Optional opaque id for one concrete step inside operation_id. Reuse the same step_id for retries of that step; use a new step_id for a distinct continuation/step within the same operation. Omit only for a single-step operation."))
                End If

                Dim required As Newtonsoft.Json.Linq.JArray = TryCast(parameters("required"), Newtonsoft.Json.Linq.JArray)
                If required Is Nothing Then
                    required = New Newtonsoft.Json.Linq.JArray()
                    parameters("required") = required
                End If

                Dim hasRequiredOperationId As Boolean =
                    required.Any(
                        Function(token As Newtonsoft.Json.Linq.JToken)
                            Return System.String.Equals(
                                If(token Is Nothing, "", token.ToString()),
                                "operation_id",
                                System.StringComparison.OrdinalIgnoreCase)
                        End Function)

                If Not hasRequiredOperationId Then
                    required.Add("operation_id")
                End If

                tool.ToolDefinition = root.ToString(Newtonsoft.Json.Formatting.None)

                Dim guidance As System.String =
                    " " & GuidanceMarker &
                    " operation_id is required and identifies one logical operation, not a tool type or file type. " &
                    "A logical operation may contain multiple concrete steps. Use optional step_id to distinguish those steps: " &
                    "reuse the same operation_id+step_id for retries/reformulations of the same failed step, and use a new step_id for a distinct continuation within the same operation. " &
                    "If step_id is omitted, the operation is treated as one legacy single step. " &
                    "Never change operation_id or step_id merely to bypass a failed-step retry or terminal-step guard."

                If System.String.IsNullOrWhiteSpace(tool.ToolInstructionsPrompt) Then
                    tool.ToolInstructionsPrompt = If(tool.ToolName, "") & ":" & guidance
                ElseIf tool.ToolInstructionsPrompt.IndexOf(GuidanceMarker, System.StringComparison.Ordinal) < 0 Then
                    tool.ToolInstructionsPrompt &= guidance
                End If
            Catch ex As System.Exception
                ' Schema augmentation must never make an otherwise valid tool unavailable.
            End Try

            Return tool
        End Function

        ''' <summary>
        ''' Returns a copy of the arguments for physical execution. For participating
        ''' tools, operation_id is intentionally stripped because it belongs to the
        ''' orchestration layer and must not change existing tool transports/contracts.
        ''' </summary>
        Public Shared Function BuildExecutionArguments(
            tool As SharedLibrary.ModelConfig,
            arguments As System.Collections.Generic.IDictionary(Of System.String, Object)) As System.Collections.Generic.Dictionary(Of System.String, Object)

            Dim result As New System.Collections.Generic.Dictionary(Of System.String, Object)(System.StringComparer.OrdinalIgnoreCase)

            If arguments IsNot Nothing Then
                For Each pair As System.Collections.Generic.KeyValuePair(Of System.String, Object) In arguments
                    result(pair.Key) = pair.Value
                Next
            End If

            If HasCapability(tool) Then
                result.Remove("operation_id")
                result.Remove("step_id")
            End If

            Return result
        End Function

        Public Shared Function TryGetTopLevelOperationIdentity(
            arguments As System.Collections.Generic.IDictionary(Of System.String, Object),
            ByRef operationId As System.String,
            ByRef stepId As System.String) As Boolean

            operationId = ""
            stepId = ""
            If arguments Is Nothing Then Return False

            Dim raw As Object = Nothing
            If Not arguments.TryGetValue("operation_id", raw) OrElse raw Is Nothing Then Return False

            operationId = raw.ToString().Trim()
            If operationId = "" Then Return False

            raw = Nothing
            If arguments.TryGetValue("step_id", raw) AndAlso raw IsNot Nothing Then
                stepId = raw.ToString().Trim()
            End If

            Return True
        End Function

        Public Shared Function TryGetTopLevelOperationId(
            arguments As System.Collections.Generic.IDictionary(Of System.String, Object),
            ByRef operationId As System.String) As Boolean

            Dim stepId As System.String = ""
            Return TryGetTopLevelOperationIdentity(arguments, operationId, stepId)
        End Function

        ''' <summary>
        ''' Marks a capability-tagged top-level operation step as succeeded after the host has
        ''' received a successful non-zero-change tool result. Structured mutation tools
        ''' that already use per-task applied semantics remain governed by ApplyToolResult.
        ''' </summary>
        Public Shared Sub MarkSucceededAfterSuccessfulExecution(
            tool As SharedLibrary.ModelConfig,
            arguments As System.Collections.Generic.IDictionary(Of System.String, Object),
            responseText As System.String,
            registry As ExplicitOperationRegistry)

            If registry Is Nothing OrElse Not HasCapability(tool) Then Return
            If ToolCallSequencing.IsZeroChangeOperationResult(responseText) Then Return

            Dim operationId As System.String = ""
            Dim stepId As System.String = ""
            If TryGetTopLevelOperationIdentity(arguments, operationId, stepId) Then
                registry.MarkSucceeded(operationId, stepId)
            End If
        End Sub
    End Class

End Namespace
