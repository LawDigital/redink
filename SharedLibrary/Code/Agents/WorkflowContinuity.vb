' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: WorkflowContinuity.vb
' Purpose:
'   Maintains host-agnostic workflow identity and continuity state across tooling-loop
'   iterations, retries, sub-agents, source grounding and finalization.
'
' Architecture / Function:
'   - Uses AsyncLocal workflow scope so Word, Outlook and Excel can attach one stable
'     WorkflowId and runtime state to nested asynchronous operations.
'   - Tracks tool success/failure, current phase, active skill/agent context, source
'     references, output references, memory-grounding metadata and unresolved failures.
'   - Produces checkpoint/log envelopes for diagnostics and recovery without making
'     host-specific routing decisions itself.
'   - Shared orchestration components update this state; host loops consume it for
'     diagnostics, continuation and completion checks.
' =============================================================================


Option Strict On
Option Explicit On

Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Agents

    Public Class SessionMemoryMetadata
        Public Property WorkflowId As String = ""
        Public Property Source As String = ""
        Public Property ContentKind As String = "unknown"
        Public Property RelatedTool As String = ""
        Public Property RelatedAgent As String = ""
        Public Property RelatedSkill As String = ""
        Public Property CreatedAt As DateTime
        Public Property TrustLevel As String = "advisory"
        Public Property TrustedForRuntime As Boolean = False
    End Class

    Public Class WorkflowSourceRecord
        Public Property SourceId As String = ""
        Public Property WorkflowId As String = ""
        Public Property Title As String = ""
        Public Property Provider As String = ""
        Public Property SourceType As String = ""
        Public Property Reference As String = ""
        Public Property RetrievedAt As DateTime
        Public Property Summary As String = ""
        Public Property RelatedTool As String = ""
        Public Property UsedInOutput As Boolean = False
    End Class

    Public Class WorkflowRuntimeState
        Public Property WorkflowId As String = ""
        Public Property HostPipeline As String = ""
        Public Property ActiveSkillName As String = ""
        Public Property CurrentPhase As String = ""
        Public Property LastSuccessfulTool As String = ""
        Public Property LastFailedTool As String = ""
        Public Property UnresolvedToolFailure As Boolean = False
        Public Property LastStructuredToolResultRef As String = ""
        Public Property LastKnownOutputReference As String = ""
        Public Property LastKnownSourceRefs As List(Of String) = New List(Of String)()
        Public Property ToolCallSuccessCount As Integer = 0
        Public Property ToolCallFailureCount As Integer = 0
        Public Property RetryCount As Integer = 0
        Public Property CreatedAt As DateTime
        Public Property UpdatedAt As DateTime
        Public Property Authoritative As Boolean = True
    End Class

    Friend NotInheritable Class WorkflowCheckpointEnvelope
        Public Property WorkflowId As String = ""
        Public Property HostPipeline As String = ""
        Public Property CheckpointKind As String = ""
        Public Property WrittenAt As DateTime
        Public Property RuntimeState As WorkflowRuntimeState
    End Class

    Public NotInheritable Class WorkflowContinuity

        Private Sub New()
        End Sub

        Private Shared ReadOnly _sync As New Object()
        Private Shared ReadOnly _states As New Dictionary(Of String, WorkflowRuntimeState)(StringComparer.OrdinalIgnoreCase)
        Private Shared ReadOnly _currentWorkflowId As New AsyncLocal(Of String)()
        Private Shared ReadOnly _currentHostPipeline As New AsyncLocal(Of String)()

        Public Shared ReadOnly Property CurrentWorkflowId As String
            Get
                Return If(_currentWorkflowId.Value, "")
            End Get
        End Property

        Public Shared ReadOnly Property CurrentHostPipeline As String
            Get
                Return If(_currentHostPipeline.Value, "")
            End Get
        End Property

        Public Shared Function CreateWorkflowId() As String
            Return "wf_" & Guid.NewGuid().ToString("N")
        End Function

        Public Shared Function BeginWorkflowScope(workflowId As String, hostPipeline As String) As IDisposable
            Return New WorkflowScope(workflowId, hostPipeline)
        End Function

        Private NotInheritable Class WorkflowScope
            Implements IDisposable

            Private ReadOnly _previousWorkflowId As String
            Private ReadOnly _previousHostPipeline As String

            Public Sub New(workflowId As String, hostPipeline As String)
                _previousWorkflowId = If(_currentWorkflowId.Value, "")
                _previousHostPipeline = If(_currentHostPipeline.Value, "")
                _currentWorkflowId.Value = If(workflowId, "")
                _currentHostPipeline.Value = If(hostPipeline, "")
            End Sub

            Public Sub Dispose() Implements IDisposable.Dispose
                _currentWorkflowId.Value = _previousWorkflowId
                _currentHostPipeline.Value = _previousHostPipeline
            End Sub
        End Class

        Public Shared Function BuildWorkflowLogLabel(workflowId As String,
                                                     phase As String,
                                                     Optional toolName As String = "",
                                                     Optional agentName As String = "",
                                                     Optional hostName As String = "") As String
            Dim parts As New List(Of String)()

            If Not String.IsNullOrWhiteSpace(workflowId) Then
                parts.Add("[workflowId: " & workflowId.Trim() & "]")
            End If

            If Not String.IsNullOrWhiteSpace(phase) Then
                parts.Add("[phase: " & phase.Trim() & "]")
            End If

            If Not String.IsNullOrWhiteSpace(toolName) Then
                parts.Add("[tool: " & toolName.Trim() & "]")
            End If

            If Not String.IsNullOrWhiteSpace(agentName) Then
                parts.Add("[agent: " & agentName.Trim() & "]")
            End If

            If Not String.IsNullOrWhiteSpace(hostName) Then
                parts.Add("[host: " & hostName.Trim() & "]")
            End If

            Return String.Join(" ", parts)
        End Function


        Public Shared Function ComposeWorkflowLogMessage(message As String,
                                                         workflowId As String,
                                                         phase As String,
                                                         Optional toolName As String = "",
                                                         Optional agentName As String = "",
                                                         Optional hostName As String = "",
                                                         Optional leadingMarker As String = "") As String
            Dim marker As String = If(leadingMarker, "").Trim()
            Dim coreMessage As String = If(message, "").Trim()
            Dim metadata As String = BuildWorkflowLogLabel(workflowId, phase, toolName, agentName, hostName)

            Dim parts As New List(Of String)()

            If marker <> "" Then
                parts.Add(marker)
            End If

            If coreMessage <> "" Then
                parts.Add(coreMessage)
            End If

            Dim result As String = String.Join(" ", parts).Trim()

            If metadata <> "" Then
                If result <> "" Then
                    result &= " " & metadata
                Else
                    result = metadata
                End If
            End If

            Return result.Trim()
        End Function


        Public Shared Function NormalizeContentKind(value As String) As String
            Select Case If(value, "").Trim().ToLowerInvariant()
                Case "runtime_state", "tool_result", "source_record", "note", "summary", "draft"
                    Return If(value, "").Trim().ToLowerInvariant()
                Case Else
                    Return "unknown"
            End Select
        End Function

        Public Shared Function NormalizeSource(value As String) As String
            Select Case If(value, "").Trim().ToLowerInvariant()
                Case "host", "tool", "agent", "model", "user"
                    Return If(value, "").Trim().ToLowerInvariant()
                Case Else
                    Return "unknown"
            End Select
        End Function

        Public Shared Function NormalizeTrustLevel(value As String) As String
            Select Case If(value, "").Trim().ToLowerInvariant()
                Case "authoritative", "advisory", "validated", "unvalidated"
                    Return If(value, "").Trim().ToLowerInvariant()
                Case Else
                    Return "advisory"
            End Select
        End Function

        Public Shared Function EnsureMetadataDefaults(metadata As SessionMemoryMetadata) As SessionMemoryMetadata
            Dim result As SessionMemoryMetadata = If(metadata, New SessionMemoryMetadata())

            If String.IsNullOrWhiteSpace(result.WorkflowId) Then
                result.WorkflowId = CurrentWorkflowId
            End If

            If result.CreatedAt = DateTime.MinValue Then
                result.CreatedAt = DateTime.UtcNow
            End If

            If String.IsNullOrWhiteSpace(result.Source) Then
                result.Source = If(String.IsNullOrWhiteSpace(result.WorkflowId), "unknown", "model")
            End If

            result.Source = NormalizeSource(result.Source)
            result.ContentKind = NormalizeContentKind(result.ContentKind)

            If result.Source = "host" Then
                result.TrustedForRuntime = True
            End If

            If String.IsNullOrWhiteSpace(result.TrustLevel) Then
                result.TrustLevel = If(result.TrustedForRuntime, "authoritative", "advisory")
            Else
                result.TrustLevel = NormalizeTrustLevel(result.TrustLevel)
            End If

            Return result
        End Function

        Public Shared Function StartWorkflow(workflowId As String, hostPipeline As String) As WorkflowRuntimeState
            If String.IsNullOrWhiteSpace(workflowId) Then
                workflowId = CreateWorkflowId()
            End If

            SyncLock _sync
                Dim state = GetOrLoadUnlocked(workflowId, hostPipeline)
                Dim nowUtc = DateTime.UtcNow

                If state.CreatedAt = DateTime.MinValue Then
                    state.CreatedAt = nowUtc
                End If

                state.UpdatedAt = nowUtc
                state.HostPipeline = If(hostPipeline, "")
                state.CurrentPhase = "workflow_started"
                state.Authoritative = True

                WriteCheckpointUnlocked(state, "workflow_started")
                Return CloneState(state)
            End SyncLock
        End Function

        Public Shared Function AttachWorkflow(workflowId As String, hostPipeline As String) As WorkflowRuntimeState
            If String.IsNullOrWhiteSpace(workflowId) Then
                workflowId = CreateWorkflowId()
            End If

            SyncLock _sync
                Dim state = GetOrLoadUnlocked(workflowId, hostPipeline)

                If state.CreatedAt = DateTime.MinValue Then
                    Dim nowUtc = DateTime.UtcNow
                    state.CreatedAt = nowUtc
                    state.UpdatedAt = nowUtc
                    state.CurrentPhase = "workflow_started"
                    state.Authoritative = True
                    WriteCheckpointUnlocked(state, "workflow_started")
                End If

                Return CloneState(state)
            End SyncLock
        End Function

        Public Shared Function GetState(workflowId As String) As WorkflowRuntimeState
            If String.IsNullOrWhiteSpace(workflowId) Then Return Nothing

            SyncLock _sync
                Dim state = GetOrLoadUnlocked(workflowId, "")
                Return CloneState(state)
            End SyncLock
        End Function

        Public Shared Function NoteSkillLoaded(workflowId As String,
                                               hostPipeline As String,
                                               skillName As String) As Boolean
            If String.IsNullOrWhiteSpace(workflowId) Then Return False

            SyncLock _sync
                Dim state = GetOrLoadUnlocked(workflowId, hostPipeline)
                state.ActiveSkillName = If(skillName, "")
                state.CurrentPhase = "skill_loaded"
                state.UpdatedAt = DateTime.UtcNow
                Return WriteCheckpointUnlocked(state, "skill_loaded")
            End SyncLock
        End Function

        Public Shared Function NoteToolCallResult(workflowId As String,
                                                  hostPipeline As String,
                                                  toolName As String,
                                                  succeeded As Boolean,
                                                  resultRef As String,
                                                  outputReference As String,
                                                  sourceRefs As IEnumerable(Of String),
                                                  retryCount As Integer) As Boolean
            If String.IsNullOrWhiteSpace(workflowId) Then Return False

            SyncLock _sync
                Dim state = GetOrLoadUnlocked(workflowId, hostPipeline)
                state.CurrentPhase = If(succeeded, "tool_call_succeeded", "tool_call_failed")
                state.UpdatedAt = DateTime.UtcNow
                state.RetryCount = Math.Max(0, retryCount)

                If succeeded Then
                    state.LastSuccessfulTool = If(toolName, "")
                    state.ToolCallSuccessCount += 1
                    state.UnresolvedToolFailure = False
                Else
                    state.LastFailedTool = If(toolName, "")
                    state.ToolCallFailureCount += 1
                    state.UnresolvedToolFailure = True
                End If

                If Not String.IsNullOrWhiteSpace(resultRef) Then
                    state.LastStructuredToolResultRef = resultRef
                End If

                If Not String.IsNullOrWhiteSpace(outputReference) Then
                    state.LastKnownOutputReference = outputReference
                End If

                If sourceRefs IsNot Nothing Then
                    state.LastKnownSourceRefs =
                        sourceRefs.
                            Where(Function(x) Not String.IsNullOrWhiteSpace(x)).
                            Select(Function(x) x.Trim()).
                            Distinct(StringComparer.OrdinalIgnoreCase).
                            Take(5).
                            ToList()
                End If

                Return WriteCheckpointUnlocked(state, state.CurrentPhase)
            End SyncLock
        End Function

        Public Shared Function NoteSubAgentInvoked(workflowId As String,
                                                   hostPipeline As String,
                                                   agentName As String) As Boolean
            If String.IsNullOrWhiteSpace(workflowId) Then Return False

            SyncLock _sync
                Dim state = GetOrLoadUnlocked(workflowId, hostPipeline)
                state.CurrentPhase = "sub_agent_invoked"
                state.UpdatedAt = DateTime.UtcNow
                Return WriteCheckpointUnlocked(state, "sub_agent_invoked")
            End SyncLock
        End Function

        Public Shared Function NoteSubAgentReturned(workflowId As String,
                                                    hostPipeline As String,
                                                    agentName As String,
                                                    succeeded As Boolean) As Boolean
            If String.IsNullOrWhiteSpace(workflowId) Then Return False

            SyncLock _sync
                Dim state = GetOrLoadUnlocked(workflowId, hostPipeline)
                state.CurrentPhase = "sub_agent_returned"
                state.UpdatedAt = DateTime.UtcNow
                If Not succeeded Then
                    state.UnresolvedToolFailure = True
                End If
                Return WriteCheckpointUnlocked(state, "sub_agent_returned")
            End SyncLock
        End Function

        Public Shared Function NoteMemoryReferenceCreated(workflowId As String,
                                                          hostPipeline As String,
                                                          memoryKey As String,
                                                          metadata As SessionMemoryMetadata,
                                                          Optional sourceRecord As WorkflowSourceRecord = Nothing) As Boolean
            If String.IsNullOrWhiteSpace(workflowId) Then Return False

            SyncLock _sync
                Dim state = GetOrLoadUnlocked(workflowId, hostPipeline)
                state.UpdatedAt = DateTime.UtcNow

                Dim contentKind As String = NormalizeContentKind(If(metadata?.ContentKind, "unknown"))

                If contentKind = "source_record" Then
                    state.CurrentPhase = "source_reference_created"

                    Dim newRefs As New List(Of String)(state.LastKnownSourceRefs)

                    If sourceRecord IsNot Nothing Then
                        If Not String.IsNullOrWhiteSpace(sourceRecord.SourceId) Then
                            newRefs.Add(sourceRecord.SourceId)
                        ElseIf Not String.IsNullOrWhiteSpace(sourceRecord.Reference) Then
                            newRefs.Add(sourceRecord.Reference)
                        ElseIf Not String.IsNullOrWhiteSpace(sourceRecord.Title) Then
                            newRefs.Add(sourceRecord.Title)
                        End If
                    ElseIf Not String.IsNullOrWhiteSpace(memoryKey) Then
                        newRefs.Add(memoryKey)
                    End If

                    state.LastKnownSourceRefs =
                        newRefs.
                            Where(Function(x) Not String.IsNullOrWhiteSpace(x)).
                            Select(Function(x) x.Trim()).
                            Distinct(StringComparer.OrdinalIgnoreCase).
                            Take(5).
                            ToList()

                    Return WriteCheckpointUnlocked(state, "source_reference_created")
                End If

                state.CurrentPhase = "memory_reference_created"

                If contentKind = "tool_result" AndAlso Not String.IsNullOrWhiteSpace(memoryKey) Then
                    state.LastStructuredToolResultRef = memoryKey
                End If

                Return WriteCheckpointUnlocked(state, "memory_reference_created")
            End SyncLock
        End Function

        Public Shared Function NoteFinalStatus(workflowId As String,
                                               hostPipeline As String,
                                               isBlocked As Boolean) As Boolean
            If String.IsNullOrWhiteSpace(workflowId) Then Return False

            SyncLock _sync
                Dim state = GetOrLoadUnlocked(workflowId, hostPipeline)
                state.CurrentPhase = If(isBlocked, "final_blocked", "final_complete")
                state.UpdatedAt = DateTime.UtcNow
                Return WriteCheckpointUnlocked(state, state.CurrentPhase)
            End SyncLock
        End Function

        Public Shared Function BuildPromptContextBlock(workflowId As String,
                                                       Optional maxMemoryStubs As Integer = 4,
                                                       Optional maxSourceStubs As Integer = 4,
                                                       Optional includeRecentWorkflowMemoryStubs As Boolean = True) As String

            If String.IsNullOrWhiteSpace(workflowId) Then Return ""

            Dim state = GetState(workflowId)
            If state Is Nothing Then Return ""

            Dim currentEntries =
                SessionMemory.ListByWorkflowId(
                    workflowId,
                    maxItems:=Math.Max(8, maxMemoryStubs + maxSourceStubs + 2))

            Dim recentWorkflowEntries As List(Of SessionMemoryEntry) = New List(Of SessionMemoryEntry)()

            If includeRecentWorkflowMemoryStubs AndAlso currentEntries.Count = 0 Then
                recentWorkflowEntries =
                    SessionMemory.ListMostRecentWorkflowEntries(
                        excludedWorkflowId:=workflowId,
                        maxItems:=Math.Max(8, maxMemoryStubs + maxSourceStubs + 2))
            End If

            Dim memoryEntries = FilterPromptEntries(currentEntries, includeSourceRecords:=False, maxItems:=maxMemoryStubs)
            Dim sourceEntries = FilterPromptEntries(currentEntries, includeSourceRecords:=True, maxItems:=maxSourceStubs)

            Dim recentMemoryEntries = FilterPromptEntries(recentWorkflowEntries, includeSourceRecords:=False, maxItems:=maxMemoryStubs)
            Dim recentSourceEntries = FilterPromptEntries(recentWorkflowEntries, includeSourceRecords:=True, maxItems:=maxSourceStubs)
            Dim recentWorkflowId As String = GetEntriesWorkflowId(recentWorkflowEntries)

            Dim sb As New StringBuilder()
            sb.AppendLine("[RUNTIME_CONTEXT]")
            sb.AppendLine("Compact host-authored runtime context:")
            sb.AppendLine("- workflowId: " & state.WorkflowId)
            sb.AppendLine("- hostPipeline: " & state.HostPipeline)
            sb.AppendLine("- authoritativeRuntimeState: true")

            If Not String.IsNullOrWhiteSpace(state.CurrentPhase) Then
                sb.AppendLine("- currentPhase: " & state.CurrentPhase)
            End If

            If Not String.IsNullOrWhiteSpace(state.ActiveSkillName) Then
                sb.AppendLine("- activeSkillName: " & state.ActiveSkillName)
            End If

            If Not String.IsNullOrWhiteSpace(state.LastSuccessfulTool) Then
                sb.AppendLine("- lastSuccessfulTool: " & state.LastSuccessfulTool)
            End If

            If Not String.IsNullOrWhiteSpace(state.LastFailedTool) Then
                sb.AppendLine("- lastFailedTool: " & state.LastFailedTool)
            End If

            sb.AppendLine("- unresolvedToolFailure: " & If(state.UnresolvedToolFailure, "true", "false"))

            If Not String.IsNullOrWhiteSpace(state.LastStructuredToolResultRef) Then
                sb.AppendLine("- lastStructuredToolResultRef: " & state.LastStructuredToolResultRef)
            End If

            If Not String.IsNullOrWhiteSpace(state.LastKnownOutputReference) Then
                sb.AppendLine("- lastKnownOutputReference: " & state.LastKnownOutputReference)
            End If

            If state.LastKnownSourceRefs IsNot Nothing AndAlso state.LastKnownSourceRefs.Count > 0 Then
                sb.AppendLine("- lastKnownSourceRefs: " & String.Join(", ", state.LastKnownSourceRefs))
            End If

            If memoryEntries.Count > 0 Then
                sb.AppendLine("Memory stubs (retrieve full content explicitly with memory_get):")
                For Each entry In memoryEntries
                    sb.AppendLine("- " & BuildMemoryStub(entry))
                Next
            End If

            If sourceEntries.Count > 0 Then
                sb.AppendLine("Source stubs (retrieve full content explicitly with memory_get):")
                For Each entry In sourceEntries
                    sb.AppendLine("- " & BuildSourceStub(entry))
                Next
            End If

            If includeRecentWorkflowMemoryStubs AndAlso (recentMemoryEntries.Count > 0 OrElse recentSourceEntries.Count > 0) Then
                sb.AppendLine("Recent workflow memory stubs (retrieve full content explicitly with memory_get):")
                If Not String.IsNullOrWhiteSpace(recentWorkflowId) Then
                    sb.AppendLine("- recentWorkflowId: " & recentWorkflowId)
                End If

                For Each entry In recentMemoryEntries
                    sb.AppendLine("- " & BuildMemoryStub(entry))
                Next

                For Each entry In recentSourceEntries
                    sb.AppendLine("- " & BuildSourceStub(entry))
                Next
            End If

            sb.AppendLine("[/RUNTIME_CONTEXT]")
            Return sb.ToString().Trim()
        End Function


        Private Shared Function FilterPromptEntries(entries As IEnumerable(Of SessionMemoryEntry),
                                                    includeSourceRecords As Boolean,
                                                    maxItems As Integer) As List(Of SessionMemoryEntry)
            If entries Is Nothing Then
                Return New List(Of SessionMemoryEntry)()
            End If

            Return entries.
                Where(
                    Function(e)
                        If e Is Nothing Then Return False

                        Dim isSourceRecord As Boolean =
                            String.Equals(
                                NormalizeContentKind(If(e.Metadata?.ContentKind, "unknown")),
                                "source_record",
                                StringComparison.OrdinalIgnoreCase)

                        Return isSourceRecord = includeSourceRecords
                    End Function).
                OrderByDescending(
                    Function(e)
                        If e Is Nothing Then Return DateTime.MinValue
                        If e.UpdatedAt <> DateTime.MinValue Then Return e.UpdatedAt
                        Return e.CreatedAt
                    End Function).
                Take(Math.Max(0, maxItems)).
                ToList()
        End Function

        Private Shared Function GetEntriesWorkflowId(entries As IEnumerable(Of SessionMemoryEntry)) As String
            If entries Is Nothing Then Return ""

            For Each entry In entries
                Dim workflowId As String = If(entry?.Metadata?.WorkflowId, "").Trim()
                If workflowId <> "" Then
                    Return workflowId
                End If
            Next

            Return ""
        End Function


        Public Shared Function TryParseSourceRecord(entry As SessionMemoryEntry, ByRef record As WorkflowSourceRecord) As Boolean
            record = Nothing
            If entry Is Nothing Then Return False
            Return TryParseSourceRecord(entry.Value, record)
        End Function

        Public Shared Function TryParseSourceRecord(value As JToken, ByRef record As WorkflowSourceRecord) As Boolean
            record = Nothing

            Dim obj As JObject = TryCast(value, JObject)
            If obj Is Nothing Then Return False

            Dim sourceRecord As New WorkflowSourceRecord() With {
                .SourceId = GetScalarString(obj, "sourceId", "source_id", "id"),
                .WorkflowId = GetScalarString(obj, "workflowId", "workflow_id"),
                .Title = GetScalarString(obj, "title", "name"),
                .Provider = GetScalarString(obj, "provider"),
                .SourceType = GetScalarString(obj, "sourceType", "source_type"),
                .Reference = GetScalarString(obj, "reference", "ref", "url"),
                .Summary = GetScalarString(obj, "summary", "snippet", "shortSummary", "short_summary"),
                .RelatedTool = GetScalarString(obj, "relatedTool", "related_tool")
            }

            Dim usedInOutputValue As Boolean = False
            Dim usedInOutputToken As JToken = obj("usedInOutput")
            If usedInOutputToken IsNot Nothing Then
                If usedInOutputToken.Type = JTokenType.Boolean Then
                    usedInOutputValue = usedInOutputToken.Value(Of Boolean)()
                Else
                    Boolean.TryParse(usedInOutputToken.ToString(), usedInOutputValue)
                End If
            End If
            sourceRecord.UsedInOutput = usedInOutputValue

            Dim retrievedAtText As String = GetScalarString(obj, "retrievedAt", "retrieved_at")
            Dim parsedDate As DateTime
            If DateTime.TryParse(retrievedAtText, parsedDate) Then
                sourceRecord.RetrievedAt = parsedDate.ToUniversalTime()
            End If

            If String.IsNullOrWhiteSpace(sourceRecord.SourceId) AndAlso
               String.IsNullOrWhiteSpace(sourceRecord.Title) AndAlso
               String.IsNullOrWhiteSpace(sourceRecord.Reference) Then
                Return False
            End If

            record = sourceRecord
            Return True
        End Function

        Public Shared Function ExtractStructuredResultReference(rawContent As String) As String
            Dim match = Regex.Match(If(rawContent, ""), "\[memory:(?<key>[^\]]+)\]", RegexOptions.IgnoreCase)
            If match.Success Then
                Return match.Groups("key").Value.Trim()
            End If

            Dim token = TryParseJson(rawContent)
            If token Is Nothing Then Return ""

            Dim candidates As String() = {
                "memory_key",
                "memoryKey",
                "result.memory_key",
                "result.memoryKey",
                "output_reference",
                "result.output_reference",
                "reference",
                "result.reference"
            }

            For Each candidate In candidates
                Dim value = SelectScalar(token, candidate)
                If Not String.IsNullOrWhiteSpace(value) Then
                    Return value
                End If
            Next

            Return ""
        End Function

        Public Shared Function ExtractOutputReference(rawContent As String) As String
            Dim token = TryParseJson(rawContent)
            If token Is Nothing Then Return ""

            Dim candidates As String() = {
                "path",
                "saved_path",
                "output_reference",
                "reference",
                "result.path",
                "result.saved_path",
                "result.output_reference",
                "results[0].path",
                "results[0].reference"
            }

            For Each candidate In candidates
                Dim value = SelectScalar(token, candidate)
                If Not String.IsNullOrWhiteSpace(value) Then
                    Return value
                End If
            Next

            Return ""
        End Function

        Public Shared Function ExtractSourceReferences(rawContent As String) As List(Of String)
            Dim token = TryParseJson(rawContent)
            Dim results As New List(Of String)()

            If token Is Nothing Then
                Return results
            End If

            CollectSourceReferences(token, results)

            Return results.
                Where(Function(x) Not String.IsNullOrWhiteSpace(x)).
                Select(Function(x) x.Trim()).
                Distinct(StringComparer.OrdinalIgnoreCase).
                Take(5).
                ToList()
        End Function

#If DEBUG Then
        Friend Shared Sub ResetForTests()
            SyncLock _sync
                _states.Clear()
            End SyncLock
        End Sub
#End If

        Private Shared Function GetOrLoadUnlocked(workflowId As String, hostPipeline As String) As WorkflowRuntimeState
            Dim state As WorkflowRuntimeState = Nothing

            If _states.TryGetValue(workflowId, state) Then
                Return state
            End If

            state = LoadCheckpointUnlocked(workflowId)

            If state Is Nothing Then
                Dim nowUtc = DateTime.UtcNow
                state = New WorkflowRuntimeState() With {
                    .WorkflowId = workflowId,
                    .HostPipeline = If(hostPipeline, ""),
                    .CreatedAt = nowUtc,
                    .UpdatedAt = nowUtc,
                    .Authoritative = True
                }
            ElseIf Not String.IsNullOrWhiteSpace(hostPipeline) Then
                state.HostPipeline = hostPipeline
            End If

            _states(workflowId) = state
            Return state
        End Function

        Private Shared Function LoadCheckpointUnlocked(workflowId As String) As WorkflowRuntimeState
            Try
                Dim path As String = GetCheckpointPath(workflowId)
                If Not File.Exists(path) Then Return Nothing

                Dim raw = File.ReadAllText(path, Encoding.UTF8)
                Dim envelope = JsonConvert.DeserializeObject(Of WorkflowCheckpointEnvelope)(raw)

                If envelope Is Nothing OrElse envelope.RuntimeState Is Nothing Then
                    Return Nothing
                End If

                Return envelope.RuntimeState
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Function WriteCheckpointUnlocked(state As WorkflowRuntimeState, checkpointKind As String) As Boolean
            Try
                Dim dir = GetPrivateWorkflowDirectory()
                If Not Directory.Exists(dir) Then Directory.CreateDirectory(dir)

                Dim envelope As New WorkflowCheckpointEnvelope() With {
                    .WorkflowId = state.WorkflowId,
                    .HostPipeline = state.HostPipeline,
                    .CheckpointKind = checkpointKind,
                    .WrittenAt = DateTime.UtcNow,
                    .RuntimeState = CloneState(state)
                }

                File.WriteAllText(
                    GetCheckpointPath(state.WorkflowId),
                    JsonConvert.SerializeObject(envelope, Formatting.None),
                    Encoding.UTF8)

                Debug.WriteLine(
                    ComposeWorkflowLogMessage(
                        "Checkpoint written.",
                        state.WorkflowId,
                        checkpointKind,
                        hostName:=state.HostPipeline) &
                    " [checkpointWritten: true]")
                Return True
            Catch ex As Exception
                Debug.WriteLine(
                    ComposeWorkflowLogMessage(
                        "Checkpoint write failed.",
                        state.WorkflowId,
                        checkpointKind,
                        hostName:=state.HostPipeline) &
                    " [checkpointWritten: false] [error: " & ex.Message & "]")
                Return False
            End Try
        End Function

        Private Shared Function CloneState(state As WorkflowRuntimeState) As WorkflowRuntimeState
            If state Is Nothing Then Return Nothing

            Return New WorkflowRuntimeState() With {
                .WorkflowId = state.WorkflowId,
                .HostPipeline = state.HostPipeline,
                .ActiveSkillName = state.ActiveSkillName,
                .CurrentPhase = state.CurrentPhase,
                .LastSuccessfulTool = state.LastSuccessfulTool,
                .LastFailedTool = state.LastFailedTool,
                .UnresolvedToolFailure = state.UnresolvedToolFailure,
                .LastStructuredToolResultRef = state.LastStructuredToolResultRef,
                .LastKnownOutputReference = state.LastKnownOutputReference,
                .LastKnownSourceRefs = New List(Of String)(If(state.LastKnownSourceRefs, New List(Of String)())),
                .ToolCallSuccessCount = state.ToolCallSuccessCount,
                .ToolCallFailureCount = state.ToolCallFailureCount,
                .RetryCount = state.RetryCount,
                .CreatedAt = state.CreatedAt,
                .UpdatedAt = state.UpdatedAt,
                .Authoritative = state.Authoritative
            }
        End Function

        Private Shared Function GetPrivateWorkflowDirectory() As String
            Dim root As String = TryGetSharedPath("INI_AgentResourcesPathLocal")

            If String.IsNullOrWhiteSpace(root) Then
                root = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "RedInk")
            End If

            Return Path.Combine(root, ".session", "workflow_runtime")
        End Function

        Public Shared Function GetCheckpointPath(workflowId As String) As String
            Dim safeWorkflowId As String = Regex.Replace(If(workflowId, "").Trim(), "[^A-Za-z0-9_\-]", "_")
            If safeWorkflowId = "" Then safeWorkflowId = "workflow"
            Return Path.Combine(GetPrivateWorkflowDirectory(), safeWorkflowId & ".json")
        End Function

        Private Shared Function BuildMemoryStub(entry As SessionMemoryEntry) As String
            Dim metadata As SessionMemoryMetadata = EnsureMetadataDefaults(entry.Metadata)
            Dim parts As New List(Of String) From {
                SessionMemory.BuildStub(entry),
                "contentKind=" & metadata.ContentKind,
                "source=" & metadata.Source
            }

            If Not String.IsNullOrWhiteSpace(metadata.WorkflowId) Then
                parts.Add("workflowId=" & metadata.WorkflowId)
            End If

            If Not String.IsNullOrWhiteSpace(metadata.RelatedTool) Then
                parts.Add("relatedTool=" & metadata.RelatedTool)
            End If

            If Not String.IsNullOrWhiteSpace(metadata.RelatedAgent) Then
                parts.Add("relatedAgent=" & metadata.RelatedAgent)
            End If

            If Not String.IsNullOrWhiteSpace(metadata.RelatedSkill) Then
                parts.Add("relatedSkill=" & metadata.RelatedSkill)
            End If

            parts.Add("trust=" & If(metadata.TrustedForRuntime, "authoritative", "advisory"))

            Return String.Join(" | ", parts)
        End Function

        Private Shared Function BuildSourceStub(entry As SessionMemoryEntry) As String
            Dim metadata As SessionMemoryMetadata = EnsureMetadataDefaults(entry.Metadata)
            Dim record As WorkflowSourceRecord = Nothing

            If TryParseSourceRecord(entry, record) Then
                Dim typeLabel As String = If(
                    Not String.IsNullOrWhiteSpace(record.Provider),
                    record.Provider,
                    record.SourceType)

                Dim summaryText As String = If(record.Summary, "").Trim()
                If summaryText.Length > 120 Then
                    summaryText = summaryText.Substring(0, 117) & "..."
                End If

                Dim retrievedText As String = ""
                If record.RetrievedAt <> DateTime.MinValue Then
                    retrievedText = record.RetrievedAt.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
                End If

                Dim parts As New List(Of String) From {
                    SessionMemory.BuildStub(entry),
                    "title=" & If(record.Title, ""),
                    "contentKind=" & metadata.ContentKind,
                    "source=" & metadata.Source
                }

                If Not String.IsNullOrWhiteSpace(metadata.WorkflowId) Then
                    parts.Add("workflowId=" & metadata.WorkflowId)
                End If

                If Not String.IsNullOrWhiteSpace(metadata.RelatedTool) Then
                    parts.Add("relatedTool=" & metadata.RelatedTool)
                End If

                If Not String.IsNullOrWhiteSpace(metadata.RelatedAgent) Then
                    parts.Add("relatedAgent=" & metadata.RelatedAgent)
                End If

                If Not String.IsNullOrWhiteSpace(typeLabel) Then
                    parts.Add("type=" & typeLabel)
                End If

                If retrievedText <> "" Then
                    parts.Add("retrievedAt=" & retrievedText)
                End If

                If summaryText <> "" Then
                    parts.Add("summary=" & summaryText)
                End If

                Return String.Join(" | ", parts)
            End If

            Return String.Join(
                " | ",
                New String() {
                    SessionMemory.BuildStub(entry),
                    "contentKind=" & metadata.ContentKind,
                    "source=" & metadata.Source,
                    If(String.IsNullOrWhiteSpace(metadata.WorkflowId), "", "workflowId=" & metadata.WorkflowId)
                }.Where(Function(part) Not String.IsNullOrWhiteSpace(part)))
        End Function

        Private Shared Function TryParseJson(rawContent As String) As JToken
            Try
                If String.IsNullOrWhiteSpace(rawContent) Then Return Nothing
                Return JToken.Parse(rawContent)
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Function SelectScalar(root As JToken, path As String) As String
            If root Is Nothing OrElse String.IsNullOrWhiteSpace(path) Then Return ""

            Try
                Dim token = root.SelectToken(path, errorWhenNoMatch:=False)
                If token Is Nothing OrElse token.Type = JTokenType.Null Then Return ""

                If token.Type = JTokenType.String OrElse
                   token.Type = JTokenType.Integer OrElse
                   token.Type = JTokenType.Float OrElse
                   token.Type = JTokenType.Boolean Then
                    Return token.ToString().Trim()
                End If
            Catch
            End Try

            Return ""
        End Function

        Private Shared Sub CollectSourceReferences(token As JToken, results As List(Of String))
            If token Is Nothing OrElse results Is Nothing Then Return

            Dim obj As JObject = TryCast(token, JObject)
            If obj IsNot Nothing Then
                For Each propertyName In New String() {"sourceId", "source_id", "reference", "ref"}
                    Dim value = GetScalarString(obj, propertyName)
                    If Not String.IsNullOrWhiteSpace(value) Then
                        results.Add(value)
                    End If
                Next

                For Each prop In obj.Properties()
                    CollectSourceReferences(prop.Value, results)
                Next

                Return
            End If

            Dim arr As JArray = TryCast(token, JArray)
            If arr IsNot Nothing Then
                For Each item In arr
                    CollectSourceReferences(item, results)
                Next
            End If
        End Sub

        Private Shared Function GetScalarString(obj As JObject, ParamArray names() As String) As String
            If obj Is Nothing OrElse names Is Nothing Then Return ""

            For Each name In names
                Dim token As JToken = obj(name)
                If token Is Nothing OrElse token.Type = JTokenType.Null Then Continue For

                If token.Type = JTokenType.String OrElse
                   token.Type = JTokenType.Integer OrElse
                   token.Type = JTokenType.Float OrElse
                   token.Type = JTokenType.Boolean Then
                    Dim value As String = token.ToString().Trim()
                    If value <> "" Then
                        Return value
                    End If
                End If
            Next

            Return ""
        End Function

        Private Shared Function TryGetSharedPath(propertyName As String) As String
            Try
                Dim asm = GetType(SharedLibrary.SharedContext).Assembly

                For Each typeFullName In {"SharedLibrary.SharedProperties", "SharedLibrary.SharedContext"}
                    Dim t = asm.GetType(typeFullName, throwOnError:=False, ignoreCase:=False)
                    If t Is Nothing Then Continue For

                    Dim pi = t.GetProperty(
                        propertyName,
                        Reflection.BindingFlags.Public Or Reflection.BindingFlags.Static Or Reflection.BindingFlags.Instance)

                    If pi Is Nothing Then Continue For

                    Dim getter = pi.GetGetMethod()
                    If getter Is Nothing OrElse Not getter.IsStatic Then Continue For

                    Dim value As Object = pi.GetValue(Nothing, Nothing)
                    If TypeOf value Is String AndAlso Not String.IsNullOrWhiteSpace(CStr(value)) Then
                        Return CStr(value)
                    End If
                Next
            Catch
            End Try

            Return Nothing
        End Function

    End Class

End Namespace