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
        Public Property ContinuationState As ToolCallSequencing.ToolingRunContinuationSnapshot
    End Class

    Public NotInheritable Class WorkflowContinuity

        Private Sub New()
        End Sub

        Public Const DefaultContinuationRetentionDays As System.Int32 = 7

        Private Shared ReadOnly _sync As New Object()
        Private Shared ReadOnly _states As New Dictionary(Of String, WorkflowRuntimeState)(StringComparer.OrdinalIgnoreCase)
        Private Shared ReadOnly _continuations As New Dictionary(Of String, ToolCallSequencing.ToolingRunContinuationSnapshot)(StringComparer.OrdinalIgnoreCase)
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

        Public Shared Function ResumeWorkflow(workflowId As String,
                                              hostPipeline As String,
                                              Optional continuationRetentionDays As System.Int32 = 0) As WorkflowRuntimeState
            If String.IsNullOrWhiteSpace(workflowId) Then
                Throw New System.ArgumentException("A workflow id is required for continuation.", NameOf(workflowId))
            End If

            SyncLock _sync
                Dim state As WorkflowRuntimeState = GetOrLoadUnlocked(workflowId, System.String.Empty)
                Dim requestedHost As System.String = If(hostPipeline, System.String.Empty).Trim()
                If requestedHost <> System.String.Empty AndAlso
                   state IsNot Nothing AndAlso
                   Not System.String.IsNullOrWhiteSpace(state.HostPipeline) AndAlso
                   Not System.String.Equals(state.HostPipeline, requestedHost, System.StringComparison.OrdinalIgnoreCase) Then
                    Throw New System.InvalidOperationException("The requested workflow continuation belongs to a different host pipeline.")
                End If
                Dim snapshot As ToolCallSequencing.ToolingRunContinuationSnapshot = Nothing
                If Not _continuations.TryGetValue(workflowId, snapshot) OrElse snapshot Is Nothing Then
                    Throw New System.InvalidOperationException("The requested workflow has no resumable continuation checkpoint.")
                End If

                If IsContinuationExpired(snapshot) Then
                    _continuations.Remove(workflowId)
                    DeleteContinuationArtifactsUnlocked(workflowId)
                    Throw New System.InvalidOperationException("The requested workflow continuation checkpoint has expired.")
                End If

                snapshot.ResumeCount += 1
                RefreshContinuationExpiryUnlocked(
                    snapshot,
                    If(continuationRetentionDays > 0, continuationRetentionDays, snapshot.RetentionDays))
                state.CurrentPhase = "workflow_resumed"
                state.UpdatedAt = DateTime.UtcNow
                If requestedHost <> System.String.Empty Then state.HostPipeline = requestedHost
                state.Authoritative = True
                WriteCheckpointUnlocked(state, "workflow_resumed")
                Return CloneState(state)
            End SyncLock
        End Function

        Public Shared Function TryGetContinuationSnapshot(
            workflowId As String,
            ByRef snapshot As ToolCallSequencing.ToolingRunContinuationSnapshot) As Boolean

            snapshot = Nothing
            If String.IsNullOrWhiteSpace(workflowId) Then Return False

            SyncLock _sync
                GetOrLoadUnlocked(workflowId, "")
                Dim stored As ToolCallSequencing.ToolingRunContinuationSnapshot = Nothing
                If Not _continuations.TryGetValue(workflowId, stored) OrElse stored Is Nothing Then Return False
                If IsContinuationExpired(stored) Then
                    _continuations.Remove(workflowId)
                    DeleteContinuationArtifactsUnlocked(workflowId)
                    Return False
                End If
                snapshot = CloneContinuationSnapshot(stored)
                Return snapshot IsNot Nothing
            End SyncLock
        End Function

        Public Shared Function HasContinuationSnapshot(workflowId As String) As Boolean
            Dim snapshot As ToolCallSequencing.ToolingRunContinuationSnapshot = Nothing
            Return TryGetContinuationSnapshot(workflowId, snapshot)
        End Function

        ''' <summary>
        ''' Explicitly abandons one resumable continuation without deleting the workflow's
        ''' diagnostic runtime history. The continuation snapshot and cached artifacts are
        ''' removed and the checkpoint is rewritten without resumable state.
        ''' </summary>
        Public Shared Function InvalidateContinuation(
            workflowId As System.String,
            Optional expectedContinuationKey As System.String = "",
            Optional expectedHostPipeline As System.String = "") As System.Boolean

            Dim id As System.String = If(workflowId, System.String.Empty).Trim()
            If id = System.String.Empty Then Return False

            SyncLock _sync
                Dim checkpointPath As System.String = GetCheckpointPath(id)
                If Not _continuations.ContainsKey(id) AndAlso Not System.IO.File.Exists(checkpointPath) Then Return False

                Dim state As WorkflowRuntimeState = GetOrLoadUnlocked(id, System.String.Empty)
                Dim snapshot As ToolCallSequencing.ToolingRunContinuationSnapshot = Nothing
                If Not _continuations.TryGetValue(id, snapshot) OrElse snapshot Is Nothing Then Return False

                Dim expectedKey As System.String = If(expectedContinuationKey, System.String.Empty).Trim()
                If expectedKey <> System.String.Empty AndAlso
                   Not System.String.Equals(If(snapshot.ContinuationKey, System.String.Empty).Trim(), expectedKey, System.StringComparison.Ordinal) Then
                    Return False
                End If

                Dim expectedHost As System.String = If(expectedHostPipeline, System.String.Empty).Trim()
                If expectedHost <> System.String.Empty AndAlso
                   (state Is Nothing OrElse
                    Not System.String.Equals(If(state.HostPipeline, System.String.Empty), expectedHost, System.StringComparison.OrdinalIgnoreCase)) Then
                    Return False
                End If

                _continuations.Remove(id)
                DeleteContinuationArtifactsUnlocked(id)

                If state Is Nothing Then Return True
                state.UpdatedAt = System.DateTime.UtcNow
                Return WriteCheckpointUnlocked(state, "continuation_invalidated")
            End SyncLock
        End Function

        Public Shared Function FindBlockedContinuationWorkflowId(
            continuationKey As String,
            Optional hostPipeline As String = "") As String

            Dim key As String = If(continuationKey, "").Trim()
            If key = "" Then Return ""

            SyncLock _sync
                Try
                    Dim dir As String = GetPrivateWorkflowDirectory()
                    If Not Directory.Exists(dir) Then Return ""
                    PruneStaleContinuationArtifactsUnlocked(dir)

                    Dim bestWorkflowId As String = ""
                    Dim bestWrittenAt As DateTime = DateTime.MinValue

                    For Each checkpointPath As String In Directory.GetFiles(dir, "*.json", SearchOption.TopDirectoryOnly)
                        Dim envelope As WorkflowCheckpointEnvelope = LoadCheckpointEnvelopeFromPathUnlocked(checkpointPath)
                        If envelope Is Nothing OrElse envelope.RuntimeState Is Nothing OrElse envelope.ContinuationState Is Nothing Then Continue For
                        If Not String.Equals(envelope.RuntimeState.CurrentPhase, "final_blocked", StringComparison.OrdinalIgnoreCase) Then Continue For
                        If IsContinuationExpired(envelope.ContinuationState) Then Continue For
                        If Not String.Equals(If(envelope.ContinuationState.ContinuationKey, "").Trim(), key, StringComparison.Ordinal) Then Continue For
                        If Not String.IsNullOrWhiteSpace(hostPipeline) AndAlso
                           Not String.Equals(If(envelope.RuntimeState.HostPipeline, ""), hostPipeline, StringComparison.OrdinalIgnoreCase) Then Continue For

                        If envelope.WrittenAt >= bestWrittenAt Then
                            bestWrittenAt = envelope.WrittenAt
                            bestWorkflowId = If(envelope.WorkflowId, "").Trim()
                        End If
                    Next

                    Return bestWorkflowId
                Catch
                    Return ""
                End Try
            End SyncLock
        End Function

        Public Shared Function GetLatestBlockedContinuationUserLanguage(
            continuationKey As String,
            Optional hostPipeline As String = "") As String

            Dim key As String = If(continuationKey, "").Trim()
            If key = "" Then Return ""

            SyncLock _sync
                Try
                    Dim dir As String = GetPrivateWorkflowDirectory()
                    If Not Directory.Exists(dir) Then Return ""

                    Dim bestLanguage As String = ""
                    Dim bestWrittenAt As DateTime = DateTime.MinValue

                    For Each checkpointPath As String In Directory.GetFiles(dir, "*.json", SearchOption.TopDirectoryOnly)
                        Dim envelope As WorkflowCheckpointEnvelope = LoadCheckpointEnvelopeFromPathUnlocked(checkpointPath)
                        If envelope Is Nothing OrElse envelope.RuntimeState Is Nothing OrElse envelope.ContinuationState Is Nothing Then Continue For
                        If Not String.Equals(envelope.RuntimeState.CurrentPhase, "final_blocked", StringComparison.OrdinalIgnoreCase) Then Continue For
                        If Not String.Equals(If(envelope.ContinuationState.ContinuationKey, "").Trim(), key, StringComparison.Ordinal) Then Continue For
                        If Not String.IsNullOrWhiteSpace(hostPipeline) AndAlso
                           Not String.Equals(If(envelope.RuntimeState.HostPipeline, ""), hostPipeline, StringComparison.OrdinalIgnoreCase) Then Continue For

                        If envelope.WrittenAt >= bestWrittenAt Then
                            bestWrittenAt = envelope.WrittenAt
                            bestLanguage = ExtractContinuationUserLanguage(envelope.ContinuationState)
                        End If
                    Next

                    Return bestLanguage
                Catch
                    Return ""
                End Try
            End SyncLock
        End Function

        Public Shared Function BuildContinuationRetryHint(userLanguage As String) As String
            Select Case NormalizeUserLanguageKey(userLanguage)
                Case "de"
                    Return "Wenn Sie diesen technisch blockierten Vorgang fortsetzen möchten, ohne bereits erfolgreich erledigte Arbeit absichtlich neu zu erzeugen, antworten Sie auf diese E-Mail mit `retry` als erster Zeile. Darunter können Sie weitere Anweisungen ergänzen. Bereits verifizierte Ergebnisse werden nach Möglichkeit wiederverwendet."
                Case "fr"
                    Return "Si vous souhaitez poursuivre cette exécution techniquement bloquée sans recréer délibérément le travail déjà effectué avec succès, répondez à cet e-mail avec `retry` sur la première ligne. Vous pouvez ajouter d’autres instructions en dessous. Les résultats déjà vérifiés seront réutilisés dans la mesure du possible."
                Case "it"
                    Return "Se desideri continuare questa esecuzione tecnicamente bloccata senza ricreare deliberatamente il lavoro già completato con successo, rispondi a questa e-mail inserendo `retry` nella prima riga. Puoi aggiungere ulteriori istruzioni nelle righe successive. I risultati già verificati verranno riutilizzati ove possibile."
                Case Else
                    Return "If you want me to continue this technically blocked run without deliberately recreating work that already succeeded, reply to this email with `retry` as the first line. You may add further instructions below it. Existing verified results will be reused where possible."
            End Select
        End Function

        Public Shared Function BuildContinuationUnavailableNotice(userLanguage As String) As String
            Select Case NormalizeUserLanguageKey(userLanguage)
                Case "de"
                    Return "Für diese Unterhaltung konnte kein fortsetzbarer blockierter Vorgang mehr gefunden werden. Ich habe keinen neuen Lauf gestartet, weil dadurch Dateien oder andere bereits erfolgreich ausgeführte Aktionen doppelt erzeugt werden könnten. Wenn Sie die Aufgabe neu starten möchten, senden oder formulieren Sie sie bitte erneut als neue Anfrage."
                Case "fr"
                    Return "Je n’ai plus trouvé d’exécution bloquée pouvant être reprise pour cette conversation. Je n’ai pas lancé une nouvelle exécution, car cela pourrait recréer des fichiers ou répéter d’autres actions déjà effectuées avec succès. Si vous souhaitez recommencer la tâche, veuillez la renvoyer ou la reformuler comme une nouvelle demande."
                Case "it"
                    Return "Non è stato possibile trovare per questa conversazione un’esecuzione bloccata ancora riprendibile. Non ho avviato una nuova esecuzione, perché ciò potrebbe ricreare file o ripetere altre azioni già completate con successo. Se desideri ricominciare l’attività, inviala o formulala nuovamente come nuova richiesta."
                Case Else
                    Return "I could not find a resumable blocked workflow checkpoint for this conversation. I have not started a fresh run, because doing so could repeat file creation or other actions that may already have succeeded. Please resend or restate the task if you want to start it again as a new request."
            End Select
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

        Public Shared Function NoteFinalStatus(
            workflowId As String,
            hostPipeline As String,
            isBlocked As Boolean,
            Optional sequencingState As ToolCallSequencing.ToolingRunState = Nothing,
            Optional originalUserRequestRaw As String = "",
            Optional lastUserFacingResponse As String = "",
            Optional continuationKey As String = "",
            Optional continuationRetentionDays As System.Int32 = DefaultContinuationRetentionDays) As Boolean

            If String.IsNullOrWhiteSpace(workflowId) Then Return False

            SyncLock _sync
                Dim state = GetOrLoadUnlocked(workflowId, hostPipeline)
                state.CurrentPhase = If(isBlocked, "final_blocked", "final_complete")
                state.UpdatedAt = DateTime.UtcNow
                PruneStaleContinuationArtifactsUnlocked(GetPrivateWorkflowDirectory())

                If isBlocked AndAlso sequencingState IsNot Nothing Then
                    Dim prior As ToolCallSequencing.ToolingRunContinuationSnapshot = Nothing
                    _continuations.TryGetValue(workflowId, prior)

                    Dim originalRequest As String = BoundedContinuationText(originalUserRequestRaw, 16000)
                    If prior IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(prior.OriginalUserRequestRaw) Then
                        originalRequest = prior.OriginalUserRequestRaw
                    End If

                    Dim effectiveKey As String = If(continuationKey, "").Trim()
                    If effectiveKey = "" AndAlso prior IsNot Nothing Then effectiveKey = If(prior.ContinuationKey, "").Trim()

                    Dim snapshot As ToolCallSequencing.ToolingRunContinuationSnapshot =
                        sequencingState.CreateContinuationSnapshot(
                            originalRequest,
                            BoundedContinuationText(lastUserFacingResponse, 16000),
                            effectiveKey)

                    If prior IsNot Nothing Then
                        snapshot.CreatedUtc = prior.CreatedUtc
                        snapshot.ResumeCount = prior.ResumeCount
                    End If

                    RefreshContinuationExpiryUnlocked(snapshot, continuationRetentionDays)
                    CacheContinuationArtifactsUnlocked(workflowId, snapshot)
                    _continuations(workflowId) = snapshot
                ElseIf Not isBlocked Then
                    ' Do not delete the durable artifact cache here. Host delivery happens
                    ' after finalization and may still reference these restored paths. Stale
                    ' continuation caches are pruned separately by age.
                    _continuations.Remove(workflowId)
                End If

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

            Dim continuation As ToolCallSequencing.ToolingRunContinuationSnapshot = Nothing
            If TryGetContinuationSnapshot(workflowId, continuation) AndAlso continuation IsNot Nothing AndAlso continuation.ResumeCount > 0 Then
                sb.AppendLine("[CROSS_RUN_CONTINUATION]")
                sb.AppendLine("This is an explicit continuation of a previously blocked workflow.")
                sb.AppendLine("Reuse successful operations and registered artifacts from the restored host state. Do not repeat a successful mutating operation merely because this is a new mail/run. Revalidate existing outputs first and retry only unresolved or missing work.")
                If Not String.IsNullOrWhiteSpace(continuation.OriginalUserRequestRaw) Then
                    sb.AppendLine("Original request from the blocked workflow:")
                    sb.AppendLine(BoundedContinuationText(continuation.OriginalUserRequestRaw, 12000))
                End If
                Dim continuationDetails As String = BuildContinuationPromptDetails(continuation)
                If continuationDetails <> "" Then sb.AppendLine(continuationDetails)
                If Not String.IsNullOrWhiteSpace(continuation.LastUserFacingResponse) Then
                    sb.AppendLine("Prior provisional user-facing response:")
                    sb.AppendLine(BoundedContinuationText(continuation.LastUserFacingResponse, 6000))
                End If
                sb.AppendLine("[/CROSS_RUN_CONTINUATION]")
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

        Private Shared Function GetOrLoadUnlocked(workflowId As String, hostPipeline As String) As WorkflowRuntimeState
            Dim state As WorkflowRuntimeState = Nothing

            If _states.TryGetValue(workflowId, state) Then
                Return state
            End If

            Dim envelope As WorkflowCheckpointEnvelope = LoadCheckpointEnvelopeUnlocked(workflowId)
            state = If(envelope Is Nothing, Nothing, envelope.RuntimeState)
            If envelope IsNot Nothing AndAlso envelope.ContinuationState IsNot Nothing Then
                If Not IsContinuationExpired(envelope.ContinuationState) Then
                    _continuations(workflowId) = envelope.ContinuationState
                Else
                    DeleteContinuationArtifactsUnlocked(workflowId)
                End If
            End If

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

        Private Shared Function LoadCheckpointEnvelopeUnlocked(workflowId As String) As WorkflowCheckpointEnvelope
            Return LoadCheckpointEnvelopeFromPathUnlocked(GetCheckpointPath(workflowId))
        End Function

        Private Shared Function LoadCheckpointEnvelopeFromPathUnlocked(path As String) As WorkflowCheckpointEnvelope
            Try
                If String.IsNullOrWhiteSpace(path) OrElse Not File.Exists(path) Then Return Nothing
                Dim raw As String = File.ReadAllText(path, Encoding.UTF8)
                Return JsonConvert.DeserializeObject(Of WorkflowCheckpointEnvelope)(raw)
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Function WriteCheckpointUnlocked(state As WorkflowRuntimeState, checkpointKind As String) As Boolean
            Try
                Dim dir = GetPrivateWorkflowDirectory()
                If Not Directory.Exists(dir) Then Directory.CreateDirectory(dir)

                Dim continuation As ToolCallSequencing.ToolingRunContinuationSnapshot = Nothing
                _continuations.TryGetValue(state.WorkflowId, continuation)

                Dim envelope As New WorkflowCheckpointEnvelope() With {
                    .WorkflowId = state.WorkflowId,
                    .HostPipeline = state.HostPipeline,
                    .CheckpointKind = checkpointKind,
                    .WrittenAt = DateTime.UtcNow,
                    .RuntimeState = CloneState(state),
                    .ContinuationState = CloneContinuationSnapshot(continuation)
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

        Private Shared Function BuildContinuationPromptDetails(
            snapshot As ToolCallSequencing.ToolingRunContinuationSnapshot) As String

            If snapshot Is Nothing OrElse String.IsNullOrWhiteSpace(snapshot.SequencingStateJson) Then Return ""
            Try
                Dim root As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(snapshot.SequencingStateJson)
                Dim sb As New StringBuilder()

                Dim successfulTools As Newtonsoft.Json.Linq.JArray = TryCast(root("SuccessfulToolsThisRun"), Newtonsoft.Json.Linq.JArray)
                If successfulTools IsNot Nothing AndAlso successfulTools.Count > 0 Then
                    Dim names As List(Of String) = successfulTools.
                        Values(Of String)().
                        Where(Function(x) Not String.IsNullOrWhiteSpace(x)).
                        Distinct(StringComparer.OrdinalIgnoreCase).
                        Take(12).
                        ToList()
                    If names.Count > 0 Then sb.AppendLine("Previously successful tools: " & String.Join(", ", names))
                End If

                Dim artifacts As Newtonsoft.Json.Linq.JArray = TryCast(root("RegisteredDeliverableArtifacts"), Newtonsoft.Json.Linq.JArray)
                If artifacts IsNot Nothing AndAlso artifacts.Count > 0 Then
                    sb.AppendLine("Restored registered artifacts:")
                    For Each token As Newtonsoft.Json.Linq.JToken In artifacts.Take(12)
                        Dim artifact As Newtonsoft.Json.Linq.JObject = TryCast(token, Newtonsoft.Json.Linq.JObject)
                        If artifact Is Nothing Then Continue For
                        Dim sessionPath As String = If(artifact.Value(Of String)("SessionPath"), "")
                        Dim logicalId As String = If(artifact.Value(Of String)("LogicalDeliverableId"), "")
                        Dim slotId As String = If(artifact.Value(Of String)("OutputSlotId"), "")
                        Dim stateName As String = If(artifact.Value(Of String)("LifecycleState"), "")
                        sb.AppendLine("- file=" & If(String.IsNullOrWhiteSpace(sessionPath), "(unavailable)", System.IO.Path.GetFileName(sessionPath)) &
                                      "; logicalDeliverableId=" & logicalId &
                                      "; outputSlotId=" & slotId &
                                      "; lifecycle=" & stateName &
                                      "; exists=" & If(Not String.IsNullOrWhiteSpace(sessionPath) AndAlso System.IO.File.Exists(sessionPath), "true", "false"))
                    Next
                End If

                Dim lastResult As String = If(root.Value(Of String)("LastStructuredToolResult"), "")
                If Not String.IsNullOrWhiteSpace(lastResult) Then
                    sb.AppendLine("Last structured tool result excerpt:")
                    sb.AppendLine(BoundedContinuationText(lastResult, 4000))
                End If

                Return sb.ToString().TrimEnd()
            Catch
                Return ""
            End Try
        End Function

        Private Shared Function CloneContinuationSnapshot(
            snapshot As ToolCallSequencing.ToolingRunContinuationSnapshot) As ToolCallSequencing.ToolingRunContinuationSnapshot

            If snapshot Is Nothing Then Return Nothing
            Try
                Dim json As String = JsonConvert.SerializeObject(snapshot, Formatting.None)
                Return JsonConvert.DeserializeObject(Of ToolCallSequencing.ToolingRunContinuationSnapshot)(json)
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Function BoundedContinuationText(value As String, maxChars As Integer) As String
            Dim text As String = If(value, "")
            If maxChars <= 0 OrElse text.Length <= maxChars Then Return text
            Return text.Substring(0, maxChars) & Environment.NewLine & "[... continuation text truncated ...]"
        End Function

        Private Shared Sub CacheContinuationArtifactsUnlocked(
            workflowId As String,
            snapshot As ToolCallSequencing.ToolingRunContinuationSnapshot)

            If snapshot Is Nothing OrElse String.IsNullOrWhiteSpace(snapshot.SequencingStateJson) Then Return

            Try
                Dim root As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(snapshot.SequencingStateJson)
                Dim artifacts As Newtonsoft.Json.Linq.JArray = TryCast(root("RegisteredDeliverableArtifacts"), Newtonsoft.Json.Linq.JArray)
                If artifacts Is Nothing OrElse artifacts.Count = 0 Then Return

                Dim cacheRoot As String = GetContinuationArtifactDirectory(workflowId)
                Directory.CreateDirectory(cacheRoot)
                Dim totalBytes As Long = 0
                Const MaxCachedBytes As Long = 100L * 1024L * 1024L

                Dim index As Integer = 0
                For Each token As Newtonsoft.Json.Linq.JToken In artifacts
                    Dim artifact As Newtonsoft.Json.Linq.JObject = TryCast(token, Newtonsoft.Json.Linq.JObject)
                    If artifact Is Nothing Then Continue For
                    Dim sourcePath As String = If(artifact.Value(Of String)("SessionPath"), "").Trim()
                    If sourcePath = "" OrElse Not File.Exists(sourcePath) Then Continue For

                    Dim info As New FileInfo(sourcePath)
                    If info.Length < 0 OrElse totalBytes + info.Length > MaxCachedBytes Then Continue For
                    totalBytes += info.Length

                    index += 1
                    Dim artifactId As String = Regex.Replace(If(artifact.Value(Of String)("ArtifactId"), "").Trim(), "[^A-Za-z0-9_\-]", "_")
                    If artifactId = "" Then artifactId = "artifact_" & index.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    Dim artifactDir As String = Path.Combine(cacheRoot, artifactId)
                    Directory.CreateDirectory(artifactDir)
                    Dim targetPath As String = Path.Combine(artifactDir, Path.GetFileName(sourcePath))
                    If Not String.Equals(Path.GetFullPath(sourcePath), Path.GetFullPath(targetPath), StringComparison.OrdinalIgnoreCase) Then
                        File.Copy(sourcePath, targetPath, overwrite:=True)
                    End If
                    artifact("SessionPath") = targetPath
                Next

                snapshot.SequencingStateJson = root.ToString(Formatting.None)
            Catch
                ' Continuation caching is best effort. The snapshot still preserves operation
                ' state and will fail closed if a required physical artifact is unavailable.
            End Try
        End Sub

        Private Shared Function GetContinuationArtifactDirectory(workflowId As String) As String
            Dim safeWorkflowId As String = Regex.Replace(If(workflowId, "").Trim(), "[^A-Za-z0-9_\-]", "_")
            If safeWorkflowId = "" Then safeWorkflowId = "workflow"
            Return Path.Combine(GetPrivateWorkflowDirectory(), safeWorkflowId & "_artifacts")
        End Function

        Private Shared Function NormalizeContinuationRetentionDays(value As System.Int32) As System.Int32
            If value > 0 Then Return value
            Return DefaultContinuationRetentionDays
        End Function

        Private Shared Sub RefreshContinuationExpiryUnlocked(
            snapshot As ToolCallSequencing.ToolingRunContinuationSnapshot,
            retentionDays As System.Int32)

            If snapshot Is Nothing Then Return
            Dim effectiveDays As System.Int32 = NormalizeContinuationRetentionDays(retentionDays)
            Dim nowUtc As System.DateTime = System.DateTime.UtcNow
            snapshot.RetentionDays = effectiveDays
            snapshot.UpdatedUtc = nowUtc
            snapshot.ExpiresUtc = nowUtc.AddDays(effectiveDays)
        End Sub

        Private Shared Function GetContinuationExpiresUtc(
            snapshot As ToolCallSequencing.ToolingRunContinuationSnapshot) As System.DateTime

            If snapshot Is Nothing Then Return System.DateTime.MinValue
            If snapshot.ExpiresUtc <> System.DateTime.MinValue Then
                Return snapshot.ExpiresUtc.ToUniversalTime()
            End If

            Dim anchorUtc As System.DateTime = snapshot.UpdatedUtc
            If anchorUtc = System.DateTime.MinValue Then anchorUtc = snapshot.CreatedUtc
            If anchorUtc = System.DateTime.MinValue Then Return System.DateTime.MinValue
            If anchorUtc.Kind <> System.DateTimeKind.Utc Then anchorUtc = anchorUtc.ToUniversalTime()

            Return anchorUtc.AddDays(NormalizeContinuationRetentionDays(snapshot.RetentionDays))
        End Function

        Private Shared Function IsContinuationExpired(
            snapshot As ToolCallSequencing.ToolingRunContinuationSnapshot) As System.Boolean

            Dim expiresUtc As System.DateTime = GetContinuationExpiresUtc(snapshot)
            If expiresUtc = System.DateTime.MinValue Then Return True
            Return expiresUtc <= System.DateTime.UtcNow
        End Function

        Private Shared Function ExtractContinuationUserLanguage(
            snapshot As ToolCallSequencing.ToolingRunContinuationSnapshot) As System.String

            If snapshot Is Nothing OrElse System.String.IsNullOrWhiteSpace(snapshot.SequencingStateJson) Then Return System.String.Empty
            Try
                Dim root As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(snapshot.SequencingStateJson)
                Return If(root.Value(Of System.String)("UserLanguage"), System.String.Empty).Trim()
            Catch
                Return System.String.Empty
            End Try
        End Function

        Private Shared Function NormalizeUserLanguageKey(value As System.String) As System.String
            Dim language As System.String = If(value, System.String.Empty).Trim().ToLowerInvariant()
            Dim separator As System.Int32 = language.IndexOfAny(New System.Char() {"-"c, "_"c})
            If separator > 0 Then language = language.Substring(0, separator)
            Return language
        End Function

        Private Shared Sub PruneStaleContinuationArtifactsUnlocked(workflowDirectory As String)
            Try
                If String.IsNullOrWhiteSpace(workflowDirectory) OrElse Not Directory.Exists(workflowDirectory) Then Return

                Dim knownArtifactDirectories As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
                For Each checkpointPath As String In Directory.GetFiles(workflowDirectory, "*.json", SearchOption.TopDirectoryOnly)
                    Try
                        Dim envelope As WorkflowCheckpointEnvelope = LoadCheckpointEnvelopeFromPathUnlocked(checkpointPath)
                        If envelope Is Nothing OrElse envelope.ContinuationState Is Nothing Then Continue For
                        Dim workflowId As System.String = If(envelope.WorkflowId, System.String.Empty).Trim()
                        If workflowId = System.String.Empty Then Continue For

                        Dim artifactDirectory As System.String = GetContinuationArtifactDirectory(workflowId)
                        knownArtifactDirectories.Add(artifactDirectory)
                        If IsContinuationExpired(envelope.ContinuationState) AndAlso Directory.Exists(artifactDirectory) Then
                            Directory.Delete(artifactDirectory, recursive:=True)
                        End If
                    Catch
                    End Try
                Next

                ' Orphaned cache directories have no checkpoint from which to derive an exact
                ' expiry. Keep the historical seven-day safety net only for those orphans.
                Dim orphanCutoffUtc As DateTime = DateTime.UtcNow.AddDays(-DefaultContinuationRetentionDays)
                For Each artifactDirectory As String In Directory.GetDirectories(workflowDirectory, "*_artifacts", SearchOption.TopDirectoryOnly)
                    Try
                        If knownArtifactDirectories.Contains(artifactDirectory) Then Continue For
                        Dim info As New DirectoryInfo(artifactDirectory)
                        If info.LastWriteTimeUtc < orphanCutoffUtc Then Directory.Delete(artifactDirectory, recursive:=True)
                    Catch
                    End Try
                Next
            Catch
            End Try
        End Sub

        Private Shared Sub DeleteContinuationArtifactsUnlocked(workflowId As String)
            Try
                Dim artifactDirectory As String = GetContinuationArtifactDirectory(workflowId)
                If Directory.Exists(artifactDirectory) Then Directory.Delete(artifactDirectory, recursive:=True)
            Catch
            End Try
        End Sub

        Private Shared Function GetPrivateWorkflowDirectory() As String
            Dim root As String = TryGetSharedPath("INI_AgentResourcesPathLocal")

            If Not System.String.IsNullOrWhiteSpace(root) Then
                root = SharedLibrary.SharedMethods.ExpandEnvironmentVariables(root)
            End If

            If System.String.IsNullOrWhiteSpace(root) Then
                root = System.IO.Path.Combine(
                    System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData),
                    "RedInk")
            End If

            Return System.IO.Path.Combine(root, ".session", "workflow_runtime")
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