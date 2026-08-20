' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.AgentHost.vb
' Purpose: Implements ISubAgentHost for managing isolated sub-agent execution within the Outlook host.
'          Handles model swapping, tool registry snapshots, tool scope filtering, and execution isolation.
'
' Architecture:
'  - RunIsolatedToolingLoopAsync(): Entry point for sub-agent requests.
'      - Captures parent model configuration scope.
'      - Swaps to special-task model (from INI_AlternateModelPath).
'      - Initializes tool scope via SubAgentToolScopeInitializer with parent registry snapshot.
'      - Executes single isolated ExecuteToolingLoop pass with clean message history.
'      - Restores parent model configuration on completion.
'      - Returns final model response or error payload.
'  - Tool Registry Snapshot:
'      - Captures parent's AuthoritativeToolRegistrySnapshot before sub-agent run.
'      - Passes snapshot to tool scope initializer for safe tool selection.
'      - Filters allowed tools based on AllowedToolNames (whitelist).
'  - Execution Tracking:
'      - Maintains SubAgentInvocationCount and per-agent invocation counters.
'      - Logs nested invocation depth and registry state.
'      - Routes all logs through ToolingFileLogger with [subagent] marker.
'  - Error Handling:
'      - Returns structured error payloads (SubAgentRuntimeHardening) on:
'          - Missing parent registry snapshot.
'          - Unresolved required tools.
'          - Model configuration failures.
' =============================================================================

Option Strict Off
Option Explicit On

Imports System.Threading
Imports System.Threading.Tasks
Imports SharedLibrary
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedContext
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn
    Implements Agents.ISubAgentHost

    ''' <summary>
    ''' Sub-agent isolated tooling-loop pass. Snapshots the current model, swaps to the agent's
    ''' resolved special-task-model (fallback agentdefaultmodel), narrows tools, runs ONE
    ''' ExecuteToolingLoop with a clean message history, then restores the model.
    ''' </summary>
    Public Async Function RunIsolatedToolingLoopAsync(
        request As SharedLibrary.Agents.SubAgentRunRequest,
        ct As CancellationToken) As Task(Of String) _
        Implements SharedLibrary.Agents.ISubAgentHost.RunIsolatedToolingLoopAsync

        If request Is Nothing Then
            Return Newtonsoft.Json.JsonConvert.SerializeObject(
                New With {
                    Key .summary = "The requested sub-agent run is invalid.",
                    Key .result = New With {Key .status = "blocked", Key .reason = "missing_subagent_run_request"}
                })
        End If

        If request.RunnerRetryIndex < 0 OrElse request.RunnerRetryIndex > 1 Then
            Return Newtonsoft.Json.JsonConvert.SerializeObject(
                New With {
                    Key .summary = "The requested sub-agent retry index is invalid.",
                    Key .result = New With {Key .status = "blocked", Key .reason = "invalid_subagent_retry_index", Key .retry_index = request.RunnerRetryIndex}
                })
        End If

        Dim scope = CaptureModelConfigScope(_context)
        Dim modelKey As String = If(String.IsNullOrWhiteSpace(request.SpecialModelKey), "agentdefaultmodel", request.SpecialModelKey)
        Dim swapped As Boolean = False

        Dim subAgentTaskId As String =
            If(request.SubAgentTaskId, "").Trim()

        If subAgentTaskId = "" Then
            Return Newtonsoft.Json.JsonConvert.SerializeObject(
                New With {
                    Key .summary = "The requested sub-agent task is missing its explicit identity.",
                    Key .result = New With {Key .status = "blocked", Key .reason = "missing_subagent_task_id", Key .agent = request.AgentName}
                })
        End If

        If System.String.IsNullOrWhiteSpace(request.ExpectedArtifactsJson) Then
            Return Newtonsoft.Json.JsonConvert.SerializeObject(
                New With {
                    Key .summary = "The requested sub-agent task is missing its explicit expected-artifact contract.",
                    Key .result = New With {Key .status = "blocked", Key .reason = "missing_expected_artifacts", Key .subagent_task_id = subAgentTaskId, Key .agent = request.AgentName}
                })
        End If

        If _activeToolingContext Is Nothing OrElse
           _activeToolingContext.SequencingState Is Nothing OrElse
           _activeToolingContext.SequencingState.SubAgentTaskRegistry Is Nothing Then

            Return Newtonsoft.Json.JsonConvert.SerializeObject(
                New With {
                    Key .summary = "The requested sub-agent task cannot start without the parent task registry.",
                    Key .result = New With {Key .status = "blocked", Key .reason = "subagent_task_registry_unavailable", Key .subagent_task_id = subAgentTaskId, Key .agent = request.AgentName}
                })
        End If

        If Not _activeToolingContext.SequencingState.SubAgentTaskRegistry.TryBegin(
            request.AgentName,
            subAgentTaskId,
            allowActiveContinuation:=(request.RunnerRetryIndex = 1)) Then

            Return Newtonsoft.Json.JsonConvert.SerializeObject(
                New With {
                    Key .summary = "The requested sub-agent task is already active or terminal.",
                    Key .result = New With {Key .status = "unresolved", Key .reason = "subagent_task_not_startable", Key .subagent_task_id = subAgentTaskId, Key .agent = request.AgentName}
                })
        End If

        Try
            If String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                Throw New InvalidOperationException(
                    $"Sub-agent model '{modelKey}' could not be resolved because INI_AlternateModelPath is empty.")
            End If

            Try
                swapped = GetSpecialTaskModel(_context, INI_AlternateModelPath, modelKey)

                If Not swapped AndAlso Not modelKey.Equals("agentdefaultmodel", StringComparison.OrdinalIgnoreCase) Then
                    ToolingFileLogger.LogWarn(
                        $"[subagent-host] Model '{modelKey}' was not found. Falling back to 'agentdefaultmodel'.")
                    swapped = GetSpecialTaskModel(_context, INI_AlternateModelPath, "agentdefaultmodel")
                End If
            Catch ex As Exception
                Throw New InvalidOperationException(
                    $"Sub-agent model '{modelKey}' could not be resolved via GetSpecialTaskModel: {ex.Message}", ex)
            End Try

            If Not swapped Then
                Throw New InvalidOperationException(
                    $"Sub-agent model '{modelKey}' could not be resolved via GetSpecialTaskModel.")
            End If

            ToolingFileLogger.LogStep(
                SharedLibrary.Agents.WorkflowContinuity.ComposeWorkflowLogMessage(
                    "Sub-agent model resolved via GetSpecialTaskModel('" & modelKey & "').",
                    If(request.WorkflowId, ""),
                    "sub_agent_host_ready",
                    agentName:=If(request.AgentName, ""),
                    hostName:="Outlook",
                    leadingMarker:="[subagent]"))

            ToolingFileLogger.LogStep(
                SharedLibrary.Agents.WorkflowContinuity.ComposeWorkflowLogMessage(
                    "Secondary API enforced for isolated sub-agent execution.",
                    If(request.WorkflowId, ""),
                    "sub_agent_host_ready",
                    agentName:=If(request.AgentName, ""),
                    hostName:="Outlook",
                    leadingMarker:="[subagent]"))

            ToolingFileLogger.LogStep(
                SharedLibrary.Agents.WorkflowContinuity.ComposeWorkflowLogMessage(
                    "Sub-agent host initialized.",
                    If(request.WorkflowId, ""),
                    "sub_agent_host_ready",
                    agentName:=If(request.AgentName, ""),
                    hostName:="Outlook",
                    leadingMarker:="[subagent]") &
                " [allowedTools: " &
                If(request.AllowedToolNames Is Nothing,
                   "(default-host-scope)",
                   If(request.AllowedToolNames.Count = 0,
                      "(none)",
                      String.Join(", ", request.AllowedToolNames))) &
                "]")

            ToolingFileLogger.LogStep(
                SharedLibrary.Agents.WorkflowContinuity.ComposeWorkflowLogMessage(
                    "Shared tool-call sequencing is active.",
                    If(request.WorkflowId, ""),
                    "sub_agent_host_ready",
                    agentName:=If(request.AgentName, ""),
                    hostName:="Outlook",
                    leadingMarker:="[subagent]"))

            Dim registrySource As String = "parent_registry_snapshot"
            Dim parentRunId As String = "(no_parent_run)"
            Dim invocationIndex As Integer = 1
            Dim sameAgentInvocationCount As Integer = 1

            If _activeToolingContext IsNot Nothing Then
                parentRunId = If(_activeToolingContext.RunId, "(no_parent_run)")
                _activeToolingContext.SubAgentInvocationCount += 1
                invocationIndex = _activeToolingContext.SubAgentInvocationCount

                Dim existingAgentInvocationCount As Integer = 0

                If _activeToolingContext.SubAgentInvocationCountsByAgent.TryGetValue(request.AgentName, existingAgentInvocationCount) Then
                    sameAgentInvocationCount = existingAgentInvocationCount + 1
                    _activeToolingContext.SubAgentInvocationCountsByAgent(request.AgentName) = sameAgentInvocationCount
                Else
                    sameAgentInvocationCount = 1
                    _activeToolingContext.SubAgentInvocationCountsByAgent(request.AgentName) = 1
                End If
            End If

            Dim parentRegistrySnapshot As SharedLibrary.Agents.ToolRegistry = Nothing

            If _activeToolingContext IsNot Nothing Then
                parentRegistrySnapshot = _activeToolingContext.AuthoritativeToolRegistrySnapshot
            End If

            Dim authoritativeSnapshotAvailable As Boolean = (parentRegistrySnapshot IsNot Nothing)
            Dim authoritativeSnapshot As SharedLibrary.Agents.ToolRegistry =
    If(authoritativeSnapshotAvailable,
       parentRegistrySnapshot.Snapshot(),
       Nothing)

            If Not authoritativeSnapshotAvailable Then
                registrySource = "parent_registry_snapshot_missing"
            End If

            Dim snapshotToolCount As Integer =
    If(authoritativeSnapshot Is Nothing,
       0,
       authoritativeSnapshot.ListNames().Count)

            Dim expandedAllowedToolNames As IReadOnlyList(Of String) =
    ExpandAllowedToolNamesForRegistry(
        request.AllowedToolNames,
        authoritativeSnapshot)

            Dim preflight = SharedLibrary.Agents.SubAgentToolScopeInitializer.Initialize(
    authoritativeSnapshot,
    expandedAllowedToolNames)

            Dim requestedNamesText As String =
    If(expandedAllowedToolNames Is Nothing OrElse expandedAllowedToolNames.Count = 0,
       "(none)",
       String.Join(", ", expandedAllowedToolNames))

            Dim resolvedNamesText As String =
    If(preflight.ResolvedToolNames.Count = 0,
       "(none)",
       String.Join(", ", preflight.ResolvedToolNames))

            Dim missingNamesText As String =
    If(preflight.MissingToolNames.Count = 0,
       "(none)",
       String.Join(", ", preflight.MissingToolNames))

            Dim finalSelectedNamesText As String =
    If(preflight.FinalSelectedToolNames.Count = 0,
       "(none)",
       String.Join(", ", preflight.FinalSelectedToolNames))

            Dim repeatedInvocationLabel As String =
    If(sameAgentInvocationCount > 1, "later", "first")

            ToolingFileLogger.LogStep(
                SharedLibrary.Agents.WorkflowContinuity.ComposeWorkflowLogMessage(
                    "Sub-agent tool scope initialized.",
                    If(request.WorkflowId, ""),
                    "sub_agent_scope_init",
                    agentName:=If(request.AgentName, ""),
                    hostName:="Outlook",
                    leadingMarker:="[subagent]") &
                " [parentRunId: " & parentRunId & "]" &
                " [invocationIndex: " & invocationIndex & "]" &
                " [parentRegistrySnapshotExists: " & authoritativeSnapshotAvailable & "]" &
                " [snapshotToolCount: " & snapshotToolCount & "]" &
                " [requestedAllowedTools: " & requestedNamesText & "]" &
                " [resolvedTools: " & resolvedNamesText & "]" &
                " [missingTools: " & missingNamesText & "]" &
                " [finalSelectedTools: " & finalSelectedNamesText & "]" &
                " [sameAgentInvocation: " & repeatedInvocationLabel & "]")

            If Not authoritativeSnapshotAvailable Then
                Dim payload = SharedLibrary.Agents.SubAgentRuntimeHardening.BuildParentRegistryMissingPayload(
        requestedToolNames:=preflight.RequestedToolNames)

                ToolingFileLogger.LogError(
        "[subagent-host] Parent tool registry snapshot missing before sub-agent model call.",
        details:=$"parentRunId={parentRunId}; invocationIndex={invocationIndex}; agent={If(request.AgentName, "")}; requested={requestedNamesText}")

                If subAgentTaskId <> "" AndAlso
                   _activeToolingContext IsNot Nothing AndAlso
                   _activeToolingContext.SequencingState IsNot Nothing AndAlso
                   _activeToolingContext.SequencingState.SubAgentTaskRegistry IsNot Nothing Then

                    _activeToolingContext.SequencingState.SubAgentTaskRegistry.MarkBlocked(
                        request.AgentName,
                        subAgentTaskId,
                        "parent_registry_snapshot_missing")
                End If

                Return payload
            End If

            If preflight.HasRequestedTools AndAlso
   (preflight.HasMissingRequestedTools OrElse preflight.HasMissingFinalToolNames) Then

                Dim missingRequiredToolNames As List(Of String) = preflight.MissingFinalToolNames
                If missingRequiredToolNames.Count = 0 Then
                    missingRequiredToolNames = New List(Of String)(preflight.MissingToolNames)
                End If

                Dim payload = SharedLibrary.Agents.SubAgentRuntimeHardening.BuildRequiredToolMissingPayload(
        missingRequiredToolNames,
        requestedToolNames:=preflight.RequestedToolNames,
        resolvedToolNames:=preflight.ResolvedToolNames)

                ToolingFileLogger.LogError(
        "[subagent-host] Sub-agent required tools could not be resolved before model call.",
        details:=$"parentRunId={parentRunId}; invocationIndex={invocationIndex}; agent={If(request.AgentName, "")}; requested={requestedNamesText}; resolved={resolvedNamesText}; missing={missingNamesText}; finalSelected={finalSelectedNamesText}")

                If subAgentTaskId <> "" AndAlso
                   _activeToolingContext IsNot Nothing AndAlso
                   _activeToolingContext.SequencingState IsNot Nothing AndAlso
                   _activeToolingContext.SequencingState.SubAgentTaskRegistry IsNot Nothing Then

                    _activeToolingContext.SequencingState.SubAgentTaskRegistry.MarkBlocked(
                        request.AgentName,
                        subAgentTaskId,
                        "required_subagent_tools_missing")
                End If

                Return payload
            End If

            Dim allTools As List(Of ModelConfig) = authoritativeSnapshot.MaterializeAll()

            ToolingFileLogger.LogStep("[subagent-host] Resolved sub-agent tool universe: " &
                          If(allTools.Count = 0,
                             "(none)",
                             String.Join(", ", allTools.Select(Function(t) t.ToolName))))

            Dim result = Await ExecuteToolingLoop(
                    sysCommand:=request.SystemPrompt,
                    userText:="",
                    selectedTools:=allTools,
                    useSecondAPI:=True,
                    fullPromptOverride:=request.UserMessage,
                    hideSplash:=True,
                    hideLogWindow:=True,
                    subAgentMode:=True,
                    subAgentAllowedToolNames:=expandedAllowedToolNames,
                    subAgentSpecialModelKey:=request.SpecialModelKey,
                    subAgentAuthoritativeRegistry:=authoritativeSnapshot,
                    subAgentRegistrySource:=registrySource,
                    subAgentParentRunId:=parentRunId,
                    subAgentInvocationIndex:=invocationIndex,
                    subAgentAgentInvocationCount:=sameAgentInvocationCount,
                    subAgentName:=request.AgentName,
                    workflowId:=request.WorkflowId,
                    finalResponseContract:=SharedLibrary.Agents.ToolingFinalResponseContract.RawCallerText,
                    subAgentExpectedArtifactsJson:=request.ExpectedArtifactsJson).ConfigureAwait(False)

            If subAgentTaskId <> "" AndAlso
               _activeToolingContext IsNot Nothing AndAlso
               _activeToolingContext.SequencingState IsNot Nothing AndAlso
               _activeToolingContext.SequencingState.SubAgentTaskRegistry IsNot Nothing Then

                Dim normalized =
                    SharedLibrary.Agents.SubAgentRuntimeHardening.NormalizeFinalOutput(
                        result,
                        jsonRequired:=True)

                Dim errorCode As String =
                    If(normalized Is Nothing, "", normalized.GetErrorCode())

                Dim isRetryableEmptyResult As Boolean =
                    normalized IsNot Nothing AndAlso
                    normalized.IsError AndAlso
                    (String.Equals(
                        errorCode,
                        SharedLibrary.Agents.SubAgentRuntimeHardening.EmptyResultCode,
                        StringComparison.OrdinalIgnoreCase) OrElse
                     String.Equals(
                        errorCode,
                        SharedLibrary.Agents.SubAgentRuntimeHardening.ModelEmptyResponseCode,
                        StringComparison.OrdinalIgnoreCase))

                If isRetryableEmptyResult AndAlso request.RunnerRetryIndex = 0 Then
                    ' Deliberately leave the task Active. SubAgentRunner owns the one
                    ' permitted internal empty-response retry and will call this host
                    ' again with the same SubAgentTaskId and RunnerRetryIndex = 1.

                ElseIf normalized IsNot Nothing AndAlso Not normalized.IsError Then

                    _activeToolingContext.SequencingState.SubAgentTaskRegistry.MarkCompleted(
                        request.AgentName,
                        subAgentTaskId)

                ElseIf isRetryableEmptyResult Then

                    _activeToolingContext.SequencingState.SubAgentTaskRegistry.MarkUnresolved(
                        request.AgentName,
                        subAgentTaskId,
                        If(errorCode, ""))

                Else
                    _activeToolingContext.SequencingState.SubAgentTaskRegistry.MarkBlocked(
                        request.AgentName,
                        subAgentTaskId,
                        If(errorCode, "subagent_failed"))
                End If
            End If

            Return If(result, "")
        Catch ex As Exception
            If subAgentTaskId <> "" AndAlso
               _activeToolingContext IsNot Nothing AndAlso
               _activeToolingContext.SequencingState IsNot Nothing AndAlso
               _activeToolingContext.SequencingState.SubAgentTaskRegistry IsNot Nothing Then

                _activeToolingContext.SequencingState.SubAgentTaskRegistry.MarkBlocked(
                    request.AgentName,
                    subAgentTaskId,
                    "subagent_host_exception")
            End If

            Throw
        Finally
            RestoreModelConfigScope(_context, scope)
        End Try
    End Function

End Class
