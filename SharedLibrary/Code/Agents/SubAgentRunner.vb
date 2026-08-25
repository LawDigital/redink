' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: SubAgentRunner.vb
' Purpose: Orchestrates invocation of a sub-agent (Claude-style AGENT.md):
'           1. Locks AgentGate as OWNER for the whole run (nested LLM/MCP calls
'              are re-entrant against the same logical owner).
'           2. Composes clean system prompt from AGENT.md body (no Inky.md,
'              no parent system prompt) and runs ONE isolated tooling-loop via host.
'           3. Parses final text as {summary, result} JSON; preserves direct JSON
'              objects/arrays as structured results instead of stringifying.
'           4. Validates final output is not semantically empty.
'           5. Retries agent_empty_result exactly once with stricter reminder.
'           6. Optionally stores result in SessionMemory and returns compact
'              tool-response JSON with memory stub.
'
' Concurrency: Owner-scope on AgentGate ensures only one sub-agent runs globally.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Diagnostics
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Agents

    Public NotInheritable Class SubAgentRunner

        Private Sub New()
        End Sub

        Private Const RetryReminderText As String =
            "Your previous response was empty or unusable. Return one usable final answer in the requested format. If you cannot complete the task, return a structured error object."

        Public Shared Async Function InvokeAsync(host As ISubAgentHost,
                                         agentName As String,
                                         task As String,
                                         Optional contextBlob As String = Nothing,
                                         Optional storeResultInMemory As Boolean = True,
                                         Optional workflowId As String = Nothing,
                                         Optional subAgentTaskId As String = Nothing,
                                         Optional cancellationToken As CancellationToken = Nothing,
                                         Optional expectedArtifactsJson As String = Nothing) As Task(Of String)

            If host Is Nothing Then
                Return BuildInfrastructureErrorPayload(agentName, "no_host", "invoke", "No sub-agent host is available.")
            End If

            Dim ag = AgentResources.FindAgent(agentName)
            If ag Is Nothing Then
                Return BuildInfrastructureErrorPayload(agentName, "agent_not_found", "invoke", "The requested sub-agent was not found.")
            End If

            Dim effectiveWorkflowId As String =
        If(String.IsNullOrWhiteSpace(workflowId), WorkflowContinuity.CurrentWorkflowId, workflowId)

            Return Await InvokeResolvedAsync(
        host,
        ag,
        task,
        contextBlob,
        storeResultInMemory,
        effectiveWorkflowId,
        subAgentTaskId,
        cancellationToken,
        expectedArtifactsJson).ConfigureAwait(False)
        End Function

        Friend Shared Async Function InvokeResolvedAsync(host As ISubAgentHost,
                                                 ag As AgentDescriptor,
                                                 task As String,
                                                 Optional contextBlob As String = Nothing,
                                                 Optional storeResultInMemory As Boolean = True,
                                                 Optional workflowId As String = Nothing,
                                                 Optional subAgentTaskId As String = Nothing,
                                                 Optional cancellationToken As CancellationToken = Nothing,
                                                 Optional expectedArtifactsJson As String = Nothing) As Task(Of String)
            If host Is Nothing Then
                Return BuildInfrastructureErrorPayload(If(ag?.Name, ""), "no_host", "invoke", "No sub-agent host is available.")
            End If

            If ag Is Nothing OrElse String.IsNullOrWhiteSpace(ag.Name) Then
                Return BuildInfrastructureErrorPayload(If(ag?.Name, ""), "agent_not_found", "invoke", "The requested sub-agent was not found.")
            End If

            If String.IsNullOrWhiteSpace(task) Then
                Return BuildInfrastructureErrorPayload(ag.Name, "missing_task", "invoke", "The delegated sub-agent task is empty.")
            End If

            Dim normalizedSubAgentTaskId As String = If(subAgentTaskId, "").Trim()
            If normalizedSubAgentTaskId = "" Then
                Return BuildInfrastructureErrorPayload(ag.Name, "missing_subagent_task_id", "invoke", "Every delegated sub-agent task requires an explicit opaque subagent_task_id.")
            End If

            If String.IsNullOrWhiteSpace(expectedArtifactsJson) Then
                Return BuildInfrastructureErrorPayload(ag.Name, "missing_expected_artifacts", "invoke", "Every delegated sub-agent task requires an explicit expected_artifacts JSON array; use [] for no user-final artifacts.")
            End If

            Dim expectedArtifactsToken As Newtonsoft.Json.Linq.JToken = Nothing
            Try
                expectedArtifactsToken = Newtonsoft.Json.Linq.JToken.Parse(expectedArtifactsJson)
            Catch ex As System.Exception
                Return BuildInfrastructureErrorPayload(ag.Name, "invalid_expected_artifacts", "invoke", "expected_artifacts is not valid JSON: " & ex.Message)
            End Try

            If expectedArtifactsToken Is Nothing OrElse
               expectedArtifactsToken.Type <> Newtonsoft.Json.Linq.JTokenType.Array Then
                Return BuildInfrastructureErrorPayload(ag.Name, "invalid_expected_artifacts", "invoke", "expected_artifacts must be a JSON array; use [] for no user-final artifacts.")
            End If

            For Each expectedArtifactToken As Newtonsoft.Json.Linq.JToken In DirectCast(expectedArtifactsToken, Newtonsoft.Json.Linq.JArray)
                Dim expectedArtifactObject As Newtonsoft.Json.Linq.JObject = TryCast(expectedArtifactToken, Newtonsoft.Json.Linq.JObject)
                If expectedArtifactObject Is Nothing Then
                    Return BuildInfrastructureErrorPayload(ag.Name, "invalid_expected_artifacts", "invoke", "Each expected_artifacts item must be an object with explicit opaque logical_deliverable_id and output_slot_id values.")
                End If

                Dim logicalDeliverableId As String = If(expectedArtifactObject.Value(Of String)("logical_deliverable_id"), "").Trim()
                Dim outputSlotId As String = If(expectedArtifactObject.Value(Of String)("output_slot_id"), "").Trim()
                If logicalDeliverableId = "" OrElse outputSlotId = "" Then
                    Return BuildInfrastructureErrorPayload(ag.Name, "invalid_expected_artifacts", "invoke", "Each expected_artifacts item requires non-empty opaque logical_deliverable_id and output_slot_id values.")
                End If
            Next

            Dim effectiveWorkflowId As String =
        If(String.IsNullOrWhiteSpace(workflowId), WorkflowContinuity.CurrentWorkflowId, workflowId)

            If Not String.IsNullOrWhiteSpace(effectiveWorkflowId) Then
                Dim checkpointWritten =
            WorkflowContinuity.NoteSubAgentInvoked(
                effectiveWorkflowId,
                WorkflowContinuity.CurrentHostPipeline,
                ag.Name)

                Debug.WriteLine(
            WorkflowContinuity.BuildWorkflowLogLabel(
                effectiveWorkflowId,
                "sub_agent_invoked",
                agentName:=ag.Name) &
            " checkpoint_written=" & If(checkpointWritten, "true", "false"))
            End If

            Dim sys As New StringBuilder()

            If Not String.IsNullOrWhiteSpace(ag.Description) Then
                sys.AppendLine(ag.Description.Trim())
                sys.AppendLine()
            End If

            sys.Append(ag.LoadBody())
            sys.AppendLine()
            sys.AppendLine()
            sys.AppendLine("Output contract: your FINAL message MUST be one usable final answer in the requested format. If you cannot complete the task, return a structured error object.")

            Dim baseUserMessage As New StringBuilder()
            baseUserMessage.Append("Task:").AppendLine().Append(If(task, "").Trim())

            If Not String.IsNullOrWhiteSpace(contextBlob) Then
                baseUserMessage.AppendLine().AppendLine().AppendLine("Context (provided by parent):")
                baseUserMessage.Append(contextBlob.Trim())
            End If

            Dim lockedExpectedArtifactsJson As String =
                expectedArtifactsToken.ToString(Newtonsoft.Json.Formatting.None)

            baseUserMessage.AppendLine().AppendLine()
            baseUserMessage.AppendLine("Locked expected-artifact contract (provided by parent/orchestrator):")
            baseUserMessage.AppendLine(lockedExpectedArtifactsJson)
            baseUserMessage.AppendLine("Use these opaque logical_deliverable_id/output_slot_id values unchanged for any user-facing final artifacts. Do not add, rename, derive, or substitute slots. If the array is empty, do not produce a user-facing final artifact.")

            Dim allowedTools As IReadOnlyList(Of String) =
        If(ag.AllowedTools Is Nothing,
           CType(Array.Empty(Of String)(), IReadOnlyList(Of String)),
           ag.AllowedTools.AsReadOnly())

            Dim optionalTools As IReadOnlyList(Of String) =
        If(ag.OptionalTools Is Nothing,
           CType(Array.Empty(Of String)(), IReadOnlyList(Of String)),
           ag.OptionalTools.AsReadOnly())

            Dim retryCount As Integer = 0
            Dim userMessageForRun As String = baseUserMessage.ToString()

            Await AgentGate.EnterAsync(cancellationToken).ConfigureAwait(False)
            AgentGate.MarkCurrentFlowAsOwner()

            Try
                Do
                    Dim req As New SubAgentRunRequest With {
                .AgentName = ag.Name,
                .SystemPrompt = sys.ToString(),
                .UserMessage = userMessageForRun,
                .SpecialModelKey = If(String.IsNullOrWhiteSpace(ag.Model), "agentdefaultmodel", ag.Model),
                .AllowedToolNames = allowedTools,
                .OptionalToolNames = optionalTools,
                .MaxIterations = 0,
                .TimeoutSeconds = ag.TimeoutSeconds,
                .WorkflowId = effectiveWorkflowId,
                .SubAgentTaskId = normalizedSubAgentTaskId,
                .RunnerRetryIndex = retryCount,
                .ExpectedArtifactsJson = lockedExpectedArtifactsJson
            }

                    Debug.WriteLine(
                WorkflowContinuity.BuildWorkflowLogLabel(
                    effectiveWorkflowId,
                    "sub_agent_invoked",
                    agentName:=ag.Name) &
                " allowed_tools=" & FormatAllowedTools(req.AllowedToolNames) &
                " optional_tools=" & FormatAllowedTools(req.OptionalToolNames) &
                " retry=" & retryCount)

                    Dim finalText As String = Nothing

                    Try
                        finalText = Await host.RunIsolatedToolingLoopAsync(req, cancellationToken).ConfigureAwait(False)
                    Catch oce As OperationCanceledException
                        Throw
                    Catch ex As Exception
                        Return BuildInfrastructureErrorPayload(ag.Name, "agent_failed", "invoke", ex.Message)
                    End Try

                    Dim normalized = SubAgentRuntimeHardening.NormalizeFinalOutput(finalText, jsonRequired:=True)
                    LogFinalOutputDiagnostics(ag.Name, req.AllowedToolNames, normalized, retryCount)

                    If Not normalized.IsError Then
                        Dim resp As JObject = normalized.ToJObject()
                        resp("agent") = ag.Name

                        If storeResultInMemory Then
                            Try
                                Dim key As String = "agent_" & ag.Name & "_" & DateTime.UtcNow.ToString("yyyyMMddHHmmssfff")
                                Dim storedResult As JToken = If(normalized.Result Is Nothing, JValue.CreateNull(), normalized.Result.DeepClone())

                                Dim metadata As New SessionMemoryMetadata With {
                            .WorkflowId = effectiveWorkflowId,
                            .Source = "agent",
                            .ContentKind = "tool_result",
                            .RelatedAgent = ag.Name,
                            .CreatedAt = DateTime.UtcNow,
                            .TrustedForRuntime = False,
                            .TrustLevel = "advisory"
                        }

                                Dim entry = SessionMemory.Put(
                            key,
                            If(normalized.Summary, "Result of sub-agent '" & ag.Name & "'."),
                            storedResult,
                            tags:={"agent", ag.Name},
                            metadata:=metadata)

                                resp("memory_key") = entry.Key
                                resp("stub") = SessionMemory.BuildStub(entry)
                            Catch
                            End Try
                        End If

                        If Not String.IsNullOrWhiteSpace(effectiveWorkflowId) Then
                            Dim checkpointWritten =
                        WorkflowContinuity.NoteSubAgentReturned(
                            effectiveWorkflowId,
                            WorkflowContinuity.CurrentHostPipeline,
                            ag.Name,
                            succeeded:=True)

                            Debug.WriteLine(
                        WorkflowContinuity.BuildWorkflowLogLabel(
                            effectiveWorkflowId,
                            "sub_agent_returned",
                            agentName:=ag.Name) &
                        " success=true checkpoint_written=" & If(checkpointWritten, "true", "false"))
                        End If

                        Return resp.ToString(Formatting.None)
                    End If

                    Dim retryableErrorCodes As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
                SubAgentRuntimeHardening.EmptyResultCode,
                SubAgentRuntimeHardening.ModelEmptyResponseCode
            }

                    If retryCount = 0 AndAlso retryableErrorCodes.Contains(normalized.GetErrorCode()) Then
                        retryCount += 1
                        userMessageForRun = BuildRetryUserMessage(baseUserMessage.ToString(), normalized, allowedTools)
                        Continue Do
                    End If

                    Dim errResp As JObject = normalized.ToJObject()
                    errResp("agent") = ag.Name
                    errResp("retryCount") = retryCount

                    If Not String.IsNullOrWhiteSpace(effectiveWorkflowId) Then
                        Dim checkpointWritten =
                    WorkflowContinuity.NoteSubAgentReturned(
                        effectiveWorkflowId,
                        WorkflowContinuity.CurrentHostPipeline,
                        ag.Name,
                        succeeded:=False)

                        Debug.WriteLine(
                    WorkflowContinuity.BuildWorkflowLogLabel(
                        effectiveWorkflowId,
                        "sub_agent_returned",
                        agentName:=ag.Name) &
                    " success=false checkpoint_written=" & If(checkpointWritten, "true", "false"))
                    End If

                    Return errResp.ToString(Formatting.None)
                Loop
            Finally
                AgentGate.UnmarkCurrentFlowAsOwner()
                AgentGate.Release()
            End Try
        End Function

        Private Shared Function BuildRetryUserMessage(baseUserMessage As String,
                                              normalized As SubAgentRuntimeHardening.NormalizedEnvelope,
                                              allowedTools As IReadOnlyList(Of String)) As String
            Dim sb As New StringBuilder(If(baseUserMessage, "").TrimEnd())

            If sb.Length > 0 Then
                sb.AppendLine()
                sb.AppendLine()
            End If

            sb.AppendLine(RetryReminderText)

            Dim errObj As JObject = Nothing
            If normalized IsNot Nothing Then
                errObj = normalized.Error
            End If

            Dim lastToolName As String = If(errObj?.Value(Of String)("lastToolName"), "")
            Dim lastToolResultSummary As String = If(errObj?.Value(Of String)("lastToolResultSummary"), "")
            Dim retryHint As String = If(errObj?.Value(Of String)("retryHint"), "")
            Dim compactedToolResponse As Boolean = False

            If errObj IsNot Nothing AndAlso errObj("compactedToolResponse") IsNot Nothing Then
                compactedToolResponse = errObj.Value(Of Boolean)("compactedToolResponse")
            End If

            sb.AppendLine("Recovery requirements:")
            sb.AppendLine("- Return the required final JSON object now, or call one smaller follow-up tool.")
            sb.AppendLine("- Do not return empty content.")
            sb.AppendLine("- Do not restate or paste a large raw tool response.")

            If Not String.IsNullOrWhiteSpace(lastToolName) Then
                sb.AppendLine("- Last successful tool: " & lastToolName)
            End If

            If Not String.IsNullOrWhiteSpace(lastToolResultSummary) Then
                sb.AppendLine("- Last tool result summary: " & lastToolResultSummary)
            End If

            If compactedToolResponse Then
                sb.AppendLine("- The prior tool result was compacted. If you need more source text, request a smaller chunk using max_chars and start_char/offset.")
            End If

            If Not String.IsNullOrWhiteSpace(retryHint) Then
                sb.AppendLine("- Retry hint: " & retryHint)
            End If

            If allowedTools IsNot Nothing AndAlso allowedTools.Count > 0 Then
                sb.AppendLine("- Available tools: " & String.Join(", ", allowedTools))
            End If

            Return sb.ToString().TrimEnd()
        End Function

        Private Shared Function BuildInfrastructureErrorPayload(agentName As String,
                                                                errorCode As String,
                                                                phase As String,
                                                                message As String) As String
            Dim obj As New JObject(
                New JProperty("summary", If(message, "Sub-agent failed.")),
                New JProperty("result", JValue.CreateNull()),
                New JProperty("resultKind", "error"),
                New JProperty("rawLength", 0),
                New JProperty("error", New JObject(
                    New JProperty("code", errorCode),
                    New JProperty("phase", phase),
                    New JProperty("message", If(message, "")))))

            If Not String.IsNullOrWhiteSpace(agentName) Then
                obj("agent") = agentName
            End If

            Return obj.ToString(Formatting.None)
        End Function

        Private Shared Sub LogFinalOutputDiagnostics(agentName As String,
                                                     allowedTools As IReadOnlyList(Of String),
                                                     normalized As SubAgentRuntimeHardening.NormalizedEnvelope,
                                                     retryCount As Integer)
            Dim errorCode As String = ""
            If normalized IsNot Nothing Then
                errorCode = normalized.GetErrorCode()
            End If

            Debug.WriteLine(
                $"[SubAgentRunner] agent='{agentName}' allowed_tools={FormatAllowedTools(allowedTools)} final_len={If(normalized?.RawLength, 0)} resultKind={If(normalized?.ResultKind, "")} retry={retryCount} errorCode={If(errorCode, "")}")
        End Sub

        Private Shared Function FormatAllowedTools(allowedTools As IReadOnlyList(Of String)) As String
            If allowedTools Is Nothing Then Return "(default-host-scope)"
            If allowedTools.Count = 0 Then Return "(none)"
            Return String.Join(", ", allowedTools)
        End Function

    End Class

End Namespace
