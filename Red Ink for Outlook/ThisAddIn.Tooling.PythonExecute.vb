' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.Tooling.PythonExecute.vb
' Purpose: Integrates the shared `python_execute` tool into the Outlook host,
'          including configuration, input resolution, output publishing, and
'          host-service bridging for the active tooling loop.
'
' Architecture:
'  - Tool Registration:
'      - `TryConfigureAndBuildPythonExecuteTool()` parses `INI_PythonAgentPath`,
'        validates availability, configures shared core options, and builds the
'        advertised `ModelConfig` when the executor is usable.
'  - Input / Output Bridging:
'      - `ResolvePythonInputFile()` resolves staged session attachments first,
'        then falls back to workspace-relative paths.
'      - `PublishPythonAgentOutput()` routes produced files either into the
'        session sink (chat / AutoPilot) or into the connected workspace.
'  - Tool Execution:
'      - `ExecutePythonExecuteTool()` converts Outlook `ToolCall` state into the
'        shared `PythonExecuteTool` execution flow and maps results back into
'        `ToolResponse`.
'  - Host-Mediated Capabilities:
'      - `BuildPythonHostServiceHandler()` exposes only the currently selected
'        LLM / web retrieval / search capabilities to the Python worker.
' =============================================================================


Option Explicit On
Option Strict Off

Imports System.Threading
Imports System.Threading.Tasks
Imports SharedLibrary
Imports SharedLibrary.SharedLibrary

Partial Public Class ThisAddIn

    ''' <summary>
    ''' Configures the secure python_execute tool from INI_PythonAgentPath and, when available,
    ''' builds its ModelConfig. Returns False (tool unregistered) when the executor path is not
    ''' set, the exe is missing, or authenticity verification fails.
    ''' </summary>
    Private Function TryConfigureAndBuildPythonExecuteTool(ByRef modelConfig As ModelConfig) As Boolean
        modelConfig = Nothing

        Dim configuration = Agents.PythonExecuteToolConfig.Parse(SharedMethods.ExpandEnvironmentVariables(INI_PythonAgentPath))
        If configuration Is Nothing Then
            Return False
        End If

        Try
            Agents.PythonExecuteTool.Configure(New Agents.PythonExecuteToolCoreOptions() With {
                .AgentConfiguration = configuration,
                .InputFileResolver = Function(rel) ResolvePythonInputFile(rel),
                .PublishOutputFile = Sub(output) PublishPythonAgentOutput(output),
                .HostServiceHandler = Nothing
            })
        Catch ex As Exception
            ToolingFileLogger.LogWarn("python_execute is unavailable and will not be advertised.", ex:=ex)
            Return False
        End Try

        Return Agents.PythonExecuteTool.TryBuild(_context, modelConfig, toolPriority:=996, displaySuffix:=InternalToolSuffix)
    End Function

    ''' <summary>
    ''' Resolves a python_execute input reference. Session attachments (files dropped into the
    ''' Local Agent window or produced by prior tool calls, staged in the per-session temp
    ''' directory) are preferred, mirroring ExecuteWorkspaceRead. Falls back to a workspace-relative
    ''' path when no attachment matches.
    ''' </summary>
    Private Function ResolvePythonInputFile(rel As String) As Agents.RedInkPythonAgentInputFile
        Dim staged As AutoPilotAttachmentInfo = StageWorkspaceFile(rel)
        If staged Is Nothing Then
            staged = FindAttachment(rel)
        End If

        If staged IsNot Nothing AndAlso
           Not String.IsNullOrWhiteSpace(staged.TempFilePath) AndAlso
           System.IO.File.Exists(staged.TempFilePath) Then
            Return New Agents.RedInkPythonAgentInputFile(staged.TempFilePath, rel)
        End If

        Return New Agents.RedInkPythonAgentInputFile(ResolveWorkspacePath(rel), rel)
    End Function

    ''' <summary>
    ''' Routes a python_execute output artifact back to the correct destination for the active host:
    ''' the session sink (Local Chat / AutoPilot / Scheduler) so it is visible to read_attachment,
    ''' list_attachments, and post-run delivery; otherwise the connected workspace, falling back to
    ''' the default writable root (Desktop) when no workspace exists.
    ''' </summary>
    Private Sub PublishPythonAgentOutput(output As Agents.RedInkPythonAgentOutput)
        If output Is Nothing OrElse String.IsNullOrWhiteSpace(output.FullPath) OrElse Not System.IO.File.Exists(output.FullPath) Then
            Return
        End If

        ' The core has already validated containment/size/hash and populated PublishedSubPath.
        Dim subPath As String = If(output.PublishedSubPath, "").Trim()
        If subPath.Length = 0 Then subPath = System.IO.Path.GetFileName(output.FullPath)
        subPath = subPath.Replace("/"c, System.IO.Path.DirectorySeparatorChar).TrimStart(System.IO.Path.DirectorySeparatorChar)
        Dim fileName As String = System.IO.Path.GetFileName(subPath)
        If String.IsNullOrWhiteSpace(fileName) Then Return

        ' Session sink (Local Chat / AutoPilot / Scheduler): name-addressed, so flatten with a
        ' collision-safe suffix so distinct outputs sharing a filename do not overwrite each other.
        If _chatAgentActive OrElse _apActive Then
            Try
                Dim tempDir As String = EnsureChatAgentTempDir()
                Dim destPath As String = System.IO.Path.Combine(tempDir, fileName)
                Dim counter As Integer = 1
                While System.IO.File.Exists(destPath)
                    destPath = System.IO.Path.Combine(tempDir,
                        System.IO.Path.GetFileNameWithoutExtension(fileName) & $"_{counter}" & System.IO.Path.GetExtension(fileName))
                    counter += 1
                End While
                ' Copy the produced file into the per-session temp dir and REGISTER it so it becomes
                ' usable by future calls (read_attachment, list_attachments, workspace save) and appears
                ' as a clickable session chip in the browser. Registration keeps the file inside the
                ' temp dir, so CollectResultAttachments still picks it up and ChatAgentCollectAndCopyOutputs
                ' still delivers it to Desktop\Inky\<timestamp>\. IsToolOutput distinguishes produced
                ' artifacts from user uploads.
                System.IO.File.Copy(output.FullPath, destPath, overwrite:=False)
                RegisterSessionFile(destPath, "Produced by python_execute", isToolOutput:=True)
                Return
            Catch ex As Exception
                ToolingFileLogger.LogWarn("Failed to publish python_execute output to the session.", ex:=ex)
            End Try
        End If

        ' Workspace (path-addressed): preserve the output subpath; fall back to Desktop when no
        ' workspace is connected. Containment is enforced before writing.
        Try
            Dim baseRoot As String = ResolveWorkspacePath("", allowRoot:=True)
            If String.IsNullOrWhiteSpace(baseRoot) Then
                baseRoot = SharedLibrary.Agents.PathPolicy.GetDefaultWritableRoot()
            End If
            baseRoot = System.IO.Path.GetFullPath(baseRoot)
            Dim targetPath As String = System.IO.Path.GetFullPath(System.IO.Path.Combine(baseRoot, subPath))
            Dim containmentPrefix As String = baseRoot.TrimEnd(System.IO.Path.DirectorySeparatorChar) & System.IO.Path.DirectorySeparatorChar
            If Not targetPath.StartsWith(containmentPrefix, StringComparison.OrdinalIgnoreCase) Then
                targetPath = System.IO.Path.Combine(baseRoot, fileName)
            End If
            Dim targetDir As String = System.IO.Path.GetDirectoryName(targetPath)
            If Not String.IsNullOrWhiteSpace(targetDir) AndAlso Not System.IO.Directory.Exists(targetDir) Then
                System.IO.Directory.CreateDirectory(targetDir)
            End If
            System.IO.File.Copy(output.FullPath, targetPath, overwrite:=True)
        Catch ex As Exception
            ToolingFileLogger.LogWarn("Failed to publish python_execute output to the workspace.", ex:=ex)
        End Try
    End Sub

    ''' <summary>
    ''' Host wrapper bridging the Outlook ToolCall/ToolResponse/ToolExecutionContext types
    ''' to the host-agnostic PythonExecuteTool core in SharedLibrary.
    ''' </summary>
    Private Async Function ExecutePythonExecuteTool(toolCall As ToolCall,
                                                    context As ToolExecutionContext,
                                                    Optional cancellationToken As CancellationToken = Nothing) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId,
            .ToolName = toolCall.ToolName,
            .OriginalCallJson = toolCall.RawJson,
            .ResultKind = Agents.PythonExecuteTool.ToolName
        }

        Dim result As Agents.PythonExecuteToolCoreResult = Nothing
        Dim cancelled As Boolean = False
        Dim unexpected As Exception = Nothing

        cancellationToken.ThrowIfCancellationRequested()

        ' Pre-execution guard: reject an unchanged deterministic resubmission before starting the worker, so a
        ' prior code-repair/diagnostic outcome that produced no code change does not spawn another identical
        ' worker run. This is not counted as a worker attempt and does not mutate the advisor session state.
        Dim pythonUnchangedRejection As String = Nothing
        If Agents.PythonExecuteRepairAdvisor.ShouldRejectUnchangedResubmission(context, toolCall.Arguments, pythonUnchangedRejection) Then
            response.Success = False
            response.ErrorCode = "UNCHANGED_RESUBMISSION_REJECTED"
            response.ErrorMessage = "The submitted Python program is unchanged from the previous deterministic failure."
            response.Response = pythonUnchangedRejection
            context.Log("Rejected unchanged Python resubmission before worker startup.", "warn")
            If _apActive OrElse _chatAgentActive Then ApDashboardLog("   ⛔ Rejected unchanged Python resubmission before worker startup.", "warn")
            ToolingFileLogger.LogRawResponseStub("Internal tool (python_execute)", response.Response)
            Return response
        End If

        ' Pre-execution guard: reject a destructive (non-minimal) repair before starting the worker. The task
        ' remains repairable so the model can submit a correct minimal repair; no worker attempt is consumed.
        Dim pythonSuspiciousRejection As String = Nothing
        If Agents.PythonExecuteRepairAdvisor.ShouldRejectSuspiciousRepair(context, toolCall.Arguments, pythonSuspiciousRejection) Then
            response.Success = False
            response.ErrorCode = "SUSPICIOUS_REPAIR_REJECTED"
            response.ErrorMessage = "The proposed change is not a minimal repair and was not executed."
            response.Response = pythonSuspiciousRejection
            context.Log("Rejected a non-minimal (destructive) Python repair before worker startup.", "warn")
            If _apActive OrElse _chatAgentActive Then ApDashboardLog("   ⛔ Rejected a non-minimal Python repair before worker startup.", "warn")
            ToolingFileLogger.LogRawResponseStub("Internal tool (python_execute)", response.Response)
            Return response
        End If

        context.Log("Running secure Python script...")

        Dim allowedOperations As New List(Of String)()
        Dim hostServiceHandler As Agents.RedInkPythonAgentDelegatingHostServiceHandler = BuildPythonHostServiceHandler(context, allowedOperations)

        Try
            result = Await Agents.PythonExecuteTool.ExecuteDetailedAsync(
                _context,
                toolCall.Arguments,
                cancellationToken,
                Sub(message)
                    context.Log(message, "step")
                    If _apActive OrElse _chatAgentActive Then ApDashboardLog("🐍 " & message, "step")
                End Sub,
                Sub(message) context.Log(message, "diag"),
                Sub(message)
                    context.Log(message, "warn")
                    If _apActive OrElse _chatAgentActive Then ApDashboardLog("   ⚠ " & message, "warn")
                End Sub,
                Sub(message) context.Log(message, "diag"),
                hostServiceHandler,
                allowedOperations)
        Catch ex As OperationCanceledException
            cancelled = True
        Catch ex As Exception
            unexpected = ex
        End Try

        If cancelled Then
            response.Success = False
            response.ErrorCode = "SESSION_CANCELLED"
            response.ErrorMessage = "Operation was cancelled."
            response.Response = Agents.PythonExecuteTool.CreateHostFailurePayload("cancelled", response.ErrorCode)
            context.Log("Python execution was cancelled.", "warn")
        ElseIf unexpected IsNot Nothing Then
            context.Log(unexpected.ToString(), "diag")
            response.Success = False
            response.ErrorCode = "INTERNAL_BROKER_ERROR"
            response.ErrorMessage = "Python execution failed."
            response.Response = Agents.PythonExecuteTool.CreateHostFailurePayload("failed", response.ErrorCode)
            context.Log("Python execution failed.", "warn")
        ElseIf result Is Nothing Then
            response.Success = False
            response.ErrorCode = "BROKER_EXITED_WITHOUT_RESPONSE"
            response.ErrorMessage = "The secure Python executor stopped unexpectedly."
            response.Response = Agents.PythonExecuteTool.CreateHostFailurePayload("failed", response.ErrorCode)
            context.Log("Python execution failed.", "warn")
        Else
            response.Response = result.Payload
            response.Success = result.Success
            response.ErrorCode = If(result.Success, String.Empty, result.ErrorCode)
            response.ErrorMessage = If(result.Success, String.Empty, result.ErrorMessage)

            ' Task-postcondition gate (distinct from worker success): a clean process exit is not accepted as a
            ' completed task when the run produced no observable/valid result (no published result and no output
            ' file, or an empty declared output file). Convert such a run into a failure so the normal repair
            ' loop annotates it and the model is required to actually produce the requested output.
            If response.Success Then
                Dim pythonIncompletePayload As String = Nothing
                If Agents.PythonExecuteRepairAdvisor.TryBuildIncompleteTaskPayload(context, toolCall.Arguments, response.Response, pythonIncompletePayload) Then
                    response.Success = False
                    response.ErrorCode = "TASK_POSTCONDITION_FAILED"
                    response.ErrorMessage = "The Python run exited successfully but did not produce a valid observable result."
                    response.Response = pythonIncompletePayload
                    context.Log("Python run exited successfully but produced no valid observable result; treating as incomplete.", "warn")
                    If _apActive OrElse _chatAgentActive Then ApDashboardLog("   ⚠ Python run produced no valid observable result; treating as incomplete.", "warn")
                End If
            End If

            ' First-failure reroute gate: when python_execute is acting as a fallback and a specialized tool that
            ' shares its capability is available, a first failure should re-route to the specialized tool rather
            ' than entering the sticky Python repair loop. Fully capability-driven and tool-agnostic. Scoped, like
            ' the redundant-success nudge, to the MAIN loop only (sub-agents/skills may use Python freely).
            Dim pythonRerouteReason As String = Nothing
            Dim pythonReroutePayload As String = Nothing
            Dim isSubAgentLoopForRoute As Boolean =
                If(context.LogPrefix, "").TrimStart().StartsWith("[subagent]", StringComparison.OrdinalIgnoreCase)
            If Not response.Success AndAlso Not isSubAgentLoopForRoute AndAlso
               Agents.PythonExecuteRepairAdvisor.TryBuildRerouteInsteadOfRepairPayload(
                   context, toolCall.Arguments, response.Response,
                   Agents.PythonExecuteRepairAdvisor.HasCapableNonFallbackAlternative(context.SelectedTools, toolCall.ToolName),
                   pythonReroutePayload, pythonRerouteReason) Then

                response.ErrorCode = "REROUTE_TO_ALTERNATIVE"
                response.ErrorMessage = "A specialized tool can perform this operation; not entering the Python repair loop."
                response.Response = pythonReroutePayload
                context.Log("python_execute failed on first fallback attempt while a specialized tool is available; advising reroute instead of repair.", "warn")
                If _apActive OrElse _chatAgentActive Then ApDashboardLog("   ⚠ Python fallback failed first attempt; advising reroute to a specialized tool.", "warn")

                If String.IsNullOrEmpty(context.PendingContinuationGuardPrompt) Then
                    context.PendingContinuationGuardPrompt = pythonRerouteReason
                    context.PendingGuardTitle = "HOST FALLBACK-REROUTE GUARD"
                    context.PendingRejectedTurnExplanation =
                        "Your previous turn used python_execute as a fallback and it failed on its first attempt, but a specialized tool for this operation is available."
                    context.PendingRejectedAssistantTurn = ""
                End If

                ToolingFileLogger.LogRawResponseStub("Internal tool (python_execute)", response.Response)
                Return response
            End If

            ' Annotate the model-facing payload with retry-vs-repair semantics and attempt history. This
            ' single wrapper serves both Outlook loops: the Local Agent (_chatAgentActive) and AutoPilot
            ' (_apActive). The ToolExecutionContext keys the per-session repair history. A terminal outcome
            ' (repair budget exhausted or non-recoverable) is flagged on the response so the tooling loop can
            ' abort via its existing tool-error abort path instead of iterating further.
            Dim pythonTerminalReason As String = Nothing
            Dim pythonRedundantReason As String = Nothing
            response.Response = Agents.PythonExecuteRepairAdvisor.Annotate(context, toolCall.Arguments, response.Response, response.Success, pythonTerminalReason, pythonRedundantReason)
            If Not String.IsNullOrEmpty(pythonTerminalReason) Then
                response.RepairLoopTerminal = True
                response.RepairLoopTerminalReason = pythonTerminalReason
                If _apActive OrElse _chatAgentActive Then ApDashboardLog("   ⛔ " & pythonTerminalReason, "warn")
            End If

            ' Soft finalization nudge for a redundant, result-less re-verification run. Scoped to the MAIN loop
            ' only (never sub-agents/skills, which may legitimately run Python repeatedly) and only when no other
            ' continuation guard is already pending. Advisory: the run stays successful and non-terminal.
            Dim isSubAgentLoop As Boolean =
                If(context.LogPrefix, "").TrimStart().StartsWith("[subagent]", StringComparison.OrdinalIgnoreCase)
            If Not isSubAgentLoop AndAlso
               Not String.IsNullOrEmpty(pythonRedundantReason) AndAlso
               String.IsNullOrEmpty(context.PendingContinuationGuardPrompt) Then
                context.PendingContinuationGuardPrompt = pythonRedundantReason
                context.PendingGuardTitle = "HOST REDUNDANT-PYTHON GUARD"
                context.PendingRejectedTurnExplanation =
                    "Your previous turn re-ran python_execute only to reproduce a result that already exists and has been validated."
                context.PendingRejectedAssistantTurn = ""
                context.Log("Redundant successful python_execute run detected; nudging finalization on the next turn.", "warn")
                If _apActive OrElse _chatAgentActive Then ApDashboardLog("   ⚠ Redundant python_execute run detected; nudging finalization.", "warn")
            End If
        End If

        ToolingFileLogger.LogRawResponseStub("Internal tool (python_execute)", response.Response)
        Return response
    End Function

    ''' <summary>
    ''' Builds the python_execute host-service handler for the current tooling loop. LLM is always
    ''' enabled. web.get is enabled only when the loop already exposes a web content retriever
    ''' (retrieve_web_content or its alias web_content_retriever). web.search is enabled only when
    ''' the loop exposes web_grounding or internet_search, preferring web_grounding when both are
    ''' present. The privacy constraint for search queries is applied only when
    ''' INI_EnablePrivacyForSearch is enabled, reusing the existing search-privacy wording.
    ''' </summary>
    Private Function BuildPythonHostServiceHandler(context As ToolExecutionContext, allowedOperations As List(Of String)) As Agents.RedInkPythonAgentDelegatingHostServiceHandler
        Dim handler As New Agents.RedInkPythonAgentDelegatingHostServiceHandler()

        ' llm.complete is always available.
        allowedOperations.Add("llm.complete")
        handler.LlmAsync =
            Async Function(systemPrompt As String, userPrompt As String, ct As System.Threading.CancellationToken) As Task(Of String)
                Dim perCallTimeoutMs As Integer = SharedLibrary.Agents.HostToolRegistration.GetPerCallLlmTimeoutMs(
                    INI_Timeout,
                    New String() {Agents.PythonExecuteTool.ToolName},
                    If(systemPrompt, "").Length,
                    If(userPrompt, "").Length)
                Return Await SharedMethods.LLM(_context, systemPrompt, userPrompt, Timeout:=perCallTimeoutMs, cancellationToken:=ct, ToolExecution:=True).ConfigureAwait(False)
            End Function

        Dim selected As List(Of ModelConfig) = If(context IsNot Nothing, context.SelectedTools, Nothing)

        If PythonLoopHasTool(selected, InternalWebToolName) OrElse PythonLoopHasTool(selected, "web_content_retriever") Then
            allowedOperations.Add("web.get")
            handler.WebGetAsync =
                Async Function(url As String, maxChars As Integer, ct As System.Threading.CancellationToken) As Task(Of String)
                    Return Await Agents.RedInkPythonAgentWebGet.RetrieveAsync(
                        url,
                        maxChars,
                        Async Function(pageUrl As String, pageMax As Integer, pageCt As System.Threading.CancellationToken) As Task(Of String)
                            Dim pageResult = Await RetrieveWebsiteContent_WebView2(pageUrl, pageMax, False, Nothing, False)
                            Return If(pageResult IsNot Nothing, pageResult.TextContent, String.Empty)
                        End Function,
                        ct).ConfigureAwait(False)
                End Function
        End If

        Dim hasGrounding As Boolean = PythonLoopHasTool(selected, "web_grounding")
        Dim hasSearch As Boolean = PythonLoopHasTool(selected, InternalSearchToolName)
        If hasGrounding OrElse hasSearch Then
            allowedOperations.Add("web.search")
            handler.WebSearchAsync =
                Async Function(query As String, maxResults As Integer, ct As System.Threading.CancellationToken) As Task(Of System.Collections.Generic.IReadOnlyList(Of Agents.RedInkPythonAgentWebSearchItem))
                    Dim text As String
                    If hasGrounding Then
                        Dim args As New Dictionary(Of String, Object) From {{"query", query}}
                        text = Await SharedLibrary.Agents.WebGroundingTool.ExecuteAsync(_context, args, ct).ConfigureAwait(False)
                    Else
                        Dim syntheticCall As New ToolCall() With {.ToolName = InternalSearchToolName, .Arguments = New Dictionary(Of String, Object) From {{"query", query}}}
                        Dim searchResponse As ToolResponse = Await ExecuteInternalSearchTool(syntheticCall, context, ct).ConfigureAwait(False)
                        text = If(searchResponse IsNot Nothing, searchResponse.Response, String.Empty)
                    End If
                    Dim results As New List(Of Agents.RedInkPythonAgentWebSearchItem)()
                    If Not String.IsNullOrWhiteSpace(text) Then
                        results.Add(New Agents.RedInkPythonAgentWebSearchItem() With {.Title = If(hasGrounding, "Web grounding result", "Internet search result"), .Url = String.Empty, .Snippet = text})
                    End If
                    Return results
                End Function
        End If

        Return handler
    End Function

    Private Shared Function PythonLoopHasTool(selected As List(Of ModelConfig), toolName As String) As Boolean
        If selected Is Nothing OrElse String.IsNullOrWhiteSpace(toolName) Then Return False
        Return selected.Any(Function(t) t IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(t.ToolName) AndAlso t.ToolName.Equals(toolName, StringComparison.OrdinalIgnoreCase))
    End Function

End Class
