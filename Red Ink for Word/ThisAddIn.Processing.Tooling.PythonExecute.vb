' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.Processing.Tooling.PythonExecute.vb
' Purpose: Integrates the shared `python_execute` tool into the Word host,
'          including configuration, output publication, and host-service
'          bridging for Word tooling execution.
'
' Architecture / How it works:
'  - Tool Registration:
'      - `TryConfigureAndBuildPythonExecuteTool()` parses `INI_PythonAgentPath`,
'        validates the configured executor, and builds the shared tool config.
'  - File Handling:
'      - Input files are resolved through `PathPolicy` into the allowed readable
'        workspace/Desktop scope.
'      - `PublishPythonAgentOutput()` writes validated outputs back through the
'        Word host's writable workspace/Desktop path policy.
'  - Tool Execution:
'      - `ExecutePythonExecuteTool()` bridges Word `ToolCall` execution to the
'        host-agnostic `PythonExecuteTool` core and returns a `ToolResponse`.
'  - Host-Mediated Capabilities:
'      - `BuildPythonHostServiceHandler()` selectively exposes LLM, `web.get`,
'        and search capabilities based on the tools active in the current loop.
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
                .InputFileResolver = Function(rel) New Agents.RedInkPythonAgentInputFile(
                                         SharedLibrary.Agents.PathPolicy.Resolve(rel, SharedLibrary.Agents.PathAccess.Read), rel),
                .PublishOutputFile = Sub(output) PublishPythonAgentOutput(output),
                .HostServiceHandler = Nothing
            })
        Catch ex As Exception
            ToolingFileLogger.LogWarn("python_execute is unavailable (configuration invalid or Python Agent below the minimum required version) and will not be advertised.", ex:=ex)
            Try
                Global.SharedLibrary.SharedLibrary.UpdateHandler.WriteUpdateLog("[ToolAvailability] python_execute unavailable; error=" & CompactPythonAvailabilityLogValue(ex.Message))
            Catch logEx As System.Exception
                System.Diagnostics.Trace.WriteLine("Could not write python_execute availability diagnostics: " & logEx.ToString())
            End Try
            Return False
        End Try

        Dim availabilityErrorCode As System.String = System.String.Empty
        If Not Agents.PythonExecuteTool.IsAvailable(availabilityErrorCode) Then
            Try
                Global.SharedLibrary.SharedLibrary.UpdateHandler.WriteUpdateLog(
                    "[ToolAvailability] python_execute unavailable; code=" & If(availabilityErrorCode, "(none)") &
                    "; path=" & CompactPythonAvailabilityLogValue(configuration.ExecutablePath))
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine("Could not write python_execute availability diagnostics: " & ex.ToString())
            End Try
            Return False
        End If

        modelConfig = Agents.PythonExecuteTool.Build(_context, toolPriority:=996, displaySuffix:=InternalToolSuffix)
        Return True
    End Function

    ''' <summary>
    ''' Writes a python_execute output artifact to the Word host's default writable root
    ''' (the connected workspace, or the Desktop when no workspace is maintained).
    ''' </summary>
    Private Shared Function CompactPythonAvailabilityLogValue(value As System.String) As System.String
        If System.String.IsNullOrWhiteSpace(value) Then Return "(none)"
        Return value.Replace(System.Environment.NewLine, " ").Replace(System.Convert.ToChar(13), " "c).Replace(System.Convert.ToChar(10), " "c).Trim()
    End Function

    Private Sub PublishPythonAgentOutput(output As Agents.RedInkPythonAgentOutput)
        If output Is Nothing OrElse String.IsNullOrWhiteSpace(output.FullPath) OrElse Not System.IO.File.Exists(output.FullPath) Then
            Return
        End If

        ' The core has already validated containment/size/hash and populated PublishedSubPath.
        Dim subPath As String = If(output.PublishedSubPath, "").Trim()
        If subPath.Length = 0 Then subPath = System.IO.Path.GetFileName(output.FullPath)
        subPath = subPath.Replace("/"c, System.IO.Path.DirectorySeparatorChar).TrimStart(System.IO.Path.DirectorySeparatorChar)
        If String.IsNullOrWhiteSpace(System.IO.Path.GetFileName(subPath)) Then Return

        Try
            Dim targetRoot As String = ""

            Try
                Dim workspace As SharedLibrary.Agents.WorkspaceState = SharedLibrary.Agents.WorkspaceTools.Active
                If workspace IsNot Nothing AndAlso
                   workspace.AllowWrite AndAlso
                   Not String.IsNullOrWhiteSpace(workspace.RootPath) AndAlso
                   System.IO.Directory.Exists(workspace.RootPath) Then

                    targetRoot = System.IO.Path.GetFullPath(workspace.RootPath)
                End If
            Catch ex As System.Exception
            End Try

            If String.IsNullOrWhiteSpace(targetRoot) Then
                Dim stagingRoot As String = SharedLibrary.Agents.PathPolicy.SessionStagingRoot
                If Not String.IsNullOrWhiteSpace(stagingRoot) AndAlso
                   System.IO.Directory.Exists(stagingRoot) Then

                    targetRoot = System.IO.Path.GetFullPath(stagingRoot)
                End If
            End If

            If String.IsNullOrWhiteSpace(targetRoot) Then
                Throw New System.InvalidOperationException(
                    "No writable Word workspace or session staging root is available for python_execute output.")
            End If

            ' Preserve output subdirectories, but contain the result strictly under the
            ' selected writable root. Without a workspace this is the per-run staging area;
            ' WordCollectAndCopyOutputs later publishes user deliverables atomically to
            ' Desktop\Inky\<timestamp> before staging cleanup.
            Dim targetPath As String = System.IO.Path.GetFullPath(System.IO.Path.Combine(targetRoot, subPath))
            Dim containmentPrefix As String = targetRoot.TrimEnd(System.IO.Path.DirectorySeparatorChar) & System.IO.Path.DirectorySeparatorChar
            If Not targetPath.Equals(targetRoot, StringComparison.OrdinalIgnoreCase) AndAlso
               Not targetPath.StartsWith(containmentPrefix, StringComparison.OrdinalIgnoreCase) Then

                targetPath = System.IO.Path.Combine(targetRoot, System.IO.Path.GetFileName(subPath))
            End If

            Dim targetDir As String = System.IO.Path.GetDirectoryName(targetPath)
            If Not String.IsNullOrWhiteSpace(targetDir) AndAlso Not System.IO.Directory.Exists(targetDir) Then
                System.IO.Directory.CreateDirectory(targetDir)
            End If
            System.IO.File.Copy(output.FullPath, targetPath, overwrite:=True)
        Catch ex As Exception
            ToolingFileLogger.LogWarn("Failed to publish python_execute output.", ex:=ex)
        End Try
    End Sub

    ''' <summary>
    ''' Host wrapper bridging the Word ToolCall/ToolResponse/ToolExecutionContext types
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
                Sub(message) context.Log(message, "step"),
                Sub(message) context.Log(message, "diag"),
                Sub(message) context.Log(message, "warn"),
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

                Dim originalPythonFailureCode As System.String = If(response.ErrorCode, System.String.Empty)
                Dim originalPythonFailureMessage As System.String = If(response.ErrorMessage, System.String.Empty)
                Dim originalPythonFailurePayload As System.String = If(response.Response, System.String.Empty)
                context.Log(
                    "python_execute original failure before fallback reroute: code=" &
                    If(originalPythonFailureCode.Length = 0, "(none)", originalPythonFailureCode) &
                    "; message=" & If(originalPythonFailureMessage.Length = 0, "(none)", originalPythonFailureMessage),
                    "diag")
                ToolingFileLogger.LogRawResponseStub("Internal tool (python_execute original failure before reroute)", originalPythonFailurePayload)

                response.ErrorCode = "REROUTE_TO_ALTERNATIVE"
                response.ErrorMessage = "A specialized tool can perform this operation; not entering the Python repair loop."
                response.Response = pythonReroutePayload
                context.Log("python_execute failed on first fallback attempt while a specialized tool is available; advising reroute instead of repair.", "warn")

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

            ' Annotate the model-facing payload with retry-vs-repair semantics and attempt history so the
            ' Word tooling loop stops guessing nonexistent APIs after deterministic Python errors. The
            ' ToolExecutionContext keys the per-session repair history. A terminal outcome (repair budget
            ' exhausted or non-recoverable) is flagged on the response so the tooling loop can abort via its
            ' existing tool-error abort path instead of iterating further.
            Dim pythonTerminalReason As String = Nothing
            Dim pythonRedundantReason As String = Nothing
            response.Response = Agents.PythonExecuteRepairAdvisor.Annotate(context, toolCall.Arguments, response.Response, response.Success, pythonTerminalReason, pythonRedundantReason)
            If Not String.IsNullOrEmpty(pythonTerminalReason) Then
                response.RepairLoopTerminal = True
                response.RepairLoopTerminalReason = pythonTerminalReason
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
            End If
        End If

        ToolingFileLogger.LogRawResponseStub("Internal tool (python_execute)", response.Response)
        Return response
    End Function

    ''' <summary>
    ''' Builds the python_execute host-service handler for the current Word tooling loop. LLM is
    ''' always enabled; web.get is enabled only when a web content retriever is selected; web.search
    ''' is enabled only when web_grounding or internet_search is selected, preferring web_grounding.
    ''' </summary>
    Private Function BuildPythonHostServiceHandler(context As ToolExecutionContext, allowedOperations As List(Of String)) As Agents.RedInkPythonAgentDelegatingHostServiceHandler
        Dim handler As New Agents.RedInkPythonAgentDelegatingHostServiceHandler()

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
                            Dim pageResult = Await RetrieveWebsiteContent_WebView2Detailed(pageUrl, pageMax, expandCollapsed:=False, includeLinks:=False, linkExtensions:=Nothing)
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
                        Dim searchResponse As ToolResponse = Await ExecuteInternalSearchTool(syntheticCall, context).ConfigureAwait(False)
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
