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
                Return Await SharedMethods.LLM(_context, systemPrompt, userPrompt, cancellationToken:=ct, ToolExecution:=True).ConfigureAwait(False)
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
