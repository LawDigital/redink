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
            Return False
        End Try

        Return Agents.PythonExecuteTool.TryBuild(_context, modelConfig, toolPriority:=996, displaySuffix:=InternalToolSuffix)
    End Function

    ''' <summary>
    ''' Writes a python_execute output artifact to the Word host's default writable root
    ''' (the connected workspace, or the Desktop when no workspace is maintained).
    ''' </summary>
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
            ' PathPolicy.Resolve enforces workspace/Desktop containment; the subpath preserves
            ' output subdirectories so distinct outputs sharing a filename do not collide.
            Dim targetPath As String = SharedLibrary.Agents.PathPolicy.Resolve(subPath, SharedLibrary.Agents.PathAccess.Write)
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

            ' Annotate the model-facing payload with retry-vs-repair semantics and attempt history so the
            ' Word tooling loop stops guessing nonexistent APIs after deterministic Python errors. The
            ' ToolExecutionContext keys the per-session repair history. A terminal outcome (repair budget
            ' exhausted or non-recoverable) is flagged on the response so the tooling loop can abort via its
            ' existing tool-error abort path instead of iterating further.
            Dim pythonTerminalReason As String = Nothing
            response.Response = Agents.PythonExecuteRepairAdvisor.Annotate(context, toolCall.Arguments, response.Response, response.Success, pythonTerminalReason)
            If Not String.IsNullOrEmpty(pythonTerminalReason) Then
                response.RepairLoopTerminal = True
                response.RepairLoopTerminalReason = pythonTerminalReason
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
