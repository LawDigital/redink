' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.Processing.Tooling.ToolExecution.vb
' Purpose: Tool call execution dispatcher and external/MCP tool invocation.
'
' Responsibilities:
'  - Main ExecuteToolCall() dispatcher routing to internal vs. external handlers.
'  - Workspace tool execution (extract_text, read_many, staged file access).
'  - Internal tool routing (web, search, knowledge stores, M365, agents, skills).
'  - External tool execution via APICall templates.
'  - MCP SSE (Server-Sent Events) transport integration.
'  - MCP Streamable HTTP transport integration.
'  - OAuth2 token refresh for tool services.
'  - Tool parameter substitution and default resolution.
'  - Error extraction and structured failure handling.
'  - Cancellation token propagation.
'
' Architecture:
'  - Dispatch by tool name to appropriate executor (internal vs. external).
'  - Support lazy tool loading for on-demand tool initialization.
'  - Apply tool configuration and restore backup after execution.
'  - Log tool call parameters, responses, and errors.
'  - Handle structured agent failures returned as ToolResponse envelopes.
'
' External Dependencies:
'  - SharedLibrary.Agents for agent-layer tools (memory, skills, agents, workspace).
'  - SharedLibrary.M365ToolService for M365 integration.
'  - ExecuteInternalWebTool, ExecuteInternalSearchTool, ExecuteInternalKnowledgeTool (internal).
'  - ExecuteExternalTool for transport-backed tools.
'  - LLM() for tool service calls (external/MCP JSON mode).
' =============================================================================


Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods


Partial Public Class ThisAddIn


    ''' <summary>
    ''' Executes a single tool call using an internal tool implementation or an external tool configuration.
    ''' Internal tools: <c>web_content_retriever</c> and <c>internet_search</c> (when search is enabled).
    ''' </summary>
    ''' <param name="toolCall">Tool call extracted from the LLM response.</param>
    ''' <param name="toolConfig">Tool configuration selected for this call.</param>
    ''' <param name="context">Tool execution context for logging and state collection.</param>
    ''' <returns>Tool execution response.</returns>
    Public Async Function ExecuteToolCall(toolCall As ToolCall, toolConfig As ModelConfig, context As ToolExecutionContext, Optional cancellationToken As System.Threading.CancellationToken = Nothing) As Task(Of ToolResponse)


        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId,
            .ToolName = toolCall.ToolName
        }

        SharedLogger.LogAgentToolCall(_context, _context.RDV, "Word_Agent", toolCall.ToolName)

        ' Build condensed parameter summary for log window
        Dim paramSummary As String = BuildCondensedParamSummary(toolCall.Arguments)
        Dim visibleToolLabel As String = Regex.Replace(If(toolCall.ToolName, ""), "_+", " ").Trim()
        Dim suppressGenericProgressLog As Boolean =
            SharedLibrary.Agents.WebGroundingTool.IsWebGroundingTool(toolCall.ToolName)

        If visibleToolLabel = "" Then
            visibleToolLabel = "processing step"
        End If

        If Not suppressGenericProgressLog Then
            context.Log("Running " & visibleToolLabel & "...")
        End If

        context.Log($"Executing tool: {toolCall.ToolName}{paramSummary}", "diag")

        If context IsNot Nothing AndAlso context.SequencingState IsNot Nothing Then
            If toolCall.Arguments Is Nothing Then
                toolCall.Arguments = New System.Collections.Generic.Dictionary(Of String, Object)(System.StringComparer.OrdinalIgnoreCase)
            End If

            Dim restoredRetryInvariants As String = ""
            Dim retryInvariantError As String = ""
            If Not SharedLibrary.Agents.ToolCallSequencing.EnforceRetryInvariantArguments(
                toolCall.ToolName,
                toolCall.Arguments,
                context.SequencingState,
                restoredRetryInvariants,
                retryInvariantError) Then

                context.LogWarn("Retry fidelity guard rejected a tool call.", details:=retryInvariantError)
                response.Success = False
                response.ErrorMessage = retryInvariantError
                response.Response = retryInvariantError
                Return response
            End If

            If restoredRetryInvariants <> "" Then
                context.Log("Retry fidelity guard restored: " & restoredRetryInvariants, "diag")
            End If
        End If

        Try
            cancellationToken.ThrowIfCancellationRequested()

            ' ── python_execute: secure sandboxed Python execution ──
            If toolCall.ToolName.Equals(Agents.PythonExecuteTool.ToolName, StringComparison.OrdinalIgnoreCase) Then
                Return Await ExecutePythonExecuteTool(toolCall, context, cancellationToken)
            End If

            ' ── workspace_extract_text: read any supported file via GetFileContent (unified extractor) ──
            If toolCall.ToolName.Equals(SharedLibrary.Agents.WorkspaceTools.ToolExtractText, StringComparison.OrdinalIgnoreCase) Then
                Dim relPath As String = GetToolArgumentString(toolCall.Arguments, "path")
                Dim maxChars As Integer = 12000
                Dim startChar As Integer = 0
                Dim startPage As Integer = 0
                Dim endPage As Integer = 0

                Try
                    If toolCall.Arguments IsNot Nothing AndAlso toolCall.Arguments.ContainsKey("max_chars") Then
                        Integer.TryParse(toolCall.Arguments("max_chars").ToString(), maxChars)
                    End If

                    If toolCall.Arguments IsNot Nothing AndAlso toolCall.Arguments.ContainsKey("start_char") Then
                        Integer.TryParse(toolCall.Arguments("start_char").ToString(), startChar)
                    ElseIf toolCall.Arguments IsNot Nothing AndAlso toolCall.Arguments.ContainsKey("offset") Then
                        Integer.TryParse(toolCall.Arguments("offset").ToString(), startChar)
                    End If

                    If toolCall.Arguments IsNot Nothing AndAlso toolCall.Arguments.ContainsKey("start_page") Then
                        Integer.TryParse(toolCall.Arguments("start_page").ToString(), startPage)
                    End If

                    If toolCall.Arguments IsNot Nothing AndAlso toolCall.Arguments.ContainsKey("end_page") Then
                        Integer.TryParse(toolCall.Arguments("end_page").ToString(), endPage)
                    End If
                Catch
                End Try

                maxChars = Math.Min(Math.Max(maxChars, 1000), 500000)
                startChar = Math.Max(startChar, 0)

                Dim ws = SharedLibrary.Agents.WorkspaceTools.Active
                If ws Is Nothing OrElse Not ws.AllowRead OrElse String.IsNullOrWhiteSpace(ws.RootPath) Then
                    response.Success = False
                    response.ErrorMessage = "No readable workspace is connected."
                    GoTo __AfterDispatch
                End If

                Dim fullPath As String = ""
                Try
                    fullPath = SharedLibrary.Agents.PathPolicy.Resolve(relPath, SharedLibrary.Agents.PathAccess.Read)
                Catch ex As Exception
                    response.Success = False
                    response.ErrorMessage = "Invalid workspace path: " & ex.Message
                    GoTo __AfterDispatch
                End Try

                If String.IsNullOrWhiteSpace(fullPath) OrElse Not IO.File.Exists(fullPath) Then
                    response.Success = False
                    response.ErrorMessage = "Workspace file not found: " & If(relPath, "")
                    GoTo __AfterDispatch
                End If

                Dim extracted As String = ""
                Try
                    extracted = Await GetFileContent(fullPath, Silent:=True, DoOCR:=True, AskUser:=False)
                Catch ex As Exception
                    response.Success = False
                    response.ErrorMessage = "Extraction failed: " & ex.Message
                    GoTo __AfterDispatch
                End Try

                Dim fullText As String = If(extracted, "")
                Dim totalChars As Integer = fullText.Length
                Dim safeStart As Integer = Math.Min(startChar, totalChars)
                Dim remaining As Integer = Math.Max(totalChars - safeStart, 0)
                Dim takeChars As Integer = Math.Min(maxChars, remaining)
                Dim chunk As String = If(takeChars > 0, fullText.Substring(safeStart, takeChars), "")
                Dim truncated As Boolean = (safeStart + takeChars) < totalChars
                Dim nextOffset As Integer = safeStart + takeChars

                Dim payload As New JObject(
        New JProperty("path", If(relPath, "")),
        New JProperty("text", chunk),
        New JProperty("excerpt", BuildResultExcerpt(chunk, 800)),
        New JProperty("total_chars", totalChars),
        New JProperty("start_char", safeStart),
        New JProperty("returned_chars", takeChars),
        New JProperty("truncated", truncated),
        New JProperty("continuation", "If more content is needed, call workspace_extract_text again with start_char=next_offset and a suitable max_chars value."))

                If truncated Then
                    payload("next_offset") = nextOffset
                End If

                If startPage > 0 Then
                    payload("start_page") = startPage
                End If

                If endPage > 0 Then
                    payload("end_page") = endPage
                End If

                response.Success = True
                response.Response = payload.ToString(Formatting.None)
                ToolingFileLogger.LogRawResponseStub($"Internal tool ({toolCall.ToolName})", response.Response)
                GoTo __AfterDispatch
            End If

            ' ── workspace_read_many: shared UTF-8 text reader for multiple files ──
            If toolCall.ToolName.Equals(SharedLibrary.Agents.WorkspaceTools.ToolReadMany, StringComparison.OrdinalIgnoreCase) Then
                Dim ws = SharedLibrary.Agents.WorkspaceTools.Active
                If ws Is Nothing OrElse Not ws.AllowRead OrElse String.IsNullOrWhiteSpace(ws.RootPath) Then
                    response.Success = False
                    response.ErrorMessage = "No readable workspace is connected."
                    GoTo __AfterDispatch
                End If

                response.Response = SharedLibrary.Agents.WorkspaceTools.Execute(toolCall.ToolName, toolCall.Arguments)
                response.Success = True
                ToolingFileLogger.LogRawResponseStub($"Internal tool ({toolCall.ToolName})", response.Response)
                GoTo __AfterDispatch
            End If

            ' ── workspace_extract_text_many: extract text from multiple files via GetFileContent ──
            If toolCall.ToolName.Equals(SharedLibrary.Agents.WorkspaceTools.ToolExtractTextMany, StringComparison.OrdinalIgnoreCase) Then
                Dim manyMaxFiles As Integer = 20
                Dim manyMaxCharsPerFile As Integer = 100000
                Try
                    If toolCall.Arguments IsNot Nothing AndAlso toolCall.Arguments.ContainsKey("max_files") Then
                        Integer.TryParse(toolCall.Arguments("max_files").ToString(), manyMaxFiles)
                    End If
                    If toolCall.Arguments IsNot Nothing AndAlso toolCall.Arguments.ContainsKey("max_chars_per_file") Then
                        Integer.TryParse(toolCall.Arguments("max_chars_per_file").ToString(), manyMaxCharsPerFile)
                    End If
                Catch
                End Try
                manyMaxFiles = Math.Min(Math.Max(manyMaxFiles, 1), 100)
                manyMaxCharsPerFile = Math.Min(Math.Max(manyMaxCharsPerFile, 1000), 500000)

                Dim manyWs = SharedLibrary.Agents.WorkspaceTools.Active
                If manyWs Is Nothing OrElse Not manyWs.AllowRead OrElse String.IsNullOrWhiteSpace(manyWs.RootPath) Then
                    response.Success = False
                    response.ErrorMessage = "No readable workspace is connected."
                    GoTo __AfterDispatch
                End If

                Dim manyPaths As List(Of String) = GetToolArgumentStringList(toolCall.Arguments, "paths")
                Dim manyRequestedCount As Integer = manyPaths.Count
                Dim manySelected As List(Of String) = manyPaths.Take(manyMaxFiles).ToList()
                Dim manyItems As New List(Of Object)()

                For Each manyRelPath In manySelected
                    Dim manyFullPath As String = ""
                    Try
                        manyFullPath = SharedLibrary.Agents.PathPolicy.Resolve(manyRelPath, SharedLibrary.Agents.PathAccess.Read)
                    Catch ex As Exception
                        manyItems.Add(New With {Key .path = manyRelPath, Key .error = "invalid_path", Key .message = ex.Message})
                        Continue For
                    End Try

                    If String.IsNullOrWhiteSpace(manyFullPath) OrElse Not IO.File.Exists(manyFullPath) Then
                        manyItems.Add(New With {Key .path = manyRelPath, Key .error = "not_found", Key .message = "File not found."})
                        Continue For
                    End If

                    Try
                        Dim manyExtracted As String = Await GetFileContent(manyFullPath, Silent:=True, DoOCR:=True, AskUser:=False)
                        Dim manyTruncated As Boolean = False
                        If Not String.IsNullOrWhiteSpace(manyExtracted) AndAlso manyExtracted.Length > manyMaxCharsPerFile Then
                            manyExtracted = manyExtracted.Substring(0, manyMaxCharsPerFile) & Environment.NewLine & "[Truncated at " & manyMaxCharsPerFile & " characters.]"
                            manyTruncated = True
                        End If
                        manyItems.Add(New With {
                            Key .path = manyFullPath,
                            Key .truncated = manyTruncated,
                            Key .text = If(manyExtracted, "")
                        })
                    Catch ex As Exception
                        manyItems.Add(New With {Key .path = manyRelPath, Key .error = "extraction_failed", Key .message = ex.Message})
                    End Try
                Next

                response.Success = True
                response.Response = Newtonsoft.Json.JsonConvert.SerializeObject(New With {
                    Key .requested_count = manyRequestedCount,
                    Key .processed_count = manySelected.Count,
                    Key .items = manyItems
                })
                ToolingFileLogger.LogRawResponseStub($"Internal tool ({toolCall.ToolName})", response.Response)
                GoTo __AfterDispatch
            End If

            If SharedLibrary.Agents.WorkspaceTools.IsWorkspaceTool(toolCall.ToolName) Then
                response.Response = SharedLibrary.Agents.WorkspaceTools.Execute(toolCall.ToolName, toolCall.Arguments)
                response.Success = True

                Try
                    Dim wsToken As JToken = JToken.Parse(response.Response)
                    If wsToken.Type = JTokenType.Object Then
                        Dim wsObj = DirectCast(wsToken, JObject)
                        Dim errToken = wsObj("error")
                        If errToken IsNot Nothing AndAlso errToken.Type <> JTokenType.Null AndAlso errToken.ToString().Trim() <> "" Then
                            response.Success = False
                            response.ErrorMessage = If(wsObj("message")?.ToString(), errToken.ToString())
                        End If
                    End If
                Catch
                End Try

                ToolingFileLogger.LogRawResponseStub($"Internal tool ({toolCall.ToolName})", response.Response)
                GoTo __AfterDispatch
            End If

            If toolCall.ToolName.Equals(SharedLibrary.Agents.ToolLoaderTool.LoaderToolName, StringComparison.OrdinalIgnoreCase) Then
                response = ExecuteToolLoaderCall(toolCall, context)
                ToolingFileLogger.LogRawResponseStub($"Internal tool ({toolCall.ToolName})", response.Response)
                GoTo __AfterDispatch
            End If

            If SharedLibrary.Agents.ToolDescribeTool.IsDescribeTool(toolCall.ToolName) Then
                Dim describeRegistry As SharedLibrary.Agents.ToolRegistry =
                    If(context.AuthoritativeToolRegistrySnapshot,
                       If(context.AuthoritativeToolRegistry, context.AllowedToolRegistry))
                response.Response = SharedLibrary.Agents.ToolDescribeTool.Execute(toolCall.Arguments, describeRegistry)
                response.Success = Not String.IsNullOrWhiteSpace(response.Response)
                ToolingFileLogger.LogRawResponseStub($"Internal tool ({toolCall.ToolName})", response.Response)
                GoTo __AfterDispatch
            End If

            If SharedLibrary.Agents.ContextExpandTool.IsContextExpandTool(toolCall.ToolName) Then
                response.Response = SharedLibrary.Agents.ContextExpandTool.Execute(toolCall.Arguments)
                response.Success = Not String.IsNullOrWhiteSpace(response.Response)
                ToolingFileLogger.LogRawResponseStub($"Internal tool ({toolCall.ToolName})", response.Response)
                GoTo __AfterDispatch
            End If

            If SharedLibrary.Agents.AskUserTool.IsAskUserTool(toolCall.ToolName) Then
                response.Response = SharedLibrary.Agents.AskUserTool.Execute(toolCall.Arguments)
                response.Success = Not String.IsNullOrWhiteSpace(response.Response)
                ToolingFileLogger.LogRawResponseStub($"Internal tool ({toolCall.ToolName})", response.Response)
                GoTo __AfterDispatch
            End If

            If SharedLibrary.Agents.ContextCompactTool.IsContextCompactTool(toolCall.ToolName) Then
                response.Response = SharedLibrary.Agents.ContextCompactTool.Execute(
                    toolCall.Arguments,
                    SharedLibrary.Agents.WorkflowContinuity.CurrentWorkflowId)
                response.Success = Not String.IsNullOrWhiteSpace(response.Response)
                ToolingFileLogger.LogRawResponseStub($"Internal tool ({toolCall.ToolName})", response.Response)
                GoTo __AfterDispatch
            End If

            ' Agent layer (memory_*, skill_use, agent_*) — single-line dispatcher.
            If SharedLibrary.Agents.AgentToolRouter.IsAgentLayerTool(toolCall.ToolName) Then
                cancellationToken.ThrowIfCancellationRequested()
                Dim __agentJson = Await SharedLibrary.Agents.AgentToolRouter.TryHandleAsync(
        toolCall.ToolName, toolCall.Arguments, CType(Me, SharedLibrary.Agents.ISubAgentHost), cancellationToken, _context).ConfigureAwait(False)

                response.Response = If(__agentJson, "")
                response.Success = Not String.IsNullOrWhiteSpace(response.Response)

                ApplyStructuredAgentResult(response, context)

                If Not response.Success AndAlso String.IsNullOrWhiteSpace(response.ErrorMessage) Then
                    response.ErrorMessage = "Agent-layer tool returned no usable result."
                ElseIf response.Success AndAlso SharedLibrary.Agents.JsRunTool.IsJsTool(toolCall.ToolName) AndAlso IsEmptyJsRunResult(response.Response) Then
                    response.Success = False
                    response.ResultKind = "error"
                    response.ErrorCode = "agent_empty_result"
                    response.ErrorMessage = "js_run returned no usable result. Ensure the script explicitly returns the computed value."
                    response.Response = "{""summary"":""Sub-agent returned no usable result."",""result"":null,""resultKind"":""error"",""error"":{""code"":""agent_empty_result"",""phase"":""final_output_parse"",""message"":""Sub-agent returned no usable final result.""}}"
                End If

                If String.Equals(toolCall.ToolName, SharedLibrary.Agents.SkillInvokeTool.ToolName, StringComparison.OrdinalIgnoreCase) AndAlso response.Success Then
                    Dim __skillName As String = Nothing
                    If toolCall.Arguments IsNot Nothing AndAlso toolCall.Arguments.ContainsKey("name") Then
                        __skillName = System.Convert.ToString(toolCall.Arguments("name"))
                    End If
                    If Not String.IsNullOrWhiteSpace(__skillName) Then
                        ToolingFileLogger.SetTopLevelSkillName(__skillName)
                    End If
                End If

                ToolingFileLogger.LogSubAgentReturn($"Agent-layer tool ({toolCall.ToolName})", response.Response)
                GoTo __AfterDispatch
            ElseIf toolCall.ToolName.StartsWith("skill_", StringComparison.OrdinalIgnoreCase) Then
                Dim skillArgs As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)

                skillArgs("name") = toolCall.ToolName.Substring("skill_".Length)

                If toolCall.Arguments IsNot Nothing Then
                    For Each kvp In toolCall.Arguments
                        If Not skillArgs.ContainsKey(kvp.Key) Then
                            skillArgs(kvp.Key) = kvp.Value
                        End If
                    Next

                    If Not skillArgs.ContainsKey("input") AndAlso toolCall.Arguments.ContainsKey("instruction") Then
                        skillArgs("input") = toolCall.Arguments("instruction")
                    End If
                End If

                response.Response = SharedLibrary.Agents.SkillInvokeTool.Execute(skillArgs)
                response.Success = Not String.IsNullOrWhiteSpace(response.Response)

                ApplyStructuredAgentResult(response, context)

                If response.Success Then
                    LoadSkillAllowedToolsFromResponse(response.Response, context)
                ElseIf String.IsNullOrWhiteSpace(response.ErrorMessage) Then
                    response.ErrorMessage = "Skill invocation returned no usable result."
                End If

                ToolingFileLogger.SetTopLevelSkillName(toolCall.ToolName.Substring("skill_".Length))
                ToolingFileLogger.LogSubAgentReturn($"Agent-layer skill ({toolCall.ToolName})", response.Response)
                GoTo __AfterDispatch
            End If


            If toolCall.ToolName.Equals(InternalWebToolName, StringComparison.OrdinalIgnoreCase) Then
                response = Await ExecuteInternalWebTool(toolCall, context)
                ToolingFileLogger.LogRawResponseStub($"Internal tool ({toolCall.ToolName})", response.Response)

            ElseIf toolCall.ToolName.Equals(InternalDownloadWebFilesToolName, StringComparison.OrdinalIgnoreCase) Then
                response = Await ExecuteInternalDownloadWebFilesTool(toolCall, context)
                ToolingFileLogger.LogRawResponseStub($"Internal tool ({toolCall.ToolName})", response.Response)

            ElseIf SharedLibrary.Agents.WebGroundingTool.IsWebGroundingTool(toolCall.ToolName) Then
                response = Await ExecuteWebGroundingTool(toolCall, context, cancellationToken)
                ToolingFileLogger.LogRawResponseStub($"Internal tool ({toolCall.ToolName})", response.Response)

            ElseIf toolCall.ToolName.Equals(InternalSearchToolName, StringComparison.OrdinalIgnoreCase) Then
                response = Await ExecuteInternalSearchTool(toolCall, context)
                ToolingFileLogger.LogRawResponseStub($"Internal tool ({toolCall.ToolName})", response.Response)

            ElseIf IsInternalKnowledgeToolName(toolCall.ToolName) Then
                response = Await ExecuteInternalKnowledgeTool(toolCall, context)
                ToolingFileLogger.LogRawResponseStub($"Internal tool ({toolCall.ToolName})", response.Response)

            ElseIf SharedLibrary.SharedLibrary.M365ToolService.IsM365ToolName(toolCall.ToolName) Then
                response = Await ExecuteInternalM365Tool(toolCall, context)
                ToolingFileLogger.LogRawResponseStub($"Internal tool ({toolCall.ToolName})", response.Response)

            Else
                response = Await ExecuteExternalTool(toolCall, toolConfig, context)
                ToolingFileLogger.LogRawResponseStub($"Tool LLM() ({toolCall.ToolName})", response.Response)
            End If

__AfterDispatch:

            ' Log completion with excerpt
            If response.Success Then
                If Not suppressGenericProgressLog Then
                    context.Log("Finished " & visibleToolLabel & ".", "success")
                End If

                context.Log($"Tool {toolCall.ToolName} completed: {BuildResultExcerpt(response.Response, 160)}", "diag")
            Else
                context.LogError(
                        $"Tool {toolCall.ToolName} failed: {If(response.ErrorMessage, "Tool failed.")}",
                        details:=$"CallId={toolCall.CallId}; Result={BuildResultExcerpt(If(response.Response, ""), 200)}",
                        userMessage:="A processing step could not be completed.")
            End If

        Catch ex As OperationCanceledException
            response.Success = False
            response.ErrorMessage = "Operation was cancelled"
            context.Log($"Tool {toolCall.ToolName} cancelled")
            ToolingFileLogger.LogWarn($"Tool {toolCall.ToolName} cancelled.")

        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            context.Log($"Tool {toolCall.ToolName} error: {ex.Message}")
            ToolingFileLogger.LogError($"Tool {toolCall.ToolName} execution error.", ex:=ex)
        End Try

        Return response
    End Function




    ''' <summary>
    ''' Executes an external tool by applying its <see cref="ModelConfig"/> to <c>_context</c>, preparing
    ''' the tool API call payload, and invoking <c>LLM</c> in JSON response mode.
    ''' </summary>
    ''' <param name="toolCall">Tool call extracted from the LLM response.</param>
    ''' <param name="toolConfig">Tool configuration to apply for this call.</param>
    ''' <param name="context">Tool execution context for logging and diagnostics.</param>
    ''' <returns>Tool response containing the tool service result or an error.</returns>
    Private Async Function ExecuteExternalTool(toolCall As ToolCall, toolConfig As ModelConfig, context As ToolExecutionContext) As Task(Of ToolResponse)
        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId,
            .ToolName = toolCall.ToolName
        }

        Try
            Dim apiCallTemplate = toolConfig.ToolAPICall
            If String.IsNullOrWhiteSpace(apiCallTemplate) Then
                apiCallTemplate = toolConfig.APICall
            End If

            If String.IsNullOrWhiteSpace(apiCallTemplate) Then
                ' Advertised but no transport. Return a STRUCTURED error so the orchestrator
                ' can force a fallback attempt instead of letting the model declare 'blocked'.
                response.Success = False
                response.ResultKind = "error"
                response.ErrorCode = "tool_not_executable"
                response.ErrorMessage =
                "The tool '" & If(toolCall.ToolName, "") &
                "' is advertised but not executable on this host (no transport template). " &
                "Choose a different tool. If the user authorized a fallback, attempt it before declaring blocked."
                response.Response = Agents.ToolExecutorRegistry.BuildNotExecutablePayload(
                toolCall.ToolName,
                If(context IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(context.HostKind),
                   context.HostKind,
                   "Word"))
                ToolingFileLogger.LogError(
                "Tool advertised but not executable on this host.",
                details:=$"ToolName='{toolCall.ToolName}'")
                Return response
            End If

            Dim apiCall = apiCallTemplate
            Dim parameterSchemas = GetToolParameterSchemas(toolConfig)
            Dim requiredParameters = GetToolRequiredParameters(toolConfig)

            For Each kvp In toolCall.Arguments
                Dim placeholder = "{" & kvp.Key & "}"
                Dim schemaToken As JToken = Nothing
                parameterSchemas.TryGetValue(kvp.Key, schemaToken)

                Dim value As String
                If kvp.Value Is Nothing Then
                    value = ""
                ElseIf TypeOf kvp.Value Is JToken Then
                    Dim jt = DirectCast(kvp.Value, JToken)
                    If jt.Type = JTokenType.String Then
                        value = FormatToolValueForPlaceholder(jt.Value(Of String)(), schemaToken)
                    Else
                        value = jt.ToString(Formatting.None)
                    End If
                Else
                    value = FormatToolValueForPlaceholder(kvp.Value.ToString(), schemaToken)
                End If

                apiCall = apiCall.Replace(placeholder, value)
            Next

            Dim unreplacedPattern As New Regex("\{([a-zA-Z_][a-zA-Z0-9_]*)\}")
            Dim unreplacedMatches = unreplacedPattern.Matches(apiCall)

            If unreplacedMatches.Count > 0 Then
                ToolingFileLogger.LogWarn(
                    "Unreplaced placeholders found in tool APICall (defaults will be applied if available).",
                    details:=$"ToolName='{toolCall.ToolName}'; Count={unreplacedMatches.Count}; APICall='{apiCall}'")

                Dim toolDefaults As Dictionary(Of String, String) = Nothing
                If Not String.IsNullOrWhiteSpace(toolConfig.ToolParameterDefaults) Then
                    Try
                        toolDefaults = JsonConvert.DeserializeObject(Of Dictionary(Of String, String))(toolConfig.ToolParameterDefaults)
                    Catch ex As Exception
                        ToolingFileLogger.LogWarn(
                            "ToolParameterDefaults parse failed.",
                            details:=$"ToolName='{toolCall.ToolName}'; ToolParameterDefaults='{toolConfig.ToolParameterDefaults}'",
                            ex:=ex)
                    End Try
                End If

                For Each m As Match In unreplacedMatches
                    Dim placeholderName = m.Groups(1).Value
                    Dim schemaToken As JToken = Nothing
                    parameterSchemas.TryGetValue(placeholderName, schemaToken)

                    Dim shouldRemoveProperty As Boolean = False
                    Dim replacement As String = ResolveToolDefaultValue(
                        placeholderName,
                        toolDefaults,
                        schemaToken,
                        requiredParameters.Contains(placeholderName),
                        shouldRemoveProperty)

                    If shouldRemoveProperty Then
                        apiCall = RemoveToolArgumentPlaceholderProperty(apiCall, placeholderName)
                    Else
                        apiCall = apiCall.Replace(m.Value, replacement)
                    End If
                Next

                Dim remainingMatches = unreplacedPattern.Matches(apiCall)
                If remainingMatches.Count > 0 Then
                    Dim remainingNames = remainingMatches.
                        Cast(Of Match)().
                        Select(Function(m) m.Groups(1).Value).
                        Distinct(StringComparer.OrdinalIgnoreCase).
                        ToList()

                    response.Success = False
                    response.ErrorMessage = $"Unreplaced placeholders remain after applying defaults: {String.Join(", ", remainingNames)}"

                    ToolingFileLogger.LogError(
                        "Unreplaced placeholders remain in tool APICall after applying defaults.",
                        details:=$"ToolName='{toolCall.ToolName}'; RemainingCount={remainingMatches.Count}; Remaining='{String.Join(", ", remainingNames)}'; APICall='{apiCall}'")

                    Return response
                End If
            End If

            If toolConfig.OAuth2 Then
                toolConfig.DecodedAPI = Await SharedMethods.GetFreshAccessToken(
                    _context,
                    toolConfig.OAuth2ClientMail,
                    toolConfig.OAuth2Scopes,
                    toolConfig.APIKey,
                    toolConfig.OAuth2Endpoint,
                    toolConfig.OAuth2ATExpiry,
                    True,
                    False).ConfigureAwait(False)

                If String.IsNullOrWhiteSpace(toolConfig.DecodedAPI) Then
                    response.Success = False
                    response.ErrorMessage = "OAuth2 authentication failed."
                    ToolingFileLogger.LogError(
                        "OAuth2 authentication failed before MCP tool execution.",
                        details:=$"ToolName='{toolCall.ToolName}'")
                    Return response
                End If
            End If

            If Not String.IsNullOrWhiteSpace(toolConfig.Endpoint) AndAlso
               toolConfig.Endpoint.StartsWith(SharedMethods.MCP_SSE_PREFIX, StringComparison.OrdinalIgnoreCase) Then

                Dim sseBase = toolConfig.Endpoint.Substring(SharedMethods.MCP_SSE_PREFIX.Length)
                Dim resolvedHeaderB = If(toolConfig.HeaderB, "").Replace("{apikey}", If(toolConfig.DecodedAPI, ""))

                context.Log($"SSE transport: executing tool {toolCall.ToolName} via {sseBase}")
                ToolingFileLogger.LogStep($"SSE round-trip for {toolCall.ToolName} at {sseBase}")
                ToolingFileLogger.LogStep($"SSE request body: {apiCall}")

                Dim sseAttemptedRefresh As Boolean = False
                Dim sseEx As Exception = Nothing
                Dim sseDone As Boolean = False

                Do
                    sseEx = Nothing
                    resolvedHeaderB = If(toolConfig.HeaderB, "").Replace("{apikey}", If(toolConfig.DecodedAPI, ""))

                    Try
                        Dim rawResult = Await SharedMethods.ExecuteMCPSSEToolCall(
                            _context,
                            sseBase, apiCall,
                            If(toolConfig.HeaderA, ""), resolvedHeaderB,
                            CInt(Math.Min(If(toolConfig.Timeout > 0, toolConfig.Timeout, 60000L), Integer.MaxValue)))

                        ToolingFileLogger.LogRawResponseStub($"SSE tool result ({toolCall.ToolName})", rawResult)

                        If Not String.IsNullOrWhiteSpace(rawResult) Then
                            Dim toolErrorMessage As String = ""
                            response.Response = rawResult

                            If TryExtractToolServiceErrorMessage(rawResult, toolErrorMessage) Then
                                response.Success = False
                                response.ErrorMessage = toolErrorMessage
                                ToolingFileLogger.LogWarn(
                                    "SSE tool service returned a logical error.",
                                    details:=$"ToolName='{toolCall.ToolName}'; Error='{toolErrorMessage}'")
                            Else
                                response.Success = True
                            End If
                        Else
                            response.Success = False
                            response.ErrorMessage = "Empty response from SSE tool service"
                            ToolingFileLogger.LogError("Empty SSE response.", details:=$"ToolName='{toolCall.ToolName}'")
                        End If

                        sseDone = True

                    Catch ex As Exception
                        sseEx = ex
                    End Try

                    If sseDone Then Exit Do
                    If sseAttemptedRefresh Then Exit Do
                    If Not ShouldRetryMCPAfterUnauthorized(toolConfig, sseEx) Then Exit Do

                    sseAttemptedRefresh = True
                    ToolingFileLogger.LogWarn(
                        "SSE tool call returned Unauthorized. Forcing MCP OAuth refresh and retrying once.",
                        details:=$"ToolName='{toolCall.ToolName}'; SseBase='{sseBase}'")

                    Dim sseRefreshOk As Boolean = Await ForceRefreshToolOAuthToken(toolConfig, toolCall.ToolName).ConfigureAwait(False)
                    If Not sseRefreshOk Then Exit Do
                Loop

                If Not sseDone AndAlso sseEx IsNot Nothing Then
                    response.Success = False
                    response.ErrorMessage = $"SSE tool call failed: {sseEx.Message}"
                    ToolingFileLogger.LogError("SSE tool call failed.",
                        details:=$"ToolName='{toolCall.ToolName}'; SseBase='{sseBase}'", ex:=sseEx)
                End If

                Return response
            End If

            If IsMCPStreamableToolCall(toolConfig.Endpoint, apiCall) Then
                Dim mcpUrl As String = If(toolConfig.Endpoint, "")
                If mcpUrl.StartsWith(SharedMethods.MCP_STREAMABLE_PREFIX, StringComparison.OrdinalIgnoreCase) Then
                    mcpUrl = mcpUrl.Substring(SharedMethods.MCP_STREAMABLE_PREFIX.Length)
                End If

                Dim resolvedHeaderB = If(toolConfig.HeaderB, "").Replace("{apikey}", If(toolConfig.DecodedAPI, ""))

                context.Log($"MCP Streamable HTTP: executing tool {toolCall.ToolName} via {mcpUrl}")
                ToolingFileLogger.LogStep($"MCP Streamable HTTP round-trip for {toolCall.ToolName} at {mcpUrl}")
                ToolingFileLogger.LogStep($"MCP Streamable HTTP request body: {apiCall}")

                Dim streamAttemptedRefresh As Boolean = False
                Dim streamEx As Exception = Nothing
                Dim streamDone As Boolean = False

                Do
                    streamEx = Nothing
                    resolvedHeaderB = If(toolConfig.HeaderB, "").Replace("{apikey}", If(toolConfig.DecodedAPI, ""))

                    Try
                        Dim rawResult = Await SharedMethods.ExecuteMCPStreamableToolCall(
                            mcpUrl,
                            apiCall,
                            If(toolConfig.HeaderA, ""),
                            resolvedHeaderB,
                            CInt(Math.Min(If(toolConfig.Timeout > 0, toolConfig.Timeout, 60000L), Integer.MaxValue)))

                        ToolingFileLogger.LogRawResponseStub($"MCP Streamable HTTP tool result ({toolCall.ToolName})", rawResult)

                        If Not String.IsNullOrWhiteSpace(rawResult) Then
                            Dim toolErrorMessage As String = ""
                            response.Response = rawResult

                            If TryExtractToolServiceErrorMessage(rawResult, toolErrorMessage) Then
                                response.Success = False
                                response.ErrorMessage = toolErrorMessage
                                ToolingFileLogger.LogWarn(
                                    "MCP Streamable HTTP tool service returned a logical error.",
                                    details:=$"ToolName='{toolCall.ToolName}'; Error='{toolErrorMessage}'")
                            Else
                                response.Success = True
                            End If
                        Else
                            response.Success = False
                            response.ErrorMessage = "Empty response from MCP Streamable HTTP tool service"
                            ToolingFileLogger.LogError(
                                "Empty MCP Streamable HTTP response.",
                                details:=$"ToolName='{toolCall.ToolName}'; Endpoint='{mcpUrl}'")
                        End If

                        streamDone = True

                    Catch ex As Exception
                        streamEx = ex
                    End Try

                    If streamDone Then Exit Do
                    If streamAttemptedRefresh Then Exit Do
                    If Not ShouldRetryMCPAfterUnauthorized(toolConfig, streamEx) Then Exit Do

                    streamAttemptedRefresh = True
                    ToolingFileLogger.LogWarn(
                        "MCP Streamable HTTP tool call returned Unauthorized. Forcing MCP OAuth refresh and retrying once.",
                        details:=$"ToolName='{toolCall.ToolName}'; Endpoint='{mcpUrl}'")

                    Dim streamRefreshOk As Boolean = Await ForceRefreshToolOAuthToken(toolConfig, toolCall.ToolName).ConfigureAwait(False)
                    If Not streamRefreshOk Then Exit Do
                Loop

                If Not streamDone AndAlso streamEx IsNot Nothing Then
                    response.Success = False
                    response.ErrorMessage = $"MCP Streamable HTTP tool call failed: {streamEx.Message}"
                    ToolingFileLogger.LogError(
                        "MCP Streamable HTTP tool call failed.",
                        details:=$"ToolName='{toolCall.ToolName}'; Endpoint='{mcpUrl}'",
                        ex:=streamEx)
                End If

                Return response
            End If

            Dim backupConfig = GetCurrentConfig(_context)

            Try
                Dim errorFlag As Boolean = False
                ApplyModelConfig(_context, toolConfig, errorFlag)
                If errorFlag Then
                    response.Success = False
                    response.ErrorMessage = "Failed to apply tool configuration"
                    ToolingFileLogger.LogError("Failed to apply tool configuration.", details:=$"ToolName='{toolCall.ToolName}'")
                    Return response
                End If

                _context.INI_APICall_2 = apiCall

                Dim originalResponse = _context.INI_Response_2
                _context.INI_Response_2 = "JSON"

                context.Log($"Calling external service for tool: {toolCall.ToolName}")

                ToolingFileLogger.LogPreToolLlmCallSnapshot(_context)

                Dim configuredToolTimeoutMs As Long = If(_context.INI_Timeout_2 > 0, _context.INI_Timeout_2, 60000)
                Dim perCallToolTimeoutMs As Integer = SharedLibrary.Agents.HostToolRegistration.GetPerCallLlmTimeoutMs(
                    configuredToolTimeoutMs,
                    New String() {toolCall.ToolName},
                    If(_context.INI_APICall_ToolInstructions_2, "").Length,
                    If(_context.INI_APICall_ToolResponses_2, "").Length)
                If perCallToolTimeoutMs <> configuredToolTimeoutMs Then
                    context.Log($"[PERF] Tool LLM timeout elevated: tool={toolCall.ToolName}; baseMs={configuredToolTimeoutMs}; effectiveMs={perCallToolTimeoutMs}.", "diag")
                End If

                Dim result = Await LLM("", "", "", "", perCallToolTimeoutMs, True, True)

                ToolingFileLogger.LogRawResponseStub($"Tool LLM() result ({toolCall.ToolName})", result)

                _context.INI_Response_2 = originalResponse

                If Not String.IsNullOrWhiteSpace(result) Then
                    Dim toolErrorMessage As String = ""

                    response.Response = result

                    If TryExtractToolServiceErrorMessage(result, toolErrorMessage) Then
                        response.Success = False
                        response.ErrorMessage = toolErrorMessage
                        ToolingFileLogger.LogWarn(
                            "Tool service returned a logical error.",
                            details:=$"ToolName='{toolCall.ToolName}'; Error='{toolErrorMessage}'")
                    Else
                        response.Success = True
                    End If
                Else
                    response.Success = False
                    response.ErrorMessage = "Empty response from tool service"
                    ToolingFileLogger.LogError("Empty response from tool service.", details:=$"ToolName='{toolCall.ToolName}'; APICall='{apiCall}'")
                End If

            Finally
                RestoreDefaults(_context, backupConfig)
            End Try

        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            ToolingFileLogger.LogError("Tool execution error.", details:=$"ToolName='{toolCall.ToolName}'", ex:=ex)
        End Try

        Return response
    End Function




End Class
