' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: AgentToolRouter.vb
' Purpose: Centralized dispatch for the agent-layer tools so call-site wiring in
'          the existing tooling loops is a single line: TryHandleAsync(...).
'          Returns tool response string when handled, Nothing if not recognized.
'
' Tools Routed:
'  - memory_put / memory_get / memory_list / memory_delete (SessionMemory)
'  - skill_use (SkillInvokeTool)
'  - agent_<name> (SubAgentRunner for sub-agent delegation)
'  - text_* tools (TextTools for plain file operations)
'  - file_* tools (FileTools for binary-safe file/directory operations)
'  - workspace_* tools (WorkspaceTools for workspace-scoped operations)
'  - word_* / worddoc_* tools (Word document operations)
'  - js_run (WebView2 JavaScript sandbox)
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Threading
Imports System.Threading.Tasks
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace Agents

    Public NotInheritable Class AgentToolRouter

        Private Sub New()
        End Sub

        Public Const AgentToolPrefix As String = "agent_"

        ''' <summary>
        ''' Tries to handle an agent-layer tool call. Returns the tool response string when
        ''' handled, or Nothing if the tool is not in this layer.
        ''' </summary>
        Public Shared Async Function TryHandleAsync(toolName As String,
                                                    arguments As IDictionary(Of String, Object),
                                                    host As ISubAgentHost,
                                                    Optional cancellationToken As CancellationToken = Nothing,
                                                    Optional sharedContext As ISharedContext = Nothing) As Task(Of String)
            If String.IsNullOrWhiteSpace(toolName) Then Return Nothing

            If MemoryTools.IsMemoryTool(toolName) Then
                Return MemoryTools.Execute(toolName, arguments)
            End If

            If TextTools.IsTextTool(toolName) Then
                Return Await TextTools.ExecuteAsync(toolName, arguments, sharedContext, cancellationToken).ConfigureAwait(False)
            End If

            If FileTools.IsFileTool(toolName) Then
                Return FileTools.Execute(toolName, arguments)
            End If

            If WorkspaceTools.IsWorkspaceTool(toolName) Then
                Return WorkspaceTools.Execute(toolName, arguments)
            End If

            If WordTools.IsWordTool(toolName) Then
                Return WordTools.Execute(toolName, arguments)
            End If

            If WordDocTools.IsWordDocTool(toolName) Then
                Return WordDocTools.Execute(toolName, arguments)
            End If

            If JsRunTool.IsJsTool(toolName) Then
                Return Await JsRunTool.ExecuteAsync(arguments, cancellationToken).ConfigureAwait(False)
            End If

            If String.Equals(toolName, SkillInvokeTool.ToolName, StringComparison.OrdinalIgnoreCase) Then
                Return SkillInvokeTool.Execute(arguments)
            End If

            If toolName.StartsWith(AgentToolPrefix, StringComparison.OrdinalIgnoreCase) Then
                Dim agentName = toolName.Substring(AgentToolPrefix.Length)
                Dim task = GetStr(arguments, "task")
                Dim ctxBlob = GetStr(arguments, "context")
                Return Await SubAgentRunner.InvokeAsync(host, agentName, task, ctxBlob,
                                                       storeResultInMemory:=True,
                                                       cancellationToken:=cancellationToken).
                                            ConfigureAwait(False)
            End If

            Return Nothing
        End Function

        ''' <summary>True if the tool name belongs to the agent layer (memory_*, skill_use, agent_*).</summary>
        Public Shared Function IsAgentLayerTool(toolName As String) As Boolean
            If String.IsNullOrWhiteSpace(toolName) Then Return False
            If MemoryTools.IsMemoryTool(toolName) Then Return True
            If TextTools.IsTextTool(toolName) Then Return True
            If FileTools.IsFileTool(toolName) Then Return True
            If WorkspaceTools.IsWorkspaceTool(toolName) Then Return True
            If WordTools.IsWordTool(toolName) Then Return True
            If WordDocTools.IsWordDocTool(toolName) Then Return True
            If JsRunTool.IsJsTool(toolName) Then Return True
            If String.Equals(toolName, SkillInvokeTool.ToolName, StringComparison.OrdinalIgnoreCase) Then Return True
            If toolName.StartsWith(AgentToolPrefix, StringComparison.OrdinalIgnoreCase) Then Return True
            Return False
        End Function





        Private Shared Function GetStr(args As IDictionary(Of String, Object), name As String) As String
            If args Is Nothing Then Return ""
            Dim v As Object = Nothing
            If Not args.TryGetValue(name, v) OrElse v Is Nothing Then Return ""
            Return System.Convert.ToString(v)
        End Function

    End Class

End Namespace
