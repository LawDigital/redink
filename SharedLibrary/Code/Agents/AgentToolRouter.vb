' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: AgentToolRouter.vb
' Purpose:
'   Host-neutral dispatch boundary for shared agent/tooling capabilities. It accepts
'   an already-authorized tool name plus normalized arguments and delegates execution
'   to the owning shared tool implementation, returning Nothing only when the tool is
'   outside this shared layer.
'
' Architecture / Function:
'   - Routes memory, text/file/workspace, skill/sub-agent, Word-document, JavaScript,
'     browser, context, Python and related shared tools without duplicating host logic.
'   - Relies on PathPolicy, host interfaces (ISubAgentHost/IWordDocumentHost) and the
'     individual tool classes for capability checks, containment and execution rules.
'   - Is an execution router, not a discovery/authorization source: ToolRegistry,
'     ToolRegistryBuilder and HostToolRegistration determine what may be exposed.
'   - Keeps Word/Outlook/Excel tooling loops symmetrical by centralizing common dispatch.
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

            If BrowserTools.IsBrowserTool(toolName) Then
                Return Await BrowserTools.ExecuteAsync(toolName, arguments, cancellationToken, sharedContext).ConfigureAwait(False)
            End If

            If JsRunTool.IsJsTool(toolName) Then
                Return Await JsRunTool.ExecuteAsync(arguments, sharedContext, cancellationToken).ConfigureAwait(False)
            End If

            If String.Equals(toolName, SkillInvokeTool.ToolName, StringComparison.OrdinalIgnoreCase) Then
                Return SkillInvokeTool.Execute(arguments)
            End If

            If toolName.StartsWith(AgentToolPrefix, StringComparison.OrdinalIgnoreCase) Then
                Dim agentName As String =
                    toolName.Substring(AgentToolPrefix.Length)

                Dim task As String =
                    GetStr(arguments, "task")

                If String.IsNullOrWhiteSpace(task) Then
                    Return "{""summary"":""Sub-agent invocation rejected.""," &
                           """result"":null," &
                           """resultKind"":""error""," &
                           """error"":{""code"":""missing_task""," &
                           """phase"":""agent_router_validation""," &
                           """message"":""Every agent_<name> invocation requires a non-empty task.""}}"
                End If

                Dim ctxBlob As String =
                    GetStr(arguments, "context")

                Dim subAgentTaskId As String =
                    GetStr(arguments, "subagent_task_id").Trim()

                If subAgentTaskId = "" Then
                    Return "{""summary"":""Sub-agent invocation rejected.""," &
                           """result"":null," &
                           """resultKind"":""error""," &
                           """error"":{""code"":""missing_subagent_task_id""," &
                           """phase"":""agent_router_validation""," &
                           """message"":""Every agent_<name> invocation requires an explicit opaque subagent_task_id.""}}"
                End If

                Dim expectedArtifactsRaw As Object = Nothing
                If arguments Is Nothing OrElse
                   Not arguments.TryGetValue("expected_artifacts", expectedArtifactsRaw) OrElse
                   expectedArtifactsRaw Is Nothing Then

                    Return "{""summary"":""Sub-agent invocation rejected.""," &
                           """result"":null," &
                           """resultKind"":""error""," &
                           """error"":{""code"":""missing_expected_artifacts""," &
                           """phase"":""agent_router_validation""," &
                           """message"":""Every agent_<name> invocation requires expected_artifacts. Use [] for an explicitly non-file-producing delegated task.""}}"
                End If

                Dim expectedArtifactsToken As Newtonsoft.Json.Linq.JToken = Nothing
                Try
                    expectedArtifactsToken = Newtonsoft.Json.Linq.JToken.FromObject(expectedArtifactsRaw)
                Catch
                End Try

                If expectedArtifactsToken Is Nothing OrElse
                   expectedArtifactsToken.Type <> Newtonsoft.Json.Linq.JTokenType.Array Then

                    Return "{""summary"":""Sub-agent invocation rejected.""," &
                           """result"":null," &
                           """resultKind"":""error""," &
                           """error"":{""code"":""invalid_expected_artifacts""," &
                           """phase"":""agent_router_validation""," &
                           """message"":""expected_artifacts must be a JSON array; use [] for no expected final files.""}}"
                End If

                For Each expectedArtifactToken As Newtonsoft.Json.Linq.JToken In
                    DirectCast(expectedArtifactsToken, Newtonsoft.Json.Linq.JArray)

                    Dim expectedArtifactObject As Newtonsoft.Json.Linq.JObject =
                        TryCast(expectedArtifactToken, Newtonsoft.Json.Linq.JObject)

                    If expectedArtifactObject Is Nothing Then
                        Return "{""summary"":""Sub-agent invocation rejected.""," &
                               """result"":null," &
                               """resultKind"":""error""," &
                               """error"":{""code"":""invalid_expected_artifacts""," &
                               """phase"":""agent_router_validation""," &
                               """message"":""Each expected_artifacts item must be an object containing explicit logical_deliverable_id and output_slot_id values.""}}"
                    End If

                    Dim logicalDeliverableId As String =
                        If(expectedArtifactObject.Value(Of String)("logical_deliverable_id"), "").Trim()

                    Dim outputSlotId As String =
                        If(expectedArtifactObject.Value(Of String)("output_slot_id"), "").Trim()

                    If logicalDeliverableId = "" OrElse outputSlotId = "" Then
                        Return "{""summary"":""Sub-agent invocation rejected.""," &
                               """result"":null," &
                               """resultKind"":""error""," &
                               """error"":{""code"":""invalid_expected_artifacts""," &
                               """phase"":""agent_router_validation""," &
                               """message"":""Each expected_artifacts item requires non-empty opaque logical_deliverable_id and output_slot_id values.""}}"
                    End If
                Next

                Dim expectedArtifactsJson As String =
                    expectedArtifactsToken.ToString(Newtonsoft.Json.Formatting.None)

                Return Await SubAgentRunner.InvokeAsync(
                    host,
                    agentName,
                    task,
                    ctxBlob,
                    storeResultInMemory:=True,
                    subAgentTaskId:=subAgentTaskId,
                    cancellationToken:=cancellationToken,
                    expectedArtifactsJson:=expectedArtifactsJson).
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
            If BrowserTools.IsBrowserTool(toolName) Then Return True
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
