' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ISubAgentHost.vb
' Purpose: Contract by which SharedLibrary asks the host add-in (Word or Outlook)
'          to run an isolated tooling-loop pass for a sub-agent. The host owns
'          the loop implementation; this interface only carries the inputs the
'          sub-agent needs and returns the final assistant text.
'
' The host implementation is expected to:
'   - Start a *clean* message history (no parent system prompt, no parent turns).
'   - Use SubAgentRunRequest.SystemPrompt as the only system prompt
'     (this already includes the AGENT.md body composed by SubAgentRunner).
'   - Use SubAgentRunRequest.UserMessage as the single user turn.
'   - Restrict tool availability to SubAgentRunRequest.AllowedToolNames plus any
'     SubAgentRunRequest.OptionalToolNames that exist in the authoritative registry.
'     Missing required tools block the run; missing optional tools are ignored.
'   - Honor SubAgentRunRequest.SpecialModelKey by resolving the corresponding
'     special-task-model and using it for this run only (restoring the previous
'     model afterwards), falling back to "agentdefaultmodel".
'   - Return the final assistant text (the runner will try to parse it as JSON).
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Threading
Imports System.Threading.Tasks

Namespace Agents

    Public Class SubAgentRunRequest
        Public Property AgentName As String
        Public Property SystemPrompt As String
        Public Property UserMessage As String
        Public Property SpecialModelKey As String
        Public Property AllowedToolNames As IReadOnlyList(Of String)
        Public Property OptionalToolNames As IReadOnlyList(Of String)
        Public Property MaxIterations As Integer
        Public Property TimeoutSeconds As Integer
        Public Property WorkflowId As String

        ''' <summary>
        ''' Opaque caller-supplied identity of one logical sub-agent task.
        ''' The same logical task must reuse the same id. No semantic inference is used.
        ''' </summary>
        Public Property SubAgentTaskId As String

        ''' <summary>
        ''' Zero-based retry index controlled exclusively by SubAgentRunner.
        ''' 0 is the initial isolated run; 1 is the single permitted internal
        ''' empty-response recovery run. This is runtime state, not model input.
        ''' </summary>
        Public Property RunnerRetryIndex As Integer

        ''' <summary>
        ''' Exact JSON array supplied by the parent agent_<name> invocation.
        ''' [] means the delegated task is explicitly non-file-producing.
        ''' The nested tooling run locks this contract and may not broaden it.
        ''' </summary>
        Public Property ExpectedArtifactsJson As String = Nothing
    End Class

    Public Interface ISubAgentHost
        Function RunIsolatedToolingLoopAsync(request As SubAgentRunRequest,
                                             cancellationToken As CancellationToken) As Task(Of String)
    End Interface

End Namespace
