' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: InkyPromptBuilder.vb
' Purpose: Composes the system-prompt addendum that the MAIN tooling loop (not
'          sub-agents) prepends to its existing system prompt:
'            1. Inky.md content (central, then local; local overrides via concat).
'            2. Brief listing of active skills/agents with invocation instructions.
'
' Architecture:
'  - Sub-agents NEVER use this builder (locked decision B).
'  - Called once per main-loop initialization with selected skill/agent tool names.
'  - Returns empty string when nothing is configured or active.
'  - Includes skill author mode notice when SkillAuthorMode.IsActive.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Text

Namespace Agents

    Public NotInheritable Class InkyPromptBuilder

        Private Sub New()
        End Sub

        ''' <summary>
        ''' Builds an addendum to be prepended (or appended) to the main loop's system prompt.
        ''' Returns "" when nothing is configured / nothing is selected.
        ''' </summary>
        ''' <param name="selectedSkillToolNames">Tool names of the form "skill_&lt;name&gt;" that are active for this run.</param>
        ''' <param name="selectedAgentToolNames">Tool names of the form "agent_&lt;name&gt;" that are active for this run.</param>
        Public Shared Function Build(Optional selectedSkillToolNames As IEnumerable(Of String) = Nothing,
                                     Optional selectedAgentToolNames As IEnumerable(Of String) = Nothing,
                                     Optional isAutoPilot As System.Boolean = False) As String
            Dim sb As New StringBuilder()

            Dim inky As String = If(isAutoPilot, AgentResources.InkyMdForAutoPilot, AgentResources.InkyMd)
            If Not String.IsNullOrWhiteSpace(inky) Then
                sb.AppendLine(If(isAutoPilot, "# Project guidance (AutoPilot)", "# Project guidance (Inky.md)"))
                sb.AppendLine(inky.Trim())
                sb.AppendLine()
            End If

            Dim skillLines As New List(Of String)()
            If selectedSkillToolNames IsNot Nothing Then
                For Each toolName In selectedSkillToolNames
                    If String.IsNullOrWhiteSpace(toolName) Then Continue For
                    Dim skName = StripPrefix(toolName, "skill_")
                    Dim sk = AgentResources.FindSkill(skName)
                    If sk Is Nothing Then Continue For
                    skillLines.Add("- " & sk.Name & ": " & If(sk.Description, "").Trim())
                Next
            End If

            Dim agentLines As New List(Of String)()
            If selectedAgentToolNames IsNot Nothing Then
                For Each toolName In selectedAgentToolNames
                    If String.IsNullOrWhiteSpace(toolName) Then Continue For
                    Dim agName = StripPrefix(toolName, "agent_")
                    Dim ag = AgentResources.FindAgent(agName)
                    If ag Is Nothing Then Continue For
                    agentLines.Add("- " & ag.Name & ": " & If(ag.Description, "").Trim())
                Next
            End If

            If skillLines.Count > 0 OrElse agentLines.Count > 0 Then
                sb.AppendLine("# Available skills and agents")
                If skillLines.Count > 0 Then
                    sb.AppendLine("Skills (load their instructions with skill_use before applying them):")
                    For Each l In skillLines : sb.AppendLine(l) : Next
                End If
                If agentLines.Count > 0 Then
                    If skillLines.Count > 0 Then sb.AppendLine()
                    sb.AppendLine("Sub-agents (delegate via the matching agent_* tool; they run in an isolated context and return {summary, result}):")
                    For Each l In agentLines : sb.AppendLine(l) : Next
                End If
                sb.AppendLine()
            End If

            If SkillAuthorMode.IsActive AndAlso Not isAutoPilot Then
                sb.AppendLine()
                sb.Append(SharedLibrary.SharedMethods.Default_SP_Add_AgentLayer_AuthorMode)
            End If

            Return sb.ToString().TrimEnd()
        End Function

        Private Shared Function StripPrefix(s As String, prefix As String) As String
            If s Is Nothing Then Return ""
            If s.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) Then
                Return s.Substring(prefix.Length)
            End If
            Return s
        End Function

    End Class

End Namespace