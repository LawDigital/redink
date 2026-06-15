' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ToolRegistryBuilder.vb
' Purpose: Convenience helpers that fill a ToolRegistry from the data sources
'          existing tooling loops already produce (List(Of ModelConfig)) and from
'          discovered Skills/Agents. Lazy registration where possible.
'
' Builders:
'  - FromModelConfigs(...): wraps existing list of tools into registry (eager).
'  - AddModelConfigs(...): adds tools to an existing registry.
'  - AddSkills(...): registers discovered skills as lazy skill_<name> entries.
'  - AddAgents(...): registers discovered agents as lazy agent_<name> entries.
'  - This file is additive; does not alter existing call sites.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO

Namespace Agents

    Public NotInheritable Class ToolRegistryBuilder

        Private Sub New()
        End Sub

        Private Shared Function ToSafeToolSuffix(value As String) As String
            If String.IsNullOrWhiteSpace(value) Then
                Return ""
            End If

            Dim sb As New System.Text.StringBuilder()
            Dim lastWasUnderscore As Boolean = False

            For Each ch As Char In value.Trim()
                If Char.IsLetterOrDigit(ch) Then
                    sb.Append(Char.ToLowerInvariant(ch))
                    lastWasUnderscore = False
                ElseIf Not lastWasUnderscore Then
                    sb.Append("_"c)
                    lastWasUnderscore = True
                End If
            Next

            Return sb.ToString().Trim("_"c)
        End Function

        ''' <summary>
        ''' Wraps an existing list of <see cref="ModelConfig"/> tools into a registry.
        ''' Entries are registered eagerly (the ModelConfig is already built and cheap to
        ''' keep). Use this to mirror the current tooling-loop tool lists without changing
        ''' callers.
        ''' </summary>
        Public Shared Function FromModelConfigs(tools As IEnumerable(Of SharedLibrary.ModelConfig),
                                                Optional category As String = "builtin") As ToolRegistry
            Dim reg As New ToolRegistry()
            If tools Is Nothing Then Return reg
            For Each mc In tools
                If mc Is Nothing Then Continue For
                If Not mc.Tool Then Continue For
                If String.IsNullOrWhiteSpace(mc.ToolName) Then Continue For
                reg.RegisterEager(mc, category:=category)
            Next
            Return reg
        End Function

        ''' <summary>
        ''' Adds an existing list of <see cref="ModelConfig"/> tools to an already-built registry.
        ''' </summary>
        Public Shared Sub AddModelConfigs(target As ToolRegistry,
                                          tools As IEnumerable(Of SharedLibrary.ModelConfig),
                                          Optional category As String = "builtin")
            If target Is Nothing OrElse tools Is Nothing Then Return
            For Each mc In tools
                If mc Is Nothing OrElse Not mc.Tool Then Continue For
                If String.IsNullOrWhiteSpace(mc.ToolName) Then Continue For
                target.RegisterEager(mc, category:=category)
            Next
        End Sub

        ''' <summary>
        ''' Registers all discovered skills as LAZY entries under the tool name
        ''' <c>skill.&lt;skillName&gt;</c>. The skill body, scripts and references
        ''' are NOT loaded here; they will be loaded on first invocation by
        ''' SkillInvokeTool (step 6).
        ''' </summary>
        Public Shared Sub AddSkills(target As ToolRegistry,
                                    skills As IEnumerable(Of SkillDescriptor))
            If target Is Nothing OrElse skills Is Nothing Then Return
            For Each sk In skills
                If sk Is Nothing OrElse String.IsNullOrWhiteSpace(sk.Name) Then Continue For

                Dim localSk = sk ' capture
                Dim toolSuffix As String = ToSafeToolSuffix(localSk.Name)
                If String.IsNullOrWhiteSpace(toolSuffix) Then Continue For

                Dim toolName As String = "skill_" & toolSuffix
                Dim shortDesc As String = If(localSk.Description, "")
                Dim originLabel As String = If(localSk.IsLocal, "local", "central")

                target.RegisterLazy(
                    New ToolManifest With {
                        .Name = toolName,
                        .Description = shortDesc,
                        .Category = "skill",
                        .Source = localSk.FilePath
                    },
                    Function() BuildSkillStubConfig(localSk, toolName, originLabel))
            Next
        End Sub

        ''' <summary>
        ''' Registers all discovered agents as LAZY entries under the tool name
        ''' <c>agent.&lt;agentName&gt;</c>. The agent body and model are not bound
        ''' here; SubAgentRunner (step 5) will resolve them when invoked.
        ''' </summary>
        Public Shared Sub AddAgents(target As ToolRegistry,
                                    agents As IEnumerable(Of AgentDescriptor))
            If target Is Nothing OrElse agents Is Nothing Then Return
            For Each ag In agents
                If ag Is Nothing OrElse String.IsNullOrWhiteSpace(ag.Name) Then Continue For

                Dim localAg = ag ' capture
                Dim toolSuffix As String = ToSafeToolSuffix(localAg.Name)
                If String.IsNullOrWhiteSpace(toolSuffix) Then Continue For

                Dim toolName As String = "agent_" & toolSuffix
                Dim shortDesc As String = If(localAg.Description, "")
                Dim originLabel As String = If(localAg.IsLocal, "local", "central")

                target.RegisterLazy(
                    New ToolManifest With {
                        .Name = toolName,
                        .Description = shortDesc,
                        .Category = "agent",
                        .Source = localAg.FilePath
                    },
                    Function() BuildAgentStubConfig(localAg, toolName, originLabel))
            Next
        End Sub

        ' --------------------------------------------------------------- stub builders

        ''' <summary>
        ''' Builds a placeholder <see cref="ModelConfig"/> describing a skill as a callable
        ''' tool. The real execution wiring is provided in step 6 (skill.use). For now this
        ''' allows the registry to surface skills in the UI selection list and to be passed
        ''' through the existing tooling-loop plumbing once skills are wired in.
        ''' </summary>
        Private Shared Function BuildSkillStubConfig(sk As SkillDescriptor,
                                                     toolName As String,
                                                     originLabel As String) As SharedLibrary.ModelConfig
            Dim instr As New System.Text.StringBuilder()
            instr.Append(toolName).Append(": ")
            If Not String.IsNullOrWhiteSpace(sk.Description) Then
                instr.Append(sk.Description.Trim())
            Else
                instr.Append("Invokes the '").Append(sk.Name).Append("' skill.")
            End If
            instr.Append(" (Skill, ").Append(originLabel).Append(".)")

            Dim def As String =
                "{""name"":""" & JsonEscape(toolName) & """," &
                """description"":""" & JsonEscape(If(sk.Description, "Invokes the skill.")) & """," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """input"":{""type"":""string"",""description"":""Task or input for the skill.""}}," &
                """required"":[""input""]}}"

            Return New SharedLibrary.ModelConfig() With {
                .ToolName = toolName,
                .ToolInstructionsPrompt = instr.ToString(),
                .ToolDefinition = def,
                .ModelDescription = "Skill: " & sk.Name & " (" & originLabel & ")",
                .Tool = True,
                .ToolPriority = 500,
                .ToolErrorHandling = "skip"
            }
        End Function

        ''' <summary>
        ''' Builds a placeholder <see cref="ModelConfig"/> describing a sub-agent as a callable
        ''' tool (the future <c>agent.delegate</c> path). The real isolated-context execution
        ''' is implemented in step 5 (SubAgentRunner).
        ''' </summary>
        Private Shared Function BuildAgentStubConfig(ag As AgentDescriptor,
                                                     toolName As String,
                                                     originLabel As String) As SharedLibrary.ModelConfig
            Dim instr As New System.Text.StringBuilder()
            instr.Append(toolName).Append(": ")
            If Not String.IsNullOrWhiteSpace(ag.Description) Then
                instr.Append(ag.Description.Trim())
            Else
                instr.Append("Delegates a task to the '").Append(ag.Name).Append("' sub-agent.")
            End If
            instr.Append(" Returns a JSON object {summary, result}. ")
            instr.Append("(Agent, ").Append(originLabel).Append(
                If(String.IsNullOrWhiteSpace(ag.Model), ".)", ", model='" & ag.Model & "'.)"))

            Dim def As String =
                "{""name"":""" & JsonEscape(toolName) & """," &
                """description"":""" & JsonEscape(If(ag.Description, "Delegates a task to a sub-agent.")) & """," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """task"":{""type"":""string"",""description"":""Self-contained task description for the sub-agent.""}," &
                """context"":{""type"":""string"",""description"":""Optional context blob (text) the parent wants to pass through. Defaults to none.""}}," &
                """required"":[""task""]}}"

            Return New SharedLibrary.ModelConfig() With {
                .ToolName = toolName,
                .ToolInstructionsPrompt = instr.ToString(),
                .ToolDefinition = def,
                .ModelDescription = "Agent: " & ag.Name & " (" & originLabel & ")",
                .Tool = True,
                .ToolPriority = 400,
                .ToolErrorHandling = "skip"
            }
        End Function

        Private Shared Function JsonEscape(s As String) As String
            If s Is Nothing Then Return ""
            Dim sb As New System.Text.StringBuilder(s.Length + 8)
            For Each c In s
                Select Case c
                    Case ChrW(34) : sb.Append("\""")
                    Case "\"c : sb.Append("\\")
                    Case ChrW(8) : sb.Append("\b")
                    Case ChrW(9) : sb.Append("\t")
                    Case ChrW(10) : sb.Append("\n")
                    Case ChrW(12) : sb.Append("\f")
                    Case ChrW(13) : sb.Append("\r")
                    Case Else
                        If AscW(c) < 32 Then
                            sb.Append("\u").Append(AscW(c).ToString("X4"))
                        Else
                            sb.Append(c)
                        End If
                End Select
            Next
            Return sb.ToString()
        End Function

    End Class

End Namespace
