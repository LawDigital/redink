' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: SkillInvokeTool.vb
' Purpose: Implements the universal "skill_use" tool. The model calls
'          skill_use(name, input?) and receives the SKILL.md body (loaded lazily)
'          along with an inventory of the skill's scripts/ and references/ dirs.
'
' Architecture:
'  - Lazy-loads skill bodies on first access via AgentResources.FindSkill().
'  - Returns skill instructions + inventory (names + sizes) as JSON.
'  - Model follows those instructions in subsequent turns.
'  - Text and script bodies are NOT auto-loaded; model fetches what it needs.
'  - Binary reference/script assets are discovered by inventory and can be
'    materialized with file_* tools when the loaded skill allows them.
'  - Security: allowed-tools communicated; enforcement by host runner.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Agents

    Public NotInheritable Class SkillInvokeTool

        Private Sub New()
        End Sub

        Public Const ToolName As String = "skill_use"

        Public Shared Function Build() As SharedLibrary.ModelConfig
            Dim def =
                "{""name"":""" & ToolName & """," &
                """description"":""Load and apply a Skill (Claude-style SKILL.md). Returns the skill's instructions and an inventory of its scripts/ and references/ files. Read text files with text_read, materialize binary reference or script assets with the appropriate file_* tools when allowed, and execute scripts with js_run. Use this when a relevant skill is offered above and the user's task matches."",""parameters"":{" &
                """type"":""object""," &
                """properties"":{" &
                """name"":{""type"":""string"",""description"":""The skill name (matches the Skill listed above).""}," &
                """input"":{""type"":""string"",""description"":""Optional input or sub-task description for the skill.""}}," &
                """required"":[""name""]}}"

            Return New SharedLibrary.ModelConfig() With {
                .ToolName = ToolName,
                .ToolDefinition = def,
                .ToolInstructionsPrompt = ToolName & ": Load a Skill's instructions (lazy). Call this once per skill, then follow its directions in subsequent turns, using text_read for text resources and the appropriate file_* tools for binary reference assets when the skill allows them.",
                .ModelDescription = "Skill loader",
                .Tool = True,
                .ToolPriority = 940,
                .ToolErrorHandling = "skip"
            }
        End Function

        ''' <summary>
        ''' Executes the skill_use call. Returns a JSON string suitable for the tool response.
        ''' Caller passes the dictionary from ToolCall.Arguments.
        ''' </summary>
        Public Shared Function Execute(arguments As IDictionary(Of String, Object)) As String
            Try
                Dim name As String = GetStr(arguments, "name")
                Dim input As String = GetStr(arguments, "input")

                If String.IsNullOrWhiteSpace(name) Then
                    name = GetStr(arguments, "tool")
                End If

                If String.IsNullOrWhiteSpace(name) Then
                    name = GetStr(arguments, "skill")
                End If

                If String.IsNullOrWhiteSpace(name) Then
                    Return JsonConvert.SerializeObject(New With {Key .error = "missing_name"})
                End If

                ' Canonical, agnostic skill resolution: try the name exactly as provided
                ' first, then fall back to a single "skill_" prefix strip only if that
                ' actually resolves. This avoids double-stripping when the tool name
                ' (e.g. "skill_<slug>") has already been reduced to the skill's own name,
                ' which may itself legitimately start with "skill_".
                Dim sk = AgentResources.FindSkill(name)

                If sk Is Nothing AndAlso name.StartsWith("skill_", StringComparison.OrdinalIgnoreCase) Then
                    Dim strippedName As String = name.Substring("skill_".Length)
                    If Not String.IsNullOrWhiteSpace(strippedName) Then
                        Dim strippedSkill = AgentResources.FindSkill(strippedName)
                        If strippedSkill IsNot Nothing Then
                            name = strippedName
                            sk = strippedSkill
                        End If
                    End If
                End If

                If sk Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {Key .error = "skill_not_found", Key .name = name})
                End If

                Dim body As String = sk.LoadBody()
                Dim scripts As List(Of Object) = InventoryDir(sk.ScriptsDir)
                Dim references As List(Of Object) = InventoryDir(sk.ReferencesDir)

                Dim result As New JObject()
                result("name") = sk.Name
                result("description") = If(sk.Description, "")
                result("origin") = If(sk.IsLocal, "local", "central")
                result("dir") = sk.DirectoryPath
                result("network_allowed") = sk.Network
                result("allowed_tools") = If(sk.AllowedTools Is Nothing,
                                             New JArray(),
                                             JArray.FromObject(sk.AllowedTools))
                result("instructions") = body
                result("scripts") = JArray.FromObject(scripts)
                result("references") = JArray.FromObject(references)

                ' Provide a discovery index of all skills and agents with their exact
                ' file paths. Without this, an authoring skill has to guess where a
                ' resource lives, which leads to failed reads and accidental new files.
                result("resource_index") = BuildResourceIndex()

                If Not String.IsNullOrWhiteSpace(input) Then result("input") = input

                Return result.ToString(Formatting.None)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {Key .error = "skill_invoke_failed", Key .message = ex.Message})
            End Try
        End Function

        Private Shared Function BuildResourceIndex() As JObject
            Dim idx As New JObject()

            Dim skillsArr As New JArray()
            Try
                For Each s In AgentResources.Skills
                    If s Is Nothing Then Continue For
                    Dim o As New JObject()
                    o("name") = If(s.Name, "")
                    o("origin") = If(s.IsLocal, "local", "central")
                    o("file") = If(s.FilePath, "")
                    o("dir") = If(s.DirectoryPath, "")
                    skillsArr.Add(o)
                Next
            Catch
            End Try

            Dim agentsArr As New JArray()
            Try
                For Each a In AgentResources.Agents
                    If a Is Nothing Then Continue For
                    Dim o As New JObject()
                    o("name") = If(a.Name, "")
                    o("origin") = If(a.IsLocal, "local", "central")
                    o("file") = If(a.FilePath, "")
                    o("dir") = If(a.DirectoryPath, "")
                    agentsArr.Add(o)
                Next
            Catch
            End Try

            idx("skills") = skillsArr
            idx("agents") = agentsArr

            Dim localRoot As String = If(AgentResources.ConfiguredLocalPath, "")
            Dim centralRoot As String = If(AgentResources.ConfiguredCentralPath, "")
            Dim authorActive As Boolean = SkillAuthorMode.IsActive
            Dim allowCentral As Boolean = authorActive AndAlso SkillAuthorMode.AllowCentralWrites

            idx("local_root") = localRoot
            idx("central_root") = centralRoot
            idx("author_mode_active") = authorActive
            idx("local_writes_allowed") = authorActive
            idx("central_writes_allowed") = allowCentral

            ' Deterministic target root for NEW resources: local is writable only while author mode is
            ' active; central is only writable when it was additionally and explicitly enabled.
            Dim newResourceRoot As String =
                If(allowCentral AndAlso Not String.IsNullOrWhiteSpace(centralRoot), centralRoot, localRoot)
            idx("new_resource_root") = newResourceRoot

            idx("create_hint") =
                "ALWAYS write NEW resources under new_resource_root using ABSOLUTE paths — never a relative path, " &
                "because a relative path resolves into the temporary workspace, not the resource tree. " &
                "New skill: new_resource_root + '\skills\<name>\SKILL.md'. " &
                "New agent: new_resource_root + '\agents\<name>\AGENT.md' (or new_resource_root + '\agents\<name>.md'). " &
                "Parent folders are created automatically."

            idx("root_choice_hint") =
                If(String.IsNullOrWhiteSpace(centralRoot),
                   "Only a local resource root is configured; create and edit everything under local_root.",
                   If(allowCentral,
                      "Both a local and a central root are configured and central writing is ENABLED. " &
                      "Create NEW shared resources under central_root; create user-private resources under local_root. " &
                      "When unsure, prefer local_root.",
                      "Both a local and a central root are configured but central writing is DISABLED. " &
                      "Create ALL new resources under local_root. Never write under central_root."))

            idx("note") =
                If(authorActive, "", "Author mode is OFF: skills and agents are READ-ONLY; do not attempt to create or edit any resource files. ") &
                "To modify an EXISTING resource, edit the exact 'file' path shown for it (each entry carries its 'origin' = local or central). " &
                "Do not invent new paths for existing resources, and do not copy a central resource into local_root unless the user asks to fork it."

            Return idx
        End Function

        Private Shared Function InventoryDir(dir As String) As List(Of Object)
            Dim list As New List(Of Object)
            If String.IsNullOrWhiteSpace(dir) OrElse Not Directory.Exists(dir) Then Return list
            Try
                For Each f In Directory.EnumerateFiles(dir, "*", SearchOption.AllDirectories)
                    Try
                        Dim fi As New FileInfo(f)
                        Dim rel = f.Substring(dir.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                        list.Add(New With {
                            Key .path = rel,
                            Key .size = fi.Length
                        })
                    Catch
                    End Try
                Next
            Catch
            End Try
            Return list
        End Function

        Private Shared Function GetStr(args As IDictionary(Of String, Object), name As String) As String
            If args Is Nothing Then Return ""
            Dim v As Object = Nothing
            If Not args.TryGetValue(name, v) OrElse v Is Nothing Then Return ""
            Return System.Convert.ToString(v)
        End Function

    End Class

End Namespace
