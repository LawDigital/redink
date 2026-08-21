' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ToolLoaderTool.vb
' Purpose: Internal manifest-only loader tool used to keep the first model pass
'          small. The model first sees only a compact index plus this loader.
'          When it asks for one or more tools by name, the host materializes
'          them and exposes their full definitions on the next iteration.
'
' Architecture:
'  - Lazy-loading trigger: ShouldUseLazyLoading(selectedTools) returns true
'    when tool count exceeds DefaultLazyLoadThreshold (8).
'  - Build(manifests) creates the loader config with tool index.
'  - ExtractRequestedToolNames(arguments) parses tool/tools array from request.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Collections
Imports System.Linq
Imports System.Text
Imports Newtonsoft.Json.Linq

Namespace Agents

    Public NotInheritable Class ToolLoaderTool

        Public Const LoaderToolName As String = "tool_loader"
        Public Const DefaultLazyLoadThreshold As Integer = 8

        Private Sub New()
        End Sub

        Public Shared Function ShouldUseLazyLoading(selectedTools As IEnumerable(Of SharedLibrary.ModelConfig)) As Boolean
            If selectedTools Is Nothing Then Return False

            Dim count As Integer = 0

            For Each tool In selectedTools
                If tool Is Nothing OrElse String.IsNullOrWhiteSpace(tool.ToolName) Then Continue For
                count += 1
                If count > DefaultLazyLoadThreshold Then
                    Return True
                End If
            Next

            Return False
        End Function

        Public Shared Function Build(manifests As IEnumerable(Of ToolManifest)) As SharedLibrary.ModelConfig
            Dim items = If(manifests, Enumerable.Empty(Of ToolManifest)()).
                Where(Function(m)
                          Return m IsNot Nothing AndAlso
                                 Not String.IsNullOrWhiteSpace(m.Name) AndAlso
                                 Not m.Name.Equals(LoaderToolName, StringComparison.OrdinalIgnoreCase)
                      End Function).
                OrderBy(Function(m) m.Name, StringComparer.OrdinalIgnoreCase).
                ToList()

            Dim sb As New StringBuilder()
            sb.Append("tool_loader: Loads one or more allowed tools by exact name so their full schema and instructions become available in a later iteration. ")
            sb.Append("Use this first when you decide a specific tool is needed. ")
            sb.Append("TURN RULE: A newly loaded tool is NOT callable in the same turn in which you load it; its schema only becomes available from the NEXT assistant turn onward. ")
            sb.Append("Never emit tool_loader and a call to a just-loaded tool in the same turn, and never call a tool you have not already loaded in an earlier turn. ")
            sb.Append("Plan ahead: in a single tool_loader call, load ALL tools you expect to need for the whole task (use the 'tools' array), then wait for the next turn and call them. ")
            sb.Append("If a tool is already exposed with its full schema this turn, call it directly and do NOT load it again via tool_loader. ")
            sb.Append("If you decide to use a skill or an agent, load it first; once its instructions list the tools it needs, load ALL of those required tools together in a single tool_loader call before invoking them. ")
            sb.Append("WEB TOOL ROUTING (BINDING): Use web_grounding to discover relevant public pages when the site or URL is not yet known. Use retrieve_web_content for a known, mostly static URL when readable text and ordinary links are sufficient. Load the Playwright browser tools lazily only when the task needs them. Prefer browser_open -> browser_snapshot, then browser_interact only when needed, followed by a fresh browser_snapshot when the task is to explore or scan a specific website, find links/pages/downloads on that site, inspect menus/navigation, follow pagination, or handle JavaScript/dynamically rendered content that simple retrieval may miss. If retrieve_web_content reports navigation failure, no links, or incomplete client-side content for a specific site, load browser_open, browser_snapshot and browser_interact together via tool_loader, then use browser_open + browser_snapshot next instead of repeating web_grounding. If a browser_snapshot already exposes the needed link/control, continue through the browser instead of restarting with web_grounding. ")
            sb.Append("SKILL/AGENT WORK: If the task is to create, install, modify, convert, review, or diagnose a Skill or Agent (even when the user does not name the skill-author), you MUST load and use the skill-author skill and the resource filesystem tools (file_make_dir, file_copy, text_write) writing under the resource root. Never satisfy such a task with workspace_write or by creating a folder in the temporary workspace - workspace outputs are temporary and do not install a skill. ")
            sb.Append("TOOLING-RUN DIAGNOSIS: If the task is to analyze, diagnose, debug, review, or explain what happened in a PREVIOUS tooling run (for example 'how did the last run go', 'analyze the last tool/python run', or 'did you really read the log'), treat this as a skill-author diagnostics task (its diagnostics section is the authority): you MUST load and use the skill-author skill together with text_read, then use skill_use.resource_index.diagnostics_files as the authoritative inventory of exact available diagnostics files. If that inventory is empty or absent, STOP and report that no deterministic diagnostics file inventory is available in this run. Do NOT reconstruct a run from session memory, the active document, or agent_workspace_* / memory_list searches, and do NOT use js_run or python_execute to probe the host filesystem for diagnostics files. ")
            sb.Append("Available tool index:")

            For Each item In items
                sb.AppendLine()
                sb.Append("- ").Append(item.Name)

                If Not String.IsNullOrWhiteSpace(item.Category) Then
                    sb.Append(" [").Append(item.Category.Trim()).Append("]")
                End If

                Dim shortDesc As String = Shrink(item.Description, 120)
                If shortDesc <> "" Then
                    sb.Append(": ").Append(shortDesc)
                End If
            Next

            Dim def As String =
                "{""name"":""tool_loader""," &
                """description"":""Loads one or more allowed tools by exact name so they become available for subsequent tool calls.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """tools"":{""type"":""array"",""items"":{""type"":""string""},""description"":""Exact tool names to load from the available tool index.""}," &
                """tool"":{""type"":""string"",""description"":""Single exact tool name to load (alternative to tools).""}," &
                """reason"":{""type"":""string"",""description"":""Optional short reason for loading the tool or tools.""}" &
                "},""additionalProperties"":false}}"

            Return New SharedLibrary.ModelConfig() With {
                .ToolName = LoaderToolName,
                .ToolInstructionsPrompt = sb.ToString(),
                .ToolDefinition = def,
                .ModelDescription = "Tool Loader (internal)",
                .Tool = True,
                .ToolPriority = -1000,
                .ToolErrorHandling = "skip"
            }
        End Function

        Public Shared Function ExtractRequestedToolNames(arguments As Dictionary(Of String, Object)) As List(Of String)
            Dim result As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            If arguments Is Nothing Then
                Return result.ToList()
            End If

            If arguments.ContainsKey("tool") Then
                AddNames(arguments("tool"), result)
            End If

            If arguments.ContainsKey("tools") Then
                AddNames(arguments("tools"), result)
            End If

            Return result.
                Where(Function(s) Not String.IsNullOrWhiteSpace(s)).
                Select(Function(s) s.Trim()).
                ToList()
        End Function

        Private Shared Sub AddNames(value As Object, target As HashSet(Of String))
            If value Is Nothing OrElse target Is Nothing Then Return

            If TypeOf value Is JValue Then
                AddNames(DirectCast(value, JValue).Value, target)
                Return
            End If

            If TypeOf value Is String Then
                Dim raw As String = DirectCast(value, String).Trim()
                If raw = "" Then Return

                raw = raw.Replace(";", ",").
                          Replace(vbCrLf, ",").
                          Replace(vbCr, ",").
                          Replace(vbLf, ",")

                For Each part In raw.Split({","c}, StringSplitOptions.RemoveEmptyEntries)
                    Dim name As String = part.Trim()
                    If name <> "" Then
                        target.Add(name)
                    End If
                Next

                Return
            End If

            If TypeOf value Is JArray Then
                For Each item As JToken In DirectCast(value, JArray)
                    AddNames(item, target)
                Next
                Return
            End If

            If TypeOf value Is IEnumerable Then
                For Each item As Object In DirectCast(value, IEnumerable)
                    AddNames(item, target)
                Next
                Return
            End If

            Dim fallback As String = value.ToString().Trim()
            If fallback <> "" Then
                target.Add(fallback)
            End If
        End Sub

        Private Shared Function Shrink(value As String, maxLength As Integer) As String
            Dim text As String = If(value, "").Replace(vbCr, " ").Replace(vbLf, " ").Trim()

            While text.Contains("  ")
                text = text.Replace("  ", " ")
            End While

            If text.Length <= maxLength Then
                Return text
            End If

            Return text.Substring(0, maxLength - 1).TrimEnd() & "…"
        End Function

    End Class

    ''' <summary>
    ''' Internal manifest-only routing handshake for top-level tooling runs.
    ''' It exposes only skill/agent names and descriptions; resource bodies stay lazy.
    ''' </summary>
    Public NotInheritable Class CapabilityRoutingTool

        Public Const ResolverToolName As String = "resolve_capability_route"
        Public Const KindSkill As String = "skill"
        Public Const KindAgent As String = "agent"
        Public Const KindNone As String = "none"

        Private Sub New()
        End Sub

        Public Shared Function Build(manifests As IEnumerable(Of ToolManifest)) As SharedLibrary.ModelConfig
            Dim candidates = If(manifests, Enumerable.Empty(Of ToolManifest)()).
                Where(Function(m)
                          If m Is Nothing OrElse String.IsNullOrWhiteSpace(m.Name) Then Return False
                          Return String.Equals(m.Category, KindSkill, StringComparison.OrdinalIgnoreCase) OrElse
                                 String.Equals(m.Category, KindAgent, StringComparison.OrdinalIgnoreCase)
                      End Function).
                OrderBy(Function(m)
                            If String.Equals(m.Category, KindSkill, StringComparison.OrdinalIgnoreCase) Then Return 0
                            Return 1
                        End Function).
                ThenBy(Function(m) m.Name, StringComparer.OrdinalIgnoreCase).
                ToList()

            If candidates.Count = 0 Then Return Nothing

            Dim sb As New StringBuilder()
            sb.Append("CAPABILITY ROUTING GATE (MANDATORY FIRST DECISION): Before any substantive tool work, resolve whether the user's requested workflow has a specifically applicable skill or, only if no such skill applies, a specifically applicable top-level agent. ")
            sb.Append("Skills have semantic precedence over agents because a skill represents the prescribed workflow and may itself delegate to agents. ")
            sb.Append("Choose a skill only when its description clearly matches the requested workflow/task type, not merely a broad topic. ")
            sb.Append("If no skill clearly matches, choose an agent only when its description fits the user's whole top-level task; do not choose an agent that is evidently a bounded worker/helper/criterion checker intended for delegation unless the user's entire request is exactly that bounded task. ")
            sb.Append("If neither applies, resolve kind='none'. Do not perform web/search/knowledge/compliance/research/document-analysis or other substantive tooling before this routing decision. ")
            sb.Append("Use only the metadata below; do not preload skill or agent bodies. Call resolve_capability_route exactly once with kind='skill' or kind='agent' plus the exact candidate name, or kind='none' with no name. Candidates:")

            For Each item In candidates
                sb.AppendLine()
                sb.Append("- ").Append(item.Name).Append(" [").Append(item.Category.Trim()).Append("]")
                Dim shortDesc As String = ShrinkRouting(item.Description, 700)
                If shortDesc <> "" Then sb.Append(": ").Append(shortDesc)
            Next

            Dim def As String =
                "{""name"":""resolve_capability_route""," &
                """description"":""Mandatory first routing handshake: select the specifically applicable skill, otherwise a top-level agent, otherwise none.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """kind"":{""type"":""string"",""enum"":[""skill"",""agent"",""none""],""description"":""Routing result. Skills take precedence when a specifically applicable skill exists.""}," &
                """name"":{""type"":""string"",""description"":""Exact advertised skill/agent tool name. Omit or leave empty when kind is none.""}," &
                """reason"":{""type"":""string"",""description"":""One short reason based only on the advertised names/descriptions.""}" &
                "},""required"":[""kind""],""additionalProperties"":false}}"

            Return New SharedLibrary.ModelConfig() With {
                .ToolName = ResolverToolName,
                .ToolInstructionsPrompt = sb.ToString(),
                .ToolDefinition = def,
                .ModelDescription = "Capability Router (internal)",
                .Tool = True,
                .ToolPriority = -1100,
                .ToolErrorHandling = "skip"
            }
        End Function

        Public Shared Function IsResolverToolName(toolName As String) As Boolean
            Return Not String.IsNullOrWhiteSpace(toolName) AndAlso
                   toolName.Trim().Equals(ResolverToolName, StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function ShrinkRouting(value As String, maxLength As Integer) As String
            Dim text As String = If(value, "").Replace(vbCr, " ").Replace(vbLf, " ").Trim()
            While text.Contains("  ")
                text = text.Replace("  ", " ")
            End While
            If text.Length <= maxLength Then Return text
            Return text.Substring(0, maxLength - 1).TrimEnd() & "…"
        End Function

    End Class

End Namespace
