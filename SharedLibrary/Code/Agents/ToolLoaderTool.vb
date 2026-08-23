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
'  - Build(manifests) creates the loader config with tool index and binding routing
'    guidance, including source-format preservation before artifact creation.
'  - Bootstrap preflight also classifies whether a user-supplied artifact is the
'    authoritative format/layout carrier so hosts can validate creator routing.
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

            Dim hasProcessWordDocument As System.Boolean =
                items.Any(Function(m) m.Name.Equals("process_word_document", System.StringComparison.OrdinalIgnoreCase))
            Dim hasPdfToWord As System.Boolean =
                items.Any(Function(m) m.Name.Equals("pdf_to_word", System.StringComparison.OrdinalIgnoreCase))

            If hasProcessWordDocument Then
                sb.Append("SOURCE FORMAT / EXISTING ARTIFACT ROUTING (BINDING): When the user identifies an attachment or other supplied artifact as the formatting, layout, design, master, style, or structural model for the requested output, treat that source as a FORMAT CARRIER, not merely as content. This remains true when the user says create, new, template, boilerplate, sample, generic, or similar wording. A user-selected source format carrier takes precedence over any implicit/default DESIGN REPOSITORY design or template. Turning an existing Word, PowerPoint, or Excel artifact into a generic/template version by replacing substantive content with placeholders or generic text is a TRANSFORMATION and should preserve the existing native structure; use process_word_document or the appropriate native mutation tools rather than rebuilding with create_word_document, create_powerpoint, or create_excel_spreadsheet. ")
                If hasPdfToWord Then
                    sb.Append("When a PDF is the requested format/layout carrier for a Word deliverable, use pdf_to_word FIRST and then transform the resulting DOCX with process_word_document or native Word tools. Do not choose extract_pdf_text + create_word_document for a format-preservation request merely because text extraction is easier; that path loses the source layout. Use OCR/text reconstruction only as a controlled fallback when structure-preserving conversion cannot provide a usable editable source, and do not claim exact format preservation after such a fallback. ")
                End If
                sb.Append("If a creator genuinely remains necessary while the user-supplied source is the intended visual/format authority, set use_repository_default_design=false unless the user explicitly requested a particular repository design. Do not silently mix an implicit repository default with a user-supplied format carrier. ")
            End If

            sb.Append("WEB TOOL ROUTING (BINDING): Use web_grounding to discover relevant public pages when the site or URL is not yet known. Use retrieve_web_content for a known, mostly static URL when readable text and ordinary links are sufficient. Load the Playwright browser tools lazily only when the task needs them. Prefer browser_open -> browser_snapshot, then browser_interact only when needed, followed by a fresh browser_snapshot when the task is to explore or scan a specific website, find links/pages/downloads on that site, inspect menus/navigation, follow pagination, or handle JavaScript/dynamically rendered content that simple retrieval may miss. If retrieve_web_content reports navigation failure, no links, or incomplete client-side content for a specific site, load browser_open, browser_snapshot and browser_interact together via tool_loader, then use browser_open + browser_snapshot next instead of repeating web_grounding. If a browser_snapshot already exposes the needed link/control, continue through the browser instead of restarting with web_grounding. ")
            sb.Append("POWERPOINT TEMPLATE ROUTING: If the user requests a named PowerPoint design shown in the DESIGN REPOSITORY, load create_powerpoint and pass that exact design_name. If the user explicitly refers to a .potx/.pptx stored in USER HOME FILES and manage_user_files is advertised, load manage_user_files and create_powerpoint together, resolve the requested file with manage_user_files action='use', and pass the resulting loaded filename as template_attachment_name before claiming that the template is unavailable. Do not guess layout numbers or corporate design details. ")
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


    ''' <summary>
    ''' One-call bootstrap classifier for top-level tooling runs. It combines the
    ''' response-language decision, memory-grounding classification, capability
    ''' routing, and the first safe lazy-load decision. The first bootstrap load is
    ''' intentionally limited to the selected skill/agent capability; ordinary tools
    ''' remain dynamically loadable after the selected capability has been entered.
    ''' </summary>
    Public NotInheritable Class ToolingBootstrapPreflight

        Public NotInheritable Class Decision
            Public Property JsonValid As Boolean
            Public Property Language As String = ""
            Public Property LanguageValid As Boolean
            Public Property MemoryDecision As ToolCallSequencing.MemoryGroundingIntentDecision =
                New ToolCallSequencing.MemoryGroundingIntentDecision()
            Public Property MemoryValid As Boolean
            Public Property RouteKind As String = CapabilityRoutingTool.KindNone
            Public Property RouteName As String = ""
            Public Property RouteReason As String = ""
            Public Property BootstrapLoad As String = ""
            Public Property SourceFormatAuthority As System.Boolean
            Public Property SourceFormatAuthorityReason As System.String = ""
            Public Property SourceFormatAuthorityValid As System.Boolean
            Public Property SourceFormatAuthorityApplied As System.Boolean
            Public Property RoutingSyntaxValid As Boolean
            Public Property LanguageApplied As Boolean
            Public Property MemoryApplied As Boolean
            Public Property RouteApplied As Boolean
            Public Property NormalizedOutput As String = ""
            Public Property ParseError As String = ""
        End Class

        Private Sub New()
        End Sub

        Public Shared Function BuildSystemPrompt(manifests As IEnumerable(Of ToolManifest)) As String
            Dim candidates = If(manifests, Enumerable.Empty(Of ToolManifest)()).
                Where(Function(m)
                          If m Is Nothing OrElse String.IsNullOrWhiteSpace(m.Name) Then Return False
                          Return String.Equals(m.Category, CapabilityRoutingTool.KindSkill, StringComparison.OrdinalIgnoreCase) OrElse
                                 String.Equals(m.Category, CapabilityRoutingTool.KindAgent, StringComparison.OrdinalIgnoreCase)
                      End Function).
                OrderBy(Function(m)
                            If String.Equals(m.Category, CapabilityRoutingTool.KindSkill, StringComparison.OrdinalIgnoreCase) Then Return 0
                            Return 1
                        End Function).
                ThenBy(Function(m) m.Name, StringComparer.OrdinalIgnoreCase).
                ToList()

            Dim sb As New StringBuilder()
            sb.Append("BOOTSTRAP PREFLIGHT: make five bounded decisions for the latest user request in one pass: ")
            sb.Append("(1) response language, (2) session-memory grounding mode, (3) capability route, (4) the first safe bootstrap load, and (5) whether a user-supplied artifact is explicitly authoritative for output formatting/layout/structure. ")
            sb.Append("Do not perform the substantive task. Do not use external knowledge. Treat <LATEST_USER_REQUEST_RAW> as authoritative. ")
            sb.Append("Return EXACTLY one raw JSON object and nothing else, with exactly these fields: ")
            sb.Append("{""language"":""de-CH"",""memoryGroundingMode"":""none"",""memoryReason"":""short reason"",""shouldExposeRecentMemoryStubs"":false,""explicitStoredMemoryRequired"":false,""routeKind"":""none"",""routeName"":"""",""routeReason"":""short reason"",""bootstrapLoad"":"""",""sourceFormatAuthority"":false,""sourceFormatReason"":""short reason""}. ")
            sb.Append("LANGUAGE: identify the language in which the assistant should answer the latest request; prefer a BCP-47 tag when clear. ")
            sb.Append("MEMORY: memoryGroundingMode MUST be exactly one of none, optional, or required. Use required ONLY when the latest request explicitly requires stored Memory, remembered stored content, prior saved results, or previous saved workflow outputs. If stored Memory may help but is not explicitly demanded, use optional. New self-contained tasks normally use none. Set explicitStoredMemoryRequired=true only for an explicit demand. Base this decision on semantic meaning, not language-specific keywords. ")
            sb.Append("ROUTING: routeKind MUST be exactly one of skill, agent, or none. A specifically applicable workflow skill has semantic precedence. Choose a skill only when its description clearly matches the workflow/task type, not merely the topic. Only if no skill specifically applies may you select a top-level agent whose description fits the whole task. Do not select bounded worker/helper agents for broader tasks. Otherwise choose none. ")
            sb.Append("BOOTSTRAP LOAD: when routeKind is skill or agent, bootstrapLoad MUST equal routeName exactly. When routeKind is none, bootstrapLoad MUST be empty. Do not propose ordinary tools here; they remain available through normal lazy loading after bootstrap. ")
            sb.Append("SOURCE FORMAT AUTHORITY: set sourceFormatAuthority=true ONLY when the latest request explicitly makes a supplied/attached/current artifact authoritative for the requested output formatting, layout, design, master, styles, or native structure (for example preserve/copy/keep that format while genericizing or replacing content). Do not set it merely because an attachment supplies facts/content/reference material. Base this on semantic intent, not language-specific keywords. sourceFormatReason is one short reason. ")
            sb.Append("Use only the advertised capability metadata below; do not preload or infer resource bodies. Candidates:")

            For Each item In candidates
                sb.AppendLine()
                sb.Append("- ").Append(item.Name).Append(" [").Append(item.Category.Trim()).Append("]")
                Dim shortDesc As String = Shrink(item.Description, 700)
                If shortDesc <> "" Then sb.Append(": ").Append(shortDesc)
            Next

            Return sb.ToString()
        End Function

        Public Shared Function BuildUserPrompt(latestUserRequestRaw As String,
                                               Optional hostTaskSummary As String = "") As String
            Dim sb As New StringBuilder()
            sb.AppendLine("[BOOTSTRAP_TASK_CONTEXT]")
            sb.AppendLine("<LATEST_USER_REQUEST_RAW>")
            sb.AppendLine(If(latestUserRequestRaw, ""))
            sb.AppendLine("</LATEST_USER_REQUEST_RAW>")
            If Not String.IsNullOrWhiteSpace(hostTaskSummary) Then
                sb.AppendLine("<HOST_TASK_SUMMARY>")
                sb.AppendLine(hostTaskSummary.Trim())
                sb.AppendLine("</HOST_TASK_SUMMARY>")
            End If
            sb.AppendLine("[/BOOTSTRAP_TASK_CONTEXT]")
            Return sb.ToString().TrimEnd()
        End Function

        Public Shared Function ParseDecision(raw As String) As Decision
            Dim result As New Decision()
            Dim normalized As String = NormalizeJsonObject(raw)
            result.NormalizedOutput = normalized

            If String.IsNullOrWhiteSpace(normalized) Then
                result.ParseError = "empty_bootstrap_output"
                Return result
            End If

            Try
                Dim obj As JObject = JObject.Parse(normalized)
                result.JsonValid = True

                Dim languageToken As JToken = obj("language")
                If languageToken IsNot Nothing AndAlso languageToken.Type = JTokenType.String Then
                    result.Language = languageToken.Value(Of String)().Trim()
                    result.LanguageValid = result.Language <> ""
                End If

                Dim memoryObj As New JObject From {
                    {"memoryGroundingMode", obj("memoryGroundingMode")},
                    {"reason", obj("memoryReason")},
                    {"shouldExposeRecentMemoryStubs", obj("shouldExposeRecentMemoryStubs")},
                    {"explicitStoredMemoryRequired", obj("explicitStoredMemoryRequired")}
                }
                Dim memoryNormalized As String = ""
                Dim memoryError As String = ""
                result.MemoryDecision = ToolCallSequencing.ParseMemoryGroundingIntentClassifierDecision(
                    memoryObj.ToString(Newtonsoft.Json.Formatting.None),
                    memoryNormalized,
                    memoryError)
                result.MemoryValid = result.MemoryDecision IsNot Nothing AndAlso result.MemoryDecision.IsValid

                result.RouteKind = If(obj.Value(Of String)("routeKind"), "").Trim().ToLowerInvariant()
                result.RouteName = If(obj.Value(Of String)("routeName"), "").Trim()
                result.RouteReason = If(obj.Value(Of String)("routeReason"), "").Trim()
                result.BootstrapLoad = If(obj.Value(Of String)("bootstrapLoad"), "").Trim()

                Dim sourceFormatToken As JToken = obj("sourceFormatAuthority")
                If sourceFormatToken IsNot Nothing AndAlso sourceFormatToken.Type = JTokenType.Boolean Then
                    result.SourceFormatAuthority = sourceFormatToken.Value(Of System.Boolean)()
                    result.SourceFormatAuthorityValid = True
                End If
                result.SourceFormatAuthorityReason = If(obj.Value(Of System.String)("sourceFormatReason"), System.String.Empty).Trim()

                If result.RouteKind = CapabilityRoutingTool.KindNone Then
                    result.RoutingSyntaxValid =
                        String.IsNullOrWhiteSpace(result.RouteName) AndAlso
                        String.IsNullOrWhiteSpace(result.BootstrapLoad)
                    If result.RoutingSyntaxValid Then
                        result.RouteName = ""
                        result.BootstrapLoad = ""
                    End If
                ElseIf (result.RouteKind = CapabilityRoutingTool.KindSkill OrElse
                        result.RouteKind = CapabilityRoutingTool.KindAgent) AndAlso
                       result.RouteName <> "" AndAlso
                       String.Equals(result.RouteName, result.BootstrapLoad, StringComparison.OrdinalIgnoreCase) Then
                    result.RoutingSyntaxValid = True
                End If

                If Not result.LanguageValid Then result.ParseError = AppendError(result.ParseError, "invalid_language")
                If Not result.MemoryValid Then result.ParseError = AppendError(result.ParseError, "invalid_memory_decision")
                If Not result.RoutingSyntaxValid Then result.ParseError = AppendError(result.ParseError, "invalid_route_or_bootstrap_load")
                If Not result.SourceFormatAuthorityValid Then result.ParseError = AppendError(result.ParseError, "invalid_source_format_authority")

                Return result
            Catch ex As System.Exception
                result.ParseError = ex.Message
                Return result
            End Try
        End Function

        Private Shared Function NormalizeJsonObject(raw As String) As String
            Dim text As String = If(raw, "").Trim()
            If text = "" Then Return ""

            If text.StartsWith("```", StringComparison.Ordinal) Then
                Dim firstBreak As Integer = text.IndexOfAny(New Char() {ChrW(10), ChrW(13)})
                If firstBreak >= 0 Then text = text.Substring(firstBreak + 1)
                Dim closingFence As Integer = text.LastIndexOf("```", StringComparison.Ordinal)
                If closingFence >= 0 Then text = text.Substring(0, closingFence)
                text = text.Trim()
            End If

            Dim firstBrace As Integer = text.IndexOf("{"c)
            Dim lastBrace As Integer = text.LastIndexOf("}"c)
            If firstBrace >= 0 AndAlso lastBrace > firstBrace Then
                text = text.Substring(firstBrace, lastBrace - firstBrace + 1)
            End If
            Return text.Trim()
        End Function

        Private Shared Function AppendError(existing As String, value As String) As String
            If String.IsNullOrWhiteSpace(existing) Then Return value
            Return existing & ";" & value
        End Function

        Private Shared Function Shrink(value As String, maxLength As Integer) As String
            Dim text As String = If(value, "").Replace(vbCr, " ").Replace(vbLf, " ").Trim()
            While text.Contains("  ")
                text = text.Replace("  ", " ")
            End While
            If text.Length <= maxLength Then Return text
            Return text.Substring(0, maxLength - 1).TrimEnd() & "…"
        End Function

    End Class

End Namespace
