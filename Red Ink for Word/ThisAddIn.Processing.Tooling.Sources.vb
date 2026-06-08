' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.Processing.Tooling.Sources.vb
' Purpose: Internal knowledge store tool management and source resolution.
'
' Responsibilities:
'  - Identify knowledge store tools by encoded hash tokens.
'  - Build knowledge store tool names (prefix + SHA256 token).
'  - Resolve tool names to KnowledgeStoreCatalog definitions.
'  - Implement hash-based token matching for indexed stores.
'  - Support fallback matching (single store, prefix-based fuzzy match).
'  - Log token resolution diagnostics and ambiguity handling.
'  - Provide common prefix matching for truncated hash recovery.
'
' Architecture:
'  - One-way hash encoding (SHA256) for knowledge store identification.
'  - Exact match preferred, fallback to single-store default, then fuzzy prefix match.
'  - Prevents tool name collisions across knowledge stores.
'
' External Dependencies:
'  - System.Security.Cryptography for SHA256 hashing.
'  - KnowledgeStoreCatalog for store definitions.
'  - ToolingFileLogger for diagnostics.
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


    Private Function IsInternalKnowledgeToolName(toolName As String) As Boolean
        Return Not String.IsNullOrWhiteSpace(toolName) AndAlso
               toolName.StartsWith(InternalKnowledgeToolNamePrefix, StringComparison.OrdinalIgnoreCase)
    End Function

    Private Function BuildInternalKnowledgeToolName(store As KnowledgeStoreCatalog.KnowledgeStoreDefinition) As String
        If store Is Nothing Then Return ""
        Return InternalKnowledgeToolNamePrefix & EncodeToolToken(store.StoreId)
    End Function
    Private Function GetKnowledgeStoreForToolName(toolName As String) As KnowledgeStoreCatalog.KnowledgeStoreDefinition
        If Not IsInternalKnowledgeToolName(toolName) Then
            Return Nothing
        End If

        Dim encodedToken As String = toolName.Substring(InternalKnowledgeToolNamePrefix.Length)

        ' Hash-based tokens are one-way — match by recomputing the hash for
        ' each known store and comparing against the token in the tool name.
        Dim indexedStores = GetIndexedKnowledgeStores()
        If indexedStores Is Nothing OrElse indexedStores.Count = 0 Then Return Nothing

        ' Exact hash match
        For Each store In indexedStores
            Dim expectedHash = EncodeToolToken(store.StoreId)
            If String.Equals(encodedToken, expectedHash, StringComparison.OrdinalIgnoreCase) Then
                Return store
            End If
        Next

        ' If there is only one knowledge store, return it directly — no ambiguity
        If indexedStores.Count = 1 Then
            ToolingFileLogger.LogWarn(
                "Knowledge tool name did not match any store hash; " &
                "falling back to the only available store.",
                details:=$"ToolName='{toolName}', Token='{encodedToken}'")
            Return indexedStores(0)
        End If

        ' Multiple stores: fuzzy match by longest common prefix of the token
        ' (handles cases where the LLM truncates the hash)
        Dim bestMatch As KnowledgeStoreCatalog.KnowledgeStoreDefinition = Nothing
        Dim bestMatchLen As Integer = 0

        For Each store In indexedStores
            Dim expectedName = BuildInternalKnowledgeToolName(store)
            If String.IsNullOrWhiteSpace(expectedName) Then Continue For

            Dim commonLen = GetCommonPrefixLength(toolName, expectedName)
            If commonLen > InternalKnowledgeToolNamePrefix.Length AndAlso commonLen > bestMatchLen Then
                bestMatchLen = commonLen
                bestMatch = store
            End If
        Next

        If bestMatch IsNot Nothing Then
            ToolingFileLogger.LogWarn(
                $"Knowledge tool name partially matched store '{bestMatch.Name}' " &
                $"(prefix match: {bestMatchLen} of {toolName.Length} chars).",
                details:=$"ToolName='{toolName}', Token='{encodedToken}'")
        End If

        Return bestMatch
    End Function

    Private Shared Function GetCommonPrefixLength(a As String, b As String) As Integer
        If a Is Nothing OrElse b Is Nothing Then Return 0
        Dim maxLen = Math.Min(a.Length, b.Length)
        For i As Integer = 0 To maxLen - 1
            If Char.ToUpperInvariant(a(i)) <> Char.ToUpperInvariant(b(i)) Then Return i
        Next
        Return maxLen
    End Function

    Private Function GetIndexedKnowledgeStores() As List(Of KnowledgeStoreCatalog.KnowledgeStoreDefinition)
        Dim result As New List(Of KnowledgeStoreCatalog.KnowledgeStoreDefinition)()

        Try
            Dim stores = KnowledgeStoreCatalog.GetActiveStores(_context)

            result = stores.
                Where(Function(s)
                          If s Is Nothing Then Return False

                          Try
                              Dim manifest = KnowledgeStoreManifest.Load(s)
                              Return manifest IsNot Nothing AndAlso
                                     manifest.Entries IsNot Nothing AndAlso
                                     manifest.Entries.Count > 0
                          Catch
                              Return False
                          End Try
                      End Function).
                OrderBy(Function(s) If(KnowledgeStoreCatalog.GetDisplayLabel(s), "").ToLowerInvariant()).
                ToList()
        Catch
        End Try

        Return result
    End Function

    Private Function BuildInternalKnowledgeToolDefinition(store As KnowledgeStoreCatalog.KnowledgeStoreDefinition) As String
        Dim displayLabel As String = KnowledgeStoreCatalog.GetDisplayLabel(store)
        Dim toolName As String = BuildInternalKnowledgeToolName(store)

        Dim schema = KnowledgeStoreSchema.LoadOrCreate(store.ResolvedSourcePath)
        Dim contentHint As String = ""
        If schema IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(schema.ToolingDescription) Then
            contentHint = " " & schema.ToolingDescription.Trim()
        End If

        Dim definition As New JObject(
        New JProperty("name", toolName),
        New JProperty("description",
            $"Searches only the user's Knowledge Store '{displayLabel}'.{contentHint} " &
            "Use query for retrieval and task_prompt for the user's actual task if you want verbatim task-relevant excerpts from the retrieved source files. " &
            "Use this for the user's own materials in that source, not for public-web lookup."),
        New JProperty("parameters",
            New JObject(
                New JProperty("type", "object"),
                New JProperty("properties",
                    New JObject(
                        New JProperty("query",
                            New JObject(
                                New JProperty("type", "string"),
                                New JProperty("description", "Optional natural-language retrieval query to search within this Knowledge Store.")
                            )
                        ),
                        New JProperty("task_prompt",
                            New JObject(
                                New JProperty("type", "string"),
                                New JProperty("description", $"Optional full user task/request. If supplied, the resolver may read up to {SharedLibrary.SharedLibrary.KnowledgeTriggerHelper.MaxRelevantExtractDocuments} top matching source documents and return only verbatim passages relevant to this task. This is separate from the retrieval query.")
                            )
                        ),
                        New JProperty("tag",
                            New JObject(
                                New JProperty("type", "string"),
                                New JProperty("description", "Optional tag filter within this Knowledge Store.")
                            )
                        ),
                        New JProperty("max_results",
                            New JObject(
                                New JProperty("type", "integer"),
                                New JProperty("description", "Optional maximum number of results to retrieve.")
                            )
                        )
                    )
                ),
                New JProperty("additionalProperties", False)
            )
        )
    )

        Return definition.ToString(Formatting.None)
    End Function

    Private Function BuildInternalKnowledgeToolInstructionsPrompt(store As KnowledgeStoreCatalog.KnowledgeStoreDefinition) As String
        Dim displayLabel As String = KnowledgeStoreCatalog.GetDisplayLabel(store)

        Dim schema = KnowledgeStoreSchema.LoadOrCreate(store.ResolvedSourcePath)
        Dim contentHint As String = ""
        If schema IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(schema.ToolingDescription) Then
            contentHint = " Content: " & schema.ToolingDescription.Trim()
        End If

        Return $"Searches only the user's Knowledge Store '{displayLabel}'.{contentHint} " &
        "Provide query (optional retrieval hint), task_prompt (optional full user task), tag (optional), and max_results (optional). " &
        $"If task_prompt is supplied, the resolver may read up to {SharedLibrary.SharedLibrary.KnowledgeTriggerHelper.MaxRelevantExtractDocuments} top matching source documents and return only verbatim passages relevant to that task. " &
        "If query is omitted, the tool may still return the most relevant documents from that store or all documents matching the tag. " &
        "Do NOT use this tool for public information or general knowledge. " &
        $"When citing results, mention the document name and store name '{displayLabel}'."
    End Function

    Private Function GetInternalKnowledgeTool(store As KnowledgeStoreCatalog.KnowledgeStoreDefinition) As ModelConfig
        Dim displayLabel As String = KnowledgeStoreCatalog.GetDisplayLabel(store)

        Return New ModelConfig() With {
            .ToolName = BuildInternalKnowledgeToolName(store),
            .ToolInstructionsPrompt = BuildInternalKnowledgeToolInstructionsPrompt(store),
            .ToolDefinition = BuildInternalKnowledgeToolDefinition(store),
            .ModelDescription = $"Knowledge Store: {displayLabel}{InternalToolSuffix}",
            .Tool = True,
            .ToolPriority = 997,
            .ToolErrorHandling = "skip"
        }
    End Function

    Private Function GetInternalKnowledgeTool(toolName As String) As ModelConfig
        Dim store = GetKnowledgeStoreForToolName(toolName)
        If store Is Nothing Then Return Nothing
        Return GetInternalKnowledgeTool(store)
    End Function

    Private Function GetInternalKnowledgeTools() As List(Of ModelConfig)
        Return GetIndexedKnowledgeStores().
            Select(Function(store) GetInternalKnowledgeTool(store)).
            Where(Function(tool) tool IsNot Nothing).
            ToList()
    End Function


    Private Function BuildKnowledgeToolStoreInventoryLine() As String
        Dim storeLabels As New List(Of String)()

        Try
            Dim stores = KnowledgeStoreCatalog.GetActiveStores(_context)

            For Each store In stores
                If store Is Nothing Then Continue For

                Dim label As String = KnowledgeStoreCatalog.GetDisplayLabel(store)
                If String.IsNullOrWhiteSpace(label) Then Continue For

                Try
                    Dim manifest = KnowledgeStoreManifest.Load(store)
                    If manifest IsNot Nothing AndAlso manifest.Entries IsNot Nothing AndAlso manifest.Entries.Count > 0 Then
                        label &= $" ({manifest.Entries.Count} docs)"
                    End If
                Catch
                End Try

                If Not storeLabels.Any(Function(x) String.Equals(x, label, StringComparison.OrdinalIgnoreCase)) Then
                    storeLabels.Add(label)
                End If
            Next
        Catch
        End Try

        If storeLabels.Count = 0 Then
            Return ""
        End If

        Return "Knowledge stores currently available: " & String.Join(", ", storeLabels) & "."
    End Function

    Private Function GetAvailableKnowledgeStoreNames() As List(Of String)
        Dim storeNames As New List(Of String)()

        Try
            Dim stores = KnowledgeStoreCatalog.GetActiveStores(_context)

            For Each store In stores
                If store Is Nothing Then Continue For

                Dim displayLabel = KnowledgeStoreCatalog.GetDisplayLabel(store)
                If Not String.IsNullOrWhiteSpace(displayLabel) Then
                    If Not storeNames.Any(Function(x) String.Equals(x, displayLabel, StringComparison.OrdinalIgnoreCase)) Then
                        storeNames.Add(displayLabel)
                    End If
                End If

                Dim plainName As String = If(store.Name, "").Trim()
                If plainName <> "" Then
                    If Not storeNames.Any(Function(x) String.Equals(x, plainName, StringComparison.OrdinalIgnoreCase)) Then
                        storeNames.Add(plainName)
                    End If
                End If
            Next
        Catch
        End Try

        Return storeNames
    End Function


    Private Function BuildInternalKnowledgeToolDefinition() As String
        Dim storeInventory As String = BuildKnowledgeToolStoreInventoryLine()
        Dim descriptionSuffix As String = If(String.IsNullOrWhiteSpace(storeInventory), "", " " & storeInventory)

        Dim definition As New JObject(
        New JProperty("name", InternalKnowledgeToolName),
        New JProperty("description",
            "Searches the user's local knowledge stores (their own curated document collections). " &
            "This tool mirrors the Freestyle knowledge trigger functionality. " &
            "You can either use structured arguments (query/store/tag/max_results) or pass the exact Freestyle trigger syntax via raw_trigger. " &
            "Use this for the user's own materials, not for public-web lookup." & descriptionSuffix),
        New JProperty("parameters",
            New JObject(
                New JProperty("type", "object"),
                New JProperty("properties",
                    New JObject(
                        New JProperty("query",
                            New JObject(
                                New JProperty("type", "string"),
                                New JProperty("description", "Optional natural-language knowledge-store query. If omitted, the tool can still search broadly or within a given store/tag scope.")
                            )
                        ),
                        New JProperty("store",
                            New JObject(
                                New JProperty("type", "string"),
                                New JProperty("description", "Optional knowledge store name. Use exactly one of the store names exposed in the tool instructions.")
                            )
                        ),
                        New JProperty("tag",
                            New JObject(
                                New JProperty("type", "string"),
                                New JProperty("description", "Optional tag filter, equivalent to Freestyle syntax 'tag:YourTag'.")
                            )
                        ),
                        New JProperty("raw_trigger",
                            New JObject(
                                New JProperty("type", "string"),
                                New JProperty("description", "Optional exact Freestyle-style trigger. Examples: '(kb)', '(kb:termination without notice)', '(kb:store:Policies confidentiality)', '(kb:tag:NDA confidentiality)'. If supplied, this takes precedence over query/store/tag.")
                            )
                        ),
                        New JProperty("max_results",
                            New JObject(
                                New JProperty("type", "integer"),
                                New JProperty("description", "Optional maximum number of results to retrieve. Best effort; the resolver may still enforce its own cap.")
                            )
                        )
                    )
                ),
                New JProperty("additionalProperties", False)
            )
        )
    )

        Return definition.ToString(Formatting.None)
    End Function

    Private Function BuildInternalKnowledgeToolInstructionsPrompt() As String
        Dim kbTrigger As String = SharedLibrary.SharedLibrary.KnowledgeTriggerHelper.KbTrigger
        Dim kbTriggerPrefix As String = SharedLibrary.SharedLibrary.KnowledgeTriggerHelper.KbTriggerPrefix

        Dim sb As New StringBuilder()
        sb.Append("knowledge_search: Searches the user's local knowledge stores — the user's own curated document collections such as contracts, policies, briefs, manuals, templates, emails, and reference material. ")
        sb.Append("This tool supports the same search semantics as the Freestyle knowledge trigger. ")
        sb.Append("Prefer structured arguments for normal calls: query (optional), store (optional), tag (optional), max_results (optional). ")
        sb.Append("If you need exact parity with Freestyle, pass raw_trigger using the literal syntax. ")
        sb.Append($"Valid trigger forms include '{kbTrigger}', '{kbTriggerPrefix}your query)', '{kbTriggerPrefix}store:StoreName your query)', '{kbTriggerPrefix}tag:TagName your query)', and combinations such as '{kbTriggerPrefix}store:StoreName tag:TagName your query)'. ")
        sb.Append("If store is omitted, all stores are searched. If query is omitted but store and/or tag is provided, the tool still performs a scoped retrieval. If everything is omitted, it performs a broad cross-store retrieval. ")

        Dim storeInventory As String = BuildKnowledgeToolStoreInventoryLine()
        If Not String.IsNullOrWhiteSpace(storeInventory) Then
            sb.Append(storeInventory & " ")
            sb.Append("Use the store names exactly as listed. ")
        End If

        sb.Append("Do NOT use this tool for public information or general knowledge — use your own knowledge or internet_search for that. ")
        sb.Append("When citing results, mention the document name and store name.")

        Return sb.ToString()
    End Function

    Private Function BuildKnowledgeToolTrigger(query As String, storeName As String, tagName As String, rawTrigger As String) As String
        Dim kbTrigger As String = SharedLibrary.SharedLibrary.KnowledgeTriggerHelper.KbTrigger
        Dim kbTriggerPrefix As String = SharedLibrary.SharedLibrary.KnowledgeTriggerHelper.KbTriggerPrefix

        Dim QuoteIfNeeded As Func(Of String, String) =
        Function(value As String) As String
            Dim trimmed As String = If(value, "").Trim()
            If trimmed = "" Then Return ""

            If (trimmed.StartsWith("""", StringComparison.Ordinal) AndAlso trimmed.EndsWith("""", StringComparison.Ordinal)) OrElse
               (trimmed.StartsWith("'", StringComparison.Ordinal) AndAlso trimmed.EndsWith("'", StringComparison.Ordinal)) Then
                Return trimmed
            End If

            If trimmed.IndexOf(" "c) >= 0 OrElse trimmed.IndexOf(ControlChars.Tab) >= 0 Then
                Return """" & trimmed.Replace("""", """""") & """"
            End If

            Return trimmed
        End Function

        If Not String.IsNullOrWhiteSpace(rawTrigger) Then
            Dim normalized As String = rawTrigger.Trim()

            If String.Equals(normalized, kbTrigger, StringComparison.OrdinalIgnoreCase) Then
                Return kbTrigger
            End If

            If normalized.StartsWith(kbTriggerPrefix, StringComparison.OrdinalIgnoreCase) Then
                If Not normalized.EndsWith(")") Then
                    normalized &= ")"
                End If
                Return normalized
            End If

            If normalized.StartsWith("(kb", StringComparison.OrdinalIgnoreCase) Then
                If Not normalized.EndsWith(")") Then
                    normalized &= ")"
                End If
                Return normalized
            End If

            Return kbTriggerPrefix & normalized.TrimEnd(")"c).Trim() & ")"
        End If

        If Not String.IsNullOrWhiteSpace(query) AndAlso
       String.IsNullOrWhiteSpace(storeName) AndAlso
       String.IsNullOrWhiteSpace(tagName) Then

            Dim normalizedQuery As String = query.Trim()

            If String.Equals(normalizedQuery, kbTrigger, StringComparison.OrdinalIgnoreCase) Then
                Return kbTrigger
            End If

            If normalizedQuery.IndexOf(kbTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return normalizedQuery
            End If

            If normalizedQuery.IndexOf(kbTriggerPrefix, StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return normalizedQuery
            End If

            If normalizedQuery.StartsWith("(kb", StringComparison.OrdinalIgnoreCase) Then
                If Not normalizedQuery.EndsWith(")") Then
                    normalizedQuery &= ")"
                End If
                Return normalizedQuery
            End If

            Return kbTrigger
        End If

        Dim parts As New List(Of String)()

        If Not String.IsNullOrWhiteSpace(storeName) Then
            parts.Add("store:" & QuoteIfNeeded(storeName))
        End If

        If Not String.IsNullOrWhiteSpace(tagName) Then
            parts.Add("tag:" & QuoteIfNeeded(tagName))
        End If

        If Not String.IsNullOrWhiteSpace(query) Then
            parts.Add(query.Trim())
        End If

        If parts.Count = 0 Then
            Return kbTrigger
        End If

        Return kbTriggerPrefix & String.Join(" ", parts).Trim() & ")"
    End Function



    Private Function NormalizeKnowledgeToolContentFileUris(content As String) As String
        If String.IsNullOrWhiteSpace(content) Then Return If(content, "")

        Return Regex.Replace(
            content,
            "<KSDOCUMENT\b[^>]*>",
            Function(m As Match) NormalizeKnowledgeToolDocumentTag(m.Value),
            RegexOptions.IgnoreCase Or RegexOptions.Singleline)
    End Function

    Private Function NormalizeKnowledgeToolDocumentTag(tag As String) As String
        If String.IsNullOrWhiteSpace(tag) Then Return tag

        Dim storeName As String = GetKnowledgeToolDocumentAttribute(tag, "store")
        If String.IsNullOrWhiteSpace(storeName) Then Return tag

        Dim store = KnowledgeStoreCatalog.GetStoreByName(storeName, _context)
        If store Is Nothing Then Return tag

        Dim rewritten As String = tag
        rewritten = RewriteKnowledgeToolDocumentAttribute(rewritten, "sourcePath", store.ResolvedSourcePath)

        Dim wikiRoot As String = KnowledgeStoreCatalog.GetWikiPath(store, createIfMissing:=False)
        If Not String.IsNullOrWhiteSpace(wikiRoot) Then
            rewritten = RewriteKnowledgeToolDocumentAttribute(rewritten, "wikiPath", wikiRoot)
        End If

        Return rewritten
    End Function

    Private Shared Function GetKnowledgeToolDocumentAttribute(tag As String, attributeName As String) As String
        If String.IsNullOrWhiteSpace(tag) OrElse String.IsNullOrWhiteSpace(attributeName) Then Return ""

        Dim m As Match = Regex.Match(
            tag,
            "\b" & Regex.Escape(attributeName) & "=""""(?<v>[^""]*)""""",
            RegexOptions.IgnoreCase)

        If Not m.Success Then Return ""

        Return System.Net.WebUtility.HtmlDecode(m.Groups("v").Value)
    End Function

    Private Shared Function RewriteKnowledgeToolDocumentAttribute(tag As String, attributeName As String, basePath As String) As String
        Dim currentValue As String = GetKnowledgeToolDocumentAttribute(tag, attributeName)
        If String.IsNullOrWhiteSpace(currentValue) Then Return tag

        Dim rewrittenValue As String = BuildKnowledgeToolAbsoluteFileUri(currentValue, basePath)
        If String.IsNullOrWhiteSpace(rewrittenValue) Then Return tag

        Return Regex.Replace(
            tag,
            "\b" & Regex.Escape(attributeName) & "=""""[^""]*""""",
            attributeName & "=""" & System.Security.SecurityElement.Escape(rewrittenValue) & """",
            RegexOptions.IgnoreCase)
    End Function

    Private Shared Function BuildKnowledgeToolAbsoluteFileUri(value As String, basePath As String) As String
        If String.IsNullOrWhiteSpace(value) Then Return ""

        Dim candidate As String = System.Net.WebUtility.HtmlDecode(value).Trim()
        If candidate = "" Then Return ""

        Try
            If candidate.StartsWith("file://", StringComparison.OrdinalIgnoreCase) Then
                candidate = New Uri(candidate).LocalPath
            End If
        Catch
        End Try

        candidate = candidate.Replace("/"c, Path.DirectorySeparatorChar)

        If Path.IsPathRooted(candidate) AndAlso File.Exists(candidate) Then
            Return New Uri(Path.GetFullPath(candidate)).AbsoluteUri
        End If

        If Not String.IsNullOrWhiteSpace(basePath) Then
            Dim trimmed As String = candidate.TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            Dim combined As String = Path.GetFullPath(Path.Combine(basePath, trimmed))
            If File.Exists(combined) Then
                Return New Uri(combined).AbsoluteUri
            End If
        End If

        Return ""
    End Function

    ''' <summary>
    ''' Creates a built-in internal knowledge store search tool configuration as a <see cref="ModelConfig"/>.
    ''' Only meaningful when <c>INI_KnowledgeStorePath</c> or <c>INI_KnowledgeStorePathLocal</c> is configured
    ''' and at least one knowledge store has an indexed manifest.
    ''' </summary>
    ''' <returns>Internal knowledge search tool configuration.</returns>
    Public Function GetInternalKnowledgeTool() As ModelConfig
        Return New ModelConfig() With {
        .ToolName = InternalKnowledgeToolName,
        .ToolInstructionsPrompt = BuildInternalKnowledgeToolInstructionsPrompt(),
        .ToolDefinition = BuildInternalKnowledgeToolDefinition(),
        .ModelDescription = "Knowledge Store Search" & InternalToolSuffix,
        .Tool = True,
        .ToolPriority = 997,
        .ToolErrorHandling = "skip"
    }
    End Function


    ''' <summary>
    ''' Executes the internal knowledge store search tool by querying the merged index
    ''' via KnowledgeQueryService and returning tagged document content blocks.
    ''' </summary>
    ''' <param name="toolCall">Tool call payload containing <c>query</c> and optional <c>max_results</c>.</param>
    ''' <param name="context">Tool execution context for logging and diagnostics.</param>
    ''' <returns>Tool response containing relevant document content or an error.</returns>
    Private Async Function ExecuteInternalKnowledgeTool(toolCall As ToolCall, context As ToolExecutionContext) As Task(Of ToolResponse)
        Dim response As New ToolResponse() With {
        .CallId = toolCall.CallId,
        .ToolName = toolCall.ToolName
    }

        Try
            Dim boundStore = GetKnowledgeStoreForToolName(toolCall.ToolName)

            If boundStore Is Nothing Then
                response.Success = False
                response.ErrorMessage = "The selected Knowledge Store source could not be resolved."
                ToolingFileLogger.LogWarn("Internal knowledge tool: bound store could not be resolved.",
                details:=$"ToolName='{toolCall.ToolName}'")
                Return response
            End If

            Dim storeLabel As String = KnowledgeStoreCatalog.GetDisplayLabel(boundStore)
            Dim query As String = GetToolArgumentString(toolCall.Arguments, "query")
            Dim taskPrompt As String = GetToolArgumentString(toolCall.Arguments, "task_prompt")
            Dim tagName As String = GetToolArgumentString(toolCall.Arguments, "tag")
            Dim rawTrigger As String = GetToolArgumentString(toolCall.Arguments, "raw_trigger")
            Dim requestedStore As String = GetToolArgumentString(toolCall.Arguments, "store")

            If String.IsNullOrWhiteSpace(taskPrompt) Then
                If context IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(context.LatestUserRequestRaw) Then
                    taskPrompt = context.LatestUserRequestRaw.Trim()
                ElseIf context IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(context.HostTaskSummary) Then
                    taskPrompt = context.HostTaskSummary.Trim()
                ElseIf Not String.IsNullOrWhiteSpace(query) Then
                    taskPrompt = query.Trim()
                End If
            End If

            Dim maxResults As Integer = 5
            If toolCall.Arguments.ContainsKey("max_results") Then
                Dim mr As Integer
                If Integer.TryParse(toolCall.Arguments("max_results")?.ToString(), mr) Then
                    maxResults = Math.Min(Math.Max(1, mr), 10)
                End If
            End If

            context.Log($"Knowledge store source: {storeLabel}")
            ToolingFileLogger.LogStep(
            $"Knowledge store source: '{storeLabel}'; query='{query}'; task_prompt='{taskPrompt}'; tag='{tagName}'; raw_trigger='{rawTrigger}'; max_results={maxResults}")

            Dim kbRequest As KnowledgeTriggerHelper.KnowledgeRequest = Nothing

            If Not String.IsNullOrWhiteSpace(rawTrigger) Then
                kbRequest = KnowledgeTriggerHelper.TryParseKnowledgeTrigger(rawTrigger)

                If kbRequest Is Nothing Then
                    response.Success = True
                    response.Response = "The supplied raw_trigger could not be parsed."
                    Return response
                End If
            Else
                kbRequest = New KnowledgeTriggerHelper.KnowledgeRequest() With {
                .LoadAll = String.IsNullOrWhiteSpace(query) AndAlso String.IsNullOrWhiteSpace(tagName),
                .SearchQuery = If(query, "").Trim(),
                .RawTrigger = If(String.IsNullOrWhiteSpace(query) AndAlso String.IsNullOrWhiteSpace(tagName),
                                 KnowledgeTriggerHelper.KbTrigger,
                                 ""),
                .OriginalParameter = If(query, "").Trim()
            }

                If Not String.IsNullOrWhiteSpace(tagName) Then
                    kbRequest.Tags = tagName.Split(","c).
                    Select(Function(t) t.Trim()).
                    Where(Function(t) Not String.IsNullOrWhiteSpace(t)).
                    ToArray()
                    kbRequest.HasExplicitTagFilter = kbRequest.Tags IsNot Nothing AndAlso kbRequest.Tags.Length > 0
                End If
            End If

            Dim boundStoreName As String = If(boundStore.Name, "").Trim()

            If Not String.IsNullOrWhiteSpace(requestedStore) AndAlso
           Not String.IsNullOrWhiteSpace(boundStoreName) AndAlso
           Not requestedStore.Equals(boundStoreName, StringComparison.OrdinalIgnoreCase) Then

                context.Log($"Ignoring requested store '{requestedStore}' because this tool is bound to '{boundStoreName}'.")
                ToolingFileLogger.LogStep($"Ignoring requested store '{requestedStore}' because tool is bound to '{boundStoreName}'.")
            End If

            If Not String.IsNullOrWhiteSpace(boundStoreName) Then
                kbRequest.StoreName = boundStoreName
                kbRequest.HasExplicitStoreFilter = True
            ElseIf Not String.IsNullOrWhiteSpace(requestedStore) Then
                kbRequest.StoreName = requestedStore.Trim()
                kbRequest.HasExplicitStoreFilter = True
            End If

            Dim resolveOptions As New KnowledgeTriggerHelper.KnowledgeResolveOptions With {
    .TaskPrompt = taskPrompt,
    .IncludeRelevantExtracts = Not String.IsNullOrWhiteSpace(taskPrompt),
    .IncludeFullDocumentContent = False,
    .MaxResults = maxResults,
    .ForceSemanticSearch = Not String.IsNullOrWhiteSpace(query)
}

            Dim resolved = Await KnowledgeTriggerHelper.ResolveKnowledgeAsync(
            request:=kbRequest,
            context:=_context,
            options:=resolveOptions).ConfigureAwait(False)

            If String.IsNullOrWhiteSpace(resolved.Content) Then
                response.Success = True
                response.Response = If(String.IsNullOrWhiteSpace(resolved.StatusMessage),
                                   $"No relevant documents found in Knowledge Store '{storeLabel}'.",
                                   resolved.StatusMessage)
                Return response
            End If

            response.Success = True
            response.Response = NormalizeKnowledgeToolContentFileUris(resolved.Content)

            context.Log($"Knowledge search returned content ({resolved.Content.Length:N0} chars) from '{storeLabel}'.", "success")

        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = $"Knowledge store search failed: {ex.Message}"
            ToolingFileLogger.LogError("Internal knowledge tool error.", ex:=ex)
        End Try

        Return response
    End Function

    Private Function IsMCPStreamableToolCall(endpoint As String, apiCall As String) As Boolean
        If String.IsNullOrWhiteSpace(endpoint) Then
            Return False
        End If

        If endpoint.StartsWith(SharedMethods.MCP_STREAMABLE_PREFIX, StringComparison.OrdinalIgnoreCase) Then
            Return True
        End If

        If endpoint.StartsWith(SharedMethods.MCP_SSE_PREFIX, StringComparison.OrdinalIgnoreCase) Then
            Return False
        End If

        If String.IsNullOrWhiteSpace(apiCall) Then
            Return False
        End If

        Try
            Dim requestObj As JObject = JObject.Parse(apiCall)

            Return String.Equals(
                If(requestObj("jsonrpc")?.ToString(), ""),
                "2.0",
                StringComparison.OrdinalIgnoreCase) AndAlso
                String.Equals(
                    If(requestObj("method")?.ToString(), ""),
                    "tools/call",
                    StringComparison.OrdinalIgnoreCase)
        Catch
            Return False
        End Try
    End Function

    Private Function ShouldRetryMCPAfterUnauthorized(toolConfig As ModelConfig, ex As Exception) As Boolean
        If toolConfig Is Nothing OrElse Not toolConfig.OAuth2 Then Return False
        If ex Is Nothing Then Return False

        Dim message As String = If(ex.Message, "")
        Return message.IndexOf("401", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               message.IndexOf("Unauthorized", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               message.IndexOf("Invalid or expired access token", StringComparison.OrdinalIgnoreCase) >= 0
    End Function

    Private Async Function ForceRefreshToolOAuthToken(toolConfig As ModelConfig, toolName As String) As Task(Of Boolean)
        Try
            toolConfig.DecodedAPI = Await SharedMethods.GetFreshAccessToken(
                _context,
                toolConfig.OAuth2ClientMail,
                toolConfig.OAuth2Scopes,
                toolConfig.APIKey,
                toolConfig.OAuth2Endpoint,
                toolConfig.OAuth2ATExpiry,
                True,
                False,
                forceRefresh:=True).ConfigureAwait(False)

            If String.IsNullOrWhiteSpace(toolConfig.DecodedAPI) Then
                ToolingFileLogger.LogError(
                    "Forced MCP OAuth refresh returned an empty token.",
                    details:=$"ToolName='{toolName}'")
                Return False
            End If

            Return True

        Catch refreshEx As Exception
            ToolingFileLogger.LogError(
                "Forced MCP OAuth refresh failed.",
                details:=$"ToolName='{toolName}'",
                ex:=refreshEx)
            Return False
        End Try
    End Function




End Class
