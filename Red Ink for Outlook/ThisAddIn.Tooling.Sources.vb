' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.Tooling.Sources.vb
' Purpose: Manages internal knowledge store search tools and knowledge source indexing.
'          Handles tool definition generation, store discovery, and fuzzy token matching.
'
' Architecture:
'  - Knowledge Store Tool Definition:
'      - GetInternalKnowledgeTool(): Factory for per-store knowledge search tool configs.
'      - BuildInternalKnowledgeToolDefinition(): Generates JSON schema for tool parameters.
'      - BuildInternalKnowledgeToolInstructionsPrompt(): Creates instructions for LLM usage.
'  - Token-Based Tool Naming:
'      - EncodeToolToken(): SHA256 hash of store ID to fixed-length 64-char token (safe for API names).
'      - DecodeToolToken(): Returns empty string (hash is one-way).
'      - IsInternalKnowledgeToolName(): Checks InternalKnowledgeToolNamePrefix.
'      - GetKnowledgeStoreForToolName(): Matches tool name token to store via hash recomputation.
'  - Store Discovery & Indexing:
'      - GetIndexedKnowledgeStores(): Returns active stores with valid manifests, sorted by display label.
'      - BuildKnowledgeToolStoreInventoryLine(): Human-readable store list with document counts.
'      - GetAvailableKnowledgeStoreNames(): Returns both display labels and plain store names.
'  - Fuzzy Matching:
'      - Handles LLM truncation of tool names via longest common prefix matching.
'      - Falls back to single-store detection when only one store is available.
'      - Logs warnings when matches are inexact.
'  - Knowledge Tool Trigger Building:
'      - BuildKnowledgeToolTrigger(): Constructs Freestyle-compatible (kb) trigger syntax.
'      - Supports raw trigger pass-through, query-only, store-scoped, and tag-filtered searches.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.Reflection
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Microsoft.Web.WebView2.Core
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedContext
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn


    Private Const InternalKnowledgeToolNamePrefix As String = "knowledge_search_store_"

    Private Function EncodeToolToken(value As String) As String
        If String.IsNullOrWhiteSpace(value) Then Return ""

        ' Use a SHA256 hash (hex, lowercase) to produce a fixed-length token
        ' that is always valid for API function names and stays well under
        ' the 128-character name limit imposed by model APIs (e.g. Gemini).
        Using sha = System.Security.Cryptography.SHA256.Create()
            Dim hashBytes = sha.ComputeHash(Encoding.UTF8.GetBytes(value))
            Dim sb As New StringBuilder(hashBytes.Length * 2)
            For Each b In hashBytes
                sb.Append(b.ToString("x2"))
            Next
            Return sb.ToString()   ' 64 hex chars, always
        End Using
    End Function


    Private Function DecodeToolToken(value As String) As String
        ' Hash-based tokens are one-way; decoding is no longer possible.
        ' Callers must use GetKnowledgeStoreForToolName which matches
        ' by recomputing hashes against known stores.
        Return ""
    End Function

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

        ' Load schema to get the user-authored tooling description
        Dim schema = KnowledgeStoreSchema.LoadOrCreate(store.ResolvedSourcePath)
        Dim contentHint As String = ""
        If schema IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(schema.ToolingDescription) Then
            contentHint = " " & schema.ToolingDescription.Trim()
        End If

        Dim definition As New JObject(
            New JProperty("name", toolName),
            New JProperty("description",
                $"Searches only the user's Knowledge Store '{displayLabel}'.{contentHint} Use this for the user's own materials in that source, not for public-web lookup."),
            New JProperty("parameters",
                New JObject(
                    New JProperty("type", "object"),
                    New JProperty("properties",
                        New JObject(
                            New JProperty("query",
                                New JObject(
                                    New JProperty("type", "string"),
                                    New JProperty("description", "Optional natural-language query to search within this Knowledge Store.")
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

        ' Load schema to get the user-authored tooling description
        Dim schema = KnowledgeStoreSchema.LoadOrCreate(store.ResolvedSourcePath)
        Dim contentHint As String = ""
        If schema IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(schema.ToolingDescription) Then
            contentHint = " Content: " & schema.ToolingDescription.Trim()
        End If

        Return $"Searches only the user's Knowledge Store '{displayLabel}'.{contentHint} " &
            "Provide query (optional), tag (optional), and max_results (optional). " &
            "If query is omitted, the tool may return the most relevant documents from that store or all documents matching the tag. " &
            "Do NOT use this tool for public information or general knowledge. " &
            $"When citing results, mention the document name And store name '{displayLabel}'."
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

    Private Function BuildInternalKnowledgeToolDefinition() As String
        Dim storeInventory As String = BuildKnowledgeToolStoreInventoryLine()
        Dim descriptionSuffix As String = If(String.IsNullOrWhiteSpace(storeInventory), "", " " & storeInventory)

        Dim definition As New JObject(
        New JProperty("name", InternalKnowledgeToolName),
        New JProperty("description",
            "Searches the user's local knowledge stores (their own curated document collections). " &
            "This tool mirrors the Freestyle knowledge trigger functionality. " &
            "Use query for retrieval and task_prompt for the user's actual task if you want verbatim task-relevant excerpts from the retrieved source files. " &
            "You can either use structured arguments (query/task_prompt/store/tag/max_results) or pass the exact Freestyle trigger syntax via raw_trigger. " &
            "Use this for the user's own materials, not for public-web lookup." & descriptionSuffix),
        New JProperty("parameters",
            New JObject(
                New JProperty("type", "object"),
                New JProperty("properties",
                    New JObject(
                        New JProperty("query",
                            New JObject(
                                New JProperty("type", "string"),
                                New JProperty("description", "Optional retrieval query for the Knowledge Store search.")
                            )
                        ),
                        New JProperty("task_prompt",
                            New JObject(
                                New JProperty("type", "string"),
                                New JProperty("description", $"Optional full user task/request. If supplied, the resolver may read up to {SharedLibrary.SharedLibrary.KnowledgeTriggerHelper.MaxRelevantExtractDocuments} top matching source documents and return only verbatim passages relevant to this task. This is separate from the retrieval query.")
                            )
                        ),
                        New JProperty("store",
                            New JObject(
                                New JProperty("type", "string"),
                                New JProperty("description", "Optional knowledge store name. Multi-word store names are supported.")
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
                                New JProperty("description", "Optional exact Freestyle-style trigger. Examples: '(kb)', '(kb:termination without notice)', '(kb:store:""My Store"" confidentiality)', '(kb:tag:NDA confidentiality)'. If supplied, this takes precedence over query/store/tag.")
                            )
                        ),
                        New JProperty("max_results",
                            New JObject(
                                New JProperty("type", "integer"),
                                New JProperty("description", "Optional maximum number of Knowledge Store matches to include. Best effort; the resolver still enforces its own hard cap.")
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
        sb.Append("Prefer structured arguments for normal calls: query (retrieval hint), task_prompt (the user's actual task), store (optional), tag (optional), max_results (optional). ")
        sb.Append($"If task_prompt is supplied, the resolver may read up to {SharedLibrary.SharedLibrary.KnowledgeTriggerHelper.MaxRelevantExtractDocuments} top matching source documents and return only verbatim passages relevant to that task. ")
        sb.Append("If you need exact parity with Freestyle, pass raw_trigger using the literal syntax. ")
        sb.Append($"Valid trigger forms include '{kbTrigger}', '{kbTriggerPrefix}your query)', '{kbTriggerPrefix}store:""My Store"" your query)', '{kbTriggerPrefix}tag:TagName your query)', and combinations such as '{kbTriggerPrefix}store:""My Store"" tag:TagName your query)'. ")
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

    Private Function GetToolArgumentString(arguments As Dictionary(Of String, Object), key As String) As String
        If arguments Is Nothing OrElse Not arguments.ContainsKey(key) OrElse arguments(key) Is Nothing Then
            Return ""
        End If

        Dim value = arguments(key)

        If TypeOf value Is JValue Then
            Return DirectCast(value, JValue).ToString().Trim()
        End If

        Return value.ToString().Trim()
    End Function

    Private Sub TrySetLateBoundProperty(target As Object, propertyName As String, value As Object)
        If target Is Nothing Then Return

        Try
            Dim prop = target.GetType().GetProperty(
            propertyName,
            BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.IgnoreCase)

            If prop Is Nothing OrElse Not prop.CanWrite Then Return

            Dim convertedValue As Object = value

            If value IsNot Nothing Then
                Dim targetType = If(Nullable.GetUnderlyingType(prop.PropertyType), prop.PropertyType)
                convertedValue = System.Convert.ChangeType(value, targetType, Globalization.CultureInfo.InvariantCulture)
            End If

            prop.SetValue(target, convertedValue, Nothing)
        Catch
        End Try
    End Sub

    Private Function TryGetLateBoundString(target As Object, propertyName As String) As String
        If target Is Nothing Then
            Return ""
        End If

        Try
            Dim prop = target.GetType().GetProperty(
            propertyName,
            BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.IgnoreCase)

            If prop Is Nothing OrElse Not prop.CanRead Then
                Return ""
            End If

            Dim value = prop.GetValue(target, Nothing)
            Return If(value, "").ToString()
        Catch
            Return ""
        End Try
    End Function



    ''' <summary>
    ''' Executes the internal internet search tool by querying the configured search engine,
    ''' extracting result URLs via response masks, fetching qualifying page content, and returning
    ''' tagged result blocks including the search query and all visited URLs for transparency.
    ''' </summary>
    ''' <param name="toolCall">Tool call payload containing <c>query</c> and optional <c>max_results</c>/<c>max_depth</c>.</param>
    ''' <param name="context">Tool execution context for logging and diagnostics.</param>
    ''' <param name="cancellationToken">Optional cancellation token to abort execution.</param>
    ''' <returns>Tool response containing search results or an error.</returns>
    Private Async Function ExecuteInternalSearchTool(toolCall As ToolCall, context As ToolExecutionContext, Optional cancellationToken As System.Threading.CancellationToken = Nothing) As Task(Of ToolResponse)
        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId,
            .ToolName = toolCall.ToolName
        }

        Try
            cancellationToken.ThrowIfCancellationRequested()

            ' ── Validate search configuration ────────────────────────────
            If Not INI_ISearch OrElse String.IsNullOrWhiteSpace(INI_ISearch_URL) Then
                response.Success = False
                response.ErrorMessage = "Internet search is not configured or not enabled."
                ToolingFileLogger.LogWarn("Internal search tool: search not enabled/configured.",
                    details:=$"INI_ISearch={INI_ISearch}; INI_ISearch_URL='{INI_ISearch_URL}'")
                Return response
            End If

            ' ── Extract and validate parameters ──────────────────────────
            Dim query As String = ""
            If toolCall.Arguments.ContainsKey("query") Then
                query = If(toolCall.Arguments("query")?.ToString(), "").Trim()
            End If

            If String.IsNullOrWhiteSpace(query) Then
                response.Success = False
                response.ErrorMessage = "No search query provided."
                ToolingFileLogger.LogWarn("Internal search tool: empty query.",
                    details:=$"CallId={toolCall.CallId}; Args={JsonConvert.SerializeObject(toolCall.Arguments)}")
                Return response
            End If

            ' ── PII / confidential data safety net ───────────────────────
            ' Block queries that contain obvious personal data patterns.
            ' Only enforced when privacy protection is enabled.
            Dim enforcePrivacyForPII As Boolean = True  ' default for non-AutoPilot callers
            If _apConfig IsNot Nothing Then
                enforcePrivacyForPII = _apConfig.EnablePrivacyProtection
            ElseIf _context IsNot Nothing Then
                enforcePrivacyForPII = _context.INI_EnablePrivacyForSearch
            End If

            If enforcePrivacyForPII Then
                Dim piiPatterns As String() = {
                "\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b",
                "\b\d{3}[-.\s]?\d{3}[-.\s]?\d{4}\b",
                "\b\+?\d{1,3}[-.\s]?\(?\d{1,4}\)?[-.\s]?\d{2,4}[-.\s]?\d{2,4}[-.\s]?\d{0,4}\b",
                "\b\d{3}-\d{2}-\d{4}\b",
                "\b\d{2}[\./]\d{2}[\./]\d{2,4}\b(?=.*\d{2}[\./]\d{2}[\./]\d{2,4})",
                "\b[A-Z]{2}\d{2}\s?\d{4}\s?\d{4}\s?\d{4}\s?\d{4}\s?\d{0,2}\b",
                "\b(?:4\d{3}|5[1-5]\d{2}|6011|3[47]\d{2})[-\s]?\d{4}[-\s]?\d{4}[-\s]?\d{4}\b",
                "\bAHV[\s-]?\d{3}[\.\s]?\d{4}[\.\s]?\d{4}[\.\s]?\d{2}\b"
            }
                For Each piiPattern In piiPatterns
                    If Regex.IsMatch(query, piiPattern, RegexOptions.IgnoreCase) Then
                        response.Success = False
                        response.ErrorMessage = "Search query blocked: appears to contain personal or confidential data."
                        ToolingFileLogger.LogWarn("Internal search tool: query blocked by PII filter.",
                        details:=$"CallId={toolCall.CallId}; Pattern='{piiPattern}'")
                        context.Log("  ⚠ Search query blocked — contains data that appears personal or confidential.", "warn")
                        Return response
                    End If
                Next
            End If

            ' Clamp max_results to server limit (INI_ISearch_Tries)
            Dim maxResults As Integer = INI_ISearch_Results
            If toolCall.Arguments.ContainsKey("max_results") Then
                Dim requested As Integer
                If Integer.TryParse(toolCall.Arguments("max_results")?.ToString(), requested) AndAlso requested > 0 Then
                    maxResults = Math.Min(requested, INI_ISearch_Tries)
                End If
            End If

            ' Clamp max_depth to server limit (INI_ISearch_MaxDepth)
            Dim maxDepth As Integer = 0
            If toolCall.Arguments.ContainsKey("max_depth") Then
                Dim requested As Integer
                If Integer.TryParse(toolCall.Arguments("max_depth")?.ToString(), requested) AndAlso requested >= 0 Then
                    maxDepth = Math.Min(requested, INI_ISearch_MaxDepth)
                End If
            End If

            context.Log($"Internet search: query='{query}', max_results={maxResults}, max_depth={maxDepth}")
            ToolingFileLogger.LogStep($"Search query: '{query}'; max_results={maxResults}; max_depth={maxDepth}; engine={INI_ISearch_Name}")

            ' ── Perform the HTTP search request ──────────────────────────
            Dim searchUrl As String = INI_ISearch_URL & Uri.EscapeDataString(query)
            context.Log($"  Search URL: {searchUrl}")

            cancellationToken.ThrowIfCancellationRequested()

            Dim searchResponse As String = ""
            Using httpClient As New System.Net.Http.HttpClient()
                httpClient.DefaultRequestHeaders.UserAgent.ParseAdd(
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36")
                httpClient.Timeout = TimeSpan.FromSeconds(30)

                searchResponse = Await httpClient.GetStringAsync(searchUrl)
            End Using

            If String.IsNullOrWhiteSpace(searchResponse) Then
                response.Success = False
                response.ErrorMessage = "Search engine returned an empty response."
                ToolingFileLogger.LogWarn("Internal search tool: empty search response.",
                    details:=$"searchUrl={searchUrl}")
                Return response
            End If

            ' ── Extract unique URLs using response masks ─────────────────
            Dim urlPattern As String = Regex.Escape(INI_ISearch_ResponseMask1) & "(.*?)" & Regex.Escape(INI_ISearch_ResponseMask2)
            Dim matches As MatchCollection = Regex.Matches(searchResponse, urlPattern)

            Dim extractedUrls As New List(Of String)()
            For Each m As Match In matches
                Dim rawUrl As String = m.Groups(1).Value
                Dim decodedUrl As String = System.Net.WebUtility.UrlDecode(rawUrl.Replace(INI_ISearch_ResponseMask1, ""))

                If Not extractedUrls.Contains(decodedUrl) AndAlso IsSafeWebUrl(decodedUrl) Then
                    extractedUrls.Add(decodedUrl)
                End If

                If extractedUrls.Count >= INI_ISearch_Tries Then Exit For
            Next

            context.Log($"  Extracted {extractedUrls.Count} unique URL(s) from search results")
            ToolingFileLogger.LogStep($"Extracted URLs: {extractedUrls.Count}")

            If extractedUrls.Count = 0 Then
                response.Success = False
                response.ErrorMessage = "No result URLs could be extracted from the search engine response."
                ToolingFileLogger.LogWarn("Internal search tool: no URLs extracted.",
                    details:=$"searchUrl={searchUrl}; ResponseMask1='{INI_ISearch_ResponseMask1}'; ResponseMask2='{INI_ISearch_ResponseMask2}'")
                Return response
            End If

            ' ── Fetch content from each result URL ───────────────────────
            Dim results As New Text.StringBuilder()
            Dim visitedUrls As New List(Of String)()
            Dim resultIndex As Integer = 0

            ' Header: report the search query and engine
            results.AppendLine($"<SEARCH_QUERY>{query}</SEARCH_QUERY>")
            results.AppendLine($"<SEARCH_ENGINE>{If(INI_ISearch_Name, "Search")}</SEARCH_ENGINE>")
            results.AppendLine()

            For Each url In extractedUrls
                If resultIndex >= maxResults Then Exit For
                cancellationToken.ThrowIfCancellationRequested()

                Try
                    context.Log($"  Fetching result: {url}")
                    visitedUrls.Add(url)

                    Dim content As String = ""

                    If UseWebView2 Then
                        Dim pageResult = Await RetrieveWebsiteContent_WebView2(url, ISearch_MaxChars)
                        content = If(pageResult?.TextContent, "")
                    Else
                        Using httpClient As New System.Net.Http.HttpClient()
                            httpClient.DefaultRequestHeaders.UserAgent.ParseAdd(
                                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
                            httpClient.Timeout = TimeSpan.FromSeconds(30)
                            content = Await RetrieveWebsiteContent(url, maxDepth, httpClient)
                        End Using
                    End If

                    ' Apply character cap (ISearch_MaxChars) for WebView2 results that exceed it
                    If Not String.IsNullOrWhiteSpace(content) AndAlso ISearch_MaxChars > 0 AndAlso content.Length > ISearch_MaxChars Then
                        content = content.Substring(0, ISearch_MaxChars)
                    End If

                    ' Discard noise (pages shorter than ISearch_MinChars)
                    If Not String.IsNullOrWhiteSpace(content) AndAlso content.Length >= ISearch_MinChars Then
                        resultIndex += 1
                        results.AppendLine($"<SEARCHRESULT_{resultIndex}_URL>{url}</SEARCHRESULT_{resultIndex}_URL>")
                        results.AppendLine($"<SEARCHRESULT_{resultIndex}>")
                        results.AppendLine(content)
                        results.AppendLine($"</SEARCHRESULT_{resultIndex}>")
                        results.AppendLine()
                        context.Log($"  Result #{resultIndex}: {content.Length} chars from {url}")
                    Else
                        Dim charCount = If(content Is Nothing, 0, content.Length)
                        context.Log($"  Skipped (too short: {charCount} chars, min {ISearch_MinChars}): {url}")
                        ToolingFileLogger.LogStep($"Search result skipped (too short: {charCount} < {ISearch_MinChars}): {url}")
                    End If

                Catch ex As OperationCanceledException
                    Throw ' Re-throw cancellation
                Catch ex As Exception
                    context.Log($"  Error fetching {url}: {ex.Message}")
                    ToolingFileLogger.LogError("Internal search tool fetch error.", details:=$"url={url}", ex:=ex)
                End Try
            Next

            ' Footer: report all visited URLs for transparency
            results.AppendLine("<URLS_VISITED>")
            For Each vUrl In visitedUrls
                results.AppendLine($"  {vUrl}")
            Next
            results.AppendLine("</URLS_VISITED>")

            context.Log($"Search complete: {resultIndex} qualifying result(s) from {visitedUrls.Count} URL(s) visited")

            response.Response = results.ToString()
            response.Success = True

        Catch ex As OperationCanceledException
            response.Success = False
            response.ErrorMessage = "Operation was cancelled"
            ToolingFileLogger.LogWarn("Internal search tool cancelled.")

        Catch ex As System.Net.Http.HttpRequestException
            response.Success = False
            response.ErrorMessage = $"Search HTTP error: {ex.Message}"
            ToolingFileLogger.LogError("Internal search tool HTTP error.", ex:=ex)

        Catch ex As TaskCanceledException
            response.Success = False
            response.ErrorMessage = $"Search request timed out: {ex.Message}"
            ToolingFileLogger.LogError("Internal search tool timeout.", ex:=ex)

        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            ToolingFileLogger.LogError("Internal search tool error.", ex:=ex)
        End Try

        Return response
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

    ''' <summary>
    ''' Removes knowledge-store source files from the temp directory that were
    ''' not cited by the LLM in its final response. Files produced by other tools
    ''' (process_word_document, merge_pdfs, etc.) are never affected.
    ''' Must be called BEFORE CollectResultAttachments so the directory scan
    ''' does not re-pick up uncited knowledge files.
    ''' </summary>
    Private Sub RemoveUncitedKnowledgeSourceCopies(llmResponseText As String)
        If _apKnowledgeSourceCopies.Count = 0 Then Return
        If String.IsNullOrWhiteSpace(llmResponseText) Then Return

        Dim toRemove As New List(Of String)()

        For Each filePath In _apKnowledgeSourceCopies
            Dim fileName = Path.GetFileName(filePath)
            Dim baseName = Path.GetFileNameWithoutExtension(filePath)

            Dim cited = llmResponseText.IndexOf(fileName, StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                    (baseName.Length >= 4 AndAlso llmResponseText.IndexOf(baseName, StringComparison.OrdinalIgnoreCase) >= 0)

            If Not cited Then toRemove.Add(filePath)
        Next

        ' Safety: if ALL knowledge files are uncited, keep them (LLM may have paraphrased)
        If toRemove.Count = _apKnowledgeSourceCopies.Count Then Return

        For Each uncitedPath In toRemove
            Try
                If File.Exists(uncitedPath) Then File.Delete(uncitedPath)
            Catch
            End Try
            _apKnowledgeSourceCopies.Remove(uncitedPath)

            ' Also remove from OutputFiles so CollectResultAttachments doesn't find it
            If _apCurrentAttachments IsNot Nothing Then
                For Each att In _apCurrentAttachments
                    If att.OutputFiles IsNot Nothing Then att.OutputFiles.Remove(uncitedPath)
                Next
            End If
        Next
    End Sub




    ''' <summary>
    ''' Filters result file paths to only those whose filename or base name appears in
    ''' the LLM's final response text. Prevents uncited knowledge-store source files
    ''' from being delivered. Deletes uncited files from disk so they are not picked up
    ''' by the fallback scan in CollectResultAttachments. Returns all files as a safety
    ''' fallback if none matched (the LLM may have paraphrased).
    ''' </summary>
    Friend Shared Function FilterAttachmentsByCitation(resultFiles As List(Of String), llmResponseText As String) As List(Of String)
        If resultFiles Is Nothing OrElse resultFiles.Count = 0 Then Return If(resultFiles, New List(Of String)())
        If String.IsNullOrWhiteSpace(llmResponseText) Then Return resultFiles

        Dim cited As New List(Of String)()
        Dim uncited As New List(Of String)()

        For Each filePath In resultFiles
            Dim fileName = Path.GetFileName(filePath)
            Dim baseName = Path.GetFileNameWithoutExtension(filePath)

            ' Match full filename or base name (≥4 chars to avoid false positives)
            If llmResponseText.IndexOf(fileName, StringComparison.OrdinalIgnoreCase) >= 0 OrElse
           (baseName.Length >= 4 AndAlso llmResponseText.IndexOf(baseName, StringComparison.OrdinalIgnoreCase) >= 0) Then
                cited.Add(filePath)
            Else
                uncited.Add(filePath)
            End If
        Next

        ' Safety fallback: if nothing matched, LLM may have paraphrased — keep all
        If cited.Count = 0 Then Return resultFiles

        ' Remove uncited files from temp dir so CollectResultAttachments's
        ' fallback directory scan doesn't re-pick them up
        For Each uncitedPath In uncited
            Try
                If File.Exists(uncitedPath) Then File.Delete(uncitedPath)
            Catch
            End Try
        Next

        Return cited
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
    Private Shared Function IsSkipToolErrorHandling(toolConfig As ModelConfig) As Boolean
        Return toolConfig IsNot Nothing AndAlso
               String.Equals(If(toolConfig.ToolErrorHandling, "skip"), "skip", StringComparison.OrdinalIgnoreCase)
    End Function

    Private Function IsStructuredErrorToolResponse(toolResponse As ToolResponse) As Boolean
        If toolResponse Is Nothing OrElse String.IsNullOrWhiteSpace(toolResponse.Response) Then
            Return False
        End If

        Dim errorCode As String = ""
        Dim resultKind As String = ""
        Return SharedLibrary.Agents.SubAgentRuntimeHardening.TryGetEnvelopeErrorInfo(
            toolResponse.Response,
            errorCode,
            resultKind)
    End Function

    Private Function IsSkippedStructuredAgentFailure(toolCall As ToolCall,
                                                     toolResponse As ToolResponse,
                                                     toolConfig As ModelConfig,
                                                     subAgentMode As Boolean) As Boolean
        If subAgentMode Then Return False
        If toolCall Is Nothing OrElse String.IsNullOrWhiteSpace(toolCall.ToolName) Then Return False
        If Not toolCall.ToolName.StartsWith("agent_", StringComparison.OrdinalIgnoreCase) Then Return False
        If toolResponse Is Nothing OrElse toolResponse.Success Then Return False
        If Not IsSkipToolErrorHandling(toolConfig) Then Return False
        Return IsStructuredErrorToolResponse(toolResponse)
    End Function



End Class
