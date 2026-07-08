' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: KnowledgeTriggerHelper.vb
' Purpose:
'   Parses Knowledge Store inline triggers such as `(kb)` and `(kb:...)`,
'   resolves matching Knowledge Store content, and assembles prompt-ready
'   knowledge blocks for injection into downstream AI workflows.
'
' Responsibilities:
'   - Detect simple and parameterized Knowledge Store triggers in user prompts.
'   - Parse trigger syntax into structured Knowledge Store requests.
'   - Resolve store-, tag-, and keyword-filtered content from active stores.
'   - Remove or normalize trigger syntax in prompts where needed.
'   - Return bounded, prompt-ready knowledge blocks with source context.
'
' Notes:
'   - This helper bridges user-facing `(kb:...)` syntax and the underlying
'     Knowledge Store retrieval pipeline.
'   - It is used by flows such as Freestyle and other Knowledge Store-enabled
'     prompt entry points.
' =============================================================================


Option Strict On
Option Explicit On

Imports System.IO
Imports System.Security
Imports System.Text
Imports System.Text.RegularExpressions
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary


    ''' <summary>
    ''' Parses (kb:...) triggers from user prompts and resolves them to Knowledge Store content.
    ''' </summary>
    Public Class KnowledgeTriggerHelper

#Region "Constants"

        Public Const MaxRelevantExtractDocuments As Integer = 4
        Private Const MaxCharsPerRelevantExtractDocument As Integer = 30000


        ''' <summary>Simple trigger for Knowledge Store invocation.</summary>
        Public Const KbTrigger As String = "(kb)"

        ''' <summary>Trigger prefix with parameter for filtered invocation.</summary>
        Public Const KbTriggerPrefix As String = "(kb:"

        ''' <summary>Trigger suffix.</summary>
        Private Const KbTriggerSuffix As String = ")"

        ''' <summary>Explicit store prefix inside (kb:...).</summary>
        Private Const StorePrefix As String = "store:"

        ''' <summary>Explicit tag prefix inside (kb:...).</summary>
        Private Const TagPrefix As String = "tag:"

        Private Const MaxDocuments As Integer = 10
        Private Const MaxTotalChars As Integer = 200000

        Private Const KnowledgeOpenTag As String = "<KNOWLEDGESTORE>"
        Private Const KnowledgeCloseTag As String = "</KNOWLEDGESTORE>"


#End Region

#Region "Data Model"


        Public Class KnowledgeResolveOptions
            Public Property TaskPrompt As String = ""
            Public Property IncludeRelevantExtracts As Boolean = False
            Public Property IncludeFullDocumentContent As Boolean = False
            Public Property MaxResults As Integer = 0
            Public Property ForceSemanticSearch As Boolean = False
        End Class


        ''' <summary>
        ''' Describes a parsed Knowledge Store request extracted from a user prompt.
        ''' </summary>
        Public Class KnowledgeRequest
            Public Property LoadAll As Boolean
            Public Property Tags As String()
            Public Property StoreName As String
            Public Property SearchQuery As String
            Public Property RawTrigger As String = ""
            Public Property OriginalParameter As String = ""
            Public Property HasExplicitStoreFilter As Boolean = False
            Public Property HasExplicitTagFilter As Boolean = False
        End Class

#End Region

#Region "Trigger Detection"

        ''' <summary>
        ''' Returns True if the prompt contains a Knowledge Store trigger.
        ''' </summary>
        Public Shared Function HasKnowledgeTrigger(prompt As String) As Boolean
            If String.IsNullOrWhiteSpace(prompt) Then Return False
            Return prompt.IndexOf(KbTrigger, StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                   prompt.IndexOf(KbTriggerPrefix, StringComparison.OrdinalIgnoreCase) >= 0
        End Function

#End Region

#Region "Trigger Parsing"

        ''' <summary>
        ''' Parses a (kb:...) trigger from the prompt and returns a KnowledgeRequest.
        ''' Returns Nothing if no trigger is found.
        ''' </summary>
        Public Shared Function TryParseKnowledgeTrigger(prompt As String) As KnowledgeRequest
            If String.IsNullOrWhiteSpace(prompt) Then Return Nothing

            ' Check for parameterized trigger first: (kb:...)
            Dim prefixIdx = prompt.IndexOf(KbTriggerPrefix, StringComparison.OrdinalIgnoreCase)
            If prefixIdx >= 0 Then
                Dim closeIdx = prompt.IndexOf(KbTriggerSuffix, prefixIdx + KbTriggerPrefix.Length, StringComparison.OrdinalIgnoreCase)
                If closeIdx > prefixIdx Then
                    Dim rawTrigger = prompt.Substring(prefixIdx, closeIdx - prefixIdx + KbTriggerSuffix.Length)
                    Dim parameter = prompt.Substring(prefixIdx + KbTriggerPrefix.Length, closeIdx - prefixIdx - KbTriggerPrefix.Length).Trim()

                    If String.IsNullOrWhiteSpace(parameter) Then
                        Return New KnowledgeRequest() With {
                            .LoadAll = True,
                            .RawTrigger = rawTrigger,
                            .OriginalParameter = ""
                        }
                    End If

                    Dim parsed = ParseParameter(parameter, rawTrigger)
                    parsed.OriginalParameter = parameter
                    Return parsed
                End If
            End If

            ' Check for simple trigger: (kb)
            Dim simpleIdx = prompt.IndexOf(KbTrigger, StringComparison.OrdinalIgnoreCase)
            If simpleIdx >= 0 Then
                Return New KnowledgeRequest() With {
                    .LoadAll = True,
                    .RawTrigger = KbTrigger,
                    .OriginalParameter = ""
                }
            End If

            Return Nothing
        End Function

        ''' <summary>
        ''' Parses the parameter string inside (kb:...).
        ''' Supports: "store:X", "tag:Y", "store:X tag:Y", plain CSV tags/store names.
        ''' Store and tag values may be quoted to allow spaces: store:"My Store"
        ''' </summary>
        Private Shared Function ParseParameter(parameter As String, rawTrigger As String) As KnowledgeRequest
            Dim req As New KnowledgeRequest() With {
                .RawTrigger = rawTrigger,
                .OriginalParameter = parameter
            }

            Dim hasStorePrefix = parameter.IndexOf(StorePrefix, StringComparison.OrdinalIgnoreCase) >= 0
            Dim hasTagPrefix = parameter.IndexOf(TagPrefix, StringComparison.OrdinalIgnoreCase) >= 0

            req.HasExplicitStoreFilter = hasStorePrefix
            req.HasExplicitTagFilter = hasTagPrefix

            If hasStorePrefix OrElse hasTagPrefix Then
                ' Tokenize respecting quoted values: store:"My Store" tag:"Some Tag" free text
                Dim tokens = TokenizeWithQuotes(parameter)
                For Each token In tokens
                    If token.StartsWith(StorePrefix, StringComparison.OrdinalIgnoreCase) Then
                        req.StoreName = Unquote(token.Substring(StorePrefix.Length)).Trim()
                    ElseIf token.StartsWith(TagPrefix, StringComparison.OrdinalIgnoreCase) Then
                        Dim tagVal = Unquote(token.Substring(TagPrefix.Length)).Trim()
                        Dim tagParts = tagVal.Split(","c).
                            Select(Function(t) t.Trim()).
                            Where(Function(t) Not String.IsNullOrWhiteSpace(t)).
                            ToArray()
                        req.Tags = tagParts
                    Else
                        req.SearchQuery = If(String.IsNullOrWhiteSpace(req.SearchQuery), token, req.SearchQuery & " " & token)
                    End If
                Next

                Return req
            End If

            Dim csvParts = parameter.Split(","c).
                Select(Function(p) p.Trim()).
                Where(Function(p) Not String.IsNullOrWhiteSpace(p)).
                ToArray()

            If csvParts.Length > 0 Then
                req.Tags = csvParts
                req.SearchQuery = parameter
            End If

            Return req
        End Function

        ''' <summary>
        ''' Splits a string by spaces but keeps quoted segments (double quotes) as single tokens.
        ''' E.g. <c>store:"My Store" hello</c> → <c>{"store:My Store", "hello"}</c>.
        ''' The prefix (e.g. "store:") is preserved and the quotes are stripped from the value portion.
        ''' </summary>
        Public Shared Function TokenizeWithQuotes(input As String) As List(Of String)
            Dim result As New List(Of String)()
            If String.IsNullOrWhiteSpace(input) Then Return result

            ' Match: prefix:"quoted value" | prefix:'quoted value' | unquoted_token
            Dim pattern As String = "(\w+:""[^""]*""|\w+:'[^']*'|\S+)"
            For Each m As Match In Regex.Matches(input, pattern)
                Dim token = m.Value

                ' For prefixed tokens like store:"My Store", strip the quotes from the value part
                Dim colonIdx = token.IndexOf(":"c)
                If colonIdx >= 0 AndAlso colonIdx < token.Length - 1 Then
                    Dim prefix = token.Substring(0, colonIdx + 1)
                    Dim value = token.Substring(colonIdx + 1)
                    result.Add(prefix & Unquote(value))
                Else
                    result.Add(token)
                End If
            Next

            Return result
        End Function

        ''' <summary>
        ''' Removes surrounding double or single quotes from a string if present.
        ''' </summary>
        Private Shared Function Unquote(value As String) As String
            If String.IsNullOrEmpty(value) OrElse value.Length < 2 Then Return value
            If (value.StartsWith("""", StringComparison.Ordinal) AndAlso value.EndsWith("""", StringComparison.Ordinal)) OrElse
               (value.StartsWith("'", StringComparison.Ordinal) AndAlso value.EndsWith("'", StringComparison.Ordinal)) Then
                Return value.Substring(1, value.Length - 2)
            End If
            Return value
        End Function

        ''' <summary>
        ''' Removes the Knowledge Store trigger from the prompt and returns the cleaned text.
        ''' </summary>
        Public Shared Function StripKnowledgeTrigger(prompt As String, request As KnowledgeRequest) As String
            If request Is Nothing OrElse String.IsNullOrWhiteSpace(request.RawTrigger) Then Return prompt
            Dim idx = prompt.IndexOf(request.RawTrigger, StringComparison.OrdinalIgnoreCase)
            If idx < 0 Then Return prompt
            Dim result = (prompt.Substring(0, idx) & prompt.Substring(idx + request.RawTrigger.Length)).Trim()
            result = Regex.Replace(result, " {2,}", " ")
            Return result
        End Function

#End Region

#Region "Knowledge Resolution"

        Private Shared Function ShouldUseSemanticFirst(request As KnowledgeRequest) As Boolean
            If request Is Nothing Then Return False
            If request.LoadAll Then Return False
            If String.IsNullOrWhiteSpace(request.SearchQuery) Then Return False

            If request.HasExplicitTagFilter Then Return False

            Dim parameter = If(request.OriginalParameter, "").Trim()

            If String.IsNullOrWhiteSpace(parameter) Then Return False

            If parameter.IndexOf(","c) >= 0 Then Return False

            If parameter.IndexOf(" "c) >= 0 Then Return True

            Return False
        End Function

        ''' <summary>
        ''' Resolves a KnowledgeRequest to assembled document content ready for LLM injection.
        ''' Synchronous compatibility wrapper over the async implementation.
        ''' </summary>
        Public Shared Function ResolveKnowledge(
        request As KnowledgeRequest,
        context As ISharedContext,
        Optional options As KnowledgeResolveOptions = Nothing) As (Content As String, StatusMessage As String)

            Return ResolveKnowledgeAsync(request, context, options).GetAwaiter().GetResult()
        End Function

        Private Shared Function FindWikiPagePathForEntry(store As KnowledgeStoreCatalog.KnowledgeStoreDefinition,
                                                         entry As KnowledgeStoreManager.KnowledgeEntry) As String
            If store Is Nothing OrElse entry Is Nothing Then Return ""

            Dim wikiDir = Path.Combine(store.ResolvedSourcePath, KnowledgeStoreCatalog.MetadataFolder, KnowledgeStoreCatalog.WikiFolder)
            If Not Directory.Exists(wikiDir) Then Return ""

            Dim normalizedEntryPath = NormalizePathForCompare(entry.FilePath)
            Dim normalizedTitle = NormalizeKey(If(entry.Title, ""))

            For Each file In Directory.GetFiles(wikiDir, "*.md", SearchOption.AllDirectories)
                Dim name = Path.GetFileName(file)
                If name.Equals(KnowledgeStoreCatalog.IndexFile, StringComparison.OrdinalIgnoreCase) Then Continue For
                If name.Equals(KnowledgeStoreCatalog.LogFile, StringComparison.OrdinalIgnoreCase) Then Continue For
                If name.Equals("health_report.md", StringComparison.OrdinalIgnoreCase) Then Continue For

                Try
                    Dim content = System.IO.File.ReadAllText(file, Encoding.UTF8)
                    Dim frontMatter = GetFrontMatter(content)

                    Dim pageTitle = GetFrontMatterScalar(frontMatter, "title")
                    If NormalizeKey(pageTitle) = normalizedTitle AndAlso Not String.IsNullOrWhiteSpace(normalizedTitle) Then
                        Return file
                    End If

                    Dim sourcePaths = GetFrontMatterList(frontMatter, "source_paths")
                    For Each sourcePath In sourcePaths
                        If NormalizePathForCompare(sourcePath) = normalizedEntryPath AndAlso
                           Not String.IsNullOrWhiteSpace(normalizedEntryPath) Then
                            Return file
                        End If
                    Next
                Catch
                End Try
            Next

            Return ""
        End Function

        Private Shared Function GetFrontMatter(content As String) As String
            If String.IsNullOrWhiteSpace(content) Then Return ""
            Dim normalized = content.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
            If Not normalized.StartsWith("---" & vbLf, StringComparison.Ordinal) Then Return ""

            Dim endMarker = vbLf & "---" & vbLf
            Dim endPos = normalized.IndexOf(endMarker, 4, StringComparison.Ordinal)
            If endPos < 0 Then Return ""

            Return normalized.Substring(4, endPos - 4)
        End Function

        Private Shared Function GetFrontMatterScalar(frontMatter As String, key As String) As String
            If String.IsNullOrWhiteSpace(frontMatter) Then Return ""

            For Each rawLine In frontMatter.Split({vbCrLf, vbLf, vbCr}, StringSplitOptions.None)
                Dim line = rawLine.Trim()
                If line.StartsWith(key & ":", StringComparison.OrdinalIgnoreCase) Then
                    Dim value = line.Substring(key.Length + 1).Trim()
                    If value.StartsWith("""", StringComparison.Ordinal) AndAlso
                       value.EndsWith("""", StringComparison.Ordinal) AndAlso
                       value.Length >= 2 Then
                        value = value.Substring(1, value.Length - 2)
                    End If
                    Return value.Trim()
                End If
            Next

            Return ""
        End Function

        Private Shared Function GetFrontMatterList(frontMatter As String, key As String) As List(Of String)
            Dim result As New List(Of String)()
            If String.IsNullOrWhiteSpace(frontMatter) Then Return result

            Dim lines = frontMatter.Split({vbCrLf, vbLf, vbCr}, StringSplitOptions.None)
            Dim inTarget As Boolean = False

            For Each rawLine In lines
                Dim line = rawLine.TrimEnd()

                If line.Trim().StartsWith(key & ":", StringComparison.OrdinalIgnoreCase) Then
                    inTarget = True
                    Continue For
                End If

                If inTarget Then
                    Dim trimmed = line.Trim()
                    If trimmed.StartsWith("- ", StringComparison.Ordinal) Then
                        Dim value = trimmed.Substring(2).Trim()
                        If value.StartsWith("""", StringComparison.Ordinal) AndAlso
                               value.EndsWith("""", StringComparison.Ordinal) AndAlso
                               value.Length >= 2 Then
                            value = value.Substring(1, value.Length - 2)
                        End If
                        ' Strip surrounding single quotes (YAML single-quoted scalars)
                        If value.StartsWith("'", StringComparison.Ordinal) AndAlso
                               value.EndsWith("'", StringComparison.Ordinal) AndAlso
                               value.Length >= 2 Then
                            value = value.Substring(1, value.Length - 2).Replace("''", "'")
                        End If
                        If Not String.IsNullOrWhiteSpace(value) Then result.Add(value)
                    ElseIf trimmed.Contains(":"c) Then
                        Exit For
                    End If
                End If
            Next

            Return result
        End Function

        Private Shared Function NormalizePathForCompare(pathValue As String) As String
            If String.IsNullOrWhiteSpace(pathValue) Then Return ""
            Try
                Dim expanded = SharedMethods.ExpandEnvironmentVariables(pathValue)
                Return Path.GetFullPath(expanded).Trim().ToUpperInvariant()
            Catch
                Return pathValue.Trim().ToUpperInvariant()
            End Try
        End Function

        Private Shared Function NormalizeKey(value As String) As String
            Return Regex.Replace(If(value, "").ToLowerInvariant(), "[^a-z0-9]+", "").Trim()
        End Function

#End Region

#Region "Knowledge Resolution"

        ''' <summary>
        ''' Resolves a KnowledgeRequest to assembled document content ready for LLM injection.
        ''' Uses the multi-store catalog and manifest infrastructure. If a search query is provided,
        ''' it uses semantic vector search for ranking.
        ''' </summary>
        Public Shared Async Function ResolveKnowledgeAsync(
        request As KnowledgeRequest,
        context As ISharedContext,
        Optional options As KnowledgeResolveOptions = Nothing) As Task(Of (Content As String, StatusMessage As String))

            If request Is Nothing Then Return ("", "")

            If Not KnowledgeStoreCatalog.IsConfigured(context) Then
                Return ("", "Knowledge Store is not configured.")
            End If

            Dim stores = KnowledgeStoreCatalog.GetActiveStores(context)
            If stores.Count = 0 Then
                Return ("", "No active Knowledge Stores found.")
            End If

            Dim effectiveMaxDocuments As Integer = MaxDocuments
            If options IsNot Nothing AndAlso options.MaxResults > 0 Then
                effectiveMaxDocuments = Math.Min(Math.Max(1, options.MaxResults), MaxDocuments)
            End If

            Dim useRelevantExtracts As Boolean =
        options IsNot Nothing AndAlso
        options.IncludeRelevantExtracts AndAlso
        Not String.IsNullOrWhiteSpace(options.TaskPrompt)

            Dim targetStores As List(Of KnowledgeStoreCatalog.KnowledgeStoreDefinition) = Nothing
            Dim resolvedStoreFromTag As Boolean = False

            If Not String.IsNullOrWhiteSpace(request.StoreName) Then
                Dim matches = KnowledgeStoreCatalog.GetStoresByName(request.StoreName, context)

                If matches.Count = 0 AndAlso request.HasExplicitStoreFilter Then
                    matches = TryConsumeStoreNamePartsFromSearchQuery(request, context)
                End If

                If matches.Count > 0 Then
                    targetStores = matches
                Else
                    Return ("", $"Knowledge Store '{request.StoreName}' not found.")
                End If
            ElseIf request.Tags IsNot Nothing AndAlso request.Tags.Length = 1 AndAlso Not request.LoadAll Then
                Dim singleVal = request.Tags(0)
                Dim storeMatches = KnowledgeStoreCatalog.GetStoresByName(singleVal, context)

                If storeMatches.Count > 0 Then
                    targetStores = storeMatches
                    resolvedStoreFromTag = True
                    request = New KnowledgeRequest() With {
                .LoadAll = True,
                .StoreName = singleVal,
                .RawTrigger = request.RawTrigger,
                .OriginalParameter = request.OriginalParameter,
                .HasExplicitStoreFilter = True
            }
                End If
            End If

            If targetStores Is Nothing Then
                targetStores = stores
            End If

            Dim forceSemanticSearch As Boolean =
    options IsNot Nothing AndAlso
    options.ForceSemanticSearch AndAlso
    Not request.LoadAll AndAlso
    Not String.IsNullOrWhiteSpace(request.SearchQuery)

            If forceSemanticSearch OrElse ShouldUseSemanticFirst(request) Then
                Dim semanticMatches = Await ResolveSemanticMatchesAsync(
            request:=request,
            context:=context,
            targetStores:=targetStores,
            maxResults:=effectiveMaxDocuments).ConfigureAwait(False)

                If semanticMatches.Count = 0 Then
                    Return ("", $"No semantic knowledge found matching '{request.SearchQuery}'.")
                End If

                Dim semanticContext As String
                If useRelevantExtracts Then
                    semanticContext = Await BuildKnowledgeContextWithRelevantSectionsAsync(
                semanticMatches,
                request,
                context,
                options,
                MaxTotalChars).ConfigureAwait(False)

                    If String.IsNullOrWhiteSpace(semanticContext) Then
                        semanticContext = KnowledgeQueryService.BuildKnowledgeContext(semanticMatches, MaxTotalChars)
                    End If
                Else
                    semanticContext = KnowledgeQueryService.BuildKnowledgeContext(semanticMatches, MaxTotalChars)
                End If

                If String.IsNullOrWhiteSpace(semanticContext) Then
                    Return ("", $"No readable semantic knowledge found matching '{request.SearchQuery}'.")
                End If

                Dim semanticMsg = $"Loaded semantic knowledge for '{request.SearchQuery}'"
                If Not String.IsNullOrWhiteSpace(request.StoreName) Then
                    semanticMsg &= $" in store '{request.StoreName}'"
                End If

                If useRelevantExtracts Then
                    semanticMsg &= " with verbatim task-relevant excerpts"
                End If

                Return (semanticContext, semanticMsg & ".")
            End If

            Dim allMatches As New List(Of Tuple(Of KnowledgeStoreManager.KnowledgeEntry, String, String))()

            For Each store In targetStores
                Dim manifest = KnowledgeStoreManifest.Load(store)
                Dim wikiDir = KnowledgeStoreCatalog.GetWikiPath(store, createIfMissing:=False)

                For Each entry In manifest.Entries
                    If String.IsNullOrWhiteSpace(entry.FilePath) Then Continue For

                    Dim include As Boolean = False

                    If request.LoadAll Then
                        include = True
                    ElseIf request.Tags IsNot Nothing AndAlso request.Tags.Length > 0 Then
                        If entry.Tags IsNot Nothing Then
                            For Each reqTag In request.Tags
                                If entry.Tags.Any(Function(t) t.Equals(reqTag, StringComparison.OrdinalIgnoreCase)) Then
                                    include = True
                                    Exit For
                                End If
                            Next
                        End If

                        If Not include AndAlso Not String.IsNullOrWhiteSpace(request.SearchQuery) Then
                            Dim sq = request.SearchQuery.ToLowerInvariant()
                            If If(entry.Title, "").ToLowerInvariant().Contains(sq) Then include = True
                            If Not include AndAlso entry.Keywords IsNot Nothing Then
                                If entry.Keywords.Any(Function(k) k.IndexOf(sq, StringComparison.OrdinalIgnoreCase) >= 0) Then
                                    include = True
                                End If
                            End If
                            If Not include AndAlso If(entry.Summary, "").ToLowerInvariant().Contains(sq) Then
                                include = True
                            End If
                        End If
                    Else
                        include = True
                    End If

                    If include Then
                        Dim wikiPath As String = FindWikiPagePathForEntry(store, entry)
                        allMatches.Add(Tuple.Create(entry, store.Name, wikiPath))
                    End If
                Next
            Next

            If allMatches.Count = 0 Then
                If Not request.LoadAll AndAlso Not String.IsNullOrWhiteSpace(request.SearchQuery) Then
                    Dim semanticMatches = Await ResolveSemanticMatchesAsync(
                request:=request,
                context:=context,
                targetStores:=targetStores,
                maxResults:=effectiveMaxDocuments).ConfigureAwait(False)

                    If semanticMatches.Count > 0 Then
                        Dim semanticContext As String
                        If useRelevantExtracts Then
                            semanticContext = Await BuildKnowledgeContextWithRelevantSectionsAsync(
                        semanticMatches,
                        request,
                        context,
                        options,
                        MaxTotalChars).ConfigureAwait(False)

                            If String.IsNullOrWhiteSpace(semanticContext) Then
                                semanticContext = KnowledgeQueryService.BuildKnowledgeContext(semanticMatches, MaxTotalChars)
                            End If
                        Else
                            semanticContext = KnowledgeQueryService.BuildKnowledgeContext(semanticMatches, MaxTotalChars)
                        End If

                        If Not String.IsNullOrWhiteSpace(semanticContext) Then
                            Dim semanticMsg = $"Loaded semantic knowledge for '{request.SearchQuery}'"
                            If Not String.IsNullOrWhiteSpace(request.StoreName) Then
                                semanticMsg &= $" in store '{request.StoreName}'"
                            End If
                            If useRelevantExtracts Then
                                semanticMsg &= " with verbatim task-relevant excerpts"
                            End If
                            Return (semanticContext, semanticMsg & ".")
                        End If
                    End If
                End If

                Dim filterDesc = ""
                If request.Tags IsNot Nothing Then filterDesc = String.Join(", ", request.Tags)
                If Not String.IsNullOrWhiteSpace(request.StoreName) Then filterDesc = $"store '{request.StoreName}'"
                If String.IsNullOrWhiteSpace(filterDesc) AndAlso Not String.IsNullOrWhiteSpace(request.SearchQuery) Then
                    filterDesc = request.SearchQuery
                End If

                Return ("", $"No Knowledge Store documents found matching '{filterDesc}'.")
            End If

            Dim entriesToLoad = allMatches.Take(effectiveMaxDocuments).ToList()
            Dim richMatches As New List(Of KnowledgeQueryService.KnowledgeMatch)()

            For Each item In entriesToLoad
                Dim entry = item.Item1
                Dim sName = item.Item2
                Dim wikiPath = item.Item3

                richMatches.Add(New KnowledgeQueryService.KnowledgeMatch With {
            .Entry = entry,
            .StoreName = sName,
            .WikiPagePath = wikiPath,
            .Title = If(entry.Title, Path.GetFileNameWithoutExtension(If(entry.FilePath, "Unknown"))),
            .Summary = If(entry.Summary, ""),
            .SourcePath = ResolveSourcePath(entry.FilePath, stores.FirstOrDefault(Function(s) s.Name.Equals(sName, StringComparison.OrdinalIgnoreCase))),
            .MatchReason = "kb-trigger"
        })
            Next

            Dim richContext As String
            If useRelevantExtracts Then
                richContext = Await BuildKnowledgeContextWithRelevantSectionsAsync(
            richMatches,
            request,
            context,
            options,
            MaxTotalChars).ConfigureAwait(False)

                If String.IsNullOrWhiteSpace(richContext) Then
                    richContext = KnowledgeQueryService.BuildKnowledgeContext(richMatches, MaxTotalChars)
                End If
            Else
                richContext = KnowledgeQueryService.BuildKnowledgeContext(richMatches, MaxTotalChars)
            End If

            If String.IsNullOrWhiteSpace(richContext) Then
                Return ("", "Knowledge Store matches were found, but no readable content could be assembled.")
            End If

            Dim statusMsg = $"Loaded {entriesToLoad.Count} document(s)"
            If Not String.IsNullOrWhiteSpace(request.StoreName) Then
                statusMsg &= $" from store '{request.StoreName}'"
            End If
            If request.Tags IsNot Nothing AndAlso request.Tags.Length > 0 AndAlso Not resolvedStoreFromTag Then
                statusMsg &= $" for '{String.Join(", ", request.Tags)}'"
            End If
            If useRelevantExtracts Then
                statusMsg &= " with verbatim task-relevant excerpts"
            End If
            statusMsg &= "."

            Return (richContext, statusMsg)
        End Function

        Private Shared Function TryConsumeStoreNamePartsFromSearchQuery(
        request As KnowledgeRequest,
        context As ISharedContext) As List(Of KnowledgeStoreCatalog.KnowledgeStoreDefinition)

            Dim result As New List(Of KnowledgeStoreCatalog.KnowledgeStoreDefinition)()

            If request Is Nothing OrElse
       String.IsNullOrWhiteSpace(request.StoreName) OrElse
       String.IsNullOrWhiteSpace(request.SearchQuery) Then
                Return result
            End If

            Dim tokens = TokenizeWithQuotes(request.SearchQuery)
            If tokens.Count = 0 Then
                Return result
            End If

            Dim maxParts As Integer = Math.Min(tokens.Count, 6)

            For i As Integer = maxParts To 1 Step -1
                Dim candidateStoreName = (request.StoreName.Trim() & " " & String.Join(" ", tokens.Take(i))).Trim()
                Dim storeMatches = KnowledgeStoreCatalog.GetStoresByName(candidateStoreName, context)

                If storeMatches.Count > 0 Then
                    request.StoreName = candidateStoreName
                    request.SearchQuery = String.Join(" ", tokens.Skip(i)).Trim()
                    Return storeMatches
                End If
            Next

            Return result
        End Function

        Private Shared Async Function ResolveSemanticMatchesAsync(
        request As KnowledgeRequest,
        context As ISharedContext,
        targetStores As List(Of KnowledgeStoreCatalog.KnowledgeStoreDefinition),
        maxResults As Integer) As Task(Of List(Of KnowledgeQueryService.KnowledgeMatch))

            Dim quoteIfNeeded As Func(Of String, String) =
        Function(value As String) As String
            Dim trimmed As String = If(value, "").Trim()
            If trimmed = "" Then Return ""

            If trimmed.IndexOf(" "c) >= 0 OrElse trimmed.IndexOf(ControlChars.Tab) >= 0 Then
                Return """" & trimmed.Replace("""", """""") & """"
            End If

            Return trimmed
        End Function

            Dim queryParts As New List(Of String)()

            If request.Tags IsNot Nothing AndAlso request.Tags.Length > 0 Then
                For Each tag In request.Tags
                    Dim cleanedTag As String = If(tag, "").Trim()
                    If cleanedTag <> "" Then
                        queryParts.Add("tag:" & quoteIfNeeded(cleanedTag))
                    End If
                Next
            End If

            If Not String.IsNullOrWhiteSpace(request.SearchQuery) Then
                queryParts.Add(request.SearchQuery.Trim())
            End If

            Dim queryText As String = String.Join(" ", queryParts).Trim()

            If String.IsNullOrWhiteSpace(queryText) Then
                Return New List(Of KnowledgeQueryService.KnowledgeMatch)()
            End If

            Dim matches = Await KnowledgeQueryService.ResolveQueryAsync(
        query:=queryText,
        context:=context,
        maxResults:=Math.Max(maxResults * 4, 12)).ConfigureAwait(False)

            If targetStores IsNot Nothing AndAlso targetStores.Count > 0 Then
                Dim allowedStoreNames = New HashSet(Of String)(
            targetStores.
                Select(Function(s) If(s.Name, "").Trim()).
                Where(Function(n) Not String.IsNullOrWhiteSpace(n)),
            StringComparer.OrdinalIgnoreCase)

                matches = matches.
            Where(Function(m) allowedStoreNames.Contains(If(m.StoreName, "").Trim())).
            ToList()
            End If

            Return matches.
        OrderByDescending(Function(m) m.Score).
        ThenBy(Function(m) m.Title, StringComparer.OrdinalIgnoreCase).
        Take(maxResults).
        ToList()
        End Function


        Private Shared Async Function BuildKnowledgeContextWithRelevantSectionsAsync(
        matches As List(Of KnowledgeQueryService.KnowledgeMatch),
        request As KnowledgeRequest,
        context As ISharedContext,
        options As KnowledgeResolveOptions,
        maxTotalChars As Integer) As Task(Of String)

            If matches Is Nothing OrElse matches.Count = 0 Then Return ""

            Dim sb As New StringBuilder()
            sb.AppendLine("<KNOWLEDGESTORE>")
            sb.AppendLine("The following source-backed Knowledge Store documents have been provided for context.")
            sb.AppendLine("IMPORTANT: When citing information, ALWAYS prefer the original source file link over the wiki page link.")
            sb.AppendLine("If a sourcePath is provided and non-empty, ALWAYS cite it as: [Source Title](sourcePath)")
            sb.AppendLine("Only fall back to wikiPath if no sourcePath is available for that document.")
            sb.AppendLine("Do NOT invent links. Use only the explicit sourcePath or wikiPath values provided in the KSDOCUMENT attributes.")
            sb.AppendLine($"When available, the resolver has read up to {MaxRelevantExtractDocuments} top-matching documents and extracted only verbatim passages relevant to the user's task prompt.")
            sb.AppendLine()

            Dim totalChars As Integer = 0
            Dim docsReadForExtracts As Integer = 0

            For Each m In matches
                Dim remaining = maxTotalChars - totalChars
                If remaining <= 0 Then Exit For

                Dim documentContent As String = ReadKnowledgeMatchContent(m)
                Dim readable As Boolean = Not String.IsNullOrWhiteSpace(documentContent)
                Dim extracted As String = ""
                Dim note As String = ""
                Dim bodyContent As String = ""

                If readable AndAlso
           docsReadForExtracts < MaxRelevantExtractDocuments AndAlso
           options IsNot Nothing AndAlso
           options.IncludeRelevantExtracts AndAlso
           Not String.IsNullOrWhiteSpace(options.TaskPrompt) Then

                    docsReadForExtracts += 1

                    extracted = Await ExtractRelevantSectionsAsync(
                match:=m,
                documentContent:=documentContent,
                taskPrompt:=options.TaskPrompt,
                retrievalPrompt:=If(request?.SearchQuery, ""),
                context:=context).ConfigureAwait(False)
                End If

                If Not String.IsNullOrWhiteSpace(extracted) Then
                    bodyContent =
                "<KSRELEVANTSECTIONS basis=""task_prompt"">" & vbCrLf &
                extracted.Trim() & vbCrLf &
                "</KSRELEVANTSECTIONS>"

                    If options IsNot Nothing AndAlso options.IncludeFullDocumentContent Then
                        bodyContent &= vbCrLf &
                               "<KSFULLDOCUMENT>" & vbCrLf &
                               documentContent.Trim() & vbCrLf &
                               "</KSFULLDOCUMENT>"
                    End If
                ElseIf readable Then
                    bodyContent = documentContent
                ElseIf Not String.IsNullOrWhiteSpace(m.Summary) Then
                    note = "The source file could not be read. Only stored summary metadata is available for this match."
                    bodyContent =
                "<KSMETADATAONLY>" & vbCrLf &
                m.Summary.Trim() & vbCrLf &
                "</KSMETADATAONLY>"
                Else
                    note = "The source file and wiki page could not be read for this match."
                    bodyContent =
                "<KSMETADATAONLY>" & vbCrLf &
                "No readable document content was available for this match." & vbCrLf &
                "</KSMETADATAONLY>"
                End If

                If String.IsNullOrWhiteSpace(bodyContent) Then Continue For

                If bodyContent.Length > remaining Then
                    bodyContent = bodyContent.Substring(0, remaining) & vbCrLf & "[... truncated ...]"
                End If

                sb.AppendLine(BuildKnowledgeDocumentBlock(
            match:=m,
            bodyContent:=bodyContent,
            readable:=readable,
            excerpted:=Not String.IsNullOrWhiteSpace(extracted),
            note:=note))
                sb.AppendLine()

                totalChars += bodyContent.Length
            Next

            sb.AppendLine("</KNOWLEDGESTORE>")
            Return sb.ToString()
        End Function

        Private Shared Function ReadKnowledgeMatchContent(match As KnowledgeQueryService.KnowledgeMatch) As String
            If match Is Nothing Then Return ""

            Dim content As String = ""

            If Not String.IsNullOrWhiteSpace(match.SourcePath) Then
                content = TryReadKnowledgeFileContent(match.SourcePath)
            End If

            If String.IsNullOrWhiteSpace(content) AndAlso match.Entry IsNot Nothing Then
                content = KnowledgeStoreManager.GetContent(match.Entry)
            End If

            If String.IsNullOrWhiteSpace(content) AndAlso
       Not String.IsNullOrWhiteSpace(match.WikiPagePath) AndAlso
       File.Exists(match.WikiPagePath) Then
                Try
                    content = File.ReadAllText(match.WikiPagePath, Encoding.UTF8)
                Catch
                End Try
            End If

            Return If(content, "")
        End Function

        Private Shared Function TryReadKnowledgeFileContent(filePath As String) As String
            If String.IsNullOrWhiteSpace(filePath) Then Return ""

            Try
                Dim expandedPath = SharedMethods.ExpandEnvironmentVariables(filePath)
                If String.IsNullOrWhiteSpace(expandedPath) OrElse Not File.Exists(expandedPath) Then
                    Return ""
                End If

                Dim ext = Path.GetExtension(expandedPath).ToLowerInvariant()

                Select Case ext
                    Case ".txt", ".ini", ".csv", ".log", ".json", ".xml", ".html", ".htm",
                 ".md", ".yaml", ".yml", ".vb", ".cs", ".js", ".ts", ".py",
                 ".java", ".cpp", ".c", ".h", ".sql"
                        Return File.ReadAllText(expandedPath, Encoding.UTF8)

                    Case ".docx"
                        Return SharedMethods.ReadDocxSandboxed(expandedPath)

                    Case ".pdf"
                        Dim t = SharedMethods.ReadPdfAsText(expandedPath, ReturnErrorInsteadOfEmpty:=False, DoOCR:=False, AskUser:=False)
                        t.Wait()
                        Return If(t.Result, "")

                    Case ".rtf"
                        Try
                            Using rtb As New System.Windows.Forms.RichTextBox()
                                rtb.LoadFile(expandedPath, System.Windows.Forms.RichTextBoxStreamType.RichText)
                                Return rtb.Text
                            End Using
                        Catch
                            Return File.ReadAllText(expandedPath, Encoding.UTF8)
                        End Try

                    Case ".xlsx"
                        Return SharedMethods.ReadXlsxSandboxed(expandedPath)

                    Case ".pptx"
                        Return SharedMethods.ReadPptxSandboxed(expandedPath)

                    Case Else
                        Return File.ReadAllText(expandedPath, Encoding.UTF8)
                End Select
            Catch
                Return ""
            End Try
        End Function

        Private Shared Async Function ExtractRelevantSectionsAsync(
        match As KnowledgeQueryService.KnowledgeMatch,
        documentContent As String,
        taskPrompt As String,
        retrievalPrompt As String,
        context As ISharedContext) As Task(Of String)

            If String.IsNullOrWhiteSpace(taskPrompt) OrElse String.IsNullOrWhiteSpace(documentContent) Then
                Return ""
            End If

            Dim workingContent As String = documentContent.Trim()
            If workingContent.Length > MaxCharsPerRelevantExtractDocument Then
                workingContent = workingContent.Substring(0, MaxCharsPerRelevantExtractDocument)
            End If

            Dim promptSystem As String =
        "You extract only verbatim passages from a document. " &
        "Return only passages that are directly useful for answering the USER TASK. " &
        "The RETRIEVAL HINT is only supporting context and must not narrow or broaden the USER TASK incorrectly. " &
        "Rules: " &
        "1. Copy text verbatim from the DOCUMENT only. " &
        "2. Do not paraphrase, summarize, normalize, or explain. " &
        "3. Return the smallest useful set of excerpts. " &
        "4. Return XML only in this exact shape: <EXCERPTS><EXCERPT>...</EXCERPT></EXCERPTS>. " &
        "5. If nothing is relevant, return <EXCERPTS />."

            Dim promptUser As New StringBuilder()
            promptUser.AppendLine("USER TASK:")
            promptUser.AppendLine(taskPrompt.Trim())
            promptUser.AppendLine()

            If Not String.IsNullOrWhiteSpace(retrievalPrompt) Then
                promptUser.AppendLine("RETRIEVAL HINT:")
                promptUser.AppendLine(retrievalPrompt.Trim())
                promptUser.AppendLine()
            End If

            promptUser.AppendLine("DOCUMENT TITLE:")
            promptUser.AppendLine(If(match?.Title, ""))
            promptUser.AppendLine()
            promptUser.AppendLine("DOCUMENT TEXT START")
            promptUser.AppendLine(workingContent)
            promptUser.AppendLine("DOCUMENT TEXT END")

            Dim llmResponse As String = ""

            Try
                llmResponse = Await SharedMethods.LLM(
            context:=context,
            promptSystem:=promptSystem,
            promptUser:=promptUser.ToString(),
            Hidesplash:=True).ConfigureAwait(False)
            Catch
                Return ""
            End Try

            If String.IsNullOrWhiteSpace(llmResponse) Then
                Return ""
            End If

            Dim excerpts As New List(Of String)()

            For Each excerptMatch As Match In Regex.Matches(
        llmResponse,
        "<EXCERPT>(.*?)</EXCERPT>",
        RegexOptions.IgnoreCase Or RegexOptions.Singleline)

                Dim excerpt As String = excerptMatch.Groups(1).Value.Trim()
                If String.IsNullOrWhiteSpace(excerpt) Then Continue For
                If Not IsVerbatimExcerpt(workingContent, excerpt) Then Continue For
                If excerpts.Any(Function(x) String.Equals(x, excerpt, StringComparison.Ordinal)) Then Continue For

                excerpts.Add(excerpt)
            Next

            If excerpts.Count = 0 Then
                Return ""
            End If

            Return String.Join(vbCrLf & vbCrLf, excerpts)
        End Function

        Private Shared Function IsVerbatimExcerpt(documentContent As String, excerpt As String) As Boolean
            If String.IsNullOrWhiteSpace(documentContent) OrElse String.IsNullOrWhiteSpace(excerpt) Then
                Return False
            End If

            Dim normalizedDocument = documentContent.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
            Dim normalizedExcerpt = excerpt.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)

            Return normalizedDocument.Contains(normalizedExcerpt)
        End Function

        Private Shared Function BuildKnowledgeDocumentBlock(
        match As KnowledgeQueryService.KnowledgeMatch,
        bodyContent As String,
        readable As Boolean,
        excerpted As Boolean,
        note As String) As String

            Dim sourceUri = BuildKnowledgeFileUri(match.SourcePath)
            Dim wikiUri = BuildKnowledgeFileUri(match.WikiPagePath)
            Dim sourceLabel = If(String.IsNullOrWhiteSpace(match.SourcePath), "", Path.GetFileName(match.SourcePath))
            Dim wikiLabel = If(String.IsNullOrWhiteSpace(match.WikiPagePath), "", Path.GetFileName(match.WikiPagePath))
            Dim title = If(match.Title, Path.GetFileNameWithoutExtension(If(match.SourcePath, "Unknown")))

            Dim sb As New StringBuilder()
            sb.AppendLine(
        $"<KSDOCUMENT title=""{SecurityElement.Escape(If(title, ""))}"" " &
        $"store=""{SecurityElement.Escape(If(match.StoreName, ""))}"" " &
        $"sourcePath=""{SecurityElement.Escape(sourceUri)}"" " &
        $"sourceLabel=""{SecurityElement.Escape(sourceLabel)}"" " &
        $"wikiPath=""{SecurityElement.Escape(wikiUri)}"" " &
        $"wikiLabel=""{SecurityElement.Escape(wikiLabel)}"" " &
        $"matchReason=""{SecurityElement.Escape(If(match.MatchReason, ""))}"" " &
        $"readable=""{If(readable, "true", "false")}"" " &
        $"excerpted=""{If(excerpted, "true", "false")}"">")

            If Not String.IsNullOrWhiteSpace(note) Then
                sb.AppendLine("<KSNOTE>")
                sb.AppendLine(note)
                sb.AppendLine("</KSNOTE>")
            End If

            sb.AppendLine(bodyContent)
            sb.AppendLine("</KSDOCUMENT>")

            Return sb.ToString()
        End Function

        Private Shared Function BuildKnowledgeFileUri(filePath As String) As String
            If String.IsNullOrWhiteSpace(filePath) Then Return ""

            Try
                Dim candidate = filePath.Trim()

                If candidate.StartsWith("file://", StringComparison.OrdinalIgnoreCase) Then
                    Return candidate
                End If

                If Not Path.IsPathRooted(candidate) Then
                    Return candidate.Replace("\"c, "/"c)
                End If

                Dim fullPath = Path.GetFullPath(candidate)
                Dim uri As New Uri(fullPath)
                Return uri.AbsoluteUri
            Catch
                Try
                    Dim candidate = filePath.Trim()

                    If candidate.StartsWith("file://", StringComparison.OrdinalIgnoreCase) Then
                        Return candidate
                    End If

                    If Not Path.IsPathRooted(candidate) Then
                        Return candidate.Replace("\"c, "/"c)
                    End If

                    Dim fullPath = Path.GetFullPath(candidate)
                    Return "file:///" & fullPath.Replace("\"c, "/"c).Replace(" ", "%20")
                Catch
                    Return ""
                End Try
            End Try
        End Function

#End Region

#Region "Helpers"

        ''' <summary>
        ''' Resolves a potentially relative source file path against the store's root directory.
        ''' Returns the original path if it is already absolute or the store is unavailable.
        ''' </summary>
        Private Shared Function ResolveSourcePath(filePath As String, store As KnowledgeStoreCatalog.KnowledgeStoreDefinition) As String
            If String.IsNullOrWhiteSpace(filePath) Then Return ""
            Try
                If Path.IsPathRooted(filePath) AndAlso File.Exists(filePath) Then
                    Return filePath
                End If
                If store IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(store.ResolvedSourcePath) Then
                    Dim combined = Path.GetFullPath(Path.Combine(store.ResolvedSourcePath, filePath))
                    If File.Exists(combined) Then Return combined
                End If
                Return filePath
            Catch
                Return filePath
            End Try
        End Function

        Private Shared Function SanitizeFileName(name As String) As String
            Dim invalid = Path.GetInvalidFileNameChars()
            Dim cleaned = New String(name.Where(Function(c) Not invalid.Contains(c)).ToArray())
            If cleaned.Length > 100 Then cleaned = cleaned.Substring(0, 100)
            Return If(String.IsNullOrWhiteSpace(cleaned), "unnamed", cleaned.Trim())
        End Function

#End Region

    End Class

End Namespace
