' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: KnowledgeQueryService.vb
' Purpose:
'   Resolves Knowledge Store queries across active stores and builds source-
'   backed context blocks for downstream prompt and LLM use.
'
' Responsibilities:
'   - Parse Knowledge Store queries including optional `tag:` and `store:`
'     filters.
'   - Enumerate active stores and respect store-level schema constraints.
'   - Search wiki content, manifests, and source-backed content as appropriate.
'   - Merge, rank, and de-duplicate matches across stores.
'   - Build prompt-ready context blocks with clear document and store provenance.
'
' Notes:
'   - This is the query-time retrieval layer of the Knowledge Store subsystem.
'   - It is used by Knowledge Store-enabled prompt flows and related retrieval
'     helpers.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary

    Public Class KnowledgeQueryService

        Public Class KnowledgeMatch
            Public Property Entry As KnowledgeStoreManager.KnowledgeEntry
            Public Property Score As Double = 0.0
            Public Property StoreName As String = ""
            Public Property WikiPagePath As String = ""
            Public Property Title As String = ""
            Public Property Summary As String = ""
            Public Property SourcePath As String = ""
            Public Property MatchReason As String = ""
        End Class

        Private Class WikiPageCandidate
            Public Property FilePath As String = ""
            Public Property RelativePath As String = ""
            Public Property Title As String = ""
            Public Property Summary As String = ""
            Public Property SourcePath As String = ""
            Public Property Content As String = ""
            Public Property Kind As String = ""
        End Class

        Public Shared Function ResolveQuery(query As String,
                                            context As ISharedContext,
                                            Optional maxResults As Integer = 6) As List(Of KnowledgeMatch)
            Return ResolveQueryAsync(query, context, maxResults).GetAwaiter().GetResult()
        End Function

        Public Shared Async Function ResolveQueryAsync(query As String,
                                                       context As ISharedContext,
                                                       Optional maxResults As Integer = 6) As Task(Of List(Of KnowledgeMatch))
            Dim resultsByKey As New Dictionary(Of String, KnowledgeMatch)(StringComparer.OrdinalIgnoreCase)

            If String.IsNullOrWhiteSpace(query) OrElse context Is Nothing Then
                Return New List(Of KnowledgeMatch)()
            End If

            If Not KnowledgeStoreCatalog.IsConfigured(context) Then
                Return New List(Of KnowledgeMatch)()
            End If

            Dim tagFilter As String = ""
            Dim storeFilter As String = ""
            Dim freeTerms As New List(Of String)()

            For Each token In KnowledgeTriggerHelper.TokenizeWithQuotes(query)
                If token.StartsWith("tag:", StringComparison.OrdinalIgnoreCase) Then
                    tagFilter = token.Substring(4).Trim()
                ElseIf token.StartsWith("store:", StringComparison.OrdinalIgnoreCase) Then
                    storeFilter = token.Substring(6).Trim()
                Else
                    freeTerms.Add(token)
                End If
            Next

            Dim searchTerms = String.Join(" ", freeTerms).Trim()
            Dim keywords = Tokenize(searchTerms)

            For Each store In KnowledgeStoreCatalog.GetActiveStores(context)
                If Not String.IsNullOrWhiteSpace(storeFilter) AndAlso
                   Not store.Name.Equals(storeFilter, StringComparison.OrdinalIgnoreCase) Then
                    Continue For
                End If

                Dim schema = KnowledgeStoreSchema.LoadOrCreate(store.ResolvedSourcePath)
                If IsIgnoredBySchema(searchTerms, schema) Then
                    Continue For
                End If

                Dim wikiDir = KnowledgeStoreCatalog.GetWikiPath(store, createIfMissing:=False)
                Dim pageCandidates = LoadWikiPageCandidates(wikiDir)
                Dim pagesByPath = pageCandidates.ToDictionary(Function(p) p.FilePath, StringComparer.OrdinalIgnoreCase)

                ' 1. Wiki-first lexical scoring
                For Each page In pageCandidates
                    Dim score = ScoreWikiPage(page, keywords, searchTerms, schema)
                    If score <= 0 Then Continue For

                    Dim key = BuildMatchKey(store.Name, page.FilePath, page.SourcePath, page.Title)
                    AddOrUpdateMatch(resultsByKey, key,
                                     New KnowledgeMatch With {
                                         .Score = score,
                                         .StoreName = store.Name,
                                         .WikiPagePath = page.FilePath,
                                         .Title = page.Title,
                                         .Summary = page.Summary,
                                         .SourcePath = ResolveEntrySourcePath(page.SourcePath, store.ResolvedSourcePath),
                                         .MatchReason = "wiki-keyword"
                                     })
                Next

                ' 2. Semantic re-ranking over wiki pages / indexed chunks
                Try
                    Dim semanticResults = Await KnowledgeEmbeddingService.SearchAsync(
                        store.ResolvedSourcePath,
                        query,
                        Math.Max(maxResults * 4, 12),
                        context).ConfigureAwait(False)

                    Dim grouped = semanticResults.
                        GroupBy(Function(r) r.FilePath, StringComparer.OrdinalIgnoreCase).
                        Select(Function(g) New With {
                            .FilePath = g.Key,
                            .Score = g.Max(Function(x) x.Score)
                        })

                    For Each item In grouped
                        Dim page As WikiPageCandidate = Nothing
                        If pagesByPath.TryGetValue(item.FilePath, page) Then
                            Dim semanticScore = item.Score * 10.0
                            If semanticScore > 0 Then
                                Dim key = BuildMatchKey(store.Name, page.FilePath, page.SourcePath, page.Title)
                                AddOrUpdateMatch(resultsByKey, key,
                                                 New KnowledgeMatch With {
                                                     .Score = semanticScore,
                                                     .StoreName = store.Name,
                                                     .WikiPagePath = page.FilePath,
                                                     .Title = page.Title,
                                                     .Summary = page.Summary,
                                                     .SourcePath = ResolveEntrySourcePath(page.SourcePath, store.ResolvedSourcePath),
                                                     .MatchReason = "wiki-semantic"
                                                 })
                            End If
                        End If
                    Next
                Catch
                    ' Semantic search is an optional enhancement. Query resolution must still work without it.
                End Try

                ' 3. Manifest/raw fallback
                Dim manifest = KnowledgeStoreManifest.Load(store)
                For Each entry In manifest.Entries
                    If String.IsNullOrWhiteSpace(entry.FilePath) Then Continue For

                    If Not String.IsNullOrWhiteSpace(tagFilter) Then
                        If entry.Tags Is Nothing OrElse
                           Not entry.Tags.Any(Function(t) t.Equals(tagFilter, StringComparison.OrdinalIgnoreCase)) Then
                            Continue For
                        End If
                    End If

                    Dim score As Double = 0.0
                    If keywords.Count = 0 AndAlso Not String.IsNullOrWhiteSpace(tagFilter) Then
                        score = 1.0
                    Else
                        Dim titleLower = If(entry.Title, "").ToLowerInvariant()
                        Dim summaryLower = If(entry.Summary, "").ToLowerInvariant()

                        For Each kw In keywords
                            If titleLower.Contains(kw) Then score += 3.0
                            If summaryLower.Contains(kw) Then score += 1.0

                            If entry.Keywords IsNot Nothing AndAlso
                               entry.Keywords.Any(Function(k) k.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0) Then
                                score += 1.5
                            End If

                            If entry.Tags IsNot Nothing AndAlso
                               entry.Tags.Any(Function(t) t.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0) Then
                                score += 2.0
                            End If
                        Next
                    End If

                    If score > 0 Then
                        Dim key = BuildMatchKey(store.Name, "", entry.FilePath, If(entry.Title, entry.FilePath))
                        AddOrUpdateMatch(resultsByKey, key,
                                         New KnowledgeMatch With {
                                             .Entry = entry,
                                             .Score = score,
                                             .StoreName = store.Name,
                                             .Title = If(entry.Title, Path.GetFileNameWithoutExtension(entry.FilePath)),
                                             .Summary = If(entry.Summary, ""),
                                             .SourcePath = ResolveEntrySourcePath(entry.FilePath, store.ResolvedSourcePath),
                                             .MatchReason = "manifest-fallback"
                                         })
                    End If
                Next
            Next

            Return resultsByKey.Values.
                OrderByDescending(Function(r) r.Score).
                ThenBy(Function(r) r.Title, StringComparer.OrdinalIgnoreCase).
                Take(maxResults).
                ToList()
        End Function

        ''' <summary>
        ''' Resolves a manifest entry file path to an absolute path using the store root.
        ''' </summary>
        Private Shared Function ResolveEntrySourcePath(filePath As String, storeRoot As String) As String
            If String.IsNullOrWhiteSpace(filePath) Then Return ""
            Try
                If Path.IsPathRooted(filePath) AndAlso File.Exists(filePath) Then Return filePath
                If Not String.IsNullOrWhiteSpace(storeRoot) Then
                    Dim combined = Path.GetFullPath(Path.Combine(storeRoot, filePath))
                    If File.Exists(combined) Then Return combined
                End If
                Return filePath
            Catch
                Return filePath
            End Try
        End Function

        Public Shared Function BuildKnowledgeContext(matches As List(Of KnowledgeMatch),
                                                     Optional maxTotalChars As Integer = 80000) As String
            If matches Is Nothing OrElse matches.Count = 0 Then Return ""

            Dim sb As New StringBuilder()
            sb.AppendLine("<KNOWLEDGESTORE>")
            sb.AppendLine("The following wiki pages and/or source documents have been provided for context.")
            sb.AppendLine("IMPORTANT: When citing information, ALWAYS prefer the original source file link over the wiki page link.")
            sb.AppendLine("If a sourcePath is provided and non-empty, ALWAYS cite it as: [Source Title](sourcePath)")
            sb.AppendLine("Only fall back to wikiPath if no sourcePath is available for that document.")
            sb.AppendLine("Do NOT invent links. Use only the explicit sourcePath or wikiPath values provided in the KSDOCUMENT attributes.")
            sb.AppendLine("Do not cite only the document name when an explicit clickable path is available.")
            sb.AppendLine()

            Dim totalChars As Integer = 0

            For Each m In matches
                Dim content As String = ""

                If Not String.IsNullOrWhiteSpace(m.WikiPagePath) AndAlso File.Exists(m.WikiPagePath) Then
                    content = File.ReadAllText(m.WikiPagePath, Encoding.UTF8)
                ElseIf m.Entry IsNot Nothing Then
                    content = KnowledgeStoreManager.GetContent(m.Entry)
                End If

                If String.IsNullOrWhiteSpace(content) Then Continue For

                Dim remaining = maxTotalChars - totalChars
                If remaining <= 0 Then Exit For
                If content.Length > remaining Then
                    content = content.Substring(0, remaining) & vbCrLf & "[... truncated ...]"
                End If

                Dim sourceUri = BuildFileUri(m.SourcePath)
                Dim wikiUri = BuildFileUri(m.WikiPagePath)
                Dim sourceLabel = If(String.IsNullOrWhiteSpace(m.SourcePath), "", Path.GetFileName(m.SourcePath))
                Dim wikiLabel = If(String.IsNullOrWhiteSpace(m.WikiPagePath), "", Path.GetFileName(m.WikiPagePath))

                sb.AppendLine(
                    $"<KSDOCUMENT title=""{System.Security.SecurityElement.Escape(If(m.Title, ""))}"" " &
                    $"store=""{System.Security.SecurityElement.Escape(If(m.StoreName, ""))}"" " &
                    $"sourcePath=""{System.Security.SecurityElement.Escape(sourceUri)}"" " &
                    $"sourceLabel=""{System.Security.SecurityElement.Escape(sourceLabel)}"" " &
                    $"wikiPath=""{System.Security.SecurityElement.Escape(wikiUri)}"" " &
                    $"wikiLabel=""{System.Security.SecurityElement.Escape(wikiLabel)}"" " &
                    $"matchReason=""{System.Security.SecurityElement.Escape(If(m.MatchReason, ""))}"">")
                sb.AppendLine(content)
                sb.AppendLine("</KSDOCUMENT>")
                sb.AppendLine()

                totalChars += content.Length
            Next

            sb.AppendLine("</KNOWLEDGESTORE>")
            Return sb.ToString()
        End Function

        Public Shared Function ResolveAndBuild(query As String,
                                               context As ISharedContext,
                                               Optional maxResults As Integer = 6,
                                               Optional maxTotalChars As Integer = 80000) As String
            Return ResolveAndBuildAsync(query, context, maxResults, maxTotalChars).GetAwaiter().GetResult()
        End Function

        Public Shared Async Function ResolveAndBuildAsync(query As String,
                                                          context As ISharedContext,
                                                          Optional maxResults As Integer = 6,
                                                          Optional maxTotalChars As Integer = 80000) As Task(Of String)
            Dim matches = Await ResolveQueryAsync(query, context, maxResults).ConfigureAwait(False)

            For Each storeName In matches.Select(Function(m) m.StoreName).Distinct(StringComparer.OrdinalIgnoreCase)
                Dim store = KnowledgeStoreCatalog.GetStoreByName(storeName, context)
                If store IsNot Nothing Then
                    KnowledgeWikiService.AppendOperationalLog(store.ResolvedSourcePath, "query", query)
                End If
            Next

            Return BuildKnowledgeContext(matches, maxTotalChars)
        End Function

        Public Shared Async Function FileQueryResultAsync(query As String,
                                                          answerMarkdown As String,
                                                          context As ISharedContext,
                                                          Optional preferredStoreName As String = "",
                                                          Optional preferredTitle As String = "") As Task(Of Integer)
            If String.IsNullOrWhiteSpace(query) OrElse
               String.IsNullOrWhiteSpace(answerMarkdown) OrElse
               context Is Nothing Then
                Return 0
            End If

            Dim matches = Await ResolveQueryAsync(query, context, 12).ConfigureAwait(False)
            Dim storesToUse As New List(Of KnowledgeStoreCatalog.KnowledgeStoreDefinition)()

            If Not String.IsNullOrWhiteSpace(preferredStoreName) Then
                Dim preferredStore = KnowledgeStoreCatalog.GetStoreByName(preferredStoreName, context)
                If preferredStore IsNot Nothing Then
                    storesToUse.Add(preferredStore)
                End If
            Else
                For Each storeName In matches.
                    Select(Function(m) m.StoreName).
                    Where(Function(s) Not String.IsNullOrWhiteSpace(s)).
                    Distinct(StringComparer.OrdinalIgnoreCase)

                    Dim store = KnowledgeStoreCatalog.GetStoreByName(storeName, context)
                    If store IsNot Nothing Then
                        storesToUse.Add(store)
                    End If
                Next
            End If

            Dim filedCount As Integer = 0

            For Each store In storesToUse.
                GroupBy(Function(s) s.Name, StringComparer.OrdinalIgnoreCase).
                Select(Function(g) g.First())

                Dim schema = KnowledgeStoreSchema.LoadOrCreate(store.ResolvedSourcePath)
                If schema Is Nothing OrElse Not schema.QueryFilingEnabled Then
                    Continue For
                End If

                Dim storeSourcePaths = matches.
                    Where(Function(m) m.StoreName.Equals(store.Name, StringComparison.OrdinalIgnoreCase)).
                    Select(Function(m) m.SourcePath).
                    Where(Function(p) Not String.IsNullOrWhiteSpace(p)).
                    Distinct(StringComparer.OrdinalIgnoreCase).
                    ToList()

                Dim saved = Await KnowledgeWikiService.FileQueryResultAsync(
                    kbRootPath:=store.ResolvedSourcePath,
                    queryText:=query,
                    answerMarkdown:=answerMarkdown,
                    context:=context,
                    preferredTitle:=preferredTitle,
                    sourcePaths:=storeSourcePaths).ConfigureAwait(False)

                If saved Then
                    filedCount += 1
                End If
            Next

            Return filedCount
        End Function

        Private Shared Function LoadWikiPageCandidates(wikiDir As String) As List(Of WikiPageCandidate)
            Dim result As New List(Of WikiPageCandidate)()
            If String.IsNullOrWhiteSpace(wikiDir) OrElse Not Directory.Exists(wikiDir) Then Return result

            For Each file In Directory.GetFiles(wikiDir, "*.md", SearchOption.AllDirectories)
                Dim name = Path.GetFileName(file)
                If name.Equals(KnowledgeStoreCatalog.IndexFile, StringComparison.OrdinalIgnoreCase) Then Continue For
                If name.Equals(KnowledgeStoreCatalog.LogFile, StringComparison.OrdinalIgnoreCase) Then Continue For
                If name.Equals("health_report.md", StringComparison.OrdinalIgnoreCase) Then Continue For

                Try
                    Dim content = System.IO.File.ReadAllText(file, Encoding.UTF8)
                    Dim frontMatter = GetFrontMatter(content)

                    result.Add(New WikiPageCandidate With {
                        .FilePath = file,
                        .RelativePath = NormalizeRelativePath(GetRelativePathCompat(wikiDir, file)),
                        .Title = If(GetFrontMatterScalar(frontMatter, "title"), Path.GetFileNameWithoutExtension(file)),
                        .Summary = If(GetFrontMatterScalar(frontMatter, "summary"), ""),
                        .SourcePath = GetFirstSourcePath(frontMatter),
                        .Kind = If(GetFrontMatterScalar(frontMatter, "kind"), ""),
                        .Content = content
                    })
                Catch
                End Try
            Next

            Return result
        End Function

        Private Shared Function ScoreWikiPage(page As WikiPageCandidate,
                                              keywords As HashSet(Of String),
                                              rawQuery As String,
                                              schema As KnowledgeStoreSchema) As Double
            Dim score As Double = 0.0
            If page Is Nothing Then Return 0

            If IsIgnoredBySchema(page.Title & " " & page.Summary, schema) Then
                score -= 2.0
            End If

            If keywords.Count = 0 Then
                If Not String.IsNullOrWhiteSpace(rawQuery) Then score = 0.5
                Return score
            End If

            Dim titleLower = If(page.Title, "").ToLowerInvariant()
            Dim summaryLower = If(page.Summary, "").ToLowerInvariant()
            Dim contentLower = If(page.Content, "").ToLowerInvariant()
            Dim kindLower = If(page.Kind, "").ToLowerInvariant()

            For Each kw In keywords
                If titleLower.Contains(kw) Then score += 4.0
                If summaryLower.Contains(kw) Then score += 2.0
                If contentLower.Contains(kw) Then score += 0.75
            Next

            If kindLower = "analysis" Then score += 0.2
            If page.Content.IndexOf("## Related Pages", StringComparison.OrdinalIgnoreCase) >= 0 Then score += 0.25
            If page.Content.IndexOf("## Sources", StringComparison.OrdinalIgnoreCase) >= 0 Then score += 0.25

            Return score
        End Function

        Private Shared Function IsIgnoredBySchema(text As String, schema As KnowledgeStoreSchema) As Boolean
            If schema Is Nothing OrElse schema.IgnoredTopics Is Nothing OrElse schema.IgnoredTopics.Length = 0 Then
                Return False
            End If

            Dim haystack = If(text, "")
            For Each ignored In schema.IgnoredTopics
                If Not String.IsNullOrWhiteSpace(ignored) AndAlso
                   haystack.IndexOf(ignored, StringComparison.OrdinalIgnoreCase) >= 0 Then
                    Return True
                End If
            Next

            Return False
        End Function

        Private Shared Sub AddOrUpdateMatch(resultsByKey As Dictionary(Of String, KnowledgeMatch),
                                            key As String,
                                            incoming As KnowledgeMatch)
            If String.IsNullOrWhiteSpace(key) OrElse incoming Is Nothing Then Return

            Dim existing As KnowledgeMatch = Nothing
            If resultsByKey.TryGetValue(key, existing) Then
                existing.Score += incoming.Score

                If String.IsNullOrWhiteSpace(existing.Title) Then existing.Title = incoming.Title
                If String.IsNullOrWhiteSpace(existing.Summary) Then existing.Summary = incoming.Summary
                If String.IsNullOrWhiteSpace(existing.SourcePath) Then existing.SourcePath = incoming.SourcePath
                If String.IsNullOrWhiteSpace(existing.WikiPagePath) Then existing.WikiPagePath = incoming.WikiPagePath
                If existing.Entry Is Nothing Then existing.Entry = incoming.Entry

                If String.IsNullOrWhiteSpace(existing.MatchReason) Then
                    existing.MatchReason = incoming.MatchReason
                ElseIf Not String.IsNullOrWhiteSpace(incoming.MatchReason) AndAlso
                       existing.MatchReason.IndexOf(incoming.MatchReason, StringComparison.OrdinalIgnoreCase) < 0 Then
                    existing.MatchReason &= "+" & incoming.MatchReason
                End If
            Else
                resultsByKey(key) = incoming
            End If
        End Sub

        Private Shared Function BuildMatchKey(storeName As String,
                                              wikiPagePath As String,
                                              sourcePath As String,
                                              title As String) As String
            If Not String.IsNullOrWhiteSpace(wikiPagePath) Then
                Return storeName & "|W|" & wikiPagePath
            End If

            If Not String.IsNullOrWhiteSpace(sourcePath) Then
                Return storeName & "|S|" & sourcePath
            End If

            Return storeName & "|T|" & title
        End Function

        Private Shared Function Tokenize(text As String) As HashSet(Of String)
            Return Regex.Split(If(text, "").ToLowerInvariant(), "[^a-z0-9äöüß]+").
                Where(Function(t) t.Length >= 2).
                ToHashSet(StringComparer.OrdinalIgnoreCase)
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

        Private Shared Function GetFirstSourcePath(frontMatter As String) As String
            If String.IsNullOrWhiteSpace(frontMatter) Then Return ""

            Dim lines = frontMatter.Split({vbCrLf, vbLf, vbCr}, StringSplitOptions.None)
            Dim inTarget As Boolean = False

            For Each rawLine In lines
                Dim line = rawLine.TrimEnd()

                If line.Trim().StartsWith("source_paths:", StringComparison.OrdinalIgnoreCase) Then
                    inTarget = True
                    Continue For
                End If

                If inTarget Then
                    Dim trimmed = line.Trim()
                    If trimmed.StartsWith("- ", StringComparison.Ordinal) Then
                        Dim value = trimmed.Substring(2).Trim()
                        ' Strip surrounding double quotes
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
                        Return value
                    ElseIf trimmed.Contains(":"c) Then
                        Exit For
                    End If
                End If
            Next

            Return ""
        End Function

        Private Shared Function BuildFileUri(filePath As String) As String
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

        Private Shared Function NormalizeRelativePath(pathValue As String) As String
            Return If(pathValue, "").Replace("\"c, "/"c).Trim()
        End Function

        Private Shared Function GetRelativePathCompat(basePath As String, targetPath As String) As String
            Try
                If String.IsNullOrWhiteSpace(basePath) Then Return targetPath
                If String.IsNullOrWhiteSpace(targetPath) Then Return ""

                Dim normalizedBase = Path.GetFullPath(basePath)
                Dim normalizedTarget = Path.GetFullPath(targetPath)

                If Not normalizedBase.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal) Then
                    normalizedBase &= Path.DirectorySeparatorChar
                End If

                Dim baseUri As New Uri(normalizedBase, UriKind.Absolute)
                Dim targetUri As New Uri(normalizedTarget, UriKind.Absolute)

                Dim relativeUri = baseUri.MakeRelativeUri(targetUri)
                Dim relativePath = Uri.UnescapeDataString(relativeUri.ToString())

                Return relativePath.Replace("/"c, Path.DirectorySeparatorChar)
            Catch
                Return targetPath
            End Try
        End Function


    End Class

End Namespace
