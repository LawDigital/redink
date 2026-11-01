' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: KnowledgeStoreManager.vb
' Purpose:
'   Defines the core Knowledge Store entry model and manages persisted
'   document-level index data across configured stores.
'
' Responsibilities:
'   - Provide the `KnowledgeEntry` data model used by indexing and retrieval.
'   - Load Knowledge Store index files from disk.
'   - Merge central and local index data with de-duplication and provenance.
'   - Save writable local index entries back to disk.
'   - Expose basic helper checks for Knowledge Store availability and indexing.
'
' Notes:
'   - This file manages document-entry index persistence.
'   - Per-store manifest handling is delegated to `KnowledgeStoreManifest`.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary

    ''' <summary>
    ''' Manages Knowledge Store index persistence and provides the KnowledgeEntry data model.
    ''' </summary>
    Public Class KnowledgeStoreManager

#Region "Data Model"

        ''' <summary>
        ''' Represents a single document entry in the Knowledge Store index.
        ''' </summary>
        Public Class KnowledgeEntry
            ''' <summary>File path (may contain environment variables like %AppData%).</summary>
            Public Property FilePath As String

            ''' <summary>Human-readable title extracted or derived from the document.</summary>
            Public Property Title As String

            ''' <summary>Short summary of the document content (up to ~2000 chars).</summary>
            Public Property Summary As String

            ''' <summary>Keywords extracted via term-frequency analysis.</summary>
            Public Property Keywords As String()

            ''' <summary>Full or truncated text snapshot of the document (up to ~80K chars).</summary>
            <JsonIgnore>
            Public Property ContentSnapshot As String

            ''' <summary>Date and time the document was last indexed.</summary>
            Public Property IndexedDate As DateTime

            ''' <summary>User-assigned tags for filtered retrieval (e.g., "contracts", "NDA").</summary>
            Public Property Tags As String()

            ''' <summary>
            ''' Runtime flag indicating whether this entry originates from the central index.
            ''' Not serialized to disk.
            ''' </summary>
            <JsonIgnore>
            Public Property IsFromCentralIndex As Boolean
        End Class

#End Region

#Region "Configuration Checks"

        ''' <summary>
        ''' Returns True if at least one Knowledge Store path is configured.
        ''' </summary>
        Public Shared Function IsConfigured(context As ISharedContext) As Boolean
            Return Not String.IsNullOrWhiteSpace(context.INI_KnowledgeStorePath) OrElse
                   Not String.IsNullOrWhiteSpace(context.INI_KnowledgeStorePathLocal)
        End Function

        ''' <summary>
        ''' Returns True if the local Knowledge Store path is configured and writable.
        ''' </summary>
        Public Shared Function HasLocalIndex(context As ISharedContext) As Boolean
            Return Not String.IsNullOrWhiteSpace(context.INI_KnowledgeStorePathLocal)
        End Function

#End Region

#Region "Load — Merged Index"

        ''' <summary>
        ''' Loads and merges the central and local Knowledge Store indexes.
        ''' Local entries take precedence over central entries with the same FilePath.
        ''' </summary>
        ''' <param name="context">Shared context providing the two index paths.</param>
        ''' <returns>Merged list of KnowledgeEntry objects, or an empty list on error.</returns>
        Public Shared Function LoadMergedIndex(context As ISharedContext) As List(Of KnowledgeEntry)
            Dim merged As New List(Of KnowledgeEntry)
            Dim seenPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            ' Load local entries first (they take precedence)
            Dim localEntries = LoadIndexFile(context.INI_KnowledgeStorePathLocal, isFromCentral:=False)
            For Each entry In localEntries
                merged.Add(entry)
                If Not String.IsNullOrWhiteSpace(entry.FilePath) Then
                    seenPaths.Add(NormalizePath(entry.FilePath))
                End If
            Next

            ' Load central entries, skip duplicates
            Dim centralEntries = LoadIndexFile(context.INI_KnowledgeStorePath, isFromCentral:=True)
            For Each entry In centralEntries
                Dim normalized = NormalizePath(If(entry.FilePath, ""))
                If Not String.IsNullOrWhiteSpace(normalized) AndAlso seenPaths.Contains(normalized) Then
                    Continue For ' Local entry already exists for this path
                End If
                merged.Add(entry)
                If Not String.IsNullOrWhiteSpace(normalized) Then
                    seenPaths.Add(normalized)
                End If
            Next

            Return merged
        End Function

        ''' <summary>
        ''' Loads a single index file from disk and returns its entries.
        ''' </summary>
        ''' <param name="rawPath">Path (may contain environment variables) to the JSON index file.</param>
        ''' <param name="isFromCentral">If True, marks all loaded entries as central.</param>
        ''' <returns>List of entries, or empty list if file is missing or invalid.</returns>
        Private Shared Function LoadIndexFile(rawPath As String, isFromCentral As Boolean) As List(Of KnowledgeEntry)
            Dim result As New List(Of KnowledgeEntry)
            If String.IsNullOrWhiteSpace(rawPath) Then Return result

            Try
                Dim expandedPath = SharedMethods.ExpandEnvironmentVariables(rawPath.Trim())
                If String.IsNullOrWhiteSpace(expandedPath) OrElse Not File.Exists(expandedPath) Then
                    Return result
                End If

                Dim json = File.ReadAllText(expandedPath, System.Text.Encoding.UTF8)
                If String.IsNullOrWhiteSpace(json) Then Return result

                Dim arr As JArray = JArray.Parse(json)
                For Each token As JToken In arr
                    Try
                        Dim entry = ParseEntry(token, isFromCentral)
                        If entry IsNot Nothing Then
                            result.Add(entry)
                        End If
                    Catch ex As Exception
                        Debug.WriteLine($"KnowledgeStoreManager: Skipping malformed entry: {ex.Message}")
                    End Try
                Next
            Catch ex As Exception
                Debug.WriteLine($"KnowledgeStoreManager: Error loading index '{rawPath}': {ex.Message}")
            End Try

            Return result
        End Function

        ''' <summary>
        ''' Parses a single JSON token into a KnowledgeEntry.
        ''' </summary>
        Private Shared Function ParseEntry(token As JToken, isFromCentral As Boolean) As KnowledgeEntry
            If token.Type <> JTokenType.Object Then Return Nothing
            Dim obj = CType(token, JObject)

            Dim entry As New KnowledgeEntry() With {
                .FilePath = obj.Value(Of String)("FilePath"),
                .Title = If(obj.Value(Of String)("Title"), ""),
                .Summary = If(obj.Value(Of String)("Summary"), ""),
                .ContentSnapshot = If(obj.Value(Of String)("ContentSnapshot"), ""),
                .IsFromCentralIndex = isFromCentral
            }

            ' Parse IndexedDate
            Dim dateStr = obj.Value(Of String)("IndexedDate")
            If Not String.IsNullOrWhiteSpace(dateStr) Then
                Dim parsed As DateTime
                If DateTime.TryParse(dateStr, parsed) Then
                    entry.IndexedDate = parsed
                End If
            End If

            ' Parse Keywords array
            Dim kwToken = obj("Keywords")
            If kwToken IsNot Nothing AndAlso kwToken.Type = JTokenType.Array Then
                entry.Keywords = kwToken.ToObject(Of String())()
            End If

            ' Parse Tags array
            Dim tagToken = obj("Tags")
            If tagToken IsNot Nothing AndAlso tagToken.Type = JTokenType.Array Then
                entry.Tags = tagToken.ToObject(Of String())()
            End If

            Return entry
        End Function

#End Region

#Region "Save — Local Index"

        ''' <summary>
        ''' Saves local Knowledge Store entries to the local index file.
        ''' Entries flagged as central are excluded automatically.
        ''' </summary>
        ''' <param name="entries">All entries (mixed central and local). Only local entries are written.</param>
        ''' <param name="context">Shared context providing KnowledgeStorePathLocal.</param>
        Public Shared Sub SaveLocalIndex(entries As List(Of KnowledgeEntry), context As ISharedContext)
            If String.IsNullOrWhiteSpace(context.INI_KnowledgeStorePathLocal) Then
                Throw New InvalidOperationException("KnowledgeStorePathLocal is not configured.")
            End If

            Dim expandedPath = SharedMethods.ExpandEnvironmentVariables(context.INI_KnowledgeStorePathLocal.Trim())
            If String.IsNullOrWhiteSpace(expandedPath) Then
                Throw New InvalidOperationException("KnowledgeStorePathLocal could not be expanded.")
            End If

            ' Filter to local entries only
            Dim localEntries = entries.Where(Function(e) Not e.IsFromCentralIndex).ToList()

            ' Build JSON array
            Dim arr As New JArray()
            For Each entry In localEntries
                Dim obj As New JObject()
                obj("FilePath") = If(entry.FilePath, "")
                obj("Title") = If(entry.Title, "")
                obj("Summary") = If(entry.Summary, "")
                obj("IndexedDate") = entry.IndexedDate.ToString("o") ' ISO 8601

                If entry.Keywords IsNot Nothing AndAlso entry.Keywords.Length > 0 Then
                    obj("Keywords") = New JArray(entry.Keywords)
                End If

                If entry.Tags IsNot Nothing AndAlso entry.Tags.Length > 0 Then
                    obj("Tags") = New JArray(entry.Tags)
                End If

                ' ContentSnapshot is intentionally excluded from the saved index
                ' to keep file sizes manageable. It is re-read on demand via KnowledgeIndexer.

                arr.Add(obj)
            Next

            ' Ensure directory exists
            Dim dir = Path.GetDirectoryName(expandedPath)
            If Not String.IsNullOrWhiteSpace(dir) AndAlso Not Directory.Exists(dir) Then
                Directory.CreateDirectory(dir)
            End If

            ' Write with formatting for human readability
            Dim jsonText = arr.ToString(Formatting.Indented)
            File.WriteAllText(expandedPath, jsonText, System.Text.Encoding.UTF8)
        End Sub

#End Region

#Region "Query Helpers"

        ''' <summary>
        ''' Finds entries matching a tag filter (case-insensitive).
        ''' </summary>
        ''' <param name="entries">The merged index to search.</param>
        ''' <param name="tag">Tag to filter by.</param>
        ''' <returns>Entries whose Tags array contains the specified tag.</returns>
        Public Shared Function FindByTag(entries As List(Of KnowledgeEntry), tag As String) As List(Of KnowledgeEntry)
            If String.IsNullOrWhiteSpace(tag) Then Return entries
            Return entries.Where(Function(e)
                                     Return e.Tags IsNot Nothing AndAlso
                                            e.Tags.Any(Function(t) String.Equals(t, tag, StringComparison.OrdinalIgnoreCase))
                                 End Function).ToList()
        End Function

        ''' <summary>
        ''' Finds entries whose Title, Summary, or Keywords match a search query (case-insensitive substring).
        ''' </summary>
        ''' <param name="entries">The merged index to search.</param>
        ''' <param name="query">Search text.</param>
        ''' <returns>Entries matching the query, ordered by relevance (title match first).</returns>
        Public Shared Function Search(entries As List(Of KnowledgeEntry), query As String) As List(Of KnowledgeEntry)
            If String.IsNullOrWhiteSpace(query) Then Return entries

            Dim q = query.Trim()
            Dim results As New List(Of Tuple(Of KnowledgeEntry, Integer)) ' entry, score

            For Each entry In entries
                Dim score As Integer = 0

                ' Title match (highest weight)
                If Not String.IsNullOrWhiteSpace(entry.Title) AndAlso
                   entry.Title.IndexOf(q, StringComparison.OrdinalIgnoreCase) >= 0 Then
                    score += 10
                End If

                ' Keyword match
                If entry.Keywords IsNot Nothing Then
                    For Each kw In entry.Keywords
                        If kw.IndexOf(q, StringComparison.OrdinalIgnoreCase) >= 0 Then
                            score += 5
                            Exit For
                        End If
                    Next
                End If

                ' Tag match
                If entry.Tags IsNot Nothing Then
                    For Each t In entry.Tags
                        If t.IndexOf(q, StringComparison.OrdinalIgnoreCase) >= 0 Then
                            score += 5
                            Exit For
                        End If
                    Next
                End If

                ' Summary match (lowest weight)
                If Not String.IsNullOrWhiteSpace(entry.Summary) AndAlso
                   entry.Summary.IndexOf(q, StringComparison.OrdinalIgnoreCase) >= 0 Then
                    score += 2
                End If

                If score > 0 Then
                    results.Add(Tuple.Create(entry, score))
                End If
            Next

            Return results.OrderByDescending(Function(r) r.Item2).
                           Select(Function(r) r.Item1).
                           ToList()
        End Function

        ''' <summary>
        ''' Returns the full content of a Knowledge Store entry by reading the file from disk.
        ''' Falls back to the stored ContentSnapshot if the file is not accessible.
        ''' </summary>
        ''' <param name="entry">The entry to load content for.</param>
        ''' <returns>The document text, or an empty string if unavailable.</returns>
        Public Shared Function GetContent(entry As KnowledgeEntry) As String
            If entry Is Nothing Then Return ""

            ' Try reading fresh from disk
            Try
                Dim expandedPath = SharedMethods.ExpandEnvironmentVariables(If(entry.FilePath, ""))
                If Not String.IsNullOrWhiteSpace(expandedPath) AndAlso File.Exists(expandedPath) Then
                    Dim ext = Path.GetExtension(expandedPath).ToLowerInvariant()
                    Select Case ext
                        Case ".txt", ".ini", ".csv", ".log", ".json", ".xml", ".html", ".htm",
                             ".md", ".yaml", ".yml", ".vb", ".cs", ".js", ".ts", ".py",
                             ".java", ".cpp", ".c", ".h", ".sql"
                            Return File.ReadAllText(expandedPath, System.Text.Encoding.UTF8)
                        Case ".docx"
                            Return SharedMethods.ReadDocxSandboxed(expandedPath)
                        Case ".pdf"
                            Dim t = SharedMethods.ReadPdfAsText(expandedPath, ReturnErrorInsteadOfEmpty:=False, DoOCR:=False, AskUser:=False)
                            t.Wait()
                            Return If(t.Result, "")
                        Case ".rtf"
                            Try
                                Return SharedMethods.LoadRtfFileAsPlainText(expandedPath)
                            Catch
                                Return File.ReadAllText(expandedPath, System.Text.Encoding.UTF8)
                            End Try
                        Case ".xlsx"
                            Return SharedMethods.ReadXlsxSandboxed(expandedPath)
                        Case ".pptx"
                            Return SharedMethods.ReadPptxSandboxed(expandedPath)
                        Case Else
                            Return File.ReadAllText(expandedPath, System.Text.Encoding.UTF8)
                    End Select
                End If
            Catch ex As Exception
                Debug.WriteLine($"KnowledgeStoreManager.GetContent: Error reading '{entry.FilePath}': {ex.Message}")
            End Try

            ' Fall back to cached snapshot
            Return If(entry.ContentSnapshot, "")
        End Function

#End Region

#Region "Path Utilities"

        ''' <summary>
        ''' Normalizes a file path for deduplication comparison.
        ''' Expands environment variables and converts to uppercase.
        ''' </summary>
        Private Shared Function NormalizePath(rawPath As String) As String
            If String.IsNullOrWhiteSpace(rawPath) Then Return ""
            Try
                Dim expanded = SharedMethods.ExpandEnvironmentVariables(rawPath.Trim())
                Return expanded.ToUpperInvariant()
            Catch
                Return rawPath.Trim().ToUpperInvariant()
            End Try
        End Function

#End Region

    End Class

End Namespace