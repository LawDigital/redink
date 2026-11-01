' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: KnowledgeIndexer.vb
' Purpose:
'   Indexes source documents for a Knowledge Store by extracting readable text,
'   deriving metadata, and optionally triggering wiki and embedding updates.
'
' Responsibilities:
'   - Read supported source file types through shared text-extraction helpers.
'   - Normalize extracted document text and derive summary metadata.
'   - Optionally use LLM-assisted indexing for richer metadata generation.
'   - Produce populated `KnowledgeStoreManager.KnowledgeEntry` objects.
'   - Coordinate optional downstream wiki generation and embedding refresh.
'
' Notes:
'   - This service is used by both background and foreground indexing flows.
'   - Supported source types are defined by `SupportedExtensions`.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary

    ''' <summary>
    ''' Indexes documents for Knowledge Store inclusion by extracting text, titles, and keywords.
    ''' </summary>
    Public Class KnowledgeIndexer

        ''' <summary>
        ''' File extensions supported for Knowledge Store indexing.
        ''' </summary>
        Public Shared ReadOnly SupportedExtensions As String() = {
            ".txt", ".rtf", ".docx", ".pdf", ".pptx", ".xlsx",
            ".json", ".xml", ".html", ".htm", ".md", ".csv",
            ".yaml", ".yml", ".ini", ".log",
            ".vb", ".cs", ".js", ".ts", ".py", ".java",
            ".cpp", ".c", ".h", ".sql"
        }

        ''' <summary>
        ''' Maximum characters to extract from a document for the content summary.
        ''' </summary>
        Private Const MaxSummaryChars As Integer = 2000

        ''' <summary>
        ''' Maximum characters to store as the full content snapshot.
        ''' </summary>
        Private Const MaxContentChars As Integer = 80000

        ''' <summary>
        ''' Maximum characters sent to the LLM for summary/keyword generation.
        ''' </summary>
        Private Const MaxLLMInputChars As Integer = 12000

        ''' <summary>
        ''' System prompt for LLM-based summary and keyword extraction.
        ''' </summary>
        Private Const LLMIndexPrompt As String =
            "You are an indexing assistant. Given the following document excerpt, produce:" & vbCrLf &
            "1. A concise title (max 100 chars) on the first line, prefixed with ""TITLE: ""." & vbCrLf &
            "2. A summary of the document (max 500 chars) on the next line(s), prefixed with ""SUMMARY: ""." & vbCrLf &
            "3. A comma-separated list of 10-20 keywords on the last line, prefixed with ""KEYWORDS: ""." & vbCrLf &
            "Output ONLY these three sections. No markdown, no numbering, no extra text."

        ''' <summary>
        ''' Indexes a single document, triggers agentic Wiki generation if requested, and returns a KnowledgeEntry.
        ''' </summary>
        ''' <param name="filePath">Full or environment-variable path to the raw source file.</param>
        ''' <param name="kbRootPath">The root directory of the current Knowledge Store.</param>
        ''' <param name="context">Shared context for configuration.</param>
        ''' <param name="useLLMIndex">When True, invokes the Wiki Agent to generate persistent Markdown structure.</param>
        ''' <returns>A populated KnowledgeEntry, or Nothing if the file cannot be read.</returns>
        Public Shared Async Function IndexDocumentAsync(
                filePath As String,
                kbRootPath As String,
                context As ISharedContext,
                Optional useLLMIndex As Boolean = False,
                Optional isBackground As Boolean = False) As Task(Of KnowledgeStoreManager.KnowledgeEntry)

            Dim expandedPath = SharedMethods.ExpandEnvironmentVariables(If(filePath, "").Trim().Trim(""""c, "'"c))
            If Not File.Exists(expandedPath) Then Return Nothing

            ' Check extension
            Dim ext = Path.GetExtension(expandedPath).ToLowerInvariant()
            If Not SupportedExtensions.Contains(ext) Then Return Nothing

            ' Ensure the physical Raw directory exists to persist the document if it isn't there already
            Dim rawPath = Path.Combine(kbRootPath, ".redink", "Raw")
            If Not Directory.Exists(rawPath) Then Directory.CreateDirectory(rawPath)

            Dim finalSourcePath As String = expandedPath
            If Not expandedPath.StartsWith(rawPath, StringComparison.OrdinalIgnoreCase) Then
                Try
                    finalSourcePath = BuildRawSnapshotPath(kbRootPath, expandedPath)

                    Dim finalDir = Path.GetDirectoryName(finalSourcePath)
                    If Not String.IsNullOrWhiteSpace(finalDir) AndAlso Not Directory.Exists(finalDir) Then
                        Directory.CreateDirectory(finalDir)
                    End If

                    File.Copy(expandedPath, finalSourcePath, overwrite:=True)
                Catch ex As Exception
                    finalSourcePath = expandedPath
                End Try
            End If

            ' Read content using robust advanced extractor
            Dim content As String = ""
            Try
                content = Await ReadFileForIndexAsync(finalSourcePath, context)
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"KnowledgeIndexer: Error reading '{finalSourcePath}': {ex.Message}")
                KnowledgeWikiService.LogWikiError(kbRootPath, finalSourcePath, $"File could not be read: {ex.Message}")
                Return Nothing
            End Try

            If String.IsNullOrWhiteSpace(content) Then
                KnowledgeWikiService.LogWikiError(kbRootPath, finalSourcePath, "File resulted in empty text payload.")
                Return Nothing
            End If

            ' Default: local extraction (Keyword / Regex)
            Dim title = ExtractTitle(content, finalSourcePath)
            Dim summary = BuildSummary(content)
            Dim keywords = ExtractKeywords(content)

            ' LLM enrichment — Triggers the Wiki Agent Architecture
            If useLLMIndex Then
                Try
                    ' 1. Orchestrator: Generate physical Wiki Page and update vector database
                    Await KnowledgeWikiService.IngestSourceAsync(
                        kbRootPath,
                        If(File.Exists(expandedPath), expandedPath, finalSourcePath),
                        context,
                        isBackground:=isBackground)

                    ' 2. Metadata enrichment
                    Dim backupConfig As ModelConfig = Nothing
                    Dim useAlternateAPI As Boolean = False
                    Dim llmResponse As String = ""

                    Try
                        If Not String.IsNullOrWhiteSpace(context.INI_AlternateModelPath) Then
                            backupConfig = SharedMethods.GetCurrentConfig(context)
                            If SharedMethods.GetSpecialTaskModel(context, context.INI_AlternateModelPath, "KnowledgeStore") Then
                                useAlternateAPI = True
                            Else
                                backupConfig = Nothing
                            End If
                        End If

                        Dim llmAttempts As Integer = Math.Max(1, KnowledgeWikiService.IngestionLlmEmptyRetryCount + 1)

                        For attempt As Integer = 1 To llmAttempts
                            If isBackground Then
                                If KnowledgeWikiService.IngestionLlmThrottleBackgroundMs > 0 Then
                                    Await Task.Delay(KnowledgeWikiService.IngestionLlmThrottleBackgroundMs).ConfigureAwait(False)
                                End If
                            Else
                                If KnowledgeWikiService.IngestionLlmThrottleForegroundMs > 0 Then
                                    Await Task.Delay(KnowledgeWikiService.IngestionLlmThrottleForegroundMs).ConfigureAwait(False)
                                End If
                            End If

                            If content.Length <= KnowledgeWikiService.MaxChunkChars Then
                                llmResponse = Await SharedMethods.LLM(
                                    context:=context,
                                    promptSystem:=LLMIndexPrompt,
                                    promptUser:=content,
                                    UseSecondAPI:=useAlternateAPI,
                                    Hidesplash:=True).ConfigureAwait(False)
                            Else
                                llmResponse = Await KnowledgeWikiService.ProcessLargeTextInChunksAsync(
                                    context:=context,
                                    systemPrompt:=LLMIndexPrompt,
                                    fullText:=content,
                                    useAlternateAPI:=useAlternateAPI,
                                    fileName:=Path.GetFileName(finalSourcePath),
                                    isBackground:=isBackground).ConfigureAwait(False)
                            End If

                            If Not String.IsNullOrWhiteSpace(llmResponse) Then
                                Exit For
                            End If

                            If attempt < llmAttempts AndAlso KnowledgeWikiService.IngestionLlmEmptyRetryDelayMs > 0 Then
                                Await Task.Delay(KnowledgeWikiService.IngestionLlmEmptyRetryDelayMs).ConfigureAwait(False)
                            End If
                        Next

                    Finally
                        If backupConfig IsNot Nothing Then
                            SharedMethods.RestoreDefaults(context, backupConfig)
                        End If
                    End Try

                    If Not String.IsNullOrWhiteSpace(llmResponse) Then
                        Dim parsed = ParseLLMIndexResponse(llmResponse)
                        If Not String.IsNullOrWhiteSpace(parsed.Title) Then title = parsed.Title
                        If Not String.IsNullOrWhiteSpace(parsed.Summary) Then summary = parsed.Summary
                        If parsed.Keywords IsNot Nothing AndAlso parsed.Keywords.Length > 0 Then keywords = parsed.Keywords
                    Else
                        KnowledgeWikiService.LogWikiError(kbRootPath, finalSourcePath, "LLM catalog generation returned empty response.")
                    End If
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine($"KnowledgeIndexer: Agent generation failed for '{finalSourcePath}': {ex.Message}")
                    KnowledgeWikiService.LogWikiError(kbRootPath, finalSourcePath, $"Agent Generation exception: {ex.Message}")
                End Try
            End If

            ' Create entry for the ephemeral catalog indexer
            Dim entry As New KnowledgeStoreManager.KnowledgeEntry() With {
                    .FilePath = KnowledgeStoreManifest.NormalizeStoredPathForStore(kbRootPath, expandedPath),
                    .Title = title,
                    .Summary = summary,
                    .Keywords = keywords,
                    .ContentSnapshot = If(content.Length > MaxContentChars,
                                          content.Substring(0, MaxContentChars),
                                          content),
                    .IndexedDate = DateTime.Now,
                    .Tags = Nothing,
                    .IsFromCentralIndex = False
                }

            Return entry
        End Function


        Private Shared Function BuildRawSnapshotPath(kbRootPath As String, originalPath As String) As String
            Dim rawRoot = Path.Combine(kbRootPath, ".redink", "Raw")
            Dim relativePath As String = Path.GetFileName(originalPath)

            Try
                Dim fullKbRoot = Path.GetFullPath(kbRootPath)
                Dim fullOriginal = Path.GetFullPath(originalPath)

                If fullOriginal.StartsWith(fullKbRoot & Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) Then
                    relativePath = GetRelativePathCompat(fullKbRoot, fullOriginal)
                End If
            Catch
            End Try

            relativePath = relativePath.TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            Return Path.Combine(rawRoot, relativePath)
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
                Return Path.GetFileName(targetPath)
            End Try
        End Function


        ''' <summary>
        ''' Parses the structured LLM response into title, summary, and keywords.
        ''' </summary>
        Private Shared Function ParseLLMIndexResponse(response As String) As (Title As String, Summary As String, Keywords As String())
            Dim title As String = ""
            Dim summary As String = ""
            Dim keywords As String() = Nothing

            Dim lines = response.Split({vbCrLf, vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)

            Dim summaryLines As New List(Of String)()
            Dim inSummary As Boolean = False

            For Each line In lines
                Dim trimmed = line.Trim()

                If trimmed.StartsWith("TITLE:", StringComparison.OrdinalIgnoreCase) Then
                    title = trimmed.Substring("TITLE:".Length).Trim()
                    inSummary = False
                ElseIf trimmed.StartsWith("SUMMARY:", StringComparison.OrdinalIgnoreCase) Then
                    summaryLines.Add(trimmed.Substring("SUMMARY:".Length).Trim())
                    inSummary = True
                ElseIf trimmed.StartsWith("KEYWORDS:", StringComparison.OrdinalIgnoreCase) Then
                    inSummary = False
                    Dim kwText = trimmed.Substring("KEYWORDS:".Length).Trim()
                    keywords = kwText.Split({","c, ";"c}, StringSplitOptions.RemoveEmptyEntries).
                        Select(Function(k) k.Trim().ToLowerInvariant()).
                        Where(Function(k) k.Length >= 2).
                        Distinct().
                        Take(25).
                        ToArray()
                ElseIf inSummary Then
                    summaryLines.Add(trimmed)
                End If
            Next

            summary = String.Join(" ", summaryLines).Trim()
            If summary.Length > MaxSummaryChars Then
                summary = summary.Substring(0, MaxSummaryChars) & "..."
            End If

            Return (title, summary, keywords)
        End Function

        ''' <summary>
        ''' Reads file content using appropriate advanced tooling based on extension.
        ''' </summary>
        Private Shared Async Function ReadFileForIndexAsync(expandedPath As String, context As ISharedContext) As Task(Of String)
            Dim ext = Path.GetExtension(expandedPath).ToLowerInvariant()

            Select Case ext
                Case ".txt", ".ini", ".csv", ".log", ".json", ".xml", ".html", ".htm",
                     ".md", ".yaml", ".yml", ".vb", ".cs", ".js", ".ts", ".py",
                     ".java", ".cpp", ".c", ".h", ".sql"
                    Return IO.File.ReadAllText(expandedPath, System.Text.Encoding.UTF8)

                Case ".docx"
                    Return SharedMethods.ReadDocxSandboxed(expandedPath)

                Case ".pdf"
                    ' Utilizes the robust PDF extraction logic, forcing silent automatic OCR for high quality results
                    Dim pdfResult = Await SharedMethods.ReadPdfAsTextEx(expandedPath, ReturnErrorInsteadOfEmpty:=False, DoOCR:=True, AskUser:=False, context:=context)
                    Return If(pdfResult.Content, "")

                Case ".rtf"
                    Try
                        Return SharedMethods.LoadRtfFileAsPlainText(expandedPath)
                    Catch
                        Return IO.File.ReadAllText(expandedPath, System.Text.Encoding.UTF8)
                    End Try

                Case ".xlsx"
                    Return SharedMethods.ReadXlsxSandboxed(expandedPath)

                Case ".pptx"
                    Return SharedMethods.ReadPptxSandboxed(expandedPath)

                Case Else
                    Return IO.File.ReadAllText(expandedPath, System.Text.Encoding.UTF8)
            End Select
        End Function

        ''' <summary>
        ''' Extracts a title from the content.
        ''' </summary>
        Private Shared Function ExtractTitle(content As String, filePath As String) As String
            Dim headingMatch = Regex.Match(content, "^#{1,3}\s+(.+)$", RegexOptions.Multiline)
            If headingMatch.Success Then
                Dim heading = headingMatch.Groups(1).Value.Trim()
                If heading.Length > 0 AndAlso heading.Length <= 200 Then Return heading
            End If

            Dim lines = content.Split({vbCrLf, vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)
            For Each line In lines
                Dim trimmed = line.Trim()
                If trimmed.Length >= 3 AndAlso trimmed.Length <= 200 Then
                    Return trimmed
                End If
            Next

            Return Path.GetFileNameWithoutExtension(filePath)
        End Function

        ''' <summary>
        ''' Builds a short summary from the document content.
        ''' </summary>
        Private Shared Function BuildSummary(content As String) As String
            Dim cleaned = Regex.Replace(content, "\s+", " ").Trim()
            If cleaned.Length <= MaxSummaryChars Then Return cleaned
            Dim cutoff = cleaned.LastIndexOf(" "c, MaxSummaryChars)
            If cutoff < MaxSummaryChars \ 2 Then cutoff = MaxSummaryChars
            Return cleaned.Substring(0, cutoff).Trim() & "..."
        End Function

        ''' <summary>
        ''' Extracts keywords using frequency analysis.
        ''' </summary>
        Private Shared Function ExtractKeywords(content As String) As String()
            Dim stopWords As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
                "the", "a", "an", "and", "or", "but", "in", "on", "at", "to", "for",
                "of", "with", "by", "from", "is", "are", "was", "were", "be", "been",
                "being", "have", "has", "had", "do", "does", "did", "will", "would",
                "could", "should", "may", "might", "shall", "can", "this", "that",
                "these", "those", "it", "its", "not", "no", "nor", "as", "if", "then",
                "than", "so", "such", "der", "die", "das", "und", "oder", "in", "von",
                "zu", "den", "dem", "des", "ein", "eine", "einer", "einem", "einen",
                "mit", "auf", "ist", "sind", "war", "wird", "werden", "hat", "haben",
                "le", "la", "les", "de", "du", "des", "un", "une", "et", "ou", "en"
            }

            Dim words = Regex.Matches(content, "\b[a-zA-ZÀ-ÿ]{3,}\b")
            Dim freq As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
            For Each m As Match In words
                Dim w = m.Value
                If stopWords.Contains(w) Then Continue For
                If Not freq.ContainsKey(w) Then freq(w) = 0
                freq(w) += 1
            Next

            Return freq.OrderByDescending(Function(kv) kv.Value).
                Take(20).
                Select(Function(kv) kv.Key.ToLowerInvariant()).
                ToArray()
        End Function

        ''' <summary>
        ''' Reads source text using the same robust extraction logic used for indexing.
        ''' </summary>
        Public Shared Async Function ReadSourceTextAsync(filePath As String, context As ISharedContext) As Task(Of String)
            Dim expandedPath = SharedMethods.ExpandEnvironmentVariables(If(filePath, "").Trim().Trim(""""c, "'"c))
            If String.IsNullOrWhiteSpace(expandedPath) OrElse Not File.Exists(expandedPath) Then Return ""
            Return Await ReadFileForIndexAsync(expandedPath, context).ConfigureAwait(False)
        End Function

    End Class
End Namespace