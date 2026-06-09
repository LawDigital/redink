' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.Commands.Freestyle.vb
' Purpose: Implements freestyle LLM prompt execution with support for markup,
'          comments, slides, library search, internet search, and various
'          output modes (clipboard, pane, in-place replacement, etc.).
'
' Architecture:
'  - Command Entry Points: FreeStyleNM (normal model), FreeStyleAM (alternate model),
'    FreeStyleRepeat (re-execute last freestyle command with saved settings).
'  - Prompt Prefix System: User prompts can start with prefixes to control output mode:
'    * Markup prefixes (markup:, markupw:, markupdiff:, markupdiffw:, markupregex:)
'    * Output prefixes (clip:, newdoc:, pane:, slides:, bubbles:, pushback:)
'    * Action prefixes (replace:, add:, pure:)
'  - Trigger System: In-prompt triggers modify behavior (e.g., {all}, {lib}, {net},
'    {chunk}, {doc}, {mystyle}, {object}, {multimodel}).
'  - Model Selection: Supports primary/secondary models with optional alternate model
'    configuration and multi-model execution.
'  - Tooling Support: Allows model tooling selection when supported by the model.
'  - Format Preservation: Configurable formatting retention (character/paragraph level)
'    with special handling for fields and styles.
'  - External Content Integration: Supports embedding external files ({doc} trigger),
'    additional Word documents ({adddoc} trigger), library/internet search results,
'    and custom style prompts (MyStyle).
'  - Special Commands: Handles utility commands (encode, decode, version, reset, etc.)
'    for configuration and diagnostics.
'  - Progress & Cancellation: User can abort long-running operations.
'  - External Dependencies: SharedLibrary.SharedMethods for UI, LLM calls, file I/O;
'    NetOffice.PowerPointApi for slide deck manipulation; DocumentFormat.OpenXml for
'    document processing.
'
' Notes:
'  - The FreeStyle method is the core orchestrator with extensive parameter parsing.
'  - Configuration settings (INI_*) control default behaviors and available features.
'  - Prompt library integration (INI_PromptLib) provides pre-defined prompt templates.
'  - Track point markup (TPMarkup) allows referencing specific user revisions.
' =============================================================================

Option Explicit On
Option Strict On

Imports System.Diagnostics
Imports System.Globalization
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Office2010.CustomUI
Imports DocumentFormat.OpenXml.Office2016.Drawing.Charts
Imports DocumentFormat.OpenXml.Presentation
Imports DocumentFormat.OpenXml.Wordprocessing
Imports Google.Rpc.Context.AttributeContext.Types
Imports Microsoft.Office.Interop.PowerPoint
Imports Microsoft.Office.Interop.Word
Imports NetOffice.PowerPointApi
Imports SharedLibrary
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports Windows.Media
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn


    ''' <summary>
    ''' Helper class to track file loading results for the Freestyle external file embedding feature.
    ''' </summary>
    Private Class FileLoadingContext
        Public Property GlobalDocumentCounter As Integer = 0
        Public Property LoadedFiles As New List(Of Tuple(Of String, Integer))() ' (path, charCount)
        Public Property FailedFiles As New List(Of String)()
        Public Property IgnoredFilesPerDir As New Dictionary(Of String, List(Of String))(StringComparer.OrdinalIgnoreCase) ' directory -> list of ignored file paths

        ''' <summary>OCR mode: 0=not set, 1=enable for all, 2=ask individually</summary>
        Public Property DirectoryOCRMode As Integer = 0

        ''' <summary>Tracks the current directory being processed (for including path in output).</summary>
        Public Property CurrentDirectoryPath As String = ""

        ''' <summary>Total number of files expected to be loaded (for determining if filename should be included).</summary>
        Public Property ExpectedFileCount As Integer = 0

        ''' <summary>PDFs that heuristics suggest may contain images/scanned content but OCR was not performed.</summary>
        Public Property PdfsWithPossibleImages As New List(Of String)()

        ''' <summary>Files that returned empty/whitespace content from GetFileContentEx.</summary>
        Public Property EmptyContentFiles As New List(Of String)()

        ''' <summary>Maximum files to load from a single directory.</summary>
        Public Const MaxFilesPerDirectory As Integer = 50

        ''' <summary>Whether to recurse into subdirectories when loading a directory.</summary>
        Public Const RecurseDirectories As Boolean = False

        ''' <summary>Ask user confirmation if directory has more than this many files.</summary>
        Public Const ConfirmDirectoryFileCount As Integer = 10

        ''' <summary>
        ''' Supported file extensions for directory scanning, aligned with
        ''' <see cref="GetFileContentEx"/> (text-based, Office, PDF, email, and binary/media).
        ''' </summary>
        Public Shared ReadOnly SupportedExtensions As String() = {
            ".txt", ".rtf", ".ini", ".csv", ".log",
            ".json", ".xml", ".html", ".htm",
            ".md", ".yaml", ".yml",
            ".vb", ".cs", ".js", ".ts", ".py", ".java", ".cpp", ".c", ".h", ".sql",
            ".doc", ".docx", ".xlsx", ".pptx",
            ".pdf",
            ".eml", ".msg",
            ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff", ".tif", ".webp", ".svg",
            ".mp3", ".wav", ".ogg", ".flac", ".m4a", ".aac", ".wma", ".opus", ".webm",
            ".mp4", ".avi", ".mkv", ".mov", ".wmv"
        }
    End Class

    ''' <summary>
    ''' Checks if a trigger placeholder at a given index is wrapped in XML tags.
    ''' </summary>
    ''' <param name="prompt">The prompt string to check.</param>
    ''' <param name="idx">The index of the trigger in the prompt.</param>
    ''' <param name="trigger">The trigger text to look for.</param>
    ''' <returns>True if the trigger is wrapped in XML tags, False otherwise.</returns>
    Private Function IsWrappedInXml(prompt As String, idx As Integer, trigger As String) As Boolean
        Dim wrappedPattern As String = "<(?<name>[A-Za-z][\w\-]*)\b[^>]*>\s*" & Regex.Escape(trigger) & "\s*</\k<name>>"
        Dim matches As MatchCollection = Regex.Matches(prompt, wrappedPattern, RegexOptions.IgnoreCase)
        For Each m As Match In matches
            If idx >= m.Index AndAlso idx < m.Index + m.Length Then
                Return True
            End If
        Next
        Return False
    End Function


    ''' <summary>
    ''' Loads content from a URL using WebView2 and wraps it in document tags.
    ''' </summary>
    ''' <param name="url">The URL to retrieve content from.</param>
    ''' <param name="ctx">The file loading context for tracking state.</param>
    ''' <returns>The content wrapped in document tags, or empty string on failure.</returns>
    Private Async Function LoadUrlContentAsync(url As String, ctx As FileLoadingContext) As Task(Of String)
        Try
            ' Retrieve content using WebView2
            Dim content As String = Await RetrieveWebsiteContent_WebView2(url, 0)

            If String.IsNullOrWhiteSpace(content) Then
                ctx.FailedFiles.Add(url)
                Return ""
            End If

            ctx.GlobalDocumentCounter += 1
            Dim docNum As Integer = ctx.GlobalDocumentCounter
            ctx.LoadedFiles.Add(Tuple.Create(url, content.Length))

            ' Always include URL info since it's a web resource
            Dim openTag As String = "<document" & docNum.ToString() & " url=""" & System.Security.SecurityElement.Escape(url) & """>"
            Dim closeTag As String = "</document" & docNum.ToString() & ">"

            Return openTag & content & closeTag

        Catch ex As Exception
            ctx.FailedFiles.Add(url)
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Asynchronously loads a single file and returns its content wrapped in numbered
    ''' XML document tags (e.g. <c>&lt;document1&gt;...&lt;/document1&gt;</c>).
    ''' </summary>
    ''' <remarks>
    ''' <para>
    ''' Delegates to <see cref="GetFileContentEx"/> for the actual file reading, which
    ''' supports text files, RTF, Word (.doc/.docx), Excel (.xlsx), PowerPoint (.pptx),
    ''' PDF (with optional OCR), and email (.eml/.msg).
    ''' </para>
    ''' <para>
    ''' OCR behaviour is determined by <paramref name="isFromDirectory"/> and the
    ''' <see cref="FileLoadingContext.DirectoryOCRMode"/> setting:
    ''' <list type="bullet">
    '''   <item>Mode 1 — OCR enabled for all files without prompting.</item>
    '''   <item>Mode 2 — OCR enabled, user asked per file.</item>
    '''   <item>Mode 3 — OCR skipped entirely.</item>
    '''   <item>Mode 0 / not set — OCR enabled with per-file user prompt (default).</item>
    ''' </list>
    ''' For individually selected files (<paramref name="isFromDirectory"/> = False),
    ''' OCR is always offered with a user prompt when available.
    ''' </para>
    ''' <para>
    ''' When <paramref name="isWrapped"/> is True, the raw file content is returned
    ''' without document tags (the caller already provides XML wrapping). Otherwise,
    ''' the content is enclosed in <c>&lt;documentN&gt;</c> tags that include optional
    ''' <c>directory</c> and <c>filename</c> attributes when multiple files are being
    ''' loaded.
    ''' </para>
    ''' <para>
    ''' Tracking: successfully loaded files are added to <see cref="FileLoadingContext.LoadedFiles"/>;
    ''' files that return empty content go to <see cref="FileLoadingContext.EmptyContentFiles"/>;
    ''' files that throw exceptions go to <see cref="FileLoadingContext.FailedFiles"/>;
    ''' PDFs flagged as potentially incomplete go to <see cref="FileLoadingContext.PdfsWithPossibleImages"/>.
    ''' </para>
    ''' </remarks>
    ''' <param name="filePath">Absolute path to the file to load.</param>
    ''' <param name="isWrapped">
    ''' When True, the trigger was already wrapped in user-supplied XML tags;
    ''' only the raw file content is returned (no <c>&lt;documentN&gt;</c> envelope).
    ''' </param>
    ''' <param name="ctx">Shared <see cref="FileLoadingContext"/> for document numbering,
    ''' OCR mode, and result tracking across all files in the current operation.</param>
    ''' <param name="isFromDirectory">
    ''' True when this file is being loaded as part of a directory scan (uses the
    ''' directory-level OCR mode); False for individually selected files.
    ''' </param>
    ''' <returns>
    ''' The file content (optionally wrapped in XML document tags), or an empty string
    ''' if the file could not be read or contained no extractable text.
    ''' </returns>
    Private Async Function LoadSingleFileAsync(filePath As String, isWrapped As Boolean, ctx As FileLoadingContext, Optional isFromDirectory As Boolean = False, Optional askWorksheetSelection As Boolean = False) As Task(Of String)
        Try
            Dim doOCR As Boolean = False
            Dim askUser As Boolean = False
            ' Check if OCR is available at all
            Dim ocrAvailable As Boolean = SharedMethods.IsOcrAvailable(_context)

            ' Determine OCR behavior based on context and availability
            If ocrAvailable Then
                If isFromDirectory Then
                    If ctx.DirectoryOCRMode = 1 Then
                        doOCR = True
                        askUser = False
                    ElseIf ctx.DirectoryOCRMode = 2 Then
                        doOCR = True
                        askUser = True
                    ElseIf ctx.DirectoryOCRMode = 3 Then
                        ' Skip OCR entirely
                        doOCR = False
                        askUser = False
                    Else
                        doOCR = True
                        askUser = True
                    End If
                Else
                    doOCR = True
                    askUser = True
                End If
            Else
                doOCR = False
                askUser = False
            End If

            ' Keep normal file loading silent, but allow worksheet picker for explicit single-workbook loads
            Dim silentLoad As Boolean = Not askWorksheetSelection

            ' Load file content with determined OCR settings            
            Dim fileResult = Await GetFileContentEx(
                optionalFilePath:=filePath,
                Silent:=silentLoad,
                DoOCR:=doOCR,
                AskUser:=askUser,
                AskWorksheetSelection:=askWorksheetSelection,
                OcrAdditionalInstruction:=Add_OcrMarkdownInstruction
            )

            ' Track PDFs that may have incomplete content
            If fileResult.PdfMayBeIncomplete Then
                ctx.PdfsWithPossibleImages.Add(filePath)
            End If

            If String.IsNullOrWhiteSpace(fileResult.Content) Then
                ctx.EmptyContentFiles.Add(filePath)
                Return ""
            End If

            ctx.GlobalDocumentCounter += 1
            Dim docNum As Integer = ctx.GlobalDocumentCounter
            ctx.LoadedFiles.Add(Tuple.Create(filePath, fileResult.Content.Length))

            If isWrapped Then
                Return fileResult.Content
            Else
                Dim includeFileInfo As Boolean = (ctx.ExpectedFileCount > 1) OrElse isFromDirectory

                Dim openTag As String
                Dim closeTag As String = "</document" & docNum.ToString() & ">"

                If includeFileInfo Then
                    Dim fileName As String = Path.GetFileName(filePath)
                    If isFromDirectory AndAlso Not String.IsNullOrEmpty(ctx.CurrentDirectoryPath) Then
                        openTag = "<document" & docNum.ToString() & " directory=""" & ctx.CurrentDirectoryPath & """ filename=""" & fileName & """>"
                    Else
                        openTag = "<document" & docNum.ToString() & " filename=""" & fileName & """>"
                    End If
                Else
                    openTag = "<document" & docNum.ToString() & ">"
                End If

                Return openTag & fileResult.Content & closeTag
            End If
        Catch ex As Exception
            ctx.FailedFiles.Add(filePath)
        Return ""
        End Try
    End Function




    ''' <summary>
    ''' Loads all supported files from a directory and returns their combined content.
    ''' Prompts for OCR only when there are multiple PDF files in the directory.
    ''' </summary>
    ''' <param name="dirPath">The directory path to load files from.</param>
    ''' <param name="isWrapped">If True, the directory trigger was already wrapped in XML (not used for individual files).</param>
    ''' <param name="ctx">The file loading context for tracking state.</param>
    ''' <param name="ensureProgressBar">
    ''' When True and the progress bar has not yet been started by the caller,
    ''' this method will start it before loading files. The caller sets this to
    ''' True when it defers progress bar creation to avoid showing it during
    ''' user-interaction dialogs.
    ''' </param>
    ''' <returns>Combined content of all files, or "ABORT" if user cancelled, or empty string on error.</returns>
    Private Async Function LoadDirectoryFilesAsync(dirPath As String, isWrapped As Boolean, ctx As FileLoadingContext, Optional ensureProgressBar As Boolean = False) As Task(Of String)
        Try
            If Not Directory.Exists(dirPath) Then
                ctx.FailedFiles.Add(dirPath)
                Return ""
            End If

            Dim searchOption As SearchOption = If(FileLoadingContext.RecurseDirectories, SearchOption.AllDirectories, SearchOption.TopDirectoryOnly)
            Dim allFiles As String() = Directory.GetFiles(dirPath, "*.*", searchOption)

            ' Filter to supported extensions
            Dim supportedFiles As New List(Of String)()
            Dim ignoredFiles As New List(Of String)()

            For Each f In allFiles
                Dim ext As String = Path.GetExtension(f).ToLowerInvariant()
                If FileLoadingContext.SupportedExtensions.Contains(ext) Then
                    supportedFiles.Add(f)
                Else
                    ignoredFiles.Add(f)
                End If
            Next

            If ignoredFiles.Count > 0 Then
                ctx.IgnoredFilesPerDir(dirPath) = ignoredFiles
            End If

            ' Check if too many files
            If supportedFiles.Count > FileLoadingContext.MaxFilesPerDirectory Then
                Dim truncateAnswer As Integer = ShowCustomYesNoBox(
                $"The directory '{dirPath}' contains {supportedFiles.Count} supported files, but the maximum is {FileLoadingContext.MaxFilesPerDirectory}." & vbCrLf & vbCrLf &
                $"Only the first {FileLoadingContext.MaxFilesPerDirectory} files will be loaded. Continue?",
                "Yes, continue", "No, abort")
                If truncateAnswer <> 1 Then
                    Return "ABORT"
                End If
                supportedFiles = supportedFiles.Take(FileLoadingContext.MaxFilesPerDirectory).ToList()
            ElseIf supportedFiles.Count > FileLoadingContext.ConfirmDirectoryFileCount Then
                Dim confirmAnswer As Integer = ShowCustomYesNoBox(
                $"The directory '{dirPath}' contains {supportedFiles.Count} files to load. Continue?",
                "Yes, continue", "No, abort")
                If confirmAnswer <> 1 Then
                    Return "ABORT"
                End If
            End If

            If supportedFiles.Count = 0 Then
                ShowCustomMessageBox($"No supported files found in directory '{dirPath}'.")
                Return ""
            End If

            ' Count PDF files in the directory
            Dim pdfFiles As List(Of String) = supportedFiles.Where(Function(f) Path.GetExtension(f).Equals(".pdf", StringComparison.OrdinalIgnoreCase)).ToList()
            Dim pdfCount As Integer = pdfFiles.Count

            ' Reset directory OCR mode for this directory
            ctx.DirectoryOCRMode = 0

            ' Only ask about OCR if there are multiple PDF files in the directory AND OCR is available
            If pdfCount > 1 AndAlso SharedMethods.IsOcrAvailable(_context) Then
                Dim skipOcrChosen As Boolean = False

                Dim ocrChoice As Integer = ShowCustomYesNoBox(
                    $"The directory contains {pdfCount} PDF files." & vbCrLf & vbCrLf &
                    "How would you like to handle OCR for these PDF files?" & vbCrLf &
                    "(OCR uses the LLM to extract text from scanned documents and can take a moment per file.)",
                    "Enable OCR for all PDFs",
                    "Choose individually for each PDF",
                    extraButtonText:="Skip OCR entirely",
                    extraButtonAction:=Sub()
                                           skipOcrChosen = True
                                       End Sub,
                    CloseAfterExtra:=True)

                If skipOcrChosen Then
                    ' Skip OCR entirely
                    ctx.DirectoryOCRMode = 3
                ElseIf ocrChoice = 0 Then
                    ' User closed the dialog - abort
                    Return "ABORT"
                ElseIf ocrChoice = 1 Then
                    ' Enable OCR for all without asking
                    ctx.DirectoryOCRMode = 1
                ElseIf ocrChoice = 2 Then
                    ' Ask individually for each PDF
                    ctx.DirectoryOCRMode = 2
                End If
            End If

            ' Set the current directory path for inclusion in document tags
            ctx.CurrentDirectoryPath = dirPath

            ' Update expected file count to include these directory files
            ctx.ExpectedFileCount += supportedFiles.Count

            ' Start progress bar now — all user dialogs are done, actual loading begins
            If ensureProgressBar Then
                ProgressBarModule.CancelOperation = False
                ProgressBarModule.GlobalProgressMax = supportedFiles.Count
                ProgressBarModule.GlobalProgressValue = 0
                ProgressBarModule.GlobalProgressLabel = "Loading directory files..."
                ShowProgressBarInSeparateThread(AN & " File Loading (may include OCR)", "Loading directory files...")
            End If

            ' Load all files - each file gets its own document wrapper with incrementing counter
            Dim resultBuilder As New System.Text.StringBuilder()
            Dim dirFileCount As Integer = supportedFiles.Count
            Dim dirFileIndex As Integer = 0
            For Each filePath In supportedFiles
                ' Check for user cancellation
                If ProgressBarModule.CancelOperation Then
                    Return "ABORT"
                End If

                dirFileIndex += 1
                ProgressBarModule.GlobalProgressLabel = $"Loading file {dirFileIndex} of {dirFileCount}: {Path.GetFileName(filePath)}..."
                ProgressBarModule.GlobalProgressMax = dirFileCount
                ProgressBarModule.GlobalProgressValue = dirFileIndex - 1

                ' Always wrap individual files from directory with document tags
                ' Pass isFromDirectory=True to use directory OCR mode for PDFs
                Dim fileContent As String = Await LoadSingleFileAsync(filePath, False, ctx, isFromDirectory:=True)
                If Not String.IsNullOrWhiteSpace(fileContent) Then
                    resultBuilder.Append(fileContent)
                End If

                ProgressBarModule.GlobalProgressValue = dirFileIndex
            Next

            ' Clear the current directory path after processing
            ctx.CurrentDirectoryPath = ""

            Return resultBuilder.ToString()
        Catch ex As Exception
            ctx.FailedFiles.Add(dirPath)
            Return ""
        End Try
    End Function


    ''' <summary>
    ''' Processes all external file/directory/URL triggers in the prompt ({doc}, {dir}, {url}, inline URLs, and fixed paths).
    ''' Processes triggers in positional order to maintain correct document numbering.
    ''' </summary>
    ''' <param name="prompt">The prompt to process.</param>
    ''' <returns>A tuple containing: (success flag, modified prompt). If success is False, the operation should abort.</returns>
    Private Async Function ProcessExternalFileTriggers(prompt As String) As Task(Of (Success As Boolean, ModifiedPrompt As String))
        ' Check if any file triggers are present
        Dim hasExtTriggers As Boolean = prompt.IndexOf(ExtTrigger, StringComparison.OrdinalIgnoreCase) >= 0
        Dim hasDirTriggers As Boolean = prompt.IndexOf(ExtDirTrigger, StringComparison.OrdinalIgnoreCase) >= 0
        Dim hasUrlTriggers As Boolean = prompt.IndexOf(ExtUrlTrigger, StringComparison.OrdinalIgnoreCase) >= 0
        Dim hasFixedPaths As Boolean = False

        ' Pattern for inline URLs like {https://example.com} or {http://example.com}
        Dim inlineUrlPattern As String = "\{(https?://[^}]+)\}"
        Dim hasInlineUrls As Boolean = Regex.IsMatch(prompt, inlineUrlPattern, RegexOptions.IgnoreCase)

        ' Prepare fixed path pattern
        Dim fixedPrefix As String = ""
        Dim fixedSuffix As String = ""
        Dim patternFixed As String = ""

        If Not String.IsNullOrWhiteSpace(ExtTriggerFixed) AndAlso
       ExtTriggerFixed.IndexOf("[path]", StringComparison.OrdinalIgnoreCase) >= 0 Then
            Dim pathTokenIndex As Integer = ExtTriggerFixed.IndexOf("[path]", StringComparison.OrdinalIgnoreCase)
            fixedPrefix = ExtTriggerFixed.Substring(0, pathTokenIndex)
            fixedSuffix = ExtTriggerFixed.Substring(pathTokenIndex + "[path]".Length)
            If Not (String.IsNullOrEmpty(fixedPrefix) AndAlso String.IsNullOrEmpty(fixedSuffix)) Then
                patternFixed = Regex.Escape(fixedPrefix) & "(?<path>.*?)" & Regex.Escape(fixedSuffix)
                hasFixedPaths = Regex.IsMatch(prompt, patternFixed, RegexOptions.IgnoreCase)
            End If
        End If

        ' No triggers present - nothing to do
        If Not hasExtTriggers AndAlso Not hasDirTriggers AndAlso Not hasFixedPaths AndAlso Not hasUrlTriggers AndAlso Not hasInlineUrls Then
            Return (True, prompt)
        End If

        ' Reset cancellation flag from any previous operation
        ProgressBarModule.CancelOperation = False

        ' Create context for tracking
        Dim ctx As New FileLoadingContext()

        ' Count total expected files/triggers to determine if filenames should be included
        ' This is an estimate - actual count may differ after user selections
        Dim extTriggerCount As Integer = Regex.Matches(prompt, Regex.Escape(ExtTrigger), RegexOptions.IgnoreCase).Count
        Dim dirTriggerCount As Integer = Regex.Matches(prompt, Regex.Escape(ExtDirTrigger), RegexOptions.IgnoreCase).Count
        Dim urlTriggerCount As Integer = Regex.Matches(prompt, Regex.Escape(ExtUrlTrigger), RegexOptions.IgnoreCase).Count
        Dim inlineUrlCount As Integer = Regex.Matches(prompt, inlineUrlPattern, RegexOptions.IgnoreCase).Count
        Dim fixedPathCount As Integer = 0
        If Not String.IsNullOrEmpty(patternFixed) Then
            ' Count only fixed path matches that are not already matched by other trigger types
            ' (e.g., the pattern {[path]} -> \{(?<path>.*?)\} would also match {doc}, {dir}, {url})
            For Each m As Match In Regex.Matches(prompt, patternFixed, RegexOptions.IgnoreCase)
                Dim candidatePath As String = If(m.Groups("path").Value, "").Trim()
                ' Remove quotes if present
                If (candidatePath.Length >= 2 AndAlso candidatePath.StartsWith("""") AndAlso candidatePath.EndsWith("""")) OrElse
                   (candidatePath.Length >= 2 AndAlso candidatePath.StartsWith("'") AndAlso candidatePath.EndsWith("'")) Then
                    candidatePath = candidatePath.Substring(1, candidatePath.Length - 2).Trim()
                End If
                ' Skip if this match is actually one of the known triggers
                Dim matchText As String = m.Value
                If String.Equals(matchText, ExtTrigger, StringComparison.OrdinalIgnoreCase) Then Continue For
                If String.Equals(matchText, ExtDirTrigger, StringComparison.OrdinalIgnoreCase) Then Continue For
                If String.Equals(matchText, ExtUrlTrigger, StringComparison.OrdinalIgnoreCase) Then Continue For
                ' Also skip if it doesn't look like a path
                Dim looksLikePath As Boolean = candidatePath.Contains("\") OrElse candidatePath.Contains("/") OrElse candidatePath.Contains(":")
                If Not looksLikePath Then Continue For
                fixedPathCount += 1
            Next
        End If
        ctx.ExpectedFileCount = extTriggerCount + fixedPathCount + urlTriggerCount + inlineUrlCount
        ' Note: dirTriggerCount files are added in LoadDirectoryFilesAsync

        ' NOTE: OCR prompt is handled per-directory when multiple PDFs are found,
        ' or per-file with AskUser=True for individual files/single PDFs

        ' === Track overall file processing progress ===
        Dim totalTriggers As Integer = extTriggerCount + dirTriggerCount + urlTriggerCount + inlineUrlCount + fixedPathCount
        Dim processedTriggers As Integer = 0
        Dim progressBarStarted As Boolean = False

        ' === Process all triggers in positional order ===
        ' We need to find the first trigger of any type, process it, then repeat
        ' This ensures document numbering follows the order in the prompt

        ' searchStartIndex tracks where to begin searching for the next trigger.
        ' After each replacement, it advances past the inserted content so that
        ' file/directory content (which may contain trigger-like text such as {doc},
        ' {dir}, or {path}) is never re-scanned as a new trigger.
        Dim searchStartIndex As Integer = 0

        Dim continueProcessing As Boolean = True
        While continueProcessing
            ' Check for user cancellation via progress bar
            If progressBarStarted AndAlso ProgressBarModule.CancelOperation Then
                Return (False, prompt)
            End If

            ' Find position of each trigger type, starting from searchStartIndex
            Dim extIdx As Integer = prompt.IndexOf(ExtTrigger, searchStartIndex, StringComparison.OrdinalIgnoreCase)
            Dim dirIdx As Integer = prompt.IndexOf(ExtDirTrigger, searchStartIndex, StringComparison.OrdinalIgnoreCase)
            Dim urlIdx As Integer = prompt.IndexOf(ExtUrlTrigger, searchStartIndex, StringComparison.OrdinalIgnoreCase)
            Dim fixedIdx As Integer = -1
            Dim fixedMatch As Match = Nothing
            Dim inlineUrlIdx As Integer = -1
            Dim inlineUrlMatch As Match = Nothing

            ' Check for inline URL pattern {https://...}
            inlineUrlMatch = Regex.Match(prompt, inlineUrlPattern, RegexOptions.IgnoreCase)
            ' Only consider matches at or after searchStartIndex
            While inlineUrlMatch IsNot Nothing AndAlso inlineUrlMatch.Success AndAlso inlineUrlMatch.Index < searchStartIndex
                inlineUrlMatch = inlineUrlMatch.NextMatch()
            End While
            If inlineUrlMatch IsNot Nothing AndAlso inlineUrlMatch.Success Then
                inlineUrlIdx = inlineUrlMatch.Index
            Else
                inlineUrlMatch = Nothing
            End If

            If Not String.IsNullOrEmpty(patternFixed) Then
                fixedMatch = Regex.Match(prompt, patternFixed, RegexOptions.IgnoreCase Or RegexOptions.Singleline)
                ' Only consider matches at or after searchStartIndex
                While fixedMatch IsNot Nothing AndAlso fixedMatch.Success AndAlso fixedMatch.Index < searchStartIndex
                    fixedMatch = fixedMatch.NextMatch()
                End While
                If fixedMatch IsNot Nothing AndAlso fixedMatch.Success Then
                    ' Check if it looks like a path
                    Dim candidatePath As String = If(fixedMatch.Groups("path").Value, "").Trim()
                    ' Remove quotes if present
                    If (candidatePath.Length >= 2 AndAlso candidatePath.StartsWith("""") AndAlso candidatePath.EndsWith("""")) OrElse
                   (candidatePath.Length >= 2 AndAlso candidatePath.StartsWith("'") AndAlso candidatePath.EndsWith("'")) Then
                        candidatePath = candidatePath.Substring(1, candidatePath.Length - 2).Trim()
                    End If
                    Dim looksLikePath As Boolean = candidatePath.Contains("\") OrElse candidatePath.Contains("/") OrElse candidatePath.Contains(":")
                    If looksLikePath Then
                        fixedIdx = fixedMatch.Index
                    Else
                        fixedMatch = Nothing ' Not a path, ignore this match
                    End If
                Else
                    fixedMatch = Nothing
                End If
            End If

            ' Determine which trigger comes first
            Dim minIdx As Integer = Integer.MaxValue
            Dim triggerType As String = ""

            If extIdx >= 0 AndAlso extIdx < minIdx Then
                minIdx = extIdx
                triggerType = "ext"
            End If
            If dirIdx >= 0 AndAlso dirIdx < minIdx Then
                minIdx = dirIdx
                triggerType = "dir"
            End If
            If urlIdx >= 0 AndAlso urlIdx < minIdx Then
                minIdx = urlIdx
                triggerType = "url"
            End If
            If inlineUrlIdx >= 0 AndAlso inlineUrlIdx < minIdx Then
                minIdx = inlineUrlIdx
                triggerType = "inlineurl"
            End If
            If fixedIdx >= 0 AndAlso fixedIdx < minIdx Then
                minIdx = fixedIdx
                triggerType = "fixed"
            End If

            ' No more triggers found
            If triggerType = "" Then
                continueProcessing = False
                Continue While
            End If

            ' Process the first trigger found
            Select Case triggerType
                Case "ext"
                    ' Process {doc} trigger - individual file, use OCR with AskUser=True
                    Globals.ThisAddIn.DragDropFormLabel = "Select a file to include"
                    Dim selectedFile As String = ""

                    Using frm As New DragDropForm(DragDropMode.FileOnly)
                        If frm.ShowDialog() = DialogResult.OK Then
                            selectedFile = frm.SelectedFilePath
                        End If
                    End Using

                    Dim replacementText As String = ""

                    If Not String.IsNullOrWhiteSpace(selectedFile) Then
                        Dim isWrapped As Boolean = IsWrappedInXml(prompt, extIdx, ExtTrigger)
                        Dim askWorksheetSelection As Boolean =
                            Path.GetExtension(selectedFile).Equals(".xlsx", StringComparison.OrdinalIgnoreCase)

                        ' Start progress bar only when no worksheet-selection dialog is expected
                        If Not askWorksheetSelection AndAlso Not progressBarStarted Then
                            progressBarStarted = True
                            ProgressBarModule.CancelOperation = False
                            ProgressBarModule.GlobalProgressMax = System.Math.Max(1, totalTriggers)
                            ProgressBarModule.GlobalProgressValue = 0
                            ProgressBarModule.GlobalProgressLabel = "Loading external files..."
                            ShowProgressBarInSeparateThread(AN & " File Loading (may include OCR)", "Loading external files...")
                        End If

                        If progressBarStarted Then
                            ProgressBarModule.GlobalProgressLabel = $"Loading file {processedTriggers + 1} of {totalTriggers}: {Path.GetFileName(selectedFile)}..."
                            ProgressBarModule.GlobalProgressMax = totalTriggers
                        End If

                        replacementText = Await LoadSingleFileAsync(
                            selectedFile,
                            isWrapped,
                            ctx,
                            isFromDirectory:=False,
                            askWorksheetSelection:=askWorksheetSelection)

                        ' Start progress bar after selection if it was deferred
                        If askWorksheetSelection AndAlso Not progressBarStarted Then
                            progressBarStarted = True
                            ProgressBarModule.CancelOperation = False
                            ProgressBarModule.GlobalProgressMax = System.Math.Max(1, totalTriggers)
                            ProgressBarModule.GlobalProgressValue = processedTriggers
                            ProgressBarModule.GlobalProgressLabel = "Loading external files..."
                            ShowProgressBarInSeparateThread(AN & " File Loading (may include OCR)", "Loading external files...")
                        End If
                    Else
                        Dim answer As Integer = ShowCustomYesNoBox(
                        "No file selected. Do you want to continue or abort?",
                        "Continue", "Abort")
                        If answer <> 1 Then
                            If progressBarStarted Then ProgressBarModule.CancelOperation = True
                            Return (False, prompt)
                        End If
                    End If

                    prompt = prompt.Substring(0, extIdx) & replacementText & prompt.Substring(extIdx + ExtTrigger.Length)
                    ' Advance past the replacement so loaded content is not re-scanned
                    searchStartIndex = extIdx + replacementText.Length
                    processedTriggers += 1
                    If progressBarStarted Then
                        ProgressBarModule.GlobalProgressMax = totalTriggers
                        ProgressBarModule.GlobalProgressValue = processedTriggers
                    End If

                Case "dir"
                    ' Process {dir} trigger - directory processing with OCR prompt for multiple PDFs
                    Globals.ThisAddIn.DragDropFormLabel = "Select a directory to include"
                    Dim selectedDir As String = ""

                    Using frm As New DragDropForm(DragDropMode.DirectoryOnly)
                        If frm.ShowDialog() = DialogResult.OK Then
                            selectedDir = frm.SelectedFilePath
                        End If
                    End Using

                    Dim replacementText As String = ""

                    If Not String.IsNullOrWhiteSpace(selectedDir) Then
                        Dim isWrapped As Boolean = IsWrappedInXml(prompt, dirIdx, ExtDirTrigger)
                        ' Do NOT start progress bar here — LoadDirectoryFilesAsync shows user
                        ' dialogs (file count confirmation, OCR choice) first and will start the
                        ' progress bar itself once all questions have been answered.
                        replacementText = Await LoadDirectoryFilesAsync(selectedDir, isWrapped, ctx, ensureProgressBar:=Not progressBarStarted)
                        progressBarStarted = True

                        If replacementText = "ABORT" Then
                            ProgressBarModule.CancelOperation = True
                            Return (False, prompt)
                        End If
                    Else
                        Dim answer As Integer = ShowCustomYesNoBox(
                        "No directory selected. Do you want to continue or abort?",
                        "Continue", "Abort")
                        If answer <> 1 Then
                            If progressBarStarted Then ProgressBarModule.CancelOperation = True
                            Return (False, prompt)
                        End If
                    End If

                    prompt = prompt.Substring(0, dirIdx) & replacementText & prompt.Substring(dirIdx + ExtDirTrigger.Length)
                    ' Advance past the replacement so loaded content is not re-scanned
                    searchStartIndex = dirIdx + replacementText.Length
                    processedTriggers += 1
                    If progressBarStarted Then
                        ProgressBarModule.GlobalProgressMax = totalTriggers
                        ProgressBarModule.GlobalProgressValue = processedTriggers
                    End If

                Case "url"
                    ' Process {url} trigger - prompt user for a URL to retrieve
                    Dim userUrl As String = SLib.ShowCustomInputBox("Please enter the URL to retrieve content from:", $"{AN} - URL Content", True, "https://")

                    Dim replacementText As String = ""

                    If Not String.IsNullOrWhiteSpace(userUrl) AndAlso Not userUrl.Equals("esc", StringComparison.OrdinalIgnoreCase) Then
                        userUrl = userUrl.Trim()

                        ' Start progress bar on first actual file load (after all user dialogs)
                        If Not progressBarStarted Then
                            progressBarStarted = True
                            ProgressBarModule.CancelOperation = False
                            ProgressBarModule.GlobalProgressMax = System.Math.Max(1, totalTriggers)
                            ProgressBarModule.GlobalProgressValue = 0
                            ProgressBarModule.GlobalProgressLabel = "Loading external files..."
                            ShowProgressBarInSeparateThread(AN & " File Loading (may include OCR)", "Loading external files...")
                        End If
                        ProgressBarModule.GlobalProgressLabel = $"Retrieving URL {processedTriggers + 1} of {totalTriggers}: {userUrl}..."
                        ProgressBarModule.GlobalProgressMax = totalTriggers

                        Try
                            replacementText = Await LoadUrlContentAsync(userUrl, ctx)
                        Catch
                        End Try

                        If String.IsNullOrWhiteSpace(replacementText) Then
                            Dim answer As Integer = ShowCustomYesNoBox(
                                $"Could not retrieve content from '{userUrl}'. Do you want to continue or abort?",
                                "Continue", "Abort")
                            If answer <> 1 Then
                                ProgressBarModule.CancelOperation = True
                                Return (False, prompt)
                            End If
                        End If
                    Else
                        Dim answer As Integer = ShowCustomYesNoBox(
                            "No URL provided. Do you want to continue or abort?",
                            "Continue", "Abort")
                        If answer <> 1 Then
                            If progressBarStarted Then ProgressBarModule.CancelOperation = True
                            Return (False, prompt)
                        End If
                    End If

                    prompt = prompt.Substring(0, urlIdx) & replacementText & prompt.Substring(urlIdx + ExtUrlTrigger.Length)
                    ' Advance past the replacement so loaded content is not re-scanned
                    searchStartIndex = urlIdx + replacementText.Length
                    processedTriggers += 1
                    If progressBarStarted Then
                        ProgressBarModule.GlobalProgressMax = totalTriggers
                        ProgressBarModule.GlobalProgressValue = processedTriggers
                    End If

                Case "inlineurl"
                    ' Process inline URL trigger like {https://example.com}
                    Dim inlineUrl As String = inlineUrlMatch.Groups(1).Value.Trim()

                    ' Start progress bar on first actual file load (after all user dialogs)
                    If Not progressBarStarted Then
                        progressBarStarted = True
                        ProgressBarModule.CancelOperation = False
                        ProgressBarModule.GlobalProgressMax = System.Math.Max(1, totalTriggers)
                        ProgressBarModule.GlobalProgressValue = 0
                        ProgressBarModule.GlobalProgressLabel = "Loading external files..."
                        ShowProgressBarInSeparateThread(AN & " File Loading (may include OCR)", "Loading external files...")
                    End If
                    ProgressBarModule.GlobalProgressLabel = $"Retrieving URL {processedTriggers + 1} of {totalTriggers}: {inlineUrl}..."
                    ProgressBarModule.GlobalProgressMax = totalTriggers

                    Dim replacementText As String = ""

                    Try
                        replacementText = Await LoadUrlContentAsync(inlineUrl, ctx)
                    Catch
                    End Try

                    If String.IsNullOrWhiteSpace(replacementText) Then
                        ctx.FailedFiles.Add(inlineUrl)
                    End If

                    prompt = prompt.Substring(0, inlineUrlMatch.Index) & replacementText & prompt.Substring(inlineUrlMatch.Index + inlineUrlMatch.Length)
                    ' Advance past the replacement so loaded content is not re-scanned
                    searchStartIndex = inlineUrlMatch.Index + replacementText.Length
                    processedTriggers += 1
                    If progressBarStarted Then
                        ProgressBarModule.GlobalProgressMax = totalTriggers
                        ProgressBarModule.GlobalProgressValue = processedTriggers
                    End If

                Case "fixed"
                    ' Process fixed path trigger
                    Dim tokenText As String = fixedMatch.Value
                    Dim candidatePath As String = If(fixedMatch.Groups("path").Value, "").Trim()

                    ' Remove quotes if present
                    If (candidatePath.Length >= 2 AndAlso candidatePath.StartsWith("""") AndAlso candidatePath.EndsWith("""")) OrElse
                   (candidatePath.Length >= 2 AndAlso candidatePath.StartsWith("'") AndAlso candidatePath.EndsWith("'")) Then
                        candidatePath = candidatePath.Substring(1, candidatePath.Length - 2).Trim()
                    End If

                    ' Expand environment variables
                    Dim expandedPath As String = SLib.ExpandEnvironmentVariables(candidatePath)
                    If Not String.IsNullOrWhiteSpace(expandedPath) Then
                        candidatePath = expandedPath
                    End If

                    Dim replacementText As String = ""

                    ' Check if it's wrapped in XML
                    Dim wrappedPatternTemplate As String = "<(?<name>[A-Za-z][\w\-]*)\b[^>]*>[^<]*" & Regex.Escape(tokenText) & "[^<]*</\k<name>>"
                    Dim isWrapped As Boolean = Regex.IsMatch(prompt, wrappedPatternTemplate, RegexOptions.IgnoreCase Or RegexOptions.Singleline)

                    ' Determine if it's a file or directory
                    If File.Exists(candidatePath) Then
                        Dim askWorksheetSelection As Boolean =
                            Path.GetExtension(candidatePath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase)

                        ' Start progress bar only when no worksheet-selection dialog is expected
                        If Not askWorksheetSelection AndAlso Not progressBarStarted Then
                            progressBarStarted = True
                            ProgressBarModule.CancelOperation = False
                            ProgressBarModule.GlobalProgressMax = System.Math.Max(1, totalTriggers)
                            ProgressBarModule.GlobalProgressValue = 0
                            ProgressBarModule.GlobalProgressLabel = "Loading external files..."
                            ShowProgressBarInSeparateThread(AN & " File Loading (may include OCR)", "Loading external files...")
                        End If

                        If progressBarStarted Then
                            ProgressBarModule.GlobalProgressLabel = $"Loading {processedTriggers + 1} of {totalTriggers}: {Path.GetFileName(candidatePath)}..."
                            ProgressBarModule.GlobalProgressMax = totalTriggers
                        End If

                        replacementText = Await LoadSingleFileAsync(
                            candidatePath,
                            isWrapped,
                            ctx,
                            isFromDirectory:=False,
                            askWorksheetSelection:=askWorksheetSelection)

                        ' Start progress bar after selection if it was deferred
                        If askWorksheetSelection AndAlso Not progressBarStarted Then
                            progressBarStarted = True
                            ProgressBarModule.CancelOperation = False
                            ProgressBarModule.GlobalProgressMax = System.Math.Max(1, totalTriggers)
                            ProgressBarModule.GlobalProgressValue = processedTriggers
                            ProgressBarModule.GlobalProgressLabel = "Loading external files..."
                            ShowProgressBarInSeparateThread(AN & " File Loading (may include OCR)", "Loading external files...")
                        End If

                    ElseIf Directory.Exists(candidatePath) Then
                        ' Directory - defer progress bar to LoadDirectoryFilesAsync (user dialogs first)
                        replacementText = Await LoadDirectoryFilesAsync(candidatePath, isWrapped, ctx, ensureProgressBar:=Not progressBarStarted)
                        progressBarStarted = True
                        If replacementText = "ABORT" Then
                            ProgressBarModule.CancelOperation = True
                            Return (False, prompt)
                        End If
                    Else
                        ctx.FailedFiles.Add(candidatePath)
                        replacementText = ""
                    End If

                    prompt = prompt.Substring(0, fixedMatch.Index) & replacementText & prompt.Substring(fixedMatch.Index + fixedMatch.Length)
                    ' Advance past the replacement so loaded content is not re-scanned
                    searchStartIndex = fixedMatch.Index + replacementText.Length
                    processedTriggers += 1
                    If progressBarStarted Then
                        ProgressBarModule.GlobalProgressMax = totalTriggers
                        ProgressBarModule.GlobalProgressValue = processedTriggers
                    End If
            End Select
        End While

        ' Close the progress bar
        If progressBarStarted Then
            ProgressBarModule.CancelOperation = True
        End If

        ' === Show summary and confirm before proceeding ===
        If ctx.LoadedFiles.Count > 0 OrElse ctx.FailedFiles.Count > 0 OrElse ctx.IgnoredFilesPerDir.Count > 0 OrElse ctx.PdfsWithPossibleImages.Count > 0 OrElse ctx.EmptyContentFiles.Count > 0 Then
            Dim summary As New System.Text.StringBuilder()
            summary.AppendLine("File inclusion summary:")
            summary.AppendLine("")

            If ctx.LoadedFiles.Count > 0 Then
                summary.AppendLine($"Successfully loaded ({ctx.LoadedFiles.Count} files):")
                Dim totalChars As Integer = 0
                For Each item In ctx.LoadedFiles
                    summary.AppendLine($"  • {item.Item1} ({item.Item2:N0} chars)")
                    totalChars += item.Item2
                Next
                summary.AppendLine($"  Total: {totalChars:N0} characters")
                summary.AppendLine("")
            End If

            If ctx.PdfsWithPossibleImages.Count > 0 Then
                summary.AppendLine($"⚠ PDFs that may contain images/scans ({ctx.PdfsWithPossibleImages.Count} file(s)):")
                For Each f In ctx.PdfsWithPossibleImages
                    summary.AppendLine($"  • {Path.GetFileName(f)}")
                Next
                summary.AppendLine("  (Text extraction may be incomplete - OCR was not available or not performed)")
                summary.AppendLine("")
            End If

            If ctx.EmptyContentFiles.Count > 0 Then
                summary.AppendLine($"⚠ Files with no extractable content ({ctx.EmptyContentFiles.Count} file(s)):")
                For Each f In ctx.EmptyContentFiles
                    summary.AppendLine($"  • {Path.GetFileName(f)} ({IO.Path.GetExtension(f).ToLowerInvariant()})")
                Next
                summary.AppendLine("  (These files were read but returned no text content)")
                summary.AppendLine("")
            End If

            If ctx.FailedFiles.Count > 0 Then
                summary.AppendLine($"Failed to load ({ctx.FailedFiles.Count} items):")
                For Each f In ctx.FailedFiles
                    summary.AppendLine($"  • {f}")
                Next
                summary.AppendLine("")
            End If

            If ctx.IgnoredFilesPerDir.Count > 0 Then
                Dim totalIgnored As Integer = ctx.IgnoredFilesPerDir.Values.Sum(Function(lst) lst.Count)
                summary.AppendLine($"Skipped due to unsupported file type ({totalIgnored} file(s)):")
                For Each kvp In ctx.IgnoredFilesPerDir
                    summary.AppendLine($"  Directory: {kvp.Key}")
                    ' Group ignored files by extension for readability
                    Dim byExtension = kvp.Value.
                        GroupBy(Function(f) IO.Path.GetExtension(f).ToLowerInvariant()).
                        OrderByDescending(Function(g) g.Count())
                    For Each extGroup In byExtension
                        Dim extLabel As String = If(String.IsNullOrEmpty(extGroup.Key), "(no extension)", extGroup.Key)
                        If extGroup.Count() <= 3 Then
                            ' List individual filenames when there are few
                            For Each f In extGroup
                                summary.AppendLine($"    • {IO.Path.GetFileName(f)}")
                            Next
                        Else
                            ' Summarize when there are many of the same type
                            summary.AppendLine($"    • {extGroup.Count()} {extLabel} files (e.g., {IO.Path.GetFileName(extGroup.First())})")
                        End If
                    Next
                Next
                summary.AppendLine($"  Supported types: {String.Join(", ", FileLoadingContext.SupportedExtensions)}")
                summary.AppendLine("")
            End If

            Dim proceedAnswer As Integer = ShowCustomYesNoBox(
            summary.ToString().TrimEnd() & vbCrLf & vbCrLf & "Do you want to proceed with these files?",
            "Proceed", "Abort")

            If proceedAnswer <> 1 Then
                Return (False, prompt)
            End If
        End If

        Return (True, prompt)
    End Function

    ''' <summary>
    ''' Stores the model configuration from the last freestyle command using alternate model.
    ''' </summary>
    Public Shared LastFreestyleModelConfig As ModelConfig

    ''' <summary>
    ''' Indicates whether the last freestyle command used the alternate model (True) or normal model (False).
    ''' </summary>
    Public Shared LastFreestyleWasAM As Boolean = False

    ''' <summary>
    ''' Stores the prompt text from the last freestyle command for repeat functionality.
    ''' </summary>
    Public Shared LastFreestylePrompt As String = ""

    ''' <summary>
    ''' Executes a freestyle command using the normal (primary) model.
    ''' Saves the command parameters to settings for potential repeat execution.
    ''' </summary>
    Public Async Sub FreeStyleNM()
        If INILoadFail() OrElse Not IsDocumentEditable() Then Return
        FreeStyle(False)

        My.Settings.LastFreestyleModelConfig = Nothing
        My.Settings.LastFreestyleWasAM = False
        My.Settings.LastFreestylePrompt = My.Settings.LastPrompt
        My.Settings.Save()

        Dim result = Globals.Ribbons.Ribbon1.InitializeAppAsync()

    End Sub

    ''' <summary>
    ''' Executes a freestyle command using the alternate (secondary) model.
    ''' Prompts for model selection if alternate model path is configured.
    ''' Saves the command parameters and model configuration to settings for potential repeat execution.
    ''' </summary>
    Public Async Sub FreeStyleAM()
        If INILoadFail() OrElse Not IsDocumentEditable() Then Return

        If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then

            If Not ShowModelSelection(_context, INI_AlternateModelPath) Then
                originalConfigLoaded = False
                Return
            End If

        End If

        LastFreestyleModelConfig = GetCurrentConfig(_context)

        FreeStyle(True)

        My.Settings.LastFreestyleModelConfig = LastFreestyleModelConfig
        My.Settings.LastFreestyleWasAM = True
        My.Settings.LastFreestylePrompt = My.Settings.LastPrompt
        My.Settings.Save()

        Dim result = Globals.Ribbons.Ribbon1.InitializeAppAsync()

    End Sub

    ''' <summary>
    ''' Re-executes the last freestyle command using the saved prompt and model configuration.
    ''' Restores alternate model settings if the last command used alternate model.
    ''' Shows error message if no previous freestyle command is stored.
    ''' </summary>
    Public Async Sub FreeStyleRepeat()
        If INILoadFail() OrElse Not IsDocumentEditable() Then Return

        Dim LastFreestylePrompt As String = My.Settings.LastFreestylePrompt

        originalConfig = GetCurrentConfig(_context)
        originalConfigLoaded = False

        If String.IsNullOrWhiteSpace(LastFreestylePrompt) Then
            ShowCustomMessageBox("No last Freestyle command has been stored.")
            Return
        End If

        If My.Settings.LastFreestyleWasAM Then
            LastFreestyleModelConfig = My.Settings.LastFreestyleModelConfig

            If LastFreestyleModelConfig IsNot Nothing Then
                Dim ErrorFlag As Boolean = True
                ApplyModelConfig(_context, LastFreestyleModelConfig, ErrorFlag)
                If ErrorFlag Then
                    ShowCustomMessageBox("There was an error assigning the last model configuration. Aborting.")
                    Return
                End If
                originalConfigLoaded = True
            End If
        End If

        FreeStyle(My.Settings.LastFreestyleWasAM, My.Settings.LastFreestylePrompt)

    End Sub

    ''' <summary>
    ''' Core freestyle command processor. Handles prompt input, parses prefixes and triggers,
    ''' manages external content integration, and orchestrates LLM invocation with selected options.
    ''' </summary>
    ''' <param name="UseSecondAPI">True to use alternate/secondary model, False for primary model.</param>
    ''' <param name="LastPrompt">Optional pre-populated prompt (used by FreeStyleRepeat). Empty string prompts user for input.</param>
    ''' <remarks>
    ''' This method implements a complex state machine that:
    ''' 1. Initializes variables for all supported options and modes
    ''' 2. Builds instruction text based on available features and selection state
    ''' 3. Prompts user for input (unless LastPrompt provided)
    ''' 4. Processes special utility commands (encode, decode, version, etc.)
    ''' 5. Parses prompt prefixes to determine output mode
    ''' 6. Processes in-prompt triggers to enable features ({all}, {lib}, {net}, etc.)
    ''' 7. Handles external content embedding (files, documents, MyStyle, library/internet search)
    ''' 8. Constructs system and user prompts based on selected mode
    ''' 9. Invokes ProcessSelectedText with appropriate parameters
    ''' 10. Restores original configuration if alternate model was used
    ''' </remarks>
    Public Async Sub FreeStyle(UseSecondAPI As Boolean, Optional LastPrompt As String = "")
        If INILoadFail() OrElse Not IsDocumentEditable() Then Return
        Try
            ' Initialize prompt and system variables
            OtherPrompt = ""
            SysPrompt = ""
            InsertDocs = ""
            MyStyleInsert = ""

            CurrentDate = "(Current Date: " & DateTime.Now.ToString("dd-MMM-yyyy", CultureInfo.GetCultureInfo("en-US")) & ")"

            ' Initialize option flags for various processing modes
            Dim NoText As Boolean = False
            Dim DoMarkup As Boolean = False
            Dim DoClipboard As Boolean = False
            Dim DoBubbles As Boolean = False
            Dim DoInplace As Boolean = Override(INI_ReplaceText2, INI_ReplaceText2Override)
            Dim MarkupMethod As Integer = Override(INI_MarkupMethodWord, INI_MarkupMethodWordOverride)
            Dim DoLib As Boolean = False
            Dim DoNet As Boolean = False
            Dim DoTPMarkup As Boolean = False
            Dim TPMarkupName As String = ""
            Dim KeepFormatCap = INI_KeepFormatCap
            Dim DoKeepFormat As Boolean = INI_KeepFormat2
            Dim DoKeepParaFormat As Boolean = INI_KeepParaFormatInline
            Dim DoFileObject As Boolean = False
            Dim DoFileObjectClip As Boolean = False
            Dim DoPane As Boolean = False
            Dim DoNewDoc As Boolean = False
            Dim DoChunks As Boolean = False
            Dim ChunkSize As Integer = 1
            Dim NoFormatAndFieldSaving As Boolean = False
            Dim DoSlides As Boolean = False
            Dim DoChart As Integer = 0
            Dim DoMyStyle As Boolean = False
            Dim DoMultiModel As Boolean = True
            Dim DoBubblesExtract As Boolean = False
            Dim DoPushback As Boolean = False
            Dim DoFiles As Boolean = False
            Dim DoAssemble As Boolean = False
            Dim DoShowModel As Boolean = False
            Dim DoKB As Boolean = False

            ' Build instruction strings for user guidance
            Dim MarkupInstruct As String = $"start With '{MarkupPrefixAll}' for markups"
            Dim InplaceInstruct As String = $"with '{InPlacePrefix}'/'{AddPrefix} for replacing/adding to the selection"
            Dim BubblesInstruct As String = $"with '{BubblesPrefix}' for having your text commented"
            Dim PushbackInstruct As String = $"with '{PushbackPrefix}' for responding to comments only"
            Dim SlidesInstruct As String = $"with '{SlidesPrefix}' for adding to a Powerpoint file"
            Dim ChartInstruct As String = $"with '{ChartPrefix}'/'{ChartPrefix}' for creating a chart (normal or for webapps)"
            Dim ClipboardInstruct As String = $"with '{ClipboardPrefix}', '{NewdocPrefix}' or '{PanePrefix}' for separate output"
            Dim PromptLibInstruct As String = If(INI_PromptLib, " or press 'OK' for the prompt library", "")
            Dim ExtInstruct As String = $"; include '{ExtTrigger}' or '{ExtTriggerFixed}' (multiple times) for including the text of (a) file(s) (txt, docx, pdf), {ExtDirTrigger} for a directory of text files, {ExtUrlTrigger} for URL content, or '{AddDocTrigger}' for an open Word doc"
            Dim TPMarkupInstruct As String = $"; add '{TPMarkupTriggerInstruct}' if revisions [of user] should be pointed out to the LLM"
            Dim NoFormatInstruct As String = $"; add '{NoFormatTrigger2}'/'{KFTrigger2}'/'{KPFTrigger2}/{SameAsReplaceTrigger}' for overriding formatting defaults"
            Dim AllInstruct As String = $"; add '{AllTrigger}' to select all"
            Dim MyStyleInstruct As String = $"; add '{MyStyleTrigger}' to apply your personal style"
            Dim LibInstruct As String = $"; add '{LibTrigger}' for library search"
            Dim NetInstruct As String = $"; add '{NetTrigger}' for internet search"
            Dim KBInstruct As String = $"; add '{KnowledgeTriggerHelper.KbTrigger}' to search all stores or use '{KnowledgeTriggerHelper.KbTriggerPrefix}your query)' / '{KnowledgeTriggerHelper.KbTriggerPrefix}store:StoreName your query)' / '{KnowledgeTriggerHelper.KbTriggerPrefix}tag:TagName your query)' for Knowledge Store retrieval"
            Dim PureInstruct As String = $"; use '{PurePrefix}' for direct prompting"
            Dim FileInstruct As String = $"; use '{FilePrefix}' for modifying file(s)"
            Dim AssembleInstruct As String = $"; use '{AssemblePrefix}' for assembling a document from templates"
            Dim ChunkInstruct As String = $"; add '{ChunkTrigger}' for iterating through the text"
            Dim BubblesExtractInstruct As String = $"; add '{BubblesExtractTrigger}' for including bubble comments"
            Dim ObjectInstruct As String = $"; add '{ObjectTrigger}'/'{ObjectTrigger2}' for adding a file object"
            Dim MultiModelInstruct As String = $"; add '{MultiModelTrigger}' for multiple models, and {ShowModel} to include the model name in the output"
            Dim ToolSelectionInstruct As String = $"; add '{ToolSelectionTrigger}' to permit {ToolFriendlyName.ToLower} selection"
            Dim LastPromptInstruct As String = If(String.IsNullOrWhiteSpace(My.Settings.LastPrompt), "", "; Ctrl-P for your last prompt")
            Dim FormInstruct As String = $"; use '{FormPrefix}' for completing tables/form fields in a Word document"
            Dim FileObject As String = ""
            Dim SlideDeck As String = ""
            Dim DoForm As Boolean = False

            Dim DefaultPrefix As String = INI_DefaultPrefix
            Dim DefaultPrefixText As String = ""

            Dim application As Word.Application = Globals.ThisAddIn.Application
            Dim selection As Microsoft.Office.Interop.Word.Selection = application.Selection

            ' Check if no text is selected (insertion point only)
            If selection.Type = WdSelectionType.wdSelectionIP Then NoText = True

            ' Check if the selected model supports tooling (can call tools)f
            Dim modelSupportsTool As Boolean = False
            If UseSecondAPI Then
                ' Check via SharedMethods - based on APICall_ToolInstructions
                modelSupportsTool = Not String.IsNullOrWhiteSpace(INI_APICall_ToolInstructions_2) OrElse
                                    SharedMethods.ModelSupportsTooling(LastFreestyleModelConfig)
            End If

            Dim ToolTriggerAvailable As Boolean =
                Not UseSecondAPI AndAlso
                SharedMethods.HasToolingCapableSpecialTaskModel(_context, INI_AlternateModelPath, "ToolDefaultModel")

            Dim ToolTriggerInstruct As String =
                If(ToolTriggerAvailable, $"; add '{ToolTrigger}' to perform an agentic search of your selected {ToolFriendlyName.ToLower}", "")

            ' Build additional instruction text based on configuration and selection state
            Dim AddOnInstruct As String = AllInstruct

            If Not NoText Then
                AddOnInstruct += NoFormatInstruct.Replace("; add", ", ")
                AddOnInstruct += TPMarkupInstruct.Replace("; add", ", ")
                AddOnInstruct += ChunkInstruct.Replace("; add", ", ")
                AddOnInstruct += BubblesExtractInstruct.Replace("; add", ", ")
            End If
            If INI_Lib Then
                AddOnInstruct += LibInstruct.Replace("; add", ",")
            End If
            If INI_ISearch Then
                AddOnInstruct += NetInstruct.Replace("; add", ", ")
            End If
            If Not String.IsNullOrWhiteSpace(INI_KnowledgeStorePath) OrElse
               Not String.IsNullOrWhiteSpace(INI_KnowledgeStorePathLocal) Then
                AddOnInstruct += KBInstruct.Replace("; add", ", ")
            End If
            If Not String.IsNullOrWhiteSpace(INI_MyStylePath) Then
                AddOnInstruct += MyStyleInstruct.Replace("; add", ", ")
            End If
            If UseSecondAPI Then
                If Not String.IsNullOrWhiteSpace(INI_APICall_Object_2) Then
                    AddOnInstruct += ObjectInstruct.Replace("; add", ",")
                    DoFileObject = True
                End If
                If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                    AddOnInstruct += MultiModelInstruct.Replace("; add", ", ")
                End If

            Else
                If Not String.IsNullOrWhiteSpace(INI_APICall_Object) Then
                    AddOnInstruct += ObjectInstruct.Replace("; add", ",")
                    DoFileObject = True
                End If
            End If

            If Not String.IsNullOrWhiteSpace(ToolTriggerInstruct) Then
                AddOnInstruct += ToolTriggerInstruct.Replace("; add", ", ")
            End If

            If UseSecondAPI OrElse ToolTriggerAvailable Then
                If (Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) And modelSupportsTool) Or ToolTriggerAvailable Then
                    AddOnInstruct += ToolSelectionInstruct.Replace("; add", ", ")
                End If
            End If

            ' Format the instruction list with proper grammar (replace last comma with ", and")
            Dim lastCommaIndex As Integer = AddOnInstruct.LastIndexOf(","c)
            If lastCommaIndex <> -1 Then
                AddOnInstruct = AddOnInstruct.Substring(0, lastCommaIndex) & ", and" & AddOnInstruct.Substring(lastCommaIndex + 1)
            End If

            If DefaultPrefix.Trim() <> "" Then
                DefaultPrefixText = $" (default prefix: '{DefaultPrefix}')"
            End If

            Dim InsertButtons As System.Tuple(Of String, String, String)() = {
                            System.Tuple.Create("📄", "Include document {doc}", "{doc}"),
                            System.Tuple.Create("📁", "Include directory {dir}", "{dir}"),
                            System.Tuple.Create("📑", "Include other open document (adddoc)", "(adddoc)"),
                            System.Tuple.Create("📎", "Include file object (file)", "(file)")
                        }

            ' Prompt user for input if not provided via LastPrompt parameter
            If LastPrompt.Trim() = "" Then
                If Not NoText Then
                    ' Offer optional buttons for common prefix shortcuts when text is selected
                    Dim OptionalButtons As System.Tuple(Of String, String, String)() = {
                            System.Tuple.Create("OK, use window", $"Use this to automatically insert '{ClipboardPrefix}' as a prefix.", ClipboardPrefix),
                            System.Tuple.Create("OK, use pane", $"Use this to automatically insert '{PanePrefix}' as a prefix.", PanePrefix),
                            System.Tuple.Create("OK, do a markup", $"Use this to automatically insert '{MarkupPrefixDiff}' as a prefix.", MarkupPrefixDiff),
                            System.Tuple.Create("OK, use bubbles", $"Use this to automatically insert '{BubblesPrefix}' as a prefix.", BubblesPrefix)
                        }


                    OtherPrompt = SLib.ShowCustomInputBox($"Please provide the prompt you wish to execute on the selected text ({MarkupInstruct}, {ClipboardInstruct}, {InplaceInstruct}, {BubblesInstruct}, {PushbackInstruct}, {ChartInstruct} or {SlidesInstruct}){PromptLibInstruct}{ExtInstruct}{AddOnInstruct}{PureInstruct}{LastPromptInstruct}{DefaultPrefixText}:", $"{AN} Freestyle (using " & If(UseSecondAPI, INI_Model_2, INI_Model) & ")", False, "", My.Settings.LastPrompt, OptionalButtons, InsertButtons, Context:=_context).Trim()
                Else
                    ' Offer limited optional buttons when no text is selected
                    Dim OptionalButtons As System.Tuple(Of String, String, String)() = {
                            System.Tuple.Create("OK, use window", $"Use this to automatically insert '{ClipboardPrefix}' as a prefix.", ClipboardPrefix),
                            System.Tuple.Create("OK, use pane", $"Use this to automatically insert '{PanePrefix}' as a prefix.", PanePrefix)
                        }
                    OtherPrompt = SLib.ShowCustomInputBox($"Please provide the prompt you wish to execute ({ClipboardInstruct}, {ChartInstruct} or {SlidesInstruct}){PromptLibInstruct}{ExtInstruct}{AddOnInstruct}{PureInstruct}{FileInstruct}{AssembleInstruct}{FormInstruct}{LastPromptInstruct}{DefaultPrefixText}:", $"{AN} Freestyle (using " & If(UseSecondAPI, INI_Model_2, INI_Model) & ")", False, "", My.Settings.LastPrompt, OptionalButtons, InsertButtons, Context:=_context).Trim()
                End If
            Else
                OtherPrompt = LastPrompt
            End If

            'Debug.WriteLine($"OtherPrompt: '{OtherPrompt}'")

            SelectedText = ""

            ' === Special utility commands (can execute without text selection) ===

            ' --- Help picker: show all short commands and let user choose one ---
            If String.Equals(OtherPrompt.Trim(), "help", StringComparison.OrdinalIgnoreCase) Or String.Equals(OtherPrompt.Trim(), "?", StringComparison.OrdinalIgnoreCase) Then

                Dim items As New List(Of SLib.SelectionItem)()
                Dim idToCommand As New Dictionary(Of Integer, String)()

                Dim id As Integer = 1
                Dim AddItem =
                    Sub(cmd As String, desc As String)
                        items.Add(New SLib.SelectionItem($"{cmd}: {desc}", id))
                        idToCommand(id) = cmd
                        id += 1
                    End Sub

                ' GENERAL
                AddItem("domain", "Show the current domain and any configured domain restrictions.")
                AddItem("model", "Show the primary model, timeout, and (if set) max output token length.")
                AddItem("terms", "Insert configured usage restrictions/permissions into the document.")
                AddItem("version", "Show the installed Red Ink version.")
                AddItem("switch", "Temporarily swap primary and secondary models.")
                AddItem("clientname", "Copy and show this PC's client identifier (used for UpdateClients and CentralConfigClients).")
                AddItem("license", "Show license information and access license manage dialog.")
                AddItem("offlinelicense", "Create an offline license key using a private key.")

                ' CONFIG / MENU
                AddItem("settings", "Open the settings dialog.")
                AddItem("reload", "Reload the configuration from disk and rebuild menus.")
                AddItem("reset", "Reset local configuration to defaults and rebuild menus.")
                AddItem("cleanmenu", "Remove old context menus and rebuild them.")

                ' INI UPDATE / SIGNING
                AddItem("iniupdate", "Check for and apply configuration updates.")
                AddItem("iniupdateignored", "Open the ignore list for configuration updates.")
                AddItem("iniupdateignore", "Open the ignore list for configuration updates.")
                AddItem("iniload", "Import/apply configuration via the configuration import workflow.")
                AddItem("inirollback", "Roll back the last imported configuration change (creates a backup).")
                AddItem("iniupdatekeys", "Open signature management (sign/validate configs, manage keys).")
                AddItem("signtool", "Open signature management (sign/validate configs, manage keys).")
                AddItem("iniupdatebatch", "Open batch signing.")
                AddItem("signbatch", "Open batch signing.")

                ' PROMPTS / LOGS
                AddItem("clearlastprompt", "Clear the stored last Freestyle prompt and repeat state.")
                AddItem("promptlog", "Show/edit the cached Freestyle prompt log.")
                AddItem("logstat", $"Compile and show {AN} usage startistics based on collected logs.")
                AddItem("logfile", $"Open the {AN} local log file (for events re update, license, etc.).")

                ' MYSTYLE
                AddItem("definemystyle", "Create/update your MyStyle prompts.")
                AddItem("editmystyle", "Open the MyStyle prompt file in an editor.")

                ' SECURITY / REGISTRY (selection required)
                AddItem("encode", "Encode the selected text (e.g., API key) and copy it to the clipboard.")
                AddItem("decode", "Decode the selected encoded text and copy it to the clipboard.")
                AddItem("inipath", "Save the selected text as the INI path in the registry.")
                AddItem("codebasis", "Save the selected text as the CodeBasis value in the registry.")

                ' JSON / TEMPLATES (selection required)
                AddItem("generateresponsetemplate", "Generate a JSON response template from selected JSON + description.")
                AddItem("generateresponsekey", "Generate a JSON response key from selected JSON + description.")
                AddItem("mcp", "Import tools from an MCP server and generate INI sections.")

                ' CLIPBOARD / INSERTION
                AddItem("insertclipboard", "Insert clipboard content at the cursor position.")
                AddItem("insertclip", "Insert clipboard content at the cursor position.")

                ' KNOWLEDGE STORE
                AddItem("kbindex", "Index new/changed files across all active knowledge stores.")
                AddItem("kbreindex", "Force full re-index of all knowledge stores (regenerates all metadata, uses API credits).")
                AddItem("kbrefreshvectors", "Rebuild embeddings from existing wiki pages only (use after changing the embedding model).")
                AddItem("kbaddstore", "Add a new Knowledge Store (Name|Path).")
                AddItem("kbstore", "Show the list of Knowledge Stores and their status.")
                AddItem("kbschema", "Open the selected Knowledge Store schema in the internal JSON editor.")
                AddItem("kbhealth", "Run an AI health check/lint on the active Wiki (finds orphans/duplicates).")
                AddItem("kbrepair", "Run an AI repair operation on the active Wiki (fixes issues found during health check).")
                AddItem("cliptowiki", "Store the clipboard text in the knowledgebase wiki.")
                ' TOOLS / SOURCES
                AddItem("setagents", "Select sources/tools available For agentic models (session scope).")
                AddItem("loadurl", "Retrieve the text Of a particular URL given.")
                AddItem("translator", "Open a widget that provides you With an On-the-fly translation.")
                AddItem("drawio", "Open a draw.io For editing chart files, optionally With Internet blocking.")
                AddItem("drawioconverter", "Convert a draw.io flow chart To a HTML mini-web-app.")
                AddItem("pptxconvert", "Convert a PowerPoint presentation To a different template format.")
                AddItem("talktome", "Start a widget that allows you to control Red Ink via speech")

                ' PRIVACY / TRANSFORMS
                AddItem("anonymize", "Anonymize/redact the current selection (no LLM Call).")
                AddItem("convertmarkdown", "Convert Markdown In the selected text To Word formatting.")

                ' AUDIO / SPEECH
                AddItem("speech", "Start speech transcription (Transcriptor).")
                AddItem("read", "Create audio (TTS) from the selected text.")
                AddItem("readlocal", "Read the selected text Using local TTS (no cloud Call).")
                AddItem("voices", "Select a Single cloud TTS voice.")
                AddItem("voices2", "Select multiple cloud TTS voices.")
                AddItem("voiceslocal", "Select the local TTS voice.")
                AddItem("createpodcast", "Create a podcast from the selected text.")
                AddItem("readpodcast", "Play/read a podcast based On the current selection.")

                ' DOCUMENT / CLAUSES
                AddItem("doccheck", "Run the document check.")
                AddItem("learndocstyle", "Learn document style — extract a style template (trainer document Or automatic/AI).")
                AddItem("applydocstyle", "Apply a style template.")
                AddItem("findclause", "Search For a clause In the clause library/database.")
                AddItem("addclause", "Add a clause To the clause library/database.")
                AddItem("splitpdf", "Split a PDF into separate exhibits based On its content.")
                AddItem("stamper", "Apply an exhibit stamp To a PDF based On its filename.")

                ' WEB AGENT
                AddItem("webagentcreator", "Create/modify web agent scripts.")
                AddItem("webagent", "Run the web agent (requires configured script paths).")

                ' ANALYSIS
                AddItem("findhiddenprompts", "Scan the document For hidden prompts.")

                Dim chosen As Integer = SLib.SelectValue(items,
                            1,
                            "Select a Freestyle Short command (Esc To cancel):",
                            $"{AN} Freestyle - Help"
                        )

                If chosen <= 0 OrElse Not idToCommand.ContainsKey(chosen) Then
                    Return
                End If

                OtherPrompt = idToCommand(chosen)
                ' Continue normal execution: the short-command checks below will now match.
            End If

            If Not NoText Then

                SelectedText = selection.Text

                ' Store selected text as code basis in registry
                If String.Equals(OtherPrompt.Trim(), "codebasis", StringComparison.OrdinalIgnoreCase) Then
                    SLib.WriteToRegistry(RemoveCR(RegPath_CodeBasis), RemoveCR(selection.Text))
                    selection.Range.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                    Return
                End If

                ' Store selected text as INI path in registry
                If OtherPrompt.StartsWith("inipath", StringComparison.OrdinalIgnoreCase) Then
                    SLib.WriteToRegistry(RemoveCR(RegPath_IniPath), RemoveCR(selection.Text))
                    selection.Range.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                    Return
                End If

                ' Encode selected text (e.g., API key) and copy to clipboard
                If String.Equals(OtherPrompt.Trim(), "encode", StringComparison.OrdinalIgnoreCase) Then
                    Dim Key As String = CodeAPIKey(RemoveCR(selection.Text))
                    SLib.PutInClipboard(Key)
                    selection.Range.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                    selection.TypeText(vbCrLf & "Encoded key (also in clipboard):" & vbCrLf & Key)
                    selection.ParagraphFormat.Hyphenation = CInt(False)
                    SLib.PutInClipboard(Key)
                    Return
                End If

                ' Decode selected text and copy to clipboard
                If String.Equals(OtherPrompt.Trim(), "decode", StringComparison.OrdinalIgnoreCase) Then
                    Dim Key As String = DeCodeAPIKey(RemoveCR(selection.Text))
                    SLib.PutInClipboard(Key)
                    selection.Range.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                    selection.TypeText(vbCrLf & "Decoded key (also in clipboard):" & vbCrLf & Key)
                    selection.ParagraphFormat.Hyphenation = CInt(False)
                    Return
                End If

                ' Convert selected markdown text to formatted Word content
                If OtherPrompt.StartsWith("convertmarkdown", StringComparison.OrdinalIgnoreCase) Then
                    Dim trailingCR = (SelectedText.EndsWith(vbCrLf) Or SelectedText.EndsWith(vbLf) Or SelectedText.EndsWith(vbCr))
                    InsertTextWithMarkdown(selection, SelectedText, trailingCR, True)
                    Return
                End If

            End If

            ' Knowledge store indexing commands
            If OtherPrompt.Equals("cliptowiki", StringComparison.OrdinalIgnoreCase) Then
                Dim stores = KnowledgeStoreCatalog.LoadAll(_context)
                If stores.Count = 0 Then
                    ShowCustomMessageBox("No Knowledge Store is configured. Use 'kbaddstore Name|Path' first.")
                    Return
                End If

                Dim selectedStorePath As String = stores(0).ResolvedSourcePath
                If stores.Count > 1 Then
                    Dim items As New List(Of SLib.SelectionItem)()
                    Dim idx As Integer = 1
                    For Each s In stores
                        items.Add(New SLib.SelectionItem(s.Name & " (" & s.ResolvedSourcePath & ")", idx))
                        idx += 1
                    Next
                    Dim chosen As Integer = SLib.SelectValue(items, 1, "Select a Knowledge Store:", "Knowledge Store")
                    If chosen <= 0 Then Return
                    selectedStorePath = stores(chosen - 1).ResolvedSourcePath
                End If

                Await KnowledgeWikiService.CreatePageFromClipboardAsync(selectedStorePath, My.Computer.Clipboard.GetText(), _context)
                ShowCustomMessageBox("Clipboard saved to Wiki!")
                Return
            End If

            ' Run Knowledge Base health check / linter
            If OtherPrompt.Equals("kbhealth", StringComparison.OrdinalIgnoreCase) Then
                Dim stores = KnowledgeStoreCatalog.LoadAll(_context)
                If stores.Count = 0 Then
                    ShowCustomMessageBox("No Knowledge Store is configured. Use 'kbaddstore Name|Path' first.")
                    Return
                End If

                Dim selectedStorePath As String = stores(0).ResolvedSourcePath
                If stores.Count > 1 Then
                    Dim items As New List(Of SLib.SelectionItem)()
                    Dim idx As Integer = 1
                    For Each s In stores
                        items.Add(New SLib.SelectionItem(s.Name & " (" & s.ResolvedSourcePath & ")", idx))
                        idx += 1
                    Next
                    Dim chosen As Integer = SLib.SelectValue(items, 1, "Select Knowledge Store to lint:", "Knowledge Store")
                    If chosen <= 0 Then Return
                    selectedStorePath = stores(chosen - 1).ResolvedSourcePath
                End If

                Dim report = Await KnowledgeWikiService.LintWikiAsync(selectedStorePath, _context)
                SP_MergePrompt_Cached = SP_MergePrompt

                Dim response = ShowCustomWindow("Here is the health report of the Knowledge Store you have selected:", report, "You can copy it to the clipboard, if you wish.", "{AN} Knowledge Store")

                Return
            End If

            If OtherPrompt.Equals("kbschema", StringComparison.OrdinalIgnoreCase) Then
                Dim stores = KnowledgeStoreCatalog.LoadAll(_context)
                If stores.Count = 0 Then
                    ShowCustomMessageBox("No Knowledge Store is configured. Use 'kbaddstore Name|Path' first.")
                    Return
                End If

                Dim selectedStore As KnowledgeStoreCatalog.KnowledgeStoreDefinition = stores(0)

                If stores.Count > 1 Then
                    Dim items As New List(Of SLib.SelectionItem)()
                    Dim idx As Integer = 1
                    For Each s In stores
                        items.Add(New SLib.SelectionItem(s.Name & " (" & s.ResolvedSourcePath & ")", idx))
                        idx += 1
                    Next

                    Dim chosen As Integer = SLib.SelectValue(items, 1, "Select Knowledge Store schema to open:", "Knowledge Store")
                    If chosen <= 0 Then Return

                    selectedStore = stores(chosen - 1)
                End If

                If String.IsNullOrWhiteSpace(selectedStore.ResolvedSourcePath) Then
                    ShowCustomMessageBox("The selected Knowledge Store does not have a valid source path.", $"{AN} Knowledge Store")
                    Return
                End If

                KnowledgeStoreSchema.LoadOrCreate(selectedStore.ResolvedSourcePath)

                Dim schemaPath = KnowledgeStoreSchema.GetSchemaPath(selectedStore.ResolvedSourcePath)
                If String.IsNullOrWhiteSpace(schemaPath) Then
                    ShowCustomMessageBox("Could not resolve the schema path for the selected Knowledge Store.", $"{AN} Knowledge Store")
                    Return
                End If

                SharedMethods.ShowTextFileEditor(
                    schemaPath,
                    $"Edit schema for Knowledge Store '{selectedStore.Name}'.",
                    ForceJson:=True,
                    _context:=_context)

                Return
            End If

            ' Run Knowledge Base repair only
            If OtherPrompt.Equals("kbrepair", StringComparison.OrdinalIgnoreCase) Then
                Dim stores = KnowledgeStoreCatalog.LoadAll(_context)
                If stores.Count = 0 Then
                    ShowCustomMessageBox("No Knowledge Store is configured. Use 'kbaddstore Name|Path' first.")
                    Return
                End If

                Dim selectedStorePath As String = stores(0).ResolvedSourcePath
                If stores.Count > 1 Then
                    Dim items As New List(Of SLib.SelectionItem)()
                    Dim idx As Integer = 1
                    For Each s In stores
                        items.Add(New SLib.SelectionItem(s.Name & " (" & s.ResolvedSourcePath & ")", idx))
                        idx += 1
                    Next
                    Dim chosen As Integer = SLib.SelectValue(items, 1, "Select Knowledge Store to repair:", "Knowledge Store")
                    If chosen <= 0 Then Return
                    selectedStorePath = stores(chosen - 1).ResolvedSourcePath
                End If

                Dim summary = Await KnowledgeWikiService.ApplyWikiHealthFixesAsync(
                    kbRootPath:=selectedStorePath,
                    context:=_context,
                    includeLlmRepairs:=True)

                SP_MergePrompt_Cached = SP_MergePrompt

                Dim response = ShowCustomWindow(
                    "Here is the summary of the Knowledge Store repair:",
                    summary,
                    "You can copy it to the clipboard, if you wish.",
                    $"{AN} Knowledge Store")

                Return
            End If

            ' Knowledge store indexing commands
            If OtherPrompt.Equals("kbindex", StringComparison.OrdinalIgnoreCase) OrElse
               OtherPrompt.Equals("kbreindex", StringComparison.OrdinalIgnoreCase) Then

                If Not KnowledgeStoreCatalog.IsConfigured(_context) Then
                    ShowCustomMessageBox("No Knowledge Store catalog is configured. Set 'KnowledgeStorePath' or 'KnowledgeStorePathLocal' in your configuration.", $"{AN} Knowledge Store")
                    Return
                End If

                Dim forceReindex As Boolean = OtherPrompt.Equals("kbreindex", StringComparison.OrdinalIgnoreCase)
                Dim stores = KnowledgeStoreCatalog.GetActiveStores(_context)

                If stores.Count = 0 Then
                    ShowCustomMessageBox("No active Knowledge Stores found.", $"{AN} Knowledge Store")
                    Return
                End If

                Dim selectedStoreName As String = ""

                If stores.Count = 1 Then
                    selectedStoreName = stores(0).StoreId
                Else
                    Dim items As New List(Of SLib.SelectionItem)()
                    items.Add(New SLib.SelectionItem("All active Knowledge Stores", 1))

                    Dim idx As Integer = 2
                    For Each s In stores
                        items.Add(New SLib.SelectionItem(KnowledgeStoreCatalog.GetDisplayLabel(s), idx))
                        idx += 1
                    Next

                    Dim chosen As Integer = SLib.SelectValue(
                                        items,
                                        1,
                                        "Select which Knowledge Store(s) to index:",
                                        "Knowledge Store")

                    If chosen <= 0 Then Return

                    If chosen > 1 Then
                        selectedStoreName = stores(chosen - 2).StoreId
                    End If
                End If

                If forceReindex Then
                    Dim targetText As String = If(String.IsNullOrWhiteSpace(selectedStoreName),
                                                  "all active Knowledge Stores",
                                                  $"Knowledge Store '{selectedStoreName}'")

                    Dim answer As Integer = ShowCustomYesNoBox(
                        $"This will force a full re-index of {targetText}, regenerating all metadata and Wiki summaries. This may take a while and use API credits. Continue?",
                        "Yes, re-index", "No, cancel")

                    If answer <> 1 Then Return
                End If

                Await RunForegroundKnowledgeStoreIndexAsync(
                    storeName:=selectedStoreName,
                    forceReindex:=forceReindex)

                Return
            End If

            If OtherPrompt.Equals("kbrefreshvectors", StringComparison.OrdinalIgnoreCase) OrElse
               OtherPrompt.Equals("kbrefreshembeddings", StringComparison.OrdinalIgnoreCase) Then

                If Not KnowledgeStoreCatalog.IsConfigured(_context) Then
                    ShowCustomMessageBox("No Knowledge Store catalog is configured. Set 'KnowledgeStorePath' or 'KnowledgeStorePathLocal' in your configuration.", $"{AN} Knowledge Store")
                    Return
                End If

                Dim stores = KnowledgeStoreCatalog.GetActiveStores(_context)

                If stores.Count = 0 Then
                    ShowCustomMessageBox("No active Knowledge Stores found.", $"{AN} Knowledge Store")
                    Return
                End If

                Dim targetStores As New List(Of KnowledgeStoreCatalog.KnowledgeStoreDefinition)()

                If stores.Count = 1 Then
                    targetStores.Add(stores(0))
                Else
                    Dim items As New List(Of SLib.SelectionItem)()
                    items.Add(New SLib.SelectionItem("All active Knowledge Stores", 1))

                    Dim idx As Integer = 2
                    For Each s In stores
                        items.Add(New SLib.SelectionItem($"{s.Name} ({s.ResolvedSourcePath})", idx))
                        idx += 1
                    Next

                    Dim chosen As Integer = SLib.SelectValue(
                        items,
                        1,
                        "Select which Knowledge Store(s) to rebuild embeddings for:",
                        "Knowledge Store")

                    If chosen <= 0 Then Return

                    If chosen = 1 Then
                        targetStores.AddRange(stores)
                    Else
                        targetStores.Add(stores(chosen - 2))
                    End If
                End If

                Dim CountEmbeddableWikiPages As Func(Of String, Integer) =
                    Function(kbStoreRoot As String) As Integer
                        If String.IsNullOrWhiteSpace(kbStoreRoot) Then Return 0

                        Dim wikiRoot As String = Path.Combine(kbStoreRoot, ".redink", KnowledgeStoreCatalog.WikiFolder)
                        If Not Directory.Exists(wikiRoot) Then Return 0

                        Dim count As Integer = 0

                        For Each filePath In Directory.GetFiles(wikiRoot, "*.md", SearchOption.AllDirectories)
                            Dim name As String = Path.GetFileName(filePath)

                            If name.Equals(KnowledgeStoreCatalog.IndexFile, StringComparison.OrdinalIgnoreCase) Then Continue For
                            If name.Equals(KnowledgeStoreCatalog.LogFile, StringComparison.OrdinalIgnoreCase) Then Continue For
                            If name.Equals("health_report.md", StringComparison.OrdinalIgnoreCase) Then Continue For

                            count += 1
                        Next

                        Return count
                    End Function

                Dim totalPages As Integer = 0
                For Each store In targetStores
                    totalPages += CountEmbeddableWikiPages(store.ResolvedSourcePath)
                Next

                If totalPages = 0 Then
                    ShowCustomMessageBox("No wiki pages were found to rebuild embeddings for.", $"{AN} Knowledge Store")
                    Return
                End If

                Dim targetText As String = If(targetStores.Count = 1,
                                              $"Knowledge Store '{targetStores(0).Name}'",
                                              "all active Knowledge Stores")

                Dim answer As Integer = ShowCustomYesNoBox(
                    $"This will rebuild the embedding index for {targetText} from the existing wiki pages using the currently configured embedding model. Continue?",
                    "Yes, rebuild embeddings", "No, cancel")

                If answer <> 1 Then Return

                ProgressBarModule.GlobalProgressValue = 0
                ProgressBarModule.GlobalProgressMax = totalPages
                ProgressBarModule.GlobalProgressLabel = "Preparing embedding rebuild..."
                ProgressBarModule.CancelOperation = False
                ProgressBarModule.ShowProgressBarInSeparateThread(
                    $"{AN} Knowledge Store — Embeddings",
                    "Refreshing embeddings...")

                Dim rebuiltPages As Integer = 0
                Dim progressOffset As Integer = 0

                Try
                    For Each store In targetStores
                        If ProgressBarModule.CancelOperation Then Exit For

                        Dim storePageCount As Integer = CountEmbeddableWikiPages(store.ResolvedSourcePath)

                        rebuiltPages += Await KnowledgeEmbeddingService.RebuildAllWikiEmbeddingsAsync(
                            kbRootPath:=store.ResolvedSourcePath,
                            context:=_context,
                            progressPrefix:=store.Name,
                            progressOffset:=progressOffset,
                            progressTotal:=totalPages)

                        progressOffset += storePageCount
                    Next
                Finally
                    Dim wasCancelled As Boolean = ProgressBarModule.CancelOperation
                    ProgressBarModule.CancelOperation = True

                    If wasCancelled Then
                        ShowCustomMessageBox(
                            $"Embedding rebuild was cancelled. Refreshed {rebuiltPages} wiki page embedding set(s).",
                            $"{AN} Knowledge Store")
                    Else
                        ShowCustomMessageBox(
                            $"Embedding rebuild complete. Refreshed {rebuiltPages} wiki page embedding set(s) across {targetStores.Count} store(s).",
                            $"{AN} Knowledge Store")
                    End If
                End Try

                Return
            End If

            ' Add a new Knowledge Store via freestyle command
            If OtherPrompt.StartsWith("kbaddstore", StringComparison.OrdinalIgnoreCase) Then
                Dim parts = OtherPrompt.Substring("kbaddstore".Length).Trim().Split("|"c)
                If parts.Length < 2 OrElse String.IsNullOrWhiteSpace(parts(0)) OrElse String.IsNullOrWhiteSpace(parts(1)) Then
                    ShowCustomMessageBox(
                        "Usage: kbaddstore Name|Path" & vbCrLf &
                        "Example: kbaddstore Research|%UserProfile%\Documents\Research",
                        $"{AN} Knowledge Store")
                Else
                    Dim storeName As String = parts(0).Trim()
                    Dim rawPath As String = parts(1).Trim()

                    ' 1. Resolve environment variables (e.g. %UserProfile%)
                    Dim resolvedPath As String = SLib.ExpandEnvironmentVariables(rawPath)

                    Try
                        ' 2. Physically create the root directory if it doesn't exist
                        If Not System.IO.Directory.Exists(resolvedPath) Then
                            System.IO.Directory.CreateDirectory(resolvedPath)
                        End If

                        ' 3. Pre-initialize the Wiki and Raw subfolders so it's ready natively
                        KnowledgeWikiService.InitializeWikiStructure(resolvedPath)

                        ' 4. Save definition to the logical catalog
                        Dim def = KnowledgeStoreCatalog.CreateDefinition(storeName, rawPath, _context)
                        Dim allDefs = KnowledgeStoreCatalog.LoadAll(_context)
                        allDefs.Add(def)
                        KnowledgeStoreCatalog.SaveLocalCatalog(allDefs, _context)

                        ShowCustomMessageBox($"Knowledge Store '{def.Name}' successfully created physically at:{vbCrLf}{resolvedPath}", $"{AN} Knowledge Store")
                    Catch ex As Exception
                        ShowCustomMessageBox($"Could not create directory at '{resolvedPath}'. Error: {ex.Message}", $"{AN} Knowledge Store")
                    End Try
                End If
                Return
            End If

            If String.Equals(OtherPrompt.Trim(), "kbstore", StringComparison.OrdinalIgnoreCase) Then
                Using frm As New KnowledgeStoreForm(_context)
                    frm.ShowDialog()
                End Using
                Return
            End If

            If String.Equals(OtherPrompt.Trim(), "pptxconvert", StringComparison.OrdinalIgnoreCase) Then
                RetemplatePresentation_UI()
                Return
            End If

            If String.Equals(OtherPrompt.Trim(), "talktome", StringComparison.OrdinalIgnoreCase) Then
                Globals.ThisAddIn.ShowTalkToMeWidget()
                Return
            End If

            If String.Equals(OtherPrompt.Trim(), "offlinelicense", StringComparison.OrdinalIgnoreCase) Then
                SharedMethods.ShowOfflineDomainLicenseGeneratorDialog()
                Return
            End If


            ' Decode serial 
            If String.Equals(OtherPrompt.Trim(), "decodeserial", StringComparison.OrdinalIgnoreCase) Then
                DecodeSerial(selection)
                Return
            End If

            ' Create serial and copy to clipboard
            If String.Equals(OtherPrompt.Trim(), "encodeserial", StringComparison.OrdinalIgnoreCase) Then
                EncodeSerial(selection)
                Return
            End If

            ' Display domain configuration information
            If String.Equals(OtherPrompt.Trim(), "domain", StringComparison.OrdinalIgnoreCase) Then
                ShowCustomMessageBox($"{AN} is running in the domain '{GetDomain()}' and configured to run in {If(String.IsNullOrEmpty(SLib.allowedDomains), "any domain ('alloweddomains' has not been set).", "'" & SLib.allowedDomains & "'.")}", "")
                Return
            End If

            ' Display primary model configuration
            If String.Equals(OtherPrompt.Trim(), "model", StringComparison.OrdinalIgnoreCase) Then
                ShowCustomMessageBox("I am using the " & INI_Model & " model as my primary model with a default timeout of " & (INI_Timeout / 1000) & " seconds (" & Microsoft.VisualBasic.Strings.Format(INI_Timeout / 60000, "0.00") & " minutes)." & If(INI_MaxOutputToken > 0, "The maximum output token length is " & INI_MaxOutputToken & ".", ""))
                Return
            End If

            ' Display usage restrictions/permissions from configuration
            If String.Equals(OtherPrompt.Trim(), "terms", StringComparison.OrdinalIgnoreCase) Then
                selection.Range.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                selection.TypeText(vbCrLf & If(INI_UsageRestrictions = "", "No usage restrictions or permissions have been defined in the configuration file.", "The defined usage restrictions or permissions defined in the configuration file are: " & INI_UsageRestrictions) & vbCrLf)
                Return
            End If

            ' Anonymize selected text (redact sensitive information)
            If String.Equals(OtherPrompt.Trim(), "anonymize", StringComparison.OrdinalIgnoreCase) Then
                AnonymizeSelection()
                Return
            End If

            ' Insert clipboard content at current position
            If OtherPrompt.StartsWith("insertclipboard", StringComparison.OrdinalIgnoreCase) OrElse OtherPrompt.StartsWith("insertclip", StringComparison.OrdinalIgnoreCase) Then
                Call InsertClipboard()
                Return
            End If

            ' Generate response template/key from JSON payload and natural language description
            If OtherPrompt.StartsWith("generateresponsekey", StringComparison.OrdinalIgnoreCase) Or OtherPrompt.StartsWith("generateresponsetemplate", StringComparison.OrdinalIgnoreCase) Then

                If NoText Then
                    ShowCustomMessageBox("No text has been selected. Select the text containing both the JSON payload to interpret and what you want the output to look like (by referencing to the JSON fields and structure in natural text).")
                    Return
                End If

                Dim response As String = Await LLM(SP_GenerateResponseKey & vbCrLf & Code_JsonTemplateFormatter, vbCrLf & SelectedText, "", "", 0, UseSecondAPI)

                selection.Range.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                selection.InsertAfter(vbCrLf & vbCrLf & response)

                Return
            End If

            ' Open MyStyle prompt file in text editor
            If OtherPrompt.StartsWith("editmystyle", StringComparison.OrdinalIgnoreCase) Then
                SLib.ShowTextFileEditor(ExpandEnvironmentVariables(INI_MyStylePath), "Edit your MyStyle prompt file (use 'Define MyStyle' to create new prompts automatically):")
                Return
            End If

            ' Create or update MyStyle prompts
            If OtherPrompt.StartsWith("definemystyle", StringComparison.OrdinalIgnoreCase) Then
                DefineMyStyle()
                Return
            End If

            ' Show and edit prompt log
            If OtherPrompt.StartsWith("promptlog", StringComparison.OrdinalIgnoreCase) Then
                ShowAndEditPromptLog()
                Return
            End If

            ' Create or modify web agent script
            If OtherPrompt.StartsWith("webagentcreator", StringComparison.OrdinalIgnoreCase) Then
                CreateModifyWebAgentScript()
                Return
            End If

            ' Execute web agent
            If String.Equals(OtherPrompt.Trim(), "webagent", StringComparison.OrdinalIgnoreCase) Then
                WebAgent()
                Return
            End If

            ' Find hidden prompts in document
            If String.Equals(OtherPrompt.Trim(), "findhiddenprompts", StringComparison.OrdinalIgnoreCase) Then
                FindHiddenPrompts()
                Return
            End If

            ' Load the content of an URL                                 
            If String.Equals(OtherPrompt.Trim(), "loadurl", StringComparison.OrdinalIgnoreCase) Then
                Dim url As String = SLib.ShowCustomInputBox("Please enter the URL to retrieve the content from:", $"{AN} - Load URL Content", True, "https://")
                If String.IsNullOrWhiteSpace(url) OrElse url.ToLower() = "esc" Then Return

                InfoBox.ShowInfoBox($"Retrieving content from {url.Trim()} ...")

                ' Run WebView2 on background and wait synchronously to stay on UI thread
                Dim content As String = ""
                Dim webViewTask As Task(Of String) = RetrieveWebsiteContent_WebView2(url.Trim(), 0)

                ' Wait for completion while pumping messages to keep UI responsive
                While Not webViewTask.IsCompleted
                    System.Windows.Forms.Application.DoEvents()
                    System.Threading.Thread.Sleep(50)
                End While

                ' Get the result (we're still on the original UI thread)
                If webViewTask.Status = TaskStatus.RanToCompletion Then
                    content = webViewTask.Result
                End If

                InfoBox.ShowInfoBox("")

                If String.IsNullOrWhiteSpace(content) Then
                    ShowCustomMessageBox("Could not retrieve content from the specified URL.")
                    Return
                End If

                Debug.WriteLine($"[loadurl] Inserting {content.Length} characters into document")

                ' We're on the UI thread - safe to access Word objects
                selection.Range.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                selection.InsertAfter(vbCrLf & vbCrLf & content)

                Return
            End If

            ' Test functionality using redinktest.txt from desktop
            If OtherPrompt.StartsWith("redinktest", StringComparison.OrdinalIgnoreCase) Then

                Dim desktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                Dim filePath As String = System.IO.Path.Combine(desktopPath, "redinktest.txt")
                If File.Exists(filePath) Then
                    Dim testtextorig As String = File.ReadAllText(filePath).Replace("\n", vbCrLf)
                    Dim testtext As String = SLib.ShowCustomWindow("Testfile content:", testtextorig, "", AN, False, True, True, True)
                    If testtext <> "" And testtext <> "Pane" Then
                        If testtext = "Markdown" Then
                            Globals.ThisAddIn.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                            Globals.ThisAddIn.Application.Selection.TypeParagraph()
                            Globals.ThisAddIn.Application.Selection.TypeParagraph()
                            InsertTextWithMarkdown(Globals.ThisAddIn.Application.Selection, vbCrLf & testtextorig, False)
                            Dim patternx As String = "\{\{(WFLD|WENT|WFNT):.*?\}\}"
                            If Regex.IsMatch(testtextorig, patternx) Then
                                Dim rng As Range = wordApp.Selection.Range
                                RestoreSpecialTextElements(rng)
                                rng.Document.Fields.Update()
                            End If
                        Else
                            SLib.PutInClipboard(testtext)
                        End If
                    ElseIf testtext = "Pane" Then
                        SP_MergePrompt_Cached = SP_MergePrompt
                        ShowPaneAsync(
                                                                        "Test Pane",
                                                                        testtextorig,
                                                                        "",
                                                                        AN,
                                                                        noRTF:=False,
                                                                        insertMarkdown:=True
                                                                        )
                    End If
                    Return
                Else
                    Return
                End If
            End If

            ' Switch primary and secondary models temporarily
            If String.Equals(OtherPrompt.Trim(), "switch", StringComparison.OrdinalIgnoreCase) Then
                selection.Range.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                If INI_SecondAPI Then
                    SwitchModels(_context)
                    ShowCustomMessageBox("You have temporarily switched the two configured models. Primary is now '" & INI_Model & "', and secondary is '" & INI_Model_2 & "'.")
                Else
                    ShowCustomMessageBox("You have defined only one model ('" & INI_Model & "').")
                End If
                Return
            End If

            ' Display version and license information
            If String.Equals(OtherPrompt.Trim(), "version", StringComparison.OrdinalIgnoreCase) Then
                ShowCustomMessageBox("You are using " & Version & $" of {AN}.", AN)
                Return
            End If

            ' Get computername (e.g., for UpdateClients parameter)
            If String.Equals(OtherPrompt.Trim(), "clientname", StringComparison.OrdinalIgnoreCase) Then
                SLib.PutInClipboard(GetCurrentClientIdentifier())
                ShowCustomMessageBox("Your hostname / client name is '" & GetCurrentClientIdentifier() & "' (also in the clipboard).", AN)
                Return
            End If


            ' Signature Management for Update INI Key Functionality
            If String.Equals(OtherPrompt.Trim(), "iniupdatekeys", StringComparison.OrdinalIgnoreCase) OrElse String.Equals(OtherPrompt.Trim(), "signtool", StringComparison.OrdinalIgnoreCase) Then
                ShowSignatureManagementDialog()
                Return
            End If


            If String.Equals(OtherPrompt.Trim(), "iniupdatebatch", StringComparison.OrdinalIgnoreCase) OrElse String.Equals(OtherPrompt.Trim(), "signbatch", StringComparison.OrdinalIgnoreCase) Then
                ShowBatchSigningDialog()
                Return
            End If

            If String.Equals(OtherPrompt.Trim(), "iniupdateignored", StringComparison.OrdinalIgnoreCase) OrElse String.Equals(OtherPrompt.Trim(), "iniupdateignore", StringComparison.OrdinalIgnoreCase) Then
                ShowIgnoredParametersDialog()
                Return
            End If

            If String.Equals(OtherPrompt.Trim(), "translator", StringComparison.OrdinalIgnoreCase) Then
                ShowQuickTranslate()
                Return
            End If

            If String.Equals(OtherPrompt.Trim(), "drawioconverter", StringComparison.OrdinalIgnoreCase) Or String.Equals(OtherPrompt.Trim(), "chart", StringComparison.OrdinalIgnoreCase) Then
                ConvertDrawioToHtml()
                Return
            End If

            If String.Equals(OtherPrompt.Trim(), "drawio", StringComparison.OrdinalIgnoreCase) Or String.Equals(OtherPrompt.Trim(), "chart", StringComparison.OrdinalIgnoreCase) Then
                OpenExistingDrawioFileForEditing()
                Return
            End If


            ' Signature Management for importing INI keys            
            If String.Equals(OtherPrompt.Trim(), "iniload", StringComparison.OrdinalIgnoreCase) Or String.Equals(OtherPrompt.Trim(), "iniupdateignore", StringComparison.OrdinalIgnoreCase) Then

                If IniImportManager.RunImportFromVariableConfigurationWindow(_context, Nothing) Then
                    Dim answer = ShowCustomYesNoBox("Your main configuration settings have changed. You need to reload them for them to become active. Proceed?", "Yes, reload", "No, load later")
                    If answer = 1 Then
                        ' Mark config as not loaded so InitializeConfig will re-read from disk
                        _context.INIloaded = False
                        ' Reload configuration from disk into memory
                        InitializeConfig(False, True)
                        ' Refresh the UI with the newly loaded values
                        _context.MenusAdded = False
                    End If
                End If
                Return
            End If


            ' Signature Management for importing INI keys
            If String.Equals(OtherPrompt.Trim(), "inirollback", StringComparison.OrdinalIgnoreCase) Or String.Equals(OtherPrompt.Trim(), "iniupdateignore", StringComparison.OrdinalIgnoreCase) Then

                If ShowCustomYesNoBox($"Do you really want to roll back you last configuration file change? A new backup will be created", "Yes, rollback", "No") = 1 Then
                    If IniImportManager.TryRollbackLastBackup(_context, Nothing) Then
                        Dim answer = ShowCustomYesNoBox("Your main configuration settings have changed. You need to reload them for them to become active. Proceed?", "Yes, reload", "No, load later")
                        If answer = 1 Then
                            ' Mark config as not loaded so InitializeConfig will re-read from disk
                            _context.INIloaded = False
                            ' Reload configuration from disk into memory
                            InitializeConfig(False, True)
                            ' Refresh the UI with the newly loaded values
                            _context.MenusAdded = False
                        End If
                    End If
                    Return
                Else
                    Return
                End If
            End If


            ' Check for INI updates and apply if available
            If String.Equals(OtherPrompt.Trim(), "iniupdate", StringComparison.OrdinalIgnoreCase) Then
                Dim answer As Boolean = CheckForIniUpdates(_context)
                If answer Then
                    ShowCustomMessageBox("Updates to the .ini file(s) have been applied.")
                Else
                    ShowCustomMessageBox("No updates were applied. Either no updates were found or you chose not to apply them.")
                End If
                Return
            End If


            ' Reset local configuration to defaults (with confirmation)
            If String.Equals(OtherPrompt.Trim(), "reset", StringComparison.OrdinalIgnoreCase) Then
                If ShowCustomYesNoBox($"Do you really want to reset your local configuration file and settings (if any) by removing non-mandatory entries? The current configuration file '{AN2}.ini' will NOT be saved to a '.bak' file. If you only want to reload the configuration settings for giving up any temporary changes, use 'reload' instead.", "Yes", "No") = 1 Then
                    INIloaded = False
                    ResetLocalAppConfig(_context)
                    MenusAdded = False
                    AddContextMenu()
                    ShowCustomMessageBox($"Following the reset, the configuration file '{AN2}.ini' has been be reloaded.")
                End If
                Return
            End If

#If DEBUG Then
            If String.Equals(OtherPrompt.Trim(), "lt", StringComparison.OrdinalIgnoreCase) Then
                SharedMethods.TestLicenseSystem()
                Return
            End If
#End If

            If String.Equals(OtherPrompt.Trim(), "license", StringComparison.OrdinalIgnoreCase) Then
                SharedMethods.ShowLicenseManagementDialog()
                Return
            End If

            ' Open the centralized usage Log File
            If String.Equals(OtherPrompt.Trim(), "logstat", StringComparison.OrdinalIgnoreCase) Then
                SharedLogger.AnalyzeLogs(_context)
                Return
            End If

            If String.Equals(OtherPrompt.Trim(), "tablefill", StringComparison.OrdinalIgnoreCase) Then
                directCompleteWordDocumentTables()
                Return
            End If

            If String.Equals(OtherPrompt.Trim(), "mcp", StringComparison.OrdinalIgnoreCase) Then
                ImportMCPServer()
                Return
            End If

            If String.Equals(OtherPrompt.Trim(), "splitpdf", StringComparison.OrdinalIgnoreCase) Then
                Globals.ThisAddIn.SplitPdfByExhibits()
                Return
            End If

            If String.Equals(OtherPrompt.Trim(), "stamper", StringComparison.OrdinalIgnoreCase) Then
                Globals.ThisAddIn.StampExhibitPDF()
                Return
            End If

            ' Open the Red Ink Local Log File
            If String.Equals(OtherPrompt.Trim(), "logfile", StringComparison.OrdinalIgnoreCase) Then
                Dim logPath As String = ""
                Try
                    logPath = System.IO.Path.Combine(
                                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                                AN2,
                                LogFileName)

                    If System.IO.File.Exists(logPath) Then
                        SLib.ShowTextFileEditor(logPath, $"{AN} Local Log File '{logPath}':", True, _context)
                    Else
                        ShowCustomMessageBox($"Local logfile at '{logPath}' not found.")
                    End If
                Catch ex As Exception
                    ShowCustomMessageBox($"Error opening local logfile at '{logPath}': " & ex.Message)
                End Try
                Return
            End If

            ' Start speech transcription
            If String.Equals(OtherPrompt.Trim(), "speech", StringComparison.OrdinalIgnoreCase) Then
                Transcriptor()
                Return

            End If

            ' Read selected text using local text-to-speech
            If OtherPrompt.StartsWith("readlocal", StringComparison.OrdinalIgnoreCase) Then
                SpeakSelectedText()
                Return

            End If

            ' Clear last saved freestyle prompt
            If OtherPrompt.StartsWith("clearlastprompt", StringComparison.OrdinalIgnoreCase) Then
                My.Settings.LastPrompt = ""
                My.Settings.LastFreestylePrompt = ""
                My.Settings.LastFreestyleModelConfig = Nothing
                My.Settings.LastFreestyleWasAM = False
                My.Settings.Save()
                Dim resultx = Globals.Ribbons.Ribbon1.InitializeAppAsync()
                ShowCustomMessageBox($"The last Freestyle prompt has been cleared.")

                Return

            End If

            ' Select local text-to-speech voice by number
            If OtherPrompt.StartsWith("voiceslocal", StringComparison.OrdinalIgnoreCase) Then
                SelectVoiceByNumber()
                Return
            End If

            ' Select cloud text-to-speech voices (multi-voice mode)
            If OtherPrompt.StartsWith("voices2", StringComparison.OrdinalIgnoreCase) Then
                Using frm As New TTSSelectionForm("Select the voices you wish to use.", $"{AN} Text-to-Speech - Select Voices", True)
                    If frm.ShowDialog() = DialogResult.OK Then
                        Dim selectedVoices As List(Of String) = frm.SelectedVoices
                        Dim outputPath As String = frm.SelectedOutputPath
                        If selectedVoices.Count > 0 Then
                            MessageBox.Show("Selected Voice(s): " & String.Join(", ", selectedVoices))
                        Else
                            MessageBox.Show("No voices selected.")
                        End If

                        If outputPath = "" Then
                            MessageBox.Show("Temporary output selected.")
                        Else
                            MessageBox.Show("Output path: " & outputPath)
                        End If
                    Else
                        MessageBox.Show("Voice selection was cancelled.")
                    End If
                End Using

                Return
            End If

            ' Select cloud text-to-speech voices (single-voice mode)
            If String.Equals(OtherPrompt.Trim(), "voices", StringComparison.OrdinalIgnoreCase) Then
                Using frm As New TTSSelectionForm("Select the voices you wish to use.", $"{AN} Text-to-Speech - Select Voices", False)
                    If frm.ShowDialog() = DialogResult.OK Then
                        Dim selectedVoices As List(Of String) = frm.SelectedVoices
                        Dim outputPath As String = frm.SelectedOutputPath
                        If selectedVoices.Count > 0 Then
                            MessageBox.Show("Selected Voice(s): " & String.Join(", ", selectedVoices))
                        Else
                            MessageBox.Show("No voices selected.")
                        End If

                        If outputPath = "" Then
                            MessageBox.Show("Temporary output selected.")
                        Else
                            MessageBox.Show("Output path: " & outputPath)
                        End If
                    Else
                        MessageBox.Show("Voice selection was cancelled.")
                    End If
                End Using

                Return
            End If

            ' Run document check
            If OtherPrompt.StartsWith("doccheck", StringComparison.OrdinalIgnoreCase) Then
                RunDocCheck()
                Return
            End If

            ' Run learn doc style
            If OtherPrompt.StartsWith("learndocstyle", StringComparison.OrdinalIgnoreCase) Then
                ExtractParagraphStylesToJson()
                Return
            End If

            ' Run apply doc style
            If OtherPrompt.StartsWith("applydocstyle", StringComparison.OrdinalIgnoreCase) Then
                ApplyStyleTemplate()
                Return
            End If


            ' Find clause in library/database
            If OtherPrompt.StartsWith("findclause", StringComparison.OrdinalIgnoreCase) Then
                FindClause()
                Return
            End If

            ' Add clause from library/database
            If OtherPrompt.StartsWith("addclause", StringComparison.OrdinalIgnoreCase) Then
                AddClause()
                Return
            End If

            ' Create podcast from selected text
            If OtherPrompt.StartsWith("createpodcast", StringComparison.OrdinalIgnoreCase) Then
                CreatePodcast()
                Return
            End If

            ' Read/play existing podcast
            If OtherPrompt.StartsWith("readpodcast", StringComparison.OrdinalIgnoreCase) Then
                ReadPodcast(selection.Text)
                Return
            End If

            ' Create audio from selected text
            If String.Equals(OtherPrompt.Trim(), "read", StringComparison.OrdinalIgnoreCase) Then
                CreateAudio()
                Return
            End If

            ' Clean and rebuild context menu
            If OtherPrompt.StartsWith("cleanmenu", StringComparison.OrdinalIgnoreCase) Then
                RemoveOldContextMenu()
                RemoveVeryOldContextMenu()
                MenusAdded = False
                AddContextMenu()
                Return
            End If

            ' Reload configuration from file
            If String.Equals(OtherPrompt.Trim(), "reload", StringComparison.OrdinalIgnoreCase) Then
                INIloaded = False
                InitializeConfig(False, True)
                MenusAdded = False
                AddContextMenu()
                ShowCustomMessageBox($"The configuration file '{AN2}.ini' has been be reloaded.")
                Return
            End If

            ' Show settings dialog
            If String.Equals(OtherPrompt.Trim(), "settings", StringComparison.OrdinalIgnoreCase) Then
                ShowSettings()
                Return
            End If

            ' Select tools
            If String.Equals(OtherPrompt.Trim(), "set" & ToolFriendlyName.ToLower, StringComparison.OrdinalIgnoreCase) Then
                Dim selectedToolsForSessionTemp As List(Of ModelConfig) = Nothing
                selectedToolsForSessionTemp = SelectToolsForSession(True)
                Return
            End If

            ' === Prompt library integration ===

            ' Show prompt library selector if prompt is empty and library is enabled
            If String.IsNullOrEmpty(OtherPrompt) And OtherPrompt <> "ESC" And INI_PromptLib Then

                Dim promptlibresult As (String, Boolean, Boolean, Boolean)

                promptlibresult = ShowPromptSelector(INI_PromptLibPath, INI_PromptLibPathLocal, Not NoText, Not NoText)

                OtherPrompt = promptlibresult.Item1
                DoMarkup = promptlibresult.Item2
                DoBubbles = promptlibresult.Item3
                DoClipboard = promptlibresult.Item4

                If OtherPrompt = "" Then
                    Return
                End If
            Else
                If String.IsNullOrEmpty(OtherPrompt) Or OtherPrompt = "ESC" Then Return
            End If

            ' === Default prefix handling ===

            ' Add default prefix if prompt doesn't start with a recognized prefix (word ending with colon)
            If Not String.IsNullOrWhiteSpace(OtherPrompt) Then
                Dim firstWord As String = OtherPrompt.Split({" "c}, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault()
                If firstWord IsNot Nothing AndAlso Not firstWord.EndsWith(":"c) Then

                    Dim prefix As String = DefaultPrefix.Trim()

                    ' Ensure prefix ends with colon
                    If prefix <> "" AndAlso Not prefix.EndsWith(":"c) Then
                        prefix &= ":"
                    End If

                    OtherPrompt = prefix & " " & OtherPrompt.Trim()
                    OtherPrompt = OtherPrompt.Trim()
                End If
            End If

            ' Save prompt to settings for potential repeat/recall
            My.Settings.LastPrompt = OtherPrompt
            My.Settings.Save()

            ' Process parameter placeholders in prompt (e.g., {param:PromptText})
            If Not SharedMethods.ProcessParameterPlaceholders(OtherPrompt) Then
                ShowCustomMessageBox("Freestyle canceled.", $"{AN} Freestyle")
                Exit Sub
            End If

            ' === In-prompt trigger processing ===

            Dim toolTriggerDetected As Boolean = False
            Dim toolTriggerConfig As ModelConfig = Nothing

            If UseSecondAPI AndAlso OtherPrompt.IndexOf(ToolTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                Dim freestyleToolDefaultModel As ModelConfig = Nothing

                If modelSupportsTool Then
                    ShowCustomMessageBox(
                        $"The {ToolTrigger} trigger is not available in Freestyle (2nd). The currently selected model already supports agentic processing. Please remove {ToolTrigger} and, if needed, use '{ToolSelectionTrigger}' to choose the available {ToolFriendlyName.ToLower}.")
                ElseIf SharedMethods.TryGetSpecialTaskModelConfig(
                    _context,
                    INI_AlternateModelPath,
                    "ToolDefaultModel",
                    freestyleToolDefaultModel) Then

                    ShowCustomMessageBox(
                        $"The {ToolTrigger} trigger is not available in Freestyle (2nd). Please either choose a model that directly supports agentic processing, or use Freestyle with {ToolTrigger} to invoke the predefined agentic model.")
                Else
                    ShowCustomMessageBox(
                        $"The {ToolTrigger} trigger is not available in Freestyle (2nd), and no predefined agentic model has been configured. Please choose a model that directly supports agentic processing.")
                End If

                Return
            End If

            If Not UseSecondAPI AndAlso OtherPrompt.IndexOf(ToolTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                toolTriggerDetected = True

                If Not SharedMethods.TryGetSpecialTaskModelConfig(
                    _context,
                    INI_AlternateModelPath,
                    "ToolDefaultModel",
                    toolTriggerConfig) Then

                    ShowCustomMessageBox(
                        $"The {ToolTrigger} trigger was requested, but no model with 'ToolDefaultModel=True' was found in the alternate model configuration. Please add a ToolDefaultModel entry to your configuration file.")
                    Return
                End If

                If Not SharedMethods.ModelSupportsTooling(toolTriggerConfig) Then
                    ShowCustomMessageBox(
                        $"The {ToolTrigger} trigger found a ToolDefaultModel, but it does not support {ToolFriendlyName.ToLower}. Please check the model's APICall_ToolInstructions setting.")
                    Return
                End If

                OtherPrompt = OtherPrompt.Replace(ToolTrigger, "").Trim()

                If Not originalConfigLoaded Then
                    originalConfig = GetCurrentConfig(_context)
                    originalConfigLoaded = True
                End If

                Dim ErrorFlag As Boolean = False
                ApplyModelConfig(_context, toolTriggerConfig, ErrorFlag)
                If ErrorFlag Then
                    ShowCustomMessageBox("There was an error assigning the tooling model configuration. Aborting.")
                    Return
                End If

                UseSecondAPI = True
                modelSupportsTool = True
            End If

            ' (all) trigger: Select entire document
            If OtherPrompt.IndexOf(AllTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(AllTrigger, "").Trim()
                Dim document As Word.Document = application.ActiveDocument
                document.Content.Select()
                NoText = False
            End If

            ' (lib) trigger: Enable library search
            If OtherPrompt.IndexOf(LibTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(LibTrigger, "").Trim()
                DoLib = True
            End If

            ' (kb) or (kb:...) trigger: Query knowledge store and inject context
            If KnowledgeTriggerHelper.HasKnowledgeTrigger(OtherPrompt) Then
                DoKB = True
            End If

            ' Track point markup trigger: Enable revision tracking in markup
            If OtherPrompt.IndexOf(TPMarkupTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(TPMarkupTrigger, "").Trim()
                DoTPMarkup = True
            End If

            ' (interate) trigger: Enable chunked processing (iterate through paragraphs)
            If OtherPrompt.IndexOf(ChunkTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(ChunkTrigger, "").Trim()
                DoChunks = True
            End If

            ' Bubble extract trigger: Include existing bubble comments in prompt
            If OtherPrompt.IndexOf(BubblesExtractTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(BubblesExtractTrigger, "").Trim()
                DoBubblesExtract = True
                If DoChunks Then
                    ShowCustomMessageBox($"The '{BubblesExtractTrigger}' option cannot be used together with '{ChunkTrigger}' - the bubble comments will not be extracted.")
                    DoBubblesExtract = False
                End If
            End If

            ' (model) trigger: Include model name in the output
            If OtherPrompt.IndexOf(ShowModel, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(ShowModel, "").Trim()
                DoShowModel = True
            End If

            ' === Formatting override triggers ===

            ' No format triggers: Disable formatting preservation
            If OtherPrompt.IndexOf(NoFormatTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(NoFormatTrigger, "").Trim()
                KeepFormatCap = 1
            End If
            If OtherPrompt.IndexOf(NoFormatTrigger2, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(NoFormatTrigger2, "").Trim()
                KeepFormatCap = 1
            End If

            ' Keep format triggers: Enable character-level formatting preservation
            If OtherPrompt.IndexOf(KFTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(KFTrigger, "").Trim()
                DoKeepFormat = True
            End If
            If OtherPrompt.IndexOf(KFTrigger2, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(KFTrigger2, "").Trim()
                DoKeepFormat = True
            End If

            ' Keep paragraph format triggers: Enable paragraph-level formatting preservation
            If OtherPrompt.IndexOf(KPFTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(KPFTrigger, "").Trim()
                DoKeepParaFormat = True
            End If
            If OtherPrompt.IndexOf(KPFTrigger2, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(KPFTrigger2, "").Trim()
                DoKeepParaFormat = True
            End If

            ' Same as replace trigger: Use replace-mode formatting behavior for add mode
            If Not DoInplace Then
                If OtherPrompt.IndexOf(SameAsReplaceTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                    OtherPrompt = OtherPrompt.Replace(SameAsReplaceTrigger, "").Trim()
                Else
                    NoFormatAndFieldSaving = True
                End If
            End If

            ' === File object triggers ===

            ' (object) trigger: Attach file object to LLM request
            If DoFileObject AndAlso OtherPrompt.IndexOf(ObjectTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(ObjectTrigger, "(a file object follows)").Trim()
            ElseIf DoFileObject AndAlso OtherPrompt.IndexOf(ObjectTrigger2, StringComparison.OrdinalIgnoreCase) >= 0 Then
                ' {objectclip} trigger: Use clipboard content as file object
                OtherPrompt = OtherPrompt.Replace(ObjectTrigger2, "(a clipboard object follows)").Trim()
                DoFileObjectClip = True
            Else
                DoFileObject = False
            End If

            ' === Track point markup with specific user name ===

            ' (markup:username) pattern: Extract username for targeted revision tracking
            Dim pattern As String = Regex.Escape(TPMarkupTriggerL) & "(.*?)" & Regex.Escape(TPMarkupTriggerR)
            Dim match As Match = Regex.Match(OtherPrompt, pattern, RegexOptions.IgnoreCase)
            If match.Success Then
                TPMarkupName = match.Groups(1).Value
                DoTPMarkup = True
                OtherPrompt = Regex.Replace(OtherPrompt, pattern, String.Empty, RegexOptions.IgnoreCase)
            End If

            ' === Prefix-based mode selection ===

            Dim ShowPrefixRequiresSelection As Action(Of String) =
    Sub(prefix As String)
        ShowCustomMessageBox($"The '{prefix}' prefix is not available when no text is selected.")
    End Sub

            ' Process prompt prefix to determine output mode and remove prefix from prompt
            If OtherPrompt.StartsWith(ClipboardPrefix, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(ClipboardPrefix.Length).Trim()
                DoClipboard = True
                DoChunks = False
            ElseIf OtherPrompt.StartsWith(ClipboardPrefix2, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(ClipboardPrefix2.Length).Trim()
                DoClipboard = True
                DoChunks = False
            ElseIf OtherPrompt.StartsWith(NewdocPrefix, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(NewdocPrefix.Length).Trim()
                DoClipboard = True
                DoChunks = False
                DoNewDoc = True
            ElseIf OtherPrompt.StartsWith(BubblesPrefix, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(BubblesPrefix.Length).Trim()
                DoBubbles = True
            ElseIf OtherPrompt.StartsWith(SlidesPrefix, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(SlidesPrefix.Length).Trim()
                DoSlides = True
                DoClipboard = True
                DoChunks = False
            ElseIf OtherPrompt.StartsWith(ChartPrefix, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(ChartPrefix.Length).Trim()
                DoChart = 1
                DoClipboard = True
                DoChunks = False
            ElseIf OtherPrompt.StartsWith(ChartPrefixApp, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(ChartPrefixApp.Length).Trim()
                DoChart = 2
                DoClipboard = True
                DoChunks = False
            ElseIf OtherPrompt.StartsWith(InPlacePrefix, StringComparison.OrdinalIgnoreCase) Then
                If NoText Then
                    ShowPrefixRequiresSelection(InPlacePrefix)
                    Return
                End If
                OtherPrompt = OtherPrompt.Substring(InPlacePrefix.Length).Trim()
                DoInplace = True
            ElseIf OtherPrompt.StartsWith(AddPrefix, StringComparison.OrdinalIgnoreCase) Then
                If NoText Then
                    ShowPrefixRequiresSelection(AddPrefix)
                    Return
                End If
                OtherPrompt = OtherPrompt.Substring(AddPrefix.Length).Trim()
                DoInplace = False
            ElseIf OtherPrompt.StartsWith(AddPrefix2, StringComparison.OrdinalIgnoreCase) Then
                If NoText Then
                    ShowPrefixRequiresSelection(AddPrefix2)
                    Return
                End If
                OtherPrompt = OtherPrompt.Substring(AddPrefix2.Length).Trim()
                DoInplace = False
            ElseIf OtherPrompt.StartsWith(MarkupPrefix, StringComparison.OrdinalIgnoreCase) Then
                If NoText Then
                    ShowPrefixRequiresSelection(MarkupPrefix)
                    Return
                End If
                OtherPrompt = OtherPrompt.Substring(MarkupPrefix.Length).Trim()
                DoMarkup = True
            ElseIf OtherPrompt.StartsWith(MarkupPrefixRegex, StringComparison.OrdinalIgnoreCase) Then
                If NoText Then
                    ShowPrefixRequiresSelection(MarkupPrefixRegex)
                    Return
                End If
                OtherPrompt = OtherPrompt.Substring(MarkupPrefixRegex.Length).Trim()
                DoMarkup = True
                MarkupMethod = 4
            ElseIf OtherPrompt.StartsWith(MarkupPrefixWord, StringComparison.OrdinalIgnoreCase) Then
                If NoText Then
                    ShowPrefixRequiresSelection(MarkupPrefixWord)
                    Return
                End If
                OtherPrompt = OtherPrompt.Substring(MarkupPrefixWord.Length).Trim()
                DoMarkup = True
                MarkupMethod = 1
            ElseIf OtherPrompt.StartsWith(MarkupPrefixDiffW, StringComparison.OrdinalIgnoreCase) Then
                If NoText Then
                    ShowPrefixRequiresSelection(MarkupPrefixDiffW)
                    Return
                End If
                OtherPrompt = OtherPrompt.Substring(MarkupPrefixDiffW.Length).Trim()
                DoMarkup = True
                MarkupMethod = 3
            ElseIf OtherPrompt.StartsWith(MarkupPrefixDiff, StringComparison.OrdinalIgnoreCase) Then
                If NoText Then
                    ShowPrefixRequiresSelection(MarkupPrefixDiff)
                    Return
                End If
                OtherPrompt = OtherPrompt.Substring(MarkupPrefixDiff.Length).Trim()
                DoMarkup = True
                MarkupMethod = 2
            ElseIf OtherPrompt.StartsWith(PanePrefix, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(PanePrefix.Length).Trim()
                DoPane = True
                DoClipboard = True
                DoChunks = False
            ElseIf OtherPrompt.StartsWith(PushbackPrefix, StringComparison.OrdinalIgnoreCase) Then
                If NoText Then
                    ShowPrefixRequiresSelection(PushbackPrefix)
                    Return
                End If
                OtherPrompt = OtherPrompt.Substring(PushbackPrefix.Length).Trim()
                DoPushback = True
                DoChunks = False
                DoBubblesExtract = True
            ElseIf OtherPrompt.StartsWith(PushbackPrefix2, StringComparison.OrdinalIgnoreCase) Then
                If NoText Then
                    ShowPrefixRequiresSelection(PushbackPrefix2)
                    Return
                End If
                OtherPrompt = OtherPrompt.Substring(PushbackPrefix2.Length).Trim()
                DoPushback = True
                DoChunks = False
                DoBubblesExtract = True
            ElseIf OtherPrompt.StartsWith(FilePrefix, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(FilePrefix.Length).Trim()
                DoFiles = True
            ElseIf OtherPrompt.StartsWith(FilePrefix2, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(FilePrefix2.Length).Trim()
                DoFiles = True
            ElseIf OtherPrompt.StartsWith(AssemblePrefix, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(AssemblePrefix.Length).Trim()
                DoAssemble = True
            ElseIf OtherPrompt.StartsWith(FormPrefix, StringComparison.OrdinalIgnoreCase) Then
                If Not NoText Then
                    ShowCustomMessageBox($"The '{FormPrefix}' prefix is available only when no text is selected.")
                    Return
                End If
                OtherPrompt = OtherPrompt.Substring(FormPrefix.Length).Trim()
                DoForm = True
            End If

            ' (net) trigger: Enable internet search
            If OtherPrompt.IndexOf(NetTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(NetTrigger, "").Trim()
                DoNet = True
            End If

            ' === TOOLING TRIGGER HANDLING ===

            ' Check if user wants to (re)select tools
            Dim DoToolSelection As Boolean = False
            If OtherPrompt.IndexOf(ToolSelectionTrigger, StringComparison.OrdinalIgnoreCase) >= 0 And Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                If modelSupportsTool Then
                    OtherPrompt = OtherPrompt.Replace(ToolSelectionTrigger, "").Trim()
                    DoToolSelection = True
                End If
            End If

            ' If model supports tooling, handle tool selection
            Dim selectedToolsForSession As List(Of ModelConfig) = Nothing
            If modelSupportsTool AndAlso Not DoFiles Then
                selectedToolsForSession = SelectToolsForSession(DoToolSelection, ToolFriendlyName)
                If selectedToolsForSession Is Nothing Then
                    ' User cancelled tool selection
                    If originalConfigLoaded Then
                        RestoreDefaults(_context, originalConfig)
                        originalConfigLoaded = False
                    End If
                    Return
                End If
            End If


            ' === Multi-model selection ===

            ' (multimodel) trigger: Prompt for multiple model selection
            SelectedAlternateModels = Nothing
            If UseSecondAPI AndAlso Not toolTriggerDetected AndAlso Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) AndAlso OtherPrompt.IndexOf(MultiModelTrigger, StringComparison.OrdinalIgnoreCase) >= 0 AndAlso Not DoFiles Then
                If Not DoMarkup AndAlso Not DoBubbles AndAlso Not DoPushback AndAlso Not DoSlides AndAlso DoChart = 0 Then
                    If Not ShowMultipleModelSelection(_context, INI_AlternateModelPath) OrElse SelectedAlternateModels Is Nothing OrElse SelectedAlternateModels.Count = 0 Then
                        Return
                    End If
                Else
                    ShowCustomMessageBox($"The multi-model feature cannot be used together with markup, bubbles or slides - will continue only with the model you already selected.")
                End If
                OtherPrompt = OtherPrompt.Replace(MultiModelTrigger, "").Trim()
            End If

            ' === MyStyle prompt integration ===

            ' (mystyle) trigger: Select and apply personal style prompt
            If Not String.IsNullOrWhiteSpace(INI_MyStylePath) AndAlso OtherPrompt.IndexOf(MyStyleTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                Dim StylePath As String = ExpandEnvironmentVariables(INI_MyStylePath)
                If Not IO.File.Exists(StylePath) Then
                    ShowCustomMessageBox("No MyStyle prompt file has been found. You may have to first create a MyStyle prompt. Go to 'Analyze' and use 'Define MyStyle' to do so - will abort.")
                    Return
                End If
                OtherPrompt = OtherPrompt.Replace(MyStyleTrigger, "").Trim()
                MyStyleInsert = MyStyleHelpers.SelectPromptFromMyStyle(StylePath, "Word", 0, "Choose the style prompt to apply …", $"{AN} MyStyle", False)
                If MyStyleInsert = "ERROR" Then Return
                If MyStyleInsert = "NONE" OrElse String.IsNullOrWhiteSpace(MyStyleInsert) Then Return
                DoMyStyle = True
            End If

            ' === Additional document integration ===

            ' (adddoc) trigger: Gather content from other open Word documents
            If Not String.IsNullOrEmpty(OtherPrompt) AndAlso OtherPrompt.IndexOf(AddDocTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then

                InsertDocs = GatherSelectedDocuments()
                Debug.WriteLine($"GatherSelectedDocs returned: {Left(InsertDocs, 3000)}")
                If String.IsNullOrWhiteSpace(InsertDocs) Then
                    ShowCustomMessageBox("No content was found or an error occurred in gathering the additional document(s) - will abort.")
                    Return
                ElseIf InsertDocs.StartsWith("ERROR", StringComparison.OrdinalIgnoreCase) Then
                    ShowCustomMessageBox($"An error occured gathering the additional document(s) ({InsertDocs.Substring(6).Trim()}) - will abort.")
                    Return
                ElseIf InsertDocs.StartsWith("NONE", StringComparison.OrdinalIgnoreCase) Then
                    ShowCustomMessageBox($"There are no other documents to add - will abort.")
                    Return
                End If
                OtherPrompt = Regex.Replace(OtherPrompt, Regex.Escape(AddDocTrigger), "", RegexOptions.IgnoreCase)
            End If

            OtherPromptUnfilled = OtherPrompt.Trim()

            ' === External file/directory embedding ===

            ' Handles {doc}, {dir}, {url} and {path} triggers with unified document numbering

            Dim fileResult = Await ProcessExternalFileTriggers(OtherPrompt)
            If Not fileResult.Success Then
                Return
            End If
            OtherPrompt = fileResult.ModifiedPrompt


            ' === Assemble dispatch (after all external triggers are resolved) ===
            ' InsertDocs now contains (adddoc) content, and OtherPrompt has had
            ' {doc}/{dir}/{url} triggers replaced with loaded document content.
            ' Both the user instruction and the additional context are fully populated.

            If DoAssemble Then
                ' Combine any embedded document content from {doc}/{dir}/{url} triggers
                ' with (adddoc) content as additional context for the assembler.
                Dim assembleContext As String = ""
                If Not String.IsNullOrWhiteSpace(InsertDocs) Then
                    assembleContext = InsertDocs
                End If

                Try
                    Await AssembleDocumentFromTemplates(OtherPrompt, assembleContext, UseSecondAPI)
                Catch ex As System.Exception
                    ShowCustomMessageBox("Error in Freestyle ('Assemble:'): " & ex.Message, "Error")
                End Try
                If UseSecondAPI And originalConfigLoaded Then
                    RestoreDefaults(_context, originalConfig)
                    originalConfigLoaded = False
                End If
                Return
            End If

            ' ======== COMPLETEFORMS mode (after all other processing, including file embedding, is done) ========

            If DoForm Then
                Try
                    Await CompleteWordDocumentTables(OtherPrompt, UseSecondAPI)
                Catch ex As System.Exception
                    ShowCustomMessageBox("Error in Freestyle ('Form:'): " & ex.Message, "Error")
                End Try

                If UseSecondAPI AndAlso originalConfigLoaded Then
                    RestoreDefaults(_context, originalConfig)
                    originalConfigLoaded = False
                End If

                Return
            End If


            ' === File object selection (for LLM APIs that support file attachments) ===

            If DoFileObject And Not DoFiles Then
                If DoFileObjectClip Then
                    ' Use clipboard content as file object
                    FileObject = "clipboard"
                Else
                    ' Prompt user to select file
                    DragDropFormLabel = "All file types that are supported by your LLM."
                    DragDropFormFilter = "Supported Files|*.*"
                    FileObject = GetFileName()
                    DragDropFormLabel = ""
                    DragDropFormFilter = ""
                    If String.IsNullOrWhiteSpace(FileObject) Then
                        ShowCustomMessageBox("No file object has been selected - will abort. You can try again (use Ctrl-P to re-insert your prompt).")
                        Return
                    End If
                End If
            End If

            ' === PowerPoint slide deck selection or creation ===

            If DoSlides Then
                DragDropFormLabel = "A Powerpoint (pptx) file (or cancel to create one)."
                DragDropFormFilter = "Supported Files|*.pptx"
                SlideDeck = GetFileName()
                DragDropFormLabel = ""
                DragDropFormFilter = ""

                ' If no file selected, offer to create new presentation
                If String.IsNullOrWhiteSpace(SlideDeck) Then

                    Dim CreatePPTX As Integer = ShowCustomYesNoBox(
                         "You have not provided a Powerpoint file. Do you want create a new one?", "Yes", "No, abort")
                    If CreatePPTX <> 1 Then
                        ShowCustomMessageBox("No Powerpoint file has been selected - will abort. You can try again (use Ctrl-P to re-insert your prompt).")
                        Return
                    End If

                    ' Default to Desktop\NewPresentation.pptx
                    If String.IsNullOrWhiteSpace(SlideDeck) Then
                        Dim desktop As String = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop)
                        SlideDeck = System.IO.Path.Combine(desktop, "NewPresentation.pptx")
                    End If

                    ' Ensure unique filename: NewPresentation (2).pptx, (3).pptx, etc.
                    If System.IO.File.Exists(SlideDeck) Then
                        Dim dir As String = System.IO.Path.GetDirectoryName(SlideDeck)
                        Dim name As String = System.IO.Path.GetFileNameWithoutExtension(SlideDeck)
                        Dim ext As String = System.IO.Path.GetExtension(SlideDeck)
                        Dim i As Integer = 2
                        Do
                            Dim candidate As String = System.IO.Path.Combine(dir, name & " (" & i.ToString() & ")" & ext)
                            If Not System.IO.File.Exists(candidate) Then
                                SlideDeck = candidate
                                Exit Do
                            End If
                            i += 1
                        Loop
                    End If

                    ' Create blank PowerPoint presentation
                    Dim pptApp As NetOffice.PowerPointApi.Application = Nothing
                    Dim presentation As NetOffice.PowerPointApi.Presentation = Nothing

                    Try
                        pptApp = New NetOffice.PowerPointApi.Application()

                        ' Make PowerPoint visible (intentional per requirements)
                        pptApp.Visible = NetOffice.OfficeApi.Enums.MsoTriState.msoTrue
                        pptApp.DisplayAlerts = NetOffice.PowerPointApi.Enums.PpAlertLevel.ppAlertsNone
                        pptApp.WindowState = NetOffice.PowerPointApi.Enums.PpWindowState.ppWindowNormal

                        ' Create new presentation with window
                        presentation = pptApp.Presentations.Add(NetOffice.OfficeApi.Enums.MsoTriState.msoTrue)

                        ' Save as Open XML format (.pptx)
                        presentation.SaveAs(
                            SlideDeck,
                            NetOffice.PowerPointApi.Enums.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,
                            NetOffice.OfficeApi.Enums.MsoTriState.msoFalse
                        )

                        ' Close presentation and quit PowerPoint
                        presentation.Close()
                        pptApp.Quit()

                    Catch comEx As System.Runtime.InteropServices.COMException
                        ' Handle COM-specific errors when creating PowerPoint file
                        ShowCustomMessageBox("PowerPoint COM error while creating file:" & vbCrLf &
                     "Message: " & comEx.Message & vbCrLf &
                     "HResult: 0x" & comEx.HResult.ToString("X8"))
                        Return

                    Catch ex As System.Exception
                        ' Handle general errors when creating PowerPoint file
                        ShowCustomMessageBox("Error while creating PowerPoint file: " & ex.ToString())
                        Return

                    Finally
                        ' Dispose COM objects in correct order to prevent memory leaks and zombie processes
                        If presentation IsNot Nothing Then
                            Try : presentation.Dispose() : Catch : End Try
                            presentation = Nothing
                        End If
                        If pptApp IsNot Nothing Then
                            Try : pptApp.Quit() : Catch : End Try
                            Try : pptApp.Dispose() : Catch : End Try
                            pptApp = Nothing
                        End If
                    End Try


                End If
            End If

            ' === User confirmation for processing without text selection ===

            ' Prompt user to process full document when bubbles or chunks mode is active but no text selected
            If NoText AndAlso (DoBubbles Or DoChunks) AndAlso Not DoFiles Then
                Dim FullDocument As Integer = ShowCustomYesNoBox("You have not selected text. Ask the LLM to comment on the full document?", "Yes", "No, abort")
                If FullDocument = 1 Then
                    Dim document As Word.Document = application.ActiveDocument
                    document.Content.Select()
                    NoText = False
                Else
                    Return
                End If
            End If

            ' Prompt user to process full document when markup mode is active but no text selected
            ' Not used anymore under normal circumstances (gates introduced above), but kept as a fallback in case of misconfiguration or unexpected states
            If NoText AndAlso DoMarkup AndAlso Not DoFiles Then
                Dim FullDocument As Integer = ShowCustomYesNoBox("You have not selected text. Do the markup on the full document?", "Yes", "No, abort")
                If FullDocument = 1 Then
                    Dim document As Word.Document = application.ActiveDocument
                    document.Content.Select()
                    NoText = False
                Else
                    Return
                End If
            End If

            ' Confirm markup placement when configuration specifies append mode
            If Not DoInplace AndAlso DoMarkup AndAlso Not Dofiles Then
                Dim AppendMarkup As Integer = ShowCustomYesNoBox("You have asked for a markup to be created, but according to the configuration, it will not replace your current selection but added to it at the end. Is this really what you want?", "Yes, add markup ", "No, replace text with markup")
                If AppendMarkup = 0 Then
                    Return
                ElseIf AppendMarkup = 2 Then
                    DoInplace = True
                    NoFormatAndFieldSaving = False
                End If
            End If

            ' === System prompt construction based on selected mode ===

            If DoFiles Then
                Try
                    CorrectWordDocuments(SP_Freestyle_Document & " " & MyStyleInsert & " " & InsertDocs, "_freestyle", UseSecondAPI, True)
                Catch ex As System.Exception
                    ' Handle any unexpected errors during freestyle execution
                    ShowCustomMessageBox("Error in Freestyle ('File:'): " & ex.Message, "Error")
                End Try
                Return
            End If

            ' Handle pure prefix mode: use prompt directly as system prompt without additional processing
            If OtherPrompt.StartsWith(PurePrefix, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(PurePrefix.Length).Replace("(a file object follows)", "").Replace("(a clipboard object follows)", "").Trim()
                SysPrompt = OtherPrompt
                DoClipboard = True
                DoChunks = False
            Else
                ' Construct system prompt based on selected feature mode
                If DoLib Then
                    ' Library search mode: consult library and update SysPrompt
                    Dim isSuccess As Boolean = Await ConsultLibrary(DoMarkup)
                    If Not isSuccess Then Return
                ElseIf DoNet Then
                    ' Internet search mode: consult internet and update SysPrompt
                    Dim isSuccess As Boolean = Await ConsultInternet(DoMarkup)
                    If Not isSuccess Then Return
                ElseIf NoText Then
                    ' No text selected: use freestyle prompt without text processing
                    SysPrompt = SP_FreestyleNoText
                Else
                    ' Standard text processing mode
                    SysPrompt = SP_FreestyleText
                    If DoBubbles Then SysPrompt = SysPrompt & " " & SP_Add_Bubbles
                    If DoPushback Then SysPrompt = SysPrompt & " " & SP_Add_BubblesReply
                    If INI_MarkdownBubbles Then FormatInstruction = SP_Add_Bubbles_Format Else FormatInstruction = ""
                End If

                ' (kb) / (kb:...) trigger: Knowledge store RAG
                If DoKB Then
                    Dim kbRequest = KnowledgeTriggerHelper.TryParseKnowledgeTrigger(OtherPrompt)
                    If kbRequest IsNot Nothing Then
                        Dim strippedPrompt = KnowledgeTriggerHelper.StripKnowledgeTrigger(OtherPrompt, kbRequest)
                        Dim knowledgeTaskPrompt As String = strippedPrompt.Trim()

                        If String.IsNullOrWhiteSpace(strippedPrompt) Then
                            If Not String.IsNullOrWhiteSpace(kbRequest.SearchQuery) Then
                                OtherPrompt = kbRequest.SearchQuery.Trim()
                            ElseIf kbRequest.Tags IsNot Nothing AndAlso kbRequest.Tags.Length > 0 Then
                                OtherPrompt = "Answer based on the provided Knowledge Store content, focusing on: " &
                                              String.Join(", ", kbRequest.Tags)
                            ElseIf Not String.IsNullOrWhiteSpace(kbRequest.StoreName) Then
                                OtherPrompt = "Answer based on the provided Knowledge Store content from store '" &
                                              kbRequest.StoreName & "'."
                            Else
                                OtherPrompt = "Answer based on the provided Knowledge Store content."
                            End If
                        Else
                            OtherPrompt = strippedPrompt
                        End If

                        Dim kbResolveOptions As KnowledgeTriggerHelper.KnowledgeResolveOptions = Nothing
                        If Not String.IsNullOrWhiteSpace(knowledgeTaskPrompt) Then
                            kbResolveOptions = New KnowledgeTriggerHelper.KnowledgeResolveOptions With {
                                .TaskPrompt = knowledgeTaskPrompt,
                                .IncludeRelevantExtracts = True,
                                .IncludeFullDocumentContent = False
                            }
                        End If

                        Dim kbSplash As New SLib.SplashScreen("Querying Knowledge Store...   ")
                        kbSplash.Show()
                        System.Windows.Forms.Application.DoEvents()

                        Dim kbResolved As (Content As String, StatusMessage As String)
                        Try
                            kbResolved = Await KnowledgeTriggerHelper.ResolveKnowledgeAsync(kbRequest, _context, kbResolveOptions)
                        Finally
                            If kbSplash.InvokeRequired Then
                                kbSplash.Invoke(Sub()
                                                    kbSplash.Close()
                                                    kbSplash.Dispose()
                                                End Sub)
                            Else
                                kbSplash.Close()
                                kbSplash.Dispose()
                            End If
                        End Try

                        If Not String.IsNullOrWhiteSpace(kbResolved.Content) Then
                            SysPrompt = SysPrompt & vbCrLf & vbCrLf &
                                "The following documents from the user's knowledge store are provided as reference material. " &
                                "Use them to answer the user's question. " &
                                "When citing information, prefer clickable markdown citations. " &
                                "If a KSDOCUMENT element provides a sourcePath attribute, cite it as [Source](sourcePath). " &
                                "If a wikiPath attribute is available and helpful, you may also cite [Wiki](wikiPath). " &
                                "Do not invent links and do not fabricate paths. Use only the paths explicitly provided in the KSDOCUMENT metadata." & vbCrLf &
                                kbResolved.Content
                        Else
                            ShowCustomMessageBox(If(String.IsNullOrWhiteSpace(kbResolved.StatusMessage),
                                                    "No documents found in the Knowledge Store.",
                                                    kbResolved.StatusMessage), $"{AN} Knowledge Store")
                            Exit Sub
                        End If
                    End If
                End If
            End If

            ' === Chunk processing configuration ===

            ' Prompt user for chunk size when chunk mode is active
            If DoChunks Then
                Dim response As String = SLib.ShowCustomInputBox($"How many paragraphs shall be treated at the same time (max. 25)?", "Iterate through the text", True, ChunkSize.ToString()).Trim()
                If Not Integer.TryParse(response, ChunkSize) Then ChunkSize = 0
                If response = "" OrElse response.ToLower() = "esc" OrElse ChunkSize = 0 Then Return
                If ChunkSize > 25 Then ChunkSize = 25
            Else
                ChunkSize = 0
            End If

            'Debug.WriteLine("Freestyle Prompt: " & SysPrompt)

            ' === Execute LLM processing with configured parameters ===

            ' Invoke ProcessSelectedText with all configured options
            Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SysPrompt), True, DoKeepFormat, DoKeepParaFormat, DoInplace, DoMarkup, MarkupMethod, DoClipboard, DoBubbles, False, UseSecondAPI, KeepFormatCap, DoTPMarkup, TPMarkupName, False, FileObject, DoPane, ChunkSize, NoFormatAndFieldSaving, DoNewDoc, SlideDeck, InsertDocs <> "", DoMyStyle, DoBubblesExtract, DoPushback, selectedToolsForSession, DoChart, DoShowModel)

        Catch ex As System.Exception
            ' Handle any unexpected errors during freestyle execution
            ShowCustomMessageBox("Error in Freestyle: " & ex.Message)
        Finally
            If UseSecondAPI AndAlso originalConfigLoaded Then
                RestoreDefaults(_context, originalConfig)
                originalConfigLoaded = False
            End If
        End Try
    End Sub

End Class

