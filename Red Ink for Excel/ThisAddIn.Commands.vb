' Part of "Red Ink for Excel"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.Commands.vb
' Purpose: Defines command methods invoked by the UI for translation, correction,
'          improvement, anonymization, shortening, party switching, freestyle prompt
'          execution, chat display, help window, and settings management. Methods
'          gather user input, set state variables, and delegate processing to
'          ProcessSelectedRange or other helper routines.
'
' Architecture:
'   - Each public async command:
'       1. Calls Application.DoEvents.
'       2. Optionally acquires the current Excel.Range selection.
'       3. Gathers user input via ShowCustomInputBox / prompt library.
'       4. Sets global/context variables (TranslateLanguage, Context, OldParty/NewParty, etc.).
'       5. Invokes ProcessSelectedRange with a system prompt constant (SP_*), passing flags
'          controlling range vs cell-by-cell vs formula handling, color check, pane usage,
'          file object inclusion, batch path, worksheet content, and shortening percentage.
'   - Freestyle(...) parses prefix triggers (CellByCellPrefix, CellByCellPrefix2, TextPrefix,
'       TextPrefix2, BubblesPrefix, PanePrefix, BatchPrefix, PurePrefix, ExtTrigger, ExtTriggerFixed,
'       ExtWSTrigger, ColorTrigger, ObjectTrigger, ObjectTrigger2) to derive execution mode
'       and augment OtherPrompt with file content, worksheet text, color analysis, file object
'       reference, clipboard object, or batch directory processing.
'   - Batch processing: collects a target insertion line (LineNumber) and directory path
'       (BatchPath), validating presence of allowedExtensions.
'   - File/worksheet inclusion: replaces ExtTrigger and ExtTriggerFixed occurrences with tagged file content; adds
'       additional worksheet text via GatherSelectedWorksheets when ExtWSTrigger is present.
'   - Settings: builds dictionaries of setting names/descriptions and displays a settings window;
'       refreshes menus via AddContextMenu using a SplashScreen.
'   - ShowChatForm: creates and positions frmAIChat with persisted size/location (My.Settings).
'   - HelpMeInky: lazy-instantiates HelpMeInky and calls ShowRaised.
'   - Return values: Methods declare a Boolean result variable from ProcessSelectedRange or other
'       operations; result is not explicitly returned (no Return statement) in this file.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn


    ' Add these at the top of the ThisAddIn partial class, before the command methods

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
        ''' <see cref="GetFileContentEx"/> (text-based, Office, PDF, email, and
        ''' binary/media types handled via LLM).
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
    ''' Asynchronously loads a single file and returns its content wrapped in numbered
    ''' XML document tags (e.g. <c>&lt;document1&gt;...&lt;/document1&gt;</c>).
    ''' </summary>
    ''' <remarks>
    ''' <para>
    ''' Delegates to <see cref="GetFileContentEx"/> for the actual file reading, which
    ''' supports text files, RTF, Word (.doc/.docx), Excel (.xlsx), PowerPoint (.pptx),
    ''' PDF (with optional OCR), email (.eml/.msg), and binary/media files (images,
    ''' audio, video) via LLM extraction.
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
    Private Async Function LoadSingleFileAsync(filePath As String,
                                               isWrapped As Boolean,
                                               ctx As FileLoadingContext,
                                               Optional isFromDirectory As Boolean = False,
                                               Optional askWorksheetSelection As Boolean = False) As Task(Of String)
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
                AskWorksheetSelection:=askWorksheetSelection)

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
            If Not IO.Directory.Exists(dirPath) Then
                ctx.FailedFiles.Add(dirPath)
                Return ""
            End If

            Dim searchOption As IO.SearchOption = If(FileLoadingContext.RecurseDirectories, IO.SearchOption.AllDirectories, IO.SearchOption.TopDirectoryOnly)
            Dim allFiles As String() = IO.Directory.GetFiles(dirPath, "*.*", searchOption)

            ' Filter to supported extensions
            Dim supportedFiles As New List(Of String)()
            Dim ignoredFiles As New List(Of String)()

            For Each f In allFiles
                Dim ext As String = IO.Path.GetExtension(f).ToLowerInvariant()
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
            Dim pdfFiles As List(Of String) = supportedFiles.Where(Function(f) IO.Path.GetExtension(f).Equals(".pdf", StringComparison.OrdinalIgnoreCase)).ToList()
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


            ' If pdfCount <= 1 or OCR not available, we don't set DirectoryOCRMode

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
                ProgressBarModule.GlobalProgressLabel = $"Loading file {dirFileIndex} of {dirFileCount}: {IO.Path.GetFileName(filePath)}..."
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
    ''' Processes all external file/directory triggers in the prompt ({doc}, {dir}, and fixed paths).
    ''' Processes triggers in positional order to maintain correct document numbering.
    ''' </summary>
    ''' <param name="prompt">The prompt to process.</param>
    ''' <returns>A tuple containing: (success flag, modified prompt). If success is False, the operation should abort.</returns>
    Private Async Function ProcessExternalFileTriggers(prompt As String) As Task(Of (Success As Boolean, ModifiedPrompt As String))
        ' Check if any file triggers are present
        Dim hasExtTriggers As Boolean = prompt.IndexOf(ExtTrigger, StringComparison.OrdinalIgnoreCase) >= 0
        Dim hasDirTriggers As Boolean = prompt.IndexOf(ExtDirTrigger, StringComparison.OrdinalIgnoreCase) >= 0
        Dim hasFixedPaths As Boolean = False

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
        If Not hasExtTriggers AndAlso Not hasDirTriggers AndAlso Not hasFixedPaths Then
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
        Dim fixedPathCount As Integer = 0
        If Not String.IsNullOrEmpty(patternFixed) Then
            ' Count only fixed path matches that are not already matched by other trigger types
            ' (e.g., the pattern {[path]} -> \{(?<path>.*?)\} would also match {doc}, {dir})
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
                ' Also skip if it doesn't look like a path
                Dim looksLikePath As Boolean = candidatePath.Contains("\") OrElse candidatePath.Contains("/") OrElse candidatePath.Contains(":")
                If Not looksLikePath Then Continue For
                fixedPathCount += 1
            Next
        End If
        ctx.ExpectedFileCount = extTriggerCount + fixedPathCount
        ' Note: dirTriggerCount files are added in LoadDirectoryFilesAsync

        ' NOTE: OCR prompt is handled per-directory when multiple PDFs are found,
        ' or per-file with AskUser=True for individual files/single PDFs

        ' === Track overall file processing progress ===
        Dim totalTriggers As Integer = extTriggerCount + dirTriggerCount + fixedPathCount
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
            Dim fixedIdx As Integer = -1
            Dim fixedMatch As Match = Nothing

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
                            IO.Path.GetExtension(selectedFile).Equals(".xlsx", StringComparison.OrdinalIgnoreCase)

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
                            ProgressBarModule.GlobalProgressLabel = $"Loading file {processedTriggers + 1} of {totalTriggers}: {IO.Path.GetFileName(selectedFile)}..."
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
                    If IO.File.Exists(candidatePath) Then
                        Dim askWorksheetSelection As Boolean =
                            IO.Path.GetExtension(candidatePath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase)

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
                            ProgressBarModule.GlobalProgressLabel = $"Loading {processedTriggers + 1} of {totalTriggers}: {IO.Path.GetFileName(candidatePath)}..."
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

                    ElseIf IO.Directory.Exists(candidatePath) Then
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
    ''' Translates the selected range to INI_Language1 using SP_Translate. Sets TranslateLanguage.
    ''' </summary>
    ''' <returns>Task(Of Boolean). Result variable not explicitly returned.</returns>
    Public Async Function InLanguage1() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        TranslateLanguage = INI_Language1
        Dim result As Boolean = Await ProcessSelectedRange(SP_Translate, True, False, False, False, True, False)
        Return result
    End Function

    ''' <summary>
    ''' Translates the selected range to INI_Language2 using SP_Translate. Sets TranslateLanguage.
    ''' </summary>
    ''' <returns>Task(Of Boolean). Result variable not explicitly returned.</returns>
    Public Async Function InLanguage2() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        TranslateLanguage = INI_Language2
        Dim result As Boolean = Await ProcessSelectedRange(SP_Translate, True, False, False, False, True, False)
        Return result
    End Function

    ''' <summary>
    ''' Prompts for a target language, selects current range if available, translates using SP_Translate.
    ''' </summary>
    ''' <returns>Task(Of Boolean). Result variable not explicitly returned. Exits early if no language provided.</returns>
    Public Async Function InOther() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        Dim selectedRange As Excel.Range = TryCast(Globals.ThisAddIn.Application.Selection, Excel.Range)
        TranslateLanguage = Global.SharedLibrary.SharedLibrary.SharedMethods.PromptForTargetLanguage(
            "Enter your target language:",
            $"{AN} Translate",
            "",
            _context)
        If Not String.IsNullOrEmpty(TranslateLanguage) Then
            If selectedRange IsNot Nothing Then
                selectedRange.Select()
            End If
            Dim result As Boolean = Await ProcessSelectedRange(SP_Translate, True, False, False, False, True, False)
            Return result
        End If

        Return False
    End Function

    ''' <summary>
    ''' Prompts for a target language and translates including formulas (third flag True) using SP_Translate.
    ''' </summary>
    ''' <returns>Task(Of Boolean). Result variable not explicitly returned.</returns>
    Public Async Function InOtherFormulas() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        TranslateLanguage = Global.SharedLibrary.SharedLibrary.SharedMethods.PromptForTargetLanguage(
            "Enter your target language:",
            $"{AN} Translate",
            "",
            _context)
        If Not String.IsNullOrEmpty(TranslateLanguage) Then
            Dim result As Boolean = Await ProcessSelectedRange(SP_Translate, True, False, True, False, True, False)
            Return result
        End If

        Return False
    End Function

    ''' <summary>
    ''' Performs correction on selected range using SP_Correct.
    ''' </summary>
    ''' <returns>Task(Of Boolean). Result variable not explicitly returned.</returns>
    Public Async Function Correct() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        Dim result As Boolean = Await ProcessSelectedRange(SP_Correct, True, False, False, False, True, False)
        Return result
    End Function

    ''' <summary>
    ''' Prompts for optional context, defaults to "n/a" if blank, then improves writing using SP_WriteNeatly.
    ''' Requires a selected range.
    ''' </summary>
    ''' <returns>Task(Of Boolean). Returns False early if no selection. Result variable not explicitly returned.</returns>
    Public Async Function Improve() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        Dim selectedRange As Excel.Range = TryCast(Globals.ThisAddIn.Application.Selection, Excel.Range)
        If selectedRange Is Nothing Then
            ShowCustomMessageBox("Please select the cells to be processed.")
            Return False
        End If
        Context = Trim(SLib.ShowCustomInputBox("Please provide the context that should be taken into account, if any:", $"{AN} Write Neatly", True))
        If String.IsNullOrWhiteSpace(Context) Then
            Context = "n/a"
        End If
        If selectedRange IsNot Nothing Then
            selectedRange.Select()
        End If
        Dim result As Boolean = Await ProcessSelectedRange(SP_WriteNeatly, True, False, False, False, True, False)
        Return result
    End Function

    ''' <summary>
    ''' Anonymizes selected range content using SP_Anonymize.
    ''' </summary>
    ''' <returns>Task(Of Boolean). Result variable not explicitly returned.</returns>
    Public Async Function Anonymize() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        Dim result As Boolean = Await ProcessSelectedRange(SP_Anonymize, True, False, False, False, True, False)
        Return result
    End Function

    ''' <summary>
    ''' Prompts for a shortening percentage (1–99), computes average/max length of unprotected non-formula cells,
    ''' then shortens content using SP_Shorten with percentage parameter.
    ''' </summary>
    ''' <returns>Task(Of Boolean). Returns False early on invalid selection or input. Result variable not explicitly returned.</returns>
    Public Async Function Shorten() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        Dim selectedRange As Excel.Range = TryCast(Globals.ThisAddIn.Application.Selection, Excel.Range)
        If selectedRange Is Nothing Then
            ShowCustomMessageBox("Please select the cells to be processed.")
            Return False
        End If
        Dim totalLength As Integer = 0
        Dim maxLength As Integer = 0
        Dim cellCount As Integer = 0
        For Each cell As Excel.Range In selectedRange.Cells
            If Not CellProtected(cell) AndAlso Not CBool(cell.HasFormula) Then
                Dim cellText As String = CStr(cell.Value)
                If Not String.IsNullOrEmpty(cellText) Then
                    Dim textLength As Integer = GetNumberOfWords(cellText)
                    totalLength += textLength
                    If textLength > maxLength Then
                        maxLength = textLength
                    End If
                    cellCount += 1
                End If
            End If
        Next
        Dim averageLength As Double = If(cellCount > 0, totalLength / cellCount, 0)
        Dim UserInput As String
        Dim ShortenPercentValue As Integer = 0
        Do
            UserInput = Trim(SLib.ShowCustomInputBox($"Enter the percentage by which the text of each selected cell should be shortened (the cells have have of average {averageLength:n1} words and {maxLength} at max; {ShortenPercent}% will cut approx. " & (averageLength * ShortenPercent / 100) & " words in average):", $"{AN} Shortener", True, CStr(ShortenPercent) & "%"))
            If String.IsNullOrEmpty(UserInput) Then
                Return False
            End If
            UserInput = UserInput.Replace("%", "").Trim()
            If Integer.TryParse(UserInput, ShortenPercentValue) AndAlso ShortenPercentValue >= 1 AndAlso ShortenPercentValue <= 99 Then
                Exit Do
            Else
                ShowCustomMessageBox("Please enter a valid percentage between 1 and 99.")
            End If
        Loop
        If ShortenPercentValue = 0 Then Return False
        If selectedRange IsNot Nothing Then
            selectedRange.Select()
        End If
        Dim result As Boolean = Await ProcessSelectedRange(SP_Shorten, True, False, False, False, True, False, ShortenPercentValue, False)
        Return result
    End Function

    ''' <summary>
    ''' Prompts for old and new party names separated by a comma, sets OldParty/NewParty, then switches occurrences using SP_SwitchParty.
    ''' Requires a selected range.
    ''' </summary>
    ''' <returns>Task(Of Boolean). Returns False early if selection missing or input invalid. Result variable not explicitly returned.</returns>
    Public Async Function SwitchParty() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        Dim selectedRange As Excel.Range = TryCast(Globals.ThisAddIn.Application.Selection, Excel.Range)
        If selectedRange Is Nothing Then
            ShowCustomMessageBox("Please select the cells to be processed.")
            Return False
        End If
        Dim UserInput As String
        Do
            UserInput = Trim(SLib.ShowCustomInputBox("Please provide the original party name and the new party name, separated by a comma (example: Elvis Presley, Taylor Swift):", $"{AN} Switch Party", True))
            If String.IsNullOrEmpty(UserInput) Then
                Return False
            End If
            Dim parts() As String = UserInput.Split(","c)
            If parts.Length = 2 Then
                OldParty = parts(0).Trim()
                NewParty = parts(1).Trim()
                Exit Do
            Else
                ShowCustomMessageBox("Please enter two names separated by a comma.")
            End If
        Loop
        If selectedRange IsNot Nothing Then
            selectedRange.Select()
        End If
        Dim result As Boolean = Await ProcessSelectedRange(SP_SwitchParty, True, False, False, False, True, False)
        Return result
    End Function

    ''' <summary>
    ''' Shows or creates the chat form (frmAIChat) and restores size/location from My.Settings; brings form to front.
    ''' </summary>
    Public Sub ShowChatForm()
        If chatForm Is Nothing OrElse chatForm.IsDisposed Then
            chatForm = New frmAIChat(_context)
            ' Set the location and size before showing the form
            If My.Settings.FormLocation <> System.Drawing.Point.Empty AndAlso My.Settings.FormSize <> Size.Empty Then
                chatForm.StartPosition = FormStartPosition.Manual
                chatForm.Location = My.Settings.FormLocation
                chatForm.Size = My.Settings.FormSize
            Else
                ' Default to center screen if no settings are available
                chatForm.StartPosition = FormStartPosition.Manual
                Dim screenBounds As System.Drawing.Rectangle = Screen.PrimaryScreen.WorkingArea
                chatForm.Location = New System.Drawing.Point((screenBounds.Width - chatForm.Width) \ 2, (screenBounds.Height - chatForm.Height) \ 2)
                chatForm.Size = New Size(650, 500) ' Set default size if needed
            End If
        End If
        ' Show and bring the form to the front
        chatForm.Show()
        chatForm.BringToFront()
    End Sub

    ''' <summary>
    ''' Freestyle execution without alternate model (UseSecondAPI = False). Delegates to Freestyle(False).
    ''' </summary>
    ''' <returns>Task(Of Boolean). Result variable not explicitly returned.</returns>
    Public Async Function FreestyleNM() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        Dim result As Boolean = Await Freestyle(False)
        Return result
    End Function

    ''' <summary>
    ''' Freestyle execution with optional alternate model selection (if INI_AlternateModelPath provided).
    ''' Restores defaults if originalConfigLoaded is True after execution.
    ''' </summary>
    ''' <returns>Task(Of Boolean). Result variable not explicitly returned. Exits early if model selection canceled.</returns>
    Public Async Function FreestyleAM() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
            If Not ShowModelSelection(_context, INI_AlternateModelPath) Then
                originalConfigLoaded = False
                Return False
            End If
        End If
        Dim result As Boolean = Await Freestyle(True)
        If originalConfigLoaded Then
            RestoreDefaults(_context, originalConfig)
            originalConfigLoaded = False
        End If
        Return result
    End Function

    ''' <summary>
    ''' Central freestyle method parsing prefixes/triggers to set execution flags (cell-by-cell, text-only, comments, pane, batch, color check,
    ''' file object, worksheet inclusion). Handles prompt library, stored last prompt, file content insertion, worksheet gathering,
    ''' batch directory selection, line number input, and delegates to ProcessSelectedRange with appropriate system prompt or raw prompt.
    ''' </summary>
    ''' <param name="UseSecondAPI">True to use second API/model path; affects object trigger availability.</param>
    ''' <returns>Task(Of Boolean). Result variables not explicitly returned. Returns False early in multiple validation failure cases.</returns>
    Public Async Function Freestyle(ByVal UseSecondAPI As Boolean) As Task(Of Boolean)
        Dim selectedRange As Excel.Range = TryCast(Globals.ThisAddIn.Application.Selection, Excel.Range)
        Dim NoSelectedCells As Boolean = False
        Dim DoClipboard As Boolean = False
        Dim DoFormulas As Boolean = True
        Dim DoBubbles As Boolean = False
        Dim DoColor As Boolean = False
        Dim DoFileObject As Boolean = False
        Dim DoFileObjectClip As Boolean = False
        Dim DoPane As Boolean = False
        Dim DoBatch As Boolean = False
        Dim BatchPath As String = ""
        Dim LastPromptInstruct As String = If(String.IsNullOrWhiteSpace(My.Settings.LastPrompt), "", "; Ctrl-P for your last prompt")
        Dim PureInstruct As String = $"; use '{PurePrefix}' for direct prompting"
        Dim DefaultPrefix As String = INI_DefaultPrefix
        Dim DefaultPrefixText As String = ""
        If selectedRange Is Nothing Then
            NoSelectedCells = True
        End If
        Dim DoRange As Boolean = True
        Dim CBCInstruct As String = $"with '{CellByCellPrefix}' or '{CellByCellPrefix2} if the instruction should be executed cell-by-cell"
        Dim TextInstruct As String = $"use '{TextPrefix}' or '{TextPrefix2}' if the instruction should apply cell-by-cell, but only to text cells"
        Dim BatchInstruct As String = $"use '{BatchPrefix}' if to process a directory of files"
        Dim BubblesInstruct As String = $"use '{BubblesPrefix}' for inserting comments only; without it, you can still ask for values/formulas and comments together"
        Dim PaneInstruct As String = $"use '{PanePrefix}' for using the pane"
        Dim ExtInstruct As String = $"; insert '{ExtTrigger}' or '{ExtTriggerFixed}' (multiple times) for including the text of (a) file(s) (txt, docx, pdf), {ExtDirTrigger} for a directory of text files, or '{ExtWSTrigger}' to add more worksheet(s)"
        Dim AddonInstruct As String = $"; add '{ColorTrigger}' to check for colorcodes"
        Dim NoFormulasInstruct As String = $"; add '{NoFormulasTrigger}' to ignore formulas in cells"
        Dim ObjectInstruct As String = $"; add '{ObjectTrigger}'/'{ObjectTrigger2}' for adding a file object"
        Dim FileObject As String = ""
        Dim InsertWS As String = ""
        If UseSecondAPI Then
            If Not String.IsNullOrWhiteSpace(INI_APICall_Object_2) Then
                AddonInstruct += ObjectInstruct.Replace("; add", ",")
                DoFileObject = True
            End If
        Else
            If Not String.IsNullOrWhiteSpace(INI_APICall_Object) Then
                AddonInstruct += ObjectInstruct.Replace("; add", ",")
                DoFileObject = True
            End If
        End If

        AddonInstruct += NoFormulasInstruct.Replace("; add", ",")

        GoTo SkipPromptWin

        Dim PromptLibInstruct As String = ""
        If INI_PromptLib Then
            PromptLibInstruct = " or press 'OK' or '/' for the prompt library"
        End If
        If DefaultPrefix.Trim() <> "" Then
            DefaultPrefixText = $" (default prefix: '{DefaultPrefix}')"
        End If
        Dim OptionalButtons As System.Tuple(Of String, String, String)() = {
            System.Tuple.Create("OK, use pane", $"Use this to automatically insert '{PanePrefix}' as a prefix.", PanePrefix)
        }

        Dim InsertButtons As System.Tuple(Of String, String, String)() = {
                            System.Tuple.Create("📄", "Include document {doc}", "{doc}"),
                            System.Tuple.Create("📁", "Include directory {dir}", "{dir}"),
                            Tuple.Create("📊", "Include other open worksheet (addws)", "(addws)"),
                            System.Tuple.Create("📎", "Include file object (file)", "(file)")
                        }

        If Not NoSelectedCells Then
            OtherPrompt = Trim(SLib.ShowCustomInputBox($"Please provide the prompt you wish to execute on the selected cells (start {CBCInstruct}; {TextInstruct}; {PaneInstruct}; {BatchInstruct}; {BubblesInstruct})" & PromptLibInstruct & PureInstruct & ExtInstruct & AddonInstruct & LastPromptInstruct & DefaultPrefixText & ":", $"{AN} Freestyle (using " & If(UseSecondAPI, INI_Model_2, INI_Model) & ")", False, "", My.Settings.LastPrompt, OptionalButtons, InsertButtons, _context))
        Else
            OtherPrompt = Trim(SLib.ShowCustomInputBox($"Please provide the prompt you wish to execute {PromptLibInstruct} (the result will be shown to you before inserting anything into your worksheet); {PaneInstruct}{BatchInstruct}{PureInstruct}{ExtInstruct}{AddonInstruct}{LastPromptInstruct}{DefaultPrefixText}:", $"{AN} Freestyle (using " & If(UseSecondAPI, INI_Model_2, INI_Model) & ")", False, "", My.Settings.LastPrompt, OptionalButtons, InsertButtons, _context))
            DoRange = True
        End If

SkipPromptWin:

        ' =============================================================
        ' FREESTYLE PROMPT UI
        ' =============================================================

        Dim freestylePromptUiState As String = ""

        Try
            freestylePromptUiState = CStr(My.Settings.Item("FreestylePromptUiState"))
        Catch
            freestylePromptUiState = ""
        End Try

        Dim promptOptions As New SLib.FreestylePromptOptions() With {
            .Title = $"{AN} Freestyle",
            .Heading = "What would you like Red Ink to do?",
            .ModeCaption = "Processing",
            .ModelText = If(UseSecondAPI, INI_Model_2, INI_Model),
            .ContextStatusText = If(NoSelectedCells, "No cells selected", "The current Excel selection will be used as context"),
            .LastPrompt = My.Settings.LastPrompt,
            .PromptLibraryEnabled = INI_PromptLib,
            .Context = _context,
            .CallerId = "excel.freestyle",
            .PersistedState = freestylePromptUiState,
            .RestorePersistedState = False
        }

        ' =============================================================
        ' PROCESSING / OUTPUT PREFIXES
        '
        ' Covers:
        ' CellByCellPrefix
        ' CellByCellPrefix2
        ' TextPrefix
        ' TextPrefix2
        ' BubblesPrefix
        ' PanePrefix
        ' BatchPrefix
        ' PurePrefix
        ' =============================================================

        Dim modeDefault As New SLib.FreestylePromptMode() With {
            .Id = "default",
            .Text = "Configured default",
            .Prefix = System.String.Empty,
            .ManualSyntax = DefaultPrefix,
            .IsDefault = True
        }

        promptOptions.Modes.Add(modeDefault)


        Dim modeCellByCell As New SLib.FreestylePromptMode() With {
            .Id = "cell-by-cell",
            .Text = "Cell by cell",
            .Prefix = CellByCellPrefix,
            .ManualSyntax = CellByCellPrefix & " / " & CellByCellPrefix2,
            .IsAvailable = Not NoSelectedCells,
            .UnavailableReason = "Cell-by-cell processing requires a selected Excel range."
        }

        modeCellByCell.Prefixes.Add(CellByCellPrefix)
        modeCellByCell.Prefixes.Add(CellByCellPrefix2)

        promptOptions.Modes.Add(modeCellByCell)


        Dim modeTextOnly As New SLib.FreestylePromptMode() With {
            .Id = "text-only",
            .Text = "Text cells only",
            .Prefix = TextPrefix,
            .ManualSyntax = TextPrefix & " / " & TextPrefix2,
            .IsAvailable = Not NoSelectedCells,
            .UnavailableReason = "Text-cell processing requires a selected Excel range."
        }

        modeTextOnly.Prefixes.Add(TextPrefix)
        modeTextOnly.Prefixes.Add(TextPrefix2)

        promptOptions.Modes.Add(modeTextOnly)


        Dim modeComments As New SLib.FreestylePromptMode() With {
            .Id = "comments",
            .Text = "Comments only",
            .Prefix = BubblesPrefix,
            .ManualSyntax = BubblesPrefix,
            .IsAvailable = Not NoSelectedCells,
            .UnavailableReason = "Comment output requires a selected Excel range."
        }

        modeComments.Prefixes.Add(BubblesPrefix)

        promptOptions.Modes.Add(modeComments)


        Dim modePane As New SLib.FreestylePromptMode() With {
            .Id = "pane",
            .Text = "Side pane",
            .Prefix = PanePrefix,
            .ManualSyntax = PanePrefix
        }

        modePane.Prefixes.Add(PanePrefix)

        promptOptions.Modes.Add(modePane)


        Dim modeBatch As New SLib.FreestylePromptMode() With {
            .Id = "batch",
            .Text = "Batch files",
            .Prefix = BatchPrefix,
            .ManualSyntax = BatchPrefix
        }

        modeBatch.Prefixes.Add(BatchPrefix)

        promptOptions.Modes.Add(modeBatch)


        Dim modePure As New SLib.FreestylePromptMode() With {
            .Id = "pure",
            .Text = "Direct prompt",
            .Prefix = PurePrefix,
            .ManualSyntax = PurePrefix
        }

        modePure.Prefixes.Add(PurePrefix)

        promptOptions.Modes.Add(modePure)

        promptOptions.QuickButtons.Add(
            New SLib.FreestylePromptQuickButton() With {
                .Id = "quick-pane",
                .Text = "Run (pane)",
                .Description = "Run immediately with the side-pane prefix.",
                .Prefix = PanePrefix
            })

        ' =============================================================
        ' ADD CONTEXT
        '
        ' Covers:
        ' ExtTrigger
        ' ExtTriggerFixed
        ' ExtDirTrigger
        ' ExtWSTrigger
        ' ObjectTrigger
        ' ObjectTrigger2
        ' =============================================================

        promptOptions.InsertOptions.Add(
            New SLib.FreestylePromptInsertOption() With {
                .Id = "file",
                .Text = "Doc",
                .Description = "Select and include the text of a document (" & ExtTrigger & ").",
                .InsertText = ExtTrigger
            })

        If Not System.String.IsNullOrWhiteSpace(ExtTriggerFixed) AndAlso ExtTriggerFixed.IndexOf("[path]", System.StringComparison.OrdinalIgnoreCase) >= 0 Then

            promptOptions.InsertOptions.Add(
                New SLib.FreestylePromptInsertOption() With {
                    .Id = "file-path",
                    .Text = "Doc path",
                    .Description = "Enter a specific document path and insert the completed path command into the prompt.",
                    .RequiresValue = True,
                    .ValuePrompt = "Enter the full path of the document to include:",
                    .ValueTitle = $"{AN} Freestyle - Doc path",
                    .ValueTemplate = ExtTriggerFixed,
                    .ValuePlaceholder = "[path]"
                })

        End If

        promptOptions.InsertOptions.Add(
            New SLib.FreestylePromptInsertOption() With {
                .Id = "folder",
                .Text = "Folder",
                .Description = "Select and include files from a directory (" & ExtDirTrigger & ").",
                .InsertText = ExtDirTrigger
            })

        promptOptions.InsertOptions.Add(
            New SLib.FreestylePromptInsertOption() With {
                .Id = "worksheet",
                .Text = "Other worksheets",
                .Description = "Include one or more other open worksheets (" & ExtWSTrigger & ").",
                .InsertText = ExtWSTrigger
            })

        If DoFileObject Then

            promptOptions.InsertOptions.Add(
                New SLib.FreestylePromptInsertOption() With {
                    .Id = "file-object",
                    .Text = "File object",
                    .Description = "Attach a file as an LLM object (" & ObjectTrigger & ").",
                    .InsertText = ObjectTrigger,
                    .MaxOccurrences = 1
                })

            promptOptions.InsertOptions.Add(
                New SLib.FreestylePromptInsertOption() With {
                    .Id = "clipboard-object",
                    .Text = "Clipboard object",
                    .Description = "Use the clipboard as an LLM object (" & ObjectTrigger2 & ").",
                    .InsertText = ObjectTrigger2
                })

        End If

        ' =============================================================
        ' MORE OPTIONS
        '
        ' Covers:
        ' ColorTrigger
        ' NoFormulasTrigger
        ' =============================================================

        Dim processingSection As New SLib.FreestylePromptSection() With {
            .Id = "processing",
            .Caption = "Processing"
        }

        processingSection.Options.Add(
            New SLib.FreestylePromptToggleOption() With {
                .Id = "colour-codes",
                .Text = "Check cell colours",
                .Trigger = ColorTrigger,
                .ManualSyntax = ColorTrigger,
                .Description = "Include cell colour information when processing the selection."
            })

        processingSection.Options.Add(
            New SLib.FreestylePromptToggleOption() With {
                .Id = "no-formulas",
                .Text = "Ignore formulas",
                .Trigger = NoFormulasTrigger,
                .ManualSyntax = NoFormulasTrigger,
                .Description = "Do not process formulas in cells."
            })

        promptOptions.Sections.Add(processingSection)

        ' =============================================================
        ' SHOW
        ' =============================================================

        Dim promptResult As SLib.FreestylePromptResult = SLib.ShowFreestylePromptForm(promptOptions)

        If promptResult Is Nothing OrElse Not promptResult.Accepted Then
            Return False
        End If

        Try
            My.Settings.Item("FreestylePromptUiState") = promptResult.PersistedState
            My.Settings.Save()
        Catch
            ' non-critical
        End Try

        OtherPrompt = SLib.ComposeFreestylePrompt(promptResult).Trim()


        If String.IsNullOrEmpty(OtherPrompt) And OtherPrompt <> "ESC" And INI_PromptLib Then
            Dim promptlibresult As (String, Boolean, Boolean, Boolean)
            promptlibresult = ShowPromptSelector(INI_PromptLibPath, INI_PromptLibPathLocal, Nothing, Nothing)
            OtherPrompt = promptlibresult.Item1
            DoClipboard = promptlibresult.Item4
            If OtherPrompt = "" Then
                Return False
            End If
        Else
            If String.IsNullOrEmpty(OtherPrompt) Or OtherPrompt = "ESC" Then Return False
        End If
        ' Check if OtherPrompt starts with a word ending with a colon
        If Not String.IsNullOrWhiteSpace(OtherPrompt) Then
            Dim firstWord As String = OtherPrompt.Split({" "c}, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault()
            If firstWord IsNot Nothing AndAlso Not firstWord.EndsWith(":"c) Then
                Dim prefix As String = DefaultPrefix.Trim()
                ' Ensure prefix ends with colon and space
                If prefix <> "" AndAlso Not prefix.EndsWith(":"c) Then
                    prefix &= ":"
                End If
                OtherPrompt = prefix & " " & OtherPrompt.Trim()
                OtherPrompt = OtherPrompt.Trim()
            End If
        End If
        My.Settings.LastPrompt = OtherPrompt
        My.Settings.Save()
        If Not SharedMethods.ProcessParameterPlaceholders(OtherPrompt) Then
            ShowCustomMessageBox("Freestyle canceled.", $"{AN} Freestyle")
            Return False
        End If
        If OtherPrompt.StartsWith(CellByCellPrefix, StringComparison.OrdinalIgnoreCase) And DoFormulas Then
            OtherPrompt = OtherPrompt.Substring(CellByCellPrefix.Length).Trim()
            DoRange = False
        End If
        If OtherPrompt.StartsWith(CellByCellPrefix2, StringComparison.OrdinalIgnoreCase) And DoFormulas Then
            OtherPrompt = OtherPrompt.Substring(CellByCellPrefix2.Length).Trim()
            DoRange = False
        End If
        If OtherPrompt.StartsWith(TextPrefix, StringComparison.OrdinalIgnoreCase) And DoFormulas Then
            OtherPrompt = OtherPrompt.Substring(TextPrefix.Length).Trim()
            DoRange = False
            DoFormulas = False
        End If
        If OtherPrompt.StartsWith(TextPrefix2, StringComparison.OrdinalIgnoreCase) And DoFormulas Then
            OtherPrompt = OtherPrompt.Substring(TextPrefix2.Length).Trim()
            DoRange = False
            DoFormulas = False
        End If
        If OtherPrompt.StartsWith(BubblesPrefix, StringComparison.OrdinalIgnoreCase) And selectedRange IsNot Nothing Then
            OtherPrompt = OtherPrompt.Substring(BubblesPrefix.Length).Trim()
            DoBubbles = True
            DoRange = True
        End If
        If OtherPrompt.StartsWith(PanePrefix, StringComparison.OrdinalIgnoreCase) And DoRange Then
            OtherPrompt = OtherPrompt.Substring(PanePrefix.Length).Trim()
            DoPane = True
            DoRange = True
        End If
        If OtherPrompt.StartsWith(BatchPrefix, StringComparison.OrdinalIgnoreCase) Then
            OtherPrompt = OtherPrompt.Substring(BatchPrefix.Length).Trim()
            DoPane = False
            DoRange = True
            DoBatch = True
            Try
                ' --- Excel context (as requested) ---
                Dim activeCell As Microsoft.Office.Interop.Excel.Range = Application.ActiveCell
                Dim ws As Microsoft.Office.Interop.Excel.Worksheet = CType(Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                Dim currentRow As System.Int32 = 1
                If activeCell IsNot Nothing Then
                    currentRow = System.Convert.ToInt32(activeCell.Row, System.Globalization.CultureInfo.InvariantCulture)
                End If
                Dim maxRow As System.Int32 = System.Convert.ToInt32(ws.Rows.Count, System.Globalization.CultureInfo.InvariantCulture)
                ' --- Loop until user provides a valid line number or cancels/ESC ---
                Dim lineNumberAnswer As System.String = System.String.Empty
                Do
                    lineNumberAnswer = ShowCustomInputBox(
                        "Please provide the starting line number at which the results should be inserted (1 for first line, 2 for second line etc.):", $"{AN} Freestyle Batch",
                        True,
                        currentRow.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    )
                    If lineNumberAnswer Is Nothing Then
                        Return Nothing
                    End If
                    lineNumberAnswer = lineNumberAnswer.Trim()
                    If lineNumberAnswer.Length = 0 Then
                        Return Nothing
                    End If
                    If System.String.Compare(lineNumberAnswer, "ESC", System.StringComparison.OrdinalIgnoreCase) = 0 Then
                        Return Nothing
                    End If
                    Dim parsedRow As System.Int32
                    If Not System.Int32.TryParse(lineNumberAnswer, parsedRow) OrElse parsedRow < 1 OrElse parsedRow > maxRow Then
                        ShowCustomMessageBox("The provided value is not a valid line number between 1 and " &
                                             maxRow.ToString(System.Globalization.CultureInfo.InvariantCulture) & ".")
                    Else
                        LineNumber = parsedRow
                        Exit Do
                    End If
                Loop
                ' --- Directory picker (browse or type) ---
                Dim initialPath As System.String
                If (Application.ActiveWorkbook IsNot Nothing) AndAlso
                   (Application.ActiveWorkbook.Path IsNot Nothing) AndAlso
                   (Application.ActiveWorkbook.Path.Length > 0) Then
                    initialPath = Application.ActiveWorkbook.Path
                Else
                    initialPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments)
                End If
                Using dlg As New System.Windows.Forms.FolderBrowserDialog()
                    dlg.Description = "Select the directory that contains the batch text files."
                    dlg.ShowNewFolderButton = False
                    dlg.SelectedPath = initialPath
                    Dim result As System.Windows.Forms.DialogResult = dlg.ShowDialog()
                    If result <> System.Windows.Forms.DialogResult.OK Then
                        Return Nothing
                    End If
                    Dim selectedPath As System.String = dlg.SelectedPath
                    If System.String.IsNullOrWhiteSpace(selectedPath) OrElse Not System.IO.Directory.Exists(selectedPath) Then
                        ShowCustomMessageBox("No directory was selected or it does not exist.")
                        Return Nothing
                    End If
                    Dim hasAny As System.Boolean = False
                    ' Enumerate files and check extensions
                    For Each filePath As System.String In System.IO.Directory.EnumerateFiles(selectedPath, "*.*", System.IO.SearchOption.TopDirectoryOnly)
                        Dim ext As System.String = System.IO.Path.GetExtension(filePath)
                        If allowedExtensions.Contains(ext) Then
                            hasAny = True
                            Exit For
                        End If
                    Next
                    ' Handle case when no allowed files exist
                    If Not hasAny Then
                        ShowCustomMessageBox(
                            "The selected directory does not contain any files of the expected types: " &
                            System.String.Join(", ", allowedExtensions) & "."
                        )
                        Return Nothing
                    End If
                    BatchPath = selectedPath
                End Using
            Catch ex As System.Exception
                ShowCustomMessageBox("GetLineNumber and Path resulted in an Error: " & ex.Message)
                Return Nothing
            End Try
        End If
        If DoFileObject AndAlso OtherPrompt.IndexOf(ObjectTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
            OtherPrompt = OtherPrompt.Replace(ObjectTrigger, "(a file object follows)").Trim()
        ElseIf DoFileObject AndAlso OtherPrompt.IndexOf(ObjectTrigger2, StringComparison.OrdinalIgnoreCase) >= 0 Then
            OtherPrompt = OtherPrompt.Replace(ObjectTrigger2, "(a file object follows)").Trim()
            DoFileObjectClip = True
        Else
            DoFileObject = False
        End If
        If selectedRange IsNot Nothing Then
            selectedRange.Select()
        End If
        If Not String.IsNullOrEmpty(OtherPrompt) And OtherPrompt.IndexOf(ColorTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
            DoColor = True
            OtherPrompt = Regex.Replace(OtherPrompt, Regex.Escape(ColorTrigger), "", RegexOptions.IgnoreCase)
        End If

        If Not String.IsNullOrEmpty(OtherPrompt) And OtherPrompt.IndexOf(NoFormulasTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
            DoFormulas = False
            OtherPrompt = Regex.Replace(OtherPrompt, Regex.Escape(NoFormulasTrigger), "", RegexOptions.IgnoreCase)
        End If

        ' === External file/directory embedding ===
        ' Handles {doc}, {dir}, and {path} triggers with unified document numbering

        Dim fileResult = Await ProcessExternalFileTriggers(OtherPrompt)
        If Not fileResult.Success Then
            Return False
        End If
        OtherPrompt = fileResult.ModifiedPrompt

        Debug.WriteLine(OtherPrompt)

        ' == worksheet embedding ===

        If Not String.IsNullOrEmpty(OtherPrompt) And OtherPrompt.IndexOf(ExtWSTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
            If Not DoRange Then
                ShowCustomMessageBox($"{ExtWSTrigger} cannot be combined with cell by cell processing - exiting.")
                Return False
            End If
            InsertWS = GatherSelectedWorksheets(True)
            Debug.WriteLine($"GatherSelectedWorksheets returned: {Left(InsertWS, 3000)}")
            If String.IsNullOrWhiteSpace(InsertWS) Then
                ShowCustomMessageBox("No content was found or an error occurred in gathering the additional worksheet(s) - exiting.")
                Return False
            ElseIf InsertWS.StartsWith("ERROR", StringComparison.OrdinalIgnoreCase) Then
                ShowCustomMessageBox($"An error occured gathering the additional worksheet(s) ({InsertWS.Substring(6).Trim()}) - exiting.")
                Return False
            ElseIf InsertWS.StartsWith("NONE", StringComparison.OrdinalIgnoreCase) Then
                ShowCustomMessageBox($"There are no other worksheets to add - exiting.")
                Return False
            End If
            OtherPrompt = Regex.Replace(OtherPrompt, Regex.Escape(ExtWSTrigger), "", RegexOptions.IgnoreCase)
        End If
        If DoFileObject Then
            If DoFileObjectClip Then
                FileObject = "clipboard"
            Else
                DragDropFormLabel = "All file types that are supported by your LLM."
                DragDropFormFilter = "Supported Files|*.*"
                FileObject = GetFileName()
                DragDropFormLabel = ""
                DragDropFormFilter = ""
                If String.IsNullOrWhiteSpace(FileObject) Then
                    ShowCustomMessageBox("No file object has been selected - will abort. You can try again (use Ctrl-P to re-insert your prompt).")
                    Return False
                End If
            End If
        End If
        If OtherPrompt.StartsWith(PurePrefix, StringComparison.OrdinalIgnoreCase) Then
            OtherPrompt = OtherPrompt.Substring(PurePrefix.Length).Trim()
            Dim result As Boolean = Await ProcessSelectedRange(OtherPrompt, True, DoRange, DoFormulas, DoBubbles, False, UseSecondAPI, 0, True, DoColor, DoPane, FileObject, InsertWS)
        Else
            If Not NoSelectedCells Then
                If DoRange Then
                    FormulaInstruction = INI_FormulaInstruction
                    Dim result As Boolean = Await ProcessSelectedRange(SP_RangeOfCells, True, DoRange, DoFormulas, DoBubbles, False, UseSecondAPI, 0, True, DoColor, DoPane, FileObject, InsertWS, BatchPath)
                Else
                    Dim result As Boolean = Await ProcessSelectedRange(SP_FreestyleText, True, DoRange, DoFormulas, DoBubbles, False, UseSecondAPI, 0, True, DoColor, DoPane, FileObject, InsertWS)
                End If
            Else
                Dim result As Boolean = Await ProcessSelectedRange(SP_RangeOfCells, True, DoRange, DoFormulas, DoBubbles, False, UseSecondAPI, 0, True, DoColor, DoPane, FileObject, InsertWS, BatchPath)
            End If
        End If
    End Function




    Private _win As HelpMeInky = Nothing

    ''' <summary>
    ''' Lazy-instantiates HelpMeInky window (HelpMeInky) and displays it using ShowRaised.
    ''' </summary>
    Public Sub HelpMeInky()
        If _win Is Nothing OrElse _win.IsDisposed Then
            _win = New HelpMeInky(
                _context,
                RDV,
                New System.Globalization.CultureInfo(
                    Globals.ThisAddIn.Application.LanguageSettings.LanguageID(Microsoft.Office.Core.MsoAppLanguageID.msoLanguageIDUI)).DisplayName)
        End If
        ' No owner needed
        _win.ShowRaised()
    End Sub

    ''' <summary>
    ''' Builds settings name/description dictionaries, shows settings window, then updates menus via AddContextMenu wrapped in a SplashScreen.
    ''' </summary>
    Public Sub ShowSettings()
        Dim Settings As New Dictionary(Of String, String) From {
            {"SimpleMenuOverride", "Simple menu"},
            {"Temperature", "Temperature of {model}"},
            {"Timeout", "Timeout of {model}"},
            {"Temperature_2", "Temperature of {model2}"},
            {"Timeout_2", "Timeout of {model2}"},
            {"RestrictedModelAccessCode", "Restricted model access code(s)"},
            {"DoubleS", "Convert '" & ChrW(223) & "' to 'ss'"},
            {"NoEmDash", "Convert em to en dash"},
            {"Ignore", "Activate 'Ignore' prompt (for 'prompt injection' protection)"},
            {"PreCorrection", "Additional instruction for prompts"},
            {"PostCorrection", "Prompt to apply after queries"},
            {"Language1", "Default translation language 1"},
            {"Language2", "Default translation language 2"},
            {"PromptLibPath", "Prompt library file"},
            {"PromptLibPathLocal", "Prompt library file (local)"},
            {"DefaultPrefix", "Default prefix to use in 'Freestyle'"},
            {"Location", "Location information to use, e.g., in 'Freestyle'"},
            {"KnowledgeStorePath", "Knowledge store file (central)"},
            {"KnowledgeStorePathLocal", "Knowledge store file (local)"},
            {"FormulaInstruction", "Additional formula instructions:"}
        }
        Dim SettingsTips As New Dictionary(Of String, String) From {
            {"SimpleMenuOverride", "Only show the simple menu features in this add-in (can be defined)"},
            {"Temperature", "The higher, the more creative the LLM will be (0.0-2.0)"},
            {"Timeout", "In milliseconds"},
            {"Temperature_2", "The higher, the more creative the LLM will be (0.0-2.0)"},
            {"Timeout_2", "In milliseconds"},
            {"RestrictedModelAccessCode", "Optional personal access code list for restricted alternate models. Use commas or semicolons to separate multiple codes. A restricted model is only shown when one stored code matches the model's 'RestrictedAccess' value."},
            {"DoubleS", "For Switzerland"},
            {"NoEmDash", "This will convert long dashes typically generated by LLMs but that are not commonly used (thus suggesting that the text has been AI generated)"},
            {"Ignore", "Allow system prompts to use {Ignore} as a placeholder for text to ignore, such as malicious prompt injections; Freestyle and some other commands use {Ignore}; the chatbots have an independent protection"},
            {"PreCorrection", "Add prompting text that will be added to all basic requests (e.g., for special language tasks)"},
            {"PostCorrection", "Add a prompt that will be applied to each result before it is further processed (slow!)"},
            {"Language1", "The language (in English) that will be used for the first quick access button in the ribbon"},
            {"Language2", "The language (in English) that will be used for the second quick access button in the ribbon"},
            {"PromptLibPath", "The filename (including path, support environmental variables) for your prompt library (if any)"},
            {"PromptLibPathLocal", "The filename (including path, support environmental variables) for your local prompt library (if any)"},
            {"DefaultPrefix", "You can define here the default prefix to use within 'Freestyle' if no other prefix is used (will be added authomatically)."},
            {"Location", "Provide location information (e.g., 'We are in Zurich, Switzerland') to be used in 'Freestyle', chatbot and some other prompts that contain {Location} to get more location specific results."},
            {"KnowledgeStorePath", "The file path for the central knowledge store index (supports env variables); used by the (kb) trigger"},
            {"KnowledgeStorePathLocal", "The file path for the local knowledge store index (supports env variables); used by the (kb) trigger"},
            {"FormulaInstruction", "This optional instruction will be added to the Freestyle system prompt for steering the creation of formulas (e.g., to have them created in a local language using local separators)."}
        }
        ShowSettingsWindow(Settings, SettingsTips)
        Dim splash As New SplashScreen("Updating menu following your changes ...")
        splash.Show()
        splash.Refresh()

        ' The Excel context menus (CommandBars popups) are built once at startup and are
        ' structurally static - their menu items never change, only the behaviour of the
        ' OnAction VBA macros, which read the current INI values at click time. Tearing
        ' them down and rebuilding the CommandBars popups after a settings change is
        ' unnecessary and, on the Word side, was the cause of a non-deterministic COM
        ' fail-fast (exit code 0xC0000409). We therefore do NOT rebuild the context menus
        ' after a settings change; we only refresh the ribbon, which is managed WinForms
        ' UI and safe to update.
        Try
            If Globals.Ribbons.Ribbon1 IsNot Nothing Then
                Globals.Ribbons.Ribbon1.ApplyRibbonVisibilityConfiguration()
            End If
        Catch
            ' non-critical
        End Try

        splash.Close()
    End Sub

End Class
