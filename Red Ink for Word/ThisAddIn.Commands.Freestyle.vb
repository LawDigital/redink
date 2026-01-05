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
Imports DocumentFormat.OpenXml.Presentation
Imports DocumentFormat.OpenXml.Wordprocessing
Imports Microsoft.Office.Interop.PowerPoint
Imports Microsoft.Office.Interop.Word
Imports NetOffice.PowerPointApi
Imports SharedLibrary
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn


    ''' <summary>
    ''' Helper class to track file loading results for the Freestyle external file embedding feature.
    ''' </summary>
    Private Class FileLoadingContext
        Public Property GlobalDocumentCounter As Integer = 0
        Public Property LoadedFiles As New List(Of Tuple(Of String, Integer))() ' (path, charCount)
        Public Property FailedFiles As New List(Of String)()
        Public Property IgnoredFilesPerDir As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

        ''' <summary>OCR mode: 0=not set, 1=enable for all, 2=ask individually</summary>
        Public Property DirectoryOCRMode As Integer = 0

        ''' <summary>Tracks the current directory being processed (for including path in output).</summary>
        Public Property CurrentDirectoryPath As String = ""

        ''' <summary>Total number of files expected to be loaded (for determining if filename should be included).</summary>
        Public Property ExpectedFileCount As Integer = 0

        ''' <summary>Maximum files to load from a single directory.</summary>
        Public Const MaxFilesPerDirectory As Integer = 50

        ''' <summary>Whether to recurse into subdirectories when loading a directory.</summary>
        Public Const RecurseDirectories As Boolean = False

        ''' <summary>Ask user confirmation if directory has more than this many files.</summary>
        Public Const ConfirmDirectoryFileCount As Integer = 10

        ''' <summary>Supported file extensions for GetFileContent.</summary>
        Public Shared ReadOnly SupportedExtensions As String() = {
        ".txt", ".rtf", ".doc", ".docx", ".pdf", ".pptx", ".ini", ".csv", ".log",
        ".json", ".xml", ".html", ".htm", ".md", ".vb", ".cs", ".js", ".ts",
        ".py", ".java", ".cpp", ".c", ".h", ".sql", ".yaml", ".yml"
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
    ''' Loads a single file and returns its content wrapped in document tags.
    ''' For individual files or single PDFs, OCR is enabled with AskUser=True.
    ''' When multiple files are loaded, includes filename (and directory path for directory files) in the tags.
    ''' </summary>
    ''' <param name="filePath">The path to the file to load.</param>
    ''' <param name="isWrapped">If True, returns raw content without document tags.</param>
    ''' <param name="ctx">The file loading context for tracking state.</param>
    ''' <param name="isFromDirectory">If True, file is being loaded as part of a directory batch.</param>
    ''' <returns>The file content, optionally wrapped in document tags.</returns>
    Private Async Function LoadSingleFileAsync(filePath As String, isWrapped As Boolean, ctx As FileLoadingContext, Optional isFromDirectory As Boolean = False) As Task(Of String)
        Try
            Dim doOCR As Boolean = True
            Dim askUser As Boolean = True

            ' Determine OCR behavior based on context
            If isFromDirectory Then
                ' File is part of a directory batch - use the directory OCR mode
                If ctx.DirectoryOCRMode = 1 Then
                    ' Enable OCR for all without asking
                    doOCR = True
                    askUser = False
                ElseIf ctx.DirectoryOCRMode = 2 Then
                    ' Ask individually for each file
                    doOCR = True
                    askUser = True
                Else
                    ' OCR mode not set (shouldn't happen for directories with multiple PDFs)
                    doOCR = True
                    askUser = True
                End If
            Else
                ' Individual file or single PDF - always enable OCR with AskUser=True
                doOCR = True
                askUser = True
            End If

            ' Load file content with determined OCR settings
            Dim content As String = Await GetFileContent(
                optionalFilePath:=filePath,
                Silent:=True,
                DoOCR:=doOCR,
                AskUser:=askUser
            )

            If String.IsNullOrWhiteSpace(content) Then
                ctx.FailedFiles.Add(filePath)
                Return ""
            End If

            ctx.GlobalDocumentCounter += 1
            Dim docNum As Integer = ctx.GlobalDocumentCounter
            ctx.LoadedFiles.Add(Tuple.Create(filePath, content.Length))

            If isWrapped Then
                Return content
            Else
                ' Determine if we should include filename/path info
                ' Include if: multiple files expected, or loading from directory
                Dim includeFileInfo As Boolean = (ctx.ExpectedFileCount > 1) OrElse isFromDirectory

                Dim openTag As String
                Dim closeTag As String = "</document" & docNum.ToString() & ">"

                If includeFileInfo Then
                    Dim fileName As String = Path.GetFileName(filePath)
                    If isFromDirectory AndAlso Not String.IsNullOrEmpty(ctx.CurrentDirectoryPath) Then
                        ' Include both directory path and filename for directory files
                        openTag = "<document" & docNum.ToString() & " directory=""" & ctx.CurrentDirectoryPath & """ filename=""" & fileName & """>"
                    Else
                        ' Include just the filename for multiple individual files
                        openTag = "<document" & docNum.ToString() & " filename=""" & fileName & """>"
                    End If
                Else
                    ' Single file - no need to include filename
                    openTag = "<document" & docNum.ToString() & ">"
                End If

                Return openTag & content & closeTag
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
    ''' <returns>Combined content of all files, or "ABORT" if user cancelled, or empty string on error.</returns>
    Private Async Function LoadDirectoryFilesAsync(dirPath As String, isWrapped As Boolean, ctx As FileLoadingContext) As Task(Of String)
        Try
            If Not Directory.Exists(dirPath) Then
                ctx.FailedFiles.Add(dirPath)
                Return ""
            End If

            Dim searchOption As SearchOption = If(FileLoadingContext.RecurseDirectories, SearchOption.AllDirectories, SearchOption.TopDirectoryOnly)
            Dim allFiles As String() = Directory.GetFiles(dirPath, "*.*", searchOption)

            ' Filter to supported extensions
            Dim supportedFiles As New List(Of String)()
            Dim ignoredCount As Integer = 0

            For Each f In allFiles
                Dim ext As String = Path.GetExtension(f).ToLowerInvariant()
                If FileLoadingContext.SupportedExtensions.Contains(ext) Then
                    supportedFiles.Add(f)
                Else
                    ignoredCount += 1
                End If
            Next

            If ignoredCount > 0 Then
                ctx.IgnoredFilesPerDir(dirPath) = ignoredCount
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

            ' Only ask about OCR if there are multiple PDF files in the directory
            If pdfCount > 1 Then
                Dim ocrChoice As Integer = ShowCustomYesNoBox(
                    $"The directory '{dirPath}' contains {pdfCount} PDF files." & vbCrLf & vbCrLf &
                    "Some PDFs may require OCR (optical character recognition) to extract text from scanned documents or images." & vbCrLf & vbCrLf &
                    "How would you like to handle OCR for these PDF files?",
                    "Enable OCR for all PDFs",
                    "Choose individually for each PDF",
                    "Close this window to abort")

                If ocrChoice = 0 Then
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
            ' If pdfCount <= 1, we don't set DirectoryOCRMode, and individual files will use AskUser=True

            ' Set the current directory path for inclusion in document tags
            ctx.CurrentDirectoryPath = dirPath

            ' Update expected file count to include these directory files
            ctx.ExpectedFileCount += supportedFiles.Count

            ' Load all files - each file gets its own document wrapper with incrementing counter
            Dim resultBuilder As New System.Text.StringBuilder()
            For Each filePath In supportedFiles
                ' Always wrap individual files from directory with document tags
                ' Pass isFromDirectory=True to use directory OCR mode for PDFs
                Dim fileContent As String = Await LoadSingleFileAsync(filePath, False, ctx, isFromDirectory:=True)
                If Not String.IsNullOrWhiteSpace(fileContent) Then
                    resultBuilder.Append(fileContent)
                End If
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

        ' Create context for tracking
        Dim ctx As New FileLoadingContext()

        ' Count total expected files/triggers to determine if filenames should be included
        ' This is an estimate - actual count may differ after user selections
        Dim extTriggerCount As Integer = Regex.Matches(prompt, Regex.Escape(ExtTrigger), RegexOptions.IgnoreCase).Count
        Dim dirTriggerCount As Integer = Regex.Matches(prompt, Regex.Escape(ExtDirTrigger), RegexOptions.IgnoreCase).Count
        Dim fixedPathCount As Integer = 0
        If Not String.IsNullOrEmpty(patternFixed) Then
            fixedPathCount = Regex.Matches(prompt, patternFixed, RegexOptions.IgnoreCase).Count
        End If
        ctx.ExpectedFileCount = extTriggerCount + fixedPathCount
        ' Note: dirTriggerCount files are added in LoadDirectoryFilesAsync

        ' NOTE: OCR prompt is handled per-directory when multiple PDFs are found,
        ' or per-file with AskUser=True for individual files/single PDFs

        ' === Process all triggers in positional order ===
        ' We need to find the first trigger of any type, process it, then repeat
        ' This ensures document numbering follows the order in the prompt

        Dim continueProcessing As Boolean = True
        While continueProcessing
            ' Find position of each trigger type
            Dim extIdx As Integer = prompt.IndexOf(ExtTrigger, StringComparison.OrdinalIgnoreCase)
            Dim dirIdx As Integer = prompt.IndexOf(ExtDirTrigger, StringComparison.OrdinalIgnoreCase)
            Dim fixedIdx As Integer = -1
            Dim fixedMatch As Match = Nothing

            If Not String.IsNullOrEmpty(patternFixed) Then
                fixedMatch = Regex.Match(prompt, patternFixed, RegexOptions.IgnoreCase Or RegexOptions.Singleline)
                If fixedMatch.Success Then
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
                        ' Individual file - isFromDirectory=False means OCR with AskUser=True
                        replacementText = Await LoadSingleFileAsync(selectedFile, isWrapped, ctx, isFromDirectory:=False)
                    Else
                        Dim answer As Integer = ShowCustomYesNoBox(
                        "No file selected. Do you want to continue or abort?",
                        "Continue", "Abort")
                        If answer <> 1 Then Return (False, prompt)
                    End If

                    prompt = prompt.Substring(0, extIdx) & replacementText & prompt.Substring(extIdx + ExtTrigger.Length)

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
                        replacementText = Await LoadDirectoryFilesAsync(selectedDir, isWrapped, ctx)

                        If replacementText = "ABORT" Then
                            Return (False, prompt)
                        End If
                    Else
                        Dim answer As Integer = ShowCustomYesNoBox(
                        "No directory selected. Do you want to continue or abort?",
                        "Continue", "Abort")
                        If answer <> 1 Then Return (False, prompt)
                    End If

                    prompt = prompt.Substring(0, dirIdx) & replacementText & prompt.Substring(dirIdx + ExtDirTrigger.Length)

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
                        ' Individual file - isFromDirectory=False means OCR with AskUser=True
                        replacementText = Await LoadSingleFileAsync(candidatePath, isWrapped, ctx, isFromDirectory:=False)
                    ElseIf Directory.Exists(candidatePath) Then
                        ' Directory - will prompt for OCR if multiple PDFs found
                        replacementText = Await LoadDirectoryFilesAsync(candidatePath, isWrapped, ctx)
                        If replacementText = "ABORT" Then
                            Return (False, prompt)
                        End If
                    Else
                        ctx.FailedFiles.Add(candidatePath)
                        replacementText = ""
                    End If

                    prompt = prompt.Substring(0, fixedMatch.Index) & replacementText & prompt.Substring(fixedMatch.Index + fixedMatch.Length)
            End Select
        End While

        ' === Show summary and confirm before proceeding ===
        If ctx.LoadedFiles.Count > 0 OrElse ctx.FailedFiles.Count > 0 OrElse ctx.IgnoredFilesPerDir.Count > 0 Then
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

            If ctx.FailedFiles.Count > 0 Then
                summary.AppendLine($"Failed to load ({ctx.FailedFiles.Count} items):")
                For Each f In ctx.FailedFiles
                    summary.AppendLine($"  • {f}")
                Next
                summary.AppendLine("")
            End If

            If ctx.IgnoredFilesPerDir.Count > 0 Then
                summary.AppendLine("Ignored unsupported files:")
                For Each kvp In ctx.IgnoredFilesPerDir
                    summary.AppendLine($"  • {kvp.Key}: {kvp.Value} file(s)")
                Next
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
        If INILoadFail() Then Return
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
        If INILoadFail() Then Return

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
        If INILoadFail() Then Return

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
        If INILoadFail() Then Return
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
            Dim DoMyStyle As Boolean = False
            Dim DoMultiModel As Boolean = True
            Dim DoBubblesExtract As Boolean = False
            Dim DoPushback As Boolean = False

            ' Build instruction strings for user guidance
            Dim MarkupInstruct As String = $"start With '{MarkupPrefixAll}' for markups"
            Dim InplaceInstruct As String = $"with '{InPlacePrefix}'/'{AddPrefix} for replacing/adding to the selection"
            Dim BubblesInstruct As String = $"with '{BubblesPrefix}' for having your text commented"
            Dim PushbackInstruct As String = $"with '{PushbackPrefix}' for responding to comments only"
            Dim SlidesInstruct As String = $"with '{SlidesPrefix}' for adding to a Powerpoint file"
            Dim ClipboardInstruct As String = $"with '{ClipboardPrefix}', '{NewdocPrefix}' or '{PanePrefix}' for separate output"
            Dim PromptLibInstruct As String = If(INI_PromptLib, " or press 'OK' for the prompt library", "")
            Dim ExtInstruct As String = $"; include '{ExtTrigger}' or '{ExtTriggerFixed}' (multiple times) for including the text of (a) file(s) (txt, docx, pdf), {ExtDirTrigger} for a directory of text files, or '{AddDocTrigger}' for an open Word doc"
            Dim TPMarkupInstruct As String = $"; add '{TPMarkupTriggerInstruct}' if revisions [of user] should be pointed out to the LLM"
            Dim NoFormatInstruct As String = $"; add '{NoFormatTrigger2}'/'{KFTrigger2}'/'{KPFTrigger2}/{SameAsReplaceTrigger}' for overriding formatting defaults"
            Dim AllInstruct As String = $"; add '{AllTrigger}' to select all"
            Dim MyStyleInstruct As String = $"; add '{MyStyleTrigger}' to apply your personal style"
            Dim LibInstruct As String = $"; add '{LibTrigger}' for library search"
            Dim NetInstruct As String = $"; add '{NetTrigger}' for internet search"
            Dim PureInstruct As String = $"; use '{PurePrefix}' for direct prompting"
            Dim ChunkInstruct As String = $"; add '{ChunkTrigger}' for iterating through the text"
            Dim BubblesExtractInstruct As String = $"; add '{BubblesExtractTrigger}' for including bubble comments"
            Dim ObjectInstruct As String = $"; add '{ObjectTrigger}'/'{ObjectTrigger2}' for adding a file object"
            Dim MultiModelInstruct As String = $"; add '{MultiModelTrigger}' for multiple models"
            Dim ToolSelectionInstruct As String = $"; add '{ToolSelectionTrigger}' to permit {ToolFriendlyName.ToLower} selection"
            Dim LastPromptInstruct As String = If(String.IsNullOrWhiteSpace(My.Settings.LastPrompt), "", "; Ctrl-P for your last prompt")
            Dim FileObject As String = ""
            Dim SlideDeck As String = ""

            Dim DefaultPrefix As String = INI_DefaultPrefix
            Dim DefaultPrefixText As String = ""

            Dim application As Word.Application = Globals.ThisAddIn.Application
            Dim selection As Microsoft.Office.Interop.Word.Selection = application.Selection

            ' Check if no text is selected (insertion point only)
            If selection.Type = WdSelectionType.wdSelectionIP Then NoText = True

            ' Check if the selected model supports tooling (can call tools)
            Dim modelSupportsTool As Boolean = False
            If UseSecondAPI Then
                ' Check via SharedMethods - based on APICall_ToolInstructions
                modelSupportsTool = Not String.IsNullOrWhiteSpace(INI_APICall_ToolInstructions_2) OrElse
                                    SharedMethods.ModelSupportsTooling(LastFreestyleModelConfig)
            End If

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
                If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) And modelSupportsTool Then
                    AddOnInstruct += ToolSelectionInstruct.Replace("; add", ", ")
                End If

            Else
                If Not String.IsNullOrWhiteSpace(INI_APICall_Object) Then
                    AddOnInstruct += ObjectInstruct.Replace("; add", ",")
                    DoFileObject = True
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

            ' Prompt user for input if not provided via LastPrompt parameter
            If LastPrompt.Trim() = "" Then
                If Not NoText Then
                    ' Offer optional buttons for common prefix shortcuts when text is selected
                    Dim OptionalButtons As System.Tuple(Of String, String, String)() = {
                            System.Tuple.Create("OK, use window", $"Use this to automatically insert '{ClipboardPrefix}' as a prefix.", ClipboardPrefix),
                            System.Tuple.Create("OK, use pane", $"Use this to automatically insert '{PanePrefix}' as a prefix.", PanePrefix),
                            System.Tuple.Create("OK, do a markup", $"Use this to automatically insert '{MarkupPrefixDiff}' as a prefix.", MarkupPrefixDiff)
                        }

                    OtherPrompt = SLib.ShowCustomInputBox($"Please provide the prompt you wish to execute on the selected text ({MarkupInstruct}, {ClipboardInstruct}, {InplaceInstruct}, {BubblesInstruct}, {PushbackInstruct} or {SlidesInstruct}){PromptLibInstruct}{ExtInstruct}{AddOnInstruct}{PureInstruct}{LastPromptInstruct}{DefaultPrefixText}:", $"{AN} Freestyle (using " & If(UseSecondAPI, INI_Model_2, INI_Model) & ")", False, "", My.Settings.LastPrompt, OptionalButtons).Trim()
                Else
                    ' Offer limited optional buttons when no text is selected
                    Dim OptionalButtons As System.Tuple(Of String, String, String)() = {
                            System.Tuple.Create("OK, use window", $"Use this to automatically insert '{ClipboardPrefix}' as a prefix.", ClipboardPrefix),
                            System.Tuple.Create("OK, use pane", $"Use this to automatically insert '{PanePrefix}' as a prefix.", PanePrefix)
                        }
                    OtherPrompt = SLib.ShowCustomInputBox($"Please provide the prompt you wish to execute ({ClipboardInstruct} or {SlidesInstruct}){PromptLibInstruct}{ExtInstruct}{AddOnInstruct}{PureInstruct}{LastPromptInstruct}{DefaultPrefixText}:", $"{AN} Freestyle (using " & If(UseSecondAPI, INI_Model_2, INI_Model) & ")", False, "", My.Settings.LastPrompt, OptionalButtons).Trim()
                End If
            Else
                OtherPrompt = LastPrompt
            End If

            Debug.WriteLine($"OtherPrompt: '{OtherPrompt}'")

            SelectedText = ""

            ' === Special utility commands (executed when text is selected) ===

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
                AddItem("clientname", "Copy and show this PC's client identifier (used for UpdateClients).")

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

                ' CLIPBOARD / INSERTION
                AddItem("insertclipboard", "Insert clipboard content at the cursor position.")
                AddItem("insertclip", "Insert clipboard content at the cursor position.")

                ' TOOLS / SOURCES
                AddItem("setsources", "Select sources/tools available for tooling-capable models (session scope).")

                ' PRIVACY / TRANSFORMS
                AddItem("anonymize", "Anonymize/redact the current selection (no LLM call).")
                AddItem("convertmarkdown", "Convert Markdown in the selected text to Word formatting.")

                ' AUDIO / SPEECH
                AddItem("speech", "Start speech transcription (Transcriptor).")
                AddItem("read", "Create audio (TTS) from the selected text.")
                AddItem("readlocal", "Read the selected text using local TTS (no cloud call).")
                AddItem("voices", "Select a single cloud TTS voice.")
                AddItem("voices2", "Select multiple cloud TTS voices.")
                AddItem("voiceslocal", "Select the local TTS voice.")
                AddItem("createpodcast", "Create a podcast from the selected text.")
                AddItem("readpodcast", "Play/read a podcast based on the current selection.")

                ' DOCUMENT / CLAUSES
                AddItem("doccheck", "Run the document check.")
                AddItem("learndocstyle", "Extract paragraph styles (learn document style).")
                AddItem("applydocstyle", "Apply a style template.")
                AddItem("findclause", "Search for a clause in the clause library/database.")
                AddItem("addclause", "Add a clause to the clause library/database.")

                ' WEB AGENT
                AddItem("webagentcreator", "Create/modify web agent scripts.")
                AddItem("webagent", "Run the web agent (requires configured script paths).")

                ' ANALYSIS
                AddItem("findhiddenprompts", "Scan the document for hidden prompts.")

                Dim chosen As Integer = SLib.SelectValue(items,
                            1,
                            "Select a Freestyle short command (Esc to cancel):",
                            $"{AN} Freestyle - Help"
                        )

                If chosen <= 0 OrElse Not idToCommand.ContainsKey(chosen) Then
                    Return
                End If

                OtherPrompt = idToCommand(chosen)
                ' Continue normal execution: the short-command checks below will now match.
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
                ShowCustomMessageBox("Your client name is '" & GetCurrentClientIdentifier() & "' (also in the clipboard).", AN)
                Return
            End If


            ' Signature Management for Update INI Key Functionality
            If String.Equals(OtherPrompt.Trim(), "iniupdatekeys", StringComparison.OrdinalIgnoreCase) OrElse String.Equals(OtherPrompt.Trim(), "signtool", StringComparison.OrdinalIgnoreCase) Then
                ShowSignatureManagementDialog()
                Return
            End If

            ' Batch Signing for Update INI Key Functionality
            If String.Equals(OtherPrompt.Trim(), "iniupdatebatch", StringComparison.OrdinalIgnoreCase) OrElse String.Equals(OtherPrompt.Trim(), "signbatch", StringComparison.OrdinalIgnoreCase) Then
                ShowBatchSigningDialog()
                Return
            End If

            ' Signature Management for Update INI Key Functionality
            If String.Equals(OtherPrompt.Trim(), "iniupdateignored", StringComparison.OrdinalIgnoreCase) OrElse String.Equals(OtherPrompt.Trim(), "iniupdateignore", StringComparison.OrdinalIgnoreCase) Then
                ShowIgnoredParametersDialog()
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

            ' Reset local configuration to defaults (with confirmation)
            If String.Equals(OtherPrompt.Trim(), "logstat", StringComparison.OrdinalIgnoreCase) Then
                SharedLogger.AnalyzeLogs(_context)
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

            ' {all} trigger: Select entire document
            If OtherPrompt.IndexOf(AllTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(AllTrigger, "").Trim()
                Dim document As Word.Document = application.ActiveDocument
                document.Content.Select()
                NoText = False
            End If

            ' {lib} trigger: Enable library search
            If OtherPrompt.IndexOf(LibTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(LibTrigger, "").Trim()
                DoLib = True
            End If

            ' Track point markup trigger: Enable revision tracking in markup
            If OtherPrompt.IndexOf(TPMarkupTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(TPMarkupTrigger, "").Trim()
                DoTPMarkup = True
            End If

            ' {chunk} trigger: Enable chunked processing (iterate through paragraphs)
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

            ' {object} trigger: Attach file object to LLM request
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
            ElseIf OtherPrompt.StartsWith(InPlacePrefix, StringComparison.OrdinalIgnoreCase) And Not NoText Then
                OtherPrompt = OtherPrompt.Substring(InPlacePrefix.Length).Trim()
                DoInplace = True
            ElseIf OtherPrompt.StartsWith(AddPrefix, StringComparison.OrdinalIgnoreCase) And Not NoText Then
                OtherPrompt = OtherPrompt.Substring(AddPrefix.Length).Trim()
                DoInplace = False
            ElseIf OtherPrompt.StartsWith(AddPrefix2, StringComparison.OrdinalIgnoreCase) And Not NoText Then
                OtherPrompt = OtherPrompt.Substring(AddPrefix2.Length).Trim()
                DoInplace = False
            ElseIf OtherPrompt.StartsWith(MarkupPrefix, StringComparison.OrdinalIgnoreCase) And Not NoText Then
                OtherPrompt = OtherPrompt.Substring(MarkupPrefix.Length).Trim()
                DoMarkup = True
            ElseIf OtherPrompt.StartsWith(MarkupPrefixRegex, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(MarkupPrefixRegex.Length).Trim()
                DoMarkup = True
                MarkupMethod = 4
            ElseIf OtherPrompt.StartsWith(MarkupPrefixWord, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(MarkupPrefixWord.Length).Trim()
                DoMarkup = True
                MarkupMethod = 1
            ElseIf OtherPrompt.StartsWith(MarkupPrefixDiffW, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(MarkupPrefixDiffW.Length).Trim()
                DoMarkup = True
                MarkupMethod = 3
            ElseIf OtherPrompt.StartsWith(MarkupPrefixDiff, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(MarkupPrefixDiff.Length).Trim()
                DoMarkup = True
                MarkupMethod = 2
            ElseIf OtherPrompt.StartsWith(PanePrefix, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(PanePrefix.Length).Trim()
                DoPane = True
                DoClipboard = True
                DoChunks = False
            ElseIf OtherPrompt.StartsWith(PushbackPrefix, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(PushbackPrefix.Length).Trim()
                DoPushback = True
                DoChunks = False
                DoBubblesExtract = True
            ElseIf OtherPrompt.StartsWith(PushbackPrefix2, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(PushbackPrefix2.Length).Trim()
                DoPushback = True
                DoChunks = False
                DoBubblesExtract = True
            End If

            ' {net} trigger: Enable internet search
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
            If modelSupportsTool Then
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

            ' {multimodel} trigger: Prompt for multiple model selection
            SelectedAlternateModels = Nothing
            If UseSecondAPI AndAlso Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) AndAlso OtherPrompt.IndexOf(MultiModelTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                If Not DoMarkup AndAlso Not DoBubbles AndAlso Not DoPushback AndAlso Not DoSlides Then
                    If Not ShowMultipleModelSelection(_context, INI_AlternateModelPath) OrElse SelectedAlternateModels Is Nothing OrElse SelectedAlternateModels.Count = 0 Then
                        Return
                    End If
                Else
                    ShowCustomMessageBox($"The multi-model feature cannot be used together with markup, bubbles or slides - will continue only with the model you already selected.")
                End If
                OtherPrompt = OtherPrompt.Replace(MultiModelTrigger, "").Trim()
            End If

            ' === MyStyle prompt integration ===

            ' {mystyle} trigger: Select and apply personal style prompt
            If Not String.IsNullOrWhiteSpace(INI_MyStylePath) And OtherPrompt.IndexOf(MyStyleTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
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

            ' {adddoc} trigger: Gather content from other open Word documents
            If Not String.IsNullOrEmpty(OtherPrompt) And OtherPrompt.IndexOf(AddDocTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then

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


            ' === External file/directory embedding ===

            ' Handles {doc}, {dir}, and {path} triggers with unified document numbering

            Dim fileResult = Await ProcessExternalFileTriggers(OtherPrompt)
            If Not fileResult.Success Then
                Return
            End If
            OtherPrompt = fileResult.ModifiedPrompt

            Debug.WriteLine(OtherPrompt)


            ' === File object selection (for LLM APIs that support file attachments) ===

            If DoFileObject Then
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
            If NoText AndAlso (DoBubbles Or DoChunks) Then
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
            If NoText AndAlso DoMarkup Then
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
            If Not DoInplace AndAlso DoMarkup Then
                Dim AppendMarkup As Integer = ShowCustomYesNoBox("You have asked for a markup to be created, but according to the configuration, it will not replace your current selection but added to it at the end. Is this really what you want?", "Yes, add markup ", "No, replace text with markup")
                If AppendMarkup = 0 Then
                    Return
                ElseIf AppendMarkup = 2 Then
                    DoInplace = True
                    NoFormatAndFieldSaving = False
                End If
            End If

            ' === System prompt construction based on selected mode ===

            ' Handle pure prefix mode: use prompt directly as system prompt without additional processing
            If OtherPrompt.StartsWith(PurePrefix, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(PurePrefix.Length).Replace("(a file object follows)", "").Replace("(a clipboard object follows)", "").Trim()
                SysPrompt = OtherPrompt
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

            Debug.WriteLine("Freestyle Prompt: " & SysPrompt)

            ' === Execute LLM processing with configured parameters ===

            ' Invoke ProcessSelectedText with all configured options            
            Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SysPrompt), True, DoKeepFormat, DoKeepParaFormat, DoInplace, DoMarkup, MarkupMethod, DoClipboard, DoBubbles, False, UseSecondAPI, KeepFormatCap, DoTPMarkup, TPMarkupName, False, FileObject, DoPane, ChunkSize, NoFormatAndFieldSaving, DoNewDoc, SlideDeck, InsertDocs <> "", DoMyStyle, DoBubblesExtract, DoPushback, selectedToolsForSession)

            ' Restore original model configuration if alternate model was used
            If UseSecondAPI And originalConfigLoaded Then
                RestoreDefaults(_context, originalConfig)
                originalConfigLoaded = False
            End If

        Catch ex As System.Exception
            ' Handle any unexpected errors during freestyle execution
            MessageBox.Show("Error in Freestyle: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub



End Class
