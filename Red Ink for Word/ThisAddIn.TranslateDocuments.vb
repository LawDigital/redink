' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.TranslateDocuments.vb
' Purpose: Translates or corrects Word documents while preserving 100% of formatting by
'          editing only OpenXML text nodes (<w:t>) inside DOCX parts.
'
' Architecture / Key Ideas:
'  - OpenXML Processing: Operates directly on DOCX XML and modifies only <w:t>
'    nodes, preserving styles, runs, fields, layout, and document structure.
'  - Paragraph Grouping: Collects all visible text runs from a paragraph,
'    translates/corrects the paragraph as a unit, then redistributes the result
'    back across the original run boundaries.
'  - Batch Processing (token-safe): Paragraphs are processed in batches to
'    stay within LLM token/character limits. Each batch contains a bounded
'    number of paragraphs (TranslateParagraphsPerBatch) and is further reduced
'    if the combined character count would exceed TranslateMaxCharsPerBatch.
'  - Context Windows: Each batch includes a small window of already-processed
'    preceding paragraphs and unprocessed following paragraphs to help the LLM
'    preserve meaning, terminology, and tone across batch boundaries.
'  - Pure Text to LLM: Only plain, visible text is sent to the LLM—no XML, no
'    formatting codes, and no markup—ensuring maximum formatting preservation.
'  - Correction Mode: When correcting, creates a Word compare document showing
'    all changes between original and corrected versions.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.IO.Compression

Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Word
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ''' <summary>
    ''' Defines the document processing mode.
    ''' </summary>
    Public Enum DocumentProcessMode
        Translate
        Correct
    End Enum

    ''' <summary>
    ''' Number of preceding paragraphs to include as processed context.
    ''' </summary>
    Private Const TranslateContextBefore As Integer = 3

    ''' <summary>
    ''' Number of following paragraphs to include as unprocessed context.
    ''' </summary>
    Private Const TranslateContextAfter As Integer = 2

    ''' <summary>
    ''' Number of paragraphs to process per LLM batch.
    ''' </summary>
    Private Const TranslateParagraphsPerBatch As Integer = 10

    ''' <summary>
    ''' Maximum characters per batch to avoid token limits.
    ''' </summary>
    Private Const TranslateMaxCharsPerBatch As Integer = 15000

    ''' <summary>
    ''' Paragraph count threshold for "large document" warning.
    ''' </summary>
    Private Const TranslateLargeDocThreshold As Integer = 200

    ''' <summary>
    ''' Suffix used for corrected document filenames.
    ''' </summary>
    Private Const CorrectedFileSuffix As String = "_corrected"

    ''' <summary>
    ''' Suffix used for compare document filenames.
    ''' </summary>
    Private Const CompareFileSuffix As String = "_corrected_compare"

    ''' <summary>
    ''' Represents a text run (w:t element) with its content and XML reference.
    ''' </summary>
    Private Class TranslateTextRunInfo
        Public Property TextNode As System.Xml.XmlNode
        Public Property OriginalText As String
    End Class

    ''' <summary>
    ''' Represents a paragraph with its text runs for translation.
    ''' </summary>
    Private Class TranslateParagraphInfo
        Public Property Index As Integer
        Public Property TextRuns As List(Of TranslateTextRunInfo)
        Public Property FullText As String  ' Combined text from all runs (plain text only)
        Public Property TranslatedText As String
        Public Property IsEmpty As Boolean
    End Class

    ''' <summary>
    ''' Entry point for translation: prompts for file/directory and target language, then translates documents.
    ''' </summary>
    Public Async Sub TranslateWordDocuments()
        Await ProcessWordDocuments(DocumentProcessMode.Translate)
    End Sub

    ' --- Correction overrides (optional) ---
    Private _correctPromptOverride As String = Nothing
    Private _correctSuffixOverride As String = Nothing
    Private _useSecondAPI As Boolean = False

    ''' <summary>
    ''' Entry point for correction (default prompt/suffix).
    ''' </summary>
    Public Async Sub CorrectWordDocuments()
        Await CorrectWordDocuments(Nothing, Nothing)
    End Sub

    ''' <summary>
    ''' Entry point for correction with optional prompt/suffix overrides.
    ''' </summary>
    ''' <summary>
    ''' Entry point for correction with optional prompt/suffix overrides.
    ''' </summary>
    Public Async Function CorrectWordDocuments(Optional promptOverride As String = Nothing,
                                          Optional correctedSuffixOverride As String = Nothing,
                                          Optional UseSecondAPI As Boolean = False) As System.Threading.Tasks.Task
        _correctPromptOverride = If(String.IsNullOrWhiteSpace(promptOverride), Nothing, promptOverride)
        _useSecondAPI = UseSecondAPI

        If String.IsNullOrWhiteSpace(correctedSuffixOverride) Then
            _correctSuffixOverride = Nothing
        Else
            Dim s As String = correctedSuffixOverride.Trim()
            If Not s.StartsWith("_"c) Then s = "_" & s
            _correctSuffixOverride = s
        End If

        Await ProcessWordDocuments(DocumentProcessMode.Correct)

        ' reset so subsequent runs go back to defaults
        _correctPromptOverride = Nothing
        _correctSuffixOverride = Nothing
        _useSecondAPI = False

        If UseSecondAPI And originalConfigLoaded Then
            RestoreDefaults(_context, originalConfig)
            originalConfigLoaded = False
        End If


    End Function


    ''' <summary>
    ''' Main processing method that handles both translation and correction modes.
    ''' </summary>
    ''' <param name="mode">The processing mode (Translate or Correct).</param>
    Private Async Function ProcessWordDocuments(mode As DocumentProcessMode) As System.Threading.Tasks.Task
        Dim selectedPath As String = ""
        Dim modeVerb As String = If(mode = DocumentProcessMode.Translate, "translate", "correct")
        Dim modeVerbPast As String = If(mode = DocumentProcessMode.Translate, "translated", "corrected")
        Dim modeVerbGerund As String = If(mode = DocumentProcessMode.Translate, "Translating", "Correcting")
        Dim modeNoun As String = If(mode = DocumentProcessMode.Translate, "Translation", "Correction")

        ' Effective correction suffix (used throughout for correction mode)
        Dim effectiveCorrectedSuffix As String = If(String.IsNullOrWhiteSpace(_correctSuffixOverride), CorrectedFileSuffix, _correctSuffixOverride)
        Globals.ThisAddIn.DragDropFormLabel = $"Select a Word document or folder to {modeVerb}"
        Globals.ThisAddIn.DragDropFormFilter = "Word Documents|*.doc;*.docx|Word Document (*.docx)|*.docx|Word 97-2003 (*.doc)|*.doc"

        Try
            Using frm As New DragDropForm(DragDropMode.FileOrDirectory)
                If frm.ShowDialog() = DialogResult.OK Then
                    selectedPath = frm.SelectedFilePath
                End If
            End Using
        Finally
            Globals.ThisAddIn.DragDropFormLabel = ""
            Globals.ThisAddIn.DragDropFormFilter = ""
        End Try

        If String.IsNullOrWhiteSpace(selectedPath) Then Exit Function

        Dim isDirectory As Boolean = Directory.Exists(selectedPath)
        Dim isFile As Boolean = File.Exists(selectedPath)

        If Not isFile AndAlso Not isDirectory Then
            ShowCustomMessageBox("The selected path does not exist.")
            Exit Function
        End If

        ' Collect files
        Dim filesToProcess As New List(Of String)()
        Dim wordExtensions As String() = {".doc", ".docx"}

        If isFile Then
            Dim ext As String = Path.GetExtension(selectedPath).ToLowerInvariant()
            If wordExtensions.Contains(ext) Then
                filesToProcess.Add(selectedPath)
            Else
                ShowCustomMessageBox($"File type '{ext}' is not supported.")
                Exit Function
            End If
        Else
            Dim recurseChoice As Integer = ShowCustomYesNoBox(
                $"Include Word documents from subdirectories?",
                "Yes, include subdirectories", "No, top directory only")
            If recurseChoice = 0 Then Exit Function

            Dim searchOption As SearchOption = If(recurseChoice = 1, SearchOption.AllDirectories, SearchOption.TopDirectoryOnly)

            ' Get all files and filter by exact extension match to avoid duplicates
            Dim allFiles = Directory.GetFiles(selectedPath, "*.*", searchOption)
            For Each f In allFiles
                Dim ext As String = Path.GetExtension(f).ToLowerInvariant()
                If ext = ".doc" OrElse ext = ".docx" Then
                    filesToProcess.Add(f)
                End If
            Next

            If filesToProcess.Count = 0 Then
                ShowCustomMessageBox("No Word documents found.")
                Exit Function
            End If
        End If

        ' Get target language (only for translation mode)
        Dim targetLanguage As String = Nothing
        Dim targetLanguageToken As String = Nothing

        If mode = DocumentProcessMode.Translate Then
            Dim defaultLanguage As String = If(String.IsNullOrWhiteSpace(INI_Language1), "English", INI_Language1)
            targetLanguage = ShowCustomInputBox(
                "Enter your target language (e.g., English, German, French):",
                AN & " Translate Word Files", True, defaultLanguage)

            If String.IsNullOrWhiteSpace(targetLanguage) Then Exit Function
            targetLanguage = targetLanguage.Trim()

            ' Normalize for file matching (also used for output filenames)
            targetLanguageToken = NormalizeLanguageTokenForFilename(targetLanguage)
            If String.IsNullOrWhiteSpace(targetLanguageToken) Then
                ShowCustomMessageBox("Invalid target language.")
                Exit Function
            End If
        Else
            ' For correction mode, use effective suffix
            targetLanguageToken = effectiveCorrectedSuffix.TrimStart("_"c)
        End If

        ' Build groups keyed by "base name"
        Dim groups As New Dictionary(Of String, (BaseFiles As List(Of String), ProcessedFiles As List(Of String)))(StringComparer.OrdinalIgnoreCase)

        For Each f In filesToProcess
            Dim ext As String = Path.GetExtension(f)
            If Not ext.Equals(".doc", StringComparison.OrdinalIgnoreCase) AndAlso Not ext.Equals(".docx", StringComparison.OrdinalIgnoreCase) Then Continue For

            Dim dir As String = Path.GetDirectoryName(f)
            Dim nameWithoutExt As String = Path.GetFileNameWithoutExtension(f)
            Dim impliedBase As String = TryGetImpliedBaseName(nameWithoutExt, targetLanguageToken)

            Dim groupBaseName As String = If(impliedBase, nameWithoutExt)
            Dim key As String = dir & "|" & groupBaseName

            If Not groups.ContainsKey(key) Then
                groups(key) = (New List(Of String)(), New List(Of String)())
            End If

            If impliedBase Is Nothing Then
                groups(key).BaseFiles.Add(f)
            Else
                groups(key).ProcessedFiles.Add(f)
            End If
        Next

        ' Partition groups
        Dim pairedGroups As New List(Of String)()          ' have base and processed
        Dim processedOnlyGroups As New List(Of String)()   ' processed exists, no base

        Dim groupsDedup As New Dictionary(Of String, (BaseFiles As List(Of String), ProcessedFiles As List(Of String)))(StringComparer.OrdinalIgnoreCase)

        For Each kvp In groups
            Dim baseFiles As List(Of String) = Dedup(kvp.Value.BaseFiles)
            Dim procFiles As List(Of String) = Dedup(kvp.Value.ProcessedFiles)

            groupsDedup(kvp.Key) = (baseFiles, procFiles)

            If baseFiles.Count > 0 AndAlso procFiles.Count > 0 Then
                pairedGroups.Add(kvp.Key)
            ElseIf baseFiles.Count = 0 AndAlso procFiles.Count > 0 Then
                processedOnlyGroups.Add(kvp.Key)
            End If
        Next

        ' After this, use groupsDedup everywhere instead of groups.
        groups = groupsDedup

        ' 1) Handle normal paired case
        If pairedGroups.Count > 0 Then
            Dim exampleKey As String = pairedGroups(0)
            Dim exBase As String = Path.GetFileName(groups(exampleKey).BaseFiles(0))
            Dim exProc As String = Path.GetFileName(groups(exampleKey).ProcessedFiles(0))

            Dim suffixDesc As String = If(mode = DocumentProcessMode.Translate,
    $"'{targetLanguage}' translation (suffix '_{targetLanguageToken}')",
    $"modified version (suffix '{effectiveCorrectedSuffix}')")

            Dim msg As New StringBuilder()
            msg.AppendLine($"Found {pairedGroups.Count} document(s) that already have an existing {suffixDesc}.")
            msg.AppendLine()
            msg.AppendLine($"If you skip: both the base file and its existing {modeNoun.ToLower()} will be skipped.")
            msg.AppendLine()
            msg.AppendLine($"If you re-{modeVerb}: the existing {modeNoun.ToLower()} file(s) will be deleted, and the base file will be {modeVerbPast} again.")
            msg.AppendLine()
            msg.AppendLine("Example:")
            msg.AppendLine($"  Base:        {exBase}")
            msg.AppendLine($"  {modeNoun}: {exProc}")

            Dim choice As Integer = ShowCustomYesNoBox(
                msg.ToString().TrimEnd(),
                "Skip these documents", $"Delete {modeNoun.ToLower()}s and re-{modeVerb}")

            If choice = 0 Then Exit Function

            Dim toExclude As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            If choice = 1 Then
                ' Skip both base and processed
                For Each k In pairedGroups
                    For Each p In groups(k).BaseFiles : toExclude.Add(p) : Next
                    For Each p In groups(k).ProcessedFiles : toExclude.Add(p) : Next
                Next
            Else
                ' Delete processed file(s), keep base
                For Each k In pairedGroups
                    For Each procPath In groups(k).ProcessedFiles
                        Try
                            If File.Exists(procPath) Then File.Delete(procPath)
                            ' Also delete compare file if it exists (for correction mode)
                            If mode = DocumentProcessMode.Correct Then
                                Dim comparePath As String = GetCompareFilePath(procPath)
                                If File.Exists(comparePath) Then File.Delete(comparePath)
                            End If
                        Catch ex As Exception
                            ' If we cannot delete, safest is to skip this pair (avoid overwriting surprises)
                            For Each p In groups(k).BaseFiles : toExclude.Add(p) : Next
                            For Each p In groups(k).ProcessedFiles : toExclude.Add(p) : Next
                        End Try
                    Next

                    ' Never process processed files themselves in this mode
                    For Each p In groups(k).ProcessedFiles : toExclude.Add(p) : Next
                Next
            End If

            filesToProcess = filesToProcess.Where(Function(p) Not toExclude.Contains(p)).ToList()
        End If

        If filesToProcess.Count = 0 Then
            ShowCustomMessageBox($"No documents remaining for {modeNoun.ToLower()}.")
            Exit Function
        End If

        ' 2) Edge case: processed-only files (no base)
        If processedOnlyGroups.Count > 0 Then
            Dim exampleKey As String = processedOnlyGroups(0)
            Dim exProc As String = Path.GetFileName(groups(exampleKey).ProcessedFiles(0))

            Dim suffixDesc As String = If(mode = DocumentProcessMode.Translate,
    $"'{targetLanguage}' (suffix '_{targetLanguageToken}')",
    $"modified (suffix '{effectiveCorrectedSuffix}')")

            Dim msg2 As New StringBuilder()
            msg2.AppendLine($"Found {processedOnlyGroups.Count} {modeNoun.ToLower()}-only file(s) for {suffixDesc}, without a matching base file.")
            msg2.AppendLine()
            msg2.AppendLine($"{modeVerb.Substring(0, 1).ToUpper()}{modeVerb.Substring(1)} these files too (treat them as base files)?")
            msg2.AppendLine()
            msg2.AppendLine("Example:")
            msg2.AppendLine($"  {exProc}")

            Dim choice2 As Integer = ShowCustomYesNoBox(
                msg2.ToString().TrimEnd(),
                $"Yes, {modeVerb} them too", "No, skip them")

            If choice2 = 0 Then Exit Function

            If choice2 <> 1 Then
                Dim toExclude2 As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                For Each k In processedOnlyGroups
                    For Each p In groups(k).ProcessedFiles : toExclude2.Add(p) : Next
                Next
                filesToProcess = filesToProcess.Where(Function(p) Not toExclude2.Contains(p)).ToList()
            End If
        End If

        If filesToProcess.Count = 0 Then
            ShowCustomMessageBox($"No documents remaining for {modeNoun.ToLower()}.")
            Exit Function
        End If

        ' Confirm if many files to process
        If filesToProcess.Count > 10 Then
            Dim confirmMsg As String = If(mode = DocumentProcessMode.Translate,
                $"Ready to translate {filesToProcess.Count} document(s) to {targetLanguage}. Continue?",
                $"Ready to correct {filesToProcess.Count} document(s). Continue?")
            Dim confirm As Integer = ShowCustomYesNoBox(confirmMsg, "Yes, continue", "No, abort")
            If confirm <> 1 Then Exit Function
        End If

        ' Process
        ProgressBarModule.GlobalProgressValue = 0
        ProgressBarModule.GlobalProgressMax = filesToProcess.Count
        ProgressBarModule.GlobalProgressLabel = "Initializing..."
        ProgressBarModule.CancelOperation = False
        ProgressBarModule.ShowProgressBarInSeparateThread(AN & " " & modeNoun, "Starting...")

        Dim successCount As Integer = 0
        Dim failedFiles As New List(Of String)()

        Try
            For i As Integer = 0 To filesToProcess.Count - 1
                If ProgressBarModule.CancelOperation Then Exit For

                Dim filePath As String = filesToProcess(i)
                Dim fileName As String = Path.GetFileName(filePath)

                ProgressBarModule.GlobalProgressValue = i
                ProgressBarModule.GlobalProgressLabel = $"{modeVerbGerund} {i + 1}/{filesToProcess.Count}: {fileName}"

                Try
                    Dim dir As String = Path.GetDirectoryName(filePath)
                    Dim nameWithoutExt As String = Path.GetFileNameWithoutExtension(filePath)
                    Dim outputPath As String

                    If mode = DocumentProcessMode.Translate Then
                        outputPath = Path.Combine(dir, $"{nameWithoutExt}_{targetLanguageToken}.docx")
                    Else
                        outputPath = Path.Combine(dir, $"{nameWithoutExt}{effectiveCorrectedSuffix}.docx")
                    End If

                    Dim success As Boolean = Await ProcessDocumentViaOpenXml(filePath, outputPath, targetLanguage, mode)

                    If success Then
                        ' For correction mode, create compare document
                        If mode = DocumentProcessMode.Correct Then
                            Dim compareSuccess As Boolean = CreateWordCompareDocument(filePath, outputPath)
                            If Not compareSuccess Then
                                failedFiles.Add($"{fileName}: Corrected but compare document creation failed")
                                successCount += 1 ' Still count as success since correction worked
                            Else
                                successCount += 1
                            End If
                        Else
                            successCount += 1
                        End If
                    Else
                        failedFiles.Add($"{fileName}: {modeNoun} failed")
                    End If
                Catch ex As Exception
                    failedFiles.Add($"{fileName}: {ex.Message}")
                End Try
            Next
        Finally
            ProgressBarModule.CancelOperation = True
        End Try

        ' Summary
        Dim summary As New StringBuilder()
        If (successCount + failedFiles.Count) < filesToProcess.Count Then
            summary.AppendLine("Operation was cancelled.")
            summary.AppendLine()
        End If

        summary.AppendLine($"Successfully {modeVerbPast}: {successCount} file(s)")
        If mode = DocumentProcessMode.Translate Then
            summary.AppendLine($"Target language: {targetLanguage}")
        Else
            summary.AppendLine($"Compare documents created with tracked changes")
        End If

        If failedFiles.Count > 0 Then
            summary.AppendLine()
            summary.AppendLine($"Failed: {failedFiles.Count} file(s)")
            For Each f In failedFiles.Take(10)
                summary.AppendLine($"  • {f}")
            Next
            If failedFiles.Count > 10 Then
                summary.AppendLine($"  ... and {failedFiles.Count - 10} more")
            End If
            SharedMethods.PutInClipboard(String.Join(vbCrLf, failedFiles))
            summary.AppendLine("(Log copied to clipboard)")
        End If

        ShowCustomMessageBox(summary.ToString().TrimEnd(), AN & " " & modeNoun)
    End Function

    ''' <summary>
    ''' Gets the compare file path from a corrected file path.
    ''' </summary>
    Private Function GetCompareFilePath(correctedPath As String) As String
        Dim dir As String = Path.GetDirectoryName(correctedPath)
        Dim nameWithoutExt As String = Path.GetFileNameWithoutExtension(correctedPath)
        Dim ext As String = Path.GetExtension(correctedPath)

        Dim effectiveCorrectedSuffix As String = If(String.IsNullOrWhiteSpace(_correctSuffixOverride), CorrectedFileSuffix, _correctSuffixOverride)
        Dim effectiveCompareSuffix As String = effectiveCorrectedSuffix & "_compare"

        ' Replace effective suffix with compare suffix
        If nameWithoutExt.EndsWith(effectiveCorrectedSuffix, StringComparison.OrdinalIgnoreCase) Then
            nameWithoutExt = nameWithoutExt.Substring(0, nameWithoutExt.Length - effectiveCorrectedSuffix.Length) & effectiveCompareSuffix
        Else
            nameWithoutExt = nameWithoutExt & effectiveCompareSuffix
        End If

        Return Path.Combine(dir, nameWithoutExt & ext)
    End Function

    ''' <summary>
    ''' Creates a Word compare document showing differences between original and corrected documents.
    ''' </summary>
    ''' <param name="originalPath">Path to the original document.</param>
    ''' <param name="correctedPath">Path to the corrected document.</param>
    ''' <returns>True if successful, False otherwise.</returns>
    Private Function CreateWordCompareDocument(originalPath As String, correctedPath As String) As Boolean
        Dim wordApp As Word.Application = Nothing
        Dim originalDoc As Word.Document = Nothing
        Dim correctedDoc As Word.Document = Nothing
        Dim compareDoc As Word.Document = Nothing

        Try
            wordApp = Globals.ThisAddIn.Application
            Dim wasScreenUpdating As Boolean = wordApp.ScreenUpdating
            wordApp.ScreenUpdating = False

            ' Determine compare output path
            Dim comparePath As String = GetCompareFilePath(correctedPath)

            ' Open both documents
            originalDoc = wordApp.Documents.Open(originalPath, ReadOnly:=True, Visible:=False, AddToRecentFiles:=False)
            correctedDoc = wordApp.Documents.Open(correctedPath, ReadOnly:=True, Visible:=False, AddToRecentFiles:=False)

            ' Create comparison document using Word's built-in compare feature
            ' This preserves all formatting and shows changes as tracked changes
            compareDoc = wordApp.CompareDocuments(
                OriginalDocument:=originalDoc,
                RevisedDocument:=correctedDoc,
                Destination:=WdCompareDestination.wdCompareDestinationNew,
                Granularity:=WdGranularity.wdGranularityWordLevel,
                CompareFormatting:=True,
                CompareCaseChanges:=True,
                CompareWhitespace:=True,
                CompareTables:=True,
                CompareHeaders:=True,
                CompareFootnotes:=True,
                CompareTextboxes:=True,
                CompareFields:=True,
                CompareComments:=True,
                RevisedAuthor:=wordApp.UserName,
                IgnoreAllComparisonWarnings:=True
            )

            ' Save the compare document
            compareDoc.SaveAs2(comparePath, WdSaveFormat.wdFormatXMLDocument)

            ' Close documents
            compareDoc.Close(WdSaveOptions.wdDoNotSaveChanges)
            compareDoc = Nothing

            correctedDoc.Close(WdSaveOptions.wdDoNotSaveChanges)
            correctedDoc = Nothing

            originalDoc.Close(WdSaveOptions.wdDoNotSaveChanges)
            originalDoc = Nothing

            wordApp.ScreenUpdating = wasScreenUpdating

            Return True

        Catch ex As Exception
            Debug.WriteLine($"CreateWordCompareDocument error: {ex.Message}")
            Return False
        Finally
            ' Cleanup
            If compareDoc IsNot Nothing Then
                Try : compareDoc.Close(WdSaveOptions.wdDoNotSaveChanges) : Catch : End Try
            End If
            If correctedDoc IsNot Nothing Then
                Try : correctedDoc.Close(WdSaveOptions.wdDoNotSaveChanges) : Catch : End Try
            End If
            If originalDoc IsNot Nothing Then
                Try : originalDoc.Close(WdSaveOptions.wdDoNotSaveChanges) : Catch : End Try
            End If
            If wordApp IsNot Nothing Then
                Try : wordApp.ScreenUpdating = True : Catch : End Try
            End If
        End Try
    End Function


    ''' <summary>
    ''' Normalizes a human-entered language name (e.g., "English (US)") into a safe token
    ''' for use in filenames by replacing non-alphanumeric characters with underscores.
    ''' </summary>
    ''' <param name="language">The target language entered by the user.</param>
    ''' <returns>A filename-safe token (e.g., <c>English_US</c>), or an empty string if invalid.</returns>
    Private Shared Function NormalizeLanguageTokenForFilename(language As String) As String
        If String.IsNullOrWhiteSpace(language) Then Return ""

        Dim s As String = language.Trim()

        ' Replace all non-letter/digit with underscore, collapse multiples, trim underscores
        s = Regex.Replace(s, "[^\p{L}\p{Nd}]+", "_")
        s = Regex.Replace(s, "_{2,}", "_")
        s = s.Trim("_"c)

        Return s
    End Function

    ''' <summary>
    ''' Infers a "base" filename from a processed filename by removing a trailing
    ''' <c>_{token}</c> suffix (optionally followed by a copy/counter suffix).
    ''' </summary>
    ''' <param name="fileBaseName">Filename without extension.</param>
    ''' <param name="token">Normalized token (filename-safe) - language token or "corrected".</param>
    ''' <returns>The inferred base name if it matches; otherwise <c>Nothing</c>.</returns>

    Private Shared Function TryGetImpliedBaseName(fileBaseName As String, token As String) As String
        If String.IsNullOrWhiteSpace(fileBaseName) OrElse String.IsNullOrWhiteSpace(token) Then Return Nothing

        Dim escaped As String = Regex.Escape(token)

        ' Matches:
        '   ABC_English
        '   ABC_English (1)
        '   ABC_English_1
        '   ABC_corrected
        '   ABC_corrected_compare (also strip compare suffix)
        Dim m As Match = Regex.Match(
        fileBaseName,
        "^(.*)_" & escaped & "(?:_compare)?(?:\s*\(\d+\)|_\d+)?$",
        RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant)

        If Not m.Success Then Return Nothing

        Dim basePart As String = m.Groups(1).Value
        If String.IsNullOrWhiteSpace(basePart) Then Return Nothing

        Return basePart
    End Function

    ''' <summary>
    ''' Adds a file path to a list only if it is non-empty and exists on disk.
    ''' </summary>
    ''' <param name="list">Target list to add into.</param>
    ''' <param name="path">Candidate path.</param>
    Private Shared Sub AddIfExists(list As List(Of String), path As String)
        If Not String.IsNullOrWhiteSpace(path) AndAlso File.Exists(path) Then list.Add(path)
    End Sub

    ''' <summary>
    ''' De-duplicates a sequence of file paths using case-insensitive comparison.
    ''' </summary>
    ''' <param name="paths">Input paths.</param>
    ''' <returns>A new list containing distinct paths.</returns>

    Private Shared Function Dedup(paths As IEnumerable(Of String)) As List(Of String)
        Return paths.Distinct(StringComparer.OrdinalIgnoreCase).ToList()
    End Function


    ''' <summary>
    ''' Checks whether a filename (without extension) already appears to represent a translation
    ''' for the current language token (supports common "(1)" and "_1" copy suffixes).
    ''' </summary>
    ''' <param name="baseName">Filename without extension.</param>
    ''' <param name="languageToken">Normalized language token (filename-safe).</param>
    ''' <returns><c>True</c> if the name likely represents a translation; otherwise <c>False</c>.</returns>

    Private Shared Function IsLikelyTranslationFile(baseName As String, languageToken As String) As Boolean
        If String.IsNullOrWhiteSpace(baseName) OrElse String.IsNullOrWhiteSpace(languageToken) Then Return False

        Dim escaped As String = Regex.Escape(languageToken)

        ' End of name:
        '   _<token>
        '   _<token> (digits)
        '   _<token>_digits
        Dim pattern As String = "_.?" & escaped & "(?:\s*\(\d+\)|_\d+)?$"
        ' Note: "_.?" is intentionally NOT used here; keep strict underscore.
        pattern = "_" & escaped & "(?:\s*\(\d+\)|_\d+)?$"

        Return Regex.IsMatch(baseName, pattern, RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant)
    End Function


    ''' <summary>
    ''' Processes a document using OpenXML direct text node manipulation.
    ''' </summary>
    Private Async Function ProcessDocumentViaOpenXml(
        inputPath As String,
        outputPath As String,
        targetLanguage As String,
        mode As DocumentProcessMode) As System.Threading.Tasks.Task(Of Boolean)

        Dim tempDocxPath As String = Nothing
        Dim wordApp As Word.Application = Nothing
        Dim doc As Word.Document = Nothing

        Try
            ' If .doc, convert to .docx first
            If Path.GetExtension(inputPath).ToLowerInvariant() = ".doc" Then
                tempDocxPath = Path.Combine(Path.GetTempPath(), $"{AN2}_conv_{Guid.NewGuid():N}.docx")
                wordApp = Globals.ThisAddIn.Application
                wordApp.ScreenUpdating = False
                doc = wordApp.Documents.Open(inputPath, ReadOnly:=True, Visible:=False, AddToRecentFiles:=False)
                doc.SaveAs2(tempDocxPath, WdSaveFormat.wdFormatXMLDocument)
                doc.Close(WdSaveOptions.wdDoNotSaveChanges)
                doc = Nothing
                wordApp.ScreenUpdating = True
            Else
                tempDocxPath = inputPath
            End If

            ' Copy to output path first (we'll modify the copy)
            File.Copy(tempDocxPath, outputPath, overwrite:=True)

            ' Process the DOCX via OpenXML
            Dim success As Boolean = Await ProcessDocxOpenXml(outputPath, targetLanguage, mode)

            Return success

        Catch ex As Exception
            Debug.WriteLine($"ProcessDocumentViaOpenXml error: {ex.Message}")
            Throw
        Finally
            If doc IsNot Nothing Then
                Try : doc.Close(WdSaveOptions.wdDoNotSaveChanges) : Catch : End Try
            End If
            ' Clean up temp file if we created one
            If tempDocxPath IsNot Nothing AndAlso tempDocxPath <> inputPath AndAlso File.Exists(tempDocxPath) Then
                Try : File.Delete(tempDocxPath) : Catch : End Try
            End If
        End Try
    End Function

    ''' <summary>
    ''' Processes a DOCX file using OpenXML to translate or correct text nodes.
    ''' </summary>
    Private Async Function ProcessDocxOpenXml(docxPath As String, targetLanguage As String, mode As DocumentProcessMode) As System.Threading.Tasks.Task(Of Boolean)
        ' DOCX is a ZIP file - extract, modify document.xml, repack
        Dim tempDir As String = Path.Combine(Path.GetTempPath(), $"{AN2}_xml_{Guid.NewGuid():N}")

        Try
            ' Extract DOCX
            ZipFile.ExtractToDirectory(docxPath, tempDir)

            ' Set up namespace manager (reused for all XML files)
            Dim nsMgr As System.Xml.XmlNamespaceManager = Nothing

            ' === Process document.xml ===
            Dim documentXmlPath As String = Path.Combine(tempDir, "word", "document.xml")
            If Not File.Exists(documentXmlPath) Then
                ShowCustomMessageBox("Invalid DOCX structure - document.xml not found.")
                Return False
            End If

            Dim xmlDoc As New System.Xml.XmlDocument()
            xmlDoc.PreserveWhitespace = True
            xmlDoc.Load(documentXmlPath)

            nsMgr = New System.Xml.XmlNamespaceManager(xmlDoc.NameTable)
            nsMgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

            ' Extract paragraphs with their text runs
            Dim paragraphs As List(Of TranslateParagraphInfo) = ExtractTranslateParagraphsFromXml(xmlDoc, nsMgr)

            If paragraphs.Count = 0 OrElse paragraphs.All(Function(p) p.IsEmpty) Then
                ShowCustomMessageBox("No processable text found.")
                Return False
            End If

            ' Warn for large documents
            Dim processableCount As Integer = paragraphs.Where(Function(p) Not p.IsEmpty).Count()
            Dim modeVerb As String = If(mode = DocumentProcessMode.Translate, "translating", "correcting")
            If processableCount > TranslateLargeDocThreshold Then
                Dim continueChoice As Integer = ShowCustomYesNoBox(
                $"Document has {processableCount} paragraphs. This may take several minutes. Continue?",
                "Yes, continue", "No, skip")
                If continueChoice <> 1 Then Return False
            End If

            ' Process paragraphs in batches (sending ONLY plain text to LLM)            
            Dim success As Boolean = Await ProcessParagraphBatches(paragraphs, targetLanguage, mode, Path.GetFileName(docxPath))
            If Not success Then Return False

            ' Apply processed text back to XML nodes
            ApplyTranslationsToXml(paragraphs)

            ' Save modified document.xml
            xmlDoc.Save(documentXmlPath)

            ' === Process comments.xml (if exists) ===
            Dim commentsXmlPath As String = Path.Combine(tempDir, "word", "comments.xml")
            If File.Exists(commentsXmlPath) Then
                Dim commentsSuccess As Boolean = Await ProcessCommentsXml(commentsXmlPath, targetLanguage, mode, Path.GetFileName(docxPath))
                ' Continue even if comments fail - main document is more important
            End If

            ' === Process headers and footers ===
            Await ProcessHeadersFooters(tempDir, targetLanguage, mode, Path.GetFileName(docxPath))

            ' === Process footnotes and endnotes ===
            Await ProcessFootnotesEndnotes(tempDir, targetLanguage, mode, Path.GetFileName(docxPath))

            ' Repack DOCX
            File.Delete(docxPath)
            ZipFile.CreateFromDirectory(tempDir, docxPath, CompressionLevel.Optimal, False)

            Return True

        Finally
            ' Cleanup
            If Directory.Exists(tempDir) Then
                Try : Directory.Delete(tempDir, recursive:=True) : Catch : End Try
            End If
        End Try
    End Function

    ''' <summary>
    ''' Processes comments.xml to translate or correct comment text.
    ''' </summary>
    Private Async Function ProcessCommentsXml(commentsXmlPath As String, targetLanguage As String, mode As DocumentProcessMode, mainFileName As String) As System.Threading.Tasks.Task(Of Boolean)
        Try
            Dim xmlDoc As New System.Xml.XmlDocument()
            xmlDoc.PreserveWhitespace = True
            xmlDoc.Load(commentsXmlPath)

            Dim nsMgr As New System.Xml.XmlNamespaceManager(xmlDoc.NameTable)
            nsMgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

            ' Extract paragraphs from comments (comments contain w:p elements just like document.xml)
            Dim paragraphs As List(Of TranslateParagraphInfo) = ExtractTranslateParagraphsFromXml(xmlDoc, nsMgr)

            Dim processableParagraphs = paragraphs.Where(Function(p) Not p.IsEmpty).ToList()
            If processableParagraphs.Count = 0 Then Return True

            ' Process comment paragraphs
            Dim success As Boolean = Await ProcessParagraphBatches(paragraphs, targetLanguage, mode, $"{mainFileName} (Comments)")
            If Not success Then Return False

            ' Apply processed text
            ApplyTranslationsToXml(paragraphs)

            ' Save
            xmlDoc.Save(commentsXmlPath)
            Return True

        Catch ex As Exception
            Debug.WriteLine($"ProcessCommentsXml error: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Processes header and footer XML files.
    ''' </summary>
    Private Async Function ProcessHeadersFooters(tempDir As String, targetLanguage As String, mode As DocumentProcessMode, mainFileName As String) As System.Threading.Tasks.Task

        Dim wordDir As String = Path.Combine(tempDir, "word")
        If Not Directory.Exists(wordDir) Then Exit Function

        ' Headers: header1.xml, header2.xml, header3.xml, etc.
        ' Footers: footer1.xml, footer2.xml, footer3.xml, etc.
        Dim patterns As String() = {"header*.xml", "footer*.xml"}

        For Each pattern In patterns
            For Each filePath In Directory.GetFiles(wordDir, pattern)
                Try
                    Dim xmlDoc As New System.Xml.XmlDocument()
                    xmlDoc.PreserveWhitespace = True
                    xmlDoc.Load(filePath)

                    Dim nsMgr As New System.Xml.XmlNamespaceManager(xmlDoc.NameTable)
                    nsMgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

                    Dim paragraphs As List(Of TranslateParagraphInfo) = ExtractTranslateParagraphsFromXml(xmlDoc, nsMgr)

                    Dim processableParagraphs = paragraphs.Where(Function(p) Not p.IsEmpty).ToList()
                    If processableParagraphs.Count = 0 Then Continue For

                    Dim componentType As String = If(Path.GetFileName(filePath).StartsWith("header", StringComparison.OrdinalIgnoreCase), "Headers", "Footers")
                    Dim success As Boolean = Await ProcessParagraphBatches(paragraphs, targetLanguage, mode, $"{mainFileName} ({componentType})")
                    If success Then
                        ApplyTranslationsToXml(paragraphs)
                        xmlDoc.Save(filePath)
                    End If

                Catch ex As Exception
                    Debug.WriteLine($"ProcessHeadersFooters error for {Path.GetFileName(filePath)}: {ex.Message}")
                    ' Continue with other files
                End Try
            Next
        Next
    End Function
    ''' <summary>
    ''' Extracts paragraph information from document XML.
    ''' Only extracts plain text - no formatting codes sent to LLM.
    ''' Preserves space information for accurate redistribution.
    ''' </summary>
    Private Function ExtractTranslateParagraphsFromXml(xmlDoc As System.Xml.XmlDocument, nsMgr As System.Xml.XmlNamespaceManager) As List(Of TranslateParagraphInfo)
        Dim paragraphs As New List(Of TranslateParagraphInfo)()

        ' Find all w:p (paragraph) elements
        Dim paraNodes As System.Xml.XmlNodeList = xmlDoc.SelectNodes("//w:p", nsMgr)
        Dim paraIndex As Integer = 0

        For Each paraNode As System.Xml.XmlNode In paraNodes
            Dim paraInfo As New TranslateParagraphInfo() With {
            .Index = paraIndex,
            .TextRuns = New List(Of TranslateTextRunInfo)(),
            .TranslatedText = Nothing
        }

            ' Find all w:t (text) elements within this paragraph
            Dim textNodes As System.Xml.XmlNodeList = paraNode.SelectNodes(".//w:t", nsMgr)
            Dim fullTextBuilder As New StringBuilder()

            For Each textNode As System.Xml.XmlNode In textNodes
                Dim text As String = textNode.InnerText

                ' Check if this run needs a space before it
                ' Word sometimes splits "word1 word2" into separate runs without explicit space
                If fullTextBuilder.Length > 0 AndAlso text.Length > 0 Then
                    Dim lastChar As Char = fullTextBuilder(fullTextBuilder.Length - 1)
                    Dim firstChar As Char = text(0)

                    ' If previous text doesn't end with space/punctuation and current doesn't start with space/punctuation,
                    ' check if we need to infer a space based on the XML structure
                    If Not Char.IsWhiteSpace(lastChar) AndAlso Not Char.IsWhiteSpace(firstChar) Then
                        ' Check xml:space attribute - if "preserve" is set, spaces are explicit
                        Dim xmlSpaceAttr = textNode.Attributes("xml:space")
                        Dim preserveSpace As Boolean = xmlSpaceAttr IsNot Nothing AndAlso xmlSpaceAttr.Value = "preserve"

                        ' If not preserving space and the run is in a new w:r element, 
                        ' Word may have intended a space (common when formatting changes mid-word)
                        ' However, we should NOT add a space if the original didn't have one
                        ' The issue is the REVERSE - we're losing spaces that WERE there
                    End If
                End If

                paraInfo.TextRuns.Add(New TranslateTextRunInfo() With {
                .TextNode = textNode,
                .OriginalText = text
            })
                fullTextBuilder.Append(text)
            Next

            paraInfo.FullText = fullTextBuilder.ToString()
            paraInfo.IsEmpty = String.IsNullOrWhiteSpace(paraInfo.FullText)

            paragraphs.Add(paraInfo)
            paraIndex += 1
        Next

        Return paragraphs
    End Function

    ''' <summary>
    ''' Processes paragraphs in batches with context windows.
    ''' Sends ONLY plain text to the LLM - no XML, no formatting codes.
    ''' </summary>
    Private Async Function ProcessParagraphBatches(
        paragraphs As List(Of TranslateParagraphInfo),
        targetLanguage As String,
        mode As DocumentProcessMode,
        Optional fileContext As String = "") As System.Threading.Tasks.Task(Of Boolean)

        Dim processableParagraphs = paragraphs.Where(Function(p) Not p.IsEmpty).ToList()
        If processableParagraphs.Count = 0 Then Return True

        ' Select appropriate system prompt based on mode
        Dim systemPrompt As String
        If mode = DocumentProcessMode.Translate Then
            TranslateLanguage = targetLanguage
            systemPrompt = InterpolateAtRuntime(SP_Translate_Document)
        Else
            Dim effectivePrompt As String = If(String.IsNullOrWhiteSpace(_correctPromptOverride), SP_Correct_Document, _correctPromptOverride)
            systemPrompt = InterpolateAtRuntime(effectivePrompt)
        End If

        Dim batchIndex As Integer = 0
        Dim totalBatches As Integer = CInt(Math.Ceiling(processableParagraphs.Count / TranslateParagraphsPerBatch))
        Dim modeVerbGerund As String = If(mode = DocumentProcessMode.Translate, "Translating", "Correcting")

        While batchIndex < processableParagraphs.Count
            If ProgressBarModule.CancelOperation Then Return False

            ' Determine batch boundaries
            Dim batchStart As Integer = batchIndex
            Dim batchEnd As Integer = Math.Min(batchIndex + TranslateParagraphsPerBatch - 1, processableParagraphs.Count - 1)

            ' Adjust for character limit
            Dim batchChars As Integer = 0
            For j As Integer = batchStart To batchEnd
                batchChars += processableParagraphs(j).FullText.Length
                If batchChars > TranslateMaxCharsPerBatch AndAlso j > batchStart Then
                    batchEnd = j - 1
                    Exit For
                End If
            Next

            ' Build prompt with ONLY plain text - no formatting codes
            Dim promptBuilder As New StringBuilder()

            ' Context Before (already processed - plain text only)
            Dim contextBeforeStart As Integer = Math.Max(0, batchStart - TranslateContextBefore)
            If contextBeforeStart < batchStart Then
                promptBuilder.AppendLine("[CONTEXT BEFORE - for reference only]")
                For j As Integer = contextBeforeStart To batchStart - 1
                    Dim p = processableParagraphs(j)
                    promptBuilder.AppendLine(If(p.TranslatedText, p.FullText))
                Next
                promptBuilder.AppendLine()
            End If

            ' Paragraphs to process (plain text only)
            promptBuilder.AppendLine("[TEXTTOPROCESS]")
            Dim batchNumber As Integer = 1
            For j As Integer = batchStart To batchEnd
                promptBuilder.AppendLine($"[{batchNumber}] {processableParagraphs(j).FullText}")
                batchNumber += 1
            Next
            promptBuilder.AppendLine("[/TEXTTOPROCESS]")
            promptBuilder.AppendLine()

            ' Context After (upcoming - plain text only)
            Dim contextAfterEnd As Integer = Math.Min(processableParagraphs.Count - 1, batchEnd + TranslateContextAfter)
            If contextAfterEnd > batchEnd Then
                promptBuilder.AppendLine("[CONTEXT AFTER - for reference only]")
                For j As Integer = batchEnd + 1 To contextAfterEnd
                    promptBuilder.AppendLine(processableParagraphs(j).FullText)
                Next
            End If

            ' Update progress - include file context if provided
            Dim currentBatch As Integer = CInt(Math.Floor(batchIndex / TranslateParagraphsPerBatch)) + 1
            If Not String.IsNullOrEmpty(fileContext) Then
                ProgressBarModule.GlobalProgressLabel = $"{fileContext} - batch {currentBatch}/{totalBatches}"
            Else
                ProgressBarModule.GlobalProgressLabel = $"{modeVerbGerund} batch {currentBatch}/{totalBatches}"
            End If

            ' Call LLM with pure text only
            Dim response As String = Await SharedMethods.LLM(
            _context, systemPrompt, promptBuilder.ToString(),
            "", "", 0, _useSecondAPI, True)

            If String.IsNullOrWhiteSpace(response) Then
                Dim modeNoun As String = If(mode = DocumentProcessMode.Translate, "Translation", "Correction")
                ShowCustomMessageBox($"LLM returned empty response. {modeNoun} incomplete.")
                Return False
            End If

            ' Parse and store results
            ParseTranslateResponse(response, processableParagraphs, batchStart, batchEnd)

            batchIndex = batchEnd + 1
        End While

        Return True
    End Function
    ''' <summary>
    ''' Parses LLM response and stores translations/corrections.
    ''' </summary>
    Private Sub ParseTranslateResponse(
        response As String,
        paragraphs As List(Of TranslateParagraphInfo),
        batchStart As Integer,
        batchEnd As Integer)

        ' Pattern: [n] followed by content until next [n] or end
        Dim pattern As New Regex("\[(\d+)\]\s*(.*?)(?=\s*\[\d+\]|$)", RegexOptions.Singleline)
        Dim matches = pattern.Matches(response)

        For Each m As Match In matches
            Dim num As Integer
            If Integer.TryParse(m.Groups(1).Value, num) Then
                Dim absoluteIndex As Integer = batchStart + num - 1
                If absoluteIndex >= batchStart AndAlso absoluteIndex <= batchEnd AndAlso absoluteIndex < paragraphs.Count Then
                    Dim processed As String = m.Groups(2).Value.Trim()
                    ' Clean any stray markers
                    processed = Regex.Replace(processed, "^\[/?TEXTTOPROCESS\]", "", RegexOptions.IgnoreCase).Trim()
                    processed = Regex.Replace(processed, "\[/?TEXTTOPROCESS\]$", "", RegexOptions.IgnoreCase).Trim()
                    ' Also clean old TRANSLATE markers for backward compatibility
                    processed = Regex.Replace(processed, "^\[/?TRANSLATE\]", "", RegexOptions.IgnoreCase).Trim()
                    processed = Regex.Replace(processed, "\[/?TRANSLATE\]$", "", RegexOptions.IgnoreCase).Trim()
                    paragraphs(absoluteIndex).TranslatedText = processed
                End If
            End If
        Next
    End Sub

    ''' <summary>
    ''' Applies translated/corrected text back to XML nodes, preserving all formatting.
    ''' Uses character-based distribution with guaranteed spacing.
    ''' </summary>
    Private Sub ApplyTranslationsToXml(paragraphs As List(Of TranslateParagraphInfo))
        For Each para In paragraphs
            If para.IsEmpty OrElse String.IsNullOrEmpty(para.TranslatedText) Then Continue For
            If para.TextRuns.Count = 0 Then Continue For

            Dim translatedText As String = para.TranslatedText

            ' Simple case: only one run
            If para.TextRuns.Count = 1 Then
                SetTextNodeWithSpacePreserve(para.TextRuns(0).TextNode, translatedText)
                Continue For
            End If

            Dim totalOriginalLength As Integer = para.FullText.Length
            If totalOriginalLength = 0 Then
                SetTextNodeWithSpacePreserve(para.TextRuns(0).TextNode, translatedText)
                For idx As Integer = 1 To para.TextRuns.Count - 1
                    SetTextNodeWithSpacePreserve(para.TextRuns(idx).TextNode, "")
                Next
                Continue For
            End If

            ' Character-based proportional distribution
            Dim translatedLength As Integer = translatedText.Length
            Dim currentPos As Integer = 0
            Dim cumulativeOriginal As Integer = 0

            For runIdx As Integer = 0 To para.TextRuns.Count - 1
                Dim run = para.TextRuns(runIdx)
                Dim originalRunLength As Integer = run.OriginalText.Length
                cumulativeOriginal += originalRunLength

                If runIdx = para.TextRuns.Count - 1 Then
                    ' Last run gets everything remaining
                    Dim remaining As String = If(currentPos < translatedLength,
                                           translatedText.Substring(currentPos),
                                           "")
                    SetTextNodeWithSpacePreserve(run.TextNode, remaining)
                Else
                    ' Calculate end position based on cumulative proportion
                    Dim proportion As Double = cumulativeOriginal / CDbl(totalOriginalLength)
                    Dim targetEndPos As Integer = CInt(Math.Round(proportion * translatedLength))
                    targetEndPos = Math.Min(targetEndPos, translatedLength)

                    ' Don't go backwards
                    If targetEndPos <= currentPos Then
                        SetTextNodeWithSpacePreserve(run.TextNode, "")
                        Continue For
                    End If

                    ' Try to break at a word boundary (space)
                    Dim endPos As Integer = targetEndPos
                    Dim foundSpace As Boolean = False

                    If endPos < translatedLength AndAlso endPos > currentPos Then
                        ' Check if we're already at a space
                        If translatedText(endPos) = " "c Then
                            endPos = endPos + 1  ' Include the space
                            foundSpace = True
                        Else
                            ' Search forward first (up to 15 chars)
                            For searchPos As Integer = endPos To Math.Min(endPos + 15, translatedLength - 1)
                                If translatedText(searchPos) = " "c Then
                                    endPos = searchPos + 1  ' Include the space
                                    foundSpace = True
                                    Exit For
                                End If
                            Next

                            ' If not found forward, search backward
                            If Not foundSpace Then
                                For searchPos As Integer = endPos - 1 To currentPos Step -1
                                    If translatedText(searchPos) = " "c Then
                                        endPos = searchPos + 1  ' Include the space
                                        foundSpace = True
                                        Exit For
                                    End If
                                Next
                            End If

                            ' Extend forward to include the entire current word (until next space or end)
                            If Not foundSpace Then
                                ' Find the end of the current word
                                Dim wordEnd As Integer = endPos
                                While wordEnd < translatedLength AndAlso translatedText(wordEnd) <> " "c
                                    wordEnd += 1
                                End While
                                ' Include the space after the word if present
                                If wordEnd < translatedLength AndAlso translatedText(wordEnd) = " "c Then
                                    endPos = wordEnd + 1
                                Else
                                    endPos = wordEnd
                                End If
                            End If
                        End If
                    End If

                    endPos = Math.Min(endPos, translatedLength)
                    Dim runText As String = translatedText.Substring(currentPos, endPos - currentPos)
                    SetTextNodeWithSpacePreserve(run.TextNode, runText)
                    currentPos = endPos
                End If
            Next

            ' Final pass removed - it was incorrectly adding spaces between runs
            ' that were legitimately split mid-formatting (e.g., bold changes within a word)
            ' The fix above ensures we never split mid-word, so this is no longer needed
        Next
    End Sub

    ''' <summary>
    ''' Sets the text content of a WordprocessingML <c>w:t</c> node and ensures
    ''' <c>xml:space="preserve"</c> is present when leading/trailing/multiple spaces exist,
    ''' preventing Word from trimming whitespace on load.
    ''' </summary>
    ''' <param name="textNode">The <c>w:t</c> node to modify.</param>
    ''' <param name="text">The new text to assign.</param>

    Private Sub SetTextNodeWithSpacePreserve(textNode As System.Xml.XmlNode, text As String)
        textNode.InnerText = text

        ' If text has leading or trailing space, we MUST set xml:space="preserve"
        ' otherwise Word will trim the whitespace when loading the document
        If text.Length > 0 AndAlso (text.StartsWith(" ") OrElse text.EndsWith(" ") OrElse text.Contains("  ")) Then
            Dim xmlSpaceAttr = textNode.Attributes("xml:space")
            If xmlSpaceAttr Is Nothing Then
                xmlSpaceAttr = textNode.OwnerDocument.CreateAttribute("xml", "space", "http://www.w3.org/XML/1998/namespace")
                textNode.Attributes.Append(xmlSpaceAttr)
            End If
            xmlSpaceAttr.Value = "preserve"
        End If
    End Sub

    ''' <summary>
    ''' Processes footnotes.xml and endnotes.xml to translate or correct their text.
    ''' </summary>
    Private Async Function ProcessFootnotesEndnotes(tempDir As String, targetLanguage As String, mode As DocumentProcessMode, mainFileName As String) As System.Threading.Tasks.Task
        Dim wordDir As String = Path.Combine(tempDir, "word")
        If Not Directory.Exists(wordDir) Then Exit Function

        ' Footnotes and Endnotes files
        Dim files As String() = {"footnotes.xml", "endnotes.xml"}

        For Each fileName In files
            Dim filePath As String = Path.Combine(wordDir, fileName)
            If Not File.Exists(filePath) Then Continue For

            Try
                Dim xmlDoc As New System.Xml.XmlDocument()
                xmlDoc.PreserveWhitespace = True
                xmlDoc.Load(filePath)

                Dim nsMgr As New System.Xml.XmlNamespaceManager(xmlDoc.NameTable)
                nsMgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

                Dim paragraphs As List(Of TranslateParagraphInfo) = ExtractTranslateParagraphsFromXml(xmlDoc, nsMgr)

                Dim processableParagraphs = paragraphs.Where(Function(p) Not p.IsEmpty).ToList()
                If processableParagraphs.Count = 0 Then Continue For

                Dim componentType As String = If(fileName = "footnotes.xml", "Footnotes", "Endnotes")
                Dim success As Boolean = Await ProcessParagraphBatches(paragraphs, targetLanguage, mode, $"{mainFileName} ({componentType})")
                If success Then
                    ApplyTranslationsToXml(paragraphs)
                    xmlDoc.Save(filePath)
                End If

            Catch ex As Exception
                Debug.WriteLine($"ProcessFootnotesEndnotes error for {fileName}: {ex.Message}")
                ' Continue with other files
            End Try
        Next
    End Function
End Class