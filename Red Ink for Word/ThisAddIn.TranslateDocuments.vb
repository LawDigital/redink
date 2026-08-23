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
'  - Context Windows: Each independent LLM batch includes a small window of already-processed
'    preceding paragraphs and unprocessed following paragraphs. The model is explicitly told
'    that it does not see or remember the rest of the document; global instructions/dictionaries
'    remain authoritative across all batches.
'  - Pure Text to LLM: Only plain, visible text is sent to the LLM—no XML, no
'    formatting codes, and no markup—ensuring maximum formatting preservation.
'  - Formatting Preservation (default): Inserts | markers only at
'    real formatting-group boundaries (not every internal Word run split). The LLM
'    preserves those boundaries, and text is redistributed only within runs that
'    have identical formatting. Manual Word breaks are protected as structural anchors.
'  - Correction Mode: When correcting, creates a Word compare document showing
'    all changes between original and corrected versions.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.IO.Compression
Imports System.Runtime.InteropServices

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
    ''' The marker character inserted at run boundaries when formatting-aware mode is enabled.
    ''' </summary>
    Private Const RunBoundaryMarker As String = "|"

    ''' <summary>
    ''' Marker inserted at protected structural boundaries in text sent to the LLM.
    ''' U+2016 DOUBLE VERTICAL LINE — virtually never appears in legal documents.
    ''' </summary>
    Private Const NoteRefMarker As String = "‖"

    ''' <summary>
    ''' Represents a non-breaking space occurrence with surrounding word context for restoration.
    ''' </summary>
    Private Class NonBreakingSpaceInfo
        ''' <summary>The non-breaking space character (U+00A0, U+202F, etc.).</summary>
        Public Property SpaceChar As Char
        ''' <summary>The word immediately before the non-breaking space (Nothing if at start).</summary>
        Public Property WordBefore As String
        ''' <summary>The word immediately after the non-breaking space (Nothing if at end).</summary>
        Public Property WordAfter As String
    End Class

    ''' <summary>
    ''' Characters treated as non-breaking spaces that must be preserved.
    ''' U+00A0 = non-breaking space (geschütztes Leerzeichen)
    ''' U+202F = narrow no-break space (schmales geschütztes Leerzeichen)
    ''' U+2007 = figure space (ziffernbreites Leerzeichen)
    ''' </summary>
    Private Shared ReadOnly NonBreakingSpaceChars As Char() = {ChrW(&HA0), ChrW(&H202F), ChrW(&H2007)}


    ''' <summary>
    ''' Represents a text run (w:t element) with its content and XML reference.
    ''' </summary>
    Private Class TranslateTextRunInfo
        Public Property TextNode As System.Xml.XmlNode
        Public Property OriginalText As String
        Public Property HasNoteReferenceBefore As Boolean
        Public Property FormattingKey As String
    End Class

    ''' <summary>
    ''' Represents a paragraph with its text runs for translation.
    ''' </summary>
    Private Class TranslateParagraphInfo
        Public Property Index As Integer
        Public Property TextRuns As List(Of TranslateTextRunInfo)
        Public Property FullText As String  ' Combined text from all runs (plain text only)
        Public Property MarkerText As String ' Text with | markers at real formatting boundaries
        Public Property TranslatedText As String
        Public Property IsEmpty As Boolean
    End Class


    ''' <summary>
    ''' Returns a stable key for run formatting and structural scope. Adjacent runs with
    ''' the same key can be treated as one formatting group even if Word split them internally.
    ''' </summary>
    Private Shared Function GetTranslateRunFormattingKey(textNode As System.Xml.XmlNode) As String
        If textNode Is Nothing Then Return ""

        Dim runNode As System.Xml.XmlNode = textNode.ParentNode
        While runNode IsNot Nothing AndAlso runNode.LocalName <> "r"
            runNode = runNode.ParentNode
        End While
        If runNode Is Nothing Then Return ""

        Dim sb As New System.Text.StringBuilder()

        For Each child As System.Xml.XmlNode In runNode.ChildNodes
            If child.NodeType = System.Xml.XmlNodeType.Element AndAlso child.LocalName = "rPr" Then
                sb.Append(child.OuterXml)
                Exit For
            End If
        Next

        ' Never merge across structural wrappers such as hyperlinks, tracked changes,
        ' content controls, etc., even when visible formatting is identical.
        Dim ancestor As System.Xml.XmlNode = runNode.ParentNode
        While ancestor IsNot Nothing AndAlso ancestor.LocalName <> "p"
            If ancestor.NodeType = System.Xml.XmlNodeType.Element Then
                sb.Append("|scope:")
                sb.Append(ancestor.LocalName)
                sb.Append(":")
                sb.Append(System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(ancestor).ToString())
            End If
            ancestor = ancestor.ParentNode
        End While

        Return sb.ToString()
    End Function

    Private Shared Function BuildTranslateFormattingGroups(
        textRuns As List(Of TranslateTextRunInfo),
        Optional runIndices As List(Of Integer) = Nothing) As List(Of List(Of Integer))

        Dim groups As New List(Of List(Of Integer))()
        If textRuns Is Nothing OrElse textRuns.Count = 0 Then Return groups

        Dim indices As New List(Of Integer)()
        If runIndices Is Nothing Then
            For i As Integer = 0 To textRuns.Count - 1
                indices.Add(i)
            Next
        Else
            indices.AddRange(runIndices)
        End If

        Dim currentGroup As List(Of Integer) = Nothing
        Dim previousKey As String = Nothing
        Dim havePrevious As Boolean = False

        For Each runIdx As Integer In indices
            If runIdx < 0 OrElse runIdx >= textRuns.Count Then Continue For
            Dim run As TranslateTextRunInfo = textRuns(runIdx)
            If run Is Nothing OrElse String.IsNullOrEmpty(run.OriginalText) Then Continue For

            Dim currentKey As String = If(run.FormattingKey, "")
            If Not havePrevious OrElse Not String.Equals(previousKey, currentKey, StringComparison.Ordinal) Then
                currentGroup = New List(Of Integer)()
                groups.Add(currentGroup)
            End If

            currentGroup.Add(runIdx)
            previousKey = currentKey
            havePrevious = True
        Next

        Return groups
    End Function


    Private Shared Function IsRemoteDocumentPath(documentPath As String) As Boolean
        If String.IsNullOrWhiteSpace(documentPath) Then
            Return False
        End If

        Dim uri As Uri = Nothing
        If Uri.TryCreate(documentPath, UriKind.Absolute, uri) Then
            Return Not uri.IsFile
        End If

        Return documentPath.StartsWith("http://", StringComparison.OrdinalIgnoreCase) OrElse
           documentPath.StartsWith("https://", StringComparison.OrdinalIgnoreCase)
    End Function

    Private Shared Function CreateDocumentCopyFromSavedSource(sourcePath As String, destinationPath As String, fileFormat As WdSaveFormat) As Boolean
        ' Local, saved source: plain file copy — no Word interop needed at all.
        If Not IsRemoteDocumentPath(sourcePath) AndAlso File.Exists(sourcePath) Then
            File.Copy(sourcePath, destinationPath, overwrite:=True)
            Return File.Exists(destinationPath)
        End If

        ' Remote source (e.g., SharePoint): materialize a local copy through Word using SaveAs2.
        ' Prefer REUSING the already-running primary Word instance so we do NOT start a second
        ' Word process (which would load any document-management add-in a second time — a known
        ' crash / re-entrancy vector). Only if reuse is not possible do we fall back to an
        ' isolated secondary instance.
        If TryCreateRemoteCopyViaPrimaryInstance(sourcePath, destinationPath, fileFormat) Then
            Return True
        End If

        Return CreateRemoteCopyViaSecondaryInstance(sourcePath, destinationPath, fileFormat)
    End Function

    ''' <summary>
    ''' Attempts to create a local copy of a remote (e.g., SharePoint) document by reusing the
    ''' primary Word instance. Only proceeds if Word actually opens a NEW document; if the source
    ''' was already open (e.g., it is the active document), the returned reference points at the
    ''' live document and must not be overwritten via SaveAs2, so this returns False to let the
    ''' caller fall back to an isolated secondary instance.
    ''' </summary>
    Private Shared Function TryCreateRemoteCopyViaPrimaryInstance(sourcePath As String, destinationPath As String, fileFormat As WdSaveFormat) As Boolean
        Dim wordApp As Word.Application = Nothing
        Dim openedDoc As Word.Document = Nothing

        Try
            wordApp = Globals.ThisAddIn.Application
            If wordApp Is Nothing OrElse wordApp.Documents Is Nothing Then Return False

            Dim countBefore As Integer = wordApp.Documents.Count

            openedDoc = wordApp.Documents.Open(
                FileName:=sourcePath,
                ReadOnly:=True,
                Visible:=False,
                AddToRecentFiles:=False)

            ' If the document count did not grow, Word returned an already-open document
            ' (the live active document). Do NOT SaveAs2 over it and do NOT close it.
            If wordApp.Documents.Count <= countBefore Then
                openedDoc = Nothing
                Return False
            End If

            openedDoc.SaveAs2(
                FileName:=destinationPath,
                FileFormat:=fileFormat,
                AddToRecentFiles:=False)

            Return File.Exists(destinationPath)

        Catch
            Return False
        Finally
            If openedDoc IsNot Nothing Then
                Try : openedDoc.Close(WdSaveOptions.wdDoNotSaveChanges) : Catch : End Try
                Try : Marshal.FinalReleaseComObject(openedDoc) : Catch : End Try
            End If
        End Try
    End Function

    ''' <summary>
    ''' Fallback that creates a local copy of a remote document in an isolated secondary Word
    ''' instance. Used only when the primary instance cannot be reused safely (e.g., the remote
    ''' source is already open as the active document).
    ''' </summary>
    Private Shared Function CreateRemoteCopyViaSecondaryInstance(sourcePath As String, destinationPath As String, fileFormat As WdSaveFormat) As Boolean
        Dim secondaryWordApp As Word.Application = Nothing
        Dim secondaryDoc As Word.Document = Nothing

        Try
            secondaryWordApp = New Word.Application With {
            .Visible = False,
            .DisplayAlerts = WdAlertLevel.wdAlertsNone
        }

            secondaryDoc = secondaryWordApp.Documents.Open(
            FileName:=sourcePath,
            ReadOnly:=True,
            Visible:=False,
            AddToRecentFiles:=False)

            secondaryDoc.SaveAs2(
            FileName:=destinationPath,
            FileFormat:=fileFormat,
            AddToRecentFiles:=False)

            Return File.Exists(destinationPath)

        Finally
            If secondaryDoc IsNot Nothing Then
                Try : secondaryDoc.Close(WdSaveOptions.wdDoNotSaveChanges) : Catch : End Try
                Try : Marshal.FinalReleaseComObject(secondaryDoc) : Catch : End Try
            End If

            If secondaryWordApp IsNot Nothing Then
                Try : secondaryWordApp.Quit(WdSaveOptions.wdDoNotSaveChanges) : Catch : End Try
                Try : Marshal.FinalReleaseComObject(secondaryWordApp) : Catch : End Try
            End If
        End Try
    End Function

    ''' <summary>
    ''' Creates a temporary working copy of the currently active Word document for the document-processing pipeline.
    ''' If the document has unsaved changes, the user can save first, use the last saved version, or cancel.
    ''' </summary>
    Friend Function TryCreateActiveDocumentProcessingCopy(ByRef tempCopyPath As String) As Boolean
        Dim wordApp As Word.Application = Nothing
        Dim activeDoc As Word.Document = Nothing
        Dim tempFolder As String = Nothing

        Try
            wordApp = Globals.ThisAddIn.Application
            If wordApp Is Nothing OrElse wordApp.Documents Is Nothing OrElse wordApp.Documents.Count = 0 Then
                ShowCustomMessageBox("No Word document is currently open.")
                Return False
            End If

            activeDoc = wordApp.ActiveDocument
            If activeDoc Is Nothing Then
                ShowCustomMessageBox("No active Word document was found.")
                Return False
            End If

            Dim sourceName As String = activeDoc.Name
            If String.IsNullOrWhiteSpace(sourceName) Then
                ShowCustomMessageBox("The active document could not be identified.")
                Return False
            End If

            If String.IsNullOrWhiteSpace(activeDoc.Path) Then
                ShowCustomMessageBox("The active document has never been saved. Please save it first, then try again.")
                Return False
            End If

            Dim ext As String = Path.GetExtension(sourceName).ToLowerInvariant()
            If ext <> ".doc" AndAlso ext <> ".docx" Then
                ShowCustomMessageBox($"The active document type '{ext}' is not supported for this workflow.")
                Return False
            End If

            tempFolder = Path.Combine(Path.GetTempPath(), $"{AN2}_active_doc_{Guid.NewGuid():N}")
            Directory.CreateDirectory(tempFolder)
            tempCopyPath = Path.Combine(tempFolder, Path.GetFileName(sourceName))

            If Not activeDoc.Saved Then
                Dim saveChoice As Integer = ShowCustomYesNoBox(
                "The active Word document has unsaved changes." & vbCrLf & vbCrLf &
                "How would you like to continue?",
                "Save current version and continue",
                "Use last saved version",
                AN & " Active Document",
                extraButtonText:="Cancel",
                extraButtonAction:=Sub()
                                   End Sub,
                CloseAfterExtra:=True)

                If saveChoice = 0 Then
                    tempCopyPath = Nothing
                    Return False
                End If

                If saveChoice = 1 Then
                    activeDoc.Save()
                End If
            End If

            Dim targetFormat As WdSaveFormat =
            If(ext = ".doc", WdSaveFormat.wdFormatDocument97, WdSaveFormat.wdFormatXMLDocument)

            If Not CreateDocumentCopyFromSavedSource(activeDoc.FullName, tempCopyPath, targetFormat) Then
                Throw New IOException("The temporary working copy could not be created.")
            End If

            Return True

        Catch ex As Exception
            tempCopyPath = Nothing

            If Not String.IsNullOrWhiteSpace(tempFolder) AndAlso Directory.Exists(tempFolder) Then
                Try : Directory.Delete(tempFolder, recursive:=True) : Catch : End Try
            End If

            ShowCustomMessageBox($"Could not prepare the active document: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Returns a non-conflicting file path in the specified directory.
    ''' If the requested filename already exists, appends " (n)" before the extension.
    ''' </summary>
    Private Shared Function GetNonConflictingFilePath(directoryPath As String, fileName As String) As String
        Dim candidatePath As String = Path.Combine(directoryPath, fileName)
        If Not File.Exists(candidatePath) Then Return candidatePath

        Dim baseName As String = Path.GetFileNameWithoutExtension(fileName)
        Dim ext As String = Path.GetExtension(fileName)
        Dim counter As Integer = 1

        Do
            candidatePath = Path.Combine(directoryPath, $"{baseName} ({counter}){ext}")
            If Not File.Exists(candidatePath) Then Return candidatePath
            counter += 1
        Loop
    End Function

    ''' <summary>
    ''' Moves a file to the current user's Desktop using a non-conflicting filename if needed.
    ''' </summary>
    Private Shared Function MoveFileToDesktop(sourcePath As String) As String
        If String.IsNullOrWhiteSpace(sourcePath) Then Return Nothing
        If Not File.Exists(sourcePath) Then
            Throw New FileNotFoundException("The source file to move was not found.", sourcePath)
        End If

        Dim desktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
        Dim destinationPath As String = GetNonConflictingFilePath(desktopPath, Path.GetFileName(sourcePath))

        File.Move(sourcePath, destinationPath)
        Return destinationPath
    End Function


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
    Private _isFreestyle As Boolean = False

    ''' <summary>
    ''' Whether to use formatting-aware marker mode for the current processing run.
    ''' </summary>
    Private _useFormattingMarkers As Boolean = False

    ''' <summary>
    ''' Entry point for correction (default prompt/suffix).
    ''' </summary>
    Public Async Sub CorrectWordDocuments()
        Await CorrectWordDocuments(Nothing, Nothing, False, False)
    End Sub

    ''' <summary>
    ''' Entry point for anonymization (default prompt/suffix).
    ''' </summary>
    Public Async Sub AnonymizeWordDocuments()
        OtherPromptUnfilled = ""
        Await CorrectWordDocuments(SP_Anonymize_Document, "_anonymized", False, True)
    End Sub

    ''' <summary>
    ''' Entry point for switching parties (default prompt/suffix).
    ''' </summary>
    Public Async Sub SwitchPartiesDocuments()
        OtherPromptUnfilled = ""
        Await CorrectWordDocuments(InterpolateAtRuntime(SP_SwitchParty_Document), "_newparties", False, True)
    End Sub


    ''' <summary>
    ''' Entry point for correction with optional prompt/suffix overrides.
    ''' </summary>
    ''' <summary>
    ''' Entry point for correction with optional prompt/suffix overrides.
    ''' </summary>
    Public Async Function CorrectWordDocuments(Optional promptOverride As String = Nothing,
                                          Optional correctedSuffixOverride As String = Nothing,
                                          Optional UseSecondAPI As Boolean = False,
                                          Optional IsFreestyle As Boolean = False) As Threading.Tasks.Task

        _correctPromptOverride = If(String.IsNullOrWhiteSpace(promptOverride), Nothing, promptOverride)
        _useSecondAPI = UseSecondAPI
        _isFreestyle = IsFreestyle

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
        _isFreestyle = False

        If UseSecondAPI And originalConfigLoaded Then
            RestoreDefaults(_context, originalConfig)
            originalConfigLoaded = False
        End If


    End Function


    ''' <summary>
    ''' Main processing method that handles both translation and correction modes.
    ''' Supports both Word (.doc/.docx) and PowerPoint (.pptx) files with automatic dispatch.
    ''' </summary>
    ''' <param name="mode">The processing mode (Translate or Correct).</param>
    ''' <summary>
    ''' Main processing method that handles both translation and correction modes.
    ''' Supports both Word (.doc/.docx) and PowerPoint (.pptx) files with automatic dispatch.
    ''' </summary>
    ''' <param name="mode">The processing mode (Translate or Correct).</param>
    Private Async Function ProcessWordDocuments(mode As DocumentProcessMode) As System.Threading.Tasks.Task
        Dim selectedPath As String = ""
        Dim usedActiveDocument As Boolean = False
        Dim activeDocumentTempFolder As String = Nothing
        Dim keepActiveDocumentTempFolder As Boolean = False
        Dim resultFilePaths As New List(Of String)()

        Dim modeVerb As String = If(mode = DocumentProcessMode.Translate, "translate", If(_isFreestyle, "adapt", "correct"))
        Dim modeVerbPast As String = If(mode = DocumentProcessMode.Translate, "translated", If(_isFreestyle, "adapted", "corrected"))
        Dim modeVerbGerund As String = If(mode = DocumentProcessMode.Translate, "Translating", If(_isFreestyle, "Adapting", "Correcting"))
        Dim modeNoun As String = If(mode = DocumentProcessMode.Translate, "Translation", If(_isFreestyle, "Adaption", "Correction"))

        ' Effective correction suffix (used throughout for correction mode)
        Dim effectiveCorrectedSuffix As String = If(String.IsNullOrWhiteSpace(_correctSuffixOverride), CorrectedFileSuffix, _correctSuffixOverride)

        If INI_AllowLegacyDocFiles Then
            Globals.ThisAddIn.DragDropFormLabel = $"Select a Word or PowerPoint document or folder to {modeVerb}, or use the currently active Word document"
            Globals.ThisAddIn.DragDropFormFilter = "Supported Documents|*.doc;*.docx;*.pptx|Word Documents|*.doc;*.docx|Word Document (*.docx)|*.docx|Word 97-2003 (*.doc)|*.doc|PowerPoint (*.pptx)|*.pptx"
        Else
            Globals.ThisAddIn.DragDropFormLabel = $"Select a Word or PowerPoint document or folder to {modeVerb}, or use the currently active Word document"
            Globals.ThisAddIn.DragDropFormFilter = "Supported Documents|*.docx;*.pptx|Word Document (*.docx)|*.docx|PowerPoint (*.pptx)|*.pptx"
        End If

        Try
            Using frm As New DragDropForm(DragDropMode.FileOrDirectory, allowUseActiveDocument:=True)
                Dim __safeDialogOwner507 As System.Windows.Forms.IWin32Window = SharedLibrary.SharedLibrary.SharedMethods.ResolveSameThreadDialogOwner()
                If If(__safeDialogOwner507 IsNot Nothing, frm.ShowDialog(__safeDialogOwner507), frm.ShowDialog()) = DialogResult.OK Then
                    selectedPath = frm.SelectedFilePath
                    usedActiveDocument = frm.UsedActiveDocument

                    If usedActiveDocument AndAlso Not String.IsNullOrWhiteSpace(selectedPath) Then
                        activeDocumentTempFolder = Path.GetDirectoryName(selectedPath)
                    End If
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
        Dim supportedExtensions As String() = If(INI_AllowLegacyDocFiles,
                                                  {".doc", ".docx", ".pptx"},
                                                  {".docx", ".pptx"})

        If isFile Then
            Dim ext As String = Path.GetExtension(selectedPath).ToLowerInvariant()
            If supportedExtensions.Contains(ext) Then
                filesToProcess.Add(selectedPath)
            Else
                ShowCustomMessageBox($"File type '{ext}' is not supported.")
                Exit Function
            End If
        Else
            Dim recurseChoice As Integer = ShowCustomYesNoBox(
                $"Include documents from subdirectories?",
                "Yes, include subdirectories", "No, top directory only")
            If recurseChoice = 0 Then Exit Function

            Dim searchOption As SearchOption = If(recurseChoice = 1, SearchOption.AllDirectories, SearchOption.TopDirectoryOnly)

            Dim allFiles = Directory.GetFiles(selectedPath, "*.*", searchOption)
            For Each f In allFiles
                Dim ext As String = Path.GetExtension(f).ToLowerInvariant()
                If supportedExtensions.Contains(ext) Then
                    filesToProcess.Add(f)
                End If
            Next

            If filesToProcess.Count = 0 Then
                ShowCustomMessageBox("No Word or PowerPoint documents found.")
                Exit Function
            End If
        End If

        ' Get target language (only for translation mode)
        Dim targetLanguage As String = Nothing
        Dim targetLanguageToken As String = Nothing

        If mode = DocumentProcessMode.Translate Then
            Dim defaultLanguage As String = If(String.IsNullOrWhiteSpace(INI_Language1), "English", INI_Language1)
            targetLanguage = Global.SharedLibrary.SharedLibrary.SharedMethods.PromptForTargetLanguage(
                "Enter your target language (e.g., English, German, French):",
                AN & " Translate Files",
                defaultLanguage,
                _context)

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
            If Not ext.Equals(".doc", StringComparison.OrdinalIgnoreCase) AndAlso
               Not ext.Equals(".docx", StringComparison.OrdinalIgnoreCase) AndAlso
               Not ext.Equals(".pptx", StringComparison.OrdinalIgnoreCase) Then Continue For

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

        ' Formatting preservation is the default for interactive processing.
        ' The reduced marker model exposes only real formatting transitions, so there is no
        ' separate user choice and no additional LLM call.
        _useFormattingMarkers = True

        ' Confirm if many files to process
        If filesToProcess.Count > 10 Then
            ' Count file types for informative message
            Dim wordCount As Integer = filesToProcess.Where(Function(f) Not Path.GetExtension(f).Equals(".pptx", StringComparison.OrdinalIgnoreCase)).Count
            Dim pptxCount As Integer = filesToProcess.Where(Function(f) Path.GetExtension(f).Equals(".pptx", StringComparison.OrdinalIgnoreCase)).Count
            Dim fileDesc As String = $"{filesToProcess.Count} document(s)"
            If wordCount > 0 AndAlso pptxCount > 0 Then
                fileDesc = $"{filesToProcess.Count} document(s) ({wordCount} Word, {pptxCount} PowerPoint)"
            End If

            Dim confirmMsg As String = If(mode = DocumentProcessMode.Translate,
                $"Ready to translate {fileDesc} to {targetLanguage}. Continue?",
                $"Ready to correct {fileDesc}. Continue?")
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
                Dim fileExt As String = Path.GetExtension(filePath).ToLowerInvariant()
                Dim isPptx As Boolean = (fileExt = ".pptx")

                ProgressBarModule.GlobalProgressValue = i
                ProgressBarModule.GlobalProgressLabel = $"{modeVerbGerund} {i + 1}/{filesToProcess.Count}: {fileName}"

                Try
                    Dim dir As String = Path.GetDirectoryName(filePath)
                    Dim nameWithoutExt As String = Path.GetFileNameWithoutExtension(filePath)
                    Dim outputExt As String = If(isPptx, ".pptx", ".docx")
                    Dim outputPath As String
                    Dim compareOutputPath As String = Nothing

                    If mode = DocumentProcessMode.Translate Then
                        outputPath = Path.Combine(dir, $"{nameWithoutExt}_{targetLanguageToken}{outputExt}")
                    Else
                        outputPath = Path.Combine(dir, $"{nameWithoutExt}{effectiveCorrectedSuffix}{outputExt}")
                        If Not isPptx Then
                            compareOutputPath = GetCompareFilePath(outputPath)
                        End If
                    End If

                    Dim success As Boolean

                    If isPptx Then
                        ' PowerPoint: copy then process via PPTX OpenXML pipeline
                        File.Copy(filePath, outputPath, overwrite:=True)
                        success = Await ProcessPptxOpenXml(outputPath, targetLanguage, mode)
                    Else
                        ' Word: existing pipeline (handles .doc → .docx conversion)
                        success = Await ProcessDocumentViaOpenXml(filePath, outputPath, targetLanguage, mode)
                    End If

                    If success Then
                        Dim compareCreated As Boolean = False

                        ' For correction mode on Word docs, create compare document.
                        If mode = DocumentProcessMode.Correct AndAlso Not isPptx Then
                            compareCreated = CreateWordCompareDocument(filePath, outputPath)
                            If Not compareCreated Then
                                failedFiles.Add($"{fileName}: Corrected but compare document creation failed")
                            End If
                        End If

                        Dim desktopMoveSucceeded As Boolean = True
                        Dim finalOutputPath As String = outputPath
                        Dim finalComparePath As String = If(compareCreated, compareOutputPath, Nothing)

                        If usedActiveDocument Then
                            Try
                                Dim movedOutputPath As String = MoveFileToDesktop(outputPath)
                                If Not String.IsNullOrWhiteSpace(movedOutputPath) Then
                                    finalOutputPath = movedOutputPath
                                End If

                                If compareCreated AndAlso Not String.IsNullOrWhiteSpace(compareOutputPath) AndAlso File.Exists(compareOutputPath) Then
                                    Dim movedComparePath As String = MoveFileToDesktop(compareOutputPath)
                                    If Not String.IsNullOrWhiteSpace(movedComparePath) Then
                                        finalComparePath = movedComparePath
                                    End If
                                End If
                            Catch ex As Exception
                                desktopMoveSucceeded = False
                                keepActiveDocumentTempFolder = True
                                failedFiles.Add($"{fileName}: Processed but moving result(s) to the Desktop failed - {ex.Message}")
                            End Try
                        End If

                        If desktopMoveSucceeded Then
                            If Not String.IsNullOrWhiteSpace(finalOutputPath) AndAlso File.Exists(finalOutputPath) Then
                                resultFilePaths.Add(finalOutputPath)
                            End If
                            If Not String.IsNullOrWhiteSpace(finalComparePath) AndAlso File.Exists(finalComparePath) Then
                                resultFilePaths.Add(finalComparePath)
                            End If
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
            _useFormattingMarkers = False
        End Try

        If usedActiveDocument AndAlso
           Not keepActiveDocumentTempFolder AndAlso
           Not String.IsNullOrWhiteSpace(activeDocumentTempFolder) AndAlso
           Directory.Exists(activeDocumentTempFolder) Then
            Try
                Directory.Delete(activeDocumentTempFolder, recursive:=True)
            Catch
            End Try
        End If

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
            ' Only mention compare documents if any Word files were processed
            Dim hasWordFiles As Boolean = filesToProcess.Any(Function(f) Not Path.GetExtension(f).Equals(".pptx", StringComparison.OrdinalIgnoreCase))
            If hasWordFiles Then
                summary.AppendLine($"Compare documents created with tracked changes")
            End If
        End If

        Dim distinctResultPaths As List(Of String) = resultFilePaths.
            Where(Function(p) Not String.IsNullOrWhiteSpace(p)).
            Distinct(StringComparer.OrdinalIgnoreCase).
            ToList()

        If distinctResultPaths.Count > 0 Then
            summary.AppendLine()
            summary.AppendLine(If(usedActiveDocument, "Result files saved to Desktop:", "Created result files:"))
            For Each resultPath In distinctResultPaths.Take(20)
                summary.AppendLine($"  • {Path.GetFileName(resultPath)}")
            Next
            If distinctResultPaths.Count > 20 Then
                summary.AppendLine($"  ... and {distinctResultPaths.Count - 20} more")
            End If
        End If

        If usedActiveDocument AndAlso keepActiveDocumentTempFolder AndAlso Not String.IsNullOrWhiteSpace(activeDocumentTempFolder) Then
            summary.AppendLine()
            summary.AppendLine("Temporary files were kept because moving the result failed:")
            summary.AppendLine($"  {activeDocumentTempFolder}")
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
        Dim tempCompareDir As String = Nothing

        Try
            ' The comparison owns its Word instance. Do not reuse Globals.ThisAddIn.Application:
            ' the hidden comparison process must be closed deterministically and must not leave
            ' comparison documents or state in the user's interactive Word instance.
            wordApp = New Word.Application()
            wordApp.Visible = False
            wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone
            wordApp.ScreenUpdating = False
            Try : wordApp.Options.UpdateLinksAtOpen = False : Catch : End Try

            ' Determine final compare output path.
            Dim comparePath As String = GetCompareFilePath(correctedPath)

            ' Work entirely on local copies inside a private %TEMP% folder so the compare
            ' interop never opens documents that live in a document-management or cloud-synced
            ' location.
            tempCompareDir = Path.Combine(Path.GetTempPath(), $"{AN2}_compare_{Guid.NewGuid():N}")
            Directory.CreateDirectory(tempCompareDir)

            Dim tempOriginalPath As String = Path.Combine(tempCompareDir, "original" & Path.GetExtension(originalPath))
            Dim tempCorrectedPath As String = Path.Combine(tempCompareDir, "corrected" & Path.GetExtension(correctedPath))
            Dim tempComparePath As String = Path.Combine(tempCompareDir, "compare.docx")

            File.Copy(originalPath, tempOriginalPath, overwrite:=True)
            File.Copy(correctedPath, tempCorrectedPath, overwrite:=True)

            originalDoc = wordApp.Documents.Open(
                FileName:=tempOriginalPath,
                ReadOnly:=True,
                Visible:=False,
                AddToRecentFiles:=False)

            correctedDoc = wordApp.Documents.Open(
                FileName:=tempCorrectedPath,
                ReadOnly:=True,
                Visible:=False,
                AddToRecentFiles:=False)

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
                RevisedAuthor:=GetMarkupAuthorOrCurrent(wordApp),
                IgnoreAllComparisonWarnings:=True
            )

            compareDoc.SaveAs2(tempComparePath, WdSaveFormat.wdFormatXMLDocument)

            compareDoc.Close(WdSaveOptions.wdDoNotSaveChanges)
            Try : Marshal.FinalReleaseComObject(compareDoc) : Catch : End Try
            compareDoc = Nothing

            correctedDoc.Close(WdSaveOptions.wdDoNotSaveChanges)
            Try : Marshal.FinalReleaseComObject(correctedDoc) : Catch : End Try
            correctedDoc = Nothing

            originalDoc.Close(WdSaveOptions.wdDoNotSaveChanges)
            Try : Marshal.FinalReleaseComObject(originalDoc) : Catch : End Try
            originalDoc = Nothing

            ' Move the finished compare file to its final destination with plain file I/O.
            File.Copy(tempComparePath, comparePath, overwrite:=True)

            Return File.Exists(comparePath)

        Catch ex As Exception
            Debug.WriteLine($"CreateWordCompareDocument error: {ex.Message}")
            Return False
        Finally
            If compareDoc IsNot Nothing Then
                Try : compareDoc.Close(WdSaveOptions.wdDoNotSaveChanges) : Catch : End Try
                Try : Marshal.FinalReleaseComObject(compareDoc) : Catch : End Try
            End If
            If correctedDoc IsNot Nothing Then
                Try : correctedDoc.Close(WdSaveOptions.wdDoNotSaveChanges) : Catch : End Try
                Try : Marshal.FinalReleaseComObject(correctedDoc) : Catch : End Try
            End If
            If originalDoc IsNot Nothing Then
                Try : originalDoc.Close(WdSaveOptions.wdDoNotSaveChanges) : Catch : End Try
                Try : Marshal.FinalReleaseComObject(originalDoc) : Catch : End Try
            End If
            If wordApp IsNot Nothing Then
                Try : wordApp.Quit(WdSaveOptions.wdDoNotSaveChanges) : Catch : End Try
                Try : Marshal.FinalReleaseComObject(wordApp) : Catch : End Try
            End If
            If Not String.IsNullOrWhiteSpace(tempCompareDir) AndAlso Directory.Exists(tempCompareDir) Then
                Try : Directory.Delete(tempCompareDir, recursive:=True) : Catch : End Try
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
            Dim success As Boolean = Await ProcessParagraphBatches(paragraphs, targetLanguage, mode, Path.GetFileName(docxPath), "main body")
            If Not success Then Return False

            ' Apply processed text back to XML nodes
            ApplyTranslationsToXml(paragraphs)

            ' Expand multi-line translations into separate XML paragraphs
            ExpandMultiLineParagraphs(paragraphs, nsMgr)

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
            Dim success As Boolean = Await ProcessParagraphBatches(paragraphs, targetLanguage, mode, $"{mainFileName} (Comments)", "comments")
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

                    Dim isHeader As Boolean = Path.GetFileName(filePath).StartsWith("header", StringComparison.OrdinalIgnoreCase)
                    Dim componentType As String = If(isHeader, "Headers", "Footers")
                    Dim contentScope As String = If(isHeader, "header", "footer")
                    Dim success As Boolean = Await ProcessParagraphBatches(paragraphs, targetLanguage, mode, $"{mainFileName} ({componentType})", contentScope)
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
    ''' When formatting markers are enabled, also builds marker-annotated text.
    ''' Detects footnote/endnote references, manual line breaks, AND complex field boundaries (cross-references,
    ''' merge fields, etc.) so that text redistribution never shifts content across them.
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
            .TranslatedText = Nothing,
            .MarkerText = Nothing
        }

            ' Build a set of w:t nodes that must be excluded from processing because
            ' they belong to or are immediately adjacent to footnote/endnote reference elements
            ' INSIDE footnotes.xml / endnotes.xml (w:footnoteRef / w:endnoteRef).
            Dim refNodes As New HashSet(Of System.Xml.XmlNode)()

            Dim noteRefElements As System.Xml.XmlNodeList = paraNode.SelectNodes(".//w:r[w:footnoteRef or w:endnoteRef]", nsMgr)
            For Each refRunNode As System.Xml.XmlNode In noteRefElements
                ' Skip any w:t inside the same run as the ref (usually none, but be safe)
                For Each tInRef As System.Xml.XmlNode In refRunNode.SelectNodes(".//w:t", nsMgr)
                    refNodes.Add(tInRef)
                Next
                ' Also skip the w:t in the immediately following sibling run if it is
                ' just whitespace (the separator space between the sign and the text)
                Dim nextSibling As System.Xml.XmlNode = refRunNode.NextSibling
                While nextSibling IsNot Nothing AndAlso nextSibling.NodeType <> System.Xml.XmlNodeType.Element
                    nextSibling = nextSibling.NextSibling
                End While
                If nextSibling IsNot Nothing AndAlso nextSibling.LocalName = "r" Then
                    Dim siblingTexts As System.Xml.XmlNodeList = nextSibling.SelectNodes(".//w:t", nsMgr)
                    If siblingTexts.Count = 1 Then
                        Dim sibText As String = siblingTexts(0).InnerText
                        If String.IsNullOrWhiteSpace(sibText) AndAlso sibText.Length <= 2 Then
                            refNodes.Add(siblingTexts(0))
                        End If
                    End If
                End If
            Next

            ' Build a set of w:r nodes that contain a w:footnoteReference or w:endnoteReference
            ' (document body references). These runs have no w:t but act as positional anchors.
            ' We need to know when a w:t run is preceded by such a reference so that redistribution
            ' does not shift text (especially spaces) across the reference boundary.
            Dim bodyRefRunNodes As New HashSet(Of System.Xml.XmlNode)()
            Dim bodyRefElements As System.Xml.XmlNodeList = paraNode.SelectNodes(".//w:r[w:footnoteReference or w:endnoteReference]", nsMgr)
            For Each refRunNode As System.Xml.XmlNode In bodyRefElements
                bodyRefRunNodes.Add(refRunNode)
            Next

            ' ─── Complex field detection ───
            ' Complex fields (cross-references, merge fields, etc.) use w:fldChar elements:
            '   <w:r><w:fldChar w:fldCharType="begin"/></w:r>
            '   <w:r><w:instrText>REF _Ref123 \h</w:instrText></w:r>
            '   <w:r><w:fldChar w:fldCharType="separate"/></w:r>
            '   <w:r><w:t>displayed value</w:t></w:r>     ← this w:t IS extracted
            '   <w:r><w:fldChar w:fldCharType="end"/></w:r>
            '
            ' The w:r containing fldChar begin/separate/end have no w:t and sit between
            ' text runs — exactly like footnote references. We treat them as boundaries
            ' so that text redistribution never shifts content across a field.
            '
            ' We also skip w:t nodes that are inside the field code region (between begin
            ' and separate) — these contain instrText-like content that should not be processed.
            ' Note: w:instrText uses a different element name and is already excluded by
            ' the ".//w:t" selector, but some generators put field codes in w:t nodes.
            Dim fldCharRuns As New HashSet(Of System.Xml.XmlNode)()
            Dim fldCharElements As System.Xml.XmlNodeList = paraNode.SelectNodes(".//w:r[w:fldChar]", nsMgr)
            For Each fldCharRunNode As System.Xml.XmlNode In fldCharElements
                fldCharRuns.Add(fldCharRunNode)
            Next

            ' Track field nesting to identify w:t nodes inside field code regions
            ' (between fldChar begin and fldChar separate) — these must be excluded.
            ' Also track w:t nodes inside field result regions (between separate and end)
            ' — these are displayed values that we extract but treat as bounded.
            If fldCharRuns.Count > 0 Then
                Dim fieldDepth As Integer = 0
                Dim inFieldCode As Boolean = False ' between begin and separate at depth 1+

                For Each childNode As System.Xml.XmlNode In paraNode.ChildNodes
                    If childNode.NodeType <> System.Xml.XmlNodeType.Element Then Continue For
                    If childNode.LocalName <> "r" Then Continue For

                    ' Check for fldChar in this run
                    Dim fldChar As System.Xml.XmlNode = childNode.SelectSingleNode("w:fldChar", nsMgr)
                    If fldChar IsNot Nothing Then
                        Dim fldType As String = ""
                        Dim fldTypeAttr = fldChar.Attributes("w:fldCharType")
                        If fldTypeAttr IsNot Nothing Then fldType = fldTypeAttr.Value

                        Select Case fldType
                            Case "begin"
                                fieldDepth += 1
                                inFieldCode = True
                            Case "separate"
                                inFieldCode = False
                            Case "end"
                                fieldDepth -= 1
                                If fieldDepth <= 0 Then
                                    fieldDepth = 0
                                    inFieldCode = False
                                End If
                        End Select
                        Continue For ' fldChar runs have no w:t to process
                    End If

                    ' If we're inside a field code region, exclude any w:t in this run
                    If inFieldCode Then
                        For Each tNode As System.Xml.XmlNode In childNode.SelectNodes(".//w:t", nsMgr)
                            refNodes.Add(tNode)
                        Next
                    End If
                Next
            End If

            ' Protect manual Word line breaks (w:br / w:cr) as structural boundaries.
            ' They are not text nodes, so without an explicit boundary the LLM sees both sides
            ' as one continuous string and redistribution can move words across the break.
            Dim manualBreakBeforeTextNodes As New HashSet(Of System.Xml.XmlNode)()
            Dim orderedTextAndBreakNodes As System.Xml.XmlNodeList =
                paraNode.SelectNodes(".//w:t | .//w:br | .//w:cr", nsMgr)
            Dim havePreviousProcessableText As Boolean = False
            Dim manualBreakPending As Boolean = False

            For Each orderedNode As System.Xml.XmlNode In orderedTextAndBreakNodes
                If orderedNode.LocalName = "br" OrElse orderedNode.LocalName = "cr" Then
                    If havePreviousProcessableText Then manualBreakPending = True
                    Continue For
                End If

                If orderedNode.LocalName = "t" Then
                    If refNodes.Contains(orderedNode) Then Continue For
                    If manualBreakPending Then
                        manualBreakBeforeTextNodes.Add(orderedNode)
                    End If
                    havePreviousProcessableText = True
                    manualBreakPending = False
                End If
            Next

            ' Find all w:t (text) elements within this paragraph
            Dim textNodes As System.Xml.XmlNodeList = paraNode.SelectNodes(".//w:t", nsMgr)
            Dim fullTextBuilder As New StringBuilder()

            For Each textNode As System.Xml.XmlNode In textNodes
                ' Skip text nodes that belong to the footnote/endnote reference sign or its separator,
                ' or that are inside a complex field code region
                If refNodes.Contains(textNode) Then Continue For

                Dim text As String = textNode.InnerText

                ' Determine if a footnoteReference/endnoteReference run OR a fldChar run
                ' exists between the previous w:t's w:r and this w:t's w:r.
                ' Walk backward from this run's parent through preceding sibling w:r elements.
                Dim hasNoteRefBefore As Boolean =
                    paraInfo.TextRuns.Count > 0 AndAlso manualBreakBeforeTextNodes.Contains(textNode)
                If paraInfo.TextRuns.Count > 0 AndAlso (bodyRefRunNodes.Count > 0 OrElse fldCharRuns.Count > 0) Then
                    Dim thisRun As System.Xml.XmlNode = textNode.ParentNode ' w:r containing this w:t
                    ' Walk backward through preceding sibling elements
                    Dim prevEl As System.Xml.XmlNode = thisRun.PreviousSibling
                    While prevEl IsNot Nothing
                        If prevEl.NodeType = System.Xml.XmlNodeType.Element Then
                            If prevEl.LocalName = "r" Then
                                If bodyRefRunNodes.Contains(prevEl) OrElse fldCharRuns.Contains(prevEl) Then
                                    ' Found a reference or field run between previous text run and this one
                                    hasNoteRefBefore = True
                                    Exit While
                                End If
                                ' Found a normal w:r — check if it contains a w:t that is in our TextRuns
                                Dim prevTexts As System.Xml.XmlNodeList = prevEl.SelectNodes(".//w:t", nsMgr)
                                Dim foundPrevTextRun As Boolean = False
                                For Each pt As System.Xml.XmlNode In prevTexts
                                    If Not refNodes.Contains(pt) Then
                                        foundPrevTextRun = True
                                        Exit For
                                    End If
                                Next
                                If foundPrevTextRun Then
                                    ' Reached the previous text-bearing run without hitting a reference
                                    Exit While
                                End If
                            End If
                        End If
                        prevEl = prevEl.PreviousSibling
                    End While
                End If

                If fullTextBuilder.Length > 0 AndAlso text.Length > 0 Then
                    Dim lastChar As Char = fullTextBuilder(fullTextBuilder.Length - 1)
                    Dim firstChar As Char = text(0)

                    If Not Char.IsWhiteSpace(lastChar) AndAlso Not Char.IsWhiteSpace(firstChar) Then
                        Dim xmlSpaceAttr = textNode.Attributes("xml:space")
                        Dim preserveSpace As Boolean = xmlSpaceAttr IsNot Nothing AndAlso xmlSpaceAttr.Value = "preserve"
                    End If
                End If

                paraInfo.TextRuns.Add(New TranslateTextRunInfo() With {
                    .TextNode = textNode,
                    .OriginalText = text,
                    .HasNoteReferenceBefore = hasNoteRefBefore,
                    .FormattingKey = GetTranslateRunFormattingKey(textNode)
                })
                ' Insert note-reference marker into FullText at the boundary
                If hasNoteRefBefore Then
                    fullTextBuilder.Append(NoteRefMarker)
                End If

                fullTextBuilder.Append(text)
            Next

            paraInfo.FullText = fullTextBuilder.ToString()
            paraInfo.IsEmpty = String.IsNullOrWhiteSpace(paraInfo.FullText)

            ' Build marker-annotated text when formatting markers are enabled
            If _useFormattingMarkers AndAlso Not paraInfo.IsEmpty AndAlso paraInfo.TextRuns.Count > 1 Then
                paraInfo.MarkerText = BuildMarkerAnnotatedText(paraInfo.TextRuns)
            End If

            paragraphs.Add(paraInfo)
            paraIndex += 1
        Next

        Return paragraphs
    End Function

    ''' <summary>
    ''' Builds a marker-annotated version of a paragraph's text by inserting | at run boundaries.
    ''' Only inserts markers between non-empty runs where formatting actually changes.
    ''' Also preserves ‖ markers at protected structural boundaries (notes, fields, manual breaks).
    ''' </summary>
    ''' <param name="textRuns">The text runs of the paragraph.</param>
    ''' <returns>The annotated text with | markers, or Nothing if only one effective run.</returns>
    Private Shared Function BuildMarkerAnnotatedText(textRuns As List(Of TranslateTextRunInfo)) As String
        If textRuns Is Nothing OrElse textRuns.Count <= 1 Then Return Nothing

        Dim sb As New StringBuilder()
        Dim markerCount As Integer = 0
        Dim previousFormattingKey As String = Nothing
        Dim havePreviousVisibleRun As Boolean = False
        Dim hardBoundarySincePreviousVisibleRun As Boolean = False

        For i As Integer = 0 To textRuns.Count - 1
            Dim run As TranslateTextRunInfo = textRuns(i)
            Dim runText As String = If(run.OriginalText, "")

            ' A note/field reference or manual break already is a hard positional boundary. Do not add
            ' a redundant formatting marker across it; markers inside each side still work.
            If run.HasNoteReferenceBefore Then
                sb.Append(NoteRefMarker)
                hardBoundarySincePreviousVisibleRun = True
            End If

            If runText.Length > 0 Then
                Dim currentFormattingKey As String = If(run.FormattingKey, "")

                If havePreviousVisibleRun AndAlso
                   Not hardBoundarySincePreviousVisibleRun AndAlso
                   Not String.Equals(previousFormattingKey, currentFormattingKey, StringComparison.Ordinal) Then
                    sb.Append(RunBoundaryMarker)
                    markerCount += 1
                End If

                sb.Append(runText)
                previousFormattingKey = currentFormattingKey
                havePreviousVisibleRun = True
                hardBoundarySincePreviousVisibleRun = False
            Else
                sb.Append(runText)
            End If
        Next

        ' No real formatting transition: do not expose Word's internal run fragmentation.
        If markerCount = 0 Then Return Nothing

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' Processes paragraphs in batches with context windows.
    ''' Sends ONLY plain text to the LLM - no XML, no formatting codes.
    ''' When formatting markers are enabled, includes | markers and instructs the LLM to preserve them.
    ''' </summary>
    Private Async Function ProcessParagraphBatches(
    paragraphs As List(Of TranslateParagraphInfo),
    targetLanguage As String,
    mode As DocumentProcessMode,
    Optional fileContext As String = "",
    Optional contentScope As String = "") As System.Threading.Tasks.Task(Of Boolean)

        Dim processableParagraphs = paragraphs.Where(Function(p) Not p.IsEmpty).ToList()
        If processableParagraphs.Count = 0 Then Return True

        Dim nbspRecords As New Dictionary(Of Integer, List(Of NonBreakingSpaceInfo))()
        For i As Integer = 0 To processableParagraphs.Count - 1
            Dim recorded = RecordNonBreakingSpaces(processableParagraphs(i).FullText)
            If recorded IsNot Nothing Then
                nbspRecords(i) = recorded
            End If
        Next

        Dim systemPrompt As String
        If mode = DocumentProcessMode.Translate Then
            TranslateLanguage = targetLanguage
            systemPrompt = InterpolateAtRuntime(SP_Translate_Document)
        Else
            Dim effectivePrompt As String = If(String.IsNullOrWhiteSpace(_correctPromptOverride), SP_Correct_Document, _correctPromptOverride)
            systemPrompt = InterpolateAtRuntime(effectivePrompt)
        End If

        systemPrompt = systemPrompt & vbCrLf & vbCrLf &
            "BATCH SCOPE / DOCUMENT-WIDE CONSISTENCY: This document is processed in independent sequential batches. " &
            "In this model call you do NOT have access to the full document and you do NOT retain hidden memory of earlier or later batches. " &
            "You see only the current [TEXTTOPROCESS] plus the small [CONTEXT BEFORE] and [CONTEXT AFTER] windows supplied with this call. " &
            "Treat the processing instruction, target language, translation dictionary, configured correction rules, and any user-supplied instruction as document-wide rules that apply consistently to every batch. " &
            "Use terminology and style choices visible in [CONTEXT BEFORE] as processed precedent unless a document-wide rule says otherwise; [CONTEXT AFTER] is unprocessed source context. " &
            "Do not invent facts, definitions, terminology choices, or other conventions based on parts of the document you cannot see."

        If _useFormattingMarkers Then
            systemPrompt = systemPrompt & vbCrLf & vbCrLf & SP_Add_Markers
        End If

        Dim hasNoteRefMarkers As Boolean = processableParagraphs.Any(Function(p) p.FullText.Contains(NoteRefMarker))
        If hasNoteRefMarkers Then
            systemPrompt = systemPrompt & vbCrLf & vbCrLf &
            "Some paragraphs contain the character ‖ (double vertical line). " &
            "This marks a protected structural boundary such as a footnote/endnote reference, field boundary, or manual Word line break. " &
            "CRITICAL: Keep each ‖ at EXACTLY the same position relative to the surrounding words. " &
            "Do NOT move, add, or remove any ‖ characters."
        End If

        If Not String.IsNullOrWhiteSpace(contentScope) Then
            systemPrompt = systemPrompt & vbCrLf & vbCrLf &
                $"Section context: The text to process comes from the document's {contentScope}. " &
                "Use this only as context for tone, conventions, and typical phrasing. " &
                "Do not add labels, headings, or explanations, and do not mention the section unless the source text already does."
        End If

        Dim batchIndex As Integer = 0
        Dim totalBatches As Integer = CInt(Math.Ceiling(processableParagraphs.Count / TranslateParagraphsPerBatch))
        Dim modeVerbGerund As String = If(mode = DocumentProcessMode.Translate, "Translating", "Correcting")

        ' Switch progress bar to batch-level granularity for this file
        ProgressBarModule.GlobalProgressMax = totalBatches
        ProgressBarModule.GlobalProgressValue = 0

        ' ── If not already using SecondAPI, try OfflineDocs alternate model ──
        Dim offlineDocsBackup As ModelConfig = Nothing
        Dim offlineDocsApplied As Boolean = False

        If Not _isFreestyle AndAlso Not _useSecondAPI AndAlso Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
            Try
                If GetSpecialTaskModel(_context, INI_AlternateModelPath, "OfflineDocs") Then
                    offlineDocsBackup = GetCurrentConfig(_context)
                    offlineDocsApplied = True
                    _useSecondAPI = True
                End If
            Catch
            End Try
        End If

        Try
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

                ' Context Before (already processed - plain text only, no markers)
                Dim contextBeforeStart As Integer = Math.Max(0, batchStart - TranslateContextBefore)
                If contextBeforeStart < batchStart Then
                    promptBuilder.AppendLine("[CONTEXT BEFORE - for reference only]")
                    For j As Integer = contextBeforeStart To batchStart - 1
                        Dim p = processableParagraphs(j)
                        ' Context uses plain translated text (no markers) for readability
                        Dim contextText As String = If(p.TranslatedText, p.FullText)
                        ' Strip any markers from translated text for clean context
                        If _useFormattingMarkers AndAlso contextText IsNot Nothing Then
                            contextText = contextText.Replace(RunBoundaryMarker, "")
                        End If
                        If contextText IsNot Nothing Then
                            contextText = contextText.Replace(NoteRefMarker, "")
                        End If
                        promptBuilder.AppendLine(contextText)
                    Next
                    promptBuilder.AppendLine()
                End If

                ' Paragraphs to process (plain text, or marker-annotated if enabled)
                promptBuilder.AppendLine("[TEXTTOPROCESS]")
                Dim batchNumber As Integer = 1
                For j As Integer = batchStart To batchEnd
                    ' Use marker-annotated text when available, otherwise plain text
                    Dim textToSend As String = If(_useFormattingMarkers AndAlso processableParagraphs(j).MarkerText IsNot Nothing,
                                                   processableParagraphs(j).MarkerText,
                                                   processableParagraphs(j).FullText)
                    promptBuilder.AppendLine($"[{batchNumber}] {textToSend}")
                    batchNumber += 1
                Next
                promptBuilder.AppendLine("[/TEXTTOPROCESS]")
                promptBuilder.AppendLine()

                ' Context After (upcoming - plain text only, no markers)
                Dim contextAfterEnd As Integer = Math.Min(processableParagraphs.Count - 1, batchEnd + TranslateContextAfter)
                If contextAfterEnd > batchEnd Then
                    promptBuilder.AppendLine("[CONTEXT AFTER - for reference only]")
                    For j As Integer = batchEnd + 1 To contextAfterEnd
                        Dim ctxAfterText As String = processableParagraphs(j).FullText
                        If ctxAfterText IsNot Nothing Then
                            ctxAfterText = ctxAfterText.Replace(NoteRefMarker, "")
                        End If
                        promptBuilder.AppendLine(ctxAfterText)
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

                Dim Response As String
                If _isFreestyle Then
                    Response = Await SharedMethods.LLM(
                    _context, systemPrompt, promptBuilder.ToString(),
                    "", "", 0, _useSecondAPI, True, AddUserPrompt:=OtherPromptUnfilled)
                Else
                    Response = Await SharedMethods.LLM(
                    _context, systemPrompt, promptBuilder.ToString(),
                    "", "", 0, _useSecondAPI, True)
                End If
                _isFreestyle = False

                If String.IsNullOrWhiteSpace(Response) Then
                    Dim modeNoun As String = If(mode = DocumentProcessMode.Translate, "Translation", "Correction")
                    ShowCustomMessageBox($"LLM returned empty response. {modeNoun} incomplete.")
                    Return False
                End If

                ' Parse and store results
                ParseTranslateResponse(Response, processableParagraphs, batchStart, batchEnd)

                ' ─── Restore non-breaking spaces in translated paragraphs ───
                For j As Integer = batchStart To batchEnd
                    If processableParagraphs(j).TranslatedText IsNot Nothing AndAlso nbspRecords.ContainsKey(j) Then
                        processableParagraphs(j).TranslatedText = RestoreNonBreakingSpaces(
                            processableParagraphs(j).TranslatedText, nbspRecords(j))
                    End If
                Next

                batchIndex = batchEnd + 1

                ' Advance progress bar after each batch completes
                ProgressBarModule.GlobalProgressValue = currentBatch
            End While

            Return True

        Finally
            ' Restore OfflineDocs model if we applied it
            If offlineDocsApplied Then
                _useSecondAPI = False
                If offlineDocsBackup IsNot Nothing Then
                    RestoreDefaults(_context, offlineDocsBackup)
                End If
            End If
        End Try
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
    ''' Determines whether the processed paragraph text is visibly unchanged.
    ''' If unchanged, the original XML runs are left untouched so Word formatting,
    ''' bookmarks, fields, proofing runs, and other run-level metadata remain exactly
    ''' as they were.
    ''' </summary>
    ''' <param name="para">The paragraph being processed.</param>
    ''' <param name="translatedText">The LLM-processed paragraph text.</param>
    ''' <returns><c>True</c> if the visible paragraph text did not change.</returns>
    Private Function IsParagraphTextVisiblyUnchanged(para As TranslateParagraphInfo, translatedText As String) As Boolean
        If para Is Nothing OrElse translatedText Is Nothing Then Return False

        Dim originalVisibleText As String = If(para.FullText, "")
        Dim processedVisibleText As String = translatedText

        If para.TextRuns IsNot Nothing AndAlso para.TextRuns.Any(Function(r) r.HasNoteReferenceBefore) Then
            originalVisibleText = originalVisibleText.Replace(NoteRefMarker, "")
            processedVisibleText = processedVisibleText.Replace(NoteRefMarker, "")
        End If

        ' Only strip formatting markers when this paragraph was actually sent in marker mode
        ' and the original visible text itself does not already contain pipe characters.
        If para.MarkerText IsNot Nothing AndAlso Not originalVisibleText.Contains(RunBoundaryMarker) Then
            processedVisibleText = processedVisibleText.Replace(RunBoundaryMarker, "")
        End If

        Return String.Equals(originalVisibleText, processedVisibleText, StringComparison.Ordinal)
    End Function

    ''' <summary>
    ''' Applies translated/corrected text back to XML nodes, preserving all formatting.
    ''' Partitions runs at protected structural boundaries so that text
    ''' redistribution never moves content across a reference anchor.
    ''' When the LLM preserves ‖ markers, splits directly at their position (they are
    ''' authoritative). Falls back to anchor + word-boundary logic only when ‖ markers
    ''' are missing. The ‖ character is purely a positional placeholder — it indicates
    ''' where a footnote/endnote/field reference or manual line break sits but does not create an additional
    ''' formatting split beyond the existing run partition.
    ''' </summary>
    Private Sub ApplyTranslationsToXml(paragraphs As List(Of TranslateParagraphInfo))
        For Each para In paragraphs
            If para.IsEmpty OrElse String.IsNullOrEmpty(para.TranslatedText) Then Continue For
            If para.TextRuns.Count = 0 Then Continue For

            Dim translatedText As String = para.TranslatedText

            ' If the visible text did not change, do not touch any w:t nodes.
            ' This is the safest way to prevent formatting drift caused by re-splitting
            ' text across Word runs.
            If IsParagraphTextVisiblyUnchanged(para, translatedText) Then Continue For

            ' Simple case: only one run
            If para.TextRuns.Count = 1 Then
                If _useFormattingMarkers Then
                    translatedText = translatedText.Replace(RunBoundaryMarker, "")
                End If
                translatedText = translatedText.Replace(NoteRefMarker, "")
                SetTextNodeWithSpacePreserve(para.TextRuns(0).TextNode, translatedText)
                Continue For
            End If

            ' ─── Check if any protected structural boundaries exist ───
            Dim hasNoteRefBoundaries As Boolean = para.TextRuns.Any(Function(r) r.HasNoteReferenceBefore)

            ' ─── DIAGNOSTIC: Log paragraph state ───
            If hasNoteRefBoundaries Then
                Debug.WriteLine($"")
                Debug.WriteLine($"══════════════════════════════════════════════════")
                Debug.WriteLine($"NOTEREF DIAG - Paragraph {para.Index}, Runs={para.TextRuns.Count}, HasNoteRefBoundaries={hasNoteRefBoundaries}")
                Debug.WriteLine($"  FullText:       [{para.FullText}]")
                Debug.WriteLine($"  MarkerText:     [{para.MarkerText}]")
                Debug.WriteLine($"  TranslatedText: [{translatedText}]")
                For ri As Integer = 0 To para.TextRuns.Count - 1
                    Debug.WriteLine($"  Run[{ri}]: HasNoteRefBefore={para.TextRuns(ri).HasNoteReferenceBefore}, OrigText=[{para.TextRuns(ri).OriginalText}]")
                Next
            End If

            ' Fast path for paragraphs without note/field/manual-break anchors. Paragraphs with anchors
            ' are handled below segment-by-segment, where formatting markers can also be used safely.
            If Not hasNoteRefBoundaries Then
                If _useFormattingMarkers AndAlso para.MarkerText IsNot Nothing Then
                    Dim markerApplied As Boolean = TryApplyMarkerBasedDistribution(para, translatedText)
                    If markerApplied Then Continue For
                    translatedText = translatedText.Replace(RunBoundaryMarker, "")
                End If

                If _useFormattingMarkers Then
                    translatedText = translatedText.Replace(RunBoundaryMarker, "")
                End If

                ' No footnote boundaries: proportional distribution across all runs
                translatedText = translatedText.Replace(NoteRefMarker, "")
                DistributeProportional(para, translatedText)
                Continue For
            End If

            ' ─── From here: paragraph HAS protected structural boundaries ───
            ' Keep formatting markers until each note-delimited segment is applied.

            ' Partition runs into segments at protected structural boundaries
            Dim segments As New List(Of List(Of Integer))()
            Dim currentSegment As New List(Of Integer)()

            For runIdx As Integer = 0 To para.TextRuns.Count - 1
                If para.TextRuns(runIdx).HasNoteReferenceBefore AndAlso currentSegment.Count > 0 Then
                    segments.Add(currentSegment)
                    currentSegment = New List(Of Integer)()
                End If
                currentSegment.Add(runIdx)
            Next
            If currentSegment.Count > 0 Then segments.Add(currentSegment)

            Debug.WriteLine($"  Segments: {segments.Count} (expected ‖ count: {segments.Count - 1})")
            For si As Integer = 0 To segments.Count - 1
                Dim runList As String = String.Join(",", segments(si))
                Dim segOrigText As New StringBuilder()
                For Each ri In segments(si)
                    segOrigText.Append(para.TextRuns(ri).OriginalText)
                Next
                Debug.WriteLine($"    Segment[{si}]: Runs=[{runList}], OrigText=[{segOrigText}]")
            Next

            ' Safety: if partitioning produced only 1 segment, fall back to proportional
            If segments.Count <= 1 Then
                Debug.WriteLine($"  → Only 1 segment, falling back to proportional")
                If _useFormattingMarkers Then
                    translatedText = translatedText.Replace(RunBoundaryMarker, "")
                End If
                translatedText = translatedText.Replace(NoteRefMarker, "")
                DistributeProportional(para, translatedText)
                Continue For
            End If

            ' ─── Split translated text across segments ───
            Dim expectedNoteRefCount As Integer = segments.Count - 1
            Dim actualNoteRefCount As Integer = 0
            For Each ch As Char In translatedText
                If ch = NoteRefMarker(0) Then actualNoteRefCount += 1
            Next

            Debug.WriteLine($"  ‖ count: expected={expectedNoteRefCount}, actual={actualNoteRefCount}")

            Dim segmentTexts As String() = Nothing

            If actualNoteRefCount = expectedNoteRefCount Then
                segmentTexts = translatedText.Split(New String() {NoteRefMarker}, StringSplitOptions.None)

                Debug.WriteLine($"  Split on ‖ produced {segmentTexts.Length} pieces (need {segments.Count})")
                For pi As Integer = 0 To segmentTexts.Length - 1
                    Debug.WriteLine($"    Piece[{pi}]: [{segmentTexts(pi)}]")
                Next

                If segmentTexts.Length = segments.Count Then
                    ' Normalize spaces at segment boundaries
                    For segIdx As Integer = 0 To segmentTexts.Length - 2
                        Dim nextSegFirstRunIdx As Integer = segments(segIdx + 1)(0)
                        Dim nextRunOriginal As String = para.TextRuns(nextSegFirstRunIdx).OriginalText
                        If nextRunOriginal.Length > 0 AndAlso Char.IsWhiteSpace(nextRunOriginal(0)) Then
                            Dim curSeg As String = segmentTexts(segIdx)
                            If curSeg.Length > 0 AndAlso curSeg(curSeg.Length - 1) = " "c Then
                                Dim trimmed As String = curSeg.TrimEnd(" "c)
                                Dim movedSpaces As String = curSeg.Substring(trimmed.Length)
                                segmentTexts(segIdx) = trimmed
                                segmentTexts(segIdx + 1) = movedSpaces & segmentTexts(segIdx + 1)
                                Debug.WriteLine($"    Space normalize: moved trailing space from Piece[{segIdx}] to Piece[{segIdx + 1}]")
                            End If
                        End If
                    Next
                Else
                    Debug.WriteLine($"  → Piece count mismatch! Falling to anchor fallback.")
                    segmentTexts = Nothing
                End If
            End If

            If segmentTexts Is Nothing Then
                Debug.WriteLine($"  → Using ANCHOR FALLBACK")
                If _useFormattingMarkers Then
                    translatedText = translatedText.Replace(RunBoundaryMarker, "")
                End If
                translatedText = translatedText.Replace(NoteRefMarker, "")

                Dim totalOrigLen As Integer = para.FullText.Replace(NoteRefMarker, "").Length
                Dim translatedLen As Integer = translatedText.Length

                If totalOrigLen = 0 Then
                    SetTextNodeWithSpacePreserve(para.TextRuns(0).TextNode, translatedText)
                    For idx As Integer = 1 To para.TextRuns.Count - 1
                        SetTextNodeWithSpacePreserve(para.TextRuns(idx).TextNode, "")
                    Next
                    Continue For
                End If

                Dim segmentOrigTexts As New List(Of String)()
                For Each seg In segments
                    Dim sb As New StringBuilder()
                    For Each ri In seg
                        sb.Append(para.TextRuns(ri).OriginalText)
                    Next
                    segmentOrigTexts.Add(sb.ToString())
                Next

                Dim splitPoints As New List(Of Integer)()
                Dim searchFrom As Integer = 0

                For segIdx As Integer = 0 To segments.Count - 2
                    Dim segOrigText As String = segmentOrigTexts(segIdx)
                    Dim nextSegOrigText As String = segmentOrigTexts(segIdx + 1)
                    Dim cumulativeOrigLen As Integer = 0
                    For si As Integer = 0 To segIdx
                        cumulativeOrigLen += segmentOrigTexts(si).Length
                    Next
                    Dim proportion As Double = cumulativeOrigLen / CDbl(totalOrigLen)
                    Dim proportionalEnd As Integer = CInt(Math.Round(proportion * translatedLen))
                    proportionalEnd = Math.Min(proportionalEnd, translatedLen)

                    Dim splitPos As Integer = FindNoteRefSplitPos(
                        translatedText, segOrigText, searchFrom, proportionalEnd, translatedLen, nextSegOrigText)
                    splitPos = Math.Min(splitPos, translatedLen)

                    If splitPos > searchFrom AndAlso splitPos <= translatedLen Then
                        Dim nextSegFirstRunIdx As Integer = segments(segIdx + 1)(0)
                        Dim nextRunOriginal As String = para.TextRuns(nextSegFirstRunIdx).OriginalText
                        If nextRunOriginal.Length > 0 AndAlso Char.IsWhiteSpace(nextRunOriginal(0)) Then
                            While splitPos > searchFrom AndAlso translatedText(splitPos - 1) = " "c
                                splitPos -= 1
                            End While
                        End If
                    End If

                    Debug.WriteLine($"    Anchor split[{segIdx}]: pos={splitPos}, searchFrom={searchFrom}, proportionalEnd={proportionalEnd}")
                    splitPoints.Add(splitPos)
                    searchFrom = splitPos
                Next
                splitPoints.Add(translatedLen)

                segmentTexts = New String(segments.Count - 1) {}
                Dim st As Integer = 0
                For i As Integer = 0 To segments.Count - 1
                    Dim en As Integer = splitPoints(i)
                    segmentTexts(i) = If(en > st, translatedText.Substring(st, en - st), "")
                    st = en
                Next

                For pi As Integer = 0 To segmentTexts.Length - 1
                    Debug.WriteLine($"    FallbackPiece[{pi}]: [{segmentTexts(pi)}]")
                Next
            End If

            ' ─── Distribute text within each segment independently ───
            For segIdx As Integer = 0 To segments.Count - 1
                Dim segText As String = segmentTexts(segIdx)
                Dim segRunIndices As List(Of Integer) = segments(segIdx)

                If _useFormattingMarkers Then
                    If TryApplyMarkerBasedDistributionToRuns(para, segRunIndices, segText) Then
                        Continue For
                    End If
                    segText = segText.Replace(RunBoundaryMarker, "")
                End If

                If segRunIndices.Count = 1 Then
                    Debug.WriteLine($"  Assign Segment[{segIdx}] → Run[{segRunIndices(0)}]: [{segText}]")
                    SetTextNodeWithSpacePreserve(para.TextRuns(segRunIndices(0)).TextNode, segText)
                Else
                    Dim segOrigLen As Integer = 0
                    For Each ri In segRunIndices
                        segOrigLen += para.TextRuns(ri).OriginalText.Length
                    Next
                    Dim segTransLen As Integer = segText.Length
                    Dim segCurrentPos As Integer = 0
                    Dim segCumOrig As Integer = 0

                    For ri As Integer = 0 To segRunIndices.Count - 1
                        Dim runIdx As Integer = segRunIndices(ri)
                        Dim run = para.TextRuns(runIdx)
                        segCumOrig += run.OriginalText.Length

                        If ri = segRunIndices.Count - 1 Then
                            Dim remaining As String = If(segCurrentPos < segTransLen,
                                                         segText.Substring(segCurrentPos), "")
                            Debug.WriteLine($"  Assign Segment[{segIdx}] → Run[{runIdx}] (last): [{remaining}]")
                            SetTextNodeWithSpacePreserve(run.TextNode, remaining)
                        Else
                            If segOrigLen = 0 Then
                                Debug.WriteLine($"  Assign Segment[{segIdx}] → Run[{runIdx}]: [] (segOrigLen=0)")
                                SetTextNodeWithSpacePreserve(run.TextNode, "")
                                Continue For
                            End If

                            Dim prop As Double = segCumOrig / CDbl(segOrigLen)
                            Dim tgtEnd As Integer = CInt(Math.Round(prop * segTransLen))
                            tgtEnd = Math.Min(tgtEnd, segTransLen)

                            If tgtEnd <= segCurrentPos Then
                                Debug.WriteLine($"  Assign Segment[{segIdx}] → Run[{runIdx}]: [] (tgtEnd<=currentPos)")
                                SetTextNodeWithSpacePreserve(run.TextNode, "")
                                Continue For
                            End If

                            Dim endPos As Integer = tgtEnd
                            Dim foundSpace As Boolean = False

                            If endPos < segTransLen AndAlso endPos > segCurrentPos Then
                                If segText(endPos) = " "c Then
                                    endPos += 1
                                    foundSpace = True
                                Else
                                    Dim searchMax As Integer = Math.Min(endPos + 15, segTransLen - 1)
                                    For sp As Integer = endPos To searchMax
                                        If segText(sp) = " "c Then
                                            endPos = sp + 1
                                            foundSpace = True
                                            Exit For
                                        End If
                                    Next
                                    If Not foundSpace Then
                                        For sp As Integer = endPos - 1 To segCurrentPos Step -1
                                            If segText(sp) = " "c Then
                                                endPos = sp + 1
                                                foundSpace = True
                                                Exit For
                                            End If
                                        Next
                                    End If
                                    If Not foundSpace Then
                                        Dim we As Integer = endPos
                                        While we < segTransLen AndAlso segText(we) <> " "c
                                            we += 1
                                        End While
                                        endPos = we
                                    End If
                                End If
                            End If

                            endPos = Math.Min(endPos, segTransLen)
                            Dim assignedText As String = segText.Substring(segCurrentPos, endPos - segCurrentPos)
                            Debug.WriteLine($"  Assign Segment[{segIdx}] → Run[{runIdx}]: [{assignedText}]")
                            SetTextNodeWithSpacePreserve(run.TextNode, assignedText)
                            segCurrentPos = endPos
                        End If
                    Next
                End If
            Next

            Debug.WriteLine($"══════════════════════════════════════════════════")

        Next
    End Sub

    ''' <summary>
    ''' Proportional distribution fallback when no note-reference boundaries exist.
    ''' </summary>
    Private Sub DistributeProportional(para As TranslateParagraphInfo, translatedText As String)
        Dim totalOrigLen As Integer = para.FullText.Length
        Dim translatedLen As Integer = translatedText.Length

        If totalOrigLen = 0 Then
            SetTextNodeWithSpacePreserve(para.TextRuns(0).TextNode, translatedText)
            For idx As Integer = 1 To para.TextRuns.Count - 1
                SetTextNodeWithSpacePreserve(para.TextRuns(idx).TextNode, "")
            Next
            Return
        End If

        Dim currentPos As Integer = 0
        Dim cumulativeOriginal As Integer = 0

        For runIdx As Integer = 0 To para.TextRuns.Count - 1
            Dim run = para.TextRuns(runIdx)
            cumulativeOriginal += run.OriginalText.Length

            If runIdx = para.TextRuns.Count - 1 Then
                Dim remaining = If(currentPos < translatedLen, translatedText.Substring(currentPos), "")
                SetTextNodeWithSpacePreserve(run.TextNode, remaining)
            Else
                Dim proportion As Double = cumulativeOriginal / CDbl(totalOrigLen)
                Dim targetEndPos = CInt(Math.Round(proportion * translatedLen))
                targetEndPos = Math.Min(targetEndPos, translatedLen)

                If targetEndPos <= currentPos Then
                    SetTextNodeWithSpacePreserve(run.TextNode, "")
                    Continue For
                End If

                Dim endPos As Integer = targetEndPos
                Dim foundSpace As Boolean = False

                If endPos < translatedLen AndAlso endPos > currentPos Then
                    If translatedText(endPos) = " "c Then
                        endPos += 1
                        foundSpace = True
                    Else
                        Dim searchMax As Integer = Math.Min(endPos + 15, translatedLen - 1)
                        For searchPos As Integer = endPos To searchMax
                            If translatedText(searchPos) = " "c Then
                                endPos = searchPos + 1
                                foundSpace = True
                                Exit For
                            End If
                        Next
                        If Not foundSpace Then
                            For searchPos As Integer = endPos - 1 To currentPos Step -1
                                If translatedText(searchPos) = " "c Then
                                    endPos = searchPos + 1
                                    foundSpace = True
                                    Exit For
                                End If
                            Next
                        End If
                        If Not foundSpace Then
                            Dim wordEnd As Integer = endPos
                            While wordEnd < translatedLen AndAlso translatedText(wordEnd) <> " "c
                                wordEnd += 1
                            End While
                            If wordEnd < translatedLen AndAlso translatedText(wordEnd) = " "c Then
                                endPos = wordEnd + 1
                            Else
                                endPos = wordEnd
                            End If
                        End If
                    End If
                End If

                endPos = Math.Min(endPos, translatedLen)
                SetTextNodeWithSpacePreserve(run.TextNode, translatedText.Substring(currentPos, endPos - currentPos))
                currentPos = endPos
            End If
        Next
    End Sub


    ''' <summary>
    ''' Finds the best split position in the translated text for a note-reference boundary
    ''' by anchoring on the last word(s) of the original segment text, or the first word(s)
    ''' of the next segment text. Falls back to proportional position with backward
    ''' word-boundary seeking if both anchoring strategies fail.
    ''' </summary>
    ''' <param name="translatedText">The full translated paragraph text.</param>
    ''' <param name="originalSegText">The original text of the segment before the note reference.</param>
    ''' <param name="currentPos">Start of the current unassigned region in translatedText.</param>
    ''' <param name="targetEndPos">Proportional end position (fallback).</param>
    ''' <param name="translatedLength">Total length of translatedText.</param>
    ''' <param name="nextSegOrigText">The original text of the segment after the note reference (Nothing if unavailable).</param>
    ''' <returns>The character index in translatedText where the split should occur.</returns>
    Private Shared Function FindNoteRefSplitPos(
            translatedText As String,
            originalSegText As String,
            currentPos As Integer,
            targetEndPos As Integer,
            translatedLength As Integer,
            Optional nextSegOrigText As String = Nothing) As Integer

        ' ── Strategy 1: Anchor on the LAST word(s) of the current segment ──
        ' The original segment ended with specific text before the footnote reference.
        ' Find that same trailing text in the translated text to locate where the
        ' footnote should be placed. Try progressively shorter trailing fragments.

        If Not String.IsNullOrWhiteSpace(originalSegText) AndAlso originalSegText.Length >= 2 Then
            Dim origTrimmed As String = originalSegText.TrimEnd()
            If origTrimmed.Length >= 2 Then
                Dim words As String() = origTrimmed.Split(" "c)
                Dim maxWords As Integer = Math.Min(words.Length, 3)

                For tryWords As Integer = maxWords To 1 Step -1
                    Dim fragment As String = String.Join(" ", words, words.Length - tryWords, tryWords)
                    If fragment.Length < 2 Then Continue For

                    Dim searchRegion As String = translatedText.Substring(currentPos)
                    Dim fragIdx As Integer = searchRegion.IndexOf(fragment, StringComparison.OrdinalIgnoreCase)

                    If fragIdx >= 0 Then
                        Dim splitPos As Integer = currentPos + fragIdx + fragment.Length
                        ' Verify it's at a word boundary
                        If splitPos >= translatedLength OrElse
                           translatedText(splitPos) = " "c OrElse
                           Char.IsPunctuation(translatedText(splitPos)) Then
                            Return splitPos
                        End If
                    End If
                Next
            End If
        End If

        ' ── Strategy 2: Anchor on the FIRST word(s) of the NEXT segment ──
        ' If the trailing text was deleted/changed by the LLM, look for the start
        ' of the next segment's text and split just before it.

        If Not String.IsNullOrWhiteSpace(nextSegOrigText) AndAlso nextSegOrigText.Length >= 2 Then
            Dim nextTrimmed As String = nextSegOrigText.TrimStart()
            If nextTrimmed.Length >= 2 Then
                Dim nextWords As String() = nextTrimmed.Split(" "c)
                Dim maxNextWords As Integer = Math.Min(nextWords.Length, 3)

                For tryWords As Integer = maxNextWords To 1 Step -1
                    Dim fragment As String = String.Join(" ", nextWords, 0, tryWords)
                    If fragment.Length < 2 Then Continue For

                    ' Search forward from a reasonable position (don't match too early)
                    ' Use half the proportional position as the earliest plausible start
                    Dim earliestSearch As Integer = Math.Max(currentPos, CInt(currentPos + (targetEndPos - currentPos) * 0.4))
                    If earliestSearch >= translatedLength Then Continue For

                    Dim searchRegion As String = translatedText.Substring(earliestSearch)
                    Dim fragIdx As Integer = searchRegion.IndexOf(fragment, StringComparison.OrdinalIgnoreCase)

                    If fragIdx >= 0 Then
                        Dim matchStart As Integer = earliestSearch + fragIdx
                        ' Split just before this fragment — walk back over any whitespace
                        Dim splitPos As Integer = matchStart
                        While splitPos > currentPos AndAlso translatedText(splitPos - 1) = " "c
                            splitPos -= 1
                        End While
                        ' Verify the character before the split is a word boundary
                        If splitPos <= currentPos Then Continue For
                        If splitPos >= translatedLength OrElse
                           Char.IsWhiteSpace(translatedText(splitPos)) OrElse
                           (splitPos > 0 AndAlso (Char.IsPunctuation(translatedText(splitPos - 1)) OrElse
                                                   Char.IsWhiteSpace(translatedText(splitPos - 1)))) Then
                            Return splitPos
                        End If
                    End If
                Next
            End If
        End If

        ' ── Fallback: backward word-boundary search from proportional position ──
        Dim endPos As Integer = Math.Min(targetEndPos, translatedLength)
        If endPos > currentPos AndAlso endPos < translatedLength Then
            If translatedText(endPos) = " "c Then Return endPos

            For searchPos As Integer = endPos - 1 To currentPos Step -1
                If translatedText(searchPos) = " "c Then
                    Return searchPos + 1
                End If
            Next
        End If

        Return endPos
    End Function


    ''' <summary>
    ''' Attempts to distribute translated text across runs using | markers from the LLM response.
    ''' The markers indicate where formatting boundaries should fall in the translated text.
    ''' </summary>
    ''' <param name="para">The paragraph info with run structure.</param>
    ''' <param name="translatedText">The translated text potentially containing | markers.</param>
    ''' <returns>True if markers were successfully applied; False to fall back to proportional mode.</returns>
    Private Function TryApplyMarkerBasedDistribution(para As TranslateParagraphInfo, translatedText As String) As Boolean
        If para Is Nothing OrElse para.TextRuns Is Nothing Then Return False

        Dim runIndices As New List(Of Integer)()
        For i As Integer = 0 To para.TextRuns.Count - 1
            runIndices.Add(i)
        Next

        Return TryApplyMarkerBasedDistributionToRuns(para, runIndices, translatedText)
    End Function

    ''' <summary>
    ''' Applies marker-delimited text to real formatting groups rather than to every raw Word run.
    ''' Internal run splits with identical formatting therefore cannot move bold/italic boundaries.
    ''' </summary>
    Private Function TryApplyMarkerBasedDistributionToRuns(
        para As TranslateParagraphInfo,
        runIndices As List(Of Integer),
        translatedText As String) As Boolean

        If para Is Nothing OrElse para.TextRuns Is Nothing OrElse runIndices Is Nothing Then Return False

        Dim groups As List(Of List(Of Integer)) = BuildTranslateFormattingGroups(para.TextRuns, runIndices)
        If groups.Count = 0 Then Return False

        Dim expectedMarkerCount As Integer = groups.Count - 1
        Dim actualMarkerCount As Integer = 0
        For Each ch As Char In translatedText
            If ch = RunBoundaryMarker(0) Then actualMarkerCount += 1
        Next

        If actualMarkerCount <> expectedMarkerCount Then
            Debug.WriteLine($"Formatting marker mismatch for paragraph {para.Index}: expected {expectedMarkerCount}, got {actualMarkerCount}. Falling back.")
            Return False
        End If

        Dim segments As String() = translatedText.Split(New String() {RunBoundaryMarker}, StringSplitOptions.None)
        If segments.Length <> groups.Count Then Return False

        For groupIdx As Integer = 0 To groups.Count - 1
            DistributeTextAcrossTranslateRuns(para, groups(groupIdx), segments(groupIdx))
        Next

        Return True
    End Function

    ''' <summary>
    ''' Redistributes one formatting-group text only across runs that all have the same formatting.
    ''' Any approximation remains inside an identical formatting group.
    ''' </summary>
    Private Sub DistributeTextAcrossTranslateRuns(
        para As TranslateParagraphInfo,
        runIndices As List(Of Integer),
        text As String)

        If runIndices Is Nothing OrElse runIndices.Count = 0 Then Return

        If runIndices.Count = 1 Then
            SetTextNodeWithSpacePreserve(para.TextRuns(runIndices(0)).TextNode, text)
            Return
        End If

        Dim totalOrigLen As Integer = 0
        For Each runIdx As Integer In runIndices
            totalOrigLen += para.TextRuns(runIdx).OriginalText.Length
        Next

        If totalOrigLen <= 0 Then
            SetTextNodeWithSpacePreserve(para.TextRuns(runIndices(0)).TextNode, text)
            For i As Integer = 1 To runIndices.Count - 1
                SetTextNodeWithSpacePreserve(para.TextRuns(runIndices(i)).TextNode, "")
            Next
            Return
        End If

        Dim translatedLen As Integer = text.Length
        Dim currentPos As Integer = 0
        Dim cumulativeOriginal As Integer = 0

        For groupRunPos As Integer = 0 To runIndices.Count - 1
            Dim runIdx As Integer = runIndices(groupRunPos)
            Dim run As TranslateTextRunInfo = para.TextRuns(runIdx)
            cumulativeOriginal += run.OriginalText.Length

            If groupRunPos = runIndices.Count - 1 Then
                Dim remaining As String = If(currentPos < translatedLen, text.Substring(currentPos), "")
                SetTextNodeWithSpacePreserve(run.TextNode, remaining)
                Continue For
            End If

            Dim targetEndPos As Integer = CInt(Math.Round((cumulativeOriginal / CDbl(totalOrigLen)) * translatedLen))
            targetEndPos = Math.Min(targetEndPos, translatedLen)

            If targetEndPos <= currentPos Then
                SetTextNodeWithSpacePreserve(run.TextNode, "")
                Continue For
            End If

            Dim endPos As Integer = targetEndPos
            Dim foundSpace As Boolean = False

            If endPos < translatedLen AndAlso endPos > currentPos Then
                If text(endPos) = " "c Then
                    endPos += 1
                    foundSpace = True
                Else
                    Dim searchMax As Integer = Math.Min(endPos + 15, translatedLen - 1)
                    For searchPos As Integer = endPos To searchMax
                        If text(searchPos) = " "c Then
                            endPos = searchPos + 1
                            foundSpace = True
                            Exit For
                        End If
                    Next

                    If Not foundSpace Then
                        For searchPos As Integer = endPos - 1 To currentPos Step -1
                            If text(searchPos) = " "c Then
                                endPos = searchPos + 1
                                foundSpace = True
                                Exit For
                            End If
                        Next
                    End If

                    If Not foundSpace Then
                        Dim wordEnd As Integer = endPos
                        While wordEnd < translatedLen AndAlso text(wordEnd) <> " "c
                            wordEnd += 1
                        End While
                        endPos = wordEnd
                    End If
                End If
            End If

            endPos = Math.Min(endPos, translatedLen)
            SetTextNodeWithSpacePreserve(run.TextNode, text.Substring(currentPos, endPos - currentPos))
            currentPos = endPos
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

                Dim isFootnotes As Boolean = fileName.Equals("footnotes.xml", StringComparison.OrdinalIgnoreCase)
                Dim componentType As String = If(isFootnotes, "Footnotes", "Endnotes")
                Dim contentScope As String = If(isFootnotes, "footnotes", "endnotes")
                Dim success As Boolean = Await ProcessParagraphBatches(paragraphs, targetLanguage, mode, $"{mainFileName} ({componentType})", contentScope)
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


    ''' <summary>
    ''' Scans text for non-breaking space characters and records their surrounding
    ''' word context so they can be restored after LLM processing.
    ''' </summary>
    ''' <param name="text">The original text to scan.</param>
    ''' <returns>A list of non-breaking space occurrences with context, or Nothing if none found.</returns>
    Private Shared Function RecordNonBreakingSpaces(text As String) As List(Of NonBreakingSpaceInfo)
        If String.IsNullOrEmpty(text) Then Return Nothing

        Dim hasNbsp As Boolean = False
        For Each ch As Char In text
            If NonBreakingSpaceChars.Contains(ch) Then
                hasNbsp = True
                Exit For
            End If
        Next
        If Not hasNbsp Then Return Nothing

        Dim results As New List(Of NonBreakingSpaceInfo)()

        ' Tokenize: split on regular spaces and non-breaking spaces, keeping delimiters
        ' We walk the string character by character to capture context around each nbsp
        Dim i As Integer = 0
        While i < text.Length
            If NonBreakingSpaceChars.Contains(text(i)) Then
                Dim info As New NonBreakingSpaceInfo() With {
                    .SpaceChar = text(i)
                }

                ' Find word before: scan backward from i-1 to find the token
                If i > 0 Then
                    Dim wordEnd As Integer = i - 1
                    ' Skip any regular whitespace between the word and the nbsp
                    While wordEnd >= 0 AndAlso text(wordEnd) = " "c
                        wordEnd -= 1
                    End While
                    If wordEnd >= 0 Then
                        Dim wordStart As Integer = wordEnd
                        While wordStart > 0 AndAlso text(wordStart - 1) <> " "c AndAlso Not NonBreakingSpaceChars.Contains(text(wordStart - 1))
                            wordStart -= 1
                        End While
                        info.WordBefore = text.Substring(wordStart, wordEnd - wordStart + 1)
                    End If
                End If

                ' Find word after: scan forward from i+1 to find the token
                If i < text.Length - 1 Then
                    Dim wordStart As Integer = i + 1
                    ' Skip any regular whitespace between the nbsp and the next word
                    While wordStart < text.Length AndAlso text(wordStart) = " "c
                        wordStart += 1
                    End While
                    If wordStart < text.Length Then
                        Dim wordEnd As Integer = wordStart
                        While wordEnd < text.Length - 1 AndAlso text(wordEnd + 1) <> " "c AndAlso Not NonBreakingSpaceChars.Contains(text(wordEnd + 1))
                            wordEnd += 1
                        End While
                        info.WordAfter = text.Substring(wordStart, wordEnd - wordStart + 1)
                    End If
                End If

                results.Add(info)
            End If
            i += 1
        End While

        Return If(results.Count > 0, results, Nothing)
    End Function

    ''' <summary>
    ''' Restores non-breaking spaces in text that has been processed by the LLM,
    ''' using the recorded word context to find where they should be placed.
    ''' Only restores a non-breaking space when both surrounding words match.
    ''' </summary>
    ''' <param name="text">The LLM-processed text (with regular spaces).</param>
    ''' <param name="nbspInfos">The recorded non-breaking space positions from the original text.</param>
    ''' <returns>The text with non-breaking spaces restored where context matches.</returns>
    Private Shared Function RestoreNonBreakingSpaces(text As String, nbspInfos As List(Of NonBreakingSpaceInfo)) As String
        If String.IsNullOrEmpty(text) OrElse nbspInfos Is Nothing OrElse nbspInfos.Count = 0 Then Return text

        ' For each recorded non-breaking space, find the matching word pair in the text
        ' and replace the regular space between them with the original non-breaking space character.
        ' Process in reverse order of match position to keep indices stable.
        Dim replacements As New SortedList(Of Integer, Char)()

        For Each info In nbspInfos
            ' Both words must be present for a safe match
            If String.IsNullOrEmpty(info.WordBefore) OrElse String.IsNullOrEmpty(info.WordAfter) Then Continue For

            ' Build a pattern: wordBefore + one or more regular spaces + wordAfter
            ' Use Regex.Escape to handle special characters in the words
            Dim pattern As String = Regex.Escape(info.WordBefore) & "( +)" & Regex.Escape(info.WordAfter)
            Dim m As Match = Regex.Match(text, pattern, RegexOptions.None)

            ' Find the first match that hasn't already been claimed by another replacement
            While m.Success
                Dim spaceGroupStart As Integer = m.Groups(1).Index
                Dim spaceGroupLen As Integer = m.Groups(1).Length

                ' Check if this space position is already claimed
                Dim alreadyClaimed As Boolean = False
                For spIdx As Integer = spaceGroupStart To spaceGroupStart + spaceGroupLen - 1
                    If replacements.ContainsKey(spIdx) Then
                        alreadyClaimed = True
                        Exit For
                    End If
                Next

                If Not alreadyClaimed Then
                    ' Replace the first space character in the group with the non-breaking space
                    replacements(spaceGroupStart) = info.SpaceChar
                    Exit While
                End If

                m = m.NextMatch()
            End While
        Next

        If replacements.Count = 0 Then Return text

        ' Apply replacements (from end to start to preserve indices)
        Dim sb As New StringBuilder(text)
        For i As Integer = replacements.Count - 1 To 0 Step -1
            Dim pos As Integer = replacements.Keys(i)
            sb(pos) = replacements.Values(i)
        Next

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' Expands paragraphs whose translated text contains line breaks into multiple
    ''' sibling &lt;w:p&gt; elements, cloning the original paragraph's formatting.
    ''' Must be called AFTER ApplyTranslationsToXml.
    ''' </summary>
    Private Sub ExpandMultiLineParagraphs(paragraphs As List(Of TranslateParagraphInfo),
                                          nsMgr As System.Xml.XmlNamespaceManager)
        For Each para In paragraphs
            If para.IsEmpty OrElse String.IsNullOrEmpty(para.TranslatedText) Then Continue For
            If para.TextRuns.Count = 0 Then Continue For

            Dim lines As String() = para.TranslatedText.Split({vbCrLf, vbLf, vbCr}, StringSplitOptions.None)
            If lines.Length <= 1 Then Continue For

            ' Find the parent <w:p> node from the first text run
            Dim paraNode As System.Xml.XmlNode = para.TextRuns(0).TextNode
            While paraNode IsNot Nothing AndAlso paraNode.LocalName <> "p"
                paraNode = paraNode.ParentNode
            End While
            If paraNode Is Nothing OrElse paraNode.ParentNode Is Nothing Then Continue For

            Dim parentNode As System.Xml.XmlNode = paraNode.ParentNode

            ' Set the first line on the existing paragraph
            SetAllTextInParagraph(paraNode, nsMgr, lines(0).TrimEnd())

            ' Clone and insert additional paragraphs for remaining lines
            Dim insertAfter As System.Xml.XmlNode = paraNode
            For lineIdx As Integer = 1 To lines.Length - 1
                Dim lineText As String = lines(lineIdx).TrimEnd()
                If String.IsNullOrEmpty(lineText) Then Continue For

                Dim newPara As System.Xml.XmlNode = paraNode.CloneNode(deep:=True)
                SetAllTextInParagraph(newPara, nsMgr, lineText)
                parentNode.InsertAfter(newPara, insertAfter)
                insertAfter = newPara
            Next
        Next
    End Sub

    ''' <summary>
    ''' Sets all w:t nodes in a paragraph to a single text value (first run gets
    ''' the text, remaining runs are cleared).
    ''' </summary>
    Private Sub SetAllTextInParagraph(paraNode As System.Xml.XmlNode,
                                      nsMgr As System.Xml.XmlNamespaceManager,
                                      text As String)
        Dim tNodes = paraNode.SelectNodes(".//w:r/w:t", nsMgr)
        If tNodes Is Nothing OrElse tNodes.Count = 0 Then Return

        For i As Integer = 0 To tNodes.Count - 1
            If i = 0 Then
                SetTextNodeWithSpacePreserve(tNodes(i), text)
            Else
                SetTextNodeWithSpacePreserve(tNodes(i), "")
            End If
        Next
    End Sub

End Class
