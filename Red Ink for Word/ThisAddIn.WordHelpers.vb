' Copyright (c) LawDigital Ltd., Switzerland.
' All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.WordHelpers.vb
' Purpose: Word-centric helper utilities for the add-in, including document and
'          selection comparison workflows, change extraction/analysis for LLM
'          summarization, markup/time-span analytics, PDF operations (flattening,
'          splitting, exhibit stamping), comment management, and various editing
'          helpers (content control removal, regex search/replace, file
'          import/export utilities).
'
' Architecture:
' - Comparison UI (HTML viewer):
'     - `CompareActiveDocWithOtherOpenDoc`:
'         - If a text selection exists, routes to `CompareSelectedTextRanges`.
'         - Otherwise compares the active document against another open document
'           chosen by the user.
'         - Generates a comparison document via `Word.Application.CompareDocuments`,
'           exports it to Filtered HTML (temp folder), and displays it with
'           `ShowHTMLCustomMessageBox`.
'         - Provides post-actions (buttons) such as: summarize changes, export to PDF,
'           copy result to clipboard (with formatting), open result in Word,
'           switch comparison direction (⇄), and run further "compare selected"
'           operations.
'     - `CompareSelectedTextRanges`:
'         - Captures two user-selected ranges (non-modal prompts so Word stays usable),
'           compares via `CompareDocuments` on temporary documents, exports to HTML,
'           and displays with the same action buttons.
'
' - Change Extraction + LLM Summarization:
'     - `ExtractChangesWithMarkupTags`:
'         - Extracts revisions into XML-like tags (`<ins>`, `<del>`) and captures
'           comments, footnotes, and endnotes for downstream processing.
'     - `SummarizeComparisonChangesAsync`:
'         - Submits extracted markup to `SharedMethods.LLM` with `SP_Markup` and renders
'           the Markdown result to HTML (Markdig) for display.
'     - `SummarizeDocumentChanges` / `ExtractRevisionsAndCommentsWithMarkup`:
'         - Summarizes revisions/comments from a selection or whole document with optional
'           date filtering and comment cutoffs relative to earliest revision.
'         - Ignores pure formatting revisions for output but uses them for date calculation.
'
' - PDF Operations:
'     - `FlattenPdfToImages` / `FlattenPdfToImageOnly`:
'         - Rasterizes PDF pages to image-only PDFs, removing text layers.
'         - Prefers Windows.Data.Pdf (Win10+ built-in renderer) via `FlattenPdfViaWindowsApi`;
'           falls back to PdfiumViewer when unavailable.
'         - Requires HAS_WINRT compilation constant (auto-detected from Windows SDK).
'         - Supports per-page JPEG encoding at configurable DPI, metadata preservation,
'           recursive directory processing, and progress bar with cancellation.
'     - `RenderPageViaWindowsPdf` / `RenderAllPagesViaWindowsPdf`:
'         - Single-page and full-document Windows.Data.Pdf rendering helpers
'           used by both flattening and exhibit stamping (Stamper.vb).
'         - Execute on STA threads with timeout protection.
'     - `SplitPdfByExhibits`:
'         - Extracts text page-by-page (PdfPig, with optional OCR fallback),
'           sends to LLM for exhibit boundary detection, and splits via PdfSharp.
'         - Supports configurable output language for descriptions and user
'           confirmation of proposed splits.
'
' - Export/IO Helpers:
'     - `ReadHtmlWithEncodingDetection` detects encoding of Word-generated Filtered HTML
'       (BOM/meta charset; fallback to Windows-1252), used before displaying in the HTML viewer.
'     - `ExportComparisonToPdfFromHtml` / `ExportComparisonToPdf` reopens exported HTML
'       in Word and exports to PDF, prompting for output path and ensuring unique filenames.
'     - `ExportFileContentToText` converts files (PDF, DOCX, TXT, etc.) to plain text,
'       with optional OCR for PDFs and batch processing support.
'
' - Comment Management:
'     - `NewRemoveRIPrefixFromComments` / `RemoveRIPrefixFromComments`:
'         - Removes leading "RI: " prefix from Word comments (including threaded replies).
'         - Optionally removes comment date/time via Open XML post-processing
'           (`RemoveCommentDateTimesViaOpenXml`) by saving to temp DOCX, stripping
'           w:date attributes from comments.xml and commentsExtended/Extensible parts.
'
' - Editing Utilities:
'     - Content control removal helpers (`RemoveContentControlsRespectSelection`,
'       `RemoveAllContentControlsKeepContents`, `RemoveContentControlsInRangeKeepContents`)
'       remove controls while preserving text/formatting and safely handling protected docs.
'     - `RegexSearchReplace` provides multi-pattern regex operations with persistent settings.
'     - `AcceptFormatting` accepts formatting-only revisions with escape-to-cancel UX.
'     - `CalculateUserMarkupTimeSpan` calculates markup/comment time spans by author with optional date filter.
'     - `CompareSelectionHalves` splits selected paragraphs into two halves and compares them.
'
' Threading / UX:
' - Uses non-modal top-most dialogs for selection capture, enabling the user to interact with Word
'   while prompts are visible.
' - Long-running work uses progress/splash patterns, and asynchronous UI rendering uses STA threads
'   where required by WinForms/embedded browser controls.
' - Windows.Data.Pdf calls execute on dedicated STA threads with 2–5 minute timeout protection.
'
' Build Configuration (Windows SDK / WinRT):
' - The HAS_WINRT conditional compilation constant is auto-defined in the .vbproj
'   when Windows.winmd is found at the expected SDK path.
' - Currently requires Windows SDK 10.0.26100.0 (see .vbproj for update instructions).
' - When HAS_WINRT is NOT defined, all Windows.Data.Pdf code is compiled out and
'   PdfiumViewer is used as fallback (reduced rendering fidelity for complex PDFs).
'
' Dependencies:
' - Microsoft Office Interop Word (`Microsoft.Office.Interop.Word`)
' - Markdig (Markdown → HTML rendering)
' - PdfSharp (PDF creation/manipulation)
' - PdfPig (PDF text extraction / page counting)
' - PdfiumViewer (PDF rendering fallback when Windows.Data.Pdf unavailable)
' - Windows.Data.Pdf (Win10+ built-in PDF renderer, via WinRT projection)
' - DocumentFormat.OpenXml (Open XML post-processing for comment date removal)
' - SharedLibrary (`SharedLibrary.SharedMethods`) for dialogs, clipboard helpers, LLM access,
'   progress bar, splash screen, drag & drop forms, and OCR utilities
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Markdig
Imports Microsoft.Office.Interop.Word
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports Slib = SharedLibrary.SharedLibrary.SharedMethods



Partial Public Class ThisAddIn

    ''' <summary>
    ''' Represents a segment of a PDF to be extracted as a separate exhibit file.
    ''' </summary>
    Friend Class SplitSegment
        ''' <summary>1-based start page number.</summary>
        Public Property StartPage As Integer
        ''' <summary>1-based end page number (inclusive).</summary>
        Public Property EndPage As Integer
        ''' <summary>Descriptive name for the exhibit.</summary>
        Public Property Name As String
    End Class

    ''' <summary>
    ''' CSS styling applied to all HTML summary windows for consistent formatting of LLM-generated content.
    ''' </summary>
    Private Const SummaryHtmlStyle As String =
        "<style>" &
        "body { font-family: 'Segoe UI', Tahoma, Arial, sans-serif; font-size: 10pt; line-height: 1.5; padding: 20px; margin: 0; }" &
        "ul, ol { margin-left: 20px; }" &
        "li { margin-bottom: 6px; }" &
        "h1, h2, h3 { color: #333; }" &
        "strong { color: #003366; }" &
        "code { background: #f6f8fa; padding: 2px 4px; border-radius: 3px; }" &
        "pre { background: #f6f8fa; padding: 10px; border-radius: 4px; overflow-x: auto; }" &
        "</style>"

    ''' <summary>
    ''' Compares the active Word document with another open document selected by the user.
    ''' Generates a comparison document, exports it to filtered HTML, and displays the result
    ''' with buttons for summarization, PDF export, Word export, and selection-based comparison.
    ''' If text is selected, redirects to Compare Selected flow regardless of other open documents.
    ''' </summary>
    Public Shared Sub CompareActiveDocWithOtherOpenDoc()
        ' Acquire the running Word instance
        Dim wordAppObj As Object = Nothing
        Try
            wordAppObj = System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application")
        Catch
            ShowCustomMessageBox("Microsoft Word is not running or cannot be accessed.", AN)
            Exit Sub
        End Try

        Dim wordApp As Microsoft.Office.Interop.Word.Application = TryCast(wordAppObj, Microsoft.Office.Interop.Word.Application)
        If wordApp Is Nothing Then
            ShowCustomMessageBox("Unable to access the Word application.", AN)
            Exit Sub
        End If

        ' Ensure there is an active document
        If wordApp.Documents Is Nothing OrElse wordApp.Documents.Count = 0 Then
            ShowCustomMessageBox("No document is open in Word.", AN)
            Exit Sub
        End If

        Dim activeDoc As Microsoft.Office.Interop.Word.Document = Nothing
        Try
            activeDoc = wordApp.ActiveDocument
        Catch
        End Try
        If activeDoc Is Nothing Then
            ShowCustomMessageBox("No active document detected in Word.", AN)
            Exit Sub
        End If

        ' CHECK FOR TEXT SELECTION FIRST - if text is selected, always go to Compare Selected
        Dim hasSelection As Boolean = False
        Try
            Dim sel As Microsoft.Office.Interop.Word.Selection = wordApp.Selection
            If sel IsNot Nothing AndAlso sel.Range IsNot Nothing AndAlso sel.Start <> sel.End Then
                hasSelection = True
            End If
        Catch
        End Try

        If hasSelection Then
            ' Text is selected - redirect to Compare Selected regardless of other documents
            CompareSelectedTextRanges(wordApp)
            Exit Sub
        End If

        ' No text selected - proceed with document comparison
        ' Build the list of other open documents
        Dim otherDocs As New List(Of Microsoft.Office.Interop.Word.Document)()
        For Each d As Microsoft.Office.Interop.Word.Document In wordApp.Documents
            If Not Object.ReferenceEquals(d, activeDoc) Then
                otherDocs.Add(d)
            End If
        Next

        If otherDocs.Count = 0 Then
            ShowCustomMessageBox("No other open document found to compare against. To compare text selections, first select text in the document.", AN)
            Exit Sub
        End If

        ' Pick the second document (auto if only one; otherwise ask via SLib.SelectValue)
        Dim docToCompare As Microsoft.Office.Interop.Word.Document = Nothing
        If otherDocs.Count = 1 Then
            docToCompare = otherDocs(0)
        Else
            Dim items As New List(Of Slib.SelectionItem)()
            Dim indexToDoc As New Dictionary(Of Integer, Microsoft.Office.Interop.Word.Document)()
            Dim idx As Integer = 1
            For Each d In otherDocs
                Dim disp As String
                Try
                    disp = If(String.IsNullOrEmpty(d.Name), "(unnamed document)", d.Name)
                Catch
                    disp = "(document)"
                End Try
                items.Add(New Slib.SelectionItem(disp, idx))
                indexToDoc(idx) = d
                idx += 1
            Next

            Dim chosenIdx As Integer = Slib.SelectValue(items, 1, "Select the document to compare with:", $"{AN} Compare")
            If chosenIdx <= 0 OrElse Not indexToDoc.ContainsKey(chosenIdx) Then Exit Sub
            docToCompare = indexToDoc(chosenIdx)
        End If

        ' Delegate to the shared comparison + display helper
        CompareDocumentsAndShowResult(wordApp, activeDoc, docToCompare)
    End Sub

    ''' <summary>
    ''' Performs document comparison between originalDoc and revisedDoc, exports the result to
    ''' filtered HTML, and displays it in the HTML viewer with action buttons including
    ''' "Switch Direction" to re-compare with swapped documents.
    ''' </summary>
    ''' <param name="wordApp">The Word application instance.</param>
    ''' <param name="originalDoc">The document treated as the original (baseline).</param>
    ''' <param name="revisedDoc">The document treated as the revised version.</param>
    Private Shared Sub CompareDocumentsAndShowResult(
        wordApp As Microsoft.Office.Interop.Word.Application,
        originalDoc As Microsoft.Office.Interop.Word.Document,
        revisedDoc As Microsoft.Office.Interop.Word.Document)

        ' Store document names for PDF naming
        Dim originalDocName As String = Path.GetFileNameWithoutExtension(If(originalDoc.Name, "Original"))
        Dim revisedDocName As String = Path.GetFileNameWithoutExtension(If(revisedDoc.Name, "Revised"))
        Dim originalDocPath As String = Nothing
        Try
            originalDocPath = originalDoc.Path
        Catch
        End Try

        ' Run the comparison and display - NOT on a separate thread since ShowHTMLCustomMessageBox handles threading
        Dim compareDoc As Microsoft.Office.Interop.Word.Document = Nothing
        Dim tempHtmlPath As String = Nothing
        Dim tempFolder As String = Nothing

        ' UI suppression to reduce flicker
        Dim prevScreenUpdating As Boolean = True
        Dim prevAlerts As Microsoft.Office.Interop.Word.WdAlertLevel = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsAll
        Dim prevWindow As Microsoft.Office.Interop.Word.Window = Nothing

        ' Store extracted changes for LLM summarization
        Dim extractedChangesText As String = Nothing

        Try
            prevScreenUpdating = wordApp.ScreenUpdating
            prevAlerts = wordApp.DisplayAlerts
            wordApp.ScreenUpdating = False
            wordApp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone
            prevWindow = wordApp.ActiveWindow

            ' Create comparison document
            compareDoc = wordApp.CompareDocuments(
                OriginalDocument:=originalDoc,
                RevisedDocument:=revisedDoc,
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
                CompareMoves:=True,
                RevisedAuthor:=GetMarkupAuthorOrCurrent(wordApp),
                IgnoreAllComparisonWarnings:=False
            )
            If compareDoc Is Nothing Then
                wordApp.DisplayAlerts = prevAlerts
                wordApp.ScreenUpdating = prevScreenUpdating
                ShowCustomMessageBox("Word did not produce a comparison document.", AN)
                Exit Sub
            End If

            ' Keep its window hidden (best effort)
            Try
                If compareDoc.Windows IsNot Nothing AndAlso compareDoc.Windows.Count > 0 Then
                    compareDoc.Windows(1).Visible = False
                End If
            Catch
            End Try

            ' Extract changes with markup tags for LLM summarization
            extractedChangesText = ExtractChangesWithMarkupTags(compareDoc)

            ' Export to filtered HTML
            tempFolder = Path.Combine(Path.GetTempPath(), $"{AN2}_compare_" & Guid.NewGuid().ToString("N"))
            Directory.CreateDirectory(tempFolder)
            tempHtmlPath = Path.Combine(tempFolder, "comparison.htm")

            compareDoc.SaveAs2(FileName:=tempHtmlPath, FileFormat:=WdSaveFormat.wdFormatFilteredHTML)

            ' Close the comparison document NOW - we have the HTML
            Try
                compareDoc.Close(WdSaveOptions.wdDoNotSaveChanges)
            Catch
            End Try
            compareDoc = Nothing

            ' Restore focus ASAP to reduce flicker
            Try
                prevWindow?.Activate()
            Catch
            End Try

            ' Restore UI settings now
            wordApp.DisplayAlerts = prevAlerts
            wordApp.ScreenUpdating = prevScreenUpdating

            ' Capture references for closures
            Dim capturedWordApp As Microsoft.Office.Interop.Word.Application = wordApp
            Dim capturedOriginalDoc As Microsoft.Office.Interop.Word.Document = originalDoc
            Dim capturedRevisedDoc As Microsoft.Office.Interop.Word.Document = revisedDoc
            Dim capturedOriginalDocName As String = originalDocName
            Dim capturedRevisedDocName As String = revisedDocName
            Dim capturedOriginalDocPath As String = originalDocPath
            Dim capturedTempFolder As String = tempFolder
            Dim capturedTempHtmlPath As String = tempHtmlPath

            ' Read HTML with proper encoding detection
            Dim htmlContent As String = ReadHtmlWithEncodingDetection(capturedTempHtmlPath)

            ' Inject base href and Mark of the Web for security
            Dim baseHref As String = $"<base href=""file:///{capturedTempFolder.Replace("\", "/")}/"">"
            Dim motw As String = "<!-- saved from url=(0016)http://localhost -->"
            htmlContent = htmlContent.Replace("<head>", "<head>" & vbCrLf & motw & vbCrLf & baseHref)

            ' Build additional buttons array
            Dim additionalButtons As New List(Of System.Tuple(Of String, System.Action, Boolean))()

            ' Switch Direction button (⇄ icon via Unicode)
            additionalButtons.Add(New System.Tuple(Of String, System.Action, Boolean)(
                "⇄",
                Sub() CompareDocumentsAndShowResult(capturedWordApp, capturedRevisedDoc, capturedOriginalDoc),
                True))

            ' Summarize Changes button
            If Not String.IsNullOrWhiteSpace(extractedChangesText) Then
                additionalButtons.Add(New System.Tuple(Of String, System.Action, Boolean)(
                    "Summarize Changes",
                    Sub() SummarizeComparisonChangesAsync(extractedChangesText),
                    False))
            End If

            ' Send to PDF button
            Dim pdfDefaultPath As String = If(String.IsNullOrEmpty(capturedOriginalDocPath) OrElse Not Directory.Exists(capturedOriginalDocPath),
                                              Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                                              capturedOriginalDocPath)
            Dim pdfDefaultName As String = $"Compare_{SanitizeFileName(capturedOriginalDocName)}_{SanitizeFileName(capturedRevisedDocName)}"
            additionalButtons.Add(New System.Tuple(Of String, System.Action, Boolean)(
                "Send to PDF",
                Sub()
                    If Not File.Exists(capturedTempHtmlPath) Then
                        ShowCustomMessageBox("The comparison file is no longer available.", AN)
                        Return
                    End If
                    ExportComparisonToPdfFromHtml(capturedTempHtmlPath, capturedWordApp, pdfDefaultName, pdfDefaultPath)
                End Sub,
                False))

            ' Copy to Clipboard button - copy the comparison WITH formatting/markup
            additionalButtons.Add(New System.Tuple(Of String, System.Action, Boolean)(
                "Copy to Clipboard",
                Sub()
                    Try
                        If Not File.Exists(capturedTempHtmlPath) Then
                            ShowCustomMessageBox("The comparison file is no longer available.", AN)
                            Return
                        End If
                        ' Open the HTML file temporarily in Word and copy the content with formatting
                        Dim tempDoc As Microsoft.Office.Interop.Word.Document = Nothing
                        Dim prevScreenUpdating_2nd As Boolean = capturedWordApp.ScreenUpdating
                        Dim prevAlerts_2nd As Microsoft.Office.Interop.Word.WdAlertLevel = capturedWordApp.DisplayAlerts
                        Try
                            capturedWordApp.ScreenUpdating = False
                            capturedWordApp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone

                            tempDoc = capturedWordApp.Documents.Open(
                                FileName:=capturedTempHtmlPath,
                                ReadOnly:=True,
                                Visible:=False)

                            ' Select all content and copy to clipboard (preserves formatting and markup)
                            tempDoc.Content.Copy()

                            tempDoc.Close(WdSaveOptions.wdDoNotSaveChanges)
                            tempDoc = Nothing

                            capturedWordApp.DisplayAlerts = prevAlerts_2nd
                            capturedWordApp.ScreenUpdating = prevScreenUpdating_2nd

                            ShowCustomMessageBox("Comparison copied to clipboard with formatting. You can now paste it into any Word document.", AN)
                        Finally
                            If tempDoc IsNot Nothing Then
                                Try
                                    tempDoc.Close(WdSaveOptions.wdDoNotSaveChanges)
                                Catch
                                End Try
                            End If
                            capturedWordApp.DisplayAlerts = prevAlerts_2nd
                            capturedWordApp.ScreenUpdating = prevScreenUpdating_2nd
                        End Try
                    Catch ex As Exception
                        ShowCustomMessageBox($"Failed to copy to clipboard: {ex.Message}", AN)
                    End Try
                End Sub,
                False))

            ' Send to Document button
            additionalButtons.Add(New System.Tuple(Of String, System.Action, Boolean)(
                "Send to Document",
                Sub()
                    Try
                        If Not File.Exists(capturedTempHtmlPath) Then
                            ShowCustomMessageBox("The comparison file is no longer available.", AN)
                            Return
                        End If
                        ' Open the HTML file in Word as a new document
                        Dim newDoc As Microsoft.Office.Interop.Word.Document = capturedWordApp.Documents.Open(
                            FileName:=capturedTempHtmlPath,
                            ReadOnly:=False,
                            Visible:=True)
                        newDoc.Activate()
                    Catch ex As Exception
                        ShowCustomMessageBox($"Failed to open in Word: {ex.Message}", AN)
                    End Try
                End Sub,
                False))

            ' Compare Selected button
            additionalButtons.Add(New System.Tuple(Of String, System.Action, Boolean)(
                "Compare Selected",
                Sub() CompareSelectedTextRanges(capturedWordApp),
                False))

            ' Define cleanup action
            Dim cleanupAction As System.Action =
                    Sub()
                        Try
                            If Directory.Exists(capturedTempFolder) Then
                                Directory.Delete(capturedTempFolder, recursive:=True)
                            End If
                        Catch
                            ' Ignore cleanup errors
                        End Try
                    End Sub

            ' Show result with all buttons - cleanup happens when dialog closes
            ' Title includes direction hint: Original → Revised
            Dim directionHint As String = $" ({capturedOriginalDocName} → {capturedRevisedDocName})"
            ShowHTMLCustomMessageBox(htmlContent, $"{AN} Word Active Compare{directionHint}", additionalButtons:=additionalButtons.ToArray(), onClose:=cleanupAction)

        Catch ex As System.Exception
            ' Safety close on error
            If compareDoc IsNot Nothing Then
                Try
                    compareDoc.Close(WdSaveOptions.wdDoNotSaveChanges)
                Catch
                End Try
            End If

            wordApp.DisplayAlerts = prevAlerts
            wordApp.ScreenUpdating = prevScreenUpdating

            ShowCustomMessageBox($"Comparison failed: {ex.Message}", AN)
        End Try
    End Sub

    ''' <summary>
    ''' Reads an HTML file with proper encoding detection.
    ''' Word's filtered HTML often uses Windows-1252 encoding.
    ''' </summary>
    Private Shared Function ReadHtmlWithEncodingDetection(filePath As String) As String
        If Not File.Exists(filePath) Then Return String.Empty

        ' First, read as bytes to detect encoding
        Dim bytes As Byte() = File.ReadAllBytes(filePath)

        ' Check for BOM
        Dim encoding As System.Text.Encoding = Nothing

        If bytes.Length >= 3 AndAlso bytes(0) = &HEF AndAlso bytes(1) = &HBB AndAlso bytes(2) = &HBF Then
            encoding = System.Text.Encoding.UTF8
        ElseIf bytes.Length >= 2 AndAlso bytes(0) = &HFF AndAlso bytes(1) = &HFE Then
            encoding = System.Text.Encoding.Unicode
        ElseIf bytes.Length >= 2 AndAlso bytes(0) = &HFE AndAlso bytes(1) = &HFF Then
            encoding = System.Text.Encoding.BigEndianUnicode
        End If

        ' If no BOM, try to detect from meta charset tag
        If encoding Is Nothing Then
            ' Read first 1024 bytes as ASCII to find charset
            Dim headerText As String = System.Text.Encoding.ASCII.GetString(bytes, 0, Math.Min(1024, bytes.Length))

            ' Look for charset in meta tag
            Dim charsetMatch As Match = Regex.Match(headerText, "charset\s*=\s*[""']?([^""'\s>]+)", RegexOptions.IgnoreCase)
            If charsetMatch.Success Then
                Dim charsetName As String = charsetMatch.Groups(1).Value.Trim()
                Try
                    encoding = System.Text.Encoding.GetEncoding(charsetName)
                Catch
                    ' Invalid charset name, fall back
                End Try
            End If
        End If

        ' Default to Windows-1252 for Word HTML (common for Western European)
        If encoding Is Nothing Then
            Try
                encoding = System.Text.Encoding.GetEncoding(1252) ' Windows-1252
            Catch
                encoding = System.Text.Encoding.UTF8
            End Try
        End If

        Return encoding.GetString(bytes)
    End Function

    ''' <summary>
    ''' Allows the user to compare two selected text ranges from open documents.
    ''' First prompts for the first selection, then for the second, and performs a text comparison.
    ''' Uses non-modal topmost dialogs to allow Word document access during selection.
    ''' Shows the comparison result in the HTML viewer.
    ''' </summary>
    Private Shared Sub CompareSelectedTextRanges(wordApp As Microsoft.Office.Interop.Word.Application)
        Try
            ' Check if text is already selected - use it as first selection
            Dim firstText As String = Nothing
            Try
                Dim sel1 As Microsoft.Office.Interop.Word.Selection = wordApp.Selection
                If sel1 IsNot Nothing AndAlso sel1.Range IsNot Nothing AndAlso sel1.Start <> sel1.End Then
                    firstText = sel1.Range.Text
                End If
            Catch
            End Try

            ' If no text selected, prompt user to select using non-modal dialog
            If String.IsNullOrWhiteSpace(firstText) Then
                Dim step1Result As Integer = ShowCustomYesNoBox(
                    "Please select the FIRST text range to compare in any open document, then click 'Selection Ready'.",
                    "Selection Ready",
                    "Cancel",
                    $"{AN} Compare Selected - Step 1",
                    nonModal:=True)

                If step1Result <> 1 Then
                    Return ' User cancelled
                End If

                Try
                    Dim sel1 As Microsoft.Office.Interop.Word.Selection = wordApp.Selection
                    If sel1 IsNot Nothing AndAlso sel1.Range IsNot Nothing AndAlso sel1.Start <> sel1.End Then
                        firstText = sel1.Range.Text
                    End If
                Catch
                End Try

                If String.IsNullOrWhiteSpace(firstText) Then
                    ShowCustomMessageBox("No text was selected for the first range. Operation cancelled.", AN)
                    Return
                End If

                ' Now prompt for second selection
                Dim step2Result As Integer = ShowCustomYesNoBox(
                    "First selection captured. Now please select the SECOND text range to compare (can be in the same or different document), then click 'Selection Ready'.",
                    "Selection Ready",
                    "Cancel",
                    $"{AN} Compare Selected - Step 2",
                    nonModal:=True)

                If step2Result <> 1 Then
                    Return ' User cancelled
                End If
            Else
                ' Inform user that we're using the current selection and ask for second
                Dim step2Result As Integer = ShowCustomYesNoBox(
                    $"First selection captured ({firstText.Length} characters).{vbCrLf}{vbCrLf}Now please select the SECOND text range to compare (can be in the same or different document), then click 'Selection Ready'.",
                    "Selection Ready",
                    "Cancel",
                    $"{AN} Compare Selected - Step 2",
                    nonModal:=True)

                If step2Result <> 1 Then
                    Return ' User cancelled
                End If
            End If

            Dim secondText As String = Nothing
            Try
                Dim sel2 As Microsoft.Office.Interop.Word.Selection = wordApp.Selection
                If sel2 IsNot Nothing AndAlso sel2.Range IsNot Nothing AndAlso sel2.Start <> sel2.End Then
                    secondText = sel2.Range.Text
                End If
            Catch
            End Try

            If String.IsNullOrWhiteSpace(secondText) Then
                ShowCustomMessageBox("No text was selected for the second range. Operation cancelled.", AN)
                Return
            End If

            ' Check if both texts are identical
            If firstText = secondText Then
                ShowCustomMessageBox("The two selected text ranges are identical. No differences to show.", AN)
                Return
            End If

            ' Create temporary documents and compare, then show in HTML viewer
            Dim tempDoc1 As Microsoft.Office.Interop.Word.Document = Nothing
            Dim tempDoc2 As Microsoft.Office.Interop.Word.Document = Nothing
            Dim compareDoc As Microsoft.Office.Interop.Word.Document = Nothing
            Dim tempHtmlPath As String = Nothing
            Dim tempFolder As String = Nothing

            Dim prevScreenUpdating As Boolean = wordApp.ScreenUpdating
            Dim prevAlerts As Microsoft.Office.Interop.Word.WdAlertLevel = wordApp.DisplayAlerts

            Try
                wordApp.ScreenUpdating = False
                wordApp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone

                ' Create temporary documents with the selected text
                tempDoc1 = wordApp.Documents.Add(Visible:=False)
                tempDoc1.Content.Text = firstText

                tempDoc2 = wordApp.Documents.Add(Visible:=False)
                tempDoc2.Content.Text = secondText

                ' Compare the temporary documents
                compareDoc = wordApp.CompareDocuments(
                    OriginalDocument:=tempDoc1,
                    RevisedDocument:=tempDoc2,
                    Destination:=WdCompareDestination.wdCompareDestinationNew,
                    Granularity:=WdGranularity.wdGranularityWordLevel,
                    CompareFormatting:=False,
                    CompareCaseChanges:=True,
                    CompareWhitespace:=False,
                    CompareTables:=True,
                    CompareHeaders:=False,
                    CompareFootnotes:=False,
                    CompareTextboxes:=False,
                    CompareFields:=False,
                    CompareComments:=False,
                    CompareMoves:=True,
                    RevisedAuthor:=GetMarkupAuthorOrCurrent(wordApp),
                    IgnoreAllComparisonWarnings:=True
                )

                ' Close temp documents immediately
                Try
                    tempDoc1.Close(WdSaveOptions.wdDoNotSaveChanges)
                Catch
                End Try
                tempDoc1 = Nothing

                Try
                    tempDoc2.Close(WdSaveOptions.wdDoNotSaveChanges)
                Catch
                End Try
                tempDoc2 = Nothing

                If compareDoc Is Nothing Then
                    wordApp.DisplayAlerts = prevAlerts
                    wordApp.ScreenUpdating = prevScreenUpdating
                    ShowCustomMessageBox("Word did not produce a comparison result.", AN)
                    Return
                End If

                ' Keep comparison doc window hidden
                Try
                    If compareDoc.Windows IsNot Nothing AndAlso compareDoc.Windows.Count > 0 Then
                        compareDoc.Windows(1).Visible = False
                    End If
                Catch
                End Try

                ' Extract changes for summarization
                Dim extractedChangesText As String = ExtractChangesWithMarkupTags(compareDoc)

                ' Export to filtered HTML
                tempFolder = Path.Combine(Path.GetTempPath(), $"{AN2}_compare_sel_" & Guid.NewGuid().ToString("N"))
                Directory.CreateDirectory(tempFolder)
                tempHtmlPath = Path.Combine(tempFolder, "comparison.htm")

                compareDoc.SaveAs2(FileName:=tempHtmlPath, FileFormat:=WdSaveFormat.wdFormatFilteredHTML)

                ' Close the comparison document NOW - we have the HTML
                Try
                    compareDoc.Close(WdSaveOptions.wdDoNotSaveChanges)
                Catch
                End Try
                compareDoc = Nothing

                ' Restore UI
                wordApp.DisplayAlerts = prevAlerts
                wordApp.ScreenUpdating = prevScreenUpdating

                ' Capture for closures
                Dim capturedWordApp As Microsoft.Office.Interop.Word.Application = wordApp
                Dim capturedTempFolder As String = tempFolder
                Dim capturedTempHtmlPath As String = tempHtmlPath

                ' Read HTML with proper encoding detection
                Dim htmlContent As String = ReadHtmlWithEncodingDetection(capturedTempHtmlPath)

                ' Inject base href and Mark of the Web for security
                Dim baseHref As String = $"<base href=""file:///{capturedTempFolder.Replace("\", "/")}/"">"
                Dim motw As String = "<!-- saved from url=(0016)http://localhost -->"
                htmlContent = htmlContent.Replace("<head>", "<head>" & vbCrLf & motw & vbCrLf & baseHref)

                ' Build additional buttons
                Dim additionalButtons As New List(Of System.Tuple(Of String, System.Action, Boolean))()

                ' Summarize Changes button
                If Not String.IsNullOrWhiteSpace(extractedChangesText) Then
                    additionalButtons.Add(New System.Tuple(Of String, System.Action, Boolean)(
                        "Summarize Changes",
                        Sub() SummarizeComparisonChangesAsync(extractedChangesText),
                        False))
                End If

                ' Send to PDF button
                Dim desktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                additionalButtons.Add(New System.Tuple(Of String, System.Action, Boolean)(
                    "Send to PDF",
                    Sub()
                        If Not File.Exists(capturedTempHtmlPath) Then
                            ShowCustomMessageBox("The comparison file is no longer available.", AN)
                            Return
                        End If
                        ExportComparisonToPdfFromHtml(capturedTempHtmlPath, capturedWordApp, "Compare_Selection", desktopPath)
                    End Sub,
                    False))

                ' Copy to Clipboard button - copy the comparison WITH formatting/markup
                additionalButtons.Add(New System.Tuple(Of String, System.Action, Boolean)(
                    "Copy to Clipboard",
                    Sub()
                        Try
                            If Not File.Exists(capturedTempHtmlPath) Then
                                ShowCustomMessageBox("The comparison file is no longer available.", AN)
                                Return
                            End If
                            ' Open the HTML file temporarily in Word and copy the content with formatting
                            Dim tempDoc As Microsoft.Office.Interop.Word.Document = Nothing
                            Dim prevScreenUpdating_2nd As Boolean = capturedWordApp.ScreenUpdating
                            Dim prevAlerts_2nd As Microsoft.Office.Interop.Word.WdAlertLevel = capturedWordApp.DisplayAlerts
                            Try
                                capturedWordApp.ScreenUpdating = False
                                capturedWordApp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone

                                tempDoc = capturedWordApp.Documents.Open(
                                    FileName:=capturedTempHtmlPath,
                                    ReadOnly:=True,
                                    Visible:=False)

                                ' Select all content and copy to clipboard (preserves formatting and markup)
                                tempDoc.Content.Copy()

                                tempDoc.Close(WdSaveOptions.wdDoNotSaveChanges)
                                tempDoc = Nothing

                                capturedWordApp.DisplayAlerts = prevAlerts_2nd
                                capturedWordApp.ScreenUpdating = prevScreenUpdating_2nd

                                ShowCustomMessageBox("Comparison copied to clipboard with formatting. You can now paste it into any Word document.", AN)
                            Finally
                                If tempDoc IsNot Nothing Then
                                    Try
                                        tempDoc.Close(WdSaveOptions.wdDoNotSaveChanges)
                                    Catch
                                    End Try
                                End If
                                capturedWordApp.DisplayAlerts = prevAlerts_2nd
                                capturedWordApp.ScreenUpdating = prevScreenUpdating_2nd
                            End Try
                        Catch ex As Exception
                            ShowCustomMessageBox($"Failed to copy to clipboard: {ex.Message}", AN)
                        End Try
                    End Sub,
                    False))

                ' Send to Document button
                additionalButtons.Add(New System.Tuple(Of String, System.Action, Boolean)(
                    "Send to Document",
                    Sub()
                        Try
                            If Not File.Exists(capturedTempHtmlPath) Then
                                ShowCustomMessageBox("The comparison file is no longer available.", AN)
                                Return
                            End If
                            ' Open the HTML file in Word as a new document
                            Dim newDoc As Microsoft.Office.Interop.Word.Document = capturedWordApp.Documents.Open(
                                FileName:=capturedTempHtmlPath,
                                ReadOnly:=False,
                                Visible:=True)
                            newDoc.Activate()
                        Catch ex As Exception
                            ShowCustomMessageBox($"Failed to open in Word: {ex.Message}", AN)
                        End Try
                    End Sub,
                    False))

                ' Compare Selected button (for another comparison)
                additionalButtons.Add(New System.Tuple(Of String, System.Action, Boolean)(
                    "Compare Selected",
                    Sub() CompareSelectedTextRanges(capturedWordApp),
                    False))

                ' Define cleanup action
                Dim cleanupAction As System.Action =
                    Sub()
                        Try
                            If Directory.Exists(capturedTempFolder) Then
                                Directory.Delete(capturedTempFolder, recursive:=True)
                            End If
                        Catch
                            ' Ignore cleanup errors
                        End Try
                    End Sub

                ' Show result in HTML viewer - cleanup happens when dialog closes
                ShowHTMLCustomMessageBox(htmlContent, $"{AN} Text Selection Compare", additionalButtons:=additionalButtons.ToArray(), onClose:=cleanupAction)

            Catch ex As Exception
                ' Cleanup on error
                If compareDoc IsNot Nothing Then
                    Try
                        compareDoc.Close(WdSaveOptions.wdDoNotSaveChanges)
                    Catch
                    End Try
                End If

                wordApp.DisplayAlerts = prevAlerts
                wordApp.ScreenUpdating = prevScreenUpdating

                ShowCustomMessageBox($"Failed to compare selected text: {ex.Message}", AN)
            Finally
                ' Cleanup temp documents if still open
                If tempDoc1 IsNot Nothing Then
                    Try
                        tempDoc1.Close(WdSaveOptions.wdDoNotSaveChanges)
                    Catch
                    End Try
                End If
                If tempDoc2 IsNot Nothing Then
                    Try
                        tempDoc2.Close(WdSaveOptions.wdDoNotSaveChanges)
                    Catch
                    End Try
                End If
            End Try

        Catch ex As Exception
            ShowCustomMessageBox($"Failed to compare selected text: {ex.Message}", AN)
        End Try
    End Sub

    ''' <summary>
    ''' Exports a comparison to PDF by reopening the HTML file as a Word document and exporting it.
    ''' Prompts user for the filename with a default value that can be changed.
    ''' </summary>
    Private Shared Sub ExportComparisonToPdfFromHtml(
        htmlFilePath As String,
        wordApp As Microsoft.Office.Interop.Word.Application,
        defaultFileName As String,
        defaultPath As String)

        If String.IsNullOrEmpty(htmlFilePath) OrElse Not File.Exists(htmlFilePath) Then
            ShowCustomMessageBox("Comparison HTML file is not available.", AN)
            Return
        End If

        Try
            ' Build proposed full path
            Dim proposedFullPath As String = Path.Combine(defaultPath, defaultFileName & ".pdf")

            ' Prompt user for filename
            Dim userInput As String = ShowCustomInputBox(
                "Enter the full path and filename for the PDF (You can change the filename or the entire path):",
                $"{AN} Export to PDF",
                True,
                proposedFullPath)

            If String.IsNullOrWhiteSpace(userInput) Then
                Return ' User cancelled
            End If

            userInput = userInput.Trim()

            ' Ensure it has .pdf extension
            If Not userInput.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) Then
                userInput = userInput & ".pdf"
            End If

            ' Extract directory and ensure it exists
            Dim outputDir As String = Path.GetDirectoryName(userInput)
            If String.IsNullOrEmpty(outputDir) Then
                outputDir = defaultPath
                userInput = Path.Combine(outputDir, Path.GetFileName(userInput))
            End If

            If Not Directory.Exists(outputDir) Then
                Try
                    Directory.CreateDirectory(outputDir)
                Catch ex As Exception
                    ShowCustomMessageBox($"Cannot create directory: {ex.Message}", AN)
                    Return
                End Try
            End If

            Dim outputPath As String = userInput

            ' Ensure unique filename if file exists
            If File.Exists(outputPath) Then
                Dim baseName As String = Path.GetFileNameWithoutExtension(outputPath)
                Dim counter As Integer = 1
                While File.Exists(outputPath)
                    outputPath = Path.Combine(outputDir, $"{baseName}_{counter}.pdf")
                    counter += 1
                End While
            End If

            ' Open the HTML file in Word, export to PDF, then close
            Dim tempDoc As Microsoft.Office.Interop.Word.Document = Nothing
            Dim prevScreenUpdating As Boolean = wordApp.ScreenUpdating
            Dim prevAlerts As Microsoft.Office.Interop.Word.WdAlertLevel = wordApp.DisplayAlerts

            Try
                wordApp.ScreenUpdating = False
                wordApp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone

                ' Open the HTML file
                tempDoc = wordApp.Documents.Open(
                    FileName:=htmlFilePath,
                    ReadOnly:=True,
                    Visible:=False)

                ' Export to PDF
                tempDoc.ExportAsFixedFormat(
                    OutputFileName:=outputPath,
                    ExportFormat:=WdExportFormat.wdExportFormatPDF,
                    OpenAfterExport:=False,
                    OptimizeFor:=WdExportOptimizeFor.wdExportOptimizeForPrint,
                    Range:=WdExportRange.wdExportAllDocument,
                    Item:=WdExportItem.wdExportDocumentWithMarkup,
                    IncludeDocProps:=True,
                    KeepIRM:=True,
                    CreateBookmarks:=WdExportCreateBookmarks.wdExportCreateHeadingBookmarks,
                    DocStructureTags:=True,
                    BitmapMissingFonts:=True,
                    UseISO19005_1:=False)

                ' Close the temp document
                tempDoc.Close(WdSaveOptions.wdDoNotSaveChanges)
                tempDoc = Nothing

            Finally
                If tempDoc IsNot Nothing Then
                    Try
                        tempDoc.Close(WdSaveOptions.wdDoNotSaveChanges)
                    Catch
                    End Try
                End If
                wordApp.DisplayAlerts = prevAlerts
                wordApp.ScreenUpdating = prevScreenUpdating
            End Try

            ' Ask user if they want to open the PDF
            Dim openChoice As Integer = ShowCustomYesNoBox(
                $"PDF exported successfully to:{vbCrLf}{vbCrLf}{outputPath}{vbCrLf}{vbCrLf}Do you want to open it now?",
                "Yes, open PDF",
                "No")

            If openChoice = 1 Then
                Try
                    System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outputPath) With {.UseShellExecute = True})
                Catch ex As Exception
                    ShowCustomMessageBox($"Could not open PDF: {ex.Message}", AN)
                End Try
            End If

        Catch ex As Exception
            ShowCustomMessageBox($"Failed to export PDF: {ex.Message}", AN)
        End Try
    End Sub

    ''' <summary>
    ''' Exports the comparison document to PDF in the same directory as the original document.
    ''' Uses naming convention "Compare_OriginalName_RevisedName.pdf".
    ''' Prompts user for filename.
    ''' </summary>
    Private Shared Sub ExportComparisonToPdf(
        compareDoc As Microsoft.Office.Interop.Word.Document,
        wordApp As Microsoft.Office.Interop.Word.Application,
        originalDocName As String,
        revisedDocName As String,
        originalDocPath As String)

        If compareDoc Is Nothing Then
            ShowCustomMessageBox("Comparison document is not available.", AN)
            Return
        End If

        Try
            ' Determine output directory (same as original document, or Documents folder)
            Dim outputDir As String = originalDocPath
            If String.IsNullOrEmpty(outputDir) OrElse Not Directory.Exists(outputDir) Then
                outputDir = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            End If

            ' Build proposed filename
            Dim sanitizedOriginal As String = SanitizeFileName(originalDocName)
            Dim sanitizedRevised As String = SanitizeFileName(revisedDocName)
            Dim proposedName As String = $"Compare_{sanitizedOriginal}_{sanitizedRevised}.pdf"
            Dim proposedFullPath As String = Path.Combine(outputDir, proposedName)

            ' Prompt user for filename
            Dim userInput As String = ShowCustomInputBox(
                "Enter the full path and filename for the PDF:" & vbCrLf & vbCrLf &
                "(You can change the filename or the entire path)",
                $"{AN} Export to PDF",
                True,
                proposedFullPath)

            If String.IsNullOrWhiteSpace(userInput) Then
                Return ' User cancelled
            End If

            userInput = userInput.Trim()

            ' Ensure it has .pdf extension
            If Not userInput.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) Then
                userInput = userInput & ".pdf"
            End If

            ' Extract directory and ensure it exists
            Dim finalDir As String = Path.GetDirectoryName(userInput)
            If String.IsNullOrEmpty(finalDir) Then
                finalDir = outputDir
                userInput = Path.Combine(finalDir, Path.GetFileName(userInput))
            End If

            If Not Directory.Exists(finalDir) Then
                Try
                    Directory.CreateDirectory(finalDir)
                Catch ex As Exception
                    ShowCustomMessageBox($"Cannot create directory: {ex.Message}", AN)
                    Return
                End Try
            End If

            Dim outputPath As String = userInput

            ' Ensure unique filename if file exists
            If File.Exists(outputPath) Then
                Dim baseName As String = Path.GetFileNameWithoutExtension(outputPath)
                Dim counter As Integer = 1
                While File.Exists(outputPath)
                    outputPath = Path.Combine(finalDir, $"{baseName}_{counter}.pdf")
                    counter += 1
                End While
            End If

            ' Export to PDF
            Dim prevScreenUpdating As Boolean = wordApp.ScreenUpdating
            Dim prevAlerts As Microsoft.Office.Interop.Word.WdAlertLevel = wordApp.DisplayAlerts
            Try
                wordApp.ScreenUpdating = False
                wordApp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone

                compareDoc.ExportAsFixedFormat(
                    OutputFileName:=outputPath,
                    ExportFormat:=WdExportFormat.wdExportFormatPDF,
                    OpenAfterExport:=False,
                    OptimizeFor:=WdExportOptimizeFor.wdExportOptimizeForPrint,
                    Range:=WdExportRange.wdExportAllDocument,
                    Item:=WdExportItem.wdExportDocumentWithMarkup,
                    IncludeDocProps:=True,
                    KeepIRM:=True,
                    CreateBookmarks:=WdExportCreateBookmarks.wdExportCreateHeadingBookmarks,
                    DocStructureTags:=True,
                    BitmapMissingFonts:=True,
                    UseISO19005_1:=False)
            Finally
                wordApp.DisplayAlerts = prevAlerts
                wordApp.ScreenUpdating = prevScreenUpdating
            End Try

            ' Ask user if they want to open the PDF
            Dim openChoice As Integer = ShowCustomYesNoBox(
                $"PDF exported successfully to:{vbCrLf}{vbCrLf}{outputPath}{vbCrLf}{vbCrLf}Do you want to open it now?",
                "Yes, open PDF",
                "No")

            If openChoice = 1 Then
                Try
                    System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outputPath) With {.UseShellExecute = True})
                Catch ex As Exception
                    ShowCustomMessageBox($"Could not open PDF: {ex.Message}", AN)
                End Try
            End If

        Catch ex As Exception
            ShowCustomMessageBox($"Failed to export PDF: {ex.Message}", AN)
        End Try
    End Sub


    ''' <summary>
    ''' Sanitizes a filename by removing invalid characters.
    ''' </summary>
    Private Shared Function SanitizeFileName(name As String) As String
        If String.IsNullOrEmpty(name) Then Return "Document"
        Dim invalidChars As Char() = Path.GetInvalidFileNameChars()
        Dim result As String = name
        For Each c In invalidChars
            result = result.Replace(c, "_"c)
        Next
        ' Also remove spaces for cleaner filenames
        result = result.Replace(" "c, "_"c)
        ' Limit length
        If result.Length > 50 Then result = result.Substring(0, 50)
        Return result
    End Function



    ''' <summary>
    ''' Prompts the user for a file or directory, flattens PDF files into image-only PDFs,
    ''' preserving visual appearance but removing all text layers (useful for security/distribution).
    ''' Shows a progress bar and allows cancellation.
    ''' </summary>
    Public Async Sub FlattenPdfToImages()
        Dim selectedPath As String = ""

        ' Show DragDropForm for file or directory selection
        Globals.ThisAddIn.DragDropFormLabel = "Select a PDF file or folder containing PDFs to flatten"

        Try
            Using frm As New DragDropForm(DragDropMode.FileOrDirectory)
                If frm.ShowDialog() = DialogResult.OK Then
                    selectedPath = frm.SelectedFilePath
                End If
            End Using
        Finally
            Globals.ThisAddIn.DragDropFormLabel = ""
        End Try

        If String.IsNullOrWhiteSpace(selectedPath) Then
            Return
        End If

        ' Determine if it's a file or directory
        Dim isDirectory As Boolean = IO.Directory.Exists(selectedPath)
        Dim isFile As Boolean = IO.File.Exists(selectedPath)

        If Not isFile AndAlso Not isDirectory Then
            ShowCustomMessageBox("The selected path does not exist.")
            Return
        End If

        ' Collect PDF files to process
        Dim filesToProcess As New List(Of String)()

        If isFile Then
            Dim ext As String = IO.Path.GetExtension(selectedPath).ToLowerInvariant()
            If ext = ".pdf" Then
                filesToProcess.Add(selectedPath)
            Else
                ShowCustomMessageBox($"File type '{ext}' is not supported. Only PDF files can be flattened.")
                Return
            End If
        Else
            ' Directory - ask user about recursion
            Dim recurseChoice As Integer = ShowCustomYesNoBox(
            "Do you want to include PDF files from subdirectories?",
            "Yes, include subdirectories",
            "No, top directory only")

            If recurseChoice = 0 Then
                Return ' User aborted
            End If

            Dim searchOption As IO.SearchOption = If(recurseChoice = 1, IO.SearchOption.AllDirectories, IO.SearchOption.TopDirectoryOnly)

            ' Collect all PDF files
            Dim allFiles As String() = IO.Directory.GetFiles(selectedPath, "*.pdf", searchOption)
            filesToProcess.AddRange(allFiles)

            If filesToProcess.Count = 0 Then
                ShowCustomMessageBox("No PDF files found in the selected directory.")
                Return
            End If

            ' Confirm if many files
            If filesToProcess.Count > 10 Then
                Dim confirmAnswer As Integer = ShowCustomYesNoBox(
                $"The directory contains {filesToProcess.Count} PDF files to process. Continue?",
                "Yes, continue", "No, abort")
                If confirmAnswer <> 1 Then
                    Return
                End If
            End If
        End If

        ' Ask user about output location
        Dim useSubdirectory As Boolean = False
        Dim outputSubdirName As String = "flattened"

        If filesToProcess.Count > 1 Then
            Dim outputChoice As Integer = ShowCustomYesNoBox(
            "Where should the flattened PDF files be saved?",
            $"In a subdirectory '{outputSubdirName}'",
            "Same location as original files (with '_flat' suffix)")

            If outputChoice = 0 Then
                Return ' User aborted
            End If
            useSubdirectory = (outputChoice = 1)
        End If

        ' Ask for DPI setting
        ' 200 DPI is recommended for OCR compatibility while keeping file size small
        ' 300 DPI is standard for high-quality OCR
        ' 150 DPI is minimum for readable text
        Dim dpiInput As String = ShowCustomInputBox(
        "Enter the output DPI (dots per inch):" & vbCrLf & vbCrLf &
        "• 150 DPI = Smaller files, lower quality (not recommended for OCR)" & vbCrLf &
        "• 200 DPI = Balanced size/quality (minimum for reliable OCR)" & vbCrLf &
        "• 300 DPI = High quality, larger files (recommended for OCR)",
        AN & " Flatten PDF to Images",
        True,
        "200")

        If dpiInput Is Nothing Then
            Return ' User cancelled
        End If

        Dim dpi As Integer = 200
        If Not Integer.TryParse(dpiInput.Trim(), dpi) OrElse dpi < 72 OrElse dpi > 600 Then
            ShowCustomMessageBox("Invalid DPI value. Please enter a number between 72 and 600.")
            Return
        End If

        ' Create output subdirectory if needed
        Dim outputBaseDir As String = ""
        If useSubdirectory Then
            If isDirectory Then
                outputBaseDir = IO.Path.Combine(selectedPath, outputSubdirName)
            Else
                outputBaseDir = IO.Path.Combine(IO.Path.GetDirectoryName(selectedPath), outputSubdirName)
            End If

            Try
                If Not IO.Directory.Exists(outputBaseDir) Then
                    IO.Directory.CreateDirectory(outputBaseDir)
                End If
            Catch ex As Exception
                ShowCustomMessageBox($"Failed to create output directory: {ex.Message}")
                Return
            End Try
        End If

        ' Show progress bar
        ProgressBarModule.GlobalProgressValue = 0
        ProgressBarModule.GlobalProgressMax = filesToProcess.Count
        ProgressBarModule.GlobalProgressLabel = "Initializing..."
        ProgressBarModule.CancelOperation = False
        ProgressBarModule.ShowProgressBarInSeparateThread(AN & " Flatten PDF to Images", "Starting PDF flattening...")

        ' Process files
        Dim successCount As Integer = 0
        Dim failedFiles As New List(Of String)()

        Try
            For i As Integer = 0 To filesToProcess.Count - 1
                ' Check for cancellation
                If ProgressBarModule.CancelOperation Then
                    Exit For
                End If

                Dim filePath As String = filesToProcess(i)
                Dim fileName As String = IO.Path.GetFileName(filePath)

                ' Update progress
                ProgressBarModule.GlobalProgressValue = i
                ProgressBarModule.GlobalProgressLabel = $"Processing file {i + 1} of {filesToProcess.Count}: {fileName}"

                Try
                    ' Determine output path
                    Dim outputPath As String
                    Dim nameWithoutExt As String = IO.Path.GetFileNameWithoutExtension(filePath)

                    If useSubdirectory Then
                        ' Preserve relative directory structure if processing subdirectories
                        If isDirectory Then
                            Dim relativePath As String = filePath.Substring(selectedPath.Length).TrimStart(IO.Path.DirectorySeparatorChar)
                            Dim relativeDir As String = IO.Path.GetDirectoryName(relativePath)
                            If Not String.IsNullOrEmpty(relativeDir) Then
                                Dim targetDir As String = IO.Path.Combine(outputBaseDir, relativeDir)
                                If Not IO.Directory.Exists(targetDir) Then
                                    IO.Directory.CreateDirectory(targetDir)
                                End If
                            End If
                            outputPath = IO.Path.Combine(outputBaseDir, IO.Path.ChangeExtension(relativePath, ".pdf"))
                        Else
                            outputPath = IO.Path.Combine(outputBaseDir, fileName)
                        End If
                    Else
                        ' Same location with suffix
                        Dim dir As String = IO.Path.GetDirectoryName(filePath)
                        outputPath = IO.Path.Combine(dir, nameWithoutExt & "_flat.pdf")
                    End If

                    ' Flatten the PDF
                    Await System.Threading.Tasks.Task.Run(Sub() FlattenPdfToImageOnly(filePath, outputPath, dpi))
                    successCount += 1

                Catch ex As Exception
                    failedFiles.Add($"{fileName}: {ex.Message}")
                End Try
            Next

        Finally
            ' Close progress bar
            ProgressBarModule.CancelOperation = True
        End Try

        ' Build summary
        Dim wasCancelled As Boolean = (successCount + failedFiles.Count) < filesToProcess.Count

        Dim summary As New System.Text.StringBuilder()

        If wasCancelled Then
            summary.AppendLine("Operation was cancelled by user.")
            summary.AppendLine()
        End If

        summary.AppendLine($"Successfully flattened: {successCount} file(s)")
        summary.AppendLine($"Output DPI: {dpi}")

        If useSubdirectory Then
            summary.AppendLine($"Output directory: {outputBaseDir}")
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

            ' Copy detailed log to clipboard
            Dim clipboardLog As New System.Text.StringBuilder()
            clipboardLog.AppendLine("=== FAILED FILES ===")
            For Each f In failedFiles
                clipboardLog.AppendLine(f)
            Next
            SharedMethods.PutInClipboard(clipboardLog.ToString().TrimEnd())
            summary.AppendLine()
            summary.AppendLine("(Detailed log copied to clipboard)")
        End If

        ShowCustomMessageBox(summary.ToString().TrimEnd(), AN & " Flatten PDF to Images")
    End Sub





    ''' <summary>
    ''' Converts a PDF to an image-only PDF by rasterizing each page via the Windows 10+
    ''' built-in PDF renderer (Windows.Data.Pdf). This API is maintained by Microsoft via
    ''' Windows Update, handles all modern PDF features (transparency groups, layered image
    ''' XObjects, form XObjects, compositing modes), and requires no third-party native DLLs.
    ''' Falls back to PdfiumViewer if the Windows API is unavailable.
    ''' </summary>
    ''' <param name="inputPath">Input PDF file path</param>
    ''' <param name="outputPath">Output PDF file path</param>
    ''' <param name="dpi">Rasterization DPI (default 200 for balance of size/quality)</param>
    Private Shared Sub FlattenPdfToImageOnly(inputPath As String, outputPath As String, Optional dpi As Integer = 200)

        Dim dbg As Action(Of String) = Sub(msg) System.Diagnostics.Debug.WriteLine($"[FlattenPDF] {msg}")
        dbg($"=== FlattenPdfToImageOnly START === dpi={dpi}")
        dbg($"Input:  {inputPath}")
        dbg($"Output: {outputPath}")

        Dim usedWindowsPdf As Boolean = False

#If HAS_WINRT Then
        If Environment.OSVersion.Version.Major >= 10 Then
            Try
                FlattenPdfViaWindowsApi(inputPath, outputPath, dpi)
                usedWindowsPdf = True
                dbg("Rendered via Windows.Data.Pdf — all visual elements preserved")
            Catch ex As Exception
                dbg($"Windows.Data.Pdf failed: {ex.Message} — falling back to PdfiumViewer")
                usedWindowsPdf = False
            End Try
        Else
            dbg($"OS version {Environment.OSVersion.Version} — Windows.Data.Pdf not available, using PdfiumViewer")
        End If
#End If

        If Not usedWindowsPdf Then
            dbg("WARNING: PdfiumViewer fallback may not preserve layered image objects with transparency")
            FlattenPdfToImageOnlyViaPdfium(inputPath, outputPath, dpi)
        End If

        dbg("=== FlattenPdfToImageOnly END ===")
    End Sub

    Private Shared Sub FlattenPdfViaWindowsApi(inputPath As String, outputPath As String, dpi As Integer)
#If Not HAS_WINRT Then
        Throw New PlatformNotSupportedException("Windows.Data.Pdf is not available (SDK not installed at build time).")
#Else
        Dim dbg As Action(Of String) = Sub(msg) System.Diagnostics.Debug.WriteLine($"[FlattenPDF-WinPdf] {msg}")

        Dim renderException As Exception = Nothing

        Dim staThread As New System.Threading.Thread(
            Sub()
                Try
                    FlattenPdfViaWindowsApiCore(inputPath, outputPath, dpi, dbg)
                Catch ex As Exception
                    renderException = ex
                End Try
            End Sub)

        staThread.SetApartmentState(System.Threading.ApartmentState.STA)
        staThread.IsBackground = True
        staThread.Start()

        If Not staThread.Join(TimeSpan.FromMinutes(5)) Then
            Try : staThread.Abort() : Catch : End Try
            Throw New TimeoutException("Windows.Data.Pdf rendering timed out after 5 minutes.")
        End If

        If renderException IsNot Nothing Then
            Throw renderException
        End If
#End If
    End Sub

    Private Shared Sub FlattenPdfViaWindowsApiCore(inputPath As String, outputPath As String,
                                                    dpi As Integer,
                                                    dbg As Action(Of String))
#If Not HAS_WINRT Then
        Throw New PlatformNotSupportedException("Windows.Data.Pdf is not available (SDK not installed at build time).")
#Else
        ' Open the PDF via Windows.Data.Pdf
        Dim storageFile As Windows.Storage.StorageFile =
            Windows.Storage.StorageFile.GetFileFromPathAsync(inputPath).GetAwaiter().GetResult()

        Dim winPdf As Windows.Data.Pdf.PdfDocument =
            Windows.Data.Pdf.PdfDocument.LoadFromFileAsync(storageFile).GetAwaiter().GetResult()

        dbg($"Opened via Windows.Data.Pdf — {winPdf.PageCount} page(s)")

        Dim outDoc As New PdfSharp.Pdf.PdfDocument()

        ' Copy metadata from source
        Try
            Using srcDoc As PdfSharp.Pdf.PdfDocument = PdfSharp.Pdf.IO.PdfReader.Open(inputPath, PdfSharp.Pdf.IO.PdfDocumentOpenMode.InformationOnly)
                CopyPdfMetadata(srcDoc, outDoc)
            End Using
        Catch
        End Try

        ' Scale factor: Windows.Data.Pdf renders at 96 DPI by default.
        Dim scaleFactor As Double = dpi / 96.0

        For pageIndex As UInteger = 0 To CUInt(winPdf.PageCount - 1)
            Dim page As Windows.Data.Pdf.PdfPage = winPdf.GetPage(pageIndex)

            Dim pageWidthPt As Double = page.Size.Width * 72.0 / 96.0
            Dim pageHeightPt As Double = page.Size.Height * 72.0 / 96.0

            Dim renderWidthPx As UInteger = CUInt(System.Math.Round(page.Size.Width * scaleFactor))
            Dim renderHeightPx As UInteger = CUInt(System.Math.Round(page.Size.Height * scaleFactor))

            dbg($"Page {pageIndex + 1}: {pageWidthPt:F0}x{pageHeightPt:F0} pts, render {renderWidthPx}x{renderHeightPx} px")

            Using renderStream As New Windows.Storage.Streams.InMemoryRandomAccessStream()
                Dim renderOptions As New Windows.Data.Pdf.PdfPageRenderOptions()
                renderOptions.DestinationWidth = renderWidthPx
                renderOptions.DestinationHeight = renderHeightPx
                renderOptions.BitmapEncoderId = Windows.Graphics.Imaging.BitmapEncoder.PngEncoderId

                page.RenderToStreamAsync(renderStream, renderOptions).GetAwaiter().GetResult()

                renderStream.Seek(0)
                Dim netStream As System.IO.Stream = System.IO.WindowsRuntimeStreamExtensions.AsStreamForRead(renderStream)

                Using ms As New System.IO.MemoryStream()
                    netStream.CopyTo(ms)
                    ms.Position = 0

                    Using bmp As New System.Drawing.Bitmap(ms)
                        Using jpegMs As New System.IO.MemoryStream()
                            Dim jpegEncoder As System.Drawing.Imaging.ImageCodecInfo = GetJpegEncoder()
                            If jpegEncoder IsNot Nothing Then
                                Dim encoderParams As New System.Drawing.Imaging.EncoderParameters(1)
                                encoderParams.Param(0) = New System.Drawing.Imaging.EncoderParameter(
                                    System.Drawing.Imaging.Encoder.Quality, 85L)
                                bmp.Save(jpegMs, jpegEncoder, encoderParams)
                            Else
                                bmp.Save(jpegMs, System.Drawing.Imaging.ImageFormat.Png)
                            End If

                            jpegMs.Position = 0

                            Dim outPage As PdfSharp.Pdf.PdfPage = outDoc.AddPage()
                            outPage.Width = PdfSharp.Drawing.XUnit.FromPoint(pageWidthPt)
                            outPage.Height = PdfSharp.Drawing.XUnit.FromPoint(pageHeightPt)

                            Using xgfx As PdfSharp.Drawing.XGraphics = PdfSharp.Drawing.XGraphics.FromPdfPage(outPage)
                                Using ximg As PdfSharp.Drawing.XImage = PdfSharp.Drawing.XImage.FromStream(jpegMs)
                                    xgfx.DrawImage(ximg, 0, 0, outPage.Width.Point, outPage.Height.Point)
                                End Using
                            End Using
                        End Using
                    End Using
                End Using
            End Using

            page.Dispose()
        Next

        outDoc.Save(outputPath)
        outDoc.Close()

        dbg($"Output saved: {New System.IO.FileInfo(outputPath).Length:N0} bytes")
#End If
    End Sub


    ''' <summary>
    ''' Renders a single page of a PDF to a JPEG MemoryStream using Windows.Data.Pdf (Win10+).
    ''' Returns Nothing if the API is unavailable or fails (caller should fall back to PdfiumViewer).
    ''' </summary>
    Friend Shared Function RenderPageViaWindowsPdf(inputPath As String, pageIndex As Integer, dpi As Integer) As Tuple(Of MemoryStream, Double, Double)
#If Not HAS_WINRT Then
        Return Nothing
#Else
        If Environment.OSVersion.Version.Major < 10 Then Return Nothing

        Dim result As Tuple(Of MemoryStream, Double, Double) = Nothing
        Dim renderException As Exception = Nothing

        Dim staThread As New System.Threading.Thread(
            Sub()
                Try
                    result = RenderPageViaWindowsPdfCore(inputPath, pageIndex, dpi)
                Catch ex As Exception
                    renderException = ex
                End Try
            End Sub)

        staThread.SetApartmentState(System.Threading.ApartmentState.STA)
        staThread.IsBackground = True
        staThread.Start()

        If Not staThread.Join(TimeSpan.FromMinutes(2)) Then
            Try : staThread.Abort() : Catch : End Try
            Return Nothing
        End If

        If renderException IsNot Nothing Then Return Nothing
        Return result
#End If
    End Function

    ''' <summary>
    ''' Core STA implementation for single-page Windows.Data.Pdf rendering.
    ''' </summary>
    Private Shared Function RenderPageViaWindowsPdfCore(inputPath As String, pageIndex As Integer, dpi As Integer) As Tuple(Of MemoryStream, Double, Double)
#If Not HAS_WINRT Then
        Return Nothing
#Else
        Dim storageFile As Windows.Storage.StorageFile =
            Windows.Storage.StorageFile.GetFileFromPathAsync(inputPath).GetAwaiter().GetResult()

        Dim winPdf As Windows.Data.Pdf.PdfDocument =
            Windows.Data.Pdf.PdfDocument.LoadFromFileAsync(storageFile).GetAwaiter().GetResult()

        If pageIndex < 0 OrElse pageIndex >= CInt(winPdf.PageCount) Then Return Nothing

        Dim page As Windows.Data.Pdf.PdfPage = winPdf.GetPage(CUInt(pageIndex))

        Dim pageWidthPt As Double = page.Size.Width * 72.0 / 96.0
        Dim pageHeightPt As Double = page.Size.Height * 72.0 / 96.0

        Dim scaleFactor As Double = dpi / 96.0
        Dim renderWidthPx As UInteger = CUInt(System.Math.Round(page.Size.Width * scaleFactor))
        Dim renderHeightPx As UInteger = CUInt(System.Math.Round(page.Size.Height * scaleFactor))

        Using renderStream As New Windows.Storage.Streams.InMemoryRandomAccessStream()
            Dim renderOptions As New Windows.Data.Pdf.PdfPageRenderOptions()
            renderOptions.DestinationWidth = renderWidthPx
            renderOptions.DestinationHeight = renderHeightPx
            renderOptions.BitmapEncoderId = Windows.Graphics.Imaging.BitmapEncoder.PngEncoderId

            page.RenderToStreamAsync(renderStream, renderOptions).GetAwaiter().GetResult()

            renderStream.Seek(0)
            Dim netStream As System.IO.Stream = System.IO.WindowsRuntimeStreamExtensions.AsStreamForRead(renderStream)

            Using pngMs As New MemoryStream()
                netStream.CopyTo(pngMs)
                pngMs.Position = 0

                Using bmp As New System.Drawing.Bitmap(pngMs)
                    Dim jpegMs As New MemoryStream()
                    Dim jpegEncoder As System.Drawing.Imaging.ImageCodecInfo = GetJpegEncoder()
                    If jpegEncoder IsNot Nothing Then
                        Dim encoderParams As New System.Drawing.Imaging.EncoderParameters(1)
                        encoderParams.Param(0) = New System.Drawing.Imaging.EncoderParameter(
                            System.Drawing.Imaging.Encoder.Quality, 85L)
                        bmp.Save(jpegMs, jpegEncoder, encoderParams)
                    Else
                        bmp.Save(jpegMs, System.Drawing.Imaging.ImageFormat.Png)
                    End If
                    jpegMs.Position = 0

                    page.Dispose()
                    Return Tuple.Create(jpegMs, pageWidthPt, pageHeightPt)
                End Using
            End Using
        End Using
#End If
    End Function

    ''' <summary>
    ''' Renders all pages of a PDF to a new image-only PdfSharp document using Windows.Data.Pdf.
    ''' Returns Nothing if the API is unavailable or fails.
    ''' </summary>
    Friend Shared Function RenderAllPagesViaWindowsPdf(inputPath As String, dpi As Integer) As PdfSharp.Pdf.PdfDocument
#If Not HAS_WINRT Then
        Return Nothing
#Else
        If Environment.OSVersion.Version.Major < 10 Then Return Nothing

        Dim result As PdfSharp.Pdf.PdfDocument = Nothing
        Dim renderException As Exception = Nothing

        Dim staThread As New System.Threading.Thread(
            Sub()
                Try
                    result = RenderAllPagesViaWindowsPdfCore(inputPath, dpi)
                Catch ex As Exception
                    renderException = ex
                End Try
            End Sub)

        staThread.SetApartmentState(System.Threading.ApartmentState.STA)
        staThread.IsBackground = True
        staThread.Start()

        If Not staThread.Join(TimeSpan.FromMinutes(5)) Then
            Try : staThread.Abort() : Catch : End Try
            Return Nothing
        End If

        If renderException IsNot Nothing Then Return Nothing
        Return result
#End If
    End Function

    ''' <summary>
    ''' Core STA implementation for full-document Windows.Data.Pdf rendering.
    ''' </summary>
    Private Shared Function RenderAllPagesViaWindowsPdfCore(inputPath As String, dpi As Integer) As PdfSharp.Pdf.PdfDocument
#If Not HAS_WINRT Then
        Return Nothing
#Else
        Dim storageFile As Windows.Storage.StorageFile =
            Windows.Storage.StorageFile.GetFileFromPathAsync(inputPath).GetAwaiter().GetResult()

        Dim winPdf As Windows.Data.Pdf.PdfDocument =
            Windows.Data.Pdf.PdfDocument.LoadFromFileAsync(storageFile).GetAwaiter().GetResult()

        Dim outDoc As New PdfSharp.Pdf.PdfDocument()
        Dim scaleFactor As Double = dpi / 96.0

        For pageIndex As UInteger = 0 To CUInt(winPdf.PageCount - 1)
            Dim page As Windows.Data.Pdf.PdfPage = winPdf.GetPage(pageIndex)

            Dim pageWidthPt As Double = page.Size.Width * 72.0 / 96.0
            Dim pageHeightPt As Double = page.Size.Height * 72.0 / 96.0

            Dim renderWidthPx As UInteger = CUInt(System.Math.Round(page.Size.Width * scaleFactor))
            Dim renderHeightPx As UInteger = CUInt(System.Math.Round(page.Size.Height * scaleFactor))

            Using renderStream As New Windows.Storage.Streams.InMemoryRandomAccessStream()
                Dim renderOptions As New Windows.Data.Pdf.PdfPageRenderOptions()
                renderOptions.DestinationWidth = renderWidthPx
                renderOptions.DestinationHeight = renderHeightPx
                renderOptions.BitmapEncoderId = Windows.Graphics.Imaging.BitmapEncoder.PngEncoderId

                page.RenderToStreamAsync(renderStream, renderOptions).GetAwaiter().GetResult()

                renderStream.Seek(0)
                Dim netStream As System.IO.Stream = System.IO.WindowsRuntimeStreamExtensions.AsStreamForRead(renderStream)

                Using pngMs As New MemoryStream()
                    netStream.CopyTo(pngMs)
                    pngMs.Position = 0

                    Using bmp As New System.Drawing.Bitmap(pngMs)
                        Using jpegMs As New MemoryStream()
                            Dim jpegEncoder As System.Drawing.Imaging.ImageCodecInfo = GetJpegEncoder()
                            If jpegEncoder IsNot Nothing Then
                                Dim encoderParams As New System.Drawing.Imaging.EncoderParameters(1)
                                encoderParams.Param(0) = New System.Drawing.Imaging.EncoderParameter(
                                    System.Drawing.Imaging.Encoder.Quality, 85L)
                                bmp.Save(jpegMs, jpegEncoder, encoderParams)
                            Else
                                bmp.Save(jpegMs, System.Drawing.Imaging.ImageFormat.Png)
                            End If

                            jpegMs.Position = 0

                            Dim outPage As PdfSharp.Pdf.PdfPage = outDoc.AddPage()
                            outPage.Width = PdfSharp.Drawing.XUnit.FromPoint(pageWidthPt)
                            outPage.Height = PdfSharp.Drawing.XUnit.FromPoint(pageHeightPt)

                            Using xgfx As PdfSharp.Drawing.XGraphics = PdfSharp.Drawing.XGraphics.FromPdfPage(outPage)
                                Using ximg As PdfSharp.Drawing.XImage = PdfSharp.Drawing.XImage.FromStream(jpegMs)
                                    xgfx.DrawImage(ximg, 0, 0, outPage.Width.Point, outPage.Height.Point)
                                End Using
                            End Using
                        End Using
                    End Using
                End Using
            End Using

            page.Dispose()
        Next

        Return outDoc
#End If
    End Function




    <System.Runtime.InteropServices.DllImport("kernel32.dll", CharSet:=System.Runtime.InteropServices.CharSet.Auto, SetLastError:=True)>
    Private Shared Function LoadLibrary(lpFileName As String) As IntPtr
    End Function

    ''' <summary>
    ''' Rasterizes a PDF to an image-only PDF via PdfiumViewer.
    ''' Used as fallback if Windows.Data.Pdf is unavailable (pre-Windows 10).
    ''' </summary>
    Private Shared Sub FlattenPdfToImageOnlyViaPdfium(inputPath As String, outputPath As String,
                                                      Optional dpi As Integer = 200,
                                                      Optional metadataSourcePath As String = Nothing)

        PdfRedactionService.EnsurePdfiumLoadedPublic()

        Using pdf As PdfiumViewer.PdfDocument = PdfiumViewer.PdfDocument.Load(inputPath)
            Dim outDoc As New PdfSharp.Pdf.PdfDocument()

            Dim metaPath As String = If(String.IsNullOrEmpty(metadataSourcePath), inputPath, metadataSourcePath)
            Try
                Using srcDoc As PdfSharp.Pdf.PdfDocument = PdfSharp.Pdf.IO.PdfReader.Open(metaPath, PdfSharp.Pdf.IO.PdfDocumentOpenMode.InformationOnly)
                    CopyPdfMetadata(srcDoc, outDoc)
                End Using
            Catch
            End Try

            For pageIndex As Integer = 0 To pdf.PageCount - 1
                Dim sizePt As System.Drawing.SizeF = pdf.PageSizes(pageIndex)
                Dim widthPx As Integer = CInt(System.Math.Round(sizePt.Width / 72.0 * dpi))
                Dim heightPx As Integer = CInt(System.Math.Round(sizePt.Height / 72.0 * dpi))

                Dim renderFlags As PdfiumViewer.PdfRenderFlags =
                    PdfiumViewer.PdfRenderFlags.Annotations Or
                    PdfiumViewer.PdfRenderFlags.LcdText Or
                    PdfiumViewer.PdfRenderFlags.ForPrinting

                Using rendered As System.Drawing.Image = pdf.Render(pageIndex, widthPx, heightPx, dpi, dpi, renderFlags)
                    Dim outPage As PdfSharp.Pdf.PdfPage = outDoc.AddPage()
                    outPage.Width = PdfSharp.Drawing.XUnit.FromPoint(sizePt.Width)
                    outPage.Height = PdfSharp.Drawing.XUnit.FromPoint(sizePt.Height)

                    Using ms As New System.IO.MemoryStream()
                        Dim jpegEncoder As System.Drawing.Imaging.ImageCodecInfo = GetJpegEncoder()
                        If jpegEncoder IsNot Nothing Then
                            Dim encoderParams As New System.Drawing.Imaging.EncoderParameters(1)
                            encoderParams.Param(0) = New System.Drawing.Imaging.EncoderParameter(
                                System.Drawing.Imaging.Encoder.Quality, 85L)
                            rendered.Save(ms, jpegEncoder, encoderParams)
                        Else
                            rendered.Save(ms, System.Drawing.Imaging.ImageFormat.Png)
                        End If

                        ms.Position = 0
                        Using xgfx As PdfSharp.Drawing.XGraphics = PdfSharp.Drawing.XGraphics.FromPdfPage(outPage)
                            Using ximg As PdfSharp.Drawing.XImage = PdfSharp.Drawing.XImage.FromStream(ms)
                                xgfx.DrawImage(ximg, 0, 0, outPage.Width.Point, outPage.Height.Point)
                            End Using
                        End Using
                    End Using
                End Using
            Next

            outDoc.Save(outputPath)
            outDoc.Close()
        End Using
    End Sub



    ''' <summary>
    ''' Gets the JPEG image encoder for file size optimization.
    ''' </summary>
    Private Shared Function GetJpegEncoder() As System.Drawing.Imaging.ImageCodecInfo
        Dim codecs As System.Drawing.Imaging.ImageCodecInfo() = System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders()
        For Each codec In codecs
            If codec.MimeType = "image/jpeg" Then
                Return codec
            End If
        Next
        Return Nothing
    End Function

    ''' <summary>
    ''' Copies PDF metadata from source to destination document.
    ''' </summary>
    ''' <summary>
    ''' Copies PDF metadata from source to destination document.
    ''' Handles DateTimeKind conversion to avoid DateTimeInvalidLocalFormat exceptions.
    ''' </summary>
    Private Shared Sub CopyPdfMetadata(src As PdfSharp.Pdf.PdfDocument, dest As PdfSharp.Pdf.PdfDocument)
        If src Is Nothing OrElse dest Is Nothing Then Return

        Try
            dest.Info.Title = src.Info.Title
            dest.Info.Author = src.Info.Author
            dest.Info.Subject = src.Info.Subject
            dest.Info.Keywords = src.Info.Keywords
            dest.Info.Creator = src.Info.Creator

            ' Handle CreationDate - convert UTC to Local to avoid DateTimeInvalidLocalFormat MDA
            If src.Info.CreationDate <> Date.MinValue Then
                Dim creationDate As DateTime = src.Info.CreationDate
                If creationDate.Kind = DateTimeKind.Utc Then
                    creationDate = DateTime.SpecifyKind(creationDate.ToLocalTime(), DateTimeKind.Local)
                ElseIf creationDate.Kind = DateTimeKind.Unspecified Then
                    creationDate = DateTime.SpecifyKind(creationDate, DateTimeKind.Local)
                End If
                dest.Info.CreationDate = creationDate
            End If

            ' Handle ModificationDate - convert UTC to Local to avoid DateTimeInvalidLocalFormat MDA
            If src.Info.ModificationDate <> Date.MinValue Then
                Dim modificationDate As DateTime = src.Info.ModificationDate
                If modificationDate.Kind = DateTimeKind.Utc Then
                    modificationDate = DateTime.SpecifyKind(modificationDate.ToLocalTime(), DateTimeKind.Local)
                ElseIf modificationDate.Kind = DateTimeKind.Unspecified Then
                    modificationDate = DateTime.SpecifyKind(modificationDate, DateTimeKind.Local)
                End If
                dest.Info.ModificationDate = modificationDate
            End If
        Catch
            ' Ignore metadata copy errors
        End Try
    End Sub



    ''' <summary>
    ''' Prompts the user for a file or directory, reads content using GetFileContent(),
    ''' and saves the extracted text as a .txt file at the same location (or optionally in a subdirectory).
    ''' Shows a progress bar and allows cancellation. If 2+ PDF files are present and OCR is available,
    ''' prompts user once for OCR preference. Optionally flattens PDFs to image-only before OCR
    ''' to ensure all visual elements (layered images, transparency groups) are captured.
    ''' </summary>
    Public Async Sub ExportFileContentToText()
        Dim selectedPath As String = ""

        ' Show DragDropForm for file or directory selection
        Globals.ThisAddIn.DragDropFormLabel = "Select a file or folder to convert to text"

        Try
            Using frm As New DragDropForm(DragDropMode.FileOrDirectory)
                If frm.ShowDialog() = DialogResult.OK Then
                    selectedPath = frm.SelectedFilePath
                End If
            End Using
        Finally
            Globals.ThisAddIn.DragDropFormLabel = ""
        End Try

        If String.IsNullOrWhiteSpace(selectedPath) Then
            Return
        End If

        ' Determine if it's a file or directory
        Dim isDirectory As Boolean = IO.Directory.Exists(selectedPath)
        Dim isFile As Boolean = IO.File.Exists(selectedPath)

        If Not isFile AndAlso Not isDirectory Then
            ShowCustomMessageBox("The selected path does not exist.")
            Return
        End If

        ' Build supported extensions list aligned with GetFileContentEx
        Dim supportedExtensionsList As New List(Of String)({
            ".txt", ".rtf", ".docx", ".pdf", ".xlsx", ".pptx", ".ini", ".csv", ".log",
            ".json", ".xml", ".html", ".htm", ".md", ".yaml", ".yml",
            ".vb", ".cs", ".js", ".ts", ".py", ".java", ".cpp", ".c", ".h", ".sql",
            ".eml", ".msg",
            ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff", ".tif", ".webp", ".svg",
            ".mp3", ".wav", ".ogg", ".flac", ".m4a", ".aac", ".wma", ".opus", ".webm",
            ".mp4", ".avi", ".mkv", ".mov", ".wmv"
        })

        ' Conditionally include .doc (legacy Word format — may execute macros via COM)
        If INI_AllowLegacyDocFiles Then
            supportedExtensionsList.Add(".doc")
        End If

        ' Filter out binary/media types the current model can't handle
        Dim supportedExtensions As String() = supportedExtensionsList.
            Where(Function(ext) Not IsBinaryMediaExtension(ext) OrElse IsModelCapableForExtension(_context, ext)).
            ToArray()

        ' Collect files to process
        Dim filesToProcess As New List(Of String)()
        Dim unsupportedFiles As New List(Of String)()
        Dim legacyDocFilesSkipped As New List(Of String)()

        If isFile Then
            Dim ext As String = IO.Path.GetExtension(selectedPath).ToLowerInvariant()
            If supportedExtensions.Contains(ext) Then
                filesToProcess.Add(selectedPath)
            ElseIf ext = ".doc" AndAlso Not INI_AllowLegacyDocFiles Then
                ShowCustomMessageBox($"Legacy .doc files are disabled for security reasons (AllowLegacyDocFiles is off). Save as .docx first.")
                Return
            Else
                ShowCustomMessageBox($"File type '{ext}' is not supported for text extraction.")
                Return
            End If
        Else
            ' Directory - ask user about recursion
            Dim recurseChoice As Integer = ShowCustomYesNoBox(
                "Do you want to include files from subdirectories?",
                "Yes, include subdirectories",
                "No, top directory only")

            If recurseChoice = 0 Then
                Return ' User aborted
            End If

            Dim searchOption As IO.SearchOption = If(recurseChoice = 1, IO.SearchOption.AllDirectories, IO.SearchOption.TopDirectoryOnly)

            ' Collect all files
            Dim allFiles As String() = IO.Directory.GetFiles(selectedPath, "*.*", searchOption)

            For Each f In allFiles
                Dim ext As String = IO.Path.GetExtension(f).ToLowerInvariant()
                If supportedExtensions.Contains(ext) Then
                    filesToProcess.Add(f)
                ElseIf ext = ".doc" AndAlso Not INI_AllowLegacyDocFiles Then
                    legacyDocFilesSkipped.Add(f)
                Else
                    unsupportedFiles.Add(f)
                End If
            Next

            If filesToProcess.Count = 0 Then
                ShowCustomMessageBox($"No supported files found in the selected directory." &
                    If(unsupportedFiles.Count > 0, $" ({unsupportedFiles.Count} unsupported file(s) were ignored.)", "") &
                    If(legacyDocFilesSkipped.Count > 0, $" ({legacyDocFilesSkipped.Count} .doc file(s) skipped — AllowLegacyDocFiles is off.)", ""))
                Return
            End If

            ' Confirm if many files
            If filesToProcess.Count > 10 Then
                Dim confirmAnswer As Integer = ShowCustomYesNoBox(
                    $"The directory contains {filesToProcess.Count} supported files to process." &
                    If(unsupportedFiles.Count > 0, $" ({unsupportedFiles.Count} unsupported files will be ignored.)", "") &
                    If(legacyDocFilesSkipped.Count > 0, $" ({legacyDocFilesSkipped.Count} .doc files skipped — disabled for security.)", "") &
                    " Continue?",
                    "Yes, continue", "No, abort")
                If confirmAnswer <> 1 Then
                    Return
                End If
            End If
        End If

        ' Ask user about output location
        Dim useSubdirectory As Boolean = False
        Dim outputSubdirName As String = "extracted_text"

        If filesToProcess.Count > 1 Then
            Dim outputChoice As Integer = ShowCustomYesNoBox(
                "Where should the extracted text files be saved?",
                $"In a subdirectory '{outputSubdirName}'",
                "Same location as original files")

            If outputChoice = 0 Then
                Return ' User aborted
            End If
            useSubdirectory = (outputChoice = 1)
        End If

        ' Count PDF files to determine OCR behavior
        Dim pdfFiles As List(Of String) = filesToProcess.Where(
            Function(f) IO.Path.GetExtension(f).Equals(".pdf", StringComparison.OrdinalIgnoreCase)
        ).ToList()
        Dim pdfCount As Integer = pdfFiles.Count

        ' Determine OCR settings
        Dim doOcr As Boolean = False
        Dim askUserPerFile As Boolean = False
        Dim flattenBeforeOcr As Boolean = False
        Dim useMarkdownOutputForFlattenedOcr As Boolean = False
        Dim ocrMarkdownInstruction As String = String.Empty

        If pdfCount >= 2 AndAlso SharedMethods.IsOcrAvailable(_context) Then
            Dim ocrChoice As Integer = ShowCustomYesNoBox(
                $"There are {pdfCount} PDF files to process." & vbCrLf & vbCrLf &
                "Some PDFs may require OCR (optical character recognition) to extract text from scanned documents or images." & vbCrLf & vbCrLf &
                "How would you like to handle OCR for these PDF files?",
                "Enable OCR for all PDFs",
                "Skip OCR for all PDFs")

            If ocrChoice = 0 Then
                Return ' User aborted
            ElseIf ocrChoice = 1 Then
                doOcr = True
                askUserPerFile = False
            Else
                doOcr = False
                askUserPerFile = False
            End If
        ElseIf pdfCount = 1 AndAlso SharedMethods.IsOcrAvailable(_context) Then
            ' Single PDF - still ask once upfront, but don't interrupt during processing
            Dim ocrChoice As Integer = ShowCustomYesNoBox(
                "The PDF may require OCR to extract text from scanned content." & vbCrLf & vbCrLf &
                "Do you want to enable OCR?",
                "Yes, enable OCR",
                "No, skip OCR")
            If ocrChoice = 0 Then
                Return ' User aborted
            End If
            doOcr = (ocrChoice = 1)
            askUserPerFile = False ' User already decided — don't interrupt during processing
        End If

        ' Offer PDF flattening before OCR when OCR is enabled
        If doOcr AndAlso pdfCount >= 1 Then
            Dim flattenChoice As Integer = ShowCustomYesNoBox(
                "Some scanned PDFs contain layered images or transparency groups that OCR may miss." & vbCrLf & vbCrLf &
                "Would you like to flatten PDFs to image-only format before OCR? " &
                "This rasterizes each page (like printing to image), ensuring all visual elements are captured." & vbCrLf & vbCrLf &
                "⚠ This adds processing time (a few seconds per page) but can significantly improve OCR results for complex PDFs.",
                "Yes, flatten first (recommended for scans)",
                "No, OCR directly")
            If flattenChoice = 0 Then
                Return ' User aborted
            End If
            flattenBeforeOcr = (flattenChoice = 1)
        End If

        ' Only ask about Markdown if OCR is enabled AND PDFs will be flattened first
        If doOcr AndAlso flattenBeforeOcr Then
            Dim markdownChoice As Integer = ShowCustomYesNoBox(
                "Do you want OCR results for those flattened PDFs to be saved as Markdown files (.md) instead of plain text (.txt)?" & vbCrLf & vbCrLf &
                "If yes, the OCR prompt will explicitly request Markdown output and the saved file extension will be changed to .md.",
                "Yes, save OCR results as Markdown",
                "No, keep plain text")
            If markdownChoice = 0 Then
                Return ' User aborted
            End If

            useMarkdownOutputForFlattenedOcr = (markdownChoice = 1)

            If useMarkdownOutputForFlattenedOcr Then
                ocrMarkdownInstruction = Add_OcrMarkdownInstruction
            End If
        End If

        ' Create output subdirectory if needed
        Dim outputBaseDir As String = ""
        If useSubdirectory Then
            If isDirectory Then
                outputBaseDir = IO.Path.Combine(selectedPath, outputSubdirName)
            Else
                outputBaseDir = IO.Path.Combine(IO.Path.GetDirectoryName(selectedPath), outputSubdirName)
            End If

            Try
                If Not IO.Directory.Exists(outputBaseDir) Then
                    IO.Directory.CreateDirectory(outputBaseDir)
                End If
            Catch ex As Exception
                ShowCustomMessageBox($"Failed to create output directory: {ex.Message}")
                Return
            End Try
        End If

        ' Show progress bar
        ProgressBarModule.GlobalProgressValue = 0
        ProgressBarModule.GlobalProgressMax = filesToProcess.Count
        ProgressBarModule.GlobalProgressLabel = "Initializing..."
        ProgressBarModule.CancelOperation = False
        ProgressBarModule.ShowProgressBarInSeparateThread(AN & " Convert to Text", "Starting text extraction...")

        ' Process files
        Dim successCount As Integer = 0
        Dim failedFiles As New List(Of String)()
        Dim emptyContentFiles As New List(Of String)()
        Dim flattenedPdfCount As Integer = 0

        Try
            For i As Integer = 0 To filesToProcess.Count - 1
                ' Check for cancellation
                If ProgressBarModule.CancelOperation Then
                    Exit For
                End If

                Dim filePath As String = filesToProcess(i)
                Dim fileName As String = IO.Path.GetFileName(filePath)

                ' Update progress
                ProgressBarModule.GlobalProgressValue = i
                ProgressBarModule.GlobalProgressLabel = $"Processing file {i + 1} of {filesToProcess.Count}: {fileName}"

                Try
                    ' Determine OCR settings for this file
                    Dim isPdf As Boolean = IO.Path.GetExtension(filePath).Equals(".pdf", StringComparison.OrdinalIgnoreCase)
                    Dim useOcrForThisFile As Boolean = isPdf AndAlso doOcr
                    Dim useMarkdownForThisFile As Boolean =
                        isPdf AndAlso useOcrForThisFile AndAlso flattenBeforeOcr AndAlso useMarkdownOutputForFlattenedOcr
                    Dim outputExtension As String = If(useMarkdownForThisFile, ".md", ".txt")

                    ' If flattening is requested for PDFs, flatten to temp file first
                    Dim effectiveFilePath As String = filePath
                    Dim tempFlattenedPath As String = Nothing

                    If isPdf AndAlso flattenBeforeOcr AndAlso useOcrForThisFile Then
                        Try
                            ProgressBarModule.GlobalProgressLabel = $"Flattening PDF {i + 1} of {filesToProcess.Count}: {fileName}..."
                            tempFlattenedPath = IO.Path.Combine(IO.Path.GetTempPath(), $"{AN2}_flatten_{Guid.NewGuid():N}.pdf")
                            Await System.Threading.Tasks.Task.Run(Sub() FlattenPdfToImageOnly(filePath, tempFlattenedPath, 200))
                            effectiveFilePath = tempFlattenedPath
                            flattenedPdfCount += 1
                            ProgressBarModule.GlobalProgressLabel = $"OCR processing file {i + 1} of {filesToProcess.Count}: {fileName}..."
                        Catch ex As Exception
                            ' Flattening failed — fall back to original PDF
                            effectiveFilePath = filePath
                            If tempFlattenedPath IsNot Nothing Then
                                Try : If IO.File.Exists(tempFlattenedPath) Then IO.File.Delete(tempFlattenedPath)
                                Catch : End Try
                            End If
                            tempFlattenedPath = Nothing
                        End Try
                    End If

                    Try
                        ' Read file content — askUser=False ensures Hidesplash=True in PerformOCR/ReadBinaryFileViaLLM,
                        ' preventing the countdown splash from interrupting batch processing.
                        ' The user has already made all choices upfront (OCR on/off, flatten on/off, Markdown on/off).
                        Dim content As String = Await GetFileContent(
                            effectiveFilePath,
                            Silent:=True,
                            DoOCR:=useOcrForThisFile,
                            AskUser:=False,
                            OcrAdditionalInstruction:=If(useMarkdownForThisFile, ocrMarkdownInstruction, ""), ShowOCRProgress:=True)

                        If String.IsNullOrWhiteSpace(content) Then
                            emptyContentFiles.Add($"{fileName} ({IO.Path.GetExtension(filePath).ToLowerInvariant()})")
                            Continue For
                        End If

                        If content.StartsWith("Error", StringComparison.OrdinalIgnoreCase) AndAlso content.Length < 200 Then
                            failedFiles.Add($"{fileName}: {content}")
                            Continue For
                        End If

                        ' Determine output path
                        Dim outputPath As String
                        If useSubdirectory Then
                            ' Preserve relative directory structure if processing subdirectories
                            If isDirectory Then
                                Dim relativePath As String = filePath.Substring(selectedPath.Length).TrimStart(IO.Path.DirectorySeparatorChar)
                                Dim relativeDir As String = IO.Path.GetDirectoryName(relativePath)
                                If Not String.IsNullOrEmpty(relativeDir) Then
                                    Dim targetDir As String = IO.Path.Combine(outputBaseDir, relativeDir)
                                    If Not IO.Directory.Exists(targetDir) Then
                                        IO.Directory.CreateDirectory(targetDir)
                                    End If
                                End If
                                outputPath = IO.Path.Combine(outputBaseDir, relativePath & outputExtension)
                            Else
                                outputPath = IO.Path.Combine(outputBaseDir, fileName & outputExtension)
                            End If
                        Else
                            outputPath = filePath & outputExtension
                        End If

                        IO.File.WriteAllText(outputPath, content, System.Text.Encoding.UTF8)
                        successCount += 1

                    Finally
                        ' Clean up temp flattened PDF
                        If tempFlattenedPath IsNot Nothing Then
                            Try : If IO.File.Exists(tempFlattenedPath) Then IO.File.Delete(tempFlattenedPath)
                            Catch : End Try
                        End If
                    End Try

                Catch ex As Exception
                    failedFiles.Add($"{fileName}: {ex.Message}")
                End Try
            Next

        Finally
            ' Close progress bar
            ProgressBarModule.CancelOperation = True
        End Try

        ' Build summary
        Dim wasCancelled As Boolean = (successCount + failedFiles.Count + emptyContentFiles.Count) < filesToProcess.Count

        Dim summary As New System.Text.StringBuilder()

        If wasCancelled Then
            summary.AppendLine("Operation was cancelled by user.")
            summary.AppendLine()
        End If

        summary.AppendLine($"Successfully converted: {successCount} file(s)")

        If useSubdirectory Then
            summary.AppendLine($"Output directory: {outputBaseDir}")
        End If

        If doOcr Then
            summary.AppendLine($"OCR was enabled for PDF files.")
            If flattenBeforeOcr Then
                summary.AppendLine($"PDFs were flattened to images before OCR ({flattenedPdfCount} file(s)).")
                If useMarkdownOutputForFlattenedOcr Then
                    summary.AppendLine("Flattened PDF OCR results were saved as Markdown (.md).")
                End If
            End If
        End If

        If emptyContentFiles.Count > 0 Then
            summary.AppendLine()
            summary.AppendLine($"⚠ Files with no extractable content ({emptyContentFiles.Count} file(s)):")
            For Each f In emptyContentFiles.Take(10)
                summary.AppendLine($"  • {f}")
            Next
            If emptyContentFiles.Count > 10 Then
                summary.AppendLine($"  ... and {emptyContentFiles.Count - 10} more")
            End If
            summary.AppendLine("  (These files were read but returned no text content)")
        End If

        If unsupportedFiles.Count > 0 Then
            summary.AppendLine()
            summary.AppendLine($"Skipped due to unsupported file type ({unsupportedFiles.Count} file(s)):")
            ' Group by extension for readability
            Dim byExtension = unsupportedFiles.
                GroupBy(Function(f) IO.Path.GetExtension(f).ToLowerInvariant()).
                OrderByDescending(Function(g) g.Count())
            For Each extGroup In byExtension
                Dim extLabel As String = If(String.IsNullOrEmpty(extGroup.Key), "(no extension)", extGroup.Key)
                If extGroup.Count() <= 3 Then
                    For Each f In extGroup
                        summary.AppendLine($"  • {IO.Path.GetFileName(f)}")
                    Next
                Else
                    summary.AppendLine($"  • {extGroup.Count()} {extLabel} files (e.g., {IO.Path.GetFileName(extGroup.First())})")
                End If
            Next
        End If

        If legacyDocFilesSkipped.Count > 0 Then
            summary.AppendLine()
            summary.AppendLine($"Skipped .doc files ({legacyDocFilesSkipped.Count} — disabled for security):")
            For Each f In legacyDocFilesSkipped.Take(5)
                summary.AppendLine($"  • {IO.Path.GetFileName(f)}")
            Next
            If legacyDocFilesSkipped.Count > 5 Then
                summary.AppendLine($"  ... and {legacyDocFilesSkipped.Count - 5} more")
            End If
        End If

        ' Build detailed failure/skip log for clipboard
        Dim clipboardLog As New System.Text.StringBuilder()

        If failedFiles.Count > 0 Then
            summary.AppendLine()
            summary.AppendLine($"Failed: {failedFiles.Count} file(s)")
            For Each f In failedFiles.Take(10)
                summary.AppendLine($"  • {f}")
            Next
            If failedFiles.Count > 10 Then
                summary.AppendLine($"  ... and {failedFiles.Count - 10} more")
            End If

            clipboardLog.AppendLine("=== FAILED FILES ===")
            For Each f In failedFiles
                clipboardLog.AppendLine(f)
            Next
            clipboardLog.AppendLine()
        End If

        If emptyContentFiles.Count > 0 Then
            clipboardLog.AppendLine("=== EMPTY CONTENT FILES ===")
            For Each f In emptyContentFiles
                clipboardLog.AppendLine(f)
            Next
            clipboardLog.AppendLine()
        End If

        If unsupportedFiles.Count > 0 Then
            clipboardLog.AppendLine("=== UNSUPPORTED FILES (Ignored) ===")
            For Each f In unsupportedFiles
                clipboardLog.AppendLine(IO.Path.GetFileName(f))
            Next
            clipboardLog.AppendLine()
        End If

        If legacyDocFilesSkipped.Count > 0 Then
            clipboardLog.AppendLine("=== LEGACY .DOC FILES (Disabled for security) ===")
            For Each f In legacyDocFilesSkipped
                clipboardLog.AppendLine(IO.Path.GetFileName(f))
            Next
            clipboardLog.AppendLine()
        End If

        ' Copy log to clipboard if there were any issues
        If failedFiles.Count > 0 OrElse emptyContentFiles.Count > 0 OrElse unsupportedFiles.Count > 0 OrElse legacyDocFilesSkipped.Count > 0 Then
            Dim logText As String = clipboardLog.ToString().TrimEnd()
            SharedMethods.PutInClipboard(logText)
            summary.AppendLine()
            summary.AppendLine("(Detailed log copied to clipboard)")
        End If

        ShowCustomMessageBox(summary.ToString().TrimEnd(), AN & " Convert to Text")
    End Sub


    ''' <summary>
    ''' Extracts text from a comparison document with revisions marked using &lt;ins&gt; and &lt;del&gt; tags,
    ''' comments marked with &lt;comment&gt; tags, and footnotes/endnotes included.
    ''' </summary>
    ''' <param name="compareDoc">The comparison document produced by Word.CompareDocuments.</param>
    ''' <returns>XML-tagged string with revisions and comments; empty string if compareDoc is Nothing.</returns>
    ''' <remarks>
    ''' Iterates paragraphs and extracts revision-marked text. Unchanged paragraphs are included as plain text.
    ''' Comments, footnotes, and endnotes are appended in separate XML-like sections.
    ''' </remarks>
    Private Shared Function ExtractChangesWithMarkupTags(compareDoc As Microsoft.Office.Interop.Word.Document) As String
        If compareDoc Is Nothing Then Return String.Empty

        Dim sb As New System.Text.StringBuilder()

        Try
            ' Build text with revision markup
            For Each para As Microsoft.Office.Interop.Word.Paragraph In compareDoc.Paragraphs
                Dim paraText As New System.Text.StringBuilder()
                Dim rng As Microsoft.Office.Interop.Word.Range = para.Range

                ' Process revisions in this paragraph
                For Each rev As Microsoft.Office.Interop.Word.Revision In rng.Revisions
                    Try
                        Dim revText As String = If(rev.Range.Text, String.Empty)
                        If String.IsNullOrEmpty(revText) Then Continue For

                        Select Case rev.Type
                            Case WdRevisionType.wdRevisionInsert
                                paraText.Append($"<ins>{revText}</ins>")
                            Case WdRevisionType.wdRevisionDelete
                                paraText.Append($"<del>{revText}</del>")
                            Case WdRevisionType.wdRevisionMovedFrom
                                paraText.Append($"<del>[moved from:]{revText}</del>")
                            Case WdRevisionType.wdRevisionMovedTo
                                paraText.Append($"<ins>[moved to:]{revText}</ins>")
                            Case Else
                                paraText.Append($"<ins>[{rev.Type}:]{revText}</ins>")
                        End Select
                    Catch
                    End Try
                Next

                ' If no revisions in this paragraph, add the plain text
                If paraText.Length = 0 Then
                    Try
                        paraText.Append(If(rng.Text, String.Empty))
                    Catch
                    End Try
                End If

                sb.AppendLine(paraText.ToString())
            Next

            ' Process comments
            If compareDoc.Comments IsNot Nothing AndAlso compareDoc.Comments.Count > 0 Then
                sb.AppendLine()
                sb.AppendLine("<comments>")
                For Each cmt As Microsoft.Office.Interop.Word.Comment In compareDoc.Comments
                    Try
                        Dim author As String = If(cmt.Author, "Unknown")
                        Dim commentText As String = If(cmt.Range.Text, String.Empty)
                        Dim scopeText As String = String.Empty
                        Try
                            scopeText = If(cmt.Scope.Text, String.Empty)
                        Catch
                        End Try

                        sb.AppendLine($"<comment author=""{System.Security.SecurityElement.Escape(author)}"" scope=""{System.Security.SecurityElement.Escape(scopeText)}"">{System.Security.SecurityElement.Escape(commentText)}</comment>")
                    Catch
                    End Try
                Next
                sb.AppendLine("</comments>")
            End If

            ' Process footnotes
            If compareDoc.Footnotes IsNot Nothing AndAlso compareDoc.Footnotes.Count > 0 Then
                sb.AppendLine()
                sb.AppendLine("<footnotes>")
                For Each fn As Microsoft.Office.Interop.Word.Footnote In compareDoc.Footnotes
                    Try
                        Dim fnText As String = If(fn.Range.Text, String.Empty)
                        sb.AppendLine($"<footnote index=""{fn.Index}"">{System.Security.SecurityElement.Escape(fnText)}</footnote>")
                    Catch
                    End Try
                Next
                sb.AppendLine("</footnotes>")
            End If

            ' Process endnotes
            If compareDoc.Endnotes IsNot Nothing AndAlso compareDoc.Endnotes.Count > 0 Then
                sb.AppendLine()
                sb.AppendLine("<endnotes>")
                For Each en As Microsoft.Office.Interop.Word.Endnote In compareDoc.Endnotes
                    Try
                        Dim enText As String = If(en.Range.Text, String.Empty)
                        sb.AppendLine($"<endnote index=""{en.Index}"">{System.Security.SecurityElement.Escape(enText)}</endnote>")
                    Catch
                    End Try
                Next
                sb.AppendLine("</endnotes>")
            End If

        Catch ex As System.Exception
            sb.AppendLine($"[Error extracting changes: {ex.Message}]")
        End Try

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' Submits extracted changes to the LLM for summarization and displays the result in a new HTML window.
    ''' Runs on a separate STA thread to avoid blocking the UI.
    ''' </summary>
    ''' <param name="extractedChangesText">XML-tagged text containing revisions and comments extracted from comparison document.</param>
    ''' <remarks>
    ''' Uses SP_Markup system prompt. LLM result (Markdown format) is converted to HTML via Markdig.
    ''' Thread is configured as STA (Single-Threaded Apartment) for ShowHTMLCustomMessageBox compatibility.
    ''' </remarks>
    Private Shared Sub SummarizeComparisonChangesAsync(extractedChangesText As String)
        ' Run the LLM call and display on a new thread to avoid blocking
        Dim t As New Threading.Thread(
            Sub()
                Try
                    ' Build the prompt
                    Dim userPrompt As String = "<TEXTTOPROCESS>" & vbCrLf & extractedChangesText & vbCrLf & "</TEXTTOPROCESS>"

                    ' System prompt for change analysis
                    Dim systemPrompt As String = SP_Markup

                    Dim llmResult As String = String.Empty
                    Try
                        llmResult = LLM(
                            _context,
                            systemPrompt,
                            userPrompt,
                            "",
                            "",
                            0,
                            False,
                            False).GetAwaiter().GetResult()
                    Catch ex As System.Exception
                        llmResult = $"Error calling LLM: {ex.Message}"
                    End Try

                    ' Convert Markdown to HTML using Markdig
                    Dim htmlResult As String
                    Try
                        Dim pipeline = New Markdig.MarkdownPipelineBuilder().UseAdvancedExtensions().Build()
                        Dim bodyHtml As String = Markdig.Markdown.ToHtml(If(llmResult, String.Empty), pipeline)

                        htmlResult = "<!DOCTYPE html><html><head><meta charset=""utf-8"">" &
                                     SummaryHtmlStyle &
                                     "</head><body>" &
                                     bodyHtml &
                                     "</body></html>"
                    Catch ex As System.Exception
                        htmlResult = $"<html><body><pre>{System.Security.SecurityElement.Escape(If(llmResult, ex.Message))}</pre></body></html>"
                    End Try

                    ShowHTMLCustomMessageBox(htmlResult, $"{AN} Change Summary")

                Catch ex As System.Exception
                    ShowCustomMessageBox($"Failed to summarize changes: {ex.Message}", AN)
                End Try
            End Sub)
        t.SetApartmentState(Threading.ApartmentState.STA)
        t.IsBackground = True
        t.Start()
    End Sub





    ''' <summary>
    ''' Extracts revisions and comments from the active document or selection based on a date filter,
    ''' then summarizes them using the LLM and displays the result.
    ''' </summary>
    Public Shared Async Sub SummarizeDocumentChanges()
        Try
            Dim app As Microsoft.Office.Interop.Word.Application = Nothing
            Dim doc As Microsoft.Office.Interop.Word.Document = Nothing

            Try
                app = Globals.ThisAddIn.Application
                doc = app.ActiveDocument
            Catch
                ShowCustomMessageBox("No active document found.", AN)
                Exit Sub
            End Try

            If doc Is Nothing Then
                ShowCustomMessageBox("No active document found.", AN)
                Exit Sub
            End If

            ' Determine scope: selection or entire document
            Dim sel As Microsoft.Office.Interop.Word.Selection = app.Selection
            Dim useEntireDoc As Boolean = (sel Is Nothing OrElse sel.Range Is Nothing OrElse sel.Start = sel.End)
            Dim scopeRange As Microsoft.Office.Interop.Word.Range = If(useEntireDoc, doc.Content, sel.Range)
            Dim scopeDescription As String = If(useEntireDoc, "the entire document", "the selected text")

            ' Prompt for date filter
            Dim defaultDate As String = System.DateTime.Now.AddDays(-7).ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture)
            Dim userDateInput As String = ShowCustomInputBox(
                $"Enter the earliest date for changes to include (leave empty to include all tracked changes and only comments made not older than 60 minutes before the first change).{vbCrLf}{vbCrLf}Changes from {scopeDescription} will be analyzed.",
                $"{AN} Summarize Changes",
                True,
                defaultDate)

            If userDateInput Is Nothing Then
                ' User cancelled
                Exit Sub
            End If

            userDateInput = userDateInput.Trim()

            Dim filterDate As System.DateTime? = Nothing
            Dim filterByDate As Boolean = False

            If Not String.IsNullOrEmpty(userDateInput) Then
                Dim parsed As System.DateTime
                If System.DateTime.TryParse(userDateInput, System.Globalization.CultureInfo.CurrentCulture, System.Globalization.DateTimeStyles.None, parsed) Then
                    filterDate = parsed.Date
                    filterByDate = True
                Else
                    ShowCustomMessageBox("Invalid date format. Operation aborted.", AN)
                    Exit Sub
                End If
            End If

            ' Extract revisions and comments
            Dim extractedText As String = ExtractRevisionsAndCommentsWithMarkup(doc, scopeRange, filterDate, filterByDate)

            If String.IsNullOrWhiteSpace(extractedText) Then
                Dim dateInfo As String = If(filterByDate, $" on or after {filterDate.Value:yyyy-MM-dd}", "")
                ShowCustomMessageBox($"No revisions or comments found{dateInfo} in {scopeDescription}.", AN)
                Exit Sub
            End If

            ' Build the prompt
            Dim userPrompt As String = "<TEXTTOPROCESS>" & vbCrLf & extractedText & vbCrLf & "</TEXTTOPROCESS>"

            ' System prompt for change analysis
            Dim systemPrompt As String = SP_Markup

            Dim llmResult As String = String.Empty
            Try
                llmResult = Await SharedMethods.LLM(
                    _context,
                    systemPrompt,
                    userPrompt,
                    "",
                    "",
                    0,
                    False,
                    False)
            Catch ex As System.Exception
                llmResult = $"Error calling LLM: {ex.Message}"
            End Try

            ' Convert Markdown to HTML using Markdig
            Dim htmlResult As String
            Try
                Dim pipeline = New Markdig.MarkdownPipelineBuilder().UseAdvancedExtensions().Build()
                Dim bodyHtml As String = Markdig.Markdown.ToHtml(If(llmResult, String.Empty), pipeline)

                Dim dateFilterInfo As String = If(filterByDate,
                    $"<p style='color:#666; font-size:9pt;'>Covering changes/comments from {filterDate.Value:yyyy-MM-dd} onwards in {scopeDescription}</p>",
                    $"<p style='color:#666; font-size:9pt;'>Covering all tracked changes in {scopeDescription} (and comments not older than 60 minutes before the first change)</p>")

                htmlResult = "<!DOCTYPE html><html><head><meta charset=""utf-8"">" &
                             SummaryHtmlStyle &
                             "</head><body>" &
                             dateFilterInfo &
                             bodyHtml &
                             "</body></html>"
            Catch ex As System.Exception
                htmlResult = $"<html><body><pre>{System.Security.SecurityElement.Escape(If(llmResult, ex.Message))}</pre></body></html>"
            End Try

            ' Show the result
            ShowHTMLCustomMessageBox(htmlResult, $"{AN} Change Summary")

        Catch ex As System.Exception
            ShowCustomMessageBox($"Failed to summarize changes: {ex.Message}", AN)
        End Try
    End Sub

    ''' <summary>
    ''' Extracts revisions and comments from a document range with markup tags.
    ''' Uses the same format as ExtractChangesWithMarkupTags for LLM compatibility.
    ''' Ignores pure formatting revisions for output but uses them for comment date calculation.
    ''' If filterByDate is True, includes revisions and comments on or after filterDate.
    ''' If filterByDate is False, includes all substantive revisions and comments added since first revision minus 60 minutes.
    ''' </summary>
    Private Shared Function ExtractRevisionsAndCommentsWithMarkup(
        doc As Microsoft.Office.Interop.Word.Document,
        scopeRange As Microsoft.Office.Interop.Word.Range,
        filterDate As System.DateTime?,
        filterByDate As Boolean) As String

        If doc Is Nothing OrElse scopeRange Is Nothing Then Return String.Empty

        Dim sb As New System.Text.StringBuilder()
        Dim hasContent As Boolean = False

        ' Revision types that are pure formatting (to be ignored in output, but used for date calculation)
        Dim formattingTypes As New HashSet(Of Integer)({
            CInt(WdRevisionType.wdRevisionProperty),
            CInt(WdRevisionType.wdRevisionParagraphNumber),
            CInt(WdRevisionType.wdRevisionParagraphProperty),
            CInt(WdRevisionType.wdRevisionSectionProperty),
            CInt(WdRevisionType.wdRevisionStyle),
            CInt(WdRevisionType.wdRevisionStyleDefinition),
            CInt(WdRevisionType.wdRevisionTableProperty)
        })

        Try
            ' Find the earliest revision date (including formatting revisions) for comment filtering
            Dim earliestRevisionDate As System.DateTime? = Nothing
            For Each rev As Microsoft.Office.Interop.Word.Revision In scopeRange.Revisions
                Try
                    If Not earliestRevisionDate.HasValue OrElse rev.Date < earliestRevisionDate.Value Then
                        earliestRevisionDate = rev.Date
                    End If
                Catch
                End Try
            Next

            ' Calculate comment cutoff date: earliest revision minus 60 minutes
            Dim commentCutoffDate As System.DateTime? = Nothing
            If earliestRevisionDate.HasValue Then
                commentCutoffDate = earliestRevisionDate.Value.AddMinutes(-60)
            End If

            ' Collect substantive revisions within the scope range
            Dim revisionList As New List(Of Microsoft.Office.Interop.Word.Revision)()

            For Each rev As Microsoft.Office.Interop.Word.Revision In scopeRange.Revisions
                Try
                    ' Skip pure formatting revisions for output
                    If formattingTypes.Contains(CInt(rev.Type)) Then
                        Continue For
                    End If

                    Dim includeRevision As Boolean = False

                    If filterByDate Then
                        ' Include if revision date >= filter date
                        If rev.Date >= filterDate.Value Then
                            includeRevision = True
                        End If
                    Else
                        ' No date filter: include all substantive revisions
                        includeRevision = True
                    End If

                    If includeRevision Then
                        revisionList.Add(rev)
                    End If
                Catch
                End Try
            Next

            ' Sort revisions by position in document
            revisionList.Sort(Function(a, b)
                                  Try
                                      Return a.Range.Start.CompareTo(b.Range.Start)
                                  Catch
                                      Return 0
                                  End Try
                              End Function)

            ' Build revision output using same format as ExtractChangesWithMarkupTags
            For Each rev In revisionList
                Try
                    Dim revText As String = If(rev.Range.Text, String.Empty)
                    If String.IsNullOrEmpty(revText) Then Continue For

                    Select Case rev.Type
                        Case WdRevisionType.wdRevisionInsert
                            sb.AppendLine($"<ins>{revText}</ins>")
                            hasContent = True
                        Case WdRevisionType.wdRevisionDelete
                            sb.AppendLine($"<del>{revText}</del>")
                            hasContent = True
                        Case WdRevisionType.wdRevisionMovedFrom
                            sb.AppendLine($"<del>[moved from:]{revText}</del>")
                            hasContent = True
                        Case WdRevisionType.wdRevisionMovedTo
                            sb.AppendLine($"<ins>[moved to:]{revText}</ins>")
                            hasContent = True
                        Case Else
                            ' Other non-formatting revision types
                            sb.AppendLine($"<ins>[{rev.Type}:]{revText}</ins>")
                            hasContent = True
                    End Select
                Catch
                End Try
            Next

            ' Collect comments within the scope range
            Dim commentList As New List(Of Microsoft.Office.Interop.Word.Comment)()

            For Each cmt As Microsoft.Office.Interop.Word.Comment In doc.Comments
                Try
                    ' Check if comment scope overlaps with our range
                    Dim cmtStart As Integer = -1
                    Dim cmtEnd As Integer = -1
                    Try
                        cmtStart = cmt.Scope.Start
                        cmtEnd = cmt.Scope.End
                    Catch
                        Try
                            cmtStart = cmt.Reference.Start
                            cmtEnd = cmt.Reference.End
                        Catch
                            Continue For
                        End Try
                    End Try

                    ' Check if comment is within scope range
                    If cmtStart >= scopeRange.Start AndAlso cmtEnd <= scopeRange.End Then
                        Dim includeComment As Boolean = False

                        If filterByDate Then
                            ' Include if comment date >= filter date
                            If cmt.Date >= filterDate.Value Then
                                includeComment = True
                            End If
                        Else
                            ' No date filter: include comments added since first revision minus 60 minutes
                            If commentCutoffDate.HasValue Then
                                If cmt.Date >= commentCutoffDate.Value Then
                                    includeComment = True
                                End If
                            Else
                                ' No revisions found, so no comments to include
                                includeComment = False
                            End If
                        End If

                        If includeComment Then
                            commentList.Add(cmt)
                        End If
                    End If
                Catch
                End Try
            Next

            ' Sort comments by position
            commentList.Sort(Function(a, b)
                                 Try
                                     Return a.Scope.Start.CompareTo(b.Scope.Start)
                                 Catch
                                     Return 0
                                 End Try
                             End Function)

            ' Build comments output using same format as ExtractChangesWithMarkupTags
            If commentList.Count > 0 Then
                sb.AppendLine()
                sb.AppendLine("<comments>")
                For Each cmt In commentList
                    Try
                        Dim author As String = If(cmt.Author, "Unknown")
                        Dim commentText As String = If(cmt.Range.Text, String.Empty)
                        Dim scopeText As String = String.Empty
                        Try
                            scopeText = If(cmt.Scope.Text, String.Empty)
                        Catch
                        End Try

                        sb.AppendLine($"<comment author=""{System.Security.SecurityElement.Escape(author)}"" scope=""{System.Security.SecurityElement.Escape(scopeText)}"">{System.Security.SecurityElement.Escape(commentText)}</comment>")
                        hasContent = True
                    Catch
                    End Try
                Next
                sb.AppendLine("</comments>")
            End If

        Catch ex As System.Exception
            sb.AppendLine($"[Error extracting changes: {ex.Message}]")
        End Try

        Return If(hasContent, sb.ToString(), String.Empty)
    End Function

    ''' <summary>
    ''' Removes Word content controls from the current selection or entire document while preserving text and formatting.
    ''' Prompts user if no selection exists.
    ''' </summary>
    ''' <remarks>
    ''' If selection exists: Removes controls overlapping the selection.
    ''' If no selection: Prompts user to confirm document-wide removal.
    ''' Unlocks LockContentControl and LockContents properties before deletion.
    ''' Temporarily disables TrackRevisions during removal to avoid tracked changes.
    ''' </remarks>
    Public Sub RemoveContentControlsRespectSelection()
        Try
            Dim app As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application
            Dim sel As Microsoft.Office.Interop.Word.Selection = app.Selection

            Dim hasSelection As Boolean = (sel IsNot Nothing AndAlso sel.Range IsNot Nothing AndAlso sel.Range.Start <> sel.Range.End)
            Dim removedCount As Integer = 0

            If hasSelection Then
                removedCount = RemoveContentControlsInRangeKeepContents(sel.Range)
            Else
                Dim Answer As Integer = ShowCustomYesNoBox("No text selection detected. Do you want to remove ALL content controls in the entire document?", "Yes", "No, abort")
                If Answer = 1 Then
                    removedCount = RemoveAllContentControlsKeepContents(app)
                Else
                    ShowCustomMessageBox("Operation aborted.")
                    Exit Sub
                End If
            End If

            ShowCustomMessageBox("Successfully removed " & removedCount.ToString() & " content control(s). Text and formatting were preserved.")
        Catch ex As System.Exception
            ShowCustomMessageBox("Error while removing content controls: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Removes all content controls from the active document while preserving contents and formatting.
    ''' </summary>
    ''' <param name="app">The Word Application instance.</param>
    ''' <returns>The number of content controls successfully removed.</returns>
    ''' <exception cref="System.Exception">Thrown if document is protected.</exception>
    ''' <remarks>
    ''' Controls are sorted inner-first (by descending start position) to handle nested controls correctly.
    ''' TrackRevisions is temporarily disabled during removal.
    ''' </remarks>
    Public Function RemoveAllContentControlsKeepContents(app As Microsoft.Office.Interop.Word.Application) As System.Int32
        Dim doc As Microsoft.Office.Interop.Word.Document = app.ActiveDocument
        If doc Is Nothing Then Return 0

        If doc.ProtectionType <> Microsoft.Office.Interop.Word.WdProtectionType.wdNoProtection Then
            Throw New System.Exception("Document is protected; cannot remove content controls.")
        End If

        Dim beforeCount As Integer = doc.ContentControls.Count

        ' Snapshot and sort inner-first
        Dim list As New List(Of Microsoft.Office.Interop.Word.ContentControl)(beforeCount)
        For i As Integer = 1 To beforeCount
            list.Add(doc.ContentControls(i))
        Next
        list.Sort(Function(a, b) b.Range.Start.CompareTo(a.Range.Start))

        Dim trackWasOn As Boolean = doc.TrackRevisions
        doc.TrackRevisions = False
        Try
            For Each cc In list
                Try
                    If cc Is Nothing Then Continue For
                    If cc.LockContentControl Then cc.LockContentControl = False
                    If cc.LockContents Then cc.LockContents = False
                    cc.Delete(False) ' keep contents/formatting
                Catch
                    ' Continue with other controls
                End Try
            Next
        Finally
            doc.TrackRevisions = trackWasOn
        End Try

        Dim afterCount As Integer = doc.ContentControls.Count
        Return System.Math.Max(0, beforeCount - afterCount)
    End Function

    ''' <summary>
    ''' Removes content controls that overlap with the specified range while preserving contents and formatting.
    ''' </summary>
    ''' <param name="rng">The Word range defining the scope for content control removal.</param>
    ''' <returns>The number of content controls successfully removed.</returns>
    ''' <exception cref="System.Exception">Thrown if range is Nothing or document is protected.</exception>
    ''' <remarks>
    ''' Only processes controls in the same story type as the range.
    ''' Controls are sorted inner-first to handle nested controls correctly.
    ''' TrackRevisions is temporarily disabled during removal.
    ''' </remarks>
    Public Function RemoveContentControlsInRangeKeepContents(ByVal rng As Microsoft.Office.Interop.Word.Range) As System.Int32
        If rng Is Nothing Then Throw New System.Exception("Selection range is not available.")
        Dim doc As Microsoft.Office.Interop.Word.Document = rng.Document
        If doc Is Nothing Then Return 0

        If doc.ProtectionType <> Microsoft.Office.Interop.Word.WdProtectionType.wdNoProtection Then
            Throw New System.Exception("Document is protected; cannot remove content controls.")
        End If

        ' Collect all controls overlapping the selection, same story only
        Dim allCcs As Microsoft.Office.Interop.Word.ContentControls = doc.ContentControls
        Dim list As New List(Of Microsoft.Office.Interop.Word.ContentControl)
        For i As Integer = 1 To allCcs.Count
            Dim cc = allCcs(i)
            If cc.Range Is Nothing Then Continue For
            If cc.Range.StoryType <> rng.StoryType Then Continue For
            If cc.Range.Start < rng.End AndAlso cc.Range.End > rng.Start Then
                list.Add(cc)
            End If
        Next

        If list.Count = 0 Then Return 0

        ' Sort inner-first
        list.Sort(Function(a, b) b.Range.Start.CompareTo(a.Range.Start))

        Dim removed As Integer = 0
        Dim trackWasOn As Boolean = doc.TrackRevisions
        doc.TrackRevisions = False
        Try
            For Each cc In list
                Try
                    If cc Is Nothing Then Continue For
                    If cc.LockContentControl Then cc.LockContentControl = False
                    If cc.LockContents Then cc.LockContents = False
                    cc.Delete(False)
                    removed += 1
                Catch
                    ' Ignore and continue
                End Try
            Next
        Finally
            doc.TrackRevisions = trackWasOn
        End Try

        Return removed
    End Function


    ''' <summary>
    ''' Prompts user to select a file and inserts its content at the current cursor position.
    ''' For PDFs, if OCR is available, the user can choose a formatting-preserving import
    ''' that uses flattening + OCR + Markdown insertion.
    ''' </summary>
    Public Async Sub ImportTextFile()
        Dim app As Word.Application = Globals.ThisAddIn.Application
        Dim sel As Word.Range = app.Selection.Range
        Dim filePath As String = GetFileName()

        If String.IsNullOrWhiteSpace(filePath) Then
            Exit Sub
        End If

        Dim docText As String = String.Empty
        Dim ext As String = IO.Path.GetExtension(filePath).ToLowerInvariant()
        Dim preserveFormattingImport As Boolean = False
        Dim tempFlattenedPath As String = Nothing
        Dim splash As Slib.SplashScreen = Nothing

        Try
            If ext = ".pdf" AndAlso SharedMethods.IsOcrAvailable(_context) Then
                Dim preserveChoice As Integer = ShowCustomYesNoBox(
                    "Do you want to preserve the formatting/structure when importing this PDF?" & vbCrLf & vbCrLf &
                    "If yes, the PDF will be flattened, OCR will be performed, and the OCR result will be inserted as Markdown so that Word can render the formatting.",
                    "Yes, preserve formatting",
                    "No, import as plain text",
                    AN & " Import File")

                If preserveChoice = 0 Then
                    Exit Sub
                End If

                preserveFormattingImport = (preserveChoice = 1)
            End If

            If preserveFormattingImport Then
                Dim effectiveFilePath As String = filePath

                splash = New Slib.SplashScreen("Preparing PDF for OCR ...")
                splash.Show()
                splash.Refresh()
                System.Windows.Forms.Application.DoEvents()

                Try
                    splash.UpdateMessage("Flattening PDF for OCR ...")
                    tempFlattenedPath = IO.Path.Combine(IO.Path.GetTempPath(), $"{AN2}_import_flatten_{Guid.NewGuid():N}.pdf")
                    Await System.Threading.Tasks.Task.Run(Sub() FlattenPdfToImageOnly(filePath, tempFlattenedPath, 200))
                    effectiveFilePath = tempFlattenedPath
                Catch
                    effectiveFilePath = filePath
                    If tempFlattenedPath IsNot Nothing Then
                        Try
                            If IO.File.Exists(tempFlattenedPath) Then
                                IO.File.Delete(tempFlattenedPath)
                            End If
                        Catch
                        End Try
                    End If
                    tempFlattenedPath = Nothing
                End Try

                splash.UpdateMessage("AI is extracting the PDF content via OCR ...")
                System.Windows.Forms.Application.DoEvents()

                docText = Await GetFileContent(
                    effectiveFilePath,
                    False,
                    True,
                    False,
                    False,
                    OcrAdditionalInstruction:=Add_OcrMarkdownInstruction,
                    ShowOCRProgress:=True)

                If String.IsNullOrWhiteSpace(docText) Then
                    Exit Sub
                End If

                splash.UpdateMessage("Inserting OCR result into Word ...")
                System.Windows.Forms.Application.DoEvents()
            Else
                docText = Await GetFileContent(filePath, False, Not String.IsNullOrWhiteSpace(INI_APICall_Object))

                If String.IsNullOrWhiteSpace(docText) Then
                    Exit Sub
                End If
            End If

            sel.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
            sel.Select()

            If preserveFormattingImport Then
                InsertTextWithMarkdown(app.Selection, docText, False)

                Dim pattern As String = "\{\{(WFLD|WENT|WFNT):.*?\}\}"
                If Regex.IsMatch(docText, pattern) Then
                    Dim rng As Word.Range = app.Selection.Range
                    RestoreSpecialTextElements(rng)
                    rng.Document.Fields.Update()
                End If
            Else
                sel.Text = docText
                sel.Select()
            End If

        Finally
            If splash IsNot Nothing Then
                Try
                    splash.Close()
                    splash.Dispose()
                Catch
                End Try
            End If

            If tempFlattenedPath IsNot Nothing Then
                Try
                    If IO.File.Exists(tempFlattenedPath) Then
                        IO.File.Delete(tempFlattenedPath)
                    End If
                Catch
                End Try
            End If
        End Try
    End Sub

    ''' <summary>
    ''' Accepts formatting-only revisions (style, paragraph properties, etc.) in the current selection or document
    ''' while leaving text insertion/deletion/move revisions unchanged. Displays progress splash screen with ESC cancellation.
    ''' </summary>
    ''' <remarks>
    ''' Formatting revision types accepted: wdRevisionProperty, wdRevisionParagraphNumber, wdRevisionParagraphProperty,
    ''' wdRevisionSectionProperty, wdRevisionStyle, wdRevisionStyleDefinition, wdRevisionTableProperty.
    ''' Structural revisions (insert/delete/move) may contain embedded formatting that cannot be separated.
    ''' User can abort by pressing ESC key during processing.
    ''' </remarks>
    Public Sub AcceptFormatting()

        Dim sel As Word.Range = Globals.ThisAddIn.Application.Selection.Range
        Dim formatChangeCount As Integer = 0
        Dim docRef As String = "in the selected text"

        ' Ensure a selection is made (use content if selection empty)
        If sel Is Nothing OrElse sel.Start = sel.End Then
            sel = Globals.ThisAddIn.Application.ActiveDocument.Content
            docRef = "in the document"
        End If

        ' Quick exit if no revisions at all
        If sel.Revisions.Count = 0 Then
            ShowCustomMessageBox($"No revisions found {docRef}. Note: Formatting embedded in insert/delete revisions would also count as those insert/delete changes.")
            Return
        End If

        Dim splash As New Slib.SplashScreen("Accepting formatting-only revisions... press 'Esc' to abort")
        splash.Show()
        splash.Refresh()

        ' Revision types treated as pure formatting (will be accepted)
        Dim formattingTypes As Word.WdRevisionType() = {
            Word.WdRevisionType.wdRevisionProperty,
            Word.WdRevisionType.wdRevisionParagraphNumber,
            Word.WdRevisionType.wdRevisionParagraphProperty,
            Word.WdRevisionType.wdRevisionSectionProperty,
            Word.WdRevisionType.wdRevisionStyle,
            Word.WdRevisionType.wdRevisionStyleDefinition,
            Word.WdRevisionType.wdRevisionTableProperty
        }

        Dim formattingSet As New HashSet(Of Integer)(formattingTypes.Select(Function(t) CInt(t)))

        ' Structural revision types that may carry embedded formatting
        Dim structuralTypes As Word.WdRevisionType() = {
            Word.WdRevisionType.wdRevisionInsert,
            Word.WdRevisionType.wdRevisionDelete,
            Word.WdRevisionType.wdRevisionMovedFrom,
            Word.WdRevisionType.wdRevisionMovedTo
        }
        Dim structuralSet As New HashSet(Of Integer)(structuralTypes.Select(Function(t) CInt(t)))

        Dim aborted As Boolean = False

        ' Accept formatting-only revisions
        For Each rev As Word.Revision In sel.Revisions
            System.Windows.Forms.Application.DoEvents()

            If (GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0 Then
                aborted = True
                Exit For
            End If
            If (GetAsyncKeyState(VK_ESCAPE) And 1) <> 0 Then
                aborted = True
                Exit For
            End If

            If formattingSet.Contains(CInt(rev.Type)) Then
                Try
                    rev.Accept()
                    formatChangeCount += 1
                Catch
                    ' Ignore failures; continue
                End Try
            End If
        Next

        splash.Close()

        ' Count remaining structural revisions (potentially with embedded formatting)
        Dim embeddedStructuralCount As Integer = sel.Revisions.Cast(Of Word.Revision)().Count(Function(r) structuralSet.Contains(CInt(r.Type)))

        ' Build final message
        Dim msg As New System.Text.StringBuilder
        If aborted Then
            msg.AppendLine("Operation aborted by user (Esc).")
            If formatChangeCount > 0 Then
                msg.AppendLine($"{formatChangeCount} formatting revision(s) were accepted before abort.")
            Else
                msg.AppendLine("No formatting revisions were accepted before abort.")
            End If
        Else
            If formatChangeCount > 0 Then
                msg.AppendLine($"{formatChangeCount} formatting revision(s) {docRef} (including paragraph numbering) have been accepted.")
            Else
                msg.AppendLine($"No pure formatting revisions were found {docRef}.")
            End If
        End If

        ' Always inform about possible embedded formatting
        If embeddedStructuralCount > 0 Then
            msg.AppendLine()
            msg.AppendLine($"{embeddedStructuralCount} insertion/deletion/move revision(s) remain. Some formatting applied during those changes cannot be accepted separately.")
        Else
            msg.AppendLine()
            msg.AppendLine("Note: If formatting was applied during text insertions/deletions/moves, it is part of those tracked text changes and cannot be accepted without accepting the text change itself.")
        End If

        ShowCustomMessageBox(msg.ToString())
    End Sub

    ''' <summary>Cached regex pattern from previous RegexSearchReplace invocation.</summary>
    Private Shared LastRegexPattern As String = String.Empty

    ''' <summary>Cached regex options from previous RegexSearchReplace invocation.</summary>
    Private Shared LastRegexOptions As String = String.Empty

    ''' <summary>Cached replacement text from previous RegexSearchReplace invocation.</summary>
    Private Shared LastRegexReplace As String = String.Empty

    ''' <summary>
    ''' Performs multi-pattern regex search and replace operations on the current selection or entire document.
    ''' Supports persistent pattern memory and validation before execution.
    ''' </summary>
    ''' <remarks>
    ''' Workflow: Prompts for patterns (one per line) → prompts for options (i/m/s/c/r/e flags) →
    ''' prompts for replacements (one per line, matching pattern count) → validates all patterns →
    ''' performs replacements or highlights first match if no replacement provided.
    ''' Previous patterns/options/replacements are cached in static fields (LastRegexPattern, LastRegexOptions, LastRegexReplace).
    ''' Aborts without changes if pattern count mismatches replacement count or any pattern is invalid.
    ''' </remarks>
    Public Sub RegexSearchReplace()
        Dim sel As Word.Range = Globals.ThisAddIn.Application.Selection.Range
        Dim docRef As String = "in the selected text"

        ' Ensure a selection is made
        If sel Is Nothing OrElse String.IsNullOrWhiteSpace(sel.Text) Then
            Globals.ThisAddIn.Application.ActiveDocument.Content.Select()
            sel = Globals.ThisAddIn.Application.Selection.Range
            docRef = "in the document"
        End If

        ' Step 1: Get regex patterns
        Dim regexPattern As String = ShowCustomInputBox("Step 1: Enter your Regex pattern(s), one per line (more info about Regex: vischerlnk.com/regexinfo):", "Regex Search & Replace", False, LastRegexPattern)?.Trim()
        If String.IsNullOrEmpty(regexPattern) Then Return

        ' Step 2: Get regex options
        Dim optionsInput As String = ShowCustomInputBox("Enter regex option(s) (i for IgnoreCase, m for Multiline, s for Singleline, c for Compiled, r for RightToLeft, e for ExplicitCapture):", "Regex Search & Replace", True, LastRegexOptions)

        Dim regexOptions As RegexOptions = RegexOptions.None

        If Not String.IsNullOrEmpty(optionsInput) Then
            ' Add specific options based on user input
            If optionsInput.Contains("i") Then regexOptions = regexOptions Or RegexOptions.IgnoreCase
            If optionsInput.Contains("m") Then regexOptions = regexOptions Or RegexOptions.Multiline
            If optionsInput.Contains("s") Then regexOptions = regexOptions Or RegexOptions.Singleline
            If optionsInput.Contains("c") Then regexOptions = regexOptions Or RegexOptions.Compiled
            If optionsInput.Contains("r") Then regexOptions = regexOptions Or RegexOptions.RightToLeft
            If optionsInput.Contains("e") Then regexOptions = regexOptions Or RegexOptions.ExplicitCapture
        End If

        ' Step 3: Get replacement text
        Dim replacementText As String = ShowCustomInputBox("Step 2: Enter your replacement text(s), one on each line, matching to your pattern(s) (leave empty or cancel to only search for the first hit):", "Regex Search & Replace", False, LastRegexReplace)

        ' Update the last-used regex pattern and options
        LastRegexPattern = regexPattern
        LastRegexOptions = optionsInput
        LastRegexReplace = replacementText

        ' Split patterns and replacements into lines
        Dim patterns() As String = regexPattern.Split(New String() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)
        Dim replacements() As String = If(Not String.IsNullOrEmpty(replacementText), replacementText.Split(New String() {Environment.NewLine}, StringSplitOptions.None), Nothing)

        ' Check if patterns and replacements match
        If replacements IsNot Nothing AndAlso patterns.Length <> replacements.Length Then
            ShowCustomMessageBox("The number of regex patterns does not match the number of replacement lines. Aborting without any replacements done.")
            Return
        End If

        ' Validate all regex patterns first
        For Each pattern As String In patterns
            Try
                Dim regexTest As New Regex(pattern, regexOptions)
            Catch ex As ArgumentException
                ShowCustomMessageBox($"Your regex pattern '{pattern}' is invalid ({ex.Message}). Aborting without any replacements done.")
                Return
            End Try
        Next

        ' Perform replacements after validation
        Dim totalReplacements As Integer = 0

        For i As Integer = 0 To patterns.Length - 1
            Dim pattern As String = patterns(i)
            Dim replacement As String = If(replacements IsNot Nothing, replacements(i), Nothing)

            Dim regex As New Regex(pattern, regexOptions)

            If Not String.IsNullOrEmpty(replacement) Then
                ' Perform replacement
                Dim replacementCount As Integer = 0
                sel.Text = regex.Replace(sel.Text, Function(match)
                                                       replacementCount += 1
                                                       Return replacement
                                                   End Function)
                totalReplacements += replacementCount
            Else
                ' Perform search only
                Dim match As Match = regex.Match(sel.Text)
                If match.Success Then
                    ' Highlight the first match
                    sel.Start = sel.Start + match.Index
                    sel.End = sel.Start + match.Length
                    Globals.ThisAddIn.Application.Selection.Select()
                    Globals.ThisAddIn.Application.ActiveWindow.ScrollIntoView(sel, True)
                    Return
                Else
                    ShowCustomMessageBox($"No matches found for '{pattern}' {docRef}.")
                    Return
                End If
            End If
        Next

        If replacements IsNot Nothing Then
            ShowCustomMessageBox($"{totalReplacements} replacement(s) made {docRef}.")
        Else
            ShowCustomMessageBox("Search complete. No replacements were made.")
        End If
    End Sub

    ''' <summary>
    ''' Calculates the time span between the first and last revision or comment by a specified user
    ''' (or all users) in the current selection or document. Supports optional earliest date filtering.
    ''' </summary>
    ''' <remarks>
    ''' Workflow: Prompts for user name (empty = all users) → prompts for earliest date filter →
    ''' iterates revisions and comments in scope → tracks min/max timestamps → displays time span breakdown
    ''' (days/hours/minutes) and user list if processing all users.
    ''' If no selection: Operates on entire document.
    ''' If date filter provided: Only considers revisions/comments on or after that date.
    ''' </remarks>
    Public Sub CalculateUserMarkupTimeSpan()

        Try
            Dim app As Microsoft.Office.Interop.Word.Application = Nothing
            Dim doc As Microsoft.Office.Interop.Word.Document = Nothing

            Try
                app = Globals.ThisAddIn.Application
                doc = app.ActiveDocument
            Catch
            End Try

            If app Is Nothing OrElse doc Is Nothing Then
                ShowCustomMessageBox("No active document found.", AN)
                Exit Sub
            End If

            Dim firstTimestamp As Date
            Dim lastTimestamp As Date
            Dim found As Boolean
            Dim userNames As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            Dim docRef As String = "in the selected text"

            ' Initialize
            found = False
            firstTimestamp = #1/1/1900#
            lastTimestamp = #1/1/1900#

            ' Prompt for user name
            Dim userName As String = String.Empty
            Try
                userName = app.UserName
            Catch
            End Try

            Dim userInput As String = ShowCustomInputBox("Please enter the name of the user (leave empty for all users):", "Markup Time Span", True, userName)

            ' ShowCustomInputBox returns "" on Cancel in SimpleInput mode — treat as abort
            If userInput Is Nothing Then
                Exit Sub
            End If

            userInput = userInput.Trim()

            ' Prompt for earliest date
            Dim earliestDate As System.DateTime = System.DateTime.MinValue
            Dim earliestDateFiltered As Boolean = False

            Dim userDateInput As String = ShowCustomInputBox(
                    "Please enter the earliest date (and time, if you wish) to consider (leave empty for no filter):",
                    "Markup Time Span",
                    True,
                    System.DateTime.Now.AddDays(-2).ToString(System.Globalization.CultureInfo.CurrentCulture)
                )

            If userDateInput Is Nothing Then
                Exit Sub
            End If

            userDateInput = userDateInput.Trim()

            Dim parsed As System.DateTime
            If String.IsNullOrEmpty(userDateInput) Then
                earliestDateFiltered = False
            ElseIf System.DateTime.TryParse(
                      userDateInput,
                      System.Globalization.CultureInfo.CurrentCulture,
                      System.Globalization.DateTimeStyles.None,
                      parsed
                  ) Then
                earliestDate = parsed
                earliestDateFiltered = True
            Else
                ShowCustomMessageBox("Improper date/time format - will abort.")
                Exit Sub
            End If

            ' Check selection
            Dim selRange As Word.Range = Nothing
            Try
                Dim sel As Word.Selection = app.Selection
                If sel IsNot Nothing AndAlso sel.Range IsNot Nothing AndAlso sel.Range.Start <> sel.Range.End Then
                    selRange = sel.Range
                End If
            Catch
            End Try

            If selRange Is Nothing Then
                Try
                    doc.Content.Select()
                    selRange = app.Selection.Range
                Catch
                End Try
                docRef = "in the document"
            End If

            If selRange Is Nothing Then
                ShowCustomMessageBox("Unable to access the document content.", AN)
                Exit Sub
            End If

            ' Process revisions
            Try
                Dim docRevisions As Word.Revisions = selRange.Revisions
                If docRevisions IsNot Nothing Then
                    For Each rev As Word.Revision In docRevisions
                        Try
                            Dim revAuthor As String = Nothing
                            Dim revDate As Date

                            Try
                                revAuthor = rev.Author
                                revDate = rev.Date
                            Catch
                                Continue For
                            End Try

                            If String.IsNullOrEmpty(revAuthor) Then revAuthor = "(unknown)"

                            If (String.IsNullOrEmpty(userInput) OrElse revAuthor.Equals(userInput, StringComparison.OrdinalIgnoreCase)) _
                                   AndAlso (Not earliestDateFiltered OrElse revDate >= earliestDate) Then
                                ' Update timestamps
                                If Not found Then
                                    firstTimestamp = revDate
                                    lastTimestamp = revDate
                                    found = True
                                Else
                                    If revDate < firstTimestamp Then firstTimestamp = revDate
                                    If revDate > lastTimestamp Then lastTimestamp = revDate
                                End If
                                ' Collect user names if processing all
                                If Not userNames.ContainsKey(revAuthor) Then
                                    userNames(revAuthor) = revAuthor
                                End If
                            End If
                        Catch
                            ' Skip problematic revision
                        End Try
                    Next
                End If
            Catch
                ' Revisions collection inaccessible (e.g., protected document)
            End Try

            ' Process comments
            Try
                Dim comments As Word.Comments = selRange.Comments
                If comments IsNot Nothing Then
                    For Each comment As Word.Comment In comments
                        Try
                            Dim cmtAuthor As String = Nothing
                            Dim cmtDate As Date

                            Try
                                cmtAuthor = comment.Author
                                cmtDate = comment.Date
                            Catch
                                Continue For
                            End Try

                            If String.IsNullOrEmpty(cmtAuthor) Then cmtAuthor = "(unknown)"

                            If (String.IsNullOrEmpty(userInput) OrElse cmtAuthor.Equals(userInput, StringComparison.OrdinalIgnoreCase)) _
                                   AndAlso (Not earliestDateFiltered OrElse cmtDate >= earliestDate) Then

                                ' Update timestamps
                                If Not found Then
                                    firstTimestamp = cmtDate
                                    lastTimestamp = cmtDate
                                    found = True
                                Else
                                    If cmtDate < firstTimestamp Then firstTimestamp = cmtDate
                                    If cmtDate > lastTimestamp Then lastTimestamp = cmtDate
                                End If
                                ' Collect user names if processing all
                                If Not userNames.ContainsKey(cmtAuthor) Then
                                    userNames(cmtAuthor) = cmtAuthor
                                End If
                            End If
                        Catch
                            ' Skip problematic comment
                        End Try
                    Next
                End If
            Catch
                ' Comments collection inaccessible
            End Try

            ' Display results
            If found Then
                Dim timeDiff As Double = DateDiff(DateInterval.Minute, firstTimestamp, lastTimestamp)
                If timeDiff < 0 Then timeDiff = System.Math.Abs(timeDiff)

                Dim timeSpan As String = System.Math.Floor(timeDiff / 1440).ToString() & " days, " &
                       ((timeDiff Mod 1440) \ 60).ToString("00") & " hours, " &
                       (timeDiff Mod 60).ToString("00") & " minutes"

                ' Format timestamps without seconds
                Dim formattedFirstTimestamp As String = firstTimestamp.ToString("dd/MM/yyyy HH:mm")
                Dim formattedLastTimestamp As String = lastTimestamp.ToString("dd/MM/yyyy HH:mm")

                Dim outputUserNames As String
                If String.IsNullOrEmpty(userInput) Then
                    ' Display all users
                    outputUserNames = "Users involved:" & vbCrLf
                    For Each kvp In userNames
                        outputUserNames &= "- " & kvp.Value & vbCrLf
                    Next
                Else
                    outputUserNames = "User: " & userInput
                End If
                ShowCustomMessageBox(outputUserNames & vbCrLf &
                    If(earliestDateFiltered, "Earliest considered: " & earliestDate.ToString("dd/MM/yyyy HH:mm") & vbCrLf, "") &
                    "First markup/comment: " & formattedFirstTimestamp & vbCrLf &
                    "Last markup/comment: " & formattedLastTimestamp & vbCrLf &
                    "Time span: " & timeSpan)
            Else
                If String.IsNullOrEmpty(userInput) Then
                    ShowCustomMessageBox($"No markups or comments found {docRef}.")
                Else
                    ShowCustomMessageBox("No markups or comments found for user '" & userInput & $"' {docRef}.")
                End If
            End If

        Catch ex As System.Exception
            ShowCustomMessageBox("Error in CalculateUserMarkupTimeSpan: " & ex.Message, AN)
        End Try
    End Sub

    ''' <summary>
    ''' Splits the selected paragraphs into two halves and compares them using the configured comparison method.
    ''' Requires an even number of non-empty paragraphs in the selection.
    ''' </summary>
    ''' <remarks>
    ''' Process: Counts non-empty paragraphs → validates even count → splits at midpoint →
    ''' extracts text (excluding final paragraph marks) → calls CompareAndInsert or CompareAndInsertComparedoc
    ''' based on INI_MarkupMethodHelper setting (1 = CompareDoc method, other = direct insert method).
    ''' Empty paragraphs (length ≤ 1 after trim) are ignored in counting.
    ''' For markup method 2 (surgical), arguments are swapped: text2 becomes the "original" and text1
    ''' becomes the "revised", because secondRange already contains text2 and the surgical method
    ''' patches the target range in-place (it expects the range to contain the original text).
    ''' </remarks>
    Public Sub CompareSelectionHalves()

        Dim sel As Word.Range
        Dim nonEmptyParaCount As Long
        Dim halfParaCount As Long
        Dim firstRange As Word.Range
        Dim secondRange As Word.Range
        Dim paraIndices() As Long
        Dim i As Long, index As Long

        ' Get the selected text
        sel = Globals.ThisAddIn.Application.Selection.Range

        ' Count non-empty paragraphs and store their indices
        ReDim paraIndices(0 To sel.Paragraphs.Count - 1)
        index = 0
        For i = 1 To sel.Paragraphs.Count
            If Len(sel.Paragraphs(i).Range.Text.Trim()) > 1 Then ' Greater than 1 to account for paragraph mark
                index += 1
                paraIndices(index - 1) = i
            End If
        Next

        ' Update nonEmptyParaCount
        nonEmptyParaCount = index

        ' If number of non-empty paragraphs is uneven or zero, abort
        If nonEmptyParaCount Mod 2 <> 0 Or nonEmptyParaCount = 0 Then
            ShowCustomMessageBox("The number of non-empty paragraphs in the selection is uneven or zero. Please select an even number of non-empty paragraphs.")
            Return
        End If

        ' Determine the halfway point
        halfParaCount = nonEmptyParaCount \ 2

        ' Get the first half and second half ranges
        firstRange = sel.Paragraphs(paraIndices(0)).Range
        firstRange.End = sel.Paragraphs(paraIndices(halfParaCount - 1)).Range.End

        secondRange = sel.Paragraphs(paraIndices(halfParaCount)).Range
        secondRange.End = sel.Paragraphs(paraIndices(nonEmptyParaCount - 1)).Range.End

        ' Get text from the first and second range without the final paragraph marks
        Dim text1 As String = Left(firstRange.Text, Len(firstRange.Text) - 1)
        Dim text2 As String = Left(secondRange.Text, Len(secondRange.Text) - 1)

        If INI_MarkupMethodHelper = 1 Then
            CompareAndInsertComparedoc(text1, text2, secondRange)
        ElseIf INI_MarkupMethodHelper = 2 Then

            CompareAndInsertSurgical(text2, text1, secondRange)
        Else
            CompareAndInsert(text1, text2, secondRange, INI_MarkupMethodHelper = 3, "These are the differences of the second (set of) paragraph(s) of the text selected:", True)
        End If
    End Sub


    ''' <summary>
    ''' Removes leading "RI: " prefix from Word comments (including threaded replies) in the selection or entire document.
    ''' Optionally also removes comment timestamps via Open XML post-processing.
    ''' Uses a progress bar and supports cancellation.
    ''' </summary>
    Public Sub NewRemoveRIPrefixFromComments()

        Dim riPrefix As String = $"{AN5}: "

        Try
            Dim app As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application
            Dim doc As Microsoft.Office.Interop.Word.Document = Nothing

            Try
                doc = app.ActiveDocument
            Catch
            End Try

            If doc Is Nothing Then
                ShowCustomMessageBox("No active document found.", AN)
                Exit Sub
            End If

            Dim sel As Microsoft.Office.Interop.Word.Selection = app.Selection
            Dim hasSelection As Boolean = (sel IsNot Nothing AndAlso sel.Range IsNot Nothing AndAlso sel.Range.Start <> sel.Range.End)
            Dim scopeRange As Microsoft.Office.Interop.Word.Range = If(hasSelection, sel.Range, doc.Content)

            Dim candidates As New List(Of Microsoft.Office.Interop.Word.Comment)()

            ' Collect candidates that start with the prefix (base comments + replies)
            For Each c As Microsoft.Office.Interop.Word.Comment In doc.Comments
                Dim inScope As Boolean = False

                Try
                    Dim cStart As Integer = c.Scope.Start
                    Dim cEnd As Integer = c.Scope.End
                    inScope = (cStart >= scopeRange.Start AndAlso cEnd <= scopeRange.End)
                Catch
                    Try
                        Dim cStart As Integer = c.Reference.Start
                        Dim cEnd As Integer = c.Reference.End
                        inScope = (cStart >= scopeRange.Start AndAlso cEnd <= scopeRange.End)
                    Catch
                        inScope = False
                    End Try
                End Try

                If Not inScope Then Continue For

                Try
                    Dim text As String = If(c.Range.Text, String.Empty)
                    If text.StartsWith(riPrefix, StringComparison.Ordinal) Then
                        candidates.Add(c)
                    End If
                Catch
                End Try

                ' Threaded replies
                Try
                    If c.Replies IsNot Nothing AndAlso c.Replies.Count > 0 Then
                        For Each r As Microsoft.Office.Interop.Word.Comment In c.Replies
                            Try
                                Dim replyText As String = If(r.Range.Text, String.Empty)
                                If replyText.StartsWith(riPrefix, StringComparison.Ordinal) Then
                                    candidates.Add(r)
                                End If
                            Catch
                            End Try
                        Next
                    End If
                Catch
                End Try
            Next

            If candidates.Count = 0 Then
                ShowCustomMessageBox($"No '{AN5}:' prefixes found in comments for the current scope.", AN)
                Exit Sub
            End If

            ' Ask whether to also remove the date/time from the processed comments
            Dim alsoRemoveDateTime As Boolean = False
            Dim dateTimeAnswer As Integer = ShowCustomYesNoBox(
                $"Found {candidates.Count} comment(s) with '{AN5}:' prefix." & vbCrLf & vbCrLf &
                "Do you also want to remove the date/time from these comments?" & vbCrLf &
                "(The date/time will be removed entirely.)",
                "Yes, also remove date/time",
                "No, only remove prefix",
                AN & $" Remove {AN5} Prefix")
            If dateTimeAnswer = 0 Then
                Exit Sub ' User closed/cancelled
            End If
            alsoRemoveDateTime = (dateTimeAnswer = 1)

            ' Collect the 1-based Interop Comment.Index for each candidate BEFORE modifying text
            ' (Index is stable as long as we don't add/remove comments)
            Dim commentIndicesToProcess As New List(Of Integer)()
            If alsoRemoveDateTime Then
                For Each cmt As Microsoft.Office.Interop.Word.Comment In candidates
                    Try
                        commentIndicesToProcess.Add(cmt.Index)
                    Catch
                    End Try
                Next
            End If

            ProgressBarModule.GlobalProgressValue = 0
            ProgressBarModule.GlobalProgressMax = candidates.Count
            ProgressBarModule.GlobalProgressLabel = "Processing comments..."
            ProgressBarModule.CancelOperation = False
            ProgressBarModule.ShowProgressBarInSeparateThread(AN & $" Remove {AN5} Prefix", "Starting...")

            Dim removedCount As Integer = 0

            For i As Integer = 0 To candidates.Count - 1
                System.Windows.Forms.Application.DoEvents()

                If ProgressBarModule.CancelOperation Then Exit For
                If (GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0 Then
                    ProgressBarModule.CancelOperation = True
                    Exit For
                End If

                ProgressBarModule.GlobalProgressValue = i + 1
                ProgressBarModule.GlobalProgressLabel = $"Processing {i + 1} of {candidates.Count}..."

                Try
                    Dim cmt As Microsoft.Office.Interop.Word.Comment = candidates(i)
                    Dim text As String = If(cmt.Range.Text, String.Empty)
                    If text.StartsWith(riPrefix, StringComparison.Ordinal) Then
                        cmt.Range.Text = text.Substring(riPrefix.Length)
                        removedCount += 1
                    End If
                Catch
                End Try
            Next

            ProgressBarModule.CancelOperation = True

            ' If user requested date/time removal, do it via Open XML post-processing
            Dim dateTimeRemoved As Boolean = False
            If alsoRemoveDateTime AndAlso removedCount > 0 AndAlso commentIndicesToProcess.Count > 0 Then
                dateTimeRemoved = RemoveCommentDateTimesViaOpenXml(doc, app, commentIndicesToProcess)
            End If

            Dim scopeInfo As String = If(hasSelection, "the selected text", "the entire document")
            If ProgressBarModule.CancelOperation AndAlso removedCount < candidates.Count Then
                ShowCustomMessageBox($"Operation cancelled. Removed '{AN5}:' prefix from {removedCount} comment(s) in {scopeInfo}." &
                    If(dateTimeRemoved, " Date/time was also removed.", ""), AN)
            Else
                ShowCustomMessageBox($"Removed '{AN5}:' prefix from {removedCount} comment(s) in {scopeInfo}." &
                    If(dateTimeRemoved, " Date/time was also removed.", ""), AN)
            End If

        Catch ex As System.Exception
            ProgressBarModule.CancelOperation = True
            ShowCustomMessageBox($"Error in RemoveRIPrefixFromComments: {ex.Message}", AN)
        End Try
    End Sub

    ''' <summary>
    ''' Removes comment date/time stamps by saving the document to a temp DOCX, modifying the
    ''' w:date attributes in comments.xml (and commentsExtended/commentsExtensible.xml) via Open XML,
    ''' then reopening the document.
    ''' Instead of fragile index or text matching, processes ALL comments whose text (after prefix
    ''' removal) matches any of the expected stripped texts. Falls back to processing all comments
    ''' by the add-in's author if no matches are found.
    ''' </summary>
    ''' <param name="doc">The active Word document.</param>
    ''' <param name="app">The Word application instance.</param>
    ''' <param name="commentIndices">1-based Interop Comment.Index values (used only for count validation).</param>
    ''' <returns>True if date/time removal succeeded; False otherwise.</returns>
    Private Shared Function RemoveCommentDateTimesViaOpenXml(
            doc As Microsoft.Office.Interop.Word.Document,
            app As Microsoft.Office.Interop.Word.Application,
            commentIndices As List(Of Integer)) As Boolean

        Dim tempPath As String = Nothing
        Dim originalPath As String = Nothing
        Dim wasReadOnly As Boolean = False

        Try
            ' Get original path
            Try
                originalPath = doc.FullName
            Catch
                Return False
            End Try

            If String.IsNullOrEmpty(originalPath) Then Return False

            ' Check if the document is in a format that supports Open XML
            Dim ext As String = System.IO.Path.GetExtension(originalPath).ToLowerInvariant()
            If ext <> ".docx" AndAlso ext <> ".docm" Then
                ShowCustomMessageBox("Date/time removal requires a .docx or .docm document. Please save as .docx first.", AN)
                Return False
            End If

            wasReadOnly = doc.ReadOnly

            ' ─────────────────────────────────────────────────────────────────
            ' STEP 1: Before saving, collect the w:id values from the Interop
            '         comments we want to process. We do this by saving the doc
            '         first, reading the XML, and matching by author name.
            '         The author for add-in comments is the current Word UserName
            '         at time of creation, but comment text starts with the prefix.
            ' ─────────────────────────────────────────────────────────────────

            ' Collect the actual comment texts (already stripped of prefix by caller)
            ' so we can match in Open XML by content
            Dim strippedTexts As New HashSet(Of String)(StringComparer.Ordinal)
            For Each idx In commentIndices
                Try
                    If idx >= 1 AndAlso idx <= doc.Comments.Count Then
                        Dim cmtText As String = If(doc.Comments(idx).Range.Text, "")
                        ' The prefix has already been removed by the caller, so this is the post-strip text
                        strippedTexts.Add(cmtText.Trim())
                    End If
                Catch
                End Try
            Next

            ' Save to temp
            tempPath = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                $"{AN2}_cmtdate_{Guid.NewGuid():N}.docx")

            Dim prevScreenUpdating As Boolean = app.ScreenUpdating
            Dim prevAlerts As Microsoft.Office.Interop.Word.WdAlertLevel = app.DisplayAlerts
            Try
                app.ScreenUpdating = False
                app.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone
                doc.SaveAs2(tempPath, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument)
                doc.Close(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges)
            Finally
                app.DisplayAlerts = prevAlerts
                app.ScreenUpdating = prevScreenUpdating
            End Try

            ' ─────────────────────────────────────────────────────────────────
            ' STEP 2: Open XML — remove w:date from matching comments
            ' ─────────────────────────────────────────────────────────────────
            Dim modified As Boolean = False
            Dim modifiedWIds As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim modifiedParaIds As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            Using wordDoc As DocumentFormat.OpenXml.Packaging.WordprocessingDocument =
                    DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(tempPath, True)

                Dim commentsPart = wordDoc.MainDocumentPart?.WordprocessingCommentsPart
                If commentsPart IsNot Nothing AndAlso commentsPart.Comments IsNot Nothing Then

                    Dim wNs As System.Xml.Linq.XNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    Dim w14Ns As System.Xml.Linq.XNamespace = "http://schemas.microsoft.com/office/word/2010/wordml"

                    ' Load the entire comments part as XDocument for reliable attribute manipulation
                    Dim commentsXDoc As System.Xml.Linq.XDocument = Nothing
                    Using readStream As System.IO.Stream = commentsPart.GetStream(System.IO.FileMode.Open, System.IO.FileAccess.Read)
                        commentsXDoc = System.Xml.Linq.XDocument.Load(readStream)
                    End Using

                    Dim commentsChanged As Boolean = False

                    For Each commentElem In commentsXDoc.Descendants(wNs + "comment").ToList()

                        ' Extract plain text from w:t elements
                        Dim commentPlainText As String = String.Join("", commentElem.Descendants(wNs + "t").Select(Function(t) t.Value))
                        Dim trimmedText As String = commentPlainText.Trim()

                        ' Match: does this comment's text match one of the stripped texts?
                        If Not strippedTexts.Contains(trimmedText) Then Continue For

                        ' Collect w:id for extended parts
                        Dim wIdAttr = commentElem.Attribute(wNs + "id")
                        If wIdAttr IsNot Nothing AndAlso Not String.IsNullOrEmpty(wIdAttr.Value) Then
                            modifiedWIds.Add(wIdAttr.Value)
                        End If

                        ' Collect paraId values
                        For Each paraElem In commentElem.Descendants(wNs + "p")
                            Dim paraIdAttr = paraElem.Attribute(w14Ns + "paraId")
                            If paraIdAttr IsNot Nothing AndAlso Not String.IsNullOrEmpty(paraIdAttr.Value) Then
                                modifiedParaIds.Add(paraIdAttr.Value)
                            End If
                        Next

                        ' Remove the w:date attribute
                        Dim dateAttr = commentElem.Attribute(wNs + "date")
                        If dateAttr IsNot Nothing Then
                            dateAttr.Remove()
                            commentsChanged = True
                            modified = True
                        End If

                        ' Remove from set so we don't match duplicates unnecessarily
                        strippedTexts.Remove(trimmedText)
                    Next

                    If commentsChanged Then
                        Using writeStream As System.IO.Stream = commentsPart.GetStream(System.IO.FileMode.Create, System.IO.FileAccess.Write)
                            commentsXDoc.Save(writeStream)
                        End Using
                    End If
                End If

                ' ─────────────────────────────────────────────────────────────────
                ' STEP 3: Remove dateUtc from commentsExtended/commentsExtensible
                ' ─────────────────────────────────────────────────────────────────
                If modified AndAlso (modifiedParaIds.Count > 0 OrElse modifiedWIds.Count > 0) Then
                    Try
                        For Each extPart In wordDoc.MainDocumentPart.Parts
                            Try
                                Dim uri As String = If(extPart.OpenXmlPart?.Uri?.ToString(), "")
                                If Not (uri.Contains("commentsExtensible") OrElse
                                        uri.Contains("commentsExtended") OrElse
                                        uri.Contains("commentsEx")) Then Continue For

                                Using stream As System.IO.Stream = extPart.OpenXmlPart.GetStream(
                                    System.IO.FileMode.Open, System.IO.FileAccess.ReadWrite)

                                    Dim xdoc As System.Xml.Linq.XDocument = System.Xml.Linq.XDocument.Load(stream)
                                    Dim anyChanged As Boolean = False

                                    For Each elem In xdoc.Descendants().ToList()
                                        Dim shouldModify As Boolean = False

                                        ' Match by paraId
                                        Dim pIdAttr = elem.Attributes().FirstOrDefault(
                                            Function(a) a.Name.LocalName.Equals("paraId", StringComparison.OrdinalIgnoreCase))
                                        If pIdAttr IsNot Nothing AndAlso modifiedParaIds.Contains(pIdAttr.Value) Then
                                            shouldModify = True
                                        End If

                                        ' Match by commentId (w:id reference)
                                        If Not shouldModify Then
                                            Dim cIdAttr = elem.Attributes().FirstOrDefault(
                                                Function(a) a.Name.LocalName.Equals("commentId", StringComparison.OrdinalIgnoreCase))
                                            If cIdAttr IsNot Nothing AndAlso modifiedWIds.Contains(cIdAttr.Value) Then
                                                shouldModify = True
                                            End If
                                        End If

                                        If Not shouldModify Then Continue For

                                        For Each attr In elem.Attributes().ToList()
                                            If attr.Name.LocalName.Equals("dateUtc", StringComparison.OrdinalIgnoreCase) Then
                                                attr.Remove()
                                                anyChanged = True
                                            End If
                                        Next
                                    Next

                                    If anyChanged Then
                                        stream.SetLength(0)
                                        xdoc.Save(stream)
                                    End If
                                End Using
                            Catch
                            End Try
                        Next
                    Catch
                    End Try
                End If
            End Using

            ' ─────────────────────────────────────────────────────────────────
            ' STEP 4: Reopen the document
            ' ─────────────────────────────────────────────────────────────────
            Dim prevScreenUpdating2 As Boolean = app.ScreenUpdating
            Dim prevAlerts2 As Microsoft.Office.Interop.Word.WdAlertLevel = app.DisplayAlerts
            Try
                app.ScreenUpdating = False
                app.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone

                Dim reopened As Microsoft.Office.Interop.Word.Document = app.Documents.Open(
                    FileName:=tempPath,
                    ReadOnly:=False,
                    Visible:=True)

                reopened.SaveAs2(originalPath, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument)
            Finally
                app.DisplayAlerts = prevAlerts2
                app.ScreenUpdating = prevScreenUpdating2
            End Try

            Return modified

        Catch ex As System.Exception
            Try
                If Not String.IsNullOrEmpty(originalPath) AndAlso System.IO.File.Exists(originalPath) Then
                    app.Documents.Open(FileName:=originalPath, ReadOnly:=wasReadOnly, Visible:=True)
                ElseIf Not String.IsNullOrEmpty(tempPath) AndAlso System.IO.File.Exists(tempPath) Then
                    app.Documents.Open(FileName:=tempPath, ReadOnly:=False, Visible:=True)
                End If
            Catch
            End Try
            Return False
        Finally
            Try
                If Not String.IsNullOrEmpty(tempPath) AndAlso System.IO.File.Exists(tempPath) Then
                    System.IO.File.Delete(tempPath)
                End If
            Catch
            End Try
        End Try
    End Function


    ''' <summary>
    ''' Extracts plain text from an Open XML Comment element by collecting text from all w:t descendants.
    ''' This is more reliable than InnerText which may include unexpected whitespace from XML structure.
    ''' </summary>
    Private Shared Function ExtractPlainTextFromOpenXmlComment(comment As DocumentFormat.OpenXml.Wordprocessing.Comment) As String
        If comment Is Nothing Then Return Nothing
        Dim sb As New System.Text.StringBuilder()
        For Each textElement In comment.Descendants(Of DocumentFormat.OpenXml.Wordprocessing.Text)()
            sb.Append(textElement.Text)
        Next
        Return sb.ToString()
    End Function

    ''' <summary>
    ''' Normalizes comment text for comparison by removing control characters (BEL, vertical tab,
    ''' form feed, etc.) that Word Interop and Open XML may handle differently.
    ''' </summary>
    Private Shared Function NormalizeCommentText(text As String) As String
        If text Is Nothing Then Return String.Empty
        Dim sb As New System.Text.StringBuilder(text.Length)
        For Each ch As Char In text
            ' Skip control characters except standard whitespace (space, tab, CR, LF)
            If Char.IsControl(ch) AndAlso ch <> " "c AndAlso ch <> vbTab(0) AndAlso ch <> vbCr(0) AndAlso ch <> vbLf(0) Then
                Continue For
            End If
            sb.Append(ch)
        Next
        ' Normalize line endings
        Return sb.ToString().Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).TrimEnd()
    End Function

    ''' <summary>
    ''' Removes leading "RI: " prefix from Word comments (including threaded replies) in the selection or entire document.
    ''' Uses a progress bar and supports cancellation.
    ''' </summary>
    Public Sub RemoveRIPrefixFromComments()

        Dim riPrefix As String = $"{AN5}: "

        Try
            Dim app As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application
            Dim doc As Microsoft.Office.Interop.Word.Document = Nothing

            Try
                doc = app.ActiveDocument
            Catch
            End Try

            If doc Is Nothing Then
                ShowCustomMessageBox("No active document found.", AN)
                Exit Sub
            End If

            Dim sel As Microsoft.Office.Interop.Word.Selection = app.Selection
            Dim hasSelection As Boolean = (sel IsNot Nothing AndAlso sel.Range IsNot Nothing AndAlso sel.Range.Start <> sel.Range.End)
            Dim scopeRange As Microsoft.Office.Interop.Word.Range = If(hasSelection, sel.Range, doc.Content)

            Dim candidates As New List(Of Microsoft.Office.Interop.Word.Comment)()

            ' Collect candidates that start with the prefix (base comments + replies)
            For Each c As Microsoft.Office.Interop.Word.Comment In doc.Comments
                Dim inScope As Boolean = False

                Try
                    Dim cStart As Integer = c.Scope.Start
                    Dim cEnd As Integer = c.Scope.End
                    inScope = (cStart >= scopeRange.Start AndAlso cEnd <= scopeRange.End)
                Catch
                    Try
                        Dim cStart As Integer = c.Reference.Start
                        Dim cEnd As Integer = c.Reference.End
                        inScope = (cStart >= scopeRange.Start AndAlso cEnd <= scopeRange.End)
                    Catch
                        inScope = False
                    End Try
                End Try

                If Not inScope Then Continue For

                Try
                    Dim text As String = If(c.Range.Text, String.Empty)
                    If text.StartsWith(riPrefix, StringComparison.Ordinal) Then
                        candidates.Add(c)
                    End If
                Catch
                End Try

                ' Threaded replies
                Try
                    If c.Replies IsNot Nothing AndAlso c.Replies.Count > 0 Then
                        For Each r As Microsoft.Office.Interop.Word.Comment In c.Replies
                            Try
                                Dim replyText As String = If(r.Range.Text, String.Empty)
                                If replyText.StartsWith(riPrefix, StringComparison.Ordinal) Then
                                    candidates.Add(r)
                                End If
                            Catch
                            End Try
                        Next
                    End If
                Catch
                End Try
            Next

            If candidates.Count = 0 Then
                ShowCustomMessageBox($"No '{AN5}:' prefixes found in comments for the current scope.", AN)
                Exit Sub
            End If

            ProgressBarModule.GlobalProgressValue = 0
            ProgressBarModule.GlobalProgressMax = candidates.Count
            ProgressBarModule.GlobalProgressLabel = "Processing comments..."
            ProgressBarModule.CancelOperation = False
            ProgressBarModule.ShowProgressBarInSeparateThread(AN & $" Remove {AN5} Prefix", "Starting...")

            Dim removedCount As Integer = 0

            For i As Integer = 0 To candidates.Count - 1
                System.Windows.Forms.Application.DoEvents()

                If ProgressBarModule.CancelOperation Then Exit For
                If (GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0 Then
                    ProgressBarModule.CancelOperation = True
                    Exit For
                End If

                ProgressBarModule.GlobalProgressValue = i + 1
                ProgressBarModule.GlobalProgressLabel = $"Processing {i + 1} of {candidates.Count}..."

                Try
                    Dim cmt As Microsoft.Office.Interop.Word.Comment = candidates(i)
                    Dim text As String = If(cmt.Range.Text, String.Empty)
                    If text.StartsWith(riPrefix, StringComparison.Ordinal) Then
                        cmt.Range.Text = text.Substring(riPrefix.Length)
                        removedCount += 1
                    End If
                Catch
                End Try
            Next

            ProgressBarModule.CancelOperation = True

            Dim scopeInfo As String = If(hasSelection, "the selected text", "the entire document")
            If ProgressBarModule.CancelOperation AndAlso removedCount < candidates.Count Then
                ShowCustomMessageBox($"Operation cancelled. Removed '{AN5}:' prefix from {removedCount} comment(s) in {scopeInfo}.", AN)
            Else
                ShowCustomMessageBox($"Removed '{AN5}:' prefix from {removedCount} comment(s) in {scopeInfo}.", AN)
            End If

        Catch ex As System.Exception
            ProgressBarModule.CancelOperation = True
            ShowCustomMessageBox($"Error in RemoveRIPrefixFromComments: {ex.Message}", AN)
        End Try
    End Sub

    ''' <summary>
    ''' Reads a PDF page by page, extracts text (with OCR if needed), sends the page texts
    ''' to the LLM to determine how to split the PDF into separate exhibits/documents,
    ''' and saves each split segment into a new subdirectory alongside the source PDF.
    ''' Output filenames follow the pattern "Exhibit 01 - Brief Description.pdf".
    ''' Shows a progress bar and supports cancellation.
    ''' </summary>
    Public Async Sub SplitPdfByExhibits()
        Dim selectedPath As String = ""

        ' Show DragDropForm for PDF file selection
        Globals.ThisAddIn.DragDropFormLabel = "Select a PDF file to split into exhibits"

        Try
            Using frm As New DragDropForm(DragDropMode.FileOnly)
                If frm.ShowDialog() = DialogResult.OK Then
                    selectedPath = frm.SelectedFilePath
                End If
            End Using
        Finally
            Globals.ThisAddIn.DragDropFormLabel = ""
        End Try

        If String.IsNullOrWhiteSpace(selectedPath) Then
            Return
        End If

        If Not IO.File.Exists(selectedPath) Then
            ShowCustomMessageBox("The selected file does not exist.", AN)
            Return
        End If

        Dim ext As String = IO.Path.GetExtension(selectedPath).ToLowerInvariant()
        If ext <> ".pdf" Then
            ShowCustomMessageBox($"File type '{ext}' is not supported. Only PDF files can be split.", AN)
            Return
        End If

        ' Determine OCR availability and ask user
        Dim doOcr As Boolean = False
        If SharedMethods.IsOcrAvailable(_context) Then
            Dim ocrChoice As Integer = ShowCustomYesNoBox(
                "Some pages may contain scanned content that requires OCR (optical character recognition) to extract text." & vbCrLf & vbCrLf &
                "Do you want to enable OCR for pages where text extraction yields little or no content?",
                "Yes, enable OCR",
                "No, skip OCR",
                AN & " Split PDF by Exhibits")
            If ocrChoice = 0 Then
                Return ' User cancelled
            End If
            doOcr = (ocrChoice = 1)
        End If

        ' Prompt user for output language for file descriptions
        ' If empty, no description will be added to filenames (just 01.pdf, 02.pdf, etc.)
        Dim useDescription As Boolean = True
        Dim languageInput As String = ShowCustomInputBox(
            "Enter the language for exhibit descriptions in the filenames (e.g., English, German, French)." & vbCrLf & vbCrLf &
            "Leave empty if you do not want descriptions added to the filenames (files will be named 01.pdf, 02.pdf, etc.).",
            AN & " Split PDF by Exhibits",
            True,
            INI_Language1)

        If languageInput Is Nothing Then
            Return ' User cancelled (pressed Cancel/Esc)
        End If

        languageInput = languageInput.Trim()

        If String.IsNullOrEmpty(languageInput) Then
            ' No description wanted; set OutputLanguage to English for internal LLM use
            OutputLanguage = "English"
            useDescription = False
        Else
            OutputLanguage = languageInput
            useDescription = True
        End If

        ' ─────────────────────────────────────────────────────────────────
        ' STEP 1: Extract text page by page
        ' ─────────────────────────────────────────────────────────────────
        Dim pageCount As Integer = 0

        ' Get page count first
        Try
            Using pdfDoc As UglyToad.PdfPig.PdfDocument = UglyToad.PdfPig.PdfDocument.Open(selectedPath)
                pageCount = pdfDoc.NumberOfPages
            End Using
        Catch ex As Exception
            ShowCustomMessageBox($"Failed to open PDF: {ex.Message}", AN)
            Return
        End Try

        If pageCount = 0 Then
            ShowCustomMessageBox("The PDF has no pages.", AN)
            Return
        End If

        If pageCount = 1 Then
            ShowCustomMessageBox("The PDF has only one page and cannot be split.", AN)
            Return
        End If

        ' Show progress bar for text extraction
        ProgressBarModule.GlobalProgressValue = 0
        ProgressBarModule.GlobalProgressMax = pageCount
        ProgressBarModule.GlobalProgressLabel = "Extracting text from pages..."
        ProgressBarModule.CancelOperation = False
        ProgressBarModule.ShowProgressBarInSeparateThread(AN & " Split PDF by Exhibits", "Reading PDF pages...")

        Dim pageTexts As New Dictionary(Of Integer, String)()
        Dim ocrWasUsed As Boolean = False

        Try
            ' Extract text using PdfPig page by page
            Using pdfDoc As UglyToad.PdfPig.PdfDocument = UglyToad.PdfPig.PdfDocument.Open(selectedPath)
                For pageIdx As Integer = 1 To pdfDoc.NumberOfPages
                    If ProgressBarModule.CancelOperation Then Exit For

                    ProgressBarModule.GlobalProgressValue = pageIdx
                    ProgressBarModule.GlobalProgressLabel = $"Extracting text from page {pageIdx} of {pdfDoc.NumberOfPages}..."

                    Dim pageText As String = String.Empty

                    Try
                        Dim page = pdfDoc.GetPage(pageIdx)
                        pageText = String.Join(" ", page.GetWords().Select(Function(w) w.Text))
                    Catch
                        ' Page text extraction failed
                    End Try

                    ' If text is sparse and OCR is enabled, use OCR for this page
                    If doOcr AndAlso (String.IsNullOrWhiteSpace(pageText) OrElse pageText.Trim().Length < 50) Then
                        Try
                            ProgressBarModule.GlobalProgressLabel = $"OCR processing page {pageIdx} of {pdfDoc.NumberOfPages}..."

                            ' Extract single page to temp PDF for OCR
                            Dim tempPagePdf As String = IO.Path.Combine(IO.Path.GetTempPath(), $"{AN2}_splitocr_{Guid.NewGuid():N}.pdf")
                            Try
                                Await System.Threading.Tasks.Task.Run(Sub() ExtractSinglePagePdf(selectedPath, pageIdx, tempPagePdf))

                                Dim ocrResult As String = Await ReadPdfAsText(tempPagePdf, False, True, False, _context)
                                If Not String.IsNullOrWhiteSpace(ocrResult) AndAlso ocrResult.Trim().Length > pageText.Trim().Length Then
                                    pageText = ocrResult
                                    ocrWasUsed = True
                                End If
                            Finally
                                Try
                                    If IO.File.Exists(tempPagePdf) Then IO.File.Delete(tempPagePdf)
                                Catch
                                End Try
                            End Try
                        Catch
                            ' OCR failed for this page; use whatever text we have
                        End Try
                    End If

                    pageTexts(pageIdx) = If(pageText, String.Empty)
                Next
            End Using

        Catch ex As Exception
            ProgressBarModule.CancelOperation = True
            ShowCustomMessageBox($"Failed to extract text from PDF: {ex.Message}", AN)
            Return
        End Try

        ' Close extraction progress
        ProgressBarModule.CancelOperation = True

        If pageTexts.Count = 0 Then
            ShowCustomMessageBox("No text could be extracted from the PDF.", AN)
            Return
        End If

        ' ─────────────────────────────────────────────────────────────────
        ' STEP 2: Send page texts to LLM for split analysis
        ' ─────────────────────────────────────────────────────────────────

        ' Build the page-by-page content for the LLM
        Dim pageContentSb As New System.Text.StringBuilder()

        ' Dynamically scale per-page character budget to stay within token limits
        ' Target: ~200K chars total (~50K tokens) to leave room for system prompt + response
        Const MaxTotalChars As Integer = 200000
        Dim charsPerPage As Integer = Math.Max(300, MaxTotalChars \ Math.Max(1, pageTexts.Count))
        ' Cap at 2250 so small documents still get the full truncation treatment
        charsPerPage = Math.Min(charsPerPage, 2250)
        ' Reserve 250 chars for the tail portion; the rest goes to the head
        Dim tailChars As Integer = Math.Min(250, charsPerPage \ 4)
        Dim headChars As Integer = charsPerPage - tailChars

        For Each kvp In pageTexts.OrderBy(Function(k) k.Key)
            Dim truncatedText As String = kvp.Value
            ' Truncate pages that exceed the per-page budget, keeping head + tail
            If truncatedText.Length > charsPerPage Then
                truncatedText = truncatedText.Substring(0, headChars) & " [...] " & truncatedText.Substring(truncatedText.Length - tailChars)
            End If
            pageContentSb.AppendLine($"--- PAGE {kvp.Key} ---")
            pageContentSb.AppendLine(If(String.IsNullOrWhiteSpace(truncatedText), "[No text extracted from this page]", truncatedText))
            pageContentSb.AppendLine()
        Next

        Dim systemPrompt As String = SP_SplitPDF

        Dim userPrompt As String = $"The PDF has {pageCount} pages. Here is the text content page by page:" & vbCrLf & vbCrLf & pageContentSb.ToString()

        ' ── Try to use a dedicated SplitPDF model (if configured in alternate models INI) ──
        Dim backupConfig As ModelConfig = Nothing
        Dim useSplitPdfModel As Boolean = False
        Dim useSecondApi As Boolean = False
        Dim timeout As Long = 0

        Try
            If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                useSplitPdfModel = GetSpecialTaskModel(_context, INI_AlternateModelPath, "SplitPDF")
            End If
        Catch
        End Try

        If useSplitPdfModel Then
            backupConfig = GetCurrentConfig(_context)
            useSecondApi = True
            timeout = If(_context.INI_Timeout_2 > 0, _context.INI_Timeout_2, _context.INI_Timeout)
        Else
            useSecondApi = False
            timeout = _context.INI_Timeout
        End If

        Dim llmResponse As String = String.Empty
        Try
            llmResponse = Await SharedMethods.LLM(
                _context,
                InterpolateAtRuntime(systemPrompt),
                userPrompt,
                "",
                "",
                timeout,
                useSecondApi,
                False)
        Catch ex As Exception
            ShowCustomMessageBox($"Failed to analyze PDF structure: {ex.Message}", AN)
            Return
        Finally
            If useSplitPdfModel AndAlso backupConfig IsNot Nothing Then
                RestoreDefaults(_context, backupConfig)
            End If
        End Try

        ' ─────────────────────────────────────────────────────────────────
        ' STEP 3: Parse LLM response into split segments
        ' ─────────────────────────────────────────────────────────────────
        Dim segments As New List(Of SplitSegment)()

        Try
            ' Clean response: strip code fences if present
            Dim cleanedResponse As String = llmResponse.Trim()
            If cleanedResponse.StartsWith("```") Then
                Dim firstNewline As Integer = cleanedResponse.IndexOf(vbLf)
                If firstNewline >= 0 Then
                    cleanedResponse = cleanedResponse.Substring(firstNewline + 1)
                End If
                If cleanedResponse.EndsWith("```") Then
                    cleanedResponse = cleanedResponse.Substring(0, cleanedResponse.Length - 3).Trim()
                End If
            End If

            Dim jsonArray As Newtonsoft.Json.Linq.JArray = Newtonsoft.Json.Linq.JArray.Parse(cleanedResponse)

            For Each item As Newtonsoft.Json.Linq.JObject In jsonArray
                Dim seg As New SplitSegment()
                seg.StartPage = CInt(item("start_page"))
                seg.EndPage = CInt(item("end_page"))
                seg.Name = CStr(If(item("name"), "Unnamed"))
                segments.Add(seg)
            Next
        Catch ex As Exception
            ShowCustomMessageBox($"Failed to parse the AI response as valid JSON: {ex.Message}" & vbCrLf & vbCrLf &
                                 "Raw response:" & vbCrLf & llmResponse, AN)
            Return
        End Try

        If segments.Count = 0 Then
            ShowCustomMessageBox("The AI returned no segments. Unable to split the PDF.", AN)
            Return
        End If

        ' Validate segments
        For Each seg In segments
            If seg.StartPage < 1 OrElse seg.EndPage > pageCount OrElse seg.StartPage > seg.EndPage Then
                ShowCustomMessageBox($"Invalid page range detected: pages {seg.StartPage}-{seg.EndPage} for '{seg.Name}'. Total pages: {pageCount}.", AN)
                Return
            End If
        Next

        ' Sort by start page
        segments.Sort(Function(a, b) a.StartPage.CompareTo(b.StartPage))

        If segments.Count = 1 Then
            ShowCustomMessageBox("The AI determined this PDF is a single coherent document and cannot be meaningfully split." & vbCrLf & vbCrLf &
                                 $"Document: {segments(0).Name} (pages {segments(0).StartPage}-{segments(0).EndPage})", AN)
            Return
        End If

        ' Show proposed split to user for confirmation
        Dim previewSb As New System.Text.StringBuilder()
        previewSb.AppendLine($"The AI proposes splitting this {pageCount}-page PDF into {segments.Count} parts:")
        previewSb.AppendLine()
        For i As Integer = 0 To segments.Count - 1
            Dim seg = segments(i)
            Dim pageInfo As String = If(seg.StartPage = seg.EndPage, $"page {seg.StartPage}", $"pages {seg.StartPage}-{seg.EndPage}")
            previewSb.AppendLine($"  {(i + 1):D2}. {seg.Name} ({pageInfo})")
        Next
        previewSb.AppendLine()
        previewSb.AppendLine("Do you want to proceed with this split?")

        Dim confirmChoice As Integer = ShowCustomYesNoBox(
            previewSb.ToString(),
            "Yes, split PDF",
            "No, cancel",
            AN & " Split PDF by Exhibits")
        If confirmChoice <> 1 Then
            Return
        End If

        ' ─────────────────────────────────────────────────────────────────
        ' STEP 4: Split the PDF using PdfSharp
        ' ─────────────────────────────────────────────────────────────────
        Dim pdfDir As String = IO.Path.GetDirectoryName(selectedPath)
        Dim pdfBaseName As String = IO.Path.GetFileNameWithoutExtension(selectedPath)
        Dim outputDir As String = IO.Path.Combine(pdfDir, SanitizeFileName(pdfBaseName) & "_split")

        ' Ensure unique output directory
        If IO.Directory.Exists(outputDir) Then
            Dim counter As Integer = 1
            While IO.Directory.Exists(outputDir & $"_{counter}")
                counter += 1
            End While
            outputDir = outputDir & $"_{counter}"
        End If

        Try
            IO.Directory.CreateDirectory(outputDir)
        Catch ex As Exception
            ShowCustomMessageBox($"Failed to create output directory: {ex.Message}", AN)
            Return
        End Try

        ' Show progress bar for splitting
        ProgressBarModule.GlobalProgressValue = 0
        ProgressBarModule.GlobalProgressMax = segments.Count
        ProgressBarModule.GlobalProgressLabel = "Splitting PDF..."
        ProgressBarModule.CancelOperation = False
        ProgressBarModule.ShowProgressBarInSeparateThread(AN & " Split PDF by Exhibits", "Splitting PDF into exhibits...")

        Dim successCount As Integer = 0
        Dim failedSegments As New List(Of String)()

        Try
            ' Load the source PDF once with PdfSharp
            Using srcDoc As PdfSharp.Pdf.PdfDocument = PdfSharp.Pdf.IO.PdfReader.Open(selectedPath, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import)

                For i As Integer = 0 To segments.Count - 1
                    If ProgressBarModule.CancelOperation Then Exit For

                    Dim seg = segments(i)
                    ProgressBarModule.GlobalProgressValue = i + 1
                    ProgressBarModule.GlobalProgressLabel = $"Creating exhibit {i + 1} of {segments.Count}: {seg.Name}..."

                    Try
                        ' Build filename based on whether descriptions are wanted
                        Dim outputFileName As String
                        If useDescription Then
                            Dim sanitizedName As String = SanitizeFileName(seg.Name)
                            outputFileName = $"{(i + 1):D2} - {sanitizedName}.pdf"
                        Else
                            outputFileName = $"{(i + 1):D2}.pdf"
                        End If
                        Dim outputPath As String = IO.Path.Combine(outputDir, outputFileName)

                        ' Create a new document with just the pages for this segment
                        Using outDoc As New PdfSharp.Pdf.PdfDocument()
                            For pageNum As Integer = seg.StartPage To seg.EndPage
                                ' PdfSharp uses 0-based indexing
                                Dim srcPage As PdfSharp.Pdf.PdfPage = srcDoc.Pages(pageNum - 1)
                                outDoc.AddPage(srcPage)
                            Next

                            outDoc.Save(outputPath)
                        End Using

                        successCount += 1

                    Catch ex As Exception
                        failedSegments.Add($"{seg.Name}: {ex.Message}")
                    End Try
                Next
            End Using

        Catch ex As Exception
            ProgressBarModule.CancelOperation = True
            ShowCustomMessageBox($"Failed to split PDF: {ex.Message}", AN)
            Return
        Finally
            ProgressBarModule.CancelOperation = True
        End Try

        ' ─────────────────────────────────────────────────────────────────
        ' STEP 5: Show summary
        ' ─────────────────────────────────────────────────────────────────
        Dim summary As New System.Text.StringBuilder()
        summary.AppendLine($"Successfully created {successCount} exhibit file(s).")
        summary.AppendLine($"Output directory: {outputDir}")

        If ocrWasUsed Then
            summary.AppendLine("OCR was used for some pages with little or no text.")
        End If

        If failedSegments.Count > 0 Then
            summary.AppendLine()
            summary.AppendLine($"Failed: {failedSegments.Count} segment(s)")
            For Each f In failedSegments
                summary.AppendLine($"  • {f}")
            Next
        End If

        ' Ask if user wants to open the output directory
        Dim openChoice As Integer = ShowCustomYesNoBox(
            summary.ToString() & vbCrLf & vbCrLf & "Do you want to open the output directory?",
            "Yes, open directory",
            "No",
            AN & " Split PDF by Exhibits")

        If openChoice = 1 Then
            Try
                System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outputDir) With {.UseShellExecute = True})
            Catch ex As Exception
                ShowCustomMessageBox($"Could not open directory: {ex.Message}", AN)
            End Try
        End If
    End Sub

    ''' <summary>
    ''' Extracts a single page from a PDF file and saves it as a separate PDF.
    ''' Used for per-page OCR processing during exhibit splitting.
    ''' </summary>
    ''' <param name="inputPath">Source PDF file path.</param>
    ''' <param name="pageNumber">1-based page number to extract.</param>
    ''' <param name="outputPath">Destination PDF file path for the single page.</param>
    Private Shared Sub ExtractSinglePagePdf(inputPath As String, pageNumber As Integer, outputPath As String)
        Using srcDoc As PdfSharp.Pdf.PdfDocument = PdfSharp.Pdf.IO.PdfReader.Open(inputPath, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import)
            Using outDoc As New PdfSharp.Pdf.PdfDocument()
                outDoc.AddPage(srcDoc.Pages(pageNumber - 1))
                outDoc.Save(outputPath)
            End Using
        End Using
    End Sub

    Private NotInheritable Class MarkupAuthorScope
        Implements IDisposable

        Private ReadOnly _app As Word.Application
        Private ReadOnly _originalUserName As String
        Private ReadOnly _originalUserInitials As String
        Private _shouldRestore As Boolean

        Public Sub New(app As Word.Application)
            _app = app
            If _app Is Nothing Then Exit Sub

            Try
                _originalUserName = If(_app.UserName, String.Empty)
            Catch
                _originalUserName = String.Empty
            End Try

            Try
                _originalUserInitials = If(_app.UserInitials, String.Empty)
            Catch
                _originalUserInitials = String.Empty
            End Try

            Dim markupAuthor As String = GetMarkupAuthorOrCurrent(_app)
            If String.IsNullOrWhiteSpace(markupAuthor) Then Exit Sub
            If String.Equals(markupAuthor, _originalUserName, StringComparison.Ordinal) Then Exit Sub

            Try
                _app.UserName = markupAuthor
                _app.UserInitials = BuildMarkupAuthorInitials(markupAuthor, _originalUserInitials)
                _shouldRestore = True
            Catch
                _shouldRestore = False
            End Try
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            If Not _shouldRestore OrElse _app Is Nothing Then Return

            Try
                _app.UserName = _originalUserName
            Catch
            End Try

            Try
                _app.UserInitials = _originalUserInitials
            Catch
            End Try
        End Sub
    End Class

    Friend Shared Function GetMarkupAuthorOrCurrent(app As Word.Application) As String
        Dim markupAuthor As String = String.Empty

        Try
            markupAuthor = If(Globals.ThisAddIn.INI_MarkupAuthor, String.Empty).Trim()
        Catch
            markupAuthor = String.Empty
        End Try

        If Not String.IsNullOrWhiteSpace(markupAuthor) Then
            Return markupAuthor
        End If

        If app Is Nothing Then Return String.Empty

        Try
            Return If(app.UserName, String.Empty)
        Catch
            Return String.Empty
        End Try
    End Function

    Private Shared Function BuildMarkupAuthorInitials(markupAuthor As String, fallbackInitials As String) As String
        If String.IsNullOrWhiteSpace(markupAuthor) Then Return fallbackInitials

        Dim parts As String() =
            markupAuthor.Trim().Split(New Char() {" "c, ControlChars.Tab}, StringSplitOptions.RemoveEmptyEntries)

        If parts.Length = 0 Then Return fallbackInitials

        Dim initials As String = String.Empty

        For Each part As String In parts
            If part.Length > 0 Then
                initials &= Char.ToUpperInvariant(part(0))
                If initials.Length >= 3 Then Exit For
            End If
        Next

        If initials.Length = 1 AndAlso parts(0).Length > 1 Then
            initials &= Char.ToUpperInvariant(parts(0)(1))
        End If

        Return If(initials.Length > 0, initials, fallbackInitials)
    End Function

    Friend Shared Function BeginMarkupAuthorScope(app As Word.Application) As IDisposable
        Return New MarkupAuthorScope(app)
    End Function




End Class




