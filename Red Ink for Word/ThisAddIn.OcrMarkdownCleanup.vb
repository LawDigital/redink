' Part of "Red Ink for Word"
' Explicit OCR cleanup entry points. Source selection/files are never overwritten.
Option Explicit On
Option Strict On
Option Infer On

Partial Public Class ThisAddIn
    Private Function SelectOcrMarkdownCleanupMode(includeDisabled As System.Boolean) As SharedLibrary.SharedLibrary.SharedMethods.OcrMarkdownCleanupMode
        Dim items As New System.Collections.Generic.List(Of SharedLibrary.SharedLibrary.SharedMethods.SelectionItem)()
        If includeDisabled Then items.Add(New SharedLibrary.SharedLibrary.SharedMethods.SelectionItem("Keep OCR output unchanged (default)", 1))
        items.Add(New SharedLibrary.SharedLibrary.SharedMethods.SelectionItem("Join wrapped prose lines; keep all text", 2))
        items.Add(New SharedLibrary.SharedLibrary.SharedMethods.SelectionItem("Also remove exact repeated page headers; log the cleanup report", 3))
        Dim selected As System.Int32 = SharedLibrary.SharedLibrary.SharedMethods.SelectValue(items, If(includeDisabled, 1, 2),
            "Optional OCR cleanup. Numbers, footers, signatures, code, footnote identifiers and ambiguous hyphens are retained. Header removal needs actual page boundaries (form feeds in a text file).",
            AN & " OCR cleanup")
        Select Case selected
            Case 2
                Return SharedLibrary.SharedLibrary.SharedMethods.OcrMarkdownCleanupMode.JoinWrappedLines
            Case 3
                Return SharedLibrary.SharedLibrary.SharedMethods.OcrMarkdownCleanupMode.JoinWrappedLinesAndRepeatedHeaders
            Case Else
                Return SharedLibrary.SharedLibrary.SharedMethods.OcrMarkdownCleanupMode.None
        End Select
    End Function

    Private Function SelectPdfMarkdownPreparationOptions(ByRef legacyMode As SharedLibrary.SharedLibrary.SharedMethods.OcrMarkdownCleanupMode) As SharedLibrary.SharedLibrary.SharedMethods.PdfMarkdownPreparationOptions
        legacyMode = SharedLibrary.SharedLibrary.SharedMethods.OcrMarkdownCleanupMode.None
        Dim items As New System.Collections.Generic.List(Of SharedLibrary.SharedLibrary.SharedMethods.SelectionItem) From {
            New SharedLibrary.SharedLibrary.SharedMethods.SelectionItem("Keep extracted Markdown unchanged (default)", 1),
            New SharedLibrary.SharedLibrary.SharedMethods.SelectionItem("Reading copy: prose flow and repeated page artifacts", 2),
            New SharedLibrary.SharedLibrary.SharedMethods.SelectionItem("Document structure: footnotes, margins, headings and page continuations", 3),
            New SharedLibrary.SharedLibrary.SharedMethods.SelectionItem("Existing conservative OCR cleanup", 4)}
        Dim choice As System.Int32 = SharedLibrary.SharedLibrary.SharedMethods.SelectValue(items, 1,
            "Choose Markdown preparation for the PDFs in this conversion. Structure preparation uses additional model requests; the original wording is retained. This choice does not change your configuration.", AN & " PDF to Markdown")
        If choice = 0 Then Return Nothing
        Dim options As New SharedLibrary.SharedLibrary.SharedMethods.PdfMarkdownPreparationOptions()
        If choice = 1 Then Return options
        If choice = 4 Then
            legacyMode = SelectOcrMarkdownCleanupMode(True)
            Dim rawParameters As SharedLibrary.SharedLibrary.SharedMethods.InputParameter() = {
                New SharedLibrary.SharedLibrary.SharedMethods.InputParameter("Also save an uncleaned .md copy", False)}
            If Not SharedLibrary.SharedLibrary.SharedMethods.ShowCustomVariableInputForm("The cleanup report is copied to the clipboard at the end. Only the final Markdown is saved unless selected below.", AN & " OCR cleanup", rawParameters) Then Return Nothing
            options.SaveRawCopy = System.Convert.ToBoolean(rawParameters(0).Value, System.Globalization.CultureInfo.InvariantCulture)
            Return options
        End If
        Dim structured As System.Boolean = choice = 3
        Dim parameters As SharedLibrary.SharedLibrary.SharedMethods.InputParameter() = {
            New SharedLibrary.SharedLibrary.SharedMethods.InputParameter("Join wrapped prose lines (preserve code, lists, tables and explicit breaks)", True),
            New SharedLibrary.SharedLibrary.SharedMethods.InputParameter("Remove repeated page-edge headers/footers and verified page-number sequences", True),
            New SharedLibrary.SharedLibrary.SharedMethods.InputParameter("Collect all identified footnotes at the document end (Markdown notes)", structured),
            New SharedLibrary.SharedLibrary.SharedMethods.InputParameter("Represent marginal numbers/text consistently", structured),
            New SharedLibrary.SharedLibrary.SharedMethods.InputParameter("Reconcile heading levels across the entire document", structured),
            New SharedLibrary.SharedLibrary.SharedMethods.InputParameter("Join confirmed paragraph continuations across pages", structured),
            New SharedLibrary.SharedLibrary.SharedMethods.InputParameter("Use model structure analysis (additional requests; needed for structural inference)", structured),
            New SharedLibrary.SharedLibrary.SharedMethods.InputParameter("Also save an uncleaned .md copy", False)}
        If Not SharedLibrary.SharedLibrary.SharedMethods.ShowCustomVariableInputForm(
            "These options apply to all PDFs in this conversion. Uncertain or invalid transformations retain the source and are reported in the detailed log copied to the clipboard at the end (Ctrl+V to paste). All identified footnotes are collected when selected; unmatched references are retained and logged. Printed OCR footnotes need model structure analysis; existing Markdown notes do not. Without model analysis, only existing footnotes, prose lines and repeated page-edge patterns are processed. For PDFs read without OCR, the existing layout reader uses source geometry; the later Markdown pass preserves its tables and notes.",
            AN & " Markdown preparation", parameters, checkboxRightClearance:=24) Then Return Nothing
        options.JoinProseLines = System.Convert.ToBoolean(parameters(0).Value, System.Globalization.CultureInfo.InvariantCulture)
        options.RemovePageArtifacts = System.Convert.ToBoolean(parameters(1).Value, System.Globalization.CultureInfo.InvariantCulture)
        options.CollectFootnotes = System.Convert.ToBoolean(parameters(2).Value, System.Globalization.CultureInfo.InvariantCulture)
        options.NormalizeMarginNotes = System.Convert.ToBoolean(parameters(3).Value, System.Globalization.CultureInfo.InvariantCulture)
        options.NormalizeHeadings = System.Convert.ToBoolean(parameters(4).Value, System.Globalization.CultureInfo.InvariantCulture)
        options.JoinPageParagraphs = System.Convert.ToBoolean(parameters(5).Value, System.Globalization.CultureInfo.InvariantCulture)
        options.UseModelStructure = System.Convert.ToBoolean(parameters(6).Value, System.Globalization.CultureInfo.InvariantCulture)
        options.SaveRawCopy = System.Convert.ToBoolean(parameters(7).Value, System.Globalization.CultureInfo.InvariantCulture)
        Return options
    End Function

    Private Async Function RunOcrMarkdownCleanupCommandAsync() As System.Threading.Tasks.Task
        Dim sourceText As System.String = System.String.Empty
        Dim sourcePath As System.String = Nothing
        Dim currentSelection As Microsoft.Office.Interop.Word.Selection = Me.Application.Selection
        If currentSelection IsNot Nothing AndAlso currentSelection.Range.Start <> currentSelection.Range.End Then
            If currentSelection.Tables.Count > 0 Then
                SharedLibrary.SharedLibrary.SharedMethods.ShowCustomMessageBox("Select plain OCR/Markdown text outside Word tables, or run ocrclean without a selection to choose a .md/.txt file.")
                Return
            End If
            sourceText = currentSelection.Text
        Else
            sourcePath = GetFileName()
            If System.String.IsNullOrWhiteSpace(sourcePath) Then Return
            Dim extension As System.String = System.IO.Path.GetExtension(sourcePath)
            If Not System.String.Equals(extension, ".md", System.StringComparison.OrdinalIgnoreCase) AndAlso
               Not System.String.Equals(extension, ".txt", System.StringComparison.OrdinalIgnoreCase) Then
                SharedLibrary.SharedLibrary.SharedMethods.ShowCustomMessageBox("OCR cleanup accepts .md and .txt files. First extract PDF content with the Convert PDFs/Files helper.")
                Return
            End If
            sourceText = System.IO.File.ReadAllText(sourcePath)
        End If
        If System.String.IsNullOrWhiteSpace(sourceText) Then
            SharedLibrary.SharedLibrary.SharedMethods.ShowCustomMessageBox("The selected OCR text is empty.")
            Return
        End If
        Dim mode As SharedLibrary.SharedLibrary.SharedMethods.OcrMarkdownCleanupMode = SharedLibrary.SharedLibrary.SharedMethods.OcrMarkdownCleanupMode.None
        Dim options As SharedLibrary.SharedLibrary.SharedMethods.PdfMarkdownPreparationOptions = SelectPdfMarkdownPreparationOptions(mode)
        If options Is Nothing Then Return
        Dim cleaned As SharedLibrary.SharedLibrary.SharedMethods.OcrMarkdownCleanupResult
        If options.Enabled Then
            Dim pages As System.String() = sourceText.Split(New System.Char() {Microsoft.VisualBasic.ChrW(12)}, System.StringSplitOptions.None)
            options.ShowProgressWindow = True
            options.SourceName = If(sourcePath Is Nothing, "Selected OCR/Markdown text", System.IO.Path.GetFileName(sourcePath))
            Try
                cleaned = Await SharedLibrary.SharedLibrary.SharedMethods.PdfMarkdownPreparer.PrepareAsync(pages, options, _context)
            Catch ex As System.OperationCanceledException
                Return
            End Try
        ElseIf mode <> SharedLibrary.SharedLibrary.SharedMethods.OcrMarkdownCleanupMode.None Then
            cleaned = SharedLibrary.SharedLibrary.SharedMethods.OcrMarkdownCleaner.Clean(sourceText, mode)
        Else
            Return
        End If
        SharedLibrary.SharedLibrary.SharedMethods.PutInClipboard(cleaned.Report)
        Dim logNotice As System.String = "Detailed cleanup log copied to clipboard (Ctrl+V to paste)."
        If Not cleaned.ValidationPassed Then
            SharedLibrary.SharedLibrary.SharedMethods.ShowCustomMessageBox("Markdown preparation was skipped; the original source was retained." & System.Environment.NewLine & logNotice & System.Environment.NewLine & cleaned.Report)
            Return
        End If
        If sourcePath IsNot Nothing Then
            Dim outputPath As System.String = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(sourcePath),
                System.IO.Path.GetFileNameWithoutExtension(sourcePath) & " (bereinigt)" & System.IO.Path.GetExtension(sourcePath))
            Dim cleanBase As System.String = System.IO.Path.GetFileNameWithoutExtension(outputPath)
            Dim suffix As System.Int32 = 2
            While System.IO.File.Exists(outputPath)
                outputPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(sourcePath), cleanBase & " (" & suffix.ToString(System.Globalization.CultureInfo.InvariantCulture) & ")" & System.IO.Path.GetExtension(sourcePath))
                suffix += 1
            End While
            System.IO.File.WriteAllText(outputPath, cleaned.Content, New System.Text.UTF8Encoding(False))
            SharedLibrary.SharedLibrary.SharedMethods.ShowCustomMessageBox("Cleaned copy: " & outputPath & System.Environment.NewLine & System.Environment.NewLine & logNotice & System.Environment.NewLine & cleaned.Report)
        Else
            ' A separate document preserves the original range, native styles and source formatting.
            Dim document As Microsoft.Office.Interop.Word.Document = Me.Application.Documents.Add()
            document.Content.Text = cleaned.Content.Replace(Microsoft.VisualBasic.vbCrLf, Microsoft.VisualBasic.vbCr).Replace(Microsoft.VisualBasic.vbLf, Microsoft.VisualBasic.vbCr)
            SharedLibrary.SharedLibrary.SharedMethods.ShowCustomMessageBox("Cleaned OCR text was placed in a new document. The source selection was retained." & System.Environment.NewLine & System.Environment.NewLine & logNotice & System.Environment.NewLine & cleaned.Report)
        End If
        SharedLibrary.SharedLogger.Log(_context, _context.RDV, "Explicit OCR cleanup command completed; source retained, detailed audit presented.")
    End Function
End Class
