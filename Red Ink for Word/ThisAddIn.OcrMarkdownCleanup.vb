' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.OcrMarkdownCleanup.vb
' PURPOSE
'   Offers optional OCR-to-Markdown cleanup modes without overwriting original documents.
'
' KEY ARCHITECTURE
'   • UI: Selects optional cleanup levels for OCR-derived Markdown preparation.
'   • Safety: Processing never overwrites the original source document or file.
' =============================================================================

' Explicit OCR cleanup entry points. Source selection/files are never overwritten.
Option Explicit On
Option Strict On
Option Infer On

Partial Public Class ThisAddIn
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
