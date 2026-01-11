' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: SharedMethods.FileImporter.vb
' Purpose: Provides helper functions to read text from common document formats
'          (plain text, RTF, Word documents, and PDF), returning either extracted
'          text or an error string (depending on caller preference).
'
' Architecture:
'  - Text files: Normalizes the input path, validates existence, then reads UTF-8
'    (with BOM detection) via `StreamReader`.
'  - RTF: Loads the file contents and uses a hidden `RichTextBox` to convert RTF
'    markup to plain text.
'  - Word: Uses Office interop; attempts to attach to an existing Word instance,
'    otherwise creates an invisible instance; opens the document read-only and
'    returns `doc.Content.Text`.
'  - PDF: Uses UglyToad.PdfPig to iterate pages and extract text with multiple
'    fallback strategies; optionally runs OCR via an LLM call when heuristics
'    indicate that the PDF likely contains scanned images / poor text layer.
'
' External Dependencies:
'  - Microsoft.Office.Interop.Word (Word automation / COM interop)
'  - System.Windows.Forms.RichTextBox (RTF-to-text conversion)
'  - UglyToad.PdfPig (PDF parsing and text extraction)
'  - SharedLibrary.SharedContext.ISharedContext (OCR model/config access)
'  - Internal helpers used here: `ShowCustomYesNoBox`, `ShowCustomMessageBox`,
'    `LLM`, `GetSpecialTaskModel`, `RestoreDefaults`, and related configuration
'    fields (`originalConfigLoaded`, `originalConfig`).
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Word
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary
    Partial Public Class SharedMethods

        ''' <summary>
        ''' Result from reading a file, including content and metadata about potential incompleteness.
        ''' </summary>
        Public Class FileReadResult
            ''' <summary>The extracted text content from the file.</summary>
            Public Property Content As String = ""

            ''' <summary>True if the file was a PDF and heuristics suggested it may contain images but OCR was not performed.</summary>
            Public Property PdfMayBeIncomplete As Boolean = False

            Public Sub New()
            End Sub

            Public Sub New(content As String, pdfMayBeIncomplete As Boolean)
                Me.Content = content
                Me.PdfMayBeIncomplete = pdfMayBeIncomplete
            End Sub
        End Class

        ''' <summary>
        ''' Result from reading a PDF file, including content and metadata about OCR status.
        ''' </summary>
        Public Class PdfReadResult
            ''' <summary>The extracted text content from the PDF.</summary>
            Public Property Content As String = ""

            ''' <summary>True if heuristics suggested OCR but it was not performed (OCR unavailable or user declined).</summary>
            Public Property OcrWasSkippedDueToHeuristics As Boolean = False

            Public Sub New()
            End Sub

            Public Sub New(content As String, ocrSkipped As Boolean)
                Me.Content = content
                Me.OcrWasSkippedDueToHeuristics = ocrSkipped
            End Sub
        End Class

        ''' <summary>
        ''' Reads a text file as UTF-8 (with BOM detection) and returns its contents.
        ''' </summary>
        ''' <param name="filePath">Path to the file to read.</param>
        ''' <param name="ReturnErrorInsteadOfEmpty">
        ''' If <c>True</c>, returns an error message string on failure; otherwise returns an empty string.
        ''' </param>
        ''' <returns>The file contents, or an error string / empty string depending on <paramref name="ReturnErrorInsteadOfEmpty"/>.</returns>
        Public Shared Function ReadTextFile(filePath As String, Optional ReturnErrorInsteadOfEmpty As Boolean = True) As String
            Try
                ' Normalize and check the path
                filePath = Path.GetFullPath(filePath)
                If Not File.Exists(filePath) Then
                    Return If(ReturnErrorInsteadOfEmpty, "Error: File not found.", "")
                End If

                ' Use StreamReader for reading
                Using reader As New StreamReader(filePath, System.Text.Encoding.UTF8, True)
                    Dim content As String = reader.ReadToEnd()
                    Return content
                End Using
            Catch ex As System.Exception
                Return If(ReturnErrorInsteadOfEmpty, $"Error reading file: {ex.Message}", "")
            End Try
        End Function

        ''' <summary>
        ''' Reads an RTF file and returns its plain-text representation.
        ''' </summary>
        ''' <param name="rtfPath">Path to the RTF file to read.</param>
        ''' <param name="ReturnErrorInsteadOfEmpty">
        ''' If <c>True</c>, returns an error message string on failure; otherwise returns an empty string.
        ''' </param>
        ''' <returns>The extracted plain text, or an error string / empty string depending on <paramref name="ReturnErrorInsteadOfEmpty"/>.</returns>
        Public Shared Function ReadRtfAsText(ByVal rtfPath As String, Optional ReturnErrorInsteadOfEmpty As Boolean = True) As String
            Try
                Dim rtfContent As String = File.ReadAllText(rtfPath)
                Using rtb As New RichTextBox()
                    rtb.Visible = False
                    rtb.Rtf = rtfContent
                    Return rtb.Text
                End Using
            Catch ex As System.Exception
                Return If(ReturnErrorInsteadOfEmpty, $"Error reading RTF: {ex.Message}", "")
            End Try
        End Function

        ''' <summary>
        ''' Reads a Word document via Office interop and returns the document's text content.
        ''' </summary>
        ''' <param name="docPath">Path to the Word document to open.</param>
        ''' <param name="ReturnErrorInsteadOfEmpty">
        ''' If <c>True</c>, returns an error message string on failure; otherwise returns an empty string.
        ''' </param>
        ''' <returns>The extracted text, or an error string / empty string depending on <paramref name="ReturnErrorInsteadOfEmpty"/>.</returns>
        Public Shared Function ReadWordDocument(ByVal docPath As String, Optional ReturnErrorInsteadOfEmpty As Boolean = True) As String
            Dim app As Microsoft.Office.Interop.Word.Application = Nothing
            Dim doc As Document = Nothing
            Dim createdNewInstance As Boolean = False

            Try
                Try
                    ' Try to attach to an existing Word instance.
                    app = CType(Marshal.GetActiveObject("Word.Application"), Microsoft.Office.Interop.Word.Application)
                Catch ex As System.Exception
                    ' If Word is not running, create a new Word application.
                    app = New Microsoft.Office.Interop.Word.Application With {.Visible = False}
                    createdNewInstance = True
                End Try

                ' Open the Word document in read-only mode                
                Dim fileName As Object = docPath
                doc = app.Documents.Open(fileName, [ReadOnly]:=True, Visible:=False)

                ' Extract the content text
                Dim text As String = doc.Content.Text

                ' Close the document without saving changes
                doc.Close(SaveChanges:=False)

                ' Return the extracted text
                Return text

            Catch ex As System.Exception
                ' Ensure the document is closed in case of an error
                If doc IsNot Nothing Then
                    doc.Close(SaveChanges:=False)
                End If

                ' Return the error message (or empty string if ReturnErrorInsteadOfEmpty=False)
                Return If(ReturnErrorInsteadOfEmpty, $"Error reading Word document: {ex.Message}", "")

            Finally
                ' Only quit the application if it was newly created by this method
                If app IsNot Nothing AndAlso createdNewInstance Then
                    app.Quit()
                End If
            End Try
        End Function

        ''' <summary>
        ''' Reads a PDF using PdfPig and returns extracted text; optionally performs OCR via an LLM call
        ''' when heuristics indicate the PDF contains little or low-quality extractable text.
        ''' </summary>
        ''' <param name="pdfPath">Path to the PDF file to read.</param>
        ''' <param name="ReturnErrorInsteadOfEmpty">
        ''' If <c>True</c>, returns an error message string on failure; otherwise returns an empty string.
        ''' </param>
        ''' <param name="DoOCR">If <c>True</c>, enables OCR heuristics and (if confirmed) OCR execution.</param>
        ''' <param name="AskUser">If <c>True</c>, prompts the user before performing OCR.</param>
        ''' <param name="context">Shared context used for OCR-capable model configuration and LLM invocation.</param>
        ''' <returns>A PdfReadResult containing the extracted text and whether OCR was skipped despite being suggested.</returns>
        Public Shared Async Function ReadPdfAsTextEx(ByVal pdfPath As String,
                                            Optional ByVal ReturnErrorInsteadOfEmpty As Boolean = True,
                                            Optional ByVal DoOCR As Boolean = False,
                                            Optional ByVal AskUser As Boolean = True,
                                            Optional ByVal context As ISharedContext = Nothing) As Task(Of PdfReadResult)

            Dim result As New PdfReadResult()

            Try
                If String.IsNullOrWhiteSpace(pdfPath) OrElse Not IO.File.Exists(pdfPath) Then
                    result.Content = If(ReturnErrorInsteadOfEmpty, "Error: File not found or path is empty.", "")
                    Return result
                End If

                Dim sb As New System.Text.StringBuilder()
                Dim pageCount As Integer = 0
                Dim totalChars As Integer = 0
                Dim hasLowQualityText As Boolean = False
                Dim reasons As New List(Of String)()

                Using document As UglyToad.PdfPig.PdfDocument = UglyToad.PdfPig.PdfDocument.Open(pdfPath)
                    pageCount = document.NumberOfPages

                    For Each page As UglyToad.PdfPig.Content.Page In document.GetPages()
                        Dim pageText As String = page.Text
                        sb.AppendLine(pageText)
                        totalChars += If(pageText IsNot Nothing, pageText.Length, 0)

                        ' Check for low-quality text indicators
                        If pageText IsNot Nothing Then
                            Dim words = pageText.Split({" "c, vbCr(0), vbLf(0)}, StringSplitOptions.RemoveEmptyEntries)
                            Dim avgWordLen = If(words.Length > 0, words.Average(Function(w) w.Length), 0)
                            If avgWordLen < 2 AndAlso words.Length > 10 Then
                                hasLowQualityText = True
                            End If
                        End If
                    Next
                End Using

                Dim extractedText As String = sb.ToString().Trim()

                ' Heuristics to determine if OCR might be needed
                Dim shouldSuggestOcr As Boolean = False
                Dim avgCharsPerPage As Double = If(pageCount > 0, totalChars / pageCount, 0)

                If pageCount > 0 AndAlso avgCharsPerPage < 100 Then
                    shouldSuggestOcr = True
                    reasons.Add($"Very little text extracted ({avgCharsPerPage:F0} chars/page average)")
                End If

                If hasLowQualityText Then
                    shouldSuggestOcr = True
                    reasons.Add("Text appears to be low quality (possibly garbled OCR or image-based)")
                End If

                If String.IsNullOrWhiteSpace(extractedText) AndAlso pageCount > 0 Then
                    shouldSuggestOcr = True
                    reasons.Add("No text could be extracted from any page")
                End If

                ' Disable OCR if no OCR-capable call is configured or context missing
                If DoOCR AndAlso (context Is Nothing OrElse Not IsOcrAvailable(context)) Then
                    DoOCR = False
                End If

                ' If DoOCR is disabled → just return whatever text we found (or empty string)
                If Not DoOCR Then
                    ' If we would have suggested OCR but it's not available, flag it
                    If shouldSuggestOcr Then
                        result.OcrWasSkippedDueToHeuristics = True
                    End If
                    result.Content = extractedText
                    Return result
                End If

                If shouldSuggestOcr Then
                    ' Check if OCR is actually available
                    If Not IsOcrAvailable(context) Then
                        ' OCR would be suggested but is not available - silently continue with extracted text
                        Debug.WriteLine("OCR suggested by heuristics but not available - skipping OCR prompt.")
                        result.OcrWasSkippedDueToHeuristics = True
                        result.Content = extractedText
                        Return result
                    End If

                    If AskUser Then
                        Dim formattedReasons As String = String.Join(Environment.NewLine, reasons.ConvertAll(Function(r) "- " & r))
                        Dim msg As String = $"The PDF appears to contain little or no extractable text:" & Environment.NewLine & Environment.NewLine &
                                            formattedReasons & Environment.NewLine & Environment.NewLine &
                                            "It's likely that the document consists mainly of scanned images." & Environment.NewLine & Environment.NewLine &
                                            "Would you like AI to perform OCR to extract text (if supported by your configured model)?"
                        Dim userChoice As Integer = ShowCustomYesNoBox(msg, "Yes, try OCR", "No, use what you have")
                        If userChoice <> 1 Then
                            result.OcrWasSkippedDueToHeuristics = True
                            result.Content = extractedText
                            Return result
                        End If
                    End If

                    Dim ocrText As String = Await PerformOCR(pdfPath, context)
                    If Not String.IsNullOrWhiteSpace(ocrText) Then
                        result.Content = ocrText
                        Return result
                    Else
                        ' OCR was attempted but returned empty - content may be incomplete
                        result.OcrWasSkippedDueToHeuristics = True
                    End If
                End If

                result.Content = extractedText
                Return result

            Catch ex As System.Exception
                result.Content = If(ReturnErrorInsteadOfEmpty, $"Error reading PDF: {ex.Message}", "")
                Return result
            End Try
        End Function

        ''' <summary>
        ''' Reads a PDF using PdfPig and returns extracted text (backward compatible wrapper).
        ''' </summary>
        Public Shared Async Function ReadPdfAsText(ByVal pdfPath As String,
                                            Optional ByVal ReturnErrorInsteadOfEmpty As Boolean = True,
                                            Optional ByVal DoOCR As Boolean = False,
                                            Optional ByVal AskUser As Boolean = True,
                                            Optional ByVal context As ISharedContext = Nothing) As Task(Of String)
            Dim result = Await ReadPdfAsTextEx(pdfPath, ReturnErrorInsteadOfEmpty, DoOCR, AskUser, context)
            Return result.Content
        End Function

        ''' <summary>
        ''' Extracts plain text content from a single PDF page using multiple strategies:
        ''' content-order extraction, word/line reconstruction, and finally a letter-gap heuristic.
        ''' </summary>
        ''' <param name="page">The PDF page to extract text from.</param>
        ''' <returns>Extracted page text (may be empty).</returns>
        Private Shared Function ExtractPageTextFromPdf(page As UglyToad.PdfPig.Content.Page) As String
            ' 1) Try PdfPig’s content-order extractor (good spacing/reading order on many PDFs)
            Try
                Dim t As String = UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor.ContentOrderTextExtractor.GetText(page)
                If Not String.IsNullOrWhiteSpace(t) AndAlso (t.Contains(" ") OrElse t.Contains(vbTab) OrElse t.Contains(vbCr) OrElse t.Contains(vbLf)) Then
                    Return t
                End If
            Catch
                ' Older PdfPig versions or certain pages may not support this path; ignore and fallback.
            End Try

            ' 2) Word-based reconstruction using Nearest-Neighbour extractor (higher recall on tricky PDFs)
            Try
                Dim words As System.Collections.Generic.IEnumerable(Of UglyToad.PdfPig.Content.Word) =
            page.GetWords(UglyToad.PdfPig.DocumentLayoutAnalysis.WordExtractor.NearestNeighbourWordExtractor.Instance)

                If words IsNot Nothing AndAlso words.Count > 0 Then
                    ' Group words into lines by baseline with a tolerant threshold
                    Dim baselineTol As Double = Math.Max(0.5, page.Height * 0.002) ' ~0.2% of page height
                    Dim lines As New System.Collections.Generic.List(Of System.Collections.Generic.List(Of UglyToad.PdfPig.Content.Word))()

                    For Each w In words.OrderByDescending(Function(x) x.BoundingBox.Bottom).ThenBy(Function(x) x.BoundingBox.Left)
                        Dim placed As Boolean = False
                        For Each ln In lines
                            Dim ref = ln(0)
                            If Math.Abs(w.BoundingBox.Bottom - ref.BoundingBox.Bottom) <= baselineTol Then
                                ln.Add(w)
                                placed = True
                                Exit For
                            End If
                        Next
                        If Not placed Then
                            lines.Add(New System.Collections.Generic.List(Of UglyToad.PdfPig.Content.Word) From {w})
                        End If
                    Next

                    Dim sbLine As New System.Text.StringBuilder()
                    Dim first As Boolean = True
                    For Each ln In lines.OrderByDescending(Function(l) l.Average(Function(w) w.BoundingBox.Bottom))
                        If Not first Then sbLine.AppendLine()
                        first = False
                        Dim lineText = String.Join(" ", ln.OrderBy(Function(w) w.BoundingBox.Left).Select(Function(w) w.Text))
                        sbLine.Append(lineText)
                    Next

                    Dim s = sbLine.ToString()
                    If Not String.IsNullOrWhiteSpace(s) Then
                        Return s
                    End If
                End If
            Catch
                ' Ignore and fallback
            End Try

            ' 3) Letter-gap heuristic: insert spaces based on horizontal gaps; break lines on baseline changes
            Dim letters = page.Letters
            If letters Is Nothing OrElse letters.Count = 0 Then Return String.Empty

            Dim ordered = letters.OrderByDescending(Function(l) l.GlyphRectangle.Bottom).ThenBy(Function(l) l.GlyphRectangle.Left)
            Dim sb As New System.Text.StringBuilder()
            Dim prev As UglyToad.PdfPig.Content.Letter = Nothing

            For Each l In ordered
                If prev IsNot Nothing Then
                    Dim sameLine = Math.Abs(l.GlyphRectangle.Bottom - prev.GlyphRectangle.Bottom) <= Math.Max(0.5, prev.GlyphRectangle.Height * 0.6)
                    If Not sameLine Then
                        sb.AppendLine()
                    Else
                        Dim gap = l.GlyphRectangle.Left - prev.GlyphRectangle.Right
                        Dim spaceThreshold = Math.Max(prev.GlyphRectangle.Width * 0.6, 0.5) ' tune if needed
                        If gap > spaceThreshold Then sb.Append(" ")
                    End If
                End If
                sb.Append(l.Value)
                prev = l
            Next

            Return sb.ToString()
        End Function

        ''' <summary>
        ''' Performs OCR on a PDF by invoking an LLM call with the PDF path as a binary object input.
        ''' </summary>
        ''' <param name="pdfPath">Path to the PDF file to OCR.</param>
        ''' <param name="context">Shared context containing model and API configuration.</param>
        ''' <returns>OCR result text, or an empty string if OCR is not available or fails.</returns>
        Private Shared Async Function PerformOCR(ByVal pdfPath As String, context As ISharedContext) As Task(Of String)

            ' Use the comprehensive OCR availability check
            If Not IsOcrAvailable(context) Then
                ShowCustomMessageBox($"OCR is not available with your current model configuration.")
                Return ""
            End If

            Dim UseSecondAPI As Boolean = False
            Dim TimeOut = context.INI_Timeout

            If Not String.IsNullOrWhiteSpace(context.INI_AlternateModelPath) Then
                If Not GetSpecialTaskModel(context, context.INI_AlternateModelPath, "OCR") Then
                    originalConfigLoaded = False
                    UseSecondAPI = False
                Else
                    UseSecondAPI = True
                    TimeOut = context.INI_Timeout_2
                End If
            End If

            Dim result As System.String = Await LLM(context, context.SP_InsertClipboard, "", "", "", TimeOut * 2, UseSecondAPI, False, "", pdfPath)

            ' Restore model if temporarily switched
            If UseSecondAPI AndAlso originalConfigLoaded Then
                RestoreDefaults(context, originalConfig)
                originalConfigLoaded = False
            End If

            Return result

        End Function

        ''' <summary>
        ''' Determines whether OCR is available based on the configured model capabilities.
        ''' </summary>
        ''' <param name="context">Shared context containing model and API configuration.</param>
        ''' <returns>True if OCR is available, False otherwise.</returns>
        Public Shared Function IsOcrAvailable(context As ISharedContext) As Boolean
            If context Is Nothing Then Return False

            ' First check: If alternate model path is configured, check for OCR-capable secondary model
            If Not String.IsNullOrWhiteSpace(context.INI_AlternateModelPath) Then
                ' Save original config before GetSpecialTaskModel potentially changes it
                Dim savedConfig As ModelConfig = GetCurrentConfig(context)
                Dim savedConfigLoaded As Boolean = originalConfigLoaded

                Try
                    If GetSpecialTaskModel(context, context.INI_AlternateModelPath, "OCR") Then
                        ' OCR model found - restore config and return True
                        RestoreDefaults(context, savedConfig)
                        originalConfigLoaded = savedConfigLoaded
                        Return True
                    End If
                Catch
                    ' If GetSpecialTaskModel fails, continue to check primary API
                Finally
                    ' Ensure config is restored even if exception occurred
                    RestoreDefaults(context, savedConfig)
                    originalConfigLoaded = savedConfigLoaded
                End Try
            End If

            ' Second check: Evaluate the INI_APICall_Object configuration
            Return IsApiCallObjectOcrCapable(context.INI_APICall_Object)
        End Function

        ''' <summary>
        ''' Checks if the given APICall_Object configuration string supports PDF/OCR.
        ''' </summary>
        ''' <param name="apiCallObject">The INI_APICall_Object or INI_APICall_Object_2 string.</param>
        ''' <returns>True if OCR/PDF is supported, False otherwise.</returns>
        Private Shared Function IsApiCallObjectOcrCapable(apiCallObject As String) As Boolean
            ' If null or empty, OCR is not available
            If String.IsNullOrWhiteSpace(apiCallObject) Then
                Return False
            End If

            ' Check if the string contains segment separators (¦)
            Dim segments As String() = apiCallObject.Split(New Char() {"¦"c}, StringSplitOptions.RemoveEmptyEntries)

            ' Track if we found any segment without a filter (means all types supported)
            ' or any segment with a filter that includes PDF
            Dim hasUnfilteredSegment As Boolean = False
            Dim hasPdfFilter As Boolean = False
            Dim allSegmentsHaveFilters As Boolean = True

            For Each segment As String In segments
                Dim trimmedSegment As String = segment.Trim()

                ' Check if this segment has a filter (starts with [...])
                If trimmedSegment.StartsWith("[") Then
                    ' Extract the filter content between [ and ]
                    Dim closeBracketIdx As Integer = trimmedSegment.IndexOf("]"c)
                    If closeBracketIdx > 1 Then
                        Dim filterContent As String = trimmedSegment.Substring(1, closeBracketIdx - 1)

                        ' Check if the filter contains application/pdf or pdf
                        If filterContent.IndexOf("application/pdf", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                           filterContent.IndexOf("pdf", StringComparison.OrdinalIgnoreCase) >= 0 Then
                            hasPdfFilter = True
                        End If
                    End If
                Else
                    ' No filter on this segment - means it accepts all types
                    hasUnfilteredSegment = True
                    allSegmentsHaveFilters = False
                End If
            Next

            ' OCR is available if:
            ' 1. There's at least one segment without a filter (accepts all), OR
            ' 2. There's a segment with a filter that includes PDF
            If hasUnfilteredSegment Then
                Return True
            End If

            If hasPdfFilter Then
                Return True
            End If

            ' If all segments have filters and none include PDF, OCR is not available
            If allSegmentsHaveFilters Then
                Return False
            End If

            ' Default: if we have content but couldn't parse filters, assume capable
            Return True
        End Function


    End Class

End Namespace