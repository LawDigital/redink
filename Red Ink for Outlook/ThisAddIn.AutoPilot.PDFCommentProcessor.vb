' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved.
' For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.AutoPilot.PDFCommentProcessor.vb
' Purpose:
'   Adds review comments to PDF attachments by creating native PDF annotations
'   (highlight + popup), driven by LLM-generated quote/comment pairs.
'
' Architecture:
'  - Extraction:
'      * Reads PDF text and per-character coordinates using PdfPig.
'      * Builds a searchable text stream plus positional map for precise anchoring.
'  - LLM protocol:
'      * Uses the same bubble protocol as AutoPilot comment processing
'        (`text@@comment§§§...`) and reuses `APParseBubblesResponse`.
'      * Supports single-pass and batched mode for large PDFs with context windows.
'  - Matching:
'      * Resolves quoted text spans with multi-tier matching:
'        exact, case-insensitive, mapped whitespace and physical line-end dehyphenation.
'      * Requires a complete, unique quote within its original review batch.
'      * Converts verified spans into contiguous word fragments by page; no prefix guessing.
'  - Injection:
'      * Writes annotations with PdfSharp low-level dictionaries:
'          - `/Highlight` + `/Popup` for matched quotes
'          - `/Text` labelled sticky notes for unmatched/ambiguous comments
'      * Adds author metadata, collision-safe annotation names, timestamps and color.
'      * Reports anchor counts and full quote/source/geometry diagnostics to the tool log.
'  - Output behavior:
'      * Works on a copied output PDF, preserves input unchanged.
'      * Saves modified annotations directly into the destination file.
'
' Dependencies:
'  - PdfPig (`UglyToad.PdfPig`) for text/position extraction.
'  - PdfSharp for annotation object creation and PDF write-back.
'  - AutoPilot comment pipeline helpers from:
'      * `ThisAddIn.AutoPilot.CommentProcessor.vb`
'        (`APParseBubblesResponse` and prompt constants).
'
' Notes:
'  - Unmatched comments are placed as stacked sticky notes on page 1 (top-right),
'    with overlap avoidance and fallback column wrapping.
'  - Optional custom author names are supported; non-default author mode prefixes
'    comment text with the Red Ink marker.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports PdfSharp.Pdf
Imports PdfSharp.Pdf.IO
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports PDFP = UglyToad.PdfPig

Partial Public Class ThisAddIn

    ' ═══════════════════════════════════════════════════════════════════════════
    '  PDF COMMENT PROCESSOR — DATA CLASSES
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Represents a single character's position within the PDF for text-matching.
    ''' </summary>
    Private Class PdfCharPosition
        Public Property PageIndex As Integer       ' 0-based page index
        Public Property WordIndex As System.Int32 = -1 ' Extraction identity; prevents quads bridging columns
        Public Property X As Double                ' Left edge in PDF points
        Public Property Y As Double                ' Bottom edge in PDF points (PDF coordinate system)
        Public Property Width As Double            ' Character width in points
        Public Property Height As Double           ' Character height in points
        Public Property PageWidth As Double        ' Page width for fallback positioning
        Public Property PageHeight As Double       ' Page height for coordinate transforms
    End Class

    ''' <summary>
    ''' Represents one contiguous word/line fragment on a specific page.
    ''' </summary>
    Private Class PdfHighlightRect
        Public Property PageIndex As Integer       ' 0-based
        Public Property X As Double                ' Left
        Public Property Y As Double                ' Bottom (PDF coords)
        Public Property Width As Double
        Public Property Height As Double
    End Class

    ''' <summary>One model comment together with its immutable review window.</summary>
    Private Class PdfCommentEntry
        Inherits APCommentEntry

        Public Property SourceStart As System.Int32
        Public Property SourceLength As System.Int32
        Public Property BatchNumber As System.Int32
        Public Property CommentId As System.Int32
        Public Property Anchor As PdfTextMatch
    End Class

    ''' <summary>A complete source span, or an explicit reason why it is not safe to anchor.</summary>
    Private Class PdfTextMatch
        Public Property Start As System.Int32 = -1
        Public Property Length As System.Int32
        Public Property Method As System.String = System.String.Empty
        Public Property CandidateCount As System.Int32
        Public Property FailureReason As System.String = "quote_not_found"
    End Class

    ''' <summary>Counts comments, not the number of per-page annotation objects.</summary>
    Private Class PdfCommentProcessingReport
        Public Property TotalCount As System.Int32
        Public Property HighlightedCount As System.Int32
        Public Property UnanchoredCount As System.Int32
        Public Property BatchCount As System.Int32
        Public Property UnconfirmedBatchCount As System.Int32
    End Class

    ' ═══════════════════════════════════════════════════════════════════════════
    '  PDF COMMENT PROCESSOR — CONSTANTS
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Maximum characters intended for a single LLM batch when processing PDFs.
    ''' </summary>
    Private Const AP_PdfCommentMaxCharsPerBatch As Integer = 15000

    ''' <summary>
    ''' Vertical spacing between stacked unmatched sticky notes (in points).
    ''' </summary>
    Private Const AP_PdfStickyNoteSpacing As Double = 30.0

    ''' <summary>
    ''' Size of the sticky note icon rectangle (in points).
    ''' </summary>
    Private Const AP_PdfStickyNoteSize As Double = 24.0

    ''' <summary>
    ''' Right margin offset for sticky note placement (in points from right edge).
    ''' </summary>
    Private Const AP_PdfStickyNoteRightMargin As Double = 30.0

    ''' <summary>
    ''' Top margin offset for the first sticky note (in points from top edge).
    ''' </summary>
    Private Const AP_PdfStickyNoteTopMargin As Double = 30.0

    ' ═══════════════════════════════════════════════════════════════════════════
    '  PDF COMMENT PROCESSOR — MAIN ENTRY POINT
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Processes a PDF file by having the LLM add comment annotations to relevant sections
    ''' based on the given instruction.
    ''' </summary>
    ''' <param name="author">Optional author name for the comments. If Nothing or empty, defaults to AN6 ("Inky").
    ''' When a custom author is used, each comment is prefixed with "RI: " (AN5).</param>
    Private Async Function CommentPdfForAutoPilot(inputPath As System.String, outputPath As System.String,
                                                   instruction As System.String, ct As System.Threading.CancellationToken,
                                                   Optional author As System.String = Nothing,
                                                   Optional processingReport As PdfCommentProcessingReport = Nothing,
                                                   Optional diagnosticLog As System.Action(Of System.String) = Nothing) As System.Threading.Tasks.Task(Of System.Boolean)
        Dim effectiveAuthor As System.String = If(System.String.IsNullOrWhiteSpace(author), AN6, author.Trim())
        Dim usePrefix As System.Boolean = Not effectiveAuthor.Equals(AN6, System.StringComparison.OrdinalIgnoreCase)
        Dim report As PdfCommentProcessingReport = If(processingReport, New PdfCommentProcessingReport())
        report.TotalCount = 0
        report.HighlightedCount = 0
        report.UnanchoredCount = 0
        report.BatchCount = 0
        report.UnconfirmedBatchCount = 0

        Try
            ApDashboardLog("PdfCommentProcessor: starting PDF comment processing", "step")

            Dim fullText As System.String = Nothing
            Dim charPositions As System.Collections.Generic.List(Of PdfCharPosition) = Nothing
            APPdfExtractTextAndPositions(inputPath, fullText, charPositions)

            If System.String.IsNullOrWhiteSpace(fullText) OrElse charPositions Is Nothing OrElse charPositions.Count = 0 Then
                ApDashboardLog("PdfCommentProcessor: PDF contains no extractable text", "warn")
                Return False
            End If

            ApDashboardLog($"PdfCommentProcessor: extracted {fullText.Length} characters ({charPositions.Count} positioned chars)", "step")

            ' Use the same PDF anchor contract and source-window metadata for one or many batches.
            Dim comments As System.Collections.Generic.List(Of PdfCommentEntry) =
                Await APPdfGetCommentsFromLLM(fullText, instruction, ct, report, diagnosticLog)
            APPdfWriteDiagnostic(diagnosticLog, "PdfCommentReviewBatches=" &
                Newtonsoft.Json.JsonConvert.SerializeObject(New With {
                    .batches = report.BatchCount,
                    .unconfirmed_review_batches = report.UnconfirmedBatchCount
                }))

            If comments Is Nothing OrElse comments.Count = 0 Then
                ApDashboardLog("PdfCommentProcessor: LLM returned no comments", "warn")
                Return False
            End If

            report.TotalCount = comments.Count
            ApDashboardLog($"PdfCommentProcessor: {comments.Count} comments to insert", "step")
            EnsureApPdfSharpFontResolver()
            System.IO.File.Copy(inputPath, outputPath, overwrite:=True)

            Using pdfDoc As PdfSharp.Pdf.PdfDocument = PdfSharp.Pdf.IO.PdfReader.Open(outputPath, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Modify)
                Dim commentPrefix As System.String = If(usePrefix, AN5 & ": ", System.String.Empty)
                Dim unmatchedEntries As New System.Collections.Generic.List(Of PdfCommentEntry)()
                Dim commentId As System.Int32 = 0

                For Each entry As PdfCommentEntry In comments
                    ct.ThrowIfCancellationRequested()
                    commentId += 1
                    entry.CommentId = commentId

                    Dim match As PdfTextMatch = Nothing
                    Dim highlightRects As System.Collections.Generic.List(Of PdfHighlightRect) =
                        APPdfFindTextRects(fullText, charPositions, entry, pdfDoc.Pages.Count, match)
                    entry.Anchor = match
                    APPdfLogAnchorDiagnostic(diagnosticLog, commentId, entry, match, highlightRects, fullText)

                    If highlightRects Is Nothing OrElse highlightRects.Count = 0 Then
                        unmatchedEntries.Add(entry)
                        Continue For
                    End If

                    Dim commentBody As System.String = commentPrefix & entry.CommentText
                    APPdfInsertHighlightAnnotation(pdfDoc, highlightRects, commentBody, effectiveAuthor, commentId)
                    report.HighlightedCount += 1
                Next

                ' Preserve every unresolved comment as an explicitly unanchored note.
                If unmatchedEntries.Count > 0 Then
                    Dim firstPage As PdfSharp.Pdf.PdfPage = pdfDoc.Pages(0)
                    Dim pageW As System.Double = firstPage.Width.Point
                    Dim pageH As System.Double = firstPage.Height.Point
                    Dim rowsPerColumn As System.Int32 = System.Math.Max(1,
                        CInt(System.Math.Floor((pageH - AP_PdfStickyNoteTopMargin - AP_PdfStickyNoteSize) / AP_PdfStickyNoteSpacing)) + 1)

                    For i As System.Int32 = 0 To unmatchedEntries.Count - 1
                        ct.ThrowIfCancellationRequested()
                        Dim noteX As System.Double = System.Math.Max(0,
                            pageW - AP_PdfStickyNoteRightMargin - AP_PdfStickyNoteSize -
                            (i \ rowsPerColumn) * (AP_PdfStickyNoteSize + 10))
                        Dim noteY As System.Double = System.Math.Max(0,
                            pageH - AP_PdfStickyNoteTopMargin - (i Mod rowsPerColumn) * AP_PdfStickyNoteSpacing)
                        Dim entry As PdfCommentEntry = unmatchedEntries(i)
                        Dim noteBody As System.String = commentPrefix & entry.CommentText &
                            vbCrLf & vbCrLf & "[Unanchored comment; reason: " & entry.Anchor.FailureReason &
                            "; referenced text: """ & entry.QuotedText & """]"

                        APPdfInsertStickyNoteAnnotation(pdfDoc, 0, noteX, noteY, noteBody, effectiveAuthor, entry.CommentId)
                        report.UnanchoredCount += 1
                    Next

                    ApDashboardLog($"PdfCommentProcessor: {report.UnanchoredCount} comments could not be anchored safely; retained as labelled sticky notes", "warn")
                End If

                If report.HighlightedCount + report.UnanchoredCount <> report.TotalCount Then
                    Throw New System.InvalidOperationException("PDF comment accounting did not preserve every parsed comment.")
                End If

                pdfDoc.Save(outputPath)
            End Using

            APPdfWriteDiagnostic(diagnosticLog, "PdfCommentSummary=" &
                Newtonsoft.Json.JsonConvert.SerializeObject(New With {
                    .stage = "saved",
                    .total = report.TotalCount,
                    .anchored = report.HighlightedCount,
                    .unanchored = report.UnanchoredCount,
                    .anchoring_complete = (report.UnanchoredCount = 0),
                    .review_batches = report.BatchCount,
                    .unconfirmed_review_batches = report.UnconfirmedBatchCount
                }))
            ApDashboardLog($"PdfCommentProcessor: saved {report.HighlightedCount} anchored comments and {report.UnanchoredCount} unanchored notes", "success")
            Return True

        Catch ex As System.Exception
            System.Diagnostics.Debug.WriteLine("CommentPdfForAutoPilot error: " & ex.ToString())
            APPdfWriteDiagnostic(diagnosticLog, "PdfCommentFailure=" &
                Newtonsoft.Json.JsonConvert.SerializeObject(New With {.stage = "failed", .error = ex.ToString()}))
            ApDashboardLog("PdfCommentProcessor: error - " & ex.Message, "error")
            Return False
        End Try
    End Function
    ' ═══════════════════════════════════════════════════════════════════════════
    '  TEXT EXTRACTION WITH POSITIONS (via PdfPig)
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Extracts the full concatenated text from a PDF and builds a parallel list
    ''' of character positions (page, x, y, width, height) for text-matching.
    ''' Uses PdfPig's word-level extraction to correctly insert spaces between words,
    ''' since PDF content streams do not contain explicit space characters — spaces
    ''' are implied by glyph positioning gaps.
    ''' </summary>
    Private Shared Sub APPdfExtractTextAndPositions(pdfPath As String,
                                                    ByRef fullText As String,
                                                    ByRef positions As List(Of PdfCharPosition))
        Dim sb As New StringBuilder()
        positions = New List(Of PdfCharPosition)()
        Dim wordIndex As System.Int32 = 0

        Using document As PDFP.PdfDocument = PDFP.PdfDocument.Open(pdfPath)
            For pageIdx = 0 To document.NumberOfPages - 1
                Dim page = document.GetPage(pageIdx + 1) ' PdfPig is 1-based
                Dim pageWidth = page.Width
                Dim pageHeight = page.Height

                ' Use PdfPig's word extraction to get proper word boundaries.
                ' Each Word contains its constituent Letters with positions,
                ' and we insert a space between consecutive words.
                Dim words = page.GetWords().ToList()
                Dim isFirstWordOnPage As Boolean = True

                For Each word As PDFP.Content.Word In words
                    wordIndex += 1
                    ' Insert space before this word (unless it's the first on the page
                    ' or the first word after a page separator)
                    If Not isFirstWordOnPage Then
                        ' Detect line breaks: if this word's Y position differs significantly
                        ' from what came before, insert a newline instead of a space
                        Dim lastPos = If(positions.Count > 0, positions(positions.Count - 1), Nothing)
                        Dim isNewLine As Boolean = False

                        If lastPos IsNot Nothing AndAlso lastPos.PageIndex = pageIdx Then
                            Dim yDiff = Math.Abs(word.BoundingBox.Bottom - lastPos.Y)
                            Dim lineHeight = If(lastPos.Height > 0, lastPos.Height, 12.0)
                            isNewLine = (yDiff > lineHeight * 0.5)
                        End If

                        Dim separator = If(isNewLine, vbLf, " ")
                        sb.Append(separator)
                        ' Add dummy position entry for the separator character
                        positions.Add(New PdfCharPosition() With {
                            .PageIndex = pageIdx,
                            .X = 0, .Y = 0, .Width = 0, .Height = 0,
                            .PageWidth = pageWidth,
                            .PageHeight = pageHeight
                        })
                    End If
                    isFirstWordOnPage = False

                    ' Append each letter of the word with its position
                    For Each letter As PDFP.Content.Letter In word.Letters
                        Dim ch = letter.Value
                        If String.IsNullOrEmpty(ch) Then Continue For

                        For Each c In ch
                            sb.Append(c)
                            positions.Add(New PdfCharPosition() With {
                                .PageIndex = pageIdx,
                                .WordIndex = wordIndex,
                                .X = letter.GlyphRectangle.Left,
                                .Y = letter.GlyphRectangle.Bottom,
                                .Width = letter.GlyphRectangle.Width,
                                .Height = letter.GlyphRectangle.Height,
                                .PageWidth = pageWidth,
                                .PageHeight = pageHeight
                            })
                        Next
                    Next
                Next

                ' Add a space between pages to prevent cross-page word merging
                If pageIdx < document.NumberOfPages - 1 Then
                    sb.Append(" ")
                    positions.Add(New PdfCharPosition() With {
                        .PageIndex = pageIdx,
                        .X = 0, .Y = 0, .Width = 0, .Height = 0,
                        .PageWidth = pageWidth,
                        .PageHeight = pageHeight
                    })
                End If
            Next
        End Using

        fullText = sb.ToString()
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    '  LLM CALL — BATCHED (for large PDFs without paragraph structure)
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Processes the PDF text in batches, sending each batch to the LLM with
    ''' surrounding context so the LLM understands the flow.
    ''' </summary>
    Private Async Function APPdfGetCommentsFromLLM(fullText As System.String,
                                                   instruction As System.String,
                                                   ct As System.Threading.CancellationToken,
                                                   report As PdfCommentProcessingReport,
                                                   diagnosticLog As System.Action(Of System.String)) As System.Threading.Tasks.Task(Of System.Collections.Generic.List(Of PdfCommentEntry))
        Dim allComments As New System.Collections.Generic.List(Of PdfCommentEntry)()

        ' Build the system prompt once
        Dim bubblesPrompt As String = InterpolateAtRuntime(SP_Add_Bubbles)
        bubblesPrompt = bubblesPrompt.Replace("{FormatInstruction}",
            "Provide each comment as plain text without any markdown formatting.")

        Dim systemPrompt As String =
            "You are a professional document reviewer. " &
            "Apply the following instruction to the document text provided." & vbCrLf & vbCrLf &
            "INSTRUCTION: " & instruction & vbCrLf & vbCrLf &
            bubblesPrompt & vbCrLf & vbCrLf &
            "IMPORTANT: The text between [TEXTTOPROCESS] and [/TEXTTOPROCESS] is the section you must review and comment on. " &
            "Text between [CONTEXT_BEFORE] and [/CONTEXT_BEFORE] or [CONTEXT_AFTER] and [/CONTEXT_AFTER] is provided " &
            "for context only — do NOT comment on context sections." & vbCrLf & vbCrLf &
            "PDF ANCHOR CONTRACT (takes precedence over any conflicting generic quote/line-break rule above): " &
            "For each comment, quote a complete, contiguous passage that actually states the provision or fact discussed in that comment. " &
            "Copy it exactly from TEXTTOPROCESS, including digits, punctuation, hyphens and physical line breaks. " &
            "Multiline quotes are allowed and required when the relevant passage wraps across PDF lines. " &
            "Do not substitute a nearby heading or unrelated short sentence just because it fits on one line. " &
            "Do not abbreviate with ellipses or join disjoint passages. If a quote occurs more than once in TEXTTOPROCESS, " &
            "include immediately adjacent source text until it is unambiguous. Quote only the current TEXTTOPROCESS window, never the context blocks. " &
            "Keep the existing text@@comment and comment-separator protocol."

        Dim contextSize As Integer = 500 ' Characters of context before/after each batch
        Dim totalLen = fullText.Length
        Dim estimatedBatches As Integer = CInt(Math.Ceiling(totalLen / CDbl(AP_PdfCommentMaxCharsPerBatch)))

        ApDashboardLog($"PdfCommentProcessor: {totalLen} chars to process (~{estimatedBatches} batches)", "step")

        Dim batchStart As Integer = 0
        Dim currentBatch As Integer = 0

        While batchStart < totalLen
            ct.ThrowIfCancellationRequested()
            currentBatch += 1
            report.BatchCount += 1

            ' Determine batch end
            Dim batchEnd = Math.Min(batchStart + AP_PdfCommentMaxCharsPerBatch, totalLen)

            ' Try to break at a word boundary
            If batchEnd < totalLen Then
                Dim lastSpace = fullText.LastIndexOf(" "c, batchEnd, Math.Min(200, batchEnd - batchStart))
                If lastSpace > batchStart Then batchEnd = lastSpace
            End If

            Dim charCount = batchEnd - batchStart
            ApDashboardLog($"PdfCommentProcessor: batch {currentBatch}/{estimatedBatches} (chars {batchStart + 1}-{batchEnd}, {charCount} chars)", "step")

            ' Build context before
            Dim ctxBeforeStart = Math.Max(0, batchStart - contextSize)
            Dim contextBefore = If(ctxBeforeStart < batchStart, fullText.Substring(ctxBeforeStart, batchStart - ctxBeforeStart), "")

            ' Main text
            Dim mainText = fullText.Substring(batchStart, charCount)

            ' Build context after
            Dim ctxAfterEnd = Math.Min(totalLen, batchEnd + contextSize)
            Dim contextAfter = If(batchEnd < ctxAfterEnd, fullText.Substring(batchEnd, ctxAfterEnd - batchEnd), "")

            ' Assemble user prompt with context markers
            Dim userPrompt As New StringBuilder()
            If contextBefore.Length > 0 Then
                userPrompt.AppendLine("[CONTEXT_BEFORE]")
                userPrompt.AppendLine(contextBefore)
                userPrompt.AppendLine("[/CONTEXT_BEFORE]")
                userPrompt.AppendLine()
            End If

            userPrompt.AppendLine("[TEXTTOPROCESS]")
            userPrompt.AppendLine(mainText)
            userPrompt.AppendLine("[/TEXTTOPROCESS]")

            If contextAfter.Length > 0 Then
                userPrompt.AppendLine()
                userPrompt.AppendLine("[CONTEXT_AFTER]")
                userPrompt.AppendLine(contextAfter)
                userPrompt.AppendLine("[/CONTEXT_AFTER]")
            End If

            ' Call LLM for this batch
            Dim llmResponse = Await LLM(systemPrompt, userPrompt.ToString(),
                                         UseSecondAPI:=False,
                                         HideSplash:=True, EnsureUI:=False,
                                         cancellationToken:=ct)

            Dim returnedCommentCount As System.Int32 = 0
            If Not String.IsNullOrWhiteSpace(llmResponse) Then
                Dim batchComments = APParseBubblesResponse(llmResponse)
                If batchComments IsNot Nothing AndAlso batchComments.Count > 0 Then
                    returnedCommentCount = batchComments.Count
                    For Each comment As APCommentEntry In batchComments
                        allComments.Add(New PdfCommentEntry() With {
                            .QuotedText = comment.QuotedText,
                            .CommentText = comment.CommentText,
                            .SourceStart = batchStart,
                            .SourceLength = charCount,
                            .BatchNumber = currentBatch
                        })
                    Next
                    ApDashboardLog($"PdfCommentProcessor: batch {currentBatch} returned {batchComments.Count} comments", "step")
                Else
                    report.UnconfirmedBatchCount += 1
                    ApDashboardLog($"PdfCommentProcessor: batch {currentBatch} returned no usable comment response; review coverage is unconfirmed (this may also mean no findings)", "warn")
                End If
            Else
                report.UnconfirmedBatchCount += 1
                ApDashboardLog($"PdfCommentProcessor: batch {currentBatch} returned an empty response; review coverage is unconfirmed (this may also mean no findings)", "warn")
            End If

            APPdfWriteDiagnostic(diagnosticLog, "PdfCommentReviewBatch=" &
                Newtonsoft.Json.JsonConvert.SerializeObject(New With {
                    .batch = currentBatch, .sourceStart = batchStart, .sourceLength = charCount,
                    .returnedComments = returnedCommentCount,
                    .responseCharacters = If(llmResponse, System.String.Empty).Length,
                    .reviewState = If(returnedCommentCount > 0, "comments_received", "unconfirmed"),
                    .anchorContract = "pdf_full_quote_v1"
                }))
            batchStart = batchEnd
        End While

        ApDashboardLog($"PdfCommentProcessor: all {currentBatch} batches completed, {allComments.Count} total comments collected", "success")
        Return allComments
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TEXT SPAN → HIGHLIGHT RECTANGLES
    ' ═══════════════════════════════════════════════════════════════════════════


    ''' <summary>
    ''' Finds the character positions matching the search text and groups them
    ''' into per-line highlight rectangles, one per visual line.
    ''' </summary>
    Private Shared Function APPdfFindTextSpan(fullText As System.String,
                                              searchText As System.String,
                                              sourceStart As System.Int32,
                                              sourceLength As System.Int32) As PdfTextMatch
        Dim result As New PdfTextMatch()
        If fullText Is Nothing OrElse sourceStart < 0 OrElse sourceLength <= 0 OrElse
           sourceStart > fullText.Length OrElse sourceLength > fullText.Length - sourceStart Then
            result.FailureReason = "invalid_source_window"
            Return result
        End If
        If System.String.IsNullOrWhiteSpace(searchText) Then
            result.FailureReason = "empty_quote"
            Return result
        End If

        Dim cleanedSearch As System.String = searchText.Trim()
        If cleanedSearch.Length >= 2 Then
            Dim first As System.Char = cleanedSearch(0)
            Dim last As System.Char = cleanedSearch(cleanedSearch.Length - 1)
            If (first = """"c AndAlso last = """"c) OrElse
               (first = Microsoft.VisualBasic.ChrW(&H201C) AndAlso (last = Microsoft.VisualBasic.ChrW(&H201D) OrElse last = first)) OrElse
               (first = Microsoft.VisualBasic.ChrW(&H201E) AndAlso (last = Microsoft.VisualBasic.ChrW(&H201D) OrElse last = Microsoft.VisualBasic.ChrW(&H201C))) OrElse
               (first = Microsoft.VisualBasic.ChrW(&HAB) AndAlso last = Microsoft.VisualBasic.ChrW(&HBB)) Then
                cleanedSearch = cleanedSearch.Substring(1, cleanedSearch.Length - 2).Trim()
            End If
        End If
        If cleanedSearch.Length = 0 Then
            result.FailureReason = "empty_quote"
            Return result
        End If

        Dim sourceText As System.String = fullText.Substring(sourceStart, sourceLength)
        ' Use the same layout rules at word boundaries, including before the exact tier.
        ' Otherwise "Client" could incorrectly match the first half of "Client<soft-hyphen>Company".
        Dim boundaryText As System.String = Nothing
        Dim boundaryStarts As System.Collections.Generic.List(Of System.Int32) = Nothing
        Dim boundaryEnds As System.Collections.Generic.List(Of System.Int32) = Nothing
        APPdfBuildAnchorText(fullText, True, boundaryText, boundaryStarts, boundaryEnds)
        If APPdfTryMatchFullQuote(fullText, sourceText, cleanedSearch, System.StringComparison.Ordinal,
                                sourceStart, Nothing, Nothing, boundaryText, boundaryStarts, "exact", result) Then Return result
        If APPdfTryMatchFullQuote(fullText, sourceText, cleanedSearch, System.StringComparison.OrdinalIgnoreCase,
                                sourceStart, Nothing, Nothing, boundaryText, boundaryStarts, "case_insensitive", result) Then Return result

        ' Both normalization modes retain a start AND end map for every output character.
        ' Digits, punctuation and ordinary word boundaries are never discarded.
        For Each dehyphenate As System.Boolean In New System.Boolean() {False, True}
            Dim normalizedSource As System.String = Nothing
            Dim sourceStarts As System.Collections.Generic.List(Of System.Int32) = Nothing
            Dim sourceEnds As System.Collections.Generic.List(Of System.Int32) = Nothing
            APPdfBuildAnchorText(sourceText, dehyphenate, normalizedSource, sourceStarts, sourceEnds)

            Dim normalizedSearch As System.String = Nothing
            Dim searchStarts As System.Collections.Generic.List(Of System.Int32) = Nothing
            Dim searchEnds As System.Collections.Generic.List(Of System.Int32) = Nothing
            APPdfBuildAnchorText(cleanedSearch, dehyphenate, normalizedSearch, searchStarts, searchEnds)
            Dim method As System.String = If(dehyphenate, "line_end_dehyphenation", "normalized_whitespace")
            If APPdfTryMatchFullQuote(fullText, normalizedSource, normalizedSearch, System.StringComparison.OrdinalIgnoreCase,
                                    sourceStart, sourceStarts, sourceEnds, boundaryText, boundaryStarts, method, result) Then Return result
        Next

        ' No prefix, digit-stripping, compact-whitespace or global/out-of-batch fallback.
        Return result
    End Function

    ''' <summary>
    ''' True means the tier found one or more complete candidates. Ambiguity stops resolution.
    ''' </summary>
    Private Shared Function APPdfTryMatchFullQuote(fullText As System.String,
                                                   sourceText As System.String,
                                                   quote As System.String,
                                                   comparison As System.StringComparison,
                                                   sourceOffset As System.Int32,
                                                   startMap As System.Collections.Generic.List(Of System.Int32),
                                                   endMap As System.Collections.Generic.List(Of System.Int32),
                                                   boundaryText As System.String,
                                                   boundaryStarts As System.Collections.Generic.List(Of System.Int32),
                                                   method As System.String,
                                                   result As PdfTextMatch) As System.Boolean
        If quote.Length = 0 Then Return False
        Dim searchFrom As System.Int32 = 0
        Dim firstMatch As System.Int32 = -1
        Dim candidateCount As System.Int32 = 0

        While searchFrom <= sourceText.Length - quote.Length
            Dim found As System.Int32 = sourceText.IndexOf(quote, searchFrom, comparison)
            If found < 0 Then Exit While
            Dim endExclusive As System.Int32 = found + quote.Length
            Dim originalStart As System.Int32 = sourceOffset + If(startMap Is Nothing, found, startMap(found))
            Dim originalEnd As System.Int32 = sourceOffset + If(endMap Is Nothing, endExclusive, endMap(endExclusive - 1))
            ' A quote ending in "1" must not silently match the beginning of "11".
            Dim leftBoundary As System.Boolean = found = 0 OrElse
                Not (APPdfIsWordCharacter(quote(0)) AndAlso APPdfIsWordCharacter(sourceText(found - 1)))
            Dim rightBoundary As System.Boolean = endExclusive = sourceText.Length OrElse
                Not (APPdfIsWordCharacter(quote(quote.Length - 1)) AndAlso APPdfIsWordCharacter(sourceText(endExclusive)))
            ' A batch boundary is not evidence of a word boundary in the PDF.
            leftBoundary = leftBoundary AndAlso (originalStart = 0 OrElse
                Not (APPdfIsWordCharacter(quote(0)) AndAlso APPdfIsWordCharacter(fullText(originalStart - 1))))
            rightBoundary = rightBoundary AndAlso (originalEnd = fullText.Length OrElse
                Not (APPdfIsWordCharacter(quote(quote.Length - 1)) AndAlso APPdfIsWordCharacter(fullText(originalEnd))))
            If leftBoundary AndAlso rightBoundary AndAlso
               APPdfHasLayoutWordBoundaries(boundaryText, boundaryStarts, originalStart, originalEnd) Then
                If firstMatch < 0 Then firstMatch = found
                candidateCount += 1
            End If
            searchFrom = found + 1
        End While

        If candidateCount = 0 Then Return False
        result.Method = method
        result.CandidateCount = candidateCount
        If candidateCount <> 1 Then
            result.FailureReason = "ambiguous_quote"
            Return True
        End If

        Dim start As System.Int32 = If(startMap Is Nothing, firstMatch, startMap(firstMatch))
        Dim finish As System.Int32 = If(endMap Is Nothing, firstMatch + quote.Length, endMap(firstMatch + quote.Length - 1))
        result.Start = sourceOffset + start
        result.Length = finish - start
        result.FailureReason = System.String.Empty
        Return True
    End Function

    Private Shared Function APPdfHasLayoutWordBoundaries(boundaryText As System.String,
                                                          boundaryStarts As System.Collections.Generic.List(Of System.Int32),
                                                          start As System.Int32, finish As System.Int32) As System.Boolean
        Dim first As System.Int32 = boundaryStarts.BinarySearch(start)
        If first < 0 Then first = Not first
        Dim after As System.Int32 = boundaryStarts.BinarySearch(finish)
        If after < 0 Then after = Not after
        If first >= after OrElse first >= boundaryText.Length Then Return False
        Dim leftBoundary As System.Boolean = first = 0 OrElse
            Not (APPdfIsWordCharacter(boundaryText(first)) AndAlso APPdfIsWordCharacter(boundaryText(first - 1)))
        Dim rightBoundary As System.Boolean = after = boundaryText.Length OrElse
            Not (APPdfIsWordCharacter(boundaryText(after - 1)) AndAlso APPdfIsWordCharacter(boundaryText(after)))
        Return leftBoundary AndAlso rightBoundary
    End Function

    Private Shared Function APPdfIsWordCharacter(value As System.Char) As System.Boolean
        Dim category As System.Globalization.UnicodeCategory = System.Char.GetUnicodeCategory(value)
        Return System.Char.IsLetterOrDigit(value) OrElse
            category = System.Globalization.UnicodeCategory.NonSpacingMark OrElse
            category = System.Globalization.UnicodeCategory.SpacingCombiningMark OrElse
            category = System.Globalization.UnicodeCategory.ConnectorPunctuation
    End Function

    ''' <summary>
    ''' Collapses whitespace. The optional layout tier removes soft hyphens and a letter-
    ''' hyphen-letter split across one physical line break; ordinary hyphens/digits remain.
    ''' Each mapped end is exclusive.
    ''' </summary>
    Private Shared Sub APPdfBuildAnchorText(text As System.String,
                                            dehyphenate As System.Boolean,
                                            ByRef normalized As System.String,
                                            ByRef startMap As System.Collections.Generic.List(Of System.Int32),
                                            ByRef endMap As System.Collections.Generic.List(Of System.Int32))
        Dim builder As New System.Text.StringBuilder(text.Length)
        startMap = New System.Collections.Generic.List(Of System.Int32)(text.Length)
        endMap = New System.Collections.Generic.List(Of System.Int32)(text.Length)
        Dim i As System.Int32 = 0
        While i < text.Length
            Dim ch As System.Char = text(i)
            If dehyphenate AndAlso (ch = "-"c OrElse ch = Microsoft.VisualBasic.ChrW(&H2010) OrElse
                                   ch = Microsoft.VisualBasic.ChrW(&HAD)) Then
                Dim nextLetter As System.Int32 = i + 1
                Dim lineBreaks As System.Int32 = 0
                While nextLetter < text.Length AndAlso System.Char.IsWhiteSpace(text(nextLetter))
                    If text(nextLetter) = Microsoft.VisualBasic.ChrW(10) OrElse
                       (text(nextLetter) = Microsoft.VisualBasic.ChrW(13) AndAlso
                        (nextLetter + 1 >= text.Length OrElse text(nextLetter + 1) <> Microsoft.VisualBasic.ChrW(10))) Then
                        lineBreaks += 1
                    End If
                    nextLetter += 1
                End While
                If lineBreaks = 1 AndAlso i > 0 AndAlso System.Char.IsLetter(text(i - 1)) AndAlso
                   nextLetter < text.Length AndAlso System.Char.IsLetter(text(nextLetter)) Then
                    i = nextLetter
                    Continue While
                End If
                If ch = Microsoft.VisualBasic.ChrW(&HAD) Then
                    i += 1
                    Continue While
                End If
            End If

            Dim originalStart As System.Int32 = i
            If System.Char.IsWhiteSpace(ch) Then
                i += 1
                While i < text.Length AndAlso System.Char.IsWhiteSpace(text(i))
                    i += 1
                End While
                builder.Append(" "c)
            Else
                builder.Append(ch)
                i += 1
            End If
            startMap.Add(originalStart)
            endMap.Add(i)
        End While
        normalized = builder.ToString()
    End Sub

    ''' <summary>Resolves a complete source quote, then builds small, contiguous word quads.</summary>
    Private Shared Function APPdfFindTextRects(fullText As System.String,
                                               charPositions As System.Collections.Generic.List(Of PdfCharPosition),
                                               entry As PdfCommentEntry,
                                               pageCount As System.Int32,
                                               ByRef match As PdfTextMatch) As System.Collections.Generic.List(Of PdfHighlightRect)
        match = APPdfFindTextSpan(fullText, entry.QuotedText, entry.SourceStart, entry.SourceLength)
        If Not System.String.IsNullOrEmpty(match.FailureReason) Then Return Nothing
        If charPositions Is Nothing OrElse charPositions.Count <> fullText.Length Then
            match.FailureReason = "position_map_mismatch"
            Return Nothing
        End If
        If match.Start < entry.SourceStart OrElse match.Length <= 0 OrElse
           match.Start + match.Length > entry.SourceStart + entry.SourceLength Then
            match.FailureReason = "match_outside_source_window"
            Return Nothing
        End If

        Dim matchedPositions As New System.Collections.Generic.List(Of PdfCharPosition)()
        For i As System.Int32 = match.Start To match.Start + match.Length - 1
            Dim pos As PdfCharPosition = charPositions(i)
            If pos Is Nothing OrElse pos.PageIndex < 0 OrElse pos.PageIndex >= pageCount OrElse
               Not APPdfIsFiniteCoordinate(pos.X) OrElse Not APPdfIsFiniteCoordinate(pos.Y) OrElse
               Not APPdfIsFiniteCoordinate(pos.Width) OrElse Not APPdfIsFiniteCoordinate(pos.Height) OrElse
               pos.Width < 0 OrElse pos.Height < 0 Then
                match.FailureReason = "invalid_geometry"
                Return Nothing
            End If
            If pos.Width > 0 OrElse pos.Height > 0 Then matchedPositions.Add(pos)
        Next
        If matchedPositions.Count = 0 Then
            match.FailureReason = "no_positioned_characters"
            Return Nothing
        End If

        Dim rects As New System.Collections.Generic.List(Of PdfHighlightRect)()
        Dim lineChars As New System.Collections.Generic.List(Of PdfCharPosition) From {matchedPositions(0)}
        For i As System.Int32 = 1 To matchedPositions.Count - 1
            Dim prev As PdfCharPosition = matchedPositions(i - 1)
            Dim curr As PdfCharPosition = matchedPositions(i)
            Dim yTolerance As System.Double = System.Math.Max(prev.Height * 0.5, 2.0)
            Dim horizontalGap As System.Double = System.Math.Max(curr.X - (prev.X + prev.Width),
                                                                 prev.X - (curr.X + curr.Width))
            Dim sameFragment As System.Boolean =
                curr.PageIndex = prev.PageIndex AndAlso curr.WordIndex = prev.WordIndex AndAlso
                System.Math.Abs(curr.Y - prev.Y) <= yTolerance AndAlso
                horizontalGap <= System.Math.Max(System.Math.Max(prev.Height, curr.Height), 2.0)
            If sameFragment Then
                lineChars.Add(curr)
            Else
                rects.Add(APPdfBuildLineRect(lineChars))
                lineChars.Clear()
                lineChars.Add(curr)
            End If
        Next
        rects.Add(APPdfBuildLineRect(lineChars))
        If rects.Any(Function(r) r.Width <= 0 OrElse r.Height <= 0 OrElse
                     Not APPdfIsFiniteCoordinate(r.X + r.Width) OrElse Not APPdfIsFiniteCoordinate(r.Y + r.Height)) Then
            match.FailureReason = "invalid_geometry"
            Return Nothing
        End If
        Return rects
    End Function

    Private Shared Function APPdfIsFiniteCoordinate(value As System.Double) As System.Boolean
        Return Not System.Double.IsNaN(value) AndAlso Not System.Double.IsInfinity(value)
    End Function
    ''' <summary>
    ''' Builds a rectangle from a contiguous fragment of one extracted word on one visual line.
    ''' </summary>
    Private Shared Function APPdfBuildLineRect(lineChars As List(Of PdfCharPosition)) As PdfHighlightRect
        Dim minX = lineChars.Min(Function(c) c.X)
        Dim minY = lineChars.Min(Function(c) c.Y)
        Dim maxX = lineChars.Max(Function(c) c.X + c.Width)
        Dim maxY = lineChars.Max(Function(c) c.Y + c.Height)

        Return New PdfHighlightRect() With {
            .PageIndex = lineChars(0).PageIndex,
            .X = minX,
            .Y = minY,
            .Width = maxX - minX,
            .Height = maxY - minY
        }
    End Function



    Private Shared Sub APPdfLogAnchorDiagnostic(diagnosticLog As System.Action(Of System.String),
                                                 commentId As System.Int32,
                                                 entry As PdfCommentEntry,
                                                 match As PdfTextMatch,
                                                 rectangles As System.Collections.Generic.List(Of PdfHighlightRect),
                                                 fullText As System.String)
        Dim actualText As System.String = System.String.Empty
        If match.Start >= 0 AndAlso match.Length > 0 AndAlso match.Start <= fullText.Length AndAlso
           match.Length <= fullText.Length - match.Start Then
            actualText = fullText.Substring(match.Start, match.Length)
        End If
        Dim safeRectangles As System.Collections.Generic.List(Of PdfHighlightRect) =
            If(rectangles, New System.Collections.Generic.List(Of PdfHighlightRect)())
        APPdfWriteDiagnostic(diagnosticLog, "PdfCommentAnchor=" &
            Newtonsoft.Json.JsonConvert.SerializeObject(New With {
                .stage = "resolved",
                .commentId = commentId,
                .batch = entry.BatchNumber,
                .status = If(System.String.IsNullOrEmpty(match.FailureReason), "anchored", "unanchored"),
                .reason = match.FailureReason,
                .method = match.Method,
                .candidateCount = match.CandidateCount,
                .sourceStart = entry.SourceStart,
                .sourceLength = entry.SourceLength,
                .start = match.Start,
                .length = match.Length,
                .quote = entry.QuotedText,
                .matchedText = actualText,
                .comment = entry.CommentText,
                .rectangles = safeRectangles.Select(Function(r) New With {
                    .page = r.PageIndex + 1, .x = r.X, .y = r.Y, .width = r.Width, .height = r.Height
                }).ToArray()
            }))
    End Sub

    Private Shared Sub APPdfWriteDiagnostic(diagnosticLog As System.Action(Of System.String),
                                             message As System.String)
        System.Diagnostics.Debug.WriteLine(message)
        If diagnosticLog Is Nothing Then Return
        Try
            diagnosticLog(message)
        Catch ex As System.Exception
            System.Diagnostics.Debug.WriteLine("PdfCommentProcessor: diagnostic sink failed: " & ex.ToString())
        End Try
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    '  PDF ANNOTATION INJECTION (PdfSharp low-level dictionary manipulation)
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Inserts a /Highlight annotation with an associated /Popup annotation for a matched
    ''' text span. The highlight covers all line rectangles; the popup is anchored near the
    ''' first highlight rectangle.
    ''' </summary>
    Private Shared Sub APPdfInsertHighlightAnnotation(pdfDoc As PdfDocument,
                                                       highlightRects As List(Of PdfHighlightRect),
                                                       commentText As String,
                                                       author As String,
                                                       commentId As Integer)
        ' Group rectangles by page
        Dim rectsByPage = highlightRects.GroupBy(Function(r) r.PageIndex)

        For Each pageGroup In rectsByPage
            Dim pageIdx = pageGroup.Key
            If pageIdx < 0 OrElse pageIdx >= pdfDoc.Pages.Count Then Continue For

            Dim page = pdfDoc.Pages(pageIdx)
            Dim pageRects = pageGroup.ToList()

            ' Compute the overall bounding rectangle for the /Rect entry
            Dim overallMinX = pageRects.Min(Function(r) r.X)
            Dim overallMinY = pageRects.Min(Function(r) r.Y)
            Dim overallMaxX = pageRects.Max(Function(r) r.X + r.Width)
            Dim overallMaxY = pageRects.Max(Function(r) r.Y + r.Height)

            ' Build QuadPoints array — 8 values per line rectangle
            ' PDF spec: QuadPoints are x1,y1,...,x4,y4 for each quad (top-left, top-right, bottom-left, bottom-right)
            Dim quadPointsArray As New PdfArray(pdfDoc)
            For Each rect In pageRects
                Dim x1 = rect.X                ' left
                Dim x2 = rect.X + rect.Width   ' right
                Dim y1 = rect.Y                ' bottom
                Dim y2 = rect.Y + rect.Height  ' top
                ' PDF QuadPoints order: top-left, top-right, bottom-left, bottom-right
                quadPointsArray.Elements.Add(New PdfReal(x1))   ' top-left X
                quadPointsArray.Elements.Add(New PdfReal(y2))   ' top-left Y
                quadPointsArray.Elements.Add(New PdfReal(x2))   ' top-right X
                quadPointsArray.Elements.Add(New PdfReal(y2))   ' top-right Y
                quadPointsArray.Elements.Add(New PdfReal(x1))   ' bottom-left X
                quadPointsArray.Elements.Add(New PdfReal(y1))   ' bottom-left Y
                quadPointsArray.Elements.Add(New PdfReal(x2))   ' bottom-right X
                quadPointsArray.Elements.Add(New PdfReal(y1))   ' bottom-right Y
            Next

            ' Create the /Highlight annotation
            Dim highlight As New PdfDictionary(pdfDoc)
            highlight.Elements.Add("/Type", New PdfName("/Annot"))
            highlight.Elements.Add("/Subtype", New PdfName("/Highlight"))
            highlight.Elements.Add("/Rect", APPdfMakeRect(pdfDoc, overallMinX, overallMinY, overallMaxX, overallMaxY))
            highlight.Elements.Add("/QuadPoints", quadPointsArray)
            highlight.Elements.Add("/Contents", New PdfString(commentText))
            highlight.Elements.Add("/T", New PdfString(author))
            highlight.Elements.Add("/CreationDate", New PdfString(APPdfFormatDate(DateTime.UtcNow)))
            highlight.Elements.Add("/M", New PdfString(APPdfFormatDate(DateTime.UtcNow)))
            highlight.Elements.Add("/NM", New PdfString("RI_Comment_" & commentId.ToString(System.Globalization.CultureInfo.InvariantCulture) & "_" & System.Guid.NewGuid().ToString("N")))
            highlight.Elements.Add("/F", New PdfInteger(4))  ' Print flag
            highlight.Elements.Add("/C", APPdfMakeColorArray(pdfDoc, 1.0, 0.85, 0.0)) ' Yellow highlight
            highlight.Elements.Add("/CA", New PdfReal(0.4))  ' Semi-transparent highlight opacity

            ' Create the /Popup annotation
            Dim popupX = Math.Min(overallMaxX + 5, page.Width.Point - 210)
            Dim popupY = Math.Max(overallMinY - 100, 5)
            Dim popup As New PdfDictionary(pdfDoc)
            popup.Elements.Add("/Type", New PdfName("/Annot"))
            popup.Elements.Add("/Subtype", New PdfName("/Popup"))
            popup.Elements.Add("/Rect", APPdfMakeRect(pdfDoc, popupX, popupY, popupX + 200, popupY + 100))
            popup.Elements.Add("/Open", New PdfBoolean(False))

            ' Register both objects so PdfSharp assigns indirect references
            pdfDoc.Internals.AddObject(highlight)
            pdfDoc.Internals.AddObject(popup)

            ' Cross-reference highlight ↔ popup
            highlight.Elements.Add("/Popup", popup.Reference)
            popup.Elements.Add("/Parent", highlight.Reference)

            ' Add both annotations to the page's /Annots array
            Dim annots = page.Elements.GetArray("/Annots")
            If annots Is Nothing Then
                annots = New PdfArray(pdfDoc)
                page.Elements.Add("/Annots", annots)
            End If
            annots.Elements.Add(highlight.Reference)
            annots.Elements.Add(popup.Reference)
        Next
    End Sub

    ''' <summary>
    ''' Inserts a /Text (sticky note) annotation at a specific position on a page.
    ''' Used for unmatched comments that couldn't be placed on their referenced text.
    ''' Each sticky note gets a unique position so they are individually selectable.
    ''' </summary>
    Private Shared Sub APPdfInsertStickyNoteAnnotation(pdfDoc As PdfDocument,
                                                        pageIndex As Integer,
                                                        x As Double, y As Double,
                                                        commentText As String,
                                                        author As String,
                                                        commentId As Integer)
        If pageIndex < 0 OrElse pageIndex >= pdfDoc.Pages.Count Then Return

        Dim page = pdfDoc.Pages(pageIndex)

        ' Create the /Text annotation (sticky note icon)
        Dim stickyNote As New PdfDictionary(pdfDoc)
        stickyNote.Elements.Add("/Type", New PdfName("/Annot"))
        stickyNote.Elements.Add("/Subtype", New PdfName("/Text"))
        stickyNote.Elements.Add("/Rect", APPdfMakeRect(pdfDoc, x, y, x + AP_PdfStickyNoteSize, y + AP_PdfStickyNoteSize))
        stickyNote.Elements.Add("/Contents", New PdfString(commentText))
        stickyNote.Elements.Add("/T", New PdfString(author))
        stickyNote.Elements.Add("/Name", New PdfName("/Comment"))  ' Comment icon style
        stickyNote.Elements.Add("/CreationDate", New PdfString(APPdfFormatDate(DateTime.UtcNow)))
        stickyNote.Elements.Add("/M", New PdfString(APPdfFormatDate(DateTime.UtcNow)))
        stickyNote.Elements.Add("/NM", New PdfString("RI_Unmatched_" & commentId.ToString(System.Globalization.CultureInfo.InvariantCulture) & "_" & System.Guid.NewGuid().ToString("N")))
        stickyNote.Elements.Add("/F", New PdfInteger(4))  ' Print flag
        stickyNote.Elements.Add("/C", APPdfMakeColorArray(pdfDoc, 1.0, 0.6, 0.0))  ' Orange for unmatched
        stickyNote.Elements.Add("/Open", New PdfBoolean(False))

        ' Create associated popup
        Dim popup As New PdfDictionary(pdfDoc)
        popup.Elements.Add("/Type", New PdfName("/Annot"))
        popup.Elements.Add("/Subtype", New PdfName("/Popup"))
        popup.Elements.Add("/Rect", APPdfMakeRect(pdfDoc, x - 200, y - 100, x, y))
        popup.Elements.Add("/Open", New PdfBoolean(False))

        pdfDoc.Internals.AddObject(stickyNote)
        pdfDoc.Internals.AddObject(popup)

        stickyNote.Elements.Add("/Popup", popup.Reference)
        popup.Elements.Add("/Parent", stickyNote.Reference)

        Dim annots = page.Elements.GetArray("/Annots")
        If annots Is Nothing Then
            annots = New PdfArray(pdfDoc)
            page.Elements.Add("/Annots", annots)
        End If
        annots.Elements.Add(stickyNote.Reference)
        annots.Elements.Add(popup.Reference)
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    '  PDF HELPER FUNCTIONS
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Creates a PDF /Rect array [x1 y1 x2 y2].
    ''' </summary>
    Private Shared Function APPdfMakeRect(pdfDoc As PdfDocument,
                                          x1 As Double, y1 As Double,
                                          x2 As Double, y2 As Double) As PdfArray
        Dim arr As New PdfArray(pdfDoc)
        arr.Elements.Add(New PdfReal(x1))
        arr.Elements.Add(New PdfReal(y1))
        arr.Elements.Add(New PdfReal(x2))
        arr.Elements.Add(New PdfReal(y2))
        Return arr
    End Function

    ''' <summary>
    ''' Creates a PDF color array [r g b] with values in 0.0–1.0 range.
    ''' </summary>
    Private Shared Function APPdfMakeColorArray(pdfDoc As PdfDocument,
                                                r As Double, g As Double, b As Double) As PdfArray
        Dim arr As New PdfArray(pdfDoc)
        arr.Elements.Add(New PdfReal(r))
        arr.Elements.Add(New PdfReal(g))
        arr.Elements.Add(New PdfReal(b))
        Return arr
    End Function

    ''' <summary>
    ''' Formats a DateTime as a PDF date string: D:YYYYMMDDHHmmSSOHH'mm'
    ''' </summary>
    Private Shared Function APPdfFormatDate(dt As DateTime) As String
        Return "D:" & dt.ToString("yyyyMMddHHmmss") & "Z"
    End Function

End Class