' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.Processing.FormatSaveAndRestore.vb
' Purpose: Extracts Word range content together with formatting placeholders,
'          serializes special elements (notes, fields, highlights, paragraph formats),
'          and restores them back into the document for Red Ink for Word workflows.
'
' Architecture:
'  - Paragraph Format Capture: Builds ParagraphFormatStructure arrays and injects {{PFOR:n}} markers.
'  - Markdown Conversion: Iteratively replaces bold/italic/underline/strike/highlight formatting
'    with Markdown-compatible tokens while preserving range integrity.
'  - Placeholder Pipeline: Scans ranges for footnotes, endnotes, fields, and paragraph formats,
'    converts them to inline placeholders, sorts them, and merges them into the extracted text.
'  - Restoration: Replays placeholders back into Word by recreating notes, fields, and formatting.
'  - Safety: Uses ProgressScope/ProgressBarModule for cancellation, preserves selection settings,
'    and avoids altering user formatting when not explicitly requested.
'  - Dependencies: Microsoft.Office.Interop.Word, SharedLibrary.SharedLibrary, ProgressScope,
'    and ParagraphFormatStructure definitions elsewhere in the add-in.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports Microsoft.Office.Interop.Word
Imports SharedLibrary.SharedLibrary

Partial Public Class ThisAddIn

    ''' <summary>
    ''' Applies previously captured paragraph formatting metadata to the supplied range.
    ''' </summary>
    ''' <param name="rng">Target range whose paragraphs are updated.</param>
    Public Sub ApplyParagraphFormat(ByRef rng As Word.Range)
        Dim maxParaStylesCount As Integer = paragraphFormat.Length
        Dim paraCount As Integer = rng.Paragraphs.Count

        If paraCount = 0 Then Exit Sub

        For i As Integer = 1 To paraCount
            If i - 1 >= maxParaStylesCount Then Exit For

            Dim pf As ParagraphFormatStructure = paragraphFormat(i - 1)
            Dim pRange As Word.Range = rng.Paragraphs(i).Range

            '--- 1. paragraph style ------------------------------------------------
            If pf.Style IsNot Nothing Then
                Try
                    pRange.Style = pf.Style
                Catch ex As System.Exception
                    ' handle / log if necessary
                End Try
            End If

            '--- 2. character-level attributes – use them *only when supplied* -----
            With pRange.Font
                If Not String.IsNullOrEmpty(pf.FontName) Then .Name = pf.FontName
                If pf.FontSize.HasValue Then .Size = pf.FontSize.Value
                If pf.FontBold.HasValue Then .Bold = pf.FontBold.Value
                If pf.FontItalic.HasValue Then .Italic = pf.FontItalic.Value
                If pf.FontUnderline.HasValue Then .Underline = pf.FontUnderline.Value
                If pf.FontColor.HasValue Then .Color = pf.FontColor.Value
            End With

            '--- 3. list formatting -----------------------------------------------
            If pf.HasListFormat AndAlso pf.ListTemplate IsNot Nothing Then
                Try
                    If pRange.ListFormat.ListType <> Word.WdListType.wdListNoNumbering Then
                        pRange.ListFormat.RemoveNumbers()
                    End If

                    pRange.ListFormat.ApplyListTemplateWithLevel(
                        ListTemplate:=pf.ListTemplate,
                        ContinuePreviousList:=pf.ListLevel > 0,
                        ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList,
                        DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior)
                    pRange.ListFormat.ListLevelNumber = pf.ListLevel
                Catch ex As System.Exception
                    ' handle / log if necessary
                End Try
            End If

            '--- 4. paragraph-level attributes (spacing restored exactly) ----------
            With pRange.ParagraphFormat
                .Alignment = pf.Alignment

                ' Important flags first
                .DisableLineHeightGrid = pf.DisableLineHeightGrid

                ' Spacing rule, then value
                .LineSpacingRule = pf.LineSpacingRule
                .LineSpacing = pf.LineSpacing

                ' Auto spacing flags, then explicit values only if not auto
                .SpaceBeforeAuto = pf.SpaceBeforeAuto
                .SpaceAfterAuto = pf.SpaceAfterAuto
                If Not pf.SpaceBeforeAuto Then .SpaceBefore = pf.SpaceBefore
                If Not pf.SpaceAfterAuto Then .SpaceAfter = pf.SpaceAfter
            End With
        Next
    End Sub

    ''' <summary>
    ''' Ensures {{PFOR:n}} markers begin on their own line, except for {{PFOR:0}}.
    ''' </summary>
    ''' <param name="input">Text containing paragraph format markers.</param>
    ''' <returns>Adjusted text with corrected marker placement.</returns>
    Public Function CorrectPFORMarkers(ByVal input As String) As String
        Try
            Dim output As New StringBuilder()
            Dim i As Integer = 0
            Dim length As Integer = input.Length

            While i < length
                ' Detect PFOR markers
                If i <= length - 9 AndAlso input.Substring(i, 7) = "{{PFOR:" Then
                    ' Check if it's "PFOR:0"
                    Dim endIndex As Integer = input.IndexOf("}}", i)
                    If endIndex <> -1 Then
                        Dim markerContent As String = input.Substring(i + 7, endIndex - (i + 7)) ' Extract "nnn"
                        If markerContent = "0" Then
                            ' If it's PFOR:0, copy as-is and move the pointer
                            output.Append(input.Substring(i, endIndex - i + 2))
                            i = endIndex + 2
                            Continue While
                        End If
                    End If

                    ' Check preceding character
                    If output.Length > 0 Then
                        Dim prevChar As Char = output(output.Length - 1)
                        If prevChar <> vbCr AndAlso prevChar <> vbLf Then
                            output.Append(vbCrLf) ' Add newline before the marker
                        End If
                    End If

                    ' Append the marker
                    Dim markerEnd As Integer = input.IndexOf("}}", i) + 2
                    output.Append(input.Substring(i, markerEnd - i))
                    i = markerEnd
                Else
                    ' Copy character-by-character
                    output.Append(input(i))
                    i += 1
                End If
            End While

            Return output.ToString()
        Catch ex As System.Exception
            Debug.WriteLine("An error occurred while correcting PFOR markers: " & ex.Message, ex)
        End Try
    End Function

    ''' <summary>
    ''' Stores positional and replacement data for placeholder tokens.
    ''' </summary>
    Private Structure PlaceholderInfo
        Public Offset As Integer     'offset relative to rng.Start (0-based)
        Public Length As Integer     'chars to skip in Word range
        Public Token As String       'replacement text ({{WFNT:…}}, {{PFOR:…}}, …)
    End Structure

    ''' <summary>
    ''' Orders placeholders by offset and then by length.
    ''' </summary>
    Private Shared ReadOnly PlaceholderComparer As Comparison(Of PlaceholderInfo) =
    Function(a, b)
        If a.Offset <> b.Offset Then
            Return a.Offset.CompareTo(b.Offset)
        End If
        Return a.Length.CompareTo(b.Length)
    End Function

    ''' <summary>
    ''' Extracts text from a Word range, replaces special elements with inline placeholders,
    ''' and optionally converts formatting to Markdown-compatible markers.
    ''' </summary>
    ''' <param name="workingrange">Range whose content is extracted.</param>
    ''' <param name="PreserveParagraphFormatInline">True to capture per-paragraph formatting.</param>
    ''' <param name="DoMarkdown">True to emit Markdown markers for styles.</param>
    ''' <returns>Serialized text containing placeholder tokens.</returns>
    Public Function GetTextWithSpecialElementsInline(
        ByVal workingrange As Word.Range,
        PreserveParagraphFormatInline As Boolean, DoMarkdown As Boolean) As String

        'Dim splash As New SLib.SplashScreen("Extracting text and format (longer text/tables may take time) ...")
        'splash.Show()
        'splash.Refresh()

        Dim app As Word.Application = CType(workingrange.Application, Word.Application)
        Dim oldSU As Boolean = app.ScreenUpdating
        Dim oldSpell As Boolean = app.Options.CheckSpellingAsYouType
        Dim oldGrammar As Boolean = app.Options.CheckGrammarAsYouType
        app.Options.CheckSpellingAsYouType = False
        app.Options.CheckGrammarAsYouType = False
        app.ScreenUpdating = False

        ' ---- Progress + Cancel helpers ----
        Dim current As Integer = 0
        Dim total As Integer = 5 ' baseline: setup + bookmark + sort + build + finalize

        ' We’ll accumulate an estimated total before starting UI; then refine later if needed.
        ' Markdown steps (12 calls: 6 bold/italic + 1 underline inline + 1 strike para + 1 strike inline + 2 highlight + 1 bold/italic inline already counted)
        If DoMarkdown Then total += 12

        ' Temporary dup range for counting without side effects
        Dim tmpRng As Word.Range = workingrange.Duplicate
        Dim rngDoc As Word.Document = tmpRng.Document

        ' Count footnotes/endnotes within range
        Dim fnCount As Integer = 0
        For Each fn As Word.Footnote In rngDoc.Footnotes
            If fn.Reference.Start >= tmpRng.Start AndAlso fn.Reference.Start < tmpRng.End Then fnCount += 1
        Next
        Dim enCount As Integer = 0
        For Each en As Word.Endnote In rngDoc.Endnotes
            If en.Reference.Start >= tmpRng.Start AndAlso en.Reference.Start < tmpRng.End Then enCount += 1
        Next

        ' Count fields and paragraphs in our range
        Dim fieldsCount As Integer = tmpRng.Fields.Count
        Dim parasInRange As Integer = If(PreserveParagraphFormatInline AndAlso tmpRng.Paragraphs IsNot Nothing, tmpRng.Paragraphs.Count, 0)

        total += fnCount + enCount + fieldsCount + parasInRange

        ' Set up UI scope (uses ProgressBarModule under the hood)        
        Using scope As New ProgressScope("Extracting text …", "Initializing …", System.Math.Max(1, total), useDpiForm:=False)
            Dim Cancelled As Func(Of Boolean) =
                Function() As Boolean
                    Return scope.CancelRequested OrElse ProgressBarModule.CancelOperation
                End Function

            Try
                '──────────── 0)  Preparation (range clone, settings) ───────────────
                Dim rng As Word.Range = workingrange.Duplicate
                Dim expandedEnd As Boolean = False
                If rng.End < rng.Document.Content.End - 1 Then
                    rng.End = rng.End + 1
                    expandedEnd = True
                End If

                ProgressScope.Report(System.Threading.Interlocked.Increment(current), label:="Prepping selection …")
                If Cancelled() Then Return rng.Text

                '──────────── Apply formatting ────────────────────────────

                Debug.WriteLine($"4-1 Range Start = {rng.Start} Selection Start = {Application.Selection.Start}")
                Debug.WriteLine($"Range End = {rng.End} Selection End = {Application.Selection.End}")

                ' 0a) Markdown for combinations & single formats (with CR handling)
                Dim origSel As Word.Range = app.Selection.Range.Duplicate

                If DoMarkdown Then
                    ' 1) Bold + Italic  (paragraph)
                    ReplaceWithinRange(rng,
                            Sub(f)
                                f.Font.Bold = True
                                f.Font.Italic = True
                                f.Font.Underline = Word.WdUnderline.wdUnderlineNone
                                f.Text = "([!^13]@)^13"
                                f.MatchWildcards = True
                            End Sub,
                            "***\1***^13",
                            Sub(rep)
                                rep.Bold = False
                                rep.Italic = False
                            End Sub)
                    ProgressScope.Report(System.Threading.Interlocked.Increment(current), label:="Markdown: bold+italic (para) …")
                    If Cancelled() Then Return rng.Text

                    ' 2) Bold + Italic  (inline)
                    ReplaceWithinRange(rng,
                            Sub(f)
                                f.Font.Bold = True
                                f.Font.Italic = True
                                f.Font.Underline = Word.WdUnderline.wdUnderlineNone
                                f.Text = ""
                                f.MatchWildcards = False
                            End Sub,
                            "***^&***",
                            Sub(rep)
                                rep.Bold = False
                                rep.Italic = False
                            End Sub)
                    ProgressScope.Report(System.Threading.Interlocked.Increment(current), label:="Markdown: bold+italic (inline) …")
                    If Cancelled() Then Return rng.Text

                    Debug.WriteLine($"4-2 Range Start = {rng.Start} Selection Start = {Application.Selection.Start}")
                    Debug.WriteLine($"Range End = {rng.End} Selection End = {Application.Selection.End}")

                    ' 3) Bold only  (paragraph)
                    ReplaceWithinRange(rng,
                            Sub(f)
                                f.Font.Bold = True
                                f.Text = "([!^13]@)^13"
                                f.MatchWildcards = True
                            End Sub,
                            "**\1**^13",
                            Sub(rep)
                                rep.Bold = False
                            End Sub)
                    ProgressScope.Report(System.Threading.Interlocked.Increment(current), label:="Markdown: bold (para) …")
                    If Cancelled() Then Return rng.Text

                    ' 4) Bold only  (inline)
                    ReplaceWithinRange(rng,
                            Sub(f)
                                f.Font.Bold = True
                                f.Text = ""
                                f.MatchWildcards = False
                            End Sub,
                            "**^&**",
                            Sub(rep)
                                rep.Bold = False
                            End Sub)
                    ProgressScope.Report(System.Threading.Interlocked.Increment(current), label:="Markdown: bold (inline) …")
                    If Cancelled() Then Return rng.Text

                    Debug.WriteLine($"4-3 Range Start = {rng.Start} Selection Start = {Application.Selection.Start}")
                    Debug.WriteLine($"Range End = {rng.End} Selection End = {Application.Selection.End}")

                    ' 5) Italic only  (paragraph)
                    ReplaceWithinRange(rng,
                            Sub(f)
                                f.Font.Italic = True
                                f.Text = "([!^13]@)^13"
                                f.MatchWildcards = True
                            End Sub,
                            "*\1*^13",
                            Sub(rep)
                                rep.Italic = False
                            End Sub)
                    ProgressScope.Report(System.Threading.Interlocked.Increment(current), label:="Markdown: italic (para) …")
                    If Cancelled() Then Return rng.Text

                    ' 6) Italic only  (inline)
                    ReplaceWithinRange(rng,
                            Sub(f)
                                f.Font.Italic = True
                                f.Text = ""
                                f.MatchWildcards = False
                            End Sub,
                            "*^&*",
                            Sub(rep)
                                rep.Italic = False
                            End Sub)
                    ProgressScope.Report(System.Threading.Interlocked.Increment(current), label:="Markdown: italic (inline) …")
                    If Cancelled() Then Return rng.Text

                    Debug.WriteLine($"4-4 Range Start = {rng.Start} Selection Start = {Application.Selection.Start}")
                    Debug.WriteLine($"Range End = {rng.End} Selection End = {Application.Selection.End}")

                    ' 8) Underline  (inline)
                    ReplaceWithinRange(rng,
                            Sub(f)
                                f.Font.Underline = Word.WdUnderline.wdUnderlineSingle
                                f.Text = ""
                                f.MatchWildcards = False
                            End Sub,
                            "<u>^&</u>",
                            Sub(rep)
                                rep.Underline = Word.WdUnderline.wdUnderlineNone
                            End Sub)
                    ProgressScope.Report(System.Threading.Interlocked.Increment(current), label:="Markdown: underline (inline) …")
                    If Cancelled() Then Return rng.Text

                    ' 9) Strikethrough  (paragraph)
                    ReplaceWithinRange(rng,
                            Sub(f)
                                f.Font.StrikeThrough = True
                                f.Text = "([!^13]@)^13"
                                f.MatchWildcards = True
                            End Sub,
                            "~~\1~~^13",
                            Sub(rep)
                                rep.StrikeThrough = False
                            End Sub)
                    ProgressScope.Report(System.Threading.Interlocked.Increment(current), label:="Markdown: strike (para) …")
                    If Cancelled() Then Return rng.Text

                    '10) Strikethrough  (inline)
                    ReplaceWithinRange(rng,
                            Sub(f)
                                f.Font.StrikeThrough = True
                                f.Text = ""
                                f.MatchWildcards = False
                            End Sub,
                            "~~^&~~",
                            Sub(rep)
                                rep.StrikeThrough = False
                            End Sub)
                    ProgressScope.Report(System.Threading.Interlocked.Increment(current), label:="Markdown: strike (inline) …")
                    If Cancelled() Then Return rng.Text

                    ' Paragraph-like + color suffix:
                    ReplaceWithinRange_Highlight(Application.Selection.Range, keepParagraphBreakOutside:=True, includeColorInTag:=True)
                    ProgressScope.Report(System.Threading.Interlocked.Increment(current), label:="Markdown: highlight (para-like) …")
                    If Cancelled() Then Return rng.Text

                    ' Preserve highlight color for non-yellow: <mark:color>…</mark>
                    ReplaceWithinRange_Highlight(Application.Selection.Range, includeColorInTag:=True)
                    ProgressScope.Report(System.Threading.Interlocked.Increment(current), label:="Markdown: highlight (inline) …")
                    If Cancelled() Then Return rng.Text
                End If

                If expandedEnd Then
                    rng.End = rng.End - 1
                End If
                rng.Select()

                Debug.WriteLine($"4-5 Range Start = {rng.Start} Selection Start = {Application.Selection.Start}")
                Debug.WriteLine($"Range End = {rng.End} Selection End = {Application.Selection.End}")

                '──────────── Prepare placeholders ────────────────────────────

                Dim doc As Word.Document = workingrange.Application.ActiveDocument
                Dim bmName As String = "__TMP_RNG_" & Guid.NewGuid().ToString("N")

                ' Create bookmark (stores start & end)
                doc.Bookmarks.Add(Name:=bmName, Range:=rng)
                ProgressScope.Report(System.Threading.Interlocked.Increment(current), label:="Preparing placeholders …")
                If Cancelled() Then Return rng.Text

                ' Switch retrieval mode
                With rng.TextRetrievalMode
                    .IncludeHiddenText = True
                    .IncludeFieldCodes = True
                End With

                ' Read bookmark and reset range
                Dim bmRange As Word.Range = doc.Bookmarks(bmName).Range
                rng.SetRange(bmRange.Start, bmRange.End)

                ' Cleanup
                doc.Bookmarks(bmName).Delete()

                Dim placeholders As New List(Of PlaceholderInfo)

                '──────────── Collect footnotes & endnotes ────────────────────────────
                Dim processed As Integer

                processed = 0
                For Each fn As Microsoft.Office.Interop.Word.Footnote In rng.Document.Footnotes
                    If fn.Reference.Start >= rng.Start AndAlso fn.Reference.Start < rng.End Then
                        Dim s As Integer = System.Math.Max(fn.Reference.Start, rng.Start)
                        Dim e As Integer = System.Math.Min(fn.Reference.End, rng.End)
                        placeholders.Add(New PlaceholderInfo With {
                            .Offset = s - rng.Start,
                            .Length = e - s,
                            .Token = $"{{{{WFNT:{fn.Range.Text}}}}}"
                        })
                        processed += 1
                        ProgressScope.Report(System.Threading.Interlocked.Increment(current), label:=$"Footnotes … {processed}/{fnCount}")
                        If Cancelled() Then Exit For
                    End If
                Next
                If Cancelled() Then
                    ' Return partially processed content (insert what we can with current placeholders)
                    Dim fullTextCancel As String = rng.Text
                    Dim sbCancel As New System.Text.StringBuilder(fullTextCancel.Length + placeholders.Count * 16)
                    Dim lastPosCancel As Integer = 0
                    For Each ph In placeholders
                        If ph.Offset > lastPosCancel Then
                            sbCancel.Append(fullTextCancel.Substring(lastPosCancel, ph.Offset - lastPosCancel))
                        End If
                        sbCancel.Append(ph.Token)
                        lastPosCancel = ph.Offset + ph.Length
                    Next
                    If lastPosCancel < fullTextCancel.Length Then
                        sbCancel.Append(fullTextCancel.Substring(lastPosCancel))
                    End If
                    Return sbCancel.ToString()
                End If

                processed = 0
                For Each en As Microsoft.Office.Interop.Word.Endnote In rng.Document.Endnotes
                    If en.Reference.Start >= rng.Start AndAlso en.Reference.Start < rng.End Then
                        Dim s As Integer = System.Math.Max(en.Reference.Start, rng.Start)
                        Dim e As Integer = System.Math.Min(en.Reference.End, rng.End)
                        placeholders.Add(New PlaceholderInfo With {
                            .Offset = s - rng.Start,
                            .Length = e - s,
                            .Token = $"{{{{WENT:{en.Range.Text}}}}}"
                        })
                        processed += 1
                        ProgressScope.Report(System.Threading.Interlocked.Increment(current), label:=$"Endnotes … {processed}/{enCount}")
                        If Cancelled() Then Exit For
                    End If
                Next
                If Cancelled() Then
                    Dim fullTextCancel As String = rng.Text
                    Dim sbCancel As New System.Text.StringBuilder(fullTextCancel.Length + placeholders.Count * 16)
                    Dim lastPosCancel As Integer = 0
                    For Each ph In placeholders
                        If ph.Offset > lastPosCancel Then
                            sbCancel.Append(fullTextCancel.Substring(lastPosCancel, ph.Offset - lastPosCancel))
                        End If
                        sbCancel.Append(ph.Token)
                        lastPosCancel = ph.Offset + ph.Length
                    Next
                    If lastPosCancel < fullTextCancel.Length Then
                        sbCancel.Append(fullTextCancel.Substring(lastPosCancel))
                    End If
                    Return sbCancel.ToString()
                End If

                '──────────── Fields – determine entire field ──────────────
                Const WD_FIELD_BEGIN As Integer = 19   'Chr(19)
                Const WD_FIELD_END As Integer = 21     'Chr(21)

                processed = 0
                For Each fld As Word.Field In rng.Fields
                    Dim codeText As String = fld.Code.Text.Trim()

                    ' A) determine exact field begin
                    Dim fldStartAbs As Integer = fld.Code.Start
                    Do While fldStartAbs > rng.Start AndAlso
                        AscW(rng.Characters(fldStartAbs - rng.Start + 1).Text) <> WD_FIELD_BEGIN
                        fldStartAbs -= 1
                    Loop
                    If AscW(rng.Characters(fldStartAbs - rng.Start + 1).Text) <> WD_FIELD_BEGIN Then
                        processed += 1
                        ProgressScope.Report(System.Threading.Interlocked.Increment(current), label:=$"Fields … {processed}/{fieldsCount}")
                        If Cancelled() Then Exit For
                        Continue For
                    End If

                    ' B) search field end (0x15)
                    Dim scanAbs As Integer = fldStartAbs
                    Do While scanAbs < rng.End
                        Dim relIdx As Integer = scanAbs - rng.Start + 1
                        If AscW(rng.Characters(relIdx).Text) = WD_FIELD_END Then Exit Do
                        scanAbs += 1
                    Loop

                    If scanAbs >= rng.End Then
                        scanAbs = fld.Result.End
                        If scanAbs >= rng.End Then
                            processed += 1
                            ProgressScope.Report(System.Threading.Interlocked.Increment(current), label:=$"Fields … {processed}/{fieldsCount}")
                            If Cancelled() Then Exit For
                            Continue For
                        End If
                    End If

                    Dim fldEndAbs As Integer = scanAbs
                    Dim fldLength As Integer = fldEndAbs - fldStartAbs + 1

                    Dim token As String
                    If fld.Type = Word.WdFieldType.wdFieldHyperlink Then
                        Dim disp As String = StripSurroundingUTags(fld.Result.Text)
                        Dim dispB64 As String = System.Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(If(disp, String.Empty)))
                        token = $"{{{{WFLD:{codeText}|||{dispB64}}}}}"
                    Else
                        token = $"{{{{WFLD:{codeText}}}}}"
                    End If

                    placeholders.Add(New PlaceholderInfo With {
                        .Offset = fldStartAbs - rng.Start,
                        .Length = fldLength,
                        .Token = token
                    })

                    processed += 1
                    ProgressScope.Report(System.Threading.Interlocked.Increment(current), label:=$"Fields … {processed}/{fieldsCount}")
                    If Cancelled() Then Exit For
                Next
                If Cancelled() Then
                    Dim fullTextCancel As String = rng.Text
                    Dim sbCancel As New System.Text.StringBuilder(fullTextCancel.Length + placeholders.Count * 16)
                    Dim lastPosCancel As Integer = 0
                    For Each ph In placeholders
                        If ph.Offset > lastPosCancel Then
                            sbCancel.Append(fullTextCancel.Substring(lastPosCancel, ph.Offset - lastPosCancel))
                        End If
                        sbCancel.Append(ph.Token)
                        lastPosCancel = ph.Offset + ph.Length
                    Next
                    If lastPosCancel < fullTextCancel.Length Then
                        sbCancel.Append(fullTextCancel.Substring(lastPosCancel))
                    End If
                    Return sbCancel.ToString()
                End If

                '──────────── Paragraph placeholders (optional) ───────────────────────
                If PreserveParagraphFormatInline AndAlso rng.Paragraphs.Count > 0 Then
                    Dim paraCountLocal As Integer = rng.Paragraphs.Count
                    ReDim paragraphFormat(paraCountLocal - 1)
                    Array.Clear(paragraphFormat, 0, paragraphFormat.Length)

                    processed = 0
                    For i As Integer = 1 To paraCountLocal
                        Dim p As Word.Paragraph = rng.Paragraphs(i)

                        Dim fmt As New ParagraphFormatStructure With {
                            .Style = p.Style,
                            .FontName = p.Range.Font.Name,
                            .FontSize = p.Range.Font.Size,
                            .FontBold = p.Range.Font.Bold,
                            .FontItalic = p.Range.Font.Italic,
                            .FontUnderline = p.Range.Font.Underline,
                            .FontColor = p.Range.Font.Color,
                            .ListType = p.Range.ListFormat.ListType,
                            .ListTemplate = If(p.Range.ListFormat.ListType <> Word.WdListType.wdListNoNumbering,
                                               p.Range.ListFormat.ListTemplate, Nothing),
                            .ListLevel = If(p.Range.ListFormat.ListType <> Word.WdListType.wdListNoNumbering,
                                               p.Range.ListFormat.ListLevelNumber, 0),
                            .ListNumber = If(p.Range.ListFormat.ListType <> Word.WdListType.wdListNoNumbering,
                                               p.Range.ListFormat.ListValue, 0),
                            .HasListFormat = p.Range.ListFormat.ListType <> Word.WdListType.wdListNoNumbering,
                            .Alignment = p.Alignment,
                            .LineSpacing = p.LineSpacing,
                            .SpaceBefore = p.SpaceBefore,
                            .SpaceAfter = p.SpaceAfter
                        }

                        paragraphFormat(i - 1) = fmt

                        placeholders.Add(New PlaceholderInfo With {
                            .Offset = p.Range.Start - rng.Start,
                            .Length = 0,
                            .Token = $"{{{{PFOR:{i - 1}}}}}"
                        })

                        processed += 1
                        ProgressScope.Report(System.Threading.Interlocked.Increment(current), label:=$"Paragraph formats … {processed}/{paraCountLocal}")
                        If Cancelled() Then Exit For
                    Next

                    If Cancelled() Then
                        Dim fullTextCancel As String = rng.Text
                        Dim sbCancel As New System.Text.StringBuilder(fullTextCancel.Length + placeholders.Count * 16)
                        Dim lastPosCancel As Integer = 0
                        For Each ph In placeholders
                            If ph.Offset > lastPosCancel Then
                                sbCancel.Append(fullTextCancel.Substring(lastPosCancel, ph.Offset - lastPosCancel))
                            End If
                            sbCancel.Append(ph.Token)
                            lastPosCancel = ph.Offset + ph.Length
                        Next
                        If lastPosCancel < fullTextCancel.Length Then
                            sbCancel.Append(fullTextCancel.Substring(lastPosCancel))
                        End If
                        Return sbCancel.ToString()
                    End If
                End If

                '──────────── Sort placeholders (Offset ↑, Length ↑) ──────────
                Debug.WriteLine($"4-6 Range Start = {rng.Start} Selection Start = {Application.Selection.Start}")
                Debug.WriteLine($"Range End = {rng.End} Selection End = {Application.Selection.End}")

                placeholders.Sort(PlaceholderComparer)
                ProgressScope.Report(System.Threading.Interlocked.Increment(current), label:="Sorting placeholders …")
                If Cancelled() Then
                    Dim fullTextCancel As String = rng.Text
                    Dim sbCancel As New System.Text.StringBuilder(fullTextCancel.Length + placeholders.Count * 16)
                    Dim lastPosCancel As Integer = 0
                    For Each ph In placeholders
                        If ph.Offset > lastPosCancel Then
                            sbCancel.Append(fullTextCancel.Substring(lastPosCancel, ph.Offset - lastPosCancel))
                        End If
                        sbCancel.Append(ph.Token)
                        lastPosCancel = ph.Offset + ph.Length
                    Next
                    If lastPosCancel < fullTextCancel.Length Then
                        sbCancel.Append(fullTextCancel.Substring(lastPosCancel))
                    End If
                    Return sbCancel.ToString()
                End If

                Debug.WriteLine("placeholders: " & String.Join(", ", placeholders.Select(Function(ph) $"[Offset={ph.Offset}, Length={ph.Length}, Token={ph.Token}]")))

                ' Increase max to include per-placeholder building progress granularity
                Dim newMax As Integer = System.Math.Max(total, current + System.Math.Max(1, placeholders.Count) + 1)
                ProgressScope.Report(current, max:=newMax, label:="Building result …")

                ' ───── Insert placeholders ────────────────────────────────
                Dim fullText As String = rng.Text
                Dim sbInline As New System.Text.StringBuilder(fullText.Length + placeholders.Count * 16)
                Dim lastPos As Integer = 0

                Dim idx As Integer = 0
                For Each ph As PlaceholderInfo In placeholders
                    If ph.Offset > lastPos Then
                        sbInline.Append(fullText.Substring(lastPos, ph.Offset - lastPos))
                    End If

                    sbInline.Append(ph.Token)
                    lastPos = ph.Offset + ph.Length

                    idx += 1
                    ProgressScope.Report(System.Threading.Interlocked.Increment(current), label:=$"Building … {idx}/{placeholders.Count}")
                    If Cancelled() Then
                        ' Finish by appending the remainder, so partial output remains valid
                        If lastPos < fullText.Length Then
                            sbInline.Append(fullText.Substring(lastPos))
                        End If
                        Return sbInline.ToString()
                    End If
                Next

                If lastPos < fullText.Length Then
                    sbInline.Append(fullText.Substring(lastPos))
                End If

                Debug.WriteLine($"4-7 Range Start = {rng.Start} Selection Start = {Application.Selection.Start}")
                Debug.WriteLine($"Range End = {rng.End} Selection End = {Application.Selection.End}")

                ProgressScope.Report(System.Threading.Interlocked.Increment(current), label:="Done")
                Return sbInline.ToString()

            Catch ex As System.Exception
                Debug.WriteLine("Error in GetTextWithSpecialElementsInline: " & ex.Message)
                Return workingrange.Text
            Finally
                app.Options.CheckSpellingAsYouType = oldSpell
                app.Options.CheckGrammarAsYouType = oldGrammar
                app.ScreenUpdating = oldSU
                'splash.Close()
            End Try
        End Using
    End Function

    ''' <summary>
    ''' Removes outer <u> … </u> tags (case-insensitive) from the supplied string.
    ''' </summary>
    ''' <param name="input">Text that may contain underline tags.</param>
    ''' <returns>Inner text without wrapping underline tags.</returns>
    Private Shared Function StripSurroundingUTags(input As String) As String
        If String.IsNullOrEmpty(input) Then Return input
        Dim s = input.Trim()

        ' Fast path
        If s.StartsWith("<u>", StringComparison.OrdinalIgnoreCase) AndAlso s.EndsWith("</u>", StringComparison.OrdinalIgnoreCase) Then
            Return s.Substring(3, s.Length - 7)
        End If

        ' Robust path with optional inner whitespace/newlines
        Dim m = System.Text.RegularExpressions.Regex.Match(
        s,
        "^\s*<u>\s*(.*?)\s*</u>\s*$",
        System.Text.RegularExpressions.RegexOptions.Singleline Or System.Text.RegularExpressions.RegexOptions.IgnoreCase
    )
        If m.Success Then Return m.Groups(1).Value

        Return input
    End Function

    ''' <summary>
    ''' Generic helper that finds formatted fragments inside a range and wraps them with the provided markers.
    ''' </summary>
    ''' <param name="rng">Range to scan.</param>
    ''' <param name="configureFind">Action configuring Word.Find before execution.</param>
    ''' <param name="replacementText">String containing \1 or ^& tokens.</param>
    ''' <param name="tweakReplacement">Action clearing formatting on inserted tokens.</param>
    Private Sub ReplaceWithinRange(
    ByVal rng As Microsoft.Office.Interop.Word.Range,
    ByVal configureFind As System.Action(Of Microsoft.Office.Interop.Word.Find),
    ByVal replacementText As System.String,
    ByVal tweakReplacement As System.Action(Of Microsoft.Office.Interop.Word.Font))

        Dim includeHidden As Boolean = True ' handle hidden text properly during trimming

        If rng Is Nothing Then Return

        Dim doc As Microsoft.Office.Interop.Word.Document = rng.Document

        Dim originalStart As System.Int32 = rng.Start
        Dim allowedEnd As System.Int32 = rng.End
        Dim currentPosition As System.Int32 = originalStart

        Dim NormalizePart As System.Func(Of System.String, System.String) =
        Function(s As System.String) As System.String
            If System.String.IsNullOrEmpty(s) Then Return System.String.Empty
            Return s.Replace("^13", vbCr)
        End Function

        Dim isWildcard As System.Boolean = replacementText.Contains("\1")
        Dim isInline As System.Boolean = replacementText.Contains("^&")
        If Not isWildcard AndAlso Not isInline Then
            isInline = True
            replacementText = replacementText & "^&"
        End If

        ' Safety/Cancel/Progress guards  (minimal-invasive)
        Dim prevPosition As System.Int32 = -1
        Dim iterations As System.Int32 = 0
        Dim maxIterations As System.Int32 = System.Math.Max(100000, System.Math.Max(1, allowedEnd - originalStart) * 8)

        Do While currentPosition < allowedEnd
            ' Honor cancel quickly and keep UI responsive
            If ProgressBarModule.CancelOperation Then Exit Do
            System.Windows.Forms.Application.DoEvents()

            Dim searchRange As Microsoft.Office.Interop.Word.Range = doc.Range(Start:=currentPosition, [End]:=allowedEnd)
            Dim f As Microsoft.Office.Interop.Word.Find = searchRange.Find

            f.ClearFormatting()
            f.Replacement.ClearFormatting()
            configureFind(f)
            f.Forward = True
            f.Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
            f.Format = True

            If Not f.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceNone) Then
                Exit Do
            End If

            Dim foundStart As System.Int32 = searchRange.Start
            Dim foundEnd As System.Int32 = searchRange.End
            Dim foundText As System.String = searchRange.Text

            ' Zero-length or non-advancing matches → force advance 1 char
            If foundEnd <= foundStart Then
                currentPosition = System.Math.Min(System.Math.Max(foundStart, currentPosition) + 1, allowedEnd)
                GoTo ContinueLoop
            End If

            ' Table end-of-cell marker handling
            Dim endsWithCellMark As System.Boolean = foundText.EndsWith(vbCr & ChrW(7), System.StringComparison.Ordinal)
            Dim isCellMarkOnly As System.Boolean = (foundText = ChrW(7))
            Dim isCrOrCellBoundary As System.Boolean = (foundText = vbCr) OrElse (foundText = vbCr & ChrW(7)) OrElse isCellMarkOnly

            If isWildcard Then
                ' Handle patterns like "**\1**^13" – paragraph-like rule
                Dim rep As System.String = replacementText
                Dim parts = rep.Replace("^13", System.String.Empty).Split(New System.String() {"\1"}, 2, System.StringSplitOptions.None)
                Dim prefix As System.String = NormalizePart(If(parts.Length > 0, parts(0), System.String.Empty))
                Dim suffix As System.String = NormalizePart(If(parts.Length > 1, parts(1), System.String.Empty))
                Dim prefixLen As System.Int32 = prefix.Length
                Dim suffixLen As System.Int32 = suffix.Length

                Dim endsWithParaCr As System.Boolean = foundText.EndsWith(vbCr, System.StringComparison.Ordinal) OrElse endsWithCellMark
                Dim groupEnd As System.Int32 = foundEnd
                If endsWithCellMark Then
                    groupEnd = foundEnd - 2 ' strip vbCr + cell mark from the group
                ElseIf foundText.EndsWith(vbCr, System.StringComparison.Ordinal) Then
                    groupEnd = foundEnd - 1   ' strip vbCr from the group
                End If

                ' Trim trailing spaces/tabs from the group so tokens hug content
                While groupEnd > foundStart
                    'Dim ch As Char = doc.Range(Start:=groupEnd - 1, [End]:=groupEnd).Text(0)
                    Dim ch As Char = SafeGetSingleChar(OneCharRange(doc, groupEnd - 1, includeHidden))
                    If ch = ChrW(0) Then Exit While
                    If ch = " "c OrElse ch = vbTab Then
                        groupEnd -= 1
                    Else
                        Exit While
                    End If
                End While

                Dim projectedEnd As System.Int32 = foundEnd + prefixLen + suffixLen
                If projectedEnd > allowedEnd Then Exit Do

                ' 1) Unformat group content to avoid re-matching
                If tweakReplacement IsNot Nothing AndAlso groupEnd > foundStart Then
                    Dim contentRange As Microsoft.Office.Interop.Word.Range = doc.Range(Start:=foundStart, [End]:=groupEnd)
                    tweakReplacement(contentRange.Font)
                End If
                ' 2) Also unformat trailing marks (¶ or ¶+cell) so they don't match formatting rules
                If tweakReplacement IsNot Nothing AndAlso endsWithParaCr Then
                    Dim trailingLen As System.Int32 = If(endsWithCellMark, 2, 1)
                    Dim paraRange As Microsoft.Office.Interop.Word.Range = doc.Range(Start:=groupEnd, [End]:=groupEnd + trailingLen)
                    tweakReplacement(paraRange.Font)
                End If

                ' 3) Insert tokens around the trimmed group only
                If prefixLen > 0 AndAlso groupEnd >= foundStart Then
                    Dim groupRange As Microsoft.Office.Interop.Word.Range = doc.Range(Start:=foundStart, [End]:=groupEnd)
                    groupRange.InsertBefore(prefix)
                End If
                If suffixLen > 0 AndAlso groupEnd >= foundStart Then
                    Dim groupRange As Microsoft.Office.Interop.Word.Range = doc.Range(Start:=foundStart, [End]:=groupEnd)
                    groupRange.InsertAfter(suffix)
                End If

                ' 4) Ensure tokens themselves are not formatted
                If tweakReplacement IsNot Nothing Then
                    If prefixLen > 0 Then
                        Dim leadingTok As Microsoft.Office.Interop.Word.Range = doc.Range(Start:=foundStart, [End]:=foundStart + prefixLen)
                        tweakReplacement(leadingTok.Font)
                    End If
                    If suffixLen > 0 Then
                        Dim trailingStart As System.Int32 = groupEnd + prefixLen
                        Dim trailingTok As Microsoft.Office.Interop.Word.Range = doc.Range(Start:=trailingStart, [End]:=trailingStart + suffixLen)
                        tweakReplacement(trailingTok.Font)
                    End If
                End If

                currentPosition = foundEnd + prefixLen + suffixLen

            Else
                ' Inline formatting rule: e.g., "**^&**"
                ' Skip pure boundary matches – avoid "**¶**" and "**cellMark**"
                If isCrOrCellBoundary Then
                    If tweakReplacement IsNot Nothing Then tweakReplacement(searchRange.Font)
                    currentPosition = foundEnd
                    allowedEnd = rng.End
                    GoTo ContinueLoop
                End If

                Dim parts = replacementText.Split(New System.String() {"^&"}, 2, System.StringSplitOptions.None)
                Dim prefix As System.String = NormalizePart(If(parts.Length > 0, parts(0), System.String.Empty))
                Dim suffix As System.String = NormalizePart(If(parts.Length > 1, parts(1), System.String.Empty))
                Dim prefixLen As System.Int32 = prefix.Length
                Dim suffixLen As System.Int32 = suffix.Length

                ' Trim leading/trailing whitespace and cell-mark from the matched range
                Dim contentStart As System.Int32 = foundStart
                Dim contentEnd As System.Int32 = foundEnd
                Dim ft As System.String = foundText

                ' Handle end-of-cell marker pair (vbCr + Chr(7)) at the end
                If ft.EndsWith(vbCr & ChrW(7), System.StringComparison.Ordinal) Then
                    contentEnd -= 2
                    ft = ft.Substring(0, ft.Length - 2)
                End If
                ' Trim trailing paragraph mark
                If ft.EndsWith(vbCr, System.StringComparison.Ordinal) Then
                    contentEnd -= 1
                    ft = ft.Substring(0, ft.Length - 1)
                End If
                ' Trim trailing spaces/tabs
                While contentEnd > contentStart
                    'Dim ch As Char = doc.Range(Start:=contentEnd - 1, [End]:=contentEnd).Text(0)
                    Dim ch As Char = SafeGetSingleChar(OneCharRange(doc, contentEnd - 1, includeHidden))
                    If ch = ChrW(0) Then Exit While
                    If ch = " "c OrElse ch = vbTab Then
                        contentEnd -= 1
                    Else
                        Exit While
                    End If
                End While
                ' Trim leading spaces/tabs
                While contentStart < contentEnd
                    'Dim ch As Char = doc.Range(Start:=contentStart, [End]:=contentStart + 1).Text(0)
                    Dim ch As Char = SafeGetSingleChar(OneCharRange(doc, contentStart, includeHidden))
                    If ch = ChrW(0) Then Exit While
                    If ch = " "c OrElse ch = vbTab Then
                        contentStart += 1
                    Else
                        Exit While
                    End If
                End While

                ' If nothing remains after trimming, just advance
                If contentEnd <= contentStart Then
                    If tweakReplacement IsNot Nothing Then tweakReplacement(searchRange.Font)
                    currentPosition = foundEnd
                    allowedEnd = rng.End
                    GoTo ContinueLoop
                End If

                Dim projectedEnd As System.Int32 = foundEnd + prefixLen + suffixLen
                If projectedEnd > allowedEnd Then Exit Do

                ' 1) Unformat the matched content (only the trimmed content)
                If tweakReplacement IsNot Nothing Then
                    Dim contentRangeUF As Microsoft.Office.Interop.Word.Range = doc.Range(Start:=contentStart, [End]:=contentEnd)
                    tweakReplacement(contentRangeUF.Font)
                End If

                ' 2) Insert opening token at contentStart, then adjust indices
                If prefixLen > 0 Then
                    Dim openTokPos As Microsoft.Office.Interop.Word.Range = doc.Range(Start:=contentStart, [End]:=contentStart)
                    openTokPos.InsertBefore(prefix)
                    ' After inserting prefix, the content shifts right
                    contentStart += prefixLen
                    contentEnd += prefixLen
                End If

                ' 3) Insert closing token at the updated contentEnd
                If suffixLen > 0 Then
                    Dim closeTokPos As Microsoft.Office.Interop.Word.Range = doc.Range(Start:=contentEnd, [End]:=contentEnd)
                    closeTokPos.InsertAfter(suffix)
                End If

                ' 4) Unformat tokens too, using their actual positions
                If tweakReplacement IsNot Nothing Then
                    If prefixLen > 0 Then
                        Dim leadingTok As Microsoft.Office.Interop.Word.Range = doc.Range(Start:=contentStart - prefixLen, [End]:=contentStart)
                        tweakReplacement(leadingTok.Font)
                    End If
                    If suffixLen > 0 Then
                        Dim trailingTok As Microsoft.Office.Interop.Word.Range = doc.Range(Start:=contentEnd, [End]:=contentEnd + suffixLen)
                        tweakReplacement(trailingTok.Font)
                    End If
                End If

                ' Advance past the original found end plus tokens
                currentPosition = foundEnd + prefixLen + suffixLen
            End If

            allowedEnd = rng.End

ContinueLoop:
            ' Progress safety: ensure forward movement
            If currentPosition <= prevPosition Then
                currentPosition = System.Math.Min(currentPosition + 1, allowedEnd)
            End If
            prevPosition = currentPosition

            iterations += 1
            If iterations >= maxIterations Then Exit Do
        Loop

        rng.SetRange(Start:=originalStart, [End]:=allowedEnd)
    End Sub


    ''' <summary>
    ''' Gets the first character from a single-character range without throwing.
    ''' </summary>
    ''' <param name="r">A Word range expected to span one character.</param>
    ''' <returns>
    ''' The first character of <paramref name="r"/> when retrievable; otherwise <see cref="ChrW"/>(0) when no character is available
    ''' (for example because the character is hidden/special under the current text retrieval mode).
    ''' </returns>
    ''' <remarks>
    ''' This is a defensive helper to reduce COM-related exceptions and edge cases where <see cref="Word.Range.Text"/> can be empty.
    ''' </remarks>
    Private Shared Function SafeGetSingleChar(r As Word.Range) As Char
        If r Is Nothing Then Return ChrW(0)

        Dim t As String = r.Text
        If Not String.IsNullOrEmpty(t) Then Return t(0)

        ' Fallback: often works when Range.Text is Nothing for hidden/special content
        Try
            If r.Characters IsNot Nothing AndAlso r.Characters.Count >= 1 Then
                Dim ct As String = r.Characters(1).Text
                If Not String.IsNullOrEmpty(ct) Then Return ct(0)
            End If
        Catch
            ' ignore COM glitches; treat as non-retrievable
        End Try

        Return ChrW(0)
    End Function

    ''' <summary>
    ''' Creates a one-character range and aligns its text retrieval mode for hidden text.
    ''' </summary>
    ''' <param name="doc">Document used to create the range.</param>
    ''' <param name="startPos">Start position for the one-character range.</param>
    ''' <param name="includeHidden">When true, sets <see cref="Word.TextRetrievalMode.IncludeHiddenText"/> on the created range.</param>
    ''' <returns>A Word range spanning exactly one character.</returns>
    ''' <remarks>
    ''' Some ranges can throw when setting <see cref="Word.Range.TextRetrievalMode"/> depending on the story/context; errors are ignored.
    ''' </remarks>
    Private Shared Function OneCharRange(doc As Word.Document, startPos As Integer, includeHidden As Boolean) As Word.Range
        Dim r As Word.Range = doc.Range(Start:=startPos, [End]:=startPos + 1)
        Try
            r.TextRetrievalMode.IncludeHiddenText = includeHidden
            ' Keep default behavior for field codes unless you explicitly want them here.
            ' r.TextRetrievalMode.IncludeFieldCodes = False
        Catch
            ' ignore; some ranges can throw depending on context
        End Try
        Return r
    End Function

    ''' <summary>
    ''' Replaces highlighted portions with <mark> tags, optionally retaining paragraph breaks
    ''' and including highlight color metadata.
    ''' </summary>
    ''' <param name="rng">Range to process.</param>
    ''' <param name="keepParagraphBreakOutside">True to keep trailing paragraph marks outside the tag.</param>
    ''' <param name="includeColorInTag">True to emit data-ri-color attributes for non-yellow highlights.</param>
    Private Sub ReplaceWithinRange_Highlight(
    ByVal rng As Microsoft.Office.Interop.Word.Range,
    Optional ByVal keepParagraphBreakOutside As Boolean = False,
    Optional ByVal includeColorInTag As Boolean = False
)
        If rng Is Nothing Then Exit Sub

        Dim searchRng As Microsoft.Office.Interop.Word.Range = rng.Duplicate
        With searchRng.Find
            .ClearFormatting()
            .Text = ""
            .Format = True
            .Highlight = True
            .MatchWildcards = False
            .Forward = True
            .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
        End With

        While searchRng.Find.Execute()
            If ProgressBarModule.CancelOperation Then Exit While
            System.Windows.Forms.Application.DoEvents()

            Dim found As Microsoft.Office.Interop.Word.Range = searchRng.Duplicate
            Dim txt As System.String = found.Text
            If System.String.IsNullOrEmpty(txt) Then
                searchRng.Start = System.Math.Min(searchRng.Start + 1, rng.End)
                searchRng.End = rng.End
                If searchRng.Start >= rng.End Then Exit While
                Continue While
            End If

            Dim trailing As System.String = ""
            If keepParagraphBreakOutside Then
                If txt.EndsWith(vbCr & ChrW(7), System.StringComparison.Ordinal) Then
                    trailing = vbCr & ChrW(7) ' paragraph + cell
                    txt = txt.Substring(0, txt.Length - 2)
                ElseIf txt.EndsWith(vbCr, System.StringComparison.Ordinal) Then
                    trailing = vbCr
                    txt = txt.Substring(0, txt.Length - 1)
                End If
            End If

            ' Trim leading/trailing spaces/tabs so <mark> hugs the content
            ' (do this after removing the trailing paragraph/cell marker above)
            Dim startIdx As Integer = 0
            Dim endIdx As Integer = txt.Length
            While startIdx < endIdx AndAlso (txt(startIdx) = " "c OrElse txt(startIdx) = vbTab)
                startIdx += 1
            End While
            While endIdx > startIdx AndAlso (txt(endIdx - 1) = " "c OrElse txt(endIdx - 1) = vbTab)
                endIdx -= 1
            End While
            If startIdx > 0 OrElse endIdx < txt.Length Then
                txt = If(endIdx > startIdx, txt.Substring(startIdx, endIdx - startIdx), String.Empty)
            End If

            ' Capture color before clearing
            Dim hi As Microsoft.Office.Interop.Word.WdColorIndex = found.HighlightColorIndex

            Dim openTag As System.String = "<mark>"
            If includeColorInTag Then
                Dim suffix = HighlightIndexToMarkSuffix(hi)
                If suffix.Length > 0 Then
                    openTag = "<mark" & suffix & ">"
                End If
            End If

            found.Text = openTag & txt & "</mark>" & trailing

            ' Clear highlight from inserted text
            found.HighlightColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdNoHighlight

            searchRng.Start = found.End
            searchRng.End = rng.End
            If searchRng.Start >= rng.End Then Exit While
        End While
    End Sub

    ''' <summary>
    ''' Maps Word highlight colors to data-ri-color attribute suffixes.
    ''' </summary>
    ''' <param name="idx">Word highlight color index.</param>
    ''' <returns>Suffix appended to the opening &lt;mark&gt; tag or empty string.</returns>
    Private Shared Function HighlightIndexToMarkSuffix(idx As Word.WdColorIndex) As String
        ' Return a valid-HTML attribute suffix for use on the opening <mark ...> tag.
        ' Example usage (caller concatenates): "<mark" & suffix & ">"
        ' Note: Closing tag should remain plain "</mark>" (no attributes).
        Select Case idx
            Case Word.WdColorIndex.wdNoHighlight, Word.WdColorIndex.wdYellow : Return ""        ' default <mark>
            Case Word.WdColorIndex.wdBrightGreen : Return " data-ri-color=""brightgreen"""
            Case Word.WdColorIndex.wdTurquoise : Return " data-ri-color=""turquoise"""
            Case Word.WdColorIndex.wdPink : Return " data-ri-color=""pink"""
            Case Word.WdColorIndex.wdBlue : Return " data-ri-color=""blue"""
            Case Word.WdColorIndex.wdRed : Return " data-ri-color=""red"""
            Case Word.WdColorIndex.wdDarkBlue : Return " data-ri-color=""darkblue"""
            Case Word.WdColorIndex.wdTeal : Return " data-ri-color=""teal"""
            Case Word.WdColorIndex.wdGreen : Return " data-ri-color=""green"""
            Case Word.WdColorIndex.wdViolet : Return " data-ri-color=""violet"""
            Case Word.WdColorIndex.wdDarkRed : Return " data-ri-color=""darkred"""
            Case Word.WdColorIndex.wdDarkYellow : Return " data-ri-color=""darkyellow"""
            Case Word.WdColorIndex.wdGray25 : Return " data-ri-color=""gray25"""
            Case Word.WdColorIndex.wdGray50 : Return " data-ri-color=""gray50"""
            Case Word.WdColorIndex.wdBlack : Return " data-ri-color=""black"""
            Case Else : Return ""
        End Select
    End Function

    ''' <summary>
    ''' Restores serialized placeholder tokens within a range by recreating footnotes,
    ''' endnotes, fields, and paragraph formatting.
    ''' </summary>
    ''' <param name="workingrange">Range containing placeholder tokens to restore.</param>
    Private Sub RestoreSpecialTextElements(workingrange As Word.Range)

        Try
            Dim doc As Word.Document = Globals.ThisAddIn.Application.ActiveDocument

            Dim StartOfRange As Integer = workingrange.Start
            Dim EndOfRange As Integer = workingrange.End
            Dim EndOfDocument As Boolean = If(EndOfRange < doc.Content.End, False, True)
            If doc.Bookmarks.Exists("RTEX1") Then
                doc.Bookmarks("RTEX1").Delete()
            End If
            If Not EndOfDocument Then
                doc.Bookmarks.Add("RTEX1", doc.Range(EndOfRange, EndOfRange))
            End If

            ' Process Footnotes
            ProcessInTextPlaceholders(workingrange, doc, "WFNT:", AddressOf AddFootnote)
            workingrange.Start = StartOfRange
            If doc.Bookmarks.Exists("RTEX1") Then
                workingrange.End = doc.Bookmarks("RTEX1").Range.Start
            End If
            If EndOfDocument Then workingrange.End = doc.Content.End

            ' Process Endnotes
            ProcessInTextPlaceholders(workingrange, doc, "WENT:", AddressOf AddEndnote)
            workingrange.Start = StartOfRange
            If doc.Bookmarks.Exists("RTEX1") Then
                workingrange.End = doc.Bookmarks("RTEX1").Range.Start
            End If
            If EndOfDocument Then workingrange.End = doc.Content.End

            ' Process Fields
            ProcessInTextPlaceholders(workingrange, doc, "WFLD:", AddressOf AddField)
            workingrange.Start = StartOfRange
            If doc.Bookmarks.Exists("RTEX1") Then
                workingrange.End = doc.Bookmarks("RTEX1").Range.Start
            End If
            If EndOfDocument Then workingrange.End = doc.Content.End

            ' Process Formatting
            ProcessInTextPlaceholders(workingrange, doc, "PFOR:", AddressOf AddFormat)
            workingrange.Start = StartOfRange
            If doc.Bookmarks.Exists("RTEX1") Then
                workingrange.End = doc.Bookmarks("RTEX1").Range.Start
                doc.Bookmarks("RTEX1").Delete()
            End If
            If EndOfDocument Then workingrange.End = doc.Content.End

        Catch ex As System.Exception
            'MsgBox("An error occurred: " & ex.Message, MsgBoxStyle.Critical)
        End Try

    End Sub

    ''' <summary>
    ''' Searches a range for placeholders with the specified prefix and invokes a callback
    ''' to recreate the associated element (footnote, endnote, field, or format).
    ''' </summary>
    ''' <param name="workingrange">Range to scan.</param>
    ''' <param name="doc">Owning document.</param>
    ''' <param name="placeholderPrefix">Token prefix (e.g., WFNT:).</param>
    ''' <param name="addNoteAction">Callback that handles reinsertion.</param>
    Private Sub ProcessInTextPlaceholders(ByRef workingrange As Word.Range, doc As Word.Document, placeholderPrefix As String, addNoteAction As Action(Of Word.Document, Word.Range, String))
        Dim PreserveRange As Range = workingrange
        With workingrange.Find
            .Text = "\{\{" & placeholderPrefix & "*\}\}"
            .MatchWildcards = True
            Do While .Execute()
                Dim startPos As Integer = workingrange.Start
                Dim endPos As Integer = workingrange.End

                ' Extract note or field text by trimming prefix and suffix
                Dim placeholderText As String = workingrange.Text
                Dim noteText As String = placeholderText.Substring(placeholderPrefix.Length + 2, placeholderText.Length - (placeholderPrefix.Length + 4))

                ' Remove placeholder
                workingrange.Text = ""

                ' Add footnote, endnote, or field
                Dim insertionRange As Word.Range = doc.Range(startPos, startPos)
                addNoteAction.Invoke(doc, insertionRange, noteText)

                ' Adjust range position for the next match
                ' workingrange.Start = endPos + 1
                workingrange.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                If placeholderPrefix = "PFOR:" Then
                    workingrange.MoveStart(Unit:=Word.WdUnits.wdParagraph, Count:=1)
                    If workingrange.Start < doc.Content.End Then
                        Dim nextChar As String = doc.Range(workingrange.Start, workingrange.Start + 1).Text
                        ' In Word the paragraph mark is typically Chr(13) (which is vbCr)
                        If nextChar = vbCr Then
                            workingrange.MoveStart(Unit:=Word.WdUnits.wdCharacter, Count:=1)
                        End If
                        If nextChar = vbLf Then
                            workingrange.MoveStart(Unit:=Word.WdUnits.wdCharacter, Count:=1)
                        End If
                    End If
                Else
                    workingrange.MoveStart(Unit:=Word.WdUnits.wdCharacter, Count:=1)
                End If
            Loop
        End With
    End Sub

    ''' <summary>
    ''' Adds a footnote at the specified location.
    ''' </summary>
    Private Sub AddFootnote(doc As Word.Document, insertionRange As Word.Range, noteText As String)
        doc.Footnotes.Add(Range:=insertionRange, Text:=noteText)
    End Sub

    ''' <summary>
    ''' Adds an endnote at the specified location.
    ''' </summary>
    Private Sub AddEndnote(doc As Word.Document, insertionRange As Word.Range, noteText As String)
        doc.Endnotes.Add(Range:=insertionRange, Text:=noteText)
    End Sub

    ''' <summary>
    ''' Inserts a Word field and restores its code plus optional display text.
    ''' </summary>
    Private Sub AddField(doc As Word.Document, insertionRange As Word.Range, fieldText As String)
        ' fieldText may be either just the code
        ' or "code|||base64(displayText)" for hyperlink fields
        Dim fieldCode As String = fieldText
        Dim displayText As String = Nothing

        Dim parts = fieldText.Split(New String() {"|||"}, 2, StringSplitOptions.None)
        If parts.Length = 2 Then
            fieldCode = parts(0).Trim()
            Try
                displayText = System.Text.Encoding.UTF8.GetString(System.Convert.FromBase64String(parts(1)))
            Catch
                ' Fallback if it wasn't base64-encoded (keeps behavior robust)
                displayText = parts(1)
            End Try
        End If

        ' Insert the field and apply the code
        Dim fieldRange As Word.Range = insertionRange.Duplicate
        Dim field As Word.Field = doc.Fields.Add(fieldRange)
        field.Code.Text = fieldCode
        field.Update()

        ' If we captured a display text (e.g., from a hyperlink), restore it and lock the field
        If Not String.IsNullOrEmpty(displayText) Then
            field.Result.Text = displayText
            ' Keep the display text stable; remove this if you want updates to overwrite it later
            field.Locked = True
        End If
    End Sub

    ''' <summary>
    ''' Applies a captured paragraph format (style, font, list, spacing) to the paragraph
    ''' containing the specified insertion range.
    ''' </summary>
    Private Sub AddFormat(doc As Word.Document, insertionRange As Word.Range, formatIndexText As String)
        Try
            ' Parse the format index from the input text
            Dim formatIndex As Integer = Integer.Parse(formatIndexText.Trim())

            ' Ensure the format index is within bounds
            If formatIndex >= 0 AndAlso formatIndex < paragraphFormat.Length Then
                ' Retrieve the specific paragraph format
                Dim format = paragraphFormat(formatIndex)

                ' Expand the range to the entire paragraph
                Dim targetRange As Word.Range = insertionRange.Paragraphs(1).Range
                If targetRange.End > targetRange.Start Then
                    targetRange.End = targetRange.End - 1  ' Exclude the paragraph mark
                End If

                With targetRange
                    ' Apply the stored style
                    If format.Style IsNot Nothing Then .Style = format.Style

                    ' Apply the stored font formatting
                    With .Font
                        If format.FontName IsNot Nothing Then .Name = format.FontName
                        If format.FontSize > 0 Then .Size = format.FontSize
                        .Bold = format.FontBold
                        .Italic = format.FontItalic
                        .Underline = format.FontUnderline
                        .Color = format.FontColor
                    End With

                    ' Apply list formatting if applicable
                    If format.HasListFormat AndAlso format.ListTemplate IsNot Nothing Then
                        Try
                            .ListFormat.ApplyListTemplateWithLevel(
                            ListTemplate:=format.ListTemplate,
                            ContinuePreviousList:=If(format.ListNumber > 0, True, False),
                            ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList,
                            DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior
                        )
                            .ListFormat.ListLevelNumber = format.ListLevel
                        Catch ex As System.Exception
                            'MsgBox("Error applying list format: " & ex.Message, MsgBoxStyle.Exclamation)
                        End Try
                    End If

                    ' Apply paragraph alignment
                    .ParagraphFormat.Alignment = format.Alignment

                    ' Apply line spacing
                    .ParagraphFormat.LineSpacing = format.LineSpacing

                    ' Apply spacing before and after
                    .ParagraphFormat.SpaceBefore = format.SpaceBefore
                    .ParagraphFormat.SpaceAfter = format.SpaceAfter
                End With
            Else
                ' MsgBox("Invalid format index: " & formatIndex, MsgBoxStyle.Exclamation)
            End If
        Catch ex As System.Exception
            ' MsgBox("Error applying format: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub
End Class