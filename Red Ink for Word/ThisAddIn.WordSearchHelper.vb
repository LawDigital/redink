' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.WordSearchHelper.vb
' Purpose: Provides advanced text search functionality for long text fragments
'          in Word documents, handling text normalization, wildcard patterns,
'          and revision tracking.
'
' Architecture:
'  - Multi-Strategy Search: Implements four search strategies in fallback order:
'    1. Plain literal search (for strings ≤255 chars)
'    2. Escaped wildcard search with IgnoreSpace enabled
'    3. Anchored search using start/end word patterns
'    4. Windowed canonical text comparison with position backtracking
'  - Text Canonicalization: Normalizes text using Unicode KC form, removes noise
'    characters (control chars, zero-width spaces), collapses whitespace/hyphens,
'    converts German ß/ẞ to SS, and converts to uppercase.
'  - Position Mapping: Maintains character-to-document-position backmaps to
'    translate canonical string matches back to Word Range coordinates.
'  - Revision Handling: Temporarily switches to wdRevisionsViewFinal to exclude
'    deleted text from searches when skipDeleted is True.
'  - Word API Limits: Respects 255-character limit for Find.Text patterns;
'    dynamically reduces anchor word count when patterns exceed this limit.
'  - Timeout & Cancellation: Supports configurable timeout (default 10s) and
'    CancellationToken for responsive UI.
'  - Debug Support: Conditional debug output when ENABLE_SLICE_DEBUG is True
'    and debugger is attached.
'
' External Dependencies:
'  - Microsoft.Office.Interop.Word
'  - System.Text.RegularExpressions
'  - System.Threading (CancellationToken)
' =============================================================================

Option Explicit On
Option Strict Off

Public Module WordSearchHelper

    ' =========================================================================
    ' Module-Level Fields
    ' =========================================================================

    ''' <summary>
    ''' Regular expression to match Unicode escape sequences (\uXXXX or \uXXXXXX).
    ''' </summary>
    Private ReadOnly RX_U As New _
    System.Text.RegularExpressions.Regex("\\u([0-9A-Fa-f]{4,6})",
        System.Text.RegularExpressions.RegexOptions.Compiled Or
        System.Text.RegularExpressions.RegexOptions.CultureInvariant)

    ''' <summary>
    ''' Regular expression to match three consecutive periods (ellipsis).
    ''' </summary>
    Private ReadOnly RX_ELL As New _
    System.Text.RegularExpressions.Regex("\.\.\.",
        System.Text.RegularExpressions.RegexOptions.Compiled Or
        System.Text.RegularExpressions.RegexOptions.CultureInvariant)

    ''' <summary>
    ''' Toggle detailed per-slice logging when debugger is attached.
    ''' </summary>
    Private Const ENABLE_SLICE_DEBUG As Boolean = True

    ' =========================================================================
    ' Public Methods
    ' =========================================================================

    <System.Diagnostics.Conditional("DEBUG")>
    Private Sub LogHelperDiag(message As String)
        If String.IsNullOrWhiteSpace(message) Then Return
        System.Diagnostics.Debug.WriteLine($"{DateTime.Now:HH:mm:ss.fff} [FindLongTextAnchoredFast] {message}")
    End Sub

    ''' <summary>
    ''' Searches for a long text fragment in the specified Word selection using
    ''' multiple fallback strategies. Updates the selection to the found text.
    ''' </summary>
    ''' <param name="sel">Word selection object to search within and update.</param>
    ''' <param name="findText">Text fragment to find.</param>
    ''' <param name="skipDeleted">When True, excludes deleted revisions from search.</param>
    ''' <param name="nWords">Number of anchor words to use for start/end patterns (default 4).</param>
    ''' <param name="cancel">Cancellation token to abort long-running searches.</param>
    ''' <param name="timeoutSeconds">Maximum search duration in seconds (default 10).</param>
    ''' <param name="searchOriginal">When True, searches in Original view (hides inserted revisions)</param>
    ''' <returns>True if text was found and selection updated; False otherwise.</returns>
    ''' <remarks>
    ''' Search scope contract:
    '''   - If <paramref name="sel"/> has Start &lt; End, the helper searches inside that range.
    '''   - If <paramref name="sel"/> is collapsed (Start = End), the helper searches
    '''     from that caret position forward to the end of the main story.
    '''
    ''' Callers that want to scan the whole document MUST set the selection to
    ''' [doc.Content.Start, doc.Content.End] immediately before every call,
    ''' because intervening Word UI activity (DoEvents, Comments.Add, balloon
    ''' refresh, etc.) may collapse the live Selection between calls.
    ''' </remarks>
    Public Function FindLongTextAnchoredFast(
        ByRef sel As Microsoft.Office.Interop.Word.Selection,
        ByVal findText As System.String,
        Optional ByVal skipDeleted As System.Boolean = True,
        Optional ByVal nWords As System.Int32 = 4,
        Optional ByVal cancel As System.Threading.CancellationToken = Nothing,
        Optional ByVal timeoutSeconds As System.Int32 = 10,
        Optional ByVal searchOriginal As System.Boolean = False
    ) As System.Boolean

        Dim wordApp As Microsoft.Office.Interop.Word.Application = sel.Application

        ' Preserve original view settings
        Dim view = wordApp.ActiveWindow.View
        Dim origRevView = view.RevisionsView
        Dim origShowRev = view.ShowRevisionsAndComments
        Dim viewChanged1 As Boolean = False
        Dim viewChanged2 As Boolean = False

        ' Set the appropriate view for searching
        If searchOriginal Then
            ' Search against the original (baseline) text — hide insertions, show deletions
            If view.RevisionsView <> Microsoft.Office.Interop.Word.WdRevisionsView.wdRevisionsViewOriginal Then
                view.RevisionsView = Microsoft.Office.Interop.Word.WdRevisionsView.wdRevisionsViewOriginal
                viewChanged1 = True
            End If
            If view.ShowRevisionsAndComments Then
                view.ShowRevisionsAndComments = False
                viewChanged2 = True
            End If
        ElseIf skipDeleted Then
            ' Search against the accepted (final) text — hide deleted revisions
            If view.RevisionsView <> Microsoft.Office.Interop.Word.WdRevisionsView.wdRevisionsViewFinal Then
                view.RevisionsView = Microsoft.Office.Interop.Word.WdRevisionsView.wdRevisionsViewFinal
                viewChanged1 = True
            End If
            If view.ShowRevisionsAndComments Then
                view.ShowRevisionsAndComments = False
                viewChanged2 = True
            End If
        End If

        Dim timedOut As Boolean = False

        Try
            ' Debug tracking variables
            Dim _dbgLastSlice As System.String = ""
            Dim _dbgLastIdx As System.Int32 = -1

            Dim t0 As System.DateTime = System.DateTime.UtcNow

            Dim doc As Microsoft.Office.Interop.Word.Document = sel.Document
            Dim mainStory As Microsoft.Office.Interop.Word.Range =
                doc.StoryRanges(Microsoft.Office.Interop.Word.WdStoryType.wdMainTextStory).Duplicate

            LogHelperDiag($"ENTRY sel=[{sel.Range.Start},{sel.Range.End}] mainStory=[{mainStory.Start},{mainStory.End}] needleLen={findText.Length}")

            ' Determine search area.
            ' A collapsed selection must search from the caret position forward,
            ' not from the beginning of the document.
            Dim area As Microsoft.Office.Interop.Word.Range
            Dim sStart As System.Int32 = System.Math.Max(sel.Range.Start, mainStory.Start)
            Dim sEnd As System.Int32

            If sel.Range.Start = sel.Range.End Then
                sEnd = mainStory.End
            Else
                sEnd = System.Math.Min(sel.Range.End, mainStory.End)
            End If

            If sEnd < sStart Then sEnd = sStart
            area = doc.Range(Start:=sStart, End:=sEnd)

            LogHelperDiag($"AREA computed area=[{area.Start},{area.End}]")

            ' STRATEGY 0: Plain literal search (no wildcards, no IgnoreSpace)
            If findText.Length <= 255 Then
                Dim rngPlain As Microsoft.Office.Interop.Word.Range = area.Duplicate
                With rngPlain.Find
                    .ClearFormatting() : .Replacement.ClearFormatting()
                    .Font.Reset() : .ParagraphFormat.Reset()
                    .Text = findText
                    .Forward = True : .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
                    .MatchCase = False : .MatchWholeWord = False
                    .MatchWildcards = False : .Format = False
                    .IgnoreSpace = False
                End With
                Dim hitPlain As System.Boolean
                Try : hitPlain = rngPlain.Find.Execute() : Catch : hitPlain = False : End Try
                If hitPlain Then
                    If rngPlain.Start < area.Start OrElse rngPlain.End > area.End Then
                        LogHelperDiag($"STRATEGY 0 REJECTED out-of-range hit area=[{area.Start},{area.End}] hit=[{rngPlain.Start},{rngPlain.End}]")
                    Else
                        LogHelperDiag($"STRATEGY 0 HIT area=[{area.Start},{area.End}] hit=[{rngPlain.Start},{rngPlain.End}]")
                        sel.SetRange(rngPlain.Start, rngPlain.End)
                        Return True
                    End If
                Else
                    LogHelperDiag($"STRATEGY 0 MISS findTextLen={findText.Length}")
                End If
            End If

            ' STRATEGY 1: Escaped wildcard search with IgnoreSpace enabled
            Dim litPat As System.String = EscapeForWordWildcard(findText)
            If litPat.Length <= 255 Then
                Dim rngLit As Microsoft.Office.Interop.Word.Range = area.Duplicate
                With rngLit.Find
                    .ClearFormatting() : .Replacement.ClearFormatting()
                    .Font.Reset() : .ParagraphFormat.Reset()
                    .Text = litPat
                    .Forward = True : .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
                    .MatchCase = False : .MatchWildcards = True
                    .Format = False : .IgnoreSpace = True
                End With
                If rngLit.Find.Execute() Then
                    If rngLit.Start < area.Start OrElse rngLit.End > area.End Then
                        LogHelperDiag($"STRATEGY 1 REJECTED out-of-range hit area=[{area.Start},{area.End}] hit=[{rngLit.Start},{rngLit.End}]")
                    Else
                        LogHelperDiag($"STRATEGY 1 HIT area=[{area.Start},{area.End}] hit=[{rngLit.Start},{rngLit.End}]")
                        sel.SetRange(rngLit.Start, rngLit.End)
                        Return True
                    End If
                Else
                    LogHelperDiag($"STRATEGY 1 MISS litPatLen={litPat.Length}")
                End If
            End If

            ' STRATEGY 2: Prepare needle for anchored/canonical search
            Dim raw() As System.String = WordSearchHelper.RawWords(findText)
            Dim canonNeedle As System.String = Canonicalise(findText, True)
            If canonNeedle.Length = 0 Then
                Return False
            End If

            ' Try full wildcard pattern if it fits within 255-character limit
            Dim fullWildcardPattern As System.String = BuildWildcardProbe(raw)
            If fullWildcardPattern.Length <= 255 Then
                Dim rngFull As Microsoft.Office.Interop.Word.Range = area.Duplicate
                With rngFull.Find
                    .ClearFormatting() : .Replacement.ClearFormatting()
                    .Font.Reset() : .ParagraphFormat.Reset()
                    .Text = fullWildcardPattern
                    .Forward = True : .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
                    .MatchCase = False : .MatchWildcards = True
                    .Format = False : .IgnoreSpace = True
                End With
                ' A full wildcard probe such as "\[*\]" (produced for a bracketed
                ' token like "[TestA]") matches any bracketed run and would wrongly
                ' hit "[TestB]". Validate each candidate against the needle using the
                ' same exact canonical-equality contract that strategies 3 and 4 use,
                ' and skip forward on a mismatch so a later valid hit is not lost.
                ' Dash/whitespace tolerance is preserved because Canonicalise folds
                ' those variants away, so legitimate fuzzy matches still compare equal.
                Dim s2Found As System.Boolean : Try : s2Found = rngFull.Find.Execute() : Catch : s2Found = False : End Try
                Dim s2Guard As System.Int32 = 0
                Do While s2Found
                    s2Guard += 1
                    If s2Guard > 10000 Then Exit Do

                    If rngFull.Start < area.Start OrElse rngFull.End > area.End Then
                        LogHelperDiag($"STRATEGY 2 REJECTED out-of-range hit area=[{area.Start},{area.End}] hit=[{rngFull.Start},{rngFull.End}]")
                        Exit Do
                    End If

                    Dim canonFound As System.String = Canonicalise(rngFull.Text, True)
                    If System.String.Equals(canonFound, canonNeedle, System.StringComparison.Ordinal) Then
                        LogHelperDiag($"STRATEGY 2 HIT area=[{area.Start},{area.End}] hit=[{rngFull.Start},{rngFull.End}]")
                        sel.SetRange(rngFull.Start, rngFull.End)
                        Return True
                    End If

                    LogHelperDiag($"STRATEGY 2 REJECTED canonical mismatch hit=[{rngFull.Start},{rngFull.End}] canonFoundLen={canonFound.Length} canonNeedleLen={canonNeedle.Length}")

                    Dim s2NextStart As System.Int32 = rngFull.End
                    If s2NextStart <= rngFull.Start Then s2NextStart = rngFull.Start + 1
                    If s2NextStart >= area.End Then Exit Do

                    rngFull.SetRange(s2NextStart, area.End)
                    Try : s2Found = rngFull.Find.Execute() : Catch : s2Found = False : End Try
                Loop

                If Not s2Found Then
                    LogHelperDiag($"STRATEGY 2 MISS fullWildcardLen={fullWildcardPattern.Length}")
                End If
            End If

            If raw.Length < 2 Then Return False

            ' Adjust anchor word count to stay within 255-character limit
            nWords = System.Math.Min(nWords, raw.Length \ 2)
            If nWords < 1 Then nWords = 1

            Do While nWords > 1 AndAlso BuildWildcardProbe(raw.Take(nWords).ToArray()).Length > 255
                nWords -= 1
            Loop

            ' Build start and end anchor patterns
            Dim startPat As System.String = BuildWildcardProbe(raw.Take(nWords).ToArray())
            Dim endWords() As System.String = raw.Skip(raw.Length - nWords).ToArray()
            Dim endPat As System.String = BuildWildcardProbe(endWords)

            ' Calculate expected occurrence count for end pattern
            Dim occur As System.Int32 = CountOccurrences(findText, System.String.Join(" "c, endWords))
            If startPat = endPat Then occur = System.Math.Max(2, occur)

            LogHelperDiag($"STRATEGY 3 PREP nWords={nWords} startPatLen={startPat.Length} endPatLen={endPat.Length} occur={occur} canonNeedleLen={canonNeedle.Length}")

            ' STRATEGY 3: Anchored search using start/end patterns
            Using sRng As New RangeProxy(area.Duplicate)
                Dim fS As Microsoft.Office.Interop.Word.Find = sRng.Range.Find
                With fS
                    .ClearFormatting() : .Replacement.ClearFormatting()
                    .Font.Reset() : .ParagraphFormat.Reset()
                    .Text = startPat
                    .Forward = True : .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
                    .MatchCase = False : .MatchWildcards = True
                    .Format = False : .IgnoreSpace = True
                End With

                Dim okS As System.Boolean : Try : okS = fS.Execute() : Catch : okS = False : End Try
                LogHelperDiag($"STRATEGY 3 first start-anchor execute returned {okS}")
                Dim s3Iter As Integer = 0
                While okS
                    s3Iter += 1
                    LogHelperDiag($"STRATEGY 3 iter={s3Iter} startAnchor=[{sRng.Range.Start},{sRng.Range.End}]")
                    ' Check timeout
                    If (System.DateTime.UtcNow - t0).TotalSeconds > timeoutSeconds Then
                        timedOut = True : Exit While
                    End If
                    cancel.ThrowIfCancellationRequested()

                    Dim posStart As System.Int32 = sRng.Range.Start
                    Dim searchFrom As System.Int32 = sRng.Range.End

                    ' Search for end pattern from current position
                    Dim eRng As Microsoft.Office.Interop.Word.Range = doc.Range(searchFrom, area.End)
                    Dim fE As Microsoft.Office.Interop.Word.Find = eRng.Find
                    With fE
                        .ClearFormatting() : .Replacement.ClearFormatting()
                        .Font.Reset() : .ParagraphFormat.Reset()
                        .Text = endPat
                        .Forward = True : .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
                        .MatchCase = False : .MatchWildcards = True
                        .Format = False : .IgnoreSpace = True
                    End With
                    Dim okE As System.Boolean : Try : okE = fE.Execute() : Catch : okE = False : End Try

                    ' Find the nth occurrence of the end pattern
                    For i As System.Int32 = 2 To occur
                        If Not okE Then Exit For
                        eRng.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
                        Try : okE = fE.Execute() : Catch : okE = False : End Try
                    Next

                    If okE Then
                        ' Extract slice and build canonical form with position backmap
                        Dim sliceTxt As System.String
                        Dim back As System.Collections.Generic.IReadOnlyList(Of System.Int32)
                        VisibleSlice(doc, posStart, eRng.End - posStart, skipDeleted, sliceTxt, back)

                        If ENABLE_SLICE_DEBUG AndAlso System.Diagnostics.Debugger.IsAttached Then
                            System.Diagnostics.Debug.WriteLine(sliceTxt & System.Environment.NewLine)
                        End If

                        Dim canSlice As System.String
                        Dim backCanon As System.Collections.Generic.List(Of System.Int32)
                        CanonicaliseWithBackMap(sliceTxt, True, back, canSlice, backCanon)

                        Dim idx As System.Int32 = canSlice.IndexOf(canonNeedle, System.StringComparison.Ordinal)
                        _dbgLastSlice = canSlice
                        _dbgLastIdx = idx

                        LogHelperDiag($"STRATEGY 3 iter={s3Iter} canSliceLen={canSlice.Length} canonNeedleLen={canonNeedle.Length} idx={idx}")

                        If idx >= 0 Then
                            Dim endIdx As System.Int32 = System.Math.Min(idx + canonNeedle.Length - 1, backCanon.Count - 1)

                            ' Debug backmap integrity check
                            If System.Diagnostics.Debugger.IsAttached Then
                                For z As Integer = 1 To backCanon.Count - 1
                                    If backCanon(z) < backCanon(z - 1) Then
                                        System.Diagnostics.Debug.WriteLine(
                                                    "BACKMAP BREAK: index=" & z &
                                                    "  prev=" & backCanon(z - 1) &
                                                    "  curr=" & backCanon(z) &
                                                    "  Δ=" & (backCanon(z) - backCanon(z - 1))
                                                )
                                        Exit For
                                    End If
                                Next

                                System.Diagnostics.Debug.WriteLine("HIT idx=" & idx &
                                        " maps to Start=" & backCanon(idx) &
                                        " End=" & backCanon(endIdx))
                            End If

                            LogHelperDiag($"STRATEGY 3 HIT area=[{area.Start},{area.End}] hit=[{backCanon(idx)},{backCanon(endIdx) + 1}] posStart={posStart}")
                            sel.SetRange(backCanon(idx), backCanon(endIdx) + 1)
                            Return True
                        End If
                    End If

                    ' Move to next start pattern occurrence
                    sRng.CollapseEndPlusOne()
                    If sRng.Range.Start >= area.End Then Exit While
                    Try : okS = fS.Execute() : Catch : okS = False : End Try
                End While
            End Using

            ' STRATEGY 4: Windowed fallback using canonical text comparison
            Dim winSize As System.Int32 = 12000
            Dim overlap As System.Int32 = 400
            Dim p As System.Int32 = area.Start
            LogHelperDiag($"STRATEGY 4 START area=[{area.Start},{area.End}] winSize={winSize} overlap={overlap}")
            Dim s4Window As Integer = 0
            While p < area.End
                s4Window += 1
                ' Check timeout
                If (System.DateTime.UtcNow - t0).TotalSeconds > timeoutSeconds Then
                    timedOut = True : Exit While
                End If
                cancel.ThrowIfCancellationRequested()

                Dim len As System.Int32 = System.Math.Min(winSize, area.End - p)
                Dim sliceTxt As System.String
                Dim back As System.Collections.Generic.IReadOnlyList(Of System.Int32)
                VisibleSlice(doc, p, len, skipDeleted, sliceTxt, back)

                Dim canSlice As System.String
                Dim backCanon As System.Collections.Generic.List(Of System.Int32)
                CanonicaliseWithBackMap(sliceTxt, True, back, canSlice, backCanon)

                Dim idx As System.Int32 = canSlice.IndexOf(canonNeedle, System.StringComparison.Ordinal)
                LogHelperDiag($"STRATEGY 4 window#{s4Window} p={p} len={len} canSliceLen={canSlice.Length} idx={idx}")

                ' Production code with boundary checking
                If idx >= 0 Then
                    Dim endIdx As System.Int32 = System.Math.Min(idx + canonNeedle.Length - 1, backCanon.Count - 1)

                    Dim selStart As System.Int32 = backCanon(idx)
                    Dim selEnd As System.Int32 = backCanon(endIdx)

                    ' Enforce document boundaries
                    selStart = System.Math.Max(selStart, doc.Content.Start)
                    selEnd = System.Math.Min(selEnd + 1, doc.Content.End)

                    ' Validate range before selection
                    If selEnd <= selStart Then
                        Continue While
                    End If

                    Try
                        Dim testRange As Microsoft.Office.Interop.Word.Range = doc.Range(selStart, selEnd)

                        ' Extend range if necessary to capture full text
                        While testRange.Text.Length < findText.Length AndAlso testRange.End < doc.Content.End - 1
                            testRange.End = testRange.End + 1
                        End While

                        If testRange.Start >= doc.Content.Start AndAlso testRange.End <= doc.Content.End Then
                            LogHelperDiag($"STRATEGY 4 HIT area=[{area.Start},{area.End}] window p={p} hit=[{testRange.Start},{testRange.End}]")
                            sel.SetRange(testRange.Start, testRange.End)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(testRange)
                            Return True
                        End If

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(testRange)
                    Catch ex As System.Runtime.InteropServices.COMException
                        ' Continue searching on COM errors
                        If System.Diagnostics.Debugger.IsAttached Then
                            System.Diagnostics.Debug.WriteLine($"Range error at Start={selStart}, End={selEnd}: {ex.Message}")
                        End If
                    End Try
                End If

                ' Move to next window with overlap
                p += System.Math.Max(1, winSize - overlap)
            End While

            ' Debug output when all strategies fail
            If System.Diagnostics.Debugger.IsAttached Then
                Dim sliceLen As Integer = If(_dbgLastSlice Is Nothing, 0, _dbgLastSlice.Length)
                Dim containsStr As String
                If sliceLen = 0 OrElse canonNeedle.Length = 0 Then
                    containsStr = "n/a"
                Else
                    containsStr = _dbgLastSlice.Contains(canonNeedle).ToString()
                End If

                System.Diagnostics.Debug.WriteLine("===== FindLongTextAnchoredFast: FINAL DEBUG =====")
                System.Diagnostics.Debug.WriteLine("  findText        = '" & findText & "'")
                System.Diagnostics.Debug.WriteLine("  lastIdx         = " & _dbgLastIdx)
                System.Diagnostics.Debug.WriteLine("  needle.Length   = " & canonNeedle.Length)
                System.Diagnostics.Debug.WriteLine("  slice.Length    = " & sliceLen)
                System.Diagnostics.Debug.WriteLine("  contains?       = " & containsStr)
                Dim previewLen As System.Int32 = 200
                Dim startEx As System.String = If(sliceLen <= previewLen,
                                              _dbgLastSlice,
                                              _dbgLastSlice.Substring(0, previewLen) & "…")
                Dim endEx As System.String = If(sliceLen <= previewLen, "",
                                            "…" & _dbgLastSlice.Substring(sliceLen - previewLen))
                System.Diagnostics.Debug.WriteLine("  slice excerpt start: '" & startEx & "'")
                If endEx <> "" Then
                    System.Diagnostics.Debug.WriteLine("  slice excerpt end:   '" & endEx & "'")
                End If
                If timedOut Then
                    System.Diagnostics.Debug.WriteLine("  NOTE: search aborted due to timeout.")
                End If
                System.Diagnostics.Debug.WriteLine("===============================================")
            End If

            Return False

        Finally
            ' Restore original view settings
            If viewChanged1 Then view.RevisionsView = origRevView
            If viewChanged2 Then view.ShowRevisionsAndComments = origShowRev
        End Try
    End Function

    ' =========================================================================
    ' Private Helper Methods - Text Canonicalization
    ' =========================================================================

    ''' <summary>
    ''' Converts source text to canonical form with character-to-position backmap.
    ''' Applies Unicode normalization (KC), removes noise characters, collapses
    ''' whitespace/hyphens, converts ß/ẞ to SS, and converts to uppercase.
    ''' </summary>
    ''' <param name="src">Source text to canonicalize.</param>
    ''' <param name="collapseWS">When True, collapses consecutive whitespace to single space.</param>
    ''' <param name="backIn">Input position map (character index to document position).</param>
    ''' <param name="canonOut">Output canonical string with leading/trailing whitespace trimmed.</param>
    ''' <param name="canonBack">Output position map for canonical string.</param>
    Private Sub CanonicaliseWithBackMap(
        ByVal src As String,
        ByVal collapseWS As Boolean,
        ByVal backIn As System.Collections.Generic.IReadOnlyList(Of Integer),
        ByRef canonOut As String,
        ByRef canonBack As System.Collections.Generic.List(Of Integer))

        src = PrepareNeedle(src).Normalize(System.Text.NormalizationForm.FormKC)

        Dim sb As New System.Text.StringBuilder(src.Length)
        Dim back As New System.Collections.Generic.List(Of Integer)(src.Length)
        Dim pendingSpace As Boolean = False

        For i As Integer = 0 To src.Length - 1
            Dim ch As Char = src(i)
            If IsDocNoise(ch) Then Continue For

            Dim code As Integer = AscW(ch)
            ' Detect whitespace and hyphen variants
            Dim isHyphenOrWs As Boolean = System.Char.IsWhiteSpace(ch) OrElse code = &HA0
            If Not isHyphenOrWs Then
                Select Case code
                    Case &H2010, &H2011, &H2013, &H2014, &HAD, 45 ' Hyphens, en/em dash, soft hyphen
                        isHyphenOrWs = True
                End Select
            End If

            If isHyphenOrWs Then
                pendingSpace = True
            Else
                ' Emit pending space before next character
                If pendingSpace AndAlso collapseWS Then
                    sb.Append(" "c)
                    Dim mapIdx As Integer = System.Math.Min(i, System.Math.Max(backIn.Count - 1, 0))
                    back.Add(If(backIn.Count > 0, backIn(mapIdx), 0))
                End If
                pendingSpace = False

                ' Handle German ß/ẞ expansion to SS
                Select Case AscW(ch)
                    Case &HDF, &H1E9E ' ß, ẞ
                        sb.Append("S"c) : sb.Append("S"c)
                        Dim mi As Integer = System.Math.Min(i, System.Math.Max(backIn.Count - 1, 0))
                        Dim m As Integer = If(backIn.Count > 0, backIn(mi), 0)
                        back.Add(m) : back.Add(m)
                    Case Else
                        Dim up As Char = System.Char.ToUpperInvariant(ch)
                        sb.Append(up)
                        Dim mi As Integer = System.Math.Min(i, System.Math.Max(backIn.Count - 1, 0))
                        back.Add(If(backIn.Count > 0, backIn(mi), 0))
                End Select
            End If
        Next

        ' Trim leading and trailing whitespace
        Dim s As String = sb.ToString()
        Dim start As Integer = 0
        While start < s.Length AndAlso System.Char.IsWhiteSpace(s.Chars(start))
            start += 1
        End While
        Dim [end] As Integer = s.Length
        While [end] > start AndAlso System.Char.IsWhiteSpace(s.Chars([end] - 1))
            [end] -= 1
        End While

        canonOut = If([end] > start, s.Substring(start, [end] - start), System.String.Empty)
        canonBack = If([end] > start, back.GetRange(start, [end] - start), New System.Collections.Generic.List(Of Integer)())

        If System.Diagnostics.Debugger.IsAttached Then
            System.Diagnostics.Debug.WriteLine("canonOut=""" & canonOut & """")
            System.Diagnostics.Debug.WriteLine("canonBack length=" & canonBack.Count)
        End If
    End Sub

    ''' <summary>
    ''' Converts text to canonical form without position mapping.
    ''' </summary>
    ''' <param name="src">Source text to canonicalize.</param>
    ''' <param name="collapseWS">When True, collapses consecutive whitespace to single space.</param>
    ''' <returns>Canonical string with leading/trailing whitespace trimmed.</returns>
    Private Function Canonicalise(ByVal src As String, ByVal collapseWS As Boolean) As String
        src = PrepareNeedle(src).Normalize(System.Text.NormalizationForm.FormKC)

        Dim sb As New System.Text.StringBuilder(src.Length)
        Dim pendingSpace As Boolean = False

        For Each ch As Char In src
            If IsDocNoise(ch) Then Continue For

            Dim code As Integer = AscW(ch)
            Dim isHyphenOrWs As Boolean = System.Char.IsWhiteSpace(ch) OrElse code = &HA0
            If Not isHyphenOrWs Then
                Select Case code
                    Case &H2010, &H2011, &H2013, &H2014, &HAD, 45
                        isHyphenOrWs = True
                End Select
            End If

            If isHyphenOrWs Then
                pendingSpace = True
            Else
                If pendingSpace AndAlso collapseWS Then sb.Append(" "c)
                pendingSpace = False
                sb.Append(CanonizeDocChar(ch))
            End If
        Next
        Return sb.ToString().Trim()
    End Function

    ''' <summary>
    ''' Prepares text by converting Unicode escapes (\uXXXX) to actual characters
    ''' and replacing three-dot sequences with ellipsis character (U+2026).
    ''' </summary>
    ''' <param name="txt">Text to prepare.</param>
    ''' <returns>Prepared text.</returns>
    Private Function PrepareNeedle(ByVal txt As String) As String
        txt = RX_U.Replace(txt,
            Function(m) System.Char.ConvertFromUtf32(
                System.Convert.ToInt32(m.Groups(1).Value, 16)))
        Return RX_ELL.Replace(txt, ChrW(&H2026))
    End Function

    ''' <summary>
    ''' Determines whether a character should be excluded from canonical text.
    ''' Excludes control characters (except tab/LF/CR), zero-width spaces,
    ''' direction marks, object replacement characters, and Word-specific codes.
    ''' </summary>
    ''' <param name="ch">Character to test.</param>
    ''' <returns>True if character should be excluded; False otherwise.</returns>
    Private Function IsDocNoise(ByVal ch As Char) As Boolean
        Dim code As Integer = AscW(ch)
        If code < 32 AndAlso code <> 9 AndAlso code <> 10 AndAlso code <> 13 Then Return True
        Select Case code
            Case &HA0, &H200B, &H200C, &H200D, &H2060, ' NBSP, zero-width spaces, word joiner
                 &H200E To &H200F, &H202A To &H202E,   ' Direction marks, embedding controls
                 1, 19, 20, 21, &HFFFA, &HFFFB, &HFFFC ' Word field codes, object replacement
                Return True
        End Select
        Return False
    End Function

    ''' <summary>
    ''' Converts a character to its canonical form for case-insensitive comparison.
    ''' Handles German ß/ẞ (converts to "SS"); all others converted to uppercase.
    ''' </summary>
    ''' <param name="ch">Character to canonicalize.</param>
    ''' <returns>Canonical representation (single uppercase character or "SS").</returns>
    Private Function CanonizeDocChar(ByVal ch As Char) As String
        Select Case AscW(ch)
            Case &HDF, &H1E9E : Return "SS" ' ß, ẞ
            Case Else : Return System.Char.ToUpperInvariant(ch)
        End Select
    End Function

    ' =========================================================================
    ' Private Helper Methods - Wildcard Pattern Construction
    ' =========================================================================

    ''' <summary>
    ''' Constructs a Word wildcard pattern from an array of words. Escapes
    ''' special characters, replaces hyphens with "?" to match variants,
    ''' and handles bracketed content by replacing with "\[*\]".
    ''' </summary>
    ''' <param name="words">Array of words to convert to pattern.</param>
    ''' <returns>Word wildcard pattern string.</returns>
    Private Function BuildWildcardProbe(ByVal words() As String) As String
        Dim sb As New System.Text.StringBuilder(words.Length * 14)
        Dim i As Integer = 0
        While i < words.Length
            If i > 0 Then sb.Append(" "c)

            Dim w As String = words(i)
            If w.Contains("["c) Then
                ' Skip forward to closing bracket
                While i < words.Length AndAlso Not words(i).Contains("]"c)
                    i += 1
                End While
                sb.Append("\[*\]")
                If i < words.Length Then
                    Dim rest As String = words(i).Substring(words(i).IndexOf("]"c) + 1)
                    If rest <> "" Then sb.Append(EscapeForWordWildcard(rest))
                End If
            Else
                ' Replace hyphen variants with wildcard "?"
                w = w.Replace("-"c, "?"c) _
                     .Replace(ChrW(&H2010), "?"c) _
                     .Replace(ChrW(&H2011), "?"c) _
                     .Replace(ChrW(&H2013), "?"c) _
                     .Replace(ChrW(&H2014), "?"c) _
                     .Replace(ChrW(&HAD), "?"c)
                sb.Append(EscapeForWordWildcard(w))
            End If
            i += 1
        End While
        Return sb.ToString()
    End Function

    ''' <summary>
    ''' Escapes special wildcard characters for Word's Find.Text pattern.
    ''' Escapes: ? * @ [ ] ( ) { } \ < >
    ''' </summary>
    ''' <param name="s">String to escape.</param>
    ''' <returns>Escaped string safe for wildcard patterns.</returns>
    Private Function EscapeForWordWildcard(ByVal s As String) As String
        If s = "" Then Return ""
        Dim sb As New System.Text.StringBuilder(s.Length * 2)
        For Each ch As Char In s
            Select Case ch
                Case "?"c, "*"c, "@"c, "["c, "]"c, "("c, ")"c,
                     "{"c, "}"c, "\"c, "<"c, ">"c
                    sb.Append("\"c)
            End Select
            sb.Append(ch)
        Next
        Return sb.ToString()
    End Function

    ' =========================================================================
    ' Private Helper Methods - Text Extraction & Analysis
    ' =========================================================================

    ''' <summary>
    ''' Extracts text slice from document and builds character-to-position map.
    ''' </summary>
    ''' <param name="doc">Word document to extract from.</param>
    ''' <param name="absStart">Absolute starting position in document.</param>
    ''' <param name="sliceLen">Requested length of slice.</param>
    ''' <param name="skipDeleted">Currently unused parameter.</param>
    ''' <param name="visOut">Output text slice.</param>
    ''' <param name="mapBack">Output map from string index to document position.</param>
    Private Sub VisibleSlice(
    ByVal doc As Microsoft.Office.Interop.Word.Document,
    ByVal absStart As System.Int32,
    ByVal sliceLen As System.Int32,
    ByVal skipDeleted As System.Boolean,
    ByRef visOut As System.String,
    ByRef mapBack As System.Collections.Generic.IReadOnlyList(Of System.Int32))

        ' Extract raw text with safety margin
        Dim rawEnd As System.Int32 = System.Math.Min(doc.Content.End, absStart + sliceLen + 500)
        Dim rawRng As Microsoft.Office.Interop.Word.Range = doc.Range(absStart, rawEnd)
        Dim rawTxt As System.String = rawRng.Text
        Dim rawLen As System.Int32 = rawTxt.Length

        Dim take As System.Int32 = System.Math.Min(sliceLen, rawLen)
        If take < 0 Then take = 0
        visOut = If(rawLen > take, rawTxt.Substring(0, take), rawTxt)

        ' Build position mapping for each character
        Dim backList As New System.Collections.Generic.List(Of System.Int32)(visOut.Length)

        If visOut.Length > 0 Then
            Dim tempRng As Microsoft.Office.Interop.Word.Range = doc.Range(absStart, absStart)
            Dim currentPos As System.Int32 = absStart

            For i As System.Int32 = 0 To visOut.Length - 1
                backList.Add(currentPos)

                ' Advance to next character position
                If i < visOut.Length - 1 Then
                    Try
                        tempRng.SetRange(currentPos, rawEnd)
                        tempRng.MoveStart(Microsoft.Office.Interop.Word.WdUnits.wdCharacter, 1)
                        currentPos = tempRng.Start
                    Catch
                        ' Fallback to linear increment
                        currentPos += 1
                    End Try
                End If
            Next

            System.Runtime.InteropServices.Marshal.ReleaseComObject(tempRng)
        End If

        mapBack = backList
    End Sub

    ''' <summary>
    ''' Splits text into words after preparing Unicode escapes and ellipses.
    ''' Splits on space, tab, LF, and CR.
    ''' </summary>
    ''' <param name="src">Source text to split.</param>
    ''' <returns>Array of words (empty entries removed).</returns>
    Private Function RawWords(ByVal src As String) As String()
        src = RX_U.Replace(src, Function(m) _
            System.Char.ConvertFromUtf32(System.Convert.ToInt32(m.Groups(1).Value, 16)))
        src = RX_ELL.Replace(src, ChrW(&H2026))
        src = src.Normalize(System.Text.NormalizationForm.FormKC)
        Return src.Split(New Char() {" "c, ChrW(9), ChrW(10), ChrW(13)},
                         System.StringSplitOptions.RemoveEmptyEntries)
    End Function

    ''' <summary>
    ''' Counts occurrences of a substring within text (case-insensitive,
    ''' using canonical form).
    ''' </summary>
    ''' <param name="txt">Text to search within.</param>
    ''' <param name="subTxt">Substring to find.</param>
    ''' <returns>Number of non-overlapping occurrences.</returns>
    Private Function CountOccurrences(ByVal txt As String, ByVal subTxt As String) As Integer
        txt = Canonicalise(txt, True)
        subTxt = Canonicalise(subTxt, True)
        Dim cnt As Integer = 0
        Dim pos As Integer = txt.IndexOf(subTxt, System.StringComparison.OrdinalIgnoreCase)
        While pos <> -1
            cnt += 1
            pos = txt.IndexOf(subTxt, pos + subTxt.Length, System.StringComparison.OrdinalIgnoreCase)
        End While
        Return cnt
    End Function

    ' =========================================================================
    ' Private Helper Class
    ' =========================================================================

    ''' <summary>
    ''' Wrapper for Word Range that implements IDisposable to ensure COM cleanup.
    ''' </summary>
    Private NotInheritable Class RangeProxy
        Implements System.IDisposable

        Friend ReadOnly Range As Microsoft.Office.Interop.Word.Range
        Private ReadOnly ptr As Object

        ''' <summary>
        ''' Initializes a new RangeProxy wrapping the specified Word Range.
        ''' </summary>
        ''' <param name="r">Word Range to wrap.</param>
        Friend Sub New(ByVal r As Microsoft.Office.Interop.Word.Range)
            Range = r
            ptr = r
        End Sub

        ''' <summary>
        ''' Collapses range to end position and advances by one character.
        ''' </summary>
        Friend Sub CollapseEndPlusOne()
            Range.Collapse(
                Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
            Range.SetRange(Range.Start + 1, Range.Start + 1)
        End Sub

        ''' <summary>
        ''' Releases the underlying COM object.
        ''' </summary>
        Public Sub Dispose() Implements System.IDisposable.Dispose
            If ptr IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ptr)
            End If
        End Sub
    End Class

End Module
