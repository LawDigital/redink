' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.Processing.SurgicalInsert.vb
' Purpose: Implements the surgical tracked-change insertion engine used by the
'          Word processing pipeline. It diffs original serialized text against
'          revised output and replays only localized edits into the live Word
'          document while preserving existing structure wherever possible.
'
' Responsibilities:
'  - Preserve editor state, tracking state, and local formatting-sensitive context
'    while applying surgical markup.
'  - Tokenize and restore protected placeholders for fields, footnotes, endnotes,
'    paragraph-format markers, and structural break characters.
'  - Build word-level and sentence-aware diff runs, including patience-diff fallback
'    behavior for difficult rewrite regions.
'  - Map diff output onto a live visible atom stream that excludes tracked deletions
'    and keeps exact Word position spans for safe edit planning.
'  - Plan tracked deletions/insertions as pending edits and apply them right-to-left
'    to avoid position drift.
'  - Preserve boundary whitespace and original break-character patterns, and perform
'    localized cleanup/debug tracing after edits are applied.
'
' Architecture:
'  - `CompareAndInsertSurgical` is the main orchestration routine for placeholder
'    tokenization, break normalization, diff construction, live-range alignment,
'    pending-edit planning, and final tracked application.
'  - Diff construction is layered: tokenized placeholder/break protection first,
'    then sentence-aware collapse rules, then word-level/patience diff where needed.
'  - Application is atom-map-based rather than `Word.Find`-driven, so unchanged text
'    is aligned against the live visible document stream with exact `[Start, End)`
'    spans before any edit is executed.
'  - Final edits are applied in reverse order and followed by local whitespace cleanup
'    to keep revisions readable and structurally stable.
'
' External Dependencies:
'  - Microsoft.Office.Interop.Word for live range inspection, tracked revisions,
'    field/revision enumeration, and document mutation.
'  - DiffPlex for inline diff/LCS fallback used during token-level comparison.
'  - SharedLibrary.SharedMethods for splash/UI helpers and shared add-in utilities.
'  - System.Windows.Forms for UI message pumping during long-running surgical passes.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports DiffPlex
Imports DiffPlex.DiffBuilder
Imports DiffPlex.DiffBuilder.Model
Imports DocumentFormat.OpenXml
Imports Markdig
Imports Microsoft.Office.Interop.Word
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ' Collapse only clear rewrites. The earlier 0.82 similarity-only threshold treated ordinary
    ' editing such as "scheduled to depart on" -> "scheduled for" and "more than" -> "over" as
    ' sentence replacement. A sentence should be replaced wholesale only when both the common-token
    ' similarity is low enough and the changed-token mass is large enough.
    Const materialRewriteSimilarityThreshold As Double = 0.72
    Const materialRewriteChangedTokenRatioThreshold As Double = 0.28
    Const materialRewriteMinimumChangedTokens As Integer = 10    ' ========================== Surgical Markup (MarkupMethod 2 & 5) ==========================

    ''' <summary>
    ''' Represents a single word-level run from the DiffPlex output: Unchanged, Inserted, or Deleted.
    ''' Consecutive words of the same ChangeType are merged into a single run.
    ''' </summary>
    Private Structure DiffRun
        Public RunType As DiffPlex.DiffBuilder.Model.ChangeType
        Public Text As String
    End Structure

    Private Enum SpecialPlaceholderKind
        None = 0
        WFLD = 1
        WFNT = 2
        WENT = 3
        PFOR = 4
    End Enum

    ''' <summary>
    ''' Applies an in-place tracked-change patch to <paramref name="targetRange"/> by diffing the
    ''' original serialized text against the revised LLM output and then replaying only the minimal
    ''' edits into the live Word document.
    '''
    ''' This routine is the add-in's "surgical" markup engine. Unlike the legacy markup paths
    ''' (`InsertMarkupText`, full-range replacement, or Word compare-document workflows), it does
    ''' not rebuild the entire selected range. Instead, it keeps the existing Word content in place
    ''' and mutates only the regions that actually changed. This is important because a full rebuild
    ''' can destroy or destabilize live Word constructs such as:
    '''
    ''' - existing tracked revisions outside the changed fragment,
    ''' - fields and merge-field placeholders,
    ''' - footnote and endnote reference anchors,
    ''' - paragraph/list structure,
    ''' - manual line breaks and paragraph breaks,
    ''' - formatting in unchanged runs.
    '''
    ''' High-level algorithm:
    '''
    ''' 1. Preserve editor state
    '''    Screen updating and document tracking are temporarily adjusted so the patch can be applied
    '''    deterministically and without visual flicker. The original Word state is restored in
    '''    <c>Finally</c>.
    '''
    ''' 2. Tokenize protected placeholders
    '''    Special inline placeholders such as <c>{{WFLD:...}}</c>, <c>{{WFNT:...}}</c>,
    '''    <c>{{WENT:...}}</c>, and <c>{{PFOR:...}}</c> are first replaced with stable synthetic
    '''    tokens shared between original and revised text. This prevents DiffPlex from splitting
    '''    placeholder internals and allows them to survive round-trips through the diff pipeline.
    '''
    ''' 3. Normalize structural breaks into explicit tokens
    '''    Paragraph breaks, line feeds, and manual line breaks are converted to explicit diff tokens
    '''    before tokenization. This allows the diff engine to reason about structural separators as
    '''    first-class units instead of burying them inside word tokens. After diffing, those tokens
    '''    are converted back to their original Word representations.
    '''
    ''' 4. Build a word-oriented diff
    '''    The method uses DiffPlex <c>InlineDiffBuilder</c> on token streams rather than raw text.
    '''    Tokens are joined line-by-line, diffed, and then reassembled into <c>DiffRun</c> items.
    '''    Consecutive runs of the same <c>ChangeType</c> are merged, except for structural tokens
    '''    such as whitespace, placeholders, and break tokens, which are intentionally kept isolated
    '''    so Word-boundary-sensitive edits can be handled safely.
    '''
    ''' 5. Restore placeholder text and break characters
    '''    Once the raw diff runs have been built, synthetic placeholder tokens are restored to their
    '''    original placeholder text and break markers are converted back to live Word characters.
    '''
    ''' 6. Skip dangerous no-op structural replacements
    '''    Break-only replacement clusters are explicitly ignored. This protects manual line breaks,
    '''    paragraph breaks, and bullet/list structure from being rewritten merely because the model
    '''    returned a different textual newline convention.
    '''
    ''' 7. Walk the live Word range from left to right
    '''    A duplicate live cursor is collapsed to the start of <paramref name="targetRange"/> and
    '''    then advanced through the document while each diff run is processed:
    '''
    '''    - Unchanged text is used as an anchor to advance the cursor.
    '''    - Placeholder runs are advanced using Word-aware placeholder handling.
    '''    - Whitespace runs are matched character-by-character.
    '''    - Break runs are matched against Word's internal break representation.
    '''    - Text runs are located via <c>Range.Find</c> with normalization suitable for Word.
    '''
    '''    The cursor is never advanced blindly unless anchor matching fails, in which case a bounded
    '''    estimate is used as a fallback to preserve forward progress.
    '''
    ''' 8. Apply change clusters as tracked deletions/insertions
    '''    Adjacent non-unchanged runs are grouped into a single cluster. Each cluster is converted
    '''    into one or more <c>SurgicalOperationCandidate</c> instances representing possible ways
    '''    to perform the edit in Word.
    '''
    '''    Replacement clusters and pure deletions are handled differently:
    '''
    '''    - Replacements prefer consuming a trailing whitespace run together with the replaced word
    '''      so that Word does not swallow or duplicate separators around the tracked change.
    '''    - Pure deletions absorb adjacent whitespace only when the deleted content is word-like;
    '''      punctuation-only deletions intentionally avoid absorbing surrounding spaces.
    '''
    '''    Each candidate is searched in the live document beginning at the current cursor position.
    '''    On the first successful match, the method performs a tracked deletion and then, if needed,
    '''    a tracked insertion at the same logical point.
    '''
    ''' 9. Final visual cleanup
    '''    After all clusters have been processed, the affected region is temporarily viewed in Final
    '''    view and repeated double spaces are collapsed in the rendered result. This cleanup is kept
    '''    local to the edited range.
    '''
    ''' Cancellation / responsiveness:
    '''
    ''' - The loop periodically pumps the Windows message queue via <c>Application.DoEvents()</c>.
    ''' - <c>Esc</c> is checked repeatedly so the user can abort a long-running patch.
    '''
    ''' Design goals:
    '''
    ''' - preserve unchanged Word content exactly where possible,
    ''' - minimize COM mutations,
    ''' - keep revisions readable and localized,
    ''' - protect placeholders and structural elements,
    ''' - avoid rebuilding the entire selection.
    '''
    ''' Limitations:
    '''
    ''' Because Word's live text model is not a plain string, exact anchor matching can still fail in
    ''' the presence of complex revisions, fields, list formatting, or unexpected whitespace/break
    ''' normalization by the model. In such cases the method falls back conservatively and logs the
    ''' decision through the debug trace instead of forcing a potentially destructive rewrite.
    ''' </summary>
    ''' <param name="text1">Original serialized text seen by the diff engine.</param>
    ''' <param name="text2">Revised serialized text returned by the LLM after post-processing.</param>
    ''' <param name="targetRange">Live Word range to be patched in place with tracked revisions.</param>
    ''' <param name="trailingCR">True when the original source ended in a trailing paragraph break that must be preserved.</param>
    Private Sub CompareAndInsertSurgical(ByVal text1 As String, ByVal text2 As String, ByVal targetRange As Microsoft.Office.Interop.Word.Range, Optional ByVal trailingCR As Boolean = False)
        Dim wordApp As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application
        Dim doc As Microsoft.Office.Interop.Word.Document = targetRange.Document

        Dim originalTrack As Boolean = doc.TrackRevisions
        Dim originalUpdate As Boolean = wordApp.ScreenUpdating
        Dim originalSmartCutPaste As Boolean = wordApp.Options.SmartCutPaste

        Const uiYieldIntervalMs As Integer = 40
        Dim manualLineBreak As String = ChrW(11).ToString()

        Dim splash As SLib.SplashScreen = Nothing

        Try
            splash = New SLib.SplashScreen("Applying markup... press 'Esc' to abort")

            Dim splashOwnerHwnd As IntPtr = GetWordMainWindowHandle()
            If splashOwnerHwnd <> IntPtr.Zero Then
                splash.Show(New SharedLibrary.SharedLibrary.WindowWrapper(splashOwnerHwnd))
            Else
                splash.Show()
            End If
            splash.Refresh()

            wordApp.ScreenUpdating = False
            doc.TrackRevisions = False
            wordApp.Options.SmartCutPaste = False

            Debug.WriteLine("SurgicalMarkup: text1 length=" & text1.Length & " text2 length=" & text2.Length)

            Dim originalTrailingInlineWhitespace As String = GetTrailingInlineBoundaryWhitespace(text1)

            If Not trailingCR Then
                text2 = text2.TrimEnd(vbCr, vbLf).TrimEnd(vbCr, vbLf)
            End If

            If originalTrailingInlineWhitespace.Length > 0 Then
                text2 = EnsureTrailingInlineBoundaryWhitespace(text2, originalTrailingInlineWhitespace)
            End If

            ' ======================================================================
            ' STEP 1: Tokenize special placeholders with stable shared IDs
            ' ======================================================================
            Dim mergefields As List(Of String) = TokenizeSpecialPlaceholdersForDiff(text1, text2)

            ' ======================================================================
            ' STEP 2: Normalize line breaks to explicit tokens
            ' ======================================================================
            text1 = text1.Replace(manualLineBreak, " {vbVt} ").
                      Replace(vbCrLf, " {vbCrLf} ").
                      Replace(vbCr, " {vbCrLf} ").
                      Replace(vbLf, " {vbCrLf} ")
            text2 = text2.Replace(manualLineBreak, " {vbVt} ").
                      Replace(vbCrLf, " {vbCrLf} ").
                      Replace(vbCr, " {vbCrLf} ").
                      Replace(vbLf, " {vbCrLf} ")

            Debug.WriteLine("==== Surgical raw text1 ====")
            Debug.WriteLine(DebugVisualizeToken(text1))
            Debug.WriteLine("==== Surgical raw text2 ====")
            Debug.WriteLine(DebugVisualizeToken(text2))

            ' ======================================================================
            ' STEP 3+4: Build diff runs
            ' ======================================================================
            Dim runs As List(Of DiffRun) = BuildSurgicalDiffRuns(text1, text2)

            DebugDumpRuns("Runs after STEP 4 (before restore)", runs)

            ' ======================================================================
            ' STEP 5: Restore placeholders and line breaks in runs
            ' ======================================================================
            For riRestore As Integer = 0 To runs.Count - 1
                Dim r As DiffRun = runs(riRestore)

                r.Text = RestoreTokenizedSpecialPlaceholders(r.Text, mergefields)

                r.Text = r.Text.Replace("{vbVt}", manualLineBreak)
                r.Text = r.Text.Replace("{vbCr}", "{vbCrLf}")
                r.Text = r.Text.Replace("{vbLf}", "{vbCrLf}")
                r.Text = r.Text.Replace("{vbCrLf}", vbCrLf)

                runs(riRestore) = r
            Next

            DebugDumpRuns("Runs after STEP 5 (after restore)", runs)

            ' ======================================================================
            ' STEP 6: Quick exit if unchanged
            ' ======================================================================
            Dim hasChanges As Boolean = runs.Any(Function(r) r.RunType <> ChangeType.Unchanged)
            If Not hasChanges Then
                Debug.WriteLine("SurgicalMarkup: No changes detected, skipping.")
                Return
            End If

            ' ======================================================================
            ' STEP 7: Strip physical-element placeholders from changed runs.
            '
            ' These placeholders represent elements that already exist in the document.
            ' They must not be inserted or deleted as literal text when a neighbouring
            ' sentence is rewritten. Unchanged placeholders remain as alignment markers.
            ' ======================================================================
            For riStrip As Integer = 0 To runs.Count - 1
                If runs(riStrip).RunType = ChangeType.Deleted OrElse runs(riStrip).RunType = ChangeType.Inserted Then
                    Dim r As DiffRun = runs(riStrip)
                    r.Text = System.Text.RegularExpressions.Regex.Replace(
                    r.Text,
                    "\{\{(?:WFLD|WFNT|WENT|PFOR):.*?\}\}",
                    String.Empty,
                    System.Text.RegularExpressions.RegexOptions.Singleline)
                    runs(riStrip) = r
                End If
            Next

            Using BeginMarkupAuthorScope(wordApp)

                ' ======================================================================
                ' STEP 8: Resolve diff runs against a live atom map.
                '
                ' The previous implementation used Word.Find over the live document for
                ' every unchanged anchor and deletion cluster. That is fragile when the
                ' live range already contains tracked deletions, because Word.Find sees
                ' text that is not present in Final-view text1. It is also fragile around
                ' repeated text, because a successful Find may be the wrong occurrence.
                '
                ' This version builds one ordered atom map from the live target range.
                ' Existing tracked deletions are marked and omitted from the visible atom
                ' stream, so the diff walk consumes the same Final-view stream that was
                ' diffed. Every atom already carries its exact live [Start, End) range,
                ' so delete and insert positions are collected without per-cluster Find.
                ' ======================================================================
                Dim atoms As List(Of SurgicalTextAtom) = BuildSurgicalVisibleAtomMap(targetRange)
                Dim atomIndex As Integer = 0
                Dim ri As Integer = 0
                Dim uiYieldStopwatch As Stopwatch = Stopwatch.StartNew()
                Dim pendingEdits As New List(Of PendingSurgicalEdit)()

                Debug.WriteLine("SurgicalMarkup: visible atom count=" & atoms.Count)

                Do While ri < runs.Count
                    If uiYieldStopwatch.ElapsedMilliseconds >= uiYieldIntervalMs Then
                        System.Windows.Forms.Application.DoEvents()
                        uiYieldStopwatch.Restart()
                    End If

                    If (GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0 Then Exit Do
                    If (GetAsyncKeyState(VK_ESCAPE) And 1) <> 0 Then Exit Do

                    Dim run As DiffRun = runs(ri)

                    If run.Text Is Nothing OrElse run.Text.Length = 0 Then
                        ri += 1
                        Continue Do
                    End If

                    Debug.WriteLine($"RUN {ri:000}: {run.RunType} atomIndex={atomIndex} text='{DebugVisualizeToken(run.Text)}'")

                    Select Case run.RunType
                        Case ChangeType.Unchanged
                            Dim unchangedTokens As List(Of String) = SurgicalTokensForRunText(run.Text)
                            If unchangedTokens.Count = 0 Then
                                ri += 1
                                Continue Do
                            End If

                            Dim unchangedStart As Integer = atomIndex
                            Dim unchangedEnd As Integer = atomIndex

                            If SurgicalTokensMatchAt(atoms, unchangedTokens, atomIndex) Then
                                unchangedEnd = atomIndex + unchangedTokens.Count
                                atomIndex = unchangedEnd
                            ElseIf unchangedTokens.All(Function(token As String) SurgicalIsWeakAnchorToken(token)) Then
                                ' Spaces and break-padding tokens are often synthetic artefacts of the
                                ' diff-normalisation layer. Never resync globally on them; doing so can
                                ' jump to a much later space and make every later real edit disappear.
                                Debug.WriteLine($"    Weak unchanged run skipped without resync. text='{DebugVisualizeToken(run.Text)}'")
                            Else
                                Dim unchangedSearchLimit As Integer = System.Math.Min(atoms.Count, atomIndex + System.Math.Max(64, unchangedTokens.Count * 4))

                                If SurgicalTryFindTokenSequence(atoms, unchangedTokens, atomIndex, unchangedSearchLimit, unchangedStart, unchangedEnd) Then
                                    Debug.WriteLine($"    Resynced unchanged run locally: {atomIndex} -> {unchangedEnd}; skipped visible atoms={unchangedStart - atomIndex}")
                                    atomIndex = unchangedEnd
                                Else
                                    Debug.WriteLine($"    Unchanged run not found in atom map; leaving atom index unchanged. text='{DebugVisualizeToken(run.Text)}'")
                                End If
                            End If

                            ri += 1

                        Case ChangeType.Deleted, ChangeType.Inserted
                            Dim clusterStartRunIndex As Integer = ri
                            Dim cluster As New List(Of DiffRun)()

                            Do While ri < runs.Count AndAlso runs(ri).RunType <> ChangeType.Unchanged
                                If runs(ri).Text IsNot Nothing AndAlso runs(ri).Text.Length > 0 Then
                                    cluster.Add(runs(ri))
                                End If
                                ri += 1
                            Loop

                            If cluster.Count = 0 Then
                                Continue Do
                            End If

                            Dim deletedBuilder As New System.Text.StringBuilder()
                            Dim insertedBuilder As New System.Text.StringBuilder()

                            For Each clusterRun As DiffRun In cluster
                                Select Case clusterRun.RunType
                                    Case ChangeType.Deleted
                                        deletedBuilder.Append(clusterRun.Text)
                                    Case ChangeType.Inserted
                                        insertedBuilder.Append(clusterRun.Text)
                                End Select
                            Next

                            Dim deletedText As String = NormalizeForWordFind(deletedBuilder.ToString())
                            Dim insertedText As String = NormalizeForWordFind(insertedBuilder.ToString())

                            If deletedText.Length > 0 AndAlso insertedText.Length > 0 AndAlso
                               IsOnlyBreakCharacters(deletedText) AndAlso
                               IsOnlyBreakCharacters(insertedText) Then

                                Debug.WriteLine($"    Break-only replacement skipped; preserving original break(s). delete='{DebugVisualizeToken(deletedText)}' insert='{DebugVisualizeToken(insertedText)}'")
                                Continue Do
                            End If

                            Dim deletedTokens As List(Of String) = SurgicalTokenizeComparableText(deletedText)
                            Dim anchorPos As Integer = SurgicalAnchorPositionFromAtomIndex(atoms, atomIndex, targetRange.End)

                            Debug.WriteLine($"    Cluster delete='{DebugVisualizeToken(deletedText)}' insert='{DebugVisualizeToken(insertedText)}' atomIndex={atomIndex} anchor={anchorPos}")

                            If deletedTokens.Count = 0 Then
                                If Not String.IsNullOrEmpty(insertedText) Then
                                    pendingEdits.Add(New PendingSurgicalEdit With {
                                    .DeleteStart = anchorPos,
                                    .DeleteEnd = anchorPos,
                                    .InsertText = insertedText
                                })
                                    Debug.WriteLine($"    INSERT planned @{anchorPos} text='{DebugVisualizeToken(insertedText)}'")
                                End If

                                Continue Do
                            End If

                            Dim deleteStartIndex As Integer = atomIndex
                            Dim deleteEndIndex As Integer = atomIndex
                            Dim deleteMatched As Boolean = False

                            If SurgicalTokensMatchAt(atoms, deletedTokens, atomIndex) Then
                                deleteStartIndex = atomIndex
                                deleteEndIndex = atomIndex + deletedTokens.Count
                                deleteMatched = True
                            Else
                                Dim nextAnchorTokens As List(Of String) = SurgicalFindNextUnchangedAnchorTokens(runs, ri)
                                Dim nextAnchorStart As Integer = atoms.Count
                                Dim nextAnchorEnd As Integer = atoms.Count

                                If nextAnchorTokens.Count > 0 Then
                                    If Not SurgicalTryFindTokenSequence(atoms, nextAnchorTokens, atomIndex, atoms.Count, nextAnchorStart, nextAnchorEnd) Then
                                        nextAnchorStart = atoms.Count
                                        nextAnchorEnd = atoms.Count
                                    End If
                                End If

                                If SurgicalTryFindTokenSequence(atoms, deletedTokens, atomIndex, nextAnchorStart, deleteStartIndex, deleteEndIndex) Then
                                    deleteMatched = True
                                    Debug.WriteLine($"    Delete tokens found after resync: {atomIndex} -> {deleteStartIndex}/{deleteEndIndex}")
                                ElseIf nextAnchorStart >= atomIndex Then
                                    ' The diff cluster sits between the already-consumed original atom and
                                    ' the next unchanged anchor. If exact token matching fails, use that
                                    ' bounded original-side region rather than dropping the cluster.
                                    deleteStartIndex = atomIndex
                                    deleteEndIndex = nextAnchorStart
                                    deleteMatched = (deleteEndIndex > deleteStartIndex)
                                    Debug.WriteLine($"    Delete tokens not found; using bounded cluster region {deleteStartIndex}..{deleteEndIndex} before next unchanged anchor.")
                                End If
                            End If

                            If deleteMatched Then
                                Dim deleteAnchor As Integer = SurgicalAnchorPositionFromAtomIndex(atoms, deleteStartIndex, targetRange.End)
                                SurgicalAppendPendingEditForAtomRange(pendingEdits, atoms, deleteStartIndex, deleteEndIndex, insertedText, deleteAnchor)
                                atomIndex = deleteEndIndex

                                Debug.WriteLine($"    EDIT planned atoms={deleteStartIndex}..{deleteEndIndex} anchor={deleteAnchor} insert='{DebugVisualizeToken(insertedText)}'")
                            Else
                                ' Last-resort safety: a failed cluster should not corrupt the document by
                                ' deleting a guessed live range. Keep the diff walk aligned by consuming to
                                ' the next unchanged anchor if one can be found; otherwise leave the atom
                                ' pointer unchanged and log the failure.
                                Dim resyncTokens As List(Of String) = SurgicalFindNextUnchangedAnchorTokens(runs, ri)
                                Dim resyncStart As Integer = atomIndex
                                Dim resyncEnd As Integer = atomIndex
                                If resyncTokens.Count > 0 AndAlso SurgicalTryFindTokenSequence(atoms, resyncTokens, atomIndex, atoms.Count, resyncStart, resyncEnd) Then
                                    Debug.WriteLine($"    Cluster not applied; resynced to next unchanged anchor at atom {resyncStart}.")
                                    atomIndex = resyncStart
                                Else
                                    Debug.WriteLine($"    Cluster not applied and no resync anchor was found. delete='{DebugVisualizeToken(deletedText)}' insert='{DebugVisualizeToken(insertedText)}'")
                                End If
                            End If

                        Case Else
                            Debug.WriteLine($"SurgicalMarkup: Unexpected ChangeType '{run.RunType}' at run {ri}")
                            ri += 1
                    End Select
                Loop

                ' ==================================================================
                ' STEP 9: Apply the planned edits right-to-left.
                ' ==================================================================
                pendingEdits.Sort(Function(a, b)
                                      Dim startCompare As Integer = b.DeleteStart.CompareTo(a.DeleteStart)
                                      If startCompare <> 0 Then Return startCompare
                                      Return b.DeleteEnd.CompareTo(a.DeleteEnd)
                                  End Function)

                Debug.WriteLine("SurgicalMarkup: pending edit count=" & pendingEdits.Count)

                doc.TrackRevisions = True

                For Each edit As PendingSurgicalEdit In pendingEdits
                    If (GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0 Then Exit For

                    Dim contentEnd As Integer = doc.Content.End
                    Dim editStart As Integer = System.Math.Max(targetRange.Start, System.Math.Min(edit.DeleteStart, contentEnd))
                    Dim editEnd As Integer = System.Math.Max(editStart, System.Math.Min(edit.DeleteEnd, contentEnd))

                    If editEnd > editStart Then
                        doc.Range(editStart, editEnd).Delete()
                    End If

                    If Not String.IsNullOrEmpty(edit.InsertText) Then
                        Dim insertRange As Microsoft.Office.Interop.Word.Range = doc.Range(editStart, editStart)
                        insertRange.InsertAfter(edit.InsertText)
                    End If
                Next

                doc.TrackRevisions = False
                CollapseDoubleSpacesInFinalView(doc, targetRange.Start, targetRange.End)
            End Using

        Catch ex As System.Exception
            Debug.WriteLine("SurgicalMarkup error: " & ex.Message & vbCrLf & ex.StackTrace)
        Finally
            If splash IsNot Nothing Then
                Try
                    splash.Close()
                Catch
                    ' Ignore: splash may already be closed/disposed.
                End Try
                splash = Nothing
            End If

            doc.TrackRevisions = originalTrack
            wordApp.ScreenUpdating = originalUpdate
            wordApp.Options.SmartCutPaste = originalSmartCutPaste

            wordApp.Selection.SetRange(targetRange.Start, targetRange.End)
        End Try
    End Sub

    Public Sub ApplySurgicalReplacement(ByVal originalText As String, ByVal newText As String, ByVal targetRange As Microsoft.Office.Interop.Word.Range, Optional ByVal trailingCR As Boolean = False)
        CompareAndInsertSurgical(originalText, newText, targetRange, trailingCR)
    End Sub

    ''' <summary>
    ''' Returns the exact inline whitespace suffix of the original selection. This suffix is a
    ''' selection boundary, not edited content: when a user selects a sentence inside a paragraph,
    ''' Word often includes the following space. The model usually omits that trailing separator,
    ''' and a normal diff would then delete it. Paragraph marks and manual line breaks are excluded
    ''' because they are structural content handled separately by the break-token logic.
    ''' </summary>
    Private Shared Function GetTrailingInlineBoundaryWhitespace(ByVal value As String) As String
        If String.IsNullOrEmpty(value) Then
            Return String.Empty
        End If

        Dim index As Integer = value.Length - 1
        Do While index >= 0 AndAlso IsInlineBoundaryWhitespace(value(index))
            index -= 1
        Loop

        If index = value.Length - 1 Then
            Return String.Empty
        End If

        Return value.Substring(index + 1)
    End Function

    ''' <summary>
    ''' Reprojects the original selected range's trailing inline whitespace onto the revised text so
    ''' the boundary separator is preserved and not marked as a deletion.
    ''' </summary>
    Private Shared Function EnsureTrailingInlineBoundaryWhitespace(ByVal value As String, ByVal trailingWhitespace As String) As String
        If String.IsNullOrEmpty(trailingWhitespace) Then
            Return If(value, String.Empty)
        End If

        Dim text As String = If(value, String.Empty)
        Dim index As Integer = text.Length - 1
        Do While index >= 0 AndAlso IsInlineBoundaryWhitespace(text(index))
            index -= 1
        Loop

        Return text.Substring(0, index + 1) & trailingWhitespace
    End Function

    Private Shared Function IsInlineBoundaryWhitespace(ByVal ch As Char) As Boolean
        Return ch = " "c OrElse
               ch = ControlChars.Tab OrElse
               ch = ChrW(160)
    End Function

    ''' <summary>
    ''' Builds the surgical diff runs for the tokenized original/revised text. When
    ''' <see cref="SurgicalSentenceDiffEnabled"/> is False this is exactly the original word-level
    ''' behavior. When True, a sentence-first pass is performed: fully rewritten sentences (below the
    ''' retention threshold, with no placeholder or break token) are collapsed into a single tracked
    ''' delete + insert, while all other regions are delegated back to the word-level engine so their
    ''' behavior and safety are unchanged.
    ''' </summary>
    ''' <param name="text1">Tokenized original text (placeholders and breaks already tokenized).</param>
    ''' <param name="text2">Tokenized revised text (placeholders and breaks already tokenized).</param>
    ''' <returns>The ordered list of diff runs consumed by the surgical apply loop.</returns>
    Private Shared Function BuildSurgicalDiffRuns(ByVal text1 As String, ByVal text2 As String) As System.Collections.Generic.List(Of DiffRun)

        If Not SurgicalSentenceDiffEnabled Then
            ' Strict surgical mode: keep the historical word-level markup policy, but still use the
            ' newer atom-map apply phase. This is useful for proofreading/typo correction, where
            ' sentence replacement would be visually too coarse.
            Return BuildWordLevelDiffRuns(text1, text2)
        End If

        ' Sentence-aware surgical mode: the atom-map apply phase can only produce sentence-level
        ' markup when the diff stream tells it that a sentence or sentence group is a replacement.
        ' A pure word-level patience diff is correct for small edits, but it over-anchors heavily
        ' rewritten paragraphs on common words such as "the", "will", "evidence", and party names.
        ' Therefore this mode builds the diff in three layers: paragraph order, sentence alignment,
        ' then word-level diff only for sentences that are still genuinely close.
        Dim paragraphs1 As System.Collections.Generic.List(Of SurgicalParagraphChunk) = SplitTokenizedIntoParagraphChunks(text1)
        Dim paragraphs2 As System.Collections.Generic.List(Of SurgicalParagraphChunk) = SplitTokenizedIntoParagraphChunks(text2)
        RemoveSyntheticEmptyParagraphChunks(paragraphs2)

        Dim runs As New System.Collections.Generic.List(Of DiffRun)()

        If paragraphs1.Count = paragraphs2.Count Then
            For paragraphIndex As Integer = 0 To paragraphs1.Count - 1
                AppendParagraphDiffRuns(runs, paragraphs1(paragraphIndex).Text, paragraphs2(paragraphIndex).Text)
                AppendParagraphBreakRuns(runs, paragraphs1(paragraphIndex).BreakText, paragraphs2(paragraphIndex).BreakText)
            Next
            Return runs
        End If

        ' If the paragraph counts really changed, align at sentence level over the whole stream.
        ' Paragraph break tokens remain hard anchors and will not be hidden inside collapsed sentences.
        AppendParagraphDiffRuns(runs, text1, text2)
        Return runs
    End Function

    Private Structure SurgicalParagraphChunk
        Public Text As String
        Public BreakText As String
    End Structure

    Private Structure SurgicalSentenceAlignmentItem
        Public OldIndex As Integer
        Public NewIndex As Integer
    End Structure

    ''' <summary>
    ''' Splits the tokenized text into paragraph bodies plus the break token that followed each body.
    ''' The normalization step intentionally wrapped every break token in spaces. Those synthetic
    ''' padding spaces do not exist in Word's live text, so exactly one injected space is removed on
    ''' each side of a paragraph break while genuine additional spaces are preserved.
    ''' </summary>
    Private Shared Function SplitTokenizedIntoParagraphChunks(ByVal text As String) As System.Collections.Generic.List(Of SurgicalParagraphChunk)
        Dim chunks As New System.Collections.Generic.List(Of SurgicalParagraphChunk)()
        Dim current As New System.Text.StringBuilder()
        Dim skipOneInjectedSpaceAfterBreak As Boolean = False

        For Each token As String In TokenizeDiffUnits(If(text, String.Empty))
            If IsDiffLineBreakToken(token) Then
                RemoveOneTrailingInjectedSpace(current)
                chunks.Add(New SurgicalParagraphChunk With {
                    .Text = current.ToString(),
                    .BreakText = token
                })
                current.Clear()
                skipOneInjectedSpaceAfterBreak = True
            Else
                If skipOneInjectedSpaceAfterBreak AndAlso token = " " Then
                    skipOneInjectedSpaceAfterBreak = False
                Else
                    current.Append(token)
                    skipOneInjectedSpaceAfterBreak = False
                End If
            End If
        Next

        chunks.Add(New SurgicalParagraphChunk With {
            .Text = current.ToString(),
            .BreakText = String.Empty
        })

        Return chunks
    End Function

    Private Shared Sub RemoveOneTrailingInjectedSpace(ByVal builder As System.Text.StringBuilder)
        If builder Is Nothing OrElse builder.Length = 0 Then
            Return
        End If

        If builder(builder.Length - 1) = " "c Then
            builder.Length -= 1
        End If
    End Sub

    Private Shared Sub RemoveSyntheticEmptyParagraphChunks(ByVal chunks As System.Collections.Generic.List(Of SurgicalParagraphChunk))
        If chunks Is Nothing OrElse chunks.Count = 0 Then
            Return
        End If

        Dim index As Integer = 0
        Do While index < chunks.Count
            Dim chunk As SurgicalParagraphChunk = chunks(index)
            Dim isLast As Boolean = (index = chunks.Count - 1)

            If Not isLast AndAlso String.IsNullOrWhiteSpace(chunk.Text) Then
                ' Gemini and similar models often serialize paragraph separation as a blank line even
                ' when instructed not to. Treat such empty chunks as separator noise so the original
                ' Word paragraph structure remains the structural source of truth.
                chunks.RemoveAt(index)
            Else
                index += 1
            End If
        Loop
    End Sub

    Private Shared Sub AppendParagraphBreakRuns(
    ByVal runs As System.Collections.Generic.List(Of DiffRun),
    ByVal originalBreakText As String,
    ByVal revisedBreakText As String)

        If Not String.IsNullOrEmpty(originalBreakText) AndAlso Not String.IsNullOrEmpty(revisedBreakText) Then
            ' Preserve the existing Word paragraph mark. Replacing break tokens merely because the
            ' model emitted vbLf/vbCrLf differently is not a content edit and can damage lists.
            runs.Add(New DiffRun With {.RunType = ChangeType.Unchanged, .Text = originalBreakText})
        ElseIf Not String.IsNullOrEmpty(originalBreakText) Then
            runs.Add(New DiffRun With {.RunType = ChangeType.Deleted, .Text = originalBreakText})
        ElseIf Not String.IsNullOrEmpty(revisedBreakText) Then
            runs.Add(New DiffRun With {.RunType = ChangeType.Inserted, .Text = revisedBreakText})
        End If
    End Sub

    Private Shared Sub AppendParagraphDiffRuns(
    ByVal runs As System.Collections.Generic.List(Of DiffRun),
    ByVal originalParagraphText As String,
    ByVal revisedParagraphText As String)

        Dim oldText As String = If(originalParagraphText, String.Empty)
        Dim newText As String = If(revisedParagraphText, String.Empty)

        If oldText.Length = 0 AndAlso newText.Length = 0 Then
            Return
        End If

        If oldText.Length = 0 Then
            runs.Add(New DiffRun With {.RunType = ChangeType.Inserted, .Text = newText})
            Return
        End If

        If newText.Length = 0 Then
            runs.Add(New DiffRun With {.RunType = ChangeType.Deleted, .Text = oldText})
            Return
        End If

        If SentenceGroupContainsPlaceholder(oldText) OrElse SentenceGroupContainsPlaceholder(newText) Then
            ' Physical Word elements are hard anchors, but the text around them may still be a
            ' sentence-level rewrite. Split around equal placeholders, diff each plain-text segment
            ' with the normal paragraph/sentence rules, and emit the placeholder itself unchanged.
            ' This prevents field/footnote placeholders from forcing a whole translated sentence
            ' back into fragile word-by-word replacement.
            AppendAnchoredParagraphDiffRuns(runs, oldText, newText)
            Return
        End If

        Dim oldSentences As System.Collections.Generic.List(Of String) = SplitTokenizedIntoSentences(oldText)
        Dim newSentences As System.Collections.Generic.List(Of String) = SplitTokenizedIntoSentences(newText)

        If oldSentences.Count = 0 OrElse newSentences.Count = 0 Then
            AppendSentenceGroupRuns(runs, oldSentences, newSentences)
            Return
        End If

        Dim alignment As System.Collections.Generic.List(Of SurgicalSentenceAlignmentItem) = SurgicalAlignSentences(oldSentences, newSentences)
        Dim pendingDeleted As New System.Collections.Generic.List(Of String)()
        Dim pendingInserted As New System.Collections.Generic.List(Of String)()

        For Each item As SurgicalSentenceAlignmentItem In alignment
            If item.OldIndex >= 0 AndAlso item.NewIndex >= 0 Then
                If pendingDeleted.Count > 0 OrElse pendingInserted.Count > 0 Then
                    AppendSentenceGroupRuns(runs, pendingDeleted, pendingInserted)
                    pendingDeleted.Clear()
                    pendingInserted.Clear()
                End If

                AppendSentencePairRuns(runs, oldSentences(item.OldIndex), newSentences(item.NewIndex))
            ElseIf item.OldIndex >= 0 Then
                pendingDeleted.Add(oldSentences(item.OldIndex))
            ElseIf item.NewIndex >= 0 Then
                pendingInserted.Add(newSentences(item.NewIndex))
            End If
        Next

        If pendingDeleted.Count > 0 OrElse pendingInserted.Count > 0 Then
            AppendSentenceGroupRuns(runs, pendingDeleted, pendingInserted)
            pendingDeleted.Clear()
            pendingInserted.Clear()
        End If
    End Sub

    Private Shared Sub AppendAnchoredParagraphDiffRuns(
    ByVal runs As System.Collections.Generic.List(Of DiffRun),
    ByVal originalParagraphText As String,
    ByVal revisedParagraphText As String)

        Dim segments1 As New System.Collections.Generic.List(Of String)()
        Dim anchors1 As New System.Collections.Generic.List(Of String)()
        Dim segments2 As New System.Collections.Generic.List(Of String)()
        Dim anchors2 As New System.Collections.Generic.List(Of String)()

        SplitAtDiffAnchors(If(originalParagraphText, String.Empty), segments1, anchors1)
        SplitAtDiffAnchors(If(revisedParagraphText, String.Empty), segments2, anchors2)

        Dim anchorsMatch As Boolean = anchors1.Count > 0 AndAlso anchors1.Count = anchors2.Count
        If anchorsMatch Then
            For anchorIndex As Integer = 0 To anchors1.Count - 1
                If Not String.Equals(anchors1(anchorIndex), anchors2(anchorIndex), StringComparison.Ordinal) Then
                    anchorsMatch = False
                    Exit For
                End If
            Next
        End If

        If Not anchorsMatch Then
            ' Different placeholder sequences means the model has added, removed, or reordered a
            ' physical Word element. Do not collapse that case; use the conservative anchor-aware
            ' word diff so existing elements are not accidentally displaced.
            runs.AddRange(BuildWordLevelDiffRuns(originalParagraphText, revisedParagraphText))
            Return
        End If

        For segmentIndex As Integer = 0 To segments1.Count - 1
            Dim oldSegment As String = segments1(segmentIndex)
            Dim newSegment As String = segments2(segmentIndex)

            If oldSegment.Length > 0 OrElse newSegment.Length > 0 Then
                AppendParagraphDiffRuns(runs, oldSegment, newSegment)
            End If

            If segmentIndex < anchors1.Count Then
                runs.Add(New DiffRun With {.RunType = ChangeType.Unchanged, .Text = anchors1(segmentIndex)})
            End If
        Next
    End Sub

    Private Shared Function SurgicalAlignSentences(
    ByVal oldSentences As System.Collections.Generic.List(Of String),
    ByVal newSentences As System.Collections.Generic.List(Of String)) As System.Collections.Generic.List(Of SurgicalSentenceAlignmentItem)

        Dim result As New System.Collections.Generic.List(Of SurgicalSentenceAlignmentItem)()
        Dim oldCount As Integer = If(oldSentences Is Nothing, 0, oldSentences.Count)
        Dim newCount As Integer = If(newSentences Is Nothing, 0, newSentences.Count)

        If oldCount = 0 Then
            For j As Integer = 0 To newCount - 1
                result.Add(New SurgicalSentenceAlignmentItem With {.OldIndex = -1, .NewIndex = j})
            Next
            Return result
        End If

        If newCount = 0 Then
            For i As Integer = 0 To oldCount - 1
                result.Add(New SurgicalSentenceAlignmentItem With {.OldIndex = i, .NewIndex = -1})
            Next
            Return result
        End If

        Dim scores(oldCount, newCount) As Double
        Dim moves(oldCount, newCount) As Integer
        Const gapPenalty As Double = 0.42
        Const minimumUsefulSentenceSimilarity As Double = 0.22

        For i As Integer = 1 To oldCount
            scores(i, 0) = scores(i - 1, 0) - gapPenalty
            moves(i, 0) = 2
        Next

        For j As Integer = 1 To newCount
            scores(0, j) = scores(0, j - 1) - gapPenalty
            moves(0, j) = 3
        Next

        For i As Integer = 1 To oldCount
            For j As Integer = 1 To newCount
                Dim similarity As Double = SentenceSimilarityRatio(oldSentences(i - 1), newSentences(j - 1))
                Dim matchScore As Double = scores(i - 1, j - 1) + If(similarity >= minimumUsefulSentenceSimilarity, similarity, -gapPenalty)
                Dim deleteScore As Double = scores(i - 1, j) - gapPenalty
                Dim insertScore As Double = scores(i, j - 1) - gapPenalty

                If matchScore >= deleteScore AndAlso matchScore >= insertScore Then
                    scores(i, j) = matchScore
                    moves(i, j) = 1
                ElseIf deleteScore >= insertScore Then
                    scores(i, j) = deleteScore
                    moves(i, j) = 2
                Else
                    scores(i, j) = insertScore
                    moves(i, j) = 3
                End If
            Next
        Next

        Dim ai As Integer = oldCount
        Dim aj As Integer = newCount
        Do While ai > 0 OrElse aj > 0
            Dim moveCode As Integer = moves(ai, aj)

            If ai > 0 AndAlso aj > 0 AndAlso moveCode = 1 Then
                result.Add(New SurgicalSentenceAlignmentItem With {.OldIndex = ai - 1, .NewIndex = aj - 1})
                ai -= 1
                aj -= 1
            ElseIf ai > 0 AndAlso (aj = 0 OrElse moveCode = 2) Then
                result.Add(New SurgicalSentenceAlignmentItem With {.OldIndex = ai - 1, .NewIndex = -1})
                ai -= 1
            ElseIf aj > 0 Then
                result.Add(New SurgicalSentenceAlignmentItem With {.OldIndex = -1, .NewIndex = aj - 1})
                aj -= 1
            Else
                Exit Do
            End If
        Loop

        result.Reverse()
        Return result
    End Function

    ''' <summary>
    ''' Emits diff runs for one aligned sentence change group. Unlike the earlier implementation,
    ''' differing sentence counts are not delegated blindly to word-level diff. Heavy summarisation
    ''' commonly turns two or three old sentences into one revised sentence; delegating that case to
    ''' word-level diff is exactly what leaves rewritten paragraphs under-marked.
    ''' </summary>
    Private Shared Sub AppendSentenceGroupRuns(
    ByVal runs As System.Collections.Generic.List(Of DiffRun),
    ByVal deletedSentences As System.Collections.Generic.List(Of String),
    ByVal insertedSentences As System.Collections.Generic.List(Of String))

        Dim deletedText As String = String.Concat(If(deletedSentences, New System.Collections.Generic.List(Of String)()))
        Dim insertedText As String = String.Concat(If(insertedSentences, New System.Collections.Generic.List(Of String)()))

        If deletedText.Length = 0 AndAlso insertedText.Length = 0 Then
            Return
        End If

        If deletedText.Length = 0 Then
            runs.Add(New DiffRun With {.RunType = ChangeType.Inserted, .Text = insertedText})
            Return
        End If

        If insertedText.Length = 0 Then
            runs.Add(New DiffRun With {.RunType = ChangeType.Deleted, .Text = deletedText})
            Return
        End If

        If ShouldCollapseSentenceReplacement(deletedText, insertedText) Then
            runs.Add(New DiffRun With {.RunType = ChangeType.Deleted, .Text = deletedText})
            runs.Add(New DiffRun With {.RunType = ChangeType.Inserted, .Text = insertedText})
        Else
            runs.AddRange(BuildWordLevelDiffRuns(deletedText, insertedText))
        End If
    End Sub

    ''' <summary>
    ''' Emits diff runs for one positionally aligned sentence pair. Small edits are still diffed at
    ''' word level. Once the sentence-level LCS says the wording has materially changed, the whole
    ''' sentence is replaced so the markup matches the user's mental model of a rewrite.
    ''' </summary>
    Private Shared Sub AppendSentencePairRuns(
    ByVal runs As System.Collections.Generic.List(Of DiffRun),
    ByVal deletedSentence As String,
    ByVal insertedSentence As String)

        If ShouldCollapseSentenceReplacement(deletedSentence, insertedSentence) Then
            If Not String.IsNullOrEmpty(deletedSentence) Then
                runs.Add(New DiffRun With {.RunType = ChangeType.Deleted, .Text = deletedSentence})
            End If
            If Not String.IsNullOrEmpty(insertedSentence) Then
                runs.Add(New DiffRun With {.RunType = ChangeType.Inserted, .Text = insertedSentence})
            End If
        Else
            runs.AddRange(BuildWordLevelDiffRuns(deletedSentence, insertedSentence))
        End If
    End Sub

    Private Shared Function ShouldCollapseSentenceReplacement(ByVal deletedSentence As String, ByVal insertedSentence As String) As Boolean
        If String.IsNullOrWhiteSpace(deletedSentence) OrElse String.IsNullOrWhiteSpace(insertedSentence) Then
            Return False
        End If

        If SentenceGroupContainsPlaceholder(deletedSentence) OrElse SentenceGroupContainsPlaceholder(insertedSentence) Then
            Return False
        End If

        If SentenceGroupContainsBreakToken(deletedSentence) OrElse SentenceGroupContainsBreakToken(insertedSentence) Then
            Return False
        End If

        Dim deletedCore As String = deletedSentence.Trim()
        Dim insertedCore As String = insertedSentence.Trim()

        If deletedCore.Length < SurgicalSentenceMinLength OrElse insertedCore.Length < SurgicalSentenceMinLength Then
            Return False
        End If

        Dim deletedTokens As System.Collections.Generic.List(Of String) = GetComparableSentenceTokens(deletedCore)
        Dim insertedTokens As System.Collections.Generic.List(Of String) = GetComparableSentenceTokens(insertedCore)

        If deletedTokens.Count < 7 OrElse insertedTokens.Count < 7 Then
            Return False
        End If

        Dim lcsLength As Integer = ComparableTokenLcsLength(deletedTokens, insertedTokens)
        Dim similarity As Double = (2.0 * lcsLength) / (deletedTokens.Count + insertedTokens.Count)
        Dim changedTokenCount As Integer = (deletedTokens.Count - lcsLength) + (insertedTokens.Count - lcsLength)
        Dim changedTokenRatio As Double = changedTokenCount / (deletedTokens.Count + insertedTokens.Count)



        Return changedTokenCount >= materialRewriteMinimumChangedTokens AndAlso
               similarity < materialRewriteSimilarityThreshold AndAlso
               changedTokenRatio >= materialRewriteChangedTokenRatioThreshold
    End Function


    Private Shared Sub SplitAtDiffAnchors(
    ByVal text As String,
    ByVal segments As List(Of String),
    ByVal anchors As List(Of String))

        segments.Clear()
        anchors.Clear()

        Dim sb As New System.Text.StringBuilder()
        For Each token As String In TokenizeDiffUnits(text)
            If IsDiffPlaceholderToken(token) OrElse IsDiffLineBreakToken(token) Then
                segments.Add(sb.ToString())
                sb.Clear()
                anchors.Add(token)
            Else
                sb.Append(token)
            End If
        Next
        segments.Add(sb.ToString())
    End Sub

    ''' <summary>
    ''' Anchor-aware word-level diff. Placeholders ([[MF#]]) and break tokens map to physical,
    ''' already-existing document elements (footnotes, fields, paragraph marks). If the diff engine is
    ''' allowed to treat them as ordinary movable tokens, a heavily rewritten neighbourhood lets
    ''' DiffPlex report them as a "move" (delete at the old spot + insert at the new one), which
    ''' physically displaces the element - e.g. a footnote jumping across a paragraph boundary. When
    ''' both sides expose the SAME ordered anchor sequence, only the text BETWEEN anchors is diffed and
    ''' each anchor is emitted as an Unchanged run, pinning it to its original position. Otherwise the
    ''' original whole-stream diff is used so behavior is unchanged.
    ''' </summary>
    ''' <param name="text1">Tokenized original text.</param>
    ''' <param name="text2">Tokenized revised text.</param>
    ''' <returns>The list of merged diff runs.</returns>
    Private Shared Function BuildWordLevelDiffRuns(ByVal text1 As String, ByVal text2 As String) As List(Of DiffRun)

        Dim segments1 As New List(Of String)()
        Dim anchors1 As New List(Of String)()
        Dim segments2 As New List(Of String)()
        Dim anchors2 As New List(Of String)()

        SplitAtDiffAnchors(text1, segments1, anchors1)
        SplitAtDiffAnchors(text2, segments2, anchors2)

        Dim anchorsMatch As Boolean = anchors1.Count > 0 AndAlso anchors1.Count = anchors2.Count
        If anchorsMatch Then
            For ai As Integer = 0 To anchors1.Count - 1
                If Not String.Equals(anchors1(ai), anchors2(ai), StringComparison.Ordinal) Then
                    anchorsMatch = False
                    Exit For
                End If
            Next
        End If

        If Not anchorsMatch Then
            Return BuildWordLevelDiffRunsCore(text1, text2)
        End If

        Dim runs As New List(Of DiffRun)()
        For si As Integer = 0 To segments1.Count - 1
            If segments1(si).Length > 0 OrElse segments2(si).Length > 0 Then
                runs.AddRange(BuildWordLevelDiffRunsCore(segments1(si), segments2(si)))
            End If

            If si < anchors1.Count Then
                runs.Add(New DiffRun With {.RunType = ChangeType.Unchanged, .Text = anchors1(si)})
            End If
        Next

        Return runs
    End Function

    ''' <summary>
    ''' Original word-level diff-run builder extracted from the surgical engine. Tokenizes both
    ''' sides, runs the DiffPlex inline diff, and merges consecutive same-type tokens into runs while
    ''' keeping line-break, placeholder, and whitespace tokens isolated.
    ''' </summary>
    ''' <param name="text1">Tokenized original text.</param>
    ''' <param name="text2">Tokenized revised text.</param>
    ''' <returns>The list of merged diff runs.</returns>
    Private Shared Function BuildWordLevelDiffRunsCore(ByVal text1 As String, ByVal text2 As String) As List(Of DiffRun)

        Dim tokenList1 As List(Of String) = TokenizeDiffUnits(text1)
        Dim tokenList2 As List(Of String) = TokenizeDiffUnits(text2)

        ' Language-independent patience diff: unique-common tokens anchor the alignment,
        ' so repeated filler tokens (spaces, "the", "of", "Mr.") can never scatter a
        ' reordered span into tiny interleaved matches. One typed entry per token is
        ' produced; the merge loop below applies the same isolation rules as before.
        Dim typedTokens As List(Of DiffRun) = BuildPatienceTokenDiff(tokenList1, tokenList2)

        Dim runs As New List(Of DiffRun)(System.Math.Max(4, typedTokens.Count))
        Dim currentRunType As ChangeType = ChangeType.Unchanged
        Dim currentRunWords As New List(Of String)
        Dim lastTokenText As String = Nothing

        For i As Integer = 0 To typedTokens.Count - 1
            Dim tokenType As ChangeType = typedTokens(i).RunType
            Dim wordText As String = If(typedTokens(i).Text, String.Empty)
            If wordText.Length = 0 Then Continue For

            Dim isLB As Boolean = IsDiffLineBreakToken(wordText)
            Dim isPlaceholderToken As Boolean = IsDiffPlaceholderToken(wordText)
            Dim isWhitespaceToken As Boolean = IsDiffWhitespaceToken(wordText)

            Dim prevWasLB As Boolean = (currentRunWords.Count > 0 AndAlso IsDiffLineBreakToken(lastTokenText))
            Dim prevWasPlaceholder As Boolean = (currentRunWords.Count > 0 AndAlso IsDiffPlaceholderToken(lastTokenText))
            Dim prevWasWhitespace As Boolean = (currentRunWords.Count > 0 AndAlso IsDiffWhitespaceToken(lastTokenText))

            If (tokenType <> currentRunType OrElse
            isLB OrElse
            isPlaceholderToken OrElse
            isWhitespaceToken OrElse
            prevWasLB OrElse
            prevWasPlaceholder OrElse
            prevWasWhitespace) AndAlso currentRunWords.Count > 0 Then

                runs.Add(New DiffRun With {
                .RunType = currentRunType,
                .Text = String.Concat(currentRunWords)
            })
                currentRunWords.Clear()
            End If

            currentRunType = tokenType
            currentRunWords.Add(wordText)
            lastTokenText = wordText
        Next

        If currentRunWords.Count > 0 Then
            runs.Add(New DiffRun With {
            .RunType = currentRunType,
            .Text = String.Concat(currentRunWords)
        })
        End If

        Return runs
    End Function

    ''' <summary>
    ''' Computes a patience diff over the token streams and returns one typed <see cref="DiffRun"/>
    ''' per token (Unchanged/Deleted/Inserted). Anchors are tokens that occur exactly once on each
    ''' side, so alignment never depends on language-specific rules or tunable thresholds.
    ''' </summary>
    Private Shared Function BuildPatienceTokenDiff(
    ByVal tokens1 As List(Of String),
    ByVal tokens2 As List(Of String)) As List(Of DiffRun)

        Dim output As New List(Of DiffRun)()
        PatienceDiffRange(tokens1, tokens2, 0, tokens1.Count, 0, tokens2.Count, output)
        Return output
    End Function

    ''' <summary>
    ''' Recursive patience diff on the index ranges [lo1,hi1) x [lo2,hi2). Trims the common
    ''' prefix/suffix, aligns on unique-common anchors (via the longest increasing subsequence of
    ''' their positions), and recurses into the gaps. Ranges with no unique anchors fall back to the
    ''' standard DiffPlex LCS, which is always a small, bounded slice here.
    ''' </summary>
    Private Shared Sub PatienceDiffRange(
    ByVal a As List(Of String),
    ByVal b As List(Of String),
    ByVal lo1 As Integer, ByVal hi1 As Integer,
    ByVal lo2 As Integer, ByVal hi2 As Integer,
    ByVal output As List(Of DiffRun))

        ' Common prefix.
        Do While lo1 < hi1 AndAlso lo2 < hi2 AndAlso String.Equals(a(lo1), b(lo2), StringComparison.Ordinal)
            output.Add(New DiffRun With {.RunType = ChangeType.Unchanged, .Text = a(lo1)})
            lo1 += 1
            lo2 += 1
        Loop

        ' Common suffix (collected reversed; appended once the middle is emitted).
        Dim suffix As New List(Of DiffRun)()
        Do While lo1 < hi1 AndAlso lo2 < hi2 AndAlso String.Equals(a(hi1 - 1), b(hi2 - 1), StringComparison.Ordinal)
            suffix.Add(New DiffRun With {.RunType = ChangeType.Unchanged, .Text = a(hi1 - 1)})
            hi1 -= 1
            hi2 -= 1
        Loop
        suffix.Reverse()

        If lo1 >= hi1 AndAlso lo2 >= hi2 Then
            output.AddRange(suffix)
            Return
        End If

        If lo1 >= hi1 Then
            For j As Integer = lo2 To hi2 - 1
                output.Add(New DiffRun With {.RunType = ChangeType.Inserted, .Text = b(j)})
            Next
            output.AddRange(suffix)
            Return
        End If

        If lo2 >= hi2 Then
            For i As Integer = lo1 To hi1 - 1
                output.Add(New DiffRun With {.RunType = ChangeType.Deleted, .Text = a(i)})
            Next
            output.AddRange(suffix)
            Return
        End If

        Dim anchors As List(Of KeyValuePair(Of Integer, Integer)) =
        FindUniqueCommonAnchors(a, b, lo1, hi1, lo2, hi2)

        If anchors.Count = 0 Then
            FallbackTokenDiff(a, b, lo1, hi1, lo2, hi2, output)
            output.AddRange(suffix)
            Return
        End If

        Dim lis As List(Of KeyValuePair(Of Integer, Integer)) = LongestIncreasingSubsequenceByValue(anchors)

        Dim prev1 As Integer = lo1
        Dim prev2 As Integer = lo2
        For Each anchor As KeyValuePair(Of Integer, Integer) In lis
            PatienceDiffRange(a, b, prev1, anchor.Key, prev2, anchor.Value, output)
            output.Add(New DiffRun With {.RunType = ChangeType.Unchanged, .Text = a(anchor.Key)})
            prev1 = anchor.Key + 1
            prev2 = anchor.Value + 1
        Next
        PatienceDiffRange(a, b, prev1, hi1, prev2, hi2, output)

        output.AddRange(suffix)
    End Sub

    ''' <summary>
    ''' Returns the position pairs of tokens that occur exactly once in a(lo1..hi1) and exactly once
    ''' in b(lo2..hi2), sorted by their position in a. These are the patience anchors.
    ''' </summary>
    Private Shared Function FindUniqueCommonAnchors(
    ByVal a As List(Of String),
    ByVal b As List(Of String),
    ByVal lo1 As Integer, ByVal hi1 As Integer,
    ByVal lo2 As Integer, ByVal hi2 As Integer) As List(Of KeyValuePair(Of Integer, Integer))

        Dim aCount As New Dictionary(Of String, Integer)(StringComparer.Ordinal)
        Dim aPos As New Dictionary(Of String, Integer)(StringComparer.Ordinal)
        For i As Integer = lo1 To hi1 - 1
            Dim t As String = a(i)
            Dim c As Integer
            If aCount.TryGetValue(t, c) Then
                aCount(t) = c + 1
            Else
                aCount(t) = 1
                aPos(t) = i
            End If
        Next

        Dim bCount As New Dictionary(Of String, Integer)(StringComparer.Ordinal)
        Dim bPos As New Dictionary(Of String, Integer)(StringComparer.Ordinal)
        For j As Integer = lo2 To hi2 - 1
            Dim t As String = b(j)
            Dim c As Integer
            If bCount.TryGetValue(t, c) Then
                bCount(t) = c + 1
            Else
                bCount(t) = 1
                bPos(t) = j
            End If
        Next

        Dim result As New List(Of KeyValuePair(Of Integer, Integer))()
        For Each kvp As KeyValuePair(Of String, Integer) In aCount
            If kvp.Value = 1 Then
                Dim bc As Integer
                If bCount.TryGetValue(kvp.Key, bc) AndAlso bc = 1 Then
                    result.Add(New KeyValuePair(Of Integer, Integer)(aPos(kvp.Key), bPos(kvp.Key)))
                End If
            End If
        Next

        result.Sort(Function(x, y) x.Key.CompareTo(y.Key))
        Return result
    End Function

    ''' <summary>
    ''' Given anchors already sorted by their a-position, returns the longest strictly increasing
    ''' subsequence by b-position (patience sort with predecessor links).
    ''' </summary>
    Private Shared Function LongestIncreasingSubsequenceByValue(
    ByVal anchors As List(Of KeyValuePair(Of Integer, Integer))) As List(Of KeyValuePair(Of Integer, Integer))

        Dim result As New List(Of KeyValuePair(Of Integer, Integer))()
        Dim n As Integer = anchors.Count
        If n = 0 Then Return result

        Dim tails As New List(Of Integer)()
        Dim tailIdx As New List(Of Integer)()
        Dim prev(n - 1) As Integer
        For i As Integer = 0 To n - 1
            prev(i) = -1
        Next

        For i As Integer = 0 To n - 1
            Dim v As Integer = anchors(i).Value

            Dim lo As Integer = 0
            Dim hi As Integer = tails.Count
            Do While lo < hi
                Dim mid As Integer = (lo + hi) \ 2
                If tails(mid) < v Then
                    lo = mid + 1
                Else
                    hi = mid
                End If
            Loop

            If lo > 0 Then prev(i) = tailIdx(lo - 1)

            If lo = tails.Count Then
                tails.Add(v)
                tailIdx.Add(i)
            Else
                tails(lo) = v
                tailIdx(lo) = i
            End If
        Next

        Dim k As Integer = tailIdx(tailIdx.Count - 1)
        Do While k <> -1
            result.Add(anchors(k))
            k = prev(k)
        Loop
        result.Reverse()
        Return result
    End Function

    ''' <summary>
    ''' Standard DiffPlex LCS diff over a small token slice, used only when a patience range has no
    ''' unique-common anchors. Emits one typed <see cref="DiffRun"/> per token.
    ''' </summary>
    Private Shared Sub FallbackTokenDiff(
    ByVal a As List(Of String),
    ByVal b As List(Of String),
    ByVal lo1 As Integer, ByVal hi1 As Integer,
    ByVal lo2 As Integer, ByVal hi2 As Integer,
    ByVal output As List(Of DiffRun))

        Dim slice1 As New List(Of String)(System.Math.Max(0, hi1 - lo1))
        For i As Integer = lo1 To hi1 - 1
            slice1.Add(a(i))
        Next
        Dim slice2 As New List(Of String)(System.Math.Max(0, hi2 - lo2))
        For j As Integer = lo2 To hi2 - 1
            slice2.Add(b(j))
        Next

        Dim words1 As String = String.Join(Environment.NewLine, slice1)
        Dim words2 As String = String.Join(Environment.NewLine, slice2)

        Dim diffBuilder As New InlineDiffBuilder(New Differ())
        Dim diffResult As DiffPaneModel = diffBuilder.BuildDiffModel(words1, words2)

        For Each line As DiffPiece In diffResult.Lines
            Dim wordText As String = If(line.Text, String.Empty)
            If wordText.Length = 0 Then Continue For

            Select Case line.Type
                Case ChangeType.Inserted
                    output.Add(New DiffRun With {.RunType = ChangeType.Inserted, .Text = wordText})
                Case ChangeType.Deleted
                    output.Add(New DiffRun With {.RunType = ChangeType.Deleted, .Text = wordText})
                Case Else
                    output.Add(New DiffRun With {.RunType = ChangeType.Unchanged, .Text = wordText})
            End Select
        Next
    End Sub

    ''' <summary>
    ''' Splits tokenized text into lossless sentence segments. Concatenating the returned segments
    ''' reproduces the input exactly. Boundaries are placed after sentence terminators ('.', '!',
    ''' '?') that are followed by whitespace, and immediately after any paragraph/line-break token so
    ''' a segment never spans a break. Trailing whitespace stays on the segment that owns it.
    ''' </summary>
    ''' <param name="tokenizedText">Tokenized text (placeholders/breaks already tokenized).</param>
    ''' <returns>Ordered, lossless list of sentence segments.</returns>
    Private Shared Function SplitTokenizedIntoSentences(ByVal tokenizedText As String) As System.Collections.Generic.List(Of String)

        Dim result As New System.Collections.Generic.List(Of String)()
        If String.IsNullOrEmpty(tokenizedText) Then
            Return result
        End If

        Dim tokens As System.Collections.Generic.List(Of String) = TokenizeDiffUnits(tokenizedText)
        Dim current As New System.Text.StringBuilder()
        Dim i As Integer = 0

        Do While i < tokens.Count
            current.Append(tokens(i))

            If IsSentenceEndingPunctuationToken(tokens(i)) AndAlso IsSafeSentenceBoundary(tokens, i) Then
                Dim consumeTo As Integer = i + 1

                Do While consumeTo < tokens.Count AndAlso IsSentenceClosingPunctuationToken(tokens(consumeTo))
                    current.Append(tokens(consumeTo))
                    consumeTo += 1
                Loop

                If consumeTo < tokens.Count AndAlso IsDiffWhitespaceToken(tokens(consumeTo)) Then
                    current.Append(tokens(consumeTo))
                    consumeTo += 1
                End If

                If current.Length > 0 Then
                    result.Add(current.ToString())
                    current.Clear()
                End If

                i = consumeTo
            Else
                i += 1
            End If
        Loop

        If current.Length > 0 Then
            result.Add(current.ToString())
        End If

        Return result
    End Function

    Private Shared Function IsSentenceEndingPunctuationToken(ByVal token As String) As Boolean
        Return String.Equals(token, ".", StringComparison.Ordinal) OrElse
               String.Equals(token, "!", StringComparison.Ordinal) OrElse
               String.Equals(token, "?", StringComparison.Ordinal)
    End Function

    Private Shared Function IsSentenceClosingPunctuationToken(ByVal token As String) As Boolean
        Return String.Equals(token, ")", StringComparison.Ordinal) OrElse
               String.Equals(token, "]", StringComparison.Ordinal) OrElse
               String.Equals(token, "}", StringComparison.Ordinal) OrElse
               String.Equals(token, """", StringComparison.Ordinal) OrElse
               String.Equals(token, "'", StringComparison.Ordinal)
    End Function

    Private Shared Function IsSafeSentenceBoundary(ByVal tokens As System.Collections.Generic.List(Of String), ByVal punctuationIndex As Integer) As Boolean
        Dim previousWord As String = FindPreviousComparableToken(tokens, punctuationIndex - 1)
        If IsLikelyAbbreviationBeforePeriod(tokens(punctuationIndex), previousWord, tokens, punctuationIndex) Then
            Return False
        End If

        Dim nextIndex As Integer = punctuationIndex + 1
        Do While nextIndex < tokens.Count AndAlso IsSentenceClosingPunctuationToken(tokens(nextIndex))
            nextIndex += 1
        Loop

        Do While nextIndex < tokens.Count AndAlso IsDiffWhitespaceToken(tokens(nextIndex))
            nextIndex += 1
        Loop

        If nextIndex >= tokens.Count Then
            Return True
        End If

        If IsDiffLineBreakToken(tokens(nextIndex)) Then
            Return True
        End If

        Dim nextWord As String = FindNextComparableToken(tokens, nextIndex)
        If String.IsNullOrEmpty(nextWord) Then
            Return True
        End If

        Dim firstChar As Char = nextWord(0)
        Return Char.IsUpper(firstChar) OrElse Char.IsDigit(firstChar)
    End Function

    Private Shared Function IsLikelyAbbreviationBeforePeriod(
    ByVal punctuationToken As String,
    ByVal previousWord As String,
    ByVal tokens As System.Collections.Generic.List(Of String),
    ByVal punctuationIndex As Integer) As Boolean

        If Not String.Equals(punctuationToken, ".", StringComparison.Ordinal) Then
            Return False
        End If

        If String.IsNullOrEmpty(previousWord) Then
            Return False
        End If

        Dim abbreviation As String = previousWord.Trim("."c)
        If abbreviation.Length = 0 Then
            Return False
        End If

        If abbreviation.Length = 1 Then
            Return True
        End If

        Dim knownAbbreviations As New System.Collections.Generic.HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
            "Mr", "Ms", "Mrs", "Dr", "Prof", "Hon", "Jr", "Sr",
            "No", "Nr", "Art", "Abs", "Sec", "Fig", "Ref",
            "e.g", "i.e", "etc", "vs", "z.B", "d.h", "u.a"
        }

        If knownAbbreviations.Contains(abbreviation) Then
            Return True
        End If

        If abbreviation.Length <= 2 AndAlso ContainsLetter(abbreviation) Then
            Return True
        End If

        ' Initial chains such as "J. K. Rowling" or abbreviations such as "e.g." should not split.
        If punctuationIndex >= 2 AndAlso String.Equals(tokens(punctuationIndex - 2), ".", StringComparison.Ordinal) Then
            Return True
        End If

        Return False
    End Function

    Private Shared Function FindPreviousComparableToken(ByVal tokens As System.Collections.Generic.List(Of String), ByVal startIndex As Integer) As String
        If tokens Is Nothing Then
            Return String.Empty
        End If

        For i As Integer = System.Math.Min(startIndex, tokens.Count - 1) To 0 Step -1
            If IsComparableSentenceToken(tokens(i)) Then
                Return tokens(i)
            End If
        Next

        Return String.Empty
    End Function

    Private Shared Function FindNextComparableToken(ByVal tokens As System.Collections.Generic.List(Of String), ByVal startIndex As Integer) As String
        If tokens Is Nothing Then
            Return String.Empty
        End If

        For i As Integer = System.Math.Max(0, startIndex) To tokens.Count - 1
            If IsComparableSentenceToken(tokens(i)) Then
                Return tokens(i)
            End If
        Next

        Return String.Empty
    End Function

    Private Shared Function IsComparableSentenceToken(ByVal token As String) As Boolean
        If String.IsNullOrEmpty(token) Then
            Return False
        End If

        If IsDiffWhitespaceToken(token) OrElse IsDiffLineBreakToken(token) OrElse IsDiffPlaceholderToken(token) Then
            Return False
        End If

        For Each ch As Char In token
            If Char.IsLetterOrDigit(ch) Then
                Return True
            End If
        Next

        Return False
    End Function

    Private Shared Function ContainsLetter(ByVal value As String) As Boolean
        If String.IsNullOrEmpty(value) Then
            Return False
        End If

        For Each ch As Char In value
            If Char.IsLetter(ch) Then
                Return True
            End If
        Next

        Return False
    End Function

    Private Shared Function SentenceRetentionRatio(ByVal oldSentence As String, ByVal newSentence As String) As Double

        Dim oldTokens As System.Collections.Generic.List(Of String) = GetComparableSentenceTokens(oldSentence)
        Dim newTokens As System.Collections.Generic.List(Of String) = GetComparableSentenceTokens(newSentence)

        If oldTokens.Count = 0 Then
            Return 0.0
        End If

        Dim lcsLength As Integer = ComparableTokenLcsLength(oldTokens, newTokens)
        Return lcsLength / oldTokens.Count
    End Function

    Private Shared Function SentenceSimilarityRatio(ByVal oldSentence As String, ByVal newSentence As String) As Double
        Dim oldTokens As System.Collections.Generic.List(Of String) = GetComparableSentenceTokens(oldSentence)
        Dim newTokens As System.Collections.Generic.List(Of String) = GetComparableSentenceTokens(newSentence)

        If oldTokens.Count = 0 AndAlso newTokens.Count = 0 Then
            Return 1.0
        End If

        If oldTokens.Count = 0 OrElse newTokens.Count = 0 Then
            Return 0.0
        End If

        Dim lcsLength As Integer = ComparableTokenLcsLength(oldTokens, newTokens)
        Return (2.0 * lcsLength) / (oldTokens.Count + newTokens.Count)
    End Function

    Private Shared Function CountComparableSentenceTokens(ByVal value As String) As Integer
        Return GetComparableSentenceTokens(value).Count
    End Function

    Private Shared Function GetComparableSentenceTokens(ByVal value As String) As System.Collections.Generic.List(Of String)
        Dim result As New System.Collections.Generic.List(Of String)()

        If String.IsNullOrEmpty(value) Then
            Return result
        End If

        For Each token As String In TokenizeDiffUnits(value)
            If IsComparableSentenceToken(token) Then
                result.Add(token.ToLowerInvariant())
            End If
        Next

        Return result
    End Function

    Private Shared Function ComparableTokenLcsLength(
    ByVal oldTokens As System.Collections.Generic.List(Of String),
    ByVal newTokens As System.Collections.Generic.List(Of String)) As Integer

        If oldTokens Is Nothing OrElse newTokens Is Nothing OrElse oldTokens.Count = 0 OrElse newTokens.Count = 0 Then
            Return 0
        End If

        Dim previous(newTokens.Count) As Integer
        Dim current(newTokens.Count) As Integer

        For i As Integer = 1 To oldTokens.Count
            For j As Integer = 1 To newTokens.Count
                If String.Equals(oldTokens(i - 1), newTokens(j - 1), StringComparison.Ordinal) Then
                    current(j) = previous(j - 1) + 1
                Else
                    current(j) = System.Math.Max(previous(j), current(j - 1))
                End If
            Next

            Dim swap As Integer() = previous
            previous = current
            current = swap
            System.Array.Clear(current, 0, current.Length)
        Next

        Return previous(newTokens.Count)
    End Function

    Private Shared Function SentenceGroupContainsPlaceholder(ByVal value As String) As Boolean
        If String.IsNullOrEmpty(value) Then Return False
        Return value.IndexOf("[[MF", StringComparison.Ordinal) >= 0
    End Function

    ''' <summary>
    ''' Returns True when the text contains a tokenized paragraph/line-break marker. Such sentences
    ''' are never collapsed so paragraph and line structure is preserved by the word-level path.
    ''' </summary>
    ''' <param name="value">Tokenized text to inspect.</param>
    ''' <returns>True if a break token is present; otherwise False.</returns>
    Private Shared Function SentenceGroupContainsBreakToken(ByVal value As String) As Boolean
        If String.IsNullOrEmpty(value) Then Return False
        Return value.IndexOf("{vbCrLf}", StringComparison.Ordinal) >= 0 OrElse
           value.IndexOf("{vbCr}", StringComparison.Ordinal) >= 0 OrElse
           value.IndexOf("{vbLf}", StringComparison.Ordinal) >= 0 OrElse
           value.IndexOf("{vbVt}", StringComparison.Ordinal) >= 0
    End Function

    ''' <summary>
    ''' Represents one candidate surgical edit operation for a diff cluster.
    ''' Stores the text to delete, the text to insert, the search start position,
    ''' and whether a trailing whitespace token should be consumed after success.
    ''' </summary>
    Private Class SurgicalOperationCandidate
        Public Property DeleteText As String
        Public Property InsertText As String
        Public Property SearchStart As Integer
        Public Property ConsumeTrailingWhitespaceRun As Boolean
    End Class

    ''' <summary>
    ''' One planned surgical edit resolved against the pristine original text: the document
    ''' range to delete (empty when only inserting) and the text to insert. All edits are
    ''' collected during the diff walk and then applied right-to-left so an applied edit never
    ''' shifts the positions of edits that have not been applied yet.
    ''' </summary>
    Private Structure PendingSurgicalEdit
        Public DeleteStart As Integer
        Public DeleteEnd As Integer
        Public InsertText As String
    End Structure

    Private Structure SurgicalTextAtom
        Public LiveStart As Integer
        Public LiveEnd As Integer
        Public KeyText As String
        Public Spans As List(Of SurgicalPositionSpan)
    End Structure

    Private Structure SurgicalLiveCharAtom
        Public LiveStart As Integer
        Public LiveEnd As Integer
        Public Text As String
        Public IsTrackedDeletion As Boolean
    End Structure

    Private Structure SurgicalPositionSpan
        Public StartPos As Integer
        Public EndPos As Integer
    End Structure

    Private Structure SurgicalFieldAtomSpan
        Public StartPos As Integer
        Public EndPos As Integer
        Public KeyText As String
    End Structure

    ''' <summary>
    ''' Builds the visible, original-side atom stream used by the surgical apply phase.
    ''' Existing tracked deletions are read once from <paramref name="targetRange"/> and omitted
    ''' from the returned atom list, so this stream corresponds to the Final-view text that was
    ''' diffed while each atom still carries its live Word range. If a tracked deletion splits a
    ''' visible word or whitespace run, the atom keeps multiple live spans but a single KeyText.
    ''' </summary>
    Private Shared Function BuildSurgicalVisibleAtomMap(ByVal targetRange As Microsoft.Office.Interop.Word.Range) As List(Of SurgicalTextAtom)
        Dim result As New List(Of SurgicalTextAtom)()

        If targetRange Is Nothing Then
            Return result
        End If

        Dim liveChars As List(Of SurgicalLiveCharAtom) = BuildSurgicalLiveCharAtoms(targetRange)
        Dim fieldSpans As List(Of SurgicalFieldAtomSpan) = BuildSurgicalFieldAtomSpans(targetRange)
        Dim fieldIndex As Integer = 0
        Dim index As Integer = 0

        Do While index < liveChars.Count
            Dim livePos As Integer = liveChars(index).LiveStart

            Do While fieldIndex < fieldSpans.Count AndAlso fieldSpans(fieldIndex).EndPos <= livePos
                fieldIndex += 1
            Loop

            If fieldIndex < fieldSpans.Count AndAlso
               fieldSpans(fieldIndex).StartPos <= livePos AndAlso livePos < fieldSpans(fieldIndex).EndPos Then

                result.Add(SurgicalCreateFieldAtom(fieldSpans(fieldIndex).KeyText, fieldSpans(fieldIndex).StartPos, fieldSpans(fieldIndex).EndPos))

                Dim skipEnd As Integer = fieldSpans(fieldIndex).EndPos
                Do While index < liveChars.Count AndAlso liveChars(index).LiveStart < skipEnd
                    index += 1
                Loop

                fieldIndex += 1
                Continue Do
            End If

            If liveChars(index).IsTrackedDeletion OrElse String.IsNullOrEmpty(liveChars(index).Text) Then
                index += 1
                Continue Do
            End If

            Dim ch As Char = liveChars(index).Text(0)

            If SurgicalIsParagraphBreakChar(ch) Then
                result.Add(SurgicalCreateTextAtom(vbCr, liveChars(index).LiveStart, liveChars(index).LiveEnd))
                index += 1

                If ch = ControlChars.Cr AndAlso index < liveChars.Count AndAlso
                   Not liveChars(index).IsTrackedDeletion AndAlso
                   liveChars(index).Text = vbLf Then

                    index += 1
                End If

            ElseIf ch = ChrW(11) Then
                result.Add(SurgicalCreateTextAtom(ChrW(11).ToString(), liveChars(index).LiveStart, liveChars(index).LiveEnd))
                index += 1

            ElseIf ch = ChrW(2) Then
                result.Add(SurgicalCreateTextAtom(ChrW(2).ToString(), liveChars(index).LiveStart, liveChars(index).LiveEnd))
                index += 1

            ElseIf SurgicalIsHorizontalWhitespaceChar(ch) Then
                result.Add(SurgicalReadVisibleRunAtom(liveChars, index, Function(c As Char) SurgicalIsHorizontalWhitespaceChar(c)))

            ElseIf SurgicalIsWordChar(ch) Then
                result.Add(SurgicalReadVisibleRunAtom(liveChars, index, Function(c As Char) SurgicalIsWordChar(c)))

            Else
                result.Add(SurgicalCreateTextAtom(ch.ToString(), liveChars(index).LiveStart, liveChars(index).LiveEnd))
                index += 1
            End If
        Loop

        Return result
    End Function

    Private Shared Function SurgicalReadVisibleRunAtom(
    ByVal liveChars As List(Of SurgicalLiveCharAtom),
    ByRef index As Integer,
    ByVal belongsToRun As System.Func(Of Char, Boolean)) As SurgicalTextAtom

        Dim sb As New System.Text.StringBuilder()
        Dim spans As New List(Of SurgicalPositionSpan)()

        Do While index < liveChars.Count
            If liveChars(index).IsTrackedDeletion OrElse String.IsNullOrEmpty(liveChars(index).Text) Then
                index += 1
                Continue Do
            End If

            Dim ch As Char = liveChars(index).Text(0)
            If Not belongsToRun(ch) Then
                Exit Do
            End If

            sb.Append(ch)
            SurgicalAppendSpan(spans, liveChars(index).LiveStart, liveChars(index).LiveEnd)
            index += 1
        Loop

        Return SurgicalCreateTextAtom(sb.ToString(), spans)
    End Function

    Private Shared Function SurgicalCreateTextAtom(ByVal keyText As String, ByVal liveStart As Integer, ByVal liveEnd As Integer) As SurgicalTextAtom
        Dim spans As New List(Of SurgicalPositionSpan) From {
        New SurgicalPositionSpan With {.StartPos = liveStart, .EndPos = liveEnd}
    }

        Return SurgicalCreateTextAtom(keyText, spans)
    End Function

    Private Shared Function SurgicalCreateTextAtom(ByVal keyText As String, ByVal spans As List(Of SurgicalPositionSpan)) As SurgicalTextAtom
        Dim atom As New SurgicalTextAtom With {
        .KeyText = If(keyText, String.Empty),
        .Spans = If(spans, New List(Of SurgicalPositionSpan)())
    }

        If atom.Spans.Count > 0 Then
            atom.LiveStart = atom.Spans(0).StartPos
            atom.LiveEnd = atom.Spans(atom.Spans.Count - 1).EndPos
        End If

        Return atom
    End Function

    Private Shared Function SurgicalCreateFieldAtom(ByVal keyText As String, ByVal liveStart As Integer, ByVal liveEnd As Integer) As SurgicalTextAtom
        Return New SurgicalTextAtom With {
        .KeyText = If(keyText, String.Empty),
        .LiveStart = liveStart,
        .LiveEnd = liveEnd,
        .Spans = New List(Of SurgicalPositionSpan)()
    }
    End Function

    Private Shared Sub SurgicalAppendSpan(ByVal spans As List(Of SurgicalPositionSpan), ByVal startPos As Integer, ByVal endPos As Integer)
        If spans Is Nothing OrElse endPos <= startPos Then
            Return
        End If

        If spans.Count > 0 AndAlso spans(spans.Count - 1).EndPos = startPos Then
            Dim lastSpan As SurgicalPositionSpan = spans(spans.Count - 1)
            lastSpan.EndPos = endPos
            spans(spans.Count - 1) = lastSpan
        Else
            spans.Add(New SurgicalPositionSpan With {.StartPos = startPos, .EndPos = endPos})
        End If
    End Sub

    Private Shared Function BuildSurgicalFieldAtomSpans(ByVal targetRange As Microsoft.Office.Interop.Word.Range) As List(Of SurgicalFieldAtomSpan)
        Dim spans As New List(Of SurgicalFieldAtomSpan)()

        If targetRange Is Nothing Then
            Return spans
        End If

        Try
            Dim doc As Microsoft.Office.Interop.Word.Document = targetRange.Document
            If doc Is Nothing Then
                Return spans
            End If

            For Each field As Microsoft.Office.Interop.Word.Field In doc.Fields
                If field Is Nothing Then
                    Continue For
                End If

                Dim codeRange As Microsoft.Office.Interop.Word.Range = Nothing
                Dim resultRange As Microsoft.Office.Interop.Word.Range = Nothing

                Try
                    codeRange = field.Code
                Catch ex As System.Exception
                    codeRange = Nothing
                End Try

                Try
                    resultRange = field.Result
                Catch ex As System.Exception
                    resultRange = Nothing
                End Try

                If resultRange Is Nothing Then
                    Continue For
                End If

                Dim resultStart As Integer = resultRange.Start
                Dim resultEnd As Integer = resultRange.End

                If resultEnd <= resultStart Then
                    resultEnd = resultStart + 1
                End If

                If resultEnd <= targetRange.Start OrElse resultStart >= targetRange.End Then
                    Continue For
                End If

                Dim spanStart As Integer = System.Math.Max(targetRange.Start, resultStart)
                Dim spanEnd As Integer = System.Math.Min(targetRange.End, resultEnd)

                If spanEnd <= spanStart Then
                    Continue For
                End If

                Dim fieldCode As String = String.Empty
                If codeRange IsNot Nothing Then
                    Try
                        fieldCode = codeRange.Text
                    Catch ex As System.Exception
                        fieldCode = String.Empty
                    End Try
                End If

                Dim keyText As String = BuildCanonicalWfldPlaceholderKey(fieldCode)
                If String.Equals(keyText, "WFLD:", StringComparison.Ordinal) Then
                    Continue For
                End If

                spans.Add(New SurgicalFieldAtomSpan With {
                .StartPos = spanStart,
                .EndPos = spanEnd,
                .KeyText = keyText
            })
            Next
        Catch ex As System.Exception
            Debug.WriteLine("BuildSurgicalFieldAtomSpans failed: " & ex.Message)
        End Try

        spans.Sort(Function(a, b)
                       Dim startCompare As Integer = a.StartPos.CompareTo(b.StartPos)
                       If startCompare <> 0 Then Return startCompare
                       Return b.EndPos.CompareTo(a.EndPos)
                   End Function)

        If spans.Count <= 1 Then
            Return spans
        End If

        Dim filtered As New List(Of SurgicalFieldAtomSpan)()
        For Each span As SurgicalFieldAtomSpan In spans
            If filtered.Count > 0 AndAlso
               span.StartPos = filtered(filtered.Count - 1).StartPos AndAlso
               span.EndPos = filtered(filtered.Count - 1).EndPos AndAlso
               String.Equals(span.KeyText, filtered(filtered.Count - 1).KeyText, StringComparison.Ordinal) Then

                Continue For
            End If

            If filtered.Count > 0 AndAlso span.StartPos < filtered(filtered.Count - 1).EndPos Then
                ' Prefer the first outer/earlier field span. Nested field results must not create
                ' duplicate logical WFLD atoms.
                Continue For
            End If

            filtered.Add(span)
        Next

        Return filtered
    End Function

    Private Shared Function BuildSurgicalLiveCharAtoms(ByVal targetRange As Microsoft.Office.Interop.Word.Range) As List(Of SurgicalLiveCharAtom)
        Dim result As New List(Of SurgicalLiveCharAtom)()

        If targetRange Is Nothing Then
            Return result
        End If

        Dim deletionSpans As List(Of SurgicalPositionSpan) = BuildSurgicalTrackedDeletionSpans(targetRange)
        Dim liveText As String = targetRange.Text
        If liveText Is Nothing Then liveText = String.Empty

        Dim rangeStart As Integer = targetRange.Start
        Dim rangeEnd As Integer = targetRange.End
        Dim positionSpan As Integer = System.Math.Max(0, rangeEnd - rangeStart)

        If liveText.Length = positionSpan Then
            Dim deletionIndex As Integer = 0

            For i As Integer = 0 To liveText.Length - 1
                Dim liveStart As Integer = rangeStart + i
                result.Add(New SurgicalLiveCharAtom With {
                .LiveStart = liveStart,
                .LiveEnd = liveStart + 1,
                .Text = liveText.Substring(i, 1),
                .IsTrackedDeletion = SurgicalPositionIsInsideSpan(liveStart, deletionSpans, deletionIndex)
            })
            Next

            Return result
        End If

        Debug.WriteLine($"SurgicalMarkup: live text/position skew detected; falling back to range-position atom scan. textLen={liveText.Length} posSpan={positionSpan}")

        Dim slowDeletionIndex As Integer = 0
        For pos As Integer = rangeStart To rangeEnd - 1
            Dim piece As String = String.Empty

            Try
                Dim pieceRange As Microsoft.Office.Interop.Word.Range = targetRange.Document.Range(pos, pos + 1)
                piece = pieceRange.Text
            Catch ex As System.Exception
                piece = String.Empty
            End Try

            If String.IsNullOrEmpty(piece) Then
                Continue For
            End If

            If piece.Length = 1 Then
                result.Add(New SurgicalLiveCharAtom With {
                .LiveStart = pos,
                .LiveEnd = pos + 1,
                .Text = piece,
                .IsTrackedDeletion = SurgicalPositionIsInsideSpan(pos, deletionSpans, slowDeletionIndex)
            })
            Else
                For pieceIndex As Integer = 0 To piece.Length - 1
                    result.Add(New SurgicalLiveCharAtom With {
                    .LiveStart = pos,
                    .LiveEnd = pos + 1,
                    .Text = piece.Substring(pieceIndex, 1),
                    .IsTrackedDeletion = SurgicalPositionIsInsideSpan(pos, deletionSpans, slowDeletionIndex)
                })
                Next
            End If
        Next

        Return result
    End Function

    Private Shared Function BuildSurgicalTrackedDeletionSpans(ByVal targetRange As Microsoft.Office.Interop.Word.Range) As List(Of SurgicalPositionSpan)
        Dim spans As New List(Of SurgicalPositionSpan)()

        If targetRange Is Nothing Then
            Return spans
        End If

        Try
            For Each revision As Microsoft.Office.Interop.Word.Revision In targetRange.Revisions
                If revision.Type = Microsoft.Office.Interop.Word.WdRevisionType.wdRevisionDelete Then
                    Dim spanStart As Integer = System.Math.Max(targetRange.Start, revision.Range.Start)
                    Dim spanEnd As Integer = System.Math.Min(targetRange.End, revision.Range.End)

                    If spanEnd > spanStart Then
                        spans.Add(New SurgicalPositionSpan With {
                        .StartPos = spanStart,
                        .EndPos = spanEnd
                    })
                    End If
                End If
            Next
        Catch ex As System.Exception
            Debug.WriteLine("BuildSurgicalTrackedDeletionSpans failed: " & ex.Message)
        End Try

        spans.Sort(Function(a, b) a.StartPos.CompareTo(b.StartPos))

        If spans.Count <= 1 Then
            Return spans
        End If

        Dim merged As New List(Of SurgicalPositionSpan)()
        Dim current As SurgicalPositionSpan = spans(0)

        For i As Integer = 1 To spans.Count - 1
            If spans(i).StartPos <= current.EndPos Then
                current.EndPos = System.Math.Max(current.EndPos, spans(i).EndPos)
            Else
                merged.Add(current)
                current = spans(i)
            End If
        Next

        merged.Add(current)
        Return merged
    End Function

    Private Shared Function SurgicalPositionIsInsideSpan(
    ByVal position As Integer,
    ByVal spans As List(Of SurgicalPositionSpan),
    ByRef spanIndex As Integer) As Boolean

        If spans Is Nothing OrElse spans.Count = 0 Then
            Return False
        End If

        Do While spanIndex < spans.Count AndAlso spans(spanIndex).EndPos <= position
            spanIndex += 1
        Loop

        Return spanIndex < spans.Count AndAlso spans(spanIndex).StartPos <= position AndAlso position < spans(spanIndex).EndPos
    End Function

    Private Shared Function SurgicalTokensForRunText(ByVal runText As String) As List(Of String)
        Dim kind As SpecialPlaceholderKind = SpecialPlaceholderKind.None

        If TryGetSpecialPlaceholderKind(runText, kind) Then
            Select Case kind
                Case SpecialPlaceholderKind.WFNT, SpecialPlaceholderKind.WENT
                    Return New List(Of String) From {ChrW(2).ToString()}
                Case SpecialPlaceholderKind.WFLD
                    Return New List(Of String) From {CanonicalizeSpecialPlaceholderForDiff(runText)}
                Case SpecialPlaceholderKind.PFOR
                    Return New List(Of String)()
            End Select
        End If

        Return SurgicalTokenizeComparableText(NormalizeForWordFind(runText))
    End Function

    Private Shared Function SurgicalTokenizeComparableText(ByVal value As String) As List(Of String)
        Dim result As New List(Of String)()

        If String.IsNullOrEmpty(value) Then
            Return result
        End If

        Dim normalized As String = NormalizeForWordFind(value)
        Dim index As Integer = 0

        Do While index < normalized.Length
            Dim ch As Char = normalized(index)

            If SurgicalIsParagraphBreakChar(ch) Then
                result.Add(vbCr)
                index += 1

                If ch = ControlChars.Cr AndAlso index < normalized.Length AndAlso normalized(index) = ControlChars.Lf Then
                    index += 1
                End If

            ElseIf ch = ChrW(11) Then
                result.Add(ChrW(11).ToString())
                index += 1

            ElseIf ch = ChrW(2) Then
                result.Add(ChrW(2).ToString())
                index += 1

            ElseIf SurgicalIsHorizontalWhitespaceChar(ch) Then
                Dim sb As New System.Text.StringBuilder()

                Do While index < normalized.Length AndAlso SurgicalIsHorizontalWhitespaceChar(normalized(index))
                    sb.Append(normalized(index))
                    index += 1
                Loop

                result.Add(sb.ToString())

            ElseIf SurgicalIsWordChar(ch) Then
                Dim sb As New System.Text.StringBuilder()

                Do While index < normalized.Length AndAlso SurgicalIsWordChar(normalized(index))
                    sb.Append(normalized(index))
                    index += 1
                Loop

                result.Add(sb.ToString())

            Else
                result.Add(ch.ToString())
                index += 1
            End If
        Loop

        Return result
    End Function

    Private Shared Function SurgicalTokensMatchAt(ByVal atoms As List(Of SurgicalTextAtom), ByVal expectedTokens As List(Of String), ByVal atomIndex As Integer) As Boolean
        If expectedTokens Is Nothing Then
            Return True
        End If

        If expectedTokens.Count = 0 Then
            Return True
        End If

        If atoms Is Nothing OrElse atomIndex < 0 OrElse atomIndex + expectedTokens.Count > atoms.Count Then
            Return False
        End If

        For tokenIndex As Integer = 0 To expectedTokens.Count - 1
            If Not String.Equals(atoms(atomIndex + tokenIndex).KeyText, expectedTokens(tokenIndex), StringComparison.Ordinal) Then
                Return False
            End If
        Next

        Return True
    End Function

    Private Shared Function SurgicalTryFindTokenSequence(
    ByVal atoms As List(Of SurgicalTextAtom),
    ByVal expectedTokens As List(Of String),
    ByVal startIndex As Integer,
    ByVal limitIndexExclusive As Integer,
    ByRef foundStart As Integer,
    ByRef foundEnd As Integer) As Boolean

        foundStart = startIndex
        foundEnd = startIndex

        If expectedTokens Is Nothing OrElse expectedTokens.Count = 0 Then
            Return True
        End If

        If atoms Is Nothing OrElse atoms.Count = 0 Then
            Return False
        End If

        Dim safeStart As Integer = System.Math.Max(0, System.Math.Min(startIndex, atoms.Count))
        Dim safeLimit As Integer = System.Math.Max(safeStart, System.Math.Min(limitIndexExclusive, atoms.Count))

        If expectedTokens.Count > safeLimit - safeStart Then
            Return False
        End If

        For candidateStart As Integer = safeStart To safeLimit - expectedTokens.Count
            If SurgicalTokensMatchAt(atoms, expectedTokens, candidateStart) Then
                foundStart = candidateStart
                foundEnd = candidateStart + expectedTokens.Count
                Return True
            End If
        Next

        Return False
    End Function

    Private Shared Function SurgicalFindNextUnchangedAnchorTokens(ByVal runs As List(Of DiffRun), ByVal startRunIndex As Integer) As List(Of String)
        Dim fallback As List(Of String) = Nothing

        If runs Is Nothing Then
            Return New List(Of String)()
        End If

        For runIndex As Integer = System.Math.Max(0, startRunIndex) To runs.Count - 1
            If runs(runIndex).RunType <> ChangeType.Unchanged Then
                Continue For
            End If

            Dim tokens As List(Of String) = SurgicalTokensForRunText(runs(runIndex).Text)
            If tokens.Count = 0 Then
                Continue For
            End If

            If fallback Is Nothing Then
                fallback = tokens
            End If

            If tokens.Any(Function(token) Not SurgicalIsWeakAnchorToken(token)) Then
                Return tokens
            End If
        Next

        If fallback IsNot Nothing Then
            Return fallback
        End If

        Return New List(Of String)()
    End Function

    Private Shared Function SurgicalIsWeakAnchorToken(ByVal token As String) As Boolean
        If String.IsNullOrEmpty(token) Then
            Return True
        End If

        If IsDiffWhitespaceToken(token) Then
            Return True
        End If

        If token = vbCr OrElse token = vbLf OrElse token = vbCrLf OrElse token = ChrW(11).ToString() Then
            Return True
        End If

        Return False
    End Function

    Private Shared Function SurgicalAnchorPositionFromAtomIndex(ByVal atoms As List(Of SurgicalTextAtom), ByVal atomIndex As Integer, ByVal fallbackEnd As Integer) As Integer
        If atoms Is Nothing OrElse atoms.Count = 0 Then
            Return fallbackEnd
        End If

        If atomIndex <= 0 Then
            Return atoms(0).LiveStart
        End If

        If atomIndex >= atoms.Count Then
            Return fallbackEnd
        End If

        Return atoms(atomIndex).LiveStart
    End Function

    Private Shared Sub SurgicalAppendPendingEditForAtomRange(
    ByVal pendingEdits As List(Of PendingSurgicalEdit),
    ByVal atoms As List(Of SurgicalTextAtom),
    ByVal startIndex As Integer,
    ByVal endIndex As Integer,
    ByVal insertText As String,
    ByVal fallbackAnchor As Integer)

        If pendingEdits Is Nothing Then
            Return
        End If

        Dim safeStart As Integer = System.Math.Max(0, startIndex)
        Dim safeEnd As Integer = endIndex

        If atoms IsNot Nothing Then
            safeEnd = System.Math.Min(endIndex, atoms.Count)
        Else
            safeEnd = safeStart
        End If

        Dim insertAttached As Boolean = False
        Dim currentStart As Integer = -1
        Dim currentEnd As Integer = -1

        For atomIndex As Integer = safeStart To safeEnd - 1
            Dim atom As SurgicalTextAtom = atoms(atomIndex)

            If atom.Spans Is Nothing OrElse atom.Spans.Count = 0 Then
                Continue For
            End If

            For Each span As SurgicalPositionSpan In atom.Spans
                If currentStart < 0 Then
                    currentStart = span.StartPos
                    currentEnd = span.EndPos
                ElseIf span.StartPos = currentEnd Then
                    currentEnd = span.EndPos
                Else
                    pendingEdits.Add(New PendingSurgicalEdit With {
                    .DeleteStart = currentStart,
                    .DeleteEnd = currentEnd,
                    .InsertText = If(insertAttached, Nothing, insertText)
                })
                    insertAttached = True

                    currentStart = span.StartPos
                    currentEnd = span.EndPos
                End If
            Next
        Next

        If currentStart >= 0 Then
            pendingEdits.Add(New PendingSurgicalEdit With {
            .DeleteStart = currentStart,
            .DeleteEnd = currentEnd,
            .InsertText = If(insertAttached, Nothing, insertText)
        })
            insertAttached = True
        End If

        If Not insertAttached AndAlso Not String.IsNullOrEmpty(insertText) Then
            pendingEdits.Add(New PendingSurgicalEdit With {
            .DeleteStart = fallbackAnchor,
            .DeleteEnd = fallbackAnchor,
            .InsertText = insertText
        })
        End If
    End Sub

    Private Shared Function SurgicalIsHorizontalWhitespaceChar(ByVal ch As Char) As Boolean
        Return ch = " "c OrElse ch = ControlChars.Tab
    End Function

    Private Shared Function SurgicalIsParagraphBreakChar(ByVal ch As Char) As Boolean
        Return ch = ControlChars.Cr OrElse ch = ControlChars.Lf
    End Function

    Private Shared Function SurgicalIsWordChar(ByVal ch As Char) As Boolean
        If ch = "_"c Then
            Return True
        End If

        Dim category As System.Globalization.UnicodeCategory = System.Globalization.CharUnicodeInfo.GetUnicodeCategory(ch)

        Select Case category
            Case System.Globalization.UnicodeCategory.UppercaseLetter,
                 System.Globalization.UnicodeCategory.LowercaseLetter,
                 System.Globalization.UnicodeCategory.TitlecaseLetter,
                 System.Globalization.UnicodeCategory.ModifierLetter,
                 System.Globalization.UnicodeCategory.OtherLetter,
                 System.Globalization.UnicodeCategory.NonSpacingMark,
                 System.Globalization.UnicodeCategory.SpacingCombiningMark,
                 System.Globalization.UnicodeCategory.DecimalDigitNumber,
                 System.Globalization.UnicodeCategory.LetterNumber,
                 System.Globalization.UnicodeCategory.OtherNumber

                Return True
        End Select

        Return False
    End Function

    ''' <summary>
    ''' Determines whether a pure deletion should absorb one adjacent whitespace run.
    ''' Whitespace is absorbed only for word-like deletions so that accepting the
    ''' tracked deletion does not leave double spaces. Punctuation-only deletions
    ''' deliberately do not absorb whitespace.
    ''' </summary>
    ''' <param name="deletedText">Normalized deleted text for the cluster.</param>
    ''' <returns>
    ''' <see langword="True"/> when adjacent whitespace should be absorbed;
    ''' otherwise <see langword="False"/>.
    ''' </returns>
    Private Shared Function ShouldAbsorbWhitespaceForPureDeletion(ByVal deletedText As String) As Boolean
        If String.IsNullOrEmpty(deletedText) Then
            Return False
        End If

        If IsOnlyBreakCharacters(deletedText) Then
            Return False
        End If

        For Each ch As Char In deletedText
            If Char.IsLetterOrDigit(ch) Then
                Return True
            End If
        Next

        Return False
    End Function

    ''' <summary>
    ''' Checks whether the supplied text consists only of break characters
    ''' used by Word processing in this file, namely carriage return, line feed,
    ''' or manual line break (<c>ChrW(11)</c>).
    ''' </summary>
    ''' <param name="value">Text to inspect.</param>
    ''' <returns>
    ''' <see langword="True"/> when the text contains only break characters;
    ''' otherwise <see langword="False"/>.
    ''' </returns>
    Private Shared Function IsOnlyBreakCharacters(ByVal value As String) As Boolean
        If String.IsNullOrEmpty(value) Then
            Return False
        End If

        For Each ch As Char In value
            If ch <> ControlChars.Cr AndAlso
           ch <> ControlChars.Lf AndAlso
           ch <> ChrW(11) Then

                Return False
            End If
        Next

        Return True
    End Function


    ''' <summary>
    ''' Reprojects the original break-character sequence onto revised text.
    ''' This preserves the original mix of paragraph breaks, line feeds,
    ''' and manual line breaks after LLM post-processing, provided that the
    ''' number of break slots still matches after trimming trailing extras.
    ''' </summary>
    ''' <param name="originalText">Original source text whose break characters must be preserved.</param>
    ''' <param name="revisedText">Revised text returned by the model.</param>
    ''' <returns>
    ''' The revised text with original break characters restored where possible;
    ''' otherwise the unchanged revised text.
    ''' </returns>
    Private Shared Function ReprojectOriginalBreakCharacters(ByVal originalText As String, ByVal revisedText As String) As String
        If String.IsNullOrEmpty(originalText) OrElse String.IsNullOrEmpty(revisedText) Then
            Return revisedText
        End If

        Dim originalBreaks As List(Of String) = ExtractBreakTokens(originalText)
        If originalBreaks.Count = 0 Then
            Return revisedText
        End If

        Dim revisedBreaks As List(Of String) = ExtractBreakTokens(revisedText)
        Dim revisedParts As List(Of String) = SplitTextByBreakTokens(revisedText)

        ' LLMs often add one or more trailing line breaks. Remove only surplus
        ' trailing breaks where the trailing text part is empty.
        Do While revisedBreaks.Count > originalBreaks.Count AndAlso
             revisedParts.Count > 1 AndAlso
             revisedParts(revisedParts.Count - 1).Length = 0

            revisedParts.RemoveAt(revisedParts.Count - 1)
            revisedBreaks.RemoveAt(revisedBreaks.Count - 1)
        Loop

        ' Existing post-processing may have trimmed a trailing break. If the original
        ' had trailing breaks, restore missing trailing break slots without changing text.
        Do While revisedBreaks.Count < originalBreaks.Count AndAlso
             TextEndsWithBreak(originalText)

            revisedBreaks.Add(String.Empty)
            revisedParts.Add(String.Empty)
        Loop

        If revisedBreaks.Count <> originalBreaks.Count Then
            Debug.WriteLine($"BreakReproject skipped: original count={originalBreaks.Count}, revised count={revisedBreaks.Count}")
            Return revisedText
        End If

        Dim sb As New System.Text.StringBuilder(revisedText.Length + originalBreaks.Sum(Function(b) b.Length))

        For i As Integer = 0 To originalBreaks.Count - 1
            If i < revisedParts.Count Then
                sb.Append(revisedParts(i))
            End If

            sb.Append(originalBreaks(i))
        Next

        If revisedParts.Count > originalBreaks.Count Then
            sb.Append(revisedParts(originalBreaks.Count))
        End If

        Dim result As String = sb.ToString()
        Debug.WriteLine($"BreakReproject applied: break count={originalBreaks.Count}")

        Return result
    End Function

    ''' <summary>
    ''' Splits text into content segments separated by break tokens.
    ''' Break tokens themselves are not returned; each list element represents
    ''' the text between two consecutive breaks.
    ''' </summary>
    ''' <param name="value">Text to split.</param>
    ''' <returns>A list of non-break text segments in source order.</returns>

    Private Shared Function SplitTextByBreakTokens(ByVal value As String) As List(Of String)
        Dim result As New List(Of String)
        Dim sb As New System.Text.StringBuilder()

        If String.IsNullOrEmpty(value) Then
            result.Add(String.Empty)
            Return result
        End If

        Dim i As Integer = 0
        Do While i < value.Length
            Dim ch As Char = value.Chars(i)

            If ch = ControlChars.Cr Then
                result.Add(sb.ToString())
                sb.Clear()

                If i + 1 < value.Length AndAlso value.Chars(i + 1) = ControlChars.Lf Then
                    i += 2
                Else
                    i += 1
                End If

            ElseIf ch = ControlChars.Lf OrElse ch = ChrW(11) Then
                result.Add(sb.ToString())
                sb.Clear()
                i += 1

            Else
                sb.Append(ch)
                i += 1
            End If
        Loop

        result.Add(sb.ToString())
        Return result
    End Function

    ''' <summary>
    ''' Determines whether the supplied text ends with any supported break character,
    ''' including carriage return, line feed, or manual line break (<c>ChrW(11)</c>).
    ''' </summary>
    ''' <param name="value">Text to inspect.</param>
    ''' <returns>
    ''' <see langword="True"/> when the text ends with a break character;
    ''' otherwise <see langword="False"/>.
    ''' </returns>
    Private Shared Function TextEndsWithBreak(ByVal value As String) As Boolean
        If String.IsNullOrEmpty(value) Then Return False

        Dim lastChar As Char = value.Chars(value.Length - 1)
        Return lastChar = ControlChars.Cr OrElse
           lastChar = ControlChars.Lf OrElse
           lastChar = ChrW(11)
    End Function


    ''' <summary>
    ''' Extracts all break tokens from the supplied text while preserving their original form,
    ''' including CRLF, CR, LF, and Word manual line breaks (<c>ChrW(11)</c>).
    ''' </summary>
    ''' <param name="value">Input text to scan for break characters.</param>
    ''' <returns>A list of break tokens in source order.</returns>

    Private Shared Function ExtractBreakTokens(ByVal value As String) As List(Of String)
        Dim result As New List(Of String)

        If String.IsNullOrEmpty(value) Then
            Return result
        End If

        Dim i As Integer = 0
        Do While i < value.Length
            Dim ch As Char = value.Chars(i)

            If ch = ControlChars.Cr Then
                If i + 1 < value.Length AndAlso value.Chars(i + 1) = ControlChars.Lf Then
                    result.Add(vbCrLf)
                    i += 2
                Else
                    result.Add(vbCr)
                    i += 1
                End If
            ElseIf ch = ControlChars.Lf Then
                result.Add(vbLf)
                i += 1
            ElseIf ch = ChrW(11) Then
                result.Add(ChrW(11).ToString())
                i += 1
            Else
                i += 1
            End If
        Loop

        Return result
    End Function

    ''' <summary>
    ''' Normalizes text for Word <c>Find</c> operations by converting all line-ending
    ''' variants to Word's internal paragraph marker representation.
    ''' </summary>
    ''' <param name="value">Input text to normalize.</param>
    ''' <returns>Text normalized for Word search operations.</returns>

    Private Shared Function NormalizeForWordFind(ByVal value As String) As String
        If String.IsNullOrEmpty(value) Then
            Return String.Empty
        End If

        Return value.Replace(vbCrLf, vbCr).Replace(vbLf, vbCr)
    End Function

    ''' <summary>
    ''' Normalizes a line-break run for range-based comparisons against Word text.
    ''' Paragraph break variants are converted to Word's internal paragraph marker,
    ''' while manual line breaks are preserved unchanged.
    ''' </summary>
    ''' <param name="value">The break run to normalize.</param>
    ''' <returns>A normalized break sequence suitable for Word range matching.</returns>

    Private Shared Function NormalizeLineBreakRunForWordRange(ByVal value As String) As String
        If String.IsNullOrEmpty(value) Then
            Return String.Empty
        End If

        Dim manualLineBreak As String = ChrW(11).ToString()

        Return value.
        Replace(manualLineBreak, manualLineBreak).
        Replace(vbCrLf, vbCr).
        Replace(vbLf, vbCr)
    End Function


    ''' <summary>
    ''' Executes a Word <c>Find</c> on the supplied range using a plain-text search.
    ''' Long search text is truncated to the configured maximum length to avoid
    ''' expensive or fragile searches on very large fragments.
    ''' </summary>
    ''' <param name="searchRange">The Word range to search.</param>
    ''' <param name="findText">The text to find.</param>
    ''' <param name="wasTruncated">
    ''' Set to <see langword="True"/> when <paramref name="findText"/> was truncated
    ''' before executing the search.
    ''' </param>
    ''' <param name="maxLength">Maximum search-text length allowed for the operation.</param>
    ''' <returns>
    ''' <see langword="True"/> when the text was found; otherwise <see langword="False"/>.
    ''' </returns>

    Private Shared Function ExecuteWordFind(
    ByVal searchRange As Range,
    ByVal findText As String,
    ByRef wasTruncated As Boolean,
    Optional ByVal maxLength As Integer = 200) As Boolean

        wasTruncated = False

        If searchRange Is Nothing OrElse String.IsNullOrEmpty(findText) Then
            Return False
        End If

        If findText.Length > maxLength Then
            findText = findText.Substring(0, maxLength)
            wasTruncated = True
        End If

        With searchRange.Find
            .ClearFormatting()
            .Text = findText
            .Forward = True
            .Wrap = WdFindWrap.wdFindStop
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            Return .Execute()
        End With
    End Function

    ''' <summary>
    ''' Loads the live text of [<paramref name="startPos"/>, <paramref name="endPos"/>) into an
    ''' in-memory cache used by the surgical navigation loop, avoiding one COM call per character.
    ''' The cache is only marked valid when its length equals the position span, i.e. when one
    ''' document position maps to exactly one text character (no field/footnote skew). When invalid,
    ''' callers fall back to the original per-character COM navigation so behavior is unchanged.
    ''' </summary>
    ''' <param name="doc">Document being patched.</param>
    ''' <param name="startPos">Start position of the cached window.</param>
    ''' <param name="endPos">End position of the cached window.</param>
    ''' <param name="cacheText">Receives the cached text (never Nothing).</param>
    ''' <param name="cacheValid">Receives whether the cache is safe to index by position.</param>
    Private Shared Sub LoadSurgicalNavCache(ByVal doc As Word.Document, ByVal startPos As Integer, ByVal endPos As Integer, ByRef cacheText As String, ByRef cacheValid As Boolean)
        Try
            If doc Is Nothing OrElse endPos <= startPos Then
                cacheText = String.Empty
                cacheValid = True
                Return
            End If

            cacheText = doc.Range(startPos, endPos).Text
            If cacheText Is Nothing Then cacheText = String.Empty

            ' Trust the cache only when 1 position == 1 character across the window.
            cacheValid = (cacheText.Length = (endPos - startPos))
        Catch ex As System.Exception
            cacheText = String.Empty
            cacheValid = False
        End Try
    End Sub


    ''' <summary>
    ''' Collapses visible double spaces within the edited range while the document is shown in
    ''' Final view, preventing whitespace artifacts left behind by tracked deletions or replacements.
    ''' </summary>
    ''' <param name="doc">Document containing the edited range.</param>
    ''' <param name="startPos">Approximate start of the cleanup window.</param>
    ''' <param name="endPos">Approximate end of the cleanup window.</param>
    Private Shared Sub CollapseDoubleSpacesInFinalView(
    ByVal doc As Word.Document,
    ByVal startPos As Integer,
    ByVal endPos As Integer)

        If doc Is Nothing Then
            Return
        End If

        Dim safeStart As Integer = System.Math.Max(doc.Content.Start, System.Math.Min(startPos - 8, doc.Content.End))
        Dim safeEnd As Integer = System.Math.Max(safeStart, System.Math.Min(endPos + 8, doc.Content.End))

        With doc.ActiveWindow.View
            .RevisionsView = WdRevisionsView.wdRevisionsViewFinal
            .ShowRevisionsAndComments = False
        End With

        Dim cleanupRange As Range = doc.Range(safeStart, safeEnd)
        With cleanupRange.Find
            .ClearFormatting()
            .Replacement.ClearFormatting()
            .Text = "  "
            .Replacement.Text = " "
            .Forward = True
            .Wrap = WdFindWrap.wdFindStop
            .Format = False
            .MatchWildcards = False
        End With

        cleanupRange.Find.Execute(Replace:=WdReplace.wdReplaceAll)

        With doc.ActiveWindow.View
            .RevisionsView = WdRevisionsView.wdRevisionsViewFinal
            .ShowRevisionsAndComments = True
        End With
    End Sub

    ''' <summary>
    ''' Determines whether a token is one of the synthetic placeholder tokens created for the
    ''' surgical diff pipeline, e.g. <c>[[MF0]]</c>.
    ''' </summary>
    ''' <param name="value">Token text to inspect.</param>
    ''' <returns>
    ''' <see langword="True"/> when the token represents a protected placeholder;
    ''' otherwise <see langword="False"/>.
    ''' </returns>
    Private Shared Function IsDiffPlaceholderToken(ByVal value As String) As Boolean
        If String.IsNullOrEmpty(value) Then Return False
        If value.Length < 7 Then Return False
        If Not value.StartsWith("[[MF", StringComparison.Ordinal) Then Return False
        If Not value.EndsWith("]]", StringComparison.Ordinal) Then Return False

        Dim tokenIndexText As String = value.Substring(4, value.Length - 6)
        Dim tokenIndex As Integer
        Return Integer.TryParse(tokenIndexText, tokenIndex)
    End Function

    ''' <summary>
    ''' Determines whether a token represents a structural break unit used by the surgical diff
    ''' pipeline, including normalized break markers and their restored runtime equivalents.
    ''' </summary>
    ''' <param name="value">Token text to inspect.</param>
    ''' <returns>
    ''' <see langword="True"/> when the token is treated as a line-break token;
    ''' otherwise <see langword="False"/>.
    ''' </returns>
    Private Shared Function IsDiffLineBreakToken(ByVal value As String) As Boolean
        If String.IsNullOrEmpty(value) Then Return False

        Return value = "{vbCrLf}" OrElse
           value = "{vbCr}" OrElse
           value = "{vbLf}" OrElse
           value = "{vbVt}" OrElse
           value = vbCrLf OrElse
           value = vbCr OrElse
           value = vbLf OrElse
           value = ChrW(11).ToString()
    End Function


    ''' <summary>
    ''' Tokenizes text into diff units used by the surgical diff pipeline.
    ''' Tokens include placeholder markers, normalized break markers, whitespace,
    ''' word-like runs, and single-character punctuation.
    ''' </summary>
    ''' <param name="text">Text to tokenize.</param>
    ''' <returns>A list of diff tokens in source order.</returns>
    Private Shared Function TokenizeDiffUnits(ByVal text As String) As List(Of String)
        Dim result As New List(Of String)

        If String.IsNullOrEmpty(text) Then
            Return result
        End If

        Dim pattern As String =
        "\[\[MF\d+\]\]|\{vbCrLf\}|\{vbCr\}|\{vbLf\}|\{vbVt\}|\s+|[\p{L}\p{M}\p{N}_]+|[^\s]"

        For Each m As Match In Regex.Matches(text, pattern, RegexOptions.Singleline)
            If m.Success AndAlso m.Length > 0 Then
                result.Add(m.Value)
            End If
        Next

        Return result
    End Function

    ''' <summary>
    ''' Determines whether a token consists only of horizontal whitespace
    ''' that should be treated as a standalone diff token.
    ''' </summary>
    ''' <param name="value">Token text to inspect.</param>
    ''' <returns>
    ''' <see langword="True"/> when the token contains only spaces or tabs;
    ''' otherwise <see langword="False"/>.
    ''' </returns>
    Private Shared Function IsDiffWhitespaceToken(ByVal value As String) As Boolean
        If String.IsNullOrEmpty(value) Then Return False

        For Each ch As Char In value
            If ch <> " "c AndAlso ch <> vbTab Then
                Return False
            End If
        Next

        Return True
    End Function

    ''' <summary>
    ''' Advances the surgical cursor past a special placeholder already present in the
    ''' Word document, such as field, footnote, endnote, or paragraph-format markers.
    ''' This keeps unchanged placeholders aligned with the live document text.
    ''' </summary>
    ''' <param name="doc">The Word document being edited.</param>
    ''' <param name="cursor">The current live cursor range used by the surgical patcher.</param>
    ''' <param name="rangeEnd">Upper bound of the editable target range.</param>
    ''' <param name="placeholderText">The placeholder token from the diff stream.</param>
    Private Shared Sub AdvanceCursorPastSpecialPlaceholder(
    ByVal doc As Word.Document,
    ByVal cursor As Word.Range,
    ByVal rangeEnd As Integer,
    ByVal placeholderText As String)

        If doc Is Nothing OrElse cursor Is Nothing OrElse String.IsNullOrWhiteSpace(placeholderText) Then
            Return
        End If

        Dim kind As SpecialPlaceholderKind = SpecialPlaceholderKind.None
        If Not TryGetSpecialPlaceholderKind(placeholderText, kind) Then
            Return
        End If

        Select Case kind
            Case SpecialPlaceholderKind.PFOR
                Return

            Case SpecialPlaceholderKind.WFNT, SpecialPlaceholderKind.WENT
                Dim probeEnd As Integer = System.Math.Min(rangeEnd, cursor.Start + 2)
                Dim pos As Integer = cursor.Start

                Do While pos < probeEnd
                    Dim probe As Word.Range = doc.Range(pos, System.Math.Min(pos + 1, rangeEnd))
                    If probe.Text = ChrW(2) Then
                        cursor.SetRange(pos + 1, pos + 1)
                        Return
                    End If
                    pos += 1
                Loop

                Dim fallbackPos As Integer = System.Math.Min(cursor.Start + 1, rangeEnd)
                cursor.SetRange(fallbackPos, fallbackPos)

            Case SpecialPlaceholderKind.WFLD
                Return
        End Select
    End Sub

    ''' <summary>
    ''' Determines whether a fully restored diff run consists of a single special placeholder
    ''' such as a field, footnote, endnote, or paragraph-format marker.
    ''' </summary>
    ''' <param name="value">Run text to inspect.</param>
    ''' <returns>
    ''' <see langword="True"/> when the run is exactly one recognized special placeholder;
    ''' otherwise <see langword="False"/>.
    ''' </returns>
    Private Shared Function IsSpecialPlaceholderRun(value As String) As Boolean
        Dim kind As SpecialPlaceholderKind = SpecialPlaceholderKind.None
        Return TryGetSpecialPlaceholderKind(value, kind)
    End Function

    ''' <summary>
    ''' Parses the placeholder kind prefix of a restored placeholder token and maps it to the
    ''' internal <c>SpecialPlaceholderKind</c> enumeration.
    ''' </summary>
    ''' <param name="value">Placeholder text to inspect.</param>
    ''' <param name="kind">Receives the resolved placeholder kind when parsing succeeds.</param>
    ''' <returns>
    ''' <see langword="True"/> when the placeholder kind could be recognized;
    ''' otherwise <see langword="False"/>.
    ''' </returns>
    Private Shared Function TryGetSpecialPlaceholderKind(
    ByVal value As String,
    ByRef kind As SpecialPlaceholderKind) As Boolean

        kind = SpecialPlaceholderKind.None

        If String.IsNullOrWhiteSpace(value) Then
            Return False
        End If

        Dim s As String = value.Trim()

        If s.Length < 10 Then
            Return False
        End If

        If Not s.StartsWith("{{", StringComparison.Ordinal) OrElse Not s.EndsWith("}}", StringComparison.Ordinal) Then
            Return False
        End If

        Dim colonPos As Integer = s.IndexOf(":"c)
        If colonPos <= 2 Then
            Return False
        End If

        Dim kindText As String = s.Substring(2, colonPos - 2)

        Select Case kindText
            Case "WFLD"
                kind = SpecialPlaceholderKind.WFLD
            Case "WFNT"
                kind = SpecialPlaceholderKind.WFNT
            Case "WENT"
                kind = SpecialPlaceholderKind.WENT
            Case "PFOR"
                kind = SpecialPlaceholderKind.PFOR
            Case Else
                Return False
        End Select

        Return True
    End Function

    Private Shared Function TryGetSpecialPlaceholderPayload(
    ByVal value As String,
    ByRef kind As SpecialPlaceholderKind,
    ByRef payload As String) As Boolean

        kind = SpecialPlaceholderKind.None
        payload = String.Empty

        If String.IsNullOrWhiteSpace(value) Then
            Return False
        End If

        Dim s As String = value.Trim()

        If Not s.StartsWith("{{", StringComparison.Ordinal) OrElse Not s.EndsWith("}}", StringComparison.Ordinal) Then
            Return False
        End If

        Dim colonPos As Integer = s.IndexOf(":"c)
        If colonPos <= 2 OrElse colonPos >= s.Length - 2 Then
            Return False
        End If

        Dim kindText As String = s.Substring(2, colonPos - 2)

        Select Case kindText
            Case "WFLD"
                kind = SpecialPlaceholderKind.WFLD
            Case "WFNT"
                kind = SpecialPlaceholderKind.WFNT
            Case "WENT"
                kind = SpecialPlaceholderKind.WENT
            Case "PFOR"
                kind = SpecialPlaceholderKind.PFOR
            Case Else
                Return False
        End Select

        payload = s.Substring(colonPos + 1, s.Length - colonPos - 3)
        Return True
    End Function

    Private Shared Function CanonicalizeSpecialPlaceholderForDiff(ByVal placeholder As String) As String
        Dim kind As SpecialPlaceholderKind = SpecialPlaceholderKind.None
        Dim payload As String = String.Empty

        If Not TryGetSpecialPlaceholderPayload(placeholder, kind, payload) Then
            Return If(placeholder, String.Empty)
        End If

        If kind = SpecialPlaceholderKind.WFLD Then
            Return BuildCanonicalWfldPlaceholderKey(payload)
        End If

        Return If(placeholder, String.Empty)
    End Function

    Private Shared Function BuildCanonicalWfldPlaceholderKey(ByVal fieldCode As String) As String
        Return "WFLD:" & NormalizeWordFieldCodeForDiff(fieldCode)
    End Function

    Private Shared Function NormalizeWordFieldCodeForDiff(ByVal fieldCode As String) As String
        If fieldCode Is Nothing Then
            Return String.Empty
        End If

        Dim sb As New System.Text.StringBuilder()
        Dim inQuote As Boolean = False
        Dim pendingWhitespace As Boolean = False

        For Each ch As Char In fieldCode.Trim()
            If ch = """"c Then
                If pendingWhitespace AndAlso sb.Length > 0 Then
                    sb.Append(" "c)
                    pendingWhitespace = False
                End If

                sb.Append(ch)
                inQuote = Not inQuote
            ElseIf inQuote Then
                sb.Append(ch)
            ElseIf Char.IsWhiteSpace(ch) Then
                pendingWhitespace = True
            Else
                If pendingWhitespace AndAlso sb.Length > 0 Then
                    sb.Append(" "c)
                End If

                sb.Append(ch)
                pendingWhitespace = False
            End If
        Next

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' Converts spaces and control characters into a debug-visible representation so token and run
    ''' traces can be read unambiguously in the Output window.
    ''' </summary>
    ''' <param name="value">Text to visualize for diagnostics.</param>
    ''' <returns>A printable diagnostic representation of the supplied text.</returns>
    <Conditional("DEBUG")>
    Private Shared Function DebugVisualizeToken(ByVal value As String) As String
        If value Is Nothing Then Return "<Nothing>"

        Return value.
        Replace(" ", "·").
        Replace(vbTab, "\t").
        Replace(vbCrLf, "\r\n").
        Replace(vbCr, "\r").
        Replace(vbLf, "\n")
    End Function


    ''' <summary>
    ''' Writes a numbered diagnostic dump of diff tokens to the debug output.
    ''' </summary>
    ''' <param name="caption">Heading written before the token list.</param>
    ''' <param name="tokens">Token sequence to dump.</param>
    <Conditional("DEBUG")>
    Private Shared Sub DebugDumpTokenList(ByVal caption As String, ByVal tokens As IEnumerable(Of String))
        Debug.WriteLine("==== " & caption & " ====")

        Dim i As Integer = 0
        For Each token As String In tokens
            Debug.WriteLine($"{i:000}: '{DebugVisualizeToken(token)}'")
            i += 1
        Next

        If i = 0 Then
            Debug.WriteLine("<empty>")
        End If
    End Sub

    ''' <summary>
    ''' Writes a numbered diagnostic dump of diff runs to the debug output.
    ''' </summary>
    ''' <param name="caption">Heading written before the run list.</param>
    ''' <param name="runs">Run sequence to dump.</param>
    <Conditional("DEBUG")>
    Private Shared Sub DebugDumpRuns(ByVal caption As String, ByVal runs As IEnumerable(Of DiffRun))
        Debug.WriteLine("==== " & caption & " ====")

        Dim i As Integer = 0
        For Each run As DiffRun In runs
            Debug.WriteLine($"{i:000}: {run.RunType} => '{DebugVisualizeToken(run.Text)}'")
            i += 1
        Next

        If i = 0 Then
            Debug.WriteLine("<empty>")
        End If
    End Sub


    ''' <summary>
    ''' Replaces all protected placeholder payloads in both texts with stable shared placeholder
    ''' tokens so the diff engine treats them as atomic units.
    ''' </summary>
    ''' <param name="text1">Original text to rewrite in place.</param>
    ''' <param name="text2">Revised text to rewrite in place.</param>
    ''' <returns>
    ''' A list mapping placeholder-token indexes back to their original placeholder text.
    ''' </returns>
    Private Shared Function TokenizeSpecialPlaceholdersForDiff(
    ByRef text1 As String,
    ByRef text2 As String) As List(Of String)

        Dim placeholdersByIndex As New List(Of String)
        Dim tokenByPlaceholderKey As New Dictionary(Of String, String)(StringComparer.Ordinal)

        Dim nextIndex As Integer = 0

        Dim replacePlaceholder As MatchEvaluator =
        Function(m As Match) As String
            Dim placeholder As String = m.Value
            Dim placeholderKey As String = CanonicalizeSpecialPlaceholderForDiff(placeholder)
            Dim token As String = Nothing

            If Not tokenByPlaceholderKey.TryGetValue(placeholderKey, token) Then
                token = $"[[MF{nextIndex}]]"
                tokenByPlaceholderKey.Add(placeholderKey, token)
                placeholdersByIndex.Add(placeholder)
                nextIndex += 1
            End If

            Return token
        End Function

        text1 = System.Text.RegularExpressions.Regex.Replace(
        text1,
        "\{\{.*?\}\}",
        replacePlaceholder,
        RegexOptions.Singleline)

        text2 = System.Text.RegularExpressions.Regex.Replace(
        text2,
        "\{\{.*?\}\}",
        replacePlaceholder,
        RegexOptions.Singleline)

        Return placeholdersByIndex
    End Function

    ''' <summary>
    ''' Restores synthetic placeholder tokens such as <c>[[MF0]]</c> back to their original
    ''' placeholder payloads after diff processing.
    ''' </summary>
    ''' <param name="input">Text containing synthetic placeholder tokens.</param>
    ''' <param name="placeholdersByIndex">Index-to-placeholder mapping created during tokenization.</param>
    ''' <returns>The input text with all known placeholder tokens restored.</returns>
    Private Shared Function RestoreTokenizedSpecialPlaceholders(
    ByVal input As String,
    ByVal placeholdersByIndex As List(Of String)) As String

        If String.IsNullOrEmpty(input) OrElse placeholdersByIndex Is Nothing OrElse placeholdersByIndex.Count = 0 Then
            Return input
        End If

        For idx As Integer = 0 To placeholdersByIndex.Count - 1
            input = input.Replace($"[[MF{idx}]]", placeholdersByIndex(idx))
        Next

        Return input
    End Function

    Private Structure InlineCharFormatSnapshot
        Public FontName As String
        Public FontSize As Single?
        Public Bold As Integer?
        Public Italic As Integer?
        Public Underline As Word.WdUnderline?
        Public Color As Long?
    End Structure

    Private Shared Function CaptureInlineCharFormatSnapshot(sourceRange As Word.Range) As InlineCharFormatSnapshot
        Dim result As New InlineCharFormatSnapshot

        If sourceRange Is Nothing Then Return result

        Try
            With sourceRange.Font
                If Not String.IsNullOrWhiteSpace(.Name) AndAlso .Name <> CStr(Word.WdConstants.wdUndefined) Then
                    result.FontName = .Name
                End If

                If .Size <> CSng(Word.WdConstants.wdUndefined) AndAlso .Size > 0 Then
                    result.FontSize = .Size
                End If

                If .Bold <> Word.WdConstants.wdUndefined Then
                    result.Bold = .Bold
                End If

                If .Italic <> Word.WdConstants.wdUndefined Then
                    result.Italic = .Italic
                End If

                If .Underline <> Word.WdConstants.wdUndefined Then
                    result.Underline = CType(.Underline, Word.WdUnderline)
                End If

                If .Color <> Word.WdConstants.wdUndefined Then
                    result.Color = .Color
                End If
            End With
        Catch ex As System.Exception
            Debug.WriteLine("CaptureInlineCharFormatSnapshot failed: " & ex.Message)
        End Try

        Return result
    End Function

End Class
