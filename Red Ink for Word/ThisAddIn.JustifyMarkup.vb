' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.JustifyMarkup.vb
' Purpose: Implements "Justify Markup" — stores original clause snapshots,
'          compares them against marked-up versions via LLM, and inserts
'          justification comments as Word bubbles.
'
' Architecture:
'  - Snapshot Storage: Maintains an in-memory list of named clause snapshots
'    (OriginalClauseSnapshots) that persist for the Word session lifetime.
'  - Store Original: StoreOriginalClause captures selected text with a user-
'    supplied label and stores it in the snapshot list.
'  - Justify Markup: JustifyMarkup lets the user pick a stored snapshot, sends
'    original + marked-up text to the LLM, and inserts the justification as a
'    Word comment bubble on the selection.
'  - Shared Bubble Insertion: InsertJustificationBubble is a reusable helper
'    that adds a justification comment to a given range, following the same
'    patterns as InsertMarkupReviewResults (direct range with markdown, fallback
'    to SetBubbles, last-resort AddNoticeBubbleAt).
'  - BalloonMerge Integration: BalloonMergeWithJustification extends BalloonMerge
'    by capturing original text before merging, then auto-justifying after.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.Text
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Word
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ' ========================= Data Types =========================

    ''' <summary>
    ''' Represents a stored snapshot of original clause text before markup.
    ''' </summary>
    Public Class ClauseSnapshot
        ''' <summary>User-supplied label identifying this snapshot.</summary>
        Public Property Label As String
        ''' <summary>Original clause text captured at snapshot time.</summary>
        Public Property OriginalText As String
        ''' <summary>Timestamp when the snapshot was taken.</summary>
        Public Property StoredAt As DateTime

        Public Sub New(label As String, originalText As String)
            Me.Label = label
            Me.OriginalText = originalText
            Me.StoredAt = DateTime.Now
        End Sub

        Public Overrides Function ToString() As String
            Dim preview As String = If(OriginalText.Length > 60, OriginalText.Substring(0, 60) & "...", OriginalText)
            Return $"{Label} ({StoredAt:HH:mm:ss}) - {preview}"
        End Function
    End Class

    ''' <summary>
    ''' In-memory list of original clause snapshots. Persists for the Word session lifetime.
    ''' </summary>
    Public Shared OriginalClauseSnapshots As New System.Collections.Generic.List(Of ClauseSnapshot)()

    ' ========================= Visible Text Helper =========================

    ''' <summary>
    ''' Returns the "accepted" view of a range's text by temporarily selecting it,
    ''' switching to Final view, and reading Selection.Text. Word re-evaluates the
    ''' live Selection in Final view, excluding deleted revisions and keeping inserts.
    ''' Restores the original view and selection afterwards.
    ''' Falls back to raw Range.Text when there are no revisions or on error.
    ''' </summary>
    ''' <param name="src">Word range to extract visible text from.</param>
    ''' <param name="restoreRange">Optional range to re-select after reading. Use this
    ''' to restore the selection to a comment balloon or other non-main-story location
    ''' that cannot be restored via SetRange.</param>
    ''' <param name="capturedRange">Output: the range that was selected in Final view,
    ''' re-acquired in the restored view. Use this for anchoring comments.</param>
    ''' <returns>Text as it would appear after accepting all tracked changes.</returns>
    Private Function GetVisibleTextSafe(ByVal src As Range,
                                    Optional restoreRange As Range = Nothing,
                                    Optional ByRef capturedRange As Range = Nothing) As String
        capturedRange = Nothing
        If src Is Nothing Then Return String.Empty

        Dim raw As String
        Try
            raw = src.Text
            If String.IsNullOrEmpty(raw) Then Return String.Empty
        Catch
            Return String.Empty
        End Try

        Try
            If src.Revisions.Count = 0 Then
                capturedRange = src
                Return raw
            End If
        Catch
            capturedRange = src
            Return raw
        End Try

        Try
            Dim app As Microsoft.Office.Interop.Word.Application = src.Application
            Dim docObj As Microsoft.Office.Interop.Word.Document = src.Document
            Dim docWindow As Microsoft.Office.Interop.Word.Window = docObj.ActiveWindow

            If docWindow Is Nothing Then
                capturedRange = src
                Return raw
            End If

            Dim view As Microsoft.Office.Interop.Word.View = docWindow.View

            Dim origRevView As WdRevisionsView = view.RevisionsView
            Dim origShowRevs As Boolean = view.ShowRevisionsAndComments

            Dim origActiveWindow As Microsoft.Office.Interop.Word.Window = Nothing
            Try
                origActiveWindow = app.ActiveWindow
            Catch
            End Try

            Dim origSelStart As Integer = 0
            Dim origSelEnd As Integer = 0
            Dim haveOrigSelection As Boolean = False
            Try
                origSelStart = app.Selection.Start
                origSelEnd = app.Selection.End
                haveOrigSelection = True
            Catch
            End Try

            Try
                If origActiveWindow Is Nothing OrElse Not Object.ReferenceEquals(origActiveWindow, docWindow) Then
                    docWindow.Activate()
                End If

                src.Select()

                view.RevisionsView = WdRevisionsView.wdRevisionsViewFinal
                view.ShowRevisionsAndComments = False

                Dim finalText As String = app.Selection.Text
                If finalText Is Nothing Then finalText = String.Empty

                view.RevisionsView = origRevView
                view.ShowRevisionsAndComments = origShowRevs

                src.Select()
                capturedRange = docObj.Range(app.Selection.Start, app.Selection.End)

                Try
                    If restoreRange IsNot Nothing Then
                        Dim restoreWindow As Microsoft.Office.Interop.Word.Window = restoreRange.Document.ActiveWindow
                        If restoreWindow IsNot Nothing AndAlso Not Object.ReferenceEquals(app.ActiveWindow, restoreWindow) Then
                            restoreWindow.Activate()
                        End If
                        restoreRange.Select()
                    ElseIf haveOrigSelection Then
                        If origActiveWindow IsNot Nothing AndAlso Not Object.ReferenceEquals(app.ActiveWindow, origActiveWindow) Then
                            origActiveWindow.Activate()
                        End If
                        app.Selection.SetRange(origSelStart, origSelEnd)
                    End If
                Catch
                End Try

                Return finalText
            Catch
                Try
                    view.RevisionsView = origRevView
                    view.ShowRevisionsAndComments = origShowRevs
                Catch
                End Try

                Try
                    If restoreRange IsNot Nothing Then
                        Dim restoreWindow As Microsoft.Office.Interop.Word.Window = restoreRange.Document.ActiveWindow
                        If restoreWindow IsNot Nothing AndAlso Not Object.ReferenceEquals(app.ActiveWindow, restoreWindow) Then
                            restoreWindow.Activate()
                        End If
                        restoreRange.Select()
                    ElseIf haveOrigSelection Then
                        If origActiveWindow IsNot Nothing AndAlso Not Object.ReferenceEquals(app.ActiveWindow, origActiveWindow) Then
                            origActiveWindow.Activate()
                        End If
                        app.Selection.SetRange(origSelStart, origSelEnd)
                    End If
                Catch
                End Try

                capturedRange = src
                Return raw
            End Try
        Catch
            capturedRange = src
            Return raw
        End Try
    End Function
    ' ========================= Store Original =========================

    ''' <summary>
    ''' Stores the currently selected text as an original clause snapshot.
    ''' Uses visible text (treating existing tracked changes as accepted) so the
    ''' snapshot reflects what the user sees. The user is prompted for a label.
    ''' </summary>
    Public Sub StoreOriginalClause()
        Try
            Dim app As Word.Application = Globals.ThisAddIn.Application
            Dim sel As Microsoft.Office.Interop.Word.Selection = app.Selection

            If sel.Type = WdSelectionType.wdSelectionIP Then
                ShowCustomMessageBox("Please select the clause text you want to store as the original before making markups.")
                Return
            End If

            Dim selectedText As String = GetVisibleTextSafe(sel.Range)
            If String.IsNullOrWhiteSpace(selectedText) Then
                ShowCustomMessageBox("The selection is empty. Please select the clause text to store.")
                Return
            End If

            Dim label As String = SLib.ShowCustomInputBox(
                "Enter a label for this original clause snapshot:",
                $"{AN} Store Original Clause",
                True,
                GetNextClauseLabel()).Trim()

            If String.IsNullOrEmpty(label) OrElse label = "ESC" Then Return

            OriginalClauseSnapshots.Add(New ClauseSnapshot(label, selectedText))

        Catch ex As System.Exception
            MessageBox.Show($"Error in StoreOriginalClause:{Environment.NewLine}{ex.Message}",
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Returns the next available "Clause N" label that does not collide with
    ''' any existing snapshot label in OriginalClauseSnapshots.
    ''' </summary>
    Private Shared Function GetNextClauseLabel() As String
        Dim n As Integer = 1
        Do
            Dim candidate As String = $"Clause {n}"
            Dim exists As Boolean = False
            For Each snap In OriginalClauseSnapshots
                If String.Equals(snap.Label, candidate, StringComparison.OrdinalIgnoreCase) Then
                    exists = True
                    Exit For
                End If
            Next
            If Not exists Then Return candidate
            n += 1
        Loop
    End Function

    ' ========================= Justify Markup =========================

    ''' <summary>
    ''' Justifies markups in the selected text by comparing against a stored original
    ''' clause snapshot. Sends both texts to the LLM and inserts the justification
    ''' as a Word comment bubble on the selection.
    ''' </summary>
    Public Async Sub JustifyMarkup()
        Try
            If INILoadFail() Then Return
            Dim app As Word.Application = Globals.ThisAddIn.Application
            Dim sel As Microsoft.Office.Interop.Word.Selection = app.Selection
            Dim doc As Microsoft.Office.Interop.Word.Document = sel.Range.Document

            If sel.Type = WdSelectionType.wdSelectionIP Then
                ShowCustomMessageBox("Please select the marked-up clause text you want to justify.")
                Return
            End If

            Dim markedUpText As String = SafeRangeText(sel.Range)
            If String.IsNullOrWhiteSpace(markedUpText) Then
                ShowCustomMessageBox("The selection is empty. Please select the marked-up clause text.")
                Return
            End If

            If OriginalClauseSnapshots.Count = 0 Then
                ShowCustomMessageBox("No original clause snapshots stored. Please use 'Store Original Clause' first to capture the text before markup.")
                Return
            End If

            ' Get the revised (visible) text — treats all tracked changes as accepted
            ' so the LLM sees what the clause looks like after markup.
            Dim revisedText As String = GetVisibleTextSafe(sel.Range)
            If String.IsNullOrWhiteSpace(revisedText) Then
                revisedText = markedUpText ' fall back to raw text if extraction fails
            End If

            ' Pick the original snapshot — auto-select when only one exists
            Dim picked As Integer
            If OriginalClauseSnapshots.Count = 1 Then
                picked = 1
            Else
                Dim items(OriginalClauseSnapshots.Count - 1) As SelectionItem
                For i As Integer = 0 To OriginalClauseSnapshots.Count - 1
                    items(i) = New SelectionItem(OriginalClauseSnapshots(i).ToString(), i + 1)
                Next
                picked = SelectValue(items, 1, "Choose the original clause to compare against...")
                If picked < 1 Then Return
            End If

            Dim snapshot As ClauseSnapshot = OriginalClauseSnapshots(picked - 1)

            ' Ask whether to include entire document as context
            Dim includeDoc As Integer = ShowCustomYesNoBox(
                "Do you want to provide the entire document to the LLM for additional context?",
                "Yes", "No")
            If includeDoc = 0 Then Return

            Dim documentContext As String = String.Empty
            If includeDoc = 1 Then
                documentContext = SafeRangeText(doc.Content)
            End If

            ' Call LLM for justification
            Dim justification As String = Await GetJustificationFromLLM(
                snapshot.OriginalText, revisedText, documentContext)

            If String.IsNullOrWhiteSpace(justification) Then
                ShowCustomMessageBox("The LLM did not return a justification. Please try again.")
                Return
            End If

            ' Insert justification bubble on the selection
            Await EnsureUIThread()

            ' Grab a stable range from the document (not the live Selection)
            Dim anchorRange As Microsoft.Office.Interop.Word.Range = doc.Range(sel.Range.Start, sel.Range.End)
            InsertJustificationBubble(doc, anchorRange, justification, "")

            ' Remove the used snapshot from the list
            OriginalClauseSnapshots.RemoveAt(picked - 1)

        Catch ex As System.Exception
            MessageBox.Show($"Error in JustifyMarkup:{Environment.NewLine}{ex.Message}",
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' ========================= LLM Call =========================

    ''' <summary>
    ''' Calls the LLM to generate a justification for changes between original and marked-up text.
    ''' </summary>
    ''' <param name="originalText">The original clause text before markup.</param>
    ''' <param name="markedUpText">The clause text after markup/changes.</param>
    ''' <param name="documentContext">Optional full document text for context. Empty string if not provided.</param>
    ''' <param name="contextLabel">XML tag name for the context block. Defaults to "DOCUMENTCONTEXT".</param>
    ''' <param name="contextSuffix">Instruction appended after the context block. Defaults to document-context guidance.</param>
    ''' <returns>The LLM-generated justification text.</returns>
    Private Async Function GetJustificationFromLLM(
        originalText As String,
        markedUpText As String,
        documentContext As String,
        Optional contextLabel As String = "DOCUMENTCONTEXT",
        Optional contextSuffix As String = "Use the document context above to better understand the purpose and setting of the clause.") As System.Threading.Tasks.Task(Of String)

        Dim systemPrompt As String = SP_JustifyMarkup

        Dim userPrompt As String =
            "<ORIGINALCLAUSE>" & Environment.NewLine &
            originalText & Environment.NewLine &
            "</ORIGINALCLAUSE>" & Environment.NewLine &
            "<REVISEDCLAUSE>" & Environment.NewLine &
            markedUpText & Environment.NewLine &
            "</REVISEDCLAUSE>"

        If Not String.IsNullOrWhiteSpace(documentContext) Then
            userPrompt &= Environment.NewLine &
                $"<{contextLabel}>" & Environment.NewLine &
                documentContext & Environment.NewLine &
                $"</{contextLabel}>" & Environment.NewLine &
                contextSuffix
        End If

        Return Await LLM(systemPrompt, userPrompt)
    End Function

    ' ========================= Shared Bubble Insertion =========================

    ''' <summary>
    ''' Inserts a justification comment bubble on a given range. Follows the same
    ''' three-tier pattern as InsertMarkupReviewResults in MarkupReview.vb:
    '''   1. Direct range + InsertMarkdownToComment (when INI_MarkdownBubbles is enabled)
    '''   2. Direct range + plain text comment
    '''   3. AddNoticeBubbleAt fallback (last resort)
    '''
    ''' This helper is shared by JustifyMarkup and BalloonMergeWithJustification.
    ''' </summary>
    ''' <param name="doc">Target Word document.</param>
    ''' <param name="anchorRange">Range to anchor the comment to.</param>
    ''' <param name="justificationText">The justification text to insert into the bubble.</param>
    ''' <param name="prefix">Prefix appended after AN5 in the comment (e.g. " Justify Markup").</param>
    Public Sub InsertJustificationBubble(doc As Microsoft.Office.Interop.Word.Document,
                                          anchorRange As Microsoft.Office.Interop.Word.Range,
                                          justificationText As String,
                                          Optional prefix As String = "")
        If doc Is Nothing OrElse String.IsNullOrWhiteSpace(justificationText) Then Return

        Dim app As Microsoft.Office.Interop.Word.Application = doc.Application
        Dim win As Microsoft.Office.Interop.Word.Window = Nothing
        Try : win = doc.ActiveWindow : Catch : End Try

        ' If no usable anchor range, fall back to AddNoticeBubbleAt
        If anchorRange Is Nothing Then
            AddNoticeBubbleAt(doc, 0, justificationText, prefix)
            Return
        End If

        Try
            Using BeginMarkupAuthorScope(app)
                If INI_MarkdownBubbles Then
                    ' Markdown-enabled bubble (same pattern as MarkupReview lines 1253-1263)
                    Dim cmt As Microsoft.Office.Interop.Word.Comment = Nothing
                    Try
                        cmt = doc.Comments.Add(Range:=anchorRange, Text:="")
                    Catch exAdd As System.Runtime.InteropServices.COMException
                        cmt = doc.Comments.Add(Range:=anchorRange, Text:=$"{AN5}{prefix}: {justificationText}")
                    Catch exAdd2 As System.Exception
                        cmt = doc.Comments.Add(Range:=anchorRange, Text:=$"{AN5}{prefix}: {justificationText}")
                    End Try

                    Try
                        Dim cRng As Microsoft.Office.Interop.Word.Range = cmt.Range
                        cRng.Text = ""
                        If win IsNot Nothing Then
                            Dim prevShow As System.Boolean = win.View.ShowRevisionsAndComments
                            Try
                                win.View.ShowRevisionsAndComments = True
                                InsertMarkdownToComment(cRng, $"{AN5}{prefix}: {justificationText}")
                            Finally
                                win.View.ShowRevisionsAndComments = prevShow
                            End Try
                        Else
                            InsertMarkdownToComment(cRng, $"{AN5}{prefix}: {justificationText}")
                        End If
                    Catch exMk As System.Exception
                        cmt.Range.Text = $"{AN5}{prefix}: {justificationText}"
                    End Try
                Else
                    ' Plain text bubble (same pattern as MarkupReview line 1265)
                    doc.Comments.Add(Range:=anchorRange, Text:=$"{AN5}{prefix}: {justificationText}")
                End If
            End Using

        Catch ex As System.Exception
            ' Last-resort fallback (same as MarkupReview line 1268)
            AddNoticeBubbleAt(doc, 0, justificationText, prefix)
        End Try
    End Sub

    ' ========================= BalloonMerge + Justify =========================

    ''' <summary>
    ''' Extends BalloonMerge with automatic justification. Captures the original text
    ''' before merging, performs the merge via BalloonMerge, then generates and inserts
    ''' a justification bubble for the changes.
    ''' </summary>
    ''' <param name="selectWholeParagraph">If True, selects entire paragraph(s) containing comment anchor; if False, selects only anchor text.</param>
    ''' <param name="Silent">If True, uses cached prompt without user input; if False, prompts user for merge instructions.</param>
    Public Async Function BalloonMergeWithJustification(
        ByVal selectWholeParagraph As Boolean, Silent As Boolean) As System.Threading.Tasks.Task

        Dim app As Word.Application = Globals.ThisAddIn.Application
        Dim sel As Microsoft.Office.Interop.Word.Selection = app.Selection
        Dim doc As Microsoft.Office.Interop.Word.Document = sel.Range.Document

        Try
            ' Resolve the active comment and its target range (mirrors BalloonMerge logic)
            Dim activeComment As Microsoft.Office.Interop.Word.Comment = Nothing

            If sel.StoryType = WdStoryType.wdCommentsStory Then
                For Each c As Microsoft.Office.Interop.Word.Comment In doc.Comments
                    If sel.Range.Start >= c.Range.Start AndAlso
                       sel.Range.End <= c.Range.End Then
                        activeComment = c : Exit For
                    End If
                Next
            Else
                For Each c As Microsoft.Office.Interop.Word.Comment In doc.Comments
                    Dim anchor As Range = c.Scope
                    If sel.Range.End > anchor.Start AndAlso
                       sel.Range.Start < anchor.End Then
                        activeComment = c : Exit For
                    End If
                Next
            End If

            If activeComment Is Nothing Then
                ShowCustomMessageBox(
                    "This command only works when the cursor is inside a comment " &
                    "balloon or on text that has a comment.")
                Return
            End If

            ' Capture the balloon comment text BEFORE the merge (BalloonMerge may delete it)
            Dim balloonCommentText As String = String.Empty
            Try
                balloonCommentText = activeComment.Range.Text
            Catch
            End Try

            ' Determine target range (same logic as BalloonMerge)
            Dim anchorRange As Range = activeComment.Scope
            Dim targetRange As Range

            If selectWholeParagraph Then
                If anchorRange.Paragraphs.Count > 0 Then
                    Dim firstPara As Range = anchorRange.Paragraphs(1).Range
                    Dim lastPara As Range = anchorRange.Paragraphs(anchorRange.Paragraphs.Count).Range
                    targetRange = doc.Range(firstPara.Start, lastPara.End)
                Else
                    targetRange = doc.Range(anchorRange.Start, anchorRange.Start).Paragraphs(1).Range
                End If
            Else
                targetRange = anchorRange
            End If

            targetRange = targetRange.Duplicate

            ' Capture original visible text BEFORE the merge (treats tracked changes as accepted).
            ' Pass the comment range as restoreRange so the selection returns to the
            ' comment balloon afterwards — BalloonMerge expects it there.
            Dim restoreRange As Range = If(sel.StoryType = WdStoryType.wdCommentsStory,
                                           activeComment.Range, Nothing)
            Dim dummyRange As Range = Nothing
            Dim originalText As String = GetVisibleTextSafe(targetRange, restoreRange, dummyRange)

            ' Remember the start position; after merge the selection collapses to the end
            ' of the replaced text, giving us the end boundary.
            Dim premergeStart As Integer = targetRange.Start

            ' Perform the standard BalloonMerge
            Await BalloonMerge(selectWholeParagraph, Silent)

            Await EnsureUIThread()

            Try
                Dim targetWindow As Word.Window = doc.ActiveWindow
                If targetWindow IsNot Nothing Then
                    targetWindow.Activate()
                End If
            Catch
            End Try

            If app.Selection Is Nothing OrElse
               app.Selection.Range Is Nothing OrElse
               app.Selection.Range.Document Is Nothing OrElse
               Not Object.ReferenceEquals(app.Selection.Range.Document, doc) Then
                Return
            End If

            ' Reconstruct the revised range: premergeStart -> collapsed cursor position
            Dim postmergeEnd As Integer = app.Selection.Range.Start
            If postmergeEnd <= premergeStart Then
                Return
            End If

            Dim revisedRange As Range = doc.Range(premergeStart, postmergeEnd)

            ' Read revised visible text and capture the valid range for the bubble
            Dim bubbleRange As Range = Nothing
            Dim revisedText As String = GetVisibleTextSafe(revisedRange, Nothing, bubbleRange)

            ' Only justify if text actually changed
            If String.IsNullOrWhiteSpace(revisedText) OrElse
               String.Equals(originalText.Trim(), revisedText.Trim(), StringComparison.Ordinal) Then
                Return
            End If

            ' Get justification from LLM — pass the balloon comment as context (no user prompt)
            Dim justification As String = Await GetJustificationFromLLM(
                originalText, revisedText, balloonCommentText,
                "BALLOONCOMMENT",
                $"Use the balloon comment above to better justify the revision — it contains the instruction or proposal that motivated the revision; beware, it may be of internal nature, in particular if it starts with '{AN5}:'.")

            If String.IsNullOrWhiteSpace(justification) Then
                ShowCustomMessageBox("The LLM did not return a justification for the merge changes.")
                Return
            End If

            ' Insert justification bubble using the range captured by GetVisibleTextSafe
            Dim anchorForBubble As Range = If(bubbleRange, revisedRange)
            InsertJustificationBubble(doc, anchorForBubble, justification, "")

        Catch ex As System.Exception
            MessageBox.Show(
                $"Error in BalloonMergeWithJustification:{Environment.NewLine}{ex.Message}",
                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function


End Class
