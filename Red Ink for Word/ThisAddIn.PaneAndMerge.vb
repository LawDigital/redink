' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.PaneAndMerge.vb
' Purpose: Provides custom task pane display and intelligent text merging
'          capabilities for Word document editing with LLM integration.
'
' Architecture:
'  - Pane Display: ShowPaneAsync displays formatted content in custom pane with
'    optional markdown conversion and document insertion.
'  - Intelligent Merge: Merges LLM-generated or selected text into document
'    with configurable formatting and markup options (Word track changes, diff,
'    diff window, regex).
'  - Comment Integration: BalloonMerge processes Word comment text and merges
'    into anchor text or paragraph with LLM processing.
'  - UI Thread Management: Uses EnsureUIThread and SwitchToUi for thread-safe
'    Word interop operations.
'  - ONNX Initialization: Lazy-loads local NER model for anonymization features.
'  - Markdown Support: Converts markdown to RTF for clipboard or inserts formatted
'    text into new/existing documents.
'  - Callback Pattern: IntelligentMergeCallback delegates pane selections to
'    merge handler.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Word
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ''' <summary>
    ''' Displays content in a custom task pane with optional markdown formatting.
    ''' Supports insertion into new or existing documents.
    ''' </summary>
    ''' <param name="introLine">Introduction text displayed at top of pane.</param>
    ''' <param name="bodyText">Main content to display; may contain markdown if insertMarkdown is True.</param>
    ''' <param name="finalRemark">Closing text displayed at bottom of pane.</param>
    ''' <param name="header">Pane header/title text.</param>
    ''' <param name="noRTF">If True, disables RTF formatting in pane display.</param>
    ''' <param name="insertMarkdown">If True, enables markdown conversion and insertion options.</param>
    Private Async Sub ShowPaneAsync(
                          introLine As String,
                          bodyText As String,
                          finalRemark As String,
                          header As String,
                          Optional noRTF As Boolean = False,
                          Optional insertMarkdown As Boolean = False
                        )
        Try
            Dim OriginalText As String = bodyText
            Dim result As String = ""

            Await EnsureUIThread()

            result = Await PaneManager.ShowMyPane(introLine, bodyText, finalRemark, header, noRTF, insertMarkdown, New IntelligentMergeCallback(AddressOf HandleIntelligentMerge))

            If result <> "" Then
                If result = "Markdown" Then
                    ' Ensure UI operations are on the main thread
                    Await SwitchToUi(
                    Sub()
                        Dim NewDocChoice As Integer = ShowCustomYesNoBox("Do you want to insert the text into a new Word document (if you cancel, it will be in the clipboard with formatting)?", "Yes, new", "No, into my existing doc")

                        If NewDocChoice = 1 Then
                            Dim newDoc As Word.Document = Globals.ThisAddIn.Application.Documents.Add()
                            Dim currentSelection As Word.Selection = newDoc.Application.Selection
                            currentSelection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                            InsertTextWithMarkdown(currentSelection, OriginalText, True, True)
                        ElseIf NewDocChoice = 2 Then
                            Dim currentSelection As Word.Selection = Globals.ThisAddIn.Application.Selection
                            currentSelection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                            Globals.ThisAddIn.Application.Selection.TypeParagraph()
                            Globals.ThisAddIn.Application.Selection.TypeParagraph()
                            InsertTextWithMarkdown(currentSelection, OriginalText, False)
                        Else
                            ShowCustomMessageBox("No text was inserted (but included in the clipboard as RTF).")
                            SLib.PutInClipboard(MarkdownToRtfConverter.Convert((OriginalText)))
                        End If
                    End Sub
                )
                End If
            End If

        Catch ex As System.Exception
            Debug.WriteLine("Bodytext=" & bodyText)
            Global.SharedLibrary.SharedLibrary.SharedMethods.ShowCustomMessageBox("Error in ShowPaneAsync: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Callback handler for pane text selection events. Delegates to IntelligentMerge.
    ''' </summary>
    ''' <param name="selectedText">Text selected by user in custom pane.</param>
    Private Sub HandleIntelligentMerge(selectedText As String)
        IntelligentMerge(selectedText)
    End Sub

    ''' <summary>
    ''' Intelligently merges new text into selected document text using LLM processing.
    ''' Prompts user for merge instructions and applies configured formatting/markup.
    ''' </summary>
    ''' <param name="newtext">Text to merge into document (typically from pane selection).</param>
    Public Async Sub IntelligentMerge(newtext As String)
        Dim application As Word.Application = Globals.ThisAddIn.Application
        Dim selection As Microsoft.Office.Interop.Word.Selection = application.Selection
        If selection.Type = WdSelectionType.wdSelectionIP Then
            ShowCustomMessageBox("Please select the text in your document with which your selection in the pane shall be merged.")
            Return
        End If
        OtherPrompt = SLib.ShowCustomInputBox("If you want, you can amend the prompt that will be used to intelligently merge your selection into your document:", $"{AN} Intelligent Merge", False, SP_MergePrompt_Cached, Context:=_context).Trim()
        If String.IsNullOrEmpty(OtherPrompt) Or OtherPrompt = "ESC" Then Return
        Dim result As String = Await ProcessSelectedText(OtherPrompt & " " & SP_Add_MergePrompt & " <INSERT>" & newtext & "</INSERT> ", True, INI_KeepFormat2, INI_KeepParaFormatInline, Override(INI_ReplaceText2, INI_ReplaceText2Override), INI_DoMarkupWord, Override(INI_MarkupMethodWord, INI_MarkupMethodWordOverride), False, False, True, False, INI_KeepFormatCap)
    End Sub

    ''' <summary>
    ''' Merges Word comment text into document text using LLM processing.
    ''' Selects comment anchor text or entire paragraph(s) based on parameters.
    ''' </summary>
    ''' <param name="selectWholeParagraph">If True, selects entire paragraph(s) containing comment anchor; if False, selects only anchor text.</param>
    ''' <param name="Silent">If True, uses cached prompt without user input; if False, prompts user for merge instructions.</param>
    Public Async Function BalloonMerge(
        ByVal selectWholeParagraph As Boolean, Silent As Boolean) As System.Threading.Tasks.Task

        Try
            Await EnsureUIThread()
            Dim context As BalloonMergeContext = Await CaptureBalloonMergeContextAsync(Globals.ThisAddIn.Application.Selection, selectWholeParagraph)
            Await EnsureUIThread()
            If context Is Nothing Then
                ShowCustomMessageBox("Select a comment, or place the cursor in the text that the comment refers to.", "Inky - Apply Comment")
                Return
            End If
            If Not Await ApplyBalloonMergeContextAsync(context, Silent) Then Return
            Await EnsureUIThread()
        Catch ex As System.Exception
            ReportWordCommandFailure("ApplyComment", ex, "Inky could not apply this comment. Please select the comment again and retry.")
        End Try
    End Function

    ''' <summary>One captured comment, instruction source and target, shared by merge and justification.</summary>
    Private NotInheritable Class BalloonMergeContext
        Public Property Comment As Microsoft.Office.Interop.Word.Comment
        Public Property TargetRange As Microsoft.Office.Interop.Word.Range
        Public Property OriginalSelection As Microsoft.Office.Interop.Word.Range
        Public Property SelectionWasInComment As System.Boolean
        Public Property FullCommentText As System.String
        Public Property SelectedCommentText As System.String
    End Class

    Private Async Function CaptureBalloonMergeContextAsync(selection As Microsoft.Office.Interop.Word.Selection,
                                                selectWholeParagraph As System.Boolean) As System.Threading.Tasks.Task(Of BalloonMergeContext)
        Await EnsureUIThread()
        If selection Is Nothing Then Return Nothing

        ' Capture the modern UI before reading Comment.Range (which can move Word
        ' out of its modern comments view). This await never blocks Office's STA.
        Dim evidence As WordCommentSelectionTracker.Evidence = Await CaptureCommentSelectionEvidenceAsync()
        Await EnsureUIThread()

        ' Word.Selection is live. Capture origin, coordinates and text BEFORE inspecting comments
        ' or activating ranges; later UI/COM state must not become the instruction source.
        Dim origin As Microsoft.Office.Interop.Word.Range = selection.Range.Duplicate
        Dim originStory As Microsoft.Office.Interop.Word.WdStoryType = origin.StoryType
        Dim originStart As System.Int32 = origin.Start
        Dim originEnd As System.Int32 = origin.End
        Dim originText As System.String = SafeRangeText(origin)
        Dim doc As Microsoft.Office.Interop.Word.Document = origin.Document
        Dim activeComment As Microsoft.Office.Interop.Word.Comment = Nothing

        ' A selected comment card is an explicit command target even when Word still
        ' exposes a document caret/selection through Range. Resolve that target before
        ' looking for overlapping document anchors; never guess from a visible pane.
        Dim selectedCommentRange As Microsoft.Office.Interop.Word.Range = Nothing
        Dim commentSelectionDetected As System.Boolean = False
        If Not TryResolveSelectedComment(selection, origin, activeComment,
                                         selectedCommentRange, commentSelectionDetected, evidence) Then
            If commentSelectionDetected Then
                ' A card/comment-story selection was attempted but its identity
                ' was not established. The old body selection is NOT its target.
                If doc.Comments.Count = 0 Then Return Nothing
                activeComment = ChooseCommentForCommand(doc, "Apply Comment")
            Else
                ' Main document text selection/caret takes precedence over a
                ' fallback picker. A partial overlap is sufficient; a large
                ' selection with several comments is genuinely ambiguous here.
                Dim anchored As System.Collections.Generic.List(Of Microsoft.Office.Interop.Word.Comment) =
                    FindCommandCommentsInBodyRange(doc, origin)
                If anchored.Count = 1 Then
                    activeComment = anchored(0)
                ElseIf anchored.Count > 1 Then
                    activeComment = ChooseCommentForCommand(doc, "Apply Comment", anchored)
                End If
            End If
        End If
        If activeComment Is Nothing Then Return Nothing

        Dim fullCommentText As System.String = SafeRangeText(activeComment.Range.Duplicate)
        If System.String.IsNullOrWhiteSpace(fullCommentText) Then
            Throw New WordCommandNoticeException("This comment is empty or its text is not available. Please finish editing the comment and try again.")
        End If

        Dim anchorRange As Microsoft.Office.Interop.Word.Range = activeComment.Scope.Duplicate
        If Not IsSameProcessingDocument(doc, anchorRange.Document) OrElse
           anchorRange.StoryType = Microsoft.Office.Interop.Word.WdStoryType.wdCommentsStory Then
            Throw New WordCommandNoticeException("The text linked to this comment is no longer available. Please select another comment.")
        End If
        Dim targetRange As Microsoft.Office.Interop.Word.Range
        If selectWholeParagraph Then
            If anchorRange.Paragraphs.Count > 0 Then
                Dim firstPara As Microsoft.Office.Interop.Word.Range = anchorRange.Paragraphs(1).Range
                Dim lastPara As Microsoft.Office.Interop.Word.Range = anchorRange.Paragraphs(anchorRange.Paragraphs.Count).Range
                targetRange = anchorRange.Duplicate
                targetRange.SetRange(firstPara.Start, lastPara.End)
            Else
                targetRange = anchorRange.Duplicate
                targetRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseStart)
                targetRange = targetRange.Paragraphs(1).Range
            End If
        Else
            targetRange = anchorRange.Duplicate
        End If

        Dim inComment As System.Boolean = selectedCommentRange IsNot Nothing
        Dim selectedCommentText As System.String = System.String.Empty
        If inComment Then
            ' Only a verified range inside THIS comment may provide a fragment. In
            ' particular, a residual main-text selection is never an instruction.
            selectedCommentText = If(selectedCommentRange.IsEqual(origin), originText, SafeRangeText(selectedCommentRange))
        End If
        Dim modernFragment As System.Boolean = False
        If Not inComment AndAlso evidence IsNot Nothing AndAlso
           WordCommentSelectionTracker.Matches(evidence, fullCommentText, activeComment.Author) AndAlso
           Not System.String.IsNullOrWhiteSpace(evidence.SelectedText) Then
            Dim normalizedFragment As System.String = WordCommentSelectionTracker.NormalizeText(evidence.SelectedText)
            Dim normalizedComment As System.String = WordCommentSelectionTracker.NormalizeText(fullCommentText)
            If normalizedFragment.Length > 0 AndAlso normalizedComment.IndexOf(normalizedFragment, System.StringComparison.Ordinal) >= 0 Then
                selectedCommentText = evidence.SelectedText
                modernFragment = True
            End If
        End If
        ' A modern UI text selection is not a Word.Range. Never restore the
        ' unrelated body selection as though it were a comment range (Justify).
        Return New BalloonMergeContext With {
            .Comment = activeComment,
            .TargetRange = targetRange.Duplicate,
            .OriginalSelection = If(modernFragment, Nothing, If(selectedCommentRange, origin)),
            .SelectionWasInComment = inComment OrElse modernFragment,
            .FullCommentText = fullCommentText,
            .SelectedCommentText = selectedCommentText
        }
    End Function

    ''' <summary>
    ''' Resolves only an explicit comment selection for Apply Comment and Freestyle Reply.
    ''' Does not change the selection, inspect inactive panes, or alter the general
    ''' TryResolveActiveComment policy used by the comment-text editing pipeline.
    ''' </summary>
    Private Function TryResolveSelectedComment(
        selection As Microsoft.Office.Interop.Word.Selection,
        origin As Microsoft.Office.Interop.Word.Range,
        ByRef activeComment As Microsoft.Office.Interop.Word.Comment,
        ByRef selectedCommentRange As Microsoft.Office.Interop.Word.Range,
        ByRef commentSelectionDetected As System.Boolean,
        Optional evidence As WordCommentSelectionTracker.Evidence = Nothing) As System.Boolean

        activeComment = Nothing
        selectedCommentRange = Nothing
        commentSelectionDetected = False
        If selection Is Nothing OrElse origin Is Nothing Then Return False
        Dim doc As Microsoft.Office.Interop.Word.Document = origin.Document

        ' An explicit modern-card focus is stronger than Selection.Comments, which
        ' may describe the old BODY selection. Never consult that collection to
        ' disambiguate a card. Resolve exact UI evidence against this document only.
        If evidence IsNot Nothing AndAlso evidence.DocumentGeneration >= 0 AndAlso
           _commentSelectionTracker IsNot Nothing AndAlso Not _commentSelectionTracker.IsCurrent(evidence) Then
            commentSelectionDetected = True
            LogCommentCommandResolution("focus_generation_changed; explicit_choice_required")
            Return False
        End If
        If evidence IsNot Nothing AndAlso evidence.IsComment Then
            commentSelectionDetected = True
            If evidence.WindowHandle <> New System.IntPtr(selection.Application.ActiveWindow.Hwnd) Then Return False
            If TryResolveCommentEvidence(doc, evidence, activeComment) Then
                If origin.StoryType = Microsoft.Office.Interop.Word.WdStoryType.wdCommentsStory AndAlso
                   origin.InRange(activeComment.Range) Then selectedCommentRange = origin.Duplicate
                Return True
            End If
            ' UIA can be unavailable for classic balloons. Only a genuine comment-
            ' story range may recover it; never a residual main-document selection.
            If evidence.Texts.Count > 0 OrElse
               origin.StoryType <> Microsoft.Office.Interop.Word.WdStoryType.wdCommentsStory Then Return False
        End If

        Dim diagnostic As System.String = System.String.Empty
        Dim detected As System.Boolean = False
        If TryResolveSelectedCommentInSelection(selection, origin, doc, activeComment,
                                                selectedCommentRange, detected, diagnostic) Then
            commentSelectionDetected = True
            LogCommentCommandResolution("classic_selection; " & diagnostic)
            Return True
        End If
        commentSelectionDetected = commentSelectionDetected OrElse detected
        LogCommentCommandResolution("classic_selection; " & diagnostic)
        Try
            Dim paneSelection As Microsoft.Office.Interop.Word.Selection = selection.Application.ActiveWindow.ActivePane.Selection
            If paneSelection IsNot Nothing AndAlso paneSelection.Active AndAlso
               IsSameProcessingDocument(doc, paneSelection.Document) Then
                Dim paneOrigin As Microsoft.Office.Interop.Word.Range = paneSelection.Range.Duplicate
                Dim paneDetected As System.Boolean = False
                Dim resolved As System.Boolean = TryResolveSelectedCommentInSelection(
                    paneSelection, paneOrigin, doc, activeComment, selectedCommentRange, paneDetected, diagnostic)
                commentSelectionDetected = commentSelectionDetected OrElse paneDetected
                LogCommentCommandResolution("classic_active_pane; " & diagnostic)
                If resolved Then Return True
            End If
        Catch ex As System.Exception
            LogCommentCommandResolution("classic_active_pane_unavailable; " & ex.ToString())
        End Try
        Return False
    End Function

    Private Function TryResolveSelectedCommentInSelection(
        selection As Microsoft.Office.Interop.Word.Selection,
        origin As Microsoft.Office.Interop.Word.Range,
        doc As Microsoft.Office.Interop.Word.Document,
        ByRef activeComment As Microsoft.Office.Interop.Word.Comment,
        ByRef selectedCommentRange As Microsoft.Office.Interop.Word.Range,
        ByRef commentSelectionDetected As System.Boolean,
        ByRef diagnostic As System.String) As System.Boolean

        activeComment = Nothing
        selectedCommentRange = Nothing
        commentSelectionDetected = False
        diagnostic = "no_comment_selection"
        Dim selectionStory As System.Int32 = -1
        Dim reportedCount As System.Int32 = -1
        Try
            selectionStory = CInt(selection.StoryType)
            commentSelectionDetected = selectionStory = CInt(Microsoft.Office.Interop.Word.WdStoryType.wdCommentsStory)
        Catch ex As System.Exception
            System.Diagnostics.Debug.WriteLine("Comment selection story unavailable: " & ex.Message)
        End Try
        Try
            commentSelectionDetected = commentSelectionDetected OrElse
                CBool(selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdInCommentPane))
        Catch ex As System.Exception
            System.Diagnostics.Debug.WriteLine("Comment pane information unavailable: " & ex.Message)
        End Try

        Try
            commentSelectionDetected = commentSelectionDetected OrElse
                origin.StoryType = Microsoft.Office.Interop.Word.WdStoryType.wdCommentsStory
            diagnostic = "selectionStory=" & selectionStory.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                "; rangeStory=" & CInt(origin.StoryType).ToString(System.Globalization.CultureInfo.InvariantCulture) &
                "; commentSelection=" & commentSelectionDetected.ToString()
            If Not commentSelectionDetected Then Return False

            Dim candidates As Microsoft.Office.Interop.Word.Comments = Nothing
            Try
                candidates = selection.Comments
                reportedCount = candidates.Count
            Catch ex As System.Exception
                System.Diagnostics.Debug.WriteLine("Selected comments collection unavailable: " & ex.Message)
                candidates = Nothing
            End Try
            diagnostic &= "; reportedComments=" & reportedCount.ToString(System.Globalization.CultureInfo.InvariantCulture)

            ' A sole Selection.Comments item in the MAIN story is still the
            ' comment on the old text selection; it is NOT evidence of the card.
            If origin.StoryType <> Microsoft.Office.Interop.Word.WdStoryType.wdCommentsStory Then
                diagnostic &= "; card_requires_uia_or_explicit_choice"
                Return False
            End If
            Dim matched As Microsoft.Office.Interop.Word.Comment = Nothing
            For Each entry As CommentChoice In ReadCommandCommentCatalog(doc)
                If Not AddClassicCommentRangeMatch(doc, origin, entry.Comment, matched) Then
                    diagnostic &= "; ambiguous_comment_range"
                    Return False
                End If
            Next
            If matched Is Nothing Then
                diagnostic &= "; no_matching_comment"
                Return False
            End If

            Dim matchedRange As Microsoft.Office.Interop.Word.Range = matched.Range
            If Not IsSameProcessingDocument(doc, matchedRange.Document) OrElse
               matchedRange.StoryType <> Microsoft.Office.Interop.Word.WdStoryType.wdCommentsStory Then
                diagnostic &= "; invalid_comment_document_or_story"
                Return False
            End If
            If origin.StoryType = Microsoft.Office.Interop.Word.WdStoryType.wdCommentsStory AndAlso origin.InRange(matchedRange) Then
                selectedCommentRange = origin.Duplicate
            End If
            activeComment = matched
            diagnostic &= "; resolved"
            Try
                diagnostic &= "; commentIndex=" & matched.Index.ToString(System.Globalization.CultureInfo.InvariantCulture)
            Catch ex As System.Exception
                System.Diagnostics.Debug.WriteLine("Comment index unavailable for diagnostics: " & ex.Message)
            End Try
            Return True
        Catch ex As System.Exception
            activeComment = Nothing
            selectedCommentRange = Nothing
            diagnostic &= "; unavailable; " & ex.Message
            Return False
        End Try
    End Function

    Private Shared Function AddClassicCommentRangeMatch(doc As Microsoft.Office.Interop.Word.Document,
                                                        origin As Microsoft.Office.Interop.Word.Range,
                                                        candidate As Microsoft.Office.Interop.Word.Comment,
                                                        ByRef matched As Microsoft.Office.Interop.Word.Comment) As System.Boolean
        Dim candidateRange As Microsoft.Office.Interop.Word.Range = candidate.Range
        If IsSameProcessingDocument(doc, candidateRange.Document) AndAlso origin.InRange(candidateRange) Then
            If matched IsNot Nothing AndAlso Not SameCommandComment(matched, candidate) Then Return False
            matched = candidate
        End If
        Return True
    End Function

    Private Sub LogCommentCommandResolution(message As System.String)
        LogWordCommandDiagnostic("CommentSelection", message)
    End Sub

    Private Async Function ApplyBalloonMergeContextAsync(context As BalloonMergeContext,
                                                        Silent As System.Boolean,
                                                        Optional onDocumentWriteback As System.Action(Of Microsoft.Office.Interop.Word.Range) = Nothing) As System.Threading.Tasks.Task(Of System.Boolean)
        Await EnsureUIThread()
        Dim targetRange As Microsoft.Office.Interop.Word.Range = context.TargetRange.Duplicate
        Dim newtext As System.String = context.FullCommentText
        Dim instructionSource As System.String = "full_comment"

        ' A partial comment-story selection can be deliberate, or Word's residual selection
        ' (for example just a leading label). Make the choice explicit; never silently truncate.
        If context.SelectionWasInComment AndAlso Not System.String.IsNullOrWhiteSpace(context.SelectedCommentText) Then
            If System.String.Equals(context.SelectedCommentText.Trim(), context.FullCommentText.Trim(), System.StringComparison.Ordinal) Then
                newtext = context.SelectedCommentText
                instructionSource = "selected_full_comment"
            Else
                Dim sourceItems As New System.Collections.Generic.List(Of SharedLibrary.SharedLibrary.SharedMethods.SelectionItem) From {
                    New SharedLibrary.SharedLibrary.SharedMethods.SelectionItem("Use the entire comment", 1),
                    New SharedLibrary.SharedLibrary.SharedMethods.SelectionItem("Use only the selected comment text", 2)
                }
                Dim preview As System.String = context.SelectedCommentText.Replace(vbCr, " ").Replace(vbLf, " ")
                If preview.Length > 160 Then preview = preview.Substring(0, 160) & "..."
                Dim sourceChoice As System.Int32 = SelectValue(sourceItems, 1,
                    "You selected part of the comment: " & preview & System.Environment.NewLine &
                    "Choose the instruction to apply. The entire comment is selected by default.")
                If sourceChoice < 1 Then Return False
                If sourceChoice = 2 Then
                    newtext = context.SelectedCommentText
                    instructionSource = "selected_fragment"
                End If
            End If
        End If

        If _context IsNot Nothing AndAlso _context.INI_APIDebug Then
            SharedLibrary.SharedLogger.Log(_context, _context.RDV,
                "BalloonMerge input source=" & instructionSource &
                "; fullCommentChars=" & context.FullCommentText.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                "; selectedCommentChars=" & context.SelectedCommentText.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                "; sentChars=" & newtext.Length.ToString(System.Globalization.CultureInfo.InvariantCulture))
        End If

        Await EnsureUIThread()
        ActivateDocumentRangeContext(targetRange)

        ' Get merge prompt from user or cached value

        If Not Silent Or String.IsNullOrWhiteSpace(SP_MergePrompt2) Then
            OtherPrompt = SLib.ShowCustomInputBox(
                "If you want, you can amend the prompt that will be used to " &
                "intelligently merge your comment into your document:",
                $"{AN} Intelligent Merge", False, SP_MergePrompt2, Context:=_context).Trim()

            If String.IsNullOrEmpty(OtherPrompt) OrElse OtherPrompt = "ESC" Then Return False
        Else
            OtherPrompt = SP_MergePrompt2
        End If

        ' Prompt user for markup method
        Dim items = {
            New SelectionItem("Word", 1),
            New SelectionItem("Diff", 2),
            New SelectionItem("Diff Window", 3),
            New SelectionItem("Regex", 4),
            New SelectionItem("Diff (words only)", 5),
            New SelectionItem("Diff Classic", 6),
            New SelectionItem("None", 7)
        }

        Dim DefaultItem As Integer = 7
        If INI_DoMarkupWord Then
            DefaultItem = Override(INI_MarkupMethodWord, INI_MarkupMethodWordOverride)
        End If

        Dim picked As Integer = SelectValue(items, DefaultItem, "Choose markup method ...")
        If picked < 1 Then Return False

        ' Process merge with LLM against the main document selection.
        ' Call TrueProcessSelectedText directly to avoid the comment-bubble reroute in ProcessSelectedText.
        ActivateDocumentRangeContext(targetRange)

        Dim result As String = Await TrueProcessSelectedText(
            OtherPrompt & " " & SP_Add_MergePrompt & " <INSERT>" &
            newtext & "</INSERT> ",
            True,
            INI_KeepFormat2,
            INI_KeepParaFormatInline,
            Override(INI_ReplaceText2, INI_ReplaceText2Override),
            If(picked < 7, True, False),
            If(picked < 7, picked, Override(INI_MarkupMethodWord, INI_MarkupMethodWordOverride)),
            False,
            False,
            True,
            False,
            INI_KeepFormatCap,
            onDocumentWriteback:=onDocumentWriteback)
        Await EnsureUIThread()

        Return True
    End Function

    ''' <summary>
    ''' Safely retrieves text from a Word range, returning empty string if range is Nothing or text is inaccessible.
    ''' </summary>
    ''' <param name="r">Word range to extract text from.</param>
    ''' <returns>Range text or empty string on error.</returns>
    Private Function SafeRangeText(r As Word.Range) As String
        If r Is Nothing Then Return String.Empty
        Try
            Dim t As String = r.Text
            If t Is Nothing Then t = String.Empty
            Return t
        Catch
            ' Range.Text can throw in corrupt documents
            Return String.Empty
        End Try
    End Function

    ''' <summary>
    ''' Merges comment balloon text into selected document text using LLM processing.
    ''' Requires active text selection in document.
    ''' </summary>
    ''' <param name="newtext">Comment text to merge into document selection.</param>
    Public Async Sub IntelligentMergeBalloon(newtext As String)
        Dim application As Word.Application = Globals.ThisAddIn.Application
        Dim selection As Microsoft.Office.Interop.Word.Selection = application.Selection
        If selection.Type = WdSelectionType.wdSelectionIP Then
            ShowCustomMessageBox("Please select the text in your document with which your selection in the pane shall be merged.")
            Return
        End If
        OtherPrompt = SLib.ShowCustomInputBox("If you want, you can amend the prompt that will be used to intelligently merge your selection into your document:", $"{AN} Intelligent Merge", False, SP_MergePrompt_Cached, Context:=_context).Trim()
        If String.IsNullOrEmpty(OtherPrompt) Or OtherPrompt = "ESC" Then Return
        Dim result As String = Await ProcessSelectedText(OtherPrompt & " " & SP_Add_MergePrompt & " <INSERT>" & newtext & "</INSERT> ", True, INI_KeepFormat2, INI_KeepParaFormatInline, Override(INI_ReplaceText2, INI_ReplaceText2Override), INI_DoMarkupWord, Override(INI_MarkupMethodWord, INI_MarkupMethodWordOverride), False, False, True, False, INI_KeepFormatCap)
    End Sub

    ''' <summary>
    ''' Tracks whether the ONNX NER model has been successfully initialized.
    ''' </summary>
    Public ONNX_initialized As Boolean = False

    ''' <summary>
    ''' Lazily initializes the ONNX-based Named Entity Recognition model for anonymization.
    ''' Loads model, vocabulary, and label files from configured local path.
    ''' </summary>
    ''' <returns>True if already initialized or initialization succeeds; False on error.</returns>
    Private Function EnsureInitialized() As Boolean

        If Not ONNX_initialized AndAlso Not String.IsNullOrEmpty(Globals.ThisAddIn.INI_LocalModelPath) Then
            Try
                Dim modelpath As String = System.IO.Path.Combine(ExpandEnvironmentVariables(Globals.ThisAddIn.INI_LocalModelPath), NER_Model)
                Dim vocabpath As String = System.IO.Path.Combine(ExpandEnvironmentVariables(Globals.ThisAddIn.INI_LocalModelPath), NER_Token)
                Dim labelpath As String = System.IO.Path.Combine(ExpandEnvironmentVariables(Globals.ThisAddIn.INI_LocalModelPath), NER_Label)

                OnnxAnonymizer.Initialize(modelpath, vocabpath, labelpath, 128)
                ONNX_initialized = True
                Return True
            Catch ex As Exception
                SLib.ShowCustomMessageBox($"Error loading And initializing the NER model ({ex.Message}).")
                ONNX_initialized = False
                Return False
            End Try
        Else
            Return ONNX_initialized
        End If
    End Function


    Private Sub ActivateDocumentRangeContext(ByVal targetRange As Microsoft.Office.Interop.Word.Range)
        Try
            If targetRange Is Nothing Then
                Throw New System.InvalidOperationException("The comment's document range is unavailable.")
            End If

            Dim targetWindow As Microsoft.Office.Interop.Word.Window = targetRange.Document.ActiveWindow
            If targetWindow IsNot Nothing Then targetWindow.Activate()
            targetRange.Select()

            ' Selecting a comment card can leave a different body range exposed by
            ' Word. Do not start the model/writeback path unless the intended document,
            ' story and exact range have actually become the active selection.
            Dim selected As Microsoft.Office.Interop.Word.Selection = Globals.ThisAddIn.Application.Selection
            If selected Is Nothing OrElse Not IsSameProcessingDocument(targetRange.Document, selected.Document) OrElse
               selected.StoryType <> targetRange.StoryType OrElse Not selected.Range.IsEqual(targetRange) Then
                Throw New System.InvalidOperationException("Word did not activate the comment's document range.")
            End If
        Catch ex As System.Exception
            System.Diagnostics.Debug.WriteLine($"ActivateDocumentRangeContext failed: {ex.Message}")
            LogWordCommandDiagnostic("ApplyComment.ActivateTarget", ex.ToString())
            Throw New WordCommandNoticeException(
                "Inky could not select the text linked to this comment. Nothing has been applied. Please close any open dialog in Word and try again.")
        End Try
    End Sub

End Class
