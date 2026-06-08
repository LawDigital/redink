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
            MessageBox.Show("Error in ShowPaneAsync: " & ex.Message)
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
        OtherPrompt = SLib.ShowCustomInputBox("If you want, you can amend the prompt that will be used to intelligently merge your selection into your document:", $"{AN} Intelligent Merge", False, SP_MergePrompt_Cached).Trim()
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

        Dim app As Word.Application = Globals.ThisAddIn.Application
        Dim sel As Microsoft.Office.Interop.Word.Selection = app.Selection
        Dim doc As Microsoft.Office.Interop.Word.Document = sel.Range.Document

        Dim activeComment As Microsoft.Office.Interop.Word.Comment = Nothing
        Dim newtext As String = String.Empty

        Try
            ' Find the comment at current cursor position
            If sel.StoryType = WdStoryType.wdCommentsStory Then
                For Each c As Microsoft.Office.Interop.Word.Comment In doc.Comments
                    If sel.Range.Start >= c.Range.Start AndAlso
                       sel.Range.End <= c.Range.End Then
                        activeComment = c
                        Exit For
                    End If
                Next
            Else
                For Each c As Microsoft.Office.Interop.Word.Comment In doc.Comments
                    Dim anchor As Range = c.Scope
                    If sel.Range.End > anchor.Start AndAlso
                       sel.Range.Start < anchor.End Then
                        activeComment = c
                        Exit For
                    End If
                Next
            End If

            ' Validate cursor is in or on a comment
            If activeComment Is Nothing Then
                ShowCustomMessageBox(
                    "This command only works when the cursor is inside a comment " &
                    "balloon or on text that has a comment.")
                Return
            End If

            ' Extract text from comment or selection
            Dim selectedText As String = SafeRangeText(sel.Range)

            If sel.StoryType = WdStoryType.wdCommentsStory Then
                ' Inside balloon
                If selectedText.Trim().Length = 0 Then
                    newtext = SafeRangeText(activeComment.Range)
                Else
                    newtext = selectedText
                End If
            Else
                ' On anchor in main story
                If selectedText.Trim().Length = 0 Then
                    newtext = SafeRangeText(activeComment.Range)
                Else
                    newtext = selectedText
                End If
            End If

            ' Select target range in main document
            Dim anchorRange As Range = activeComment.Scope
            Dim targetRange As Range

            If selectWholeParagraph Then
                If anchorRange.Paragraphs.Count > 0 Then
                    ' Anchor spans one or more paragraphs - select them all
                    Dim firstPara As Range = anchorRange.Paragraphs(1).Range
                    Dim lastPara As Range = anchorRange.Paragraphs(anchorRange.Paragraphs.Count).Range
                    targetRange = doc.Range(firstPara.Start, lastPara.End)
                Else
                    ' Collapsed anchor (no text selected when comment was made) - select containing paragraph
                    targetRange = doc.Range(anchorRange.Start, anchorRange.Start).Paragraphs(1).Range
                End If
            Else
                ' Only the exact anchor text
                targetRange = anchorRange.Duplicate
            End If

            targetRange = targetRange.Duplicate

            Await EnsureUIThread()
            ActivateDocumentRangeContext(targetRange)

            ' Get merge prompt from user or cached value
            ActivateDocumentRangeContext(targetRange)

            If Not Silent Or String.IsNullOrWhiteSpace(SP_MergePrompt2) Then
                OtherPrompt = SLib.ShowCustomInputBox(
                    "If you want, you can amend the prompt that will be used to " &
                    "intelligently merge your comment into your document:",
                    $"{AN} Intelligent Merge", False, SP_MergePrompt2).Trim()

                If String.IsNullOrEmpty(OtherPrompt) OrElse OtherPrompt = "ESC" Then Return
            Else
                OtherPrompt = SP_MergePrompt2
            End If

            ' Prompt user for markup method
            Dim items = {
                New SelectionItem("Word", 1),
                New SelectionItem("Diff", 2),
                New SelectionItem("Diff Window", 3),
                New SelectionItem("Regex", 4),
                New SelectionItem("Diff Classic", 5),
                New SelectionItem("None", 6)
            }

            Dim DefaultItem As Integer = 6
            If INI_DoMarkupWord Then
                DefaultItem = Override(INI_MarkupMethodWord, INI_MarkupMethodWordOverride)
            End If

            Dim picked As Integer = SelectValue(items, DefaultItem, "Choose markup method ...")
            If picked < 1 Then Return

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
                If(picked < 6, True, False),
                If(picked < 6, picked, Override(INI_MarkupMethodWord, INI_MarkupMethodWordOverride)),
                False,
                False,
                True,
                False,
                INI_KeepFormatCap)

        Catch ex As System.Exception
            MessageBox.Show(
                $"Error in BalloonMerge:{Environment.NewLine}{ex.Message}",
                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
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
        OtherPrompt = SLib.ShowCustomInputBox("If you want, you can amend the prompt that will be used to intelligently merge your selection into your document:", $"{AN} Intelligent Merge", False, SP_MergePrompt_Cached).Trim()
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


    Private Sub ActivateDocumentRangeContext(ByVal targetRange As Word.Range)
        Try
            If targetRange Is Nothing Then Return

            Dim targetWindow As Word.Window = Nothing
            Try
                targetWindow = targetRange.Document.ActiveWindow
            Catch
                targetWindow = Nothing
            End Try

            If targetWindow IsNot Nothing Then
                targetWindow.Activate()
            End If

            targetRange.Select()
        Catch ex As Exception
            Debug.WriteLine($"ActivateDocumentRangeContext failed: {ex.Message}")
        End Try
    End Sub

End Class
