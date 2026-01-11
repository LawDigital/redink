' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.Processing.vb
' Purpose: Drives the Red Ink for Word text-processing pipeline, including selection
'          validation, chunk iteration, formatting capture/restoration, LLM invocation,
'          and routing AI output into Word, panes, clipboard, bubbles, podcasts, or slides.
'
' Architecture:
'  - Entry Points & Undo: `ProcessSelectedText` enforces prerequisites, opens `WordUndoScope`,
'    handles table-aware branching, and delegates to `TrueProcessSelectedText`.
'  - Story & Chunk Management: re-centers the selection in the main story, bookmarks the
'    working range, and iterates paragraph chunks while snapshotting styles via
'    `ParagraphFormatStructure`.
'  - Formatting & Markdown/HTML: extracts HTML/inline text, performs Markdown conversion,
'    restores special fields, and reapplies paragraph/list formatting after insertions.
'  - LLM Workflow: builds prompts (incl. bubbles, TP markup, slides, MyStyle), enforces token
'    caps, supports alternate models, and post-processes responses (corrections, trimming).
'  - Markup Generation: offers Word compare, DiffPlex inline tagging, regex-based markups,
'    and revision-tag helpers (`AddMarkupTags`, `InsertMarkupText`, `SearchAndReplace`).
'  - Output & UI Surfaces: dispatches results to clipboard panes, new docs, in-place replacements,
'    podcast scripts, slide decks, bubbles/pushbacks, with progress/cancellation handling.
'  - Story Routing: Detects non-main stories (comments, footnotes/endnotes) and routes processing to
'    specialized handlers (`ProcessSelectedTextInActiveCommentBubble`, `ProcessSelectedTextInActiveFootnote`).
'
' External Dependencies:
'  - Microsoft.Office.Interop.Word for document automation and story navigation.
'  - SharedLibrary.SharedMethods for UI, prompt construction, LLM access, formatting utilities.
'  - DiffPlex for diff building, Markdig + HtmlAgilityPack for Markdown-to-HTML conversion.
'  - DocumentFormat.OpenXml, MarkdownToRtfConverter, and additional helpers referenced via SharedLibrary.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
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

    ''' <summary>
    ''' P/Invoke to enable or disable a window handle.
    ''' </summary>
    ''' <param name="hWnd">Window handle.</param>
    ''' <param name="bEnable">True to enable, False to disable.</param>
    ''' <returns>True if the window was previously enabled.</returns>
    <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function EnableWindow(hWnd As System.IntPtr, bEnable As System.Boolean) As System.Boolean
    End Function

    ''' <summary>
    ''' P/Invoke to check whether a window handle is enabled.
    ''' </summary>
    ''' <param name="hWnd">Window handle.</param>
    ''' <returns>True if enabled, False otherwise.</returns>
    <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function IsWindowEnabled(hWnd As System.IntPtr) As System.Boolean
    End Function

    ' ProcessSelectedText Parameters:
    ' - SysCommand: A string command to be processed.
    ' - CheckMaxToken: Boolean flag to check the maximum token limit.
    ' - KeepFormat: Boolean flag to maintain the formatting of the text.
    ' - ParaFormatInline: Boolean flag to format paragraphs inline.
    ' - InPlace: Boolean flag to indicate that the output should replace the selected text.
    ' - DoMarkup: Boolean flag to indicate that the output should be provided as a markup of the selected text.
    ' - MarkupMethod: Integer to indicate the markup method to be used: 1 = Word, 2 = Diff, 3 = Regex
    ' - PutInClipboard: Boolean flag to output the processed text in the clipboard.
    ' - PutInBubbles: Boolean flag to output the processed text in bubbles
    ' - SelectionMandatory: Boolean flag to enforce text selection before processing.
    ' - UseSecondAPI: Boolean flag to decide if a secondary API should be utilized.
    ' - FormattingCap: Number indicating the maximum number of characters for preserving format
    ' - DoTPMarkup: Boolean flag to indicate that markups in the output should marked.
    ' - TPMarkupname: String containing the user of whom the tags will be marked, if any.
    ' - CreatePodcast: Boolean flag to indicate that the output should be used to create a podcast.
    ' - FileObject: String containing the file path to the object to be added to the LLM request if supported by the API.
    ' - DoPane: Boolean flag to indicate that the output should be shown in a pane.
    ' - WithRevisions: Boolean flag to indicate that the input should contain Word revisions.
    ' - ChunkSize: Integer indicating how many paragraphs should be processed at once.
    ' - NoFormatAndFieldSaving: Boolean flag to indicate that no standard formatting/field saving should be applied to the selected text.
    ' - DoNewDoc: Boolean flag to indicate that the output should be placed in a new document.
    ' - SlideDeck: String containing the file path to the PowerPoint deck to be created or modified.
    ' - AddDocs: Boolean flag indicating whether insertdocs should be added
    ' - DoMyStyle: Boolean flag whether MyStyleInsert shall be used
    ' - DoBubblesExtract: Boolean flag whether text shall be extracted from bubbles
    ' - DoPushback: Boolean flag whether a pushback to bubbles shall be done
    ' - SelectedTools: List of ModelConfig indicating the tools selected for tooling runs

    ''' <summary>
    ''' Stores paragraph-level style, font, list, alignment, and spacing information for reapplication after text processing.
    ''' </summary>
    Structure ParagraphFormatStructure
        Dim Style As Word.Style
        Dim FontName As String
        Dim FontSize As Nullable(Of Integer)
        Dim FontBold As Nullable(Of Integer)  ' 1=True, 0=False, Nothing=keep
        Dim FontItalic As Nullable(Of Integer)
        Dim FontUnderline As Nullable(Of Word.WdUnderline)
        Dim FontColor As Nullable(Of Long)

        Dim ListType As Word.WdListType
        Dim ListTemplate As Word.ListTemplate
        Dim ListLevel As Integer
        Dim ListNumber As Integer
        Dim HasListFormat As Boolean

        Dim Alignment As Word.WdParagraphAlignment

        ' Line spacing + rules
        Dim LineSpacingRule As Word.WdLineSpacing
        Dim LineSpacing As Single

        ' Spacing before/after (values and auto flags)
        Dim SpaceBeforeAuto As Boolean
        Dim SpaceAfterAuto As Boolean
        Dim SpaceBefore As Single
        Dim SpaceAfter As Single

        ' Additional flags affecting spacing behavior
        Dim NoSpaceBetweenParagraphsOfSameStyle As Boolean
        Dim DisableLineHeightGrid As Boolean
    End Structure

    ''' <summary>
    ''' Module-level array holding paragraph format snapshots for the current selection.
    ''' </summary>
    Dim paragraphFormat() As ParagraphFormatStructure

    ''' <summary>
    ''' Count of paragraphs in the current selection snapshot.
    ''' </summary>
    Dim paraCount As Integer

    ''' <summary>
    ''' Entry point for processing selected text. Validates prerequisites (system prompt, selection), opens an undo scope,
    ''' handles table-aware processing (cell-by-cell or whole table), and delegates to TrueProcessSelectedText.
    ''' </summary>
    ''' <param name="SysCommand">System prompt for the LLM.</param>
    ''' <param name="CheckMaxToken">Whether to check token limits.</param>
    ''' <param name="KeepFormat">Preserve HTML formatting.</param>
    ''' <param name="ParaFormatInline">Include inline paragraph format markers.</param>
    ''' <param name="InPlace">Replace selection in place.</param>
    ''' <param name="DoMarkup">Generate markup comparing original and new text.</param>
    ''' <param name="MarkupMethod">1=Word compare, 2=Diff, 3=DiffWindow, 4=Regex.</param>
    ''' <param name="PutInClipboard">Place result in clipboard.</param>
    ''' <param name="PutInBubbles">Place result as Word comments (bubbles).</param>
    ''' <param name="SelectionMandatory">Require text selection before processing.</param>
    ''' <param name="UseSecondAPI">Use secondary API configuration.</param>
    ''' <param name="FormattingCap">Character limit for format preservation.</param>
    ''' <param name="DoTPMarkup">Mark revisions for a specific author.</param>
    ''' <param name="TPMarkupname">Author name for revision filtering.</param>
    ''' <param name="CreatePodcast">Generate podcast script output.</param>
    ''' <param name="FileObject">File path for additional LLM input.</param>
    ''' <param name="DoPane">Show result in a pane.</param>
    ''' <param name="ChunkSize">Paragraph chunk size for iteration.</param>
    ''' <param name="NoFormatAndFieldSaving">Skip format/field saving.</param>
    ''' <param name="DoNewDoc">Place result in a new document.</param>
    ''' <param name="SlideDeck">PowerPoint deck path for slide generation.</param>
    ''' <param name="AddDocs">Include insertdocs content.</param>
    ''' <param name="DoMyStyle">Use MyStyle insertion.</param>
    ''' <param name="DoBubblesExtract">Extract text from bubbles.</param>
    ''' <param name="DoPushback">Reply to bubbles.</param>
    ''' <param name="SelectedTools">Tools selected for Tooling</param>
    ''' <returns>Empty string on completion.</returns>
    Public Async Function ProcessSelectedText(
     SysCommand As String,
     CheckMaxToken As Boolean,
     KeepFormat As Boolean,
     ParaFormatInline As Boolean,
     InPlace As Boolean,
     DoMarkup As Boolean,
     MarkupMethod As Integer,
     PutInClipboard As Boolean,
     PutInBubbles As Boolean,
     SelectionMandatory As Boolean,
     UseSecondAPI As Boolean,
     FormattingCap As Integer,
     Optional DoTPMarkup As Boolean = False,
     Optional TPMarkupname As String = "",
     Optional CreatePodcast As Boolean = False,
     Optional FileObject As String = "",
     Optional DoPane As Boolean = False,
     Optional ChunkSize As Integer = 0,
     Optional NoFormatAndFieldSaving As Boolean = False,
     Optional DoNewDoc As Boolean = False,
     Optional SlideDeck As String = "",
     Optional AddDocs As Boolean = False,
     Optional DoMyStyle As Boolean = False,
     Optional DoBubblesExtract As Boolean = False,
     Optional DoPushback As Boolean = False,
     Optional SelectedTools As List(Of ModelConfig) = Nothing) As Task(Of String)


        Dim application As Word.Application = Globals.ThisAddIn.Application
        Dim selection As Microsoft.Office.Interop.Word.Selection = application.Selection

        If SysCommand = "" Then
            ShowCustomMessageBox("The (system-)prompt for the LLM is missing.")
            Return ""
        End If

        ' Route comment-bubble selections to a comment pipeline
        If selection IsNot Nothing AndAlso selection.StoryType = Word.WdStoryType.wdCommentsStory Then
            Dim bubbleSel As Word.Range = Nothing
            Try
                bubbleSel = selection.Range.Duplicate
            Catch
                bubbleSel = Nothing
            End Try


            Return Await ProcessSelectedTextInActiveCommentBubble(
                                SysCommand,
                                CheckMaxToken,
                                bubbleSel,
                                UseSecondAPI:=UseSecondAPI,
                                SelectionMandatory:=SelectionMandatory,
                                FileObject:=FileObject
                            )
        End If

        If selection IsNot Nothing AndAlso
   (selection.StoryType = Word.WdStoryType.wdFootnotesStory OrElse
    selection.StoryType = Word.WdStoryType.wdEndnotesStory) Then

            Return Await ProcessSelectedTextInActiveFootnote(
                        SysCommand,
                        CheckMaxToken,
                        UseSecondAPI,
                        SelectionMandatory,
                        FileObject
                    )
        End If

        If selection.Type = WdSelectionType.wdSelectionIP And SelectionMandatory Then
            ShowCustomMessageBox("Please select the text to be processed.")
            Return ""
        End If

        Try
            Using New WordUndoScope(application, $"{AN} Changes")

                If selection.Type = WdSelectionType.wdSelectionIP Or selection.Tables.Count = 0 Or PutInClipboard Or PutInBubbles Or DoPushback Then

                    Dim Result = Await TrueProcessSelectedText(SysCommand, CheckMaxToken, KeepFormat, ParaFormatInline, InPlace, DoMarkup, MarkupMethod, PutInClipboard, PutInBubbles, SelectionMandatory, UseSecondAPI, FormattingCap, DoTPMarkup, TPMarkupname, CreatePodcast, FileObject, DoPane, ChunkSize, NoFormatAndFieldSaving, DoNewDoc, SlideDeck, AddDocs, DoMyStyle, DoBubblesExtract, False, DoPushback, SelectedTools)

                Else

                    Dim userdialog As Integer = ShowCustomYesNoBox($"Your text contains tables. Shall each text section and each cell content be processed separately to avoid the table falling apart? This will take more time{If(SelectedAlternateModels IsNot Nothing, " and be done only with the initially selected model", "")}." & If(ChunkSize > 0, $" Your '(iterate)' parameter will apply only outside the tables.", "") & If(DoMarkup And MarkupMethod <> 2, " For the markup, the Diff markup will be used instead of the markup method choosen by you.", "") & " If you want to abort, close this window.", "No", "Yes, process each cell individually", $"{AN} Table Processing")

                    If userdialog = 0 Then Return ""

                    If userdialog = 2 Then

                        SelectedAlternateModels = Nothing

                        MarkupMethod = 2

                        Dim selRange As Range = selection.Range
                        Dim docTables As Tables = selRange.Tables

                        Dim isEntirelyWithinTable As Boolean = False
                        Dim isWholeTable As Boolean = False
                        Dim isPartialTableSelection As Boolean = False

                        If selection.Tables.Count = 1 Then
                            Dim tbl As Microsoft.Office.Interop.Word.Table = selRange.Tables(1)
                            Dim tblRange As Range = tbl.Range

                            ' Check if the selection is entirely within the table boundaries.
                            isEntirelyWithinTable = (selRange.Start >= tblRange.Start AndAlso selRange.End <= tblRange.End)

                            ' Get trimmed texts. Adjust the characters to trim as needed.
                            Dim selText As String = selRange.Text.Trim(vbCr, vbLf, " "c)
                            Dim tblText As String = tblRange.Text.Trim(vbCr, vbLf, " "c)

                            ' Compare the texts. If they differ, then the selection is not the whole table.
                            isWholeTable = (selText = tblText)

                            ' If the selection is fully contained in the table but does not equal the entire table's text,
                            ' then it is entirely within the table but is only a part of it.
                            If isEntirelyWithinTable AndAlso Not isWholeTable Then
                                isPartialTableSelection = True
                            End If
                        End If

                        If isEntirelyWithinTable Or isWholeTable Then

                            Dim sel As Word.Selection = application.Selection
                            Dim selCRange As Word.Range = sel.Range

                            ' Loop only the cells the user actually selected
                            For Each cell As Word.Cell In sel.Cells
                                ' Make a working copy of the cell's range, minus its end-of-cell marker
                                Dim cellRange As Word.Range = cell.Range.Duplicate
                                cellRange.End -= 1

                                ' Compute the overlap of selRange & cellRange
                                Dim intersection As Word.Range = selCRange.Duplicate
                                intersection.Start = System.Math.Max(cellRange.Start, selCRange.Start)
                                intersection.End = System.Math.Min(cellRange.End, selCRange.End)

                                ' If there is any overlap, process only that text
                                If intersection.Start < intersection.End Then
                                    ' Keep UI responsive
                                    System.Windows.Forms.Application.DoEvents()

                                    ' Show exactly what's being processed
                                    intersection.Select()

                                    ' Async processing call
                                    Dim result = Await TrueProcessSelectedText(
                                        SysCommand, CheckMaxToken, KeepFormat, ParaFormatInline,
                                        InPlace, DoMarkup, MarkupMethod, PutInClipboard,
                                        PutInBubbles, SelectionMandatory, UseSecondAPI,
                                        FormattingCap, DoTPMarkup, TPMarkupname, False,
                                        FileObject, DoPane, 0, NoFormatAndFieldSaving, DoNewDoc, "", AddDocs, DoMyStyle, DoBubblesExtract, True, SelectedTools:=SelectedTools)

                                    ' Throttle so Word doesn't lock up
                                    Await System.Threading.Tasks.Task.Delay(500)
                                End If
                            Next

                        Else

                            ' Sort tables by their start positions in the selection
                            Dim tableList As New List(Of Microsoft.Office.Interop.Word.Table)
                            For i As Integer = 1 To docTables.Count
                                tableList.Add(docTables(i))
                            Next
                            tableList.Sort(Function(t1, t2) t1.Range.Start.CompareTo(t2.Range.Start))

                            Dim lastPos As Integer = selRange.Start

                            Dim splash As New SLib.SplashScreen("Processing table(s)... press 'Esc' to abort")
                            splash.Show()
                            splash.Refresh()

                            Dim IsExit As Boolean = False

                            For Each tbl As Microsoft.Office.Interop.Word.Table In tableList

                                System.Windows.Forms.Application.DoEvents()

                                If (GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0 Then
                                    Exit For
                                End If

                                If (GetAsyncKeyState(VK_ESCAPE) And 1) <> 0 Or IsExit Then
                                    Exit For
                                End If

                                Dim tblStart As Integer = tbl.Range.Start
                                Dim tblEnd As Integer = tbl.Range.End

                                ' Text chunk BEFORE the table
                                If tblStart > lastPos Then
                                    Dim textChunk As Range = selRange.Duplicate
                                    textChunk.Start = lastPos
                                    textChunk.End = tblStart - 1

                                    ' Double-check you haven't snagged any table content
                                    If textChunk.Tables.Count = 0 Then
                                        ' Also verify it's not empty
                                        If textChunk.Start < textChunk.End Then
                                            textChunk.Select()
                                            Dim Result = Await TrueProcessSelectedText(SysCommand, CheckMaxToken, KeepFormat, ParaFormatInline, InPlace, DoMarkup, MarkupMethod, PutInClipboard, PutInBubbles, SelectionMandatory, UseSecondAPI, FormattingCap, DoTPMarkup, TPMarkupname, False, FileObject, DoPane, ChunkSize * -1, NoFormatAndFieldSaving, DoNewDoc, "", AddDocs, DoMyStyle, DoBubblesExtract, True, SelectedTools:=SelectedTools)
                                            Await System.Threading.Tasks.Task.Delay(500)
                                        End If
                                    Else

                                        Do
                                            textChunk.Start += 1
                                        Loop While textChunk.Tables.Count <> 0 And Not textChunk.Start = textChunk.End

                                        If textChunk.Tables.Count = 0 AndAlso textChunk.Start < textChunk.End Then
                                            textChunk.Select()
                                            Dim Result = Await TrueProcessSelectedText(SysCommand, CheckMaxToken, KeepFormat, ParaFormatInline, InPlace, DoMarkup, MarkupMethod, PutInClipboard, PutInBubbles, SelectionMandatory, UseSecondAPI, FormattingCap, DoTPMarkup, TPMarkupname, False, FileObject, DoPane, ChunkSize * -1, NoFormatAndFieldSaving, DoNewDoc, "", AddDocs, DoMyStyle, DoBubblesExtract, True, SelectedTools:=SelectedTools)
                                            Await System.Threading.Tasks.Task.Delay(500)
                                        End If

                                    End If
                                End If

                                ' Process the table itself (cells)
                                For Each row As Microsoft.Office.Interop.Word.Row In tbl.Rows
                                    System.Windows.Forms.Application.DoEvents()

                                    If (GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0 Then
                                        Exit For
                                    End If

                                    If (GetAsyncKeyState(VK_ESCAPE) And 1) <> 0 Or IsExit Then
                                        Exit For
                                    End If
                                    For Each cell As Microsoft.Office.Interop.Word.Cell In row.Cells
                                        System.Windows.Forms.Application.DoEvents()

                                        If (GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0 Then
                                            Exit For
                                        End If

                                        If (GetAsyncKeyState(VK_ESCAPE) And 1) <> 0 Or IsExit Then
                                            Exit For
                                        End If
                                        Dim cellRange As Range = cell.Range
                                        cellRange.End -= 1  ' Exclude cell marker
                                        If cellRange.Start < cellRange.End Then
                                            cellRange.Select()
                                            Dim Result = Await TrueProcessSelectedText(SysCommand, CheckMaxToken, KeepFormat, ParaFormatInline, InPlace, DoMarkup, MarkupMethod, PutInClipboard, PutInBubbles, SelectionMandatory, UseSecondAPI, FormattingCap, DoTPMarkup, TPMarkupname, False, FileObject, DoPane, 0, NoFormatAndFieldSaving, DoNewDoc, "", AddDocs, DoMyStyle, DoBubblesExtract, True, SelectedTools:=SelectedTools)
                                            Await System.Threading.Tasks.Task.Delay(500)
                                        End If
                                    Next
                                Next

                                ' Move lastPos to end of this table
                                lastPos = tblEnd + 1
                            Next

                            ' Text chunk AFTER the last table
                            If lastPos <= selRange.End And Not IsExit Then
                                Dim finalChunk As Range = selRange.Duplicate
                                finalChunk.Start = lastPos
                                finalChunk.End = selRange.End

                                If finalChunk.Tables.Count = 0 AndAlso finalChunk.Start < finalChunk.End Then

                                    finalChunk.Select()
                                    Dim text = selection.Text
                                    Dim Result = Await TrueProcessSelectedText(SysCommand, CheckMaxToken, KeepFormat, ParaFormatInline, InPlace, DoMarkup, MarkupMethod, PutInClipboard, PutInBubbles, SelectionMandatory, UseSecondAPI, FormattingCap, DoTPMarkup, TPMarkupname, False, FileObject, DoPane, ChunkSize * -1, NoFormatAndFieldSaving, DoNewDoc, "", AddDocs, DoMyStyle, DoBubblesExtract, True, SelectedTools:=SelectedTools)
                                Else
                                    Do
                                        finalChunk.Start += 1
                                    Loop While finalChunk.Tables.Count <> 0 And Not finalChunk.Start = finalChunk.End

                                    finalChunk.End = selRange.End

                                    If finalChunk.Tables.Count = 0 AndAlso finalChunk.Start < finalChunk.End Then
                                        finalChunk.Select()
                                        Dim Result = Await TrueProcessSelectedText(SysCommand, CheckMaxToken, KeepFormat, ParaFormatInline, InPlace, DoMarkup, MarkupMethod, PutInClipboard, PutInBubbles, SelectionMandatory, UseSecondAPI, FormattingCap, DoTPMarkup, TPMarkupname, False, FileObject, DoPane, ChunkSize * -1, NoFormatAndFieldSaving, DoNewDoc, "", AddDocs, DoMyStyle, DoBubblesExtract, True, SelectedTools:=SelectedTools)
                                    End If
                                End If
                            End If

                            splash.Close()
                        End If

                    ElseIf userdialog = 1 Then

                        Dim Result = Await TrueProcessSelectedText(SysCommand, CheckMaxToken, KeepFormat, ParaFormatInline, InPlace, DoMarkup, MarkupMethod, PutInClipboard, PutInBubbles, SelectionMandatory, UseSecondAPI, FormattingCap, DoTPMarkup, TPMarkupname, CreatePodcast, FileObject, DoPane, ChunkSize, NoFormatAndFieldSaving, DoNewDoc, SlideDeck, AddDocs, DoMyStyle, DoBubblesExtract, False, DoPushback, SelectedTools)

                    End If

                End If

                InsertDocs = ""

                If Not PutInClipboard Then
                    selection.Collapse(WdCollapseDirection.wdCollapseEnd)
                    selection.MoveStart(WdUnits.wdCharacter, 0)
                    selection.MoveEnd(WdUnits.wdCharacter, 0)
                End If

                Return ""

            End Using

        Catch ex As System.Exception
            Debug.WriteLine("Error in Undo: " & ex.Message)
        End Try

    End Function

    ''' <summary>
    ''' Core processing routine. Iterates paragraphs in chunks, captures formatting, invokes the LLM,
    ''' applies post-processing, and routes output to the requested surface (pane, clipboard, doc, etc.).
    ''' </summary>
    ''' <param name="SysCommand">System prompt.</param>
    ''' <param name="CheckMaxToken">Check token limits.</param>
    ''' <param name="KeepFormat">Preserve HTML formatting.</param>
    ''' <param name="ParaFormatInline">Include inline paragraph format markers.</param>
    ''' <param name="InPlace">Replace selection in place.</param>
    ''' <param name="DoMarkup">Generate markup.</param>
    ''' <param name="MarkupMethod">1=Word, 2=Diff, 3=DiffWindow, 4=Regex.</param>
    ''' <param name="PutInClipboard">Place result in clipboard.</param>
    ''' <param name="PutInBubbles">Place result as Word comments.</param>
    ''' <param name="SelectionMandatory">Require text selection.</param>
    ''' <param name="UseSecondAPI">Use secondary API.</param>
    ''' <param name="FormattingCap">Character limit for format preservation.</param>
    ''' <param name="DoTPMarkup">Mark revisions for a specific author.</param>
    ''' <param name="TPMarkupname">Author name for revision filtering.</param>
    ''' <param name="CreatePodcast">Generate podcast script.</param>
    ''' <param name="FileObject">File path for additional LLM input.</param>
    ''' <param name="DoPane">Show result in a pane.</param>
    ''' <param name="ChunkSize">Paragraph chunk size.</param>
    ''' <param name="NoFormatAndFieldSaving">Skip format/field saving.</param>
    ''' <param name="DoNewDoc">Place result in a new document.</param>
    ''' <param name="SlideDeck">PowerPoint deck path.</param>
    ''' <param name="AddDocs">Include insertdocs.</param>
    ''' <param name="DoMyStyle">Use MyStyle.</param>
    ''' <param name="DoBubblesExtract">Extract from bubbles.</param>
    ''' <param name="InTable">Processing within table context.</param>
    ''' <param name="DoPushback">Reply to bubbles.</param>
    ''' <param name="SelectedTools">The tools selected for tooling runs</param>
    ''' <returns>Empty string on completion.</returns>
    Private Async Function TrueProcessSelectedText(SysCommand As String,
       CheckMaxToken As Boolean,
       KeepFormat As Boolean,
       ParaFormatInline As Boolean,
       InPlace As Boolean,
       DoMarkup As Boolean,
       MarkupMethod As Integer,
       PutInClipboard As Boolean,
       PutInBubbles As Boolean,
       SelectionMandatory As Boolean,
       UseSecondAPI As Boolean,
       FormattingCap As Integer,
       Optional DoTPMarkup As Boolean = False,
       Optional TPMarkupname As String = "",
       Optional CreatePodcast As Boolean = False,
       Optional FileObject As String = "",
       Optional DoPane As Boolean = False,
       Optional ChunkSize As Integer = 0,
       Optional NoFormatAndFieldSaving As Boolean = False,
       Optional DoNewDoc As Boolean = False,
       Optional SlideDeck As String = "",
       Optional AddDocs As Boolean = False,
       Optional DoMyStyle As Boolean = False,
       Optional DoBubblesExtract As Boolean = False,
       Optional InTable As Boolean = False,
       Optional DoPushback As Boolean = False,
       Optional SelectedTools As List(Of ModelConfig) = Nothing) As Task(Of String)


        Dim application As Word.Application = Globals.ThisAddIn.Application
        Dim selection As Microsoft.Office.Interop.Word.Selection = application.Selection
        Dim currentdoc As Word.Document = selection.Document

        ' ============= ENSURE WE'RE IN MAIN STORY WITHOUT CHANGING SELECTION =============
        Try
            If currentdoc IsNot Nothing AndAlso selection IsNot Nothing Then
                Dim currentStory As Word.WdStoryType = selection.StoryType

                ' Only act if we're NOT already in the main text story
                If currentStory <> Word.WdStoryType.wdMainTextStory Then
                    ' Force view back to print view to get out of special editing modes
                    application.ActiveWindow.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView

                    ' Move to start of main document story without selecting anything
                    Dim mainStoryRange As Word.Range = currentdoc.StoryRanges(Word.WdStoryType.wdMainTextStory)
                    mainStoryRange.Collapse(Word.WdCollapseDirection.wdCollapseStart)
                    mainStoryRange.Select()

                    ' Collapse to insertion point (no selection)
                    selection.Collapse(Word.WdCollapseDirection.wdCollapseStart)
                End If
            End If
        Catch ex As Exception
            ' Best-effort; continue even if this fails
            Debug.WriteLine($"Warning: Could not reset to main story: {ex.Message}")
        End Try
        ' ================================================================================

        Debug.WriteLine(
                    vbCrLf & "CheckMaxToken=" & CheckMaxToken &
                    vbCrLf & "KeepFormat=" & KeepFormat &
                    vbCrLf & "ParaFormatInline=" & ParaFormatInline &
                    vbCrLf & "InPlace=" & InPlace &
                    vbCrLf & "DoMarkup=" & DoMarkup &
                    vbCrLf & "PutInClipboard=" & PutInClipboard &
                    vbCrLf & "PutInBubbles=" & PutInBubbles &
                    vbCrLf & "SelectionMandatory=" & SelectionMandatory &
                    vbCrLf & "UseSecondAPI=" & UseSecondAPI &
                    vbCrLf & "DoTPMarkup=" & DoTPMarkup &
                    vbCrLf & "CreatePodcast=" & CreatePodcast &
                    vbCrLf & "DoPane=" & DoPane &
                    vbCrLf & "DoNewDoc=" & DoNewDoc &
                    vbCrLf & "Chunksize=" & ChunkSize &
                    vbCrLf & "Fileobject=" & FileObject &
                    vbCrLf & "Slidedeck=" & SlideDeck &
                    vbCrLf & "NoFormatAndFieldSaving=" & NoFormatAndFieldSaving &
                    vbCrLf & "AddDocs=" & AddDocs &
                    vbCrLf & "DoMyStyle=" & DoMyStyle &
                    vbCrLf & "DoBubblesExtract=" & DoBubblesExtract &
                    vbCrLf & "InTable=" & InTable &
                    vbCrLf & "DoPushback=" & DoPushback
                )

        Try

            Dim SelectedText As String = ""
            Dim rng As Range
            Dim i As Integer
            Dim NoFormatting As Boolean = False
            Dim NoSelectedText As Boolean = False
            Dim trailingCR As Boolean
            Dim trailingCRcount As Integer = 0
            Dim DoSilent As Boolean = False

            If selection.Type = WdSelectionType.wdSelectionIP And SelectionMandatory Then
                Return ""
            End If

            If selection.Type = WdSelectionType.wdSelectionIP Then NoSelectedText = True

            rng = selection.Range

            Debug.WriteLine($"1Range Start = {rng.Start} Selection Start = {selection.Start}")
            Debug.WriteLine($"Range End = {rng.End} Selection End = {selection.End}")
            Debug.WriteLine(vbCrLf & Left(rng.Text, 400) & vbCrLf)

            ' Added for processing footnotes etc. too
            ' What story (main text, header, footer, footnote, etc.) am I in?
            Dim storyType As Word.WdStoryType = rng.StoryType
            ' Grab the full Range for that story from the document
            Dim storyRange As Word.Range = currentdoc.StoryRanges(storyType)

            If Not NoSelectedText Then
                If rng.Text.Length = 0 Then NoSelectedText = True
                If Not NoSelectedText And (KeepFormat Or ParaFormatInline) And FormattingCap > 0 And rng.Text.Length > FormattingCap Then NoFormatting = True
            End If

            If PutInBubbles Or PutInClipboard Or NoSelectedText Or DoPushback Then NoFormatting = True

            If PutInBubbles Or DoPushback Then
                DoMarkup = False
                PutInClipboard = False
            End If

            If PutInClipboard Then DoMarkup = False

            If DoTPMarkup Then NoFormatting = True

            If MarkupMethod = 4 And DoMarkup Then NoFormatting = True

            If ChunkSize > 0 Then
                DoSilent = True
                If DoMarkup Then
                    Select Case MarkupMethod
                        Case 1
                            Dim SilentMarkup As Integer = SLib.ShowCustomYesNoBox($"You have choosen both iterated processing and markups using Word compare. Iteration only works using the Regex method. Continue using Regex markup (the character cap will be ignored) or go without markups?", "Yes, Regex markups", "No, no markups")
                            If SilentMarkup = 1 Then
                                MarkupMethod = 4
                            ElseIf SilentMarkup = 2 Then
                                DoMarkup = False
                            Else
                                Return ""
                            End If
                        Case 2
                            Dim SilentMarkup As Integer = SLib.ShowCustomYesNoBox($"You have choosen both iterated processing and markups using Diff compare. Iteration only works using the Regex method. Continue using Diff markup (the character cap will be ignored) or go without markups?", "Yes, Diff markups", "No, no markups")
                            If SilentMarkup = 1 Then
                                MarkupMethod = 2
                            ElseIf SilentMarkup = 2 Then
                                DoMarkup = False
                            Else
                                Return ""
                            End If
                        Case 3
                            Dim SilentMarkup As Integer = SLib.ShowCustomYesNoBox($"You have choosen both iterated processing and markups using DiffW compare. Iteration only works using the Regex method. Continue using Diff markup (the character cap will be ignored) or go without markups?", "Yes, Diff markups", "No, no markups")
                            If SilentMarkup = 1 Then
                                MarkupMethod = 2
                            ElseIf SilentMarkup = 2 Then
                                DoMarkup = False
                            Else
                                Return ""
                            End If
                        Case 4
                            Dim SilentMarkup As Integer = SLib.ShowCustomYesNoBox($"You have choosen both iterated processing and markups using the Regex method. This works, but may tak a very long time (the character cap will be ignored). Continue with Regex markups or go without markups?", "Yes, Regex markups", "No, no markups")
                            If SilentMarkup = 1 Then
                                MarkupMethod = 4
                            ElseIf SilentMarkup = 2 Then
                                DoMarkup = False
                            Else
                                Return ""
                            End If
                    End Select
                End If
            End If

            If ChunkSize < 0 Then
                ChunkSize = ChunkSize * -1
                If DoMarkup Then MarkupMethod = 2  ' Force Diff when getting a negative ChunkSize (e.g., in tables)
                DoSilent = True
            End If

            Dim effectiveChunk As Integer = If(ChunkSize > 0, ChunkSize, Integer.MaxValue)

            Dim totalEndBm As Word.Bookmark

            ' Added for processing footnotes etc. too
            Dim docEnd As Integer = storyRange.End

            If selection.End < docEnd Then
                totalEndBm = currentdoc.Bookmarks.Add(
                    Name:="TotalEnd",
                    Range:=currentdoc.Range(Start:=selection.End, End:=selection.End))
            Else
                Dim endRange As Word.Range = storyRange.Duplicate
                endRange.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                totalEndBm = currentdoc.Bookmarks.Add(
                    Name:="TotalEnd",
                    Range:=endRange)
            End If

            Debug.WriteLine($"2Range Start = {rng.Start} Selection Start = {selection.Start}")
            Debug.WriteLine($"Range End = {rng.End} Selection End = {selection.End}")
            Debug.WriteLine(vbCrLf & Left(rng.Text, 400) & vbCrLf)

            Dim safeRange As Word.Range = selection.Range
            safeRange.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseStart)

            Dim nextStartBm As Word.Bookmark = currentdoc.Bookmarks.Add(
                    Name:="NextStart",
                    Range:=safeRange)

            Do While NoSelectedText OrElse (currentdoc.Bookmarks.Exists("NextStart") AndAlso currentdoc.Bookmarks.Exists("TotalEnd") AndAlso currentdoc.Bookmarks("NextStart").Range.Start < currentdoc.Bookmarks("TotalEnd").Range.Start)

                Try

                    If Not NoSelectedText Then

                        Dim curStart As Integer = currentdoc.Bookmarks("NextStart").Range.Start
                        Dim totalEnd As Integer = currentdoc.Bookmarks("TotalEnd").Range.Start

                        Do While curStart < totalEnd AndAlso currentdoc.Range(Start:=curStart, End:=System.Math.Min(curStart + 1, totalEnd)).Text = vbCr
                            curStart += 1
                        Loop
                        If curStart >= totalEnd Then Exit Do

                        ' ---- 2.1  Determine chunk end ----------------------------
                        docEnd = storyRange.End
                        Dim restRng As Word.Range = currentdoc.Range(Start:=curStart, End:=totalEnd)
                        Dim paras As Word.Paragraphs = restRng.Paragraphs

                        Dim chunkEnd As Integer

                        If paras.Count <= effectiveChunk Then
                            chunkEnd = totalEnd
                        Else
                            ' Start with the end of the effectiveChunk-th paragraph
                            Dim xxi As Integer = effectiveChunk
                            Dim paraRng As Word.Range = paras(xxi).Range
                            Dim paraText As String = paraRng.Text.Trim()

                            ' Keep extending while paragraph is empty and more paras are available
                            Do While (paraText = "" OrElse paraText = vbCr) AndAlso xxi < paras.Count
                                xxi += 1
                                paraRng = paras(xxi).Range
                                paraText = paraRng.Text.Trim()
                            Loop

                            chunkEnd = paraRng.End
                        End If

                        ' Secure boundaries to avoid Range errors
                        If chunkEnd > totalEnd Then chunkEnd = totalEnd
                        If chunkEnd <= curStart Then chunkEnd = System.Math.Min(curStart + 1, totalEnd)

                        ' ---- 2.2  Selection on this chunk ----------------------
                        selection.SetRange(Start:=curStart, End:=chunkEnd)
                        rng = selection.Range

                        If rng Is Nothing OrElse rng.Text.Trim() = "" Then Exit Do

                    End If
                Catch ex As System.Exception
                    Exit Do
                End Try

                Debug.WriteLine($"3Range Start = {rng.Start} Selection Start = {selection.Start}")
                Debug.WriteLine($"Range End = {rng.End} Selection End = {selection.End}")
                Debug.WriteLine(vbCrLf & Left(rng.Text, 400) & vbCrLf)

                paraCount = 0
                trailingCR = False
                trailingCRcount = 0

                If Not ParaFormatInline AndAlso Not NoFormatting AndAlso Not NoSelectedText AndAlso Not NoFormatAndFieldSaving Then

                    paraCount = rng.Paragraphs.Count

                    ReDim paragraphFormat(paraCount - 1)
                    Array.Clear(paragraphFormat, 0, paragraphFormat.Length)

                    For i = 1 To paraCount
                        Dim para As Word.Paragraph = rng.Paragraphs(i)
                        Dim paraRange As Word.Range = para.Range

                        ' bodyRange = text without the paragraph mark
                        Dim bodyRange As Word.Range = paraRange.Duplicate
                        bodyRange.MoveEnd(Word.WdUnits.wdCharacter, -1)

                        Try
                            ' Character-level attributes – store only when uniform
                            Dim boldV As Integer? = Nothing
                            Dim italicV As Integer? = Nothing
                            Dim underlineV As Word.WdUnderline? = Nothing
                            Dim colorV As Long? = Nothing

                            If bodyRange.Font.Bold <> Word.WdConstants.wdUndefined Then _
                                boldV = bodyRange.Font.Bold
                            If bodyRange.Font.Italic <> Word.WdConstants.wdUndefined Then _
                                italicV = bodyRange.Font.Italic
                            If bodyRange.Font.Underline <> Word.WdConstants.wdUndefined Then _
                                underlineV = CType(bodyRange.Font.Underline, Word.WdUnderline)
                            If bodyRange.Font.Color <> Word.WdConstants.wdUndefined Then _
                                colorV = bodyRange.Font.Color

                            Dim fname As String = Nothing
                            Dim fsize As Single? = Nothing
                            If bodyRange.Font.Name <> CStr(Word.WdConstants.wdUndefined) Then _
                                fname = bodyRange.Font.Name
                            If bodyRange.Font.Size <> CSng(Word.WdConstants.wdUndefined) Then _
                                fsize = bodyRange.Font.Size

                            ' Assign into the (freshly resized) array
                            paragraphFormat(i - 1) = New ParagraphFormatStructure With {
                                .Style = para.Style,
                                .FontName = fname,
                                .FontSize = fsize,
                                .FontBold = boldV,
                                .FontItalic = italicV,
                                .FontUnderline = underlineV,
                                .FontColor = colorV,
                                .ListType = bodyRange.ListFormat.ListType,
                                .ListTemplate = If(bodyRange.ListFormat.ListType <>
                                                    Word.WdListType.wdListNoNumbering,
                                                    bodyRange.ListFormat.ListTemplate, Nothing),
                                .ListLevel = If(bodyRange.ListFormat.ListType <>
                                                    Word.WdListType.wdListNoNumbering,
                                                    bodyRange.ListFormat.ListLevelNumber, 0),
                                .ListNumber = If(bodyRange.ListFormat.ListType <>
                                                    Word.WdListType.wdListNoNumbering,
                                                    bodyRange.ListFormat.ListValue, 0),
                                .HasListFormat = bodyRange.ListFormat.ListType <>
                                                    Word.WdListType.wdListNoNumbering,
                                .Alignment = para.Alignment,
                                .LineSpacing = para.LineSpacing,
                                .SpaceBefore = para.SpaceBefore,
                                .SpaceAfter = para.SpaceAfter,
                                .LineSpacingRule = para.LineSpacingRule,
                                .SpaceBeforeAuto = para.SpaceBeforeAuto,
                                .SpaceAfterAuto = para.SpaceAfterAuto,
                                .DisableLineHeightGrid = para.DisableLineHeightGrid
                                        }

                        Catch ex As System.Exception
                            ' Error extracting paragraph formatting
                        End Try
                    Next

                End If

                Debug.WriteLine($"4Range Start = {rng.Start} Selection Start = {selection.Start}")
                Debug.WriteLine($"Range End = {rng.End} Selection End = {selection.End}")
                Debug.WriteLine(vbCrLf & Left(rng.Text, 400) & vbCrLf)

                Dim raw As String = ""

                If (PutInBubbles Or PutInClipboard) AndAlso Not DoTPMarkup AndAlso rng IsNot Nothing Then
                    raw = GetVisibleText(rng)
                End If

                Debug.WriteLine($"4aRange Start = {rng.Start} Selection Start = {selection.Start}")
                Debug.WriteLine($"Range End = {rng.End} Selection End = {selection.End}")

                Dim BubblesText As String = ""

                If DoBubblesExtract Then BubblesText = BubblesExtract(rng, DoSilent Or InTable)

                Debug.WriteLine($"BubblesText = '{BubblesText}'")

                If Not NoSelectedText Then

                    If KeepFormat And Not NoFormatting Then
                        SelectedText = SLib.GetRangeHtml(rng)
                    Else
                        If NoFormatting OrElse NoFormatAndFieldSaving Then
                            If DoTPMarkup Then
                                SelectedText = AddMarkupTags(rng, TPMarkupname)
                            Else
                                SelectedText = rng.Text
                                If Not String.IsNullOrWhiteSpace(raw) Then SelectedText = raw
                            End If
                        Else
                            If INI_MarkdownConvert AndAlso Not KeepFormat AndAlso (Not DoMarkup OrElse (MarkupMethod = 3 Or MarkupMethod = 2)) AndAlso InPlace Then

                                Debug.WriteLine($"4bRange Start = {rng.Start} Selection Start = {selection.Start}")
                                Debug.WriteLine($"Range End = {rng.End} Selection End = {selection.End}")
                                Debug.WriteLine(vbCrLf & Left(rng.Text, 400) & vbCrLf)
                                Debug.WriteLine("SelectedText: " & SelectedText)

                                SelectedText = GetTextWithSpecialElementsInline(rng, If(NoFormatting, False, ParaFormatInline), True)

                                Debug.WriteLine($"4cRange Start = {rng.Start} Selection Start = {selection.Start}")
                                Debug.WriteLine($"Range End = {rng.End} Selection End = {selection.End}")
                                Debug.WriteLine(vbCrLf & Left(rng.Text, 400) & vbCrLf)
                                Debug.WriteLine("SelectedText: " & SelectedText)

                            Else
                                SelectedText = GetTextWithSpecialElementsInline(rng, If(NoFormatting, False, ParaFormatInline), False)
                                'SelectedText = LegacyGetTextWithSpecialElementsInline(rng, If(NoFormatting, False, ParaFormatInline))
                            End If
                        End If
                        trailingCR = (SelectedText.EndsWith(vbCrLf) Or SelectedText.EndsWith(vbLf) Or SelectedText.EndsWith(vbCr))
                        Dim tempText As String = SelectedText

                        Do While tempText.EndsWith(vbCrLf) Or tempText.EndsWith(vbLf) Or tempText.EndsWith(vbCr)
                            If tempText.EndsWith(vbCrLf) Then
                                trailingCRcount += 1
                                tempText = tempText.Substring(0, tempText.Length - vbCrLf.Length)
                            ElseIf tempText.EndsWith(vbLf) Then
                                trailingCRcount += 1
                                tempText = tempText.Substring(0, tempText.Length - vbLf.Length)
                            ElseIf tempText.EndsWith(vbCr) Then
                                trailingCRcount += 1
                                tempText = tempText.Substring(0, tempText.Length - vbCr.Length)
                            End If
                        Loop

                    End If

                    Debug.WriteLine($"4dRange Start = {rng.Start} Selection Start = {selection.Start}")
                    Debug.WriteLine($"Range End = {rng.End} Selection End = {selection.End}")
                    Debug.WriteLine("SelectedText: " & SelectedText)

                    Dim MaxToken As Integer = If(UseSecondAPI, INI_MaxOutputToken_2, INI_MaxOutputToken)
                    Dim EstimatedTokens As Integer = EstimateTokenCount(SelectedText)

                    If CheckMaxToken And MaxToken > 0 AndAlso EstimatedTokens > MaxToken AndAlso (InPlace Or DoMarkup) AndAlso Not DoSilent Then
                        ShowCustomMessageBox("Your selected text Is larger than the maximum output your LLM can supposedly generate. Therefore, the output may be shorter than expected based on maximum tokens supported, which Is " & MaxToken & " tokens. Your input (with formatting information, as the case may be) has an estimated to be " & EstimatedTokens & " tokens). Therefore, check whether the output Is complete.", AN, 15)
                    End If

                    If DoMarkup AndAlso MarkupMethod = 2 AndAlso Len(SelectedText) > INI_MarkupDiffCap AndAlso Not DoSilent Then
                        Dim MarkupChange As Integer = SLib.ShowCustomYesNoBox($"The selected text exceeds the defined cap for the Diff markup method at {INI_MarkupDiffCap} chars (your selection has {Len(SelectedText)} chars). {If(KeepFormat, "This may be because HTML codes have been inserted to keep the formatting (you can turn this off in the settings). ", "")}How do you want to continue?", "Use Diff in Window compare instead", "Use Diff")
                        Select Case MarkupChange
                            Case 1
                                MarkupMethod = 3
                            Case 2
                                MarkupMethod = 2
                            Case Else
                                Return ""
                        End Select
                    End If

                    If DoMarkup And MarkupMethod = 4 And Len(SelectedText) > INI_MarkupRegexCap AndAlso Not DoSilent Then
                        Dim MarkupChange As Integer = SLib.ShowCustomYesNoBox($"The selected text exceeds the defined cap for the Regex markup method at {INI_MarkupRegexCap} chars (your selection has {Len(SelectedText)} chars). {If(KeepFormat, "This may be because HTML codes have been inserted to keep the formatting (you can turn this off in the settings). ", "")}How do you want to continue?", "Use Word compare instead", "Use Regex")
                        Select Case MarkupChange
                            Case 1
                                MarkupMethod = 1
                            Case 2
                                MarkupMethod = 4
                            Case Else
                                Return ""
                        End Select
                    End If

                Else

                    SelectedText = ""

                End If

                Debug.WriteLine($"5Range Start = {rng.Start} Selection Start = {selection.Start}")
                Debug.WriteLine($"Range End = {rng.End} Selection End = {selection.End}")
                Debug.WriteLine(vbCrLf & Left(rng.Text, 400) & vbCrLf)
                Debug.WriteLine("SelectedText: " & SelectedText)

                Dim SlideInsert As String = ""

                If SlideDeck <> "" Then
                    SlideInsert = GetPresentationJson(SlideDeck)
                    Debug.WriteLine("SlideInsert = " & SlideInsert)
                    If SlideDeck = "" Then
                        Return ""
                    Else
                        SlideInsert = " <SLIDEDECK>" & SlideInsert & "</SLIDEDECK>"
                    End If
                End If

                Dim LLMResult As String = ""

                ' If tools are selected and the model supports tooling, enter the tool execution loop
                If SelectedTools IsNot Nothing AndAlso SelectedTools.Count > 0 AndAlso UseSecondAPI Then
                    LLMResult = Await ExecuteToolingLoop(
                        SysCommand,
                        SelectedText,
                        SelectedTools,
                        UseSecondAPI,
                        FileObject,
                        DoTPMarkup,
                        BubblesText,
                        NoFormatting,
                        KeepFormat,
                        SlideDeck,
                        DoMyStyle,
                        MyStyleInsert,
                        AddDocs,
                        InsertDocs,
                        SlideInsert,
                        OtherPrompt)
                Else

                    ' Other LLM calls w/o Tooling

                    If SelectedAlternateModels Is Nothing OrElse SelectedAlternateModels.Count = 0 OrElse DoMarkup OrElse PutInBubbles OrElse DoPushback OrElse SlideInsert <> "" Then

                        LLMResult = Await LLM(SysCommand & If(String.IsNullOrWhiteSpace(BubblesText), "", " " & SP_Add_BubblesExtract) & If(DoTPMarkup, " " & SP_Add_Revisions, "") & " " & If(SlideDeck = "", If(NoFormatting, "", If(KeepFormat, " " & SP_Add_KeepHTMLIntact, " " & SP_Add_KeepInlineIntact)), " " & SP_Add_Slides) & If(DoMyStyle, " " & MyStyleInsert, ""), If(NoSelectedText, If(AddDocs, " " & InsertDocs & " ", "") & SlideInsert, "<TEXTTOPROCESS>" & SelectedText & "</TEXTTOPROCESS>" & If(AddDocs, " " & InsertDocs & " ", "") & SlideInsert & " " & BubblesText), "", "", 0, UseSecondAPI, False, OtherPrompt, FileObject)

                    Else

                        For Each mc As ModelConfig In SelectedAlternateModels
                            Dim err As Boolean = False
                            ApplyModelConfig(_context, mc, err)

                            LLMResult += mc.ModelDescription & ":" & vbCrLf & vbCrLf & Await LLM(SysCommand & If(String.IsNullOrWhiteSpace(BubblesText), "", " " & SP_Add_BubblesExtract) & If(DoTPMarkup, " " & SP_Add_Revisions, "") & " " & If(SlideDeck = "", If(NoFormatting, "", If(KeepFormat, " " & SP_Add_KeepHTMLIntact, " " & SP_Add_KeepInlineIntact)), " " & SP_Add_Slides) & If(DoMyStyle, " " & MyStyleInsert, ""), If(NoSelectedText, If(AddDocs, " " & InsertDocs & " ", "") & SlideInsert, "<TEXTTOPROCESS>" & SelectedText & "</TEXTTOPROCESS>" & If(AddDocs, " " & InsertDocs & " ", "") & SlideInsert & " " & BubblesText), "", "", 0, UseSecondAPI, False, OtherPrompt, FileObject) & vbCrLf

                        Next

                    End If
                End If

                OtherPrompt = ""

                LLMResult = LLMResult.Replace("<TEXTTOPROCESS>", "").Replace("</TEXTTOPROCESS>", "")

                If Not String.IsNullOrEmpty(LLMResult) Then
                    LLMResult = Await PostCorrection(LLMResult, UseSecondAPI)
                End If

                Debug.WriteLine($"LLMResult 0 = '{LLMResult}'")

                ' Remove horizontal whitespace (incl. NBSP) between real newline tokens (CRLF/CR/LF)
                LLMResult = System.Text.RegularExpressions.Regex.Replace(LLMResult, "(\r\n|\r|\n)[^\S\r\n]+(\r\n|\r|\n)", "$1$2")

                Debug.WriteLine($"LLMResult 1 = '{LLMResult}'")

                If ParaFormatInline Then LLMResult = CorrectPFORMarkers(LLMResult)

                Debug.WriteLine($"LLMResult 2 = '{LLMResult}'")

                If DoTPMarkup Then LLMResult = RemoveMarkupTags(LLMResult)

                If Not trailingCR AndAlso LLMResult.EndsWith(ControlChars.CrLf) Then LLMResult = LLMResult.TrimEnd(ControlChars.CrLf)
                If Not trailingCR AndAlso LLMResult.EndsWith(ControlChars.Lf) Then LLMResult = LLMResult.TrimEnd(ControlChars.Lf)
                If Not trailingCR AndAlso LLMResult.EndsWith(ControlChars.Cr) Then LLMResult = LLMResult.TrimEnd(ControlChars.Cr)

                If trailingCR Then
                    LLMResult = LLMResult.TrimEnd({ControlChars.Cr, ControlChars.Lf})
                    If trailingCRcount > 1 Then
                        LLMResult &= String.Concat(Enumerable.Repeat(vbCrLf, trailingCRcount - 1))
                    End If
                End If

                Debug.WriteLine($"LLMResult 3 = '{LLMResult}'")
                Debug.WriteLine($"TrailingCR = {trailingCR} Count = {trailingCRcount}")

                Debug.WriteLine($"6Range Start = {rng.Start} Selection Start = {selection.Start}")
                Debug.WriteLine($"Range End = {rng.End} Selection End = {selection.End}")
                Debug.WriteLine(vbCrLf & Left(rng.Text, 400) & vbCrLf)

                If Not String.IsNullOrEmpty(LLMResult) Then


                    Debug.WriteLine(
                                vbCrLf & "PutInClipboard=" & PutInClipboard &
                                vbCrLf & "DoSilent=" & DoSilent &
                                vbCrLf & "DoMarkup=" & DoMarkup &
                                vbCrLf & "MarkupMethod=" & MarkupMethod &
                                vbCrLf & "DoPane=" & DoPane &
                                vbCrLf & "DoNewDoc=" & DoNewDoc &
                                vbCrLf & "PutInBubbles=" & PutInBubbles &
                                vbCrLf & "NoSelectedText=" & NoSelectedText &
                                vbCrLf & "ParaFormatInline=" & ParaFormatInline &
                                vbCrLf & "NoFormatting=" & NoFormatting &
                                vbCrLf & "NoFormatAndFieldSaving=" & NoFormatAndFieldSaving &
                                vbCrLf & "KeepFormat=" & KeepFormat &
                                vbCrLf & "Inplace=" & InPlace
                            )

                    Dim ClipPaneText1 As String = "The LLM has provided the following result (you can edit it)"
                    Dim ClipText2 As String = "You can choose whether you want to have the original text put into the clipboard Or your text with any changes you have made (without formatting), Or you can directly insert the original text in your document. If you select Cancel, nothing will be put into the clipboard."
                    Dim PaneText2 As String = "Choose to put your edited Or original text in the clipboard, Or inserted the original with formatting; the pane will close. You can also copy & paste from the pane."

                    If DoPushback Then

                        Dim app As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application
                        If _uiContext IsNot Nothing Then
                            _uiContext.Post(
                                Sub(s)
                                    Dim uiSel As Microsoft.Office.Interop.Word.Selection = app.Selection
                                    ReplyBubbles(LLMResult, uiSel, DoSilent)
                                End Sub, Nothing)
                        Else
                            ' Fallback – assume we are already on UI thread
                            Dim uiSel As Microsoft.Office.Interop.Word.Selection = app.Selection
                            ReplyBubbles(LLMResult, uiSel, DoSilent)
                        End If

                    ElseIf CreatePodcast AndAlso Not DoSilent Then
                        Dim TTSAvailable As Boolean = False

                        DetectTTSEngines()

                        If Not TTS_googleAvailable AndAlso Not TTS_openAIAvailable Then
                            TTSAvailable = False
                        Else
                            TTSAvailable = True
                        End If

                        LLMResult = NormalizeHostGuestConversation(LLMResult)

                        If TTSAvailable Then
                            Dim FinalText = ShowCustomWindow("The LLM has created the following podcast script for you (you can edit it; you do Not have to manually remove the SSML codes, if you do Not Like them)", LLMResult, "The next step Is the production of an audio file. You can choose whether you want to use the original text or your text with any changes you have made. The text will also be put in the clipboard. If you select Cancel, the original text will only be put into the clipboard.", AN, True)

                            If FinalText = "" Then
                                SLib.PutInClipboard(MarkdownToRtfConverter.Convert(LLMResult))
                            Else
                                FinalText = FinalText.Trim()
                                SLib.PutInClipboard(FinalText)
                                If FinalText.Contains("H: ") AndAlso FinalText.Contains("G: ") Then ReadPodcast(FinalText)
                            End If
                        Else
                            Dim FinalText = ShowCustomWindow("The LLM has created the following podcast script for you (you can edit it; you do Not have to manually remove the SSML codes, if you do Not Like them)", LLMResult, $"The next step Is the production of an audio file. Since you have not configured {AN} for Google, you unfortunately cannot do that here. However, you can choose whether you want the original text Or the text with your changes to put in the clipboard for further use. If you select Cancel, no text will be put in the clipboard.", AN, True)

                            If FinalText <> "" Then
                                SLib.PutInClipboard(MarkdownToRtfConverter.Convert(LLMResult))
                            Else
                                FinalText = FinalText.Trim()
                                SLib.PutInClipboard(FinalText)
                            End If
                        End If
                    ElseIf SlideInsert <> "" Then

                        Dim Jsonstring As String = CleanJsonString(LLMResult)

                        Debug.WriteLine("JsonString=" & Jsonstring)

                        If Not String.IsNullOrEmpty(Jsonstring) Then

                            If ApplyPlanToPresentation(SlideDeck, Jsonstring) Then

                                Dim TokenErrorResponse As String = ValidatePptx(SlideDeck)

                                If TokenErrorResponse = "" Then
                                    ShowCustomMessageBox($"Your slide deck at '{SlideDeck}' has been amended as per the AI's instruction. Check it out.")
                                Else
                                    ShowCustomMessageBox($"Your slide deck at '{SlideDeck}' has been amended as per the AI's instruction, but the file may show certain problems and may require a repair (internal error: {TokenErrorResponse}).")
                                End If
                            End If

                        Else
                            ShowCustomMessageBox($"There was a problem converting the AI response. You may want to retry.")
                        End If


                    ElseIf DoPane AndAlso Not DoSilent Then

                        If _uiContext IsNot Nothing Then  ' Make sure we run in the UI Thread
                            _uiContext.Post(Sub(s)
                                                SP_MergePrompt_Cached = SP_MergePrompt
                                                ShowPaneAsync(
                                        ClipPaneText1,
                                        LLMResult,
                                        PaneText2,
                                        AN,
                                        noRTF:=False,
                                        insertMarkdown:=True
                                        )
                                            End Sub, Nothing)
                        Else

                            SP_MergePrompt_Cached = SP_MergePrompt
                            ShowPaneAsync(ClipPaneText1, LLMResult, PaneText2, AN, noRTF:=False, insertMarkdown:=True)
                        End If

                    ElseIf DoNewDoc AndAlso Not DoSilent Then

                        Dim newDoc As Word.Document = Globals.ThisAddIn.Application.Documents.Add()
                        newDoc.Activate()
                        Dim newSelection As Word.Selection = Globals.ThisAddIn.Application.Selection
                        InsertTextWithMarkdown(newSelection, LLMResult, True, True)
                        Dim pattern As String = "\{\{(WFLD|WENT|WFNT):.*?\}\}"
                        rng = wordApp.Selection.Range
                        If Regex.IsMatch(LLMResult, pattern) Then
                            RestoreSpecialTextElements(rng)
                            rng.Document.Fields.Update()
                        End If

                    ElseIf PutInClipboard AndAlso Not DoSilent Then

                        Dim dialogResult As String = ""

                        If _uiContext IsNot Nothing Then
                            Dim doneEvent As New ManualResetEventSlim(False)            ' Make sure we run in the UI Thread

                            _uiContext.Post(Sub(state)
                                                Try

                                                    Dim wordHwnd As IntPtr = GetWordMainWindowHandle()

                                                    dialogResult = ShowCustomWindow(ClipPaneText1,
                                                                            LLMResult,
                                                                            ClipText2,
                                                                            AN,
                                                                            NoRTF:=False,
                                                                            Getfocus:=False,
                                                                            InsertMarkdown:=True,
                                                                            TransferToPane:=True,
                                                                            parentWindowHwnd:=wordHwnd)

                                                    If dialogResult <> "" And dialogResult <> "Pane" Then
                                                        If dialogResult = "Markdown" Then

                                                            Dim NewDocChoice As Integer = ShowCustomYesNoBox("Do you want to insert the text into a new Word document (if you cancel, it will be in the clipboard with formatting)?", "Yes, new", "No, into my existing doc")

                                                            If NewDocChoice = 1 Then
                                                                Dim newDoc As Word.Document = Globals.ThisAddIn.Application.Documents.Add()
                                                                Dim currentSelection As Word.Selection = newDoc.Application.Selection
                                                                currentSelection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                                                                InsertTextWithMarkdown(currentSelection, LLMResult, True, True)
                                                                Dim pattern As String = "\{\{(WFLD|WENT|WFNT):.*?\}\}"
                                                                If Regex.IsMatch(LLMResult, pattern) Then
                                                                    rng = currentSelection.Range
                                                                    RestoreSpecialTextElements(rng)
                                                                    rng.Document.Fields.Update()
                                                                End If
                                                            ElseIf NewDocChoice = 2 Then
                                                                Globals.ThisAddIn.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                                                                Globals.ThisAddIn.Application.Selection.TypeParagraph()
                                                                Globals.ThisAddIn.Application.Selection.TypeParagraph()
                                                                InsertTextWithMarkdown(Globals.ThisAddIn.Application.Selection, vbCrLf & LLMResult, False)
                                                                Dim pattern As String = "\{\{(WFLD|WENT|WFNT):.*?\}\}"
                                                                If Regex.IsMatch(LLMResult, pattern) Then
                                                                    rng = wordApp.Selection.Range
                                                                    RestoreSpecialTextElements(rng)
                                                                    rng.Document.Fields.Update()
                                                                End If
                                                            Else
                                                                ShowCustomMessageBox("No text was inserted (but included in the clipboard as RTF).")
                                                                SLib.PutInClipboard(MarkdownToRtfConverter.Convert((LLMResult)))
                                                            End If
                                                        Else
                                                            SLib.PutInClipboard(dialogResult)
                                                        End If
                                                    ElseIf dialogResult = "Pane" Then
                                                        SP_MergePrompt_Cached = SP_MergePrompt
                                                        ShowPaneAsync(
                                                                            ClipPaneText1,
                                                                            LLMResult,
                                                                            PaneText2,
                                                                            AN,
                                                                            noRTF:=False,
                                                                            insertMarkdown:=True
                                                                            )
                                                    End If

                                                Finally
                                                    doneEvent.Set()
                                                End Try
                                            End Sub, Nothing)


                        Else
                            dialogResult = ShowCustomWindow(
                                            ClipPaneText1,
                                            LLMResult,
                                            ClipText2,
                                            AN,
                                            NoRTF:=False,
                                            Getfocus:=False,
                                            InsertMarkdown:=True,
                                            TransferToPane:=True)

                            If dialogResult <> "" And dialogResult <> "Pane" Then
                                If dialogResult = "Markdown" Then

                                    Dim NewDocChoice As Integer = ShowCustomYesNoBox("Do you want to insert the text into a new Word document (if you cancel, it will be in the clipboard with formatting)?", "Yes, new", "No, into my existing doc")

                                    If NewDocChoice = 1 Then
                                        Dim newDoc As Word.Document = Globals.ThisAddIn.Application.Documents.Add()
                                        Dim currentSelection As Word.Selection = newDoc.Application.Selection
                                        currentSelection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                                        InsertTextWithMarkdown(currentSelection, LLMResult, True, True)
                                        Dim pattern As String = "\{\{(WFLD|WENT|WFNT):.*?\}\}"
                                        If Regex.IsMatch(LLMResult, pattern) Then
                                            rng = currentSelection.Range
                                            RestoreSpecialTextElements(rng)
                                            rng.Document.Fields.Update()
                                        End If
                                    ElseIf NewDocChoice = 2 Then
                                        Globals.ThisAddIn.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                                        Globals.ThisAddIn.Application.Selection.TypeParagraph()
                                        Globals.ThisAddIn.Application.Selection.TypeParagraph()
                                        InsertTextWithMarkdown(Globals.ThisAddIn.Application.Selection, LLMResult, False)
                                        Dim pattern As String = "\{\{(WFLD|WENT|WFNT):.*?\}\}"
                                        If Regex.IsMatch(LLMResult, pattern) Then
                                            rng = wordApp.Selection.Range
                                            RestoreSpecialTextElements(rng)
                                            rng.Document.Fields.Update()
                                        End If
                                    Else
                                        ShowCustomMessageBox("No text was inserted (but included in the clipboard as RTF).")
                                        SLib.PutInClipboard(MarkdownToRtfConverter.Convert((LLMResult)))
                                    End If
                                Else
                                    SLib.PutInClipboard(dialogResult)
                                End If
                            ElseIf dialogResult = "Pane" Then
                                SP_MergePrompt_Cached = SP_MergePrompt
                                ShowPaneAsync(
                                                    ClipPaneText1,
                                                    LLMResult,
                                                    PaneText2,
                                                    AN,
                                                    noRTF:=False,
                                                    insertMarkdown:=True
                                                    )
                            End If

                        End If

                    ElseIf PutInBubbles Then

                        'SetBubbles(LLMResult, selection, DoSilent)
                        Dim app As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application
                        If _uiContext IsNot Nothing Then
                            _uiContext.Post(
                                Sub(s)
                                    Dim uiSel As Microsoft.Office.Interop.Word.Selection = app.Selection
                                    SetBubbles(LLMResult, uiSel, DoSilent)
                                End Sub, Nothing)
                        Else
                            ' Fallback – assume we are already on UI thread
                            Dim uiSel As Microsoft.Office.Interop.Word.Selection = app.Selection
                            SetBubbles(LLMResult, uiSel, DoSilent)
                        End If

                    ElseIf MarkupMethod = 4 Then

                        Dim RegexResult = Await LLM(SP_MarkupRegex, "<ORIGINALTEXT>" & SelectedText & "</ORIGINALTEXT> /n <NEWTEXT>" & LLMResult & "</NEWTEXT>", "", "", 0, False)

                        MarkupSelectedTextWithRegex(RegexResult)

                        ' End Extended Selection Mode
                        Globals.ThisAddIn.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                    ElseIf NoSelectedText Then

                        InsertTextWithMarkdown(selection, vbCrLf & LLMResult, trailingCR, True)
                        Dim pattern As String = "\{\{(WFLD|WENT|WFNT):.*?\}\}"
                        If Regex.IsMatch(LLMResult, pattern) Then
                            rng = wordApp.Selection.Range
                            RestoreSpecialTextElements(rng)
                            rng.Document.Fields.Update()
                        End If

                        ' End Extended Selection Mode
                        Globals.ThisAddIn.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                    ElseIf KeepFormat AndAlso Not NoFormatting Then

                        SelectedText = selection.Text
                        SLib.InsertTextWithFormat(LLMResult, rng, InPlace)
                        If DoMarkup Then
                            LLMResult = SLib.RemoveHTML(LLMResult)
                            If MarkupMethod = 2 Or MarkupMethod = 3 Then
                                Dim SaveRng As Range = rng.Duplicate
                                CompareAndInsert(SelectedText, LLMResult, rng, MarkupMethod = 3, "This is the markup of the text inserted:", trailingCR)
                                If Not ParaFormatInline AndAlso Not NoFormatting AndAlso Not NoFormatAndFieldSaving Then
                                    ApplyParagraphFormat(rng)
                                End If
                                Dim pattern As String = "\{\{(WFLD|WENT|WFNT):.*?\}\}"
                                If Not NoFormatAndFieldSaving Or Regex.IsMatch(LLMResult, pattern) Then
                                    RestoreSpecialTextElements(SaveRng)
                                    SaveRng.Document.Fields.Update()
                                End If
                            Else
                                CompareAndInsertComparedoc(SelectedText, LLMResult, rng, ParaFormatInline, NoFormatting)
                                Dim pattern As String = "\{\{(WFLD|WENT|WFNT):.*?\}\}"
                                If Not NoFormatAndFieldSaving Or Regex.IsMatch(LLMResult, pattern) Then
                                    RestoreSpecialTextElements(rng)
                                    rng.Document.Fields.Update()
                                End If
                            End If
                        End If

                        ' End Extended Selection Mode
                        Globals.ThisAddIn.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                    Else

                        wordApp.ScreenUpdating = False

                        SelectedText = selection.Text

                        Debug.WriteLine($"7Range Start = {rng.Start} Selection Start = {selection.Start}")
                        Debug.WriteLine($"Range End = {rng.End} Selection End = {selection.End}")
                        Debug.WriteLine(vbCrLf & Left(rng.Text, 400) & vbCrLf)

                        If InPlace Then
                            If DoMarkup Then
                                If MarkupMethod = 2 Or MarkupMethod = 3 Then
                                    If MarkupMethod = 3 Then
                                        InsertTextWithMarkdown(selection, LLMResult, trailingCR)
                                        'If INI_MarkdownConvert Then LLMResult = RemoveMarkdownFormatting(LLMResult)
                                        rng = selection.Range
                                    Else
                                        If INI_MarkdownConvert Then LLMResult = RemoveMarkdownFormatting(LLMResult)
                                    End If
                                    Dim SaveRng As Range = rng.Duplicate
                                    CompareAndInsert(SelectedText, LLMResult, rng, MarkupMethod = 3, "This is the markup of the text inserted:", trailingCR)
                                    If Not ParaFormatInline AndAlso Not NoFormatting AndAlso Not NoFormatAndFieldSaving Then
                                        ApplyParagraphFormat(rng)
                                    End If
                                    Dim pattern As String = "\{\{(WFLD|WENT|WFNT):.*?\}\}"
                                    If Not NoFormatAndFieldSaving Or Regex.IsMatch(LLMResult, pattern) Then
                                        RestoreSpecialTextElements(SaveRng)
                                        SaveRng.Document.Fields.Update()
                                    End If
                                Else
                                    If INI_MarkdownConvert Then LLMResult = RemoveMarkdownFormatting(LLMResult)
                                    CompareAndInsertComparedoc(SelectedText, LLMResult, rng, ParaFormatInline, NoFormatting)
                                    Dim pattern As String = "\{\{(WFLD|WENT|WFNT):.*?\}\}"
                                    If Not NoFormatAndFieldSaving Or Regex.IsMatch(LLMResult, pattern) Then
                                        RestoreSpecialTextElements(rng)
                                        rng.Document.Fields.Update()
                                    End If

                                End If
                            Else
                                InsertTextWithMarkdown(selection, LLMResult, trailingCR)

                                Debug.WriteLine($"8Range Start = {rng.Start} Selection Start = {selection.Start}")
                                Debug.WriteLine($"Range End = {rng.End} Selection End = {selection.End}")
                                Debug.WriteLine(vbCrLf & Left(rng.Text, 400) & vbCrLf)

                                rng = selection.Range
                                Dim SaveRng As Range = rng.Duplicate
                                If Not ParaFormatInline AndAlso Not NoFormatting AndAlso Not NoFormatAndFieldSaving Then
                                    Debug.WriteLine($"9Range Start = {rng.Start} Selection Start = {selection.Start}")
                                    Debug.WriteLine($"Range End = {rng.End} Selection End = {selection.End}")
                                    Debug.WriteLine(vbCrLf & Left(rng.Text, 400) & vbCrLf)

                                    ApplyParagraphFormat(rng)
                                End If

                                Debug.WriteLine($"10Range Start = {rng.Start} Selection Start = {selection.Start}")
                                Debug.WriteLine($"Range End = {rng.End} Selection End = {selection.End}")

                                Debug.WriteLine($"SaveRange Start = {SaveRng.Start} Selection Start = {selection.Start}")
                                Debug.WriteLine($"SaveRange End = {SaveRng.End} Selection End = {selection.End}")

                                Dim pattern As String = "\{\{(WFLD|WENT|WFNT):.*?\}\}"
                                If Not NoFormatAndFieldSaving Or Regex.IsMatch(LLMResult, pattern) Then
                                    RestoreSpecialTextElements(SaveRng)
                                    SaveRng.Document.Fields.Update()
                                End If
                            End If

                        Else
                            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                            selection.TypeText(vbCrLf)
                            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                            rng = selection.Range
                            If DoMarkup Then
                                If MarkupMethod = 2 Or MarkupMethod = 3 Then
                                    Dim Pattern As String = ""
                                    If MarkupMethod = 3 Then
                                        Pattern = "\{\{.*?\}\}"
                                        If System.Text.RegularExpressions.Regex.IsMatch(LLMResult, Pattern) Then
                                            SLib.InsertTextWithBoldMarkers(selection, LLMResult)
                                            'If INI_MarkdownConvert Then LLMResult = RemoveMarkdownFormatting(LLMResult)
                                        Else
                                            InsertTextWithMarkdown(selection, LLMResult, trailingCR, True)
                                            'If INI_MarkdownConvert Then LLMResult = RemoveMarkdownFormatting(LLMResult)
                                        End If
                                        rng = selection.Range
                                    End If
                                    Dim SaveRng As Range = rng.Duplicate
                                    CompareAndInsert(SelectedText, LLMResult, rng.Duplicate, MarkupMethod = 3, "This is the markup of the text inserted:", trailingCR)
                                    If Not ParaFormatInline AndAlso Not NoFormatting AndAlso Not NoFormatAndFieldSaving Then
                                        ApplyParagraphFormat(rng)
                                    End If
                                    Pattern = "\{\{(WFLD|WENT|WFNT):.*?\}\}"
                                    If Not NoFormatAndFieldSaving Or Regex.IsMatch(LLMResult, Pattern) Then
                                        RestoreSpecialTextElements(SaveRng)
                                        SaveRng.Document.Fields.Update()
                                    End If
                                Else
                                    If INI_MarkdownConvert Then LLMResult = RemoveMarkdownFormatting(LLMResult)
                                    CompareAndInsertComparedoc(SelectedText, LLMResult, rng, ParaFormatInline, NoFormatting)
                                    Dim pattern As String = "\{\{(WFLD|WENT|WFNT):.*?\}\}"
                                    If Not NoFormatAndFieldSaving Or Regex.IsMatch(LLMResult, pattern) Then
                                        RestoreSpecialTextElements(rng)
                                        rng.Document.Fields.Update()
                                    End If
                                End If
                            Else
                                Dim pattern As String = "\{\{.*?\}\}"
                                If System.Text.RegularExpressions.Regex.IsMatch(LLMResult, pattern) Then
                                    SLib.InsertTextWithBoldMarkers(selection, LLMResult & vbCrLf)
                                    'If INI_MarkdownConvert Then LLMResult = RemoveMarkdownFormatting(LLMResult)
                                Else
                                    InsertTextWithMarkdown(selection, LLMResult, trailingCR, True)
                                    'If INI_MarkdownConvert Then LLMResult = RemoveMarkdownFormatting(LLMResult)
                                End If
                                rng = selection.Range
                                Dim SaveRng As Range = rng.Duplicate
                                If Not ParaFormatInline AndAlso Not NoFormatting AndAlso Not NoFormatAndFieldSaving Then
                                    ApplyParagraphFormat(rng)
                                End If
                                If Not NoFormatting Then
                                    pattern = "\{\{(WFLD|WENT|WFNT):.*?\}\}"
                                    If Not NoFormatAndFieldSaving Or Regex.IsMatch(LLMResult, pattern) Then
                                        RestoreSpecialTextElements(SaveRng)
                                        SaveRng.Document.Fields.Update()
                                    End If
                                End If
                            End If

                        End If

                        ' End Extended Selection Mode
                        Globals.ThisAddIn.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                        wordApp.ScreenUpdating = True

                    End If

                Else
                    If Not DoSilent Then ShowCustomMessageBox("The LLM did not return any content to process.")
                End If

                If NoSelectedText Or ChunkSize = 0 Then
                    Exit Do
                End If

                Try
                    If currentdoc.Bookmarks.Exists("NextStart") Then
                        Try
                            currentdoc.Bookmarks("NextStart").Delete()
                        Catch ex As System.Exception
                            '
                        End Try
                    End If

                    'nextStartBm = currentdoc.Bookmarks.Add(
                    'Name:="NextStart",
                    'Range:=currentdoc.Range(Start:=selection.End, End:=selection.End))

                    Dim totalEndPos As Integer = If(currentdoc.Bookmarks.Exists("TotalEnd"), currentdoc.Bookmarks("TotalEnd").Range.Start, selection.End)
                    Dim clampPos As Integer = System.Math.Min(selection.End, totalEndPos)
                    nextStartBm = currentdoc.Bookmarks.Add(
                            Name:="NextStart",
                            Range:=currentdoc.Range(Start:=clampPos, End:=clampPos))


                    nextStartBm.Range.Collapse(WdCollapseDirection.wdCollapseEnd)

                    If nextStartBm Is Nothing OrElse Not currentdoc.Bookmarks.Exists("NextStart") Then
                        Exit Do
                    End If

                Catch ex As System.Exception
                    Exit Do
                End Try

            Loop

        Catch ex As System.Exception

#If DEBUG Then
            Debug.WriteLine("Error: " & ex.Message)
            Debug.WriteLine("Stacktrace: " & ex.StackTrace)

            System.Diagnostics.Debugger.Break()
#End If
            MessageBox.Show("Error in TrueProcessSelectedText:  " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            INIloaded = False

        Finally

            Try
                ' Aufräumen aller temporären Bookmarks
                For yi As Integer = currentdoc.Bookmarks.Count To 1 Step -1
                    Dim bm As Word.Bookmark = currentdoc.Bookmarks(yi)
                    If bm.Name = "TotalEnd" OrElse bm.Name = "NextStart" Then bm.Delete()
                Next
            Catch ex As System.Exception
                '
            End Try

        End Try

        Return ""

    End Function


    ''' <summary>
    ''' Processes text in the active footnote or endnote story by sending either the current selection (when non-empty)
    ''' or the entire note text to the LLM, then replacing that same scope with the LLM result using the standard
    ''' Markdown-to-Word insertion pipeline.
    ''' </summary>
    ''' <param name="sysCommand">System prompt used for the LLM call.</param>
    ''' <param name="checkMaxToken">When true, estimates token count for the input text and warns if the configured output limit may be exceeded.</param>
    ''' <param name="useSecondApi">When true, routes the LLM call to the configured secondary API.</param>
    ''' <param name="selectionMandatory">When true, shows an error when the resolved input text is empty.</param>
    ''' <param name="fileObject">Optional file object parameter forwarded to the LLM call.</param>
    ''' <returns>Always returns an empty string; errors are reported via message boxes.</returns>
    ''' <remarks>
    ''' This routine is invoked when <see cref="Microsoft.Office.Interop.Word.Selection.StoryType"/> is
    ''' <see cref="Microsoft.Office.Interop.Word.WdStoryType.wdFootnotesStory"/> or
    ''' <see cref="Microsoft.Office.Interop.Word.WdStoryType.wdEndnotesStory"/>.
    ''' </remarks>
    Private Async Function ProcessSelectedTextInActiveFootnote(
    ByVal sysCommand As String,
    ByVal checkMaxToken As Boolean,
    ByVal useSecondApi As Boolean,
    ByVal selectionMandatory As Boolean,
    Optional ByVal fileObject As String = ""
) As Task(Of String)

        Try
            Dim app As Word.Application = Globals.ThisAddIn.Application
            If app Is Nothing OrElse app.Selection Is Nothing OrElse app.Selection.Range Is Nothing Then Return ""

            Dim sel As Word.Range = app.Selection.Range

            Dim doc As Word.Document = Nothing
            Try : doc = sel.Document : Catch : doc = Nothing : End Try
            If doc Is Nothing Then Return ""

            ' Must be inside footnote or endnote story.
            Dim st As Word.WdStoryType
            Try : st = sel.StoryType : Catch : Return "" : End Try

            Dim inFootnote As Boolean = (st = Word.WdStoryType.wdFootnotesStory)
            Dim inEndnote As Boolean = (st = Word.WdStoryType.wdEndnotesStory)

            If Not inFootnote AndAlso Not inEndnote Then
                ShowCustomMessageBox("Place the cursor inside a footnote/endnote (or select footnote text), then retry.")
                Return ""
            End If

            ' 1) Resolve owning note object
            Dim owningFootnote As Word.Footnote = Nothing
            Dim owningEndnote As Word.Endnote = Nothing

            If inFootnote Then
                Try
                    For Each fn As Word.Footnote In doc.Footnotes
                        Dim r As Word.Range = fn.Range
                        If sel.Start >= r.Start AndAlso sel.End <= r.End Then
                            owningFootnote = fn
                            Exit For
                        End If
                    Next
                Catch
                End Try
            ElseIf inEndnote Then
                Try
                    For Each en As Word.Endnote In doc.Endnotes
                        Dim r As Word.Range = en.Range
                        If sel.Start >= r.Start AndAlso sel.End <= r.End Then
                            owningEndnote = en
                            Exit For
                        End If
                    Next
                Catch
                End Try
            End If

            Dim noteRange As Word.Range = Nothing
            Try
                If owningFootnote IsNot Nothing Then noteRange = owningFootnote.Range.Duplicate
                If owningEndnote IsNot Nothing Then noteRange = owningEndnote.Range.Duplicate
            Catch
                noteRange = Nothing
            End Try

            If noteRange Is Nothing Then
                ShowCustomMessageBox("Cursor is in footnote/endnote story, but no owning note could be resolved.")
                Return ""
            End If

            ' 2) Decide whether the user actually has a meaningful selection
            Dim hasRealSelection As Boolean = False
            Try
                hasRealSelection = (sel.Start < sel.End AndAlso Not String.IsNullOrWhiteSpace(SafeRangeText(sel)))
            Catch
                hasRealSelection = False
            End Try

            ' 3) Input range: selection if meaningful, else whole note
            Dim inputRange As Word.Range = If(hasRealSelection, sel.Duplicate, noteRange.Duplicate)

            ' Best-effort: exclude trailing marker
            Try
                If inputRange.End > inputRange.Start Then inputRange.End -= 1
            Catch
            End Try

            Dim inputText As String = SafeRangeText(inputRange)
            If String.IsNullOrWhiteSpace(inputText) Then
                If selectionMandatory Then ShowCustomMessageBox("The footnote/endnote contains no text to process.")
                Return ""
            End If

            ' 4) Optional token warning
            If checkMaxToken Then
                Dim maxTok As Integer = If(useSecondApi, INI_MaxOutputToken_2, INI_MaxOutputToken)
                If maxTok > 0 Then
                    Dim est As Integer = EstimateTokenCount(inputText)
                    If est > maxTok Then
                        ShowCustomMessageBox($"Your footnote/endnote text may exceed the model output limits ({maxTok} tokens).")
                    End If
                End If
            End If

            ' 5) Call LLM
            Dim llmResult As String =
            Await LLM(
                promptSystem:=sysCommand,
                promptUser:="<TEXTTOPROCESS>" & inputText & "</TEXTTOPROCESS>",
                Model:="",
                Temperature:="",
                Timeout:=0,
                UseSecondAPI:=useSecondApi,
                Hidesplash:=False,
                AddUserPrompt:=OtherPrompt,
                FileObject:=fileObject
            )

            llmResult = llmResult.Replace("<TEXTTOPROCESS>", "").Replace("</TEXTTOPROCESS>", "")

            If Not String.IsNullOrEmpty(llmResult) Then
                llmResult = Await PostCorrection(llmResult, useSecondApi)
            End If

            If String.IsNullOrWhiteSpace(llmResult) Then
                ShowCustomMessageBox("The LLM did not return any content to process.")
                Return ""
            End If

            ' 6) Write back: selection or whole note
            Dim targetRange As Word.Range = If(hasRealSelection, sel.Duplicate, noteRange.Duplicate)

            Try
                If targetRange.End > targetRange.Start Then targetRange.End -= 1
            Catch
            End Try

            ' Clear + insert (footnotes accept formatted insertion like main doc text)
            Try
                targetRange.Text = ""
            Catch
            End Try

            ' Insert formatted content into footnote/endnote text (use the main document markdown pipeline)
            Dim trSel As Word.Selection = app.Selection
            trSel.SetRange(targetRange.Start, targetRange.End)
            trSel.Select()

            ' Reuse your main inserter (HTML/Markdown -> Word)
            InsertTextWithMarkdown(trSel, llmResult, TrailingCR:=False, AddTrailingIfNeeded:=True)

            Return ""

        Catch ex As Exception
            Debug.WriteLine($"ProcessSelectedTextInActiveFootnote error: {ex.Message}{vbCrLf}{ex.StackTrace}")
            ShowCustomMessageBox($"Error processing footnote/endnote: {ex.Message}")
            Return ""
        End Try
    End Function


    ''' <summary>
    ''' Captures revision-free text from a range by temporarily hiding revisions and filtering deleted content.
    ''' Returns text visible in final view, excluding deletions and moves.
    ''' </summary>
    ''' <param name="src">Word range to extract visible text from.</param>
    ''' <returns>Text with deletions omitted.</returns>
    Public Function GetVisibleText(ByVal src As Range) As String
        Try
            ' 1) Catch null/empty selection
            If src Is Nothing Then
                Return String.Empty
            End If

            ' 2) Fetch raw text once for fast path
            Dim raw As String
            Try
                raw = src.Text
                If String.IsNullOrEmpty(raw) Then
                    Return String.Empty
                End If
            Catch
                Return String.Empty
            End Try

            ' 3) Fast path: no revisions in this range → return immediately
            Dim revCount As Integer
            Try
                revCount = src.Revisions.Count
                If revCount = 0 Then
                    Return raw
                End If
            Catch
                Return raw ' If we can't access revisions count, return raw text
            End Try

            ' Alternative approach: use Word's built-in view settings
            Try
                Dim doc As Microsoft.Office.Interop.Word.Document = src.Document
                Dim origShowRevs As Boolean = doc.ShowRevisions
                Dim origPrintRevs As Boolean = doc.PrintRevisions

                ' Temporarily hide revisions to get clean text
                doc.ShowRevisions = False
                doc.PrintRevisions = False

                ' Get text without revisions
                Dim visibleText As String = src.Text

                ' Restore original settings
                doc.ShowRevisions = origShowRevs
                doc.PrintRevisions = origPrintRevs

                Return visibleText
            Catch ex As Exception
                Debug.WriteLine($"Alternative method failed: {ex.Message}")
                ' Continue with original algorithm if alternative fails
            End Try

            ' Original algorithm with better error handling
            Dim sliceStart As Integer = src.Start
            Dim sliceEnd As Integer = src.End    ' exclusive

            ' 4) Collect deleted intervals with safer revision handling
            Dim skips As New List(Of (s As Integer, e As Integer))()
            For Each rev As Revision In src.Revisions
                Try
                    ' Skip if we can't safely get revision type
                    Dim revType As WdRevisionType = rev.Type

                    If revType = WdRevisionType.wdRevisionInsert _
                OrElse revType = WdRevisionType.wdRevisionMovedTo Then
                        Continue For
                    End If

                    Dim revRange As Range = rev.Range
                    Dim fromPos As Integer = System.Math.Max(revRange.Start, sliceStart)
                    Dim toPos As Integer = System.Math.Min(revRange.End, sliceEnd)
                    If fromPos < toPos Then
                        skips.Add((fromPos, toPos))
                    End If
                Catch ex As Exception
                    ' Skip problematic revisions
                    Debug.WriteLine($"Error processing revision: {ex.Message}")
                    Continue For
                End Try
            Next

            ' 5) Merge intervals
            Dim merged As List(Of (s As Integer, e As Integer)) = MergeIntervals(skips)

            ' 6) Determine visible segments
            Dim keep As New List(Of (s As Integer, e As Integer))()
            Dim pos As Integer = sliceStart
            For Each iv In merged
                If iv.s > pos Then
                    keep.Add((pos, iv.s))
                End If
                pos = System.Math.Max(pos, iv.e)
            Next
            If pos < sliceEnd Then
                keep.Add((pos, sliceEnd))
            End If

            ' 7) Read text segment by segment
            Dim sb As New StringBuilder()
            For Each iv In keep
                Try
                    sb.Append(src.Document.Range(iv.s, iv.e).Text)
                Catch ex As Exception
                    Debug.WriteLine($"Error reading segment {iv.s}-{iv.e}: {ex.Message}")
                End Try
            Next

            Return sb.ToString()

        Catch ex As Exception
            Debug.WriteLine($"Exception in GetVisibleText: {ex.Message}{vbCrLf}{ex.StackTrace}")
            System.Diagnostics.Debugger.Break()
            ' Fall back to raw text or empty string in worst case
            Try
                Return If(src IsNot Nothing, src.Text, String.Empty)
            Catch
                Return String.Empty
            End Try
        End Try
    End Function

    ''' <summary>
    ''' Merges overlapping or adjacent integer intervals, returning a sorted list of non-overlapping ranges.
    ''' </summary>
    ''' <param name="intervals">List of start/end tuples to merge.</param>
    ''' <returns>Merged intervals.</returns>
    Private Function MergeIntervals(ByVal intervals As List(Of (s As Integer, e As Integer))) _
    As List(Of (s As Integer, e As Integer))

        Dim result As New List(Of (s As Integer, e As Integer))()
        If intervals.Count = 0 Then
            Return result
        End If

        intervals.Sort(Function(a, b) a.s.CompareTo(b.s))
        Dim cur = intervals(0)

        For i As Integer = 1 To intervals.Count - 1
            Dim nxt = intervals(i)
            If nxt.s <= cur.e Then
                cur.e = System.Math.Max(cur.e, nxt.e)
            Else
                result.Add(cur)
                cur = nxt
            End If
        Next

        result.Add(cur)
        Return result
    End Function


    ''' <summary>
    ''' Applies regex-based find/replace pairs from the LLM to the selected text with tracked changes.
    ''' Shows progress and handles cancellation; reports errors if patterns fail.
    ''' </summary>
    ''' <param name="regexResult">String containing pattern/replacement pairs from LLM.</param>
    Public Sub MarkupSelectedTextWithRegex(regexResult As String)
        Dim regexList As List(Of (Pattern As String, Replacement As String)) = ParseRegexString(regexResult)
        Dim errorCount As Integer = 0

        If regexList.Count = 0 Then
            ShowCustomMessageBox("The Regex markup method did Not work, As the the LLM delivered no valid regex patterns. You may want To retry.")
            Return
        End If

        Try
            Dim app As Word.Application = Globals.ThisAddIn.Application
            Dim selection As Microsoft.Office.Interop.Word.Selection = app.Selection

            If selection Is Nothing OrElse selection.Range Is Nothing Then
                MessageBox.Show("Error In MarkupSelectedTextWithRegex: No text selected (anymore). Can't proceed.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            'Dim splash As New Slib.Splashscreen("Applying changes... press 'Esc' to abort") 
            'splash.Show()
            'splash.Refresh()

            ShowProgressBarInSeparateThread($"{AN} Regex Markup", "Applying changes...")
            ProgressBarModule.CancelOperation = False

            ' Ensure Track Changes is enabled
            Dim originalTrackChangesSetting As Boolean = app.ActiveDocument.TrackRevisions
            Dim originalUserName As String = app.UserName
            app.ActiveDocument.TrackRevisions = True
            ' app.UserName = AN

            ' Define the character to be replaced
            Dim specialChar As String = ChrW(&HD83D)

            Dim selectedRange As Range = selection.Range
            Dim Exited As Boolean = False

            Dim regexIndex As Integer = 0

            For Each regexPair In regexList
                Try

                    System.Windows.Forms.Application.DoEvents()

                    If (GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0 Then
                        Exited = True
                        Exit For
                    End If

                    If (GetAsyncKeyState(VK_ESCAPE) And 1) <> 0 Then
                        Exited = True
                        Exit For
                    End If

                    selectedRange.Select()
                    SearchAndReplace(regexPair.Pattern, regexPair.Replacement, True, specialChar)

                    GlobalProgressMax = regexList.Count + 1

                    ' Update the current progress value and status label.
                    GlobalProgressValue = regexIndex + 1
                    GlobalProgressLabel = $"Search & Replace command {regexIndex + 1} of {regexList.Count}"

                    regexIndex += 1

                Catch ex As Exception
                    errorCount += 1
                End Try
            Next

            selectedRange.Select()

            If Not Exited Then

                GlobalProgressValue = regexIndex + 1
                GlobalProgressLabel = $"Cleaning up..."

                ' Loop through and replace occurrences of the character
                Dim replacementsMade As Boolean = False
                Do
                    With selectedRange.Find
                        .ClearFormatting()
                        .Text = specialChar
                        .Replacement.ClearFormatting()
                        .Replacement.Text = "" ' Replace with empty string
                        .Forward = True
                        .Wrap = Word.WdFindWrap.wdFindStop ' Do not loop around
                        If .Execute(Replace:=Word.WdReplace.wdReplaceOne) Then
                            replacementsMade = True
                        Else
                            Exit Do
                        End If
                    End With
                Loop
            End If

            ProgressBarModule.CancelOperation = True

            ' Restore original Track Changes setting
            app.ActiveDocument.TrackRevisions = originalTrackChangesSetting
            ' app.UserName = originalUserName

            'splash.Close()

            If errorCount > 0 Then
                ShowCustomMessageBox($"Some markups were applied. However, in {errorCount} cases this did not work because the LLM did not return the correct results. You may want to retry.")
            End If

        Catch ex As Exception
            MessageBox.Show($"Error in MarkupSelectedTextWithRegex: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Parses regex result string into pattern/replacement tuples. Splits by RegexSeparator2, then by RegexSeparator1.
    ''' Deduplicates identical pairs.
    ''' </summary>
    ''' <param name="input">Raw LLM regex output.</param>
    ''' <returns>List of pattern/replacement pairs.</returns>
    Private Function ParseRegexString(input As String) As List(Of (String, String))
        Dim result As New List(Of (String, String))
        Dim entries() As String = input.Split(New String() {RegexSeparator2}, StringSplitOptions.RemoveEmptyEntries)

        For Each entry In entries
            Dim parts() As String = entry.Split(New String() {RegexSeparator1}, StringSplitOptions.None)

            If parts.Length = 2 Then
                Dim key As String = parts(0).Trim()
                Dim value As String = parts(1).Trim()

                ' Only add if the tuple does not yet exist in result
                If Not result.Any(Function(item) item.Item1 = key AndAlso item.Item2 = value) Then
                    result.Add((key, value))
                End If
            End If
        Next

        Return result
    End Function

    ''' <summary>
    ''' Searches for oldText in chunks, replaces with newText (plus optional marker), respecting tracked changes and selection boundaries.
    ''' Advances search window after each replacement, restoring selection only if matches found.
    ''' </summary>
    ''' <param name="oldText">Text to find.</param>
    ''' <param name="newText">Replacement text.</param>
    ''' <param name="OnlySelection">Restrict to current selection.</param>
    ''' <param name="Marker">Optional character to embed in replacement.</param>
    Private Sub SearchAndReplace(oldText As String, newText As String, OnlySelection As Boolean, Marker As String)
        Dim doc As Microsoft.Office.Interop.Word.Document = Globals.ThisAddIn.Application.ActiveDocument
        Dim trackChangesEnabled As Boolean = doc.TrackRevisions
        Dim originalAuthor As String = doc.Application.UserName

        Try
            ' 1) Define working area
            Dim workRange As Microsoft.Office.Interop.Word.Range
            If OnlySelection AndAlso doc.Application.Selection IsNot Nothing _
           AndAlso doc.Application.Selection.Range.Text <> "" Then
                workRange = doc.Application.Selection.Range.Duplicate
            Else
                workRange = doc.Content.Duplicate
                OnlySelection = False
            End If

            ' 2) Insert marker into new text
            Dim newTextWithMarker As String
            If newText.Length > 2 AndAlso Marker <> "" Then
                newTextWithMarker = newText.Substring(0, newText.Length - 2) & Marker & newText.Substring(newText.Length - 2)
            Else
                newTextWithMarker = newText
            End If

            ' 3) Remember original selection
            Dim selectionStart As Integer = doc.Application.Selection.Start
            Dim selectionEnd As Integer = doc.Application.Selection.End

            ' 4) Long-chunk search
            doc.Application.Selection.SetRange(workRange.Start, workRange.End)
            Dim foundAny As Boolean = False

            Do While Globals.ThisAddIn.FindLongTextInChunks(oldText, doc.Application.Selection)
                If doc.Application.Selection Is Nothing Then Exit Do
                ' Check escape key
                If (GetAsyncKeyState(VK_ESCAPE) And 1) <> 0 Then Exit Do

                foundAny = True
                Dim selRange As Microsoft.Office.Interop.Word.Range = doc.Application.Selection.Range

                ' Check for deletions
                Dim isDeleted As Boolean = False
                For Each rev As Microsoft.Office.Interop.Word.Revision In selRange.Revisions
                    If rev.Type = Microsoft.Office.Interop.Word.WdRevisionType.wdRevisionDelete Then
                        isDeleted = True
                        Exit For
                    End If
                Next

                ' Replace
                Dim replaceStart As Integer = selRange.Start
                Dim replaceEnd As Integer = selRange.End

                If Not isDeleted Then
                    selRange.Text = newTextWithMarker

                    ' Advance replaceEnd and selectionEnd by inserted length
                    replaceEnd = replaceStart + newTextWithMarker.Length
                    selectionEnd += newTextWithMarker.Length

                    ' Hard clamp after mutation
                    selectionEnd = System.Math.Min(selectionEnd, doc.Content.End)
                    replaceEnd = System.Math.Min(replaceEnd, doc.Content.End)
                End If

                ' Clamping and range advance
                Dim newStart As Integer = System.Math.Max(doc.Content.Start, System.Math.Min(replaceEnd, doc.Content.End))

                ' Ensure the next search window is valid and non-negative length
                Dim desiredEnd As Integer = If(OnlySelection,
                                           System.Math.Min(selectionEnd, doc.Content.End),
                                           doc.Content.End)

                ' Guarantee monotonic forward progress: end >= start
                Dim newEnd As Integer = System.Math.Max(newStart, desiredEnd)

                ' Final clamps
                newStart = System.Math.Max(doc.Content.Start, System.Math.Min(newStart, doc.Content.End))
                newEnd = System.Math.Max(doc.Content.Start, System.Math.Min(newEnd, doc.Content.End))

                If newStart <= newEnd Then
                    doc.Application.Selection.SetRange(newStart, newEnd)
                Else
                    Exit Do
                End If
            Loop

            ' Restore selection only when matches found
            If foundAny Then
                selectionStart = System.Math.Max(doc.Content.Start, System.Math.Min(selectionStart, doc.Content.End))
                selectionEnd = System.Math.Max(doc.Content.Start, System.Math.Min(selectionEnd, doc.Content.End))
                If selectionStart <= selectionEnd Then
                    doc.Application.Selection.SetRange(selectionStart, selectionEnd)
                    doc.Application.Selection.Select()
                End If
            Else
                Debug.WriteLine("Note: Term not found, restore skipped.")
            End If

        Catch ex As System.Exception
            MsgBox("Error in SearchReplace: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub


    ''' <summary>
    ''' Converts Markdown text to Word formatting in the current selection, preserving original paragraph styles where possible.
    ''' </summary>
    Public Shared Sub ConvertMarkdownToWord()
        Dim app As Word.Application = Globals.ThisAddIn.Application
        Dim sel As Word.Selection = app.Selection

        ' Snapshot basic paragraph style + format for each paragraph in selection
        Dim srcParas As Word.Paragraphs = sel.Range.Paragraphs
        Dim savedStyles As New List(Of Object)()
        Dim savedFormats As New List(Of Word.ParagraphFormat)()

        For Each p As Word.Paragraph In srcParas
            Try
                savedStyles.Add(p.Range.Style)
            Catch
                savedStyles.Add(Nothing)
            End Try
            Try
                savedFormats.Add(p.Range.ParagraphFormat.Duplicate)
            Catch
                savedFormats.Add(Nothing)
            End Try
        Next

        ' Perform the conversion
        Dim selectedText As String = sel.Text
        Dim trailingCR As Boolean = (selectedText.EndsWith(vbCrLf) OrElse selectedText.EndsWith(vbLf) OrElse selectedText.EndsWith(vbCr))
        InsertTextWithMarkdown(sel, selectedText, trailingCR, True)

        ' Re-apply captured style + paragraph format (best-effort)
        Dim newParas As Word.Paragraphs = sel.Range.Paragraphs
        Dim applyCount As Integer = System.Math.Min(savedFormats.Count, newParas.Count)

        For i As Integer = 1 To applyCount
            Dim p As Word.Paragraph = newParas(i)
            Try
                If savedStyles(i - 1) IsNot Nothing Then
                    p.Range.Style = savedStyles(i - 1)
                End If
            Catch
            End Try
            Try
                If savedFormats(i - 1) IsNot Nothing Then
                    p.Range.ParagraphFormat = savedFormats(i - 1)
                End If
            Catch
            End Try
        Next
    End Sub

    ''' <summary>
    ''' Inserts Markdown-formatted text into the selection, converting it to HTML then Word formatting.
    ''' Handles trailing CR and space preservation around the inserted content.
    ''' </summary>
    ''' <param name="selection">Target Word selection.</param>
    ''' <param name="Result">Markdown text to insert.</param>
    ''' <param name="TrailingCR">Whether original text ended with CR.</param>
    ''' <param name="AddTrailingIfNeeded">Add trailing CR if needed.</param>
    Public Shared Sub InsertTextWithMarkdown(selection As Microsoft.Office.Interop.Word.Selection, Result As String, Optional TrailingCR As Boolean = False, Optional AddTrailingIfNeeded As Boolean = False)

        If selection Is Nothing Then
            MessageBox.Show("Error in InsertTextWithMarkdown: The selection object is null", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        ' Extract the range from the selection
        Dim range As Microsoft.Office.Interop.Word.Range = selection.Range

        Dim LeadingTrailingSpace As Boolean = False

        Debug.WriteLine($"IM1-Range Start = {selection.Start}")
        Debug.WriteLine($"Range End = {selection.End}")
        Debug.WriteLine("TrailingCR = " & TrailingCR)
        Debug.WriteLine(selection.Text)

        If range.Start < range.End AndAlso Not TrailingCR Then
            Dim docStart As Integer = range.Document.Content.Start
            Dim docEnd As Integer = range.Document.Content.End

            If range.Start > docStart AndAlso range.End < docEnd Then
                Dim beforerange As Range = range.Document.Range(range.Start - 1, range.Start)
                Dim afterrange As Range = range.Document.Range(range.End - 1, range.End)

                Debug.WriteLine($"Beforetext='{beforerange.Text}'")
                Debug.WriteLine($"Aftertext='{afterrange.Text}'")
                If afterrange.Text.EndsWith(" "c) OrElse afterrange.Text.StartsWith(" "c) Then
                    LeadingTrailingSpace = True
                Else
                    LeadingTrailingSpace = False
                End If
            Else
                LeadingTrailingSpace = False
            End If
        End If

        Dim insertionStart As Integer = selection.Range.Start

        Dim ResultBack As String = Result
        Try
            Result = System.Text.RegularExpressions.Regex.Unescape(Result)
        Catch
            Debug.WriteLine("Error unescaping Result with: " & Result)
            Result = ResultBack
        End Try

        Debug.WriteLine($"After Unescape: {Result}")

        Dim markdownSource As String = Result

        Result = Result.Replace(vbLf & " " & vbLf, vbLf & vbLf)

        Dim pattern As String = "((\r\n|\n|\r){2,})"
        Result = Regex.Replace(Result, pattern, Function(m As Match)
                                                    If m.Index + m.Length = Result.Length Then
                                                        Return m.Value
                                                    Else
                                                        Dim breaks As String = m.Value
                                                        Dim regexBreaks As New Regex("(\r\n|\n|\r)")
                                                        Dim splitBreaks = regexBreaks.Matches(breaks)
                                                        If splitBreaks.Count <= 1 Then Return breaks
                                                        Dim resultx As String = splitBreaks(0).Value
                                                        For i As Integer = 1 To splitBreaks.Count - 1
                                                            resultx &= vbCrLf & "&nbsp;" & vbCrLf & splitBreaks(i).Value
                                                        Next
                                                        Return resultx
                                                    End If
                                                End Function)

        ' ============= PROTECT SPECIAL PATTERNS BEFORE MARKDOWN CONVERSION =============
        ' Protect {{...}} placeholders (merge fields, special elements) - MUST come first
        Dim curlyPlaceholders As New List(Of String)()
        Result = System.Text.RegularExpressions.Regex.Replace(
            Result,
            "\{\{.*?\}\}",
            Function(m)
                curlyPlaceholders.Add(m.Value)
                Return $"[[CURLY_{curlyPlaceholders.Count - 1}]]"
            End Function,
            RegexOptions.Singleline)

        ' Protect single curly braces {word} that Markdig interprets as generic attributes
        Dim singleCurlyPlaceholders As New List(Of String)()
        Result = System.Text.RegularExpressions.Regex.Replace(
            Result,
            "\{[^{}]+\}",
            Function(m)
                singleCurlyPlaceholders.Add(m.Value)
                Return $"[[SCURLY_{singleCurlyPlaceholders.Count - 1}]]"
            End Function,
            RegexOptions.Singleline)

        ' Protect <...> angle bracket content that is NOT valid HTML
        Dim anglePlaceholders As New List(Of String)()
        Result = System.Text.RegularExpressions.Regex.Replace(
            Result,
            "<(?![/]?(?:p|br|div|span|strong|b|i|em|u|a|ul|ol|li|h[1-6]|table|tr|td|th|thead|tbody|img|hr|pre|code|blockquote)[ >/])[^>]+>",
            Function(m)
                anglePlaceholders.Add(m.Value)
                Return $"[[ANGLE_{anglePlaceholders.Count - 1}]]"
            End Function,
            RegexOptions.IgnoreCase Or RegexOptions.Singleline)
        ' ================================================================================

        Dim builder As New MarkdownPipelineBuilder()

        builder.UsePipeTables()
        builder.UseGridTables()
        builder.UseSoftlineBreakAsHardlineBreak()
        builder.UseListExtras()
        builder.UseFootnotes()
        builder.UseDefinitionLists()
        builder.UseAbbreviations()
        builder.UseAutoLinks()
        builder.UseTaskLists()
        builder.UseEmojiAndSmiley()
        builder.UseMathematics()
        builder.UseFigures()
        builder.UseAdvancedExtensions()
        builder.UseGenericAttributes()

        Dim markdownPipeline As MarkdownPipeline = builder.Build()

        Debug.WriteLine("Result=" & Result)

        Dim htmlResult As String = Markdown.ToHtml(Result, markdownPipeline).Trim

        ' ============= RESTORE PROTECTED PATTERNS AFTER MARKDOWN CONVERSION =============
        ' Restore {{...}} placeholders - DO NOT encode, these need to be processed by RestoreSpecialTextElements
        For idx As Integer = 0 To curlyPlaceholders.Count - 1
            htmlResult = htmlResult.Replace($"[[CURLY_{idx}]]", curlyPlaceholders(idx))
        Next

        ' Restore single curly braces {word} - keep as literal text
        For idx As Integer = 0 To singleCurlyPlaceholders.Count - 1
            htmlResult = htmlResult.Replace($"[[SCURLY_{idx}]]", System.Net.WebUtility.HtmlEncode(singleCurlyPlaceholders(idx)))
        Next

        ' Restore <...> placeholders (HTML-encode them so they display as text)
        For idx As Integer = 0 To anglePlaceholders.Count - 1
            htmlResult = htmlResult.Replace($"[[ANGLE_{idx}]]", System.Net.WebUtility.HtmlEncode(anglePlaceholders(idx)))
        Next
        ' ================================================================================

        ' Remove all real newlines so they are not converted as text
        htmlResult = htmlResult _
        .Replace(vbCrLf, "") _
        .Replace(vbCr, "") _
        .Replace(vbLf, "")

        ' Load the HTML into HtmlDocument
        Dim htmlDoc As New HtmlAgilityPack.HtmlDocument()
        Dim fullhtml As String
        htmlDoc.LoadHtml(htmlResult)

        fullhtml = htmlDoc.DocumentNode.OuterHtml
        Debug.WriteLine("HTML1=" & fullhtml)

        SLib.InsertTextWithFormat(fullhtml, range, True, Not TrailingCR)

        range = range.Application.Selection.Range

        Debug.WriteLine($"IM3-Range Start = {selection.Start}")
        Debug.WriteLine($"Range End = {selection.End}")
        Debug.WriteLine("LeadingTrailingSpace = " & LeadingTrailingSpace)
        Debug.WriteLine("TrailingCR = " & TrailingCR)
        Debug.WriteLine(selection.Text)

        If LeadingTrailingSpace Then
            range.Collapse(WdCollapseDirection.wdCollapseEnd)
            range.InsertAfter(" ")
        End If

        Dim InsertionEnd As Integer = range.End

        Dim doc As Microsoft.Office.Interop.Word.Document = selection.Document
        selection.SetRange(insertionStart, InsertionEnd)
        selection.Select()

        Debug.WriteLine($"IM4-Range Start = {selection.Start}")
        Debug.WriteLine($"Range End = {selection.End}")
        Debug.WriteLine("LeadingTrailingSpace = " & LeadingTrailingSpace)
        Debug.WriteLine("TrailingCR = " & TrailingCR)
        Debug.WriteLine(selection.Text)
    End Sub

    ''' <summary>
    ''' Lightweight structure holding revision metadata (start, end, text, type, author) for efficient processing without repeated COM calls.
    ''' </summary>
    Private Structure RevInfo
        Public Start As Integer
        Public EndPos As Integer  ' Using EndPos instead of End to avoid keyword conflict
        Public Text As String
        Public Type As WdRevisionType
        Public Author As String
    End Structure

    ''' <summary>
    ''' Encodes Word revisions as HTML-style markup tags (ins/del) for LLM consumption.
    ''' Optionally filters by author name. Shows splash screen progress during collection.
    ''' </summary>
    ''' <param name="rng">Range containing revisions.</param>
    ''' <param name="TPMarkupName">Optional author name filter.</param>
    ''' <returns>Text with del/ins tags.</returns>
    Public Function AddMarkupTags(ByVal rng As Range, Optional ByVal TPMarkupName As String = Nothing) As String

        Dim splash As New SLib.SplashScreen("Coding markups...  counting")
        splash.Show()
        splash.Refresh()

        ' Quick exit for ranges without revisions
        Dim revCount As Integer = 0
        Try
            revCount = rng.Revisions.Count
        Catch ex As Exception
            splash.Close()
            Return rng.Text
        End Try

        If revCount = 0 Then
            splash.Close()
            Return rng.Text
        End If

        ' Get range boundaries
        Dim rangeStart As Integer = rng.Start
        Dim rangeEnd As Integer = rng.End
        Dim resultBuilder As New StringBuilder(rng.Text.Length * 2)

        ' Create a collection to hold all revision data in memory
        Dim revInfos As New List(Of RevInfo)(revCount)

        ' Collect all revision data in a single pass to minimize COM calls
        For i As Integer = 1 To revCount
            splash.UpdateMessage($"Collecting markups... {revCount - i} left")
            Try

                Dim rev As Revision = rng.Revisions(i)

                Try
                    Dim revRange As Range = rev.Range

                    Dim revStart As Integer = revRange.Start
                    Dim revEnd As Integer = revRange.End

                    ' Only process revisions that overlap with our range
                    If revEnd > rangeStart AndAlso revStart < rangeEnd Then
                        Dim revText As String = revRange.Text
                        Dim revType As WdRevisionType = rev.Type
                        Dim revAuthor As String = rev.Author

                        ' Create a value type to store data efficiently
                        revInfos.Add(New RevInfo() With {
                        .Start = revStart,
                        .EndPos = revEnd,
                        .Text = revText,
                        .Type = revType,
                        .Author = revAuthor
                    })
                    End If
                Catch ex As Exception
                    ' Skip this revision and continue
                    Continue For
                End Try
            Catch ex As Exception
                Debug.WriteLine($"AddMarkupTags: ERROR with revision {i}: {ex.Message}")
                ' Skip and continue with next revision
                Continue For
            End Try

        Next

        ' Sort revisions by start position
        revInfos.Sort(Function(a, b) a.Start.CompareTo(b.Start))

        ' Process document with minimal COM access
        Dim currentPos As Integer = rangeStart
        Dim ii As Integer = 0

        For Each info In revInfos

            splash.UpdateMessage("Coding markups... " & revInfos.Count - ii & " left")
            ii = ii + 1

            ' Add text before this revision
            If info.Start > currentPos Then
                Try
                    Debug.WriteLine($"AddMarkupTags: Getting text before revision: {currentPos} to {info.Start}")
                    Dim beforeText As String = rng.Document.Range(currentPos, info.Start).Text
                    resultBuilder.Append(beforeText)
                Catch ex As Exception
                    Debug.WriteLine($"AddMarkupTags: Error getting text before revision: {ex.Message}")
                    ' If we can't get the text, just continue
                End Try
            End If

            ' Check if we should include markup
            Dim includeMarkup As Boolean = String.IsNullOrEmpty(TPMarkupName) OrElse
            String.Equals(info.Author, TPMarkupName, StringComparison.OrdinalIgnoreCase)

            ' Add revision text with markup
            If includeMarkup Then
                Select Case info.Type
                    Case WdRevisionType.wdRevisionDelete
                        resultBuilder.Append("<del>").Append(info.Text).Append("</del>")
                        Debug.WriteLine($"AddMarkupTags: Added delete markup: {info.Text.Length} chars")
                    Case WdRevisionType.wdRevisionInsert
                        resultBuilder.Append("<ins>").Append(info.Text).Append("</ins>")
                        Debug.WriteLine($"AddMarkupTags: Added insert markup: {info.Text.Length} chars")
                    Case Else
                        resultBuilder.Append(info.Text)
                End Select
            Else
                resultBuilder.Append(info.Text)
            End If

            ' Update position
            currentPos = info.EndPos
        Next

        ' Add any remaining text
        If currentPos < rangeEnd Then
            Try
                Dim tailText As String = rng.Document.Range(currentPos, rangeEnd).Text
                resultBuilder.Append(tailText)
            Catch ex As Exception
                ' If we can't get the remaining text, just return what we have
            End Try
        End If

        splash.Close()

        Return resultBuilder.ToString()
    End Function

    ''' <summary>
    ''' Strips del/ins markup tags from text using regex, returning plain text.
    ''' </summary>
    ''' <param name="text">Markup-tagged text.</param>
    ''' <returns>Text without tags.</returns>
    Public Function RemoveMarkupTags(text As String) As String
        ' Remove <del>, </del>, <ins>, and </ins> tags using regular expressions
        Dim result As String = System.Text.RegularExpressions.Regex.Replace(text, "<del>|</del>|<ins>|</ins>", String.Empty)
        Return result
    End Function

    ''' <summary>
    ''' Creates tracked-change comparison by using Word's CompareDocuments on temporary docs, then pastes result into target range.
    ''' Shows splash screen and minimizes temp doc windows.
    ''' </summary>
    ''' <param name="originalText">Original text.</param>
    ''' <param name="newText">Revised text.</param>
    ''' <param name="targetrange">Destination range for comparison output.</param>
    ''' <param name="paraformatinline">Whether inline paragraph format was used.</param>
    ''' <param name="noformatting">Skip formatting application.</param>
    Private Sub CompareAndInsertComparedoc(originalText As String, newText As String, targetrange As Range, Optional paraformatinline As Boolean = False, Optional noformatting As Boolean = True)

        Dim splash As New SLib.SplashScreen("Creating markup using the Word compare functionality (ignore any flickering and press 'No' if prompted) ...")
        splash.Show()
        splash.Refresh()

        Dim wordApp As Word.Application = Globals.ThisAddIn.Application
        Dim tempOriginalDoc As Word.Document = Nothing
        Dim tempNewDoc As Word.Document = Nothing
        Dim comparisonDoc As Word.Document = Nothing
        Dim originalAuthor As String = wordApp.UserName
        Dim originalScreenUpdating As Boolean = wordApp.ScreenUpdating
        Dim rng As Word.Range

        Try
            ' Disable screen updating to reduce flickers
            wordApp.ScreenUpdating = False

            ' Set the temporary author name to app
            ' wordApp.UserName = AN

            ' Create temporary documents for original and new text
            tempOriginalDoc = wordApp.Documents.Add
            tempNewDoc = wordApp.Documents.Add

            ' Minimize the windows of the temporary documents
            tempOriginalDoc.Windows(1).WindowState = Word.WdWindowState.wdWindowStateMinimize
            tempNewDoc.Windows(1).WindowState = Word.WdWindowState.wdWindowStateMinimize

            ' Insert original text into the first temporary document
            tempOriginalDoc.Content.Text = originalText

            ' Insert new text into the second temporary document
            tempNewDoc.Content.Text = newText

            ' Define the entire newly added text to be the range rng
            rng = tempNewDoc.Content
            If Not paraformatinline And Not noformatting Then
                ApplyParagraphFormat(rng)
            End If

            ' Perform the comparison
            comparisonDoc = wordApp.CompareDocuments(
                OriginalDocument:=tempOriginalDoc,
                RevisedDocument:=tempNewDoc,
                Destination:=WdCompareDestination.wdCompareDestinationNew,
                Granularity:=WdGranularity.wdGranularityWordLevel,
                CompareFormatting:=False,
                CompareCaseChanges:=False,
                CompareWhitespace:=False,
                CompareTables:=False,
                CompareHeaders:=False,
                CompareFootnotes:=False,
                CompareTextboxes:=False,
                CompareFields:=False,
                CompareComments:=False,
                CompareMoves:=False,
                RevisedAuthor:=Application.UserName
            )

            ' Copy the comparison document's content while keeping the original format
            comparisonDoc.Content.Copy()

            ' Insert the compared content at the specified range
            targetrange.Paste()

        Catch ex As System.Exception
            MessageBox.Show("Error in CompareAndInsertComparedoc: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' Restore screen updating
            wordApp.ScreenUpdating = originalScreenUpdating

            ' Restore the original author name
            'wordApp.UserName = originalAuthor

            ' Clean up temporary documents
            If tempOriginalDoc IsNot Nothing Then tempOriginalDoc.Close(SaveChanges:=False)
            If tempNewDoc IsNot Nothing Then tempNewDoc.Close(SaveChanges:=False)
            If comparisonDoc IsNot Nothing Then comparisonDoc.Close(SaveChanges:=False)

            splash.Close()

        End Try
    End Sub

    ''' <summary>
    ''' Generates inline diff markup using DiffPlex, converting word-by-word changes into [INS_START]/[DEL_START] tags.
    ''' Optionally shows result in a preview window or inserts directly into target range.
    ''' </summary>
    ''' <param name="text1">Original text.</param>
    ''' <param name="text2">Revised text.</param>
    ''' <param name="targetRange">Destination range.</param>
    ''' <param name="ShowInWindow">Display preview before inserting.</param>
    ''' <param name="TextforWindow">Window caption.</param>
    ''' <param name="paraformatinline">Inline paragraph formatting flag.</param>
    ''' <param name="noformatting">Skip formatting.</param>
    Private Sub CompareAndInsert(text1 As String, text2 As String, targetRange As Range, Optional ShowInWindow As Boolean = False, Optional TextforWindow As String = "A text with these changes will be inserted ('Esc' to abort):", Optional paraformatinline As Boolean = False, Optional noformatting As Boolean = True, Optional trailingCR As Boolean = False)
        Try

            Dim diffBuilder As New InlineDiffBuilder(New Differ())
            Dim sText As String = String.Empty

            Debug.WriteLine("A Text1 = " & text1)
            Debug.WriteLine("A Text2 = " & text2)

            If Not trailingCR Then text2 = text2.TrimEnd(vbCr, vbLf).TrimEnd(vbCr, vbLf)

            ' Pre-process the texts to replace line breaks with a unique marker
            text1 = text1.Replace(vbCrLf, " {vbCrLf} ").Replace(vbCr, " {vbCr} ").Replace(vbLf, " {vbLf} ")
            text2 = text2.Replace(vbCrLf, " {vbCrLf} ").Replace(vbCr, " {vbCr} ").Replace(vbLf, " {vbLf} ")

            ' Normalize the texts by removing extra spaces
            text1 = text1.Replace("  ", " ").Trim()
            text2 = text2.Replace("  ", " ").Trim()

            Debug.WriteLine("B Text1 = " & text1)
            Debug.WriteLine("B Text2 = " & text2)

            ' Split the texts into words and convert them into a line-by-line format
            ' In Worte splitten (ohne leere Einträge) und zeilenweise darstellen
            '--- 1) pull out all {{…}} fields into a list and replace them with placeholders:
            Dim mergefields As New List(Of String)
            text1 = System.Text.RegularExpressions.Regex.Replace(text1, "\{\{.*?\}\}",
                Function(m)
                    mergefields.Add(m.Value)
                    Return $"[[MF{mergefields.Count - 1}]]"
                End Function)
            text2 = System.Text.RegularExpressions.Regex.Replace(text2, "\{\{.*?\}\}",
                Function(m)
                    mergefields.Add(m.Value)
                    Return $"[[MF{mergefields.Count - 1}]]"
                End Function)

            ' Split the texts into words and convert them into a line-by-line format
            ' 3) In Worte splitten (ohne leere Einträge) und zeilenweise darstellen
            Dim words1 As String = String.Join(
              Environment.NewLine,
              text1.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
                    )
            Dim words2 As String = String.Join(
              Environment.NewLine,
              text2.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
                    )
            ' Generate word-based diff using DiffPlex
            Dim diffResult As DiffPaneModel = diffBuilder.BuildDiffModel(words1, words2)

            '--- 4) emit tags *per run* rather than per word:
            Dim prevType = ChangeType.Unchanged
            For i = 0 To diffResult.Lines.Count - 1
                Dim line = diffResult.Lines(i)
                Dim nextType = If(i < diffResult.Lines.Count - 1, diffResult.Lines(i + 1).Type, ChangeType.Unchanged)

                ' open tag when entering an Insert or Delete run
                If line.Type = ChangeType.Inserted AndAlso prevType <> ChangeType.Inserted Then
                    sText &= "[INS_START]"
                ElseIf line.Type = ChangeType.Deleted AndAlso prevType <> ChangeType.Deleted Then
                    sText &= "[DEL_START]"
                End If

                ' the word itself
                sText &= line.Text.Trim() & " "

                ' close tag when exiting a run
                If line.Type = ChangeType.Inserted AndAlso nextType <> ChangeType.Inserted Then
                    sText &= "[INS_END] "
                ElseIf line.Type = ChangeType.Deleted AndAlso nextType <> ChangeType.Deleted Then
                    sText &= "[DEL_END] "
                End If

                prevType = line.Type
            Next

            '--- 5) put your merge‑fields back in-place:
            For idx = 0 To mergefields.Count - 1
                sText = sText.Replace($"[[MF{idx}]]", mergefields(idx))
            Next

            Debug.WriteLine("1 = " & sText)

            ' Remove preceding and trailing spaces around placeholders
            sText = sText.Replace("{vbCr}", "{vbCrLf}")
            sText = sText.Replace("{vbLf}", "{vbCrLf}")
            sText = sText.Replace(" {vbCrLf} ", "{vbCrLf}")
            sText = sText.Replace(" {vbCrLf}", "{vbCrLf}")
            sText = sText.Replace("{vbCrLf} ", "{vbCrLf}")

            Debug.WriteLine("2 = " & sText)

            ' Remove instances of line breaks surrounded by [DEL_START] and [DEL_END]
            sText = sText.Replace("[DEL_START]{vbCrLf}[DEL_END] ", "")
            sText = sText.Replace("[DEL_START]{vbCrLf}{vbCrLf}[DEL_END] ", "")
            sText = sText.Replace("{vbCrLf}[DEL_END] ", "{vbCrLf}[DEL_END]")

            ' Include instances of line breaks surrounded by [INS_START] and [INS_END] without the [INS...] text
            sText = sText.Replace("[INS_START]{vbCrLf}[INS_END] ", "{vbCrLf}")
            sText = sText.Replace("[INS_START]{vbCrLf}{vbCrLf}[INS_END] ", "{vbCrLf}{vbCrLf}")
            sText = sText.Replace("{vbCrLf}[INS_END] ", "{vbCrLf}[INS_END]")

            ' Entferne alle überflüssigen Leerzeilen-Platzhalter am Ende

            Debug.WriteLine("3 = " & sText)

            sText = sText.Replace(vbCrLf, "").Replace(vbCr, "").Replace(vbLf, "")

            ' Replace placeholders with actual line breaks
            sText = sText.Replace("{vbCrLf}", vbCrLf)

            ' Adjust overlapping tags
            sText = sText.Replace("[DEL_END] [INS_START]", "[DEL_END][INS_START]")
            sText = sText.Replace("[INS_START][INS_END] ", "")
            'sText = RemoveInsDelTagsInPlaceholders(sText)

            ' Insert formatted text into the specified range
            If Not ShowInWindow Then
                Debug.WriteLine("Text with tags: " & vbCrLf & "'" & sText & "'" & vbCrLf & vbCrLf)
                ' Trim trailing line breaks that may have been added during diff processing
                sText = sText.TrimEnd(vbCr, vbLf).TrimEnd(vbCr, vbLf)
                InsertMarkupText(sText, targetRange)
            Else
                sText = Regex.Replace(sText, "\{\{.*?\}\}", String.Empty)

                Dim htmlContent As String = ConvertMarkupToRTF(TextforWindow & "\r\r" & sText)

                System.Threading.Tasks.Task.Run(Sub()
                                                    ShowRTFCustomMessageBox(htmlContent)
                                                End Sub)
            End If

        Catch ex As System.Exception
            MessageBox.Show("Error in CompareAndInsertText: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    ''' <summary>
    ''' Parses markup tags ([INS_START]/[DEL_START]) and applies them to Word with track changes, stripping merge fields from deleted runs.
    ''' Merges contiguous tags, inserts plain text, tracked insertions, and deletions, then cleans duplicate spaces in final view.
    ''' </summary>
    ''' <param name="inputText">Tagged markup string.</param>
    ''' <param name="targetRange">Destination Word range.</param>
    Public Sub InsertMarkupText(ByVal inputText As String, ByVal targetRange As Microsoft.Office.Interop.Word.Range)
        Dim wordApp As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application
        Dim doc As Microsoft.Office.Interop.Word.Document = wordApp.ActiveDocument

        Dim originalTrack As Boolean = doc.TrackRevisions
        Dim originalUpdate As Boolean = wordApp.ScreenUpdating

        ' Declare position variables BEFORE the Try so they are still valid in Finally:
        Dim docStart As Integer
        Dim startPos As Integer
        Dim endPosNoCR As Integer

        Try
            wordApp.ScreenUpdating = False
            doc.TrackRevisions = False


            ' A) Preserve the trailing ¶ so the next paragraph never joins in
            docStart = doc.Content.Start
            startPos = targetRange.Start
            endPosNoCR = targetRange.End

            If endPosNoCR > docStart Then
                Dim checkRange As Microsoft.Office.Interop.Word.Range =
                doc.Range(endPosNoCR - 1, endPosNoCR)
                If checkRange.Text = vbCr Then
                    endPosNoCR -= 1
                End If
            End If

            If endPosNoCR >= startPos Then
                doc.Range(startPos, endPosNoCR).Delete()
            End If

            targetRange.SetRange(startPos, startPos)

            ' Merge contiguous INS- and DEL-tags with only spaces in between
            Dim txt As String = inputText
            txt = RemoveMergeFormatFromBraces(txt)

            ' Strip merge-fields out of **closed** delete-runs:
            txt = System.Text.RegularExpressions.Regex.Replace(
            txt,
            "\[DEL_START\]([\s\S]*?)\[DEL_END\]",
            Function(m As System.Text.RegularExpressions.Match) As String
                Return "[DEL_START]" &
                       System.Text.RegularExpressions.Regex.Replace(
                           m.Groups(1).Value,
                           "\{\{(?:WFLD|WFNT|WENT|PFOR):.*?\}\}",
                           String.Empty
                       ) &
                       "[DEL_END]"
            End Function,
            System.Text.RegularExpressions.RegexOptions.Singleline
        )

            ' Strip merge-fields out of **open** delete-runs (no closing tag),
            ' but only if really no [DEL_END] follows:
            txt = System.Text.RegularExpressions.Regex.Replace(
            txt,
            "\[DEL_START\]((?:(?!\[DEL_END\]).)*)$",
            Function(m As System.Text.RegularExpressions.Match) As String
                Return "[DEL_START]" &
                       System.Text.RegularExpressions.Regex.Replace(
                           m.Groups(1).Value,
                           "\{\{(?:WFLD|WFNT|WENT|PFOR):.*?\}\}",
                           String.Empty
                       )
            End Function,
            System.Text.RegularExpressions.RegexOptions.Singleline
        )

            Debug.WriteLine("Stripped txt1 = " & txt)

            txt = System.Text.RegularExpressions.Regex.Replace(txt, "\[INS_END\](\s*)\[INS_START\]", "$1")
            txt = System.Text.RegularExpressions.Regex.Replace(txt, "\[DEL_END\](\s*)\[DEL_START\]", "$1")

            Debug.WriteLine("Stripped txt2 = " & txt)

            While txt.Length > 0
                System.Windows.Forms.Application.DoEvents()
                If (GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0 Then Exit While

                ' Locate next opening tag
                Dim insPos As Integer = txt.IndexOf("[INS_START]", StringComparison.Ordinal)
                Dim delPos As Integer = txt.IndexOf("[DEL_START]", StringComparison.Ordinal)

                Dim nextTagPos As Integer
                Dim tagType As String = Nothing
                If insPos = -1 AndAlso delPos = -1 Then
                    nextTagPos = -1
                ElseIf insPos = -1 OrElse (delPos <> -1 AndAlso delPos < insPos) Then
                    nextTagPos = delPos : tagType = "DEL"
                Else
                    nextTagPos = insPos : tagType = "INS"
                End If

                ' Plain text before the next tag
                If nextTagPos = -1 OrElse nextTagPos > 0 Then
                    Dim plain As String = If(nextTagPos = -1, txt, txt.Substring(0, nextTagPos))
                    If plain.Length > 0 Then
                        doc.TrackRevisions = False
                        targetRange.InsertAfter(plain)
                        targetRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
                    End If
                End If
                If nextTagPos = -1 Then Exit While

                If tagType = "INS" Then
                    ' INSERT block
                    txt = txt.Substring(nextTagPos + "[INS_START]".Length)
                    Dim endIns As Integer = txt.IndexOf("[INS_END]", StringComparison.Ordinal)
                    Dim insText As String = If(endIns = -1, txt, txt.Substring(0, endIns))
                    If endIns <> -1 Then txt = txt.Substring(endIns + "[INS_END]".Length)
                    doc.TrackRevisions = True
                    targetRange.InsertAfter(insText)
                    targetRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
                    doc.TrackRevisions = False

                Else
                    ' DELETION block
                    txt = txt.Substring(nextTagPos + "[DEL_START]".Length)
                    Dim endDel As Integer = txt.IndexOf("[DEL_END]", StringComparison.Ordinal)
                    Dim delText As String = If(endDel = -1, txt, txt.Substring(0, endDel))
                    If endDel <> -1 Then txt = txt.Substring(endDel + "[DEL_END]".Length)

                    ' Absorb following space/CR
                    If txt.StartsWith(" ") Then
                        delText &= " " : txt = txt.Substring(1)
                    ElseIf txt.StartsWith(vbCrLf) Then
                        delText &= vbCrLf : txt = txt.Substring(2)
                    ElseIf txt.StartsWith(vbCr) Then
                        delText &= vbCr : txt = txt.Substring(1)
                    End If

                    ' a) Insert (silent)
                    doc.TrackRevisions = False
                    targetRange.Text = delText
                    ' b) Delete (with tracking)
                    doc.TrackRevisions = True
                    targetRange.Delete()
                    doc.TrackRevisions = False
                    targetRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
                End If
            End While

        Catch ex As System.Exception
            Debug.WriteLine("InsertMarkupText error: " & ex.Message & vbCrLf & inputText)
        Finally

            ' Final-view replace test (space cleanup)

            ' 2) Activate final view
            With wordApp.ActiveWindow.View
                .RevisionsView = Microsoft.Office.Interop.Word.WdRevisionsView.wdRevisionsViewFinal
                .ShowRevisionsAndComments = False
            End With
            ' 3) Replace double spaces
            ' Temporarily turn off revisions so replacements are not logged as changes
            doc.TrackRevisions = False

            Dim endPosInserted2 As Integer = targetRange.End
            Dim insertedRange As Microsoft.Office.Interop.Word.Range =
            doc.Range(startPos, endPosInserted2)

            ' Find/Replace for two spaces → one space
            With insertedRange.Find
                .ClearFormatting()
                .Replacement.ClearFormatting()
                .Text = "  "    ' exactly two spaces
                .Replacement.Text = " "
                .Forward = True
                .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
                .Format = False
                .MatchWildcards = False
            End With

            ' As long as a replacement still occurs, repeat
            Do
                ' Execute returns True if something was replaced
            Loop While insertedRange.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)

            With wordApp.ActiveWindow.View
                .RevisionsView = Microsoft.Office.Interop.Word.WdRevisionsView.wdRevisionsViewFinal
                .ShowRevisionsAndComments = True
            End With

            ' Restore tracking to original state
            doc.TrackRevisions = originalTrack

            wordApp.ScreenUpdating = originalUpdate

            ' Set range to the full inserted length
            Dim endPosInserted As Integer = targetRange.End
            targetRange.SetRange(startPos, endPosInserted)
            wordApp.Selection.SetRange(targetRange.Start, targetRange.End)
        End Try
    End Sub


    ''' <summary>
    ''' Removes any "\* MERGEFORMAT" switch from inside {{…}} fields.
    ''' </summary>
    ''' <param name="input">Your full diff-markup string.</param>
    ''' <returns>The same string, but with MERGEFORMAT gone from inside all {{…}}.</returns>
    Function RemoveMergeFormatFromBraces(input As String) As String
        ' Process each {{…}} as one chunk
        Return System.Text.RegularExpressions.Regex.Replace(
        input,
        "\{\{(.*?)\}\}",
        Function(m As System.Text.RegularExpressions.Match) As String
            ' m.Groups(1).Value is the interior of the braces
            Dim inner As String = m.Groups(1).Value
            ' remove any \* MERGEFORMAT (case‑insensitive)
            inner = System.Text.RegularExpressions.Regex.Replace(
                inner,
                "\\\*\s*MERGEFORMAT",
                String.Empty,
                System.Text.RegularExpressions.RegexOptions.IgnoreCase
            )
            ' stitch it back together
            Return "{{" & inner & "}}"
        End Function,
        System.Text.RegularExpressions.RegexOptions.Singleline
    )
    End Function


End Class
