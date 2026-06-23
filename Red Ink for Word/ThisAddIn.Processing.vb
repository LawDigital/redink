' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.Processing.vb
' Purpose: Drives the Red Ink for Word text-processing pipeline, including selection
'          validation, chunk iteration, formatting capture/restoration, LLM invocation,
'          and routing AI output into Word, panes, clipboard, bubbles, podcasts, slides,
'          and chart/diagram outputs.
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
'  - Chart / Diagram Output: when `DoChart` is enabled, the LLM is prompted to return draw.io
'    XML (mxfile). The response is cleaned (`CleanDrawioXml`), saved as a `.drawio` file
'    (`ProcessChartResult`), and can be opened in an embedded diagrams.net editor
'    (`OpenDrawioWithWebView2` / `DrawioEditorForm`) for interactive editing and re-save.
'  - Story Routing: Detects non-main stories (comments, footnotes/endnotes) and routes processing to
'    specialized handlers (`ProcessSelectedTextInActiveCommentBubble`, `ProcessSelectedTextInActiveFootnote`).
'
' External Dependencies:
'  - Microsoft.Office.Interop.Word for document automation and story navigation.
'  - SharedLibrary.SharedMethods for UI, prompt construction, LLM access, formatting utilities.
'  - DiffPlex for diff building, Markdig + HtmlAgilityPack for Markdown-to-HTML conversion.
'  - DocumentFormat.OpenXml, MarkdownToRtfConverter, and additional helpers referenced via SharedLibrary.
'  - Microsoft.Web.WebView2 (via `DrawioEditorForm`) for embedded diagrams.net editor hosting.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.Runtime.InteropServices
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
    ' - MarkupMethod: Integer to indicate the markup method to be used: 1 = Word, 2 = Diff Surgical, 3 = DiffW, 4 = Regex, 5 = Diff Legacy
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
    ''' <param name="DoChart">Create a chart (0 = none, 1 = normal, 2 = webapp)</param>
    ''' <param name="DoShowModel">Include model name in the output</param>
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
     Optional SelectedTools As List(Of ModelConfig) = Nothing,
     Optional DoChart As Integer = 0,
     Optional DoShowModel As Boolean = False) As Task(Of String)


        Dim application As Word.Application = Globals.ThisAddIn.Application
        Dim selection As Microsoft.Office.Interop.Word.Selection = application.Selection

        If SysCommand = "" Then
            ShowCustomMessageBox("The (system-)prompt for the LLM is missing.")
            Return ""
        End If

        ' Route comment-bubble selections to a comment pipeline

        Dim activeComment As Word.Comment = Nothing
        If TryResolveActiveComment(selection, activeComment) Then
            Dim bubbleSel As Word.Range = Nothing

            Try
                If selection.StoryType = Word.WdStoryType.wdCommentsStory Then
                    bubbleSel = selection.Range.Duplicate
                ElseIf activeComment IsNot Nothing Then
                    bubbleSel = activeComment.Range.Duplicate
                End If
            Catch
                bubbleSel = Nothing
            End Try

            Return Await ProcessSelectedTextInActiveCommentBubble(
                    SysCommand,
                    CheckMaxToken,
                    bubbleSel,
                    UseSecondAPI:=UseSecondAPI,
                    SelectionMandatory:=SelectionMandatory,
                    FileObject:=FileObject,
                    DoMarkup:=DoMarkup,
                    MarkupMethod:=MarkupMethod
                )
        End If


        'If selection IsNot Nothing AndAlso selection.StoryType = Word.WdStoryType.wdCommentsStory Then
        'Dim bubbleSel As Word.Range = Nothing
        'Try
        'bubbleSel = selection.Range.Duplicate
        'Catch
        'bubbleSel = Nothing
        'End Try


        'Return Await ProcessSelectedTextInActiveCommentBubble(
        'SysCommand,
        'CheckMaxToken,
        'bubbleSel,
        'UseSecondAPI:=UseSecondAPI,
        'SelectionMandatory:=SelectionMandatory,
        'FileObject:=FileObject
        '                   )
        'End If

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

                    Dim Result = Await TrueProcessSelectedText(SysCommand, CheckMaxToken, KeepFormat, ParaFormatInline, InPlace, DoMarkup, MarkupMethod, PutInClipboard, PutInBubbles, SelectionMandatory, UseSecondAPI, FormattingCap, DoTPMarkup, TPMarkupname, CreatePodcast, FileObject, DoPane, ChunkSize, NoFormatAndFieldSaving, DoNewDoc, SlideDeck, AddDocs, DoMyStyle, DoBubblesExtract, False, DoPushback, SelectedTools, DoChart, DoShowModel)

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

                        Dim Result = Await TrueProcessSelectedText(SysCommand, CheckMaxToken, KeepFormat, ParaFormatInline, InPlace, DoMarkup, MarkupMethod, PutInClipboard, PutInBubbles, SelectionMandatory, UseSecondAPI, FormattingCap, DoTPMarkup, TPMarkupname, CreatePodcast, FileObject, DoPane, ChunkSize, NoFormatAndFieldSaving, DoNewDoc, SlideDeck, AddDocs, DoMyStyle, DoBubblesExtract, False, DoPushback, SelectedTools, DoChart, DoShowModel)

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
    ''' <param name="DoChart">Create a chart (0 = none, 1 = normal, 2 = webapp)</param>
    ''' <param name="DoShowModel">Include model name in the output</param>
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
       Optional SelectedTools As List(Of ModelConfig) = Nothing,
       Optional DoChart As Integer = 0,
       Optional DoShowModel As Boolean = False) As Task(Of String)


        Dim application As Word.Application = Globals.ThisAddIn.Application
        Dim selection As Microsoft.Office.Interop.Word.Selection = application.Selection
        Dim currentdoc As Word.Document = selection.Document

        Dim savedUiWindow As Word.Window = Nothing
        Dim savedUiSelection As Word.Range = Nothing

        Try
            savedUiWindow = application.ActiveWindow
            If application.Selection IsNot Nothing AndAlso application.Selection.Range IsNot Nothing Then
                savedUiSelection = application.Selection.Range.Duplicate
            End If
        Catch ex As Exception
            Debug.WriteLine($"Saving UI context failed: {ex.Message}")
        End Try


        ' ============= ENSURE WE'RE IN MAIN STORY WITHOUT CHANGING SELECTION =============
        Try
            If currentdoc IsNot Nothing AndAlso selection IsNot Nothing Then
                Dim currentStory As Word.WdStoryType = selection.StoryType

                ' Only act if we're NOT already in the main text story
                If currentStory <> Word.WdStoryType.wdMainTextStory Then
                    ' Force view back to print view to get out of special editing modes                    
                    currentdoc.ActiveWindow.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView

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
                    vbCrLf & "DoPushback=" & DoPushback &
                    vbCrLf & "DoChart=" & DoChart &
                    vbCrLf & "DoShowModel=" & DoShowModel
                )

        Try

            Dim SelectedText As String = ""
            Dim SerializedSelectedText As String = ""
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
                DoShowModel = False

                If DoMarkup Then
                    Select Case MarkupMethod
                        Case 1
                            Dim SilentMarkup As Integer = SLib.ShowCustomYesNoBox($"You have choosen both iterated processing and markups using Word compare. Iteration only works using the Diff method. Continue using Diff markup (the character cap will be ignored) or go without markups?", "Yes, Diff markups", "No, no markups")
                            If SilentMarkup = 1 Then
                                MarkupMethod = 2
                            ElseIf SilentMarkup = 2 Then
                                DoMarkup = False
                            Else
                                Return ""
                            End If
                        Case 3
                            Dim SilentMarkup As Integer = SLib.ShowCustomYesNoBox($"You have choosen both iterated processing and markups using DiffW compare. Iteration only works using the Diff method. Continue using normal Diff markup (the character cap will be ignored) or go without markups?", "Yes, Diff markups", "No, no markups")
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

                If Not ParaFormatInline AndAlso Not NoFormatting AndAlso Not NoSelectedText AndAlso Not NoFormatAndFieldSaving AndAlso Not (DoMarkup AndAlso MarkupMethod = 2) Then

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
                                SelectedText = GetVisibleText(rng)
                            End If
                        Else
                            If INI_MarkdownConvert AndAlso Not KeepFormat AndAlso (Not DoMarkup OrElse (MarkupMethod = 3 Or MarkupMethod = 5)) AndAlso InPlace Then

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
                            End If

                            ' ── Universal revision fix ──────────────────────────────
                            ' GetTextWithSpecialElementsInline reads rng.Text internally,
                            ' which includes deleted revision text. The placeholders it
                            ' embedded ({{WFLD:...}}, {{WFNT:...}}, **bold**, etc.) are
                            ' correct, but the base text between them contains ghost
                            ' deletions. Fix: re-read just the visible text from the
                            ' (now Markdown-converted) range and re-inject the placeholders.
                            If Not DoTPMarkup Then
                                Try
                                    Dim revCount As Integer = 0
                                    Try : revCount = rng.Revisions.Count : Catch : End Try
                                    If revCount > 0 Then
                                        ' Extract all placeholders from the processed SelectedText
                                        Dim placeholderTokens As New List(Of String)
                                        Dim stripped As String = System.Text.RegularExpressions.Regex.Replace(
                                            SelectedText, "\{\{(?:WFLD|WFNT|WENT|PFOR):[^}]*\}\}",
                                            Function(m)
                                                placeholderTokens.Add(m.Value)
                                                Return $"[[RVPH{placeholderTokens.Count - 1}]]"
                                            End Function)

                                        ' Get the visible (accepted) text from the same range
                                        ' — the range has already been Markdown-converted by
                                        '   GetTextWithSpecialElementsInline, so bold is now
                                        '   **bold** etc. in the live range too.
                                        Dim visText As String = GetVisibleText(rng)

                                        If Not String.IsNullOrWhiteSpace(visText) Then
                                            ' Re-inject placeholders into the visible text at
                                            ' their original relative positions. Because the
                                            ' visible text is shorter (no deleted chars), we
                                            ' can't use absolute offsets. Instead, use the
                                            ' surrounding context words as anchors.
                                            '
                                            ' Simplest viable approach: the visible text is the
                                            ' correct base. Placeholders mark positions of
                                            ' footnotes/fields/etc. that exist in the visible
                                            ' text too (as single reference chars). So just
                                            ' re-run GetTextWithSpecialElementsInline in Final
                                            ' view — the function will find the same footnotes/
                                            ' fields at their (now correct) visible-text offsets.

                                            Dim viewObj = currentdoc.ActiveWindow.View
                                            Dim savedRevView = viewObj.RevisionsView
                                            Dim savedShowRevs = viewObj.ShowRevisionsAndComments
                                            Dim savedSelStart = selection.Start
                                            Dim savedSelEnd = selection.End

                                            Try
                                                viewObj.RevisionsView = WdRevisionsView.wdRevisionsViewFinal
                                                viewObj.ShowRevisionsAndComments = False

                                                ' Re-run extraction in Final view — same range,
                                                ' same Markdown state, but rng.Text now excludes
                                                ' deleted revisions. Placeholder offsets will be
                                                ' correct against the visible character positions.
                                                SelectedText = GetTextWithSpecialElementsInline(
                                                    rng,
                                                    If(NoFormatting, False, ParaFormatInline),
                                                    False)  ' Markdown already applied, don't redo

                                                Debug.WriteLine($"RevisionFix: re-extracted SelectedText in Final view, len={SelectedText.Length}")
                                            Finally
                                                viewObj.RevisionsView = savedRevView
                                                viewObj.ShowRevisionsAndComments = savedShowRevs
                                                Try
                                                    selection.SetRange(savedSelStart, savedSelEnd)
                                                Catch
                                                End Try
                                            End Try
                                        End If
                                    End If
                                Catch ex As Exception
                                    Debug.WriteLine($"RevisionFix error: {ex.Message}")
                                End Try
                            End If
                            ' ── End revision fix ────────────────────────────────────
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

                    SerializedSelectedText = SelectedText

                    Debug.WriteLine($"4dRange Start = {rng.Start} Selection Start = {selection.Start}")
                    Debug.WriteLine($"Range End = {rng.End} Selection End = {selection.End}")
                    Debug.WriteLine("SelectedText: " & SelectedText)
                    Debug.WriteLine("SerializedSelectedText: " & SerializedSelectedText)

                    Dim MaxToken As Integer = If(UseSecondAPI, INI_MaxOutputToken_2, INI_MaxOutputToken)
                    Dim EstimatedTokens As Integer = EstimateTokenCount(SelectedText)

                    If CheckMaxToken And MaxToken > 0 AndAlso EstimatedTokens > MaxToken AndAlso (InPlace Or DoMarkup) AndAlso Not DoSilent Then
                        ShowCustomMessageBox("Your selected text Is larger than the maximum output your LLM can supposedly generate. Therefore, the output may be shorter than expected based On maximum tokens supported, which Is " & MaxToken & " tokens. Your input (With formatting information, As the Case may be) has an estimated To be " & EstimatedTokens & " tokens). Therefore, check whether the output Is complete.", AN, 15)
                    End If

                    If DoMarkup AndAlso (MarkupMethod = 2 OrElse MarkupMethod = 5) AndAlso Len(SelectedText) > INI_MarkupDiffCap AndAlso Not DoSilent Then
                        Dim MarkupChange As Integer = SLib.ShowCustomYesNoBox($"The selected text exceeds the defined cap For the Diff markup method at {INI_MarkupDiffCap} chars (your selection has {Len(SelectedText)} chars). {If(KeepFormat, "This may be because HTML codes have been inserted to keep the formatting (you can turn this off in the settings). ", "")}How Do you want To Continue?", "Use Diff In Window compare instead", If(MarkupMethod = 2, "Use Diff", "Use Diff Classic"))
                        Select Case MarkupChange
                            Case 1
                                MarkupMethod = 3
                            Case 2
                                ' Keep current method (2 or 5) — user chose to proceed
                            Case Else
                                Return ""
                        End Select
                    End If

                    If DoMarkup And MarkupMethod = 4 And Len(SelectedText) > INI_MarkupRegexCap AndAlso Not DoSilent Then
                        Dim MarkupChange As Integer = SLib.ShowCustomYesNoBox($"The selected text exceeds the defined cap For the Regex markup method at {INI_MarkupRegexCap} chars (your selection has {Len(SelectedText)} chars). {If(KeepFormat, "This may be because HTML codes have been inserted to keep the formatting (you can turn this off in the settings). ", "")}How Do you want To Continue?", "Use Word compare instead", "Use Regex")
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

                Dim MarkdownInstruction As String = ""
                If MarkupMethod = 2 Then MarkdownInstruction = " " & SP_Add_NoMarkdown

                ' Restore the user's original UI context before the long-running model call.
                ' The processing target is kept separately in rng / bookmarks.
                RestoreUiContext(application, savedUiWindow, savedUiSelection)

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
                        OtherPromptUnfilled, DoChart:=DoChart)
                Else

                    ' Other LLM calls w/o Tooling

                    If SelectedAlternateModels Is Nothing OrElse SelectedAlternateModels.Count = 0 OrElse DoMarkup OrElse PutInBubbles OrElse DoPushback OrElse SlideInsert <> "" OrElse DoChart > 0 Then

                        LLMResult = Await LLM(SysCommand & If(String.IsNullOrWhiteSpace(BubblesText), "", " " & SP_Add_BubblesExtract) & If(DoTPMarkup, " " & SP_Add_Revisions, "") & " " & If(SlideDeck = "", If(NoFormatting, "", If(KeepFormat, " " & SP_Add_KeepHTMLIntact, " " & SP_Add_KeepInlineIntact & MarkdownInstruction)), " " & SP_Add_Slides) & If(DoMyStyle, " " & MyStyleInsert, "") & If(DoChart > 0, " " & SP_Add_Chart & If(DoChart = 2, " " & SP_Add_Chart_App, ""), ""), If(NoSelectedText, If(AddDocs, " " & InsertDocs & " ", "") & SlideInsert, "<TEXTTOPROCESS>" & SelectedText & "</TEXTTOPROCESS>" & If(AddDocs, " " & InsertDocs & " ", "") & SlideInsert & " " & BubblesText), "", "", 0, UseSecondAPI, False, OtherPromptUnfilled, FileObject)

                    Else

                        For Each mc As ModelConfig In SelectedAlternateModels
                            Dim err As Boolean = False
                            ApplyModelConfig(_context, mc, err)

                            LLMResult += mc.ModelDescription & ":" & vbCrLf & vbCrLf & Await LLM(SysCommand & If(String.IsNullOrWhiteSpace(BubblesText), "", " " & SP_Add_BubblesExtract) & If(DoTPMarkup, " " & SP_Add_Revisions, "") & " " & If(SlideDeck = "", If(NoFormatting, "", If(KeepFormat, " " & SP_Add_KeepHTMLIntact, " " & SP_Add_KeepInlineIntact & MarkdownInstruction)), " " & SP_Add_Slides) & If(DoMyStyle, " " & MyStyleInsert, ""), If(NoSelectedText, If(AddDocs, " " & InsertDocs & " ", "") & SlideInsert, "<TEXTTOPROCESS>" & SelectedText & "</TEXTTOPROCESS>" & If(AddDocs, " " & InsertDocs & " ", "") & SlideInsert & " " & BubblesText), "", "", 0, UseSecondAPI, False, OtherPromptUnfilled, FileObject) & vbCrLf

                        Next

                    End If
                End If

                OtherPrompt = ""
                OtherPromptUnfilled = ""

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

                If Not String.IsNullOrEmpty(SelectedText) AndAlso Not String.IsNullOrEmpty(LLMResult) Then
                    LLMResult = ReprojectOriginalBreakCharacters(SelectedText, LLMResult)
                End If

                Debug.WriteLine($"LLMResult 3 = '{LLMResult}'")
                Debug.WriteLine($"TrailingCR = {trailingCR} Count = {trailingCRcount}")

                Debug.WriteLine($"6Range Start = {rng.Start} Selection Start = {selection.Start}")
                Debug.WriteLine($"Range End = {rng.End} Selection End = {selection.End}")
                Debug.WriteLine(vbCrLf & Left(rng.Text, 400) & vbCrLf)

                ' Re-bind Word's live Selection to the processing target before any writeback.
                ActivateProcessingContext(rng)
                selection = application.Selection
                rng = selection.Range

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

                    Dim ModelName As String
                    If UseSecondAPI Then
                        Dim isAlternateModel As Boolean = originalConfigLoaded AndAlso
                            Not String.IsNullOrWhiteSpace(LastAlternateModel) AndAlso
                            Not String.Equals(INI_Model_2, originalConfig.Model, StringComparison.OrdinalIgnoreCase)
                        If isAlternateModel Then
                            ' (b2) Alternate model temporarily applied to secondary slot
                            ModelName = LastAlternateModel
                        Else
                            ' (b1) Default secondary model
                            ModelName = INI_Model_2 & " (as secondary model)"
                        End If
                    Else
                        ' (a) Primary model
                        ModelName = INI_Model & " (as primary model)"
                    End If

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

                    ElseIf DoChart > 0 Then

                        ProcessChartResult(LLMResult)

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

                        If DoShowModel Then LLMResult = $"Model: {ModelName}" & vbCrLf & vbCrLf & LLMResult

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

                        If DoShowModel Then LLMResult = $"Model: {ModelName}" & vbCrLf & vbCrLf & LLMResult

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

                        If DoShowModel Then LLMResult = $"Model: {ModelName}" & vbCrLf & vbCrLf & LLMResult

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

                        If DoShowModel Then LLMResult = $"Model: {ModelName}" & vbCrLf & vbCrLf & LLMResult

                        Dim RegexResult = Await LLM(SP_MarkupRegex, "<ORIGINALTEXT>" & SelectedText & "</ORIGINALTEXT> /n <NEWTEXT>" & LLMResult & "</NEWTEXT>", "", "", 0, False)

                        MarkupSelectedTextWithRegex(RegexResult)

                        ' End Extended Selection Mode
                        Globals.ThisAddIn.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                    ElseIf NoSelectedText Then

                        If DoShowModel Then LLMResult = $"Model: {ModelName}" & vbCrLf & vbCrLf & LLMResult

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

                        If DoShowModel Then LLMResult = $"Model: {ModelName}" & vbCrLf & vbCrLf & LLMResult

                        ' ── use serialized text for diff-based markup methods; visible text otherwise ──
                        Dim revCountLate As Integer = 0
                        Dim visibleSelectedText As String = ""
                        Try : revCountLate = rng.Revisions.Count : Catch : End Try
                        If revCountLate > 0 Then
                            visibleSelectedText = GetVisibleText(rng)
                        Else
                            visibleSelectedText = selection.Text
                        End If

                        If DoMarkup AndAlso (MarkupMethod = 2 OrElse MarkupMethod = 3 OrElse MarkupMethod = 5) Then
                            SelectedText = SerializedSelectedText
                        Else
                            SelectedText = visibleSelectedText
                        End If

                        Debug.WriteLine("==== SelectedText used for markup compare ====")
                        Debug.WriteLine(SelectedText.Replace(" ", "·").Replace(vbCr, "\r").Replace(vbLf, "\n"))


                        SLib.InsertTextWithFormat(LLMResult, rng, InPlace)
                        If DoMarkup Then
                            LLMResult = SLib.RemoveHTML(LLMResult)
                            If MarkupMethod = 5 Or MarkupMethod = 3 Then
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
                            ElseIf MarkupMethod = 2 Then
                                Dim SaveRng As Range = rng.Duplicate
                                CompareAndInsertSurgical(SelectedText, LLMResult, rng, trailingCR)
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

                        If DoShowModel Then LLMResult = $"Model: {ModelName}" & vbCrLf & vbCrLf & LLMResult

                        wordApp.ScreenUpdating = False

                        ' ── use serialized text for diff-based markup methods; visible text otherwise ──
                        Dim revCountLate As Integer = 0
                        Dim visibleSelectedText As String = ""
                        Try : revCountLate = rng.Revisions.Count : Catch : End Try
                        If revCountLate > 0 Then
                            visibleSelectedText = GetVisibleText(rng)
                        Else
                            visibleSelectedText = selection.Text
                        End If

                        If DoMarkup AndAlso (MarkupMethod = 2 OrElse MarkupMethod = 3 OrElse MarkupMethod = 5) Then
                            SelectedText = SerializedSelectedText
                        Else
                            SelectedText = visibleSelectedText
                        End If

                        Debug.WriteLine("==== SelectedText used for markup compare ====")
                        Debug.WriteLine(SelectedText.Replace(" ", "·").Replace(vbCr, "\r").Replace(vbLf, "\n"))

                        Debug.WriteLine($"7Range Start = {rng.Start} Selection Start = {selection.Start}")
                        Debug.WriteLine($"Range End = {rng.End} Selection End = {selection.End}")
                        Debug.WriteLine(vbCrLf & Left(rng.Text, 400) & vbCrLf)

                        If InPlace Then
                            If DoMarkup Then
                                If MarkupMethod = 5 Or MarkupMethod = 3 Then
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
                                ElseIf MarkupMethod = 2 Then
                                    Dim SaveRng As Range = rng.Duplicate
                                    CompareAndInsertSurgical(SelectedText, LLMResult, rng, trailingCR)
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
                                If MarkupMethod = 5 Or MarkupMethod = 3 Then
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
                                ElseIf MarkupMethod = 2 Then
                                    Dim SaveRng As Range = rng.Duplicate
                                    CompareAndInsertSurgical(SelectedText, LLMResult, rng, trailingCR)
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
                AddUserPrompt:=OtherPromptUnfilled,
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
    ''' Returns the visible text of a range in Final view, excluding deleted revisions.
    ''' Falls back to raw <see cref="Range.Text"/> when there are no revisions
    ''' or when the Final-view extraction fails.
    ''' </summary>
    ''' <param name="src">Word range to extract visible text from.</param>
    ''' <returns>Text with deletions omitted and insertions preserved.</returns>

    Public Function GetVisibleText(ByVal src As Range) As String
        Try
            ' 1) Catch null/empty range
            If src Is Nothing Then Return String.Empty

            ' 2) Fetch raw text once for fast path
            Dim raw As String
            Try
                raw = src.Text
                If String.IsNullOrEmpty(raw) Then Return String.Empty
            Catch
                Return String.Empty
            End Try

            ' 3) Fast path: no revisions in this range
            Try
                If src.Revisions.Count = 0 Then Return raw
            Catch
                Return raw
            End Try

            ' 4) Use Selection + Final view to read accepted text.
            '    Switching RevisionsView changes the position space, so a stale Range
            '    object returns wrong text. The live Selection is re-evaluated by Word.            
            Try
                Dim app As Microsoft.Office.Interop.Word.Application = src.Application

                ' ── FIX: Use the window that belongs to the range's own document,
                '    NOT app.ActiveWindow. When multiple documents are open,
                '    ActiveWindow may point to a different document, causing the
                '    view switch and Select() call to operate on the wrong window.
                Dim srcDoc As Microsoft.Office.Interop.Word.Document = src.Document
                Dim docWindow As Microsoft.Office.Interop.Word.Window = srcDoc.ActiveWindow
                Dim view As Microsoft.Office.Interop.Word.View = docWindow.View

                ' Save current view state
                Dim origRevView As WdRevisionsView = view.RevisionsView
                Dim origShowRevs As Boolean = view.ShowRevisionsAndComments

                ' Save current selection (in the original view's position space)
                Dim origSelStart As Integer = app.Selection.Start
                Dim origSelEnd As Integer = app.Selection.End

                ' Remember which window was active so we can restore it
                Dim origActiveWindow As Microsoft.Office.Interop.Word.Window = app.ActiveWindow

                Try
                    ' Activate the range's document window before selecting
                    If Not Object.ReferenceEquals(origActiveWindow, docWindow) Then
                        docWindow.Activate()
                    End If

                    ' Select the target range so Word treats it as the live Selection
                    src.Select()

                    ' Switch to Final view — Selection.Text now returns accepted text
                    view.RevisionsView = WdRevisionsView.wdRevisionsViewFinal
                    view.ShowRevisionsAndComments = False

                    Dim finalText As String = app.Selection.Text
                    If finalText Is Nothing Then finalText = String.Empty
                    Return finalText
                Finally
                    ' Restore view FIRST (so position space matches the saved selection)
                    view.RevisionsView = origRevView
                    view.ShowRevisionsAndComments = origShowRevs

                    ' Restore the original selection
                    Try
                        app.Selection.SetRange(origSelStart, origSelEnd)
                    Catch
                    End Try

                    ' Restore the previously active window if we switched away
                    Try
                        If Not Object.ReferenceEquals(origActiveWindow, docWindow) Then
                            origActiveWindow.Activate()
                        End If
                    Catch
                    End Try
                End Try

            Catch ex As Exception
                Debug.WriteLine($"Selection/Final-view approach failed: {ex.Message}")
                Return raw
            End Try

        Catch ex As Exception
            Debug.WriteLine($"Exception in GetVisibleText: {ex.Message}{vbCrLf}{ex.StackTrace}")
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

            Using BeginMarkupAuthorScope(app)
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
            End Using

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

        ' Track the target document's TrackRevisions state so we can restore it
        Dim targetDoc As Word.Document = targetrange.Document
        Dim originalTrackRevisions As Boolean = targetDoc.TrackRevisions

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
                RevisedAuthor:=GetMarkupAuthorOrCurrent(wordApp)
            )

            ' Copy the comparison document's content while keeping the original format
            comparisonDoc.Content.Copy()

            ' Disable TrackRevisions on the target document BEFORE pasting.
            ' If TrackRevisions is on, Word will record the entire paste as
            ' a new tracked insertion ON TOP of the revisions already embedded
            ' in the comparison document — causing the changed parts to appear
            ' both as tracked changes and as plain (duplicated) text.
            targetDoc.TrackRevisions = False

            ' Insert the compared content at the specified range
            targetrange.Paste()

            ' Re-enable TrackRevisions so that the pasted revisions (from the
            ' comparison document) are visible and any subsequent user edits
            ' continue to be tracked.
            targetDoc.TrackRevisions = True

        Catch ex As System.Exception
            MessageBox.Show("Error in CompareAndInsertComparedoc: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' Restore screen updating
            wordApp.ScreenUpdating = originalScreenUpdating

            ' Restore the original TrackRevisions state
            targetDoc.TrackRevisions = originalTrackRevisions

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
    ''' Generates inline diff markup using shared diff logic, optionally showing the result in a preview window or inserting it into the target range.
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

            Debug.WriteLine("A Text1 = " & text1)
            Debug.WriteLine("A Text2 = " & text2)

            Dim sText As String =
                BuildInlineDiffMarkup(
                    text1,
                    text2,
                    trimTrailingLineBreaksOnRevised:=Not trailingCR)

            Debug.WriteLine("Diff markup = " & sText)

            ' Insert formatted text into the specified range
            If Not ShowInWindow Then
                Debug.WriteLine("Text with tags: " & vbCrLf & "'" & sText & "'" & vbCrLf & vbCrLf)

                ' Trim trailing line breaks that may have been added during diff processing
                sText = sText.TrimEnd(vbCr, vbLf).TrimEnd(vbCr, vbLf)

                InsertMarkupText(sText, targetRange)
            Else
                sText = Regex.Replace(sText, "\{\{.*?\}\}", String.Empty)

                Dim htmlContent As String = ConvertMarkupToRTF(TextforWindow & vbCrLf & vbCrLf & sText)

                System.Threading.Tasks.Task.Run(
                    Sub()
                        ShowRTFCustomMessageBox(htmlContent, RestoreWindow:=True)
                    End Sub)
            End If

        Catch ex As System.Exception
            MessageBox.Show("Error in CompareAndInsertText: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub



    ' ========================== Surgical Markup (MarkupMethod 2) ==========================

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
    Private Sub CompareAndInsertSurgical(text1 As String, text2 As String, targetRange As Range, Optional trailingCR As Boolean = False)
        Dim wordApp As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application
        Dim doc As Microsoft.Office.Interop.Word.Document = targetRange.Document

        Dim originalTrack As Boolean = doc.TrackRevisions
        Dim originalUpdate As Boolean = wordApp.ScreenUpdating

        Const maxFindTextLength As Integer = 200
        Const uiYieldIntervalMs As Integer = 40
        Dim manualLineBreak As String = ChrW(11).ToString()

        Try
            wordApp.ScreenUpdating = False
            doc.TrackRevisions = False

            Debug.WriteLine("SurgicalMarkup: text1 length=" & text1.Length & " text2 length=" & text2.Length)

            If Not trailingCR Then
                text2 = text2.TrimEnd(vbCr, vbLf).TrimEnd(vbCr, vbLf)
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
                          Replace(vbCr, " {vbCr} ").
                          Replace(vbLf, " {vbLf} ")
            text2 = text2.Replace(manualLineBreak, " {vbVt} ").
                          Replace(vbCrLf, " {vbCrLf} ").
                          Replace(vbCr, " {vbCr} ").
                          Replace(vbLf, " {vbLf} ")

            ' ======================================================================
            ' STEP 3: Word-level diff via DiffPlex
            ' ======================================================================
            Dim tokenList1 As List(Of String) = TokenizeDiffUnits(text1)
            Dim tokenList2 As List(Of String) = TokenizeDiffUnits(text2)

            Debug.WriteLine("==== Surgical raw text1 ====")
            Debug.WriteLine(DebugVisualizeToken(text1))
            Debug.WriteLine("==== Surgical raw text2 ====")
            Debug.WriteLine(DebugVisualizeToken(text2))
            DebugDumpTokenList("Surgical tokenList1", tokenList1)
            DebugDumpTokenList("Surgical tokenList2", tokenList2)

            Dim words1 As String = String.Join(Environment.NewLine, tokenList1)
            Dim words2 As String = String.Join(Environment.NewLine, tokenList2)

            Dim diffBuilder As New InlineDiffBuilder(New Differ())
            Dim diffResult As DiffPaneModel = diffBuilder.BuildDiffModel(words1, words2)

            ' ======================================================================
            ' STEP 4: Merge consecutive same-type tokens into runs.
            '         Line-break, placeholder, and whitespace tokens stay isolated.
            ' ======================================================================
            Dim runs As New List(Of DiffRun)(System.Math.Max(4, diffResult.Lines.Count))
            Dim currentRunType As ChangeType = ChangeType.Unchanged
            Dim currentRunWords As New List(Of String)
            Dim lastTokenText As String = Nothing

            For i As Integer = 0 To diffResult.Lines.Count - 1
                Dim line = diffResult.Lines(i)
                Dim wordText As String = If(line.Text, String.Empty)
                If wordText.Length = 0 Then Continue For

                Dim isLB As Boolean = IsDiffLineBreakToken(wordText)
                Dim isPlaceholderToken As Boolean = IsDiffPlaceholderToken(wordText)
                Dim isWhitespaceToken As Boolean = IsDiffWhitespaceToken(wordText)

                Dim prevWasLB As Boolean = (currentRunWords.Count > 0 AndAlso IsDiffLineBreakToken(lastTokenText))
                Dim prevWasPlaceholder As Boolean = (currentRunWords.Count > 0 AndAlso IsDiffPlaceholderToken(lastTokenText))
                Dim prevWasWhitespace As Boolean = (currentRunWords.Count > 0 AndAlso IsDiffWhitespaceToken(lastTokenText))

                If (line.Type <> currentRunType OrElse
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

                currentRunType = line.Type
                currentRunWords.Add(wordText)
                lastTokenText = wordText
            Next

            If currentRunWords.Count > 0 Then
                runs.Add(New DiffRun With {
                    .RunType = currentRunType,
                    .Text = String.Concat(currentRunWords)
                })
            End If

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
            ' STEP 7: Strip special placeholders from deleted runs
            ' ======================================================================
            For riStrip As Integer = 0 To runs.Count - 1
                If runs(riStrip).RunType = ChangeType.Deleted Then
                    Dim r As DiffRun = runs(riStrip)
                    r.Text = System.Text.RegularExpressions.Regex.Replace(
                        r.Text,
                        "\{\{(?:WFLD|WFNT|WENT|PFOR):.*?\}\}",
                        String.Empty)
                    runs(riStrip) = r
                End If
            Next

            Using BeginMarkupAuthorScope(wordApp)

                ' ======================================================================
                ' STEP 8: Walk runs and apply changes
                ' ======================================================================
                Dim cursor As Range = targetRange.Duplicate
                cursor.Collapse(WdCollapseDirection.wdCollapseStart)

                Dim rangeEnd As Integer = targetRange.End
                Dim ri As Integer = 0
                Dim uiYieldStopwatch As Stopwatch = Stopwatch.StartNew()

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

                    Dim runText As String = run.Text
                    Debug.WriteLine($"RUN {ri:000}: {run.RunType} cursor={cursor.Start} rangeEnd={rangeEnd} text='{DebugVisualizeToken(runText)}'")

                    Select Case run.RunType

                        Case ChangeType.Unchanged
                            Dim isLineBreakRun As Boolean = IsDiffLineBreakToken(runText)
                            Dim isWhitespaceRun As Boolean = IsDiffWhitespaceToken(runText)
                            Dim isSpecialPlaceholderRunCurrent As Boolean = IsSpecialPlaceholderRun(runText)

                            If isSpecialPlaceholderRunCurrent Then
                                Dim beforeCursor As Integer = cursor.Start

                                AdvanceCursorPastSpecialPlaceholder(doc, cursor, rangeEnd, runText)

                                Dim nextChar As String = "<eor>"
                                If cursor.Start < rangeEnd Then
                                    nextChar = DebugVisualizeToken(doc.Range(cursor.Start, System.Math.Min(cursor.Start + 1, rangeEnd)).Text)
                                End If
                                Debug.WriteLine($"    Placeholder advance: {beforeCursor} -> {cursor.Start}; nextChar='{nextChar}'")

                            ElseIf isLineBreakRun Then
                                Dim expectedBreaks As String = NormalizeLineBreakRunForWordRange(runText)
                                If expectedBreaks.Length = 0 Then expectedBreaks = vbCr

                                Dim pos As Integer = cursor.Start
                                Dim matched As Integer = 0

                                Do While matched < expectedBreaks.Length AndAlso pos < rangeEnd
                                    Dim peekRange As Range = doc.Range(pos, System.Math.Min(pos + 1, rangeEnd))
                                    Dim expectedChar As String = expectedBreaks.Substring(matched, 1)

                                    If peekRange.Text = expectedChar Then
                                        matched += 1
                                    End If
                                    pos += 1
                                Loop

                                cursor.SetRange(pos, pos)

                            ElseIf isWhitespaceRun Then
                                Dim expected As String = runText
                                Dim pos As Integer = cursor.Start
                                Dim matched As Integer = 0

                                Do While matched < expected.Length AndAlso pos < rangeEnd
                                    Dim peekRange As Range = doc.Range(pos, System.Math.Min(pos + 1, rangeEnd))
                                    Dim expectedChar As String = expected.Substring(matched, 1)

                                    If peekRange.Text = expectedChar Then
                                        matched += 1
                                        pos += 1
                                    Else
                                        Exit Do
                                    End If
                                Loop

                                cursor.SetRange(pos, pos)

                            Else
                                Dim searchRange As Range = doc.Range(cursor.Start, rangeEnd)
                                Dim runFindText As String = NormalizeForWordFind(runText)

                                Dim wasTruncated As Boolean = False
                                Dim found As Boolean = ExecuteWordFind(searchRange, runFindText, wasTruncated, maxFindTextLength)

                                If found Then
                                    If wasTruncated Then
                                        Dim fullLen As Integer = runFindText.Length
                                        Dim estimatedEnd As Integer = searchRange.Start + fullLen
                                        If estimatedEnd > rangeEnd Then estimatedEnd = rangeEnd
                                        cursor.SetRange(estimatedEnd, estimatedEnd)
                                    Else
                                        cursor.SetRange(searchRange.End, searchRange.End)
                                    End If
                                Else
                                    Debug.WriteLine($"SurgicalMarkup: Anchor not found, advancing by estimate. Run text='{Left(runText, 80)}'")
                                    Dim estimatedLen As Integer = runFindText.Length
                                    Dim newPos As Integer = System.Math.Min(cursor.Start + estimatedLen, rangeEnd)
                                    cursor.SetRange(newPos, newPos)
                                End If
                            End If

                            ri += 1

                        Case ChangeType.Deleted, ChangeType.Inserted
                            Dim clusterStartIndex As Integer = ri
                            Dim cluster As New List(Of DiffRun)

                            Do While ri < runs.Count AndAlso runs(ri).RunType <> ChangeType.Unchanged
                                If runs(ri).Text IsNot Nothing AndAlso runs(ri).Text.Length > 0 Then
                                    cluster.Add(runs(ri))
                                End If
                                ri += 1
                            Loop

                            If cluster.Count = 0 Then
                                Continue Do
                            End If

                            Dim hasDeleted As Boolean = cluster.Any(Function(x) x.RunType = ChangeType.Deleted)
                            Dim hasInserted As Boolean = cluster.Any(Function(x) x.RunType = ChangeType.Inserted)

                            Dim deletedCoreBuilder As New System.Text.StringBuilder()
                            Dim insertedCoreBuilder As New System.Text.StringBuilder()

                            For Each clusterRun As DiffRun In cluster
                                Select Case clusterRun.RunType
                                    Case ChangeType.Deleted
                                        deletedCoreBuilder.Append(clusterRun.Text)
                                    Case ChangeType.Inserted
                                        insertedCoreBuilder.Append(clusterRun.Text)
                                End Select
                            Next

                            Dim deletedCoreText As String = NormalizeForWordFind(deletedCoreBuilder.ToString())
                            Dim insertedCoreText As String = NormalizeForWordFind(insertedCoreBuilder.ToString())

                            ' Safety guard:
                            ' Never mutate a break-only replacement. This protects manual line breaks
                            ' (ChrW(11)), paragraph breaks, and bullet/list paragraphs from being
                            ' destroyed when the model returns vbLf instead of the original Word break.
                            If hasDeleted AndAlso hasInserted AndAlso
                               IsOnlyBreakCharacters(deletedCoreText) AndAlso
                               IsOnlyBreakCharacters(insertedCoreText) Then

                                Debug.WriteLine($"    Break-only replacement skipped; preserving original break(s). delete='{DebugVisualizeToken(deletedCoreText)}' insert='{DebugVisualizeToken(insertedCoreText)}'")

                                Dim advanceTo As Integer = System.Math.Min(cursor.Start + deletedCoreText.Length, rangeEnd)
                                cursor.SetRange(advanceTo, advanceTo)

                                Continue Do
                            End If

                            Dim trailingWhitespace As String = Nothing
                            Dim hasTrailingWhitespace As Boolean =
                                ri < runs.Count AndAlso
                                runs(ri).RunType = ChangeType.Unchanged AndAlso
                                IsDiffWhitespaceToken(runs(ri).Text)

                            If hasTrailingWhitespace Then
                                trailingWhitespace = runs(ri).Text
                            End If

                            Dim leadingWhitespace As String = Nothing
                            Dim hasLeadingWhitespace As Boolean =
                                clusterStartIndex > 0 AndAlso
                                runs(clusterStartIndex - 1).RunType = ChangeType.Unchanged AndAlso
                                IsDiffWhitespaceToken(runs(clusterStartIndex - 1).Text)

                            If hasLeadingWhitespace Then
                                leadingWhitespace = runs(clusterStartIndex - 1).Text
                            End If

                            Dim operationCandidates As New List(Of SurgicalOperationCandidate)

                            If hasDeleted AndAlso hasInserted Then
                                ' Replacement:
                                ' Prefer absorbing the following whitespace on both sides of the replacement.
                                ' This avoids Word swallowing the separator after a tracked word replacement,
                                ' but does not put spaces on both sides of the deleted revision.
                                If hasTrailingWhitespace Then
                                    operationCandidates.Add(New SurgicalOperationCandidate With {
                                        .DeleteText = deletedCoreText & trailingWhitespace,
                                        .InsertText = insertedCoreText & trailingWhitespace,
                                        .SearchStart = cursor.Start,
                                        .ConsumeTrailingWhitespaceRun = True
                                    })
                                End If

                                operationCandidates.Add(New SurgicalOperationCandidate With {
                                    .DeleteText = deletedCoreText,
                                    .InsertText = insertedCoreText,
                                    .SearchStart = cursor.Start,
                                    .ConsumeTrailingWhitespaceRun = False
                                })

                            ElseIf hasDeleted Then
                                ' Pure deletion:
                                ' Only absorb whitespace when deleting word-like content.
                                ' Do not absorb whitespace for punctuation-only deletions, e.g. "," -> "",
                                ' because that would turn "FlyHigh, gemäss" into "FlyHighgemäss".
                                If ShouldAbsorbWhitespaceForPureDeletion(deletedCoreText) AndAlso hasTrailingWhitespace Then
                                    operationCandidates.Add(New SurgicalOperationCandidate With {
                                        .DeleteText = deletedCoreText & trailingWhitespace,
                                        .InsertText = String.Empty,
                                        .SearchStart = cursor.Start,
                                        .ConsumeTrailingWhitespaceRun = True
                                    })
                                End If

                                operationCandidates.Add(New SurgicalOperationCandidate With {
                                    .DeleteText = deletedCoreText,
                                    .InsertText = String.Empty,
                                    .SearchStart = cursor.Start,
                                    .ConsumeTrailingWhitespaceRun = False
                                })

                                If ShouldAbsorbWhitespaceForPureDeletion(deletedCoreText) AndAlso hasLeadingWhitespace Then
                                    operationCandidates.Add(New SurgicalOperationCandidate With {
                                        .DeleteText = leadingWhitespace & deletedCoreText,
                                        .InsertText = String.Empty,
                                        .SearchStart = System.Math.Max(targetRange.Start, cursor.Start - leadingWhitespace.Length),
                                        .ConsumeTrailingWhitespaceRun = False
                                    })
                                End If

                            ElseIf hasInserted Then
                                operationCandidates.Add(New SurgicalOperationCandidate With {
                                    .DeleteText = String.Empty,
                                    .InsertText = insertedCoreText,
                                    .SearchStart = cursor.Start,
                                    .ConsumeTrailingWhitespaceRun = False
                                })
                            End If

                            Debug.WriteLine($"    Cluster delete-core='{DebugVisualizeToken(deletedCoreText)}' insert-core='{DebugVisualizeToken(insertedCoreText)}'")

                            Dim operationSucceeded As Boolean = False
                            Dim appliedOperation As SurgicalOperationCandidate = Nothing

                            For Each op As SurgicalOperationCandidate In operationCandidates
                                Debug.WriteLine($"    Trying operation delete='{DebugVisualizeToken(op.DeleteText)}' insert='{DebugVisualizeToken(op.InsertText)}' from {op.SearchStart} to {rangeEnd}")

                                If String.IsNullOrEmpty(op.DeleteText) Then
                                    appliedOperation = op
                                    operationSucceeded = True
                                    Exit For
                                End If

                                Dim candidateRange As Range = doc.Range(op.SearchStart, rangeEnd)
                                Dim candidateWasTruncated As Boolean = False

                                If ExecuteWordFind(candidateRange, op.DeleteText, candidateWasTruncated, maxFindTextLength) Then
                                    If candidateWasTruncated Then
                                        Dim fullEnd As Integer = candidateRange.Start + op.DeleteText.Length
                                        If fullEnd > rangeEnd Then fullEnd = rangeEnd
                                        candidateRange.SetRange(candidateRange.Start, fullEnd)
                                    End If

                                    Dim deleteStart As Integer = candidateRange.Start

                                    doc.TrackRevisions = True
                                    candidateRange.Delete()
                                    doc.TrackRevisions = False

                                    cursor.SetRange(deleteStart, deleteStart)
                                    rangeEnd = targetRange.End

                                    appliedOperation = op
                                    operationSucceeded = True
                                    Exit For
                                End If
                            Next

                            If Not operationSucceeded Then
                                Debug.WriteLine($"SurgicalMarkup: Cluster operation not applied. Delete='{DebugVisualizeToken(deletedCoreText)}' Insert='{DebugVisualizeToken(insertedCoreText)}'")
                                Continue Do
                            End If

                            If appliedOperation.ConsumeTrailingWhitespaceRun Then
                                ri += 1
                            End If

                            If Not String.IsNullOrEmpty(appliedOperation.InsertText) Then
                                cursor.Collapse(WdCollapseDirection.wdCollapseStart)

                                doc.TrackRevisions = True
                                cursor.InsertAfter(appliedOperation.InsertText)
                                doc.TrackRevisions = False

                                cursor.SetRange(cursor.End, cursor.End)
                                rangeEnd = targetRange.End
                            End If

                        Case Else
                            Debug.WriteLine($"SurgicalMarkup: Unexpected ChangeType '{run.RunType}' at run {ri}")
                            ri += 1

                    End Select
                Loop

                doc.TrackRevisions = False
                CollapseDoubleSpacesInFinalView(doc, targetRange.Start, targetRange.End)
            End Using

        Catch ex As System.Exception
            Debug.WriteLine("SurgicalMarkup error: " & ex.Message & vbCrLf & ex.StackTrace)
        Finally
            doc.TrackRevisions = originalTrack
            wordApp.ScreenUpdating = originalUpdate

            wordApp.Selection.SetRange(targetRange.Start, targetRange.End)
        End Try
    End Sub

    Public Sub ApplySurgicalReplacement(originalText As String, newText As String, targetRange As Range, Optional trailingCR As Boolean = False)
        CompareAndInsertSurgical(originalText, newText, targetRange, trailingCR)
    End Sub

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
        Dim tokenByPlaceholder As New Dictionary(Of String, String)(StringComparer.Ordinal)

        Dim nextIndex As Integer = 0

        Dim replacePlaceholder As MatchEvaluator =
            Function(m As Match) As String
                Dim placeholder As String = m.Value
                Dim token As String = Nothing

                If Not tokenByPlaceholder.TryGetValue(placeholder, token) Then
                    token = $"[[MF{nextIndex}]]"
                    tokenByPlaceholder.Add(placeholder, token)
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


    ''' <summary>
    ''' Parses markup tags ([INS_START]/[DEL_START]) and applies them to Word with track changes, stripping merge fields from deleted runs.
    ''' Merges contiguous tags, inserts plain text, tracked insertions, and deletions, then cleans duplicate spaces in final view.
    ''' </summary>
    ''' <param name="inputText">Tagged markup string.</param>
    ''' <param name="targetRange">Destination Word range.</param>
    Public Sub InsertMarkupText(ByVal inputText As String, ByVal targetRange As Microsoft.Office.Interop.Word.Range)
        Dim wordApp As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application
        Dim doc As Microsoft.Office.Interop.Word.Document = targetRange.Document

        Dim originalTrack As Boolean = doc.TrackRevisions
        Dim originalUpdate As Boolean = wordApp.ScreenUpdating

        ' Declare position variables BEFORE the Try so they are still valid in Finally:
        Dim docStart As Integer
        Dim startPos As Integer
        Dim endPosNoCR As Integer

        Try

            wordApp.ScreenUpdating = False
            doc.TrackRevisions = False

            Using BeginMarkupAuthorScope(wordApp)

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

            End Using

        Catch ex As System.Exception
            Debug.WriteLine("InsertMarkupText error: " & ex.Message & vbCrLf & inputText)
        Finally

            ' Final-view replace test (space cleanup)

            ' 2) Activate final view            
            With doc.ActiveWindow.View
                .RevisionsView = Microsoft.Office.Interop.Word.WdRevisionsView.wdRevisionsViewFinal
                .ShowRevisionsAndComments = False
            End With
            ' 3) Replace double spaces
            ' Temporarily turn off revisions so replacements are not logged as changes
            doc.TrackRevisions = False

            Dim endPosInserted2 As Integer = targetRange.End
            Dim insertedRange As Microsoft.Office.Interop.Word.Range =
            doc.Range(startPos, endPosInserted2)

            ' Find/Replace for runs of two or more spaces → one space
            With insertedRange.Find
                .ClearFormatting()
                .Replacement.ClearFormatting()
                .Text = "[ ]{2,}"
                .Replacement.Text = " "
                .Forward = True
                .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
                .Format = False
                .MatchWildcards = True
            End With

            insertedRange.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)

            With doc.ActiveWindow.View
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

    ' ========================== Chart Processing ==========================

    ' Add this helper enum somewhere in ThisAddIn (same file is fine)
    Private Enum DrawioOpenChoice
        None = 0
        OfflineAfterLoad = 1
        Online = 2
    End Enum

    Private Function AskHowToOpenDrawio(prompt As String, caption As String) As DrawioOpenChoice

        ' If local-only is enforced, do not offer the online option.
        If _context IsNot Nothing AndAlso _context.INI_ForceDrawioLocal Then
            Dim rLocalOnly As Integer = ShowCustomYesNoBox(
                bodyText:=prompt,
                button1Text:="Open draw.io (local processing after load)",
                button2Text:="Do not open",
                header:=caption,
                autoCloseSeconds:=Nothing,
                Defaulttext:="",
                extraButtonText:="",
                extraButtonAction:=Nothing,
                CloseAfterExtra:=True,
                nonModal:=False)

            If rLocalOnly = 1 Then Return DrawioOpenChoice.OfflineAfterLoad
            Return DrawioOpenChoice.None ' 0 abort or 2 do not open
        End If

        ' Default behavior: offer local + online + do-not-open.
        Dim chosen As DrawioOpenChoice = DrawioOpenChoice.None

        ' Two main buttons return 1/2.
        ' Third button is implemented via extraButtonText+extraButtonAction and we set chosen inside.
        Dim r As Integer = ShowCustomYesNoBox(
            bodyText:=prompt,
            button1Text:="Open draw.io (local processing after load)",
            button2Text:="Open draw.io (with internet access)",
            header:=caption,
            autoCloseSeconds:=Nothing,
            Defaulttext:="",
            extraButtonText:="Do not open",
            extraButtonAction:=Sub()
                                   chosen = DrawioOpenChoice.None
                               End Sub,
            CloseAfterExtra:=True,
            nonModal:=False)

        If r = 1 Then Return DrawioOpenChoice.OfflineAfterLoad
        If r = 2 Then Return DrawioOpenChoice.Online
        Return DrawioOpenChoice.None

    End Function

    ''' <summary>
    ''' Processes the LLM chart result by saving it as a draw.io file on the desktop
    ''' and optionally opening draw.io in a WebView2 editor window.
    ''' </summary>
    ''' <param name="xmlContent">The draw.io XML content from the LLM.</param>
    Private Sub ProcessChartResult(ByVal xmlContent As String)
        Try
            If String.IsNullOrWhiteSpace(xmlContent) Then
                ShowCustomMessageBox("The LLM did not return any chart content.")
                Return
            End If

            ' Extract + clean the XML (handles code fences AND extra text around the XML)
            Dim cleanedXml As String = ExtractAndCleanDrawioXml(xmlContent)
            If String.IsNullOrWhiteSpace(cleanedXml) Then
                ShowCustomMessageBox("The chart output did not contain valid draw.io XML.")
                Return
            End If

            ' Minimal validation (avoid writing obviously non-draw.io content)
            Dim looksLikeDrawio As Boolean =
                cleanedXml.IndexOf("<mxfile", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                cleanedXml.IndexOf("<mxGraphModel", StringComparison.OrdinalIgnoreCase) >= 0

            If Not looksLikeDrawio Then
                ShowCustomMessageBox("The chart output did not look like draw.io XML (missing <mxfile> / <mxGraphModel>).")
                Return
            End If

            ' Get the desktop path
            Dim desktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)

            ' Find the next available file number
            Dim fileCounter As Integer = 1
            Dim filePath As String
            Do
                filePath = IO.Path.Combine(desktopPath, $"AI_Chart_{fileCounter:D3}.drawio")
                If Not IO.File.Exists(filePath) Then Exit Do
                fileCounter += 1
            Loop

            ' Always save the XML content to the file first
            IO.File.WriteAllText(filePath, cleanedXml, System.Text.Encoding.UTF8)

            Dim choice As DrawioOpenChoice =
            AskHowToOpenDrawio(
                prompt:=$"The chart has been saved to your desktop as '{IO.Path.GetFileName(filePath)}'." & vbCrLf & vbCrLf &
                        "How do you want to continue?",
                caption:=$"{AN}")

            Select Case choice
                Case DrawioOpenChoice.OfflineAfterLoad
                    OpenDrawioWithWebView2(cleanedXml, filePath, disableInternetAfterLoad:=True)
                Case DrawioOpenChoice.Online
                    OpenDrawioWithWebView2(cleanedXml, filePath, disableInternetAfterLoad:=False)
                Case Else
                    Return
            End Select

        Catch ex As Exception
            Debug.WriteLine($"ProcessChartResult error: {ex.Message}{vbCrLf}{ex.StackTrace}")
            ShowCustomMessageBox($"Error processing chart result: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Opens draw.io in a WebView2-based editor form, loading the XML via postMessage.
    ''' </summary>
    ''' <param name="xmlContent">The draw.io XML content.</param>
    ''' <param name="saveFilePath">Path where the .drawio file is saved.</param>
    ''' <param name="disableInternetAfterLoad">When true, blocks http/https requests after initial load.</param>
    Private Sub OpenDrawioWithWebView2(ByVal xmlContent As String, ByVal saveFilePath As String, Optional ByVal disableInternetAfterLoad As Boolean = False)
        Try
            ' Ensure we're on the UI thread
            If _uiContext IsNot Nothing Then
                _uiContext.Post(
                    Sub(state)
                        ShowDrawioEditor(xmlContent, saveFilePath, disableInternetAfterLoad)
                    End Sub, Nothing)
            Else
                ShowDrawioEditor(xmlContent, saveFilePath, disableInternetAfterLoad)
            End If

        Catch ex As Exception
            Debug.WriteLine($"OpenDrawioWithWebView2 error: {ex.Message}")
            ShowCustomMessageBox(
                $"Could not open the diagram editor: {ex.Message}{vbCrLf}{vbCrLf}" &
                $"Your diagram has been saved to:{vbCrLf}{saveFilePath}")
        End Try
    End Sub

    ''' <summary>
    ''' Shows the draw.io editor form with WebView2.
    ''' </summary>
    Private Sub ShowDrawioEditor(ByVal xmlContent As String, ByVal saveFilePath As String, ByVal disableInternetAfterLoad As Boolean)
        Try
            Dim editorForm As New DrawioEditorForm(xmlContent, saveFilePath, disableInternetAfterLoad:=disableInternetAfterLoad)

            ' For Custom PDF export on form 
            'AddHandler editorForm.Shown, Sub()            
            '                                editorForm.ExportPdfToDevice()
            '                            End Sub

            editorForm.Show()
        Catch ex As Exception
            Debug.WriteLine($"ShowDrawioEditor error: {ex.Message}")
            ShowCustomMessageBox(
                $"Could not open the diagram editor: {ex.Message}{vbCrLf}{vbCrLf}" &
                $"Your diagram has been saved to:{vbCrLf}{saveFilePath}")
        End Try
    End Sub

    ''' <summary>
    ''' Opens an existing .drawio file using the application's drag & drop picker,
    ''' then asks whether to open draw.io offline-after-load, with internet, or not at all.
    ''' If the user cancels the file picker, offers to create a new empty diagram instead.
    ''' </summary>
    Public Sub OpenExistingDrawioFileForEditing()
        Try
            ' Configure DragDropForm for draw.io files
            Globals.ThisAddIn.DragDropFormFilter = "draw.io files (*.drawio)|*.drawio|All files (*.*)|*.*"
            Globals.ThisAddIn.DragDropFormLabel = "Drop a .drawio file here (or click Browse) to open it in the embedded draw.io editor." & vbCrLf & vbCrLf & "Close this dialog to create a new chart."
            Dim selectedPath As String = Nothing

            Using f As New DragDropForm(DragDropMode.FileOnly)
                If f.ShowDialog() <> DialogResult.OK Then
                    ' User cancelled the file picker - ask if they want to create a new empty diagram
                    Dim createNew As String = ShowCustomInputBox(
    "You did not select a file. Do you want to create a new empty diagram instead?" & vbCrLf & vbCrLf &
    "Enter a file name (or full path) for the new diagram (without extension), or leave empty to cancel:",
    $"{AN} - Create New Diagram",
    SimpleInput:=True,
    DefaultValue:="AI_Chart_New")

                    If String.IsNullOrWhiteSpace(createNew) Then
                        Return ' User cancelled or left empty
                    End If

                    Dim newFilePath As String
                    Dim userInput As String = createNew.Trim()

                    ' Check if user provided a full path (contains directory separator or drive letter)
                    If IO.Path.IsPathRooted(userInput) OrElse userInput.Contains(IO.Path.DirectorySeparatorChar) OrElse userInput.Contains(IO.Path.AltDirectorySeparatorChar) Then
                        ' User provided a path - use it directly
                        ' Ensure .drawio extension
                        If Not userInput.EndsWith(".drawio", StringComparison.OrdinalIgnoreCase) Then
                            userInput &= ".drawio"
                        End If

                        ' Validate the directory exists
                        Dim directory As String = IO.Path.GetDirectoryName(userInput)
                        If Not String.IsNullOrWhiteSpace(directory) AndAlso Not IO.Directory.Exists(directory) Then
                            Dim createDir As Integer = ShowCustomYesNoBox(
                                $"The directory '{directory}' does not exist. Create it?",
                                "Yes", "No",
                                $"{AN} - Create Directory")
                            If createDir = 1 Then
                                Try
                                    IO.Directory.CreateDirectory(directory)
                                Catch ex As Exception
                                    ShowCustomMessageBox($"Could not create directory: {ex.Message}")
                                    Return
                                End Try
                            Else
                                Return
                            End If
                        End If

                        newFilePath = userInput
                    Else
                        ' User provided just a file name - save to desktop
                        Dim sanitizedName As String = SanitizeFileName2(userInput)
                        If String.IsNullOrWhiteSpace(sanitizedName) Then
                            ShowCustomMessageBox("Invalid file name. Operation cancelled.")
                            Return
                        End If

                        Dim desktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                        newFilePath = IO.Path.Combine(desktopPath, sanitizedName & ".drawio")
                    End If

                    ' Check if file already exists and find unique name
                    Dim counter As Integer = 1
                    Dim basePath As String = IO.Path.Combine(IO.Path.GetDirectoryName(newFilePath), IO.Path.GetFileNameWithoutExtension(newFilePath))
                    While IO.File.Exists(newFilePath)
                        newFilePath = $"{basePath}_{counter:D3}.drawio"
                        counter += 1
                    End While

                    ' Create empty draw.io XML
                    Dim emptyXml As String = CreateEmptyDrawioXml()

                    ' Save the empty file
                    IO.File.WriteAllText(newFilePath, emptyXml, System.Text.Encoding.UTF8)

                    ' Ask how to open
                    Dim choice As DrawioOpenChoice =
                        AskHowToOpenDrawio(
                            prompt:=$"A new empty diagram has been created at '{IO.Path.GetFileName(newFilePath)}'.{vbCrLf}{vbCrLf}" &
                                    "Choose how to open the draw.io editor:",
                            caption:=$"{AN} - New Diagram Created")

                    Select Case choice
                        Case DrawioOpenChoice.OfflineAfterLoad
                            OpenDrawioWithWebView2(emptyXml, newFilePath, disableInternetAfterLoad:=True)

                        Case DrawioOpenChoice.Online
                            OpenDrawioWithWebView2(emptyXml, newFilePath, disableInternetAfterLoad:=False)

                        Case Else
                            ShowCustomMessageBox($"The empty diagram has been saved to:{vbCrLf}{newFilePath}")
                            Return
                    End Select

                    Return
                End If
                selectedPath = f.SelectedFilePath
            End Using

            If String.IsNullOrWhiteSpace(selectedPath) OrElse Not IO.File.Exists(selectedPath) Then
                ShowCustomMessageBox("No valid .drawio file was selected.")
                Return
            End If

            Dim xmlContent As String = IO.File.ReadAllText(selectedPath, System.Text.Encoding.UTF8)

            ' Use the shared 3-button chooser based on ShowCustomYesNoBox(extraButtonText,...)
            Dim choiceExisting As DrawioOpenChoice =
                AskHowToOpenDrawio(
                    prompt:=$"Open '{IO.Path.GetFileName(selectedPath)}' in the embedded draw.io editor?" & vbCrLf & vbCrLf &
                            "Choose the mode:",
                    caption:=$"{AN} Open draw.io")

            Select Case choiceExisting
                Case DrawioOpenChoice.OfflineAfterLoad
                    OpenDrawioWithWebView2(xmlContent, selectedPath, disableInternetAfterLoad:=True)

                Case DrawioOpenChoice.Online
                    OpenDrawioWithWebView2(xmlContent, selectedPath, disableInternetAfterLoad:=False)

                Case Else
                    Return
            End Select

        Catch ex As Exception
            Debug.WriteLine($"OpenExistingDrawioFileForEditing error: {ex.Message}")
            ShowCustomMessageBox($"Could not open the .drawio file: {ex.Message}")
        Finally
            ' Best-effort: reset global drag-drop customization so other uses are unaffected
            Try
                Globals.ThisAddIn.DragDropFormFilter = ""
                Globals.ThisAddIn.DragDropFormLabel = ""
            Catch
            End Try
        End Try
    End Sub

    ''' <summary>
    ''' Creates an empty draw.io XML structure that can be used as a starting point for new diagrams.
    ''' </summary>
    ''' <returns>A valid empty draw.io XML string.</returns>
    Private Function CreateEmptyDrawioXml() As String
        Dim timestamp As String = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
        Return $"<?xml version=""1.0"" encoding=""UTF-8""?>
<mxfile host=""embed.diagrams.net"" modified=""{timestamp}"" agent=""{AN}"" version=""21.0.0"" etag=""empty"" type=""device"">
  <diagram name=""Page-1"" id=""page1"">
    <mxGraphModel dx=""800"" dy=""600"" grid=""1"" gridSize=""10"" guides=""1"" tooltips=""1"" connect=""1"" arrows=""1"" fold=""1"" page=""1"" pageScale=""1"" pageWidth=""827"" pageHeight=""1169"" math=""0"" shadow=""0"">
      <root>
        <mxCell id=""0""/>
        <mxCell id=""1"" parent=""0""/>
      </root>
    </mxGraphModel>
  </diagram>
</mxfile>"
    End Function

    ''' <summary>
    ''' Sanitizes a file name by removing invalid characters.
    ''' </summary>
    ''' <param name="fileName">The file name to sanitize.</param>
    ''' <returns>A sanitized file name safe for use in the file system.</returns>
    Private Function SanitizeFileName2(fileName As String) As String
        If String.IsNullOrWhiteSpace(fileName) Then Return String.Empty

        ' Remove invalid file name characters
        Dim invalidChars As Char() = IO.Path.GetInvalidFileNameChars()
        Dim sanitized As String = fileName

        For Each c As Char In invalidChars
            sanitized = sanitized.Replace(c.ToString(), "")
        Next

        ' Also remove any path separators that might have slipped through
        sanitized = sanitized.Replace("/", "").Replace("\", "")

        ' Trim whitespace and dots from start/end
        sanitized = sanitized.Trim().Trim("."c)

        Return sanitized
    End Function

    ''' <summary>
    ''' Extracts draw.io XML from an LLM response and cleans it for saving/opening.
    ''' Handles Markdown code fences and “extra text around the XML” by extracting the first
    ''' <c>&lt;mxfile&gt;...&lt;/mxfile&gt;</c> (or fallback <c>&lt;mxGraphModel&gt;</c>) block.
    ''' </summary>
    ''' <param name="llmResponse">Raw LLM response potentially containing draw.io XML.</param>
    ''' <returns>
    ''' The extracted and trimmed draw.io XML, or an empty string if no suitable XML could be found.
    ''' </returns>
    Private Function ExtractAndCleanDrawioXml(ByVal llmResponse As String) As String
        If String.IsNullOrWhiteSpace(llmResponse) Then Return String.Empty

        Dim s As String = llmResponse.Trim()

        ' 1) Remove markdown fences if they wrap the whole payload.
        ' (This is a best-effort pre-clean; we still do XML extraction below.)
        If s.StartsWith("```xml", StringComparison.OrdinalIgnoreCase) Then
            s = s.Substring(6)
        ElseIf s.StartsWith("```", StringComparison.Ordinal) Then
            s = s.Substring(3)
        End If
        If s.EndsWith("```", StringComparison.Ordinal) Then
            s = s.Substring(0, s.Length - 3)
        End If
        s = s.Trim()

        ' 2) Prefer extracting the full draw.io wrapper (<mxfile>...</mxfile>).
        Dim mxfileMatch As Match = Regex.Match(
            s,
            "<mxfile[\s\S]*?</mxfile>",
            RegexOptions.IgnoreCase Or RegexOptions.Singleline)

        If mxfileMatch.Success Then
            Return mxfileMatch.Value.Trim()
        End If

        ' 3) Fallback: some outputs might contain only the model (<mxGraphModel>...</mxGraphModel>).
        Dim mxGraphMatch As Match = Regex.Match(
            s,
            "<mxGraphModel[\s\S]*?</mxGraphModel>",
            RegexOptions.IgnoreCase Or RegexOptions.Singleline)

        If mxGraphMatch.Success Then
            Return mxGraphMatch.Value.Trim()
        End If

        Return String.Empty
    End Function

    ''' <summary>
    ''' Restores the user's previous UI window and selection after processing temporarily moved
    ''' focus away from the original editing context.
    ''' </summary>
    ''' <param name="app">Word application instance.</param>
    ''' <param name="savedWindow">Previously active window.</param>
    ''' <param name="savedSelection">Previously active selection range.</param>
    Private Sub RestoreUiContext(ByVal app As Word.Application,
                             ByVal savedWindow As Word.Window,
                             ByVal savedSelection As Word.Range)
        Try
            If savedWindow IsNot Nothing Then
                savedWindow.Activate()
            End If
        Catch ex As Exception
            Debug.WriteLine($"RestoreUiContext(window) failed: {ex.Message}")
        End Try

        Try
            If savedSelection IsNot Nothing Then
                savedSelection.Select()
            End If
        Catch ex As Exception
            Debug.WriteLine($"RestoreUiContext(selection) failed: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Reactivates the live Word UI context for the supplied processing range so subsequent
    ''' write-back operations run against the intended document window and selection.
    ''' </summary>
    ''' <param name="targetRange">Range that should become the active processing selection.</param>
    Private Sub ActivateProcessingContext(ByVal targetRange As Word.Range)
        Try
            If targetRange Is Nothing Then Return

            Dim targetWindow As Word.Window = targetRange.Document.ActiveWindow
            If targetWindow IsNot Nothing Then
                targetWindow.Activate()
            End If

            targetRange.Select()
        Catch ex As Exception
            Debug.WriteLine($"ActivateProcessingContext failed: {ex.Message}")
        End Try
    End Sub

    Private Shared Function TryResolveActiveComment(
    ByVal Selection As Word.Selection,
    ByRef activeComment As Word.Comment) As Boolean

        activeComment = Nothing

        If Selection Is Nothing OrElse Selection.Range Is Nothing Then Return False

        Dim doc As Word.Document = Nothing
        Try
            doc = Selection.Range.Document
        Catch
            doc = Nothing
        End Try
        If doc Is Nothing Then Return False

        Try
            If Selection.StoryType = Word.WdStoryType.wdCommentsStory Then
                For Each c As Word.Comment In doc.Comments
                    If Selection.Range.Start >= c.Range.Start AndAlso
                       Selection.Range.End <= c.Range.End Then
                        activeComment = c
                        Return True
                    End If
                Next
            End If
        Catch
        End Try

        ' Only treat a main-text position as "inside a comment" when the user has NOT made a
        ' real text selection. A broad selection (e.g. the entire document) can merely overlap a
        ' comment's anchor/scope; in that case the user is editing the main text, not the comment,
        ' and must not be routed to the comment-processing pipeline. Selections inside the comments
        ' story are already handled by the wdCommentsStory branch above.
        Dim isInsertionPoint As Boolean =
            (Selection.Type = Word.WdSelectionType.wdSelectionIP) OrElse
            (Selection.Range.Start = Selection.Range.End)

        If isInsertionPoint AndAlso Selection.StoryType = Word.WdStoryType.wdMainTextStory Then
            Try
                Dim cursorPos As Integer = Selection.Range.Start

                For Each c As Word.Comment In doc.Comments
                    Dim anchor As Word.Range = Nothing

                    Try
                        anchor = c.Scope
                    Catch
                        Try
                            anchor = c.Reference
                        Catch
                            anchor = Nothing
                        End Try
                    End Try

                    If anchor IsNot Nothing AndAlso
                       cursorPos >= anchor.Start AndAlso
                       cursorPos <= anchor.End Then
                        activeComment = c
                        Return True
                    End If
                Next
            Catch
            End Try
        End If

        Return False
    End Function

End Class
