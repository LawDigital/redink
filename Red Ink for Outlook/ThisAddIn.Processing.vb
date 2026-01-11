' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.Processing.vb
' Purpose: Text processing helpers for the Outlook add-in: Word-based document comparison
'          (tracked changes flattened), inline diff (DiffPlex) with custom markup tags,
'          Markdown conversion from Word formatting, formatted insertion routines,
'          RTF to plain text conversion, parsing of markup tag sequences, and word count.
'
' Architecture:
' - CompareAndInsertTextCompareDocs: uses Word CompareDocuments API to build formatting markup, applies static styling, pastes into current compose window.
' - CompareAndInsertText: builds a word-level diff via DiffPlex, encodes insert/delete spans with pseudo tags for later HTML/format rendering.
' - ConvertRangeToMarkdown: traverses a Word.Range; maps headings, lists, and basic character formatting (bold/italic/underline/strikethrough) into Markdown or inline HTML escapes.
' - ReplaceWithinRange: controlled find/replace loop with boundary preservation; single replacement iterations and rollback if exceeding original range.
' - InsertFormattedTextFast / InsertFormattedText: two insertion paths (HTML fragment vs incremental tagged spans) applying blue underline for insertions and red strikethrough for deletions.
' - ParseText: linear tag parser producing parallel arrays of segments and formatting codes.
' - ConvertRtfToPlainText: regex-based RTF control word stripping and hex escape decoding.
' - RGB: utility composing a DWORD color.
' - GetSelectedTextLength: acquires selection, filters for HTML/RTF mail, counts words via whitespace split.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports DiffPlex
Imports DiffPlex.DiffBuilder
Imports DiffPlex.DiffBuilder.Model
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Outlook
Imports Microsoft.Office.Interop.Word
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ''' <summary>
    ''' Compares two text inputs using Word.CompareDocuments, applies static formatting for insertions (blue underline) and deletions (red strikethrough),
    ''' accepts revisions, and pastes the result into the active Outlook compose window.
    ''' </summary>
    Private Sub CompareAndInsertTextCompareDocs(input1 As String, input2 As String)

        Dim splash As New SplashScreen("Creating markup using the Word compare functionality (ignore any flickering and press 'No' if prompted) ...")
        splash.Show()
        splash.Refresh()
        Try
            ' Get the active inspector (compose mail window)
            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
            Dim inspector As Inspector = ComRetry(Function() outlookApp.ActiveInspector)

            ' Ensure the current item is a MailItem and in compose mode (COM-safe)
            If inspector Is Nothing Then
                System.Windows.Forms.MessageBox.Show("Error in CompareAndInsertTextCompareDocs: No active inspector.",
                                         "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            Dim curr As Object = Nothing
            Try
                curr = ComRetry(Function() inspector.CurrentItem)
            Catch
                curr = Nothing
            End Try

            If curr Is Nothing OrElse Not TypeOf curr Is Microsoft.Office.Interop.Outlook.MailItem Then
                System.Windows.Forms.MessageBox.Show("Error in CompareAndInsertTextCompareDocs: No active email item.",
                                         "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem =
    ComRetry(Function() CType(curr, Microsoft.Office.Interop.Outlook.MailItem))

            Dim editor As Object = ComRetry(Function() inspector.WordEditor)

            ' Cast the WordEditor to Word.Document
            Dim wordDoc As Document = CType(editor, Document)

            ' Create a new temporary Word application for comparison
            Dim wordApp As New Microsoft.Office.Interop.Word.Application()
            wordApp.Visible = False

            ' Create temporary documents for input1 and input2
            Dim tempDoc1 As Document = wordApp.Documents.Add()
            Dim tempDoc2 As Document = wordApp.Documents.Add()

            ' Insert the input texts into the temporary documents
            tempDoc1.Content.Text = input1
            tempDoc2.Content.Text = input2

            ' Perform the comparison
            Dim compareResult As Document = wordApp.CompareDocuments(tempDoc1, tempDoc2,
                                                        WdCompareDestination.wdCompareDestinationNew,
                                                        WdGranularity.wdGranularityWordLevel,
                                                        False, False, False, False, False, False)

            ' Convert tracked changes to static formatting
            For Each revision As Revision In compareResult.Revisions
                Select Case revision.Type
                    Case WdRevisionType.wdRevisionInsert
                        ' Insertions: Apply blue color and underline
                        revision.Range.Font.Color = WdColor.wdColorBlue
                        revision.Range.Font.Underline = WdUnderline.wdUnderlineSingle
                    Case WdRevisionType.wdRevisionDelete
                        ' Deletions: Apply red color and strikethrough
                        revision.Range.Font.Color = WdColor.wdColorRed
                        revision.Range.Font.StrikeThrough = True
                End Select
                revision.Accept() ' Accept the revision to make the formatting static
            Next

            ' Copy the comparison result to clipboard
            compareResult.Content.Copy()

            ' Paste the comparison result into the Outlook compose window at the current selection
            wordDoc.Application.Selection.PasteAndFormat(WdRecoveryType.wdFormatOriginalFormatting)

            ' Clean up
            tempDoc1.Close(False)
            tempDoc2.Close(False)
            compareResult.Close(False)
            wordApp.Quit(False)

            ' Release COM objects in reverse order of creation
            If inspector IsNot Nothing Then Marshal.ReleaseComObject(inspector) : inspector = Nothing

        Catch ex As System.Exception
            MessageBox.Show("Error in CompareAndInsertTextCompareDocs: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            splash.Close()

        End Try
    End Sub

    ''' <summary>
    ''' Produces an inline word-level diff between two texts using DiffPlex, encodes changes with custom pseudo tags,
    ''' normalizes line break markers, and inserts or displays formatted result.
    ''' </summary>
    Private Sub CompareAndInsertText(text1 As String, text2 As String, Optional ShowInWindow As Boolean = False, Optional TextforWindow As String = "A text with these changes will be inserted ('Esc' to abort):", Optional DoNotWait As Boolean = False)
        Dim diffBuilder As New InlineDiffBuilder(New Differ())
        Dim sText As String = String.Empty

        ' Pre-process the texts to replace line breaks with a unique marker
        text1 = text1.Replace(vbCrLf, " {vbCrLf} ").Replace(vbCr, " {vbCr} ").Replace(vbLf, " {vbLf} ")
        text2 = text2.Replace(vbCrLf, " {vbCrLf} ").Replace(vbCr, " {vbCr} ").Replace(vbLf, " {vbLf} ")

        ' Normalize the texts by removing extra spaces
        text1 = text1.Replace("  ", " ").Trim()
        text2 = text2.Replace("  ", " ").Trim()

        ' Split the texts into words and convert them into a line-by-line format
        Dim words1 As String = String.Join(Environment.NewLine, text1.Split(" "c))
        Dim words2 As String = String.Join(Environment.NewLine, text2.Split(" "c))

        ' Generate word-based diff using DiffPlex
        Dim diffResult As DiffPaneModel = diffBuilder.BuildDiffModel(words1, words2)

        ' Build the formatted output based on the diff results
        For Each line In diffResult.Lines
            Select Case line.Type
                Case ChangeType.Inserted
                    sText &= "[INS_START]" & line.Text.Trim() & "[INS_END] "
                Case ChangeType.Deleted
                    sText &= "[DEL_START]" & line.Text.Trim() & "[DEL_END] "
                Case ChangeType.Unchanged
                    sText &= line.Text.Trim() & " "
            End Select
        Next

        ' Remove preceding and trailing spaces around placeholders
        sText = sText.Replace("{vbCr}", "{vbCrLf}")
        sText = sText.Replace("{vbLf}", "{vbCrLf}")
        sText = sText.Replace(" {vbCrLf} ", "{vbCrLf}")
        sText = sText.Replace(" {vbCrLf}", "{vbCrLf}")
        sText = sText.Replace("{vbCrLf} ", "{vbCrLf}")

        ' Remove instances of line breaks surrounded by [DEL_START] and [DEL_END]
        sText = sText.Replace("[DEL_START]{vbCrLf}[DEL_END] ", "")

        ' Include instances of line breaks surrounded by [INS_START] and [INS_END] without the [INS...] text
        sText = sText.Replace("[INS_START]{vbCrLf}[INS_END] ", "{vbCrLf}")

        ' Replace placeholders with actual line breaks
        sText = sText.Replace("{vbCrLf}", vbCrLf)

        ' Adjust overlapping tags
        sText = sText.Replace("[DEL_END] [INS_START]", "[DEL_END][INS_START]")
        sText = sText.Replace("[INS_START][INS_END] ", "")

        ' Insert formatted text into the Word editor
        If Not ShowInWindow Then
            InsertFormattedTextFast(sText)
        Else
            Dim htmlContent As String = ConvertMarkupToRTF(TextforWindow & "\r\r" & sText)
            System.Threading.Tasks.Task.Run(Sub()
                                                ShowRTFCustomMessageBox(htmlContent)
                                            End Sub)
        End If

    End Sub

    ''' <summary>
    ''' Converts Word paragraphs and character formatting in a range to Markdown equivalents (headings, lists, bold, italic, underline via HTML <u>, strikethrough).
    ''' </summary>
    Private Shared Sub ConvertRangeToMarkdown(WorkingRange As Word.Range)

        Dim listRegex As New Regex("^(\s*)([-*+]|\d+[\.\)])\s+", RegexOptions.Compiled)

        Dim rng As Word.Range = WorkingRange.Duplicate
        If rng.End < rng.Document.Content.End - 1 Then
            rng.End = rng.End + 1
        End If

        Dim doc As Microsoft.Office.Interop.Word.Document = rng.Document

        ' 0) Headings & lists
        For Each para As Microsoft.Office.Interop.Word.Paragraph In rng.Paragraphs
            Dim styleName As String = CType(para.Style, Microsoft.Office.Interop.Word.Style).NameLocal

            Select Case styleName
                Case doc.Styles(Word.WdBuiltinStyle.wdStyleTitle).NameLocal
                    para.Range.InsertBefore("# ")
                Case doc.Styles(Word.WdBuiltinStyle.wdStyleHeading1).NameLocal
                    para.Range.InsertBefore("# ")
                Case doc.Styles(Word.WdBuiltinStyle.wdStyleHeading2).NameLocal
                    para.Range.InsertBefore("## ")
                Case doc.Styles(Word.WdBuiltinStyle.wdStyleHeading3).NameLocal
                    para.Range.InsertBefore("### ")
                Case doc.Styles(Word.WdBuiltinStyle.wdStyleHeading4).NameLocal
                    para.Range.InsertBefore("#### ")
                Case doc.Styles(Word.WdBuiltinStyle.wdStyleHeading5).NameLocal
                    para.Range.InsertBefore("##### ")
                Case doc.Styles(Word.WdBuiltinStyle.wdStyleHeading6).NameLocal
                    para.Range.InsertBefore("###### ")
            End Select

            ' — List detection
            With para.Range.ListFormat
                Try
                    ' Proceed only if a list is present
                    If .ListType <> Microsoft.Office.Interop.Word.WdListType.wdListNoNumbering Then

                        ' 1) Store needed information before RemoveNumbers
                        Dim lvl As Integer = .ListLevelNumber
                        Dim lt As Microsoft.Office.Interop.Word.WdListType = .ListType
                        Dim ls As String = .ListString.Trim()

                        ' 2) Compute prefix (4 spaces per level)
                        Dim indent As String = New String(" "c, (lvl - 1) * 4)
                        Dim prefix As String
                        Select Case lt
                            Case Microsoft.Office.Interop.Word.WdListType.wdListBullet,
                                 Microsoft.Office.Interop.Word.WdListType.wdListPictureBullet
                                prefix = indent & "- "
                            Case Microsoft.Office.Interop.Word.WdListType.wdListSimpleNumbering,
                                 Microsoft.Office.Interop.Word.WdListType.wdListOutlineNumbering,
                                 Microsoft.Office.Interop.Word.WdListType.wdListMixedNumbering,
                                 Microsoft.Office.Interop.Word.WdListType.wdListListNumOnly
                                prefix = indent & ls & " "
                            Case Else
                                prefix = indent & "- "
                        End Select

                        ' 3) Remove list formatting
                        .RemoveNumbers()

                        ' 4) Insert Markdown prefix at line start
                        Dim insertRange As Microsoft.Office.Interop.Word.Range = para.Range.Duplicate()
                        insertRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseStart)
                        insertRange.InsertBefore(prefix)

                        ' Range for inserted prefix
                        Dim prefixRange As Word.Range = insertRange.Duplicate
                        prefixRange.End = prefixRange.Start + prefix.Length

                        ' Reset font (standard formatting)
                        prefixRange.Font.Reset()

                    End If

                Catch ex As System.Exception
                    System.Diagnostics.Debug.WriteLine("Error during list conversion: " & ex.ToString())
                End Try
            End With

        Next

        ' 1) Bold + Italic (Paragraph)
        ReplaceWithinRange(rng,
                        Sub(f)
                            f.Font.Bold = True
                            f.Font.Italic = True
                            f.Font.Underline = Word.WdUnderline.wdUnderlineNone
                            f.Text = "(*)^13"
                            f.MatchWildcards = True
                        End Sub,
                        "***\1***^13",
                        Sub(rep)                          ' Disable Bold & Italic
                            rep.Bold = False
                            rep.Italic = False
                        End Sub)

        ' 2) Bold + Italic (Inline)
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

        ' 3) Bold only (Paragraph)
        ReplaceWithinRange(rng,
                        Sub(f)
                            f.Font.Bold = True
                            f.Text = "(*)^13"
                            f.MatchWildcards = True
                        End Sub,
                        "**\1**^13",
                        Sub(rep)
                            rep.Bold = False
                        End Sub)

        ' 4) Bold only (Inline)
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

        ' 5) Italic only (Paragraph)
        ReplaceWithinRange(rng,
                        Sub(f)
                            f.Font.Italic = True
                            f.Text = "(*)^13"
                            f.MatchWildcards = True
                        End Sub,
                        "*\1*^13",
                        Sub(rep)
                            rep.Italic = False
                        End Sub)

        ' 6) Italic only (Inline)
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

        ' 7) Underline (Paragraph)
        ReplaceWithinRange(rng,
                        Sub(f)
                            f.Font.Underline = Word.WdUnderline.wdUnderlineSingle
                            f.Text = "(*)^13"
                            f.MatchWildcards = True
                        End Sub,
                        "<u>\1</u>^13",
                        Sub(rep)
                            rep.Underline = Word.WdUnderline.wdUnderlineNone
                        End Sub)

        ' 8) Underline (Inline)
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

        ' 9) Strikethrough (Paragraph)
        ReplaceWithinRange(rng,
                        Sub(f)
                            f.Font.StrikeThrough = True
                            f.Text = "(*)^13"
                            f.MatchWildcards = True
                        End Sub,
                        "~~\1~~^13",
                        Sub(rep)
                            rep.StrikeThrough = False
                        End Sub)

        '10) Strikethrough (Inline)
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

        ' Restore selection
        rng.End = rng.End - 1
        rng.Select()

    End Sub

    ''' <summary>
    ''' Performs iterative find/replace operations within a constrained range; prevents replacements from exceeding original bounds.
    ''' </summary>
    Private Shared Sub ReplaceWithinRange(
    ByVal rng As Word.Range,
    ByVal configureFind As Action(Of Word.Find),
    ByVal replacementText As String,
    ByVal tweakReplacement As Action(Of Word.Font))

        Dim doc As Word.Document = rng.Document
        Dim originalStart As Long = rng.Start
        Dim originalEnd As Long = rng.End
        Dim currentPosition As Long = originalStart

        Do
            ' Create a range from current position to the end of the original range
            Dim searchRange As Word.Range = doc.Range(currentPosition, originalEnd)
            Dim f As Word.Find = searchRange.Find

            Debug.WriteLine($"Searchrange = '{searchRange.Text}'")

            f.ClearFormatting()
            f.Replacement.ClearFormatting()

            configureFind(f)
            f.Replacement.Text = replacementText
            tweakReplacement(f.Replacement.Font)

            f.Forward = True
            f.Wrap = Word.WdFindWrap.wdFindStop
            f.Format = True

            ' If no more matches, exit
            If Not f.Execute(Replace:=Word.WdReplace.wdReplaceOne) Then Exit Do

            Debug.WriteLine($"Searchrange = '{searchRange.Text}' (after change)")

            ' After replacement, searchRange now points to the match
            ' Check if this match/replacement went beyond boundary
            If searchRange.End > originalEnd Then
                Debug.WriteLine("Went too far!")
                doc.Undo()
                Exit Do
            End If

            ' Continue from end of this match
            currentPosition = searchRange.End
            originalEnd = rng.End

        Loop While currentPosition < originalEnd

        rng.SetRange(originalStart, originalEnd)
    End Sub

    ''' <summary>
    ''' Strips RTF control words, decodes hex escapes, converts \par markers to newline, returns plain UTF-16 text.
    ''' </summary>
    Private Function ConvertRtfToPlainText(rtfContent As String) As String
        If String.IsNullOrWhiteSpace(rtfContent) Then Return String.Empty

        ' Remove RTF headers and control words
        Dim plainText As String = Regex.Replace(rtfContent, "{\\.*?}|\\[a-z]+[0-9]*|[{}]", String.Empty)

        ' Decode escaped characters (e.g., \'xx)
        plainText = Regex.Replace(plainText, "\\'([0-9a-fA-F]{2})", Function(m)
                                                                        Dim hex = System.Convert.ToByte(m.Groups(1).Value, 16)
                                                                        Return Chr(hex)
                                                                    End Function)

        ' Replace RTF line breaks (\par) with actual line breaks
        plainText = Regex.Replace(plainText, "\\par", Environment.NewLine, RegexOptions.IgnoreCase)

        ' Trim the result
        plainText = plainText.Trim()

        Return plainText
    End Function

    ''' <summary>
    ''' Fast insertion of markup: converts pseudo tags to HTML spans and inserts as an HTML fragment at current selection.
    ''' </summary>
    Private Sub InsertFormattedTextFast(ByVal inputText As String)

        ' 1. Convert the pseudo-tags to plain HTML
        Dim markup As String = System.Net.WebUtility.HtmlEncode(inputText)

        ' Preserve line breaks (optional – remove if paragraphs required)
        markup = markup.Replace(vbCrLf, "<br>")

        ' Replace each tag with an inline <span>
        markup = markup.Replace("[INS_START]",
                "<span style=""color:#0000FF;text-decoration:underline;"">") _
                   .Replace("[INS_END]", "</span>") _
                   .Replace("[DEL_START]",
                "<span style=""color:#FF0000;text-decoration:line-through;"">") _
                   .Replace("[DEL_END]", "</span>")

        ' 2. Get the current Outlook inspector / Word selection
        Dim inspector As Microsoft.Office.Interop.Outlook.Inspector =
        ComRetry(Function() TryCast(Globals.ThisAddIn.Application.ActiveInspector,
                Microsoft.Office.Interop.Outlook.Inspector))

        If inspector Is Nothing Then
            System.Windows.Forms.MessageBox.Show(
            "No open mail item found.",
            "InsertFormattedTextFast",
            System.Windows.Forms.MessageBoxButtons.OK,
            System.Windows.Forms.MessageBoxIcon.Error)
            Exit Sub
        End If

        Dim wordDoc As Microsoft.Office.Interop.Word.Document =
        ComRetry(Function() TryCast(inspector.WordEditor,
                Microsoft.Office.Interop.Word.Document))

        If wordDoc Is Nothing Then
            System.Windows.Forms.MessageBox.Show(
            "Unable to access the Word editor.",
            "InsertFormattedTextFast",
            System.Windows.Forms.MessageBoxButtons.OK,
            System.Windows.Forms.MessageBoxIcon.Error)
            Exit Sub
        End If

        Dim app As Microsoft.Office.Interop.Word.Application = wordDoc.Application
        Dim selRange As Microsoft.Office.Interop.Word.Range = app.Selection.Range

        ' 3. Insert the fragment in one shot
        Dim oldScreenUpdating As Boolean = app.ScreenUpdating
        app.ScreenUpdating = False

        Try
            ' Existing helper that pastes an HTML fragment
            InsertTextWithFormat(markup, selRange, True, True)

            ' Add a CRLF AFTER insertion (Word paragraph mark)
            selRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
            selRange.InsertAfter(vbCrLf & vbCrLf)
            selRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)


        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(
            ex.Message,
            "InsertFormattedTextFast",
            System.Windows.Forms.MessageBoxButtons.OK,
            System.Windows.Forms.MessageBoxIcon.Error)

        Finally
            ' Restore Word UI state
            app.ScreenUpdating = oldScreenUpdating

            ' 4. Clean up COM objects in reverse order of creation
            If selRange IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(selRange)
                selRange = Nothing
            End If
            If wordDoc IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDoc)
                wordDoc = Nothing
            End If
            If inspector IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(inspector)
                inspector = Nothing
            End If
        End Try
    End Sub

    ''' <summary>
    ''' Inserts text chunk by chunk applying formatting based on pseudo tag boundaries ([INS_START]/[DEL_START]) with user abort capability (Esc).
    ''' </summary>
    Private Sub InsertFormattedText(inputText As String)

        Dim objInspector As Microsoft.Office.Interop.Outlook.Inspector
        Dim objWordDoc As Microsoft.Office.Interop.Word.Document
        Dim objSelection As Object
        Dim objRange As Object
        Dim TextArray() As String = {}
        Dim FormatArray() As Integer = {}
        Dim i As Integer

        ' Store original font properties
        Dim originalFontColor As Integer = 0
        Dim originalUnderline As Integer = 0
        Dim originalStrikeThrough As Boolean = False
        Dim originalBold As Boolean = False
        Dim originalItalic As Boolean = False

        ' Check if there is an active inspector (open email)
        objInspector = ComRetry(Function() TryCast(Globals.ThisAddIn.Application.ActiveInspector, Microsoft.Office.Interop.Outlook.Inspector))
        If objInspector Is Nothing Then
            MessageBox.Show("Error in InsertFormattedText: No open mail item found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        ' Get the Word editor and the current selection
        objWordDoc = ComRetry(Function() TryCast(objInspector.WordEditor, Microsoft.Office.Interop.Word.Document))
        If objWordDoc Is Nothing Then
            MessageBox.Show("Error in InsertFormattedText: Unable to access the necessary mail editor for this mail.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        objSelection = objWordDoc.Application.Selection

        ' Store original font properties
        If objSelection.Font IsNot Nothing Then
            With objSelection.Font
                originalFontColor = .Color
                originalUnderline = .Underline
                originalStrikeThrough = .StrikeThrough
                originalBold = .Bold
                originalItalic = .Italic
            End With
        End If

        Dim splash As New SplashScreen("Creating your markup ... press 'Esc' to abort")
        splash.Show()
        splash.Refresh()

        ' Parse the input text into chunks with formatting information
        ParseText(inputText, TextArray, FormatArray)

        ' Reset formatting before starting
        If objSelection.Font IsNot Nothing Then objSelection.Font.Reset()

        ' Insert each text chunk with the appropriate formatting
        For i = 0 To TextArray.Length - 1

            If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 Then
                Exit For
            End If

            If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then
                Exit For
            End If

            ' Reset formatting to original before each insertion
            If objSelection.Font IsNot Nothing Then
                With objSelection.Font
                    .Color = originalFontColor
                    .Underline = originalUnderline
                    .StrikeThrough = originalStrikeThrough
                    .Bold = originalBold
                    .Italic = originalItalic
                End With
            End If

            ' Insert the text at the current cursor position
            objSelection.Collapse(0) ' Collapse to insertion point
            objSelection.TypeText(TextArray(i))

            ' Define the range for the inserted text
            objRange = objSelection.Range
            objRange.Start = objSelection.Start - TextArray(i).Length
            objRange.End = objSelection.Start

            ' Apply formatting based on the tag
            Select Case FormatArray(i)
                Case 1 ' [INS_START]...[INS_END]: Blue underline
                    If objRange.Font IsNot Nothing Then
                        With objRange.Font
                            .Color = RGB(0, 0, 255)
                            .Underline = True
                            .StrikeThrough = False
                        End With
                    End If
                Case 2 ' [DEL_START]...[DEL_END]: Red strikethrough
                    If objRange.Font IsNot Nothing Then
                        With objRange.Font
                            .Color = RGB(255, 0, 0)
                            .StrikeThrough = True
                            .Underline = False
                        End With
                    End If
                Case Else ' Normal text
                    ' Already reset to original formatting
            End Select
        Next

        ' Ensure formatting is reset after all insertions
        If objSelection.Font IsNot Nothing Then
            With objSelection.Font
                .Color = originalFontColor
                .Underline = originalUnderline
                .StrikeThrough = originalStrikeThrough
                .Bold = originalBold
                .Italic = originalItalic
            End With
        End If

        splash.Close()

        ' Release COM objects in reverse order of creation
        If objInspector IsNot Nothing Then Marshal.ReleaseComObject(objInspector) : objInspector = Nothing
        If objWordDoc IsNot Nothing Then Marshal.ReleaseComObject(objWordDoc) : objWordDoc = Nothing

    End Sub


    ''' <summary>
    ''' Parses tagged markup text into arrays of plain text segments and associated formatting codes:
    ''' 0 = normal, 1 = insertion span, 2 = deletion span.
    ''' </summary>
    Private Sub ParseText(inputText As String, ByRef TextArray() As String, ByRef FormatArray() As Integer)
        Dim pos As Integer = 1
        Dim lenText As Integer = inputText.Length
        Dim nextTagPos As Integer
        Dim tagEndPos As Integer
        Dim tagText As String
        Dim chunkIndex As Integer = 0
        Dim tagType As Integer
        Dim nextInsPos As Integer
        Dim nextDelPos As Integer

        While pos <= lenText
            If inputText.Substring(pos - 1, System.Math.Min(11, lenText - pos + 1)) = "[INS_START]" Then
                pos += 11
                tagType = 1 ' Insert formatting
                Dim rawIndex As Integer = inputText.IndexOf("[INS_END]", pos - 1)
                If rawIndex = -1 Then
                    MessageBox.Show("Error in ParseText: Missing [INS_END] tag.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If
                tagEndPos = rawIndex + 1
                tagText = inputText.Substring(pos - 1, tagEndPos - pos)
                pos = tagEndPos + 9
            ElseIf inputText.Substring(pos - 1, System.Math.Min(11, lenText - pos + 1)) = "[DEL_START]" Then
                pos += 11
                tagType = 2 ' Delete formatting
                Dim rawIndex As Integer = inputText.IndexOf("[DEL_END]", pos - 1)
                If rawIndex = -1 Then
                    MessageBox.Show("Error in ParseText: Missing [DEL_END] tag.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If
                tagEndPos = rawIndex + 1
                tagText = inputText.Substring(pos - 1, tagEndPos - pos)
                pos = tagEndPos + 9
            Else
                tagType = 0
                nextInsPos = inputText.IndexOf("[INS_START]", pos - 1) + 1
                If nextInsPos = 0 Then nextInsPos = lenText + 1
                nextDelPos = inputText.IndexOf("[DEL_START]", pos - 1) + 1
                If nextDelPos = 0 Then nextDelPos = lenText + 1
                nextTagPos = System.Math.Min(nextInsPos, nextDelPos)
                tagText = inputText.Substring(pos - 1, nextTagPos - pos)
                pos = nextTagPos
            End If

            chunkIndex += 1
            ReDim Preserve TextArray(chunkIndex - 1)
            ReDim Preserve FormatArray(chunkIndex - 1)
            TextArray(chunkIndex - 1) = tagText
            FormatArray(chunkIndex - 1) = tagType
        End While
    End Sub

    ''' <summary>
    ''' Combines RGB components into a packed integer (blue shifted left 16, green left 8).
    ''' </summary>
    Private Function RGB(ByVal red As Integer, ByVal green As Integer, ByVal blue As Integer) As Integer
        Return red Or (green << 8) Or (blue << 16)
    End Function

    ''' <summary>
    ''' Returns word count of current selection in an HTML/RTF MailItem. Returns 0 if no valid selection or plain text format.
    ''' </summary>
    Private Function GetSelectedTextLength() As Integer
        Try
            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
            Dim inspector As Microsoft.Office.Interop.Outlook.Inspector = ComRetry(Function() outlookApp.ActiveInspector())

            ' Guard CurrentItem via ComRetry to avoid transient COM rejections
            Dim curr As Object = Nothing
            If inspector IsNot Nothing Then
                Try
                    curr = ComRetry(Function() inspector.CurrentItem)
                Catch
                    curr = Nothing
                End Try
            End If

            If inspector Is Nothing _
               OrElse curr Is Nothing _
               OrElse Not TypeOf curr Is Microsoft.Office.Interop.Outlook.MailItem Then
                Return 0
            End If

            Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem =
                CType(curr, Microsoft.Office.Interop.Outlook.MailItem)

            ' Reject plain text mails
            If mailItem.BodyFormat = Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatPlain Then
                Return 0
            End If

            ' Acquire Word editor
            Dim wordEditor As Microsoft.Office.Interop.Word.Document =
            ComRetry(Function() TryCast(inspector.WordEditor, Microsoft.Office.Interop.Word.Document))

            If wordEditor Is Nothing Then
                Return 0
            End If

            ' Selected text
            Dim selection As Microsoft.Office.Interop.Word.Selection = wordEditor.Application.Selection
            Dim selectedText As String = selection.Text

            If String.IsNullOrWhiteSpace(selectedText) Then
                Return 0
            End If

            ' Word splitting
            ' Dim words = selectedText.Split(New Char() {" "c, ControlChars.Tab, ControlChars.Cr, ControlChars.Lf},
            '                           StringSplitOptions.RemoveEmptyEntries)
            ' Return words.Length

            ' Count only real words:
            ' \b[\p{L}]+(?:['’\-‑–][\p{L}]+)*\b
            ' - one or more Unicode letters
            ' - optionally followed by groups of apostrophe/hyphen/dash + letters (e.g., don't, mother-in-law)
            ' - excludes digits and symbol-only tokens
            Dim pattern As String = "\b[\p{L}]+(?:['’\-‑–][\p{L}]+)*\b"
            Return Regex.Matches(selectedText, pattern).Count

        Catch ex As System.Exception  ' Explicitly referencing System.Exception per your guideline
            Return 0
        End Try
    End Function

End Class