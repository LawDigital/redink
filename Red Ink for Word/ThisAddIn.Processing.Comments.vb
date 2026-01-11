' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.Processing.Comments.vb
' Purpose: Integrates LLM-provided feedback with Word comments by creating,
'          replying to, extracting, and formatting comment bubbles (including
'          Markdown/HTML rendering and XML export utilities).
'
' Architecture:
'   - Comment Creation: Parses LLM responses, locates target text, and inserts
'     comments with optional Markdown rendering while suppressing UI flicker.
'   - Reply Automation: Resolves Word comment IDs/pseudo hashes and posts
'     threaded replies with markdown-aware formatting.
'   - Extraction: Filters existing comments by range/author/date and serializes
'     them as XML, including pseudo-stable identifiers.
'   - Formatting Utilities: Converts Markdown/HTML to the limited formatting
'     that Word comment balloons accept, including sanitization and fallbacks.
'   - UI Safety: Temporarily disables the Word window and turns off ScreenUpdating to reduce flicker and prevent Selection drift.
'   - Reply Fallback: Uses threaded replies when supported; otherwise appends the reply into the original comment text.
'   - Helpers: Provide safe position calculations, XML escaping, hashing, and
'     ID token parsing to keep interactions deterministic and auditable.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Data
Imports System.Diagnostics
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Wordprocessing
Imports Markdig
Imports Microsoft.Office.Interop.Word
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods

''' <summary>
''' Hosts helpers that transform LLM responses into Word comments, apply replies, format Markdown/HTML safely,
''' and export comment metadata for downstream processing.
''' </summary>
Partial Public Class ThisAddIn


    ''' <summary>
    ''' Parses an LLM bubble payload, locates each referenced text fragment, and inserts formatted Word comments while tracking failures.
    ''' </summary>
    ''' <param name="LLMResult">LLM output containing items separated by <c>"§§§"</c> and individual pairs <c>"find@@comment"</c>.</param>
    ''' <param name="Selection">Text selection that scopes the search anchors.</param>
    ''' <param name="DoSilent">When true, suppresses UI notifications.</param>
    ''' <param name="Prefix">Optional prefix prepended to every comment body.</param>
    ''' <remarks>
    ''' Errors (missing text, wrong format) are collected and optionally inserted as a summary comment at the tail of the selection.
    ''' When Markdown formatting is enabled, balloons are temporarily shown to allow Word to accept formatted insertion.
    ''' </remarks>
    Public Sub SetBubbles(LLMResult As System.String, Selection As Microsoft.Office.Interop.Word.Selection, DoSilent As System.Boolean, Optional Prefix As String = "")

        Dim responseItems() As System.String = LLMResult.Split(New System.String() {"§§§"}, System.StringSplitOptions.RemoveEmptyEntries)
        Dim wrongformatresponse As New System.Collections.Generic.List(Of System.String)
        Dim notfoundresponse As New System.Collections.Generic.List(Of System.String)

        ' Stable document reference
        Dim docRef As Microsoft.Office.Interop.Word.Document = Nothing
        Try
            If Selection IsNot Nothing AndAlso Selection.Range IsNot Nothing Then
                docRef = Selection.Range.Document
            End If
            If docRef Is Nothing AndAlso Globals.ThisAddIn IsNot Nothing AndAlso
           Globals.ThisAddIn.Application IsNot Nothing AndAlso
           Globals.ThisAddIn.Application.Documents.Count > 0 Then
                docRef = Globals.ThisAddIn.Application.ActiveDocument
            End If
        Catch ex As System.Exception
        End Try
        If docRef Is Nothing Then Exit Sub

        Dim originalRange As Microsoft.Office.Interop.Word.Range = Selection.Range.Duplicate
        Dim BubblecutHappened As System.Boolean = False
        Dim BubbleCount As System.Int32 = 0
        Dim MaxBubbles As System.Int32 = responseItems.Count

        If MaxBubbles = 0 Then
            If Not DoSilent Then ShowCustomMessageBox($"The bubble command did not result in any comment(s) by the LLM.")
            Exit Sub
        End If

        Dim splash As New SLib.SplashScreen($"Adding {MaxBubbles} bubble(s) to your text... press 'Esc' to abort")
        splash.Show()
        splash.Refresh()

        ' ─────────────────────────────────────────────────────────────────────
        ' Prevent user interaction and hide balloons while we run
        Dim app As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application
        Dim win As Microsoft.Office.Interop.Word.Window = app.ActiveWindow
        Dim hwnd As System.IntPtr = System.IntPtr.Zero
        Dim winWasEnabled As System.Boolean = True
        Dim prevShowRevs As System.Boolean = False
        Dim prevScreenUpdating As System.Boolean = True

        Try
            If win IsNot Nothing Then
                Try
                    hwnd = CType(win.Hwnd, System.IntPtr)
                Catch ex As System.Exception
                    hwnd = System.IntPtr.Zero
                End Try
                Try
                    prevShowRevs = win.View.ShowRevisionsAndComments
                    win.View.ShowRevisionsAndComments = False ' hide balloons/comments
                    win.View.RevisionsView = Microsoft.Office.Interop.Word.WdRevisionsView.wdRevisionsViewFinal
                Catch ex As System.Exception
                End Try
            End If

            Try
                prevScreenUpdating = app.ScreenUpdating
                app.ScreenUpdating = False
            Catch ex As System.Exception
            End Try

            If hwnd <> System.IntPtr.Zero Then
                Try
                    winWasEnabled = IsWindowEnabled(hwnd)
                Catch ex As System.Exception
                    winWasEnabled = True
                End Try
                Try
                    ' Disable the Word window so clicks cannot move the live Selection while ranges are being created
                    EnableWindow(hwnd, False)
                Catch ex As System.Exception
                End Try
            End If
            ' ─────────────────────────────────────────────────────────────────────

            Try
                For Each item As System.String In responseItems

                    splash.UpdateMessage($"Adding {MaxBubbles - BubbleCount} bubble(s) to your text... press 'Esc' to abort")
                    System.Windows.Forms.Application.DoEvents()

                    If (GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0 Then Exit For
                    If (GetAsyncKeyState(VK_ESCAPE) And 1) <> 0 Then Exit For

                    Dim parts() As System.String = item.Split(New System.String() {"@@"}, System.StringSplitOptions.None)
                    If parts.Length = 2 Then

                        Dim findText As System.String = parts(0).Trim().Trim("'"c).Trim(""""c)
                        Dim commentText As System.String = parts(1).Trim()

                        Dim viewChanged1 As Boolean = False
                        Dim viewChanged2 As Boolean = False
                        Dim view = app.ActiveWindow.View
                        Dim origRevView = view.RevisionsView
                        Dim origShowRev = view.ShowRevisionsAndComments

                        If view.RevisionsView <> Microsoft.Office.Interop.Word.WdRevisionsView.wdRevisionsViewFinal Then
                            view.RevisionsView = Microsoft.Office.Interop.Word.WdRevisionsView.wdRevisionsViewFinal
                            viewChanged1 = True
                        End If
                        If view.ShowRevisionsAndComments Then
                            view.ShowRevisionsAndComments = False
                            viewChanged2 = True
                        End If

                        Try
                            ' Ensure we start in MainText (best effort)
                            Try
                                If app.Selection IsNot Nothing AndAlso app.Selection.StoryType <> Microsoft.Office.Interop.Word.WdStoryType.wdMainTextStory Then
                                    app.Selection.SetRange(Start:=originalRange.Start, End:=originalRange.End)
                                End If
                            Catch exSet As System.Exception
                            End Try

                            ' === NO DoEvents between the finder and creating the stable range ===
                            If FindLongTextInChunks(findText, Selection) Then
                                Dim s As System.Int32 = -1, e As System.Int32 = -1
                                Try
                                    s = Selection.Range.Start : e = Selection.Range.End
                                Catch exGrab As System.Exception
                                    s = -1 : e = -1
                                End Try

                                If s >= 0 AndAlso e >= 0 Then
                                    Dim matchedRange As Microsoft.Office.Interop.Word.Range = Nothing
                                    Try
                                        matchedRange = docRef.Range(Start:=s, End:=e)
                                    Catch exR As System.Exception
                                        matchedRange = Nothing
                                    End Try

                                    If matchedRange IsNot Nothing AndAlso
                                   matchedRange.StoryType = Microsoft.Office.Interop.Word.WdStoryType.wdMainTextStory Then

                                        If INI_MarkdownBubbles Then
                                            Dim cmt As Microsoft.Office.Interop.Word.Comment = Nothing
                                            Try
                                                cmt = Globals.ThisAddIn.Application.ActiveDocument.Comments.Add(Range:=matchedRange, Text:="")
                                            Catch exAdd As System.Runtime.InteropServices.COMException
                                                cmt = Globals.ThisAddIn.Application.ActiveDocument.Comments.Add(Range:=matchedRange, Text:=$"{AN5}{Prefix}: ")
                                            Catch exAdd2 As System.Exception
                                                cmt = Globals.ThisAddIn.Application.ActiveDocument.Comments.Add(Range:=matchedRange, Text:=$"{AN5}{Prefix}: ")
                                            End Try

                                            Try
                                                If commentText.StartsWith("* ") Then
                                                    commentText = $"{AN5}{Prefix}:" & vbCrLf & commentText
                                                Else
                                                    commentText = $"{AN5}{Prefix}: " & commentText
                                                End If
                                                Dim cRng As Microsoft.Office.Interop.Word.Range = cmt.Range
                                                ' Ensure balloons visible during formatted insertion into the comment story
                                                Dim prevShow As System.Boolean = win.View.ShowRevisionsAndComments
                                                Try
                                                    win.View.ShowRevisionsAndComments = True
                                                    InsertMarkdownToComment(cRng, commentText) ' will not fight a hidden pane
                                                Finally
                                                    win.View.ShowRevisionsAndComments = prevShow
                                                End Try
                                            Catch exMk As System.Exception
                                                cmt.Range.Text = $"{AN5}{Prefix}: " & commentText
                                            End Try
                                        Else
                                            Globals.ThisAddIn.Application.ActiveDocument.Comments.Add(matchedRange, $"{AN5}{Prefix}: " & commentText)
                                        End If

                                        BubbleCount += 1
                                    Else
                                        notfoundresponse.Add("'" & findText & "' " & vbCrLf & ChrW(8594) & $" {AN5}{Prefix}: " & commentText & vbCrLf & vbCrLf)
                                    End If
                                Else
                                    notfoundresponse.Add("'" & findText & "' " & vbCrLf & ChrW(8594) & $" {AN5}{Prefix}: " & commentText & vbCrLf & vbCrLf)
                                End If
                            Else
                                notfoundresponse.Add("'" & findText & "' " & vbCrLf & ChrW(8594) & $" {AN5}{Prefix}: " & commentText & vbCrLf & vbCrLf)
                            End If

                        Catch ex As System.Exception
                            notfoundresponse.Add("'" & findText & "' " & vbCrLf & ChrW(8594) & $" {AN5}{Prefix}: " & commentText & " [Error: " & ex.Message & "]" & vbCrLf & vbCrLf)
                        End Try

                        If viewChanged1 Then view.RevisionsView = origRevView
                        If viewChanged2 Then view.ShowRevisionsAndComments = origShowRev

                    Else
                        If Not System.String.IsNullOrWhiteSpace(item) Then
                            wrongformatresponse.Add(item)
                        End If
                    End If

                    ' Best-effort restore
                    Try
                        Selection.SetRange(Start:=originalRange.Start, End:=originalRange.End)
                    Catch exRestore As System.Exception
                    End Try
                Next

            Finally
                'splash.Close()
            End Try

        Finally
            ' ─────────────────────────────────────────────────────────────────
            ' Restore Word window + UI
            Try
                If hwnd <> System.IntPtr.Zero Then EnableWindow(hwnd, winWasEnabled)
            Catch ex As System.Exception
            End Try
            Try
                If win IsNot Nothing Then win.View.ShowRevisionsAndComments = prevShowRevs
            Catch ex As System.Exception
            End Try
            Try
                If app IsNot Nothing Then app.ScreenUpdating = prevScreenUpdating
            Catch ex As System.Exception
            End Try
            ' ─────────────────────────────────────────────────────────────────

            If splash IsNot Nothing Then
                Try : splash.Close() : Catch : End Try
            End If

        End Try

        ' build ErrorList and add final summary comment…
        Dim ErrorList As System.String = ""

        If notfoundresponse.Count > 0 Then
            ErrorList += "The following comments could not be assigned to your text (they were not found, typically because of the LLM not acting as instructed, formatting or markup issues):" & vbCrLf & vbCrLf
            For Each itemNF As System.String In notfoundresponse
                If itemNF.Trim() <> "" Then ErrorList += Trim("- " & itemNF & vbCrLf)
            Next
            ErrorList += vbCrLf
        End If

        If wrongformatresponse.Count > 0 Then
            ErrorList += "The following responses could not be identified as bubble comments (typically because the LLM did not act as instructed):" & vbCrLf & vbCrLf
            For Each itemWF As System.String In wrongformatresponse
                If itemWF.Trim() <> "" Then ErrorList += Trim("- " & itemWF & vbCrLf)
            Next
            ErrorList += vbCrLf
        End If

        If Not System.String.IsNullOrWhiteSpace(ErrorList) Then
            If BubblecutHappened Then
                ErrorList = $"Some of the sections to which the bubble comments relate were too long for selecting. Only the initial part has been selected. This is indicated by '{BubbleCutText}' in the bubble comments, as applicable." & vbCrLf & vbCrLf & ErrorList
            End If
            If Not DoSilent Then
                ErrorList = ShowCustomWindow($"{BubbleCount} bubble comment(s) applied (Warning: complicated formatting and markups may cause misalignments of the commented portions of the text). The following errors occurred when implementing the 'bubbles' feedback of the LLM:", ErrorList, "The above error list will be included in a final comment at the end of your selection (it will also be included in the clipboard). You can have the original list included, or you can now make changes and have this version used. If you select Cancel, nothing will be put added to the document.", AN, True)
            End If
            If ErrorList <> "" AndAlso ErrorList.ToLower() <> "esc" Then
                SLib.PutInClipboard(ErrorList)

                Dim tailRange As Microsoft.Office.Interop.Word.Range = originalRange.Duplicate
                tailRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)

                If INI_MarkdownBubbles Then
                    Try
                        Dim cmt As Microsoft.Office.Interop.Word.Comment = docRef.Comments.Add(Range:=tailRange, Text:="")
                        Dim cRng As Microsoft.Office.Interop.Word.Range = cmt.Range
                        InsertMarkdownToComment(cRng, $"{AN5}{Prefix}:" & ErrorList)
                    Catch exMkSum As System.Exception
                        docRef.Comments.Add(Range:=tailRange, Text:=$"{AN5}{Prefix}: " & ErrorList)
                    End Try
                Else
                    docRef.Comments.Add(Range:=tailRange, Text:=$"{AN5}{Prefix}: " & ErrorList)
                End If
            End If
        Else
            If Not DoSilent Then
                ShowCustomMessageBox($"{BubbleCount} bubble comment(s) provided by the LLM applied to to your text (Warning: complicated formatting and markups may cause misalignments of the commented portions of the text)." &
                                 If(BubblecutHappened, $"Some of the sections to which the bubble comments relate were too long for selecting. Only the initial part has been selected. This is indicated by '{BubbleCutText}' in the bubble comments, as applicable.", ""))
            End If
        End If

    End Sub

    ''' <summary>
    ''' Parses LLM reply instructions, resolves target comments by ID/hash, and posts threaded replies or records failures.
    ''' </summary>
    ''' <param name="LLMResult">Reply payload split by <c>"§§§"</c> with each entry formatted as <c>"token@@reply"</c>.</param>
    ''' <param name="Selection">Selection used to keep focus inside the main story range.</param>
    ''' <param name="DoSilent">When true, suppresses message boxes.</param>
    ''' <param name="Prefix">Optional label prepended to reply content.</param>
    ''' <remarks>Each token can be a Word comment index, pseudo hash, or labeled pair as described in the summary block.</remarks>
    Public Sub ReplyBubbles(ByVal LLMResult As String,
                        ByVal Selection As Microsoft.Office.Interop.Word.Selection,
                        ByVal DoSilent As Boolean, Optional Prefix As String = "")

        If String.IsNullOrWhiteSpace(LLMResult) Then
            If Not DoSilent Then ShowCustomMessageBox("No replies found in the LLM result.")
            Exit Sub
        End If

        ' Resolve active Word objects (application/document/window)
        Dim app As Microsoft.Office.Interop.Word.Application = Nothing
        Dim docRef As Microsoft.Office.Interop.Word.Document = Nothing
        Dim win As Microsoft.Office.Interop.Word.Window = Nothing
        Dim hwnd As IntPtr = IntPtr.Zero

        Dim hadSel As Boolean = False
        Dim origStart As Integer = -1
        Dim origEnd As Integer = -1

        Dim prevScreenUpdating As Boolean = True
        Dim winWasEnabled As Boolean = True

        ' Split into individual reply items
        Dim items As String() = LLMResult.Split(New String() {"§§§"}, StringSplitOptions.RemoveEmptyEntries)
        If items.Length = 0 Then
            If Not DoSilent Then ShowCustomMessageBox("No replies found in the LLM result.")
            Exit Sub
        End If

        ' Collect errors
        Dim wrongformatresponse As New List(Of String)()
        Dim notfoundresponse As New List(Of String)()

        Try
            app = Globals.ThisAddIn.Application
            If app Is Nothing OrElse app.Documents Is Nothing OrElse app.Documents.Count = 0 Then
                If Not DoSilent Then ShowCustomMessageBox("No active Word document.")
                Exit Sub
            End If

            docRef = app.ActiveDocument
            win = app.ActiveWindow

            ' Capture original selection
            Try
                If Selection IsNot Nothing Then
                    origStart = Selection.Start
                    origEnd = Selection.End
                    hadSel = True
                Else
                    ' Fall back to active selection
                    If app.Selection IsNot Nothing Then
                        origStart = app.Selection.Start
                        origEnd = app.Selection.End
                        hadSel = True
                    End If
                End If
            Catch
            End Try

            ' Clamp selection to main story bounds
            Dim safeStart As Integer = docRef.Content.Start
            Dim safeEnd As Integer = docRef.Content.End
            If hadSel Then
                safeStart = System.Math.Max(docRef.Content.Start, System.Math.Min(origStart, docRef.Content.End))
                safeEnd = System.Math.Max(docRef.Content.Start, System.Math.Min(origEnd, docRef.Content.End))
                If safeEnd < safeStart Then safeEnd = safeStart
            End If

            ' Minimize UI interaction while processing
            Try
                prevScreenUpdating = app.ScreenUpdating
                app.ScreenUpdating = False
            Catch
            End Try

            If win IsNot Nothing Then
                Try
                    hwnd = CType(win.Hwnd, IntPtr)
                    winWasEnabled = IsWindowEnabled(hwnd)
                    EnableWindow(hwnd, False)
                Catch
                End Try
            End If

            ' Process each reply item
            Dim addedCount As Integer = 0
            For Each raw In items
                ' ESC to abort
                If (GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0 OrElse (GetAsyncKeyState(VK_ESCAPE) And 1) <> 0 Then Exit For

                Dim item As String = If(raw, String.Empty).Trim()
                If item.Length = 0 Then Continue For

                ' Split into "token@@reply"
                Dim sepIdx As Integer = item.IndexOf("@@")
                If sepIdx <= 0 OrElse sepIdx >= item.Length - 2 Then
                    wrongformatresponse.Add(item)
                    Continue For
                End If

                Dim idToken As String = item.Substring(0, sepIdx).Trim()
                Dim replyText As String = item.Substring(sepIdx + 2).Trim()

                If idToken.Length = 0 OrElse replyText.Length = 0 Then
                    wrongformatresponse.Add(item)
                    Continue For
                End If

                ' Parse token into (wordId, hash)
                Dim wordId As Integer? = Nothing
                Dim pseudoHash As String = Nothing
                If Not TryParseCommentIdToken(idToken, wordId, pseudoHash) Then
                    wrongformatresponse.Add(item)
                    Continue For
                End If

                replyText = AN5 & ": " & replyText

                ' Add reply using existing helper
                Dim formatted As Boolean = INI_MarkdownBubbles ' follow same setting as SetBubbles
                Dim ok As Boolean = False
                Try
                    ok = ReplyToWordComment(wordId, pseudoHash, replyText, formatted)
                Catch ex As Exception
                    ok = False
                End Try

                If ok Then
                    addedCount += 1
                Else
                    notfoundresponse.Add($"'{idToken}'{vbCrLf}{ChrW(&H2192)} reply: {replyText}")
                End If

                ' Ensure focus stays on main story (avoid leaving caret in a comment card)
                Try
                    app.ActiveWindow.Activate()
                    docRef.Range(safeStart, safeEnd).Select()
                Catch
                End Try
            Next

            ' Build and optionally show summary
            Dim errList As String = ""
            If notfoundresponse.Count > 0 Then
                errList += "The following replies could not be added (target comment not found):" & vbCrLf & vbCrLf
                For Each nf In notfoundresponse
                    If nf.Trim() <> "" Then errList += "- " & nf & vbCrLf
                Next
                errList += vbCrLf
            End If
            If wrongformatresponse.Count > 0 Then
                errList += "The following items did not match the expected 'token@@reply' format:" & vbCrLf & vbCrLf
                For Each wf In wrongformatresponse
                    If wf.Trim() <> "" Then errList += "- " & wf & vbCrLf
                Next
                errList += vbCrLf
            End If

            If String.IsNullOrWhiteSpace(errList) Then
                If Not DoSilent Then
                    ShowCustomMessageBox($"{addedCount} reply/replies added.")
                End If
            Else
                If Not DoSilent Then
                    Dim final As String = ShowCustomWindow($"{addedCount} reply/replies added. Some issues occurred:", errList, "You can edit the error list before it is inserted as a final comment at the end of your selection. Cancel to skip insertion.", AN, True)
                    If final <> "" AndAlso final.ToLower() <> "esc" Then
                        ' Insert summary comment at end of original selection (or at document end if none)
                        Dim tail As Microsoft.Office.Interop.Word.Range =
                        If(hadSel, docRef.Range(safeStart, safeEnd).Duplicate, docRef.Content.Duplicate)
                        tail.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
                        Try
                            If INI_MarkdownBubbles Then
                                Dim cmt As Microsoft.Office.Interop.Word.Comment = docRef.Comments.Add(Range:=tail, Text:="")
                                InsertMarkdownToComment(cmt.Range, $"{AN5}{Prefix}: " & final)
                            Else
                                docRef.Comments.Add(Range:=tail, Text:=$"{AN5}{Prefix}: " & final)
                            End If
                        Catch
                            docRef.Comments.Add(Range:=tail, Text:=$"{AN5}{Prefix}: " & final)
                        End Try
                    End If
                End If
            End If

        Finally
            ' Restore UI and selection
            Try
                If win IsNot Nothing AndAlso hwnd <> IntPtr.Zero Then
                    EnableWindow(hwnd, winWasEnabled)
                End If
            Catch
            End Try
            Try
                If app IsNot Nothing Then app.ScreenUpdating = prevScreenUpdating
            Catch
            End Try

            ' Force selection back into main text story (avoid comment card focus)
            Try
                If docRef IsNot Nothing Then
                    If hadSel Then
                        Dim s As Integer = System.Math.Max(docRef.Content.Start, System.Math.Min(origStart, docRef.Content.End))
                        Dim e As Integer = System.Math.Max(docRef.Content.Start, System.Math.Min(origEnd, docRef.Content.End))
                        If e < s Then e = s
                        docRef.Range(s, e).Select()
                    Else
                        docRef.Range(docRef.Content.End, docRef.Content.End).Select()
                    End If
                End If
            Catch
            End Try
        End Try
    End Sub

    ''' <summary>
    ''' Attempts to extract a Word comment index and/or pseudo hash from a flexible identifier string.
    ''' </summary>
    ''' <param name="raw">Input containing tokens such as <c>"wid:123 ph:abcdef"</c>, <c>"1234|abcdef"</c>, or a single value.</param>
    ''' <param name="wordId">Outputs the parsed Word comment index when present.</param>
    ''' <param name="pseudoHash">Outputs the parsed pseudo hash identifier when present.</param>
    ''' <returns><c>True</c> when at least one identifier was resolved; otherwise <c>False</c>.</returns>
    ''' <remarks>
    ''' The single-token fallback treats a pure numeric token as a Word comment index and treats a token of length >= 6 as a pseudo hash.
    ''' </remarks>
    Private Function TryParseCommentIdToken(ByVal raw As String,
                                              ByRef wordId As System.Nullable(Of Integer),
                                              ByRef pseudoHash As String) As Boolean
        wordId = Nothing
        pseudoHash = Nothing
        If String.IsNullOrWhiteSpace(raw) Then Return False

        Dim s As String = raw.Trim()

        ' 1) "id|hash"
        Dim pipeParts = s.Split(New Char() {"|"c}, 2, StringSplitOptions.None)
        If pipeParts.Length = 2 Then
            Dim left = pipeParts(0).Trim()
            Dim right = pipeParts(1).Trim()
            Dim idVal As Integer
            If Integer.TryParse(left, idVal) Then wordId = idVal
            If Not String.IsNullOrWhiteSpace(right) Then pseudoHash = right
            Return (wordId.HasValue OrElse Not String.IsNullOrWhiteSpace(pseudoHash))
        End If

        ' 2) labeled forms (wid/id and ph/hash/pseudohash)
        Dim idMatch = System.Text.RegularExpressions.Regex.Match(s, "(?:\bwid|\bid|\bwordid)\s*[:=]\s*(?<id>-?\d+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        If idMatch.Success Then
            Dim idVal As Integer
            If Integer.TryParse(idMatch.Groups("id").Value, idVal) Then wordId = idVal
        End If
        Dim hashMatch = System.Text.RegularExpressions.Regex.Match(s, "(?:\bph|\bhash|\bpseudohash)\s*[:=]\s*(?<hash>[A-Za-z0-9_-]{6,})", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        If hashMatch.Success Then
            pseudoHash = hashMatch.Groups("hash").Value.Trim()
        End If
        If wordId.HasValue OrElse Not String.IsNullOrWhiteSpace(pseudoHash) Then Return True

        ' 3) single-token fallback
        Dim onlyDigits As Boolean = True
        For Each ch As Char In s
            If Not Char.IsDigit(ch) Then
                onlyDigits = False
                Exit For
            End If
        Next

        If onlyDigits Then
            Dim idVal As Integer
            If Integer.TryParse(s, idVal) Then wordId = idVal
            Return True
        Else
            If s.Length >= 6 Then
                pseudoHash = s
                Return True
            End If
        End If

        Return False
    End Function

    ''' <summary>
    ''' Converts Markdown or HTML into Word-safe formatting and inserts the result into a comment range with sanitization.
    ''' </summary>
    ''' <param name="rg">Target Word range inside a comment bubble.</param>
    ''' <param name="src">Markdown or HTML source to render.</param>
    ''' <remarks>On conversion failure the content falls back to a plain text representation.</remarks>
    Public Shared Sub InsertMarkdownToComment(ByRef rg As Word.Range,
                                                 ByVal src As String)
        Try
            ' Heuristic: is this HTML already?
            Dim looksLikeHtml As Boolean =
            (src IsNot Nothing AndAlso src.IndexOf("<"c) >= 0 AndAlso src.IndexOf(">"c) > src.IndexOf("<"c))

            ' 1) Get HTML (Markdown -> HTML if needed)
            Dim html As String
            If looksLikeHtml Then
                html = src
            Else
                Dim pipeline = New MarkdownPipelineBuilder().UseAdvancedExtensions().Build()
                html = Markdig.Markdown.ToHtml(If(src, String.Empty), pipeline)
            End If

            ' 2) Load/sanitize
            Dim hdoc As New HtmlAgilityPack.HtmlDocument()
            hdoc.LoadHtml(If(html, String.Empty))

            ' Comments in modern Word only allow a small subset of formatting.
            RemoveNodesIfPresent(hdoc, "//script|//style|//img")
            FlattenTables(hdoc)              ' turn tables into plain text block(s)
            RemoveTrailingParagraph(hdoc)       ' cosmetic trim

            ParseHtmlNode(hdoc.DocumentNode, rg)

        Catch
            ' Last-resort fallback: plain text
            Dim fallback As String = SafePlainFromMarkdownOrHtml(src)
            rg.Text = fallback
        End Try
    End Sub

    ''' <summary>
    ''' Removes nodes that match the provided XPath to ensure only supported markup reaches Word comments.
    ''' </summary>
    ''' <param name="hdoc">HTML document to prune.</param>
    ''' <param name="xpath">XPath selecting disallowed elements.</param>
    Private Shared Sub RemoveNodesIfPresent(hdoc As HtmlAgilityPack.HtmlDocument, xpath As String)
        Dim nodes = hdoc.DocumentNode.SelectNodes(xpath)
        If nodes Is Nothing Then Return
        For Each n In nodes
            n.Remove()
        Next
    End Sub

    ''' <summary>
    ''' Flattens each table into tab-delimited rows so that table content remains readable in balloons.
    ''' </summary>
    ''' <param name="hdoc">HTML document that may contain tables.</param>
    Private Shared Sub FlattenTables(hdoc As HtmlAgilityPack.HtmlDocument)
        Dim tables = hdoc.DocumentNode.SelectNodes("//table")
        If tables Is Nothing Then Return

        For Each t In tables
            Dim lines As New List(Of String)

            Dim rows = t.SelectNodes(".//tr")
            If rows Is Nothing Then
                ' No rows found; remove the table to avoid leaving unsupported markup behind
                t.Remove()
                Continue For
            End If

            For Each tr In rows
                Dim cells = tr.SelectNodes("./th|./td")
                If cells Is Nothing Then Continue For
                Dim vals = cells.Select(Function(c) HtmlAgilityPack.HtmlEntity.DeEntitize(c.InnerText).Trim())
                lines.Add(String.Join(vbTab, vals))
            Next

            Dim repl = hdoc.CreateTextNode(String.Join(vbCrLf, lines))
            t.ParentNode.ReplaceChild(repl, t)
        Next
    End Sub

    ''' <summary>
    ''' Provides a Markdown/HTML to plain-text fallback that preserves readable content when rich rendering fails.
    ''' </summary>
    ''' <param name="src">Original Markdown or HTML input.</param>
    ''' <returns>Plain text representation of the source.</returns>
    Private Shared Function SafePlainFromMarkdownOrHtml(src As String) As String
        Try
            Dim looksLikeHtml As Boolean =
            CBool(src?.IndexOf("<"c) >= 0 AndAlso src.IndexOf(">"c) > src.IndexOf("<"c))
            Dim html As String
            If looksLikeHtml Then
                html = src
            Else
                Dim pipeline = New MarkdownPipelineBuilder().UseAdvancedExtensions().Build()
                html = Markdig.Markdown.ToHtml(If(src, String.Empty), pipeline)
            End If
            Dim hdoc As New HtmlAgilityPack.HtmlDocument()
            hdoc.LoadHtml(html)
            Return HtmlAgilityPack.HtmlEntity.DeEntitize(hdoc.DocumentNode.InnerText).Trim()
        Catch
            Return If(src, String.Empty)
        End Try
    End Function

    ''' <summary>
    ''' Extracts comments within a range, optionally filters by author/date, and emits an XML document describing each bubble.
    ''' </summary>
    ''' <param name="rng">Range limiting extraction; entire document is used when <c>Nothing</c> or zero-length.</param>
    ''' <param name="Silent">When true, suppresses user prompts and notifications.</param>
    ''' <param name="SortByDate">Sorts output by date when true; otherwise by document order.</param>
    ''' <returns>XML payload containing summary metadata and comment nodes, or an empty string if none were exported.</returns>
    Public Shared Function BubblesExtract(ByVal rng As Microsoft.Office.Interop.Word.Range,
                                      Optional ByVal Silent As Boolean = False, Optional ByVal SortByDate As Boolean = False) As System.String

        Dim app As Microsoft.Office.Interop.Word.Application = Nothing
        Dim doc As Microsoft.Office.Interop.Word.Document = Nothing

        Try
            ' Get Word Application
            Try
                Try
                    app = CType(System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application"), Microsoft.Office.Interop.Word.Application)
                Catch ex As System.Exception
                    app = Globals.ThisAddIn.Application
                End Try
            Catch ex As System.Exception
                If Not Silent Then ShowCustomMessageBox("Unable to access Word Application instance.")
                Return ""
            End Try

            ' Get Active Document
            Try
                doc = app.ActiveDocument
            Catch ex As System.Exception
                If Not Silent Then ShowCustomMessageBox("No active document found.")
                Return ""
            End Try
            If doc Is Nothing Then
                If Not Silent Then ShowCustomMessageBox("No active document found.")
                Return ""
            End If

            ' Rule 1: If range is Nothing or zero-length, use the whole document.
            Dim effectiveRange As Microsoft.Office.Interop.Word.Range
            If rng Is Nothing OrElse rng.Start = rng.End Then
                effectiveRange = doc.Content
            Else
                effectiveRange = rng
            End If

            ' Collect comments fully within effectiveRange
            Dim allInRange As System.Collections.Generic.List(Of Microsoft.Office.Interop.Word.Comment) =
            New System.Collections.Generic.List(Of Microsoft.Office.Interop.Word.Comment)()

            Try
                For Each c As Microsoft.Office.Interop.Word.Comment In doc.Comments
                    Dim cStart As System.Int32
                    Dim cEnd As System.Int32
                    Try
                        cStart = c.Scope.Start
                        cEnd = c.Scope.End
                    Catch
                        Try
                            cStart = c.Reference.Start
                            cEnd = c.Reference.End
                        Catch
                            Continue For
                        End Try
                    End Try
                    If cStart >= effectiveRange.Start AndAlso cEnd <= effectiveRange.End Then
                        allInRange.Add(c)
                    End If
                Next
            Catch ex As System.Exception
                If Not Silent Then ShowCustomMessageBox("Failed to collect comments from the document.")
                Return ""
            End Try

            If allInRange.Count = 0 Then
                'If Not Silent Then ShowCustomMessageBox("No comments found in the selected range (or document).")
                Return ""
            End If

            ' Build author and date sets (date grouped by day)
            Dim authors As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
            Dim dateSet As New System.Collections.Generic.HashSet(Of System.DateTime)()

            For Each c In allInRange
                Try
                    Dim a As System.String = If(c.Author, System.String.Empty)
                    If Not System.String.IsNullOrWhiteSpace(a) Then authors.Add(a)
                Catch
                End Try
                Try
                    Dim d As System.DateTime = c.Date.Date
                    If d <> System.DateTime.MinValue Then dateSet.Add(d)
                Catch
                End Try
            Next

            ' Only prompt when not silent
            Dim needPrompt As System.Boolean = (Not Silent AndAlso (authors.Count > 1 OrElse dateSet.Count > 1))

            Dim selectedAuthor As System.String = "(All authors)"
            Dim fromDateChoice As System.String = "(All dates)"
            Dim toDateChoice As System.String = "(All dates)"

            If needPrompt Then
                ' Author options
                Dim authorOptions As New System.Collections.Generic.List(Of System.String)()
                authorOptions.Add("(All authors)")
                For Each a In authors
                    authorOptions.Add(a)
                Next
                authorOptions.Sort(System.StringComparer.CurrentCultureIgnoreCase)

                ' Date options (descending by date)
                Dim dateList As New System.Collections.Generic.List(Of System.DateTime)(dateSet)
                dateList.Sort() : dateList.Reverse()
                Dim dateOptions As New System.Collections.Generic.List(Of System.String)()
                dateOptions.Add("(All dates)")
                Dim now As System.DateTime = System.DateTime.Now
                For Each d In dateList
                    Dim daysAgo As System.Int32 = System.Math.Max(0, CInt((now.Date - d.Date).TotalDays))
                    dateOptions.Add($"{d:yyyy-MM-dd} ({daysAgo} days ago)")
                Next

                Dim prmAuthor As New SLib.InputParameter() With {
                .Name = "Author",
                .Value = "(All authors)",
                .Options = New System.Collections.Generic.List(Of System.String)(authorOptions)
            }
                Dim prmFrom As New SLib.InputParameter() With {
                .Name = "From date",
                .Value = "(All dates)",
                .Options = New System.Collections.Generic.List(Of System.String)(dateOptions)
            }
                Dim prmTo As New SLib.InputParameter() With {
                .Name = "To date",
                .Value = "(All dates)",
                .Options = New System.Collections.Generic.List(Of System.String)(dateOptions)
            }

                Dim parameters As SLib.InputParameter() = {prmAuthor, prmFrom, prmTo}

                Dim ok As System.Boolean = False
                Try
                    ok = SLib.ShowCustomVariableInputForm(
                    "Select an author (or keep 'All authors') and optional date range (From/To). If only 'To' is set, all comments up to and including that date will be selected. If only 'From' is set, all comments from that date onward will be selected.",
                    $"{AN} Filter Word Comments",
                    parameters
                )
                Catch ex As System.Exception
                    If Not Silent Then ShowCustomMessageBox("Failed to open the filter dialog.")
                    ok = False
                End Try

                If Not ok Then
                    Dim proceed As Integer = SLib.ShowCustomYesNoBox("No selection made. Do you want to continue with all authors and all dates?", "Yes", "No")
                    If proceed <> 1 Then
                        Return ""
                    End If
                Else
                    selectedAuthor = System.Convert.ToString(prmAuthor.Value)
                    fromDateChoice = System.Convert.ToString(prmFrom.Value)
                    toDateChoice = System.Convert.ToString(prmTo.Value)
                    If System.String.IsNullOrWhiteSpace(selectedAuthor) Then selectedAuthor = "(All authors)"
                    If System.String.IsNullOrWhiteSpace(fromDateChoice) Then fromDateChoice = "(All dates)"
                    If System.String.IsNullOrWhiteSpace(toDateChoice) Then toDateChoice = "(All dates)"
                End If
            End If

            ' Parse "yyyy-MM-dd (…)” labels
            Dim parseDateFromLabel As System.Func(Of System.String, System.Nullable(Of System.DateTime)) =
            Function(lbl As System.String) As System.Nullable(Of System.DateTime)
                If System.String.Equals(lbl, "(All dates)", System.StringComparison.OrdinalIgnoreCase) Then Return Nothing
                Dim spaceIdx As System.Int32 = lbl.IndexOf(" "c)
                If spaceIdx <= 0 Then Return Nothing
                Dim datePart As System.String = lbl.Substring(0, spaceIdx)
                Dim d As System.DateTime
                If System.DateTime.TryParseExact(datePart, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, d) Then
                    Return d.Date
                End If
                Return Nothing
            End Function

            Dim fromDate As System.Nullable(Of System.DateTime) = parseDateFromLabel(fromDateChoice)
            Dim toDate As System.Nullable(Of System.DateTime) = parseDateFromLabel(toDateChoice)

            ' Apply filters
            Dim filtered As System.Collections.Generic.IEnumerable(Of Microsoft.Office.Interop.Word.Comment) = allInRange

            If Not System.String.Equals(selectedAuthor, "(All authors)", System.StringComparison.OrdinalIgnoreCase) Then
                filtered = filtered.Where(Function(c)
                                              Dim a As System.String = System.String.Empty
                                              Try : a = If(c.Author, System.String.Empty) : Catch : End Try
                                              Return System.String.Equals(a, selectedAuthor, System.StringComparison.OrdinalIgnoreCase)
                                          End Function)
            End If

            If fromDate.HasValue OrElse toDate.HasValue Then
                filtered = filtered.Where(Function(c)
                                              Dim d As System.DateTime
                                              Try : d = c.Date.Date : Catch : Return False : End Try
                                              If fromDate.HasValue AndAlso toDate.HasValue Then
                                                  Return d >= fromDate.Value.Date AndAlso d <= toDate.Value.Date
                                              ElseIf fromDate.HasValue Then
                                                  Return d >= fromDate.Value.Date
                                              ElseIf toDate.HasValue Then
                                                  Return d <= toDate.Value.Date
                                              Else
                                                  Return True
                                              End If
                                          End Function)
            End If

            Dim finalList As System.Collections.Generic.List(Of Microsoft.Office.Interop.Word.Comment)

            If SortByDate Then

                finalList =
            filtered.OrderBy(Function(c)
                                 Dim d As System.DateTime = System.DateTime.MinValue
                                 Try : d = c.Date : Catch : End Try
                                 Return d
                             End Function).
                     ThenBy(Function(c)
                                Try
                                    Return c.Scope.Start
                                Catch
                                    Return System.Int32.MaxValue
                                End Try
                            End Function).
                     ToList()

            Else
                ' Sort strictly by order of appearance in the document
                finalList = filtered.OrderBy(Function(c) CommentStartPos(c)).ThenBy(Function(c) CommentEndPos(c)).ToList()

            End If


            If finalList.Count = 0 Then
                If Not Silent Then ShowCustomMessageBox("No comments matched the selected filters.")
                Return ""
            End If

            ' Build XML
            Dim sb As New System.Text.StringBuilder()
            sb.AppendLine("<WORDBUBBLES>")
            sb.Append("  <Summary ")
            sb.Append($"documentName=""{xmlEscapeSafe(doc.Name)}"" ")
            sb.Append($"rangeStart=""{effectiveRange.Start}"" ")
            sb.Append($"rangeEnd=""{effectiveRange.End}"" ")
            If selectedAuthor <> "" Then sb.Append($"authorFilter=""{xmlEscapeSafe(selectedAuthor)}"" ")
            If fromDateChoice <> "" Then sb.Append($"fromFilter=""{xmlEscapeSafe(fromDateChoice)}"" ")
            If toDateChoice <> "" Then sb.Append($"toFilter=""{xmlEscapeSafe(toDateChoice)}"" ")
            sb.AppendLine($"/>")

            For Each c In finalList
                Dim author As System.String = System.String.Empty
                Dim initials As System.String = System.String.Empty
                Dim dateStr As System.String = System.String.Empty
                Dim text As System.String = System.String.Empty
                Dim referencedText As System.String = System.String.Empty
                Dim commentIndex As System.Int32 = -1

                Try : author = If(c.Author, System.String.Empty) : Catch : End Try
                Try : initials = If(c.Initial, System.String.Empty) : Catch : End Try
                Try : dateStr = c.Date.ToString("yyyy-MM-ddTHH:mm:ssK", System.Globalization.CultureInfo.InvariantCulture) : Catch : End Try
                Try : text = If(c.Range.Text, System.String.Empty) : Catch : End Try
                Try : referencedText = If(c.Scope.Text, System.String.Empty) : Catch : End Try
                Try : commentIndex = c.Index : Catch : End Try

                ' Pseudo-stable ID derived from comment metadata/content (position independent; changes if those inputs change)
                Dim idMaterial As System.String = author & "|" & initials & "|" & dateStr & "|" & referencedText & "|" & text
                Dim stableId As System.String = ComputeSha256Hex(idMaterial)

                sb.AppendLine("  <Comment " &
                          $"id=""{xmlEscapeSafe(stableId)}"" " &
                          $"wordIndex=""{commentIndex}"" " &
                          $"author=""{xmlEscapeSafe(author)}"" " &
                          $"initials=""{xmlEscapeSafe(initials)}"" " &
                          $"date=""{xmlEscapeSafe(dateStr)}"">")
                sb.AppendLine($"    <ScopeText>{xmlEscapeSafe(referencedText)}</ScopeText>")
                sb.AppendLine($"    <Content>{xmlEscapeSafe(text)}</Content>")
                sb.AppendLine("  </Comment>")
            Next

            sb.AppendLine("</WORDBUBBLES>")
            Return sb.ToString()

        Catch ex As System.Exception
            If Not Silent Then ShowCustomMessageBox($"Unexpected error: {ex.Message}")
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Retrieves a safe starting position for a comment, falling back to the reference range when necessary.
    ''' </summary>
    ''' <param name="c">Comment of interest.</param>
    ''' <returns>Start position or <see cref="Integer.MaxValue"/> if unavailable.</returns>
    Private Shared Function CommentStartPos(c As Microsoft.Office.Interop.Word.Comment) As Integer
        Try
            Return c.Scope.Start
        Catch
            Try
                Return c.Reference.Start
            Catch
                Return Integer.MaxValue
            End Try
        End Try
    End Function

    ''' <summary>
    ''' Retrieves a safe ending position for a comment, falling back to the reference range when necessary.
    ''' </summary>
    ''' <param name="c">Comment of interest.</param>
    ''' <returns>End position or <see cref="Integer.MaxValue"/> if unavailable.</returns>
    Private Shared Function CommentEndPos(c As Microsoft.Office.Interop.Word.Comment) As Integer
        Try
            Return c.Scope.End
        Catch
            Try
                Return c.Reference.End
            Catch
                Return Integer.MaxValue
            End Try
        End Try
    End Function

    ''' <summary>
    ''' Attempts to reply to a Word comment using its numeric ID or pseudo hash, optionally formatting the reply as Markdown/HTML.
    ''' </summary>
    ''' <param name="wordId">Word comment index (<c>c.Index</c>) if available.</param>
    ''' <param name="pseudoHashId">Content-derived pseudo hash identifier.</param>
    ''' <param name="replyText">Reply body (already prefixed).</param>
    ''' <param name="formatted">When true, renders reply via <see cref="InsertMarkdownToComment"/>; otherwise inserts plain text.</param>
    ''' <returns><c>True</c> when a threaded reply or fallback insertion succeeded; otherwise <c>False</c>.</returns>
    Public Shared Function ReplyToWordComment(ByVal wordId As System.Nullable(Of System.Int32),
                                              ByVal pseudoHashId As System.String,
                                              ByVal replyText As System.String,
                                              ByVal formatted As System.Boolean) As System.Boolean
        Dim app As Microsoft.Office.Interop.Word.Application = Nothing
        Dim doc As Microsoft.Office.Interop.Word.Document = Nothing

        Try
            Try
                app = CType(System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application"), Microsoft.Office.Interop.Word.Application)
            Catch ex As System.Exception
                app = Globals.ThisAddIn.Application
            End Try
        Catch ex As System.Exception
            Debug.WriteLine("ReplyToWordComment: Unable to access Word Application instance.")
            Return False
        End Try

        Try
            doc = app.ActiveDocument
        Catch ex As System.Exception
            Debug.WriteLine("ReplyToWordComment: No active document found.")
            Return False
        End Try
        If doc Is Nothing Then
            ShowCustomMessageBox("ReplyToWordComment: No active document found.")
            Return False
        End If

        If System.String.IsNullOrEmpty(replyText) Then
            Debug.WriteLine("ReplyToWordComment: Reply text is empty.")
            Return False
        End If

        ' 1) Try locate by Word-ID (Index)
        Dim target As Microsoft.Office.Interop.Word.Comment = Nothing
        If wordId.HasValue Then
            For Each c As Microsoft.Office.Interop.Word.Comment In doc.Comments
                Try
                    If c.Index = wordId.Value Then
                        target = c
                        Exit For
                    End If
                Catch
                End Try
            Next
        End If

        ' 2) Fallback: locate by pseudo hash id (content-based)
        If target Is Nothing AndAlso Not System.String.IsNullOrWhiteSpace(pseudoHashId) Then
            For Each c As Microsoft.Office.Interop.Word.Comment In doc.Comments
                Try
                    Dim author As System.String = If(c.Author, System.String.Empty)
                    Dim initials As System.String = If(c.Initial, System.String.Empty)
                    Dim dateStr As System.String = c.Date.ToString("yyyy-MM-ddTHH:mm:ssK", System.Globalization.CultureInfo.InvariantCulture)
                    Dim text As System.String = If(c.Range.Text, System.String.Empty)
                    Dim referencedText As System.String = System.String.Empty
                    Try : referencedText = If(c.Scope.Text, System.String.Empty) : Catch : referencedText = System.String.Empty : End Try

                    Dim idMaterial As System.String = author & "|" & initials & "|" & dateStr & "|" & referencedText & "|" & text
                    Dim stableId As System.String = ComputeSha256Hex(idMaterial)

                    If System.String.Equals(stableId, pseudoHashId, System.StringComparison.OrdinalIgnoreCase) Then
                        target = c
                        Exit For
                    End If
                Catch
                End Try
            Next
        End If

        If target Is Nothing Then
            Debug.WriteLine("ReplyToWordComment: Target comment not found by Word-ID or pseudo hash id.")
            Return False
        End If

        ' 3) Try to create a threaded reply if available, else fall back to appending inside the original comment range.
        Try
            ' Preferred path: threaded reply (Word 2013+)
            Dim newReply As Microsoft.Office.Interop.Word.Comment = Nothing
            Try
                ' Use the same anchor scope for the reply; text will be set via Range below.
                newReply = target.Replies.Add(target.Scope, System.String.Empty)
            Catch
                ' Some versions may require a reference range; try Reference
                Try
                    newReply = target.Replies.Add(target.Reference, System.String.Empty)
                Catch
                    newReply = Nothing
                End Try
            End Try

            If newReply IsNot Nothing Then
                If formatted Then
                    ' Use provided formatter
                    InsertMarkdownToComment(newReply.Range, replyText)
                Else
                    newReply.Range.Text = replyText
                End If
                Return True
            End If
        Catch ex As System.Exception
            ' fall through to non-threaded fallback
        End Try

        ' 4) Fallback: append "reply" inside the original comment content (non-threaded Word)
        Try
            Dim sep As System.String = System.Environment.NewLine & System.Environment.NewLine & "--- Reply ---" & System.Environment.NewLine
            If formatted Then
                ' Insert separator as plain, then formatted content after it
                target.Range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                target.Range.Text &= sep
                target.Range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                InsertMarkdownToComment(target.Range, replyText)
            Else
                target.Range.Text = (If(target.Range.Text, System.String.Empty)) & sep & replyText
            End If
            Return True
        Catch ex As System.Exception
            ShowCustomMessageBox($"ReplyToWordComment failed to add the reply: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Escapes XML-sensitive characters in the supplied string.
    ''' </summary>
    ''' <param name="s">Input value that may contain reserved XML characters.</param>
    ''' <returns>Escaped string (empty when input is <c>Nothing</c>).</returns>
    Private Shared Function xmlEscapeSafe(ByVal s As System.String) As System.String
        If s Is Nothing Then Return System.String.Empty
        Dim r As System.String = s.Replace("&", "&amp;").
                                  Replace("<", "&lt;").
                                  Replace(">", "&gt;").
                                  Replace("""", "&quot;").
                                  Replace("'", "&apos;")
        Return r
    End Function

    ''' <summary>
    ''' Computes a SHA-256 hash for the provided text and returns the hexadecimal string representation.
    ''' </summary>
    ''' <param name="input">String used to derive a pseudo-stable identifier.</param>
    ''' <returns>Lowercase hexadecimal digest.</returns>
    Private Shared Function ComputeSha256Hex(ByVal input As System.String) As System.String
        If input Is Nothing Then input = System.String.Empty
        Dim bytes As System.Byte() = System.Text.Encoding.UTF8.GetBytes(input)
        Using sha As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
            Dim hash As System.Byte() = sha.ComputeHash(bytes)
            Dim sbHex As New System.Text.StringBuilder(hash.Length * 2)
            For Each b As System.Byte In hash
                sbHex.Append(b.ToString("x2", System.Globalization.CultureInfo.InvariantCulture))
            Next
            Return sbHex.ToString()
        End Using
    End Function

    ''' <summary>
    ''' Processes either the selected text inside the active comment bubble or (when there is no meaningful selection) the entire comment text,
    ''' by sending it to the LLM and writing the formatted result back into the comment story.
    ''' </summary>
    ''' <param name="sysCommand">System prompt used for the LLM call.</param>
    ''' <param name="checkMaxToken">When true, performs a token estimate and shows a warning if output limits may be exceeded.</param>
    ''' <param name="bubbleSelection">Range that represents the current selection inside a comment story; may be <c>Nothing</c>.</param>
    ''' <param name="UseSecondAPI">When true, routes the LLM call to the configured secondary API.</param>
    ''' <param name="SelectionMandatory">When true, shows an error if the resolved input text is empty.</param>
    ''' <param name="FileObject">Optional file object parameter forwarded to the LLM call.</param>
    ''' <returns>Always returns an empty string; errors are reported via message boxes.</returns>
    Private Async Function ProcessSelectedTextInActiveCommentBubble(
    ByVal sysCommand As String,
    ByVal checkMaxToken As Boolean,
    ByVal bubbleSelection As Word.Range,
    ByVal UseSecondAPI As Boolean,
    ByVal SelectionMandatory As Boolean,
    Optional ByVal FileObject As String = ""
) As Task(Of String)

        Try
            Dim app As Word.Application = Globals.ThisAddIn.Application
            If app Is Nothing Then Return ""

            If bubbleSelection Is Nothing Then
                ' Fallback: try to grab current selection, but do not depend on it
                Try
                    bubbleSelection = app.Selection.Range.Duplicate
                Catch
                    bubbleSelection = Nothing
                End Try
            End If

            If bubbleSelection Is Nothing Then
                ShowCustomMessageBox("No comment range available.")
                Return ""
            End If

            Dim doc As Word.Document = bubbleSelection.Document
            If doc Is Nothing Then Return ""

            ' 1) Resolve owning comment
            Dim activeComment As Word.Comment = Nothing
            For Each c As Word.Comment In doc.Comments
                If bubbleSelection.Start >= c.Range.Start AndAlso bubbleSelection.End <= c.Range.End Then
                    activeComment = c
                    Exit For
                End If
            Next

            If activeComment Is Nothing Then
                ShowCustomMessageBox("Cursor is in comment story, but no owning comment could be resolved.")
                Return ""
            End If

            ' 2) Decide whether the user actually has a meaningful selection
            Dim hasRealSelection As Boolean = False
            Try
                hasRealSelection = (bubbleSelection.Start < bubbleSelection.End AndAlso
                                Not String.IsNullOrWhiteSpace(SafeRangeText(bubbleSelection)))
            Catch
                hasRealSelection = False
            End Try

            ' 3) Pick inputRange: selection if meaningful, else the whole comment
            Dim inputRange As Word.Range = If(hasRealSelection, bubbleSelection.Duplicate, activeComment.Range.Duplicate)

            ' Best-effort: exclude the trailing end-of-comment marker / paragraph mark
            Try
                If inputRange.End > inputRange.Start Then inputRange.End -= 1
            Catch
            End Try

            Dim inputText As String = SafeRangeText(inputRange)
            If String.IsNullOrWhiteSpace(inputText) Then
                ' Only complain if the comment is empty; otherwise SelectionMandatory does not apply here
                If SelectionMandatory Then
                    ShowCustomMessageBox("The comment contains no text to process.")
                End If
                Return ""
            End If

            ' 4) Optional token warning
            If checkMaxToken Then
                Dim maxTok As Integer = If(UseSecondAPI, INI_MaxOutputToken_2, INI_MaxOutputToken)
                If maxTok > 0 Then
                    Dim est As Integer = EstimateTokenCount(inputText)
                    If est > maxTok Then
                        ShowCustomMessageBox($"Your comment text may exceed the model output limits ({maxTok} tokens).")
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
                UseSecondAPI:=UseSecondAPI,
                Hidesplash:=False,
                AddUserPrompt:=OtherPrompt,
                FileObject:=FileObject
            )

            llmResult = llmResult.Replace("<TEXTTOPROCESS>", "").Replace("</TEXTTOPROCESS>", "")

            If Not String.IsNullOrEmpty(llmResult) Then
                llmResult = Await PostCorrection(llmResult, UseSecondAPI)
            End If

            If String.IsNullOrWhiteSpace(llmResult) Then
                ShowCustomMessageBox("The LLM did not return any content to process.")
                Return ""
            End If

            ' 6) Write back: selection or whole comment
            Dim targetRange As Word.Range = If(hasRealSelection, bubbleSelection.Duplicate, activeComment.Range.Duplicate)

            Try
                If targetRange.End > targetRange.Start Then targetRange.End -= 1
            Catch
            End Try

            ' Capture "before" snapshot (for post-condition)
            Dim beforeText As String = ""
            Try
                beforeText = activeComment.Range.Text
            Catch
            End Try

            ' Do the write
            Try
                targetRange.Text = ""
            Catch
            End Try

            InsertMarkdownToComment(targetRange, llmResult)

            AbortCommentEditingBestEffort(app)

            ' Force UI refresh (this is what makes the change show up if the user was editing)
            ForceCommentUiRefresh(app, activeComment)

            ' Post-condition: if comment still looks unchanged, instruct user
            Dim afterText As String = ""
            Try
                afterText = activeComment.Range.Text
            Catch
            End Try

            If String.Equals(beforeText, afterText, StringComparison.Ordinal) OrElse
   String.IsNullOrWhiteSpace(afterText) Then
                ShowCustomMessageBox("Word is currently editing this comment. Please submit/close the comment editor (Esc/click outside), then retry.")
                Return ""
            End If

            Return ""

        Catch ex As Exception
            Debug.WriteLine($"ProcessSelectedTextInActiveCommentBubble error: {ex.Message}{vbCrLf}{ex.StackTrace}")
            ShowCustomMessageBox($"Error processing comment bubble: {ex.Message}")
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Attempts to end the active comment editing UI state by sending key strokes.
    ''' </summary>
    ''' <param name="app">The active Word application instance.</param>
    ''' <remarks>
    ''' This is best-effort UI automation and depends on focus being within the comment editor UI.
    ''' </remarks>
    Private Shared Sub AbortCommentEditingBestEffort(ByVal app As Word.Application)
        If app Is Nothing Then Exit Sub

        Try
            ' Best-effort: submit/close the current comment editor.
            ' Observed sequence: TAB, TAB, ENTER.
            System.Windows.Forms.SendKeys.SendWait("{TAB}")
            System.Windows.Forms.SendKeys.SendWait("{TAB}")
            System.Windows.Forms.SendKeys.SendWait("{ENTER}")
            System.Windows.Forms.Application.DoEvents()
        Catch
            ' ignore (UI automation is inherently brittle)
        End Try
    End Sub

    ''' <summary>
    ''' Forces a best-effort UI refresh for a comment by selecting the comment anchor in the main text story.
    ''' </summary>
    ''' <param name="app">The active Word application instance.</param>
    ''' <param name="c">The comment whose anchor should be selected.</param>
    Private Shared Sub ForceCommentUiRefresh(ByVal app As Word.Application, ByVal c As Word.Comment)
        If app Is Nothing OrElse c Is Nothing Then Exit Sub

        Try
            ' Jump to the comment anchor in the main text story
            c.Scope.Select()
            app.Selection.Collapse(Word.WdCollapseDirection.wdCollapseStart)

            ' Let Word pump UI messages
            System.Windows.Forms.Application.DoEvents()

            ' Optional: jump back to comment range (not strictly needed)
            ' c.Range.Select()
        Catch
            ' Best-effort only
        End Try
    End Sub


End Class