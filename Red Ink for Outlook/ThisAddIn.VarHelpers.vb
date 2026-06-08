' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland.
' All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.VarHelpers.vb
' Purpose: Variable and interaction helpers for the Outlook add-in.
'          Current scope: user-driven selection capture and comparison of two
'          text ranges inside the Outlook compose editor (WordEditor).
'
' Architecture:
' - CompareSelectedTextRangesOutlook:
'     - Obtains the active Outlook `Inspector` (compose window) and its Word-based
'       editor via `Inspector.WordEditor`.
'     - Captures the first selection (if present) or prompts the user to select it
'       using SharedLibrary non-modal dialogs (so Outlook remains interactive).
'     - Prompts the user to select a second text range and captures it.
'     - Performs a comparison using existing processing helpers:
'         - Primary: Word compare-docs pipeline via `CompareAndInsertTextCompareDocs`
'           (Word.CompareDocuments-based markup), when enabled by configuration.
'         - Fallback: DiffPlex inline diff via `CompareAndInsertText` (typically shown
'           in a viewer window rather than inserted).
'     - Provides user feedback for missing selections and no-diff cases.
'     - Uses best-effort COM cleanup for local references to `Inspector` / `Document`.
'
' Dependencies:
' - Microsoft Office Interop:
'     - Outlook: `Microsoft.Office.Interop.Outlook.Inspector`
'     - Word: `Microsoft.Office.Interop.Word.Document`, `Selection`
' - SharedLibrary:
'     - `SharedLibrary.SharedLibrary.SharedMethods` for UI prompts and messages.
' - Processing pipeline (same add-in):
'     - `CompareAndInsertTextCompareDocs`, `CompareAndInsertText` (see `ThisAddIn.Processing.vb`).
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Imports DiffPlex
Imports DiffPlex.DiffBuilder
Imports DiffPlex.DiffBuilder.Model
Imports Microsoft.Office.Interop.Outlook
Imports Microsoft.Office.Interop.Word
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn


    ''' <summary>
    ''' Compare two text selections in an Outlook compose inspector.
    ''' Flow:
    '''  - If user already selected text, use as first selection; otherwise prompt to select it.
    '''  - Prompt to select second text.
    '''  - Compare selections using DiffPlex and display in HTML viewer with output options.
    ''' Result:
    '''  - Shows the formatted result in a window with options to copy to clipboard or insert.
    ''' </summary>
    Public Sub CompareSelectedTextRangesOutlook()
            Dim inspector As Inspector = Nothing
            Dim wordDoc As Document = Nothing
            Dim wordApp As Microsoft.Office.Interop.Word.Application = Nothing

            Try
                inspector = TryCast(Globals.ThisAddIn.Application.ActiveInspector(), Inspector)
                If inspector Is Nothing Then
                    SLib.ShowCustomMessageBox("No active compose window found. Please open an email in compose mode and try again.", AN)
                    Return
                End If

                wordDoc = TryCast(inspector.WordEditor, Document)
                If wordDoc Is Nothing Then
                    SLib.ShowCustomMessageBox("Unable to access the Word editor for this item. Please ensure you are in an HTML/RTF compose window.", AN)
                    Return
                End If

                wordApp = wordDoc.Application

                ' --- Step 1: capture first selection (if already selected) ---
                Dim firstText As String = Nothing
                Try
                    Dim sel1 As Microsoft.Office.Interop.Word.Selection = wordApp.Selection
                    If sel1 IsNot Nothing AndAlso sel1.Range IsNot Nothing AndAlso sel1.Start <> sel1.End Then
                        firstText = sel1.Range.Text
                    End If
                Catch
                End Try

                If String.IsNullOrWhiteSpace(firstText) Then
                    Dim step1 As Integer = SLib.ShowCustomYesNoBox(
                    "Please select the FIRST text range to compare in the compose window, then click 'Selection Ready'.",
                    "Selection Ready",
                    "Cancel",
                    $"{AN} Compare Selected - Step 1",
                    nonModal:=True)

                    If step1 <> 1 Then Return

                    Try
                        Dim sel1 As Microsoft.Office.Interop.Word.Selection = wordApp.Selection
                        If sel1 IsNot Nothing AndAlso sel1.Range IsNot Nothing AndAlso sel1.Start <> sel1.End Then
                            firstText = sel1.Range.Text
                        End If
                    Catch
                    End Try

                    If String.IsNullOrWhiteSpace(firstText) Then
                        SLib.ShowCustomMessageBox("No text was selected for the first range. Operation cancelled.", AN)
                        Return
                    End If
                End If

                ' --- Step 2: prompt and capture second selection ---
                Dim step2 As Integer = SLib.ShowCustomYesNoBox(
                $"First selection captured ({firstText.Length} characters).{vbCrLf}{vbCrLf}Now please select the SECOND text range to compare, then click 'Selection Ready'.",
                "Selection Ready",
                "Cancel",
                $"{AN} Compare Selected - Step 2",
                nonModal:=True)

                If step2 <> 1 Then Return

                Dim secondText As String = Nothing
                Try
                    Dim sel2 As Microsoft.Office.Interop.Word.Selection = wordApp.Selection
                    If sel2 IsNot Nothing AndAlso sel2.Range IsNot Nothing AndAlso sel2.Start <> sel2.End Then
                        secondText = sel2.Range.Text
                    End If
                Catch
                End Try

                If String.IsNullOrWhiteSpace(secondText) Then
                    SLib.ShowCustomMessageBox("No text was selected for the second range. Operation cancelled.", AN)
                    Return
                End If

                ' --- identical check ---
                If String.Equals(firstText, secondText, StringComparison.Ordinal) Then
                    SLib.ShowCustomMessageBox("The two selected text ranges are identical. No differences to show.", AN)
                    Return
                End If

                ' --- compare using DiffPlex and show with enhanced options ---
                CompareAndShowWithOptionsOutlook(firstText, secondText, wordDoc)

            Catch ex As System.Exception
                SLib.ShowCustomMessageBox($"Failed to compare selected text: {ex.Message}", AN)
            Finally
                ' Best-effort COM cleanup (avoid over-releasing objects owned by Outlook/Word)
                If wordDoc IsNot Nothing Then
                    Try : System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDoc) : Catch : End Try
                    wordDoc = Nothing
                End If
                If inspector IsNot Nothing Then
                    Try : System.Runtime.InteropServices.Marshal.ReleaseComObject(inspector) : Catch : End Try
                    inspector = Nothing
                End If
            End Try
        End Sub

    ''' <summary>
    ''' Compares two text selections using DiffPlex and displays in an HTML viewer
    ''' with option to copy with formatting.
    ''' </summary>
    ''' <param name="firstText">Original text for comparison.</param>
    ''' <param name="secondText">Revised text for comparison.</param>
    ''' <param name="wordDoc">Word document from Outlook's WordEditor for insertion.</param>
    Private Sub CompareAndShowWithOptionsOutlook(firstText As String, secondText As String, wordDoc As Document)
        Try
            ' Generate diff using DiffPlex
            Dim diffBuilder As New InlineDiffBuilder(New Differ())

            ' Pre-process the texts to handle line breaks
            Dim text1 As String = firstText.Replace(vbCrLf, " {LINEBREAK} ").Replace(vbCr, " {LINEBREAK} ").Replace(vbLf, " {LINEBREAK} ")
            Dim text2 As String = secondText.Replace(vbCrLf, " {LINEBREAK} ").Replace(vbCr, " {LINEBREAK} ").Replace(vbLf, " {LINEBREAK} ")

            ' Normalize extra spaces
            text1 = text1.Replace("  ", " ").Trim()
            text2 = text2.Replace("  ", " ").Trim()

            ' Split into words for word-level diff
            Dim words1 As String = String.Join(Environment.NewLine, text1.Split(" "c))
            Dim words2 As String = String.Join(Environment.NewLine, text2.Split(" "c))

            ' Generate word-based diff
            Dim diffResult As DiffPaneModel = diffBuilder.BuildDiffModel(words1, words2)

            ' Build HTML output with inline styles
            Dim htmlBuilder As New StringBuilder()
            htmlBuilder.AppendLine("<!DOCTYPE html>")
            htmlBuilder.AppendLine("<html><head>")
            htmlBuilder.AppendLine("<meta charset=""utf-8"">")
            htmlBuilder.AppendLine("<style>")
            htmlBuilder.AppendLine("body { font-family: 'Segoe UI', Calibri, Arial, sans-serif; font-size: 11pt; line-height: 1.5; padding: 15px; }")
            htmlBuilder.AppendLine(".ins { color: #0000FF; text-decoration: underline; }")
            htmlBuilder.AppendLine(".del { color: #FF0000; text-decoration: line-through; }")
            htmlBuilder.AppendLine("</style>")
            htmlBuilder.AppendLine("</head><body>")

            ' Build the diff output
            Dim plainTextBuilder As New StringBuilder()

            For Each line In diffResult.Lines
                Dim word As String = line.Text.Trim()
                If String.IsNullOrEmpty(word) Then Continue For

                ' Handle line break markers
                If word = "{LINEBREAK}" Then
                    htmlBuilder.Append("<br>")
                    plainTextBuilder.AppendLine()
                    Continue For
                End If

                Select Case line.Type
                    Case ChangeType.Inserted
                        htmlBuilder.Append($"<span class=""ins"">{System.Net.WebUtility.HtmlEncode(word)}</span> ")
                        plainTextBuilder.Append($"[+{word}] ")

                    Case ChangeType.Deleted
                        htmlBuilder.Append($"<span class=""del"">{System.Net.WebUtility.HtmlEncode(word)}</span> ")
                        plainTextBuilder.Append($"[-{word}] ")

                    Case ChangeType.Unchanged
                        htmlBuilder.Append($"{System.Net.WebUtility.HtmlEncode(word)} ")
                        plainTextBuilder.Append($"{word} ")
                End Select
            Next

            htmlBuilder.AppendLine("</body></html>")

            Dim htmlContent As String = htmlBuilder.ToString()
            Dim plainText As String = plainTextBuilder.ToString().Trim()

            ' Capture for closure
            Dim capturedHtml As String = htmlContent
            Dim capturedPlain As String = plainText

            ' Single button to copy with formatting
            Dim additionalButtons As New List(Of System.Tuple(Of String, System.Action, Boolean))()

            additionalButtons.Add(New System.Tuple(Of String, System.Action, Boolean)(
        "Copy",
        Sub()
            Try
                CopyHtmlToClipboard(capturedHtml, capturedPlain)
                SLib.ShowCustomMessageBox("Comparison copied to clipboard with formatting.", AN)
            Catch ex As System.Exception
                SLib.ShowCustomMessageBox($"Failed to copy: {ex.Message}", AN)
            End Try
        End Sub,
        False))

            ' Show result in HTML viewer
            SLib.ShowHTMLCustomMessageBox(htmlContent, $"{AN} Text Comparison", additionalButtons:=additionalButtons.ToArray())

        Catch ex As System.Exception
            SLib.ShowCustomMessageBox($"Failed to compare text: {ex.Message}", AN)
        End Try
    End Sub


    ''' <summary>
    ''' Copies HTML content to clipboard in a format that can be pasted with formatting.
    ''' </summary>
    Private Sub CopyHtmlToClipboard(html As String, plainText As String)
        Dim dataObject As New DataObject()

        ' Add plain text fallback
        dataObject.SetData(DataFormats.Text, plainText)
        dataObject.SetData(DataFormats.UnicodeText, plainText)

        ' Add HTML in CF_HTML format
        Dim cfHtml As String = BuildCfHtml(html)
        dataObject.SetData(DataFormats.Html, cfHtml)

        ' Add RTF format - many apps prefer RTF and paste with formatting by default
        Dim rtfContent As String = BuildRtfFromHtml(html)
        dataObject.SetData(DataFormats.Rtf, rtfContent)

        ' Use SetDataObject with copy=True to keep data available after paste
        Clipboard.SetDataObject(dataObject, True, 5, 100)
    End Sub

    ''' <summary>
    ''' Converts the diff HTML to RTF format for better paste compatibility.
    ''' </summary>
    Private Function BuildRtfFromHtml(html As String) As String
        Dim sb As New StringBuilder()
        sb.Append("{\rtf1\ansi\deff0")
        sb.Append("{\fonttbl{\f0 Segoe UI;}}")
        sb.Append("{\colortbl;\red255\green0\blue0;\red0\green0\blue255;}")  ' cf1=red, cf2=blue
        sb.Append("\f0\fs22 ")  ' 11pt font

        ' Parse HTML spans and convert to RTF
        Dim pos As Integer = 0
        Dim bodyStart As Integer = html.IndexOf("<body>", StringComparison.OrdinalIgnoreCase)
        If bodyStart >= 0 Then pos = bodyStart + 6
        Dim bodyEnd As Integer = html.IndexOf("</body>", StringComparison.OrdinalIgnoreCase)
        If bodyEnd < 0 Then bodyEnd = html.Length

        Dim content As String = html.Substring(pos, bodyEnd - pos)

        ' Process the content
        Dim i As Integer = 0
        While i < content.Length
            If content.Substring(i).StartsWith("<span class=""del"">", StringComparison.OrdinalIgnoreCase) Then
                ' Deleted text - red strikethrough
                Dim closeTag As Integer = content.IndexOf("</span>", i, StringComparison.OrdinalIgnoreCase)
                If closeTag > i Then
                    Dim startText As Integer = content.IndexOf(">"c, i) + 1
                    Dim text As String = System.Net.WebUtility.HtmlDecode(content.Substring(startText, closeTag - startText))
                    sb.Append($"{{\cf1\strike {EscapeRtf(text)}}}\strike0 ")
                    i = closeTag + 7
                    Continue While
                End If
            ElseIf content.Substring(i).StartsWith("<span class=""ins"">", StringComparison.OrdinalIgnoreCase) Then
                ' Inserted text - blue underline
                Dim closeTag As Integer = content.IndexOf("</span>", i, StringComparison.OrdinalIgnoreCase)
                If closeTag > i Then
                    Dim startText As Integer = content.IndexOf(">"c, i) + 1
                    Dim text As String = System.Net.WebUtility.HtmlDecode(content.Substring(startText, closeTag - startText))
                    sb.Append($"{{\cf2\ul {EscapeRtf(text)}}}\ul0 ")
                    i = closeTag + 7
                    Continue While
                End If
            ElseIf content.Substring(i).StartsWith("<br>", StringComparison.OrdinalIgnoreCase) Then
                sb.Append("\par ")
                i += 4
                Continue While
            ElseIf content(i) = "<"c Then
                ' Skip other HTML tags
                Dim closeAngle As Integer = content.IndexOf(">"c, i)
                If closeAngle > i Then
                    i = closeAngle + 1
                    Continue While
                End If
            End If

            ' Regular character
            If i < content.Length AndAlso content(i) <> "<"c Then
                sb.Append(EscapeRtf(content(i).ToString()))
            End If
            i += 1
        End While

        sb.Append("}")
        Return sb.ToString()
    End Function

    ''' <summary>
    ''' Builds the CF_HTML clipboard format header for HTML content.
    ''' </summary>
    Private Function BuildCfHtml(html As String) As String
            Dim sb As New StringBuilder()

            ' CF_HTML format requires specific headers
            Const startHtmlPlaceholder As String = "0000000000"
            Const endHtmlPlaceholder As String = "0000000000"
            Const startFragmentPlaceholder As String = "0000000000"
            Const endFragmentPlaceholder As String = "0000000000"

            sb.AppendLine("Version:0.9")
            sb.AppendLine($"StartHTML:{startHtmlPlaceholder}")
            sb.AppendLine($"EndHTML:{endHtmlPlaceholder}")
            sb.AppendLine($"StartFragment:{startFragmentPlaceholder}")
            sb.AppendLine($"EndFragment:{endFragmentPlaceholder}")

            Dim headerLength As Integer = sb.Length
            Dim startHtml As Integer = headerLength

            sb.Append("<!--StartFragment-->")
            Dim startFragment As Integer = sb.Length

            sb.Append(html)
            Dim endFragment As Integer = sb.Length

            sb.Append("<!--EndFragment-->")
            Dim endHtml As Integer = sb.Length

            ' Replace placeholders with actual values
            Dim result As String = sb.ToString()
            result = result.Replace($"StartHTML:{startHtmlPlaceholder}", $"StartHTML:{startHtml:D10}")
            result = result.Replace($"EndHTML:{endHtmlPlaceholder}", $"EndHTML:{endHtml:D10}")
            result = result.Replace($"StartFragment:{startFragmentPlaceholder}", $"StartFragment:{startFragment:D10}")
            result = result.Replace($"EndFragment:{endFragmentPlaceholder}", $"EndFragment:{endFragment:D10}")

            Return result
        End Function

        ''' <summary>
        ''' Inserts comparison result into the current email at cursor position.
        ''' Handles HTML, RTF, and plain text email formats.
        ''' </summary>
        Private Sub InsertComparisonIntoEmail(wordDoc As Document, html As String, plainText As String)
            If wordDoc Is Nothing Then
                Throw New InvalidOperationException("Word document is not available.")
            End If

            Dim sel As Microsoft.Office.Interop.Word.Selection = wordDoc.Application.Selection

            ' Collapse selection to end (insert after current position)
            sel.Collapse(WdCollapseDirection.wdCollapseEnd)

            ' Try to insert HTML if the email supports it
            Try
                ' Insert a paragraph break first
                sel.TypeParagraph()
                sel.TypeParagraph()

                ' For HTML/RTF emails, we can paste HTML
                CopyHtmlToClipboard(html, plainText)
                sel.PasteAndFormat(WdRecoveryType.wdFormatOriginalFormatting)

                sel.TypeParagraph()
            Catch
                ' Fallback to plain text if HTML paste fails
                sel.TypeText(vbCrLf & vbCrLf & plainText & vbCrLf)
            End Try
        End Sub

        ''' <summary>
        ''' Escapes special RTF characters.
        ''' </summary>
        Private Function EscapeRtf(text As String) As String
            If String.IsNullOrEmpty(text) Then Return String.Empty
            Return text.Replace("\", "\\").Replace("{", "\{").Replace("}", "\}")
        End Function

        ''' <summary>
        ''' Builds a complete RTF document with color table for diff formatting.
        ''' </summary>
        Private Function BuildRtfDocument(content As String) As String
            Dim sb As New StringBuilder()
            sb.Append("{\rtf1\ansi\deff0")
            sb.Append("{\colortbl;\red255\green0\blue0;\red0\green0\blue255;}")  ' cf1=red, cf2=blue
            sb.Append("\pard ")
            sb.Append(content)
            sb.Append("}")
            Return sb.ToString()
        End Function




        ''' <summary>
        ''' Compare two text selections in an Outlook compose inspector.
        ''' Flow:
        '''  - If user already selected text, use as first selection; otherwise prompt to select it.
        '''  - Prompt to select second text.
        '''  - Compare selections using Word Compare (CompareAndInsertTextCompareDocs) or DiffPlex fallback (CompareAndInsertText).
        ''' Result:
        '''  - Either inserts formatted comparison into the compose window (CompareDocs path)
        '''  - or shows the formatted result in a window (DiffPlex path).
        ''' </summary>
        Public Sub OldCompareSelectedTextRangesOutlook()
        Dim inspector As Inspector = Nothing
        Dim wordDoc As Document = Nothing
        Dim wordApp As Microsoft.Office.Interop.Word.Application = Nothing

        Try
            inspector = TryCast(Globals.ThisAddIn.Application.ActiveInspector(), Inspector)
            If inspector Is Nothing Then
                SLib.ShowCustomMessageBox("No active compose window found. Please open an email in compose mode and try again.", AN)
                Return
            End If

            wordDoc = TryCast(inspector.WordEditor, Document)
            If wordDoc Is Nothing Then
                SLib.ShowCustomMessageBox("Unable to access the Word editor for this item. Please ensure you are in an HTML/RTF compose window.", AN)
                Return
            End If

            wordApp = wordDoc.Application

            ' --- Step 1: capture first selection (if already selected) ---
            Dim firstText As String = Nothing
            Try
                Dim sel1 As Microsoft.Office.Interop.Word.Selection = wordApp.Selection
                If sel1 IsNot Nothing AndAlso sel1.Range IsNot Nothing AndAlso sel1.Start <> sel1.End Then
                    firstText = sel1.Range.Text
                End If
            Catch
            End Try

            If String.IsNullOrWhiteSpace(firstText) Then
                Dim step1 As Integer = SLib.ShowCustomYesNoBox(
                    "Please select the FIRST text range to compare in the compose window, then click 'Selection Ready'.",
                    "Selection Ready",
                    "Cancel",
                    $"{AN} Compare Selected - Step 1",
                    nonModal:=True)

                If step1 <> 1 Then Return

                Try
                    Dim sel1 As Microsoft.Office.Interop.Word.Selection = wordApp.Selection
                    If sel1 IsNot Nothing AndAlso sel1.Range IsNot Nothing AndAlso sel1.Start <> sel1.End Then
                        firstText = sel1.Range.Text
                    End If
                Catch
                End Try

                If String.IsNullOrWhiteSpace(firstText) Then
                    SLib.ShowCustomMessageBox("No text was selected for the first range. Operation cancelled.", AN)
                    Return
                End If
            End If

            ' --- Step 2: prompt and capture second selection ---
            Dim step2 As Integer = SLib.ShowCustomYesNoBox(
                $"First selection captured ({firstText.Length} characters).{vbCrLf}{vbCrLf}Now please select the SECOND text range to compare, then click 'Selection Ready'.",
                "Selection Ready",
                "Cancel",
                $"{AN} Compare Selected - Step 2",
                nonModal:=True)

            If step2 <> 1 Then Return

            Dim secondText As String = Nothing
            Try
                Dim sel2 As Microsoft.Office.Interop.Word.Selection = wordApp.Selection
                If sel2 IsNot Nothing AndAlso sel2.Range IsNot Nothing AndAlso sel2.Start <> sel2.End Then
                    secondText = sel2.Range.Text
                End If
            Catch
            End Try

            If String.IsNullOrWhiteSpace(secondText) Then
                SLib.ShowCustomMessageBox("No text was selected for the second range. Operation cancelled.", AN)
                Return
            End If

            ' --- identical check ---
            If String.Equals(firstText, secondText, StringComparison.Ordinal) Then
                SLib.ShowCustomMessageBox("The two selected text ranges are identical. No differences to show.", AN)
                Return
            End If

            ' --- compare (Word CompareDocs or fallback diff) ---
            ' Use your existing configuration switch (same pattern as Word CompareSelectionHalves).
            ' INI_MarkupMethodHelper = 1 => CompareDoc path (Word CompareDocuments)
            'If INI_MarkupMethodHelper = 1 Then
            ' Inserts formatted comparison at current cursor in compose window.
            'CompareAndInsertTextCompareDocs(firstText, secondText)
            'Else
            ' Shows the diff in a window (does not insert).
            CompareAndInsertText(
                    firstText,
                    secondText,
                    ShowInWindow:=True,
                    TextforWindow:="Comparison result (not inserted):",
                    DoNotWait:=True)
            'End If

        Catch ex As System.Exception
            SLib.ShowCustomMessageBox($"Failed to compare selected text: {ex.Message}", AN)
        Finally
            ' Best-effort COM cleanup (avoid over-releasing objects owned by Outlook/Word)
            If wordDoc IsNot Nothing Then
                Try : System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDoc) : Catch : End Try
                wordDoc = Nothing
            End If
            If inspector IsNot Nothing Then
                Try : System.Runtime.InteropServices.Marshal.ReleaseComObject(inspector) : Catch : End Try
                inspector = Nothing
            End If
        End Try
    End Sub


    Private _quickTranslateWidget As Global.SharedLibrary.SharedLibrary.QuickTranslateWidget = Nothing

    Public Sub ShowQuickTranslate()
        If _quickTranslateWidget Is Nothing OrElse _quickTranslateWidget.IsDisposed Then
            _quickTranslateWidget = New Global.SharedLibrary.SharedLibrary.QuickTranslateWidget(
                Async Function(text, lang, sourcelang, token)
                    TranslateLanguage = lang
                    SourceLanguage = sourcelang
                    Dim SysPrompt As String = SP_Translate_Multi
                    If Not String.IsNullOrWhiteSpace(SourceLanguage) Then SysPrompt = SP_Translate_Multi_Source
                    Return Await LLM(InterpolateAtRuntime(SysPrompt),
                                    "<TEXTTOPROCESS>" & text & "</TEXTTOPROCESS>",
                                    "", "", 0,
                                    UseSecondAPI:=False,
                                    HideSplash:=True,
                                    cancellationToken:=token,
                                    EnsureUI:=False)
                End Function,
                INI_Language1,
                Sub()
                    Global.SharedLibrary.SharedLibrary.SharedMethods.EditUserDictionaryFile(_context)
                End Sub)
        End If
        _quickTranslateWidget.ShowWidget()
    End Sub


End Class
