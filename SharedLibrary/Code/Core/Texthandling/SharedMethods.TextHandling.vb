' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: SharedMethods.TextHandling.vb
' Purpose: Provides helper methods for converting and inserting text into a
'          Microsoft Word document range/selection, including Markdown-to-HTML
'          conversion, HTML cleanup/simplification, and markup/format stripping.
'
' Architecture:
'  - Markdown Insertion: `InsertTextWithMarkdown` converts Markdown to HTML (Markdig),
'    normalizes line breaks, and delegates insertion to `InsertTextWithFormat`.
'  - HTML Insertion: `InsertTextWithFormat` loads HTML (HtmlAgilityPack), normalizes
'    `<br>` usage inside paragraphs/list items, applies Word-derived inline styles,
'    constructs a CF_HTML clipboard packet (UTF-8 byte offsets), and pastes into Word.
'  - Word Highlight Tags: `FixMarkTagsForWord` translates `<mark>` elements into
'    `<span>` elements with `mso-highlight:*` so Word can interpret highlights.
'  - HTML/Text Utilities: Helpers remove HTML, create normalized HTML with a
'    `<head>`/`<meta charset>`, and simplify HTML by whitelisting tags/attributes.
'  - Markdown Stripping: `RemoveMarkdownFormatting` removes a subset of Markdown
'    markers while preserving bracketed/brace-delimited regions verbatim.
'
' External Dependencies:
'  - Markdig: Markdown parsing/conversion to HTML.
'  - HtmlAgilityPack: HTML parsing, manipulation, and entity decoding.
'  - Microsoft.Office.Interop.Word: Word Range/Selection manipulation and paste APIs.
' =============================================================================


Option Strict On
Option Explicit On

Imports System.Text.RegularExpressions
Imports HtmlAgilityPack
Imports Markdig
Imports Microsoft.Office.Interop.Word
Imports DiffPlex
Imports DiffPlex.DiffBuilder
Imports DiffPlex.DiffBuilder.Model

Namespace SharedLibrary
    Partial Public Class SharedMethods


        ''' <summary>
        ''' Builds inline diff markup using [INS_START]/[INS_END] and [DEL_START]/[DEL_END] tags.
        ''' </summary>
        ''' <param name="originalText">Original text.</param>
        ''' <param name="revisedText">Revised text.</param>
        ''' <param name="trimTrailingLineBreaksOnRevised">If True, trims trailing CR/LF from the revised text before diffing.</param>
        ''' <returns>Diff-marked text suitable for <see cref="ConvertMarkupToRTF"/>.</returns>
        Public Shared Function BuildInlineDiffMarkup(
            originalText As String,
            revisedText As String,
            Optional trimTrailingLineBreaksOnRevised As Boolean = True
        ) As String

            Dim diffBuilder As New InlineDiffBuilder(New Differ())
            Dim sText As String = String.Empty

            Dim text1 As String = If(originalText, String.Empty)
            Dim text2 As String = If(revisedText, String.Empty)

            If trimTrailingLineBreaksOnRevised Then
                text2 = text2.TrimEnd(ControlChars.Cr, ControlChars.Lf).TrimEnd(ControlChars.Cr, ControlChars.Lf)
            End If

            text1 = text1.Replace(vbCrLf, " {vbCrLf} ").Replace(vbCr, " {vbCr} ").Replace(vbLf, " {vbLf} ")
            text2 = text2.Replace(vbCrLf, " {vbCrLf} ").Replace(vbCr, " {vbCr} ").Replace(vbLf, " {vbLf} ")

            text1 = text1.Replace("  ", " ").Trim()
            text2 = text2.Replace("  ", " ").Trim()

            Dim mergefields As New List(Of String)

            text1 = Regex.Replace(
                text1,
                "\{\{.*?\}\}",
                Function(m)
                    mergefields.Add(m.Value)
                    Return $"[[MF{mergefields.Count - 1}]]"
                End Function)

            text2 = Regex.Replace(
                text2,
                "\{\{.*?\}\}",
                Function(m)
                    mergefields.Add(m.Value)
                    Return $"[[MF{mergefields.Count - 1}]]"
                End Function)

            Dim words1 As String =
                String.Join(
                    Environment.NewLine,
                    text1.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries))

            Dim words2 As String =
                String.Join(
                    Environment.NewLine,
                    text2.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries))

            Dim diffResult As DiffPaneModel = diffBuilder.BuildDiffModel(words1, words2)

            Dim prevType As ChangeType = ChangeType.Unchanged

            For i As Integer = 0 To diffResult.Lines.Count - 1
                Dim line = diffResult.Lines(i)
                Dim nextType As ChangeType =
                    If(i < diffResult.Lines.Count - 1, diffResult.Lines(i + 1).Type, ChangeType.Unchanged)

                If line.Type = ChangeType.Inserted AndAlso prevType <> ChangeType.Inserted Then
                    sText &= "[INS_START]"
                ElseIf line.Type = ChangeType.Deleted AndAlso prevType <> ChangeType.Deleted Then
                    sText &= "[DEL_START]"
                End If

                sText &= If(line.Text, String.Empty).Trim() & " "

                If line.Type = ChangeType.Inserted AndAlso nextType <> ChangeType.Inserted Then
                    sText &= "[INS_END] "
                ElseIf line.Type = ChangeType.Deleted AndAlso nextType <> ChangeType.Deleted Then
                    sText &= "[DEL_END] "
                End If

                prevType = line.Type
            Next

            For idx As Integer = 0 To mergefields.Count - 1
                sText = sText.Replace($"[[MF{idx}]]", mergefields(idx))
            Next

            sText = sText.Replace("{vbCr}", "{vbCrLf}")
            sText = sText.Replace("{vbLf}", "{vbCrLf}")
            sText = sText.Replace(" {vbCrLf} ", "{vbCrLf}")
            sText = sText.Replace(" {vbCrLf}", "{vbCrLf}")
            sText = sText.Replace("{vbCrLf} ", "{vbCrLf}")

            sText = sText.Replace("[DEL_START]{vbCrLf}[DEL_END] ", "")
            sText = sText.Replace("[DEL_START]{vbCrLf}{vbCrLf}[DEL_END] ", "")
            sText = sText.Replace("{vbCrLf}[DEL_END] ", "{vbCrLf}[DEL_END]")

            sText = sText.Replace("[INS_START]{vbCrLf}[INS_END] ", "{vbCrLf}")
            sText = sText.Replace("[INS_START]{vbCrLf}{vbCrLf}[INS_END] ", "{vbCrLf}{vbCrLf}")
            sText = sText.Replace("{vbCrLf}[INS_END] ", "{vbCrLf}[INS_END]")

            sText = sText.Replace(vbCrLf, "").Replace(vbCr, "").Replace(vbLf, "")
            sText = sText.Replace("{vbCrLf}", vbCrLf)

            sText = sText.Replace("[DEL_END] [INS_START]", "[DEL_END][INS_START]")
            sText = sText.Replace("[INS_START][INS_END] ", "")

            Return sText.TrimEnd()
        End Function

        ''' <summary>
        ''' Converts the provided Markdown text to HTML and inserts it into the given Word selection.
        ''' </summary>
        ''' <param name="selection">A Word <see cref="Microsoft.Office.Interop.Word.Selection"/> (passed as <see cref="Object"/>).</param>
        ''' <param name="gptResult">The Markdown text to convert and insert.</param>
        ''' <param name="TrailingCR">If <c>True</c>, keeps trailing paragraph breaks; otherwise suppresses them.</param>
        Public Shared Sub InsertTextWithMarkdown(selection As Object,
                                                 gptResult As String,
                                                 TrailingCR As Boolean,
                                                 Optional UseHostDefaultFontColor As Boolean = False)

            Dim wordSelection As Microsoft.Office.Interop.Word.Selection = CType(selection, Microsoft.Office.Interop.Word.Selection)
            Dim wordRange As Microsoft.Office.Interop.Word.Range = wordSelection.Range

            Debug.WriteLine("ITWM: " & gptResult)

            gptResult = gptResult.Replace(vbLf & " " & vbLf, vbLf & vbLf)

            Dim pattern As String = "((\r\n|\n|\r){2,})"
            gptResult = Regex.Replace(gptResult, pattern, Function(m As Match)
                                                              ' Check whether the match reaches to the end of the string
                                                              If m.Index + m.Length = gptResult.Length Then
                                                                  ' At the end: return the line breaks as they are
                                                                  Return m.Value
                                                              Else
                                                                  ' Otherwise: insert &nbsp; between the line breaks
                                                                  Dim breaks As String = m.Value
                                                                  Dim regexBreaks As New Regex("(\r\n|\n|\r)")
                                                                  Dim splitBreaks = regexBreaks.Matches(breaks)
                                                                  If splitBreaks.Count <= 1 Then Return breaks
                                                                  Dim result As String = splitBreaks(0).Value
                                                                  For i As Integer = 1 To splitBreaks.Count - 1
                                                                      result &= vbCrLf & "&nbsp;" & vbCrLf & splitBreaks(i).Value
                                                                  Next
                                                                  Return result
                                                              End If
                                                          End Function)


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

            Dim pipeline As MarkdownPipeline = builder.Build()

            Dim htmlresult As String = Markdown.ToHtml(gptResult, pipeline)


            htmlresult = htmlresult _
                .Replace(vbCrLf, "") _
                .Replace(vbCr, "") _
                .Replace(vbLf, "")


            ' Load the HTML into HtmlDocument
            Dim htmlDoc As New HtmlAgilityPack.HtmlDocument()
            Dim fullhtml As String
            htmlDoc.LoadHtml(htmlresult)

            fullhtml = htmlDoc.DocumentNode.OuterHtml

            Debug.WriteLine("ITWM: " & fullhtml)

            InsertTextWithFormat(fullhtml, wordRange, True, Not TrailingCR, UseHostDefaultFontColor)

        End Sub


        ''' <summary>
        ''' Inserts HTML-formatted content into a Word range using CF_HTML clipboard formatting and Word paste APIs.
        ''' </summary>
        ''' <param name="formattedText">The HTML fragment to insert.</param>
        ''' <param name="range">The Word range that defines the paste target and receives the updated inserted range.</param>
        ''' <param name="ReplaceSelection">If <c>True</c>, pastes over the current selection; otherwise appends at the range end.</param>
        ''' <param name="NoTrailingCR">If <c>True</c> and <paramref name="ReplaceSelection"/> is <c>True</c>, deletes the last paragraph mark after insertion.</param>
        Public Shared Sub InsertTextWithFormat(formattedText As String,
                                               ByRef range As Microsoft.Office.Interop.Word.Range,
                                               ReplaceSelection As Boolean,
                                               Optional NoTrailingCR As Boolean = False,
                                               Optional UseHostDefaultFontColor As Boolean = False)
            Try
                If formattedText Is Nothing OrElse formattedText.Trim() = "" Then
                    Return
                End If

                ' --- 0) Clone original range start and collapse to the start ---
                Dim origRange As Microsoft.Office.Interop.Word.Range = range.Duplicate()
                origRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseStart)

                System.Diagnostics.Debug.WriteLine("PreFinalHTML=" & formattedText)

                formattedText = FixMarkTagsForWord(formattedText)

                System.Diagnostics.Debug.WriteLine("PreFinalHTML[after-mark]=" & formattedText)

                ' --- 1) Load HTML and split <br> into separate <p> elements ---
                Dim doc As New HtmlAgilityPack.HtmlDocument()
                doc.LoadHtml(formattedText)

                ' Select all <p> and <li> nodes
                Dim nodes As HtmlAgilityPack.HtmlNodeCollection = doc.DocumentNode.SelectNodes("//p | //li")
                If nodes IsNot Nothing Then
                    For Each node As HtmlAgilityPack.HtmlNode In nodes.ToList()
                        Dim segments As String() = System.Text.RegularExpressions.Regex.Split(node.InnerHtml, "<br\s*/?>", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
                        If segments.Length <= 1 Then Continue For

                        If node.Name.Equals("p", System.StringComparison.OrdinalIgnoreCase) Then
                            Dim parent As HtmlAgilityPack.HtmlNode = node.ParentNode
                            If parent Is Nothing Then Continue For

                            For Each seg As String In segments
                                Dim txt As String = seg.Trim()
                                If System.String.IsNullOrEmpty(txt) Then Continue For
                                Dim newP As HtmlAgilityPack.HtmlNode = doc.CreateElement("p")
                                newP.InnerHtml = txt
                                parent.InsertBefore(newP, node)
                            Next
                            parent.RemoveChild(node)

                        ElseIf node.Name.Equals("li", System.StringComparison.OrdinalIgnoreCase) Then
                            node.RemoveAllChildren()
                            For Each seg As String In segments
                                Dim txt As String = seg.Trim()
                                If System.String.IsNullOrEmpty(txt) Then Continue For
                                Dim newP As HtmlAgilityPack.HtmlNode = doc.CreateElement("p")
                                newP.InnerHtml = txt
                                node.AppendChild(newP)
                            Next
                        End If
                    Next
                End If

                formattedText = doc.DocumentNode.OuterHtml

                ' --- 2) Read font and paragraph properties from the first character of the range ---
                '     A collapsed range at a paragraph boundary can inherit font properties from the
                '     *following* paragraph. Reading from the first concrete character avoids this.
                Dim fontSourceRange As Microsoft.Office.Interop.Word.Range = range.Duplicate()
                If fontSourceRange.Characters.Count > 0 Then
                    fontSourceRange.SetRange(fontSourceRange.Start, fontSourceRange.Start + 1)
                End If

                Dim fontName As String = fontSourceRange.Font.Name
                Dim fontSize As Single = fontSourceRange.Font.Size
                Dim isBold As Boolean = (fontSourceRange.Font.Bold = 1)
                Dim isItalic As Boolean = (fontSourceRange.Font.Italic = 1)
                Dim fontColor As Integer = fontSourceRange.Font.Color
                Dim hexColor As String = String.Empty

                ' Guard against ambiguous values (9999999 = mixed formatting in selection)
                If fontSize <= 0 OrElse fontSize > 1000 Then fontSize = 11.0F
                If fontName Is Nothing OrElse fontName = "" Then fontName = "Calibri"

                If Not UseHostDefaultFontColor Then
                    ' Convert Word BGR color to RGB hex string
                    Dim bgr As Integer = fontColor And &HFFFFFF
                    Dim r As Integer = (bgr And &HFF)
                    Dim g As Integer = ((bgr >> 8) And &HFF)
                    Dim b As Integer = ((bgr >> 16) And &HFF)
                    hexColor = System.String.Format("#{0:X2}{1:X2}{2:X2}", r, g, b)
                End If

                Dim para As Microsoft.Office.Interop.Word.ParagraphFormat = fontSourceRange.ParagraphFormat
                Dim spaceBefore As Single = para.SpaceBefore
                Dim spaceAfter As Single = para.SpaceAfter
                Dim lineRule As Microsoft.Office.Interop.Word.WdLineSpacing = para.LineSpacingRule
                Dim rawLineSpacing As Single = para.LineSpacing

                Dim lineHeightCss As String
                Select Case lineRule
                    Case Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceSingle
                        lineHeightCss = "normal"
                    Case Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpace1pt5
                        lineHeightCss = "1.5"
                    Case Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceDouble
                        lineHeightCss = "2"
                    Case Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceMultiple
                        lineHeightCss = rawLineSpacing.ToString() & "pt"
                    Case Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly,
                 Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceAtLeast
                        lineHeightCss = rawLineSpacing.ToString() & "pt"
                    Case Else
                        lineHeightCss = "normal"
                End Select

                ' --- 3) Build CSS strings ---
                Dim cssBody As String = $"font-family:'{fontName}'; line-height:{lineHeightCss};"
                If Not UseHostDefaultFontColor Then
                    cssBody &= $" color:{hexColor};"
                End If
                Dim cssPara As String = cssBody & $" font-size:{fontSize}pt; margin-top:{spaceBefore}pt; margin-bottom:{spaceAfter}pt;"
                If isBold Then cssPara &= " font-weight:bold;"
                If isItalic Then cssPara &= " font-style:italic;"

                ' --- 4) Apply inline styles ---
                Dim allTextContainers As HtmlAgilityPack.HtmlNodeCollection = doc.DocumentNode.SelectNodes("//p | //li")
                If allTextContainers IsNot Nothing Then
                    For Each n As HtmlAgilityPack.HtmlNode In allTextContainers
                        n.SetAttributeValue("style", cssPara)
                    Next
                End If

                ' Headings (h1–h6): override only family/color/line-height
                Dim headings As HtmlAgilityPack.HtmlNodeCollection = doc.DocumentNode.SelectNodes("//h1 | //h2 | //h3 | //h4 | //h5 | //h6")
                If headings IsNot Nothing Then
                    For Each h As HtmlAgilityPack.HtmlNode In headings
                        Dim current As String = h.GetAttributeValue("style", "")
                        If Not System.String.IsNullOrWhiteSpace(current) Then
                            current = System.Text.RegularExpressions.Regex.Replace(current, "font-family\s*:\s*[^;]+;?", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase).Trim()
                        End If
                        Dim merged As String = cssBody
                        If Not System.String.IsNullOrWhiteSpace(current) Then
                            If Not merged.EndsWith(";", System.StringComparison.Ordinal) Then merged &= ";"
                            merged &= " " & current
                        End If
                        h.SetAttributeValue("style", merged.Trim())
                    Next
                End If

                formattedText = doc.DocumentNode.OuterHtml

                ' --- 5) Construct HTML fragment ---
                Dim htmlHeader As String = "<html><head><meta charset=""UTF-8""></head>" &
                                   $"<body style=""font-family:'{fontName}'; font-size:{fontSize}pt;""><!--StartFragment-->"
                Dim htmlFooter As String = "<!--EndFragment--></body></html>"

                Dim cleanedHtml As String = htmlHeader & formattedText.Trim() & htmlFooter
                cleanedHtml = CreateProperHtml(cleanedHtml).Replace(vbCr, "").Replace(vbLf, "").Replace(vbCrLf, "")

                ' --- 6) CF_HTML clipboard formatting (requires UTF-8 byte offsets) ---
                Dim preamble As String =
            $"Version:0.9{vbCrLf}" &
            $"StartHTML:00000000{vbCrLf}" &
            $"EndHTML:00000000{vbCrLf}" &
            $"StartFragment:00000000{vbCrLf}" &
            $"EndFragment:00000000{vbCrLf}"

                Dim packet As String = preamble & cleanedHtml

                Dim idxHtml As Integer = packet.IndexOf("<html>", System.StringComparison.OrdinalIgnoreCase)
                Dim idxFragStartTag As Integer = packet.IndexOf("<!--StartFragment-->", System.StringComparison.OrdinalIgnoreCase)
                Dim idxFragStart As Integer = idxFragStartTag + "<!--StartFragment-->".Length
                Dim idxFragEnd As Integer = packet.IndexOf("<!--EndFragment-->", System.StringComparison.OrdinalIgnoreCase)

                Dim enc As System.Text.Encoding = System.Text.Encoding.UTF8
                Dim startHtmlOffset As Integer = enc.GetByteCount(packet.Substring(0, idxHtml))
                Dim startFragmentOffset As Integer = enc.GetByteCount(packet.Substring(0, idxFragStart))
                Dim endFragmentOffset As Integer = enc.GetByteCount(packet.Substring(0, idxFragEnd))
                Dim endHtmlOffset As Integer = enc.GetByteCount(packet)

                Dim finalHtml As String = packet _
            .Replace("StartHTML:00000000", $"StartHTML:{startHtmlOffset:D8}") _
            .Replace("EndHTML:00000000", $"EndHTML:{endHtmlOffset:D8}") _
            .Replace("StartFragment:00000000", $"StartFragment:{startFragmentOffset:D8}") _
            .Replace("EndFragment:00000000", $"EndFragment:{endFragmentOffset:D8}")

                System.Diagnostics.Debug.WriteLine("FinalHTML=" & finalHtml)

                Dim savedClipboard As System.Windows.Forms.IDataObject = ClipboardSnapshot.Capture()
                Try

                    ' Set clipboard on STA with short retries (clipboard can be locked)
                    Dim setOk As Boolean = False
                    Dim clipboardThread As New System.Threading.Thread(
                                        Sub()
                                            For attempt As Integer = 1 To 6
                                                Try
                                                    System.Windows.Forms.Clipboard.SetText(finalHtml, System.Windows.Forms.TextDataFormat.Html)
                                                    setOk = True
                                                    Exit For
                                                Catch exClip As System.Runtime.InteropServices.ExternalException
                                                    System.Threading.Thread.Sleep(50 * attempt)
                                                Catch exAny As System.Exception
                                                    ' Unexpected – still retry
                                                    System.Threading.Thread.Sleep(50 * attempt)
                                                End Try
                                            Next
                                        End Sub)
                    clipboardThread.SetApartmentState(System.Threading.ApartmentState.STA)
                    clipboardThread.Start()
                    clipboardThread.Join()

                    If Not setOk Then
                        Throw New System.Exception("HTML could not be written to the clipboard (clipboard locked?).")
                    End If

                    ' Small delay to ensure Word reads stable data
                    System.Threading.Thread.Sleep(50)

                    ' --- 7) Paste into the Word range (with retries for timing issues) ---
                    range.Select()
                    Dim pasted As Boolean = False
                    For attempt As Integer = 1 To 4
                        Try
                            If ReplaceSelection Then
                                range.Application.Selection.PasteAndFormat(Microsoft.Office.Interop.Word.WdRecoveryType.wdFormatOriginalFormatting)
                            Else
                                range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
                                range.Select()
                                range.Application.Selection.PasteAndFormat(Microsoft.Office.Interop.Word.WdRecoveryType.wdFormatOriginalFormatting)
                            End If
                            pasted = True
                            Exit For
                        Catch exPaste As System.Runtime.InteropServices.COMException
                            System.Threading.Thread.Sleep(50 * attempt)
                        End Try
                    Next

                    If Not pasted Then
                        Throw New System.Exception("Pasting into Word failed.")
                    End If

                    System.Threading.Thread.Sleep(100)
                    range = range.Application.Selection.Range

                    ' --- 8) Optionally remove last newline character ---
                    '     Only delete if the trailing character is actually a paragraph mark,
                    '     not real content. PasteAndFormat does not always append a trailing CR.
                    If ReplaceSelection AndAlso NoTrailingCR Then
                        Dim insertedRange As Microsoft.Office.Interop.Word.Range = range.Application.Selection.Range
                        Dim delRng As Microsoft.Office.Interop.Word.Range = insertedRange.Duplicate()
                        delRng.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
                        delRng.MoveStart(Microsoft.Office.Interop.Word.WdUnits.wdCharacter, -1)

                        ' Only delete if the character is a paragraph mark (vbCr) or line feed
                        Dim trailingChar As String = delRng.Text
                        If trailingChar = vbCr OrElse trailingChar = vbLf Then
                            delRng.Delete()
                        End If

                        insertedRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
                        insertedRange.Select()
                    End If

                Finally
                    System.Threading.Thread.Sleep(100)
                    ClipboardSnapshot.Restore(savedClipboard)
                End Try

            Catch ex As System.Exception
                System.Windows.Forms.MessageBox.Show("InsertTextWithFormat Error: " & ex.Message)
            End Try
        End Sub


        ''' <summary>
        ''' Converts HTML <c>&lt;mark&gt;</c> tags into <c>&lt;span&gt;</c> tags that use Word-compatible highlight styles.
        ''' </summary>
        ''' <param name="html">The input HTML.</param>
        ''' <param name="defaultColor">The default highlight token to apply for plain <c>&lt;mark&gt;</c> tags.</param>
        ''' <returns>The transformed HTML string.</returns>
        ''' <summary>
        ''' Converts HTML <c>&lt;mark&gt;</c> tags into <c>&lt;span&gt;</c> tags that use Word-compatible highlight styles.
        ''' </summary>
        ''' <param name="html">The input HTML.</param>
        ''' <param name="defaultColor">The default highlight token to apply for plain <c>&lt;mark&gt;</c> tags.</param>
        ''' <returns>The transformed HTML string.</returns>
        ''' <summary>
        ''' Converts HTML <c>&lt;mark&gt;</c> tags into <c>&lt;span&gt;</c> tags that use Word-compatible highlight styles.
        ''' </summary>
        ''' <param name="html">The input HTML.</param>
        ''' <param name="defaultColor">The default highlight token to apply for plain <c>&lt;mark&gt;</c> tags.</param>
        ''' <returns>The transformed HTML string.</returns>
        Private Shared Function FixMarkTagsForWord(html As String, Optional defaultColor As String = "yellow") As String
            If String.IsNullOrEmpty(html) Then Return html

            ' --- Decode HTML-encoded <mark> tags first ---
            ' Handle &lt;mark&gt; ... &lt;/mark&gt; (plain)
            html = html.Replace("&lt;mark&gt;", "<mark>")
            html = html.Replace("&lt;/mark&gt;", "</mark>")

            ' Handle &lt;mark data-ri-color=&quot;...&quot;&gt; variants (with HTML-encoded quotes)
            html = System.Text.RegularExpressions.Regex.Replace(
                        html,
                        "&lt;mark\s+data-ri-color\s*=\s*(?:&quot;|&#39;|[""'])([^&""']+)(?:&quot;|&#39;|[""'])\s*&gt;",
                        Function(m)
                            Dim color = m.Groups(1).Value.Trim().ToLowerInvariant()
                            Dim css = MsoHighlightToCssColor(color)
                            Return $"<span style=""background:{css}; mso-highlight:{color}"">"
                        End Function,
                        RegexOptions.IgnoreCase)

            ' Handle &lt;mark data-ri-color=...&gt; without quotes (edge case)
            html = System.Text.RegularExpressions.Regex.Replace(
                        html,
                        "&lt;mark\s+data-ri-color\s*=\s*([^&\s>]+)\s*&gt;",
                        Function(m)
                            Dim color = m.Groups(1).Value.Trim().ToLowerInvariant()
                            Dim css = MsoHighlightToCssColor(color)
                            Return $"<span style=""background:{css}; mso-highlight:{color}"">"
                        End Function,
                        RegexOptions.IgnoreCase)

            Dim opts As RegexOptions = RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant Or RegexOptions.Singleline

            ' 1) Convert <mark data-ri-color="...">...</mark> → <span style="background:css; mso-highlight:token">...</span>
            html = System.Text.RegularExpressions.Regex.Replace(
                        html,
                        "<\s*mark\b[^>]*data-ri-color\s*=\s*['""]?(?<color>[^'""\s>]+)['""]?[^>]*>",
                        Function(m As Match)
                            Dim token = m.Groups("color").Value.Trim().ToLowerInvariant()
                            Dim css = MsoHighlightToCssColor(token)
                            Return $"<span style=""background:{css}; mso-highlight:{token}"">"
                        End Function,
                        opts)

            ' 2) Convert plain <mark>...</mark> (yellow) → <span style="...">...</span>
            html = System.Text.RegularExpressions.Regex.Replace(
                        html,
                        "<\s*mark\s*>",
                        "<span style=""mso-highlight:yellow"">",
                        opts)

            ' 3) Close tags
            html = System.Text.RegularExpressions.Regex.Replace(html, "</\s*mark\s*>", "</span>", opts)

            Return html
        End Function

        ''' <summary>
        ''' Maps a Word <c>mso-highlight</c> token to a CSS background color keyword.
        ''' </summary>
        ''' <param name="mso">The Word highlight token (for example, <c>yellow</c>).</param>
        ''' <returns>A CSS color keyword that can be used for a background fill.</returns>
        Private Shared Function MsoHighlightToCssColor(mso As String) As String
            Select Case mso
                Case "yellow" : Return "yellow"
                Case "brightgreen" : Return "lime"
                Case "turquoise" : Return "aqua"
                Case "pink" : Return "fuchsia"
                Case "blue" : Return "blue"
                Case "red" : Return "red"
                Case "darkblue" : Return "navy"
                Case "teal" : Return "teal"
                Case "green" : Return "green"
                Case "violet" : Return "purple"
                Case "darkred" : Return "maroon"
                Case "darkyellow" : Return "olive"
                Case "gray50" : Return "gray"
                Case "gray25" : Return "silver"
                Case "black" : Return "black"
                Case Else : Return "yellow"
            End Select
        End Function


        ''' <summary>
        ''' Removes a trailing carriage return or line feed from the end of a Word range (up to the last 4 characters).
        ''' </summary>
        ''' <param name="range">The Word range to modify.</param>
        Public Shared Sub RemoveTrailingCr(ByRef range As Microsoft.Office.Interop.Word.Range)
            Try
                ' Check a maximum of the last 4 characters
                Dim maxCheck As Integer = Math.Min(4, range.Characters.Count)
                For i As Integer = 1 To maxCheck
                    ' Index of the i-th last character
                    Dim idx As Integer = range.Characters.Count - i + 1
                    If range.Characters(idx).Text = vbCr Or range.Characters(idx).Text = vbLf Then
                        ' Delete the found paragraph mark and stop
                        range.Characters(idx).Delete()
                        Exit For
                    End If
                Next
            Catch ex As System.Exception
                System.Windows.Forms.MessageBox.Show("RemoveTrailingCr Error: " & ex.Message)
            End Try
        End Sub


        ''' <summary>
        ''' Removes HTML tags from the provided HTML and returns the decoded plain text.
        ''' </summary>
        ''' <param name="html">The HTML input.</param>
        ''' <returns>Plain text with HTML entities decoded.</returns>
        Public Shared Function RemoveHTML(html As String) As String

            If String.IsNullOrEmpty(html) Then
                Return String.Empty
            End If

            ' Replace <br> and </p> with vbCrLf.
            ' Handle variations like <br>, <br/>, <br />, and </p> in a case-insensitive manner
            html = Regex.Replace(html, "</p>", vbCrLf, RegexOptions.IgnoreCase)
            html = Regex.Replace(html, "<br\s*/?>", vbCrLf, RegexOptions.IgnoreCase)

            ' Load into HtmlAgilityPack to remove remaining tags and handle entities
            Dim doc As New HtmlAgilityPack.HtmlDocument()
            doc.LoadHtml(html)

            ' Get the inner text (this strips out all remaining HTML tags)
            Dim textContent As String = doc.DocumentNode.InnerText

            ' Decode HTML entities (including special characters and umlauts)
            ' HtmlEntity.DeEntitize converts HTML encoded characters to their decoded form
            textContent = HtmlEntity.DeEntitize(textContent)

            ' Remove extra line breaks or whitespace caused by replaced tags            
            textContent = Regex.Replace(textContent, "(?<!\\)\\[rnt]", Function(m)
                                                                           Select Case m.Value
                                                                               Case "\n" : Return vbLf
                                                                               Case "\r" : Return vbCr
                                                                               Case "\t" : Return vbTab
                                                                               Case Else : Return m.Value
                                                                           End Select
                                                                       End Function)

            ' Trim leading and trailing whitespace
            textContent = textContent.Trim()

            Return textContent
        End Function



        ''' <summary>
        ''' Converts text with custom change markers into an RTF document string.
        ''' </summary>
        ''' <param name="inputText">The input text containing <c>[DEL_START]..[DEL_END]</c> and/or <c>[INS_START]..[INS_END]</c> markers.</param>
        ''' <returns>An RTF string representing the input with basic formatting.</returns>
        Public Shared Function ConvertMarkupToRTF(inputText As String) As String
            ' Define the RTF header with font and color tables
            Dim rtfHeader As String =
                    "{\rtf1\ansi\deff0" &
                    "{\fonttbl{\f0\fnil\fcharset0 Calibri;}}" &
                    "{\colortbl;\red0\green0\blue0;\red0\green0\blue255;\red255\green0\blue0;}" &
                    "\f0\fs20\cf1 "

            ' Replace custom markup with RTF formatting
            Dim rtfContent As String = inputText.Replace(vbCrLf, "\r\n").Replace(vbCr, "\r").Replace(vbLf, "\n")

            ' Convert [DEL_START] ... [DEL_END] to red + strikethrough
            rtfContent = Regex.Replace(rtfContent, "\[DEL_START\](.*?)\[DEL_END\]", "{\cf3\strike $1}{\strike0}", RegexOptions.Singleline)

            ' Convert [INS_START] ... [INS_END] to blue + underline
            rtfContent = Regex.Replace(rtfContent, "\[INS_START\](.*?)\[INS_END\]", "{\cf2\ul $1}{\ul0}", RegexOptions.Singleline)

            ' Convert newlines to RTF paragraph breaks
            rtfContent = Regex.Replace(rtfContent, "(?<!\\)\\r\\n", "\par ")
            rtfContent = Regex.Replace(rtfContent, "(?<!\\)\\r", "\par ")
            rtfContent = Regex.Replace(rtfContent, "(?<!\\)\\n", "\par ")

            ' Add RTF footer
            Dim rtfFooter As String = "}"

            ' Combine and return the full RTF string
            Return rtfHeader & rtfContent & rtfFooter
        End Function

        ''' <summary>
        ''' Normalizes HTML by ensuring required elements exist, encoding text nodes, and removing <c>&lt;TEXTTOPROCESS&gt;</c> wrappers.
        ''' </summary>
        ''' <param name="inputHtml">The input HTML.</param>
        ''' <returns>The normalized HTML.</returns>
        Public Shared Function CreateProperHtml(inputHtml As String) As String
            ' 0) Normalize typographic quotes
            inputHtml = inputHtml _
                .Replace("„"c, """"c) _
                .Replace(ChrW(&H201C), """"c) _
                .Replace(ChrW(&H201D), """"c)

            ' 1) Mask entities: store all &...; sequences and replace with placeholders
            Dim entityPattern As New System.Text.RegularExpressions.Regex("(&#\d+;|&[A-Za-z]+;)")
            Dim entities As New List(Of String)
            inputHtml = entityPattern.Replace(inputHtml,
        Function(m As System.Text.RegularExpressions.Match)
            entities.Add(m.Value)
            Return "###ENTITY" & (entities.Count - 1) & "###"
        End Function)

            ' 2) Remove <TEXTTOPROCESS> wrapper
            inputHtml = inputHtml.Replace("<TEXTTOPROCESS>", "") _
                         .Replace("</TEXTTOPROCESS>", "")

            ' 3) Load HTML
            Dim htmlDoc As New HtmlAgilityPack.HtmlDocument()
            htmlDoc.LoadHtml(inputHtml)

            ' 4) Ensure <head>
            Dim headTag = htmlDoc.DocumentNode.SelectSingleNode("//head")
            If headTag Is Nothing Then
                headTag = HtmlAgilityPack.HtmlNode.CreateNode("<head></head>")
                Dim htmlTag = htmlDoc.DocumentNode.SelectSingleNode("//html")
                If htmlTag Is Nothing Then
                    htmlTag = HtmlAgilityPack.HtmlNode.CreateNode("<html></html>")
                    htmlDoc.DocumentNode.AppendChild(htmlTag)
                End If
                htmlTag.PrependChild(headTag)
            End If

            ' 5) Insert <meta charset="UTF-8"> if not present
            If Not headTag.InnerHtml.Contains("charset") Then
                headTag.InnerHtml = "<meta charset=""UTF-8"">" & headTag.InnerHtml
            End If

            ' 6) Encode all text nodes
            For Each textNode As HtmlAgilityPack.HtmlNode In
            htmlDoc.DocumentNode.DescendantsAndSelf() _
                   .Where(Function(n) n.NodeType = HtmlAgilityPack.HtmlNodeType.Text)

                Dim rawText As String = textNode.InnerText
                textNode.InnerHtml = HtmlEncodeAll(rawText)
            Next

            ' 7) Render HTML
            Dim result As String = htmlDoc.DocumentNode.OuterHtml

            ' 8) Restore masked entities
            result = System.Text.RegularExpressions.Regex.Replace(result, "###ENTITY(\d+)###",
        Function(m As System.Text.RegularExpressions.Match)
            Return entities(Integer.Parse(m.Groups(1).Value))
        End Function)

            Return result
        End Function

        ''' <summary>
        ''' Encodes reserved HTML characters and all non-ASCII characters (&gt; 127) as numeric entities.
        ''' </summary>
        ''' <param name="s">The input string.</param>
        ''' <returns>The encoded string.</returns>
        Private Shared Function HtmlEncodeAll(s As String) As String
            Dim sb As New System.Text.StringBuilder()
            For Each c As Char In s
                Select Case c
                    Case "<"c : sb.Append("&lt;")
                    Case ">"c : sb.Append("&gt;")
                    Case "&"c : sb.Append("&amp;")
                    Case """"c : sb.Append("&quot;")
                    Case "'"c : sb.Append("&#39;")
                    Case Else
                        Dim code = AscW(c)
                        If code > 127 Then
                            sb.Append("&#" & code & ";")
                        Else
                            sb.Append(c)
                        End If
                End Select
            Next
            Return sb.ToString()
        End Function



        ''' <summary>
        ''' Exports a Word range as filtered HTML and returns a simplified HTML string.
        ''' </summary>
        ''' <param name="range">The Word range to export.</param>
        ''' <returns>Simplified HTML for the provided range.</returns>
        Public Shared Function GetRangeHtml(ByVal range As Microsoft.Office.Interop.Word.Range) As String
            Dim htmlContent As String = String.Empty
            Dim tempFile As String = System.IO.Path.GetTempFileName()

            Try
                ' Save the range as a filtered HTML file
                range.ExportFragment(FileName:=tempFile, Format:=WdSaveFormat.wdFormatFilteredHTML)

                ' Read the HTML content
                htmlContent = System.IO.File.ReadAllText(tempFile)
            Finally
                ' Delete the temporary file
                If System.IO.File.Exists(tempFile) Then
                    System.IO.File.Delete(tempFile)
                End If
            End Try

            htmlContent = SimplifyHtml(htmlContent)

            Return htmlContent
        End Function

        ''' <summary>
        ''' Simplifies HTML by removing non-whitelisted tags/attributes and stripping real line breaks.
        ''' </summary>
        ''' <param name="htmlContent">The HTML to simplify.</param>
        ''' <returns>The simplified HTML.</returns>
        Public Shared Function SimplifyHtml(htmlContent As String) As String
            ' Load the HTML content into an HtmlDocument
            Dim htmlDoc As New HtmlAgilityPack.HtmlDocument()
            htmlDoc.LoadHtml(htmlContent)

            ' Process the document to remove irrelevant tags and attributes
            CleanHtmlNode(htmlDoc.DocumentNode)

            ' Get the simplified HTML
            Dim simplifiedHtml As String = htmlDoc.DocumentNode.OuterHtml

            ' Remove real line breaks
            simplifiedHtml = simplifiedHtml.Replace(vbCr, "").Replace(vbLf, "").Replace(vbCrLf, "")

            ' Return the simplified HTML
            Return simplifiedHtml
        End Function

        ''' <summary>
        ''' Cleans an HTML node tree by removing non-whitelisted elements and non-whitelisted attributes.
        ''' </summary>
        ''' <param name="node">The node to clean (processed recursively).</param>
        Public Shared Sub CleanHtmlNode(node As HtmlNode)
            If node.NodeType = HtmlNodeType.Element Then
                ' Define the allowed tags
                Dim allowedTags As HashSet(Of String) = New HashSet(Of String) From {"b", "strong", "i", "em", "u", "font", "span", "p", "ul", "ol", "li", "br"}

                ' Define the allowed attributes
                Dim allowedAttributes As HashSet(Of String) = New HashSet(Of String) From {"style", "class"}

                ' Remove attributes that are not in the allowed list
                For Each attr In node.Attributes.ToList()
                    If Not allowedAttributes.Contains(attr.Name.ToLower()) Then
                        node.Attributes.Remove(attr.Name)
                    End If
                Next

                ' If the node is not an allowed tag, replace it with its inner content
                If Not allowedTags.Contains(node.Name.ToLower()) Then
                    Dim parentNode = node.ParentNode
                    Dim innerNodes = node.ChildNodes.ToList()
                    For Each innerNode In innerNodes
                        If innerNode.Name.ToLower() = "p" OrElse innerNode.Name.ToLower() = "br" Then
                            parentNode.InsertBefore(HtmlNode.CreateNode(innerNode.OuterHtml), node)
                        Else
                            parentNode.InsertBefore(innerNode, node)
                        End If
                    Next
                    parentNode.RemoveChild(node)
                End If
            End If

            ' Recursively process child nodes
            For Each childNode In node.ChildNodes.ToList()
                CleanHtmlNode(childNode)
            Next
        End Sub


        ''' <summary>
        ''' Removes a subset of Markdown formatting markers while preserving bracketed and brace-delimited regions verbatim.
        ''' </summary>
        ''' <param name="input">The input string.</param>
        ''' <returns>The input with selected Markdown markers removed.</returns>
        ''' <exception cref="System.Exception">Thrown when processing fails.</exception>
        Public Shared Function RemoveMarkdownFormatting(ByVal input As System.String) As System.String
            Try
                If input Is Nothing Then
                    Return Nothing
                End If
                If input.Length = 0 Then
                    Return System.String.Empty
                End If

                ' --- lazily-initialized, compiled regexes (cached across calls) ---
                Static rxBoldItalic As System.Text.RegularExpressions.Regex = Nothing
                Static rxBold As System.Text.RegularExpressions.Regex = Nothing
                Static rxItalic As System.Text.RegularExpressions.Regex = Nothing
                Static rxStrike As System.Text.RegularExpressions.Regex = Nothing
                Static rxHeadings As System.Text.RegularExpressions.Regex = Nothing

                If rxBoldItalic Is Nothing Then
                    rxBoldItalic = New System.Text.RegularExpressions.Regex("\*\*\*(.+?)\*\*\*", System.Text.RegularExpressions.RegexOptions.Singleline Or System.Text.RegularExpressions.RegexOptions.Compiled Or System.Text.RegularExpressions.RegexOptions.CultureInvariant)
                End If
                If rxBold Is Nothing Then
                    rxBold = New System.Text.RegularExpressions.Regex("\*\*(.+?)\*\*", System.Text.RegularExpressions.RegexOptions.Singleline Or System.Text.RegularExpressions.RegexOptions.Compiled Or System.Text.RegularExpressions.RegexOptions.CultureInvariant)
                End If
                If rxItalic Is Nothing Then
                    rxItalic = New System.Text.RegularExpressions.Regex("(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)", System.Text.RegularExpressions.RegexOptions.Singleline Or System.Text.RegularExpressions.RegexOptions.Compiled Or System.Text.RegularExpressions.RegexOptions.CultureInvariant)
                End If
                If rxStrike Is Nothing Then
                    rxStrike = New System.Text.RegularExpressions.Regex("~~(.+?)~~", System.Text.RegularExpressions.RegexOptions.Singleline Or System.Text.RegularExpressions.RegexOptions.Compiled Or System.Text.RegularExpressions.RegexOptions.CultureInvariant)
                End If
                If rxHeadings Is Nothing Then
                    rxHeadings = New System.Text.RegularExpressions.Regex("^[ \t]*#{1,6}[ \t]+(.+?)(?:[ \t]+#+)?[ \t]*(\r?\n|$)", System.Text.RegularExpressions.RegexOptions.Multiline Or System.Text.RegularExpressions.RegexOptions.Compiled Or System.Text.RegularExpressions.RegexOptions.CultureInvariant)
                End If
                ' --- end regex cache ---

                ' 1) Find protected regions ([...] and {...}) with nesting
                Dim regions As System.Collections.Generic.List(Of System.ValueTuple(Of System.Int32, System.Int32)) = New System.Collections.Generic.List(Of System.ValueTuple(Of System.Int32, System.Int32))()
                Dim stack As System.Collections.Generic.Stack(Of System.Char) = New System.Collections.Generic.Stack(Of System.Char)()
                Dim startIdx As System.Int32 = -1

                For i As System.Int32 = 0 To input.Length - 1
                    Dim ch As System.Char = input(i)
                    If ch = "["c OrElse ch = "{"c Then
                        If stack.Count = 0 Then
                            startIdx = i
                        End If
                        stack.Push(ch)
                    ElseIf ch = "]"c OrElse ch = "}"c Then
                        If stack.Count > 0 Then
                            Dim opener As System.Char = stack.Peek()
                            Dim matches As System.Boolean = (opener = "["c AndAlso ch = "]"c) OrElse (opener = "{"c AndAlso ch = "}"c)
                            If matches Then
                                stack.Pop()
                                If stack.Count = 0 AndAlso startIdx >= 0 Then
                                    regions.Add((startIdx, i)) ' inclusive
                                    startIdx = -1
                                End If
                            End If
                        End If
                    End If
                Next

                ' 2) Mask protected regions with placeholders
                Dim masked As System.Text.StringBuilder = New System.Text.StringBuilder(input.Length + (regions.Count * 16))
                Dim placeholders As System.Collections.Generic.List(Of System.String) = New System.Collections.Generic.List(Of System.String)(regions.Count)
                Dim originals As System.Collections.Generic.List(Of System.String) = New System.Collections.Generic.List(Of System.String)(regions.Count)

                Dim lastPos As System.Int32 = 0
                For idx As System.Int32 = 0 To regions.Count - 1
                    Dim r = regions(idx)
                    If r.Item1 > lastPos Then
                        masked.Append(input, lastPos, r.Item1 - lastPos)
                    End If
                    Dim original As System.String = input.Substring(r.Item1, r.Item2 - r.Item1 + 1)
                    Dim token As System.String = "__BRMASK_" & idx.ToString(System.Globalization.CultureInfo.InvariantCulture) & "_X__"
                    masked.Append(token)
                    placeholders.Add(token)
                    originals.Add(original)
                    lastPos = r.Item2 + 1
                Next
                If lastPos < input.Length Then
                    masked.Append(input, lastPos, input.Length - lastPos)
                End If

                Dim work As System.String = masked.ToString()

                ' 3) Strip markdown on the masked text (outside protected regions)
                work = rxBoldItalic.Replace(work, "$1")
                work = rxBold.Replace(work, "$1")
                work = rxItalic.Replace(work, "$1")
                work = rxStrike.Replace(work, "$1")
                work = rxHeadings.Replace(work, "$1$2")

                ' 4) Restore protected regions verbatim
                For i As System.Int32 = 0 To placeholders.Count - 1
                    work = work.Replace(placeholders(i), originals(i))
                Next

                Return work

            Catch ex As System.Exception
                Throw New System.Exception("Error in RemoveMarkdownFormatting: " & ex.Message, ex)
            End Try
        End Function

        ''' <summary>
        ''' Inserts text into a Word selection and applies bold formatting to sections delimited by <c>**</c>.
        ''' </summary>
        ''' <param name="selection">The Word selection to insert into.</param>
        ''' <param name="gptResult">The input text containing <c>**</c> bold markers.</param>
        Public Shared Sub InsertTextWithBoldMarkers(selection As Microsoft.Office.Interop.Word.Selection, gptResult As String)

            ' Save the starting position of the insertion
            Dim startPosition As Integer = selection.Start

            ' Split the text by "**" to identify bold and regular sections
            Dim parts() As String
            parts = Split(gptResult, "**")

            ' Iterate through the parts and add text with appropriate formatting
            For i As Integer = 0 To UBound(parts)
                If i Mod 2 = 1 Then
                    ' Odd-index parts are bold
                    selection.Font.Bold = -1 ' True
                Else
                    ' Even-index parts are normal text
                    selection.Font.Bold = 0 ' False
                End If

                ' Insert the text part
                If parts(i) <> "" Then
                    selection.TypeText(parts(i))
                End If
            Next

            ' Reset bold formatting to normal after insertion
            selection.Font.Bold = 0 ' False

            ' Save the end position of the insertion
            Dim endPosition As Integer = selection.Start

            ' Select the entire inserted text
            selection.SetRange(startPosition, endPosition)
        End Sub


    End Class
End Namespace
