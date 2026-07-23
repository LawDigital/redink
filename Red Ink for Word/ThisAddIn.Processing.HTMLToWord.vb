' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.Processing.HTMLToWord.vb
' Purpose: Converts HtmlAgilityPack nodes into formatted Microsoft Word content for the
'          Red Ink add-in, including inline styling, block elements, media, and links.
'
' Architecture:
'  - Inline Rendering: RenderInline normalizes text nodes, propagates hyperlink context,
'    applies cumulative formatting delegates, and routes emojis, images, breaks, and spans.
'  - Block Rendering: ParseHtmlNode walks the DOM depth-first to translate paragraphs, headings,
'    blockquotes, lists, definition lists, tables, code blocks, inputs, and anchors into Word ranges.
'  - List State Tracking: ulLevels/ulStartPos record indentation levels so bullet and ordered lists
'    receive consistent Word list formatting across nested structures.
'  - Media Handling: InsertImageFromSrc resolves local paths or downloads remote images, inserts
'    InlineShape instances, and logs failures without throwing.
'  - Helper Utilities: CombineStyle chains formatting delegates, InsertInline/TrueInsertInline centralize
'    styled text insertion and hyperlink creation, RemoveTrailingParagraph trims trailing block nodes.
' Dependencies:
'  - HtmlAgilityPack for DOM access, Microsoft.Office.Interop.Word for document automation,
'    System.Net for remote image downloads, System.Diagnostics for tracing.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.Text.RegularExpressions
Imports HtmlAgilityPack
Imports Microsoft.Office.Interop.Word

Partial Public Class ThisAddIn

    Private Shared emojiSet As New HashSet(Of String)()
    Private Shared ReadOnly _emojiPairRegex As New System.Text.RegularExpressions.Regex(
        "[\uD83C-\uDBFF][\uDC00-\uDFFF]",
        System.Text.RegularExpressions.RegexOptions.Compiled Or
        System.Text.RegularExpressions.RegexOptions.CultureInvariant)

    ''' <summary>
    ''' Combines two formatting delegates without discarding the originals.
    ''' </summary>
    Private Shared Function CombineStyle(
        baseAction As Action(Of Microsoft.Office.Interop.Word.Range),
        additional As Action(Of Microsoft.Office.Interop.Word.Range)
    ) As Action(Of Microsoft.Office.Interop.Word.Range)

        If baseAction Is Nothing Then Return additional
        If additional Is Nothing Then Return baseAction

        Return Sub(rng As Microsoft.Office.Interop.Word.Range)
                   baseAction(rng)
                   additional(rng)
               End Sub
    End Function

    ''' <summary>
    ''' Renders inline nodes (recursively when needed) with cumulative formatting and optional hyperlink context.
    ''' </summary>
    Private Shared Sub RenderInline(
        node As HtmlAgilityPack.HtmlNode,
        rng As Microsoft.Office.Interop.Word.Range,
        styleAction As Action(Of Microsoft.Office.Interop.Word.Range),
        inheritedHref As String
    )

        ' Ignore comment nodes.
        If node.NodeType = HtmlAgilityPack.HtmlNodeType.Comment Then Return

        ' ------------------------------------------------- Leaf: #text -------------
        'If node.NodeType = HtmlAgilityPack.HtmlNodeType.Text Then
        'Dim txt As String = HtmlEntity.DeEntitize(node.InnerText)
        'If Not String.IsNullOrWhiteSpace(txt) Then
        'InsertInline(rng, txt, styleAction, inheritedHref)
        'End If
        'Return
        'End If

        ' ─── Newline handling: split only real line breaks with content.
        If node.NodeType = HtmlAgilityPack.HtmlNodeType.Text Then
            Dim rawText = HtmlEntity.DeEntitize(node.InnerText)
            Dim hasNewline = rawText.IndexOfAny({vbCr(0), vbLf(0)}) >= 0
            Dim stripped = rawText.Replace(vbCr, "").Replace(vbLf, "")

            ' 1) Whitespace-only line breaks → skip entirely.
            If hasNewline AndAlso String.IsNullOrWhiteSpace(stripped) Then
                Return
            End If

            ' 2) Mixed text with real newlines → split and insert per break.
            If hasNewline Then
                Dim parts = rawText.Split(
                    New String() {vbCrLf, vbCr, vbLf},
                    StringSplitOptions.None)
                For i = 0 To parts.Length - 1
                    Dim segment = parts(i)
                    If Not String.IsNullOrWhiteSpace(segment) Then
                        InsertInline(rng, segment, styleAction, inheritedHref)
                    End If
                    If i < parts.Length - 1 Then
                        rng.InsertAfter(vbCr)
                        rng.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                    End If
                Next
                Return
            End If

            ' 3) No newline → insert normally.
            If Not String.IsNullOrWhiteSpace(rawText) Then
                InsertInline(rng, rawText, styleAction, inheritedHref)
            End If
            Return
        End If

        ' ------------------------------------------------- Leaf: <br> --------------
        'If node.Name.Equals("br", StringComparison.OrdinalIgnoreCase) Then
        'rng.Font.Reset()
        'rng.Text = vbCr
        'rng.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
        'Return
        'End If

        If node.Name.Equals("br", StringComparison.OrdinalIgnoreCase) Then
            ' Soft line break in Word (Shift+Enter) instead of a hard paragraph.
            rng.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
            rng.InsertAfter(ChrW(11))  ' Manual line break.
            rng.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
            Return
        End If

        ' ------------------------------------------------- Leaf: <img> -------------
        If node.Name.Equals("img", StringComparison.OrdinalIgnoreCase) Then
            Dim src As String = node.GetAttributeValue("src", String.Empty)
            If Not String.IsNullOrWhiteSpace(src) Then
                ' Instead of calling InlineShapes.AddPicture directly,
                ' call the resilient helper.
                InsertImageFromSrc(rng, src)
            End If
            Return
        End If

        ' ------------------------------------------------- Leaf/Semi-Leaf: <a> -----
        Dim thisHref As String = inheritedHref
        If node.Name.Equals("a", StringComparison.OrdinalIgnoreCase) Then
            thisHref = node.GetAttributeValue("href", String.Empty)

            ' Output simple text-only anchors immediately …
            If node.ChildNodes.All(Function(c) c.NodeType = HtmlAgilityPack.HtmlNodeType.Text) Then
                Dim txtLink As String = HtmlEntity.DeEntitize(node.InnerText)
                InsertInline(rng, txtLink, styleAction, thisHref)
                Return
            End If
            ' … otherwise render the children with the same href.
        End If

        ' ------------------------------------------------- Style routing ------------
        Select Case node.Name.ToLowerInvariant()
            Case "strong", "b"
                styleAction = CombineStyle(styleAction,
                                           Sub(r) r.Font.Bold = True)

            Case "em", "i"
                styleAction = CombineStyle(styleAction,
                                           Sub(r) r.Font.Italic = True)

            Case "u"
                styleAction = CombineStyle(styleAction,
                                           Sub(r) r.Font.Underline = Word.WdUnderline.wdUnderlineSingle)

            Case "del", "s"
                styleAction = CombineStyle(styleAction,
                                           Sub(r) r.Font.StrikeThrough = True)

            Case "code"
                styleAction = CombineStyle(styleAction,
                    Sub(r)
                        r.Font.Name = "Courier New"
                        r.Font.Size = 10
                        r.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray25
                    End Sub)

            Case "sub"
                styleAction = CombineStyle(styleAction,
                                           Sub(r) r.Font.Subscript = True)

            Case "sup"
                styleAction = CombineStyle(styleAction,
                                           Sub(r) r.Font.Superscript = True)

            Case "span"
                Dim cls = node.GetAttributeValue("class", String.Empty)
                If cls.Contains("emoji") Then
                    styleAction = CombineStyle(styleAction,
                        Sub(r)
                            r.Font.Name = "Segoe UI Emoji"
                            r.Font.Color = Word.WdColor.wdColorWhite
                            r.Shading.BackgroundPatternColor =
                                System.Drawing.ColorTranslator.ToOle(
                                    System.Drawing.Color.FromArgb(0, 112, 192))
                        End Sub)
                End If
                ' No additional formatting otherwise → pass through.
        End Select

        ' ------------------------------------------------- Recursion ---------------
        For Each child In node.ChildNodes
            RenderInline(child, rng, styleAction, thisHref)
        Next
    End Sub

    ''' <summary>
    ''' Inserts an image from a src attribute into the target Word range by loading local paths or downloading URLs.
    ''' </summary>
    Private Shared Sub InsertImageFromSrc(
        rng As Microsoft.Office.Interop.Word.Range,
        src As String
    )
        If String.IsNullOrWhiteSpace(src) Then Return

        Dim fileName As String = src
        Dim tempFile As String = String.Empty
        Dim isUrl As Boolean = False

        Try
            ' URL detection.
            Dim uri = New System.Uri(src, UriKind.RelativeOrAbsolute)
            If uri.IsAbsoluteUri AndAlso
               (uri.Scheme.Equals("http", StringComparison.OrdinalIgnoreCase) _
                OrElse uri.Scheme.Equals("https", StringComparison.OrdinalIgnoreCase)) Then

                isUrl = True
                tempFile = System.IO.Path.Combine(
                    System.IO.Path.GetTempPath(),
                    System.IO.Path.GetFileName(uri.LocalPath)
                )
                Using client As New System.Net.WebClient()
                    client.DownloadFile(uri, tempFile)
                End Using
                fileName = tempFile
            End If

            ' Verify file existence.
            If Not System.IO.File.Exists(fileName) Then
                Debug.WriteLine(
                    $"Image file not found or download failed: '{fileName}'"
                )
            End If

            ' Insert InlineShape.
            Dim pic As Microsoft.Office.Interop.Word.InlineShape =
                rng.InlineShapes.AddPicture(
                    FileName:=fileName,
                    LinkToFile:=False,
                    SaveWithDocument:=True
                )
            rng.SetRange(pic.Range.End, pic.Range.End)

        Catch ex As System.Exception
            ' Log errors internally without throwing further.
            Debug.WriteLine(
                $"[InsertImageFromSrc] {ex.GetType().FullName}: {ex.Message}"
            )
            rng.InsertAfter("[Image missing]")
        Finally
            ' Remove temporary file.
            If isUrl AndAlso Not String.IsNullOrWhiteSpace(tempFile) Then
                Try
                    System.IO.File.Delete(tempFile)
                Catch ioEx As System.Exception
                    Debug.WriteLine(
                        $"[InsertImageFromSrc] Temporary file could not be deleted: {ioEx.Message}"
                    )
                End Try
            End If
        End Try
    End Sub

    ''' <summary>
    ''' Removes the trailing paragraph or break node when it is redundant at the end of the document node.
    ''' </summary>
    Private Shared Sub RemoveTrailingParagraph(htmlDoc As HtmlAgilityPack.HtmlDocument)

        Dim candidates = New String() {"p", "br"}

        For Each TagName In candidates
            Dim lastNode = htmlDoc.DocumentNode.SelectSingleNode("(//" & TagName & ")[last()]")
            If lastNode Is Nothing Then Continue For

            ' Confirm that the node is truly the last significant node.
            Dim cur = lastNode.NextSibling
            Dim canDelete As Boolean = True
            While cur IsNot Nothing
                Select Case cur.NodeType
                    Case HtmlAgilityPack.HtmlNodeType.Comment
                        ' Ignore.
                    Case HtmlAgilityPack.HtmlNodeType.Text
                        If Not String.IsNullOrWhiteSpace(cur.InnerText) Then
                            canDelete = False : Exit While
                        End If
                    Case Else
                        canDelete = False : Exit While
                End Select
                cur = cur.NextSibling
            End While

            If canDelete Then
                If TagName = "p" Then
                    ' Preserve children and remove the <p> container.
                    For Each child In lastNode.ChildNodes.ToList()
                        lastNode.ParentNode.InsertBefore(child, lastNode)
                    Next
                End If
                lastNode.Remove()
                Exit For            ' Touch only one trailing element.
            End If
        Next
    End Sub

    Private Shared ulLevels As List(Of Integer)
    Private Shared ulStartPos As Integer

    ''' <summary>
    ''' Parses an HtmlNode (and its children) into the provided Word range, tracking list nesting levels.
    ''' </summary>
    Private Shared Sub ParseHtmlNode(
        node As HtmlAgilityPack.HtmlNode,
        range As Microsoft.Office.Interop.Word.Range,
        Optional currentLevel As Integer = 0)

        ' -------------------------------- Text shortcut ---------------------------
        If Not node.HasChildNodes AndAlso node.NodeType = HtmlAgilityPack.HtmlNodeType.Text Then
            RenderInline(node, range, Nothing, String.Empty)
            Return
        End If

        If node.Name.Equals("p", StringComparison.OrdinalIgnoreCase) Then
            ' (1) Render inline content.
            RenderInline(node, range, Nothing, String.Empty)

            ' — find the next real node (skip whitespace/comments) —
            Dim nxt As HtmlAgilityPack.HtmlNode = node.NextSibling
            While nxt IsNot Nothing _
                AndAlso (nxt.NodeType = HtmlAgilityPack.HtmlNodeType.Comment _
                        OrElse (nxt.NodeType = HtmlAgilityPack.HtmlNodeType.Text _
                                AndAlso String.IsNullOrWhiteSpace(nxt.InnerText)))
                nxt = nxt.NextSibling
            End While

            ' — insert paragraph(s) only if something follows —
            If nxt IsNot Nothing Then
                ' Single paragraph break.
                range.InsertParagraphAfter()
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                ' Additional blank paragraph when the next sibling is a <p>.
                If nxt.Name.Equals("p", StringComparison.OrdinalIgnoreCase) Then
                    range.InsertParagraphAfter()
                    range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                End If
            End If

            Return
        End If

        If node.Name.Equals("li", StringComparison.OrdinalIgnoreCase) Then

            Dim isFootnoteEntry As Boolean =
                    node.GetAttributeValue("id", String.Empty) _
                        .StartsWith("fn:", StringComparison.OrdinalIgnoreCase)

            If isFootnoteEntry Then
                ' --- (A) Determine bookmark name and footnote number ---
                Dim rawId As String = node.GetAttributeValue("id", String.Empty)   ' e.g. "fn:1"
                Dim fnNum As String = rawId.Substring(rawId.IndexOf(":"c) + 1)     ' "1"
                Dim bookmarkName As String = "fn" & fnNum                          ' "fn1"

                ' --- (B) Insert superscript number and wrap with bookmark ---
                Dim bmStart As Integer = range.End
                InsertInline(
                                range,
                                fnNum,
                                Sub(r) r.Font.Superscript = True,
                                String.Empty
                            )
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                'Dim bmRange As Word.Range = range.Document.Range(bmStart, range.End)
                Dim bmRange As Word.Range = range.Duplicate
                bmRange.Start = bmStart
                bmRange.End = range.End
                range.Bookmarks.Add(Name:=bookmarkName, Range:=bmRange)

                Debug.WriteLine($"[ParseHtmlNode] Footnote Bookmark '{bookmarkName}' at Range=({bmStart},{range.End})")

                ' Space after the number.
                range.InsertAfter(" ")
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                ' --- (C) Render the complete footnote text (plus back reference arrow) ---
                ' If the <li> contains only a <p>, unwrap it.
                If node.ChildNodes.Count = 1 _
                     AndAlso node.FirstChild.Name.Equals("p", StringComparison.OrdinalIgnoreCase) Then

                    Dim pNode As HtmlAgilityPack.HtmlNode = node.FirstChild
                    For Each subNode As HtmlAgilityPack.HtmlNode In pNode.ChildNodes
                        ParseHtmlNode(subNode, range, currentLevel)
                    Next

                Else

                    For Each subNode As HtmlAgilityPack.HtmlNode In node.ChildNodes

                        Select Case subNode.Name.ToLowerInvariant()

                            Case "p"
                                ' (1) Render inline content (including <br> as manual break).
                                RenderInline(subNode, range, Nothing, String.Empty)

                                ' (2) Complete the paragraph.
                                range.InsertParagraphAfter()
                                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                                ' (3) Check whether <p> or <blockquote> follows (after whitespace/comments).
                                Dim nxt As HtmlAgilityPack.HtmlNode = subNode.NextSibling
                                While nxt IsNot Nothing _
                                          AndAlso (nxt.NodeType = HtmlAgilityPack.HtmlNodeType.Comment _
                                                   OrElse (nxt.NodeType = HtmlAgilityPack.HtmlNodeType.Text _
                                                           AndAlso String.IsNullOrWhiteSpace(nxt.InnerText)))
                                    nxt = nxt.NextSibling
                                End While

                                ' (4) If yes, insert an extra blank line.
                                If nxt IsNot Nothing _
                                       AndAlso (nxt.Name.Equals("p", StringComparison.OrdinalIgnoreCase) _
                                                OrElse nxt.Name.Equals("blockquote", StringComparison.OrdinalIgnoreCase)) Then

                                    range.InsertParagraphAfter()
                                    range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                                End If

                                Exit Select

                            Case "blockquote"
                                ' (1) Render quote paragraphs individually.
                                Dim quoteParas = subNode.SelectNodes("./p")
                                If quoteParas IsNot Nothing Then
                                    For Each pNode As HtmlAgilityPack.HtmlNode In quoteParas

                                        Dim paraStart As Integer = range.Start
                                        RenderInline(pNode, range, Nothing, String.Empty)

                                        range.InsertParagraphAfter()
                                        range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                                        ' Indent the inserted paragraph.
                                        Dim indentRg As Microsoft.Office.Interop.Word.Range = range.Duplicate
                                        indentRg.Start = paraStart
                                        indentRg.End = range.End
                                        'range.Document.Range(paraStart, range.End)

                                        indentRg.ListFormat.RemoveNumbers()
                                        indentRg.ParagraphFormat.LeftIndent +=
                                            indentRg.Application.CentimetersToPoints(0.75)

                                    Next
                                Else
                                    ' Fallback: parse recursively without <p> wrapper.
                                    For Each innerNode As HtmlAgilityPack.HtmlNode In subNode.ChildNodes
                                        ParseHtmlNode(innerNode, range, currentLevel)
                                    Next
                                End If

                                Exit Select

                            ' Render inline elements directly so RenderInline applies bold/other styles.
                            Case "#text", "strong", "b", "em", "i", "u",
                             "del", "s", "sub", "sup", "code", "span", "img", "br", "a"

                                RenderInline(subNode, range, Nothing, String.Empty)

                            ' Skip nested lists here; handled afterward.
                            Case "ul", "ol"
                                ' Intentionally empty.

                                ' All other block elements are parsed recursively.
                            Case Else
                                ParseHtmlNode(subNode, range, currentLevel)

                        End Select

                    Next

                End If

                ' --- (D) Paragraph break after each footnote and return ---
                range.InsertParagraphAfter()
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                Return
            End If

            If currentLevel > 1 Then
                ' --- (0) CR before the LI depending on list type ---
                Dim parentName = node.ParentNode.Name.ToLowerInvariant()
                If parentName = "ol" Or parentName = "ul" Then
                    ' OL: only before sub items except the first.
                    Dim sibs = node.ParentNode.SelectNodes("li")
                    If sibs IsNot Nothing Then
                        Dim idx As Integer = 0
                        For i As Integer = 0 To sibs.Count - 1
                            If sibs(i) Is node Then
                                idx = i
                                Exit For
                            End If
                        Next
                        If idx > 0 Then
                            range.InsertAfter(vbCr)
                            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                        End If
                    End If
                End If
            End If

            ' --- (1) Store level ---
            If ulLevels IsNot Nothing Then
                ulLevels.Add(currentLevel)
            End If

            ' --- (2) Remove <p> wrapper when it is the only direct child ---
            If node.ChildNodes.Count = 1 _
       AndAlso node.FirstChild.Name.Equals("p", StringComparison.OrdinalIgnoreCase) Then

                Dim pNode As HtmlAgilityPack.HtmlNode = node.FirstChild
                For Each subNode As HtmlAgilityPack.HtmlNode In pNode.ChildNodes
                    ParseHtmlNode(subNode, range, currentLevel)
                Next

            Else

                For Each subNode As HtmlAgilityPack.HtmlNode In node.ChildNodes

                    Select Case subNode.Name.ToLowerInvariant()

                        Case "blockquote"
                            ' (1) Render quote paragraphs individually.
                            Dim quoteParas = subNode.SelectNodes("./p")
                            If quoteParas IsNot Nothing Then
                                For Each pNode As HtmlAgilityPack.HtmlNode In quoteParas
                                    Dim paraStart As Integer = range.Start
                                    RenderInline(pNode, range, Nothing, String.Empty)

                                    range.InsertParagraphAfter()
                                    range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                                    ' Indent the inserted paragraph.
                                    Dim indentRg As Microsoft.Office.Interop.Word.Range = range.Duplicate
                                    indentRg.Start = paraStart
                                    indentRg.End = range.End
                                    indentRg.ListFormat.RemoveNumbers()
                                    indentRg.ParagraphFormat.LeftIndent +=
                                        indentRg.Application.CentimetersToPoints(0.75)

                                Next
                            Else
                                ' Fallback: parse recursively without <p> wrapper.
                                For Each innerNode As HtmlAgilityPack.HtmlNode In subNode.ChildNodes
                                    ParseHtmlNode(innerNode, range, currentLevel)
                                Next
                            End If

                            Exit Select

                        ' Render inline elements directly so RenderInline applies bold/other styles.
                        Case "#text", "strong", "b", "em", "i", "u",
                             "del", "s", "sub", "sup", "code", "span", "img", "br", "a"

                            RenderInline(subNode, range, Nothing, String.Empty)

                        ' Skip nested lists here; handled afterward.
                        Case "ul", "ol"
                            ' Intentionally empty.

                            ' All other block elements parsed recursively.
                        Case Else
                            ParseHtmlNode(subNode, range, currentLevel)

                    End Select

                Next

            End If

            ' --- (3) Process nested lists at the end ---
            Dim nestedUl As HtmlAgilityPack.HtmlNode = node.SelectSingleNode("ul")
            Dim nestedOl As HtmlAgilityPack.HtmlNode = node.SelectSingleNode("ol")
            If (nestedUl IsNot Nothing OrElse nestedOl IsNot Nothing) Then
                range.InsertAfter(vbCr)
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
            End If
            If nestedUl IsNot Nothing Then
                ParseHtmlNode(nestedUl, range, currentLevel + 1)
            ElseIf nestedOl IsNot Nothing Then
                ParseHtmlNode(nestedOl, range, currentLevel + 1)
            End If

            If isFootnoteEntry Then
                range.InsertParagraphAfter()
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
            End If

            Return
        End If

        Debug.WriteLine($"[ParseHtmlNode] Enter node=<{node.Name}> Range=({range.Start},{range.End})")

        ' -------------------------------- Main child-node loop ---------------------
        For Each childNode As HtmlAgilityPack.HtmlNode In node.ChildNodes

            Debug.WriteLine($"  └─ Child: <{childNode.Name}> Type={childNode.NodeType}")

            Dim nestedLinkNode As HtmlAgilityPack.HtmlNode = Nothing
            If Not childNode.Name.Equals("a", StringComparison.OrdinalIgnoreCase) Then
                nestedLinkNode = childNode.SelectSingleNode(".//a")
            End If
            Dim nestedHref As String = If(nestedLinkNode IsNot Nothing,
                                      nestedLinkNode.GetAttributeValue("href", String.Empty),
                                      String.Empty)

            Select Case childNode.Name.ToLowerInvariant()

                Case "blockquote"
                    ' (1) Paragraph before the quote.
                    'range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
                    'range.InsertParagraphAfter()
                    range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)

                    ' (2) Process only direct <p> children.
                    Dim quoteParas As HtmlAgilityPack.HtmlNodeCollection =
                        childNode.SelectNodes("./p")

                    If quoteParas IsNot Nothing Then
                        For Each pNode As HtmlAgilityPack.HtmlNode In quoteParas
                            ' Mark the start of the new paragraph.
                            Dim paraStart As Integer = range.Start

                            ' (3) Render inline content of the quote.
                            RenderInline(pNode, range, Nothing, String.Empty)

                            ' (4) Paragraph break after each quote paragraph.
                            range.InsertParagraphAfter()
                            range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)

                            ' (5) Indent the inserted paragraph.
                            Dim indentRg As Microsoft.Office.Interop.Word.Range = range.Duplicate
                            indentRg.Start = paraStart
                            indentRg.End = range.End
                            indentRg.ListFormat.RemoveNumbers()
                            indentRg.ParagraphFormat.LeftIndent +=
                                indentRg.Application.CentimetersToPoints(0.75)

                        Next
                    Else
                        ' Fallback: parse normally when no <p> exists in the blockquote.
                        For Each subNode As HtmlAgilityPack.HtmlNode In childNode.ChildNodes
                            ParseHtmlNode(subNode, range, currentLevel)
                        Next
                    End If

                    Exit Select

                    ' (6) Paragraph after the entire blockquote.
                    range.InsertParagraphAfter()
                    range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)

                    ' (7) Additional blank paragraph only when another <p> follows immediately.
                    Dim nxtBQ As HtmlAgilityPack.HtmlNode = childNode.NextSibling
                    While nxtBQ IsNot Nothing _
                              AndAlso (nxtBQ.NodeType = HtmlAgilityPack.HtmlNodeType.Comment _
                                       OrElse (nxtBQ.NodeType = HtmlAgilityPack.HtmlNodeType.Text _
                                               AndAlso String.IsNullOrWhiteSpace(nxtBQ.InnerText)))
                        nxtBQ = nxtBQ.NextSibling
                    End While

                    If nxtBQ IsNot Nothing _
                           AndAlso nxtBQ.Name.Equals("p", System.StringComparison.OrdinalIgnoreCase) Then

                        range.InsertParagraphAfter()
                        range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
                    End If

                    Exit Select

                Case "p"
                    ' Render inline content (including <br> as manual break).
                    RenderInline(childNode, range, Nothing, String.Empty)

                    ' Locate next real sibling.
                    Dim nxt As HtmlAgilityPack.HtmlNode = childNode.NextSibling
                    While nxt IsNot Nothing _
                          AndAlso (nxt.NodeType = HtmlAgilityPack.HtmlNodeType.Comment _
                                   OrElse (nxt.NodeType = HtmlAgilityPack.HtmlNodeType.Text _
                                           AndAlso String.IsNullOrWhiteSpace(nxt.InnerText)))
                        nxt = nxt.NextSibling
                    End While

                    ' Insert paragraph break only if something follows.
                    If nxt IsNot Nothing Then
                        range.InsertParagraphAfter()
                        range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                        ' Insert an extra blank line when the sibling is another <p> or <blockquote>.
                        If nxt.Name.Equals("p", StringComparison.OrdinalIgnoreCase) OrElse nxt.Name.Equals("blockquote", StringComparison.OrdinalIgnoreCase) Then
                            range.InsertParagraphAfter()
                            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                        End If
                    End If

                    Exit Select

                Case "#text", "strong", "b", "em", "i", "u", "del", "s",
                    "sub", "sup", "code", "span", "img", "br"
                    RenderInline(childNode, range, Nothing, String.Empty)

                Case "div"
                    Dim cls As String = childNode.GetAttributeValue("class", String.Empty)
                    If cls.Equals("footnotes", StringComparison.OrdinalIgnoreCase) Then
                        ' Instead of skipping: parse the OL inside the footnotes container.
                        Dim footOl As HtmlAgilityPack.HtmlNode = childNode.SelectSingleNode("ol")
                        If footOl IsNot Nothing Then
                            ' currentLevel is preserved so numbering remains consistent.
                            ParseHtmlNode(footOl, range, currentLevel)
                        End If
                        Exit Select
                    End If

                Case "br"
                    RenderInline(childNode, range, Nothing, String.Empty)

                Case "h1", "h2", "h3", "h4", "h5", "h6"
                    ' 1) Determine built-in heading style and its relative size factor.
                    Dim style As WdBuiltinStyle = WdBuiltinStyle.wdStyleNormal ' Default for 'p'
                    Dim headingSizeFactor As Single = 1.0F
                    Select Case childNode.Name.ToLower()
                        Case "h1" : style = WdBuiltinStyle.wdStyleHeading1 : headingSizeFactor = 1.5F
                        Case "h2" : style = WdBuiltinStyle.wdStyleHeading2 : headingSizeFactor = 1.3F
                        Case "h3" : style = WdBuiltinStyle.wdStyleHeading3 : headingSizeFactor = 1.15F
                        Case "h4" : style = WdBuiltinStyle.wdStyleHeading4 : headingSizeFactor = 1.1F
                        Case "h5" : style = WdBuiltinStyle.wdStyleHeading5 : headingSizeFactor = 1.05F
                        Case "h6" : style = WdBuiltinStyle.wdStyleHeading6 : headingSizeFactor = 1.0F
                    End Select

                    Dim txt As String = HtmlEntity.DeEntitize(childNode.InnerText)
                    Dim href As String = nestedHref

                    ' 2) Insert new paragraph and move range to it.
                    range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                    range.InsertParagraphAfter()
                    range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                    ' 3) Insert text.
                    Dim paraStart As Integer = range.Start
                    range.InsertAfter(txt)
                    range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                    ' 4) Determine paragraph range.
                    Dim paraRg As Word.Range = range.Duplicate
                    paraRg.Start = paraStart
                    paraRg.End = range.End
                    ' 5) Apply heading style, then scale the font size proportionally to the
                    '    document's own body (Normal) size so headings are not oversized by the
                    '    template's absolute built-in heading sizes.
                    paraRg.Style = style
                    Try
                        Dim baseSize As Single = 0.0F
                        Dim normalStyle As Word.Style =
                            CType(paraRg.Document.Styles(Word.WdBuiltinStyle.wdStyleNormal), Word.Style)
                        If normalStyle IsNot Nothing Then baseSize = normalStyle.Font.Size

                        If baseSize > 0.0F AndAlso baseSize <> CSng(Word.WdConstants.wdUndefined) Then
                            paraRg.Font.Size = CSng(System.Math.Round(baseSize * headingSizeFactor, MidpointRounding.AwayFromZero))
                        End If
                    Catch ex As System.Exception
                        Debug.WriteLine($"Heading size scaling skipped: {ex.Message}")
                    End Try

                    ' 6) Add hyperlink when present.
                    If href <> String.Empty Then
                        Dim hl As Word.Hyperlink =
                                            range.Hyperlinks.Add(
                                                Anchor:=paraRg,
                                                Address:=CStr(href)
                                            )
                        hl.Range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                        range.SetRange(hl.Range.End, hl.Range.End)
                    End If

                    ' 7) Paragraph break at the end.
                    range.InsertParagraphAfter()
                    range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                Case "a"
                    Dim cls As String = childNode.GetAttributeValue("class", String.Empty)
                    Dim href As String = childNode.GetAttributeValue("href", String.Empty)
                    Dim id As String = childNode.GetAttributeValue("id", String.Empty)

                    ' 1) Inline footnote reference inside the text?
                    If id.StartsWith("fnref:", StringComparison.OrdinalIgnoreCase) _
                       AndAlso href.StartsWith("#fn:", StringComparison.OrdinalIgnoreCase) Then

                        Debug.WriteLine("Setting Bookmark in Case 'a'")

                        ' Extract footnote number.
                        Dim fnNum As String = id.Substring(id.IndexOf(":"c) + 1)  ' e.g. "1"
                        Dim bookmarkName As String = "fn" & fnNum                         ' "fn1"

                        ' Display text (e.g. <sup>1</sup>).
                        Dim displayText As String = HtmlEntity.DeEntitize(childNode.InnerText)

                        'Hyperlink to bookmark.
                        range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                        range.Hyperlinks.Add(
                                                Anchor:=range,
                                                Address:="",                ' No external URL.
                                                SubAddress:=CStr(bookmarkName),   ' Internal target.
                                                TextToDisplay:=CStr(displayText)
                                                )
                        range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                        Exit Select    ' End only this Case block.
                    End If

                    ' 2) Back-reference arrow in the footnote list.
                    If cls.Equals("footnote-back-ref", StringComparison.OrdinalIgnoreCase) Then
                        RenderInline(childNode, range, Nothing, String.Empty)
                        Exit Select
                    End If

                    ' 3) All other anchors render normally.
                    RenderInline(childNode, range, Nothing, String.Empty)
                    Exit Select

                Case "ul"
                    ' a) Task list (checkboxes) handling as before …
                    If childNode.GetAttributeValue("class", "").Contains("contains-task-list") Then
                        For Each li As HtmlAgilityPack.HtmlNode In childNode.SelectNodes("li")
                            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                            Dim chk = li.SelectSingleNode(".//input[@type='checkbox']")
                            Dim symbol = If(chk IsNot Nothing _
                           AndAlso chk.GetAttributeValue("checked", False), "☑", "☐")
                            range.InsertAfter(symbol & " " &
                              HtmlEntity.DeEntitize(li.InnerText.Trim()) &
                              vbCr)
                            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                        Next
                        Exit Select
                    End If

                    Dim enteringTopUL = (currentLevel = 0)
                    If enteringTopUL AndAlso ulLevels Is Nothing Then
                        ulLevels = New List(Of Integer)()
                        ulStartPos = range.Start
                        Debug.WriteLine("[ul] Entering top UL – reset ulLevels")
                    End If

                    ' c) Original insertion of LI nodes (recursive).
                    Dim liNodes = childNode.SelectNodes("li")
                    If liNodes IsNot Nothing Then
                        Dim listStart = range.Start
                        For Each liNode As HtmlAgilityPack.HtmlNode In liNodes
                            ParseHtmlNode(liNode, range, currentLevel + 1)
                            range.InsertAfter(vbCr)
                            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                        Next
                        Dim ulRange As Word.Range = range.Duplicate
                        ulRange.Start = listStart   ' listStart must be in the same story as `range`
                        ulRange.End = range.End
                        ulRange.ListFormat.ApplyBulletDefault()
                        ulRange.ListFormat.ListIndent()
                        With ulRange.ParagraphFormat
                            .LeftIndent = .Application.CentimetersToPoints(0.75)
                            .FirstLineIndent = - .Application.CentimetersToPoints(0.75)
                        End With
                        range.SetRange(ulRange.End, ulRange.End)

                        ' d) At the end of the first UL: apply the stored levels.
                        If enteringTopUL Then
                            Debug.WriteLine($"[ul] At end of top UL – ulLevels.Count = {ulLevels.Count}")
                            Debug.WriteLine($"[ul] Levels array: {String.Join(",", ulLevels)}")

                            Dim paras = ulRange.Paragraphs
                            Dim maxItems = System.Math.Min(paras.Count, ulLevels.Count)
                            For i = 1 To maxItems
                                Dim p = paras(i)
                                Dim lvl = ulLevels(i - 1)
                                Debug.WriteLine($"[ul] Paragraph {i} initial level={lvl}")

                                ' Apply ListIndent() once per nested level beyond the first.
                                For stepIndent = 1 To (lvl - 1)
                                    p.Range.ListFormat.ListIndent()
                                Next
                            Next
                            ulLevels = Nothing
                        End If
                    End If
                    If enteringTopUL Then
                        ulLevels = Nothing
                    End If

                    Exit Select

                Case "ol"

                    ' a) Retrieve all <li> nodes.
                    Dim liNodes As HtmlAgilityPack.HtmlNodeCollection = childNode.SelectNodes("li")
                    If liNodes Is Nothing OrElse liNodes.Count = 0 Then
                        Debug.WriteLine("[ol] No <li> nodes found in <ol> – skipping")
                        Exit Select
                    End If

                    ' b) Read start attribute (start number).
                    Dim startAttr As Integer = 1
                    Dim startStr As String = childNode.GetAttributeValue("start", String.Empty)
                    Dim tmpInt As Integer
                    If Integer.TryParse(startStr, tmpInt) Then startAttr = tmpInt

                    ' c) Top-level OL: initialize ulLevels (shared with UL handling).
                    Dim enteringTopUL As Boolean = (currentLevel = 0)
                    If enteringTopUL AndAlso ulLevels Is Nothing Then
                        ulLevels = New List(Of Integer)()
                        ulStartPos = range.Start
                        Debug.WriteLine("[ol] Entering top OL – reset ulLevels")
                    End If

                    ' d) Render each LI recursively (LI logic manages CR and level push).
                    Dim listStart As Integer = range.Start
                    For Each liNode As HtmlAgilityPack.HtmlNode In liNodes
                        ParseHtmlNode(liNode, range, currentLevel + 1)
                        range.InsertAfter(vbCr)
                        range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                    Next

                    ' e) Capture the entire OL range.
                    Dim olRange As Word.Range = range.Duplicate
                    olRange.Start = listStart
                    olRange.End = range.End

                    ' f) Ensure range contains paragraphs.
                    If olRange.Paragraphs.Count = 0 Then
                        Debug.WriteLine("[ol] olRange contains no paragraphs – skip numbering")
                        Exit Select
                    End If

                    ' j) Apply formatting only for the top-most OL.
                    If enteringTopUL Then
                        Dim paras As Word.Paragraphs = olRange.Paragraphs

                        ' Remove previous numbering.
                        olRange.ListFormat.RemoveNumbers()

                        ' Use a multi-level template from ListGalleries.
                        Dim multiLevelTemplate As Word.ListTemplate = olRange.Application.ListGalleries(Word.WdListGalleryType.wdOutlineNumberGallery).ListTemplates(1)

                        ' Set custom start value for level 1 if specified.
                        If startAttr <> 1 Then
                            multiLevelTemplate.ListLevels(1).StartAt = startAttr
                        End If

                        ' Apply the multi-level template.
                        olRange.ListFormat.ApplyListTemplateWithLevel(
                                ListTemplate:=multiLevelTemplate,
                                ContinuePreviousList:=False,
                                ApplyTo:=Word.WdListApplyTo.wdListApplyToSelection,
                                DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior
                            )

                        ' Assign correct level per paragraph.
                        For i As Integer = 1 To paras.Count
                            Dim p As Word.Paragraph = paras(i)
                            Dim lvl As Integer = ulLevels(i - 1)
                            If lvl >= 1 AndAlso lvl <= multiLevelTemplate.ListLevels.Count Then
                                p.Range.ListFormat.ListLevelNumber = lvl
                            End If
                        Next
                        ulLevels = Nothing
                    End If

                    ' j) Place cursor after the list.
                    range.SetRange(olRange.End, olRange.End)

                    Exit Select

                Case "dl"
                    ' Definition list.
                    For Each dt As HtmlNode In childNode.SelectNodes("dt")
                        ' Term.
                        Dim term As Microsoft.Office.Interop.Word.Range = range.Duplicate
                        term.Text = HtmlEntity.DeEntitize(dt.InnerText) & vbTab
                        term.Font.Bold = True
                        term.Collapse(False)
                        range.SetRange(term.End, term.End)
                        ' Definition.
                        Dim dd As HtmlNode = dt.NextSibling
                        If dd IsNot Nothing AndAlso dd.Name.ToLower() = "dd" Then
                            Dim defn As Microsoft.Office.Interop.Word.Range = range.Duplicate
                            defn.Text = HtmlEntity.DeEntitize(dd.InnerText) & vbCr
                            defn.ParagraphFormat.LeftIndent += 18
                            defn.Collapse(False)
                            range.SetRange(defn.End, defn.End)
                        End If
                    Next

                Case "hr"

                    range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                    Dim hrPara As Word.Paragraph = range.Paragraphs.Add(range)
                    hrPara.Range.Text = ""  ' Keep empty; only the border is needed.

                    With hrPara.Range.ParagraphFormat.Borders(Word.WdBorderType.wdBorderBottom)
                        .LineStyle = Word.WdLineStyle.wdLineStyleSingle
                        .LineWidth = Word.WdLineWidth.wdLineWidth050pt
                        .Color = Word.WdColor.wdColorAutomatic
                    End With

                    Dim afterHr As Word.Range = hrPara.Range.Duplicate
                    afterHr.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                    range.SetRange(afterHr.Start, afterHr.Start)
                    range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                Case "input"
                    ' Checkbox (ContentControl).
                    If childNode.GetAttributeValue("type", String.Empty).ToLower() = "checkbox" Then
                        range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                        Dim cc As Word.ContentControl =
                                        range.ContentControls.Add(
                                            Word.WdContentControlType.wdContentControlCheckBox,
                                            range
                                        )
                        cc.Checked = (childNode.GetAttributeValue("checked", String.Empty).ToLower() = "checked")

                        range.SetRange(cc.Range.End, cc.Range.End)
                        range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                    End If

                Case "img"

                    RenderInline(childNode, range, Nothing, String.Empty)

                Case "pre"
                    ' Code block.
                    Dim codeBlock As Microsoft.Office.Interop.Word.Range = range.Duplicate
                    codeBlock.Text = HtmlEntity.DeEntitize(childNode.InnerText) & vbCr
                    codeBlock.Font.Name = "Courier New"
                    codeBlock.Font.Size = 10
                    codeBlock.ParagraphFormat.LeftIndent += 14.18
                    codeBlock.Collapse(False)
                    range.SetRange(codeBlock.End, codeBlock.End)

                Case "table"
                    '---------- 1) Retrieve top-level rows ----------------------------
                    Dim topRows As New List(Of HtmlNode)

                    'Direct <tr> plus <thead>/<tbody> children (non-recursive).
                    For Each tr As HtmlNode In childNode.SelectNodes("./tr|./thead/tr|./tbody/tr")
                        topRows.Add(tr)
                    Next
                    If topRows.Count = 0 Then Exit Select
                    '----------------------------------------------------------------

                    '---------- 2) Determine best column count ----------------------
                    Dim colCount As Integer = 0
                    For Each tr In topRows
                        Dim cells = tr.SelectNodes("th|td")
                        If cells IsNot Nothing AndAlso cells.Count > colCount Then
                            colCount = cells.Count
                        End If
                    Next
                    If colCount = 0 Then Exit Select
                    '----------------------------------------------------------------

                    '---------- 3) Create table at cursor ----------------------------
                    Dim tbl As Microsoft.Office.Interop.Word.Table =
                        range.Tables.Add(range, topRows.Count, colCount)

                    '---------- 4) Populate cells -----------------------------------
                    Dim rIdx As Integer = 1
                    For Each tr In topRows
                        Dim cells = tr.SelectNodes("th|td")
                        Dim cIdx As Integer = 1

                        If cells IsNot Nothing Then
                            For Each cell In cells
                                Dim cellRg As Word.Range = tbl.Cell(rIdx, cIdx).Range
                                'Remove the hidden cell-end character.
                                cellRg.SetRange(cellRg.Start, cellRg.End - 1)

                                ParseHtmlNode(cell, cellRg, currentLevel)          '← recursive, data preserved.

                                'Header cells bold.
                                If cell.Name.Equals("th", StringComparison.OrdinalIgnoreCase) Then
                                    cellRg.Font.Bold = True
                                End If

                                '---------- 4a) Handle colspan ------------------
                                Dim cSpan As Integer = cell.GetAttributeValue("colspan", 1)
                                If cSpan > 1 AndAlso cIdx + cSpan - 1 <= colCount Then
                                    Dim tgtCell = tbl.Cell(rIdx, cIdx + cSpan - 1)
                                    tbl.Cell(rIdx, cIdx).Merge(tgtCell)
                                    cIdx += cSpan                    'Continue after the merge.
                                Else
                                    cIdx += 1
                                End If
                                '----------------------------------------------------
                            Next
                        End If
                        rIdx += 1
                    Next

                    '---------- 5) Place cursor after table -------------------------
                    range.SetRange(tbl.Range.End, tbl.Range.End)

                Case Else

                    ParseHtmlNode(childNode, range, currentLevel)

            End Select

        Next
    End Sub

    ''' <summary>
    ''' Inserts text with optional emoji-aware splits, applying a base style and hyperlink when required.
    ''' </summary>
    Private Shared Sub InsertInline(
        ByRef mainRg As Range,
        txt As String,
        baseStyle As Action(Of Range),
        Optional href As String = "")

        ' 1) If no emoji handling is needed, use the direct path.
        If emojiSet Is Nothing OrElse emojiSet.Count = 0 OrElse Not _emojiPairRegex.IsMatch(txt) Then
            TrueInsertInline(mainRg, txt, baseStyle, href)
            Return
        End If

        ' 2) Otherwise split at emoji boundaries and apply styling conditionally.
        Dim lastPos As Integer = 0
        For Each m As Match In _emojiPairRegex.Matches(txt)
            ' (a) Everything before the emoji.
            If m.Index > lastPos Then
                Dim segment As String = txt.Substring(lastPos, m.Index - lastPos)
                TrueInsertInline(mainRg, segment, baseStyle, href)
            End If

            ' (b) The emoji itself (only when it exists in the set).
            Dim emoji As String = m.Value
            If emojiSet.Contains(emoji) Then
                Dim emojiStyle = CombineStyle(baseStyle,
                    Sub(r As Range) r.Font.Name = "Segoe UI Emoji")
                TrueInsertInline(mainRg, emoji, emojiStyle, href)
            Else
                ' If the emoji is not in the set, treat it as normal text.
                TrueInsertInline(mainRg, emoji, baseStyle, href)
            End If

            lastPos = m.Index + m.Length
        Next

        ' (c) Remaining text after the last emoji.
        If lastPos < txt.Length Then
            Dim tail As String = txt.Substring(lastPos)
            TrueInsertInline(mainRg, tail, baseStyle, href)
        End If
    End Sub

    ''' <summary>
    ''' Inserts text into the provided range, optionally creating a hyperlink and applying a style delegate.
    ''' </summary>
    Private Shared Sub TrueInsertInline(
        ByRef mainRg As Word.Range,
        txt As String,
        styleAction As Action(Of Word.Range),
        Optional href As String = "")

        mainRg.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

        Dim wrk As Word.Range = mainRg.Duplicate
        wrk.Text = txt

        ' **Reset occurs here** – remember to log before and after the reset when required.
        wrk.Font.Reset()

        If href <> "" Then
            Dim hl = mainRg.Document.Hyperlinks.Add(Anchor:=wrk, Address:=CStr(href))
            If styleAction IsNot Nothing Then
                Debug.WriteLine("[InsertInline] → Applying styleAction to hyperlink range")
                styleAction(hl.Range)
            End If
            hl.Range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
            mainRg.SetRange(hl.Range.End, hl.Range.End)
        Else
            If styleAction IsNot Nothing Then
                Debug.WriteLine("[InsertInline] → Applying styleAction to text range")
                styleAction(wrk)
            End If
            wrk.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
            mainRg.SetRange(wrk.End, wrk.End)
        End If

    End Sub

End Class
