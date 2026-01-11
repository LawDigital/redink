' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland (and Gustavo Hennig, as a licensor for the original MarkdowntoRTF code). All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: MarkdownToRtfConverter.vb
' Purpose: Converts a Markdown string to an RTF-formatted string.
'
' Architecture:
'  - Markdown Parsing: Uses Markdig with advanced extensions, pipe/grid tables, footnotes, and emoji.
'  - Preprocessing: Optionally escapes asterisks inside square brackets to prevent unintended emphasis,
'    and escapes Excel-like instruction markers at line start to prevent link parsing.
'  - Footnote Indexing: Collects footnote definitions (by label) from FootnoteGroup blocks into a lookup.
'  - RTF Construction: Emits a single RTF header (codepage 1252, font table, \uc1) and appends RTF for
'    supported Markdown block types (headings, paragraphs, lists, quotes, tables, code blocks, thematic breaks).
'  - Inline Rendering: Renders common Markdown inline constructs (literal text, emphasis, code spans,
'    line breaks, links/images, limited HTML tags, emoji, and footnote links).
'
' Notes / Known Issues:
'  - RTF output for tables uses tab-separated cells and does not emit true RTF table constructs.
'
' Licensing/Origin:
'  - The Markdown-to-RTF conversion code is adapted from the open-source project "MarkdownToRtf".
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Text
Imports Markdig
Imports Markdig.Extensions.Footnotes
Imports Markdig.Syntax

Namespace SharedLibrary

    ''' <summary>
    ''' Markdown-to-RTF conversion utilities.
    ''' </summary>
    Public Module MarkdownToRtfConverter

        ''' <summary>
        ''' Converts Markdown markup to an RTF-formatted string.
        ''' </summary>
        ''' <param name="markdownText">The input string containing Markdown markup.</param>
        ''' <param name="preserveSquareBracketLiterals">
        ''' If <see langword="True"/>, escapes asterisks inside <c>[...]</c> ranges so Markdown emphasis does not apply.
        ''' </param>
        ''' <returns>An RTF-formatted string.</returns>
        Public Function Convert(markdownText As String,
                        Optional preserveSquareBracketLiterals As Boolean = False) As String

            ' Optionally prevent markdown emphasis inside [...] (e.g., formulas like [1*2*3])
            If preserveSquareBracketLiterals AndAlso Not String.IsNullOrEmpty(markdownText) Then
                markdownText = EscapeAsterisksInsideSquareBrackets(markdownText)
            End If

            ' Escape "[Cell:" / "[Value:" / "[Formula:" / "[Comment:" at the start of a line so Markdig will not parse them as links.
            markdownText = EscapeExcelInstructionMarkers(markdownText)

            'markdownText = System.Text.RegularExpressions.Regex.Unescape(markdownText)
            markdownText = System.Text.RegularExpressions.Regex.Replace(
                markdownText,
                "^[ \t]+(?=>)",       ' "any sequence of spaces/tabs directly before a >"
                String.Empty,
                System.Text.RegularExpressions.RegexOptions.Multiline)

            Debug.WriteLine("MarkdownToRtfConverter.Convert: " & markdownText)

            ' 1) Parse Markdown
            'Dim pipeline = New Markdig.MarkdownPipelineBuilder().Build()
            Dim pipeline = New Markdig.MarkdownPipelineBuilder() _
                .UseAdvancedExtensions() _
                .UsePipeTables() _
                .UseGridTables() _
                .UseFootnotes() _
                .UseEmojiAndSmiley() _
                .Build()
            Dim document = Markdig.Markdown.Parse(markdownText, pipeline)

            ' Collect footnote definitions by label.
            Dim fnDefs As New Dictionary(Of String, Markdig.Extensions.Footnotes.Footnote)()
            For Each block In document
                If TypeOf block Is FootnoteGroup Then
                    ' Skip the group; the actual Footnote blocks are contained within it.
                    For Each fn As Markdig.Extensions.Footnotes.Footnote In CType(block, FootnoteGroup)
                        fnDefs(fn.Label) = fn
                    Next
                End If
            Next

            ' 2) Build RTF
            Dim rtfBuilder As New System.Text.StringBuilder()
            ' (1) A *single* RTF header with codepage, font table and \uc1
            rtfBuilder.AppendLine("{\rtf1\ansi\ansicpg1252\deff0")
            rtfBuilder.AppendLine("{\fonttbl{\f0\fnil\fcharset0 Arial;}{\f1\fmodern\fcharset0 Courier New;}}")
            ' \uc1 for consistent Unicode fallback rendering (\uN?)
            rtfBuilder.AppendLine("\uc1")

            ' 3) Process blocks
            For Each block In document
                If TypeOf block Is Markdig.Extensions.Tables.Table Then
                    ConvertTableBlock(rtfBuilder, CType(block, Markdig.Extensions.Tables.Table), fnDefs)
                ElseIf TypeOf block Is Markdig.Syntax.HeadingBlock Then
                    ConvertHeadingBlock(rtfBuilder, CType(block, Markdig.Syntax.HeadingBlock), fnDefs)
                ElseIf TypeOf block Is Markdig.Syntax.ParagraphBlock Then
                    ConvertParagraphBlock(rtfBuilder, CType(block, Markdig.Syntax.ParagraphBlock), fnDefs)
                ElseIf TypeOf block Is Markdig.Syntax.ListBlock Then
                    ConvertListBlock(rtfBuilder, CType(block, Markdig.Syntax.ListBlock), 0, fnDefs)
                ElseIf TypeOf block Is Markdig.Syntax.QuoteBlock Then
                    ConvertQuoteBlock(rtfBuilder, CType(block, Markdig.Syntax.QuoteBlock), 1, fnDefs)
                ElseIf TypeOf block Is Markdig.Syntax.FencedCodeBlock Then
                    ConvertCodeBlock(rtfBuilder, CType(block, Markdig.Syntax.FencedCodeBlock), fnDefs)
                    ' (2) Convert generic (e.g., indented) code blocks as well.
                ElseIf (TypeOf block Is Markdig.Syntax.CodeBlock) AndAlso Not (TypeOf block Is Markdig.Syntax.FencedCodeBlock) Then
                    ConvertCodeBlock(rtfBuilder, CType(block, Markdig.Syntax.CodeBlock))
                ElseIf TypeOf block Is Markdig.Syntax.ThematicBreakBlock Then
                    ConvertThematicBreakBlock(rtfBuilder)
                ElseIf TypeOf block Is FootnoteGroup Then
                    ' FootnoteGroup blocks are indexed earlier and are not rendered as top-level blocks here.
                End If
            Next

            ' Close the RTF document.
            rtfBuilder.AppendLine("}")
            Return rtfBuilder.ToString()
        End Function

        ''' <summary>
        ''' Escapes the opening bracket of Excel instruction markers at the start of a line
        ''' so Markdig will not parse them as links (the backslash is consumed by the parser).
        ''' </summary>
        ''' <param name="md">Markdown input.</param>
        ''' <returns>Markdown with escaped instruction markers.</returns>
        Private Function EscapeExcelInstructionMarkers(md As String) As String
            If String.IsNullOrEmpty(md) Then Return md
            ' Match start-of-line optional whitespace, then [Cell:|[Value:|[Formula:|[Comment:
            Dim pattern As String = "(?m)^(\s*)\[(Cell|Value|Formula|Comment):"
            Dim replacement As String = "$1\[$2:"
            Return System.Text.RegularExpressions.Regex.Replace(md, pattern, replacement)
        End Function

        ''' <summary>
        ''' Escapes asterisks inside square-bracket ranges so Markdown will not turn <c>*x*</c> into italics.
        ''' </summary>
        ''' <param name="input">Text to process.</param>
        ''' <returns>
        ''' Text where asterisks inside <c>[...]</c> are escaped. Example: <c>[1*2*3]</c> becomes <c>[1\*2\*3]</c>.
        ''' The backslashes are consumed by the Markdown parser; final rendered text remains <c>[1*2*3]</c>.
        ''' </returns>
        Private Function EscapeAsterisksInsideSquareBrackets(input As String) As String
            If String.IsNullOrEmpty(input) Then Return input

            Dim sb As New System.Text.StringBuilder(input.Length)
            Dim bracketDepth As Integer = 0

            Dim inInlineCode As Boolean = False
            Dim inlineTicks As Integer = 0

            Dim inFencedCode As Boolean = False
            Dim fencedTicks As Integer = 0

            Dim atLineStart As Boolean = True
            Dim i As Integer = 0

            While i < input.Length
                Dim ch As Char = input(i)

                ' Handle runs of backticks (inline spans and fenced blocks).
                If ch = "`"c Then
                    Dim start As Integer = i
                    While i < input.Length AndAlso input(i) = "`"c
                        i += 1
                    End While
                    Dim count As Integer = i - start
                    sb.Append(New String("`"c, count))

                    If inInlineCode Then
                        If count = inlineTicks Then
                            inInlineCode = False
                            inlineTicks = 0
                        End If
                    ElseIf inFencedCode Then
                        ' Close fence only at line start and with >= opening length.
                        If atLineStart AndAlso count >= fencedTicks Then
                            inFencedCode = False
                            fencedTicks = 0
                        End If
                    Else
                        If atLineStart AndAlso count >= 3 Then
                            inFencedCode = True
                            fencedTicks = count
                        Else
                            inInlineCode = True
                            inlineTicks = count
                        End If
                    End If

                    atLineStart = False
                    Continue While
                End If

                ' Track newlines for "line start" detection (for fenced blocks).
                If ch = vbCr OrElse ch = vbLf Then
                    sb.Append(ch)
                    atLineStart = True
                    i += 1
                    Continue While
                End If

                ' Inside code: pass through verbatim and do not touch bracketDepth.
                If inInlineCode OrElse inFencedCode Then
                    sb.Append(ch)
                    atLineStart = False
                    i += 1
                    Continue While
                End If

                ' Outside code: manage bracket depth and escape '*' inside [...].
                Select Case ch
                    Case "["c
                        bracketDepth += 1
                        sb.Append(ch)
                    Case "]"c
                        If bracketDepth > 0 Then bracketDepth -= 1
                        sb.Append(ch)
                    Case "*"c
                        If bracketDepth > 0 Then
                            ' Idempotent: avoid double-escaping if the previous character is already '\'.
                            If sb.Length = 0 OrElse sb(sb.Length - 1) <> "\"c Then
                                sb.Append("\"c)
                            End If
                            sb.Append("*"c)
                        Else
                            sb.Append("*"c)
                        End If
                    Case Else
                        sb.Append(ch)
                End Select

                atLineStart = False
                i += 1
            End While

            Return sb.ToString()
        End Function



        ''' <summary>
        ''' Appends an RTF horizontal rule representation.
        ''' </summary>
        ''' <param name="rtf">Target RTF builder.</param>
        Private Sub ConvertThematicBreakBlock(rtf As System.Text.StringBuilder)
            ' New paragraph + horizontal rule + new paragraph.
            rtf.AppendLine("\par")
            rtf.AppendLine("\pard\brdrb\brdrs\brdrw10\par")
        End Sub



        ''' <summary>
        ''' Appends all lines of a Markdig code line group as escaped RTF, using <c>\line</c> separators.
        ''' </summary>
        ''' <param name="rtf">Target RTF builder.</param>
        ''' <param name="linesGroup">Markdig line group to render.</param>
        Private Sub AppendCodeLines(rtf As System.Text.StringBuilder,
                                linesGroup As Markdig.Helpers.StringLineGroup)
            Dim arr = linesGroup.Lines
            If arr Is Nothing OrElse linesGroup.Count = 0 Then
                ' Nothing to output (paragraph framing is emitted by caller).
                Exit Sub
            End If
            For i = 0 To linesGroup.Count - 1
                Dim slice = arr(i).Slice
                If slice.Text Is Nothing Then
                    rtf.Append("\line ")
                    Continue For
                End If
                Dim raw As String = slice.Text.Substring(slice.Start, slice.Length)
                rtf.Append(EscapeRtf(raw)).Append("\line ")
            Next
        End Sub

        ''' <summary>
        ''' Converts a fenced Markdown code block to RTF using a monospace font.
        ''' </summary>
        ''' <param name="rtf">Target RTF builder.</param>
        ''' <param name="codeBlock">Fenced code block to convert.</param>
        ''' <param name="fnDefs">Footnote definitions lookup (not used by this overload).</param>
        Private Sub ConvertCodeBlock(
        rtf As System.Text.StringBuilder,
        codeBlock As Markdig.Syntax.FencedCodeBlock,
        fnDefs As System.Collections.Generic.Dictionary(Of String, Markdig.Extensions.Footnotes.Footnote)
    )
            If codeBlock Is Nothing Then Return
            rtf.Append("\par\f1\fs18 ")
            AppendCodeLines(rtf, codeBlock.Lines)
            rtf.Append("\f0\fs20\par")
        End Sub

        ''' <summary>
        ''' Converts a generic (indented) Markdown code block to RTF using a monospace font.
        ''' </summary>
        ''' <param name="rtf">Target RTF builder.</param>
        ''' <param name="codeBlock">Code block to convert.</param>
        Private Sub ConvertCodeBlock(
        rtf As System.Text.StringBuilder,
        codeBlock As Markdig.Syntax.CodeBlock
    )
            If codeBlock Is Nothing Then Return
            rtf.Append("\par\f1\fs18 ")
            AppendCodeLines(rtf, codeBlock.Lines)
            rtf.Append("\f0\fs20\par")
        End Sub




        ''' <summary>
        ''' Converts a Markdig table block to RTF using tab-separated cell content.
        ''' </summary>
        ''' <param name="rtf">Target RTF builder.</param>
        ''' <param name="table">Table to convert.</param>
        ''' <param name="fnDefs">Footnote definitions lookup used by inline rendering.</param>
        Private Sub ConvertTableBlock(
rtf As StringBuilder,
table As Markdig.Extensions.Tables.Table,
fnDefs As Dictionary(Of String, Markdig.Extensions.Footnotes.Footnote)
)
            ' Normalize to equal-width rows (ensures consistent cell counts).
            table.NormalizeUsingMaxWidth()

            For Each row As Markdig.Extensions.Tables.TableRow In table
                rtf.Append("\pard\sa100\fs20 ")

                For Each cell As Markdig.Extensions.Tables.TableCell In row
                    ' Render all blocks contained in this cell.
                    For Each subBlock As Markdig.Syntax.Block In cell
                        Select Case True
                            Case TypeOf subBlock Is Markdig.Syntax.ParagraphBlock
                                Dim p As Markdig.Syntax.ParagraphBlock =
                        CType(subBlock, Markdig.Syntax.ParagraphBlock)
                                ConvertInline(rtf, p.Inline, fnDefs)

                            Case TypeOf subBlock Is Markdig.Syntax.ListBlock
                                ConvertListBlock(rtf:=rtf,
                                    listBlock:=CType(subBlock, Markdig.Syntax.ListBlock),
                                    level:=0,
                                    fnDefs:=fnDefs)

                            Case TypeOf subBlock Is Markdig.Syntax.CodeBlock
                                ConvertCodeBlock(rtf, CType(subBlock, Markdig.Syntax.CodeBlock))

                                ' Additional block types (e.g., QuoteBlock) are not handled here.
                        End Select
                    Next

                    ' Cell separator.
                    rtf.Append("\tab ")
                Next

                rtf.AppendLine("\par")
            Next
        End Sub




        ''' <summary>
        ''' Converts a Markdown heading block to RTF using a size based on heading level.
        ''' </summary>
        ''' <param name="rtf">Target RTF builder.</param>
        ''' <param name="headingBlock">Heading block to convert.</param>
        ''' <param name="fnDefs">Footnote definitions lookup used by inline rendering.</param>
        Private Sub ConvertHeadingBlock(rtf As System.Text.StringBuilder, headingBlock As Markdig.Syntax.HeadingBlock, fnDefs As Dictionary(Of String, Markdig.Extensions.Footnotes.Footnote))
            Dim headingSizes() As Integer = {30, 28, 26, 24, 22, 20}
            Dim level As Integer = headingBlock.Level
            Dim size As Integer = headingSizes(System.Math.Min(level, headingSizes.Length) - 1)

            rtf.Append($"\pard\sa180\fs{size} \b ")
            ConvertInline(rtf, headingBlock.Inline, fnDefs)
            rtf.AppendLine(" \b0\par")
        End Sub

        ''' <summary>
        ''' Converts a Markdown paragraph block to an RTF paragraph.
        ''' </summary>
        ''' <param name="rtf">Target RTF builder.</param>
        ''' <param name="paragraphBlock">Paragraph block to convert.</param>
        ''' <param name="fnDefs">Footnote definitions lookup used by inline rendering.</param>
        Private Sub ConvertParagraphBlock(rtf As System.Text.StringBuilder, paragraphBlock As Markdig.Syntax.ParagraphBlock, fnDefs As Dictionary(Of String, Markdig.Extensions.Footnotes.Footnote))
            rtf.Append("\pard\sa180\fs20 ")
            ConvertInline(rtf, paragraphBlock.Inline, fnDefs)
            rtf.AppendLine("\par")
        End Sub


        ''' <summary>
        ''' Converts a Markdown list block (ordered or unordered) to RTF, supporting nesting.
        ''' </summary>
        ''' <param name="rtf">Target RTF builder.</param>
        ''' <param name="listBlock">List block to convert.</param>
        ''' <param name="level">Nesting level used to calculate indentation.</param>
        ''' <param name="fnDefs">Footnote definitions lookup used by inline rendering.</param>
        Private Sub ConvertListBlock(rtf As System.Text.StringBuilder,
                         listBlock As Markdig.Syntax.ListBlock,
                         Optional level As Integer = 0,
                                 Optional fnDefs As Dictionary(Of String, Markdig.Extensions.Footnotes.Footnote) = Nothing)

            Dim isOrdered As Boolean = listBlock.IsOrdered
            Dim indent As Integer = level * 360            ' 360 twips ≈ 0.25"
            Dim itemIndex As Integer = 0

            ' Determine the start value for ordered lists.
            Dim startNumber As Integer = 1
            If isOrdered Then
                For Each blk In listBlock
                    If TypeOf blk Is Markdig.Syntax.ListItemBlock Then
                        Dim firstLi = CType(blk, Markdig.Syntax.ListItemBlock)
                        If firstLi.Order <> 0 Then startNumber = firstLi.Order
                        Exit For
                    End If
                Next
            End If

            For Each item In listBlock
                If TypeOf item Is Markdig.Syntax.ListItemBlock Then
                    Dim li = CType(item, Markdig.Syntax.ListItemBlock)
                    itemIndex += 1

                    ' Bullet/number prefix plus a tab to align the content at the tab stop.
                    Dim prefix = If(isOrdered,
                       $"{startNumber + itemIndex - 1}. ",
                       "\u8226?\tab ")    ' Trailing space is part of the prefix.

                    ' Single \pard configuring left indent, hanging indent and tab stop.
                    rtf.Append($"\pard\li{indent}\fi-200\tx{indent + 200}\sa50\fs20 ")
                    rtf.Append(prefix)

                    ' Render all blocks within the list item.
                    For Each sb In li
                        Select Case True
                            Case TypeOf sb Is Markdig.Syntax.ParagraphBlock
                                ConvertInline(rtf, CType(sb, Markdig.Syntax.ParagraphBlock).Inline, fnDefs)

                            Case TypeOf sb Is Markdig.Syntax.ListBlock
                                rtf.AppendLine("\par")    ' Blank line before a nested list.
                                ConvertListBlock(rtf,
                                     CType(sb, Markdig.Syntax.ListBlock),
                                     level + 1, fnDefs)
                            Case TypeOf sb Is Markdig.Syntax.CodeBlock
                                rtf.AppendLine()
                                ConvertCodeBlock(rtf, CType(sb, Markdig.Syntax.CodeBlock))

                        End Select
                    Next

                    rtf.AppendLine("\par")               ' Finalize the item.
                End If
            Next
        End Sub

        ''' <summary>
        ''' Converts a Markdown QuoteBlock with indentation (supports nested quotes).
        ''' </summary>
        ''' <param name="rtf">Target RTF builder.</param>
        ''' <param name="quoteBlock">Quote block to convert.</param>
        ''' <param name="level">Nesting level used to calculate indentation.</param>
        ''' <param name="fnDefs">Footnote definitions lookup used by inline rendering.</param>
        Private Sub ConvertQuoteBlock(
rtf As System.Text.StringBuilder,
quoteBlock As Markdig.Syntax.QuoteBlock,
Optional level As Integer = 1,
Optional fnDefs As Dictionary(Of String, Markdig.Extensions.Footnotes.Footnote) = Nothing
)
            ' Left indentation per level: 360 twips ≈ 0.25 cm (as per existing code comment).
            Dim indentPerLevel As Integer = 360
            Dim indent As Integer = level * indentPerLevel

            ' \pard begins a new paragraph.
            rtf.Append($"\pard\li{indent}\sa180\fs20 ")

            ' Render each child block (typically ParagraphBlock) within the quote.
            For Each inner In quoteBlock
                If TypeOf inner Is Markdig.Syntax.ParagraphBlock Then
                    ConvertInline(rtf, CType(inner, Markdig.Syntax.ParagraphBlock).Inline, fnDefs)
                    rtf.AppendLine("\par")
                ElseIf TypeOf inner Is Markdig.Syntax.ListBlock Then
                    ' Nested list inside the quote.
                    ConvertListBlock(rtf, CType(inner, Markdig.Syntax.ListBlock), level, fnDefs)
                ElseIf TypeOf inner Is Markdig.Syntax.QuoteBlock Then
                    ' Nested quote: increase nesting level.
                    ConvertQuoteBlock(rtf, CType(inner, Markdig.Syntax.QuoteBlock), level + 1, fnDefs)
                End If
            Next

            ' Ensure the quote ends with a paragraph break.
            rtf.AppendLine("\par")
        End Sub


        ''' <summary>
        ''' Renders all inline elements within a Markdig <see cref="Markdig.Syntax.Inlines.ContainerInline"/> into RTF.
        ''' </summary>
        ''' <param name="rtf">Target RTF builder.</param>
        ''' <param name="container">Inline container to render.</param>
        ''' <param name="fnDefs">Footnote definitions lookup used when encountering footnote links.</param>
        ''' <param name="visitedFootnotes">Set used to prevent recursion when rendering footnotes.</param>
        Private Sub ConvertInline(
rtf As System.Text.StringBuilder,
container As Markdig.Syntax.Inlines.ContainerInline,
Optional fnDefs As System.Collections.Generic.Dictionary(Of String, Markdig.Extensions.Footnotes.Footnote) = Nothing,
Optional visitedFootnotes As System.Collections.Generic.HashSet(Of String) = Nothing
)
            ' Guard: some blocks (e.g., empty headings) may have Nothing for their Inline property.
            If container Is Nothing Then Return

            If visitedFootnotes Is Nothing Then
                visitedFootnotes = New System.Collections.Generic.HashSet(Of String)()
            End If

            For Each inline In container
                Select Case True

                    ' Literal text
                    Case TypeOf inline Is Markdig.Syntax.Inlines.LiteralInline
                        Dim lit = CType(inline, Markdig.Syntax.Inlines.LiteralInline)
                        rtf.Append(EscapeRtf(lit.Content.ToString()))

                    ' Emphasis (bold/italic/strikethrough/sub/superscript)
                    Case TypeOf inline Is Markdig.Syntax.Inlines.EmphasisInline
                        Dim emp = CType(inline, Markdig.Syntax.Inlines.EmphasisInline)
                        Select Case True
                            Case emp.DelimiterChar = "~"c AndAlso emp.DelimiterCount = 2
                                rtf.Append("\strike ")
                                ConvertInline(rtf, emp, fnDefs, visitedFootnotes)
                                rtf.Append("\strike0 ")
                            Case emp.DelimiterChar = "~"c AndAlso emp.DelimiterCount = 1
                                rtf.Append("{\sub ")
                                ConvertInline(rtf, emp, fnDefs, visitedFootnotes)
                                rtf.Append("\nosupersub} ")
                            Case emp.DelimiterChar = "^"c AndAlso emp.DelimiterCount = 1
                                rtf.Append("{\super ")
                                ConvertInline(rtf, emp, fnDefs, visitedFootnotes)
                                rtf.Append("\nosupersub} ")
                            Case Else
                                HandleEmphasis(rtf, emp)
                        End Select

                    ' Inline code span
                    Case TypeOf inline Is Markdig.Syntax.Inlines.CodeInline
                        Dim ci = CType(inline, Markdig.Syntax.Inlines.CodeInline)
                        rtf.Append("\f1 ")                               ' Monospace font
                        rtf.Append(EscapeRtf(ci.Content))
                        rtf.Append("\f0 ")                               ' Back to default font

                    ' Line break (hard or soft)
                    Case TypeOf inline Is Markdig.Syntax.Inlines.LineBreakInline
                        rtf.Append("\line ")

                    ' Link or image
                    Case TypeOf inline Is Markdig.Syntax.Inlines.LinkInline
                        Dim link = CType(inline, Markdig.Syntax.Inlines.LinkInline)
                        If link.IsImage Then
                            ' Image: render alt text only.
                            Dim alt As String = ""
                            If link.FirstChild IsNot Nothing AndAlso TypeOf link.FirstChild Is Markdig.Syntax.Inlines.LiteralInline Then
                                alt = CType(link.FirstChild, Markdig.Syntax.Inlines.LiteralInline).Content.ToString()
                            End If
                            rtf.Append("[Image: " & EscapeRtf(alt) & "] ")
                        Else
                            ' Hyperlink (RTF field).
                            If link.FirstChild Is Nothing Then
                                rtf.Append("{\field{\*\fldinst HYPERLINK """ & EscapeRtf(link.Url) & """}{\fldrslt " & EscapeRtf(link.Url) & "}}")
                            Else
                                rtf.Append("{\field{\*\fldinst HYPERLINK """ & EscapeRtf(link.Url) & """}{\fldrslt ")
                                ConvertInline(rtf, link, fnDefs, visitedFootnotes)
                                rtf.Append("}}")
                            End If
                        End If

                    ' HTML inline (<u>, <sup>, <sub>, otherwise escape as literal)
                    Case TypeOf inline Is Markdig.Syntax.Inlines.HtmlInline
                        Dim html = CType(inline, Markdig.Syntax.Inlines.HtmlInline).Tag.Trim()
                        Select Case True
                            Case html.StartsWith("<u", StringComparison.OrdinalIgnoreCase)
                                rtf.Append("\ul ")
                            Case html.StartsWith("</u", StringComparison.OrdinalIgnoreCase)
                                rtf.Append("\ulnone ")
                            Case html.StartsWith("<sup", StringComparison.OrdinalIgnoreCase)
                                rtf.Append("{\super ")
                            Case html.StartsWith("</sup", StringComparison.OrdinalIgnoreCase)
                                rtf.Append("\nosupersub} ")
                            Case html.StartsWith("<sub", StringComparison.OrdinalIgnoreCase)
                                rtf.Append("{\sub ")
                            Case html.StartsWith("</sub", StringComparison.OrdinalIgnoreCase)
                                rtf.Append("\nosupersub} ")
                            Case Else
                                rtf.Append(EscapeRtf(html))
                        End Select

                    ' EmojiInline
                    Case TypeOf inline Is Markdig.Extensions.Emoji.EmojiInline
                        Dim emo = CType(inline, Markdig.Extensions.Emoji.EmojiInline)
                        rtf.Append(EscapeRtf(emo.Content.ToString()))

                    ' Footnote link
                    Case TypeOf inline Is Markdig.Extensions.Footnotes.FootnoteLink
                        Dim fl = CType(inline, Markdig.Extensions.Footnotes.FootnoteLink)
                        HandleFootnoteLink(rtf, fl, fnDefs, visitedFootnotes)

                        ' Fallback handling (recursive for containers, otherwise ToString()).
                    Case Else
                        If TypeOf inline Is Markdig.Syntax.Inlines.ContainerInline Then
                            ConvertInline(rtf, CType(inline, Markdig.Syntax.Inlines.ContainerInline), fnDefs, visitedFootnotes)
                        Else
                            rtf.Append(EscapeRtf(inline.ToString()))
                        End If
                End Select
            Next
        End Sub


        ''' <summary>
        ''' Escapes a string for inclusion in RTF:
        ''' <list type="bullet">
        ''' <item><description>Escapes <c>\</c>, <c>{</c>, <c>}</c>.</description></item>
        ''' <item><description>Encodes non-ASCII as <c>\uN?</c> sequences.</description></item>
        ''' </list>
        ''' </summary>
        ''' <param name="text">Text to escape.</param>
        ''' <returns>RTF-escaped text.</returns>
        Private Function EscapeRtf(text As String) As String
            If String.IsNullOrEmpty(text) Then Return String.Empty
            Dim sb As New System.Text.StringBuilder()
            For Each c As Char In text
                Select Case c
                    Case "\"c : sb.Append("\\")
                    Case "{"c : sb.Append("\{")
                    Case "}"c : sb.Append("\}")
                    Case Else
                        If AscW(c) > 127 Then
                            ' RTF Unicode escape
                            sb.Append("\u" & AscW(c) & "?")
                        Else
                            sb.Append(c)
                        End If
                End Select
            Next
            Return sb.ToString()
        End Function

        ''' <summary>
        ''' Applies RTF formatting for bold, italic, and underline based on Markdown emphasis delimiters.
        ''' </summary>
        ''' <param name="rtf">Target RTF builder.</param>
        ''' <param name="e">Emphasis inline to render.</param>
        Private Sub HandleEmphasis(rtf As System.Text.StringBuilder, e As Markdig.Syntax.Inlines.EmphasisInline)
            Dim italic = (e.DelimiterChar = "*"c AndAlso e.DelimiterCount = 1) OrElse (e.DelimiterChar = "_"c AndAlso e.DelimiterCount = 1)
            Dim bold = (e.DelimiterChar = "*"c AndAlso e.DelimiterCount = 2)
            Dim underline = (e.DelimiterChar = "_"c AndAlso e.DelimiterCount = 2)

            If bold Then rtf.Append("\b ")
            If italic Then rtf.Append("\i ")
            If underline Then rtf.Append("\ul ")

            ConvertInline(rtf, e)

            If underline Then rtf.Append(" \ulnone")
            If italic Then rtf.Append(" \i0")
            If bold Then rtf.Append(" \b0")
        End Sub

        ' Track visited footnotes to prevent recursion while rendering nested footnote links.

        ''' <summary>
        ''' Renders a footnote link by writing an RTF <c>\footnote</c> group containing the footnote definition.
        ''' A visited set is used to prevent infinite recursion when footnotes reference footnotes.
        ''' </summary>
        ''' <param name="rtf">Target RTF builder.</param>
        ''' <param name="fl">Footnote link encountered in inline content.</param>
        ''' <param name="fnDefs">Footnote definitions lookup by label.</param>
        ''' <param name="visited">Footnote labels currently being rendered.</param>
        Private Sub HandleFootnoteLink(
    rtf As System.Text.StringBuilder,
    fl As FootnoteLink,
    fnDefs As Dictionary(Of String, Markdig.Extensions.Footnotes.Footnote),
    visited As HashSet(Of String)
)
            Dim label = fl.Footnote.Label

            ' Guard: skip if the footnote definition is missing from the lookup.
            If fnDefs Is Nothing OrElse Not fnDefs.ContainsKey(label) Then
                Return
            End If

            ' Prevent infinite recursion.
            If visited.Contains(label) Then
                Return
            End If
            visited.Add(label)

            ' Write the footnote content as an RTF footnote group.
            rtf.Append("{\footnote ")
            Dim def = fnDefs(label)
            For Each subBlk In def
                If TypeOf subBlk Is ParagraphBlock Then
                    ConvertInline(
                rtf,
                CType(subBlk, ParagraphBlock).Inline,
                fnDefs,
                visited)    ' Pass through visited.
                End If
            Next
            rtf.Append("}")

            ' Allow the same footnote label to be rendered again elsewhere.
            visited.Remove(label)
        End Sub

    End Module

End Namespace