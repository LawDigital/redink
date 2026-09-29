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
        ''' Builds the common Markdig pipeline used for HTML display. Advanced extensions are preserved,
        ''' but the Mathematics extension is removed because the embedded HTML viewers do not run MathJax/KaTeX.
        ''' </summary>
        Public Shared Function CreateMarkdownHtmlPipeline(
            Optional useSoftlineBreakAsHardlineBreak As Boolean = False
        ) As Markdig.MarkdownPipeline

            Return CreateMarkdownHtmlPipeline(
                useSoftlineBreakAsHardlineBreak,
                usePreciseSourceLocation:=False)
        End Function

        Public Shared Function CreateMarkdownHtmlPipeline(
            useSoftlineBreakAsHardlineBreak As Boolean,
            usePreciseSourceLocation As System.Boolean
        ) As Markdig.MarkdownPipeline

            Dim builder As New Markdig.MarkdownPipelineBuilder()
            builder.UseAdvancedExtensions()

            ' UseAdvancedExtensions includes Mathematics. Remove it explicitly while preserving
            ' every other advanced extension supported by the installed Markdig version.
            For extensionIndex As Integer = builder.Extensions.Count - 1 To 0 Step -1
                Dim extensionInstance As Object = builder.Extensions(extensionIndex)
                If extensionInstance IsNot Nothing AndAlso
                   String.Equals(extensionInstance.GetType().FullName,
                                 "Markdig.Extensions.Mathematics.MathExtension",
                                 StringComparison.Ordinal) Then
                    builder.Extensions.RemoveAt(extensionIndex)
                End If
            Next

            If useSoftlineBreakAsHardlineBreak Then
                builder.UseSoftlineBreakAsHardlineBreak()
            End If

            If usePreciseSourceLocation Then
                builder.UsePreciseSourceLocation()
            End If

            Return builder.Build()
        End Function


        ''' <summary>
        ''' Normalizes lightweight LaTeX-style notation emitted by LLMs before Markdown-to-HTML rendering.
        ''' Embedded Red Ink viewers do not run MathJax, so wrappers such as \(\rightarrow\) would otherwise
        ''' remain visible verbatim. Only a bounded allow-list is substituted; Markdown link destinations remain unchanged.
        ''' </summary>
        Public Shared Function NormalizeMarkdownForHtmlDisplay(markdown As String) As String
            If String.IsNullOrEmpty(markdown) Then Return If(markdown, String.Empty)

            ' Keep ordinary text behavior stable, but never infer subscript/superscript from
            ' underscores or carets outside an explicitly delimited math region. File names,
            ' identifiers and paths such as report_9.docx or %appdata%\Microsoft\Word must
            ' therefore remain unchanged. Explicit math is scanned deterministically below;
            ' no regular expression is used to guess mathematical content.
            Return NormalizeKnownLatexCodesPreservingMarkdownLinks(markdown)
        End Function

        Private Shared Function NormalizeKnownLatexCodesPreservingMarkdownLinks(value As String) As String
            If String.IsNullOrEmpty(value) Then Return If(value, String.Empty)

            Dim output As New System.Text.StringBuilder(value.Length + 16)
            Dim index As Integer = 0
            While index < value.Length
                Dim linkStart As Integer = FindNextMarkdownLinkStart(value, index)
                If linkStart < 0 Then
                    output.Append(NormalizeKnownLatexCodes(NormalizeExplicitMarkdownMath(value.Substring(index))))
                    Exit While
                End If

                If linkStart > index Then output.Append(NormalizeKnownLatexCodes(NormalizeExplicitMarkdownMath(value.Substring(index, linkStart - index))))

                Dim labelOpen As Integer = If(value(linkStart) = "!"c, linkStart + 1, linkStart)
                Dim labelClose As Integer = FindUnescapedMarkdownDelimiter(value, labelOpen + 1, "]"c)
                If labelClose < 0 OrElse labelClose + 1 >= value.Length OrElse value(labelClose + 1) <> "("c Then
                    output.Append(NormalizeKnownLatexCodes(NormalizeExplicitMarkdownMath(value.Substring(linkStart, 1))))
                    index = linkStart + 1
                    Continue While
                End If

                Dim destinationClose As Integer = FindMarkdownLinkDestinationClose(value, labelClose + 2)
                If destinationClose < 0 Then
                    output.Append(NormalizeKnownLatexCodes(NormalizeExplicitMarkdownMath(value.Substring(linkStart))))
                    Exit While
                End If

                If value(linkStart) = "!"c Then output.Append("!")
                output.Append("[")
                output.Append(NormalizeKnownLatexCodes(NormalizeExplicitMarkdownMath(value.Substring(labelOpen + 1, labelClose - labelOpen - 1))))
                output.Append("](")
                output.Append(value.Substring(labelClose + 2, destinationClose - labelClose - 2))
                output.Append(")")
                index = destinationClose + 1
            End While
            Return output.ToString()
        End Function

        Private Shared Function NormalizeKnownLatexCodes(value As String) As String
            If String.IsNullOrEmpty(value) Then Return If(value, String.Empty)

            Dim result As String = value

            ' Keep LaTeX normalization deliberately literal and bounded. Do not use regular
            ' expressions here: ordinary prose, paths, environment-variable notation, URLs,
            ' and other backslash-containing text must remain byte-for-byte unchanged unless
            ' it contains one of the explicitly supported LaTeX forms below.
            result = ReplaceSimpleBracedLatexCommand(result, "\text", "", "")
            result = ReplaceSimpleBracedLatexCommand(result, "\mathrm", "", "")
            result = ReplaceSimpleBracedLatexCommand(result, "\mathbf", "", "")
            result = ReplaceSimpleBracedLatexCommand(result, "\mathit", "", "")
            result = ReplaceSimpleBracedLatexCommand(result, "\xrightarrow", " —", "→ ")
            result = ReplaceSimpleBracedLatexCommand(result, "\xleftarrow", " ←", "— ")

            result = result.Replace("^{\circ}", "°")
            result = result.Replace("^\circ", "°")

            Dim replacements As New System.Collections.Generic.Dictionary(Of String, String)(StringComparer.Ordinal) From {
                {"\leftrightarrow", "↔"}, {"\rightarrow", "→"}, {"\leftarrow", "←"}, {"\to", "→"}, {"\gets", "←"},
                {"\Longleftrightarrow", "⟺"}, {"\Longrightarrow", "⟹"}, {"\Longleftarrow", "⟸"},
                {"\Rightarrow", "⇒"}, {"\Leftarrow", "⇐"}, {"\Leftrightarrow", "⇔"}, {"\implies", "⇒"}, {"\impliedby", "⇐"}, {"\iff", "⇔"}, {"\mapsto", "↦"},
                {"\Updownarrow", "⇕"}, {"\Uparrow", "⇑"}, {"\Downarrow", "⇓"}, {"\uparrow", "↑"}, {"\downarrow", "↓"}, {"\updownarrow", "↕"},
                {"\leq", "≤"}, {"\le", "≤"}, {"\geq", "≥"}, {"\ge", "≥"}, {"\neq", "≠"}, {"\ne", "≠"},
                {"\ll", "≪"}, {"\gg", "≫"}, {"\approx", "≈"}, {"\equiv", "≡"}, {"\propto", "∝"},
                {"\pm", "±"}, {"\mp", "∓"}, {"\times", "×"}, {"\cdot", "·"}, {"\div", "÷"}, {"\circ", "°"},
                {"\checkmark", "✓"}, {"\star", "★"}, {"\bullet", "•"}, {"\textdegree", "°"},
                {"\notin", "∉"}, {"\in", "∈"}, {"\subseteq", "⊆"}, {"\supseteq", "⊇"}, {"\subset", "⊂"}, {"\supset", "⊃"},
                {"\cap", "∩"}, {"\cup", "∪"}, {"\forall", "∀"}, {"\exists", "∃"}, {"\land", "∧"}, {"\lor", "∨"}, {"\neg", "¬"}
            }

            For Each pair As System.Collections.Generic.KeyValuePair(Of String, String) In System.Linq.Enumerable.OrderByDescending(replacements, Function(item) item.Key.Length)
                result = ReplaceBoundedLatexToken(result, pair.Key, pair.Value)
            Next

            result = ReplaceBoundedSectionLatexToken(result, "\S", "§")
            result = ReplaceBoundedSectionLatexToken(result, "\P", "¶")
            result = result.Replace("\,", " ")

            Return result
        End Function

        Private Shared Function ReplaceSimpleBracedLatexCommand(
            value As String,
            command As String,
            replacementPrefix As String,
            replacementSuffix As String
        ) As String
            If String.IsNullOrEmpty(value) OrElse String.IsNullOrEmpty(command) Then Return If(value, String.Empty)

            Dim startToken As String = command & "{"
            Dim output As New System.Text.StringBuilder(value.Length)
            Dim index As Integer = 0

            While index < value.Length
                Dim commandStart As Integer = value.IndexOf(startToken, index, StringComparison.Ordinal)
                If commandStart < 0 Then
                    output.Append(value.Substring(index))
                    Exit While
                End If

                output.Append(value.Substring(index, commandStart - index))
                Dim contentStart As Integer = commandStart + startToken.Length
                Dim commandEnd As Integer = value.IndexOf("}"c, contentStart)
                If commandEnd < 0 Then
                    output.Append(value.Substring(commandStart))
                    Exit While
                End If

                Dim content As String = value.Substring(contentStart, commandEnd - contentStart)
                If content.IndexOf("{"c) >= 0 OrElse content.IndexOf("}"c) >= 0 Then
                    output.Append(startToken)
                    index = contentStart
                    Continue While
                End If

                output.Append(replacementPrefix)
                output.Append(content)
                output.Append(replacementSuffix)
                index = commandEnd + 1
            End While

            Return output.ToString()
        End Function

        Private Shared Function ReplaceBoundedLatexToken(value As String, token As String, replacement As String) As String
            If String.IsNullOrEmpty(value) OrElse String.IsNullOrEmpty(token) Then Return If(value, String.Empty)

            Dim output As New System.Text.StringBuilder(value.Length)
            Dim index As Integer = 0

            While index < value.Length
                Dim tokenStart As Integer = value.IndexOf(token, index, StringComparison.Ordinal)
                If tokenStart < 0 Then
                    output.Append(value.Substring(index))
                    Exit While
                End If

                output.Append(value.Substring(index, tokenStart - index))
                Dim tokenEnd As Integer = tokenStart + token.Length
                Dim nextIsAsciiLetter As Boolean = tokenEnd < value.Length AndAlso IsAsciiLetter(value(tokenEnd))

                If nextIsAsciiLetter OrElse IsLikelyFilesystemPathToken(value, tokenStart, tokenEnd) Then
                    output.Append(token)
                Else
                    output.Append(replacement)
                End If
                index = tokenEnd
            End While

            Return output.ToString()
        End Function

        Private Shared Function ReplaceBoundedSectionLatexToken(value As String, token As String, replacement As String) As String
            If String.IsNullOrEmpty(value) OrElse String.IsNullOrEmpty(token) Then Return If(value, String.Empty)

            Dim output As New System.Text.StringBuilder(value.Length)
            Dim index As Integer = 0

            While index < value.Length
                Dim tokenStart As Integer = value.IndexOf(token, index, StringComparison.Ordinal)
                If tokenStart < 0 Then
                    output.Append(value.Substring(index))
                    Exit While
                End If

                output.Append(value.Substring(index, tokenStart - index))
                Dim tokenEnd As Integer = tokenStart + token.Length
                Dim boundaryIndex As Integer = tokenEnd
                If boundaryIndex + 1 < value.Length AndAlso value(boundaryIndex) = "\"c AndAlso value(boundaryIndex + 1) = ","c Then
                    boundaryIndex += 2
                End If

                Dim hasExpectedBoundary As Boolean = boundaryIndex < value.Length AndAlso IsLatexSectionBoundary(value(boundaryIndex))
                If hasExpectedBoundary AndAlso Not IsLikelyFilesystemPathToken(value, tokenStart, tokenEnd) Then
                    output.Append(replacement)
                Else
                    output.Append(token)
                End If
                index = tokenEnd
            End While

            Return output.ToString()
        End Function

        Private Shared Function IsLatexSectionBoundary(ch As Char) As Boolean
            Return Char.IsWhiteSpace(ch) OrElse Char.IsDigit(ch) OrElse ch = "("c OrElse ch = ")"c OrElse
                   ch = "."c OrElse ch = ","c OrElse ch = ":"c OrElse ch = ";"c
        End Function

        Private Shared Function IsAsciiLetter(ch As Char) As Boolean
            Return (ch >= "A"c AndAlso ch <= "Z"c) OrElse (ch >= "a"c AndAlso ch <= "z"c)
        End Function

        Private Shared Function IsLikelyFilesystemPathToken(value As String, tokenStart As Integer, tokenEnd As Integer) As Boolean
            If String.IsNullOrEmpty(value) OrElse tokenStart < 0 OrElse tokenStart >= value.Length Then Return False

            ' A following path separator makes the token a filesystem path segment, not a LaTeX command.
            If tokenEnd < value.Length AndAlso (value(tokenEnd) = "\"c OrElse value(tokenEnd) = "/"c) Then Return True

            Dim segmentStart As Integer = tokenStart - 1
            While segmentStart >= 0 AndAlso Not Char.IsWhiteSpace(value(segmentStart))
                segmentStart -= 1
            End While
            segmentStart += 1

            Dim prefixLength As Integer = tokenStart - segmentStart
            If prefixLength <= 0 Then Return False
            Dim prefix As String = value.Substring(segmentStart, prefixLength)

            If prefix.IndexOf(":\", StringComparison.Ordinal) >= 0 OrElse
               prefix.IndexOf(":/", StringComparison.Ordinal) >= 0 OrElse
               prefix.IndexOf("%", StringComparison.Ordinal) >= 0 OrElse
               prefix.StartsWith("\\", StringComparison.Ordinal) Then
                Return True
            End If

            Return False
        End Function

        Private Shared Function RemoveKnownLatexMathWrappers(value As String) As String
            If String.IsNullOrEmpty(value) Then Return If(value, String.Empty)

            Dim result As String = value
            result = RemoveKnownLatexMathWrapper(result, "\(", "\)")
            result = RemoveKnownLatexMathWrapper(result, "\[", "\]")
            result = RemoveKnownDollarMathWrappers(result)
            Return result
        End Function

        Private Shared Function RemoveKnownLatexMathWrapper(value As String, openToken As String, closeToken As String) As String
            Dim output As New System.Text.StringBuilder(value.Length)
            Dim index As Integer = 0

            While index < value.Length
                Dim openIndex As Integer = value.IndexOf(openToken, index, StringComparison.Ordinal)
                If openIndex < 0 Then
                    output.Append(value.Substring(index))
                    Exit While
                End If

                output.Append(value.Substring(index, openIndex - index))
                Dim contentStart As Integer = openIndex + openToken.Length
                Dim closeIndex As Integer = value.IndexOf(closeToken, contentStart, StringComparison.Ordinal)
                If closeIndex < 0 Then
                    output.Append(value.Substring(openIndex))
                    Exit While
                End If

                Dim content As String = value.Substring(contentStart, closeIndex - contentStart)
                If ContainsKnownLatexReplacementSymbol(content) AndAlso content.IndexOf(ControlChars.Cr) < 0 AndAlso content.IndexOf(ControlChars.Lf) < 0 Then
                    output.Append(content)
                Else
                    output.Append(openToken)
                    output.Append(content)
                    output.Append(closeToken)
                End If
                index = closeIndex + closeToken.Length
            End While

            Return output.ToString()
        End Function

        Private Shared Function RemoveKnownDollarMathWrappers(value As String) As String
            Dim output As New System.Text.StringBuilder(value.Length)
            Dim index As Integer = 0

            While index < value.Length
                Dim openIndex As Integer = FindNextUnescapedSingleDollar(value, index)
                If openIndex < 0 Then
                    output.Append(value.Substring(index))
                    Exit While
                End If

                output.Append(value.Substring(index, openIndex - index))
                Dim closeIndex As Integer = FindNextUnescapedSingleDollar(value, openIndex + 1)
                If closeIndex < 0 Then
                    output.Append(value.Substring(openIndex))
                    Exit While
                End If

                Dim content As String = value.Substring(openIndex + 1, closeIndex - openIndex - 1)
                If ContainsKnownLatexReplacementSymbol(content) AndAlso content.IndexOf(ControlChars.Cr) < 0 AndAlso content.IndexOf(ControlChars.Lf) < 0 Then
                    output.Append(content)
                Else
                    output.Append("$")
                    output.Append(content)
                    output.Append("$")
                End If
                index = closeIndex + 1
            End While

            Return output.ToString()
        End Function

        Private Shared Function FindNextUnescapedSingleDollar(value As String, startIndex As Integer) As Integer
            For index As Integer = Math.Max(0, startIndex) To value.Length - 1
                If value(index) <> "$"c Then Continue For
                If index > 0 AndAlso value(index - 1) = "\"c Then Continue For
                If index > 0 AndAlso value(index - 1) = "$"c Then Continue For
                If index + 1 < value.Length AndAlso value(index + 1) = "$"c Then Continue For
                Return index
            Next
            Return -1
        End Function

        Private Shared Function ContainsKnownLatexReplacementSymbol(value As String) As Boolean
            If String.IsNullOrEmpty(value) Then Return False
            Const knownSymbols As String = "≤≥≠≈≡∝≪≫→←↔⇒⇐⇔⟹⟸⟺↦↑↓↕⇑⇓⇕✓★•∈∉⊂⊆⊃⊇∩∪∀∃∧∨¬±∓×·÷°§¶"
            For Each symbol As Char In knownSymbols
                If value.IndexOf(symbol) >= 0 Then Return True
            Next
            Return False
        End Function

        Private Shared Function FindNextMarkdownLinkStart(value As String, startIndex As Integer) As Integer
            For i As Integer = Math.Max(0, startIndex) To value.Length - 2
                ' Escaped brackets include the LaTeX display wrapper \[...\]; they are not Markdown links.
                Dim escapedBracket As Boolean = (i > 0 AndAlso value(i - 1) = "\"c)
                If value(i) = "["c AndAlso Not escapedBracket Then Return i
                If value(i) = "!"c AndAlso value(i + 1) = "["c Then Return i
            Next
            Return -1
        End Function

        Private Shared Function FindUnescapedMarkdownDelimiter(value As String, startIndex As Integer, delimiter As Char) As Integer
            Dim escaped As Boolean = False
            For i As Integer = Math.Max(0, startIndex) To value.Length - 1
                Dim ch As Char = value(i)
                If escaped Then
                    escaped = False
                ElseIf ch = "\"c Then
                    escaped = True
                ElseIf ch = delimiter Then
                    Return i
                End If
            Next
            Return -1
        End Function

        Private Shared Function FindMarkdownLinkDestinationClose(value As String, startIndex As Integer) As Integer
            Dim depth As Integer = 0
            Dim escaped As Boolean = False
            For i As Integer = Math.Max(0, startIndex) To value.Length - 1
                Dim ch As Char = value(i)
                If escaped Then
                    escaped = False
                    Continue For
                End If
                If ch = "\"c Then
                    escaped = True
                    Continue For
                End If
                If ch = "("c Then
                    depth += 1
                ElseIf ch = ")"c Then
                    If depth = 0 Then Return i
                    depth -= 1
                End If
            Next
            Return -1
        End Function

        Private Shared Function NormalizeExplicitMarkdownMath(value As String) As String
            If System.String.IsNullOrEmpty(value) Then Return If(value, System.String.Empty)

            Dim output As New System.Text.StringBuilder(value.Length + 32)
            Dim index As System.Int32 = 0
            Dim convertedCount As System.Int32 = 0

            While index < value.Length
                ' Markdown code spans and fenced code blocks are opaque. A backtick run is
                ' copied through the matching run without inspecting its content for math.
                If value(index) = "`"c Then
                    Dim tickCount As System.Int32 = CountRepeatedCharacter(value, index, "`"c)
                    Dim codeCloseIndex As System.Int32 = FindMatchingCharacterRun(value, index + tickCount, "`"c, tickCount)
                    If codeCloseIndex >= 0 Then
                        output.Append(value.Substring(index, codeCloseIndex + tickCount - index))
                        index = codeCloseIndex + tickCount
                        Continue While
                    End If
                End If

                Dim openToken As System.String = Nothing
                Dim closeToken As System.String = Nothing
                Dim contentStart As System.Int32 = -1
                Dim allowLineBreaks As System.Boolean = False
                Dim requireMathSignal As System.Boolean = False

                If index + 1 < value.Length AndAlso value(index) = "$"c AndAlso value(index + 1) = "$"c AndAlso Not IsEscapedCharacter(value, index) Then
                    openToken = "$$"
                    closeToken = "$$"
                    contentStart = index + 2
                    allowLineBreaks = True
                ElseIf value(index) = "$"c AndAlso Not IsEscapedCharacter(value, index) Then
                    openToken = "$"
                    closeToken = "$"
                    contentStart = index + 1
                    requireMathSignal = True
                ElseIf index + 1 < value.Length AndAlso value(index) = "\"c AndAlso value(index + 1) = "("c Then
                    openToken = "\("
                    closeToken = "\)"
                    contentStart = index + 2
                ElseIf index + 1 < value.Length AndAlso value(index) = "\"c AndAlso value(index + 1) = "["c Then
                    openToken = "\["
                    closeToken = "\]"
                    contentStart = index + 2
                    allowLineBreaks = True
                End If

                If openToken Is Nothing Then
                    output.Append(value(index))
                    index += 1
                    Continue While
                End If

                Dim closeIndex As System.Int32 = FindExplicitMathClose(value, contentStart, closeToken, allowLineBreaks)
                If closeIndex < 0 Then
                    output.Append(openToken)
                    index += openToken.Length
                    Continue While
                End If

                Dim content As System.String = value.Substring(contentStart, closeIndex - contentStart)
                If requireMathSignal AndAlso Not HasConservativeInlineMathSignal(content) Then
                    output.Append(value.Substring(index, closeIndex + closeToken.Length - index))
                    index = closeIndex + closeToken.Length
                    Continue While
                End If

                Dim normalizedContent As System.String = NormalizeExplicitMathContent(content)
                output.Append(normalizedContent)
                convertedCount += 1
                index = closeIndex + closeToken.Length
            End While

            If convertedCount > 0 Then
                System.Diagnostics.Debug.WriteLine("[MDMATH] EXPLICIT_MATH converted=" & convertedCount.ToString(System.Globalization.CultureInfo.InvariantCulture))
            End If

            Return output.ToString()
        End Function

        Private Shared Function CountRepeatedCharacter(value As System.String, startIndex As System.Int32, character As System.Char) As System.Int32
            Dim count As System.Int32 = 0
            Dim index As System.Int32 = startIndex
            While index < value.Length AndAlso value(index) = character
                count += 1
                index += 1
            End While
            Return count
        End Function

        Private Shared Function FindMatchingCharacterRun(value As System.String, startIndex As System.Int32, character As System.Char, runLength As System.Int32) As System.Int32
            If runLength <= 0 Then Return -1
            Dim index As System.Int32 = System.Math.Max(0, startIndex)
            While index < value.Length
                If value(index) = character AndAlso CountRepeatedCharacter(value, index, character) = runLength Then Return index
                index += 1
            End While
            Return -1
        End Function

        Private Shared Function IsEscapedCharacter(value As System.String, index As System.Int32) As System.Boolean
            If System.String.IsNullOrEmpty(value) OrElse index <= 0 OrElse index >= value.Length Then Return False
            Dim slashCount As System.Int32 = 0
            Dim scan As System.Int32 = index - 1
            While scan >= 0 AndAlso value(scan) = "\"c
                slashCount += 1
                scan -= 1
            End While
            Return (slashCount Mod 2) = 1
        End Function

        Private Shared Function FindExplicitMathClose(value As System.String, startIndex As System.Int32, closeToken As System.String, allowLineBreaks As System.Boolean) As System.Int32
            Dim index As System.Int32 = System.Math.Max(0, startIndex)
            While index <= value.Length - closeToken.Length
                Dim ch As System.Char = value(index)
                If Not allowLineBreaks AndAlso (ch = Microsoft.VisualBasic.ControlChars.Cr OrElse ch = Microsoft.VisualBasic.ControlChars.Lf) Then Return -1

                If System.String.CompareOrdinal(value, index, closeToken, 0, closeToken.Length) = 0 Then
                    If closeToken = "$" Then
                        If Not IsEscapedCharacter(value, index) AndAlso
                           Not (index + 1 < value.Length AndAlso value(index + 1) = "$"c) Then Return index
                    ElseIf closeToken = "$$" Then
                        If Not IsEscapedCharacter(value, index) Then Return index
                    Else
                        Return index
                    End If
                End If
                index += 1
            End While
            Return -1
        End Function

        Private Shared Function HasConservativeInlineMathSignal(content As System.String) As System.Boolean
            If System.String.IsNullOrWhiteSpace(content) Then Return False

            ' Single-dollar notation is ambiguous with currency. Only normalize it when the
            ' content contains an unmistakable LaTeX/script signal. Display math ($$...$$)
            ' and explicit \(...\)/\[...\] delimiters do not need this heuristic.
            If content.IndexOf("\"c) >= 0 OrElse content.IndexOf("_"c) >= 0 OrElse content.IndexOf("^"c) >= 0 Then Return True
            If ContainsKnownLatexReplacementSymbol(content) Then Return True

            Dim trimmed As System.String = content.Trim()
            Return trimmed = ">" OrElse trimmed = "<"
        End Function

        Private Shared Function NormalizeExplicitMathContent(content As System.String) As System.String
            If System.String.IsNullOrEmpty(content) Then Return If(content, System.String.Empty)

            Dim result As System.String = content
            result = ReplaceSimpleMathFunction(result, "\frac", MathFunctionKind.Fraction)
            result = ReplaceSimpleMathFunction(result, "\sqrt", MathFunctionKind.SquareRoot)

            Dim mathReplacements As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.Ordinal) From {
                {"\sum", "∑"}, {"\prod", "∏"}, {"\int", "∫"}, {"\infty", "∞"},
                {"\alpha", "α"}, {"\beta", "β"}, {"\gamma", "γ"}, {"\delta", "δ"}, {"\pi", "π"}, {"\Omega", "Ω"},
                {"\leftrightarrow", "↔"}, {"\rightarrow", "→"}, {"\leftarrow", "←"}, {"\to", "→"}, {"\gets", "←"},
                {"\leq", "≤"}, {"\le", "≤"}, {"\geq", "≥"}, {"\ge", "≥"}, {"\neq", "≠"}, {"\ne", "≠"},
                {"\approx", "≈"}, {"\equiv", "≡"}, {"\pm", "±"}, {"\mp", "∓"}, {"\times", "×"}, {"\cdot", "·"}, {"\div", "÷"}
            }
            For Each pair As System.Collections.Generic.KeyValuePair(Of System.String, System.String) In System.Linq.Enumerable.OrderByDescending(mathReplacements, Function(item) item.Key.Length)
                result = ReplaceBoundedLatexToken(result, pair.Key, pair.Value)
            Next

            result = result.Replace("^{\circ}", "°").Replace("^\circ", "°")
            result = NormalizeSimpleLatexSubscriptsAndSuperscripts(result)
            Return result.Trim()
        End Function

        Private Enum MathFunctionKind
            Fraction
            SquareRoot
        End Enum

        Private Shared Function ReplaceSimpleMathFunction(value As System.String, command As System.String, kind As MathFunctionKind) As System.String
            If System.String.IsNullOrEmpty(value) Then Return If(value, System.String.Empty)

            Dim output As New System.Text.StringBuilder(value.Length)
            Dim index As System.Int32 = 0
            While index < value.Length
                Dim commandIndex As System.Int32 = value.IndexOf(command, index, System.StringComparison.Ordinal)
                If commandIndex < 0 Then
                    output.Append(value.Substring(index))
                    Exit While
                End If

                output.Append(value.Substring(index, commandIndex - index))
                Dim argumentStart As System.Int32 = commandIndex + command.Length
                If argumentStart >= value.Length OrElse value(argumentStart) <> "{"c Then
                    output.Append(command)
                    index = argumentStart
                    Continue While
                End If

                Dim firstClose As System.Int32 = FindBalancedBraceClose(value, argumentStart)
                If firstClose < 0 Then
                    output.Append(value.Substring(commandIndex))
                    Exit While
                End If
                Dim firstArgument As System.String = value.Substring(argumentStart + 1, firstClose - argumentStart - 1)

                If kind = MathFunctionKind.SquareRoot Then
                    output.Append("√(")
                    output.Append(NormalizeExplicitMathContent(firstArgument))
                    output.Append(")")
                    index = firstClose + 1
                    Continue While
                End If

                Dim secondStart As System.Int32 = firstClose + 1
                If secondStart >= value.Length OrElse value(secondStart) <> "{"c Then
                    output.Append(value.Substring(commandIndex, firstClose - commandIndex + 1))
                    index = firstClose + 1
                    Continue While
                End If
                Dim secondClose As System.Int32 = FindBalancedBraceClose(value, secondStart)
                If secondClose < 0 Then
                    output.Append(value.Substring(commandIndex))
                    Exit While
                End If

                Dim secondArgument As System.String = value.Substring(secondStart + 1, secondClose - secondStart - 1)
                output.Append("(")
                output.Append(NormalizeExplicitMathContent(firstArgument))
                output.Append(")/(")
                output.Append(NormalizeExplicitMathContent(secondArgument))
                output.Append(")")
                index = secondClose + 1
            End While
            Return output.ToString()
        End Function

        Private Shared Function FindBalancedBraceClose(value As System.String, openIndex As System.Int32) As System.Int32
            If openIndex < 0 OrElse openIndex >= value.Length OrElse value(openIndex) <> "{"c Then Return -1
            Dim depth As System.Int32 = 0
            For index As System.Int32 = openIndex To value.Length - 1
                If value(index) = "{"c Then
                    depth += 1
                ElseIf value(index) = "}"c Then
                    depth -= 1
                    If depth = 0 Then Return index
                End If
            Next
            Return -1
        End Function

        Private Shared Function NormalizeSimpleLatexSubscriptsAndSuperscripts(value As String) As String
            If System.String.IsNullOrEmpty(value) Then Return If(value, System.String.Empty)

            Dim output As New System.Text.StringBuilder(value.Length)
            Dim index As System.Int32 = 0
            While index < value.Length
                Dim marker As System.Char = value(index)
                If marker <> "_"c AndAlso marker <> "^"c Then
                    output.Append(marker)
                    index += 1
                    Continue While
                End If

                If index + 1 >= value.Length Then
                    output.Append(marker)
                    index += 1
                    Continue While
                End If

                Dim scriptText As System.String = Nothing
                Dim consumedLength As System.Int32 = 0
                If value(index + 1) = "{"c Then
                    Dim closeIndex As System.Int32 = FindBalancedBraceClose(value, index + 1)
                    If closeIndex > index + 2 Then
                        scriptText = value.Substring(index + 2, closeIndex - index - 2)
                        consumedLength = closeIndex - index + 1
                    End If
                Else
                    scriptText = value.Substring(index + 1, 1)
                    consumedLength = 2
                End If

                If System.String.IsNullOrEmpty(scriptText) Then
                    output.Append(marker)
                    index += 1
                    Continue While
                End If

                Dim mappedScript As System.String = MapUnicodeMathScript(scriptText, marker = "_"c)
                If mappedScript Is Nothing Then
                    output.Append(value.Substring(index, consumedLength))
                Else
                    output.Append(mappedScript)
                End If
                index += consumedLength
            End While

            Return output.ToString()
        End Function

        Private Shared Function MapUnicodeMathScript(scriptText As System.String, isSubscript As System.Boolean) As System.String
            If System.String.IsNullOrEmpty(scriptText) Then Return Nothing
            Dim output As New System.Text.StringBuilder(scriptText.Length)
            For Each ch As System.Char In scriptText
                Dim mapped As System.String = If(isSubscript, MapSubscriptCharacter(ch), MapSuperscriptCharacter(ch))
                If mapped Is Nothing Then Return Nothing
                output.Append(mapped)
            Next
            Return output.ToString()
        End Function

        Private Shared Function MapSubscriptCharacter(ch As System.Char) As System.String
            Select Case ch
                Case "0"c : Return "₀"
                Case "1"c : Return "₁"
                Case "2"c : Return "₂"
                Case "3"c : Return "₃"
                Case "4"c : Return "₄"
                Case "5"c : Return "₅"
                Case "6"c : Return "₆"
                Case "7"c : Return "₇"
                Case "8"c : Return "₈"
                Case "9"c : Return "₉"
                Case "+"c : Return "₊"
                Case "-"c : Return "₋"
                Case "="c : Return "₌"
                Case "("c : Return "₍"
                Case ")"c : Return "₎"
                Case "a"c : Return "ₐ"
                Case "e"c : Return "ₑ"
                Case "h"c : Return "ₕ"
                Case "i"c : Return "ᵢ"
                Case "j"c : Return "ⱼ"
                Case "k"c : Return "ₖ"
                Case "l"c : Return "ₗ"
                Case "m"c : Return "ₘ"
                Case "n"c : Return "ₙ"
                Case "o"c : Return "ₒ"
                Case "p"c : Return "ₚ"
                Case "r"c : Return "ᵣ"
                Case "s"c : Return "ₛ"
                Case "t"c : Return "ₜ"
                Case "u"c : Return "ᵤ"
                Case "v"c : Return "ᵥ"
                Case "x"c : Return "ₓ"
                Case Else : Return Nothing
            End Select
        End Function

        Private Shared Function MapSuperscriptCharacter(ch As System.Char) As System.String
            Select Case ch
                Case "0"c : Return "⁰"
                Case "1"c : Return "¹"
                Case "2"c : Return "²"
                Case "3"c : Return "³"
                Case "4"c : Return "⁴"
                Case "5"c : Return "⁵"
                Case "6"c : Return "⁶"
                Case "7"c : Return "⁷"
                Case "8"c : Return "⁸"
                Case "9"c : Return "⁹"
                Case "+"c : Return "⁺"
                Case "-"c : Return "⁻"
                Case "="c : Return "⁼"
                Case "("c : Return "⁽"
                Case ")"c : Return "⁾"
                Case "i"c : Return "ⁱ"
                Case "n"c : Return "ⁿ"
                Case Else : Return Nothing
            End Select
        End Function

        ''' <summary>
        ''' Validates heading and list nesting against an optional native paragraph-style map.
        ''' When a native body-style map is present, headings and lists are strict: every
        ''' encountered heading/list level must have an explicit semantic mapping.
        ''' </summary>
        Public Shared Function ValidateMarkdownParagraphStyleMap(
            markdownText As String,
            paragraphStyleMap As System.Collections.Generic.IDictionary(Of String, String),
            ByRef validationError As String
        ) As Boolean
            validationError = System.String.Empty
            If paragraphStyleMap Is Nothing OrElse paragraphStyleMap.Count = 0 OrElse System.String.IsNullOrWhiteSpace(markdownText) Then Return True

            Try
                Dim pipeline As Markdig.MarkdownPipeline = CreateMarkdownHtmlPipeline(useSoftlineBreakAsHardlineBreak:=True)
                Dim html As String = Markdig.Markdown.ToHtml(NormalizeMarkdownForHtmlDisplay(markdownText), pipeline)
                Dim htmlDoc As New HtmlAgilityPack.HtmlDocument()
                htmlDoc.LoadHtml(html)

                Dim headings As HtmlAgilityPack.HtmlNodeCollection = htmlDoc.DocumentNode.SelectNodes("//h1 | //h2 | //h3 | //h4 | //h5 | //h6")
                If headings IsNot Nothing Then
                    For Each heading As HtmlAgilityPack.HtmlNode In headings
                        If HtmlNodeHasAncestor(heading, "table") Then Continue For
                        Dim level As Integer = 0
                        If heading.Name.Length = 2 AndAlso System.Int32.TryParse(heading.Name.Substring(1), level) Then
                            Dim semantic As String = "heading" & level.ToString(System.Globalization.CultureInfo.InvariantCulture)
                            If Not paragraphStyleMap.ContainsKey(semantic) Then
                                validationError = "The selected Word design does not permit Markdown heading level " & level.ToString(System.Globalization.CultureInfo.InvariantCulture) & ". Allowed heading levels are: " & BuildMappedStyleLevelList(paragraphStyleMap, "heading") & "."
                                Return False
                            End If
                        End If
                    Next
                End If

                Dim listItems As HtmlAgilityPack.HtmlNodeCollection = htmlDoc.DocumentNode.SelectNodes("//li")
                If listItems IsNot Nothing Then
                    For Each item As HtmlAgilityPack.HtmlNode In listItems
                        If HtmlNodeHasAncestor(item, "table") Then Continue For
                        Dim owningList As HtmlAgilityPack.HtmlNode = item.ParentNode
                        If owningList Is Nothing Then Continue For

                        Dim isBullet As Boolean = owningList.Name.Equals("ul", System.StringComparison.OrdinalIgnoreCase)
                        Dim isNumbered As Boolean = owningList.Name.Equals("ol", System.StringComparison.OrdinalIgnoreCase)
                        If Not isBullet AndAlso Not isNumbered Then Continue For

                        Dim level As Integer = 1
                        Dim ancestor As HtmlAgilityPack.HtmlNode = owningList.ParentNode
                        While ancestor IsNot Nothing
                            If ancestor.Name.Equals("ul", System.StringComparison.OrdinalIgnoreCase) OrElse ancestor.Name.Equals("ol", System.StringComparison.OrdinalIgnoreCase) Then level += 1
                            ancestor = ancestor.ParentNode
                        End While

                        Dim prefix As String = If(isBullet, "bullet", "numbered")
                        Dim semantic As String = prefix & level.ToString(System.Globalization.CultureInfo.InvariantCulture)
                        If Not paragraphStyleMap.ContainsKey(semantic) Then
                            validationError = "The selected Word design does not permit " & If(isBullet, "bullet", "numbered-list") & " nesting level " & level.ToString(System.Globalization.CultureInfo.InvariantCulture) & ". Allowed " & prefix & " levels are: " & BuildMappedStyleLevelList(paragraphStyleMap, prefix) & "."
                            Return False
                        End If
                    Next
                End If

                Dim hasQuoteStyle As Boolean = paragraphStyleMap.Keys.Any(
                    Function(key As String) System.Text.RegularExpressions.Regex.IsMatch(If(key, ""), "^quote[1-9]$", System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.CultureInvariant))
                If hasQuoteStyle Then
                    Dim blockQuotes As HtmlAgilityPack.HtmlNodeCollection = htmlDoc.DocumentNode.SelectNodes("//blockquote")
                    If blockQuotes IsNot Nothing Then
                        For Each blockQuote As HtmlAgilityPack.HtmlNode In blockQuotes
                            If HtmlNodeHasAncestor(blockQuote, "table") Then Continue For
                            Dim level As Integer = 1
                            Dim ancestor As HtmlAgilityPack.HtmlNode = blockQuote.ParentNode
                            While ancestor IsNot Nothing
                                If ancestor.Name.Equals("blockquote", System.StringComparison.OrdinalIgnoreCase) Then level += 1
                                ancestor = ancestor.ParentNode
                            End While

                            Dim semantic As String = "quote" & level.ToString(System.Globalization.CultureInfo.InvariantCulture)
                            If Not paragraphStyleMap.ContainsKey(semantic) Then
                                validationError = "The selected Word design does not permit block-quote nesting level " & level.ToString(System.Globalization.CultureInfo.InvariantCulture) & ". Allowed quote levels are: " & BuildMappedStyleLevelList(paragraphStyleMap, "quote") & "."
                                Return False
                            End If
                        Next
                    End If
                End If

                Return True
            Catch ex As System.Exception
                validationError = "Markdown body-style validation failed: " & ex.Message
                Return False
            End Try
        End Function

        Private Shared Function HtmlNodeHasAncestor(node As HtmlAgilityPack.HtmlNode, ancestorName As String) As Boolean
            If node Is Nothing OrElse System.String.IsNullOrWhiteSpace(ancestorName) Then Return False
            Dim current As HtmlAgilityPack.HtmlNode = node.ParentNode
            While current IsNot Nothing
                If System.String.Equals(current.Name, ancestorName, System.StringComparison.OrdinalIgnoreCase) Then Return True
                current = current.ParentNode
            End While
            Return False
        End Function

        Private Shared Function BuildMappedStyleLevelList(
            paragraphStyleMap As System.Collections.Generic.IDictionary(Of String, String),
            prefix As String
        ) As String
            Dim levels As New System.Collections.Generic.List(Of Integer)()
            If paragraphStyleMap Is Nothing Then Return "none"
            For Each key As String In paragraphStyleMap.Keys
                If key Is Nothing OrElse Not key.StartsWith(prefix, System.StringComparison.OrdinalIgnoreCase) Then Continue For
                Dim suffix As String = key.Substring(prefix.Length)
                Dim level As Integer = 0
                If System.Int32.TryParse(suffix, level) AndAlso level > 0 Then levels.Add(level)
            Next
            If levels.Count = 0 Then Return "none"
            Return System.String.Join(", ", levels.Distinct().OrderBy(Function(level As Integer) level).Select(Function(level As Integer) level.ToString(System.Globalization.CultureInfo.InvariantCulture)))
        End Function

        ''' <summary>
        ''' Removes renderer-only boundary whitespace introduced by Markdown-to-HTML formatting.
        ''' Internal whitespace between inline nodes is preserved. This is intentionally shared by
        ''' both legacy Word HTML insertion and OOXML Word rendering so Markdig formatting newlines
        ''' cannot become visible leading/trailing spaces in generated document paragraphs.
        ''' </summary>
        Public Shared Sub NormalizeMarkdigHtmlBlockBoundaryWhitespace(htmlDoc As HtmlAgilityPack.HtmlDocument)
            If htmlDoc Is Nothing OrElse htmlDoc.DocumentNode Is Nothing Then Return

            Dim blocks As HtmlAgilityPack.HtmlNodeCollection =
                htmlDoc.DocumentNode.SelectNodes("//p | //li | //h1 | //h2 | //h3 | //h4 | //h5 | //h6 | //td | //th | //blockquote")
            If blocks Is Nothing Then Return

            For Each block As HtmlAgilityPack.HtmlNode In blocks
                If block Is Nothing Then Continue For

                Dim textNodes As System.Collections.Generic.List(Of HtmlAgilityPack.HtmlNode) =
                    block.DescendantsAndSelf().Where(
                        Function(candidate As HtmlAgilityPack.HtmlNode)
                            If candidate Is Nothing OrElse candidate.NodeType <> HtmlAgilityPack.HtmlNodeType.Text Then Return False

                            Dim ancestor As HtmlAgilityPack.HtmlNode = candidate.ParentNode
                            Do While ancestor IsNot Nothing AndAlso ancestor IsNot block
                                Dim name As String = If(ancestor.Name, "").ToLowerInvariant()
                                If name = "ul" OrElse name = "ol" OrElse name = "table" OrElse name = "pre" OrElse
                                   name = "p" OrElse name = "blockquote" OrElse System.Text.RegularExpressions.Regex.IsMatch(name, "^h[1-6]$") Then
                                    Return False
                                End If
                                ancestor = ancestor.ParentNode
                            Loop
                            Return True
                        End Function).ToList()

                If textNodes.Count = 0 Then Continue For

                Dim firstMeaningful As Integer = textNodes.FindIndex(
                    Function(item As HtmlAgilityPack.HtmlNode) Not System.String.IsNullOrWhiteSpace(HtmlAgilityPack.HtmlEntity.DeEntitize(item.InnerText)))
                Dim lastMeaningful As Integer = textNodes.FindLastIndex(
                    Function(item As HtmlAgilityPack.HtmlNode) Not System.String.IsNullOrWhiteSpace(HtmlAgilityPack.HtmlEntity.DeEntitize(item.InnerText)))

                If firstMeaningful < 0 Then Continue For

                For index As Integer = 0 To firstMeaningful - 1
                    textNodes(index).InnerHtml = System.String.Empty
                Next
                For index As Integer = lastMeaningful + 1 To textNodes.Count - 1
                    textNodes(index).InnerHtml = System.String.Empty
                Next

                Dim firstValue As String = HtmlAgilityPack.HtmlEntity.DeEntitize(textNodes(firstMeaningful).InnerText).TrimStart()
                textNodes(firstMeaningful).InnerHtml = HtmlAgilityPack.HtmlEntity.Entitize(firstValue)

                Dim lastValue As String = HtmlAgilityPack.HtmlEntity.DeEntitize(textNodes(lastMeaningful).InnerText).TrimEnd()
                textNodes(lastMeaningful).InnerHtml = HtmlAgilityPack.HtmlEntity.Entitize(lastValue)
            Next
        End Sub

        ''' <summary>
        ''' Converts the provided Markdown text to HTML and inserts it into the given Word selection.
        ''' </summary>
        ''' <param name="selection">A Word <see cref="Microsoft.Office.Interop.Word.Selection"/> (passed as <see cref="Object"/>).</param>
        ''' <param name="gptResult">The Markdown text to convert and insert.</param>
        ''' <param name="TrailingCR">If <c>True</c>, keeps trailing paragraph breaks; otherwise suppresses them.</param>
        Public Shared Sub InsertTextWithMarkdown(selection As Object,
                                                 gptResult As String,
                                                 TrailingCR As Boolean,
                                                 Optional UseHostDefaultFontColor As Boolean = False,
                                                 Optional PreserveDestinationParagraphFormatting As Boolean = False,
                                                 Optional FormattingSourceRange As Microsoft.Office.Interop.Word.Range = Nothing)

            Dim wordSelection As Microsoft.Office.Interop.Word.Selection = CType(selection, Microsoft.Office.Interop.Word.Selection)
            Dim wordRange As Microsoft.Office.Interop.Word.Range = wordSelection.Range

            Debug.WriteLine("ITWM: " & gptResult)

            Dim fullhtml As System.String = ConvertMarkdownToHtmlForWordInsertion(gptResult)

            Debug.WriteLine("ITWM: " & fullhtml)

            InsertTextWithFormat(
                fullhtml,
                wordRange,
                True,
                Not TrailingCR,
                UseHostDefaultFontColor,
                PreserveDestinationParagraphFormatting,
                FormattingSourceRange)

        End Sub

        ''' <summary>
        ''' Converts Markdown to the normalized HTML fragment consumed by the Word/Outlook
        ''' CF_HTML insertion pipeline. This is the single Markdown-to-HTML contract for Office
        ''' insertion; host-specific placeholder or document preprocessing must happen before it.
        ''' </summary>
        Public Shared Function ConvertMarkdownToHtmlForWordInsertion(ByVal markdown As System.String) As System.String
            Dim normalizedMarkdown As System.String = If(markdown, System.String.Empty)

            normalizedMarkdown = normalizedMarkdown.Replace(vbLf & " " & vbLf, vbLf & vbLf)

            Dim blankLinePattern As System.String = "((\r\n|\n|\r){2,})"
            normalizedMarkdown = System.Text.RegularExpressions.Regex.Replace(
                normalizedMarkdown,
                blankLinePattern,
                Function(match As System.Text.RegularExpressions.Match) As System.String
                    If match.Index + match.Length = normalizedMarkdown.Length Then
                        Return match.Value
                    End If

                    Dim breaks As System.String = match.Value
                    Dim regexBreaks As New System.Text.RegularExpressions.Regex("(\r\n|\n|\r)")
                    Dim splitBreaks As System.Text.RegularExpressions.MatchCollection = regexBreaks.Matches(breaks)
                    If splitBreaks.Count <= 1 Then Return breaks

                    Dim replacement As System.String = splitBreaks(0).Value
                    For breakIndex As System.Int32 = 1 To splitBreaks.Count - 1
                        replacement &= vbCrLf & "&nbsp;" & vbCrLf & splitBreaks(breakIndex).Value
                    Next
                    Return replacement
                End Function)

            Dim pipeline As Markdig.MarkdownPipeline =
                CreateMarkdownHtmlPipeline(useSoftlineBreakAsHardlineBreak:=True)

            Dim htmlResult As System.String =
                Markdig.Markdown.ToHtml(NormalizeMarkdownForHtmlDisplay(normalizedMarkdown), pipeline)

            ' Renderer-only line breaks are removed for stable CF_HTML, but literal line
            ' boundaries inside <pre> blocks are opaque source content and must survive.
            Dim protectedPreformattedBlocks As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.Ordinal)
            htmlResult = ProtectHtmlPreformattedBlocks(htmlResult, protectedPreformattedBlocks)

            htmlResult = htmlResult _
                .Replace(vbCrLf, System.String.Empty) _
                .Replace(vbCr, System.String.Empty) _
                .Replace(vbLf, System.String.Empty)

            htmlResult = RestoreHtmlPreformattedBlocks(htmlResult, protectedPreformattedBlocks)

            Dim htmlDocument As New HtmlAgilityPack.HtmlDocument()
            htmlDocument.LoadHtml(htmlResult)
            NormalizeMarkdigHtmlBlockBoundaryWhitespace(htmlDocument)

            Dim normalizedHtml As System.String = htmlDocument.DocumentNode.OuterHtml
            If normalizedMarkdown.IndexOf("**", System.StringComparison.Ordinal) >= 0 AndAlso
               normalizedHtml.IndexOf("**", System.StringComparison.Ordinal) >= 0 Then
                System.Diagnostics.Debug.WriteLine(
                    "[MDINLINE] Literal '**' survived Markdown-to-HTML conversion. Inspect source whitespace/escaping/code context before applying any delimiter rewrite.")
            End If

            Return normalizedHtml
        End Function

        Private Shared Function ProtectHtmlPreformattedBlocks(
            ByVal html As System.String,
            ByVal protectedBlocks As System.Collections.Generic.Dictionary(Of System.String, System.String)
        ) As System.String
            If System.String.IsNullOrEmpty(html) OrElse protectedBlocks Is Nothing Then Return If(html, System.String.Empty)

            Dim output As New System.Text.StringBuilder(html.Length)
            Dim index As System.Int32 = 0
            Dim blockIndex As System.Int32 = 0

            While index < html.Length
                Dim preStart As System.Int32 = html.IndexOf("<pre", index, System.StringComparison.OrdinalIgnoreCase)
                If preStart < 0 Then
                    output.Append(html.Substring(index))
                    Exit While
                End If

                Dim openEnd As System.Int32 = html.IndexOf(">"c, preStart)
                If openEnd < 0 Then
                    output.Append(html.Substring(index))
                    Exit While
                End If

                Dim preEnd As System.Int32 = html.IndexOf("</pre>", openEnd + 1, System.StringComparison.OrdinalIgnoreCase)
                If preEnd < 0 Then
                    output.Append(html.Substring(index))
                    Exit While
                End If

                output.Append(html.Substring(index, preStart - index))
                Dim blockEnd As System.Int32 = preEnd + "</pre>".Length
                Dim blockHtml As System.String = html.Substring(preStart, blockEnd - preStart)
                blockIndex += 1
                Dim token As System.String =
                    "RIPREFORMATTED" & System.Guid.NewGuid().ToString("N") &
                    "B" & blockIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) & "END"
                protectedBlocks(token) = blockHtml
                output.Append(token)
                index = blockEnd
            End While

            Return output.ToString()
        End Function

        Private Shared Function RestoreHtmlPreformattedBlocks(
            ByVal html As System.String,
            ByVal protectedBlocks As System.Collections.Generic.Dictionary(Of System.String, System.String)
        ) As System.String
            If System.String.IsNullOrEmpty(html) OrElse protectedBlocks Is Nothing OrElse protectedBlocks.Count = 0 Then
                Return If(html, System.String.Empty)
            End If

            Dim result As System.String = html
            For Each pair As System.Collections.Generic.KeyValuePair(Of System.String, System.String) In protectedBlocks
                result = result.Replace(pair.Key, pair.Value)
            Next
            Return result
        End Function

        Private Shared Sub NormalizeHtmlStrikethroughAndTaskCheckboxes(ByVal htmlDocument As HtmlAgilityPack.HtmlDocument)
            If htmlDocument Is Nothing OrElse htmlDocument.DocumentNode Is Nothing Then Return

            ' Markdig's GFM strike node is <del>. Word can interpret <del> as tracked deletion
            ' semantics rather than visual strikethrough. Convert only that explicit HTML node
            ' to a neutral span with CSS line-through; ordinary text is never inspected.
            Dim deletionNodes As HtmlAgilityPack.HtmlNodeCollection = htmlDocument.DocumentNode.SelectNodes("//del")
            If deletionNodes IsNot Nothing Then
                For Each deletionNode As HtmlAgilityPack.HtmlNode In deletionNodes.ToList()
                    Dim parent As HtmlAgilityPack.HtmlNode = deletionNode.ParentNode
                    If parent Is Nothing Then Continue For

                    Dim strikeSpan As HtmlAgilityPack.HtmlNode = htmlDocument.CreateElement("span")
                    For Each attribute As HtmlAgilityPack.HtmlAttribute In deletionNode.Attributes
                        If Not attribute.Name.Equals("style", System.StringComparison.OrdinalIgnoreCase) Then
                            strikeSpan.SetAttributeValue(attribute.Name, attribute.Value)
                        End If
                    Next

                    Dim existingStyle As System.String = deletionNode.GetAttributeValue("style", System.String.Empty).Trim()
                    If existingStyle.Length > 0 AndAlso Not existingStyle.EndsWith(";", System.StringComparison.Ordinal) Then existingStyle &= ";"
                    strikeSpan.SetAttributeValue("style", existingStyle & "text-decoration:line-through;")

                    For Each child As HtmlAgilityPack.HtmlNode In deletionNode.ChildNodes.ToList()
                        deletionNode.RemoveChild(child)
                        strikeSpan.AppendChild(child)
                    Next
                    parent.ReplaceChild(strikeSpan, deletionNode)
                Next
            End If

            ' GFM task-list state is explicit in Markdig's <input type=checkbox>. Word's HTML
            ' importer drops the control, so replace only those explicit checkbox elements with
            ' deterministic Unicode state glyphs. No textual [x]/[ ] guessing is performed.
            Dim checkboxNodes As HtmlAgilityPack.HtmlNodeCollection =
                htmlDocument.DocumentNode.SelectNodes("//input[translate(@type,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')='checkbox']")
            If checkboxNodes IsNot Nothing Then
                For Each checkboxNode As HtmlAgilityPack.HtmlNode In checkboxNodes.ToList()
                    Dim parent As HtmlAgilityPack.HtmlNode = checkboxNode.ParentNode
                    If parent Is Nothing Then Continue For
                    Dim isChecked As System.Boolean = checkboxNode.Attributes("checked") IsNot Nothing
                    parent.ReplaceChild(htmlDocument.CreateTextNode(If(isChecked, "☒", "☐")), checkboxNode)
                Next
            End If
        End Sub

        Private NotInheritable Class HtmlListImportDescriptor
            Public Property Token As System.String
            Public Property ContainerId As System.Int32
            Public Property ItemIndex As System.Int32
            Public Property Level As System.Int32
            Public Property IsOrdered As System.Boolean
            Public Property StartAt As System.Int32 = 1
        End Class

        Private NotInheritable Class HtmlListImportTarget
            Public Property Source As HtmlListImportDescriptor
            Public Property ParagraphRange As Microsoft.Office.Interop.Word.Range
            Public Property ImportedLeftIndent As System.Single
            Public Property ImportedFirstLineIndent As System.Single
        End Class

        ''' <summary>
        ''' Adds temporary, text-only anchors to HTML list items immediately before CF_HTML import.
        ''' The anchors let the post-paste renderer map each source &lt;li&gt; to the exact Word
        ''' paragraph even when Word imports a nested list item as ordinary text instead of a
        ''' native Word list. The Markdown-to-HTML DOM semantics are not changed.
        ''' </summary>
        Private Shared Function PrepareHtmlListImportMarkers(
            ByVal htmlDocument As HtmlAgilityPack.HtmlDocument
        ) As System.Collections.Generic.List(Of HtmlListImportDescriptor)

            Dim descriptors As New System.Collections.Generic.List(Of HtmlListImportDescriptor)()
            If htmlDocument Is Nothing OrElse htmlDocument.DocumentNode Is Nothing Then Return descriptors

            Dim listContainers As HtmlAgilityPack.HtmlNodeCollection =
                htmlDocument.DocumentNode.SelectNodes("//ul | //ol")
            If listContainers Is Nothing Then Return descriptors

            Dim nonce As System.String = System.Guid.NewGuid().ToString("N")
            Dim containerId As System.Int32 = 0

            For Each listContainer As HtmlAgilityPack.HtmlNode In listContainers
                Dim directItems As HtmlAgilityPack.HtmlNodeCollection = listContainer.SelectNodes("./li")
                If directItems Is Nothing OrElse directItems.Count = 0 Then Continue For

                containerId += 1

                Dim level As System.Int32 = 1
                Dim ancestor As HtmlAgilityPack.HtmlNode = listContainer.ParentNode
                Do While ancestor IsNot Nothing
                    If ancestor.Name.Equals("ul", System.StringComparison.OrdinalIgnoreCase) OrElse
                       ancestor.Name.Equals("ol", System.StringComparison.OrdinalIgnoreCase) Then
                        level += 1
                    End If
                    ancestor = ancestor.ParentNode
                Loop

                Dim isOrdered As System.Boolean =
                    listContainer.Name.Equals("ol", System.StringComparison.OrdinalIgnoreCase)

                Dim orderedStartAt As System.Int32 = 1
                If isOrdered Then
                    Dim startAttribute As System.String = listContainer.GetAttributeValue("start", "1")
                    Dim parsedStartAt As System.Int32
                    If System.Int32.TryParse(
                        startAttribute,
                        System.Globalization.NumberStyles.Integer,
                        System.Globalization.CultureInfo.InvariantCulture,
                        parsedStartAt) AndAlso parsedStartAt > 0 Then
                        orderedStartAt = parsedStartAt
                    End If
                End If

                Dim itemIndex As System.Int32 = 0
                For Each listItem As HtmlAgilityPack.HtmlNode In directItems
                    itemIndex += 1

                    Dim descriptor As New HtmlListImportDescriptor With {
                        .Token = "RILIST" & nonce &
                                 "C" & containerId.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                                 "I" & itemIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                                 "END",
                        .ContainerId = containerId,
                        .ItemIndex = itemIndex,
                        .Level = level,
                        .IsOrdered = isOrdered,
                        .StartAt = orderedStartAt
                    }

                    ' Put the marker into the first content paragraph when one exists. This avoids
                    ' creating a marker-only paragraph for the common <li><p>...</p><ol>...</ol></li>
                    ' shape while still handling inline-only list items.
                    Dim markerHost As HtmlAgilityPack.HtmlNode = listItem
                    For Each child As HtmlAgilityPack.HtmlNode In listItem.ChildNodes
                        If child.Name.Equals("ul", System.StringComparison.OrdinalIgnoreCase) OrElse
                           child.Name.Equals("ol", System.StringComparison.OrdinalIgnoreCase) Then
                            Continue For
                        End If

                        If child.Name.Equals("p", System.StringComparison.OrdinalIgnoreCase) Then
                            markerHost = child
                        End If
                        Exit For
                    Next

                    markerHost.PrependChild(htmlDocument.CreateTextNode(descriptor.Token))
                    descriptors.Add(descriptor)
                Next
            Next

            Return descriptors
        End Function

        Private Shared Function FindHtmlListImportTokenRange(
            ByVal searchScope As Microsoft.Office.Interop.Word.Range,
            ByVal token As System.String
        ) As Microsoft.Office.Interop.Word.Range

            If searchScope Is Nothing OrElse System.String.IsNullOrEmpty(token) Then Return Nothing

            Dim tokenRange As Microsoft.Office.Interop.Word.Range = searchScope.Duplicate()
            With tokenRange.Find
                .ClearFormatting()
                .Replacement.ClearFormatting()
                .Text = token
                .Forward = True
                .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
                .Format = False
                .MatchWildcards = False
            End With

            If tokenRange.Find.Execute() Then Return tokenRange
            Return Nothing
        End Function

        Private Shared Function WordListFormatMatchesHtmlKind(
            ByVal listFormat As Microsoft.Office.Interop.Word.ListFormat,
            ByVal isOrdered As System.Boolean
        ) As System.Boolean

            If listFormat Is Nothing Then Return False

            Dim listType As Microsoft.Office.Interop.Word.WdListType = listFormat.ListType
            If listType = Microsoft.Office.Interop.Word.WdListType.wdListNoNumbering Then Return False

            ' For multilevel/mixed templates ListType alone is not enough: a level can be a
            ' bullet inside an outline-numbered template. Prefer the active ListLevel's
            ' NumberStyle and use ListType only as a compatibility fallback.
            Try
                Dim levelNumber As System.Int32 =
                    System.Math.Max(1, listFormat.ListLevelNumber)
                Dim listTemplate As Microsoft.Office.Interop.Word.ListTemplate =
                    listFormat.ListTemplate
                If listTemplate IsNot Nothing Then
                    Dim numberStyle As Microsoft.Office.Interop.Word.WdListNumberStyle =
                        listTemplate.ListLevels(levelNumber).NumberStyle
                    Dim isBulletLevel As System.Boolean =
                        numberStyle = Microsoft.Office.Interop.Word.WdListNumberStyle.wdListNumberStyleBullet
                    Return If(isOrdered, Not isBulletLevel, isBulletLevel)
                End If
            Catch exLevel As System.Exception
                System.Diagnostics.Debug.WriteLine(
                    "HTML list reconciliation: active list-level type probe failed: " & exLevel.Message)
            End Try

            Dim isBulletType As System.Boolean =
                listType = Microsoft.Office.Interop.Word.WdListType.wdListBullet OrElse
                listType = Microsoft.Office.Interop.Word.WdListType.wdListPictureBullet

            Return If(isOrdered, Not isBulletType, isBulletType)
        End Function

        Private Shared Sub RemoveLiteralHtmlImporterListPrefix(
            ByVal paragraphRange As Microsoft.Office.Interop.Word.Range,
            ByVal token As System.String,
            ByVal isOrdered As System.Boolean
        )
            If paragraphRange Is Nothing OrElse System.String.IsNullOrEmpty(token) Then Return

            Dim tokenRange As Microsoft.Office.Interop.Word.Range =
                FindHtmlListImportTokenRange(paragraphRange, token)
            If tokenRange Is Nothing OrElse tokenRange.Start <= paragraphRange.Start Then Return

            Dim prefixRange As Microsoft.Office.Interop.Word.Range = paragraphRange.Duplicate()
            prefixRange.End = tokenRange.Start

            Dim prefixText As System.String = If(prefixRange.Text, System.String.Empty)
            prefixText = prefixText.Replace(Microsoft.VisualBasic.Strings.ChrW(160), " "c)

            Dim isImporterPrefix As System.Boolean
            If isOrdered Then
                isImporterPrefix = System.Text.RegularExpressions.Regex.IsMatch(
                    prefixText,
                    "^\s*\d+[\.\)]\s*$",
                    System.Text.RegularExpressions.RegexOptions.CultureInvariant)
            Else
                isImporterPrefix = System.Text.RegularExpressions.Regex.IsMatch(
                    prefixText,
                    "^\s*[\-\u2022\u00B7\u25CF\u25E6]\s*$",
                    System.Text.RegularExpressions.RegexOptions.CultureInvariant)
            End If

            If isImporterPrefix Then prefixRange.Delete()
        End Sub

        ''' <summary>
        ''' Repairs only list semantics that Word lost during CF_HTML import. Word's imported
        ''' paragraph indentation is captured first and restored after applying native list
        ''' formatting, so the HTML renderer remains the authority for visual nesting.
        ''' </summary>
        Private Shared Sub ReconcileHtmlListsAfterWordPaste(
            ByVal insertedRange As Microsoft.Office.Interop.Word.Range,
            ByVal descriptors As System.Collections.Generic.List(Of HtmlListImportDescriptor),
            ByVal destinationFontName As System.String,
            ByVal destinationFontSize As System.Single
        )
            If insertedRange Is Nothing OrElse descriptors Is Nothing OrElse descriptors.Count = 0 Then Return

            Dim targetsByContainer As New System.Collections.Generic.Dictionary(
                Of System.Int32,
                   System.Collections.Generic.List(Of HtmlListImportTarget))()

            Dim expectedByContainer As New System.Collections.Generic.Dictionary(Of System.Int32, System.Int32)()

            For Each descriptor As HtmlListImportDescriptor In descriptors
                If Not expectedByContainer.ContainsKey(descriptor.ContainerId) Then
                    expectedByContainer(descriptor.ContainerId) = 0
                End If
                expectedByContainer(descriptor.ContainerId) += 1

                Dim tokenRange As Microsoft.Office.Interop.Word.Range =
                    FindHtmlListImportTokenRange(insertedRange, descriptor.Token)
                If tokenRange Is Nothing OrElse tokenRange.Paragraphs Is Nothing OrElse tokenRange.Paragraphs.Count = 0 Then
                    System.Diagnostics.Debug.WriteLine(
                        "HTML list reconciliation: marker not found for container " &
                        descriptor.ContainerId.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                        ", item " &
                        descriptor.ItemIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) & ".")
                    Continue For
                End If

                Dim paragraphRange As Microsoft.Office.Interop.Word.Range =
                    tokenRange.Paragraphs(1).Range.Duplicate()

                Dim importedLeftIndent As System.Single = 0.0F
                Dim importedFirstLineIndent As System.Single = 0.0F
                Try
                    importedLeftIndent = paragraphRange.ParagraphFormat.LeftIndent
                    importedFirstLineIndent = paragraphRange.ParagraphFormat.FirstLineIndent
                Catch exIndent As System.Exception
                    System.Diagnostics.Debug.WriteLine(
                        "HTML list reconciliation: could not capture imported indentation: " & exIndent.Message)
                End Try

                Dim target As New HtmlListImportTarget With {
                    .Source = descriptor,
                    .ParagraphRange = paragraphRange,
                    .ImportedLeftIndent = importedLeftIndent,
                    .ImportedFirstLineIndent = importedFirstLineIndent
                }

                If Not targetsByContainer.ContainsKey(descriptor.ContainerId) Then
                    targetsByContainer(descriptor.ContainerId) =
                        New System.Collections.Generic.List(Of HtmlListImportTarget)()
                End If
                targetsByContainer(descriptor.ContainerId).Add(target)
            Next

            For Each pair As System.Collections.Generic.KeyValuePair(
                Of System.Int32,
                   System.Collections.Generic.List(Of HtmlListImportTarget)) In targetsByContainer

                Dim containerId As System.Int32 = pair.Key
                Dim targets As System.Collections.Generic.List(Of HtmlListImportTarget) = pair.Value

                Dim expectedCount As System.Int32 = 0
                If expectedByContainer.ContainsKey(containerId) Then
                    expectedCount = expectedByContainer(containerId)
                End If

                ' Never guess a partial source-to-target mapping.
                If targets.Count <> expectedCount Then
                    System.Diagnostics.Debug.WriteLine(
                        "HTML list reconciliation skipped for container " &
                        containerId.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                        ": expected " &
                        expectedCount.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                        " item marker(s), found " &
                        targets.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) & ".")
                    Continue For
                End If

                targets.Sort(
                    Function(leftTarget As HtmlListImportTarget, rightTarget As HtmlListImportTarget) As System.Int32
                        Return leftTarget.Source.ItemIndex.CompareTo(rightTarget.Source.ItemIndex)
                    End Function)

                Dim desiredOrdered As System.Boolean = targets(0).Source.IsOrdered
                Dim needsSemanticRepair As System.Boolean = False

                For Each target As HtmlListImportTarget In targets
                    Try
                        Dim expectedLevel As System.Int32 =
                            System.Math.Max(1, System.Math.Min(9, target.Source.Level))
                        Dim importedLevel As System.Int32 =
                            System.Math.Max(1, target.ParagraphRange.ListFormat.ListLevelNumber)

                        If Not WordListFormatMatchesHtmlKind(target.ParagraphRange.ListFormat, desiredOrdered) OrElse
                           importedLevel <> expectedLevel Then
                            needsSemanticRepair = True
                            System.Diagnostics.Debug.WriteLine(
                                "HTML list reconciliation: semantic mismatch for container " &
                                containerId.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                                ", item " & target.Source.ItemIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                                ": expected level=" & expectedLevel.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                                ", imported level=" & importedLevel.ToString(System.Globalization.CultureInfo.InvariantCulture) & ".")
                            Exit For
                        End If
                    Catch exListProbe As System.Exception
                        needsSemanticRepair = True
                        System.Diagnostics.Debug.WriteLine(
                            "HTML list reconciliation: list probe failed: " & exListProbe.Message)
                        Exit For
                    End Try
                Next

                If Not needsSemanticRepair Then Continue For

                ' Build one dedicated multi-level template for this source list container.
                ' This is deliberately created only when Word lost list semantics. It prevents
                ' us from mutating a gallery/document template that may be used elsewhere and
                ' lets the repaired paragraph expose the real source nesting level to later
                ' paragraph-format capture/restore code.
                Dim targetLevelNumber As System.Int32 =
                    System.Math.Max(1, System.Math.Min(9, targets(0).Source.Level))
                Dim repairTemplate As Microsoft.Office.Interop.Word.ListTemplate = Nothing

                Try
                    repairTemplate =
                        insertedRange.Document.ListTemplates.Add(OutlineNumbered:=True)

                    Dim repairLevel As Microsoft.Office.Interop.Word.ListLevel =
                        repairTemplate.ListLevels(targetLevelNumber)

                    If desiredOrdered Then
                        repairLevel.NumberStyle =
                            Microsoft.Office.Interop.Word.WdListNumberStyle.wdListNumberStyleArabic
                        repairLevel.NumberFormat =
                            "%" &
                            targetLevelNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                            "."
                        repairLevel.StartAt = System.Math.Max(1, targets(0).Source.StartAt)
                    Else
                        repairLevel.NumberStyle =
                            Microsoft.Office.Interop.Word.WdListNumberStyle.wdListNumberStyleBullet
                        repairLevel.NumberFormat = Microsoft.VisualBasic.Strings.ChrW(&H2022)
                    End If

                    ' Derive list-level geometry from Word's own HTML import before semantic
                    ' repair. For a hanging indent, LeftIndent is the text position and
                    ' LeftIndent + FirstLineIndent is the marker position.
                    Dim importedTextPosition As System.Single =
                        System.Math.Max(0.0F, targets(0).ImportedLeftIndent)
                    Dim importedMarkerPosition As System.Single =
                        System.Math.Max(
                            0.0F,
                            targets(0).ImportedLeftIndent + targets(0).ImportedFirstLineIndent)

                    If importedMarkerPosition >= importedTextPosition AndAlso importedTextPosition > 0.0F Then
                        importedMarkerPosition = System.Math.Max(0.0F, importedTextPosition - 18.0F)
                    End If

                    repairLevel.NumberPosition = importedMarkerPosition
                    repairLevel.TextPosition = importedTextPosition
                    repairLevel.TabPosition = importedTextPosition

                    ' The RILIST transport marker is NOT a formatting authority. Outlook can
                    ' import that artificial run using its own default font (for example Aptos),
                    ' even while InsertTextWithFormat correctly resolved Verdana from the real
                    ' destination text. Use that already-resolved destination base font instead.
                    ' Set all Word font-name slots because list markers and list text can use
                    ' different script slots in Outlook's WordEditor. No emphasis attributes are
                    ' touched here.
                    If Not System.String.IsNullOrWhiteSpace(destinationFontName) AndAlso
                       destinationFontName <> CStr(Microsoft.Office.Interop.Word.WdConstants.wdUndefined) Then
                        repairLevel.Font.Name = destinationFontName
                        repairLevel.Font.NameAscii = destinationFontName
                        repairLevel.Font.NameOther = destinationFontName
                        repairLevel.Font.NameFarEast = destinationFontName
                        repairLevel.Font.NameBi = destinationFontName
                    End If

                    If destinationFontSize > 0.0F AndAlso destinationFontSize < 1000.0F Then
                        repairLevel.Font.Size = destinationFontSize
                    End If

                Catch exTemplate As System.Exception
                    System.Diagnostics.Debug.WriteLine(
                        "HTML list reconciliation: could not create isolated repair template: " &
                        exTemplate.Message)
                    Continue For
                End Try

                ' If Word flattened even one item in this source list container, re-apply the
                ' entire container with the isolated template. This keeps ordered sequences
                ' continuous and gives every paragraph a true ListLevelNumber matching the HTML
                ' depth rather than a merely visual indentation.
                For targetIndex As System.Int32 = 0 To targets.Count - 1
                    Dim target As HtmlListImportTarget = targets(targetIndex)
                    Dim paragraphRange As Microsoft.Office.Interop.Word.Range = target.ParagraphRange

                    Try
                        Dim currentType As Microsoft.Office.Interop.Word.WdListType =
                            paragraphRange.ListFormat.ListType

                        If currentType = Microsoft.Office.Interop.Word.WdListType.wdListNoNumbering Then
                            ' Some Word HTML importers materialize an <ol> marker as literal "1. "
                            ' text. Because our token was the first source content, any such text can
                            ' only occur immediately before the token and can be removed without
                            ' inspecting or rewriting the user's actual list-item text.
                            RemoveLiteralHtmlImporterListPrefix(
                                paragraphRange,
                                target.Source.Token,
                                desiredOrdered)
                        Else
                            paragraphRange.ListFormat.RemoveNumbers(
                                Microsoft.Office.Interop.Word.WdNumberType.wdNumberParagraph)
                        End If

                        ' Apply the native Word list at the SOURCE level in the same COM call.
                        ' ApplyListTemplateWithLevel defaults to level 1 when ApplyLevel is omitted;
                        ' assigning ListLevelNumber afterwards is not equivalent and Word can retain
                        ' the paragraph as a flattened/non-native nested item. Existing Red Ink list
                        ' restoration code uses ApplyLevel for exactly this reason.
                        Dim applyLevel As System.Object = targetLevelNumber
                        paragraphRange.ListFormat.ApplyListTemplateWithLevel(
                            ListTemplate:=repairTemplate,
                            ContinuePreviousList:=(targetIndex > 0),
                            ApplyTo:=Microsoft.Office.Interop.Word.WdListApplyTo.wdListApplyToSelection,
                            DefaultListBehavior:=Microsoft.Office.Interop.Word.WdDefaultListBehavior.wdWord10ListBehavior,
                            ApplyLevel:=applyLevel)

                        ' Preserve the HTML importer's already-correct visual nesting as a second
                        ' invariant even after assigning the true Word list level.
                        paragraphRange.ParagraphFormat.LeftIndent = target.ImportedLeftIndent
                        paragraphRange.ParagraphFormat.FirstLineIndent = target.ImportedFirstLineIndent

                        ' ApplyListTemplateWithLevel can reset BOTH authorities independently in
                        ' Outlook: the ListLevel marker font and the paragraph text font. Reassert
                        ' the destination base family after the COM call as well. Only family/size
                        ' are restored, so Markdown Bold/Italic/Underline/Color runs remain intact.
                        Try
                            ' ApplyListTemplateWithLevel may cause WordEditor to attach an internal
                            ' copy of the template. Therefore also address the ACTUAL active level
                            ' obtained back from the paragraph after the COM call. This is the
                            ' authority that renders the visible bullet/number.
                            Dim activeListTemplate As Microsoft.Office.Interop.Word.ListTemplate =
                                paragraphRange.ListFormat.ListTemplate
                            Dim activeListLevelNumber As System.Int32 =
                                System.Math.Max(1, System.Math.Min(9, paragraphRange.ListFormat.ListLevelNumber))
                            If activeListTemplate IsNot Nothing Then
                                Dim activeListLevel As Microsoft.Office.Interop.Word.ListLevel =
                                    activeListTemplate.ListLevels(activeListLevelNumber)
                                If Not System.String.IsNullOrWhiteSpace(destinationFontName) AndAlso
                                   destinationFontName <> CStr(Microsoft.Office.Interop.Word.WdConstants.wdUndefined) Then
                                    activeListLevel.Font.Name = destinationFontName
                                    activeListLevel.Font.NameAscii = destinationFontName
                                    activeListLevel.Font.NameOther = destinationFontName
                                    activeListLevel.Font.NameFarEast = destinationFontName
                                    activeListLevel.Font.NameBi = destinationFontName
                                End If
                                If destinationFontSize > 0.0F AndAlso destinationFontSize < 1000.0F Then
                                    activeListLevel.Font.Size = destinationFontSize
                                End If
                            End If

                            Dim listTextRange As Microsoft.Office.Interop.Word.Range = paragraphRange.Duplicate()
                            If listTextRange.End > listTextRange.Start Then
                                Dim lastCharacterRange As Microsoft.Office.Interop.Word.Range = listTextRange.Duplicate()
                                lastCharacterRange.SetRange(listTextRange.End - 1, listTextRange.End)
                                Dim lastCharacter As System.String = lastCharacterRange.Text
                                If lastCharacter = vbCr OrElse lastCharacter = vbLf Then
                                    listTextRange.End -= 1
                                End If
                            End If

                            If listTextRange.End > listTextRange.Start Then
                                If Not System.String.IsNullOrWhiteSpace(destinationFontName) AndAlso
                                   destinationFontName <> CStr(Microsoft.Office.Interop.Word.WdConstants.wdUndefined) Then
                                    listTextRange.Font.Name = destinationFontName
                                    listTextRange.Font.NameAscii = destinationFontName
                                    listTextRange.Font.NameOther = destinationFontName
                                    listTextRange.Font.NameFarEast = destinationFontName
                                    listTextRange.Font.NameBi = destinationFontName
                                End If
                                If destinationFontSize > 0.0F AndAlso destinationFontSize < 1000.0F Then
                                    listTextRange.Font.Size = destinationFontSize
                                End If
                            End If

                        Catch exFontRestore As System.Exception
                            System.Diagnostics.Debug.WriteLine(
                                "HTML list reconciliation: could not restore destination base font after native list repair: " &
                                exFontRestore.Message)
                        End Try

                        If Not WordListFormatMatchesHtmlKind(paragraphRange.ListFormat, desiredOrdered) OrElse
                           paragraphRange.ListFormat.ListLevelNumber <> targetLevelNumber Then
                            Throw New System.Exception(
                                "Word did not retain the repaired native list semantics.")
                        End If

                    Catch exRepair As System.Exception
                        System.Diagnostics.Debug.WriteLine(
                            "HTML list reconciliation failed for container " &
                            containerId.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                            ", item " &
                            target.Source.ItemIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                            ": " & exRepair.Message)
                    End Try
                Next
            Next
        End Sub

        Private Shared Sub RemoveHtmlListImportMarkers(
            ByVal document As Microsoft.Office.Interop.Word.Document,
            ByVal rangeStart As System.Int32,
            ByVal rangeEnd As System.Int32,
            ByVal descriptors As System.Collections.Generic.List(Of HtmlListImportDescriptor)
        )
            If document Is Nothing OrElse descriptors Is Nothing OrElse descriptors.Count = 0 Then Return

            For Each descriptor As HtmlListImportDescriptor In descriptors
                Try
                    Dim safeEnd As System.Int32 =
                        System.Math.Min(document.Content.End, System.Math.Max(rangeStart, rangeEnd))
                    Dim searchStart As System.Object = rangeStart
                    Dim searchEnd As System.Object = safeEnd
                    Dim searchScope As Microsoft.Office.Interop.Word.Range =
                        document.Range(Start:=searchStart, End:=searchEnd)
                    Dim tokenRange As Microsoft.Office.Interop.Word.Range =
                        FindHtmlListImportTokenRange(searchScope, descriptor.Token)
                    If tokenRange IsNot Nothing Then tokenRange.Delete()
                Catch exCleanup As System.Exception
                    Throw New System.Exception(
                        "Internal HTML list marker cleanup failed.",
                        exCleanup)
                End Try
            Next
        End Sub


        Private Shared Function IsWordStructuralFormattingCharacter(ByVal character As System.Char) As System.Boolean
            Select Case Microsoft.VisualBasic.Strings.AscW(character)
                Case 7, 10, 11, 12, 13
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        Private Shared Function FindConcreteFormattingCharacter(
            ByVal candidateRange As Microsoft.Office.Interop.Word.Range,
            ByVal searchForward As System.Boolean
        ) As Microsoft.Office.Interop.Word.Range

            If candidateRange Is Nothing OrElse candidateRange.End <= candidateRange.Start Then Return Nothing

            Dim candidateText As System.String = If(candidateRange.Text, System.String.Empty)
            If candidateText.Length = 0 Then Return Nothing

            If searchForward Then
                For characterIndex As System.Int32 = 0 To candidateText.Length - 1
                    If IsWordStructuralFormattingCharacter(candidateText(characterIndex)) Then Continue For

                    Dim characterStart As System.Int32 = candidateRange.Start + characterIndex
                    If characterStart >= candidateRange.End Then Exit For
                    Dim result As Microsoft.Office.Interop.Word.Range = candidateRange.Duplicate()
                    result.SetRange(characterStart, System.Math.Min(characterStart + 1, candidateRange.End))
                    Return result
                Next
            Else
                For characterIndex As System.Int32 = candidateText.Length - 1 To 0 Step -1
                    If IsWordStructuralFormattingCharacter(candidateText(characterIndex)) Then Continue For

                    Dim characterStart As System.Int32 = candidateRange.Start + characterIndex
                    If characterStart >= candidateRange.End Then Continue For
                    Dim result As Microsoft.Office.Interop.Word.Range = candidateRange.Duplicate()
                    result.SetRange(characterStart, System.Math.Min(characterStart + 1, candidateRange.End))
                    Return result
                Next
            End If

            Return Nothing
        End Function

        ''' <summary>
        ''' Resolves a real text character for destination font inheritance. Paragraph marks,
        ''' table-cell end markers and line-break control characters are never used as the font
        ''' authority because Word/Outlook can expose host-default formatting on those markers.
        ''' </summary>
        Private Shared Function ResolveWordFormattingSourceRange(
            ByVal targetRange As Microsoft.Office.Interop.Word.Range,
            ByRef inheritInlineEmphasis As System.Boolean
        ) As Microsoft.Office.Interop.Word.Range

            inheritInlineEmphasis = True
            If targetRange Is Nothing Then Return Nothing

            If targetRange.End > targetRange.Start Then
                Dim insideTarget As Microsoft.Office.Interop.Word.Range =
                    FindConcreteFormattingCharacter(targetRange, searchForward:=True)
                If insideTarget IsNot Nothing Then Return insideTarget

                ' A non-empty selection containing only structural characters is typical for
                ' Outlook InsertAfter (the two temporary paragraph marks). In that case inherit
                ' the surrounding base font, but do not promote incidental bold/italic from the
                ' adjacent run to the entire inserted block.
                inheritInlineEmphasis = False
            End If

            Dim document As Microsoft.Office.Interop.Word.Document = targetRange.Document
            If document Is Nothing Then Return targetRange.Duplicate()

            Dim documentStart As System.Int32 = document.Content.Start
            Dim documentEnd As System.Int32 = document.Content.End
            Const probeWindow As System.Int32 = 4096

            If targetRange.Start > documentStart Then
                Dim beforeStart As System.Int32 = System.Math.Max(documentStart, targetRange.Start - probeWindow)
                Dim beforeRange As Microsoft.Office.Interop.Word.Range = document.Content.Duplicate()
                beforeRange.SetRange(beforeStart, targetRange.Start)
                Dim beforeCharacter As Microsoft.Office.Interop.Word.Range =
                    FindConcreteFormattingCharacter(beforeRange, searchForward:=False)
                If beforeCharacter IsNot Nothing Then Return beforeCharacter
            End If

            If targetRange.End < documentEnd Then
                Dim afterEnd As System.Int32 = System.Math.Min(documentEnd, targetRange.End + probeWindow)
                Dim afterRange As Microsoft.Office.Interop.Word.Range = document.Content.Duplicate()
                afterRange.SetRange(targetRange.End, afterEnd)
                Dim afterCharacter As Microsoft.Office.Interop.Word.Range =
                    FindConcreteFormattingCharacter(afterRange, searchForward:=True)
                If afterCharacter IsNot Nothing Then Return afterCharacter
            End If

            ' Controlled fallback: preserve the previous behavior only when no real text
            ' character exists nearby (for example in a completely empty document).
            Dim fallbackRange As Microsoft.Office.Interop.Word.Range = targetRange.Duplicate()
            If fallbackRange.Start = fallbackRange.End Then
                If fallbackRange.Start > documentStart Then
                    fallbackRange.SetRange(fallbackRange.Start - 1, fallbackRange.Start)
                ElseIf fallbackRange.End < documentEnd Then
                    fallbackRange.SetRange(fallbackRange.Start, fallbackRange.Start + 1)
                End If
            ElseIf fallbackRange.Start < fallbackRange.End Then
                fallbackRange.SetRange(fallbackRange.Start, fallbackRange.Start + 1)
            End If
            Return fallbackRange
        End Function

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
                                               Optional UseHostDefaultFontColor As Boolean = False,
                                               Optional PreserveDestinationParagraphFormatting As Boolean = False,
                                               Optional FormattingSourceRange As Microsoft.Office.Interop.Word.Range = Nothing)
            Try
                If formattedText Is Nothing OrElse formattedText.Trim() = "" Then
                    Return
                End If

                ' --- 0) Clone original range start and collapse to the start ---
                Dim origRange As Microsoft.Office.Interop.Word.Range = range.Duplicate()
                origRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseStart)

                System.Diagnostics.Debug.WriteLine("InsertTextWithFormat: UseHotDefaultFontColor=" & UseHostDefaultFontColor & "   NoTrailingCR=" & NoTrailingCR)
                System.Diagnostics.Debug.WriteLine("PreFinalHTML=" & formattedText)

                formattedText = FixMarkTagsForWord(formattedText)

                System.Diagnostics.Debug.WriteLine("PreFinalHTML[after-mark]=" & formattedText)

                ' --- 1) Load HTML and split <br> into separate <p> elements ---
                Dim doc As New HtmlAgilityPack.HtmlDocument()
                doc.LoadHtml(formattedText)
                NormalizeHtmlStrikethroughAndTaskCheckboxes(doc)

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
                            ' Preserve nested list structure. Markdig emits level-2+ Markdown lists as
                            ' child <ul>/<ol> nodes inside the parent <li>. The previous RemoveAllChildren()
                            ' path discarded those child lists whenever the parent item also contained a
                            ' <br>, flattening/truncating nested Markdown during Word insertion.
                            Dim nestedLists As New System.Collections.Generic.List(Of HtmlAgilityPack.HtmlNode)()
                            For Each child As HtmlAgilityPack.HtmlNode In node.ChildNodes.ToList()
                                If child.Name.Equals("ul", System.StringComparison.OrdinalIgnoreCase) OrElse
                                   child.Name.Equals("ol", System.StringComparison.OrdinalIgnoreCase) Then
                                    nestedLists.Add(child)
                                    node.RemoveChild(child)
                                End If
                            Next

                            Dim listItemSegments As String() =
                                System.Text.RegularExpressions.Regex.Split(
                                    node.InnerHtml,
                                    "<br\s*/?>",
                                    System.Text.RegularExpressions.RegexOptions.IgnoreCase)

                            node.RemoveAllChildren()
                            For Each seg As String In listItemSegments
                                Dim txt As String = seg.Trim()
                                If System.String.IsNullOrEmpty(txt) Then Continue For
                                Dim newP As HtmlAgilityPack.HtmlNode = doc.CreateElement("p")
                                newP.InnerHtml = txt
                                node.AppendChild(newP)
                            Next

                            ' Re-attach the exact nested list nodes after the parent item's own text.
                            ' This keeps existing simple-list behavior while retaining arbitrary deeper
                            ' list levels and their already-normalized descendants.
                            For Each nestedList As HtmlAgilityPack.HtmlNode In nestedLists
                                node.AppendChild(nestedList)
                            Next
                        End If
                    Next
                End If

                ' Add temporary anchors only after <br> normalization has finished, so the anchors
                ' cannot be removed by the list-item restructuring above. They are deleted again
                ' immediately after Word has imported the HTML.
                Dim listImportDescriptors As System.Collections.Generic.List(Of HtmlListImportDescriptor) =
                    PrepareHtmlListImportMarkers(doc)

                formattedText = doc.DocumentNode.OuterHtml

                ' --- 2) Read font and paragraph properties from a real text character. ---
                '     Structural characters such as paragraph marks are not formatting authorities:
                '     Outlook can expose the host default (for example Aptos) on a newly inserted
                '     paragraph mark even when the surrounding mail text is Verdana.
                Dim inheritInlineEmphasis As System.Boolean = True
                Dim fontSourceRange As Microsoft.Office.Interop.Word.Range = Nothing

                If FormattingSourceRange IsNot Nothing Then
                    ' A caller that creates a temporary insertion range can provide the actual
                    ' formatting authority from before that structural edit. This is important for
                    ' Outlook InsertAfter: searching around the temporary paragraph marks can reach
                    ' unrelated text below the insertion point and inherit its font/size.
                    fontSourceRange = FormattingSourceRange.Duplicate()
                    inheritInlineEmphasis = False
                Else
                    fontSourceRange = ResolveWordFormattingSourceRange(range, inheritInlineEmphasis)
                End If

                If fontSourceRange Is Nothing Then
                    fontSourceRange = range.Duplicate()
                End If

                Dim fontName As String = fontSourceRange.Font.Name
                Dim fontSize As Single = fontSourceRange.Font.Size
                Dim isBold As Boolean = inheritInlineEmphasis AndAlso (fontSourceRange.Font.Bold = 1)
                Dim isItalic As Boolean = inheritInlineEmphasis AndAlso (fontSourceRange.Font.Italic = 1)
                Dim fontColor As Integer = fontSourceRange.Font.Color
                Dim hexColor As String = String.Empty

                ' Guard against ambiguous values (9999999 = mixed formatting in selection).
                If fontSize <= 0 OrElse fontSize > 1000 Then fontSize = 11.0F
                If fontName Is Nothing OrElse fontName = "" Then fontName = "Calibri"

                Dim applyFontColor As System.Boolean =
                    Not UseHostDefaultFontColor AndAlso
                    fontColor <> CInt(Microsoft.Office.Interop.Word.WdConstants.wdUndefined)

                If applyFontColor Then
                    ' Convert Word BGR color to RGB hex string.
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
                If applyFontColor Then
                    cssBody &= $" color:{hexColor};"
                End If
                Dim cssPara As String = cssBody & $" font-size:{fontSize}pt; margin-top:{spaceBefore}pt; margin-bottom:{spaceAfter}pt;"
                If isBold Then cssPara &= " font-weight:bold;"
                If isItalic Then cssPara &= " font-style:italic;"

                ' --- 4) Apply inline styles ---
                Dim allTextContainers As HtmlAgilityPack.HtmlNodeCollection = doc.DocumentNode.SelectNodes("//p | //li")
                If allTextContainers IsNot Nothing AndAlso Not PreserveDestinationParagraphFormatting Then
                    For Each n As HtmlAgilityPack.HtmlNode In allTextContainers
                        n.SetAttributeValue("style", cssPara)
                    Next
                End If

                ' Headings (h1–h6): emit the target document's own heading Formatvorlagen
                ' (Überschrift 1..6 / Heading 1..6) font, size, weight and color, so pasted
                ' headings match what the user defined instead of the HTML importer's oversized
                ' browser-default heading sizing. Reads the built-in heading styles directly.
                Dim headings As HtmlAgilityPack.HtmlNodeCollection = doc.DocumentNode.SelectNodes("//h1 | //h2 | //h3 | //h4 | //h5 | //h6")
                If headings IsNot Nothing AndAlso Not PreserveDestinationParagraphFormatting Then
                    For Each h As HtmlAgilityPack.HtmlNode In headings
                        ' Resolve the built-in heading style for this level.
                        Dim headingLevel As Integer = 1
                        Integer.TryParse(h.Name.Substring(1), headingLevel)

                        Dim headingCss As String = String.Empty
                        Try
                            Dim builtin As Microsoft.Office.Interop.Word.WdBuiltinStyle
                            Select Case headingLevel
                                Case 1 : builtin = Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleHeading1
                                Case 2 : builtin = Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleHeading2
                                Case 3 : builtin = Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleHeading3
                                Case 4 : builtin = Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleHeading4
                                Case 5 : builtin = Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleHeading5
                                Case Else : builtin = Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleHeading6
                            End Select

                            Dim hStyle As Microsoft.Office.Interop.Word.Style =
                                CType(range.Document.Styles.Item(builtin), Microsoft.Office.Interop.Word.Style)
                            Dim hFont As Microsoft.Office.Interop.Word.Font = hStyle.Font

                            ' Font family (fall back to body family when undefined).
                            Dim hFontName As String = hFont.Name
                            If String.IsNullOrWhiteSpace(hFontName) OrElse
                               hFontName = CStr(Microsoft.Office.Interop.Word.WdConstants.wdUndefined) Then
                                hFontName = fontName
                            End If
                            headingCss &= $"font-family:'{hFontName}';"

                            ' Font size (only when the style defines a concrete size).
                            Dim hFontSize As Single = hFont.Size
                            If hFontSize > 0 AndAlso hFontSize < 1000 Then
                                headingCss &= $" font-size:{hFontSize.ToString(System.Globalization.CultureInfo.InvariantCulture)}pt;"
                            End If

                            ' Weight.
                            If hFont.Bold = -1 Then
                                headingCss &= " font-weight:bold;"
                            ElseIf hFont.Bold = 0 Then
                                headingCss &= " font-weight:normal;"
                            End If

                            ' Italic.
                            If hFont.Italic = -1 Then headingCss &= " font-style:italic;"

                            ' Color (skip when host default color is requested).
                            ' The built-in heading styles usually carry a *theme* color, so the
                            ' style's Font.Color only exposes a theme-index token (not a real RGB).
                            ' Resolve it against the document's theme color scheme and apply the
                            ' style's TintAndShade, mirroring how Word itself derives the shown RGB.
                            If Not UseHostDefaultFontColor Then
                                Dim headingColorHex As String = ResolveFontColorHex(hFont, range.Document)
                                If Not System.String.IsNullOrEmpty(headingColorHex) Then
                                    headingCss &= $" color:{headingColorHex};"
                                End If
                            End If
                        Catch ex As System.Exception
                            System.Diagnostics.Debug.WriteLine($"Heading style read failed (h{headingLevel}): {ex.Message}")
                            headingCss = String.Empty
                        End Try

                        ' Fall back to the previous behavior (body family/color/line-height)
                        ' when the heading style could not be read.
                        Dim merged As String
                        If System.String.IsNullOrWhiteSpace(headingCss) Then
                            merged = cssBody
                        Else
                            merged = $"line-height:{lineHeightCss}; {headingCss}"
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
                            Dim recoveryType As Microsoft.Office.Interop.Word.WdRecoveryType =
                                If(PreserveDestinationParagraphFormatting,
                                   Microsoft.Office.Interop.Word.WdRecoveryType.wdUseDestinationStylesRecovery,
                                   Microsoft.Office.Interop.Word.WdRecoveryType.wdFormatOriginalFormatting)
                            If ReplaceSelection Then
                                range.Application.Selection.PasteAndFormat(recoveryType)
                            Else
                                range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
                                range.Select()
                                range.Application.Selection.PasteAndFormat(recoveryType)
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

                    ' --- 7a) Reconcile HTML list semantics before any caller can restore
                    '     destination paragraph formatting. Word sometimes imports nested <ol>/<ul>
                    '     items as visually indented ordinary paragraphs with literal markers rather
                    '     than native Word lists. The source anchors map those paragraphs exactly;
                    '     native list semantics are repaired while Word's imported indentation is
                    '     preserved as the visual authority.
                    Dim insertedStartForLists As System.Int32 = origRange.Start
                    Try
                        If listImportDescriptors IsNot Nothing AndAlso listImportDescriptors.Count > 0 Then
                            Dim insertedEndForLists As System.Int32 =
                                range.Application.Selection.Range.End
                            Dim insertedListStart As System.Object = insertedStartForLists
                            Dim insertedListEnd As System.Object = insertedEndForLists
                            Dim insertedListRange As Microsoft.Office.Interop.Word.Range =
                                range.Document.Range(Start:=insertedListStart, End:=insertedListEnd)

                            ReconcileHtmlListsAfterWordPaste(
                                insertedListRange,
                                listImportDescriptors,
                                fontName,
                                fontSize)
                        End If
                    Catch exListRepair As System.Exception
                        System.Diagnostics.Debug.WriteLine(
                            "HTML list reconciliation skipped: " & exListRepair.Message)
                    Finally
                        ' Internal anchors are transport-only and must never survive in the document.
                        If listImportDescriptors IsNot Nothing AndAlso listImportDescriptors.Count > 0 Then
                            RemoveHtmlListImportMarkers(
                                range.Document,
                                insertedStartForLists,
                                range.Application.Selection.Range.End,
                                listImportDescriptors)
                            range = range.Application.Selection.Range
                        End If
                    End Try

                    ' --- 7b) Strip automatic heading numbering that the pasted heading
                    '     Formatvorlagen (Überschrift 1..6) may carry, when the source markdown
                    '     contained no such numbers. Genuine numbered lists (<ol>/<li>) are
                    '     body-text paragraphs and keep their numbering; only paragraphs whose
                    '     outline level is an actual heading level are cleared. This is fully
                    '     deterministic and locale-independent (no style-name matching).
                    Try
                        Dim insertedStart As Object = origRange.Start
                        Dim insertedEnd As Object = range.Application.Selection.Range.End
                        Dim insertedRng As Microsoft.Office.Interop.Word.Range =
                            range.Document.Range(insertedStart, insertedEnd)

                        For Each hpara As Microsoft.Office.Interop.Word.Paragraph In insertedRng.Paragraphs
                            Try
                                Dim isHeadingLevel As Boolean =
                                    (hpara.OutlineLevel <> Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText)

                                Dim hasAutoNumbering As Boolean =
                                    (hpara.Range.ListFormat.ListType <> Microsoft.Office.Interop.Word.WdListType.wdListNoNumbering)

                                If isHeadingLevel AndAlso hasAutoNumbering Then
                                    hpara.Range.ListFormat.RemoveNumbers(
                                        Microsoft.Office.Interop.Word.WdNumberType.wdNumberParagraph)
                                End If
                            Catch exPara As System.Exception
                                System.Diagnostics.Debug.WriteLine($"Heading numbering strip (paragraph) skipped: {exPara.Message}")
                            End Try
                        Next
                    Catch exNum As System.Exception
                        System.Diagnostics.Debug.WriteLine($"Heading numbering strip skipped: {exNum.Message}")
                    End Try

                    ' --- 7c) Reset the trailing insertion point back to the captured body font. ---
                    '     PasteAndFormat leaves the paragraph mark at the end of the inserted
                    '     content carrying the formatting of the last pasted run (e.g. a heading).
                    '     A subsequent insertion reads its font from exactly this position, so it
                    '     would otherwise inherit that stale formatting. Collapse to the end and
                    '     restore the body font so the next round matches again.
                    Try
                        Dim tailRange As Microsoft.Office.Interop.Word.Range = range.Application.Selection.Range.Duplicate()
                        tailRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
                        tailRange.Font.Name = fontName
                        tailRange.Font.Size = fontSize
                        If Not UseHostDefaultFontColor Then
                            tailRange.Font.Color = CType(fontColor, Microsoft.Office.Interop.Word.WdColor)
                        End If
                    Catch exTail As System.Exception
                        System.Diagnostics.Debug.WriteLine($"Trailing font reset skipped: {exTail.Message}")
                    End Try

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
                If PreserveDestinationParagraphFormatting Then Throw
                Global.SharedLibrary.SharedLibrary.SharedMethods.ShowCustomMessageBox("InsertTextWithFormat Error: " & ex.Message)
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
        ''' Resolves the concrete RGB of a <see cref="Microsoft.Office.Interop.Word.Font"/> to a
        ''' CSS hex string, handling both direct RGB colors and theme colors. Theme colors are
        ''' resolved against the document's theme color scheme and adjusted by the font's
        ''' <c>TintAndShade</c>, matching how Word derives the displayed color. Returns an empty
        ''' string when the color is automatic/undefined so the caller can fall back to the host
        ''' default color.
        ''' </summary>
        ''' <param name="font">The font whose color should be resolved.</param>
        ''' <param name="document">The document providing the theme color scheme.</param>
        ''' <returns>A CSS hex color string (for example, <c>#0F4761</c>) or an empty string.</returns>
        Private Shared Function ResolveFontColorHex(font As Microsoft.Office.Interop.Word.Font,
                                                    document As Microsoft.Office.Interop.Word.Document) As String
            Try
                Dim themeColor As Microsoft.Office.Interop.Word.WdThemeColorIndex = font.TextColor.ObjectThemeColor

                Dim r As Integer
                Dim g As Integer
                Dim b As Integer

                If themeColor = Microsoft.Office.Interop.Word.WdThemeColorIndex.wdNotThemeColor Then
                    ' Direct (non-theme) color: TextColor.RGB already holds the concrete RGB.
                    Dim rgb As Integer = font.TextColor.RGB
                    If rgb = CInt(Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic) OrElse
                       rgb = CInt(Microsoft.Office.Interop.Word.WdConstants.wdUndefined) OrElse
                       rgb < 0 Then
                        Return System.String.Empty
                    End If
                    r = (rgb And &HFF)
                    g = ((rgb >> 8) And &HFF)
                    b = ((rgb >> 16) And &HFF)
                Else
                    ' Theme color: look up the base RGB in the document's theme color scheme.
                    Dim schemeIndex As Microsoft.Office.Core.MsoThemeColorSchemeIndex =
                        MapThemeColorToSchemeIndex(themeColor)
                    If schemeIndex = CType(0, Microsoft.Office.Core.MsoThemeColorSchemeIndex) Then
                        Return System.String.Empty
                    End If

                    Dim baseRgb As Integer =
                        document.DocumentTheme.ThemeColorScheme.Colors(schemeIndex).RGB

                    ' Theme scheme RGB is stored as R + G*256 + B*65536 (same layout as Word RGB).
                    r = (baseRgb And &HFF)
                    g = ((baseRgb >> 8) And &HFF)
                    b = ((baseRgb >> 16) And &HFF)

                    ' Apply the tint/shade the style requests (HSL luminance adjustment).
                    ApplyTintAndShade(r, g, b, font.TextColor.TintAndShade)
                End If

                Return System.String.Format("#{0:X2}{1:X2}{2:X2}", r, g, b)

            Catch ex As System.Exception
                System.Diagnostics.Debug.WriteLine($"ResolveFontColorHex failed: {ex.Message}")
                Return System.String.Empty
            End Try
        End Function

        ''' <summary>
        ''' Maps a Word <see cref="Microsoft.Office.Interop.Word.WdThemeColorIndex"/> to the
        ''' corresponding <see cref="Microsoft.Office.Core.MsoThemeColorSchemeIndex"/> used by the
        ''' document theme color scheme.
        ''' </summary>
        ''' <param name="themeColor">The Word theme color index.</param>
        ''' <returns>The matching Office theme color scheme index, or 0 when unmapped.</returns>
        Private Shared Function MapThemeColorToSchemeIndex(
            themeColor As Microsoft.Office.Interop.Word.WdThemeColorIndex) As Microsoft.Office.Core.MsoThemeColorSchemeIndex

            Select Case themeColor
                Case Microsoft.Office.Interop.Word.WdThemeColorIndex.wdThemeColorMainDark1,
                     Microsoft.Office.Interop.Word.WdThemeColorIndex.wdThemeColorText1
                    Return Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeDark1
                Case Microsoft.Office.Interop.Word.WdThemeColorIndex.wdThemeColorMainLight1,
                     Microsoft.Office.Interop.Word.WdThemeColorIndex.wdThemeColorBackground1
                    Return Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeLight1
                Case Microsoft.Office.Interop.Word.WdThemeColorIndex.wdThemeColorMainDark2,
                     Microsoft.Office.Interop.Word.WdThemeColorIndex.wdThemeColorText2
                    Return Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeDark2
                Case Microsoft.Office.Interop.Word.WdThemeColorIndex.wdThemeColorMainLight2,
                     Microsoft.Office.Interop.Word.WdThemeColorIndex.wdThemeColorBackground2
                    Return Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeLight2
                Case Microsoft.Office.Interop.Word.WdThemeColorIndex.wdThemeColorAccent1
                    Return Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeAccent1
                Case Microsoft.Office.Interop.Word.WdThemeColorIndex.wdThemeColorAccent2
                    Return Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeAccent2
                Case Microsoft.Office.Interop.Word.WdThemeColorIndex.wdThemeColorAccent3
                    Return Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeAccent3
                Case Microsoft.Office.Interop.Word.WdThemeColorIndex.wdThemeColorAccent4
                    Return Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeAccent4
                Case Microsoft.Office.Interop.Word.WdThemeColorIndex.wdThemeColorAccent5
                    Return Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeAccent5
                Case Microsoft.Office.Interop.Word.WdThemeColorIndex.wdThemeColorAccent6
                    Return Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeAccent6
                Case Microsoft.Office.Interop.Word.WdThemeColorIndex.wdThemeColorHyperlink
                    Return Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeHyperlink
                Case Microsoft.Office.Interop.Word.WdThemeColorIndex.wdThemeColorHyperlinkFollowed
                    Return Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeFollowedHyperlink
                Case Else
                    Return CType(0, Microsoft.Office.Core.MsoThemeColorSchemeIndex)
            End Select
        End Function

        ''' <summary>
        ''' Applies Word's <c>TintAndShade</c> to an RGB triple by adjusting the HSL luminance,
        ''' mirroring the transformation Word uses to derive lighter/darker theme color variants.
        ''' </summary>
        ''' <param name="r">Red channel (0-255), modified in place.</param>
        ''' <param name="g">Green channel (0-255), modified in place.</param>
        ''' <param name="b">Blue channel (0-255), modified in place.</param>
        ''' <param name="tintAndShade">The Word TintAndShade value (-1.0 to 1.0).</param>
        Private Shared Sub ApplyTintAndShade(ByRef r As Integer, ByRef g As Integer, ByRef b As Integer, tintAndShade As Single)
            If tintAndShade = 0.0F Then Return

            Dim rd As Double = r / 255.0
            Dim gd As Double = g / 255.0
            Dim bd As Double = b / 255.0

            Dim maxC As Double = System.Math.Max(rd, System.Math.Max(gd, bd))
            Dim minC As Double = System.Math.Min(rd, System.Math.Min(gd, bd))

            Dim h As Double = 0.0
            Dim s As Double = 0.0
            Dim l As Double = (maxC + minC) / 2.0

            If maxC <> minC Then
                Dim d As Double = maxC - minC
                s = If(l > 0.5, d / (2.0 - maxC - minC), d / (maxC + minC))
                If maxC = rd Then
                    h = (gd - bd) / d + (If(gd < bd, 6.0, 0.0))
                ElseIf maxC = gd Then
                    h = (bd - rd) / d + 2.0
                Else
                    h = (rd - gd) / d + 4.0
                End If
                h /= 6.0
            End If

            ' TintAndShade > 0 lightens toward white, < 0 darkens toward black.
            If tintAndShade > 0.0F Then
                l = l * (1.0 - tintAndShade) + tintAndShade
            Else
                l = l * (1.0 + tintAndShade)
            End If

            Dim r2 As Double
            Dim g2 As Double
            Dim b2 As Double

            If s = 0.0 Then
                r2 = l : g2 = l : b2 = l
            Else
                Dim q As Double = If(l < 0.5, l * (1.0 + s), l + s - l * s)
                Dim p As Double = 2.0 * l - q
                r2 = HueToRgb(p, q, h + 1.0 / 3.0)
                g2 = HueToRgb(p, q, h)
                b2 = HueToRgb(p, q, h - 1.0 / 3.0)
            End If

            r = CInt(System.Math.Round(System.Math.Max(0.0, System.Math.Min(1.0, r2)) * 255.0))
            g = CInt(System.Math.Round(System.Math.Max(0.0, System.Math.Min(1.0, g2)) * 255.0))
            b = CInt(System.Math.Round(System.Math.Max(0.0, System.Math.Min(1.0, b2)) * 255.0))
        End Sub

        ''' <summary>
        ''' Helper for HSL-to-RGB conversion (single channel).
        ''' </summary>
        Private Shared Function HueToRgb(p As Double, q As Double, t As Double) As Double
            If t < 0.0 Then t += 1.0
            If t > 1.0 Then t -= 1.0
            If t < 1.0 / 6.0 Then Return p + (q - p) * 6.0 * t
            If t < 1.0 / 2.0 Then Return q
            If t < 2.0 / 3.0 Then Return p + (q - p) * (2.0 / 3.0 - t) * 6.0
            Return p
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
                Global.SharedLibrary.SharedLibrary.SharedMethods.ShowCustomMessageBox("RemoveTrailingCr Error: " & ex.Message)
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


        ''' <summary>
        ''' Resets paragraph spacing in the current active Word selection to single line spacing
        ''' with no space before or after paragraphs, leaving all other formatting unchanged.
        ''' If there is no text selection, the entire current story/document is selected first.
        ''' Works both in Word and in an Outlook mail editor that uses WordEditor.
        ''' </summary>
        Public Shared Sub ResetSelectedTextParagraphSpacing()
            Dim selection As Microsoft.Office.Interop.Word.Selection = TryGetActiveWordSelection()

            If selection Is Nothing Then
                selection = TryGetActiveOutlookEditorSelection()
            End If

            ResetSelectedTextParagraphSpacing(selection)
        End Sub

        ''' <summary>
        ''' Resets paragraph spacing for the supplied Word selection. The host passes its
        ''' actual selection explicitly so Word and Outlook WordEditor ranges cannot be mixed.
        ''' </summary>
        Public Shared Sub ResetSelectedTextParagraphSpacing(ByVal selection As Microsoft.Office.Interop.Word.Selection)
            Try
                If selection Is Nothing OrElse selection.Range Is Nothing Then
                    Return
                End If

                If selection.Start = selection.End Then
                    selection.WholeStory()
                End If

                If selection.Range Is Nothing OrElse selection.Range.Start = selection.Range.End Then
                    Return
                End If

                Dim paragraphFormat As Microsoft.Office.Interop.Word.ParagraphFormat = selection.Range.ParagraphFormat
                paragraphFormat.SpaceBeforeAuto = 0
                paragraphFormat.SpaceAfterAuto = 0
                paragraphFormat.SpaceBefore = 0.0F
                paragraphFormat.SpaceAfter = 0.0F
                paragraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceSingle
            Catch ex As System.Exception
                ShowCustomMessageBox("ResetSelectedTextParagraphSpacing Error: " & ex.Message, "Error")
            End Try
        End Sub
        ''' <summary>
        ''' Returns the active selection from an Outlook compose inspector's WordEditor, if available.
        ''' Uses reflection so SharedLibrary does not take a direct dependency on Outlook interop.
        ''' </summary>
        Private Shared Function TryGetActiveOutlookEditorSelection() As Microsoft.Office.Interop.Word.Selection
            Try
                Dim outlookAppObj As Object = Nothing

                Try
                    outlookAppObj = System.Runtime.InteropServices.Marshal.GetActiveObject("Outlook.Application")
                Catch
                    outlookAppObj = Nothing
                End Try

                If outlookAppObj Is Nothing Then
                    Return Nothing
                End If

                Dim inspectorObj As Object = Nothing
                Try
                    inspectorObj = outlookAppObj.GetType().InvokeMember(
                        "ActiveInspector",
                        System.Reflection.BindingFlags.InvokeMethod Or
                        System.Reflection.BindingFlags.Public Or
                        System.Reflection.BindingFlags.Instance,
                        Nothing,
                        outlookAppObj,
                        Nothing,
                        System.Globalization.CultureInfo.InvariantCulture)
                Catch
                    inspectorObj = Nothing
                End Try

                If inspectorObj Is Nothing Then
                    Return Nothing
                End If

                Dim wordEditorObj As Object = Nothing
                Try
                    wordEditorObj = inspectorObj.GetType().InvokeMember(
                        "WordEditor",
                        System.Reflection.BindingFlags.GetProperty Or
                        System.Reflection.BindingFlags.Public Or
                        System.Reflection.BindingFlags.Instance,
                        Nothing,
                        inspectorObj,
                        Nothing,
                        System.Globalization.CultureInfo.InvariantCulture)
                Catch
                    wordEditorObj = Nothing
                End Try

                Dim wordDoc As Microsoft.Office.Interop.Word.Document =
                    TryCast(wordEditorObj, Microsoft.Office.Interop.Word.Document)

                If wordDoc Is Nothing Then
                    Return Nothing
                End If

                Dim selection As Microsoft.Office.Interop.Word.Selection = Nothing
                Try
                    selection = wordDoc.Application.Selection
                Catch
                    selection = Nothing
                End Try

                If selection Is Nothing OrElse selection.Range Is Nothing Then
                    Return Nothing
                End If

                Return selection

            Catch ex As System.Exception
                System.Diagnostics.Debug.WriteLine("TryGetActiveOutlookEditorSelection Error: " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Returns the active selection from the running Word application, if available.
        ''' </summary>
        Private Shared Function TryGetActiveWordSelection() As Microsoft.Office.Interop.Word.Selection
            Try
                Dim wordAppObj As Object = Nothing

                Try
                    wordAppObj = System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application")
                Catch
                    wordAppObj = Nothing
                End Try

                Dim wordApp As Microsoft.Office.Interop.Word.Application =
                    TryCast(wordAppObj, Microsoft.Office.Interop.Word.Application)

                If wordApp Is Nothing OrElse wordApp.Selection Is Nothing OrElse wordApp.Selection.Range Is Nothing Then
                    Return Nothing
                End If

                Return wordApp.Selection

            Catch ex As System.Exception
                System.Diagnostics.Debug.WriteLine("TryGetActiveWordSelection Error: " & ex.Message)
                Return Nothing
            End Try
        End Function

    End Class
End Namespace
