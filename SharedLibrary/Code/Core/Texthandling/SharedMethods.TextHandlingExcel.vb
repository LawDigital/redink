' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: SharedMethods.TextHandlingExcel.vb
' Purpose: Provides helper methods for cleaning and decoding text that appears inside
'          Excel formulas and Excel-style string literals, including decoding of JSON
'          and HTML escapes and limited normalization/removal of RTF constructs.
'
' Architecture:
'  - Formula Scanner: `CleanExcelFormulaStrings` scans an Excel formula and applies
'    decoding/cleanup only inside Excel string literals (quoted segments), while optionally
'    stripping outer CR/LF characters outside literals.
'  - Literal Decoding: `DecodeTextLiterals` applies a decoding pipeline for malformed
'    decimal `\u###?` escapes, JSON unescaping (when backslashes are present), HTML entity
'    decoding, and conversion of some RTF escape sequences to plain text.
'  - RTF Hyperlink Normalization: `NormalizeRtfHyperlinks` replaces brace-balanced RTF
'    hyperlink field blocks with a plain-text representation: `url` or `url (display)`.
'  - RTF Cleanup: `StripRtfInline`, `CleanupRtfRemnants`, and `_CleanRtfLiteral` remove
'    selected RTF tokens and decode common RTF escape sequences.
'
' External Dependencies:
'  - `Newtonsoft.Json`: Used to deserialize a quoted string to perform JSON unescaping.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Text.RegularExpressions
Imports System.Text
Imports System.Globalization

Namespace SharedLibrary
    Partial Public Class SharedMethods

        ''' <summary>
        ''' Cleans and decodes string literals inside an Excel formula.
        ''' </summary>
        ''' <param name="formula">
        ''' The Excel formula text. If it does not start with <c>=</c>, the method treats it as plain text and calls
        ''' <see cref="DecodeTextLiterals(String)"/>.
        ''' </param>
        ''' <param name="decodeHtml">If <c>True</c>, HTML entities (for example, <c>&amp;auml;</c>) are decoded inside string literals.</param>
        ''' <param name="decodeUnicode">If <c>True</c>, selected Unicode escape sequences are decoded inside string literals.</param>
        ''' <param name="cleanRtf">If <c>True</c>, applies a lightweight RTF cleanup inside string literals.</param>
        ''' <param name="stripOuterLineBreaks">If <c>True</c>, CR/LF characters outside string literals are removed.</param>
        ''' <returns>The cleaned formula text.</returns>
        Public Shared Function CleanExcelFormulaStrings(formula As String,
                                                        Optional decodeHtml As Boolean = True,
                                                        Optional decodeUnicode As Boolean = True,
                                                        Optional cleanRtf As Boolean = True,
                                                        Optional stripOuterLineBreaks As Boolean = True) As String
            If String.IsNullOrEmpty(formula) OrElse formula(0) <> "="c Then
                Return DecodeTextLiterals(formula)
            End If

            Dim sb As New System.Text.StringBuilder(formula.Length)
            Dim i As Integer = 0
            While i < formula.Length
                Dim ch As Char = formula(i)
                If ch = """"c Then
                    ' Enter string literal
                    Dim j As Integer = i + 1
                    While j < formula.Length
                        If formula(j) = """"c Then
                            If j + 1 < formula.Length AndAlso formula(j + 1) = """"c Then
                                j += 2
                                Continue While
                            Else
                                Exit While
                            End If
                        End If
                        j += 1
                    End While
                    If j >= formula.Length Then Return formula ' Unbalanced -> return original

                    Dim encoded = formula.Substring(i + 1, j - i - 1)
                    Dim plain = encoded.Replace("""""", """")
                    Dim cleaned = plain

                    If decodeUnicode Then
                        cleaned = _rtfUnicodePattern.Replace(cleaned, _rtfUnicodeEvaluator)
                        cleaned = _jsonUnicodePattern.Replace(cleaned, _jsonUnicodeEvaluator)
                    End If
                    If decodeHtml Then
                        cleaned = System.Net.WebUtility.HtmlDecode(cleaned)
                    End If
                    If cleanRtf Then
                        cleaned = _CleanRtfLiteral(cleaned)
                    End If

                    Dim reEscaped = cleaned.Replace("""", """""")
                    sb.Append(""""c).Append(reEscaped).Append(""""c)
                    i = j + 1
                Else
                    ' Outside string literal
                    If stripOuterLineBreaks AndAlso (ch = ChrW(10) OrElse ch = ChrW(13)) Then
                        ' Skip CR / LF (and paired CRLF / LFCR)
                        i += 1
                        If i < formula.Length AndAlso ((ch = ChrW(13) AndAlso formula(i) = ChrW(10)) OrElse (ch = ChrW(10) AndAlso formula(i) = ChrW(13))) Then
                            i += 1
                        End If
                        Continue While
                    End If
                    sb.Append(ch)
                    i += 1
                End If
            End While
            Return sb.ToString()
        End Function

        ''' <summary>
        ''' Matches RTF decimal Unicode escapes of the form <c>\uN?</c> (including negative values).
        ''' </summary>
        Private Shared ReadOnly _rtfUnicodePattern As New Regex("\\u(-?\d+)\?", RegexOptions.Compiled)

        ''' <summary>
        ''' Matches JSON-style hex Unicode escapes of the form <c>\uXXXX</c>.
        ''' </summary>
        Private Shared ReadOnly _jsonUnicodePattern As New Regex("\\u([0-9A-Fa-f]{4})", RegexOptions.Compiled)


        ' Decode JSON/HTML escapes and normalize RTF hyperlink fields to "URL (display)".

        ''' <summary>
        ''' Decodes a text string by applying JSON unescaping, HTML entity decoding, and selected RTF normalization steps.
        ''' </summary>
        ''' <param name="text">The input text.</param>
        ''' <returns>The decoded text.</returns>
        Public Shared Function DecodeTextLiterals(ByVal text As String) As String
            If String.IsNullOrEmpty(text) Then Return text

            Dim result As String = text

            ' 1) Fix malformed decimal escape pattern: \u228?  (228 decimal => ä)
            Dim decFix As MatchEvaluator = Function(m As Match)
                                               Dim dec As Integer = Integer.Parse(m.Groups(1).Value)
                                               Return ChrW(dec).ToString()
                                           End Function
            result = Regex.Replace(result, "\\u(\d{1,3})\?", decFix)

            ' 2) Try to JSON-unescape if it contains backslashes (covers \uXXXX, \n, \t, \", etc.)
            If result.IndexOf("\"c) >= 0 Then
                Try
                    Dim quoted As String = ChrW(34) & result.Replace(ChrW(34), "\" & ChrW(34)) & ChrW(34)
                    result = Newtonsoft.Json.JsonConvert.DeserializeObject(Of String)(quoted)
                Catch
                End Try
            End If

            ' 3) HTML-decode entities (e.g., &auml;, &#228;)
            Try
                result = System.Net.WebUtility.HtmlDecode(result)
            Catch
            End Try

            ' 4) Convert RTF \line to newline
            result = Regex.Replace(result, "\\line\b\s*", Environment.NewLine, RegexOptions.IgnoreCase)

            ' 5) RTF decimal \u####? → Unicode char
            result = _rtfUnicodePattern.Replace(result, Function(m)
                                                            Dim dec = Integer.Parse(m.Groups(1).Value)
                                                            Return ChrW(dec)
                                                        End Function)

            ' 6) JSON hex \uXXXX → Unicode char (avoid double-decoding already converted)
            result = _jsonUnicodePattern.Replace(
                                    result,
                                    Function(m As Match)
                                        Dim hex As String = m.Groups(1).Value
                                        Dim code As Integer = Integer.Parse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture)
                                        Return ChrW(code)
                                    End Function)

            ' 7) Normalize RTF hyperlink fields (brace-aware)
            result = NormalizeRtfHyperlinks(result)

            ' 8) Cleanup for leftover RTF remnants ({{...}} and font/format switches like \f0 \fs20)
            result = CleanupRtfRemnants(result)

            Return result
        End Function

        ' Convert {\field ...{\*\fldinst ... HYPERLINK "url"...}{\fldrslt ...display...}} to "url (display)"

        ''' <summary>
        ''' Converts RTF <c>\field</c> hyperlink blocks into a plain-text representation: <c>url</c> or <c>url (display)</c>.
        ''' </summary>
        ''' <param name="s">The input text that may contain RTF hyperlink fields.</param>
        ''' <returns>The input with recognized RTF hyperlink fields normalized.</returns>
        Private Shared Function NormalizeRtfHyperlinks(ByVal s As String) As String
            If String.IsNullOrEmpty(s) Then Return s

            Dim i As Integer = 0
            Dim sb As New StringBuilder(s.Length)

            While i < s.Length
                Dim j As Integer = s.IndexOf("\field", i, StringComparison.OrdinalIgnoreCase)
                If j = -1 Then
                    sb.Append(s, i, s.Length - i)
                    Exit While
                End If

                ' Append text before the field
                sb.Append(s, i, j - i)

                ' Include all consecutive opening braces before \field (handles "{{\field ...}}")
                Dim fieldStart As Integer = j
                While fieldStart > 0 AndAlso s.Chars(fieldStart - 1) = "{"c
                    fieldStart -= 1
                End While

                Dim fieldEnd As Integer = FindMatchingBrace(s, fieldStart)
                If fieldEnd = -1 Then
                    ' Fallback: not a full RTF block; copy as-is and stop
                    sb.Append(s, fieldStart, s.Length - fieldStart)
                    Exit While
                End If

                Dim fieldBlock As String = s.Substring(fieldStart, fieldEnd - fieldStart + 1)

                ' Extract URL after HYPERLINK "..."
                Dim url As String = ExtractQuotedAfter(fieldBlock, "HYPERLINK")
                ' Extract fldrslt content block
                Dim displayRaw As String = ExtractFldrsltBlock(fieldBlock)
                Dim display As String = StripRtfInline(displayRaw)

                If Not String.IsNullOrWhiteSpace(url) Then
                    If String.IsNullOrWhiteSpace(display) Then
                        sb.Append(url)
                    Else
                        sb.Append(url).Append(" (").Append(display).Append(")")
                    End If
                Else
                    ' If no URL found, keep original block
                    sb.Append(fieldBlock)
                End If

                i = fieldEnd + 1
            End While

            Return sb.ToString()
        End Function

        ' Strip simple inline RTF (e.g., \ul, \cf1, \fs20, \b) and decode hex escapes like \'e4; drop braces.

        ''' <summary>
        ''' Removes selected inline RTF control words, decodes <c>\'xx</c> hex escapes using Windows-1252,
        ''' drops braces, and normalizes whitespace.
        ''' </summary>
        ''' <param name="s">The input text.</param>
        ''' <returns>The stripped text.</returns>
        Private Shared Function StripRtfInline(ByVal s As String) As String
            If String.IsNullOrEmpty(s) Then Return s

            ' Decode hex escapes \'xx to Unicode
            Dim hexDecoder As MatchEvaluator = Function(m As Match)
                                                   Dim hex = m.Value.Substring(2)
                                                   Dim b As Byte = System.Convert.ToByte(hex, 16)
                                                   Return Encoding.GetEncoding(1252).GetString(New Byte() {b})
                                               End Function
            s = Regex.Replace(s, "\\'[0-9a-fA-F]{2}", hexDecoder)

            ' Remove control words (e.g., \ul, \ul0, \cf1, \fs20, \b, \f0, etc.)
            s = Regex.Replace(s, "\\[a-zA-Z]+-?\d*", "")

            ' Remove remaining braces
            s = s.Replace("{", "").Replace("}", "")

            ' Collapse whitespace
            s = Regex.Replace(s, "\s+", " ").Trim()

            Return s
        End Function

        ' Remove leftover RTF tokens and braces around normalized hyperlinks, e.g. "{{URL (title)}}" and "\f0\fs20".

        ''' <summary>
        ''' Removes selected RTF remnants (for example, <c>\par</c>, <c>\tab</c>, <c>\fN</c>, <c>\fsN</c>) and
        ''' normalizes whitespace and line breaks.
        ''' </summary>
        ''' <param name="s">The input text.</param>
        ''' <returns>The cleaned text.</returns>
        Private Shared Function CleanupRtfRemnants(ByVal s As String) As String
            If String.IsNullOrEmpty(s) Then Return s

            ' Paragraph and tab controls -> line breaks / spaces
            s = Regex.Replace(s, "\s*\\par\b\s*", vbCrLf, RegexOptions.IgnoreCase)
            s = Regex.Replace(s, "\\tab\b", " ", RegexOptions.IgnoreCase)

            ' Remove paragraph formatting tokens (\pard, \liN, \fiN, \txN, \saN)
            s = Regex.Replace(s, "\\pard\b(?:\s*\\[a-z]+-?\d+)*", "", RegexOptions.IgnoreCase)
            s = Regex.Replace(s, "(?:\s*\\(?:li|fi|tx|sa|sb)-?\d+)+", "", RegexOptions.IgnoreCase)

            ' Remove paired {{ ... }} around normalized URL (and optional " (title)")
            s = Regex.Replace(s, "\{\{\s*(https?://[^\s}]+(?:\s*\([^)]+\))?)\s*\}\}", "$1", RegexOptions.IgnoreCase)

            ' If a stray closing brace directly precedes an RTF control word, drop the brace
            s = Regex.Replace(s, "\}\s*(\\[a-zA-Z])", "$1")

            ' Remove common stray inline RTF formatting outside fields
            s = Regex.Replace(s, "\\(?:fs\d+|f\d+|cf\d+|ul0?|b0?)\b", "", RegexOptions.IgnoreCase)

            ' Normalize spaces around parentheses
            s = Regex.Replace(s, "\s+\)", ")", RegexOptions.None)
            s = Regex.Replace(s, "\(\s+", "(", RegexOptions.None)

            ' Collapse multiple spaces
            s = Regex.Replace(s, " +", " ")

            ' Trim spaces at line starts/ends and collapse excessive blank lines
            s = Regex.Replace(s, "[ \t]+\r?\n", vbCrLf)
            s = Regex.Replace(s, "\r?\n[ \t]+", vbCrLf)
            s = Regex.Replace(s, "(?:\r?\n){3,}", vbCrLf & vbCrLf).Trim()

            Return s
        End Function


        ' Evaluators kept static to avoid repeated allocations.

        ''' <summary>
        ''' Evaluates <c>\uN?</c> RTF decimal Unicode escapes and replaces them with the corresponding Unicode character
        ''' when <c>N</c> is within a supported range; otherwise returns an empty string or the original match.
        ''' </summary>
        Private Shared ReadOnly _rtfUnicodeEvaluator As MatchEvaluator =
    New MatchEvaluator(Function(m As System.Text.RegularExpressions.Match)
                           Dim raw As String = m.Groups(1).Value
                           Dim dec As Integer
                           If Integer.TryParse(raw, System.Globalization.NumberStyles.Integer,
                                               System.Globalization.CultureInfo.InvariantCulture, dec) Then
                               ' RTF negative \uN? => ignore (fallback), return empty
                               If dec >= 0 AndAlso dec <= &H10FFFF Then
                                   Return ChrW(dec)
                               Else
                                   Return String.Empty
                               End If
                           End If
                           Return m.Value
                       End Function)

        ''' <summary>
        ''' Evaluates JSON-style <c>\uXXXX</c> hex Unicode escapes and replaces them with the corresponding Unicode character
        ''' when the code point is within a supported range; otherwise returns the original match.
        ''' </summary>
        Private Shared ReadOnly _jsonUnicodeEvaluator As MatchEvaluator =
    New MatchEvaluator(Function(m As System.Text.RegularExpressions.Match)
                           Dim hex As String = m.Groups(1).Value
                           Dim code As Integer
                           If Integer.TryParse(hex, System.Globalization.NumberStyles.HexNumber,
                                               System.Globalization.CultureInfo.InvariantCulture, code) Then
                               If code <= &H10FFFF Then
                                   Return ChrW(code)
                               End If
                           End If
                           Return m.Value
                       End Function)

        ' Lightweight RTF cleanup for Excel string literals only (no brace parsing).

        ''' <summary>
        ''' Performs a lightweight RTF cleanup intended for Excel string literals (without brace-balanced parsing).
        ''' </summary>
        ''' <param name="lit">The literal text to clean.</param>
        ''' <returns>The cleaned literal.</returns>
        Private Shared Function _CleanRtfLiteral(lit As String) As String
            If String.IsNullOrEmpty(lit) Then Return lit

            Dim s = lit

            ' Common paragraph / line markers -> line feed (Excel expects CHAR(10))
            s = System.Text.RegularExpressions.Regex.Replace(s, "\\par\b\s*", vbLf, RegexOptions.IgnoreCase)
            s = System.Text.RegularExpressions.Regex.Replace(s, "\\line\b\s*", vbLf, RegexOptions.IgnoreCase)

            ' Tabs
            s = System.Text.RegularExpressions.Regex.Replace(s, "\\tab\b", vbTab, RegexOptions.IgnoreCase)

            ' Unicode escapes already handled separately, but in case they appear here:
            s = _rtfUnicodePattern.Replace(s, _rtfUnicodeEvaluator)

            ' Basic hex escaped chars like \'E4 (ä)
            s = System.Text.RegularExpressions.Regex.Replace(
                                s,
                                "\\'[0-9A-Fa-f]{2}",
                                New System.Text.RegularExpressions.MatchEvaluator(
                                    Function(m As System.Text.RegularExpressions.Match)
                                        Dim hex = m.Value.Substring(2)
                                        Try
                                            Dim b As Byte = System.Convert.ToByte(hex, 16)
                                            Return System.Text.Encoding.GetEncoding(1252).GetString(New Byte() {b})
                                        Catch
                                            Return m.Value
                                        End Try
                                    End Function))

            ' Strip simple style/control words that should not appear as text
            s = System.Text.RegularExpressions.Regex.Replace(
                    s,
                    "\\(fs\d+|f\d+|cf\d+|highlight\d*|ulnone|ul|b0|b|i0|i|super|sub|strike|pard|plain)\b",
                    "",
                    RegexOptions.IgnoreCase)

            ' Remove lone braces left from inline groups like {\*\fldinst ...}
            s = System.Text.RegularExpressions.Regex.Replace(s, "[{}]", "")

            ' Collapse multiple consecutive line feeds produced by many \par
            s = System.Text.RegularExpressions.Regex.Replace(s, "[\r\n]+", vbLf)

            Return s
        End Function



        ' Find the matching closing brace for the brace at startIdx (or at the next char if not at a brace).

        ''' <summary>
        ''' Finds the matching closing brace for an opening brace starting at (or before) the specified index.
        ''' </summary>
        ''' <param name="s">The input string.</param>
        ''' <param name="startIdx">An index at or near an opening brace.</param>
        ''' <returns>The index of the matching closing brace; otherwise <c>-1</c>.</returns>
        Private Shared Function FindMatchingBrace(ByVal s As String, ByVal startIdx As Integer) As Integer
            Dim idx As Integer = startIdx
            If idx >= s.Length Then Return -1
            ' Ensure we start at an opening brace if possible
            If s.Chars(idx) <> "{"c Then
                Dim prevOpen As Integer = s.LastIndexOf("{"c, idx)
                If prevOpen >= 0 Then idx = prevOpen
                If idx >= s.Length OrElse s.Chars(idx) <> "{"c Then Return -1
            End If

            Dim depth As Integer = 0
            For k As Integer = idx To s.Length - 1
                Dim ch As Char = s.Chars(k)
                If ch = "{"c Then
                    depth += 1
                ElseIf ch = "}"c Then
                    depth -= 1
                    If depth = 0 Then
                        Return k
                    End If
                End If
            Next
            Return -1
        End Function

        ' Extract the first "...", occurring after a token, case-insensitive.

        ''' <summary>
        ''' Extracts the first quoted string (<c>"..."</c>) that occurs after the specified token (case-insensitive).
        ''' </summary>
        ''' <param name="s">The input string.</param>
        ''' <param name="token">The token to search for.</param>
        ''' <returns>The extracted quoted string without quotes; otherwise <c>Nothing</c>.</returns>
        Private Shared Function ExtractQuotedAfter(ByVal s As String, ByVal token As String) As String
            Dim p As Integer = s.IndexOf(token, StringComparison.OrdinalIgnoreCase)
            If p = -1 Then Return Nothing
            Dim q1 As Integer = s.IndexOf(""""c, p)
            If q1 = -1 Then Return Nothing
            Dim q2 As Integer = s.IndexOf(""""c, q1 + 1)
            If q2 = -1 OrElse q2 <= q1 Then Return Nothing
            Return s.Substring(q1 + 1, q2 - (q1 + 1))
        End Function

        ' Extract content inside {\fldrslt ...} (handles both {\fldrslt text} and {\fldrslt{...}}).

        ''' <summary>
        ''' Extracts the content of an RTF <c>\fldrslt</c> block, if present.
        ''' </summary>
        ''' <param name="s">An RTF field block string.</param>
        ''' <returns>The extracted content without the leading <c>\fldrslt</c> token; otherwise <c>Nothing</c>.</returns>
        Private Shared Function ExtractFldrsltBlock(ByVal s As String) As String
            Dim p As Integer = s.IndexOf("{\fldrslt", StringComparison.OrdinalIgnoreCase)
            If p = -1 Then Return Nothing

            ' The fldrslt block itself should be brace-balanced
            Dim startBrace As Integer = s.IndexOf("{"c, p)
            If startBrace = -1 Then Return Nothing

            Dim endBrace As Integer = FindMatchingBrace(s, startBrace)
            If endBrace = -1 Then Return Nothing

            ' Remove the leading {\fldrslt and surrounding braces
            Dim inner As String = s.Substring(startBrace + 1, endBrace - (startBrace + 1))
            ' Strip the control word \fldrslt
            If inner.StartsWith("\fldrslt", StringComparison.OrdinalIgnoreCase) Then
                inner = inner.Substring("\fldrslt".Length)
            End If
            Return inner
        End Function




    End Class
End Namespace