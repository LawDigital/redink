' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: JsonTemplateFormatter.vb
' Purpose: Formats JSON data using a lightweight template syntax that supports:
'          - Plain JSONPath extraction when the template contains no placeholders/loops
'          - Placeholder expansion using `{path}` with optional flags and mapping/separator rules
'          - Simple loop blocks: `{% for path %}...{% endfor %}`
'
' Template Syntax (as implemented):
'  - Plain path mode:
'      If the template contains neither `{...}` placeholders nor `{% for ... %}` loops, the template
'      is treated as a property name/pattern to be resolved via `FindJsonPropertyCustom` (Mode-dependent).
'  - Placeholders:
'      `{path}` replaces the placeholder with values from `jObj.SelectTokens(path)`.
'      Optional prefixes on the placeholder content:
'        - `html:`        converts the resolved value via `HtmlToMarkdownSimple`
'        - `nocr:`        normalizes CR/LF into spaces and collapses whitespace
'        - `htmlnocr:`    combination of `html:` and `nocr:`
'      Optional suffix after first `|`:
'        - If it contains `=` then it is treated as mappings: `key=value;key=value;...`
'        - Otherwise it is treated as a separator override for joining multiple token results
'  - Loops:
'      `{% for path %} inner {% endfor %}` selects tokens at `path`, expands arrays/objects into a
'      sequence of `JObject` items, and formats each item using `inner` as a template.
'  - Line break normalization:
'      The template treats `\n`, `\r`, `\N`, `\R` and `<cr>` (case-insensitive) as `vbCrLf`.
'
' Dependencies:
'  - Newtonsoft.Json (JObject/JToken/SelectTokens)
'  - System.Text.RegularExpressions (template parsing)
'  - System.Net (HTML decode)
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Net
Imports System.Text.RegularExpressions
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary.SharedMethods

Namespace SharedLibrary
    ''' <summary>
    ''' Provides helper functions to render JSON content into text using a lightweight template format.
    ''' </summary>
    Public Module JsonTemplateFormatter

        ''' <summary>
        ''' Formats a JSON string using the provided template.
        ''' Supports both root-level objects and arrays.
        ''' </summary>
        ''' <param name="json">JSON string to parse and format.</param>
        ''' <param name="template">Template defining extraction and formatting rules.</param>
        ''' <returns>Formatted output, or an error string if JSON parsing fails.</returns>
        Public Function FormatJsonWithTemplate(json As String, ByVal template As String) As String
            Dim jToken As JToken
            Try
                jToken = JToken.Parse(json)
            Catch ex As Newtonsoft.Json.JsonReaderException
                Return $"[Error parsing JSON: {ex.Message}]"
            End Try

            ' If root is an array, wrap it in an object with an "items" property
            If jToken.Type = JTokenType.Array Then
                Dim wrapper As New JObject()
                wrapper("items") = jToken
                NormalizeSources(wrapper)

                ' Auto-adjust template: replace {% for $ %} or {% for items %} with the correct path
                ' Also handle case where user just uses $ to mean "iterate the root array"
                template = Regex.Replace(template, "\{\%\s*for\s+\$\s*\%\}", "{% for items %}", RegexOptions.IgnoreCase)

                ' If there's no loop but template expects array iteration, we need to handle it
                ' Check if the template has a loop that would work with the wrapped structure
                If Not Regex.IsMatch(template, "\{\%\s*for\s+items\s*\%\}", RegexOptions.IgnoreCase) Then
                    ' No items loop - check if there's any loop at all
                    Dim loopMatch = Regex.Match(template, "\{\%\s*for\s+([^\s\%]+)\s*\%\}", RegexOptions.IgnoreCase)
                    If loopMatch.Success Then
                        Dim loopPath = loopMatch.Groups(1).Value
                        ' If the loop path doesn't exist in wrapper but items array has objects with that property
                        ' then we assume user wants to iterate items
                        If wrapper.SelectToken("$." & loopPath) Is Nothing Then
                            ' Replace the loop to iterate items instead, and keep inner template
                            ' This handles case like {% for results %} when array is at root
                            template = Regex.Replace(template,
                                "\{\%\s*for\s+" & Regex.Escape(loopPath) & "\s*\%\}",
                                "{% for items %}",
                                RegexOptions.IgnoreCase)
                        End If
                    End If
                End If

                Return FormatJsonWithTemplate(wrapper, template)
            End If

            Dim jObj As JObject = CType(jToken, JObject)
            NormalizeSources(jObj)
            Return FormatJsonWithTemplate(jObj, template)
        End Function

        ''' <summary>
        ''' Formats a <see cref="JObject"/> using the provided template.
        ''' </summary>
        ''' <param name="jObj">JSON object to format.</param>
        ''' <param name="template">Template defining extraction and formatting rules.</param>
        ''' <param name="Mode">
        ''' Selection behavior when the template contains no placeholders or loops:
        ''' 1 = join all matches; 2 = choose the longest non-empty match; 3 = choose the first non-empty match;
        ''' otherwise fallback to <c>FindJsonProperty</c>.
        ''' </param>
        ''' <returns>Formatted output.</returns>
        Public Function FormatJsonWithTemplate(jObj As JObject, ByVal template As String, Optional ByVal Mode As Integer = 2) As String
            If String.IsNullOrWhiteSpace(template) Then Return ""

            NormalizeSources(jObj)

            ' Normalize CRLF and placeholder markers for line breaks.
            template = template _
        .Replace("\N", vbCrLf) _
        .Replace("\n", vbCrLf) _
        .Replace("\R", vbCrLf) _
        .Replace("\r", vbCrLf)
            template = Regex.Replace(template, "<cr>", vbCrLf, RegexOptions.IgnoreCase)

            Dim hasLoop = Regex.IsMatch(template, "\{\%\s*for\s+([^\s\%]+)\s*\%\}", RegexOptions.Singleline)
            Dim hasPh = Regex.IsMatch(template, "\{([^}]+)\}")

            ' === Plain path mode ===
            If Not hasLoop AndAlso Not hasPh Then
                ' Template contains no placeholders -> treat it as a property name/pattern to locate in the JSON.
                Dim ResponseString As String = ""
                Select Case Mode
                    Case 1
                        ' 1) Join all matches in document order.
                        ResponseString = FindJsonPropertyCustom(jObj, template, SelectionMode.JoinAll)
                    Case 2
                        ' 2) Select a single match: longest non-empty.
                        ResponseString = FindJsonPropertyCustom(jObj, template, SelectionMode.LongestNonEmpty)
                    Case 3
                        ' 3) Select a single match: first non-empty.
                        ResponseString = FindJsonPropertyCustom(jObj, template, SelectionMode.FirstNonEmpty)
                    Case Else
                        ResponseString = FindJsonProperty(jObj, template)
                End Select

                Return ResponseString
            End If


            ' === Loop blocks ===
            Dim loopRegex = New Regex("\{\%\s*for\s+([^%\s]+)\s*\%\}(.*?)\{\%\s*endfor\s*\%\}", RegexOptions.Singleline Or RegexOptions.IgnoreCase)
            Dim mLoop = loopRegex.Match(template)
            While mLoop.Success
                Dim fullBlock = mLoop.Value
                Dim rawPath = mLoop.Groups(1).Value.Trim()
                Dim innerTpl = mLoop.Groups(2).Value

                Dim path = If(rawPath.StartsWith("$"), rawPath, "$." & rawPath)
                Dim tokens = jObj.SelectTokens(path)
                Dim items = tokens.SelectMany(Function(t)
                                                  If t.Type = JTokenType.Array Then
                                                      Return CType(t, JArray).OfType(Of JObject)()
                                                  ElseIf t.Type = JTokenType.Object Then
                                                      Return {CType(t, JObject)}
                                                  Else
                                                      Return Enumerable.Empty(Of JObject)()
                                                  End If
                                              End Function)

                Dim rendered = items.Select(Function(o) FormatJsonWithTemplate(o, innerTpl)).ToArray()
                template = template.Replace(fullBlock, If(rendered.Any, String.Join(vbCrLf & vbCrLf, rendered), ""))
                mLoop = loopRegex.Match(template)
            End While

            ' === Placeholders (non-greedy) ===
            Dim phRegex = New Regex("\{(.+?)\}", RegexOptions.Singleline)
            Dim result = template

            For Each mPh As Match In phRegex.Matches(template)
                Dim fullPh = mPh.Value
                Dim content = mPh.Groups(1).Value

                ' Placeholder flags.
                Dim isHtml As Boolean = False
                Dim isNoCr As Boolean = False

                If content.StartsWith("htmlnocr:", StringComparison.OrdinalIgnoreCase) Then
                    isHtml = True
                    isNoCr = True
                    content = content.Substring("htmlnocr:".Length)
                ElseIf content.StartsWith("html:", StringComparison.OrdinalIgnoreCase) Then
                    isHtml = True
                    content = content.Substring("html:".Length)
                ElseIf content.StartsWith("nocr:", StringComparison.OrdinalIgnoreCase) Then
                    isNoCr = True
                    content = content.Substring("nocr:".Length)
                End If

                ' Split only on the first "|".
                Dim parts = content.Split(New Char() {"|"c}, 2)
                Dim pathPh = parts(0).Trim()
                Dim remainder = If(parts.Length > 1, parts(1), String.Empty)

                ' Separator override (e.g. "/") or mapping definitions (contains "=").
                Dim sep As String = vbCrLf
                Dim mappings As Dictionary(Of String, String) = Nothing

                If Not String.IsNullOrEmpty(remainder) Then
                    If remainder.Contains("="c) Then
                        mappings = ParseMappings(remainder)
                    Else
                        sep = remainder.Replace("\n", vbCrLf)
                    End If
                End If

                Dim replacement = RenderTokens(jObj, pathPh, sep, isHtml, isNoCr, mappings)
                result = result.Replace(fullPh, replacement)
            Next

            Return result
        End Function


        ''' <summary>
        ''' Selection strategy used by <see cref="FindJsonPropertyCustom"/> when multiple matches exist.
        ''' </summary>
        Public Enum SelectionMode
            ''' <summary>Return the first non-empty match.</summary>
            FirstNonEmpty = 0

            ''' <summary>Return the longest non-empty match.</summary>
            LongestNonEmpty = 1

            ''' <summary>Join all non-empty matches in traversal order.</summary>
            JoinAll = 2
        End Enum

        ''' <summary>
        ''' Searches for all occurrences of a property name (case-insensitive) and returns a value based on the selected mode.
        ''' </summary>
        ''' <param name="jObj">Root JSON object.</param>
        ''' <param name="propertyName">Property name to search for (case-insensitive).</param>
        ''' <param name="mode">Selection strategy (default: <see cref="SelectionMode.JoinAll"/>).</param>
        ''' <param name="separator">Separator used when joining values (default: double CRLF).</param>
        ''' <returns>Resolved value according to the selection mode; empty string if not found.</returns>
        Public Function FindJsonPropertyCustom(jObj As JObject,
                                 propertyName As String,
                                 Optional mode As SelectionMode = SelectionMode.JoinAll,
                                 Optional separator As String = vbCrLf & vbCrLf) As String
            If jObj Is Nothing Then
                Throw New System.ArgumentNullException(NameOf(jObj))
            End If
            If String.IsNullOrWhiteSpace(propertyName) Then
                Throw New System.ArgumentException("propertyName must not be empty.", NameOf(propertyName))
            End If

            Dim matches As System.Collections.Generic.List(Of String) = CollectPropertyValues(jObj, propertyName)

            If matches Is Nothing OrElse matches.Count = 0 Then
                Return String.Empty
            End If

            Select Case mode
                Case SelectionMode.FirstNonEmpty
                    For Each s As String In matches
                        If Not String.IsNullOrWhiteSpace(s) Then
                            Return s
                        End If
                    Next
                    Return matches(0) ' If everything was empty/whitespace.

                Case SelectionMode.LongestNonEmpty
                    Dim best As String = String.Empty
                    Dim bestLen As Integer = -1
                    For Each s As String In matches
                        Dim candidate As String = If(s, String.Empty).Trim()
                        If candidate.Length > bestLen Then
                            best = candidate
                            bestLen = candidate.Length
                        End If
                    Next
                    Return best

                Case SelectionMode.JoinAll
                    Dim sb As New System.Text.StringBuilder()
                    For Each s As String In matches
                        Dim part As String = If(s, String.Empty).Trim()
                        If part.Length = 0 Then Continue For
                        If sb.Length > 0 Then sb.Append(separator)
                        sb.Append(part)
                    Next
                    Return sb.ToString()

                Case Else
                    ' Fallback.
                    Return String.Join(separator, matches)

            End Select
        End Function

        ' ---- Helper functions ---------------------------------------------------

        ''' <summary>
        ''' Collects string representations of all values for properties with the specified name by iteratively walking the JSON token tree.
        ''' </summary>
        ''' <param name="root">Root token to traverse.</param>
        ''' <param name="propertyName">Property name to match (case-insensitive).</param>
        ''' <returns>List of matched values as strings (may contain empty values).</returns>
        Private Function CollectPropertyValues(root As JToken, propertyName As String) As System.Collections.Generic.List(Of String)
            Dim results As New System.Collections.Generic.List(Of String)()
            Dim stack As New System.Collections.Generic.Stack(Of JToken)()
            stack.Push(root)

            While stack.Count > 0
                Dim node As JToken = stack.Pop()

                Select Case node.Type
                    Case JTokenType.Object
                        Dim obj As JObject = CType(node, JObject)
                        ' Push properties in reverse order to keep traversal order stable (left-to-right).
                        Dim props As System.Collections.Generic.IEnumerable(Of JProperty) = obj.Properties()
                        For Each p As JProperty In props.Reverse()
                            stack.Push(p)
                        Next

                    Case JTokenType.Property
                        Dim jp As JProperty = CType(node, JProperty)
                        If jp.Name.Equals(propertyName, StringComparison.OrdinalIgnoreCase) Then
                            Dim s As String = ConvertTokenToString(jp.Value)
                            results.Add(s)
                        End If
                        If jp.Value IsNot Nothing Then
                            stack.Push(jp.Value)
                        End If

                    Case JTokenType.Array
                        Dim arr As JArray = CType(node, JArray)
                        For i As Integer = arr.Count - 1 To 0 Step -1
                            stack.Push(arr(i))
                        Next

                    Case Else
                        ' Primitive values: no traversal needed.
                End Select
            End While

            Return results
        End Function

        ''' <summary>
        ''' Converts a JSON token to a string representation suitable for template output.
        ''' </summary>
        ''' <param name="t">Token to convert.</param>
        ''' <returns>String representation; empty string if <paramref name="t"/> is <c>Nothing</c>.</returns>
        Private Function ConvertTokenToString(t As JToken) As String
            If t Is Nothing Then Return String.Empty

            Select Case t.Type
                Case JTokenType.String
                    Return CStr(t)

                Case JTokenType.Integer, JTokenType.Float, JTokenType.Boolean, JTokenType.Null
                    Return t.ToString()

                Case JTokenType.Array, JTokenType.Object
                    ' For non-strings, return a compact JSON serialization.
                    Return t.ToString(Newtonsoft.Json.Formatting.None)

                Case Else
                    Return t.ToString()
            End Select
        End Function


        ''' <summary>
        ''' Resolves tokens from a path and renders them into a string, applying optional mapping, HTML conversion, and no-CR behavior.
        ''' </summary>
        ''' <param name="jObj">JSON object to select tokens from.</param>
        ''' <param name="path">JSONPath or relative path; if not starting with "$" or "@", "$." is prepended.</param>
        ''' <param name="sep">Separator used when joining multiple token values.</param>
        ''' <param name="isHtml">If True, values are converted via <see cref="HtmlToMarkdownSimple"/>.</param>
        ''' <param name="isNoCr">If True, line breaks are normalized to spaces and whitespace is collapsed.</param>
        ''' <param name="mappings">Optional mapping dictionary keyed by the raw token string.</param>
        ''' <returns>Rendered string, or empty string on errors/no matches.</returns>
        Private Function RenderTokens(
        jObj As JObject,
        path As String,
        sep As String,
        isHtml As Boolean,
        isNoCr As Boolean,
        mappings As Dictionary(Of String, String)
    ) As String

            Try
                If Not path.StartsWith("$") AndAlso Not path.StartsWith("@") Then
                    path = "$." & path
                End If
                Dim tokens = jObj.SelectTokens(path)
                Dim list As New List(Of String)

                For Each t In tokens
                    Dim raw = t.ToString()
                    ' Apply mapping if defined.
                    If mappings IsNot Nothing AndAlso mappings.ContainsKey(raw) Then raw = mappings(raw)
                    ' Convert HTML to Markdown-like text if requested.
                    If isHtml Then raw = HtmlToMarkdownSimple(raw)

                    ' If requested, normalize line breaks to spaces, collapse runs of whitespace, remove some bullet characters, and trim.
                    If isNoCr Then
                        ' 1) Turn all line breaks into single spaces.
                        raw = Regex.Replace(raw, "[\r\n]+", " ")

                        ' 2) Collapse any run of whitespace into one space.
                        raw = Regex.Replace(raw, "\s{2,}", " ")

                        ' 3) Remove common Unicode bullet characters only.
                        raw = Regex.Replace(raw, "[\u2022\u2023\u25E6]", String.Empty)

                        ' 4) Trim leading/trailing spaces.
                        raw = raw.Trim()
                    End If

                    list.Add(raw)
                Next

                Return If(list.Count = 0, "", String.Join(sep, list))
            Catch ex As System.Exception
                Return ""
            End Try
        End Function

        ''' <summary>
        ''' Parses mapping definitions of the form <c>key1=value1;key2=value2;...</c> into a dictionary.
        ''' </summary>
        ''' <param name="defs">Raw mapping definition string.</param>
        ''' <returns>Case-insensitive mapping dictionary.</returns>
        Private Function ParseMappings(defs As String) As Dictionary(Of String, String)
            Dim dict As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            For Each pair In defs.Split(";"c)
                Dim kv = pair.Split(New Char() {"="c}, 2)
                If kv.Length = 2 Then dict(kv(0).Trim()) = kv(1).Trim()
            Next
            Return dict
        End Function

        ''' <summary>
        ''' Converts a subset of HTML tags to a Markdown-like text representation and removes other tags.
        ''' </summary>
        ''' <param name="html">HTML input string.</param>
        ''' <returns>Converted text with line breaks and basic formatting markers.</returns>
        Public Function HtmlToMarkdownSimple(html As String) As String
            Dim s = WebUtility.HtmlDecode(html)
            ' Paragraph tags -> double line breaks.
            s = Regex.Replace(s, "</?p\s*/?>", vbCrLf & vbCrLf, RegexOptions.IgnoreCase)
            ' Line break tags.
            s = Regex.Replace(s, "<br\s*/?>", vbCrLf, RegexOptions.IgnoreCase)
            ' Bold/strong -> **text**
            s = Regex.Replace(s, "<strong>(.*?)</strong>", "**$1**", RegexOptions.IgnoreCase)
            ' Italic/em -> *text*
            s = Regex.Replace(s, "<em>(.*?)</em>", "*$1*", RegexOptions.IgnoreCase)
            ' SPAN tags -> *text*
            s = Regex.Replace(s, "<span\b[^>]*>(.*?)</span>", "*$1*", RegexOptions.IgnoreCase)
            ' List items <li> -> "- text"
            s = Regex.Replace(s, "<li>(.*?)</li>", "- $1" & vbCrLf, RegexOptions.IgnoreCase)
            ' Footnote tags <fn>...</fn> -> <sup>...</sup>
            s = Regex.Replace(s, "<fn>(.*?)</fn>", "<sup>$1</sup>", RegexOptions.IgnoreCase)
            ' Remove all remaining tags except <sup>.
            s = Regex.Replace(s, "<(?!/?sup\b)[^>]+>", String.Empty, RegexOptions.IgnoreCase)
            ' Clean up repeated line breaks.
            s = Regex.Replace(s, "(" & vbCrLf & "){3,}", vbCrLf & vbCrLf)
            Return s.Trim()
        End Function

        ''' <summary>
        ''' Normalizes the <c>sources</c> property when present as an array that may contain nested arrays,
        ''' converting entries of the form <c>[*,*,jsonObjectAsString,...]</c> into <see cref="JObject"/> items.
        ''' </summary>
        ''' <param name="jObj">JSON object to normalize in-place.</param>
        Private Sub NormalizeSources(jObj As JObject)
            Dim srcToken = jObj.SelectToken("sources")
            If srcToken IsNot Nothing AndAlso srcToken.Type = JTokenType.Array Then
                Dim newArray As New JArray()
                For Each item In CType(srcToken, JArray)
                    If item.Type = JTokenType.Array AndAlso item.Count >= 3 Then
                        Dim objStr = item(2).ToString()
                        Try
                            Dim o = JObject.Parse(objStr)
                            newArray.Add(o)
                        Catch ex As System.Exception
                            ' Skip invalid JSON string entries.
                        End Try
                    ElseIf item.Type = JTokenType.Object Then
                        newArray.Add(item)
                    End If
                Next
                jObj("sources") = newArray
            End If
        End Sub

    End Module


End Namespace