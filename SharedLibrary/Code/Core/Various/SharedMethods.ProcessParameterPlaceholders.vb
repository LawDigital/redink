' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: SharedMethods.ProcessParameterPlaceholders.vb
' Purpose: Expands `{parameterN=...}` definitions and `{parameterN}` references in a script by
'          prompting the user for values and then replacing the placeholders in-place.
'
' Placeholder syntax:
'  - Definition: `{parameterN = Description; type; default; range-or-options; options}`
'    - `N` is an integer (e.g., `parameter1`, `parameter2`, ...).
'    - `Description` is shown to the user.
'    - `type` supports: `string`, `boolean`, `integer`, `long`, `double` (case-insensitive).
'    - `default` is interpreted according to `type`.
'    - `range-or-options` and `options` are parsed depending on `type`:
'       - Numeric types: a `min-max` range (`\d+-\d+`) or an options list.
'       - String type: an options list.
'    - Options may include a display label and a code value via angle brackets:
'      `Label <CodeValue>`. The UI shows `Label` while replacements use `CodeValue`.
'  - Reference: `{parameterN}` inserts the value chosen for the corresponding definition.
'
' Replacement behavior:
'  - Definitions are replaced first, then references.
'  - If a placeholder occurrence is immediately surrounded by double quotes (ignoring whitespace),
'    the replacement is JSON-escaped via `JsonEscape` before insertion.
'
' External dependencies:
'  - `SharedMethods.InputParameter` and `ShowCustomVariableInputForm` (UI prompting)
'  - `AN` (used in the UI caption text)
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Globalization
Imports System.Text.RegularExpressions

Namespace SharedLibrary
    Partial Public Class SharedMethods

        ' Placeholder examples:
        '   {parameter1 = Description; type; default; range-or-options; options}
        '   {parameter1}  (reference/reuse)
        '   {parameter2 = Description; type; default; range-or-options; options}
        '   {parameter2}  (reference/reuse)

        ''' <summary>
        ''' Prompts for parameter values defined as `{parameterN=...}` in <paramref name="script"/> and replaces
        ''' both definitions and `{parameterN}` references with the chosen values.
        ''' </summary>
        ''' <param name="script">Script text containing parameter definitions and/or references.</param>
        ''' <returns>
        ''' <c>True</c> if the user confirmed the parameter input form or if no parameter definitions exist;
        ''' otherwise <c>False</c>.
        ''' </returns>
        Public Shared Function ProcessParameterPlaceholders(ByRef script As String) As Boolean
            ' Match `{parameterN=...}` definitions (captures N and definition body)
            Dim defRx As New System.Text.RegularExpressions.Regex("\{\s*parameter(\d+)\s*=\s*(.*?)\}", RegexOptions.IgnoreCase Or RegexOptions.Singleline)
            ' Match `{parameterN}` references (captures N)
            Dim refRx As New System.Text.RegularExpressions.Regex("\{\s*parameter(\d+)\s*\}", RegexOptions.IgnoreCase)

            ' Collect definition matches first
            Dim defMatches = defRx.Matches(script)
            If defMatches.Count = 0 Then
                ' No definitions -> nothing to prompt for; still OK (references stay literal)
                Return True
            End If

            ' First definition wins per parameter number
            Dim paramDefs As New Dictionary(Of Integer, (MatchObj As Match, Definition As String))()
            For Each m As Match In defMatches
                Dim num = Integer.Parse(m.Groups(1).Value)
                If Not paramDefs.ContainsKey(num) Then
                    paramDefs(num) = (m, m.Groups(2).Value.Trim())
                End If
            Next

            ' Prepare UI parameter list
            Dim parameterDefs As New List(Of SharedLibrary.SharedMethods.InputParameter)()
            Dim metaList As New List(Of (ParamNumber As Integer,
                                     TypeName As String,
                                     RangeTuple As (Integer?, Integer?),
                                     DisplayOptions As List(Of String),
                                     CodeOptions As List(Of String)))()

            For Each num In paramDefs.Keys.OrderBy(Function(i) i)
                Dim definition = paramDefs(num).Definition

                ' Split definition by ';' into: description; type; default; (range or options); (options)
                Dim segments = definition.Split(";"c).
                                      Select(Function(s) s.Trim()).
                                      Where(Function(s) s.Length > 0).ToArray()
                If segments.Length = 0 Then Continue For

                Dim desc = segments(0)
                Dim t = If(segments.Length > 1, segments(1).ToLowerInvariant(), "string")
                Dim defaultStr = If(segments.Length > 2, segments(2), "")

                Dim rangeTuple As (Integer?, Integer?) = (Nothing, Nothing)
                Dim optsRaw As List(Of String) = Nothing

                ' Parse numeric range or options
                If (t = "integer" OrElse t = "long" OrElse t = "double") AndAlso segments.Length > 3 Then
                    If System.Text.RegularExpressions.Regex.IsMatch(segments(3), "^\d+\s*-\s*\d+$") Then
                        Dim parts = segments(3).Split("-"c)
                        Dim minVal = Integer.Parse(parts(0))
                        Dim maxVal = Integer.Parse(parts(1))
                        rangeTuple = (minVal, maxVal)
                        If segments.Length > 4 Then
                            optsRaw = SplitOptionsRespectingAngles(segments(4))
                        End If
                    Else
                        optsRaw = SplitOptionsRespectingAngles(segments(3))
                    End If
                ElseIf t = "string" AndAlso segments.Length > 3 Then
                    ' Parse string options
                    optsRaw = SplitOptionsRespectingAngles(segments(3))
                End If

                ' Build display list and code list (supports `Label <Code>`)
                Dim displayList As List(Of String) = Nothing
                Dim codeList As List(Of String) = Nothing
                If optsRaw IsNot Nothing AndAlso optsRaw.Count > 0 Then
                    displayList = New List(Of String)()
                    codeList = New List(Of String)()
                    For Each o In optsRaw
                        If String.IsNullOrWhiteSpace(o) Then Continue For
                        Dim lbl = o
                        Dim code = o
                        Dim idx1 = o.IndexOf("<"c)
                        Dim idx2 = o.LastIndexOf(">"c)
                        If idx1 >= 0 AndAlso idx2 > idx1 Then
                            lbl = o.Substring(0, idx1).Trim()
                            code = o.Substring(idx1 + 1, idx2 - idx1 - 1).Trim()
                        End If
                        displayList.Add(lbl)
                        codeList.Add(code)
                    Next
                End If

                ' Map default code value to its display value (if options are used)
                Dim defaultDisplay As Object = defaultStr
                If codeList IsNot Nothing Then
                    Dim idxDef = codeList.IndexOf(defaultStr)
                    If idxDef >= 0 Then defaultDisplay = displayList(idxDef)
                End If

                ' Parse typed default values by declared type
                Dim typedDefault As Object
                Select Case t
                    Case "boolean"
                        Dim b As Boolean : Boolean.TryParse(defaultStr, b) : typedDefault = b
                    Case "integer"
                        Dim i As Integer : Integer.TryParse(defaultStr, i) : typedDefault = i
                    Case "long"
                        Dim l As Long : Long.TryParse(defaultStr, l) : typedDefault = l
                    Case "double"
                        Dim d As Double
                        ' Normalize comma to dot, then parse with invariant culture
                        Dim normalizedDefault As String = defaultStr.Replace(","c, "."c)
                        Double.TryParse(normalizedDefault, NumberStyles.Float, CultureInfo.InvariantCulture, d)
                        typedDefault = d
                    Case Else
                        typedDefault = defaultDisplay
                End Select

                ' Create UI input parameter entry
                If displayList IsNot Nothing Then
                    parameterDefs.Add(New SharedLibrary.SharedMethods.InputParameter(desc, typedDefault, displayList))
                Else
                    parameterDefs.Add(New SharedLibrary.SharedMethods.InputParameter(desc, typedDefault))
                End If

                metaList.Add((num, t, rangeTuple, displayList, codeList))
            Next

            If parameterDefs.Count = 0 Then Return True

            ' Prompt user for all parameters
            Dim prmArray = parameterDefs.ToArray()
            If Not ShowCustomVariableInputForm("Please configure your parameters:", $"{AN} Parameters", prmArray) Then
                Return False
            End If

            ' Build parameter number -> raw string value map (before context escaping)
            Dim valueByParamRaw As New Dictionary(Of Integer, String)()
            For i = 0 To metaList.Count - 1
                Dim meta = metaList(i)
                Dim p = prmArray(i)
                Dim rawValue = If(p.Value Is Nothing, "", p.Value.ToString()).Trim()
                Dim finalValue As String

                If meta.TypeName = "boolean" Then
                    finalValue = rawValue.ToLowerInvariant()
                Else
                    ' Map display -> code (if options were defined)
                    If meta.DisplayOptions IsNot Nothing Then
                        Dim idx = meta.DisplayOptions.IndexOf(rawValue)
                        If idx >= 0 Then
                            finalValue = meta.CodeOptions(idx)
                        Else
                            finalValue = rawValue
                        End If
                    Else
                        Dim rvLower = rawValue.ToLowerInvariant()
                        If rvLower.StartsWith("(keine auswahl)") OrElse rvLower.StartsWith("(no selection)") OrElse rawValue.StartsWith("---") Then
                            finalValue = ""
                        Else
                            finalValue = rawValue
                        End If
                    End If

                    ' Clamp numeric values to range (if provided)
                    If meta.TypeName = "integer" OrElse meta.TypeName = "long" OrElse meta.TypeName = "double" Then
                        Dim num As Double
                        ' Normalize comma to dot, then parse with invariant culture
                        Dim normalizedValue As String = finalValue.Replace(","c, "."c)
                        If Double.TryParse(normalizedValue, NumberStyles.Float, CultureInfo.InvariantCulture, num) Then
                            If meta.RangeTuple.Item1.HasValue AndAlso num < meta.RangeTuple.Item1.Value Then num = meta.RangeTuple.Item1.Value
                            If meta.RangeTuple.Item2.HasValue AndAlso num > meta.RangeTuple.Item2.Value Then num = meta.RangeTuple.Item2.Value
                            If meta.TypeName = "integer" OrElse meta.TypeName = "long" Then
                                finalValue = CLng(System.Math.Round(num)).ToString()
                            Else
                                ' Always format with dot using invariant culture
                                finalValue = num.ToString("0.###", CultureInfo.InvariantCulture)
                            End If
                        End If
                    End If
                End If

                valueByParamRaw(meta.ParamNumber) = finalValue
            Next

            ' Positional replacement (definitions first)
            Dim sb As New System.Text.StringBuilder(script)

            ' Replace definitions with context-aware escaping
            For Each m As Match In defMatches.Cast(Of Match)().OrderByDescending(Function(mm) mm.Index)
                Dim paramNumber = Integer.Parse(m.Groups(1).Value)
                Dim replRaw = If(valueByParamRaw.ContainsKey(paramNumber), valueByParamRaw(paramNumber), "")
                Dim escapeForString = IsSurroundedByQuotes(sb, m.Index, m.Length)
                Dim repl = If(escapeForString, JsonEscape(replRaw), replRaw)
                sb.Remove(m.Index, m.Length)
                sb.Insert(m.Index, repl)
            Next

            ' Replace references `{parameterN}` with the same context-aware logic
            Dim temp = sb.ToString()
            Dim refMatches = refRx.Matches(temp)
            sb = New System.Text.StringBuilder(temp)
            For Each m As Match In refMatches.Cast(Of Match)().OrderByDescending(Function(mm) mm.Index)
                If m.Groups.Count > 0 AndAlso m.Value.Contains("=") Then Continue For
                Dim paramNumber = Integer.Parse(m.Groups(1).Value)
                If valueByParamRaw.ContainsKey(paramNumber) Then
                    Dim replRaw = valueByParamRaw(paramNumber)
                    Dim escapeForString = IsSurroundedByQuotes(sb, m.Index, m.Length)
                    Dim repl = If(escapeForString, JsonEscape(replRaw), replRaw)
                    sb.Remove(m.Index, m.Length)
                    sb.Insert(m.Index, repl)
                End If
            Next

            script = sb.ToString()
            Return True
        End Function




        ''' <summary>
        ''' Splits a comma-separated option list, ignoring commas that appear inside angle brackets
        ''' (e.g., in <c>Label &lt;Code,With,Commas&gt;</c>).
        ''' </summary>
        ''' <param name="input">The raw options string.</param>
        ''' <returns>A list of option tokens (trimmed); may be empty.</returns>
        Private Shared Function SplitOptionsRespectingAngles(input As String) As List(Of String)
            Dim result As New List(Of String)()
            If String.IsNullOrWhiteSpace(input) Then Return result
            Dim buf As New System.Text.StringBuilder()
            Dim depth As Integer = 0
            For i = 0 To input.Length - 1
                Dim ch = input(i)
                If ch = "<"c Then
                    depth += 1
                    buf.Append(ch)
                ElseIf ch = ">"c AndAlso depth > 0 Then
                    depth -= 1
                    buf.Append(ch)
                ElseIf ch = ","c AndAlso depth = 0 Then
                    Dim s = buf.ToString().Trim()
                    If s.Length > 0 Then result.Add(s)
                    buf.Clear()
                Else
                    buf.Append(ch)
                End If
            Next
            Dim last = buf.ToString().Trim()
            If last.Length > 0 Then result.Add(last)
            Return result
        End Function

        ''' <summary>
        ''' Escapes backslash and double quote characters for insertion into a quoted JSON string value.
        ''' </summary>
        ''' <param name="s">Input string to escape.</param>
        ''' <returns>The minimally escaped JSON string.</returns>
        Private Shared Function JsonEscape(s As String) As String
            If String.IsNullOrEmpty(s) Then Return s
            Return s.Replace("\", "\\").Replace("""", "\""")
        End Function

        ''' <summary>
        ''' Determines whether the placeholder occurrence at <paramref name="start"/> with <paramref name="length"/>
        ''' is immediately enclosed by double quotes, ignoring surrounding whitespace.
        ''' </summary>
        ''' <param name="sb">The script builder used for replacement.</param>
        ''' <param name="start">Start index of the placeholder match.</param>
        ''' <param name="length">Length of the placeholder match.</param>
        ''' <returns><c>True</c> if a double quote appears immediately before and after the placeholder.</returns>
        Private Shared Function IsSurroundedByQuotes(sb As System.Text.StringBuilder, start As Integer, length As Integer) As Boolean
            Dim li = start - 1
            While li >= 0 AndAlso Char.IsWhiteSpace(sb(li)) : li -= 1 : End While
            Dim ri = start + length
            While ri < sb.Length AndAlso Char.IsWhiteSpace(sb(ri)) : ri += 1 : End While
            Return (li >= 0 AndAlso sb(li) = """"c) AndAlso (ri < sb.Length AndAlso sb(ri) = """"c)
        End Function



    End Class
End Namespace