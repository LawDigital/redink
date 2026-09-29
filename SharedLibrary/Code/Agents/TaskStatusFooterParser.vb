' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: TaskStatusFooterParser.vb
' Purpose: STRICT JSON parser for the <TASK_STATUS>{...}</TASK_STATUS> footer
'          (Q13). Replaces the regex-only ParseTaskStatus/StripTaskStatus pair
'          previously duplicated in Outlook and Word host files.
'
' Strict Rules (per Q13):
'  - Exactly zero or one footer allowed in final turn (two+ => Invalid).
'  - Body must parse as JSON object with string "status" field.
'  - Allowed status values: complete, blocked, continue (else => Invalid).
'  - Returns TaskStatusFooter with Kind, Reason, RawJson, position indices.
' =============================================================================


Option Explicit On
Option Strict On

Imports System.Text.RegularExpressions
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Agents

    Public Enum TaskStatusKind
        Missing
        Complete
        ContinueWork
        Blocked
        Invalid
    End Enum

    Public Class TaskStatusFooter
        Public Property Kind As TaskStatusKind
        Public Property Reason As String
        Public Property RawJson As String
        Public Property StartIndex As Integer
        Public Property EndIndex As Integer
        Public Property InvalidDetail As String
    End Class

    ''' <summary>
    ''' Strict, JSON-based parser/builder for the &lt;TASK_STATUS&gt; contract.
    ''' Strict rules (per Q13):
    '''   - Exactly zero or one footer is allowed in a final turn. Two or more => Invalid.
    '''   - The body must parse as a JSON object with a string "status" field.
    '''   - Allowed status values: complete, blocked, continue. Anything else => Invalid.
    ''' </summary>
    Public Module TaskStatusFooterParser

        Private Const OpenTag As System.String = "<TASK_STATUS>"
        Private Const CloseTag As System.String = "</TASK_STATUS>"

        Public NotInheritable Class TaskStatusEnvelopeLocation
            Public Property IsPresent As System.Boolean
            Public Property FooterCount As System.Int32
            Public Property StartIndex As System.Int32
            Public Property EndIndex As System.Int32
            Public Property RawBody As System.String = System.String.Empty
        End Class

        ''' <summary>
        ''' Locates the real top-level trailing TASK_STATUS protocol envelope. Literal tags
        ''' in fenced code, inline text, blockquotes, or earlier examples are data. Consecutive
        ''' top-level trailing envelopes are counted so strict callers can reject duplicates.
        ''' </summary>
        Public Function LocateTrailingEnvelope(text As System.String) As TaskStatusEnvelopeLocation
            Dim result As New TaskStatusEnvelopeLocation()
            Dim trimmedEnd As System.String = If(text, System.String.Empty).TrimEnd()
            If trimmedEnd = System.String.Empty Then Return result

            Dim current As TaskStatusEnvelopeLocation = LocateSingleTrailingEnvelope(trimmedEnd)
            If current Is Nothing Then Return result

            result.IsPresent = True
            result.FooterCount = 1
            result.StartIndex = current.StartIndex
            result.EndIndex = current.EndIndex
            result.RawBody = current.RawBody

            Dim prefix As System.String = trimmedEnd.Substring(0, current.StartIndex).TrimEnd()
            Do While prefix <> System.String.Empty
                Dim preceding As TaskStatusEnvelopeLocation = LocateSingleTrailingEnvelope(prefix)
                If preceding Is Nothing Then Exit Do
                result.FooterCount += 1
                prefix = prefix.Substring(0, preceding.StartIndex).TrimEnd()
            Loop

            Return result
        End Function

        Private Function LocateSingleTrailingEnvelope(text As System.String) As TaskStatusEnvelopeLocation
            If System.String.IsNullOrEmpty(text) OrElse
               Not text.EndsWith(CloseTag, System.StringComparison.OrdinalIgnoreCase) Then
                Return Nothing
            End If

            Dim closeStart As System.Int32 = text.Length - CloseTag.Length
            Dim candidateStarts As New System.Collections.Generic.List(Of System.Int32)()
            Dim searchFrom As System.Int32 = 0

            Do While searchFrom < closeStart
                Dim openStart As System.Int32 =
                    text.IndexOf(OpenTag, searchFrom, System.StringComparison.OrdinalIgnoreCase)
                If openStart < 0 OrElse openStart >= closeStart Then Exit Do

                If IsAcceptableProtocolStart(text, openStart) AndAlso
                   Not IsInsideMarkdownFence(text, openStart) Then
                    candidateStarts.Add(openStart)
                End If

                searchFrom = openStart + OpenTag.Length
            Loop

            ' Prefer a candidate whose complete body is already a valid JSON object. This lets the
            ' existing Newtonsoft parser define JSON string/escape boundaries, so literal protocol
            ' tags inside string values are data rather than competing envelope delimiters.
            For candidateIndex As System.Int32 = candidateStarts.Count - 1 To 0 Step -1
                Dim openStart As System.Int32 = candidateStarts(candidateIndex)
                Dim bodyStart As System.Int32 = openStart + OpenTag.Length
                Dim rawBody As System.String = text.Substring(bodyStart, closeStart - bodyStart).Trim()

                If IsCompleteJsonObject(rawBody) Then
                    Return BuildEnvelopeLocation(text.Length, openStart, rawBody)
                End If
            Next

            ' A single unambiguous protocol start must still be returned when the JSON is malformed,
            ' so Parse/strict sequencing can report the existing controlled malformed-footer result.
            ' With multiple non-data candidates and no valid JSON body, do not silently choose the
            ' last marker and "repair" an ambiguous response.
            If candidateStarts.Count = 1 Then
                Dim openStart As System.Int32 = candidateStarts(0)
                Dim bodyStart As System.Int32 = openStart + OpenTag.Length
                Dim rawBody As System.String = text.Substring(bodyStart, closeStart - bodyStart).Trim()
                Return BuildEnvelopeLocation(text.Length, openStart, rawBody)
            End If

            Return Nothing
        End Function

        Private Function BuildEnvelopeLocation(endIndex As System.Int32,
                                               openStart As System.Int32,
                                               rawBody As System.String) As TaskStatusEnvelopeLocation
            Return New TaskStatusEnvelopeLocation() With {
                .IsPresent = True,
                .FooterCount = 1,
                .StartIndex = openStart,
                .EndIndex = endIndex,
                .RawBody = If(rawBody, System.String.Empty)
            }
        End Function

        Private Function IsCompleteJsonObject(rawBody As System.String) As System.Boolean
            If System.String.IsNullOrWhiteSpace(rawBody) Then Return False

            Try
                Dim parsed As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(rawBody)
                Return parsed IsNot Nothing
            Catch ex As Newtonsoft.Json.JsonReaderException
                Return False
            Catch ex As System.Exception
                Return False
            End Try
        End Function

        Private Function IsAcceptableProtocolStart(text As System.String, openStart As System.Int32) As System.Boolean
            Dim lineStart As System.Int32 = 0
            If openStart > 0 Then
                lineStart = text.LastIndexOf(ControlChars.Lf, openStart - 1)
                If lineStart < 0 Then
                    lineStart = 0
                Else
                    lineStart += 1
                End If
            End If

            Dim linePrefix As System.String = text.Substring(lineStart, openStart - lineStart)
            Dim indentationColumns As System.Int32 = 0

            For Each prefixCharacter As System.Char In linePrefix
                If prefixCharacter = " "c Then
                    indentationColumns += 1
                ElseIf prefixCharacter = ControlChars.Tab Then
                    indentationColumns += 4 - (indentationColumns Mod 4)
                Else
                    ' Protocol control must begin at the top-level line start. Inline prose,
                    ' blockquote markers and other Markdown prefixes are data, not control.
                    Return False
                End If

                If indentationColumns >= 4 Then Return False
            Next

            Return True
        End Function

        Private Function IsInsideMarkdownFence(text As System.String, position As System.Int32) As System.Boolean
            Dim insideFence As System.Boolean = False
            Dim fenceCharacter As System.Char = System.Char.MinValue
            Dim fenceLength As System.Int32 = 0
            Dim lineStart As System.Int32 = 0

            Do While lineStart < position
                Dim lineEnd As System.Int32 = text.IndexOf(ControlChars.Lf, lineStart)
                If lineEnd < 0 OrElse lineEnd > position Then lineEnd = position

                Dim line As System.String = text.Substring(lineStart, lineEnd - lineStart)
                If line.Length > 0 AndAlso line.Chars(line.Length - 1) = ControlChars.Cr Then
                    line = line.Substring(0, line.Length - 1)
                End If

                Dim leadingSpaces As System.Int32 = 0
                Do While leadingSpaces < line.Length AndAlso leadingSpaces < 4 AndAlso line.Chars(leadingSpaces) = " "c
                    leadingSpaces += 1
                Loop

                If leadingSpaces <= 3 AndAlso leadingSpaces < line.Length Then
                    Dim markerCharacter As System.Char = line.Chars(leadingSpaces)
                    If markerCharacter = "`"c OrElse markerCharacter = "~"c Then
                        Dim markerLength As System.Int32 = 0
                        Do While leadingSpaces + markerLength < line.Length AndAlso
                                 line.Chars(leadingSpaces + markerLength) = markerCharacter
                            markerLength += 1
                        Loop

                        If markerLength >= 3 Then
                            If Not insideFence Then
                                insideFence = True
                                fenceCharacter = markerCharacter
                                fenceLength = markerLength
                            ElseIf markerCharacter = fenceCharacter AndAlso markerLength >= fenceLength Then
                                Dim remainder As System.String = line.Substring(leadingSpaces + markerLength)
                                If System.String.IsNullOrWhiteSpace(remainder) Then
                                    insideFence = False
                                    fenceCharacter = System.Char.MinValue
                                    fenceLength = 0
                                End If
                            End If
                        End If
                    End If
                End If

                If lineEnd >= position Then Exit Do
                lineStart = lineEnd + 1
            Loop

            Return insideFence
        End Function

        ''' <summary>Parses the trailing TASK_STATUS footer. Returns Missing if none, Invalid if malformed or duplicated.</summary>
        Public Function Parse(text As String) As TaskStatusFooter
            Dim result As New TaskStatusFooter() With {.Kind = TaskStatusKind.Missing}
            If String.IsNullOrWhiteSpace(text) Then Return result

            Dim location As TaskStatusEnvelopeLocation = LocateTrailingEnvelope(text)
            If location Is Nothing OrElse Not location.IsPresent Then Return result

            result.StartIndex = location.StartIndex
            result.EndIndex = location.EndIndex
            result.RawJson = If(location.RawBody, System.String.Empty).Trim()

            If location.FooterCount <> 1 Then
                result.Kind = TaskStatusKind.Invalid
                result.InvalidDetail = "multiple_task_status_footers"
                Return result
            End If

            Dim parsedKind As TaskStatusKind = TaskStatusKind.Invalid
            Dim parsedReason As String = ""
            Dim parsedDetail As String = ""

            Try
                Dim obj As JObject = JObject.Parse(result.RawJson)
                Dim statusToken As JToken = obj("status")
                If statusToken Is Nothing OrElse statusToken.Type <> JTokenType.String Then
                    parsedKind = TaskStatusKind.Invalid
                    parsedDetail = "missing_or_non_string_status_field"
                Else
                    Dim s As String = statusToken.ToString().Trim().ToLowerInvariant()
                    Select Case s
                        Case "complete", "done", "finished"
                            parsedKind = TaskStatusKind.Complete
                        Case "continue", "incomplete", "more", "in_progress"
                            parsedKind = TaskStatusKind.ContinueWork
                        Case "blocked", "failed", "abort", "impossible"
                            parsedKind = TaskStatusKind.Blocked
                        Case Else
                            parsedKind = TaskStatusKind.Invalid
                            parsedDetail = "unknown_status_value:" & s
                    End Select

                    Dim reasonToken As JToken = obj("reason")
                    If reasonToken IsNot Nothing AndAlso reasonToken.Type = JTokenType.String Then
                        parsedReason = reasonToken.ToString()
                    End If
                End If
            Catch ex As JsonReaderException
                parsedKind = TaskStatusKind.Invalid
                parsedDetail = "invalid_json:" & ex.Message
            Catch ex As System.Exception
                parsedKind = TaskStatusKind.Invalid
                parsedDetail = "footer_parse_error:" & ex.Message
            End Try

            result.Kind = parsedKind
            result.Reason = parsedReason
            result.InvalidDetail = parsedDetail
            Return result
        End Function

        ''' <summary>
        ''' Strips only the real trailing top-level TASK_STATUS envelope. Literal TASK_STATUS
        ''' text embedded in prose, quotes, fences, or payloads remains untouched.
        ''' </summary>
        Public Function Strip(text As String) As String
            If String.IsNullOrEmpty(text) Then Return text

            Dim cleaned As System.String = text.TrimEnd()
            Do While cleaned <> System.String.Empty
                Dim location As TaskStatusEnvelopeLocation = LocateSingleTrailingEnvelope(cleaned)
                If location Is Nothing OrElse Not location.IsPresent Then Exit Do
                cleaned = cleaned.Substring(0, location.StartIndex).TrimEnd()
            Loop

            Return cleaned
        End Function

        ''' <summary>Returns the prose part (everything BEFORE real trailing protocol footer(s), else the whole text).</summary>
        Public Function ExtractProse(text As String) As String
            If String.IsNullOrEmpty(text) Then Return ""
            Return Strip(text)
        End Function

        ''' <summary>Builds a strict, valid TASK_STATUS footer line.</summary>
        Public Function Build(status As String, reason As String) As String
            Dim normalizedStatus As String = If(status, "").Trim().ToLowerInvariant()
            If normalizedStatus = "" Then
                normalizedStatus = "blocked"
            End If

            Dim normalizedReason As String =
        Regex.Replace(
            If(reason, ""),
            "\s+",
            " ",
            RegexOptions.CultureInvariant).Trim()

            If normalizedReason = "" Then
                If String.Equals(normalizedStatus, "complete", StringComparison.OrdinalIgnoreCase) Then
                    normalizedReason = "answer ready"
                Else
                    normalizedReason = "no safe completion path"
                End If
            End If

            If normalizedReason.Length > ToolCallSequencing.TaskStatusReasonMaxChars Then
                normalizedReason = normalizedReason.Substring(0, ToolCallSequencing.TaskStatusReasonMaxChars).Trim()
            End If

            Dim obj As New JObject(
        New JProperty("status", normalizedStatus),
        New JProperty("reason", normalizedReason)
    )

            Return "<TASK_STATUS>" & obj.ToString(Formatting.None) & "</TASK_STATUS>"
        End Function

        ''' <summary>True if the kind represents an accepted terminal state.</summary>
        Public Function IsTerminal(kind As TaskStatusKind) As Boolean
            Return kind = TaskStatusKind.Complete OrElse kind = TaskStatusKind.Blocked
        End Function

    End Module

End Namespace