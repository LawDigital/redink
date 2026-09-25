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

        Private ReadOnly _envelopeRx As New Regex(
            "<TASK_STATUS>\s*(?<body>\{[\s\S]*?\})\s*</TASK_STATUS>",
            RegexOptions.Compiled Or RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant)

        Private ReadOnly _trailingEnvelopeRx As New Regex(
            "(?<footer><TASK_STATUS>\s*\{[\s\S]*?\}\s*</TASK_STATUS>)\s*$",
            RegexOptions.Compiled Or RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant)

        ''' <summary>Parses the trailing TASK_STATUS footer. Returns Missing if none, Invalid if malformed or duplicated.</summary>
        Public Function Parse(text As String) As TaskStatusFooter
            Dim result As New TaskStatusFooter() With {.Kind = TaskStatusKind.Missing}
            If String.IsNullOrWhiteSpace(text) Then Return result

            ' TASK_STATUS is a trailing parent-response contract. Literal tags inside
            ' JSON/tool/agent data are not protocol envelopes and must be ignored.
            Dim m As Match = _trailingEnvelopeRx.Match(text)
            If Not m.Success Then Return result

            Dim prefix As String = text.Substring(0, m.Index).TrimEnd()
            If _trailingEnvelopeRx.IsMatch(prefix) Then
                result.Kind = TaskStatusKind.Invalid
                result.InvalidDetail = "multiple_task_status_footers"
                Return result
            End If

            result.StartIndex = m.Index
            result.EndIndex = m.Index + m.Length
            Dim strictMatch As Match = _envelopeRx.Match(m.Groups("footer").Value)
            If Not strictMatch.Success Then
                result.Kind = TaskStatusKind.Invalid
                result.InvalidDetail = "malformed_task_status_footer"
                Return result
            End If
            result.RawJson = strictMatch.Groups("body").Value

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
        ''' Strips only a trailing top-level TASK_STATUS envelope. Literal TASK_STATUS
        ''' text embedded in agent/tool payloads is data and must remain untouched.
        ''' </summary>
        Public Function Strip(text As String) As String
            If String.IsNullOrEmpty(text) Then Return text
            Dim cleaned As String = _trailingEnvelopeRx.Replace(text, "")
            Return cleaned.TrimEnd()
        End Function

        ''' <summary>Returns the prose part (everything BEFORE the footer if present, else the whole text).</summary>
        Public Function ExtractProse(text As String) As String
            If String.IsNullOrEmpty(text) Then Return ""
            Dim parsed As TaskStatusFooter = Parse(text)
            If parsed.Kind = TaskStatusKind.Missing Then Return text.TrimEnd()
            Return text.Substring(0, parsed.StartIndex).TrimEnd()
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