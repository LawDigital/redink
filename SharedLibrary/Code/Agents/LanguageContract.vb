' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved.
' For license to use see https://redink.ai.
'
' =============================================================================
' File: LanguageContract.vb
' Purpose: Centralizes (1) the hard system-prompt rule that final user-facing
'          prose must be in the detected user language, and (2) host-side
'          post-localization decisions for final prose when the model ignored
'          that contract.
' =============================================================================

Option Explicit On
Option Strict On

Imports System.Text.RegularExpressions

Namespace Agents

    Public Module LanguageContract

        ''' <summary>
        ''' Builds the language-contract block that MUST be appended to every iteration's
        ''' system prompt. This ensures the model answers in the user's language regardless
        ''' of guard-prompt language or tool-output language.
        ''' </summary>
        Public Function BuildSystemPromptFragment(userLanguage As String) As String
            Dim lang As String = If(userLanguage, "").Trim()
            If lang = "" Then Return ""

            Dim sb As New System.Text.StringBuilder()
            sb.AppendLine("[LANGUAGE CONTRACT — MANDATORY]")
            sb.AppendLine("- The user's language for THIS run is: " & lang)
            sb.AppendLine("- All FINAL user-facing prose (the message the user reads) MUST be written in " & lang & ".")
            sb.AppendLine("- This applies to both successful and blocked final answers, and to any explanation of failures or limitations.")
            sb.AppendLine("- Tool call arguments, JSON envelopes, structured payloads, and the <TASK_STATUS> footer remain in English regardless of user language.")
            sb.AppendLine("- Do NOT switch the language of your final prose just because tool output, host guard prompts, or system text are in English.")
            sb.AppendLine("[/LANGUAGE CONTRACT]")
            Return sb.ToString().TrimEnd()
        End Function

        ''' <summary>
        ''' Returns True when the host should post-localize a final response because
        ''' the target language is non-English and the prose still does not look like
        ''' that target language.
        ''' </summary>
        Public Function ShouldPostLocalizeFinal(prose As String,
                                                userLanguage As String,
                                                finalStatus As String,
                                                maxLocalizableChars As Integer) As Boolean
            If String.IsNullOrWhiteSpace(prose) Then Return False
            If String.IsNullOrWhiteSpace(userLanguage) Then Return False
            If LooksLikeEnglish(userLanguage) Then Return False
            If prose.Length > maxLocalizableChars Then Return False

            ' Post-localization is intentionally skipped when the final response contains an
            ' embedded structured JSON payload. The language contract already requires the model
            ' to produce user-facing prose in the user's language, while JSON/tool/host payloads
            ' must remain byte-stable. Sending a mixed prose+JSON response through a translator
            ' can silently corrupt command envelopes, tool metadata, or structured values.
            If ContainsEmbeddedJsonPayload(prose) Then Return False

            If ProseLooksLikeTargetLanguage(prose, userLanguage) Then Return False

            Dim status As String = If(finalStatus, "").Trim().ToLowerInvariant()
            Return status = "blocked" OrElse status = "complete"
        End Function

        ''' <summary>
        ''' Detects a complete JSON object or array embedded anywhere inside otherwise user-facing
        ''' text. This is host-agnostic and exists only to protect structured payloads from the
        ''' optional post-localization fallback. Brackets inside JSON strings are ignored.
        ''' </summary>
        Private Function ContainsEmbeddedJsonPayload(text As String) As Boolean
            Dim raw As String = If(text, "")
            If raw = "" Then Return False

            For startIndex As Integer = 0 To raw.Length - 1
                Dim opening As Char = raw(startIndex)
                If opening <> "{"c AndAlso opening <> "["c Then Continue For

                Dim endIndex As Integer = FindBalancedJsonPayloadEnd(raw, startIndex)
                If endIndex < startIndex Then Continue For

                Dim candidate As String = raw.Substring(startIndex, endIndex - startIndex + 1)
                Try
                    Newtonsoft.Json.Linq.JToken.Parse(candidate)
                    Return True
                Catch ex As Newtonsoft.Json.JsonException
                    ' Not valid JSON at this opening delimiter. Continue scanning.
                End Try
            Next

            Return False
        End Function

        Private Function FindBalancedJsonPayloadEnd(text As String, startIndex As Integer) As Integer
            If String.IsNullOrEmpty(text) OrElse
               startIndex < 0 OrElse
               startIndex >= text.Length Then
                Return -1
            End If

            Dim firstChar As Char = text(startIndex)
            If firstChar <> "{"c AndAlso firstChar <> "["c Then Return -1

            Dim expectedClosers As New System.Collections.Generic.Stack(Of Char)()
            Dim inString As Boolean = False
            Dim escaped As Boolean = False

            For index As Integer = startIndex To text.Length - 1
                Dim currentChar As Char = text(index)

                If inString Then
                    If escaped Then
                        escaped = False
                    ElseIf currentChar = "\"c Then
                        escaped = True
                    ElseIf currentChar = """"c Then
                        inString = False
                    End If
                    Continue For
                End If

                If currentChar = """"c Then
                    inString = True
                    Continue For
                End If

                Select Case currentChar
                    Case "{"c
                        expectedClosers.Push("}"c)
                    Case "["c
                        expectedClosers.Push("]"c)
                    Case "}"c, "]"c
                        If expectedClosers.Count = 0 OrElse expectedClosers.Peek() <> currentChar Then
                            Return -1
                        End If
                        expectedClosers.Pop()
                        If expectedClosers.Count = 0 Then Return index
                End Select
            Next

            Return -1
        End Function

        Private Function LooksLikeEnglish(language As String) As Boolean
            Dim l As String = If(language, "").Trim().ToLowerInvariant()
            If l = "" Then Return False
            If l = "en" Then Return True
            If l.StartsWith("en-", StringComparison.OrdinalIgnoreCase) Then Return True
            If l = "english" Then Return True
            Return False
        End Function

        ''' <summary>
        ''' Cheap heuristic: counts language-specific letter clusters. Not perfect, but
        ''' good enough to skip retranslation for obvious matches.
        ''' </summary>
        Private Function ProseLooksLikeTargetLanguage(prose As String, language As String) As Boolean
            Dim l As String = If(language, "").Trim().ToLowerInvariant()
            If l = "" Then Return False

            If l.StartsWith("de", StringComparison.OrdinalIgnoreCase) OrElse l = "german" Then
                Return Regex.IsMatch(prose, "[äöüÄÖÜß]") OrElse
                       Regex.IsMatch(prose, "\b(?:nicht|ich|und|aber|kann|werden|wurde|leider|die|der|das|ist)\b", RegexOptions.IgnoreCase)
            End If

            If l.StartsWith("fr", StringComparison.OrdinalIgnoreCase) OrElse l = "french" Then
                Return Regex.IsMatch(prose, "[àâçéèêëîïôûùüÿñæœ]") OrElse
                       Regex.IsMatch(prose, "\b(?:je|ne|pas|nous|mais|peut|impossible|d(?:é|e)sol(?:é|e))\b", RegexOptions.IgnoreCase)
            End If

            If l.StartsWith("it", StringComparison.OrdinalIgnoreCase) OrElse l = "italian" Then
                Return Regex.IsMatch(prose, "\b(?:non|sono|ma|posso|impossibile|spiacente)\b", RegexOptions.IgnoreCase)
            End If

            If l.StartsWith("es", StringComparison.OrdinalIgnoreCase) OrElse l = "spanish" Then
                Return Regex.IsMatch(prose, "[ñáéíóúü¿¡]") OrElse
                       Regex.IsMatch(prose, "\b(?:no|pero|puedo|imposible|lo\s+siento)\b", RegexOptions.IgnoreCase)
            End If

            Return False
        End Function

    End Module

End Namespace