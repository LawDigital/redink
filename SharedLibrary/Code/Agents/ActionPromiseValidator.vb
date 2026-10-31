' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ActionPromiseValidator.vb
' Purpose: Detects model prose that ANNOUNCES a future action ("I will try to
'          create a text file instead...") but never invokes the corresponding tool.
'          This addresses failure mode #3 from the user report: promised but
'          unexecuted fallbacks.
'
' Patterns:
'  - Bilingual EN/DE/FR matching "I will/I'll/let me/ich werde/je vais <verb>"
'  - Captures fallback promises and alternative action announcements
'  - Used by ToolingOrchestrator to enforce guard prompts (Q10)
'
' KEY ARCHITECTURE
'   • Surface: ActionPromiseValidator (VB implementation).
'   • Entry points/helpers: ContainsActionPromise, BuildPromiseRecoveryGuardPrompt.
' =============================================================================

Option Explicit On
Option Strict On

Imports System.Text.RegularExpressions

Namespace Agents

    Public Module ActionPromiseValidator

        ' Bilingual EN/DE/FR patterns covering "I will <verb>", "instead I'll <verb>",
        ' "let me <verb>", "ich werde <verb>", "je vais <verb>".
        Private ReadOnly _patternStrings As String() = {
            "\b(?:I\s+will|I'?ll|I\s+am\s+going\s+to|let\s+me)\s+(?:try\s+to\s+|attempt\s+to\s+|now\s+)?(?:create|generate|write|save|export|attach|produce|download|fall(?:\s+|-)back|use|try|build|make|store|persist)\b",
            "\binstead[\s,]+I'?ll\s+(?:create|generate|write|save|export|attach|produce|download|fall(?:\s+|-)back|use|try|build|make|store|persist)\b",
            "\bI\s+will\s+try\s+to\s+create\b",
            "\b(?:ich\s+werde|ich\s+versuche|lass\s+mich|ich\s+erstelle\s+stattdessen)\s+(?:[a-zäöüß]+\s+){0,3}(?:erstellen|generieren|schreiben|speichern|exportieren|anh(?:ä|ae)ngen|herunterladen|verwenden|versuchen|bauen|machen|ablegen)\b",
            "\b(?:je\s+vais|laissez-moi|je\s+vais\s+essayer\s+de)\s+(?:[a-zàâçéèêëîïôûùüÿñæœ]+\s+){0,3}(?:cr(?:é|e)er|g(?:é|e)n(?:é|e)rer|(?:é|e)crire|sauvegarder|exporter|attacher|t(?:é|e)l(?:é|e)charger|utiliser|essayer)\b"
        }

        Private ReadOnly _regexes As Regex()

        Sub New()
            Dim list As New List(Of Regex)()
            For Each p As String In _patternStrings
                list.Add(New Regex(p, RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant Or RegexOptions.Compiled))
            Next
            _regexes = list.ToArray()
        End Sub

        ''' <summary>True if the prose contains a future-action promise (EN/DE/FR).</summary>
        Public Function ContainsActionPromise(prose As String) As Boolean
            If String.IsNullOrWhiteSpace(prose) Then Return False
            For Each rx As Regex In _regexes
                If rx.IsMatch(prose) Then Return True
            Next
            Return False
        End Function

        ''' <summary>Guard prompt instructing the model to either invoke the announced action or rewrite the prose.</summary>
        Public Function BuildPromiseRecoveryGuardPrompt() As String
            Return _
                "HOST ACTION-PROMISE GUARD: Your previous turn announced a future action " &
                "(""I will / I'll / let me / instead I'll / ich werde / je vais ..."") but did not actually invoke any tool to perform it. " &
                "In THIS turn you MUST EITHER:" & Environment.NewLine &
                "  (a) Actually invoke the announced tool now via a tool call, OR" & Environment.NewLine &
                "  (b) Rewrite the response WITHOUT announcing any further work and end with a valid <TASK_STATUS> footer." & Environment.NewLine &
                "It is NOT acceptable to declare 'blocked' while simultaneously promising to try a fallback. " &
                "If a fallback tool was authorized by the user (for example 'otherwise create a plain text file'), you MUST attempt that fallback at least once before declaring 'blocked'."
        End Function

    End Module

End Namespace