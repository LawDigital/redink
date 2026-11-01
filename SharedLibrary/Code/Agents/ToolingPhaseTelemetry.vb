
' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' =============================================================================
' File: ToolingPhaseTelemetry.vb
' Purpose:
'   Sanitized host-neutral phase-timing records and deterministic tool-phase
'   classification.
'
' Architecture / Function:
'   Formats elapsed/queue-wait/outcome tokens without performing log I/O or changing
'   execution policy.
' =============================================================================

Option Strict On
Option Explicit On

Namespace Agents

    Public NotInheritable Class ToolingPhaseTelemetry

        Private Sub New()
        End Sub

        Public Shared Function BuildRecord(phase As System.String,
                                           host As System.String,
                                           elapsedMilliseconds As System.Int64,
                                           outcome As System.String,
                                           Optional operation As System.String = Nothing,
                                           Optional queueWaitMilliseconds As System.Nullable(Of System.Int64) = Nothing) As System.String
            Dim safePhase As System.String = NormalizeToken(phase, "unknown")
            Dim safeHost As System.String = NormalizeToken(host, "unknown")
            Dim safeOutcome As System.String = NormalizeToken(outcome, "unknown")
            Dim safeOperation As System.String = NormalizeToken(operation, System.String.Empty)

            Dim parts As New System.Collections.Generic.List(Of System.String) From {
                "[PERF] phase_timing",
                "phase=" & safePhase,
                "host=" & safeHost,
                "elapsedMs=" & System.Math.Max(0L, elapsedMilliseconds).ToString(System.Globalization.CultureInfo.InvariantCulture),
                "outcome=" & safeOutcome
            }

            If Not System.String.IsNullOrWhiteSpace(safeOperation) Then
                parts.Add("operation=" & safeOperation)
            End If

            If queueWaitMilliseconds.HasValue Then
                parts.Add("queueWaitMs=" & System.Math.Max(0L, queueWaitMilliseconds.Value).ToString(System.Globalization.CultureInfo.InvariantCulture))
            End If

            Return System.String.Join("; ", parts)
        End Function

        Public Shared Function ResolveToolPhase(toolName As System.String) As System.String
            Dim normalized As System.String = If(toolName, System.String.Empty).Trim().ToLowerInvariant()
            If normalized = "tool_loader" Then
                Return "tool_loading"
            End If
            If normalized = "text_export_to_text" OrElse normalized = "extract_pdf_text" OrElse normalized = "read_attachment" Then
                Return "ocr_import"
            End If
            If normalized = "create_word_document" OrElse normalized.StartsWith("word_", System.StringComparison.Ordinal) OrElse normalized = "process_word_document" OrElse normalized = "pdf_to_word" Then
                Return "word_openxml"
            End If
            Return "tool_execution"
        End Function

        Private Shared Function NormalizeToken(value As System.String, fallback As System.String) As System.String
            Dim candidate As System.String = If(value, System.String.Empty).Trim()
            If candidate.Length = 0 Then
                candidate = If(fallback, System.String.Empty)
            End If

            Dim sb As New System.Text.StringBuilder(candidate.Length)
            For Each ch As System.Char In candidate
                If System.Char.IsLetterOrDigit(ch) OrElse ch = "_"c OrElse ch = "-"c OrElse ch = "."c Then
                    sb.Append(ch)
                Else
                    sb.Append("_"c)
                End If
            Next
            Return sb.ToString()
        End Function

    End Class

End Namespace
