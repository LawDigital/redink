' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ContextProgressPolicy.vb
' PURPOSE
'   Detects repeated context-management reads without new coverage and bounds recovery.
'
' KEY ARCHITECTURE
'   • State: Tracks covered context ranges and repeated reads within one execution.
'   • Recovery: Bounds stagnation without treating delivered ranges as verified analysis.
' =============================================================================

' Per-execution context-management progress; coverage is delivery history, NOT proof of
' active visibility or semantic analysis. Original evidence and normal rereads remain available.
Option Explicit On
Option Strict On

Namespace Agents
    Public NotInheritable Class ContextProgressPolicy
        Public Const StalledCode As System.String = "context_progress_stalled"
        Public Const RecoveryPrompt As System.String =
            "HOST CONTEXT PROGRESS GUARD: Repeated context management is revisiting substantially the same source ranges without new coverage. " &
            "Stop the compact/expand cycle and change strategy. Preserve verified findings, artifacts, exact evidence and unresolved scope. " &
            "Prefer an exposed complete-source analysis capability accepting a stored result reference or exported file; pass the reference/path directly, not a copied source dump. " &
            "Use only available capabilities, or load missing ranges and finish the analysis. A delivered range is not proof of current visibility or semantic review. " &
            "Do not invent omitted evidence, repeat completed mutations, or claim completion while required scope remains unread. If no safe route exists, report blocked with the remaining scope."

        Private NotInheritable Class Span
            Public Start As System.Int32
            Public Finish As System.Int32
        End Class
        Private ReadOnly _coverage As New System.Collections.Generic.Dictionary(Of System.String, System.Collections.Generic.List(Of Span))(System.StringComparer.OrdinalIgnoreCase)
        ' One complete verification pass over already delivered ranges is legitimate,
        ' regardless of page size. Attempts consume this allowance even if visibility
        ' suppression avoids dispatch; otherwise identical visible reads could loop forever.
        Private ReadOnly _verificationAttempts As New System.Collections.Generic.Dictionary(Of System.String, System.Collections.Generic.List(Of Span))(System.StringComparer.OrdinalIgnoreCase)
        Private ReadOnly _completedCalls As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
        Private _redundantReads As System.Int32
        Private _compactionsWithoutRead As System.Int32
        Private _recoveryIssued As System.Boolean
        Public Property LastDiagnostic As System.String = System.String.Empty

        ' Called before dispatch, including attempts suppressed by the visibility guard.
        ' 0 = continue, 1 = one strategy-recovery turn, 2 = controlled blocked result.
        Public Function Evaluate(toolName As System.String, canonicalWindowKey As System.String) As System.Int32
            If ContextCompactTool.IsContextCompactTool(toolName) Then
                _compactionsWithoutRead += 1
            ElseIf ContextExpandTool.IsContextExpandTool(toolName) Then
                Dim resultRef As System.String = System.String.Empty
                Dim start As System.Int32
                Dim finish As System.Int32
                If Not TryParseWindowKey(canonicalWindowKey, resultRef, start, finish) Then Return 0
                Dim length As System.Int64 = CLng(finish) - start
                Dim newChars As System.Int64 = CountNewCharacters(_coverage, resultRef, start, finish)
                ' A small boundary drift must not defeat the guard. First reads, including
                ' short last pages, always count as progress; overlap <= 95% also progresses.
                If length > 0 AndAlso newChars * 100L > length * 5L Then
                    LastDiagnostic = "new_coverage; ref=" & resultRef & "; newChars=" & newChars.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    Return 0
                End If
                Dim newVerificationChars As System.Int64 = CountNewCharacters(_verificationAttempts, resultRef, start, finish)
                AddSpan(_verificationAttempts, resultRef, start, finish)
                If length > 0 AndAlso newVerificationChars * 100L > length * 5L Then
                    _redundantReads = 0
                    _compactionsWithoutRead = 0
                    LastDiagnostic = "verification_pass; window=" & canonicalWindowKey
                    Return 0
                End If
                _redundantReads += 1
                LastDiagnostic = "revisited_coverage; ref=" & resultRef & "; window=" & canonicalWindowKey & "; newChars=" & newChars.ToString(System.Globalization.CultureInfo.InvariantCulture)
            Else
                Return 0
            End If
            Dim limit As System.Int32 = If(_recoveryIssued, ToolingConstants.ContextChurnRecoveryAttempts, ToolingConstants.ContextChurnInitialAttempts)
            If _redundantReads < limit AndAlso _compactionsWithoutRead < limit Then Return 0
            LastDiagnostic &= "; redundantReads=" & _redundantReads.ToString(System.Globalization.CultureInfo.InvariantCulture) & "; compactionsWithoutRead=" & _compactionsWithoutRead.ToString(System.Globalization.CultureInfo.InvariantCulture)
            If _recoveryIssued Then Return 2
            _recoveryIssued = True
            _redundantReads = 0
            _compactionsWithoutRead = 0
            Return 1
        End Function

        Public Sub RecordSuccess(toolName As System.String, responseText As System.String, callSignature As System.String)
            If ContextCompactTool.IsContextCompactTool(toolName) Then Return
            If ContextExpandTool.IsContextExpandTool(toolName) Then
                Dim key As System.String = System.String.Empty
                If Not ContextExpandTool.TryBuildCanonicalWindowKeyFromResponse(responseText, key) Then Return
                Dim resultRef As System.String = System.String.Empty
                Dim start As System.Int32
                Dim finish As System.Int32
                If Not TryParseWindowKey(key, resultRef, start, finish) Then Return
                Dim length As System.Int64 = CLng(finish) - start
                Dim newChars As System.Int64 = CountNewCharacters(_coverage, resultRef, start, finish)
                If length > 0 AndAlso newChars * 100L > length * 5L Then ResetEpisode()
                _compactionsWithoutRead = 0
                AddSpan(_coverage, resultRef, start, finish)
            ElseIf Not System.String.IsNullOrWhiteSpace(callSignature) AndAlso _completedCalls.Add(callSignature) Then
                ' A new successful work step permits later deliberate verification of old
                ' sources. Replaying an identical successful call does not reset the guard.
                ResetEpisode()
                _verificationAttempts.Clear()
            End If
        End Sub

        Private Sub ResetEpisode()
            _redundantReads = 0
            _compactionsWithoutRead = 0
            _recoveryIssued = False
        End Sub

        Private Shared Function TryParseWindowKey(key As System.String, ByRef resultRef As System.String, ByRef start As System.Int32, ByRef finish As System.Int32) As System.Boolean
            Dim parts As System.String() = If(key, System.String.Empty).Split("|"c)
            If parts.Length <> 3 OrElse System.String.IsNullOrWhiteSpace(parts(0)) Then Return False
            resultRef = parts(0)
            Return System.Int32.TryParse(parts(1), System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, start) AndAlso
                   System.Int32.TryParse(parts(2), System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, finish) AndAlso
                   start >= 0 AndAlso finish >= start
        End Function

        Private Shared Function CountNewCharacters(coverage As System.Collections.Generic.Dictionary(Of System.String, System.Collections.Generic.List(Of Span)), resultRef As System.String, start As System.Int32, finish As System.Int32) As System.Int64
            Dim count As System.Int64 = CLng(finish) - start
            Dim spans As System.Collections.Generic.List(Of Span) = Nothing
            If coverage.TryGetValue(resultRef, spans) Then
                For Each item As Span In spans
                    count -= System.Math.Max(0L, CLng(System.Math.Min(finish, item.Finish)) - System.Math.Max(start, item.Start))
                Next
            End If
            Return count
        End Function

        Private Shared Sub AddSpan(coverage As System.Collections.Generic.Dictionary(Of System.String, System.Collections.Generic.List(Of Span)), resultRef As System.String, start As System.Int32, finish As System.Int32)
            If finish <= start Then Return
            Dim spans As System.Collections.Generic.List(Of Span) = Nothing
            If Not coverage.TryGetValue(resultRef, spans) Then
                spans = New System.Collections.Generic.List(Of Span)()
                coverage.Add(resultRef, spans)
            End If
            Dim index As System.Int32 = 0
            While index < spans.Count AndAlso spans(index).Finish < start
                index += 1
            End While
            While index < spans.Count AndAlso spans(index).Start <= finish
                start = System.Math.Min(start, spans(index).Start)
                finish = System.Math.Max(finish, spans(index).Finish)
                spans.RemoveAt(index)
            End While
            spans.Insert(index, New Span With {.Start = start, .Finish = finish})
        End Sub

        ' One-shot compaction must not resurrect shelved history on the next rebuild.
        ' Validate availability before keeping a stub; failed storage retains full evidence.
        Public Shared Function CanRetainHistoricalReplay(replayContent As System.String) As System.Boolean
            If System.String.IsNullOrWhiteSpace(replayContent) Then Return False
            Try
                Dim obj As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(replayContent)
                Dim stored As ToolResultStore.StoredResult = Nothing
                Return ToolResultStore.TryGetForWorkflow(obj.Value(Of System.String)("result_ref"), WorkflowContinuity.CurrentWorkflowId, stored) AndAlso
                       stored IsNot Nothing AndAlso stored.IsDurable
            Catch ex As System.Exception
                System.Diagnostics.Trace.TraceWarning("Historical replay cannot be retained; restoring full result: " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
