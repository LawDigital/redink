' Part of "Red Ink" (SharedLibrary)
' Deterministic edits of user-authored source controls; never parses retrieved content.

' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' =============================================================================
' File: RetrievalPromptEditing.vb
' Purpose:
'   Deterministic selection, history restoration and deduplication of user-authored
'   retrieval-source controls.
'
' Architecture / Function:
'   Edits prompt control syntax only; retrieved document content is never interpreted as
'   a user scope instruction.
' =============================================================================

Option Strict On
Option Explicit On

Namespace SharedLibrary
    Public NotInheritable Class RetrievalPromptEditing
        Private Sub New()
        End Sub

        Private NotInheritable Class Span
            Friend Start As System.Int32
            Friend Length As System.Int32
            Friend Raw As System.String
            Friend Key As System.String
            Friend Provider As System.String
            Friend ScopeOnly As System.Boolean
            Friend Qualified As System.Boolean
        End Class

        Private Shared Function Spans(text As System.String) As System.Collections.Generic.List(Of Span)
            Dim result As New System.Collections.Generic.List(Of Span)()
            For Each request As SemanticArchiveRequest In SemanticArchiveTriggerHelper.Parse(text)
                If request.ErrorCode.Length > 0 OrElse request.Length <= 0 Then Continue For
                Dim selectors As New System.Collections.Generic.List(Of System.String)(request.ArchiveSelectors)
                selectors.Sort(System.StringComparer.OrdinalIgnoreCase)
                result.Add(New Span With {.Start = request.Start, .Length = request.Length, .Raw = request.RawTrigger,
                    .Key = "sa:" & Newtonsoft.Json.JsonConvert.SerializeObject(selectors), .Provider = "sa", .ScopeOnly = request.UsesSurroundingTask AndAlso request.Mode = "content",
                    .Qualified = request.ArchiveSelectors.Count > 0})
            Next
            ' KB retains its existing grammar. Only simple scope controls are edited;
            ' query/tag/legacy controls and malformed text are left byte-for-byte intact.
            Dim position As System.Int32 = 0
            While position < text.Length
                Dim opening As System.Int32 = text.IndexOf("(kb", position, System.StringComparison.OrdinalIgnoreCase)
                If opening < 0 Then Exit While
                Dim ending As System.Int32 = text.IndexOf(")"c, opening)
                If ending < 0 Then Exit While
                Dim raw As System.String = text.Substring(opening, ending - opening + 1)
                Dim request As KnowledgeTriggerHelper.KnowledgeRequest = KnowledgeTriggerHelper.TryParseKnowledgeTrigger(raw)
                If request IsNot Nothing AndAlso System.String.Equals(request.RawTrigger, raw, System.StringComparison.OrdinalIgnoreCase) Then
                    Dim scopeOnly As System.Boolean = System.Text.RegularExpressions.Regex.IsMatch(raw,
                        "^\(kb(?::\s*(?:store:\s*(?:""[^""()\r\n]+""|[^\s()]+))?\s*)?\)$",
                        System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.CultureInvariant)
                    result.Add(New Span With {.Start = opening, .Length = raw.Length, .Raw = raw, .Provider = "kb",
                        .Key = "kb:" & If(request.StoreName, System.String.Empty).Trim(), .ScopeOnly = scopeOnly, .Qualified = request.HasExplicitStoreFilter})
                End If
                position = ending + 1
            End While
            result.Sort(Function(left As Span, right As Span) left.Start.CompareTo(right.Start))
            Return result
        End Function

        ''' <summary>The source picker is a scope selection, not an append operation.
        ''' Replace scope-only controls of that provider; preserve explicit queries,
        ''' modes, the other provider and all surrounding task text.</summary>
        Public Shared Function SelectSource(prompt As System.String, insertion As System.String) As System.String
            Dim text As System.String = If(prompt, System.String.Empty)
            Dim selected As System.Collections.Generic.List(Of Span) = Spans(If(insertion, System.String.Empty))
            If selected.Count <> 1 OrElse Not selected(0).ScopeOnly Then Return text
            Dim matches As New System.Collections.Generic.List(Of Span)()
            For Each item As Span In Spans(text)
                If item.Provider = selected(0).Provider AndAlso item.ScopeOnly Then matches.Add(item)
            Next
            If matches.Count = 0 Then Return text & If(text.Length > 0 AndAlso Not System.Char.IsWhiteSpace(text(text.Length - 1)), " ", "") & insertion
            For index As System.Int32 = matches.Count - 1 To 0 Step -1
                Dim item As Span = matches(index)
                text = text.Remove(item.Start, item.Length).Insert(item.Start, If(index = 0, insertion, System.String.Empty))
            Next
            Return text
        End Function

        ''' <summary>Restore history at the caret without carrying over the previous
        ''' dialog's source-only selections. Other controls and task text remain intact.</summary>
        Public Shared Function RestoreHistory(prompt As System.String, previousPrompt As System.String,
                                              selectionStart As System.Int32) As System.String
            Dim text As System.String = If(prompt, System.String.Empty)
            Dim previous As System.String = NormalizeScopes(previousPrompt)
            If previous.Length = 0 Then Return text
            If System.String.Equals(NormalizeScopes(text).Trim(), previous.Trim(), System.StringComparison.Ordinal) Then Return NormalizeScopes(text)
            Dim providers As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            For Each item As Span In Spans(previous)
                If item.ScopeOnly Then providers.Add(item.Provider)
            Next
            Dim caret As System.Int32 = System.Math.Max(0, System.Math.Min(selectionStart, text.Length))
            Dim current As System.Collections.Generic.List(Of Span) = Spans(text)
            For index As System.Int32 = current.Count - 1 To 0 Step -1
                Dim item As Span = current(index)
                If Not item.ScopeOnly OrElse Not providers.Contains(item.Provider) Then Continue For
                If item.Start < caret Then caret -= System.Math.Min(item.Length, caret - item.Start)
                text = text.Remove(item.Start, item.Length)
            Next
            Return NormalizeScopes(text.Insert(caret, previous))
        End Function

        ''' <summary>Remove duplicate scope controls after history insertion, without
        ''' collapsing distinct explicitly queried sources or changing KB grammar.</summary>
        Public Shared Function NormalizeScopes(prompt As System.String) As System.String
            Dim text As System.String = If(prompt, System.String.Empty)
            Dim spansFound As System.Collections.Generic.List(Of Span) = Spans(text)
            Dim qualified As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            For Each item As Span In spansFound
                If item.ScopeOnly AndAlso item.Qualified Then qualified.Add(item.Provider)
            Next
            Dim seen As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
            Dim remove As New System.Collections.Generic.List(Of Span)()
            For Each item As Span In spansFound
                If Not item.ScopeOnly Then Continue For
                If (Not item.Qualified AndAlso qualified.Contains(item.Provider)) OrElse Not seen.Add(item.Key) Then remove.Add(item)
            Next
            For index As System.Int32 = remove.Count - 1 To 0 Step -1
                text = text.Remove(remove(index).Start, remove(index).Length)
            Next
            Return text
        End Function
    End Class
End Namespace
