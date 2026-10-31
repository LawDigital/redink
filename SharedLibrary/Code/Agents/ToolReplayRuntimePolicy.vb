' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ToolReplayRuntimePolicy.vb
' Purpose: Shared, host-agnostic semantics for model replay retention and runtime
'          context primitives. Content strings never determine retention.
'
' KEY ARCHITECTURE
'   • Surface: ToolingRuntimePrimitives, EvidencePropertyCandidate (VB implementation).
'   • Entry points/helpers: RenderReplayTemplate, NormalizeReplayJsonSeparators,
'     IsRequiredRuntimePrimitive.
' =============================================================================

Option Strict On
Option Explicit On

Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Agents

    Public Enum ToolReplayRetentionKind
        NormalHistorical = 0
        CurrentTurnCritical = 1
        ControlPlanePinned = 2
    End Enum

    Public NotInheritable Class ToolingRuntimePrimitives

        Private NotInheritable Class EvidencePropertyCandidate
            Public Property Name As System.String
            Public Property Value As JToken
            Public Property OriginalIndex As System.Int32
            Public Property Priority As System.Int32
        End Class

        Private Sub New()
        End Sub

        ''' <summary>
        ''' Renders replay placeholders once; inserted source evidence remains literal.
        ''' </summary>
        Public Shared Function RenderReplayTemplate(template As System.String,
                                                      replacements As System.Collections.Generic.IDictionary(Of System.String, System.String)) As System.String
            Dim source As System.String = If(template, System.String.Empty)
            If replacements Is Nothing OrElse replacements.Count = 0 Then Return source
            Dim alternatives As New System.Collections.Generic.List(Of System.String)()
            For Each key As System.String In replacements.Keys
                If key.Length = 0 Then Throw New System.ArgumentException("A replay template placeholder cannot be empty.", NameOf(replacements))
                alternatives.Add(System.Text.RegularExpressions.Regex.Escape(key))
            Next
            ' One pass over the template: literal placeholder text in inserted evidence is never re-expanded.
            Return System.Text.RegularExpressions.Regex.Replace(
                source, System.String.Join("|", alternatives),
                Function(matched As System.Text.RegularExpressions.Match) If(replacements(matched.Value), System.String.Empty))
        End Function

        ''' <summary>Repairs template separators only outside JSON strings.</summary>
        Public Shared Function NormalizeReplayJsonSeparators(payload As System.String) As System.String
            If System.String.IsNullOrEmpty(payload) Then Return If(payload, System.String.Empty)
            Dim segments As System.String() = System.Text.RegularExpressions.Regex.Split(
                payload, "(""(?:\\.|[^""\\])*"")")
            For index As System.Int32 = 0 To segments.Length - 1 Step 2
                Dim structural As System.String = segments(index)
                structural = System.Text.RegularExpressions.Regex.Replace(structural, "\[\s*\]", "[]")
                structural = System.Text.RegularExpressions.Regex.Replace(structural, ",\s*,", ",")
                structural = System.Text.RegularExpressions.Regex.Replace(structural, "\[\s*,", "[")
                structural = System.Text.RegularExpressions.Regex.Replace(structural, ",\s*\]", "]")
                segments(index) = structural
            Next
            Return System.String.Concat(segments)
        End Function

        Public Shared Function IsRequiredRuntimePrimitive(toolName As System.String) As System.Boolean
            Return ContextExpandTool.IsContextExpandTool(toolName)
        End Function

        ''' <summary>
        ''' True when a runtime primitive must deliver its current-turn result body
        ''' losslessly at least once. A context expansion that is immediately replaced by
        ''' a stub has not fulfilled the read request and may create a compaction loop.
        ''' </summary>
        Public Shared Function RequiresLosslessCurrentTurnReplay(toolName As System.String) As System.Boolean
            Return ContextExpandTool.IsContextExpandTool(toolName)
        End Function

        ''' <summary>
        ''' Builds a deterministic, provider/tool-agnostic evidence core from a structured
        ''' JSON result. Values retained in the returned core are copied verbatim; large text
        ''' fields and lower-priority fields inside high-volume array items may be omitted.
        ''' No array item is silently dropped. If a sufficiently small complete item-set
        ''' cannot be produced, False is returned and the caller must use reference replay.
        ''' </summary>
        Public Shared Function TryBuildEvidenceCore(rawContent As System.String,
                                                    maxChars As System.Int32,
                                                    ByRef evidenceCoreJson As System.String,
                                                    ByRef omittedPropertyCount As System.Int32,
                                                    ByRef exactScalarCount As System.Int32,
                                                    ByRef arrayItemPropertyLimit As System.Int32,
                                                    Optional preferredArrayItemProperties As System.Int32 = ToolingConstants.EvidenceReplayPreferredArrayItemProperties,
                                                    Optional maxExactStringChars As System.Int32 = ToolingConstants.EvidenceReplayMaxExactStringChars) As System.Boolean
            evidenceCoreJson = System.String.Empty
            omittedPropertyCount = 0
            exactScalarCount = 0
            arrayItemPropertyLimit = 0

            Dim raw As System.String = If(rawContent, System.String.Empty)
            If raw = System.String.Empty OrElse maxChars <= 0 Then Return False

            Dim source As JToken
            Try
                source = JToken.Parse(raw)
            Catch ex As System.Exception
                Return False
            End Try

            If source.Type <> JTokenType.Object AndAlso source.Type <> JTokenType.Array Then
                Return False
            End If

            Dim limits As System.Int32() = New System.Int32() {
                System.Math.Max(1, preferredArrayItemProperties),
                System.Math.Min(System.Math.Max(1, preferredArrayItemProperties), 10),
                System.Math.Min(System.Math.Max(1, preferredArrayItemProperties), 8),
                System.Math.Min(System.Math.Max(1, preferredArrayItemProperties), 6),
                System.Math.Min(System.Math.Max(1, preferredArrayItemProperties), 4)
            }

            For Each propertyLimit As System.Int32 In limits
                Dim omitted As System.Int32 = 0
                Dim exactScalars As System.Int32 = 0
                Dim projected As JToken
                Try
                    projected = BuildEvidenceCoreToken(
                        source,
                        propertyLimit,
                        False,
                        omitted,
                        exactScalars,
                        maxExactStringChars)
                Catch ex As System.InvalidOperationException
                    Return False
                End Try

                If projected Is Nothing Then Continue For

                Dim serialized As System.String = projected.ToString(Formatting.None)
                If serialized.Length <= maxChars Then
                    evidenceCoreJson = serialized
                    omittedPropertyCount = omitted
                    exactScalarCount = exactScalars
                    arrayItemPropertyLimit = propertyLimit
                    Return True
                End If
            Next

            Return False
        End Function

        ''' <summary>
        ''' True when replay content represents evidence that was compacted by reference.
        ''' Includes source-produced and nested previews, even if the replay layer did
        ''' not shorten them again. Only verified pinned control blocks are excluded.
        ''' </summary>
        Public Shared Function IsEvidenceCompactedReplay(modelReplayContent As System.String) As System.Boolean
            If System.String.IsNullOrWhiteSpace(modelReplayContent) Then Return False
            Try
                Return ContainsIncompleteReference(JToken.Parse(modelReplayContent))
            Catch ex As System.Exception
                Return False
            End Try
        End Function

        Private Shared Function ContainsIncompleteReference(token As JToken) As System.Boolean
            If token Is Nothing Then Return False
            Dim obj As JObject = TryCast(token, JObject)
            If obj IsNot Nothing Then
                Dim pinned As JToken = obj("control_plane_pinned")
                If pinned IsNot Nothing AndAlso pinned.Type = JTokenType.Boolean AndAlso pinned.Value(Of System.Boolean)() Then Return False
                Dim reference As JToken = obj("result_ref")
                Dim filePath As JToken = obj("path")
                Dim fileHash As JToken = obj("snapshot_sha256")
                Dim hasStoredReference As System.Boolean = reference IsNot Nothing AndAlso reference.Type = JTokenType.String AndAlso
                    Not System.String.IsNullOrWhiteSpace(reference.Value(Of System.String)())
                Dim hasFileReference As System.Boolean = filePath IsNot Nothing AndAlso filePath.Type = JTokenType.String AndAlso
                    Not System.String.IsNullOrWhiteSpace(filePath.Value(Of System.String)()) AndAlso
                    fileHash IsNot Nothing AndAlso fileHash.Type = JTokenType.String AndAlso
                    Not System.String.IsNullOrWhiteSpace(fileHash.Value(Of System.String)())
                If hasStoredReference OrElse hasFileReference Then
                    Dim complete As JToken = obj("content_complete")
                    Dim truncated As JToken = obj("truncated")
                    If (complete IsNot Nothing AndAlso complete.Type = JTokenType.Boolean AndAlso Not complete.Value(Of System.Boolean)()) OrElse
                       (truncated IsNot Nothing AndAlso truncated.Type = JTokenType.Boolean AndAlso truncated.Value(Of System.Boolean)()) OrElse
                       obj("evidence_core") IsNot Nothing OrElse obj("preview") IsNot Nothing OrElse obj("omitted_window_chars") IsNot Nothing Then Return True
                End If
                For Each prop As JProperty In obj.Properties()
                    If ContainsIncompleteReference(prop.Value) Then Return True
                Next
            ElseIf token.Type = JTokenType.Array Then
                For Each child As JToken In DirectCast(token, JArray)
                    If ContainsIncompleteReference(child) Then Return True
                Next
            End If
            Return False
        End Function

        Private Shared Function BuildEvidenceCoreToken(source As JToken,
                                                       arrayItemPropertyLimit As System.Int32,
                                                       inArrayItem As System.Boolean,
                                                       ByRef omittedPropertyCount As System.Int32,
                                                       ByRef exactScalarCount As System.Int32,
                                                       maxExactStringChars As System.Int32) As JToken
            If source Is Nothing Then Return Nothing

            Select Case source.Type
                Case JTokenType.Object
                    Dim sourceObject As JObject = DirectCast(source, JObject)
                    Dim candidates As New System.Collections.Generic.List(Of EvidencePropertyCandidate)()
                    Dim originalIndex As System.Int32 = 0

                    For Each prop As JProperty In sourceObject.Properties()
                        Dim child As JToken = BuildEvidenceCoreToken(
                            prop.Value,
                            arrayItemPropertyLimit,
                            False,
                            omittedPropertyCount,
                            exactScalarCount,
                            maxExactStringChars)

                        If child IsNot Nothing Then
                            candidates.Add(New EvidencePropertyCandidate() With {
                                .Name = prop.Name,
                                .Value = child,
                                .OriginalIndex = originalIndex,
                                .Priority = GetEvidencePropertyPriority(prop.Name, child)
                            })
                        End If
                        originalIndex += 1
                    Next

                    If inArrayItem AndAlso
                       arrayItemPropertyLimit > 0 AndAlso
                       candidates.Count > arrayItemPropertyLimit Then

                        Dim ranked As New System.Collections.Generic.List(Of EvidencePropertyCandidate)(candidates)
                        ranked.Sort(
                            Function(left As EvidencePropertyCandidate, right As EvidencePropertyCandidate) As System.Int32
                                Dim byPriority As System.Int32 = left.Priority.CompareTo(right.Priority)
                                If byPriority <> 0 Then Return byPriority

                                Dim byLength As System.Int32 = GetEvidenceValueLength(left.Value).CompareTo(GetEvidenceValueLength(right.Value))
                                If byLength <> 0 Then Return byLength

                                Return left.OriginalIndex.CompareTo(right.OriginalIndex)
                            End Function)

                        Dim selectedNames As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
                        For index As System.Int32 = 0 To System.Math.Min(arrayItemPropertyLimit, ranked.Count) - 1
                            selectedNames.Add(ranked(index).Name)
                        Next

                        omittedPropertyCount += candidates.Count - selectedNames.Count
                        candidates = candidates.FindAll(Function(candidate As EvidencePropertyCandidate) selectedNames.Contains(candidate.Name))
                    End If

                    candidates.Sort(Function(left As EvidencePropertyCandidate, right As EvidencePropertyCandidate) left.OriginalIndex.CompareTo(right.OriginalIndex))
                    Dim resultObject As New JObject()
                    For Each candidate As EvidencePropertyCandidate In candidates
                        resultObject(candidate.Name) = candidate.Value
                    Next
                    Return resultObject

                Case JTokenType.Array
                    Dim resultArray As New JArray()
                    For Each item As JToken In DirectCast(source, JArray)
                        Dim child As JToken = BuildEvidenceCoreToken(
                            item,
                            arrayItemPropertyLimit,
                            item IsNot Nothing AndAlso item.Type = JTokenType.Object,
                            omittedPropertyCount,
                            exactScalarCount,
                            maxExactStringChars)
                        If child IsNot Nothing Then
                            resultArray.Add(child)
                        Else
                            ' A synthetic null would be mistaken for exact source evidence.
                            ' Reject this projection; callers keep the full body or a verified reference.
                            Throw New System.InvalidOperationException("An evidence-core array item cannot be represented exactly.")
                        End If
                    Next
                    Return resultArray

                Case JTokenType.String
                    Dim value As System.String = source.Value(Of System.String)()
                    If ShouldOmitEvidenceString(value, maxExactStringChars) Then
                        omittedPropertyCount += 1
                        Return Nothing
                    End If
                    exactScalarCount += 1
                    Return source.DeepClone()

                Case Else
                    If TypeOf source Is JValue Then exactScalarCount += 1
                    Return source.DeepClone()
            End Select
        End Function

        Private Shared Function ShouldOmitEvidenceString(value As System.String, maxExactStringChars As System.Int32) As System.Boolean
            Dim normalized As System.String = If(value, System.String.Empty)
            If normalized.Length > System.Math.Max(0, maxExactStringChars) Then Return True

            Dim lower As System.String = normalized.TrimStart().ToLowerInvariant()
            If lower.StartsWith("http://", System.StringComparison.Ordinal) OrElse
               lower.StartsWith("https://", System.StringComparison.Ordinal) OrElse
               lower.StartsWith("data:", System.StringComparison.Ordinal) Then
                Return True
            End If

            Return False
        End Function

        Private Shared Function GetEvidencePropertyPriority(propertyName As System.String, value As JToken) As System.Int32
            Dim name As System.String = If(propertyName, System.String.Empty).ToLowerInvariant()

            If name = "id" OrElse name = "ref" OrElse name = "key" Then Return 0
            If ContainsAny(name, "title", "subject") OrElse name = "name" OrElse name = "label" Then Return 0

            Dim temporal As System.Boolean = ContainsAny(name, "start", "end", "date", "time")
            If temporal AndAlso ContainsAny(name, "local", "user") Then Return 0
            If ContainsAny(name, "start", "end") AndAlso Not name.Contains("utc") Then Return 1

            ' Exact identity/reference fields are provenance anchors and therefore rank ahead
            ' of descriptive enrichment when high-volume array items must be reduced.
            If name.EndsWith("_id", System.StringComparison.Ordinal) OrElse
               name.EndsWith("_ref", System.StringComparison.Ordinal) OrElse
               name.EndsWith("_key", System.StringComparison.Ordinal) Then
                Return 2
            End If

            If ContainsAny(name, "location", "organizer", "author", "source", "status", "type", "kind") Then Return 2

            If ContainsAny(name,
                           "amount", "value", "result", "answer", "count", "total", "number",
                           "percent", "rate", "version", "code", "score", "state") Then Return 3

            If temporal Then
                If name.Contains("utc") Then Return 6
                Return 3
            End If

            If name = "n" OrElse name = "ok" Then Return 4
            If ContainsAny(name, "id", "ref", "key") Then Return 7

            If value IsNot Nothing AndAlso (value.Type = JTokenType.Object OrElse value.Type = JTokenType.Array) Then Return 4
            Return 5
        End Function

        Private Shared Function ContainsAny(value As System.String, ParamArray needles As System.String()) As System.Boolean
            Dim source As System.String = If(value, System.String.Empty)
            For Each needle As System.String In needles
                If source.IndexOf(needle, System.StringComparison.OrdinalIgnoreCase) >= 0 Then Return True
            Next
            Return False
        End Function

        Private Shared Function GetEvidenceValueLength(value As JToken) As System.Int32
            If value Is Nothing Then Return 0
            If value.Type = JTokenType.String Then
                Return If(value.Value(Of System.String)(), System.String.Empty).Length
            End If
            Return value.ToString(Formatting.None).Length
        End Function

        ''' <summary>
        ''' Determines whether an already returned expansion window is available without
        ''' executing context_expand again. Same-batch results count as available even
        ''' before the next model call; older results count only while their full replay
        ''' body has not been compacted.
        ''' </summary>
        Public Shared Function IsExpansionWindowAvailableWithoutReexecution(producedIteration As System.Int32,
                                                                            currentIteration As System.Int32,
                                                                            wasCompactedForModelReplay As System.Boolean) As System.Boolean
            If currentIteration >= 0 AndAlso producedIteration = currentIteration Then
                Return True
            End If

            Return Not wasCompactedForModelReplay
        End Function

        Public Shared Function AddAvailableRequiredRuntimePrimitives(
            requestedToolNames As System.Collections.Generic.IEnumerable(Of System.String),
            registry As ToolRegistry) As System.Collections.Generic.List(Of System.String)

            Dim result As New System.Collections.Generic.List(Of System.String)()
            Dim seen As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)

            If requestedToolNames IsNot Nothing Then
                For Each requestedToolName As System.String In requestedToolNames
                    Dim normalized As System.String = If(requestedToolName, System.String.Empty).Trim()
                    If normalized <> System.String.Empty AndAlso seen.Add(normalized) Then
                        result.Add(normalized)
                    End If
                Next
            End If

            If registry IsNot Nothing AndAlso
               registry.Contains(ContextExpandTool.ToolName) AndAlso
               seen.Add(ContextExpandTool.ToolName) Then
                result.Add(ContextExpandTool.ToolName)
            End If

            Return result
        End Function

    End Class

End Namespace
