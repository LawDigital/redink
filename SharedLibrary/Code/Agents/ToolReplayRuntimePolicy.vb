' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ToolReplayRuntimePolicy.vb
' Purpose: Shared, host-agnostic semantics for model replay retention and runtime
'          context primitives. Content strings never determine retention.
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
                                                    ByRef arrayItemPropertyLimit As System.Int32) As System.Boolean
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
                ToolingConstants.EvidenceReplayPreferredArrayItemProperties,
                10,
                8,
                6,
                4
            }

            For Each propertyLimit As System.Int32 In limits
                Dim omitted As System.Int32 = 0
                Dim exactScalars As System.Int32 = 0
                Dim projected As JToken = BuildEvidenceCoreToken(
                    source,
                    propertyLimit,
                    False,
                    omitted,
                    exactScalars)

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
        ''' Control-plane and deliverable-only envelopes are deliberately excluded so the
        ''' final evidence review applies to source evidence rather than bookkeeping.
        ''' </summary>
        Public Shared Function IsEvidenceCompactedReplay(modelReplayContent As System.String) As System.Boolean
            Dim raw As System.String = If(modelReplayContent, System.String.Empty).Trim()
            If raw = System.String.Empty Then Return False

            Try
                Dim obj As JObject = JObject.Parse(raw)
                Dim controlPinned As JToken = obj("control_plane_pinned")
                If controlPinned IsNot Nothing AndAlso
                   controlPinned.Type = JTokenType.Boolean AndAlso
                   controlPinned.Value(Of System.Boolean)() Then
                    Return False
                End If

                If System.String.IsNullOrWhiteSpace(obj.Value(Of System.String)("result_ref")) Then
                    Return False
                End If

                If obj("evidence_core") IsNot Nothing Then Return True
                If obj("preview") IsNot Nothing Then Return True
                If obj("omitted_window_chars") IsNot Nothing Then Return True

                Return False
            Catch ex As System.Exception
                Return False
            End Try
        End Function

        Private Shared Function BuildEvidenceCoreToken(source As JToken,
                                                       arrayItemPropertyLimit As System.Int32,
                                                       inArrayItem As System.Boolean,
                                                       ByRef omittedPropertyCount As System.Int32,
                                                       ByRef exactScalarCount As System.Int32) As JToken
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
                            exactScalarCount)

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
                            exactScalarCount)
                        If child IsNot Nothing Then
                            resultArray.Add(child)
                        Else
                            ' Preserve array cardinality even when a large scalar item is omitted.
                            resultArray.Add(JValue.CreateNull())
                        End If
                    Next
                    Return resultArray

                Case JTokenType.String
                    Dim value As System.String = source.Value(Of System.String)()
                    If ShouldOmitEvidenceString(value) Then
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

        Private Shared Function ShouldOmitEvidenceString(value As System.String) As System.Boolean
            Dim normalized As System.String = If(value, System.String.Empty)
            If normalized.Length > ToolingConstants.EvidenceReplayMaxExactStringChars Then Return True

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
