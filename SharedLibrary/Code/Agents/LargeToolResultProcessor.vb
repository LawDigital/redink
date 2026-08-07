' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: LargeToolResultProcessor.vb
' Purpose: Reference-by-default normalization of large tool results. Applied at
'          the single dispatch choke point in each host so every heavy producer
'          (web, M365, Excel/Word/PDF/Office readers, knowledge stores) behaves
'          identically without per-producer edits.
' =============================================================================

Option Strict On
Option Explicit On

Imports Newtonsoft.Json.Linq

Namespace Agents

    Public NotInheritable Class LargeToolResultProcessor

        Private Sub New()
        End Sub

        ''' <summary>Builds a compact reference envelope for an oversized result.</summary>
        ''' <param name="workflowId">Current workflow id (for later cleanup).</param>
        ''' <param name="toolName">Producing tool name.</param>
        ''' <param name="fullContent">Full result body.</param>
        ''' <param name="thresholdChars">Above this size, the body is stored by reference.</param>
        ''' <param name="previewChars">Preview length kept inline.</param>
        ''' <param name="modelSummary">Optional summary already produced by the tool.</param>
        ''' <returns>Compact JSON envelope, or the original body if under threshold.</returns>
        Public Shared Function NormalizeIfLarge(workflowId As String,
                                                toolName As String,
                                                fullContent As String,
                                                thresholdChars As Integer,
                                                previewChars As Integer,
                                                Optional modelSummary As String = "",
                                                Optional deliverableToolNames As IReadOnlyCollection(Of String) = Nothing) As String
            Dim body As String = If(fullContent, "")
            If body.Length <= thresholdChars Then
                Return body
            End If

            ' Never normalize deliverable/reference-producing tools: the completion
            ' gate reads top-level path/output_reference fields that an envelope would
            ' hide inside a truncated preview. Optional set; skipped safely if absent.
            If Not String.IsNullOrWhiteSpace(toolName) AndAlso
               deliverableToolNames IsNot Nothing AndAlso
               deliverableToolNames.Contains(toolName.Trim()) Then
                Return body
            End If

            ' Host-agnostic safety net: if the body already exposes a top-level
            ' reference field the completion gate relies on, leave it untouched so
            ' deliverable detection works regardless of tool name or model strength.
            If HasTopLevelReference(body) Then
                Return body
            End If

            Dim stored = ToolResultStore.Put(workflowId, toolName, body)

            Dim previewLength As Integer = Math.Min(Math.Max(previewChars, 0), body.Length)
            Dim preview As String = body.Substring(0, previewLength)

            Dim summary As String = If(modelSummary, "").Trim()
            If summary = "" Then
                summary = ExtractEmbeddedSummary(body)
            End If

            Dim envelope As New JObject(
                New JProperty("ok", True),
                New JProperty("tool", If(toolName, "")),
                New JProperty("summary", summary),
                New JProperty("result_ref", stored.Ref),
                New JProperty("preview", preview),
                New JProperty("total_chars", body.Length),
                New JProperty("returned_chars", previewLength),
                New JProperty("truncated", True),
                New JProperty("next_offset", previewLength),
                New JProperty("continuation",
                    "Full content stored by reference. To read more, call context_expand with result_ref='" &
                    stored.Ref & "' and a start_char/max_chars window."))

            Return envelope.ToString(Newtonsoft.Json.Formatting.None)
        End Function

        Private Shared Function HasTopLevelReference(body As String) As Boolean
            Try
                Dim tok As JToken = JToken.Parse(body)
                Dim obj As JObject = TryCast(tok, JObject)
                If obj Is Nothing Then Return False

                Dim refFields As String() = New String() {
                    "path", "saved_path", "output_path", "file_path",
                    "output_reference", "reference", "memory_key", "memoryKey",
                    "outputArtifactRef", "output_artifact_ref", "artifact_ref",
                    "outputFilePath", "output_file_path"}

                For Each field As String In refFields
                    Dim val As JToken = obj(field)
                    If val IsNot Nothing AndAlso
                       val.Type <> JTokenType.Null AndAlso
                       Not String.IsNullOrWhiteSpace(val.ToString()) Then
                        Return True
                    End If
                Next

                Return False
            Catch
                Return False
            End Try
        End Function

        Private Shared Function ExtractEmbeddedSummary(body As String) As String
            Try
                Dim tok As JToken = JToken.Parse(body)
                If TypeOf tok Is JObject Then
                    Dim s As String = DirectCast(tok, JObject).Value(Of String)("summary")
                    If Not String.IsNullOrWhiteSpace(s) Then Return s
                End If
            Catch
            End Try
            Return ""
        End Function

    End Class

End Namespace
