' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ContextExpandTool.vb
' Purpose: Model-driven retrieval of a window from a stored large tool result.
'          Shared by Outlook and Word so both expose identical behavior.
' =============================================================================

Option Strict On
Option Explicit On

Imports Newtonsoft.Json.Linq

Namespace Agents

    Public NotInheritable Class ContextExpandTool

        Private Sub New()
        End Sub

        Public Const ToolName As String = "context_expand"

        Public Shared Function IsContextExpandTool(name As String) As Boolean
            Return Not String.IsNullOrWhiteSpace(name) AndAlso
                   name.Trim().Equals(ToolName, StringComparison.OrdinalIgnoreCase)
        End Function

        Public Shared Function Build() As SharedLibrary.ModelConfig
            Dim def As String =
                "{""name"":""" & ToolName & """," &
                """description"":""Retrieve a character window from a large tool result that was stored by reference. Large results are replaced in context by a short 'result_ref' plus a preview; call this to read more of that stored content on demand.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """result_ref"":{""type"":""string"",""description"":""The result_ref returned in a prior tool result envelope.""}," &
                """start_char"":{""type"":""integer"",""description"":""Zero-based character offset to start reading from. Defaults to 0.""}," &
                """max_chars"":{""type"":""integer"",""description"":""Maximum number of characters to return (500-100000). Defaults to 8000.""}" &
                "},""required"":[""result_ref""],""additionalProperties"":false}}"

            Return New SharedLibrary.ModelConfig() With {
                .ToolName = ToolName,
                .ToolDefinition = def,
                .ToolInstructionsPrompt = ToolName & ": Read more of a large tool result that was stored by reference. Pass the 'result_ref' from an earlier result envelope, with optional start_char and max_chars, to page through the full content.",
                .ModelDescription = "Large-result expander (internal)",
                .Tool = True,
                .ToolPriority = 938,
                .ToolErrorHandling = "skip"
            }
        End Function

        Public Shared Function Execute(arguments As Dictionary(Of String, Object)) As String
            Dim ref As String = GetString(arguments, "result_ref")
            Dim startChar As Integer = GetInt(arguments, "start_char", 0)
            Dim maxChars As Integer = GetInt(arguments, "max_chars", 8000)
            maxChars = Math.Min(Math.Max(maxChars, 500), 100000)

            Dim stored As ToolResultStore.StoredResult = Nothing
            If Not ToolResultStore.TryGet(ref, stored) Then
                Return New JObject(
                    New JProperty("ok", False),
                    New JProperty("error", New JObject(
                        New JProperty("code", "unknown_result_ref"),
                        New JProperty("message", "No stored result found for result_ref '" & If(ref, "") & "'.")))
                ).ToString(Newtonsoft.Json.Formatting.None)
            End If

            Dim window As String = ToolResultStore.GetWindow(ref, startChar, maxChars)
            Dim nextOffset As Integer = Math.Min(startChar + window.Length, stored.TotalChars)

            Return New JObject(
                New JProperty("ok", True),
                New JProperty("result_ref", ref),
                New JProperty("tool", stored.ToolName),
                New JProperty("content_window", window),
                New JProperty("start_char", startChar),
                New JProperty("returned_chars", window.Length),
                New JProperty("total_chars", stored.TotalChars),
                New JProperty("next_offset", nextOffset),
                New JProperty("truncated", nextOffset < stored.TotalChars)
            ).ToString(Newtonsoft.Json.Formatting.None)
        End Function

        Private Shared Function GetString(args As Dictionary(Of String, Object), key As String) As String
            If args Is Nothing OrElse Not args.ContainsKey(key) OrElse args(key) Is Nothing Then Return ""
            Return Convert.ToString(args(key))
        End Function

        Private Shared Function GetInt(args As Dictionary(Of String, Object), key As String, fallback As Integer) As Integer
            If args Is Nothing OrElse Not args.ContainsKey(key) OrElse args(key) Is Nothing Then Return fallback
            Dim v As Integer
            If Integer.TryParse(Convert.ToString(args(key)), v) Then Return v
            Return fallback
        End Function

    End Class

End Namespace
