' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ContextCompactTool.vb
' Purpose: Model-driven request to compact older tool results out of the active
'          context. The full bodies remain retrievable via context_expand. Shared
'          by Outlook and Word so both expose identical behavior.
' =============================================================================

Option Strict On
Option Explicit On

Imports Newtonsoft.Json.Linq

Namespace Agents

    Public NotInheritable Class ContextCompactTool

        Private Sub New()
        End Sub

        Public Const ToolName As String = "context_compact"

        Public Shared Function IsContextCompactTool(name As String) As Boolean
            Return Not String.IsNullOrWhiteSpace(name) AndAlso
                   name.Trim().Equals(ToolName, StringComparison.OrdinalIgnoreCase)
        End Function

        Public Shared Function Build() As SharedLibrary.ModelConfig
            Dim def As String =
                "{""name"":""" & ToolName & """," &
                """description"":""Request one compaction of existing older tool results when you no longer need them in full. This is not a persistent retention setting; do not alternate compaction with reloading the same evidence. The full text stays retrievable via context_expand using each result's result_ref. Optionally set keep_recent to control how many of the most recent results remain fully visible (default 0).""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """keep_recent"":{""type"":""integer"",""description"":""How many of the most recent tool results to keep fully visible; older ones are moved to the drawer. Defaults to 0.""}" &
                "},""required"":[],""additionalProperties"":false}}"

            Return New SharedLibrary.ModelConfig() With {
                .ToolName = ToolName,
                .ToolDefinition = def,
                .ToolInstructionsPrompt = ToolName & ": Request one compaction of existing older tool results when they are no longer needed in full; keep verified findings and exact supporting evidence before shelving sources (they remain retrievable via context_expand). Optional keep_recent controls how many recent results stay fully visible (default 0).",
                .ModelDescription = "Context compactor (internal)",
                .Tool = True,
                .ToolPriority = 939,
                .ToolErrorHandling = "skip"
            }
        End Function

        Public Shared Function Execute(arguments As Dictionary(Of String, Object), workflowId As String) As String
            Dim keepRecent As Integer = GetInt(arguments, "keep_recent", 0)
            keepRecent = Math.Max(0, keepRecent)

            If String.IsNullOrWhiteSpace(workflowId) Then
                Return New JObject(
                    New JProperty("ok", False),
                    New JProperty("error", New JObject(
                        New JProperty("code", "no_active_workflow"),
                        New JProperty("message", "No active workflow is available to compact.")))
                ).ToString(Newtonsoft.Json.Formatting.None)
            End If

            ToolResultStore.RequestCompaction(workflowId, keepRecent)

            Return New JObject(
                New JProperty("ok", True),
                New JProperty("tool", ToolName),
                New JProperty("keep_recent", keepRecent),
                New JProperty("note", "One compaction was requested for the next replay rebuild, not proof that it succeeded. It does not apply to future results. Do not repeatedly shelve and reload the same evidence; prefer an exposed full-source analyzer accepting input_result_ref when appropriate. The host replaces content only when the complete stored result is retrievable; on storage failure it keeps the full body. Every resulting stub is incomplete: retrieve all task-required ranges with context_expand, or use an exposed isolated analyzer accepting input_result_ref, before relying on omitted content.")
            ).ToString(Newtonsoft.Json.Formatting.None)
        End Function

        Private Shared Function GetInt(args As Dictionary(Of String, Object), key As String, fallback As Integer) As Integer
            If args Is Nothing OrElse Not args.ContainsKey(key) OrElse args(key) Is Nothing Then Return fallback
            Dim v As Integer
            If Integer.TryParse(Convert.ToString(args(key)), v) Then Return v
            Return fallback
        End Function

    End Class

End Namespace
