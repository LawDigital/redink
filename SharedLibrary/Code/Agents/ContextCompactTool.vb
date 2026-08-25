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
                """description"":""Voluntarily compact older tool results out of the active context to free space when you no longer need them in full. The full text stays retrievable via context_expand using each result's result_ref. Optionally set keep_recent to control how many of the most recent results remain fully visible (default 0).""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """keep_recent"":{""type"":""integer"",""description"":""How many of the most recent tool results to keep fully visible; older ones are moved to the drawer. Defaults to 0.""}" &
                "},""required"":[],""additionalProperties"":false}}"

            Return New SharedLibrary.ModelConfig() With {
                .ToolName = ToolName,
                .ToolDefinition = def,
                .ToolInstructionsPrompt = ToolName & ": Compact older tool results out of the active context when they are no longer needed in full (they remain retrievable via context_expand). Optional keep_recent controls how many recent results stay fully visible (default 0).",
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
                New JProperty("note", "Older tool results will be moved out of the active context on the next turn. Use context_expand with a result_ref to read any of them again.")
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
