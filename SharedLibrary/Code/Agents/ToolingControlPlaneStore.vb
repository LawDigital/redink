' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ToolingControlPlaneStore.vb
' Purpose: Keeps authoritative workflow-control payloads outside ordinary tool
'          history so budget compaction cannot silently remove active instructions.
' =============================================================================

Option Strict On
Option Explicit On

Namespace Agents

    Public NotInheritable Class ToolingControlPlaneStore

        Private Sub New()
        End Sub

        Public NotInheritable Class PinnedPayload
            Public Property Key As System.String
            Public Property WorkflowId As System.String
            Public Property Kind As System.String
            Public Property Name As System.String
            Public Property FullPayload As System.String
            Public Property TotalChars As System.Int32
            Public Property UpdatedUtc As System.DateTime
        End Class

        Private Shared ReadOnly _activeSkillPayloads As New System.Collections.Concurrent.ConcurrentDictionary(Of System.String, PinnedPayload)(System.StringComparer.Ordinal)

        Public Shared Function PinActiveSkillPayload(workflowId As System.String,
                                                     skillName As System.String,
                                                     fullPayload As System.String) As PinnedPayload
            Dim normalizedWorkflowId As System.String = If(workflowId, System.String.Empty).Trim()
            If normalizedWorkflowId = System.String.Empty Then Return Nothing

            Dim body As System.String = If(fullPayload, System.String.Empty)
            Dim entry As New PinnedPayload() With {
                .Key = "active_skill",
                .WorkflowId = normalizedWorkflowId,
                .Kind = "active_skill",
                .Name = If(skillName, System.String.Empty).Trim(),
                .FullPayload = body,
                .TotalChars = body.Length,
                .UpdatedUtc = System.DateTime.UtcNow
            }

            _activeSkillPayloads(normalizedWorkflowId) = entry
            Return entry
        End Function

        Public Shared Function TryGetActiveSkillPayload(workflowId As System.String,
                                                        ByRef payload As PinnedPayload) As System.Boolean
            payload = Nothing
            Dim normalizedWorkflowId As System.String = If(workflowId, System.String.Empty).Trim()
            If normalizedWorkflowId = System.String.Empty Then Return False
            Return _activeSkillPayloads.TryGetValue(normalizedWorkflowId, payload)
        End Function

        Public Shared Function BuildPromptContextBlock(workflowId As System.String) As System.String
            Dim payload As PinnedPayload = Nothing
            If Not TryGetActiveSkillPayload(workflowId, payload) OrElse payload Is Nothing Then Return System.String.Empty

            Dim sb As New System.Text.StringBuilder()
            sb.AppendLine("[CONTROL_PLANE_CONTEXT]")
            sb.AppendLine("Host-pinned active workflow control payload. This block is authoritative for the current workflow and is replayed independently of ordinary tool-history compaction.")
            sb.AppendLine("- kind: " & payload.Kind)
            sb.AppendLine("- name: " & payload.Name)
            sb.AppendLine("- payloadChars: " & payload.TotalChars.ToString(System.Globalization.CultureInfo.InvariantCulture))
            sb.AppendLine("[CONTROL_PLANE_PAYLOAD]")
            sb.AppendLine(payload.FullPayload)
            sb.AppendLine("[/CONTROL_PLANE_PAYLOAD]")
            sb.Append("[/CONTROL_PLANE_CONTEXT]")
            Return sb.ToString()
        End Function

        Public Shared Sub ClearWorkflow(workflowId As System.String)
            Dim normalizedWorkflowId As System.String = If(workflowId, System.String.Empty).Trim()
            If normalizedWorkflowId = System.String.Empty Then Return
            Dim removed As PinnedPayload = Nothing
            _activeSkillPayloads.TryRemove(normalizedWorkflowId, removed)
        End Sub

    End Class

End Namespace
