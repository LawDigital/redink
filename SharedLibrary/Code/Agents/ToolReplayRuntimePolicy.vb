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

Namespace Agents

    Public Enum ToolReplayRetentionKind
        NormalHistorical = 0
        CurrentTurnCritical = 1
        ControlPlanePinned = 2
    End Enum

    Public NotInheritable Class ToolingRuntimePrimitives

        Private Sub New()
        End Sub

        Public Shared Function IsRequiredRuntimePrimitive(toolName As System.String) As System.Boolean
            Return ContextExpandTool.IsContextExpandTool(toolName)
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
