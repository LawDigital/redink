' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
Option Explicit On
Option Strict On

Namespace Agents
    ' Advisory, per-run metadata only. Never changes tool dispatch, budgets or final answers.
    Public NotInheritable Class RunEffortEstimate
        Private ReadOnly _sink As System.Action(Of System.String)
        Private _observed As System.Int32
        Private _enabled As System.Boolean

        Public Const Instruction As System.String =
            "OPTIONAL WORKLOAD ESTIMATE: In the first report_progress call you already make, you may include effort_class for the ENTIRE requested task: short (a small direct task), medium (several straightforward steps), extensive (substantial research, source comparison or document work), or very_extensive (many dependent steps or extensive research and creation). Consider the full scope, expected research, source volume and verification, not just attachments or the next phase. This is workload, not reasoning intensity or minutes. If unsure, omit it. Never make an extra call or turn just to report or correct this estimate; batch it with your ordinary calls when possible. Do not repeat it, put it in user-facing text, or change your work because of this metadata."

        Public Sub New(sink As System.Action(Of System.String))
            _sink = sink
        End Sub

        Public ReadOnly Property Enabled As System.Boolean
            Get
                Return _enabled
            End Get
        End Property

        Public Function DecorateProgressDefinition(definition As System.String) As System.String
            If _sink Is Nothing Then Return definition
            Try
                Dim root As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(definition)
                Dim parameters As Newtonsoft.Json.Linq.JObject = TryCast(root("parameters"), Newtonsoft.Json.Linq.JObject)
                Dim properties As Newtonsoft.Json.Linq.JObject = If(parameters Is Nothing, Nothing, TryCast(parameters("properties"), Newtonsoft.Json.Linq.JObject))
                If properties Is Nothing Then Return definition
                properties("effort_class") = New Newtonsoft.Json.Linq.JObject(
                    New Newtonsoft.Json.Linq.JProperty("type", "string"),
                    New Newtonsoft.Json.Linq.JProperty("enum", New Newtonsoft.Json.Linq.JArray("short", "medium", "extensive", "very_extensive")),
                    New Newtonsoft.Json.Linq.JProperty("description", "Optional initial workload class for the whole task; include only in your first ordinary progress call. Omit if uncertain. Never add a call just for this field."))
                Dim decorated As System.String = root.ToString(Newtonsoft.Json.Formatting.None)
                _enabled = True
                Return decorated
            Catch ex As System.Exception
                System.Diagnostics.Trace.TraceWarning("Optional workload metadata was not enabled: " & ex.Message)
                Return definition
            End Try
        End Function

        Public Shared Function Normalize(value As System.String) As System.String
            Dim canonical As System.String = If(value, System.String.Empty).Trim().ToLowerInvariant().Replace("-", "_").Replace(" ", "_")
            Select Case canonical
                Case "short", "medium", "extensive", "very_extensive"
                    Return canonical
                Case Else
                    Return System.String.Empty
            End Select
        End Function

        Public Sub Observe(arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object))
            If Not _enabled OrElse System.Threading.Interlocked.Exchange(_observed, 1) <> 0 Then Return
            Try
                ' Only the first ordinary progress call may classify the run. Later phase updates
                ' must not retroactively label a task using its already observed duration.
                Dim raw As System.Object = Nothing
                If arguments IsNot Nothing Then
                    For Each entry As System.Collections.Generic.KeyValuePair(Of System.String, System.Object) In arguments
                        If System.String.Equals(entry.Key, "effort_class", System.StringComparison.OrdinalIgnoreCase) Then
                            raw = entry.Value
                            Exit For
                        End If
                    Next
                End If
                Dim text As System.String = TryCast(raw, System.String)
                Dim token As Newtonsoft.Json.Linq.JValue = TryCast(raw, Newtonsoft.Json.Linq.JValue)
                If token IsNot Nothing AndAlso token.Type = Newtonsoft.Json.Linq.JTokenType.String Then text = CType(token, System.String)
                Dim canonical As System.String = Normalize(text)
                If canonical.Length = 0 Then
                    System.Diagnostics.Trace.TraceInformation("No valid initial workload class; retaining the existing duration estimate.")
                    Return
                End If
                _sink.Invoke(canonical)
            Catch ex As System.Exception
                ' Optional telemetry must never fail an otherwise valid report_progress/tool turn.
                System.Diagnostics.Trace.TraceWarning("Optional workload estimate was ignored: " & ex.Message)
            End Try
        End Sub
    End Class
End Namespace
