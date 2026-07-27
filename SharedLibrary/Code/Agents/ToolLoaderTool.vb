' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ToolLoaderTool.vb
' Purpose: Internal manifest-only loader tool used to keep the first model pass
'          small. The model first sees only a compact index plus this loader.
'          When it asks for one or more tools by name, the host materializes
'          them and exposes their full definitions on the next iteration.
'
' Architecture:
'  - Lazy-loading trigger: ShouldUseLazyLoading(selectedTools) returns true
'    when tool count exceeds DefaultLazyLoadThreshold (8).
'  - Build(manifests) creates the loader config with tool index.
'  - ExtractRequestedToolNames(arguments) parses tool/tools array from request.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Collections
Imports System.Linq
Imports System.Text
Imports Newtonsoft.Json.Linq

Namespace Agents

    Public NotInheritable Class ToolLoaderTool

        Public Const LoaderToolName As String = "tool_loader"
        Public Const DefaultLazyLoadThreshold As Integer = 8

        Private Sub New()
        End Sub

        Public Shared Function ShouldUseLazyLoading(selectedTools As IEnumerable(Of SharedLibrary.ModelConfig)) As Boolean
            If selectedTools Is Nothing Then Return False

            Dim count As Integer = 0

            For Each tool In selectedTools
                If tool Is Nothing OrElse String.IsNullOrWhiteSpace(tool.ToolName) Then Continue For
                count += 1
                If count > DefaultLazyLoadThreshold Then
                    Return True
                End If
            Next

            Return False
        End Function

        Public Shared Function Build(manifests As IEnumerable(Of ToolManifest)) As SharedLibrary.ModelConfig
            Dim items = If(manifests, Enumerable.Empty(Of ToolManifest)()).
                Where(Function(m)
                          Return m IsNot Nothing AndAlso
                                 Not String.IsNullOrWhiteSpace(m.Name) AndAlso
                                 Not m.Name.Equals(LoaderToolName, StringComparison.OrdinalIgnoreCase)
                      End Function).
                OrderBy(Function(m) m.Name, StringComparer.OrdinalIgnoreCase).
                ToList()

            Dim sb As New StringBuilder()
            sb.Append("tool_loader: Loads one or more allowed tools by exact name so their full schema and instructions become available in a later iteration. ")
            sb.Append("Use this first when you decide a specific tool is needed. ")
            sb.Append("TURN RULE: A newly loaded tool is NOT callable in the same turn in which you load it; its schema only becomes available from the NEXT assistant turn onward. ")
            sb.Append("Never emit tool_loader and a call to a just-loaded tool in the same turn, and never call a tool you have not already loaded in an earlier turn. ")
            sb.Append("Plan ahead: in a single tool_loader call, load ALL tools you expect to need for the whole task (use the 'tools' array), then wait for the next turn and call them. ")
            sb.Append("If a tool is already exposed with its full schema this turn, call it directly and do NOT load it again via tool_loader. ")
            sb.Append("If you decide to use a skill or an agent, load it first; once its instructions list the tools it needs, load ALL of those required tools together in a single tool_loader call before invoking them. ")
            sb.Append("Available tool index:")

            For Each item In items
                sb.AppendLine()
                sb.Append("- ").Append(item.Name)

                If Not String.IsNullOrWhiteSpace(item.Category) Then
                    sb.Append(" [").Append(item.Category.Trim()).Append("]")
                End If

                Dim shortDesc As String = Shrink(item.Description, 120)
                If shortDesc <> "" Then
                    sb.Append(": ").Append(shortDesc)
                End If
            Next

            Dim def As String =
                "{""name"":""tool_loader""," &
                """description"":""Loads one or more allowed tools by exact name so they become available for subsequent tool calls.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """tools"":{""type"":""array"",""items"":{""type"":""string""},""description"":""Exact tool names to load from the available tool index.""}," &
                """tool"":{""type"":""string"",""description"":""Single exact tool name to load (alternative to tools).""}," &
                """reason"":{""type"":""string"",""description"":""Optional short reason for loading the tool or tools.""}" &
                "},""additionalProperties"":false}}"

            Return New SharedLibrary.ModelConfig() With {
                .ToolName = LoaderToolName,
                .ToolInstructionsPrompt = sb.ToString(),
                .ToolDefinition = def,
                .ModelDescription = "Tool Loader (internal)",
                .Tool = True,
                .ToolPriority = -1000,
                .ToolErrorHandling = "skip"
            }
        End Function

        Public Shared Function ExtractRequestedToolNames(arguments As Dictionary(Of String, Object)) As List(Of String)
            Dim result As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            If arguments Is Nothing Then
                Return result.ToList()
            End If

            If arguments.ContainsKey("tool") Then
                AddNames(arguments("tool"), result)
            End If

            If arguments.ContainsKey("tools") Then
                AddNames(arguments("tools"), result)
            End If

            Return result.
                Where(Function(s) Not String.IsNullOrWhiteSpace(s)).
                Select(Function(s) s.Trim()).
                ToList()
        End Function

        Private Shared Sub AddNames(value As Object, target As HashSet(Of String))
            If value Is Nothing OrElse target Is Nothing Then Return

            If TypeOf value Is JValue Then
                AddNames(DirectCast(value, JValue).Value, target)
                Return
            End If

            If TypeOf value Is String Then
                Dim raw As String = DirectCast(value, String).Trim()
                If raw = "" Then Return

                raw = raw.Replace(";", ",").
                          Replace(vbCrLf, ",").
                          Replace(vbCr, ",").
                          Replace(vbLf, ",")

                For Each part In raw.Split({","c}, StringSplitOptions.RemoveEmptyEntries)
                    Dim name As String = part.Trim()
                    If name <> "" Then
                        target.Add(name)
                    End If
                Next

                Return
            End If

            If TypeOf value Is JArray Then
                For Each item As JToken In DirectCast(value, JArray)
                    AddNames(item, target)
                Next
                Return
            End If

            If TypeOf value Is IEnumerable Then
                For Each item As Object In DirectCast(value, IEnumerable)
                    AddNames(item, target)
                Next
                Return
            End If

            Dim fallback As String = value.ToString().Trim()
            If fallback <> "" Then
                target.Add(fallback)
            End If
        End Sub

        Private Shared Function Shrink(value As String, maxLength As Integer) As String
            Dim text As String = If(value, "").Replace(vbCr, " ").Replace(vbLf, " ").Trim()

            While text.Contains("  ")
                text = text.Replace("  ", " ")
            End While

            If text.Length <= maxLength Then
                Return text
            End If

            Return text.Substring(0, maxLength - 1).TrimEnd() & "…"
        End Function

    End Class

End Namespace
