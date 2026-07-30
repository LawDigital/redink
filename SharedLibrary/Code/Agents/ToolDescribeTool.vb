' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ToolDescribeTool.vb
' Purpose: Internal, inspection-only "tool_describe" tool. Returns the full
'          parameter schema and usage instructions for one or more host tools
'          WITHOUT exposing them for calling. Lets the model (in particular the
'          skill-author skill) compare overlapping tools (e.g. several Word
'          editors) and pick the right one before loading it via tool_loader.
'
' Architecture:
'  - Reads from the per-run authoritative ToolRegistry, so newly registered
'    tools are described automatically without any per-tool wiring.
'  - Materializing a tool here only reads its ModelConfig; it does NOT add it to
'    the model's callable set, so there is no tool_loader-style turn delay.
'  - Filter by exact name(s) ('tool'/'tools') or a name prefix/substring
'    ('prefix'); with no arguments it returns a compact index of all tools.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Collections
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Agents

    Public NotInheritable Class ToolDescribeTool

        Public Const ToolName As String = "tool_describe"

        Private Sub New()
        End Sub

        Public Shared Function IsDescribeTool(name As String) As Boolean
            Return Not String.IsNullOrWhiteSpace(name) AndAlso
                   name.Trim().Equals(ToolName, StringComparison.OrdinalIgnoreCase)
        End Function

        Public Shared Function Build() As SharedLibrary.ModelConfig
            Dim def As String =
                "{""name"":""" & ToolName & """," &
                """description"":""Return the full parameter schema and usage instructions for one or more host tools WITHOUT making them callable. Use it to compare overlapping tools (for example several Word editors) and pick the right one before loading it with tool_loader. Filter by exact name(s) or by a name prefix/substring; call with no arguments to get a compact index of every available tool.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """tool"":{""type"":""string"",""description"":""Single exact tool name to describe.""}," &
                """tools"":{""type"":""array"",""items"":{""type"":""string""},""description"":""Several exact tool names to describe.""}," &
                """prefix"":{""type"":""string"",""description"":""Name prefix or substring to match a family of tools, for example 'word' or 'worddoc'.""}" &
                "},""additionalProperties"":false}}"

            Return New SharedLibrary.ModelConfig() With {
                .ToolName = ToolName,
                .ToolDefinition = def,
                .ToolInstructionsPrompt = ToolName & ": Inspect tools' full parameter schemas and instructions without exposing them for calling. Use it to choose between overlapping tools; then load the chosen tool with tool_loader before calling it.",
                .ModelDescription = "Tool schema inspector (internal)",
                .Tool = True,
                .ToolPriority = 939,
                .ToolErrorHandling = "skip"
            }
        End Function

        ''' <summary>
        ''' Executes a tool_describe call against the given authoritative registry.
        ''' Returns a JSON string suitable for the tool response.
        ''' </summary>
        Public Shared Function Execute(arguments As IDictionary(Of String, Object),
                                       registry As ToolRegistry) As String
            Try
                If registry Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {Key .error = "registry_unavailable"})
                End If

                Dim allManifests = registry.ListManifests()

                Dim selectedNames As New List(Of String)()
                Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

                Dim singleName As String = GetStr(arguments, "tool")
                If Not String.IsNullOrWhiteSpace(singleName) Then
                    Dim collected As New List(Of String)()
                    CollectNames(singleName, collected)
                    For Each n In collected
                        If seen.Add(n) Then selectedNames.Add(n)
                    Next
                End If

                If arguments IsNot Nothing AndAlso arguments.ContainsKey("tools") Then
                    Dim collected As New List(Of String)()
                    CollectNames(arguments("tools"), collected)
                    For Each n In collected
                        If seen.Add(n) Then selectedNames.Add(n)
                    Next
                End If

                Dim prefix As String = GetStr(arguments, "prefix")
                If Not String.IsNullOrWhiteSpace(prefix) Then
                    Dim p As String = prefix.Trim()
                    For Each m In allManifests
                        If m Is Nothing OrElse String.IsNullOrWhiteSpace(m.Name) Then Continue For
                        If m.Name.IndexOf(p, StringComparison.OrdinalIgnoreCase) >= 0 Then
                            If seen.Add(m.Name) Then selectedNames.Add(m.Name)
                        End If
                    Next
                End If

                ' No filter provided: return a compact index so the model can choose.
                If selectedNames.Count = 0 Then
                    Dim idx As New JArray()
                    For Each m In allManifests
                        If m Is Nothing OrElse String.IsNullOrWhiteSpace(m.Name) Then Continue For
                        Dim o As New JObject()
                        o("name") = m.Name
                        o("category") = If(m.Category, "")
                        o("description") = If(m.Description, "")
                        idx.Add(o)
                    Next

                    Dim indexResult As New JObject()
                    indexResult("index") = idx
                    indexResult("count") = idx.Count
                    indexResult("hint") = "Call tool_describe again with 'tool', 'tools', or 'prefix' to get full parameter schemas."
                    Return indexResult.ToString(Formatting.None)
                End If

                Dim described As New JArray()
                Dim notFound As New JArray()

                For Each n In selectedNames
                    Dim mc As SharedLibrary.ModelConfig = registry.Get(n)
                    If mc Is Nothing Then
                        notFound.Add(n)
                        Continue For
                    End If

                    Dim o As New JObject()
                    o("name") = If(mc.ToolName, n)
                    o("description") = If(mc.ModelDescription, "")
                    o("instructions") = If(mc.ToolInstructionsPrompt, "")

                    If Not String.IsNullOrWhiteSpace(mc.ToolDefinition) Then
                        Try
                            Dim defObj As JObject = JObject.Parse(mc.ToolDefinition)
                            If defObj("description") IsNot Nothing Then
                                o("summary") = defObj("description").ToString()
                            End If
                            If defObj("parameters") IsNot Nothing Then
                                o("parameters") = defObj("parameters")
                            End If
                        Catch
                            o("raw_definition") = mc.ToolDefinition
                        End Try
                    End If

                    described.Add(o)
                Next

                Dim result As New JObject()
                result("described") = described
                result("not_found") = notFound
                result("count") = described.Count
                Return result.ToString(Formatting.None)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {Key .error = "tool_describe_failed", Key .message = ex.Message})
            End Try
        End Function

        Private Shared Function GetStr(args As IDictionary(Of String, Object), key As String) As String
            If args Is Nothing OrElse Not args.ContainsKey(key) Then Return ""
            Dim v As Object = args(key)
            If v Is Nothing Then Return ""
            Return v.ToString()
        End Function

        Private Shared Sub CollectNames(value As Object, target As List(Of String))
            If value Is Nothing OrElse target Is Nothing Then Return

            If TypeOf value Is String Then
                Dim raw As String = DirectCast(value, String)
                For Each part In raw.Split({","c, ";"c}, StringSplitOptions.RemoveEmptyEntries)
                    Dim n As String = part.Trim()
                    If n <> "" Then target.Add(n)
                Next
                Return
            End If

            If TypeOf value Is JValue Then
                CollectNames(DirectCast(value, JValue).Value, target)
                Return
            End If

            If TypeOf value Is JArray Then
                For Each item As JToken In DirectCast(value, JArray)
                    CollectNames(item.ToString(), target)
                Next
                Return
            End If

            If TypeOf value Is IEnumerable Then
                For Each item As Object In DirectCast(value, IEnumerable)
                    CollectNames(item, target)
                Next
                Return
            End If

            Dim fallback As String = value.ToString().Trim()
            If fallback <> "" Then target.Add(fallback)
        End Sub

    End Class

End Namespace
