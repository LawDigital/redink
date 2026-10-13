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

        ''' <summary>
        ''' Builds a canonical identity for the exact stored character window that a
        ''' context_expand request would return. Defaults, clamping, negative starts and
        ''' end-of-content normalization intentionally match Execute().
        ''' </summary>
        Public Shared Function TryBuildCanonicalWindowKey(arguments As System.Collections.Generic.Dictionary(Of System.String, System.Object),
                                                          ByRef windowKey As System.String,
                                                          Optional context As Global.SharedLibrary.SharedLibrary.SharedContext.ISharedContext = Nothing) As System.Boolean
            windowKey = System.String.Empty

            Dim ref As System.String = GetString(arguments, "result_ref")
            If System.String.IsNullOrWhiteSpace(ref) Then Return False

            Dim stored As ToolResultStore.StoredResult = Nothing
            If Not ToolResultStore.TryGetForWorkflow(
                ref,
                WorkflowContinuity.CurrentWorkflowId,
                stored) OrElse stored Is Nothing Then Return False

            Dim body As System.String = If(stored.FullContent, System.String.Empty)
            Dim requestedStart As System.Int32 = GetInt(arguments, "start_char", 0)
            Dim requestedMax As System.Int32 = NormalizeMaxChars(GetInt(arguments, "max_chars", GetDefaultWindowChars(context)), context)
            Dim normalizedStart As System.Int32
            Dim returnedChars As System.Int32
            Dim nextOffset As System.Int32
            If Not TryNormalizeWindow(body, requestedStart, requestedMax, normalizedStart, returnedChars, nextOffset) Then Return False

            windowKey = BuildCanonicalWindowKey(ref, normalizedStart, nextOffset)
            Return True
        End Function

        ''' <summary>
        ''' Extracts the canonical window identity from a successful context_expand result.
        ''' This lets the hosts compare a requested window with the exact content that was
        ''' previously returned, without relying on provider-specific call JSON.
        ''' </summary>
        Public Shared Function TryBuildCanonicalWindowKeyFromResponse(responseText As System.String,
                                                                      ByRef windowKey As System.String) As System.Boolean
            windowKey = System.String.Empty
            If System.String.IsNullOrWhiteSpace(responseText) Then Return False

            Dim obj As Newtonsoft.Json.Linq.JObject
            Try
                obj = Newtonsoft.Json.Linq.JObject.Parse(responseText)
            Catch
                Return False
            End Try

            Dim okToken As Newtonsoft.Json.Linq.JToken = obj("ok")
            If okToken Is Nothing OrElse
               okToken.Type <> Newtonsoft.Json.Linq.JTokenType.Boolean OrElse
               Not okToken.Value(Of System.Boolean)() Then
                Return False
            End If

            Dim ref As System.String = If(obj.Value(Of System.String)("result_ref"), System.String.Empty).Trim()
            If ref = System.String.Empty Then Return False

            Dim normalizedStart As System.Int32
            Dim nextOffset As System.Int32
            If Not TryParseInt32Token(obj("start_char"), normalizedStart) OrElse
               Not TryParseInt32Token(obj("next_offset"), nextOffset) Then
                Return False
            End If

            If normalizedStart < 0 OrElse nextOffset < normalizedStart Then Return False

            windowKey = BuildCanonicalWindowKey(ref, normalizedStart, nextOffset)
            Return True
        End Function

        Private Shared Function BuildCanonicalWindowKey(ref As System.String,
                                                        normalizedStart As System.Int32,
                                                        nextOffset As System.Int32) As System.String
            Return If(ref, System.String.Empty).Trim().ToLowerInvariant() & "|" &
                   normalizedStart.ToString(System.Globalization.CultureInfo.InvariantCulture) & "|" &
                   nextOffset.ToString(System.Globalization.CultureInfo.InvariantCulture)
        End Function

        Private Shared Function GetMaximumWindowChars(context As Global.SharedLibrary.SharedLibrary.SharedContext.ISharedContext) As System.Int32
            Return System.Math.Max(1, If(context Is Nothing, ToolingConstants.ContextExpandMaxChars, context.INI_ContextExpandMaxChars))
        End Function

        Private Shared Function GetDefaultWindowChars(context As Global.SharedLibrary.SharedLibrary.SharedContext.ISharedContext) As System.Int32
            Return NormalizeMaxChars(If(context Is Nothing, ToolingConstants.ContextExpandDefaultChars, context.INI_ContextExpandDefaultChars), context)
        End Function

        Private Shared Function NormalizeMaxChars(value As System.Int32, context As Global.SharedLibrary.SharedLibrary.SharedContext.ISharedContext) As System.Int32
            Return System.Math.Min(System.Math.Max(value, 1), GetMaximumWindowChars(context))
        End Function

        Public Shared Function TryNormalizeWindow(body As System.String, requestedStart As System.Int32,
                                                   maximumChars As System.Int32, ByRef normalizedStart As System.Int32,
                                                   ByRef returnedChars As System.Int32, ByRef nextOffset As System.Int32) As System.Boolean
            normalizedStart = System.Math.Max(0, System.Math.Min(requestedStart, body.Length))
            If normalizedStart > 0 AndAlso normalizedStart < body.Length AndAlso
               System.Char.IsLowSurrogate(body(normalizedStart)) AndAlso System.Char.IsHighSurrogate(body(normalizedStart - 1)) Then
                normalizedStart -= 1
            End If
            returnedChars = System.Math.Min(maximumChars, body.Length - normalizedStart)
            nextOffset = normalizedStart + returnedChars
            If nextOffset > normalizedStart AndAlso nextOffset < body.Length AndAlso
               System.Char.IsHighSurrogate(body(nextOffset - 1)) AndAlso System.Char.IsLowSurrogate(body(nextOffset)) Then
                If returnedChars = 1 Then Return False
                returnedChars -= 1
                nextOffset -= 1
            End If
            Return True
        End Function

        Private Shared Function TryParseInt32Token(token As Newtonsoft.Json.Linq.JToken,
                                                   ByRef value As System.Int32) As System.Boolean
            value = 0
            If token Is Nothing OrElse token.Type = Newtonsoft.Json.Linq.JTokenType.Null Then Return False

            Return System.Int32.TryParse(
                token.ToString(),
                System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture,
                value)
        End Function

        Public Shared Function Build(Optional context As Global.SharedLibrary.SharedLibrary.SharedContext.ISharedContext = Nothing) As SharedLibrary.ModelConfig
            Dim maxWindow As System.Int32 = GetMaximumWindowChars(context)
            Dim defaultWindow As System.Int32 = GetDefaultWindowChars(context)
            Dim def As String =
                "{""name"":""" & ToolName & """," &
                """description"":""Retrieve a character window from a large tool result that was stored by reference. Large results are replaced in context by a short 'result_ref' plus a preview; call this to read more of that stored content on demand.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """result_ref"":{""type"":""string"",""description"":""The result_ref returned in a prior tool result envelope.""}," &
                """start_char"":{""type"":""integer"",""description"":""Zero-based UTF-16 character-unit offset to start reading from. Defaults to 0; the host keeps surrogate pairs intact and reports normalized ranges.""}," &
                """max_chars"":{""type"":""integer"",""description"":""Maximum number of UTF-16 character units to return (1-" & maxWindow.ToString(System.Globalization.CultureInfo.InvariantCulture) & "). Defaults to " & defaultWindow.ToString(System.Globalization.CultureInfo.InvariantCulture) & ".""}" &
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

        Public Shared Function Execute(arguments As Dictionary(Of String, Object), Optional context As Global.SharedLibrary.SharedLibrary.SharedContext.ISharedContext = Nothing) As String
            Dim ref As String = GetString(arguments, "result_ref")
            Dim startChar As Integer = GetInt(arguments, "start_char", 0)
            Dim maxChars As Integer = NormalizeMaxChars(GetInt(arguments, "max_chars", GetDefaultWindowChars(context)), context)

            Dim stored As ToolResultStore.StoredResult = Nothing
            If Not ToolResultStore.TryGetForWorkflow(
                ref,
                WorkflowContinuity.CurrentWorkflowId,
                stored) Then
                Return New JObject(
                    New JProperty("ok", False),
                    New JProperty("error", New JObject(
                        New JProperty("code", "unknown_result_ref"),
                        New JProperty("message", "No stored result found for result_ref '" & If(ref, "") & "'.")))
                ).ToString(Newtonsoft.Json.Formatting.None)
            End If

            ' Capture the stored body once so the returned window and navigation metadata
            ' are derived from the same immutable snapshot of this stored result.
            Dim body As System.String = If(stored.FullContent, System.String.Empty)
            Dim totalChars As System.Int32 = body.Length
            Dim normalizedStart As System.Int32
            Dim returnedChars As System.Int32
            Dim nextOffset As System.Int32
            If Not TryNormalizeWindow(body, startChar, maxChars, normalizedStart, returnedChars, nextOffset) Then
                Return New JObject(
                    New JProperty("ok", False),
                    New JProperty("error", New JObject(
                        New JProperty("code", "window_limit_too_small"),
                        New JProperty("message", "This window would split a UTF-16 surrogate pair. Request at least two characters and configure ContextExpandMaxChars accordingly; the original is retained.")))
                ).ToString(Newtonsoft.Json.Formatting.None)
            End If
            Dim window As System.String = body.Substring(normalizedStart, returnedChars)

            Return New JObject(
                New JProperty("ok", True),
                New JProperty("result_ref", ref),
                New JProperty("tool", stored.ToolName),
                New JProperty("content_complete", normalizedStart = 0 AndAlso nextOffset = totalChars),
                New JProperty("window_complete", True),
                New JProperty("is_semantic_summary", False),
                New JProperty("sha256", stored.ContentHash),
                New JProperty("hash_scope", "stored_tool_result"),
                New JProperty("character_unit", "utf16_code_units"),
                New JProperty("requested_start_char", startChar),
                New JProperty("representation", "source_window"),
                New JProperty("full_content_available", True),
                New JProperty("expires_utc", stored.ExpiresUtc),
                New JProperty("retrieval_required", "This is only the disclosed range. Whole-source tasks must inspect all required ranges; do not infer unread content. For long sources prefer complete file export and isolated file analysis rather than source dumps into the parent context."),
                New JProperty("content_window", window),
                New JProperty("start_char", normalizedStart),
                New JProperty("returned_chars", returnedChars),
                New JProperty("total_chars", totalChars),
                New JProperty("next_offset", nextOffset),
                New JProperty("truncated", normalizedStart > 0 OrElse nextOffset < totalChars)
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
