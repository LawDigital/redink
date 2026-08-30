' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: SubAgentRuntimeHardening.vb
' Purpose: Normalizes and validates sub-agent final outputs to ensure
'          semantic non-emptiness and correct JSON/text envelope handling.
'
' Envelope Handling:
'  - Recognizes {summary, result} envelope format (preserved as-is).
'  - Fallback: direct JSON objects/arrays treated as structured results.
'  - Fallback: plain text (when jsonRequired=false) with auto-generated summary.
'  - Empty check: no usable output triggers agent_empty_result error.
'
' Normalized Output:
'  - Summary (short textual description).
'  - Result (JToken: object, array, or string).
'  - ResultKind (envelope, json_object, json_array, text, error).
'  - Error envelope (when IsError=true).
' =============================================================================

Option Strict On
Option Explicit On

Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Agents

    Public NotInheritable Class SubAgentRuntimeHardening

        Private Sub New()
        End Sub

        Public Const EmptyResultSummary As String = "Sub-agent returned no usable result."
        Public Const EmptyResultCode As String = "agent_empty_result"
        Public Const EmptyResultPhase As String = "final_output_parse"
        Public Const EmptyResultMessage As String = "Sub-agent returned no usable final result."

        Public Const ModelEmptyResponseStatus As String = "blocked"
        Public Const ModelEmptyResponseCode As String = "model_empty_response"
        Public Const ModelEmptyResponsePhase As String = "main_loop"
        Public Const ModelEmptyResponseMessage As String = "The model returned no tool calls and no final answer."

        Private Const StructuredJsonFallbackSummary As String = "Sub-agent returned structured JSON."

        Public NotInheritable Class NormalizedEnvelope
            Public Property Summary As String
            Public Property Result As JToken
            Public Property ResultKind As String
            Public Property RawLength As Integer
            Public Property [Error] As JObject

            Public ReadOnly Property IsError As Boolean
                Get
                    If String.Equals(ResultKind, "error", StringComparison.OrdinalIgnoreCase) Then Return True
                    Return Not String.IsNullOrWhiteSpace(GetErrorCode())
                End Get
            End Property

            Public Function GetErrorCode() As String
                If [Error] Is Nothing Then Return ""
                Return If([Error].Value(Of String)("code"), "")
            End Function

            Public Function ToJObject() As JObject
                Dim obj As New JObject()

                obj("summary") = If(Summary, "")
                obj("result") = If(Result Is Nothing, JValue.CreateNull(), Result.DeepClone())
                obj("resultKind") = If(ResultKind, "")
                obj("rawLength") = RawLength

                If [Error] IsNot Nothing Then
                    obj("error") = [Error].DeepClone()
                End If

                Return obj
            End Function

            Public Function ToJson() As String
                Return ToJObject().ToString(Formatting.None)
            End Function
        End Class

        Public Shared Function NormalizeFinalOutput(text As String,
                                                    Optional jsonRequired As Boolean = False) As NormalizedEnvelope
            Dim rawText As String = If(text, "")
            Dim rawLength As Integer = rawText.Length
            Dim candidate As String = StripCodeFence(rawText).Trim()

            If String.IsNullOrWhiteSpace(candidate) Then
                Return BuildEmptyResultEnvelope(rawLength)
            End If

            Try

                If Not jsonRequired AndAlso Not LooksLikeJson(candidate) Then
                    Dim summary As String = BuildFallbackSummary(candidate)
                    If String.IsNullOrWhiteSpace(summary) Then
                        Return BuildEmptyResultEnvelope(rawLength)
                    End If

                    Return New NormalizedEnvelope With {
                            .Summary = summary,
                            .Result = New JValue(candidate),
                            .ResultKind = "text",
                            .RawLength = rawLength
                        }
                End If

                Dim tok As JToken = JToken.Parse(candidate)

                If TypeOf tok Is JObject Then
                    Return NormalizeObject(CType(tok, JObject), rawLength)
                End If

                If TypeOf tok Is JArray Then
                    Dim arr = CType(tok, JArray)

                    If arr.Count = 0 Then
                        Return BuildEmptyResultEnvelope(rawLength)
                    End If

                    Return New NormalizedEnvelope With {
                        .Summary = StructuredJsonFallbackSummary,
                        .Result = arr.DeepClone(),
                        .ResultKind = "json_array",
                        .RawLength = rawLength
                    }
                End If

                If jsonRequired Then
                    Return BuildEmptyResultEnvelope(rawLength)
                End If

                If Not IsUsableToken(tok) Then
                    Return BuildEmptyResultEnvelope(rawLength)
                End If

                Return New NormalizedEnvelope With {
                    .Summary = BuildFallbackSummary(candidate),
                    .Result = New JValue(candidate),
                    .ResultKind = "text",
                    .RawLength = rawLength
                }
            Catch
                If jsonRequired Then
                    Return BuildEmptyResultEnvelope(rawLength)
                End If

                Dim summary As String = BuildFallbackSummary(candidate)
                If String.IsNullOrWhiteSpace(summary) Then
                    Return BuildEmptyResultEnvelope(rawLength)
                End If

                Return New NormalizedEnvelope With {
                    .Summary = summary,
                    .Result = New JValue(candidate),
                    .ResultKind = "text",
                    .RawLength = rawLength
                }
            End Try
        End Function

        Public Shared Function BuildEmptyResultEnvelope(Optional rawLength As Integer = 0) As NormalizedEnvelope
            Return New NormalizedEnvelope With {
                .Summary = EmptyResultSummary,
                .Result = Nothing,
                .ResultKind = "error",
                .RawLength = rawLength,
                .Error = New JObject(
                    New JProperty("code", EmptyResultCode),
                    New JProperty("phase", EmptyResultPhase),
                    New JProperty("message", EmptyResultMessage))
            }
        End Function

        Public Shared Function BuildModelEmptyResponsePayload(Optional lastToolName As String = "",
                                                      Optional lastToolResultSummary As String = "",
                                                      Optional compactedToolResponse As Boolean = False,
                                                      Optional retryHint As String = "") As String
            Dim err As New JObject(
        New JProperty("code", ModelEmptyResponseCode),
        New JProperty("phase", ModelEmptyResponsePhase),
        New JProperty("message", ModelEmptyResponseMessage))

            If Not String.IsNullOrWhiteSpace(lastToolName) Then
                err("lastToolName") = lastToolName
            End If

            If Not String.IsNullOrWhiteSpace(lastToolResultSummary) Then
                err("lastToolResultSummary") = lastToolResultSummary
            End If

            If compactedToolResponse Then
                err("compactedToolResponse") = True
            End If

            If Not String.IsNullOrWhiteSpace(retryHint) Then
                err("retryHint") = retryHint
            End If

            Dim obj As New JObject(
        New JProperty("status", ModelEmptyResponseStatus),
        New JProperty("summary", EmptyResultSummary),
        New JProperty("result", JValue.CreateNull()),
        New JProperty("resultKind", "error"),
        New JProperty("error", err))

            Return obj.ToString(Formatting.None)
        End Function


        Public Const ToolScopeEmptyStatus As String = "blocked"
        Public Const ToolScopeEmptyCode As String = "subagent_tool_scope_empty"
        Public Const ToolScopeEmptyPhase As String = "tool_initialization"
        Public Const ToolScopeEmptyMessage As String = "The sub-agent allowed tool scope resolved to no callable tools."

        Public Shared Function BuildToolScopeEmptyPayload(requestedToolNames As IEnumerable(Of String),
                                                          resolvedToolNames As IEnumerable(Of String),
                                                          missingToolNames As IEnumerable(Of String)) As String

            Dim obj As New JObject(
                New JProperty("status", ToolScopeEmptyStatus),
                New JProperty("summary", ToolScopeEmptyMessage),
                New JProperty("result", JValue.CreateNull()),
                New JProperty("resultKind", "error"),
                New JProperty("error", New JObject(
                    New JProperty("code", ToolScopeEmptyCode),
                    New JProperty("phase", ToolScopeEmptyPhase),
                    New JProperty("message", ToolScopeEmptyMessage),
                    New JProperty("requestedTools", New JArray(NormalizeToolNamesForPayload(requestedToolNames).ToArray())),
                    New JProperty("resolvedTools", New JArray(NormalizeToolNamesForPayload(resolvedToolNames).ToArray())),
                    New JProperty("missingTools", New JArray(NormalizeToolNamesForPayload(missingToolNames).ToArray()))
                ))
            )

            Return obj.ToString(Formatting.None)
        End Function

        Private Shared Function NormalizeToolNamesForPayload(names As IEnumerable(Of String)) As List(Of String)
            Dim result As New List(Of String)()
            If names Is Nothing Then Return result

            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            For Each rawName In names
                Dim name As String = If(rawName, "").Trim()
                If name = "" Then Continue For
                If seen.Add(name) Then
                    result.Add(name)
                End If
            Next

            Return result
        End Function

        Public Const RequiredToolMissingSummary As String = "Sub-agent tool environment could not be initialized."
        Public Const RequiredToolMissingCode As String = "subagent_required_tool_missing"
        Public Const RequiredToolMissingPhase As String = "tool_initialization"
        Public Const RequiredToolMissingMessage As String = "One or more required sub-agent tools could not be resolved."
        Public Const ParentRegistryMissingSummary As String = "Sub-agent tool environment could not be initialized."
        Public Const ParentRegistryMissingCode As String = "subagent_parent_registry_missing"
        Public Const ParentRegistryMissingPhase As String = "tool_initialization"
        Public Const ParentRegistryMissingMessage As String = "The parent tooling-run registry snapshot was not passed to the sub-agent."


        Public Shared Function BuildParentRegistryMissingPayload(Optional message As String = Nothing,
                                                         Optional requestedToolNames As IEnumerable(Of String) = Nothing) As String

            Dim finalMessage As String =
        If(String.IsNullOrWhiteSpace(message),
           ParentRegistryMissingMessage,
           message)

            Dim normalizedRequested = NormalizeToolNamesForPayload(requestedToolNames)

            Dim obj As New JObject(
        New JProperty("summary", ParentRegistryMissingSummary),
        New JProperty("result", JValue.CreateNull()),
        New JProperty("resultKind", "error"),
        New JProperty("error", New JObject(
            New JProperty("code", ParentRegistryMissingCode),
            New JProperty("phase", ParentRegistryMissingPhase),
            New JProperty("message", finalMessage)
        ))
    )

            If normalizedRequested.Count > 0 Then
                obj("error")("requestedTools") = New JArray(normalizedRequested.ToArray())
                obj("error")("resolvedTools") = New JArray()
                obj("error")("missingTools") = New JArray(normalizedRequested.ToArray())
            End If

            Return obj.ToString(Formatting.None)
        End Function

        Public Shared Function BuildRequiredToolMissingPayload(missingToolNames As IEnumerable(Of String),
                                                               Optional message As String = Nothing,
                                                               Optional requestedToolNames As IEnumerable(Of String) = Nothing,
                                                               Optional resolvedToolNames As IEnumerable(Of String) = Nothing) As String

            Dim finalMessage As String =
                If(String.IsNullOrWhiteSpace(message),
                   RequiredToolMissingMessage,
                   message)

            Dim obj As New JObject(
                New JProperty("summary", RequiredToolMissingSummary),
                New JProperty("result", JValue.CreateNull()),
                New JProperty("resultKind", "error"),
                New JProperty("error", New JObject(
                    New JProperty("code", RequiredToolMissingCode),
                    New JProperty("phase", RequiredToolMissingPhase),
                    New JProperty("message", finalMessage),
                    New JProperty("missingTools", New JArray(NormalizeToolNamesForPayload(missingToolNames).ToArray()))
                ))
            )

            If requestedToolNames IsNot Nothing Then
                obj("error")("requestedTools") = New JArray(NormalizeToolNamesForPayload(requestedToolNames).ToArray())
            End If

            If resolvedToolNames IsNot Nothing Then
                obj("error")("resolvedTools") = New JArray(NormalizeToolNamesForPayload(resolvedToolNames).ToArray())
            End If

            Return obj.ToString(Formatting.None)
        End Function
        Public Shared Function TryGetEnvelopeErrorMessage(payload As String,
                                                          ByRef errorMessage As String) As Boolean
            errorMessage = ""
            If String.IsNullOrWhiteSpace(payload) Then Return False

            Try
                Dim obj As JObject = JObject.Parse(payload)
                Dim errObj As JObject = TryCast(obj("error"), JObject)
                If errObj Is Nothing Then Return False

                errorMessage = If(errObj.Value(Of String)("message"), "").Trim()
                Return Not String.IsNullOrWhiteSpace(errorMessage)
            Catch ex As System.Exception
                Return False
            End Try
        End Function

        Public Shared Function TryGetEnvelopeRetryable(payload As String,
                                                       ByRef retryable As Boolean) As Boolean
            retryable = False
            If String.IsNullOrWhiteSpace(payload) Then Return False

            Try
                Dim obj As JObject = JObject.Parse(payload)
                Dim errObj As JObject = TryCast(obj("error"), JObject)
                If errObj Is Nothing Then Return False

                Dim retryToken As JToken = errObj("retryable")
                If retryToken Is Nothing OrElse retryToken.Type <> JTokenType.Boolean Then Return False

                retryable = retryToken.Value(Of Boolean)()
                Return True
            Catch ex As System.Exception
                Return False
            End Try
        End Function

        Public Shared Function TryGetEnvelopeErrorInfo(payload As String,
                                                       ByRef errorCode As String,
                                                       ByRef resultKind As String) As Boolean
            errorCode = ""
            resultKind = ""

            If String.IsNullOrWhiteSpace(payload) Then Return False

            Try
                Dim obj As JObject = JObject.Parse(payload)

                resultKind = If(obj.Value(Of String)("resultKind"), "")

                ' Structured error object: {"error":{"code":"..."}}
                Dim errObj As JObject = TryCast(obj("error"), JObject)
                If errObj IsNot Nothing Then
                    errorCode = If(errObj.Value(Of String)("code"), "")
                End If

                ' Flat / string error: {"error":"skill_not_found"} or any non-null "error".
                If String.IsNullOrWhiteSpace(errorCode) Then
                    Dim errToken As JToken = obj("error")
                    If errToken IsNot Nothing AndAlso errToken.Type <> JTokenType.Null Then
                        If errToken.Type = JTokenType.String Then
                            errorCode = errToken.Value(Of String)()
                        ElseIf errToken.Type <> JTokenType.Object Then
                            errorCode = errToken.ToString()
                        End If
                        If String.IsNullOrWhiteSpace(errorCode) Then errorCode = "error"
                    End If
                End If

                ' Explicit business-failure flags: {"success":false} / {"ok":false}.
                If String.IsNullOrWhiteSpace(errorCode) Then
                    If TokenIsExplicitlyFalse(obj("success")) Then errorCode = "success_false"
                    If String.IsNullOrWhiteSpace(errorCode) AndAlso TokenIsExplicitlyFalse(obj("ok")) Then errorCode = "ok_false"
                End If

                ' Explicit status strings: {"status":"error"} / {"status":"failed"}.
                Dim statusValue As String = If(obj.Value(Of String)("status"), "").Trim()
                If String.IsNullOrWhiteSpace(errorCode) AndAlso
                   (String.Equals(statusValue, "error", StringComparison.OrdinalIgnoreCase) OrElse
                    String.Equals(statusValue, "failed", StringComparison.OrdinalIgnoreCase)) Then
                    errorCode = statusValue.ToLowerInvariant()
                End If

                If String.Equals(resultKind, "error", StringComparison.OrdinalIgnoreCase) Then
                    If String.IsNullOrWhiteSpace(resultKind) Then resultKind = "error"
                    Return True
                End If

                If Not String.IsNullOrWhiteSpace(errorCode) Then
                    If String.IsNullOrWhiteSpace(resultKind) Then resultKind = "error"
                    Return True
                End If
            Catch
            End Try

            Return False
        End Function

        Private Shared Function TokenIsExplicitlyFalse(token As JToken) As Boolean
            If token Is Nothing Then Return False

            Select Case token.Type
                Case JTokenType.Boolean
                    Return Not token.Value(Of Boolean)()
                Case JTokenType.String
                    Return String.Equals(token.Value(Of String)(), "false", StringComparison.OrdinalIgnoreCase)
                Case Else
                    Return False
            End Select
        End Function

        Private Shared Function LooksLikeJson(text As String) As Boolean
            If String.IsNullOrWhiteSpace(text) Then Return False

            Dim value As String = text.Trim()
            If value.Length = 0 Then Return False

            Select Case value(0)
                Case "{"c, "["c, """"c, "-"c
                    Return True
                Case "t"c, "f"c, "n"c
                    Return True
                Case Else
                    Return Char.IsDigit(value(0))
            End Select
        End Function

        Private Shared Function NormalizeObject(obj As JObject, rawLength As Integer) As NormalizedEnvelope
            If obj Is Nothing OrElse obj.Count = 0 Then
                Return BuildEmptyResultEnvelope(rawLength)
            End If

            If IsExplicitErrorObject(obj) Then
                Return NormalizeErrorObject(obj, rawLength)
            End If

            Dim hasSummary As Boolean = (obj("summary") IsNot Nothing)
            Dim hasResult As Boolean = (obj("result") IsNot Nothing)

            If hasSummary OrElse hasResult Then
                Dim summary As String = If(obj.Value(Of String)("summary"), "")
                Dim resultToken As JToken = obj("result")

                If Not IsUsableToken(resultToken) Then
                    Return BuildEmptyResultEnvelope(rawLength)
                End If

                If String.IsNullOrWhiteSpace(summary) Then
                    summary = BuildSummaryFromResult(resultToken)
                End If

                Return New NormalizedEnvelope With {
                    .Summary = summary,
                    .Result = resultToken.DeepClone(),
                    .ResultKind = "envelope",
                    .RawLength = rawLength
                }
            End If

            If Not IsUsableToken(obj) Then
                Return BuildEmptyResultEnvelope(rawLength)
            End If

            Return New NormalizedEnvelope With {
                .Summary = StructuredJsonFallbackSummary,
                .Result = obj.DeepClone(),
                .ResultKind = "json_object",
                .RawLength = rawLength
            }
        End Function

        Private Shared Function NormalizeErrorObject(obj As JObject, rawLength As Integer) As NormalizedEnvelope
            Dim errObj As JObject = TryCast(obj("error"), JObject)
            Dim summary As String = If(obj.Value(Of String)("summary"), "")

            If errObj Is Nothing Then
                errObj = New JObject(
                    New JProperty("code", "agent_error"),
                    New JProperty("phase", EmptyResultPhase),
                    New JProperty("message", If(obj.Value(Of String)("message"), EmptyResultMessage)))
            Else
                errObj = CType(errObj.DeepClone(), JObject)
                If String.IsNullOrWhiteSpace(errObj.Value(Of String)("phase")) Then
                    errObj("phase") = EmptyResultPhase
                End If
            End If

            If String.IsNullOrWhiteSpace(summary) Then
                summary = If(errObj.Value(Of String)("message"), "Sub-agent reported an error.")
            End If

            Return New NormalizedEnvelope With {
                .Summary = summary,
                .Result = If(obj("result") Is Nothing, Nothing, obj("result").DeepClone()),
                .ResultKind = "error",
                .RawLength = rawLength,
                .Error = errObj
            }
        End Function

        Private Shared Function IsExplicitErrorObject(obj As JObject) As Boolean
            If obj Is Nothing Then Return False

            If String.Equals(obj.Value(Of String)("resultKind"), "error", StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If

            Dim errObj As JObject = TryCast(obj("error"), JObject)
            If errObj IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(errObj.Value(Of String)("code")) Then
                Return True
            End If

            Return False
        End Function

        Private Shared Function IsUsableToken(token As JToken) As Boolean
            If token Is Nothing Then Return False

            Select Case token.Type
                Case JTokenType.Null, JTokenType.Undefined
                    Return False

                Case JTokenType.String
                    Return Not String.IsNullOrWhiteSpace(token.Value(Of String)())

                Case JTokenType.Object
                    Dim obj = CType(token, JObject)
                    If obj.Count = 0 Then Return False

                    For Each prop In obj.Properties()
                        If IsUsableToken(prop.Value) Then Return True
                    Next

                    Return False

                Case JTokenType.Array
                    Return CType(token, JArray).Count > 0

                Case JTokenType.Boolean,
                     JTokenType.Integer,
                     JTokenType.Float,
                     JTokenType.Date,
                     JTokenType.Bytes,
                     JTokenType.Guid,
                     JTokenType.Uri,
                     JTokenType.TimeSpan
                    Return True

                Case Else
                    Return Not String.IsNullOrWhiteSpace(token.ToString())
            End Select
        End Function

        Private Shared Function BuildSummaryFromResult(resultToken As JToken) As String
            If resultToken Is Nothing Then Return StructuredJsonFallbackSummary

            If resultToken.Type = JTokenType.String Then
                Return BuildFallbackSummary(resultToken.Value(Of String)())
            End If

            Return StructuredJsonFallbackSummary
        End Function

        Private Shared Function StripCodeFence(text As String) As String
            If String.IsNullOrWhiteSpace(text) Then Return If(text, "")

            Dim value As String = text.Trim()

            If value.StartsWith("```", StringComparison.Ordinal) Then
                Dim firstLf As Integer = value.IndexOf(ChrW(10))
                If firstLf >= 0 Then
                    value = value.Substring(firstLf + 1)
                End If

                If value.EndsWith("```", StringComparison.Ordinal) Then
                    value = value.Substring(0, value.Length - 3)
                End If
            End If

            Return value.Trim()
        End Function

        Private Shared Function BuildFallbackSummary(text As String) As String
            Dim line As String = If(text, "").Replace(vbCr, " ").Replace(vbLf, " ").Trim()
            If line.Length <= 160 Then Return line
            Return line.Substring(0, 157) & "..."
        End Function

    End Class

End Namespace
