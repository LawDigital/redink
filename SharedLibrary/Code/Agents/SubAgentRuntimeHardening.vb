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

        Public Const TimeoutSummary As String = "Sub-agent timed out before completing its delegated task."
        Public Const TimeoutCode As String = "subagent_timeout"
        Public Const TimeoutPhase As String = "subagent_execution"
        Public Const TimeoutMessage As String = "The isolated sub-agent attempt exceeded its execution deadline."

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
                Return SubAgentRuntimeHardening.ReadScalarText([Error]("code"))
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

        ''' <summary>
        ''' Builds the canonical retryable timeout envelope for one isolated sub-agent
        ''' attempt. A timeout is not an empty model result and must therefore never be
        ''' routed through the runner's agent_empty_result retry path.
        ''' </summary>
        Public Shared Function BuildTimeoutPayload(agentName As String,
                                                   Optional timeoutSeconds As System.Int32 = 0,
                                                   Optional message As String = Nothing,
                                                   Optional completedResults As Newtonsoft.Json.Linq.JArray = Nothing,
                                                   Optional runState As ToolCallSequencing.ToolingRunState = Nothing) As String
            Dim effectiveMessage As String = If(message, "").Trim()
            If effectiveMessage = "" Then effectiveMessage = TimeoutMessage

            Dim err As New JObject(
                New JProperty("code", TimeoutCode),
                New JProperty("phase", TimeoutPhase),
                New JProperty("message", effectiveMessage),
                New JProperty("retryable", True))

            Dim normalizedAgentName As String = If(agentName, "").Trim()
            If normalizedAgentName <> "" Then err("agent") = normalizedAgentName
            If timeoutSeconds > 0 Then err("timeoutSeconds") = timeoutSeconds

            Dim obj As New JObject(
                New JProperty("summary", TimeoutSummary),
                New JProperty("result", JValue.CreateNull()),
                New JProperty("resultKind", "error"),
                New JProperty("error", err))

            If completedResults IsNot Nothing OrElse runState IsNot Nothing Then
                Dim preserved As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(
                    ToolCallSequencing.BuildBlockedResultPayload(
                        TimeoutCode, TimeoutPhase, effectiveMessage, retryable:=True,
                        completedResults:=completedResults, runState:=runState))
                Dim preservedResult As Newtonsoft.Json.Linq.JToken = preserved("result")
                If preservedResult IsNot Nothing Then obj("result") = preservedResult.DeepClone()
            End If

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
        Public Shared Function TryGetEnvelopeErrorMessage(payload As System.String,
                                                          ByRef errorMessage As System.String) As System.Boolean
            errorMessage = System.String.Empty
            If System.String.IsNullOrWhiteSpace(payload) Then Return False
            Try
                Dim obj As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(payload)
                Dim errorToken As Newtonsoft.Json.Linq.JToken = obj("error")
                If GetErrorTokenCode(errorToken) = System.String.Empty Then Return False
                Dim errObj As Newtonsoft.Json.Linq.JObject = TryCast(errorToken, Newtonsoft.Json.Linq.JObject)
                If errObj IsNot Nothing Then errorMessage = ReadScalarText(errObj("message"))
                If errorMessage = System.String.Empty Then errorMessage = ReadScalarText(obj("message"))
                If errorMessage = System.String.Empty Then
                    errorMessage = If(errorToken.Type = Newtonsoft.Json.Linq.JTokenType.String,
                        errorToken.Value(Of System.String)(), errorToken.ToString(Newtonsoft.Json.Formatting.None))
                End If
                Return Not System.String.IsNullOrWhiteSpace(errorMessage)
            Catch ex As System.Exception
                Return False
            End Try
        End Function

        Public Shared Function TryGetEnvelopeRetryable(payload As System.String,
                                                       ByRef retryable As System.Boolean) As System.Boolean
            retryable = False
            If System.String.IsNullOrWhiteSpace(payload) Then Return False
            Try
                Dim obj As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(payload)
                Dim errObj As Newtonsoft.Json.Linq.JObject = TryCast(obj("error"), Newtonsoft.Json.Linq.JObject)
                Dim retryToken As Newtonsoft.Json.Linq.JToken = If(errObj Is Nothing, Nothing, errObj("retryable"))
                If retryToken Is Nothing Then retryToken = obj("retryable")
                If retryToken Is Nothing Then Return False
                If retryToken.Type = Newtonsoft.Json.Linq.JTokenType.Boolean Then
                    retryable = retryToken.Value(Of System.Boolean)()
                    Return True
                End If
                If retryToken.Type = Newtonsoft.Json.Linq.JTokenType.String Then
                    Return System.Boolean.TryParse(retryToken.Value(Of System.String)().Trim(), retryable)
                End If
                Return False
            Catch ex As System.Exception
                Return False
            End Try
        End Function

        Public Shared Function IsEmptyResultRetryAllowed(payload As System.String,
                                                         errorCode As System.String) As System.Boolean
            If Not System.String.Equals(errorCode, EmptyResultCode, System.StringComparison.OrdinalIgnoreCase) AndAlso
               Not System.String.Equals(errorCode, ModelEmptyResponseCode, System.StringComparison.OrdinalIgnoreCase) Then Return False
            Dim declaredRetryable As System.Boolean = True
            Return Not TryGetEnvelopeRetryable(payload, declaredRetryable) OrElse declaredRetryable
        End Function

        Public Shared Function TryGetEnvelopeErrorInfo(payload As System.String,
                                                       ByRef errorCode As System.String,
                                                       ByRef resultKind As System.String) As System.Boolean
            errorCode = System.String.Empty
            resultKind = System.String.Empty
            If System.String.IsNullOrWhiteSpace(payload) Then Return False
            Try
                Dim obj As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(payload)
                If Not IsProtocolEnvelope(obj) AndAlso Not IsLegacyOutcomeObject(obj) Then Return False
                resultKind = ReadScalarText(obj("resultKind"))
                errorCode = GetErrorTokenCode(obj("error"))
                If errorCode = System.String.Empty AndAlso TokenIsExplicitlyFalse(obj("success")) Then errorCode = "success_false"
                If errorCode = System.String.Empty AndAlso TokenIsExplicitlyFalse(obj("ok")) Then errorCode = "ok_false"
                Dim statusValue As System.String = ReadScalarText(obj("status"))
                If errorCode = System.String.Empty AndAlso
                   (System.String.Equals(statusValue, "error", System.StringComparison.OrdinalIgnoreCase) OrElse
                    System.String.Equals(statusValue, "failed", System.StringComparison.OrdinalIgnoreCase)) Then errorCode = statusValue.ToLowerInvariant()
                If System.String.Equals(resultKind, "error", System.StringComparison.OrdinalIgnoreCase) AndAlso errorCode = System.String.Empty Then errorCode = "agent_error"
                If errorCode = System.String.Empty Then Return False
                If resultKind = System.String.Empty Then resultKind = "error"
                Return True
            Catch ex As System.Exception
                Return False
            End Try
        End Function

        ' Only reserved protocol fields are interpreted. Arbitrary nested domain data is not scanned.
        Private Shared Function ReadScalarText(token As Newtonsoft.Json.Linq.JToken) As System.String
            If token Is Nothing OrElse token.Type = Newtonsoft.Json.Linq.JTokenType.Null OrElse
               token.Type = Newtonsoft.Json.Linq.JTokenType.Object OrElse token.Type = Newtonsoft.Json.Linq.JTokenType.Array Then Return System.String.Empty
            Return token.ToString().Trim()
        End Function

        Private Shared Function GetErrorTokenCode(token As Newtonsoft.Json.Linq.JToken) As System.String
            If token Is Nothing OrElse token.Type = Newtonsoft.Json.Linq.JTokenType.Null Then Return System.String.Empty
            If token.Type = Newtonsoft.Json.Linq.JTokenType.Boolean Then
                Return If(token.Value(Of System.Boolean)(), "error", System.String.Empty)
            End If
            If token.Type = Newtonsoft.Json.Linq.JTokenType.String Then Return token.Value(Of System.String)().Trim()
            Dim obj As Newtonsoft.Json.Linq.JObject = TryCast(token, Newtonsoft.Json.Linq.JObject)
            If obj IsNot Nothing Then
                If obj.Count = 0 Then Return System.String.Empty
                Dim code As System.String = ReadScalarText(obj("code"))
                Return If(code = System.String.Empty, "error", code)
            End If
            If token.Type = Newtonsoft.Json.Linq.JTokenType.Array AndAlso Not token.HasValues Then Return System.String.Empty
            Return "error"
        End Function

        Private Shared Function TokenIsExplicitlyFalse(token As JToken) As Boolean
            If token Is Nothing Then Return False

            Select Case token.Type
                Case JTokenType.Boolean
                    Return Not token.Value(Of Boolean)()
                Case JTokenType.String
                    Return System.String.Equals(token.Value(Of System.String)().Trim(), "false", System.StringComparison.OrdinalIgnoreCase)
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

            If IsProtocolEnvelope(obj) Then
                Dim summary As System.String = ReadScalarText(obj("summary"))
                Dim resultToken As JToken = obj("result")

                ' A sub-agent model call can complete successfully at the transport layer while
                ' its declared task outcome is a failure. Preserve that distinction here, at the
                ' normalization boundary, so every host sees the same contract semantics. This
                ' interpretation is deliberately restricted to the documented {summary,result}
                ' agent envelope; a direct JSON object remains arbitrary domain data.
                Dim envelopeFailureCode As System.String = GetDeclaredOutcomeFailureCode(obj)
                If Not System.String.IsNullOrWhiteSpace(envelopeFailureCode) Then
                    Return NormalizeDeclaredFailureObject(obj, resultToken, envelopeFailureCode, rawLength)
                End If

                If Not IsUsableToken(resultToken) Then
                    Return BuildEmptyResultEnvelope(rawLength)
                End If

                Dim resultObject As JObject = TryCast(resultToken, JObject)
                Dim nestedFailureCode As System.String = System.String.Empty
                Dim declaredKind As System.String = ReadScalarText(obj("resultKind")).ToLowerInvariant()
                If (declaredKind = System.String.Empty OrElse declaredKind = "envelope") AndAlso
                   (IsProtocolEnvelope(resultObject) OrElse IsLegacyOutcomeObject(resultObject)) Then
                    nestedFailureCode = GetDeclaredOutcomeFailureCode(resultObject)
                End If
                If Not System.String.IsNullOrWhiteSpace(nestedFailureCode) Then
                    Return NormalizeDeclaredFailureObject(obj, resultToken, nestedFailureCode, rawLength)
                End If

                If String.IsNullOrWhiteSpace(summary) Then
                    summary = BuildSummaryFromResult(resultToken)
                End If

                Return New NormalizedEnvelope With {
                    .Summary = summary,
                    .Result = resultToken.DeepClone(),
                    .ResultKind = If(declaredKind = "json_object" OrElse declaredKind = "json_array" OrElse declaredKind = "text", declaredKind, "envelope"),
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

        Private Shared Function NormalizeErrorObject(obj As Newtonsoft.Json.Linq.JObject, rawLength As System.Int32) As NormalizedEnvelope
            Dim errObj As Newtonsoft.Json.Linq.JObject = TryCast(obj("error"), Newtonsoft.Json.Linq.JObject)
            Dim summary As System.String = ReadScalarText(obj("summary"))
            Dim code As System.String = GetErrorTokenCode(obj("error"))
            If code = System.String.Empty Then code = "agent_error"
            If errObj Is Nothing Then
                errObj = New Newtonsoft.Json.Linq.JObject()
            Else
                errObj = DirectCast(errObj.DeepClone(), Newtonsoft.Json.Linq.JObject)
            End If
            errObj("code") = code
            If ReadScalarText(errObj("phase")) = System.String.Empty Then errObj("phase") = EmptyResultPhase
            If ReadScalarText(errObj("message")) = System.String.Empty Then
                Dim message As System.String = System.String.Empty
                If Not TryGetEnvelopeErrorMessage(obj.ToString(Newtonsoft.Json.Formatting.None), message) Then message = "Sub-agent reported an error."
                errObj("message") = message
            End If
            If errObj("retryable") Is Nothing AndAlso obj("retryable") IsNot Nothing Then errObj("retryable") = obj("retryable").DeepClone()
            If summary = System.String.Empty Then summary = ReadScalarText(errObj("message"))
            Return New NormalizedEnvelope With {
                .Summary = summary,
                .Result = If(obj("result") Is Nothing, Nothing, obj("result").DeepClone()),
                .ResultKind = "error",
                .RawLength = rawLength,
                .Error = errObj
            }
        End Function

        Private Shared Function NormalizeDeclaredFailureObject(container As JObject,
                                                               resultToken As JToken,
                                                               errorCode As System.String,
                                                               rawLength As System.Int32) As NormalizedEnvelope
            Dim normalizedCode As System.String = If(errorCode, System.String.Empty).Trim()
            If System.String.IsNullOrWhiteSpace(normalizedCode) Then normalizedCode = "agent_declared_failure"

            Dim resultObject As JObject = TryCast(resultToken, JObject)
            Dim errObj As JObject = Nothing

            If resultObject IsNot Nothing Then
                errObj = TryCast(resultObject("error"), JObject)
            End If
            If errObj Is Nothing AndAlso container IsNot Nothing Then
                errObj = TryCast(container("error"), JObject)
            End If

            If errObj Is Nothing Then
                errObj = New JObject()
            Else
                errObj = CType(errObj.DeepClone(), JObject)
            End If

            If System.String.IsNullOrWhiteSpace(ReadScalarText(errObj("code"))) Then
                errObj("code") = normalizedCode
            End If
            If System.String.IsNullOrWhiteSpace(ReadScalarText(errObj("phase"))) Then
                errObj("phase") = "declared_result_status"
            End If

            Dim summary As System.String = System.String.Empty
            If container IsNot Nothing Then
                summary = ReadScalarText(container("summary"))
            End If
            If System.String.IsNullOrWhiteSpace(summary) AndAlso resultObject IsNot Nothing Then
                summary = ReadScalarText(resultObject("message"))
                If System.String.IsNullOrWhiteSpace(summary) Then
                    summary = ReadScalarText(resultObject("reason"))
                End If
            End If
            If System.String.IsNullOrWhiteSpace(summary) Then
                summary = "Sub-agent returned declared failure outcome '" & normalizedCode & "'."
            End If

            If System.String.IsNullOrWhiteSpace(ReadScalarText(errObj("message"))) Then
                errObj("message") = summary
            End If
            If errObj("retryable") Is Nothing Then
                Dim retryToken As Newtonsoft.Json.Linq.JToken = If(resultObject Is Nothing, Nothing, resultObject("retryable"))
                If retryToken Is Nothing AndAlso container IsNot Nothing Then retryToken = container("retryable")
                If retryToken IsNot Nothing Then errObj("retryable") = retryToken.DeepClone()
            End If

            Return New NormalizedEnvelope With {
                .Summary = summary,
                .Result = If(resultToken Is Nothing, Nothing, resultToken.DeepClone()),
                .ResultKind = "error",
                .RawLength = rawLength,
                .Error = errObj
            }
        End Function

        Private Shared Function GetDeclaredOutcomeFailureCode(obj As JObject) As System.String
            If obj Is Nothing Then Return System.String.Empty

            ' Keep this intentionally narrow. The object itself is an agent outcome object,
            ' but its arbitrary domain fields are still data. Only explicit generic outcome
            ' flags and the reserved failure-status vocabulary are promoted to orchestration
            ' failure. Existing top-level structured error-envelope handling remains unchanged.
            Dim explicitError As System.String = GetErrorTokenCode(obj("error"))
            If explicitError <> System.String.Empty Then Return explicitError
            If TokenIsExplicitlyFalse(obj("success")) Then Return "success_false"
            If TokenIsExplicitlyFalse(obj("ok")) Then Return "ok_false"

            Dim statusValue As System.String = ReadScalarText(obj("status")).ToLowerInvariant()
            If statusValue = "error" OrElse
               statusValue = "failed" OrElse
               statusValue = "failure" OrElse
               statusValue = "blocked" OrElse
               statusValue.StartsWith("blocked_", System.StringComparison.Ordinal) Then

                Return statusValue
            End If

            Return System.String.Empty
        End Function

        ''' <summary>
        ''' Returns True for declared outcome codes whose semantics explicitly mean that the
        ''' delegated task cannot continue without an external prerequisite/change. Ordinary
        ''' error/failed outcomes remain subject to the configured generic recovery policy.
        ''' </summary>
        Public Shared Function IsDeclaredTerminalOutcomeErrorCode(errorCode As System.String) As System.Boolean
            Dim normalized As System.String = If(errorCode, System.String.Empty).Trim().ToLowerInvariant()
            If normalized = "blocked" Then Return True
            Return normalized.StartsWith("blocked_", System.StringComparison.Ordinal)
        End Function

        ' A reserved envelope requires an explicit discriminator or the complete legacy shape.
        ' A single domain field named summary/result/error is not envelope metadata.
        Private Shared Function IsProtocolEnvelope(obj As Newtonsoft.Json.Linq.JObject) As System.Boolean
            If obj Is Nothing Then Return False
            Dim kind As System.String = ReadScalarText(obj("resultKind")).ToLowerInvariant()
            If kind = "error" Then Return True
            ' A discriminator alone can be a domain field. Require the payload member
            ' before interpreting a non-error object as an envelope (JSON null still counts).
            If obj("result") IsNot Nothing AndAlso
               (kind = "envelope" OrElse kind = "json_object" OrElse kind = "json_array" OrElse kind = "text") Then Return True
            If obj("summary") Is Nothing OrElse obj("result") Is Nothing Then Return False
            Return HasOnlyOutcomeProperties(obj)
        End Function

        Private Shared Function HasOnlyOutcomeProperties(obj As Newtonsoft.Json.Linq.JObject) As System.Boolean
            If obj Is Nothing OrElse obj.Count = 0 Then Return False
            Dim reserved As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal) From {
                "summary", "result", "resultKind", "rawLength", "error", "status", "success", "ok", "message", "reason", "retryable", "code", "phase"
            }
            For Each prop As Newtonsoft.Json.Linq.JProperty In obj.Properties()
                If Not reserved.Contains(prop.Name) Then Return False
            Next
            Return True
        End Function

        Private Shared Function IsLegacyOutcomeObject(obj As Newtonsoft.Json.Linq.JObject) As System.Boolean
            If Not HasOnlyOutcomeProperties(obj) Then Return False
            ' Retain explicit failure reports, including structured {error:{code,...}}.
            Dim structuredError As Newtonsoft.Json.Linq.JObject = TryCast(obj("error"), Newtonsoft.Json.Linq.JObject)
            If structuredError IsNot Nothing AndAlso structuredError("code") IsNot Nothing Then Return True
            Return obj("status") IsNot Nothing OrElse obj("success") IsNot Nothing OrElse obj("ok") IsNot Nothing
        End Function

        Private Shared Function IsExplicitErrorObject(obj As Newtonsoft.Json.Linq.JObject) As System.Boolean
            If obj Is Nothing Then Return False
            Return System.String.Equals(ReadScalarText(obj("resultKind")), "error", System.StringComparison.OrdinalIgnoreCase) OrElse
                ((IsProtocolEnvelope(obj) OrElse IsLegacyOutcomeObject(obj)) AndAlso
                 GetErrorTokenCode(obj("error")) <> System.String.Empty)
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
