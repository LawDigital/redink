' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.Tooling.Parameterhandling.vb
' Purpose: Tool parameter schema handling, placeholder resolution, and signature building.
'          Prepares tool call arguments for external API invocation.
'
' Architecture:
'  - Parameter Schema Extraction:
'      - GetToolParameterSchemas(): Parses ToolDefinition JSON for parameter properties.
'      - GetToolRequiredParameters(): Extracts required parameter names from schema.
'      - GetToolParameterType(): Resolves type string (string/integer/number/boolean/array/object).
'      - GetToolParameterEnumValues(): Extracts enum constraint values.
'  - Placeholder Resolution:
'      - FormatToolValueForPlaceholder(): Type-aware formatting (bool/int/number/array/object/string).
'      - ResolveToolDefaultValue(): Applies ToolParameterDefaults when placeholder is unreplaced.
'      - RemoveToolArgumentPlaceholderProperty(): Strips optional unreplaced placeholders from JSON.
'  - Type-Specific Formatting:
'      - Boolean: Parses "true"/"false"/"1"/"0"/"yes"/"no" via TryParseBooleanLiteral().
'      - Integer: InvariantCulture parsing with Long range support.
'      - Number: InvariantCulture parsing with comma-to-dot normalization.
'      - Array: Validates and preserves JSON arrays; defaults to "[]".
'      - Object: Validates and preserves JSON objects; defaults to "{}".
'      - String: Escapes via EscapeJsonString() for JSON string literal injection.
'  - Signature Building:
'      - BuildToolCallSignature(): Creates deterministic signature for duplicate detection.
'      - BuildToolArgumentsSignature(): Normalized JSON representation of arguments dict.
'      - BuildExecutedToolSignature(): Full signature including response/error state.
'      - NormalizeToolArgumentValue(): Recursive normalization for JToken/JValue/collections.
'  - Placeholder Property Removal:
'      - Strips "{param}:"{param}"" and "{param}":{param} patterns via regex.
'      - Cleans up trailing/leading commas and empty structural remnants.
'
' Key Functions:
'  - GetToolParameterSchemas(): Schema extraction entry point.
'  - FormatToolValueForPlaceholder(): Type-aware value formatting.
'  - ResolveToolDefaultValue(): Default value resolution with schema awareness.
'  - BuildToolCallSignature(): Signature for duplicate detection.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.Reflection
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Microsoft.Web.WebView2.Core
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedContext
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn



    Private Function GetToolParameterSchemas(toolConfig As ModelConfig) As Dictionary(Of String, JToken)
        Dim result As New Dictionary(Of String, JToken)(StringComparer.OrdinalIgnoreCase)

        If toolConfig Is Nothing OrElse String.IsNullOrWhiteSpace(toolConfig.ToolDefinition) Then
            Return result
        End If

        Try
            Dim toolDefinition As JObject = JObject.Parse(toolConfig.ToolDefinition)
            Dim propertiesObject As JObject = TryCast(toolDefinition.SelectToken("parameters.properties"), JObject)

            If propertiesObject Is Nothing Then
                Return result
            End If

            For Each prop As JProperty In propertiesObject.Properties()
                result(prop.Name) = prop.Value
            Next
        Catch ex As Exception
            ToolingFileLogger.LogWarn(
                "Failed to parse tool parameter schemas.",
                details:=$"ToolName='{If(toolConfig.ToolName, "")}'",
                ex:=ex)
        End Try

        Return result
    End Function

    Private Function GetToolRequiredParameters(toolConfig As ModelConfig) As HashSet(Of String)
        Dim result As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        If toolConfig Is Nothing OrElse String.IsNullOrWhiteSpace(toolConfig.ToolDefinition) Then
            Return result
        End If

        Try
            Dim toolDefinition As JObject = JObject.Parse(toolConfig.ToolDefinition)
            Dim requiredArray As JArray = TryCast(toolDefinition.SelectToken("parameters.required"), JArray)

            If requiredArray Is Nothing Then
                Return result
            End If

            For Each item As JToken In requiredArray
                Dim name As String = If(item, "").ToString().Trim()
                If name <> "" Then
                    result.Add(name)
                End If
            Next
        Catch ex As Exception
            ToolingFileLogger.LogWarn(
                "Failed to parse required tool parameters.",
                details:=$"ToolName='{If(toolConfig.ToolName, "")}'",
                ex:=ex)
        End Try

        Return result
    End Function

    Private Function GetToolParameterTypes(schemaToken As JToken) As List(Of String)
        Dim result As New List(Of String)()

        If schemaToken Is Nothing Then
            Return result
        End If

        Dim typeToken As JToken = schemaToken("type")

        If TypeOf typeToken Is JArray Then
            For Each item As JToken In DirectCast(typeToken, JArray)
                Dim typeName As String = If(item, "").ToString().Trim().ToLowerInvariant()
                If typeName <> "" Then
                    result.Add(typeName)
                End If
            Next
        Else
            Dim typeName As String = If(typeToken, "").ToString().Trim().ToLowerInvariant()
            If typeName <> "" Then
                result.Add(typeName)
            End If
        End If

        If result.Count = 0 AndAlso schemaToken("enum") IsNot Nothing Then
            result.Add("string")
        End If

        Return result.
            Distinct(StringComparer.OrdinalIgnoreCase).
            ToList()
    End Function

    Private Function GetToolParameterType(schemaToken As JToken) As String
        Dim types As List(Of String) = GetToolParameterTypes(schemaToken)

        If types Is Nothing OrElse types.Count = 0 Then
            Return "string"
        End If

        Return types(0)
    End Function

    Private Function GetToolParameterEnumValues(schemaToken As JToken) As List(Of String)
        Dim values As New List(Of String)()

        If schemaToken Is Nothing Then
            Return values
        End If

        Dim enumArray As JArray = TryCast(schemaToken("enum"), JArray)
        If enumArray Is Nothing Then
            Return values
        End If

        For Each item As JToken In enumArray
            Dim value As String = If(item, "").ToString()
            If value <> "" Then
                values.Add(value)
            End If
        Next

        Return values
    End Function

    Private Function ToolSchemaDisallowsAdditionalProperties(toolConfig As ModelConfig) As Boolean
        If toolConfig Is Nothing OrElse String.IsNullOrWhiteSpace(toolConfig.ToolDefinition) Then
            Return False
        End If

        Try
            Dim toolDefinition As JObject = JObject.Parse(toolConfig.ToolDefinition)
            Dim additionalPropertiesToken As JToken = toolDefinition.SelectToken("parameters.additionalProperties")

            If additionalPropertiesToken Is Nothing Then
                Return False
            End If

            If additionalPropertiesToken.Type = JTokenType.Boolean Then
                Return Not additionalPropertiesToken.Value(Of Boolean)()
            End If

            Dim parsed As Boolean = False
            If Boolean.TryParse(additionalPropertiesToken.ToString(), parsed) Then
                Return Not parsed
            End If
        Catch ex As Exception
            ToolingFileLogger.LogWarn(
                "Failed to parse tool additionalProperties setting.",
                details:=$"ToolName='{If(toolConfig.ToolName, "")}'",
                ex:=ex)
        End Try

        Return False
    End Function

    Private Function TryValidateToolArgumentValueAgainstSchema(argumentName As String,
                                                               value As Object,
                                                               schemaToken As JToken,
                                                               ByRef validationError As String) As Boolean
        validationError = ""

        If schemaToken Is Nothing Then
            Return True
        End If

        Dim token As JToken = Nothing

        Try
            If value Is Nothing Then
                Return True
            End If

            If TypeOf value Is JToken Then
                token = DirectCast(value, JToken)
            Else
                token = JToken.FromObject(value)
            End If
        Catch
            validationError = $"Parameter '{argumentName}' could not be converted for schema validation."
            Return False
        End Try

        If token Is Nothing OrElse
           token.Type = JTokenType.Null OrElse
           token.Type = JTokenType.Undefined Then
            Return True
        End If

        Dim parameterTypes As List(Of String) = GetToolParameterTypes(schemaToken)
        Dim typeMatches As Boolean = (parameterTypes.Count = 0)

        For Each parameterType As String In parameterTypes
            Select Case parameterType
                Case "string"
                    typeMatches = (token.Type = JTokenType.String)

                Case "integer"
                    typeMatches = (token.Type = JTokenType.Integer)

                Case "number"
                    typeMatches = (token.Type = JTokenType.Integer OrElse token.Type = JTokenType.Float)

                Case "boolean"
                    typeMatches = (token.Type = JTokenType.Boolean)

                Case "array"
                    typeMatches = (token.Type = JTokenType.Array)

                Case "object"
                    typeMatches = (token.Type = JTokenType.Object)

                Case Else
                    typeMatches = True
            End Select

            If typeMatches Then Exit For
        Next

        If Not typeMatches Then
            validationError = $"Parameter '{argumentName}' must be of type '{String.Join("' or '", parameterTypes)}'."
            Return False
        End If

        Dim enumValues As List(Of String) = GetToolParameterEnumValues(schemaToken)
        If enumValues.Count > 0 Then
            Dim actualValue As String =
                If(token.Type = JTokenType.String,
                   token.Value(Of String)(),
                   token.ToString(Formatting.None))

            Dim matchesEnum As Boolean =
                enumValues.Any(Function(v) String.Equals(v, actualValue, StringComparison.OrdinalIgnoreCase))

            If Not matchesEnum Then
                validationError = $"Parameter '{argumentName}' must be one of: {String.Join(", ", enumValues)}."
                Return False
            End If
        End If

        If token.Type = JTokenType.Array Then
            Dim itemSchema As JToken = schemaToken("items")

            If itemSchema IsNot Nothing Then
                Dim index As Integer = 0

                For Each item As JToken In DirectCast(token, JArray)
                    Dim itemError As String = ""

                    If Not TryValidateToolArgumentValueAgainstSchema(
                        $"{argumentName}[{index}]",
                        item,
                        itemSchema,
                        itemError) Then

                        validationError = itemError
                        Return False
                    End If

                    index += 1
                Next
            End If
        End If

        Return True
    End Function

    Private Sub NormalizeSchedulerDateArgumentForValidation(arguments As Dictionary(Of String, Object),
                                                            argumentName As String)
        If arguments Is Nothing OrElse String.IsNullOrWhiteSpace(argumentName) Then
            Return
        End If

        If Not arguments.ContainsKey(argumentName) Then
            Return
        End If

        Dim rawValue As Object = arguments(argumentName)
        If rawValue Is Nothing Then
            Return
        End If

        Dim parsedUtc As DateTime
        If TryParseSchedulerUtcArgument(rawValue, parsedUtc) Then
            arguments(argumentName) = parsedUtc.ToString("yyyy-MM-ddTHH:mm:ssZ", Globalization.CultureInfo.InvariantCulture)
        End If
    End Sub

    Private Sub NormalizeToolArgumentsForSchemaValidation(toolCall As ToolCall,
                                                          arguments As Dictionary(Of String, Object))
        If toolCall Is Nothing OrElse arguments Is Nothing Then
            Return
        End If

        If String.Equals(toolCall.ToolName, AP_Tool_ManageScheduledTasks, StringComparison.OrdinalIgnoreCase) Then
            NormalizeSchedulerDateArgumentForValidation(arguments, "next_due_utc")
            NormalizeSchedulerDateArgumentForValidation(arguments, "end_date_utc")
        End If
    End Sub

    Private Function TryValidateToolCallArguments(toolCall As ToolCall,
                                                  toolConfig As ModelConfig,
                                                  ByRef validationError As String) As Boolean
        validationError = ""

        If toolCall Is Nothing OrElse toolConfig Is Nothing Then
            Return True
        End If

        If String.IsNullOrWhiteSpace(toolConfig.ToolDefinition) Then
            Return True
        End If

        Dim parameterSchemas As Dictionary(Of String, JToken) = GetToolParameterSchemas(toolConfig)
        Dim requiredParameters As HashSet(Of String) = GetToolRequiredParameters(toolConfig)

        If parameterSchemas.Count = 0 AndAlso requiredParameters.Count = 0 Then
            Return True
        End If

        Dim arguments As Dictionary(Of String, Object) =
            If(toolCall.Arguments, New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase))

        NormalizeToolArgumentsForSchemaValidation(toolCall, arguments)

        For Each requiredParameter As String In requiredParameters
            If Not arguments.ContainsKey(requiredParameter) OrElse arguments(requiredParameter) Is Nothing Then
                validationError = $"Missing required parameter '{requiredParameter}'."
                Return False
            End If

            If TypeOf arguments(requiredParameter) Is JToken AndAlso
               DirectCast(arguments(requiredParameter), JToken).Type = JTokenType.Null Then

                validationError = $"Required parameter '{requiredParameter}' must not be null."
                Return False
            End If
        Next

        If ToolSchemaDisallowsAdditionalProperties(toolConfig) Then
            For Each argumentName As String In arguments.Keys
                If Not parameterSchemas.ContainsKey(argumentName) Then
                    validationError = $"Unknown parameter '{argumentName}'."
                    Return False
                End If
            Next
        End If

        For Each kvp As KeyValuePair(Of String, Object) In arguments
            If Not parameterSchemas.ContainsKey(kvp.Key) Then Continue For
            If kvp.Value Is Nothing Then Continue For

            If TypeOf kvp.Value Is JToken AndAlso
               DirectCast(kvp.Value, JToken).Type = JTokenType.Null Then
                Continue For
            End If

            Dim argumentError As String = ""

            If Not TryValidateToolArgumentValueAgainstSchema(
                kvp.Key,
                kvp.Value,
                parameterSchemas(kvp.Key),
                argumentError) Then

                validationError = argumentError
                Return False
            End If
        Next

        Return True
    End Function

    Private Function TryParseBooleanLiteral(value As String, ByRef result As Boolean) As Boolean
        Dim normalized As String = If(value, "").Trim()

        If Boolean.TryParse(normalized, result) Then
            Return True
        End If

        Select Case normalized.ToLowerInvariant()
            Case "1", "yes", "y"
                result = True
                Return True
            Case "0", "no", "n"
                result = False
                Return True
            Case Else
                Return False
        End Select
    End Function

    Private Function FormatToolValueForPlaceholder(rawValue As String, schemaToken As JToken) As String
        Dim parameterType As String = GetToolParameterType(schemaToken)
        Dim safeValue As String = If(rawValue, "").Trim()

        Select Case parameterType
            Case "boolean"
                Dim boolValue As Boolean = False
                If TryParseBooleanLiteral(safeValue, boolValue) Then
                    Return If(boolValue, "true", "false")
                End If
                Return "false"

            Case "integer"
                Dim longValue As Long
                If Long.TryParse(safeValue, Globalization.NumberStyles.Integer, Globalization.CultureInfo.InvariantCulture, longValue) OrElse
                   Long.TryParse(safeValue, longValue) Then
                    Return longValue.ToString(Globalization.CultureInfo.InvariantCulture)
                End If
                Return "0"

            Case "number"
                Dim doubleValue As Double
                Dim normalized As String = safeValue.Replace(","c, "."c)

                If Double.TryParse(normalized, Globalization.NumberStyles.Float Or Globalization.NumberStyles.AllowThousands,
                                   Globalization.CultureInfo.InvariantCulture, doubleValue) OrElse
                   Double.TryParse(safeValue, doubleValue) Then
                    Return doubleValue.ToString(Globalization.CultureInfo.InvariantCulture)
                End If

                Return "0"

            Case "array"
                If safeValue <> "" Then
                    Try
                        Dim parsed As JToken = JToken.Parse(safeValue)
                        If parsed.Type = JTokenType.Array Then
                            Return parsed.ToString(Formatting.None)
                        End If
                    Catch
                    End Try
                End If
                Return "[]"

            Case "object"
                If safeValue <> "" Then
                    Try
                        Dim parsed As JToken = JToken.Parse(safeValue)
                        If parsed.Type = JTokenType.Object Then
                            Return parsed.ToString(Formatting.None)
                        End If
                    Catch
                    End Try
                End If
                Return "{}"

            Case Else
                Return EscapeJsonString(safeValue)
        End Select
    End Function

    Private Function ResolveToolDefaultValue(placeholderName As String,
                                             toolDefaults As IDictionary(Of String, String),
                                             schemaToken As JToken,
                                             isRequired As Boolean,
                                             ByRef shouldRemoveProperty As Boolean) As String
        shouldRemoveProperty = False

        Dim rawDefault As String = ""

        If toolDefaults IsNot Nothing AndAlso toolDefaults.ContainsKey(placeholderName) Then
            rawDefault = If(toolDefaults(placeholderName), "")
        End If

        If String.IsNullOrWhiteSpace(rawDefault) Then
            If isRequired Then
                Return "{" & placeholderName & "}"
            End If

            shouldRemoveProperty = True
            Return ""
        End If

        Return FormatToolValueForPlaceholder(rawDefault, schemaToken)
    End Function

    Private Function RemoveToolArgumentPlaceholderProperty(apiCall As String, placeholderName As String) As String
        If String.IsNullOrWhiteSpace(apiCall) OrElse String.IsNullOrWhiteSpace(placeholderName) Then
            Return apiCall
        End If

        Dim quotedPropertyName As String = Regex.Escape("""" & placeholderName & """")
        Dim rawPlaceholder As String = Regex.Escape("{" & placeholderName & "}")
        Dim quotedPlaceholder As String = Regex.Escape("""{" & placeholderName & "}""")

        Dim patterns As String() = {
            ",\s*" & quotedPropertyName & "\s*:\s*" & quotedPlaceholder,
            ",\s*" & quotedPropertyName & "\s*:\s*" & rawPlaceholder,
            quotedPropertyName & "\s*:\s*" & quotedPlaceholder & "\s*,",
            quotedPropertyName & "\s*:\s*" & rawPlaceholder & "\s*,",
            quotedPropertyName & "\s*:\s*" & quotedPlaceholder,
            quotedPropertyName & "\s*:\s*" & rawPlaceholder
        }

        Dim result As String = apiCall

        For Each pattern As String In patterns
            result = Regex.Replace(
                result,
                pattern,
                "",
                RegexOptions.Singleline Or RegexOptions.CultureInvariant)
        Next

        result = Regex.Replace(result, ",\s*(\}|\])", "$1", RegexOptions.Singleline Or RegexOptions.CultureInvariant)
        result = Regex.Replace(result, "(\{|\[)\s*,", "$1", RegexOptions.Singleline Or RegexOptions.CultureInvariant)

        Return result
    End Function


    Private Shared Function BuildToolNameSet(names As IEnumerable(Of String)) As HashSet(Of String)
        Dim result As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        If names Is Nothing Then
            Return result
        End If

        For Each name In names
            If String.IsNullOrWhiteSpace(name) Then Continue For
            result.Add(name.Trim())
        Next

        Return result
    End Function

    Private Function BuildExecutedToolSignature(toolCall As ToolCall, toolResponse As ToolResponse) As String
        Dim toolName As String = If(toolCall?.ToolName, "")
        Dim argsSig As String = BuildToolArgumentsSignature(toolCall?.Arguments)
        Dim successSig As String = If(toolResponse IsNot Nothing AndAlso toolResponse.Success, "ok", "err")
        Dim responseSig As String = If(toolResponse?.Response, "")
        Dim errorSig As String = If(toolResponse?.ErrorMessage, "")

        If responseSig.Length > 500 Then
            responseSig = responseSig.Substring(0, 500)
        End If

        If errorSig.Length > 300 Then
            errorSig = errorSig.Substring(0, 300)
        End If

        Return toolName & "|" & argsSig & "|" & successSig & "|" & responseSig & "|" & errorSig
    End Function


    Private Function BuildToolCallSignature(toolCall As ToolCall) As String
        If toolCall Is Nothing Then Return ""

        Dim parts As New List(Of String)()

        If toolCall.Arguments IsNot Nothing Then
            For Each kvp In toolCall.Arguments.OrderBy(Function(x) x.Key, StringComparer.OrdinalIgnoreCase)
                parts.Add(kvp.Key.ToLowerInvariant() & "=" & NormalizeToolArgumentValue(kvp.Value))
            Next
        End If

        Return toolCall.ToolName & "|" & String.Join(";", parts)
    End Function

    Private Function NormalizeToolArgumentValue(value As Object) As String
        If value Is Nothing Then Return "null"

        If TypeOf value Is JValue Then
            Return DirectCast(value, JValue).ToString(Formatting.None)
        End If

        If TypeOf value Is JToken Then
            Return DirectCast(value, JToken).ToString(Formatting.None)
        End If

        If TypeOf value Is IEnumerable(Of Object) AndAlso Not TypeOf value Is String Then
            Dim items = DirectCast(value, IEnumerable(Of Object)).
                Select(Function(v) NormalizeToolArgumentValue(v))
            Return "[" & String.Join(",", items) & "]"
        End If

        Return System.Convert.ToString(value, Globalization.CultureInfo.InvariantCulture)
    End Function


    Private Function BuildToolArgumentsSignature(arguments As Dictionary(Of String, Object)) As String
        If arguments Is Nothing OrElse arguments.Count = 0 Then
            Return "{}"
        End If

        Try
            Dim normalized As New JObject()

            For Each key In arguments.Keys.OrderBy(Function(k) k, StringComparer.OrdinalIgnoreCase)
                Dim value = arguments(key)

                If TypeOf value Is JToken Then
                    normalized(key) = DirectCast(value, JToken)
                ElseIf value Is Nothing Then
                    normalized(key) = JValue.CreateNull()
                Else
                    normalized(key) = JToken.FromObject(value)
                End If
            Next

            Return normalized.ToString(Formatting.None)
        Catch
            Try
                Return JsonConvert.SerializeObject(arguments)
            Catch
                Return "{}"
            End Try
        End Try
    End Function

    Private Function BuildNormalizedToolCallSignature(toolCall As ToolCall) As String
        If toolCall Is Nothing Then Return ""

        Dim toolName As String = If(toolCall.ToolName, "").Trim()
        If toolName = "" Then Return ""

        ' Append-style operations are legitimately repeatable (each call mutates document
        ' state), so they must never collapse into a duplicate-replay. A unique salt keeps
        ' their signature distinct without affecting idempotent ops like replace/insert_*.
        If IsRepeatableAppendCall(toolCall) Then
            Return toolName.ToLowerInvariant() & "|append|" & Guid.NewGuid().ToString("N")
        End If

        Return toolName.ToLowerInvariant() & "|" & BuildToolArgumentsSignature(toolCall.Arguments)
    End Function

    Private Function IsRepeatableAppendCall(toolCall As ToolCall) As Boolean
        If toolCall Is Nothing OrElse toolCall.Arguments Is Nothing Then Return False

        Dim opValue As Object = Nothing
        If Not toolCall.Arguments.TryGetValue("op", opValue) OrElse opValue Is Nothing Then
            Return False
        End If

        Return System.Convert.ToString(opValue, Globalization.CultureInfo.InvariantCulture).
            Trim().Equals("append", StringComparison.OrdinalIgnoreCase)
    End Function

    Private Function IsDuplicateSuccessGuardedTool(toolName As String) As Boolean
        If String.IsNullOrWhiteSpace(toolName) Then
            Return False
        End If

        Dim normalizedToolName As String = toolName.Trim()

        For Each deliverableToolName As String In SharedLibrary.Agents.HostToolRegistration.GetDeliverableCapableToolNames(
            SharedLibrary.Agents.ToolingHostKind.Outlook)

            If normalizedToolName.Equals(deliverableToolName, StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If
        Next

        Select Case normalizedToolName.ToLowerInvariant()
            Case "complete_word_tables",
                 "process_word_document",
                 "word_write",
                 "word_markup"
                Return True
        End Select

        Return False
    End Function

    Private Function CloneToolResponseForDuplicateReplay(source As ToolResponse,
                                                         toolCall As ToolCall) As ToolResponse
        Return New ToolResponse() With {
            .CallId = toolCall.CallId,
            .ToolName = toolCall.ToolName,
            .Response = source.Response,
            .Success = source.Success,
            .ErrorMessage = source.ErrorMessage,
            .Timestamp = DateTime.Now,
            .OriginalCallJson = toolCall.RawJson,
            .ResultKind = source.ResultKind,
            .ErrorCode = source.ErrorCode,
            .ModelReplayContent = source.ModelReplayContent,
            .ModelReplaySummary = source.ModelReplaySummary,
            .WasCompactedForModelReplay = source.WasCompactedForModelReplay,
            .NormalizedCallSignature = source.NormalizedCallSignature,
            .WasDuplicateReplay = True
        }
    End Function

    Private Function TryBuildDuplicateSuccessfulToolReplay(toolCall As ToolCall,
                                                           normalizedCallSignature As String,
                                                           context As ToolExecutionContext,
                                                           ByRef replayResponse As ToolResponse) As Boolean
        replayResponse = Nothing

        If toolCall Is Nothing OrElse context Is Nothing Then
            Return False
        End If

        If String.IsNullOrWhiteSpace(normalizedCallSignature) Then
            Return False
        End If

        If Not IsDuplicateSuccessGuardedTool(toolCall.ToolName) Then
            Return False
        End If

        If context.AllToolResponses Is Nothing OrElse context.AllToolResponses.Count = 0 Then
            Return False
        End If

        For i As Integer = context.AllToolResponses.Count - 1 To 0 Step -1
            Dim prior As ToolResponse = context.AllToolResponses(i)

            If prior Is Nothing OrElse Not prior.Success Then Continue For
            If prior.WasDuplicateReplay Then Continue For

            If String.Equals(prior.NormalizedCallSignature, normalizedCallSignature, StringComparison.Ordinal) Then
                replayResponse = CloneToolResponseForDuplicateReplay(prior, toolCall)
                Return True
            End If
        Next

        Return False
    End Function

    Private Function BuildDuplicateSuccessfulToolCallGuardPrompt(toolName As String) As String
        Dim normalizedToolName As String = If(toolName, "").Trim()
        Dim toolLabel As String =
            If(normalizedToolName = "",
               "The previous tool call",
               "The tool '" & normalizedToolName & "'")

        Return "HOST DUPLICATE TOOL CALL GUARD: " &
               toolLabel &
               " already completed successfully earlier in this same run with the same arguments. " &
               "Do NOT call it again. Reuse that result. " &
               "If the task is now finished, provide the final answer. " &
               "Otherwise choose the next distinct tool step."
    End Function


End Class
