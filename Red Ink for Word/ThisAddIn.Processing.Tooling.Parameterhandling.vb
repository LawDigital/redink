' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.Processing.Tooling.Parameterhandling.vb
' Purpose: Tool argument extraction, parsing, and formatting utilities.
'
' Responsibilities:
'  - Extract string/boolean/list arguments from tool call dictionaries.
'  - Parse parameter types (string, integer, number, boolean, array, object).
'  - Retrieve tool parameter schema and required field lists.
'  - Handle enumerated value parsing.
'  - Format tool values for template substitution (JSON escaping, type coercion).
'  - Resolve default values for unreplaced template placeholders.
'  - Remove optional parameters when no value/default exists.
'  - Build tool argument signatures for duplicate-execution detection.
'  - Normalize link extensions and parameter values.
'  - Generate condensed parameter summaries for log display.
'  - Parse boolean literals with fallback format support (yes/no, true/false, 1/0).
'
' Architecture:
'  - Centralize argument extraction to reduce per-tool parsing.
'  - Support late-bound (dynamic) type conversion via reflection.
'  - Provide JSON escaping for safe template injection.
'
' External Dependencies:
'  - Newtonsoft.Json (JValue, JToken parsing).
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods


Partial Public Class ThisAddIn



    Private Function GetToolArgumentString(arguments As Dictionary(Of String, Object), key As String) As String
        If arguments Is Nothing OrElse Not arguments.ContainsKey(key) OrElse arguments(key) Is Nothing Then
            Return ""
        End If

        Dim value = arguments(key)

        If TypeOf value Is JValue Then
            Return DirectCast(value, JValue).ToString().Trim()
        End If

        Return value.ToString().Trim()
    End Function


    Private Function GetToolArgumentBoolean(arguments As Dictionary(Of String, Object), key As String, Optional defaultValue As Boolean = False) As Boolean
        If arguments Is Nothing OrElse Not arguments.ContainsKey(key) OrElse arguments(key) Is Nothing Then
            Return defaultValue
        End If

        Dim value = arguments(key)

        Try
            If TypeOf value Is Boolean Then
                Return CBool(value)
            End If

            If TypeOf value Is JValue Then
                Dim jv = DirectCast(value, JValue)

                If jv.Type = JTokenType.Boolean Then
                    Return jv.Value(Of Boolean)()
                End If

                value = jv.ToString()
            End If

            Dim text As String = value.ToString().Trim()
            Dim parsed As Boolean

            If Boolean.TryParse(text, parsed) Then
                Return parsed
            End If

            Select Case text.ToLowerInvariant()
                Case "1", "yes", "y", "on"
                    Return True
                Case "0", "no", "n", "off"
                    Return False
            End Select
        Catch
        End Try

        Return defaultValue
    End Function

    Private Function GetToolArgumentStringList(arguments As Dictionary(Of String, Object), key As String) As List(Of String)
        Dim result As New List(Of String)()

        If arguments Is Nothing OrElse Not arguments.ContainsKey(key) OrElse arguments(key) Is Nothing Then
            Return result
        End If

        Dim value = arguments(key)

        Try
            If TypeOf value Is JArray Then
                For Each item In DirectCast(value, JArray)
                    Dim s As String = If(item, "").ToString().Trim()
                    If s <> "" Then result.Add(s)
                Next
                Return result
            End If

            If TypeOf value Is IEnumerable(Of Object) AndAlso Not TypeOf value Is String Then
                For Each item In DirectCast(value, IEnumerable(Of Object))
                    If item Is Nothing Then Continue For
                    Dim s As String = item.ToString().Trim()
                    If s <> "" Then result.Add(s)
                Next
                Return result
            End If

            Dim raw As String = value.ToString().Trim()
            If raw = "" Then Return result

            If raw.StartsWith("[") AndAlso raw.EndsWith("]") Then
                Dim arr As JArray = JArray.Parse(raw)
                For Each item In arr
                    Dim s As String = If(item, "").ToString().Trim()
                    If s <> "" Then result.Add(s)
                Next
                Return result
            End If

            For Each part In raw.Split({","c, ";"c}, StringSplitOptions.RemoveEmptyEntries)
                Dim s As String = part.Trim()
                If s <> "" Then result.Add(s)
            Next
        Catch
        End Try

        Return result
    End Function

    Private Function NormalizeLinkExtensions(values As IEnumerable(Of String)) As List(Of String)
        Dim result As New List(Of String)()

        If values Is Nothing Then
            Return result
        End If

        For Each value In values
            Dim normalized As String = If(value, "").Trim().TrimStart("."c).ToLowerInvariant()
            If normalized = "" Then Continue For

            If Not result.Any(Function(x) x.Equals(normalized, StringComparison.OrdinalIgnoreCase)) Then
                result.Add(normalized)
            End If
        Next

        Return result
    End Function



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

    Private Function GetToolParameterType(schemaToken As JToken) As String
        If schemaToken Is Nothing Then
            Return "string"
        End If

        Dim typeToken As JToken = schemaToken("type")
        Dim typeName As String = If(typeToken, "").ToString().Trim().ToLowerInvariant()

        If typeName <> "" Then
            Return typeName
        End If

        If schemaToken("enum") IsNot Nothing Then
            Return "string"
        End If

        Return "string"
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

        Dim parameterType As String = GetToolParameterType(schemaToken)
        Dim typeMatches As Boolean = True

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

        If Not typeMatches Then
            validationError = $"Parameter '{argumentName}' must be of type '{parameterType}'."
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

    Private Function TryValidateToolCallArguments(toolCall As ToolCall,
                                                  toolConfig As ModelConfig,
                                                  ByRef validationError As String) As Boolean
        validationError = ""

        If toolCall Is Nothing OrElse toolConfig Is Nothing Then
            Return True
        End If

        Dim preparedArguments As System.Collections.Generic.Dictionary(Of System.String, System.Object) = toolCall.Arguments
        Dim normalizationFailure As System.String = System.String.Empty
        If Not SharedLibrary.Agents.ToolCallSequencing.TryPrepareToolCallArgumentsForPreflight(
            toolConfig,
            toolCall.Arguments,
            preparedArguments,
            normalizationFailure) Then

            validationError = If(
                System.String.IsNullOrWhiteSpace(normalizationFailure),
                "Tool arguments could not be normalized for validation.",
                normalizationFailure)
            Return False
        End If
        toolCall.Arguments = preparedArguments

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

        Dim q As String = """"
        Dim propertyNamePattern As String = Regex.Escape(q & placeholderName & q)
        Dim rawPlaceholderPattern As String = Regex.Escape("{" & placeholderName & "}")
        Dim quotedPlaceholderPattern As String = Regex.Escape(q & "{" & placeholderName & "}" & q)

        Dim patterns As String() = {
            ",\s*" & propertyNamePattern & "\s*:\s*" & quotedPlaceholderPattern,
            ",\s*" & propertyNamePattern & "\s*:\s*" & rawPlaceholderPattern,
            propertyNamePattern & "\s*:\s*" & quotedPlaceholderPattern & "\s*,",
            propertyNamePattern & "\s*:\s*" & rawPlaceholderPattern & "\s*,",
            propertyNamePattern & "\s*:\s*" & quotedPlaceholderPattern,
            propertyNamePattern & "\s*:\s*" & rawPlaceholderPattern
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

        If names Is Nothing Then Return result

        For Each name In names
            If String.IsNullOrWhiteSpace(name) Then Continue For
            result.Add(name.Trim())
        Next

        Return result
    End Function

    ''' <summary>
    ''' Builds a condensed parameter summary for display in the log window.
    ''' </summary>
    ''' <param name="arguments">Tool call arguments dictionary.</param>
    ''' <param name="maxLength">Maximum length for each parameter value display.</param>
    ''' <returns>Formatted parameter string like " (query: 'search term', count: 10)".</returns>
    Private Function BuildCondensedParamSummary(arguments As Dictionary(Of String, Object), Optional maxLength As Integer = 50) As String
        If arguments Is Nothing OrElse arguments.Count = 0 Then
            Return ""
        End If

        Dim parts As New List(Of String)()

        For Each kvp In arguments
            Dim valueStr As String = ""
            If kvp.Value IsNot Nothing Then
                If TypeOf kvp.Value Is JArray Then
                    Dim arr = DirectCast(kvp.Value, JArray)
                    valueStr = $"[{arr.Count} items]"
                ElseIf TypeOf kvp.Value Is IEnumerable(Of Object) AndAlso Not TypeOf kvp.Value Is String Then
                    valueStr = $"[{DirectCast(kvp.Value, IEnumerable(Of Object)).Count()} items]"
                Else
                    valueStr = kvp.Value.ToString()
                    If valueStr.Length > maxLength Then
                        valueStr = valueStr.Substring(0, maxLength - 3) & "..."
                    End If
                End If
            End If

            parts.Add($"{kvp.Key}: '{valueStr}'")
        Next

        Return $" ({String.Join(", ", parts)})"
    End Function

End Class
