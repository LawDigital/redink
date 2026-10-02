' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: LogCountTool.vb
' Purpose:
'   Deterministically counts permitted skill invocations in all Red Ink log files
'   for the current host under INI_LogPath. Permission filtering happens before logs
'   are inspected so callers cannot discover unavailable skills through counts.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Text.RegularExpressions
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Agents

    Public NotInheritable Class LogCountTool

        Private Sub New()
        End Sub

        Public Const ToolName As System.String = "log_count"
        Private Const DefaultDays As System.Int32 = 7
        Private Const MaximumDays As System.Int32 = 3660
        Private Const MaximumPatternLength As System.Int32 = 256

        Public NotInheritable Class ExecutionResult
            Public Property Success As System.Boolean
            Public Property Response As System.String
            Public Property ErrorMessage As System.String
        End Class

        Private NotInheritable Class SkillCounter
            Public Property ToolName As System.String
            Public Property LogToken As System.String
            Public Property Total As System.Int32
            Public ReadOnly ByDay As New System.Collections.Generic.Dictionary(Of System.DateTime, System.Int32)()
        End Class

        Public Shared Function IsLogCountTool(name As System.String) As System.Boolean
            Return Not System.String.IsNullOrWhiteSpace(name) AndAlso
                   name.Trim().Equals(ToolName, System.StringComparison.OrdinalIgnoreCase)
        End Function

        Public Shared Function Build(context As SharedLibrary.SharedContext.ISharedContext) As SharedLibrary.ModelConfig
            If context Is Nothing OrElse System.String.IsNullOrWhiteSpace(context.INI_LogPath) Then
                Return Nothing
            End If

            Dim definition As System.String =
                "{""name"":""" & ToolName & """," &
                """description"":""Counts how often skills that are available to the current caller were invoked across the configured Red Ink log files for the current host. Permission filtering is applied before log inspection: unavailable skills are never named, counted, confirmed, or denied. The output is deterministic and pre-aggregated by skill and day. If no explicit period is supplied, the last 7 local calendar days including today are used.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """skill_pattern"":{""type"":""string"",""description"":""Exact skill tool name (with or without the 'skill_' prefix) or wildcard pattern using * and ?. Use a wildcard such as '*bkw*' when the user asks for a family/category of skills. Defaults to '*' for all currently permitted skills.""}," &
                """days"":{""type"":""integer"",""minimum"":1,""maximum"":" & MaximumDays.ToString(System.Globalization.CultureInfo.InvariantCulture) & ",""description"":""Number of local calendar days to include, counting today. Defaults to 7. Ignored when both start_date and end_date are supplied.""}," &
                """start_date"":{""type"":""string"",""description"":""Optional inclusive local start date in yyyy-MM-dd format. If specified, end_date must also be specified.""}," &
                """end_date"":{""type"":""string"",""description"":""Optional inclusive local end date in yyyy-MM-dd format. If specified, start_date must also be specified.""}" &
                "},""additionalProperties"":false}}"

            Return New SharedLibrary.ModelConfig() With {
                .ToolOnly = True,
                .Tool = True,
                .ToolName = ToolName,
                .CapabilityTags = "usage_statistics",
                .ToolPriority = 940,
                .ToolErrorHandling = "skip",
                .ModelDescription = "Permitted skill usage statistics (host logs)",
                .ToolInstructionsPrompt =
                    ToolName & ": Deterministically counts invocations of skills that the current caller is permitted to use, using only log files for the current host under INI_LogPath. " &
                    "Use skill_pattern for an exact skill name or wildcard (* and ?). For a category/fragment request such as 'BKW skills', pass a wildcard such as '*bkw*'. " &
                    "For 'last N days', pass days=N; when no period is specified the tool itself uses the last 7 local calendar days including today. " &
                    "For an explicit date range, pass both start_date and end_date as yyyy-MM-dd; the dates are inclusive. " &
                    "SECURITY BOUNDARY: the tool filters against the caller's permitted skill registry BEFORE reading/counting the log. Never infer, guess, enumerate, or mention skills absent from the returned statistic. " &
                    "If status='not_available', do not provide a numeric estimate and do not say whether the requested skill exists. " &
                    "The returned counts and breakdowns are already deterministic; present them without recalculating or supplementing them.",
                .ToolDefinition = definition
            }
        End Function

        Public Shared Function Execute(arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object),
                                       context As SharedLibrary.SharedContext.ISharedContext,
                                       hostName As System.String,
                                       permittedSkillToolNames As System.Collections.Generic.IEnumerable(Of System.String),
                                       invocationSurfaces As System.Collections.Generic.IEnumerable(Of System.String),
                                       Optional toolInvocationPermitted As System.Boolean = True) As ExecutionResult
            Try
                Dim normalizedHost As System.String = If(hostName, System.String.Empty).Trim()
                If normalizedHost = System.String.Empty Then
                    Return Failure("host_unavailable", "The current host could not be determined.")
                End If

                Dim pattern As System.String = GetString(arguments, "skill_pattern", "*").Trim()
                If pattern = System.String.Empty Then pattern = "*"
                If pattern.Length > MaximumPatternLength Then
                    Return Failure("invalid_skill_pattern", "The skill pattern is too long.")
                End If

                ' Fail closed before inspecting configuration or touching the filesystem when
                ' the caller is not permitted to invoke log_count itself. This prevents a
                ' forged/unexpected tool call from learning whether logging is configured.
                If Not toolInvocationPermitted Then
                    Return SuccessResult(BuildNotAvailableResponse(normalizedHost, pattern))
                End If

                If context Is Nothing OrElse System.String.IsNullOrWhiteSpace(context.INI_LogPath) Then
                    Return Failure("log_statistics_unavailable", "Log statistics are unavailable in the current configuration.")
                End If

                ' Permission filtering MUST happen before any log content is inspected.
                Dim permittedMatches As System.Collections.Generic.List(Of System.String) =
                    ResolvePermittedMatches(permittedSkillToolNames, pattern)

                If permittedMatches.Count = 0 Then
                    Return SuccessResult(BuildNotAvailableResponse(normalizedHost, pattern))
                End If

                Dim startDate As System.DateTime
                Dim endDate As System.DateTime
                Dim periodError As System.String = System.String.Empty
                If Not TryResolvePeriod(arguments, startDate, endDate, periodError) Then
                    Return Failure("invalid_date_range", periodError)
                End If

                If Not Global.SharedLibrary.SharedLogger.FlushPendingWrites() Then
                    Return Failure("log_flush_failed", "The log statistic could not be produced from a complete log snapshot.")
                End If

                Dim logFiles As System.String() = Nothing
                If Not Global.SharedLibrary.SharedLogger.TryGetHostLogFilePaths(context, normalizedHost, logFiles) Then
                    Return Failure("log_statistics_unavailable", "Log statistics are unavailable in the current configuration.")
                End If

                Dim counters As New System.Collections.Generic.List(Of SkillCounter)()
                For Each permittedSkill As System.String In permittedMatches
                    counters.Add(New SkillCounter() With {
                        .ToolName = permittedSkill,
                        .LogToken = Global.SharedLibrary.SharedLogger.NormalizeLogTokenPart(permittedSkill),
                        .Total = 0
                    })
                Next

                If logFiles IsNot Nothing Then
                    For Each logFile As System.String In logFiles
                        CountLogFile(logFile, startDate, endDate, counters, invocationSurfaces)
                    Next
                End If

                Return SuccessResult(BuildStatisticsResponse(normalizedHost, pattern, startDate, endDate, counters))

            Catch ex As System.Exception
                Return Failure("log_count_failed", "The log statistic could not be produced.")
            End Try
        End Function

        Private Shared Sub CountLogFile(logPath As System.String,
                                        startDate As System.DateTime,
                                        endDate As System.DateTime,
                                        counters As System.Collections.Generic.List(Of SkillCounter),
                                        invocationSurfaces As System.Collections.Generic.IEnumerable(Of System.String))
            Dim expectedKeys As New System.Collections.Generic.Dictionary(Of System.String, SkillCounter)(System.StringComparer.OrdinalIgnoreCase)

            If invocationSurfaces IsNot Nothing Then
                For Each rawSurface As System.String In invocationSurfaces
                    Dim safeSurface As System.String = Global.SharedLibrary.SharedLogger.NormalizeLogTokenPart(rawSurface)
                    If safeSurface = System.String.Empty Then Continue For

                    For Each counter As SkillCounter In counters
                        If System.String.IsNullOrWhiteSpace(counter.LogToken) Then Continue For
                        Dim invocationKey As System.String = "AgentToolCall_" & safeSurface & "_" & counter.LogToken
                        If Not expectedKeys.ContainsKey(invocationKey) Then expectedKeys.Add(invocationKey, counter)
                    Next
                Next
            End If

            If expectedKeys.Count = 0 Then Return

            Using stream As New System.IO.FileStream(
                logPath,
                System.IO.FileMode.Open,
                System.IO.FileAccess.Read,
                System.IO.FileShare.ReadWrite Or System.IO.FileShare.Delete)

                Using reader As New System.IO.StreamReader(stream, System.Text.Encoding.UTF8, detectEncodingFromByteOrderMarks:=True)
                    Do While Not reader.EndOfStream
                        Dim rawLine As System.String = reader.ReadLine()
                        If System.String.IsNullOrWhiteSpace(rawLine) Then Continue Do

                        Dim timestamp As System.DateTime
                        Dim invocationKey As System.String = System.String.Empty
                        If Not Global.SharedLibrary.SharedLogger.TryParseInvocationKey(rawLine, timestamp, invocationKey) Then Continue Do

                        Dim logDay As System.DateTime = timestamp.Date
                        If logDay < startDate OrElse logDay > endDate Then Continue Do

                        Dim counter As SkillCounter = Nothing
                        If Not expectedKeys.TryGetValue(invocationKey, counter) OrElse counter Is Nothing Then Continue Do

                        counter.Total += 1
                        If Not counter.ByDay.ContainsKey(logDay) Then counter.ByDay(logDay) = 0
                        counter.ByDay(logDay) += 1
                    Loop
                End Using
            End Using
        End Sub

        Private Shared Function TryResolvePeriod(arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object),
                                                  ByRef startDate As System.DateTime,
                                                  ByRef endDate As System.DateTime,
                                                  ByRef errorMessage As System.String) As System.Boolean
            startDate = System.DateTime.MinValue
            endDate = System.DateTime.MinValue
            errorMessage = System.String.Empty

            Dim startRaw As System.String = GetString(arguments, "start_date", System.String.Empty).Trim()
            Dim endRaw As System.String = GetString(arguments, "end_date", System.String.Empty).Trim()

            Dim hasStart As System.Boolean = startRaw <> System.String.Empty
            Dim hasEnd As System.Boolean = endRaw <> System.String.Empty

            If hasStart Xor hasEnd Then
                errorMessage = "start_date and end_date must either both be supplied or both be omitted."
                Return False
            End If

            If hasStart Then
                If Not System.DateTime.TryParseExact(
                    startRaw,
                    "yyyy-MM-dd",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None,
                    startDate) Then

                    errorMessage = "start_date must use yyyy-MM-dd format."
                    Return False
                End If

                If Not System.DateTime.TryParseExact(
                    endRaw,
                    "yyyy-MM-dd",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None,
                    endDate) Then

                    errorMessage = "end_date must use yyyy-MM-dd format."
                    Return False
                End If

                startDate = startDate.Date
                endDate = endDate.Date

                If startDate > endDate Then
                    errorMessage = "start_date must not be later than end_date."
                    Return False
                End If

                Dim inclusiveDays As System.Double = (endDate - startDate).TotalDays + 1.0R
                If inclusiveDays > MaximumDays Then
                    errorMessage = "The requested date range is too large."
                    Return False
                End If

                Return True
            End If

            Dim days As System.Int32 = DefaultDays
            Dim daysRaw As System.String = GetString(arguments, "days", System.String.Empty).Trim()
            If daysRaw <> System.String.Empty AndAlso
               Not System.Int32.TryParse(
                   daysRaw,
                   System.Globalization.NumberStyles.Integer,
                   System.Globalization.CultureInfo.InvariantCulture,
                   days) Then
                errorMessage = "days must be an integer between 1 and " & MaximumDays.ToString(System.Globalization.CultureInfo.InvariantCulture) & "."
                Return False
            End If

            If days < 1 OrElse days > MaximumDays Then
                errorMessage = "days must be between 1 and " & MaximumDays.ToString(System.Globalization.CultureInfo.InvariantCulture) & "."
                Return False
            End If

            endDate = System.DateTime.Today
            startDate = endDate.AddDays(-(days - 1))
            Return True
        End Function

        Private Shared Function ResolvePermittedMatches(permittedSkillToolNames As System.Collections.Generic.IEnumerable(Of System.String),
                                                        pattern As System.String) As System.Collections.Generic.List(Of System.String)
            Dim result As New System.Collections.Generic.List(Of System.String)()
            Dim seen As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)

            If permittedSkillToolNames Is Nothing Then Return result

            For Each rawName As System.String In permittedSkillToolNames
                Dim toolName As System.String = If(rawName, System.String.Empty).Trim()
                If toolName = System.String.Empty Then Continue For
                If Not toolName.StartsWith("skill_", System.StringComparison.OrdinalIgnoreCase) Then Continue For
                If Not seen.Add(toolName) Then Continue For

                Dim logicalName As System.String = toolName.Substring("skill_".Length)
                If SkillPatternMatches(pattern, toolName) OrElse SkillPatternMatches(pattern, logicalName) Then
                    result.Add(toolName)
                End If
            Next

            result.Sort(System.StringComparer.OrdinalIgnoreCase)
            Return result
        End Function

        Private Shared Function SkillPatternMatches(pattern As System.String,
                                                    candidate As System.String) As System.Boolean
            Dim normalizedPattern As System.String = If(pattern, System.String.Empty).Trim()
            Dim normalizedCandidate As System.String = If(candidate, System.String.Empty).Trim()

            If normalizedPattern = System.String.Empty Then normalizedPattern = "*"
            If normalizedCandidate = System.String.Empty Then Return False

            If normalizedPattern.IndexOf("*"c) < 0 AndAlso normalizedPattern.IndexOf("?"c) < 0 Then
                Return normalizedCandidate.Equals(normalizedPattern, System.StringComparison.OrdinalIgnoreCase)
            End If

            Dim regexPattern As System.String = System.Text.RegularExpressions.Regex.Escape(normalizedPattern)
            regexPattern = regexPattern.Replace("\*", ".*")
            regexPattern = regexPattern.Replace("\?", ".")
            regexPattern = "^" & regexPattern & "$"

            Return System.Text.RegularExpressions.Regex.IsMatch(
                normalizedCandidate,
                regexPattern,
                System.Text.RegularExpressions.RegexOptions.IgnoreCase Or
                System.Text.RegularExpressions.RegexOptions.CultureInvariant)
        End Function

        Private Shared Function BuildNotAvailableResponse(hostName As System.String,
                                                          pattern As System.String) As System.String
            Return New Newtonsoft.Json.Linq.JObject(
                New Newtonsoft.Json.Linq.JProperty("status", "not_available"),
                New Newtonsoft.Json.Linq.JProperty("statistic", "skill_invocation_count"),
                New Newtonsoft.Json.Linq.JProperty("host", hostName),
                New Newtonsoft.Json.Linq.JProperty("skill_pattern", pattern),
                New Newtonsoft.Json.Linq.JProperty("message", "No skill-usage statistics are available for this request under the current permissions.")
            ).ToString(Newtonsoft.Json.Formatting.None)
        End Function

        Private Shared Function BuildStatisticsResponse(hostName As System.String,
                                                        pattern As System.String,
                                                        startDate As System.DateTime,
                                                        endDate As System.DateTime,
                                                        counters As System.Collections.Generic.List(Of SkillCounter)) As System.String
            Dim total As System.Int32 = counters.Sum(Function(item As SkillCounter) item.Total)

            Dim bySkill As New Newtonsoft.Json.Linq.JArray()
            For Each counter As SkillCounter In counters.OrderBy(Function(item As SkillCounter) item.ToolName, System.StringComparer.OrdinalIgnoreCase)
                bySkill.Add(New Newtonsoft.Json.Linq.JObject(
                    New Newtonsoft.Json.Linq.JProperty("skill", counter.ToolName),
                    New Newtonsoft.Json.Linq.JProperty("invocations", counter.Total)
                ))
            Next

            Dim aggregateByDay As New System.Collections.Generic.SortedDictionary(Of System.DateTime, System.Int32)()
            For Each counter As SkillCounter In counters
                For Each entry As System.Collections.Generic.KeyValuePair(Of System.DateTime, System.Int32) In counter.ByDay
                    If Not aggregateByDay.ContainsKey(entry.Key) Then aggregateByDay(entry.Key) = 0
                    aggregateByDay(entry.Key) += entry.Value
                Next
            Next

            Dim byDay As New Newtonsoft.Json.Linq.JArray()
            Dim dayCursor As System.DateTime = startDate.Date
            Do While dayCursor <= endDate.Date
                Dim dayCount As System.Int32 = 0
                aggregateByDay.TryGetValue(dayCursor, dayCount)
                byDay.Add(New Newtonsoft.Json.Linq.JObject(
                    New Newtonsoft.Json.Linq.JProperty("date", dayCursor.ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture)),
                    New Newtonsoft.Json.Linq.JProperty("invocations", dayCount)
                ))
                dayCursor = dayCursor.AddDays(1)
            Loop

            Return New Newtonsoft.Json.Linq.JObject(
                New Newtonsoft.Json.Linq.JProperty("status", "ok"),
                New Newtonsoft.Json.Linq.JProperty("statistic", "skill_invocation_count"),
                New Newtonsoft.Json.Linq.JProperty("host", hostName),
                New Newtonsoft.Json.Linq.JProperty("source_scope", "configured_log_path_current_host_only"),
                New Newtonsoft.Json.Linq.JProperty("period", New Newtonsoft.Json.Linq.JObject(
                    New Newtonsoft.Json.Linq.JProperty("start_date", startDate.ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture)),
                    New Newtonsoft.Json.Linq.JProperty("end_date", endDate.ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture)),
                    New Newtonsoft.Json.Linq.JProperty("semantics", "inclusive_local_calendar_dates")
                )),
                New Newtonsoft.Json.Linq.JProperty("skill_pattern", pattern),
                New Newtonsoft.Json.Linq.JProperty("total_invocations", total),
                New Newtonsoft.Json.Linq.JProperty("by_skill", bySkill),
                New Newtonsoft.Json.Linq.JProperty("by_day", byDay)
            ).ToString(Newtonsoft.Json.Formatting.None)
        End Function

        Private Shared Function SuccessResult(response As System.String) As ExecutionResult
            Return New ExecutionResult() With {
                .Success = True,
                .Response = If(response, System.String.Empty),
                .ErrorMessage = System.String.Empty
            }
        End Function

        Private Shared Function Failure(code As System.String,
                                        message As System.String) As ExecutionResult
            Dim response As System.String = New Newtonsoft.Json.Linq.JObject(
                New Newtonsoft.Json.Linq.JProperty("status", "error"),
                New Newtonsoft.Json.Linq.JProperty("statistic", "skill_invocation_count"),
                New Newtonsoft.Json.Linq.JProperty("error", New Newtonsoft.Json.Linq.JObject(
                    New Newtonsoft.Json.Linq.JProperty("code", If(code, System.String.Empty)),
                    New Newtonsoft.Json.Linq.JProperty("message", If(message, System.String.Empty))
                ))
            ).ToString(Newtonsoft.Json.Formatting.None)

            Return New ExecutionResult() With {
                .Success = False,
                .Response = response,
                .ErrorMessage = If(message, System.String.Empty)
            }
        End Function

        Private Shared Function GetString(arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object),
                                           key As System.String,
                                           fallback As System.String) As System.String
            If arguments Is Nothing OrElse System.String.IsNullOrWhiteSpace(key) Then Return fallback

            Dim value As System.Object = Nothing
            If Not arguments.TryGetValue(key, value) OrElse value Is Nothing Then Return fallback

            If TypeOf value Is Newtonsoft.Json.Linq.JValue Then
                Dim token As Newtonsoft.Json.Linq.JValue = DirectCast(value, Newtonsoft.Json.Linq.JValue)
                If token.Value Is Nothing Then Return fallback
                Return System.Convert.ToString(token.Value, System.Globalization.CultureInfo.InvariantCulture)
            End If

            Return System.Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture)
        End Function



    End Class

End Namespace
