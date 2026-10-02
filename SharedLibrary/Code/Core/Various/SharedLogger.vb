' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved.
' For license to use see https://redink.ai.
'
' =============================================================================
' File: SharedLogger.vb
' Purpose:
'   Provides lightweight, privacy-preserving logging to per-user/per-host log files,
'   and includes an interactive log analysis utility for aggregating usage statistics.
'
' Notes:
'   - Logging is performed asynchronously via a single consumer queue to avoid file
'     contention and minimize impact on the caller.
'   - The module is designed to fail silently (no exceptions escape to the caller).
'   - User identity is represented by a short SHA-256 derived hash.
'
' Architecture:
'   - `Log`: Enqueues a single append operation.
'   - Background worker: Single thread drains the queue and writes to disk.
'   - `AnalyzeLogs`: Reads log files and aggregates counts and unique-user metrics.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports System.Threading
Imports NAudio
Imports Org.BouncyCastle.Utilities
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedContext
Imports SharedLibrary.SharedLibrary.SharedMethods

Public Module SharedLogger

    ' Single-threaded queue so log writes never contend with each other / file locks.
    Private ReadOnly _logQueue As New System.Collections.Concurrent.BlockingCollection(Of Action)()
    Private _logWorkerStarted As Integer = 0

    ''' <summary>
    ''' Enqueues a single log entry for asynchronous append to a per-user/per-host log file.
    ''' </summary>
    ''' <param name="context">Shared context providing configuration (e.g., log directory path).</param>
    ''' <param name="RDV">
    ''' A descriptive runtime identifier (typically contains host info and optionally a version marker like "(Vx.y)").
    ''' </param>
    ''' <param name="functionName">The name of the function or handler being logged.</param>
    ''' <remarks>
    ''' This routine is intentionally best-effort and must not throw. Log writes are serialized by a dedicated
    ''' background worker thread.
    ''' </remarks>
    Public Sub Log(context As ISharedContext, RDV As String, functionName As String)
        Try
            If context Is Nothing Then Return

            EnsureLogWorker()

            Dim logDir As String = context.INI_LogPath
            If String.IsNullOrWhiteSpace(logDir) Then Return

            Dim appCode As String = GetOfficeHostCode(RDV)
            Dim fileHash As String = GetFileHash(appCode)
            Dim userHash As String = GetUserHash()
            Dim version As String = ""
            If Not String.IsNullOrWhiteSpace(RDV) Then
                Dim start As Integer = RDV.IndexOf("(V", StringComparison.OrdinalIgnoreCase)
                If start >= 0 Then
                    start += 1 ' skip "(" so version starts with "V"
                    Dim endIdx As Integer = RDV.IndexOf(")"c, start)
                    version = If(endIdx > start, RDV.Substring(start, endIdx - start), RDV.Substring(start))
                End If
            End If
            version = version.PadRight(20)

            Dim fileName As String = $"{AN2}-" & fileHash & "-" & appCode & ".log"
            Dim fullPath As String = System.IO.Path.Combine(logDir, fileName)

            Dim line As String =
                "[" & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture) & "] " &
                userHash & " " &
                version & " " &
                functionName

            Debug.WriteLine($"[SharedLogger] Enqueue write -> dir='{logDir}', file='{fileName}', fn='{functionName}'")

            _logQueue.Add(
                Sub()
                    Try
                        ' Ensure directory exists (creates if missing).
                        System.IO.Directory.CreateDirectory(logDir)

                        Using fs As New System.IO.FileStream(
                            fullPath,
                            FileMode.Append,
                            FileAccess.Write,
                            FileShare.Read)

                            Using sw As New System.IO.StreamWriter(fs, Encoding.UTF8)
                                sw.WriteLine(line)
                            End Using
                        End Using

                        Debug.WriteLine($"[SharedLogger] Wrote line -> '{fullPath}'")

                    Catch ex As Exception
                        Debug.WriteLine($"[SharedLogger] Write failed: {ex.GetType().Name}: {ex.Message}")
                        Try
                            ' Optional diagnostic log written alongside the primary logs.
                            Dim dbgPath As String = System.IO.Path.Combine(logDir, $"{AN2}-logger-debug.log")
                            Dim dbgLine As String =
                                    "[" & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture) & "] " &
                                    "ERROR " & ex.GetType().Name & ": " & ex.Message & " | fullPath=" & fullPath
                            System.IO.Directory.CreateDirectory(logDir)
                            System.IO.File.AppendAllText(dbgPath, dbgLine & Environment.NewLine, Encoding.UTF8)
                        Catch
                        End Try

                        ' MUST stay silent
                    End Try
                End Sub)

        Catch
            ' MUST stay silent
        End Try
    End Sub

    ''' <summary>
    ''' Waits until all log writes that were queued before this call have been processed.
    ''' A queue barrier is used instead of polling the file system. Returns False on timeout
    ''' or failure; no exception escapes to the caller.
    ''' </summary>
    Public Function FlushPendingWrites(Optional timeoutMilliseconds As System.Int32 = 5000) As System.Boolean
        Try
            EnsureLogWorker()

            Dim boundedTimeout As System.Int32 = System.Math.Max(1, timeoutMilliseconds)
            Using completed As New System.Threading.ManualResetEventSlim(False)
                _logQueue.Add(Sub() completed.Set())
                Return completed.Wait(boundedTimeout)
            End Using
        Catch ex As System.Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Starts the single background log worker thread if it has not been started yet.
    ''' </summary>
    ''' <remarks>
    ''' Thread creation is guarded by an atomic compare-exchange to ensure only one worker is started.
    ''' </remarks>
    Private Sub EnsureLogWorker()
        If Interlocked.CompareExchange(_logWorkerStarted, 1, 0) <> 0 Then Return

        Dim t As New Thread(
            Sub()
                For Each work In _logQueue.GetConsumingEnumerable()
                    Try
                        work()
                    Catch
                        ' MUST stay silent
                    End Try
                Next
            End Sub)

        t.IsBackground = True
        t.Name = "SharedLoggerWorker"
        t.Start()
    End Sub

    ''' <summary>
    ''' Reads log files from the configured log directory, aggregates usage statistics, and displays the results.
    ''' </summary>
    ''' <param name="context">Shared context providing the log directory path.</param>
    ''' <remarks>
    ''' The analysis produces the following sections (both overall and per add-in):
    ''' 1. Summary: total users, total log entries, date range observed, active days.
    ''' 2. Version adoption: all versions ever seen (with counts and users) plus the "current"
    '''    version per user (only the latest version for each user).
    ''' 3. Feature usage: per feature total count, unique users, per-user average, and per-day average.
    ''' 4. Daily activity: invocations and unique users per calendar day.
    ''' 5. Top users: most active users by invocation count (privacy-preserving hashes).
    ''' </remarks>
    Public Sub AnalyzeLogs(context As ISharedContext)
        Try
            Dim startInput As String = ShowCustomInputBox("Start date (yyyy-MM-dd) or empty", $"{AN} Log Statistics", True)
            Dim endInput As String = ShowCustomInputBox("End date (yyyy-MM-dd) or empty", $"{AN} Log Statistics", True)

            Dim startDate As Nullable(Of DateTime) = ParseDate(startInput)
            Dim endDate As Nullable(Of DateTime) = ParseDate(endInput)

            If String.IsNullOrWhiteSpace(context.INI_LogPath) Then
                ShowCustomMessageBox("Log path is not configured.")
                Exit Sub
            End If

            Dim dir As String = context.INI_LogPath
            If Not System.IO.Directory.Exists(dir) Then
                ShowCustomMessageBox("Log directory does not exist or could not be accessed.")
                Exit Sub
            End If

            Dim files As String() = System.IO.Directory.GetFiles(dir, "redink-*.log")
            If files.Length = 0 Then
                ShowCustomMessageBox("No log files found.")
                Exit Sub
            End If

            ' Initialize progress bar.
            SharedLibrary.ProgressBarModule.CancelOperation = False
            SharedLibrary.ProgressBarModule.GlobalProgressMax = files.Length
            SharedLibrary.ProgressBarModule.GlobalProgressValue = 0
            SharedLibrary.ProgressBarModule.GlobalProgressLabel = "Initializing..."
            SharedLibrary.ProgressBarModule.ShowProgressBarInSeparateThread($"{AN} Log Statistics", "Analyzing log files...")

            ' ----- Data structures: "OVERALL" scope -----
            Dim allStats As New AnalysisScope()

            ' ----- Data structures: per-app scope (keyed by app code: WD/XL/OL/UK) -----
            Dim perApp As New Dictionary(Of String, AnalysisScope)(StringComparer.OrdinalIgnoreCase)

            Dim fileIndex As Integer = 0
            For Each file As String In files
                If SharedLibrary.ProgressBarModule.CancelOperation Then
                    SharedLibrary.ProgressBarModule.CancelOperation = True
                    ShowCustomMessageBox("Analysis cancelled by user.")
                    Exit Sub
                End If

                fileIndex += 1
                SharedLibrary.ProgressBarModule.GlobalProgressValue = fileIndex
                SharedLibrary.ProgressBarModule.GlobalProgressLabel = $"Processing file {fileIndex} of {files.Length}..."

                Dim appCode As String = ExtractAppCode(file)
                If Not perApp.ContainsKey(appCode) Then perApp(appCode) = New AnalysisScope()
                Dim appStats As AnalysisScope = perApp(appCode)

                For Each line As String In System.IO.File.ReadLines(file, Encoding.UTF8)
                    Dim parsed As ParsedLine = ParseLogLine(line)
                    If parsed Is Nothing Then Continue For

                    If startDate.HasValue AndAlso parsed.Time < startDate.Value Then Continue For
                    If endDate.HasValue AndAlso parsed.Time > endDate.Value Then Continue For

                    AccumulateLine(allStats, parsed)
                    AccumulateLine(appStats, parsed)
                Next
            Next

            ' Close progress bar.
            SharedLibrary.ProgressBarModule.CancelOperation = True

            ' ----- Build output -----
            Dim sb As New StringBuilder()

            sb.AppendLine("# RED INK LOG ANALYSIS")
            sb.AppendLine()
            sb.AppendLine("| Setting | Value |")
            sb.AppendLine("| --- | --- |")
            sb.AppendLine($"| Log path | {context.INI_LogPath} |")
            sb.AppendLine($"| Date range filter | {If(startDate.HasValue, startDate.Value.ToString("yyyy-MM-dd"), "(open)")} to {If(endDate.HasValue, endDate.Value.ToString("yyyy-MM-dd"), "(open)")} |")
            sb.AppendLine($"| Files scanned | {files.Length} |")
            sb.AppendLine()

            AppendSection(sb, "OVERALL", allStats)

            For Each appCode In perApp.Keys.OrderBy(Function(x) x)
                AppendSection(sb, FormatAppName(appCode), perApp(appCode))
            Next

            Dim FinalText = ShowCustomWindow(
                $"The following log statistics were compiled based on the content found in {context.INI_LogPath}",
                sb.ToString(), "", $"{AN} Log Statistics")

            If FinalText = "" Then
                PutInClipboard(MarkdownToRtfConverter.Convert(sb.ToString()))
            Else
                FinalText = FinalText.Trim()
                PutInClipboard(FinalText)
            End If

        Catch ex As System.Exception
            SharedLibrary.ProgressBarModule.CancelOperation = True
            ShowCustomMessageBox("Error in AnalyzeLogs: " & ex.Message)
        End Try
    End Sub

    '==============================
    ' ANALYSIS HELPER TYPES
    '==============================

    ''' <summary>
    ''' Holds all aggregated statistics for one scope (overall or a single app code).
    ''' </summary>
    Private Class AnalysisScope
        ' Users.
        Public Users As New HashSet(Of String)()
        Public TotalLines As Integer = 0

        ' Version stats: version -> count; version -> set of users.
        Public VersionUsage As New Dictionary(Of String, Integer)(StringComparer.Ordinal)
        Public VersionUsers As New Dictionary(Of String, HashSet(Of String))(StringComparer.Ordinal)

        ' Latest version per user (user hash -> latest timestamp, version string).
        Public UserLatestVersion As New Dictionary(Of String, KeyValuePair(Of DateTime, String))(StringComparer.Ordinal)

        ' Feature (invoked) stats: feature -> count; feature -> set of users.
        Public FeatureUsage As New Dictionary(Of String, Integer)(StringComparer.Ordinal)
        Public FeatureUsers As New Dictionary(Of String, HashSet(Of String))(StringComparer.Ordinal)

        ' Feature usage per day: feature -> (date -> set of users).
        Public FeatureByDay As New Dictionary(Of String, Dictionary(Of Date, HashSet(Of String)))(StringComparer.Ordinal)

        ' Feature usage per user: feature -> (user -> count).
        Public FeatureByUser As New Dictionary(Of String, Dictionary(Of String, Integer))(StringComparer.Ordinal)

        ' Daily activity (all features combined): date -> (total count, set of users).
        Public DailyCount As New Dictionary(Of Date, Integer)()
        Public DailyUsers As New Dictionary(Of Date, HashSet(Of String))()

        ' Other (non-invoked) lines: key -> count; key -> set of users.
        Public OtherUsage As New Dictionary(Of String, Integer)(StringComparer.Ordinal)
        Public OtherUsers As New Dictionary(Of String, HashSet(Of String))(StringComparer.Ordinal)
    End Class

    ''' <summary>
    ''' Accumulates a single parsed log line into the given analysis scope.
    ''' </summary>
    Private Sub AccumulateLine(scope As AnalysisScope, parsed As ParsedLine)
        scope.TotalLines += 1
        scope.Users.Add(parsed.UserHash)

        ' Version tracking.
        If Not String.IsNullOrWhiteSpace(parsed.Version) Then
            Dim v As String = parsed.Version.Trim()

            If Not scope.VersionUsage.ContainsKey(v) Then scope.VersionUsage(v) = 0
            scope.VersionUsage(v) += 1
            If Not scope.VersionUsers.ContainsKey(v) Then scope.VersionUsers(v) = New HashSet(Of String)()
            scope.VersionUsers(v).Add(parsed.UserHash)

            ' Track latest version per user.
            If scope.UserLatestVersion.ContainsKey(parsed.UserHash) Then
                If parsed.Time > scope.UserLatestVersion(parsed.UserHash).Key Then
                    scope.UserLatestVersion(parsed.UserHash) = New KeyValuePair(Of DateTime, String)(parsed.Time, v)
                End If
            Else
                scope.UserLatestVersion(parsed.UserHash) = New KeyValuePair(Of DateTime, String)(parsed.Time, v)
            End If
        End If

        Dim day As Date = parsed.Time.Date

        If parsed.IsInvoked Then
            ' Feature usage.
            If Not scope.FeatureUsage.ContainsKey(parsed.Key) Then scope.FeatureUsage(parsed.Key) = 0
            scope.FeatureUsage(parsed.Key) += 1
            If Not scope.FeatureUsers.ContainsKey(parsed.Key) Then scope.FeatureUsers(parsed.Key) = New HashSet(Of String)()
            scope.FeatureUsers(parsed.Key).Add(parsed.UserHash)

            ' Feature by day.
            If Not scope.FeatureByDay.ContainsKey(parsed.Key) Then
                scope.FeatureByDay(parsed.Key) = New Dictionary(Of Date, HashSet(Of String))()
            End If
            If Not scope.FeatureByDay(parsed.Key).ContainsKey(day) Then
                scope.FeatureByDay(parsed.Key)(day) = New HashSet(Of String)()
            End If
            scope.FeatureByDay(parsed.Key)(day).Add(parsed.UserHash)

            ' Feature by user.
            If Not scope.FeatureByUser.ContainsKey(parsed.Key) Then
                scope.FeatureByUser(parsed.Key) = New Dictionary(Of String, Integer)(StringComparer.Ordinal)
            End If
            If Not scope.FeatureByUser(parsed.Key).ContainsKey(parsed.UserHash) Then
                scope.FeatureByUser(parsed.Key)(parsed.UserHash) = 0
            End If
            scope.FeatureByUser(parsed.Key)(parsed.UserHash) += 1

            ' Daily totals (all features combined).
            If Not scope.DailyCount.ContainsKey(day) Then scope.DailyCount(day) = 0
            scope.DailyCount(day) += 1
            If Not scope.DailyUsers.ContainsKey(day) Then scope.DailyUsers(day) = New HashSet(Of String)()
            scope.DailyUsers(day).Add(parsed.UserHash)

        Else
            ' Other (non-invoked) lines.
            If Not scope.OtherUsage.ContainsKey(parsed.Key) Then scope.OtherUsage(parsed.Key) = 0
            scope.OtherUsage(parsed.Key) += 1
            If Not scope.OtherUsers.ContainsKey(parsed.Key) Then scope.OtherUsers(parsed.Key) = New HashSet(Of String)()
            scope.OtherUsers(parsed.Key).Add(parsed.UserHash)
        End If
    End Sub

    '==============================
    ' OUTPUT FORMATTING
    '==============================

    ''' <summary>
    ''' Formats an application code into a friendly name.
    ''' </summary>
    Private Function FormatAppName(code As String) As String
        Select Case code.ToUpperInvariant()
            Case "WD" : Return "Word (WD)"
            Case "XL" : Return "Excel (XL)"
            Case "OL" : Return "Outlook (OL)"
            Case "UK" : Return "Unknown (UK)"
            Case Else : Return code
        End Select
    End Function

    ''' <summary>
    ''' Appends a complete formatted statistics section to the output StringBuilder.
    ''' </summary>
    Private Sub AppendSection(sb As StringBuilder, title As String, scope As AnalysisScope)

        sb.AppendLine()
        sb.AppendLine("## " & title)
        sb.AppendLine()

        ' --- 1. Summary ---
        Dim totalInvocations As Integer = 0
        For Each kvp In scope.FeatureUsage
            totalInvocations += kvp.Value
        Next
        Dim activeDays As Integer = scope.DailyCount.Count

        sb.AppendLine()
        sb.AppendLine("### Summary")
        sb.AppendLine()
        sb.AppendLine("| Metric | Value |")
        sb.AppendLine("| --- | --- |")
        sb.AppendLine($"| Total users | {scope.Users.Count} |")
        sb.AppendLine($"| Total log entries | {scope.TotalLines} |")
        sb.AppendLine($"| Total feature invocations | {totalInvocations} |")
        sb.AppendLine($"| Active days | {activeDays} |")
        If activeDays > 0 Then
            sb.AppendLine($"| Avg invocations/day | {(totalInvocations / activeDays):F1} |")
            sb.AppendLine($"| Avg users/day | {(scope.DailyUsers.Values.Sum(Function(s) s.Count) / CDbl(activeDays)):F1} |")
        End If
        If scope.DailyCount.Count > 0 Then
            Dim minDay As Date = scope.DailyCount.Keys.Min()
            Dim maxDay As Date = scope.DailyCount.Keys.Max()
            sb.AppendLine($"| Date range observed | {minDay:yyyy-MM-dd} to {maxDay:yyyy-MM-dd} |")
        End If
        sb.AppendLine()

        ' --- 2. Version adoption ---
        sb.AppendLine()
        sb.AppendLine("### Version History (all versions ever seen)")
        sb.AppendLine()
        If scope.VersionUsage.Count = 0 Then
            sb.AppendLine("(no version data)")
        Else
            sb.AppendLine("| Version | Count | Users |")
            sb.AppendLine("| --- | --- | --- |")
            For Each kvp In scope.VersionUsage.OrderByDescending(Function(x) x.Value)
                Dim u As Integer = If(scope.VersionUsers.ContainsKey(kvp.Key), scope.VersionUsers(kvp.Key).Count, 0)
                sb.AppendLine($"| {kvp.Key} | {kvp.Value} | {u} |")
            Next
        End If
        sb.AppendLine()

        ' Current version adoption (latest version per user).
        sb.AppendLine()
        sb.AppendLine("### Current Versions (latest version per user)")
        sb.AppendLine()
        If scope.UserLatestVersion.Count = 0 Then
            sb.AppendLine("(no version data)")
        Else
            Dim currentVersionCounts As New Dictionary(Of String, Integer)(StringComparer.Ordinal)
            For Each kvp In scope.UserLatestVersion
                Dim v As String = kvp.Value.Value
                If Not currentVersionCounts.ContainsKey(v) Then currentVersionCounts(v) = 0
                currentVersionCounts(v) += 1
            Next
            Dim totalUsersWithVersion As Integer = scope.UserLatestVersion.Count
            sb.AppendLine("| Version | Users | % of total |")
            sb.AppendLine("| --- | --- | --- |")
            For Each kvp In currentVersionCounts.OrderByDescending(Function(x) x.Value)
                Dim pct As Double = If(totalUsersWithVersion > 0, (kvp.Value / CDbl(totalUsersWithVersion)) * 100, 0)
                sb.AppendLine($"| {kvp.Key} | {kvp.Value} | {pct:F1}% |")
            Next
        End If
        sb.AppendLine()

        ' --- 3. Feature usage ---
        sb.AppendLine()
        sb.AppendLine("### Feature Usage")
        sb.AppendLine()
        If scope.FeatureUsage.Count = 0 Then
            sb.AppendLine("(no feature data)")
        Else
            sb.AppendLine("| Feature | Total | Users | Avg/user | Avg/day | Days used |")
            sb.AppendLine("| --- | --- | --- | --- | --- | --- |")
            For Each kvp In scope.FeatureUsage.OrderByDescending(Function(x) x.Value)
                Dim fn As String = kvp.Key
                Dim total As Integer = kvp.Value
                Dim users As Integer = If(scope.FeatureUsers.ContainsKey(fn), scope.FeatureUsers(fn).Count, 0)
                Dim avgPerUser As Double = If(users > 0, total / CDbl(users), 0)
                Dim daysUsed As Integer = If(scope.FeatureByDay.ContainsKey(fn), scope.FeatureByDay(fn).Count, 0)
                Dim avgPerDay As Double = If(daysUsed > 0, total / CDbl(daysUsed), 0)
                sb.AppendLine($"| {fn} | {total} | {users} | {avgPerUser:F1} | {avgPerDay:F1} | {daysUsed} |")
            Next
        End If
        sb.AppendLine()

        ' --- 4. Daily activity ---
        sb.AppendLine()
        sb.AppendLine("### Daily Activity (all features combined)")
        sb.AppendLine()
        If scope.DailyCount.Count = 0 Then
            sb.AppendLine("(no daily data)")
        Else
            sb.AppendLine("| Date | Invocations | Unique users |")
            sb.AppendLine("| --- | --- | --- |")
            For Each logDay In scope.DailyCount.Keys.OrderBy(Function(d) d)
                Dim cnt As Integer = scope.DailyCount(logDay)
                Dim u As Integer = If(scope.DailyUsers.ContainsKey(logDay), scope.DailyUsers(logDay).Count, 0)
                sb.AppendLine($"| {logDay:yyyy-MM-dd} | {cnt} | {u} |")
            Next
        End If
        sb.AppendLine()

        ' --- 5. Top users ---
        sb.AppendLine()
        sb.AppendLine("### Top 20 Users (by total feature invocations)")
        sb.AppendLine()
        Dim userTotals As New Dictionary(Of String, Integer)(StringComparer.Ordinal)
        For Each feat In scope.FeatureByUser
            For Each ukvp In feat.Value
                If Not userTotals.ContainsKey(ukvp.Key) Then userTotals(ukvp.Key) = 0
                userTotals(ukvp.Key) += ukvp.Value
            Next
        Next
        If userTotals.Count = 0 Then
            sb.AppendLine("(no user data)")
        Else
            sb.AppendLine("| User hash | Invocations | Features used |")
            sb.AppendLine("| --- | --- | --- |")
            For Each kvp In userTotals.OrderByDescending(Function(x) x.Value).Take(20)
                Dim featuresUsed As Integer = 0
                For Each feat In scope.FeatureByUser
                    If feat.Value.ContainsKey(kvp.Key) Then featuresUsed += 1
                Next
                sb.AppendLine($"| {kvp.Key} | {kvp.Value} | {featuresUsed} |")
            Next
        End If
        sb.AppendLine()

        ' --- 6. Other log lines ---
        If scope.OtherUsage.Count > 0 Then
            sb.AppendLine()
            sb.AppendLine("### Other Log Entries (non-invoked)")
            sb.AppendLine()
            sb.AppendLine("| Entry | Count | Users |")
            sb.AppendLine("| --- | --- | --- |")
            For Each kvp In scope.OtherUsage.OrderByDescending(Function(x) x.Value)
                Dim u As Integer = If(scope.OtherUsers.ContainsKey(kvp.Key), scope.OtherUsers(kvp.Key).Count, 0)
                sb.AppendLine($"| {kvp.Key} | {kvp.Value} | {u} |")
            Next
            sb.AppendLine()
        End If
    End Sub

    '==============================
    ' LOGGING HELPERS
    '==============================

    ''' <summary>
    ''' Determines a short application code (WD/XL/OL/UK) from a host descriptor string.
    ''' </summary>
    ''' <param name="name">A host descriptor string that may include "word", "excel", or "outlook".</param>
    ''' <returns>A two-letter application code.</returns>
    Private Function GetOfficeHostCode(name As String) As String
        If String.IsNullOrWhiteSpace(name) Then Return "UK"

        Dim n = name.ToLowerInvariant()

        If n.Contains("excel") Then Return "XL"
        If n.Contains("word") Then Return "WD"
        If n.Contains("outlook") Then Return "OL"
        Return "UK"
    End Function

    ''' <summary>
    ''' Resolves all log files in the configured log directory for one Office host.
    ''' The filename filter uses the same host-code contract as <see cref="Log"/> and
    ''' intentionally does not enumerate logs for other Office hosts.
    ''' </summary>
    Public Function TryGetHostLogFilePaths(context As ISharedContext,
                                           hostName As System.String,
                                           ByRef logFiles As System.String()) As System.Boolean
        logFiles = New System.String() {}

        Try
            If context Is Nothing OrElse System.String.IsNullOrWhiteSpace(context.INI_LogPath) Then Return False
            If Not System.IO.Directory.Exists(context.INI_LogPath) Then Return False

            Dim appCode As System.String = GetOfficeHostCode(hostName)
            If System.String.Equals(appCode, "UK", System.StringComparison.OrdinalIgnoreCase) Then Return False

            Dim searchPattern As System.String = $"{AN2}-*-" & appCode & ".log"
            logFiles = System.IO.Directory.GetFiles(context.INI_LogPath, searchPattern, System.IO.SearchOption.TopDirectoryOnly)
            System.Array.Sort(Of System.String)(logFiles, System.StringComparer.OrdinalIgnoreCase)
            Return True
        Catch ex As System.Exception
            logFiles = New System.String() {}
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Produces a stable short hash used to partition log files per user and host application.
    ''' </summary>
    ''' <param name="appCode">The application code (WD/XL/OL/UK).</param>
    ''' <returns>An uppercase hexadecimal hash string.</returns>
    Private Function GetFileHash(appCode As String) As String
        Return ComputeHash(Environment.UserName & "|" &
                           Environment.MachineName & "|" &
                           appCode, 8)
    End Function

    ''' <summary>
    ''' Produces a privacy-preserving short hash for the current user.
    ''' </summary>
    ''' <returns>An uppercase hexadecimal hash string.</returns>
    Private Function GetUserHash() As String
        Return ComputeHash(Environment.UserName, 8)
    End Function

    ''' <summary>
    ''' Computes a short uppercase hexadecimal SHA-256 prefix for the provided input.
    ''' </summary>
    ''' <param name="input">Input value to hash.</param>
    ''' <param name="length">Number of bytes to include from the SHA-256 digest.</param>
    ''' <returns>An uppercase hexadecimal string of <paramref name="length"/> bytes.</returns>
    Private Function ComputeHash(input As String, length As Integer) As String
        Using sha As SHA256 = SHA256.Create()
            Dim bytes As Byte() = sha.ComputeHash(Encoding.UTF8.GetBytes(input))
            Dim sb As New StringBuilder()
            For i As Integer = 0 To length - 1
                sb.Append(bytes(i).ToString("X2"))
            Next
            Return sb.ToString()
        End Using
    End Function

    ''' <summary>
    ''' Extracts the application code from a log file name generated by this module.
    ''' </summary>
    ''' <param name="filePath">Full path to the log file.</param>
    ''' <returns>The application code suffix (e.g., WD, XL, OL, UK).</returns>
    Private Function ExtractAppCode(filePath As String) As String
        Dim name As String = System.IO.Path.GetFileNameWithoutExtension(filePath)
        Dim parts As String() = name.Split("-"c)
        Return parts(parts.Length - 1)
    End Function

    ''' <summary>
    ''' Parses a date in <c>yyyy-MM-dd</c> format.
    ''' </summary>
    ''' <param name="input">User-provided input.</param>
    ''' <returns>A parsed <see cref="DateTime"/> value, or <see langword="Nothing"/> if empty/invalid.</returns>
    Private Function ParseDate(input As String) As Nullable(Of DateTime)
        If String.IsNullOrWhiteSpace(input) Then Return Nothing
        Dim dt As DateTime
        If DateTime.TryParseExact(input, "yyyy-MM-dd", CultureInfo.InvariantCulture,
                                  DateTimeStyles.None, dt) Then
            Return dt
        End If
        Return Nothing
    End Function

    ''' <summary>
    ''' Parsed representation of a single log line.
    ''' </summary>
    Private Class ParsedLine
        ''' <summary>
        ''' Timestamp of the log entry.
        ''' </summary>
        Public Time As DateTime

        ''' <summary>
        ''' Privacy-preserving user identifier.
        ''' </summary>
        Public UserHash As String

        ''' <summary>
        ''' Version field extracted from the log line (trimmed from a fixed-width field).
        ''' </summary>
        Public Version As String

        ''' <summary>
        ''' Grouping key; for "invoked" lines this is the function name, otherwise the first token after the version field.
        ''' </summary>
        Public Key As String

        ''' <summary>
        ''' Indicates whether the log line represents an invoked function event.
        ''' </summary>
        Public IsInvoked As Boolean
    End Class

    ''' <summary>
    ''' Parses a log line written by <see cref="Log"/> into a structured representation.
    ''' </summary>
    ''' <param name="line">A raw log file line.</param>
    ''' <returns>A <see cref="ParsedLine"/> instance, or <see langword="Nothing"/> if parsing fails.</returns>
    Private Function ParseLogLine(line As String) As ParsedLine
        Try
            Dim endBracket As Integer = line.IndexOf("]")
            If endBracket < 0 Then Return Nothing

            Dim timePart As String = line.Substring(1, endBracket - 1)
            Dim rest As String = line.Substring(endBracket + 2)

            Dim firstSpace As Integer = rest.IndexOf(" "c)
            If firstSpace <= 0 Then Return Nothing

            Dim userHash As String = rest.Substring(0, firstSpace).Trim()
            Dim afterUser As String = rest.Substring(firstSpace).TrimStart()
            If String.IsNullOrWhiteSpace(afterUser) Then Return Nothing

            Dim versionField As String
            Dim remainder As String

            If afterUser.Length <= 20 Then
                versionField = afterUser.Trim()
                remainder = ""
            Else
                versionField = afterUser.Substring(0, 20).Trim()
                remainder = afterUser.Substring(20).Trim()
            End If

            If String.IsNullOrWhiteSpace(remainder) Then Return Nothing

            Dim isInvoked As Boolean =
            remainder.EndsWith(" invoked", StringComparison.OrdinalIgnoreCase) OrElse
            remainder.EndsWith(" invoked.", StringComparison.OrdinalIgnoreCase)

            Dim tokens As String() = remainder.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
            If tokens.Length < 1 Then Return Nothing

            Dim key As String
            If isInvoked Then
                key = tokens(0)
            Else
                key = tokens(0)
            End If

            Dim pl As New ParsedLine()
            pl.Time = DateTime.ParseExact(timePart, "yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture)
            pl.UserHash = userHash
            pl.Version = versionField
            pl.Key = key
            pl.IsInvoked = isInvoked
            Return pl
        Catch
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Parses one log line and returns only the timestamp and exact invocation key.
    ''' Raw remainder text and non-invocation records are never returned.
    ''' </summary>
    Friend Function TryParseInvocationKey(line As System.String,
                                          ByRef timestamp As System.DateTime,
                                          ByRef invocationKey As System.String) As System.Boolean
        timestamp = System.DateTime.MinValue
        invocationKey = System.String.Empty

        Dim parsed As ParsedLine = ParseLogLine(line)
        If parsed Is Nothing OrElse Not parsed.IsInvoked Then Return False

        timestamp = parsed.Time
        invocationKey = If(parsed.Key, System.String.Empty)
        Return invocationKey <> System.String.Empty
    End Function

    ''' <summary>
    ''' Logs one agent/tool invocation in a format that AnalyzeLogs can aggregate.
    ''' The full counting key must remain in the first token, followed by "invoked".
    ''' Example:
    '''   AgentToolCall_Outlook_AutoPilot_m365_search invoked
    ''' </summary>
    Public Sub LogAgentToolCall(context As ISharedContext, RDV As String, surface As String, toolName As String)
        Try
            Dim safeSurface As String = NormalizeLogTokenPart(surface)
            Dim safeToolName As String = NormalizeLogTokenPart(toolName)

            If String.IsNullOrWhiteSpace(safeSurface) Then safeSurface = "UnknownSurface"
            If String.IsNullOrWhiteSpace(safeToolName) Then safeToolName = "UnknownTool"

            Log(context, RDV, $"AgentToolCall_{safeSurface}_{safeToolName} invoked")
        Catch
            ' MUST stay silent
        End Try
    End Sub

    ''' <summary>
    ''' Normalizes one token fragment so it is safe for log-key aggregation.
    ''' Keeps letters, digits, and underscores; collapses separators to underscores.
    ''' </summary>
    Friend Function NormalizeLogTokenPart(value As String) As String
        If String.IsNullOrWhiteSpace(value) Then Return ""

        Dim sb As New StringBuilder()
        Dim lastWasUnderscore As Boolean = False

        For Each ch As Char In value.Trim()
            If Char.IsLetterOrDigit(ch) Then
                sb.Append(ch)
                lastWasUnderscore = False

            ElseIf ch = "_"c Then
                If Not lastWasUnderscore AndAlso sb.Length > 0 Then
                    sb.Append("_"c)
                    lastWasUnderscore = True
                End If

            ElseIf Char.IsWhiteSpace(ch) OrElse
                   ch = "-"c OrElse
                   ch = "."c OrElse
                   ch = "/"c OrElse
                   ch = "\"c OrElse
                   ch = ":"c Then

                If Not lastWasUnderscore AndAlso sb.Length > 0 Then
                    sb.Append("_"c)
                    lastWasUnderscore = True
                End If
            End If
        Next

        Return sb.ToString().Trim("_"c)
    End Function

End Module