' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: SharedMethods.FileImporter.vb
' Purpose: 
'
' Architecture:
'
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports System.Threading
Imports SharedLibrary.SharedLibrary.SharedContext
Imports SharedLibrary.SharedLibrary.SharedMethods

Public Module SharedLogger

    ' Single-threaded queue so log writes never contend with each other / file locks.
    Private ReadOnly _logQueue As New System.Collections.Concurrent.BlockingCollection(Of Action)()
    Private _logWorkerStarted As Integer = 0

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
                Version & " " &
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
                            ' Optional: a separate debug log, only when INI_APIDebug is enabled.
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

            ' --- OVERALL ---
            Dim allUsers As New HashSet(Of String)()

            Dim allInvokedUsage As New Dictionary(Of String, Integer)(StringComparer.Ordinal)
            Dim allInvokedUsers As New Dictionary(Of String, HashSet(Of String))(StringComparer.Ordinal)

            Dim allOtherUsage As New Dictionary(Of String, Integer)(StringComparer.Ordinal)
            Dim allOtherUsers As New Dictionary(Of String, HashSet(Of String))(StringComparer.Ordinal)

            Dim allInvokedUsageByDay As New Dictionary(Of String, Dictionary(Of Date, HashSet(Of String)))(StringComparer.Ordinal)

            Dim allVersionUsage As New Dictionary(Of String, Integer)(StringComparer.Ordinal)
            Dim allVersionUsers As New Dictionary(Of String, HashSet(Of String))(StringComparer.Ordinal)

            ' --- PER APP (by file name suffix: WD / XL / OL / UK) ---
            Dim appUsers As New Dictionary(Of String, HashSet(Of String))(StringComparer.OrdinalIgnoreCase)

            Dim appInvokedUsage As New Dictionary(Of String, Dictionary(Of String, Integer))(StringComparer.OrdinalIgnoreCase)
            Dim appInvokedUsers As New Dictionary(Of String, Dictionary(Of String, HashSet(Of String)))(StringComparer.OrdinalIgnoreCase)

            Dim appOtherUsage As New Dictionary(Of String, Dictionary(Of String, Integer))(StringComparer.OrdinalIgnoreCase)
            Dim appOtherUsers As New Dictionary(Of String, Dictionary(Of String, HashSet(Of String)))(StringComparer.OrdinalIgnoreCase)

            Dim appVersionUsage As New Dictionary(Of String, Dictionary(Of String, Integer))(StringComparer.OrdinalIgnoreCase)
            Dim appVersionUsers As New Dictionary(Of String, Dictionary(Of String, HashSet(Of String)))(StringComparer.OrdinalIgnoreCase)

            Dim appInvokedUsageByDay As New Dictionary(Of String, Dictionary(Of String, Dictionary(Of Date, HashSet(Of String))))(StringComparer.OrdinalIgnoreCase)

            For Each file As String In System.IO.Directory.GetFiles(dir, "redink-*.log")
                Dim appCode As String = ExtractAppCode(file)

                If Not appUsers.ContainsKey(appCode) Then appUsers(appCode) = New HashSet(Of String)()

                If Not appInvokedUsage.ContainsKey(appCode) Then appInvokedUsage(appCode) = New Dictionary(Of String, Integer)(StringComparer.Ordinal)
                If Not appInvokedUsers.ContainsKey(appCode) Then appInvokedUsers(appCode) = New Dictionary(Of String, HashSet(Of String))(StringComparer.Ordinal)

                If Not appOtherUsage.ContainsKey(appCode) Then appOtherUsage(appCode) = New Dictionary(Of String, Integer)(StringComparer.Ordinal)
                If Not appOtherUsers.ContainsKey(appCode) Then appOtherUsers(appCode) = New Dictionary(Of String, HashSet(Of String))(StringComparer.Ordinal)

                If Not appVersionUsage.ContainsKey(appCode) Then appVersionUsage(appCode) = New Dictionary(Of String, Integer)(StringComparer.Ordinal)
                If Not appVersionUsers.ContainsKey(appCode) Then appVersionUsers(appCode) = New Dictionary(Of String, HashSet(Of String))(StringComparer.Ordinal)

                If Not appInvokedUsageByDay.ContainsKey(appCode) Then
                    appInvokedUsageByDay(appCode) = New Dictionary(Of String, Dictionary(Of Date, HashSet(Of String)))(StringComparer.Ordinal)
                End If

                For Each line As String In System.IO.File.ReadAllLines(file, Encoding.UTF8)
                    Dim parsed As ParsedLine = ParseLogLine(line)
                    If parsed Is Nothing Then Continue For

                    If startDate.HasValue AndAlso parsed.Time < startDate.Value Then Continue For
                    If endDate.HasValue AndAlso parsed.Time > endDate.Value Then Continue For

                    allUsers.Add(parsed.UserHash)
                    appUsers(appCode).Add(parsed.UserHash)

                    ' --- Version stats (overall + per-app) ---
                    If Not String.IsNullOrWhiteSpace(parsed.Version) Then
                        Dim v As String = parsed.Version.Trim()

                        If Not allVersionUsage.ContainsKey(v) Then allVersionUsage(v) = 0
                        allVersionUsage(v) += 1
                        If Not allVersionUsers.ContainsKey(v) Then allVersionUsers(v) = New HashSet(Of String)()
                        allVersionUsers(v).Add(parsed.UserHash)

                        If Not appVersionUsage(appCode).ContainsKey(v) Then appVersionUsage(appCode)(v) = 0
                        appVersionUsage(appCode)(v) += 1
                        If Not appVersionUsers(appCode).ContainsKey(v) Then appVersionUsers(appCode)(v) = New HashSet(Of String)()
                        appVersionUsers(appCode)(v).Add(parsed.UserHash)
                    End If

                    If parsed.IsInvoked Then
                        ' --- Invoked stats (overall + per-app) ---
                        If Not allInvokedUsage.ContainsKey(parsed.Key) Then allInvokedUsage(parsed.Key) = 0
                        allInvokedUsage(parsed.Key) += 1
                        If Not allInvokedUsers.ContainsKey(parsed.Key) Then allInvokedUsers(parsed.Key) = New HashSet(Of String)()
                        allInvokedUsers(parsed.Key).Add(parsed.UserHash)

                        If Not appInvokedUsage(appCode).ContainsKey(parsed.Key) Then appInvokedUsage(appCode)(parsed.Key) = 0
                        appInvokedUsage(appCode)(parsed.Key) += 1
                        If Not appInvokedUsers(appCode).ContainsKey(parsed.Key) Then appInvokedUsers(appCode)(parsed.Key) = New HashSet(Of String)()
                        appInvokedUsers(appCode)(parsed.Key).Add(parsed.UserHash)

                        ' per-day (overall)
                        Dim day As Date = parsed.Time.Date
                        If Not allInvokedUsageByDay.ContainsKey(parsed.Key) Then
                            allInvokedUsageByDay(parsed.Key) = New Dictionary(Of Date, HashSet(Of String))()
                        End If
                        If Not allInvokedUsageByDay(parsed.Key).ContainsKey(day) Then
                            allInvokedUsageByDay(parsed.Key)(day) = New HashSet(Of String)()
                        End If
                        allInvokedUsageByDay(parsed.Key)(day).Add(parsed.UserHash)

                        ' per-day (per app)
                        If Not appInvokedUsageByDay(appCode).ContainsKey(parsed.Key) Then
                            appInvokedUsageByDay(appCode)(parsed.Key) = New Dictionary(Of Date, HashSet(Of String))()
                        End If
                        If Not appInvokedUsageByDay(appCode)(parsed.Key).ContainsKey(day) Then
                            appInvokedUsageByDay(appCode)(parsed.Key)(day) = New HashSet(Of String)()
                        End If
                        appInvokedUsageByDay(appCode)(parsed.Key)(day).Add(parsed.UserHash)

                    Else
                        ' --- Other lines (overall + per-app) ---
                        If Not allOtherUsage.ContainsKey(parsed.Key) Then allOtherUsage(parsed.Key) = 0
                        allOtherUsage(parsed.Key) += 1
                        If Not allOtherUsers.ContainsKey(parsed.Key) Then allOtherUsers(parsed.Key) = New HashSet(Of String)()
                        allOtherUsers(parsed.Key).Add(parsed.UserHash)

                        If Not appOtherUsage(appCode).ContainsKey(parsed.Key) Then appOtherUsage(appCode)(parsed.Key) = 0
                        appOtherUsage(appCode)(parsed.Key) += 1
                        If Not appOtherUsers(appCode).ContainsKey(parsed.Key) Then appOtherUsers(appCode)(parsed.Key) = New HashSet(Of String)()
                        appOtherUsers(appCode)(parsed.Key).Add(parsed.UserHash)
                    End If
                Next
            Next

            Dim sb As New StringBuilder()

            Dim appNameMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
            {"WD", "Word"},
            {"XL", "Excel"},
            {"OL", "Outlook"},
            {"UK", "Unknown"}
        }

            Dim formatAppName As Func(Of String, String) =
            Function(code As String)
                If appNameMap.ContainsKey(code) Then Return appNameMap(code) & " (" & code & ")"
                Return code
            End Function

            Dim appendSection As Action(Of String, HashSet(Of String),
                                  Dictionary(Of String, Integer), Dictionary(Of String, HashSet(Of String)),
                                  Dictionary(Of String, Integer), Dictionary(Of String, HashSet(Of String)),
                                  Dictionary(Of String, Integer), Dictionary(Of String, HashSet(Of String)),
                                  Dictionary(Of String, Dictionary(Of Date, HashSet(Of String)))) =
            Sub(title As String,
                users As HashSet(Of String),
                versionUsage As Dictionary(Of String, Integer), versionUsers As Dictionary(Of String, HashSet(Of String)),
                invokedUsage As Dictionary(Of String, Integer), invokedUsers As Dictionary(Of String, HashSet(Of String)),
                otherUsage As Dictionary(Of String, Integer), otherUsers As Dictionary(Of String, HashSet(Of String)),
                invokedByDay As Dictionary(Of String, Dictionary(Of Date, HashSet(Of String))))

                sb.AppendLine("=== " & title & " ===")
                sb.AppendLine("Total users: " & users.Count)

                sb.AppendLine()
                sb.AppendLine("Version usage (count; unique users):")
                For Each kvp In versionUsage.OrderByDescending(Function(x) x.Value).ThenBy(Function(x) x.Key)
                    Dim u As Integer = If(versionUsers.ContainsKey(kvp.Key), versionUsers(kvp.Key).Count, 0)
                    sb.AppendLine($"{kvp.Key}: {kvp.Value}; users: {u}")
                Next

                sb.AppendLine()
                sb.AppendLine("Functions invoked: count; unique users")
                For Each kvp In invokedUsage.OrderByDescending(Function(x) x.Value).ThenBy(Function(x) x.Key)
                    Dim u As Integer = If(invokedUsers.ContainsKey(kvp.Key), invokedUsers(kvp.Key).Count, 0)
                    sb.AppendLine($"{kvp.Key}: {kvp.Value}; users: {u}")
                Next

                If otherUsage.Count > 0 Then
                    sb.AppendLine()
                    sb.AppendLine("Other log lines (not ending with 'invoked'): count; unique users")
                    For Each kvp In otherUsage.OrderByDescending(Function(x) x.Value).ThenBy(Function(x) x.Key)
                        Dim u As Integer = If(otherUsers.ContainsKey(kvp.Key), otherUsers(kvp.Key).Count, 0)
                        sb.AppendLine($"{kvp.Key}: {kvp.Value}; users: {u}")
                    Next
                End If

                sb.AppendLine()
                sb.AppendLine("Functions invoked per day (unique users):")
                For Each fn In invokedByDay.Keys.OrderBy(Function(x) x)
                    For Each usageDay As Date In invokedByDay(fn).Keys.OrderBy(Function(d) d)
                        sb.AppendLine(fn & " | " & usageDay.ToString("yyyy-MM-dd") &
                                      " | users: " & invokedByDay(fn)(usageDay).Count)
                    Next
                Next

                sb.AppendLine()
            End Sub

            sb.AppendLine("=== RED INK LOG ANALYSIS ===")
            sb.AppendLine("Log path: " & context.INI_LogPath)
            sb.AppendLine("Date range: " &
                      If(startDate.HasValue, startDate.Value.ToString("yyyy-MM-dd"), "N/A") & " to " &
                      If(endDate.HasValue, endDate.Value.ToString("yyyy-MM-dd"), "N/A"))
            sb.AppendLine()

            appendSection(
            "OVERALL",
            allUsers,
            allVersionUsage, allVersionUsers,
            allInvokedUsage, allInvokedUsers,
            allOtherUsage, allOtherUsers,
            allInvokedUsageByDay)

            For Each appCode In appUsers.Keys.OrderBy(Function(x) x)
                appendSection(
                formatAppName(appCode),
                appUsers(appCode),
                appVersionUsage(appCode), appVersionUsers(appCode),
                appInvokedUsage(appCode), appInvokedUsers(appCode),
                appOtherUsage(appCode), appOtherUsers(appCode),
                appInvokedUsageByDay(appCode))
            Next

            ShowCustomWindow($"The following log statistics were compiled based on the content found in {context.INI_LogPath}", sb.ToString(), "", $"{AN} Log Statistics")

        Catch ex As System.Exception
            ShowCustomMessageBox("Error in AnalyzeLogs: " & ex.Message)
        End Try
    End Sub

    '==============================
    ' INTERNAL HELPERS
    '==============================

    Private Function GetOfficeHostCode(name As String) As String
        If String.IsNullOrWhiteSpace(name) Then Return "UK"

        Dim n = name.ToLowerInvariant()

        If n.Contains("excel") Then Return "XL"
        If n.Contains("word") Then Return "WD"
        If n.Contains("outlook") Then Return "OL"
        Return "UK"
    End Function

    Private Function GetFileHash(appCode As String) As String
        Return ComputeHash(Environment.UserName & "|" &
                           Environment.MachineName & "|" &
                           appCode, 8)
    End Function

    Private Function GetUserHash() As String
        Return ComputeHash(Environment.UserName, 8)
    End Function

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


    Private Function ExtractAppCode(filePath As String) As String
        Dim name As String = System.IO.Path.GetFileNameWithoutExtension(filePath)
        Dim parts As String() = name.Split("-"c)
        Return parts(parts.Length - 1)
    End Function

    Private Function ParseDate(input As String) As Nullable(Of DateTime)
        If String.IsNullOrWhiteSpace(input) Then Return Nothing
        Dim dt As DateTime
        If DateTime.TryParseExact(input, "yyyy-MM-dd", CultureInfo.InvariantCulture,
                                  DateTimeStyles.None, dt) Then
            Return dt
        End If
        Return Nothing
    End Function

    Private Class ParsedLine
        Public Time As DateTime
        Public UserHash As String
        Public Version As String

        ' Group key:
        '   - invoked lines: function handler name (e.g. RI_PrimLang2_Click)
        '   - other lines: first token after version field (best-effort grouping)
        Public Key As String

        Public IsInvoked As Boolean
    End Class

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
                ' Typical: "<FunctionName> invoked"
                key = tokens(0)
            Else
                ' Best-effort grouping for non-invoked lines
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

End Module