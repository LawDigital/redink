' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: PerformanceLogger.vb
' Purpose: Low-overhead, APIDebug-gated performance logging shared by all Office hosts.
'
' Architecture:
'  - Disabled unless ISharedContext.INI_APIDebug is True.
'  - Callers enqueue already-formatted timing records only; no file I/O occurs on the caller thread.
'  - A single background writer per process drains the queue and appends to a host-specific log under
'    %LOCALAPPDATA%\redink.
'  - Log growth is bounded in the background.
' =============================================================================

Option Strict On
Option Explicit On

Namespace SharedLibrary

    Public NotInheritable Class PerformanceLogger

        Private Const MaxLogBytes As System.Int64 = 2L * 1024L * 1024L
        Private Const KeepLogLines As System.Int32 = 4000

        Private Shared ReadOnly PendingEntries As New System.Collections.Concurrent.ConcurrentQueue(Of PerformanceLogEntry)()
        Private Shared ReadOnly ProcessId As System.Int32 = System.Diagnostics.Process.GetCurrentProcess().Id
        Private Shared _writerScheduled As System.Int32 = 0

        Private Sub New()
        End Sub

        Private NotInheritable Class PerformanceLogEntry
            Public Property HostName As System.String
            Public Property Text As System.String
        End Class

        Public Shared Function IsEnabled(context As SharedContext.ISharedContext) As System.Boolean
            Return context IsNot Nothing AndAlso context.INI_APIDebug
        End Function

        Public Shared Sub LogDuration(context As SharedContext.ISharedContext,
                                      category As System.String,
                                      operation As System.String,
                                      elapsedMilliseconds As System.Int64,
                                      Optional details As System.String = Nothing,
                                      Optional hostName As System.String = Nothing)
            If Not IsEnabled(context) Then Return

            Dim resolvedHost As System.String = ResolveHostName(context, hostName)
            Dim line As System.String = BuildLine(
                category,
                operation,
                "durationMs=" & System.Math.Max(0L, elapsedMilliseconds).ToString(System.Globalization.CultureInfo.InvariantCulture),
                details)
            Enqueue(resolvedHost, line)
        End Sub

        Public Shared Sub LogEvent(context As SharedContext.ISharedContext,
                                   category As System.String,
                                   operation As System.String,
                                   Optional details As System.String = Nothing,
                                   Optional hostName As System.String = Nothing)
            If Not IsEnabled(context) Then Return

            Dim resolvedHost As System.String = ResolveHostName(context, hostName)
            Enqueue(resolvedHost, BuildLine(category, operation, Nothing, details))
        End Sub

        Public Shared Sub LogStartupSnapshot(context As SharedContext.ISharedContext,
                                             hostName As System.String,
                                             version As System.String,
                                             timings As System.Collections.Generic.IEnumerable(Of System.String))
            If Not IsEnabled(context) OrElse timings Is Nothing Then Return

            Dim snapshot As New System.Collections.Generic.List(Of System.String)()
            For Each timing As System.String In timings
                If Not System.String.IsNullOrWhiteSpace(timing) Then
                    snapshot.Add(timing)
                End If
            Next
            If snapshot.Count = 0 Then Return

            Dim resolvedHost As System.String = ResolveHostName(context, hostName)
            Dim builder As New System.Text.StringBuilder()
            builder.AppendLine(BuildLine("Startup", "Snapshot.begin", Nothing, "version=" & SanitizeValue(version)))
            For Each timing As System.String In snapshot
                builder.AppendLine(BuildLine("Startup", "Step", Nothing, timing))
            Next
            builder.AppendLine(BuildLine("Startup", "Snapshot.end", Nothing, Nothing))
            Enqueue(resolvedHost, builder.ToString().TrimEnd(System.Environment.NewLine.ToCharArray()))
        End Sub

        Public Shared Sub Measure(context As SharedContext.ISharedContext,
                                  category As System.String,
                                  operation As System.String,
                                  action As System.Action,
                                  Optional details As System.String = Nothing,
                                  Optional hostName As System.String = Nothing)
            If action Is Nothing Then Return
            If Not IsEnabled(context) Then
                action.Invoke()
                Return
            End If

            Dim stopwatch As System.Diagnostics.Stopwatch = System.Diagnostics.Stopwatch.StartNew()
            Try
                action.Invoke()
            Finally
                stopwatch.Stop()
                LogDuration(context, category, operation, stopwatch.ElapsedMilliseconds, details, hostName)
            End Try
        End Sub

        Public Shared Function Measure(Of TResult)(context As SharedContext.ISharedContext,
                                                   category As System.String,
                                                   operation As System.String,
                                                   callback As System.Func(Of TResult),
                                                   Optional details As System.String = Nothing,
                                                   Optional hostName As System.String = Nothing) As TResult
            If callback Is Nothing Then Return Nothing
            If Not IsEnabled(context) Then
                Return callback.Invoke()
            End If

            Dim stopwatch As System.Diagnostics.Stopwatch = System.Diagnostics.Stopwatch.StartNew()
            Try
                Return callback.Invoke()
            Finally
                stopwatch.Stop()
                LogDuration(context, category, operation, stopwatch.ElapsedMilliseconds, details, hostName)
            End Try
        End Function

        Private Shared Function BuildLine(category As System.String,
                                          operation As System.String,
                                          primaryDetails As System.String,
                                          additionalDetails As System.String) As System.String
            Dim parts As New System.Collections.Generic.List(Of System.String) From {
                System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff", System.Globalization.CultureInfo.InvariantCulture),
                "processId=" & ProcessId.ToString(System.Globalization.CultureInfo.InvariantCulture),
                "threadId=" & System.Threading.Thread.CurrentThread.ManagedThreadId.ToString(System.Globalization.CultureInfo.InvariantCulture),
                "threadPool=" & System.Threading.Thread.CurrentThread.IsThreadPoolThread.ToString(),
                "category=" & SanitizeToken(category, "General"),
                "operation=" & SanitizeToken(operation, "Unknown")
            }

            If Not System.String.IsNullOrWhiteSpace(primaryDetails) Then
                parts.Add(primaryDetails)
            End If
            If Not System.String.IsNullOrWhiteSpace(additionalDetails) Then
                parts.Add(SanitizeValue(additionalDetails))
            End If

            Return System.String.Join("; ", parts)
        End Function

        Private Shared Sub Enqueue(hostName As System.String, text As System.String)
            If System.String.IsNullOrWhiteSpace(text) Then Return

            PendingEntries.Enqueue(
                New PerformanceLogEntry With {
                    .HostName = SanitizeToken(hostName, "Shared"),
                    .Text = text
                })
            ScheduleWriter()
        End Sub

        Private Shared Sub ScheduleWriter()
            If System.Threading.Interlocked.CompareExchange(_writerScheduled, 1, 0) <> 0 Then Return

            System.Threading.Tasks.Task.Run(
                Sub()
                    Try
                        DrainQueue()
                    Catch ex As System.Exception
                        System.Diagnostics.Debug.WriteLine("[PERF] Background performance logger failed: " & ex.Message)
                    Finally
                        System.Threading.Interlocked.Exchange(_writerScheduled, 0)
                        If Not PendingEntries.IsEmpty Then
                            ScheduleWriter()
                        End If
                    End Try
                End Sub)
        End Sub

        Private Shared Sub DrainQueue()
            Dim grouped As New System.Collections.Generic.Dictionary(Of System.String, System.Text.StringBuilder)(System.StringComparer.OrdinalIgnoreCase)
            Dim entry As PerformanceLogEntry = Nothing

            While PendingEntries.TryDequeue(entry)
                Dim builder As System.Text.StringBuilder = Nothing
                If Not grouped.TryGetValue(entry.HostName, builder) Then
                    builder = New System.Text.StringBuilder()
                    grouped(entry.HostName) = builder
                End If
                builder.AppendLine(entry.Text)
            End While

            For Each item As System.Collections.Generic.KeyValuePair(Of System.String, System.Text.StringBuilder) In grouped
                Try
                    WriteBatch(item.Key, item.Value.ToString())
                Catch ex As System.Exception
                    System.Diagnostics.Debug.WriteLine("[PERF] Background performance log write failed: " & ex.Message)
                End Try
            Next
        End Sub

        Private Shared Sub WriteBatch(hostName As System.String, text As System.String)
            Dim basePath As System.String = System.IO.Path.Combine(
                System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData),
                "redink")
            System.IO.Directory.CreateDirectory(basePath)

            Dim logPath As System.String = System.IO.Path.Combine(basePath, "RI_" & hostName & "_Performance.log")
            Dim mutexName As System.String = "Local\RedInk_Performance_" & SanitizeToken(hostName, "Shared")
            Dim mutex As New System.Threading.Mutex(False, mutexName)
            Dim acquired As System.Boolean = False

            Try
                Try
                    acquired = mutex.WaitOne(System.TimeSpan.FromSeconds(5))
                Catch ex As System.Threading.AbandonedMutexException
                    acquired = True
                End Try

                If Not acquired Then Return

                TrimLogIfNeeded(logPath)
                System.IO.File.AppendAllText(logPath, text, System.Text.Encoding.UTF8)
            Finally
                If acquired Then
                    Try
                        mutex.ReleaseMutex()
                    Catch ex As System.Exception
                    End Try
                End If
                mutex.Dispose()
            End Try
        End Sub

        Private Shared Sub TrimLogIfNeeded(logPath As System.String)
            If Not System.IO.File.Exists(logPath) Then Return

            Dim fileInfo As New System.IO.FileInfo(logPath)
            If fileInfo.Length <= MaxLogBytes Then Return

            Dim lines As System.String() = System.IO.File.ReadAllLines(logPath, System.Text.Encoding.UTF8)
            Dim keepFrom As System.Int32 = System.Math.Max(0, lines.Length - KeepLogLines)
            Dim keepCount As System.Int32 = lines.Length - keepFrom
            If keepCount <= 0 Then
                System.IO.File.WriteAllText(logPath, System.String.Empty, System.Text.Encoding.UTF8)
                Return
            End If

            Dim kept(keepCount - 1) As System.String
            System.Array.Copy(lines, keepFrom, kept, 0, keepCount)
            System.IO.File.WriteAllLines(logPath, kept, System.Text.Encoding.UTF8)
        End Sub

        Private Shared Function ResolveHostName(context As SharedContext.ISharedContext, explicitHostName As System.String) As System.String
            If Not System.String.IsNullOrWhiteSpace(explicitHostName) Then
                Return SanitizeToken(explicitHostName, "Shared")
            End If

            Dim rdv As System.String = If(context?.RDV, System.String.Empty)
            If rdv.IndexOf("Excel", System.StringComparison.OrdinalIgnoreCase) >= 0 Then Return "Excel"
            If rdv.IndexOf("Word", System.StringComparison.OrdinalIgnoreCase) >= 0 Then Return "Word"
            If rdv.IndexOf("Outlook", System.StringComparison.OrdinalIgnoreCase) >= 0 Then Return "Outlook"
            Return "Shared"
        End Function

        Private Shared Function SanitizeToken(value As System.String, fallback As System.String) As System.String
            Dim candidate As System.String = If(value, System.String.Empty).Trim()
            If candidate.Length = 0 Then candidate = fallback

            Dim builder As New System.Text.StringBuilder(candidate.Length)
            For Each character As System.Char In candidate
                If System.Char.IsLetterOrDigit(character) OrElse character = "_"c OrElse character = "-"c OrElse character = "."c Then
                    builder.Append(character)
                Else
                    builder.Append("_"c)
                End If
            Next
            Return builder.ToString()
        End Function

        Private Shared Function SanitizeValue(value As System.String) As System.String
            If value Is Nothing Then Return System.String.Empty
            Return value.Replace(vbCr, " ").Replace(vbLf, " ").Trim()
        End Function

    End Class

End Namespace
