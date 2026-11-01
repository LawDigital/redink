' Part of "Red Ink" (Outlook)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.AutoPilot.Timing.vb
' PURPOSE
'   Tracks AutoPilot job durations and advisory workload-based timing estimates.
'
' KEY ARCHITECTURE
'   • Samples: Stores bounded historical timings by AutoPilot workload category.
'   • Estimates: Advisory effort class adjusts durations from observed category data.
' =============================================================================

Option Explicit On
Option Strict On

Partial Public Class ThisAddIn
    Private Const AP_TimingHistoryLimit As System.Int32 = 500
    Private Const AP_TimingRecentLimit As System.Int32 = 30
    Private Const AP_EffortTimingMinimumSamples As System.Int32 = 20
    Private _apCurrentJobEffortClass As System.String = Nothing
    Private ReadOnly _apEffortTimingHistory As New System.Collections.Generic.Dictionary(Of System.String, System.Collections.Generic.List(Of System.Double))(System.StringComparer.OrdinalIgnoreCase)
    Private ReadOnly _apTimingSyncRoot As New System.Object()
    Private ReadOnly _apTimingHistory As New System.Collections.Generic.Dictionary(Of System.String, System.Collections.Generic.List(Of System.Double))(System.StringComparer.OrdinalIgnoreCase)
    Private _apTimingHistoryLoaded As System.Boolean = False

    Private Sub RecordAutoPilotJobEffort(expectedEntryId As System.String, effortClass As System.String)
        If Not System.String.Equals(expectedEntryId, _apCurrentProcessingEntryId, System.StringComparison.Ordinal) Then Return
        Dim canonical As System.String = SharedLibrary.Agents.RunEffortEstimate.Normalize(effortClass)
        If canonical.Length = 0 Then Return
        If System.Threading.Interlocked.CompareExchange(Of System.String)(_apCurrentJobEffortClass, canonical, Nothing) IsNot Nothing Then Return
        Dim category As System.String = If(_apCurrentProcessingCategory, AP_Cat_TextOnly)
        Dim estimate As System.Double = GetEffortAdjustedEstimate(category, canonical, GetCategoryEstimate(category))
        ApDashboardLog($"⏱ Initial workload estimate: {canonical}; estimated total duration {System.Math.Ceiling(estimate / 60.0):F0} min (advisory).", "step")
    End Sub

    Private Function GetEffortAdjustedEstimate(category As System.String, effortClass As System.String, baseline As System.Double) As System.Double
        Dim canonical As System.String = SharedLibrary.Agents.RunEffortEstimate.Normalize(effortClass)
        If canonical.Length = 0 Then Return baseline
        Try
            EnsureAutoPilotTimingHistoryLoaded()
            Dim seed As System.Double
            Select Case category
                Case AP_Cat_HeavyPdf : seed = AP_SeedSeconds_HeavyPdf
                Case AP_Cat_HeavyDoc : seed = AP_SeedSeconds_HeavyDoc
                Case AP_Cat_Light : seed = AP_SeedSeconds_LightAttachments
                Case Else : seed = AP_SeedSeconds_TextOnly
            End Select
            SyncLock _apTimingSyncRoot
                Dim samples As System.Collections.Generic.List(Of System.Double) = Nothing
                If _apEffortTimingHistory.TryGetValue(category & "|" & canonical, samples) AndAlso samples.Count >= AP_EffortTimingMinimumSamples Then
                    Dim recentCount As System.Int32 = System.Math.Min(AP_TimingRecentLimit, samples.Count)
                    Dim broad As System.Double = GetAutoPilotTimingPercentile(samples)
                    Dim recent As System.Double = GetAutoPilotTimingPercentile(samples.GetRange(samples.Count - recentCount, recentCount))
                    Return System.Math.Max(seed, System.Math.Max(broad, recent) * 1.25)
                End If
            End SyncLock
            ' Provisional floors only: do not multiply an already learned mixed-category estimate.
            ' Once 20 comparable observations exist, measured durations replace these factors.
            Dim factor As System.Double = 1.0
            Select Case canonical
                Case "medium" : factor = 1.5
                Case "extensive" : factor = 3.0
                Case "very_extensive" : factor = 6.0
            End Select
            Return System.Math.Max(baseline, seed * factor)
        Catch ex As System.Exception
            System.Diagnostics.Trace.TraceWarning("Workload adjustment unavailable; using the category estimate: " & ex.Message)
            Return baseline
        End Try
    End Function

    Private Shared Function ReadAutoPilotEffortHistory(document As Newtonsoft.Json.Linq.JObject) As System.Collections.Generic.Dictionary(Of System.String, System.Collections.Generic.List(Of System.Double))
        Dim loaded As New System.Collections.Generic.Dictionary(Of System.String, System.Collections.Generic.List(Of System.Double))(System.StringComparer.OrdinalIgnoreCase)
        Try
            Dim categories As Newtonsoft.Json.Linq.JObject = TryCast(document("effortCategories"), Newtonsoft.Json.Linq.JObject)
            If categories Is Nothing Then Return loaded
            For Each category As System.String In New System.String() {AP_Cat_TextOnly, AP_Cat_Light, AP_Cat_HeavyDoc, AP_Cat_HeavyPdf}
                For Each effort As System.String In New System.String() {"short", "medium", "extensive", "very_extensive"}
                    Dim key As System.String = category & "|" & effort
                    Dim values As Newtonsoft.Json.Linq.JArray = TryCast(categories(key), Newtonsoft.Json.Linq.JArray)
                    If values Is Nothing Then Continue For
                    Dim samples As New System.Collections.Generic.List(Of System.Double)()
                    For i As System.Int32 = System.Math.Max(0, values.Count - AP_TimingHistoryLimit) To values.Count - 1
                        Dim token As Newtonsoft.Json.Linq.JToken = values(i)
                        If token.Type <> Newtonsoft.Json.Linq.JTokenType.Float AndAlso token.Type <> Newtonsoft.Json.Linq.JTokenType.Integer Then Continue For
                        Dim seconds As System.Double = CType(token, System.Double)
                        If IsAutoPilotTimingSample(seconds) Then samples.Add(seconds)
                    Next
                    loaded(key) = samples
                Next
            Next
        Catch ex As System.Exception
            ' A damaged optional index must not discard the existing category history.
            loaded.Clear()
            System.Diagnostics.Trace.TraceWarning("Optional workload history was ignored: " & ex.Message)
        End Try
        Return loaded
    End Function

    Private Shared Function GetAutoPilotTimingHistoryPath() As System.String
        Return System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData),
                                      "redink", "autopilot-timing-v1.json")
    End Function

    Private Shared Function IsAutoPilotTimingCategory(category As System.String) As System.Boolean
        Return category = AP_Cat_TextOnly OrElse category = AP_Cat_Light OrElse
               category = AP_Cat_HeavyDoc OrElse category = AP_Cat_HeavyPdf
    End Function

    Private Shared Function IsAutoPilotTimingSample(seconds As System.Double) As System.Boolean
        Return seconds > 0 AndAlso Not System.Double.IsNaN(seconds) AndAlso Not System.Double.IsInfinity(seconds)
    End Function

    Private Sub EnsureAutoPilotTimingHistoryLoaded()
        SyncLock _apTimingSyncRoot
            If _apTimingHistoryLoaded Then Return
            _apTimingHistoryLoaded = True
            Dim path As System.String = GetAutoPilotTimingHistoryPath()
            For Each candidate As System.String In New System.String() {path, path & ".bak"}
                If Not System.IO.File.Exists(candidate) Then Continue For
                Try
                    If New System.IO.FileInfo(candidate).Length > 262144 Then Throw New System.IO.InvalidDataException("Timing history is too large.")
                    Dim document As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(System.IO.File.ReadAllText(candidate, System.Text.Encoding.UTF8))
                    If document("version") Is Nothing OrElse document("version").Type <> Newtonsoft.Json.Linq.JTokenType.Integer OrElse
                       CType(document("version"), System.Int32) <> 1 Then Throw New System.IO.InvalidDataException("Unsupported timing-history version.")
                    Dim categories As Newtonsoft.Json.Linq.JObject = TryCast(document("categories"), Newtonsoft.Json.Linq.JObject)
                    If categories Is Nothing Then Throw New System.IO.InvalidDataException("Missing timing categories.")
                    Dim loaded As New System.Collections.Generic.Dictionary(Of System.String, System.Collections.Generic.List(Of System.Double))(System.StringComparer.OrdinalIgnoreCase)
                    For Each category As System.String In New System.String() {AP_Cat_TextOnly, AP_Cat_Light, AP_Cat_HeavyDoc, AP_Cat_HeavyPdf}
                        Dim values As Newtonsoft.Json.Linq.JArray = TryCast(categories(category), Newtonsoft.Json.Linq.JArray)
                        If values Is Nothing Then Continue For
                        Dim samples As New System.Collections.Generic.List(Of System.Double)()
                        For i As System.Int32 = System.Math.Max(0, values.Count - AP_TimingHistoryLimit) To values.Count - 1
                            Dim token As Newtonsoft.Json.Linq.JToken = values(i)
                            If token.Type <> Newtonsoft.Json.Linq.JTokenType.Float AndAlso token.Type <> Newtonsoft.Json.Linq.JTokenType.Integer Then Continue For
                            Dim seconds As System.Double = CType(token, System.Double)
                            If IsAutoPilotTimingSample(seconds) Then samples.Add(seconds)
                        Next
                        loaded(category) = samples
                    Next
                    ' Publish only a fully parsed snapshot. A corrupt primary may fall back to its backup.
                    Dim loadedEffort As System.Collections.Generic.Dictionary(Of System.String, System.Collections.Generic.List(Of System.Double)) = ReadAutoPilotEffortHistory(document)
                    _apTimingHistory.Clear()
                    For Each entry In loaded
                        _apTimingHistory(entry.Key) = entry.Value
                    Next
                    _apEffortTimingHistory.Clear()
                    For Each entry In loadedEffort
                        _apEffortTimingHistory(entry.Key) = entry.Value
                    Next
                    ApDashboardLog("⏱ Duration history loaded (up to 500 completed jobs per category / workload class).", "step")
                    Return
                Catch ex As System.Exception
                    ApDashboardLog("⚠ Duration history could not be read: " & ex.Message, "warn")
                End Try
            Next
        End SyncLock
    End Sub

    ''' <summary>Upper-range estimate from broad and recent history, with a seed floor and 25% margin.
    ''' This is an expected TOTAL duration, never a countdown or a reason to abort work.</summary>
    Private Function GetCategoryEstimate(category As System.String) As System.Double
        Dim seed As System.Double
        Select Case category
            Case AP_Cat_HeavyPdf : seed = AP_SeedSeconds_HeavyPdf
            Case AP_Cat_HeavyDoc : seed = AP_SeedSeconds_HeavyDoc
            Case AP_Cat_Light : seed = AP_SeedSeconds_LightAttachments
            Case Else : seed = AP_SeedSeconds_TextOnly
        End Select
        EnsureAutoPilotTimingHistoryLoaded()
        SyncLock _apTimingSyncRoot
            Dim samples As System.Collections.Generic.List(Of System.Double) = Nothing
            If Not _apTimingHistory.TryGetValue(If(category, AP_Cat_TextOnly), samples) OrElse samples.Count = 0 Then Return seed
            Dim recentCount As System.Int32 = System.Math.Min(AP_TimingRecentLimit, samples.Count)
            Dim broad As System.Double = GetAutoPilotTimingPercentile(samples)
            Dim recent As System.Double = GetAutoPilotTimingPercentile(samples.GetRange(samples.Count - recentCount, recentCount))
            Return System.Math.Max(seed, System.Math.Max(broad, recent) * 1.25)
        End SyncLock
    End Function

    Private Shared Function GetAutoPilotTimingPercentile(samples As System.Collections.Generic.List(Of System.Double)) As System.Double
        Dim ordered As System.Double() = samples.ToArray()
        System.Array.Sort(ordered)
        Dim index As System.Int32 = CInt(System.Math.Ceiling(ordered.Length * 0.8)) - 1
        Return ordered(System.Math.Max(0, index))
    End Function

    Private Sub UpdateCategoryTiming(category As System.String, elapsedSeconds As System.Double, Optional effortClass As System.String = Nothing)
        If Not IsAutoPilotTimingCategory(category) OrElse Not IsAutoPilotTimingSample(elapsedSeconds) Then Return
        EnsureAutoPilotTimingHistoryLoaded()
        Try
            SyncLock _apTimingSyncRoot
                Dim samples As System.Collections.Generic.List(Of System.Double) = Nothing
                If Not _apTimingHistory.TryGetValue(category, samples) Then
                    samples = New System.Collections.Generic.List(Of System.Double)()
                    _apTimingHistory(category) = samples
                End If
                samples.Add(elapsedSeconds)
                If samples.Count > AP_TimingHistoryLimit Then samples.RemoveRange(0, samples.Count - AP_TimingHistoryLimit)
                Dim canonical As System.String = SharedLibrary.Agents.RunEffortEstimate.Normalize(effortClass)
                If canonical.Length > 0 Then
                    Dim key As System.String = category & "|" & canonical
                    Dim effortSamples As System.Collections.Generic.List(Of System.Double) = Nothing
                    If Not _apEffortTimingHistory.TryGetValue(key, effortSamples) Then
                        effortSamples = New System.Collections.Generic.List(Of System.Double)()
                        _apEffortTimingHistory(key) = effortSamples
                    End If
                    effortSamples.Add(elapsedSeconds)
                    If effortSamples.Count > AP_TimingHistoryLimit Then effortSamples.RemoveRange(0, effortSamples.Count - AP_TimingHistoryLimit)
                End If
                PersistAutoPilotTimingHistory()
            End SyncLock
        Catch ex As System.Exception
            ' Keep the learned data in memory; a later completion retries persistence.
            ApDashboardLog("⚠ Duration history could not be saved: " & ex.Message, "warn")
        End Try
    End Sub

    ' Caller holds _apTimingSyncRoot. No mail content, identities or credentials are written.
    Private Sub PersistAutoPilotTimingHistory()
        Dim path As System.String = GetAutoPilotTimingHistoryPath()
        System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(path))
        Dim tempPath As System.String = path & "." & System.Guid.NewGuid().ToString("N") & ".tmp"
        Try
            Dim document As New Newtonsoft.Json.Linq.JObject()
            document("version") = New Newtonsoft.Json.Linq.JValue(1)
            document("categories") = Newtonsoft.Json.Linq.JObject.FromObject(_apTimingHistory)
            ' Additive optional index: original v1 category arrays remain readable without it.
            document("effortCategories") = Newtonsoft.Json.Linq.JObject.FromObject(_apEffortTimingHistory)
            Dim bytes As System.Byte() = System.Text.Encoding.UTF8.GetBytes(document.ToString(Newtonsoft.Json.Formatting.None))
            Using stream As New System.IO.FileStream(tempPath, System.IO.FileMode.CreateNew, System.IO.FileAccess.Write, System.IO.FileShare.None)
                stream.Write(bytes, 0, bytes.Length)
                stream.Flush(True)
            End Using
            If System.IO.File.Exists(path) Then
                System.IO.File.Replace(tempPath, path, path & ".bak")
            Else
                System.IO.File.Move(tempPath, path)
            End If
        Finally
            If System.IO.File.Exists(tempPath) Then
                Try
                    System.IO.File.Delete(tempPath)
                Catch ex As System.Exception
                    System.Diagnostics.Trace.TraceWarning("AutoPilot timing temporary-file cleanup failed: " & ex.Message)
                End Try
            End If
        End Try
    End Sub
End Class
