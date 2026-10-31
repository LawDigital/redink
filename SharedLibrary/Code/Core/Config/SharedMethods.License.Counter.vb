' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: SharedMethods.License.Counter.vb
' Purpose: Durable monthly licence counter with sparse daily add-in load observations.
'
' Scope:
'  - Creates one local usage marker per Month + JointNetworkKey + ProductID + HostID + UserKey.
'  - Records observed local load dates; reports one user/month aggregate to each target.
'  - Retries failures without resending successful daily evidence across Office hosts.
'  - Stores lightweight delivery state and pending error events.
'  - Provides a simple evaluator UI for Word via the "licensecounter" freestyle command.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Net
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary
    Partial Public Class SharedMethods

#Region "License Counter"

        Private Const LicenseCounterFormatVersion As Integer = 1
        Private Const LicenseCounterMarkerPrefix As String = "RI_LC_"
        Private Const LicenseCounterReportPrefix As String = "RI_LC_REPORT_"
        Private Const LicenseCounterStateFileName As String = "RI_LC_State.json"
        Private Const LicenseCounterRetentionMonths As Integer = 24
        Private Const LicenseCounterWebTimeoutMs As Integer = 5000
        Private Const LicenseCounterRetryHours As Integer = 24
        Public Const LicenseCounterOutlookTimerHours As Integer = 6

        ' Reporting never calls Office or displays UI. This in-process host gate and
        ' a separate background-only profile mutex prevent overlapping deliveries.
        ' Local monthly markers are persisted before either transport gate.
        Private Shared ReadOnly LicenseCounterReportBusy As New System.Collections.Concurrent.ConcurrentDictionary(Of System.String, System.Byte)(System.StringComparer.OrdinalIgnoreCase)

        Private NotInheritable Class LicenseCounterRuntimeOptions
            Public Property PathRaw As System.String = ""
            Public Property MethodRaw As System.String = ""
            Public Property Anon As System.Boolean
            Public Property ReportIntervalMinutes As System.Int32 = DEFAULT_LICENSECOUNTERINTERVALL
            Public Property RequireAck As System.Boolean
            Public Property ClassicNetworkScope As System.String = ""
            ' Snapshot at startup; timer callbacks must not read UI/context config.
            Public Property ConfiguredProductId As System.String = ""
            Public Property ConfiguredUserId As System.String = ""
        End Class

        Private NotInheritable Class LicenseCounterReporter
            Implements System.IDisposable

            Private ReadOnly _timer As System.Threading.Timer
            Private ReadOnly _options As LicenseCounterRuntimeOptions
            Private ReadOnly _hostId As System.String
            Private _disposed As System.Int32

            Public Sub New(options As LicenseCounterRuntimeOptions, hostId As System.String)
                _options = options
                _hostId = NormalizeLicenseCounterHostId(hostId)
                _timer = New System.Threading.Timer(
                    Sub(state As System.Object)
                        If System.Threading.Volatile.Read(_disposed) <> 0 Then Return
                        Try
                            RegisterLicenseCounterUsageAndReportOptions(_options, _hostId)
                        Finally
                            ScheduleNextCheck()
                        End Try
                    End Sub,
                    Nothing,
                    System.Threading.Timeout.InfiniteTimeSpan,
                    System.Threading.Timeout.InfiniteTimeSpan)
                ' Record the startup date before background transport. Reusing scalar
                ' configuration avoids accessing Office objects on timer threads.
                RegisterLicenseCounterUsageAndReportOptions(_options, _hostId)
                ScheduleNextCheck()
            End Sub

            Private Sub ScheduleNextCheck()
                If _options.ReportIntervalMinutes = 0 OrElse System.Threading.Volatile.Read(_disposed) <> 0 Then Return
                Try
                    _timer.Change(GetLicenseCounterNextCheckDelay(_options.ReportIntervalMinutes, System.DateTime.Now), System.Threading.Timeout.InfiniteTimeSpan)
                Catch ex As System.ObjectDisposedException
                    ' Disposal may race the final callback; no rescheduling is needed.
                End Try
            End Sub

            Public Sub Dispose() Implements System.IDisposable.Dispose
                If System.Threading.Interlocked.Exchange(_disposed, 1) = 0 Then _timer.Dispose()
            End Sub
        End Class

        Private Shared Function GetLicenseCounterNextCheckDelay(intervalMinutes As System.Int32, nowLocal As System.DateTime) As System.TimeSpan
            If intervalMinutes = 0 Then Return System.Threading.Timeout.InfiniteTimeSpan
            ' A long configured retry interval must not skip a loaded calendar day.
            ' Observe just after local midnight; UTC subtraction also handles DST.
            Dim untilNextDay As System.TimeSpan = nowLocal.Date.AddDays(1).AddSeconds(1).ToUniversalTime().Subtract(nowLocal.ToUniversalTime())
            Dim interval As System.TimeSpan = System.TimeSpan.FromMinutes(intervalMinutes)
            Dim due As System.TimeSpan = If(untilNextDay < interval, untilNextDay, interval)
            Return If(due > System.TimeSpan.Zero, due, System.TimeSpan.FromSeconds(1))
        End Function

        ' <summary>
        ' Starts durable monthly registration with daily load checks and retry polling. Own and dispose
        ' the returned token in the host lifecycle. Configuration is captured on the caller
        ' thread; callbacks never read Office objects or invoke licence dialogs.
        ' </summary>
        Public Shared Function StartLicenseCounterReporting(context As ISharedContext, hostId As System.String) As System.IDisposable
            Try
                Dim options As LicenseCounterRuntimeOptions = ReadLicenseCounterRuntimeOptions(context)
                If options Is Nothing OrElse System.String.IsNullOrWhiteSpace(options.PathRaw) Then Return Nothing
                LogLicenseCounterDiagnostic("reporter_started", NormalizeLicenseCounterHostId(hostId) & "; intervalMinutes=" & options.ReportIntervalMinutes.ToString(System.Globalization.CultureInfo.InvariantCulture))
                Return New LicenseCounterReporter(options, hostId)
            Catch ex As System.Exception
                LogLicenseCounterDiagnostic("reporter_start_error", ex.GetType().Name)
                LogLicenseCounterInternal("reporter_start_error", ex.Message)
                Return Nothing
            End Try
        End Function

        Private Shared Function ReadLicenseCounterRuntimeOptions(context As ISharedContext) As LicenseCounterRuntimeOptions
            If context Is Nothing Then Return Nothing
            Dim options As New LicenseCounterRuntimeOptions() With {
                .PathRaw = context.INI_LicenseCounterPath,
                .MethodRaw = context.INI_LicenseCounterMethod,
                .Anon = context.INI_LicenseCounterAnon
            }
            options.ReportIntervalMinutes = ParseLicenseCounterIntervall(context.INI_LicenseCounterIntervall.ToString(System.Globalization.CultureInfo.InvariantCulture))
            options.ClassicNetworkScope = GetLicenseCounterOption(options.MethodRaw, "NetworkScope")
            ' License counting is independent of whether a Pro/private licence is present
            ' or the normal licence verification was explicitly disabled.
            If _licenseConfigDict IsNot Nothing Then
                Dim configured As System.String = Nothing
                If _licenseConfigDict.TryGetValue("LicenseProductID", configured) Then options.ConfiguredProductId = NormalizeLicenseCounterValue(configured)
                configured = Nothing
                If _licenseConfigDict.TryGetValue("LicenseUserID", configured) Then options.ConfiguredUserId = NormalizeLicenseCounterValue(configured)
            End If
            Dim requireAck As System.String = GetLicenseCounterOption(options.MethodRaw, "RequireAck")
            If Not System.String.IsNullOrWhiteSpace(requireAck) Then
                Select Case requireAck.Trim().ToLowerInvariant()
                    Case "true", "1"
                        options.RequireAck = True
                    Case "false", "0"
                        options.RequireAck = False
                    Case Else
                        Throw New System.FormatException("LicenseCounterMethod RequireAck must be true or false; reporting was not started.")
                End Select
            End If
            Return options
        End Function

        Private Shared Function GetLicenseCounterOption(rawValue As System.String, optionName As System.String) As System.String
            Dim result As System.String = ""
            Dim parts As System.String() = If(rawValue, "").Split(";"c)
            For i As System.Int32 = 1 To parts.Length - 1
                Dim part As System.String = parts(i).Trim()
                Dim equalsIndex As System.Int32 = part.IndexOf("="c)
                If equalsIndex <= 0 Then Continue For
                If System.String.Equals(part.Substring(0, equalsIndex).Trim(), optionName, System.StringComparison.OrdinalIgnoreCase) Then
                    result = part.Substring(equalsIndex + 1).Trim()
                End If
            Next
            Return result
        End Function

        ' Minutes; zero disables the periodic timer but retains startup reporting.
        Public Shared Function ParseLicenseCounterIntervall(value As System.String) As System.Int32
            If System.String.IsNullOrWhiteSpace(value) Then Return DEFAULT_LICENSECOUNTERINTERVALL
            Dim parsed As System.Int32
            If System.Int32.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, parsed) AndAlso parsed >= 0 AndAlso parsed <= 10080 Then Return parsed
            Throw New System.FormatException("LicenseCounterIntervall must be an integer from 0 to 10080 minutes.")
        End Function

        Private Enum LicenseCounterMethodKind
            Auto = 0
            FileOnly = 1
            WebBasic = 2
            WebExtended = 3
            LocalOnly = 4
            AppendFile = 5
            WebPathOnly = 6
        End Enum

        Private Enum LicenseCounterWebModeKind
            Basic = 0
            PathOnly = 1
            Extended = 2
        End Enum

        Private Enum LicenseCounterTargetType
            Invalid = 0
            Web = 1
            Directory = 2
            AppendFile = 3
        End Enum

        Private Class LicenseCounterTarget
            Public Property RawValue As String = ""
            Public Property NormalizedValue As String = ""
            Public Property TargetType As LicenseCounterTargetType
            Public Property TargetHash As String = ""
            Public Property Note As String = ""
        End Class

        Private Class LicenseCounterRuntimeContext
            Public Property BaseFolder As String = ""
            Public Property PathRaw As String = ""
            Public Property MethodRaw As String = ""
            Public Property MethodKind As LicenseCounterMethodKind
            Public Property ReportIntervalMinutes As System.Int32 = DEFAULT_LICENSECOUNTERINTERVALL
            Public Property RequireAck As System.Boolean
            Public Property WebMode As LicenseCounterWebModeKind
            Public Property Targets As New List(Of LicenseCounterTarget)()
            Public Property Month As String = ""
            Public Property ObservedLocalDate As System.String = ""
            Public Property JointNetworkId As String = ""
            Public Property JointNetworkKey As String = ""
            Public Property NetworkScope As String = ""
            Public Property AllowedNetworkIds As New List(Of String)()
            Public Property ActualNetworkId As String = ""
            Public Property ActualNetworkKey As String = ""
            Public Property ProductId As String = ""
            Public Property ProductIdSafe As String = ""
            Public Property HostId As String = ""
            Public Property UserId As String = ""
            Public Property UserKey As String = ""
            Public Property BillingEventId As String = ""
            Public Property HostEventId As String = ""
            Public Property FirstUseUtc As String = ""
            Public Property Anon As Boolean
            Public Property UserObf As String = ""
            Public Property EmptyUserIdFallbackUsed As Boolean
            Public Property MarkerFileName As String = ""
            Public Property MarkerFilePath As String = ""
            Public Property ResubmitInstruction As LicenseCounterResubmitInstruction
            Public Property MarkerRecord As LicenseCounterMarkerRecord
        End Class

        Private Class LicenseCounterMarkerRecord
            Public Property v As Integer = LicenseCounterFormatVersion
            Public Property t As String = "U"
            Public Property month As String = ""
            Public Property networkKey As String = ""
            Public Property networkScope As String = ""
            Public Property allowedNetworkIds As New List(Of String)()
            Public Property productId As String = ""
            Public Property productIdSafe As String = ""
            Public Property hostId As String = ""
            Public Property userKey As String = ""
            Public Property billingEventId As String = ""
            Public Property hostEventId As String = ""
            Public Property firstUseUtc As String = ""
            Public Property lastUseUtc As System.String = ""
            Public Property lastUseIh As System.String = ""
            Public Property loadedDates As New System.Collections.Generic.List(Of System.String)()
            Public Property loadedDatesIh As System.String = ""
            Public Property reportIntervalMinutes As System.Int32 = DEFAULT_LICENSECOUNTERINTERVALL
            Public Property anon As Boolean
            Public Property actualNetworkId As String = ""
            Public Property actualNetworkKey As String = ""
            Public Property userObf As String = ""
            Public Property ih As String = ""
        End Class

        Private Class LicenseCounterKnownMarker
            Public Property file As String = ""
            Public Property firstSeenUtc As String = ""
            Public Property lastSeenUtc As String = ""
        End Class

        Private Class LicenseCounterDeliveryStatus
            Public Property lastAttemptUtc As String = ""
            Public Property lastSuccessUtc As String = ""
            Public Property lastStatus As String = ""
        End Class

        Private Class LicenseCounterErrorEvent
            Public Property v As Integer = LicenseCounterFormatVersion
            Public Property t As String = "E"
            Public Property errorEventId As String = ""
            Public Property errorCode As String = ""
            Public Property relatedMonth As String = ""
            Public Property relatedFile As String = ""
            Public Property detail As String = ""
            Public Property createdUtc As String = ""
        End Class

        Private Class LicenseCounterStateFile
            Public Property v As Integer = LicenseCounterFormatVersion
            Public Property knownHostEventIds As New Dictionary(Of String, LicenseCounterKnownMarker)(StringComparer.OrdinalIgnoreCase)
            Public Property delivery As New Dictionary(Of String, Dictionary(Of String, LicenseCounterDeliveryStatus))(StringComparer.OrdinalIgnoreCase)
            Public Property pendingErrors As New List(Of LicenseCounterErrorEvent)()
            Public Property appliedResubmitTokens As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            Public Property lastUpdatedUtc As String = ""
            Public Property ih As String = ""
        End Class

        Private Class LicenseCounterTransportEvent
            Public Property v As Integer = LicenseCounterFormatVersion
            Public Property t As String = "U"
            Public Property month As String = ""
            Public Property networkKey As String = ""
            Public Property networkScope As String = ""
            Public Property allowedNetworkIds As New List(Of String)()
            Public Property actualNetworkKey As String = ""
            Public Property actualNetworkId As String = ""
            Public Property productId As String = ""
            Public Property productIdSafe As String = ""
            Public Property hostId As String = ""
            Public Property userKey As String = ""
            Public Property billingEventId As String = ""
            Public Property hostEventId As String = ""
            Public Property firstUseUtc As String = ""
            Public Property lastUseUtc As System.String = ""
            Public Property reportedUtc As System.String = ""
            Public Property loadedDates As New System.Collections.Generic.List(Of System.String)()
            Public Property loadedHosts As New System.Collections.Generic.List(Of System.String)()
            Public Property reportIntervalMinutes As System.Int32 = DEFAULT_LICENSECOUNTERINTERVALL
            Public Property anon As Boolean
            Public Property userObf As String = ""
            Public Property errorEventId As String = ""
            Public Property errorCode As String = ""
            Public Property relatedMonth As String = ""
            Public Property relatedFile As String = ""
            Public Property detail As String = ""

            <JsonIgnore>
            Public Property Source As String = ""

            <JsonIgnore>
            Public Property GroupedDeliveryKey As System.String = ""

            <JsonIgnore>
            Public Property SourceHostEventIds As New System.Collections.Generic.List(Of System.String)()

            <JsonIgnore>
            Public Property SourceMarkerFiles As New System.Collections.Generic.List(Of System.String)()

            <JsonIgnore>
            Public ReadOnly Property DeliveryKey As String
                Get
                    If String.Equals(t, "E", StringComparison.OrdinalIgnoreCase) Then
                        Return errorEventId
                    End If

                    Return If(System.String.IsNullOrWhiteSpace(GroupedDeliveryKey), hostEventId, GroupedDeliveryKey)
                End Get
            End Property
        End Class

        Private Class LicenseCounterAnalysisEvent
            Public Property EventType As String = ""
            Public Property Month As String = ""
            Public Property NetworkKey As String = ""
            Public Property NetworkScope As String = ""
            Public Property AllowedNetworkIds As New List(Of String)()
            Public Property ProductId As String = ""
            Public Property ProductIdSafe As String = ""
            Public Property HostId As String = ""
            Public Property LoadedHosts As New System.Collections.Generic.List(Of System.String)()
            Public Property UserKey As String = ""
            Public Property BillingEventId As String = ""
            Public Property HostEventId As String = ""
            Public Property Anonymous As String = ""
            Public Property UserId As String = ""
            Public Property UserObf As String = ""
            Public Property Source As String = ""
            Public Property Validity As String = "valid"
            Public Property WarningText As String = ""
            Public Property ErrorCode As String = ""
        End Class

        Private Class LicenseCounterMainRow
            Public Property Month As String = ""
            Public Property NetworkKey As String = ""
            Public Property ProductId As String = ""
            Public Property DistinctUsers As Integer
            Public Property CoveredNetworks As String = ""
        End Class

        Private Class LicenseCounterNetworkScopeRow
            Public Property NetworkKey As String = ""
            Public Property CoveredNetworks As String = ""
        End Class

        Private Class LicenseCounterHostRow
            Public Property Month As String = ""
            Public Property NetworkKey As String = ""
            Public Property ProductId As String = ""
            Public Property HostId As String = ""
            Public Property DistinctUsers As Integer
        End Class

        Private Class LicenseCounterUserRow
            Public Property Month As String = ""
            Public Property NetworkKey As String = ""
            Public Property ProductId As String = ""
            Public Property UserId As String = ""
            Public Property UserKey As String = ""
            Public Property Hosts As String = ""
        End Class

        Private Class LicenseCounterAppliedResubmitTokenRow
            Public Property Token As String = ""
            Public Property AppliedUtc As String = ""
            Public Property Source As String = ""
        End Class

        Private Class LicenseCounterAnalysisResult
            Public Property Events As New List(Of LicenseCounterAnalysisEvent)()
            Public Property Warnings As New List(Of String)()
            Public Property MainRows As New List(Of LicenseCounterMainRow)()
            Public Property HostRows As New List(Of LicenseCounterHostRow)()
            Public Property UserRows As New List(Of LicenseCounterUserRow)()
            Public Property NetworkScopeRows As New List(Of LicenseCounterNetworkScopeRow)()
            Public Property AppliedResubmitTokens As New List(Of LicenseCounterAppliedResubmitTokenRow)()
        End Class

        Private Class LicenseCounterResubmitInstruction
            Public Property ResubmitAll As Boolean
            Public Property ResubmitMonth As String = ""
            Public Property ResubmitFromUtc As Nullable(Of System.DateTime)
            Public Property Token As String = ""
        End Class

        Public Shared Sub RegisterLicenseCounterUsageAndReport(context As ISharedContext, hostId As System.String)
            Try
                RegisterLicenseCounterUsageAndReportOptions(ReadLicenseCounterRuntimeOptions(context), hostId)
            Catch ex As System.Exception
                LogLicenseCounterInternal("register_error", ex.Message)
            End Try
        End Sub

        Private Shared Sub RegisterLicenseCounterUsageAndReportOptions(options As LicenseCounterRuntimeOptions, hostId As System.String)
            Try
                Dim runtime As LicenseCounterRuntimeContext = Nothing
                If Not TryCreateLicenseCounterRuntimeContext(options, hostId, runtime) Then Return

                ' This always runs, including while an older backlog is still being sent.
                ' It therefore records month boundaries independently of the network.
                EnsureLicenseCounterMarker(runtime)
                Dim capturedRuntime As LicenseCounterRuntimeContext = runtime

                If Not LicenseCounterReportBusy.TryAdd(runtime.HostId, 0) Then Return
                Try
                    System.Threading.Tasks.Task.Run(
                        Sub()
                            Try
                                ProcessLicenseCounterBackground(capturedRuntime)
                            Catch ex As System.Exception
                                LogLicenseCounterInternal("background_error", ex.Message)
                            Finally
                                Dim removed As System.Byte
                                LicenseCounterReportBusy.TryRemove(capturedRuntime.HostId, removed)
                            End Try
                        End Sub)
                Catch
                    Dim removed As System.Byte
                    LicenseCounterReportBusy.TryRemove(runtime.HostId, removed)
                    Throw
                End Try
            Catch ex As System.Exception
                LogLicenseCounterInternal("register_error", ex.Message)
            End Try
        End Sub

        Public Shared Sub ShowLicenseCounterEvaluatorDialog(context As ISharedContext)
            If Not SharedMethods.IsOnHostUiThread AndAlso Not SharedMethods.IsOnHostNeutralSta Then
                RunOnStaThread(Sub() ShowLicenseCounterEvaluatorDialog(context), "Red Ink license counter", alwaysUseWorker:=True)
                Return
            End If
            Try
                Dim analyzeFolderFirst As Boolean = (ShowCustomYesNoBox(
                    "Analyze a folder?" & vbCrLf & vbCrLf &
                    "Choose 'Folder' for a directory scan or 'Files' to select one or more JSON/log files.",
                    "Folder",
                    "Files",
                    $"{AN} - License Counter") = 1)

                Dim sources As New List(Of String)()
                Dim recursive As Boolean = False

                If analyzeFolderFirst Then
                    Using dlg As New FolderBrowserDialog()
                        dlg.Description = "Select the folder to analyze"
                        dlg.SelectedPath = GetLicenseCounterBaseFolder()

                        Dim __safeDialogOwner316 As System.Windows.Forms.IWin32Window = Global.SharedLibrary.SharedLibrary.SharedMethods.ResolveSameThreadDialogOwner()
                        If If(__safeDialogOwner316 IsNot Nothing, dlg.ShowDialog(__safeDialogOwner316), dlg.ShowDialog()) <> DialogResult.OK OrElse String.IsNullOrWhiteSpace(dlg.SelectedPath) Then
                            Return
                        End If

                        sources.Add(dlg.SelectedPath)
                    End Using

                    recursive = (ShowCustomYesNoBox(
                        "Scan subfolders recursively?",
                        "Yes",
                        "No",
                        $"{AN} - License Counter") = 1)
                Else
                    Using dlg As New OpenFileDialog()
                        dlg.Multiselect = True
                        dlg.Filter = "JSON / Log / Text|*.json;*.log;*.txt;*.csv;*.htm;*.html|All Files|*.*"

                        Dim __safeDialogOwner333 As System.Windows.Forms.IWin32Window = Global.SharedLibrary.SharedLibrary.SharedMethods.ResolveSameThreadDialogOwner()
                        If If(__safeDialogOwner333 IsNot Nothing, dlg.ShowDialog(__safeDialogOwner333), dlg.ShowDialog()) <> DialogResult.OK OrElse dlg.FileNames Is Nothing OrElse dlg.FileNames.Length = 0 Then
                            Return
                        End If

                        sources.AddRange(dlg.FileNames)
                    End Using
                End If

                Dim fromMonth As String = ShowCustomInputBox("From Month (yyyy-MM) or empty:", $"{AN} - License Counter", True)
                Dim toMonth As String = ShowCustomInputBox("To Month (yyyy-MM) or empty:", $"{AN} - License Counter", True)

                SharedLibrary.ProgressBarModule.CancelOperation = False
                SharedLibrary.ProgressBarModule.GlobalProgressMax = 100
                SharedLibrary.ProgressBarModule.GlobalProgressValue = 0
                SharedLibrary.ProgressBarModule.GlobalProgressLabel = "Initializing..."
                SharedLibrary.ProgressBarModule.ShowProgressBarInSeparateThread($"{AN} - License Counter", "Analyzing files...")

                Dim analysis = AnalyzeLicenseCounterSources(sources, recursive, fromMonth, toMonth)

                SharedLibrary.ProgressBarModule.GlobalProgressValue = SharedLibrary.ProgressBarModule.GlobalProgressMax
                SharedLibrary.ProgressBarModule.GlobalProgressLabel = "Analysis completed."
                SharedLibrary.ProgressBarModule.CancelOperation = True

                ShowLicenseCounterAnalysisResultDialog(analysis)

            Catch ex As System.Exception
                ShowCustomMessageBox($"License counter evaluation error: {ex.Message}", $"{AN} - License Counter")
            End Try
        End Sub

        Private Shared Function TryCreateLicenseCounterRuntimeContext(options As LicenseCounterRuntimeOptions,
                                                                      hostId As String,
                                                                      ByRef runtime As LicenseCounterRuntimeContext) As Boolean
            runtime = Nothing

            Try
                If options Is Nothing Then Return False
                If System.String.IsNullOrWhiteSpace(options.PathRaw) Then Return False
                ' Reporting is not a licensing gate. A configured LicenseCounterPath
                ' records the add-in load even for LicenseNoCheck, Private or trial use.
                ' Never activate, elevate or alter the licence state here.
                Dim hasStoredPro As System.Boolean = HasStoredProLicense()
                Dim licenseKey As System.String = If(hasStoredPro, My.Settings.License_Key, "")
                Dim storedProductId As System.String = If(hasStoredPro, My.Settings.License_ProductID, "")
                ' Existing stored Pro identities take precedence to preserve the
                ' established monthly key even when the optional INI value differs.
                Dim effectiveProductId As System.String = NormalizeLicenseCounterValue(storedProductId)
                If System.String.IsNullOrWhiteSpace(effectiveProductId) Then effectiveProductId = NormalizeLicenseCounterValue(options.ConfiguredProductId)
                If System.String.IsNullOrWhiteSpace(effectiveProductId) Then
                    effectiveProductId = "UNASSIGNED"
                    LogLicenseCounterDiagnostic("unassigned_product", "host=" & NormalizeLicenseCounterHostId(hostId) & "; configure LicenseProductID for billing attribution")
                End If

                Dim actualNetworkId As System.String = NormalizeLicenseCounterValue(System.Environment.GetEnvironmentVariable("USERDNSDOMAIN")).ToLowerInvariant()
                If System.String.IsNullOrWhiteSpace(actualNetworkId) Then actualNetworkId = NormalizeLicenseCounterValue(System.Environment.UserDomainName).ToLowerInvariant()
                If System.String.IsNullOrWhiteSpace(actualNetworkId) Then actualNetworkId = "local-" & NormalizeLicenseCounterValue(System.Environment.MachineName).ToLowerInvariant()

                Dim normalizedAllowedNetworkIds As System.Collections.Generic.List(Of System.String) = Nothing
                Dim signedOfflineIdentity As System.Boolean = False
                If hasStoredPro AndAlso Not System.String.IsNullOrWhiteSpace(licenseKey) AndAlso IsOfflineDomainLicenseKey(licenseKey) Then
                    Dim info As OfflineDomainLicenseInfo = Nothing
                    Dim reason As System.String = ""
                    If TryParseAndVerifyOfflineDomainLicense(storedProductId, licenseKey, info, reason) Then
                        ' Preserve the v1 network/product identity of a verified signed licence.
                        effectiveProductId = info.ProductId
                        actualNetworkId = NormalizeLicenseCounterValue(info.MatchedDomain)
                        normalizedAllowedNetworkIds = NormalizeAllowedNetworkIds(info.AllowedDomains)
                        signedOfflineIdentity = True
                    Else
                        ' Invalid evidence is never treated as a verified licence. Counting
                        ' still proceeds using the configured reporting identity below.
                        LogLicenseCounterDiagnostic("unverified_offline_identity", "host=" & NormalizeLicenseCounterHostId(hostId) & "; using reporting scope")
                    End If
                End If

                If Not signedOfflineIdentity Then
                    ' A single NetworkScope provides stable grouping for classic/private/
                    ' unchecked deployments across multiple Office hosts and machines.
                    Dim scope As System.String = NormalizeLicenseCounterValue(options.ClassicNetworkScope).ToLowerInvariant()
                    If System.String.IsNullOrWhiteSpace(scope) Then scope = actualNetworkId
                    normalizedAllowedNetworkIds = SplitAllowedNetworkIds(scope)
                    If normalizedAllowedNetworkIds.Count = 0 Then normalizedAllowedNetworkIds.Add("UNKNOWN_NETWORK")
                End If

                ' Prefer the established Pro user identity when present. With private or
                ' unchecked installations, honour the explicitly configured LicenseUserID.
                Dim effectiveUserId As System.String = ""
                If hasStoredPro Then effectiveUserId = My.Settings.License_UserID
                If System.String.IsNullOrWhiteSpace(effectiveUserId) Then effectiveUserId = ExpandLicenseEnvironmentVariables(options.ConfiguredUserId)
                Dim fallbackUsed As System.Boolean = False
                If System.String.IsNullOrWhiteSpace(effectiveUserId) Then
                    effectiveUserId = ExpandLicenseEnvironmentVariables("%USERNAME%")
                    fallbackUsed = True
                End If

                Dim normalizedUserId As String = NormalizeLicenseCounterUserId(effectiveUserId)
                Dim normalizedHostId As String = NormalizeLicenseCounterHostId(hostId)
                Dim jointNetworkId As String = BuildJointNetworkId(normalizedAllowedNetworkIds)
                Dim observedLocalTime As System.DateTime = System.DateTime.Now
                Dim month As String = observedLocalTime.ToString("yyyy-MM", CultureInfo.InvariantCulture)
                Dim resubmitInstruction As LicenseCounterResubmitInstruction = ParseLicenseCounterResubmitInstruction(options.MethodRaw)

                If String.IsNullOrWhiteSpace(normalizedUserId) Then
                    normalizedUserId = "UNKNOWN_USER"
                End If

                Dim jointNetworkKey As String = BuildJointNetworkKey(jointNetworkId)
                Dim actualNetworkKey As String = BuildNetworkKey(actualNetworkId)
                Dim productIdSafe As String = MakeLicenseCounterSafeProductId(effectiveProductId)
                Dim userKey As String = BuildUserKey(jointNetworkKey, normalizedUserId)
                Dim billingEventId As String = BuildBillingEventId(month, jointNetworkKey, effectiveProductId, userKey)
                Dim hostEventId As String = BuildHostEventId(month, jointNetworkKey, effectiveProductId, normalizedHostId, userKey)
                Dim anon As Boolean = options.Anon
                Dim userObf As String = If(anon, "", BuildUserObf(normalizedUserId))
                Dim firstUseUtc As String = observedLocalTime.ToUniversalTime().ToString("o", CultureInfo.InvariantCulture)
                Dim markerFileName As String = $"{LicenseCounterMarkerPrefix}{month}_{jointNetworkKey}_{productIdSafe}_{normalizedHostId}_{userKey}.json"

                Dim markerRecord As New LicenseCounterMarkerRecord() With {
                    .v = LicenseCounterFormatVersion,
                    .t = "U",
                    .month = month,
                    .networkKey = jointNetworkKey,
                    .networkScope = jointNetworkId,
                    .allowedNetworkIds = New List(Of String)(normalizedAllowedNetworkIds),
                    .productId = effectiveProductId,
                    .productIdSafe = productIdSafe,
                    .hostId = normalizedHostId,
                    .userKey = userKey,
                    .billingEventId = billingEventId,
                    .hostEventId = hostEventId,
                    .firstUseUtc = firstUseUtc,
                    .anon = anon,
                    .actualNetworkId = actualNetworkId,
                    .actualNetworkKey = actualNetworkKey,
                    .userObf = userObf
                }

                markerRecord.ih = ComputeMarkerIntegrityHash(markerRecord)

                Dim parsedMethod As LicenseCounterMethodKind = ParseLicenseCounterMethod(options.MethodRaw)
                Dim parsedWebMode As LicenseCounterWebModeKind = ParseLicenseCounterWebMode(options.MethodRaw, parsedMethod)
                Dim parsedTargets As List(Of LicenseCounterTarget) = ParseLicenseCounterTargets(
                    options.PathRaw,
                    parsedMethod)

                runtime = New LicenseCounterRuntimeContext() With {
                    .BaseFolder = GetLicenseCounterBaseFolder(),
                    .PathRaw = options.PathRaw,
                    .MethodRaw = options.MethodRaw,
                    .MethodKind = parsedMethod,
                    .ReportIntervalMinutes = options.ReportIntervalMinutes,
                    .RequireAck = options.RequireAck,
                    .WebMode = parsedWebMode,
                    .Targets = parsedTargets,
                    .Month = month,
                    .ObservedLocalDate = observedLocalTime.ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture),
                    .JointNetworkId = jointNetworkId,
                    .JointNetworkKey = jointNetworkKey,
                    .NetworkScope = jointNetworkId,
                    .AllowedNetworkIds = New List(Of String)(normalizedAllowedNetworkIds),
                    .ActualNetworkId = actualNetworkId,
                    .ActualNetworkKey = actualNetworkKey,
                    .ProductId = effectiveProductId,
                    .ProductIdSafe = productIdSafe,
                    .HostId = normalizedHostId,
                    .UserId = normalizedUserId,
                    .UserKey = userKey,
                    .BillingEventId = billingEventId,
                    .HostEventId = hostEventId,
                    .FirstUseUtc = firstUseUtc,
                    .Anon = anon,
                    .UserObf = userObf,
                    .EmptyUserIdFallbackUsed = fallbackUsed,
                    .MarkerFileName = markerFileName,
                    .MarkerFilePath = Path.Combine(GetLicenseCounterBaseFolder(), markerFileName),
                    .ResubmitInstruction = resubmitInstruction,
                    .MarkerRecord = markerRecord
                }

                LogLicenseCounterInternal("runtime_context",
                                          $"Method={runtime.MethodKind}, WebMode={runtime.WebMode}, Targets={runtime.Targets.Count}, Month={runtime.Month}, HostEventId={runtime.HostEventId}, BillingEventId={runtime.BillingEventId}")
                If runtime.Targets.Count = 0 AndAlso runtime.MethodKind <> LicenseCounterMethodKind.LocalOnly Then
                    LogLicenseCounterDiagnostic("skipped", "No valid reporting targets; host=" & runtime.HostId)
                End If

                Return True

            Catch ex As System.Exception
                LogLicenseCounterDiagnostic("context_error", ex.GetType().Name)
                LogLicenseCounterInternal("context_error", ex.Message)
                Return False
            End Try
        End Function

        Private Shared Sub EnsureLicenseCounterMarker(runtime As LicenseCounterRuntimeContext)
            Try
                System.IO.Directory.CreateDirectory(runtime.BaseFolder)
                WithLicenseCounterStateLock(
                    runtime.BaseFolder,
                    Sub(state)
                        Dim record As LicenseCounterMarkerRecord = runtime.MarkerRecord
                        Dim markerExists As System.Boolean = System.IO.File.Exists(runtime.MarkerFilePath)
                        If markerExists Then
                            record = Newtonsoft.Json.JsonConvert.DeserializeObject(Of LicenseCounterMarkerRecord)(System.IO.File.ReadAllText(runtime.MarkerFilePath, System.Text.Encoding.UTF8))
                            If record Is Nothing OrElse Not System.String.Equals(record.ih, ComputeMarkerIntegrityHash(record), System.StringComparison.OrdinalIgnoreCase) OrElse
                               Not System.String.Equals(record.hostEventId, runtime.HostEventId, System.StringComparison.OrdinalIgnoreCase) Then
                                Throw New System.IO.InvalidDataException("Existing licence-counter marker failed integrity validation; it was preserved for investigation.")
                            End If
                        End If
                        ' Read optional metadata before an anonymity change alters v1 ih.
                        ' Old markers contribute monthly evidence, never invented dates.
                        Dim observedDates As System.Collections.Generic.List(Of System.String) = ReadLicenseCounterLoadedDates(record)
                        Dim previousLastUse As System.String = ReadLicenseCounterLastUseUtc(record)
                        Dim newDay As System.Boolean = Not observedDates.Contains(runtime.ObservedLocalDate)
                        Dim removeIdentity As System.Boolean = runtime.Anon AndAlso (Not record.anon OrElse Not System.String.IsNullOrWhiteSpace(record.userObf))
                        Dim changed As System.Boolean = Not markerExists OrElse newDay OrElse removeIdentity OrElse record.reportIntervalMinutes <> runtime.ReportIntervalMinutes
                        If newDay Then observedDates.Add(runtime.ObservedLocalDate)
                        If runtime.Anon Then
                            record.anon = True
                            record.userObf = ""
                        End If
                        record.loadedDates = observedDates.OrderBy(Function(value) value, System.StringComparer.Ordinal).ToList()
                        ' This timestamp now means the first load observation of the latest
                        ' observed day. Same-day checks neither move it nor rewrite the marker.
                        record.lastUseUtc = If(newDay, runtime.FirstUseUtc, previousLastUse)
                        record.reportIntervalMinutes = runtime.ReportIntervalMinutes
                        record.ih = ComputeMarkerIntegrityHash(record)
                        record.lastUseIh = ComputeLicenseCounterLastUseIntegrityHash(record)
                        record.loadedDatesIh = ComputeLicenseCounterLoadedDatesIntegrityHash(record)
                        If changed Then
                            WriteLicenseCounterJsonAtomically(runtime.MarkerFilePath, Newtonsoft.Json.JsonConvert.SerializeObject(record, Newtonsoft.Json.Formatting.Indented))
                            LogLicenseCounterInternal("marker_updated", "HostEventId=" & record.hostEventId & ", Month=" & record.month & ", LoadedDate=" & runtime.ObservedLocalDate)
                            LogLicenseCounterDiagnostic("marker_updated", "host=" & runtime.HostId & "; month=" & record.month & "; loadedDate=" & runtime.ObservedLocalDate)
                        End If
                        runtime.MarkerRecord = record
                        runtime.FirstUseUtc = record.firstUseUtc
                        EnsureKnownMarkerEntry(state, record.hostEventId, runtime.MarkerFileName)
                    End Sub)
            Catch ex As System.Exception
                LogLicenseCounterDiagnostic("marker_error", ex.GetType().Name)
                LogLicenseCounterInternal("marker_error", ex.Message)
            End Try
        End Sub

        Private Shared Function ComputeLicenseCounterLoadedDatesIntegrityHash(record As LicenseCounterMarkerRecord) As System.String
            Return HashSha256("RI-LC-LOADED-DATES-v1|" & record.ih & "|" & System.String.Join("|", If(record.loadedDates, New System.Collections.Generic.List(Of System.String)())) & "|" & LicenseCounterLicenseFingerprint)
        End Function

        Private Shared Function ReadLicenseCounterLoadedDates(record As LicenseCounterMarkerRecord) As System.Collections.Generic.List(Of System.String)
            Dim empty As New System.Collections.Generic.List(Of System.String)()
            If record.loadedDates Is Nothing OrElse record.loadedDates.Count = 0 Then Return empty
            If record.loadedDates.Count <= 31 AndAlso System.String.Equals(record.loadedDatesIh, ComputeLicenseCounterLoadedDatesIntegrityHash(record), System.StringComparison.OrdinalIgnoreCase) Then
                For Each value As System.String In record.loadedDates
                    Dim parsed As System.DateTime
                    If Not System.DateTime.TryParseExact(value, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, parsed) OrElse
                       Not System.String.Equals(parsed.ToString("yyyy-MM", System.Globalization.CultureInfo.InvariantCulture), record.month, System.StringComparison.Ordinal) Then
                        LogLicenseCounterInternal("marker_loaded_dates_invalid", "HostEventId=" & record.hostEventId & "; monthly evidence retained without invalid daily metadata.")
                        Return empty
                    End If
                Next
                Return record.loadedDates.Distinct(System.StringComparer.Ordinal).OrderBy(Function(value) value, System.StringComparer.Ordinal).ToList()
            End If
            LogLicenseCounterInternal("marker_loaded_dates_integrity_error", "HostEventId=" & record.hostEventId & "; monthly evidence retained without unverified daily metadata.")
            Return empty
        End Function

        Private Shared Function ComputeLicenseCounterLastUseIntegrityHash(record As LicenseCounterMarkerRecord) As System.String
            Return HashSha256("RI-LC-LASTUSE-v1|" & record.ih & "|" & record.lastUseUtc & "|" & LicenseCounterLicenseFingerprint)
        End Function

        Private Shared Function ReadLicenseCounterLastUseUtc(record As LicenseCounterMarkerRecord) As System.String
            If System.String.IsNullOrWhiteSpace(record.lastUseUtc) Then Return ""
            If System.String.Equals(record.lastUseIh, ComputeLicenseCounterLastUseIntegrityHash(record), System.StringComparison.OrdinalIgnoreCase) Then Return record.lastUseUtc
            LogLicenseCounterInternal("marker_freshness_integrity_error", "HostEventId=" & record.hostEventId & "; monthly evidence retained without freshness metadata.")
            Return ""
        End Function

        Private Shared Sub WriteLicenseCounterJsonAtomically(path As System.String, json As System.String)
            Dim temporaryPath As System.String = path & "." & System.Guid.NewGuid().ToString("N") & ".tmp"
            Try
                Using stream As New System.IO.FileStream(temporaryPath, System.IO.FileMode.CreateNew, System.IO.FileAccess.Write, System.IO.FileShare.None)
                    Dim preamble As System.Byte() = New System.Text.UTF8Encoding(True).GetPreamble()
                    Dim bytes As System.Byte() = System.Text.Encoding.UTF8.GetBytes(json)
                    stream.Write(preamble, 0, preamble.Length)
                    stream.Write(bytes, 0, bytes.Length)
                    stream.Flush(True)
                End Using
                If System.IO.File.Exists(path) Then
                    System.IO.File.Replace(temporaryPath, path, Nothing)
                Else
                    System.IO.File.Move(temporaryPath, path)
                End If
            Finally
                If System.IO.File.Exists(temporaryPath) Then System.IO.File.Delete(temporaryPath)
            End Try
        End Sub

        Private Shared Sub ProcessLicenseCounterBackground(runtime As LicenseCounterRuntimeContext)
            Dim mutexName As System.String = "Local\RI_LicenseCounter_Report_" & FirstHex(HashSha256(System.Environment.GetFolderPath(System.Environment.SpecialFolder.UserProfile)), 16)
            Dim hasHandle As System.Boolean = False
            Using reportMutex As New System.Threading.Mutex(False, mutexName)
                Try
                    Try
                        ' Wait on a background worker rather than dropping a concurrent
                        ' Excel/Outlook start while Word is delivering its first marker.
                        hasHandle = reportMutex.WaitOne(System.TimeSpan.FromMinutes(2))
                    Catch ex As System.Threading.AbandonedMutexException
                        hasHandle = True
                    End Try
                    If Not hasHandle Then
                        LogLicenseCounterDiagnostic("delivery_lock_timeout", "host=" & runtime.HostId & "; marker retained for later retry")
                        Return
                    End If
                    ' Hold this separate lock through the ACK save. Network I/O never
                    ' holds the local state mutex; Office UI never waits on this transport lock.
                    ProcessLicenseCounterBackgroundLocked(runtime)
                Finally
                    If hasHandle Then reportMutex.ReleaseMutex()
                End Try
            End Using
        End Sub

        Private Shared Sub ProcessLicenseCounterBackgroundLocked(runtime As LicenseCounterRuntimeContext)
            Dim pendingEvents As New List(Of LicenseCounterTransportEvent)()
            Dim pendingTargets As New System.Collections.Generic.Dictionary(Of System.String, System.Collections.Generic.HashSet(Of System.String))(System.StringComparer.OrdinalIgnoreCase)

            WithLicenseCounterStateLock(
                runtime.BaseFolder,
                Sub(state)
                    RefreshKnownMarkersAndDetectMissing(runtime.BaseFolder, state)
                    EnsureKnownMarkerEntry(state, runtime.MarkerRecord.hostEventId, runtime.MarkerFileName)
                    ApplyLicenseCounterResubmissionInstruction(runtime.BaseFolder, state, runtime.ResubmitInstruction)

                    If runtime.EmptyUserIdFallbackUsed Then
                        QueueLicenseCounterError(state,
                                                 "empty_userid",
                                                 runtime.Month,
                                                 runtime.MarkerFileName,
                                                 "LicenseUserID was empty; fallback %USERNAME% was used.")
                    End If

                    Dim enabledTargets As System.Collections.Generic.List(Of LicenseCounterTarget) = runtime.Targets.Where(Function(target) ShouldUseLicenseCounterTarget(runtime.MethodKind, target)).ToList()
                    ApplyLicenseCounterRetention(runtime.BaseFolder, state, enabledTargets, runtime.Anon)
                    pendingEvents.AddRange(CollectLicenseCounterPendingEvents(runtime.BaseFolder, state, enabledTargets, runtime.Anon))
                    For Each evt As LicenseCounterTransportEvent In pendingEvents
                        Dim targetHashes As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
                        For Each target As LicenseCounterTarget In enabledTargets
                            If ShouldAttemptLicenseCounterEvent(state, evt, target.TargetHash) Then targetHashes.Add(target.TargetHash)
                        Next
                        pendingTargets(evt.DeliveryKey) = targetHashes
                    Next
                End Sub)

            If runtime.MethodKind = LicenseCounterMethodKind.LocalOnly Then
                Return
            End If

            If pendingEvents.Count = 0 Then
                Return
            End If

            Dim deliveryResults As New List(Of Tuple(Of String, String, Boolean, String))()

            For Each evt In pendingEvents
                For Each target In runtime.Targets
                    If Not pendingTargets.ContainsKey(evt.DeliveryKey) OrElse Not pendingTargets(evt.DeliveryKey).Contains(target.TargetHash) Then Continue For
                    If Not ShouldUseLicenseCounterTarget(runtime.MethodKind, target) Then
                        LogLicenseCounterInternal("submit_skipped",
                                                  $"Event={evt.DeliveryKey}, TargetHash={target.TargetHash}, TargetType={target.TargetType}, Method={runtime.MethodKind}")
                        Continue For
                    End If

                    Dim success As Boolean = False
                    Dim statusText As String = ""

                    LogLicenseCounterInternal("submit_attempt",
                                              $"Event={evt.DeliveryKey}, Type={evt.t}, TargetHash={target.TargetHash}, TargetType={target.TargetType}")

                    Try
                        evt.reportedUtc = System.DateTime.UtcNow.ToString("o", System.Globalization.CultureInfo.InvariantCulture)
                        Select Case target.TargetType
                            Case LicenseCounterTargetType.Web
                                success = SendLicenseCounterWebEvent(target, evt, runtime.MethodKind, runtime.WebMode, statusText, runtime.RequireAck)

                            Case LicenseCounterTargetType.Directory
                                success = WriteLicenseCounterReportFile(target, evt, statusText)

                            Case LicenseCounterTargetType.AppendFile
                                success = AppendLicenseCounterReportLine(target, evt, statusText)

                            Case Else
                                statusText = "invalid_target"
                        End Select
                    Catch ex As System.Exception
                        success = False
                        statusText = ex.Message
                    End Try

                    LogLicenseCounterInternal("submit_result",
                                              $"Event={evt.DeliveryKey}, TargetHash={target.TargetHash}, Success={success}, Status={statusText}")
                    ' Preserve only an expected HTTP status or a fixed networking error
                    ' label. Never persist raw exception text, which can contain the
                    ' credential-bearing destination URL.
                    Dim safeStatus As System.String = If(success, "ok", "failed")
                    If target.TargetType = LicenseCounterTargetType.Web Then
                        If System.Text.RegularExpressions.Regex.IsMatch(If(statusText, ""), "^[1-5][0-9]{2}(:[a-z_]+)?$") Then
                            safeStatus = statusText
                        Else
                            Select Case statusText
                                Case "ConnectFailure", "Timeout", "NameResolutionFailure", "ProxyNameResolutionFailure", "TrustFailure", "SendFailure", "ReceiveFailure", "ProtocolError"
                                    safeStatus = statusText
                            End Select
                        End If
                    End If
                    LogLicenseCounterDiagnostic("delivery", "host=" & runtime.HostId & "; targetType=" & target.TargetType.ToString() & "; success=" & success.ToString() & "; status=" & safeStatus)

                    deliveryResults.Add(Tuple.Create(evt.DeliveryKey, target.TargetHash, success, statusText))
                    If success AndAlso System.String.Equals(evt.t, "U", System.StringComparison.OrdinalIgnoreCase) Then
                        ' Preserve original host delivery entries for v1 readers, resubmit
                        ' and legacy retention. The group ACK separately proves the dates.
                        For Each sourceHostEventId As System.String In evt.SourceHostEventIds
                            deliveryResults.Add(System.Tuple.Create(sourceHostEventId, target.TargetHash, True, statusText))
                        Next
                    End If
                Next
            Next

            WithLicenseCounterStateLock(
                runtime.BaseFolder,
                Sub(state)
                    For Each result In deliveryResults
                        UpdateLicenseCounterDelivery(state, result.Item1, result.Item2, result.Item3, result.Item4)
                    Next

                    Dim deliveredErrorIds As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

                    For Each evt In pendingEvents
                        If String.Equals(evt.t, "E", StringComparison.OrdinalIgnoreCase) Then
                            Dim allSucceeded As Boolean = True

                            For Each target In runtime.Targets
                                If Not ShouldUseLicenseCounterTarget(runtime.MethodKind, target) Then
                                    Continue For
                                End If

                                Dim deliveryStatus = GetLicenseCounterDeliveryStatus(state, evt.DeliveryKey, target.TargetHash)

                                If deliveryStatus Is Nothing OrElse String.IsNullOrWhiteSpace(deliveryStatus.lastSuccessUtc) Then
                                    allSucceeded = False
                                    Exit For
                                End If
                            Next

                            If allSucceeded Then
                                deliveredErrorIds.Add(evt.errorEventId)
                            End If
                        End If
                    Next

                    If deliveredErrorIds.Count > 0 Then
                        state.pendingErrors = state.pendingErrors.
                            Where(Function(p) Not deliveredErrorIds.Contains(p.errorEventId)).
                            ToList()
                    End If
                End Sub)
        End Sub

        Private Shared Function CollectLicenseCounterPendingEvents(baseFolder As System.String,
                                                                   state As LicenseCounterStateFile,
                                                                   targets As System.Collections.Generic.List(Of LicenseCounterTarget),
                                                                   Optional forceAnonymous As System.Boolean = False) As System.Collections.Generic.List(Of LicenseCounterTransportEvent)
            Dim results As New System.Collections.Generic.List(Of LicenseCounterTransportEvent)()
            Try
                For Each evt As LicenseCounterTransportEvent In CollectLicenseCounterUsageGroups(baseFolder, forceAnonymous)
                    If targets.Any(Function(target) ShouldAttemptLicenseCounterEvent(state, evt, target.TargetHash)) Then
                        LogLicenseCounterInternal("pending_usage_event", "Event=" & evt.DeliveryKey & ", Month=" & evt.month & ", LoadedDays=" & evt.loadedDates.Count.ToString(System.Globalization.CultureInfo.InvariantCulture))
                        results.Add(evt)
                    End If
                Next
                For Each errEvt As LicenseCounterErrorEvent In state.pendingErrors
                    Dim evt As New LicenseCounterTransportEvent() With {
                        .v = errEvt.v,
                        .t = "E",
                        .errorEventId = errEvt.errorEventId,
                        .errorCode = errEvt.errorCode,
                        .relatedMonth = errEvt.relatedMonth,
                        .relatedFile = errEvt.relatedFile,
                        .detail = errEvt.detail,
                        .Source = "state"
                    }
                    If targets.Any(Function(target) ShouldAttemptLicenseCounterEvent(state, evt, target.TargetHash)) Then results.Add(evt)
                Next
            Catch ex As System.Exception
                LogLicenseCounterInternal("collect_pending_error", ex.Message)
            End Try
            Return results
        End Function

        Private Shared Function CollectLicenseCounterUsageGroups(baseFolder As System.String,
                                                                 Optional forceAnonymous As System.Boolean = False) As System.Collections.Generic.List(Of LicenseCounterTransportEvent)
            Dim markers As New System.Collections.Generic.List(Of LicenseCounterTransportEvent)()
            For Each filePath As System.String In System.IO.Directory.GetFiles(baseFolder, LicenseCounterMarkerPrefix & "*.json", System.IO.SearchOption.TopDirectoryOnly)
                Dim evt As LicenseCounterTransportEvent = TryReadLicenseCounterUsageEvent(filePath)
                If evt IsNot Nothing Then markers.Add(evt)
            Next
            Dim results As New System.Collections.Generic.List(Of LicenseCounterTransportEvent)()
            For Each group In markers.GroupBy(Function(evt) evt.billingEventId, System.StringComparer.OrdinalIgnoreCase)
                ' Keep a native host/e identity for older readers; billing identity is
                ' shared by the aggregate and all its actual local host markers.
                Dim aggregate As LicenseCounterTransportEvent = group.OrderBy(Function(evt) evt.firstUseUtc, System.StringComparer.Ordinal).ThenBy(Function(evt) evt.hostId, System.StringComparer.Ordinal).First()
                aggregate.SourceHostEventIds = group.Select(Function(evt) evt.hostEventId).Distinct(System.StringComparer.OrdinalIgnoreCase).ToList()
                aggregate.SourceMarkerFiles = group.Select(Function(evt) evt.Source).ToList()
                aggregate.loadedDates = group.SelectMany(Function(evt) evt.loadedDates).Distinct(System.StringComparer.Ordinal).OrderBy(Function(value) value, System.StringComparer.Ordinal).ToList()
                aggregate.loadedHosts = group.Select(Function(evt) evt.hostId).Distinct(System.StringComparer.OrdinalIgnoreCase).OrderBy(Function(value) value, System.StringComparer.Ordinal).ToList()
                aggregate.lastUseUtc = If(group.Select(Function(evt) evt.lastUseUtc).Where(Function(value) Not System.String.IsNullOrWhiteSpace(value)).OrderByDescending(Function(value) value, System.StringComparer.Ordinal).FirstOrDefault(), "")
                aggregate.anon = forceAnonymous OrElse group.Any(Function(evt) evt.anon)
                If aggregate.anon Then aggregate.userObf = ""
                aggregate.GroupedDeliveryKey = BuildLicenseCounterGroupDeliveryKey(aggregate)
                results.Add(aggregate)
            Next
            Return results
        End Function

        Private Shared Function BuildLicenseCounterGroupDeliveryKey(evt As LicenseCounterTransportEvent) As System.String
            ' Change the ACK identity when dates OR host coverage changes. A Word
            ' acknowledgement cannot suppress a later Excel/Outlook host observation.
            Dim days As System.String = System.String.Join("|", evt.loadedDates)
            Dim hosts As System.String = System.String.Join("|", evt.loadedHosts)
            Return "d" & evt.billingEventId & "_" & FirstHex(HashSha256("RI-LC-EVIDENCE-DELIVERY-v2|" & days & "|" & hosts & "|" & evt.anon.ToString(System.Globalization.CultureInfo.InvariantCulture)), 32)
        End Function

        Private Shared Function ShouldAttemptLicenseCounterEvent(state As LicenseCounterStateFile,
                                                                 evt As LicenseCounterTransportEvent,
                                                                 targetHash As System.String) As System.Boolean
            If Not ShouldAttemptLicenseCounterDelivery(GetLicenseCounterDeliveryStatus(state, evt.DeliveryKey, targetHash)) Then Return False
            If System.String.Equals(evt.t, "U", System.StringComparison.OrdinalIgnoreCase) AndAlso evt.loadedDates.Count > 0 AndAlso evt.loadedHosts.Count = 1 Then
                ' Existing v1.1 day-only ACKs proved coverage for a single host. Preserve
                ' them on upgrade rather than replaying every historic single-host marker.
                Dim legacyDays As System.String = System.String.Join("|", evt.loadedDates)
                Dim legacyKey As System.String = "d" & evt.billingEventId & "_" & FirstHex(HashSha256("RI-LC-DAY-DELIVERY-v1|" & legacyDays & "|" & evt.anon.ToString(System.Globalization.CultureInfo.InvariantCulture)), 32)
                If Not ShouldAttemptLicenseCounterDelivery(GetLicenseCounterDeliveryStatus(state, legacyKey, targetHash)) Then Return False
            End If
            If System.String.Equals(evt.t, "U", System.StringComparison.OrdinalIgnoreCase) AndAlso evt.loadedDates.Count = 0 AndAlso evt.SourceHostEventIds.Count > 0 Then
                ' Upgrade does not replay old, already acknowledged month-only markers.
                Return evt.SourceHostEventIds.Any(Function(hostEventId) ShouldAttemptLicenseCounterDelivery(GetLicenseCounterDeliveryStatus(state, hostEventId, targetHash)))
            End If
            Return True
        End Function

        Private Shared Function ShouldAttemptLicenseCounterDelivery(status As LicenseCounterDeliveryStatus) As System.Boolean
            ' Interval/startup checks retry unacknowledged evidence only. A successful
            ' target remains suppressed until the aggregate gains another observed date.
            Return status Is Nothing OrElse System.String.IsNullOrWhiteSpace(status.lastSuccessUtc)
        End Function

        Private Shared Function ShouldUseLicenseCounterTarget(methodKind As LicenseCounterMethodKind,
                                                              target As LicenseCounterTarget) As Boolean
            Select Case methodKind
                Case LicenseCounterMethodKind.LocalOnly
                    Return False

                Case LicenseCounterMethodKind.FileOnly
                    Return target.TargetType = LicenseCounterTargetType.Directory

                Case LicenseCounterMethodKind.WebBasic, LicenseCounterMethodKind.WebExtended, LicenseCounterMethodKind.WebPathOnly
                    Return target.TargetType = LicenseCounterTargetType.Web

                Case LicenseCounterMethodKind.AppendFile
                    Return target.TargetType = LicenseCounterTargetType.AppendFile

                Case Else
                    Return target.TargetType <> LicenseCounterTargetType.Invalid
            End Select
        End Function

        Private Shared Function TryReadLicenseCounterUsageEvent(filePath As String) As LicenseCounterTransportEvent
            Try
                Dim json As String = File.ReadAllText(filePath, Encoding.UTF8)
                Dim record = JsonConvert.DeserializeObject(Of LicenseCounterMarkerRecord)(json)

                If record Is Nothing OrElse Not String.Equals(record.t, "U", StringComparison.OrdinalIgnoreCase) Then
                    Return Nothing
                End If

                If Not String.Equals(record.ih, ComputeMarkerIntegrityHash(record), StringComparison.OrdinalIgnoreCase) Then
                    Return Nothing
                End If

                Return New LicenseCounterTransportEvent() With {
                    .v = record.v,
                    .t = "U",
                    .month = record.month,
                    .networkKey = record.networkKey,
                    .networkScope = record.networkScope,
                    .allowedNetworkIds = If(record.allowedNetworkIds, New List(Of String)()),
                    .actualNetworkKey = record.actualNetworkKey,
                    .actualNetworkId = record.actualNetworkId,
                    .productId = record.productId,
                    .productIdSafe = record.productIdSafe,
                    .hostId = record.hostId,
                    .userKey = record.userKey,
                    .billingEventId = record.billingEventId,
                    .hostEventId = record.hostEventId,
                    .firstUseUtc = record.firstUseUtc,
                    .lastUseUtc = ReadLicenseCounterLastUseUtc(record),
                    .loadedDates = ReadLicenseCounterLoadedDates(record),
                    .loadedHosts = New System.Collections.Generic.List(Of System.String) From {record.hostId},
                    .reportIntervalMinutes = record.reportIntervalMinutes,
                    .anon = record.anon,
                    .userObf = If(record.anon, "", record.userObf),
                    .Source = filePath
                }

            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Function WriteLicenseCounterReportFile(target As LicenseCounterTarget,
                                                              evt As LicenseCounterTransportEvent,
                                                              ByRef statusText As String) As Boolean
            statusText = ""

            Dim dirPath As String = target.NormalizedValue

            If Not Directory.Exists(dirPath) Then
                Directory.CreateDirectory(dirPath)
            End If

            Dim fileName As String

            If String.Equals(evt.t, "E", StringComparison.OrdinalIgnoreCase) Then
                fileName = $"{LicenseCounterReportPrefix}ERROR_{evt.errorEventId}.json"
            Else
                fileName = $"{LicenseCounterReportPrefix}{evt.month}_{evt.networkKey}_{evt.productIdSafe}_{evt.hostId}_{evt.userKey}_{evt.hostEventId}.json"
            End If

            Dim fullPath As String = Path.Combine(dirPath, fileName)

            For attempt As System.Int32 = 1 To 5
                Try
                    Using reportLock As New System.IO.FileStream(fullPath & ".lock", System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.ReadWrite, System.IO.FileShare.None)
                        If System.IO.File.Exists(fullPath) Then
                            If Not System.String.Equals(evt.t, "U", System.StringComparison.OrdinalIgnoreCase) Then
                                statusText = "exists"
                                Return True
                            End If
                            Dim previous As LicenseCounterTransportEvent = Newtonsoft.Json.JsonConvert.DeserializeObject(Of LicenseCounterTransportEvent)(System.IO.File.ReadAllText(fullPath, System.Text.Encoding.UTF8))
                            If previous Is Nothing OrElse Not System.String.Equals(previous.hostEventId, evt.hostEventId, System.StringComparison.OrdinalIgnoreCase) Then
                                statusText = "existing_report_identity_mismatch"
                                Return False
                            End If
                            Dim previousDates As System.Collections.Generic.List(Of System.String) = If(previous.loadedDates, New System.Collections.Generic.List(Of System.String)())
                            Dim previousHosts As System.Collections.Generic.List(Of System.String) = If(previous.loadedHosts, New System.Collections.Generic.List(Of System.String)())
                            Dim addedEvidence As System.Boolean = evt.loadedDates.Any(Function(value) Not previousDates.Contains(value)) OrElse evt.loadedHosts.Any(Function(value) Not previousHosts.Contains(value))
                            evt.loadedDates = previousDates.Union(evt.loadedDates, System.StringComparer.Ordinal).OrderBy(Function(value) value, System.StringComparer.Ordinal).ToList()
                            evt.loadedHosts = previousHosts.Union(evt.loadedHosts, System.StringComparer.OrdinalIgnoreCase).OrderBy(Function(value) value, System.StringComparer.Ordinal).ToList()
                            Dim removeIdentity As System.Boolean = evt.anon AndAlso (Not previous.anon OrElse Not System.String.IsNullOrWhiteSpace(previous.userObf))
                            If previous.anon Then
                                evt.anon = True
                                evt.userObf = ""
                            End If
                            Dim previousUtc As System.DateTimeOffset
                            Dim incomingUtc As System.DateTimeOffset
                            If System.DateTimeOffset.TryParse(previous.lastUseUtc, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, previousUtc) AndAlso
                               System.DateTimeOffset.TryParse(evt.lastUseUtc, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, incomingUtc) AndAlso previousUtc >= incomingUtc Then
                                If Not removeIdentity AndAlso Not addedEvidence Then
                                    statusText = "exists_current"
                                    Return True
                                End If
                                evt.lastUseUtc = previous.lastUseUtc
                            End If
                            Dim previousFirst As System.DateTimeOffset
                            Dim incomingFirst As System.DateTimeOffset
                            If System.DateTimeOffset.TryParse(previous.firstUseUtc, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, previousFirst) AndAlso
                               System.DateTimeOffset.TryParse(evt.firstUseUtc, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, incomingFirst) AndAlso previousFirst < incomingFirst Then evt.firstUseUtc = previous.firstUseUtc
                        End If
                        WriteLicenseCounterJsonAtomically(fullPath, Newtonsoft.Json.JsonConvert.SerializeObject(evt, Newtonsoft.Json.Formatting.Indented))
                        statusText = "written"
                        LogLicenseCounterInternal("report_file_written", "Event=" & evt.DeliveryKey & ", TargetHash=" & target.TargetHash)
                        Return True
                    End Using
                Catch ex As System.IO.IOException
                    statusText = ex.GetType().Name
                    If attempt < 5 Then System.Threading.Thread.Sleep(GetLicenseCounterRetryDelayMs(attempt))
                End Try
            Next
            Return False
        End Function

        Private Shared Function AppendLicenseCounterReportLine(target As LicenseCounterTarget,
                                                               evt As LicenseCounterTransportEvent,
                                                               ByRef statusText As String) As Boolean
            statusText = ""

            Dim filePath As String = target.NormalizedValue
            Dim parentDir As String = Path.GetDirectoryName(filePath)

            If Not String.IsNullOrWhiteSpace(parentDir) AndAlso Not Directory.Exists(parentDir) Then
                Directory.CreateDirectory(parentDir)
            End If

            Dim line As String = JsonConvert.SerializeObject(evt, Formatting.None) & Environment.NewLine

            For attempt As Integer = 1 To 8
                Try
                    AppendLicenseCounterLineAtomically(filePath, line)
                    statusText = "appended"
                    LogLicenseCounterInternal("append_file_written", $"Path='{filePath}'")
                    Return True
                Catch ex As IOException
                    statusText = ex.Message
                    System.Threading.Thread.Sleep(GetLicenseCounterRetryDelayMs(attempt) + GetLicenseCounterAppendJitterMs())
                Catch ex As UnauthorizedAccessException
                    statusText = ex.Message
                    System.Threading.Thread.Sleep(GetLicenseCounterRetryDelayMs(attempt) + GetLicenseCounterAppendJitterMs())
                End Try
            Next

            Return False
        End Function

        Private Shared Sub AppendLicenseCounterLineAtomically(filePath As String, line As String)
            Dim fileExists As Boolean = File.Exists(filePath)
            Dim encodingWithoutBom As New UTF8Encoding(False)
            Dim encodingWithBom As New UTF8Encoding(True)
            Dim lineBytes As Byte() = encodingWithoutBom.GetBytes(line)

            Using fs As New FileStream(filePath,
                                       FileMode.Append,
                                       FileAccess.Write,
                                       FileShare.Read)
                If (Not fileExists) OrElse fs.Length = 0 Then
                    Dim bomBytes As Byte() = encodingWithBom.GetPreamble()
                    If bomBytes IsNot Nothing AndAlso bomBytes.Length > 0 Then
                        fs.Write(bomBytes, 0, bomBytes.Length)
                    End If
                End If

                fs.Write(lineBytes, 0, lineBytes.Length)
                fs.Flush(True)
            End Using
        End Sub

        Private Shared Function GetLicenseCounterAppendJitterMs() As Integer
            Dim bytes(1) As Byte

            Using rng = RandomNumberGenerator.Create()
                rng.GetBytes(bytes)
            End Using

            Dim value As Integer = BitConverter.ToUInt16(bytes, 0)
            Return 25 + (value Mod 125)
        End Function

        Private Shared Function SendLicenseCounterWebEvent(target As LicenseCounterTarget,
                                                           evt As LicenseCounterTransportEvent,
                                                           methodKind As LicenseCounterMethodKind,
                                                           webMode As LicenseCounterWebModeKind,
                                                           ByRef statusText As String,
                                                           Optional requireAck As System.Boolean = False) As Boolean
            If requireAck Then
                ' Active collectors use one canonical JSON event. A logging-only HTTP
                ' response must never retire the durable local marker in this mode.
                Dim payload As Newtonsoft.Json.Linq.JObject
                If System.String.Equals(evt.t, "E", System.StringComparison.OrdinalIgnoreCase) Then
                    payload = New Newtonsoft.Json.Linq.JObject() From {
                        {"ri_lc_v", evt.v}, {"t", "E"}, {"id", evt.errorEventId},
                        {"ec", evt.errorCode}, {"m", evt.relatedMonth},
                        {"f", evt.relatedFile}, {"detail", evt.detail}
                    }
                Else
                    payload = BuildLicenseCounterWebJsonPayload(evt)
                End If
                Return SendAcknowledgedLicenseCounterJson(target.NormalizedValue, payload, statusText)
            End If

            Dim reached As Boolean = False
            Dim statuses As New List(Of String)()

            If String.Equals(evt.t, "U", StringComparison.OrdinalIgnoreCase) Then
                Dim usagePathUrl As String = BuildLicenseCounterWebUsagePathUrl(target.NormalizedValue, evt)

                Select Case webMode
                    Case LicenseCounterWebModeKind.PathOnly
                        statuses.Add("GET-PATH:" & SendLicenseCounterHttpRequest("GET", usagePathUrl, "", reached))

                    Case LicenseCounterWebModeKind.Extended
                        Dim usageQueryUrl As String = BuildLicenseCounterWebUsageQueryUrl(target.NormalizedValue, evt)

                        statuses.Add("GET-PATH:" & SendLicenseCounterHttpRequest("GET", usagePathUrl, "", reached))
                        statuses.Add("GET-QUERY:" & SendLicenseCounterHttpRequest("GET", usageQueryUrl, "", reached))
                        statuses.Add("POST-JSON:" & SendLicenseCounterHttpRequest("POST",
                                                                                  target.NormalizedValue,
                                                                                  JsonConvert.SerializeObject(BuildLicenseCounterWebJsonPayload(evt), Formatting.None),
                                                                                  reached,
                                                                                  "application/json"))
                        statuses.Add("POST-FORM:" & SendLicenseCounterHttpRequest("POST",
                                                                                  target.NormalizedValue,
                                                                                  BuildLicenseCounterFormPayload(evt),
                                                                                  reached,
                                                                                  "application/x-www-form-urlencoded"))

                    Case Else
                        Dim usageQueryUrl As String = BuildLicenseCounterWebUsageQueryUrl(target.NormalizedValue, evt)

                        statuses.Add("GET-PATH:" & SendLicenseCounterHttpRequest("GET", usagePathUrl, "", reached))
                        statuses.Add("GET-QUERY:" & SendLicenseCounterHttpRequest("GET", usageQueryUrl, "", reached))
                End Select
            Else
                Dim errorQueryUrl As String = BuildLicenseCounterWebErrorQueryUrl(target.NormalizedValue, evt)
                statuses.Add("GET-ERROR:" & SendLicenseCounterHttpRequest("GET", errorQueryUrl, "", reached))
            End If

            statusText = String.Join(" | ", statuses.Where(Function(s) Not String.IsNullOrWhiteSpace(s)))

            LogLicenseCounterInternal("web_submission",
                                      $"TargetHash={target.TargetHash}, Method={methodKind}, WebMode={webMode}, Event={evt.DeliveryKey}, Reached={reached}, Status={statusText}")

            Return reached
        End Function

        ' Local collectors must not be forwarded through a corporate proxy.
        ' External HTTP(S) destinations retain the normal system proxy settings.
        Private Shared Function CreateLicenseCounterWebRequest(url As System.String) As System.Net.HttpWebRequest
            Dim destination As System.Uri = Nothing
            If Not System.Uri.TryCreate(url, System.UriKind.Absolute, destination) OrElse
               (destination.Scheme <> System.Uri.UriSchemeHttp AndAlso destination.Scheme <> System.Uri.UriSchemeHttps) Then
                Throw New System.UriFormatException("License counter destination must be an absolute HTTP(S) URL.")
            End If
            Dim request As System.Net.HttpWebRequest = DirectCast(System.Net.WebRequest.Create(destination), System.Net.HttpWebRequest)
            If destination.IsLoopback Then request.Proxy = Nothing
            Return request
        End Function

        Private Shared Function SendAcknowledgedLicenseCounterJson(url As System.String,
                                                                    payload As Newtonsoft.Json.Linq.JObject,
                                                                    ByRef statusText As System.String) As System.Boolean
            Try
                Dim request As System.Net.HttpWebRequest = CreateLicenseCounterWebRequest(url)
                request.Method = "POST"
                request.ContentType = "application/json"
                request.Accept = "application/json"
                request.AllowAutoRedirect = False
                request.Timeout = LicenseCounterWebTimeoutMs
                request.ReadWriteTimeout = LicenseCounterWebTimeoutMs
                request.CachePolicy = New System.Net.Cache.RequestCachePolicy(System.Net.Cache.RequestCacheLevel.NoCacheNoStore)
                Dim bytes As System.Byte() = System.Text.Encoding.UTF8.GetBytes(payload.ToString(Newtonsoft.Json.Formatting.None))
                request.ContentLength = bytes.Length
                Using stream As System.IO.Stream = request.GetRequestStream()
                    stream.Write(bytes, 0, bytes.Length)
                End Using
                Using response As System.Net.HttpWebResponse = DirectCast(request.GetResponse(), System.Net.HttpWebResponse)
                    Dim statusCode As System.Int32 = CInt(response.StatusCode)
                    statusText = statusCode.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    If statusCode < 200 OrElse statusCode >= 300 Then Return False
                    Using stream As System.IO.Stream = response.GetResponseStream()
                        Using reader As New System.IO.StreamReader(stream, System.Text.Encoding.UTF8)
                            ' Bound the ACK body; never accept an HTML login/error page.
                            Dim buffer(4096) As System.Char
                            Dim used As System.Int32 = 0
                            While used < buffer.Length
                                Dim count As System.Int32 = reader.Read(buffer, used, buffer.Length - used)
                                If count = 0 Then Exit While
                                used += count
                            End While
                            If used > 4096 Then
                                statusText &= ":ack_too_large"
                                Return False
                            End If
                            Dim ack As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(New System.String(buffer, 0, used))
                            Dim accepted As Newtonsoft.Json.Linq.JToken = ack("accepted")
                            If accepted IsNot Nothing AndAlso accepted.Type = Newtonsoft.Json.Linq.JTokenType.Boolean AndAlso accepted.Value(Of System.Boolean)() Then Return True
                        End Using
                    End Using
                    statusText &= ":ack_missing_or_rejected"
                End Using
            Catch ex As System.Net.WebException
                Dim response As System.Net.HttpWebResponse = TryCast(ex.Response, System.Net.HttpWebResponse)
                If response IsNot Nothing Then
                    Using response
                        statusText = CInt(response.StatusCode).ToString(System.Globalization.CultureInfo.InvariantCulture)
                    End Using
                Else
                    statusText = ex.Status.ToString()
                End If
            Catch ex As System.Exception
                statusText = ex.GetType().Name
            End Try
            Return False
        End Function

        Private Shared Function SendLicenseCounterHttpRequest(method As String,
                                                              url As String,
                                                              body As String,
                                                              ByRef reached As Boolean,
                                                              Optional contentType As String = "") As String
            Try
                Dim request As System.Net.HttpWebRequest = CreateLicenseCounterWebRequest(url)
                request.Method = method
                request.Timeout = LicenseCounterWebTimeoutMs
                request.ReadWriteTimeout = LicenseCounterWebTimeoutMs
                request.CachePolicy = New System.Net.Cache.RequestCachePolicy(System.Net.Cache.RequestCacheLevel.NoCacheNoStore)

                If Not String.IsNullOrWhiteSpace(contentType) Then
                    request.ContentType = contentType
                End If

                If String.Equals(method, "POST", StringComparison.OrdinalIgnoreCase) Then
                    Dim bodyBytes = Encoding.UTF8.GetBytes(body)
                    request.ContentLength = bodyBytes.Length

                    Using reqStream = request.GetRequestStream()
                        reqStream.Write(bodyBytes, 0, bodyBytes.Length)
                    End Using
                End If

                Using response = CType(request.GetResponse(), HttpWebResponse)
                    reached = True
                    Return CInt(response.StatusCode).ToString(CultureInfo.InvariantCulture)
                End Using

            Catch ex As WebException
                Dim response = TryCast(ex.Response, HttpWebResponse)

                If response IsNot Nothing Then
                    reached = True
                    Return CInt(response.StatusCode).ToString(CultureInfo.InvariantCulture)
                End If

                Return ex.Status.ToString()

            Catch ex As System.Exception
                Return ex.GetType().Name
            End Try
        End Function

        Private Shared Function BuildLicenseCounterWebUsagePathUrl(baseUrl As String, evt As LicenseCounterTransportEvent) As String
            Dim nonce As String = CreateLicenseCounterNonce()
            Dim url As String = baseUrl.TrimEnd("/"c)

            url &= $"/RI_LC_v1/m/{Uri.EscapeDataString(evt.month)}" &
                   $"/n/{Uri.EscapeDataString(evt.networkKey)}" &
                   $"/ns/{Uri.EscapeDataString(evt.networkScope)}" &
                   $"/p/{Uri.EscapeDataString(evt.productIdSafe)}" &
                   $"/h/{Uri.EscapeDataString(evt.hostId)}" &
                   $"/u/{Uri.EscapeDataString(evt.userKey)}" &
                   $"/b/{Uri.EscapeDataString(evt.billingEventId)}" &
                   $"/e/{Uri.EscapeDataString(evt.hostEventId)}" &
                   $"/a/{If(evt.anon, "1", "0")}" &
                   $"/r/{nonce}"

            If Not evt.anon AndAlso Not String.IsNullOrWhiteSpace(evt.userObf) Then
                url &= $"/du/{Uri.EscapeDataString(evt.userObf)}"
            End If

            Return url
        End Function

        Private Shared Function BuildLicenseCounterWebUsageQueryUrl(baseUrl As String, evt As LicenseCounterTransportEvent) As String
            Dim separator As String = If(baseUrl.Contains("?"), "&", "?")
            Dim sb As New StringBuilder()

            sb.Append(baseUrl)
            sb.Append(separator)
            sb.Append("ri_lc_v=1")
            sb.Append("&t=U")
            sb.Append("&m=" & Uri.EscapeDataString(evt.month))
            sb.Append("&n=" & Uri.EscapeDataString(evt.networkKey))
            sb.Append("&ns=" & Uri.EscapeDataString(evt.networkScope))
            sb.Append("&ani=" & Uri.EscapeDataString(String.Join("|", If(evt.allowedNetworkIds, New List(Of String)()))))
            sb.Append("&p=" & Uri.EscapeDataString(evt.productIdSafe))
            sb.Append("&h=" & Uri.EscapeDataString(evt.hostId))
            sb.Append("&u=" & Uri.EscapeDataString(evt.userKey))
            sb.Append("&b=" & Uri.EscapeDataString(evt.billingEventId))
            sb.Append("&e=" & Uri.EscapeDataString(evt.hostEventId))
            sb.Append("&a=" & If(evt.anon, "1", "0"))
            sb.Append("&r=" & CreateLicenseCounterNonce())
            sb.Append("&lastUseUtc=" & System.Uri.EscapeDataString(evt.lastUseUtc))
            sb.Append("&reportedUtc=" & System.Uri.EscapeDataString(evt.reportedUtc))
            sb.Append("&reportIntervalMinutes=" & evt.reportIntervalMinutes.ToString(System.Globalization.CultureInfo.InvariantCulture))
            sb.Append("&loadedDates=" & System.Uri.EscapeDataString(System.String.Join(",", evt.loadedDates)))
            sb.Append("&loadedHosts=" & System.Uri.EscapeDataString(System.String.Join(",", evt.loadedHosts)))

            If Not evt.anon AndAlso Not String.IsNullOrWhiteSpace(evt.userObf) Then
                sb.Append("&du=" & Uri.EscapeDataString(evt.userObf))
            End If

            Return sb.ToString()
        End Function

        Private Shared Function BuildLicenseCounterWebErrorQueryUrl(baseUrl As String, evt As LicenseCounterTransportEvent) As String
            Dim separator As String = If(baseUrl.Contains("?"), "&", "?")
            Dim sb As New StringBuilder()

            sb.Append(baseUrl)
            sb.Append(separator)
            sb.Append("ri_lc_v=1")
            sb.Append("&t=E")
            sb.Append("&id=" & Uri.EscapeDataString(evt.errorEventId))
            sb.Append("&ec=" & Uri.EscapeDataString(evt.errorCode))
            sb.Append("&m=" & Uri.EscapeDataString(If(evt.relatedMonth, "")))
            sb.Append("&f=" & Uri.EscapeDataString(If(evt.relatedFile, "")))
            sb.Append("&r=" & CreateLicenseCounterNonce())

            Return sb.ToString()
        End Function

        Private Shared Function BuildLicenseCounterWebJsonPayload(evt As LicenseCounterTransportEvent) As JObject
            Dim jObj As New JObject() From {
                {"ri_lc_v", evt.v},
                {"t", evt.t},
                {"m", evt.month},
                {"n", evt.networkKey},
                {"ns", evt.networkScope},
                {"p", evt.productId},
                {"h", evt.hostId},
                {"u", evt.userKey},
                {"b", evt.billingEventId},
                {"e", evt.hostEventId},
                {"a", If(evt.anon, 1, 0)},
                {"firstUseUtc", evt.firstUseUtc},
                {"lastUseUtc", evt.lastUseUtc},
                {"reportedUtc", evt.reportedUtc},
                {"reportIntervalMinutes", evt.reportIntervalMinutes},
                {"loadedDates", Newtonsoft.Json.Linq.JArray.FromObject(evt.loadedDates)},
                {"loadedHosts", Newtonsoft.Json.Linq.JArray.FromObject(evt.loadedHosts)}
            }

            If evt.allowedNetworkIds IsNot Nothing AndAlso evt.allowedNetworkIds.Count > 0 Then
                jObj("allowedNetworkIds") = JArray.FromObject(evt.allowedNetworkIds)
            End If

            If Not evt.anon AndAlso Not String.IsNullOrWhiteSpace(evt.userObf) Then
                jObj("du") = evt.userObf
            End If

            Return jObj
        End Function

        Private Shared Function BuildLicenseCounterFormPayload(evt As LicenseCounterTransportEvent) As String
            Dim parts As New List(Of String) From {
                "ri_lc_v=1",
                "t=" & Uri.EscapeDataString(evt.t),
                "m=" & Uri.EscapeDataString(evt.month),
                "n=" & Uri.EscapeDataString(evt.networkKey),
                "ns=" & Uri.EscapeDataString(evt.networkScope),
                "p=" & Uri.EscapeDataString(evt.productId),
                "h=" & Uri.EscapeDataString(evt.hostId),
                "u=" & Uri.EscapeDataString(evt.userKey),
                "b=" & Uri.EscapeDataString(evt.billingEventId),
                "e=" & Uri.EscapeDataString(evt.hostEventId),
                "a=" & If(evt.anon, "1", "0"),
                "lastUseUtc=" & System.Uri.EscapeDataString(evt.lastUseUtc),
                "reportedUtc=" & System.Uri.EscapeDataString(evt.reportedUtc),
                "reportIntervalMinutes=" & evt.reportIntervalMinutes.ToString(System.Globalization.CultureInfo.InvariantCulture),
                "loadedDates=" & System.Uri.EscapeDataString(System.String.Join(",", evt.loadedDates)),
                "loadedHosts=" & System.Uri.EscapeDataString(System.String.Join(",", evt.loadedHosts))
            }

            If evt.allowedNetworkIds IsNot Nothing AndAlso evt.allowedNetworkIds.Count > 0 Then
                parts.Add("ani=" & Uri.EscapeDataString(String.Join("|", evt.allowedNetworkIds)))
            End If

            If Not evt.anon AndAlso Not String.IsNullOrWhiteSpace(evt.userObf) Then
                parts.Add("du=" & Uri.EscapeDataString(evt.userObf))
            End If

            Return String.Join("&", parts)
        End Function

        Private Shared Function ParseLicenseCounterTargets(rawPath As String,
                                                           methodKind As LicenseCounterMethodKind) As List(Of LicenseCounterTarget)
            Dim results As New List(Of LicenseCounterTarget)()

            If String.IsNullOrWhiteSpace(rawPath) Then
                Return results
            End If

            If rawPath.Contains(";"c) AndAlso
               Not rawPath.Contains(vbCr) AndAlso
               Not rawPath.Contains(vbLf) AndAlso
               Not rawPath.Contains("|"c) Then
                LogLicenseCounterInternal("path_separator_warning",
                                          "LicenseCounterPath contains ';' but multiple targets must be separated by line breaks or '|'.")
            End If

            Dim separators = If(rawPath.Contains(vbCr) OrElse rawPath.Contains(vbLf),
                                rawPath.Split({vbCrLf, vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries),
                                rawPath.Split({"|"c}, StringSplitOptions.RemoveEmptyEntries))

            For Each raw In separators
                Dim trimmed As String = raw.Trim()

                If String.IsNullOrWhiteSpace(trimmed) Then
                    Continue For
                End If

                Dim isWebTarget As Boolean =
                    trimmed.StartsWith("http://", StringComparison.OrdinalIgnoreCase) OrElse
                    trimmed.StartsWith("https://", StringComparison.OrdinalIgnoreCase)

                Dim normalizedValue As String = If(isWebTarget, trimmed, ExpandEnvironmentVariables(trimmed))

                If String.IsNullOrWhiteSpace(normalizedValue) Then
                    LogLicenseCounterInternal("target_parse_error",
                                              $"Raw='{trimmed}' could not be expanded or normalized.")
                    Continue For
                End If

                Dim target As New LicenseCounterTarget() With {
                    .RawValue = trimmed,
                    .NormalizedValue = normalizedValue,
                    .TargetHash = "t" & FirstHex(HashSha256("RI-LC-TARGET-v1|" & normalizedValue), 16)
                }

                If isWebTarget Then
                    target.TargetType = LicenseCounterTargetType.Web
                ElseIf methodKind = LicenseCounterMethodKind.AppendFile Then
                    target.TargetType = LicenseCounterTargetType.AppendFile
                Else
                    If Directory.Exists(normalizedValue) Then
                        target.TargetType = LicenseCounterTargetType.Directory
                    ElseIf File.Exists(normalizedValue) Then
                        target.TargetType = If(methodKind = LicenseCounterMethodKind.AppendFile,
                                               LicenseCounterTargetType.AppendFile,
                                               LicenseCounterTargetType.Invalid)
                        If target.TargetType = LicenseCounterTargetType.Invalid Then
                            target.Note = "target_is_file_but_append_not_enabled"
                        End If
                    ElseIf normalizedValue.EndsWith("\", StringComparison.Ordinal) OrElse
                           normalizedValue.EndsWith("/", StringComparison.Ordinal) OrElse
                           String.IsNullOrWhiteSpace(Path.GetExtension(normalizedValue)) Then
                        target.TargetType = LicenseCounterTargetType.Directory
                    Else
                        target.TargetType = If(methodKind = LicenseCounterMethodKind.AppendFile,
                                               LicenseCounterTargetType.AppendFile,
                                               LicenseCounterTargetType.Invalid)
                        If target.TargetType = LicenseCounterTargetType.Invalid Then
                            target.Note = "target_is_file_but_append_not_enabled"
                        End If
                    End If
                End If

                LogLicenseCounterInternal("target_parsed",
                                          $"Type={target.TargetType}, Hash={target.TargetHash}, Note={target.Note}")

                results.Add(target)
            Next

            Return results
        End Function

        Private Shared Function ParseLicenseCounterMethod(rawValue As String) As LicenseCounterMethodKind
            Dim baseMethod As String = GetLicenseCounterBaseMethod(rawValue)

            Select Case baseMethod.Trim().ToLowerInvariant()
                Case "", "0", "auto"
                    Return LicenseCounterMethodKind.Auto
                Case "1", "file"
                    Return LicenseCounterMethodKind.FileOnly
                Case "2", "webbasic"
                    Return LicenseCounterMethodKind.WebBasic
                Case "3", "webextended"
                    Return LicenseCounterMethodKind.WebExtended
                Case "4", "localonly"
                    Return LicenseCounterMethodKind.LocalOnly
                Case "5", "appendfile"
                    Return LicenseCounterMethodKind.AppendFile
                Case "6", "webpathonly", "webpath"
                    Return LicenseCounterMethodKind.WebPathOnly
                Case Else
                    Return LicenseCounterMethodKind.Auto
            End Select
        End Function

        Private Shared Function GetLicenseCounterBaseMethod(rawValue As String) As String
            If String.IsNullOrWhiteSpace(rawValue) Then
                Return "Auto"
            End If

            Dim parts As String() = rawValue.Split(";"c)
            If parts.Length = 0 OrElse String.IsNullOrWhiteSpace(parts(0)) Then
                Return "Auto"
            End If

            Return parts(0).Trim()
        End Function

        Private Shared Function ParseLicenseCounterWebMode(rawValue As String,
                                                           methodKind As LicenseCounterMethodKind) As LicenseCounterWebModeKind
            Select Case methodKind
                Case LicenseCounterMethodKind.WebExtended
                    Return LicenseCounterWebModeKind.Extended

                Case LicenseCounterMethodKind.WebPathOnly
                    Return LicenseCounterWebModeKind.PathOnly

                Case LicenseCounterMethodKind.WebBasic
                    Return LicenseCounterWebModeKind.Basic
            End Select

            If methodKind <> LicenseCounterMethodKind.Auto OrElse String.IsNullOrWhiteSpace(rawValue) Then
                Return LicenseCounterWebModeKind.Basic
            End If

            Dim parts As String() = rawValue.Split(";"c)

            For i As Integer = 1 To parts.Length - 1
                Dim part As String = parts(i).Trim()

                If part.StartsWith("WebMode=", StringComparison.OrdinalIgnoreCase) Then
                    Dim modeValue As String = part.Substring("WebMode=".Length).Trim()

                    Select Case modeValue.ToLowerInvariant()
                        Case "basic", "webbasic"
                            Return LicenseCounterWebModeKind.Basic
                        Case "pathonly", "webpathonly", "webpath"
                            Return LicenseCounterWebModeKind.PathOnly
                        Case "extended", "webextended"
                            Return LicenseCounterWebModeKind.Extended
                    End Select
                End If
            Next

            Return LicenseCounterWebModeKind.Basic
        End Function

        Private Shared Function ParseLicenseCounterResubmitInstruction(rawValue As String) As LicenseCounterResubmitInstruction
            Dim instruction As New LicenseCounterResubmitInstruction()

            If String.IsNullOrWhiteSpace(rawValue) Then
                Return instruction
            End If

            Dim parts As String() = rawValue.Split(";"c)

            For i As Integer = 1 To parts.Length - 1
                Dim part As String = parts(i).Trim()

                If part.Equals("ResubmitAll", StringComparison.OrdinalIgnoreCase) Then
                    instruction.ResubmitAll = True
                ElseIf part.StartsWith("ResubmitMonth=", StringComparison.OrdinalIgnoreCase) Then
                    instruction.ResubmitMonth = part.Substring("ResubmitMonth=".Length).Trim()
                ElseIf part.StartsWith("ResubmitFrom=", StringComparison.OrdinalIgnoreCase) Then
                    Dim rawDate As String = part.Substring("ResubmitFrom=".Length).Trim()
                    Dim parsedDate As System.DateTime

                    If System.DateTime.TryParseExact(rawDate,
                                                    "yyyy-MM-dd",
                                                    CultureInfo.InvariantCulture,
                                                    DateTimeStyles.AssumeUniversal Or DateTimeStyles.AdjustToUniversal,
                                                    parsedDate) Then
                        instruction.ResubmitFromUtc = parsedDate.Date
                    End If
                ElseIf part.StartsWith("Token=", StringComparison.OrdinalIgnoreCase) Then
                    instruction.Token = part.Substring("Token=".Length).Trim()
                End If
            Next

            Return instruction
        End Function

        Private Shared Sub ApplyLicenseCounterResubmissionInstruction(baseFolder As String,
                                                                      state As LicenseCounterStateFile,
                                                                      instruction As LicenseCounterResubmitInstruction)
            If instruction Is Nothing Then
                Return
            End If

            Dim hasScope As Boolean = instruction.ResubmitAll OrElse
                                      Not String.IsNullOrWhiteSpace(instruction.ResubmitMonth) OrElse
                                      instruction.ResubmitFromUtc.HasValue

            If Not hasScope OrElse String.IsNullOrWhiteSpace(instruction.Token) Then
                Return
            End If

            If state.appliedResubmitTokens.ContainsKey(instruction.Token) Then
                Return
            End If

            For Each filePath In Directory.GetFiles(baseFolder, $"{LicenseCounterMarkerPrefix}*.json", SearchOption.TopDirectoryOnly)
                Dim evt = TryReadLicenseCounterUsageEvent(filePath)

                If evt Is Nothing Then
                    Continue For
                End If

                If ShouldResubmitLicenseCounterEvent(evt, instruction) Then
                    state.delivery.Remove(evt.DeliveryKey)
                    RemoveLicenseCounterGroupDeliveryEntries(state, evt.billingEventId)
                End If
            Next

            state.appliedResubmitTokens(instruction.Token) = System.DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture)
        End Sub

        Private Shared Function ShouldResubmitLicenseCounterEvent(evt As LicenseCounterTransportEvent,
                                                                  instruction As LicenseCounterResubmitInstruction) As Boolean
            If instruction Is Nothing Then
                Return False
            End If

            If instruction.ResubmitAll Then
                Return True
            End If

            If Not String.IsNullOrWhiteSpace(instruction.ResubmitMonth) Then
                If String.Compare(evt.month, instruction.ResubmitMonth, StringComparison.OrdinalIgnoreCase) >= 0 Then
                    Return True
                End If
            End If

            If instruction.ResubmitFromUtc.HasValue Then
                Dim firstUseUtc As System.DateTime

                If System.DateTime.TryParse(evt.firstUseUtc,
                                            CultureInfo.InvariantCulture,
                                            DateTimeStyles.RoundtripKind,
                                            firstUseUtc) Then
                    If firstUseUtc >= instruction.ResubmitFromUtc.Value Then
                        Return True
                    End If
                End If
            End If

            Return False
        End Function

        Private Shared Sub RefreshKnownMarkersAndDetectMissing(baseFolder As String, state As LicenseCounterStateFile)
            Dim currentFiles As New Dictionary(Of String, LicenseCounterMarkerRecord)(StringComparer.OrdinalIgnoreCase)

            For Each filePath In Directory.GetFiles(baseFolder, $"{LicenseCounterMarkerPrefix}*.json", SearchOption.TopDirectoryOnly)
                Try
                    Dim json As String = File.ReadAllText(filePath, Encoding.UTF8)
                    Dim marker = JsonConvert.DeserializeObject(Of LicenseCounterMarkerRecord)(json)

                    If marker IsNot Nothing AndAlso
                       String.Equals(marker.t, "U", StringComparison.OrdinalIgnoreCase) AndAlso
                       Not String.IsNullOrWhiteSpace(marker.hostEventId) Then
                        currentFiles(marker.hostEventId) = marker
                        EnsureKnownMarkerEntry(state, marker.hostEventId, Path.GetFileName(filePath))
                    End If
                Catch
                End Try
            Next

            For Each knownKey In state.knownHostEventIds.Keys.ToList()
                If Not currentFiles.ContainsKey(knownKey) Then
                    Dim missingFile As String = state.knownHostEventIds(knownKey).file
                    QueueLicenseCounterError(state, "local_marker_missing", "", missingFile, "Marker file known in state is no longer present.")
                End If
            Next
        End Sub

        Private Shared Sub ApplyLicenseCounterRetention(baseFolder As System.String,
                                                        state As LicenseCounterStateFile,
                                                        enabledTargets As System.Collections.Generic.List(Of LicenseCounterTarget),
                                                        Optional forceAnonymous As System.Boolean = False)
            Try
                For Each evt As LicenseCounterTransportEvent In CollectLicenseCounterUsageGroups(baseFolder, forceAnonymous)
                    If IsLicenseCounterMonthWithinRetention(evt.month) Then Continue For
                    ' Require proof for the latest complete date union, not an older host
                    ' ACK. LocalOnly or a new/failed target must preserve outage backlog.
                    Dim reportedEverywhere As System.Boolean = enabledTargets IsNot Nothing AndAlso enabledTargets.Count > 0
                    If reportedEverywhere Then reportedEverywhere = enabledTargets.All(Function(target) Not ShouldAttemptLicenseCounterEvent(state, evt, target.TargetHash))
                    If Not reportedEverywhere Then
                        Dim noticeId As System.String = "x" & FirstHex(HashSha256("RI-LC-RETENTION-PRESERVED-v1|" & evt.billingEventId), 32)
                        If Not state.delivery.ContainsKey(noticeId) AndAlso Not state.pendingErrors.Any(Function(item) System.String.Equals(item.errorEventId, noticeId, System.StringComparison.OrdinalIgnoreCase)) Then
                            state.pendingErrors.Add(New LicenseCounterErrorEvent() With {
                                .v = LicenseCounterFormatVersion,
                                .t = "E",
                                .errorEventId = noticeId,
                                .errorCode = "retention_preserved_unreported_marker",
                                .relatedMonth = evt.month,
                                .relatedFile = System.IO.Path.GetFileName(evt.Source),
                                .detail = "Marker exceeded normal retention but was preserved until every enabled target confirms delivery of all observed dates.",
                                .createdUtc = System.DateTime.UtcNow.ToString("o", System.Globalization.CultureInfo.InvariantCulture)
                            })
                            LogLicenseCounterInternal("retention_preserved_unreported_marker", "Event=" & evt.DeliveryKey & ", Month=" & evt.month)
                        End If
                        Continue For
                    End If
                    For Each filePath As System.String In evt.SourceMarkerFiles
                        Try
                            Dim marker As LicenseCounterTransportEvent = TryReadLicenseCounterUsageEvent(filePath)
                            If marker Is Nothing Then Continue For
                            System.IO.File.Delete(filePath)
                            If Not System.IO.File.Exists(filePath) Then
                                state.knownHostEventIds.Remove(marker.hostEventId)
                                state.delivery.Remove(marker.hostEventId)
                            End If
                        Catch ex As System.Exception
                            LogLicenseCounterInternal("retention_delete_error", "Event=" & evt.DeliveryKey & ", Error=" & ex.GetType().Name)
                        End Try
                    Next
                    If evt.SourceMarkerFiles.All(Function(filePath) Not System.IO.File.Exists(filePath)) Then RemoveLicenseCounterGroupDeliveryEntries(state, evt.billingEventId)
                Next
            Catch ex As System.Exception
                LogLicenseCounterInternal("retention_error", ex.Message)
            End Try
        End Sub

        Private Shared Sub RemoveLicenseCounterGroupDeliveryEntries(state As LicenseCounterStateFile, billingEventId As System.String)
            Dim prefix As System.String = "d" & billingEventId & "_"
            For Each key As System.String In state.delivery.Keys.Where(Function(value) value.StartsWith(prefix, System.StringComparison.OrdinalIgnoreCase)).ToList()
                state.delivery.Remove(key)
            Next
        End Sub

        Private Shared Function IsLicenseCounterMonthWithinRetention(monthValue As String) As Boolean
            Dim parsed As System.DateTime

            If Not System.DateTime.TryParseExact(monthValue & "-01",
                                                 "yyyy-MM-dd",
                                                 CultureInfo.InvariantCulture,
                                                 DateTimeStyles.None,
                                                 parsed) Then
                Return True
            End If

            Dim currentMonthStart As New System.DateTime(System.DateTime.Now.Year, System.DateTime.Now.Month, 1)
            Dim oldestKeptMonthStart As System.DateTime = currentMonthStart.AddMonths(-(LicenseCounterRetentionMonths - 1))
            Return parsed >= oldestKeptMonthStart
        End Function

        Private Shared Sub EnsureKnownMarkerEntry(state As LicenseCounterStateFile, hostEventId As String, fileName As String)
            If String.IsNullOrWhiteSpace(hostEventId) Then
                Return
            End If

            Dim nowUtc As String = System.DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture)

            If Not state.knownHostEventIds.ContainsKey(hostEventId) Then
                state.knownHostEventIds(hostEventId) = New LicenseCounterKnownMarker() With {
                    .file = fileName,
                    .firstSeenUtc = nowUtc,
                    .lastSeenUtc = nowUtc
                }
            Else
                state.knownHostEventIds(hostEventId).file = fileName
                state.knownHostEventIds(hostEventId).lastSeenUtc = nowUtc
            End If
        End Sub

        Private Shared Sub QueueLicenseCounterError(state As LicenseCounterStateFile,
                                                    errorCode As String,
                                                    relatedMonth As String,
                                                    relatedFile As String,
                                                    detail As String)
            If state.pendingErrors.Any(Function(p) p.errorCode = errorCode AndAlso
                                                   String.Equals(p.relatedMonth, relatedMonth, StringComparison.OrdinalIgnoreCase) AndAlso
                                                   String.Equals(p.relatedFile, relatedFile, StringComparison.OrdinalIgnoreCase)) Then
                Return
            End If

            state.pendingErrors.Add(New LicenseCounterErrorEvent() With {
                .v = LicenseCounterFormatVersion,
                .t = "E",
                .errorEventId = "x" & FirstHex(HashSha256("RI-LC-ERR-v1|" &
                                                          errorCode & "|" &
                                                          relatedMonth & "|" &
                                                          relatedFile & "|" &
                                                          detail & "|" &
                                                          System.DateTime.UtcNow.ToString("yyyyMMddHHmmss", CultureInfo.InvariantCulture)), 32),
                .errorCode = errorCode,
                .relatedMonth = relatedMonth,
                .relatedFile = relatedFile,
                .detail = detail,
                .createdUtc = System.DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture)
            })
        End Sub

        Private Shared Function GetLicenseCounterDeliveryStatus(state As LicenseCounterStateFile,
                                                                deliveryKey As String,
                                                                targetHash As String) As LicenseCounterDeliveryStatus
            If state Is Nothing OrElse String.IsNullOrWhiteSpace(deliveryKey) OrElse String.IsNullOrWhiteSpace(targetHash) Then
                Return Nothing
            End If

            If Not state.delivery.ContainsKey(deliveryKey) Then
                Return Nothing
            End If

            If Not state.delivery(deliveryKey).ContainsKey(targetHash) Then
                Return Nothing
            End If

            Return state.delivery(deliveryKey)(targetHash)
        End Function

        Private Shared Sub UpdateLicenseCounterDelivery(state As LicenseCounterStateFile,
                                                        deliveryKey As String,
                                                        targetHash As String,
                                                        succeeded As Boolean,
                                                        statusText As String)
            If Not state.delivery.ContainsKey(deliveryKey) Then
                state.delivery(deliveryKey) = New Dictionary(Of String, LicenseCounterDeliveryStatus)(StringComparer.OrdinalIgnoreCase)
            End If

            If Not state.delivery(deliveryKey).ContainsKey(targetHash) Then
                state.delivery(deliveryKey)(targetHash) = New LicenseCounterDeliveryStatus()
            End If

            Dim item = state.delivery(deliveryKey)(targetHash)
            Dim nowUtc As String = System.DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture)

            item.lastAttemptUtc = nowUtc
            item.lastStatus = statusText

            If succeeded Then
                item.lastSuccessUtc = nowUtc
            End If
        End Sub

        Private Shared Sub WithLicenseCounterStateLock(baseFolder As String, action As Action(Of LicenseCounterStateFile))
            Dim mutexName As String = "Local\RI_LicenseCounter_State_" &
                                      FirstHex(HashSha256(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)), 16)

            Dim hasHandle As Boolean = False
            Dim stateMutex As System.Threading.Mutex = Nothing

            Try
                stateMutex = New System.Threading.Mutex(False, mutexName)

                For attempt As Integer = 1 To 5
                    Try
                        hasHandle = stateMutex.WaitOne(GetLicenseCounterRetryDelayMs(attempt))
                        If hasHandle Then
                            Exit For
                        End If
                    Catch ex As System.Threading.AbandonedMutexException
                        hasHandle = True
                        Exit For
                    End Try
                Next

                If Not hasHandle Then
                    LogLicenseCounterDiagnostic("state_lock_timeout", "Unable to acquire the local counter-state mutex")
                    Return
                End If

                Directory.CreateDirectory(baseFolder)

                Dim state = LoadLicenseCounterState(baseFolder)
                action(state)
                SaveLicenseCounterState(baseFolder, state)

            Catch ex As System.Exception
                LogLicenseCounterDiagnostic("state_lock_error", ex.GetType().Name)
                LogLicenseCounterInternal("state_lock_error", ex.Message)
            Finally
                Try
                    If hasHandle AndAlso stateMutex IsNot Nothing Then
                        stateMutex.ReleaseMutex()
                    End If
                Catch
                End Try

                Try
                    If stateMutex IsNot Nothing Then
                        stateMutex.Dispose()
                    End If
                Catch
                End Try
            End Try
        End Sub

        Private Shared Function LoadLicenseCounterState(baseFolder As String) As LicenseCounterStateFile
            Dim statePath As String = Path.Combine(baseFolder, LicenseCounterStateFileName)

            Try
                If Not File.Exists(statePath) Then
                    Return New LicenseCounterStateFile() With {
                        .lastUpdatedUtc = System.DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture)
                    }
                End If

                Dim json As String = File.ReadAllText(statePath, Encoding.UTF8)
                Dim state = JsonConvert.DeserializeObject(Of LicenseCounterStateFile)(json)

                If state Is Nothing Then
                    Throw New InvalidDataException("State JSON is empty.")
                End If

                If Not String.Equals(state.ih, ComputeStateIntegrityHash(state), StringComparison.OrdinalIgnoreCase) Then
                    Dim rebuilt As New LicenseCounterStateFile() With {
                        .lastUpdatedUtc = System.DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture)
                    }
                    QueueLicenseCounterError(rebuilt, "state_integrity_error", "", LicenseCounterStateFileName, "State integrity hash mismatch.")
                    Return rebuilt
                End If

                Return state

            Catch ex As System.Exception
                Dim rebuilt As New LicenseCounterStateFile() With {
                    .lastUpdatedUtc = System.DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture)
                }
                QueueLicenseCounterError(rebuilt, "state_integrity_error", "", LicenseCounterStateFileName, ex.Message)
                Return rebuilt
            End Try
        End Function

        Private Shared Sub SaveLicenseCounterState(baseFolder As String, state As LicenseCounterStateFile)
            Dim statePath As String = Path.Combine(baseFolder, LicenseCounterStateFileName)

            state.lastUpdatedUtc = System.DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture)
            state.ih = ComputeStateIntegrityHash(state)

            Dim json As String = JsonConvert.SerializeObject(state, Formatting.Indented)
            WriteLicenseCounterJsonAtomically(statePath, json)
        End Sub

        Private Shared Function ComputeStateIntegrityHash(state As LicenseCounterStateFile) As String
            Dim clone As New LicenseCounterStateFile() With {
                .v = state.v,
                .knownHostEventIds = state.knownHostEventIds,
                .delivery = state.delivery,
                .pendingErrors = state.pendingErrors,
                .lastUpdatedUtc = state.lastUpdatedUtc,
                .ih = ""
            }

            Return HashSha256("RI-LC-STATE-IH-v1|" & JsonConvert.SerializeObject(clone, Formatting.None))
        End Function

        Private Shared Function ComputeMarkerIntegrityHash(record As LicenseCounterMarkerRecord) As String
            Return HashSha256(
                "RI-LC-MARKER-IH-v1|" &
                record.month & "|" &
                record.networkKey & "|" &
                record.networkScope & "|" &
                String.Join("|", If(record.allowedNetworkIds, New List(Of String)())) & "|" &
                record.productId & "|" &
                record.productIdSafe & "|" &
                record.hostId & "|" &
                record.userKey & "|" &
                record.billingEventId & "|" &
                record.hostEventId & "|" &
                record.firstUseUtc & "|" &
                record.anon.ToString(CultureInfo.InvariantCulture) & "|" &
                record.actualNetworkId & "|" &
                record.actualNetworkKey & "|" &
                If(record.userObf, "") & "|" &
                LicenseCounterLicenseFingerprint)
        End Function

        Private Shared Function GetLicenseCounterBaseFolder() As String
            Return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), AN2, "licensecounter")
        End Function

        Private Shared Function NormalizeLicenseCounterUserId(value As String) As String
            Return NormalizeLicenseCounterValue(value)
        End Function

        Private Shared Function NormalizeLicenseCounterValue(value As String) As String
            If String.IsNullOrWhiteSpace(value) Then
                Return ""
            End If

            Dim normalized As String = value.Normalize(NormalizationForm.FormC).Trim()
            normalized = normalized.Replace(vbCr, " ").Replace(vbLf, " ").Replace(vbTab, " ")
            Return normalized.Trim()
        End Function

        Private Shared Function NormalizeLicenseCounterHostId(hostId As String) As String
            Select Case If(hostId, "").Trim().ToUpperInvariant()
                Case "WD", "WORD"
                    Return "WD"
                Case "EX", "XL", "EXCEL"
                    Return "EX"
                Case "OL", "OUTLOOK"
                    Return "OL"
                Case Else
                    Return "UNKNOWN"
            End Select
        End Function

        Private Shared Function NormalizeAllowedNetworkIds(allowedDomains As List(Of String)) As List(Of String)
            If allowedDomains Is Nothing Then
                Return New List(Of String)()
            End If

            Return allowedDomains.
                Select(Function(d) NormalizeLicenseCounterValue(d)).
                Where(Function(d) Not String.IsNullOrWhiteSpace(d)).
                Distinct(StringComparer.OrdinalIgnoreCase).
                OrderBy(Function(d) d, StringComparer.OrdinalIgnoreCase).
                ToList()
        End Function

        Private Shared Function BuildJointNetworkId(allowedDomains As List(Of String)) As String
            Dim normalized As List(Of String) = NormalizeAllowedNetworkIds(allowedDomains)

            If normalized.Count = 0 Then
                Return "UNKNOWN_NETWORK"
            End If

            Return String.Join("|", normalized)
        End Function

        Private Shared Function BuildNetworkKey(normalizedNetworkId As String) As String
            Return "n" & FirstHex(HashSha256("RI-LC-NET-v1|" & normalizedNetworkId), 16)
        End Function

        Private Shared Function BuildJointNetworkKey(jointNetworkId As String) As String
            Return "n" & FirstHex(HashSha256("RI-LC-JNET-v1|" & jointNetworkId), 16)
        End Function

        Private Shared Function BuildUserKey(jointNetworkKey As String, normalizedUserId As String) As String
            Return "u" & FirstHex(HashSha256(
                "RI-LC-USER-v1|" &
                jointNetworkKey & "|" &
                LicenseCounterSalt & "|" &
                normalizedUserId), 32)
        End Function

        Private Shared Function BuildBillingEventId(month As String, jointNetworkKey As String, productId As String, userKey As String) As String
            Return "b" & FirstHex(HashSha256(
                "RI-LC-BILL-v1|" &
                month & "|" &
                jointNetworkKey & "|" &
                productId & "|" &
                userKey), 32)
        End Function

        Private Shared Function BuildHostEventId(month As String, jointNetworkKey As String, productId As String, hostId As String, userKey As String) As String
            Return "e" & FirstHex(HashSha256(
                "RI-LC-HOST-v1|" &
                month & "|" &
                jointNetworkKey & "|" &
                productId & "|" &
                hostId & "|" &
                userKey), 32)
        End Function

        Private Shared Function BuildUserObf(normalizedUserId As String) As String
            Return "o1_" & Base64UrlEncode(Encoding.UTF8.GetBytes(normalizedUserId))
        End Function

        Private Shared Function TryDecodeUserObf(userObf As String, ByRef userId As String) As Boolean
            userId = ""

            Try
                If String.IsNullOrWhiteSpace(userObf) Then
                    Return False
                End If

                If userObf.StartsWith("o1_", StringComparison.OrdinalIgnoreCase) Then
                    userId = Encoding.UTF8.GetString(Base64UrlDecode(userObf.Substring(3)))
                    Return True
                End If

                Return False

            Catch
                Return False
            End Try
        End Function

        Private Shared Function MakeLicenseCounterSafeProductId(productId As String) As String
            Dim input As String = NormalizeLicenseCounterValue(productId)

            If String.IsNullOrWhiteSpace(input) Then
                Return "UNKNOWN_PRODUCT"
            End If

            Dim sb As New StringBuilder()

            For Each ch In input
                If Char.IsLetterOrDigit(ch) OrElse ch = "-"c OrElse ch = "_"c OrElse ch = "."c Then
                    sb.Append(ch)
                Else
                    sb.Append("_"c)
                End If
            Next

            Dim result As String = sb.ToString()

            If result.Length > 80 Then
                result = result.Substring(0, 80)
            End If

            If String.IsNullOrWhiteSpace(result) Then
                Return "UNKNOWN_PRODUCT"
            End If

            Return result
        End Function

        Private Shared Function CreateLicenseCounterNonce() As String
            Dim bytes(7) As Byte
            Using rng = RandomNumberGenerator.Create()
                rng.GetBytes(bytes)
            End Using

            Return BitConverter.ToString(bytes).Replace("-", "").ToLowerInvariant()
        End Function

        Private Shared Function HashSha256(value As String) As String
            Using sha As SHA256 = SHA256.Create()
                Dim bytes = Encoding.UTF8.GetBytes(value)
                Dim hash = sha.ComputeHash(bytes)
                Dim sb As New StringBuilder(hash.Length * 2)

                For Each b In hash
                    sb.Append(b.ToString("x2", CultureInfo.InvariantCulture))
                Next

                Return sb.ToString()
            End Using
        End Function

        Private Shared Function FirstHex(value As String, length As Integer) As String
            If String.IsNullOrEmpty(value) Then
                Return ""
            End If

            If value.Length <= length Then
                Return value
            End If

            Return value.Substring(0, length)
        End Function

        Private Shared Function GetLicenseCounterRetryDelayMs(attempt As Integer) As Integer
            Select Case attempt
                Case 1
                    Return 50
                Case 2
                    Return 100
                Case 3
                    Return 250
                Case 4
                    Return 500
                Case Else
                    Return 1000
            End Select
        End Function

        Private Shared Function ExtractMonthFromKnownMarkerFileName(fileName As String) As String
            If String.IsNullOrWhiteSpace(fileName) Then
                Return ""
            End If

            Dim match = Regex.Match(fileName, "^RI_LC_(?<m>\d{4}-\d{2})_", RegexOptions.IgnoreCase)
            If match.Success Then
                Return match.Groups("m").Value
            End If

            Return ""
        End Function

        ' Minimal local reporting diagnostics. The active collector does not print
        ' access logs by default (Waitress), so support must not rely on console
        ' output to determine whether the add-in actually attempted reporting.
        ' Never write configured URLs/tokens, licence keys, user identities or payloads.
        Private Shared ReadOnly LicenseCounterDiagnosticLock As New System.Object()
        Private Shared Sub LogLicenseCounterDiagnostic(eventType As System.String, detail As System.String)
            Try
                Dim baseFolder As System.String = GetLicenseCounterBaseFolder()
                System.IO.Directory.CreateDirectory(baseFolder)
                Dim diagnosticPath As System.String = System.IO.Path.Combine(baseFolder, "RI_LC_Diagnostics.log")
                Dim cleanType As System.String = System.Text.RegularExpressions.Regex.Replace(If(eventType, "unknown"), "[^a-zA-Z0-9_-]", "_")
                ' This helper is called exclusively with fixed, deliberately sanitized
                ' metadata; never pass user-controlled error messages into it.
                Dim line As System.String = System.DateTime.UtcNow.ToString("o", System.Globalization.CultureInfo.InvariantCulture) & " " & cleanType & " " & If(detail, "") & System.Environment.NewLine
                SyncLock LicenseCounterDiagnosticLock
                    If System.IO.File.Exists(diagnosticPath) AndAlso New System.IO.FileInfo(diagnosticPath).Length > 524288 Then
                        Dim previousPath As System.String = diagnosticPath & ".old"
                        If System.IO.File.Exists(previousPath) Then System.IO.File.Delete(previousPath)
                        System.IO.File.Move(diagnosticPath, previousPath)
                    End If
                    System.IO.File.AppendAllText(diagnosticPath, line, New System.Text.UTF8Encoding(False))
                End SyncLock
            Catch
                ' Diagnostics must never interfere with licencing or the add-in.
            End Try
        End Sub

        Private Shared Sub LogLicenseCounterInternal(eventType As String, details As String)
            Try
                LogLicenseEvent($"Counter:{eventType}", details)
            Catch
            End Try
        End Sub

#End Region

#Region "License Counter Evaluation"

        Private Shared Sub ShowLicenseCounterAnalysisResultDialog(result As LicenseCounterAnalysisResult)
            Using form As New Form()
                form.Text = $"{AN} - License Counter"
                form.StartPosition = FormStartPosition.CenterScreen
                form.Width = 1100
                form.Height = 760
                form.MinimumSize = New Drawing.Size(900, 620)
                form.FormBorderStyle = FormBorderStyle.Sizable
                form.MaximizeBox = True
                form.MinimizeBox = False
                form.ShowInTaskbar = True
                form.TopMost = True
                form.AutoScaleMode = AutoScaleMode.Font
                form.Font = New Drawing.Font("Segoe UI", 9.0F, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point)

                Try
                    Global.SharedLibrary.SharedLibrary.SharedMethods.ApplyStandardDialogIcon(form)
                Catch
                End Try

                Dim txtOutput As New TextBox() With {
                    .Multiline = True,
                    .ScrollBars = ScrollBars.Both,
                    .WordWrap = False,
                    .Dock = DockStyle.Fill,
                    .Font = New Drawing.Font("Consolas", 9.0F, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point),
                    .Text = BuildLicenseCounterAnalysisText(result)
                }

                Dim panel As New FlowLayoutPanel() With {
                    .Dock = DockStyle.Bottom,
                    .AutoSize = True,
                    .FlowDirection = FlowDirection.RightToLeft,
                    .Padding = New Padding(20),
                    .WrapContents = False
                }

                Dim btnClose As New Button() With {.Text = "Close", .AutoSize = True, .Margin = New Padding(10)}
                Dim btnExportCsv As New Button() With {.Text = "Export CSV", .AutoSize = True, .Margin = New Padding(10)}
                Dim btnExportJson As New Button() With {.Text = "Export JSON", .AutoSize = True, .Margin = New Padding(10)}

                AddHandler btnClose.Click,
                    Sub()
                        form.Close()
                    End Sub

                AddHandler btnExportCsv.Click,
                    Sub()
                        Try
                            Using dlg As New FolderBrowserDialog()
                                dlg.Description = "Select folder for CSV export"
                                If dlg.ShowDialog(form) <> DialogResult.OK OrElse String.IsNullOrWhiteSpace(dlg.SelectedPath) Then
                                    Return
                                End If

                                ExportLicenseCounterCsv(result, dlg.SelectedPath)
                                ShowCustomMessageBox("CSV export completed.", $"{AN} - License Counter")
                            End Using
                        Catch ex As System.Exception
                            ShowCustomMessageBox($"CSV export failed: {ex.Message}", $"{AN} - License Counter")
                        End Try
                    End Sub

                AddHandler btnExportJson.Click,
                    Sub()
                        Try
                            Using dlg As New SaveFileDialog()
                                dlg.Filter = "JSON|*.json"
                                dlg.FileName = "licensecounter-analysis.json"

                                If dlg.ShowDialog(form) <> DialogResult.OK OrElse String.IsNullOrWhiteSpace(dlg.FileName) Then
                                    Return
                                End If

                                File.WriteAllText(dlg.FileName, JsonConvert.SerializeObject(result, Formatting.Indented), New UTF8Encoding(True))
                                ShowCustomMessageBox("JSON export completed.", $"{AN} - License Counter")
                            End Using
                        Catch ex As System.Exception
                            ShowCustomMessageBox($"JSON export failed: {ex.Message}", $"{AN} - License Counter")
                        End Try
                    End Sub

                Dim pnlHost As New Panel() With {
                    .Dock = DockStyle.Fill,
                    .Padding = New Padding(15, 15, 15, 15)
                }

                pnlHost.Controls.Add(txtOutput)

                panel.Controls.Add(btnClose)
                panel.Controls.Add(btnExportCsv)
                panel.Controls.Add(btnExportJson)
                form.Controls.Add(pnlHost)
                form.Controls.Add(panel)
                AddHandler form.Shown,
                    Sub()
                        form.TopMost = True
                        Global.SharedLibrary.SharedLibrary.SharedMethods.ForceDialogToForeground(form)
                        Global.SharedLibrary.SharedLibrary.SharedMethods.AttachForeignForegroundWatchdog(form)
                    End Sub
                Dim __safeDialogOwner1985 As System.Windows.Forms.IWin32Window = Global.SharedLibrary.SharedLibrary.SharedMethods.ResolveSameThreadDialogOwner()
                If __safeDialogOwner1985 IsNot Nothing Then
                    form.ShowDialog(__safeDialogOwner1985)
                Else
                    form.ShowDialog()
                End If
            End Using
        End Sub

        Private Shared Function AnalyzeLicenseCounterSources(sources As IEnumerable(Of String),
                                                             recursive As Boolean,
                                                             fromMonth As String,
                                                             toMonth As String) As LicenseCounterAnalysisResult
            Dim result As New LicenseCounterAnalysisResult()
            Dim files As New List(Of String)()
            Const ProgressUnitsPerFile As Integer = 1000

            Try
                For Each source In sources
                    If String.IsNullOrWhiteSpace(source) Then
                        Continue For
                    End If

                    If Directory.Exists(source) Then
                        files.AddRange(Directory.GetFiles(source,
                                                          "*.*",
                                                          If(recursive, SearchOption.AllDirectories, SearchOption.TopDirectoryOnly)))
                    ElseIf File.Exists(source) Then
                        files.Add(source)
                    Else
                        result.Warnings.Add($"Source not found: {source}")
                    End If
                Next

                Dim distinctFiles As List(Of String) = files.Distinct(StringComparer.OrdinalIgnoreCase).ToList()

                If distinctFiles.Count = 0 Then
                    SharedLibrary.ProgressBarModule.GlobalProgressMax = 100
                    SharedLibrary.ProgressBarModule.GlobalProgressValue = 100
                    SharedLibrary.ProgressBarModule.GlobalProgressLabel = "No files found."
                Else
                    SharedLibrary.ProgressBarModule.GlobalProgressMax = distinctFiles.Count * ProgressUnitsPerFile
                    SharedLibrary.ProgressBarModule.GlobalProgressValue = 0
                    SharedLibrary.ProgressBarModule.GlobalProgressLabel = $"Analyzing {distinctFiles.Count} file(s)..."
                End If

                Dim fileIndex As Integer = 0

                For Each filePath In distinctFiles
                    If SharedLibrary.ProgressBarModule.CancelOperation Then
                        result.Warnings.Add("Analysis cancelled by user.")
                        Return result
                    End If

                    fileIndex += 1

                    ParseLicenseCounterFile(filePath, result, fileIndex, distinctFiles.Count, ProgressUnitsPerFile)
                Next

                SharedLibrary.ProgressBarModule.GlobalProgressLabel = "Aggregating report data..."
                SharedLibrary.ProgressBarModule.GlobalProgressValue =
                    System.Math.Max(0, SharedLibrary.ProgressBarModule.GlobalProgressMax - 1)

                Dim filtered = result.Events.
                    Where(Function(e) e.Validity <> "invalid").
                    Where(Function(e) IsLicenseCounterMonthInRange(e.Month, fromMonth, toMonth)).
                    ToList()

                BuildLicenseCounterReports(filtered, result)
                SharedLibrary.ProgressBarModule.GlobalProgressValue = SharedLibrary.ProgressBarModule.GlobalProgressMax

            Catch ex As System.Exception
                result.Warnings.Add("Analysis error: " & ex.Message)
            End Try

            Return result
        End Function

        Private Shared Sub ParseLicenseCounterFile(filePath As String,
                                                   result As LicenseCounterAnalysisResult,
                                                   fileIndex As Integer,
                                                   totalFiles As Integer,
                                                   progressUnitsPerFile As Integer)
            Try
                Dim fileName As String = Path.GetFileName(filePath)
                Dim baseProgress As Integer = (fileIndex - 1) * progressUnitsPerFile

                SharedLibrary.ProgressBarModule.GlobalProgressValue = baseProgress
                SharedLibrary.ProgressBarModule.GlobalProgressLabel = $"Processing file {fileIndex} of {totalFiles}: {fileName}"

                Dim extension As String = Path.GetExtension(filePath)
                Dim shouldTryJsonParse As Boolean = extension.Equals(".json", StringComparison.OrdinalIgnoreCase)

                If Not shouldTryJsonParse Then
                    Try
                        Using sr As New StreamReader(filePath, Encoding.UTF8, detectEncodingFromByteOrderMarks:=True)
                            Dim firstNonWhitespace As Integer = sr.Read()

                            While firstNonWhitespace >= 0 AndAlso Char.IsWhiteSpace(ChrW(firstNonWhitespace))
                                firstNonWhitespace = sr.Read()
                            End While

                            shouldTryJsonParse = (firstNonWhitespace = AscW("{"c) OrElse firstNonWhitespace = AscW("["c))
                        End Using
                    Catch
                    End Try
                End If

                If shouldTryJsonParse Then
                    Dim json As String = File.ReadAllText(filePath, Encoding.UTF8)

                    If Path.GetFileName(filePath).Equals(LicenseCounterStateFileName, StringComparison.OrdinalIgnoreCase) Then
                        If TryParseLicenseCounterStateJson(json, filePath, result) Then
                            SharedLibrary.ProgressBarModule.GlobalProgressValue = baseProgress + progressUnitsPerFile
                            Return
                        End If
                    End If

                    If TryParseLicenseCounterJson(json, filePath, result) Then
                        SharedLibrary.ProgressBarModule.GlobalProgressValue = baseProgress + progressUnitsPerFile
                        Return
                    End If
                End If

                If TryParseLicenseCounterFromFileName(Path.GetFileName(filePath), filePath, result) Then
                    SharedLibrary.ProgressBarModule.GlobalProgressValue = baseProgress + progressUnitsPerFile
                    Return
                End If

                Dim fileLength As Long = 0

                Try
                    fileLength = New FileInfo(filePath).Length
                Catch
                    fileLength = 0
                End Try

                Dim bytesReadEstimate As Long = 0
                Dim nextProgressUnits As Integer = 1

                For Each line In File.ReadLines(filePath)
                    If SharedLibrary.ProgressBarModule.CancelOperation Then
                        Exit For
                    End If

                    If line.IndexOf("RI_LC_v1", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                       line.IndexOf("ri_lc_v=1", StringComparison.OrdinalIgnoreCase) >= 0 Then
                        TryParseLicenseCounterLogLine(line, filePath, result)
                    End If

                    If fileLength > 0 Then
                        bytesReadEstimate += Encoding.UTF8.GetByteCount(line) + 2

                        Dim completedUnits As Integer = CInt(System.Math.Truncate((bytesReadEstimate / CDbl(fileLength)) * progressUnitsPerFile))
                        completedUnits = System.Math.Max(0, System.Math.Min(progressUnitsPerFile, completedUnits))

                        If completedUnits >= nextProgressUnits Then
                            SharedLibrary.ProgressBarModule.GlobalProgressValue = baseProgress + completedUnits
                            SharedLibrary.ProgressBarModule.GlobalProgressLabel = $"Processing file {fileIndex} of {totalFiles}: {fileName}"
                            nextProgressUnits = completedUnits + 1
                        End If
                    End If
                Next

                SharedLibrary.ProgressBarModule.GlobalProgressValue = baseProgress + progressUnitsPerFile

            Catch ex As System.Exception
                result.Warnings.Add($"invalid_json: {filePath} ({ex.Message})")
            End Try
        End Sub

        Private Shared Function TryParseLicenseCounterStateJson(json As String,
                                                                source As String,
                                                                result As LicenseCounterAnalysisResult) As Boolean
            Try
                Dim jObj = JObject.Parse(json)
                Dim tokenMap = TryCast(jObj("appliedResubmitTokens"), JObject)

                If tokenMap Is Nothing Then
                    Return True
                End If

                For Each prop In tokenMap.Properties()
                    result.AppliedResubmitTokens.Add(New LicenseCounterAppliedResubmitTokenRow() With {
                        .Token = prop.Name,
                        .AppliedUtc = prop.Value.ToString(),
                        .Source = source
                    })
                Next

                Return True

            Catch ex As System.Exception
                result.Warnings.Add($"invalid_state_json: {source} ({ex.Message})")
                Return True
            End Try
        End Function

        Private Shared Function TryParseLicenseCounterJson(json As String,
                                                           source As String,
                                                           result As LicenseCounterAnalysisResult) As Boolean
            Try
                Dim jObj = JObject.Parse(json)
                Dim eventType As String = jObj.Value(Of String)("t")

                If String.Equals(eventType, "U", StringComparison.OrdinalIgnoreCase) Then
                    Dim parsedNetworkScope As String = GetJsonValue(jObj, "networkScope")
                    Dim parsedAllowedNetworkIds As List(Of String) = GetJsonStringArray(jObj, "allowedNetworkIds")

                    If parsedAllowedNetworkIds.Count = 0 Then
                        parsedAllowedNetworkIds = SplitAllowedNetworkIds(parsedNetworkScope)
                    End If

                    Dim evt As New LicenseCounterAnalysisEvent() With {
                        .EventType = "U",
                        .Month = GetJsonValue(jObj, "month"),
                        .NetworkKey = GetJsonValue(jObj, "networkKey"),
                        .NetworkScope = parsedNetworkScope,
                        .AllowedNetworkIds = parsedAllowedNetworkIds,
                        .ProductId = If(String.IsNullOrWhiteSpace(GetJsonValue(jObj, "productId")),
                                        GetJsonValue(jObj, "productIdSafe"),
                                        GetJsonValue(jObj, "productId")),
                        .ProductIdSafe = GetJsonValue(jObj, "productIdSafe"),
                        .HostId = GetJsonValue(jObj, "hostId"),
                        .LoadedHosts = GetJsonStringArray(jObj, "loadedHosts"),
                        .UserKey = GetJsonValue(jObj, "userKey"),
                        .BillingEventId = GetJsonValue(jObj, "billingEventId"),
                        .HostEventId = GetJsonValue(jObj, "hostEventId"),
                        .Anonymous = If(jObj.Value(Of Boolean?)("anon").GetValueOrDefault(False), "true", "false"),
                        .UserObf = GetJsonValue(jObj, "userObf"),
                        .Source = source,
                        .Validity = "valid"
                    }

                    If String.Equals(evt.Anonymous, "false", StringComparison.OrdinalIgnoreCase) AndAlso Not String.IsNullOrWhiteSpace(evt.UserObf) Then
                        Dim restoredUserId As String = ""
                        If TryDecodeUserObf(evt.UserObf, restoredUserId) Then
                            evt.UserId = restoredUserId
                        Else
                            evt.WarningText = "nonanonymous_event_without_user_data"
                        End If
                    End If

                    If String.IsNullOrWhiteSpace(evt.BillingEventId) AndAlso
                       Not String.IsNullOrWhiteSpace(evt.Month) AndAlso
                       Not String.IsNullOrWhiteSpace(evt.NetworkKey) AndAlso
                       Not String.IsNullOrWhiteSpace(evt.ProductId) AndAlso
                       Not String.IsNullOrWhiteSpace(evt.UserKey) Then
                        evt.BillingEventId = BuildBillingEventId(evt.Month, evt.NetworkKey, evt.ProductId, evt.UserKey)
                    End If

                    If String.IsNullOrWhiteSpace(evt.HostEventId) AndAlso
                       Not String.IsNullOrWhiteSpace(evt.Month) AndAlso
                       Not String.IsNullOrWhiteSpace(evt.NetworkKey) AndAlso
                       Not String.IsNullOrWhiteSpace(evt.ProductId) AndAlso
                       Not String.IsNullOrWhiteSpace(evt.HostId) AndAlso
                       Not String.IsNullOrWhiteSpace(evt.UserKey) Then
                        evt.HostEventId = BuildHostEventId(evt.Month, evt.NetworkKey, evt.ProductId, evt.HostId, evt.UserKey)
                    End If

                    If String.IsNullOrWhiteSpace(evt.Month) OrElse
                       String.IsNullOrWhiteSpace(evt.NetworkKey) OrElse
                       String.IsNullOrWhiteSpace(evt.ProductId) OrElse
                       String.IsNullOrWhiteSpace(evt.UserKey) Then
                        evt.Validity = "warning"
                        evt.WarningText = "incomplete_event"
                    End If

                    result.Events.Add(evt)
                    Return True
                End If

                If String.Equals(eventType, "E", StringComparison.OrdinalIgnoreCase) Then
                    result.Events.Add(New LicenseCounterAnalysisEvent() With {
                        .EventType = "E",
                        .ErrorCode = GetJsonValue(jObj, "errorCode"),
                        .Month = GetJsonValue(jObj, "relatedMonth"),
                        .Source = source,
                        .Validity = "warning",
                        .WarningText = "error_event"
                    })
                    Return True
                End If

                Return False

            Catch
                Return False
            End Try
        End Function

        Private Shared Function TryParseLicenseCounterFromFileName(fileName As String,
                                                                   source As String,
                                                                   result As LicenseCounterAnalysisResult) As Boolean
            Dim reportMatch = Regex.Match(fileName,
                                          "^RI_LC_REPORT_(?<m>\d{4}-\d{2})_(?<n>n[0-9a-f]+)_(?<p>[^_]+)_(?<h>WD|EX|OL|UNKNOWN)_(?<u>u[0-9a-f]+)_(?<e>e[0-9a-f]+)\.json$",
                                          RegexOptions.IgnoreCase)

            If reportMatch.Success Then
                result.Events.Add(New LicenseCounterAnalysisEvent() With {
                    .EventType = "U",
                    .Month = reportMatch.Groups("m").Value,
                    .NetworkKey = reportMatch.Groups("n").Value,
                    .ProductId = reportMatch.Groups("p").Value,
                    .ProductIdSafe = reportMatch.Groups("p").Value,
                    .HostId = reportMatch.Groups("h").Value.ToUpperInvariant(),
                    .UserKey = reportMatch.Groups("u").Value,
                    .HostEventId = reportMatch.Groups("e").Value,
                    .Source = source,
                    .Validity = "warning",
                    .WarningText = "filename_only_event"
                })
                Return True
            End If

            Dim markerMatch = Regex.Match(fileName,
                                          "^RI_LC_(?<m>\d{4}-\d{2})_(?<n>n[0-9a-f]+)_(?<p>[^_]+)_(?<h>WD|EX|OL|UNKNOWN)_(?<u>u[0-9a-f]+)\.json$",
                                          RegexOptions.IgnoreCase)

            If markerMatch.Success Then
                result.Events.Add(New LicenseCounterAnalysisEvent() With {
                    .EventType = "U",
                    .Month = markerMatch.Groups("m").Value,
                    .NetworkKey = markerMatch.Groups("n").Value,
                    .ProductId = markerMatch.Groups("p").Value,
                    .ProductIdSafe = markerMatch.Groups("p").Value,
                    .HostId = markerMatch.Groups("h").Value.ToUpperInvariant(),
                    .UserKey = markerMatch.Groups("u").Value,
                    .Source = source,
                    .Validity = "warning",
                    .WarningText = "filename_only_event"
                })
                Return True
            End If

            Return False
        End Function

        Private Shared Sub TryParseLicenseCounterLogLine(line As String,
                                                         source As String,
                                                         result As LicenseCounterAnalysisResult)
            Dim pathPattern As String =
                "RI_LC_v1/m/(?<m>\d{4}-\d{2})/n/(?<n>n[0-9a-f]+)(?:/ns/(?<ns>[^/\s\?&]+))?/p/(?<p>[^/\s\?&]+)/h/(?<h>WD|EX|OL|UNKNOWN)/u/(?<u>u[0-9a-f]+)/b/(?<b>b[0-9a-f]+)/e/(?<e>e[0-9a-f]+)/a/(?<a>[01])(?:/r/(?<r>[^/\s\?&]+))?(?:/du/(?<du>[^/\s\?&]+))?"
            Dim pathMatch = Regex.Match(line, pathPattern, RegexOptions.IgnoreCase)

            If pathMatch.Success Then
                Dim parsedNetworkScope As String = Uri.UnescapeDataString(pathMatch.Groups("ns").Value)

                Dim evt As New LicenseCounterAnalysisEvent() With {
                    .EventType = "U",
                    .Month = pathMatch.Groups("m").Value,
                    .NetworkKey = pathMatch.Groups("n").Value,
                    .NetworkScope = parsedNetworkScope,
                    .AllowedNetworkIds = SplitAllowedNetworkIds(parsedNetworkScope),
                    .ProductId = Uri.UnescapeDataString(pathMatch.Groups("p").Value),
                    .ProductIdSafe = Uri.UnescapeDataString(pathMatch.Groups("p").Value),
                    .HostId = pathMatch.Groups("h").Value.ToUpperInvariant(),
                    .UserKey = pathMatch.Groups("u").Value,
                    .BillingEventId = pathMatch.Groups("b").Value,
                    .HostEventId = pathMatch.Groups("e").Value,
                    .Anonymous = If(pathMatch.Groups("a").Value = "1", "true", "false"),
                    .UserObf = Uri.UnescapeDataString(pathMatch.Groups("du").Value),
                    .Source = source,
                    .Validity = "valid"
                }

                If String.Equals(evt.Anonymous, "false", StringComparison.OrdinalIgnoreCase) AndAlso Not String.IsNullOrWhiteSpace(evt.UserObf) Then
                    Dim restoredUserId As String = ""
                    If TryDecodeUserObf(evt.UserObf, restoredUserId) Then
                        evt.UserId = restoredUserId
                    End If
                End If

                result.Events.Add(evt)
                Return
            End If

            If line.IndexOf("ri_lc_v=1", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Dim parameters = ParseLicenseCounterQueryParameters(line)

                If parameters.ContainsKey("t") AndAlso String.Equals(parameters("t"), "U", StringComparison.OrdinalIgnoreCase) Then
                    Dim parsedNetworkScope As String = GetDictionaryValue(parameters, "ns")
                    Dim parsedAllowedNetworkIds As List(Of String) = SplitAllowedNetworkIds(GetDictionaryValue(parameters, "ani"))

                    If parsedAllowedNetworkIds.Count = 0 Then
                        parsedAllowedNetworkIds = SplitAllowedNetworkIds(parsedNetworkScope)
                    End If

                    Dim evt As New LicenseCounterAnalysisEvent() With {
                        .EventType = "U",
                        .Month = GetDictionaryValue(parameters, "m"),
                        .NetworkKey = GetDictionaryValue(parameters, "n"),
                        .NetworkScope = parsedNetworkScope,
                        .AllowedNetworkIds = parsedAllowedNetworkIds,
                        .ProductId = GetDictionaryValue(parameters, "p"),
                        .ProductIdSafe = GetDictionaryValue(parameters, "p"),
                        .HostId = NormalizeLicenseCounterHostId(GetDictionaryValue(parameters, "h")),
                        .LoadedHosts = GetDictionaryValue(parameters, "loadedHosts").Split(","c).Where(Function(value) Not System.String.IsNullOrWhiteSpace(value)).ToList(),
                        .UserKey = GetDictionaryValue(parameters, "u"),
                        .BillingEventId = GetDictionaryValue(parameters, "b"),
                        .HostEventId = GetDictionaryValue(parameters, "e"),
                        .Anonymous = If(GetDictionaryValue(parameters, "a") = "1", "true", "false"),
                        .UserObf = GetDictionaryValue(parameters, "du"),
                        .Source = source,
                        .Validity = "valid"
                    }

                    If String.Equals(evt.Anonymous, "false", StringComparison.OrdinalIgnoreCase) AndAlso Not String.IsNullOrWhiteSpace(evt.UserObf) Then
                        Dim restoredUserId As String = ""
                        If TryDecodeUserObf(evt.UserObf, restoredUserId) Then
                            evt.UserId = restoredUserId
                        End If
                    End If

                    result.Events.Add(evt)
                End If
            End If
        End Sub

        Private Shared Function ParseLicenseCounterQueryParameters(line As String) As Dictionary(Of String, String)
            Dim results As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

            Dim queryStart As Integer = line.IndexOf("ri_lc_v=1", StringComparison.OrdinalIgnoreCase)
            If queryStart < 0 Then
                Return results
            End If

            Dim query As String = line.Substring(queryStart)

            For Each match As Match In Regex.Matches(query, "(?<k>[A-Za-z0-9_]+)=(?<v>[^&\s""]+)")
                results(match.Groups("k").Value) = Uri.UnescapeDataString(match.Groups("v").Value)
            Next

            Return results
        End Function

        Private Shared Sub BuildLicenseCounterReports(events As List(Of LicenseCounterAnalysisEvent),
                                                      result As LicenseCounterAnalysisResult)
            Dim mainGroups As New Dictionary(Of String, HashSet(Of String))(StringComparer.OrdinalIgnoreCase)
            Dim hostGroups As New Dictionary(Of String, HashSet(Of String))(StringComparer.OrdinalIgnoreCase)
            Dim userGroups As New Dictionary(Of String, HashSet(Of String))(StringComparer.OrdinalIgnoreCase)
            Dim userMeta As New Dictionary(Of String, LicenseCounterUserRow)(StringComparer.OrdinalIgnoreCase)
            Dim networkScopes As New Dictionary(Of String, HashSet(Of String))(StringComparer.OrdinalIgnoreCase)

            For Each evt In events.Where(Function(e) String.Equals(e.EventType, "U", StringComparison.OrdinalIgnoreCase))
                Dim productIdForGrouping As String = If(String.IsNullOrWhiteSpace(evt.ProductId), evt.ProductIdSafe, evt.ProductId)
                Dim mainKey As String = $"{evt.Month}|{evt.NetworkKey}|{productIdForGrouping}"
                Dim userKey As String = $"{evt.Month}|{evt.NetworkKey}|{productIdForGrouping}|{evt.UserKey}"

                If Not mainGroups.ContainsKey(mainKey) Then
                    mainGroups(mainKey) = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                End If
                mainGroups(mainKey).Add(evt.UserKey)

                Dim observedHosts As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase) From {evt.HostId}
                For Each loadedHost As System.String In evt.LoadedHosts
                    observedHosts.Add(NormalizeLicenseCounterHostId(loadedHost))
                Next
                For Each loadedHost As System.String In observedHosts
                    Dim observedHostKey As System.String = $"{evt.Month}|{evt.NetworkKey}|{productIdForGrouping}|{loadedHost}"
                    If Not hostGroups.ContainsKey(observedHostKey) Then hostGroups(observedHostKey) = New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
                    hostGroups(observedHostKey).Add(evt.UserKey)
                Next

                If Not userGroups.ContainsKey(userKey) Then
                    userGroups(userKey) = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                    userMeta(userKey) = New LicenseCounterUserRow() With {
                        .Month = evt.Month,
                        .NetworkKey = evt.NetworkKey,
                        .ProductId = productIdForGrouping,
                        .UserId = evt.UserId,
                        .UserKey = evt.UserKey,
                        .Hosts = ""
                    }
                End If

                For Each loadedHost As System.String In observedHosts
                    userGroups(userKey).Add(loadedHost)
                Next

                If String.IsNullOrWhiteSpace(userMeta(userKey).UserId) AndAlso Not String.IsNullOrWhiteSpace(evt.UserId) Then
                    userMeta(userKey).UserId = evt.UserId
                End If

                If Not String.IsNullOrWhiteSpace(evt.NetworkKey) Then
                    If Not networkScopes.ContainsKey(evt.NetworkKey) Then
                        networkScopes(evt.NetworkKey) = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                    End If

                    Dim scopeValues As List(Of String) = If(evt.AllowedNetworkIds, New List(Of String)())

                    If scopeValues.Count = 0 AndAlso Not String.IsNullOrWhiteSpace(evt.NetworkScope) Then
                        scopeValues = SplitAllowedNetworkIds(evt.NetworkScope)
                    End If

                    For Each scopeValue In scopeValues
                        networkScopes(evt.NetworkKey).Add(scopeValue)
                    Next
                End If
            Next

            For Each kvp In mainGroups.OrderBy(Function(k) k.Key, StringComparer.OrdinalIgnoreCase)
                Dim parts = kvp.Key.Split("|"c)
                result.MainRows.Add(New LicenseCounterMainRow() With {
                    .Month = parts(0),
                    .NetworkKey = parts(1),
                    .ProductId = parts(2),
                    .DistinctUsers = kvp.Value.Count,
                    .CoveredNetworks = GetCoveredNetworksDisplay(networkScopes, parts(1))
                })
            Next

            For Each kvp In hostGroups.OrderBy(Function(k) k.Key, StringComparer.OrdinalIgnoreCase)
                Dim parts = kvp.Key.Split("|"c)
                result.HostRows.Add(New LicenseCounterHostRow() With {
                    .Month = parts(0),
                    .NetworkKey = parts(1),
                    .ProductId = parts(2),
                    .HostId = parts(3),
                    .DistinctUsers = kvp.Value.Count
                })
            Next

            For Each key In userMeta.Keys.OrderBy(Function(k) k, StringComparer.OrdinalIgnoreCase)
                userMeta(key).Hosts = String.Join(", ", userGroups(key).OrderBy(Function(h) h, StringComparer.OrdinalIgnoreCase))
                result.UserRows.Add(userMeta(key))
            Next

            For Each key In networkScopes.Keys.OrderBy(Function(k) k, StringComparer.OrdinalIgnoreCase)
                Dim coveredNetworks As String = "(not available in source data)"

                If networkScopes(key) IsNot Nothing AndAlso networkScopes(key).Count > 0 Then
                    coveredNetworks = String.Join(", ", networkScopes(key).OrderBy(Function(n) n, StringComparer.OrdinalIgnoreCase))
                End If

                result.NetworkScopeRows.Add(New LicenseCounterNetworkScopeRow() With {
                    .NetworkKey = key,
                    .CoveredNetworks = coveredNetworks
                })
            Next
        End Sub

        Private Shared Function BuildLicenseCounterAnalysisText(result As LicenseCounterAnalysisResult) As String
            Dim sb As New StringBuilder()

            sb.AppendLine("MAIN REPORT")
            sb.AppendLine("Month,NetworkKey,ProductID,DistinctUsers,CoveredNetworks")
            For Each row In result.MainRows
                sb.AppendLine($"{row.Month},{row.NetworkKey},{row.ProductId},{row.DistinctUsers},{row.CoveredNetworks}")
            Next

            sb.AppendLine()
            sb.AppendLine("HOST REPORT")
            sb.AppendLine("Month,NetworkKey,ProductID,HostID,DistinctUsers")
            For Each row In result.HostRows
                sb.AppendLine($"{row.Month},{row.NetworkKey},{row.ProductId},{row.HostId},{row.DistinctUsers}")
            Next

            If result.UserRows.Count > 0 Then
                sb.AppendLine()
                sb.AppendLine("USER DETAIL REPORT")
                sb.AppendLine("Month,NetworkKey,ProductID,UserID,UserKey,Hosts")
                For Each row In result.UserRows
                    sb.AppendLine($"{row.Month},{row.NetworkKey},{row.ProductId},{row.UserId},{row.UserKey},{row.Hosts}")
                Next
            End If

            If result.NetworkScopeRows.Count > 0 Then
                sb.AppendLine()
                sb.AppendLine("NETWORKKEY COVERAGE")
                sb.AppendLine("NetworkKey,CoveredNetworks")
                For Each row In result.NetworkScopeRows
                    sb.AppendLine($"{row.NetworkKey},{row.CoveredNetworks}")
                Next
            End If

            If result.AppliedResubmitTokens.Count > 0 Then
                sb.AppendLine()
                sb.AppendLine("APPLIED RESUBMIT TOKENS")
                sb.AppendLine("Token,AppliedUtc,Source")
                For Each row In result.AppliedResubmitTokens.
                    OrderBy(Function(r) r.AppliedUtc, StringComparer.OrdinalIgnoreCase).
                    ThenBy(Function(r) r.Token, StringComparer.OrdinalIgnoreCase)
                    sb.AppendLine($"{row.Token},{row.AppliedUtc},{row.Source}")
                Next
            End If

            If result.Warnings.Count > 0 Then
                sb.AppendLine()
                sb.AppendLine("WARNINGS")
                For Each warning In result.Warnings
                    sb.AppendLine("- " & warning)
                Next
            End If

            Return sb.ToString()
        End Function

        Private Shared Sub ExportLicenseCounterCsv(result As LicenseCounterAnalysisResult, targetFolder As String)
            Directory.CreateDirectory(targetFolder)

            Dim stamp As String = System.DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture)

            Dim mainCsv As New StringBuilder()
            mainCsv.AppendLine("Month,NetworkKey,ProductID,DistinctUsers,CoveredNetworks")
            For Each row In result.MainRows
                mainCsv.AppendLine(ToCsvRow(row.Month,
                                            row.NetworkKey,
                                            row.ProductId,
                                            row.DistinctUsers.ToString(CultureInfo.InvariantCulture),
                                            row.CoveredNetworks))
            Next
            File.WriteAllText(Path.Combine(targetFolder, $"licensecounter_main_{stamp}.csv"), mainCsv.ToString(), New UTF8Encoding(True))

            Dim hostCsv As New StringBuilder()
            hostCsv.AppendLine("Month,NetworkKey,ProductID,HostID,DistinctUsers")
            For Each row In result.HostRows
                hostCsv.AppendLine(ToCsvRow(row.Month, row.NetworkKey, row.ProductId, row.HostId, row.DistinctUsers.ToString(CultureInfo.InvariantCulture)))
            Next
            File.WriteAllText(Path.Combine(targetFolder, $"licensecounter_host_{stamp}.csv"), hostCsv.ToString(), New UTF8Encoding(True))

            If result.UserRows.Count > 0 Then
                Dim userCsv As New StringBuilder()
                userCsv.AppendLine("Month,NetworkKey,ProductID,UserID,UserKey,Hosts")
                For Each row In result.UserRows
                    userCsv.AppendLine(ToCsvRow(row.Month, row.NetworkKey, row.ProductId, row.UserId, row.UserKey, row.Hosts))
                Next
                File.WriteAllText(Path.Combine(targetFolder, $"licensecounter_user_{stamp}.csv"), userCsv.ToString(), New UTF8Encoding(True))
            End If

            If result.NetworkScopeRows.Count > 0 Then
                Dim scopeCsv As New StringBuilder()
                scopeCsv.AppendLine("NetworkKey,CoveredNetworks")
                For Each row In result.NetworkScopeRows
                    scopeCsv.AppendLine(ToCsvRow(row.NetworkKey, row.CoveredNetworks))
                Next
                File.WriteAllText(Path.Combine(targetFolder, $"licensecounter_networkscope_{stamp}.csv"), scopeCsv.ToString(), New UTF8Encoding(True))
            End If

            If result.AppliedResubmitTokens.Count > 0 Then
                Dim tokenCsv As New StringBuilder()
                tokenCsv.AppendLine("Token,AppliedUtc,Source")
                For Each row In result.AppliedResubmitTokens.
                    OrderBy(Function(r) r.AppliedUtc, StringComparer.OrdinalIgnoreCase).
                    ThenBy(Function(r) r.Token, StringComparer.OrdinalIgnoreCase)
                    tokenCsv.AppendLine(ToCsvRow(row.Token, row.AppliedUtc, row.Source))
                Next
                File.WriteAllText(Path.Combine(targetFolder, $"licensecounter_resubmittokens_{stamp}.csv"), tokenCsv.ToString(), New UTF8Encoding(True))
            End If
        End Sub

        Private Shared Function ToCsvRow(ParamArray values As String()) As String
            Return String.Join(",", values.Select(Function(v) $"""{If(v, "").Replace("""", """""")}"""))
        End Function

        Private Shared Function IsLicenseCounterMonthInRange(monthValue As String, fromMonth As String, toMonth As String) As Boolean
            If String.IsNullOrWhiteSpace(monthValue) Then
                Return False
            End If

            If Not String.IsNullOrWhiteSpace(fromMonth) AndAlso String.Compare(monthValue, fromMonth, StringComparison.OrdinalIgnoreCase) < 0 Then
                Return False
            End If

            If Not String.IsNullOrWhiteSpace(toMonth) AndAlso String.Compare(monthValue, toMonth, StringComparison.OrdinalIgnoreCase) > 0 Then
                Return False
            End If

            Return True
        End Function

        Private Shared Function GetJsonValue(jObj As JObject, propertyName As String) As String
            Dim token = jObj(propertyName)
            If token Is Nothing Then
                Return ""
            End If

            Return token.ToString()
        End Function

        Private Shared Function GetJsonStringArray(jObj As JObject, propertyName As String) As List(Of String)
            Dim token = jObj(propertyName)
            If token Is Nothing OrElse token.Type <> JTokenType.Array Then
                Return New List(Of String)()
            End If

            Return token.Values(Of String)().
                Where(Function(v) Not String.IsNullOrWhiteSpace(v)).
                Select(Function(v) v.Trim()).
                Distinct(StringComparer.OrdinalIgnoreCase).
                OrderBy(Function(v) v, StringComparer.OrdinalIgnoreCase).
                ToList()
        End Function

        Private Shared Function SplitAllowedNetworkIds(value As String) As List(Of String)
            If String.IsNullOrWhiteSpace(value) Then
                Return New List(Of String)()
            End If

            Return value.Split({"|"c}, StringSplitOptions.RemoveEmptyEntries).
                Select(Function(v) NormalizeLicenseCounterValue(v)).
                Where(Function(v) Not String.IsNullOrWhiteSpace(v)).
                Distinct(StringComparer.OrdinalIgnoreCase).
                OrderBy(Function(v) v, StringComparer.OrdinalIgnoreCase).
                ToList()
        End Function

        Private Shared Function GetCoveredNetworksDisplay(networkScopes As Dictionary(Of String, HashSet(Of String)),
                                                          networkKey As String) As String
            If String.IsNullOrWhiteSpace(networkKey) Then
                Return "(not available in source data)"
            End If

            If networkScopes Is Nothing OrElse Not networkScopes.ContainsKey(networkKey) Then
                Return "(not available in source data)"
            End If

            If networkScopes(networkKey) Is Nothing OrElse networkScopes(networkKey).Count = 0 Then
                Return "(not available in source data)"
            End If

            Return String.Join(", ", networkScopes(networkKey).OrderBy(Function(n) n, StringComparer.OrdinalIgnoreCase))
        End Function

        Private Shared Function GetDictionaryValue(dict As Dictionary(Of String, String), key As String) As String
            If dict.ContainsKey(key) Then
                Return dict(key)
            End If

            Return ""
        End Function

#End Region

    End Class
End Namespace
