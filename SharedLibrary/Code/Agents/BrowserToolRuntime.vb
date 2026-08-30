' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: BrowserToolRuntime.vb
' Purpose: Implements the host-agnostic Playwright browser session used by the
'          shared browser_open, browser_snapshot and browser_interact tools.
'
' Architecture / How it works:
'  - Owns one Playwright browser/context/page per host process and serializes all
'    access through a SemaphoreSlim.
'  - Uses Playwright's AI ARIA snapshot mode, introduced in Playwright 1.59. The
'    resulting YAML contains native refs such as [ref=e7] or frame-qualified refs such as [ref=f11e38] and includes iframe
'    content.
'  - Tracks the refs present in the latest successful snapshot. browser_interact
'    rejects unknown or stale refs before attempting an action.
'  - Resolves native AI snapshot refs through Playwright's aria-ref selector.
'    A successful non-navigating form action may retain the snapshot for a short
'    same-form sequence; potentially navigational actions invalidate it.
'  - Keeps browser logic independent from Outlook/Word COM. Both Office hosts can
'    therefore share the same tool implementation and registry wiring.
'
' Default browser strategy:
'  - First tries the installed Microsoft Edge stable channel (`msedge`) because
'    Red Ink runs on Windows Office hosts.
'  - Falls back to Playwright's bundled Chromium when configured to do so.
'  - To use bundled Chromium, install the matching Playwright browser binary for
'    the deployed Microsoft.Playwright package version.
'
' Security / behavior notes:
'  - Only absolute HTTP and HTTPS navigation is accepted by browser_open.
'  - Explicit navigation to loopback/private/link-local targets requires a user
'    approval that is kept only in memory for the current Office host process.
'  - Direct link clicks are preflighted through the same session-only policy;
'    public Internet navigation keeps the existing behavior without extra prompts.
'  - A browser context is non-persistent by default; cookies/session state live
'    for the lifetime of the current host process only.
'  - browser_interact is write-capable and may submit forms or mutate remote data.
'  - Browser sessions are headless by default. After navigation, the runtime makes a
'    conservative best-effort attempt to dismiss common cookie/consent overlays using
'    only reject/necessary-only actions; it never auto-clicks accept-all.
' =============================================================================

Option Explicit On
Option Strict On
Option Infer On

Namespace Agents

    Public NotInheritable Class BrowserToolOptions
        Public Property Headless As System.Boolean = True
        Public Property BrowserChannel As System.String = "msedge"
        Public Property FallbackToBundledChromium As System.Boolean = True
        Public Property IgnoreHTTPSErrors As System.Boolean = False
        Public Property DefaultTimeoutMs As System.Int32 = 30000
        Public Property NavigationTimeoutMs As System.Int32 = 45000
        Public Property RuntimePath As System.String = System.String.Empty
        Public Property UseLocalRuntimeCache As System.Boolean = False

        Public Function Clone() As BrowserToolOptions
            Return New BrowserToolOptions() With {
                .Headless = Headless,
                .BrowserChannel = BrowserChannel,
                .FallbackToBundledChromium = FallbackToBundledChromium,
                .IgnoreHTTPSErrors = IgnoreHTTPSErrors,
                .DefaultTimeoutMs = DefaultTimeoutMs,
                .NavigationTimeoutMs = NavigationTimeoutMs,
                .RuntimePath = RuntimePath,
                .UseLocalRuntimeCache = UseLocalRuntimeCache
            }
        End Function
    End Class

    Friend NotInheritable Class BrowserToolRuntime
        Private Enum RuntimeResolutionPreference
            Auto = 0
            SourceOnly = 1
        End Enum

        Private Shared ReadOnly Gate As New System.Threading.SemaphoreSlim(1, 1)
        Private Shared ReadOnly ConfigurationLock As New System.Object()
        Private Shared ReadOnly PrivateNetworkApprovalLock As New System.Object()
        Private Shared ReadOnly SessionApprovedPrivateOrigins As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
        Private Shared ReadOnly SnapshotRefRegex As New System.Text.RegularExpressions.Regex(
            "\[ref=([A-Za-z0-9_-]+)\]",
            System.Text.RegularExpressions.RegexOptions.Compiled Or System.Text.RegularExpressions.RegexOptions.CultureInvariant)

        Private Shared ConfiguredOptions As New BrowserToolOptions()
        Private Shared PlaywrightInstance As Microsoft.Playwright.IPlaywright
        Private Shared Browser As Microsoft.Playwright.IBrowser
        Private Shared BrowserContext As Microsoft.Playwright.IBrowserContext
        Private Shared CurrentPage As Microsoft.Playwright.IPage
        Private Shared LoadedAuthProfileKey As System.String = System.String.Empty
        Private Shared SnapshotIsValid As System.Boolean
        Private Shared SnapshotGeneration As System.Int64
        Private Shared LastSnapshotRefs As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
        Private Shared LastResolvedRuntimeRoot As System.String = System.String.Empty
        Private Shared LastResolvedRuntimeVersion As System.String = System.String.Empty
        Private Shared LastResolvedRuntimeUsesLocalCache As System.Boolean
        Private Const PlaywrightDriverUnavailableCode As System.String = "PLAYWRIGHT_DRIVER_UNAVAILABLE"
        Private Const PlaywrightBrowserUnavailableCode As System.String = "PLAYWRIGHT_BROWSER_UNAVAILABLE"
        Private Const PlaywrightDriverSearchPathEnvironmentVariable As System.String = "PLAYWRIGHT_DRIVER_SEARCH_PATH"
        Private Const PlaywrightBrowsersPathEnvironmentVariable As System.String = "PLAYWRIGHT_BROWSERS_PATH"

        ' Interactive authentication must sometimes keep the exact headed Chromium/Edge
        ' process alive because some enterprise sites bind authentication to that live
        ' browser session and require a fresh login in a new process/context. After the
        ' user confirms sign-in, Red Ink therefore hides only the browser window while
        ' retaining the same Playwright browser/context/page for automation.
        Private Const BrowserWindowHideCommand As System.Int32 = 0 ' SW_HIDE

        Private Delegate Function BrowserWindowEnumProc(
            hWnd As System.IntPtr,
            lParam As System.IntPtr
        ) As System.Boolean

        <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)>
        Private Shared Function EnumWindows(
            lpEnumFunc As BrowserWindowEnumProc,
            lParam As System.IntPtr
        ) As System.Boolean
        End Function

        <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)>
        Private Shared Function IsWindowVisible(hWnd As System.IntPtr) As System.Boolean
        End Function

        <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)>
        Private Shared Function IsWindow(hWnd As System.IntPtr) As System.Boolean
        End Function

        <System.Runtime.InteropServices.DllImport("user32.dll", CharSet:=System.Runtime.InteropServices.CharSet.Unicode, SetLastError:=True)>
        Private Shared Function GetClassName(
            hWnd As System.IntPtr,
            className As System.Text.StringBuilder,
            maxCount As System.Int32
        ) As System.Int32
        End Function

        <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)>
        Private Shared Function ShowWindow(
            hWnd As System.IntPtr,
            command As System.Int32
        ) As System.Boolean
        End Function

        Private Sub New()
        End Sub

        Public Shared Sub Configure(options As BrowserToolOptions)
            If options Is Nothing Then
                Throw New System.ArgumentNullException(NameOf(options))
            End If
            ValidateOptions(options)

            SyncLock ConfigurationLock
                ConfiguredOptions = options.Clone()
            End SyncLock
        End Sub

        Public Shared Sub ConfigureExternalRuntime(runtimePath As System.String, useLocalCache As System.Boolean)
            SyncLock ConfigurationLock
                Dim updated As BrowserToolOptions = ConfiguredOptions.Clone()
                updated.RuntimePath = If(runtimePath, System.String.Empty)
                updated.UseLocalRuntimeCache = useLocalCache
                ConfiguredOptions = updated
            End SyncLock
        End Sub

        Public Shared Async Function ExecuteAsync(
            toolName As System.String,
            arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object),
            cancellationToken As System.Threading.CancellationToken
        ) As System.Threading.Tasks.Task(Of System.String)
            Dim firstResult As System.String = Await ExecuteOnceAsync(
                toolName,
                arguments,
                cancellationToken,
                RuntimeResolutionPreference.Auto).ConfigureAwait(False)

            If Not IsPlaywrightLifecycleFailurePayload(firstResult) Then
                Return firstResult
            End If

            Dim firstPhase As System.String = ExtractErrorPhaseFromPayload(firstResult)
            Dim firstError As System.String = ExtractErrorMessageFromPayload(firstResult)
            Dim firstRuntimeRoot As System.String = System.String.Empty
            Dim firstRuntimeVersion As System.String = System.String.Empty
            Dim firstUsedLocalCache As System.Boolean = False
            GetLastRuntimeSelection(firstRuntimeRoot, firstRuntimeVersion, firstUsedLocalCache)

            Global.SharedLibrary.SharedLibrary.UpdateHandler.WriteUpdateLog(
                "[BrowserRuntime] lifecycle failure detected; tool=" & CompactLogValue(toolName) &
                "; phase=" & CompactLogValue(firstPhase) &
                "; runtime=" & CompactLogValue(firstRuntimeRoot) &
                "; version=" & CompactLogValue(firstRuntimeVersion) &
                "; localCache=" & firstUsedLocalCache.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                "; error=" & CompactLogValue(firstError) &
                "; action=reset")

            Await ResetRuntimeAfterLifecycleFailureAsync().ConfigureAwait(False)

            If System.String.Equals(toolName, BrowserTools.BrowserOpenToolName, System.StringComparison.OrdinalIgnoreCase) Then
                ' Never replay a visible authentication flow after the user-confirmation phase.
                ' At that point credentials, MFA or other remote state may already have changed.
                If System.String.Equals(firstPhase, "interactive_user_confirmation", System.StringComparison.OrdinalIgnoreCase) OrElse
                   System.String.Equals(firstPhase, "interactive_storage_state", System.StringComparison.OrdinalIgnoreCase) Then

                    Return CreateErrorPayload(
                        BrowserTools.BrowserOpenToolName,
                        "PLAYWRIGHT_RUNTIME_FAILED",
                        "The Playwright process became unusable after interactive authentication had already reached the user-confirmation phase. The browser session was reset and was not replayed automatically.",
                        False,
                        True,
                        firstPhase)
                End If

                ' A private/intranet navigation can legitimately require a visible browser after
                ' the Red Ink private-network approval has been granted. Do not convert a
                ' navigation-phase failure from authentication=auto into a terminal runtime error:
                ' return the original retryable BROWSER_OPEN_FAILED after resetting the dead
                ' session so the existing one-time interactive-authentication path can run with a
                ' fresh Playwright process.
                If IsPrivateInteractiveFallbackPhase(firstPhase) AndAlso
                   Not firstUsedLocalCache AndAlso
                   System.String.Equals(GetOptionalString(arguments, "authentication", "auto").Trim(), "auto", System.StringComparison.OrdinalIgnoreCase) AndAlso
                   IsPrivateNavigationTarget(arguments) AndAlso
                   AskUserTool.IsInteractive() Then

                    Global.SharedLibrary.SharedLibrary.UpdateHandler.WriteUpdateLog(
                        "[BrowserRuntime] private-network navigation requires fresh interactive retry; tool=" &
                        BrowserTools.BrowserOpenToolName & "; action=return_retryable")
                    Return firstResult
                End If

                ' If a structurally valid local cache nevertheless kills the Playwright
                ' driver/browser process, retry once from the configured source runtime instead
                ' of selecting the same cache again. Only after the source succeeds do we mark
                ' the local cache invalid, avoiding false quarantine for site-specific failures.
                Dim secondPreference As RuntimeResolutionPreference =
                    If(firstUsedLocalCache, RuntimeResolutionPreference.SourceOnly, RuntimeResolutionPreference.Auto)

                Dim secondResult As System.String = Await ExecuteOnceAsync(
                    toolName,
                    arguments,
                    cancellationToken,
                    secondPreference).ConfigureAwait(False)

                If Not IsPlaywrightLifecycleFailurePayload(secondResult) Then
                    If firstUsedLocalCache AndAlso
                       secondPreference = RuntimeResolutionPreference.SourceOnly AndAlso
                       SourceRetryDemonstratesHealthyRuntime(firstPhase, secondResult) Then

                        PlaywrightRuntimeResolver.MarkLocalCacheInvalid(firstRuntimeRoot, firstError)
                        Global.SharedLibrary.SharedLibrary.UpdateHandler.WriteUpdateLog(
                            "[BrowserRuntime] source runtime recovered from local-cache failure; tool=" &
                            BrowserTools.BrowserOpenToolName &
                            "; rejectedCache=" & CompactLogValue(firstRuntimeRoot))
                    End If
                    Return secondResult
                End If

                Dim secondPhase As System.String = ExtractErrorPhaseFromPayload(secondResult)
                Dim secondError As System.String = ExtractErrorMessageFromPayload(secondResult)
                Dim secondRuntimeRoot As System.String = System.String.Empty
                Dim secondRuntimeVersion As System.String = System.String.Empty
                Dim secondUsedLocalCache As System.Boolean = False
                GetLastRuntimeSelection(secondRuntimeRoot, secondRuntimeVersion, secondUsedLocalCache)

                If firstUsedLocalCache AndAlso
                   secondPreference = RuntimeResolutionPreference.SourceOnly AndAlso
                   SourceRetryDemonstratesHealthyRuntime(firstPhase, secondResult) Then

                    PlaywrightRuntimeResolver.MarkLocalCacheInvalid(firstRuntimeRoot, firstError)
                    Global.SharedLibrary.SharedLibrary.UpdateHandler.WriteUpdateLog(
                        "[BrowserRuntime] source runtime progressed beyond local-cache failure; tool=" &
                        BrowserTools.BrowserOpenToolName &
                        "; rejectedCache=" & CompactLogValue(firstRuntimeRoot) &
                        "; sourcePhase=" & CompactLogValue(secondPhase))
                End If

                Await ResetRuntimeAfterLifecycleFailureAsync().ConfigureAwait(False)

                If IsPrivateInteractiveFallbackPhase(secondPhase) AndAlso
                   System.String.Equals(GetOptionalString(arguments, "authentication", "auto").Trim(), "auto", System.StringComparison.OrdinalIgnoreCase) AndAlso
                   IsPrivateNavigationTarget(arguments) AndAlso
                   AskUserTool.IsInteractive() Then

                    Global.SharedLibrary.SharedLibrary.UpdateHandler.WriteUpdateLog(
                        "[BrowserRuntime] private-network navigation requires visible interactive retry after source check; tool=" &
                        BrowserTools.BrowserOpenToolName & "; action=return_retryable")
                    Return secondResult
                End If

                Global.SharedLibrary.SharedLibrary.UpdateHandler.WriteUpdateLog(
                    "[BrowserRuntime] lifecycle restart failed; tool=" & BrowserTools.BrowserOpenToolName &
                    "; phase=" & CompactLogValue(secondPhase) &
                    "; runtime=" & CompactLogValue(secondRuntimeRoot) &
                    "; version=" & CompactLogValue(secondRuntimeVersion) &
                    "; localCache=" & secondUsedLocalCache.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                    "; error=" & CompactLogValue(secondError))

                Return CreateErrorPayload(
                    BrowserTools.BrowserOpenToolName,
                    "PLAYWRIGHT_RUNTIME_FAILED",
                    "The Playwright driver/browser process terminated or its connection became unusable after one controlled restart. " &
                    If(System.String.IsNullOrWhiteSpace(secondPhase), System.String.Empty, "Phase: " & secondPhase & ". ") &
                    secondError,
                    False,
                    False,
                    secondPhase)
            End If

            Return CreateErrorPayload(
                toolName,
                "PLAYWRIGHT_RUNTIME_FAILED",
                "The Playwright driver/browser process terminated or its connection became unusable. The browser session was reset; call browser_open again before continuing.",
                False,
                False,
                firstPhase)
        End Function

        Private Shared Async Function ExecuteOnceAsync(
            toolName As System.String,
            arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object),
            cancellationToken As System.Threading.CancellationToken,
            runtimePreference As RuntimeResolutionPreference
        ) As System.Threading.Tasks.Task(Of System.String)
            If System.String.Equals(toolName, BrowserTools.BrowserOpenToolName, System.StringComparison.OrdinalIgnoreCase) Then
                Return Await OpenAsync(arguments, cancellationToken, runtimePreference).ConfigureAwait(False)
            End If

            If System.String.Equals(toolName, BrowserTools.BrowserSnapshotToolName, System.StringComparison.OrdinalIgnoreCase) Then
                Return Await SnapshotAsync(arguments, cancellationToken).ConfigureAwait(False)
            End If

            If System.String.Equals(toolName, BrowserTools.BrowserInteractToolName, System.StringComparison.OrdinalIgnoreCase) Then
                Return Await InteractAsync(arguments, cancellationToken).ConfigureAwait(False)
            End If

            Return CreateErrorPayload(
                toolName,
                "UNKNOWN_BROWSER_TOOL",
                "Unknown browser tool name.",
                False,
                False)
        End Function

        Private Shared Async Function ResetRuntimeAfterLifecycleFailureAsync() As System.Threading.Tasks.Task
            Await Gate.WaitAsync().ConfigureAwait(False)
            Try
                Await DisposeRuntimeAsync().ConfigureAwait(False)
            Finally
                Gate.Release()
            End Try
        End Function

        Private Shared Function IsPlaywrightLifecycleFailurePayload(payload As System.String) As System.Boolean
            If System.String.IsNullOrWhiteSpace(payload) Then
                Return False
            End If

            Try
                Dim obj As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(payload)
                If Not System.String.Equals(System.Convert.ToString(obj("status"), System.Globalization.CultureInfo.InvariantCulture), "error", System.StringComparison.OrdinalIgnoreCase) Then
                    Return False
                End If

                Dim errorObject As Newtonsoft.Json.Linq.JObject = TryCast(obj("error"), Newtonsoft.Json.Linq.JObject)
                If errorObject Is Nothing Then
                    Return False
                End If

                Dim message As System.String = System.Convert.ToString(errorObject("message"), System.Globalization.CultureInfo.InvariantCulture)
                Return IsPlaywrightLifecycleFailureMessage(message)
            Catch ex As System.Exception
                Return False
            End Try
        End Function

        Private Shared Function IsPlaywrightLifecycleFailureMessage(message As System.String) As System.Boolean
            If System.String.IsNullOrWhiteSpace(message) Then
                Return False
            End If

            Dim lifecycleFragments As System.String() = {
                "CancellationTokenSource has been disposed",
                "Process exited",
                "Target page, context or browser has been closed",
                "Target page, context or browser was closed",
                "Browser has been closed",
                "Browser closed",
                "Connection closed",
                "Connection is closed",
                "Playwright connection closed"
            }

            For Each fragment As System.String In lifecycleFragments
                If message.IndexOf(fragment, System.StringComparison.OrdinalIgnoreCase) >= 0 Then
                    Return True
                End If
            Next

            Return False
        End Function

        Private Shared Function SourceRetryDemonstratesHealthyRuntime(
            firstFailurePhase As System.String,
            sourceRetryPayload As System.String
        ) As System.Boolean
            If IsSuccessfulToolPayload(sourceRetryPayload) Then Return True
            If Not IsRuntimeInitializationPhase(firstFailurePhase) Then Return False

            Dim sourcePhase As System.String = ExtractErrorPhaseFromPayload(sourceRetryPayload)
            Return System.String.Equals(sourcePhase, "navigation", System.StringComparison.OrdinalIgnoreCase) OrElse
                   System.String.Equals(sourcePhase, "post_navigation", System.StringComparison.OrdinalIgnoreCase) OrElse
                   System.String.Equals(sourcePhase, "interactive_navigation", System.StringComparison.OrdinalIgnoreCase) OrElse
                   System.String.Equals(sourcePhase, "interactive_user_confirmation", System.StringComparison.OrdinalIgnoreCase) OrElse
                   System.String.Equals(sourcePhase, "interactive_storage_state", System.StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function IsPrivateInteractiveFallbackPhase(phase As System.String) As System.Boolean
            Return System.String.Equals(phase, "navigation", System.StringComparison.OrdinalIgnoreCase) OrElse
                   System.String.Equals(phase, "post_navigation", System.StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function IsRuntimeInitializationPhase(phase As System.String) As System.Boolean
            Return System.String.Equals(phase, "runtime_initialize", System.StringComparison.OrdinalIgnoreCase) OrElse
                   System.String.Equals(phase, "page_create", System.StringComparison.OrdinalIgnoreCase) OrElse
                   System.String.Equals(phase, "interactive_runtime_initialize", System.StringComparison.OrdinalIgnoreCase) OrElse
                   System.String.Equals(phase, "interactive_context_create", System.StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function IsSuccessfulToolPayload(payload As System.String) As System.Boolean
            If System.String.IsNullOrWhiteSpace(payload) Then Return False

            Try
                Dim obj As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(payload)
                Return System.String.Equals(
                    System.Convert.ToString(obj("status"), System.Globalization.CultureInfo.InvariantCulture),
                    "ok",
                    System.StringComparison.OrdinalIgnoreCase)
            Catch ex As System.Exception
                Return False
            End Try
        End Function

        Private Shared Function CompactLogValue(value As System.String) As System.String
            Return If(value, System.String.Empty).Replace(System.Environment.NewLine, " ").Replace(System.Convert.ToChar(13), " "c).Replace(System.Convert.ToChar(10), " "c).Trim()
        End Function

        Private Shared Function ExtractErrorMessageFromPayload(payload As System.String) As System.String
            If System.String.IsNullOrWhiteSpace(payload) Then
                Return System.String.Empty
            End If

            Try
                Dim obj As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(payload)
                Dim errorObject As Newtonsoft.Json.Linq.JObject = TryCast(obj("error"), Newtonsoft.Json.Linq.JObject)
                If errorObject Is Nothing Then
                    Return System.String.Empty
                End If
                Return System.Convert.ToString(errorObject("message"), System.Globalization.CultureInfo.InvariantCulture)
            Catch ex As System.Exception
                Return System.String.Empty
            End Try
        End Function

        Private Shared Function ExtractErrorPhaseFromPayload(payload As System.String) As System.String
            If System.String.IsNullOrWhiteSpace(payload) Then
                Return System.String.Empty
            End If

            Try
                Dim obj As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(payload)
                Dim errorObject As Newtonsoft.Json.Linq.JObject = TryCast(obj("error"), Newtonsoft.Json.Linq.JObject)
                If errorObject Is Nothing Then
                    Return System.String.Empty
                End If
                Return System.Convert.ToString(errorObject("phase"), System.Globalization.CultureInfo.InvariantCulture)
            Catch ex As System.Exception
                Return System.String.Empty
            End Try
        End Function

        Private Shared Function IsPrivateNavigationTarget(
            arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object)
        ) As System.Boolean
            Try
                Dim rawUrl As System.String = GetRequiredString(arguments, "url")
                Dim parsedUri As System.Uri = Nothing
                Dim privateAddresses As System.Collections.Generic.List(Of System.String) = Nothing
                Return TryClassifyPrivateNetworkTarget(rawUrl, parsedUri, privateAddresses)
            Catch ex As System.Exception
                Return False
            End Try
        End Function

        Private Shared Sub GetLastRuntimeSelection(
            ByRef runtimeRoot As System.String,
            ByRef runtimeVersion As System.String,
            ByRef usesLocalCache As System.Boolean
        )
            SyncLock ConfigurationLock
                runtimeRoot = LastResolvedRuntimeRoot
                runtimeVersion = LastResolvedRuntimeVersion
                usesLocalCache = LastResolvedRuntimeUsesLocalCache
            End SyncLock
        End Sub

        Public Shared Sub Shutdown()
            Gate.Wait()
            Try
                DisposeRuntimeAsync().ConfigureAwait(False).GetAwaiter().GetResult()
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine(ex.ToString())
            Finally
                Gate.Release()
            End Try
        End Sub

        Public Shared Function CreateErrorPayload(
            toolName As System.String,
            code As System.String,
            message As System.String,
            retryable As System.Boolean,
            stateMayHaveChanged As System.Boolean,
            Optional errorPhase As System.String = Nothing
        ) As System.String
            Dim errorObject As New Newtonsoft.Json.Linq.JObject(
                New Newtonsoft.Json.Linq.JProperty("code", code),
                New Newtonsoft.Json.Linq.JProperty("message", message),
                New Newtonsoft.Json.Linq.JProperty("retryable", retryable),
                New Newtonsoft.Json.Linq.JProperty("state_may_have_changed", stateMayHaveChanged))

            If Not System.String.IsNullOrWhiteSpace(errorPhase) Then
                errorObject.Add(New Newtonsoft.Json.Linq.JProperty("phase", errorPhase))
            End If

            Return New Newtonsoft.Json.Linq.JObject(
                New Newtonsoft.Json.Linq.JProperty("status", "error"),
                New Newtonsoft.Json.Linq.JProperty("tool", If(toolName, System.String.Empty)),
                New Newtonsoft.Json.Linq.JProperty("requires_snapshot", stateMayHaveChanged),
                New Newtonsoft.Json.Linq.JProperty("error", errorObject)).ToString(Newtonsoft.Json.Formatting.None)
        End Function

        Private Shared Async Function OpenAsync(
            arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object),
            cancellationToken As System.Threading.CancellationToken,
            runtimePreference As RuntimeResolutionPreference
        ) As System.Threading.Tasks.Task(Of System.String)
            Dim url As System.String = GetRequiredString(arguments, "url")
            Dim validationError As System.String = ValidateNavigationUrl(url)
            If validationError.Length > 0 Then
                Return CreateErrorPayload(
                    BrowserTools.BrowserOpenToolName,
                    "INVALID_URL",
                    validationError,
                    False,
                    False)
            End If

            If Not EnsurePrivateNavigationApproved(url) Then
                Return CreateErrorPayload(
                    BrowserTools.BrowserOpenToolName,
                    "PRIVATE_NETWORK_ACCESS_DENIED",
                    "Access to the local or private network target was not authorized by the user.",
                    False,
                    False)
            End If

            Dim options As BrowserToolOptions = GetOptionsSnapshot()
            Dim failurePhase As System.String = "preflight"
            Dim authenticationMode As System.String = GetOptionalString(arguments, "authentication", "auto").Trim().ToLowerInvariant()
            If authenticationMode <> "auto" AndAlso authenticationMode <> "interactive" AndAlso authenticationMode <> "none" Then
                Return CreateErrorPayload(
                    BrowserTools.BrowserOpenToolName,
                    "INVALID_AUTHENTICATION_MODE",
                    "authentication must be one of auto, interactive or none.",
                    False,
                    False)
            End If

            Dim authProfile As System.String = GetOptionalString(arguments, "auth_profile", Nothing)
            Dim authProfileKey As System.String = BuildAuthProfileKey(url, authProfile)
            If authProfileKey.Length = 0 Then
                Return CreateErrorPayload(
                    BrowserTools.BrowserOpenToolName,
                    "INVALID_AUTH_PROFILE",
                    "auth_profile contains unsupported characters or could not be normalized.",
                    False,
                    False)
            End If

            Dim waitUntilText As System.String = GetOptionalString(arguments, "wait_until", "domcontentloaded")
            Dim waitUntil As Microsoft.Playwright.WaitUntilState
            Dim waitUntilError As System.String = System.String.Empty
            If Not TryParseWaitUntil(waitUntilText, waitUntil, waitUntilError) Then
                Return CreateErrorPayload(
                    BrowserTools.BrowserOpenToolName,
                    "INVALID_WAIT_UNTIL",
                    waitUntilError,
                    False,
                    False)
            End If

            Dim timeoutMs As System.Int32
            Dim timeoutError As System.String = System.String.Empty
            If Not TryGetTimeout(arguments, "timeout_ms", options.NavigationTimeoutMs, timeoutMs, timeoutError) Then
                Return CreateErrorPayload(
                    BrowserTools.BrowserOpenToolName,
                    "INVALID_TIMEOUT",
                    timeoutError,
                    False,
                    False)
            End If

            Await Gate.WaitAsync(cancellationToken).ConfigureAwait(False)
            Try
                If authenticationMode = "interactive" Then
                    If Not AskUserTool.IsInteractive() Then
                        Return CreateAuthenticationRequiredPayload(url, authProfileKey, "Interactive authentication is unavailable in this unattended run. Provision the browser session in Local Chat or Word first, then retry the unattended task.")
                    End If

                    failurePhase = "interactive_authentication"
                    Dim interactiveResult As System.String = Await ProvisionInteractiveAuthenticationAsync(
                        url,
                        authProfileKey,
                        options,
                        waitUntil,
                        timeoutMs,
                        cancellationToken,
                        runtimePreference).ConfigureAwait(False)

                    If interactiveResult IsNot Nothing Then
                        Return interactiveResult
                    End If
                End If

                Dim storageState As System.String = Nothing
                If authenticationMode <> "none" Then
                    storageState = TryLoadProtectedStorageState(authProfileKey)
                    If storageState Is Nothing AndAlso
                       BrowserContext IsNot Nothing AndAlso
                       Not System.String.IsNullOrEmpty(LoadedAuthProfileKey) AndAlso
                       Not System.String.Equals(LoadedAuthProfileKey, authProfileKey, System.StringComparison.Ordinal) Then
                        ' Never let an explicitly different authentication profile inherit the
                        ' in-memory cookies/tokens of a previously loaded profile.
                        Await DisposeRuntimeAsync().ConfigureAwait(False)
                    End If
                ElseIf BrowserContext IsNot Nothing Then
                    ' Explicit opt-out means a genuinely clean context, not merely "do not
                    ' load another stored profile" while retaining cookies from the current one.
                    Await DisposeRuntimeAsync().ConfigureAwait(False)
                End If

                failurePhase = "runtime_initialize"
                Await EnsureRuntimeAsync(
                    options,
                    cancellationToken,
                    storageState,
                    If(storageState Is Nothing, System.String.Empty, authProfileKey),
                    runtimePreference).ConfigureAwait(False)

                If CurrentPage Is Nothing OrElse CurrentPage.IsClosed Then
                    failurePhase = "page_create"
                    CurrentPage = Await BrowserContext.NewPageAsync().ConfigureAwait(False)
                End If

                InvalidateSnapshot()

                Dim gotoOptions As New Microsoft.Playwright.PageGotoOptions() With {
                    .waitUntil = waitUntil,
                    .Timeout = CSng(timeoutMs)
                }

                failurePhase = "navigation"
                Await CurrentPage.GotoAsync(url, gotoOptions).ConfigureAwait(False)
                failurePhase = "post_navigation"
                SelectNewestOpenPage()
                Dim consentDismissedAfterOpen As System.Boolean =
                    Await TryDismissCommonCookieConsentAsync(CurrentPage, cancellationToken).ConfigureAwait(False)

                Dim title As System.String = Await CurrentPage.TitleAsync().ConfigureAwait(False)
                Return New Newtonsoft.Json.Linq.JObject(
                    New Newtonsoft.Json.Linq.JProperty("status", "ok"),
                    New Newtonsoft.Json.Linq.JProperty("tool", BrowserTools.BrowserOpenToolName),
                    New Newtonsoft.Json.Linq.JProperty("url", CurrentPage.Url),
                    New Newtonsoft.Json.Linq.JProperty("title", title),
                    New Newtonsoft.Json.Linq.JProperty("authentication", authenticationMode),
                    New Newtonsoft.Json.Linq.JProperty("session_state_loaded", storageState IsNot Nothing),
                    New Newtonsoft.Json.Linq.JProperty("requires_snapshot", True)).ToString(Newtonsoft.Json.Formatting.None)
            Catch ex As System.OperationCanceledException
                Return CreateErrorPayload(
                    BrowserTools.BrowserOpenToolName,
                    "CANCELLED",
                    "Browser navigation was cancelled.",
                    True,
                    True,
                    failurePhase)
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine(ex.ToString())
                If IsPlaywrightDriverUnavailable(ex) Then
                    Return CreateErrorPayload(
                        BrowserTools.BrowserOpenToolName,
                        PlaywrightDriverUnavailableCode,
                        SanitizeExceptionMessage(ex),
                        False,
                        False,
                        failurePhase)
                End If
                If IsPlaywrightBrowserUnavailable(ex) Then
                    Return CreateErrorPayload(
                        BrowserTools.BrowserOpenToolName,
                        PlaywrightBrowserUnavailableCode,
                        SanitizeExceptionMessage(ex),
                        False,
                        False,
                        failurePhase)
                End If

                Return CreateErrorPayload(
                    BrowserTools.BrowserOpenToolName,
                    "BROWSER_OPEN_FAILED",
                    SanitizeExceptionMessage(ex),
                    True,
                    True,
                    failurePhase)
            Finally
                Gate.Release()
            End Try
        End Function

        Private Shared Async Function SnapshotAsync(
            arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object),
            cancellationToken As System.Threading.CancellationToken
        ) As System.Threading.Tasks.Task(Of System.String)
            Dim options As BrowserToolOptions = GetOptionsSnapshot()
            Dim timeoutMs As System.Int32
            Dim timeoutError As System.String = System.String.Empty
            If Not TryGetTimeout(arguments, "timeout_ms", options.DefaultTimeoutMs, timeoutMs, timeoutError) Then
                Return CreateErrorPayload(
                    BrowserTools.BrowserSnapshotToolName,
                    "INVALID_TIMEOUT",
                    timeoutError,
                    False,
                    False)
            End If

            Await Gate.WaitAsync(cancellationToken).ConfigureAwait(False)
            Try
                If Not HasOpenPage() Then
                    Return CreateErrorPayload(
                        BrowserTools.BrowserSnapshotToolName,
                        "BROWSER_NOT_OPEN",
                        "No browser page is open. Call browser_open first.",
                        False,
                        False)
                End If

                SelectNewestOpenPage()

                Dim snapshotOptions As New Microsoft.Playwright.PageAriaSnapshotOptions() With {
                    .Mode = Microsoft.Playwright.AriaSnapshotMode.Ai,
                    .Timeout = CSng(timeoutMs)
                }

                Dim snapshot As System.String = Await CurrentPage.AriaSnapshotAsync(snapshotOptions).ConfigureAwait(False)
                Dim refs As System.Collections.Generic.HashSet(Of System.String) = ExtractSnapshotRefs(snapshot)
                LastSnapshotRefs = refs
                SnapshotGeneration += 1L
                SnapshotIsValid = True

                Dim title As System.String = Await CurrentPage.TitleAsync().ConfigureAwait(False)
                Return New Newtonsoft.Json.Linq.JObject(
                    New Newtonsoft.Json.Linq.JProperty("status", "ok"),
                    New Newtonsoft.Json.Linq.JProperty("tool", BrowserTools.BrowserSnapshotToolName),
                    New Newtonsoft.Json.Linq.JProperty("url", CurrentPage.Url),
                    New Newtonsoft.Json.Linq.JProperty("title", title),
                    New Newtonsoft.Json.Linq.JProperty("snapshot_generation", SnapshotGeneration),
                    New Newtonsoft.Json.Linq.JProperty("ref_count", refs.Count),
                    New Newtonsoft.Json.Linq.JProperty("snapshot", snapshot),
                    New Newtonsoft.Json.Linq.JProperty("requires_snapshot", False)).ToString(Newtonsoft.Json.Formatting.None)
            Catch ex As System.OperationCanceledException
                InvalidateSnapshot()
                Return CreateErrorPayload(
                    BrowserTools.BrowserSnapshotToolName,
                    "CANCELLED",
                    "Browser snapshot capture was cancelled.",
                    True,
                    False)
            Catch ex As System.Exception
                InvalidateSnapshot()
                System.Diagnostics.Trace.WriteLine(ex.ToString())
                Return CreateErrorPayload(
                    BrowserTools.BrowserSnapshotToolName,
                    "SNAPSHOT_FAILED",
                    SanitizeExceptionMessage(ex),
                    True,
                    False)
            Finally
                Gate.Release()
            End Try
        End Function

        Private Shared Async Function InteractAsync(
            arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object),
            cancellationToken As System.Threading.CancellationToken
        ) As System.Threading.Tasks.Task(Of System.String)
            Dim refValue As System.String = GetRequiredString(arguments, "ref")
            Dim action As System.String = GetRequiredString(arguments, "action").Trim().ToLowerInvariant()
            Dim value As System.String = GetOptionalString(arguments, "value", Nothing)

            If Not System.Text.RegularExpressions.Regex.IsMatch(refValue, "^[A-Za-z0-9_-]+$", System.Text.RegularExpressions.RegexOptions.CultureInvariant) Then
                Return CreateErrorPayload(
                    BrowserTools.BrowserInteractToolName,
                    "INVALID_REF",
                    "The ref must be a native ref from the most recent browser_snapshot, for example e7 or f11e38.",
                    False,
                    False)
            End If

            Dim actionValidationError As System.String = ValidateAction(action, value)
            If actionValidationError.Length > 0 Then
                Return CreateErrorPayload(
                    BrowserTools.BrowserInteractToolName,
                    "INVALID_ACTION_ARGUMENTS",
                    actionValidationError,
                    False,
                    False)
            End If

            Dim options As BrowserToolOptions = GetOptionsSnapshot()
            Dim timeoutMs As System.Int32
            Dim timeoutError As System.String = System.String.Empty
            If Not TryGetTimeout(arguments, "timeout_ms", options.DefaultTimeoutMs, timeoutMs, timeoutError) Then
                Return CreateErrorPayload(
                    BrowserTools.BrowserInteractToolName,
                    "INVALID_TIMEOUT",
                    timeoutError,
                    False,
                    False)
            End If

            Await Gate.WaitAsync(cancellationToken).ConfigureAwait(False)
            Try
                If Not HasOpenPage() Then
                    Return CreateErrorPayload(
                        BrowserTools.BrowserInteractToolName,
                        "BROWSER_NOT_OPEN",
                        "No browser page is open. Call browser_open first.",
                        False,
                        False)
                End If

                If Not SnapshotIsValid Then
                    Return CreateErrorPayload(
                        BrowserTools.BrowserInteractToolName,
                        "SNAPSHOT_REQUIRED",
                        "No valid browser snapshot is active. Call browser_snapshot before browser_interact.",
                        True,
                        False)
                End If

                If Not LastSnapshotRefs.Contains(refValue) Then
                    Return CreateErrorPayload(
                        BrowserTools.BrowserInteractToolName,
                        "STALE_OR_UNKNOWN_REF",
                        "The ref is not present in the most recent browser_snapshot. Take a new snapshot and use a ref from it.",
                        True,
                        False)
                End If

                SelectNewestOpenPage()
                Dim pageBefore As Microsoft.Playwright.IPage = CurrentPage
                Dim urlBefore As System.String = If(CurrentPage Is Nothing, System.String.Empty, CurrentPage.Url)
                Dim snapshotGenerationBefore As System.Int64 = SnapshotGeneration
                Dim directNavigationUrl As System.String = Nothing
                Dim locator As Microsoft.Playwright.ILocator = CurrentPage.Locator("aria-ref=" & refValue)

                If action = "fill" AndAlso Await IsSensitiveAuthenticationInputAsync(locator).ConfigureAwait(False) Then
                    Return CreateErrorPayload(
                        BrowserTools.BrowserInteractToolName,
                        "SENSITIVE_AUTH_INPUT_BLOCKED",
                        "Direct agent entry into password/authentication-secret fields is blocked. Use browser_open with authentication=interactive so the user can enter passwords, MFA codes or other secrets directly in the visible browser.",
                        False,
                        False)
                End If

                If action = "click" OrElse action = "double_click" Then
                    directNavigationUrl = Await TryGetDirectNavigationUrlAsync(locator, CurrentPage.Url).ConfigureAwait(False)
                    If Not System.String.IsNullOrWhiteSpace(directNavigationUrl) AndAlso
                       Not EnsurePrivateNavigationApproved(directNavigationUrl) Then
                        Return CreateErrorPayload(
                            BrowserTools.BrowserInteractToolName,
                            "PRIVATE_NETWORK_ACCESS_DENIED",
                            "Access to the local or private network target was not authorized by the user.",
                            False,
                            False)
                    End If
                End If

                Await ExecuteLocatorActionAsync(locator, action, value, timeoutMs).ConfigureAwait(False)
                SelectNewestOpenPage()
                Dim consentDismissedAfterInteraction As System.Boolean =
                    Await TryDismissCommonCookieConsentAsync(CurrentPage, cancellationToken).ConfigureAwait(False)

                ' A short form-completion sequence may legitimately use several refs from one snapshot
                ' (for example fill a textbox, then click its submit button). Retain that snapshot only
                ' after actions that do not themselves submit/navigate, and only while the exact page and
                ' URL are unchanged. All potentially navigational/mutating actions still force a fresh
                ' snapshot before another browser_interact call.
                Dim actionCanRetainSnapshot As System.Boolean =
                    action = "fill" OrElse
                    action = "clear" OrElse
                    action = "focus" OrElse
                    action = "hover"
                Dim samePage As System.Boolean =
                    pageBefore IsNot Nothing AndAlso
                    CurrentPage IsNot Nothing AndAlso
                    System.Object.ReferenceEquals(pageBefore, CurrentPage)
                Dim sameUrl As System.Boolean =
                    samePage AndAlso
                    System.String.Equals(urlBefore, CurrentPage.Url, System.StringComparison.OrdinalIgnoreCase)
                Dim retainSnapshot As System.Boolean =
                    actionCanRetainSnapshot AndAlso
                    samePage AndAlso
                    sameUrl AndAlso
                    Not consentDismissedAfterInteraction
                Dim navigationObserved As System.Boolean = Not samePage OrElse Not sameUrl
                If Not retainSnapshot Then
                    InvalidateSnapshot()
                End If

                Global.SharedLibrary.SharedLibrary.UpdateHandler.WriteUpdateLog(
                    "[BrowserRuntime] interaction completed; action=" & action &
                    "; ref=" & refValue &
                    "; snapshot_generation=" & snapshotGenerationBefore.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                    "; snapshot_retained=" & retainSnapshot.ToString().ToLowerInvariant() &
                    "; same_page=" & samePage.ToString().ToLowerInvariant() &
                    "; same_url=" & sameUrl.ToString().ToLowerInvariant() &
                    "; navigation_observed=" & navigationObserved.ToString().ToLowerInvariant() &
                    "; target_url_present=" & (Not System.String.IsNullOrWhiteSpace(directNavigationUrl)).ToString().ToLowerInvariant() &
                    "; consent_dismissed=" & consentDismissedAfterInteraction.ToString().ToLowerInvariant())

                Dim title As System.String = Await CurrentPage.TitleAsync().ConfigureAwait(False)
                Return New Newtonsoft.Json.Linq.JObject(
                    New Newtonsoft.Json.Linq.JProperty("status", "ok"),
                    New Newtonsoft.Json.Linq.JProperty("tool", BrowserTools.BrowserInteractToolName),
                    New Newtonsoft.Json.Linq.JProperty("ref", refValue),
                    New Newtonsoft.Json.Linq.JProperty("action", action),
                    New Newtonsoft.Json.Linq.JProperty("url", CurrentPage.Url),
                    New Newtonsoft.Json.Linq.JProperty("title", title),
                    New Newtonsoft.Json.Linq.JProperty("target_url", If(System.String.IsNullOrWhiteSpace(directNavigationUrl), Nothing, directNavigationUrl)),
                    New Newtonsoft.Json.Linq.JProperty("navigation_observed", navigationObserved),
                    New Newtonsoft.Json.Linq.JProperty("snapshot_generation", snapshotGenerationBefore),
                    New Newtonsoft.Json.Linq.JProperty("snapshot_retained", retainSnapshot),
                    New Newtonsoft.Json.Linq.JProperty("requires_snapshot", Not retainSnapshot)).ToString(Newtonsoft.Json.Formatting.None)
            Catch ex As System.OperationCanceledException
                InvalidateSnapshot()
                Return CreateErrorPayload(
                    BrowserTools.BrowserInteractToolName,
                    "CANCELLED",
                    "Browser interaction was cancelled.",
                    True,
                    True)
            Catch ex As System.Exception
                InvalidateSnapshot()
                System.Diagnostics.Trace.WriteLine(ex.ToString())
                Return CreateErrorPayload(
                    BrowserTools.BrowserInteractToolName,
                    "INTERACTION_FAILED",
                    SanitizeExceptionMessage(ex),
                    True,
                    True)
            Finally
                Gate.Release()
            End Try
        End Function

        Private Shared Async Function TryDismissCommonCookieConsentAsync(
            page As Microsoft.Playwright.IPage,
            cancellationToken As System.Threading.CancellationToken
        ) As System.Threading.Tasks.Task(Of System.Boolean)
            If page Is Nothing OrElse page.IsClosed Then
                Return False
            End If

            cancellationToken.ThrowIfCancellationRequested()

            ' Deliberately conservative: only reject / necessary-only wording.
            ' Never auto-click "accept all", because that would grant optional tracking.
            Dim preferredLabels As System.String() = {
                "Nur notwendige",
                "Nur erforderliche",
                "Nur notwendige Cookies",
                "Nur erforderliche Cookies",
                "Nicht notwendige ablehnen",
                "Alle ablehnen",
                "Ablehnen",
                "Reject all",
                "Reject optional",
                "Decline all",
                "Decline",
                "Necessary only",
                "Essential only",
                "Only necessary",
                "Only essential",
                "Tout refuser",
                "Refuser tout",
                "Refuser",
                "Uniquement nécessaires",
                "Solo necessari",
                "Rifiuta tutto",
                "Rifiuta"
            }

            For Each label As System.String In preferredLabels
                cancellationToken.ThrowIfCancellationRequested()

                Try
                    Dim locator As Microsoft.Playwright.ILocator =
                        page.GetByRole(
                            Microsoft.Playwright.AriaRole.Button,
                            New Microsoft.Playwright.PageGetByRoleOptions() With {
                                .Name = label,
                                .Exact = True
                            })

                    Dim count As System.Int32 =
                        Await locator.CountAsync().ConfigureAwait(False)

                    If count <= 0 Then
                        Continue For
                    End If

                    Dim candidate As Microsoft.Playwright.ILocator = locator.First

                    If Not Await candidate.IsVisibleAsync().ConfigureAwait(False) Then
                        Continue For
                    End If

                    Await candidate.ClickAsync(
                        New Microsoft.Playwright.LocatorClickOptions() With {
                            .Timeout = 1500.0F
                        }).ConfigureAwait(False)

                    ' Consent UIs usually disappear immediately; a short delay lets
                    ' the DOM settle without waiting for network-idle trackers.
                    Await page.WaitForTimeoutAsync(150).ConfigureAwait(False)
                    Return True
                Catch ex As System.Exception
                    ' Consent handling is best effort only. A failure must never make
                    ' browser_open/browser_interact fail; the next snapshot can expose
                    ' the banner so the model can handle it explicitly.
                End Try
            Next

            Return False
        End Function

        Private Shared Async Function ExecuteLocatorActionAsync(
            locator As Microsoft.Playwright.ILocator,
            action As System.String,
            value As System.String,
            timeoutMs As System.Int32
        ) As System.Threading.Tasks.Task
            Select Case action
                Case "click"
                    Await locator.ClickAsync(New Microsoft.Playwright.LocatorClickOptions() With {
                        .Timeout = CSng(timeoutMs)
                    }).ConfigureAwait(False)

                Case "double_click"
                    Await locator.DblClickAsync(New Microsoft.Playwright.LocatorDblClickOptions() With {
                        .Timeout = CSng(timeoutMs)
                    }).ConfigureAwait(False)

                Case "fill"
                    Await locator.FillAsync(value, New Microsoft.Playwright.LocatorFillOptions() With {
                        .Timeout = CSng(timeoutMs)
                    }).ConfigureAwait(False)

                Case "clear"
                    Await locator.ClearAsync(New Microsoft.Playwright.LocatorClearOptions() With {
                        .Timeout = CSng(timeoutMs)
                    }).ConfigureAwait(False)

                Case "press"
                    Await locator.PressAsync(value, New Microsoft.Playwright.LocatorPressOptions() With {
                        .Timeout = CSng(timeoutMs)
                    }).ConfigureAwait(False)

                Case "select"
                    Await locator.SelectOptionAsync(
                        New System.String() {value},
                        New Microsoft.Playwright.LocatorSelectOptionOptions() With {
                            .Timeout = CSng(timeoutMs)
                        }).ConfigureAwait(False)

                Case "check"
                    Await locator.CheckAsync(New Microsoft.Playwright.LocatorCheckOptions() With {
                        .Timeout = CSng(timeoutMs)
                    }).ConfigureAwait(False)

                Case "uncheck"
                    Await locator.UncheckAsync(New Microsoft.Playwright.LocatorUncheckOptions() With {
                        .Timeout = CSng(timeoutMs)
                    }).ConfigureAwait(False)

                Case "hover"
                    Await locator.HoverAsync(New Microsoft.Playwright.LocatorHoverOptions() With {
                        .Timeout = CSng(timeoutMs)
                    }).ConfigureAwait(False)

                Case "focus"
                    Await locator.FocusAsync(New Microsoft.Playwright.LocatorFocusOptions() With {
                        .Timeout = CSng(timeoutMs)
                    }).ConfigureAwait(False)

                Case Else
                    Throw New System.InvalidOperationException("Unsupported browser action: " & action)
            End Select
        End Function

        Private Shared Async Function CreatePlaywrightAsync(
            runtimePreference As RuntimeResolutionPreference
        ) As System.Threading.Tasks.Task(Of Microsoft.Playwright.IPlaywright)
            ConfigurePlaywrightRuntime(runtimePreference)
            Return Await Microsoft.Playwright.Playwright.CreateAsync().ConfigureAwait(False)
        End Function

        Private Shared Sub ConfigurePlaywrightRuntime(runtimePreference As RuntimeResolutionPreference)
            Dim options As BrowserToolOptions = GetOptionsSnapshot()
            Dim resolution As PlaywrightRuntimeResolution = Nothing
            Dim errorMessage As System.String = System.String.Empty
            Dim useLocalCache As System.Boolean = options.UseLocalRuntimeCache
            Dim prepareLocalCache As System.Boolean = options.UseLocalRuntimeCache

            If runtimePreference = RuntimeResolutionPreference.SourceOnly Then
                useLocalCache = False
                prepareLocalCache = False
            End If

            SyncLock ConfigurationLock
                LastResolvedRuntimeRoot = System.String.Empty
                LastResolvedRuntimeVersion = System.String.Empty
                LastResolvedRuntimeUsesLocalCache = False
            End SyncLock

            ' Cache preparation is intentionally non-blocking. In Auto mode a valid local cache
            ' is preferred; otherwise the configured source runtime is used immediately while a
            ' cache is prepared in the background. SourceOnly is used only as a controlled
            ' recovery path after a local runtime has demonstrated a driver/browser lifecycle
            ' failure.
            If Not PlaywrightRuntimeResolver.TryResolve(
                options.RuntimePath,
                useLocalCache,
                prepareLocalCache,
                resolution,
                errorMessage) Then

                Throw New System.InvalidOperationException(PlaywrightDriverUnavailableCode & ": " & errorMessage)
            End If

            SyncLock ConfigurationLock
                LastResolvedRuntimeRoot = If(resolution.EffectiveRoot, System.String.Empty)
                LastResolvedRuntimeVersion = If(resolution.RuntimeVersionText, System.String.Empty)
                LastResolvedRuntimeUsesLocalCache = resolution.UsesLocalCache
            End SyncLock

            System.Environment.SetEnvironmentVariable(PlaywrightDriverSearchPathEnvironmentVariable, resolution.EffectiveRoot)
            If Not System.String.IsNullOrWhiteSpace(resolution.BrowsersDirectory) AndAlso System.IO.Directory.Exists(resolution.BrowsersDirectory) Then
                System.Environment.SetEnvironmentVariable(PlaywrightBrowsersPathEnvironmentVariable, resolution.BrowsersDirectory)
            Else
                System.Environment.SetEnvironmentVariable(PlaywrightBrowsersPathEnvironmentVariable, Nothing)
            End If

            If runtimePreference = RuntimeResolutionPreference.SourceOnly Then
                Global.SharedLibrary.SharedLibrary.UpdateHandler.WriteUpdateLog(
                    "[BrowserRuntime] source runtime selected for controlled recovery; runtime=" &
                    CompactLogValue(resolution.EffectiveRoot) &
                    "; version=" & CompactLogValue(resolution.RuntimeVersionText))
            End If

            System.Diagnostics.Trace.WriteLine("Playwright runtime resolved to: " & resolution.EffectiveRoot & "; version=" & resolution.RuntimeVersionText & "; localCache=" & resolution.UsesLocalCache.ToString())
        End Sub

        Private Shared Function IsPlaywrightDriverUnavailable(ex As System.Exception) As System.Boolean
            Dim current As System.Exception = ex
            While current IsNot Nothing
                Dim message As System.String = If(current.Message, System.String.Empty)
                If message.IndexOf(PlaywrightDriverUnavailableCode, System.StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                   message.IndexOf("Driver not found:", System.StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                   message.IndexOf("PLAYWRIGHT_DRIVER_SEARCH_PATH", System.StringComparison.OrdinalIgnoreCase) >= 0 Then
                    Return True
                End If
                current = current.InnerException
            End While
            Return False
        End Function

        Private Shared Function IsPlaywrightBrowserUnavailable(ex As System.Exception) As System.Boolean
            Dim current As System.Exception = ex
            While current IsNot Nothing
                Dim message As System.String = If(current.Message, System.String.Empty)
                If message.IndexOf("Unable to launch a Playwright browser.", System.StringComparison.OrdinalIgnoreCase) >= 0 Then Return True
                current = current.InnerException
            End While
            Return False
        End Function

        Private Shared Function BuildPlaywrightDriverUnavailableMessage() As System.String
            Return "The configured external Playwright runtime is unavailable, incomplete or incompatible. Configure PlayWrightPath to a valid runtime and optionally PlayWrightUseLocalCache=True. Browser authentication retries will not resolve a missing runtime."
        End Function

        Private Shared Async Function EnsureRuntimeAsync(
            options As BrowserToolOptions,
            cancellationToken As System.Threading.CancellationToken,
            Optional storageState As System.String = Nothing,
            Optional authProfileKey As System.String = Nothing,
            Optional runtimePreference As RuntimeResolutionPreference = RuntimeResolutionPreference.Auto
        ) As System.Threading.Tasks.Task
            cancellationToken.ThrowIfCancellationRequested()

            Dim normalizedProfileKey As System.String = If(authProfileKey, System.String.Empty)
            If Browser IsNot Nothing AndAlso Browser.IsConnected AndAlso BrowserContext IsNot Nothing Then
                If storageState Is Nothing OrElse System.String.Equals(LoadedAuthProfileKey, normalizedProfileKey, System.StringComparison.Ordinal) Then
                    Return
                End If
            End If

            Await DisposeRuntimeAsync().ConfigureAwait(False)
            cancellationToken.ThrowIfCancellationRequested()

            PlaywrightInstance = Await CreatePlaywrightAsync(runtimePreference).ConfigureAwait(False)

            Dim launchedBrowser As Microsoft.Playwright.IBrowser = Nothing
            Dim channelFailure As System.Exception = Nothing

            If Not System.String.IsNullOrWhiteSpace(options.BrowserChannel) Then
                Dim channelOptions As New Microsoft.Playwright.BrowserTypeLaunchOptions() With {
                    .Headless = options.Headless,
                    .Channel = options.BrowserChannel
                }

                Try
                    launchedBrowser = Await PlaywrightInstance.Chromium.LaunchAsync(channelOptions).ConfigureAwait(False)
                Catch ex As System.Exception
                    channelFailure = ex
                End Try
            End If

            If launchedBrowser Is Nothing AndAlso
               (System.String.IsNullOrWhiteSpace(options.BrowserChannel) OrElse options.FallbackToBundledChromium) Then
                Dim bundledFailure As System.Exception = Nothing
                Dim bundledOptions As New Microsoft.Playwright.BrowserTypeLaunchOptions() With {
                    .Headless = options.Headless
                }

                Try
                    launchedBrowser = Await PlaywrightInstance.Chromium.LaunchAsync(bundledOptions).ConfigureAwait(False)
                Catch ex As System.Exception
                    bundledFailure = ex
                End Try

                If launchedBrowser Is Nothing Then
                    Dim combinedMessage As System.String = BuildBrowserLaunchFailureMessage(channelFailure, bundledFailure)
                    Throw New System.InvalidOperationException(combinedMessage, bundledFailure)
                End If
            End If

            If launchedBrowser Is Nothing Then
                Dim message As System.String = "Unable to launch configured browser channel '" & options.BrowserChannel & "'."
                If channelFailure IsNot Nothing AndAlso Not System.String.IsNullOrWhiteSpace(channelFailure.Message) Then
                    message &= " " & channelFailure.Message
                End If
                Throw New System.InvalidOperationException(message, channelFailure)
            End If

            Browser = launchedBrowser

            Dim contextOptions As New Microsoft.Playwright.BrowserNewContextOptions() With {
                .IgnoreHTTPSErrors = options.IgnoreHTTPSErrors
            }
            If Not System.String.IsNullOrWhiteSpace(storageState) Then
                contextOptions.StorageState = storageState
            End If
            BrowserContext = Await Browser.NewContextAsync(contextOptions).ConfigureAwait(False)
            LoadedAuthProfileKey = normalizedProfileKey
            BrowserContext.SetDefaultTimeout(CSng(options.DefaultTimeoutMs))
            BrowserContext.SetDefaultNavigationTimeout(CSng(options.NavigationTimeoutMs))
            CurrentPage = Await BrowserContext.NewPageAsync().ConfigureAwait(False)
            InvalidateSnapshot()
        End Function

        Private Shared Async Function DisposeRuntimeAsync() As System.Threading.Tasks.Task
            Dim contextToClose As Microsoft.Playwright.IBrowserContext = BrowserContext
            Dim browserToClose As Microsoft.Playwright.IBrowser = Browser
            Dim playwrightToDispose As Microsoft.Playwright.IPlaywright = PlaywrightInstance

            BrowserContext = Nothing
            Browser = Nothing
            CurrentPage = Nothing
            PlaywrightInstance = Nothing
            InvalidateSnapshot()

            If contextToClose IsNot Nothing Then
                Try
                    Await contextToClose.CloseAsync().ConfigureAwait(False)
                Catch ex As System.Exception
                    System.Diagnostics.Trace.WriteLine(ex.ToString())
                End Try
            End If

            If browserToClose IsNot Nothing Then
                Try
                    Await browserToClose.CloseAsync().ConfigureAwait(False)
                Catch ex As System.Exception
                    System.Diagnostics.Trace.WriteLine(ex.ToString())
                End Try
            End If

            If playwrightToDispose IsNot Nothing Then
                Try
                    playwrightToDispose.Dispose()
                Catch ex As System.Exception
                    System.Diagnostics.Trace.WriteLine(ex.ToString())
                End Try
            End If
        End Function

        Private Shared Function HasOpenPage() As System.Boolean
            Return Browser IsNot Nothing AndAlso
                   Browser.IsConnected AndAlso
                   BrowserContext IsNot Nothing AndAlso
                   CurrentPage IsNot Nothing AndAlso
                   Not CurrentPage.IsClosed
        End Function

        Private Shared Sub SelectNewestOpenPage()
            If BrowserContext Is Nothing Then
                Return
            End If

            Dim pages As System.Collections.Generic.IReadOnlyList(Of Microsoft.Playwright.IPage) = BrowserContext.Pages
            For index As System.Int32 = pages.Count - 1 To 0 Step -1
                If pages(index) IsNot Nothing AndAlso Not pages(index).IsClosed Then
                    CurrentPage = pages(index)
                    Exit For
                End If
            Next
        End Sub

        Private Shared Sub InvalidateSnapshot()
            SnapshotIsValid = False
            LastSnapshotRefs = New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
        End Sub

        Private Shared Function ExtractSnapshotRefs(snapshot As System.String) As System.Collections.Generic.HashSet(Of System.String)
            Dim refs As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
            If System.String.IsNullOrEmpty(snapshot) Then
                Return refs
            End If

            Dim matches As System.Text.RegularExpressions.MatchCollection = SnapshotRefRegex.Matches(snapshot)
            For Each match As System.Text.RegularExpressions.Match In matches
                If match.Success AndAlso match.Groups.Count > 1 Then
                    refs.Add(match.Groups(1).Value)
                End If
            Next
            Return refs
        End Function

        Private Shared Function GetOptionsSnapshot() As BrowserToolOptions
            SyncLock ConfigurationLock
                Return ConfiguredOptions.Clone()
            End SyncLock
        End Function

        Private Shared Sub ValidateOptions(options As BrowserToolOptions)
            If options.DefaultTimeoutMs < 1000 OrElse options.DefaultTimeoutMs > 120000 Then
                Throw New System.ArgumentOutOfRangeException(NameOf(options), "DefaultTimeoutMs must be between 1000 and 120000.")
            End If
            If options.NavigationTimeoutMs < 1000 OrElse options.NavigationTimeoutMs > 120000 Then
                Throw New System.ArgumentOutOfRangeException(NameOf(options), "NavigationTimeoutMs must be between 1000 and 120000.")
            End If
        End Sub

        Private Shared Function ValidateNavigationUrl(url As System.String) As System.String
            If System.String.IsNullOrWhiteSpace(url) Then
                Return "The url argument is required."
            End If

            Dim parsed As System.Uri = Nothing
            If Not System.Uri.TryCreate(url, System.UriKind.Absolute, parsed) Then
                Return "The url must be an absolute URI."
            End If

            If Not System.String.Equals(parsed.Scheme, System.Uri.UriSchemeHttp, System.StringComparison.OrdinalIgnoreCase) AndAlso
               Not System.String.Equals(parsed.Scheme, System.Uri.UriSchemeHttps, System.StringComparison.OrdinalIgnoreCase) Then
                Return "Only http:// and https:// URLs are allowed."
            End If

            Return System.String.Empty
        End Function

        Private Shared Function EnsurePrivateNavigationApproved(rawUrl As System.String) As System.Boolean
            Dim parsedUri As System.Uri = Nothing
            Dim privateAddresses As System.Collections.Generic.List(Of System.String) = Nothing
            If Not TryClassifyPrivateNetworkTarget(rawUrl, parsedUri, privateAddresses) Then
                Return True
            End If

            Dim origin As System.String = parsedUri.GetLeftPart(System.UriPartial.Authority)
            SyncLock PrivateNetworkApprovalLock
                If SessionApprovedPrivateOrigins.Contains(origin) Then
                    Global.SharedLibrary.SharedLibrary.UpdateHandler.WriteUpdateLog(
                        "[BrowserRuntime] private-network approval reused; origin=" & CompactLogValue(origin))
                    Return True
                End If
            End SyncLock

            Dim details As System.String = System.String.Empty
            If privateAddresses IsNot Nothing AndAlso privateAddresses.Count > 0 Then
                details = System.Environment.NewLine & System.Environment.NewLine &
                          "Resolved private/local address: " & System.String.Join(", ", privateAddresses.ToArray())
            End If

            If Not AskUserTool.IsInteractive() Then
                Global.SharedLibrary.SharedLibrary.UpdateHandler.WriteUpdateLog(
                    "[BrowserRuntime] private-network approval denied; origin=" & CompactLogValue(origin) & "; reason=unattended")
                Return False
            End If

            Global.SharedLibrary.SharedLibrary.UpdateHandler.WriteUpdateLog(
                "[BrowserRuntime] private-network approval requested; origin=" & CompactLogValue(origin))

            Dim message As System.String =
                "The Red Ink browser agent wants to open a local or private network address:" &
                System.Environment.NewLine & System.Environment.NewLine &
                parsedUri.AbsoluteUri &
                details &
                System.Environment.NewLine & System.Environment.NewLine &
                "Private network addresses can provide access to services on this computer or your internal network." &
                System.Environment.NewLine & System.Environment.NewLine &
                "Allow this exact origin for the current Office session?"

            Dim choice As System.Int32 = SharedLibrary.SharedMethods.ShowCustomYesNoBox(
                message,
                "Allow for this session",
                "Deny",
                "Red Ink Browser Agent - Private Network Access")

            If choice <> 1 Then
                Global.SharedLibrary.SharedLibrary.UpdateHandler.WriteUpdateLog(
                    "[BrowserRuntime] private-network approval denied; origin=" & CompactLogValue(origin) & "; reason=user")
                Return False
            End If

            SyncLock PrivateNetworkApprovalLock
                SessionApprovedPrivateOrigins.Add(origin)
            End SyncLock
            Global.SharedLibrary.SharedLibrary.UpdateHandler.WriteUpdateLog(
                "[BrowserRuntime] private-network approval granted; origin=" & CompactLogValue(origin))
            Return True
        End Function

        Private Shared Function TryClassifyPrivateNetworkTarget(
            rawUrl As System.String,
            ByRef parsedUri As System.Uri,
            ByRef privateAddresses As System.Collections.Generic.List(Of System.String)
        ) As System.Boolean
            parsedUri = Nothing
            privateAddresses = New System.Collections.Generic.List(Of System.String)()

            If System.String.IsNullOrWhiteSpace(rawUrl) OrElse
               Not System.Uri.TryCreate(rawUrl, System.UriKind.Absolute, parsedUri) Then
                Return False
            End If

            If Not System.String.Equals(parsedUri.Scheme, System.Uri.UriSchemeHttp, System.StringComparison.OrdinalIgnoreCase) AndAlso
               Not System.String.Equals(parsedUri.Scheme, System.Uri.UriSchemeHttps, System.StringComparison.OrdinalIgnoreCase) Then
                Return False
            End If

            If parsedUri.IsLoopback Then
                Return True
            End If

            Dim host As System.String = If(parsedUri.DnsSafeHost, System.String.Empty).Trim()
            If host.Length = 0 Then
                Return False
            End If

            Dim hostLower As System.String = host.ToLowerInvariant()
            Dim sensitiveName As System.Boolean =
                hostLower = "localhost" OrElse
                hostLower.EndsWith(".localhost", System.StringComparison.OrdinalIgnoreCase) OrElse
                hostLower.EndsWith(".local", System.StringComparison.OrdinalIgnoreCase) OrElse
                hostLower.EndsWith(".internal", System.StringComparison.OrdinalIgnoreCase) OrElse
                hostLower.EndsWith(".home", System.StringComparison.OrdinalIgnoreCase)

            Dim literalAddress As System.Net.IPAddress = Nothing
            If System.Net.IPAddress.TryParse(host, literalAddress) Then
                If IsPrivateOrLocalAddress(literalAddress) Then
                    privateAddresses.Add(literalAddress.ToString())
                    Return True
                End If
                Return sensitiveName
            End If

            Try
                Dim seen As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
                For Each resolvedAddress As System.Net.IPAddress In System.Net.Dns.GetHostAddresses(host)
                    If IsPrivateOrLocalAddress(resolvedAddress) Then
                        Dim addressText As System.String = resolvedAddress.ToString()
                        If seen.Add(addressText) Then
                            privateAddresses.Add(addressText)
                        End If
                    End If
                Next
            Catch ex As System.Net.Sockets.SocketException
                ' Preserve existing navigation behavior if DNS cannot be resolved during preflight.
                ' The browser will perform its normal navigation attempt and surface reachability errors.
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine(ex.ToString())
            End Try

            Return sensitiveName OrElse privateAddresses.Count > 0
        End Function

        Private Shared Function IsPrivateOrLocalAddress(address As System.Net.IPAddress) As System.Boolean
            If address Is Nothing Then Return True

            If address.IsIPv4MappedToIPv6 Then
                address = address.MapToIPv4()
            End If

            If System.Net.IPAddress.IsLoopback(address) Then Return True

            Dim bytes As System.Byte() = address.GetAddressBytes()
            If address.AddressFamily = System.Net.Sockets.AddressFamily.InterNetwork Then
                If bytes.Length <> 4 Then Return True
                If bytes(0) = 0 Then Return True
                If bytes(0) = 10 Then Return True
                If bytes(0) = 100 AndAlso bytes(1) >= 64 AndAlso bytes(1) <= 127 Then Return True
                If bytes(0) = 127 Then Return True
                If bytes(0) = 169 AndAlso bytes(1) = 254 Then Return True
                If bytes(0) = 172 AndAlso bytes(1) >= 16 AndAlso bytes(1) <= 31 Then Return True
                If bytes(0) = 192 AndAlso bytes(1) = 168 Then Return True
                Return False
            End If

            If address.AddressFamily = System.Net.Sockets.AddressFamily.InterNetworkV6 Then
                If address.Equals(System.Net.IPAddress.IPv6Any) OrElse
                   address.Equals(System.Net.IPAddress.IPv6None) OrElse
                   address.IsIPv6LinkLocal OrElse
                   address.IsIPv6SiteLocal Then
                    Return True
                End If

                If bytes.Length = 16 AndAlso (bytes(0) And &HFE) = &HFC Then
                    Return True
                End If
                Return False
            End If

            Return True
        End Function

        Private Shared Async Function TryGetDirectNavigationUrlAsync(
            locator As Microsoft.Playwright.ILocator,
            currentPageUrl As System.String
        ) As System.Threading.Tasks.Task(Of System.String)
            Try
                Dim href As System.String = Await locator.GetAttributeAsync("href").ConfigureAwait(False)
                If System.String.IsNullOrWhiteSpace(href) Then
                    Return Nothing
                End If

                Dim directUri As System.Uri = Nothing
                If System.Uri.TryCreate(href, System.UriKind.Absolute, directUri) Then
                    Return directUri.AbsoluteUri
                End If

                Dim baseUri As System.Uri = Nothing
                If System.Uri.TryCreate(currentPageUrl, System.UriKind.Absolute, baseUri) AndAlso
                   System.Uri.TryCreate(baseUri, href, directUri) Then
                    Return directUri.AbsoluteUri
                End If
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine(ex.ToString())
            End Try
            Return Nothing
        End Function

        Private Shared Async Function IsSensitiveAuthenticationInputAsync(
            locator As Microsoft.Playwright.ILocator
        ) As System.Threading.Tasks.Task(Of System.Boolean)
            Try
                Dim inputType As System.String = Await locator.GetAttributeAsync("type").ConfigureAwait(False)
                If System.String.Equals(inputType, "password", System.StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If

                Dim autocompleteValue As System.String = Await locator.GetAttributeAsync("autocomplete").ConfigureAwait(False)
                If Not System.String.IsNullOrWhiteSpace(autocompleteValue) Then
                    Dim normalizedAutocomplete As System.String = autocompleteValue.Trim().ToLowerInvariant()
                    If normalizedAutocomplete.Contains("current-password") OrElse
                       normalizedAutocomplete.Contains("new-password") OrElse
                       normalizedAutocomplete.Contains("one-time-code") Then
                        Return True
                    End If
                End If
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine(ex.ToString())
            End Try
            Return False
        End Function

        Private Shared Function BuildAuthProfileKey(url As System.String, requestedProfile As System.String) As System.String
            If Not System.String.IsNullOrWhiteSpace(requestedProfile) Then
                Dim trimmed As System.String = requestedProfile.Trim()
                If trimmed.Length > 200 Then Return System.String.Empty
                If Not System.Text.RegularExpressions.Regex.IsMatch(trimmed, "^[A-Za-z0-9][A-Za-z0-9._:@/+-]{0,199}$", System.Text.RegularExpressions.RegexOptions.CultureInvariant) Then
                    Return System.String.Empty
                End If
                Return "named:" & trimmed
            End If

            Dim uri As System.Uri = Nothing
            If Not System.Uri.TryCreate(url, System.UriKind.Absolute, uri) Then
                Return System.String.Empty
            End If
            Return "origin:" & uri.GetLeftPart(System.UriPartial.Authority).ToLowerInvariant()
        End Function

        Private Shared Function CreateAuthenticationRequiredPayload(
            url As System.String,
            authProfileKey As System.String,
            message As System.String
        ) As System.String
            Return New Newtonsoft.Json.Linq.JObject(
                New Newtonsoft.Json.Linq.JProperty("status", "error"),
                New Newtonsoft.Json.Linq.JProperty("tool", BrowserTools.BrowserOpenToolName),
                New Newtonsoft.Json.Linq.JProperty("requires_snapshot", False),
                New Newtonsoft.Json.Linq.JProperty("authentication_required", True),
                New Newtonsoft.Json.Linq.JProperty("url", url),
                New Newtonsoft.Json.Linq.JProperty("auth_profile", authProfileKey),
                New Newtonsoft.Json.Linq.JProperty("error", New Newtonsoft.Json.Linq.JObject(
                    New Newtonsoft.Json.Linq.JProperty("code", "AUTHENTICATION_REQUIRED"),
                    New Newtonsoft.Json.Linq.JProperty("message", message),
                    New Newtonsoft.Json.Linq.JProperty("retryable", False),
                    New Newtonsoft.Json.Linq.JProperty("state_may_have_changed", False)))).ToString(Newtonsoft.Json.Formatting.None)
        End Function

        Private Shared Async Function ProvisionInteractiveAuthenticationAsync(
            url As System.String,
            authProfileKey As System.String,
            options As BrowserToolOptions,
            waitUntil As Microsoft.Playwright.WaitUntilState,
            timeoutMs As System.Int32,
            cancellationToken As System.Threading.CancellationToken,
            runtimePreference As RuntimeResolutionPreference
        ) As System.Threading.Tasks.Task(Of System.String)
            cancellationToken.ThrowIfCancellationRequested()

            Dim authPlaywright As Microsoft.Playwright.IPlaywright = Nothing
            Dim authBrowser As Microsoft.Playwright.IBrowser = Nothing
            Dim authContext As Microsoft.Playwright.IBrowserContext = Nothing
            Dim result As System.String = Nothing
            Dim cancellationException As System.OperationCanceledException = Nothing
            Dim failurePhase As System.String = "interactive_runtime_initialize"
            Dim browserWindowsBefore As System.Collections.Generic.HashSet(Of System.IntPtr) = CaptureChromiumTopLevelWindows()
            Dim interactiveBrowserWindows As New System.Collections.Generic.HashSet(Of System.IntPtr)()
            Try
                authPlaywright = Await CreatePlaywrightAsync(runtimePreference).ConfigureAwait(False)
                authBrowser = Await LaunchBrowserAsync(authPlaywright, options, False).ConfigureAwait(False)

                Dim seedState As System.String = TryLoadProtectedStorageState(authProfileKey)
                If seedState Is Nothing AndAlso
                   BrowserContext IsNot Nothing AndAlso
                   (System.String.IsNullOrEmpty(LoadedAuthProfileKey) OrElse
                    System.String.Equals(LoadedAuthProfileKey, authProfileKey, System.StringComparison.Ordinal)) Then
                    Try
                        seedState = Await BrowserContext.StorageStateAsync(New Microsoft.Playwright.BrowserContextStorageStateOptions() With {
                            .IndexedDB = True
                        }).ConfigureAwait(False)
                    Catch ex As System.Exception
                        System.Diagnostics.Trace.WriteLine(ex.ToString())
                    End Try
                End If

                Dim contextOptions As New Microsoft.Playwright.BrowserNewContextOptions() With {
                    .IgnoreHTTPSErrors = options.IgnoreHTTPSErrors
                }
                If Not System.String.IsNullOrWhiteSpace(seedState) Then
                    contextOptions.StorageState = seedState
                End If

                failurePhase = "interactive_context_create"
                authContext = Await authBrowser.NewContextAsync(contextOptions).ConfigureAwait(False)
                authContext.SetDefaultTimeout(CSng(options.DefaultTimeoutMs))
                authContext.SetDefaultNavigationTimeout(CSng(options.NavigationTimeoutMs))
                Dim authPage As Microsoft.Playwright.IPage = Await authContext.NewPageAsync().ConfigureAwait(False)

                ' A visible interactive authentication session must remain usable even when the
                ' initial Playwright navigation itself cannot complete. Browser-native auth/SSO,
                ' enterprise network permission UI, redirects waiting for user interaction, or a
                ' login challenge can make GotoAsync time out before a DOM snapshot is possible.
                ' In that case the navigation has still been initiated in the visible browser and
                ' the user must be allowed to complete the challenge there.
                Dim initialNavigationWarning As System.String = Nothing
                failurePhase = "interactive_navigation"
                Try
                    Await authPage.GotoAsync(url, New Microsoft.Playwright.PageGotoOptions() With {
                        .WaitUntil = waitUntil,
                        .Timeout = CSng(timeoutMs)
                    }).ConfigureAwait(False)
                Catch ex As System.OperationCanceledException
                    Throw
                Catch ex As System.Exception
                    If IsPlaywrightLifecycleFailureMessage(SanitizeExceptionMessage(ex)) Then
                        Throw
                    End If
                    initialNavigationWarning = SanitizeExceptionMessage(ex)
                    System.Diagnostics.Trace.WriteLine("Interactive authentication initial navigation did not complete: " & ex.ToString())
                End Try

                ' Capture the top-level Chromium/Edge window created by this dedicated
                ' interactive launch. The baseline/delta approach avoids touching browser
                ' windows that were already open before Red Ink started authentication.
                MergeWindowHandles(interactiveBrowserWindows, FindNewChromiumTopLevelWindows(browserWindowsBefore))

                failurePhase = "interactive_user_confirmation"
                Dim promptArguments As New System.Collections.Generic.Dictionary(Of System.String, System.Object)(System.StringComparer.OrdinalIgnoreCase)
                Dim authenticationQuestion As System.String =
                    "Bitte melden Sie sich im geöffneten Browser vollständig an. Benutzername, Passwort, MFA-/Einmalcodes und andere Geheimnisse geben Sie ausschließlich dort ein. Wählen Sie anschließend hier 'Anmeldung abgeschlossen'."
                If Not System.String.IsNullOrWhiteSpace(initialNavigationWarning) Then
                    authenticationQuestion &= System.Environment.NewLine & System.Environment.NewLine &
                        "Hinweis: Die automatische Navigation konnte nicht vollständig abgeschlossen werden. Der sichtbare Browser bleibt absichtlich geöffnet, damit Sie dort eine Browser-, Netzwerk-, SSO- oder Anmeldeabfrage abschließen können."
                End If
                promptArguments.Add("question", authenticationQuestion)

                Dim promptOptions As New Newtonsoft.Json.Linq.JArray()
                promptOptions.Add(New Newtonsoft.Json.Linq.JObject(
                    New Newtonsoft.Json.Linq.JProperty("id", "complete"),
                    New Newtonsoft.Json.Linq.JProperty("label", "Anmeldung abgeschlossen")))
                promptOptions.Add(New Newtonsoft.Json.Linq.JObject(
                    New Newtonsoft.Json.Linq.JProperty("id", "cancel"),
                    New Newtonsoft.Json.Linq.JProperty("label", "Abbrechen")))

                promptArguments.Add("options", promptOptions)
                promptArguments.Add("allow_free_text", False)
                promptArguments.Add("multi_select", False)
                promptArguments.Add("input_type", "choice")

                Global.SharedLibrary.SharedLibrary.UpdateHandler.WriteUpdateLog(
                    "[BrowserRuntime] interactive authentication waiting for user; profile=" & CompactLogValue(authProfileKey) &
                    "; requested_url=" & CompactLogValue(url) &
                    "; page_count=" & authContext.Pages.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                    "; active_url=" & CompactLogValue(authPage.Url))

                Dim promptResultText As System.String = AskUserTool.Execute(promptArguments)
                Dim promptResult As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(promptResultText)
                Dim selected As Newtonsoft.Json.Linq.JArray = TryCast(promptResult("selected_option_ids"), Newtonsoft.Json.Linq.JArray)
                Dim completed As System.Boolean = False
                If selected IsNot Nothing Then
                    For Each token As Newtonsoft.Json.Linq.JToken In selected
                        If System.String.Equals(System.Convert.ToString(token, System.Globalization.CultureInfo.InvariantCulture), "complete", System.StringComparison.OrdinalIgnoreCase) Then
                            completed = True
                            Exit For
                        End If
                    Next
                End If

                If Not completed Then
                    result = CreateAuthenticationRequiredPayload(url, authProfileKey, "Interactive authentication was cancelled or not confirmed by the user.")
                Else
                    failurePhase = "interactive_context_handoff"
                    Dim authenticatedPage As Microsoft.Playwright.IPage = SelectNewestOpenPage(authContext, authPage)
                    Global.SharedLibrary.SharedLibrary.UpdateHandler.WriteUpdateLog(
                        "[BrowserRuntime] interactive authentication user confirmed; profile=" & CompactLogValue(authProfileKey) &
                        "; page_count=" & authContext.Pages.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                        "; continuation_url=" & CompactLogValue(If(authenticatedPage Is Nothing, System.String.Empty, authenticatedPage.Url)))

                    If authenticatedPage Is Nothing OrElse authenticatedPage.IsClosed Then
                        result = CreateAuthenticationRequiredPayload(
                            url,
                            authProfileKey,
                            "Interactive authentication was confirmed, but the visible browser no longer has an open page. Please retry the sign-in.")
                    Else
                        failurePhase = "interactive_storage_state"
                        Dim storageState As System.String = Await authContext.StorageStateAsync(New Microsoft.Playwright.BrowserContextStorageStateOptions() With {
                            .IndexedDB = True
                        }).ConfigureAwait(False)
                        SaveProtectedStorageState(authProfileKey, storageState)

                        ' IMPORTANT: keep the exact browser context in which the user completed
                        ' authentication. Rebuilding a separate context from StorageState can lose
                        ' session-only or application-specific state and can therefore navigate the
                        ' automation back to the login page even though the visible browser already
                        ' reached the authenticated application. StorageState is persisted for future
                        ' runs, but the current run continues in the verified user-controlled context.
                        failurePhase = "interactive_context_adoption"
                        Await DisposeRuntimeAsync().ConfigureAwait(False)
                        PlaywrightInstance = authPlaywright
                        Browser = authBrowser
                        BrowserContext = authContext
                        CurrentPage = authenticatedPage
                        LoadedAuthProfileKey = authProfileKey
                        InvalidateSnapshot()

                        authPlaywright = Nothing
                        authBrowser = Nothing
                        authContext = Nothing

                        Dim title As System.String = Await CurrentPage.TitleAsync().ConfigureAwait(False)

                        ' Some enterprise portals require a new sign-in whenever a new browser
                        ' process/context is created. Closing the headed browser and recreating a
                        ' headless context would therefore destroy the just-established session.
                        ' Keep the exact live authenticated context, but hide the dedicated Red Ink
                        ' browser window after the user confirms authentication.
                        MergeWindowHandles(interactiveBrowserWindows, FindNewChromiumTopLevelWindows(browserWindowsBefore))
                        Dim hiddenWindowCount As System.Int32 = HideBrowserWindows(interactiveBrowserWindows)
                        Dim interactiveWindowHidden As System.Boolean =
                            interactiveBrowserWindows.Count > 0 AndAlso
                            hiddenWindowCount = interactiveBrowserWindows.Count

                        Global.SharedLibrary.SharedLibrary.UpdateHandler.WriteUpdateLog(
                            "[BrowserRuntime] interactive authentication context adopted; profile=" & CompactLogValue(authProfileKey) &
                            "; url=" & CompactLogValue(CurrentPage.Url) &
                            "; action=continue_in_live_context" &
                            "; browser_windows_detected=" & interactiveBrowserWindows.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                            "; browser_windows_hidden=" & hiddenWindowCount.ToString(System.Globalization.CultureInfo.InvariantCulture))

                        result = New Newtonsoft.Json.Linq.JObject(
                            New Newtonsoft.Json.Linq.JProperty("status", "ok"),
                            New Newtonsoft.Json.Linq.JProperty("tool", BrowserTools.BrowserOpenToolName),
                            New Newtonsoft.Json.Linq.JProperty("url", CurrentPage.Url),
                            New Newtonsoft.Json.Linq.JProperty("title", title),
                            New Newtonsoft.Json.Linq.JProperty("authentication", "interactive"),
                            New Newtonsoft.Json.Linq.JProperty("session_state_loaded", True),
                            New Newtonsoft.Json.Linq.JProperty("session_state_was_seeded", Not System.String.IsNullOrWhiteSpace(seedState)),
                            New Newtonsoft.Json.Linq.JProperty("session_state_persisted", True),
                            New Newtonsoft.Json.Linq.JProperty("interactive_authentication_confirmed", True),
                            New Newtonsoft.Json.Linq.JProperty("live_authenticated_context_adopted", True),
                            New Newtonsoft.Json.Linq.JProperty("interactive_browser_window_hidden", interactiveWindowHidden),
                            New Newtonsoft.Json.Linq.JProperty("requires_snapshot", True)).ToString(Newtonsoft.Json.Formatting.None)
                    End If
                End If
            Catch ex As System.OperationCanceledException
                cancellationException = ex
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine(ex.ToString())
                If IsPlaywrightDriverUnavailable(ex) Then
                    result = CreateErrorPayload(
                        BrowserTools.BrowserOpenToolName,
                        PlaywrightDriverUnavailableCode,
                        SanitizeExceptionMessage(ex),
                        False,
                        False,
                        failurePhase)
                ElseIf IsPlaywrightBrowserUnavailable(ex) Then
                    result = CreateErrorPayload(
                        BrowserTools.BrowserOpenToolName,
                        PlaywrightBrowserUnavailableCode,
                        SanitizeExceptionMessage(ex),
                        False,
                        False,
                        failurePhase)
                Else
                    result = CreateErrorPayload(
                        BrowserTools.BrowserOpenToolName,
                        "INTERACTIVE_AUTHENTICATION_FAILED",
                        SanitizeExceptionMessage(ex),
                        True,
                        False,
                        failurePhase)
                End If
            End Try

            Await DisposeInteractiveAuthenticationResourcesAsync(authContext, authBrowser, authPlaywright).ConfigureAwait(False)

            If cancellationException IsNot Nothing Then
                System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(cancellationException).Throw()
            End If

            Return result
        End Function


        Private Shared Function SelectNewestOpenPage(
            context As Microsoft.Playwright.IBrowserContext,
            fallbackPage As Microsoft.Playwright.IPage
        ) As Microsoft.Playwright.IPage
            If context Is Nothing Then Return fallbackPage

            Dim pages As System.Collections.Generic.IReadOnlyList(Of Microsoft.Playwright.IPage) = context.Pages
            If pages IsNot Nothing Then
                For index As System.Int32 = pages.Count - 1 To 0 Step -1
                    Dim candidate As Microsoft.Playwright.IPage = pages(index)
                    If candidate IsNot Nothing AndAlso Not candidate.IsClosed Then
                        Return candidate
                    End If
                Next
            End If

            Return fallbackPage
        End Function

        Private Shared Function CaptureChromiumTopLevelWindows() As System.Collections.Generic.HashSet(Of System.IntPtr)
            Dim result As New System.Collections.Generic.HashSet(Of System.IntPtr)()

            Try
                Dim callback As BrowserWindowEnumProc =
                    Function(hWnd As System.IntPtr, lParam As System.IntPtr) As System.Boolean
                        If hWnd = System.IntPtr.Zero Then
                            Return True
                        End If

                        Dim className As New System.Text.StringBuilder(128)
                        If GetClassName(hWnd, className, className.Capacity) <= 0 Then
                            Return True
                        End If

                        If className.ToString().StartsWith("Chrome_WidgetWin_", System.StringComparison.OrdinalIgnoreCase) Then
                            result.Add(hWnd)
                        End If

                        Return True
                    End Function

                EnumWindows(callback, System.IntPtr.Zero)
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine(ex.ToString())
            End Try

            Return result
        End Function

        Private Shared Function FindNewChromiumTopLevelWindows(
            baseline As System.Collections.Generic.HashSet(Of System.IntPtr)
        ) As System.Collections.Generic.HashSet(Of System.IntPtr)
            Dim current As System.Collections.Generic.HashSet(Of System.IntPtr) = CaptureChromiumTopLevelWindows()
            If baseline Is Nothing OrElse baseline.Count = 0 Then
                Return current
            End If

            current.ExceptWith(baseline)
            Return current
        End Function

        Private Shared Sub MergeWindowHandles(
            target As System.Collections.Generic.HashSet(Of System.IntPtr),
            source As System.Collections.Generic.IEnumerable(Of System.IntPtr)
        )
            If target Is Nothing OrElse source Is Nothing Then
                Return
            End If

            For Each handle As System.IntPtr In source
                If handle <> System.IntPtr.Zero Then
                    target.Add(handle)
                End If
            Next
        End Sub

        Private Shared Function HideBrowserWindows(
            windowHandles As System.Collections.Generic.IEnumerable(Of System.IntPtr)
        ) As System.Int32
            If windowHandles Is Nothing Then
                Return 0
            End If

            Dim hiddenCount As System.Int32 = 0
            For Each handle As System.IntPtr In windowHandles
                Try
                    If handle = System.IntPtr.Zero OrElse Not IsWindow(handle) Then
                        Continue For
                    End If

                    If IsWindowVisible(handle) Then
                        ShowWindow(handle, BrowserWindowHideCommand)
                    End If

                    If Not IsWindowVisible(handle) Then
                        hiddenCount += 1
                    End If
                Catch ex As System.Exception
                    System.Diagnostics.Trace.WriteLine(ex.ToString())
                End Try
            Next

            Return hiddenCount
        End Function

        Private Shared Async Function DisposeInteractiveAuthenticationResourcesAsync(
            authContext As Microsoft.Playwright.IBrowserContext,
            authBrowser As Microsoft.Playwright.IBrowser,
            authPlaywright As Microsoft.Playwright.IPlaywright
        ) As System.Threading.Tasks.Task
            If authContext IsNot Nothing Then
                Try
                    Await authContext.CloseAsync().ConfigureAwait(False)
                Catch ex As System.Exception
                    System.Diagnostics.Trace.WriteLine(ex.ToString())
                End Try
            End If

            If authBrowser IsNot Nothing Then
                Try
                    Await authBrowser.CloseAsync().ConfigureAwait(False)
                Catch ex As System.Exception
                    System.Diagnostics.Trace.WriteLine(ex.ToString())
                End Try
            End If

            If authPlaywright IsNot Nothing Then
                Try
                    authPlaywright.Dispose()
                Catch ex As System.Exception
                    System.Diagnostics.Trace.WriteLine(ex.ToString())
                End Try
            End If
        End Function

        Private Shared Async Function LaunchBrowserAsync(
            playwright As Microsoft.Playwright.IPlaywright,
            options As BrowserToolOptions,
            headless As System.Boolean
        ) As System.Threading.Tasks.Task(Of Microsoft.Playwright.IBrowser)
            Dim channelFailure As System.Exception = Nothing
            If Not System.String.IsNullOrWhiteSpace(options.BrowserChannel) Then
                Try
                    Return Await playwright.Chromium.LaunchAsync(New Microsoft.Playwright.BrowserTypeLaunchOptions() With {
                        .Headless = headless,
                        .Channel = options.BrowserChannel
                    }).ConfigureAwait(False)
                Catch ex As System.Exception
                    channelFailure = ex
                End Try
            End If

            If System.String.IsNullOrWhiteSpace(options.BrowserChannel) OrElse options.FallbackToBundledChromium Then
                Try
                    Return Await playwright.Chromium.LaunchAsync(New Microsoft.Playwright.BrowserTypeLaunchOptions() With {
                        .Headless = headless
                    }).ConfigureAwait(False)
                Catch ex As System.Exception
                    Throw New System.InvalidOperationException(BuildBrowserLaunchFailureMessage(channelFailure, ex), ex)
                End Try
            End If

            Throw New System.InvalidOperationException(BuildBrowserLaunchFailureMessage(channelFailure, Nothing), channelFailure)
        End Function

        Private Shared Function GetAuthStateDirectory() As System.String
            ' Keep browser authentication state beside the existing M365 MSAL cache,
            ' but in separate browser-specific files. This intentionally mirrors the
            ' shared RedInk roaming cache location without coupling BrowserTools to M365.
            Return System.IO.Path.Combine(
                System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData),
                "RedInk")
        End Function

        Private Shared Function GetAuthStatePath(authProfileKey As System.String) As System.String
            Dim bytes As System.Byte() = System.Text.Encoding.UTF8.GetBytes(authProfileKey)
            Using sha As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
                Dim hash As System.Byte() = sha.ComputeHash(bytes)
                Dim fileName As System.String = "browserauth-" & System.BitConverter.ToString(hash).Replace("-", System.String.Empty).ToLowerInvariant() & ".state"
                Return System.IO.Path.Combine(GetAuthStateDirectory(), fileName)
            End Using
        End Function

        Private Shared Function GetAuthEntropy() As System.Byte()
            Return System.Text.Encoding.UTF8.GetBytes("Red Ink BrowserTools authenticated session state v1")
        End Function

        Private Shared Sub SaveProtectedStorageState(authProfileKey As System.String, storageState As System.String)
            If System.String.IsNullOrWhiteSpace(storageState) Then
                Throw New System.InvalidOperationException("Playwright returned an empty authenticated storage state.")
            End If

            Dim directoryPath As System.String = GetAuthStateDirectory()
            System.IO.Directory.CreateDirectory(directoryPath)
            Dim plainBytes As System.Byte() = System.Text.Encoding.UTF8.GetBytes(storageState)
            Dim protectedBytes As System.Byte() = System.Security.Cryptography.ProtectedData.Protect(
                plainBytes,
                GetAuthEntropy(),
                System.Security.Cryptography.DataProtectionScope.CurrentUser)
            Dim targetPath As System.String = GetAuthStatePath(authProfileKey)
            Dim tempPath As System.String = targetPath & ".tmp-" & System.Guid.NewGuid().ToString("N")
            System.IO.File.WriteAllBytes(tempPath, protectedBytes)
            If System.IO.File.Exists(targetPath) Then
                System.IO.File.Delete(targetPath)
            End If
            System.IO.File.Move(tempPath, targetPath)
        End Sub

        Private Shared Function TryLoadProtectedStorageState(authProfileKey As System.String) As System.String
            Try
                Dim path As System.String = GetAuthStatePath(authProfileKey)
                If Not System.IO.File.Exists(path) Then
                    Return Nothing
                End If
                Dim protectedBytes As System.Byte() = System.IO.File.ReadAllBytes(path)
                Dim plainBytes As System.Byte() = System.Security.Cryptography.ProtectedData.Unprotect(
                    protectedBytes,
                    GetAuthEntropy(),
                    System.Security.Cryptography.DataProtectionScope.CurrentUser)
                Dim state As System.String = System.Text.Encoding.UTF8.GetString(plainBytes)
                If System.String.IsNullOrWhiteSpace(state) Then Return Nothing
                Return state
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine("Browser auth state could not be loaded: " & ex.Message)
                Return Nothing
            End Try
        End Function

        Private Shared Function ValidateAction(action As System.String, value As System.String) As System.String
            Select Case action
                Case "click", "double_click", "clear", "check", "uncheck", "hover", "focus"
                    If value IsNot Nothing Then
                        Return "The value argument must be omitted for action '" & action & "'."
                    End If
                    Return System.String.Empty

                Case "fill", "press", "select"
                    If value Is Nothing Then
                        Return "The value argument is required for action '" & action & "'."
                    End If
                    Return System.String.Empty

                Case Else
                    Return "Unsupported action. Use click, double_click, fill, clear, press, select, check, uncheck, hover or focus."
            End Select
        End Function

        Private Shared Function TryParseWaitUntil(
            value As System.String,
            ByRef result As Microsoft.Playwright.WaitUntilState,
            ByRef errorMessage As System.String
        ) As System.Boolean
            Select Case If(value, System.String.Empty).Trim().ToLowerInvariant()
                Case "domcontentloaded"
                    result = Microsoft.Playwright.WaitUntilState.DOMContentLoaded
                Case "load"
                    result = Microsoft.Playwright.WaitUntilState.Load
                Case "networkidle"
                    result = Microsoft.Playwright.WaitUntilState.NetworkIdle
                Case "commit"
                    result = Microsoft.Playwright.WaitUntilState.Commit
                Case Else
                    errorMessage = "wait_until must be one of domcontentloaded, load, networkidle or commit."
                    Return False
            End Select

            errorMessage = System.String.Empty
            Return True
        End Function

        Private Shared Function TryGetTimeout(
            arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object),
            name As System.String,
            defaultValue As System.Int32,
            ByRef result As System.Int32,
            ByRef errorMessage As System.String
        ) As System.Boolean
            result = defaultValue
            errorMessage = System.String.Empty

            Dim raw As System.Object = Nothing
            If arguments Is Nothing OrElse Not TryGetArgument(arguments, name, raw) OrElse raw Is Nothing Then
                Return True
            End If

            raw = UnwrapJsonValue(raw)
            Dim parsed As System.Int32
            If Not System.Int32.TryParse(
                System.Convert.ToString(raw, System.Globalization.CultureInfo.InvariantCulture),
                System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture,
                parsed) Then
                errorMessage = name & " must be an integer."
                Return False
            End If

            If parsed < 1000 OrElse parsed > 120000 Then
                errorMessage = name & " must be between 1000 and 120000."
                Return False
            End If

            result = parsed
            Return True
        End Function

        Private Shared Function GetRequiredString(
            arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object),
            name As System.String
        ) As System.String
            Dim raw As System.Object = Nothing
            If arguments Is Nothing OrElse Not TryGetArgument(arguments, name, raw) OrElse raw Is Nothing Then
                Return System.String.Empty
            End If

            raw = UnwrapJsonValue(raw)
            Return System.Convert.ToString(raw, System.Globalization.CultureInfo.InvariantCulture)
        End Function

        Private Shared Function GetOptionalString(
            arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object),
            name As System.String,
            defaultValue As System.String
        ) As System.String
            Dim raw As System.Object = Nothing
            If arguments Is Nothing OrElse Not TryGetArgument(arguments, name, raw) OrElse raw Is Nothing Then
                Return defaultValue
            End If

            raw = UnwrapJsonValue(raw)
            If raw Is Nothing Then
                Return defaultValue
            End If
            Return System.Convert.ToString(raw, System.Globalization.CultureInfo.InvariantCulture)
        End Function

        Private Shared Function TryGetArgument(
            arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object),
            name As System.String,
            ByRef value As System.Object
        ) As System.Boolean
            If arguments.TryGetValue(name, value) Then
                Return True
            End If

            For Each pair As System.Collections.Generic.KeyValuePair(Of System.String, System.Object) In arguments
                If System.String.Equals(pair.Key, name, System.StringComparison.OrdinalIgnoreCase) Then
                    value = pair.Value
                    Return True
                End If
            Next

            value = Nothing
            Return False
        End Function

        Private Shared Function UnwrapJsonValue(value As System.Object) As System.Object
            If TypeOf value Is Newtonsoft.Json.Linq.JValue Then
                Return DirectCast(value, Newtonsoft.Json.Linq.JValue).Value
            End If
            Return value
        End Function

        Private Shared Function BuildBrowserLaunchFailureMessage(
            channelFailure As System.Exception,
            bundledFailure As System.Exception
        ) As System.String
            Dim builder As New System.Text.StringBuilder()
            builder.Append("Unable to launch a Playwright browser.")

            If channelFailure IsNot Nothing AndAlso Not System.String.IsNullOrWhiteSpace(channelFailure.Message) Then
                builder.Append(" Configured browser channel failed: ")
                builder.Append(channelFailure.Message)
            End If

            If bundledFailure IsNot Nothing AndAlso Not System.String.IsNullOrWhiteSpace(bundledFailure.Message) Then
                builder.Append(" Bundled Chromium fallback failed: ")
                builder.Append(bundledFailure.Message)
            End If

            builder.Append(" Ensure Microsoft.Playwright 1.59+ is deployed and the matching browser binary is installed, or configure an available browser channel.")
            Return builder.ToString()
        End Function

        Private Shared Function SanitizeExceptionMessage(ex As System.Exception) As System.String
            If ex Is Nothing OrElse System.String.IsNullOrWhiteSpace(ex.Message) Then
                Return "Playwright reported an unspecified browser error."
            End If

            Dim message As System.String = ex.Message.Trim()
            If message.Length > 2000 Then
                message = message.Substring(0, 2000)
            End If
            Return message
        End Function
    End Class

End Namespace
