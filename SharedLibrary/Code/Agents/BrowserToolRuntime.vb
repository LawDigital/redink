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
'    resulting YAML contains native refs such as [ref=e7] and includes iframe
'    content.
'  - Tracks the refs present in the latest successful snapshot. browser_interact
'    rejects unknown or stale refs before attempting an action.
'  - Resolves native AI snapshot refs through Playwright's aria-ref selector and
'    invalidates the snapshot immediately when an interaction is attempted.
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

        Public Function Clone() As BrowserToolOptions
            Return New BrowserToolOptions() With {
                .Headless = Headless,
                .BrowserChannel = BrowserChannel,
                .FallbackToBundledChromium = FallbackToBundledChromium,
                .IgnoreHTTPSErrors = IgnoreHTTPSErrors,
                .DefaultTimeoutMs = DefaultTimeoutMs,
                .NavigationTimeoutMs = NavigationTimeoutMs
            }
        End Function
    End Class

    Friend NotInheritable Class BrowserToolRuntime
        Private Shared ReadOnly Gate As New System.Threading.SemaphoreSlim(1, 1)
        Private Shared ReadOnly ConfigurationLock As New System.Object()
        Private Shared ReadOnly PrivateNetworkApprovalLock As New System.Object()
        Private Shared ReadOnly SessionApprovedPrivateOrigins As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
        Private Shared ReadOnly SnapshotRefRegex As New System.Text.RegularExpressions.Regex(
            "\[ref=(e[0-9]+)\]",
            System.Text.RegularExpressions.RegexOptions.Compiled Or System.Text.RegularExpressions.RegexOptions.CultureInvariant)

        Private Shared ConfiguredOptions As New BrowserToolOptions()
        Private Shared PlaywrightInstance As Microsoft.Playwright.IPlaywright
        Private Shared Browser As Microsoft.Playwright.IBrowser
        Private Shared BrowserContext As Microsoft.Playwright.IBrowserContext
        Private Shared CurrentPage As Microsoft.Playwright.IPage
        Private Shared SnapshotIsValid As System.Boolean
        Private Shared SnapshotGeneration As System.Int64
        Private Shared LastSnapshotRefs As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)

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

        Public Shared Async Function ExecuteAsync(
            toolName As System.String,
            arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object),
            cancellationToken As System.Threading.CancellationToken
        ) As System.Threading.Tasks.Task(Of System.String)
            If System.String.Equals(toolName, BrowserTools.BrowserOpenToolName, System.StringComparison.OrdinalIgnoreCase) Then
                Return Await OpenAsync(arguments, cancellationToken).ConfigureAwait(False)
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
            stateMayHaveChanged As System.Boolean
        ) As System.String
            Return New Newtonsoft.Json.Linq.JObject(
                New Newtonsoft.Json.Linq.JProperty("status", "error"),
                New Newtonsoft.Json.Linq.JProperty("tool", If(toolName, System.String.Empty)),
                New Newtonsoft.Json.Linq.JProperty("requires_snapshot", stateMayHaveChanged),
                New Newtonsoft.Json.Linq.JProperty("error", New Newtonsoft.Json.Linq.JObject(
                    New Newtonsoft.Json.Linq.JProperty("code", code),
                    New Newtonsoft.Json.Linq.JProperty("message", message),
                    New Newtonsoft.Json.Linq.JProperty("retryable", retryable),
                    New Newtonsoft.Json.Linq.JProperty("state_may_have_changed", stateMayHaveChanged)))).ToString(Newtonsoft.Json.Formatting.None)
        End Function

        Private Shared Async Function OpenAsync(
            arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object),
            cancellationToken As System.Threading.CancellationToken
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
                Await EnsureRuntimeAsync(options, cancellationToken).ConfigureAwait(False)
                If CurrentPage Is Nothing OrElse CurrentPage.IsClosed Then
                    CurrentPage = Await BrowserContext.NewPageAsync().ConfigureAwait(False)
                End If

                InvalidateSnapshot()

                Dim gotoOptions As New Microsoft.Playwright.PageGotoOptions() With {
                    .waitUntil = waitUntil,
                    .Timeout = CSng(timeoutMs)
                }

                Await CurrentPage.GotoAsync(url, gotoOptions).ConfigureAwait(False)
                SelectNewestOpenPage()
                Await TryDismissCommonCookieConsentAsync(CurrentPage, cancellationToken).ConfigureAwait(False)

                Dim title As System.String = Await CurrentPage.TitleAsync().ConfigureAwait(False)
                Return New Newtonsoft.Json.Linq.JObject(
                    New Newtonsoft.Json.Linq.JProperty("status", "ok"),
                    New Newtonsoft.Json.Linq.JProperty("tool", BrowserTools.BrowserOpenToolName),
                    New Newtonsoft.Json.Linq.JProperty("url", CurrentPage.Url),
                    New Newtonsoft.Json.Linq.JProperty("title", title),
                    New Newtonsoft.Json.Linq.JProperty("requires_snapshot", True)).ToString(Newtonsoft.Json.Formatting.None)
            Catch ex As System.OperationCanceledException
                Return CreateErrorPayload(
                    BrowserTools.BrowserOpenToolName,
                    "CANCELLED",
                    "Browser navigation was cancelled.",
                    True,
                    True)
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine(ex.ToString())
                Return CreateErrorPayload(
                    BrowserTools.BrowserOpenToolName,
                    "BROWSER_OPEN_FAILED",
                    SanitizeExceptionMessage(ex),
                    True,
                    True)
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

            If Not System.Text.RegularExpressions.Regex.IsMatch(refValue, "^e[0-9]+$", System.Text.RegularExpressions.RegexOptions.CultureInvariant) Then
                Return CreateErrorPayload(
                    BrowserTools.BrowserInteractToolName,
                    "INVALID_REF",
                    "The ref must have the form e<number>, for example e7.",
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
                        False,
                        False)
                End If

                If Not LastSnapshotRefs.Contains(refValue) Then
                    Return CreateErrorPayload(
                        BrowserTools.BrowserInteractToolName,
                        "STALE_OR_UNKNOWN_REF",
                        "The ref is not present in the most recent browser_snapshot. Take a new snapshot and use a ref from it.",
                        False,
                        False)
                End If

                SelectNewestOpenPage()
                Dim locator As Microsoft.Playwright.ILocator = CurrentPage.Locator("aria-ref=" & refValue)

                If action = "click" OrElse action = "double_click" Then
                    Dim directNavigationUrl As System.String = Await TryGetDirectNavigationUrlAsync(locator, CurrentPage.Url).ConfigureAwait(False)
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

                ' Invalidate only after security preflight. A denied click has not changed page state,
                ' while an attempted Playwright action may have partially changed it even on failure.
                InvalidateSnapshot()
                Await ExecuteLocatorActionAsync(locator, action, value, timeoutMs).ConfigureAwait(False)
                SelectNewestOpenPage()
                Await TryDismissCommonCookieConsentAsync(CurrentPage, cancellationToken).ConfigureAwait(False)

                Dim title As System.String = Await CurrentPage.TitleAsync().ConfigureAwait(False)
                Return New Newtonsoft.Json.Linq.JObject(
                    New Newtonsoft.Json.Linq.JProperty("status", "ok"),
                    New Newtonsoft.Json.Linq.JProperty("tool", BrowserTools.BrowserInteractToolName),
                    New Newtonsoft.Json.Linq.JProperty("ref", refValue),
                    New Newtonsoft.Json.Linq.JProperty("action", action),
                    New Newtonsoft.Json.Linq.JProperty("url", CurrentPage.Url),
                    New Newtonsoft.Json.Linq.JProperty("title", title),
                    New Newtonsoft.Json.Linq.JProperty("requires_snapshot", True)).ToString(Newtonsoft.Json.Formatting.None)
            Catch ex As System.OperationCanceledException
                Return CreateErrorPayload(
                    BrowserTools.BrowserInteractToolName,
                    "CANCELLED",
                    "Browser interaction was cancelled.",
                    True,
                    True)
            Catch ex As System.Exception
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
        ) As System.Threading.Tasks.Task
            If page Is Nothing OrElse page.IsClosed Then
                Return
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
                    Return
                Catch ex As System.Exception
                    ' Consent handling is best effort only. A failure must never make
                    ' browser_open/browser_interact fail; the next snapshot can expose
                    ' the banner so the model can handle it explicitly.
                End Try
            Next
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

        Private Shared Async Function EnsureRuntimeAsync(
            options As BrowserToolOptions,
            cancellationToken As System.Threading.CancellationToken
        ) As System.Threading.Tasks.Task
            cancellationToken.ThrowIfCancellationRequested()

            If Browser IsNot Nothing AndAlso Browser.IsConnected AndAlso BrowserContext IsNot Nothing Then
                Return
            End If

            Await DisposeRuntimeAsync().ConfigureAwait(False)
            cancellationToken.ThrowIfCancellationRequested()

            PlaywrightInstance = Await Microsoft.Playwright.Playwright.CreateAsync().ConfigureAwait(False)

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
            BrowserContext = Await Browser.NewContextAsync(contextOptions).ConfigureAwait(False)
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
                    Return True
                End If
            End SyncLock

            Dim details As System.String = System.String.Empty
            If privateAddresses IsNot Nothing AndAlso privateAddresses.Count > 0 Then
                details = System.Environment.NewLine & System.Environment.NewLine &
                          "Resolved private/local address: " & System.String.Join(", ", privateAddresses.ToArray())
            End If

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
                Return False
            End If

            SyncLock PrivateNetworkApprovalLock
                SessionApprovedPrivateOrigins.Add(origin)
            End SyncLock
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
