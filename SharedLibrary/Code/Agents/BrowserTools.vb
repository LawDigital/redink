' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: BrowserTools.vb
' Purpose: Exposes the shared Playwright browser tools as Red Ink `ModelConfig`
'          instances and routes execution to the host-agnostic browser runtime.
'
' Tool verbs:
'  - browser_open:     Starts/reuses the shared browser session and navigates to
'                      an absolute HTTP/HTTPS URL.
'  - browser_snapshot: Captures Playwright's AI-optimized ARIA snapshot. The
'                      snapshot contains native Playwright refs such as [ref=e7] or frame-qualified refs such as [ref=f11e38].
'  - browser_interact: Performs one browser action against a ref from the most
'                      recent browser_snapshot result.
'
' Agent loop:
'  browser_open -> browser_snapshot -> browser_interact -> browser_snapshot -> ...
'
' Dependencies:
'  - NuGet package Microsoft.Playwright 1.59 or newer.
'  - Newtonsoft.Json (already used by the Red Ink shared agent layer).
'
' Notes:
'  - The tool is host-agnostic and does not reference Outlook or Word interop.
'  - browser_interact is write-capable: clicks, form filling, selections and key
'    presses can change remote application state.
'  - Refs are intentionally accepted only from the latest valid snapshot. A
'    successful fill/clear/focus/hover may retain that snapshot when the page and
'    URL remain unchanged; responses with requires_snapshot=true require a fresh
'    browser_snapshot before another action.
' =============================================================================

Option Explicit On
Option Strict On
Option Infer On

Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace Agents

    Public NotInheritable Class BrowserTools
        Public Const BrowserOpenToolName As System.String = "browser_open"
        Public Const BrowserSnapshotToolName As System.String = "browser_snapshot"
        Public Const BrowserInteractToolName As System.String = "browser_interact"

        Private Const BrowserOpenPriority As System.Int32 = 990
        Private Const BrowserSnapshotPriority As System.Int32 = 989
        Private Const BrowserInteractPriority As System.Int32 = 988

        Private Sub New()
        End Sub

        Public Shared Function IsBrowserTool(name As System.String) As System.Boolean
            If System.String.IsNullOrWhiteSpace(name) Then
                Return False
            End If

            Return System.String.Equals(name, BrowserOpenToolName, System.StringComparison.OrdinalIgnoreCase) OrElse
                   System.String.Equals(name, BrowserSnapshotToolName, System.StringComparison.OrdinalIgnoreCase) OrElse
                   System.String.Equals(name, BrowserInteractToolName, System.StringComparison.OrdinalIgnoreCase)
        End Function

        Public Shared Function IsDisabled(sharedContext As ISharedContext) As System.Boolean
            Return sharedContext IsNot Nothing AndAlso sharedContext.INI_BrowserToolsDisable
        End Function

        Public Shared Function IsRuntimeAvailable(sharedContext As ISharedContext) As System.Boolean
            Dim ignored As System.String = System.String.Empty
            Return IsRuntimeAvailable(sharedContext, ignored)
        End Function

        Public Shared Function IsRuntimeAvailable(sharedContext As ISharedContext, ByRef errorMessage As System.String) As System.Boolean
            errorMessage = System.String.Empty

            Try
                If sharedContext Is Nothing Then
                    errorMessage = "Browser tools require a host context with PlayWrightPath configuration."
                    Return False
                End If

                Dim configuredPath As System.String = If(sharedContext.INI_PlayWrightPath, System.String.Empty)
                Dim useLocalCache As System.Boolean = sharedContext.INI_PlayWrightUseLocalCache
                BrowserToolRuntime.ConfigureExternalRuntime(configuredPath, useLocalCache)

                ' Availability stays non-blocking. When local caching is enabled the resolver starts
                ' the one-time copy in the background, but immediately returns the valid source runtime.
                Dim resolution As PlaywrightRuntimeResolution = Nothing
                Return PlaywrightRuntimeResolver.TryResolve(
                    configuredPath,
                    useLocalCache,
                    useLocalCache,
                    resolution,
                    errorMessage)
            Catch ex As System.Exception
                errorMessage = "The configured Playwright runtime could not be validated. Browser tools are unavailable for this run."
                System.Diagnostics.Trace.WriteLine("Browser runtime availability check failed: " & ex.ToString())
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Applies browser runtime options used for future browser launches.
        ''' This is optional; sensible Windows/Office defaults are used otherwise.
        ''' </summary>
        Private Shared Function CompactAvailabilityLogValue(value As System.String) As System.String
            If System.String.IsNullOrWhiteSpace(value) Then Return "(none)"
            Return value.Replace(System.Environment.NewLine, " ").Replace(System.Convert.ToChar(13), " "c).Replace(System.Convert.ToChar(10), " "c).Trim()
        End Function

        Public Shared Sub Configure(options As BrowserToolOptions)
            If options Is Nothing Then
                Throw New System.ArgumentNullException(NameOf(options))
            End If

            BrowserToolRuntime.Configure(options)
        End Sub

        Public Overloads Shared Function BuildAll() As System.Collections.Generic.List(Of ModelConfig)
            Return BuildAll(Nothing)
        End Function

        Public Overloads Shared Function BuildAll(sharedContext As ISharedContext) As System.Collections.Generic.List(Of ModelConfig)
            Dim result As New System.Collections.Generic.List(Of ModelConfig)()

            If IsDisabled(sharedContext) Then
                Return result
            End If

            Dim runtimeError As System.String = System.String.Empty
            If Not IsRuntimeAvailable(sharedContext, runtimeError) Then
                If Not System.String.IsNullOrWhiteSpace(runtimeError) Then
                    System.Diagnostics.Trace.WriteLine("Browser tools not exposed: " & runtimeError)
                    Try
                        Global.SharedLibrary.SharedLibrary.UpdateHandler.WriteUpdateLog("[ToolAvailability] browser unavailable; reason=" & CompactAvailabilityLogValue(runtimeError))
                    Catch ex As System.Exception
                        System.Diagnostics.Trace.WriteLine("Could not write browser availability diagnostics: " & ex.ToString())
                    End Try
                End If
                Return result
            End If

            result.Add(BuildBrowserOpen())
            result.Add(BuildBrowserSnapshot())
            result.Add(BuildBrowserInteract())
            Return result
        End Function

        Public Overloads Shared Function Execute(
            toolName As System.String,
            arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object)
        ) As System.String
            Return Execute(toolName, arguments, Nothing)
        End Function

        Public Overloads Shared Function Execute(
            toolName As System.String,
            arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object),
            sharedContext As ISharedContext
        ) As System.String
            Try
                Return ExecuteAsync(toolName, arguments, System.Threading.CancellationToken.None, sharedContext).ConfigureAwait(False).GetAwaiter().GetResult()
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine(ex.ToString())
                Return BrowserToolRuntime.CreateErrorPayload(
                    toolName,
                    "UNEXPECTED_HOST_ERROR",
                    "The browser tool failed before a structured result could be produced.",
                    False,
                    False)
            End Try
        End Function

        Public Overloads Shared Async Function ExecuteAsync(
            toolName As System.String,
            arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object),
            cancellationToken As System.Threading.CancellationToken
        ) As System.Threading.Tasks.Task(Of System.String)
            Return Await ExecuteAsync(toolName, arguments, cancellationToken, Nothing).ConfigureAwait(False)
        End Function

        Public Overloads Shared Async Function ExecuteAsync(
            toolName As System.String,
            arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object),
            cancellationToken As System.Threading.CancellationToken,
            sharedContext As ISharedContext
        ) As System.Threading.Tasks.Task(Of System.String)
            If Not IsBrowserTool(toolName) Then
                Return BrowserToolRuntime.CreateErrorPayload(
                    toolName,
                    "UNKNOWN_BROWSER_TOOL",
                    "Unknown browser tool name.",
                    False,
                    False)
            End If

            If IsDisabled(sharedContext) Then
                Return BrowserToolRuntime.CreateErrorPayload(
                    toolName,
                    "BROWSER_TOOLS_DISABLED",
                    "Browser tools are disabled by configuration.",
                    False,
                    False)
            End If

            Dim runtimeError As System.String = System.String.Empty
            If Not IsRuntimeAvailable(sharedContext, runtimeError) Then
                Return BrowserToolRuntime.CreateErrorPayload(
                    toolName,
                    "PLAYWRIGHT_RUNTIME_UNAVAILABLE",
                    runtimeError,
                    False,
                    False)
            End If

            Dim safeArguments As System.Collections.Generic.IDictionary(Of System.String, System.Object) = arguments
            If safeArguments Is Nothing Then
                safeArguments = New System.Collections.Generic.Dictionary(Of System.String, System.Object)(System.StringComparer.OrdinalIgnoreCase)
            End If

            Return Await BrowserToolRuntime.ExecuteAsync(toolName, safeArguments, cancellationToken).ConfigureAwait(False)
        End Function

        ''' <summary>
        ''' Closes the current Playwright context and browser, if one is running.
        ''' This is not advertised as an agent tool; hosts may call it during shutdown.
        ''' </summary>
        Public Shared Sub Shutdown()
            BrowserToolRuntime.Shutdown()
        End Sub

        Private Shared Function BuildBrowserOpen() As ModelConfig
            Return New ModelConfig() With {
                .ToolName = BrowserOpenToolName,
                .Tool = True,
                .ToolPriority = BrowserOpenPriority,
                .ToolErrorHandling = "skip",
                .ModelDescription = "Open or navigate the shared Playwright browser session",
                .ToolDefinition = BuildBrowserOpenDefinition(),
                .ToolInstructionsPrompt =
                    "Use browser_open when a specific website must be explored as a rendered browser page, especially to find links, menus, sections, downloads, pagination, JavaScript-rendered content, or controls that simple HTTP text retrieval may miss. " &
                    "Prefer web_grounding when the relevant site/page is not yet known and public-web discovery is required. Prefer retrieve_web_content for a known mostly-static URL when readable text/ordinary links are sufficient. " &
                    "Pass an absolute http:// or https:// URL. The browser runs headless by default. This tool navigates only; after it succeeds, call browser_snapshot to inspect the rendered page before attempting interaction. " &
                    "AUTHENTICATION: authentication=auto (default) silently reuses a previously provisioned, Windows-DPAPI-protected browser session for this origin when available. If the rendered page requires sign-in and a live user is available, call browser_open again for the target URL with authentication=interactive. If the initial authentication=auto navigation itself returns BROWSER_OPEN_FAILED before a snapshot can be taken, a protected enterprise/intranet site may be waiting on browser-native authentication, SSO, network permission UI or another user challenge; when a live user is available, one interactive retry for the same URL/profile is allowed. Do not loop. Errors explicitly marked retryable=false are terminal and must not trigger authentication retries; PLAYWRIGHT_DRIVER_UNAVAILABLE, PLAYWRIGHT_RUNTIME_UNAVAILABLE, PLAYWRIGHT_BROWSER_UNAVAILABLE or PLAYWRIGHT_RUNTIME_FAILED are terminal runtime errors and must not trigger authentication retries. Red Ink opens a dedicated visible Playwright browser so the user can complete username/password, multi-page sign-in, SSO, MFA or other challenges without exposing secrets to the model; the visible browser intentionally remains usable even if its initial automated navigation cannot fully complete. After the user confirms sign-in, Red Ink keeps that exact authenticated browser/context alive for the current run (some enterprise sites require a fresh login in every new browser process), persists storage state for best-effort reuse, and hides the dedicated browser window while automation continues. In unattended AutoPilot/e-mail Scheduler runs interactive authentication is forbidden and returns AUTHENTICATION_REQUIRED; never loop or guess credentials. authentication=none explicitly suppresses loading a stored session. auth_profile may be supplied only to intentionally share a provisioned session across related URLs; otherwise omit it and the normalized origin is used. " &
                    "The runtime may conservatively dismiss common reject/necessary-only cookie banners, but never grants optional tracking by clicking accept-all. Do not invent element refs and never place passwords, MFA codes, recovery codes or other authentication secrets in browser_interact values."
            }
        End Function

        Private Shared Function BuildBrowserSnapshot() As ModelConfig
            Return New ModelConfig() With {
                .ToolName = BrowserSnapshotToolName,
                .Tool = True,
                .ToolPriority = BrowserSnapshotPriority,
                .ToolErrorHandling = "skip",
                .ModelDescription = "Observe the current browser through an AI-optimized ARIA snapshot",
                .ToolDefinition = BuildBrowserSnapshotDefinition(),
                .ToolInstructionsPrompt =
                    "Use browser_snapshot to inspect the actually rendered page structure, including accessible links, buttons, menus, headings, form controls and other interactive elements. It is especially useful for scanning a specific website for relevant links/pages or for content/navigation produced by JavaScript. " &
                    "The returned YAML snapshot contains native Playwright refs such as [ref=e7] and frame-qualified refs such as [ref=f11e38]. Treat the complete ref token as a short-lived handle. Only refs from the most recent successful browser_snapshot may be passed to browser_interact. " &
                    "If a cookie/consent overlay is still present, deal with that overlay FIRST before unrelated navigation; prefer reject/necessary-only choices over accept-all. Then take a fresh browser_snapshot. If the relevant link is already visible in the snapshot, use its current ref rather than falling back to a new web_grounding search."
            }
        End Function

        Private Shared Function BuildBrowserInteract() As ModelConfig
            Return New ModelConfig() With {
                .ToolName = BrowserInteractToolName,
                .Tool = True,
                .ToolPriority = BrowserInteractPriority,
                .ToolErrorHandling = "skip",
                .ModelDescription = "Perform one Playwright action against a ref from the latest browser snapshot",
                .ToolDefinition = BuildBrowserInteractDefinition(),
                .ToolInstructionsPrompt =
                    "Use browser_interact only when navigation or another browser action is actually needed after inspecting a browser_snapshot. For site exploration, click the relevant link/menu/pagination ref rather than restarting discovery with web_grounding. " &
                    "Supported actions are click, double_click, fill, clear, press, select, check, uncheck, hover and focus. The value argument is required for fill, press and select and must be omitted for the other actions. Never put passwords, MFA/one-time codes, recovery codes or other authentication secrets in value; direct filling of detected password/one-time-code controls is rejected by the runtime. Use browser_open(authentication=interactive) for secret-bearing authentication. " &
                    "A browser_interact result explicitly tells you whether a fresh snapshot is required. If requires_snapshot=true, call browser_snapshot before any further browser_interact. If requires_snapshot=false and snapshot_retained=true, you may continue a short same-form sequence with another ref from that same snapshot (for example fill a textbox and then click its submit button). For click/double_click, target_url contains the absolute href when the clicked element exposed one; treat it only as the captured target of that element, not as proof that a document was successfully retrieved. navigation_observed tells you whether the active page or URL changed. Never claim that a result/document was opened merely because the search page itself has a URL. Never assume refs survived a click, submit, navigation or other response that requires a new snapshot. Interactions can change remote state, submit forms or trigger navigation; apply the user's intent and normal safety rules."
            }
        End Function

        Private Shared Function BuildBrowserOpenDefinition() As System.String
            Dim properties As New Newtonsoft.Json.Linq.JObject(
                New Newtonsoft.Json.Linq.JProperty("url", New Newtonsoft.Json.Linq.JObject(
                    New Newtonsoft.Json.Linq.JProperty("type", "string"),
                    New Newtonsoft.Json.Linq.JProperty("description", "Absolute http:// or https:// URL to open."))),
                New Newtonsoft.Json.Linq.JProperty("wait_until", New Newtonsoft.Json.Linq.JObject(
                    New Newtonsoft.Json.Linq.JProperty("type", "string"),
                    New Newtonsoft.Json.Linq.JProperty("enum", New Newtonsoft.Json.Linq.JArray("domcontentloaded", "load", "networkidle", "commit")),
                    New Newtonsoft.Json.Linq.JProperty("description", "Navigation readiness state. Defaults to domcontentloaded."))),
                New Newtonsoft.Json.Linq.JProperty("timeout_ms", New Newtonsoft.Json.Linq.JObject(
                    New Newtonsoft.Json.Linq.JProperty("type", "integer"),
                    New Newtonsoft.Json.Linq.JProperty("minimum", 1000),
                    New Newtonsoft.Json.Linq.JProperty("maximum", 120000),
                    New Newtonsoft.Json.Linq.JProperty("description", "Navigation timeout in milliseconds."))),
                New Newtonsoft.Json.Linq.JProperty("authentication", New Newtonsoft.Json.Linq.JObject(
                    New Newtonsoft.Json.Linq.JProperty("type", "string"),
                    New Newtonsoft.Json.Linq.JProperty("enum", New Newtonsoft.Json.Linq.JArray("auto", "interactive", "none")),
                    New Newtonsoft.Json.Linq.JProperty("description", "Authentication handling. auto (default) reuses a securely provisioned browser session when available; interactive opens a visible browser for the live user to sign in and then securely persists the resulting session; none suppresses stored-session loading."))),
                New Newtonsoft.Json.Linq.JProperty("auth_profile", New Newtonsoft.Json.Linq.JObject(
                    New Newtonsoft.Json.Linq.JProperty("type", "string"),
                    New Newtonsoft.Json.Linq.JProperty("minLength", 1),
                    New Newtonsoft.Json.Linq.JProperty("maxLength", 200),
                    New Newtonsoft.Json.Linq.JProperty("description", "Optional stable session profile name. Omit to scope the saved session to the normalized URL origin. Use only when intentionally sharing one authenticated session across related URLs."))))

            Dim required As New Newtonsoft.Json.Linq.JArray()
            required.Add("url")
            Return "{""name"":""" & BrowserOpenToolName & """,""description"":""Open or navigate the shared Playwright browser session"",""parameters"":" &
                   BuildObjectSchema(properties, required) & "}"
        End Function

        Private Shared Function BuildBrowserSnapshotDefinition() As System.String
            Dim properties As New Newtonsoft.Json.Linq.JObject(
                New Newtonsoft.Json.Linq.JProperty("timeout_ms", New Newtonsoft.Json.Linq.JObject(
                    New Newtonsoft.Json.Linq.JProperty("type", "integer"),
                    New Newtonsoft.Json.Linq.JProperty("minimum", 1000),
                    New Newtonsoft.Json.Linq.JProperty("maximum", 120000),
                    New Newtonsoft.Json.Linq.JProperty("description", "Snapshot timeout in milliseconds."))))

            Return "{""name"":""" & BrowserSnapshotToolName & """,""description"":""Observe the current browser through an AI-optimized ARIA snapshot"",""parameters"":" &
                   BuildObjectSchema(properties, New Newtonsoft.Json.Linq.JArray()) & "}"
        End Function

        Private Shared Function BuildBrowserInteractDefinition() As System.String
            Dim properties As New Newtonsoft.Json.Linq.JObject(
                New Newtonsoft.Json.Linq.JProperty("ref", New Newtonsoft.Json.Linq.JObject(
                    New Newtonsoft.Json.Linq.JProperty("type", "string"),
                    New Newtonsoft.Json.Linq.JProperty("pattern", "^[A-Za-z0-9_-]+$"),
                    New Newtonsoft.Json.Linq.JProperty("description", "Native Playwright ref from the most recent browser_snapshot, for example e7 or f11e38."))),
                New Newtonsoft.Json.Linq.JProperty("action", New Newtonsoft.Json.Linq.JObject(
                    New Newtonsoft.Json.Linq.JProperty("type", "string"),
                    New Newtonsoft.Json.Linq.JProperty("enum", New Newtonsoft.Json.Linq.JArray(
                        "click",
                        "double_click",
                        "fill",
                        "clear",
                        "press",
                        "select",
                        "check",
                        "uncheck",
                        "hover",
                        "focus")),
                    New Newtonsoft.Json.Linq.JProperty("description", "Single browser action to perform."))),
                New Newtonsoft.Json.Linq.JProperty("value", New Newtonsoft.Json.Linq.JObject(
                    New Newtonsoft.Json.Linq.JProperty("type", "string"),
                    New Newtonsoft.Json.Linq.JProperty("description", "Required for fill, press and select. For select, matches option value or label."))),
                New Newtonsoft.Json.Linq.JProperty("timeout_ms", New Newtonsoft.Json.Linq.JObject(
                    New Newtonsoft.Json.Linq.JProperty("type", "integer"),
                    New Newtonsoft.Json.Linq.JProperty("minimum", 1000),
                    New Newtonsoft.Json.Linq.JProperty("maximum", 120000),
                    New Newtonsoft.Json.Linq.JProperty("description", "Action timeout in milliseconds."))))

            Dim required As New Newtonsoft.Json.Linq.JArray()
            required.Add("ref")
            required.Add("action")
            Return "{""name"":""" & BrowserInteractToolName & """,""description"":""Perform one Playwright action against a ref from the latest browser snapshot"",""parameters"":" &
                   BuildObjectSchema(properties, required) & "}"
        End Function

        Private Shared Function BuildObjectSchema(
            properties As Newtonsoft.Json.Linq.JObject,
            required As Newtonsoft.Json.Linq.JArray
        ) As System.String
            Return New Newtonsoft.Json.Linq.JObject(
                New Newtonsoft.Json.Linq.JProperty("type", "object"),
                New Newtonsoft.Json.Linq.JProperty("properties", properties),
                New Newtonsoft.Json.Linq.JProperty("required", required),
                New Newtonsoft.Json.Linq.JProperty("additionalProperties", False)).ToString(Newtonsoft.Json.Formatting.None)
        End Function
    End Class

End Namespace
