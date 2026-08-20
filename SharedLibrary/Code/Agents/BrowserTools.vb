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
'                      snapshot contains native Playwright refs such as [ref=e7].
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
'  - Refs are intentionally accepted only from the latest valid snapshot. After
'    every interaction, a new browser_snapshot is required before another action.
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

        ''' <summary>
        ''' Applies browser runtime options used for future browser launches.
        ''' This is optional; sensible Windows/Office defaults are used otherwise.
        ''' </summary>
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
                    "Pass an absolute http:// or https:// URL. This tool navigates only; after it succeeds, call browser_snapshot to inspect the rendered page before attempting interaction. Do not invent element refs."
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
                    "The returned YAML snapshot contains Playwright refs such as [ref=e7]. Treat those refs as short-lived handles. Only refs from the most recent successful browser_snapshot may be passed to browser_interact. " &
                    "If the relevant link is already visible in the snapshot, use its current ref rather than falling back to a new internet search."
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
                    "Supported actions are click, double_click, fill, clear, press, select, check, uncheck, hover and focus. The value argument is required for fill, press and select and must be omitted for the other actions. " &
                    "After every successful or attempted interaction, call browser_snapshot again before another browser_interact. Interactions can change remote state, submit forms or trigger navigation; apply the user's intent and normal safety rules."
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
                    New Newtonsoft.Json.Linq.JProperty("description", "Navigation timeout in milliseconds."))))

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
                    New Newtonsoft.Json.Linq.JProperty("pattern", "^e[0-9]+$"),
                    New Newtonsoft.Json.Linq.JProperty("description", "Playwright ref from the most recent browser_snapshot, for example e7."))),
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
