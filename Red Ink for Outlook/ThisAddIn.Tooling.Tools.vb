' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.Tooling.Tools.vb
' Purpose:
'   Outlook Local Chat tooling support layer for host-specific tool definitions,
'   registration helpers and execution utilities that are not implemented in the shared
'   AgentToolRouter.
'
' Architecture / Function:
'   - Builds/augments the Outlook tool catalog and bridges Outlook-specific capabilities
'     into the common tooling loop while shared tools remain registered centrally.
'   - Provides web-content retrieval/link extraction, WebView2 integration, reflection
'     helpers and host adapters used by individual Outlook tools.
'   - Keeps model-facing tool schemas/instructions separate from the sequencing and
'     finalization rules in SharedLibrary.Agents.
'   - Binary downloads, M365, Python, Office and other substantial capabilities are
'     delegated to their dedicated modules rather than being implemented monolithically here.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.Reflection
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Microsoft.Web.WebView2.Core
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedContext
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn


    ''' <summary>
    ''' Enables the WebView2-based implementation for internal web content retrieval.
    ''' </summary>
    Private Const UseWebView2 As Boolean = True

    ''' <summary>
    ''' Internal tool name used in model tool calls for web content retrieval.
    ''' </summary>
    Private Const InternalWebToolName As String = "retrieve_web_content"

    ''' <summary>
    ''' Human-readable tool instructions shown to the model describing how to use the internal web tool.
    ''' </summary>
    Private Const InternalWebToolInstructionsPrompt As String =
        "retrieve_web_content: Retrieves text content from one or more URLs. Use this to fetch information from websites when needed. " &
        "By default, this tool returns readable page text only. " &
        "OPTIONAL LINK EXTRACTION: If you need downloadable links such as PDFs, explicitly set include_links=true. " &
        "To discover links hidden behind collapsed accordions, tabs, or drop-down sections, explicitly set expand_interactive_sections=true. " &
        "Optionally set link_extensions to restrict extracted links to specific extensions such as ['pdf']. " &
        "When include_links=true, the tool returns the normal text plus a structured <LINKS_n> JSON block for each URL. " &
        "IMPORTANT BOUNDARY: This tool retrieves readable text and link metadata only. It does NOT download or preserve original binary file bytes. " &
        "If you call this tool on a direct PDF, DOCX, XLSX, PPTX, ZIP, image, audio, or other binary URL, the result is text extraction or page content — not a real downloadable file object. " &
        "Therefore, NEVER use this tool as a file-downloader and NEVER save its returned text as if it were the original PDF or other binary file. " &
        "If the user wants the actual remote file saved into the workspace, use a dedicated binary download tool if available; otherwise explain that the current toolset can analyze the content but cannot save the original remote binary file. " &
        "IMPORTANT FALLBACK RULE: If include_links=true returns zero matching links, and the page may be dynamically computing or revealing links client-side, then use js_run as a fallback with allow_network=true and navigate_url set to that page. " &
        "In js_run, the code is already the body of an async function, so it must return the final value explicitly at top level. " &
        "SHAREPOINT/ONEDRIVE LIMITATION: This tool CANNOT access SharePoint, OneDrive, Microsoft Teams, or any other " &
        "authenticated cloud storage URLs. URLs containing 'sharepoint.com', 'onedrive.com', '1drv.ms', " &
        "'teams.microsoft.com', or ':f:/' point to resources that require authentication and will NOT return " &
        "useful content. UNC paths (e.g. \\server\share\file.doc) that resolve to SharePoint will also fail. " &
        "If the user asks you to retrieve content from such a link, do NOT call this tool. Instead, explain " &
        "that you cannot remotely log into authenticated cloud storage and ask the user to download the file(s) " &
        "and provide them as direct attachments."

    ''' <summary>
    ''' Canonical tool definition JSON used for the internal web tool.
    ''' </summary>
    Private Const InternalWebToolDefinition As String =
        "{""name"":""retrieve_web_content"",""description"":""Retrieves readable text from one or more web URLs. " &
        "By default this tool returns text only. " &
        "Optional behavior: set include_links=true to also extract structured hyperlinks; set expand_interactive_sections=true " &
        "to attempt opening accordions, details elements, and similar collapsed sections before extraction; set link_extensions " &
        "to filter extracted links to specific extensions such as pdf. " &
        "IMPORTANT: Cannot access SharePoint, OneDrive, Teams, or other authenticated cloud storage URLs " &
        "(sharepoint.com, onedrive.com, 1drv.ms, teams.microsoft.com, :f:/). " &
        "Do NOT call this tool for such URLs — ask the user to download and attach the file(s) instead.""," &
        """parameters"":{""type"":""object"",""properties"":{" &
        """urls"":{""type"":""array"",""items"":{""type"":""string""},""description"":""Array of URLs to retrieve content from""}," &
        """include_links"":{""type"":""boolean"",""description"":""Optional. Default false. When true, include structured extracted links in a <LINKS_n> JSON block for each URL.""}," &
        """link_extensions"":{""type"":""array"",""items"":{""type"":""string""},""description"":""Optional. Restrict extracted links to these file extensions, for example ['pdf']. Ignored unless include_links=true.""}," &
        """expand_interactive_sections"":{""type"":""boolean"",""description"":""Optional. Default false. When true, attempts to open accordions, details sections, tabs, and similar collapsed UI before extracting text and links.""}}" &
        ",""required"":[""urls""]}}"

    Private Const InternalDownloadWebFilesToolName As String = "download_web_files"

    Private Const InternalDownloadWebFilesToolInstructionsPrompt As String =
        "download_web_files: Downloads one or more remote files and saves the original binary bytes to disk. " &
        "Use this only when the user wants the actual file saved locally. " &
        "Provide urls (required array). Optionally provide target_directory (relative subfolder under the current workspace when available; otherwise under the safe download root) and overwrite (boolean, default false). " &
        "Use this tool instead of retrieve_web_content when the user wants the actual PDF or other binary file, not merely extracted text."

    Private Const InternalDownloadWebFilesToolDefinition As String =
        "{""name"":""download_web_files"",""description"":""Downloads one or more remote files and saves the original binary bytes to disk. " &
        "Use this for real file downloads, not for text extraction. Only absolute HTTP/HTTPS URLs are allowed. " &
        "Authenticated SharePoint, OneDrive, Teams, and similar cloud storage URLs are not supported.""," &
        """parameters"":{""type"":""object"",""properties"":{" &
        """urls"":{""type"":""array"",""items"":{""type"":""string""},""description"":""Array of absolute HTTP/HTTPS URLs to download.""}," &
        """target_directory"":{""type"":""string"",""description"":""Optional relative target folder. If a writable workspace is active, this is resolved under the workspace root; otherwise under the safe download root.""}," &
        """overwrite"":{""type"":""boolean"",""description"":""Optional. Default false. If false, existing files are not overwritten; unique names are created instead.""}}" &
        ",""required"":[""urls""]}}"


    ' Internet Search Tooling (available only when INI_ISearch is enabled and INI_ISearch_URL is configured)

    ''' <summary>
    ''' Internal tool name used in model tool calls for internet search.
    ''' </summary>
    Private Const InternalSearchToolName As String = "internet_search"

    ''' <summary>
    ''' Canonical tool definition JSON used for the internal search tool.
    ''' </summary>
    Private Const InternalSearchToolDefinition As String =
        "{""name"":""internet_search""," &
        """description"":""Searches the internet via the configured search engine, retrieves the top result pages, and returns their readable text content. Use this when you need up-to-date or factual information you are not confident about. PRIVACY: The query is sent to an external search engine. Never include personal data, confidential information, private names, case details, contract terms, internal identifiers, or any non-public information in the query. Only public figures, public institutions, published legislation, and other clearly public information may appear. If a useful query cannot be formed without non-public data, do not call this tool.""," &
        """parameters"":{""type"":""object"",""properties"":{" &
        """query"":{""type"":""string"",""description"":""The search query. MUST NOT contain personal data, confidential details, or any non-public information. Use only generic, anonymized, or publicly known terms.""}," &
        """max_results"":{""type"":""integer"",""description"":""Maximum number of search result pages to retrieve (default: 4, server-capped).""}," &
        """max_depth"":{""type"":""integer"",""description"":""Maximum crawl depth per result page. 0 = top-level only (default: 0, server-capped).""}},""required"":[""query""]}}"

    ''' <summary>
    ''' Human-readable tool instructions shown to the model describing how to use the internal search tool.
    ''' </summary>
    Private Const InternalSearchToolInstructionsPrompt As String =
        "internet_search: Searches the internet and returns readable text from the top result pages. " &
        "Call this tool when you need current or factual information you are not confident about. " &
        "Provide query (required string). Optionally provide max_results (integer, default 4) and max_depth (integer, default 0). " &
        "Return value includes the search query used, the URLs visited, and the page content for each qualifying result. " &
        "IMPORTANT PRIVACY CONSTRAINT: The search query is sent to an external search engine. " &
        "You MUST NOT include any personal data, confidential information, private names, " &
        "case details, contract terms, internal identifiers, email addresses, phone numbers, " &
        "account numbers, or any other non-public information in the query. " &
        "Only well-known public figures, public institutions, published legislation, " &
        "publicly available case law references, and other clearly public information may appear in queries. " &
        "If you cannot formulate a useful query without disclosing non-public information, " &
        "do NOT call this tool — instead respond based on your existing knowledge and state your uncertainty."

    ' Knowledge Store Tooling (available only when KnowledgeStorePath or KnowledgeStorePathLocal is configured)

    Private Const InternalKnowledgeToolName As String = "knowledge_search"

    Private Const InternalKnowledgeToolDefinition As String =
        "{""name"":""knowledge_search""," &
        """description"":""Searches the user's local knowledge store (a curated collection of documents such as contracts, policies, legal briefs, " &
        "manuals, and reference material) and returns the most relevant document content. Use this tool when the user's question " &
        "relates to their own documents, internal policies, past work, or reference material that would not be found on the public internet. " &
        "Do NOT use this tool for general knowledge questions or publicly available information — use your training data or internet_search instead.""," &
        """parameters"":{""type"":""object"",""properties"":{" &
        """query"":{""type"":""string"",""description"":""A natural language search query describing what information is needed from the knowledge store. " &
        "Supports optional prefixes: 'tag:tagname' to filter by tag, 'store:storename' to restrict to a specific store, " &
        "or both 'tag:tagname store:storename'. Without prefixes, all stores are searched by keyword relevance.""}," &
        """max_results"":{""type"":""integer"",""description"":""Maximum number of documents to retrieve (default: 5, max: 10).""}},""required"":[""query""]}}"

    Private Const InternalKnowledgeToolInstructionsPrompt As String =
        "knowledge_search: Searches the user's local knowledge store — a curated library of the user's own documents " &
        "(contracts, policies, briefs, manuals, templates, reference material, etc.). " &
        "Call this tool when the user's question relates to their own documents, internal policies, past work, or reference material. " &
        "Provide query (required string) describing the information needed. Optionally provide max_results (integer, default 5). " &
        "The query supports optional prefixes: 'tag:tagname' filters by document tag, 'store:storename' restricts to a specific knowledge store. " &
        "Return value is the text content of the most relevant documents, each tagged with the document name and store. " &
        "IMPORTANT: Do NOT use this tool for general knowledge or publicly available information — only for the user's own document library. " &
        "When citing information from the results, mention the document name so the user can locate the source."

    ''' <summary>Minimum characters for a search hit to be considered relevant.</summary>
    Private Const ISearch_MinChars As Integer = 500

    ''' <summary>Maximum characters per search result page (applied as a cap even with WebView2).</summary>
    Private Const ISearch_MaxChars As Integer = 4000


#Region "WebView2 Web Content Retrieval"

    ''' <summary>
    ''' Retrieves website content using WebView2 (Edge) to include JavaScript-rendered content.
    ''' A dedicated STA thread is created to host WebView2 and run a minimal message loop.
    ''' </summary>
    ''' <param name="baseUrl">Absolute HTTP/HTTPS URL to retrieve.</param>
    ''' <param name="maxChars">Maximum characters to return. Values less than or equal to 0 return full content.</param>
    ''' <returns>A task producing the extracted plain text content (possibly truncated) or an empty string.</returns>
    Private Function RetrieveWebsiteContent_WebView2(baseUrl As String,
                                                 Optional maxChars As Integer = 0,
                                                 Optional includeLinks As Boolean = False,
                                                 Optional linkExtensions As System.Collections.Generic.List(Of String) = Nothing,
                                                 Optional expandInteractiveSections As Boolean = False) As System.Threading.Tasks.Task(Of WebRetrievalResult)

        Dim tcs As New System.Threading.Tasks.TaskCompletionSource(Of WebRetrievalResult)()
        Dim normalizedExtensions As System.Collections.Generic.List(Of String) = NormalizeLinkExtensions(linkExtensions)

        Dim thread As New System.Threading.Thread(
        Sub()
            Dim result As New WebRetrievalResult() With {
                .FinalUrl = baseUrl,
                .LinksJson = "[]"
            }

            Dim form As System.Windows.Forms.Form = Nothing
            Dim webView As Microsoft.Web.WebView2.WinForms.WebView2 = Nothing
            Dim userDataFolder As String = Nothing

            Dim navigationFailed As Boolean = False
            Dim contentExtracted As Boolean = False
            Dim extractionStarted As Boolean = False

            Dim startTime As System.DateTime = System.DateTime.Now
            Dim timeout As System.TimeSpan = System.TimeSpan.FromSeconds(WebToolTimeoutSeconds)

            Try
                System.Diagnostics.Debug.WriteLine(
                    $"[WebView2] Fetching: {baseUrl} (maxChars: {If(maxChars <= 0, "unlimited", maxChars.ToString())}, includeLinks: {includeLinks}, expandInteractiveSections: {expandInteractiveSections}, timeout: {WebToolTimeoutSeconds}s)")

                If Not IsSafeWebUrl(baseUrl) Then
                    result.ErrorMessage = $"Blocked unsafe URL: {baseUrl}"
                    result.TextContent = result.ErrorMessage
                    navigationFailed = True
                    Return
                End If

                userDataFolder = SharedLibrary.SharedLibrary.SharedMethods.CreateWebView2TransientFolder()

                form = New System.Windows.Forms.Form() With {
                    .Width = 1920,
                    .Height = 4000,
                    .ShowInTaskbar = False,
                    .FormBorderStyle = System.Windows.Forms.FormBorderStyle.None,
                    .StartPosition = System.Windows.Forms.FormStartPosition.Manual,
                    .Location = New System.Drawing.Point(-32000, -32000),
                    .Opacity = 0
                }

                webView = New Microsoft.Web.WebView2.WinForms.WebView2() With {
                    .Dock = System.Windows.Forms.DockStyle.Fill
                }

                form.Controls.Add(webView)

                AddHandler webView.CoreWebView2InitializationCompleted,
                    Sub(s, e)
                        Try
                            If Not e.IsSuccess Then
                                result.ErrorMessage = $"WebView2 initialization failed: {If(e.InitializationException IsNot Nothing, e.InitializationException.Message, "Unknown error")}"
                                result.TextContent = result.ErrorMessage
                                System.Diagnostics.Debug.WriteLine($"[WebView2] {result.ErrorMessage}")
                                navigationFailed = True
                                Return
                            End If

                            System.Diagnostics.Debug.WriteLine("[WebView2] CoreWebView2 initialized")

                            With webView.CoreWebView2.Settings
                                .AreDefaultScriptDialogsEnabled = False
                                .AreDefaultContextMenusEnabled = False
                                .AreDevToolsEnabled = False
                                .IsStatusBarEnabled = False
                                .IsScriptEnabled = True
                                .IsBuiltInErrorPageEnabled = False

                                ' IMPORTANT:
                                ' This must be True because the extraction script returns its result
                                ' via chrome.webview.postMessage(...).
                                .IsWebMessageEnabled = True
                            End With

                            AddHandler webView.CoreWebView2.WebMessageReceived,
                                Sub(sender, args)
                                    Try
                                        Dim rawJson As String = args.WebMessageAsJson

                                        System.Diagnostics.Debug.WriteLine("[WebView2] WebMessage received: " & If(rawJson, "<null>"))

                                        If System.String.IsNullOrWhiteSpace(rawJson) Then
                                            Return
                                        End If

                                        Dim payload As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(rawJson)

                                        Dim errorToken As Newtonsoft.Json.Linq.JToken = payload("error")
                                        If errorToken IsNot Nothing AndAlso Not System.String.IsNullOrWhiteSpace(errorToken.ToString()) Then
                                            result.ErrorMessage = errorToken.ToString()
                                            System.Diagnostics.Debug.WriteLine("[WebView2] Script error: " & result.ErrorMessage)
                                        End If

                                        Dim sourceToken As Newtonsoft.Json.Linq.JToken = payload("source_url")
                                        If sourceToken IsNot Nothing Then
                                            result.FinalUrl = sourceToken.ToString()
                                        End If

                                        Dim textToken As Newtonsoft.Json.Linq.JToken = payload("text")
                                        If textToken IsNot Nothing Then
                                            result.TextContent = textToken.ToString()
                                        End If

                                        Dim linksToken As Newtonsoft.Json.Linq.JToken = payload("links")
                                        If linksToken IsNot Nothing AndAlso linksToken.Type = Newtonsoft.Json.Linq.JTokenType.Array Then
                                            result.LinksJson = linksToken.ToString(Newtonsoft.Json.Formatting.None)
                                        End If

                                        Dim hasText As Boolean = Not System.String.IsNullOrWhiteSpace(result.TextContent)
                                        Dim hasLinks As Boolean = False

                                        Try
                                            Dim parsedLinks As Newtonsoft.Json.Linq.JToken =
                                                Newtonsoft.Json.Linq.JToken.Parse(If(result.LinksJson, "[]"))

                                            hasLinks =
                                                parsedLinks.Type = Newtonsoft.Json.Linq.JTokenType.Array AndAlso
                                                DirectCast(parsedLinks, Newtonsoft.Json.Linq.JArray).Count > 0
                                        Catch ex As System.Exception
                                            System.Diagnostics.Debug.WriteLine("[WebView2] Link parse check failed: " & ex.Message)
                                        End Try

                                        System.Diagnostics.Debug.WriteLine(
                                            $"[WebView2] Extract result: text={If(result.TextContent, "").Length} chars, links={If(hasLinks, "yes", "no")}, elapsed={(System.DateTime.Now - startTime).TotalSeconds:F1}s")

                                        If hasText OrElse includeLinks OrElse Not System.String.IsNullOrWhiteSpace(result.ErrorMessage) Then
                                            contentExtracted = True
                                        End If

                                    Catch ex As System.Exception
                                        result.ErrorMessage = "[WebView2] WebMessage parse error: " & ex.Message
                                        System.Diagnostics.Debug.WriteLine(result.ErrorMessage)
                                        contentExtracted = True
                                    End Try
                                End Sub

                            AddHandler webView.CoreWebView2.NewWindowRequested,
                                Sub(sender, args)
                                    args.Handled = True
                                End Sub

                            AddHandler webView.CoreWebView2.PermissionRequested,
                                Sub(sender, args)
                                    args.State = Microsoft.Web.WebView2.Core.CoreWebView2PermissionState.Deny
                                End Sub

                            AddHandler webView.CoreWebView2.NavigationStarting,
                                Sub(sender, args)
                                    Try
                                        Dim uriStart As System.Uri = Nothing

                                        If System.Uri.TryCreate(args.Uri, System.UriKind.Absolute, uriStart) Then
                                            If uriStart.Scheme <> System.Uri.UriSchemeHttp AndAlso uriStart.Scheme <> System.Uri.UriSchemeHttps Then
                                                args.Cancel = True
                                            End If
                                        End If
                                    Catch ex As System.Exception
                                        System.Diagnostics.Debug.WriteLine("[WebView2] NavigationStarting error: " & ex.Message)
                                    End Try
                                End Sub

                            AddHandler webView.NavigationCompleted,
                                Sub(sender, args)
                                    Try
                                        System.Diagnostics.Debug.WriteLine($"[WebView2] Navigation completed. Success: {args.IsSuccess}, Status: {args.WebErrorStatus}")

                                        If Not args.IsSuccess Then
                                            result.ErrorMessage = $"Navigation failed: {args.WebErrorStatus}"
                                            System.Diagnostics.Debug.WriteLine($"[WebView2] {result.ErrorMessage}")
                                            navigationFailed = True
                                            Return
                                        End If

                                        If extractionStarted Then
                                            Return
                                        End If

                                        extractionStarted = True

                                        Dim extractScript As String =
                                            BuildRobustWebExtractionScript(
                                                includeLinks,
                                                expandInteractiveSections,
                                                normalizedExtensions)

                                        ' IMPORTANT:
                                        ' We do NOT parse ExecuteScriptAsync's return value anymore.
                                        ' The script sends the real result through chrome.webview.postMessage(...).
                                        webView.CoreWebView2.ExecuteScriptAsync(extractScript).ContinueWith(
                                            Sub(scriptTask As System.Threading.Tasks.Task(Of String))
                                                Try
                                                    If scriptTask.IsFaulted Then
                                                        Dim message As String = "Unknown script execution error"

                                                        If scriptTask.Exception IsNot Nothing Then
                                                            message = scriptTask.Exception.GetBaseException().Message
                                                        End If

                                                        result.ErrorMessage = "[WebView2] ExecuteScriptAsync failed: " & message
                                                        System.Diagnostics.Debug.WriteLine(result.ErrorMessage)
                                                        contentExtracted = True
                                                    Else
                                                        System.Diagnostics.Debug.WriteLine("[WebView2] Extraction script started.")
                                                    End If
                                                Catch ex As System.Exception
                                                    result.ErrorMessage = "[WebView2] ExecuteScriptAsync continuation error: " & ex.Message
                                                    System.Diagnostics.Debug.WriteLine(result.ErrorMessage)
                                                    contentExtracted = True
                                                End Try
                                            End Sub)

                                    Catch ex As System.Exception
                                        result.ErrorMessage = "[WebView2] NavigationCompleted error: " & ex.Message
                                        System.Diagnostics.Debug.WriteLine(result.ErrorMessage)
                                        contentExtracted = True
                                    End Try
                                End Sub

                            System.Diagnostics.Debug.WriteLine("[WebView2] Navigating...")
                            webView.CoreWebView2.Navigate(baseUrl)

                        Catch ex As System.Exception
                            result.ErrorMessage = "[WebView2] Initialization handler error: " & ex.Message
                            result.TextContent = result.ErrorMessage
                            System.Diagnostics.Debug.WriteLine(result.ErrorMessage)
                            navigationFailed = True
                        End Try
                    End Sub

                form.Show()

                Dim envTask As System.Threading.Tasks.Task(Of Microsoft.Web.WebView2.Core.CoreWebView2Environment) =
                    Microsoft.Web.WebView2.Core.CoreWebView2Environment.CreateAsync(Nothing, userDataFolder)

                envTask.ContinueWith(
                    Sub(t As System.Threading.Tasks.Task(Of Microsoft.Web.WebView2.Core.CoreWebView2Environment))
                        Try
                            If form Is Nothing OrElse form.IsDisposed Then
                                Return
                            End If

                            form.BeginInvoke(
                                Sub()
                                    Try
                                        If t.IsCompleted AndAlso Not t.IsFaulted Then
                                            webView.EnsureCoreWebView2Async(t.Result)
                                        Else
                                            Dim message As String = "Unknown environment creation error"

                                            If t.Exception IsNot Nothing Then
                                                message = t.Exception.GetBaseException().Message
                                            End If

                                            result.ErrorMessage = "WebView2 environment creation failed: " & message
                                            result.TextContent = result.ErrorMessage
                                            System.Diagnostics.Debug.WriteLine("[WebView2] " & result.ErrorMessage)
                                            navigationFailed = True
                                        End If
                                    Catch ex As System.Exception
                                        result.ErrorMessage = "WebView2 environment BeginInvoke failed: " & ex.Message
                                        result.TextContent = result.ErrorMessage
                                        System.Diagnostics.Debug.WriteLine("[WebView2] " & result.ErrorMessage)
                                        navigationFailed = True
                                    End Try
                                End Sub)
                        Catch ex As System.Exception
                            result.ErrorMessage = "WebView2 environment continuation failed: " & ex.Message
                            result.TextContent = result.ErrorMessage
                            System.Diagnostics.Debug.WriteLine("[WebView2] " & result.ErrorMessage)
                            navigationFailed = True
                        End Try
                    End Sub)

                While Not contentExtracted AndAlso Not navigationFailed AndAlso (System.DateTime.Now - startTime) < timeout
                    System.Windows.Forms.Application.DoEvents()
                    System.Threading.Thread.Sleep(50)
                End While

                If Not contentExtracted AndAlso Not navigationFailed Then
                    result.TimedOut = True
                    result.ErrorMessage = $"Timed out after {WebToolTimeoutSeconds} seconds while retrieving rendered content from {baseUrl}."
                    System.Diagnostics.Debug.WriteLine("[WebView2] " & result.ErrorMessage)
                End If

                Dim onlyErrorOrTimeout As Boolean =
                    System.String.IsNullOrWhiteSpace(result.TextContent) OrElse
                    result.TextContent = result.ErrorMessage

                If onlyErrorOrTimeout AndAlso Not includeLinks Then
                    Try
                        System.Diagnostics.Debug.WriteLine("[WebView2] Empty rendered content; trying HTTP fallback.")

                        Using client As New System.Net.Http.HttpClient()
                            client.Timeout = System.TimeSpan.FromSeconds(WebToolTimeoutSeconds)
                            client.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0 (Windows NT 10.0; Win64; x64) RedInk/1.0")

                            Dim fallbackHtml As String = client.GetStringAsync(baseUrl).GetAwaiter().GetResult()

                            If Not System.String.IsNullOrWhiteSpace(fallbackHtml) Then
                                Dim fallbackText As String = fallbackHtml

                                fallbackText = System.Text.RegularExpressions.Regex.Replace(
                                    fallbackText,
                                    "<script[\s\S]*?</script>",
                                    "",
                                    System.Text.RegularExpressions.RegexOptions.IgnoreCase)

                                fallbackText = System.Text.RegularExpressions.Regex.Replace(
                                    fallbackText,
                                    "<style[\s\S]*?</style>",
                                    "",
                                    System.Text.RegularExpressions.RegexOptions.IgnoreCase)

                                fallbackText = System.Text.RegularExpressions.Regex.Replace(
                                    fallbackText,
                                    "<[^>]+>",
                                    " ")

                                fallbackText = System.Net.WebUtility.HtmlDecode(fallbackText)

                                fallbackText = System.Text.RegularExpressions.Regex.Replace(
                                    fallbackText,
                                    "\s+",
                                    " ").Trim()

                                If Not System.String.IsNullOrWhiteSpace(fallbackText) Then
                                    result.TextContent = fallbackText
                                    result.UsedHttpFallback = True
                                    result.ErrorMessage = ""
                                    result.TimedOut = False

                                    System.Diagnostics.Debug.WriteLine($"[WebView2] HTTP fallback extracted {result.TextContent.Length} chars.")
                                End If
                            End If
                        End Using

                    Catch ex As System.Exception
                        If System.String.IsNullOrWhiteSpace(result.ErrorMessage) Then
                            result.ErrorMessage = $"No rendered content retrieved; HTTP fallback also failed: {ex.Message}"
                        Else
                            result.ErrorMessage &= $" HTTP fallback also failed: {ex.Message}"
                        End If

                        System.Diagnostics.Debug.WriteLine("[WebView2] " & result.ErrorMessage)
                    End Try
                End If

                If System.String.IsNullOrWhiteSpace(result.TextContent) AndAlso Not System.String.IsNullOrWhiteSpace(result.ErrorMessage) Then
                    result.TextContent = result.ErrorMessage
                End If

                System.Diagnostics.Debug.WriteLine(
                    $"[WebView2] Loop ended. Content: {If(result.TextContent, "").Length} chars, elapsed: {(System.DateTime.Now - startTime).TotalSeconds:F1}s, fallback={result.UsedHttpFallback}")

            Catch ex As System.Exception
                result.ErrorMessage = $"Error retrieving {baseUrl}: {ex.Message}"
                result.TextContent = result.ErrorMessage
                System.Diagnostics.Debug.WriteLine("[WebView2] Error: " & ex.Message)

            Finally
                Try
                    If webView IsNot Nothing Then
                        webView.Dispose()
                    End If
                Catch ex As System.Exception
                    System.Diagnostics.Debug.WriteLine("[WebView2] WebView dispose error: " & ex.Message)
                End Try

                Try
                    If form IsNot Nothing Then
                        form.Close()
                        form.Dispose()
                    End If
                Catch ex As System.Exception
                    System.Diagnostics.Debug.WriteLine("[WebView2] Form dispose error: " & ex.Message)
                End Try

                Try
                    If userDataFolder IsNot Nothing AndAlso System.IO.Directory.Exists(userDataFolder) Then
                        System.IO.Directory.Delete(userDataFolder, True)
                    End If
                Catch ex As System.Exception
                    System.Diagnostics.Debug.WriteLine("[WebView2] User data folder cleanup error: " & ex.Message)
                End Try
            End Try

            Dim finalText As String = If(result.TextContent, "").Trim()

            If maxChars > 0 AndAlso finalText.Length > maxChars Then
                Dim cutPoint As Integer = maxChars
                Dim lastPeriod As Integer = finalText.LastIndexOf("."c, maxChars - 1)
                Dim lastNewline As Integer = finalText.LastIndexOf(Microsoft.VisualBasic.ControlChars.Lf, maxChars - 1)

                If lastPeriod > maxChars * 0.8 Then
                    cutPoint = lastPeriod + 1
                ElseIf lastNewline > maxChars * 0.8 Then
                    cutPoint = lastNewline
                End If

                finalText = finalText.Substring(0, cutPoint).Trim()
                System.Diagnostics.Debug.WriteLine($"[WebView2] Trimmed to {finalText.Length} chars (limit was {maxChars})")
            End If

            If System.String.IsNullOrWhiteSpace(finalText) AndAlso Not System.String.IsNullOrWhiteSpace(result.ErrorMessage) Then
                finalText = result.ErrorMessage
            End If

            result.TextContent = finalText
            tcs.TrySetResult(result)
        End Sub)

        thread.SetApartmentState(System.Threading.ApartmentState.STA)
        thread.Start()

        Return tcs.Task
    End Function


    Private Class WebRetrievalResult
        Public Property TextContent As String = ""
        Public Property FinalUrl As String = ""
        Public Property LinksJson As String = "[]"
        Public Property TimedOut As Boolean = False
        Public Property ErrorMessage As String = ""
        Public Property UsedHttpFallback As Boolean = False
    End Class

    Private Function BuildWebRetrieverFallbackNote(includeLinks As Boolean,
                                               linkExtensions As List(Of String),
                                               linksJson As String) As String
        If Not includeLinks Then
            Return ""
        End If

        Try
            Dim parsed As JToken = JToken.Parse(If(linksJson, "[]"))
            If parsed.Type = JTokenType.Array AndAlso DirectCast(parsed, JArray).Count > 0 Then
                Return ""
            End If
        Catch
        End Try

        Dim extText As String =
            If(linkExtensions Is Nothing OrElse linkExtensions.Count = 0,
               "matching links",
               String.Join(", ", linkExtensions).ToUpperInvariant() & " links")

        Return $"No {extText} were detected in the rendered DOM. " &
        "If this page computes links client-side, stores them in script state, Or reveals them only after richer interaction, " &
        "use js_run as a fallback with allow_network=true And navigate_url set to this page. " &
        "In js_run, return the final result explicitly at top level."
    End Function

    Private Function GetToolArgumentBoolean(arguments As Dictionary(Of String, Object), key As String, Optional defaultValue As Boolean = False) As Boolean
        If arguments Is Nothing OrElse Not arguments.ContainsKey(key) OrElse arguments(key) Is Nothing Then
            Return defaultValue
        End If

        Try
            Dim value = arguments(key)

            If TypeOf value Is Boolean Then
                Return CBool(value)
            End If

            If TypeOf value Is JValue Then
                Dim jv = DirectCast(value, JValue)

                If jv.Type = JTokenType.Boolean Then
                    Return jv.Value(Of Boolean)()
                End If

                Dim jvText As String = jv.ToString().Trim()
                Dim parsedJvBool As Boolean

                If Boolean.TryParse(jvText, parsedJvBool) Then
                    Return parsedJvBool
                End If

                Select Case jvText.ToLowerInvariant()
                    Case "1", "yes", "y", "on"
                        Return True
                    Case "0", "no", "n", "off"
                        Return False
                End Select
            End If

            Dim text As String = value.ToString().Trim()
            Dim parsedBool As Boolean

            If Boolean.TryParse(text, parsedBool) Then
                Return parsedBool
            End If

            Select Case text.ToLowerInvariant()
                Case "1", "yes", "y", "on"
                    Return True
                Case "0", "no", "n", "off"
                    Return False
            End Select
        Catch
        End Try

        Return defaultValue
    End Function

    Private Function GetToolArgumentStringList(arguments As Dictionary(Of String, Object), key As String) As List(Of String)
        Dim result As New List(Of String)()

        If arguments Is Nothing OrElse Not arguments.ContainsKey(key) OrElse arguments(key) Is Nothing Then
            Return result
        End If

        Try
            Dim value = arguments(key)

            If TypeOf value Is JArray Then
                For Each item In DirectCast(value, JArray)
                    Dim s As String = If(item, "").ToString().Trim()
                    If s <> "" Then result.Add(s)
                Next

                Return result
            End If

            If TypeOf value Is IEnumerable(Of Object) AndAlso Not TypeOf value Is String Then
                For Each item In DirectCast(value, IEnumerable(Of Object))
                    If item Is Nothing Then Continue For
                    Dim s As String = item.ToString().Trim()
                    If s <> "" Then result.Add(s)
                Next

                Return result
            End If

            Dim raw As String = value.ToString().Trim()
            If raw = "" Then
                Return result
            End If

            If raw.StartsWith("[") AndAlso raw.EndsWith("]") Then
                Try
                    Dim arr As JArray = JArray.Parse(raw)
                    For Each item In arr
                        Dim s As String = If(item, "").ToString().Trim()
                        If s <> "" Then result.Add(s)
                    Next

                    Return result
                Catch
                End Try
            End If

            For Each part In raw.Split({","c, ";"c}, StringSplitOptions.RemoveEmptyEntries)
                Dim s As String = part.Trim()
                If s <> "" Then result.Add(s)
            Next
        Catch
        End Try

        Return result
    End Function

    Private Function NormalizeLinkExtensions(values As IEnumerable(Of String)) As List(Of String)
        Dim result As New List(Of String)()

        If values Is Nothing Then
            Return result
        End If

        For Each value In values
            Dim normalized As String = If(value, "").Trim().TrimStart("."c).ToLowerInvariant()
            If normalized = "" Then Continue For
            If Not result.Contains(normalized, StringComparer.OrdinalIgnoreCase) Then
                result.Add(normalized)
            End If
        Next

        Return result
    End Function

    Private Function BuildWebLinkExtractionResult(requestedUrl As String,
                                                  resolvedUrl As String,
                                                  linkExtensions As List(Of String),
                                                  linksJson As String,
                                                  Optional note As String = "") As String

        Dim linksToken As JToken = New JArray()

        If Not String.IsNullOrWhiteSpace(linksJson) Then
            Try
                Dim parsed As JToken = JToken.Parse(linksJson)
                If parsed.Type = JTokenType.Array Then
                    linksToken = parsed
                End If
            Catch
            End Try
        End If

        Dim payload As New JObject(
            New JProperty("requested_url", requestedUrl),
            New JProperty("source_url", If(String.IsNullOrWhiteSpace(resolvedUrl), requestedUrl, resolvedUrl)),
            New JProperty("filters",
                New JObject(
                    New JProperty("extensions", New JArray(If(linkExtensions, New List(Of String)()).ToArray()))
                )
            ),
            New JProperty("links", linksToken)
        )

        If Not String.IsNullOrWhiteSpace(note) Then
            payload.Add("note", note)
        End If

        Return payload.ToString(Formatting.None)
    End Function

    Private Function BuildRobustWebExtractionScript(includeLinks As Boolean,
                                                expandInteractiveSections As Boolean,
                                                allowedExtensions As System.Collections.Generic.List(Of String)) As String

        Dim script As String = <![CDATA[
(async function() {
    var includeLinks = __INCLUDE_LINKS__;
    var expandInteractiveSections = __EXPAND_INTERACTIVE__;
    var allowedExtensions = __ALLOWED_EXTENSIONS__;

    function delay(ms) {
        return new Promise(function(resolve) {
            setTimeout(resolve, ms);
        });
    }

    function normalizeText(value) {
        return (value || '').replace(/\s+/g, ' ').trim();
    }

      function isAllowedResolvedUrl(url) {
        if (!url) {
            return false;
        }

        try {
            var parsed = new URL(url, document.baseURI);
            var protocol = (parsed.protocol || '').toLowerCase();
            var hostname = (parsed.hostname || '').toLowerCase();

            if (protocol !== 'http:' && protocol !== 'https:') {
                return false;
            }

            if (hostname === 'localhost' ||
                hostname === '::1' ||
                hostname === '[::1]' ||
                /^127(?:\.\d{1,3}){3}$/.test(hostname)) {
                return false;
            }

            return true;
        } catch (e) {
            return false;
        }
    }

    function resolveUrl(value) {
        if (!value) {
            return '';
        }

        value = String(value).trim();

        if (!value || value === '#' || value.toLowerCase().indexOf('javascript:') === 0) {
            return '';
        }

        try {
            var resolved = new URL(value, document.baseURI).href;
            return isAllowedResolvedUrl(resolved) ? resolved : '';
        } catch (e) {
            return '';
        }
    }

    function getExtensionFromUrl(url) {
        try {
            var parsed = new URL(url, document.baseURI);
            var pathname = parsed.pathname || '';
            var lastSegment = pathname.split('/').pop() || '';
            var dot = lastSegment.lastIndexOf('.');

            if (dot < 0) {
                return '';
            }

            return lastSegment.substring(dot + 1).toLowerCase();
        } catch (e) {
            return '';
        }
    }

    function isVisible(el) {
        if (!el) {
            return false;
        }

        try {
            if (el.hidden) {
                return false;
            }

            var style = window.getComputedStyle(el);

            if (!style) {
                return true;
            }

            if (style.display === 'none' || style.visibility === 'hidden') {
                return false;
            }

            return true;
        } catch (e) {
            return true;
        }
    }

    function queryAllDeep(selector) {
        var results = [];
        var seen = new Set();

        function visitRoot(root) {
            if (!root || seen.has(root)) {
                return;
            }

            seen.add(root);

            try {
                if (root.querySelectorAll) {
                    root.querySelectorAll(selector).forEach(function(el) {
                        results.push(el);
                    });
                }
            } catch (e) {
            }

            try {
                var all = root.querySelectorAll ? root.querySelectorAll('*') : [];

                for (var i = 0; i < all.length; i++) {
                    var el = all[i];

                    if (el && el.shadowRoot) {
                        visitRoot(el.shadowRoot);
                    }
                }
            } catch (e) {
            }
        }

        visitRoot(document);
        return results;
    }

    function getAllElementsDeep() {
        return queryAllDeep('*');
    }

    function extractUrlsFromString(value) {
        var results = [];

        if (!value) {
            return results;
        }

        var text = String(value);
        var absoluteRegex = /https?:\/\/[^\s"'<>]+/gi;
        var match;

        while ((match = absoluteRegex.exec(text)) !== null) {
            results.push(match[0]);
        }

        var relativeRegex = /["']((?:\/|\.{1,2}\/)[^"'<>]+)["']/gi;

        while ((match = relativeRegex.exec(text)) !== null) {
            results.push(match[1]);
        }

        return results;
    }

    function matchesAllowed(url, extension, hintText) {
        if (!allowedExtensions || allowedExtensions.length === 0) {
            return true;
        }

        var ext = (extension || '').toLowerCase();

        if (ext && allowedExtensions.indexOf(ext) >= 0) {
            return true;
        }

        var haystack = ((url || '') + ' ' + (hintText || '')).toLowerCase();

        for (var i = 0; i < allowedExtensions.length; i++) {
            var allowed = allowedExtensions[i];

            if (haystack.indexOf('.' + allowed) >= 0 ||
                haystack.indexOf('=' + allowed) >= 0 ||
                haystack.indexOf('/' + allowed) >= 0 ||
                haystack.indexOf('format ' + allowed) >= 0 ||
                haystack.indexOf('type ' + allowed) >= 0 ||
                haystack.indexOf(allowed) >= 0) {
                return true;
            }
        }

        return false;
    }

    function buildHintText(el, extraText) {
        if (!el) {
            return normalizeText(extraText || '');
        }

        var className = '';

        try {
            className = typeof el.className === 'string' ? el.className : '';
        } catch (e) {
            className = '';
        }

        return normalizeText([
            extraText || '',
            el.innerText || '',
            el.textContent || '',
            el.getAttribute && el.getAttribute('aria-label'),
            el.getAttribute && el.getAttribute('title'),
            el.getAttribute && el.getAttribute('type'),
            el.getAttribute && el.getAttribute('download'),
            el.id || '',
            className
        ].join(' '));
    }

    async function waitForBodyText() {
        var started = Date.now();

        while (Date.now() - started < 8000) {
            if (document.body && normalizeText(document.body.innerText || document.body.textContent || '').length > 20) {
                return;
            }

            await delay(250);
        }
    }

    async function autoScroll() {
        var body = document.body || document.documentElement;

        if (!body) {
            return;
        }

        var totalHeight = body.scrollHeight || 0;
        var viewportHeight = window.innerHeight || 1000;
        var currentPosition = 0;
        var maxScroll = 20000;

        while (currentPosition < totalHeight && currentPosition < maxScroll) {
            window.scrollTo(0, currentPosition);
            await delay(150);

            currentPosition += viewportHeight;
            totalHeight = body.scrollHeight || totalHeight;
        }

        window.scrollTo(0, 0);
        await delay(250);
    }

    function clickIfExpandable(el) {
        if (!el) {
            return false;
        }

        if (!isVisible(el) && !expandInteractiveSections) {
            return false;
        }

        var tag = (el.tagName || '').toUpperCase();
        var ariaExpanded = (el.getAttribute && el.getAttribute('aria-expanded') || '').toLowerCase();
        var dataBsToggle = (el.getAttribute && el.getAttribute('data-bs-toggle') || '').toLowerCase();
        var dataToggle = (el.getAttribute && el.getAttribute('data-toggle') || '').toLowerCase();
        var text = normalizeText(el.innerText || el.textContent || '');

        var shouldClick =
            tag === 'SUMMARY' ||
            ariaExpanded === 'false' ||
            dataBsToggle === 'collapse' ||
            dataBsToggle === 'dropdown' ||
            dataToggle === 'collapse' ||
            dataToggle === 'dropdown' ||
            /\b(expand|show more|read more|open|attachments|documents|downloads|resources)\b/i.test(text) ||
            (el.classList && (
                el.classList.contains('accordion-button') ||
                el.classList.contains('accordion-trigger') ||
                el.classList.contains('expander')
            ));

        if (!shouldClick) {
            return false;
        }

        try {
            el.click();
            return true;
        } catch (e) {
            return false;
        }
    }

    async function expandSections() {
        if (!expandInteractiveSections) {
            return;
        }

        for (var pass = 0; pass < 5; pass++) {
            var clicked = 0;

            queryAllDeep('details').forEach(function(detailsEl) {
                if (!detailsEl.open) {
                    try {
                        detailsEl.open = true;
                        clicked++;
                    } catch (e) {
                    }
                }
            });

            var selectors = [
                'summary',
                '[aria-expanded="false"]',
                '[data-bs-toggle="collapse"]',
                '[data-toggle="collapse"]',
                '[data-bs-toggle="dropdown"]',
                '[data-toggle="dropdown"]',
                '.accordion-button',
                '.accordion-trigger',
                '.expander',
                'button[aria-expanded="false"]',
                '[role="button"][aria-expanded="false"]',
                '[aria-controls]'
            ];

            queryAllDeep(selectors.join(',')).forEach(function(el) {
                if (clickIfExpandable(el)) {
                    clicked++;
                }
            });

            if (clicked === 0) {
                break;
            }

            await delay(600);
            await autoScroll();
        }
    }

    function walk(node) {
        if (!node) {
            return '';
        }

        if (node.nodeType === 3) {
            return node.textContent || '';
        }

        if (node.nodeType !== 1) {
            return '';
        }

        var tag = node.tagName ? node.tagName.toUpperCase() : '';

        if (tag === 'SCRIPT' || tag === 'STYLE' || tag === 'NOSCRIPT') {
            return '';
        }

        if (!isVisible(node)) {
            return '';
        }

        var parts = [];

        if (node.shadowRoot && node.shadowRoot.childNodes) {
            for (var s = 0; s < node.shadowRoot.childNodes.length; s++) {
                parts.push(walk(node.shadowRoot.childNodes[s]));
            }
        }

        for (var i = 0; i < node.childNodes.length; i++) {
            parts.push(walk(node.childNodes[i]));
        }

        var inner = parts.join('');

        if (tag === 'A') {
            var href = resolveUrl(node.getAttribute('href') || '');
            var text = inner.trim();

            if (href && text) {
                if (text === href || text === decodeURIComponent(href)) {
                    return text;
                }

                return '[' + text + '](' + href + ')';
            }

            return text || '';
        }

        if (/^(DIV|P|BR|H[1-6]|LI|TR|BLOCKQUOTE|SECTION|ARTICLE|ASIDE|MAIN|DT|DD|FIGCAPTION|PRE|HEADER|FOOTER|NAV)$/.test(tag)) {
            if (tag === 'BR') {
                return '\n';
            }

            return '\n' + inner + '\n';
        }

        return inner;
    }

    function collectText() {
        var root =
            document.querySelector('main') ||
            document.querySelector('[role="main"]') ||
            document.body;

        var text = root ? walk(root) : '';

        if (!normalizeText(text) && document.body) {
            text = document.body.innerText || document.body.textContent || '';
        }

        text = text
            .replace(/\r/g, '\n')
            .replace(/\n{3,}/g, '\n\n')
            .replace(/[ \t]+/g, ' ')
            .trim();

        return text;
    }

    function collectLinks() {
        if (!includeLinks) {
            return [];
        }

        var links = [];
        var seen = new Set();

        function addCandidate(url, source, attributeName, el, explicitText) {
            var resolved = resolveUrl(url);

            if (!resolved) {
                return;
            }

            var extension = getExtensionFromUrl(resolved);
            var hintText = buildHintText(el, explicitText);

            if (!matchesAllowed(resolved, extension, hintText)) {
                return;
            }

            var key = resolved.toLowerCase();

            if (seen.has(key)) {
                return;
            }

            seen.add(key);

            links.push({
                text: hintText || resolved,
                url: resolved,
                extension: extension,
                download: !!(el && el.hasAttribute && el.hasAttribute('download')) || extension !== '',
                source: source || '',
                attribute: attributeName || '',
                visible: el ? isVisible(el) : true
            });
        }

        queryAllDeep('a[href], area[href]').forEach(function(el) {
            addCandidate(el.getAttribute('href'), 'anchor', 'href', el, '');
        });

        var attributeNames = [
            'href',
            'src',
            'data',
            'data-href',
            'data-url',
            'data-link',
            'data-src',
            'data-download',
            'data-download-url',
            'data-file',
            'data-file-url',
            'data-doc-url',
            'data-document-url'
        ];

        getAllElementsDeep().forEach(function(el) {
            for (var i = 0; i < attributeNames.length; i++) {
                var attrName = attributeNames[i];
                var attrValue = el.getAttribute && el.getAttribute(attrName);

                if (attrValue) {
                    addCandidate(attrValue, 'attribute', attrName, el, '');
                }
            }

            var scriptLikeAttrs = ['onclick', 'onmousedown', 'onmouseup', 'data-onclick'];

            for (var j = 0; j < scriptLikeAttrs.length; j++) {
                var scriptAttr = scriptLikeAttrs[j];
                var raw = el.getAttribute && el.getAttribute(scriptAttr);

                if (!raw) {
                    continue;
                }

                extractUrlsFromString(raw).forEach(function(foundUrl) {
                    addCandidate(foundUrl, 'script-attribute', scriptAttr, el, raw);
                });
            }
        });

        extractUrlsFromString(document.body ? document.body.innerText : '').forEach(function(foundUrl) {
            addCandidate(foundUrl, 'body-text', 'text', null, foundUrl);
        });

        return links;
    }

    function sendResult(payload) {
        if (window.chrome && chrome.webview && chrome.webview.postMessage) {
            chrome.webview.postMessage(payload);
        }
    }

    try {
        await waitForBodyText();
        await autoScroll();
        await expandSections();
        await autoScroll();
        await delay(500);

        sendResult({
            source_url: document.baseURI || location.href,
            title: document.title || '',
            text: collectText(),
            links: collectLinks()
        });
    } catch (err) {
        sendResult({
            source_url: document.baseURI || location.href,
            title: document.title || '',
            text: document.body ? (document.body.innerText || document.body.textContent || '') : '',
            links: [],
            error: err && err.message ? err.message : String(err)
        });
    }
})();
]]>.Value

        script = script.Replace("__INCLUDE_LINKS__", If(includeLinks, "true", "false"))
        script = script.Replace("__EXPAND_INTERACTIVE__", If(expandInteractiveSections, "true", "false"))
        script = script.Replace("__ALLOWED_EXTENSIONS__", Newtonsoft.Json.JsonConvert.SerializeObject(If(allowedExtensions, New System.Collections.Generic.List(Of String)())))

        Return script
    End Function


    ''' <summary>
    ''' Determines whether a URL is allowed to be retrieved by the internal web tool.
    ''' Only absolute HTTP/HTTPS URLs are permitted; loopback/localhost is denied.
    ''' </summary>
    ''' <param name="url">Candidate URL.</param>
    ''' <returns><c>True</c> if the URL is considered safe; otherwise <c>False</c>.</returns>

    Private Function IsSafeWebUrl(url As String) As Boolean
        Try
            Dim uriResult As Uri = Nothing
            If Not Uri.TryCreate(url, UriKind.Absolute, uriResult) Then Return False

            ' 1. Only allow HTTP and HTTPS
            If uriResult.Scheme <> Uri.UriSchemeHttp AndAlso uriResult.Scheme <> Uri.UriSchemeHttps Then
                Return False
            End If

            ' 2. Basic SSRF check: Block localhost / loopback
            If uriResult.IsLoopback Then Return False
            If uriResult.Host.ToLower().Equals("localhost") Then Return False

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Unescapes the JSON-encoded string returned by WebView2 <c>ExecuteScriptAsync</c>.
    ''' </summary>
    ''' <param name="jsonString">JSON string value returned by script execution.</param>
    ''' <returns>Unescaped string content.</returns>
    Private Function UnescapeJsonString(jsonString As String) As String
        If String.IsNullOrEmpty(jsonString) Then Return ""

        ' Remove surrounding quotes
        If jsonString.StartsWith("""") AndAlso jsonString.EndsWith("""") Then
            jsonString = jsonString.Substring(1, jsonString.Length - 2)
        End If

        ' Handle null
        If jsonString = "null" Then Return ""

        ' Unescape common sequences
        Try
            jsonString = jsonString.Replace("\n", vbLf)
            jsonString = jsonString.Replace("\r", vbCr)
            jsonString = jsonString.Replace("\t", vbTab)
            jsonString = jsonString.Replace("\\", "\")
            jsonString = jsonString.Replace("\""", """")

            ' Handle unicode escapes like \u0027
            Dim unicodePattern As String = "\\u([0-9A-Fa-f]{4})"
            jsonString = Regex.Replace(jsonString, unicodePattern,
                Function(m) ChrW(System.Convert.ToInt32(m.Groups(1).Value, 16)).ToString())

        Catch ex As Exception
            Debug.WriteLine($"[WebView2] Unescape error: {ex.Message}")
        End Try

        Return jsonString
    End Function

    ''' <summary>
    ''' Retrieves website content using an HTTP client (non-WebView2 fallback).
    ''' </summary>
    ''' <param name="url">Absolute URL to retrieve.</param>
    ''' <param name="maxDepth">Unused by this implementation; retained for caller compatibility.</param>
    ''' <param name="client">HTTP client used for retrieval.</param>
    ''' <returns>Response body as a string, or an error string on failure.</returns>
    Private Async Function RetrieveWebsiteContent(url As String, maxDepth As Integer, client As System.Net.Http.HttpClient) As Task(Of String)
        Try
            client.Timeout = TimeSpan.FromSeconds(WebToolTimeoutSeconds)

            If Not client.DefaultRequestHeaders.UserAgent.Any() Then
                client.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0 (Windows NT 10.0; Win64; x64) RedInk/1.0")
            End If

            Return Await client.GetStringAsync(url)
        Catch ex As TaskCanceledException
            Return $"Timed out after {WebToolTimeoutSeconds} seconds while retrieving {url}."
        Catch ex As OperationCanceledException
            Return $"Timed out after {WebToolTimeoutSeconds} seconds while retrieving {url}."
        Catch ex As Exception
            Return $"Error retrieving {url}: {ex.Message}"
        End Try
    End Function

#End Region






    ''' <summary>
    ''' Executes the internal web retrieval tool by fetching content for one or more URLs and returning tagged content blocks.
    ''' </summary>
    ''' <param name="toolCall">Tool call payload containing <c>url</c> or <c>urls</c> arguments.</param>
    ''' <param name="context">Tool execution context for logging and diagnostics.</param>
    ''' <param name="cancellationToken">Optional cancellation token to abort execution.</param>
    ''' <returns>Tool response containing retrieved content or an error.</returns>
    Private Async Function ExecuteInternalWebTool(toolCall As ToolCall, context As ToolExecutionContext, Optional cancellationToken As System.Threading.CancellationToken = Nothing) As Task(Of ToolResponse)
        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId,
            .ToolName = toolCall.ToolName
        }

        Try
            cancellationToken.ThrowIfCancellationRequested()

            Dim urls As New List(Of String)()

            If toolCall.Arguments.ContainsKey("urls") Then
                Dim urlsArg = toolCall.Arguments("urls")
                If TypeOf urlsArg Is JArray Then
                    For Each item In DirectCast(urlsArg, JArray)
                        urls.Add(item.ToString())
                    Next
                ElseIf TypeOf urlsArg Is IEnumerable(Of Object) Then
                    For Each item In DirectCast(urlsArg, IEnumerable(Of Object))
                        urls.Add(item.ToString())
                    Next
                ElseIf TypeOf urlsArg Is String Then
                    urls.Add(urlsArg.ToString())
                End If
            ElseIf toolCall.Arguments.ContainsKey("url") Then
                urls.Add(toolCall.Arguments("url").ToString())
            End If

            If urls.Count = 0 Then
                response.Success = False
                response.ErrorMessage = "No URLs provided"
                ToolingFileLogger.LogWarn("Internal web tool: No URLs provided.", details:=$"CallId={toolCall.CallId}; Args={JsonConvert.SerializeObject(toolCall.Arguments)}")
                Return response
            End If

            Dim includeLinks As Boolean = GetToolArgumentBoolean(toolCall.Arguments, "include_links", False)
            Dim expandInteractiveSections As Boolean = GetToolArgumentBoolean(toolCall.Arguments, "expand_interactive_sections", False)
            Dim linkExtensions As List(Of String) = NormalizeLinkExtensions(GetToolArgumentStringList(toolCall.Arguments, "link_extensions"))

            Dim invalidUrls As New List(Of String)()
            Dim sharepointPatterns As String() = {"sharepoint.com", "onedrive.com", "1drv.ms", "teams.microsoft.com", ":f:/", "/:f:/"}
            Dim blockedUrls As New List(Of String)()

            For Each url In urls
                If Not IsSafeWebUrl(url) Then
                    invalidUrls.Add(url)
                    Continue For
                End If

                Dim lowerUrl = url.ToLowerInvariant()
                For Each pattern In sharepointPatterns
                    If lowerUrl.Contains(pattern) Then
                        blockedUrls.Add(url)
                        Exit For
                    End If
                Next
            Next

            If invalidUrls.Count > 0 Then
                Dim invalidList = String.Join(", ", invalidUrls)
                response.Success = False
                response.ErrorMessage =
                    "Only absolute non-loopback HTTP/HTTPS URLs are allowed. Blocked URL(s): " & invalidList
                context.Log($"Blocked unsafe URL(s): {invalidList}", "warn")
                ToolingFileLogger.LogWarn("Internal web tool: Unsafe URL blocked.", details:=$"urls={invalidList}")
                Return response
            End If

            If blockedUrls.Count > 0 Then
                Dim blockedList = String.Join(", ", blockedUrls)
                response.Success = False
                response.ErrorMessage =
                    $"Cannot retrieve content from the following URL(s) because they point to SharePoint, OneDrive, or Microsoft Teams — " &
                    $"these are authenticated cloud storage resources that require login and cannot be accessed remotely: {blockedList}. " &
                    "Please ask the user to download the file(s) and attach them directly to the e-mail."
                context.Log($"Blocked SharePoint/OneDrive URL(s): {blockedList}", "warn")
                ToolingFileLogger.LogWarn("Internal web tool: SharePoint/OneDrive URL blocked.", details:=$"urls={blockedList}")
                Return response
            End If

            context.Log($"Retrieving content from {urls.Count} URL(s)...")

            If includeLinks Then
                context.Log($"  Structured link extraction enabled (extensions: {If(linkExtensions.Count = 0, "all", String.Join(", ", linkExtensions))}; expand interactive sections: {expandInteractiveSections})")
            End If

            Dim results As New StringBuilder()

            If UseWebView2 Then
                For i = 0 To urls.Count - 1
                    cancellationToken.ThrowIfCancellationRequested()

                    Dim requestedUrl = urls(i)

                    Try
                        context.Log($"  Fetching: {requestedUrl}")

                        Dim pageResult = Await RetrieveWebsiteContent_WebView2(
                            requestedUrl,
                            0,
                            includeLinks,
                            linkExtensions,
                            expandInteractiveSections)

                        Dim resolvedUrl As String = If(pageResult?.FinalUrl, requestedUrl)
                        Dim content As String = If(pageResult?.TextContent, "")
                        Dim linksJson As String = If(pageResult?.LinksJson, "[]")

                        results.AppendLine($"<URL_{i + 1}>{resolvedUrl}</URL_{i + 1}>")

                        If Not String.IsNullOrWhiteSpace(content) Then
                            results.AppendLine($"<CONTENT_{i + 1}>")
                            results.AppendLine(content)
                            results.AppendLine($"</CONTENT_{i + 1}>")
                        Else
                            results.AppendLine($"<CONTENT_{i + 1}>No content retrieved</CONTENT_{i + 1}>")
                            ToolingFileLogger.LogWarn("Internal web tool: No content retrieved.", details:=$"url={requestedUrl}")
                        End If

                        If includeLinks Then
                            results.AppendLine($"<LINKS_{i + 1}>")
                            results.AppendLine(
                                BuildWebLinkExtractionResult(
                                    requestedUrl,
                                    resolvedUrl,
                                    linkExtensions,
                                    linksJson,
                                    BuildWebRetrieverFallbackNote(includeLinks, linkExtensions, linksJson)))
                            results.AppendLine($"</LINKS_{i + 1}>")
                        End If

                        results.AppendLine()

                    Catch ex As OperationCanceledException
                        Throw
                    Catch ex As Exception
                        results.AppendLine($"<URL_{i + 1}>{requestedUrl}</URL_{i + 1}>")
                        results.AppendLine($"<ERROR_{i + 1}>{ex.Message}</ERROR_{i + 1}>")

                        If includeLinks Then
                            results.AppendLine($"<LINKS_{i + 1}>")
                            results.AppendLine(
                                BuildWebLinkExtractionResult(
                                    requestedUrl,
                                    requestedUrl,
                                    linkExtensions,
                                    "[]",
                                    "Link extraction failed because page retrieval failed."))
                            results.AppendLine($"</LINKS_{i + 1}>")
                        End If

                        results.AppendLine()
                        ToolingFileLogger.LogError("Internal web tool fetch error.", details:=$"url={requestedUrl}", ex:=ex)
                    End Try
                Next
            Else
                Using httpClient As New System.Net.Http.HttpClient()
                    httpClient.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
                    httpClient.Timeout = TimeSpan.FromSeconds(30)

                    For i = 0 To urls.Count - 1
                        cancellationToken.ThrowIfCancellationRequested()

                        Dim requestedUrl = urls(i)

                        Try
                            context.Log($"  Fetching: {requestedUrl}")
                            Dim content = Await RetrieveWebsiteContent(requestedUrl, INI_ISearch_MaxDepth, httpClient)

                            results.AppendLine($"<URL_{i + 1}>{requestedUrl}</URL_{i + 1}>")

                            If Not String.IsNullOrWhiteSpace(content) Then
                                results.AppendLine($"<CONTENT_{i + 1}>")
                                results.AppendLine(content)
                                results.AppendLine($"</CONTENT_{i + 1}>")
                            Else
                                results.AppendLine($"<CONTENT_{i + 1}>No content retrieved</CONTENT_{i + 1}>")
                                ToolingFileLogger.LogWarn("Internal web tool: No content retrieved.", details:=$"url={requestedUrl}")
                            End If

                            If includeLinks Then
                                results.AppendLine($"<LINKS_{i + 1}>")
                                results.AppendLine(
                                    BuildWebLinkExtractionResult(
                                        requestedUrl,
                                        requestedUrl,
                                        linkExtensions,
                                        "[]",
                                        "Structured link extraction requires WebView2 and is unavailable in the HTTP fallback path."))
                                results.AppendLine($"</LINKS_{i + 1}>")
                            End If

                            results.AppendLine()

                        Catch ex As OperationCanceledException
                            Throw
                        Catch ex As Exception
                            results.AppendLine($"<URL_{i + 1}>{requestedUrl}</URL_{i + 1}>")
                            results.AppendLine($"<ERROR_{i + 1}>{ex.Message}</ERROR_{i + 1}>")

                            If includeLinks Then
                                results.AppendLine($"<LINKS_{i + 1}>")
                                results.AppendLine(
                                    BuildWebLinkExtractionResult(
                                        requestedUrl,
                                        requestedUrl,
                                        linkExtensions,
                                        "[]",
                                        "Link extraction failed because page retrieval failed."))
                                results.AppendLine($"</LINKS_{i + 1}>")
                            End If

                            results.AppendLine()
                            ToolingFileLogger.LogError("Internal web tool fetch error.", details:=$"url={requestedUrl}", ex:=ex)
                        End Try
                    Next
                End Using
            End If

            response.Response = results.ToString()
            response.Success = True

        Catch ex As OperationCanceledException
            response.Success = False
            response.ErrorMessage = "Operation was cancelled"
            ToolingFileLogger.LogWarn("Internal web tool cancelled.")

        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            ToolingFileLogger.LogError("Internal web tool error.", ex:=ex)
        End Try

        Return response
    End Function




    Private Function BuildToolWorkflowInstructionAddendum(selectedTools As List(Of ModelConfig)) As String
        Dim sb As New StringBuilder()

        sb.AppendLine("PERSISTENCE CHECKLIST:")
        sb.AppendLine("- Remain in tool-calling mode until the whole user request is completed. Do not stop after planning, discovering files, staging files, or finishing only the first subtask.")
        sb.AppendLine("- If one tool fails, returns too little information, or only partially advances the task, and another available tool could still help, call the next suitable tool instead of giving up.")
        sb.AppendLine("- If the request applies to a folder, directory, workspace path, or a collection of files, discover or stage the collection first and then continue processing the returned items until the collection has actually been searched or analyzed.")
        sb.AppendLine("- Before giving a final answer, explicitly check whether any requested next step, remaining file, or reasonable fallback tool is still outstanding.")

        If HasToolName(selectedTools, SharedLibrary.Agents.JsRunTool.ToolName) Then
            sb.AppendLine("- For deterministic text or data operations, prefer js_run over estimating in prose.")
            sb.AppendLine("- Use js_run for exact word counts, character counts, line counts, regex extraction/counting, delimiter-based splitting, deduplication, sorting, JSON reshaping, and similar programmable transformations.")
            sb.AppendLine("- When the user asks to count, enumerate, classify by exact rule, or transform text mechanically, compute the result with js_run and return the computed value explicitly.")
        End If

        If HasToolName(selectedTools, "extract_pdf_text") Then
            sb.AppendLine("- extract_pdf_text is for a single PDF or staged/session file at a time. Never pass a directory or folder path to extract_pdf_text.")
        End If

        If HasToolName(selectedTools, "agent_workspace_find_files") OrElse
       HasToolName(selectedTools, "agent_workspace_stage") OrElse
       HasToolName(selectedTools, "agent_workspace_read") OrElse
       HasToolName(selectedTools, "workspace_inventory") OrElse
       HasToolName(selectedTools, "workspace_read") Then
            sb.AppendLine("- For local/workspace PDF collections, prefer the workspace workflow: find files, stage them if required, then read/search/extract them. Do not stop after file discovery.")
        End If

        If HasToolName(selectedTools, "agent_workspace_read") Then
            sb.AppendLine("- For one workspace-local PDF or Office file, prefer agent_workspace_read over calling extract_pdf_text directly on a workspace path.")
        End If

        If HasToolName(selectedTools, "search_in_attachments") Then
            sb.AppendLine("- When many staged PDFs must be searched for a term, prefer search_in_attachments across the staged set before falling back to repeated one-file extraction calls.")
        End If

        Return sb.ToString().Trim()
    End Function



    ''' <summary>
    ''' Creates a built-in internal internet search tool configuration as a <see cref="ModelConfig"/>.
    ''' Only meaningful when <c>INI_ISearch</c> is enabled and <c>INI_ISearch_URL</c> is configured.
    ''' </summary>
    ''' <param name="enforcePrivacy">When True, privacy constraints are included in the tool definition and instructions.</param>
    ''' <returns>Internal search tool configuration.</returns>
    Public Function GetInternalSearchTool(Optional enforcePrivacy As Boolean = True) As ModelConfig
        Dim definition As String = InternalSearchToolDefinition
        Dim instructions As String = InternalSearchToolInstructionsPrompt

        If Not enforcePrivacy Then
            ' Strip privacy constraints from the tool definition
            definition =
                "{""name"":""internet_search""," &
                """description"":""Searches the internet via the configured search engine, retrieves the top result pages, and returns their readable text content. Use this when you need up-to-date or factual information you are not confident about.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """query"":{""type"":""string"",""description"":""The search query.""}," &
                """max_results"":{""type"":""integer"",""description"":""Maximum number of search result pages to retrieve (default: 4, server-capped).""}," &
                """max_depth"":{""type"":""integer"",""description"":""Maximum crawl depth per result page. 0 = top-level only (default: 0, server-capped).""}},""required"":[""query""]}}"

            instructions =
                "internet_search: Searches the internet and returns readable text from the top result pages. " &
                "Call this tool when you need current or factual information you are not confident about. " &
                "Provide query (required string). Optionally provide max_results (integer, default 4) and max_depth (integer, default 0). " &
                "Return value includes the search query used, the URLs visited, and the page content for each qualifying result."
        End If

        Return New ModelConfig() With {
            .ToolName = InternalSearchToolName,
            .ToolInstructionsPrompt = instructions,
            .ToolDefinition = definition,
            .ModelDescription = "Internet Search (" & If(Not String.IsNullOrWhiteSpace(INI_ISearch_Name), INI_ISearch_Name, "Search") & ")" & InternalToolSuffix,
            .Tool = True,
            .ToolPriority = 998,
            .ToolErrorHandling = "skip"
        }
    End Function


    Public Function GetInternalWebGroundingTool(Optional enforcePrivacy As Boolean = True) As ModelConfig
        Return SharedLibrary.Agents.WebGroundingTool.Build(
        _context,
        enforcePrivacy:=enforcePrivacy,
        toolPriority:=997,
        displaySuffix:=InternalToolSuffix)
    End Function


    ''' <summary>
    ''' Creates a built-in internal web retrieval tool configuration as a <see cref="ModelConfig"/>.
    ''' </summary>
    ''' <returns>Internal tool configuration.</returns>
    Public Function GetInternalWebTool() As ModelConfig
        Return New ModelConfig() With {
            .ToolName = InternalWebToolName,
            .ToolInstructionsPrompt = InternalWebToolInstructionsPrompt,
            .ToolDefinition = InternalWebToolDefinition,
            .ModelDescription = "Web Content Retriever" & InternalToolSuffix,
            .Tool = True,
            .ToolPriority = 999,
            .ToolErrorHandling = "skip"
        }
    End Function

    Public Function GetInternalDownloadWebFilesTool() As ModelConfig
        Return New ModelConfig() With {
            .ToolName = InternalDownloadWebFilesToolName,
            .ToolInstructionsPrompt = InternalDownloadWebFilesToolInstructionsPrompt,
            .ToolDefinition = InternalDownloadWebFilesToolDefinition,
            .ModelDescription = "Web File Downloader" & InternalToolSuffix,
            .Tool = True,
            .ToolPriority = 996,
            .ToolErrorHandling = "skip"
        }
    End Function

    ''' <summary>Internal tool name used by the model to announce a major step (B1).</summary>
    Public Const InternalProgressReportToolName As String = "report_progress"

    ''' <summary>Returns True when the tool name is the internal report_progress tool.</summary>
    Public Function IsReportProgressToolName(toolName As String) As Boolean
        Return Not String.IsNullOrWhiteSpace(toolName) AndAlso
               toolName.Trim().Equals(InternalProgressReportToolName, StringComparison.OrdinalIgnoreCase)
    End Function

    ''' <summary>
    ''' Model-facing instruction describing when to call report_progress. Major steps only;
    ''' the note is authored in the dialogue language; it may be batched with real tool calls.
    ''' </summary>
    Public Function BuildMajorStepProgressToolInstruction() As String
        Return "PROGRESS REPORTING RULE: You MUST call the report_progress tool before the first substantive tool action of every tooling run and again whenever you begin a materially different major phase of work. Do this whether or not the user explicitly asked for progress updates. Use exactly one short, user-facing sentence in the dialogue language. Call it only for major steps, not for every minor step. You may include report_progress in the same turn as the real tool calls that follow it. Do NOT write progress as plain text. Do NOT serialize tool calls as text or JSON text. Use only native tool calls. A major work phase must not begin silently without a preceding report_progress call."
    End Function

    ''' <summary>Canonical JSON definition for the internal report_progress tool.</summary>
    Private Function BuildInternalProgressReportToolDefinition() As String
        Dim definition As New JObject(
            New JProperty("name", InternalProgressReportToolName),
            New JProperty("description",
                "Announces a single major step to the user in the dialogue language. " &
                "Use only at the start of a significant new phase of work, not on every turn. " &
                "Performs no other work and returns immediately."),
            New JProperty("parameters",
                New JObject(
                    New JProperty("type", "object"),
                    New JProperty("properties",
                        New JObject(
                            New JProperty("note",
                                New JObject(
                                    New JProperty("type", "string"),
                                    New JProperty("description", "One short, user-facing sentence in the dialogue language describing the major step you are about to start. Do not reveal tool names, arguments, or internal reasoning.")
                                )
                            )
                        )
                    ),
                    New JProperty("required", New JArray("note")),
                    New JProperty("additionalProperties", False)
                )
            )
        )

        Return definition.ToString(Formatting.None)
    End Function

    ''' <summary>Creates the built-in report_progress tool configuration (B1 major-step channel).</summary>
    Public Function GetInternalProgressReportTool() As ModelConfig
        Return New ModelConfig() With {
            .ToolName = InternalProgressReportToolName,
            .ToolInstructionsPrompt = BuildMajorStepProgressToolInstruction(),
            .ToolDefinition = BuildInternalProgressReportToolDefinition(),
            .ModelDescription = "Progress Update" & InternalToolSuffix,
            .Tool = True,
            .ToolPriority = 995,
            .ToolErrorHandling = "skip"
        }
    End Function

    Private Function GetSafeDownloadRoot() As String
        Try
            If _chatAgentActive AndAlso Not _apActive Then
                LoadChatAgentWorkspaceIfNeeded()

                If _chatAgentWorkspace IsNot Nothing AndAlso
                   _chatAgentWorkspace.AllowWrite AndAlso
                   Not String.IsNullOrWhiteSpace(_chatAgentWorkspace.RootPath) AndAlso
                   Directory.Exists(_chatAgentWorkspace.RootPath) Then
                    Return Path.GetFullPath(_chatAgentWorkspace.RootPath)
                End If
            End If
        Catch
        End Try

        Try
            Dim ws = SharedLibrary.Agents.WorkspaceTools.Active

            If ws IsNot Nothing AndAlso
               ws.AllowWrite AndAlso
               Not String.IsNullOrWhiteSpace(ws.RootPath) AndAlso
               Directory.Exists(ws.RootPath) Then
                Return Path.GetFullPath(ws.RootPath)
            End If
        Catch
        End Try

        Try
            Dim policyRoot = SharedLibrary.Agents.PathPolicy.WorkspaceRoot

            If Not String.IsNullOrWhiteSpace(policyRoot) AndAlso Directory.Exists(policyRoot) Then
                Return Path.GetFullPath(policyRoot)
            End If
        Catch
        End Try

        Throw New InvalidOperationException(
            "No writable workspace is available for download_web_files. " &
            "Connect a writable workspace first, or provide an explicit absolute target_directory.")
    End Function

    Private Function ResolveDownloadTargetDirectory(requestedDirectory As String) As String
        If String.IsNullOrWhiteSpace(requestedDirectory) Then
            Dim workspaceRoot = GetSafeDownloadRoot()
            If Not Directory.Exists(workspaceRoot) Then Directory.CreateDirectory(workspaceRoot)
            Return workspaceRoot
        End If

        If Path.IsPathRooted(requestedDirectory) Then
            Dim absoluteTarget = Path.GetFullPath(requestedDirectory)
            Dim absoluteDir = Path.GetDirectoryName(absoluteTarget)

            If String.IsNullOrWhiteSpace(absoluteDir) Then
                Throw New UnauthorizedAccessException("The absolute target_directory is invalid.")
            End If

            If Not Directory.Exists(absoluteTarget) Then Directory.CreateDirectory(absoluteTarget)
            Return absoluteTarget
        End If

        Dim root As String = GetSafeDownloadRoot()
        Dim fullPath As String = Path.GetFullPath(Path.Combine(root, requestedDirectory))

        If Not fullPath.Equals(root, StringComparison.OrdinalIgnoreCase) AndAlso
           Not fullPath.StartsWith(root & Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) Then
            Throw New UnauthorizedAccessException("Download target directory is outside the permitted workspace root.")
        End If

        If Not Directory.Exists(fullPath) Then Directory.CreateDirectory(fullPath)
        Return fullPath
    End Function

    Private Function SanitizeDownloadFileName(name As String) As String
        Dim candidate As String = If(name, "").Trim().Trim(""""c)
        If candidate = "" Then candidate = "download.bin"

        For Each invalidChar In Path.GetInvalidFileNameChars()
            candidate = candidate.Replace(invalidChar, "_"c)
        Next

        If candidate = "" Then candidate = "download.bin"
        Return candidate
    End Function

    Private Function GetExtensionFromContentType(contentType As String) As String
        Dim mediaType As String = If(contentType, "").Trim().ToLowerInvariant()

        Select Case mediaType
            Case "application/pdf" : Return ".pdf"
            Case "application/zip" : Return ".zip"
            Case "application/json" : Return ".json"
            Case "text/plain" : Return ".txt"
            Case "text/html" : Return ".html"
            Case "application/xml", "text/xml" : Return ".xml"
            Case "image/png" : Return ".png"
            Case "image/jpeg" : Return ".jpg"
            Case "image/gif" : Return ".gif"
            Case Else : Return ""
        End Select
    End Function

    Private Function BuildDownloadFileName(url As String,
                                           response As System.Net.Http.HttpResponseMessage) As String
        Dim candidate As String = ""

        Try
            Dim cd = response.Content.Headers.ContentDisposition
            If cd IsNot Nothing Then
                If Not String.IsNullOrWhiteSpace(cd.FileNameStar) Then
                    candidate = cd.FileNameStar
                ElseIf Not String.IsNullOrWhiteSpace(cd.FileName) Then
                    candidate = cd.FileName
                End If
            End If
        Catch
        End Try

        If String.IsNullOrWhiteSpace(candidate) Then
            Try
                candidate = Path.GetFileName(New Uri(url).LocalPath)
            Catch
            End Try
        End If

        candidate = SanitizeDownloadFileName(candidate)

        If Path.GetExtension(candidate) = "" Then
            Dim ext = GetExtensionFromContentType(If(response.Content.Headers.ContentType?.MediaType, ""))
            If ext <> "" Then candidate &= ext
        End If

        Return candidate
    End Function



    Private Async Function ReadResponseBytesLimitedAsync(content As System.Net.Http.HttpContent,
                                                         maxBytes As Long,
                                                         cancellationToken As System.Threading.CancellationToken) As Task(Of Byte())
        Using sourceStream = Await content.ReadAsStreamAsync().ConfigureAwait(False)
            Using ms As New MemoryStream()
                Dim buffer(8191) As Byte

                Do
                    cancellationToken.ThrowIfCancellationRequested()
                    Dim read = Await sourceStream.ReadAsync(buffer, 0, buffer.Length, cancellationToken).ConfigureAwait(False)
                    If read <= 0 Then Exit Do

                    ms.Write(buffer, 0, read)

                    If ms.Length > maxBytes Then
                        Throw New InvalidOperationException($"Remote file exceeds the maximum allowed size of {maxBytes} bytes.")
                    End If
                Loop

                Return ms.ToArray()
            End Using
        End Using
    End Function

    Private Function LooksLikeHtml(bytes As Byte()) As Boolean
        If bytes Is Nothing OrElse bytes.Length = 0 Then Return False

        Dim sampleLength = Math.Min(bytes.Length, 1024)
        Dim sample = System.Text.Encoding.UTF8.GetString(bytes, 0, sampleLength).ToLowerInvariant()

        Return sample.Contains("<html") OrElse
               sample.Contains("<!doctype html") OrElse
               sample.Contains("<body") OrElse
               sample.Contains("<head")
    End Function

    Private Function LooksLikePdf(bytes As Byte()) As Boolean
        If bytes Is Nothing OrElse bytes.Length < 5 Then Return False
        Return bytes(0) = AscW("%"c) AndAlso
               bytes(1) = AscW("P"c) AndAlso
               bytes(2) = AscW("D"c) AndAlso
               bytes(3) = AscW("F"c) AndAlso
               bytes(4) = AscW("-"c)
    End Function

    Private Async Function ExecuteInternalDownloadWebFilesTool(toolCall As ToolCall,
                                                               context As ToolExecutionContext,
                                                               cancellationToken As System.Threading.CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId,
            .ToolName = toolCall.ToolName
        }

        Dim urls As List(Of String) = GetToolArgumentStringList(toolCall.Arguments, "urls")
        Dim overwrite As Boolean = GetToolArgumentBoolean(toolCall.Arguments, "overwrite", False)
        Dim targetDirectory As String = GetToolArgumentString(toolCall.Arguments, "target_directory")

        If urls Is Nothing OrElse urls.Count = 0 Then
            response.Success = False
            response.ErrorCode = "missing_urls"
            response.ErrorMessage = "No URLs were provided to download_web_files."
            response.Response = New JObject(
                New JProperty("ok", False),
                New JProperty("error", response.ErrorMessage)
            ).ToString(Formatting.None)
            Return response
        End If

        Dim downloadRoot As String = ResolveSafeWebDownloadDirectory(targetDirectory)
        Directory.CreateDirectory(downloadRoot)

        Dim downloaded As New JArray()
        Dim failed As New JArray()

        context?.Log($"Downloading {urls.Count} web file(s)...")

        For Each rawUrl As String In urls
            Dim url As String = If(rawUrl, "").Trim()

            If cancellationToken.IsCancellationRequested Then
                failed.Add(New JObject(
                    New JProperty("url", url),
                    New JProperty("error", "Operation was canceled.")
                ))
                Exit For
            End If

            If Not IsSafeWebUrl(url) Then
                failed.Add(New JObject(
                    New JProperty("url", url),
                    New JProperty("error", "URL is not an allowed absolute HTTP/HTTPS URL.")
                ))
                Continue For
            End If

            context?.Log($"Downloading: {url}")

            Try
                Using timeoutCts As New System.Threading.CancellationTokenSource(TimeSpan.FromSeconds(WebToolTimeoutSeconds))
                    Using linkedCts As System.Threading.CancellationTokenSource =
                        System.Threading.CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, timeoutCts.Token)

                        Using client As New System.Net.Http.HttpClient()
                            client.Timeout = TimeSpan.FromSeconds(WebToolTimeoutSeconds)
                            client.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0 (Windows NT 10.0; Win64; x64) RedInk/1.0")

                            Using httpResponse As System.Net.Http.HttpResponseMessage =
                                Await client.GetAsync(url, System.Net.Http.HttpCompletionOption.ResponseHeadersRead, linkedCts.Token)

                                httpResponse.EnsureSuccessStatusCode()

                                Dim contentLength As Long = -1
                                If httpResponse.Content.Headers.ContentLength.HasValue Then
                                    contentLength = httpResponse.Content.Headers.ContentLength.Value
                                End If

                                If contentLength > MaxDownloadedWebFileBytes Then
                                    failed.Add(New JObject(
                                        New JProperty("url", url),
                                        New JProperty("error", $"File is too large ({contentLength:N0} bytes). Limit is {MaxDownloadedWebFileBytes:N0} bytes.")
                                    ))
                                    Continue For
                                End If

                                Dim fileName As String = GetDownloadFileName(url, httpResponse)
                                Dim targetPath As String = Path.Combine(downloadRoot, fileName)

                                If Not overwrite Then
                                    targetPath = GetUniqueFilePath(targetPath)
                                End If

                                Dim totalBytes As Long = 0

                                Using sourceStream As Stream = Await httpResponse.Content.ReadAsStreamAsync()
                                    Using targetStream As New FileStream(targetPath, FileMode.Create, FileAccess.Write, FileShare.None)
                                        Dim buffer(81919) As Byte

                                        While True
                                            Dim read As Integer = Await sourceStream.ReadAsync(buffer, 0, buffer.Length, linkedCts.Token)
                                            If read = 0 Then Exit While

                                            totalBytes += read

                                            If totalBytes > MaxDownloadedWebFileBytes Then
                                                targetStream.Close()

                                                Try
                                                    File.Delete(targetPath)
                                                Catch
                                                End Try

                                                Throw New InvalidOperationException($"File exceeded the maximum allowed size of {MaxDownloadedWebFileBytes:N0} bytes.")
                                            End If

                                            Await targetStream.WriteAsync(buffer, 0, read, linkedCts.Token)
                                        End While
                                    End Using
                                End Using

                                downloaded.Add(New JObject(
                                    New JProperty("url", url),
                                    New JProperty("path", targetPath),
                                    New JProperty("file_name", Path.GetFileName(targetPath)),
                                    New JProperty("bytes", totalBytes)
                                ))
                            End Using
                        End Using
                    End Using
                End Using

            Catch ex As OperationCanceledException
                Dim timeoutMessage As String = $"Timed out after {WebToolTimeoutSeconds} seconds while downloading {url}."
                failed.Add(New JObject(
                    New JProperty("url", url),
                    New JProperty("error", timeoutMessage)
                ))
                context?.LogWarn(timeoutMessage)

            Catch ex As Exception
                failed.Add(New JObject(
                    New JProperty("url", url),
                    New JProperty("error", ex.Message)
                ))
                context?.LogWarn($"Download failed: {url}", details:=ex.Message)
            End Try
        Next

        Dim payload As New JObject(
            New JProperty("ok", failed.Count = 0),
            New JProperty("download_directory", downloadRoot),
            New JProperty("downloaded", downloaded),
            New JProperty("failed", failed),
            New JProperty("summary", $"{downloaded.Count} file(s) downloaded, {failed.Count} failed.")
        )

        response.Success = failed.Count = 0
        response.Response = payload.ToString(Formatting.None)

        If failed.Count > 0 Then
            response.ErrorCode = "download_failed"
            response.ErrorMessage = $"{failed.Count} download(s) failed. See tool response for details."
        End If

        Return response
    End Function

    Private Function GetToolArgumentString(arguments As Dictionary(Of String, Object), key As String, Optional defaultValue As String = "") As String
        If arguments Is Nothing OrElse Not arguments.ContainsKey(key) OrElse arguments(key) Is Nothing Then
            Return defaultValue
        End If

        Try
            Return If(arguments(key).ToString(), "").Trim()
        Catch
            Return defaultValue
        End Try
    End Function

    Private Function ResolveSafeWebDownloadDirectory(targetDirectory As String) As String
        Dim root As String = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
            "Red Ink Downloads")

        Dim requested As String = If(targetDirectory, "").Trim()

        If requested = "" Then
            Return root
        End If

        requested = requested.Replace("/"c, Path.DirectorySeparatorChar)

        If Path.IsPathRooted(requested) Then
            requested = requested.TrimStart(Path.GetPathRoot(requested).ToCharArray())
        End If

        Dim fullPath As String = Path.GetFullPath(Path.Combine(root, requested))
        Dim fullRoot As String = Path.GetFullPath(root)

        If Not fullPath.StartsWith(fullRoot, StringComparison.OrdinalIgnoreCase) Then
            Return root
        End If

        Return fullPath
    End Function

    Private Function GetDownloadFileName(url As String, response As System.Net.Http.HttpResponseMessage) As String
        Dim fileName As String = ""

        Try
            Dim disposition = response.Content.Headers.ContentDisposition
            If disposition IsNot Nothing Then
                fileName = If(disposition.FileNameStar, disposition.FileName)
                fileName = If(fileName, "").Trim().Trim(""""c)
            End If
        Catch
        End Try

        If fileName = "" Then
            Try
                Dim uri As New Uri(url)
                fileName = Path.GetFileName(uri.LocalPath)
            Catch
            End Try
        End If

        If fileName = "" Then
            fileName = "download-" & DateTime.Now.ToString("yyyyMMdd-HHmmss") & ".bin"
        End If

        Return MakeSafeFileName(fileName)
    End Function

    Private Function MakeSafeFileName(fileName As String) As String
        Dim safeName As String = If(fileName, "").Trim()

        For Each c As Char In Path.GetInvalidFileNameChars()
            safeName = safeName.Replace(c, "_"c)
        Next

        If safeName = "" Then
            safeName = "download.bin"
        End If

        Return safeName
    End Function


End Class
