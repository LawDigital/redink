' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.Processing.SearchGrounding.vb
' Purpose: Collects LLM-driven search terms, performs configured internet searches, and enriches prompts with crawled content.
'
' Workflow:
'  - ConsultInternet: Builds search prompts, optionally requests approval, runs searches, and prepares follow-up prompts.
'  - PerformSearchGrounding: Executes the HTTP search, extracts unique URLs using response masks, retrieves qualifying content, and aggregates it.
'  - RetrieveWebsiteContent/CrawlWebsite: Crawl discovered pages within bounded depth, timeout, cancellation/error limits, while harvesting paragraph text.
'  - CrawlContext/GetAbsoluteUrl: Maintain crawl state, enforce dedupe/error thresholds, and normalize relative links.
'  - RetrieveWebsiteContent_WebView2/UnescapeJsonString: (Optional) Use WebView2 to render JavaScript pages, scroll for lazy-loaded content, and extract cleaned body text.
'
' External Dependencies:
'  - SharedLibrary.SharedLibrary.SharedMethods for interpolation, UI dialogues, and LLM access.
'  - HtmlAgilityPack for HTML parsing (HttpClient crawl path).
'  - Microsoft.Web.WebView2.Core / Microsoft.Web.WebView2.WinForms for JS-rendered content retrieval (optional path).
' =============================================================================

Option Explicit On
Option Strict On

Imports System.Diagnostics
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports HtmlAgilityPack
Imports Microsoft.Web.WebView2.Core
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary.SharedMethods

''' <summary>
''' Provides search-grounding helpers that connect Word instructions with external internet content.
''' </summary>
Partial Public Class ThisAddIn

    Const UseWebView2ForISearch As Boolean = True ' Set to True to enable WebView2-based content retrieval

    ''' <summary>
    ''' Coordinates LLM-based search-term generation, optional approval, remote search execution, and prompt selection.
    ''' </summary>
    ''' <param name="DoMarkup">Determines whether the markup-specific system prompt is used when selected text exists.</param>
    ''' <returns>True when search preparation completes successfully; otherwise False.</returns>
    Public Function ConsultInternet(DoMarkup As Boolean) As Task(Of Boolean)
        Dim tcs As New TaskCompletionSource(Of Boolean)()

        Try
            InfoBox.ShowInfoBox("Asking the LLM to determine the necessary searchterms for your instruction ...")

            Dim SysPromptTemp As String
            Dim SearchResults As List(Of String)

            SysPromptTemp = InterpolateAtRuntime(INI_ISearch_SearchTerm_SP)

            ' Use polling instead of Await to stay on UI thread
            Dim llmTask As Task(Of String) = LLM(SysPromptTemp, If(SelectedText = "", "", "<TEXTTOPROCESS>" & SelectedText & "</TEXTTOPROCESS>"), "", "", 0)

            While Not llmTask.IsCompleted
                System.Windows.Forms.Application.DoEvents()
                System.Threading.Thread.Sleep(50)
            End While

            If llmTask.Status = TaskStatus.RanToCompletion Then
                SearchTerms = llmTask.Result
            Else
                SearchTerms = ""
            End If

            If String.IsNullOrWhiteSpace(SearchTerms) Then
                InfoBox.ShowInfoBox("")
                ShowCustomMessageBox("The LLM failed to establish searchterms. Will abort.")
                tcs.SetResult(False)
                Return tcs.Task
            End If

            If INI_ISearch_Approve Then
                InfoBox.ShowInfoBox("")
                Dim approveresult As Integer = ShowCustomYesNoBox("These are the searchterms that the LLM wants to issue to " & INI_ISearch_Name & ": {SearchTerms}", "Approve", "Abort", $"{AN} Internet Search", 5, " = 'Approve'")
                If approveresult = 0 Or approveresult = 2 Then
                    tcs.SetResult(False)
                    Return tcs.Task
                End If
            End If

            InfoBox.ShowInfoBox($"Now using {INI_ISearch_Name} to search for '{SearchTerms}' ...")

            ' Use polling instead of Await for PerformSearchGrounding
            Dim searchTask As Task(Of List(Of String)) = PerformSearchGrounding(SearchTerms, INI_ISearch_URL, INI_ISearch_ResponseMask1, INI_ISearch_ResponseMask2, INI_ISearch_Tries, INI_ISearch_MaxDepth)

            While Not searchTask.IsCompleted
                System.Windows.Forms.Application.DoEvents()
                System.Threading.Thread.Sleep(50)
            End While

            If searchTask.Status = TaskStatus.RanToCompletion Then
                SearchResults = searchTask.Result
            Else
                SearchResults = New List(Of String)()
            End If

            SearchResult = String.Join(Environment.NewLine, SearchResults.Select(Function(result, index) $"<SEARCHRESULT{index + 1}>{result}</SEARCHRESULT{index + 1}>"))

            InfoBox.ShowInfoBox($"Having the LLM execute your instruction using also the {SearchResults.Count} result(s) from the Internet search ...", 3)
            If DoMarkup And Not String.IsNullOrWhiteSpace(SelectedText) Then
                SysPrompt = InterpolateAtRuntime(INI_ISearch_Apply_SP_Markup)
            Else
                SysPrompt = InterpolateAtRuntime(INI_ISearch_Apply_SP)
            End If

            tcs.SetResult(True)

        Catch ex As System.Exception
            MessageBox.Show("Error in ConsultInternet: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            tcs.SetResult(False)
        End Try

        Return tcs.Task
    End Function

    ''' <summary>
    ''' Executes the configured internet search, extracts unique URLs using response masks, retrieves qualifying content, and returns the collected snippets.
    ''' </summary>
    Public Function PerformSearchGrounding(SGTerms As String, ISearch_URL As String, ISearch_ResponseMask1 As String, ISearch_ResponseMask2 As String, ISearch_Tries As Integer, ISearch_MaxDepth As Integer) As Task(Of List(Of String))
        Dim tcs As New TaskCompletionSource(Of List(Of String))()
        Dim results As New List(Of String)

        Using httpClient As New HttpClient()
            Try
                ' Construct the search URL
                Dim searchUrl As String = ISearch_URL & Uri.EscapeDataString(SGTerms)

                InfoBox.ShowInfoBox($"Searching {searchUrl} ...")

                ' Get search results HTML
                httpClient.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36")
                httpClient.Timeout = TimeSpan.FromSeconds(30)

                ' Use polling instead of Await
                Dim httpTask As Task(Of String) = httpClient.GetStringAsync(searchUrl)

                While Not httpTask.IsCompleted
                    System.Windows.Forms.Application.DoEvents()
                    System.Threading.Thread.Sleep(50)
                End While

                Dim searchResponse As String = ""
                If httpTask.Status = TaskStatus.RanToCompletion Then
                    searchResponse = httpTask.Result
                Else
                    Throw New HttpRequestException("Failed to get search results")
                End If

                InfoBox.ShowInfoBox($"Extracting URLs ...")

                ' Extract URLs using the defined start and mask
                Dim urlPattern As String = Regex.Escape(ISearch_ResponseMask1) & "(.*?)" & Regex.Escape(ISearch_ResponseMask2)
                Dim matches As MatchCollection = Regex.Matches(searchResponse, urlPattern)

                Dim extractedUrls As New List(Of String)
                Dim URLList As String = "URLS found so far:" & vbCrLf & vbCrLf
                For Each match As Match In matches
                    Dim rawUrl As String = match.Groups(1).Value
                    Dim decodedUrl As String = WebUtility.UrlDecode(rawUrl.Replace(ISearch_ResponseMask1, ""))

                    If Not extractedUrls.Contains(decodedUrl) Then
                        extractedUrls.Add(decodedUrl)
                        URLList += decodedUrl & vbCrLf
                        InfoBox.ShowInfoBox(URLList)
                    End If

                    If extractedUrls.Count >= ISearch_Tries Then Exit For
                Next

                ' Visit each extracted URL and retrieve content
                For Each url In extractedUrls
                    Try
                        Dim content As String = ""

                        If UseWebView2ForISearch Then
                            ' Use polling for WebView2
                            Dim webViewTask As Task(Of String) = RetrieveWebsiteContent_WebView2(url, 0)

                            While Not webViewTask.IsCompleted
                                System.Windows.Forms.Application.DoEvents()
                                System.Threading.Thread.Sleep(50)
                            End While

                            If webViewTask.Status = TaskStatus.RanToCompletion Then
                                content = webViewTask.Result
                            End If
                        Else
                            ' Use polling for HttpClient crawl
                            Dim crawlTask As Task(Of String) = RetrieveWebsiteContent(url, ISearch_MaxDepth, httpClient)

                            While Not crawlTask.IsCompleted
                                System.Windows.Forms.Application.DoEvents()
                                System.Threading.Thread.Sleep(50)
                            End While

                            If crawlTask.Status = TaskStatus.RanToCompletion Then
                                content = crawlTask.Result
                            End If
                        End If

                        If Not String.IsNullOrWhiteSpace(content) Then
                            If Len(content) > ISearch_MinChars Then
                                results.Add(content)
                                InfoBox.ShowInfoBox($"{url} resulted in: " & Left(content.Replace(vbCr, "").Replace(vbLf, "").Replace(vbCrLf, ""), 1000))
                            End If
                        End If
                    Catch ex As Exception
                        Debug.WriteLine($"Error retrieving content from URL: {url} - {ex.Message}")
                    End Try
                Next

            Catch ex As HttpRequestException
                ShowCustomMessageBox("An error occurred when searching and analyzing the Internet (HTTP request error: " & ex.Message & ")")
            Catch ex As TaskCanceledException
                ShowCustomMessageBox("An error occurred when searching and analyzing the Internet (request timed-out or was canceled: " & ex.Message & ")")
            Catch ex As Exception
                ShowCustomMessageBox("An error occurred when searching and analyzing the Internet (" & ex.Message & ")")
            Finally
                InfoBox.ShowInfoBox("")
            End Try
        End Using

        tcs.SetResult(results)
        Return tcs.Task
    End Function
    ''' <summary>
    ''' Crawls a base URL up to the specified depth, captures paragraph text, strips HTML, and enforces timeout plus error limits.
    ''' </summary>
    ''' <param name="baseUrl">Starting URL that seeds the crawl.</param>
    ''' <param name="subTries">Maximum link depth to follow.</param>
    ''' <param name="httpClient">HttpClient instance reused for the crawl.</param>
    ''' <returns>Plain-text content (limited to ISearch_MaxChars) harvested from the crawl.</returns>
    Private Async Function RetrieveWebsiteContent(
                        baseUrl As String,
                        subTries As Integer,
                        httpClient As HttpClient
                    ) As Task(Of String)

        ' Use the shared HttpClient instance provided by the caller for the entire crawl
        Dim client As HttpClient = httpClient

        ' Create the shared context object
        Dim context As New CrawlContext With {
                    .VisitedUrls = New HashSet(Of String)(),
                    .ContentBuilder = New StringBuilder(),
                    .ErrorCount = 0,
                    .MaxErrors = ISearch_MaxCrawlErrors
                }

        ' Create one CancellationTokenSource for the entire crawl duration
        Using cts As New CancellationTokenSource(TimeSpan.FromSeconds(INI_ISearch_Timeout))
            ' Call the CrawlWebsite function with the context
            '   - 'subTries' is your maxDepth
            '   - '0' is your currentDepth

            Await CrawlWebsite(
                        currentUrl:=baseUrl,
                        maxDepth:=subTries,
                        currentDepth:=0,
                        httpClient:=client,
                        context:=context,
                        cancellationToken:=cts.Token,
                        timeOutSeconds:=CInt(INI_ISearch_Timeout)
                         )

        End Using

        ' Return plain text with HTML tags removed (up to ISearch_MaxChars)
        Return Left(
                    Regex.Replace(context.ContentBuilder.ToString(), "<.*?>", String.Empty).Trim(),
                    ISearch_MaxChars
                    )
    End Function

    ''' <summary>
    ''' Holds crawl state across recursive invocations so that visited urls, aggregated content, and error limits stay consistent.
    ''' </summary>
    Public Class CrawlContext
        Public Property VisitedUrls As HashSet(Of String)
        Public Property ContentBuilder As StringBuilder
        Public Property ErrorCount As Integer
        Public Property MaxErrors As Integer
    End Class

    ''' <summary>
    ''' Recursively crawls a URL, collecting paragraph text and following links while honoring depth, timeout, and error thresholds.
    ''' </summary>
    ''' <param name="currentUrl">The URL currently being crawled.</param>
    ''' <param name="maxDepth">Maximum depth allowed for recursion.</param>
    ''' <param name="currentDepth">Current recursion depth.</param>
    ''' <param name="httpClient">HttpClient used for fetching page content.</param>
    ''' <param name="context">Shared crawl context for deduplication and aggregation.</param>
    ''' <param name="cancellationToken">Cancellation token controlling crawl timeout.</param>
    ''' <param name="timeOutSeconds">Timeout used when the caller does not supply a token.</param>
    ''' <returns>Empty string because content is accumulated through the shared context.</returns>
    Private Async Function CrawlWebsite(
    currentUrl As String,
    maxDepth As Integer,
    currentDepth As Integer,
    httpClient As HttpClient,
    context As CrawlContext,
    Optional cancellationToken As CancellationToken = Nothing,
    Optional timeOutSeconds As Integer = 10
) As Task(Of String)

        ' If the function has no valid CancellationToken, create one that cancels after 30 seconds
        Dim localCts As CancellationTokenSource = Nothing
        If cancellationToken = CancellationToken.None Then
            localCts = New CancellationTokenSource(TimeSpan.FromSeconds(timeOutSeconds))
            cancellationToken = localCts.Token
        End If

        Dim results As String = ""

        ' If we've already exceeded the max errors, abort quickly
        If context.ErrorCount >= context.MaxErrors Then
            Return results
        End If

        ' Early exit if depth is exceeded or already visited
        If currentDepth > maxDepth OrElse context.VisitedUrls.Contains(currentUrl) Then
            Return results
        End If

        Try
            context.VisitedUrls.Add(currentUrl)

            ' Use the cancellation token to abort if it exceeds the specified time
            Dim response As HttpResponseMessage = Await httpClient.GetAsync(currentUrl, cancellationToken)
            Dim pageHtml As String = Await response.Content.ReadAsStringAsync()

            Dim doc As New HtmlAgilityPack.HtmlDocument()
            doc.LoadHtml(pageHtml)

            ' Safely extract paragraph text
            Dim pNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//p")
            If pNodes IsNot Nothing Then
                For Each node In pNodes
                    context.ContentBuilder.AppendLine(node.InnerText.Trim())
                Next
            End If

            ' Follow links if depth permits
            If currentDepth < maxDepth Then
                Dim links As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//a[@href]")
                If links IsNot Nothing Then
                    For Each link In links
                        Dim hrefValue As String = link.GetAttributeValue("href", "").Trim()
                        Dim absoluteUrl As String = GetAbsoluteUrl(currentUrl, hrefValue)
                        ' You should already have a GetAbsoluteUrl function that resolves relative paths

                        If Not String.IsNullOrEmpty(absoluteUrl) Then
                            Await CrawlWebsite(
                            absoluteUrl,
                            maxDepth,
                            currentDepth + 1,
                            httpClient,
                            context,
                            cancellationToken,
                            timeOutSeconds
                        )

                            ' If error count has now exceeded the limit, stop immediately
                            If context.ErrorCount >= context.MaxErrors Then
                                Exit For
                            End If
                        End If
                    Next
                End If
            End If

        Catch ex As System.Threading.Tasks.TaskCanceledException
            ' Decide if a cancellation/timeout should increment errorCount
            context.ErrorCount += 1
            Debug.WriteLine($"Task canceled while crawling URL: {currentUrl} - {ex.Message}")

        Catch ex As System.Exception
            context.ErrorCount += 1
            Debug.WriteLine($"Error crawling URL: {currentUrl} - {ex.Message}")
        Finally
            If localCts IsNot Nothing Then
                localCts.Dispose()
            End If
        End Try

        Return results
    End Function

    ''' <summary>
    ''' Resolves a possibly relative link against the provided base URL and returns the absolute URL string.
    ''' </summary>
    ''' <param name="baseUrl">Page URL that acts as the anchor for relative links.</param>
    ''' <param name="relativeUrl">Relative or absolute href value extracted from a link.</param>
    ''' <returns>Absolute URL when resolvable; otherwise an empty string.</returns>
    Private Function GetAbsoluteUrl(baseUrl As String, relativeUrl As String) As String
        Try
            Dim baseUri As New Uri(baseUrl)
            Dim absoluteUri As New Uri(baseUri, relativeUrl)
            Return absoluteUri.ToString()
        Catch ex As Exception
            ' Invalid relative URL handling
            Return String.Empty
        End Try
    End Function


    '========================= WebView 2 Alternative Implementation ============================

    ''' <summary>
    ''' Document extensions that should be downloaded and processed instead of rendered via WebView2.
    ''' </summary>
    Private Shared ReadOnly DocumentExtensions As String() = {".pdf", ".docx", ".pptx", ".txt", ".rtf"}
    ' NOTE: .doc (legacy binary format) is excluded - requires Word Interop which can execute macros

    ''' <summary>
    ''' Maximum file size for document downloads (50 MB).
    ''' </summary>
    Private Const MaxDocumentDownloadBytes As Long = 50 * 1024 * 1024

    Private Class WebRetrievalResult
        Public Property TextContent As String = ""
        Public Property FinalUrl As String = ""
        Public Property LinksJson As String = "[]"
    End Class

    ''' <summary>
    ''' Retrieves website content using WebView2 (Edge-based, built into Windows).
    ''' Handles JavaScript rendering without external browser downloads.
    ''' For document URLs (PDF, Word, PowerPoint, TXT), downloads and extracts text using safe parsers.
    ''' </summary>
    ''' <param name="baseUrl">The URL to fetch.</param>
    ''' <param name="maxChars">Maximum characters to return. 0 or negative = no limit.</param>
    ''' <param name="expandCollapsed">When True, attempts to expand collapsed/hidden content before extraction.</param>
    Private Async Function RetrieveWebsiteContent_WebView2(baseUrl As String,
                                                           Optional maxChars As Integer = 0,
                                                           Optional expandCollapsed As Boolean = True) As Task(Of String)

        Dim pageResult = Await RetrieveWebsiteContent_WebView2Detailed(
            baseUrl,
            maxChars,
            expandCollapsed,
            includeLinks:=False,
            linkExtensions:=Nothing).ConfigureAwait(False)

        Return If(pageResult?.TextContent, "")
    End Function

    Private Function RetrieveWebsiteContent_WebView2Detailed(baseUrl As String,
                                                         Optional maxChars As Integer = 0,
                                                         Optional expandCollapsed As Boolean = True,
                                                         Optional includeLinks As Boolean = False,
                                                         Optional linkExtensions As System.Collections.Generic.List(Of String) = Nothing) As System.Threading.Tasks.Task(Of WebRetrievalResult)

        Dim tcs As New System.Threading.Tasks.TaskCompletionSource(Of WebRetrievalResult)()

        Dim normalizedExtensions As System.Collections.Generic.List(Of String) =
        If(linkExtensions, New System.Collections.Generic.List(Of String)()).
        Select(Function(x) If(x, "").Trim().TrimStart("."c).ToLowerInvariant()).
        Where(Function(x) x <> "").
        Distinct(System.StringComparer.OrdinalIgnoreCase).
        ToList()

        Dim documentContent As String = Nothing

        If TryExtractDocumentContent(baseUrl, documentContent) Then
            If maxChars > 0 AndAlso documentContent.Length > maxChars Then
                documentContent = TrimToSentenceBoundary(documentContent, maxChars)
            End If

            tcs.TrySetResult(New WebRetrievalResult() With {
            .TextContent = documentContent,
            .FinalUrl = baseUrl,
            .LinksJson = "[]"
        })

            Return tcs.Task
        End If

        Dim thread As New System.Threading.Thread(
        Sub()
            Dim result As New WebRetrievalResult() With {
                .TextContent = "",
                .FinalUrl = baseUrl,
                .LinksJson = "[]"
            }

            Dim form As System.Windows.Forms.Form = Nothing
            Dim webView As Microsoft.Web.WebView2.WinForms.WebView2 = Nothing
            Dim userDataFolder As String = ""

            Dim navigationFinished As Boolean = False
            Dim navigationFailed As Boolean = False
            Dim contentExtracted As Boolean = False
            Dim extractionStarted As Boolean = False

            Dim startTime As System.DateTime = System.DateTime.Now
            Dim timeout As System.TimeSpan = System.TimeSpan.FromSeconds(90)

            Try
                System.Diagnostics.Debug.WriteLine(
                    $"[WebView2] Fetching: {baseUrl} (maxChars: {If(maxChars <= 0, "unlimited", maxChars.ToString())}, includeLinks: {includeLinks}, expandCollapsed: {expandCollapsed})")

                If Not IsSafeWebUrl(baseUrl) Then
                    System.Diagnostics.Debug.WriteLine($"[WebView2] Blocked unsafe URL: {baseUrl}")
                    navigationFailed = True
                    result.TextContent = $"Blocked unsafe URL: {baseUrl}"
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
                                Dim initMessage As String =
                                    If(e.InitializationException IsNot Nothing,
                                       e.InitializationException.Message,
                                       "Unknown initialization error")

                                System.Diagnostics.Debug.WriteLine("[WebView2] CoreWebView2 initialization failed: " & initMessage)
                                result.TextContent = "WebView2 initialization failed: " & initMessage
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

                                ' Wichtig: Muss True sein, weil das Script per chrome.webview.postMessage(...)
                                ' an VB.NET zurückmeldet.
                                .IsWebMessageEnabled = True
                            End With

                            AddHandler webView.CoreWebView2.WebMessageReceived,
                                Sub(sender, args)
                                    Try
                                        Dim rawJson As String = args.WebMessageAsJson

                                        System.Diagnostics.Debug.WriteLine("[WebView2] WebMessage received: " & If(rawJson, "<null>"))

                                        If System.String.IsNullOrWhiteSpace(rawJson) Then
                                            contentExtracted = True
                                            Return
                                        End If

                                        Dim payload As Newtonsoft.Json.Linq.JObject =
                                            Newtonsoft.Json.Linq.JObject.Parse(rawJson)

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

                                        Dim errorToken As Newtonsoft.Json.Linq.JToken = payload("error")
                                        If errorToken IsNot Nothing AndAlso System.String.IsNullOrWhiteSpace(result.TextContent) Then
                                            result.TextContent = "WebView2 extraction error: " & errorToken.ToString()
                                        End If

                                        System.Diagnostics.Debug.WriteLine(
                                            $"[WebView2] Extract result: text={If(result.TextContent, "").Length} chars, elapsed={(System.DateTime.Now - startTime).TotalSeconds:F1}s")

                                        contentExtracted = True

                                    Catch ex As System.Exception
                                        System.Diagnostics.Debug.WriteLine("[WebView2] WebMessage parse error: " & ex.Message)
                                        result.TextContent = "WebView2 WebMessage parse error: " & ex.Message
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
                                        navigationFinished = True

                                        System.Diagnostics.Debug.WriteLine(
                                            $"[WebView2] Navigation completed. Success: {args.IsSuccess}, Status: {args.WebErrorStatus}")

                                        If Not args.IsSuccess Then
                                            result.TextContent = $"Navigation failed: {args.WebErrorStatus}"
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
                                                expandCollapsed,
                                                normalizedExtensions)

                                        ' Wichtig:
                                        ' ExecuteScriptAsync startet nur noch das Script.
                                        ' Das eigentliche Resultat kommt über WebMessageReceived.
                                        webView.CoreWebView2.ExecuteScriptAsync(extractScript).ContinueWith(
                                            Sub(extractTask As System.Threading.Tasks.Task(Of String))
                                                Try
                                                    If extractTask.IsFaulted Then
                                                        Dim message As String = "Unknown script execution error"

                                                        If extractTask.Exception IsNot Nothing Then
                                                            message = extractTask.Exception.GetBaseException().Message
                                                        End If

                                                        System.Diagnostics.Debug.WriteLine("[WebView2] ExecuteScriptAsync failed: " & message)
                                                        result.TextContent = "WebView2 ExecuteScriptAsync failed: " & message
                                                        contentExtracted = True
                                                    Else
                                                        System.Diagnostics.Debug.WriteLine("[WebView2] Extraction script started.")
                                                    End If

                                                Catch ex As System.Exception
                                                    System.Diagnostics.Debug.WriteLine("[WebView2] ExecuteScriptAsync continuation error: " & ex.Message)
                                                    result.TextContent = "WebView2 ExecuteScriptAsync continuation error: " & ex.Message
                                                    contentExtracted = True
                                                End Try
                                            End Sub)

                                    Catch ex As System.Exception
                                        System.Diagnostics.Debug.WriteLine("[WebView2] NavigationCompleted error: " & ex.Message)
                                        result.TextContent = "WebView2 NavigationCompleted error: " & ex.Message
                                        contentExtracted = True
                                    End Try
                                End Sub

                            webView.CoreWebView2.Navigate(baseUrl)

                        Catch ex As System.Exception
                            System.Diagnostics.Debug.WriteLine("[WebView2] Initialization handler error: " & ex.Message)
                            result.TextContent = "WebView2 initialization handler error: " & ex.Message
                            navigationFailed = True
                        End Try
                    End Sub

                form.Show()

                Dim env As System.Threading.Tasks.Task(Of Microsoft.Web.WebView2.Core.CoreWebView2Environment) =
                    Microsoft.Web.WebView2.Core.CoreWebView2Environment.CreateAsync(Nothing, userDataFolder)

                env.ContinueWith(
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

                                            System.Diagnostics.Debug.WriteLine("[WebView2] Environment creation failed: " & message)
                                            result.TextContent = "WebView2 environment creation failed: " & message
                                            navigationFailed = True
                                        End If

                                    Catch ex As System.Exception
                                        System.Diagnostics.Debug.WriteLine("[WebView2] Environment BeginInvoke error: " & ex.Message)
                                        result.TextContent = "WebView2 environment BeginInvoke error: " & ex.Message
                                        navigationFailed = True
                                    End Try
                                End Sub)

                        Catch ex As System.Exception
                            System.Diagnostics.Debug.WriteLine("[WebView2] Environment continuation error: " & ex.Message)
                            result.TextContent = "WebView2 environment continuation error: " & ex.Message
                            navigationFailed = True
                        End Try
                    End Sub)

                While Not contentExtracted AndAlso Not navigationFailed AndAlso (System.DateTime.Now - startTime) < timeout
                    System.Windows.Forms.Application.DoEvents()
                    System.Threading.Thread.Sleep(50)
                End While

                If Not contentExtracted AndAlso Not navigationFailed Then
                    System.Diagnostics.Debug.WriteLine($"[WebView2] Timed out after {timeout.TotalSeconds:F0} seconds while retrieving {baseUrl}")
                    result.TextContent = $"Timed out after {timeout.TotalSeconds:F0} seconds while retrieving {baseUrl}."
                End If

            Catch ex As System.Exception
                System.Diagnostics.Debug.WriteLine("[WebView2] Error: " & ex.Message)
                result.TextContent = "WebView2 error: " & ex.Message

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
                    If Not System.String.IsNullOrEmpty(userDataFolder) AndAlso System.IO.Directory.Exists(userDataFolder) Then
                        System.IO.Directory.Delete(userDataFolder, True)
                    End If
                Catch ex As System.Exception
                    System.Diagnostics.Debug.WriteLine("[WebView2] User data folder cleanup error: " & ex.Message)
                End Try
            End Try

            Dim finalText As String = If(result.TextContent, "").Trim()

            If maxChars > 0 AndAlso finalText.Length > maxChars Then
                finalText = TrimToSentenceBoundary(finalText, maxChars)
            End If

            result.TextContent = finalText

            System.Diagnostics.Debug.WriteLine(
                $"[WebView2] Loop ended. Content: {If(result.TextContent, "").Length} chars, elapsed: {(System.DateTime.Now - startTime).TotalSeconds:F1}s")

            tcs.TrySetResult(result)
        End Sub)

        thread.SetApartmentState(System.Threading.ApartmentState.STA)
        thread.Start()

        Return tcs.Task
    End Function

    Private Function BuildRobustWebExtractionScript(includeLinks As Boolean,
                                                expandInteractiveSections As Boolean,
                                                allowedExtensions As System.Collections.Generic.List(Of String)) As String

        Dim script As String = <![CDATA[
(async function() {
    var includeLinks = __INCLUDE_LINKS__;
    var expandInteractiveSections = __EXPAND_INTERACTIVE__;
    var allowedExtensions = __ALLOWED_EXTENSIONS__;

    function sendResult(payload) {
        try {
            if (window.chrome && chrome.webview && chrome.webview.postMessage) {
                chrome.webview.postMessage(payload);
            }
        } catch (e) {
        }
    }

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
    ''' Attempts to download and extract text content from a document URL (PDF, Word, PowerPoint, TXT).
    ''' Uses safe parsers that do not execute macros or scripts.
    ''' </summary>
    ''' <param name="url">The URL to check and potentially download.</param>
    ''' <param name="content">Output parameter containing extracted text if successful.</param>
    ''' <returns>True if the URL was a document and content was extracted; otherwise False.</returns>
    Private Function TryExtractDocumentContent(url As String, ByRef content As String) As Boolean
        content = Nothing

        Try
            ' Validate URL
            If Not IsSafeWebUrl(url) Then Return False

            Dim uri As New Uri(url)
            Dim path As String = uri.AbsolutePath.ToLowerInvariant()

            ' Check for query string parameters that might indicate document type
            Dim queryExt As String = ""
            If Not String.IsNullOrEmpty(uri.Query) Then
                For Each ext In DocumentExtensions
                    If uri.Query.ToLowerInvariant().Contains(ext) Then
                        queryExt = ext
                        Exit For
                    End If
                Next
            End If

            ' Determine extension from path or query
            Dim extension As String = IO.Path.GetExtension(path).ToLowerInvariant()
            If String.IsNullOrEmpty(extension) OrElse Not DocumentExtensions.Contains(extension) Then
                extension = queryExt
            End If

            ' If no document extension found, return False to use WebView2
            If String.IsNullOrEmpty(extension) OrElse Not DocumentExtensions.Contains(extension) Then
                Return False
            End If

            Debug.WriteLine($"[Document] Detected document URL: {url} (extension: {extension})")

            ' Download the file to a temp location            
            Dim tempFile As String = IO.Path.Combine(IO.Path.GetTempPath(), $"RedInk_Download_{Guid.NewGuid()}{extension}")

            Using client As New HttpClient()
                client.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
                client.Timeout = TimeSpan.FromSeconds(60)

                ' Use polling for synchronous-style download with size check
                Dim responseTask = client.GetAsync(url, HttpCompletionOption.ResponseHeadersRead)
                While Not responseTask.IsCompleted
                    System.Windows.Forms.Application.DoEvents()
                    System.Threading.Thread.Sleep(50)
                End While

                If responseTask.Status <> TaskStatus.RanToCompletion Then
                    Debug.WriteLine($"[Document] Download failed for: {url}")
                    Return False
                End If

                Using response = responseTask.Result
                    ' Check HTTP status
                    If Not response.IsSuccessStatusCode Then
                        Debug.WriteLine($"[Document] HTTP {response.StatusCode} for: {url}")
                        Return False
                    End If

                    ' Check content length if available
                    If response.Content.Headers.ContentLength.HasValue Then
                        If response.Content.Headers.ContentLength.Value > MaxDocumentDownloadBytes Then
                            Debug.WriteLine($"[Document] File too large ({response.Content.Headers.ContentLength.Value} bytes): {url}")
                            Return False
                        End If
                    End If

                    ' Download with size limit enforcement
                    Dim readTask = response.Content.ReadAsByteArrayAsync()
                    While Not readTask.IsCompleted
                        System.Windows.Forms.Application.DoEvents()
                        System.Threading.Thread.Sleep(50)
                    End While

                    If readTask.Status <> TaskStatus.RanToCompletion Then
                        Debug.WriteLine($"[Document] Failed to read content for: {url}")
                        Return False
                    End If

                    Dim fileBytes = readTask.Result
                    If fileBytes.Length > MaxDocumentDownloadBytes Then
                        Debug.WriteLine($"[Document] File too large ({fileBytes.Length} bytes): {url}")
                        Return False
                    End If

                    File.WriteAllBytes(tempFile, fileBytes)
                    Debug.WriteLine($"[Document] Downloaded {fileBytes.Length} bytes to: {tempFile}")
                End Using
            End Using

            Try
                ' Extract content based on file type using SAFE parsers only
                Select Case extension
                    Case ".pdf"
                        ' PdfPig is safe - pure .NET parser, no JS/macro execution
                        Dim pdfTask = ReadPdfAsText(tempFile, True, False, False, _context)
                        While Not pdfTask.IsCompleted
                            System.Windows.Forms.Application.DoEvents()
                            System.Threading.Thread.Sleep(50)
                        End While
                        If pdfTask.Status = TaskStatus.RanToCompletion Then
                            content = pdfTask.Result
                        End If

                    Case ".docx"
                        ' Open XML SDK is safe - only reads XML structure, no macro execution
                        content = ReadDocxSandboxed(tempFile)

                    Case ".pptx"
                        ' Open XML SDK is safe - only reads XML, no macro execution
                        content = GetPresentationJson(tempFile)

                    Case ".txt"
                        ' Plain text is safe
                        content = ReadTextFile(tempFile, False)

                    Case ".rtf"
                        ' RichTextBox parsing is safe - no macro execution
                        content = ReadRtfAsText(tempFile, False)

                        ' NOTE: .doc (legacy binary) is NOT supported here due to macro risks
                        ' It would require Word Interop which can execute embedded macros
                End Select

                Debug.WriteLine($"[Document] Extracted {If(content?.Length, 0)} chars from {extension} file")

            Finally
                ' Clean up temp file
                Try
                    If File.Exists(tempFile) Then File.Delete(tempFile)
                Catch
                    ' Ignore cleanup errors
                End Try
            End Try

            Return Not String.IsNullOrWhiteSpace(content)

        Catch ex As Exception
            Debug.WriteLine($"[Document] Error extracting document content: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Trims text to a maximum length, attempting to cut at a sentence boundary.
    ''' </summary>
    ''' <param name="text">The text to trim.</param>
    ''' <param name="maxChars">Maximum number of characters.</param>
    ''' <returns>Trimmed text.</returns>
    Private Function TrimToSentenceBoundary(text As String, maxChars As Integer) As String
        If String.IsNullOrEmpty(text) OrElse text.Length <= maxChars Then
            Return text
        End If

        Dim cutPoint As Integer = maxChars
        Dim lastPeriod As Integer = text.LastIndexOf("."c, maxChars - 1)
        Dim lastNewline As Integer = text.LastIndexOf(vbLf, maxChars - 1)

        If lastPeriod > maxChars * 0.8 Then
            cutPoint = lastPeriod + 1
        ElseIf lastNewline > maxChars * 0.8 Then
            cutPoint = lastNewline
        End If

        Dim result = text.Substring(0, cutPoint).Trim()
        Debug.WriteLine($"[Document] Trimmed to {result.Length} chars (limit was {maxChars})")
        Return result
    End Function


    ''' <summary>
    ''' Validates that a URL is safe to crawl (Http/Https only, no local files/loopback).
    ''' </summary>
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

            ' (Optional) Advanced: DNS resolve to check for private IP ranges (192.168.x.x, 10.x.x.x)

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function


    ''' <summary>
    ''' Unescapes a JSON string returned by ExecuteScriptAsync.
    ''' </summary>
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
                Function(m) ChrW(Convert.ToInt32(m.Groups(1).Value, 16)).ToString())

        Catch ex As Exception
            Debug.WriteLine($"[WebView2] Unescape error: {ex.Message}")
        End Try

        Return jsonString
    End Function

End Class
