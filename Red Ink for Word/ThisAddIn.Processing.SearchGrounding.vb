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
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports Microsoft.Web.WebView2.Core

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

            CurrentDate = DateAndTime.Now.ToString("MMMM d, yyyy")

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
    ''' Retrieves website content using WebView2 (Edge-based, built into Windows).
    ''' Handles JavaScript rendering without external browser downloads.
    ''' </summary>
    ''' <param name="baseUrl">The URL to fetch.</param>
    ''' <param name="maxChars">Maximum characters to return. 0 or negative = no limit.</param>
    Private Function RetrieveWebsiteContent_WebView2(baseUrl As String, Optional maxChars As Integer = 0) As Task(Of String)
        Dim tcs As New TaskCompletionSource(Of String)()

        ' Create a new STA thread for WebView2 (it requires its own message loop)
        Dim thread As New System.Threading.Thread(
            Sub()
                Dim result As String = ""
                Dim form As Form = Nothing
                Dim webView As Microsoft.Web.WebView2.WinForms.WebView2 = Nothing

                Try
                    Debug.WriteLine($"[WebView2] Fetching: {baseUrl} (maxChars: {If(maxChars <= 0, "unlimited", maxChars.ToString())})")

                    ' Create a temporary user data folder
                    Dim userDataFolder As String = Path.Combine(Path.GetTempPath(), "RedInkWebView2")
                    Directory.CreateDirectory(userDataFolder)

                    ' Create the form - larger size to trigger more content loading
                    form = New Form() With {
                        .Width = 1920,
                        .Height = 4000,
                        .ShowInTaskbar = False,
                        .FormBorderStyle = FormBorderStyle.None,
                        .StartPosition = FormStartPosition.Manual,
                        .Location = New System.Drawing.Point(-5000, -5000),
                        .Opacity = 0
                    }

                    webView = New Microsoft.Web.WebView2.WinForms.WebView2() With {
                        .Dock = DockStyle.Fill
                    }

                    form.Controls.Add(webView)

                    Dim navigationCompleted As Boolean = False
                    Dim navigationSuccess As Boolean = False
                    Dim contentExtracted As Boolean = False

                    ' Set up event handlers
                    AddHandler webView.CoreWebView2InitializationCompleted,
                        Sub(s, e)
                            If e.IsSuccess Then
                                Debug.WriteLine("[WebView2] CoreWebView2 initialized, navigating...")
                                webView.CoreWebView2.Navigate(baseUrl)
                            Else
                                Debug.WriteLine($"[WebView2] Initialization failed: {e.InitializationException?.Message}")
                                navigationCompleted = True
                            End If
                        End Sub

                    AddHandler webView.NavigationCompleted,
                        Sub(s, e)
                            navigationSuccess = e.IsSuccess
                            Debug.WriteLine($"[WebView2] Navigation completed. Success: {e.IsSuccess}, Status: {e.WebErrorStatus}")

                            If e.IsSuccess Then
                                ' Use a timer to wait for JS rendering - 5 seconds
                                Dim timer As New System.Windows.Forms.Timer() With {.Interval = 5000}
                                AddHandler timer.Tick,
                                    Sub(ts, te)
                                        timer.Stop()
                                        timer.Dispose()

                                        Try
                                            ' Scroll to trigger lazy loading
                                            Dim scrollScript As String = "
(async function() {
    var totalHeight = document.body.scrollHeight;
    var viewportHeight = window.innerHeight || 1000;
    var currentPosition = 0;
    
    while (currentPosition < totalHeight) {
        window.scrollTo(0, currentPosition);
        await new Promise(r => setTimeout(r, 200));
        currentPosition += viewportHeight;
        totalHeight = document.body.scrollHeight;
    }
    
    window.scrollTo(0, 0);
    await new Promise(r => setTimeout(r, 300));
    return 'done';
})();
"
                                            webView.CoreWebView2.ExecuteScriptAsync(scrollScript).ContinueWith(
                                                Sub(scrollTask)
                                                    ' Wait after scrolling
                                                    System.Threading.Thread.Sleep(2000)

                                                    form.BeginInvoke(
                                                        Sub()
                                                            Try
                                                                ' Simple extraction - get full body text
                                                                Dim extractScript As String = "
(function() {
    // Remove script/style/noscript to reduce noise
    var toRemove = document.querySelectorAll('script, style, noscript, nav, footer, header');
    toRemove.forEach(function(el) { try { el.remove(); } catch(e) {} });
    
    // Get body text
    var text = document.body ? (document.body.innerText || '') : '';
    
    // Clean up whitespace
    text = text.replace(/\n{3,}/g, '\n\n').replace(/[ \t]+/g, ' ').trim();
    
    return text;
})();
"
                                                                webView.CoreWebView2.ExecuteScriptAsync(extractScript).ContinueWith(
                                                                    Sub(t)
                                                                        form.BeginInvoke(
                                                                            Sub()
                                                                                Try
                                                                                    If t.IsCompleted AndAlso Not t.IsFaulted Then
                                                                                        result = UnescapeJsonString(t.Result)
                                                                                        Debug.WriteLine($"[WebView2] Extracted {result.Length} chars (full content)")
                                                                                    End If
                                                                                Catch ex As Exception
                                                                                    Debug.WriteLine($"[WebView2] Extract error: {ex.Message}")
                                                                                End Try
                                                                                contentExtracted = True
                                                                            End Sub)
                                                                    End Sub)
                                                            Catch ex As Exception
                                                                Debug.WriteLine($"[WebView2] Script error: {ex.Message}")
                                                                contentExtracted = True
                                                            End Try
                                                        End Sub)
                                                End Sub)
                                        Catch ex As Exception
                                            Debug.WriteLine($"[WebView2] Timer error: {ex.Message}")
                                            contentExtracted = True
                                        End Try
                                    End Sub
                                timer.Start()
                            Else
                                navigationCompleted = True
                            End If
                        End Sub

                    ' Show form and start async initialization
                    form.Show()

                    ' Initialize WebView2 asynchronously
                    Dim env = CoreWebView2Environment.CreateAsync(Nothing, userDataFolder)
                    env.ContinueWith(
                        Sub(t)
                            form.BeginInvoke(
                                Sub()
                                    If t.IsCompleted AndAlso Not t.IsFaulted Then
                                        webView.EnsureCoreWebView2Async(t.Result)
                                    Else
                                        Debug.WriteLine($"[WebView2] Environment creation failed")
                                        navigationCompleted = True
                                    End If
                                End Sub)
                        End Sub)

                    ' Run message loop with timeout - 90 seconds
                    Dim startTime As DateTime = DateTime.Now
                    Dim timeout As TimeSpan = TimeSpan.FromSeconds(90)

                    While Not contentExtracted AndAlso Not navigationCompleted AndAlso (DateTime.Now - startTime) < timeout
                        System.Windows.Forms.Application.DoEvents()
                        System.Threading.Thread.Sleep(50)
                    End While

                    Debug.WriteLine($"[WebView2] Loop ended. Content: {result.Length} chars, elapsed: {(DateTime.Now - startTime).TotalSeconds:F1}s")

                Catch ex As Exception
                    Debug.WriteLine($"[WebView2] Error: {ex.Message}")
                Finally
                    Try : webView?.Dispose() : Catch : End Try
                    Try : form?.Close() : form?.Dispose() : Catch : End Try
                End Try

                ' Apply character limit only if explicitly requested (maxChars > 0)
                Dim finalResult As String = result.Trim()
                If maxChars > 0 AndAlso finalResult.Length > maxChars Then
                    ' Try to cut at a sentence boundary
                    Dim cutPoint As Integer = maxChars
                    Dim lastPeriod As Integer = finalResult.LastIndexOf("."c, maxChars - 1)
                    Dim lastNewline As Integer = finalResult.LastIndexOf(vbLf, maxChars - 1)

                    If lastPeriod > maxChars * 0.8 Then
                        cutPoint = lastPeriod + 1
                    ElseIf lastNewline > maxChars * 0.8 Then
                        cutPoint = lastNewline
                    End If

                    finalResult = finalResult.Substring(0, cutPoint).Trim()
                    Debug.WriteLine($"[WebView2] Trimmed to {finalResult.Length} chars (limit was {maxChars})")
                End If

                tcs.TrySetResult(finalResult)
            End Sub)

        thread.SetApartmentState(System.Threading.ApartmentState.STA)
        thread.Start()

        Return tcs.Task
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