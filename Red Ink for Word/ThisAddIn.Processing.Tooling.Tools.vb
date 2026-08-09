' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.Processing.Tooling.Tools.vb
' Purpose: Internal tool utilities, web link extraction, and reflection helpers.
'
' Responsibilities:
'  - Build web link extraction result payloads (JSON).
'  - Extract and filter web links by file extension.
'  - Construct result metadata (requested URL, source URL, filter info).
'  - Late-bound property getters/setters via reflection.
'  - Type-safe property value conversion (int, bool, string, enum).
'  - Handle WebView2 configuration and browser automation setup.
'  - Support OCR initialization and execution.
'  - Manage large downloaded file handling (50MB limit).
'
' Architecture:
'  - Reflection-based late binding to support dynamic object manipulation.
'  - Culture-invariant type conversions (ISO date formats, numeric parsing).
'
' External Dependencies:
'  - System.Reflection for property manipulation.
'  - Newtonsoft.Json for payload construction.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    Private Const MaxDownloadedWebFileBytes As Long = 50L * 1024L * 1024L


    Const UseWebView2 = True

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


    Private Sub TrySetLateBoundProperty(target As Object, propertyName As String, value As Object)
        If target Is Nothing Then Return

        Try
            Dim prop = target.GetType().GetProperty(
            propertyName,
            BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.IgnoreCase)

            If prop Is Nothing OrElse Not prop.CanWrite Then Return

            Dim convertedValue As Object = value

            If value IsNot Nothing Then
                Dim targetType = If(Nullable.GetUnderlyingType(prop.PropertyType), prop.PropertyType)
                convertedValue = System.Convert.ChangeType(value, targetType, Globalization.CultureInfo.InvariantCulture)
            End If

            prop.SetValue(target, convertedValue, Nothing)
        Catch
        End Try
    End Sub

    Private Function TryGetLateBoundString(target As Object, propertyName As String) As String
        If target Is Nothing Then
            Return ""
        End If

        Try
            Dim prop = target.GetType().GetProperty(
            propertyName,
            BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.IgnoreCase)

            If prop Is Nothing OrElse Not prop.CanRead Then
                Return ""
            End If

            Dim value = prop.GetValue(target, Nothing)
            Return If(value, "").ToString()
        Catch
            Return ""
        End Try
    End Function


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

    Private Function GetUniqueDownloadPath(path As String) As String
        If Not File.Exists(path) Then Return path

        Dim dir = System.IO.Path.GetDirectoryName(path)
        Dim name = System.IO.Path.GetFileNameWithoutExtension(path)
        Dim ext = System.IO.Path.GetExtension(path)

        For i As Integer = 2 To 1000
            Dim candidate = System.IO.Path.Combine(dir, $"{name} ({i}){ext}")
            If Not File.Exists(candidate) Then Return candidate
        Next

        Return System.IO.Path.Combine(dir, $"{name}_{Guid.NewGuid():N}{ext}")
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
                                                               Optional cancellationToken As System.Threading.CancellationToken = Nothing) As Task(Of ToolResponse)
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
                response.ErrorMessage = "No URLs provided."
                Return response
            End If

            Dim blockedPatterns As String() = {"sharepoint.com", "onedrive.com", "1drv.ms", "teams.microsoft.com", ":f:/", "/:f:/"}
            For Each url In urls
                Dim lowerUrl = url.ToLowerInvariant()
                If blockedPatterns.Any(Function(pattern) lowerUrl.Contains(pattern)) Then
                    response.Success = False
                    response.ErrorMessage = "Authenticated SharePoint/OneDrive/Teams URLs are not supported by download_web_files."
                    Return response
                End If

                If Not IsSafeWebUrl(url) Then
                    response.Success = False
                    response.ErrorMessage = $"Blocked unsafe URL: {url}"
                    Return response
                End If
            Next

            Dim overwrite As Boolean = GetToolArgumentBoolean(toolCall.Arguments, "overwrite", False)
            Dim targetDirectory As String = GetToolArgumentString(toolCall.Arguments, "target_directory")
            Dim resolvedTargetDirectory As String = ResolveDownloadTargetDirectory(targetDirectory)
            context.Log($"Resolved download target directory: {resolvedTargetDirectory}")

            context.Log($"Downloading {urls.Count} remote file(s) to: {resolvedTargetDirectory}")

            Dim results As New JArray()

            Using handler As New System.Net.Http.HttpClientHandler() With {.AllowAutoRedirect = True}
                Using client As New System.Net.Http.HttpClient(handler)
                    client.Timeout = TimeSpan.FromSeconds(90)
                    client.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")

                    For Each url In urls
                        cancellationToken.ThrowIfCancellationRequested()

                        Dim item As New JObject(New JProperty("url", url))

                        Try
                            context.Log($"  Downloading: {url}")

                            Using request As New System.Net.Http.HttpRequestMessage(System.Net.Http.HttpMethod.Get, url)
                                Using httpResponse = Await client.SendAsync(
                                    request,
                                    System.Net.Http.HttpCompletionOption.ResponseHeadersRead,
                                    cancellationToken).ConfigureAwait(False)

                                    If Not httpResponse.IsSuccessStatusCode Then
                                        item("ok") = False
                                        item("status") = CInt(httpResponse.StatusCode)
                                        item("error") = $"HTTP {(CInt(httpResponse.StatusCode)).ToString()} {httpResponse.ReasonPhrase}"
                                        results.Add(item)
                                        Continue For
                                    End If

                                    Dim contentType As String = If(httpResponse.Content.Headers.ContentType?.MediaType, "")
                                    item("content_type") = contentType

                                    If contentType.StartsWith("text/", StringComparison.OrdinalIgnoreCase) OrElse
                                       contentType.IndexOf("html", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                                       contentType.IndexOf("xml", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                                       contentType.IndexOf("json", StringComparison.OrdinalIgnoreCase) >= 0 Then

                                        item("ok") = False
                                        item("error") = "Remote response is textual/HTML, not a downloadable binary file."
                                        results.Add(item)
                                        Continue For
                                    End If

                                    Dim bytes = Await ReadResponseBytesLimitedAsync(
                                        httpResponse.Content,
                                        MaxDownloadedWebFileBytes,
                                        cancellationToken).ConfigureAwait(False)

                                    If bytes Is Nothing OrElse bytes.Length = 0 Then
                                        item("ok") = False
                                        item("error") = "Empty response body."
                                        results.Add(item)
                                        Continue For
                                    End If

                                    If LooksLikeHtml(bytes) Then
                                        item("ok") = False
                                        item("error") = "Remote response appears to be HTML, not the original binary file."
                                        results.Add(item)
                                        Continue For
                                    End If

                                    Dim fileName As String = BuildDownloadFileName(url, httpResponse)
                                    Dim targetPath As String = Path.Combine(resolvedTargetDirectory, fileName)

                                    If Not overwrite Then
                                        targetPath = GetUniqueDownloadPath(targetPath)
                                    End If

                                    Dim ext = Path.GetExtension(targetPath).ToLowerInvariant()
                                    If ext = ".pdf" OrElse String.Equals(contentType, "application/pdf", StringComparison.OrdinalIgnoreCase) Then
                                        If Not LooksLikePdf(bytes) Then
                                            item("ok") = False
                                            item("error") = "Downloaded content does not have a valid PDF signature."
                                            results.Add(item)
                                            Continue For
                                        End If
                                    End If

                                    File.WriteAllBytes(targetPath, bytes)

                                    item("ok") = True
                                    item("path") = targetPath
                                    item("file_name") = Path.GetFileName(targetPath)
                                    item("size_bytes") = bytes.Length
                                    results.Add(item)
                                End Using
                            End Using

                        Catch ex As OperationCanceledException
                            Throw
                        Catch ex As Exception
                            item("ok") = False
                            item("error") = ex.Message
                            results.Add(item)
                        End Try
                    Next
                End Using
            End Using

            response.Success = True
            response.Response = New JObject(
                        New JProperty("target_directory", resolvedTargetDirectory),
                        New JProperty("results", results)
                    ).ToString(Formatting.None)
            Return response

        Catch ex As OperationCanceledException
            response.Success = False
            response.ErrorMessage = "Operation was cancelled"
            Return response
        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            Return response
        End Try
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


    Private Async Function ExecuteWebGroundingTool(
        toolCall As ToolCall,
        context As ToolExecutionContext,
        Optional cancellationToken As CancellationToken = Nothing) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
        .CallId = toolCall.CallId,
        .ToolName = toolCall.ToolName
    }

        Try
            response.Response =
            Await SharedLibrary.Agents.WebGroundingTool.ExecuteAsync(
                _context,
                toolCall.Arguments,
                cancellationToken,
                logStep:=Sub(message)
                             context.Log(message)
                         End Sub,
                logInfo:=Sub(message)
                             context.Log(message)
                         End Sub,
                logWarn:=Sub(message)
                             context.Log(message, "warn")
                         End Sub)

            response.Success = Not String.IsNullOrWhiteSpace(response.Response)

            If Not response.Success Then
                response.ErrorMessage = "web_grounding returned no usable result."
                response.Response = response.ErrorMessage
            End If

        Catch ex As OperationCanceledException
            response.Success = False
            response.ErrorMessage = "Operation was cancelled."
            response.Response = response.ErrorMessage

        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error during web grounding: {ex.Message}"
            context.Log(response.Response, "warn")
        End Try

        Return response
    End Function

    ''' <summary>
    ''' Executes the internal web retrieval tool by fetching content for one or more URLs and returning tagged content blocks.
    ''' </summary>
    ''' <param name="toolCall">Tool call payload containing <c>url</c> or <c>urls</c> arguments.</param>
    ''' <param name="context">Tool execution context for logging and diagnostics.</param>
    ''' <returns>Tool response containing retrieved content or an error.</returns>
    Private Async Function ExecuteInternalWebTool(toolCall As ToolCall, context As ToolExecutionContext) As Task(Of ToolResponse)
        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId,
            .ToolName = toolCall.ToolName
        }

        Try
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
                    "Please ask the user to download the file(s) and provide them directly."
                context.Log($"Blocked SharePoint/OneDrive URL(s): {blockedList}", "warn")
                ToolingFileLogger.LogWarn("Internal web tool: SharePoint/OneDrive URL blocked.", details:=$"urls={blockedList}")
                Return response
            End If
            context.Log($"Retrieving content from {urls.Count} URL(s)...")

            Dim results As New StringBuilder()

            If UseWebView2 Then
                For i = 0 To urls.Count - 1
                    Dim requestedUrl = urls(i)

                    Try
                        context.Log($"  Fetching: {requestedUrl}")

                        Dim pageResult = Await RetrieveWebsiteContent_WebView2Detailed(
                            requestedUrl,
                            0,
                            expandCollapsed:=expandInteractiveSections,
                            includeLinks:=includeLinks,
                            linkExtensions:=linkExtensions)

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
                Using httpClient As New HttpClient()
                    httpClient.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
                    httpClient.Timeout = TimeSpan.FromSeconds(30)

                    For i = 0 To urls.Count - 1
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
                                        "Structured link extraction is unavailable in the HTTP fallback path."))
                                results.AppendLine($"</LINKS_{i + 1}>")
                            End If

                            results.AppendLine()

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

        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            ToolingFileLogger.LogError("Internal web tool error.", ex:=ex)
        End Try

        Return response
    End Function


    ''' <summary>
    ''' Executes the internal internet search tool by querying the configured search engine,
    ''' extracting result URLs via response masks, fetching qualifying page content, and returning
    ''' tagged result blocks including the search query and all visited URLs for transparency.
    ''' </summary>
    ''' <param name="toolCall">Tool call payload containing <c>query</c> and optional <c>max_results</c>/<c>max_depth</c>.</param>
    ''' <param name="context">Tool execution context for logging and diagnostics.</param>
    ''' <returns>Tool response containing search results or an error.</returns>
    Private Async Function ExecuteInternalSearchTool(toolCall As ToolCall, context As ToolExecutionContext) As Task(Of ToolResponse)
        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId,
            .ToolName = toolCall.ToolName
        }

        Try
            ' ── Validate search configuration ────────────────────────────
            If Not INI_ISearch OrElse String.IsNullOrWhiteSpace(INI_ISearch_URL) Then
                response.Success = False
                response.ErrorMessage = "Internet search is not configured or not enabled."
                ToolingFileLogger.LogWarn("Internal search tool: search not enabled/configured.",
                    details:=$"INI_ISearch={INI_ISearch}; INI_ISearch_URL='{INI_ISearch_URL}'")
                Return response
            End If

            ' ── Extract and validate parameters ──────────────────────────
            Dim query As String = ""
            If toolCall.Arguments.ContainsKey("query") Then
                query = If(toolCall.Arguments("query")?.ToString(), "").Trim()
            End If

            If String.IsNullOrWhiteSpace(query) Then
                response.Success = False
                response.ErrorMessage = "No search query provided."
                ToolingFileLogger.LogWarn("Internal search tool: empty query.",
                    details:=$"CallId={toolCall.CallId}; Args={JsonConvert.SerializeObject(toolCall.Arguments)}")
                Return response
            End If

            ' ── PII / confidential data safety net ───────────────────────
            ' Block queries that contain obvious personal data patterns.
            ' This is a last-resort filter; the model is instructed not to include such data,
            ' but defense-in-depth requires a code-level check before the query leaves the system.
            Dim piiPatterns As String() = {
                "\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b",
                "\b\d{3}[-.\s]?\d{3}[-.\s]?\d{4}\b",
                "\b\+?\d{1,3}[-.\s]?\(?\d{1,4}\)?[-.\s]?\d{2,4}[-.\s]?\d{2,4}[-.\s]?\d{0,4}\b",
                "\b\d{3}-\d{2}-\d{4}\b",
                "\b\d{2}[\./]\d{2}[\./]\d{2,4}\b(?=.*\d{2}[\./]\d{2}[\./]\d{2,4})",
                "\b[A-Z]{2}\d{2}\s?\d{4}\s?\d{4}\s?\d{4}\s?\d{4}\s?\d{0,2}\b",
                "\b(?:4\d{3}|5[1-5]\d{2}|6011|3[47]\d{2})[-\s]?\d{4}[-\s]?\d{4}[-\s]?\d{4}\b",
                "\bAHV[\s-]?\d{3}[\.\s]?\d{4}[\.\s]?\d{4}[\.\s]?\d{2}\b"
            }

            For Each piiPattern In piiPatterns
                If Regex.IsMatch(query, piiPattern, RegexOptions.IgnoreCase) Then
                    response.Success = False
                    response.ErrorMessage = "Search query blocked: appears to contain personal or confidential data."
                    ToolingFileLogger.LogWarn("Internal search tool: query blocked by PII filter.",
                        details:=$"CallId={toolCall.CallId}; Pattern='{piiPattern}'")
                    context.Log("  ⚠ Search query blocked — contains data that appears personal or confidential.", "warn")
                    Return response
                End If
            Next

            ' Clamp max_results to server limit (INI_ISearch_Tries)
            Dim maxResults As Integer = INI_ISearch_Results
            If toolCall.Arguments.ContainsKey("max_results") Then
                Dim requested As Integer
                If Integer.TryParse(toolCall.Arguments("max_results")?.ToString(), requested) AndAlso requested > 0 Then
                    maxResults = Math.Min(requested, INI_ISearch_Tries)
                End If
            End If

            ' Clamp max_depth to server limit (INI_ISearch_MaxDepth)
            Dim maxDepth As Integer = 0
            If toolCall.Arguments.ContainsKey("max_depth") Then
                Dim requested As Integer
                If Integer.TryParse(toolCall.Arguments("max_depth")?.ToString(), requested) AndAlso requested >= 0 Then
                    maxDepth = Math.Min(requested, INI_ISearch_MaxDepth)
                End If
            End If

            context.Log($"Internet search: query='{query}', max_results={maxResults}, max_depth={maxDepth}")
            ToolingFileLogger.LogStep($"Search query: '{query}'; max_results={maxResults}; max_depth={maxDepth}; engine={INI_ISearch_Name}")

            ' ── Perform the HTTP search request ──────────────────────────
            Dim searchUrl As String = INI_ISearch_URL & Uri.EscapeDataString(query)
            context.Log($"  Search URL: {searchUrl}")

            Dim searchResponse As String = ""
            Using httpClient As New HttpClient()
                httpClient.DefaultRequestHeaders.UserAgent.ParseAdd(
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36")
                httpClient.Timeout = TimeSpan.FromSeconds(30)

                searchResponse = Await httpClient.GetStringAsync(searchUrl)
            End Using

            If String.IsNullOrWhiteSpace(searchResponse) Then
                response.Success = False
                response.ErrorMessage = "Search engine returned an empty response."
                ToolingFileLogger.LogWarn("Internal search tool: empty search response.",
                    details:=$"searchUrl={searchUrl}")
                Return response
            End If

            ' ── Extract unique URLs using response masks ─────────────────
            Dim urlPattern As String = Regex.Escape(INI_ISearch_ResponseMask1) & "(.*?)" & Regex.Escape(INI_ISearch_ResponseMask2)
            Dim matches As MatchCollection = Regex.Matches(searchResponse, urlPattern)

            Dim extractedUrls As New List(Of String)()
            For Each m As Match In matches
                Dim rawUrl As String = m.Groups(1).Value
                Dim decodedUrl As String = WebUtility.UrlDecode(rawUrl.Replace(INI_ISearch_ResponseMask1, ""))

                If Not extractedUrls.Contains(decodedUrl) AndAlso IsSafeWebUrl(decodedUrl) Then
                    extractedUrls.Add(decodedUrl)
                End If

                If extractedUrls.Count >= INI_ISearch_Tries Then Exit For
            Next

            context.Log($"  Extracted {extractedUrls.Count} unique URL(s) from search results")
            ToolingFileLogger.LogStep($"Extracted URLs: {extractedUrls.Count}")

            If extractedUrls.Count = 0 Then
                response.Success = False
                response.ErrorMessage = "No result URLs could be extracted from the search engine response."
                ToolingFileLogger.LogWarn("Internal search tool: no URLs extracted.",
                    details:=$"searchUrl={searchUrl}; ResponseMask1='{INI_ISearch_ResponseMask1}'; ResponseMask2='{INI_ISearch_ResponseMask2}'")
                Return response
            End If

            ' ── Fetch content from each result URL ───────────────────────
            Dim results As New StringBuilder()
            Dim visitedUrls As New List(Of String)()
            Dim resultIndex As Integer = 0

            ' Header: report the search query and engine
            results.AppendLine($"<SEARCH_QUERY>{query}</SEARCH_QUERY>")
            results.AppendLine($"<SEARCH_ENGINE>{If(INI_ISearch_Name, "Search")}</SEARCH_ENGINE>")
            results.AppendLine()

            For Each url In extractedUrls
                If resultIndex >= maxResults Then Exit For

                Try
                    context.Log($"  Fetching result: {url}")
                    visitedUrls.Add(url)

                    Dim content As String = ""

                    If UseWebView2 Then
                        content = Await RetrieveWebsiteContent_WebView2(url, ISearch_MaxChars)
                    Else
                        Using httpClient As New HttpClient()
                            httpClient.DefaultRequestHeaders.UserAgent.ParseAdd(
                                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
                            httpClient.Timeout = TimeSpan.FromSeconds(30)
                            content = Await RetrieveWebsiteContent(url, maxDepth, httpClient)
                        End Using
                    End If

                    ' Apply character cap (ISearch_MaxChars) for WebView2 results that exceed it
                    If Not String.IsNullOrWhiteSpace(content) AndAlso ISearch_MaxChars > 0 AndAlso content.Length > ISearch_MaxChars Then
                        content = content.Substring(0, ISearch_MaxChars)
                    End If

                    ' Discard noise (pages shorter than ISearch_MinChars)
                    If Not String.IsNullOrWhiteSpace(content) AndAlso content.Length >= ISearch_MinChars Then
                        resultIndex += 1
                        results.AppendLine($"<SEARCHRESULT_{resultIndex}_URL>{url}</SEARCHRESULT_{resultIndex}_URL>")
                        results.AppendLine($"<SEARCHRESULT_{resultIndex}>")
                        results.AppendLine(content)
                        results.AppendLine($"</SEARCHRESULT_{resultIndex}>")
                        results.AppendLine()
                        context.Log($"  Result #{resultIndex}: {content.Length} chars from {url}")
                    Else
                        Dim charCount = If(content Is Nothing, 0, content.Length)
                        context.Log($"  Skipped (too short: {charCount} chars, min {ISearch_MinChars}): {url}")
                        ToolingFileLogger.LogStep($"Search result skipped (too short: {charCount} < {ISearch_MinChars}): {url}")
                    End If

                Catch ex As Exception
                    context.Log($"  Error fetching {url}: {ex.Message}")
                    ToolingFileLogger.LogError("Internal search tool fetch error.", details:=$"url={url}", ex:=ex)
                End Try
            Next

            ' Footer: report all visited URLs for transparency
            results.AppendLine("<URLS_VISITED>")
            For Each vUrl In visitedUrls
                results.AppendLine($"  {vUrl}")
            Next
            results.AppendLine("</URLS_VISITED>")

            context.Log($"Search complete: {resultIndex} qualifying result(s) from {visitedUrls.Count} URL(s) visited")

            response.Response = results.ToString()
            response.Success = True

        Catch ex As HttpRequestException
            response.Success = False
            response.ErrorMessage = $"Search HTTP error: {ex.Message}"
            ToolingFileLogger.LogError("Internal search tool HTTP error.", ex:=ex)

        Catch ex As TaskCanceledException
            response.Success = False
            response.ErrorMessage = $"Search request timed out: {ex.Message}"
            ToolingFileLogger.LogError("Internal search tool timeout.", ex:=ex)

        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            ToolingFileLogger.LogError("Internal search tool error.", ex:=ex)
        End Try

        Return response
    End Function



End Class
