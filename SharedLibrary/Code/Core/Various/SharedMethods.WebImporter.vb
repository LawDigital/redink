' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: SharedMethods.WebImporter.vb
' Purpose: Shared URL-import helpers for drag/drop scenarios across Word, Excel,
'          and Outlook.
'
' Behavior:
'  - Accepts HTTP/HTTPS URLs only.
'  - If the URL points to a downloadable document, downloads the original bytes
'    to a temp file and returns that path.
'  - If the URL points to a website/HTML page, uses the shared WebView2 sandbox
'    to render/extract visible text and stores that text in a temp .txt file.
'  - Also provides shared drag/drop URL extraction helpers for browser drops and
'    Windows .url InternetShortcut files.
'  - Does not expose preliminary website response sizes, because those often only
'    reflect the initial HTML payload and can be misleading compared to the final
'    rendered/extracted text size.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Net.Sockets
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Newtonsoft.Json.Linq

Namespace SharedLibrary
    Partial Public Class SharedMethods

        Private Class WebImportPageResult
            Public Property TextContent As String = ""
            Public Property FinalUrl As String = ""
        End Class

        ''' <summary>
        ''' Returns the default set of supported URL-import extensions for Office hosts.
        ''' This list is intended for remote file downloads, not HTML pages.
        ''' </summary>
        Public Shared Function GetDefaultWebImportSupportedExtensions(Optional includeLegacyDoc As Boolean = False,
                                                                     Optional includeLegacyExcelPowerPoint As Boolean = False) As List(Of String)
            Dim extensions As New List(Of String) From {
                ".txt", ".ini", ".csv", ".log", ".json", ".xml", ".md", ".yaml", ".yml",
                ".vb", ".cs", ".js", ".ts", ".py", ".java", ".cpp", ".c", ".h", ".sql",
                ".rtf", ".docx", ".xlsx", ".pptx", ".pdf", ".msg", ".eml"
            }

            If includeLegacyDoc Then
                extensions.Add(".doc")
            End If

            If includeLegacyExcelPowerPoint Then
                extensions.Add(".xls")
                extensions.Add(".ppt")
            End If

            Return extensions
        End Function

        ''' <summary>
        ''' Retrieves content from a web URL and stores it as a temporary file.
        ''' Documents are downloaded as original files; websites are retrieved via WebView2 and stored as text.
        ''' </summary>
        Public Shared Async Function CreateTempFileFromUrlAsync(url As String,
                                                                supportedExtensions As IEnumerable(Of String)) As Task(Of String)
            Try
                Dim normalizedUrl As String = NormalizeWebImportUrl(url)
                If normalizedUrl = "" Then
                    Return ""
                End If

                Dim allowedExtensions As HashSet(Of String) = NormalizeWebImportExtensions(supportedExtensions)

                For attempt As Integer = 1 To 2
                    Dim downloadedFilePath As String =
                        Await TryDownloadUrlDocumentToTempFileAsync(normalizedUrl, allowedExtensions).ConfigureAwait(False)

                    If Not String.IsNullOrWhiteSpace(downloadedFilePath) AndAlso File.Exists(downloadedFilePath) Then
                        Return downloadedFilePath
                    End If

                    Dim pageResult As WebImportPageResult =
                        Await RetrieveWebsiteTextAsync(normalizedUrl).ConfigureAwait(False)

                    If pageResult IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(pageResult.TextContent) Then
                        Dim sourceUrl As String = normalizedUrl
                        If Not String.IsNullOrWhiteSpace(pageResult.FinalUrl) Then
                            sourceUrl = pageResult.FinalUrl
                        End If

                        Dim tempTextPath As String =
                            Path.Combine(Path.GetTempPath(), $"{BuildWebImportTextFileStem(sourceUrl)}_{Guid.NewGuid():N}.txt")

                        Dim tempTextContent As String = "Source URL: " & sourceUrl & vbCrLf & vbCrLf & pageResult.TextContent
                        File.WriteAllText(tempTextPath, tempTextContent, Encoding.UTF8)

                        Return tempTextPath
                    End If

                    If attempt < 2 Then
                        Await Task.Delay(1500).ConfigureAwait(False)
                    End If
                Next

                Return ""

            Catch ex As Exception
                Debug.WriteLine($"CreateTempFileFromUrlAsync failed for '{url}': {ex.Message}")
                Return ""
            End Try
        End Function

        ''' <summary>
        ''' Attempts to extract a dropped HTTP/HTTPS link from drag data.
        ''' </summary>
        Public Shared Function TryGetDroppedInternetLink(data As IDataObject, ByRef url As String) As Boolean
            url = ""

            If data Is Nothing Then
                Return False
            End If

            Dim candidate As String = TryGetUrlFromDataFormat(data, "UniformResourceLocatorW")
            If IsSupportedInternetUrl(candidate, url) Then
                Return True
            End If

            candidate = TryGetUrlFromDataFormat(data, "UniformResourceLocator")
            If IsSupportedInternetUrl(candidate, url) Then
                Return True
            End If

            If data.GetDataPresent(DataFormats.UnicodeText) Then
                candidate = TryCast(data.GetData(DataFormats.UnicodeText), String)
                If IsSupportedInternetUrl(candidate, url) Then
                    Return True
                End If
            End If

            If data.GetDataPresent(DataFormats.Text) Then
                candidate = TryCast(data.GetData(DataFormats.Text), String)
                If IsSupportedInternetUrl(candidate, url) Then
                    Return True
                End If
            End If

            Return False
        End Function

        ''' <summary>
        ''' Attempts to extract the URL from a Windows internet-shortcut file.
        ''' </summary>
        Public Shared Function TryReadInternetShortcutUrl(shortcutPath As String, ByRef url As String) As Boolean
            url = ""

            Try
                If String.IsNullOrWhiteSpace(shortcutPath) OrElse Not File.Exists(shortcutPath) Then
                    Return False
                End If

                If Not String.Equals(Path.GetExtension(shortcutPath), ".url", StringComparison.OrdinalIgnoreCase) Then
                    Return False
                End If

                For Each line As String In File.ReadAllLines(shortcutPath)
                    If line.StartsWith("URL=", StringComparison.OrdinalIgnoreCase) Then
                        Return IsSupportedInternetUrl(line.Substring(4), url)
                    End If
                Next
            Catch
            End Try

            Return False
        End Function

        ''' <summary>
        ''' Validates and normalizes an HTTP/HTTPS URL.
        ''' </summary>
        Public Shared Function IsSupportedInternetUrl(candidate As String, ByRef normalizedUrl As String) As Boolean
            normalizedUrl = NormalizeWebImportUrl(candidate)
            Return normalizedUrl <> ""
        End Function

        Private Shared Function NormalizeWebImportUrl(url As String) As String
            Try
                Dim candidate As String = RemoveCR(If(url, "").Trim())
                If candidate = "" Then
                    Return ""
                End If

                Dim uriValue As Uri = Nothing
                If Not Uri.TryCreate(candidate, UriKind.Absolute, uriValue) Then
                    Return ""
                End If

                If uriValue.Scheme <> Uri.UriSchemeHttp AndAlso uriValue.Scheme <> Uri.UriSchemeHttps Then
                    Return ""
                End If

                If Not IsSafeWebImportUrl(uriValue.AbsoluteUri) Then
                    Return ""
                End If

                Return uriValue.AbsoluteUri
            Catch
                Return ""
            End Try
        End Function

        Private Shared Function NormalizeWebImportExtensions(supportedExtensions As IEnumerable(Of String)) As HashSet(Of String)
            Dim result As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            If supportedExtensions Is Nothing Then
                Return result
            End If

            For Each extension As String In supportedExtensions
                Dim candidate As String = If(extension, "").Trim().ToLowerInvariant()
                If candidate = "" Then
                    Continue For
                End If

                If Not candidate.StartsWith(".") Then
                    candidate = "." & candidate
                End If

                result.Add(candidate)
            Next

            Return result
        End Function

        Private Shared Async Function TryDownloadUrlDocumentToTempFileAsync(url As String,
                                                                            allowedExtensions As HashSet(Of String)) As Task(Of String)
            Try
                Using client As New HttpClient()
                    client.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
                    client.Timeout = TimeSpan.FromSeconds(60)

                    Using response As HttpResponseMessage =
                        Await client.GetAsync(url, HttpCompletionOption.ResponseHeadersRead).ConfigureAwait(False)

                        If response Is Nothing OrElse Not response.IsSuccessStatusCode OrElse response.Content Is Nothing Then
                            Return ""
                        End If

                        If response.Content.Headers.ContentLength.HasValue AndAlso
                           response.Content.Headers.ContentLength.Value > DEFAULT_WEBIMPORT_MAX_DOWNLOAD_BYTES Then
                            Return ""
                        End If

                        Dim effectiveUrl As String = url
                        If response.RequestMessage IsNot Nothing AndAlso response.RequestMessage.RequestUri IsNot Nothing Then
                            effectiveUrl = response.RequestMessage.RequestUri.AbsoluteUri
                        End If

                        Dim mediaType As String = If(response.Content.Headers.ContentType?.MediaType, "").Trim().ToLowerInvariant()

                        If mediaType = "text/html" OrElse mediaType = "application/xhtml+xml" Then
                            Return ""
                        End If

                        Dim urlExtension As String = GetWebImportDownloadExtensionFromUrl(effectiveUrl, allowedExtensions)
                        Dim dispositionFileName As String = GetWebImportContentDispositionFileName(response)
                        Dim dispositionExtension As String = Path.GetExtension(If(dispositionFileName, "")).ToLowerInvariant()
                        Dim contentTypeExtension As String = GetWebImportDownloadExtensionFromContentType(mediaType)

                        Dim finalExtension As String = ""

                        If IsDownloadableWebImportExtension(urlExtension, allowedExtensions) Then
                            finalExtension = urlExtension
                        ElseIf IsDownloadableWebImportExtension(dispositionExtension, allowedExtensions) Then
                            finalExtension = dispositionExtension
                        ElseIf IsDownloadableWebImportExtension(contentTypeExtension, allowedExtensions) Then
                            finalExtension = contentTypeExtension
                        End If

                        If finalExtension = "" Then
                            Return ""
                        End If

                        Dim fileBytes As Byte() =
                            Await ReadWebImportResponseBytesLimitedAsync(
                                response.Content,
                                DEFAULT_WEBIMPORT_MAX_DOWNLOAD_BYTES,
                                CancellationToken.None).ConfigureAwait(False)

                        If fileBytes Is Nothing OrElse fileBytes.Length = 0 Then
                            Return ""
                        End If

                        If LooksLikeWebImportHtml(fileBytes) Then
                            Return ""
                        End If

                        Dim fileName As String = BuildWebImportDownloadFileName(effectiveUrl, response, finalExtension)
                        Dim tempFilePath As String = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}_{fileName}")

                        File.WriteAllBytes(tempFilePath, fileBytes)
                        Return tempFilePath
                    End Using
                End Using

            Catch ex As Exception
                Debug.WriteLine($"TryDownloadUrlDocumentToTempFileAsync failed for '{url}': {ex.Message}")
                Return ""
            End Try
        End Function

        Private Shared Function GetWebImportDownloadExtensionFromUrl(url As String,
                                                                     allowedExtensions As HashSet(Of String)) As String
            Try
                Dim uriValue As Uri = Nothing
                If Not Uri.TryCreate(url, UriKind.Absolute, uriValue) Then
                    Return ""
                End If

                Dim extension As String = Path.GetExtension(uriValue.AbsolutePath).ToLowerInvariant()
                If IsDownloadableWebImportExtension(extension, allowedExtensions) Then
                    Return extension
                End If

                Dim queryText As String = If(uriValue.Query, "").ToLowerInvariant()

                For Each candidateExtension As String In allowedExtensions
                    If IsWebsiteMarkupExtension(candidateExtension) Then
                        Continue For
                    End If

                    If queryText.Contains(candidateExtension) Then
                        Return candidateExtension
                    End If
                Next
            Catch
            End Try

            Return ""
        End Function

        Private Shared Function GetWebImportDownloadExtensionFromContentType(contentType As String) As String
            Select Case If(contentType, "").Trim().ToLowerInvariant()
                Case "text/plain"
                    Return ".txt"
                Case "text/csv", "application/csv"
                    Return ".csv"
                Case "text/tab-separated-values"
                    Return ".tsv"
                Case "application/json", "text/json"
                    Return ".json"
                Case "application/xml", "text/xml"
                    Return ".xml"
                Case "text/markdown"
                    Return ".md"
                Case "application/rtf", "text/rtf"
                    Return ".rtf"
                Case "application/pdf"
                    Return ".pdf"
                Case "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    Return ".docx"
                Case "application/msword"
                    Return ".doc"
                Case "application/vnd.ms-excel"
                    Return ".xls"
                Case "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    Return ".xlsx"
                Case "application/vnd.ms-powerpoint"
                    Return ".ppt"
                Case "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    Return ".pptx"
                Case "message/rfc822"
                    Return ".eml"
                Case "application/vnd.ms-outlook"
                    Return ".msg"
                Case Else
                    Return ""
            End Select
        End Function

        Private Shared Function IsDownloadableWebImportExtension(extension As String,
                                                                 allowedExtensions As HashSet(Of String)) As Boolean
            Dim candidate As String = If(extension, "").Trim().ToLowerInvariant()

            If candidate = "" Then
                Return False
            End If

            If IsWebsiteMarkupExtension(candidate) Then
                Return False
            End If

            If allowedExtensions Is Nothing OrElse allowedExtensions.Count = 0 Then
                Return False
            End If

            Return allowedExtensions.Contains(candidate)
        End Function

        Private Shared Function IsWebsiteMarkupExtension(extension As String) As Boolean
            Dim candidate As String = If(extension, "").Trim().ToLowerInvariant()
            Return candidate = ".html" OrElse candidate = ".htm"
        End Function

        Private Shared Function GetWebImportContentDispositionFileName(response As HttpResponseMessage) As String
            Try
                If response Is Nothing OrElse response.Content Is Nothing Then
                    Return ""
                End If

                Dim disposition = response.Content.Headers.ContentDisposition
                If disposition Is Nothing Then
                    Return ""
                End If

                If Not String.IsNullOrWhiteSpace(disposition.FileNameStar) Then
                    Return disposition.FileNameStar.Trim().Trim(""""c)
                End If

                If Not String.IsNullOrWhiteSpace(disposition.FileName) Then
                    Return disposition.FileName.Trim().Trim(""""c)
                End If
            Catch
            End Try

            Return ""
        End Function

        Private Shared Function BuildWebImportDownloadFileName(url As String,
                                                               response As HttpResponseMessage,
                                                               fallbackExtension As String) As String
            Dim candidate As String = GetWebImportContentDispositionFileName(response)

            If String.IsNullOrWhiteSpace(candidate) Then
                Try
                    candidate = Path.GetFileName(New Uri(url).LocalPath)
                Catch
                    candidate = ""
                End Try
            End If

            candidate = If(candidate, "").Trim().Trim(""""c)

            For Each invalidChar As Char In Path.GetInvalidFileNameChars()
                candidate = candidate.Replace(invalidChar, "_"c)
            Next

            If String.IsNullOrWhiteSpace(candidate) Then
                candidate = "download" & fallbackExtension
            End If

            Dim currentExtension As String = Path.GetExtension(candidate).ToLowerInvariant()

            If currentExtension = "" Then
                candidate &= fallbackExtension
            ElseIf IsWebsiteMarkupExtension(currentExtension) Then
                candidate = Path.GetFileNameWithoutExtension(candidate) & fallbackExtension
            ElseIf currentExtension <> fallbackExtension Then
                candidate = Path.GetFileNameWithoutExtension(candidate) & fallbackExtension
            End If

            Return candidate
        End Function

        Private Shared Function BuildWebImportTextFileStem(url As String) As String
            Dim candidate As String = "redink-url"

            Try
                Dim uriValue As Uri = Nothing
                If Uri.TryCreate(url, UriKind.Absolute, uriValue) Then
                    candidate = $"{uriValue.Host}{uriValue.AbsolutePath}".Trim("/"c).Replace("/"c, "_"c)
                End If
            Catch
            End Try

            If String.IsNullOrWhiteSpace(candidate) Then
                candidate = "redink-url"
            End If

            For Each invalidChar As Char In Path.GetInvalidFileNameChars()
                candidate = candidate.Replace(invalidChar, "_"c)
            Next

            If candidate.Length > 80 Then
                candidate = candidate.Substring(0, 80)
            End If

            Return candidate
        End Function

        Private Shared Async Function ReadWebImportResponseBytesLimitedAsync(content As HttpContent,
                                                                             maxBytes As Long,
                                                                             cancellationToken As CancellationToken) As Task(Of Byte())
            Using sourceStream = Await content.ReadAsStreamAsync().ConfigureAwait(False)
                Using ms As New MemoryStream()
                    Dim buffer(8191) As Byte

                    Do
                        cancellationToken.ThrowIfCancellationRequested()

                        Dim read As Integer =
                            Await sourceStream.ReadAsync(buffer, 0, buffer.Length, cancellationToken).ConfigureAwait(False)

                        If read <= 0 Then
                            Exit Do
                        End If

                        ms.Write(buffer, 0, read)

                        If ms.Length > maxBytes Then
                            Throw New InvalidOperationException($"Remote file exceeds the maximum allowed size of {maxBytes} bytes.")
                        End If
                    Loop

                    Return ms.ToArray()
                End Using
            End Using
        End Function

        Private Shared Function LooksLikeWebImportHtml(bytes As Byte()) As Boolean
            If bytes Is Nothing OrElse bytes.Length = 0 Then
                Return False
            End If

            Dim sampleLength As Integer = Math.Min(bytes.Length, 1024)
            Dim sample As String = Encoding.UTF8.GetString(bytes, 0, sampleLength).ToLowerInvariant()

            Return sample.Contains("<html") OrElse
                   sample.Contains("<!doctype html") OrElse
                   sample.Contains("<body") OrElse
                   sample.Contains("<head")
        End Function

        Private Shared Async Function RetrieveWebsiteTextAsync(url As String) As Task(Of WebImportPageResult)
            Dim result As New WebImportPageResult() With {
                .FinalUrl = url,
                .TextContent = ""
            }

            Try
                If Not Agents.WebView2JsSandbox.IsConfigured Then
                    Return result
                End If

                Dim rawJson As String =
                    Await Agents.WebView2JsSandbox.RunAsync(
                        code:=BuildWebImportExtractionScript(),
                        timeoutMs:=90000,
                        allowNetwork:=True,
                        navigateUrl:=url,
                        waitAfterLoadMs:=1500,
                        waitForSelector:="body").ConfigureAwait(False)

                If String.IsNullOrWhiteSpace(rawJson) Then
                    Return result
                End If

                Dim root As JObject = JObject.Parse(rawJson)
                Dim okToken As JToken = root("ok")

                If okToken Is Nothing OrElse okToken.Type <> JTokenType.Boolean OrElse Not okToken.Value(Of Boolean)() Then
                    Return result
                End If

                Dim resultToken As JToken = root("result")
                If resultToken Is Nothing Then
                    Return result
                End If

                If resultToken.Type = JTokenType.Object Then
                    Dim sourceToken As JToken = resultToken("source_url")
                    If sourceToken IsNot Nothing Then
                        result.FinalUrl = sourceToken.ToString()
                    End If

                    Dim textToken As JToken = resultToken("text")
                    If textToken IsNot Nothing Then
                        result.TextContent = textToken.ToString().Trim()
                    End If
                ElseIf resultToken.Type = JTokenType.String Then
                    result.TextContent = resultToken.ToString().Trim()
                End If

                Return result

            Catch ex As Exception
                Debug.WriteLine($"RetrieveWebsiteTextAsync failed for '{url}': {ex.Message}")
                Return result
            End Try
        End Function

        Private Shared Function BuildWebImportExtractionScript() As String
            Return <![CDATA[
const delay = ms => new Promise(resolve => setTimeout(resolve, ms));

function isVisible(el) {
    try {
        if (!el) return false;
        const style = window.getComputedStyle(el);
        return style &&
               style.display !== 'none' &&
               style.visibility !== 'hidden' &&
               style.opacity !== '0';
    } catch (e) {
        return false;
    }
}

function textOf(el) {
    try {
        return String((el.innerText || el.textContent || '') + ' ' + (el.getAttribute('aria-label') || '')).toLowerCase();
    } catch (e) {
        return '';
    }
}

function tryExpandOnce() {
    let clicked = 0;

    try {
        const nodes = document.querySelectorAll('button, summary, [role="button"], [aria-expanded="false"], [aria-label]');

        for (const node of nodes) {
            if (clicked >= 50) break;
            if (!isVisible(node)) continue;

            const text = textOf(node);
            const shouldClick =
                node.getAttribute('aria-expanded') === 'false' ||
                text.includes('show more') ||
                text.includes('read more') ||
                text.includes('expand') ||
                text.includes('continue reading') ||
                text.includes('load more') ||
                text.includes('mehr anzeigen') ||
                text.includes('mehr lesen');

            if (!shouldClick) continue;

            try {
                node.click();
                clicked++;
            } catch (e) {
            }
        }
    } catch (e) {
    }

    return clicked;
}

async function autoScroll() {
    let lastHeight = -1;

    for (let i = 0; i < 10; i++) {
        window.scrollTo(0, document.body ? document.body.scrollHeight : 0);
        await delay(700);

        const currentHeight = Math.max(
            document.body ? document.body.scrollHeight : 0,
            document.documentElement ? document.documentElement.scrollHeight : 0
        );

        if (currentHeight === lastHeight) {
            break;
        }

        lastHeight = currentHeight;
    }

    window.scrollTo(0, 0);
    await delay(250);
}

tryExpandOnce();
await delay(400);
await autoScroll();
tryExpandOnce();
await delay(800);

let text = '';

try {
    text = document.body && typeof document.body.innerText === 'string'
        ? document.body.innerText
        : '';
} catch (e) {
    text = '';
}

text = String(text || '')
    .replace(/\u00A0/g, ' ')
    .replace(/\r/g, '')
    .replace(/\n{3,}/g, '\n\n')
    .replace(/[ \t]+\n/g, '\n')
    .trim();

return {
    source_url: window.location && window.location.href ? window.location.href : '',
    text: text
};
]]>.Value
        End Function

        Private Shared Function TryGetUrlFromDataFormat(data As IDataObject, format As String) As String
            Try
                If Not data.GetDataPresent(format) Then
                    Return ""
                End If

                Dim rawData As Object = data.GetData(format)

                If TypeOf rawData Is String Then
                    Return DirectCast(rawData, String)
                End If

                Dim memoryStream As MemoryStream = TryCast(rawData, MemoryStream)
                If memoryStream IsNot Nothing Then
                    Dim bytes As Byte() = memoryStream.ToArray()
                    Dim unicodeText As String = Encoding.Unicode.GetString(bytes).TrimEnd(ChrW(0))
                    If unicodeText <> "" Then
                        Return unicodeText
                    End If

                    Return Encoding.ASCII.GetString(bytes).TrimEnd(ChrW(0))
                End If
            Catch
            End Try

            Return ""
        End Function

        Private Shared Function IsSafeWebImportUrl(url As String) As Boolean
            Try
                Dim uriResult As Uri = Nothing
                If Not Uri.TryCreate(url, UriKind.Absolute, uriResult) Then
                    Return False
                End If

                If uriResult.Scheme <> Uri.UriSchemeHttp AndAlso uriResult.Scheme <> Uri.UriSchemeHttps Then
                    Return False
                End If

                If uriResult.IsLoopback Then
                    Return False
                End If

                If Not String.IsNullOrWhiteSpace(uriResult.UserInfo) Then
                    Return False
                End If

                Dim host As String = If(uriResult.Host, "").Trim().ToLowerInvariant()
                If host = "" Then
                    Return False
                End If

                If host = "localhost" OrElse
                   host.EndsWith(".local", StringComparison.OrdinalIgnoreCase) OrElse
                   host.EndsWith(".internal", StringComparison.OrdinalIgnoreCase) OrElse
                   host.EndsWith(".home", StringComparison.OrdinalIgnoreCase) Then
                    Return False
                End If

                Dim literalIp As IPAddress = Nothing
                If IPAddress.TryParse(host, literalIp) Then
                    Return Not IsPrivateWebImportIpAddress(literalIp)
                End If

                Try
                    For Each resolvedAddress As IPAddress In Dns.GetHostAddresses(uriResult.DnsSafeHost)
                        If IsPrivateWebImportIpAddress(resolvedAddress) Then
                            Return False
                        End If
                    Next
                Catch
                    ' If DNS resolution fails here, allow the normal request to determine reachability.
                End Try

                Return True
            Catch
                Return False
            End Try
        End Function

        Private Shared Function IsPrivateWebImportIpAddress(address As IPAddress) As Boolean
            If address Is Nothing Then
                Return True
            End If

            If IPAddress.IsLoopback(address) Then
                Return True
            End If

            Dim bytes As Byte() = address.GetAddressBytes()

            If address.AddressFamily = AddressFamily.InterNetwork Then
                If bytes.Length <> 4 Then
                    Return True
                End If

                If bytes(0) = 10 Then Return True
                If bytes(0) = 127 Then Return True
                If bytes(0) = 169 AndAlso bytes(1) = 254 Then Return True
                If bytes(0) = 172 AndAlso bytes(1) >= 16 AndAlso bytes(1) <= 31 Then Return True
                If bytes(0) = 192 AndAlso bytes(1) = 168 Then Return True
                If bytes(0) = 100 AndAlso bytes(1) >= 64 AndAlso bytes(1) <= 127 Then Return True
                If bytes(0) = 0 Then Return True

                Return False
            End If

            If address.AddressFamily = AddressFamily.InterNetworkV6 Then
                If address.IsIPv6LinkLocal OrElse address.IsIPv6SiteLocal Then
                    Return True
                End If

                If bytes.Length = 16 AndAlso (bytes(0) And &HFE) = &HFC Then
                    Return True
                End If

                Return False
            End If

            Return True
        End Function

    End Class
End Namespace
