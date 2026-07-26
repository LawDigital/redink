' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: RedInkPythonAgentWebGet.vb
' Purpose: Implements the host-agnostic `python_execute` `web.get` helper for
'          safe page retrieval and supported document extraction.
'
' Architecture / How it works:
'  - Validates that only absolute HTTP/HTTPS URLs are accepted and blocks known
'    authenticated cloud-storage hosts that cannot be fetched anonymously.
'  - Uses a lightweight HEAD request plus URL extension fallback to classify the
'    target as HTML/page content or as a supported downloadable document.
'  - Routes ordinary page retrieval through the host-provided HTML/WebView2
'    retriever so page extraction stays host-controlled and consistent.
'  - Downloads supported documents into a private temp folder, enforces size
'    caps, extracts text with sandbox-safe readers, prefixes compact metadata,
'    and deletes temp files on completion.
' =============================================================================
Option Explicit On
Option Infer On

Imports System.Net.Http
Imports System.Threading
Imports System.Threading.Tasks

Namespace Agents

    ''' <summary>
    ''' Host-agnostic implementation of the python_execute web.get operation.
    ''' - Ordinary web pages / HTML are routed to the host's WebView2 retriever (htmlPageRetriever).
    ''' - File-typed documents (PDF, DOCX, PPTX, XLSX, TXT) are downloaded into a host-owned temp
    '''   file, parsed with the sandboxed (non-Office-host) readers, prefixed with compact metadata,
    '''   and the temp file is deleted. The result is always a string, matching agent_api.web_get().
    ''' - ZIP/binary/unsupported/oversized/undecodable content fails with a typed host error and is
    '''   never decoded as text, base64-embedded, saved, or published.
    ''' </summary>
    Public NotInheritable Class RedInkPythonAgentWebGet

        Private Const MaximumDownloadBytes As System.Int64 = 33554432L ' 32 MiB policy cap for document extraction
        Private Shared ReadOnly BlockedHostMarkers As System.String() = {"sharepoint.com", "onedrive.com", "1drv.ms", "teams.microsoft.com", ":f:/", "/:f:/"}

        Private Sub New()
        End Sub

        Public Shared Async Function RetrieveAsync(url As System.String,
                                                   maximumCharacters As System.Int32,
                                                   htmlPageRetriever As System.Func(Of System.String, System.Int32, CancellationToken, Task(Of System.String)),
                                                   cancellationToken As CancellationToken) As Task(Of System.String)

            Dim uri As System.Uri = Nothing
            If System.String.IsNullOrWhiteSpace(url) OrElse
               Not System.Uri.TryCreate(url, System.UriKind.Absolute, uri) OrElse
               (uri.Scheme <> System.Uri.UriSchemeHttp AndAlso uri.Scheme <> System.Uri.UriSchemeHttps) Then
                Throw New RedInkPythonAgentHostCallException("WEB_URL_INVALID", False, "Only absolute HTTP/HTTPS URLs are allowed.")
            End If

            Dim lowerUrl As System.String = url.ToLowerInvariant()
            For Each marker As System.String In BlockedHostMarkers
                If lowerUrl.Contains(marker) Then
                    Throw New RedInkPythonAgentHostCallException("WEB_AUTHENTICATED_STORAGE_UNSUPPORTED", False, "Authenticated cloud storage URLs are not supported.")
                End If
            Next

            ' Classify by Content-Type via a lightweight HEAD; documents are downloaded and parsed,
            ' everything else (html/pages/unknown) is routed to the host WebView2 retriever.
            Dim documentKind As DocumentKind = Await ClassifyAsync(uri, cancellationToken).ConfigureAwait(False)

            If documentKind = DocumentKind.NotADocument Then
                If htmlPageRetriever Is Nothing Then
                    Throw New RedInkPythonAgentHostCallException("WEB_CONTENT_TYPE_UNSUPPORTED", False, "No page retriever is available for this content.")
                End If
                Dim pageText As System.String = Await htmlPageRetriever(url, maximumCharacters, cancellationToken).ConfigureAwait(False)
                Return Truncate(If(pageText, System.String.Empty), maximumCharacters)
            End If

            Return Await DownloadExtractAndFormatAsync(uri, documentKind, maximumCharacters, cancellationToken).ConfigureAwait(False)
        End Function

        Private Enum DocumentKind
            NotADocument
            Pdf
            Docx
            Xlsx
            Pptx
            Text
        End Enum

        Private Shared Async Function ClassifyAsync(uri As System.Uri, cancellationToken As CancellationToken) As Task(Of DocumentKind)
            Dim contentType As System.String = System.String.Empty
            Dim contentLength As System.Nullable(Of System.Int64) = Nothing
            Try
                Using client As New HttpClient()
                    client.Timeout = System.TimeSpan.FromSeconds(30)
                    Using request As New HttpRequestMessage(HttpMethod.Head, uri)
                        Using response As HttpResponseMessage = Await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(False)
                            If response.IsSuccessStatusCode Then
                                If response.Content IsNot Nothing AndAlso response.Content.Headers IsNot Nothing Then
                                    If response.Content.Headers.ContentType IsNot Nothing Then
                                        contentType = If(response.Content.Headers.ContentType.MediaType, System.String.Empty)
                                    End If
                                    contentLength = response.Content.Headers.ContentLength
                                End If
                            End If
                        End Using
                    End Using
                End Using
            Catch ex As OperationCanceledException
                Throw
            Catch
                ' HEAD is advisory only; fall back to treating this as a page.
                Return DocumentKind.NotADocument
            End Try

            If contentLength.HasValue AndAlso contentLength.Value > MaximumDownloadBytes Then
                Throw New RedInkPythonAgentHostCallException("WEB_RESPONSE_TOO_LARGE", False, "The remote resource exceeds the size limit.")
            End If

            Dim kind As DocumentKind = KindFromContentType(contentType)
            If kind = DocumentKind.NotADocument Then
                ' Octet-stream or missing type: use the URL extension as a secondary signal.
                kind = KindFromExtension(uri)
            End If
            Return kind
        End Function

        Private Shared Function KindFromContentType(mediaType As System.String) As DocumentKind
            Select Case If(mediaType, System.String.Empty).Trim().ToLowerInvariant()
                Case "application/pdf" : Return DocumentKind.Pdf
                Case "application/vnd.openxmlformats-officedocument.wordprocessingml.document" : Return DocumentKind.Docx
                Case "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" : Return DocumentKind.Xlsx
                Case "application/vnd.openxmlformats-officedocument.presentationml.presentation" : Return DocumentKind.Pptx
                Case "text/plain" : Return DocumentKind.Text
                Case Else : Return DocumentKind.NotADocument
            End Select
        End Function

        Private Shared Function KindFromExtension(uri As System.Uri) As DocumentKind
            Dim ext As System.String
            Try
                ext = System.IO.Path.GetExtension(uri.LocalPath).ToLowerInvariant()
            Catch
                Return DocumentKind.NotADocument
            End Try
            Select Case ext
                Case ".pdf" : Return DocumentKind.Pdf
                Case ".docx" : Return DocumentKind.Docx
                Case ".xlsx" : Return DocumentKind.Xlsx
                Case ".pptx" : Return DocumentKind.Pptx
                Case ".txt" : Return DocumentKind.Text
                Case Else : Return DocumentKind.NotADocument
            End Select
        End Function

        Private Shared Async Function DownloadExtractAndFormatAsync(uri As System.Uri, kind As DocumentKind, maximumCharacters As System.Int32, cancellationToken As CancellationToken) As Task(Of System.String)
            Dim tempDir As System.String = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "ri_webget_" & System.Guid.NewGuid().ToString("N"))
            System.IO.Directory.CreateDirectory(tempDir)
            Dim tempFile As System.String = System.IO.Path.Combine(tempDir, "download" & ExtensionFor(kind))
            Dim finalUrl As System.String = uri.AbsoluteUri
            Dim contentType As System.String = System.String.Empty
            Try
                Using client As New HttpClient()
                    client.Timeout = System.TimeSpan.FromSeconds(60)
                    Using response As HttpResponseMessage = Await client.GetAsync(uri, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(False)
                        If Not response.IsSuccessStatusCode Then
                            Throw New RedInkPythonAgentHostCallException("WEB_DOCUMENT_EXTRACTION_FAILED", False, "The remote resource returned an error status.")
                        End If
                        If response.RequestMessage IsNot Nothing AndAlso response.RequestMessage.RequestUri IsNot Nothing Then
                            finalUrl = response.RequestMessage.RequestUri.AbsoluteUri
                        End If
                        If response.Content IsNot Nothing AndAlso response.Content.Headers IsNot Nothing AndAlso response.Content.Headers.ContentType IsNot Nothing Then
                            contentType = If(response.Content.Headers.ContentType.MediaType, System.String.Empty)
                        End If
                        Await WriteLimitedAsync(response.Content, tempFile, cancellationToken).ConfigureAwait(False)
                    End Using
                End Using

                Dim extracted As System.String = Await ExtractAsync(kind, tempFile).ConfigureAwait(False)
                If System.String.IsNullOrWhiteSpace(extracted) OrElse (extracted.Length < 200 AndAlso extracted.TrimStart().StartsWith("Error", System.StringComparison.OrdinalIgnoreCase)) Then
                    Throw New RedInkPythonAgentHostCallException("WEB_DOCUMENT_EXTRACTION_FAILED", False, "The document could not be parsed.")
                End If

                Dim header As New System.Text.StringBuilder()
                header.AppendLine("Source URL: " & uri.AbsoluteUri)
                header.AppendLine("Final URL: " & finalUrl)
                header.AppendLine("Content-Type: " & If(contentType, System.String.Empty))
                header.AppendLine("Filename: " & FileNameFor(uri, kind))
                header.AppendLine()
                header.Append(extracted)
                Return Truncate(header.ToString(), maximumCharacters)

            Catch ex As OperationCanceledException
                Throw
            Catch ex As RedInkPythonAgentHostCallException
                Throw
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine(ex.ToString())
                Throw New RedInkPythonAgentHostCallException("WEB_DOCUMENT_EXTRACTION_FAILED", False, "The document could not be retrieved or parsed.")
            Finally
                Try
                    If System.IO.Directory.Exists(tempDir) Then
                        System.IO.Directory.Delete(tempDir, True)
                    End If
                Catch
                End Try
            End Try
        End Function

        Private Shared Async Function WriteLimitedAsync(content As HttpContent, destinationPath As System.String, cancellationToken As CancellationToken) As Task
            Using source As System.IO.Stream = Await content.ReadAsStreamAsync().ConfigureAwait(False)
                Using destination As New System.IO.FileStream(destinationPath, System.IO.FileMode.Create, System.IO.FileAccess.Write, System.IO.FileShare.None)
                    Dim buffer(8191) As System.Byte
                    Dim total As System.Int64 = 0
                    Do
                        cancellationToken.ThrowIfCancellationRequested()
                        Dim read As System.Int32 = Await source.ReadAsync(buffer, 0, buffer.Length, cancellationToken).ConfigureAwait(False)
                        If read <= 0 Then Exit Do
                        total += read
                        If total > MaximumDownloadBytes Then
                            Throw New RedInkPythonAgentHostCallException("WEB_RESPONSE_TOO_LARGE", False, "The remote resource exceeds the size limit.")
                        End If
                        destination.Write(buffer, 0, read)
                    Loop
                End Using
            End Using
        End Function

        Private Shared Async Function ExtractAsync(kind As DocumentKind, path As System.String) As Task(Of System.String)
            Select Case kind
                Case DocumentKind.Pdf
                    Return Await SharedLibrary.SharedMethods.ReadPdfAsText(path, ReturnErrorInsteadOfEmpty:=False, DoOCR:=False, AskUser:=False).ConfigureAwait(False)
                Case DocumentKind.Docx
                    Return SharedLibrary.SharedMethods.ReadDocxSandboxed(path)
                Case DocumentKind.Xlsx
                    Return SharedLibrary.SharedMethods.ReadXlsxSandboxed(path)
                Case DocumentKind.Pptx
                    Return SharedLibrary.SharedMethods.ReadPptxSandboxed(path)
                Case DocumentKind.Text
                    Return System.IO.File.ReadAllText(path, System.Text.Encoding.UTF8)
                Case Else
                    Throw New RedInkPythonAgentHostCallException("WEB_CONTENT_TYPE_UNSUPPORTED", False, "Unsupported content type.")
            End Select
        End Function

        Private Shared Function ExtensionFor(kind As DocumentKind) As System.String
            Select Case kind
                Case DocumentKind.Pdf : Return ".pdf"
                Case DocumentKind.Docx : Return ".docx"
                Case DocumentKind.Xlsx : Return ".xlsx"
                Case DocumentKind.Pptx : Return ".pptx"
                Case DocumentKind.Text : Return ".txt"
                Case Else : Return ".bin"
            End Select
        End Function

        Private Shared Function FileNameFor(uri As System.Uri, kind As DocumentKind) As System.String
            Dim name As System.String = System.String.Empty
            Try
                name = System.IO.Path.GetFileName(uri.LocalPath)
            Catch
            End Try
            If System.String.IsNullOrWhiteSpace(name) Then name = "download" & ExtensionFor(kind)
            Return name
        End Function

        Private Shared Function Truncate(value As System.String, maximumCharacters As System.Int32) As System.String
            If maximumCharacters > 0 AndAlso value.Length > maximumCharacters Then
                Return value.Substring(0, maximumCharacters)
            End If
            Return value
        End Function

    End Class

End Namespace
