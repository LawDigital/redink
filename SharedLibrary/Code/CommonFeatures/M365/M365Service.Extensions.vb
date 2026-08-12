' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: M365Service.Extensions.vb
' Purpose: Builds on M365Service core (auth/search/retrieval) with three
'          higher-level features:
'
'   1) Attachment expansion (ExpandAttachmentsAsync)
'      Recursively unpacks every attachment of an Outlook message:
'        - fileAttachment      → downloaded via /$value
'        - itemAttachment      → expanded with $expand=microsoft.graph.itemAttachment/item($expand=attachments)
'                                Nested fileAttachments are decoded from contentBytes
'                                (base64) and written to disk; the inner message
'                                is materialised as an M365Message
'        - referenceAttachment → sourceUrl recorded; if it is a Graph "shares"
'                                URL we attempt to follow it and download
'
'   2) Plain-text extraction (Get…AsTextAsync, GetHitAsTextAsync)
'      Produces an AI-friendly plain-text representation of any M365 entity,
'      reusing the existing SharedMethods extractors:
'        - Mail   → headers + body (HTML stripped) + recursively extracted attachments
'        - Files  → DriveItem downloaded → ReadDocxSandboxed / ReadXlsxSandboxed /
'                   ReadPptxSandboxed / ReadEmlSandboxed / ReadPdfAsText / ReadTextFile
'        - Events / Chat / OneNote / People → structured plain-text rendering
'      A single GetHitAsTextAsync(hit) dispatches by source so callers can write:
'           Dim docs = result.Hits.Select(Async Function(h) Await M365Service.GetHitAsTextAsync(_context, h))
'
'   3) Thread retrieval
'      - GetMailThreadAsync(conversationId)  → all messages in a mail thread
'      - GetChatThreadAsync(chatId)          → 1:1/group Teams chat history
'      - GetChannelThreadAsync(team,channel,rootMessageId)
'                                            → channel post + replies
' =============================================================================

Option Explicit On
Option Strict On

Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.Linq
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary

    ' ─── Public DTOs / options for the extensions ───────────────────────────

    ''' <summary>Type of an attachment after expansion.</summary>
    Public Enum M365ExpandedAttachmentKind
        FileAttachment
        ItemAttachment      ' embedded message / event / contact
        ReferenceAttachment ' link to OneDrive / SharePoint
        Unknown
    End Enum



    ''' <summary>One node in the (potentially nested) attachment tree of a message.</summary>
    Public Class M365ExpandedAttachment
        Public Property Kind As M365ExpandedAttachmentKind
        Public Property Name As String
        Public Property ContentType As String
        Public Property Size As Long
        Public Property IsInline As Boolean
        ''' <summary>Local path on disk (for file/reference attachments that downloaded successfully).</summary>
        Public Property LocalPath As String
        ''' <summary>For referenceAttachment: original cloud URL.</summary>
        Public Property ReferenceUrl As String
        ''' <summary>For itemAttachment of type message: the embedded message's parsed content.</summary>
        Public Property NestedMessage As M365Message
        ''' <summary>For itemAttachment of type event: minimal subject / start.</summary>
        Public Property NestedEventSubject As String
        ''' <summary>Recursively expanded children (only for itemAttachment when its inner message has its own attachments).</summary>
        Public Property Children As New List(Of M365ExpandedAttachment)()
        ''' <summary>Non-fatal error message captured while expanding this node.</summary>
        Public Property ErrorMessage As String
    End Class

    ''' <summary>Options for <see cref="M365Service.ExpandAttachmentsAsync"/>.</summary>
    Public Class M365ExpandOptions
        ''' <summary>Maximum recursion depth for itemAttachment trees (default 4).</summary>
        Public Property MaxDepth As Integer = 4
        ''' <summary>If True, inline attachments are also downloaded (defaults False — usually noise).</summary>
        Public Property IncludeInline As Boolean = False
        ''' <summary>If True, attempt to download referenceAttachments via Graph /shares.</summary>
        Public Property FollowReferenceAttachments As Boolean = True
        ''' <summary>Skip individual attachments larger than this size (bytes); 0 = no limit.</summary>
        Public Property MaxBytesPerAttachment As Long = 0
    End Class

    ''' <summary>Options for the plain-text extractors.</summary>
    Public Class M365TextOptions
        ''' <summary>Hard cap on the returned plain text length; 0 = unlimited.</summary>
        Public Property MaxChars As Integer = 0
        ''' <summary>If True, attachments of mail messages are extracted and appended.</summary>
        Public Property IncludeAttachments As Boolean = True
        ''' <summary>If True, inline attachments are also extracted.</summary>
        Public Property IncludeInlineAttachments As Boolean = False
        ''' <summary>If True, OCR heuristics are enabled when reading PDFs (uses LLM if configured).</summary>
        Public Property OcrPdf As Boolean = False
        ''' <summary>If True, prompt the user before running OCR on a PDF.</summary>
        Public Property AskUserForOcr As Boolean = False
        ''' <summary>Maximum recursion depth when expanding mail attachments (passed through to ExpandAttachmentsAsync).</summary>
        Public Property AttachmentMaxDepth As Integer = 3
        ''' <summary>If non-empty, downloaded files are kept here. Otherwise a temp folder is created and removed afterwards.</summary>
        Public Property WorkingFolder As String
        ''' <summary>If True (default), temp working folder is deleted at the end.</summary>
        Public Property CleanupWorkingFolder As Boolean = True
    End Class

    ''' <summary>Per-attachment plain text snippet captured during text extraction.</summary>
    Public Class M365AttachmentText
        Public Property Name As String
        Public Property LocalPath As String
        Public Property Text As String
        Public Property Errors As String
    End Class

    ''' <summary>Aggregate plain-text result.</summary>
    Public Class M365TextResult
        Public Property Source As M365SearchSources
        Public Property Id As String
        Public Property Title As String
        Public Property WebUrl As String
        ''' <summary>Concatenated AI-friendly plain text (header block + body + attachments).</summary>
        Public Property Text As String
        Public Property Truncated As Boolean
        Public Property AttachmentTexts As New List(Of M365AttachmentText)()
        Public Property Errors As New List(Of String)()
        ''' <summary>Folder containing any downloaded files. <c>Nothing</c> when cleanup ran.</summary>
        Public Property WorkingFolder As String
    End Class

    ''' <summary>Options for thread retrieval.</summary>
    Public Class M365ThreadOptions
        ''' <summary>Maximum total messages to return (default 200).</summary>
        Public Property MaxMessages As Integer = 200
        ''' <summary>If True, return oldest → newest. If False, newest → oldest (Graph default).</summary>
        Public Property Ascending As Boolean = True
        Public Property [From] As Date?
        Public Property [To] As Date?
        ''' <summary>For mail threads: include body in each message (default True).</summary>
        Public Property IncludeMailBody As Boolean = True
    End Class


    Partial Public Module M365Service

        ' ═══════════════════════════════════════════════════════════════════════
        '   1)  ATTACHMENT EXPANSION
        ' ═══════════════════════════════════════════════════════════════════════

        ''' <summary>
        ''' Recursively expands every attachment of a message:
        '''  - downloads fileAttachments to <paramref name="targetFolder"/>,
        '''  - parses itemAttachments (embedded messages / events) and recurses,
        '''  - records referenceAttachments and (optionally) downloads them.
        ''' Returns the full tree as <see cref="M365ExpandedAttachment"/>.
        ''' </summary>
        Public Async Function ExpandAttachmentsAsync(context As ISharedContext,
                                                     messageId As String,
                                                     targetFolder As String,
                                                     Optional options As M365ExpandOptions = Nothing,
                                                     Optional ct As CancellationToken = Nothing) As Task(Of List(Of M365ExpandedAttachment))
            If options Is Nothing Then options = New M365ExpandOptions()
            If String.IsNullOrWhiteSpace(targetFolder) Then
                Throw New ArgumentException("targetFolder is required", NameOf(targetFolder))
            End If

            Dim normalizedMessageId = ToGraphUrlSafeId(messageId)
            If String.IsNullOrWhiteSpace(normalizedMessageId) Then
                Throw New ArgumentException("messageId is required", NameOf(messageId))
            End If

            Directory.CreateDirectory(targetFolder)

            Dim token = Await GetAccessTokenAsync(context, ct).ConfigureAwait(False)
            Return Await ExpandAttachmentsRecursive(context, token, "/me/messages/" & Uri.EscapeDataString(normalizedMessageId),
                                                    targetFolder, options, depth:=0, ct:=ct).ConfigureAwait(False)
        End Function

        ''' <summary>
        ''' Internal worker. <paramref name="messagePathRel"/> is the Graph-relative path of the
        ''' "owning" message (e.g. <c>/me/messages/AAMk…</c>) so we can build URLs for both
        ''' top-level and nested attachment retrieval.
        ''' </summary>
        Private Async Function ExpandAttachmentsRecursive(context As ISharedContext,
                                                          token As String,
                                                          messagePathRel As String,
                                                          targetFolder As String,
                                                          options As M365ExpandOptions,
                                                          depth As Integer,
                                                          ct As CancellationToken) As Task(Of List(Of M365ExpandedAttachment))
            Dim out As New List(Of M365ExpandedAttachment)()

            ' Fetch attachments of this message; expand nested itemAttachment/item (with its own attachments)
            ' so we can decode contentBytes inline without an extra round-trip per attachment.
            Dim url = $"{GraphV1}{messagePathRel}/attachments?$expand=microsoft.graph.itemAttachment/item($expand=attachments)"
            Dim resp As JObject = Nothing
            Dim primaryError As Exception = Nothing
            Try
                resp = Await GraphGetAsync(token, url, ct).ConfigureAwait(False)
            Catch ex As Exception
                primaryError = ex
            End Try

            ' Some servers/permissions reject the $expand on nested items; retry without it.
            If resp Is Nothing Then
                Try
                    resp = Await GraphGetAsync(token, $"{GraphV1}{messagePathRel}/attachments", ct).ConfigureAwait(False)
                Catch ex2 As Exception
                    out.Add(New M365ExpandedAttachment() With {.Kind = M365ExpandedAttachmentKind.Unknown,
                                                                .ErrorMessage = ex2.Message})
                    Return out
                End Try
            End If

            Dim arr = TryCast(resp("value"), JArray)
            If arr Is Nothing Then Return out

            For Each at In arr
                ct.ThrowIfCancellationRequested()
                Dim ao = TryCast(at, JObject)
                If ao Is Nothing Then Continue For

                Dim node = New M365ExpandedAttachment() With {
                    .Id = SafeStr(ao, "id"),
                    .Name = SafeStr(ao, "name"),
                    .ContentType = SafeStr(ao, "contentType"),
                    .Size = CLng(If(ao("size"), 0L)),
                    .IsInline = CBool(If(ao("isInline"), False))
                }

                If node.IsInline AndAlso Not options.IncludeInline Then Continue For
                If options.MaxBytesPerAttachment > 0 AndAlso node.Size > options.MaxBytesPerAttachment Then
                    node.ErrorMessage = $"Skipped: size {node.Size} exceeds limit {options.MaxBytesPerAttachment}"
                    out.Add(node)
                    Continue For
                End If

                Dim odataType = SafeStr(ao, "@odata.type").ToLowerInvariant()
                Try
                    If odataType.Contains("fileattachment") Then
                        node.Kind = M365ExpandedAttachmentKind.FileAttachment
                        Dim attId = SafeStr(ao, "id")
                        Dim safeName = MakeSafeFileName(node.Name, "attachment.bin")
                        Dim localPath = UniquePath(System.IO.Path.Combine(targetFolder, safeName))

                        ' Prefer inline contentBytes (avoids an extra HTTP call) when present
                        Dim contentBytes = SafeStr(ao, "contentBytes")
                        If Not String.IsNullOrEmpty(contentBytes) Then
                            File.WriteAllBytes(localPath, System.Convert.FromBase64String(contentBytes))
                        Else
                            Dim relAtt = $"{messagePathRel}/attachments/{Uri.EscapeDataString(attId)}/$value"
                            Await DownloadRawAsync(token, GraphV1 & relAtt, localPath, ct).ConfigureAwait(False)
                        End If
                        node.LocalPath = localPath

                    ElseIf odataType.Contains("itemattachment") Then
                        node.Kind = M365ExpandedAttachmentKind.ItemAttachment
                        Dim inner = TryCast(ao("item"), JObject)
                        If inner IsNot Nothing Then
                            Dim innerType = SafeStr(inner, "@odata.type").ToLowerInvariant()
                            If innerType.Contains("message") Then
                                ' Build a M365Message from the inline expansion.
                                node.NestedMessage = ParseMessage(inner,
                                    M365MessageFields.Body Or M365MessageFields.Recipients Or M365MessageFields.AttachmentsList)

                                ' Also bring in nested attachments that the inline expand returned.
                                Dim nestedAtts = TryCast(inner("attachments"), JArray)
                                If nestedAtts IsNot Nothing AndAlso nestedAtts.Count > 0 AndAlso depth + 1 < options.MaxDepth Then
                                    For Each na In nestedAtts
                                        ct.ThrowIfCancellationRequested()
                                        Dim child = Await DecodeInlineAttachmentAsync(token, CType(na, JObject),
                                                                                      targetFolder, options, depth + 1, ct).ConfigureAwait(False)
                                        If child IsNot Nothing Then node.Children.Add(child)
                                    Next
                                End If

                            ElseIf innerType.Contains("event") Then
                                node.NestedEventSubject = SafeStr(inner, "subject")
                            End If
                        End If

                    ElseIf odataType.Contains("referenceattachment") Then
                        node.Kind = M365ExpandedAttachmentKind.ReferenceAttachment
                        Dim refUrl = SafeStr(ao, "sourceUrl")
                        node.ReferenceUrl = refUrl
                        If options.FollowReferenceAttachments AndAlso Not String.IsNullOrEmpty(refUrl) Then
                            Try
                                Dim localPath = UniquePath(System.IO.Path.Combine(targetFolder, MakeSafeFileName(node.Name, "reference.bin")))
                                If Await TryDownloadSharedUrlAsync(token, refUrl, localPath, ct).ConfigureAwait(False) Then
                                    node.LocalPath = localPath
                                End If
                            Catch ex As Exception
                                node.ErrorMessage = "ReferenceAttachment download failed: " & ex.Message
                            End Try
                        End If

                    Else
                        node.Kind = M365ExpandedAttachmentKind.Unknown
                    End If

                Catch ex As Exception
                    node.ErrorMessage = ex.Message
                End Try

                out.Add(node)
            Next

            Return out
        End Function

        ''' <summary>
        ''' Decodes an attachment that was returned inline (via <c>$expand=…/item($expand=attachments)</c>).
        ''' For nested fileAttachments we just persist the base64 bytes; for nested itemAttachments we
        ''' record the inner message metadata but do not recurse further unless depth budget remains.
        ''' </summary>
        Private Async Function DecodeInlineAttachmentAsync(token As String,
                                                           ao As JObject,
                                                           targetFolder As String,
                                                           options As M365ExpandOptions,
                                                           depth As Integer,
                                                           ct As CancellationToken) As Task(Of M365ExpandedAttachment)
            If ao Is Nothing Then Return Nothing

            Dim node As New M365ExpandedAttachment() With {
                .Id = SafeStr(ao, "id"),
                .Name = SafeStr(ao, "name"),
                .ContentType = SafeStr(ao, "contentType"),
                .Size = CLng(If(ao("size"), 0L)),
                .IsInline = CBool(If(ao("isInline"), False))
            }
            If node.IsInline AndAlso Not options.IncludeInline Then Return Nothing
            If options.MaxBytesPerAttachment > 0 AndAlso node.Size > options.MaxBytesPerAttachment Then
                node.ErrorMessage = $"Skipped: size {node.Size} exceeds limit {options.MaxBytesPerAttachment}"
                Return node
            End If

            Dim odataType = SafeStr(ao, "@odata.type").ToLowerInvariant()
            Try
                If odataType.Contains("fileattachment") Then
                    node.Kind = M365ExpandedAttachmentKind.FileAttachment
                    Dim cb = SafeStr(ao, "contentBytes")
                    If Not String.IsNullOrEmpty(cb) Then
                        Dim localPath = UniquePath(System.IO.Path.Combine(targetFolder, MakeSafeFileName(node.Name, "nested.bin")))
                        File.WriteAllBytes(localPath, System.Convert.FromBase64String(cb))
                        node.LocalPath = localPath
                    Else
                        node.ErrorMessage = "Nested fileAttachment had no contentBytes (deeper retrieval needed)."
                    End If
                ElseIf odataType.Contains("itemattachment") Then
                    node.Kind = M365ExpandedAttachmentKind.ItemAttachment
                    Dim inner = TryCast(ao("item"), JObject)
                    If inner IsNot Nothing Then
                        Dim innerType = SafeStr(inner, "@odata.type").ToLowerInvariant()
                        If innerType.Contains("message") Then
                            node.NestedMessage = ParseMessage(inner, M365MessageFields.Body Or M365MessageFields.Recipients)
                        ElseIf innerType.Contains("event") Then
                            node.NestedEventSubject = SafeStr(inner, "subject")
                        End If
                    End If
                ElseIf odataType.Contains("referenceattachment") Then
                    node.Kind = M365ExpandedAttachmentKind.ReferenceAttachment
                    node.ReferenceUrl = SafeStr(ao, "sourceUrl")
                Else
                    node.Kind = M365ExpandedAttachmentKind.Unknown
                End If
            Catch ex As Exception
                node.ErrorMessage = ex.Message
            End Try

            Return node
        End Function

        ''' <summary>
        ''' Streams a Graph URL's bytes to disk.
        ''' </summary>
        Private Async Function DownloadRawAsync(token As String, url As String, targetPath As String,
                                                ct As CancellationToken) As Task
            Using req As New HttpRequestMessage(HttpMethod.Get, url)
                req.Headers.Authorization = New AuthenticationHeaderValue("Bearer", token)
                Using resp = Await _http.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct).ConfigureAwait(False)
                    If Not resp.IsSuccessStatusCode Then
                        Dim body = ""
                        Try : body = Await resp.Content.ReadAsStringAsync().ConfigureAwait(False) : Catch : End Try
                        Throw BuildGraphException(resp.StatusCode, body)
                    End If
                    Directory.CreateDirectory(System.IO.Path.GetDirectoryName(targetPath))
                    Using fs As New FileStream(targetPath, FileMode.Create, FileAccess.Write, FileShare.None)
                        Await resp.Content.CopyToAsync(fs).ConfigureAwait(False)
                    End Using
                End Using
            End Using
        End Function

        ''' <summary>
        ''' Best-effort download of a referenceAttachment cloud URL via Graph /shares.
        ''' Returns True if the file was written to <paramref name="targetPath"/>.
        ''' </summary>
        Private Async Function TryDownloadSharedUrlAsync(token As String, refUrl As String,
                                                        targetPath As String,
                                                        ct As CancellationToken) As Task(Of Boolean)
            Try
                ' Graph encodes a sharing URL as: u! + base64url-without-padding.
                Dim encoded = "u!" & System.Convert.ToBase64String(Encoding.UTF8.GetBytes(refUrl)) _
                                          .TrimEnd("="c).Replace("/"c, "_"c).Replace("+"c, "-"c)
                Dim metaUrl = $"{GraphV1}/shares/{encoded}/driveItem"
                Dim meta = Await GraphGetAsync(token, metaUrl, ct).ConfigureAwait(False)
                Dim downloadUrl = SafeStr(meta, "@microsoft.graph.downloadUrl")
                If String.IsNullOrEmpty(downloadUrl) Then Return False
                Using req As New HttpRequestMessage(HttpMethod.Get, downloadUrl)
                    Using resp = Await _http.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct).ConfigureAwait(False)
                        If Not resp.IsSuccessStatusCode Then Return False
                        Directory.CreateDirectory(Path.GetDirectoryName(targetPath))
                        Using fs As New FileStream(targetPath, FileMode.Create, FileAccess.Write, FileShare.None)
                            Await resp.Content.CopyToAsync(fs).ConfigureAwait(False)
                        End Using
                    End Using
                End Using
                Return True
            Catch ex As Exception
                Debug.WriteLine($"[M365] TryDownloadSharedUrl failed: {ex.Message}")
                Return False
            End Try
        End Function

        Private Function UniquePath(suggested As String) As String
            If Not File.Exists(suggested) Then Return suggested
            Dim dir = Path.GetDirectoryName(suggested)
            Dim baseName = Path.GetFileNameWithoutExtension(suggested)
            Dim ext = Path.GetExtension(suggested)
            Dim n As Integer = 1
            Dim p = suggested
            While File.Exists(p)
                p = Path.Combine(dir, baseName & $"_{n}" & ext)
                n += 1
            End While
            Return p
        End Function

        ' ═══════════════════════════════════════════════════════════════════════
        '   2)  PLAIN-TEXT EXTRACTION
        ' ═══════════════════════════════════════════════════════════════════════

        ''' <summary>
        ''' Convenience dispatcher: turns any <see cref="M365SearchHit"/> into AI-friendly plain text.
        ''' Calls the appropriate per-source extractor.
        ''' </summary>
        Public Async Function GetHitAsTextAsync(context As ISharedContext,
                                                hit As M365SearchHit,
                                                Optional options As M365TextOptions = Nothing,
                                                Optional ct As CancellationToken = Nothing) As Task(Of M365TextResult)
            If hit Is Nothing Then Throw New ArgumentNullException(NameOf(hit))
            If options Is Nothing Then options = New M365TextOptions()

            Select Case hit.Source
                Case M365SearchSources.Mail




                    Return Await GetMessageAsTextAsync(context, hit.Id, options, ct).ConfigureAwait(False)
                Case M365SearchSources.OneDrive, M365SearchSources.SharePoint
                    Return Await GetDriveItemAsTextAsync(
                        context,
                        driveItemId:=hit.Id,
                        driveId:=hit.ParentId,
                        webUrl:=hit.WebUrl,
                        options:=options,
                        ct:=ct).ConfigureAwait(False)
                Case M365SearchSources.Calendar
                    Return Await GetEventAsTextAsync(context, hit.Id, ct).ConfigureAwait(False)
                Case M365SearchSources.Teams
                    Return Await GetChatMessageAsTextAsync(context, hit.Id, chatId:=hit.ParentId, ct:=ct).ConfigureAwait(False)
                Case M365SearchSources.OneNote
                    Return Await GetOneNotePageAsTextAsync(context, hit.Id, ct).ConfigureAwait(False)
                Case M365SearchSources.People
                    Return New M365TextResult() With {
                        .Source = hit.Source, .Id = hit.Id, .Title = hit.Title,
                        .Text = $"{hit.Title}{vbCrLf}{hit.Author}{vbCrLf}{hit.Summary}".Trim()
                    }
                Case Else
                    Return New M365TextResult() With {
                        .Source = hit.Source, .Id = hit.Id, .Title = hit.Title,
                        .Text = hit.Summary
                    }
            End Select
        End Function

        Private Async Function GetDriveItemByWebUrlAsync(context As ISharedContext,
                                                         webUrl As String,
                                                         Optional ct As CancellationToken = Nothing) As Task(Of M365DriveItem)
            If String.IsNullOrWhiteSpace(webUrl) Then Return Nothing

            Dim token = Await GetAccessTokenAsync(context, ct).ConfigureAwait(False)
            Dim encoded = "u!" & System.Convert.ToBase64String(Encoding.UTF8.GetBytes(webUrl.Trim())) _
                .TrimEnd("="c).Replace("/"c, "_"c).Replace("+"c, "-"c)

            Dim url = $"{GraphV1}/shares/{encoded}/driveItem"
            Dim j = Await GraphGetAsync(token, url, ct).ConfigureAwait(False)
            Return ParseDriveItem(j)
        End Function

        ''' <summary>
        ''' Extracts a mail message as plain text: header block + body (HTML stripped) +
        ''' recursively extracted attachments.
        ''' </summary>
        Public Async Function GetMessageAsTextAsync(context As ISharedContext,
                                                    messageId As String,
                                                    Optional options As M365TextOptions = Nothing,
                                                    Optional ct As CancellationToken = Nothing) As Task(Of M365TextResult)
            If options Is Nothing Then options = New M365TextOptions()
            Dim r As New M365TextResult() With {.Source = M365SearchSources.Mail, .Id = messageId}

            Dim msg = Await GetMessageAsync(context, messageId,
                M365MessageFields.Body Or M365MessageFields.Recipients Or M365MessageFields.AttachmentsList,
                ct).ConfigureAwait(False)
            If msg Is Nothing Then
                r.Errors.Add("Message not found.")
                Return r
            End If
            r.Title = msg.Subject
            r.WebUrl = If(msg.WebLink, "")

            Dim sb As New StringBuilder()
            sb.AppendLine("From: " & FormatAddr(msg.From, msg.FromAddress))
            If msg.To_.Count > 0 Then sb.AppendLine("To: " & String.Join("; ", msg.To_))
            If msg.Cc.Count > 0 Then sb.AppendLine("Cc: " & String.Join("; ", msg.Cc))
            AppendDateLines(sb, "Sent", msg.SentUtc)
            AppendDateLines(sb, "Received", msg.ReceivedUtc)
            sb.AppendLine("Subject: " & If(msg.Subject, ""))
            sb.AppendLine()
            Dim body = If(msg.Body, "")
            If String.Equals(msg.BodyContentType, "html", StringComparison.OrdinalIgnoreCase) Then body = StripHtml(body)
            sb.AppendLine(body.Trim())

            If options.IncludeAttachments AndAlso msg.HasAttachments Then
                Dim work = EnsureWorkingFolder(options, "ri_m365_msg_")
                r.WorkingFolder = work
                Try
                    Dim expandOpts As New M365ExpandOptions() With {
                        .MaxDepth = options.AttachmentMaxDepth,
                        .IncludeInline = options.IncludeInlineAttachments,
                        .FollowReferenceAttachments = True
                    }
                    Dim atts = Await ExpandAttachmentsAsync(context, messageId, work, expandOpts, ct).ConfigureAwait(False)
                    For Each a In atts
                        ct.ThrowIfCancellationRequested()
                        Await AppendAttachmentTextAsync(context, a, sb, r, options, ct).ConfigureAwait(False)
                    Next
                Finally
                    If options.CleanupWorkingFolder AndAlso String.IsNullOrEmpty(options.WorkingFolder) Then
                        Try : Directory.Delete(work, True) : r.WorkingFolder = Nothing : Catch : End Try
                    End If
                End Try
            End If

            r.Text = TruncateIfNeeded(sb.ToString().Trim(), options, r)
            Return r
        End Function

        Private Sub AppendDateLines(sb As StringBuilder, label As String, value As DateTime?)
            If sb Is Nothing OrElse String.IsNullOrWhiteSpace(label) Then Return

            If Not value.HasValue Then
                sb.AppendLine(label & ": ")
                Return
            End If

            Dim utc = value.Value.ToUniversalTime()
            Dim localValue As DateTimeOffset = New DateTimeOffset(utc, TimeSpan.Zero).ToLocalTime()

            Debug.WriteLine(
                $"[M365Date] {label}: rawUtc='{If(value.HasValue, value.Value.ToString("o"), "")}' " &
                $"local='{localValue.ToString("dd MMM yyyy HH:mm", Globalization.CultureInfo.InvariantCulture)} local (UTC{localValue.ToString("zzz", Globalization.CultureInfo.InvariantCulture)})' " &
                $"utc='{utc.ToString("dd MMM yyyy HH:mm 'UTC'", Globalization.CultureInfo.InvariantCulture)}'")

            sb.AppendLine(
                label & ": " &
                localValue.ToString("dd MMM yyyy HH:mm", Globalization.CultureInfo.InvariantCulture) &
                " local (UTC" &
                localValue.ToString("zzz", Globalization.CultureInfo.InvariantCulture) &
                ")")
            sb.AppendLine(label & "Utc: " & utc.ToString("dd MMM yyyy HH:mm 'UTC'", Globalization.CultureInfo.InvariantCulture))
        End Sub

        ''' <summary>
        ''' Recursively renders an expanded attachment node into <paramref name="sb"/> as plain text,
        ''' invoking the appropriate file extractor.
        ''' </summary>
        Private Async Function AppendAttachmentTextAsync(context As ISharedContext,
                                                 a As M365ExpandedAttachment,
                                                 sb As StringBuilder,
                                                 r As M365TextResult,
                                                 options As M365TextOptions,
                                                 ct As CancellationToken) As Task
            sb.AppendLine()
            sb.AppendLine($"═══ Attachment: {If(a.Name, "(unnamed)")} ({a.Kind}) ═══")
            If Not String.IsNullOrEmpty(a.ErrorMessage) Then
                sb.AppendLine($"[Error: {a.ErrorMessage}]")
                r.Errors.Add($"{a.Name}: {a.ErrorMessage}")
            End If

            Select Case a.Kind
                Case M365ExpandedAttachmentKind.FileAttachment, M365ExpandedAttachmentKind.ReferenceAttachment
                    If Not String.IsNullOrEmpty(a.LocalPath) Then
                        Dim text = Await ExtractTextFromLocalFileAsync(context, a.LocalPath, options, ct).ConfigureAwait(False)
                        sb.AppendLine(text.Text)
                        r.AttachmentTexts.Add(New M365AttachmentText() With {
                    .Name = a.Name, .LocalPath = a.LocalPath, .Text = text.Text, .Errors = text.ErrorMessage})
                        If Not String.IsNullOrEmpty(text.ErrorMessage) Then r.Errors.Add($"{a.Name}: {text.ErrorMessage}")
                    ElseIf a.Kind = M365ExpandedAttachmentKind.ReferenceAttachment Then
                        sb.AppendLine($"[Reference URL: {a.ReferenceUrl}]")
                    End If

                Case M365ExpandedAttachmentKind.ItemAttachment
                    If a.NestedMessage IsNot Nothing Then
                        sb.AppendLine("From: " & FormatAddr(a.NestedMessage.From, a.NestedMessage.FromAddress))
                        If a.NestedMessage.To_.Count > 0 Then sb.AppendLine("To: " & String.Join("; ", a.NestedMessage.To_))
                        If a.NestedMessage.Cc.Count > 0 Then sb.AppendLine("Cc: " & String.Join("; ", a.NestedMessage.Cc))
                        AppendDateLines(sb, "Sent", a.NestedMessage.SentUtc)
                        AppendDateLines(sb, "Received", a.NestedMessage.ReceivedUtc)
                        sb.AppendLine("Subject: " & If(a.NestedMessage.Subject, ""))
                        sb.AppendLine()

                        Dim body = If(a.NestedMessage.Body, "")
                        If String.Equals(a.NestedMessage.BodyContentType, "html", StringComparison.OrdinalIgnoreCase) Then body = StripHtml(body)
                        sb.AppendLine(body.Trim())

                        For Each child In a.Children
                            Await AppendAttachmentTextAsync(context, child, sb, r, options, ct).ConfigureAwait(False)
                        Next

                    ElseIf a.NestedEvent IsNot Nothing Then
                        Dim ev = a.NestedEvent
                        sb.AppendLine("Event: " & If(ev.Subject, ""))
                        If Not String.IsNullOrEmpty(ev.Organizer) Then sb.AppendLine("Organizer: " & ev.Organizer)
                        If Not String.IsNullOrEmpty(ev.Location) Then sb.AppendLine("Location: " & ev.Location)
                        AppendDateLines(sb, "Start", ev.StartUtc)
                        AppendDateLines(sb, "End", ev.EndUtc)
                        If ev.Attendees.Count > 0 Then sb.AppendLine("Attendees: " & String.Join("; ", ev.Attendees))
                        sb.AppendLine()
                        sb.AppendLine(If(ev.BodyPreview, "").Trim())

                    ElseIf Not String.IsNullOrEmpty(a.NestedEventSubject) Then
                        sb.AppendLine($"[Embedded event: {a.NestedEventSubject}]")
                    Else
                        sb.AppendLine("[Embedded item — no expandable content]")
                    End If

                Case Else
                    sb.AppendLine("[Unknown attachment type]")
            End Select
        End Function



        ''' <summary>
        ''' Downloads a DriveItem (OneDrive / SharePoint) to a temp file and returns its plain text.
        ''' </summary>
        Public Async Function GetDriveItemAsTextAsync(context As ISharedContext,
                                                      driveItemId As String,
                                                      Optional driveId As String = Nothing,
                                                      Optional webUrl As String = Nothing,
                                                      Optional options As M365TextOptions = Nothing,
                                                      Optional ct As CancellationToken = Nothing) As Task(Of M365TextResult)
            If options Is Nothing Then options = New M365TextOptions()
            Dim r As New M365TextResult() With {.Source = M365SearchSources.OneDrive, .Id = driveItemId}

            Dim meta As M365DriveItem = Nothing
            Dim lookupErrors As New List(Of String)()

            If Not String.IsNullOrWhiteSpace(driveId) Then
                Try
                    meta = Await GetDriveItemAsync(context, driveItemId, driveId, ct).ConfigureAwait(False)
                Catch ex As Exception
                    lookupErrors.Add(ex.Message)
                End Try
            ElseIf Not String.IsNullOrWhiteSpace(webUrl) Then
                Try
                    meta = Await GetDriveItemByWebUrlAsync(context, webUrl, ct).ConfigureAwait(False)
                Catch ex As Exception
                    lookupErrors.Add("/shares metadata lookup failed: " & ex.Message)
                End Try

                If meta Is Nothing Then
                    Try
                        meta = Await GetDriveItemAsync(context, driveItemId, Nothing, ct).ConfigureAwait(False)
                    Catch ex As Exception
                        lookupErrors.Add("/me/drive metadata lookup failed: " & ex.Message)
                    End Try
                End If
            Else
                Try
                    meta = Await GetDriveItemAsync(context, driveItemId, Nothing, ct).ConfigureAwait(False)
                Catch ex As Exception
                    lookupErrors.Add(ex.Message)
                End Try
            End If

            If meta Is Nothing Then
                r.WebUrl = If(webUrl, "")
                If lookupErrors.Count = 0 Then
                    r.Errors.Add("DriveItem not found.")
                Else
                    r.Errors.Add("DriveItem not found. " & String.Join(" | ", lookupErrors.Distinct()))
                End If
                Return r
            End If

            If String.IsNullOrWhiteSpace(driveId) Then driveId = If(meta.DriveId, "")
            If String.IsNullOrWhiteSpace(webUrl) Then webUrl = If(meta.WebUrl, "")
            If Not String.IsNullOrWhiteSpace(webUrl) AndAlso
               webUrl.IndexOf("sharepoint", StringComparison.OrdinalIgnoreCase) >= 0 Then
                r.Source = M365SearchSources.SharePoint
            End If

            r.Id = If(meta.Id, driveItemId)
            r.Title = meta.Name
            r.WebUrl = If(meta.WebUrl, webUrl)

            If meta.IsFolder Then
                r.Text = "[Folder] " & meta.Name
                Return r
            End If

            Dim work = EnsureWorkingFolder(options, "ri_m365_file_")
            r.WorkingFolder = work
            Dim path = UniquePath(System.IO.Path.Combine(work, MakeSafeFileName(meta.Name, "file.bin")))

            Try
                If Not String.IsNullOrWhiteSpace(driveId) Then
                    Await DownloadFileAsync(context, meta.Id, path, driveId, ct).ConfigureAwait(False)
                ElseIf Not String.IsNullOrWhiteSpace(webUrl) Then
                    Dim token = Await GetAccessTokenAsync(context, ct).ConfigureAwait(False)
                    Dim ok = Await TryDownloadSharedUrlAsync(token, webUrl, path, ct).ConfigureAwait(False)
                    If Not ok Then
                        Await DownloadFileAsync(context, meta.Id, path, Nothing, ct).ConfigureAwait(False)
                    End If
                Else
                    Await DownloadFileAsync(context, meta.Id, path, Nothing, ct).ConfigureAwait(False)
                End If

                Dim ex = Await ExtractTextFromLocalFileAsync(context, path, options, ct).ConfigureAwait(False)
                Dim sb As New StringBuilder()
                sb.AppendLine($"File: {meta.Name}")
                sb.AppendLine($"Path: {meta.Path}")
                sb.AppendLine($"LastModified: {If(meta.LastModifiedUtc.HasValue, meta.LastModifiedUtc.Value.ToString("u"), "")}")
                sb.AppendLine($"By: {meta.LastModifiedBy}")
                sb.AppendLine()
                sb.AppendLine(ex.Text)
                r.Text = TruncateIfNeeded(sb.ToString().Trim(), options, r)
                If Not String.IsNullOrEmpty(ex.ErrorMessage) Then r.Errors.Add(ex.ErrorMessage)
            Catch exDl As Exception
                r.Errors.Add("Download/extract failed: " & exDl.Message)
            Finally
                If options.CleanupWorkingFolder AndAlso String.IsNullOrEmpty(options.WorkingFolder) Then
                    Try : Directory.Delete(work, True) : r.WorkingFolder = Nothing : Catch : End Try
                End If
            End Try

            Return r
        End Function

        ''' <summary>
        ''' Plain-text rendering of an Outlook calendar event.
        ''' </summary>
        Public Async Function GetEventAsTextAsync(context As ISharedContext,
                                          eventId As String,
                                          Optional ct As CancellationToken = Nothing) As Task(Of M365TextResult)
            Dim r As New M365TextResult() With {.Source = M365SearchSources.Calendar, .Id = eventId}
            Dim ev = Await GetEventAsync(context, eventId, ct).ConfigureAwait(False)
            If ev Is Nothing Then
                r.Errors.Add("Event not found.")
                Return r
            End If

            r.Title = ev.Subject
            r.WebUrl = If(ev.WebLink, "")

            Dim sb As New StringBuilder()
            sb.AppendLine("Subject: " & If(ev.Subject, ""))
            sb.AppendLine("Organizer: " & If(ev.Organizer, ""))
            sb.AppendLine("Location: " & If(ev.Location, ""))
            AppendDateLines(sb, "Start", ev.StartUtc)
            AppendDateLines(sb, "End", ev.EndUtc)
            If ev.Attendees.Count > 0 Then sb.AppendLine("Attendees: " & String.Join("; ", ev.Attendees))
            sb.AppendLine()
            sb.AppendLine(If(ev.BodyPreview, ""))

            r.Text = sb.ToString().Trim()
            Return r
        End Function


        ''' <summary>Plain-text rendering of a single Teams chat message.</summary>
        Public Async Function GetChatMessageAsTextAsync(context As ISharedContext,
                                                messageId As String,
                                                Optional chatId As String = Nothing,
                                                Optional teamId As String = Nothing,
                                                Optional channelId As String = Nothing,
                                                Optional ct As CancellationToken = Nothing) As Task(Of M365TextResult)
            Dim r As New M365TextResult() With {.Source = M365SearchSources.Teams, .Id = messageId}
            Dim cm = Await GetChatMessageAsync(context, messageId, chatId, teamId, channelId, ct).ConfigureAwait(False)
            If cm Is Nothing Then
                r.Errors.Add("Chat message not found.")
                Return r
            End If

            r.Title = TruncatePlain(StripHtml(If(cm.Content, "")), 80)

            Dim sb As New StringBuilder()
            sb.AppendLine("From: " & If(cm.From, ""))
            AppendDateLines(sb, "At", cm.CreatedUtc)
            sb.AppendLine()

            Dim body = If(cm.Content, "")
            If String.Equals(cm.ContentType, "html", StringComparison.OrdinalIgnoreCase) Then body = StripHtml(body)
            sb.AppendLine(body.Trim())

            r.Text = sb.ToString().Trim()
            Return r
        End Function

        ''' <summary>Plain-text rendering of a OneNote page (HTML stripped).</summary>
        Public Async Function GetOneNotePageAsTextAsync(context As ISharedContext,
                                                        pageId As String,
                                                        Optional ct As CancellationToken = Nothing) As Task(Of M365TextResult)
            Dim r As New M365TextResult() With {.Source = M365SearchSources.OneNote, .Id = pageId}
            Dim p = Await GetOneNotePageAsync(context, pageId, includeContent:=True, ct:=ct).ConfigureAwait(False)
            If p Is Nothing Then
                r.Errors.Add("OneNote page not found.")
                Return r
            End If
            r.Title = p.Title
            Dim sb As New StringBuilder()
            sb.AppendLine($"Title: {p.Title}")
            sb.AppendLine($"LastModified: {If(p.LastModifiedUtc.HasValue, p.LastModifiedUtc.Value.ToString("u"), "")}")
            sb.AppendLine()
            sb.AppendLine(StripHtml(If(p.HtmlContent, "")))
            r.Text = sb.ToString().Trim()
            Return r
        End Function

        ''' <summary>
        ''' Dispatches a downloaded file to the right SharedMethods extractor based on extension.
        ''' Mirrors the dispatch in SharedMethods.SandboxedReaders so that all current/future
        ''' formats are supported uniformly.
        ''' </summary>
        Private Async Function ExtractTextFromLocalFileAsync(context As ISharedContext,
                                                             path As String,
                                                             options As M365TextOptions,
                                                             ct As CancellationToken) As Task(Of (Text As String, ErrorMessage As String))
            Try
                If String.IsNullOrEmpty(path) OrElse Not File.Exists(path) Then
                    Return ("", "File not found: " & path)
                End If
                Dim ext = System.IO.Path.GetExtension(path).ToLowerInvariant()
                Dim text As String = Nothing
                Select Case ext
                    Case ".docx"
                        text = SharedMethods.ReadDocxSandboxed(path)
                    Case ".xlsx"
                        text = SharedMethods.ReadXlsxSandboxed(path)
                    Case ".pptx"
                        text = SharedMethods.ReadPptxSandboxed(path)
                    Case ".eml"
                        text = SharedMethods.ReadEmlSandboxed(path)
                    Case ".pdf"
                        text = Await SharedMethods.ReadPdfAsText(path,
                                                                  ReturnErrorInsteadOfEmpty:=False,
                                                                  DoOCR:=options.OcrPdf,
                                                                  AskUser:=options.AskUserForOcr,
                                                                  context:=context).ConfigureAwait(False)
                    Case ".rtf"
                        text = SharedMethods.ReadRtfAsText(path, False)
                    Case ".doc"
                        If SharedMethods.INI_AllowLegacyDocFiles_Cached Then
                            text = SharedMethods.ReadWordDocument(path, False)
                        Else
                            text = "[Skipped: .doc format disabled for security]"
                        End If
                    Case ".txt", ".csv", ".log", ".json", ".xml", ".html", ".htm",
                         ".md", ".yaml", ".yml", ".ini", ".tsv"
                        text = SharedMethods.ReadTextFile(path, False)
                    Case Else
                        ' Last resort for binary/media files: hand off to the LLM (vision/audio capable models).
                        Try
                            Dim viaLlm = Await SharedMethods.ReadBinaryFileViaLLM(path, context, "", askUser:=False).ConfigureAwait(False)
                            text = If(String.IsNullOrWhiteSpace(viaLlm), "[Unsupported file type for plain-text extraction: " & ext & "]", viaLlm)
                        Catch ex As Exception
                            text = "[Unsupported file type: " & ext & " — " & ex.Message & "]"
                        End Try
                End Select
                Return (If(text, ""), "")
            Catch ex As Exception
                Return ("", ex.Message)
            End Try
        End Function

        ' ═══════════════════════════════════════════════════════════════════════
        '   3)  THREADS
        ' ═══════════════════════════════════════════════════════════════════════

        ''' <summary>
        ''' Returns every message in a mail conversation, ordered per <see cref="M365ThreadOptions.Ascending"/>.
        ''' </summary>
        ''' <summary>
        ''' Returns every message in a mail conversation, ordered per <see cref="M365ThreadOptions.Ascending"/>.
        ''' </summary>
        Public Async Function GetMailThreadAsync(context As ISharedContext,
                                                 conversationId As String,
                                                 Optional options As M365ThreadOptions = Nothing,
                                                 Optional ct As CancellationToken = Nothing) As Task(Of List(Of M365Message))
            If options Is Nothing Then options = New M365ThreadOptions()
            Dim token = Await GetAccessTokenAsync(context, ct).ConfigureAwait(False)

            Dim selectFields = New List(Of String) From {
                "id", "subject", "from", "receivedDateTime", "sentDateTime", "importance",
                "hasAttachments", "bodyPreview", "internetMessageId", "conversationId", "webLink",
                "toRecipients", "ccRecipients"
            }
            If options.IncludeMailBody Then selectFields.Add("body")

            ' IMPORTANT:
            ' Do NOT send $orderby together with this filter. In some mailboxes Graph
            ' returns "InefficientFilter". We fetch unsorted and sort client-side.
            Dim filter = New StringBuilder($"conversationId eq '{conversationId.Replace("'", "''")}'")
            If options.From.HasValue Then filter.Append($" and receivedDateTime ge {options.From.Value:yyyy-MM-ddTHH:mm:ssZ}")
            If options.To.HasValue Then filter.Append($" and receivedDateTime le {options.To.Value:yyyy-MM-ddTHH:mm:ssZ}")

            Dim url = $"{GraphV1}/me/messages?" &
                      $"$select={String.Join(",", selectFields)}" &
                      $"&$filter={Uri.EscapeDataString(filter.ToString())}" &
                      $"&$top={Math.Min(options.MaxMessages, 100)}"

            Dim raw = Await GraphPageAsync(token, url, options.MaxMessages, ct).ConfigureAwait(False)

            Dim out As New List(Of M365Message)()
            For Each j In raw
                Dim flags = M365MessageFields.Recipients
                If options.IncludeMailBody Then flags = flags Or M365MessageFields.Body
                out.Add(ParseMessage(j, flags))
            Next

            If options.Ascending Then
                out = out.
                    OrderBy(Function(m) If(m.ReceivedUtc, If(m.SentUtc, DateTime.MinValue))).
                    ToList()
            Else
                out = out.
                    OrderByDescending(Function(m) If(m.ReceivedUtc, If(m.SentUtc, DateTime.MinValue))).
                    ToList()
            End If

            Return out
        End Function

        ''' <summary>
        ''' Returns the message history of a 1:1 / group Teams chat, paginated.
        ''' </summary>
        Public Async Function GetChatThreadAsync(context As ISharedContext,
                                                 chatId As String,
                                                 Optional options As M365ThreadOptions = Nothing,
                                                 Optional ct As CancellationToken = Nothing) As Task(Of List(Of M365ChatMessage))
            If options Is Nothing Then options = New M365ThreadOptions()
            Dim token = Await GetAccessTokenAsync(context, ct).ConfigureAwait(False)
            Dim url = $"{GraphV1}/chats/{Uri.EscapeDataString(chatId)}/messages?$top={Math.Min(options.MaxMessages, 50)}"
            Dim pages = Await GraphPageAsync(token, url, options.MaxMessages, ct).ConfigureAwait(False)
            Dim out = pages.Select(Function(j) ParseChatMessage(j, chatId, Nothing, Nothing)) _
                            .Where(Function(c) ApplyChatDateFilter(c, options)) _
                            .ToList()
            If options.Ascending Then
                out = out.OrderBy(Function(c) c.CreatedUtc).ToList()
            Else
                out = out.OrderByDescending(Function(c) c.CreatedUtc).ToList()
            End If
            Return out
        End Function

        ''' <summary>
        ''' Returns a Teams channel post (the "root" message) plus all its replies.
        ''' </summary>
        Public Async Function GetChannelThreadAsync(context As ISharedContext,
                                                    teamId As String,
                                                    channelId As String,
                                                    rootMessageId As String,
                                                    Optional options As M365ThreadOptions = Nothing,
                                                    Optional ct As CancellationToken = Nothing) As Task(Of List(Of M365ChatMessage))
            If options Is Nothing Then options = New M365ThreadOptions()
            Dim token = Await GetAccessTokenAsync(context, ct).ConfigureAwait(False)
            Dim out As New List(Of M365ChatMessage)()

            ' Root message
            Try
                Dim rootJson = Await GraphGetAsync(token,
                    $"{GraphV1}/teams/{Uri.EscapeDataString(teamId)}/channels/{Uri.EscapeDataString(channelId)}/messages/{Uri.EscapeDataString(rootMessageId)}",
                    ct).ConfigureAwait(False)
                out.Add(ParseChatMessage(rootJson, Nothing, teamId, channelId))
            Catch ex As Exception
                Debug.WriteLine($"[M365] Channel root fetch failed: {ex.Message}")
            End Try

            ' Replies
            Dim url = $"{GraphV1}/teams/{Uri.EscapeDataString(teamId)}/channels/{Uri.EscapeDataString(channelId)}/messages/{Uri.EscapeDataString(rootMessageId)}/replies?$top={Math.Min(options.MaxMessages, 50)}"
            Dim pages = Await GraphPageAsync(token, url, options.MaxMessages, ct).ConfigureAwait(False)
            For Each j In pages
                Dim cm = ParseChatMessage(j, Nothing, teamId, channelId)
                If ApplyChatDateFilter(cm, options) Then out.Add(cm)
            Next

            If options.Ascending Then
                out = out.OrderBy(Function(c) c.CreatedUtc).ToList()
            Else
                out = out.OrderByDescending(Function(c) c.CreatedUtc).ToList()
            End If
            Return out
        End Function

        Private Function ApplyChatDateFilter(cm As M365ChatMessage, options As M365ThreadOptions) As Boolean
            If cm Is Nothing Then Return False
            If options.From.HasValue AndAlso (Not cm.CreatedUtc.HasValue OrElse cm.CreatedUtc.Value < options.From.Value) Then Return False
            If options.To.HasValue AndAlso (Not cm.CreatedUtc.HasValue OrElse cm.CreatedUtc.Value > options.To.Value) Then Return False
            Return True
        End Function

        Private Function ParseChatMessage(j As JObject, chatId As String, teamId As String, channelId As String) As M365ChatMessage
            If j Is Nothing Then Return Nothing
            Dim cm As New M365ChatMessage() With {
                .RawJson = j,
                .Id = SafeStr(j, "id"),
                .ChatId = chatId,
                .TeamId = teamId,
                .ChannelId = channelId,
                .CreatedUtc = TryDate(j, "createdDateTime"),
                .WebUrl = SafeStr(j, "webUrl")
            }
            Dim body = TryCast(j("body"), JObject)
            If body IsNot Nothing Then
                cm.ContentType = SafeStr(body, "contentType")
                cm.Content = SafeStr(body, "content")
            End If
            Dim fromUser = TryCast(TryCast(j("from"), JObject)?("user"), JObject)
            If fromUser IsNot Nothing Then cm.From = SafeStr(fromUser, "displayName")
            Return cm
        End Function

        ' ═══════════════════════════════════════════════════════════════════════
        '   Local helpers
        ' ═══════════════════════════════════════════════════════════════════════

        Private Function EnsureWorkingFolder(options As M365TextOptions, prefix As String) As String
            If Not String.IsNullOrEmpty(options.WorkingFolder) Then
                Directory.CreateDirectory(options.WorkingFolder)
                Return options.WorkingFolder
            End If
            Dim p = Path.Combine(Path.GetTempPath(), prefix & Guid.NewGuid().ToString("N"))
            Directory.CreateDirectory(p)
            Return p
        End Function

        Private Function FormatAddr(name As String, addr As String) As String
            If String.IsNullOrEmpty(name) AndAlso String.IsNullOrEmpty(addr) Then Return ""
            If String.IsNullOrEmpty(name) Then Return addr
            If String.IsNullOrEmpty(addr) Then Return name
            Return $"{name} <{addr}>"
        End Function

        Private Function TruncateIfNeeded(s As String, options As M365TextOptions, r As M365TextResult) As String
            If options Is Nothing OrElse options.MaxChars <= 0 Then Return s
            If s.Length <= options.MaxChars Then Return s
            r.Truncated = True
            Return s.Substring(0, options.MaxChars) & vbCrLf & "[…truncated]"
        End Function

    End Module

End Namespace
