' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: M365Service.Threads.vb
' Purpose: Higher-level conveniences building on M365Service:
'
'   (a) GetMailThreadAsTextAsync / GetChatThreadAsTextAsync /
'       GetChannelThreadAsTextAsync — flatten an entire thread into a single
'       AI-ready transcript. Mail threads include full attachment extraction
'       via the existing ExpandAttachmentsAsync + extractor pipeline.
'
'   (b) ExpandEventInAttachmentAsync — fetches the full body of a calendar
'       event that arrived as an itemAttachment of a mail message. Pair this
'       with ExpandAttachmentsAsync results: when a node has
'       Kind = ItemAttachment and NestedEventSubject set (or NestedMessage Is
'       Nothing), call this helper with the parent message id + node.Id to
'       populate node.NestedEvent.
'
' Adds two backing additions:
'   - Partial Class M365ExpandedAttachment: adds .Id and .NestedEvent fields.
'   - Private ParseEventInline / ParseEventBody: shared by the helper and the
'     mail-message renderer.
'
' KEY ARCHITECTURE
'   • Surface: M365ExpandedAttachment, M365Service (VB implementation).
'   • Entry points/helpers: GetMailThreadAsTextAsync, GetChatThreadAsTextAsync,
'     GetChannelThreadAsTextAsync.
' =============================================================================

Option Explicit On
Option Strict On

Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary

    ''' <summary>
    ''' Adds an attachment id (needed for follow-up Graph calls) and a parsed
    ''' nested event payload to attachment nodes returned by
    ''' <see cref="M365Service.ExpandAttachmentsAsync"/>.
    ''' </summary>
    Partial Public Class M365ExpandedAttachment
        ''' <summary>Graph attachment id, used by <see cref="M365Service.ExpandEventInAttachmentAsync"/>.</summary>
        Public Property Id As String
        ''' <summary>For itemAttachment whose inner type is "event": the parsed event details.</summary>
        Public Property NestedEvent As M365Event
    End Class

    Partial Public Module M365Service

        ' ═══════════════════════════════════════════════════════════════════════
        '   (a)  THREAD → PLAIN TEXT
        ' ═══════════════════════════════════════════════════════════════════════

        ''' <summary>
        ''' Flattens an entire mail conversation into a single AI-ready transcript.
        ''' Each message is rendered with header block + body (HTML stripped) +
        ''' (optionally) recursively extracted attachments.
        ''' </summary>
        Public Async Function GetMailThreadAsTextAsync(context As ISharedContext,
                                                       conversationId As String,
                                                       Optional textOptions As M365TextOptions = Nothing,
                                                       Optional threadOptions As M365ThreadOptions = Nothing,
                                                       Optional ct As CancellationToken = Nothing) As Task(Of M365TextResult)
            If textOptions Is Nothing Then textOptions = New M365TextOptions()
            If threadOptions Is Nothing Then threadOptions = New M365ThreadOptions()
            ' Mail threads need full body for the transcript; force it on.
            threadOptions.IncludeMailBody = True

            Dim r As New M365TextResult() With {.Source = M365SearchSources.Mail, .Id = conversationId}

            Dim msgs As List(Of M365Message)
            Try
                msgs = Await GetMailThreadAsync(context, conversationId, threadOptions, ct).ConfigureAwait(False)
            Catch ex As Exception
                r.Errors.Add("Thread fetch failed: " & ex.Message)
                Return r
            End Try

            If msgs.Count = 0 Then
                r.Errors.Add("Thread is empty.")
                Return r
            End If

            r.Title = If(msgs(0).Subject, conversationId)

            ' All messages share one working folder so their attachments coexist
            ' on disk for the lifetime of the call.
            Dim work = EnsureWorkingFolder(textOptions, "ri_m365_thread_")
            r.WorkingFolder = work

            Dim sb As New StringBuilder()
            sb.AppendLine($"=== Mail thread: {r.Title} ({msgs.Count} message(s)) ===")
            sb.AppendLine()

            Try
                For i As Integer = 0 To msgs.Count - 1
                    ct.ThrowIfCancellationRequested()
                    Dim m = msgs(i)
                    sb.AppendLine($"───── Message {i + 1} of {msgs.Count} ─────")
                    Await RenderMailMessageAsync(context, m, work, textOptions, sb, r, ct).ConfigureAwait(False)
                    sb.AppendLine()
                Next
            Finally
                If textOptions.CleanupWorkingFolder AndAlso String.IsNullOrEmpty(textOptions.WorkingFolder) Then
                    Try : Directory.Delete(work, True) : r.WorkingFolder = Nothing : Catch : End Try
                End If
            End Try

            r.Text = TruncateIfNeeded(sb.ToString().Trim(), textOptions, r)
            Return r
        End Function

        ''' <summary>
        ''' Flattens a 1:1 / group Teams chat into a transcript.
        ''' Chat attachments (file links inside messages) are recorded but not
        ''' downloaded — feed the message URLs to <see cref="ExpandAttachmentsAsync"/>
        ''' or <see cref="DownloadFileAsync"/> separately if you need their content.
        ''' </summary>
        Public Async Function GetChatThreadAsTextAsync(context As ISharedContext,
                                                       chatId As String,
                                                       Optional textOptions As M365TextOptions = Nothing,
                                                       Optional threadOptions As M365ThreadOptions = Nothing,
                                                       Optional ct As CancellationToken = Nothing) As Task(Of M365TextResult)
            If textOptions Is Nothing Then textOptions = New M365TextOptions()
            If threadOptions Is Nothing Then threadOptions = New M365ThreadOptions()

            Dim r As New M365TextResult() With {.Source = M365SearchSources.Teams, .Id = chatId}
            Dim msgs As List(Of M365ChatMessage)
            Try
                msgs = Await GetChatThreadAsync(context, chatId, threadOptions, ct).ConfigureAwait(False)
            Catch ex As Exception
                r.Errors.Add("Chat fetch failed: " & ex.Message)
                Return r
            End Try

            r.Title = $"Chat {chatId}"
            Dim sb As New StringBuilder()
            sb.AppendLine($"=== Teams chat ({msgs.Count} message(s)) ===")
            sb.AppendLine()
            For Each m In msgs
                ct.ThrowIfCancellationRequested()
                RenderChatMessage(m, sb)
                sb.AppendLine()
            Next
            r.Text = TruncateIfNeeded(sb.ToString().Trim(), textOptions, r)
            Return r
        End Function

        ''' <summary>
        ''' Flattens a Teams channel post (root message + its replies) into a transcript.
        ''' </summary>
        Public Async Function GetChannelThreadAsTextAsync(context As ISharedContext,
                                                          teamId As String,
                                                          channelId As String,
                                                          rootMessageId As String,
                                                          Optional textOptions As M365TextOptions = Nothing,
                                                          Optional threadOptions As M365ThreadOptions = Nothing,
                                                          Optional ct As CancellationToken = Nothing) As Task(Of M365TextResult)
            If textOptions Is Nothing Then textOptions = New M365TextOptions()
            If threadOptions Is Nothing Then threadOptions = New M365ThreadOptions()

            Dim r As New M365TextResult() With {.Source = M365SearchSources.Teams, .Id = rootMessageId}
            Dim msgs As List(Of M365ChatMessage)
            Try
                msgs = Await GetChannelThreadAsync(context, teamId, channelId, rootMessageId,
                                                   threadOptions, ct).ConfigureAwait(False)
            Catch ex As Exception
                r.Errors.Add("Channel fetch failed: " & ex.Message)
                Return r
            End Try

            r.Title = $"Channel post {rootMessageId}"
            Dim sb As New StringBuilder()
            sb.AppendLine($"=== Teams channel thread ({msgs.Count} message(s)) ===")
            sb.AppendLine()
            For Each m In msgs
                ct.ThrowIfCancellationRequested()
                RenderChatMessage(m, sb)
                sb.AppendLine()
            Next
            r.Text = TruncateIfNeeded(sb.ToString().Trim(), textOptions, r)
            Return r
        End Function

        ''' <summary>
        ''' Renders a single mail message into <paramref name="sb"/>, optionally appending
        ''' attachments (extracted via <see cref="ExpandAttachmentsAsync"/> +
        ''' <see cref="AppendAttachmentTextAsync"/>).
        ''' Reused by both single-message and thread paths.
        ''' </summary>
        Private Async Function RenderMailMessageAsync(context As ISharedContext,
                                                      msg As M365Message,
                                                      workingFolder As String,
                                                      options As M365TextOptions,
                                                      sb As StringBuilder,
                                                      aggregate As M365TextResult,
                                                      ct As CancellationToken) As Task
            sb.AppendLine("From: " & FormatAddr(msg.From, msg.FromAddress))
            If msg.To_.Count > 0 Then sb.AppendLine("To: " & String.Join("; ", msg.To_))
            If msg.Cc.Count > 0 Then sb.AppendLine("Cc: " & String.Join("; ", msg.Cc))
            AppendDateLines(sb, "Sent", msg.SentUtc)
            AppendDateLines(sb, "Received", msg.ReceivedUtc)
            sb.AppendLine($"Subject: {If(msg.Subject, "")}")
            sb.AppendLine()
            Dim body = If(msg.Body, "")
            If String.Equals(msg.BodyContentType, "html", StringComparison.OrdinalIgnoreCase) Then body = StripHtml(body)
            sb.AppendLine(body.Trim())

            If options.IncludeAttachments AndAlso msg.HasAttachments AndAlso Not String.IsNullOrEmpty(msg.Id) Then
                Try
                    Dim expandOpts As New M365ExpandOptions() With {
                        .MaxDepth = options.AttachmentMaxDepth,
                        .IncludeInline = options.IncludeInlineAttachments,
                        .FollowReferenceAttachments = True
                    }
                    Dim atts = Await ExpandAttachmentsAsync(context, msg.Id, workingFolder, expandOpts, ct).ConfigureAwait(False)
                    For Each a In atts
                        Await AppendAttachmentTextAsync(context, a, sb, aggregate, options, ct).ConfigureAwait(False)
                    Next
                Catch ex As Exception
                    aggregate.Errors.Add($"Attachments for message {msg.Id}: {ex.Message}")
                End Try
            End If
        End Function

        Private Sub RenderChatMessage(m As M365ChatMessage, sb As StringBuilder)
            If m Is Nothing Then Return

            sb.AppendLine("From: " & If(m.From, "(unknown)"))
            AppendDateLines(sb, "Date", m.CreatedUtc)

            Dim body = If(m.Content, "")
            If String.Equals(m.ContentType, "html", StringComparison.OrdinalIgnoreCase) Then body = StripHtml(body)
            sb.AppendLine(body.Trim())
        End Sub

        ' ═══════════════════════════════════════════════════════════════════════
        '   (b)  EXPAND EVENT ITEM-ATTACHMENT
        ' ═══════════════════════════════════════════════════════════════════════

        ''' <summary>
        ''' For an itemAttachment whose inner type is <c>event</c>, fetches the full event
        ''' (subject, organizer, start/end, attendees, body) and returns it as a
        ''' <see cref="M365Event"/>. If the supplied attachment node is supplied,
        ''' its <see cref="M365ExpandedAttachment.NestedEvent"/> is also populated.
        ''' </summary>
        ''' <param name="messageId">Id of the mail message that owns the attachment.</param>
        ''' <param name="attachmentId">Id of the itemAttachment node.</param>
        ''' <param name="node">Optional node to populate in-place.</param>
        Public Async Function ExpandEventInAttachmentAsync(context As ISharedContext,
                                                           messageId As String,
                                                           attachmentId As String,
                                                           Optional node As M365ExpandedAttachment = Nothing,
                                                           Optional ct As CancellationToken = Nothing) As Task(Of M365Event)
            If String.IsNullOrEmpty(messageId) Then Throw New ArgumentException("messageId required", NameOf(messageId))
            If String.IsNullOrEmpty(attachmentId) Then Throw New ArgumentException("attachmentId required", NameOf(attachmentId))

            Dim token = Await GetAccessTokenAsync(context, ct).ConfigureAwait(False)
            Dim url = $"{GraphV1}/me/messages/{Uri.EscapeDataString(messageId)}/attachments/{Uri.EscapeDataString(attachmentId)}?$expand=microsoft.graph.itemAttachment/item"
            Dim j = Await GraphGetAsync(token, url, ct).ConfigureAwait(False)
            Dim inner = TryCast(j("item"), JObject)
            If inner Is Nothing Then Return Nothing

            Dim innerType = SafeStr(inner, "@odata.type").ToLowerInvariant()
            If Not innerType.Contains("event") Then Return Nothing

            Dim ev = ParseEventInline(inner)
            If node IsNot Nothing Then
                node.NestedEvent = ev
                If String.IsNullOrEmpty(node.NestedEventSubject) Then node.NestedEventSubject = ev.Subject
            End If
            Return ev
        End Function

        ''' <summary>
        ''' Convenience: walks a list of expanded attachments and resolves every
        ''' itemAttachment-of-type-event node by fetching its full body. Useful right
        ''' after <see cref="ExpandAttachmentsAsync"/> when the caller wants events
        ''' fully materialised before rendering.
        ''' </summary>
        Public Async Function ResolveEventAttachmentsAsync(context As ISharedContext,
                                                           messageId As String,
                                                           attachments As IEnumerable(Of M365ExpandedAttachment),
                                                           Optional ct As CancellationToken = Nothing) As Task
            If attachments Is Nothing Then Return
            For Each a In attachments
                ct.ThrowIfCancellationRequested()
                If a Is Nothing Then Continue For
                If a.Kind = M365ExpandedAttachmentKind.ItemAttachment _
                   AndAlso a.NestedEvent Is Nothing _
                   AndAlso (a.NestedMessage Is Nothing) _
                   AndAlso Not String.IsNullOrEmpty(a.Id) Then
                    Try
                        Await ExpandEventInAttachmentAsync(context, messageId, a.Id, a, ct).ConfigureAwait(False)
                    Catch ex As Exception
                        a.ErrorMessage = If(a.ErrorMessage, "") & " [event-resolve: " & ex.Message & "]"
                    End Try
                End If
                ' Recurse into children that may themselves contain event item-attachments.
                If a.Children IsNot Nothing AndAlso a.Children.Count > 0 Then
                    Await ResolveEventAttachmentsAsync(context, messageId, a.Children, ct).ConfigureAwait(False)
                End If
            Next
        End Function

        ''' <summary>Parses an inline event JObject into a <see cref="M365Event"/>.</summary>
        Private Function ParseEventInline(j As JObject) As M365Event
            If j Is Nothing Then Return Nothing
            Dim ev As New M365Event() With {.RawJson = j}
            ev.Id = SafeStr(j, "id")
            ev.Subject = SafeStr(j, "subject")
            ev.BodyPreview = SafeStr(j, "bodyPreview")
            ev.WebLink = SafeStr(j, "webLink")
            ev.IsAllDay = CBool(If(j("isAllDay"), False))
            ev.StartUtc = TryDate(TryCast(j("start"), JObject), "dateTime")
            ev.EndUtc = TryDate(TryCast(j("end"), JObject), "dateTime")
            Dim body = TryCast(j("body"), JObject)
            If body IsNot Nothing Then
                Dim text = SafeStr(body, "content")
                If String.Equals(SafeStr(body, "contentType"), "html", StringComparison.OrdinalIgnoreCase) Then
                    text = StripHtml(text)
                End If
                If Not String.IsNullOrEmpty(text) Then ev.BodyPreview = text
            End If
            Dim org = TryCast(TryCast(j("organizer"), JObject)?("emailAddress"), JObject)
            If org IsNot Nothing Then ev.Organizer = SafeStr(org, "name")
            Dim loc = TryCast(j("location"), JObject)
            If loc IsNot Nothing Then ev.Location = SafeStr(loc, "displayName")
            Dim att = TryCast(j("attendees"), JArray)
            If att IsNot Nothing Then
                For Each a In att
                    Dim ea = TryCast(TryCast(a("emailAddress"), JObject), JObject)
                    Dim disp = SafeStr(ea, "name")
                    Dim addr = SafeStr(ea, "address")
                    ev.Attendees.Add(If(String.IsNullOrEmpty(disp), addr, $"{disp} <{addr}>"))
                Next
            End If
            Return ev
        End Function

    End Module

End Namespace