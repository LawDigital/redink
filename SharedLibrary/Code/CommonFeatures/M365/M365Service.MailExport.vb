' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' Server inventory/export is independent of the Outlook offline cache and never uses search ranking.

' =============================================================================
' File: M365Service.MailExport.vb
' Purpose:
'   Graph folder-mail inventory and message/attachment retrieval for export beyond the
'   Outlook offline cache.
'
' Architecture / Function:
'   Uses paged server collection retrieval rather than search ranking; callers retain
'   export scope and cancellation.
' =============================================================================

Option Strict On
Option Explicit On

Namespace SharedLibrary
    Partial Public Module M365Service
        Public NotInheritable Class ServerExportMail
            Public Property Message As M365Message
            Public Property OutlookEntryId As System.String
        End Class

        Public Async Function ListServerExportMailAsync(context As SharedContext.ISharedContext,
                                                         mailbox As System.String,
                                                         folderNames As System.Collections.Generic.IList(Of System.String),
                                                         ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of System.Collections.Generic.List(Of ServerExportMail))
            Dim token = Await GetAccessTokenAsync(context, ct).ConfigureAwait(False)
            Dim user = Await GraphGetAsync(token, GraphV1 & "/me?$select=mail,userPrincipalName", ct).ConfigureAwait(False)
            If Not System.String.Equals(SafeStr(user, "mail"), mailbox, System.StringComparison.OrdinalIgnoreCase) AndAlso
               Not System.String.Equals(SafeStr(user, "userPrincipalName"), mailbox, System.StringComparison.OrdinalIgnoreCase) Then Throw New System.UnauthorizedAccessException("Sign in to Microsoft 365 with the mailbox selected in Outlook: " & mailbox)
            Dim prefix = GraphV1 & "/users/" & System.Uri.EscapeDataString(mailbox)
            Dim folderId As System.String = "msgfolderroot"
            For Each name In folderNames
                Dim children = Await ExportCollectionAsync(token, prefix & "/mailFolders/" & System.Uri.EscapeDataString(folderId) & "/childFolders?includeHiddenFolders=true&$select=id,displayName&$top=100", ct).ConfigureAwait(False)
                Dim matches As New System.Collections.Generic.List(Of Newtonsoft.Json.Linq.JObject)()
                For Each child In children
                    If System.String.Equals(SafeStr(child, "displayName"), name, System.StringComparison.OrdinalIgnoreCase) Then matches.Add(child)
                Next
                If matches.Count <> 1 Then Throw New System.IO.InvalidDataException("Server folder could not be resolved uniquely: " & name)
                folderId = SafeStr(matches(0), "id")
            Next
            Dim records = Await ExportCollectionAsync(token, prefix & "/mailFolders/" & System.Uri.EscapeDataString(folderId) & "/messages?$select=id,subject,from,receivedDateTime,sentDateTime,hasAttachments,internetMessageId&$top=100", ct).ConfigureAwait(False)
            Dim result As New System.Collections.Generic.List(Of ServerExportMail)()
            For start As System.Int32 = 0 To records.Count - 1 Step 100
                ct.ThrowIfCancellationRequested()
                Dim ids As New Newtonsoft.Json.Linq.JArray()
                For index As System.Int32 = start To System.Math.Min(start + 99, records.Count - 1)
                    ids.Add(SafeStr(records(index), "id"))
                Next
                Dim mapping = Await GraphPostAsync(token, prefix & "/translateExchangeIds", New Newtonsoft.Json.Linq.JObject(New Newtonsoft.Json.Linq.JProperty("inputIds", ids), New Newtonsoft.Json.Linq.JProperty("sourceIdType", "restId"), New Newtonsoft.Json.Linq.JProperty("targetIdType", "entryId")), ct).ConfigureAwait(False)
                Dim values = TryCast(mapping("value"), Newtonsoft.Json.Linq.JArray)
                If values Is Nothing Then Throw New System.IO.InvalidDataException("Server message identity conversion returned no records.")
                Dim entryIds As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.Ordinal)
                For Each value In values
                    Dim source = CStr(value("sourceId"))
                    Dim target = CStr(value("targetId"))
                    If System.String.IsNullOrWhiteSpace(source) OrElse System.String.IsNullOrWhiteSpace(target) Then Throw New System.IO.InvalidDataException("A server message identity could not be converted.")
                    ' Graph entryId uses URL-safe base64: '/' => '_', '+' => '-', trailing padding encoded by its count.
                    Dim padding = target(target.Length - 1)
                    If padding < "0"c OrElse padding > "2"c Then Throw New System.IO.InvalidDataException("Invalid binary Exchange ID encoding.")
                    Dim base64 = target.Substring(0, target.Length - 1).Replace("-"c, "+"c).Replace("_"c, "/"c) & New System.String("="c, System.Int32.Parse(padding.ToString(), System.Globalization.CultureInfo.InvariantCulture))
                    entryIds(source) = System.BitConverter.ToString(System.Convert.FromBase64String(base64)).Replace("-", "")
                Next
                For index As System.Int32 = start To System.Math.Min(start + 99, records.Count - 1)
                    Dim record = records(index)
                    Dim id = SafeStr(record, "id")
                    If Not entryIds.ContainsKey(id) Then Throw New System.IO.InvalidDataException("Incomplete server message identity conversion.")
                    result.Add(New ServerExportMail() With {.Message = ParseMessage(record, M365MessageFields.Headers), .OutlookEntryId = entryIds(id)})
                Next
            Next
            Return result
        End Function

        Private Async Function ExportCollectionAsync(token As System.String, url As System.String, ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of System.Collections.Generic.List(Of Newtonsoft.Json.Linq.JObject))
            Dim result As New System.Collections.Generic.List(Of Newtonsoft.Json.Linq.JObject)()
            Dim visited As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            Dim identities As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            While Not System.String.IsNullOrWhiteSpace(url)
                ct.ThrowIfCancellationRequested()
                If Not url.StartsWith(GraphV1 & "/", System.StringComparison.Ordinal) OrElse Not visited.Add(url) Then Throw New System.IO.InvalidDataException("Invalid or repeated server collection continuation.")
                Dim response = Await GraphGetAsync(token, url, ct).ConfigureAwait(False)
                Dim records = TryCast(response("value"), Newtonsoft.Json.Linq.JArray)
                If records Is Nothing Then Throw New System.IO.InvalidDataException("Incomplete server collection response.")
                For Each record In records
                    Dim item = TryCast(record, Newtonsoft.Json.Linq.JObject)
                    If item Is Nothing Then Throw New System.IO.InvalidDataException("Invalid server collection item.")
                    Dim identity = SafeStr(item, "id")
                    If System.String.IsNullOrWhiteSpace(identity) Then Throw New System.IO.InvalidDataException("Server collection item has no identity.")
                    If identities.Add(identity) Then result.Add(item)
                Next
                url = SafeStr(response, "@odata.nextLink")
            End While
            Return result
        End Function

        Public Async Function GetServerExportMailAsync(context As SharedContext.ISharedContext, mailbox As System.String, id As System.String, ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of M365Message)
            Dim token = Await GetAccessTokenAsync(context, ct).ConfigureAwait(False)
            Dim value = Await GraphGetAsync(token, GraphV1 & "/users/" & System.Uri.EscapeDataString(mailbox) & "/messages/" & System.Uri.EscapeDataString(id) & "?$select=id,subject,from,receivedDateTime,sentDateTime,hasAttachments,body,internetMessageId,toRecipients,ccRecipients,bccRecipients", ct).ConfigureAwait(False)
            Dim body = TryCast(value("body"), Newtonsoft.Json.Linq.JObject)
            If body Is Nothing OrElse body("content") Is Nothing OrElse body("content").Type <> Newtonsoft.Json.Linq.JTokenType.String Then Throw New System.IO.InvalidDataException("Server response did not include a valid mail body.")
            Return ParseMessage(value, M365MessageFields.Body Or M365MessageFields.Recipients)
        End Function

        Public Async Function ListServerExportAttachmentsAsync(context As SharedContext.ISharedContext, mailbox As System.String, id As System.String, ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of System.Collections.Generic.List(Of M365AttachmentInfo))
            Dim token = Await GetAccessTokenAsync(context, ct).ConfigureAwait(False)
            Dim records = Await ExportCollectionAsync(token, GraphV1 & "/users/" & System.Uri.EscapeDataString(mailbox) & "/messages/" & System.Uri.EscapeDataString(id) & "/attachments?$select=id,name,contentType,size,isInline", ct).ConfigureAwait(False)
            Dim result As New System.Collections.Generic.List(Of M365AttachmentInfo)()
            For Each record In records
                result.Add(New M365AttachmentInfo() With {.Id = SafeStr(record, "id"), .Name = SafeStr(record, "name"), .ContentType = SafeStr(record, "contentType"), .Size = CLng(record("size")), .IsInline = CBool(record("isInline")), .OdataType = SafeStr(record, "@odata.type")})
            Next
            Return result
        End Function

        Public Async Function DownloadServerExportAttachmentAsync(context As SharedContext.ISharedContext, mailbox As System.String, id As System.String, attachmentId As System.String, path As System.String, ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task
            Dim token = Await GetAccessTokenAsync(context, ct).ConfigureAwait(False)
            Dim url = GraphV1 & "/users/" & System.Uri.EscapeDataString(mailbox) & "/messages/" & System.Uri.EscapeDataString(id) & "/attachments/" & System.Uri.EscapeDataString(attachmentId) & "/$value"
            Using request As New System.Net.Http.HttpRequestMessage(System.Net.Http.HttpMethod.Get, url)
                request.Headers.Authorization = New System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token)
                Using response = Await _http.SendAsync(request, System.Net.Http.HttpCompletionOption.ResponseHeadersRead, ct).ConfigureAwait(False)
                    If Not response.IsSuccessStatusCode Then Await ThrowGraphErrorAsync(response).ConfigureAwait(False)
                    Using stream As New System.IO.FileStream(path, System.IO.FileMode.CreateNew, System.IO.FileAccess.Write)
                        Using content = Await response.Content.ReadAsStreamAsync().ConfigureAwait(False)
                            Dim buffer(81919) As System.Byte
                            While True
                                Dim count = Await content.ReadAsync(buffer, 0, buffer.Length, ct).ConfigureAwait(False)
                                If count = 0 Then Exit While
                                Await stream.WriteAsync(buffer, 0, count, ct).ConfigureAwait(False)
                            End While
                        End Using
                    End Using
                End Using
            End Using
        End Function
    End Module
End Namespace
