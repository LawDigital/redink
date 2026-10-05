' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
Option Strict On
Option Explicit On

Partial Public Class ThisAddIn
    Private NotInheritable Class ServerFolderExport
        Public Property Mailbox As System.String
        Public Property Messages As System.Collections.Generic.List(Of Global.SharedLibrary.SharedLibrary.M365Service.ServerExportMail)
    End Class

    Private Function ExportCancellationTimer(cts As System.Threading.CancellationTokenSource, state As PstExportState) As System.Windows.Forms.Timer
        Dim timer As New System.Windows.Forms.Timer() With {.Interval = 200}
        AddHandler timer.Tick, Sub()
                                   If Global.SharedLibrary.SharedLibrary.ProgressBarModule.CancelOperation Then
                                       state.Cancelled = True
                                       cts.Cancel()
                                   End If
                               End Sub
        timer.Start()
        Return timer
    End Function

    Private Async Function PrepareServerExportAsync(folder As Microsoft.Office.Interop.Outlook.MAPIFolder, state As PstExportState) As System.Threading.Tasks.Task
        Try
            If folder.Store.ExchangeStoreType = Microsoft.Office.Interop.Outlook.OlExchangeStoreType.olNotExchange Then Return
            If System.String.IsNullOrWhiteSpace(INI_M365ClientId) Then Throw New System.InvalidOperationException("Complete server coverage requires the configured Microsoft 365 client ID and Mail.Read access; alternatively set Outlook's offline cache to All and download all folders before export.")
            Dim mailbox As System.String = Nothing
            For Each account As Microsoft.Office.Interop.Outlook.Account In Me.Application.Session.Accounts
                If account.DeliveryStore IsNot Nothing AndAlso System.String.Equals(account.DeliveryStore.StoreID, folder.StoreID, System.StringComparison.OrdinalIgnoreCase) Then
                    mailbox = account.SmtpAddress
                    Exit For
                End If
            Next
            If System.String.IsNullOrWhiteSpace(mailbox) Then Throw New System.InvalidOperationException("Server coverage cannot be verified for this shared/archive store: no matching Outlook delivery account.")
            Dim names As New System.Collections.Generic.List(Of System.String)()
            Dim rootId = folder.Store.GetRootFolder().EntryID
            Dim current As Microsoft.Office.Interop.Outlook.MAPIFolder = folder
            While Not System.String.Equals(current.EntryID, rootId, System.StringComparison.OrdinalIgnoreCase)
                names.Add(current.Name)
                current = TryCast(current.Parent, Microsoft.Office.Interop.Outlook.MAPIFolder)
                If current Is Nothing Then Throw New System.IO.InvalidDataException("Outlook folder hierarchy could not be resolved.")
            End While
            names.Reverse()
            Dim localIds As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
            For Each snapshot In GetSortedItemSnapshots(folder)
                localIds.Add(snapshot.EntryID)
            Next
            Using cts As New System.Threading.CancellationTokenSource(), timer = ExportCancellationTimer(cts, state)
                Dim records = Await Global.SharedLibrary.SharedLibrary.M365Service.ListServerExportMailAsync(_context, mailbox, names, cts.Token)
                state.ServerFolders(folder.EntryID) = New ServerFolderExport() With {.Mailbox = mailbox, .Messages = records}
                For Each record In records
                    If Not localIds.Contains(record.OutlookEntryId) Then state.ServerExtraCount += 1
                Next
            End Using
        Catch ex As System.OperationCanceledException
            state.Cancelled = True
        Catch ex As System.Exception
            state.DownloadFailureCount += 1
            AppendError(state, "SERVER-COVERAGE|Folder=" & TryGetFolderPath(folder) & "|Message=" & ex.Message)
        End Try
    End Function

    Private Async Function ExportServerMailAsync(record As Global.SharedLibrary.SharedLibrary.M365Service.ServerExportMail,
                                                  mailbox As System.String,
                                                  folder As Microsoft.Office.Interop.Outlook.MAPIFolder,
                                                  outputDirectory As System.String,
                                                  options As PstExportOptions,
                                                  state As PstExportState) As System.Threading.Tasks.Task
        Using cts As New System.Threading.CancellationTokenSource(), timer = ExportCancellationTimer(cts, state)
            Dim mail = Await Global.SharedLibrary.SharedLibrary.M365Service.GetServerExportMailAsync(_context, mailbox, record.Message.Id, cts.Token)
            Dim attachments = Await Global.SharedLibrary.SharedLibrary.M365Service.ListServerExportAttachmentsAsync(_context, mailbox, mail.Id, cts.Token)
            state.NextItemNumber += 1
            Dim itemId = "MAIL" & state.NextItemNumber.ToString("000000", System.Globalization.CultureInfo.InvariantCulture)
            Dim path = System.IO.Path.Combine(outputDirectory, itemId & ".txt")
            Dim body = If(System.String.Equals(mail.BodyContentType, "html", System.StringComparison.OrdinalIgnoreCase), HtmlToPlainText(mail.Body), If(mail.Body, ""))
            Dim text As New System.Text.StringBuilder()
            text.AppendLine("ItemId: " & itemId)
            text.AppendLine("Source: Microsoft 365 server download")
            text.AppendLine("Folder: " & TryGetFolderPath(folder))
            text.AppendLine("Subject: " & mail.Subject)
            text.AppendLine("From: " & mail.From & " <" & mail.FromAddress & ">")
            text.AppendLine("To: " & System.String.Join("; ", mail.To_))
            text.AppendLine("Cc: " & System.String.Join("; ", mail.Cc))
            text.AppendLine("Bcc: " & System.String.Join("; ", mail.Bcc))
            text.AppendLine("Received: " & FormatNullableDate(mail.ReceivedUtc))
            text.AppendLine("Sent: " & FormatNullableDate(mail.SentUtc))
            text.AppendLine("=== BODY / PRIMARY TEXT ===")
            text.AppendLine(body)
            text.AppendLine("=== ATTACHMENTS ===")
            text.AppendLine("Count: " & attachments.Count.ToString(System.Globalization.CultureInfo.InvariantCulture))
            Dim outputs As New System.Collections.Generic.List(Of System.String)()
            Dim before = state.PlaceholderCount
            Dim index As System.Int32 = 0
            For Each attachment In attachments
                cts.Token.ThrowIfCancellationRequested()
                index += 1
                Dim attachmentId = itemId & "-ATT" & index.ToString("000", System.Globalization.CultureInfo.InvariantCulture)
                Dim rendered As System.String = Nothing
                Dim temporaryDirectory = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "RedInk-ServerExport-" & System.Guid.NewGuid().ToString("N"))
                Try
                    System.IO.Directory.CreateDirectory(temporaryDirectory)
                    If attachment.OdataType = "#microsoft.graph.referenceAttachment" Then Throw New System.NotSupportedException("Cloud reference attachment: remote target was not retrieved.")
                    Dim filename = "attachment" & System.IO.Path.GetExtension(SanitizeFileNameForFile(If(System.String.IsNullOrWhiteSpace(attachment.Name), "attachment.bin", attachment.Name)))
                    If attachment.OdataType = "#microsoft.graph.itemAttachment" Then filename &= ".eml"
                    Dim temporaryPath = System.IO.Path.Combine(temporaryDirectory, filename)
                    Await Global.SharedLibrary.SharedLibrary.M365Service.DownloadServerExportAttachmentAsync(_context, mailbox, mail.Id, attachment.Id, temporaryPath, cts.Token)
                    rendered = Await ExtractTextFromSavedAttachmentAsync(temporaryPath, options)
                    If System.String.IsNullOrWhiteSpace(rendered) OrElse rendered.StartsWith("Error:", System.StringComparison.OrdinalIgnoreCase) Then Throw New System.IO.InvalidDataException(If(rendered, "No extractable text."))
                Catch ex As System.OperationCanceledException
                    Throw
                Catch ex As System.Exception
                    state.PlaceholderCount += 1
                    rendered = "[Attachment placeholder: " & ex.Message & "]"
                    AppendError(state, "SERVER-ATTACHMENT|ItemId=" & itemId & "|Name=" & attachment.Name & "|Message=" & ex.Message)
                Finally
                    Try
                        If System.IO.Directory.Exists(temporaryDirectory) Then System.IO.Directory.Delete(temporaryDirectory, True)
                    Catch ex As System.Exception
                        AppendError(state, "SERVER-TEMP-CLEANUP|Message=" & ex.Message)
                    End Try
                End Try
                Dim attachmentText = "Attachment: " & attachment.Name & System.Environment.NewLine & rendered
                If options.InlineAttachments Then
                    text.AppendLine(attachmentText)
                Else
                    Dim filename = attachmentId & ".txt"
                    System.IO.File.WriteAllText(System.IO.Path.Combine(outputDirectory, filename), attachmentText, New System.Text.UTF8Encoding(False))
                    outputs.Add(filename)
                    text.AppendLine("Attachment: " & attachment.Name & " -> " & filename)
                End If
            Next
            cts.Token.ThrowIfCancellationRequested()
            System.IO.File.WriteAllText(path, text.ToString(), New System.Text.UTF8Encoding(False))
            state.IndexLines.Add(CsvLine(itemId, MakeRelativePath(options.OutputRootDirectory, path), TryGetFolderPath(folder), "Mail", "IPM.Note", mail.Subject,
                                        mail.From & "; " & System.String.Join("; ", mail.To_), FormatNullableDate(mail.ReceivedUtc), attachments.Count.ToString(System.Globalization.CultureInfo.InvariantCulture),
                                        System.String.Join(";", outputs), body.Length.ToString(System.Globalization.CultureInfo.InvariantCulture), If(state.PlaceholderCount > before, "Partial", "OK")))
            state.ExportedItemCount += 1
        End Using
    End Function
End Class
