' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.AutoPilot.OutgoingDelivery.vb
' Purpose:
'   Central AutoPilot outgoing-mail delivery planning for Outlook transport-size
'   limits. Normal deliveries remain unchanged. If attachments would make the
'   primary message exceed the conservative safe transport size, the primary
'   message is sent without result attachments and the attachments are delivered
'   in one or more follow-up mails sized by estimated MIME transport bytes.
'
' Invariants:
'  - The operational transport ceiling is 35 MB; AutoPilot targets 30 MB after
'    a 5 MB safety margin.
'  - Attachment planning happens after outgoing-attachment sanitization/ZIP.
'  - The primary reply remains unchanged except for a deterministic delivery
'    notice when splitting is required.
'  - Follow-up mails inherit the same AutoPilot loop marker, category, sending
'    account and cleanup-group metadata as the primary message.
'  - A single attachment that cannot fit safely by itself is never submitted
'    blindly; its filename is disclosed in the primary delivery notice.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Linq

Partial Public Class ThisAddIn

    ''' <summary>Operational Outlook/Exchange message-size ceiling used for AutoPilot delivery planning.</summary>
    Private Const AP_OutgoingMessageTransportLimitBytes As Long = 35L * 1024L * 1024L

    ''' <summary>Reserved headroom for MIME headers, store-specific overhead and estimation error.</summary>
    Private Const AP_OutgoingMessageSafetyMarginBytes As Long = 5L * 1024L * 1024L

    ''' <summary>Conservative per-message target used by the AutoPilot outgoing-delivery pipeline.</summary>
    Private Const AP_MaxSafeOutgoingMessageBytes As Long = AP_OutgoingMessageTransportLimitBytes - AP_OutgoingMessageSafetyMarginBytes

    ''' <summary>Conservative expansion factor for Base64 MIME attachment transport encoding.</summary>
    Private Const AP_MimeAttachmentExpansionFactor As Double = 1.4R

    ''' <summary>Fixed reserve for ordinary MIME/message headers and Outlook transport metadata.</summary>
    Private Const AP_OutgoingMessageFixedOverheadBytes As Long = 512L * 1024L

    ''' <summary>Reserved body/header space in attachment-only follow-up messages.</summary>
    Private Const AP_OutgoingFollowUpReserveBytes As Long = 512L * 1024L

    Private NotInheritable Class AutoPilotOutgoingDeliveryPlan
        Public Property PrimaryAttachments As New System.Collections.Generic.List(Of String)()
        Public Property FollowUpBatches As New System.Collections.Generic.List(Of System.Collections.Generic.List(Of String))()
        Public Property UndeliverableAttachments As New System.Collections.Generic.List(Of String)()

        Public ReadOnly Property UsesAttachmentFollowUps As Boolean
            Get
                Return FollowUpBatches IsNot Nothing AndAlso FollowUpBatches.Count > 0
            End Get
        End Property
    End Class

    ''' <summary>
    ''' Sanitizes outgoing deliverables once, estimates MIME transport size, and either leaves
    ''' the current one-message behavior intact or produces attachment-only follow-up batches.
    ''' </summary>
    Private Function PrepareAutoPilotOutgoingDelivery(primaryHtmlBody As String,
                                                       resultAttachments As System.Collections.Generic.List(Of String)) As AutoPilotOutgoingDeliveryPlan
        Dim plan As New AutoPilotOutgoingDeliveryPlan()
        Dim sanitizedAttachments As System.Collections.Generic.List(Of String) =
            SanitizeOutgoingAttachmentsForDelivery(resultAttachments)

        If sanitizedAttachments.Count = 0 Then Return plan

        Dim estimatedCombinedBytes As Long = EstimateAutoPilotOutgoingMessageBytes(primaryHtmlBody, sanitizedAttachments)
        If estimatedCombinedBytes <= AP_MaxSafeOutgoingMessageBytes Then
            plan.PrimaryAttachments.AddRange(sanitizedAttachments)
            Return plan
        End If

        Dim currentBatch As New System.Collections.Generic.List(Of String)()
        Dim currentBatchBytes As Long = AP_OutgoingMessageFixedOverheadBytes + AP_OutgoingFollowUpReserveBytes

        For Each attachPath As String In sanitizedAttachments
            Dim attachmentBytes As Long = EstimateAutoPilotMimeAttachmentBytes(attachPath)
            Dim singleMessageBytes As Long =
                AP_OutgoingMessageFixedOverheadBytes + AP_OutgoingFollowUpReserveBytes + attachmentBytes

            If singleMessageBytes > AP_MaxSafeOutgoingMessageBytes Then
                plan.UndeliverableAttachments.Add(attachPath)
                Continue For
            End If

            If currentBatch.Count > 0 AndAlso
               currentBatchBytes + attachmentBytes > AP_MaxSafeOutgoingMessageBytes Then
                plan.FollowUpBatches.Add(currentBatch)
                currentBatch = New System.Collections.Generic.List(Of String)()
                currentBatchBytes = AP_OutgoingMessageFixedOverheadBytes + AP_OutgoingFollowUpReserveBytes
            End If

            currentBatch.Add(attachPath)
            currentBatchBytes += attachmentBytes
        Next

        If currentBatch.Count > 0 Then plan.FollowUpBatches.Add(currentBatch)

        ApDashboardLog(
            $"📨 Outgoing message estimated at {FormatAutoPilotDeliveryBytes(estimatedCombinedBytes)}; safe target is {FormatAutoPilotDeliveryBytes(AP_MaxSafeOutgoingMessageBytes)}. " &
            $"Primary message will be sent without result attachments; follow-up batches={plan.FollowUpBatches.Count}; individually too large={plan.UndeliverableAttachments.Count}.",
            "warn")

        Return plan
    End Function

    Private Shared Function EstimateAutoPilotOutgoingMessageBytes(primaryHtmlBody As String,
                                                                   attachments As System.Collections.Generic.IEnumerable(Of String)) As Long
        Dim total As Long = AP_OutgoingMessageFixedOverheadBytes

        If Not String.IsNullOrEmpty(primaryHtmlBody) Then
            total += System.Text.Encoding.UTF8.GetByteCount(primaryHtmlBody)
        End If

        If attachments IsNot Nothing Then
            For Each attachPath As String In attachments
                total += EstimateAutoPilotMimeAttachmentBytes(attachPath)
            Next
        End If

        Return total
    End Function

    Private Shared Function EstimateAutoPilotMimeAttachmentBytes(attachPath As String) As Long
        If String.IsNullOrWhiteSpace(attachPath) OrElse Not System.IO.File.Exists(attachPath) Then
            Throw New System.IO.FileNotFoundException(
                "An outgoing deliverable disappeared before transport-size estimation.",
                If(attachPath, ""))
        End If

        Dim rawBytes As Long = New System.IO.FileInfo(attachPath).Length
        Dim expanded As Double = CDbl(rawBytes) * AP_MimeAttachmentExpansionFactor
        Return System.Convert.ToInt64(System.Math.Ceiling(expanded)) + 4096L
    End Function

    ''' <summary>Builds the deterministic notice inserted into the primary message when attachments are split.</summary>
    Private Function BuildAutoPilotAttachmentSplitNoticeHtml(plan As AutoPilotOutgoingDeliveryPlan) As String
        If plan Is Nothing Then Return ""
        If Not plan.UsesAttachmentFollowUps AndAlso plan.UndeliverableAttachments.Count = 0 Then Return ""

        Dim sb As New System.Text.StringBuilder()
        sb.Append("<div style='margin-top:16px;padding:10px 12px;border:1px solid #d0d0d0;background:#f7f7f7;font-family:Arial,sans-serif;font-size:10pt;'>")

        If plan.UsesAttachmentFollowUps Then
            sb.Append("<strong>Attachment delivery notice:</strong> The result attachments would make this e-mail too large for reliable Outlook delivery. ")
            sb.Append("They will therefore be sent in ")
            sb.Append(plan.FollowUpBatches.Count.ToString(System.Globalization.CultureInfo.InvariantCulture))
            sb.Append(If(plan.FollowUpBatches.Count = 1, " separate follow-up e-mail.", " separate follow-up e-mails."))
        End If

        If plan.UndeliverableAttachments.Count > 0 Then
            If plan.UsesAttachmentFollowUps Then sb.Append("<br/><br/>")
            sb.Append("<strong>Not deliverable by e-mail because the individual file is too large:</strong> ")
            sb.Append(String.Join(", ", plan.UndeliverableAttachments.Select(
                Function(path) System.Net.WebUtility.HtmlEncode(System.IO.Path.GetFileName(path)))))
            sb.Append(".")
        End If

        sb.Append("</div>")
        Return sb.ToString()
    End Function

    ''' <summary>
    ''' Sends attachment-only follow-up messages synchronously on the Outlook UI thread. Each part
    ''' receives the same cleanup-group metadata as the primary mail so retention/deletion treats
    ''' the complete multipart delivery as one AutoPilot conversation bundle.
    ''' </summary>
    Private Sub SendAutoPilotAttachmentFollowUps(recipientAddresses As System.Collections.Generic.IEnumerable(Of String),
                                                 baseSubject As String,
                                                 plan As AutoPilotOutgoingDeliveryPlan,
                                                 sendAccount As Microsoft.Office.Interop.Outlook.Account,
                                                 cleanupGroupId As String,
                                                 cleanupIsEligible As Boolean,
                                                 cleanupAnsweredUtc As DateTime?,
                                                 cleanupDeleteAfterUtc As DateTime?,
                                                 Optional disclaimerHtml As System.String = "")
        If plan Is Nothing OrElse Not plan.UsesAttachmentFollowUps Then Return

        Dim recipients As New System.Collections.Generic.List(Of String)()
        If recipientAddresses IsNot Nothing Then
            recipients = recipientAddresses.
                Where(Function(value) Not String.IsNullOrWhiteSpace(value)).
                Select(Function(value) value.Trim()).
                Distinct(System.StringComparer.OrdinalIgnoreCase).
                ToList()
        End If

        If recipients.Count = 0 Then
            Throw New System.InvalidOperationException(
                "Attachment follow-up delivery has no resolved recipient address.")
        End If

        For batchIndex As Integer = 0 To plan.FollowUpBatches.Count - 1
            Dim followUp As Microsoft.Office.Interop.Outlook.MailItem = Nothing

            Try
                followUp = Application.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem)

                For Each address As String In recipients
                    Dim recipient As Microsoft.Office.Interop.Outlook.Recipient = Nothing
                    Try
                        recipient = followUp.Recipients.Add(address)
                        recipient.Type = Microsoft.Office.Interop.Outlook.OlMailRecipientType.olTo
                    Finally
                        If recipient IsNot Nothing Then
                            Try
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(recipient)
                            Catch
                            End Try
                        End If
                    End Try
                Next

                If Not followUp.Recipients.ResolveAll() Then
                    Throw New System.InvalidOperationException(
                        "Could not resolve one or more recipients for an AutoPilot attachment follow-up.")
                End If

                Dim partNumber As Integer = batchIndex + 1
                Dim partCount As Integer = plan.FollowUpBatches.Count
                followUp.Subject = $"{If(baseSubject, "")} [Attachments {partNumber}/{partCount}]".Trim()
                followUp.BodyFormat = Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatHTML
                followUp.HTMLBody = BuildAutoPilotAttachmentFollowUpHtml(partNumber, partCount) & If(disclaimerHtml, System.String.Empty) & BuildAutoPilotFooter()

                For Each attachPath As String In plan.FollowUpBatches(batchIndex)
                    If String.IsNullOrWhiteSpace(attachPath) OrElse Not System.IO.File.Exists(attachPath) Then
                        Throw New System.IO.FileNotFoundException(
                            "An outgoing deliverable disappeared before attachment follow-up creation.",
                            If(attachPath, ""))
                    End If

                    followUp.Attachments.Add(
                        attachPath,
                        Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue,
                        ,
                        System.IO.Path.GetFileName(attachPath))
                Next

                Try
                    followUp.PropertyAccessor.SetProperty(AP_LoopHeaderProperty, AP_LoopHeaderValue)
                Catch
                End Try
                Try
                    followUp.Categories = AP_CategoryName
                Catch
                End Try

                If Not String.IsNullOrWhiteSpace(cleanupGroupId) Then
                    StampCleanupMetadata(
                        followUp,
                        cleanupGroupId,
                        isEligible:=cleanupIsEligible,
                        answeredUtc:=cleanupAnsweredUtc,
                        deleteAfterUtc:=cleanupDeleteAfterUtc,
                        saveItem:=False)
                End If

                If sendAccount IsNot Nothing Then followUp.SendUsingAccount = sendAccount

                Dim sentSubject As String = followUp.Subject
                Dim sentTo As String = followUp.To
                followUp.Send()

                Try
                    MoveLastSentToInkyReplies(
                        cleanupGroupId,
                        cleanupIsEligible,
                        cleanupAnsweredUtc,
                        cleanupDeleteAfterUtc,
                        sentSubject,
                        sentTo)
                Catch
                End Try

                ApDashboardLog(
                    $"✓ SENT attachment follow-up {partNumber}/{partCount} ({plan.FollowUpBatches(batchIndex).Count} attachment(s)).",
                    "info")

            Catch ex As System.Exception
                ApDashboardLog(
                    $"ERROR sending attachment follow-up {batchIndex + 1}/{plan.FollowUpBatches.Count}: {ex.Message}",
                    "error")
                Throw
            Finally
                If followUp IsNot Nothing Then
                    Try
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(followUp)
                    Catch
                    End Try
                End If
            End Try
        Next
    End Sub

    Private Shared Function BuildAutoPilotAttachmentFollowUpHtml(partNumber As Integer,
                                                                  partCount As Integer) As String
        Return "<div style='font-family:Arial,sans-serif;font-size:11pt;'>" &
               "This is attachment delivery " &
               partNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) &
               " of " &
               partCount.ToString(System.Globalization.CultureInfo.InvariantCulture) &
               " for the preceding AutoPilot message. The attachments are sent separately because the combined message would exceed the safe Outlook delivery size." &
               "</div>"
    End Function

    Private Shared Function FormatAutoPilotDeliveryBytes(value As Long) As String
        Return (CDbl(value) / 1024.0R / 1024.0R).ToString("0.0", System.Globalization.CultureInfo.InvariantCulture) & " MB"
    End Function

End Class
