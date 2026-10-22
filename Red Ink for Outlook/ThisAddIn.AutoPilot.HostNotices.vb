' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.AutoPilot.HostNotices.vb
' Purpose:
'   Immutable host-owned response obligations for incoming attachments that
'   AutoPilot skipped. Their content is composed by Outlook, not the model.
'
' Invariants:
'   - Immutable per-request facts and already-localized message wording.
'   - Reuse existing LLM host-message localization with no supported-language list.
'   - Preserve structured facts independently from LLM-translated prose.
'   - Missing or invalid translation, facts or placeholders fails before Send.
'   - Deferred approval and final delivery reuse the same localized snapshot.
' =============================================================================

Option Explicit On
Option Strict Off

Partial Public Class ThisAddIn

    Private NotInheritable Class AutoPilotHostNoticeAttachment
        Public ReadOnly Property FileName As System.String
        Public ReadOnly Property SizeBytes As System.Int64

        Public Sub New(fileName As System.String, sizeBytes As System.Int64)
            If System.String.IsNullOrWhiteSpace(fileName) OrElse sizeBytes < 0 Then
                Throw New System.ArgumentException("A skipped attachment requires a name and non-negative size.")
            End If
            Me.FileName = fileName
            Me.SizeBytes = sizeBytes
        End Sub
    End Class

    ' <summary>Immutable host facts; no locale-specific prose is stored in the snapshot.</summary>
    Private NotInheritable Class AutoPilotHostNotice
        Public ReadOnly Property Id As System.String
        Public ReadOnly Property ReasonCode As System.String
        Public ReadOnly Property ConfiguredLimitBytes As System.Int64
        Public ReadOnly Property Attachments As System.Collections.Generic.IReadOnlyList(Of AutoPilotHostNoticeAttachment)
        Public ReadOnly Property Wording As AutoPilotHostNoticeText

        Public Sub New(id As System.String,
                       reasonCode As System.String,
                       configuredLimitBytes As System.Int64,
                       attachments As System.Collections.Generic.IEnumerable(Of AutoPilotHostNoticeAttachment),
                       Optional wording As AutoPilotHostNoticeText = Nothing)
            If System.String.IsNullOrWhiteSpace(id) OrElse
               System.String.IsNullOrWhiteSpace(reasonCode) OrElse
               configuredLimitBytes <= 0 OrElse attachments Is Nothing Then
                Throw New System.ArgumentException("Host notice identity, reason and limit must be complete.")
            End If
            Dim copied As New System.Collections.Generic.List(Of AutoPilotHostNoticeAttachment)()
            For Each item As AutoPilotHostNoticeAttachment In attachments
                If item Is Nothing Then Throw New System.ArgumentException("Null attachment in host notice.")
                copied.Add(item)
            Next
            If copied.Count = 0 Then Throw New System.ArgumentException("Host notice attachment list is empty.")
            Me.Id = id
            Me.ReasonCode = reasonCode
            Me.ConfiguredLimitBytes = configuredLimitBytes
            Me.Attachments = copied.AsReadOnly()
            Me.Wording = wording
        End Sub
    End Class

    Private NotInheritable Class AutoPilotHostNoticeText
        Public ReadOnly Property Heading As System.String
        Public ReadOnly Property Description As System.String
        Public ReadOnly Property Resolution As System.String

        Public Sub New(heading As System.String, description As System.String, resolution As System.String)
            If System.String.IsNullOrWhiteSpace(heading) OrElse
               System.String.IsNullOrWhiteSpace(description) OrElse
               System.String.IsNullOrWhiteSpace(resolution) Then
                Throw New System.ArgumentException("A localized host notice is incomplete.")
            End If
            Me.Heading = heading
            Me.Description = description
            Me.Resolution = resolution
        End Sub
    End Class

    ' <summary>Captures skipped incoming files and the actual configured limit.</summary>
    Private Shared Function BuildAutoPilotOversizedAttachmentNotices(
            oversizedAttachments As System.Collections.Generic.IEnumerable(Of AutoPilotAttachmentInfo),
            maxAttachmentBytes As System.Int64) As System.Collections.Generic.IReadOnlyList(Of AutoPilotHostNotice)

        If maxAttachmentBytes <= 0 Then
            Throw New System.ArgumentOutOfRangeException(NameOf(maxAttachmentBytes),
                "A positive configured input-attachment limit is required.")
        End If

        Dim attachments As New System.Collections.Generic.List(Of AutoPilotHostNoticeAttachment)()
        If oversizedAttachments IsNot Nothing Then
            For Each item As AutoPilotAttachmentInfo In oversizedAttachments
                If item Is Nothing OrElse Not item.IsOverSizeLimit OrElse
                   System.String.IsNullOrWhiteSpace(item.OriginalFileName) OrElse
                   item.SizeBytes < 0 Then
                    Throw New System.InvalidOperationException("Incomplete metadata for a skipped input attachment.")
                End If
                attachments.Add(New AutoPilotHostNoticeAttachment(item.OriginalFileName, item.SizeBytes))
            Next
        End If
        If attachments.Count = 0 Then Return System.Array.Empty(Of AutoPilotHostNotice)()

        Dim notice As New AutoPilotHostNotice(
            "inbound-attachment-size-limit",
            "attachment_exceeds_configured_input_limit",
            maxAttachmentBytes,
            attachments)
        Dim snapshot As New System.Collections.Generic.List(Of AutoPilotHostNotice) From {notice}
        Return snapshot.AsReadOnly()
    End Function

    ' Verify model-translated templates without trying to infer/guess the language.
    Private Shared Sub VerifyAutoPilotHostNoticeToken(
            translated As System.String,
            placeholder As System.String)
        If System.String.IsNullOrWhiteSpace(translated) OrElse
           translated.Split(New System.String() {placeholder}, System.StringSplitOptions.None).Length <> 2 Then
            Throw New System.InvalidOperationException("Mandatory host translation failed its placeholder contract: " & placeholder)
        End If
    End Sub

    ' Async localization happens ONCE before any send or deferred approval.
    ' The same existing generic LLM helper is used for every language.
    Private Async Function LocalizeAutoPilotHostNoticesAsync(
            notices As System.Collections.Generic.IReadOnlyList(Of AutoPilotHostNotice),
            targetLanguage As System.String,
            cancellationToken As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of System.Collections.Generic.IReadOnlyList(Of AutoPilotHostNotice))

        ValidateAutoPilotHostNotices(notices)
        If notices Is Nothing OrElse notices.Count = 0 Then Return System.Array.Empty(Of AutoPilotHostNotice)()
        If System.String.IsNullOrWhiteSpace(targetLanguage) Then
            Throw New System.InvalidOperationException("Mandatory host notice requires an explicit response language.")
        End If

        Dim result As New System.Collections.Generic.List(Of AutoPilotHostNotice)()
        For Each notice As AutoPilotHostNotice In notices
            If notice.ReasonCode <> "attachment_exceeds_configured_input_limit" Then
                Throw New System.InvalidOperationException("Unknown mandatory host notice reason: " & notice.ReasonCode)
            End If
            Dim heading As System.String = Await LocalizeHostMessageIfNeededAsync(
                "Attachment processing notice", targetLanguage, _apUseSecondApi, True, cancellationToken, strictTranslation:=True)
            Dim description As System.String = Await LocalizeHostMessageIfNeededAsync(
                "The following attachments could not be processed because they exceed the configured size limit of {{LIMIT}}:",
                targetLanguage, _apUseSecondApi, True, cancellationToken, strictTranslation:=True)
            Dim resolution As System.String = Await LocalizeHostMessageIfNeededAsync(
                "Please split these files into smaller parts and resend them, or process them locally with the {{ADDIN}} add-in on your computer.",
                targetLanguage, _apUseSecondApi, True, cancellationToken, strictTranslation:=True)
            VerifyAutoPilotHostNoticeToken(description, "{{LIMIT}}")
            VerifyAutoPilotHostNoticeToken(resolution, "{{ADDIN}}")
            Dim wording As New AutoPilotHostNoticeText(heading, description, resolution)
            result.Add(New AutoPilotHostNotice(notice.Id, notice.ReasonCode,
                                               notice.ConfiguredLimitBytes, notice.Attachments, wording))
        Next
        Return result.AsReadOnly()
    End Function

    Private Shared Function ResolveAutoPilotHostNoticeWording(notice As AutoPilotHostNotice) As AutoPilotHostNoticeText
        If notice Is Nothing OrElse notice.Wording Is Nothing Then
            Throw New System.InvalidOperationException("Mandatory host notice lacks validated translated wording.")
        End If
        Dim limitMb As System.String = (notice.ConfiguredLimitBytes / (1024.0R * 1024.0R)).ToString(
            "F0", System.Globalization.CultureInfo.InvariantCulture) & " MB"
        VerifyAutoPilotHostNoticeToken(notice.Wording.Description, "{{LIMIT}}")
        VerifyAutoPilotHostNoticeToken(notice.Wording.Resolution, "{{ADDIN}}")
        Return New AutoPilotHostNoticeText(
            notice.Wording.Heading,
            notice.Wording.Description.Replace("{{LIMIT}}", limitMb),
            notice.Wording.Resolution.Replace("{{ADDIN}}", AN))
    End Function

    Private Shared Function AutoPilotHostNoticeAttachmentDetail(
            attachment As AutoPilotHostNoticeAttachment) As System.String
        Return attachment.FileName & " (" &
            (attachment.SizeBytes / (1024.0R * 1024.0R)).ToString("F1", System.Globalization.CultureInfo.InvariantCulture) & " MB)"
    End Function

    ' <summary>Validates the host-owned response contract before display and send.</summary>
    Private Shared Sub ValidateAutoPilotHostNotices(
            notices As System.Collections.Generic.IReadOnlyList(Of AutoPilotHostNotice))
        If notices Is Nothing Then Return
        Dim seen As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
        For Each notice As AutoPilotHostNotice In notices
            If notice Is Nothing OrElse
               System.String.IsNullOrWhiteSpace(notice.Id) OrElse
               System.String.IsNullOrWhiteSpace(notice.ReasonCode) OrElse
               notice.ConfiguredLimitBytes <= 0 OrElse
               notice.Attachments Is Nothing OrElse notice.Attachments.Count = 0 Then
                Throw New System.InvalidOperationException("An AutoPilot host notice is incomplete.")
            End If
            If Not seen.Add(notice.Id) Then
                Throw New System.InvalidOperationException("Duplicate AutoPilot host notice identity: " & notice.Id)
            End If
            For Each attachment As AutoPilotHostNoticeAttachment In notice.Attachments
                If attachment Is Nothing OrElse System.String.IsNullOrWhiteSpace(attachment.FileName) OrElse
                   attachment.SizeBytes < 0 Then
                    Throw New System.InvalidOperationException("Host notice contains invalid attachment facts.")
                End If
            Next
        Next
    End Sub

    ' <summary>Renders every notice separately from model-generated HTML.</summary>
    Private Shared Function RenderAutoPilotHostNoticesHtml(
            notices As System.Collections.Generic.IReadOnlyList(Of AutoPilotHostNotice),
            explicitLanguage As System.String) As System.String

        ValidateAutoPilotHostNotices(notices)
        If notices Is Nothing OrElse notices.Count = 0 Then Return System.String.Empty

        Dim html As New System.Text.StringBuilder()
        Dim rendered As System.Int32 = 0
        For Each notice As AutoPilotHostNotice In notices
            Dim wording As AutoPilotHostNoticeText = ResolveAutoPilotHostNoticeWording(notice)
            html.Append("<div data-redink-host-notice-id=""")
            html.Append(System.Net.WebUtility.HtmlEncode(notice.Id))
            html.Append(""" style='margin:16px 0;padding:10px 12px;border:1px solid #d0d0d0;background:#f7f7f7;font-family:Arial,sans-serif;font-size:10pt;'>")
            html.Append("<strong>").Append(System.Net.WebUtility.HtmlEncode(wording.Heading)).Append("</strong><br/>")
            html.Append(System.Net.WebUtility.HtmlEncode(wording.Description))
            html.Append("<ul style='margin:6px 0 8px 0;padding-left:24px;'>")
            For Each detail As AutoPilotHostNoticeAttachment In notice.Attachments
                html.Append("<li>").Append(System.Net.WebUtility.HtmlEncode(AutoPilotHostNoticeAttachmentDetail(detail))).Append("</li>")
            Next
            html.Append("</ul>")
            html.Append(System.Net.WebUtility.HtmlEncode(wording.Resolution)).Append("</div>")
            rendered += 1
        Next
        If rendered <> notices.Count Then
            Throw New System.InvalidOperationException("Composed reply is missing mandatory AutoPilot host notices.")
        End If
        Return html.ToString()
    End Function

    ' <summary>Uses the same structured snapshot and locale in operator approval.</summary>
    Private Shared Function AppendAutoPilotHostNoticesToPreview(
            draftResponse As System.String,
            notices As System.Collections.Generic.IReadOnlyList(Of AutoPilotHostNotice),
            explicitLanguage As System.String) As System.String

        ValidateAutoPilotHostNotices(notices)
        If notices Is Nothing OrElse notices.Count = 0 Then Return If(draftResponse, System.String.Empty)

        Dim plain As New System.Text.StringBuilder(If(draftResponse, System.String.Empty))
        For Each notice As AutoPilotHostNotice In notices
            Dim wording As AutoPilotHostNoticeText = ResolveAutoPilotHostNoticeWording(notice)
            plain.AppendLine().AppendLine().AppendLine("── " & wording.Heading & " ──")
            plain.AppendLine(wording.Description)
            For Each detail As AutoPilotHostNoticeAttachment In notice.Attachments
                plain.AppendLine("  • " & AutoPilotHostNoticeAttachmentDetail(detail))
            Next
            plain.AppendLine(wording.Resolution)
        Next
        Return plain.ToString()
    End Function

    ' <summary>Inform the model about unavailable input, not the sender notification wording.</summary>
    Private Shared Function BuildAutoPilotHostNoticeModelContext(
            notices As System.Collections.Generic.IReadOnlyList(Of AutoPilotHostNotice)) As System.String

        ValidateAutoPilotHostNotices(notices)
        If notices Is Nothing OrElse notices.Count = 0 Then Return System.String.Empty

        Dim sb As New System.Text.StringBuilder()
        sb.AppendLine().AppendLine("[HOST INPUT AVAILABILITY — INFORMATION ONLY]")
        For Each notice As AutoPilotHostNotice In notices
            sb.AppendLine("Reason: " & notice.ReasonCode)
            sb.AppendLine("Configured attachment limit (bytes): " & notice.ConfiguredLimitBytes.ToString(System.Globalization.CultureInfo.InvariantCulture))
            For Each detail As AutoPilotHostNoticeAttachment In notice.Attachments
                sb.AppendLine("  - " & detail.FileName & " (" & detail.SizeBytes.ToString(System.Globalization.CultureInfo.InvariantCulture) & " bytes, excluded)")
            Next
        Next
        sb.AppendLine("Only the available attachments are accessible to your tools. Do not claim excluded attachments were processed.")
        sb.AppendLine("The Outlook host adds the mandatory sender notice. Do not repeat that notice or suggest other tools, services or workarounds for excluded attachments.")
        sb.AppendLine("[/HOST INPUT AVAILABILITY]")
        Return sb.ToString()
    End Function


End Class
