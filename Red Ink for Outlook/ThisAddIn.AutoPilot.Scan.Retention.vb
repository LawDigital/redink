' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.AutoPilot.Scan.Retention.vb
' Purpose: Scan thread seeding uses the existing sender-scoped retention tree.
'
' Architecture:
'  - Part of the AutoPilot Scan Handler; configuration and recipient authority
'    remain in the validated, persisted scan register.
'  - Document/model content never grants transport or filesystem permissions.
'  - Host lifecycle, shared UI and retained operation checkpoints are reused.
' =============================================================================
Option Explicit On
Option Strict On
Option Infer On
Imports System.Linq
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn
    Private Shared Function ScanThreadKey(subject As System.String) As System.String
        Dim match = System.Text.RegularExpressions.Regex.Match(If(subject, ""), "\[RI-SCAN:([a-f0-9]{32})\]", System.Text.RegularExpressions.RegexOptions.CultureInvariant)
        If Not match.Success Then Return ""
        Return "scan-" & match.Groups(1).Value
    End Function

    Private Function ResolveScanThreadKey(mail As AutoPilotMailInfo) As System.String
        Dim legacy = ScanThreadKey(mail.Subject)
        If legacy <> "" Then Return legacy
        If System.String.IsNullOrWhiteSpace(mail.ConversationID) Then Return ""
        Dim userRoot = GetCanonicalUserDirPath(mail.SenderEmail)
        Dim threads = System.IO.Path.Combine(userRoot, AP_ThreadRetentionSubdir)
        If Not System.IO.Directory.Exists(threads) Then Return ""
        Dim matches As New System.Collections.Generic.List(Of System.String)()
        For Each directory In System.IO.Directory.GetDirectories(threads, "scan-*")
            If Not IsPathContained(directory, userRoot) Then Continue For
            Dim metaPath = System.IO.Path.Combine(directory, AP_ThreadMetaFileName)
            If Not System.IO.File.Exists(metaPath) Then Continue For
            Dim meta = Newtonsoft.Json.Linq.JObject.Parse(System.IO.File.ReadAllText(metaPath))
            Dim conversations = TryCast(meta("ScanConversationIds"), Newtonsoft.Json.Linq.JArray)
            If conversations IsNot Nothing AndAlso conversations.Any(Function(v) System.String.Equals(CStr(v), mail.ConversationID, System.StringComparison.Ordinal)) Then matches.Add(System.IO.Path.GetFileName(directory))
        Next
        ' Ambiguous conversations never grant access to more than one retained package.
        Return If(matches.Count = 1, matches(0), "")
    End Function

    Private Function AuthorizedScanThread(mail As AutoPilotMailInfo) As System.Boolean
        If mail Is Nothing OrElse Not _apActive OrElse _apConfig Is Nothing OrElse Not _apConfig.EnableScanProcessing Then Return False
        Try
            Dim key = ResolveScanThreadKey(mail)
            If key = "" Then Return False
            Dim userRoot = GetCanonicalUserDirPath(mail.SenderEmail)
            Dim path = System.IO.Path.Combine(userRoot, AP_ThreadRetentionSubdir, key)
            If Not IsPathContained(path, userRoot) Then Return False
            If path Is Nothing Then Return False
            Dim metaPath = System.IO.Path.Combine(path, AP_ThreadMetaFileName)
            If Not System.IO.File.Exists(metaPath) Then Return False
            Dim meta = Newtonsoft.Json.Linq.JObject.Parse(System.IO.File.ReadAllText(metaPath))
            If meta.Value(Of System.Boolean?)("ScanSeed").GetValueOrDefault() <> True Then Return False
            If Not System.String.Equals(meta.Value(Of System.String)("Sender"), mail.SenderEmail, System.StringComparison.OrdinalIgnoreCase) Then Return False
            Dim expiry = GetThreadDeleteAfterUtc(path)
            If Not expiry.HasValue OrElse expiry.Value <= System.DateTime.UtcNow Then Return False
            Dim config = ScanInbox.Registry.Load(GetScanRegisterPath())
            Dim person = config.Recipients.SingleOrDefault(Function(p) p.Active AndAlso p.Id = meta.Value(Of System.String)("RecipientId") AndAlso
                p.SmtpAddress.Equals(mail.SenderEmail, System.StringComparison.OrdinalIgnoreCase))
            Return person IsNot Nothing AndAlso config.RuleSets.Any(Function(r) r.Active AndAlso r.Id = meta.Value(Of System.String)("RuleSetId") AndAlso r.AllowedRecipientIds.Contains(person.Id))
        Catch ex As System.Exception
            ApDashboardLog("Scan thread authorization unavailable: " & ex.Message, "warn")
            Return False
        End Try
    End Function

    Private Function ResolveRetentionConversationKey(mail As AutoPilotMailInfo) As System.String
        If AuthorizedScanThread(mail) Then Return ResolveScanThreadKey(mail)
        Return ComputeThreadConversationKey(mail)
    End Function

    Private Sub UpdateScanThreadDeliveryState(operation As ScanInbox.Operation, config As ScanInbox.Configuration)
        For Each id In operation.Decision.ToIds.Concat(operation.Decision.CcIds).Distinct()
            Dim person = config.Recipients.Single(Function(p) p.Id = id)
            Dim directory = GetThreadDirSafe(person.SmtpAddress, ("scan-" & operation.Id), False)
            Dim metaPath = System.IO.Path.Combine(directory, AP_ThreadMetaFileName)
            Dim meta = Newtonsoft.Json.Linq.JObject.Parse(System.IO.File.ReadAllText(metaPath))
            meta("ScanConversationIds") = Newtonsoft.Json.Linq.JArray.FromObject(operation.Deliveries.Where(Function(d) d.ToIds.Concat(d.CcIds).Contains(id) AndAlso Not System.String.IsNullOrWhiteSpace(d.ConversationId)).Select(Function(d) d.ConversationId).Distinct())
            ScanInbox.Registry.AtomicWrite(metaPath, meta.ToString())
            Dim path = System.IO.Path.Combine(GetThreadFilesDirSafe(directory, False), "scan-package.json")
            Dim payload = Newtonsoft.Json.Linq.JObject.Parse(System.IO.File.ReadAllText(path))
            payload("Status") = operation.Status
            payload("Deliveries") = Newtonsoft.Json.Linq.JArray.FromObject(operation.Deliveries.Select(Function(d) New With {.Id = d.Id, .Status = d.Status, .SubmittedUtc = d.SubmittedUtc}))
            ScanInbox.Registry.AtomicWrite(path, payload.ToString())
            ' AtomicWrite backup is not an additional retained input artifact.
            If System.IO.File.Exists(path & ".bak") Then System.IO.File.Delete(path & ".bak")
        Next
    End Sub

    Private Sub ResetUnsubmittedScanRetention(operation As ScanInbox.Operation)
        If operation.Deliveries.Any(Function(d) d.Status <> "pending") Then Throw New System.InvalidOperationException("Cannot reset retention after a possible delivery submission.")
        Dim key = ("scan-" & operation.Id)
        If key = "" Then Throw New System.IO.InvalidDataException("Invalid scan thread identity during reanalysis.")
        For Each userRoot In EnumerateUserStorageDirectoriesForAdminAndMaintenance()
            Dim threadDir = System.IO.Path.Combine(userRoot, AP_ThreadRetentionSubdir, key)
            If Not System.IO.Directory.Exists(threadDir) Then Continue For
            If Not IsPathContained(threadDir, userRoot) Then Throw New System.IO.InvalidDataException("Unsafe scan retention path.")
            Dim manifest = System.IO.Path.Combine(threadDir, AP_ThreadFilesSubdir, "scan-package.json")
            If Not System.IO.File.Exists(manifest) Then Throw New System.IO.InvalidDataException("Incomplete retained scan seed needs explicit repair before reanalysis.")
            Dim payload = Newtonsoft.Json.Linq.JObject.Parse(System.IO.File.ReadAllText(manifest))
            If payload.Value(Of System.String)("PackageId") <> operation.Id Then Throw New System.IO.InvalidDataException("Retained package identity mismatch; reanalysis stopped.")
            System.IO.Directory.Delete(threadDir, True)
        Next
    End Sub

    Private Sub SeedScanThread(operation As ScanInbox.Operation, rule As ScanInbox.RuleSet, recipient As ScanInbox.RegisteredRecipient, root As System.String)
        Dim key = ("scan-" & operation.Id)
        If key = "" Then Throw New System.IO.InvalidDataException("Scan thread has no stable operation identity.")
        Dim threadDir = GetThreadDirSafe(recipient.SmtpAddress, key, True)
        Dim existingMeta = System.IO.Path.Combine(threadDir, AP_ThreadMetaFileName)
        If System.IO.File.Exists(existingMeta) Then
            Dim expiry = GetThreadDeleteAfterUtc(threadDir)
            If expiry.HasValue AndAlso expiry.Value > System.DateTime.UtcNow Then Return
            ' An expired seed cannot be silently revived with a new retention period.
            Throw New System.IO.InvalidDataException("Scan retention expired before pending delivery completed; manual review required.")
        End If
        Dim filesDir = GetThreadFilesDirSafe(threadDir, True)
        If filesDir Is Nothing Then Throw New System.IO.IOException("Cannot persist scan thread files.")
        ' No private per-person output or future reply is included in this immutable seed.
        For Each doc In operation.Documents
            Dim source = ScanOperationFile(root, doc.RelativePath)
            Dim dest = System.IO.Path.Combine(filesDir, ScanRetainedFileName(operation, doc))
            If Not IsPathContained(dest, filesDir) Then Throw New System.IO.InvalidDataException("Unsafe retained scan filename.")
            System.IO.File.Copy(source, dest, True)
            Using stream = System.IO.File.OpenRead(dest)
                Using sha = System.Security.Cryptography.SHA256.Create()
                    Dim actual = System.BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", "").ToLowerInvariant()
                    If actual <> doc.Sha256 Then Throw New System.IO.IOException("Retained original hash mismatch.")
                End Using
            End Using
        Next
        Dim retained = New With {.Documents = operation.Documents, .AttachmentNames = operation.AttachmentNames, .Extraction = operation.Extraction, .Decision = operation.Decision,
            .ConfigurationRevision = operation.ConfigurationRevision, .ConfigurationHash = operation.ConfigurationHash,
            .Status = "retained_before_delivery", .PackageId = operation.Id}
        ScanInbox.Registry.AtomicWrite(System.IO.Path.Combine(filesDir, "scan-package.json"), Newtonsoft.Json.JsonConvert.SerializeObject(retained, Newtonsoft.Json.Formatting.Indented))
        Dim meta As New Newtonsoft.Json.Linq.JObject From {
            {"ScanSeed", True}, {"Sender", recipient.SmtpAddress}, {"RecipientId", recipient.Id}, {"RuleSetId", rule.Id},
            {"Mailbox", operation.Mailbox}, {"Subject", operation.Subject}, {"ConversationId", ""}, {"LastTouchedUtc", System.DateTime.UtcNow.ToString("o")},
            {"DeleteAfterUtc", System.DateTime.UtcNow.AddDays(rule.RetentionDays).ToString("o")}}
        ScanInbox.Registry.AtomicWrite(System.IO.Path.Combine(threadDir, AP_ThreadMetaFileName), meta.ToString())
        Dim packageBackup = System.IO.Path.Combine(filesDir, "scan-package.json.bak")
        If System.IO.File.Exists(packageBackup) Then System.IO.File.Delete(packageBackup)
        If Not GetThreadDeleteAfterUtc(threadDir).HasValue Then Throw New System.IO.IOException("Scan retention metadata was not committed.")
    End Sub

    ' Existing Reply path stays private unless the inbound message explicitly includes
    ' other currently authorized participants. Original scanner To/Cc is never consulted.
    Private Function ScanThreadMailbox(info As AutoPilotMailInfo) As System.String
        If Not AuthorizedScanThread(info) Then Return ""
        Dim directory = GetThreadDirSafe(info.SenderEmail, ResolveScanThreadKey(info), False)
        Dim meta = Newtonsoft.Json.Linq.JObject.Parse(System.IO.File.ReadAllText(System.IO.Path.Combine(directory, AP_ThreadMetaFileName)))
        Return If(meta.Value(Of System.String)("Mailbox"), "")
    End Function

    Private Function ScanReplyCcAddresses(original As Microsoft.Office.Interop.Outlook.MailItem, reply As Microsoft.Office.Interop.Outlook.MailItem) As System.Collections.Generic.List(Of System.String)
        Dim result As New System.Collections.Generic.List(Of System.String)()
        If _apConfig Is Nothing OrElse Not _apConfig.EnableScanProcessing Then Return result
        If Not AuthorizedScanThread(ExtractMailInfo(original)) Then Return result
        For index As System.Int32 = 1 To reply.Recipients.Count
            Dim person As Microsoft.Office.Interop.Outlook.Recipient = Nothing
            Try
                person = reply.Recipients.Item(index)
                If person.Type = CInt(Microsoft.Office.Interop.Outlook.OlMailRecipientType.olCC) Then result.Add(ScanRecipientAddress(person))
            Finally
                ReleaseScanCom(person)
            End Try
        Next
        Return result
    End Function

    Private Sub ConfigureScanThreadReply(original As Microsoft.Office.Interop.Outlook.MailItem, reply As Microsoft.Office.Interop.Outlook.MailItem)
        If _apConfig Is Nothing OrElse Not _apConfig.EnableScanProcessing Then Return
        Dim info = ExtractMailInfo(original)
        If Not AuthorizedScanThread(info) Then Return
        Dim config = ScanInbox.Registry.Load(GetScanRegisterPath())
        Dim key = ResolveScanThreadKey(info)
        Dim participants As New System.Collections.Generic.List(Of System.String) From {info.SenderEmail}
        Dim incoming As Microsoft.Office.Interop.Outlook.Recipients = Nothing
        Try
            incoming = original.Recipients
            For i As System.Int32 = 1 To incoming.Count
                Dim item As Microsoft.Office.Interop.Outlook.Recipient = Nothing
                Try
                    item = incoming.Item(i)
                    If item.Type <> CInt(Microsoft.Office.Interop.Outlook.OlMailRecipientType.olTo) AndAlso item.Type <> CInt(Microsoft.Office.Interop.Outlook.OlMailRecipientType.olCC) Then Continue For
                    Dim address = ScanRecipientAddress(item)
                    Dim candidate = config.Recipients.SingleOrDefault(Function(p) p.Active AndAlso p.SmtpAddress.Equals(address, System.StringComparison.OrdinalIgnoreCase))
                    If candidate Is Nothing Then Continue For
                    Dim otherInfo As New AutoPilotMailInfo With {.Subject = info.Subject, .ConversationID = info.ConversationID, .SenderEmail = candidate.SmtpAddress}
                    If AuthorizedScanThread(otherInfo) Then participants.Add(candidate.SmtpAddress)
                Finally
                    ReleaseScanCom(item)
                End Try
            Next
        Finally
            ReleaseScanCom(incoming)
        End Try
        reply.To = ScanInbox.Registry.Address(info.SenderEmail)
        reply.CC = System.String.Join(";", participants.Distinct(System.StringComparer.OrdinalIgnoreCase).Where(Function(a) Not a.Equals(info.SenderEmail, System.StringComparison.OrdinalIgnoreCase)))
        reply.BCC = ""
    End Sub
End Class
