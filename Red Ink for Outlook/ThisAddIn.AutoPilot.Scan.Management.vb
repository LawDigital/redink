' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.AutoPilot.Scan.Management.vb
' Purpose: Scan Handler administration and host-only runtime actions.
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
    Private _scanDashboard As System.Windows.Forms.Form

    Friend Sub ShowScanInboxDashboard()
        If _scanDashboard IsNot Nothing AndAlso Not _scanDashboard.IsDisposed Then
            _scanDashboard.Show()
            _scanDashboard.BringToFront()
            Return
        End If
        EnsureAutoPilotSettingsRestoredFromRegistry()
        Dim host As New ScanInbox.AdminHost With {
            .GetMailboxes = AddressOf GetScanAdminMailboxes,
            .RuntimeSummary = Function() If(_apActive, "AutoPilot running", "AutoPilot stopped") & " · Handler " & If(_apActive AndAlso _apConfig IsNot Nothing AndAlso _apConfig.EnableScanProcessing, "running", If(My.Settings.AP_EnableScanProcessing, "enabled", "disabled")) & If(_scanSweepRunning <> 0, " · Search running", ""),
            .ReadActivity = Function() System.Threading.Tasks.Task.Run(Function() ReadScanAdminActivity()),
            .ReadSearchProgress = Function() System.Threading.Tasks.Task.Run(Function() ReadScanAdminSearchProgress())}
        _scanDashboard = New ScanInbox.AdminDashboard(GetScanRegisterPath(), host)
        _scanDashboard.Show()
    End Sub

    Private Function GetScanAdminMailboxes() As System.Collections.Generic.List(Of System.String)
        Dim result As New System.Collections.Generic.List(Of System.String)()
        Dim accounts As Microsoft.Office.Interop.Outlook.Accounts = Nothing
        Try
            accounts = Application.Session.Accounts
            For index As System.Int32 = 1 To accounts.Count
                Dim account As Microsoft.Office.Interop.Outlook.Account = Nothing
                Try
                    account = accounts.Item(index)
                    If Not System.String.IsNullOrWhiteSpace(account.SmtpAddress) Then result.Add(account.SmtpAddress)
                Finally
                    ReleaseScanCom(account)
                End Try
            Next
        Finally
            ReleaseScanCom(accounts)
        End Try
        Return result
    End Function

    Private Function ReadScanAdminActivity() As System.Collections.Generic.List(Of ScanInbox.AdminActivity)
        Dim result As New System.Collections.Generic.List(Of ScanInbox.AdminActivity)()
        Dim names As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.Ordinal)
        Try
            For Each rule In ScanInbox.Registry.Load(GetScanRegisterPath()).RuleSets
                names(rule.Id) = rule.Name
            Next
        Catch ex As System.Exception
            ' Activity remains readable when the configuration needs repair.
        End Try
        For Each directory In System.IO.Directory.GetDirectories(GetScanStateRoot())
            Dim path = System.IO.Path.Combine(directory, "operation.json")
            If Not System.IO.File.Exists(path) Then Continue For
            Dim id = System.IO.Path.GetFileName(directory)
            Try
                Dim operation = ScanInbox.Registry.Parse(Of ScanInbox.Operation)(System.IO.File.ReadAllText(path))
                If operation.Id <> id Then Throw New System.IO.InvalidDataException("Package identity does not match its storage directory.")
                Dim displayName As System.String = Nothing
                If Not names.TryGetValue(operation.RuleSetId, displayName) Then displayName = operation.RuleSetId
                result.Add(New ScanInbox.AdminActivity With {.Id = id, .Received = operation.ReceivedUtc, .Mailbox = operation.Mailbox, .RuleSet = displayName,
                    .Status = operation.Status, .Deliveries = operation.Deliveries.Where(Function(d) d.Status = "submitted").Count().ToString() & "/" & operation.Deliveries.Count.ToString(),
                    .Issue = operation.ErrorMessage, .Operation = operation})
            Catch ex As System.Exception
                result.Add(New ScanInbox.AdminActivity With {.Id = id, .Status = "unreadable", .RuleSet = "Unavailable", .Issue = ex.Message})
            End Try
        Next
        Return result
    End Function

    Private Function ReadScanAdminSearchProgress() As System.Collections.Generic.List(Of ScanInbox.AdminSearchProgress)
        Dim result As New System.Collections.Generic.List(Of ScanInbox.AdminSearchProgress)()
        For Each path In System.IO.Directory.GetFiles(GetScanStateRoot(), "cursor-*.json")
            Try
                Dim value = Newtonsoft.Json.Linq.JObject.Parse(System.IO.File.ReadAllText(path))
                result.Add(New ScanInbox.AdminSearchProgress With {.Mailbox = value.Value(Of System.String)("Mailbox"), .ThroughUtc = value.Value(Of System.DateTime)("ThroughUtc").ToUniversalTime()})
            Catch ex As System.Exception
                result.Add(New ScanInbox.AdminSearchProgress With {.Mailbox = "Unreadable checkpoint", .Issue = ex.Message})
            End Try
        Next
        Return result
    End Function

End Class
